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
   Caption         =   "����ҽ������վ"
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
         Caption         =   "��������"
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
            Text            =   "�Һŵ�"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "�����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "�Ա�"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "����"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "��"
            Object.Width           =   635
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "��"
            Object.Width           =   635
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Key             =   "_����"
            Text            =   "����"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ʱ��"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "����ҽ��"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Key             =   "���￨��"
            Text            =   "���￨��"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "��������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Key             =   "��ҽ���"
            Text            =   "��ҽ���"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Key             =   "��ҽ���"
            Text            =   "��ҽ���"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
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
         Caption         =   "����:"
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmOutDoctorStation.frx":058A
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      DefaultCardType =   "���￨"
      IDKindWidth     =   555
      FindPatiShowName=   0   'False
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Text            =   "�Һŵ�"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ա�"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_����"
         Text            =   "����"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ʱ��"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "����ҽ��"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "���￨��"
         Text            =   "���￨��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "��Ⱦ��"
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
            Caption         =   "���뷶��(&I)"
            Height          =   350
            Left            =   6480
            TabIndex        =   102
            Top             =   2040
            Width           =   1200
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "ȫ�ı༭(&U)"
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
            ToolTipText     =   "�鿴���й�����Ϣ"
            Top             =   2078
            Width           =   260
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "ȡ��ǩ��(&Q)"
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
               ToolTipText     =   "�밴 * �ż�ѡ��"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "(���ﲡ��)"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4920
            TabIndex        =   75
            Top             =   1780
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lbl��ʾ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���벡��ʱ�� ~ ������ȡ��ѡ��ʾ�ʾ��."
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   4920
            TabIndex        =   74
            Top             =   2115
            Width           =   3420
         End
         Begin VB.Label lblҽ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "������"
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
            Caption         =   "��  ��"
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
            Caption         =   "����ʷ"
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
            Caption         =   "��  ��"
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
            Caption         =   "��ȥʷ"
            Height          =   540
            Index           =   2
            Left            =   120
            TabIndex        =   69
            Top             =   907
            Width           =   180
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblҽ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ��:"
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
            Caption         =   "�ֲ�ʷ"
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
         Begin VB.Label lblTitle���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
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
         Begin VB.Label lblTitle������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
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
         Begin VB.Label lblTitleҽ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ����:"
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
         Begin VB.Label lblTitle���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
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
         Begin VB.Label lblTitle�ѱ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�:"
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
            Caption         =   "���:"
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
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   5955
            TabIndex        =   16
            Top             =   1155
            Width           =   675
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H00FDFDFD&
            Caption         =   "����"
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
                  Key             =   "չ��"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorStation.frx":15FD
                  Key             =   "�۵�"
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
            Caption         =   "��ʷ"
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
               Name            =   "����"
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
            Caption         =   "��"
            Height          =   220
            Index           =   0
            Left            =   2720
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1170
            Width           =   240
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "��"
            Height          =   220
            Index           =   1
            Left            =   2720
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1470
            Width           =   240
         End
         Begin MSMask.MaskEdBox txt�������� 
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
         Begin MSMask.MaskEdBox txt�������� 
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
         Begin MSMask.MaskEdBox txt����ʱ�� 
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
         Begin MSMask.MaskEdBox txt����ʱ�� 
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
               Caption         =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "�ֻ���:"
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
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "ȥ��:"
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
            Caption         =   "������ַ:"
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
            Caption         =   "ժҪ:"
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
            Caption         =   "����:"
            Height          =   180
            Index           =   14
            Left            =   5235
            TabIndex        =   81
            Top             =   1770
            Width           =   450
         End
         Begin VB.Label lbl��ƾ��� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "���ն�ƾ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8415
            TabIndex        =   78
            Top             =   1185
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lbl�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "סַ:"
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
            Caption         =   "��������:"
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
            Caption         =   "����:"
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
            Caption         =   "����:"
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
            Caption         =   "ְҵ:"
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
            Caption         =   "���֤:"
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
            Caption         =   "��λ:"
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
            Caption         =   "��λ�绰:"
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
            Caption         =   "��ͥ�绰:"
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
            Caption         =   "�໤��:"
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
            Caption         =   "����ʱ��:"
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
         Text            =   "�Һŵ�"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ա�"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_����"
         Text            =   "����"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "��������"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "����ҽ��"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "����ʱ��"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "�Һ�ʱ��"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Key             =   "���￨��"
         Text            =   "���￨��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ת��״̬"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "������"
            TextSave        =   "������"
            Object.ToolTipText     =   "����״̬(�����������)"
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
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":2BB3
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":314D
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":36E7
            Key             =   "ת��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":3C81
            Key             =   "�ܾ�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":421B
            Key             =   "��ͣ"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":47B5
            Key             =   "��Ϣ"
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
         Text            =   "�Һŵ�"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ա�"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_����"
         Text            =   "����"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "����ʱ��"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "���￨��"
         Text            =   "���￨��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ת��״̬"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "��Ⱦ��"
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
         Text            =   "�Һŵ�"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ա�"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_����"
         Text            =   "����"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "�Һ�ʱ��"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "���￨��"
         Text            =   "���￨��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ת��״̬"
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
         Text            =   "�Һŵ�"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ա�"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "��"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ԤԼҽ��"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ԤԼʱ��"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "���֤��"
         Object.Width           =   2541
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "���￨��"
         Text            =   "���￨��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "��������"
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
    pt���� = 0
    pt���� = 1
    pt���� = 2
    ptת�� = 3
    ptԤԼ = 4
    pt���� = 5
    pt�Ŷӽк� = 6
End Enum

Private Enum OPT_ENUM
    opt���� = 0
    opt���� = 1
End Enum

Private Enum cboEnum
    cbo�Ա� = 0
    cbo���� = 1
    cboְҵ = 2
    cboѪѹ��λ = 4
    cbo����ʱ�� = 5
    cboȥ�� = 6
End Enum

Private Enum lineDoc
    lineY1 = 0
    lineY2 = 1
    lineX1 = 2
    lineX2 = 3
End Enum

Private Enum txtEnum    'һ��Ҫ�������
    txt���� = 0
    txt���� = 1
    txt���֤�� = 2
    txt��λ���� = 3
    txt��λ�绰 = 4
    txt��ͥ��ַ = 5
    txt�໤�� = 6
    txt��ͥ�绰 = 7
    txt����ժҪ = 8
    txt������ַ = 9
    txt���� = 10
    txt�ֻ��� = 11
End Enum

Private Enum cmdEnum
    cmd��λ���� = 0
    cmd��ͥ��ַ = 1
End Enum

Private Enum rtfEnum
    txt���� = 0
    txt����ʷ = 3
    txt�ֲ�ʷ = 1
    txt���� = 4
    txt��ȥʷ = 2
End Enum

Private Enum lblEditEnum
    lbl���� = 0
    lbl��λ = 3
    lbl��λ�绰 = 4
    lbl��ͥ�绰 = 7
    lblժҪ = 8
    lbl���� = 14
    lbl���֤ = 20
    lbl����ʱ�� = 21
    lblȥ�� = 15
    lbl�������� = 24
    lbl���� = 13
    lbl�ֻ��� = 11
End Enum

Private Enum lblShowEnum
    lbl�ѱ� = 0
    lbl���� = 1
    lbl���� = 2
    lblҽ���� = 3
    lbl������ = 4
End Enum

Private Enum Msg_Type '��Ϣ�������
    mΣ��ֵ = 1
    m��Ⱦ�� = 2
    m������� = 3
    m��Ѫ��� = 4
    m��Ѫ��� = 5
    m��Ѫ��Ӧ = 6
End Enum
 
Private Enum NOTIFYREPORT_COLUMN
    c_ͼ�� = 0
    C_����ID = 1
    C_No = 2
    c_���� = 3
    c_����� = 4
    C_����ʱ�� = 5
    C_״̬ = 6
    '������
    C_��Ϣ = 7
    C_��� = 8
    C_���� = 9
    C_ҵ�� = 10
    C_�Һ�ID = 11
    C_ID = 12
End Enum

Private Type PatiInfo
    ���� As PatiType
    ����� As String
    �Һ�ID As Long
    �Һŵ� As String
    ����ID As Long
    ���� As String
    ���� As Integer
    ������ As String
    �Һ�ʱ�� As Date
    ����ת�� As Boolean
    �����ļ�id As Long
    ����id As Long
    ������ As String
    �Ƿ�ǩ�� As Boolean
    �Ա� As String
    ����״�� As String
    ���� As String
    ���� As String
    ���� As String
    �����ص� As String
    ��Ⱦ���ϴ� As Long
    ��ͥ��ַ�ʱ� As String
    ��λ�ʱ� As String
    ����֤�� As String
    ���ڵ�ַ As String
    ���ڵ�ַ�ʱ� As String
    ����   As String
    Email As String
    QQ As String
    ����ID As Long
End Type

Private Type ty_Queue
    strQueuePrivs As String '�Ŷӽк�����ģ��Ȩ��
    str����վ�� As String     '���е�վ��:��Ϊ��վ��;����Ϊ����վ��
    byt�Ŷӽк�ģʽ As Byte '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    int�������� As Integer  '0-������,>0��ʾ��������
    bln���к����� As Boolean   '�����Ƿ񺬻�������
    blnҽ���������� As Boolean  'true:��ʾҽ����������;False-ҽ������������
    strCurrQueueName As String '��ǰ��������
    lngcurr�Һ�ID As Long '��ǰ�Һ�ID
End Type
Private mty_Queue As ty_Queue

'�����������
Private Type COND_FILTER
    Begin As Date
    End As Date
    ����ID As Long
    ҽ�� As String
    �Һŵ� As String
    ����� As String
    ���￨ As String
    ���� As String
End Type
Private mvCondFilter As COND_FILTER

'�Ӵ��������
Private mclsEMR As Object  '�°没��zlRichEMR.clsDockEMR
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
Private mobjKernel As zlPublicAdvice.clsPublicAdvice          '�ٴ����Ĳ���

'�������ñ���
Private mint���ﷶΧ As Integer '1-����,2-������,3-������
Private mlng�������ID As Long
Private mstr�������� As String
Private mstr����ҽ�� As String
Private mblnҪ����� As Boolean
Private mintRefresh As Integer '���ﲡ��ˢ�¼��(s)
Private mbln�Զ����� As Boolean
Private mlng�Զ����� As Long
Private mbln���к���� As Boolean
Private mArrDate As Variant

Private mblnDocInput As Boolean    '��ʾ�����������
Private mblnPatiDetail As Boolean  '��ʾ������ϸ��Ϣ
Private mblnPatiChange As Boolean '������Ϣ������ݸı�
Private mblnPatiEditable As Boolean '�Ƿ������޸Ĳ�����Ϣ

Private mlng������� As Long '0-������ 1-��ֹ 2-��ʾ �����:57566
Private mlng��ǰ����ʱ�� As Long  '����Ҫ��ԤԼ�Ž��ս��п���ʱ,��ֵ����ԤԼ�ſ�����ǰ���յķ����� �����:57566
Private mblnUseTYT As Boolean 'ʹ��̫Ԫͨ�ӿ�
Private mint����������Դ As Integer 'ҽ��վ�Ĺ���������Դ
Private mintOutPreTime As Integer

'�����������
Private mrsAller As ADODB.Recordset '���˹�����¼
Private mstrIDCard As String '����Զ�ˢ���������֤��
Private WithEvents mobjIDCard As clsIDCard '���֤����
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object 'IC������
Private mblnUnRefresh As Boolean
Private mstrPrivs As String
Private mlngModul As Long, mstrVirutalPrivs As String
Private mintActive As PatiType '��ǰѡ�������ڵ��б�����
Private mPatiInfo As PatiInfo '��ʷ�����¼�е�,��һ��Ϊ��ǰ��
Private mlng����ID As Long, mstr�Һŵ� As String, mlng����ID As Long '�����嵥�е�
Private mlng�Һ�ID As Long
Private mintFindType As Integer '0-�����￨,1-�����,2-�Һŵ�,3-��������,4-���֤,5-IC��
Private mstrFindType As String '�洢��ǰ�������͵�����
Private mblnIsInit As Boolean 'idkind�ؼ���ʼ����־
Private mblnFindTypeEnabled As Boolean
Private mobjPatient As Object
Private mblnΣ��ֵ As Boolean '��Σ��ֵ��Ȩ��

'ҽ�ƿ�
Private mobjSquareCard As Object      '���������
Private mstrCardKind As String        '��������󷵻صĿ��õ�ҽ�ƿ�
Private Enum CardProperty
    CP���� = 0
    CPȫ�� = 1
    CP�ɶ��� = 2
    CP�����ID = 3
    CP���ų��� = 4
    CPȱʡ��� = 5
    CP�����ʻ� = 6
    CP����������ʾ = 7
End Enum

Private mstrPrePati As String
Private mintPreTime As Integer
Private mblnMouseDown As Boolean
Private mlngCommunityID As Long '�Զ�ִ�е���������
Private mbytSize As Byte '���� 0-С���壨9�����壩��1-�����壨12�����壩
Private mblnTabTmp As Boolean
Private mblnSizeTmp As Boolean

Private mblnMsgOk As Boolean '�Ƿ�����Ϣ����
Private mblnFirstMsg As Boolean 'mblnFirstMsg=false ��ʾ��ҽ��վ��ĵ�һ����Ϣ
Private mintNotify As Integer 'ҽ�������Զ�ˢ�¼��(����)
Private mintNotifyDay As Integer '���Ѷ������ڵ�ҽ��
Private mstrNotifyAdvice As String '���ѵ�ҽ������
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln��Ϣ���� As Boolean
Private mstrPreNotify As String
Private mblnMaskID As Boolean     '�Ƿ����֤������ʾ

Private Sub cboEdit_Click(Index As Integer)
    Dim datCur As Date, datRes As Date
    '�༭״̬
    If cboEdit(Index).List(cboEdit(Index).ListIndex) <> cboEdit(Index).Tag Then
        Call SetPermitEscape(False)
    End If
    If Index = cbo����ʱ�� Then
        If cboEdit(cbo����ʱ��).ListIndex <= 0 Then Exit Sub
        If Trim(txtEdit(txt����).Text) = "" Then Exit Sub
        datCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        Select Case cboEdit(cbo����ʱ��).ListIndex
                Case 1 'Сʱ
                    datRes = DateAdd("n", -1 * Val(txtEdit(txt����).Text) * 60, CDate(datCur))
                Case 2 '��
                    datRes = DateAdd("h", -1 * Val(txtEdit(txt����).Text) * 24, CDate(datCur))
                Case 3 '��
                    datRes = DateAdd("d", -1 * 7 * Val(txtEdit(txt����).Text), CDate(datCur))
                Case 4 '��
                    datRes = DateAdd("M", -1 * Int(Val(txtEdit(txt����).Text)), CDate(datCur))
                    datRes = DateAdd("d", -1 * (Val(txtEdit(txt����).Text) - Int(Val(txtEdit(txt����).Text))) * 30, datRes)
                Case 5 '��
                    If Val(txtEdit(txt����).Text) < 100 Then
                        datRes = DateAdd("yyyy", -1 * Int(Val(txtEdit(txt����).Text)), CDate(datCur))
                        datRes = DateAdd("d", -1 * (Val(txtEdit(txt����).Text) - Int(Val(txtEdit(txt����).Text))) * 365, datRes)
                    Else
                        MsgBox "����ʱ�����㲻�ܳ���100�ꡣ", vbInformation, gstrSysName
                        txtEdit(txt����).SetFocus: Exit Sub
                    End If
        End Select
        txt��������.Text = Format(CDate(datRes), "YYYY-MM-DD")
        If cboEdit(cbo����ʱ��).ListIndex < 3 Then
            txt����ʱ��.Text = Format(CDate(datRes), "HH:mm")
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
            Case cboְҵ
                If SendMessage(cboEdit(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
                lngidx = MatchIndex(cboEdit(Index).hwnd, KeyAscii)
                If lngidx <> -2 Then cboEdit(Index).ListIndex = lngidx
            Case cboȥ��
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
                'ȡ��ʱ�ָ�ԭ����ѡ��
                Call zlControl.CboSetIndex(cboSelectTime.hwnd, mintOutPreTime)
                Exit Sub
            End If
        ElseIf intDateCount = 0 Then
            '����  86114
            mvCondFilter.Begin = Format(datCurr, "yyyy-MM-dd 00:00:00")
            mvCondFilter.End = Format(datCurr, "yyyy-MM-dd 23:59:59")
        Else
            mvCondFilter.End = datCurr
            mvCondFilter.Begin = mvCondFilter.End - intDateCount
        End If
    End If
    'ѡ����ʱ��֮������Һŵ�����
    mvCondFilter.�Һŵ� = ""
    mvCondFilter.���￨ = ""
    mvCondFilter.����� = ""
    mvCondFilter.���� = ""
    '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
    Call zlDatabase.SetPara("���ﲡ�˽������", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
    Call zlDatabase.SetPara("���ﲡ�˿�ʼ���", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
    cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
    lbl����ʱ��.ToolTipText = cboSelectTime.ToolTipText
    mintOutPreTime = cboSelectTime.ListIndex
    
    Call LoadPatients("0010")
End Sub

Private Sub cmdOtherFilter_Click()
    Dim datCurr As Date
    
    With mvCondFilter
        .����ID = IIf(.����ID = 0, mlng�������ID, .����ID)
        If frmPatiFilter.ShowMe(Me, .Begin, .End, .����ID, .ҽ��, .�Һŵ�, .�����, .���￨, .����, mstrPrivs) Then
            datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            Call Cbo.SetIndex(cboSelectTime.hwnd, 5)
            '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
            Call zlDatabase.SetPara("���ﲡ�˽������", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call zlDatabase.SetPara("���ﲡ�˿�ʼ���", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
            lbl����ʱ��.ToolTipText = cboSelectTime.ToolTipText
            mintOutPreTime = cboSelectTime.ListIndex
            Call LoadPatients("0010")
        End If
    End With
End Sub

Private Sub cmdAller_Click()
    Dim i As Long, strTmp As String
    Dim objBar As CommandBar
    
    mrsAller.Filter = "����ID=" & mPatiInfo.����ID & " and �Һŵ�<>'" & mPatiInfo.�Һŵ� & "'"
    If mrsAller.RecordCount > 0 Then
        Set objBar = cbsMain.Add("������¼", xtpBarPopup)
        With mrsAller
            For i = 1 To .RecordCount
                If Not IsNull(!�Һ�ʱ��) Then
                    strTmp = Format(!����ʱ��, "yyyy-MM-dd HH:mm") & ",�������:" & Nvl(!�Һſ���) & "," & Nvl(!ҩ����)
                Else
                    strTmp = Format(!����ʱ��, "yyyy-MM-dd HH:mm") & ",��" & !��ҳID & "��סԺ:" & Nvl(!סԺ����) & "," & Nvl(!ҩ����)
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
    
    If MsgBox("��ȷʵҪ�����Ѹı�����ݣ����¶�ȡ�ò��˵���Ϣ��", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    mintPreTime = -1
    Call cboRegist_Click
    Call LoadAllerInfo(mrsAller)
            
    Call SetPermitEscape(True)
End Sub

Private Sub ExecuteOK()
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean
    Dim lng����ID As Long, blnDoc As Boolean
    Dim objLvw As ListView
    arrSQL = Array()
    
    If InStr(mstrPrivs, "������ҳ") > 0 Then
        If Not CheckOutMediRec Then Exit Sub
    
        Call GetSQLOutMediRec(arrSQL)
    End If
    
    If mblnDocInput And PicOutDoc.Tag = "2" Then
        blnDoc = True
        If mPatiInfo.����id = 0 Then lng����ID = zlDatabase.GetNextId("���Ӳ�����¼")
        Call GetSQLOutDoc(arrSQL, lng����ID)
    End If
    
    
    '�ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        
        If blnDoc Then
            If ReadRTFData(lng����ID) = False Then GoTo errH
            If SaveRTFData(lng����ID) = False Then GoTo errH
        End If
        
        '��������ͬ��
        If Not gobjCommunity Is Nothing And mPatiInfo.������ <> "" Then
            If Not gobjCommunity.UpdateInfo(glngSys, p����ҽ��վ, mPatiInfo.����, mPatiInfo.������, mPatiInfo.����ID, mPatiInfo.�Һ�ID) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
    gcnOracle.CommitTrans: blnTrans = False
    If HaveRIS Then
        If gobjRis.HISModPati(1, mlng����ID, mlng�Һ�ID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    
    If mintActive = pt���� Then
        '����:
        Set objLvw = lvwPatiHZ
    ElseIf mintActive = ptת�� Then
        Set objLvw = lvwIncept
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    
    '�����б��е����������䣬�Ա�
    If txtEdit(txt����).Text <> objLvw.SelectedItem.SubItems(2) Then
         objLvw.SelectedItem.SubItems(2) = txtEdit(txt����).Text
    End If
    If cboEdit(cbo�Ա�).Text <> objLvw.SelectedItem.SubItems(3) Then
        objLvw.SelectedItem.SubItems(3) = cboEdit(cbo�Ա�).Text
    End If
    If txtEdit(txt����).Text & cboEdit(cbo����).Text <> objLvw.SelectedItem.SubItems(4) Then
        objLvw.SelectedItem.SubItems(4) = txtEdit(txt����).Text & cboEdit(cbo����).Text
    End If
    If mintActive <> ptԤԼ Then
        If IIf(optState(opt����).Value, "��", "") <> objLvw.SelectedItem.SubItems(6) Then
            objLvw.SelectedItem.SubItems(6) = IIf(optState(opt����).Value, "��", "")
        End If
    End If
   
    
    'ˢ��mPatiInfo�����Լ��Ӵ�������(����)
    mintPreTime = -1
    Call cboRegist_Click        '���ڲ���IDû�䣬SubWinRefreshData��û��ˢ�²����嵥
    If blnDoc Then
        With mPatiInfo
            Call mclsEPRs.zlRefresh(.����ID, .�Һ�ID, mlng����ID, mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
        End With
    End If
    
    Call ShowAller '���¶�ȡ������Ϣ
    Call SetPermitEscape(True)
 
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ReadRTFData(ByVal lng����ID As Long) As Boolean
'���ܣ���ȡ�����ļ���RTF���ݵ�editor�ؼ���
    Dim strZipFile As String, strTempFile As String
    Dim lngRecID As Long
    
    If mPatiInfo.����id = 0 Then
        lngRecID = lng����ID
    Else
        lngRecID = mPatiInfo.����id
    End If
    
    On Error GoTo errH
    strZipFile = zlBlobRead(5, lngRecID)
    strTempFile = zlFileUnzip(strZipFile)
    edtEditor.OpenDoc strTempFile
    
     'ɾ����ʱ�ļ�
    Kill strTempFile
    Kill strZipFile
   
    ReadRTFData = True
    Exit Function
errH:
    ReadRTFData = False
End Function

Private Function SaveRTFData(ByVal lng����ID As Long, Optional blnSign As Boolean) As Boolean
'���ܣ����没�˲�����ʽRTF����
'������
    Dim strZipFile As String, strTempFile As String, i As Long
    Dim bFinded As Boolean, lngStartPos As Long, lngEndPos As Long, arrTmp As Variant
    Dim strContent As String, lngRecID As Long
    
    If mPatiInfo.����id = 0 Then
        lngRecID = lng����ID
    Else
        lngRecID = mPatiInfo.����id
    End If
    
    If blnSign = False Then
        '�滻�������
        edtEditor.Freeze
        edtEditor.ForceEdit = True
        
        For i = 0 To lblDoc.UBound
            bFinded = FindOutLinePosition(edtEditor, CStr(lblDoc(i).Tag), lngStartPos, lngEndPos)
            If bFinded Then
                strContent = rtfEdit(i).Text    'ȥ��β���Ļس�����
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
        'Ҫ�����ݸ���
        If mPatiInfo.����id = 0 Then Call ElementsUpdate(lngRecID)
    End If
    
    On Error GoTo errH
    strTempFile = App.Path & "\TMP.rtf"
    If Dir(strTempFile) <> "" Then Kill strTempFile
    edtEditor.SaveDoc strTempFile
    'ѹ���ļ�
    strZipFile = zlFileZip(strTempFile)
    '�����ʽ
    zlBlobSave 5, lngRecID, strZipFile
    
    'ɾ����ʱ�ļ�
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
        Dim lngLen As Long, str���� As String
        str���� = Ele.�����ı�
        lngLen = Len(str����)
        With edtThis
            .Freeze
            strOldTag = .Tag
            .Tag = "EleToString"
            bForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = str����
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
    ByVal lngҽ��ID As Long) As String

    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5]) From Dual"
    err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�滻��", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lngҽ��ID)
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

Private Function ElementsUpdate(ByVal lng����ID As Long) As Boolean
'���ܣ�����Editor�ؼ��е��滻Ҫ�����ݣ��Ա㱣��ΪRTF�ļ�
    Dim ThisElements As New zlRichEPR.cEPRElements
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, lngKey As Long
    Dim bFinded As Boolean, bNeeded As Boolean, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long

    strSQL = "Select ������,ID From ���Ӳ������� Where �ļ�ID= [1] And �������� = 4 And ��ֹ��=0 and �������� =0 And �滻�� =1 order by ������ "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    For i = 1 To rsTmp.RecordCount
        lngKey = ThisElements.Add(Nvl(rsTmp("������"), 0))
        ThisElements("K" & lngKey).GetElementFromDB cprET_�������༭, rsTmp("ID"), True
        rsTmp.MoveNext
    Next

     For i = 1 To ThisElements.Count
        If ThisElements(i).�滻�� = 1 Then
            ThisElements(i).�����ı� = GetReplaceEleValue(ThisElements(i).Ҫ������, mPatiInfo.����ID, mPatiInfo.�Һ�ID, 1, 0)
            bFinded = FindNextKey(edtEditor, 0, "E", ThisElements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            ThisElements(i).Refresh edtEditor
        End If
        If ThisElements(i).�滻�� = 1 And ThisElements(i).�Զ�ת�ı� Then
            EleToString edtEditor, ThisElements(i)     '�Զ�ת��Ϊ���ı�����ʱ��ɾ����Ҫ�أ�
        End If
    Next
    Set ThisElements = Nothing
End Function


Public Function FindOutLinePosition(ByRef edtThis As Object, ByVal strOName As String, ByRef lngS As Long, lngE As Long) As Boolean
'���ܣ�����ָ����������ƣ�������������ı�����ֹλ��
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
                Do While lngE > lngS + 1    'ȥ��β���Ļس�����
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
    
    strSQL = "Select Zl_Lob_Read([1],[2],[3]" & IIf(blnMoved, ",1", "") & ") as Ƭ�� From Dual"
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
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
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
    Dim objFSO As New Scripting.FileSystemObject    'FSO����
    
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
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
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
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = .TOM.TextDocument.Range(i - 2, i - 1)
                lngKSS = i - 2 'ת��Ϊ0��ʼ������λ�á�
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
'���ܣ������ҳ�������ݺϷ���
'���أ�
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str���֤ As String, str�������� As String, lng�Ա� As Long
    Dim str���� As String, i As Long, j As Long
    
    
    '��Ŀ����ĳ��ȼ��
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "�������ݹ��������顣(����Ŀ������� " & objTmp.MaxLength & " ���ַ��� " & objTmp.MaxLength \ 2 & " ������)")
                Exit Function
            End If
        End If
    Next
    For Each objTmp In rtfEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "�������ݹ��������顣(����Ŀ������� " & objTmp.MaxLength & " ���ַ��� " & objTmp.MaxLength \ 2 & " ������)")
                Exit Function
            End If
        End If
    Next
    
    '�������ݵ���Ч�Լ��
    '-----------------------------------------------------------------------------------------
    
    curDate = zlDatabase.Currentdate
    
            
    '���֤������
    '�����֤�Ž�����֤
    If mblnMaskID Then
        str���֤ = txtEdit(txt���֤��).Tag
    Else
        str���֤ = txtEdit(txt���֤��).Text
    End If
    If str���֤ <> "" And lblEdit(20).Tag <> str���֤ Then
        If Len(str���֤) <> 15 And Len(str���֤) <> 18 Then
            Call ShowMessage(txtEdit(txt���֤��), "���֤����ĳ��Ȳ���ȷ��ӦΪ15λ��18λ��")
            Exit Function
        End If

        If Len(str���֤) = 15 Then
            str�������� = Mid(str���֤, 7, 6)
            str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
            lng�Ա� = Val(Right(str���֤, 1))
        Else
            str�������� = Mid(str���֤, 7, 8)
            str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
            lng�Ա� = Val(Mid(str���֤, 17, 1))
        End If
        If Not IsDate(str��������) Then
            If ShowMessage(txtEdit(txt���֤��), "���֤�����еĳ���������Ϣ����ȷ���Ƿ������", True) = vbNo Then Exit Function
        ElseIf IsDate(txt��������.Text) Then
            If Format(str��������, "yyyy-MM-dd") <> Format(txt��������.Text, "yyyy-MM-dd") Then
                If ShowMessage(txtEdit(txt���֤��), "���֤�����еĳ���������Ϣ�벡�˵ĳ������ڲ������Ƿ������", True) = vbNo Then Exit Function
            End If
        End If
        If (lng�Ա� Mod 2 = 1 And InStr(cboEdit(cbo�Ա�).Text, "Ů") > 0) Or (lng�Ա� Mod 2 = 0 And InStr(cboEdit(cbo�Ա�).Text, "��") > 0) Then
            If ShowMessage(txtEdit(txt���֤��), "���֤�����е��Ա���Ϣ�벡�˵��Ա𲻷����Ƿ������", True) = vbNo Then Exit Function
        End If
    End If
    
    '����ҩ������
    With vsAller
        For i = 0 To .Cols - 1
            If Trim(.TextMatrix(0, i)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(0, i)) > 60 Then
                    .Col = i
                    Call ShowMessage(vsAller, "����ҩ����̫����ֻ����60���ַ���30�����֡�")
                    Exit Function
                End If
                For j = i + 1 To .Cols - 1
                    If Trim(.TextMatrix(0, j)) <> "" Then
                        If .TextMatrix(0, j) = .TextMatrix(0, i) Then
                            .Col = i
                            Call ShowMessage(vsAller, "���ִ���������ͬ�Ĺ���ҩ�")
                            Exit Function
                        ElseIf Val(.ColData(i)) <> 0 And Val(.ColData(j)) = Val(.ColData(i)) Then
                            .Col = i
                            Call ShowMessage(vsAller, "���ִ���������ͬ�Ĺ���ҩ�")
                            Exit Function
                        ElseIf .TextMatrix(1, i) <> "" And .TextMatrix(1, i) = .TextMatrix(1, j) Then
                            .Col = i
                            Call ShowMessage(vsAller, "���ִ���������ͬ�Ĺ���ҩ�")
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    '����ʱ����
    If txt��������.Text <> "____-__-__" Then
        If Not IsDate(txt��������.Text) Then
            Call ShowMessage(txt��������, "��������ȷ�ķ������ڡ�")
            Exit Function
        Else
            If txt����ʱ��.Text <> "__:__" Then
                If Not IsDate(txt����ʱ��.Text) Then
                    Call ShowMessage(txt����ʱ��, "��������ȷ�ķ���ʱ�䡣")
                    Exit Function
                End If
            End If
            
            If txt��������.Text & IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text) _
                >= Format(curDate, txt��������.Format & IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Format)) Then
                Call ShowMessage(txt��������, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                Exit Function
            End If
        End If
    End If
    
    CheckOutMediRec = True
End Function

Private Sub GetSQLOutMediRec(ByRef arrSQL As Variant)
    '���ܣ�����������ҳ�ĸ�����Ϣ
    Dim i As Integer, lngCnt As Long, blnExist As Boolean, curDate As Date
    Dim str���� As String, str���� As String
    Dim lng��λID As Long
    Dim strTmpSQL As String
    
    curDate = zlDatabase.Currentdate
    
    If IsDate(txt��������.Text) Then
        If IsDate(txt����ʱ��.Text) Then
            str���� = "To_Date('" & Format(txt��������.Text & " " & txt����ʱ��.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            str���� = "To_Date('" & Format(txt��������.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        End If
    Else
       str���� = "NULL"
    End If
    If Trim(txtEdit(txt��λ����).Text) <> "" Then
        lng��λID = Val(txtEdit(txt��λ����).Tag)
    End If
    
    '������Ϣ
    str���� = "NULL"
    If IsDate(txt��������.Text) Then
        If IsDate(txt����ʱ��.Text) Then
            str���� = "To_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI')"
        Else
            str���� = "To_Date('" & txt��������.Text & "','YYYY-MM-DD')"
        End If
    End If
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    With mPatiInfo
    arrSQL(UBound(arrSQL)) = "ZL_������Ϣ_��ҳ����(" & _
        mPatiInfo.����ID & ",'" & mPatiInfo.����� & "','" & txtEdit(txt����).Text & "'," & _
        "'" & NeedName(cboEdit(cbo�Ա�).Text) & "','" & txtEdit(txt����).Text & cboEdit(cbo����).Text & "'," & _
        "'" & .���� & "','" & .���� & "','" & .���� & "','" & .���� & "','" & NeedName(cboEdit(cboְҵ).Text) & "'," & _
        str���� & ",'" & .�����ص� & "','" & txtEdit(txt���֤��).Tag & "','" & .����֤�� & "','" & .����״�� & "'," & _
        "'" & lblShow(lbl����).Caption & "','" & txtEdit(txt��ͥ��ַ).Text & "','" & txtEdit(txt��ͥ�绰).Text & "'," & _
        "'" & .��ͥ��ַ�ʱ� & "','" & .���ڵ�ַ & "','" & .���ڵ�ַ�ʱ� & "'," & ZVal(lng��λID) & "," & _
        "'" & txtEdit(txt��λ����).Text & "','" & txtEdit(txt��λ�绰).Text & "','" & .��λ�ʱ� & "'," & _
        "Null,Null,Null,Null,'" & .Email & "','" & .QQ & "','" & txtEdit(txt�໤��).Text & "','" & mstr�Һŵ� & "'," & _
        IIf(optState(opt����).Value, 1, 0) & ",'" & txtEdit(txt����ժҪ).Text & "'," & .��Ⱦ���ϴ� & "," & str���� & "," & _
        "'" & txtEdit(txt������ַ).Text & "')"
    End With
    
    '�����ֻ���
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_������Ϣ_������Ϣ(" & mPatiInfo.����ID & ",'�ֻ���','" & txtEdit(txt�ֻ���).Text & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mPatiInfo.����ID & ",'ȥ��','" & cboEdit(cboȥ��).Text & "'," & cboRegist.ItemData(cboRegist.ListIndex) & ")"
    strTmpSQL = ucPatiVitalSigns.GetSaveSQL(mlng����ID, cboRegist.ItemData(cboRegist.ListIndex))
    If strTmpSQL <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strTmpSQL
    End If
    
    
            
    '����ҩ��
    With vsAller
        '��������δ�仯ʱ������ɾ������
        lngCnt = 0: blnExist = False
        For i = 0 To .Cols - 1
            If CStr(.Cell(flexcpData, 0, i)) <> "" Then
                blnExist = True '���ֻ��һ���У��򲻵���ɾ��
                If .Cell(flexcpData, 0, i) = .TextMatrix(0, i) Then    'ɾ����������յ�
                    lngCnt = lngCnt + 1
                End If
            End If
        Next
        If blnExist And lngCnt <> .Cols - 1 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Delete(" & mPatiInfo.����ID & "," & mPatiInfo.�Һ�ID & ",3)"
        End If
        
        If blnExist = False Or blnExist And lngCnt <> .Cols - 1 Then
        For i = 0 To .Cols - 1
            If Trim(.TextMatrix(0, i)) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Insert(" & mPatiInfo.����ID & "," & mPatiInfo.�Һ�ID & "," & _
                    "3," & ZVal(.ColData(i)) & ",'" & .TextMatrix(0, i) & "',1," & _
                    "To_Date('" & Format(IIf(.Cell(flexcpData, 1, i) & "" = "", curDate, .Cell(flexcpData, 1, i)), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI:SS'),Null,'" & .TextMatrix(1, i) & "')"
            End If
        Next
        End If
    End With
End Sub


Private Sub GetSQLOutDoc(ByRef arrSQL As Variant, ByVal lng����ID As Long)
'���ܣ���֯��ݲ��������ݱ���SQL
'������lng����ID-����ʱ������ȡ�Ĳ���ID
    Dim i As Long, k As Long
    Dim strTmp(5) As String
    
    If mPatiInfo.����id = 0 Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then Exit Sub     '����ʱ�����û�������ݣ��򲻱���
    End If
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    With mPatiInfo
        If .����id = 0 Then
            If rtfEdit(txt����).Locked Then Exit Sub
            arrSQL(UBound(arrSQL)) = "Zl_�����ﲡ��_Update(1," & mPatiInfo.����ID & "," & _
                .�Һ�ID & "," & mlng����ID & "," & .�����ļ�id & "," & lng����ID & ",'" & UserInfo.���� & "','" & _
                Trim(rtfEdit(txt����).Text) & "','" & Trim(rtfEdit(txt����ʷ).Text) & "','" & Trim(rtfEdit(txt�ֲ�ʷ).Text) & "','" & _
                Trim(rtfEdit(txt����).Text) & "','" & Trim(rtfEdit(txt��ȥʷ).Text) & "')"
        Else
            k = 0
            For i = 0 To rtfEdit.UBound
                If rtfEdit(i).Locked = False Then
                    strTmp(i) = rtfEdit(i).Tag & "|" & Trim(rtfEdit(i).Text)
                    k = k + 1
                End If
            Next
            If k = 0 Then Exit Sub
            
            arrSQL(UBound(arrSQL)) = "Zl_�����ﲡ��_Update(2," & mPatiInfo.����ID & "," & _
                .�Һ�ID & "," & mlng����ID & ",0," & .����id & ",'" & UserInfo.���� & "','" & _
                strTmp(0) & "','" & strTmp(3) & "','" & strTmp(1) & "','" & strTmp(4) & "','" & strTmp(2) & "')"
        End If
    End With
End Sub

Private Function GetEPRDoc() As zlRichEPR.cEPRDocument
'���ܣ���ȡ�����ļ���RTF���ݵ�editor�ؼ��У��������ĵ�����
    Dim objDoc As New zlRichEPR.cEPRDocument
   
    objDoc.InitEPRDoc cprEM_�޸�, cprET_�������༭, mPatiInfo.����id, cprPF_����, mPatiInfo.����ID, mPatiInfo.�Һ�ID, , mPatiInfo.����ID
    If objDoc.ReadFileStructure(edtEditor) = True Then
        Set GetEPRDoc = objDoc
    End If
End Function

Private Sub cmdImportEPRDemo_Click()
    Dim objImportEPRDemo As New frmImportEPRDemo
    Dim rsDemo As New Recordset
    
    If mPatiInfo.����id <> 0 Then
        MsgBox "�ò����Ѿ������˲����ļ��������ٵ��뷶�ġ�", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If objImportEPRDemo.ShowMe(Me, mPatiInfo.�����ļ�id, mPatiInfo.����ID, mPatiInfo.�Һ�ID, rsDemo) > 0 Then
        Call SetDocData(rsDemo, 1)
        Call SetPermitEscape(False)
        Call SetRTFEditFontSize
    End If
End Sub

Private Sub cmdSign_Click()
    Dim i As Long, str�������� As String, strSource As String, strSQL As String
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim patiSign As cEPRSign, objEPRDoc As cEPRDocument
    
    If mPatiInfo.�Ƿ�ǩ�� = False Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then
            MsgBox "�������벡����Ϣ���ٽ���ǩ����", vbInformation, gstrSysName
            Exit Sub    '����ʱ�����û�������ݣ��򲻱���
        End If
                       
        If mblnPatiChange Then
            Call ExecuteOK
            If mblnPatiChange Then Exit Sub    '����ʧ�����ټ���
        End If
                     
        If edtEditor.Text = "" Then
            If ReadRTFData(mPatiInfo.����id) = False Then Exit Sub
        End If
        
        strSource = edtEditor.Text
        '76491,δ֪BUG,���õ����㣬��δ��������±���
        If cmdSign.Visible And cmdSign.Enabled Then cmdSign.SetFocus
        Set patiSign = frmOutDocterSign.ShowMe(Me, strSource, mPatiInfo.����ID, mPatiInfo.�Һ�ID)
        If patiSign Is Nothing Then Exit Sub
        With patiSign
            .Key = "1"
            str�������� = .ǩ����ʽ & ";" & .ǩ������ & ";" & .֤��ID & ";" & IIf(.��ʾ��ǩ, 1, 0) & ";" & _
                    Format(.ǩ��ʱ��, "yyyy-mm-dd hh:mm:ss") & ";" & .��ʾʱ�� & ";" & .ǩ��Ҫ��
                    
            strSQL = "Zl_�����ﲡ��_ǩ��(1," & mPatiInfo.����id & ",'" & str�������� & "','" & UserInfo.���� & "','" & _
                    .ǰ������ & "','" & .ʱ��� & "','" & .ǩ������ & "','" & .ǩ����Ϣ & "')"
        End With
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.InsertIntoEditor(edtEditor, Len(edtEditor.Text), , objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
    Else
        If MsgBox("��ȷ��Ҫȡ��ǩ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        Set patiSign = GetSign(mPatiInfo.����id)
        If patiSign Is Nothing Then Exit Sub
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.DeleteFromEditor(edtEditor, objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
        
        strSQL = "Zl_�����ﲡ��_ǩ��(0," & mPatiInfo.����id & ")"
    End If
    
   
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        If SaveRTFData(mPatiInfo.����id, True) = False Then GoTo errH
    gcnOracle.CommitTrans: blnTrans = False
    
        
    Call LoadDocData
    Call SetPermitEdit
    Call SetPermitEscape(True)
    Call PicBasis_Resize
    With mPatiInfo
        Call mclsEPRs.zlRefresh(mPatiInfo.����ID, .�Һ�ID, mlng����ID, mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
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
Private Function GetSign(ByVal lng����ID As Long) As cEPRSign
'���ܣ���ȡ��ǰ�û���ǩ������
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim OneSign As New cEPRSign, intSign As Integer, strUserName As String
    
    strUserName = UserInfo.����
    intSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    If intSign = 1 Then
        strSQL = "Select ǩ�� From ��Ա�� Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
        If rsTemp.RecordCount > 0 Then
            If Not IsNull(rsTemp!ǩ��) Then strUserName = rsTemp!ǩ��
        End If
    End If
    strSQL = "Select Id,������ From ���Ӳ������� Where �ļ�id= [1] And ��������=8 And Instr(';'||�����ı�||';',[2])>0 Order By ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, ";" & strUserName & ";")
    If rsTemp.RecordCount > 0 Then
        OneSign.Key = Nvl(rsTemp!������, 0)
        If OneSign.GetSignFromDB(rsTemp!ID) = True Then Set GetSign = OneSign
    End If
End Function

Private Sub cmdUpdate_Click()
    Dim blnDoc As Boolean
    
    If InStr(";" & GetPrivFunc(glngSys, p���ﲡ������) & ";", ";������д;") > 0 Then
        blnDoc = mlng����ID <> 0 And mlng����ID = mPatiInfo.����ID And _
                 (mPatiInfo.����id = 0 And mPatiInfo.�����ļ�id <> 0 Or mPatiInfo.����id <> 0) And (mintActive = pt���� Or mintActive = pt����)
        If blnDoc And mPatiInfo.����id <> 0 And lblҽ��(1).Tag = "0" Then   'û���޸����˲�����Ȩ��
            blnDoc = mPatiInfo.������ = UserInfo.����
        End If
        
        If blnDoc Then
            If mobjEPRDoc Is Nothing Then
                Set mobjEPRDoc = New zlRichEPR.cEPRDocument
            End If
            If mPatiInfo.����id = 0 And mPatiInfo.�����ļ�id <> 0 Then '���û���½����½�
                Call mobjEPRDoc.InitEPRDoc(0, 2, mPatiInfo.�����ļ�id, 1, mPatiInfo.����ID, mPatiInfo.�Һ�ID, , mPatiInfo.����ID, , False)
            Else
                Call mobjEPRDoc.InitEPRDoc(1, 2, mPatiInfo.����id, 1, mPatiInfo.����ID, mPatiInfo.�Һ�ID, , mPatiInfo.����ID, , False)
            End If
            Call mobjEPRDoc.ShowEPREditor(Me)
        Else
            MsgBox "��ǰ���������޸ġ�", vbInformation, Me.Caption
        End If
    Else
        MsgBox "��û�в�����д��Ȩ�ޡ�", vbInformation, Me.Caption
    End If
End Sub



Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    '��� Pane���ڹرգ����ǵ�ǰ��Ĳ����б����ֹ�ر�
    If Action = PaneActionCollapsing And mintActive > -1 Then
        Select Case mintActive
            Case pt����, pt����, pt����
                If Pane.ID = mintActive + 1 Then
                    Cancel = True
                End If
            Case ptת��
                If Pane.ID = 4 Then
                    Cancel = True
                End If
            Case ptԤԼ
                If Pane.ID = 5 Then
                    Cancel = True
                End If
            Case pt����
                If Pane.ID = 7 Then
                    Cancel = True
                End If
            Case pt�Ŷӽк�
                If Pane.ID = pt�Ŷӽк� Then
                    Cancel = True
                End If
        End Select
    End If
End Sub

Private Sub Form_Activate()
    If Check�Ŷӽк� Then
        DoEvents
        mobjQueue.SetFocus
    End If
 
    '����ʱ��ѡ���κβ���
    'If lvwPati(pt����).Visible And lvwPati(pt����).Enabled Then Call lvwPati(pt����).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        If picSentence.Visible Then
            Call HideWordInput   '���شʾ�����
        ElseIf mblnPatiEditable And mblnPatiChange Then
            Call ExecutePaitCancel
            Call PicBasis_Resize
        End If
    End If
    '����
    PatiIdentify.ActiveFastKey
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[|']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If

    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And Not Me.ActiveControl Is PatiIdentify And mstrFindType = "���￨" And PatiIdentify.Enabled And PatiIdentify.Visible Then
        If Not (ActiveControl.Container Is PicBasis Or ActiveControl.Container Is PicOutDoc Or ActiveControl.Container Is picSentence) Then
            PatiIdentify.Text = UCase(Chr(KeyAscii))
            PatiIdentify.NotAutoSel = True
            PatiIdentify.SetFocus
        End If
    End If
End Sub

Private Sub mclsAdvices_VSKeyPress(KeyAscii As Integer)
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And mstrFindType = "���￨" And PatiIdentify.Enabled And PatiIdentify.Visible Then
        picFind.SetFocus
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        PatiIdentify.NotAutoSel = True
        PatiIdentify.SetFocus
    End If
End Sub

Private Sub InitQueuePara(Optional blnOnlyRefreshҽ���������� As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ���ŶӽкŲ���
    '���ƣ����˺�
    '���ڣ�2010-06-07 16:23:31
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    '���˺�:'�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    Dim bytType As Byte
    If blnOnlyRefreshҽ���������� Then GoTo RefreshDoctor:
    
    mty_Queue.strQueuePrivs = ";" & GetPrivFunc(glngSys, p�Ŷӽк�����ģ��) & ";"
    mty_Queue.byt�Ŷӽк�ģʽ = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, p����������))
    mty_Queue.str����վ�� = zlDatabase.GetPara("Զ�˺���վ��", glngSys, p�Ŷӽк�����ģ��)
    
RefreshDoctor:
    If mty_Queue.byt�Ŷӽк�ģʽ = 1 Then
   
        mty_Queue.blnҽ���������� = Val(zlGetLocaleComputerNamePara("�ŶӺ���վ��", glngSys, p����������, "0", mty_Queue.str����վ��)) = 1
    Else
         mty_Queue.blnҽ���������� = False
    End If
    If mty_Queue.blnҽ���������� Then
        mty_Queue.int�������� = Val(zlDatabase.GetPara("ҽ����������", glngSys, p����ҽ��վ))
    Else
        mty_Queue.int�������� = 0
    End If
    mty_Queue.bln���к����� = Val(zlDatabase.GetPara("��������������", glngSys, p����ҽ��վ, "1")) = 1
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
    Call mclsMipModule.InitMessage(glngSys, p����ҽ��վ, GetInsidePrivs(p����ҽ��վ))
    Call AddMipModule(mclsMipModule)
    
    Set mclsDisease = New zl9Disease.clsDisease
    Call mclsDisease.InitDisease(gcnOracle, Me, glngSys, glngModul, mstrPrivs, mclsMipModule)
    Set mclsDisDoc = New zl9Disease.cDockDisease
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    Call InitQueuePara
    Call InitRegist
    'ҽ������ˢ������
    mstrNotifyAdvice = zlDatabase.GetPara("�Զ�ˢ������", glngSys, p����ҽ��վ, "0")
    mintNotifyDay = Val(zlDatabase.GetPara("�Զ�ˢ�²�����������", glngSys, p����ҽ��վ, 1))
    mintNotify = Val(zlDatabase.GetPara("�Զ�ˢ�²������ļ��", glngSys, p����ҽ��վ))
    mbln��Ϣ���� = Val(zlDatabase.GetPara("����������ʾ", glngSys, p����ҽ��վ)) = 1
    mblnMaskID = Val(zlDatabase.GetPara("���֤������ʾ", glngSys)) = 1
    mblnΣ��ֵ = InStr(GetInsidePrivs(p����ҽ��վ), ";Σ��ֵ����;") > 0
    
    
    '�ȶ��������˵���������Ҫ�ж�
    blnTmp = InStr(GetInsidePrivs(p���ﲡ������), "������д") > 0
    If blnTmp Then
        lblҽ��(1).Tag = IIf(InStr(GetInsidePrivs(p���ﲡ������), "���˲���") > 0, 1, 0)
        
        mblnDocInput = Val(zlDatabase.GetPara("��ʾ�����������", glngSys, p����ҽ��վ, 0, , , intType)) = 1
        blnHave = IIf(InStr(GetInsidePrivs(1070), "ǩ��Ȩ") > 0, True, False)
        lblҽ��(1).Caption = UserInfo.����
    Else
        mblnDocInput = False
        blnHave = False
    End If
    mblnPatiDetail = Val(zlDatabase.GetPara("��ʾ������ϸ��Ϣ", glngSys, p����ҽ��վ, 0, , , intType)) = 1
    If mblnPatiDetail Then
        Set picExpand.Picture = ilexpand.ListImages("�۵�").Picture
    Else
        Set picExpand.Picture = ilexpand.ListImages("չ��").Picture
    End If
    
    cmdSign.Visible = blnHave
    lblҽ��(0).Visible = blnHave
    lblҽ��(1).Visible = blnHave
    PicOutDoc.Visible = mblnDocInput
    mArrDate = Array(txt��������, txt����ʱ��, txt��������, txt����ʱ��)
    
    'һ��ͨ������ʼ������tbcSub_SelectedChanged֮ǰ���Ա㴫�ݸ�ҽ������
     'zlGetIDKindStr�л��Զ�����Ϊ����8λ����
    mstrCardKind = "��|���￨|0|0|8|0|0|0;��|��ʶ��|0|0|0|0|0|0;��|�Һŵ�|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0;�ɣ�|�ɣÿ�|1|0|0|0|0|0"
    If Check�Ŷӽк� = True Then mstrCardKind = mstrCardKind & ";��|�ŶӺ�|0|0|0|0|0|0;ҽ|ҽ����|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    If Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
        Else
            mstrCardKind = mobjSquareCard.zlGetIDKindStr(mstrCardKind)
        End If
    End If
    Call PatiIdentify.zlInit(Me, glngSys, p����ҽ��վ, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
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
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
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
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 380, 550, DockLeftOf, Nothing)
    objPane.Title = "���ﲡ��"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 280, 190, DockBottomOf, objPane)
    objPane.Title = "���ﲡ��"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(7, 280, 250, DockBottomOf, objPane)
    objPane.Title = "���ﲡ��"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(3, 280, 550, DockBottomOf, objPane)
    objPane.Title = "���ﲡ��"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(8, 280, 180, DockBottomOf, objPane)
    objPane.Title = "��Ϣ����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(4, 380, 350, DockTopOf, dkpMain.Panes(1))
    objPane.Title = "ת�ﲡ��"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    objPane.AttachTo dkpMain.Panes(1)
    Set objPane = Me.dkpMain.CreatePane(5, 380, 550, DockTopOf, dkpMain.Panes(1))
    objPane.Title = "ԤԼ����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    objPane.AttachTo dkpMain.Panes(1)
     
     
    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p�°����ﲡ��, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "���Ӳ���")
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
        mcolSubForm.Add mclsEMR.zlGetForm, "_�²���"
    End If
    mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_����"
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        If GetInsidePrivs(p����ҽ���´�) <> "" Then
            '�ȼ���ҽ����ԭ��:�����������ӿڣ����ͻ���û����������ʱ������ȼ����Ŷӽкź����ҽ����ʱ��
            '�ӡ�������Ϣ���л�����ҽ����Ϣ�����򵯳�Msgbox���� �����:67995
            .InsertItem(intIdx, "ҽ����Ϣ", mcolSubForm("_ҽ��").hwnd, 0).Tag = "ҽ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p���ﲡ������) <> "" Then
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�°����ﲡ��, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "���Ӳ���", picTmp.hwnd, 0).Tag = "�²���": intIdx = intIdx + 1
        End If
        '����ṩ�Ŀ�Ƭ
        Call CreatePlugInOK(p����ҽ��վ)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p����ҽ��վ)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p����ҽ��վ, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        If .ItemCount = 0 Then
            MsgBox "��û��ʹ������ҽ������վ��Ȩ�ޡ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = zlDatabase.GetPara("ҽ������", glngSys, p����ҽ��վ)
        
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '���⼤���¼�
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
        End If
        'ֻ����ѡ����Ӵ���
        Call tbcSub_SelectedChanged(.Selected)
    End With
            
    '��ȡ��������
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
    
    Call GetLocalSetting '���ز���
    
    Call InitCondFilter '���ﲡ�˹�������
    Call InitReportColumn
    Call InitPatiData
    Call LoadPatients '��ʾ����
    Call LoadNotify '��Ϣ����
    
    dkpMain(4).Hidden = True     'ȱʡ����ʾ���ﲡ������
    
    '�ŵ�ҽ�����棬����Ϊҽ���к�����ҩ������������ڣ�������ʾ��󣬻ᵼ���ŶӽкŴ��ڽ��á�
    If Check�Ŷӽк� = True Then
        '����Ƿ�����Ŷӽк�
        Set objPane = Me.dkpMain.CreatePane(6, 380, 550, DockTopOf, dkpMain.Panes(1))
        objPane.Title = "�Ŷӽк�"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable
        objPane.AttachTo dkpMain.Panes(1)
        'mobjQueue.zlSetToolIcon 24, True
    End If
    
    '����ָ�:�������ִ��
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        '��ָ�Panne�ı���,Tag�����
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
    
    '����һ�����վ�������������֮���ٰ󶨾��
    For i = 1 To Me.dkpMain.PanesCount
        If Me.dkpMain.Panes(i).ID = 4 Then
            Me.dkpMain.Panes(i).Handle = lvwIncept.hwnd '����ʱû�м���AttachPane����,����ǿ�и�ֵ
        ElseIf Me.dkpMain.Panes(i).ID = 5 Then
            Me.dkpMain.Panes(i).Handle = lvwReserve.hwnd '����ʱû�м���AttachPane����,����ǿ�и�ֵ
        ElseIf Me.dkpMain.Panes(i).ID = pt�Ŷӽк� Then
            Me.dkpMain.Panes(i).Handle = mobjQueue.zlGetForm.hwnd '����ʱû�м���AttachPane����,����ǿ�и�ֵ
        ElseIf Me.dkpMain.Panes(i).ID = 7 Then  '����
            Me.dkpMain.Panes(i).Handle = lvwPatiHZ.hwnd
        ElseIf Me.dkpMain.Panes(i).ID = 3 Then  '����
            Me.dkpMain.Panes(i).Handle = picYZ.hwnd
        ElseIf Me.dkpMain.Panes(i).ID = 8 Then
            Me.dkpMain.Panes(i).Handle = rptNotify.hwnd
        Else
            Me.dkpMain.Panes(i).Handle = lvwPati(Me.dkpMain.Panes(i).ID - 1).hwnd
        End If
    Next
    
    dkpMain.Panes(1).Select
    
    '����ȱʡ���ҷ�ʽ
    arrType = Split(mstrCardKind, ";")
    For i = 1 To UBound(arrType) + 1
        If i = mintFindType Then
            PatiIdentify.objIDKind.IDKind = i
            Exit For
        End If
    Next
    
    
    '������������
    picPatiInput.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    Call zlControl.CboSetWidth(cboRegist.hwnd, cboRegist.Width * 1.1)
    Call zlControl.CboSetWidth(cboEdit(cboְҵ).hwnd, cboRegist.Width * 3)
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    Call RefreshTitle
    If Check�Ŷӽк� = True Then
       '����Ƿ�����Ŷӽк�
       Call ReshDataQueue
       For i = 1 To dkpMain.PanesCount
           If dkpMain.Panes(i).Title Like "*�Ŷӽк�*" Then
               dkpMain.Panes(i).Select: Exit For
           End If
       Next
    End If
    Call RefreshPass

    If ISPassShowCard Then Call Hide���￨����
    ucPatiVitalSigns.LabToTxt = -20
    ucPatiVitalSigns.XDis = 100
    mblnUnRefresh = False
End Sub

Private Sub RefreshPass()
    '�Ƿ����̫Ԫͨ�ӿڲ���
    mblnUseTYT = False
    If gbytPass = 3 Then
        If gint����������Դ = 0 Then
            mint����������Դ = Val(zlDatabase.GetPara("����������Դ", glngSys, p����ҽ��վ, "0"))
        End If
        mblnUseTYT = gint����������Դ = 0 And mint����������Դ = 1 Or gint����������Դ = 2
    End If
    '����̫Ԫͨ�ӿڶ��󣬴���ʧ�ܣ�������̫Ԫͨ
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
    Dim str�Һŵ� As String, strCardNO As String
    Dim rsTmp As Recordset
    Dim str����ID As String, str���ID As String
    Dim intFindTypeTmp As Integer
    Dim strPictureFile As String
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
 
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
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
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S 'С����
        If mbytSize <> 0 Then
            mbytSize = 0
            Call zlDatabase.SetPara("����", mbytSize, glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '������
        If mbytSize <> 1 Then
            mbytSize = 1
            Call zlDatabase.SetPara("����", mbytSize, glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Find '����
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '��ʱ��Ҫ��λһ��
            If PatiIdentify.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            PatiIdentify.SetFocus
        End If
    Case conMenu_View_FindNext '������һ��
        If PatiIdentify.Text = "" And mstrIDCard = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True, IIf(PatiIdentify.Text = "", mstrIDCard, ""))
        End If
    Case 3564 'ԤԼ�Ǽ�
        Call gobjVitualExpense.zlExecuteCommandBars(Me, Control, str�Һŵ�, mlng����ID)
    Case conMenu_Edit_AppRequest
        Call gobjVitualExpense.zlExecuteCommandBars(Me, Control, str�Һŵ�, mlng����ID)
    Case conMenu_Edit_OpenArrangement
        Call gobjVitualExpense.zlOpenStopedPlanBySN(Me, p����ҽ��վ, , , UserInfo.ID)
    Case conMenu_View_PatInfor  '��ʾ������ϸ��Ϣ
        mblnPatiDetail = Not mblnPatiDetail
        Call zlDatabase.SetPara("��ʾ������ϸ��Ϣ", IIf(mblnPatiDetail, 1, 0), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
        Call picExpand_Click
        
    Case conMenu_View_PatiInput  '��ʾ����������
        mblnDocInput = Not mblnDocInput
        Call zlDatabase.SetPara("��ʾ�����������", IIf(mblnDocInput, 1, 0), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
        
        PicOutDoc.Visible = mblnDocInput
        Call cbsMain_Resize
        
        If mblnDocInput Then Call LoadDocData   '����ʾʱû�ж�ȡ
        Call SetPermitEdit
        
        Call PicBasis_Resize
        Call PicPatiInfo_Resize
                
    Case conMenu_View_Busy '����״̬
        Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
    Case conMenu_View_Refresh 'ˢ��
        Call LoadPatients("110111")
        Call LoadNotify
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_File_Parameter '��������
        frmOutStationSetup.mstrPrivs = mstrPrivs
        frmOutStationSetup.Show 1, Me

        If gblnOK Then
            intFindTypeTmp = mintFindType
            Call GetLocalSetting
            mintFindType = intFindTypeTmp
            Call LoadPatients
            Call InitQueuePara
        End If
        If Check�Ŷӽк� Then
            Call ReshDataQueue
        End If
    Case conMenu_Tool_KssAudit '������ҩ���
        On Error Resume Next
        Call frmKSSExamine.Show(0, Me)
    Case conMenu_Tool_CISMed  '�ٴ��Թ�ҩ
        Call Set�ٴ��Թ�ҩ(Me)
     Case conMenu_Tool_TransAudit '��Ѫ��˹���
        On Error Resume Next
        Call frmTransfuseExamine.ShowMe(Me, 1)
    Case conMenu_Tool_Archive '���Ӳ�������
        Call frmArchiveView.ShowArchive(Me, mPatiInfo.����ID, mPatiInfo.�Һ�ID)
    Case conMenu_Tool_Reference_1 '������ϲο�
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '������Ŀ��������
        Call Set������Ŀ��������
        
    Case conMenu_Tool_Community * 100# + 1 '���������֤
        Call ExecuteCommunityIdentify
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '������������
        If Not gobjCommunity Is Nothing And mPatiInfo.���� <> 0 And mPatiInfo.�Һ�ID <> 0 Then
            If gobjCommunity.CommunityFunc(glngSys, mlngModul, Val(Control.Parameter), mPatiInfo.����, mPatiInfo.������, mPatiInfo.����ID, mPatiInfo.�Һ�ID) Then
                Call LoadPatients
            End If
        End If
    Case conMenu_Tool_MedRec '������ҳ
        If mclsInOutMedRec Is Nothing Then
            Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
            Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, p����ҽ��վ, mclsMipModule, gobjCommunity, gclsInsure)
        End If
        If mclsInOutMedRec.ShowOutMedRecEdit(Me, mPatiInfo.�Һŵ�, mstrPrivs, IIf(mstr�Һŵ� = mPatiInfo.�Һŵ� And (mPatiInfo.���� = pt���� Or mPatiInfo.���� = pt����), 0, 1), strPictureFile) Then
'            If strPictureFile <> "" And strPictureFile <> "0" Then
'                Call ReadPatPricture(mlng����ID, imgPatient, strPictureFile)
'                picPatient.Visible = True
'            ElseIf strPictureFile = "" Then
'                picPatient.Visible = False
'            End If
'            Call LoadPatients("110")
        End If
        Call RefreshPass
    Case conMenu_File_MedRecSetup '��ҳ��ӡ����
        'Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me)
    Case conMenu_File_MedRecPreview '��ҳԤ��
        'Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me, "����ID=" & mlng����ID, "NO=" & mPatiInfo.�Һŵ�, 1)
    Case conMenu_File_MedRecPrint '��ҳ��ӡ
        'Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me, "����ID=" & mlng����ID, "NO=" & mPatiInfo.�Һŵ�, 2)
    Case conMenu_Manage_Regist '���˹Һ�
        Control.Enabled = False
        Call ExecuteRegist
        Control.Enabled = True
    Case conMenu_Manage_Bespeak 'ԤԼ�Һ�
        Control.Enabled = False
        Call ExecuteBespeak
        Control.Enabled = True
    Case conMenu_File_Print_Bespeak '�ش�ԤԼ�Һŵ�
        Control.Enabled = False
        Call ExecuteBespeakPrint
        Control.Enabled = True
    Case conMenu_Manage_Transfer_Send '����ת��
        Call ExecuteTransferSend
    Case conMenu_Manage_Transfer_Cancel 'ȡ��ת��
        Call ExecuteTransferCancel
    Case conMenu_Manage_Transfer_Incept '����ת��
        Call ExecuteTransferIncept
    Case conMenu_Manage_Transfer_Refuse 'ת��ܾ�
        Call ExecuteTransferRefuse
    Case conMenu_Manage_Transfer_Force 'ǿ������
        str�Һŵ� = frmForceGet.ShowMe(Me, mstrPrivs, mlng�������ID, mobjSquareCard)
        If str�Һŵ� <> "" Then
            If lvwPati(pt����).Visible Then
                Call LoadPatients("110011", pt����, str�Һŵ�)
                lvwPati(pt����).SetFocus
            Else
                Call LoadPatients("110011")
            End If
        End If
    Case conMenu_Manage_Receive '���˽���
        Call ExecuteReceive
    Case conMenu_Manage_Cancel 'ȡ������
        Call ExecuteCancel
    Case conMenu_Manage_Finish '��ɽ���
        Call ExecuteFinish
    Case conMenu_Manage_Redo '�ָ�����
        Call ExecuteRedo
    Case conMenu_Manage_ReBack '��ͣ����
          Call ExecuteStopAndReuse(False)
    Case conMenu_Manage_ReBackCancel '�ָ���ͣ����
          Call ExecuteStopAndReuse(True)
    Case conMenu_Edit_Transf_Save   '���没����Ϣ
        Call ExecuteOK
        Call HideWordInput
        Call PicBasis_Resize
    Case conMenu_Edit_Transf_Cancle 'ȡ��������Ϣ
        Call ExecutePaitCancel
        Call HideWordInput
        Call PicBasis_Resize
   Case conmenu_View_Leave  '��ʾ�����ﲡ��
         mblnShowLeavePati = Not mblnShowLeavePati
         Control.Checked = mblnShowLeavePati
        Call LoadPatients("10000")
    Case conmenu_Edit_Leave     '���˲�����
        If Set���˹Һ�״̬(-1) Then
            Call LoadPatients("10000")
            Call ReshDataQueue
        End If
    Case conmenu_Edit_Wait      '���˾���
        If Set���˹Һ�״̬(0) Then
            Call LoadPatients("10000")
            Call ReshDataQueue
        End If
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_Tool_HealthCard  '���񽡿���
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlHealthArchivesShow(Me, p����ҽ��վ, mlng����ID, "")
        End If
    Case conMenu_Edit_TraReactionRecord '��Ѫ��Ӧ
        Call FuncTraReactionRecord(Me, 0, p����ҽ���´�)
    Case conMenu_Tool_Positive '���Խ���鿴
        i = GetOne���Խ��
        If i <> 0 Then Call mclsDisease.ShowDisRegist(Me, 1, i, mlng����ID, 0, mstr�Һŵ�)
    Case conMenu_Tool_Critical 'Σ��ֵ�鿴����
        Call ExecuteCritical
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            With mPatiInfo
                If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    If mlng�������ID = 0 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                    Else
                        Set rsTmp = zlDatabase.OpenSQLRecord("Select ���� From ���ű� Where ID=[1]", Me.Caption, mlng�������ID)
                        If rsTmp.EOF Then Exit Sub
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "��������=" & rsTmp!���� & "|=" & mlng�������ID)
                    End If
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "����ID=" & .����ID, "�����=" & .�����, "�Һŵ�=" & .�Һŵ�, "����=" & .����)
                End If
            End With
        Else
            If Check�Ŷӽк� = True Then
                mobjQueue.zlExecuteCommandBars Control
            End If
            Select Case Me.tbcSub.Selected.Tag
            Case "ҽ��"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "�²���"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, p����ҽ��վ, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mPatiInfo.����ID, 0, mPatiInfo.�Һŵ�)
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
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Send, "ת�ﲡ��(&S)", -1, False)
                objControl.IconId = conMenu_Manage_Transfer
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Cancel, "ȡ��ת��(&C)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Incept, "ת�����(&I)", -1, False)
                objControl.IconId = conMenu_Manage_Receive
                objControl.BeginGroup = True
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Refuse, "ת��ܾ�(&R)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "ǿ������(&F)", -1, False)
                objControl.BeginGroup = True
            End If
        End With
    Case conMenu_Tool_Community '��������
        mlngCommunityID = 0
        With CommandBar.Controls
            .DeleteAll
            If Not gobjCommunity Is Nothing Then
                '������֤
                If mPatiInfo.���� = 0 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Tool_Community * 100# + 1, "�����֤(&V)")
                End If
                
                '��������
                If mPatiInfo.���� <> 0 Then
                    strFunc = gobjCommunity.GetCommunityFunc(glngSys, p����ҽ��վ, mPatiInfo.����)
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
       Case "ҽ��"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "����"
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
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case 3564
        Control.Enabled = InStr(mstrVirutalPrivs, ";ԤԼ�Ǽ�;") > 0
    Case conMenu_Edit_AppRequest
        Control.Enabled = InStr(mstrVirutalPrivs, ";ԤԼ�Ǽ�;") > 0
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_FontSize_S 'С����
        Control.Checked = Not (mbytSize = 1)
    Case conMenu_View_FontSize_L '������
        Control.Checked = (mbytSize = 1)
    Case conMenu_View_PatiInput '��ʾ����������
        Control.Checked = mblnDocInput
    Case conMenu_View_PatInfor      '��ʾ������ϸ��Ϣ
        Control.Checked = mblnPatiDetail
        
    Case conMenu_View_Busy '����״̬
        Control.Checked = lblRoom.BackColor = COLOR_BUSY
    Case conMenu_Tool_KssAudit  '������ҩ���
        If GetInsidePrivs(p������ҩ���) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_TransAudit '��Ѫ�ּ�����
        If GetInsidePrivs(p��Ѫ��˹���) = "" Or Not gbln��Ѫ�ּ����� Then
            Control.Visible = False
        End If
    Case conMenu_Tool_CISMed  '�ٴ��Թ�ҩ
        If InStr(GetInsidePrivs(p����ҽ��վ), ";�ٴ��Թ�ҩ;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_Archive '���Ӳ�������
        If GetInsidePrivs(p���Ӳ�������) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    Case conMenu_Tool_HealthCard  '���񽡿���
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Tool_Reference_1 '������ϲο�
        If GetInsidePrivs(p������ϲο�) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 'ҩƷ�����Ʋο�
        If GetInsidePrivs(pҩƷ���Ʋο�) = "" Then Control.Visible = False
    Case conMenu_Tool_Community '�����˵�
        If gobjCommunity Is Nothing Then
            Control.Visible = False
        End If
    Case conMenu_Edit_TraReactionRecord '��Ѫ��Ӧ
        Control.Visible = InStr(1, GetInsidePrivs(9005, , 2200), "��Ѫ��Ӧ�Ǽ�") <> 0
        Control.Enabled = Control.Visible And gblnѪ��ϵͳ
    Case conMenu_Manage_FeeItemSet '������Ŀ��������,û��Ȩ��ʱ�ɲ鿴
                
    Case conMenu_Tool_Community * 100# + 1 '���������֤
        Control.Enabled = mlng����ID <> 0 And mPatiInfo.���� = 0 And (mPatiInfo.���� = pt���� Or mPatiInfo.���� = pt����) And InStr(mstrPrivs, "���˽���") > 0
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '������������
        Control.Enabled = mlng����ID <> 0 And mPatiInfo.���� <> 0
    Case conMenu_Tool_MedRec '������ҳ
        If InStr(mstrPrivs, "������ҳ") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    Case conMenu_File_MedRec '��ҳ��ӡ
        If InStr(mstrPrivs, "��ӡ��ҳ") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    
    Case conMenu_ManagePopup '������˵�
        If InStr(mstrPrivs, ";���˽���;") = 0 Then Control.Visible = False
    Case conMenu_Manage_Regist '���˹Һ�
        If InStr(mstrPrivs, ";���˹Һ�;") = 0 Then Control.Visible = False
    Case conMenu_Manage_Bespeak 'ԤԼ�Һ�
        If InStr(mstrPrivs, ";ԤԼ�Һ�;") = 0 Then Control.Visible = False
    Case conMenu_Edit_OpenArrangement
        If InStr(mstrPrivs, ";ԤԼ�Һ�;") = 0 And InStr(mstrPrivs, ";���˹Һ�;") = 0 And InStr(mstrVirutalPrivs, ";ԤԼ�Ǽ�;") = 0 Then Control.Visible = False
    Case conMenu_File_Print_Bespeak
      Control.Visible = InStr(mstrPrivs, ";ԤԼ�Һŵ�;") > 0 And lvwReserve.Visible     '56274
      Control.Enabled = lvwReserve.Visible And Not lvwReserve.SelectedItem Is Nothing
    Case conMenu_Manage_Transfer 'ת�ﴦ��
        If InStr(mstrPrivs, "���˽���") = 0 _
            And InStr(mstrPrivs, "����ת��") = 0 _
                And InStr(mstrPrivs, "���ﲡ��") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Manage_Transfer_Send '����ת��
        If InStr(mstrPrivs, "����ת��") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt���� Or mintActive = pt����)
            If blnEnabled Then
                blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
                If blnEnabled Then
                    'Ŀǰ����"��ת��/���ѽ���"״̬
                    With lvwPati(mintActive).SelectedItem.ListSubItems(5)
                        blnEnabled = .Tag = "" Or Val(.Tag) = 1
                    End With
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Manage_Transfer_Cancel 'ȡ��ת��
        If InStr(mstrPrivs, "����ת��") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt���� Or mintActive = pt����)
            If blnEnabled Then
                blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
                If blnEnabled Then
                    'Ŀǰ����ת��"������/�Ѿܾ�"״̬
                    With lvwPati(mintActive).SelectedItem.ListSubItems(5)
                        blnEnabled = Val(.Tag) = 0 And .Tag <> "" Or Val(.Tag) = -1
                    End With
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conmenu_View_Leave  '��ʾ�����ﲡ��
            Control.Checked = mblnShowLeavePati
            'Control.Enabled = (mintActive = pt����)
    Case conmenu_Edit_Leave
            blnEnabled = (mintActive = pt����)
            If blnEnabled Then
                blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
                If blnEnabled Then
                    'ֻ�������ĺ��ﲡ�˲ſ���ȡ������
                    With lvwPati(mintActive).SelectedItem.ListSubItems(9)
                        blnEnabled = Val(.Tag) = 0
                    End With
                End If
            End If
            Control.Enabled = blnEnabled
    Case conmenu_Edit_Wait
        blnEnabled = mintActive = pt����
        If blnEnabled Then
            blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
            If blnEnabled Then
                'Ŀǰ����ת��"������/�Ѿܾ�"״̬
                With lvwPati(mintActive).SelectedItem.ListSubItems(9)
                    blnEnabled = Val(.Tag) = -1
                End With
            End If
        End If
        Control.Enabled = blnEnabled
        
    Case conMenu_Manage_Transfer_Incept, conMenu_Manage_Transfer_Refuse 'ת�����,ת��ܾ�
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            'ת���б��в��ˣ��Ҵ���"������"״̬
            blnEnabled = Not lvwIncept.SelectedItem Is Nothing And lvwIncept.Visible
            Control.Enabled = blnEnabled
        End If
        
    Case conMenu_Manage_Transfer_Force 'ǿ������
        If InStr(mstrPrivs, "���˽���") = 0 Or InStr(mstrPrivs, "���ﲡ��") = 0 Then Control.Visible = False
    Case conMenu_Manage_ReBack '��ͣ����:�����
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            If Not lvwPati(pt����).SelectedItem Is Nothing And mintActive = pt���� And lvwPati(pt����).Visible Then
                '0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
                Control.Enabled = Val(lvwPati(pt����).SelectedItem.ListSubItems(8).Tag) < 2
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Manage_ReBackCancel '�ָ���ͣ����
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            If Not lvwPatiHZ.SelectedItem Is Nothing And lvwPatiHZ.Visible And mintActive = pt���� Then
                ' 0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
                Control.Enabled = Val(lvwPatiHZ.SelectedItem.ListSubItems(8).Tag) = 2
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Manage_Receive '���˽���
        If InStr(mstrPrivs, "���˽���") = 0 Or (mty_Queue.blnҽ���������� And mbln���к����) Then
            Control.Enabled = False
            Control.Visible = False
        Else
            Control.Visible = True
            '���ԤԼ�ҺŲ��˿���ֱ�ӽ��ת�ﲡ�˲�ͨ���������
            blnEnabled = False
            If lvwPati(pt����).Visible And lvwReserve.Visible Then
                blnEnabled = mintActive = pt���� And Not lvwPati(pt����).SelectedItem Is Nothing And Me.ActiveControl Is lvwPati(pt����) _
                    Or Not lvwReserve.SelectedItem Is Nothing And Me.ActiveControl Is lvwReserve
            ElseIf lvwPati(pt����).Visible Then
                blnEnabled = mintActive = pt���� And Not lvwPati(pt����).SelectedItem Is Nothing
            ElseIf lvwReserve.Visible Then
                blnEnabled = mintActive = ptԤԼ And Not lvwReserve.SelectedItem Is Nothing
            End If
            Control.Enabled = blnEnabled    '�������жϵ�ǰ�Ƿ�Ϊת�ﲡ���б���Ϊ�����ת���б�Ļ���blnEnabled�Ѿ���False
             
        End If
    Case conMenu_Manage_Cancel 'ȡ������
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mintActive = pt���� And Not lvwPati(pt����).SelectedItem Is Nothing And lvwPati(pt����).Visible
        End If
    Case conMenu_Manage_Finish '��ɾ���
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        ElseIf mintActive = pt���� Then
            Control.Enabled = Not lvwPati(pt����).SelectedItem Is Nothing And lvwPati(pt����).Visible
        ElseIf mintActive = pt���� Then
            Control.Enabled = Not lvwPatiHZ.SelectedItem Is Nothing And lvwPatiHZ.Visible
        Else
            Control.Enabled = False
        End If
    Case conMenu_Manage_Redo '�ָ�����
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mintActive = pt���� And Not lvwPati(pt����).SelectedItem Is Nothing And lvwPati(pt����).Visible
            If blnEnabled Then 'ֻ�ָܻ��������ѵĲ���(������Ȩ�޿���ǿ������)
                blnEnabled = lvwPati(pt����).SelectedItem.ListSubItems(4).Tag = UserInfo.����
            End If
            Control.Enabled = blnEnabled
        End If
    Case Else
        '60075:������,2013-04-03,���ⲿ��ҽ����ӡ��Ԥ���˵��Ĵ�����ֲ���˴�,��ǰ�ķ�ʽ�����޷���������ģ��ĸ����¼�
        If (Control.ID = conMenu_File_Print Or Control.ID = conMenu_File_Preview Or Control.ID = conMenu_Help_Help) Then
            If tbcSub.Selected.Tag = "ҽ��" Then
                Control.Visible = False
                Exit Sub
            Else
                Control.Visible = True
            End If
        End If
        If Check�Ŷӽк� Then mobjQueue.zlUpdateCommandBars Control
        Select Case tbcSub.Selected.Tag
        Case "ҽ��"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "�²���"
            Call mclsEMR.zlUpdateCommandBars(Control)
        End Select
        '������ҩ����
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
'���ܣ�ˢ���Ӵ���˵���������
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    
    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If
    
    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hwnd)
        
    Me.Caption = "����ҽ������վ - " & objItem.Caption & "(��ǰ�û���" & UserInfo.���� & ")"
    
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '���������¼���
    Call MainDefCommandBar
'    If Check�Ŷӽк� And dkpMain.Panes(pt�Ŷӽк�).Selected Then
'        mobjQueue.zlDefCommandBars cbsMain
'    End If

    '�Ӵ������¼���
    Select Case objItem.Tag
    Case "ҽ��"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 0, gobjPlugIn, mobjSquareCard)
    Case "����"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "�²���"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag)
            Call zlPlugInErrH(err, "GetButtomName")
            '�����˵�
            If strName <> "" Then Call PlugInInSideBar(cbsMain, strName)
            err.Clear: On Error GoTo 0
        End If
    End Select
    
    
    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
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
    
    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ������ݼ�״̬
    If mlng����ID = 0 Or (mintActive = pt���� And mPatiInfo.�Һŵ� = mstr�Һŵ�) Then
        '�����ԤԼ���ˣ����ξ���û��ҽ���Ͳ�������
        'Ҫ���Ӵ��尴�����ݴ������
        Select Case objItem.Tag
        Case "ҽ��"
            Call mclsAdvices.zlRefresh(0, "", False)
        Case "����"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "�²���"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 1)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
            Select Case objItem.Tag
            Case "ҽ��"
                Call mclsAdvices.zlRefresh(.����ID, .�Һŵ�, mstr�Һŵ� = .�Һŵ� And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, , , mclsMipModule)
            Case "����"
                Call mclsEPRs.zlRefresh(.����ID, .�Һ�ID, mlng����ID, mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
            Case "�²���"
                Call mclsEMR.zlRefresh(.����ID, .�Һ�ID, mlng����ID, .����, 1)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng����ID, mstr�Һŵ�, 0, .����ת��, 0, 0)
                    Call zlPlugInErrH(err, "RefreshForm")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End With
    End If
    Call SetFontSize(Not Me.Visible)
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim strFunName As String

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��") '����
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_File_MedRec, "��ҳ��ӡ(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_File_MedRecSetup, "��ӡ����(&S)", -1, False
            .Add xtpControlButton, conMenu_File_MedRecPreview, "��ӡԤ��(&V)", -1, False
            .Add xtpControlButton, conMenu_File_MedRecPrint, "��ӡ��ҳ(&P)", -1, False
        End With
        '56274
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_Bespeak, "�ش�ԤԼ�Һŵ�(&P)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "����(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "���˹Һ�(&H)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ�Һ�(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AppRequest, "ԤԼ�Ǽ�(&A)")
        Set objControl = .Add(xtpControlButton, 3564, "ԤԼ�Ǽǹ���(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OpenArrangement, "����ͣ�ﰲ��(&P)")
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Leave, "���˲�����(&L)", -1, False): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Wait, "���˴���(&W)", -1, False)
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Transfer, "ת�ﴦ��(&C)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "���˽���(&Z)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "ȡ������(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "��ɽ���(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Redo, "�ָ�����(&R)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "�����(&S)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "ȡ������(&R)")
        objControl.IconId = conMenu_Edit_Reuse
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "�����С(&N)") '����
        With objPopup.CommandBar.Controls
             .Add xtpControlButton, conMenu_View_FontSize_S, "С����(&S)", -1, False '����(С�����ӦС��Ƭ���������Ӧ��Ƭ)
             .Add xtpControlButton, conMenu_View_FontSize_L, "������(&L)", -1, False '����
        End With
        objPopup.BeginGroup = True

        If InStr(GetInsidePrivs(p���ﲡ������), "������д") > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_View_PatiInput, "��ʾ�����������(&I)")
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "��ʾ������ϸ��Ϣ(&D)")
        Set objControl = .Add(xtpControlButton, conmenu_View_Leave, "��ʾ�����ﲡ��(&4)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Busy, "����æ(&M)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "������ת(&J)")
        
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Community, "��������(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_KssAudit, "������ҩ���(&K)")
        objControl.IconId = 3551
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "��Ѫ��˹���(&M)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_CISMed, "�ٴ��Թ�ҩ(&J)")
        objControl.IconId = 3901
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)"): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
        
        If gblnѪ��ϵͳ = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReactionRecord, "��Ѫ��Ӧ��¼"): objControl.BeginGroup = True
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Positive, "���Խ��")
            objControl.IconId = 3551
        If mblnΣ��ֵ Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Critical, "Σ��ֵ")
                objControl.IconId = 4113
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "������Ŀ��������(&C)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "������ҳ(&M)")
        objControl.BeginGroup = True
        objControl.ToolTipText = "�༭��鿴��ҳ��Ϣ"
        On Error Resume Next
            If mobjSquareCard.zlHealthArchiveIsSHow(Me, p����ҽ��վ, strFunName, "") Then
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
    
    '����������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�"): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlPopup, conMenu_Manage_Transfer, "ת��")
        objPopup.ID = conMenu_Manage_Transfer
        objPopup.IconId = conMenu_Manage_Transfer
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "�����")
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "ȡ��")
        objControl.IconId = conMenu_Edit_Reuse
        objControl.ToolTipText = "ȡ������"
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "��ҳ")
        objControl.BeginGroup = True
        objControl.ToolTipText = "�༭��鿴��ҳ��Ϣ"
        Set objPopup = .Add(xtpControlPopup, conMenu_Tool_Community, "����")
        objPopup.ID = conMenu_Tool_Community
        objPopup.IconId = conMenu_Tool_Community
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "��Ѫ���")
        objControl.IconId = 3551
                If strFunName <> "" Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                objControl.ToolTipText = strFunName
                objControl.IconId = 3208
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����")
        objControl.BeginGroup = True
        objControl.Enabled = False
        objControl.ToolTipText = "���没�������Ϣ���޸�"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "����")
        objControl.Enabled = False
        objControl.ToolTipText = "�������������Ϣ���޸ģ�ESC��"
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����") '����
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With

    '���������⴦��
    '-----------------------------------------------------
    '���˵��Ҳ�Ĳ���
    With cbsMain.ActiveMenuBar.Controls
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = picTmphwnd.hwnd
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyH, conMenu_Manage_Regist '�Һ�
        .Add 0, vbKeyF7, conMenu_Manage_Receive '����
        .Add 0, vbKeyF8, conMenu_Manage_Finish '��ɾ���
        .Add FCONTROL, vbKeyB, conMenu_View_Busy '����״̬
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, vbKeyF, conMenu_View_Find '���Ҳ���
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        .Add 0, vbKeyF12, conMenu_File_Parameter '��������
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF6, conMenu_View_Jump '��ת
        .Add 0, vbKeyF1, conMenu_Help_Help '����
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save '����
    End With
    
    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
'        .AddHiddenCommand conMenu_File_Excel '�����Excel
'        .AddHiddenCommand conMenu_View_Jump '��ת
    End With
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1260_2")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mblnPatiChange
End Sub

Private Sub InitRegist()
    '��ʼ���Һ�
    Set gobjVitualExpense = New clsRegist
    gobjVitualExpense.zlInitCommon glngSys, gcnOracle, gstrDBUser
    gobjVitualExpense.zlInitData 1
    mstrVirutalPrivs = GetPrivFunc(glngSys, 9000)
End Sub

Private Sub lblҽ��_Click(Index As Integer)
    lbl��������.Visible = Not lbl��������.Visible
End Sub

Private Sub lvwIncept_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwIncept, ColumnHeader.Index)
End Sub
Private Sub lvwIncept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LvwItemClick(ptת��, Item)
End Sub

Private Sub lvwReserve_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwReserve, ColumnHeader.Index)
End Sub

Private Sub lvwReserve_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LvwItemClick(ptԤԼ, Item)
End Sub

Private Sub mclsAdvices_Activate()
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByRef blnNo As Boolean)
'���ܣ���������뼲������ �õ������༭��
'      blnOnChek    �Ƿ�ֻ���д�Ⱦ�����濨��д���
'      str����ID    ����ID
'      str���ID   ���ID
'blnNO �Ƿ�Ҫ��д��Ⱦ�����濨
    Call OpenEPRDoc(mobjEPRDoc, Me, mPatiInfo.����ID, mPatiInfo.�Һ�ID, mPatiInfo.����ID, str����ID, str���ID, 1, , False, blnOnChek, blnNo)
End Sub

Private Sub mclsAdvices_EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean)
'���ܣ�Ҫ�������������
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, p����ҽ��վ, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    If Not mclsInOutMedRec.ShowOutMedRecEdit(ParentForm, �Һŵ�, mstrPrivs) Then
        Succeed = False
    Else
        Succeed = mclsInOutMedRec.IsDiagInput
    End If
End Sub

Private Sub mclsAdvices_RequestRefresh()
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    Call LoadPatients
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclsAdvices_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
'���ܣ����༭��ʽ��ӡ����
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr���Ʊ���, ����ID, Not Preview, True)
End Sub

Private Sub mclsAdvices_ViewPACSImage(ByVal ҽ��ID As Long)
'���ܣ�PACS��Ƭ����
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(ҽ��ID, Me, mPatiInfo.����ת��)
    End If
End Sub

Private Sub SetPermitEscape(ByVal blnOk As Boolean)
    Dim i As Long
        
    If blnOk Then
        'Call SetPermitEscape(true)֮ǰ�����ȵ�SetPermitEdit
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
            
    blnDo = mlng����ID <> 0 And mlng����ID = mPatiInfo.����ID And InStr(mstrPrivs, "������ҳ") > 0 And (mintActive = pt���� Or mintActive = pt����)
    blnBasis = InStr(mstrPrivs, "�޸Ļ�����Ϣ") > 0
        
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
        If i = txt��λ���� Then
            txtEdit(i).Locked = Not blnDo
            If blnDo And Val("" & txtEdit(txt��λ����).Tag) <> 0 Then
                txtEdit(i).Locked = InStr(GetInsidePrivs(p����ҽ��վ), "��Լ���˵Ǽ�") = 0
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
        If i = cmd��λ���� Then
            cmdEdit(i).Enabled = Not txtEdit(txt��λ����).Locked
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
        blnDoc = mlng����ID <> 0 And mlng����ID = mPatiInfo.����ID And _
                 (mPatiInfo.����id = 0 And mPatiInfo.�����ļ�id <> 0 Or mPatiInfo.����id <> 0 And mPatiInfo.�Ƿ�ǩ�� = False) And (mintActive = pt���� Or mintActive = pt����)
        If blnDoc And mPatiInfo.����id <> 0 And lblҽ��(1).Tag = "0" Then   'û���޸����˲�����Ȩ��
            blnDoc = mPatiInfo.������ = UserInfo.����
        End If
       
        k = 0
        For i = 0 To rtfEdit.Count - 1
            rtfEdit(i).Locked = Not blnDoc Or InStr(rtfEdit(i).Tag, ",") > 0   '���ڶ�������ʱ(������ȫ�ı༭����)�����������޸�
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

        If mPatiInfo.����id = 0 Or mPatiInfo.����id <> 0 And mPatiInfo.�Ƿ�ǩ�� = False Then
            cmdSign.Caption = "ǩ��(&S)"
        Else
            cmdSign.Caption = "ȡ��ǩ��(&S)"
        End If
        cmdSign.Enabled = mlng����ID <> 0 And mlng����ID = mPatiInfo.����ID And (mPatiInfo.����id = 0 And mPatiInfo.�����ļ�id <> 0 Or mPatiInfo.����id <> 0) And (mintActive = pt���� Or mintActive = pt����)
        
        If cmdSign.Enabled And mPatiInfo.����id <> 0 And lblҽ��(1).Tag = "0" Then   'û���޸����˲�����Ȩ��
            cmdSign.Enabled = mPatiInfo.������ = UserInfo.����
        End If
        cmdUpdate.Enabled = cmdSign.Enabled
    End If
                
    mblnPatiEditable = blnDo Or blnDoc
    
    '���˻�����Ϣ���������Ա����䣬�������ڲ������޸�
    txtEdit(txt����).BackColor = &H8000000F: txtEdit(txt����).Locked = True: txtEdit(txt����).TabStop = False
    cboEdit(cbo�Ա�).BackColor = &H8000000F: cboEdit(cbo�Ա�).Locked = True: cboEdit(cbo�Ա�).TabStop = False
    txt��������.BackColor = &H8000000F: txt��������.Enabled = False
    txt����ʱ��.BackColor = &H8000000F: txt����ʱ��.Enabled = False
    txtEdit(txt����).BackColor = &H8000000F: txtEdit(txt����).Locked = True: txtEdit(txt����).TabStop = False
    cboEdit(cbo����).BackColor = &H8000000F: cboEdit(cbo����).Locked = True: cboEdit(cbo����).TabStop = False

End Sub

Private Sub mclsEPRs_RequestRefresh()
    If mblnDocInput Then
        Call LoadDocData
        Call SetPermitEdit
        Call PicBasis_Resize
    End If
End Sub


Private Function CheckIsAskNextQueue(Optional strҵ��ID As String = "") As Boolean
   '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ҽ���Ƿ����������һ������
    '���ƣ����˺�
    '����:����,����true,���򷵻�False
    '���ڣ�2010-06-09 16:48:30
    '˵��������׼:��ʵ���Ѻ���Ϊ׼(ֻ����ɺ󣬲����ٽ�)(����:37442)
    '   ȡ��:��������(�������������)+�ѽ����+ת��<��������
    '------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, lngCount As Long, rsTemp As ADODB.Recordset
    Dim strSQL As String, strLimit As String, strResult As String, arrCheck As Variant
    
    If Val(strҵ��ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strҵ��ID))
           strResult = Nvl(rsTemp!chk) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "�Ƿ����?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Function
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Function
              End If
              
           End If
    End If
    
    
    If mty_Queue.blnҽ���������� = False Or mty_Queue.int�������� <= 0 Then
        CheckIsAskNextQueue = True: Exit Function
    End If
    '0:�Ŷ��У�1:�����У�2�������ţ�3����ͣ��4����ɾ��6�����7���Ѻ���
    'mty_Queue.bln���к�����
    
    '����:44250
    strLimit = ",0,4," & IIf(mty_Queue.bln���к�����, "", ",6,")
    strSQL = "" & _
    "   Select Count(distinct B.ID) as Count From ���˹Һż�¼ B ,�ŶӽкŶ��� A" & _
    "   Where A.ҵ��ID=B.ID And A.ҵ������=0  " & _
    "               And instr([4],','||A.�Ŷ�״̬||',')=0   And B.��¼����=1 And B.��¼״̬=1" & _
    "               And A.ҽ������||''=[1]   " & IIf(mty_Queue.bln���к�����, " And nvl(A.�������,0) = 0", "") & _
    "               And (  (nvl(B.����,0)=1  and B.����ʱ��>=Sysdate-[3] ) or   (nvl(B.����,0)<>1  and B.����ʱ��>=Sysdate-[2] )) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), IIf(gint����Һ����� = 0, 1, gint����Һ�����), strLimit)
    lngCount = Val(Nvl(rsTemp!Count))

    If lngCount >= mty_Queue.int�������� Then
            MsgBox "���ֻ����" & mty_Queue.int�������� & "�����ﲡ��,�����ٽ��к��У�", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    CheckIsAskNextQueue = True
End Function
 
Private Sub mclsInOutMedRec_Closed(ByVal blnEditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String)
'���ܣ������¼�
' strTag=������Ϣ�����ڴ洢���ﲡ����Ƭ�ļ���·�����Ժ���չʱ����|�ָ�
    Dim strPictureFile As String, blnNo As Boolean
    If Not blnEditCancel Then
        If InStr(";" & GetPrivFunc(glngSys, p���ﲡ������) & ";", ";������д;") > 0 Then
            If mobjEPRDoc Is Nothing Then
                Set mobjEPRDoc = New zlRichEPR.cEPRDocument
            End If
            Call OpenEPRDoc(mobjEPRDoc, Me, mPatiInfo.����ID, mPatiInfo.�Һ�ID, mPatiInfo.����ID, str����ID, str���ID, 1, , False, , blnNo)
            If blnNo Then
                Call mclsDisease.EditNotFillReason(Me, mPatiInfo.����ID, mPatiInfo.�Һ�ID, 1)
            End If
        End If
        strPictureFile = Trim(Split(strTag & "|", "|")(0))
         '���������޸�Ϊ��ģ��ʾ����˲���ͨ����������ֵ��ˢ�½��棬ͨ��Closedʱ��ˢ��
        If strPictureFile <> "" And strPictureFile <> "0" Then
            Call ReadPatPricture(mPatiInfo.����ID, imgPatient, strPictureFile)
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
        Call mclsEPRs.zlRefresh(.����ID, .�Һ�ID, mlng����ID, mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
        Call SetPermitEdit
    End With
End Sub

Private Sub mobjQueue_OnQueueExecuteAfter(ByVal strҵ��ID As String, ByVal byt�������� As Byte)
    '------------------------------------------------------------------------------------------------------------------------
    '��Σ�byt��������-0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    '------------------------------------------------------------------------------------------------------------------------
    If mty_Queue.blnҽ���������� = False Then Exit Sub
    If byt�������� <> 1 Then Exit Sub

    
    '����ˢ�²�����Ϣ
    Call LoadPatients("1000")
End Sub

Private Sub mobjQueue_OnQueueExecuteBefore(ByVal strҵ��ID As String, ByVal byt�������� As Byte, blnCancel As Boolean, strNewQueueName As String)
    Dim strSQL As String, rsTemp As ADODB.Recordset
   ' byt�������� -0 - ����, 1 - ֱ��, 2 - ����, 3 - ��ͣ, 4 - ��ɾ���, 5 - �㲥
   
    If InStr(1, "15", byt��������) = 0 Then Exit Sub
    If CheckIsAskNextQueue(strҵ��ID) = False Then blnCancel = True: Exit Sub
    
    strSQL = "SELECT a.ID,a.No,a.����ID,a.ִ�в���ID,A.ִ��״̬ From ���˹Һż�¼ A,�ŶӽкŶ��� B  " & _
        "  where  a.ID=b.ҵ��id and b.ҵ������=0 and a.ID=[1] and nvl(b.�Ŷ�״̬,0)=0 And a.��¼���� in(1,2) And a.��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strҵ��ID))
    If rsTemp.EOF Then Exit Sub
    
    '68736:������,2014-02-18,ת�ﲡ��û��������Ϣ
    If byt�������� = 1 Then
        If Isת�ﲡ��(strҵ��ID) Then
            If CheckTransferDetail(strҵ��ID) = False Then
                strSQL = "ZL_���˹Һż�¼_�������� ('" & Nvl(rsTemp!NO) & "'," & Val(Nvl(rsTemp!����ID)) & ",'" & mstr�������� & "','" & UserInfo.���� & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            Exit Sub
        End If
    End If
    
    If InStr(1, "12", Val(Nvl(rsTemp!ִ��״̬))) > 0 Then
        '1-��ɾ���,2-���ھ���:��Ҫ�ǵڶ��κ���
        'Ӧ����:����Ѿ������,ҽ�������,�в���ȥ����,�ٸ���������
        Exit Sub
    End If
    
    '��������_In Integer := 1
    strSQL = "ZL_���˹Һż�¼_�������� ('" & Nvl(rsTemp!NO) & "'," & Val(Nvl(rsTemp!����ID)) & ",'" & mstr�������� & "','" & UserInfo.���� & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),0)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
End Sub

Private Function CheckTransferDetail(strID As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'����:����ת�ﲡ���Ƿ���������Ϣ
'���:strID-strҵ��ID
'����:True ����ת�ﲡ����������Ϣ False ����ת�ﲡ����������Ϣ
'����:������
'����:2014-02-18
'��ע:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    
    strSQL = "Select ���� From �ŶӽкŶ��� Where ҵ��Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
    '�ŶӽкŶ���û�м�¼,������
    If rsTemp.EOF Then CheckTransferDetail = True: Exit Function
    If Nvl(rsTemp!����) = "" Then CheckTransferDetail = False: Exit Function
    CheckTransferDetail = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Isת�ﲡ��(strҵ��ID As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '����:���ò����Ƿ���ת�ﲡ�˲���δ����
    '���:strҵ��ID
    '����:True ����Ϊת�ﲡ�� False ����Ϊ��ͨ����
    '����:����
    '��������:2012-9-14
    '�����:51514
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    strSQL = _
    "   Select Count(ID) as �Ƿ�Ϊת�ﲡ�� From ���˹Һż�¼ Where ID=[1] And Nvl(ת�����ID,0) <> 0 And Nvl(ת��״̬,0)=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҵ��ID)
    If rsTemp.EOF Then Isת�ﲡ�� = False
    Isת�ﲡ�� = rsTemp!�Ƿ�Ϊת�ﲡ�� > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mobjQueue_OnRecevieDiagnose(ByVal strҵ��ID As String, ByVal lngҵ������ As Long)
    '����:
    Dim objControl As CommandBarControl
    Dim strNO As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim bln���� As Boolean, arrCheck As Variant, strResult As String
    Dim blnת�ﲡ�� As Boolean '�����:51514
    Dim datCurr As Date
    
    If lngҵ������ <> 0 Then Exit Sub
    On Error GoTo errH
     If Val(strҵ��ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strҵ��ID))
           strResult = Nvl(rsTmp!chk) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "�Ƿ����?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Sub
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Sub
              End If
              
           End If
    End If
    strSQL = "Select ����ID,ִ����,NO,��¼��־,ִ��״̬,��¼����,����,�����,id as �Һ�id,����,���� From ���˹Һż�¼ Where  ID=[1]  "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҵ��ID)
    If rsTmp.EOF Then
        MsgBox "�ò���û�йҺż�¼���ܽ��", vbInformation, gstrSysName
        Call LoadPatients("100001"): Exit Sub
    End If
    
    '�����:57566
    If Check�������("����", rsTmp!NO) = False Then Exit Sub
    
    '0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
    If Val(rsTmp!ִ��״̬) = 1 Then
        MsgBox "�ò����Ѿ���ɾ���,�����ٽ��о��������", vbInformation, gstrSysName
        Call LoadPatients("100001"): Exit Sub
    ElseIf Val(rsTmp!ִ��״̬) = -1 Then
        MsgBox "�ò����Ѿ����Ϊ������,�����ٽ��о��������", vbInformation, gstrSysName
        Call LoadPatients("100001"): Exit Sub
    End If
    strNO = Nvl(rsTmp!NO)
    
    'ת����� �����:51514
    blnת�ﲡ�� = Isת�ﲡ��(strҵ��ID)
    If blnת�ﲡ�� Then
        strSQL = "Zl_���˹Һż�¼_ת��('" & strNO & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        'ˢ�²���λ����
        If lvwPati(pt����).Visible Then
            Call LoadPatients("11011", pt����, strNO)
        Else
            Call LoadPatients("11011")
        End If
    End If
    
    '����ԤԼ�Һŵ�
    datCurr = zlDatabase.Currentdate
    If Val("" & rsTmp!��¼����) = 2 Then
        If Val(zlDatabase.GetPara("����ҺŻ��۵�", glngSys, p����ҽ��վ, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, strNO, mstr��������, 0, "") Then Exit Sub
        Else
            strSQL = "Zl_����ԤԼ�Һ�_����('" & strNO & "','" & mstr�������� & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Else
        If Val(Nvl(rsTmp!ִ��״̬)) = 0 Then
            '�����ҺŽ���
            strSQL = "zl_���˽���(" & Val(Nvl(rsTmp!����ID)) & ",'" & strNO & "',Null,'" & UserInfo.���� & "','" & mstr�������� & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Else
            'Zl_���˽���
            strSQL = "Zl_���˽���("
            '  ����id_In     ������Ϣ.����id%Type,
            strSQL = strSQL & "" & Val(Nvl(rsTmp!����ID)) & ","
            '  No_In         ���˹Һż�¼.NO%Type,
            strSQL = strSQL & "'" & strNO & "',"
            '  ִ�в���id_In ���˹Һż�¼.ִ�в���id%Type,
            strSQL = strSQL & "" & IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID) & ","
            '  ִ����_In     ���˹Һż�¼.ִ����%Type,
            strSQL = strSQL & "'" & IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��) & "',"
            '  ����_In       ���˹Һż�¼.����%Type := Null,
            strSQL = strSQL & "'" & mstr�������� & "',"
            '  ��Ǽ���_In   ���˹Һż�¼.����%Type := 0,
            strSQL = strSQL & "0,"
            '  ����_In Integer:=0
            strSQL = strSQL & "1,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            bln���� = True
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
        
    mstr�Һŵ� = strNO
    mlng����ID = Val(Nvl(rsTmp!����ID))
    
    '���ﻼ�߽�����Ϣ����
    Call ZLHIS_CIS_009(mclsMipModule, mlng����ID, Nvl(rsTmp!����), Nvl(rsTmp!�����), 0, 0, Nvl(rsTmp!�Һ�ID), Nvl(rsTmp!����, 0), Nvl(rsTmp!����, 0), datCurr, mlng�������ID, , mstr��������, UserInfo.����)
    
    'ˢ�²���λ����
    On Error GoTo 0
    If lvwPati(pt����).Visible Then
        Call LoadPatients("110001", pt����, strNO)
        lvwPati(pt����).SetFocus
    Else
        Call LoadPatients("110001")
    End If
    '���������Զ����ù���
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng����ID <> 0 And mPatiInfo.���� <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    
    Call CreatePlugInOK(p����ҽ��վ)
    '����������ҽӿ�
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicReceive(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID)
        Call zlPlugInErrH(err, "ClinicReceive")
        err.Clear: On Error GoTo errH
    End If
    
    '����֮���Զ�����ҽ���´�״̬
    If mlng�Զ����� = 1 And bln���� = False Then
        If tbcSub.Selected.Tag <> "ҽ��" Then tbcSub.Item(0).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    ElseIf mlng�Զ����� = 2 And bln���� = False Then
        If tbcSub.Selected.Tag <> "����" Then tbcSub.Item(1).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr�Һŵ�)
    End If
    '�����ŶӽкŶ���(����ˢ��)
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjQueue_OnSelectionChanged(ByVal blnIsCallingList As Boolean, objReportRow As Object, cbrMain As Object)
    If mty_Queue.blnҽ���������� Then
        mobjQueue.zlCommandBarSet 7, blnIsCallingList Or Not mbln���к����
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
        lngPatiID = objHisPati.����ID
    End If
    
    Call ExecuteFindPati(False, , blnCard, lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index: mstrFindType = objCard.����
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
    If picExpand.Picture Is ilexpand.ListImages("չ��").Picture Then
        Set picExpand.Picture = ilexpand.ListImages("�۵�").Picture
        mblnPatiDetail = True
    Else
        Set picExpand.Picture = ilexpand.ListImages("չ��").Picture
        mblnPatiDetail = False
    End If
    Call zlDatabase.SetPara("��ʾ������ϸ��Ϣ", IIf(mblnPatiDetail, 1, 0), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
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
        lbl����ʱ��.Left = 100
    cboSelectTime.Left = lbl����ʱ��.Left + lbl����ʱ��.Width + 15
    cmdOtherFilter.Left = cboSelectTime.Left + cboSelectTime.Width + 50
    lvwPati(pt����).Top = cboSelectTime.Top + cboSelectTime.Height + 30
    lvwPati(pt����).Width = picYZ.Width
    lvwPati(pt����).Height = picYZ.Height - lvwPati(pt����).Top
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
            '�������λس������ת
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
            If Mid(.Text, .SelStart, 1) = "`" Or Mid(.Text, .SelStart, 1) = "��" Then
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
        Case txt����
            strType = "��������"
        Case txt����ʷ
            strType = "����ʷ"
        Case txt�ֲ�ʷ
            strType = "�ֲ�ʷ"
        Case txt����
            strType = "���һ����"
        Case txt��ȥʷ
            strType = "����ʷ"
        End Select
                
        strSentence = frmSentenceSel.ShowMe(Me, mPatiInfo.�����ļ�id, mPatiInfo.�Ա�, mPatiInfo.����״��, strType, txtSentence.Text, picSentence.hwnd, blnCancel)
        If strSentence <> "" Then
            rtfEdit(Val(picSentence.Tag)).SelText = strSentence
            Call HideWordInput
        Else
            If Not blnCancel Then
                MsgBox "û���ҵ�ƥ��Ĵʾ䡣", vbInformation, gstrSysName
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
    Case txt����
        strType = "��������"
    Case txt����ʷ
        strType = "����ʷ"
    Case txt�ֲ�ʷ
        strType = "�ֲ�ʷ"
    Case txt����
        strType = "���һ����"
    Case txt��ȥʷ
        strType = "����ʷ"
    End Select
    
    strSentence = frmSentenceSel.ShowMe(Me, mPatiInfo.�����ļ�id, mPatiInfo.�Ա�, mPatiInfo.����״��, strType)
    If strSentence <> "" Then
        rtfEdit(Val(picSentence.Tag)).SelText = strSentence
        Call HideWordInput
    End If
End Sub

Private Sub txtSentence_LostFocus()
    If Not frmSentenceSel.mblnShow Then
        Call HideWordInput   '���شʾ�����
    End If
End Sub


Private Sub ShowWordInput(ByRef txtThis As RichTextBox)
'���ܣ���ʾ�ʾ�����
    Dim vPos As POINTAPI
    
    If txtThis.Visible And txtThis.Enabled And Not txtThis.Locked Then
        picSentence.Tag = txtThis.Index '�����Ա����ط��غ�λ
        
        If txtThis.Text = "" Then picPatiInput.Tag = "UnChange": txtThis.Text = " " '����Ҫ��һ�����ַ����ܷ���������
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
'���ܣ����شʾ�����
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
    If Index = txt���֤�� Then
        If txtEdit(txt���֤��).Tag <> "������Change�¼�" Then
            lngPos = InStr(txtEdit(txt���֤��).Text, "*")
            lngLen = Len(Mid(txtEdit(txt���֤��).Text, 13, 2))
            Select Case lngPos
                Case 0
                    txtEdit(txt���֤��).Tag = txtEdit(txt���֤��).Text
                Case Else
                    txtEdit(txt���֤��).Tag = Mid(txtEdit(txt���֤��).Text, 1, lngPos - 1)
                    txtEdit(txt���֤��).Text = txtEdit(txt���֤��).Tag
                    txtEdit(txt���֤��).SelStart = Len(txtEdit(txt���֤��).Text)
            End Select
        End If
    End If
    Call SetPermitEscape(False)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtEdit(Index))
    
    Select Case Index
        Case txt��λ����, txt��ͥ��ַ, txt�໤��, txt����ժҪ
            Call zlCommFun.OpenIme(True)
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt��ͥ��ַ) And txtEdit(Index).Text <> "" Then
            '�����������
            strSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!����
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt��λ���� And txtEdit(Index).Text <> "" Then
            '���빤����λ
            strSQL = "Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������λ", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If InStr(GetInsidePrivs(p����ҽ��վ), "��Լ���˵Ǽ�") > 0 Then txtEdit(Index).Tag = Val(rsTmp!ID)
                If txtEdit(txt��λ�绰).Text = "" Then
                    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!�绰)
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
        '�ǿ��ư���
        
        'ѡ���ݼ�
        If KeyAscii = Asc("*") Then
            'ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
            If Index = txt��ͥ��ַ Then
                KeyAscii = 0
                Call cmdEdit_Click(cmd��ͥ��ַ)
                Exit Sub
            ElseIf Index = txt��λ���� Then
                KeyAscii = 0
                Call cmdEdit_Click(cmd��λ����)
                Exit Sub
            End If
        End If
        
        '�������볤��
        If txtEdit(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(Index).Text) > txtEdit(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '������������
        Select Case Index
'            Case txt���� '��������¼����
'                strMask = "1234567890"
            'Case txt�������� 'MaskEdit������
                'strMask = "1234567890-"
            Case txt���֤��
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Case txt��ͥ�绰, txt��λ�绰
                strMask = "1234567890-()"
            Case txt�ֻ���
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
    
        Case txt����
            If Trim(txtEdit(txt����).Text) <> "" Then
                If IsNumeric(txtEdit(txt����).Text) Then
                    If Val(txtEdit(txt����).Text) <= 0 Then
                        MsgBox "����ʱ������ֵ����Ϊ������", vbInformation, gstrSysName
                        txtEdit(txt����).SetFocus: Exit Sub
                    End If
                Else
                    MsgBox "����ʱ������ֵ����Ϊ���֡�", vbInformation, gstrSysName
                    txtEdit(txt����).Text = "": txtEdit(txt����).SetFocus: Exit Sub
                End If
            Else
                 Exit Sub
            End If
            If cboEdit(cbo����ʱ��).ListIndex <= 0 Then Exit Sub
            datCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            Select Case cboEdit(cbo����ʱ��).ListIndex
                Case 1 'Сʱ
                    datRes = DateAdd("n", -1 * Val(txtEdit(txt����).Text) * 60, CDate(datCur))
                Case 2 '��
                    datRes = DateAdd("h", -1 * Val(txtEdit(txt����).Text) * 24, CDate(datCur))
                Case 3 '��
                    datRes = DateAdd("d", -1 * 7 * Val(txtEdit(txt����).Text), CDate(datCur))
                Case 4 '��
                    datRes = DateAdd("M", -1 * Int(Val(txtEdit(txt����).Text)), CDate(datCur))
                    datRes = DateAdd("d", -1 * (Val(txtEdit(txt����).Text) - Int(Val(txtEdit(txt����).Text))) * 30, datRes)
                Case 5 '��
                    If Val(txtEdit(txt����).Text) < 100 Then
                        datRes = DateAdd("yyyy", -1 * Int(Val(txtEdit(txt����).Text)), CDate(datCur))
                        datRes = DateAdd("d", -1 * (Val(txtEdit(txt����).Text) - Int(Val(txtEdit(txt����).Text))) * 365, datRes)
                    Else
                        MsgBox "����ʱ�����㲻�ܳ���100�ꡣ", vbInformation, gstrSysName
                        txtEdit(txt����).SetFocus: Exit Sub
                    End If
            End Select
            txt��������.Text = Format(CDate(datRes), "YYYY-MM-DD")
            If cboEdit(cbo����ʱ��).ListIndex < 3 Then
                txt����ʱ��.Text = Format(CDate(datRes), "HH:mm")
            End If
     Case txt�ֻ���
        If Not IsNumeric(Trim(txtEdit(txt�ֻ���).Text)) And txtEdit(txt�ֻ���).Text <> "" Then
            MsgBox "��ǰ¼����ֻ��Ÿ�ʽ����ȷ��������¼��!", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End Select
End Sub

Private Sub txt��������_Change()
    Call SetPermitEscape(False)
    
    If IsDate(txt��������.Text) Then
        txt����ʱ��.Enabled = True
    Else
        txt����ʱ��.Enabled = False
        txt����ʱ��.Text = "__:__"
    End If
End Sub

Private Sub txt��������_GotFocus()
    Call zlControl.TxtSelAll(txt��������)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
        txt��������.Text = "____-__-__": Cancel = True
    ElseIf txt��������.Text = "____-__-__" Then
        If txt����ʱ��.Text <> "__:__" Then
            txt����ʱ��.Text = "__:__"
        End If
    End If
    Call PicBasis_Resize
End Sub

Private Sub txt����ʱ��_Change()
    Call SetPermitEscape(False)
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.Text = "__:__": Cancel = True
    End If
    Call PicBasis_Resize
End Sub

Private Sub cmdEdit_Click(Index As Integer)
'˵����ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
        
    Select Case Index
        Case cmd��ͥ��ַ
            'ѡ���������
            strSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            vPoint = GetCoordPos(txtEdit(txt��ͥ��ַ).Container.hwnd, txtEdit(txt��ͥ��ַ).Left, txtEdit(txt��ͥ��ַ).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.x, vPoint.y, txtEdit(txt��ͥ��ַ).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtEdit(txt��ͥ��ַ).SetFocus
            Else
                txtEdit(txt��ͥ��ַ).Text = rsTmp!����
                txtEdit(txt��ͥ��ַ).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case cmd��λ����
            'ѡ��λ��Ϣ
            strSQL = "Select ID,�ϼ�ID,ĩ��,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��" & _
                " From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            vPoint = GetCoordPos(txtEdit(txt��λ����).Container.hwnd, txtEdit(txt��λ����).Left, txtEdit(txt��λ����).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "��Լ��λ", , , , , True, True, vPoint.x, vPoint.y, txtEdit(txt��λ����).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""��Լ��λ""���ݣ����ȵ���Լ��λ���������á�", vbInformation, gstrSysName
                End If
                txtEdit(txt��λ����).Tag = ""
                txtEdit(txt��λ����).SetFocus
            Else
                txtEdit(txt��λ����).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If InStr(GetInsidePrivs(p����ҽ��վ), "��Լ���˵Ǽ�") > 0 Then txtEdit(txt��λ����).Tag = Val(rsTmp!ID)
                If txtEdit(txt��λ�绰).Text = "" Then
                    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!�绰)
                End If
                txtEdit(txt��λ����).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
    End Select
End Sub

Private Sub cboRegist_Click()
'���ܣ�ѡ��ĳ����ʷ�����¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    If cboRegist.ListIndex = -1 Then
        '����ǰ�б�������ˢ���Ӵ���
        Call ClearPatiInfo
        
        cboRegist.Tag = ""
        Call SetPermitEdit
        Call PicBasis_Resize
        'ˢ���Ӵ�������
        Call SubWinRefreshData(tbcSub.Selected)
        
        '��ȡ�򵥲�������
        If mblnDocInput Then Call LoadDocData
        Exit Sub
    End If
    If cboRegist.ListIndex = mintPreTime Then Exit Sub
    mintPreTime = cboRegist.ListIndex
       
    cboRegist.Tag = "Loading"   '���ڶԱ༭�ؼ������Ƿ�ı���ж�ʱ�ſ�����ʱ�ĳ��θı�
    mblnPatiChange = False              '���ڼ�¼�Ƿ��޸Ĺ����ݣ��ж��Ƿ���Ҫ����
    
    On Error GoTo errH
    strSQL = "Select E.����,B.Id,B.NO,B.�����,B.����,B.�Ա�,B.����,A.��������,B.ҽ�Ƹ��ʽ,A.ְҵ," & _
        "   A.�ѱ�,A.����,A.ҽ����,B.����,A.����ģʽ,B.����ʱ��,B.ִ����,B.ִ��״̬,B.ִ��ʱ��," & _
        "   B.ִ�в���ID as ����ID,B.����,B.����,D.������,C.���� as ����,B.����,B.ժҪ," & _
        "   A.���֤��,A.�໤��,A.��ͥ��ַ,A.��ͥ�绰,A.������λ,A.��ͬ��λid,A.��λ�绰,B.����ʱ��,B.������ַ," & _
        "   A.����,A.����,A.����,A.����״��,A.��ͥ��ַ�ʱ�,A.��λ�ʱ�,A.�����ص�,B.��Ⱦ���ϴ�,A.����֤��,a.���ڵ�ַ,a.���ڵ�ַ�ʱ�,a.����,a.email,a.qq,A.��������,A.����ID,A.�ֻ���" & _
        " From ������Ϣ A,���˹Һż�¼ B,���ű� C,����������Ϣ D,�ҺŰ��� E" & _
        " Where A.����ID=B.����ID And B.ID=[1] And B.ִ�в���ID=C.ID" & _
        " And B.����ID=D.����ID(+) And B.����=D.����(+) And B.�ű�=E.����(+)"
        '��ID��ȡ�Һż�¼�����üӼ�¼���ʡ�״̬������
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboRegist.ItemData(cboRegist.ListIndex))
    With rsTmp
        txtEdit(txt����).Text = "" & !����
        txtEdit(txt����).Tag = "" & !����
        '��ʾ������ɫ
        If Not IsNull(!����) And Nvl(rsTmp!��������) = "" Then
            txtEdit(txt����).ForeColor = &HC0&
        Else
            txtEdit(txt����).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
        End If
        
        Call zlControl.CboLocate(cboEdit(cbo�Ա�), "" & !�Ա�)
                
        If Not IsNull(!��������) Then
            txt��������.Text = Format(!��������, "yyyy-MM-dd")
            If Format(!��������, "HH:mm") <> "00:00" Then txt����ʱ��.Text = Format(!��������, "HH:mm")
        End If

        Call LoadOldData("" & !����, txtEdit(txt����), cboEdit(cbo����))
        
        Call zlControl.CboLocate(cboEdit(cboְҵ), "" & !ְҵ)
        If Not IsNull(!����ʱ��) Then
            txt��������.Text = Format(!����ʱ��, "yyyy-MM-dd")
            txt����ʱ��.Text = Format(!����ʱ��, "HH:mm")
            If txt����ʱ��.Text = "00:00" Then txt����ʱ��.Text = "__:__"
        Else
            txt��������.Text = "____-__-__": txt����ʱ��.Text = "__:__"
        End If
        txtEdit(txt������ַ).Text = Nvl(rsTmp!������ַ)
        lbl��.Visible = Nvl(!����, 0) <> 0
        lblRec.Visible = Nvl(!����ģʽ, 0) <> 0
                Call picPatiInput_Resize
        
        txtEdit(txt���֤��).Tag = "������Change�¼�"
        txtEdit(txt���֤��).Text = "" & !���֤��
        If zlCommFun.ActualLen(txtEdit(txt���֤��).Tag) > 12 And mblnMaskID Then   '�������֤������
            txtEdit(txt���֤��).Text = Mid(txtEdit(txt���֤��).Text, 1, 12) & String(Len(Mid(txtEdit(txt���֤��).Text, 13, 2)), "*") & Mid(txtEdit(txt���֤��).Text, 15)
        End If
        txtEdit(txt���֤��).Tag = "" & !���֤��
                lblEdit(20).Tag = "" & !���֤��  '�������ݱ������ж�ʱ�õ�
        
        If Val("" & !����) = 1 Then
            optState(opt����).Value = True
        Else
            optState(opt����).Value = True
        End If
        txtEdit(txt��λ����).Text = "" & !������λ
        txtEdit(txt��λ����).Tag = Val("" & !��ͬ��λid)
                
        txtEdit(txt��λ�绰).Text = "" & !��λ�绰
        txtEdit(txt��ͥ��ַ).Text = "" & !��ͥ��ַ
        txtEdit(txt�໤��).Text = "" & !�໤��
        txtEdit(txt��ͥ�绰).Text = "" & !��ͥ�绰
        txtEdit(txt�ֻ���).Text = "" & !�ֻ���
        txtEdit(txt����ժҪ).Text = "" & !ժҪ
                        
        lblShow(lbl�ѱ�).Caption = "" & !�ѱ�
        lblShow(lbl����).Caption = "" & !ҽ�Ƹ��ʽ
        lblShow(lbl����).Caption = "" & !����
        lblShow(lblҽ����).Caption = "" & !ҽ����
        If IsNull(!������) Then
            lblShow(lbl������).Caption = ""
            lblShow(lbl������).Visible = False
            lblTitle������.Visible = False
        Else
            lblShow(lbl������).Caption = "" & !������
            lblShow(lbl������).Visible = True
            lblTitle������.Visible = True
        End If
                
        '���
        lblDiag(1).Caption = GetPatiDiagnose(Val(rsTmp!����ID & ""), cboRegist.ItemData(cboRegist.ListIndex), 1)
        
        '������Ϣ
        If mintActive = ptת�� Then
            mPatiInfo.���� = ptת��
        Else
            mPatiInfo.���� = Decode(Nvl(!ִ��״̬, 0), 0, 0, 2, 1, 1, 2)
        End If
        mPatiInfo.����� = Nvl(!�����)
        mPatiInfo.�Һ�ID = !ID
        mPatiInfo.����ID = !����ID
        mPatiInfo.�Һŵ� = !NO
        mPatiInfo.����ID = !����ID
        mPatiInfo.���� = Nvl(!����)
        mPatiInfo.���� = Nvl(!����, 0)
        mPatiInfo.������ = Nvl(!������)
        mPatiInfo.�Һ�ʱ�� = !����ʱ��
        mPatiInfo.�Ա� = "" & !�Ա�
        mPatiInfo.����״�� = "" & !����״��
        
        mPatiInfo.���� = "" & !����
        mPatiInfo.���� = "" & !����
        mPatiInfo.���� = "" & !����
        mPatiInfo.�����ص� = "" & !�����ص�
        mPatiInfo.��Ⱦ���ϴ� = Val("" & !��Ⱦ���ϴ�)
        mPatiInfo.��ͥ��ַ�ʱ� = "" & !��ͥ��ַ�ʱ�
        mPatiInfo.��λ�ʱ� = "" & !��λ�ʱ�
        mPatiInfo.����֤�� = "" & !����֤��
        mPatiInfo.���ڵ�ַ = "" & !���ڵ�ַ
        mPatiInfo.���ڵ�ַ�ʱ� = "" & !���ڵ�ַ�ʱ�
        mPatiInfo.���� = "" & !����
        mPatiInfo.Email = "" & !Email
        mPatiInfo.QQ = "" & !QQ
        
        If mPatiInfo.���� = pt���� Then
            mPatiInfo.����ת�� = zlDatabase.NOMoved("���˹Һż�¼", !NO)
        Else
            mPatiInfo.����ת�� = False
        End If
        picPatient.Visible = ReadPatPricture(mPatiInfo.����ID, imgPatient)
    End With
    '������Ϣ
    Call ucPatiVitalSigns.LoadPatiVitalSigns(mPatiInfo.����ID, cboRegist.ItemData(cboRegist.ListIndex))
    strSQL = "Select ��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And (����ID=[2] Or ����ID is Null) Order by Nvl(����ID,999999999)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.����ID, cboRegist.ItemData(cboRegist.ListIndex))
    rsTmp.Filter = "��Ϣ��='ȥ��'"
    'ɾ��������ӵ�
    If cboEdit(cboȥ��).ListCount <> 0 Then
        If cboEdit(cboȥ��).ItemData(cboEdit(cboȥ��).ListCount - 1) = -1 Then
            cboEdit(cboȥ��).RemoveItem (cboEdit(cboȥ��).ListCount - 1)
        End If
    End If
    If Not rsTmp.EOF Then
        If Not zlControl.CboLocate(cboEdit(cboȥ��), Nvl(rsTmp!��Ϣֵ)) Then
            cboEdit(cboȥ��).AddItem Nvl(rsTmp!��Ϣֵ)
            cboEdit(cboȥ��).ItemData(cboEdit(cboȥ��).NewIndex) = -1
        End If
        cboEdit(cboȥ��).Text = Nvl(rsTmp!��Ϣֵ)
    Else
        cboEdit(cboȥ��).ListIndex = 0
    End If
    cboEdit(cbo����ʱ��).ListIndex = 0
    
    For i = 0 To cboEdit.UBound
        If i = 3 Then i = i + 2
        If cboEdit(i).ListIndex <> -1 Then
            cboEdit(i).Tag = cboEdit(i).List(cboEdit(i).ListIndex)
        Else
            cboEdit(i).Tag = ""
        End If
    Next
    txtEdit(txt����).Text = ""
    txtEdit(txt����).Tag = ""
    cboEdit(cbo����ʱ��).ListIndex = 0
    cboEdit(cbo����ʱ��).Tag = ""
    
    Call ShowAller
    
    'ˢ���Ӵ�������
    Call SubWinRefreshData(tbcSub.Selected)
    
    '��ȡ�򵥲�������
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
'���ܣ����ÿ����������
'������intType=0������ȡ��intType=1���ĵ��룬����ղ���ID
    Dim i As Long, j As Long, arrTmp As Variant
    Dim strContent As String
    
    With rsTmp
        If .RecordCount > 0 Then
            arrTmp = Split("-10,2,3,5,6", ",") '��������,�ֲ�ʷ,����ʷ,����ʷ,�����
            For i = 0 To UBound(arrTmp)
                .Filter = "Ԥ�����id=" & arrTmp(i)
                rtfEdit(i).Text = ""
                If intType = 1 Then
                    '���뷶�ĺ����ж����á�
                    rtfEdit(i).Locked = False
                    rtfEdit(i).BackColor = HColor
                End If
                For j = 1 To .RecordCount
                    If j = 1 Then
                        strContent = "" & !�����ı�
                        If InStr(strContent, lblDoc(i).Tag) = 1 Then strContent = Mid(strContent, Len(lblDoc(i).Tag) + 1)
                        rtfEdit(i).Text = strContent
                        If intType = 0 Then rtfEdit(i).Tag = !ID
                    Else
                        rtfEdit(i).Text = rtfEdit(i).Text & vbCrLf & !�����ı�
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
    If Not blnLoading Then cboRegist.Tag = "Loading"    '����text��ֵʱ����text_Change�¼��иı������ؼ��Ŀ���״̬
    
    For i = 0 To rtfEdit.UBound
        rtfEdit(i).Text = ""
        rtfEdit(i).Tag = ""
    Next
    lbl��������.Caption = ""    '�����ڼ�����Ա��ԭ�����磺ѡ����ٴʾ�ʱ��û���г�Ԥ�ƵĴʾ䣬�ɸ��ݲ����ļ����Ʋ��Ƿ���������ٴʾ��Ӧ
    
    'ֻ��ʾ�򵥲���ģʽ�²������ļ�
    strSQL = "Select id,�ļ�id,ǩ������,��������,������ From ���Ӳ�����¼ A Where ����id = [1] And ��ҳid = [2] And �������� = 1" & vbNewLine & _
            " And Exists(Select 1 From �����ļ��б� B Where A.�ļ�ID = B.ID And B.���� = '3')"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.����ID, mPatiInfo.�Һ�ID)
    If rsTmp.RecordCount > 0 Then
        If SetCompendsTag(Val("" & rsTmp!�ļ�ID)) Then
            mPatiInfo.�����ļ�id = Val("" & rsTmp!�ļ�ID)
            mPatiInfo.�Ƿ�ǩ�� = IIf(Val("" & rsTmp!ǩ������) > 0, True, False)
            mPatiInfo.����id = rsTmp!ID
            mPatiInfo.������ = "" & rsTmp!������
            lbl��������.Caption = "" & rsTmp!��������
                                            
            '��ȡ����µĶ����ı�,��������Ϊ-1��ʾ��ٱ����ı�
            strSQL = "Select A.Ԥ�����id, B.�����ı�, B.ID" & vbNewLine & _
                    "From ���Ӳ������� A, ���Ӳ������� B" & vbNewLine & _
                    "Where A.�ļ�id = [1] And A.�������� = 1 And A.Ԥ�����id+0 In(-10,5,2,6,3)" & vbNewLine & _
                    "      And B.��id = A.ID And B.�������� = 2 Order By A.Ԥ�����id, B.�������"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.����id)
            Call SetDocData(rsTmp, 0)
        End If
    Else
        If lbl��.Visible Then
            strSQL = " And (R.�¼� = '����'  OR R.�¼� IS NUll)"
        Else
            If optState(opt����).Value Then
                strSQL = " And (R.�¼� = '����' Or R.�¼� = '����'  OR R.�¼� IS NUll)"
            Else
                strSQL = " And (R.�¼� = '����' Or R.�¼� = '����'  OR R.�¼� IS NUll )"
            End If
        End If
        'ϵͳ��������(��)�ﲡ���ҶԵ�ǰ�������ã�����5���̶�Ԥ�����,����ʾ����¼�����.
        strSQL = "Select F.ID, F.���� as ��������" & vbNewLine & _
                "From (Select F.ID, F.ͨ��, A.����id, F.����,Decode(R.�¼�,Null,2,1) �¼�" & vbNewLine & _
                "       From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� R" & vbNewLine & _
                "       Where F.ID = A.�ļ�id(+) And F.ID = R.�ļ�id(+) And F.���� = 1 And F.����= '3'" & strSQL & ") F" & vbNewLine & _
                "Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [2]" & vbNewLine & _
                "Order By F.�¼�,F.ͨ�� Desc,F.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.�Һ�ID, mPatiInfo.����ID)
        If rsTmp.RecordCount > 0 Then
            mPatiInfo.�����ļ�id = rsTmp!ID
            lbl��������.Caption = "" & rsTmp!��������
            If SetCompendsTag(mPatiInfo.�����ļ�id) = False Then
                mPatiInfo.�����ļ�id = 0: lbl��������.Caption = ""
            End If
        Else
            mPatiInfo.�����ļ�id = 0: lbl��������.Caption = ""
        End If
        
        mPatiInfo.����id = 0
        mPatiInfo.�Ƿ�ǩ�� = False
    End If
     '��������
     Call SetRTFEditFontSize
     
    PicOutDoc.Tag = ""
    If Not blnLoading Then cboRegist.Tag = ""
    cmdImportEPRDemo.Visible = mPatiInfo.�����ļ�id <> 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function SetCompendsTag(ByVal lng�����ļ�id As Long) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    
    strSQL = "Select Decode(A.Ԥ�����id, -10, 0, 5, 3, 2, 1, 6, 4, 3, 2) As ���, B.�����ı�" & vbNewLine & _
            "From �����ļ��ṹ A, �����ļ��ṹ B" & vbNewLine & _
            "Where A.�ļ�id = [1] And A.Ԥ�����id+0 In (-10,5,2,6,3) And A.Id = B.��id And B.�������� = 2" & vbNewLine & _
            "Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�����ļ�id)
    If rsTmp.RecordCount > 0 And rsTmp.RecordCount <= 5 Then
        If rsTmp!��� & "" = "0" Then  '�����������
            For i = 0 To rsTmp.RecordCount - 1
                lblDoc(Val(rsTmp!��� & "")).Tag = rsTmp!�����ı�       '���ڱ���Rtf�ļ��滻����ʱ��λ
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
'���ܣ����֤ʶ��ɹ��󼤻�
    mstrIDCard = strID
    If mstrFindType = "�������֤" Then
        PatiIdentify.Text = mstrIDCard
    Else
        PatiIdentify.Text = "" '�������(Ŀǰ�������������²��ܼ���)��
    End If
    Call ExecuteFindPati(False, mstrIDCard)
End Sub

Private Function CheckHaveAdvice(ByVal lng����ID As Long, ByVal str�Һŵ� As String) As Boolean
'���ܣ��жϲ����Ƿ���ҽ��
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "select 1 from ����ҽ����¼ where ����ID=[1] and �Һŵ�=[2] and rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str�Һŵ�)
    CheckHaveAdvice = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ˢ���Ӵ�����漰����
'˵����������Ϊ�л����濨Ƭ����
    Dim objControl As CommandBarControl
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    
    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ
     
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "�²���"
                Set objItem = tbcSub.InsertItem(Index, "���Ӳ���", mcolSubForm("_�²���").hwnd, 0)
                objItem.Tag = "�²���"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
     
    'ˢ���Ӵ����Ӧ��CommandBar
    Call SubWinDefCommandBar(Item)
    
    'ˢ���Ӵ�������
    Call SubWinRefreshData(Item)
    
    If Visible Then mfrmActive.SetFocus
    
    '�Զ�����һ������/����/���ﲡ��/�����ҽ����������ҽ�������ж�û��ҽ��������
    If Item.Tag = "����" And mlng�Զ����� = 1 Then
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr�Һŵ�)
        '��Ϊִ��������Ƿ�ģ̬���壬������mclsAdvices��mclsEPRs��active������ mblnUnRefresh = False
    ElseIf Item.Tag = "ҽ��" And mlng�Զ����� = 2 Then
        If CheckHaveAdvice(mlng����ID, mstr�Һŵ�) = False Then
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
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
    ElseIf Item.ID = pt�Ŷӽк� Then
        Item.Handle = mobjQueue.zlGetForm.hwnd
    ElseIf Item.ID = 7 Then  '����
        Item.Handle = lvwPatiHZ.hwnd
    ElseIf Item.ID = 3 Then '����
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
        PicBasis.Height = fraLine(cbo����).Height + fraLine(cbo����).Top
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
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("���˲��ҷ�ʽ", mintFindType, glngSys, p����ҽ��վ, blnSetup)

    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("ҽ������", tbcSub.Selected.Tag, glngSys, p����ҽ��վ, blnSetup)
    End If
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    End If

    '���������̶�����һ���ؼ�����ʽ���棬����վ���������һ���Ǵ�ӡ����̶���ͼ����ʽ,������ָ�Ϊ������ť����ʽ
    If Me.Visible Then  'Form_load���˳�ʱ������
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

    '--�ر������ŶӵĴ���
    If Not mobjQueue Is Nothing Then
        Call mobjQueue.CloseWindows
        Set mobjQueue = Nothing
    End If
    If Not mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = Nothing
    End If
    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
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
    
    '�����:57566
    mlng������� = 0
    mlng��ǰ����ʱ�� = 0
    
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
    'MouseDown����GotFocusִ��
    If Not mblnMouseDown And Not lvwPati(Index).SelectedItem Is Nothing Then
        Call lvwPati_ItemClick(Index, lvwPati(Index).SelectedItem)
    End If
End Sub

Private Sub lvwPati_DblClick(Index As Integer)
'���ܣ�˫���Զ��������ɽ���
    Dim objControl As CommandBarControl
    Dim objItem As ListItem
    Dim vPoint As POINTAPI
    
    Call GetCursorPos(vPoint)
    Call ScreenToClient(lvwPati(Index).hwnd, vPoint)
    Set objItem = lvwPati(Index).HitTest(vPoint.x * Screen.TwipsPerPixelX, vPoint.y * Screen.TwipsPerPixelY)
    If Not objItem Is Nothing And InStr(mstrPrivs, "���˽���") > 0 Then
        If Index = pt���� Then
            Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
        ElseIf Index = pt���� Then
            Set objControl = cbsMain.FindControl(, conMenu_Manage_Finish, True, True)
        End If
        If Not objControl Is Nothing Then
            If objControl.Enabled Then Call cbsMain_Update(objControl) '�״�ִ�У�û����ʾ�˵�ǰ���¼�û��ִ��
            If objControl.Enabled Then objControl.Execute
        End If
    End If
End Sub

Private Sub LvwItemClick(ByVal Index As Integer, ByVal Item As MSComctlLib.ListItem)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ���
    '���:index-5����;0-����;1-����;2-��ɾ���
    '����:���˺�
    '����:2011-01-17 10:59:25
    '��Ҫ�Ǽ����˻����,��Ҫ�ڵ����ص��б�ʱ,������ص�����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, strTmp As String
    Dim intCount As Integer
    Dim objLvw As ListView
    Dim str���֤�� As String
    Dim str����IDs As String
    
    'Index:5-����
    If Index = pt���� Then
        Set objLvw = lvwPatiHZ
    ElseIf Index = ptԤԼ Then
        Set objLvw = lvwReserve
    ElseIf Index = ptת�� Then
        Set objLvw = lvwIncept
    Else
        Set objLvw = lvwPati(Index)
    End If
    
    If objLvw.SelectedItem Is Nothing Then Exit Sub '���������
    With objLvw.SelectedItem
        '��ǰ��б�
        mintActive = Index
        If .Key = mstrPrePati Then Exit Sub
        mstrPrePati = .Key
        '��ǰѡ���˵��б��вſ��Կ���ѡ����,�Ա�����
        lvwPatiHZ.HideSelection = Index <> pt����
        For i = 0 To lvwPati.UBound
            lvwPati(i).HideSelection = i <> Index
        Next
        mstr�Һŵ� = .Text
        mlng����ID = Val("" & .Tag) 'ԤԼ���˿���δ����
        mlng����ID = Val(.ListSubItems(3).Tag)
        str���֤�� = .ListSubItems(6).Tag
        
        LockWindowUpdate Me.hwnd
        
        '��֤���֤��
        If str���֤�� <> "" Then
            If mobjPatient Is Nothing Then
                On Error Resume Next
                Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                err.Clear: On Error GoTo 0
                If mobjPatient Is Nothing Then
                    MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
                Else
                    Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
                End If
            End If
            strTmp = ""
            If Not mobjPatient Is Nothing Then
                If mobjPatient.CheckPatiIdcard(str���֤��) Then
                    strTmp = str���֤��
                End If
            End If
            str���֤�� = strTmp
        End If
        
        On Error GoTo errH
        
        If str���֤�� <> "" Then
            strSQL = "select a.����id from ������Ϣ a where a.����id<>[1] and a.���֤��=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, str���֤��)
            Do While Not rsTmp.EOF
                str����IDs = str����IDs & "," & rsTmp!����ID
                rsTmp.MoveNext
            Loop
            If str����IDs <> "" Then
                str����IDs = mlng����ID & str����IDs
            End If
        End If
        
        
        If str����IDs = "" Then
            '��ȡ"��ʷ��"�����¼
            strSQL = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ���� From ���˹Һż�¼ A,���ű� B" & _
                " Where A.ִ�в���ID=B.ID And A.����ID=[1] And A.����ʱ��<=[2] And A.��¼����=1 And A.��¼״̬=1 Order by A.����ʱ�� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, CDate(Item.ListSubItems(2).Tag))
        Else
            strSQL = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ���� From ���˹Һż�¼ A,���ű� B" & _
                " Where A.ִ�в���ID=B.ID And A.����ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
                " And A.����ʱ��<=[2] And A.��¼����=1 And A.��¼״̬=1 Order by A.����ʱ�� Desc,a.����ʱ�� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs, CDate(Item.ListSubItems(2).Tag))
        End If
        
        cboRegist.Clear
        Do While Not rsTmp.EOF
            cboRegist.AddItem Format(rsTmp!ʱ��, "YYMMdd") & rsTmp!����
            cboRegist.ItemData(cboRegist.NewIndex) = rsTmp!ID
            If rsTmp!NO = mstr�Һŵ� Then
                mlng�Һ�ID = rsTmp!ID
                Call zlControl.CboSetIndex(cboRegist.hwnd, cboRegist.NewIndex)
            End If
            
            '���ն�ƾ���
            If Format(rsTmp!ʱ��, "yyyy-MM-dd") = Format(CDate(Item.ListSubItems(2).Tag), "yyyy-MM-dd") Then
                intCount = intCount + 1
            End If
            
            rsTmp.MoveNext
        Loop
        If cboRegist.ListIndex = -1 Then
            Call zlControl.CboSetIndex(cboRegist.hwnd, 0)
        End If
        
        lbl��ƾ���.Visible = intCount > 1 And mintActive = pt����
        
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
    Call LoadPatiAllergy(mPatiInfo.����ID, , mrsAller)
    mrsAller.Filter = "����ID=" & mPatiInfo.����ID & " and �Һŵ�<>'" & mPatiInfo.�Һŵ� & "'"
    cmdAller.Enabled = mrsAller.RecordCount > 0
    
    Call LoadAllerInfo(mrsAller)
End Sub

Private Sub LoadAllerInfo(ByRef rsTmp As ADODB.Recordset)
    Dim i As Long
    Dim lngRow As Long
    
    With vsAller
        .Clear
        .Cols = 0   '��������м��п�
        .Cols = 1
        .Rows = 2
        .RowHidden(1) = True
        '��ʾ���ιҺŵĹ�����¼�����޸�
        If rsTmp.State = 1 Then
            rsTmp.Filter = "�Һŵ�='" & mstr�Һŵ� & "'"
            If rsTmp.RecordCount > 0 Then
                .Cols = rsTmp.RecordCount + 1
                For i = 0 To rsTmp.RecordCount - 1
                    '������Դ�Ŀ������ظ�
                    lngRow = -1
                    If Not IsNull(rsTmp!ҩ��ID) Then
                        lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                    ElseIf Not IsNull(rsTmp!ҩ����) Then
                        lngRow = .FindRow(CStr(rsTmp!ҩ����), 0)
                    End If
                    If lngRow = -1 Then
                        .TextMatrix(0, i) = "" & rsTmp!ҩ����
                        .Cell(flexcpData, 0, i) = .TextMatrix(0, i)   '�����ж��Ƿ��޸�
                        .Cell(flexcpData, 1, i) = rsTmp!����ʱ�� & ""
                        .ColData(i) = rsTmp!ҩ��ID & ""
                        .TextMatrix(1, i) = rsTmp!����Դ���� & ""
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
    
    If Button = 2 And InStr(mstrPrivs, "���˽���") > 0 Then
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
        
    cboEdit(cboȥ��).Width = IIf(mbytSize = 1, 1440, 1150)
    fraLine(cboȥ��).Width = cboEdit(cboȥ��).Width - 30
        
    lbl��ƾ���.Left = fraLine(cboȥ��).Left + fraLine(cboȥ��).Width + 30
    lbl��ƾ���.Top = optState(opt����).Top + 20
        
    lbl��.Top = fraLine(cbo�Ա�).Top - 60
    lbl��.Left = lngLeft - lbl��.Width - 30
    
    lblRec.Top = lbl��.Top
    lblRec.Left = lbl��.Left - lblRec.Width - 10
    
    If lbl��.Visible = False Then lblRec.Left = lbl��.Left
    
    lngTmp = lngLeft
    If lbl��.Visible Then
        lngTmp = lbl��.Left
    End If
    If lblRec.Visible Then
        lngTmp = lblRec.Left
    End If
    picExpand.Left = lngTmp - picExpand.Width - 150
        
    txtEdit(txt����ժҪ).Width = PicBasis.Width - txtEdit(txt����ժҪ).Left - 30
    txtEdit(txt������ַ).Width = PicBasis.Width - txtEdit(txt������ַ).Left - 30
    vsAller.Width = txtEdit(txt����ժҪ).Width - picPrompt.Width - 30 + txtEdit(txt��ͥ�绰).Left - txtEdit(txt�໤��).Left + txtEdit(txt��ͥ�绰).Width
    lblDiag(1).Width = PicPatiInfo.Width - lblDiag(1).Left
           
    
    '���1024*786���������⴦��
    '---------------------------------------------------------------------------------------------------------------------
    If mblnDocInput Then
        If lngLeft - 200 > fraLine(cbo����).Left Then
            rtfEdit(txt����).Width = (picPatiInput.ScaleWidth - lblDoc(txt����).Width * 4 - 100) / 2
        Else
            rtfEdit(txt����).Width = fraLine(cbo�Ա�).Left + fraLine(cbo�Ա�).Width - rtfEdit(txt����).Left
        End If
        
        lngCount = 1
        rtfEdit(txt����).Left = lblDoc(txt����).Left + lblDoc(txt����).Width + 70
        lblDoc(txt����).Top = 0
        For i = 1 To rtfEdit.Count - 1
            If rtfEdit(i).Visible Then
                lngCount = lngCount + 1
                lblDoc(i).Left = IIf(lngCount Mod 2 = 1, lblDoc(txt����).Left, rtfEdit(txt����).Left + rtfEdit(txt����).Width + 100)
                rtfEdit(i).Left = lblDoc(i).Left + lblDoc(i).Width + 70
                rtfEdit(i).Width = IIf(lngCount Mod 2 = 1, rtfEdit(txt����).Width, lngLeft - rtfEdit(i).Left - 100)
                lblDoc(i).Top = ((lngCount - 1) \ 2) * rtfEdit(i).Height - (((lngCount - 1) \ 2) * 15)
                rtfEdit(i).Top = ((lngCount - 1) \ 2) * rtfEdit(i).Height - (((lngCount - 1) \ 2) * 15)
            End If
        Next
        
        
        cmdSign.Left = lngLeft - 100 - cmdSign.Width
        cmdSign.Top = ((lngCount) \ 2) * rtfEdit(txt����).Height + 50
        cmdUpdate.Left = cmdSign.Left
        cmdUpdate.Top = cmdSign.Top + cmdSign.Height + 20
        cmdImportEPRDemo.Left = cmdUpdate.Left - cmdImportEPRDemo.Width - 50
        cmdImportEPRDemo.Top = cmdUpdate.Top
        
        lblҽ��(1).Left = cmdSign.Left - lblҽ��(1).Width - 100
        lblҽ��(0).Left = lblҽ��(1).Left - lblҽ��(0).Width - 20
        lblҽ��(0).Top = ((lngCount) \ 2) * rtfEdit(txt����).Height + 150
        lblҽ��(1).Top = lblҽ��(0).Top
        
        picPrompt.Left = rtfEdit(txt����).Left + rtfEdit(txt����).Width + 120
        picPrompt.Top = ((lngCount) \ 2) * rtfEdit(txt����).Height + 550
        lbl��ʾ.Left = picPrompt.Left + picPrompt.Width + 60
        lbl��ʾ.Top = ((lngCount) \ 2) * rtfEdit(txt����).Height + 550
        
        lbl��������.Left = rtfEdit(txt����).Left + rtfEdit(txt����).Width + 400
        lbl��������.Top = ((lngCount) \ 2) * rtfEdit(txt����).Height + 250
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
 
    strSQL = "Select ����,���� From �Ա�"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cbo�Ա�), rsTmp, True)
               
    Call SetCboFromList(Array("��", "��", "��"), cboEdit(cbo����), 0)
    Call SetCboFromList(Array(" ", "Сʱǰ", "��ǰ", "��ǰ", "��ǰ", "��ǰ"), cboEdit(cbo����ʱ��), 0)
    
    strSQL = "Select ����,���� From ְҵ"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cboְҵ), rsTmp, True)
    
    strSQL = "Select ����, ���� From ����ȥ��"
    cboEdit(cboȥ��).Clear
    cboEdit(cboȥ��).AddItem ("")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cboȥ��), rsTmp, False)
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
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "15����"
        .ItemData(.NewIndex) = 15
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    
    '���ﲡ��ʱ�䷶Χ
    curDate = zlDatabase.Currentdate
    
    intStart = Val(zlDatabase.GetPara("���ﲡ�˽������", glngSys, p����ҽ��վ, "0", Array(lbl����ʱ��, cboSelectTime), InStr(";" & mstrPrivs & ";", ";��������;") > 0))
    If lbl����ʱ��.ForeColor <> vbBlue Then
        '˽�в���
        mvCondFilter.End = Format(curDate, "yyyy-MM-dd 23:59:59")
        mvCondFilter.Begin = Format(mvCondFilter.End, "yyyy-MM-dd 00:00:00")
                If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
    Else
        'ϵͳ����(�ָ��ɹ���Ա���õ�ֵ����ֹͨ��)
        mvCondFilter.End = Format(curDate + intStart, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("���ﲡ�˿�ʼ���", glngSys, p����ҽ��վ, "7", Array(lbl����ʱ��, cboSelectTime), InStr(";" & mstrPrivs & ";", ";��������;") > 0))
        If intDay > 7 Then intDay = 7
        mvCondFilter.Begin = Format(mvCondFilter.End - intDay, "yyyy-MM-dd 00:00:00")
        cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
        lbl����ʱ��.ToolTipText = cboSelectTime.ToolTipText
        If intDay = 7 And intStart = 0 Then
            cboSelectTime.ListIndex = 1
                ElseIf intDay = 0 And intStart = 0 Then
                        cboSelectTime.ListIndex = 0
        Else
            cboSelectTime.ListIndex = 4
        End If
    End If
    
    'ȱʡҽ������
    mvCondFilter.ҽ�� = UserInfo.����
    
    '������ȱʡ
    mvCondFilter.�Һŵ� = ""
    mvCondFilter.���￨ = ""
    mvCondFilter.����ID = 0
    mvCondFilter.����� = ""
    mvCondFilter.���� = ""
    
End Sub

Private Sub GetLocalSetting()
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    '���ﷶΧ��1=�ұ��˺ŵĲ���,2=�����Ҳ���,3=�����Ҳ���
    Dim strSQL As String, rsTmp As Recordset, intType As Integer
    Dim str���˽������ As String '�����:57566
    
    mint���ﷶΧ = Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2"))
    mstr�������� = zlDatabase.GetPara("��������", glngSys, p����ҽ��վ)
    mlng�������ID = Val(zlDatabase.GetPara("�������", glngSys, p����ҽ��վ))
    On Error GoTo errH
    strSQL = "Select Distinct B.ID,B.����,B.����,A.ȱʡ" & _
        " From ������Ա A,���ű� B,��������˵�� C" & _
        " Where A.����ID=B.ID And B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And A.��ԱID=[1] And b.ID=[2]" & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, mlng�������ID)
    If rsTmp.RecordCount = 0 Then mlng�������ID = 0
    mblnҪ����� = Val(zlDatabase.GetPara("ֻ�����Ѿ�����Ĳ���", glngSys, p����ҽ��վ)) <> 0
    
    '���ﲡ��
    If InStr(mstrPrivs, "���ﲡ��") > 0 Then
        mstr����ҽ�� = zlDatabase.GetPara("����ҽ��", glngSys, p����ҽ��վ, UserInfo.����)
    Else
        mstr����ҽ�� = UserInfo.����
    End If
    
    '�Զ�������
    mbln�Զ����� = Val(zlDatabase.GetPara("�ҵ����˺��Զ�����", glngSys, p����ҽ��վ)) <> 0
    mlng�Զ����� = Val(zlDatabase.GetPara("������Զ�����", glngSys, p����ҽ��վ))
    
    'ҽ���������к���������
    mbln���к���� = Val(zlDatabase.GetPara("ҽ���������к���������", glngSys, p����ҽ��վ)) <> 0
    '��������
    mbytSize = zlDatabase.GetPara("����", glngSys, p����ҽ��վ, "0")


    mintFindType = Val(zlDatabase.GetPara("���˲��ҷ�ʽ", glngSys, p����ҽ��վ, "1", , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    
    '�����:57566
    str���˽������ = CStr(zlDatabase.GetPara("���˽������", glngSys, p����ҽ��վ))
    If str���˽������ <> "" Then
        mlng������� = Val(Left(str���˽������, 1))
        If UBound(Split(str���˽������, "|")) >= 1 Then
            mlng��ǰ����ʱ�� = Val(Split(str���˽������, "|")(1))
        End If
    End If
    
    '�����Զ�ˢ��
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
'���ܣ���ȡ�����б�
'������intActive,strActNO=ˢ�º���Ҫ��λ���б������Ͳ��˹Һŵ�(�����)
'      ע���������ָ����intActive,�����Ҫ����strRefeshˢ���б���
'      strRefesh=�ֱ��Ƿ�ˢ��ָ�����б��ֱ�Ϊ"���������ת�ԤԼ,����"
    Dim rsPati As New ADODB.Recordset
    Dim objItem As ListItem, intIdx As PatiType
    Dim strKeep As String, strPrePati As String
    Dim strSQL As String, i As Long, j As Long
    Dim strTime As String, blnRefresh As Boolean
    Dim objLvw As ListView
    Dim lngColor As Long, lngPatiTypeIdx As Long
    Dim rs��Ⱦ��״̬ As ADODB.Recordset
    Dim blnDo��Ⱦ��״̬ As Boolean
    Dim bln��ҽ As Boolean
    
    strPrePati = mstrPrePati '��ΪҪ�ƻ�,�����ʱ��¼
    
    Screen.MousePointer = 11
    On Error GoTo errH
    mblnUnRefresh = True
    strSQL = "select  m.����id,m.id,m.no,max(m.��¼) as ��¼,max(m.��д) as ��д,max(m.״̬) as ״̬ from" & vbNewLine & _
        "(select a.����id,a.id, a.no,1 as ��¼,0 as ��д,0 as ״̬ from ���˹Һż�¼ a,�������Լ�¼ b" & vbNewLine & _
        "where a.no=b.�Һŵ� and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.����id,a.id, a.no,0 as ��¼,1 as ��д,0 as ״̬" & vbNewLine & _
        "From ���˹Һż�¼ A, ���Ӳ�����¼ C, �����ļ��б� D" & vbNewLine & _
        "Where c.�ļ�id = d.Id And d.���� = 5 And a.����id = c.����id And a.id = c.��ҳid and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.����id,a.id, a.no,0 as ��¼,1 as ��д,e.����״̬ as ״̬" & vbNewLine & _
        "From ���˹Һż�¼ A,���Ӳ�����¼ C,�����ļ��б� D,�����걨��¼ E" & vbNewLine & _
        "Where a.����id = c.����id And a.id = c.��ҳid and c.id=e.�ļ�id and d.����=5 and e.�ļ�id =d.id and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1) M" & vbNewLine & _
        "group by m.����id,m.id,m.no"
    Set rs��Ⱦ��״̬ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ҽ��)
    If rs��Ⱦ��״̬.RecordCount > 0 Then blnDo��Ⱦ��״̬ = True
    
    For intIdx = 0 To lvwPati.UBound + 3
        If Mid(strRefesh, intIdx + 1, 1) = "1" Then
            If intIdx = pt���� Then    '���ﲡ��
                '���ﷶΧ
                If mint���ﷶΧ = 1 Then
                    strSQL = " And B.ִ����||''=[2]" '�ұ��˺�
                    If mblnҪ����� Then strSQL = strSQL & " And B.���� is Not NULL"
                ElseIf mint���ﷶΧ = 2 Then
                    '������
                    If mlng�������ID <> 0 Then
                        strSQL = " And B.����=[3] And b.ִ�в���id+0 =[4] And (B.ִ����||''=[2] Or B.ִ���� Is Null) "
                    Else    '10.28��ǰѡ����ʱû�ж�����
                        strSQL = " And B.����=[3] And (B.ִ����||''=[2] Or B.ִ���� Is Null) " & _
                            "And Exists (Select ����id" & vbNewLine & _
                            " From �ҺŰ��� F, ������Ա D" & vbNewLine & _
                            " Where D.��Աid = [6] And F.����id = D.����id And b.ִ�в���id = F.����id)"
                    End If
                ElseIf mint���ﷶΧ = 3 Then
                    strSQL = " And B.ִ�в���ID+0=[4] And (B.ִ����||''=[2] Or B.ִ���� Is Null)" '������
                    If mblnҪ����� Then strSQL = strSQL & " And B.���� is Not NULL"
                End If
                
                strSQL = _
                    " Select /*+ Rule*/B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����," & _
                    "       B.����ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����," & _
                    "       B.����,B.����,B.����ʱ��,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
                    "       B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
                    " From ������Ϣ A,���˹Һż�¼ B,���ű� C" & _
                    " Where B.����ID=A.����ID And (Nvl(B.ִ��״̬,0)=0 or nvl(B.ִ��״̬,0)=[5]) And B.ת�����ID=C.ID(+) And B.��¼����=1 And B.��¼״̬=1" & _
                    "       And B.ִ��ʱ�� is Null And B.����ʱ�� <= Trunc(Sysdate)+1-1/24/60/60 " & strSQL & _
                    IIf(gint��ͨ�Һ����� = gint����Һ�����, " And B.����ʱ��>=Sysdate-" & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), _
                    " And B.����ʱ�� >= Sysdate-" & IIf(gint��ͨ�Һ����� > gint����Һ�����, gint��ͨ�Һ�����, gint����Һ�����) & " And B.����ʱ��>=Sysdate-Decode(B.����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ")") & _
                    " Order By Decode(B.����ʱ��,NULL,2,1),B.����ʱ��,B.NO"
                '"Sysdate-Decode(B.����"��������ʧЧ�����Լ��˶��������
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "δ��", UserInfo.����, mstr��������, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mblnShowLeavePati, -1, 0), UserInfo.ID)
            ElseIf intIdx = pt���� Then '���ﲡ��
                strSQL = _
                    " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����," & _
                    " B.ִ��ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
                    " B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
                    " From ������Ϣ A,���˹Һż�¼ B,���ű� C" & _
                    " Where B.����ID=A.����ID And B.ת�����ID=C.ID(+)" & _
                    " And B.ִ��״̬=2 and nvl(B.��¼��־,0)<=1 And B.ִ����||''=[1] And B.��¼����=1 And B.��¼״̬=1" & _
                    " Order By B.NO"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ҽ��)
            ElseIf intIdx = pt���� Then '���ﲡ��
                strSQL = "Select /*+ Rule*/" & vbNewLine & _
                    " Distinct(b.No), b.����id, b.�����, b.����, b.�Ա�, b.����, b.����, b.����, b.����, b.ִ��ʱ�� As ʱ��, a.���￨��, a.���֤��, a.Ic����, a.����, b.����ʱ��, b.ִ�в���id," & vbNewLine & _
                    " b.ִ����, b.ִ��״̬, b.��¼��־, a.��������," & vbNewLine & _
                    "First_Value(Decode(Sign(h.������� - 10), -1, h.�������, '')) Over(Partition By h.����id, h.��ҳid Order By Sign(h.������� - 10), Decode(h.��¼��Դ, 4, 0, h.��¼��Դ) Desc, Decode(h.�������, 1, 1, 0) Desc, h.��ϴ���) As ��ҽ���," & vbNewLine & _
                    "First_Value(Decode(Sign(h.������� - 10), 1, h.�������, '')) Over(Partition By h.����id, h.��ҳid Order By -Sign(h.������� - 10), Decode(h.��¼��Դ, 4, 0, h.��¼��Դ) Desc, Decode(h.�������, 11,11, 0) Desc, h.��ϴ���) As ��ҽ���" & vbNewLine & _
                    "From ������Ϣ A, ���˹Һż�¼ B, ������ϼ�¼ H" & IIf(mvCondFilter.���￨ <> "", ",����ҽ�ƿ���Ϣ C, ҽ�ƿ���� D", "") & vbNewLine & _
                    "Where b.����id = a.����id And h.����id(+) = b.����id And h.��ҳid(+) = b.id And b.ִ��״̬ + 0 = 1 And b.��¼���� = 1 And b.��¼״̬ = 1" & _
                     IIf(mvCondFilter.���￨ <> "", " And c.����id = a.����id And c.�����id = d.Id And d.�Ƿ�̶� = 1 And d.���� = '���￨' ", "")
              
                If mvCondFilter.�Һŵ� <> "" Then
                    strSQL = strSQL & " And B.NO=[5]"
                ElseIf mvCondFilter.����� <> "" Then
                    strSQL = strSQL & " And A.�����=[6]"
                ElseIf mvCondFilter.���￨ <> "" Then
                    strSQL = strSQL & " And C.����=[7]"
                
                Else
                    strSQL = strSQL & " And B.ִ��ʱ�� Between To_Date('" & Format(mvCondFilter.Begin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mvCondFilter.End, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    strSQL = strSQL & IIf(mvCondFilter.ҽ�� = "", "", " And B.ִ����||''=[3]")
                    If mvCondFilter.����ID <> 0 Then strSQL = strSQL & " And B.ִ�в���ID+0=[4]"
                                        If mvCondFilter.���� <> "" Then strSQL = strSQL & " And A.����=[8]"
                End If
                
                If zlDatabase.DateMoved(mvCondFilter.Begin) Then
                    strSQL = strSQL & " Union ALL " & Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
                End If

                strSQL = strSQL & " Order By NO Desc"
                
                With mvCondFilter
                    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "δ��", "δ��", .ҽ��, .����ID, .�Һŵ�, .�����, .���￨, .����)
                End With
            ElseIf intIdx = pt���� Then    '���ﲡ��
                strSQL = _
                    " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����," & _
                    " B.ִ��ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
                    " B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
                    " From ������Ϣ A,���˹Һż�¼ B,���ű� C" & _
                    " Where B.����ID=A.����ID And B.ת�����ID=C.ID(+) And B.��¼����=1 And B.��¼״̬=1" & _
                    " And B.ִ��״̬=2 and nvl(B.��¼��־,0) in (2,3) And B.ִ����||''=[1]" & _
                    " Order By B.NO"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ҽ��)
            ElseIf intIdx = lvwPati.UBound + 1 Then   'ת�ﲡ��
                '���ﷶΧ
                If mint���ﷶΧ = 1 Then
                    strSQL = " And B.ת��ҽ��=[2]" 'ת���˺�
                ElseIf mint���ﷶΧ = 2 Then
                    'ת�����ң���������ת�ģ�����ҽ�������ѻ���δָ������ҽ��
                    strSQL = " And B.ת������=[3] And B.ת�����ID=[4] And Nvl(B.ִ����,'��')<>[2] And (B.ת��ҽ��=[2] Or B.ת��ҽ�� Is NULL)"
                ElseIf mint���ﷶΧ = 3 Then
                    'ת�����ң���������ת�ģ�����ҽ�������ѻ���δָ������ҽ��
                    strSQL = " And B.ת�����ID=[4] And Nvl(B.ִ����,'��')<>[2] And (B.ת��ҽ��=[2] Or B.ת��ҽ�� Is NULL)"
                End If
                strSQL = _
                    " Select /*+ Rule*/B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����,B.ִ����," & _
                    " B.����ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,B.����ʱ��,B.ת�����ID as ִ�в���ID," & _
                    " B.ת��״̬,C.���� as ת�����,B.���� as ת������,B.ִ���� as ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
                    " From ������Ϣ A,���˹Һż�¼ B,���ű� C" & _
                    " Where B.����ID=A.����ID And B.ת��״̬=0 And B.ִ�в���ID=C.ID And B.��¼����=1 And B.��¼״̬=1" & strSQL & _
                    IIf(gint��ͨ�Һ����� = gint����Һ�����, " And B.����ʱ��>=Sysdate-" & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), _
                    " And B.����ʱ�� >= Sysdate-" & IIf(gint��ͨ�Һ����� > gint����Һ�����, gint��ͨ�Һ�����, gint����Һ�����) & " And B.����ʱ��>=Sysdate-Decode(B.����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ")") & _
                    " Order By B.NO"
                '"Sysdate-Decode(B.����"��������ʧЧ�����Լ��˶��������
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "δ��", UserInfo.����, mstr��������, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), 0, 0)
            ElseIf intIdx = lvwPati.UBound + 2 Then   'ԤԼ����
                '���ﷶΧ
                If mint���ﷶΧ = 1 Then
                    strSQL = " And A.ִ����||''=[1]" '�ұ��˺�
                                                            
                ElseIf mint���ﷶΧ = 2 Or mint���ﷶΧ = 3 Then '�����ң�ԤԼ�Һŵķ�ҩ���������Ԥ�ţ�û�����ң���������
                    strSQL = " And A.ִ�в���ID+0=[2] And (A.ִ����||''=[1] Or A.ִ���� Is Null)"
                End If


                '�������ڵ�ʱ��Σ��ñ����ӵķ�ʽ�������
                strTime = _
                    "Select ʱ��� From ʱ��� Where" & _
                    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                    " Between" & _
                    " Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-09 '||To_Char(��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS'))" & _
                    " And" & _
                    " '3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS'))" & _
                    " Or" & _
                    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                    " Between" & _
                    " '3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS')" & _
                    " And" & _
                    " Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-11 '||To_Char(��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS')))"
                
                'ȡ���ڵ���������Ӧ���ŵ�ʱ���
                strTime = " And Decode(To_Char(SysDate,'D'),'1',B.����,'2',B.��һ,'3',B.�ܶ�,'4',B.����,'5',B.����,'6',B.����,'7',B.����,NULL) IN(" & strTime & ")"
                strSQL = "Select A.NO,A.����ID,A.��ʶ�� as �����,A.����,A.�Ա�,A.����,A.�Ӱ��־ as ����,A.ִ����," & _
                    " A.����ʱ�� as ʱ��,C.���￨��,C.���֤��,C.IC����,C.����,A.����ʱ��,A.ִ�в���ID,0 as ִ��״̬,0 as ��¼��־,C.��������" & _
                    " From ������ü�¼ A,�ҺŰ��� B,������Ϣ C" & _
                    " Where A.���㵥λ=B.���� And A.����ID=C.����ID(+) And A.���=1" & _
                    " And A.��¼����=4 And A.��¼״̬=0 " & strTime & strSQL & _
                    " And A.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID))
            End If
            
            '��¼�б�����ѡ�еĲ���
            strKeep = ""
            If intIdx = lvwPati.UBound + 1 Then 'ת�ﲡ��
                If Not lvwIncept.SelectedItem Is Nothing Then
                    strKeep = lvwIncept.SelectedItem.Key
                End If
                lvwIncept.ListItems.Clear
            ElseIf intIdx = lvwPati.UBound + 2 Then 'ԤԼ����
                If Not lvwReserve.SelectedItem Is Nothing Then
                    strKeep = lvwReserve.SelectedItem.Key
                End If
                lvwReserve.ListItems.Clear
            ElseIf intIdx = pt���� Then   '���ﲡ��
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
                If intIdx = lvwPati.UBound + 1 Then 'ת�ﲡ��
                    Set objItem = lvwIncept.ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , intIdx + 1)
                ElseIf intIdx = lvwPati.UBound + 2 Then 'ԤԼ����
                    Set objItem = lvwReserve.ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , 1)
                ElseIf intIdx = 5 Then  'pt����
                    Set objItem = lvwPatiHZ.ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , 1)
                Else
                    Set objItem = lvwPati(intIdx).ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , intIdx + 1)
                End If
                objItem.SubItems(1) = Nvl(rsPati!�����)
                objItem.SubItems(2) = Nvl(rsPati!����)
                objItem.SubItems(3) = Nvl(rsPati!�Ա�)
                objItem.SubItems(4) = Nvl(rsPati!����)
                objItem.SubItems(5) = IIf(Nvl(rsPati!����, 0) <> 0, "��", "")
                
                If intIdx = lvwPati.UBound + 2 Then
                    'ԤԼ����
                    objItem.SubItems(6) = Nvl(rsPati!ִ����)
                    objItem.SubItems(7) = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
                    objItem.SubItems(8) = Nvl(rsPati!���֤��)
                    objItem.SubItems(9) = Nvl(rsPati!���￨��)
                    objItem.SubItems(10) = Nvl(rsPati!��������)
                    lngPatiTypeIdx = 10
                Else
                    objItem.SubItems(6) = IIf(Nvl(rsPati!����, 0) <> 0, "��", "")
                    objItem.SubItems(7) = IIf(Nvl(rsPati!����, 0) <> 0, "��", "")
                    If intIdx = pt���� Then
                        objItem.SubItems(8) = Nvl(rsPati!����)
                        objItem.SubItems(9) = Nvl(rsPati!ִ����)
                        objItem.ListSubItems(9).Tag = Nvl(rsPati!ִ��״̬)
                        objItem.SubItems(10) = LPAD(Nvl(rsPati!����), 5, " ")
                        objItem.SubItems(11) = Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(12) = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(13) = Nvl(rsPati!���￨��)
                        objItem.SubItems(14) = Nvl(rsPati!��������)
                        lngPatiTypeIdx = 14
                    ElseIf intIdx = pt���� Then
                        objItem.SubItems(8) = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(9) = Nvl(rsPati!ִ����)
                        objItem.SubItems(10) = Nvl(rsPati!���￨��)
                        objItem.SubItems(11) = Nvl(rsPati!��������)
                        objItem.SubItems(12) = Nvl(rsPati!��ҽ���)
                        objItem.SubItems(13) = Nvl(rsPati!��ҽ���)
                        If rsPati!��ҽ��� & "" <> "" Then bln��ҽ = True
                        lngPatiTypeIdx = 13
                    ElseIf intIdx = pt���� Then  '����
                        If Nvl(rsPati!��¼��־, "0") = 2 Then
                                objItem.SmallIcon = "��ͣ"
                        End If
                        objItem.SubItems(8) = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(10) = Nvl(rsPati!���￨��)
                        objItem.SubItems(11) = Nvl(rsPati!��������)
                        lngPatiTypeIdx = 11
                        '��Ӵ�Ⱦ��״̬
                        strSQL = ""
                        If blnDo��Ⱦ��״̬ Then
                            rs��Ⱦ��״̬.Filter = "no='" & rsPati!NO & "'"
                            If Not rs��Ⱦ��״̬.EOF Then strSQL = Get��Ⱦ��״̬(Val(rs��Ⱦ��״̬!��¼ & ""), Val(rs��Ⱦ��״̬!��д & ""), Val(rs��Ⱦ��״̬!״̬ & ""))
                        End If
                        objItem.SubItems(12) = strSQL
                    Else
                        objItem.SubItems(8) = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(9) = Nvl(rsPati!���￨��)
                        objItem.SubItems(10) = Nvl(rsPati!��������)
                        lngPatiTypeIdx = 10
                        If intIdx = pt���� Then
                            '��Ӵ�Ⱦ��״̬
                            strSQL = ""
                            If blnDo��Ⱦ��״̬ Then
                                rs��Ⱦ��״̬.Filter = "no='" & rsPati!NO & "'"
                                If Not rs��Ⱦ��״̬.EOF Then strSQL = Get��Ⱦ��״̬(Val(rs��Ⱦ��״̬!��¼ & ""), Val(rs��Ⱦ��״̬!��д & ""), Val(rs��Ⱦ��״̬!״̬ & ""))
                            End If
                            objItem.SubItems(12) = strSQL
                        End If
                    End If
                End If
                objItem.ListSubItems(1).Tag = Nvl(rsPati!���￨��)
                objItem.ListSubItems(2).Tag = Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                objItem.ListSubItems(3).Tag = Nvl(rsPati!ִ�в���ID, 0)
                objItem.ListSubItems(4).Tag = Nvl(rsPati!ִ����)
                objItem.ListSubItems(5).Tag = "" '�����¼ת��״̬
                objItem.ListSubItems(6).Tag = Nvl(rsPati!���֤��)
                objItem.ListSubItems(7).Tag = Nvl(rsPati!IC����)
                objItem.ListSubItems(8).Tag = Val(Nvl(rsPati!��¼��־))
                objItem.Tag = rsPati!����ID
                
                'ת��״̬:��ʾ�����һ��
                If intIdx = pt���� Or intIdx = pt���� Then
                    If intIdx = pt���� Then
                        j = lvwPati(intIdx).ColumnHeaders.Count - 2
                    Else
                        j = lvwPati(intIdx).ColumnHeaders.Count - 1
                    End If
                
                    objItem.ListSubItems(5).Tag = Nvl(rsPati!ת��״̬) 'Null��0��ͬ
                    If Not IsNull(rsPati!ת��״̬) Then
                        If rsPati!ת��״̬ = 0 Then
                            '�Ѿ�ת��
                            objItem.SmallIcon = "ת��"
                            objItem.SubItems(j) = "���Է�����,����:" & rsPati!ת����� & _
                                IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                                IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                        ElseIf rsPati!ת��״̬ = -1 Then
                            '�Ѿܾ�ת��
                            objItem.SmallIcon = "�ܾ�"
                            objItem.SubItems(j) = "�Է��Ѿܾ�,����:" & rsPati!ת����� & _
                                IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                                IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                        ElseIf rsPati!ת��״̬ = 1 Then
                            '�ѽ���ת��
                        End If
                    End If
                ElseIf intIdx = lvwPati.UBound + 1 Then
                    'ת�ﲡ��
                    objItem.SmallIcon = "����"
                    objItem.SubItems(lvwIncept.ColumnHeaders.Count - 1) = "������ת��,����:" & rsPati!ת����� & _
                        IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                        IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                End If
                
                '��ʾ������ɫ
                lngColor = zlDatabase.GetPatiColor(Nvl(rsPati!��������))
                objItem.ListSubItems(1).ForeColor = lngColor
                objItem.ListSubItems(lngPatiTypeIdx).ForeColor = lngColor
                
                '���ղ����ú�ɫ��ʾ
                If Not IsNull(rsPati!����) And rsPati!�������� & "" = "" Then
                    objItem.ListSubItems(1).ForeColor = &HC0&
                    objItem.ListSubItems(lngPatiTypeIdx).ForeColor = &HC0&
                End If
                
                '�����־��ɫͻ����ʾ
                If Nvl(rsPati!����, 0) <> 0 Then
                    objItem.ListSubItems(5).ForeColor = vbRed
                End If
                
                '��λ��ָ������
                If objItem.Key = "_" & strActNO Then
                    mstrPrePati = "_" & strActNO '���⼤���¼�
                    objItem.Selected = True
                    strKeep = ""
                End If
                '��λ��ԭ�Ȳ���
                If objItem.Key = strKeep And Me.Visible Then
                    mstrPrePati = strKeep '���⼤���¼�
                    objItem.Selected = True
                End If
                  If intIdx = pt���� Then
                        If Val(objItem.ListSubItems(9).Tag) = -1 Then
                            objItem.ForeColor = &H808080
                            For j = 1 To objItem.ListSubItems.Count
                                objItem.ListSubItems(j).ForeColor = &H808080
                            Next
                        End If
                  End If
                rsPati.MoveNext
            Next
            
            '��ҽ���Ϊ��ʱ����
            If intIdx = pt���� Then
                lvwPati(pt����).ColumnHeaders(14).Width = IIf(bln��ҽ, 3000, 0)
            End If
            
            'ˢ������æ��״̬
            If intIdx = pt���� Then
                Call SetRoomState(lvwPati(intIdx).ListItems.Count > 0)
            End If
        End If
    Next
    
    '�����б����
    Call RefreshTitle
    
    '�����ǰ��б�δҪ��ˢ����������,���ظ�����ˢ��
    blnRefresh = True
    If mintActive <> -1 Then
        If mintActive = pt���� Then
            '����
            Set objLvw = lvwPatiHZ
        ElseIf mintActive = ptԤԼ Then
            Set objLvw = lvwReserve
        ElseIf mintActive = ptת�� Then
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
        
    'ȷ��ˢ�º������б�ˢ��������:ȱʡ��Ϊ���ﲡ���б�
    If blnRefresh Then
        If intActive = -1 Then intActive = mintActive
        If intActive = pt���� Then
            '����
            Set objLvw = lvwPatiHZ
        ElseIf intActive = ptԤԼ Then
            Set objLvw = lvwReserve
        ElseIf intActive = ptת�� Then
            Set objLvw = lvwIncept
        ElseIf intActive <> -1 Then
            Set objLvw = lvwPati(intActive)
        End If
            
        If intActive = -1 Then
            If lvwPati(pt����).ListItems.Count > 0 Then
                intActive = 0
            ElseIf lvwPati(pt����).ListItems.Count > 0 Then
                intActive = 1
            End If
        ElseIf objLvw.ListItems.Count = 0 Then
            If lvwPati(pt����).ListItems.Count > 0 Then
                intActive = 0
            ElseIf lvwPati(pt����).ListItems.Count > 0 Then
                intActive = 1
            Else
                intActive = -1
            End If
        End If
        
        If intActive = pt���� Then
            '����
            Set objLvw = lvwPatiHZ
        ElseIf intActive = ptԤԼ Then
            Set objLvw = lvwReserve
        ElseIf intActive = ptת�� Then
            Set objLvw = lvwIncept
        ElseIf intActive <> -1 Then
            Set objLvw = lvwPati(intActive)
        End If
        
        'ˢ�²��˵��������
        mintActive = -1
        'Ĭ������һ��ֵ
        If Not Me.Visible Then mintActive = 1
        mstrPrePati = ""
        If intActive <> -1 And Me.Visible Then
            objLvw.SelectedItem.EnsureVisible
            Call LvwItemClick(CInt(intActive), objLvw.SelectedItem)
            'Call lvwPati_ItemClick(CInt(intActive), lvwPati(intActive).SelectedItem)
        Else
            
            '����ǰ�б�������ˢ���Ӵ���
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
'���ܣ������б����
    Dim i As Integer
    
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).Title Like "���ﲡ��*" Then
            dkpMain.Panes(i).Title = "���ﲡ��" & IIf(lvwPati(pt����).ListItems.Count = 0, "", ":" & lvwPati(pt����).ListItems.Count & "��")
        ElseIf dkpMain.Panes(i).Title Like "*���ﲡ��*" Then
            If mstr����ҽ�� <> UserInfo.���� Then
                dkpMain.Panes(i).Title = mstr����ҽ�� & "�ľ��ﲡ��" & IIf(lvwPati(pt����).ListItems.Count = 0, "", ":" & lvwPati(pt����).ListItems.Count & "��")
            Else
                dkpMain.Panes(i).Title = UserInfo.���� & "�ľ��ﲡ��" & IIf(lvwPati(pt����).ListItems.Count = 0, "", ":" & lvwPati(pt����).ListItems.Count & "��")
            End If
        ElseIf dkpMain.Panes(i).Title Like "���ﲡ��*" Then
            dkpMain.Panes(i).Title = "���ﲡ��" & IIf(lvwPati(pt����).ListItems.Count = 0, "", ":" & lvwPati(pt����).ListItems.Count & "��")
        ElseIf dkpMain.Panes(i).Title Like "ת�ﲡ��*" Then
            dkpMain.Panes(i).Title = "ת�ﲡ��" & IIf(lvwIncept.ListItems.Count = 0, "", ":" & lvwIncept.ListItems.Count & "��")
        ElseIf dkpMain.Panes(i).Title Like "ԤԼ����*" Then
            dkpMain.Panes(i).Title = "ԤԼ����" & IIf(lvwReserve.ListItems.Count = 0, "", ":" & lvwReserve.ListItems.Count & "��")
        ElseIf dkpMain.Panes(i).Title Like "���ﲡ��*" Then
            dkpMain.Panes(i).Title = "���ﲡ��" & IIf(lvwPatiHZ.ListItems.Count = 0, "", ":" & lvwPatiHZ.ListItems.Count & "��")
        End If
    Next
End Sub

Private Sub ClearPatiInfo()
'���ܣ��������������ص���ʾ��Ϣ
    Dim i As Long
    
    cboRegist.Tag = "Loading"
    mlng����ID = 0
    mstr�Һŵ� = ""
    mlng����ID = 0
    mlng�Һ�ID = 0
    mPatiInfo.���� = 0
    mPatiInfo.����� = ""
    mPatiInfo.�Һŵ� = ""
    mPatiInfo.�Һ�ID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.���� = ""
    mPatiInfo.���� = 0
    mPatiInfo.������ = ""
    mPatiInfo.�Һ�ʱ�� = CDate(0)
    mPatiInfo.����ת�� = False
    mPatiInfo.�����ļ�id = 0
    mPatiInfo.����id = 0
    mPatiInfo.�Ƿ�ǩ�� = False
    mPatiInfo.������ = ""
    mPatiInfo.����״�� = ""
    mPatiInfo.�Ա� = ""
    mPatiInfo.���� = ""
    mPatiInfo.���� = ""
    mPatiInfo.���� = ""
    mPatiInfo.�����ص� = ""
    mPatiInfo.��Ⱦ���ϴ� = 0
    mPatiInfo.��ͥ��ַ�ʱ� = ""
    mPatiInfo.��λ�ʱ� = ""
    mPatiInfo.����֤�� = ""
    mPatiInfo.����ID = 0
        
    cboRegist.Clear
    lbl��ƾ���.Visible = False
    lbl��.Visible = False
    lblRec.Visible = False
            
    For i = 0 To txtEdit.Count - 1
        txtEdit(i).Text = ""
    Next
    If mblnDocInput Then
        For i = 0 To rtfEdit.Count - 1
            rtfEdit(i).Text = ""
        Next
    End If
    txt��������.Text = "____-__-__"
    txt��������.Text = "____-__-__"
    txt����ʱ��.Text = "__:__"
    txt����ʱ��.Text = "__:__"
            
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
'���ܣ����˹Һ�
    Dim strCommon As String, intAtom As Integer
    Dim strNO As String, blnPrice As Boolean
    Dim objControl As CommandBarControl
    
    blnPrice = Val(zlDatabase.GetPara("����ҺŻ��۵�", glngSys, p����ҽ��վ, 1)) = 1
    
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    mblnUnRefresh = True
    '��������(����Ϸ�������)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    strNO = gobjRegist.StationRegister(Me, gcnOracle, glngSys, mstr��������, InStr(mstrPrivs, "�Һŷѱ����") = 0, blnPrice, , gstrDBUser)
    Call GlobalDeleteAtom(intAtom)
        
    'ˢ�²���λ���չҺŵĲ�����
    If strNO <> "" And lvwPati(pt����).Visible Then
        Call LoadPatients("11000", pt����, strNO)
        lvwPati(pt����).SetFocus
        
        '����֮���Զ�����ҽ���´�״̬
        If mlng�Զ����� = 1 Then
            If tbcSub.Selected.Tag <> "ҽ��" Then tbcSub.Item(0).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
            Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        ElseIf mlng�Զ����� = 2 Then
            If tbcSub.Selected.Tag <> "����" Then tbcSub.Item(1).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
            mblnUnRefresh = True
            Call mclsEPRs.zlOpenDefaultEPR(mstr�Һŵ�)
        End If
    Else
        Call LoadPatients("11000")
    End If
    mblnUnRefresh = False
End Sub

Private Sub ExecuteBespeak()
'���ܣ�ԤԼ�Һ�
    Dim strCommon As String, intAtom As Integer, strNO As String
            
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    mblnUnRefresh = True
    '��������(����Ϸ�������)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    strNO = gobjRegist.StationBespeak(Me, gcnOracle, glngSys, "", InStr(mstrPrivs, "�Һŷѱ����") = 0, mlng����ID, gstrDBUser)
    Call GlobalDeleteAtom(intAtom)
    mblnUnRefresh = False
End Sub

Private Sub ExecuteBespeakPrint()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش�ԤԼ�Һŵ�
    '����:���˺�
    '����:2012-12-24 10:55:39
    '˵��:
    '����:56274
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
    '��������(����Ϸ�������)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
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
'���ܣ�����ת��
    Dim rsTmp As New ADODB.Recordset
    Dim lng����ID As Long, str���� As String
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strSQL As String, objLvw As ListView
    
    If mintActive = pt���� Then
        Set objLvw = lvwPatiHZ
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    
    With objLvw.SelectedItem
        If mstr�Һŵ� = "" Then
            MsgBox "����ѡ���ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If mintActive = pt���� Then
            If zlDatabase.NOMoved("������ü�¼", mstr�Һŵ�, "��¼����=", "4") Then
                MsgBox "�ò��˵ĹҺŷ����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '���Һŵ�ʱ��
        If BillExpend(mstr�Һŵ�) Then
            MsgBox "�ò��˹Һ��ѳ�����Ч�����������ٽ���ת�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        On Error GoTo errH
        
        '�����ھ���Ĳ��˵ļ��
        If mintActive = pt���� Or mintActive = pt���� Then
            If InStr(GetInsidePrivs(p����ҽ��վ), "����ҽ��ת��") > 0 Then
                '����Ƿ���δ���͵�ҽ��
                strSQL = "Select ID From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬=1 And Nvl(ִ������,0)<>0 And Rownum = 1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
                If Not rsTmp.EOF Then
                    MsgBox "�ò��˻���δ����ҽ����ֻ�н�����ҽ�����ͺ���ܽ���ת�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else    'ֻҪ�¹�ҽ��(���������ϵ�)��˵��������Ϊ�ѷ�����������ת������¹Һ�
                strSQL = "Select ID From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬ <> 4 And Rownum = 1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
                If Not rsTmp.EOF Then
                    MsgBox "�Ѿ��Ըò����¹�ҽ����������ת���ɾ��������ҽ�����ٽ��У��������¹Һš�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        If Not frmRegistPlan.ShowMe(Me, mstr�Һŵ�, lng����ID, str����, strҽ��, lngҽ��ID) Then mblnUnRefresh = False: Exit Sub
        
        'ִ��ת��
        strSQL = "Zl_���˹Һż�¼_ת��('" & mstr�Һŵ� & "',0," & lng����ID & ",'" & str���� & "','" & strҽ�� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        '���ﻼ��ת����Ϣ����
        Call ZLHIS_CIS_007(mclsMipModule, mlng����ID, Trim(txtEdit(txt����).Text), mPatiInfo.�����, mPatiInfo.�Һ�ID, mlng�������ID, , lng����ID, , lngҽ��ID, strҽ��, str����, UserInfo.����)
        
        Call zlShowQuence(mstr�Һŵ�)
        'ˢ�½���
        Call LoadPatients("11011")
    End With
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlShowQuence(ByVal strNO As String)
    '����:��ʾ�ŶӽкŶ��е��º�
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If Check�Ŷӽк� = False Then Exit Sub
    strSQL = "Select �ŶӺ��� From �ŶӽкŶ��� Where ҵ������=0 and ҵ��ID in (Select ID From ���˹Һż�¼ where NO=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    MsgBox "ע��:" & vbCrLf & "    �ò������½������ŶӴ���,�Ӻ�Ϊ:[ " & Nvl(rsTemp!�ŶӺ���) & " ]", vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub ExecuteTransferRefuse()
'���ܣ�ת��ܾ�
    Dim strSQL As String
        
    On Error GoTo errH
    
    With lvwIncept.SelectedItem
        If MsgBox("ȷʵҪ�ܾ���ת�ﲡ��""" & .SubItems(2) & """��", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        strSQL = "Zl_���˹Һż�¼_ת��('" & .Text & "',-1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End With
    'ˢ�½���
    Call LoadPatients("11011")
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferCancel(Optional ByVal blnMsg As Boolean = True)
'���ܣ�ȡ��ת��
    Dim strSQL As String
    Dim objLvw As ListView
    On Error GoTo errH
    If mintActive = pt���� Then
        Set objLvw = lvwPatiHZ
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    With objLvw.SelectedItem
        If blnMsg Then
            If MsgBox("ȷʵҪȡ������""" & .SubItems(2) & """��ת����", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
        strSQL = "Zl_���˹Һż�¼_ת��('" & .Text & "',Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End With
    
    'ˢ�½���
    Call LoadPatients("11011")
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferIncept()
'���ܣ�����ת��
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lvwIncept.SelectedItem Is Nothing Then Exit Sub
    
    With lvwIncept.SelectedItem
        If MsgBox(.SubItems(lvwIncept.ColumnHeaders.Count - 1) & vbCrLf & vbCrLf & "ȷ�Ͻ��ո�ת�ﲡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
 
        strSQL = "Zl_���˹Һż�¼_ת��('" & .Text & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        If HaveRIS Then
            If gobjRis.HISModPati(1, mlng����ID, mlng�Һ�ID) <> 1 Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
        ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
        Call mclsAdvices.zlRefresh(0, "", False)
        'ˢ�²���λ����
        If lvwPati(pt����).Visible Then
            Call LoadPatients("11011", pt����, .Text)
            lvwPati(pt����).SetFocus
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
    '����:���˽���
    '����:blnIsCard-�Ƿ���ˢ�����ý���ԤԼ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSQL As String, strNO As String
    Dim blnReserve As Boolean
    Dim datCurr As Date
   
    On Error GoTo errH
    If lvwPati(pt����).Visible And lvwReserve.Visible Then
        blnReserve = Me.ActiveControl Is lvwReserve
    Else
        blnReserve = lvwReserve.Visible
    End If
    datCurr = zlDatabase.Currentdate
    If blnReserve Then
        '��ԤԼ�ҺŲ��˽��н���
        If lvwReserve.SelectedItem Is Nothing Then Exit Sub
        
        '�����:57566
        If Check�������("����", Mid(lvwReserve.SelectedItem.Key, 2)) = False Then Exit Sub
        
        '����ҽ��վԤԼ����ʱ���ùҺŲ����Ľ��սӿڽ��п۷ѵĹ���
        If Val(zlDatabase.GetPara("����ҺŻ��۵�", glngSys, p����ҽ��վ, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, Mid(lvwReserve.SelectedItem.Key, 2), mstr��������, PatiIdentify.objIDKind.GetCurCard.�ӿ����, PatiIdentify.Text) Then Exit Sub
        Else
            With lvwReserve.SelectedItem
                strNO = Mid(.Key, 2)
                strSQL = "Zl_����ԤԼ�Һ�_����('" & strNO & "','" & mstr�������� & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End With
        End If
        
        'ˢ�²���λ����
        On Error GoTo 0
        If lvwPati(pt����).Visible Then
            Call LoadPatients("11001", pt����, strNO)
            lvwPati(pt����).SetFocus
        Else
            Call LoadPatients("11001")
        End If
    Else
        '�����:57566
        If Check�������("����", mstr�Һŵ�) = False Then Exit Sub
        '�������ҺŲ��˽��н���
        strSQL = "Select ִ���� From ���˹Һż�¼ Where ����ID+0=[1] And NO=[2] And Nvl(ִ��״̬,0)<>0 And ��¼����=1 And ��¼״̬=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            MsgBox "�ò�������" & IIf(IsNull(rsTmp!ִ����), "����ҽ��", "ҽ����" & rsTmp!ִ���� & " ") & "���", vbInformation, gstrSysName
            Call LoadPatients("100"): Exit Sub
        End If
        
        strSQL = "Select ִ���� From ���˹Һż�¼ Where ����ID+0=[1] And NO=[2] And Nvl(ִ��״̬,0)=0 And ��¼����=1 And ��¼״̬=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        If rsTmp.EOF Then
            MsgBox "�ò������˺ţ����ܽ��", vbInformation, gstrSysName
            Call LoadPatients("100"): Exit Sub
        End If
        
        strSQL = "zl_���˽���(" & mlng����ID & ",'" & mstr�Һŵ� & "',Null,'" & UserInfo.���� & "','" & mstr�������� & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        'ˢ�²���λ����
        On Error GoTo 0
        If lvwPati(pt����).Visible Then
            Call LoadPatients("110", pt����, mstr�Һŵ�)
            lvwPati(pt����).SetFocus
        Else
            Call LoadPatients("110")
        End If
    End If
    '���ﻼ�߽�����Ϣ����
    Call ZLHIS_CIS_009(mclsMipModule, mlng����ID, Trim(txtEdit(txt����).Text), mPatiInfo.�����, Val(ucPatiVitalSigns.value���), Val(ucPatiVitalSigns.value����), mlng�Һ�ID, IIf(optState(opt����).Value, 1, 0), IIf(lbl��.Visible, 1, 0), datCurr, mlng�������ID, , mstr��������, UserInfo.����)

    '���������Զ����ù���
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng����ID <> 0 And mPatiInfo.���� <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    Call CreatePlugInOK(p����ҽ��վ)
    '����������ҽӿ�
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicReceive(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID)
        Call zlPlugInErrH(err, "ClinicReceive")
        err.Clear: On Error GoTo errH
    End If
    
    '����֮���Զ�����ҽ���´�״̬
    If mlng�Զ����� = 1 Then
        If tbcSub.Selected.Tag <> "ҽ��" Then tbcSub.Item(0).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    ElseIf mlng�Զ����� = 2 Then
        If tbcSub.Selected.Tag <> "����" Then tbcSub.Item(1).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr�Һŵ�)
    End If
    '�����ŶӽкŶ���(����ˢ��)
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteCancel()
'���ܣ�ȡ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If BillExpend(mstr�Һŵ�) Then
        MsgBox "�ò��˹Һ��ѳ�����Ч��������������ȡ�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    On Error GoTo errH
    
    'ֻ��ȡ���Լ�����Ĳ���
    strSQL = "Select ִ���� From ���˹Һż�¼ Where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.�Һ�ID)
    If rsTmp!ִ���� <> UserInfo.���� Then
        MsgBox "ֻ��ȡ���Լ�����Ĳ��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ToDo:ȡ������ʱ�������ݵļ��
    'ҽ�����ݵļ��
    strSQL = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ҽ��״̬ IN(1,8) And ����ID+0=[1] And �Һŵ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    If Nvl(rsTmp!ҽ��, 0) > 0 Then
        MsgBox "�ò��������¿����ѷ��͵�ҽ��������ȡ�����" & vbCrLf & _
            "���ȷʵҪȡ��������Ƚ���Щҽ��ɾ�������ϡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSQL = "Zl_���˽���_Cancel(" & mlng����ID & ",'" & mstr�Һŵ� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    'ˢ�²���λ����
    If lvwPati(pt����).Visible Then
        Call LoadPatients("110", pt����, mstr�Һŵ�)
        lvwPati(pt����).SetFocus
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
'���ܣ���ɽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnTran As Boolean
    Dim str����IDs As String, str���IDs As String
    Dim lng�Һ�id As Long
    Dim objLvw As ListView
    
    On Error GoTo errH
    If mintActive = pt���� Then
        Set objLvw = lvwPatiHZ
    Else
        Set objLvw = lvwPati(pt����)
    End If
    
    If objLvw.SelectedItem Is Nothing Then Exit Sub
    '����б�ʱ�䲻ˢ�²����������
    strSQL = "select 1 from ���˹Һż�¼ where no=[1] and ִ����=[2] And ִ��״̬=2 And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�, mstr����ҽ��)
    If rsTmp.EOF Then
        MsgBox """" & objLvw.SelectedItem.SubItems(2) & """���ܱ�����ҽ��ǿ��������գ������ԡ�", vbInformation, gstrSysName
        Call LoadPatients
        Call ReshDataQueue
        Exit Sub
    End If
    'ToDo:��ɽ���ʱ�������ݵļ��
    
    If objLvw.SelectedItem.ListSubItems(5).Tag = "0" Then
        If MsgBox("��ǰ����""" & objLvw.SelectedItem.SubItems(2) & """�Ѿ�ת��Ƿ�Ҫȡ��ת�������ɽ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            Call ExecuteTransferCancel(False)
            Call ExecuteFinish
            Exit Sub
        End If
    End If
    '����Ƿ������Чҽ��
    strSQL = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬<>4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    If Nvl(rsTmp!ҽ��, 0) = 0 Then
        If MsgBox("δ��""" & objLvw.SelectedItem.SubItems(2) & """�´��κ���Ч��ҽ����ȷʵҪ��ɽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    '����Ƿ����δ���͵�ҽ��
    strSQL = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬=1 And Nvl(ִ������,0)<>0 And Nvl(Ƥ�Խ��,'��')<>'����'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    If Nvl(rsTmp!ҽ��, 0) > 0 Then
        MsgBox """" & objLvw.SelectedItem.SubItems(2) & """����δ���͵�ҽ����������ɽ��", vbInformation, gstrSysName
        Exit Sub
    End If
    '���δ��д�ļ���֤������
    strSQL = "Select ��ҳID,����ID,���ID From ������ϼ�¼ Where ȡ��ʱ�� is Null And ����ID=[1] And ��ҳID=(Select ID From ���˹Һż�¼ Where NO=[2] And ��¼����=1 And ��¼״̬=1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    Do While Not rsTmp.EOF
        If lng�Һ�id = 0 Then lng�Һ�id = rsTmp!��ҳID
        If Not IsNull(rsTmp!����id) Then str����IDs = str����IDs & "," & rsTmp!����id
        If Not IsNull(rsTmp!���id) Then str���IDs = str���IDs & "," & rsTmp!���id
        rsTmp.MoveNext
    Loop
    If str����IDs <> "" Or str���IDs <> "" Then
        If InStr(";" & GetPrivFunc(glngSys, p���ﲡ������) & ";", ";������д;") > 0 Then
            If Not CheckDiseaseFile(Me, mlng����ID, lng�Һ�id, mlng�������ID, Mid(str����IDs, 2), Mid(str���IDs, 2), , True) Then Exit Sub
        End If
    End If
    
    '��ȡ��Ҫ����Ϣ�������ӿڵ���:����߾��ﲡ�˱��ξ���Ϊ׼,�ұ߿��ܵ�ǰѡ�����ʷ����
    strSQL = "Select A.ID,A.����,B.������ From ���˹Һż�¼ A,����������Ϣ B Where A.����ID=B.����ID(+) And A.��¼����=1 And A.��¼״̬=1 And A.����=B.����(+) And A.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    
    'ִ�й���
    '-----------------------------------
    gcnOracle.BeginTrans: blnTran = True
    
    strSQL = "Zl_���˽������(" & mlng����ID & ",'" & mstr�Һŵ� & "','" & mstr�������� & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
         
    If Not gobjCommunity Is Nothing And Nvl(rsTmp!����, 0) <> 0 Then
        '��������������Ϣ�ύ
        If Not gobjCommunity.ClinicSubmit(glngSys, mlngModul, rsTmp!����, Nvl(rsTmp!������), mlng����ID, rsTmp!ID) Then
            gcnOracle.RollbackTrans: blnTran = False: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False

    '����������ҽӿ�
    Call CreatePlugInOK(p����ҽ��վ)
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicFinish(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID)
        Call zlPlugInErrH(err, "ClinicFinish")
        err.Clear: On Error GoTo errH
    End If
    
    'һ��ͨ�����ϴ�
    If Not mobjICCard Is Nothing Then
        strSQL = "Select 1 From һ��ͨĿ¼ Where ����=2 And Rownum=1"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            mobjICCard.UploadSwap mlng����ID, ""
        End If
    End If
    'ˢ��:����λ�������б�
    Call LoadPatients
    Call ReshDataQueue
    
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteRedo()
'�ָ�����
    Dim strSQL As String
    
    'ֻ����������ݱ��е�
    If BillExpend(mstr�Һŵ�) Then
        MsgBox "�ò��˹Һ��ѳ�����Ч�������������ٻָ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintActive = pt���� Then
        If zlDatabase.NOMoved("���˹Һż�¼", mstr�Һŵ�) Then
            MsgBox "�ùҺż�¼�Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '��ǰҽ����ɵĲ��˲ſ���ֱ�ӻָ�(������Ȩ�޿���ǿ������)
    With lvwPati(pt����).SelectedItem
        If .ListSubItems(4).Tag <> UserInfo.���� Then
            MsgBox "�ò��˲���������ɾ���ģ�����ֱ�ӻָ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    On Error GoTo errH
    strSQL = "zl_���˽������_Cancel(" & mlng����ID & ",'" & mstr�Һŵ� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    'ˢ�²���λ����
    If lvwPati(pt����).Visible Then
        Call LoadPatients("011001", pt����, mstr�Һŵ�)
        lvwPati(pt����).SetFocus
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
'���ܣ��������������֤
    Dim arrSQL As Variant, i As Long
    Dim colInfo As New Collection
    Dim int���� As Integer, str������ As String
    Dim str�������� As String
        
    If gobjCommunity Is Nothing Or mlng����ID = 0 Or mPatiInfo.�Һ�ID = 0 Or mPatiInfo.���� <> 0 Then Exit Sub
    
    If Not gobjCommunity.Identify(glngSys, p����ҽ��վ, int����, str������, colInfo, mPatiInfo.����ID, mPatiInfo.�Һ�ID) Then Exit Sub
    
    arrSQL = Array()
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����������Ϣ_Insert(" & mPatiInfo.����ID & "," & int���� & ",'" & str������ & "',1,Sysdate)"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    str�������� = GetColItem(colInfo, "��������")
    If IsDate(str��������) Then
        str�������� = "To_Date('" & Format(str��������, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
    Else
        str�������� = "Null"
    End If
    arrSQL(UBound(arrSQL)) = "Zl_���˹Һż�¼_������֤(" & mPatiInfo.����ID & "," & mPatiInfo.�Һ�ID & "," & int���� & "," & _
        "'" & GetColItem(colInfo, "����") & "','" & GetColItem(colInfo, "�Ա�") & "','" & GetColItem(colInfo, "����") & "'," & _
        str�������� & ",'" & GetColItem(colInfo, "�����ص�") & "','" & GetColItem(colInfo, "���֤��") & "'," & _
        "'" & GetColItem(colInfo, "����") & "','" & GetColItem(colInfo, "����") & "','" & GetColItem(colInfo, "����״��") & "'," & _
        "'" & GetColItem(colInfo, "ְҵ") & "','" & GetColItem(colInfo, "��ͥ��ַ") & "','" & GetColItem(colInfo, "��ͥ�绰") & "'," & _
        "'" & GetColItem(colInfo, "��ͥ��ַ�ʱ�") & "','" & GetColItem(colInfo, "������λ") & "','" & GetColItem(colInfo, "��λ�绰") & "'," & _
        "'" & GetColItem(colInfo, "��λ�ʱ�") & "','" & GetColItem(colInfo, "��ϵ������") & "','" & GetColItem(colInfo, "��ϵ�˹�ϵ") & "'," & _
        "'" & GetColItem(colInfo, "��ϵ�˵绰") & "','" & GetColItem(colInfo, "��ϵ�˵�ַ") & "','" & GetColItem(colInfo, "���ڵ�ַ") & "','" & GetColItem(colInfo, "���ڵ�ַ�ʱ�") & "')"
    
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
'���ܣ���������æ��״̬
    On Error GoTo DBError
    gcnOracle.Execute "Update �������� Set ȱʡ��־=" & IIf(blnBusy, 1, 0) & " Where ����='" & mstr�������� & "' And ȱʡ��־<>" & IIf(blnBusy, 1, 0)
    On Error GoTo 0
    
    Me.stbThis.Panels(3).Text = "����" + IIf(blnBusy, "æ", "��")
    Me.lblRoom.BackColor = IIf(blnBusy, COLOR_BUSY, COLOR_FREE)
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("����ˢ�¼��", glngSys, p����ҽ��վ, 180))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
    If mintRefresh = 0 Then
        timRefresh.Enabled = False
    Else
        timRefresh.Interval = 1000 '�̶�Ϊ1����
        timRefresh.Enabled = True
    End If
End Sub
Private Sub timRefresh_Timer()
    Static lngSecond As Long
    Static strPreTime1 As String
    Dim curTime As Date
    
    If mbln��Ϣ���� Then
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
        If mclsMipModule.IsConnect Then 'ʹ������Ϣƽ̨���µ�ˢ�²���
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
    lngSecond = lngSecond + 1 '����
    If lngSecond Mod mintRefresh = 0 Then
        lngSecond = 0
        Call LoadPatients("100111")
        Call ReshDataQueue
    End If
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String, Optional ByVal blnIsCard As Boolean _
                            , Optional ByVal lngPatiID As Long)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
'      strIDCard=����ֵʱ����ʾ�̶������֤�Ų���
'      blnIsCard=�Ƿ���ˢ�����ý���ԤԼ����
    Static blnReStart As Boolean
    Dim intIdx As PatiType, i As Long
    Dim objControl As CommandBarControl
    Dim objLvw As ListView
    Dim lngReserve As Long  '��ͷ�ҵ�ʱ������ԤԼ��
    Dim blnQueueFind As Boolean
    
    If mintActive = -1 Or mintActive = ptת�� Then
        PatiIdentify.Text = "": Exit Sub
    End If
    lngReserve = 1
    
    '��������ʽ���Һ��Զ�ˢ���֤�ļ���������ȡ��
    If strIDCard = "" And PatiIdentify.Text <> "" Then mstrIDCard = ""
    
    If Not blnNext And mstrFindType = "�Һŵ�" Then
        PatiIdentify.Text = GetFullNO(PatiIdentify.Text, 12)
    End If
    PatiIdentify.SetFocus
    If mintActive = pt���� Then
        Set objLvw = lvwPatiHZ
    ElseIf mintActive = ptԤԼ Then
        Set objLvw = lvwReserve
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    
    '��ʼ������
    If Not blnNext Or blnReStart Or objLvw.SelectedItem Is Nothing Then
        intIdx = pt���� - lngReserve: i = 1
    Else
        intIdx = mintActive
        '=3Ϊ����
        If intIdx = pt���� Then intIdx = pt���� - 2
        i = objLvw.SelectedItem.Index + 1
    End If
    
     '���Ҳ���
    If lngPatiID = 0 And Not mobjSquareCard Is Nothing And mstrFindType <> "���￨" And mstrFindType <> "��ʶ��" And mstrFindType <> "�Һŵ�" And mstrFindType <> "����" And mstrFindType <> "�������֤" Then
        If mstrFindType = "IC��" Then
            Call mobjSquareCard.zlGetPatiID("IC��", PatiIdentify.Text, , lngPatiID)
        Else
            Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����), PatiIdentify.Text, , lngPatiID)
        End If
    End If
    
    '���Ҳ���
    If Check�Ŷӽк� = True Then
        blnQueueFind = mobjQueue.FindQueue(IIf(PatiIdentify.objIDKind.GetCurCard.�ӿ���� > 0, _
                            PatiIdentify.objIDKind.GetCurCard.�ӿ����, _
                            IIf(PatiIdentify.objIDKind.GetCurCard.���� = "��ʶ��", "�����", PatiIdentify.objIDKind.GetCurCard.����)), _
                            PatiIdentify.Text)
    End If
    If blnQueueFind = False Then
        For intIdx = intIdx To lvwPati.UBound + 2
            If intIdx = lvwPati.UBound + 1 Then
                Set objLvw = lvwPatiHZ
            ElseIf intIdx = pt���� - lngReserve Or intIdx = lvwPati.UBound + 2 Then
                Set objLvw = lvwReserve
            Else
                Set objLvw = lvwPati(intIdx)
            End If
            For i = i To objLvw.ListItems.Count
                With objLvw.ListItems(i)
                    If strIDCard <> "" Then '���֤�Զ�ʶ��ǿ������
                        If UCase(.ListSubItems(6).Tag) = UCase(strIDCard) Then Exit For
                    Else
                        If Val(.Tag) = lngPatiID And lngPatiID <> 0 Then Exit For
                        Select Case mstrFindType
                            Case "���￨"
                                If .ListSubItems(1).Tag = PatiIdentify.Text Then Exit For
                            Case "��ʶ��"
                                If .SubItems(1) = PatiIdentify.Text Then Exit For '�����
                            Case "�Һŵ�"
                                If UCase(.Text) = UCase(PatiIdentify.Text) Then Exit For '���ݺ�
                            Case "����"
                                If .SubItems(2) Like "*" & PatiIdentify.Text & "*" Then Exit For
                            Case "�������֤"
                                If UCase(.ListSubItems(6).Tag) = UCase(PatiIdentify.Text) Then Exit For
                            Case "IC��"
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
            ElseIf intIdx = pt���� - lngReserve Or intIdx = lvwPati.UBound + 2 Then
                Set objLvw = lvwReserve
                If intIdx = pt���� - lngReserve Then intIdx = ptԤԼ
            Else
                Set objLvw = lvwPati(intIdx)
            End If
            mstrPrePati = objLvw.ListItems(i).Key
            objLvw.ListItems(i).Selected = True
            objLvw.SelectedItem.EnsureVisible
            
            mstrPrePati = ""
            If intIdx = lvwPati.UBound + 1 Then
                Call LvwItemClick(pt����, objLvw.SelectedItem)
            ElseIf intIdx <> lvwPati.UBound + 2 Then
                Call lvwPati_ItemClick(CInt(intIdx), objLvw.SelectedItem)
            ElseIf intIdx = ptԤԼ Then
                Call LvwItemClick(ptԤԼ, objLvw.SelectedItem)
            End If
            
            If Not objLvw.Visible Then
                For i = 1 To dkpMain.PanesCount
                    If dkpMain.Panes(i).Handle = objLvw.hwnd Then
                        dkpMain.Panes(i).Select
                    End If
                Next
            End If
            If objLvw.Visible Then objLvw.SetFocus
            
            '�ҵ����Զ����н���,ԤԼ�����Զ�����
            If (mbln�Զ����� And intIdx = pt����) Or intIdx = ptԤԼ Then
                cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                If intIdx = ptԤԼ Then
                    If mstrFindType = "��ʶ��" Or mstrFindType = "�Һŵ�" Or mstrFindType = "����" Or mstrFindType = "�������֤" Then Exit Sub
                    Call ExecuteReceive(blnIsCard)
                Else
                    Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then Call cbsMain_Update(objControl) '�״�ִ�У�û����ʾ�˵�ǰ���¼�û��ִ��
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
        Else
            blnReStart = True
            MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub ucPatiVitalSigns_Change(ByVal int��� As Integer)
    Call SetPermitEscape(False)
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsAller.ComboList = "..."
    vsAller.FocusRect = flexFocusSolid
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int�Ա� As Integer
    
    With vsAller
        Call RefreshPass
        If mblnUseTYT Then
            strSQL = gobjPass.inputAllergy()
            If strSQL <> "" Then
                Call SetAllerInput(Col, , strSQL)
                Call AllerEnterNextCell
            End If
        Else
            If cboEdit(cbo�Ա�).Text Like "*��*" Then
                int�Ա� = 1
            ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                int�Ա� = 2
            End If
            
            strSQL = _
                " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
                " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
                " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                " Union All" & _
                " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
                " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                " From ������ĿĿ¼ A,ҩƷ���� B" & _
                " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
                IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û��ҩƷ���ݿ���ѡ��", vbInformation, gstrSysName
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
                If MsgBox("ȷʵҪ����������ҩ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnDo = True
                End If
            Else
                blnDo = .Col <> .Cols - 1
            End If
            
            If blnDo Then
                If CStr(.Cell(flexcpData, 0, .Col)) <> "" Then  'ɾ����ǰ�����������
                    .ColWidth(.Col) = 0
                    '.ColHidden(.Col) = True    '��ʹ��hidden����Ϊ�������к����һ�в�����.col-1
                    .TextMatrix(0, .Col) = ""
                    
                    If .Col = .Cols - 1 Then
                        Call AllerAddCol
                    Else
                        .Col = .Cols - 1
                        .ShowCell 0, .Col
                    End If
                    
                    Call SetPermitEscape(False)
                Else   'ɾ��һ��ʱ�����ƺ�����
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
            '���ֱ�����뺺�ֵ�����
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
                .ComboList = "" 'ʹ��ť״̬��������״̬
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
    Dim int�Ա�  As Integer
    
    With vsAller
        If .EditText = "" Then
            .Cell(flexcpData, Row, Col) = ""
            If vsAller.Tag = "KeyPress" Then Call AllerEnterNextCell
        ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
            If vsAller.Tag = "KeyPress" Then Call AllerEnterNextCell
        Else
            strInput = UCase(.EditText)
            If cboEdit(cbo�Ա�).Text Like "*��*" Then
                int�Ա� = 1
            ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                int�Ա� = 2
            End If
            strSQL = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ," & _
                " B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
                " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
                " And (A.���� Like [1] Or A.���� Like [2] Or C.���� Like [2] Or C.���� Like [2])" & _
                IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[3])", "") & _
                Decode(gint����, 0, " And C.����=[4]", 1, " And C.����=[4]", "") & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " Order by A.����"
            
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ��", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                strInput & "%", gstrLike & strInput & "%", int�Ա�, gint���� + 1)
            If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                Cancel = True
            Else
                Call SetAllerInput(Col, rsTmp): .EditText = .Text
                If vsAller.Tag = "KeyPress" Or vsAller.Col = vsAller.Cols - 1 Then Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngCol As Long, Optional rsInput As ADODB.Recordset, Optional ByVal strTYTInput As String)
'���ܣ��������ҩ�������
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
                .TextMatrix(0, lngCol) = "" & rsInput!����
            Else    '��������¼��
                
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
'���ܣ�����һ������������
    With vsAller
        .Cols = .Cols + 1
        .ShowCell 0, .Cols - 1
        
        .ColWidth(.Cols - 1) = 1200
        .ColAlignment(.Cols - 1) = flexAlignLeftCenter
    End With
End Sub
Private Function Check�Ŷӽк�() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ʹ����ŶӽкŹ���
    '���أ��ŶӽкŹ������еĶ��Ϸ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-06 10:19:43
    '˵��������: Ȩ�޺Ϸ����;�������Ŷӽкŵ�;�����Ŷӽкųɹ�!
    '------------------------------------------------------------------------------------------------------------------------
    '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    If mty_Queue.byt�Ŷӽк�ģʽ = 0 Then GoTo GOEND:
    If Not (InStr(mty_Queue.strQueuePrivs, ";����;") > 0) Then GoTo GOEND:
    If mty_Queue.blnҽ���������� = False And mty_Queue.byt�Ŷӽк�ģʽ = 1 Then GoTo GOEND:
    
    err = 0: On Error GoTo GOEND:
    If mobjQueue Is Nothing Then
        Set mobjQueue = CreateObject("zlQueueManage.clsQueueManage")
        err = 0: On Error GoTo ErrHand:
        mobjQueue.zlInitVar gcnOracle, glngSys, 0, IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), mty_Queue.strQueuePrivs, "", False
        mobjQueue.zlSetToolIcon 24, True
        mobjQueue.IsShowFindTools = False
    End If
    Check�Ŷӽк� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
GOEND:
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    Set mobjQueue = Nothing

End Function
Private Sub ReshDataQueue()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ���Ŷӽк�����
    '���ƣ����˺�
    '���ڣ�2010-06-07 15:27:57
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim varQueue() As String, strTemp As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim str���� As String, strҽ�� As String, str���� As String
    Dim intType As Integer
    
    If mobjQueue Is Nothing Then Exit Sub
    If Check�Ŷӽк� = False Then Exit Sub
    '��ȡ��صĶ�������
    '���ﷶΧ��1=�ұ��˺ŵĲ���,2=�����Ҳ���,3=�����Ҳ���
    mint���ﷶΧ = Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2"))
    Dim strQueue() As String
    
    ReDim Preserve strQueue(1 To 1) As String
    str���� = IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID)
    strQueue(1) = str����
    strҽ�� = IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��)
    str���� = mstr��������
    intType = 1
    Select Case mint���ﷶΧ
    Case 1   '1=�ұ��˺ŵĲ���
        If Not mty_Queue.blnҽ���������� Then
           strҽ�� = UserInfo.����  '64696,������,2014-01-08,�õ�¼��Ա�����������ŶӽкŶ���
        End If
        If mlng�������ID = 0 Then strQueue(1) = ""
        intType = 3
    Case 2  '2=�����Ҳ���
        If Not mty_Queue.blnҽ���������� Then
           str���� = mstr��������
        End If
        If mlng�������ID = 0 Then strQueue(1) = ""
        intType = 2
    Case 3  '3=�����Ҳ���
    End Select
    
    '��Ҫ�Ŷ�û�н����Ĳ���
    strSQL = "" & _
    "   Select distinct  /*+ Rule*/  c.ҵ��ID From ���˹Һż�¼ A ,�ŶӽкŶ���  C" & _
    "   Where A.id=C.ҵ��ID and C.��������=[1]  and nvl(C.ҵ������,0)=0 and nvl(A.����ID,0) =0 And a.��¼����=1 And a.��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    With rsTemp
        strTemp = ""
        Do While Not .EOF
            strTemp = strTemp & "," & Val(Nvl(rsTemp!ҵ��id))
            .MoveNext
        Loop
        If strTemp <> "" Then strTemp = "0|" & Mid(strTemp, 2)
    End With
    Call mobjQueue.zlRefresh(strQueue, mty_Queue.strCurrQueueName, mty_Queue.lngcurr�Һ�ID, str����, strҽ��, strTemp, intType)
End Sub
 
Private Sub zlQueueStartus(intType As Integer, strNO As String, lng����ID As Long)
  '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ܲ�����,
    '��Σ�2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-����ȡ������
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-03 14:15:46
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte
    If Check�Ŷӽк� = False Then Exit Sub
    
    strSQL = "SELECT ID,ִ�в���ID,����,ִ���� From ���˹Һż�¼ where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    strQueueName = Nvl(rsTemp!ִ�в���ID)
    If Nvl(rsTemp!ִ����) <> "" Then
        strQueueName = strQueueName & ":" & Nvl(rsTemp!ִ����)
    ElseIf Nvl(rsTemp!����) <> "" Then
        strQueueName = strQueueName & ":" & Nvl(rsTemp!����)
    End If
    
    lngID = Val(Nvl(rsTemp!ID))
    Select Case intType
    Case 3   ' ���˲�����;
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 3
    Case 4, 6   '���˴���,'����ȡ������
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 0
    Case 5  '������ɾ���
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    End Select
End Sub

Private Function Set���˹Һ�״̬(ByVal lngState As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ò��˹Һ�״̬
    '��Σ�lngState : -1- ���˲�����
    '                         0-���˴���
    '���Σ�
    '���أ��Ƿ����óɹ������˲�����ʱ����ɾ�����۵��ݣ����ٴ����ô���ʱ�����ò��ɹ� ����False ,�����������True
    '���ƣ����˺�
    '���ڣ�2010-06-03 15:24:48
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str����NO As String
    
    If mstr�Һŵ� = "" Then Exit Function
    
    On Error GoTo errH
    
    If lngState = -1 Then
        '��鲡���Ƿ������Ч��ҽ��
        strSQL = "Select 1 From ����ҽ����¼ Where ����id = [1] And �Һŵ� = [2]  And ҽ��״̬ <> -1 And ҽ��״̬ <> 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            MsgBox "�ò��˴�����Чҽ��,��������Ϊ������!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If

    '��ȡ�ҺŻ��۵���Ϣ
    strSQL = "Select ժҪ From ������ü�¼ Where NO = [1] And ��¼���� = 4 And ��¼״̬ = 1 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    If Not rsTmp.EOF Then
        If rsTmp!ժҪ & "" <> "" And InStr(rsTmp!ժҪ & "", "����:") <> 0 Then
            '��ȡ�ҺŻ��۵���Ϣ,�жϹҺŻ��۵��Ƿ���ڣ������ڣ�����������״̬����Ϊ����
            str����NO = Mid(rsTmp!ժҪ & "", Len("����:") + 1)
            strSQL = "Select 1 From ������ü�¼ Where NO = [1] And Mod(��¼����,10) = 1 And ��¼״̬ = 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����NO)
            If rsTmp.EOF Then
                If lngState = 0 Then '����Ϊ����
                    MsgBox "�ùҺŵ��Ļ��۷��ò����ڣ����˺ź����¹Һ�!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
            Else
                If lngState = -1 Then '����Ϊ������
                    If MsgBox("�ò��˴��ڹҺŵ��Ļ��۷��ã�����Ϊ������ʱ��ɾ���ùҺŵ��Ļ��۷��ã�" & vbCrLf & "���Ҳ����ٻָ�Ϊ����,�Ƿ����?��", vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    
    gcnOracle.BeginTrans
        strSQL = "Zl_���˹Һż�¼_״̬ ('" & mstr�Һŵ� & "'," & lngState & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call zlQueueStartus(IIf(lngState = -1, 3, 4), mstr�Һŵ�, mlng����ID)
        'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    gcnOracle.CommitTrans
    MsgBox "�����ɹ�!", vbInformation, gstrSysName
    
    Set���˹Һ�״̬ = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ExecuteStopAndReuse(ByVal bln���� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ծ��ﲡ�˽�����ͣ������������
    '���:bln����-true:�����Ѿ�ͣ�õľ��ﲡ��
    '����:
    '����:
    '����:���˺�
    '����:2010-12-08 20:26:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bln��ͣ As Boolean
    Dim strNO As String, rsTemp As ADODB.Recordset
    Dim objLvw As ListView
    If Not bln���� Then
        Set objLvw = lvwPati(pt����)
    Else
        Set objLvw = lvwPatiHZ
    End If
    With objLvw
        If .SelectedItem Is Nothing Then Exit Sub
        bln��ͣ = .SelectedItem.SmallIcon = "��ͣ"
        If bln���� Then
            If bln��ͣ = False Then
                MsgBox "ע��:" & vbCrLf & "    �ò��˻�δ��ͣ����,���ܽ��лָ���ͣ����!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        Else
            If bln��ͣ Then
                MsgBox "ע��:" & vbCrLf & "    �ò��˻�������ͣ����,���ܽ�����ͣ����!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        strNO = .SelectedItem.Text
        strSQL = "Select ID From ���˹Һż�¼ where NO=[1] And ��¼����=1 And ��¼״̬=1"
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.EOF Then
            Exit Sub
        End If
    End With
    If Not bln���� Then
        'Zl_���˹Һż�¼_����
        strSQL = "Zl_���˹Һż�¼_����("
        '  Id_In         ���˹Һż�¼.ID%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!ID)) & ","
        '  ��ִ�п���_In ���˹Һż�¼.ִ�в���id%Type,
        strSQL = strSQL & "NULL,"
        '  ������_In     ���˹Һż�¼.����%Type,
        strSQL = strSQL & "NULL,"
        '  ��ҽ��_In     ���˹Һż�¼.ִ����%Type,
        strSQL = strSQL & "NULL,"
        '  �����_In Integer:=0
        strSQL = strSQL & "1)"
        '--�����_In :0-�������;1-���Ϊ��Ҫ����
    Else
        'Zl_���˹Һż�¼_ȡ������
        strSQL = "Zl_���˹Һż�¼_ȡ������("
        '  Id_In         ���˹Һż�¼.ID%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!ID)) & ","
        '  �����_In Integer:=0
        strSQL = strSQL & "1)"
    End If
    On Error GoTo errHandle
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    'ˢ��:����λ�������б�
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
    'MouseDown����GotFocusִ��
    If Not mblnMouseDown And Not lvwPatiHZ.SelectedItem Is Nothing Then
        Call lvwPatiHZ_ItemClick(lvwPatiHZ.SelectedItem)
    End If
End Sub

Private Sub lvwPatiHZ_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LvwItemClick(pt����, Item)
End Sub

Private Sub lvwPatiHZ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMouseDown = True
End Sub
Private Sub lvwPatiHZ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim objItem As ListItem
    mblnMouseDown = False
    If Button = 2 And InStr(mstrPrivs, "���˽���") > 0 Then
        Set objItem = lvwPatiHZ.HitTest(x, y)
        If Not objItem Is Nothing Then
            Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub Set������Ŀ��������()
    Dim lng����ID As Long
    
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    If mlng����ID = 0 Then
        lng����ID = mPatiInfo.����ID
    Else
        lng����ID = mlng����ID
    End If
    If lng����ID = 0 Then
        lng����ID = UserInfo.����ID
    End If
        
    Call gobjCISBase.CallSetClinicCharge(lng����ID, 1, Me, gcnOracle, glngSys, gstrDBUser, E�������, InStr(GetInsidePrivs(p����ҽ��վ), ";������Ŀ��������;") = 0)
End Sub
Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'����: ���н��������ͳһ����
'����: blnSetMainFont �Ƿ���������������(���������ӽ����л�)
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
        Case "ҽ��"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "����"
            Call mclsEPRs.SetFontSize(mbytSize)
                Case "�²���"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
    End Select
End Sub

Private Sub SetPatiInfoPosition()
'���ܣ����ò�����Ϣ�Ŀؼ�λ��
    Dim lngDistance1 As Long
    Dim lngDistance2 As Long
    
    lngDistance1 = 30
    lngDistance2 = 120
    Call SetCtrlPosOnLine(False, 0, lblTitle�ѱ�, lngDistance1, lblShow(lbl�ѱ�), lngDistance2, lblTitle����, lngDistance1, lblShow(lbl����), lngDistance2, lblTitle����, _
        lngDistance1, lblShow(lbl����), lngDistance2, lblTitleҽ����, lngDistance1, lblShow(lblҽ����), lngDistance2, lblTitle������, lngDistance1, lblShow(lbl������))
    
    lblDiag(0).Top = lblTitle�ѱ�.Top + lblTitle�ѱ�.Height + 90
    Call SetCtrlPosOnLine(False, 0, lblDiag(0), lngDistance1, lblDiag(1))
    picFind.Width = IIf(mbytSize = 1, 670, 475)
    lblFind.Width = picFind.Width
    Call SetCtrlPosOnLine(False, 0, picFind, 10, PatiIdentify)
        picFind.Height = PatiIdentify.Height
End Sub

Private Sub SetPicBasisFontSizeAndPosition()
'���ܣ����ò�����Ϣ������
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
    lbl��.FontName = "����"
    lbl��.FontSize = IIf(mbytSize = 0, 14, 18)
    
    lblRec.FontName = "����"
    lblRec.FontSize = IIf(mbytSize = 0, 14, 18)
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    
    '������ҿؼ��Ĵ�С����
    fraLine(3).Width = cboRegist.Width
    fraLine(3).Height = cboRegist.Height
    fraLine(3).Left = lblRegistInput.Left + lblRegistInput.Width + 20
    fraLine(3).Top = IIf(mbytSize = 0, 130, 160)
    lblRegistInput.Top = fraLine(3).Top + 30
    fraRegistInput.Width = fraLine(3).Left + fraLine(3).Width - 60
    fraRegistInput.Height = lblRegistInput.Top + lblRegistInput.Height + 110
    
    '�Ա����������б��Լ���ؿؼ��Ĵ�С����
    For i = 0 To cboEdit.UBound
        If i = 3 Then i = i + 2
        cboEdit(i).Width = Me.TextWidth("������")
        fraLine(i).Width = cboEdit(i).Width
        fraLine(i).Height = cboEdit(i).Height + cboEdit(i).Top - 30
        fraLine(i).Top = lblEdit(lbl����).Top + lblEdit(lbl����).Height - fraLine(i).Height - 30
    Next
    cmdAller.Height = Me.TextHeight("��") * IIf(mbytSize = 0, 2, 1.5)
    cmdAller.Width = Me.TextWidth(cmdAller.Caption & "��")
    Call SetCtrlPosOnLine(True, -1, fraRegistInput, 100, lblEdit(lbl��������), 60, lblEdit(cboְҵ), 60, lblEdit(lbl��λ), 60, lblEdit(txt��ͥ��ַ), 180, lblEdit(lbl���֤), 180, ucPatiVitalSigns)
    Call SetCtrlPosOnLine(False, 0, fraRegistInput, lngDistance2, lblEdit(txt����), lngDistance1, txtEdit(txt����), lngDistance2, fraLine(cbo�Ա�))
    Call SetCtrlPosOnLine(False, 0, lblEdit(lbl��������), lngDistance1, txt��������, lngDistance1 * 0.5, txt����ʱ��, lngDistance2, lblEdit(txt����), lngDistance1, txtEdit(txt����), lngDistance1, fraLine(cbo����))

    cboEdit(cboְҵ).Width = fraRegistInput.Width + fraRegistInput.Left - fraLine(cboְҵ).Left
    fraLine(cboְҵ).Width = cboEdit(cboְҵ).Width
    fraLine(cbo����ʱ��).Width = cboEdit(cbo����ʱ��).Width
    Call SetPatiPictureSize '������Ƭ��С
    Call SetCtrlPosOnLine(False, 0, lblEdit(cboְҵ), lngDistance1, fraLine(cboְҵ), lngDistance2, lblEdit(lbl����ʱ��), lngDistance1, txtEdit(txt����), lngDistance1 * 0.5, fraLine(cbo����ʱ��), lngDistance1, txt��������, lngDistance1, txt����ʱ��, lngDistance2, lblEdit(txt������ַ), lngDistance1, txtEdit(txt������ַ))
    Call SetCtrlPosOnLine(False, 0, lblEdit(lbl��λ), lngDistance1, txtEdit(txt��λ����), -1 * cmdEdit(cmd��λ����).Width, cmdEdit(cmd��λ����), lngDistance2, lblEdit(lbl��λ�绰), lngDistance1, txtEdit(txt��λ�绰), lngDistance2, optState(opt����), lngDistance1, optState(opt����), lngDistance2, lblEdit(lblȥ��), lngDistance1, fraLine(cboȥ��), lngDistance2, lbl��ƾ���)
    
    Call SetCtrlPosOnLine(False, 0, lblEdit(txt��ͥ��ַ), lngDistance1, txtEdit(txt��ͥ��ַ), -1 * cmdEdit(cmd��λ����).Width, cmdEdit(cmd��ͥ��ַ), lngDistance2, lblEdit(txt�໤��), lngDistance1, txtEdit(txt�໤��), lngDistance2, lblEdit(lbl��ͥ�绰), lngDistance1, txtEdit(txt��ͥ�绰), lngDistance2, lblEdit(lblժҪ), lngDistance1, txtEdit(txt����ժҪ))  'lblEdit(txt����ѹ), lngDistance1, txtEdit(txt����ѹ), lngDistance1, lblEdit(txt����ѹ), lngDistance1, txtEdit(txt����ѹ), lngDistance1, fraLine(cboѪѹ��λ))
    Call SetCtrlPosOnLine(False, 0, lblEdit(lbl���֤), lngDistance1, txtEdit(txt���֤��), lngDistance2, lblEdit(lbl�ֻ���), lngDistance1, txtEdit(txt�ֻ���), lngDistance2, lblEdit(lbl����), lngDistance1, cmdAller, -30, vsAller)

End Sub

Private Sub SetPatiPictureSize()
    picPatient.Left = fraLine(cbo�Ա�).Left + fraLine(cbo�Ա�).Width
    If picPatient.Left < fraLine(cbo����).Left + fraLine(cbo����).Width Then
        picPatient.Left = fraLine(cbo����).Left + fraLine(cbo����).Width
    End If
    picPatient.Left = picPatient.Left + 100
    
    '������Ƭ��С
    picPatient.Height = lblEdit(lbl��������).Top + lblEdit(lbl��������).Height - picPatient.Top
    picPatient.Width = picPatient.Height * 1.25
    imgPatient.Height = picPatient.Height - 75
    imgPatient.Width = picPatient.Width - 75
End Sub

Private Sub SetPicOutDocFontSizeAndPosition()
'���ܣ����ò��˲�ʷ��Ϣ������弰�ؼ���С
    Dim i As Long
    
    For i = 0 To lblDoc.UBound
        lblDoc(i).Width = Me.TextWidth("��")
        lblDoc(i).Height = Me.TextHeight("��") * 3
    Next
    lbl��ʾ.Left = rtfEdit(txt����).Left
    lbl��������.Left = rtfEdit(txt����).Left
    cmdSign.Left = rtfEdit(txt����).Left + rtfEdit(txt����).Width - cmdSign.Width
    
    '����Fontsizeʱ�ᴥ��change�¼�
    mblnSizeTmp = True
    Call SetRTFEditFontSize
    mblnSizeTmp = False
End Sub

Private Sub SetRTFEditFontSize()
'���ܣ����ò��˲�ʷ��Ϣ�����������
    Dim i As Long
    
    For i = 0 To rtfEdit.UBound
        Call SetPublicRTFFont(rtfEdit(i), IIf(mbytSize = 0, 9, 12))
    Next
End Sub

Private Function Check�������(str���� As String, strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˽������
    '���:str���� -��ǰ���� strNo - �Һŵ��ݺ�
    '����:
    '����:
    '����:����
    '����:2013-1-17 20:26:59
    '�����:57566
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHanl:
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim rsԤԼʱ�� As Recordset
    Dim strMsg As String
    
    If mlng������� = 0 Then Check������� = True: Exit Function
    
    strSQL = "" & _
    "   Select  Nvl(A.ԤԼʱ��,nvl(����ʱ��,sysdate)) - " & mlng��ǰ����ʱ�� & "/24/60 as �Һ�ʱ��  " & _
    "   From ���˹Һż�¼ A " & _
    "   Where No=[1] And Nvl(A.ԤԼʱ��,nvl(����ʱ��,sysdate))- " & mlng��ǰ����ʱ�� & "*1/24/60>sysdate"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Check������� = True: Exit Function
    strMsg = "�ò�����Ҫ��" & Format(rsTemp!�Һ�ʱ��, "yyyy-mm-dd HH:MM:SS") & "����������" & str����
    If mlng������� = 2 Then
        Check������� = (MsgBox(strMsg & ",��ȷ��Ҫ����" & str���� & "��", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes)
    Else
        MsgBox strMsg & ",������" & str����, vbInformation, gstrSysName
    End If
    Exit Function
ErrHanl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'���ܣ���������ҽ��վ���յ�����Ϣ
    Dim objXML As zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim bytˢ�� As Byte  'ˢ�·�ʽ��1-�����б�2��ת���б�
    Dim blnˢ�� As Boolean
    
    On Error GoTo errH
    
    Set objXML = New zl9ComLib.clsXML
    Call objXML.OpenXMLDocument(strMsgContent)
    Select Case strMsgItemIdentity '��ȡ�Һż�¼id
        Case "ZLHIS_REGIST_001", "ZLHIS_REGIST_002" '���ﻼ�߹Һţ��������֪ͨ����ȡһ����ˢ��һ�η�ʽ������ǵ�һ����Ϣ������ˢ�¡�
            bytˢ�� = 1
            Call objXML.GetSingleNodeValue("register_id", strTmp)
        Case "ZLHIS_CIS_007" '���ﻼ��ת���ʱˢ�£���Ϣ������ʱ���ˢ�£�ֻˢ��ת���б�
            bytˢ�� = 2
            Call objXML.GetSingleNodeValue("clinic_id", strTmp)
    End Select
    
    If strTmp = "" Then Exit Sub
    
    strSQL = "Select ִ����,����,ִ�в���id,ת��ҽ��,ת������,ת�����id From ���˹Һż�¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strTmp))
    
    If bytˢ�� = 1 Then
        If mint���ﷶΧ = 1 And rsTmp!ִ���� & "" = UserInfo.���� And (Not mblnҪ����� Or mblnҪ����� And rsTmp!���� & "" <> "") Then
            blnˢ�� = True
        Else
            If (mint���ﷶΧ = 2 And rsTmp!���� & "" = mstr�������� Or mint���ﷶΧ = 3 And (Not mblnҪ����� Or mblnҪ����� And rsTmp!���� & "" <> "")) And _
                Val(rsTmp!ִ�в���ID & "") = IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID) And _
                (rsTmp!ִ���� & "" = "" Or rsTmp!ִ���� & "" = UserInfo.����) Then
                
                blnˢ�� = True
            End If
        End If
        
        If blnˢ�� Then
            mblnMsgOk = True
            If Not mblnFirstMsg Then     '�ǵ�һ����Ϣ
                mblnFirstMsg = True
                Call RefeshByMsg
            End If
        End If
    ElseIf bytˢ�� = 2 Then
        If mint���ﷶΧ = 1 And rsTmp!ת��ҽ�� & "" = UserInfo.���� Then
            blnˢ�� = True
        Else
            If (mint���ﷶΧ = 2 And rsTmp!ת������ & "" = mstr�������� Or mint���ﷶΧ = 3) And _
                Val(rsTmp!ת�����ID & "") = IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID) And _
                UserInfo.���� <> IIf("" = rsTmp!ִ���� & "", "��", rsTmp!ִ����) And _
                (rsTmp!ת��ҽ�� & "" = "" Or rsTmp!ת��ҽ�� & "" = UserInfo.����) Then
                
                blnˢ�� = True
            End If
        End If
        
        If blnˢ�� Then
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
'���ܣ�������Ϣƽ̨��ʹ�õ�ˢ�·�ʽ
    Dim strTmp As String
    
    If Not mblnMsgOk Then Exit Sub
    '���ԤԼ�б�ɼ���һ��ˢ��
    strTmp = "1000" & IIf(lvwReserve.Visible, 1, 0)
    Call LoadPatients(strTmp)
    Call ReshDataQueue
    mblnMsgOk = False
End Sub

Private Sub Hide���￨����()
'���ܣ�������ĸ��������б��е�   ���￨��  ��  ����Ϊ����
    Dim lngIndex As Long
    Dim strTmp As String
        
    strTmp = "���￨��"
    
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
'���ܣ����� ListView �б��е�ָ���е�������ֵ
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
        Set objCol = .Columns.Add(c_ͼ��, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_����ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_No, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_�����, "�����", 62, True)
        Set objCol = .Columns.Add(C_����ʱ��, "����ʱ��", 60, True)
        Set objCol = .Columns.Add(C_״̬, "״̬", 150, True)
         
        Set objCol = .Columns.Add(C_��Ϣ, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ҵ��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_�Һ�ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ID, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            If objCol.Index <> C_��� Or objCol.Index <> C_���� Then objCol.Sortable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û����������..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '���� ����
        .SortOrder.Add .Columns(C_���)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns(C_����)
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
    If Mid(mstrNotifyAdvice, mΣ��ֵ, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(mstrNotifyAdvice, m��Ⱦ��, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_032,ZLHIS_CIS_033"
    If Mid(mstrNotifyAdvice, m�������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_RECIPEAUDIT_001"
    If Mid(mstrNotifyAdvice, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_001"   '������Ѫ�����̲��д���Ϣ�Ͳ���
    If Mid(mstrNotifyAdvice, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_004"   '������Ѫ�����̲��д���Ϣ�Ͳ���
    If Mid(mstrNotifyAdvice, m��Ѫ��Ӧ, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_006"  '����Ѫ����д���Ϣ�Ͳ���
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
       
    strSQL = "Select b.id,a.����id,a.NO,a.id as �Һ�ID,a.�����,a.����,a.ִ��ʱ�� as ����ʱ��,b.��Ϣ����,b.���ͱ���, b.ҵ���ʶ, b.���ȳ̶�, b.�Ǽ�ʱ��,a.����,b.������Դ" & _
        " From ҵ����Ϣ�嵥 B, ���˹Һż�¼ A" & _
        " Where b.����id=a.Id And a.ִ����||''=[1]  And b.�Ǽ�ʱ��>=Trunc(Sysdate-" & (mintNotifyDay - 1) & ")" & _
        " And Nvl(b.�Ƿ�����,0)=0 And instr(','||[2]||',',','||b.���ͱ���||',')>0 AND substr(b.���ѳ���,1,1)='1' " & _
        " Order By b.���ȳ̶� Desc, b.�Ǽ�ʱ�� Desc"
    
    Screen.MousePointer = 11

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr����ҽ��, strTmp)
    
    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!���ͱ���
        Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
            strTag = strTag & "<TB>" & rsTmp!���ͱ��� & "," & rsTmp!ID
            blnDo = True
        Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!ҵ���ʶ & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!ҵ���ʶ
                blnDo = True
            End If
        Case "ZLHIS_BLOOD_006"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!���ͱ��� & ":" & rsTmp!����ID & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!���ͱ��� & ":" & rsTmp!����ID
                blnDo = True
            End If
        Case Else
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ��� & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ���
                blnDo = True
            End If
        End Select
        
        If blnDo Then
            Call AddReportRow(rsTmp!����ID & "," & rsTmp!�Һ�ID, rsTmp!����ID, rsTmp!NO, Nvl(rsTmp!����), Nvl(rsTmp!�����), Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm"), _
                 Nvl(rsTmp!��Ϣ����), rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), Nvl(rsTmp!ҵ���ʶ), rsTmp!������Դ & "", _
                 Nvl(rsTmp!����, 0), rsTmp!�Һ�ID, rsTmp!ID)
            blnDo = False
        End If
        rsTmp.MoveNext
    Next
    rptNotify.Populate 'ȱʡ��ѡ���κ���
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln��Ϣ���� Then
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
'���ܣ�����Ϣ�����б�������һ��
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objItemIcon As ReportRecordItem
    Dim strRowID As String '�����б��е�Ψһ��ʶ��"����id,��ҳid,��Ϣ����"
    Dim strNO As String
    Dim strҵ�� As String
    Dim str������Դ As String
    Dim int���ȼ� As Integer
    Dim int���� As Integer
    Dim Index As Integer
    
    On Error GoTo errH
     
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tagֵ ����ID,�Һ�ID
    Set objItem = objRecord.AddItem(""): objItem.Icon = 6
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  'NO
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '����
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '�����
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '����ʱ��
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '״̬������
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '��Ϣ���
    objRecord.AddItem strNO: Index = Index + 1
    
    int���ȼ� = Val(arrInput(Index))                     '���ȼ�
    objRecord.AddItem int���ȼ�: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '�Ǽ�����
    
    strҵ�� = arrInput(Index): Index = Index + 1              'ҵ���ʶ
    str������Դ = arrInput(Index): Index = Index + 1          '������Դ
    
    int���� = arrInput(Index): Index = Index + 1
    objRecord.AddItem strҵ��
    
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1   '�Һ�ID
    objRecord.AddItem Val(arrInput(Index)) '��ϢID��ҵ����Ϣ�嵥.ID
    
    If int���ȼ� > 1 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            If int���ȼ� = 3 Then
                objRecord.Item(Index).ForeColor = &HC0&
            End If
            objRecord.Item(Index).Bold = True
        Next
    End If
    '���ղ����ú�ɫ��ʾ
    If int���� > 0 And int���ȼ� <> 3 Then
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
'���ܣ������յ�����Ϣ���������б���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "select a.NO,a.����,a.ִ����,a.�����,a.ִ��ʱ��,a.���� from ���˹Һż�¼ a where a.id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!����ID & ""))

    If mstr����ҽ�� = rsTmp!ִ���� & "" Then
        '�ж��б��Ƿ��Ѿ���������Ϣ��
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_��Ϣ).Value = rsMsg!���ͱ��� And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!����ID & "," & rsMsg!����ID) Then
                    Exit Sub
                End If
            End If
        Next
        
        Call AddReportRow(rsMsg!����ID & "," & rsMsg!����ID, rsMsg!����ID, rsMsg!NO, rsTmp!����, Nvl(rsTmp!�����), Format(rsTmp!ִ��ʱ�� & "", "yyyy-MM-dd HH:mm"), Nvl(rsMsg!��Ϣ����), _
             rsMsg!���ͱ��� & "", rsMsg!���ȳ̶� & "", Format(rsMsg!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!ҵ���ʶ & "", rsMsg!������Դ & "", Nvl(rsTmp!����, 0), rsMsg!����ID, 0)
        
        rptNotify.Populate
         
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'���ܣ��Զ�����ҽ��У�ԡ�ȷ��ֹͣ��ִ�н���
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng����ID As Long
    Dim lngҽ��ID As Long, lng�Һ�id As Long, lng��ϢID As Long
    Dim strҵ�� As String, blnOk As Boolean
    Dim blnFinded As Boolean
    Dim strTmp As String
    Dim strNO As String
    Dim str�Һŵ� As String
    Dim str��Ϣ���� As String
    Dim i As Long
    Dim strPatis As String
    Dim blnOnePati As Boolean
    Dim blnTmp As Boolean
    
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_��Ϣ).Value
                strҵ�� = .Item(C_ҵ��).Value
                str�Һŵ� = .Item(C_No).Value
                str��Ϣ���� = .Item(C_״̬).Value
                lng����ID = Val(.Item(C_����ID).Value)
                lng�Һ�id = Val(.Item(C_�Һ�ID).Value)
                lng��ϢID = Val(.Item(C_ID).Value)
                lngIndex = .Index
            End With
    
            blnTmp = True
            
            If str�Һŵ� <> mstr�Һŵ� Then blnTmp = LocatePati(str�Һŵ�)
            
            If strNO = "ZLHIS_RECIPEAUDIT_001" Then
                '�Ƚ���Ƭ�л���ҽ����Ƭ������Ҳ˵�
                Call LocatedCard("ҽ��")
                cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                If str��Ϣ���� = "�������ϸ�" Then
                    '������Ϣ���ʹ���
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_Send, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                Else
                    'ҽ���༭����
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_CIS_032" Then
                Call mclsDisease.ShowDisRegist(Me, 1, Val(strҵ��), lng����ID, 0, str�Һŵ�)
            End If
            
            If strNO = "ZLHIS_BLOOD_006" Then
                If gobjPublicBlood Is Nothing And gblnѪ��ϵͳ Then InitObjBlood
                blnOk = gobjPublicBlood.zlIsBloodMessageDone(2, lng����ID, lng�Һ�id, 1, mlng����ID)
                If blnOk Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                Else
                    If FuncTraReaction(Val(strҵ��), mlngModul, False, IIf(InStr(1, strҵ��, ":") > 0, Val(Split(strҵ��, ":")(1)), 0)) Then
                        If gobjPublicBlood.zlIsBloodMessageDone(2, lng����ID, lng�Һ�id, 1, mlng����ID) Then
                            Call rptNotify.Records.RemoveAt(lngIndex)
                        End If
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_CIS_033" Then
            '��Ⱦ�����淴�޸���Ϣ�Ķ�
                blnOk = ReadMsgCIS033(lng����ID, lng�Һ�id, strҵ��, lng��ϢID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            
            If strNO <> "ZLHIS_CIS_033" And strNO <> "ZLHIS_BLOOD_006" Then
                blnOk = ReadMsg(lng����ID, lng�Һ�id, strNO, strҵ��, lng��ϢID, str�Һŵ�)
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
    Dim str�Һŵ� As String
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    str�Һŵ� = rptNotify.SelectedRows(0).Record.Item(C_No).Value
 
    If str�Һŵ� <> mstr�Һŵ� Then Call LocatePati(str�Һŵ�)
    
End Sub

Private Function ReadMsg(ByVal lng����ID As Long, ByVal lng�Һ�id As Long, ByVal strNO As String, ByVal strҵ�� As String, ByVal lng��ϢID As Long, ByVal str�Һŵ� As String) As Boolean
'���ܣ��Ķ���Ϣ
'˵������Ϣ�Ķ���ʽĿǰ��3�֣�����Ϣ�������Ķ�����ϢID�Ķ�����ҵ���ʶ�Ķ�
    Dim strSQL As String
    Dim lng����ID As Long
    Dim strҽ��ID As String
    Dim blnDo As Boolean
    Dim lngΣ��ֵID As Long  '���δ����Σ��ֵ��¼ID
    Dim strSQLReadMsg As String
    Dim blnHisΣ��ֵ As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objControl As CommandBarControl
    
    If mlng�������ID = 0 Then
        lng����ID = UserInfo.����ID
    Else
        lng����ID = mlng�������ID
    End If
    blnDo = True
    
    On Error GoTo errH
    
    strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng�Һ�id & ",'" & strNO & "',1,'" & UserInfo.���� & "'," & lng����ID
    Select Case strNO
    Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
        strSQL = strSQL & ",null,null,'" & strҵ�� & "'"
    Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
        strSQL = strSQL & ",null," & lng��ϢID
    End Select
    strSQL = strSQL & ")"
    
    strSQLReadMsg = strSQL
    
    If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
        If mblnΣ��ֵ Then
            'Σ��ֵ��Ϣ��ش���
            Call mobjKernel.ShowDealCritical(Me, lng����ID, 0, str�Һŵ�, lngΣ��ֵID)
            
            If lngΣ��ֵID <> 0 Then
                '����Ϣ����Ϊ����
                Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
                '�����LISΣ��ֵ����LIS�ӿ�
                If strNO = "ZLHIS_LIS_003" Then
                    Call InitObjLis(p����ҽ��վ)
                    If Not gobjLIS Is Nothing Then
                        strSQL = "select a.�걾id,a.�������,a.ȷ���� from ����Σ��ֵ��¼ a where a.id=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngΣ��ֵID)
                        If Not rsTmp.EOF Then
                            Call gobjLIS.WriteNotifyToLis(Val(rsTmp!�걾ID & ""), rsTmp!ȷ���� & "", rsTmp!������� & "")
                        End If
                    End If
                End If
            End If
            Call SetCriticalAdvice(lngΣ��ֵID)
            blnHisΣ��ֵ = True
        End If
    End If
    
    If Not blnHisΣ��ֵ Then
        If strNO = "ZLHIS_LIS_003" Then
            If strҵ�� <> "" Then
                strҽ��ID = strҵ��
                Call InitObjLis(p����ҽ��վ)
                If Not gobjLIS Is Nothing Then
                    blnDo = gobjLIS.GetReadNotify(Me, strҽ��ID, UserInfo.����)
                End If
            End If
        End If
        If strNO = "ZLHIS_BLOOD_004" Then
            '��Ѫ�����Ϣ���Ķ�״̬������Ѫ�ⲿ���ڲ����ٴ�����ִ���Ķ���Ϣ����
            strSQL = "select 1 from ����ҽ����¼ a where a.�Һŵ�=[1] and a.ҽ��״̬=1 and a.�������='K' and a.��鷽��='1' and a.���״̬=1 and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�Һŵ�)
            If Not rsTmp.EOF Then
                '��������ݣ��򵯳�ҽ���޸Ľ��棬�������в�ִ����Ϣ�Ķ�SQL���
                '�Ƚ���Ƭ�л���ҽ����Ƭ������Ҳ˵�
                Call LocatedCard("ҽ��")
                cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                'ҽ���༭����
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
'���ܣ�ͨ���Һŵ���λ����ǰ���Լ����б��ھ����б�ͻ����б����ҡ�
    Dim blnTmp As Boolean
    Dim objLvw As ListView
    Dim i As Long
    Dim objItem As MSComctlLib.ListItem
    Dim lngIndex As Long
    
    Set objLvw = lvwPati(pt����)
    lngIndex = pt����
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
        lngIndex = pt����
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

Private Sub mclsDisease_PatiTransfer(ByVal lng����ID As Long, ByVal str�Һ�No As String)
'���ܣ���Ⱦ�����Խ��津���¼�ת�
    Call ExecuteTransferSend
End Sub

Private Function GetOne���Խ��() As Long
'���ܣ���ȡһ��ָ�������Խ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    If mlng����ID = 0 Then
        MsgBox "��ѡ��һ�����ˡ�", vbInformation, gstrSysName
        Exit Function
    End If

   strSQL = "Select A.ID, '����' As ��Դ,a.��¼״̬, a.����id,  b.����,  b.�Ա�,  b.����, e.���� As ����, " & vbNewLine & _
        "b.����� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, f.���� As �Ǽǿ���, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, " & vbNewLine & _
        "a.����ʱ��, a.�������˵�� " & vbNewLine & _
        "From �������Լ�¼ A, ���˹Һż�¼ B,  ���ű� E, ���ű� F " & vbNewLine & _
        "Where  a.����id = b.����id And a.�Һŵ� = b.No  And " & vbNewLine & _
        "a.�Ǽǿ���ID = f.Id(+) And b.ִ�в���id = e.Id(+) And a.�Һŵ�=[1] order by a.�ͼ�ʱ�� desc"

   Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)

    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò��������Խ����¼��", vbInformation, gstrSysName
        Exit Function
    ElseIf rsTmp.RecordCount = 1 Then
        GetOne���Խ�� = Val(rsTmp!ID & "")
        Exit Function
    End If
    
    GetOne���Խ�� = mclsDisease.ShowPatiDis(rsTmp, Me)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadMsgCIS033(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str��ʶ As String, ByVal lng��ϢID As Long) As Boolean
'���ܣ���Ⱦ�����淴�޸���Ϣ�Ķ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lng�ļ�ID As Long
    Dim lng����ID As Long
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
 
    lng�ļ�ID = Val(Split(str��ʶ, ",")(0))
    
    strSQL = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        If mlng�������ID = 0 Then
            lng����ID = UserInfo.����ID
        Else
            lng����ID = mlng�������ID
        End If
        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.���� & "'," & lng����ID & ",null," & lng��ϢID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    
    '�������޸ı���
    Call mclsDisDoc.ModifyDiseaseDoc(Me, lng�ļ�ID, mlng����ID, mlng�Һ�ID, 1, mlng����ID)
    
    strSQL = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        If mlng�������ID = 0 Then
            lng����ID = UserInfo.����ID
        Else
            lng����ID = mlng�������ID
        End If
        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.���� & "'," & lng����ID & ",null," & lng��ϢID & ")"
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
'���ܣ���λ��ָ����ҳǩ��Ƭ���ڲ�ҳǩ
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

Private Sub SetCriticalAdvice(ByVal lng��¼ID As Long)
'���ܣ�ȷ����Σ��ֵ�󵯳�ҽ���´���棬�ղŵ�ǰ�����ҽ���뱾�εļ�¼������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim objControl As Object
    
    On Error GoTo errH
    If lng��¼ID = 0 Then Exit Sub
    strSQL = "select 1 from ����Σ��ֵ��¼ a where a.id=[1] and a.�Ƿ�Σ��ֵ=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    
    If Not rsTmp.EOF Then
        '�����´�ҽ���Ĵ���
        If tbcSub.Tag <> "ҽ��" Then
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    If tbcSub.Item(i).Tag = "ҽ��" Then
                        tbcSub.Item(i).Selected = True
                        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                        Exit For
                    End If
                End If
            Next
        End If
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then
                objControl.Parameter = lng��¼ID
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
'���ܣ�Σ��ֵ��ش���
    Dim lngΣ��ֵID As Long  '���δ����Σ��ֵ��¼ID
    
    Call mobjKernel.ShowDealCritical(Me, mlng����ID, 0, mstr�Һŵ�, lngΣ��ֵID)
    
    Call SetCriticalAdvice(lngΣ��ֵID)
End Sub
