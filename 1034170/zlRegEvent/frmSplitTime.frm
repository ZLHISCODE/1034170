VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplitTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ʱ�������"
   ClientHeight    =   3810
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8220
   Icon            =   "frmSplitTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdChange 
      Caption         =   "תΪ�Զ���(C)"
      Height          =   350
      Left            =   1890
      TabIndex        =   17
      Top             =   3375
      Width           =   1380
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   720
      TabIndex        =   18
      Top             =   3375
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6030
      TabIndex        =   16
      Top             =   3375
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4875
      TabIndex        =   15
      Top             =   3375
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   90
      Top             =   930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picStd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   0
      ScaleHeight     =   3300
      ScaleWidth      =   8220
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   15
      Width           =   8220
      Begin VB.PictureBox pic��ǰ��ɫ 
         BackColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   3405
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   43
         Top             =   2985
         Width           =   270
      End
      Begin MSMask.MaskEdBox txt��ҹS 
         Height          =   270
         Left            =   2505
         TabIndex        =   6
         Top             =   2565
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt��ҹE 
         Height          =   270
         Left            =   3840
         TabIndex        =   7
         Top             =   2565
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtǰҹS 
         Height          =   270
         Left            =   2505
         TabIndex        =   4
         Top             =   2115
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtǰҹE 
         Height          =   270
         Left            =   3840
         TabIndex        =   5
         Top             =   2115
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt����S 
         Height          =   270
         Left            =   2505
         TabIndex        =   0
         Top             =   1215
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt����E 
         Height          =   270
         Left            =   3840
         TabIndex        =   1
         Top             =   1215
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt����S 
         Height          =   270
         Left            =   2505
         TabIndex        =   2
         Top             =   1680
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt����E 
         Height          =   270
         Left            =   3840
         TabIndex        =   3
         Top             =   1680
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt��ҹȱʡ 
         Height          =   270
         Left            =   5265
         TabIndex        =   34
         Top             =   2565
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtǰҹȱʡ 
         Height          =   270
         Left            =   5265
         TabIndex        =   35
         Top             =   2115
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt����ȱʡ 
         Height          =   270
         Left            =   5265
         TabIndex        =   36
         Top             =   1680
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt������ǰ 
         Height          =   270
         Left            =   6705
         TabIndex        =   39
         Top             =   1215
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt������ǰ 
         Height          =   270
         Left            =   6705
         TabIndex        =   40
         Top             =   1680
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtǰҹ��ǰ 
         Height          =   270
         Left            =   6705
         TabIndex        =   41
         Top             =   2115
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt��ҹ��ǰ 
         Height          =   270
         Left            =   6705
         TabIndex        =   42
         Top             =   2565
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt����ȱʡ 
         Height          =   270
         Left            =   5265
         TabIndex        =   55
         Top             =   1215
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡʱ��"
         Height          =   180
         Left            =   5445
         TabIndex        =   56
         Top             =   855
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "�ź�ʱ��"
         Height          =   180
         Left            =   6885
         TabIndex        =   54
         Top             =   855
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   4
         Left            =   3675
         TabIndex        =   53
         Top             =   855
         Width           =   90
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   4020
         TabIndex        =   52
         Top             =   855
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   2685
         TabIndex        =   51
         Top             =   855
         Width           =   720
      End
      Begin VB.Line Line6 
         X1              =   6540
         X2              =   6540
         Y1              =   795
         Y2              =   2910
      End
      Begin VB.Line Line5 
         X1              =   5055
         X2              =   5055
         Y1              =   795
         Y2              =   2910
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "ʱ���"
         Height          =   180
         Left            =   1275
         TabIndex        =   50
         Top             =   855
         Width           =   540
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   795
         X2              =   7980
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label9 
         Caption         =   $"frmSplitTime.frx":000C
         Height          =   585
         Left            =   780
         TabIndex        =   49
         Top             =   135
         Width           =   5670
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ�ź�ʱ��ҺŰ�����ʾ��ɫ"
         Height          =   180
         Left            =   780
         TabIndex        =   44
         ToolTipText     =   "������ǰ�ź�ʱ���ڹҺ�ʱ�����б�����ʾ����ɫ������ǰʱ�䴦����ǰ�ź�ʱ���뿪ʼʱ��֮��ʱ���ô���ɫ��ʾ�ú��룬��������"
         Top             =   3030
         Width           =   2520
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmSplitTime.frx":00CA
         Top             =   90
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   1275
         X2              =   1275
         Y1              =   1125
         Y2              =   2910
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ҹ"
         Height          =   180
         Left            =   1920
         TabIndex        =   30
         Top             =   2610
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ǰҹ"
         Height          =   180
         Left            =   1920
         TabIndex        =   29
         Top             =   2175
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1920
         TabIndex        =   28
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1920
         TabIndex        =   27
         Top             =   1725
         Width           =   360
      End
      Begin VB.Line Line2 
         X1              =   1800
         X2              =   1800
         Y1              =   1125
         Y2              =   2910
      End
      Begin VB.Line Line3 
         X1              =   2370
         X2              =   2370
         Y1              =   795
         Y2              =   2910
      End
      Begin VB.Label Label1 
         Caption         =   "ȫ        ��"
         Height          =   930
         Left            =   960
         TabIndex        =   26
         Top             =   1545
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1365
         TabIndex        =   25
         Top             =   1425
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ҹ��"
         Height          =   180
         Left            =   1365
         TabIndex        =   24
         Top             =   2370
         Width           =   360
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   1800
         X2              =   7980
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   1275
         X2              =   7980
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   1800
         X2              =   7980
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   0
         Left            =   3675
         TabIndex        =   23
         Top             =   1260
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   1
         Left            =   3675
         TabIndex        =   22
         Top             =   2145
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   2
         Left            =   3675
         TabIndex        =   21
         Top             =   1725
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   3
         Left            =   3675
         TabIndex        =   20
         Top             =   2610
         Width           =   90
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         Height          =   2145
         Left            =   780
         Top             =   780
         Width           =   7200
      End
   End
   Begin VB.PictureBox picCus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   0
      ScaleHeight     =   3330
      ScaleWidth      =   8175
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   8175
      Begin VB.PictureBox pic��ǰ��ɫ 
         BackColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   2850
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   47
         Top             =   2985
         Width           =   270
      End
      Begin MSMask.MaskEdBox txt���� 
         Height          =   300
         Left            =   2940
         TabIndex        =   11
         Top             =   2625
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt��ʼ 
         Height          =   300
         Left            =   2010
         TabIndex        =   10
         Top             =   2625
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lvwSeg 
         Height          =   2355
         Left            =   210
         TabIndex        =   8
         Top             =   165
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   4154
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ʱ���"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "��ʼʱ��"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "��ֹʱ��"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ȱʡԤԼʱ��"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "��ǰ�ź�ʱ��"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox txtʱ��� 
         Height          =   300
         Left            =   825
         MaxLength       =   4
         TabIndex        =   9
         Top             =   2625
         Width           =   720
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   495
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSplitTime.frx":0994
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   6945
         TabIndex        =   14
         Top             =   1830
         Width           =   1100
      End
      Begin VB.CommandButton cmdModi 
         Caption         =   "�޸�(&M)"
         Height          =   350
         Left            =   6945
         TabIndex        =   13
         Top             =   1320
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   6945
         TabIndex        =   12
         Top             =   825
         Width           =   1100
      End
      Begin MSMask.MaskEdBox mkTxtȱʡ 
         Height          =   300
         Left            =   5025
         TabIndex        =   37
         Top             =   2625
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt��ǰ 
         Height          =   300
         Left            =   7080
         TabIndex        =   46
         Top             =   2625
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ�ź�ʱ��ҺŰ�����ʾ��ɫ"
         Height          =   180
         Left            =   240
         TabIndex        =   48
         Top             =   3030
         Width           =   2520
      End
      Begin VB.Label lbl��ǰ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ�ź�ʱ��"
         Height          =   180
         Left            =   5955
         TabIndex        =   45
         Top             =   2685
         Width           =   1080
      End
      Begin VB.Label lblȱʡ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡԤԼʱ��"
         Height          =   180
         Left            =   3900
         TabIndex        =   38
         Top             =   2685
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Χ"
         Height          =   180
         Left            =   1590
         TabIndex        =   33
         Top             =   2685
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ���"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   2685
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSplitTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrColor As String

Private Sub cmdAdd_Click()
    Dim ObjItem As ListItem, i As Integer
    
    If Not OneValid Then Exit Sub
    
    For i = 1 To lvwSeg.ListItems.Count
        If lvwSeg.ListItems(i).Text = txtʱ���.Text Then
            MsgBox "��ʱ��������Ѿ����ڣ�", vbInformation, gstrSysName
            txtʱ���.SetFocus: Exit Sub
        End If
    Next
    If Checkȱʡʱ�� = False Then Exit Sub
    Set ObjItem = lvwSeg.ListItems.Add(, , txtʱ���.Text, , 1)
    ObjItem.SubItems(1) = txt��ʼ.Text
    ObjItem.SubItems(2) = txt����.Text
    ObjItem.Selected = True
    ObjItem.EnsureVisible
    lvwSeg.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    Dim ObjItem As ListItem
    
    If Not CheckValid Then Exit Sub
    
    If picStd.Visible Then
        lvwSeg.ListItems.Clear
        Set ObjItem = lvwSeg.ListItems.Add(, , "����", , 1)
        ObjItem.SubItems(1) = txt����S.Text
        ObjItem.SubItems(2) = txt����E.Text
        ObjItem.SubItems(3) = txt����ȱʡ.Text
        ObjItem.SubItems(4) = txt������ǰ.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "����", , 1)
        ObjItem.SubItems(1) = txt����S.Text
        ObjItem.SubItems(2) = txt����E.Text
        ObjItem.SubItems(3) = txt����ȱʡ.Text
        ObjItem.SubItems(4) = txt������ǰ.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "����", , 1)
        ObjItem.SubItems(1) = txt����S.Text
        ObjItem.SubItems(2) = txt����E.Text
        ObjItem.SubItems(3) = txt����ȱʡ.Text
        ObjItem.SubItems(4) = txt������ǰ.Text
        ObjItem.Selected = True
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "ǰҹ", , 1)
        ObjItem.SubItems(1) = txtǰҹS.Text
        ObjItem.SubItems(2) = txtǰҹE.Text
        ObjItem.SubItems(3) = txtǰҹȱʡ.Text
        ObjItem.SubItems(4) = txtǰҹ��ǰ.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "��ҹ", , 1)
        ObjItem.SubItems(1) = txt��ҹS.Text
        ObjItem.SubItems(2) = txt��ҹE.Text
        ObjItem.SubItems(3) = txt��ҹȱʡ.Text
        ObjItem.SubItems(4) = txt��ҹ��ǰ.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "ҹ��", , 1)
        ObjItem.SubItems(1) = txtǰҹS.Text
        ObjItem.SubItems(2) = txt��ҹE.Text
        ObjItem.SubItems(3) = txtǰҹȱʡ.Text
        ObjItem.SubItems(4) = txtǰҹ��ǰ.Text
        ObjItem.Selected = True
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "ȫ��", , 1)
        ObjItem.SubItems(1) = txt����S.Text
        ObjItem.SubItems(2) = txt��ҹE.Text
        ObjItem.SubItems(3) = txt��ҹȱʡ.Text
        ObjItem.SubItems(4) = txt��ҹ��ǰ.Text
        ObjItem.Selected = True
                    
        txt����S.Text = "__:__:__"
        txt����S.Text = "__:__:__"
        txtǰҹS.Text = "__:__:__"
        txt��ҹS.Text = "__:__:__"
        
        lvwSeg.ListItems(1).Selected = True
        Call lvwSeg_ItemClick(lvwSeg.SelectedItem)
        
        cmdChange.Caption = "תΪ��׼(&C)"
        picStd.Visible = False
        picCus.Visible = True
        lvwSeg.SetFocus
    Else
        Call SetStandard
        
        cmdChange.Caption = "תΪ�Զ���(&C)"
        lvwSeg.ListItems.Clear
        picCus.Visible = False
        picStd.Visible = True
        txt����S.SetFocus
    End If
End Sub

Private Sub cmdDel_Click()
    Dim intIdx As Integer
    
    If lvwSeg.SelectedItem Is Nothing Then
        MsgBox "û�п���ɾ����ʱ��Σ�", vbInformation, gstrSysName
        lvwSeg.SetFocus: Exit Sub
    End If
    
    intIdx = lvwSeg.SelectedItem.Index
    
    lvwSeg.ListItems.Remove intIdx
    
    If lvwSeg.ListItems.Count > 0 Then
        If intIdx <= lvwSeg.ListItems.Count Then
            lvwSeg.ListItems(intIdx).Selected = True
        Else
            lvwSeg.ListItems(lvwSeg.ListItems.Count).Selected = True
        End If
        lvwSeg.SelectedItem.EnsureVisible
        Call lvwSeg_ItemClick(lvwSeg.SelectedItem)
    Else
        txtʱ���.Text = ""
        txt��ʼ.Text = "__:__:__"
        txt����.Text = "__:__:__"
        mkTxtȱʡ.Text = "__:__:__"
        txt��ǰ.Text = "__:__:__"
    End If
    
    lvwSeg.SetFocus
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Function CheckValid() As Boolean
    If picStd.Visible Then
        If Not IsDate(txt����S.Text) Then
            MsgBox "����Ŀ�ʼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt����S.SetFocus: Exit Function
        End If
        If Not IsDate(txt����E.Text) Then
            MsgBox "�������ֹʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt����S.SetFocus: Exit Function
        End If
        If Not IsDate(txt����ȱʡ.Text) Then
            MsgBox "�����ȱʡԤԼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt����ȱʡ.SetFocus: Exit Function
        End If
        If Not IsDate(txt����S.Text) Then
            MsgBox "����Ŀ�ʼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt����S.SetFocus: Exit Function
        End If
        If Not IsDate(txt����E.Text) Then
            MsgBox "�������ֹʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt����S.SetFocus: Exit Function
        End If
        If Not IsDate(txt����ȱʡ.Text) Then
            MsgBox "�����ȱʡԤԼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt����ȱʡ.SetFocus: Exit Function
        End If
        
        If Not (IIf(txt����S.Text = "00:00:00", "24:00:00", txt����S.Text) < IIf(txt����S.Text = "00:00:00", "24:00:00", txt����S.Text)) Then
            MsgBox "���翪ʼʱ��Ӧ��С�����翪ʼʱ�䣡", vbInformation, gstrSysName
            txt����S.SetFocus: Exit Function
        End If
        
        If Replace(Replace(txt������ǰ.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txt������ǰ.Text) Then
                MsgBox "����ķź�ʱ�����ò���ȷ��", vbInformation, gstrSysName
                txt������ǰ.SetFocus: Exit Function
            End If
            If Format(txt������ǰ.Text, "HH:MM:SS") > Format(txt����S.Text, "HH:MM:SS") Then
                MsgBox "����ķź�ʱ�䲻�ܴ��ڿ�ʼʱ�䣡", vbInformation, gstrSysName
                txt������ǰ.SetFocus: Exit Function
            End If
        End If
        
        If Replace(Replace(txt������ǰ.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txt������ǰ.Text) Then
                MsgBox "����ķź�ʱ�����ò���ȷ��", vbInformation, gstrSysName
                txt������ǰ.SetFocus: Exit Function
            End If
            If Format(txt������ǰ.Text, "HH:MM:SS") > Format(txt����S.Text, "HH:MM:SS") Then
                MsgBox "����ķź�ʱ�䲻�ܴ��ڿ�ʼʱ�䣡", vbInformation, gstrSysName
                txt������ǰ.SetFocus: Exit Function
            End If
        End If
        
        If Replace(Replace(txtǰҹ��ǰ.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txtǰҹ��ǰ.Text) Then
                MsgBox "ǰҹ�ķź�ʱ�����ò���ȷ��", vbInformation, gstrSysName
                txtǰҹ��ǰ.SetFocus: Exit Function
            End If
            If Format(txtǰҹ��ǰ.Text, "HH:MM:SS") > Format(txtǰҹS.Text, "HH:MM:SS") Then
                MsgBox "ǰҹ�ķź�ʱ�䲻�ܴ��ڿ�ʼʱ�䣡", vbInformation, gstrSysName
                txtǰҹ��ǰ.SetFocus: Exit Function
            End If
        End If
        
        If Replace(Replace(txt��ҹ��ǰ.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txt��ҹ��ǰ.Text) Then
                MsgBox "��ҹ�ķź�ʱ�����ò���ȷ��", vbInformation, gstrSysName
                txt��ҹ��ǰ.SetFocus: Exit Function
            End If
            If Format(txt��ҹ��ǰ.Text, "HH:MM:SS") > Format(txt��ҹS.Text, "HH:MM:SS") Then
                MsgBox "��ҹ�ķź�ʱ�䲻�ܴ��ڿ�ʼʱ�䣡", vbInformation, gstrSysName
                txt��ҹ��ǰ.SetFocus: Exit Function
            End If
        End If
        
        If Not IsDate(txtǰҹS.Text) Then
            MsgBox "ǰҹ�Ŀ�ʼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txtǰҹS.SetFocus: Exit Function
        End If
        If Not IsDate(txtǰҹE.Text) Then
            MsgBox "ǰҹ����ֹʱ�����ò���ȷ��", vbInformation, gstrSysName
            txtǰҹS.SetFocus: Exit Function
        End If
        If Not IsDate(txtǰҹȱʡ.Text) Then
            MsgBox "ǰҹ��ȱʡԤԼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txtǰҹȱʡ.SetFocus: Exit Function
        End If
        
        If Not IsDate(txt��ҹS.Text) Then
            MsgBox "��ҹ�Ŀ�ʼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt��ҹS.SetFocus: Exit Function
        End If
        If Not IsDate(txt��ҹE.Text) Then
            MsgBox "��ҹ����ֹʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt��ҹS.SetFocus: Exit Function
        End If
        If Not IsDate(txt��ҹȱʡ.Text) Then
            MsgBox "��ҹ��ȱʡԤԼʱ�����ò���ȷ��", vbInformation, gstrSysName
            txt��ҹȱʡ.SetFocus: Exit Function
        End If
        If Not (IIf(txtǰҹS.Text = "00:00:00", "24:00:00", txtǰҹS.Text) < IIf(txt��ҹS.Text = "00:00:00", "24:00:00", txt��ҹS.Text)) Then
            MsgBox "ǰҹ��ʼʱ��Ӧ��С�ں�ҹ��ʼʱ�䣡", vbInformation, gstrSysName
            txtǰҹS.SetFocus: Exit Function
        End If
    Else
        If lvwSeg.ListItems.Count = 0 Then
            MsgBox "������������һ��ʱ��Σ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckValid = True
End Function

Private Function OneValid() As Boolean
    If Trim(txtʱ���.Text) = "" Then
        MsgBox "��������ʱ������ƣ�", vbInformation, gstrSysName
        txtʱ���.SetFocus: Exit Function
    End If
    If zlCommFun.ActualLen(txtʱ���.Text) > 4 Then
        MsgBox "ʱ�������ֻ��Ϊ�������ֻ�4����ĸ��", vbInformation, gstrSysName
        txtʱ���.SetFocus: Exit Function
    End If
    If Not IsDate(txt��ʼ.Text) Then
        MsgBox "��ʼʱ�����ò���ȷ��", vbInformation, gstrSysName
        txt��ʼ.SetFocus: Exit Function
    End If
    If Not IsDate(txt����.Text) Then
        MsgBox "����ʱ�����ò���ȷ��", vbInformation, gstrSysName
        txt����.SetFocus: Exit Function
    End If
    If txt��ʼ.Text = txt����.Text Then
        MsgBox "��ʼ�ͽ���ʱ�䲻Ӧ����ͬ��", vbInformation, gstrSysName
        txt����.SetFocus: Exit Function
    End If
    If Checkȱʡʱ�� = False Then Exit Function
    
    OneValid = True
End Function

Private Sub cmdModi_Click()
    Dim i As Integer
    
    If lvwSeg.SelectedItem Is Nothing Then
        MsgBox "û��ʱ��ο����޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Not OneValid Then Exit Sub
    
    For i = 1 To lvwSeg.ListItems.Count
        If i <> lvwSeg.SelectedItem.Index And lvwSeg.ListItems(i).Text = txtʱ���.Text Then
            MsgBox "��ʱ��������Ѿ����ڣ�", vbInformation, gstrSysName
            txtʱ���.SetFocus: Exit Sub
        End If
    Next
    If Checkȱʡʱ�� = False Then Exit Sub
    
    lvwSeg.SelectedItem.Text = txtʱ���.Text
    lvwSeg.SelectedItem.SubItems(1) = txt��ʼ.Text
    lvwSeg.SelectedItem.SubItems(2) = txt����.Text
    lvwSeg.SelectedItem.SubItems(3) = mkTxtȱʡ.Text
    lvwSeg.SelectedItem.SubItems(4) = txt��ǰ.Text
    
    lvwSeg.SetFocus
End Sub
Private Function Checkȱʡʱ��() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ȱʡʱ���Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2012-03-12 14:46:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strDate As String
    
    If mkTxtȱʡ.Text Like "*_*" Then
        MsgBox "ȱʡԤԼʱ������,����������!", vbInformation + vbOKOnly, gstrSysName
        If mkTxtȱʡ.Enabled Then mkTxtȱʡ.SetFocus
        Exit Function
    End If
    
    If IsDate(mkTxtȱʡ.Text) = False Then
        MsgBox "ȱʡԤԼʱ���ʽ����,����������!", vbInformation + vbOKOnly, gstrSysName
        If mkTxtȱʡ.Enabled Then mkTxtȱʡ.SetFocus
        Exit Function
    End If
    
    If Replace(Replace(txt��ǰ.Text, "_", ""), ":", "") <> "" Then
        If IsDate(txt��ǰ.Text) = False Then
            MsgBox "�ź�ʱ���ʽ����,����������!", vbInformation + vbOKOnly, gstrSysName
            If txt��ǰ.Enabled Then txt��ǰ.SetFocus
            Exit Function
        End If
    End If
    
    strDate = "2010-01-01 "
    If CDate("2010-01-01 " & txt��ʼ.Text) > CDate("2010-01-01 " & txt����.Text) Then
        strDate = "2010-01-02 "
    End If
    
    If CDate(strDate & mkTxtȱʡ.Text) < CDate("2010-01-01 " & txt��ʼ.Text) _
        Or CDate(strDate & mkTxtȱʡ.Text) > CDate(strDate & txt����.Text) Then
        MsgBox "ȱʡԤԼʱ�������ʱ�䷶Χ��,����������!", vbInformation + vbOKOnly, gstrSysName
        If mkTxtȱʡ.Enabled Then mkTxtȱʡ.SetFocus
        Exit Function
    End If
    
    If Replace(Replace(txt��ǰ.Text, "_", ""), ":", "") <> "" Then
        If Format(txt��ʼ.Text, "HH:MM:SS") < Format(txt��ǰ.Text, "HH:MM:SS") Then
            MsgBox "�ź�ʱ�����С�ڿ�ʼʱ��,����������!", vbInformation + vbOKOnly, gstrSysName
            If txt��ǰ.Enabled Then txt��ǰ.SetFocus
            Exit Function
        End If
    End If
    
    Checkȱʡʱ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub cmdOK_Click()
    Dim arrSQL() As String, i As Integer, blnTrans As Boolean
    
    '�����Լ��
    If Not CheckValid Then Exit Sub
    
    ReDim arrSQL(0)
    arrSQL(0) = "zl_ʱ���_Clear"
    
    If picStd.Visible Then
        ReDim Preserve arrSQL(7)
        arrSQL(1) = "zl_ʱ���_INSERT('����',To_Date('" & Format(txt����S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt����E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt����ȱʡ.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt������ǰ.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt������ǰ.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "')"
        arrSQL(2) = "zl_ʱ���_INSERT('��ҹ',To_Date('" & Format(txt��ҹS.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt��ҹE.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt��ҹȱʡ.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt��ҹ��ǰ.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt��ҹ��ǰ.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "')"
        arrSQL(3) = "zl_ʱ���_INSERT('ǰҹ',To_Date('" & Format(txtǰҹS.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txtǰҹE.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txtǰҹȱʡ.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txtǰҹ��ǰ.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txtǰҹ��ǰ.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "')"
        arrSQL(4) = "zl_ʱ���_INSERT('ȫ��',To_Date('" & Format(txt����S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt��ҹE.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt��ҹȱʡ.Text, "HH:MM:SS") & "','HH24:MI:SS')," & "Null,'" & mstrColor & "')"
        arrSQL(5) = "zl_ʱ���_INSERT('����',To_Date('" & Format(txt����S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt����E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt����ȱʡ.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt������ǰ.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt������ǰ.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "')"
        arrSQL(6) = "zl_ʱ���_INSERT('����',To_Date('" & Format(txt����S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt����E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt����ȱʡ.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt������ǰ.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt������ǰ.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "')"
        arrSQL(7) = "zl_ʱ���_INSERT('ҹ��',To_Date('" & Format(txtǰҹS.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt��ҹE.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt��ҹȱʡ.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txtǰҹ��ǰ.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txtǰҹ��ǰ.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "')"
    Else
        ReDim Preserve arrSQL(lvwSeg.ListItems.Count)
        For i = 1 To lvwSeg.ListItems.Count
            With lvwSeg.ListItems(i)
                arrSQL(i) = "zl_ʱ���_INSERT('" & .Text & "',To_Date('" & Format(.SubItems(1), "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(.SubItems(2), "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(.SubItems(3), "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(.SubItems(4), "_", ""), ":", "") = "", "Null", "To_Date('" & Format(.SubItems(4), "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "')"
            End With
        Next
    End If
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(arrSQL(i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Call FillData
End Sub

Public Function FillData() As Boolean
'����:������ʱ���װ�뵽msfTime
    Dim rsTime As New ADODB.Recordset
    Dim strSQL As String, ObjItem As ListItem
    
    Dim vBegin As Date, vEnd As Date, blnStd As Boolean

    On Error GoTo errH
    

    strSQL = "Select ʱ���,To_Char(��ʼʱ��,'HH24:MI:SS') As ��ʼʱ��,to_char(��ֹʱ��,'HH24:MI:SS') As ��ֹʱ�� ,to_char(ȱʡʱ��,'HH24:MI:SS') As ȱʡʱ��,to_char(��ǰʱ��,'HH24:MI:SS') As ��ǰʱ��,��ǰ��ɫ  From ʱ��� Where ���� Is Null And վ�� Is Null"
    Set rsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With rsTime
        If Not .EOF Then
            '�ж��Ƿ���ϱ�׼ʱ��ε�����
            mstrColor = Nvl(rsTime!��ǰ��ɫ)
            If mstrColor = "" Then mstrColor = &H0&
            pic��ǰ��ɫ(0).BackColor = mstrColor
            pic��ǰ��ɫ(1).BackColor = mstrColor
            
            blnStd = rsTime.RecordCount = 7
            If blnStd Then
                rsTime.Filter = "ʱ���='����' or ʱ���='����' or ʱ���='����' or ʱ���='ǰҹ' or ʱ���='��ҹ' or ʱ���='ҹ��' or ʱ���='ȫ��'"
                blnStd = blnStd And rsTime.RecordCount = 7
            End If
            
            If blnStd Then
                '����Ľ���=����Ŀ�ʼ
                rsTime.Filter = "ʱ���='����'": vEnd = rsTime!��ֹʱ��
                rsTime.Filter = "ʱ���='����'": vBegin = rsTime!��ʼʱ��
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                
                '����Ľ���=ǰҹ�Ŀ�ʼ
                rsTime.Filter = "ʱ���='����'": vEnd = rsTime!��ֹʱ��
                rsTime.Filter = "ʱ���='ǰҹ'": vBegin = rsTime!��ʼʱ��
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                
                'ǰҹ�Ľ���=��ҹ�Ŀ�ʼ
                rsTime.Filter = "ʱ���='ǰҹ'": vEnd = rsTime!��ֹʱ��
                rsTime.Filter = "ʱ���='��ҹ'": vBegin = rsTime!��ʼʱ��
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                
                '��ҹ�Ľ���=����Ŀ�ʼ
                rsTime.Filter = "ʱ���='��ҹ'": vEnd = rsTime!��ֹʱ��
                rsTime.Filter = "ʱ���='����'": vBegin = rsTime!��ʼʱ��
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                '--------------------------------------------------------------------------
                '����Ŀ�ʼ=����Ŀ�ʼ
                rsTime.Filter = "ʱ���='����'": vEnd = rsTime!��ʼʱ��
                rsTime.Filter = "ʱ���='����'": vBegin = rsTime!��ʼʱ��
                blnStd = blnStd And vEnd = vBegin
                
                '����Ľ���=����Ľ���
                rsTime.Filter = "ʱ���='����'": vEnd = rsTime!��ֹʱ��
                rsTime.Filter = "ʱ���='����'": vBegin = rsTime!��ֹʱ��
                blnStd = blnStd And vEnd = vBegin
                
                'ҹ��Ŀ�ʼ=ǰҹ�Ŀ�ʼ
                rsTime.Filter = "ʱ���='ҹ��'": vEnd = rsTime!��ʼʱ��
                rsTime.Filter = "ʱ���='ǰҹ'": vBegin = rsTime!��ʼʱ��
                blnStd = blnStd And vEnd = vBegin
                
                'ҹ��Ľ���=��ҹ�Ľ���
                rsTime.Filter = "ʱ���='ҹ��'": vEnd = rsTime!��ֹʱ��
                rsTime.Filter = "ʱ���='��ҹ'": vBegin = rsTime!��ֹʱ��
                blnStd = blnStd And vEnd = vBegin
                
                'ȫ�յĿ�ʼ=����Ŀ�ʼ
                rsTime.Filter = "ʱ���='ȫ��'": vEnd = rsTime!��ʼʱ��
                rsTime.Filter = "ʱ���='����'": vBegin = rsTime!��ʼʱ��
                blnStd = blnStd And vEnd = vBegin
                
                'ȫ�յĽ���=��ҹ�Ľ���
                rsTime.Filter = "ʱ���='ȫ��'": vEnd = rsTime!��ֹʱ��
                rsTime.Filter = "ʱ���='��ҹ'": vBegin = rsTime!��ֹʱ��
                blnStd = blnStd And vEnd = vBegin
            End If
            
            .Filter = 0
            .MoveFirst
            If blnStd Then
                Do Until .EOF
                    Select Case .Fields("ʱ���").Value
                    Case "��ҹ"
                        txt��ҹS.Text = IIf(IsNull(.Fields("��ʼʱ��").Value), "__:__:__", .Fields("��ʼʱ��").Value)
                        txt��ҹE.Text = IIf(IsNull(.Fields("��ֹʱ��").Value), "__:__:__", .Fields("��ֹʱ��").Value)
                        txt��ҹȱʡ.Text = IIf(IsNull(.Fields("ȱʡʱ��").Value), "__:__:__", .Fields("ȱʡʱ��").Value)
                        txt��ҹ��ǰ.Text = IIf(IsNull(.Fields("��ǰʱ��").Value), "__:__:__", .Fields("��ǰʱ��").Value)
                    Case "ǰҹ"
                        txtǰҹS.Text = IIf(IsNull(.Fields("��ʼʱ��").Value), "__:__:__", .Fields("��ʼʱ��").Value)
                        txtǰҹE.Text = IIf(IsNull(.Fields("��ֹʱ��").Value), "__:__:__", .Fields("��ֹʱ��").Value)
                        txtǰҹȱʡ.Text = IIf(IsNull(.Fields("ȱʡʱ��").Value), "__:__:__", .Fields("ȱʡʱ��").Value)
                        txtǰҹ��ǰ.Text = IIf(IsNull(.Fields("��ǰʱ��").Value), "__:__:__", .Fields("��ǰʱ��").Value)
                    Case "����"
                        txt����S.Text = IIf(IsNull(.Fields("��ʼʱ��").Value), "__:__:__", .Fields("��ʼʱ��").Value)
                        txt����E.Text = IIf(IsNull(.Fields("��ֹʱ��").Value), "__:__:__", .Fields("��ֹʱ��").Value)
                        txt����ȱʡ.Text = IIf(IsNull(.Fields("ȱʡʱ��").Value), "__:__:__", .Fields("ȱʡʱ��").Value)
                        txt������ǰ.Text = IIf(IsNull(.Fields("��ǰʱ��").Value), "__:__:__", .Fields("��ǰʱ��").Value)
                    Case "����"
                        txt����S.Text = IIf(IsNull(.Fields("��ʼʱ��").Value), "__:__:__", .Fields("��ʼʱ��").Value)
                        txt����E.Text = IIf(IsNull(.Fields("��ֹʱ��").Value), "__:__:__", .Fields("��ֹʱ��").Value)
                        txt����ȱʡ.Text = IIf(IsNull(.Fields("ȱʡʱ��").Value), "__:__:__", .Fields("ȱʡʱ��").Value)
                        txt������ǰ.Text = IIf(IsNull(.Fields("��ǰʱ��").Value), "__:__:__", .Fields("��ǰʱ��").Value)
                    End Select
                    rsTime.MoveNext
                Loop
            Else
                Do Until .EOF
                    Set ObjItem = lvwSeg.ListItems.Add(, , !ʱ���, , 1)
                    ObjItem.SubItems(1) = IIf(IsNull(!��ʼʱ��), "__:__:__", !��ʼʱ��)
                    ObjItem.SubItems(2) = IIf(IsNull(!��ֹʱ��), "__:__:__", !��ֹʱ��)
                    ObjItem.SubItems(3) = IIf(IsNull(!ȱʡʱ��), "__:__:__", !ȱʡʱ��)
                    ObjItem.SubItems(4) = IIf(IsNull(!��ǰʱ��), "__:__:__", !��ǰʱ��)
                    rsTime.MoveNext
                Loop
                lvwSeg.ListItems(1).Selected = True
                Call lvwSeg_ItemClick(lvwSeg.SelectedItem)
                
                cmdChange.Caption = "תΪ��׼(&C)"
                picStd.Visible = False
                picCus.Visible = True
            End If
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If ActiveControl.Name = "cmdOK" Or ActiveControl.Name = "cmdCancel" Then Exit Sub
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub lvwSeg_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtʱ���.Text = Item.Text
    txt��ʼ.Text = Item.SubItems(1)
    txt����.Text = Item.SubItems(2)
    mkTxtȱʡ.Text = Item.SubItems(3)
    txt��ǰ.Text = Item.SubItems(4)
End Sub

Private Sub pic��ǰ��ɫ_Click(Index As Integer)
    dlgColor.ShowColor
    mstrColor = dlgColor.Color
    pic��ǰ��ɫ(0).BackColor = mstrColor
    pic��ǰ��ɫ(1).BackColor = mstrColor
End Sub

Private Sub txt��ҹS_GotFocus()
    SelAll txt��ҹS
End Sub

Private Sub txt��ҹS_LostFocus()
    If IsDate(txt��ҹS.Text) Then
        Me.txtǰҹE.Text = Format(DateAdd("s", -1, CDate(Me.txt��ҹS.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub txt����_Change()
    If mkTxtȱʡ.Text Like "*_*" Then mkTxtȱʡ.Text = txt����
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
End Sub

Private Sub txt��ʼ_GotFocus()
    SelAll txt��ʼ
End Sub

Private Sub txtǰҹS_GotFocus()
    SelAll txtǰҹS
End Sub

Private Sub txtǰҹS_LostFocus()
    If IsDate(txtǰҹS.Text) Then
        Me.txt����E.Text = Format(DateAdd("s", -1, CDate(Me.txtǰҹS.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub txt����S_GotFocus()
    SelAll txt����S
End Sub

Private Sub txt����S_LostFocus()
    If IsDate(txt����S.Text) Then
        Me.txt��ҹE.Text = Format(DateAdd("s", -1, CDate(Me.txt����S.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub txtʱ���_GotFocus()
    SelAll txtʱ���
End Sub

Private Sub txt��ǰ_GotFocus()
    SelAll txt��ǰ
End Sub

Private Sub txt����S_GotFocus()
    SelAll txt����S
End Sub

Private Sub txt����S_LostFocus()
    If IsDate(Me.txt����S.Text) Then
        Me.txt����E.Text = Format(DateAdd("s", -1, CDate(Me.txt����S.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub SetStandard()
'���ܣ����Զ���ʱ���ת��Ϊ��׼ʱ���
    Dim i As Integer
    
    For i = 1 To lvwSeg.ListItems.Count
        With lvwSeg.ListItems(i)
            Select Case .Text
                Case "����"
                    txt����S.Text = .SubItems(1)
                    txt����E.Text = .SubItems(2)
                Case "����"
                    txt����S.Text = .SubItems(1)
                    txt����E.Text = .SubItems(2)
                Case "����", "����"
                    If Not IsDate(txt����S.Text) Then txt����S.Text = .SubItems(1)
                    If Not IsDate(txt����E.Text) Then txt����E.Text = .SubItems(2)
                Case "ǰҹ", "��ҹ"
                    txtǰҹS.Text = .SubItems(1)
                    txtǰҹE.Text = .SubItems(2)
                Case "��ҹ", "��ҹ"
                    txt��ҹS.Text = .SubItems(1)
                    txt��ҹE.Text = .SubItems(2)
                Case "ҹ��", "����"
                    If Not IsDate(txtǰҹS.Text) Then txtǰҹS.Text = .SubItems(1)
                    If Not IsDate(txt��ҹE.Text) Then txt��ҹE.Text = .SubItems(2)
                Case "ȫ��", "ȫ��"
                    If Not IsDate(txt����S.Text) Then txt����S.Text = .SubItems(1)
                    If Not IsDate(txt��ҹE.Text) Then txt��ҹE.Text = .SubItems(2)
            End Select
        End With
    Next
End Sub
