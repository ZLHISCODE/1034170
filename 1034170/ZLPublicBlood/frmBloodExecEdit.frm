VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBloodExecEdit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ѫִ�еǼ�"
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
   StartUpPosition =   1  '����������
   Begin VB.Frame fraExe 
      Caption         =   "��ע����"
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
            Caption         =   "��"
            Height          =   240
            Index           =   1
            Left            =   630
            TabIndex        =   26
            Top             =   45
            Width           =   525
         End
         Begin VB.OptionButton optLeakage 
            Caption         =   "��"
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
      End
      Begin VB.ComboBox cboִ���� 
         Height          =   300
         Index           =   2
         Left            =   4845
         TabIndex        =   22
         Text            =   "cboִ����"
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
            Caption         =   "��"
            Height          =   240
            Index           =   4
            Left            =   0
            TabIndex        =   29
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton optBegin 
            Caption         =   "��"
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
         ToolTipText     =   "�༭(F4)"
         Top             =   990
         Width           =   255
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Index           =   1
         Left            =   1290
         TabIndex        =   20
         Top             =   210
         Width           =   930
      End
      Begin MSMask.MaskEdBox txt��Ӧʱ�� 
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
            Name            =   "����"
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
            Name            =   "����"
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
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         Caption         =   "��/��"
         Height          =   180
         Index           =   1
         Left            =   2265
         TabIndex        =   77
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblLeakage 
         AutoSize        =   -1  'True
         Caption         =   "��Ѫ��λ������©"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   23
         Top             =   630
         Width           =   1440
      End
      Begin VB.Label lblPeople 
         AutoSize        =   -1  'True
         Caption         =   "ִ �� ��"
         Height          =   180
         Index           =   2
         Left            =   4050
         TabIndex        =   21
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblRe 
         AutoSize        =   -1  'True
         Caption         =   "������Ӧ"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   31
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReactionTime 
         AutoSize        =   -1  'True
         Caption         =   "��Ӧʱ��"
         Height          =   180
         Index           =   2
         Left            =   4050
         TabIndex        =   33
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReaction 
         AutoSize        =   -1  'True
         Caption         =   "��Ѫ��Ӧ"
         Height          =   180
         Index           =   2
         Left            =   4050
         TabIndex        =   27
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "15���Ӻ����"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   19
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame fraExe 
      Caption         =   "��ע�˶�"
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
         Text            =   "����Ա"
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
         Text            =   "����Ա"
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
         Caption         =   "�˶���֤"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�˶�ʱ��"
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
         Caption         =   "�� �� ��"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   70
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         Caption         =   "�� �� ��"
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
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4500
      TabIndex        =   59
      Top             =   7710
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   5625
      TabIndex        =   60
      Top             =   7710
      Width           =   1100
   End
   Begin VB.TextBox txtִ��ժҪ 
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
      Caption         =   "��ע����"
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
            Caption         =   "��"
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   45
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton optLeakage 
            Caption         =   "��"
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
         ToolTipText     =   "�༭(F4)"
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
         ToolTipText     =   "�༭(F4)"
         Top             =   240
         Width           =   255
      End
      Begin MSMask.MaskEdBox txtִ��ʱ�� 
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
            Name            =   "����"
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
      Begin VB.ComboBox cboִ���� 
         Height          =   300
         Index           =   1
         Left            =   4845
         TabIndex        =   42
         Text            =   "cboִ����"
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
            Caption         =   "��"
            Height          =   240
            Index           =   3
            Left            =   1335
            TabIndex        =   50
            Top             =   45
            Width           =   525
         End
         Begin VB.OptionButton optBegin 
            Caption         =   "��"
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
      Begin MSMask.MaskEdBox txt��Ӧʱ�� 
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "��Ѫ��λ������©"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   43
         Top             =   630
         Width           =   1440
      End
      Begin VB.Label lblExeTime 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   38
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblPeople 
         AutoSize        =   -1  'True
         Caption         =   "ִ �� ��"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   41
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblReaction 
         AutoSize        =   -1  'True
         Caption         =   "��Ѫ��Ӧ"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   47
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblReactionTime 
         AutoSize        =   -1  'True
         Caption         =   "��Ӧʱ��"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   53
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblRe 
         AutoSize        =   -1  'True
         Caption         =   "������Ӧ"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   51
         Top             =   1020
         Width           =   720
      End
   End
   Begin VB.Frame fraExe 
      Caption         =   "��ע��ʼ"
      Height          =   1725
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1230
      Width           =   6705
      Begin VB.ComboBox cbo���� 
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
         ToolTipText     =   "�༭(F4)"
         Top             =   240
         Width           =   255
      End
      Begin MSMask.MaskEdBox txtִ��ʱ�� 
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
            Name            =   "����"
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
         ToolTipText     =   "�༭(F4)"
         Top             =   990
         Width           =   255
      End
      Begin MSMask.MaskEdBox txt��Ӧʱ�� 
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
            Name            =   "����"
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
            Caption         =   "��"
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton optBegin 
            Caption         =   "��"
            Height          =   240
            Index           =   1
            Left            =   1335
            TabIndex        =   11
            Top             =   45
            Width           =   525
         End
      End
      Begin VB.ComboBox cboִ���� 
         Height          =   300
         Index           =   0
         Left            =   4845
         TabIndex        =   5
         Text            =   "cboִ����"
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
            Name            =   "����"
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
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         Caption         =   "��/��"
         Height          =   180
         Index           =   0
         Left            =   2265
         TabIndex        =   76
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblRe 
         AutoSize        =   -1  'True
         Caption         =   "������Ӧ"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   12
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReactionTime 
         AutoSize        =   -1  'True
         Caption         =   "��Ӧʱ��"
         Height          =   180
         Index           =   0
         Left            =   4050
         TabIndex        =   14
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReaction 
         AutoSize        =   -1  'True
         Caption         =   "��Ѫ��Ӧ"
         Height          =   180
         Index           =   0
         Left            =   4050
         TabIndex        =   8
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ǰ15���ӵ���"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label lblPeople 
         AutoSize        =   -1  'True
         Caption         =   "ִ �� ��"
         Height          =   180
         Index           =   0
         Left            =   4050
         TabIndex        =   4
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblExeTime 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��"
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
      Begin VB.TextBox txt�������� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   195
         TabIndex        =   63
         Top             =   0
         Width           =   1005
      End
      Begin VB.TextBox txt�������� 
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
      Begin MSComCtl2.DTPicker dtpҪ��ʱ�� 
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
      Caption         =   "���ѣ���ע��ʼ������4h�����ѪҺ��ע"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ִ��ժҪ"
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
Private mstrȱʡ��Ѫ��Ӧ As String
Private mlngModul As Long
Private mlng�շ�ID As Long
Private mlngҽ��ID As Long, mlng���ID As Long
Private mlng���ͺ� As Long
Private mlng����ID As Long
Private mlngִ�п���ID As Long
Private mstrPrivs As String
Private mblnOk As Boolean
Private mstr����ʱ�� As String 'ѪҺ�Ľ���ʱ��
Private mintѪ���� As Integer, mint��ִ��Ѫ���� As Integer
Private mintTimerCount As Integer
Private mblnReturn As Boolean  'ִ���˿�������ƥ�����
Private mblnOnlyRead As Boolean '�Ƿ���ֻ��ģʽ

Private Type ExeInfo
    �շ�ID As Long
    ��ʼִ���� As String
    ��ʼʱ��  As String
    ǰ15���ӵ��� As Integer
    ��עǰ��Ѫ��Ӧ  As String
    ��עǰ��Ӧʱ�� As String
    ��ע��ִ���� As String
    ��15���ӵ��� As Integer   '��ע15���Ӻ�ĵ���
    ��ע����Ѫ��Ӧ  As String
    ��ע�з�Ӧʱ�� As String
    ����ִ���� As String
    ����ʱ��  As String
    ��ע����Ѫ��Ӧ  As String
    ��ע��Ӧʱ�� As String
    ִ�п���ID  As Long
    �˲��� As String
    ������ As String
    �˶�ʱ�� As String
    �Ǽ��� As String
    �Ǽ�ʱ��  As String
    ժҪ  As String
    ��Ѫ��λ��© As String
End Type
Private mExeInfo As ExeInfo
Private mExeInfoSave As ExeInfo

Public Function ShowEdit(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lngҽ��ID As Long, _
    ByVal lng���ͺ� As Long, ByVal lng����ID As Long, ByVal lng�շ�ID As Long, ByVal lngִ�п���ID As Long, Optional ByVal strPrivs As String, Optional ByVal blnOnlyRead As Boolean) As Boolean
    mblnOk = True
    mlngModul = lngModul
    mlngҽ��ID = lngҽ��ID
    mlng���ͺ� = lng���ͺ�
    mlng����ID = lng����ID
    mlng�շ�ID = lng�շ�ID
    mlngִ�п���ID = lngִ�п���ID
    mstrPrivs = strPrivs
    mblnOnlyRead = blnOnlyRead
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
    ShowEdit = mblnOk
End Function

Private Sub ClearExeInfo()
    With mExeInfo
        .�շ�ID = 0
        .��ʼִ���� = ""
        .��ʼʱ�� = ""
        .ǰ15���ӵ��� = 0
        .��עǰ��Ѫ��Ӧ = ""
        .��עǰ��Ӧʱ�� = ""
        .��ע��ִ���� = ""
        .��15���ӵ��� = 0
        .��ע����Ѫ��Ӧ = ""
        .��ע�з�Ӧʱ�� = ""
        .����ִ���� = ""
        .����ʱ�� = ""
        .��ע����Ѫ��Ӧ = ""
        .��ע��Ӧʱ�� = ""
        .ִ�п���ID = 0
        .�˲��� = ""
        .������ = ""
        .�˶�ʱ�� = ""
        .�Ǽ��� = ""
        .�Ǽ�ʱ�� = ""
        .ժҪ = ""
        .��Ѫ��λ��© = ""
    End With
End Sub

Private Sub cbo����_Click(Index As Integer)
    If cbo����(Index).ListIndex = 2 Or cbo����(Index).ListIndex = 3 Then
        lbl��λ(Index).Visible = False
    Else
        lbl��λ(Index).Visible = True
    End If
End Sub

Private Sub cbo����_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call cbo����_Click(Index)
End Sub

Private Sub cboִ����_Click(Index As Integer)
    Call cboִ����_KeyPress(Index, vbKeyReturn)
End Sub

Private Sub cboִ����_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strPersons As String
    Dim blnOk As Boolean
    
    mblnReturn = False
    If cboִ����(Index).ListIndex <> -1 Then cboִ����(Index).Tag = cboִ����(Index).ListIndex
    strPersons = cboִ����(Index).Text
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        If cboִ����(Index).Text <> "" Then
            Set rsTmp = GetDataToPersons(cboִ����(Index).Text)
            If Not rsTmp Is Nothing Then
                If rsTmp.State = adStateOpen Then
                    blnOk = Not rsTmp.EOF
                End If
            End If
            If blnOk Then
                Call FindCboIndex(cboִ����(Index), rsTmp!id)
            Else
                cboִ����(Index).ListIndex = Val(cboִ����(Index).Tag)
            End If
            Call gobjControl.CboSetIndex(cboִ����(Index).hWnd, Val(cboִ����(Index).ListIndex))
        Else
            cboִ����(Index).ListIndex = Val(cboִ����(Index).Tag)
        End If
        If strPersons = cboִ����(Index).Text Then
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cboִ����_Validate(Index As Integer, Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call gobjControl.CboSetIndex(cboִ����(Index).hWnd, Val(cboִ����(Index).Tag))
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
        Set objControl = txtִ��ʱ��(0)
        lngLeft = fraExe(0).Left
        lngTop = fraExe(0).Top
    ElseIf Index = 1 Then
        Set objControl = txtִ��ʱ��(1)
        lngLeft = fraExe(0).Left
        lngTop = fraExe(0).Top
    ElseIf Index = 2 Then
        Set objControl = txt��Ӧʱ��(0)
        lngLeft = fraExe(1).Left
        lngTop = fraExe(1).Top
    ElseIf Index = 3 Then
        Set objControl = txt��Ӧʱ��(1)
        lngLeft = fraExe(1).Left
        lngTop = fraExe(1).Top
    ElseIf Index = 4 Then
        Set objControl = txt��Ӧʱ��(2)
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
    Dim dbl�������� As Double, dblʣ����� As Double
    Dim blnTrans As Boolean, strsql As String
    Dim i As Integer, intIndex As Integer
    Dim str��Ѫ��λ������© As String
    Dim str��ʼִ��ʱ�� As String
    Dim rsTmp As New Recordset
    Dim arrMsg As Variant
    
    If lblLink.Tag <> "�Ѻ˶�" Then
        MsgBox "���Ƚ�����עǰ�˶ԣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If IsDate(txtִ��ʱ��(0).Text) = False Then
        MsgBox "��ע��ʼʱ�䲻����Ч�����ڸ�ʽ��", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txtִ��ʱ��(0))
        Exit Sub
    End If
    
    If Trim(cboִ����(0).Text) = "" Then
        MsgBox "��ѡ����ע��ʼִ���ˡ�", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(cboִ����(0))
        Exit Sub
    End If
    
    '117041:��ʼʱ����ͬ����ʱ��������һ��ʱ���Զ���һ��
    str��ʼִ��ʱ�� = txtִ��ʱ��(0).Text
    '��鱾��ִ��ʱ���Ƿ�����ϴ�ִ��ʱ��
    If IsDate(txtִ��ʱ��(0).Tag) Then
        If CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")) < CDate(Format(txtִ��ʱ��(0).Tag, "yyyy-MM-dd HH:mm")) Then
            MsgBox "����ִ��ʱ�䲻��С���ϴ�ִ��ʱ�� " & Format(txtִ��ʱ��(0).Tag, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txtִ��ʱ��(0))
            Exit Sub
        ElseIf CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")) = CDate(Format(txtִ��ʱ��(0).Tag, "yyyy-MM-dd HH:mm")) Then
            str��ʼִ��ʱ�� = Format(DateAdd("s", 1, CDate(Format(txtִ��ʱ��(0).Tag, "yyyy-MM-dd HH:mm:ss"))), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
     If CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")) < CDate(Format(txtCheck(2).Text, "yyyy-MM-dd HH:mm")) Then
        MsgBox "����ִ��ʱ�䲻��С�ں˶�ʱ�䡣", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txtִ��ʱ��(0))
        Exit Sub
    End If
    
    If cbo����(0).ListIndex <= 0 Then
        If LenB(StrConv(cbo����(0).Text, vbFromUnicode)) > 3 Or (Not IsNumeric(cbo����(0).Text) And cbo����(0).Text <> "") Then
            MsgBox "����¼���ǰ15���ӵ���ֻ�������֣������ֻ����¼��3λ���֣�", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(cbo����(0))
            Exit Sub
        End If
    End If
    
    '��ע�����е�У��
    If cbo����(1).ListIndex <= 0 Then
        If LenB(StrConv(cbo����(1).Text, vbFromUnicode)) > 3 Or (Not IsNumeric(cbo����(1).Text) And cbo����(1).Text <> "") Then
            MsgBox "����¼���15���Ӻ����ֻ�������֣������ֻ����¼��3λ���֣�", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(cbo����(1))
            Exit Sub
        End If
    End If
    
    If (cbo����(1).Text <> "" Or optLeakage(1).Value = True Or optBegin(5).Value = True) And Trim(cboִ����(2).Text) = "" Then
        MsgBox "��ѡ����ע����ִ���ˡ�", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(cboִ����(2))
        Exit Sub
    End If
    
    '��ע������У��
    If Not IsDate(txtִ��ʱ��(1).Text) And txtִ��ʱ��(1).Text <> "____-__-__ __:__" Then
        MsgBox "��д����ע����ʱ��ʱ����������Ч�����ڸ�ʽ��", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txtִ��ʱ��(1))
        Exit Sub
    End If
    
    If IsDate(txtִ��ʱ��(1).Text) And Trim(cboִ����(1).Text) = "" Then
        MsgBox "��ѡ����ע����ִ���ˡ�", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(cboִ����(1))
        Exit Sub
    End If
    
    If optLeakage(3).Value = True And Not IsDate(txtִ��ʱ��(1).Text) Then
        MsgBox "��ע���������Ѫ��λ����©ʱ���������д��ע����ʱ�䣡", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txtִ��ʱ��(1))
        Exit Sub
    End If
    '��ע����ѡ����Ѫ��Ӧ�������д����ʱ��
    If optBegin(3).Value = True And Not IsDate(txtִ��ʱ��(1).Text) Then
        MsgBox "��ע������Ĵ�����Ѫ��Ӧʱ���������д��ע����ʱ�䣡", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txtִ��ʱ��(1))
        Exit Sub
    End If
    '��Ѫ��Ӧ��дͳһУ��
    For i = 1 To 5 Step 2
        If optBegin(i).Value = True Then  '¼�����뷴Ӧ
            intIndex = i \ 2
            If cboReaction(intIndex).ListIndex = -1 Then
                MsgBox "����Ѫ��Ӧʱ������¼�벻����Ӧ�����", vbInformation, gstrSysName
                Call gobjControl.ControlSetFocus(cboReaction(intIndex))
                Exit Sub
            End If
            If Not IsDate(txt��Ӧʱ��(intIndex).Text) Then
                MsgBox "��Ѫ��Ӧʱ�䲻����Ч�����ڸ�ʽ��", vbInformation, gstrSysName
                Call gobjControl.ControlSetFocus(txt��Ӧʱ��(intIndex))
                Exit Sub
            End If
            
            If CDate(Format(txt��Ӧʱ��(intIndex).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")) Then
                MsgBox "��Ѫ��Ӧʱ�䲻��С����ע��ʼʱ�䣡", vbInformation, gstrSysName
                Call gobjControl.ControlSetFocus(txt��Ӧʱ��(intIndex))
                Exit Sub
            End If
            '��ע�����е���Ѫ��Ӧʱ�����С����ע����ʱ��
            If intIndex = 2 And IsDate(txtִ��ʱ��(1).Text) Then
                If CDate(Format(txt��Ӧʱ��(intIndex).Text, "YYYY-MM-DD HH:mm")) >= CDate(Format(txtִ��ʱ��(1).Text, "YYYY-MM-DD HH:mm")) Then
                    MsgBox "��ע�����е���Ѫ��Ӧʱ�����С����ע����ʱ�䣡", vbInformation, gstrSysName
                    Call gobjControl.ControlSetFocus(txt��Ӧʱ��(intIndex))
                    Exit Sub
                End If
            End If
            '��ע��������Ѫ��Ӧʱ�������ڵ�����ע����ʱ��
            If intIndex = 1 And IsDate(txtִ��ʱ��(1).Text) Then
                If CDate(Format(txt��Ӧʱ��(intIndex).Text, "YYYY-MM-DD HH:mm")) < CDate(Format(txtִ��ʱ��(1).Text, "YYYY-MM-DD HH:mm")) Then
                    MsgBox "��ע��������Ѫ��Ӧʱ�䲻��С����ע����ʱ�䣡", vbInformation, gstrSysName
                    Call gobjControl.ControlSetFocus(txt��Ӧʱ��(intIndex))
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    If optBegin(1).Value = True And optBegin(5).Value = True Then
        If CDate(Format(txt��Ӧʱ��(2).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(txt��Ӧʱ��(0).Text, "YYYY-MM-DD HH:mm")) Then
            MsgBox "��ע�����е���Ѫ��Ӧʱ����������ע��ʼ����Ѫ��Ӧʱ�䣡", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt��Ӧʱ��(2))
            Exit Sub
        End If
    End If
    
    If optBegin(1).Value = True And optBegin(3).Value = True Then
        If CDate(Format(txt��Ӧʱ��(1).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(txt��Ӧʱ��(0).Text, "YYYY-MM-DD HH:mm")) Then
            MsgBox "��ע��������Ѫ��Ӧʱ����������ע��ʼ����Ѫ��Ӧʱ�䣡", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt��Ӧʱ��(1))
            Exit Sub
        End If
    End If
    
    If optBegin(3).Value = True And optBegin(5).Value = True Then
        If CDate(Format(txt��Ӧʱ��(1).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(txt��Ӧʱ��(2).Text, "YYYY-MM-DD HH:mm")) Then
            MsgBox "��ע��������Ѫ��Ӧʱ����������ע�����е���Ѫ��Ӧʱ�䣡", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt��Ӧʱ��(1))
            Exit Sub
        End If
    End If
    
    If IsDate(txtִ��ʱ��(1).Text) Then
         If CDate(Format(txtִ��ʱ��(1).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")) Then
            MsgBox "��ע����ʱ����������ע��ʼʱ�䣡", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txtִ��ʱ��(1))
            Exit Sub
         End If
    End If
    
    If gobjCommFun.ActualLen(txtִ��ժҪ.Text) > txtִ��ժҪ.MaxLength Then
        MsgBox "ִ��ժҪ���ݹ��࣬������� " & txtִ��ժҪ.MaxLength \ 2 & " �����ֻ� " & txtִ��ժҪ.MaxLength & " ���ַ���", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txtִ��ժҪ)
        Exit Sub
    End If
    dbl�������� = Val(txt��������.Text)
    dblʣ����� = gobjComlib.FormatEx(Val(txt��������.Text) - Val(txt��������.Tag), 5)
    If mintѪ���� > mint��ִ��Ѫ���� Then
        dbl�������� = gobjComlib.FormatEx(dblʣ����� / (mintѪ���� - mint��ִ��Ѫ����), 5)
    Else
        dbl�������� = gobjComlib.FormatEx(dbl�������� / mintѪ����, 5)
    End If
    If Val(txt��������.Tag) + dbl�������� > Val(txt��������.Text) Then
        dbl�������� = gobjComlib.FormatEx(Val(txt��������.Text) - Val(txt��������.Tag), 5)
    End If
    
    '���ݸ�ֵ
    With mExeInfoSave
        .�շ�ID = mlng�շ�ID
        .��ʼִ���� = Trim(gobjCommFun.GetNeedName(cboִ����(0).Text))
        .��ʼʱ�� = Format(str��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
        If cbo����(0).ListIndex > 0 Then
            .ǰ15���ӵ��� = cbo����(0).ItemData(cbo����(0).ListIndex)
        Else
            .ǰ15���ӵ��� = Val(cbo����(0).Text)
        End If
        .��עǰ��Ѫ��Ӧ = IIf(optBegin(0).Value = True, "", cboReaction(0).Text)
        .��עǰ��Ӧʱ�� = IIf(optBegin(0).Value = True, "", Format(txt��Ӧʱ��(0).Text, "yyyy-MM-dd HH:mm:ss"))
        .��ע��ִ���� = Trim(gobjCommFun.GetNeedName(cboִ����(2).Text))
        If .��ע��ִ���� = "" Then
            .��15���ӵ��� = 0
            .��ע����Ѫ��Ӧ = ""
            .��ע�з�Ӧʱ�� = ""
            str��Ѫ��λ������© = "0"
        Else
            If cbo����(1).ListIndex > 0 Then
                .��15���ӵ��� = cbo����(1).ItemData(cbo����(1).ListIndex)
            Else
                .��15���ӵ��� = Val(cbo����(1).Text)
            End If
            .��ע����Ѫ��Ӧ = IIf(optBegin(4).Value = True, "", cboReaction(2).Text)
            .��ע�з�Ӧʱ�� = IIf(optBegin(4).Value = True, "", Format(txt��Ӧʱ��(2).Text, "yyyy-MM-dd HH:mm:ss"))
            str��Ѫ��λ������© = IIf(optLeakage(0).Value = True, 0, 1)
        End If
        If Not IsDate(txtִ��ʱ��(1).Text) Then
            .��ע����Ѫ��Ӧ = ""
            .��ע��Ӧʱ�� = ""
            .����ִ���� = ""
            .����ʱ�� = ""
        Else
            .��ע����Ѫ��Ӧ = IIf(optBegin(2).Value = True, "", cboReaction(1).Text)
            .��ע��Ӧʱ�� = IIf(optBegin(2).Value = True, "", Format(txt��Ӧʱ��(1).Text, "yyyy-MM-dd HH:mm:ss"))
            .����ִ���� = Trim(gobjCommFun.GetNeedName(cboִ����(1).Text))
            .����ʱ�� = Format(txtִ��ʱ��(1).Text, "yyyy-MM-dd HH:mm:ss")
            str��Ѫ��λ������© = str��Ѫ��λ������© & IIf(optLeakage(2).Value = True, 0, 1)
        End If
        If .��ע��ִ���� = "" And Not IsDate(txtִ��ʱ��(1).Text) Then
            .��Ѫ��λ��© = ""
        Else
            .��Ѫ��λ��© = str��Ѫ��λ������©
        End If
        
        .�˲��� = txtCheck(0).Text
        .������ = txtCheck(1).Text
        .�˶�ʱ�� = txtCheck(2).Text
        .ִ�п���ID = mlng����ID
        .�Ǽ��� = UserInfo.����
        .�Ǽ�ʱ�� = ""
        .ժҪ = txtִ��ժҪ.Text
    End With
    
    Call SetMessages(arrMsg)
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    If mExeInfo.��ʼʱ�� = "" Then
'        If mlngִ�п���ID <> mlng����ID Then
'            strsql = "Zl_����ҽ������_���ұ��(" & mlngҽ��ID & "," & mlng���ͺ� & "," & mlng����ID & ")"
'            Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
'        End If
        
        strsql = "ZL_����ҽ��ִ��_Insert(" & mlngҽ��ID & "," & mlng���ͺ� & "," & _
            "To_Date('" & Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl�������� & ",'" & txtִ��ժҪ.Text & "','" & gobjCommFun.GetNeedName(cboִ����(0).Text) & "'," & _
            "To_Date('" & Format(str��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            1 & "," & "0," & 1 & ",'','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    Else
        strsql = "ZL_����ҽ��ִ��_Update(To_Date('" & mExeInfo.��ʼʱ�� & "','YYYY-MM-DD HH24:MI:SS')," & mlngҽ��ID & "," & mlng���ͺ� & "," & _
            "To_Date('" & Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl�������� & ",'" & txtִ��ժҪ.Text & "','" & gobjCommFun.GetNeedName(cboִ����(0).Text) & "'," & _
            "To_Date('" & Format(str��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & "," & 1 & ",NULL," & 1 & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    End If
    
     '����ʱ�ĺ˶�
    If lblLink.Enabled = True Then
        strsql = "Zl_ѪҺִ�м�¼_Check(" & mlng�շ�ID & ",'" & mExeInfoSave.�˲��� & "','" & mExeInfoSave.������ & "',To_Date('" & mExeInfoSave.�˶�ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),'" & imgMore.Tag & "')"
        Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    End If
    
    strsql = "zl_ѪҺִ�м�¼_Update(" & mlng�շ�ID & ",'" & mExeInfoSave.��ʼִ���� & "',To_Date('" & mExeInfoSave.��ʼʱ�� & "','YYYY-MM-DD HH24:MI:SS')," & _
        IIf(mExeInfoSave.ǰ15���ӵ��� = 0, "NULL", mExeInfoSave.ǰ15���ӵ���) & ",'" & mExeInfoSave.��עǰ��Ѫ��Ӧ & "'," & IIf(mExeInfoSave.��עǰ��Ӧʱ�� = "", "NULL", "To_Date('" & mExeInfoSave.��עǰ��Ӧʱ�� & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & mExeInfoSave.��ע��ִ���� & "'," & _
        IIf(mExeInfoSave.��15���ӵ��� = 0, "NULL", mExeInfoSave.��15���ӵ���) & ",'" & mExeInfoSave.��ע����Ѫ��Ӧ & "'," & IIf(mExeInfoSave.��ע�з�Ӧʱ�� = "", "NULL", "To_Date('" & mExeInfoSave.��ע�з�Ӧʱ�� & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & _
        mExeInfoSave.����ִ���� & "'," & IIf(mExeInfoSave.����ʱ�� = "", "NULL", "To_Date('" & mExeInfoSave.����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')") & "," & _
        "'" & mExeInfoSave.��ע����Ѫ��Ӧ & "'," & IIf(mExeInfoSave.��ע��Ӧʱ�� = "", "NULL", "To_Date('" & mExeInfoSave.��ע��Ӧʱ�� & "','YYYY-MM-DD HH24:MI:SS')") & "," & _
        mExeInfoSave.ִ�п���ID & ",'" & mExeInfoSave.��Ѫ��λ��© & "','" & mExeInfoSave.�Ǽ��� & "'," & IIf(mExeInfoSave.�Ǽ�ʱ�� = "", "NULL", "To_Date('" & mExeInfoSave.�Ǽ�ʱ�� & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & mExeInfoSave.ժҪ & "')"
    Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    
    '����������������
    For i = 0 To 2
        strsql = UCPatiVS(i).GetSaveSQL(mlng�շ�ID, i + 1)
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
    Dim lng����ID As Long, lng����ID As Long, lng����ID As Long, lngִ�п���id As Long
    Dim lng����id As Long
    Dim int������Դ As Integer
    Dim str���Ѳ��� As String
    arrSQL = Array()
    strSQL = "select a.��ҳid,a.�Һŵ�,a.����id,a.���˿���id,a.������Դ, b.ִ�в���id from ����ҽ����¼ a,ѪҺ��Ѫ��¼ b,ѪҺ���ͼ�¼ c " & _
            "WHERE a.id = b.����id AND b.id = c.�䷢id AND a.id = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng���ID)
    If rsTmp.State = adStateClosed Then Exit Function
    If rsTmp.RecordCount = 0 Then Exit Function
    lng����ID = Val(rsTmp!����id)
    lng����ID = Val(rsTmp!���˿���id)
    int������Դ = Val(rsTmp!������Դ)
    lngִ�п���id = Val(rsTmp!ִ�в���id)
    If int������Դ = 2 Then
        lng����id = Val(rsTmp!��ҳid)
        strSQL = "select ��ǰ����id from ������ҳ where ����id = [1] and ��ҳid = [2]  "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����id)
        lng����ID = Val(rsTmp!��ǰ����id)
    Else
        lng����ID = Val(rsTmp!���˿���id)
        strSQL = "select id �Һ�id from ���˹Һż�¼ where no = [1] and ����id = [2] "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTmp!�Һŵ� & "", Val(rsTmp!����id))
        lng����id = Val(rsTmp!�Һ�ID)
    End If
    strSQL = "select ID,���ͱ���,ҵ���ʶ from ҵ����Ϣ�嵥 where ����ID = [1] and ����id = [2] and �Ƿ����� = 0 "
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����id)
    'ȷ��Ҫ���ѵĲ���
    str���Ѳ��� = IIf(Val(lng����ID) = 0, "", lng����ID)
    If lng����ID <> lng����ID Then
        If str���Ѳ��� = "" Then
            str���Ѳ��� = IIf(lng����ID = 0, "", lng����ID)
        Else
            str���Ѳ��� = str���Ѳ��� & IIf(lng����ID = 0, "", "," & lng����ID)
        End If
    End If
    '��ѯ�Ƿ���ڱ�ҽ����Ѫ����Ѫ��Ӧ��Ϣ
    rsTmp.Filter = "���ͱ��� = 'ZLHIS_BLOOD_006' And ҵ���ʶ = '" & mlng���ID & ":" & mlng�շ�ID & "'"
    If (optBegin(1).Value Or optBegin(3).Value Or optBegin(5).Value) Then
        If rsTmp.RecordCount = 0 Then
            strSQL = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng����id & ","  '����id ����id
            strSQL = strSQL & Val(lng����ID) & ","     '�������id
            strSQL = strSQL & Val(lng����ID) & ","      '���ﲡ��id
            strSQL = strSQL & int������Դ & ","                                      '������Դ
            strSQL = strSQL & "'������Ѫ��Ӧ���뼰ʱ��д��Ѫ��Ӧ����','"             '��Ϣ����
            strSQL = strSQL & IIf(Val(int������Դ) = 1, "1000", "0100") & "','ZLHIS_BLOOD_006',"     ' ���ѳ���, ���ͱ���
            strSQL = strSQL & "'" & mlng���ID & ":" & mlng�շ�ID & "',"                      'ҵ���ʶ�����id:�շ�id��
            strSQL = strSQL & "1,0,NULL,'" & str���Ѳ��� & "',NULL)"                                                   '���ȳ̶ȣ��Ƿ����ģ��Ǽ�ʱ��,���Ѳ���
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    Else    '����Ѫ��Ӧ�����ѯ�Ƿ�����з�Ӧ��Ϣ�����У���Ϊ�Ѷ���
        If rsTmp.RecordCount > 0 Then
            strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_006',"
            strSQL = strSQL & "3,'" & UserInfo.���� & "'," & lng����ID & ",NULL,"
            strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    End If
    
    rsTmp.Filter = "���ͱ��� = 'ZLHIS_BLOOD_007' And ҵ���ʶ = '" & mlng���ID & ":" & mlngҽ��ID & ":" & mlng�շ�ID & "'"
    If IsDate(txtִ��ʱ��(1).Text) Then
        If rsTmp.RecordCount = 0 Then
            strSQL = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng����id & ","  '����id ����id
            strSQL = strSQL & Val(lng����ID) & ","      '�������id
            strSQL = strSQL & Val(lng����ID) & ","      '���ﲡ��id
            strSQL = strSQL & int������Դ & ","                                      '������Դ
            strSQL = strSQL & "'��Ѫ��ɣ�����24Сʱ���ջ�Ѫ����','"                         '��Ϣ����
            strSQL = strSQL & IIf(Val(int������Դ) = 1, "0001", "0010") & "','ZLHIS_BLOOD_007',"     ' ���ѳ���, ���ͱ���
            strSQL = strSQL & "'" & mlng���ID & ":" & mlngҽ��ID & ":" & mlng�շ�ID & "',"                      'ҵ���ʶ�����id:�շ�id��
            strSQL = strSQL & "1,0,NULL,'" & IIf(Val(int������Դ) = 1, lngִ�п���ID, lng����ID)  & "',NULL)"                                                   '���ȳ̶ȣ��Ƿ����ģ��Ǽ�ʱ��,���Ѳ���                                                      '
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    Else    '��Ѫ����Ҫ���գ����ѯ�Ƿ����Ѫ����Ϣ�����У���Ϊ�Ѷ���
        If rsTmp.RecordCount > 0 Then
            strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_007',"
            strSQL = strSQL & IIf(Val(int������Դ) = 1, 4, 3) & ",'" & UserInfo.���� & "'," & lng����ID & ",NULL,"
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
        Set objControl = txtִ��ʱ��(0)
    ElseIf intIndex = 1 Then
        Set objControl = txtִ��ʱ��(1)
    ElseIf intIndex = 2 Then
        Set objControl = txt��Ӧʱ��(0)
    ElseIf intIndex = 3 Then
        Set objControl = txt��Ӧʱ��(1)
    ElseIf intIndex = 4 Then
        Set objControl = txt��Ӧʱ��(2)
    Else
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Exit Sub
    End If
    
    'ȡֵ
    If IsDate(objControl.Text) Then
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(objControl.Text, "yyyy-MM-dd HH:mm"), 12, 5)
    Else
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
    End If
    objControl.Text = strDate
        
    If intIndex = 0 Or intIndex = 1 Then
        Call txtִ��ʱ��_Validate(intIndex, False)
    ElseIf intIndex = 2 Or intIndex = 3 Or intIndex = 4 Then
        Set objControl = txtִ��ʱ��(1)
        Call txt��Ӧʱ��_Validate(intIndex - 2, False)
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
            Set objControl = txtִ��ʱ��(0)
        ElseIf intIndex = 1 Then
            Set objControl = txtִ��ʱ��(1)
        ElseIf intIndex = 2 Then
            Set objControl = txt��Ӧʱ��(0)
        ElseIf intIndex = 3 Then
            Set objControl = txt��Ӧʱ��(1)
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
    '��ʾ����
    lngScrH = GetSystemMetrics(17) * 15 '��Ļ���ø߶�
    If Me.Top + Me.Height > lngScrH Then
        Me.Top = lngScrH - Me.Height
    End If
    If mblnOnlyRead = True Then
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.name <> "cboִ����" And Me.ActiveControl.name <> "UCPatiVS" Then
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
    Dim lng������� As Long
    
    On Error GoTo ErrHand
    mintTimerCount = 0
    mblnAcTive = True
    Call ClearExeInfo
    strsql = "Select ����ʱ��,����״̬ From ѪҺ���ͼ�¼ where �շ�ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng�շ�ID)
    If rsTmp.EOF Then
        MsgBox "ѪҺ��δ����,���ܽ���ִ�еǼǣ�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If gbln���պ����ִ�� = True And mblnOnlyRead = False Then
        If Not (Val("" & rsTmp!����״̬) = 1 Or Val("" & rsTmp!����״̬) = 3) Then
            MsgBox "ѪҺ��δ����,���ܽ���ִ�еǼǣ�", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    If IsDate("" & rsTmp!����ʱ��) Then
        mstr����ʱ�� = Format("" & rsTmp!����ʱ��, "YYYY-MM-DD HH:mm")
    Else
        mstr����ʱ�� = ""
    End If
    strsql = _
        " Select �շ�id, ��ʼִ����, ��ʼʱ��, ǰ15���ӵ���, ��עǰ��Ѫ��Ӧ, ��עǰ��Ӧʱ��, ��15���ӵ���,��ע��ִ����,��ע����Ѫ��Ӧ, ��ע�з�Ӧʱ��,  " & vbNewLine & _
                "����ִ����, ����ʱ��, ִ�п���id, ��ע����Ѫ��Ӧ, ��ע��Ӧʱ��,��Ѫ��λ��©,�˲���, ������, �˶�ʱ��, �Ǽ���,�Ǽ�ʱ��, ժҪ" & vbNewLine & _
        " From ѪҺִ�м�¼ where �շ�ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng�շ�ID)
    If rsTmp.RecordCount > 0 Then
        With mExeInfo
            .�շ�ID = Val("" & rsTmp("�շ�ID"))
            .��ʼִ���� = "" & rsTmp("��ʼִ����")
            .��ʼʱ�� = "" & rsTmp("��ʼʱ��")
            .ǰ15���ӵ��� = Val("" & rsTmp("ǰ15���ӵ���"))
            .��עǰ��Ѫ��Ӧ = "" & rsTmp("��עǰ��Ѫ��Ӧ")
            .��עǰ��Ӧʱ�� = "" & rsTmp("��עǰ��Ӧʱ��")
            .��ע��ִ���� = "" & rsTmp("��ע��ִ����")
            .��15���ӵ��� = Val("" & rsTmp("��15���ӵ���"))
            .��ע����Ѫ��Ӧ = "" & rsTmp("��ע����Ѫ��Ӧ")
            .��ע�з�Ӧʱ�� = "" & rsTmp("��ע�з�Ӧʱ��")
            .����ִ���� = "" & rsTmp("����ִ����")
            .����ʱ�� = "" & rsTmp("����ʱ��")
            .��ע����Ѫ��Ӧ = "" & rsTmp("��ע����Ѫ��Ӧ")
            .��ע��Ӧʱ�� = "" & rsTmp("��ע��Ӧʱ��")
            .ִ�п���ID = Val("" & rsTmp("ִ�п���ID"))
            .�˲��� = "" & rsTmp("�˲���")
            .������ = "" & rsTmp("������")
            .�˶�ʱ�� = "" & rsTmp("�˶�ʱ��")
            .��Ѫ��λ��© = "" & rsTmp("��Ѫ��λ��©")
            .�Ǽ��� = "" & rsTmp("�Ǽ���")
            .�Ǽ�ʱ�� = "" & rsTmp("�Ǽ�ʱ��")
            .ժҪ = "" & rsTmp("ժҪ")
        End With
    End If
    mstrȱʡ��Ѫ��Ӧ = ""
    
    '����
    With cbo����(0)
        .Clear
        .AddItem 15: .ItemData(.NewIndex) = 15
        .AddItem 30: .ItemData(.NewIndex) = 30
        .AddItem "����": .ItemData(.NewIndex) = -1
        .AddItem "��ѹ": .ItemData(.NewIndex) = -2
        .ListIndex = 0
    End With
    With cbo����(1)
        .Clear
        .AddItem 15: .ItemData(.NewIndex) = 15
        .AddItem 30: .ItemData(.NewIndex) = 30
        .AddItem "����": .ItemData(.NewIndex) = -1
        .AddItem "��ѹ": .ItemData(.NewIndex) = -2
    End With
    '��Ѫ��Ӧ��ȡ
    strsql = "Select ����,ȱʡ��־ From ��Ѫ��Ӧ"
    Call gobjDatabase.OpenRecordset(rsTmp, strsql, "��Ѫ��Ӧ")
    cboReaction(0).Clear
    cboReaction(1).Clear
    cboReaction(2).Clear
    Do While Not rsTmp.EOF
        cboReaction(0).AddItem "" & rsTmp!����
        cboReaction(1).AddItem "" & rsTmp!����
        cboReaction(2).AddItem "" & rsTmp!����
        If Val(rsTmp!ȱʡ��־) = 1 Then mstrȱʡ��Ѫ��Ӧ = "" & rsTmp!����
    rsTmp.MoveNext
    Loop
    'ִ���˶�ȡ
    cboִ����(0).Clear: cboִ����(0).Tag = -1
    cboִ����(1).Clear: cboִ����(1).Tag = -1
    cboִ����(2).Clear: cboִ����(2).Tag = -1
    cboִ����(2).AddItem ""
    cboִ����(2).ItemData(cboִ����(2).NewIndex) = 0
    '��ȡִ����(������Ա)
    Set rsTmp = GetDataToPersons
    For i = 1 To rsTmp.RecordCount
        cboִ����(0).AddItem rsTmp!��� & "-" & rsTmp!����
        cboִ����(0).ItemData(cboִ����(0).NewIndex) = Val("" & rsTmp!id)
        cboִ����(1).AddItem rsTmp!��� & "-" & rsTmp!����
        cboִ����(1).ItemData(cboִ����(1).NewIndex) = Val("" & rsTmp!id)
        cboִ����(2).AddItem rsTmp!��� & "-" & rsTmp!����
        cboִ����(2).ItemData(cboִ����(2).NewIndex) = Val("" & rsTmp!id)
        If mExeInfo.��ʼʱ�� = "" Then
            If rsTmp!id = UserInfo.id Then
                cboִ����(0).ListIndex = cboִ����(0).NewIndex
            End If
        Else
            If rsTmp!���� = mExeInfo.��ʼִ���� Then
                cboִ����(0).ListIndex = cboִ����(0).NewIndex
            End If
        End If
        If mExeInfo.��ע��ִ���� = "" Then
            If rsTmp!id = UserInfo.id Then
                lngFindCenterIndex = cboִ����(2).NewIndex
            End If
        Else
            If rsTmp!���� = mExeInfo.��ע��ִ���� Then
                cboִ����(2).ListIndex = cboִ����(2).NewIndex
            End If
        End If
        
        If mExeInfo.����ʱ�� = "" Then
            If rsTmp!id = UserInfo.id Then
                lngFindEndIndex = cboִ����(1).NewIndex
            End If
        Else
            If rsTmp!���� = mExeInfo.����ִ���� Then
                cboִ����(1).ListIndex = cboִ����(1).NewIndex
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If mlngModul = pҽ������վ Then
        If Val(gobjDatabase.GetPara(51, 100)) = 1 Then
            Me.cboִ����(0).Enabled = False
            If cboִ����(1).ListCount > 0 And cboִ����(1).ListIndex = -1 Then cboִ����(1).ListIndex = lngFindEndIndex
            Me.cboִ����(1).Enabled = False
            If cboִ����(2).ListCount > 0 And cboִ����(2).ListIndex = -1 Then cboִ����(2).ListIndex = lngFindCenterIndex
            Me.cboִ����(2).Enabled = False
        End If
    End If
    
    If mExeInfo.�˶�ʱ�� <> "" Then
        txtCheck(0).Text = mExeInfo.�˲���
        txtCheck(1).Text = mExeInfo.������
        txtCheck(2).Text = Format(mExeInfo.�˶�ʱ��, "YYYY-MM-DD HH:mm")
        lblLink.Enabled = False
        lblLink.Tag = "�Ѻ˶�"
    Else
        txtCheck(0).Text = ""
        txtCheck(1).Text = ""
        txtCheck(2).Text = ""
        lblLink.Enabled = True
        lblLink.Tag = ""
    End If
    'ִ�������ȡ
    mintѪ���� = 0
    mint��ִ��Ѫ���� = 0
    txt��������.Tag = ""
    txtִ��ʱ��(0).Text = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    With mExeInfo
        If mExeInfo.��ʼʱ�� = "" Then '����
            '��ȡ�ϴ�ִ����Ϣ
            strsql = _
                " Select Max(ִ��ʱ��) as LastDate,Sum(��������) as curNum" & _
                " From ����ҽ��ִ��" & _
                " Where ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlngҽ��ID, mlng���ͺ�)
            If Not rsTmp.EOF Then
                txtִ��ʱ��(0).Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss") '�ϴ�ʵ��ִ��ʱ��
                txt��������.Tag = Nvl(rsTmp!curNum, 0) 'ѪҺҽ����ִ�д����ܺͣ�ÿ��ִ��һ��Ѫ��ִ�д���Ϊ1
            End If
            
            '���㱾��ִ��Ӧ�õ�Ҫ��ʱ��
            strsql = "Select A.��������,Nvl(B.���id, B.ID) ��ID,C.���㵥λ,A.�״�ʱ��,A.ĩ��ʱ��,Decode(B.������Դ, 2, Decode(A.��¼����, 1, 1, Decode(A.�������, 1, 1, 2)), 1) ��������," & _
                " B.��ʼִ��ʱ��,Decode(B.ҽ����Ч,0,B.ִ����ֹʱ��,null) as ִ����ֹʱ��,B.�ϴ�ִ��ʱ��,B.ִ��ʱ�䷽��," & _
                " B.ִ��Ƶ��,B.Ƶ�ʴ���,B.Ƶ�ʼ��,B.�����λ,B.����ID,b    .��ҳID,c.���,c.��������,c.ִ�з���,C.���㷽ʽ,B.ҽ����Ч,Nvl(b.�ܸ�����, 1) as �ܸ�����,NVL(B.��������,1) AS ��������,A.NO " & _
                " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C" & _
                " Where A.ҽ��ID=B.ID And B.������ĿID=C.ID" & _
                " And A.ҽ��ID=[1] And A.���ͺ�=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlngҽ��ID, mlng���ͺ�)
            dtpҪ��ʱ��.Value = rsTmp!��ʼִ��ʱ�� '��Ѫҽ����Ϊһ����ִ�е�����
            txt��������.Text = Val(rsTmp!�������� & "")
            mlng���ID = rsTmp!��ID
            '��ѯ�����������Ѿ��˷ѻ����ʵ�ҽ��ִ�д���
            lng������� = Get�������(mlngҽ��ID, rsTmp!NO & "", rsTmp!��� & "", Val(rsTmp!�������� & ""))
            mintѪ���� = GetBloodNum
            mint��ִ��Ѫ���� = gobjComlib.FormatEx(Val(txt��������.Tag) * mintѪ���� / Val(txt��������.Text), 0) '�ϴ�ִ�е�Ѫ�������Ѿ���5λС��������������Ϳ�������
            If lng������� = 1 Then '������Ѫҽ�� ��� mlng������� ��ֵֻ���� 0 �� 1 ��Ϊֻ����һ�Ρ�
                MsgBox "��ҽ����ص����Ѿ��˷ѻ����ʣ�������ִ�С�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            ElseIf lng������� = 0 Then
                If mint��ִ��Ѫ���� >= mintѪ���� Then
                    MsgBox "��ҽ�����η�������ִ�� " & mintѪ���� & "������ǰ�Ѿ�ִ���� " & mint��ִ��Ѫ���� & " ����������ִ�С�", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
            txt��������.Text = 1 'ÿ��ִ��Ĭ��Ϊһ��
        Else '�޸�
            '�ϴ���ִ�е���һЩ����(���㱾��)
            strsql = "Select " & _
                " Max(ִ��ʱ��) as LastDate," & _
                " Sum(��������) as curNum" & _
                " From ����ҽ��ִ��" & _
                " Where ִ��ʱ��<[3] And ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(mExeInfo.��ʼʱ��))
            If Not rsTmp.EOF Then
                txt��������.Tag = Nvl(rsTmp!curNum, 0) '�ϴ�Ϊֹʵ����ִ�е���������
                txtִ��ʱ��(0).Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss")   '�ϴ�ʵ��ִ��ʱ��
            End If
            
            strsql = "Select A.Ҫ��ʱ��,Nvl(C.���id, C.ID) ��ID,A.ִ��ʱ��,A.��������,A.ִ��ժҪ,A.ִ�н��,A.ִ����,B.��������,Decode(C.������Դ, 2, Decode(B.��¼����, 1, 1, Decode(B.�������, 1, 1, 2)), 1) ��������,D.���㵥λ,Decode(c.ҽ����Ч,0,c.ִ����ֹʱ��,null) as ִ����ֹʱ�� ,d.���,d.��������,d.ִ�з���,c.����ID,c.��ҳID,B.NO" & _
                " From ����ҽ��ִ�� A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D" & _
                " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID And C.������ĿID=D.ID" & _
                " And A.ҽ��ID=[1] And A.���ͺ�=[2] And A.ִ��ʱ��=[3]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(mExeInfo.��ʼʱ��))
            dtpҪ��ʱ��.Value = rsTmp!Ҫ��ʱ��
            txt��������.Text = gobjComlib.FormatEx(Nvl(rsTmp!��������), 5)
            txtִ��ժҪ.Text = "" & rsTmp!ִ��ժҪ
            txtִ��ʱ��(0).Text = Format(rsTmp!ִ��ʱ��, "YYYY-MM-DD HH:mm")
            txt��������.Text = Val(rsTmp!�������� & "")
            mlng���ID = rsTmp!��ID
            gobjComlib.cbo.SetText cboִ����(0), rsTmp!ִ����
            
            mintѪ���� = GetBloodNum
            mint��ִ��Ѫ���� = gobjComlib.FormatEx(Val(txt��������.Tag) * mintѪ���� / Val(txt��������.Text), 0) '���ε�ִ�е�Ѫ����
            txt��������.Text = gobjComlib.FormatEx(Val("" & rsTmp!��������) * mintѪ����, 0)
            
            Select Case mExeInfo.ǰ15���ӵ���
                Case 0
                    cbo����(0).Text = ""
                Case -1
                    cbo����(0).ListIndex = 2
                Case -2
                    cbo����(0).ListIndex = 3
                Case Else
                    cbo����(0).Text = mExeInfo.ǰ15���ӵ���
            End Select
            If mExeInfo.��עǰ��Ѫ��Ӧ <> "" Then
                optBegin(1).Value = True
                Call gobjControl.CboLocate(cboReaction(0), mExeInfo.��עǰ��Ѫ��Ӧ)
                txt��Ӧʱ��(0).Text = Format(mExeInfo.��עǰ��Ӧʱ��, "YYYY-MM-DD HH:mm")
            End If
            
            Select Case mExeInfo.��15���ӵ���
                Case 0
                    cbo����(1).Text = ""
                Case -1
                    cbo����(1).ListIndex = 2
                Case -2
                    cbo����(1).ListIndex = 3
                Case Else
                    cbo����(1).Text = mExeInfo.��15���ӵ���
            End Select
            gobjComlib.cbo.SetText cboִ����(2), mExeInfo.��ע��ִ����
            If Val(Mid(.��Ѫ��λ��©, 1, 1)) = 1 Then optLeakage(1).Value = True
            If mExeInfo.��ע����Ѫ��Ӧ <> "" Then
                optBegin(5).Value = True
                Call gobjControl.CboLocate(cboReaction(2), mExeInfo.��ע����Ѫ��Ӧ)
                txt��Ӧʱ��(2).Text = Format(mExeInfo.��ע�з�Ӧʱ��, "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(Format(mExeInfo.����ʱ��, "YYYY-MM-DD HH:mm")) Then
                txtִ��ʱ��(1).Text = Format(mExeInfo.����ʱ��, "YYYY-MM-DD HH:mm")
                gobjComlib.cbo.SetText cboִ����(1), mExeInfo.����ִ����
                If mExeInfo.��ע����Ѫ��Ӧ <> "" Then
                    optBegin(3).Value = True
                    Call gobjControl.CboLocate(cboReaction(1), mExeInfo.��ע����Ѫ��Ӧ)
                    txt��Ӧʱ��(1).Text = Format(mExeInfo.��ע��Ӧʱ��, "YYYY-MM-DD HH:mm")
                End If
                If Val(Mid(.��Ѫ��λ��©, 2, 1)) = 1 Then optLeakage(3).Value = True
            End If
        End If
    End With
    cmdDate(0).Tag = txtִ��ʱ��(0).Text
    If cboִ����(0).ListIndex <> -1 Then cboִ����(0).Tag = cboִ����(0).ListIndex
    If cboִ����(1).ListIndex <> -1 Then cboִ����(1).Tag = cboִ����(1).ListIndex
    If cboִ����(2).ListIndex <> -1 Then cboִ����(2).Tag = cboִ����(2).ListIndex
    
    Call UCPatiVS(0).LoadPatiVitalSigns(mlng�շ�ID, 1)  '��Ѫǰ��������
    Call UCPatiVS(1).LoadPatiVitalSigns(mlng�շ�ID, 2)  '��Ѫ����������
    Call UCPatiVS(2).LoadPatiVitalSigns(mlng�շ�ID, 3)  '��Ѫ����������
    
    Call SetFaceEnabledFalse
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub SetFaceEnabledFalse()
'���ܣ����ÿؼ��Ŀ�����
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
'��ȡ����ҽ�����͵�����
    Dim rsTemp As New ADODB.Recordset
    Dim strsql As String
    On Error GoTo ErrHand
    strsql = "Select Count(�շ�id)  ���� From ѪҺ���ͼ�¼ a, ѪҺ��Ѫ��¼ b Where a.�䷢id = b.Id And b.����id = [1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng���ID)
    GetBloodNum = rsTemp!����
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
    If mExeInfo.�˶�ʱ�� <> "" Then
        MsgBox "�ô�ѪҺ�Ѿ��˶ԣ��������ٴκ˶ԣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    blnOk = frmUserCheck.ShowMe(Me, mlngModul, mlng����ID, mlng����ID, mstr����ʱ��, "", True, ִ�к˶�)
    If blnOk = True Then
        strCheckOper = frmUserCheck.SendAndTakeOper
        strCheckTime = frmUserCheck.SendTime
        strCheckResult = frmUserCheck.CheckResult
        
        txtCheck(0).Text = Split(strCheckOper, "'")(0)
        txtCheck(1).Text = Split(strCheckOper, "'")(1)
        txtCheck(2).Text = strCheckTime
        imgMore.Tag = strCheckResult
        lblLink.Tag = "�Ѻ˶�"
        If IsDate(txtִ��ʱ��(0).Text) Then
            If Format(txtִ��ʱ��(0).Text, "YYYY-MM-DD HH:mm") < Format(strCheckTime, "YYYY-MM-DD HH:mm") Then
                txtִ��ʱ��(0).Text = Format(strCheckTime, "YYYY-MM-DD HH:mm")
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
       If cboReaction(intIndex).ListIndex < 0 Then Call gobjControl.CboLocate(cboReaction(intIndex), mstrȱʡ��Ѫ��Ӧ)
    End If
    cboReaction(intIndex).Enabled = blnEable
    txt��Ӧʱ��(intIndex).Enabled = blnEable
    cmdDate(intIndex + 2).Enabled = blnEable
End Sub

Private Sub cbo����_GotFocus(Index As Integer)
    Call gobjControl.TxtSelAll(cbo����(Index))
End Sub

Private Sub cbo����_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    If Chr(KeyAscii) = 0 Then
        If cbo����(Index).Text = "" Or cbo����(Index).SelLength = Len(cbo����(Index).Text) Or cbo����(Index).SelStart = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If KeyAscii <> vbKeyReturn And KeyAscii <> 8 Then
        If cbo����(Index).SelLength = 0 And Len(cbo����(Index).Text) > 2 Then
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

Private Sub txt��Ӧʱ��_GotFocus(Index As Integer)
    Call gobjComlib.os.OpenImeByName
'    Call gobjControl.TxtSelAll(txt��Ӧʱ��(Index))
    txt��Ӧʱ��(Index).SelStart = 0
    txt��Ӧʱ��(Index).SelLength = Len(txt��Ӧʱ��(Index).Text)
End Sub

Private Sub txt��Ӧʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If Index = 0 Then
            Call cmdDate_Click(2)
        ElseIf Index = 1 Then
            Call cmdDate_Click(3)
        End If
    End If
End Sub

Private Sub txt��Ӧʱ��_Validate(Index As Integer, Cancel As Boolean)
    Dim intIndex As Integer
    If Not IsDate(txt��Ӧʱ��(Index).Text) Then
        If txt��Ӧʱ��(Index).Text <> "____-__-__ __:__" Then
            Cancel = True
            Call txt��Ӧʱ��_GotFocus(Index)
            Exit Sub
        Else
            If IsDate(cmdDate(Index + 2).Tag) Then
                txt��Ӧʱ��(Index).Text = Format(cmdDate(Index + 2).Tag, "YYYY-MM-DD HH:mm")
            End If
        End If
    Else
        txt��Ӧʱ��(Index).Text = Format(txt��Ӧʱ��(Index).Text, "YYYY-MM-DD HH:mm")
        cmdDate(Index + 2).Tag = txt��Ӧʱ��(Index).Text
    End If
End Sub

Private Sub txtִ��ʱ��_GotFocus(Index As Integer)
    Call gobjComlib.os.OpenImeByName
    txtִ��ʱ��(Index).SelStart = 0: txtִ��ʱ��(Index).SelLength = Len(txtִ��ʱ��(Index).Text)
End Sub

Private Sub txtִ��ʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If Index = 0 Then
            Call cmdDate_Click(0)
        ElseIf Index = 1 Then
            Call cmdDate_Click(1)
        End If
    End If
End Sub

Private Sub txtִ��ʱ��_Validate(Index As Integer, Cancel As Boolean)
    Dim strDate As String
    If Not IsDate(txtִ��ʱ��(Index).Text) Then
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
        If txtִ��ʱ��(Index).Text <> "____-__-__ __:__" Then
            Cancel = True
            If IsDate(strDate) Then txtִ��ʱ��(Index).Text = strDate
            Call txtִ��ʱ��_GotFocus(Index)
            Exit Sub
        Else
            If Index = 0 And IsDate(strDate) Then
                 If IsDate(strDate) Then txtִ��ʱ��(Index).Text = strDate
            End If
        End If
    Else
        txtִ��ʱ��(Index).Text = Format(txtִ��ʱ��(Index).Text, "YYYY-MM-DD HH:mm")
        If Index = 0 Then
            cmdDate(0).Tag = txtִ��ʱ��(Index).Text
        ElseIf Index = 1 Then
            cmdDate(1).Tag = txtִ��ʱ��(Index).Text
        End If
    End If
End Sub

Private Sub txtִ��ժҪ_GotFocus()
    Call gobjControl.TxtSelAll(txtִ��ժҪ)
End Sub

Private Sub txtִ��ժҪ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function Get�������(ByVal lngҽ��ID As Long, ByVal strNO As String, ByVal str������� As String, ByVal int�������� As Integer) As Long
'���ܣ���ȡĳ��ҽ������ĳ��ҽ������������ʵ�ҽ��ִ�д���
'       lngҽ��ID ����ҽ��ID
'       strNo  ����NO
'       str������� ��ҽ�����������
'       int�������� 1-������ã�2-סԺ����
    Dim rsTmp As ADODB.Recordset, strsql As String, strTable As String
    strTable = IIf(int�������� = 1, "������ü�¼", "סԺ���ü�¼")
    On Error GoTo errH
    
    strsql = "Select -1 * Sum(Nvl(a.����, 1) * a.���� / b.����) As ���������" & vbNewLine & _
            "From " & strTable & " A, ����ҽ���Ƽ� B" & vbNewLine & _
            "Where a.ҽ����� = [1] And A.NO=[3] And b.ҽ��id = a.ҽ����� And b.�շ�ϸĿid = a.�շ�ϸĿid And Nvl(B.��������,0)=0 And a.��¼״̬ = 2 And a.��¼���� in(1,2,11) And a.�۸񸸺� Is Null And" & vbNewLine & _
            "      a.�շ���� Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
            " (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1)"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, lngҽ��ID, str�������, strNO)
    If rsTmp.RecordCount <> 0 Then
        Get������� = Val(rsTmp!��������� & "")
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function GetDataToPersons(Optional ByVal strIn As String = "") As ADODB.Recordset
'������Ӧ���ҵ�ҽ����Ա��Ϣ
    Dim strsql As String, strNewSQL As String, strWhere As String
    Dim blnYn As Boolean
    
    
    On Error GoTo ErrHand
    If strIn <> "" Then blnYn = True
    
    'ҽ����վ��ֻ��ִ��������Ŀ���ܿ����������ҵ�ҽ�����ٴ���ʿվ���ܻ����ȫԱ���˵�Ȩ�ޣ���Ҫ���ϲ���Ա����(���ڲ����������ˣ�Ҫô�������Ǹò����ģ�Ҫô���Ǿ���ȫԱ����Ȩ��)
    If InStr(mstrPrivs, "ִ��������Ŀ") > 0 Or Not (mlngModul = pҽ������վ) Then
        strNewSQL = " Union " & vbNewLine & _
                            " Select " & UserInfo.id & " id,'" & UserInfo.��� & "' ���,'" & UserInfo.���� & "' ����,'" & UserInfo.���� & "' ���� From Dual "
    End If
        
    If Not mlngModul = pҽ������վ Then
        strWhere = "  Exists (Select 1 From ��Ա����˵�� Where ��Աid = a.Id And Instr(',ҽ��,��ʿ,', ',' || ��Ա���� || ',', 1) <> 0)"
    End If
    
    '��ǰ��¼����Ա������ʾ��ǰ��
    If strNewSQL = "" Then
        strsql = "Select a.Id, a.���, a.����, a.����" & vbNewLine & _
            " From ��Ա�� a, ������Ա b" & vbNewLine & _
            " Where a.Id = b.��Աid And b.����id = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And" & vbNewLine & _
            "      (a.վ�� = ' & gstrNodeNo & ' Or a.վ�� Is Null) " & vbNewLine & _
            IIf(blnYn, " And (A.��� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & vbNewLine & _
            IIf(strWhere = "", "", " And " & strWhere) & " Order by Decode(a.id," & IIf(blnYn = True, "[4]", "[2]") & ",0,1),a.���"
    Else
        strsql = "Select a.Id, a.���, a.����, a.����" & vbNewLine & _
            " From ��Ա�� a, ������Ա b" & vbNewLine & _
            " Where a.Id = b.��Աid And b.����id = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And" & vbNewLine & _
            "      (a.վ�� = ' & gstrNodeNo & ' Or a.վ�� Is Null)" & IIf(strWhere = "", "", " And " & strWhere) & vbNewLine & _
            strNewSQL
        If blnYn Then
            strsql = " Select a.Id, a.���, a.����, a.���� From (" & strsql & ") a" & vbNewLine & _
                " Where (A.��� Like [2] Or A.���� Like [3] Or A.���� Like [3])  Order by Decode(a.id,[4],0,1),a.���"
        Else
            strsql = " Select a.Id, a.���, a.����, a.���� From (" & strsql & ") a Order by Decode(a.id,[2],0,1),a.���"
        End If
    End If
    If blnYn = True Then
        Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng����ID, UCase(strIn) & "%", gstrLike & UCase(strIn) & "%", UserInfo.id)
    Else
        Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng����ID, UserInfo.id)
    End If
    
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

