VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frm���ŷ�ҩ����New 
   Caption         =   "ҩƷ���ŷ�ҩ"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10260
   DrawStyle       =   1  'Dash
   Icon            =   "frmҩƷ���ŷ�ҩnew.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmҩƷ���ŷ�ҩnew.frx":030A
   ScaleHeight     =   9015
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Visible         =   0   'False
   Begin VB.Timer TimerReturn 
      Interval        =   10000
      Left            =   7680
      Top             =   240
   End
   Begin VB.Frame fraColorStateSend 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      TabIndex        =   72
      Top             =   6240
      Visible         =   0   'False
      Width           =   6705
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1200
         Picture         =   "frmҩƷ���ŷ�ҩnew.frx":09F4
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   79
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00D7D7FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   5080
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   78
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   77
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBDBDB&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3240
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   76
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   75
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   0
         Picture         =   "frmҩƷ���ŷ�ҩnew.frx":7246
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   74
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   73
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "����ҩ��"
         Height          =   180
         Index           =   4
         Left            =   1460
         TabIndex        =   86
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "ȱҩ"
         Height          =   180
         Index           =   3
         Left            =   5500
         TabIndex        =   85
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   2
         Left            =   4440
         TabIndex        =   84
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "�ܷ�"
         Height          =   180
         Index           =   1
         Left            =   3600
         TabIndex        =   83
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ"
         Height          =   180
         Index           =   0
         Left            =   2790
         TabIndex        =   82
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "��ΣҩƷ"
         Height          =   180
         Index           =   5
         Left            =   260
         TabIndex        =   81
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   6
         Left            =   6420
         TabIndex        =   80
         Top             =   30
         Width           =   180
      End
   End
   Begin VB.PictureBox picCondition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   120
      ScaleHeight     =   8535
      ScaleWidth      =   3615
      TabIndex        =   12
      Top             =   0
      Width           =   3615
      Begin VB.PictureBox picConOther 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3612
         Left            =   30
         ScaleHeight     =   3615
         ScaleWidth      =   3375
         TabIndex        =   47
         Top             =   3600
         Width           =   3375
         Begin VB.CheckBox chkWithNotAudited 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�����������뵥��"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            TabIndex        =   88
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdҩƷ���� 
            Height          =   250
            Left            =   2985
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":77D0
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   840
            Width           =   270
         End
         Begin VB.CommandButton cmd��ҩ;�� 
            Height          =   250
            Left            =   2985
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":830A
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   480
            Width           =   270
         End
         Begin VB.TextBox txtҩƷ���� 
            Height          =   300
            Left            =   840
            TabIndex        =   65
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txt��ҩ;�� 
            Height          =   300
            Left            =   840
            TabIndex        =   64
            Top             =   480
            Width           =   2415
         End
         Begin VB.Frame fraLineH2 
            Height          =   50
            Left            =   0
            TabIndex        =   63
            Top             =   120
            Width           =   3525
         End
         Begin VB.OptionButton opt��Χ 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ҩ����"
            Height          =   225
            Index           =   2
            Left            =   840
            TabIndex        =   62
            Top             =   2160
            Width           =   1125
         End
         Begin VB.OptionButton opt��Χ 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ҩ����"
            Height          =   225
            Index           =   1
            Left            =   2040
            TabIndex        =   61
            Top             =   1920
            Width           =   1125
         End
         Begin VB.OptionButton opt��Χ 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��������"
            Height          =   225
            Index           =   0
            Left            =   840
            TabIndex        =   60
            Top             =   1920
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.ComboBox Cboҽ������ 
            Height          =   276
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2400
            Width           =   2415
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ӥ��ҩƷ"
            Height          =   180
            Index           =   1
            Left            =   2160
            TabIndex        =   58
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ҩƷ"
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   57
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkDanger 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ȡ��ΣҩƷ"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            TabIndex        =   56
            Top             =   3060
            Width           =   1695
         End
         Begin VB.CheckBox chkDangerType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "A��"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   3348
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkDangerType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "B��"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   54
            Top             =   3348
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkDangerType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "C��"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   53
            Top             =   3348
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ҩ"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   1440
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ҩ"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   1
            Left            =   1440
            TabIndex        =   51
            Top             =   1440
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����I��"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   50
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����II��"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   3
            Left            =   1440
            TabIndex        =   49
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkToxicologyType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ȡ�Ķ������"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            TabIndex        =   48
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblҩƷ���� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ҩƷ����"
            Height          =   180
            Left            =   0
            TabIndex        =   69
            Top             =   900
            Width           =   720
         End
         Begin VB.Label lbl��ҩ;�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ҩ;��"
            Height          =   180
            Left            =   0
            TabIndex        =   68
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "����Χ"
            Height          =   180
            Left            =   0
            TabIndex        =   67
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Lblҽ������ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   66
            Top             =   2460
            Width           =   720
         End
      End
      Begin VB.PictureBox picDeptList 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         ScaleHeight     =   1335
         ScaleWidth      =   3375
         TabIndex        =   38
         Top             =   6960
         Width           =   3375
         Begin VB.Frame fraLineH3 
            Height          =   50
            Left            =   0
            TabIndex        =   44
            Top             =   120
            Width           =   3525
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "ˢ���嵥"
            Height          =   375
            Left            =   2040
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":8E44
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdRefreshDept 
            Caption         =   "ˢ�¿���"
            Height          =   375
            Left            =   1320
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdListSel 
            Height          =   255
            Left            =   50
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":93CE
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   300
            Width           =   255
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ȫѡ"
            Enabled         =   0   'False
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   40
            Top             =   337
            Width           =   735
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ȫѡ"
            Enabled         =   0   'False
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   39
            Top             =   337
            Width           =   735
         End
         Begin MSComctlLib.TreeView tvwList 
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1296
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   476
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            ImageList       =   "imgTvw"
            Appearance      =   1
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
         Begin MSComctlLib.TreeView tvwList 
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1296
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   476
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            ImageList       =   "imgTvw"
            Appearance      =   1
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
      End
      Begin VB.PictureBox picConMain 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   30
         ScaleHeight     =   3615
         ScaleWidth      =   3375
         TabIndex        =   13
         Top             =   0
         Width           =   3375
         Begin VB.Frame fraLineH1 
            Height          =   50
            Left            =   0
            TabIndex        =   29
            Top             =   480
            Width           =   3405
         End
         Begin VB.ComboBox cbo��ҩҩ�� 
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   120
            Width           =   2415
         End
         Begin VB.ComboBox cboʱ�䷶Χ 
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Left            =   840
            TabIndex        =   26
            Top             =   1680
            Width           =   2415
         End
         Begin VB.CheckBox chkSend 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��Ժ��ҩ"
            Height          =   180
            Index           =   1
            Left            =   2160
            TabIndex        =   25
            Top             =   2040
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSend 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ȡҩ"
            Height          =   180
            Index           =   2
            Left            =   840
            TabIndex        =   24
            Top             =   2280
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSend 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ժ����ҩ"
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   23
            Top             =   2040
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.PictureBox picSendType 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   2895
            TabIndex        =   21
            Top             =   2880
            Width           =   2895
            Begin VB.CheckBox chkSendType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�Զ��巢ҩ���ͣ���̬����"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.PictureBox picShowOther 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            MouseIcon       =   "frmҩƷ���ŷ�ҩnew.frx":9F08
            ScaleHeight     =   270
            ScaleWidth      =   2655
            TabIndex        =   18
            Tag             =   "0"
            Top             =   3240
            Width           =   2655
            Begin VB.PictureBox picUpOrDown 
               BackColor       =   &H00FFEDDD&
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   2400
               Picture         =   "frmҩƷ���ŷ�ҩnew.frx":A212
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   19
               Top             =   0
               Width           =   270
            End
            Begin VB.Label lblComment 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "��ʾ��������"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   0
               TabIndex        =   20
               Top             =   45
               Width           =   1080
            End
         End
         Begin VB.PictureBox picShowSendType 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            MouseIcon       =   "frmҩƷ���ŷ�ҩnew.frx":A554
            ScaleHeight     =   270
            ScaleWidth      =   2655
            TabIndex        =   15
            Tag             =   "0"
            Top             =   2520
            Width           =   2655
            Begin VB.PictureBox picUpOrDown1 
               BackColor       =   &H00FFEDDD&
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   2400
               Picture         =   "frmҩƷ���ŷ�ҩnew.frx":A85E
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   16
               Top             =   0
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "��ʾ������ҩ����"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   0
               TabIndex        =   17
               Top             =   45
               Width           =   1440
            End
         End
         Begin VB.CommandButton cmdIC 
            Caption         =   "����"
            Height          =   300
            Left            =   2640
            TabIndex        =   14
            Top             =   1680
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComCtl2.DTPicker Dtp����ʱ�� 
            Height          =   315
            Left            =   840
            TabIndex        =   30
            Top             =   1320
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   84803587
            CurrentDate     =   39998
         End
         Begin MSComCtl2.DTPicker Dtp��ʼʱ�� 
            Height          =   300
            Left            =   840
            TabIndex        =   31
            Top             =   960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   84803587
            CurrentDate     =   39998
         End
         Begin VB.Label lbl��ҩҩ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ҩҩ��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   0
            TabIndex        =   37
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lblʱ�䷶Χ 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ʱ�䷶Χ"
            Height          =   180
            Left            =   0
            TabIndex        =   36
            Top             =   660
            Width           =   720
         End
         Begin VB.Label lblPatiInputType 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ�š�"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   35
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label lblTimeEnd 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   34
            Top             =   1387
            Width           =   720
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʼʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   33
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label lbl��ҩ���� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ҩ����"
            Height          =   180
            Left            =   0
            TabIndex        =   32
            Top             =   2160
            Width           =   720
         End
      End
   End
   Begin VB.Frame fraColorStateReturn 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   4440
      Width           =   3840
      Begin VB.PictureBox picColorStateReturn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   0
         Picture         =   "frmҩƷ���ŷ�ҩnew.frx":ABA0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   89
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picColorStateReturn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateReturn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1995
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateReturn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblColorStateReturn 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   3
         Left            =   260
         TabIndex        =   90
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColorStateReturn 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ"
         Height          =   180
         Index           =   0
         Left            =   1455
         TabIndex        =   11
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateReturn 
         AutoSize        =   -1  'True
         Caption         =   "ԭʼ"
         Height          =   180
         Index           =   1
         Left            =   2400
         TabIndex        =   10
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateReturn 
         AutoSize        =   -1  'True
         Caption         =   "�ѷ�ҩ"
         Height          =   180
         Index           =   2
         Left            =   3285
         TabIndex        =   9
         Top             =   30
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList imgPacker 
      Left            =   5520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":113F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":1198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":11F26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerAuto 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3840
      Top             =   240
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   4080
      ScaleHeight     =   1935
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   960
      Width           =   3015
      Begin VB.Frame fraLineV1 
         Height          =   2085
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList imgTvw 
      Left            =   6240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":124C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":12A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":12FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":1358E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLvwSel 
      Left            =   6840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":13B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":13E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":1415C
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":144AE
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw��ҩ;�� 
      Height          =   345
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLvwSel"
      SmallIcons      =   "imgLvwSel"
      ColHdrIcons     =   "imgLvwSel"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView LvwҩƷ���� 
      Height          =   345
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLvwSel"
      SmallIcons      =   "imgLvwSel"
      ColHdrIcons     =   "imgLvwSel"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   87
      Top             =   8655
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmҩƷ���ŷ�ҩnew.frx":14800
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6482
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "δ�������������0��   "
            TextSave        =   "δ�������������0��   "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "�ְ���"
            TextSave        =   "�ְ���"
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   5040
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmҩƷ���ŷ�ҩnew.frx":15094
      Left            =   4440
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm���ŷ�ҩ����New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''�����˵�
Private Const mconMenu_TypePopup = 3000                  '��ҩ;������
Private mTypeCount As Integer                            '��ҩ;�����������

Private Const mconMenu_SortPopup = 6000                  '����ʽ
Private Const mconMenu_SortPopup_ByName = 6001           '�����б���������ʽ��������
Private Const mconMenu_SortPopup_ByBedNo = 6002          '����λ��

Public mblnEnter As Boolean                              '�Ƿ��ܽ���

Private mblnStartPacker As Boolean                       '�Ƿ�����ҩƷ�ְ����ӿ�
Private mblnPackerConnect As Boolean                     '�ְ����ӿ����ݿ��Ƿ�����
Private mlngҩ��ID As Long
Private mstrҩ������ As String

Private mstrCardType As String   '���п���𣬸�ʽ������|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);��
Private mintCardCount As Integer  '������
Private mobjSquareCard As Object             'һ��ͨ�ӿ�

Private mblnFreshDeptList As Boolean
Private mblnStart As Boolean

Private mstrDeptNode As String      '��ѡҩ����Ӧ��վ��
Private mRsDept As Recordset
Private mblnCheck As Boolean        '���ͬ�鷢ҩ�Ƿ���Ҫ���
Private rstemp As Recordset
Private mrsApplyforcredit As Recordset      '���ڼ�¼������������ĵ���
Private mblnIs�������� As Boolean
Private mclsComLib As Object
Private mobjDrugMAC As Object       '��ҩ�ӿڲ���
Private mobjPlugIn As Object             '��ҽӿڶ���
Private mstr����id As String

'��Ϣ��ض������
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule
Attribute mobjMipModule.VB_VarHelpID = -1
Private mrsReceiveMsg As ADODB.Recordset    '���յ�����Ϣ��¼��
Private mdateBegin As Date

Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

''''����
'�б�����
Private Enum mDeptType
    ��ҩ = 0
    ��ҩ = 1
End Enum

Private Enum mListType
    ��ҩ = 0
    ���� = 1
    ȱҩ = 2
    �ܷ� = 3
    ��ҩ = 4
    ���� = 5
End Enum

'ʱ�䷶Χ
Private Enum mTimeRange
    ���� = 0
    ������ = 1
    ������ = 2
    ָ��ʱ�䷶Χ = 3
End Enum

'¼����Ϣ����
Private Enum mInputType
    סԺ�� = 0
    ���� = 1
    ���� = 2
    NO = 3
    ����ID = 4
    ��ҩ�� = 5
    ��ҩ�� = 6
    ��ҩ���� = 7
    IC�� = 8
End Enum

'ִ��״̬
Private Enum mState
    ȱҩ = 0
    ��ҩ = 1
    �ܷ� = 2
    ������ = 3
    �ܷ�_�ָ� = 4
    �ܷ�_������ = 5
    ��ҩ = 6
    ��ҩ_ԭʼ��¼ = 7
    ��ҩ_��ҩ��¼ = 8
    ��ҩ_��ҩ��¼ = 9
    ת����¼ = 10
End Enum

'��ҩ;����ҩƷ����ѡ��
Private Enum mSel
    ��ҩ;�� = 0
    ҩƷ���� = 1
End Enum

'��ҩ�б���ɫ
Private Enum mSendListColor
    SendState = 0
    RejectState = 1
    UnProcessState = 2
    ShortageState = 3
End Enum

'��ҩ�б���ɫ
Private Enum mReturnListColor
    ReturnState = 0
    OriginalState = 1
    SendedState = 2
End Enum

'''����

'Ĭ�ϵĴ����С
Private Const mcstlngWinNormalWidth As Long = 13275
Private Const mcstlngWinNormalHeight As Long = 8500

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private mlngMyWindow As Long

Private mdate�ϴ�ˢ��ʱ�� As Date                       '��¼�ϴ�ˢ��ʱϵͳʱ��

Private mstrPrivs As String                                 'Ȩ�޴�
Private mlngMode As Long                                    'ģ���

Private mcur���ܷ�ҩ�� As Currency

'��ѯ��ʾ
Private Type TYPE_FindWar
    blnNoAsk_Dept_Send As Boolean                      '��ѯʱ�����ʱ��ʾ���Ƿ��´β�����ʾ����ҩʱ
    blnNoAsk_Dept_Sended As Boolean                    '��ѯʱ�����ʱ��ʾ���Ƿ��´β�����ʾ����ҩʱ
    blnNoAsk_Rec As Boolean                            '��ѯ��ϸ��¼����ʱ��ʾ���Ƿ��´β�����ʾ
    blnProc_Dept_Send As Boolean                       '��ѯ�����б��Ƿ��������ҩʱ
    blnProc_Dept_Sended As Boolean                     '��ѯ�����б��Ƿ��������ҩʱ
    blnProc_Rec As Boolean                             '��ѯ��ϸ��¼���Ƿ����
End Type
Private mFindWar As TYPE_FindWar

Private mfrmDetail As New frm���ŷ�ҩ�嵥

Private mblnExistOtherSendType As Boolean                   '�Ƿ����Զ���ķ�ҩ����
Private mblnCard As Boolean                                 '�Ƿ�ˢ���￨

Public BlnSetPara As Boolean                                '�������ô����Ƿ�ȷ�����˳�
Public BlnRefresh As Boolean                                '���������Ƿ���������,����ˢ��
Private mblnInput As Boolean                                '�Ƿ���ͨ��¼�벡����Ϣ��ʽ����������

Private mstr����ID�� As String                              '��ҩ����ID
Private mstr�������� As String                              '��ҩ��������

'�������ݼ�
Private mrsDeptList As ADODB.Recordset                      '���ݲ����б�ʵ�ʹ�ѡ������ù�ѡ�Ĳ��š�NO����Ϊ��Ҫ����������ȡ��ϸ����
Private mrsSendData As ADODB.Recordset                      '����ҩƷ��¼��
Private mrsReturnData As ADODB.Recordset                    '��ҩҩƷ��¼��
Private mrsChargeOff As New ADODB.Recordset                   '������ʾ���������¼
Private mrsChargeOffMain As New ADODB.Recordset               '��������
Private mrs��ҩ;�� As ADODB.Recordset
Private mrs��Ʒ���� As ADODB.Recordset
Private mrsPASS As ADODB.Recordset                          'PASS�����ݼ�

'ҽ���ӿ�
Private gclsInsure As New clsInsure
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

'Ȩ��
Private Type Type_Privs
    bln����ҩ�� As Boolean
    bln��ҩ As Boolean
    bln��ҩ As Boolean
    bln������ҩ���Ĵ��� As Boolean
    bln���˽��ʴ��� As Boolean
    bln���˳�Ժ���˴��� As Boolean
    blnȱҩ���� As Boolean
    bln�ܷ� As Boolean
    bln������ҩ��� As Boolean
    blnҽ����ѯ As Boolean
    bln�������� As Boolean
    bln��ҩ���� As Boolean
    bln�޸��������� As Boolean
    blnֹͣ��ҩ As Boolean
    bln�ָ���ҩ As Boolean
End Type
Private mPrives As Type_Privs


'�б��ѯ����
Private Type Type_Condition
    '��Ҫ����
    lngҩ��ID As Long
    str��ʼʱ�� As String
    str����ʱ�� As String
    int��ҩ���� As Integer
    str������ҩ���� As String
        
    '¼����Ϣ����ѡ����
    strסԺ�� As String
    str���� As String
    str���� As String
    strNo As String
    lng����ID As Long
    str���￨ As String
    str��ҩ�� As String
    cur��ҩ�� As Currency
    lng��ҩ����ID As Long
    strIC�� As String
    
    
    '��Ҫ����
    str��ҩ;�� As String
    strҩƷ���� As String
    int����Χ As Integer
    intҽ������ As Integer
    int�������� As Integer
    
    '��������
    int����ģʽ As Integer
    str������ As String
    bln��ʾ��ҩ�������� As Boolean
    bln��ʾ������ҩ���� As Boolean
    bln��ʾ��ҩ��ҩ�� As Boolean
    int��ʾ��ҩ����ģʽ As Integer          '0-�����ҡ����ˡ�NO��֯��1-����ҩ�š����ҡ����ˡ�NO��֯
End Type
Private mcondition As Type_Condition

'ʹ�õ��Ĳ���������ϵͳ�����������������򱾻�ע���
Private Type Type_Params
    '�������е�ϵͳ����
    bln����δ��˴�����ҩ As Boolean
    bln����ҽ�������Ϻ���ҩ As Boolean
    int����λ�� As Integer
    bln��˻��۵� As Boolean
    intЧ����ʾ��ʽ As Integer          '0-��ʧЧ����ʾ��1-����Ч����ʾ
    intҩƷ������ʾ As Integer          '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
    bln������     As Boolean          '�Ƿ����ô������ϵͳ
    
    '�������е���������
    intDays As Integer
    int��ҩ���� As Integer
    bln��ҩ��ǩ�� As Boolean
    blnȱҩ��� As Boolean
    bln��ҩ��ǩ�� As Boolean
    int�Զ�ˢ��δ��ҩ�嵥 As Integer
    blnҩƷ���� As Boolean
    bln���ܷ�ҩ As Boolean
    int����ģʽ As Integer
    intҽ������ As Integer
    bln������ʾ As Boolean
    str������ As String
    str������� As String
    str��ֵ���� As String
    str��Σ���� As String
    str��Σ���� As String
    lngҩ��ID As Long
    int�Զ���ӡ As Integer
    bln�������� As Boolean
    bln���̵��� As Boolean
    int��ѯ��ҩ���� As Integer
    int��ѯ��ҩ���� As Integer
    lng����¼�� As Long
    bln��˳�Ժ�������� As Boolean
    int��ҩ�嵥��ӡ As Integer
    intCheck As Integer
    int��ҩʱ���ҽ�� As Integer
    bln����������ʱ���ܽ�����ҩ���� As Boolean
    bln���ط�ҩʱ�����ҩ���� As Boolean
    
    '�����
    IntCheckStock As Integer
    
    '�ⷿ��λ
    strUnit As String
    
    '���ú�����ҩPASS
    blnStarPass As Boolean
    
    '��������
    bln�������� As Boolean
    
    '������Դ
    strSourceDep As String
    
    'ע������
    intҩƷ���Ʊ�����ʾ As String       '0��ҩƷ���������ƣ�1��ҩƷ���룻2��ҩƷ����
    intFont As Integer                  '�����
    StrFindStyle As String              '����ƥ��
    int����ģʽ As Integer
    int�������� As Integer                  '�����б��У���������ʽ��1-��������2-����λ
    blnOnlyShowDept As Boolean              '�����б��Ƿ����ʾ��������
    intShowDept As Integer                  '0-��ʾ���п���;1-��ʾ�ٴ�����;2-��ʾҽ������;3-��ʾ���˲���
    blnShowReject As Boolean                '��ȡ�ܷ�ҩƷ��0-����ȡ�ܷ�ҩƷ��1-��ȡ�ܷ�ҩƷ
    intAdviceType As Integer                'ҽ�����ͣ�0-�������е���;1-��������ҽ��;2-������ʱҽ��;3-��ͨ���ʵ���;4-��������ҽ��
    blnSort As Boolean                      'ҽ���б���Ұ�ҽ�������ʱ������
    
    'ע����������װ�����
    int��ͣ���� As Integer              '��ͣ��ҩʱ���װ����������:0-����;1-��ͣ����
    str��װ������ As String             '�������ݵ����ͣ���ʽ��00������1λ��ʾ������������2λ��ʾ����������0����ʾ��������1����ʾ����
    str��װ������ As String             '�������ݵļ��ͣ��������ƴ����á�;���ָ�������ǡ����С����ʾ���м���
End Type
Private mParams As Type_Params

Private Function CheckDangerDrug(ByVal rsData As ADODB.Recordset) As Boolean
    '����ΣҩƷ�������ΣҩƷ��Ҫ��������ʱ�������Ƿ������ͨҩƷ
    Dim bln��ͨҩƷ As Boolean
    Dim bln��ΣҩƷ As Boolean
    Dim lngҩƷid As Long
    
    If mParams.str��Σ���� = "" Then
        CheckDangerDrug = True
        Exit Function
    End If
    
    With rsData
        .Filter = "ִ��״̬=" & mState.��ҩ
        .Sort = "ҩƷID Asc"
        
        Do While Not .EOF
            If lngҩƷid <> !ҩƷid Then
                If !��ΣҩƷ = 0 Then
                    bln��ͨҩƷ = True
                ElseIf InStr(1, mParams.str��Σ����, !��ΣҩƷ) > 0 Then
                    bln��ΣҩƷ = True
                End If
                
                If bln��ͨҩƷ And bln��ΣҩƷ Then
                    MsgBox "��ʾ����ΣҩƷ���ܺ���ͨҩƷ���ܷ�ҩ��", vbInformation, gstrSysName
                    CheckDangerDrug = False
                    Exit Function
                End If
                    
                lngҩƷid = !ҩƷid
            End If
            .MoveNext
        Loop
    End With
    
    CheckDangerDrug = True
End Function

Private Sub DrugStoreWork_PrintBill()
    '��ӡ��ҩ����
    Dim intAllFormat As Integer
    
    If GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat") <> "" Then
        intAllFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    Else
        intAllFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    End If
    
    If mParams.int�Զ���ӡ = 2 Then
        If MsgBox("�Ƿ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            If intAllFormat = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                    "��ҩ�ⷿ=" & mcondition.lngҩ��ID, _
                    "��ҩ��=" & mcur���ܷ�ҩ��, _
                    "��ҩ����=" & mstr�������� & "|" & " IN (" & mstr����ID�� & ")", _
                    "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"), _
                    "PrintEmpty=0", 2)
            Else
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                    "��ҩ�ⷿ=" & mcondition.lngҩ��ID, _
                    "��ҩ��=" & mcur���ܷ�ҩ��, _
                    "��ҩ����=" & mstr�������� & "|" & " IN (" & mstr����ID�� & ")", _
                    "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"), _
                    "ReportFormat=" & mfrmDetail.Get��ǰ��ҩ����ʽ, "PrintEmpty=0", 2)
            End If
        End If
    ElseIf mParams.int�Զ���ӡ = 1 Then
        If intAllFormat = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "��ҩ�ⷿ=" & mcondition.lngҩ��ID, _
                "��ҩ��=" & mcur���ܷ�ҩ��, _
                "��ҩ����=" & mstr�������� & "|" & " IN (" & mstr����ID�� & ")", _
                "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"), _
                "PrintEmpty=0", 2)
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "��ҩ�ⷿ=" & mcondition.lngҩ��ID, _
                "��ҩ��=" & mcur���ܷ�ҩ��, _
                "��ҩ����=" & mstr�������� & "|" & " IN (" & mstr����ID�� & ")", _
                "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"), _
                "ReportFormat=" & mfrmDetail.Get��ǰ��ҩ����ʽ, "PrintEmpty=0", 2)
        End If
    End If
End Sub

Private Sub DrugStoreWork_SendToPacker(ByVal rsData As ADODB.Recordset)
    Dim str�շ�ids As String, strMessage As String
    Dim arr�շ�ids As Variant
    Dim lng��ǰ����id As Long
    Dim n As Integer
    
    On Error GoTo errHandle
    
    If mblnPackerConnect = True And Not mobjDrugMAC Is Nothing Then
        If TypeName(mobjDrugMAC) = "clsDrugMachine" Then
            '�½ӿ�
            
            With rsData
                .Filter = "ִ��״̬=" & mState.��ҩ & " And ��ҩ����>0"
                .Sort = "��ҩ����ID,����id"
        
                If .EOF Then Exit Sub
                            
                arr�շ�ids = Array()
                
                '����ҩ���ŷ����ϴ�
                Do While Not .EOF
                    If lng��ǰ����id <> !��ҩ����ID Then
                        If str�շ�ids <> "" Then
                            ReDim Preserve arr�շ�ids(UBound(arr�շ�ids) + 1)
                            arr�շ�ids(UBound(arr�շ�ids)) = str�շ�ids
                        End If
                        
                        lng��ǰ����id = !��ҩ����ID
                        str�շ�ids = "2|" & !�շ�Id
                    Else
                        str�շ�ids = str�շ�ids & ";" & !�շ�Id
                    End If
                    
                    .MoveNext
                    
                    If .EOF And str�շ�ids <> "" Then
                        '����û�м�¼ʱ���뵽����
                        ReDim Preserve arr�շ�ids(UBound(arr�շ�ids) + 1)
                        arr�շ�ids(UBound(arr�շ�ids)) = str�շ�ids
                    End If
                Loop
                
                For n = 0 To UBound(arr�շ�ids)
                    mobjDrugMAC.Operation gstrDbUser, Val("21-��ϸ�ϴ�"), CStr(arr�շ�ids(n)), strMessage
                Next
            End With
        Else
            '�����Ͻӿ�
            Call PackerTransDetail_DYEY(rsData)
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PackerTransDetail_DLSY(ByVal rsData As ADODB.Recordset)
    '�Զ���ҩ�����ݴ��䣺������Ժר��
    'ֱ�Ӵ���HIS�˵��м��
    
''��ˮ��
'PrescriptionNo
''���
'Seqno
''С���־
'Group_No
''������
'MachineNo
''����״̬
'ProcFlg
''����ID
'PatientID
''��������
'PatientName
''�����Ա�
'Sex
''����סԺ��־
'IOFlg
''��������
'WardCd
''��������
'WardName
''��λ��
'BedNo
''��������
'PrescriptionDate
''�״���ҩ����
'TakeDate
''��ʼʱ��
'TakeTime
''����ʱ��
'LastTime
''�������
'Presc_Class
''ҩƷ����
'Drugcd
''ҩƷ����
'DrugName
''��ҩ��λ
'DispensedUnit
''��ҩ����
'Dispense_days
''�÷�
'Freq_desc
''����ʱ��
'Freq_desc_Detail
''�ʱ��
'MakeRecTime

End Sub
Private Sub PackerTransDetail_DYEY(ByVal rsData As ADODB.Recordset)
    '�Զ���ҩ�����ݴ��䣺��ҽ��Ժר��
    '���ýӿں��������м����ݿ�
    Dim str���� As String
    Dim rsTmp As ADODB.Recordset
    Dim lng��ǰ���� As Long
    Dim str��ϸ As String
    Dim strReturn As String
    Dim str�ְ��豸��� As String
    Dim strTmp As String
    Dim strFilter As String
    Dim strDetail As String
    
    On Error GoTo errHandle
    
    If mblnStartPacker = False Or mblnPackerConnect = False Then Exit Sub
    If mParams.int��ͣ���� = 1 Then Exit Sub
    If Val(Mid(mParams.str��װ������, 1, 1)) = 0 And Val(Mid(mParams.str��װ������, 2, 1)) = 0 Then Exit Sub
    If mParams.str��װ������ = "" Then Exit Sub
    
    str�ְ��豸��� = "1"
    
    If mlngҩ��ID <> mcondition.lngҩ��ID Or mstrҩ������ = "" Then
        mlngҩ��ID = mcondition.lngҩ��ID
        gstrSQL = "select ���� from ���ű� where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩ������", mlngҩ��ID)
        mstrҩ������ = rsTmp!����
    End If
    
    With rsData
        If Val(Mid(mParams.str��װ������, 1, 1)) = 1 Then
            strFilter = "ִ��״̬=" & mState.��ҩ & " And ����='����' And ��ҩ����>0"
        End If
        If Val(Mid(mParams.str��װ������, 2, 1)) = 1 Then
            If strFilter <> "" Then
                strFilter = "(" & strFilter & ")"
                strFilter = strFilter & " Or (ִ��״̬=" & mState.��ҩ & " And ����='����' And ��ҩ����>0) "
            Else
                strFilter = "ִ��״̬=" & mState.��ҩ & " And ����='����' And ��ҩ����>0 "
            End If
        End If
        
        .Filter = strFilter
        .Sort = "��ҩ����ID"
        
        If .EOF Then Exit Sub
        
        lng��ǰ���� = !��ҩ����ID
        str���� = !��ҩ���ű��� & ";" & mstrҩ������ & ";" & str�ְ��豸��� & IIf(gstrUnitName = GCST_UNIT_YQMY, ";" & CStr(mcur���ܷ�ҩ��), "")
        Do While Not .EOF
            If lng��ǰ���� <> !��ҩ����ID Then
                '��ǰ���Ų�һ��ʱ���������ݣ�������û�д��ݳɹ����շ�ID
                If str��ϸ <> "" Then
                    strReturn = IIf(strReturn = "", "", strReturn & ";") & mobjDrugMAC.TranDrugPacker(str���� & "|" & str��ϸ)
                End If
                
                '����ָ����ǰ����
                lng��ǰ���� = !��ҩ����ID
                str���� = !��ҩ���ű��� & ";" & mstrҩ������ & ";" & str�ְ��豸��� & IIf(gstrUnitName = GCST_UNIT_YQMY, ";" & CStr(mcur���ܷ�ҩ��), "")
                str��ϸ = GetMediPackerDetail(!�շ�Id, mParams.str��װ������, !����)
            Else
                strDetail = GetMediPackerDetail(!�շ�Id, mParams.str��װ������, !����)
                If strDetail <> "" Then
                    str��ϸ = IIf(str��ϸ = "", "", str��ϸ & "|") & strDetail
                End If
            End If
            
            .MoveNext
            
            If .EOF And str��ϸ <> "" Then
                '����û�м�¼ʱ���������ݣ�������û�д��ݳɹ����շ�ID
                strReturn = IIf(strReturn = "", "", strReturn & ";") & mobjDrugMAC.TranDrugPacker(str���� & "|" & str��ϸ)
            End If
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ExecuteWriteOffByMessage(ByVal objMsgBar As CommandBarControl)
    'ͨ����Ϣ���ó�������
    '���û��������Ϣ��Ŀ�����ݳ�����Ҫ�Ĺؼ���Ϣ
    '���ݸ�ʽΪ������ʱ��,����id|����ʱ��,����id
    '�����ĳ��������Ϣ����ôȡ�˵��б������Ϣ�����ݣ����������ִ�еģ���ô���ݼ�¼����������Ϣ
    Dim strMsg As String
    
    If Not objMsgBar Is Nothing Then
        If objMsgBar.Parameter <> "" Then
            strMsg = objMsgBar.Parameter
        Else
            With mrsReceiveMsg
                If mrsReceiveMsg.RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        strMsg = IIf(strMsg = "", "", strMsg & "|") & !����ʱ�� & "," & !����ID
                        .MoveNext
                    Loop
                End If
            End With
        End If
       
        '����������˴���
        Call ShowWindow_ReVerify(strMsg)
    End If
End Sub

Private Sub SetMessageBar(ByVal rsMsg As ADODB.Recordset)
    '������Ϣ�˵�
    '��ɾ���Ӳ˵����ٸ��ݼ�¼���е����������Ӳ˵�
    '������ʱ���˵���ʾ�ж�����Ϣ�����û���κ���Ϣ�������ظ��˵�
    '����Ӳ˵�ʱ�������Ϣ����5������ֻ��ʾ5��������Ϣ
    '�����Ϣ����1��ʱ���������һ���Ӳ˵���ȫ����ˡ�
    'strDelMsg����Ϊ��ʱ��ɾ����Ϣ��¼���ж�Ӧ����Ŀ
    Dim cbrControlMain As CommandBarPopup
    Dim cbrControlPopup As CommandBarControl
    Dim IntCount As Integer
    Dim intTemp As Integer
    Dim blnTemp As Boolean
                
    If mobjMipModule Is Nothing Then Exit Sub
    
    If rsMsg Is Nothing Then Exit Sub
    
    Set cbrControlMain = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_File_Message)
    cbrControlMain.Visible = True
    If Not cbrControlMain Is Nothing Then
        Set cbrControlMain = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_File_Message)
        cbrControlMain.Visible = True
        If rsMsg.RecordCount > 0 Then rsMsg.MoveFirst
        If rsMsg.RecordCount = 0 Then
            cbrControlMain.Visible = False
        Else
            cbrControlMain.Caption = "����Ϣ����" & "(" & rsMsg.RecordCount & ")"
            
            For Each cbrControlPopup In cbrControlMain.CommandBar.Controls
                If Not rsMsg.EOF And intTemp <= 5 Then
                    cbrControlPopup.Caption = Format(rsMsg!����ʱ��, "mm-dd hh:mm") & " " & rsMsg!���� & " " & rsMsg!סԺ��
                    cbrControlPopup.Parameter = rsMsg!����ʱ�� & "," & rsMsg!����ID
                    cbrControlPopup.Visible = True
                    rsMsg.MoveNext
                Else
                    If intTemp < cbrControlMain.CommandBar.Controls.count Then
                        cbrControlPopup.Visible = False
                    Else
                        blnTemp = True
                    End If
                End If
                
                intTemp = intTemp + 1
            Next
                
            For IntCount = intTemp + 1 To rsMsg.RecordCount
                If IntCount <= 5 Then
                    Set cbrControlPopup = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_File_Message + IntCount, Format(rsMsg!����ʱ��, "mm-dd hh:mm") & " " & rsMsg!���� & " " & rsMsg!סԺ��)
                    cbrControlPopup.Parameter = rsMsg!����ʱ�� & "," & rsMsg!����ID
                Else
                    Exit For
                End If
                rsMsg.MoveNext
            Next
            If IntCount > 2 And (blnTemp = True Or intTemp < 6) Then
                Set cbrControlPopup = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_File_Message + IntCount, "ȫ�����")
            End If
        End If
    End If
End Sub

Private Sub chkDanger_Click()
    chkDangerType(0).Enabled = (chkDanger.Value = 1)
    chkDangerType(1).Enabled = chkDangerType(0).Enabled
    chkDangerType(2).Enabled = chkDangerType(0).Enabled
    
    
End Sub

Private Sub chkDangerType_Click(Index As Integer)
    Dim objChk As CheckBox
    Dim blnAllUnCheck As Boolean
    
    If mblnStart = False Then Exit Sub
    
    blnAllUnCheck = True
    
    For Each objChk In chkDangerType
        If objChk.Value = 1 Then
            blnAllUnCheck = False
        End If
    Next
    
    If blnAllUnCheck = True Then
        chkDangerType(Index).Value = 1
    End If
End Sub


Private Sub chkToxicologyType_Click()
    Me.chkToxicology(0).Enabled = (Me.chkToxicologyType.Value = 1)
    Me.chkToxicology(1).Enabled = Me.chkToxicology(0).Enabled
    Me.chkToxicology(2).Enabled = Me.chkToxicology(0).Enabled
    Me.chkToxicology(3).Enabled = Me.chkToxicology(0).Enabled
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    If Not txtInput.Locked And txtInput.Text = "" And Me.ActiveControl Is txtInput And strNo <> "" Then
        txtInput.Text = strNo
        
        If txtInput.Text = "" Then
            Call mobjICCard.SetEnabled(False)
        Else
            If mParams.int����ģʽ <> mInputType.IC�� Then
                mParams.int����ģʽ = mInputType.IC��
                Call SetInputState(mParams.int����ģʽ)
            
                DoEvents
            End If
            
            Call txtInput_KeyPress(vbKeyReturn)
        End If
    End If
End Sub
Private Function DrugStoreWork_CheckSend(ByVal rsData As ADODB.Recordset) As Boolean
    '��ҩ���
    Dim rsGroupCheck As ADODB.Recordset
    Dim strCheckMsg As String
    
    On Error GoTo errHandle
    
    '���ȼ���ΣҩƷ
    If CheckDangerDrug(rsData) = False Then Exit Function
    
    '���洢�ⷿ
    If CheckDrugStock(rsData) = False Then Exit Function
    
    '����Ƿ������������δ��˵ĵ���
    If CheckNotAudited(rsData) = False Then Exit Function
    
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    If Not CheckCorrelation(0, rsData) Then Exit Function
    
    '����������
    If CheckShortage(rsData, True, strCheckMsg) = False Then
        '�����
        If mParams.IntCheckStock = 2 Then
            '��治���ֹ��ҩ
            MsgBox "����ҩƷʵ�ʿ���������㣬���ܷ�ҩ��" & vbCrLf & strCheckMsg, vbInformation, gstrSysName
            
            If mParams.blnȱҩ��� Then
                Call mfrmDetail.RefreshList(mListType.��ҩ, mrsSendData, mrsChargeOff)
            End If
            Exit Function
        ElseIf mParams.IntCheckStock = 1 Then
            '��治�㣬����
            If MsgBox("����ҩƷʵ�ʿ���������㣬�Ƿ������ҩ��" & vbCrLf & strCheckMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If mParams.blnȱҩ��� Then
                    Call mfrmDetail.RefreshList(mListType.��ҩ, mrsSendData, mrsChargeOff)
                End If
                Exit Function
            End If
        End If
    End If
    
    '����״̬��������
    Set rsGroupCheck = rsData.Clone
    With rsData
        .Filter = "ִ��״̬=" & mState.��ҩ
        .Sort = "�շ�ID"
        Do While Not .EOF
            '��鵥��״̬
            If DeptSendWork_CheckBill(1, !�շ�Id, mParams.bln����δ��˴�����ҩ) > 0 Then Exit Function
            
            '������״̬
            If Not mblnCheck Then
                If CheckGroupSend(rsGroupCheck, Val(!���ID), !NO) = False Then Exit Function
            End If
            
            .MoveNext
        Loop
    End With
    
    DrugStoreWork_CheckSend = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AutoRefresh()
    '�Զ�ˢ��ֻ���δ��ҩƷ�嵥
    Dim dateCurr As Date
        
    '���������С��ʱ�˳�
    If Me.WindowState = 1 Then Exit Sub
    
    '�������ڲ��ǵ�ǰ����ʱ�˳�
    If mlngMyWindow = 0 Then
        mlngMyWindow = GetActiveWindow()
    Else
        If mlngMyWindow <> GetActiveWindow() Then Exit Sub
    End If
    
    '�������δ��ҩ��������Զ�ˢ�²���Ϊ0ʱ�˳�
    If tbcDetail.Selected.Index <> mListType.��ҩ Or mParams.int�Զ�ˢ��δ��ҩ�嵥 = 0 Then Exit Sub
    
    '���ݵ�ǰʱ�����ϴ�ˢ��ʱ�����������Ƿ�ˢ��
    dateCurr = zlDatabase.Currentdate
    If DateDiff("s", mdate�ϴ�ˢ��ʱ��, dateCurr) < mParams.int�Զ�ˢ��δ��ҩ�嵥 * 60 Then Exit Sub
    
    TimerAuto.Enabled = False
    
    'ˢ������
    cmdRefresh_Click

'    MsgBox "Ok��" & "[" & Format(dateCurr, "yyyy-mm-dd hh:mm:ss") & "]" & "[" & Format(mdate�ϴ�ˢ��ʱ��, "yyyy-mm-dd hh:mm:ss") & "]"
'    mdate�ϴ�ˢ��ʱ�� = zldatabase.Currentdate
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub BillPrint_Restore()
    '���ܣ���ӡ��ҩ֪ͨ��
    Dim strDate As String
    
    strDate = Format(mfrmDetail.GetReturnDate, "yyyy-MM-dd HH:mm:ss")
    
    If strDate = "" Then Exit Sub
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, _
        "��ҩʱ��=" & strDate, _
        "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), _
        "��ҩ�ⷿ=" & mcondition.lngҩ��ID, _
        2)
End Sub


Private Sub BillPrint_Total()
    Dim rsTmp As ADODB.Recordset
    Dim strҩ�� As String, str���� As String
    Dim str��ҩ As String
    Dim str��ҩ���� As String
    Dim str��ҩ����ID As String
    Dim var��ҩ�� As Variant
    Dim intAllFormat As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select ����,���� From ���ű� Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ǰҩ��������]", mcondition.lngҩ��ID)

    If Not rsTmp.RecordCount <= 0 Then strҩ�� = "(" & rsTmp!���� & ")" & rsTmp!����
    
    If GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat") <> "" Then
        intAllFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    Else
        intAllFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    End If
    
    If tbcDetail.Selected.Index = mListType.��ҩ Then
        str��ҩ = mfrmDetail.GetSendedInfo
                
        If str��ҩ <> "" Then
            str��ҩ���� = Split(str��ҩ, "|")(0)
            str��ҩ����ID = Split(str��ҩ, "|")(1)
            var��ҩ�� = Split(str��ҩ, "|")(2)
        End If
        
        If intAllFormat = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "��ҩ�ⷿ=" & strҩ�� & "|" & mcondition.lngҩ��ID, _
                "��ҩ��=" & var��ҩ��, _
                "��ҩ����=" & str��ҩ���� & "|" & " IN (" & str��ҩ����ID & ")")
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "��ҩ�ⷿ=" & strҩ�� & "|" & mcondition.lngҩ��ID, _
                "��ҩ��=" & var��ҩ��, _
                "��ҩ����=" & str��ҩ���� & "|" & " IN (" & str��ҩ����ID & ")", "ReportFormat=" & mfrmDetail.Get��ǰ��ҩ����ʽ)
        End If
    Else
        If intAllFormat = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "��ҩ�ⷿ=" & strҩ�� & "|" & mcondition.lngҩ��ID, _
                "��ҩ����=" & mstr�������� & "|" & " IN (" & mstr����ID�� & ")", _
                "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"))
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "��ҩ�ⷿ=" & strҩ�� & "|" & mcondition.lngҩ��ID, _
                "��ҩ����=" & mstr�������� & "|" & " IN (" & mstr����ID�� & ")", _
                "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"), "ReportFormat=" & mfrmDetail.Get��ǰ��ҩ����ʽ)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub BillPrint_Wait()
    Dim rsTmp As New ADODB.Recordset
    Dim str��ʾ As String, str�� As String
    Dim strҩ�� As String, i As Long
    Dim n As Integer

    '�ⷿ����
    On Error GoTo errHandle
    gstrSQL = "Select ���� From ���ű� Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ǰҩ��������]", mcondition.lngҩ��ID)

    strҩ�� = rsTmp!���� & "|" & mcondition.lngҩ��ID

    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1342_1", Me, _
        "סԺҩ��=" & strҩ��, "סԺ����=" & mstr�������� & "|" & " IN (" & mstr����ID�� & ")", _
        "��ʼʱ��=" & mcondition.str��ʼʱ��, "����ʱ��=" & mcondition.str����ʱ��, 1)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckShortage(ByRef rsSendData As ADODB.Recordset, ByVal blnSendCheck As Boolean, Optional ByRef strMsg As String) As Boolean
    'ȱҩ���
    '1��blnSendCheck=False���Զ�ȱҩ���ʱȡ���ݵĿ���������ͻ��ܷ�ҩ�����Ƚ�
    '2��blnSendCheck=True����ҩ���ʱȡ��ǰ���ݿ�Ŀ���������ͻ��ܷ�ҩ�����Ƚ�
    
    Dim rsData As ADODB.Recordset
    Dim dblSum As Double
    Dim dblStock As Double
    Dim str��ǰҩƷ As String
    Dim blnTmp As Boolean       '�Ƿ����µ�ȱҩ
    Dim strƷ�� As String
    Dim IntCount As Integer
    
    blnTmp = True
    
    If mParams.IntCheckStock = 0 Then
        CheckShortage = True
        Exit Function
    End If
    
    rsSendData.Filter = "ִ��״̬=" & mState.��ҩ
    rsSendData.Sort = "ҩƷID,����,NO"

    With rsSendData
        Do While Not .EOF
            If str��ǰҩƷ <> !ҩƷid & ";" & !���� Then
                If blnSendCheck = True Then
                    dblSum = MediWork_GetMediRealAmount(mcondition.lngҩ��ID, Val(!ҩƷid), Val(!����))
                Else
                    dblSum = NVL(!�������, 0)
                End If
                
                str��ǰҩƷ = !ҩƷid & ";" & !����
            End If
            
            dblSum = dblSum - !��ҩ����
                
            If dblSum < 0 Then
                If strƷ�� <> !Ʒ�� Then
                    strƷ�� = !Ʒ��
                    
                    IntCount = IntCount + 1
                    If IntCount < 6 Then
                        strMsg = IIf(strMsg = "", strƷ��, strMsg & vbCrLf & strƷ��)
                    End If
                End If
                
                If !ִ��״̬ <> mState.ȱҩ Then
                    If mParams.blnȱҩ��� Then
                        !ִ��״̬ = mState.ȱҩ
                        !״̬ = "ȱҩ"
                        .Update
                    End If
                    blnTmp = False
                End If
            End If
            
            .MoveNext
        Loop
        
        rsSendData.Filter = ""
        rsSendData.Sort = ""
    End With
    
    If strMsg <> "" Then
        If IntCount > 5 Then strMsg = strMsg & vbCrLf & "��������" & IntCount - 5 & "��ȱҩҩƷ......"
    End If
    
    CheckShortage = blnTmp
End Function

Private Function CheckNotAudited(ByRef rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim bln�������� As Boolean
    Dim bln������ As Boolean
    
    On Error GoTo errHandle
    
    Call InitApplyforcredit
    
    CheckNotAudited = True
    bln�������� = True
    
    gstrSQL = "Select ���� As ������������ From ���˷������� Where ����id = [1] And ״̬ = 0"
    
    With rsData
        .Filter = "ִ��״̬=" & mState.��ҩ
        .Sort = "ҩƷID Asc"
        
        Do While Not .EOF
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ������������δ��˵ĵ���", rsData!����ID)
            
            If rsTmp.RecordCount > 0 Then
                bln�������� = False
                
                With mrsApplyforcredit
                    .AddNew
                    
                    !ִ��״̬ = rsData!ִ��״̬
                    !����ID = rsData!����ID
                    !�շ�Id = rsData!�շ�Id
                    !NO = rsData!NO
                    !ҩƷ���� = rsData!ҩƷ����
                    !���� = rsData!����
                    !���� = rsData!����
                    !������������ = zlStr.FormatEx(rsTmp!������������ / rsData!��װ, 5) & rsData!��λ
                    !���� = rsData!����
                    !�Ա� = rsData!�Ա�
                    !���� = rsData!����
                    !��ҩ���� = rsData!��ҩ����
                    !���� = rsData!����
                    !���˿��� = rsData!����
                End With

            End If
            
            .MoveNext
        Loop
    End With
    
    '�Ժ�����������ĵ��ݽ��д���
    If bln�������� = False Then
        Call frm���ŷ�ҩ���������嵥.ShowCard(Me, mrsApplyforcredit, bln������)
        
        '���Ӵ��巵���û��Ƿ����ִ�в���������ȡ�������ֹ��������
        CheckNotAudited = bln������
        If CheckNotAudited = False Then Exit Function
        
        '����ȡ�����͵ĵ��ݵ�ִ��״̬
        mrsApplyforcredit.Filter = "ִ��״̬ = 3"
        If mrsApplyforcredit.RecordCount > 0 Then
            Do While Not mrsApplyforcredit.EOF
                rsData.Filter = "�շ�ID = " & mrsApplyforcredit!�շ�Id
                rsData!ִ��״̬ = 3
                rsData.Update
                mrsApplyforcredit.MoveNext
            Loop
        End If
        
        rsData.Filter = ""
    End If
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDrugStock(ByVal rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Integer
    Dim lngҩƷid As Long
    
    CheckDrugStock = True
    With rsData
        .Filter = "ִ��״̬=" & mState.��ҩ
        .Sort = "ҩƷID Asc"
        
        Do While Not .EOF
            If lngҩƷid <> !ҩƷid Then
                If MediWork_CheckStorageStock(mcondition.lngҩ��ID, Val(!ҩƷid)) = False Then
                    MsgBox !Ʒ�� & "δ���ô洢�ⷿ�����ܷ�ҩ��", vbInformation, gstrSysName
                    CheckDrugStock = False
                    Exit Function
                End If
                    
                lngҩƷid = !ҩƷid
            End If
            .MoveNext
        Loop
    End With
    
    CheckDrugStock = True
End Function
Private Sub ClearData(ByVal intType As Integer)
    '����������ݣ���������б����ݣ������ϸ����
    
    ClearTreeView IIf(intType = mListType.��ҩ, 0, 1)
    ClearDetailList intType
End Sub

Private Sub ClearDetailList(ByVal intType As Integer)
    '�����ϸ�б�
    If intType <> mListType.��ҩ Then
        mfrmDetail.ClearList mListType.��ҩ
    Else
        mfrmDetail.ClearList mListType.��ҩ
    End If
End Sub

Private Sub ClearTreeView(ByVal intType As Integer)
    tvwList(intType).Nodes.Clear
    tvwList(intType).Tag = 1
    chkAll(intType).Value = 0
End Sub

Private Function DrugStoreWork_SendProc(ByVal rsData As ADODB.Recordset, ByVal StrCurDate As String) As Boolean
   '����ҩ����
    Dim lng����ID As Long
    Dim strID���δ� As String         '��ʽ���շ�ID,����|�շ�ID,����...
    Dim strID�� As String             '��ʽ���շ�ID,�շ�ID...
    Dim blnBeginTrans As Boolean
    Dim str��ҩ�� As String
    Dim str��ҩ�� As String
    Dim str�˲��� As String
    Dim blnUpdate As Boolean
    Dim strǩ����¼ As String
    Dim strSQL As String
    Dim arrSql As Variant
    Dim lngRow As Long
    Dim strFilter As String
    Dim blnIsCommit As Boolean        '�Ƿ��������ύ
    Dim strInputID  As String
    Dim rsSign As ADODB.Recordset     '���ڴ������ǩ��
    Dim strReserve As String
    
    arrSql = Array()
    
    '��ҩ��ǩ��
    TimerAuto.Enabled = False
    mstr����id = ""
    str��ҩ�� = ""
    If mParams.bln��ҩ��ǩ�� = True Then
        str��ҩ�� = zlDatabase.UserIdentify(Me, "��ҩ��ǩ��", glngSys, 1342, "")
        If str��ҩ�� = "" Then
            TimerAuto.Enabled = True
            Exit Function
        End If
    End If
    TimerAuto.Enabled = True
    
    'ȡ��ҩ��
    str��ҩ�� = mfrmDetail.Get��ǰ��ҩ��
    
    'ȡ�����
    str�˲��� = mfrmDetail.Get��ǰ�˲���
    
    '��¡��ҩ���ݼ����ڵ���ǩ����������ѭ����ֱ���÷�ҩ���ݼ�
    Set rsSign = rsData.Clone
    
    On Error GoTo errHandle
    
    '��ҩ��������ID������ҩ��
    With rsData
        .Filter = "ִ��״̬=" & mState.��ҩ
        
        '���밴����ID��ҩƷID����
        .Sort = "����ID Asc ,ҩƷID Asc"
        
        Do While Not .EOF
            If lng����ID = 0 Then
                lng����ID = !����ID
            End If
                
            '����ID��ͬʱ��
            If lng����ID = !����ID Then
                '���������ַ�������3950ʱ���ύ��������ַ���Ϊ4000��
'                If Len(strID���δ�) > 3950 Then
'                    gstrSQL = "Zl_ҩƷ�շ���¼_������ҩ("
'                    '�շ�ID�����δ�
'                    gstrSQL = gstrSQL & "'" & strID���δ� & "'"
'                    '�ⷿID
'                    gstrSQL = gstrSQL & "," & mcondition.lngҩ��ID
'                    '�����
'                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
'                    '�������
'                    gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
'                    '��ҩ��ʽ
'                    gstrSQL = gstrSQL & ",3"
'                    '��ҩ��
'                    gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
'                    '���ܷ�ҩ��
'                    gstrSQL = gstrSQL & "," & mcur���ܷ�ҩ��
'                    '����λ��
'                    gstrSQL = gstrSQL & "," & mParams.int����λ��
'                    '��ҩ��
'                    gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
'                     '��˴�����
'                    gstrSQL = gstrSQL & ",'" & str�˲��� & "'"
'                    gstrSQL = gstrSQL & ")"
'
'                    gcnOracle.BeginTrans
'                    blnBeginTrans = True
'                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-�����˷�ҩ")
'
'                    If mParams.bln��˻��۵� = True Then
'                        gstrSQL = "Zl_סԺ���ʼ�¼_��ҩ���("
'                        '�շ�ID��
'                        gstrSQL = gstrSQL & "'" & strID�� & "'"
'                        '����Ա���
'                        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
'                        '����Ա����
'                        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
'                        '���ʱ��
'                        gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
'                        gstrSQL = gstrSQL & ")"
'                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-סԺ�������")
'                    End If
'
'                    '��ҩ���ʾҽ�����ͨ��
'                    If strInputID <> "" And mParams.int��ҩʱ���ҽ�� = 1 Then
'                        gstrSQL = "Zl_��Һ��ҩ��¼_���("
'                        'ҽ��ID
'                        gstrSQL = gstrSQL & "'" & strInputID & "'"
'                        gstrSQL = gstrSQL & ")"
'
'                        Call zldatabase.ExecuteProcedure(gstrSQL, "�״�ִ��ҽ�����")
'                    End If
'
'                    gcnOracle.CommitTrans
'                    blnIsCommit = True
'                    blnBeginTrans = False
'                    blnUpdate = True
'                    strFilter = IIf(strFilter = "", "(����id=" & lng����ID & " and ִ��״̬=1)", strFilter & " or (����id=" & lng����ID & " and ִ��״̬=1)")
'                    lng����ID = 0
'                    strID���δ� = !�շ�Id & "," & NVL(!����, 0)
'                    strID�� = !�շ�Id
'                    strInputID = ""
'                    If NVL(!ҽ��id, 0) <> 0 And Not (!��� = "E" And !ִ�з��� = 1 And mblnIs��������) Then
'                        strInputID = !ҽ��id & ",3|"
'                    End If
'                Else
                    strID���δ� = IIf(strID���δ� = "", !�շ�Id & "," & NVL(!����, 0), strID���δ� & "|" & !�շ�Id & "," & NVL(!����, 0))
                    strID�� = IIf(strID�� = "", !�շ�Id, strID�� & "," & !�շ�Id)
                    If InStr(1, strInputID, !ҽ��id & ",3|") < 1 And NVL(!ҽ��id, 0) <> 0 And Not (!��� = "E" And !ִ�з��� = 1 And mblnIs��������) Then
                        strInputID = strInputID & !ҽ��id & ",3|"
                    End If
'                End If
            Else
                '�������ID��ͬ���ύ����
                gstrSQL = "Zl_ҩƷ�շ���¼_������ҩ("
                '�շ�ID�����δ�
                gstrSQL = gstrSQL & "'" & strID���δ� & "'"
                '�ⷿID
                gstrSQL = gstrSQL & "," & mcondition.lngҩ��ID
                '�����
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '�������
                gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                '��ҩ��ʽ
                gstrSQL = gstrSQL & ",3"
                '��ҩ��
                gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
                '���ܷ�ҩ��
                gstrSQL = gstrSQL & "," & mcur���ܷ�ҩ��
                '����λ��
                gstrSQL = gstrSQL & "," & mParams.int����λ��
                '��ҩ��
                gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
                '��˴�����
                gstrSQL = gstrSQL & ",'" & str�˲��� & "'"
                gstrSQL = gstrSQL & ")"
                
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
'                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-�����˷�ҩ")
                                    
                If mParams.bln��˻��۵� = True Then
                    gstrSQL = "Zl_סԺ���ʼ�¼_��ҩ���("
                    '�շ�ID��
                    gstrSQL = gstrSQL & "'" & strID�� & "'"
                    '����Ա���
                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                    '����Ա����
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '���ʱ��
                    gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
'                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-סԺ�������")
                End If
                
                '��ҩ���ʾҽ�����ͨ��
                If strInputID <> "" And mParams.int��ҩʱ���ҽ�� = 1 Then
                    gstrSQL = "Zl_��Һ��ҩ��¼_���("
                    'ҽ��ID
                    gstrSQL = gstrSQL & "'" & strInputID & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
'                    Call zldatabase.ExecuteProcedure(gstrSQL, "�״�ִ��ҽ�����")
                End If
                
                'ǩ��ʧ�ܣ���ҩ�����˳�
                If gblnESign���ŷ�ҩ = True And gblnESignUserStoped = False Then
                    mstr����id = IIf(mstr����id = "", "����ID <>" & lng����ID, mstr����id & " And ����ID <>" & lng����ID)
                    gstrSQL = Signature(rsSign, StrCurDate, str��ҩ��, lng����ID)
                    If gstrSQL = "" Then Exit Function
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                gcnOracle.BeginTrans
                blnBeginTrans = True
                
                For lngRow = 0 To UBound(arrSql)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-����ǩ��")
                Next
                
                gcnOracle.CommitTrans
                blnIsCommit = True
                blnBeginTrans = False
                blnUpdate = True
                strFilter = IIf(strFilter = "", "(����id=" & lng����ID & " and ִ��״̬=1)", strFilter & " or (����id=" & lng����ID & " and ִ��״̬=1)")
                lng����ID = !����ID
                strID���δ� = !�շ�Id & "," & NVL(!����, 0)
                strID�� = !�շ�Id
                arrSql = Array()
                strInputID = ""
                If NVL(!ҽ��id, 0) <> 0 And Not (!��� = "E" And !ִ�з��� = 1 And mblnIs��������) Then
                    strInputID = !ҽ��id & ",3|"
                End If
            End If
            
            .MoveNext
            
            '�������û�м�¼���Ҵ����ַ�����Ϊ�գ����ύ����
            If .EOF And strID���δ� <> "" Then
                gstrSQL = "Zl_ҩƷ�շ���¼_������ҩ("
                '�շ�ID�����δ�
                gstrSQL = gstrSQL & "'" & strID���δ� & "'"
                '�ⷿID
                gstrSQL = gstrSQL & "," & mcondition.lngҩ��ID
                '�����
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '�������
                gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                '��ҩ��ʽ
                gstrSQL = gstrSQL & ",3"
                '��ҩ��
                gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
                '���ܷ�ҩ��
                gstrSQL = gstrSQL & "," & mcur���ܷ�ҩ��
                '����λ��
                gstrSQL = gstrSQL & "," & mParams.int����λ��
                '��ҩ��
                gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
                '��˴�����
                gstrSQL = gstrSQL & ",'" & str�˲��� & "'"
                gstrSQL = gstrSQL & ")"

                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
'                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-�����˷�ҩ")
                                   
                If mParams.bln��˻��۵� = True Then
                    gstrSQL = "Zl_סԺ���ʼ�¼_��ҩ���("
                    '�շ�ID��
                    gstrSQL = gstrSQL & "'" & strID�� & "'"
                    '����Ա���
                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                    '����Ա����
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '���ʱ��
                    gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                    gstrSQL = gstrSQL & ")"
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
'                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-סԺ�������")
                End If
                
                '��ҩ���ʾҽ�����ͨ��
                If strInputID <> "" And mParams.int��ҩʱ���ҽ�� = 1 Then
                    gstrSQL = "Zl_��Һ��ҩ��¼_���("
                    'ҽ��ID
                    gstrSQL = gstrSQL & "'" & strInputID & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
'                    Call zldatabase.ExecuteProcedure(gstrSQL, "�״�ִ��ҽ�����")
                End If
                
                If gblnESign���ŷ�ҩ = True And gblnESignUserStoped = False Then
                    mstr����id = IIf(mstr����id = "", "����ID <>" & lng����ID, mstr����id & " And ����ID <>" & lng����ID)
                    gstrSQL = Signature(rsSign, StrCurDate, str��ҩ��, lng����ID)
                    If gstrSQL = "" Then Exit Function
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                gcnOracle.BeginTrans
                blnBeginTrans = True
                
                For lngRow = 0 To UBound(arrSql)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-����ǩ��")
                Next

                gcnOracle.CommitTrans
                blnIsCommit = True
                blnBeginTrans = False
            End If
        Loop
    End With
    
    '���÷�ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing Then
        mobjPlugIn.DrugSendBySumID mcondition.lngҩ��ID, mcur���ܷ�ҩ��, strReserve
    End If
    
    DrugStoreWork_SendProc = True
    Exit Function
errHandle:
    '����ѿ������񣬲���δ�ύ�������ʱ�ع�����
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
'        If strFilter <> "" Then
'            rsData.Filter = strFilter
'            Call Signature(rsData, StrCurDate, str��ҩ��)
'        End If
    End If
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '���ύ�����ݣ���ӡ�ύ�Ļ�������
    If blnIsCommit = True Then
        Call DrugStoreWork_PrintBill
    End If
End Function

Private Function Signature(ByVal rsData As Recordset, ByVal StrCurDate As String, ByVal str��ҩ�� As String, ByVal lng����ID As Long) As String
    Dim strǩ����¼ As String
    Dim strSQL As String
    Dim rstemp As Recordset
    Dim lngǩ��id As Long
    Dim str�շ�ID As String
    Dim lngRow As Long
    Dim arrSql As Variant
    Dim blnTrans As Boolean
    
    On Error GoTo errHandle
    
    arrSql = Array()
    rsData.Filter = "����id=" & lng����ID
    '����ǩ������
    If gblnESign���ŷ�ҩ = True And gblnESignUserStoped = False Then
        If rsData.RecordCount > 0 Then
            If GetSignatureRecoredGather(EsignTache.send, rsData, mcondition.lngҩ��ID, str��ҩ��, gstrUserName, StrCurDate, strǩ����¼) = False Then
                Exit Function
            End If
            
            If strǩ����¼ <> "" Then
                strSQL = "Select ҩƷǩ����¼_Id.Nextval ǩ��id From Dual"
                Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡǩ��id")
                
                lngǩ��id = Val(rstemp!ǩ��ID)
                
                str�շ�ID = Mid(Mid(strǩ����¼, 1, Len(strǩ����¼) - 1), InStrRev(Mid(strǩ����¼, 1, Len(strǩ����¼) - 1), "'") + 1)
                strǩ����¼ = Mid(Mid(strǩ����¼, 1, Len(strǩ����¼) - 1), 1, InStrRev(Mid(strǩ����¼, 1, Len(strǩ����¼) - 1), "'") - 1)
                
                strSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & "'" & str�շ�ID & "'," & lngǩ��id & ")"
                
'                For lngRow = 0 To UBound(Split(str�շ�ID, ","))
'                    If Len(strTemp & "," & Split(str�շ�ID, ",")(lngRow)) >= 4000 Then
'                        strSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & "'" & str�շ�ID & "'," & lngǩ��id & ")"
'
'                        ReDim Preserve arrSql(UBound(arrSql) + 1)
'                        arrSql(UBound(arrSql)) = strSQL
'
'                        strTemp = Split(str�շ�ID, ",")(lngRow)
'                    Else
'                        strTemp = IIf(strTemp = "", Split(str�շ�ID, ",")(lngRow), strTemp & "," & Split(str�շ�ID, ",")(lngRow))
'                    End If
'
'                    If lngRow = UBound(Split(str�շ�ID, ",")) And strǩ����¼ <> "" Then
'                        strSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & "'" & strTemp & "'," & lngǩ��id & ")"
'
'                        ReDim Preserve arrSql(UBound(arrSql) + 1)
'                        arrSql(UBound(arrSql)) = strSQL
'                    End If
'                Next
                
                Signature = strSQL
'                gcnOracle.BeginTrans
'                blnTrans = True
'                For lngRow = 0 To UBound(arrSql)
'                    Call zldatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-����ǩ��")
'                Next
'                gcnOracle.CommitTrans
                blnTrans = False
                rsData.Filter = mstr����id
            Else
                Signature = ""
                MsgBox "�Է�ҩ�˵���ǩ��ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    Exit Function
errHandle:
'    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function










Private Function DrugStoreWork_StayProc(ByVal StrCurDate As String) As Boolean
    '������������
    Dim rsData As ADODB.Recordset
    Dim Str�ڼ� As String
    Dim arrSql As Variant
    Dim lngRow As Long
    Dim int���淽ʽ As Integer
    
    'ȡ�������ݼ�
    Set rsData = mfrmDetail.GetStayRecord
    arrSql = Array()
    
    int���淽ʽ = Val(zlDatabase.GetPara("������������", glngSys, ģ���.ҩƷ����))
    With rsData
        Str�ڼ� = Format(StrCurDate, IIf(int���淽ʽ = 0, "yyyy", "yyyymm"))
        .Filter = ""
        
        Do While Not .EOF
            gstrSQL = "ZL_ҩƷ�����¼_INSERT("
            '�ڼ�
            gstrSQL = gstrSQL & "'" & Str�ڼ� & "'"
            '���ܷ�ҩ��
            gstrSQL = gstrSQL & "," & mcur���ܷ�ҩ��
            '�ⷿID
            gstrSQL = gstrSQL & "," & mcondition.lngҩ��ID
            'ҩƷID
            gstrSQL = gstrSQL & "," & !ҩƷid
            '����
            gstrSQL = gstrSQL & "," & !����
            '��������
            gstrSQL = gstrSQL & "," & !��������
            '���ۼ�
            gstrSQL = gstrSQL & "," & !����
            '������
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '��������
            gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
            '��ҩ����ID
            gstrSQL = gstrSQL & "," & !��ҩ����ID
            '�Ǽ�ʱ��
            gstrSQL = gstrSQL & ",To_Date('" & !����Ǽ�ʱ�� & "','yyyy-MM-dd hh24:mi:ss')"
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
                
            .MoveNext
        Loop
        
    End With
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For lngRow = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-��������")
    Next
    gcnOracle.CommitTrans
    
    DrugStoreWork_StayProc = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DrugStoreWork_CancelVerifyProc(ByVal StrCurDate As String) As Boolean
    '������������
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim int��˱�־ As Integer
    Dim bln�Ƿ�����ҩ As Boolean
    Dim str������� As String
    Dim lngPre����id As Long
    Dim strҩƷID As String
    Dim strPreNo As String
    Dim lngPre������� As Long
    Dim dblSum As Double
    Dim rsData As ADODB.Recordset
    Dim arrSql As Variant
    Dim blnBeginTrans As Boolean
    Dim strWriteOffInfo As String
    Dim strReturnInfo As String
    Dim strReserve As String
    
    arrSql = Array()
    
    'ǰ�������ǻ������ʼ�¼һ����ҩ
    If mParams.bln���ܷ�ҩ = True Then
        'ȡ�������ݼ�
        Set rsData = mfrmDetail.GetChargeOffRecord
    
        If rsData.State <> 0 Then
            rsData.Filter = "ִ�б�־=1"
            rsData.Sort = "ҩƷid,No,����id,�շ�id"
            If rsData.RecordCount > 0 Then
                With rsData
                    '��ʼ��ҽ������
                    gclsInsure.InitOracle gcnOracle
                    Do While Not .EOF
                        If !��˱�־ = 1 And !�������� <> 0 Then
                            If IsOutPatient(mstrPrivs, !����, !NO, 2, 2) = False Then Exit Function
                            If IsReceiptBalance_Charge(1, mstrPrivs, !����, !NO, !�������, 2, 2) = False Then Exit Function
                        End If
                
                        If !��˱�־ <> 0 Then
                            If lngPre����id <> !����ID Then
                                '�������ʼ�¼����
                                gstrSQL = "zl_���˷�������_Audit("
                                '����ID
                                gstrSQL = gstrSQL & !����ID
                                '����ʱ��
                                gstrSQL = gstrSQL & ",To_Date('" & !����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                                '�����
                                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                                '���ʱ��
                                gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                                '��˱�־
                                gstrSQL = gstrSQL & "," & !��˱�־
                                gstrSQL = gstrSQL & ")"

                                ReDim Preserve arrSql(UBound(arrSql) + 1)
                                arrSql(UBound(arrSql)) = gstrSQL
                
                                lngPre����id = !����ID
                                
                                '��¼��ǰ������˵ļ�¼������ʱ��Ͳ���ID�����ڸ���������Ϣ�˵�
                                If strWriteOffInfo = "" Then
                                    strWriteOffInfo = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & !����ID
                                ElseIf InStr(strWriteOffInfo & "|", Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & !����ID & "|") = 0 Then
                                    strWriteOffInfo = strWriteOffInfo & "|" & Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & !����ID
                                End If
                                
                            End If
                        End If
                        
                        '��ҩ����
                        If !��˱�־ = 1 And !�������� <> 0 Then
                            gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
                            '�շ�ID
                            gstrSQL = gstrSQL & !�շ�Id
                            '�����
                            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                            '���ʱ��
                            gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                            '����
                            gstrSQL = gstrSQL & "," & IIf(IsNull(!����), "NULL", IIf(Mid(!����, 1, 1) = "(", "NULL", "'" & Mid(!����, 1, 8) & "'"))
                            'Ч��
                            gstrSQL = gstrSQL & "," & IIf(IsNull(!Ч��), "NULL", IIf(!Ч�� = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
                            '����
                            gstrSQL = gstrSQL & "," & IIf(IsNull(!����), "NULL", "'" & !���� & "'")
                            '��ҩ����
                            gstrSQL = gstrSQL & "," & !��������
                            '��ҩ�ⷿ
                            gstrSQL = gstrSQL & ",NULL"
                            '��ҩ��
                            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                            '����λ��
                            gstrSQL = gstrSQL & "," & mParams.int����λ��
                            '����
                            gstrSQL = gstrSQL & ",2"
                            '���ܷ�ҩ��
                            gstrSQL = gstrSQL & "," & mcur���ܷ�ҩ��
                            gstrSQL = gstrSQL & ")"

                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = gstrSQL
                
                            bln�Ƿ�����ҩ = True
                            
                            If InStr("," & strҩƷID & ",", "," & !ҩƷid & ",") = 0 Then
                                strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & !ҩƷid
                            End If
                            
                            strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(!�շ�Id) & "," & Val(!��������)
                        
                            '���ʴ���
                            strPreNo = !NO
                            lngPre������� = !�������
                            dblSum = dblSum + !��������
                            
                            .MoveNext
                            If .EOF Then
                                .MovePrevious
                                str������� = !������� & ":" & dblSum
                
                                gstrSQL = "ZL_סԺ���ʼ�¼_Delete("
                                'NO
                                gstrSQL = gstrSQL & "'" & !NO & "'"
                                '��ţ�������
                                gstrSQL = gstrSQL & ",'" & str������� & "'"
                                '����Ա���
                                gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                                '����Ա����
                                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                                '��¼����
                                gstrSQL = gstrSQL & "," & !��¼����
                                '����״̬
                                gstrSQL = gstrSQL & ",1"
                                gstrSQL = gstrSQL & ")"

                                ReDim Preserve arrSql(UBound(arrSql) + 1)
                                arrSql(UBound(arrSql)) = gstrSQL
                
                                'ҽ������
                                If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                                    MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                                    MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                                            "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                                End If
                                .MoveNext
                            Else
                                If strPreNo <> !NO Or (strPreNo = !NO And lngPre������� <> !�������) Then
                                    .MovePrevious
                                    str������� = !������� & ":" & dblSum
                                    
                                    gstrSQL = "ZL_סԺ���ʼ�¼_Delete("
                                    'NO
                                    gstrSQL = gstrSQL & "'" & !NO & "'"
                                    '��ţ�������
                                    gstrSQL = gstrSQL & ",'" & str������� & "'"
                                    '����Ա���
                                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                                    '����Ա����
                                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                                    '��¼����
                                    gstrSQL = gstrSQL & "," & !��¼����
                                    '����״̬
                                    gstrSQL = gstrSQL & ",1"
                                    gstrSQL = gstrSQL & ")"

                                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                                    arrSql(UBound(arrSql)) = gstrSQL
                    
                                    'ҽ������
                                    If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                                        MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                                        MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                                                "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                                    End If
                                    
                                    dblSum = 0
                                    .MoveNext
                                End If
                            End If
                            .MovePrevious
                        End If
                        .MoveNext
                    Loop
                End With
                
                '���д�����ҩ��������
                gcnOracle.BeginTrans
                blnBeginTrans = True
                
                For i = 0 To UBound(arrSql)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "DrugStoreWork_CancelVerifyProc")
                Next
            
                'ҽ�������������ϴ�������ʱ�ϴ�
                If strMCNO <> "" Then
                    arrMCRec = Split(strMCNO, "|")
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                gcnOracle.RollbackTrans
                                GoTo errHandle
                            End If
                        End If
                    Next
                End If
                
                gcnOracle.CommitTrans
                blnBeginTrans = False
                
                'ҽ�������������ϴ�����ɺ��ϴ�
                If strMCNO <> "" Then
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                MsgBox "����""" & CStr(arrMCPar(0)) & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    'ɾ����Ϣ��¼�����Ѿ���˹�����Ϣ��¼
    If strWriteOffInfo <> "" And Not mobjMipModule Is Nothing Then
        If Not mrsReceiveMsg Is Nothing Then
            If mrsReceiveMsg.RecordCount > 0 Then
                With mrsReceiveMsg
                    .MoveFirst
                    Do While Not .EOF
                        If InStr(strWriteOffInfo & "|", !����ʱ�� & "," & !����ID & "|") > 0 Then
                            .Delete adAffectCurrent
                        End If
                        
                        .MoveNext
                    Loop
                End With
                '������Ϣ�˵�
                Call SetMessageBar(mrsReceiveMsg)
            End If
        End If
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ�����ҩ Then
        mobjPlugIn.DrugReturnByID mlngҩ��ID, strReturnInfo, CDate(StrCurDate), strReserve
    End If
        
    DrugStoreWork_CancelVerifyProc = True
    Exit Function
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FindRow()
    Dim StrFind As String
    
    If tbcDetail.Selected.Index <> mListType.��ҩ And tbcDetail.Selected.Index <> mListType.��ҩ Then Exit Sub
    
    TimerAuto.Enabled = False
    StrFind = Frm���ŷ�ҩ��λ.ShowMe(mcondition.lngҩ��ID, Me, mstrPrivs)
    
    If StrFind <> "" Then
        mfrmDetail.FindRecord tbcDetail.Selected.Index, StrFind
    End If
    
    TimerAuto.Enabled = True
End Sub

Private Sub FindRowNext()
    If tbcDetail.Selected.Index <> mListType.��ҩ And tbcDetail.Selected.Index <> mListType.��ҩ Then Exit Sub
    
    mfrmDetail.FindRecord tbcDetail.Selected.Index
End Sub

Private Sub GetCondition()
    '��������
    Dim dteTime As Date
    Dim n As Integer
    
    dteTime = zlDatabase.Currentdate
    
    With mcondition
        'ҩ��ID
        .lngҩ��ID = cbo��ҩҩ��.ItemData(cbo��ҩҩ��.ListIndex)
        
        'ʱ�䷶Χ
        Select Case cboʱ�䷶Χ.ListIndex
            Case mTimeRange.����
                .str��ʼʱ�� = Format(dteTime, "yyyy-mm-dd") & " 00:00:00"
                .str����ʱ�� = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
            Case mTimeRange.������
                .str��ʼʱ�� = Format(DateAdd("d", -1, dteTime), "yyyy-mm-dd") & " 00:00:00"
                .str����ʱ�� = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
            Case mTimeRange.������
                .str��ʼʱ�� = Format(DateAdd("d", -2, dteTime), "yyyy-mm-dd") & " 00:00:00"
                .str����ʱ�� = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
            Case mTimeRange.ָ��ʱ�䷶Χ
                .str��ʼʱ�� = Format(Dtp��ʼʱ��.Value, "yyyy-mm-dd hh:mm:ss")
                .str����ʱ�� = Format(Dtp����ʱ��.Value, "yyyy-mm-dd hh:mm:ss")
            Case Else
                .str��ʼʱ�� = Format(dteTime, "yyyy-mm-dd") & " 00:00:00"
                .str����ʱ�� = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
        End Select
        
        '¼����Ϣ
        .strסԺ�� = ""
        .str���� = ""
        .str���� = ""
        .strNo = ""
        .lng����ID = -1
        .str���￨ = ""
        .str��ҩ�� = ""
        .cur��ҩ�� = 0
        .lng��ҩ����ID = -1
        .strIC�� = ""
        
        If Trim(txtInput.Text) <> "" Then
            Select Case Val(lblPatiInputType.Tag)
                Case mInputType.סԺ��
                    If InStr(txtInput.Text, "-") > 0 Then
                        .strסԺ�� = Mid(Trim(txtInput.Text), 1, InStr(txtInput.Text, "-") - 1)
                    Else
                        .strסԺ�� = Trim(txtInput.Text)
                    End If
                Case mInputType.����
'                    If mblnCard = True Then
'                        .lng����ID = Val(txtInput.Tag)
'                    Else
'                        .str���� = Trim(txtInput.Text)
'                    End If
                    .lng����ID = Val(txtInput.Tag)
                Case mInputType.����
                    '���ڴ��Ų�Ψһ��תΪ�ò���ID����ѯ
                    .lng����ID = Val(txtInput.Tag)
                    
'                    If InStr(TxtInput.Text, "-") > 0 Then
'                        .str���� = Mid(Trim(TxtInput.Text), 1, InStr(TxtInput.Text, "-") - 1)
'                    Else
'                        .str���� = Trim(TxtInput.Text)
'                    End If
                Case mInputType.NO
                    If InStr(txtInput.Text, "-") > 0 Then
                        .strNo = Mid(Trim(txtInput.Text), 1, InStr(txtInput.Text, "-") - 1)
                    Else
                        .strNo = Trim(txtInput.Text)
                    End If
                Case mInputType.����ID
                    If InStr(txtInput.Text, "-") > 0 Then
                        .lng����ID = Mid(Trim(txtInput.Text), 1, InStr(txtInput.Text, "-") - 1)
                    Else
                        .lng����ID = Val(Trim(txtInput.Text))
                    End If
                Case mInputType.��ҩ��
                    .str��ҩ�� = Trim(txtInput.Text)
                Case mInputType.��ҩ��
                    .cur��ҩ�� = Val(Trim(txtInput.Text))
                Case mInputType.��ҩ����
                    .lng��ҩ����ID = Val(txtInput.Tag)
                Case mInputType.IC��
                    .lng����ID = Val(txtInput.Tag)
                Case Else
                    '���������ѿ������ز���ID
                    .lng����ID = Val(txtInput.Tag)
            End Select
        End If
        
        '��ҩ����
        '0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
        If chkSend(0).Value = 1 And chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
            .int��ҩ���� = 0
        ElseIf chkSend(0).Value = 1 And chkSend(2).Value = 1 Then
            .int��ҩ���� = 1
        ElseIf chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
            .int��ҩ���� = 3
        ElseIf chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
            .int��ҩ���� = 6
        ElseIf chkSend(0).Value = 1 Then
            .int��ҩ���� = 5
        ElseIf chkSend(1).Value = 1 Then
            .int��ҩ���� = 2
        ElseIf chkSend(2).Value = 1 Then
            .int��ҩ���� = 4
        End If
        
        '�Զ��巢ҩ����
        .str������ҩ���� = ""
        If mblnExistOtherSendType = True Then
            For n = 0 To chkSendType.UBound
                If chkSendType(n).Value = 1 Then
                    .str������ҩ���� = IIf(.str������ҩ���� = "", "", .str������ҩ���� & ",") & chkSendType(n).Caption
                End If
            Next
        End If
        
        '��ҩ;��
        If Trim(txt��ҩ;��.Text) = "" Or InStr(Trim(txt��ҩ;��.Text), "���и�ҩ;��") > 0 Then
            .str��ҩ;�� = ""
        Else
            .str��ҩ;�� = Trim(txt��ҩ;��.Text)
        End If
        
        'ҩƷ����
        If Trim(txtҩƷ����.Text) = "" Or InStr(Trim(txtҩƷ����.Text), "����ҩƷ����") > 0 Then
            .strҩƷ���� = ""
        Else
            .strҩƷ���� = Trim(txtҩƷ����.Text)
        End If
        
        '����Χ
        If Me.opt��Χ(1).Value = True Then
            .int����Χ = 1
        ElseIf Me.opt��Χ(2).Value = True Then
            .int����Χ = 2
        Else
            .int����Χ = 0
        End If
        
        'ҽ������
        .intҽ������ = 0
        If Cboҽ������.ListIndex <> -1 Then .intҽ������ = Cboҽ������.ListIndex
                
        '��������
        If chkType(0).Value = 1 And chkType(1).Value = 1 Then
            .int�������� = 2
        ElseIf chkType(1).Value = 1 Then
            .int�������� = 1
        Else
            .int�������� = 0
        End If
        
        '����ģʽ
        .int����ģʽ = mParams.int����ģʽ
        
        '������
        .str������ = mParams.str������
        
        '��ҩ����
        .bln��ʾ��ҩ�������� = mParams.bln��������
        
        '���й��̵���
        .bln��ʾ������ҩ���� = mParams.bln���̵���
        
        '��ҩ/��ҩ��
        .bln��ʾ��ҩ��ҩ�� = False
    End With
End Sub

Private Sub GetPrivs()
    'Ȩ��
    mPrives.bln����ҩ�� = IsInString(mstrPrivs, "����ҩ��", ";")
    mPrives.bln��ҩ = IsInString(mstrPrivs, "��ҩ", ";")
    mPrives.bln��ҩ = IsInString(mstrPrivs, "��ҩ", ";")
    mPrives.bln������ҩ���Ĵ��� = IsInString(mstrPrivs, "������ҩ���Ĵ���", ";")
    mPrives.bln���˽��ʴ��� = IsInString(mstrPrivs, "���˽��ʴ���", ";")
    mPrives.bln���˳�Ժ���˴��� = IsInString(mstrPrivs, "���˳�Ժ���˴���", ";")
    mPrives.blnȱҩ���� = IsInString(mstrPrivs, "ȱҩ����", ";")
    mPrives.bln�ܷ� = IsInString(mstrPrivs, "�ܷ�", ";")
    mPrives.bln������ҩ��� = IsInString(mstrPrivs, "������ҩ���", ";")
    mPrives.blnҽ����ѯ = IsInString(mstrPrivs, "ҽ����ѯ", ";")
    mPrives.bln�������� = IsInString(mstrPrivs, "��������", ";")
    mPrives.bln��ҩ���� = IsInString(mstrPrivs, "��ҩ����", ";")
    mPrives.bln�޸��������� = IsInString(mstrPrivs, "�޸���������", ";")
    mPrives.blnֹͣ��ҩ = IsInString(mstrPrivs, "ֹͣ��ҩ", ";")
    mPrives.bln�ָ���ҩ = IsInString(mstrPrivs, "�ָ���ҩ", ";")
End Sub

Private Sub GetDeptListRecord(ByVal rsData As ADODB.Recordset)
    Set mrsDeptList = New ADODB.Recordset
    
    With mrsDeptList
        If .State = 1 Then .Close
        
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable                  '��ҩ����ID
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable                  '��ҩ��������
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable                  'ҩƷ�շ���¼NO��
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable                  'ҩƷ�շ���¼ID
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable                  'ҩƷ�շ���¼ҩƷID
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If mParams.bln������ And tbcDetail.Selected.Index = mListType.��ҩ Then
            rsData.Filter = "(�����=1 and ���id<>0) or ���id=0"
        Else
            rsData.Filter = ""
        End If
        
        Do While Not rsData.EOF
            .AddNew
            !����ID = rsData!Id
            !�������� = rsData!��������
            !NO = rsData!NO
            !�շ�Id = rsData!�շ�Id
            !ҩƷid = rsData!ҩƷid
            !����ID = rsData!����ID
            !�������� = rsData!����
            !ִ��״̬ = 1
            
            .Update
            
            rsData.MoveNext
        Loop
    End With
End Sub
Private Sub GetSendDeptTreeView(ByRef rsData As ADODB.Recordset)
    'ˢ�´���ҩ��������
    Dim objNode As Node
    Dim objItem As listItem
    Dim lng��ǰ���� As Long
    Dim str��ǰ���� As String
    Dim str��ǰ��ҩ�� As String
    Dim str��ǰ����Key As String
    Dim lng��ǰ����ID As Long
    Dim str��ǰ�������� As String
    Dim str��ǰNO As String
    Dim int����ҩƷ�� As Integer
    Dim lng��ǰҩƷ As Long
    Dim strType As String
    Dim arr���� As Variant
    Dim i As Integer
    Dim j As Integer
    Dim count As Integer
    

    If rsData.EOF Then
        Set objNode = tvwList(mDeptType.��ҩ).Nodes.Add(, , "_������", "δ�ҵ����������ļ�¼")
        tvwList(mDeptType.��ҩ).Checkboxes = False
        tvwList(mDeptType.��ҩ).Tag = "0"
        chkAll(mDeptType.��ҩ).Enabled = False

        mfrmDetail.ClearList mListType.��ҩ
        Exit Sub
    End If
    
    '�������ݼ��Ľ����֯���������ң�����ҩƷ���ࣩ����ҩ�ţ�������ҩ����û���⼶�������ˡ����ݺ���ʾ����
    
    chkAll(mDeptType.��ҩ).Enabled = True
    tvwList(mDeptType.��ҩ).Checkboxes = True
    arr���� = Array()
    With tvwList(mDeptType.��ҩ)
        If Not rsData.EOF Then
            '��¼���п�������
            rsData.Sort = "��������,ID"
            
            If mParams.blnSort Then
                Set rstemp = New Recordset
                With rstemp
                    If .State = 1 Then .Close
                    .Fields.Append "ID", adDouble, 18, adFldIsNullable
                    .Fields.Append "��������", adLongVarChar, 40, adFldIsNullable
                    .Fields.Append "����ʱ��", adLongVarChar, 40, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                End With
            End If
            
            Do While Not rsData.EOF
                If lng��ǰ���� <> rsData!Id Then
                    lng��ǰ���� = rsData!Id
                    If mParams.blnSort Then
                        rstemp.AddNew
                        rstemp!Id = rsData!Id
                        rstemp!�������� = rsData!��������
                        rstemp!����ʱ�� = Format(rsData!��������, "yyyy-mm-dd hh:mm:ss")
                    Else
                        ReDim Preserve arr����(UBound(arr����) + 1)
                        arr����(UBound(arr����)) = lng��ǰ���� & "|" & rsData!��������
                    End If
                End If
                rsData.MoveNext
            Loop
            
            If mParams.blnSort Then
                rstemp.Sort = "����ʱ��"
                rstemp.MoveFirst
                Do While Not rstemp.EOF
                    ReDim Preserve arr����(UBound(arr����) + 1)
                    arr����(UBound(arr����)) = rstemp!Id & "|" & rstemp!��������
                    rstemp.MoveNext
                Loop
            End If
            
            
            '��������֯��������
            For i = 0 To UBound(arr����)
                If mParams.bln������ Then
                    rsData.Filter = "(�����=1 and ���id<>0 and ID= '" & Split(arr����(i), "|")(0) & "') or (���id=0 and ID=' " & Split(arr����(i), "|")(0) & "')"
                Else
                    rsData.Filter = "ID= '" & Split(arr����(i), "|")(0) & "' "
                End If
                
                '���㵱ǰ����ҩƷ����
                rsData.Sort = "ҩƷID"
                lng��ǰҩƷ = 0
                int����ҩƷ�� = 0
                Do While Not rsData.EOF
                    If lng��ǰҩƷ <> rsData!ҩƷid Then
                        int����ҩƷ�� = int����ҩƷ�� + 1
                        lng��ǰҩƷ = rsData!ҩƷid
                    End If
                    rsData.MoveNext
                Loop
                
                Set objNode = .Nodes.Add(, , "D_" & Split(arr����(i), "|")(0), Split(arr����(i), "|")(1) & "��" & int����ҩƷ�� & "��ҩƷ������", 1)
                objNode.Expanded = False
                
                If mParams.blnOnlyShowDept = False Then
                    '�ȹ�������ҩ�ŵļ�¼
                    If mParams.bln������ Then
                        rsData.Filter = "(�����=1 and ���id<>0 And ��ҩ��<>0 and ID= '" & Split(arr����(i), "|")(0) & "') or (���id=0 And ��ҩ��<>0 and ID=' " & Split(arr����(i), "|")(0) & "')"
                    Else
                        rsData.Filter = "ID= '" & Split(arr����(i), "|")(0) & "' And ��ҩ��<>0"
                    End If
                    
                    If mParams.int�������� = 1 Then
                        rsData.Sort = "��ҩ��,����,����ID,Ӥ����,NO"
                    Else
                        rsData.Sort = "��ҩ��,����,����,����ID,Ӥ����,NO"
                    End If
                    
                    str��ǰ��ҩ�� = ""
                    str��ǰ����Key = ""
                    str��ǰNO = ""
                    Do While Not rsData.EOF
                        If str��ǰ��ҩ�� <> rsData!��ҩ�� Then
                            str��ǰ��ҩ�� = rsData!��ҩ��
                            str��ǰ����Key = ""        '��ͬ��ҩ�ſ��ܴ�����ͬ�Ĳ��ˣ�����ҩ�Ų�ͬʱ����ʼ������ϢҪ���
                            
                            Set objNode = .Nodes.Add("D_" & Split(arr����(i), "|")(0), 4, "R_" & Split(arr����(i), "|")(0) & str��ǰ��ҩ��, str��ǰ��ҩ��, 2)
                            objNode.Expanded = False
                        End If
                        
                        If str��ǰ����Key & lng��ǰ����ID <> rsData!���� & "(" & IIf(IsNull(rsData!����), "", rsData!���� & "�� ") & rsData!�Ա� & " " & rsData!���� & ")" & rsData!����ID Then
                            If IIf(IsNull(rsData!����), "", rsData!����) <> "" Then
                                str��ǰ����Key = rsData!���� & "(" & rsData!���� & "�� " & rsData!�Ա� & " " & rsData!���� & ")"
                            Else
                                str��ǰ����Key = rsData!���� & "(" & rsData!�Ա� & " " & rsData!���� & ")"
                            End If
                            lng��ǰ����ID = rsData!����ID
                            
                            Set objNode = .Nodes.Add("R_" & Split(arr����(i), "|")(0) & str��ǰ��ҩ��, 4, "P_" & Split(arr����(i), "|")(0) & str��ǰ��ҩ�� & str��ǰ����Key & rsData!����ID, str��ǰ����Key, 3)
                            objNode.ForeColor = IIf(IsNull(rsData!��ɫ), vbBlack, rsData!��ɫ)
                            objNode.Expanded = False
                        End If
                        
                        If ((str��ǰNO <> rsData!NO) Or (str��ǰNO = rsData!NO And str��ǰ�������� <> rsData!����)) Then
                            str��ǰNO = rsData!NO
                            str��ǰ�������� = rsData!����  '��������ĸ�׺�Ӥ��Ϊͬһ�ŵ���NO�����
                            
                            strType = IIf(NVL(rsData!ҽ�����, 0) = 0, IIf(rsData!�����־ = 1 Or rsData!�����־ = 4, "������ʵ�", IIf(rsData!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(rsData!����) = True, "סԺ���ʵ�", IIf(rsData!���� Like "0*", "����", IIf(rsData!���� Like "1*", "����", "���ʱ�"))))
                            strType = strType & " " & Format(rsData!��������, "mm-dd hh:mm:ss")
                            
                            Set objNode = .Nodes.Add("P_" & Split(arr����(i), "|")(0) & str��ǰ��ҩ�� & str��ǰ����Key & rsData!����ID, 4, "N" & str��ǰ����Key & rsData!����ID & "_" & str��ǰNO & "_" & rsData!����, str��ǰNO & "(" & strType & ")", 4)
                            objNode.Expanded = False
                            objNode.Tag = rsData!NO & "|" & rsData!����ID & "|" & rsData!����
                            If rsData!�ܷ� = 1 Then
                                objNode.ForeColor = vbRed
                                objNode.Text = objNode.Text & "(�Ѿܷ�)"
                            End If
                        End If
                        
                        rsData.MoveNext
                    Loop
                    
                    '��������ҩ�ŵļ�¼������ҩ�ž�û����ҩ���⼶��
                    If mParams.bln������ Then
                        rsData.Filter = "(�����=1 and ���id<>0 And ��ҩ��=0 and ID= '" & Split(arr����(i), "|")(0) & "') or (���id=0 And ��ҩ��=0 and ID=' " & Split(arr����(i), "|")(0) & "')"
                    Else
                        rsData.Filter = "ID= '" & Split(arr����(i), "|")(0) & "' And ��ҩ��=0"
                    End If
                    If mParams.int�������� = 1 Then
                        rsData.Sort = "����,����ID,Ӥ����,NO"
                    Else
                        rsData.Sort = "����,����,����ID,Ӥ����,NO"
                    End If
                    str��ǰ����Key = ""
                    str��ǰNO = ""
                    str��ǰ�������� = ""
                    Do While Not rsData.EOF
                        If str��ǰ����Key & lng��ǰ����ID <> rsData!���� & "(" & IIf(IsNull(rsData!����), "", rsData!���� & "�� ") & rsData!�Ա� & " " & rsData!���� & ")" & rsData!����ID Then
                            If IIf(IsNull(rsData!����), "", rsData!����) <> "" Then
                                str��ǰ����Key = rsData!���� & "(" & rsData!���� & "�� " & rsData!�Ա� & " " & rsData!���� & ")"
                            Else
                                str��ǰ����Key = rsData!���� & "(" & rsData!�Ա� & " " & rsData!���� & ")"
                            End If
                            lng��ǰ����ID = rsData!����ID
                            
                            Set objNode = .Nodes.Add("D_" & Split(arr����(i), "|")(0), 4, "P_" & Split(arr����(i), "|")(0) & str��ǰ����Key & rsData!����ID, str��ǰ����Key, 3)
                            objNode.ForeColor = IIf(IsNull(rsData!��ɫ), vbBlack, rsData!��ɫ)
                            objNode.Expanded = False
                        End If
                        
                        If ((str��ǰNO <> rsData!NO) Or (str��ǰNO = rsData!NO And str��ǰ�������� <> rsData!����)) Then
                            str��ǰNO = rsData!NO
                            str��ǰ�������� = rsData!����  '��������ĸ�׺�Ӥ��Ϊͬһ�ŵ���NO�����
                            
                            strType = IIf(NVL(rsData!ҽ�����, 0) = 0, IIf(rsData!�����־ = 1 Or rsData!�����־ = 4, "������ʵ�", IIf(rsData!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(rsData!����) = True, "סԺ���ʵ�", IIf(rsData!���� Like "0*", "����", IIf(rsData!���� Like "1*", "����", "���ʱ�"))))
                            strType = strType & " " & Format(rsData!��������, "mm-dd hh:mm:ss")
                            Set objNode = .Nodes.Add("P_" & Split(arr����(i), "|")(0) & str��ǰ����Key & rsData!����ID, 4, "N" & str��ǰ����Key & rsData!����ID & "_" & str��ǰNO & "_" & rsData!����, str��ǰNO & "(" & strType & ")", 4)
                            objNode.Expanded = False
                            objNode.Tag = rsData!NO & "|" & rsData!����ID & "|" & rsData!����
                            If rsData!�ܷ� = 1 Then
                                objNode.ForeColor = vbRed
                                objNode.Text = objNode.Text & "(�Ѿܷ�)"
                            End If
                        End If
                        
                        rsData.MoveNext
                    Loop
                End If
            Next
        End If
    End With
End Sub
Private Sub GetReturnDeptTreeView(ByRef rsData As ADODB.Recordset)
    'ˢ����ҩ��������
    Dim objNode As Node
    Dim objItem As listItem
    Dim lng����ID As Long
    Dim str��ǰ����Key As String
    Dim lng��ǰ����ID As Long
    Dim str��ǰ�������� As String
    Dim str��ǰNO As String
    Dim strType As String
    
    Dim arr���� As Variant
    Dim i As Integer
    Dim j As Integer
    Dim count As Integer

    If rsData.EOF Then
        Set objNode = tvwList(mDeptType.��ҩ).Nodes.Add(, , "_������", "δ�ҵ����������ļ�¼")
        tvwList(mDeptType.��ҩ).Checkboxes = False
        tvwList(mDeptType.��ҩ).Tag = "0"
        chkAll(mDeptType.��ҩ).Enabled = False

        mfrmDetail.ClearList mListType.��ҩ
        Exit Sub
    End If
    
    '�������ݼ��Ľ����֯�������ַ�ʽ��
    '1������ҩ�š����ң�����ҩƷ���ࣩ�����ˡ����ݺ���ʾ������
    '2�������ң�����ҩƷ���ࣩ�����ˡ����ݺ���ʾ������
    
    chkAll(mDeptType.��ҩ).Enabled = True
    tvwList(mDeptType.��ҩ).Checkboxes = True
    arr���� = Array()
    With tvwList(mDeptType.��ҩ)
        If Not rsData.EOF Then
            '��¼���п�������
            rsData.Sort = "��������,ID"
            Do While Not rsData.EOF
                If lng����ID <> rsData!Id Then
                    ReDim Preserve arr����(UBound(arr����) + 1)
                    lng����ID = rsData!Id
                    arr����(UBound(arr����)) = rsData!Id & "|" & rsData!��������
                End If
                rsData.MoveNext
            Loop
    
            '��������֯��������
            For i = 0 To UBound(arr����)
                rsData.Filter = "ID= '" & Split(arr����(i), "|")(0) & "' "
                
                '���㵱ǰ����ҩƷ����
                rsData.Sort = "ҩƷID"
                
                Set objNode = .Nodes.Add(, , "D_" & Split(arr����(i), "|")(0), Split(arr����(i), "|")(1), 1)
                objNode.Expanded = False
                
                If mParams.blnOnlyShowDept = False Then
                    rsData.Filter = "ID= '" & Split(arr����(i), "|")(0) & "'"
                    If mParams.int�������� = 1 Then
                        rsData.Sort = "����,����ID,Ӥ����,NO"
                    Else
                        rsData.Sort = "����,����,����ID,Ӥ����,NO"
                    End If
                    str��ǰ����Key = ""
                    str��ǰNO = ""
                    str��ǰ�������� = ""
                    Do While Not rsData.EOF
                        If str��ǰ����Key & lng��ǰ����ID <> rsData!���� & "(" & IIf(IsNull(rsData!����), "", rsData!���� & "�� ") & rsData!�Ա� & " " & rsData!���� & ")" & rsData!����ID Then
                            If IIf(IsNull(rsData!����), "", rsData!����) <> "" Then
                                str��ǰ����Key = rsData!���� & "(" & rsData!���� & "�� " & rsData!�Ա� & " " & rsData!���� & ")"
                            Else
                                str��ǰ����Key = rsData!���� & "(" & rsData!�Ա� & " " & rsData!���� & ")"
                            End If
                            lng��ǰ����ID = rsData!����ID
                            
                            Set objNode = .Nodes.Add("D_" & Split(arr����(i), "|")(0), 4, "P_" & Split(arr����(i), "|")(0) & str��ǰ����Key & rsData!����ID, str��ǰ����Key, 3)
                            objNode.ForeColor = IIf(IsNull(rsData!��ɫ), vbBlack, rsData!��ɫ)
                            objNode.Expanded = False
                        End If
                        
                        If (str��ǰNO <> rsData!NO) Or (str��ǰNO = rsData!NO And str��ǰ�������� <> rsData!����) Then
                            str��ǰNO = rsData!NO
                            str��ǰ�������� = rsData!����  '��������ĸ�׺�Ӥ��Ϊͬһ�ŵ���NO�����
                            
                            strType = IIf(NVL(rsData!ҽ�����, 0) = 0, IIf(rsData!�����־ = 1 Or rsData!�����־ = 4, "������ʵ�", IIf(rsData!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(rsData!����) = True, "סԺ���ʵ�", IIf(rsData!���� Like "0*", "����", IIf(rsData!���� Like "1*", "����", "���ʱ�"))))
                            strType = strType & " " & Format(rsData!��������, "mm-dd hh:mm:ss")
                            
                            Set objNode = .Nodes.Add("P_" & Split(arr����(i), "|")(0) & str��ǰ����Key & rsData!����ID, 4, "N" & str��ǰ����Key & rsData!����ID & "_" & str��ǰNO & Split(arr����(i), "|")(0) & "_" & rsData!����, str��ǰNO & "(" & strType & ")", 4)
                            objNode.Tag = rsData!NO & "|" & rsData!����ID & "|" & rsData!����
                            objNode.Expanded = False
                        End If
                        
                        rsData.MoveNext
                    Loop
                End If
            Next
        End If
    End With
    
End Sub
Private Function GetDrugFormat() As Integer
    Dim strSave As String
    Dim arrColumn
    
    'ȡ��ҩƷ���Ƶĸ�ʽ��ʽ
    strSave = zlDatabase.GetPara("������", glngSys, 1342)
    If strSave = "" Then strSave = "0|ҩƷ����,0|������,0|Ӣ����,0|����,0|����ҽ��,0|״̬,0|����,0|NO,0|����Ա,0|����,0|����,0|סԺ��,0|���,0|����,0|����,0|��,0|����,0|������,0|׼����,0|��ҩ��,0|����,0|���,0|����,0|Ƶ��,0|�÷�,0|����ʱ��,0|˵��,0|����Ա,0|��ҩʱ��,0|��/��ҩ��,0|�ⷿ��λ"
    arrColumn = Split(strSave, ",")
    GetDrugFormat = Val(Split(arrColumn(0), "|")(0))
End Function

Private Sub ReturnSelected��ҩ;��(ByVal intType As Integer)
    'intType:0-˫����ҩ;���б�ʱ��1-��ҩ;���б��а��س�ʱ
    Dim n As Integer
    
    With Lvw��ҩ;��
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt��ҩ;��.Tag = ""
        Me.txt��ҩ;��.Text = ""
        
        '���ѡ����ȫѡ������ȡ���и�ҩ;����
        If .ListItems(1).Checked Then
            Me.txt��ҩ;��.Tag = ""
            Me.txt��ҩ;��.Text = "���и�ҩ;��"
            .Visible = False
            Exit Sub
        End If
        For n = 1 To .ListItems.count
            If .ListItems(n).Checked Then
                Me.txt��ҩ;��.Tag = IIf(Me.txt��ҩ;��.Tag = "", Mid(.ListItems(n).Key, 2), Me.txt��ҩ;��.Tag & "," & Mid(.ListItems(n).Key, 2))
                Me.txt��ҩ;��.Text = IIf(Me.txt��ҩ;��.Text = "", .ListItems(n).Text, Me.txt��ҩ;��.Text & "," & .ListItems(n).Text)
            End If
        Next
        
        If intType = 0 Then
            '�����ǰ˫���ĸ�ҩ;��δ��ѡ�ϣ�����ǰ˫���ĸ�ҩ;��Ҳ���뵽�༭����
            If .SelectedItem.Checked = False Then
                .SelectedItem.Checked = True
                Me.txt��ҩ;��.Tag = IIf(Me.txt��ҩ;��.Tag = "", Mid(.SelectedItem.Key, 2), Me.txt��ҩ;��.Tag & "," & Mid(.SelectedItem.Key, 2))
                Me.txt��ҩ;��.Text = IIf(Me.txt��ҩ;��.Text = "", .SelectedItem.Text, Me.txt��ҩ;��.Text & "," & .SelectedItem.Text)
            End If
            
            '���ѡ����ȫѡ������ȡ���и�ҩ;����
            If .ListItems(1).Checked Then
                Me.txt��ҩ;��.Tag = ""
                Me.txt��ҩ;��.Text = "���и�ҩ;��"
                .Visible = False
                Exit Sub
            End If
        End If
        
        .Visible = False
    End With
End Sub

Private Sub ReturnSelected����(ByVal intType As Integer)
    'intType:0-˫�������б�ʱ��1-�����б��а��س�ʱ
    Dim n As Integer
    
    With LvwҩƷ����
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txtҩƷ����.Text = ""
        
        '���ѡ����ȫѡ������ȡ���и�ҩ;����
        If .ListItems(1).Checked Then
             Me.txtҩƷ����.Text = "����ҩƷ����"
            .Visible = False
            Exit Sub
        End If
        
        For n = 1 To .ListItems.count
            If .ListItems(n).Checked Then
                Me.txtҩƷ����.Text = IIf(Me.txtҩƷ����.Text = "", Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1), Me.txtҩƷ����.Text & "," & Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1))
            End If
        Next
        
        If intType = 0 Then
            '�����ǰ˫���ĸ�ҩ;��δ��ѡ�ϣ�����ǰ˫���ĸ�ҩ;��Ҳ���뵽�༭����
            If .SelectedItem.Checked = False Then
                .SelectedItem.Checked = True
                Me.txtҩƷ����.Text = IIf(Me.txtҩƷ����.Text = "", Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1), Me.txtҩƷ����.Text & "," & Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1))
            End If
            
            If .ListItems(1).Checked Then
                 Me.txtҩƷ����.Text = "����ҩƷ����"
                .Visible = False
                Exit Sub
            End If
        End If
        
        .Visible = False
    End With
End Sub

Private Sub InitSendRec()
    Set mrsSendData = New ADODB.Recordset
    With mrsSendData
        If .State = 1 Then .Close
        
        '�ü�¼��Ӧ�ĵ�����Ϣ
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable              'ҩƷ�շ�ID
        .Fields.Append "���", adDouble, 18, adFldIsNullable                'ҩƷ�շ����
        .Fields.Append "��¼״̬", adDouble, 2, adFldIsNullable             'ҩƷ�շ���¼�ļ�¼״̬
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable           '���˵������˱�ҽ����������������
        .Fields.Append "����", adDouble, 2, adFldIsNullable                  '��ҩ�����У���ʾԺ����ҩ����Ժ��ҩ����ȡҩ������
        .Fields.Append "���շ�", adDouble, 2, adFldIsNullable               '�Ƿ����շѣ�1�����շ�
        .Fields.Append "����", adDouble, 18, adFldIsNullable                'ҩƷ�շ��������ͣ�8�������շѵ���9��סԺ���˵���10��סԺ���˱�
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable              'ҩƷ�շ�NO��
        .Fields.Append "����Ա", adLongVarChar, 20, adFldIsNullable         'סԺ���ü�¼�еĲ���Ա
        .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable           'ҩƷ�շ���¼��ժҪ
        .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable       'סԺ���ü�¼�еĵǼ�ʱ��
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable         'ҩƷ�շ���¼����ҩ��
        .Fields.Append "�����", adLongVarChar, 20, adFldIsNullable         'סԺ���ü�¼�еĲ���Ա
        .Fields.Append "��ҳID", adDouble, 18, adFldIsNullable              'סԺ���ü�¼�е���ҳID
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
    
        .Fields.Append "�����", adDouble, 18, adFldIsNullable            '���ԡ�����ҽ����¼���ġ�������������ں�����ҩ��PASS��
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable              '������ҽ����¼����ID��סԺ���ü�¼���ġ�ҽ����š�
        .Fields.Append "���id", adDouble, 18, adFldIsNullable              '���ԡ�����ҽ����¼���ġ����ID�������ڷ���
        .Fields.Append "������", adDouble, 1, adFldIsNullable
        .Fields.Append "��ҩĿ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 1000, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƥ�Խ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable              '���ԡ�סԺ���ü�¼���ġ����˿���ID��
        .Fields.Append "��ҩ���ű���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable       '��ʱ��ҩƷ�շ���¼���ġ��Է�����ID����Ӧ�Ĳ���
        .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable          '��ʱ��ҩƷ�շ���¼���ġ��Է�����ID��
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "ҽ������", adLongVarChar, 40, adFldIsNullable
        
        '������Ϣ
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable           '���˿���
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����ҽ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ɫ", adDouble, 18, adFldIsNullable
        
        'ҩƷ��Ϣ
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "Ʒ��", adLongVarChar, 50, adFldIsNullable           'ҩƷ���ƣ�0��ҩƷ���������ƣ�1��ҩƷ���룻2��ҩƷ����
        .Fields.Append "������", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "Ӣ����", adLongVarChar, 80, adFldIsNullable         '���ԡ�������Ŀ���������ɿ����Ż�
        .Fields.Append "�䷽����", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable                 '�������ԣ����ԡ�ҩƷ���
        .Fields.Append "��", adLongVarChar, 50, adFldIsNullable             '��ҩ���������ԡ�ҩƷ�շ���¼��
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ԭʼ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�����λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ִ��ʱ�䷽��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ƶ�ʴ���", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩʱ��", adLongVarChar, 1000, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable            '��С��λ���������жϿ����
        .Fields.Append "������λ", adLongVarChar, 20, adFldIsNullable       '���ԡ�������ĿĿ¼���ġ����㵥λ�������ں�����ҩ��PASS��
        .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ҩƷ���������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��װ", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ����", adDouble, 18, adFldIsNullable            '�շ���¼�е�ʵ�������������ȽϿ��
        .Fields.Append "��ΣҩƷ", adDouble, 2, adFldIsNullable
        .Fields.Append "�Ƿ�Ƥ��", adDouble, 2, adFldIsNullable
        .Fields.Append "ҩʦ��˱�־", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ�з���", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        
        'ҩƷ���顢��ֵ��Ϣ���ۺϹ���������ҩƷ���͡��ɿ����Ż���
        .Fields.Append "�������", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��ֵ����", adLongVarChar, 10, adFldIsNullable
        
        '���ԡ�ҩƷ�����޶�͡�ҩƷ��桱�����ݲ�����ҩƷ�������ɿ����Ż�
        .Fields.Append "�ⷿ��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        
        .Fields.Append "��������", adDouble, 18, adFldIsNullable            '��¼���ҶԸ�ҩƷ�ļƻ������������ɸ��ݲ����Ż�
        .Fields.Append "����Ǽ�ʱ��", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable         '����������һ��SQL�����Ը��������Ż�
        
        .Fields.Append "λ��", adDouble, 18, adFldIsNullable                '���ڶ�λ
        .Fields.Append "״̬", adLongVarChar, 10, adFldIsNullable           '״̬����ҩ���ܷ���������
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable             '״̬���ڲ���ʶ��0��ȱҩ��1����ҩ��2���ܷ���3��������
        
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        
        '���ڵ���ǩ��
        .Fields.Append "������id", adDouble, 18, adFldIsNullable
        .Fields.Append "���ϵ��", adDouble, 18, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrsChargeOff = New ADODB.Recordset
    With mrsChargeOff
        If .State = 1 Then .Close
        .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�շ����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "׼������", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "��װ", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҳID", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ִ�б�־", adDouble, 2, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҽ������", adLongVarChar, 40, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
        
    Set mrsChargeOffMain = New ADODB.Recordset
    With mrsChargeOffMain
        If .State = 1 Then .Close
        .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "׼������", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub LoadCustomSet()
    Dim str��ҩ���� As String
    Dim intSendType As Integer
    Dim n As Integer
   
    mParams.blnShowReject = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��ʾ�ܷ�", "0")) = 1)
    mParams.blnSort = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "������ʱ������", "0")) = 1)
    mParams.blnOnlyShowDept = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "�����б�", "0")) = 1)
    mParams.intShowDept = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��ʾ����", "0"))
    mParams.int�������� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��������", "1"))
    mParams.int����ģʽ = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "����ģʽ", 0))
    If mParams.int����ģʽ < 0 Then
        mParams.int����ģʽ = 0
    End If
    
    mParams.intAdviceType = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����", "ҽ������", "0"))
    
    If mParams.intAdviceType >= 0 And mParams.intAdviceType <= Cboҽ������.ListCount - 1 Then
        Cboҽ������.ListIndex = mParams.intAdviceType
    End If
    
    '������ҩ����
    str��ҩ���� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����", "��ҩ����", "")
    If mblnExistOtherSendType = True And str��ҩ���� <> "" Then
        For n = 0 To chkSendType.UBound
            If InStr(1, "," & str��ҩ���� & ",", "," & chkSendType(n).Caption & ",") > 0 Then
                chkSendType(n).Value = 1
            End If
        Next
        picShowSendType_Click
    End If
    
    '��ҩ����
    intSendType = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����", "ԭ��ҩ����", "0"))
    
    chkSend(0).Value = 0
    chkSend(1).Value = 0
    chkSend(2).Value = 0
    
    If intSendType < 0 Or intSendType > 6 Then
        intSendType = 0
    End If
    
    If intSendType = 0 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
        chkSend(2).Value = 1
    ElseIf intSendType = 1 Then
        chkSend(0).Value = 1
        chkSend(2).Value = 1
    ElseIf intSendType = 3 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
    ElseIf intSendType = 6 Then
        chkSend(1).Value = 1
        chkSend(2).Value = 1
    ElseIf intSendType = 5 Then
        chkSend(0).Value = 1
    ElseIf intSendType = 2 Then
        chkSend(1).Value = 1
    ElseIf intSendType = 4 Then
        chkSend(2).Value = 1
    End If
End Sub
Private Sub RefreshChargeOffDetail()
    '���³���������ϸ
    Dim strSubUnit As String
    Dim rstemp As ADODB.Recordset
    Dim strCon As String
    Dim strTmpCon As String
    Dim str����ʱ�� As String
    Dim lng��ҩ����ID As Long
    
    Dim lng����id As Long
    Dim dbl׼������ As Double
    
    Dim str����ID�� As String
    Dim lng���� As Long
    Dim strҩƷID�� As String
    Dim lngҩƷid As Long
    Dim strSqlҩ�� As String
    
    'Ҫ����ӦȨ�޺Ͳ���ʱ���ܽ������˴���
    If mPrives.bln��ҩ���� = False Or mParams.bln���ܷ�ҩ = False Then Exit Sub
    
    With mrsDeptList
        .Filter = ""
        .Sort = "����ID,ҩƷID"
        Do While Not .EOF
            If !ִ��״̬ = 1 Then
                If lng���� <> !����ID Then
                    lng���� = !����ID
                    str����ID�� = str����ID�� & IIf(str����ID�� = "", "", ",") & !����ID
                End If
                
                If lngҩƷid <> !ҩƷid Then
                    lngҩƷid = !ҩƷid
                    strҩƷID�� = strҩƷID�� & IIf(strҩƷID�� = "", "", ",") & !ҩƷid
                End If
            End If
            
            .MoveNext
        Loop
    End With
    If str����ID�� = "" Then Exit Sub
        
    '��λ����װ����
    Select Case mParams.strUnit
    Case "�ۼ۵�λ"
        strSubUnit = "X.���㵥λ ��λ,1 ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    Case "���ﵥλ"
        strSubUnit = "D.���ﵥλ ��λ,D.�����װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    Case "סԺ��λ"
        strSubUnit = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    Case "ҩ�ⵥλ"
        strSubUnit = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    End Select
    
    If mcondition.strNo <> "" Then
    ElseIf mcondition.strסԺ�� <> "" Then
        strCon = strCon & " And B.��ʶ��=[4] "
    ElseIf mcondition.str���� <> "" Then
        strCon = strCon & " And B.���� Like [5] "
    ElseIf mcondition.lng����ID <> -1 Then
        strCon = strCon & " And B.����ID=[6] "
    ElseIf mcondition.str���� <> "" Then
        strCon = strCon & " And B.���� = [7] "
    End If
    
    If mParams.intҩƷ������ʾ = 1 Then
        strSqlҩ�� = "'['||X.����||']'|| Nvl(K.����,X.����) As ҩƷ����,"
    Else
        strSqlҩ�� = "'['||X.����||']'|| X.���� As ҩƷ����,"
    End If
    
    gstrSQL = "Select Distinct " & strSqlҩ�� & "K.���� As ��Ʒ��," & _
        " C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����,C.����, C.Ч��, F.����, P.���� As ��������,E.���� As ��ҩ����,E.Id As ��ҩ����Id, " & _
        " A.����id, B.��� As �������, B.��¼����, B.��ҳID, B.����id, A.����ʱ��, " & strSubUnit & " " & _
        " From ���˷������� A, סԺ���ü�¼ B," & _
        " (Select A.ID, A.����, A.NO, A.���, A.ҩƷid, A.����, A.����,A.����, A.Ч��, A.����id, B.ʵ������ " & _
            " From ҩƷ�շ���¼ A, " & _
            " (Select C.����, C.NO, C.���, C.ҩƷid, Sum(Nvl(C.����, 1) * C.ʵ������) As ʵ������ " & _
            " From ҩƷ�շ���¼ C, ���˷������� A, סԺ���ü�¼ B " & _
            " Where A.�������=1 And A.����id = B.ID And B.NO = C.NO And B.ID = C.����id And A.״̬ = 0 " & _
            " And C.���� In (9, 10) And C.������� Is Not Null And C.�ⷿid = [1] And Instr([3], ',' || A.�շ�ϸĿid || ',') > 0 " & strTmpCon

    '�ų�������Һ�������Ĺ����в����ĵ���
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
    
    gstrSQL = gstrSQL & " Group By C.����, C.NO, C.���, C.ҩƷid " & _
            " Having Sum(Nvl(C.����, 1) * C.ʵ������) > 0) B" & _
            " Where A.NO = B.NO And A.���� = B.���� And A.ҩƷid + 0 = B.ҩƷid And A.��� = B.��� And A.����� Is Not Null " & _
            " And (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0))C, " & _
        " ҩƷ��� D, �շ���ĿĿ¼ X, �շ���Ŀ���� K, ���ű� P, ������ҳ F, ���ű� E " & _
        " Where A.�������=1 And A.����id = B.ID And B.NO = C.NO And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID And B.����id = F.����id And B.��ҳid = F.��ҳid  And A.���벿��id = E.ID " & strCon & _
        " And X.Id = K.�շ�ϸĿID(+) AND K.����(+)=3  And B.ִ�в���id = [1] And Instr([2], ',' || A.���벿��id || ',') > 0 And A.����� Is Null And A.״̬ = 0 "
    
    If mParams.bln��˳�Ժ�������� = False Then
        gstrSQL = gstrSQL & " And F.��Ժ���� Is Null "
    End If
        
    gstrSQL = gstrSQL & " Order By A.����ʱ��, C.����, C.NO, C.��� Desc "
    
    On Error GoTo errHandle
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", _
        mcondition.lngҩ��ID, _
        "," & str����ID�� & ",", _
        "," & strҩƷID�� & ",", _
        mcondition.strסԺ��, _
        mcondition.str����, _
        mcondition.lng����ID, _
        mcondition.str����)
    
    If rstemp.EOF Then
        Exit Sub
    End If
    
    Do While Not rstemp.EOF
        With mrsChargeOff
            .AddNew
            !ҩƷ���� = rstemp!ҩƷ����
            !��ҩ���� = rstemp!��ҩ����
            !��ҩ����ID = rstemp!��ҩ����ID
            !���� = rstemp!����
            !NO = rstemp!NO
            !ҩƷid = rstemp!ҩƷid
            !����ʱ�� = Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            !����ID = rstemp!����ID
            !�շ���� = rstemp!�շ����
            !���� = rstemp!����
            !���� = rstemp!����
            !Ч�� = rstemp!Ч��
            
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And NVL(!Ч��) <> "" Then
                '����Ϊ��Ч��
                !Ч�� = Format(DateAdd("D", -1, !Ч��), "yyyy-mm-dd")
            End If
            
            !׼������ = rstemp!׼������
            !�������� = rstemp!��������
            !��װ = rstemp!��װ
            !��λ = rstemp!��λ
            !�շ�Id = rstemp!�շ�Id
            !��ҳid = IIf(IsNull(rstemp!��ҳid), 0, rstemp!��ҳid)
            !������� = rstemp!�������
            !���� = rstemp!����
            !����ID = rstemp!����ID
            !��¼���� = rstemp!��¼����
            !���� = IIf(IsNull(rstemp!����), 0, rstemp!����)
            !��˱�־ = 0
            !ִ�б�־ = 0
            
            .Update
        End With
        
        With mrsChargeOffMain
'            dbl׼������ = dbl׼������ + rstemp!׼������
            If lng��ҩ����ID <> rstemp!��ҩ����ID Or str����ʱ�� <> Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss") Or lng����id <> rstemp!����ID Then
                .AddNew
                !��ҩ����ID = rstemp!��ҩ����ID
                !ҩƷid = rstemp!ҩƷid
                !����ʱ�� = Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                !����ID = rstemp!����ID
                !׼������ = rstemp!׼������
                !�������� = rstemp!��������
                
                .Update
                
                dbl׼������ = 0
            Else
                !׼������ = !׼������ + rstemp!׼������
                .Update
            End If
            lng��ҩ����ID = rstemp!��ҩ����ID
            str����ʱ�� = Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            lng����id = rstemp!����ID
        End With
        
        rstemp.MoveNext
    Loop
    
    'ֻ����ҩ�嵥��Ӧ��ҩƷ������ҩ����ID��ҩƷIDΪ׼��
    mrsChargeOff.MoveFirst
    Do While Not mrsChargeOff.EOF
        mrsSendData.Filter = "ִ��״̬=" & mState.��ҩ
        mrsSendData.Sort = "��ҩ����id,ҩƷid"
        Do While Not mrsSendData.EOF
            If mrsChargeOff!��ҩ����ID = mrsSendData!��ҩ����ID And mrsChargeOff!ҩƷid = mrsSendData!ҩƷid Then
                mrsChargeOff!��˱�־ = 1
                mrsChargeOff.Update
            End If
            mrsSendData.MoveNext
        Loop
        mrsChargeOff.MoveNext
    Loop
        
    Call AutoExpendQuantity
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AutoExpendQuantity()
    '���ǵ�ͬһ����ID��Ӧ����շ�ID���������Ҫ�����������ֽ⵽����շ���¼��
    '�ֽ��ԭ���ǰ���Ŵ�����ȷ��䣨�Ѱ���Ž�������
    Dim n As Integer
    Dim dbl׼������ As Double
    Dim dblʣ������ As Double
    Dim int�շ���� As Integer
    Dim lng����id As Long
    Dim lngҩƷid As Long
    Dim str����ʱ�� As String
    
    With mrsChargeOff
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl׼������ = !׼������

            If lng����id = !����ID And lngҩƷid = !ҩƷid And str����ʱ�� = !����ʱ�� Then

            Else
                dblʣ������ = !��������
            End If

            If dblʣ������ >= dbl׼������ Then
                dblʣ������ = dblʣ������ - dbl׼������
                !�������� = dbl׼������
            Else
                !�������� = dblʣ������
                dblʣ������ = 0
            End If

            lng����id = !����ID
            lngҩƷid = !ҩƷid
            str����ʱ�� = !����ʱ��

            .Update
            .MoveNext
        Next
    End With
    
    '��������������׼�����������־Ϊ�ܾ����
    With mrsChargeOffMain
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            mrsChargeOff.Filter = "ҩƷID=" & !ҩƷid & _
                " And ����ID=" & !����ID & _
                " And ����ʱ��='" & !����ʱ�� & "'"
            If mrsChargeOff.RecordCount > 0 Then
                If !׼������ < !�������� Then
                    Do While Not mrsChargeOff.EOF
                        mrsChargeOff!��˱�־ = 2
                        mrsChargeOff.Update
                        mrsChargeOff.MoveNext
                    Loop
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub InitReturnRec()
    '�ѷ�������¼��
    Set mrsReturnData = New ADODB.Recordset
    With mrsReturnData
        If .State = 1 Then .Close
        
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable                  '��ҩ�����У���ʾԺ����ҩ����Ժ��ҩ����ȡҩ������
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "Ʒ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "Ӣ����", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�������", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable
        .Fields.Append "��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "׼����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ҩ��", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ʵ������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ⷿ��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��װ", adDouble, 18, adFldIsNullable
        .Fields.Append "��ΣҩƷ", adDouble, 2, adFldIsNullable
        
        .Fields.Append "����Ա", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩʱ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
       
        .Fields.Append "�����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "���id", adDouble, 18, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "ҩƷ���������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "��ҩ��", adDouble, 18, adFldIsNullable
        
        .Fields.Append "ҽ������", adLongVarChar, 40, adFldIsNullable

        .Fields.Append "״̬", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
        
        .Fields.Append "ת��", adDouble, 1, adFldIsNullable
        .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 40, adFldIsNullable
        
        .Fields.Append "��������", adDouble, 1, adFldIsNullable         '�����ڲ���[����������ʱ���ܽ�����ҩ����],����������ʱ�����ֶβŽ��ܶ�Ӧ��ֵ
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub InitApplyforcredit()
    '������������ļ�¼��
    Set mrsApplyforcredit = New ADODB.Recordset
    With mrsApplyforcredit
        If .State = 1 Then .Close
        
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable              'ҩƷ�շ�ID
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable             '״̬���ڲ���ʶ��0��ȱҩ��1����ҩ��2���ܷ���3��������
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "���˿���", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub InitMsgRec()
    '��Ϣ���ռ�¼��
    Set mrsReceiveMsg = New ADODB.Recordset
    With mrsReceiveMsg
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 40, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
Private Function LoadSendRecord(ByVal rsData As ADODB.Recordset) As Boolean
    'װ�ط�ҩ���ݼ�
    Dim intState As Integer
    Dim strState As String
    
    On Error GoTo errHandle
    
    With rsData
        Do While Not .EOF
            mrsSendData.AddNew
            
            mrsSendData!�շ�Id = !�շ�Id
            mrsSendData!��� = !���
            mrsSendData!��¼״̬ = !��¼״̬
            mrsSendData!���� = IIf(NVL(!ҽ�����, 0) = 0, IIf(!�����־ = 1 Or !�����־ = 4, "������ʵ�", IIf(!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(!����) = True, "סԺ���ʵ�", IIf(!���� Like "0*", "����", IIf(!���� Like "1*", "����", "���ʱ�"))))
            mrsSendData!���� = IIf(IsNull(!����), 0, !����)
            mrsSendData!���շ� = !���շ�
            mrsSendData!���� = !����
            mrsSendData!NO = !NO
            mrsSendData!����Ա = IIf(IsNull(!����Ա����), "", !����Ա����)
            mrsSendData!˵�� = IIf(IsNull(!˵��), "", !˵��)
            mrsSendData!����ʱ�� = IIf(IsNull(!�Ǽ�ʱ��), "", Format(!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"))
            mrsSendData!��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            mrsSendData!����� = IIf(IsNull(!�����), "", !�����)
            mrsSendData!��ҳid = !��ҳid
            
            mrsSendData!������� = !�������
            mrsSendData!����� = !�����
            mrsSendData!ҽ��id = !ҽ��id
            mrsSendData!���ID = IIf(IsNull(!���ID), 0, !���ID)
            mrsSendData!������ = !������
            mrsSendData!��ҩĿ�� = NVL(!��ҩĿ��)
            mrsSendData!��ҩ���� = !��ҩ����
            
            mrsSendData!�������� = IIf(IsNull(!��������), "", Format(!��������, "yyyy-MM-dd HH:mm:ss"))
            
            '����ʵ����ҩ���������
            If IsNull(!��ʼִ��ʱ��) Or IsNull(!�״�ʱ��) Or IsNull(!ĩ��ʱ��) Or IsNull(!Ƶ�ʼ��) Or IsNull(!�����λ) Or IsNull(!ִ��ʱ�䷽��) Or (Not mParams.bln���ط�ҩʱ�����ҩ����) Then
                mrsSendData!��ҩʱ�� = ""
                mrsSendData!��ҩ���� = ""
            Else
                mrsSendData!��ҩʱ�� = Calc���ڷֽ�ʱ��(CDate(!�״�ʱ��), CDate(!ĩ��ʱ��), "", !ִ��ʱ�䷽��, !Ƶ�ʴ���, !Ƶ�ʼ��, !�����λ, CDate(!��ʼִ��ʱ��))
                mrsSendData!��ҩ���� = UBound(Split(mrsSendData!��ҩʱ��, ",")) + 1
                
                '�Է�ҩʱ�����¸�ʽ��
                mrsSendData!��ҩʱ�� = Format��ҩʱ��(mrsSendData!��ҩʱ��, !�����λ, !ҽ��id, mrsSendData!��������, !ִ��ʱ�䷽��)
            End If
            
            mrsSendData!Ƥ�Խ�� = !Ƥ�Խ��
            mrsSendData!����ʱ�� = !����ʱ��
            
            mrsSendData!����ID = !����ID
            mrsSendData!��ҩ���ű��� = !��ҩ���ű���
            mrsSendData!��ҩ���� = !��ҩ����
            mrsSendData!��ҩ����ID = !��ҩ����ID
            mrsSendData!��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            
            mrsSendData!����ID = !����ID
            mrsSendData!���� = !����
            mrsSendData!�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
            mrsSendData!סԺ�� = NVL(!��ʶ��)
            mrsSendData!���� = !����
            mrsSendData!���� = !����
            mrsSendData!����ҽ�� = !����ҽ��
            mrsSendData!���� = !����
            mrsSendData!�������� = !��������
            mrsSendData!��ɫ = IIf(IsNull(!��ɫ), vbBlack, !��ɫ)
            
            mrsSendData!ҩƷid = !ҩƷid
            mrsSendData!ҩ��ID = !ҩ��ID
            mrsSendData!Ʒ�� = !Ʒ��
            mrsSendData!������ = IIf(IsNull(!������), "", !������)
            mrsSendData!Ӣ���� = IIf(IsNull(!Ӣ����), "", !Ӣ����)
            mrsSendData!�䷽���� = IIf(IsNull(!�䷽����), "", !�䷽����)
            mrsSendData!��� = !���
            mrsSendData!���� = IIf(IsNull(!����), "", !����)
            mrsSendData!�Ƿ�Ƥ�� = !�Ƿ�Ƥ��
            
            mrsSendData!���� = IIf(IsNull(!����), 0, !����)
            mrsSendData!���� = IIf(IsNull(!����), "", !����)
            mrsSendData!Ч�� = IIf(IsNull(!Ч��), "", !Ч��)
            mrsSendData!���� = IIf(IsNull(!����), 0, !����)
            mrsSendData!�� = IIf(IsNull(!��), 1, !��)
            mrsSendData!���� = FormatEx(IIf(IsNull(!����), 1, !����) / !��װ, 5) & !��λ
            mrsSendData!���� = FormatEx(!���� * !��װ, 5)
            
            mrsSendData!��λ = !��λ
            mrsSendData!��װ = !��װ
            
'            mrsSendData!��� = Format(!���, "#####0.00;-#####0.00; ;")
            mrsSendData!��� = !���
            mrsSendData!���� = IIf(IsNull(!����), "", FormatEx(!����, 5) & NVL(!���㵥λ) & "(" & FormatEx(!���� / !����ϵ�� / !��װ, 5) & !��λ & ")")
            mrsSendData!ԭʼ���� = IIf(IsNull(!����), "", FormatEx(!����, 5) & NVL(!���㵥λ))
            mrsSendData!Ƶ�� = IIf(IsNull(!Ƶ��), "", !Ƶ��)
            mrsSendData!�����λ = IIf(IsNull(!�����λ), "", !�����λ)
            mrsSendData!ִ��ʱ�䷽�� = IIf(IsNull(!ִ��ʱ�䷽��), "", !ִ��ʱ�䷽��)
            mrsSendData!Ƶ�ʴ��� = IIf(IsNull(!ִ��ʱ�䷽��), 0, !Ƶ�ʴ���)
            mrsSendData!�������� = IIf(IsNull(!ִ��ʱ�䷽��), 0, !��������)
            mrsSendData!ʵ������ = FormatEx(Val(IIf(IsNull(!����), 1, !����)) * (Val(IIf(IsNull(!��), 1, !��))) / !��װ, 5)
            mrsSendData!������λ = NVL(!���㵥λ)
            mrsSendData!�÷� = IIf(IsNull(!�÷�), "", !�÷�)
            
            mrsSendData!��ҩ���� = IIf(IsNull(!����), 1, !����)
            
            mrsSendData!��ΣҩƷ = IIf(IsNull(!��ΣҩƷ), 0, !��ΣҩƷ)
            
            mrsSendData!������� = IIf(IsNull(!�������), "", !�������)
            mrsSendData!��ֵ���� = IIf(IsNull(!��ֵ����), "", !��ֵ����)
            
            mrsSendData!�ⷿ��λ = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
            mrsSendData!������� = !�������
            
            mrsSendData!�������� = FormatEx(IIf(IsNull(!��������), 0, !��������) / !��װ, 5)
            mrsSendData!����Ǽ�ʱ�� = IIf(IsNull(!����Ǽ�ʱ��), "", Format(!����Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"))
            
            mrsSendData!��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            
            mrsSendData!ҽ������ = IIf(IsNull(!ҽ������), "", !ҽ������)

            mrsSendData!ҩƷ��������� = !Ʒ��
            mrsSendData!ҩƷ���� = !ҩƷ����
            mrsSendData!ҩƷ���� = !ҩƷ����
            
            mrsSendData!������� = !�������
            
            mrsSendData!������id = !������id
            mrsSendData!���ϵ�� = !���ϵ��
            mrsSendData!������ = IIf(IsNull(!������), "", !������)
            mrsSendData!��ҩ���� = IIf(IsNull(!��ҩ����), "", Format(!��ҩ����, "yyyy-MM-dd HH:mm:ss"))
            mrsSendData!����ID = !����ID
            mrsSendData!ҩʦ��˱�־ = NVL(!ҩʦ��˱�־, 0)
            mrsSendData!ִ�з��� = NVL(!ִ�з���, 0)
            mrsSendData!��� = NVL(!���, 0)
            
            mrsSendData!λ�� = .AbsolutePosition
            
            '����Ƿ�����ҩ
            intState = mState.��ҩ
            If !���շ� = 0 Then intState = mState.������
            If Not IsNull(!˵��) Then
                intState = IIf(!˵�� = "�ܷ�", mState.�ܷ�_������, intState)
            End If
            If mParams.bln����δ��˴�����ҩ = False Then
                If IsNull(!�����) Then
                    intState = mState.������
                Else
                    If Trim(!�����) = "" Then intState = mState.������
                End If
            ElseIf intState = mState.������ Then
                intState = mState.��ҩ
            End If
            
            '��鶾����࣬��ֵ���࣬��Σ����
            If intState <> mState.������ Then
                If mParams.str������� <> "" And !������� <> "" Then
                    If InStr("," & mParams.str������� & ",", "," & !������� & ",") > 0 Then
                        intState = mState.������
                    End If
                End If
                If mParams.str��ֵ���� <> "" And !��ֵ���� <> "" Then
                    If InStr("," & mParams.str��ֵ���� & ",", "," & !��ֵ���� & ",") > 0 Then
                        intState = mState.������
                    End If
                End If
                If mParams.str��Σ���� <> "" And !��ΣҩƷ <> "" Then
                    If InStr("," & mParams.str��Σ���� & ",", "," & !��ΣҩƷ & ",") > 0 Then
                        intState = mState.������
                    End If
                End If
            End If
            
'            If !��¼״̬ > 1 Then
'                intState = mState.������
'            End If
            
            mrsSendData!ִ��״̬ = intState
            
            Select Case intState
                Case mState.ȱҩ
                    strState = "ȱҩ"
                Case mState.��ҩ
                    strState = "��ҩ"
                Case mState.�ܷ�
                    strState = "�ܷ�"
                Case mState.������, mState.�ܷ�_������
                    strState = "������"
            End Select
            mrsSendData!״̬ = strState
            
            mrsSendData.Update
            
            .MoveNext
        Loop
        
        'ȱҩ���
        If mParams.blnȱҩ��� = True Then
            Call CheckShortage(mrsSendData, False)
        End If
    End With
    
    LoadSendRecord = True
    Exit Function
errHandle:
    MsgBox "�����ڲ���¼��ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
    Call InitSendRec
    Exit Function
End Function

Private Function LoadReturnRecord(ByVal rsData As ADODB.Recordset) As Boolean
    Dim dblSumSended As Double '�ѷ�����
    
    On Error GoTo errHandle
    
    With rsData
        Do While Not .EOF
            mrsReturnData.AddNew
            mrsReturnData!�շ�Id = !�շ�Id
            mrsReturnData!ҩƷid = !ҩƷid
            mrsReturnData!���� = !����
            mrsReturnData!��ҩ����ID = !��ҩ����ID
            mrsReturnData!���� = IIf(NVL(!ҽ�����, 0) = 0, IIf(!�����־ = 1 Or !�����־ = 4, "������ʵ�", IIf(!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(!����) = True, "סԺ���ʵ�", IIf(!���� Like "0*", "����", IIf(!���� Like "1*", "����", "���ʱ�"))))
            mrsReturnData!���� = IIf(IsNull(!����), 0, !����)
            mrsReturnData!NO = !NO
            mrsReturnData!���� = !����
            mrsReturnData!��� = !���
            mrsReturnData!������� = !�������
            mrsReturnData!����ID = !����ID
            mrsReturnData!���� = !����
            mrsReturnData!���� = IIf(IsNull(!����), "", !����)
            mrsReturnData!�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
            mrsReturnData!סԺ�� = NVL(!��ʶ��)
            mrsReturnData!Ʒ�� = !Ʒ��
            mrsReturnData!������ = !������
            mrsReturnData!Ӣ���� = !Ӣ����
            mrsReturnData!��� = IIf(IsNull(!���), "", !���)
            mrsReturnData!���� = IIf(IsNull(!����), "", !����)
            mrsReturnData!������� = NVL(!�������)
            mrsReturnData!���� = IIf(IsNull(!����), 0, !����)
            mrsReturnData!���� = IIf(IsNull(!����), 0, !����)
            mrsReturnData!���� = IIf(IsNull(!����), "", !����)
            mrsReturnData!Ч�� = IIf(IsNull(!Ч��), "", !Ч��)
            mrsReturnData!�� = IIf(IsNull(!��), 1, !��)
            mrsReturnData!���� = FormatEx(IIf(IsNull(!����), 1, !����) / !��װ, 5) & !��λ
'            If !�ɲ��� <> 1 Then
'                mrsReturnData!������ = FormatEx(IIf(IsNull(!��������), 1, !��������) / !��װ, 5)
'                mrsReturnData!׼���� = FormatEx(IIf(IsNull(!׼����), 1, !׼����) / !��װ, 5)
'                mrsReturnData!��ҩ�� = FormatEx(IIf(IsNull(!׼����), 1, !׼����) / !��װ, 5)
'            Else
                dblSumSended = GetSumSended(!����, !NO, !ҩƷid, !���)
                mrsReturnData!������ = FormatEx((Val(IIf(IsNull(!����), 1, !����)) * (Val(IIf(IsNull(!��), 1, !��))) - dblSumSended) / !��װ, 5)
                mrsReturnData!׼���� = FormatEx(dblSumSended / !��װ, 5)
                mrsReturnData!��ҩ�� = FormatEx(dblSumSended / !��װ, 5)
'            End If
            mrsReturnData!��װ = !��װ
            mrsReturnData!��λ = !��λ
            mrsReturnData!���� = FormatEx(!���� * !��װ, 5)
            mrsReturnData!��� = !���
            mrsReturnData!���� = IIf(IsNull(!����), "", FormatEx(!����, 5) & NVL(!���㵥λ) & "(" & FormatEx(!���� / !����ϵ�� / !��װ, 5) & !��λ & ")")
            mrsReturnData!������λ = NVL(!���㵥λ)
            mrsReturnData!Ƶ�� = IIf(IsNull(!Ƶ��), "", !Ƶ��)
            mrsReturnData!�÷� = IIf(IsNull(!�÷�), "", !�÷�)
            mrsReturnData!˵�� = IIf(IsNull(!˵��), "", !˵��)
            mrsReturnData!����Ա = IIf(IsNull(!�����), "", !�����)
            mrsReturnData!��ҩʱ�� = IIf(IsNull(!��ҩʱ��), "", !��ҩʱ��)
            
            mrsReturnData!ҽ������ = IIf(IsNull(!ҽ������), "", !ҽ������)

            mrsReturnData!��ΣҩƷ = IIf(IsNull(!��ΣҩƷ), 0, !��ΣҩƷ)
            
            mrsReturnData!����� = !�����
            mrsReturnData!ҽ��id = !ҽ��id
            mrsReturnData!��ҩ�� = !��ҩ��
            mrsReturnData!ʵ������ = dblSumSended
            mrsReturnData!�ⷿ��λ = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
            mrsReturnData!ת�� = Val(!ת��)
            
            mrsReturnData!ҩƷ��������� = !Ʒ��
            mrsReturnData!ҩƷ���� = !ҩƷ����
            mrsReturnData!ҩƷ���� = !ҩƷ����
            mrsReturnData!����ʱ�� = !����ʱ��
            
            mrsReturnData!��ҩ�� = !��ҩ��
            
            mrsReturnData!���ID = IIf(IsNull(!���ID), 0, !���ID)
            
            If Val(!ת��) = -1 Then
                mrsReturnData!ִ��״̬ = mState.ת����¼
            ElseIf Val(!�ɲ���) = 1 Then
                mrsReturnData!ִ��״̬ = mState.��ҩ_ԭʼ��¼
            ElseIf Val(!�ɲ���) = 2 Then
                mrsReturnData!ִ��״̬ = mState.��ҩ_��ҩ��¼
            ElseIf Val(!�ɲ���) = 3 Then
                mrsReturnData!ִ��״̬ = mState.��ҩ_��ҩ��¼
            End If
            
            mrsReturnData!״̬ = "������"
            
            If mParams.bln����������ʱ���ܽ�����ҩ���� Then
                mrsReturnData!�������� = !��������
            End If
            
            mrsReturnData.Update
            
            .MoveNext
        Loop
    End With
    
    LoadReturnRecord = True
    Exit Function
errHandle:
    MsgBox "�����ڲ���¼��ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
    Call InitReturnRec
    Exit Function
End Function

Private Function Format��ҩʱ��(ByVal str��ҩʱ�� As String, ByVal str�����λ As String, ByVal lngҽ��id As Long, ByVal str�������� As String, ByVal strִ��ʱ�䷽�� As String) As String
    '�����ܡ�:�Է�ҩʱ�䰴��λ���и�ʽ��
    Dim str����ʱ�� As String
    Dim str�ֽ�ʱ�� As String
    Dim str����ʱ�� As String
    Dim str����ִ��ʱ�� As String
    Dim strƽ��ִ��ʱ�� As String
    Dim strTemp As String   '����ð������ʱ�Ļ���
    Dim arrTemp As Variant
    Dim i As Integer
    Dim j As Integer
    
    If str�����λ = "��" Then
        If Is�״�ִ��(lngҽ��id, str��������) And InStr(strִ��ʱ�䷽��, ",") > 0 Then
            '����ִ��
            '------------------------------------
            str����ʱ�� = Mid(Split(str��ҩʱ��, ",")(0), 1, 10)    '����-��-�ա�
            
            For i = 0 To UBound(Split(str��ҩʱ��, ","))
                str����ʱ�� = Split(str��ҩʱ��, ",")(i)
                str�ֽ�ʱ�� = Mid(str����ʱ��, 12, 5)   '��ʱ:�֡�
                
                If InStr(str����ʱ��, str����ʱ��) > 0 Then
                    '����ִ��ʱ���
                    If Not InStr(str����ִ��ʱ��, str�ֽ�ʱ��) > 0 Then
                        str����ִ��ʱ�� = str����ִ��ʱ�� & IIf(str����ִ��ʱ�� = "", "", ",") & str�ֽ�ʱ��
                    End If
                Else
                    '����ִ��ʱ���
                    If Not InStr(strƽ��ִ��ʱ��, str�ֽ�ʱ��) > 0 Then
                        strƽ��ִ��ʱ�� = strƽ��ִ��ʱ�� & IIf(strƽ��ִ��ʱ�� = "", "", ",") & str�ֽ�ʱ��
                    End If
                End If
            Next
        Else
            'ƽ��ִ��
            '------------------------------------
            For i = 0 To UBound(Split(str��ҩʱ��, ","))
                str����ʱ�� = Split(str��ҩʱ��, ",")(i)
                str�ֽ�ʱ�� = Mid(str����ʱ��, 12, 5)   '��ʱ:�֡�
                
                If Not InStr(strƽ��ִ��ʱ��, str�ֽ�ʱ��) > 0 Then
                    strƽ��ִ��ʱ�� = strƽ��ִ��ʱ�� & IIf(strƽ��ִ��ʱ�� = "", "", ",") & str�ֽ�ʱ��
                End If
            Next
        End If
        
        '��ʱ����絽������(ð���㷨)
        '------------------------------------
        arrTemp = Split(strƽ��ִ��ʱ��, ",")
        
        For i = 0 To UBound(arrTemp) - 1
            For j = 0 To UBound(arrTemp) - i - 1
                If Mid(arrTemp(j), 1, 2) > Mid(arrTemp(j + 1), 1, 2) Then
                    strTemp = arrTemp(j)
                    arrTemp(j) = arrTemp(j + 1)
                    arrTemp(j + 1) = strTemp
                End If
            Next
        Next
        
        strƽ��ִ��ʱ�� = ""
        
        For i = 0 To UBound(arrTemp)
            strƽ��ִ��ʱ�� = strƽ��ִ��ʱ�� & IIf(strƽ��ִ��ʱ�� = "", "", ",") & arrTemp(i)
        Next
        '------------------------------------
        
        Format��ҩʱ�� = IIf(str����ִ��ʱ�� = "", "", "����ִ�С�" & str����ִ��ʱ�� & "��") & IIf(str����ִ��ʱ�� <> "" And strƽ��ִ��ʱ�� <> "", "��", "") & IIf(strƽ��ִ��ʱ�� = "", "", "ִ��ʱ�䡾" & strƽ��ִ��ʱ�� & "��")
    Else
        For i = 0 To UBound(Split(str��ҩʱ��, ","))
            str����ʱ�� = Split(str��ҩʱ��, ",")(i)
            str�ֽ�ʱ�� = Mid(str����ʱ��, 6, 11)   '����-�� ʱ:�֡�
            
            Format��ҩʱ�� = Format��ҩʱ�� & IIf(Format��ҩʱ�� = "", "", ",") & str�ֽ�ʱ��
        Next
    End If
End Function

Private Function Is�״�ִ��(ByVal lngҽ��id As Long, ByVal str�������� As String) As Boolean
    '���ܣ��жϵ�ǰ�ĵ����Ƿ�Ϊҽ���״�ִ��
    
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 1" & vbNewLine & _
            "From (Select ����ʱ�� From (Select ����ʱ�� From ����ҽ������ Where ҽ��id = [1] Order By ����ʱ��) Where Rownum = 1)" & vbNewLine & _
            "Where ����ʱ�� = [2]"
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�Ϊҽ���״�ִ��", _
                lngҽ��id, _
                CDate(str��������))
    
    If rstemp.EOF Then
        Is�״�ִ�� = False
    Else
        Is�״�ִ�� = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSumSended(ByVal Int���� As Integer, ByVal strNo As String, ByVal lngҩƷid As Long, ByVal int��� As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "Select Sum(Nvl(����, 1) * ʵ������) �ѷ����� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ҩƷID+0 = [3] And ��� = [4] And ������� Is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ѷ�����", Int����, strNo, lngҩƷid, int���)
    
    If Not rsTmp.EOF Then
        GetSumSended = rsTmp!�ѷ�����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get��ҩʱ��(ByVal str��ʼִ��ʱ�� As String, ByVal str�״�ʱ�� As String, ByVal strĩ��ʱ�� As String, _
                             ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, ByVal strִ��ʱ�䷽�� As String, _
                             ByVal lngҽ��id As Long, ByVal intҽ��Ч�� As Integer) As String
    '�����ܡ�:��ȡ�����ҩʱ��
    Dim rstemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select zl_adviceexetimes([1],[2],[3],[4],[5],null,[6],[7],[8]) From Dual "

    Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "ȡʵ�ʷ�ҩ����", _
            lngҽ��id, _
            CDate(str�״�ʱ��), _
            CDate(strĩ��ʱ��), _
            strִ��ʱ�䷽��, _
            CDate(str��ʼִ��ʱ��), _
            intƵ�ʼ��, _
            str�����λ, _
            intҽ��Ч��)
    
    If Not rstemp.EOF Then
        Get��ҩʱ�� = IIf(IsNull(rstemp.Fields(0).Value), "", rstemp.Fields(0).Value)
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Loadȡ�Զ��巢ҩ����()
    '��ȡ��ҩ���ͣ�����̬���ӷ�ҩ����ѡ���
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    Set rsData = DeptSendWork_Get�Զ��巢ҩ����
    
    With rsData
        mblnExistOtherSendType = Not .EOF
        picShowSendType.Visible = mblnExistOtherSendType
        picSendType.Visible = mblnExistOtherSendType
        
        If mblnExistOtherSendType = False Then
            Exit Sub
        Else
            chkSendType(0).Caption = rsData!����
            chkSendType(0).Width = 150 + LenB(chkSendType(0).Caption) * 128
            If rsData.RecordCount > 1 Then
                rsData.MoveNext
                For i = 2 To rsData.RecordCount
                    Load chkSendType(i - 1)
                    chkSendType(i - 1).Visible = True
                    chkSendType(i - 1).Caption = rsData!����
                    chkSendType(i - 1).Width = 150 + LenB(chkSendType(i - 1).Caption) * 128
                    rsData.MoveNext
                Next
            End If
        End If
    End With
End Sub

Private Sub Loadʱ�䷶Χ()
    Dim dteTime As Date
    
    With cboʱ�䷶Χ
        .Clear
        .AddItem "0-����"
        .AddItem "1-������"
        .AddItem "2-������"
        .AddItem "3-ָ��ʱ�䷶Χ"
        
        .ListIndex = IIf(mParams.intDays < 4, mParams.intDays, 3)
        
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call picConMain_Resize
                Call picCondition_Resize
            End If
            .Tag = .ListIndex
        End If
    End With
    
    dteTime = zlDatabase.Currentdate
    Dtp��ʼʱ��.Value = Format(DateAdd("d", -1 * mParams.intDays, dteTime), "yyyy-MM-dd 00:00:00")
    Dtp����ʱ��.Value = Format(dteTime, "yyyy-MM-dd") & " 23:59:59"
    mdateBegin = dteTime
End Sub
Private Sub RefreshData()
    Dim intType As Integer
    
    'ˢ�����ݣ�ˢ�¿����б���Ĭ��ȫ����ѡ����ˢ����ϸ�嵥
    DoEvents
    cmdRefreshDept_Click
    
    DoEvents
    intType = IIf(tbcDetail.Selected.Index = 0, 0, 1)
    chkAll(intType).Value = 1
    chkAll_Click intType
    
    DoEvents
    cmdRefresh_Click
End Sub

Private Sub SaveCustomSet()
    Dim str��ҩ���� As String
    Dim intSendType As Integer
    Dim n As Integer
    
    '��ҩ����
    '0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If chkSend(0).Value = 1 And chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        intSendType = 0
    ElseIf chkSend(0).Value = 1 And chkSend(2).Value = 1 Then
        intSendType = 1
    ElseIf chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
        intSendType = 3
    ElseIf chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        intSendType = 6
    ElseIf chkSend(0).Value = 1 Then
        intSendType = 5
    ElseIf chkSend(1).Value = 1 Then
        intSendType = 2
    ElseIf chkSend(2).Value = 1 Then
        intSendType = 4
    End If

    '������ҩ����
    If mblnExistOtherSendType = True Then
        For n = 0 To chkSendType.UBound
            If chkSendType(n).Value = 1 Then
                str��ҩ���� = IIf(str��ҩ���� = "", "", str��ҩ���� & ",") & chkSendType(n).Caption
            End If
        Next
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��ʾ�ܷ�", IIf(mParams.blnShowReject, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "������ʱ������", IIf(mParams.blnSort, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "�����б�", IIf(mParams.blnOnlyShowDept, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��ʾ����", mParams.intShowDept
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��������", mParams.int��������
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "����ģʽ", mParams.int����ģʽ
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����", "ҽ������", mParams.intAdviceType
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����", "��ҩ����", str��ҩ����
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����", "ԭ��ҩ����", intSendType
End Sub
Private Sub SetColorState()
    '��ҩ�б����ɫ״̬
    picColorStateSend(mSendListColor.SendState).BackColor = mListColor.State_Send
    picColorStateSend(mSendListColor.RejectState).BackColor = mListColor.State_Reject
    picColorStateSend(mSendListColor.UnProcessState).BackColor = mListColor.State_UnProcess
    picColorStateSend(mSendListColor.ShortageState).BackColor = mListColor.State_Shortage
    
    '��ҩ�б����ɫ״̬
    picColorStateReturn(mReturnListColor.ReturnState).BackColor = mListColor.Return_Returned
    picColorStateReturn(mReturnListColor.OriginalState).BackColor = mListColor.Return_Original
    picColorStateReturn(mReturnListColor.SendedState).BackColor = mListColor.Return_Sended
End Sub

Private Sub SetCommandBar(ByVal intType As Integer)
    '1������ϵͳ������Ȩ�޵ȸı�˵�״̬
    '2�����ݵ�ǰҳ��͵�ǰѡ�����ϸ��¼���ı�˵�״̬
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl

    Select Case intType
        Case mListType.��ҩ
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = True
            End If
                
            '��֤ǩ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign���ŷ�ҩ = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '�ܷ��ָ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            'ȫѡ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            'ȫ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
        Case mListType.����
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = True
            End If
            
            '��֤ǩ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign���ŷ�ҩ = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '�ܷ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '�ܷ��ָ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            'ȫѡ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            'ȫ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
        Case mListType.�ܷ�
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
           '��֤ǩ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign���ŷ�ҩ = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '�ܷ��ָ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            'ȫѡ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            'ȫ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
        Case mListType.ȱҩ
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '��֤ǩ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign���ŷ�ҩ = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '�ܷ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '�ܷ��ָ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            'ȫѡ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            'ȫ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
        Case mListType.��ҩ
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '��֤ǩ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign���ŷ�ҩ = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '�ܷ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '�ܷ��ָ�
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '��ҩ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = True
            End If
            
            'ȫѡ
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            
            'ȫ��
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
    End Select
End Sub

Private Sub RefreshSendDept()
    'ˢ�´���ҩ�����б�
    Dim rsData As ADODB.Recordset
    Dim strTmpSql As String
    Dim strDanger As String
    Dim strToxicology As String
   
    '''select
    gstrSQL = "Select " & IIf(mParams.strSourceDep = "", "", "/*+rule*/") & "  distinct H.ID, H.���� As ��������, Nvl(A.��ҩ��, 0) As ��ҩ��, Decode(Nvl(c.Ӥ����,0), 0, Nvl(b.����, c.����), z.Ӥ������) ����, C.����ID, Decode(Nvl(c.Ӥ����,0), 0, Nvl(p.�Ա�, c.�Ա�), z.Ӥ���Ա�) �Ա�, Decode(Nvl(c.Ӥ����,0), 0, p.����, Ceil(Sysdate - z.����ʱ��) || '��') ����, S.����, S.NO, S.ҩƷid, " & _
        " Decode(Nvl(C.ҽ�����, 0), 0, 0, 1) ҽ�����, C.�����־, Nvl(S.����, 0) ����, S.ID As �շ�id, S.�������� ��������, 0 As �ܷ�, Nvl(B.��ǰ����,'') As ����,W.��ɫ,c.Ӥ����" & IIf(mParams.bln������, ", nvl(q.�����,0) �����,nvl(q.id,0) ���id", "")

    '''from
    gstrSQL = gstrSQL & " From סԺ���ü�¼ C, ҩƷ�շ���¼ S, ������Ϣ B, ҩƷ��� D, ҩƷ���� T, δ��ҩƷ��¼ A,������ҳ P,���ű� H,�������� W, ������������¼ Z " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([17]) As zlTools.t_NumList)) E ")
    gstrSQL = gstrSQL & IIf(mblnIs�������� And mParams.intCheck = 1, ",����ҽ����¼ F", "")
    
    gstrSQL = gstrSQL & IIf(mParams.bln������, ",��������¼ Q,���������ϸ K ", "")
    
    '''where
    gstrSQL = gstrSQL & " Where A.�Է�����id = H.ID" & IIf(mParams.strSourceDep = "", "", " And A.�Է�����id=E.Column_Value ") & _
        " And C.����id = B.����id And C.����id=P.����ID And C.��ҳid=P.��ҳid And A.���� = S.���� And A.NO = S.NO And Nvl(A.�ⷿid,[1]) = Nvl(S.�ⷿid,[1]) And S.����id = C.ID And c.����id = z.����id(+) And c.Ӥ���� = z.���(+) And C.��ҳid=Z.��ҳid(+) " & _
        IIf(mblnIs�������� And mParams.intCheck = 1, "And F.id(+)=C.ҽ����� And (F.id is null or (F.id is not null and F.ҩʦ��˱�־=1)) ", "") & _
        " And Nvl(A.�ⷿid,[1]) = Nvl(C.ִ�в���id,[1]) And S.ҩƷid = D.ҩƷid And D.ҩ��id = T.ҩ��id And P.��������=W.����(+) " & _
        " And (H.����ʱ�� Is Null Or H.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
        " And A.�������� Between [2] And [3] And S.������� Is Null "
        
    gstrSQL = gstrSQL & IIf(mParams.bln������, " and c.ҽ�����=k.ҽ��id(+) and Q.id(+)=K.��id and K.����ύ(+)=1", "")
    
    'վ�����
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (H.վ�� = [16] Or H.վ�� Is Null) "
    End If
    
    '��ǰҩ��
    gstrSQL = gstrSQL & " And Nvl(A.�ⷿid,[1]) + 0 = [1] "
    
    '¼����Ϣ
    If mcondition.strסԺ�� <> "" Then
        gstrSQL = gstrSQL & " And P.סԺ�� = [4] "
    ElseIf mcondition.str���� <> "" Then
        '���ڴ��Ų�Ψһ��תΪͨ������ID����ѯ
        gstrSQL = gstrSQL & " And C.����ID = [9] "
    ElseIf mcondition.str���￨ <> "" Then
        gstrSQL = gstrSQL & " And B.���￨�� = [6] "
    ElseIf mcondition.str���� <> "" Then
        gstrSQL = gstrSQL & " And P.���� = [7] "
    ElseIf mcondition.strNo <> "" Then
        gstrSQL = gstrSQL & " And A.NO = [8] "
    ElseIf mcondition.lng����ID <> -1 Or (Val(lblPatiInputType.Tag) = mInputType.IC�� And Me.txtInput.Text <> "") Then
        gstrSQL = gstrSQL & " And C.����ID = [9] "
    ElseIf mcondition.str��ҩ�� <> "" Then
        gstrSQL = gstrSQL & " And A.��ҩ�� = [10] "
    ElseIf mcondition.lng��ҩ����ID <> -1 Then
        gstrSQL = gstrSQL & " And A.�Է�����id + 0 = [11] "
    End If
    
    '����ģʽ:0-����,1-���ʵ�,2-���ʱ�
    If mcondition.int����ģʽ = 0 Then
        gstrSQL = gstrSQL & " And A.���� IN(9,10)"
    ElseIf mcondition.int����ģʽ = 1 Then
        gstrSQL = gstrSQL & " And A.����=9"
    ElseIf mcondition.int����ģʽ = 2 Then
        gstrSQL = gstrSQL & " And A.����=10"
    End If
    
    '������
    If mcondition.str������ <> "���м�����" Then
        gstrSQL = gstrSQL & " And S.������ = [12] "
    End If
    
    'ҽ������:0-����,1-����,2-����,3-��ͨ
    '�õ����Ƿ���д�����Ƿ�ҽ��������ҩƷ����
    If mcondition.intҽ������ = 0 Then
    ElseIf mcondition.intҽ������ = 1 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 2 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 3 Then
        gstrSQL = gstrSQL & " And (Nvl(C.ҽ�����,0) + 0 =0 Or S.���� Is Null) "
    ElseIf mcondition.intҽ������ = 4 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_') And Nvl(C.ҽ�����,0) + 0 > 0 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If mcondition.int��ҩ���� = 0 Then
    ElseIf mcondition.int��ҩ���� = 1 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 2 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 3 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 4 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 5 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 6 Then
        gstrSQL = gstrSQL & " And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4')"
    End If
    
    '����Χ
    Select Case mcondition.int����Χ
    Case 1
        gstrSQL = gstrSQL & " And S.ʵ������>=0"
    Case 2
        gstrSQL = gstrSQL & " And S.ʵ������<0"
    End Select
    
    '�������ͣ����˻�Ӥ��
    If mcondition.int�������� = 0 Then
        gstrSQL = gstrSQL & " And Nvl(C.Ӥ����, 0) = 0 "
    ElseIf mcondition.int�������� = 1 Then
        gstrSQL = gstrSQL & " And Nvl(C.Ӥ����, 0) > 0 "
    End If
    
'    '�Ƿ���ʾ��������
'    If mcondition.bln��ʾ��ҩ�������� = False Then
'        gstrSQL = gstrSQL & " And S.��¼״̬ = 1"
'    Else
'        gstrSQL = gstrSQL & " And Mod(S.��¼״̬, 3) = 1"
'    End If
    
    gstrSQL = gstrSQL & " And Mod(S.��¼״̬, 3) = 1"
    
    '��ҩ;��
    If mcondition.str��ҩ;�� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [13] || ',',',' || S.�÷� || ',') > 0 "
    End If
    
    'ҩƷ����
    If mcondition.strҩƷ���� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [14] || ',',',' || T.ҩƷ���� || ',') > 0 "
    End If
    
    '������ҩ����
    If mcondition.str������ҩ���� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [15] || ',',',' || D.��ҩ���� || ',') > 0 "
    End If
    
    '��������
    If Trim(txtInput.Text) = "" Then
        If mParams.intShowDept = 1 Then
            gstrSQL = gstrSQL & " And H.id In (Select ����id From ��������˵�� Where �������� = '�ٴ�' And ������� In (2, 3)) "
        ElseIf mParams.intShowDept = 2 Then
            gstrSQL = gstrSQL & " And H.id In (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����','Ӫ��') And ������� IN(2,3)) "
        ElseIf mParams.intShowDept = 3 Then
            gstrSQL = gstrSQL & " And H.id In (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3)) "
        End If
    End If

    '�ų�������Һ�������Ĺ����в����ĵ���
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = S.ID) "
    
    '�ų���δ��ҩƷ�����ʼ�¼
    If chkWithNotAudited.Value = 0 Then
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ���˷������� X " & _
            " Where X.������� = 0 And X.״̬+0 = 0 And X.�շ�ϸĿid+0 = S.ҩƷid And X.����id = S.����id)"
    End If
    
    '��ΣҩƷ
    If chkDanger.Value = 1 Then
        If chkDangerType(0).Value = 1 Then strDanger = IIf(strDanger = "", 1, strDanger & "," & 1)
        If chkDangerType(1).Value = 1 Then strDanger = IIf(strDanger = "", 2, strDanger & "," & 2)
        If chkDangerType(2).Value = 1 Then strDanger = IIf(strDanger = "", 3, strDanger & "," & 3)
    End If
    
    '�������
    If Me.chkToxicologyType.Value = 1 Then
        If Me.chkToxicology(0).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(0).Caption, strToxicology & "," & Me.chkToxicology(0).Caption)
        If Me.chkToxicology(1).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(1).Caption, strToxicology & "," & Me.chkToxicology(1).Caption)
        If Me.chkToxicology(2).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(2).Caption, strToxicology & "," & Me.chkToxicology(2).Caption)
        If Me.chkToxicology(3).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(3).Caption, strToxicology & "," & Me.chkToxicology(3).Caption)
    End If
    
    If strDanger <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [18] || ',' , ',' || Nvl(D.��ΣҩƷ,0) || ',') > 0 "
    
    If strToxicology <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [19] || ',' , ',' || T.������� || ',') > 0 "
    
    If mParams.blnShowReject = True Then
        '�ϲ��ܷ���¼
        '�ϲ��ܷ���¼
        strTmpSql = " (Select A.����, A.NO, A.����id, A.��ҳid, A.����, Nvl(B.���ȼ�, 0) ���ȼ�, A.�Է�����id, A.�ⷿid, A.��ҩ����, A.��������, A.���շ�, Null As ��ҩ��," & _
                " 0 As ��ӡ״̬, 0 As δ����, A.��Ʒ�ϸ�֤ As ��ҩ�� " & _
                " From (Select B.����, B.NO, A.����id, A.��ҳid, A.����, Decode(A.��¼״̬, 0, 0, 1) ���շ�, B.�Է�����id, B.�ⷿid, " & _
                " B.��ҩ���� , B.��������, C.���, B.��Ʒ�ϸ�֤ " & _
                " From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C " & _
                 IIf(mblnIs�������� And mParams.intCheck = 1, ",����ҽ����¼ F", "") & _
                " Where A.ID = B.����id + 0 And B.���� In (9, 10) And B.������� Is Null And B.ժҪ = '�ܷ�' And " & _
                IIf(mblnIs�������� And mParams.intCheck = 1, " F.id(+)=A.ҽ����� And (F.id is null or (F.id is not null and F.ҩʦ��˱�־=1)) And ", "") & _
                " Nvl(B.�ⷿid,[1]) + 0 = [1] And B.�������� Between [2] And [3] And A.����id = C.����id(+)) A, ��� B " & _
                " Where B.����(+) = A.���) "
        
        strTmpSql = Replace(gstrSQL, "δ��ҩƷ��¼", strTmpSql)
        strTmpSql = Replace(strTmpSql, "0 As �ܷ�", "1 As �ܷ�")
        
        gstrSQL = gstrSQL & " Union All " & strTmpSql
    End If
    
    '''order by
    gstrSQL = gstrSQL & " Order By ��������,�������� desc, ID, ��ҩ��, ����, NO "
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "��ȡ����ҩ���һ���", _
        mcondition.lngҩ��ID, _
        CDate(mcondition.str��ʼʱ��), _
        CDate(mcondition.str����ʱ��), _
        mcondition.strסԺ��, _
        mcondition.str����, _
        mcondition.str���￨, _
        mcondition.str����, _
        mcondition.strNo, _
        mcondition.lng����ID, _
        mcondition.str��ҩ��, _
        mcondition.lng��ҩ����ID, _
        mcondition.str������, _
        mcondition.str��ҩ;��, _
        mcondition.strҩƷ����, _
        mcondition.str������ҩ����, _
        mstrDeptNode, _
        mParams.strSourceDep, _
        strDanger, _
        strToxicology)
        
    If mParams.bln������ Then
        rsData.Filter = "(�����=1 and ���id<>0) or ���id=0"
    End If
    
    '���²�������
    Call GetSendDeptTreeView(rsData)
    
    '���²��������Ӧ���շ���¼���ݼ�
    Call GetDeptListRecord(rsData)
    
    Me.MousePointer = 0
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub RefreshSendDetail()
    'ˢ�´���ҩ��ϸ�б�
    Dim rsData As ADODB.Recordset
    Dim strSql��ҩ�� As String
    Dim str�շ�ID�� As String
    Dim lng��ǰ���� As Long
    Dim str��ǰNO As String
    Dim strSqlTmp As String
    Dim strSqlUnion As String
    Dim i As Integer
    Dim strArr�շ�id As Variant
    Dim ArrTmp As Variant
    Dim IntCount As Integer
    Dim strTmp As String
    
    If Val(tvwList(mDeptType.��ҩ).Tag) = 0 Then Exit Sub
    On Error GoTo errHandle
    '���ݲ����б�ʵ�ʹ�ѡ�����������ID��NO���շ�ID����֯����
    If mrsDeptList Is Nothing Then Exit Sub
    mrsDeptList.Filter = ""
    mrsDeptList.Sort = "����ID,NO,�շ�ID"
    mstr����ID�� = ""
    mstr�������� = ""
    With mrsDeptList
        Do While Not .EOF
            If !ִ��״̬ = 1 Then
                If lng��ǰ���� <> !����ID Then
                    mstr����ID�� = mstr����ID�� & IIf(mstr����ID�� = "", "", ",") & !����ID
                    mstr�������� = mstr�������� & IIf(mstr�������� = "", "", ",") & !��������
                    lng��ǰ���� = !����ID
                End If
                
                If InStr(1, "," & str�շ�ID�� & ",", "," & !�շ�Id & ",") = 0 Then
                    str�շ�ID�� = str�շ�ID�� & IIf(str�շ�ID�� = "", "", ",") & !�շ�Id
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    If str�շ�ID�� = "" Then Exit Sub
    
    '�ֽ��շ�ID��
    '�շ�ID������4Kʱ�ֳ�С��4K�Ĵ����󶨱���ʱ������������Ϊ4K�ַ���
    strArr�շ�id = Array()
    ArrTmp = Split(str�շ�ID�� & ",", ",")
    IntCount = UBound(ArrTmp)
    
    '��ѯ��ʾ
    If WarRecoredCount(IntCount) = False Then Exit Sub

    If Len(str�շ�ID��) >= 4000 Then
        For i = 0 To IntCount
            If ArrTmp(i) <> "" Then
                If Len(IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)) >= 4000 Then
                    ReDim Preserve strArr�շ�id(UBound(strArr�շ�id) + 1)
                    strArr�շ�id(UBound(strArr�շ�id)) = strTmp
                    strTmp = ArrTmp(i)
                Else
                    strTmp = IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)
                End If
            End If
                   
            If i = IntCount Then
                ReDim Preserve strArr�շ�id(UBound(strArr�շ�id) + 1)
                strArr�շ�id(UBound(strArr�շ�id)) = strTmp
            End If
        Next
    Else
        ReDim Preserve strArr�շ�id(UBound(strArr�շ�id) + 1)
        strArr�շ�id(UBound(strArr�շ�id)) = str�շ�ID��
    End If
    
    '''select
    gstrSQL = "SELECT /*+rule*/ Distinct A.*, Nvl(C.��������,0) As ��������, c.����Ǽ�ʱ�� " & IIf(mcondition.bln��ʾ��ҩ��ҩ�� = True, ", B.��ҩ��", ",'' As ��ҩ��") & " FROM ("
    
    strSqlTmp = "SELECT DISTINCT S.ID As �շ�ID,to_char(s.Ч��,'yyyy-mm-dd') Ч��,S.��¼״̬,S.ҩƷID,S.����id,NVL(N.���շ�,0) ���շ�,P.���� ����,S.��ҩ��,C.������ ����ҽ��,C.����Ա���� �����,S.����,S.����, " & _
             " S.NO,S.���,C.����ID,Nvl(C.��ҳID,0) As ��ҳID,Nvl(C.����,'(δ����)') As ����,Decode(Nvl(c.Ӥ����,0), 0, Nvl(Q.����, C.����), U.Ӥ������) ����,Decode(Nvl(c.Ӥ����,0), 0, Nvl(Q.�Ա�, C.�Ա�), U.Ӥ���Ա�) �Ա�,C.�����־,C.��ʶ��,C.����Ա����,S.���� ��,S.ʵ������ ����," & _
             " NVL(D.ҩ������,0) ����,Nvl(D.��ΣҩƷ,0) As ��ΣҩƷ,X.���,T.�������,T.��ֵ����,Nvl(T.������,0) ������,C.�Ǽ�ʱ��,H.���� As ��ҩ���ű���,H.���� As ��ҩ����,H.Id As ��ҩ����Id," & _
             " S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����, Ceil((s.ʵ������ * d.����ϵ��) / Nvl(s.����, 1)) As ��������," & _
             " C.ҽ�����,I.���㵥λ,NVL(S.����,NVL(X.����,'')) ����,nvl(M.�����,-1) �����,M.Ƥ�Խ��,Nvl(M.����ʱ��,C.�Ǽ�ʱ��) As ����ʱ��,decode(m.��ҩĿ��,1,'Ԥ��',2,'����',3,'Ԥ��������','') ��ҩĿ��,m.��ҩ����,D.ҩ��ID,nvl(C.ҽ�����,-1) ҽ��id," & IIf(mParams.blnҩƷ���� = True, "L.", "'' ") & "�ⷿ��λ," & _
             " M.���ID,M.ҩʦ��˱�־,C.���˿���ID As ����ID,C.��� �������," & IIf(mParams.blnҩƷ���� = True, "Decode(Sign(Nvl(K.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) ", "0 ") & " �������, Z.���� As Ӣ����, Decode(Nvl(c.Ӥ����,0), 0, Q.����, Ceil(Sysdate - U.����ʱ��) || '��') ����,Q.��������,W.��ɫ,N.��ҩ��, " & _
             IIf(mParams.intҩƷ������ʾ = 0 Or mParams.intҩƷ������ʾ = 2, "NVL(E.����,'')", "Decode(E.����,Null,'',X.����)") & " As ������, " & _
             "'['||X.����||']'||" & IIf(mParams.intҩƷ������ʾ = 1, "NVL(E.����,X.����)", "X.����") & " As Ʒ��,nvl(N.����,'') �䷽����," & _
             "X.����" & " As ҩƷ����," & IIf(mParams.intҩƷ������ʾ = 1, "NVL(E.����,X.����)", "X.����") & " As ҩƷ����,s.������id,s.���ϵ��,s.������,s.��������,s.��ҩ����,Nvl(t.�Ƿ�Ƥ��,0) As �Ƿ�Ƥ��,F.ִ�з���,F.���,D.����ϵ��, m.ҽ������, m.�����λ, m.ִ��ʱ�䷽��, m.Ƶ�ʴ���, m.��ʼִ��ʱ��, j.�״�ʱ��, j.ĩ��ʱ��, m.Ƶ�ʼ��, m.�ϴδ�ӡʱ��, m.ҽ����Ч "
'    '���Է��飨���ID��Ϊ1��
'    strSqlTmp = "SELECT DISTINCT S.ID As �շ�ID,S.��¼״̬,S.ҩƷID,NVL(N.���շ�,0) ���շ�,P.���� ����,S.��ҩ��,C.������ ����ҽ��,C.����Ա���� �����,S.����,S.����," & _
'             " S.NO,S.���,C.����ID,C.����,C.����,C.�����־,C.��ʶ��,C.����Ա����,S.���� ��,S.ʵ������ ����," & _
'             " NVL(D.ҩ������,0) ����,X.���,T.�������,T.��ֵ����,C.�Ǽ�ʱ��,H.���� As ��ҩ����,H.Id As ��ҩ����Id," & _
'             " S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����," & _
'             " C.ҽ�����,I.���㵥λ,NVL(S.����,NVL(X.����,'')) ����,nvl(M.�����,-1) �����,nvl(C.ҽ�����,-1) ҽ��id," & IIf(mParams.blnҩƷ���� = True, "L.", "'' ") & "�ⷿ��λ," & _
'             " 1 ���ID,C.���˿���ID As ����ID,C.��� �������," & IIf(mParams.blnҩƷ���� = True, "Decode(Sign(Nvl(K.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) ", "0 ") & " �������, Z.���� As Ӣ����, R.����, N.��ҩ��, " & _
'             IIf(mParams.intҩƷ������ʾ = 0 Or mParams.intҩƷ������ʾ = 2, "NVL(E.����,'')", "Decode(E.����,Null,'',X.����)") & " As ������, " & _
'             "'['||X.����||']'||" & IIf(mParams.intҩƷ������ʾ = 1, "NVL(E.����,X.����)", "X.����") & " As Ʒ��," & _
'             "X.����" & " As ҩƷ����," & IIf(mParams.intҩƷ������ʾ = 1, "NVL(E.����,X.����)", "X.����") & " As ҩƷ����,s.������id,s.���ϵ��,s.������,s.��������,s.��ҩ����"
           
    '��λ����
    Select Case mParams.strUnit
    Case "�ۼ۵�λ"
        strSqlTmp = strSqlTmp & ",X.���㵥λ ��λ,1 ��װ "
    Case "���ﵥλ"
        strSqlTmp = strSqlTmp & ",D.���ﵥλ ��λ,D.�����װ ��װ "
    Case "סԺ��λ"
        strSqlTmp = strSqlTmp & ",D.סԺ��λ ��λ,D.סԺ��װ ��װ "
    Case "ҩ�ⵥλ"
        strSqlTmp = strSqlTmp & ",D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ "
    End Select
    
    'ȱҩ���
    If mParams.blnȱҩ��� = True Then
        strSqlTmp = strSqlTmp & " ,A.ʵ������ As ������� "
    Else
        strSqlTmp = strSqlTmp & " ,0 As ������� "
    End If
    
    '''from
    strSqlTmp = strSqlTmp & " FROM ҩƷ�շ���¼ S,סԺ���ü�¼ C,����ҽ����¼ M,����ҽ����¼ G,δ��ҩƷ��¼ N,�շ���Ŀ���� E,�շ���ĿĿ¼ X,������ĿĿ¼ I,������ĿĿ¼ N,������ĿĿ¼ F,����ҽ������ J," & _
             " ҩƷ��� D,ҩƷ���� T," & IIf(mParams.blnҩƷ���� = True, "ҩƷ�����޶� L,", "") & "������Ŀ���� Z,���ű� P,���ű� H,������Ϣ R,������ҳ Q,�������� W,������������¼ U "
             
    '���շ�ID�����ǵģ��������շ�ID��Ϊ����
    strSqlTmp = strSqlTmp & " ,Table(Cast(f_Num2List([15]) As zlTools.t_NumList)) G "
    
    If mParams.blnҩƷ���� = True Then
        strSqlTmp = strSqlTmp & ",(Select �ⷿid, ҩƷid, Nvl(Sum(ʵ������), 0) ������� From ҩƷ��� Where ���� = 1 And �ⷿid = [1] Group By �ⷿid, ҩƷid) K "
    End If
    
    If mParams.blnȱҩ��� = True Then
        strSqlTmp = strSqlTmp & ",(Select �ⷿid, ҩƷid, ʵ������, Nvl(����, 0) ���� From ҩƷ��� Where ���� = 1 And �ⷿid = [1]) A "
    End If
             
    strSqlTmp = strSqlTmp & " WHERE S.NO=N.NO AND S.����=N.���� AND NVL(S.�ⷿID,[1])+0=NVL(N.�ⷿID,[1]) AND S.����ID=C.ID And S.ҩƷID=D.ҩƷID And c.����id = u.����id(+) And c.Ӥ���� = u.���(+) And C.��ҳid=U.��ҳid(+) " & _
            " And C.����id = R.����id And C.����id=Q.����id And C.��ҳid=Q.��ҳid And Q.��������=W.����(+) And j.ҽ��id(+) = c.ҽ����� And j.No(+) = c.No And j.��¼����(+) = c.��¼���� " & _
            " AND S.�Է�����ID+0=H.ID AND S.����� IS NULL AND NVL(S.�ⷿID,[1])+0=[1] " & _
            " AND C.���˿���ID=P.id And d.ҩƷID=X.ID and D.ҩ��ID=T.ҩ��ID AND D.ҩ��ID=I.ID and C.ҽ�����=M.ID(+) and M.���id=G.id(+) and G.�䷽id=N.id(+) and G.������Ŀid=F.id(+) " & _
            " And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 " & IIf(mParams.blnҩƷ���� = True, " And S.ҩƷID=L.ҩƷID(+) And Nvl(S.�ⷿID,[1])=L.�ⷿID(+) ", "") & _
            " AND D.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
            " And nvl(S.��ҩ��ʽ,-999)<>-1 " & _
            " And S.���� In(9,10)  And N.�������� Between [2] And [3] "
    
    strSqlTmp = strSqlTmp & " And S.ID= G.Column_Value "
    
    If mParams.blnҩƷ���� = True Then
        strSqlTmp = strSqlTmp & " And Nvl(S.�ⷿid, [1]) + 0 = K.�ⷿid(+) And S.ҩƷid = K.ҩƷid(+) "
    End If
    
    If mParams.blnȱҩ��� = True Then
        strSqlTmp = strSqlTmp & " And Nvl(S.�ⷿid, [1]) + 0 = A.�ⷿid(+) And S.ҩƷid = A.ҩƷid(+) And Nvl(S.����, 0) = A.����(+) "
    End If
    
    '¼����Ϣ
    If mcondition.strסԺ�� <> "" Then
        strSqlTmp = strSqlTmp & " And Q.סԺ�� = [8] "
    ElseIf mcondition.str���� <> "" Then
        strSqlTmp = strSqlTmp & " And R.��ǰ���� = [9] "
    ElseIf mcondition.str���￨ <> "" Then
        strSqlTmp = strSqlTmp & " And R.���￨�� = [10] "
    ElseIf mcondition.str���� <> "" Then
        strSqlTmp = strSqlTmp & " And N.���� = [11] "
    ElseIf mcondition.strNo <> "" Then
        strSqlTmp = strSqlTmp & " And N.NO = [12] "
    ElseIf mcondition.lng����ID <> -1 Then
        strSqlTmp = strSqlTmp & " And N.����ID = [13] "
    ElseIf mcondition.str��ҩ�� <> "" Then
        strSqlTmp = strSqlTmp & " And N.��ҩ�� = [14] "
    End If
    
    '����ģʽ:0-����,1-���ʵ�,2-���ʱ�
    If mcondition.int����ģʽ = 1 Then
        strSqlTmp = strSqlTmp & " And S.����=9"
    ElseIf mcondition.int����ģʽ = 2 Then
        strSqlTmp = strSqlTmp & " And S.����=10"
    End If
    
    '������
    If mcondition.str������ <> "���м�����" Then
        strSqlTmp = strSqlTmp & " And S.������ = [7] "
    End If
    
    'ҽ������:0-����,1-����,2-����,3-��ͨ
    '�õ����Ƿ���д�����Ƿ�ҽ��������ҩƷ����
    If mcondition.intҽ������ = 0 Then
    ElseIf mcondition.intҽ������ = 1 Then
        strSqlTmp = strSqlTmp & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 2 Then
        strSqlTmp = strSqlTmp & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 3 Then
        strSqlTmp = strSqlTmp & " And (Nvl(C.ҽ�����,0) + 0 =0 Or S.���� Is Null) "
    ElseIf mcondition.intҽ������ = 4 Then
        strSqlTmp = strSqlTmp & " And S.���� Is Not Null And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_') And Nvl(C.ҽ�����,0) + 0 > 0 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If mcondition.int��ҩ���� = 0 Then
    ElseIf mcondition.int��ҩ���� = 1 Then
        strSqlTmp = strSqlTmp & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 2 Then
        strSqlTmp = strSqlTmp & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 3 Then
        strSqlTmp = strSqlTmp & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 4 Then
        strSqlTmp = strSqlTmp & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 5 Then
        strSqlTmp = strSqlTmp & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 6 Then
        strSqlTmp = strSqlTmp & " And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4')"
    End If
    
    '����Χ
    Select Case mcondition.int����Χ
    Case 1
        strSqlTmp = strSqlTmp & " And S.ʵ������>=0"
    Case 2
        strSqlTmp = strSqlTmp & " And S.ʵ������<0"
    End Select
    
    '�������ͣ����˻�Ӥ��
    If mcondition.int�������� = 0 Then
        strSqlTmp = strSqlTmp & " And Nvl(C.Ӥ����, 0) = 0 "
    ElseIf mcondition.int�������� = 1 Then
        strSqlTmp = strSqlTmp & " And Nvl(C.Ӥ����, 0) > 0 "
    End If
    
    '��ҩ;��
    If mcondition.str��ҩ;�� <> "" Then
        strSqlTmp = strSqlTmp & " And Instr(',' || [4] || ',',',' || S.�÷� || ',') > 0 "
    End If
    
    'ҩƷ����
    If mcondition.strҩƷ���� <> "" Then
        strSqlTmp = strSqlTmp & " And Instr(',' || [5] || ',',',' || T.ҩƷ���� || ',') > 0 "
    End If
    
    '������ҩ����
    If mcondition.str������ҩ���� <> "" Then
        strSqlTmp = strSqlTmp & " And Instr(',' || [6] || ',',',' || D.��ҩ���� || ',') > 0 "
    End If
    
    '��������
    If Trim(txtInput.Text) = "" Then
    
    
        If mParams.intShowDept = 1 Then
            strSqlTmp = strSqlTmp & " And H.id In (Select ����id From ��������˵�� Where �������� = '�ٴ�' And ������� In (2, 3)) "
        ElseIf mParams.intShowDept = 2 Then
            strSqlTmp = strSqlTmp & " And H.id In (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����','Ӫ��') And ������� IN(2,3)) "
        ElseIf mParams.intShowDept = 3 Then
            strSqlTmp = strSqlTmp & " And H.id In (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3)) "
        End If
    End If
    
    '�ų�������Һ�������Ĺ����в����ĵ���
    strSqlTmp = strSqlTmp & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = S.ID) "
    
    '�ϲ��ܷ���¼
    strSqlUnion = " (Select A.����, A.NO, A.����id, A.��ҳid, A.����, Nvl(B.���ȼ�, 0) ���ȼ�, A.�Է�����id, A.�ⷿid, A.��ҩ����, A.��������, A.���շ�, Null As ��ҩ��," & _
            " 0 As ��ӡ״̬, 0 As δ����, A.��Ʒ�ϸ�֤ As ��ҩ�� " & _
            " From (Select B.����, B.NO, A.����id, Nvl(A.��ҳID,0) As ��ҳID,A.����, Decode(A.��¼״̬, 0, 0, 1) ���շ�, B.�Է�����id, B.�ⷿid, " & _
            " B.��ҩ���� , B.��������, C.���, B.��Ʒ�ϸ�֤ " & _
            " From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C " & _
            " Where A.ID = B.����id + 0 And B.������� Is Null And B.ժҪ = '�ܷ�' And " & _
            " Nvl(B.�ⷿid,[1]) = [1] And B.�������� Between [2] And [3] And A.����id = C.����id(+)) A, ��� B " & _
            " Where B.����(+) = A.���) "
            
    strSqlTmp = strSqlTmp & " Union All " & Replace(strSqlTmp, "δ��ҩƷ��¼", strSqlUnion)
    
    gstrSQL = gstrSQL & strSqlTmp & ") A "
    
    gstrSQL = gstrSQL & ", (Select ҩƷid,�ⷿid,����id,Nvl(ʵ������, ��������) As ��������, �Ǽ�ʱ�� As ����Ǽ�ʱ�� From ҩƷ����ƻ�  Where ״̬ <> 1) C "
    
    '�����һ����ҩ����ҩ��
    If mcondition.bln��ʾ��ҩ��ҩ�� = True Then
        strSql��ҩ�� = ",(Select a.���� ,a.No,a.���,a.������ ��ҩ�� From ҩƷ�շ���¼ a," & _
                " (Select s.����,s.No,s.���, Max(s.��¼״̬) ��¼״̬ " & _
                " From ҩƷ�շ���¼ s, δ��ҩƷ��¼ n " & _
                " Where s.No = n.No And s.���� = n.���� And Nvl(s.�ⷿid, [1]) + 0 = Nvl(n.�ⷿid, [1]) And " & _
                " Nvl(s.�ⷿid, [1]) + 0 = [1] " & _
                " And Nvl(s.��ҩ��ʽ, -999) <> -1 And " & _
                " Mod(s.��¼״̬, 3) = 2 And s.���� In (9, 10) " & _
                " Group By s.����,s.No,s.���) b " & _
                " Where a.����=b.���� And a.No=b.No And a.���=b.��� And a.��¼״̬=b.��¼״̬) B "
        gstrSQL = gstrSQL & strSql��ҩ��
    End If
    
    gstrSQL = gstrSQL & " Where A.��ҩ����id = C.����id(+) And C.�ⷿid(+) = [1] And A.ҩƷid = C.ҩƷid(+) "
    
    If mcondition.bln��ʾ��ҩ��ҩ�� = True Then
        gstrSQL = gstrSQL & " And A.���� = B.����(+) And A.No = B.No(+) And A.��� = B.���(+) "
    End If
    
    '�ų���δ��ҩƷ�����ʼ�¼
    If chkWithNotAudited.Value = 0 Then
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ���˷������� X " & _
            " Where X.������� = 0 And X.״̬+0 = 0 And X.�շ�ϸĿid+0 = A.ҩƷid And X.����id = A.����id) "
    End If
    
    gstrSQL = gstrSQL & "  Order By a.����,a.No,a.������� "
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    Call AviShow(Me)
    
    Call InitSendRec
    
    '�����շ�ID����������Ŀ��ѭ��ִ��
    For i = 0 To UBound(strArr�շ�id)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", _
            mcondition.lngҩ��ID, CDate(mcondition.str��ʼʱ��), CDate(mcondition.str����ʱ��), mcondition.str��ҩ;��, mcondition.strҩƷ����, _
            mcondition.str������ҩ����, mcondition.str������, mcondition.strסԺ��, mcondition.str����, mcondition.str���￨, _
            mcondition.str����, mcondition.strNo, mcondition.lng����ID, mcondition.str��ҩ��, _
            CStr(strArr�շ�id(i)))
            
        If Not rsData.EOF Then
            'װ�ط�ҩ���ݼ�
            If LoadSendRecord(rsData) = False Then
                Me.MousePointer = 0
                Exit Sub
            End If
        End If
    Next
    
    If mrsSendData.RecordCount > 0 Then
        'װ���������ݼ�
        Call RefreshChargeOffDetail
        '���Ӵ��崫�����ݼ�
        Call mfrmDetail.RefreshList(mListType.��ҩ, mrsSendData, mrsChargeOff)
    End If
    
    Me.MousePointer = 0
    Call AviShow(Me, False)
    Exit Sub
errHandle:
    Me.MousePointer = 0
    Call AviShow(Me, False)
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshReturnDetail()
    'ˢ����ҩ��ϸ�б�
    Dim rsData As ADODB.Recordset
    Dim strSql��ҩ�� As String
    Dim strSqlSelect As String
    Dim i As Integer
    Dim strArr�շ�id As Variant
    Dim ArrTmp As Variant
    Dim IntCount As Integer
    Dim strTmp As String
    Dim str�շ�ID�� As String
    
    If Val(tvwList(mDeptType.��ҩ).Tag) = 0 Then Exit Sub
    
    '���ݲ����б��ѹ�ѡ�������֯��Ҫ������
    If mrsDeptList Is Nothing Then Exit Sub
    mrsDeptList.Filter = ""
    With mrsDeptList
        Do While Not .EOF
            If !ִ��״̬ = 1 Then
                If InStr(1, "," & str�շ�ID�� & ",", "," & !�շ�Id & ",") = 0 Then
                    str�շ�ID�� = str�շ�ID�� & IIf(str�շ�ID�� = "", "", ",") & !�շ�Id
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    If str�շ�ID�� = "" Then Exit Sub
    
    '�ֽ��շ�ID��
    '�շ�ID������4Kʱ�ֳ�С��4K�Ĵ����󶨱���ʱ������������Ϊ4K�ַ���
    strArr�շ�id = Array()
    ArrTmp = Split(str�շ�ID�� & ",", ",")
    IntCount = UBound(ArrTmp)
    
    '��ѯ��ʾ
    If WarRecoredCount(IntCount) = False Then Exit Sub
    
    If Len(str�շ�ID��) >= 4000 Then
        For i = 0 To IntCount
            If ArrTmp(i) <> "" Then
                If Len(IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)) >= 4000 Then
                    ReDim Preserve strArr�շ�id(UBound(strArr�շ�id) + 1)
                    strArr�շ�id(UBound(strArr�շ�id)) = strTmp
                    strTmp = ArrTmp(i)
                Else
                    strTmp = IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)
                End If
            End If
                   
            If i = IntCount Then
                ReDim Preserve strArr�շ�id(UBound(strArr�շ�id) + 1)
                strArr�շ�id(UBound(strArr�շ�id)) = strTmp
            End If
        Next
    Else
        ReDim Preserve strArr�շ�id(UBound(strArr�շ�id) + 1)
        strArr�շ�id(UBound(strArr�շ�id)) = str�շ�ID��
    End If
    
    
    '��λ����
    Select Case mParams.strUnit
    Case "�ۼ۵�λ"
        strSqlSelect = "X.���㵥λ ��λ,1 ��װ,"
    Case "���ﵥλ"
        strSqlSelect = "D.���ﵥλ ��λ,D.�����װ ��װ,"
    Case "סԺ��λ"
        strSqlSelect = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,"
    Case "ҩ�ⵥλ"
        strSqlSelect = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,"
    End Select
        
    strSqlSelect = strSqlSelect & IIf(mParams.intҩƷ������ʾ = 0 Or mParams.intҩƷ������ʾ = 2, "NVL(A.����,'')", "Decode(A.����,Null,'',X.����)") & " As ������, " & _
             "'['||X.����||']'||" & IIf(mParams.intҩƷ������ʾ = 1, "NVL(A.����,X.����)", "X.����") & " As Ʒ��," & _
             "X.����" & " As ҩƷ����," & IIf(mParams.intҩƷ������ʾ = 1, "NVL(A.����,X.����)", "X.����") & " As ҩƷ����,"

    gstrSQL = " SELECT /*+rule*/ DISTINCT S.ID As �շ�ID,S.����,S.ҩƷID,S.NO,S.���,S.����,H.ID As ��ҩ����ID,P.���� ����,C.�����־,C.��ʶ��,C.����ID,C.����,Decode(Nvl(c.Ӥ����,0), 0, Nvl(W.����, C.����), U.Ӥ������) ����,Decode(Nvl(c.Ӥ����,0), 0, Nvl(W.�Ա�, C.�Ա�), U.Ӥ���Ա�) �Ա�," & _
             " NVL(D.ҩ������,0) ����,Nvl(D.��ΣҩƷ,0) As ��ΣҩƷ,X.���,T.�������,TO_CHAR(Q.����ʱ��,'YYYY-MM-DD HH24:MI:SS') ����ʱ��," & _
             strSqlSelect & _
             " S.���� ��,S.ʵ������ ����,S.��������,S.�ѷ����� ׼����,DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����,to_char(S.Ч��,'yyyy-mm-dd') Ч��," & _
             " S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ��ҩʱ��,S.�����,S.�������,�ɲ���,C.ҽ�����,I.���㵥λ," & _
             " NVL(S.����,NVL(X.����,'')) ����,nvl(M.�����,-1) �����,nvl(C.ҽ�����,-1) ҽ��id,S.��ҩ��," & IIf(mParams.blnҩƷ���� = True, "L.", "'' ") & "�ⷿ��λ, " & _
             " M.���ID,c.��� �������,Z.���� As Ӣ����,0 As ת��, S.��ҩ��,D.����ϵ��,m.ҽ������ " & IIf(mParams.bln����������ʱ���ܽ�����ҩ����, " , Decode(e.״̬, 0, 1, 0) As �������� ", "") & _
             " FROM "
    gstrSQL = gstrSQL & _
             "          (SELECT A.ID,A.NO,A.����,A.���,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0) ����," & _
             "              NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
             "              A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID,1 �ɲ���,A.����," & _
             "              decode(nvl(A.������,''),'','',Decode(A.��¼״̬,1,'(��)'||A.������," & _
             "              decode(Mod(A.��¼״̬,3),0,'(��)'||A.������,1,'(��)'||A.������,2,'(��)'||A.������))) ��ҩ��,Nvl(A.���ܷ�ҩ��, 0) ��ҩ��,A.������ " & _
             "          FROM ҩƷ�շ���¼ A," & _
             "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
             "          FROM ҩƷ�շ���¼ A,Table(Cast(f_Num2List([15]) As zlTools.t_NumList)) G " & _
             "          WHERE A.ID= G.Column_Value And A.����� IS NOT NULL" & _
             "          AND A.�ⷿID+0=[1] AND A.������� BETWEEN [2] AND [3] " & _
             "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B" & _
             "          WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� And A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0) "
    gstrSQL = gstrSQL & _
             "          UNION" & _
             "          SELECT A.ID,A.NO,A.����,A.���,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0)," & _
             "          NVL(A.����,1) ����,A.ʵ������,0 ������,0 �ѷ�����,A.��¼״̬," & _
             "          A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID," & _
             "          DECODE(A.��¼״̬,1,1,DECODE(MOD(A.��¼״̬,3),0,1,MOD(A.��¼״̬,3)+1)) �ɲ���,A.����," & _
             "          decode(nvl(A.������,''),'','',Decode(A.��¼״̬,1,'(��)'||A.������," & _
             "          decode(Mod(A.��¼״̬,3),0,'(��)'||A.������,1,'(��)'||A.������,2,'(��)'||A.������))) ��ҩ��,Nvl(A.���ܷ�ҩ��, 0) ��ҩ��,A.������ " & _
             "          FROM ҩƷ�շ���¼ A,Table(Cast(f_Num2List([15]) As zlTools.t_NumList)) G " & _
             "          WHERE A.ID= G.Column_Value And A.����� IS NOT NULL AND NOT (��¼״̬=1 OR MOD(��¼״̬,3)=0)" & _
             "          AND A.�ⷿID+0=[1] AND A.������� BETWEEN [2] AND [3] " & _
             "          ) S,"
    gstrSQL = gstrSQL & "" & _
             "      סԺ���ü�¼ C,���ű� P,ҩƷ��� D,�շ���ĿĿ¼ X,�շ���Ŀ���� A,ҩƷ���� T,������ĿĿ¼ I,����ҽ����¼ M,����ҽ������ Q,������Ϣ R,������ҳ W," & IIf(mParams.blnҩƷ���� = True, "ҩƷ�����޶� L,", "") & "������Ŀ���� Z,���ű� H, ������������¼ U " & IIf(mParams.bln����������ʱ���ܽ�����ҩ����, " , (Select ����id, ״̬ From ���˷������� Where ״̬ = 0) E ", "")
     
     
    '''where
    gstrSQL = gstrSQL & " WHERE S.ҩƷID=D.ҩƷID And C.����id=W.����id And C.��ҳid=W.��ҳid AND D.ҩ��ID=T.ҩ��ID AND d.ҩƷID=x.ID AND C.���˿���ID+0=P.ID AND D.ҩ��ID=I.ID and C.ҽ�����=M.ID(+)  and C.ҽ�����=Q.ҽ��id(+) And c.No = q.No(+) And c.����id = u.����id(+) And c.Ӥ���� = u.���(+) And C.��ҳid=U.��ҳid(+) " & _
             " And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 " & IIf(mParams.blnҩƷ���� = True, " And S.ҩƷID=L.ҩƷID(+) And Nvl(S.�ⷿID,[1])=L.�ⷿID(+) ", "") & _
             " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 And C.����id = R.����id " & _
             " AND S.����ID=C.ID And S.���� IN(9,10) " & _
             " AND S.����� IS NOT NULL And s.�Է�����id + 0 = h.Id " & IIf(mParams.bln����������ʱ���ܽ�����ҩ����, " And s.����id = e.����id(+) ", "")
    
    '¼����Ϣ
    If mcondition.strסԺ�� <> "" Then
        gstrSQL = gstrSQL & " And W.סԺ�� = [8] "
    ElseIf mcondition.str���� <> "" Then
        gstrSQL = gstrSQL & " And R.��ǰ���� = [9] "
    ElseIf mcondition.str���￨ <> "" Then
        gstrSQL = gstrSQL & " And R.���￨�� = [10] "
    ElseIf mcondition.str���� <> "" Then
        gstrSQL = gstrSQL & " And C.���� = [11] "
    ElseIf mcondition.strNo <> "" Then
        gstrSQL = gstrSQL & " And C.NO = [12] "
    ElseIf mcondition.lng����ID <> -1 Then
        gstrSQL = gstrSQL & " And C.����ID = [13] "
    ElseIf mcondition.cur��ҩ�� <> 0 Then
        gstrSQL = gstrSQL & " And S.��ҩ�� = [14] "
    End If
    
    '����ģʽ:0-����,1-���ʵ�,2-���ʱ�
    If mcondition.int����ģʽ = 1 Then
        gstrSQL = gstrSQL & " And S.����=9"
    ElseIf mcondition.int����ģʽ = 2 Then
        gstrSQL = gstrSQL & " And S.����=10"
    End If
    
    '������
    If mcondition.str������ <> "���м�����" Then
        gstrSQL = gstrSQL & " And S.������ = [7] "
    End If
    
    'ҽ������:0-����,1-����,2-����,3-��ͨ
    '�õ����Ƿ���д�����Ƿ�ҽ��������ҩƷ����
    If mcondition.intҽ������ = 0 Then
    ElseIf mcondition.intҽ������ = 1 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 2 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 3 Then
        gstrSQL = gstrSQL & " And (Nvl(C.ҽ�����,0) + 0 =0 Or S.���� Is Null) "
    ElseIf mcondition.intҽ������ = 4 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_') And Nvl(C.ҽ�����,0) + 0 > 0 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If mcondition.int��ҩ���� = 0 Then
    ElseIf mcondition.int��ҩ���� = 1 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 2 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 3 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 4 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 5 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 6 Then
        gstrSQL = gstrSQL & " And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4')"
    End If
    
    '�������ͣ����˻�Ӥ��
    If mcondition.int�������� = 0 Then
        gstrSQL = gstrSQL & " And Nvl(C.Ӥ����, 0) = 0 "
    ElseIf mcondition.int�������� = 1 Then
        gstrSQL = gstrSQL & " And Nvl(C.Ӥ����, 0) > 0 "
    End If
    
    '��ҩ;��
    If mcondition.str��ҩ;�� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [4] || ',',',' || S.�÷� || ',') > 0 "
    End If
    
    'ҩƷ����
    If mcondition.strҩƷ���� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [5] || ',',',' || T.ҩƷ���� || ',') > 0 "
    End If
    
    '������ҩ����
    If mcondition.str������ҩ���� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [6] || ',',',' || D.��ҩ���� || ',') > 0 "
    End If
    
    Dim blnMoved As Boolean
    Dim strSQL As String
    '�ж��Ƿ���ڲ���������ת��
    blnMoved = zlDatabase.DateMoved(mcondition.str��ʼʱ��)
    If blnMoved Then
        'SQL����¼��Ż��ܣ����κ�һ����ϸҪô���ߣ�Ҫô�󱸣���ˣ���UNION��ʽ����
        strSQL = gstrSQL
        strSQL = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
        strSQL = Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼")
        strSQL = Replace(strSQL, "0 As ת��", "1 As ת��")
        
        gstrSQL = gstrSQL & " UNION ALL " & strSQL
    End If
    
    gstrSQL = gstrSQL & " Order By No,����,�������"
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    Call AviShow(Me)
    Call InitReturnRec
    
    '�����շ�ID����������Ŀ��ѭ��ִ��
    For i = 0 To UBound(strArr�շ�id)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mcondition.lngҩ��ID, _
            CDate(mcondition.str��ʼʱ��), _
            CDate(mcondition.str����ʱ��), _
            mcondition.str��ҩ;��, _
            mcondition.strҩƷ����, _
            mcondition.str������ҩ����, _
            mcondition.str������, _
            mcondition.strסԺ��, _
            mcondition.str����, _
            mcondition.str���￨, _
            mcondition.str����, _
            mcondition.strNo, _
            mcondition.lng����ID, _
            mcondition.cur��ҩ��, _
            CStr(strArr�շ�id(i)))
        
        If Not rsData.EOF Then
            If LoadReturnRecord(rsData) = False Then
                Me.MousePointer = 0
                Exit Sub
            End If
        End If
    Next
    
    If mrsReturnData.RecordCount > 0 Then
        Call mfrmDetail.RefreshList(mListType.��ҩ, mrsReturnData)
    End If
    
    Me.MousePointer = 0
    Call AviShow(Me, False)
    Exit Sub
errHandle:
    Me.MousePointer = 0
    Call AviShow(Me, False)
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'
Private Sub GetParams()
    'ȡ��ģ���õ��Ĳ�����Ϣ
    Dim int��� As Integer
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "select ���� from ҩƷ���ľ��� where ����=0 and ��� = 1 And ���� = 4 And ��λ = 5"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����")
    If rstemp.RecordCount = 0 Then
        int��� = 2
    Else
        int��� = rstemp!����
    End If
    With mParams
        '�������е�ϵͳ����
        .bln����δ��˴�����ҩ = (gtype_UserSysParms.P6_δ��˼��ʴ�����ҩ = 1)
        .bln����ҽ�������Ϻ���ҩ = (gtype_UserSysParms.P68_����ҩ�������Ϻ���ҩ = 1)
        .int����λ�� = int���
        .bln��˻��۵� = (gtype_UserSysParms.P81_ִ�к��Զ���˻��۵� = 1)
        .intЧ����ʾ��ʽ = gtype_UserSysParms.P149_Ч����ʾ��ʽ
        .intҩƷ������ʾ = gintҩƷ������ʾ
        
        '���������еĲ���
        '��������
        .lngҩ��ID = Val(zlDatabase.GetPara("��ҩҩ��", glngSys, 1342))
        .int����ģʽ = Val(zlDatabase.GetPara("����ģʽ", glngSys, 1342))
        .intDays = Val(zlDatabase.GetPara("��ѯ����", glngSys, 1342)) - 1
        .int�Զ�ˢ��δ��ҩ�嵥 = Val(zlDatabase.GetPara("�Զ�ˢ��δ��ҩ�嵥", glngSys, 1342))
        .str������ = zlDatabase.GetPara("������", glngSys, 1342, "���м�����")
        .bln���ܷ�ҩ = (Val(zlDatabase.GetPara("��ҩʱ������ҩ���ʼ�¼", glngSys, 1342, 0)) = 1)
        .bln������ʾ = (Val(zlDatabase.GetPara("�����һ�����ʾ�����嵥", glngSys, 1342)) = 1)
        .int��ҩ���� = Val(zlDatabase.GetPara("��ҩ����", glngSys, 1342))
        .bln��ҩ��ǩ�� = (Val(zlDatabase.GetPara("��ҩ��ǩ��", glngSys, 1342)) = 1)
        .bln��ҩ��ǩ�� = (Val(zlDatabase.GetPara("��ҩ��ǩ��", glngSys, 1342)) = 1)
        .bln��˳�Ժ�������� = (Val(zlDatabase.GetPara("��˳�Ժ���˵���������", glngSys, 1342, 0)) = 1)
        .bln������ = (gtype_UserSysParms.P240_ҩ��������� = 2 Or gtype_UserSysParms.P240_ҩ��������� = 3)
        .bln����������ʱ���ܽ�����ҩ���� = (Val(zlDatabase.GetPara("����������ʱ���ܽ�����ҩ����", glngSys, 1342)) = 1)
        .bln���ط�ҩʱ�����ҩ���� = (Val(zlDatabase.GetPara("���ط�ҩʱ�����ҩ����", glngSys, 1342, 0)) = 1)
        
        '��������
        .blnȱҩ��� = (Val(zlDatabase.GetPara("ȱҩ���", glngSys, 1342, 1)) = 1)
        .int�Զ���ӡ = Val(zlDatabase.GetPara("�Զ���ӡ", glngSys, 1342))
        .blnҩƷ���� = (Val(zlDatabase.GetPara("�ⷿ��λ�����������ʾ", glngSys, 1342, 0)) = 1)
        .str������� = zlDatabase.GetPara("�������", glngSys, 1342)
        .str��ֵ���� = zlDatabase.GetPara("��ֵ����", glngSys, 1342)
        .str��Σ���� = zlDatabase.GetPara("��Σ����", glngSys, 1342, "")
        .str��Σ���� = zlDatabase.GetPara("��ΣҩƷ����", glngSys, 1342, "")
        .int��ҩ�嵥��ӡ = Val(zlDatabase.GetPara("��ӡ��ҩ�嵥", glngSys, 1342))
        .intCheck = Val(zlDatabase.GetPara("��˸�ҩ������������", glngSys, 1345))
        .int��ҩʱ���ҽ�� = Val(zlDatabase.GetPara("��ҩʱ���ҽ��", glngSys, 1342))
        
        .intҽ������ = Val(zlDatabase.GetPara("ҽ������", glngSys, 1342))
        
        .bln�������� = (Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\Frm���ŷ�ҩ����", "��ʾ��ҩ��������", 1)) = 1)
        
        .intҩƷ���Ʊ�����ʾ = GetDrugFormat
        
        '��ѯ��ʾ
        .int��ѯ��ҩ���� = Val(zlDatabase.GetPara("��ѯ��ҩ����", glngSys, 1342, 7))
        .int��ѯ��ҩ���� = Val(zlDatabase.GetPara("��ѯ��ҩ����", glngSys, 1342, 3))
        .lng����¼�� = Val(zlDatabase.GetPara("��ѯ��ϸ��¼��", glngSys, 1342, 3000))
        
        '�����
        .IntCheckStock = MediWork_GetCheckStockRule(.lngҩ��ID)
    
        '�ⷿ��λ
        .strUnit = GetSpecUnit(.lngҩ��ID, gintסԺҩ��)
        
        '�Ƿ�����PASS
        .blnStarPass = gintPass <> 0 And mPrives.bln������ҩ��� = True
        
        '��������
        .bln�������� = CheckIsCenter(.lngҩ��ID)
        
        '�������ã���Դ����
        .strSourceDep = zlDatabase.GetPara("��Դ����", glngSys, 1342)
        
        'ע������
        .intFont = Val(zlDatabase.GetPara("����", glngSys, 1342))
        .StrFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
        
        'ע����������װ�����
        .int��ͣ���� = Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��ͣ����", "0"))
        .int��ͣ���� = IIf(.int��ͣ���� = 1, 1, 0)
        .str��װ������ = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��������", "11")
        .str��װ������ = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "ѡ�����", "����")
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshReturnDept()
    'ˢ����ҩ�����б�
    Dim rsData As ADODB.Recordset
    Dim strDanger As String
    Dim strToxicology As String
    
    '''select
    gstrSQL = "Select" & IIf(mParams.strSourceDep = "", "", "/*+rule*/") & "  H.ID, H.���� As ��������, S.���ܷ�ҩ�� As ��ҩ��, Decode(Nvl(c.Ӥ����,0), 0, Nvl(b.����, c.����), z.Ӥ������) ����, B.����ID, Decode(Nvl(c.Ӥ����,0), 0, Nvl(p.�Ա�, c.�Ա�), z.Ӥ���Ա�) �Ա�, Decode(Nvl(c.Ӥ����,0), 0, p.����, Ceil(Sysdate - z.����ʱ��) || '��') ����, S.����, S.NO, S.ҩƷid, " & _
        " Decode(Nvl(C.ҽ�����, 0), 0, 0, 1) ҽ�����, C.�����־, Nvl(S.����, 0) ����, S.ID As �շ�id, S.��������, Nvl(B.��ǰ����,'') As ����,W.��ɫ,c.Ӥ���� "

    
    '''from
    gstrSQL = gstrSQL & " From ҩƷ�շ���¼ S, סԺ���ü�¼ C, ������Ϣ B, ҩƷ��� D, ҩƷ���� T, ������ҳ P, ���ű� H,�������� W, ������������¼ Z " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([17]) As zlTools.t_NumList)) E ")
    
    '''where
    gstrSQL = gstrSQL & " Where S.�Է�����id = H.ID" & IIf(mParams.strSourceDep = "", "", " And S.�Է�����id=E.Column_Value ") & _
        " And C.����id = B.����id And C.����id=P.����id And C.��ҳid=P.��ҳid And C.NO = S.NO And S.����id = C.ID And c.����id = z.����id(+) And c.Ӥ���� = z.���(+) And C.��ҳid=Z.��ҳid(+) " & _
        " And S.�ⷿid = C.ִ�в���id And S.ҩƷid = D.ҩƷid And D.ҩ��id = T.ҩ��id And P.��������=W.����(+) " & _
        " And (H.����ʱ�� Is Null Or H.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
        " And S.������� Between [2] And [3] And S.����� IS NOT NULL "
    
    'վ�����
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (H.վ�� = [16] Or H.վ�� Is Null) "
    End If
    
    '��ǰҩ��
    gstrSQL = gstrSQL & " And S.�ⷿid + 0 = [1] "
    
    '¼����Ϣ
    If mcondition.strסԺ�� <> "" Then
        gstrSQL = gstrSQL & " And P.סԺ�� = [4] "
    ElseIf mcondition.str���� <> "" Then
        '���ڴ��Ų�Ψһ��תΪͨ������ID����ѯ
        gstrSQL = gstrSQL & " And B.����ID+0 = [9] "
    ElseIf mcondition.str���￨ <> "" Then
        gstrSQL = gstrSQL & " And B.���￨�� = [6] "
    ElseIf mcondition.str���� <> "" Then
        gstrSQL = gstrSQL & " And P.���� = [7] "
    ElseIf mcondition.strNo <> "" Then
        gstrSQL = gstrSQL & " And S.NO = [8] "
    ElseIf mcondition.lng����ID <> -1 Then
        gstrSQL = gstrSQL & " And B.����ID+0 = [9] "
    ElseIf mcondition.cur��ҩ�� <> 0 Then
        gstrSQL = gstrSQL & " And S.���ܷ�ҩ�� = [10] "
    ElseIf mcondition.lng��ҩ����ID <> -1 Then
        gstrSQL = gstrSQL & " And S.�Է�����id + 0 = [11] "
    End If
    
    '����ģʽ:0-����,1-���ʵ�,2-���ʱ�
    If mcondition.int����ģʽ = 0 Then
        gstrSQL = gstrSQL & " And S.���� IN(9,10)"
    ElseIf mcondition.int����ģʽ = 1 Then
        gstrSQL = gstrSQL & " And S.����=9"
    ElseIf mcondition.int����ģʽ = 2 Then
        gstrSQL = gstrSQL & " And S.����=10"
    End If
    
    '������
    If mcondition.str������ <> "���м�����" Then
        gstrSQL = gstrSQL & " And S.������ = [12] "
    End If
    
    'ҽ������:0-����,1-����,2-����,3-��ͨ
    '�õ����Ƿ���д�����Ƿ�ҽ��������ҩƷ����
    If mcondition.intҽ������ = 0 Then
    ElseIf mcondition.intҽ������ = 1 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 2 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf mcondition.intҽ������ = 3 Then
        gstrSQL = gstrSQL & " And (Nvl(C.ҽ�����,0) + 0 =0 Or S.���� Is Null) "
    ElseIf mcondition.intҽ������ = 4 Then
        gstrSQL = gstrSQL & " And S.���� Is Not Null And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_') And Nvl(C.ҽ�����,0) + 0 > 0 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If mcondition.int��ҩ���� = 0 Then
    ElseIf mcondition.int��ҩ���� = 1 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 2 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf mcondition.int��ҩ���� = 3 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 4 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 5 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf mcondition.int��ҩ���� = 6 Then
        gstrSQL = gstrSQL & " And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4')"
    End If
    
    '�������ͣ����˻�Ӥ��
    If mcondition.int�������� = 0 Then
        gstrSQL = gstrSQL & " And Nvl(C.Ӥ����, 0) = 0 "
    ElseIf mcondition.int�������� = 1 Then
        gstrSQL = gstrSQL & " And Nvl(C.Ӥ����, 0) > 0 "
    End If
    
    '��ҩ;��
    If mcondition.str��ҩ;�� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [13] || ',',',' || S.�÷� || ',') > 0 "
    End If
    
    'ҩƷ����
    If mcondition.strҩƷ���� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [14] || ',',',' || T.ҩƷ���� || ',') > 0 "
    End If
    
    '������ҩ����
    If mcondition.str������ҩ���� <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [15] || ',',',' || D.��ҩ���� || ',') > 0 "
    End If
    
    '��������
    If Trim(txtInput.Text) = "" Then
        If mParams.intShowDept = 1 Then
            gstrSQL = gstrSQL & " And H.id In (Select ����id From ��������˵�� Where �������� = '�ٴ�' And ������� In (2, 3)) "
        ElseIf mParams.intShowDept = 2 Then
            gstrSQL = gstrSQL & " And H.id In (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����','Ӫ��') And ������� IN(2,3)) "
        ElseIf mParams.intShowDept = 3 Then
            gstrSQL = gstrSQL & " And H.id In (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3)) "
        End If
    End If
    
    '�ų�������Һ�������Ĺ����в����ĵ���
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where y.�շ�id = s.Id) "
    
    '��ΣҩƷ
    If chkDanger.Value = 1 Then
        If chkDangerType(0).Value = 1 Then strDanger = IIf(strDanger = "", 1, strDanger & "," & 1)
        If chkDangerType(1).Value = 1 Then strDanger = IIf(strDanger = "", 2, strDanger & "," & 2)
        If chkDangerType(2).Value = 1 Then strDanger = IIf(strDanger = "", 3, strDanger & "," & 3)
    End If
    If strDanger <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [18] || ',' , ',' || Nvl(D.��ΣҩƷ,0) || ',') > 0 "
    
    '�������
    If Me.chkToxicologyType.Value = 1 Then
        If Me.chkToxicology(0).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(0).Caption, strToxicology & "," & Me.chkToxicology(0).Caption)
        If Me.chkToxicology(1).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(1).Caption, strToxicology & "," & Me.chkToxicology(1).Caption)
        If Me.chkToxicology(2).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(2).Caption, strToxicology & "," & Me.chkToxicology(2).Caption)
        If Me.chkToxicology(3).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(3).Caption, strToxicology & "," & Me.chkToxicology(3).Caption)
    End If
    
    If strToxicology <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [19] || ',' , ',' || T.������� || ',') > 0 "
    
    '''order by
    gstrSQL = gstrSQL & " Order By H.����, ��ҩ��, B.����, S.NO "
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "��ȡ����ҩ���һ���", _
        mcondition.lngҩ��ID, _
        CDate(mcondition.str��ʼʱ��), _
        CDate(mcondition.str����ʱ��), _
        mcondition.strסԺ��, _
        mcondition.str����, _
        mcondition.str���￨, _
        mcondition.str����, _
        mcondition.strNo, _
        mcondition.lng����ID, _
        mcondition.cur��ҩ��, _
        mcondition.lng��ҩ����ID, _
        mcondition.str������, _
        mcondition.str��ҩ;��, _
        mcondition.strҩƷ����, _
        mcondition.str������ҩ����, _
        mstrDeptNode, _
        mParams.strSourceDep, _
        strDanger, _
        strToxicology)
    
    '���²�������
    Call GetReturnDeptTreeView(rsData)
    
    '���²��������Ӧ���շ���¼���ݼ�
    Call GetDeptListRecord(rsData)
    
    Me.MousePointer = 0
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckAdvice(ByVal rsData As ADODB.Recordset) As Boolean
    '�ȼ���Ƿ�������ҩ��ҽ����
    Dim rsTmp As ADODB.Recordset
    
    CheckAdvice = False
    On Error GoTo errHandle
    If mParams.bln����ҽ�������Ϻ���ҩ = True Then
        CheckAdvice = True
        Exit Function
    End If
    
    With rsData
        .Filter = "ִ��״̬=" & mState.��ҩ
        
        Do While Not .EOF
            gstrSQL = "select ���� From ҩƷ�շ���¼ Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ�������]", CLng(!�շ�Id))
            
            If (rsTmp!���� Like "1*") Then       '����
                gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From סԺ���ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", CLng(!�շ�Id))
                
                If Not rsTmp.EOF Then
                    If (rsTmp!�����־ = 1 Or rsTmp!�����־ = 4) And rsTmp!ҽ����� <> 0 Then
                        gstrSQL = "Select decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rsTmp!ҽ�����))
                        
                        If rsTmp!���� = 0 Then
                            MsgBox "[" & " & !NO & " & "]�е�ҩƷ[" & !Ʒ�� & "]��Ӧ��ҽ����δ���ϣ�������ҩ��", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    CheckAdvice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitComandBars()
    '��ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim IntCount As Integer
    Dim strCardName As String
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = frmPublic.imgPublic.Icons
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsMain.ActiveMenuBar.Title = "�˵�"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.Id = mconMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ(&P)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Excel, "�����&Excel��")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrint, "���ݴ�ӡ(&B)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrintTotal, "��ӡ�����嵥(&C)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrintRestore, "��ӡ��ҩ֪ͨ��(&R)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrintWait, "��ӡҩƷ��ҩ��(&W)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Parameter, "��������(&T)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Verify, "��ҩ(&V)")
        cbrControlMain.Visible = mPrives.bln��ҩ
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Reject, "�ܷ�ȷ��(&H)")
        cbrControlMain.Visible = mPrives.bln�ܷ�
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, "�ܷ��ָ�(&H)")
        cbrControlMain.Visible = mPrives.bln�ܷ�
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Return, "��ҩ(&R)")
        cbrControlMain.Visible = mPrives.bln��ҩ
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_EMR, "������ѯ(&Z)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_ReturnOther, "������ҩ���Ĵ���(&T)")
        cbrControlMain.Visible = mPrives.bln������ҩ���Ĵ���
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_VerifySign, "��֤ǩ��(&S)")
        cbrControlMain.Visible = gblnESign���ŷ�ҩ
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_ReVerify, "ҩƷ��ҩ����(&B)")
        cbrControlMain.Visible = mPrives.bln��ҩ����
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_StopFlag, "ֹͣ��ҩ���(&S)")
        cbrControlMain.Visible = (mPrives.blnֹͣ��ҩ = True Or mPrives.bln�ָ���ҩ = True)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Packer, "�ְ����ӿ�����(&P)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Hot_IC, "��IC��(&I)")
        cbrControlMain.Visible = False
    End With
    
'    '�Զ�����ҩ���ò˵�
'    If Not gobjPackerMZ Is Nothing Then
'        Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_AutoSend, "ҩ���Զ�������(&V)", -1, False)
'        cbrMenuBar.Id = mconMenu_AutoSend
'    End If
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.Id = mconMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_ToolBar, "������(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrControl.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_StatusBar, "״̬��(&S)")
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_FontSize, "����(&F)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_1, "С����(&S)", -1, False)
        If mParams.intFont = 0 Then cbrControl.Checked = True
        cbrControl.Parameter = 0
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_2, "������(&M)", -1, False)
        If mParams.intFont = 1 Then cbrControl.Checked = True
        cbrControl.Parameter = 1
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_3, "������(&B)", -1, False)
        If mParams.intFont = 2 Then cbrControl.Checked = True
        cbrControl.Parameter = 2
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Find, "����(&L)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_FindNext, "������һ��(&N)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_SelAll, "ȫѡ(&A)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_ClsAll, "ȫ��(&C)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��(&R)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.Id = mconMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "��������(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB�ϵ�����")
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, "������ҳ(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "������̳(&F)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_About, "����(&A)��")
        cbrControlMain.BeginGroup = True
    End With
    
    '���˵��Ҳ����Ϣ��ʾ�����˵������յ���Ϣʱ��̬����
    With cbsMain.ActiveMenuBar.Controls
        Set cbrMenuBar = .Add(xtpControlPopup, mconMenu_File_Message, "����Ϣ����")
        cbrMenuBar.Id = mconMenu_File_Message
        cbrMenuBar.flags = xtpFlagRightAlign
        cbrMenuBar.Visible = mPrives.bln��ҩ����
    End With
        
    '�����
    With Me.cbsMain.KeyBindings
'        .Add FCONTROL, Asc("S"), mconMenu_Edit_Save
'        .Add FCONTROL, Asc("Z"), mconMenu_Edit_Untread
'        .Add FCONTROL, Asc("M"), mconMenu_Edit_Modify
'        .Add FSHIFT, VK_DELETE, mconMenu_Edit_Delete
        .Add FCONTROL, VK_F4, mconMenu_Edit_Dept_Hot_IC
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
    End With

    '���ò����ò˵�
    With Me.cbsMain.Options
        .AddHiddenCommand mconMenu_File_PrintSet
        .AddHiddenCommand mconMenu_File_Excel
    End With
    
    '����¼����Ŀ�����˵�
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_InputPopup, "¼��(&I)", -1, False)
    cbrMenuBar.Id = mconMenu_InputPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_HosNumber, "סԺ��(&0)")
        cbrControlMain.Parameter = "ס|סԺ��|0|||||"
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_Name, "����(&1)")
        cbrControlMain.Parameter = "��|����|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_BedNumber, "����(&2)")
        cbrControlMain.Parameter = "��|����|0||||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_NO, "���ݺ�(&3)")
        cbrControlMain.Parameter = "��|���ݺ�|0||||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_Ident, "����ID(&4)")
        cbrControlMain.Parameter = "��|����ID|0||||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_ReceiveNO, "��ҩ��(&5)")
        cbrControlMain.Parameter = "��|��ҩ��|0||||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_BatchSendNO, "���ܷ�ҩ��(&6)")
        cbrControlMain.Parameter = "��|��ҩ��|0||||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_Dept, "��ҩ����(&7)")
        cbrControlMain.Parameter = "��|��ҩ����|0||||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_ICCard, "IC��(&8)")
        cbrControlMain.Parameter = "IC|IC����|1|||||"
        
        '��̬ȡ����ҽ�ƿ�����Ҫ�����п���
        If mstrCardType <> "" Then
            mintCardCount = UBound(Split(mstrCardType, ";")) + 1
            For IntCount = 0 To UBound(Split(mstrCardType, ";"))
                'ȡ���п�����
                strCardName = Split(Split(mstrCardType, ";")(IntCount), "|")(1)
                
                'ȥ���ظ����ص�����
                If InStr(";סԺ��;����;����;���ݺ�;����ID;��ҩ��;���ܷ�ҩ��;��ҩ����;IC��;", ";" & strCardName & ";") = 0 Then
                    Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Dept_ICCard + IntCount + 1, strCardName & "(&" & IntCount + 9 & ")")
                
                    '���濨��Ϣ
                    cbrControlMain.Parameter = Split(mstrCardType, ";")(IntCount)
                End If
                
                If IntCount = 0 Then
                    cbrControlMain.BeginGroup = True
                End If
            Next
        End If
    End With
    
    '���ò����б���Ŀ�����˵�
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ListPopup, "��Ŀ(&I)", -1, False)
    cbrMenuBar.Id = mconMenu_ListPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowReject, "�����ܷ�ҩƷ(&R)")
        cbrControlMain.Checked = mParams.blnShowReject
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_OnlyShowDept, "����ʾ����(&0)")
        cbrControlMain.Checked = mParams.blnOnlyShowDept
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowOther, "��ʾ��ϸ��Ϣ(&1)")
        cbrControlMain.Checked = Not mParams.blnOnlyShowDept
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowAll, "��ʾ���п���(&A)")
        cbrControlMain.Checked = (mParams.intShowDept = 0)
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowClin, "��ʾ�ٴ�����(&C)")
        cbrControlMain.Checked = (mParams.intShowDept = 1)
    
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowTech, "��ʾҽ������(&T)")
        cbrControlMain.Checked = (mParams.intShowDept = 2)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowArea, "��ʾ���˲���(&B)")
        cbrControlMain.Checked = (mParams.intShowDept = 3)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_Sort, "���Ұ�ҽ������ʱ������(&D)")
        cbrControlMain.Checked = mParams.blnSort
        cbrControlMain.BeginGroup = True
    End With
    
    '���ø�ҩ;�����൯���˵�
    Set rsData = DeptSendWork_Get��ҩ;������
    
    If rsData.RecordCount > 0 Then
        Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_TypePopup, "����(&T)", -1, False)
        cbrMenuBar.Id = mconMenu_TypePopup
        cbrMenuBar.Visible = False
        
        mTypeCount = rsData.RecordCount
        With cbrMenuBar.CommandBar.Controls
            For i = 1 To rsData.RecordCount
                Set cbrControlMain = .Add(xtpControlButton, mconMenu_TypePopup + i, rsData!����)
                rsData.MoveNext
            Next
        End With
    End If
    
    '���ò����б��в�������˵�
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_SortPopup, "��������(&P)", -1, False)
    cbrMenuBar.Id = mconMenu_SortPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByName, "����������(&0)")
        cbrControlMain.Checked = (mParams.int�������� = 1)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByBedNo, "����λ����(&1)")
        cbrControlMain.Checked = (mParams.int�������� = 2)
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Verify, "��ҩ")
        cbrControlMain.Visible = mPrives.bln��ҩ
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Reject, "�ܷ�")
        cbrControlMain.Visible = mPrives.bln�ܷ�
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, "�ָ�")
        cbrControlMain.Visible = mPrives.bln�ܷ�
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Return, "��ҩ")
        cbrControlMain.Visible = mPrives.bln��ҩ
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_VerifySign, "��֤ǩ��")
        cbrControlMain.Visible = gblnESign���ŷ�ҩ

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_ReVerify, "����")
        cbrControlMain.Visible = mPrives.bln��ҩ����
        cbrControlMain.BeginGroup = True
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_EMR, "������ѯ")
'        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��")
        cbrControlMain.BeginGroup = True

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
        cbrControlMain.BeginGroup = True
    End With
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
End Sub


Private Sub InitPanes()
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane
    Dim objPaneList As Pane
    Dim objPaneDetail As Pane
    
    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_Dept_Condition, 225, 100, DockLeftOf, Nothing)
    objPaneCon.Title = "��������"
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
'    objPaneCon.MaxTrackSize.SetSize 290, 500
    
'    Set objPaneList = Me.dkpMain.CreatePane(mconPane_SelDept, 290, 250, DockBottomOf, objPaneCon)
'    objPaneList.Title = "��������"
'    objPaneList.Options = PaneNoCloseable Or PaneNoFloatable
End Sub
Private Sub InitTabControl()
    '��ʼ����ҳ�ؼ�
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(0, "δ��ҩƷ�嵥(&0)", mfrmDetail.hWnd, 0).Tag = "δ��ҩƷ�嵥_"
        .InsertItem(1, "�����嵥(&1)", mfrmDetail.hWnd, 0).Tag = "�����嵥_"
        .InsertItem(2, "ȱҩ�嵥(&2)", mfrmDetail.hWnd, 0).Tag = "ȱҩ�嵥_"
        .InsertItem(3, "�ܷ�ҩ�嵥(&3)", mfrmDetail.hWnd, 0).Tag = "�ܷ�ҩ�嵥_"
        .InsertItem(4, "�ѷ�ҩ�嵥(&4)", mfrmDetail.hWnd, 0).Tag = "�ѷ�ҩ�嵥_"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
End Sub


Private Sub Load��ҩ;��()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get��ҩ;��
    
    With Lvw��ҩ;��
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "���и�ҩ;��", 1, 1
        .ListItems(.ListItems.count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, rsData!�÷�, 1, 1
            .ListItems(.ListItems.count).Checked = True
            .ListItems(.ListItems.count).Tag = rsData!����
            rsData.MoveNext
        Loop
    End With
End Sub
Private Function Load��ҩҩ��() As Boolean
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    Dim intIndex As Integer
    
    Set rsData = DeptSendWork_GetDrugstore(mstrPrivs, glngUserId, gstrNodeNo)
    
    If rsData.EOF Then
        If IsInString(mstrPrivs, "����ҩ��", ";") Then
            strMsg = "���ʼ��ҩ�������Ź���"
        Else
            strMsg = "�㲻��ҩ��������Ա�����ܲ�����ģ�飡"
        End If
        
        MsgBox strMsg, vbInformation, gstrSysName
        Load��ҩҩ�� = False
        Exit Function
    Else
        
        rsData.Filter = "id=" & mParams.lngҩ��ID
        If rsData.EOF Then
            Call ResetParams(True)
        End If
        
        rsData.Filter = ""
        With cbo��ҩҩ��
            .Clear
            
            Do While Not rsData.EOF
                .AddItem rsData!����
                .ItemData(.NewIndex) = rsData!Id
                
                If rsData!Id = mParams.lngҩ��ID Then intIndex = .NewIndex
                
                rsData.MoveNext
            Loop
            
            .ListIndex = intIndex
            
            .Tag = .ItemData(intIndex)
        End With
        
        Load��ҩҩ�� = True
    End If
End Function

Private Sub LoadҩƷ����(ByVal lngҩ��ID As Long)
    Dim rsData As ADODB.Recordset
    Dim bln��ҩ�ⷿ As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ��������˵�� " & _
         " Where �������� Like '��ҩ%' And ����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鲿������]", Val(cbo��ҩҩ��.ItemData(cbo��ҩҩ��.ListIndex)))
    
    If Not rsData.EOF Then bln��ҩ�ⷿ = True
    
    Set rsData = DeptSendWork_Get����(lngҩ��ID)
    
    With LvwҩƷ����
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "����ҩƷ����", 1, 1
        .ListItems(.ListItems.count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, rsData!����, 1, 1
            .ListItems(.ListItems.count).Checked = True
            rsData.MoveNext
        Loop
        If bln��ҩ�ⷿ Then
           .ListItems.Add , "_" & .ListItems.count + 1, "0-����", 1, 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Loadҽ������()
    '����ҽ������
    With Cboҽ������
        .Clear
        .AddItem "0-�������е���"
        .AddItem "1-��������ҽ��"
        .AddItem "2-������ʱҽ��"
        .AddItem "3-��ͨ���ʵ���"
        .AddItem "4-��������ҽ��"
'        .ListIndex = Lngҽ������
    End With
End Sub
Private Sub ResetParams(Optional ByVal blnNext As Boolean)
    Dim intFixedCol As Integer
    Dim dateCurDate As Date
    Dim i As Integer
    
    BlnSetPara = False
    With Frm���ŷ�ҩ��������
        .strPrivs = mstrPrivs
        .blnStartPacker = (TypeName(mobjDrugMAC) = "clsDrugPacker" And mblnStartPacker)
        .Show 1, Me
    End With
    
    If BlnSetPara Then
        '����ȡ����
        Call GetParams
        If blnNext = True Then Exit Sub
        '����ҩ��
        If Val(cbo��ҩҩ��.Tag) <> mParams.lngҩ��ID Then
            For i = 0 To cbo��ҩҩ��.ListCount - 1
                If Val(cbo��ҩҩ��.ItemData(i)) = mParams.lngҩ��ID Then
                    cbo��ҩҩ��.Tag = cbo��ҩҩ��.ItemData(i)
                    cbo��ҩҩ��.ListIndex = i
                    Exit For
                End If
            Next
            
            ClearDetailList IIf(tbcDetail.Selected.Index = 0, mListType.��ҩ, mListType.��ҩ)
            
            mstrDeptNode = GetDeptStationNode(mParams.lngҩ��ID)
        End If
        
        mfrmDetail.Load�˲��� (mParams.lngҩ��ID)
        
        Call Loadʱ�䷶Χ
        
        Call SetPacker
        
        '�����Ӵ��ڵĲ���
        mfrmDetail.SetParams
        
        '����ˢ����ϸ
        Call cmdRefreshDept_Click
        Call cmdRefresh_Click
    End If
End Sub

Private Sub SetListItemCheck(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    '�б���ʾ��ʽ��������ʾ��ʽ���Ƿ�����ܷ�
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_ListPopup)
    If Not cbrMenuBar Is Nothing Then
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            If cbrControl.Id = mconMenu_List_ShowReject And Control.Id = mconMenu_List_ShowReject Then
                cbrControl.Checked = Not cbrControl.Checked
                mParams.blnShowReject = cbrControl.Checked
            ElseIf (cbrControl.Id > mconMenu_ListPopup And cbrControl.Id <= mconMenu_List_ShowOther) _
                And (Control.Id > mconMenu_ListPopup And Control.Id <= mconMenu_List_ShowOther) Then
                cbrControl.Checked = (cbrControl.Id = Control.Id)
                If cbrControl.Id = mconMenu_List_OnlyShowDept Then
                    mParams.blnOnlyShowDept = cbrControl.Checked
                End If
            ElseIf (cbrControl.Id >= mconMenu_List_ShowAll And cbrControl.Id <= mconMenu_List_ShowArea) _
                And (Control.Id >= mconMenu_List_ShowAll And Control.Id <= mconMenu_List_ShowArea) Then
                cbrControl.Checked = (cbrControl.Id = Control.Id)
                mParams.intShowDept = Control.Id - mconMenu_List_ShowAll
            ElseIf cbrControl.Id = mconMenu_List_Sort And Control.Id = mconMenu_List_Sort Then
                cbrControl.Checked = Not cbrControl.Checked
                mParams.blnSort = cbrControl.Checked
            End If
        Next
    End If
End Sub

Private Sub SetPacker()
'    If gtype_UserSysParms.P222_ҩ���Զ�����ҩ�ӿ� = 0 Then
'        'δ����ʱ����ʾ��װ��ͼ��
'        Me.stbThis.Panels(5).Visible = False
'    Else
'        '��������״̬��ʾ��ͬ�ķ�ҩ��ͼ��
'        If gtype_UserSysParms.P222_ҩ���Զ�����ҩ�ӿ� = 1 And gobjPackerZY Is Nothing Then
'            '��������ʱ
'            Me.stbThis.Panels(5).Picture = imgPacker.ListImages(1).Picture
'            Me.stbThis.Panels(5).Enabled = True
'        Else
'            'δ����״̬
'            Me.stbThis.Panels(5).Picture = imgPacker.ListImages(2).Picture
'            Me.stbThis.Panels(5).Enabled = False
'        End If
'    End If
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Packer, , True)

    If mblnStartPacker = False Then
        If Not cbrMenu Is Nothing Then
            cbrMenu.Visible = False
        End If
        
        'δ����ʱ����ʾ��װ��ͼ��
        Me.stbThis.Panels(5).Visible = False
    Else
        If Not cbrMenu Is Nothing Then
            cbrMenu.Visible = True
        End If
        
        '��������״̬��ʾ��ͬ�İ�װ��ͼ��
        If mblnPackerConnect = True Then
            If mParams.int��ͣ���� = 0 Then
                '��������ʱ
                Me.stbThis.Panels(5).Picture = imgPacker.ListImages(1).Picture
            Else
                '��ͣ����ʱ
                Me.stbThis.Panels(5).Picture = imgPacker.ListImages(3).Picture
            End If
            
            Me.stbThis.Panels(5).Enabled = True
        Else
            'δ����״̬
            Me.stbThis.Panels(5).Picture = imgPacker.ListImages(2).Picture
            Me.stbThis.Panels(5).Enabled = False
        End If
    End If
End Sub

Private Sub ShowOtherConditon()
    picShowOther.Tag = Abs(Val(picShowOther.Tag) - 1)
    picUpOrDown.Picture = imgLvwSel.ListImages(Val(picShowOther.Tag) + 3).Picture
    Call picCondition_Resize
End Sub

Private Sub ShowWindow_ReturnOther()
    TimerAuto.Enabled = False
    
    Set frm������ҩ.In_PlugIn = mobjPlugIn
    If Not frm������ҩ.ShowEditor(Me, mcondition.lngҩ��ID, False, mParams.int����λ��, mstrPrivs) Then
        TimerAuto.Enabled = True
        Exit Sub
    End If
    
    DoEvents
    
    TimerAuto.Enabled = True
End Sub

Private Sub ShowWindow_ReVerify(ByVal strWriteOffMsg As String)
    Dim strWriteOffInfo As String   '������˽��淵�ص��ϴβ���������˹�����Ϣ������ʱ��,����id|����ʱ��,����id...
    
    TimerAuto.Enabled = False
    
    BlnRefresh = False
    
    strWriteOffInfo = FrmҩƷ����.ShowForm(Me, mcondition.lngҩ��ID, mParams.strUnit, _
        mParams.int����λ��, mstrCardType, mParams.int��ҩ�嵥��ӡ, strWriteOffMsg, _
        mobjSquareCard, mobjPlugIn)
    
    If BlnRefresh Then
        'ɾ����Ϣ��¼�����Ѿ���˹�����Ϣ��¼
        If strWriteOffInfo <> "" Then
            If Not mrsReceiveMsg Is Nothing Then
                If mrsReceiveMsg.RecordCount > 0 Then
                    With mrsReceiveMsg
                        .MoveFirst
                        Do While Not .EOF
                            If InStr(strWriteOffInfo & "|", Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & !����ID & "|") > 0 Then
                                .Delete adAffectCurrent
                            End If
                            
                            .MoveNext
                        Loop
                    End With
                    '������Ϣ�˵�
                    Call SetMessageBar(mrsReceiveMsg)
                End If
            End If
        End If
        
        cmdRefresh_Click
    End If
    
    DoEvents
    
    TimerAuto.Enabled = True
End Sub

Private Sub ShowWindow_StopFlag()
    Dim frmFlag As New Frm���ٷ�ҩ������־
    
    TimerAuto.Enabled = False
    BlnRefresh = True
    
    frmFlag.In_����� = mParams.IntCheckStock
    frmFlag.gstrParentName = "Frm���ŷ�ҩ����New"
    frmFlag.ShowMe Me, Val(cbo��ҩҩ��.ItemData(cbo��ҩҩ��.ListIndex))
    
    If BlnRefresh Then
        cmdRefresh_Click
    End If
    
    DoEvents
    TimerAuto.Enabled = True
End Sub



Private Sub UpdateDeptListRecord(ByVal intType As Integer)
    '���ݲ�������Ĺ�ѡ����������ݼ�
    Dim i As Integer
    
    If Val(tvwList(intType).Tag) = 0 Then Exit Sub
    
    If mrsDeptList Is Nothing Then Exit Sub
    If mrsDeptList.State = 0 Then Exit Sub
    
    With tvwList(intType)
        If mParams.blnOnlyShowDept = False Then
            '�������ʾ��ϸ��Ϣ����ʵ�ʹ�ѡ��NO�ű��
            For i = 1 To .Nodes.count
                If Mid(.Nodes(i).Key, 1, 1) = "N" Then
                    mrsDeptList.Filter = "NO='" & Split(.Nodes(i).Tag, "|")(0) & "' and ����id=" & Val(Split(.Nodes(i).Tag, "|")(1)) & " and ��������='" & Split(.Nodes(i).Tag, "|")(2) & "'"
                    Do While Not mrsDeptList.EOF
                        mrsDeptList!ִ��״̬ = IIf(.Nodes(i).Checked = True, 1, 0)
                        mrsDeptList.Update
                        
                        mrsDeptList.MoveNext
                    Loop
                End If
            Next
        Else
            '�������ʾ�������ƣ���ʵ�ʹ�ѡ�Ĳ��ű��
            For i = 1 To .Nodes.count
                If Mid(.Nodes(i).Key, 1, 1) = "D" Then
                    mrsDeptList.Filter = "����ID=" & Mid(.Nodes(i).Key, 3) & ""
                    Do While Not mrsDeptList.EOF
                        mrsDeptList!ִ��״̬ = IIf(.Nodes(i).Checked = True, 1, 0)
                        mrsDeptList.Update
                        
                        mrsDeptList.MoveNext
                    Loop
                End If
            Next
        End If
    End With
    
    mrsDeptList.Filter = ""
End Sub

Private Function WarRecoredCount(ByVal lngCount As Long) As Boolean
    Dim intProc As Integer
    
    If mFindWar.blnNoAsk_Rec = True Then
        WarRecoredCount = mFindWar.blnProc_Rec
        Exit Function
    End If
    
    intProc = vbYes
    
    '��ѯ��¼������ʱ����
    If mFindWar.blnNoAsk_Rec = False Then
        If lngCount > mParams.lng����¼�� Then
            intProc = frmMsgBox.ShowMsgBox("��ѯ������Ҫ�ܳ�ʱ�䣬�Ƿ������", Me)
            mFindWar.blnNoAsk_Rec = (intProc = vbIgnore Or intProc = vbCancel)
            mFindWar.blnProc_Rec = (intProc = vbYes Or intProc = vbIgnore)
        End If
    End If
    
    WarRecoredCount = mFindWar.blnProc_Rec
End Function

Private Function WarTimeArea() As Boolean
    Dim intDateDiff As Integer
    Dim intProc As Integer
    
    '��ѯʱ����
    intDateDiff = DateDiff("d", CDate(mcondition.str��ʼʱ��), CDate(mcondition.str����ʱ��))
    
    'С�ڲ�ѯʱ�����������������
    If tbcDetail.Selected.Index = mListType.��ҩ Then
        If intDateDiff <= mParams.int��ѯ��ҩ���� Then
            WarTimeArea = True
            Exit Function
        End If
    Else
        If intDateDiff <= mParams.int��ѯ��ҩ���� Then
            WarTimeArea = True
            Exit Function
        End If
    End If
    
    '���ڲ�ѯʱ��ʱ������ϴ�ѡ����ǲ�����ʾ�����ϴ�ѡ���������
    If tbcDetail.Selected.Index = mListType.��ҩ Then
        If mFindWar.blnNoAsk_Dept_Sended = True Then
            WarTimeArea = mFindWar.blnProc_Dept_Sended
            Exit Function
        End If
    Else
        If mFindWar.blnNoAsk_Dept_Send = True Then
            WarTimeArea = mFindWar.blnProc_Dept_Send
            Exit Function
        End If
    End If
    
    '��ʾ��ʾ
    If tbcDetail.Selected.Index = mListType.��ҩ Then
        If intDateDiff > mParams.int��ѯ��ҩ���� Then
            intProc = frmMsgBox.ShowMsgBox("��ѯ������Ҫ�ܳ�ʱ�䣬�Ƿ������", Me)
                
            mFindWar.blnNoAsk_Dept_Sended = (intProc = vbIgnore Or intProc = vbCancel)
            mFindWar.blnProc_Dept_Sended = (intProc = vbYes Or intProc = vbIgnore)
        End If
        WarTimeArea = mFindWar.blnProc_Dept_Sended
    Else
        If intDateDiff > mParams.int��ѯ��ҩ���� Then
            intProc = frmMsgBox.ShowMsgBox("��ѯ������Ҫ�ܳ�ʱ�䣬�Ƿ������", Me)
                
            mFindWar.blnNoAsk_Dept_Send = (intProc = vbIgnore Or intProc = vbCancel)
            mFindWar.blnProc_Dept_Send = (intProc = vbYes Or intProc = vbIgnore)
        End If
        WarTimeArea = mFindWar.blnProc_Dept_Send
    End If
End Function

Private Sub Cbo��ҩҩ��_Click()
    If mblnStart = False Then Exit Sub
    
    With cbo��ҩҩ��
        If Val(.Tag) <> Val(.ItemData(.ListIndex)) Then
            Call LoadҩƷ����(Val(.ItemData(.ListIndex)))
            .Tag = Val(.ItemData(.ListIndex))
            
            mcondition.lngҩ��ID = Val(.Tag)
            
            mfrmDetail.Load�˲��� (mcondition.lngҩ��ID)
            
            If Not gobjESign Is Nothing Then
                gblnESign���ŷ�ҩ = EsignIsOpen(mcondition.lngҩ��ID)
            End If
            
            mstrDeptNode = GetDeptStationNode(Val(.Tag))
            
            zlDatabase.SetPara "��ҩҩ��", mcondition.lngҩ��ID, glngSys, 1342
            mblnIs�������� = Is��������(Val(.Tag))
            
            '�����
            mParams.IntCheckStock = MediWork_GetCheckStockRule(Val(.Tag))
                
            '�ⷿ��λ
            mParams.strUnit = GetSpecUnit(Val(.Tag), gintסԺҩ��)
            
            '�����Ӵ��ڵĲ���
            mfrmDetail.SetParams
            
            '����б�
            ClearTreeView IIf(tbcDetail.Selected.Index = mListType.��ҩ, 1, 0)
            
            Select Case tbcDetail.Selected.Index
                Case mListType.��ҩ, mListType.����, mListType.�ܷ�
                    ClearDetailList mListType.��ҩ
                Case mListType.��ҩ
                    ClearDetailList mListType.��ҩ
            End Select
            
            Call SetCommandBar(tbcDetail.Selected.Index)
        End If
    End With
End Sub

Private Function Is��������(ByVal lngҩ��ID As Long)
    'Is��������
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    strTmp = "select ����id from ��������˵�� where ����id=[1] and ��������='��������'"
    Set rsSQL = zlDatabase.OpenSQLRecord(strTmp, "Is��������", lngҩ��ID)
    Is�������� = Not (rsSQL.EOF)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboʱ�䷶Χ_Click()
    With cboʱ�䷶Χ
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call picConMain_Resize
                Call picCondition_Resize
            End If
            .Tag = .ListIndex
        End If
    End With
End Sub


Private Sub Cboҽ������_Click()
    mParams.intAdviceType = Cboҽ������.ListIndex
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim strReturn As String
    
    Select Case Control.Id
        '''''�ļ�
        Case mconMenu_File_PrintSet     '��ӡ����
            zlPrintSet
        Case mconMenu_File_Preview      '��ӡԤ��
            zlSubPrint 2
        Case mconMenu_File_Print        '��ӡ
            zlSubPrint 1
        Case mconMenu_File_Excel        '�����Excel
            zlSubPrint 3
        Case mconMenu_File_Dept_BillPrintTotal              '��ӡ�����嵥
            Call BillPrint_Total
        Case mconMenu_File_Dept_BillPrintRestore            '��ӡ��ҩ֪ͨ��
            Call BillPrint_Restore
        Case mconMenu_File_Dept_BillPrintWait               '��ӡҩƷ��ҩ��
            Call BillPrint_Wait
        Case mconMenu_File_Parameter                    '��������
            ResetParams
        Case mconMenu_File_Exit                         '�˳�
            Unload Me
        
        ''''�༭
        Case mconMenu_Edit_Dept_Verify                    '��ҩ
            Call DrugStoreWork_Send
        Case mconMenu_Edit_Dept_Reject                    '�ܷ�ȷ��
            Call DrugStoreWork_Reject
        Case mconMenu_Edit_Dept_RejectRestore             '�ܷ��ָ�
            Call DrugStoreWork_RejectRestore
        Case mconMenu_Edit_Dept_Return                    '��ҩ
            Call DrugStoreWork_Return
        
        Case mconMenu_Edit_Dept_ReturnOther               '������ҩ������
            ShowWindow_ReturnOther
        Case mconMenu_Edit_Dept_ReVerify                  'ҩƷ��ҩ����
            Call ShowWindow_ReVerify("")
        Case mconMenu_Edit_Dept_StopFlag                  'ֹͣ��ҩ���
            ShowWindow_StopFlag
        Case mconMenu_Edit_Dept_VerifySign                    '��֤ǩ��
            If gblnESign���ŷ�ҩ = True Then mfrmDetail.VerifySign
            
        ''''�鿴
        Case mconMenu_View_ToolBar_Button               '��׼��ť
            Control.Checked = Not Control.Checked
            Me.cbsMain(2).Visible = Control.Checked
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Text                 '�ı���ǩ
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Size                 '��ͼ��
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_StatusBar                    '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3                   '�ֺ�����
            mParams.intFont = Val(Control.Parameter)
            Call SetFontSize
            Call zlDatabase.SetPara("����", mParams.intFont, glngSys, 1342)
        Case mconMenu_View_Find                         '����
            FindRow
        Case mconMenu_View_FindNext                     '������һ��
            FindRowNext
        Case mconMenu_View_SelAll                       'ȫѡ
            If Not mfrmDetail Is Nothing Then mfrmDetail.SetAllReturn
        Case mconMenu_View_ClsAll                       'ȫ��
            If Not mfrmDetail Is Nothing Then mfrmDetail.SetAllNotReturn
        Case mconMenu_View_Refresh                      'ˢ��
            cmdRefresh_Click
        
        ''''����
        Case mconMenu_Help_Help                         '����
            Call ShowHelp(App.ProductName, Me.hWnd, "Frm���ŷ�ҩ����")
        Case mconMenu_Help_Web                          'WEB�ϵ�����
        Case mconMenu_Help_Web_Home                     '������ҳ
            Call zlHomePage(Me.hWnd)
        Case mconMenu_Help_Web_Forum                    '������̳
            Call zlWebForum(Me.hWnd)
        Case mconMenu_Help_Web_Mail                     '���ͷ���
            Call zlMailTo(Me.hWnd)
        Case mconMenu_Help_About                        '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case mconMenu_List_ShowReject, mconMenu_List_OnlyShowDept, mconMenu_List_ShowOther, mconMenu_List_ShowAll, mconMenu_List_ShowClin, mconMenu_List_ShowTech, mconMenu_List_ShowArea, mconMenu_List_Sort
            '�б�Ϳ�����ʾ��ʽ
            Call SetListItemCheck(Control)
        Case mconMenu_Edit_Dept_Packer
            If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
                strReturn = mobjDrugMAC.DrugPackerSet(gcnOracle, mblnPackerConnect)
                mblnPackerConnect = (Left(strReturn, 1) = 1)
                
                '��������ͼ��״̬
                Call SetPacker
            End If
        Case mconMenu_File_Exit                      '�˳�
            Unload Me
        
        ''''�����ȼ�
        Case mconMenu_Edit_Dept_Hot_IC
            If lblPatiInputType.Tag = mInputType.IC�� Then
                Call cmdIC_Click
            End If
        Case Else
            If Control.Id > 401 And Control.Id < 499 Then
                'ִ���Զ��屨��
                Call BillPrint_Custom(Control)
            End If
            
            ''''�����˵�
            If Control.Id >= mconMenu_Input_Dept_HosNumber And Control.Id <= mconMenu_Input_Dept_HosNumber + 8 + mintCardCount Then
                '¼�뷽ʽ
                Call SetInputPopupCheck(Control)
            End If
            
'            'ҩ���Զ���ҩ�ӿڲ˵�
'            If Control.Id > mconMenu_AutoSend And Control.Id < mconMenu_AutoSend + 10 Then
'                gobjPackerZY.SetInterface Control.Id - mconMenu_AutoSend - 1, mParams.lngҩ��ID
'            End If
    End Select
    
    ''''ҩƷ��ҩ;������
    If Control.Id > mconMenu_TypePopup And Control.Id < mconMenu_TypePopup + mTypeCount + 1 Then
        Dim i As Integer
        Dim objPopup As CommandBarControl
        Dim strType As String
        
        Control.Checked = Not Control.Checked
        
        For i = 1 To mTypeCount
            Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_TypePopup + i, , True)
            If Not objPopup Is Nothing Then
                If objPopup.Checked = True Then
                    strType = strType & ";" & objPopup.Caption & ";"
                End If
            End If
        Next
        
        With Lvw��ҩ;��
            For i = 1 To .ListItems.count
                If InStr(1, strType, ";" & .ListItems(i).Tag & ";") > 0 Then
                    .ListItems(i).Checked = True
                Else
                    .ListItems(i).Checked = False
                End If
            Next
        End With
    End If
    
    '�������򵯳��˵�
    If Control.Id > mconMenu_SortPopup And Control.Id < mconMenu_SortPopup + 10 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_SortPopup)
        If Not objPopup Is Nothing Then
            For Each cbrControl In objPopup.CommandBar.Controls
                cbrControl.Checked = False
            Next
        End If
        
        Control.Checked = True
        If mParams.int�������� <> Control.Id - mconMenu_SortPopup Then
            mParams.int�������� = Control.Id - mconMenu_SortPopup
            cmdRefreshDept_Click
        End If
    End If
    
    '��Ϣ���Ѳ˵�
    If Control.Id > mconMenu_File_Message And Control.Id < mconMenu_File_Message + 10000 Then
        Call ExecuteWriteOffByMessage(Control)
    End If
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '��ӡ�Զ��屨��
    'Ĭ�ϲ�����ҩƷ=ҩƷid��ҩ��=ҩ��id������ID=����id��סԺ��=סԺ�ţ�NO=����NO����������=ҩƷ�շ���¼.����
    
    Dim str��ǰ���� As String
    Dim Int���� As Integer, strNo As String
    Dim lngҩƷid As Long
    Dim strName As String
    
    str��ǰ���� = mfrmDetail.GetRecordInfo
    
    If str��ǰ���� <> "" Then
        Int���� = Val(Split(str��ǰ����, "|")(0))
        strNo = Split(str��ǰ����, "|")(1)
        lngҩƷid = Val(Split(str��ǰ����, "|")(2))
    End If
    
    strName = Split(Control.Parameter, ",")(1)
    
    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
        "ҩƷ=" & IIf(lngҩƷid = 0, "", lngҩƷid), _
        "ҩ��=" & IIf(mcondition.lngҩ��ID = 0, "", mcondition.lngҩ��ID), _
        "����ID=" & IIf(mcondition.lng����ID = 0 Or mcondition.lng����ID = -1, "", mcondition.lng����ID), _
        "סԺ��=" & mcondition.strסԺ��, _
        "NO=" & strNo, _
        "��������=" & IIf(Int���� = 0, "", Int����))
End Sub
Private Sub zlSubPrint(ByVal bytMode As Byte)
    'bytMode��1-��ӡ��2-Ԥ����3-�����Excel
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim strTitle As String
    
    'ȡ��ӡ�б����
    Set ObjThis = mfrmDetail.GetPrintObject(True)
    
    If ObjThis Is Nothing Then
        mfrmDetail.GetPrintObject False
        Exit Sub
    End If
    
    Select Case tbcDetail.Selected.Index
        Case mListType.��ҩ
            strTitle = "ҩƷ��ҩ�嵥"
        Case mListType.����
            strTitle = "ҩƷ���ܷ�ҩ�嵥"
        Case mListType.�ܷ�
            strTitle = "ҩƷ�ܷ��嵥"
        Case mListType.ȱҩ
            strTitle = "ҩƷȱҩ�嵥"
        Case mListType.��ҩ
            strTitle = "ҩƷ��ҩ�嵥"
    End Select
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ӡ��:" & gstrUserName
    ObjAppRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ʼʱ��:" & Format(Dtp��ʼʱ��.Value, "yyyy-MM-dd HH:mm:ss")
    ObjAppRow.Add "����ʱ��:" & Format(Dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = strTitle
    Set objPrint.Body = ObjThis
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
    mfrmDetail.GetPrintObject False
End Sub
Private Sub SetFontSize()
    Dim intFont As Integer
    Dim stdfnt As StdFont
    
    Select Case mParams.intFont
        Case 0
            intFont = 9
        Case 1
            intFont = 11
        Case 2
            intFont = 15
        Case Else
            intFont = 9
    End Select
    
    mfrmDetail.SetFontSize intFont
    
    If Not tbcDetail.PaintManager.Font Is Nothing Then
        With tbcDetail
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = intFont
End Sub
Private Sub SetInputPopupCheck(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_InputPopup)
    If Not cbrMenuBar Is Nothing Then
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            cbrControl.Checked = (cbrControl.Id = Control.Id)
        Next
        
        If Split(Control.Caption, "(")(0) = "��ҩ����" Then
            lblPatiInputType.Caption = Split(Control.Caption, "(")(0)
        Else
            lblPatiInputType.Caption = Split(Control.Caption, "(")(0) & "��"
        End If

        lblPatiInputType.Tag = Control.Id - mconMenu_Input_Dept_HosNumber
        
        mParams.int����ģʽ = Val(lblPatiInputType.Tag)
        txtInput.Text = ""
        If txtInput.Visible And txtInput.Enabled Then txtInput.SetFocus
        
        zlfuncCard_SetText txtInput, Control.Parameter
        
        Call picConMain_Resize
    End If
End Sub
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
        
    With fraColorStateSend
        .ZOrder 0
        .Top = stbThis.Top + 90
        .Left = stbThis.Panels(2).Left + stbThis.Panels(2).Width - .Width - 50
    End With
    
    With fraColorStateReturn
        .ZOrder 0
        .Top = fraColorStateSend.Top
        .Left = stbThis.Panels(2).Left + stbThis.Panels(2).Width - .Width - 50
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3       '����
            Control.Checked = Val(Control.Parameter) = mParams.intFont
    End Select
End Sub

Private Sub chkAll_Click(Index As Integer)
    Dim i As Long
    
    If chkAll(Index).Value = 2 Then Exit Sub
    
    With tvwList(Index)
        For i = 1 To .Nodes.count
            If .Nodes(i).Parent Is Nothing Then
                .Nodes(i).Checked = (chkAll(Index).Value = 1)
                TvwCheckNode .Nodes(i), .Nodes(i).Checked
            End If
        Next
    End With
End Sub


Private Sub chkSend_Click(Index As Integer)
    Dim objChk As CheckBox
    Dim blnAllUnCheck As Boolean
    
    If mblnStart = False Then Exit Sub
    
    blnAllUnCheck = True
    
    For Each objChk In chkSend
        If objChk.Value = 1 Then
            blnAllUnCheck = False
        End If
    Next
    
    If blnAllUnCheck = True Then
        chkSend(Index).Value = 1
    End If
End Sub

Private Sub cmdIC_Click()
    Dim strOutXML As String
    
    If Val(lblPatiInputType.Tag) = mInputType.IC�� Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtInput.Text = mobjICCard.Read_Card()
            If txtInput.Text <> "" Then Call txtInput_KeyPress(vbKeyReturn)
        End If
    Else
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlReadCard(Me, mlngMode, Val(Split(txtInput.Tag, "|")(gCardFormat.�����ID)), True, "", txtInput.Text, strOutXML)
            If txtInput.Text <> "" Then Call txtInput_KeyPress(vbKeyReturn)
        End If
    End If
End Sub

Private Sub cmdListSel_Click()
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_ListPopup)
    If Not objPopup Is Nothing Then
        For Each cbrControl In objPopup.CommandBar.Controls
            If Trim(txtInput.Text) = "" Then
                If cbrControl.Id >= mconMenu_List_ShowAll And cbrControl.Id <= mconMenu_List_ShowArea Then
                    cbrControl.Visible = True
                End If
            Else
                If cbrControl.Id >= mconMenu_List_ShowAll And cbrControl.Id <= mconMenu_List_ShowArea Then
                    cbrControl.Visible = False
                End If
            End If
            
            If cbrControl.Id = mconMenu_List_ShowReject Then
                cbrControl.Visible = Not (tbcDetail.Selected.Index = mListType.��ҩ)
            End If
        Next
        
        objPopup.CommandBar.ShowPopup
    End If
End Sub


Private Sub cmdRefresh_Click()
    If Val(tvwList(IIf(tbcDetail.Selected.Index = 4, mDeptType.��ҩ, mDeptType.��ҩ)).Tag) = 0 Then Exit Sub
    
    GetCondition
    
    mdate�ϴ�ˢ��ʱ�� = zlDatabase.Currentdate
    
    Call UpdateDeptListRecord(IIf(tbcDetail.Selected.Index = 4, mDeptType.��ҩ, mDeptType.��ҩ))
    
    Select Case tbcDetail.Selected.Index
        Case mListType.��ҩ, mListType.����, mListType.�ܷ�
            ClearDetailList mListType.��ҩ
            
            Call RefreshSendDetail
        Case mListType.��ҩ
            ClearDetailList mListType.��ҩ
            
            Call RefreshReturnDetail
    End Select
End Sub

Private Sub cmdRefreshDept_Click()
    Dim blnExecute As Boolean
    
    If mblnFreshDeptList = True Then Exit Sub
    
    mblnFreshDeptList = True
    
    'ˢ��ʱ������б�
    ClearTreeView IIf(tbcDetail.Selected.Index = mListType.��ҩ, 1, 0)
    
    Call GetCondition
    
    If mblnInput = True Then
        blnExecute = True
    ElseIf WarTimeArea = True Then
        blnExecute = True
    End If
    
    If blnExecute Then
        Call AviShow(Me)
    
        Select Case tbcDetail.Selected.Index
            Case mListType.��ҩ, mListType.����
                Call RefreshSendDept
            Case mListType.��ҩ
                Call RefreshReturnDept
        End Select
        
        Call AviShow(Me, False)
    End If
    
    mblnFreshDeptList = False
End Sub
Private Sub cmd��ҩ;��_Click()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With Lvw��ҩ;��
        .Visible = True
        
        .Top = picCondition.Top + picConOther.Top + txt��ҩ;��.Top + txt��ҩ;��.Height + lngTop
        .Left = picCondition.Left + picConOther.Left + txt��ҩ;��.Left
        .Width = txt��ҩ;��.Width * 3
        .Height = picDeptList.Height + picConOther.Height - txt��ҩ;��.Top - txt��ҩ;��.Height - 50
        
        .SetFocus
        .ZOrder 0
    End With
End Sub

Private Sub cmdҩƷ����_Click()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With LvwҩƷ����
        .Visible = True
        
        .Top = picCondition.Top + picConOther.Top + txtҩƷ����.Top + txtҩƷ����.Height + lngTop
        .Left = picCondition.Left + picConOther.Left + txtҩƷ����.Left
        .Width = txtҩƷ����.Width * 2
        .Height = picDeptList.Height + picConOther.Height - txtҩƷ����.Top - txtҩƷ����.Height - 50
        
        .SetFocus
        .ZOrder 0
    End With
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hWnd
        Case 2
            Item.Handle = picDeptList.hWnd
        Case 3
'            Item.Handle = tbcDetail.hWnd
            
    End Select
End Sub

Private Sub Form_Activate()
    Call picConMain_Resize
    Call picCondition_Resize
    
    TimerAuto.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lvw��ҩ;��.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw��ҩ;��)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw��ҩ;��)
            End If
        End If
    End If
    
    If LvwҩƷ����.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(LvwҩƷ����)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(LvwҩƷ����)
            End If
        End If
    End If
    
    '����
    If tbcDetail.Selected.Index = mListType.��ҩ Or tbcDetail.Selected.Index = mListType.��ҩ Then
        If KeyCode = vbKeyF3 Then
            FindRowNext
        End If
    End If
    
    'Ctrl+F4  ��IC��
    If KeyCode = vbKeyF4 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then
            If lblPatiInputType.Tag = mInputType.IC�� Then
                Call cmdIC_Click
            End If
        End If
    End If
End Sub

Private Sub UnSelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.count
        UserListView.ListItems(n).Checked = False
    Next
End Sub
Private Sub SelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.count
        UserListView.ListItems(n).Checked = True
    Next
End Sub
Private Sub Form_Load()
    Dim strStart As String
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnStart = False
    mblnEnter = False
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    Me.Width = mcstlngWinNormalWidth
    Me.Height = mcstlngWinNormalHeight
    
    On Error Resume Next
    
    'IC���ӿ�
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    'һ��ͨ�ӿ�
    mstrCardType = zlfuncCard_Ini(mobjSquareCard, Me, mlngMode)
    
    '��ʼ��������ʾ
    mParams.blnShowReject = False
    mParams.int�������� = 1
    
    '��ʼ����ѯ���Ѳ���
    With mFindWar
        .blnNoAsk_Dept_Send = False
        .blnNoAsk_Dept_Sended = False
        .blnProc_Dept_Send = True
        .blnProc_Dept_Sended = True
        .blnNoAsk_Rec = False
        .blnProc_Rec = True
    End With
    
    'ȡȨ��
    Call GetPrivs
    
    'ȡ����
    Call GetParams
    
    Call SetFontSize
    
    '��ʼ������
    mcondition.lngҩ��ID = mParams.lngҩ��ID
    mstrDeptNode = GetDeptStationNode(mParams.lngҩ��ID)
   
    If Load��ҩҩ�� = False Then Exit Sub
    '�Ƿ����
    mblnEnter = True
    
    Call Loadʱ�䷶Χ
    Call Loadȡ�Զ��巢ҩ����
    
    Call Loadҽ������
    
    Call Load��ҩ;��
    Call LoadҩƷ����(Val(cbo��ҩҩ��.ItemData(cbo��ҩҩ��.ListIndex)))
    
    Call SetColorState
    
    '------------------------------------------------------------------
    'ҩƷ�ְ����ӿ�
    mblnStartPacker = False
    mblnPackerConnect = False
    
    Set mclsComLib = New zl9ComLib.clsComLib
    
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        Set mobjDrugMAC = Nothing
        '�����½ӿ�
        Set mobjDrugMAC = CreateObject("zlDrugMachine.clsDrugMachine")
        If err.Number <> 0 Then
            '��ξɽӿ�
            Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
        End If
    Else
        Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
    End If
    
    err.Clear: On Error GoTo 0
    
    If TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        '�½ӿ�
        ''��ȡ�ӿڵ�Ȩ��
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�")) & ";"
        If strPrivs Like "*;����;*" Then
            
            mblnPackerConnect = mobjDrugMAC.Init(1, mclsComLib, strMessage)
        Else
            mblnPackerConnect = False
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        '�ɽӿ�
        
        '�������ע�����Ϊ0��ʾδ����סԺҩ���ӿ�
        strStart = GetSetting("ZLSOFT", "����ģ��\�Զ���ҩ��", "����סԺҩ��")
        If Not mobjDrugMAC Is Nothing And strStart <> "0" Then
            mblnStartPacker = True
            If mobjDrugMAC.DBConnect Then
                mblnPackerConnect = True
            Else
                mblnPackerConnect = False
                MsgBox "ҩƷ�ְ����ӿ����ݿ�δ���������ӣ����ܴ������ݣ�" & vbCrLf & "��ʾ��������ڲ˵���ѡ���ֶ������������ӡ�", vbInformation, gstrSysName
            End If
        End If
        
        'ҩ���Զ���ҩ���ӿ���ز˵���״̬������
        Call SetPacker
    Else
        mblnPackerConnect = False
    End If
    
    '��ҽӿ�
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    
    '------------------------------------------------------------------
    
    '�ŵ�InitComandBarsǰ�棬������Щ��ť���Ի�������Ч
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        '�ָ����Ի�����
        LoadCustomSet
    End If
    
    '��ʼ���˵�������ҳ��Ƚ��沼��
    Call InitComandBars
    Call InitPanes
    Call InitTabControl
    
    '����Զ��屨��
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrprivs)
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        '�ָ�¼��״̬
        Call SetInputState(mParams.int����ģʽ)
        
        '�ָ�����
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
    
    Me.picColorStateSend(6).BackColor = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\1345", "δ���ҽ����ɫ", 33023)
    
    Call RestoreWinState(Me, App.ProductName)
   
    '����ǩ���ӿڿ���
    gblnESign���ŷ�ҩ = EsignIsOpen(mParams.lngҩ��ID)
    gblnESignUserStoped = False
    If gblnESign���ŷ�ҩ = True Then
        On Error Resume Next
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        err.Clear: On Error GoTo 0
        If Not gobjESign Is Nothing Then
            If Not gobjESign.Initialize(gcnOracle, glngSys) Then
                Set gobjESign = Nothing
                gblnESign���ŷ�ҩ = False
            Else
                gblnESign���ŷ�ҩ = True
                gblnESignUserStoped = gobjESign.CertificateStoped(gstrUserName)
            End If
        Else
            gblnESign���ŷ�ҩ = False
        End If
    End If
    
    mblnIs�������� = Is��������(mParams.lngҩ��ID)
    Call Cbo��ҩҩ��_Click
    
    '��ʼ����Ϣ����
    If mPrives.bln��ҩ���� Then
        err = 0
        On Error Resume Next
        Set mobjMipModule = New zl9ComLib.clsMipModule
        Call mobjMipModule.InitMessage(glngSys, mlngMode, mstrPrivs)
        Call AddMipModule(mobjMipModule)
        
        If Not mobjMipModule Is Nothing Then
            Call InitMsgRec
        End If
    End If
    mblnStart = True
End Sub


Private Sub SetInputState(ByVal intType As Integer)
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Input_Dept_HosNumber + intType, , True)
    If Not cbrControl Is Nothing Then
        SetInputPopupCheck cbrControl
    End If
End Sub
Private Sub SetSendTypePosition()
    '������ҩ����ѡ���λ��
    Dim n As Integer
    Dim dbl����� As Double
    Dim dblTmp As Double
    Dim dblSumTmp As Double
    Dim int���� As Integer
    Dim dblCheckControlH As Double
    Const cst������ = 50
    Const cst�о� = 50
    
    picSendType.Visible = mblnExistOtherSendType
    picShowSendType.Visible = mblnExistOtherSendType
    
    If picShowSendType.Visible = False Then Exit Sub
    
    If chkSendType.UBound > 0 Then
        dbl����� = picSendType.Width - 100
        dblCheckControlH = chkSendType(0).Height
        picSendType.Height = chkSendType(0).Height + 75
        
        int���� = 0
        dblSumTmp = chkSendType(0).Width + cst������
        For n = 1 To chkSendType.UBound
            dblTmp = chkSendType(n).Width + dblSumTmp
            
            If dblTmp <= dbl����� Then
                chkSendType(n).Top = chkSendType(n - 1).Top
                chkSendType(n).Left = chkSendType(n - 1).Left + chkSendType(n - 1).Width + cst������
                dblSumTmp = dblSumTmp + chkSendType(n).Width + cst������
            Else
                '�����У������������ؼ�λ��
                int���� = int���� + 1
                chkSendType(n).Left = chkSendType(0).Left
                chkSendType(n).Top = chkSendType(0).Top + (dblCheckControlH + cst�о�) * int����
                dblSumTmp = chkSendType(n).Width + cst������

                picSendType.Height = chkSendType(n).Top + chkSendType(n).Height + 50
            End If
        Next
    End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < mcstlngWinNormalWidth Then Me.Width = mcstlngWinNormalWidth
    If Me.Height < mcstlngWinNormalHeight Then Me.Height = mcstlngWinNormalHeight
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngMyWindow = 0
    mblnFreshDeptList = False
    
    'ж��IC��ˢ���ӿ�
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    'ж��ҩƷ�Զ���ҩ���ӿ�
    Set mobjDrugMAC = Nothing
    Set mclsComLib = Nothing
    
    'ж��һ��ͨ�ӿ�
    mstrCardType = ""
    Call zlfuncCard_Unload(mobjSquareCard)
    
    'ж�����õĴ���
    If Not mfrmDetail Is Nothing Then
        Unload mfrmDetail
        Set mfrmDetail = Nothing
    End If
    
    '���洰�ڼ�����
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
        
        Call SaveWinState(Me, App.ProductName)
    
        '������Ի�����
        SaveCustomSet
    End If
    
    'ж����Ϣ����
    If Not mobjMipModule Is Nothing Then
        Call mobjMipModule.CloseMessage
        Call DelMipModule(mobjMipModule)
        Set mobjMipModule = Nothing
    End If
    
    'ж����ҽӿ�
    Call zlPlugIn_Unload(mobjPlugIn)
End Sub

Private Sub lblComment_Click()
    ShowOtherConditon
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Button = 1 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_InputPopup)
        If Not objPopup Is Nothing Then
            Set cbrControl = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Input_Dept_ReceiveNO, , True)
            cbrControl.Enabled = tbcDetail.Selected.Index <> mListType.��ҩ
            
            Set cbrControl = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Input_Dept_BatchSendNO, , True)
            cbrControl.Enabled = tbcDetail.Selected.Index = mListType.��ҩ
            
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub Lvw��ҩ;��_DblClick()
    ReturnSelected��ҩ;�� 0
End Sub

Private Sub Lvw��ҩ;��_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw��ҩ;��
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "���и�ҩ;��" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub Lvw��ҩ;��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ReturnSelected��ҩ;�� 1
    End If
End Sub

Private Sub Lvw��ҩ;��_LostFocus()
    Lvw��ҩ;��.Visible = False
End Sub


Private Sub Lvw��ҩ;��_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_TypePopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub LvwҩƷ����_DblClick()
    ReturnSelected���� 0
End Sub

Private Sub LvwҩƷ����_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With LvwҩƷ����
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "����ҩƷ����" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub LvwҩƷ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ReturnSelected���� 1
    End If
End Sub

Private Sub LvwҩƷ����_LostFocus()
    LvwҩƷ����.Visible = False
End Sub

Private Sub mobjMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '���յ���ϢҪ��֤��Ϣ����Ч�ԣ���ҩ����
    '������Ϣ���ݼ�������̬������Ϣ��Ŀ�����˵�
    '��Ϣ�����˵������ʾ5����������5��ʱ������һ����ʾ��ȫ�����ˡ�
    Dim objXML As New zl9ComLib.clsXML
    Dim rsMsg As ADODB.Recordset
    Dim blnValid As Boolean
    Dim str���� As String
    Dim str����id As String
    Dim str���� As String
    Dim strסԺ�� As String
    Dim str����ʱ�� As String
    Dim strSQL As String
    Dim rstemp As Recordset
    Dim i As Integer
    
'    'ZLHIS_CHARGE_001
'    patient_info ������Ϣ
'    patient_id ����id
'    patient_name ����
'    identity_card ���֤��
'    in_number סԺ��
'    out_number �����
'    cancel_reqeust ��������
'    cancel_charge
'       charge_id ����id
'       request_kind �������
'       request_time ����ʱ��
'       request_person ������Ա
'       cancel_item_id ������Ŀid
'       cancel_item_title ������Ŀ
'       calcel_num ��������
'       audit_dept_id ��˲���id
'       audit_dept_title ��˲���


    '��Ϣ����Ϊ��ʱ�˳�
    
    If mobjMipModule Is Nothing Then Exit Sub
    
    '��Ϣ��������ʧ��ʱ��������Ϣ
    If mobjMipModule.IsConnect = False Then Exit Sub
    
    If objXML Is Nothing Then Exit Sub
    '��XML�ļ�
    objXML.OpenXMLDocument strMsgContent
    
    '�����Ϣ�Ƿ���Ч����Ҫ�Ǽ��ҩ��
    If objXML.GetMultiNodeRecord("cancel_charge", rsMsg) = False Then Exit Sub
    If rsMsg Is Nothing Then Exit Sub
    If rsMsg.RecordCount = 0 Then Exit Sub
    
    blnValid = False
    Do While Not rsMsg.EOF
        If rsMsg("node_name").Value = "audit_dept_id" Then
            If Val(rsMsg("node_value").Value) = mcondition.lngҩ��ID Then
                blnValid = True
                Exit Do
            End If
        End If
        rsMsg.MoveNext
    Loop
    If blnValid = False Then Exit Sub
    
    '�������Ч��Ϣ�������Ϣ���ݼ�
'    str���� = ""
'    If objXML.GetSingleNodeValue("patient_id", str����id, xsString) = False Then Exit Sub
'    If objXML.GetSingleNodeValue("patient_name", str����, xsString) = False Then Exit Sub
'    If objXML.GetSingleNodeValue("in_number", strסԺ��, xsString) = False Then Exit Sub
'    If objXML.GetSingleNodeValue("request_time", str����ʱ��, xsString) = False Then Exit Sub
    
    Call mobjMipModule.ShowMessage(strMsgItemIdentity, "���µ��������룬�����Աע������Ϣ�б��в鿴�ʹ���", "��Ϣ����")
    
    '��Ϣ��Ч������ݿ��ȡ��Ϣ
    strSQL = "select distinct A.����ʱ��,B.����id,B.����,C.סԺ�� from ���˷������� A,סԺ���ü�¼ B,������ҳ C where A.����ID=B.ID And B.����ID=C.����ID And B.��ҳid=C.��ҳid And A.��˲���ID=[1] And A.����ʱ��>[2] and A.����� is null and A.״̬=0"
    Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "", mcondition.lngҩ��ID, mdateBegin)
    
    Call InitMsgRec
    With mrsReceiveMsg
        For i = 1 To rstemp.RecordCount
            .AddNew
            !���� = ""
            !����ID = Val(rstemp!����ID)
            !���� = NVL(rstemp!����, "")
            !סԺ�� = NVL(rstemp!סԺ��, "")
            !����ʱ�� = Format(rstemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
            .Update
            
            rstemp.MoveNext
        Next
    End With
    
    '������Ϣ�˵�
    
    Call SetMessageBar(mrsReceiveMsg)
End Sub

Private Sub picColorStateSend_Click(Index As Integer)
    On Error GoTo errHandle
    
    If Index = 6 Then
        cmdialog.CancelError = True
        cmdialog.ShowColor
        picColorStateSend(6).BackColor = cmdialog.Color
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\1345", "δ���ҽ����ɫ", Val(picColorStateSend(6).BackColor)
        Exit Sub
    End If
errHandle:
End Sub

Private Sub picCondition_Resize()
    On Error Resume Next
    
    With picConMain
        .Width = picCondition.Width
    End With
    
    With picConOther
        If Val(picShowOther.Tag) = 1 Then
            .Visible = True
            .Left = picConMain.Left
            .Top = picConMain.Top + picConMain.Height - 60
            .Width = picCondition.Width
        Else
            .Visible = False
        End If
    End With
    
    With picDeptList
        If Val(picShowOther.Tag) = 1 Then
            .Top = picConOther.Top + picConOther.Height
        Else
            .Top = picConMain.Top + picConMain.Height
        End If
        
        .Left = picConMain.Left
        .Width = picCondition.Width
        .Height = picCondition.Height - .Top - 50
    End With
End Sub

Private Sub picConMain_Resize()
    On Error Resume Next
    
    With cbo��ҩҩ��
        .Width = picConMain.Width - .Left - 50
    End With
    
    With fraLineH1
        .Width = picConMain.Width + 150
    End With
    
    With cboʱ�䷶Χ
        .Left = cbo��ҩҩ��.Left
        .Width = cbo��ҩҩ��.Width
    End With
    
    If cboʱ�䷶Χ.ListIndex <> 3 Then
        lblTimeBegin.Visible = False
        Dtp��ʼʱ��.Visible = False
        lblTimeEnd.Visible = False
        Dtp����ʱ��.Visible = False
        
        With lblPatiInputType
            .Top = lblʱ�䷶Χ.Top + lblʱ�䷶Χ.Height + 180
        End With
        
        With txtInput
            .Top = cboʱ�䷶Χ.Top + cboʱ�䷶Χ.Height + 60
            .Width = cboʱ�䷶Χ.Width
        End With
    Else
        lblTimeBegin.Visible = True
        Dtp��ʼʱ��.Visible = True
        lblTimeEnd.Visible = True
        Dtp����ʱ��.Visible = True
        
        With lblTimeBegin
            .Top = lblʱ�䷶Χ.Top + lblʱ�䷶Χ.Height + 180
        End With
        
        With Dtp��ʼʱ��
            .Top = cboʱ�䷶Χ.Top + cboʱ�䷶Χ.Height + 60
            .Width = cbo��ҩҩ��.Width
        End With
        
        With lblTimeEnd
            .Top = lblTimeBegin.Top + lblTimeBegin.Height + 180
        End With
        
        With Dtp����ʱ��
            .Top = Dtp��ʼʱ��.Top + Dtp��ʼʱ��.Height + 60
            .Width = cbo��ҩҩ��.Width
        End With
        
        With lblPatiInputType
            .Top = lblTimeEnd.Top + lblTimeEnd.Height + 180
        End With
        
        With txtInput
            .Top = Dtp����ʱ��.Top + Dtp����ʱ��.Height + 60
            .Width = cbo��ҩҩ��.Width
        End With
    End If
    
    With cmdIC
        .Visible = (Val(Split(txtInput.Tag, "|")(gCardFormat.ˢ����־)) = 1)
        .Top = txtInput.Top
        .Left = picConMain.Width - .Width - 80
        
        If Val(Split(txtInput.Tag, "|")(gCardFormat.ˢ����־)) = 1 Then
            txtInput.Width = .Left - txtInput.Left - 50
        Else
            txtInput.Width = cbo��ҩҩ��.Width
        End If
    End With
    
    With chkSend(0)
        .Top = txtInput.Top + txtInput.Height + 60
    End With
    
    With chkSend(1)
        .Top = chkSend(0).Top
    End With
    
    If picConMain.Width > chkSend(1).Left + chkSend(1).Width + chkSend(2).Width + 200 Then
        chkSend(2).Top = chkSend(1).Top
        chkSend(2).Left = chkSend(1).Left + chkSend(1).Width + 100
        lbl��ҩ����.Top = chkSend(0).Top
    Else
        chkSend(2).Top = chkSend(0).Top + chkSend(0).Height + 50
        chkSend(2).Left = chkSend(0).Left
        lbl��ҩ����.Top = chkSend(0).Top + 100
    End If
    
    '�Զ��巢ҩ���͵�λ��
    Call SetSendTypePosition
    If picShowSendType.Visible = True Then
        picShowSendType.Top = chkSend(2).Top + chkSend(2).Height + 100
        picShowSendType.Width = picConMain.Width - 50
        picSendType.Left = picShowSendType.Left + 240
        picSendType.Top = picShowSendType.Top + picShowSendType.Height + 50
        picSendType.Width = picConMain.Width - picSendType.Left
        
        If Val(picShowSendType.Tag) = 1 Then
            picSendType.Visible = True
            picShowOther.Top = picSendType.Top + picSendType.Height + 50
        Else
            picSendType.Visible = False
            picShowOther.Top = picShowSendType.Top + picShowSendType.Height + 50
        End If
    Else
        picShowOther.Top = chkSend(2).Top + chkSend(2).Height + 50
    End If
    
    With picShowOther
        .Left = lbl��ҩҩ��.Left
        .Width = picConMain.Width - 50
    End With
    
    With picConMain
        .Height = picShowOther.Top + picShowOther.Height
    End With
    
    With Lvw��ҩ;��
        .Top = picConOther.Top + txt��ҩ;��.Top + txt��ҩ;��.Height
        .Left = picConOther.Left + txt��ҩ;��.Left
        .Width = txt��ҩ;��.Width
        .Height = picDeptList.Height + picConOther.Height - txt��ҩ;��.Top - txt��ҩ;��.Height - 50
    End With
    
    With LvwҩƷ����
        .Top = picConOther.Top + txtҩƷ����.Top + txtҩƷ����.Height
        .Left = picConOther.Left + txtҩƷ����.Left
        .Width = txtҩƷ����.Width
        .Height = picDeptList.Height + picConOther.Height - txtҩƷ����.Top - txtҩƷ����.Height - 50
    End With
End Sub



Private Sub picConOther_Resize()
    On Error Resume Next
    
    With fraLineH2
        .Width = picConOther.Width + 150
    End With
    
    With Cboҽ������
        .Width = picConOther.Width - .Left - 50
    End With
    
    With cmd��ҩ;��
        .Left = picConOther.Width - .Width - 50
        If .Left < txt��ҩ;��.Left + 100 Then .Left = txt��ҩ;��.Left + 100
    End With
    
    With txt��ҩ;��
        .Width = cmd��ҩ;��.Left - .Left + cmd��ҩ;��.Width
    End With
    
    With cmdҩƷ����
        .Left = picConOther.Width - .Width - 50
        If .Left < txtҩƷ����.Left + 100 Then .Left = txtҩƷ����.Left + 100
    End With
    
    With txtҩƷ����
        .Width = cmdҩƷ����.Left - .Left + cmdҩƷ����.Width
    End With
    
    With picConOther
        .Height = chkDangerType(0).Top + chkDangerType(0).Height
    End With
End Sub
Private Sub picDeptList_Resize()
    On Error Resume Next
    
    With fraLineH3
        .Width = picDeptList.Width + 150
    End With
    
    With cmdRefresh
        .Left = picDeptList.Width - .Width - 100
    End With
    
    With cmdRefreshDept
        .Left = cmdRefresh.Left - .Width - 50
    End With
    
    With tvwList(mDeptType.��ҩ)
        .Top = cmdRefreshDept.Top + cmdRefreshDept.Height + 50
        .Left = 0
        .Width = picDeptList.Width - 100
        .Height = picDeptList.Height - .Top - 50
    End With
    
    With tvwList(mDeptType.��ҩ)
        .Top = tvwList(mDeptType.��ҩ).Top
        .Left = 0
        .Width = tvwList(mDeptType.��ҩ).Width
        .Height = tvwList(mDeptType.��ҩ).Height
    End With
End Sub
Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLineV1
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLineV1.Left + 50
        .Width = picDetail.Width - fraLineV1.Width
        .Height = picDetail.Height - 50
    End With
End Sub

Private Sub picShowOther_Click()
    ShowOtherConditon
End Sub


Private Sub picShowOther_Resize()
    With picUpOrDown
        .Left = picShowOther.Width - .Width
        .Top = 0
    End With
End Sub


Private Sub picShowSendType_Click()
    picShowSendType.Tag = Abs(Val(picShowSendType.Tag) - 1)
    picUpOrDown1.Picture = imgLvwSel.ListImages(Val(picShowSendType.Tag) + 3).Picture
    
    picSendType.Visible = (Val(picShowSendType.Tag) = 1)
    Call picConMain_Resize
    Call picCondition_Resize
End Sub

Private Sub picShowSendType_Resize()
    With picUpOrDown1
        .Left = picShowSendType.Width - .Width
        .Top = 0
    End With
End Sub


Private Sub picUpOrDown_Click()
    ShowOtherConditon
End Sub

Private Sub picUpOrDown1_Click()
    picShowSendType_Click
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 3 Then
        Call ShowWindow_ReVerify("")
    End If
End Sub

Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '�л�δ��ҩ������δ��ҩ�嵥�������嵥��ȱҩ�嵥���ܷ�ҩ�嵥�����ѷ�ҩ�嵥����ҩ�嵥��
    Dim cbrControl As CommandBarControl
    
    Call mfrmDetail.ShowList(Item.Index, Val(cbo��ҩҩ��.ItemData(cbo��ҩҩ��.ListIndex)))
    Call SetCommandBar(Item.Index)
    
    Select Case Item.Index
        Case mListType.��ҩ, mListType.����, mListType.�ܷ�, mListType.ȱҩ
            Me.dkpMain.FindPane(mconPane_Dept_Condition).Title = "��������(��ҩģʽ)"
            
            tvwList(mDeptType.��ҩ).Visible = True
            chkAll(mDeptType.��ҩ).Visible = True
            
            tvwList(mDeptType.��ҩ).Visible = False
            chkAll(mDeptType.��ҩ).Visible = False
            
            chkWithNotAudited.Enabled = True
        Case mListType.��ҩ
            Me.dkpMain.FindPane(mconPane_Dept_Condition).Title = "��������(��ҩģʽ)"
            
            tvwList(mDeptType.��ҩ).Visible = True
            chkAll(mDeptType.��ҩ).Visible = True
            
            tvwList(mDeptType.��ҩ).Visible = False
            chkAll(mDeptType.��ҩ).Visible = False
            
            chkWithNotAudited.Enabled = False
    End Select
    
    fraColorStateSend.Visible = (Item.Index = mListType.��ҩ)
    fraColorStateReturn.Visible = (Item.Index = mListType.��ҩ)
    
    lblPatiInputType.Caption = "סԺ�š�"
    lblPatiInputType.Tag = 0
    txtInput.Text = ""
    
    Set cbrControl = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Input_Dept_HosNumber, , True)
    If Not cbrControl Is Nothing Then
        SetInputPopupCheck cbrControl
    End If
End Sub

Private Sub DrugStoreWork_Send()
    'ҩ����������ҩ
    Dim rsSendData As ADODB.Recordset
    Dim StrCurDate As String
    
'    On Error GoTo errHandle
    
    mblnCheck = False
    
    'ȡ��ҩ���ݼ�
    Set rsSendData = mfrmDetail.GetSendRecord
    
    If rsSendData Is Nothing Then Exit Sub
    
    If MsgBox("��ȷ��Ҫ��ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign���ŷ�ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
    '��ҩ���
    If DrugStoreWork_CheckSend(rsSendData) = False Then Exit Sub
    
    'ȡϵͳʱ��
    StrCurDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    'ȡ���ܷ�ҩ��
    mcur���ܷ�ҩ�� = Val(zlDatabase.GetNextNo(20))
    
    '��ҩ����
    If DrugStoreWork_SendProc(rsSendData, StrCurDate) = False Then Exit Sub
        
    '���洦��
    If DrugStoreWork_StayProc(StrCurDate) = False Then Exit Sub
    
    '���ʴ���
    If DrugStoreWork_CancelVerifyProc(StrCurDate) = False Then Exit Sub
    
    '��ҩƷ�ְ�����������
    Call DrugStoreWork_SendToPacker(rsSendData)

    '��ӡ���ܵ���
    Call DrugStoreWork_PrintBill
    
    '��ҩ����²����б����ϸ����
    cmdRefreshDept_Click
    
    mfrmDetail.AfterSendRefresh
    
    If mcur���ܷ�ҩ�� > 0 Then
        stbThis.Panels(2).Text = "�ϴη�ҩ�ţ�" & mcur���ܷ�ҩ�� & ""
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TimerReturn_Timer()
    Dim strSQL As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    strSQL = "select count(����id) ���� from (Select distinct A.����id,A.����ʱ�� " & vbNewLine & _
        "From ���˷������� A, ҩƷ�շ���¼ B" & vbNewLine & _
        "Where A.����id = B.����id And Not Exists" & vbNewLine & _
        " (Select 1 From ��Һ��ҩ���� C Where C.�շ�id = B.ID) And ��˲���id = [1] And ����ʱ�� Between Trunc(Sysdate) And" & vbNewLine & _
        "      Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And ���ʱ�� Is Null and (B.��¼״̬=1 or mod(B.��¼״̬,3)=0))"

    Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "", mParams.lngҩ��ID)
    
    Me.stbThis.Panels(3).Text = "δ�������������" & rstemp!���� & "��"
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckGroupSend(ByVal rsGroupRec As ADODB.Recordset, ByVal lng���ID As Long, ByVal strNo As String) As Boolean
    '���ͬ��ҩƷ�Ƿ��ܹ�����
    'ǰ����ҩ������������������
    'ͬ��ҩƷ��ֻ�е����ж��Ƿ�ҩ״̬����������ȱҩ���ܷ������������ܷ�ҩ
    Dim i As Integer
    
    'Ĭ��������
    CheckGroupSend = True
    
    '���������������޸ù���
    If mParams.bln�������� = False Then Exit Function
    
    '�޷���Ĳ���
    If lng���ID = 0 Then Exit Function
    
    '���ݴ����NO�����ID���ж��Ƿ����ҩƷ���ܷ�ҩ
    With rsGroupRec
        .Filter = "NO='" & strNo & "'" & " And ���ID = " & lng���ID
        
        If .EOF Then Exit Function
        
        Do While Not .EOF
            'ֻҪ����ִ��״̬��Ϊ1���Ͳ��ܷ�ҩ�������ΣҩƷ��ҩ��ʽѡ���˸�ΣҩƷ���࣬��ô���Բ�������ΣҩƷ
            If !ִ��״̬ <> 1 And InStr(1, mParams.str��Σ����, !��ΣҩƷ) = 0 And Not mblnCheck Then
                If MsgBox("ͬ��ҩƷ�ķ�ҩ״̬��һ�£��Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mblnCheck = True
                    CheckGroupSend = True
                Else
                    mblnCheck = False
                    CheckGroupSend = False
                End If
                Exit Function
            End If
            
            .MoveNext
        Loop
    End With
End Function

Private Function Check��������(ByRef rsData As ADODB.Recordset) As Boolean
    '����:����Ƿ�����������ļ�¼�������޳�
    'return:False-����Ҫ������ʾ;True-��Ҫ�������ҩ�����������ʾ
    Dim str�շ�ID As String
    Dim rstemp As ADODB.Recordset
    
    Check�������� = False
    
    On Error GoTo errHandle
    
    With rsData
        .Filter = "ִ��״̬=" & mState.��ҩ
        .Sort = "�շ�ID"
        Do While Not .EOF
            str�շ�ID = str�շ�ID & IIf(str�շ�ID = "", "", ",") & !�շ�Id
            
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select " & vbNewLine & _
            " a.Id As �շ�id" & vbNewLine & _
            "From ҩƷ�շ���¼ A, ���˷������� B, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) C" & vbNewLine & _
            "Where a.����id = b.����id And a.Id = c.Column_Value And b.״̬ + 0 = 0"
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�շ�ID)
    
    
    '�޳��������������¼���շ�ID
    If Not rstemp.EOF Then
        Do While Not rstemp.EOF
            rsData.Filter = "�շ�ID=" & rstemp!�շ�Id
            rsData!ִ��״̬ = mState.������
            
            rstemp.MoveNext
        Loop
        
        Check�������� = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckCorrelation(ByVal intType As Integer, ByVal rsSendData As ADODB.Recordset) As Boolean
    'intType:0-��ҩ;1-��ҩ
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    Dim strNo As String, lng���� As Long, str��� As String, lng����ID As Long
    
    With rsSendData
        .Filter = "ִ��״̬=" & IIf(intType = 0, mState.��ҩ, mState.��ҩ)
        
        Do While Not .EOF
            strNo = !NO & !����
            lng���� = !����
            strNo = !NO
            lng����ID = !����ID
            str��� = NVL(!�������)
            If Not IsReceiptBalance_Charge(intType, mstrPrivs, lng����, strNo, str���, 2, 2) Then Exit Function
            If Not IsOutPatient(mstrPrivs, lng����, strNo, 2, 2) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function
Private Sub DrugStoreWork_Reject()
    'ҩ���������ܷ�ҩ
    Dim rsSendData As ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim arrSql As Variant
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    'ȡ��ҩ���ݼ�
    Set rsSendData = mfrmDetail.GetSendRecord
    arrSql = Array()
    
    With rsSendData
        .Filter = "ִ��״̬=" & mState.�ܷ�
        .Sort = "ҩƷID Asc"
        
        If .EOF Then Exit Sub
        
        Do While Not .EOF
            '��鵥��״̬
            If DeptSendWork_CheckBill(0, !�շ�Id, mParams.bln����δ��˴�����ҩ) > 0 Then Exit Sub
            
            .MoveNext
        Loop
        
        .MoveFirst
        
        
        
        Do While Not .EOF
            gstrSQL = "zl_ҩƷ�շ���¼_���žܷ�(" & !�շ�Id & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            .MoveNext
        Loop
    End With
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For lngRow = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-���þܷ�ҩƷ")
    Next
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    mfrmDetail.AfterRejectRefresh
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DrugStoreWork_RejectRestore()
    'ҩ���������ܷ�ҩ�ָ�
    Dim rsSendData As ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim lngRow As Long
    Dim arrSql As Variant
    
    On Error GoTo errHandle
    
    'ȡ��ҩ���ݼ�
    Set rsSendData = mfrmDetail.GetSendRecord
    arrSql = Array()
    
    With rsSendData
        .Filter = "ִ��״̬=" & mState.�ܷ�_�ָ�
        .Sort = "ҩƷID Asc"

        Do While Not .EOF
            gstrSQL = "zl_ҩƷ�շ���¼_���Żָ�(" & !�շ�Id & ")"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            .MoveNext
        Loop
        
        gcnOracle.BeginTrans
        blnBeginTrans = True
        For lngRow = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-�ָ��ܷ�ҩƷ")
        Next
        gcnOracle.CommitTrans
        blnBeginTrans = False
    End With
    
    mfrmDetail.AfterRejectRestoreRefresh
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub DrugStoreWork_Return()
    'ҩ����������ҩ
    Dim rsReturnData As ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim str��ҩ�� As String
    Dim dbl��ҩ�� As Double
    Dim str�۸�ʧЧ��ʾ As String
    Dim strҩƷID As String
    Dim strDate As String
    Dim bln�Ƿ�����ҩ As Boolean
    Dim blnIsReturn As Boolean
    Dim arrSql As Variant
    Dim i As Integer
    Dim strǩ����¼ As String
    Dim strSQL As String
    Dim rstemp As Recordset
    Dim int��ҩ As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    Dim bln����������ʾ As Boolean
    
    On Error GoTo errHandle
    
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign���ŷ�ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
    arrSql = Array()
    
    'ȡ��ҩ���ݼ�
    Set rsReturnData = mfrmDetail.GetReturnRecord
    
    If rsReturnData Is Nothing Then Exit Sub
    
    If MsgBox("��ȷ��Ҫ��ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    If Not CheckCorrelation(1, rsReturnData) Then Exit Sub
    
    '���ҽ��
    If CheckAdvice(rsReturnData) = False Then Exit Sub
    
    '��ҩ��ǩ��
    str��ҩ�� = ""
    If mParams.bln��ҩ��ǩ�� = True Then
        str��ҩ�� = zlDatabase.UserIdentify(Me, "��ҩ��ǩ��", glngSys, 1342, "��ҩ")
        If str��ҩ�� = "" Then
            Exit Sub
        End If
    End If
    
    '���ԭ�������������ڷ���
    '������Ż�Ч��Ϊ�գ�����ȡ���û����루���Ӵ��������У������룩
    
    '����״̬���
    With rsReturnData
        .Filter = "ִ��״̬=" & mState.��ҩ
        .Sort = "�շ�ID"
        Do While Not .EOF
            '��鵥��״̬
            If DeptSendWork_CheckBill(2, !�շ�Id, mParams.bln����δ��˴�����ҩ) > 0 Then Exit Sub
            
            .MoveNext
        Loop
    End With
    
    '���á�����������ʱ���ܽ�����ҩ����������,���¼���Ƿ��м����������ļ�¼����ֹδ��ʱˢ���嵥�󣬴������������ҩƷ����ҩ
    If mParams.bln����������ʱ���ܽ�����ҩ���� Then
        If Check��������(rsReturnData) Then
            bln����������ʾ = True
        End If
    End If
    
    '��ҩ
    With rsReturnData
        .Filter = "ִ��״̬=" & mState.��ҩ
        .Sort = "ҩƷID Asc"
        
        strDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        
        Do While Not .EOF
            If Val(!��ҩ��) = Val(!׼����) Then
                dbl��ҩ�� = Val(!ʵ������)
            Else
                dbl��ҩ�� = Val(!��ҩ��) * Val(!��װ)
            End If
            
            If dbl��ҩ�� <> 0 Then
                blnIsReturn = False
                If CheckPrice(!�շ�Id, str�۸�ʧЧ��ʾ) = False Then
                    If MsgBox("ҩƷ[" & !Ʒ�� & "(" & !��� & ")]" & str�۸�ʧЧ��ʾ, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        blnIsReturn = True
                    End If
                Else
                    blnIsReturn = True
                End If
                
                If blnIsReturn = True Then
                    gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
                    '�շ�ID
                    gstrSQL = gstrSQL & !�շ�Id
                    '�����
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '���ʱ��
                    gstrSQL = gstrSQL & ",To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                    '����
                    gstrSQL = gstrSQL & "," & IIf(IsNull(!����), "NULL", IIf(Mid(!����, 1, 1) = "(", "NULL", "'" & Mid(!����, 1, 8) & "'"))
                    'Ч��
                    gstrSQL = gstrSQL & "," & IIf(IsNull(!Ч��), "NULL", IIf(!Ч�� = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
                    '����
                    gstrSQL = gstrSQL & "," & IIf(IsNull(!����), "NULL", "'" & !���� & "'")
                    '��ҩ��
                    gstrSQL = gstrSQL & "," & dbl��ҩ��
                    '��ҩ�ⷿ
                    gstrSQL = gstrSQL & ",NULL"
                    '��ҩ��
                    gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
                    '����λ��
                    gstrSQL = gstrSQL & "," & mParams.int����λ��
                    '����
                    gstrSQL = gstrSQL & ",2"
                    '���ܷ�ҩ��
                    gstrSQL = gstrSQL & ",Null"
                    gstrSQL = gstrSQL & ")"
    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                    
                    bln�Ƿ�����ҩ = True
                    
                    If InStr("," & strҩƷID & ",", "," & !ҩƷid & ",") = 0 Then
                        strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & !ҩƷid
                    End If
                    
                    strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(!�շ�Id) & "," & dbl��ҩ��
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    '��ʾͣ��ҩƷ
    If strҩƷID <> "" Then
        int��ҩ = 1
        Call CheckStopMedi(strҩƷID, int��ҩ)
        If int��ҩ = 2 Then Exit Sub
    End If
    
    '���д�����ҩ����
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-ҩƷ��ҩ")
    Next
    
    '����ǩ������
    If UBound(arrSql) >= 0 And gblnESign���ŷ�ҩ = True And gblnESignUserStoped = False Then
        With rsReturnData
            .Filter = "ִ��״̬=" & mState.��ҩ
            
            '���밴����ID��ҩƷID����
            .Sort = "���� Asc ,NO Asc"
            Do While Not .EOF
                strǩ����¼ = ""
                strSQL = "Select id From ҩƷ�շ���¼ Where mod(��¼״̬,3)=2 and no=[1] And ����=[2] And �ⷿid=[3] and �������=[4]"
                Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "", !NO, !����, mcondition.lngҩ��ID, CDate(strDate))
                
                If GetSignatureRecored(EsignTache.returnStep, !����, !NO, mcondition.lngҩ��ID, strǩ����¼, rstemp!Id) = False Then
                    If blnBeginTrans = True Then gcnOracle.RollbackTrans
                    Exit Sub
                End If
                
                If strǩ����¼ <> "" Then
                    strSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption & "-����ǩ��")
                Else
                    gcnOracle.RollbackTrans
                    MsgBox "����ҩ�˵���ǩ��ʧ�ܣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                .MoveNext
            Loop
        End With
    End If
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '��ӡ����
    If bln�Ƿ�����ҩ = True Then
        If mParams.int��ҩ�嵥��ӡ = 2 Then
            If MsgBox("����Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & strDate, "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "��ҩ�ⷿ=" & mcondition.lngҩ��ID, 2)
            End If
        ElseIf mParams.int��ҩ�嵥��ӡ = 1 Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & strDate, "��װϵ��=" & IIf(mParams.strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "��ҩ�ⷿ=" & mcondition.lngҩ��ID, 2)
        End If
        
    Else
        If bln����������ʾ Then
            MsgBox "�������������ļ�¼��������ҩʧ�ܡ�" & vbCrLf & "������ˢ���嵥�󲢺˶ԣ�"
        Else
            MsgBox "����û����ҩ��"
        End If
        
        Exit Sub
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ�����ҩ Then
        mobjPlugIn.DrugReturnByID mcondition.lngҩ��ID, strReturnInfo, CDate(strDate), strReserve
    End If
    
    '��Ϣ��ʾ-��������
    If bln����������ʾ Then
        MsgBox "�������������ļ�¼������ҩƷδ�ɹ���ҩ��" & vbCrLf & "������ˢ���嵥�󲢺˶ԣ�"
    End If
    
    '��ҩ����²����б����ϸ����
    cmdRefreshDept_Click
    
    mfrmDetail.AfterReturnRefresh
    
    Exit Sub
errHandle:
    '����ѿ������񣬲���δ�ύ�������ʱ�ع�����
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub tvwList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        If Mid(tvwList(Index).SelectedItem.Key, 1, 2) = "P_" Then
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_SortPopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub tvwList_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)
    Dim i As Long
    Dim blnAllChecked As Boolean
    Dim blnAllUnChecked As Boolean
     
    Call TvwCheckNode(Node, Node.Checked, True)
    Call TvwSetParentNode(tvwList(Index), Node, Node.Checked)
    
    blnAllChecked = True
    blnAllUnChecked = True
    
    With tvwList(Index)
        For i = 1 To .Nodes.count
            If .Nodes(i).Checked = True Then
                blnAllUnChecked = False
            Else
                blnAllChecked = False
            End If
        Next
    End With
    
    If blnAllChecked = True Then
        chkAll(Index).Value = 1
    ElseIf blnAllUnChecked = True Then
        chkAll(Index).Value = 0
    Else
        chkAll(Index).Value = 2
    End If
End Sub


Private Sub txtInput_Change()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtInput.Text = "" And Me.ActiveControl Is txtInput)
End Sub

Private Sub txtInput_GotFocus()
    Call SelAll(txtInput)

    If Not mobjICCard Is Nothing And txtInput.Text = "" Then
        Call mobjICCard.SetEnabled(True)
    End If
End Sub


Private Sub TxtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         Call TxtInput_Validate(True)
    End If
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    mblnCard = False
    
    If Val(lblPatiInputType.Tag) = mInputType.סԺ�� Or Val(lblPatiInputType.Tag) = mInputType.����ID Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Then Exit Sub
        KeyAscii = 0
    ElseIf Val(lblPatiInputType.Tag) = mInputType.���� Then
        '�������
        mblnCard = zlCommFun.InputIsCard(txtInput, KeyAscii, glngSys)
    End If
    
    If Val(lblPatiInputType.Tag) > 8 Then
        '�����������ѿ�
        If InStr(":��;��?��''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        If Len(txtInput.Text) = txtInput.MaxLength - 1 And KeyAscii <> 8 Then
            txtInput.Text = txtInput.Text & Chr(KeyAscii)
            txtInput.SelStart = Len(txtInput.Text)
            KeyAscii = 0
        End If
        
'        mblnCard = (KeyAscii <> 8 And Len(txtInput.Text) = txtInput.MaxLength - 1 And txtInput.SelLength <> Len(txtInput.Text))
    End If
End Sub

Private Sub txtInput_LostFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub TxtInput_Validate(Cancel As Boolean)
    Dim strDeptInfo As String
    Dim strInput As String
    
    'ȡ�������ƣ����˵�ǰ����������ȡ������¼
    '��ȡ��������Ϣ�󣬷���������ʽ��������Ϣ-��������
    '���������Ϣ�����ؿ���ID����������
    
    mblnInput = False
    
    If InStr(Trim(txtInput.Text), "-") > 0 Then
        'ȡ��-��ǰ���������Ϣ
        strInput = Mid(Trim(txtInput.Text), 1, InStr(Trim(txtInput.Text), "-") - 1)
    Else
        strInput = Trim(txtInput.Text)
    End If
    
    If strInput = "" Then Exit Sub
    
    If Val(lblPatiInputType.Tag) = mInputType.NO Then
        If IsNumeric(strInput) Then
            strInput = GetFullNO(strInput, 14)
        End If
    End If
    
    strDeptInfo = GetPatiInfo(Val(lblPatiInputType.Tag), strInput)
    
    txtInput.Tag = ""
    If strDeptInfo <> "" Then
        Select Case Val(lblPatiInputType.Tag)
        Case mInputType.����
            If mblnCard = True Then
                txtInput.Text = UCase(strInput)
                txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
            Else
                txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
                txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
            End If
'        Case mInputType.���￨
'            txtInput.PasswordChar = ""
'            txtInput.MaxLength = 0
'            txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
'            txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case mInputType.��ҩ��, mInputType.��ҩ��
            txtInput.Text = strDeptInfo
        Case mInputType.��ҩ����
            '���ز���ID����������
            txtInput.Text = Split(strDeptInfo, ",")(1)
            txtInput.Tag = Split(strDeptInfo, ",")(0)
        Case mInputType.IC��
            txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case mInputType.����
            '����ʵ��ͨ������ID����ѯ��������ʾ����-������Ϣ��Tag��¼����ID
            txtInput.Text = strInput & "-" & Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case Else
            If Val(lblPatiInputType.Tag) > 8 Then
                '�������ѿ������ز���ID
                txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
                txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
            Else
                txtInput.Text = strInput & "-" & Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            End If
        End Select
    Else
        txtInput.Tag = 0
    End If
    
    If Val(lblPatiInputType.Tag) <> mInputType.��ҩ���� Then
        mblnInput = True
    End If
        
    'ˢ�²����б�
    DoEvents
    cmdRefreshDept_Click
    
    '�Զ�����Ϊȫѡ������ȡ��ϸ��¼
    If chkAll(IIf(tbcDetail.Selected.Index <> mListType.��ҩ, 0, 1)).Enabled = True Then
        chkAll(IIf(tbcDetail.Selected.Index <> mListType.��ҩ, 0, 1)).Value = 1
        Call chkAll_Click(IIf(tbcDetail.Selected.Index <> mListType.��ҩ, 0, 1))

        DoEvents
        Call cmdRefresh_Click
    End If
    
    tbcDetail.SetFocus
    
    mblnInput = False
    
End Sub

Private Function GetPatiInfo(ByVal intType As Integer, ByVal strInfo As String) As String
    'intType��mInputType����Ŀֵ
    '���ز�����Ϣ����ǰ������ID�Ͳ������ƣ���������Ϣ��ID��������
    '��ʽ��13,һ����|1,����
    Dim rstemp As ADODB.Recordset
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim lngH As Long
    Dim blnCancel As Boolean
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    If intType = mInputType.סԺ�� Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select Nvl(A.��ǰ����id, A.��Ժ����id) As ����ID, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
            " From ������ҳ A, ������Ϣ B, ���ű� C, ������ҳ P " & _
            " Where A.����id = B.����id And A.��ҳid = B.��ҳid and B.����id = P.����id And Nvl(A.��ǰ����id, A.��Ժ����id) = C.ID(+) And P.סԺ�� = [1]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", strInfo)
    ElseIf intType = mInputType.����ID Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select Nvl(A.��ǰ����id, A.��Ժ����id) As ����ID, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
            " From ������ҳ A, ������Ϣ B, ���ű� C " & _
            " Where A.����id = B.����id And A.��ҳid = B.��ҳid And Nvl(A.��ǰ����id, A.��Ժ����id) = C.ID(+) And A.����id = [1]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", Val(strInfo))
    ElseIf intType = mInputType.NO Then
        gstrSQL = "Select Distinct Nvl(A.���˲���id, ���˿���id) As ����ID, B.���� || '-' || B.���� As ��������, A.����id, A.���� As �������� " & _
            " From סԺ���ü�¼ A, ���ű� B " & _
            " Where Nvl(A.���˲���id, ���˿���id) = B.ID(+) And A.NO = [1] And A.�����־=2 And A.ִ�в���id = [2] "
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", strInfo, mcondition.lngҩ��ID)
    ElseIf intType = mInputType.���� Then
        '���ſ��ܲ�Ψһ�������б�ѡ��
        gstrSQL = "Select Rownum As ID, B.���� As ��������, C.���� || '-' || C.���� As ��������, Nvl(A.��ǰ����id, A.��Ժ����id) As ����ID, B.����id " & _
            " From ������ҳ A, ������Ϣ B, ���ű� C " & _
            " Where A.����id = B.����id And A.��ҳid = B.��ҳid And Nvl(A.��ǰ����id, A.��Ժ����id) = C.ID(+) And B.��ǰ���� = [1]"
            
        vRect = GetControlRect(txtInput.hWnd)
        lngH = txtInput.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
        
        Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ȡ������Ϣ", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strInfo)
        If blnCancel = True Then Exit Function
    ElseIf intType = mInputType.���� Then
        If mblnCard = True Then
            gstrSQL = "Select /*+rule*/ Nvl(A.��ǰ����id, A.��Ժ����id) As ����ID, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
                " From ������ҳ A, ������Ϣ B, ���ű� C " & _
                " Where A.����id = B.����id And A.��ҳid = B.��ҳid And Nvl(A.��ǰ����id, A.��Ժ����id) = C.ID(+) And B.���￨�� = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", strInfo)
        Else
            '�������ƿ��ܻ����ظ��������б�ѡ��
            gstrSQL = "Select /*+rule*/ Rownum As ID, ��������, ����ID, ��������, ����id" & _
                " From (Select Distinct B.���� As ��������, B.����id, Nvl(A.��ǰ����id, A.��Ժ����id) As ����ID, C.���� || '-' || C.���� As �������� " & _
                " From ������ҳ A, ������Ϣ B, ���ű� C " & _
                " Where A.����id = B.����id And A.��ҳid = B.��ҳid And Nvl(A.��ǰ����id, A.��Ժ����id) = C.ID(+) And B.���� Like [1])"
            
            vRect = GetControlRect(txtInput.hWnd)
            lngH = txtInput.Height
            sngX = vRect.Left - 15
            sngY = vRect.Top
            
            Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ȡ������Ϣ", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strInfo & "%")
            If blnCancel = True Then Exit Function
        End If
    ElseIf intType = mInputType.��ҩ���� Then
        gstrSQL = " Select ID,����,���� From ���ű� " & _
             " Where ID in (Select ����ID From ��������˵�� Where �������� In ('�ٴ�','���','����','����','����','Ӫ��','����') And ������� IN(2,3))" & _
             " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) And (���� Like [1] Or ���� Like [1] Or ���� Like [2])" & _
             " Order By ����"
        
        vRect = GetControlRect(txtInput.hWnd)
        lngH = txtInput.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
        
        Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ȡ������Ϣ", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, UCase(strInfo) & "%", IIf(gstrMatchMethod = "0", "%", "") & UCase(strInfo) & "%")
        If blnCancel = True Then Exit Function
        
        If rstemp Is Nothing Then Exit Function
        If rstemp.EOF Then Exit Function
        
        GetPatiInfo = rstemp!Id & "," & "[" & rstemp!���� & "]" & rstemp!����
        Exit Function
    ElseIf intType = mInputType.IC�� Then
        If Not mobjSquareCard Is Nothing Then
            'ͨ����ID�Ϳ��Ų��Ҳ���ID
            Call mobjSquareCard.zlGetPatiID("IC��", UCase(txtInput.Text), False, lng����ID)
        End If
        
        If lng����ID > 0 Then
            gstrSQL = "Select Nvl(A.��ǰ����id, A.��Ժ����id) As ����ID, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
                " From ������ҳ A, ������Ϣ B, ���ű� C " & _
                " Where A.����id = B.����id And A.��ҳid = B.��ҳid And Nvl(A.��ǰ����id, A.��Ժ����id) = C.ID(+) And B.����id = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", lng����ID)
        End If
    ElseIf intType > 8 Then
        '���ѿ�
        If Not mobjSquareCard Is Nothing Then
            'ͨ����ID�Ϳ��Ų��Ҳ���ID
            Call mobjSquareCard.zlGetPatiID(Split(Split(mstrCardType, ";")(intType - 9), "|")(3), txtInput.Text, True, lng����ID)
        End If
        
        If lng����ID > 0 Then
            gstrSQL = "Select Nvl(A.��ǰ����id, A.��Ժ����id) As ����ID, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
                " From ������ҳ A, ������Ϣ B, ���ű� C " & _
                " Where A.����id = B.����id And A.��ҳid = B.��ҳid And Nvl(A.��ǰ����id, A.��Ժ����id) = C.ID(+) And A.����id = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", lng����ID)
        End If
    Else
        GetPatiInfo = strInfo
        Exit Function
    End If
    
    If rstemp Is Nothing Then Exit Function
    If rstemp.EOF Then Exit Function
    
    GetPatiInfo = rstemp!����ID & "," & rstemp!�������� & "|" & rstemp!����ID & "," & rstemp!��������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub LoadDept()
    On Error GoTo errHandle
    gstrSQL = "select A.id,A.���� from ���ű� A" & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) B ") & " where A.id=B.Column_Value"
    Set mRsDept = zlDatabase.OpenSQLRecord(gstrSQL, "LoadDept", mParams.strSourceDep)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub




