VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmStationRegist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��վ�Һ�"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   Icon            =   "frmStationRegist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   2940
      Picture         =   "frmStationRegist.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "��������"
      Top             =   600
      Width           =   350
   End
   Begin VB.PictureBox picPayMoney 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   6300
      ScaleHeight     =   360
      ScaleWidth      =   1575
      TabIndex        =   36
      Top             =   4942
      Width           =   1635
      Begin VB.Label lblPayMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   780
         TabIndex        =   37
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.PictureBox picInfo 
      Height          =   2925
      Left            =   15
      ScaleHeight     =   2865
      ScaleWidth      =   7845
      TabIndex        =   30
      Top             =   1950
      Width           =   7905
      Begin VB.CheckBox chkBook 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6435
         TabIndex        =   8
         Top             =   2543
         Width           =   1485
      End
      Begin VB.ComboBox cboDoctor 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   525
         Width           =   3390
      End
      Begin VB.ComboBox cboAppointStyle 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   525
         Width           =   2355
      End
      Begin VB.ComboBox cboArrangeNo 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   45
         Width           =   3390
      End
      Begin VB.ComboBox cboRemark 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         TabIndex        =   7
         Top             =   2490
         Width           =   5430
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   1440
         Left            =   75
         TabIndex        =   31
         Top             =   975
         Width           =   7770
         _cx             =   13705
         _cy             =   2540
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStationRegist.frx":0B14
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   585
         Width           =   480
      End
      Begin VB.Label lblAppointStyle 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4365
         TabIndex        =   35
         Top             =   585
         Width           =   960
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "�ű�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lblLimit 
         AutoSize        =   -1  'True
         Caption         =   "�޺�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4860
         TabIndex        =   33
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   32
         Top             =   2550
         Width           =   480
      End
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   795
      ScaleHeight     =   360
      ScaleWidth      =   1575
      TabIndex        =   28
      Top             =   4950
      Width           =   1635
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   825
         TabIndex        =   29
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   3
      Left            =   -45
      TabIndex        =   24
      Top             =   5430
      Width           =   11000
   End
   Begin VB.ComboBox cboPayMode 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4335
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4935
      Width           =   1950
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   2
      Left            =   -30
      TabIndex        =   19
      Top             =   1440
      Width           =   11000
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   1
      Left            =   -30
      TabIndex        =   18
      Top             =   480
      Width           =   11000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6180
      TabIndex        =   11
      Top             =   5520
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4800
      TabIndex        =   10
      Top             =   5520
      Width           =   1300
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   705
      TabIndex        =   12
      Top             =   5520
      Width           =   1300
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   0
      Left            =   -60
      TabIndex        =   16
      Top             =   960
      Width           =   11000
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   360
      Left            =   705
      TabIndex        =   15
      Top             =   600
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   635
      Appearance      =   2
      IDKindStr       =   "��|��������￨|0|0|0|0|0|;ҽ|ҽ����|0|0|0|0|0|;��|���֤��|1|0|0|0|0|;��|�����|0|0|0|0|0|"
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "����"
      IDKind          =   -1
      DefaultCardType =   "0"
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1290
      TabIndex        =   1
      Top             =   600
      Width           =   1650
   End
   Begin VB.ComboBox cboNO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6060
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "��"
      Height          =   345
      Left            =   7605
      TabIndex        =   27
      Top             =   1568
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   360
      Left            =   5520
      TabIndex        =   3
      Top             =   1560
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   93782018
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picRoom 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5505
      ScaleHeight     =   300
      ScaleWidth      =   2325
      TabIndex        =   43
      Top             =   1560
      Width           =   2385
      Begin VB.Label lblRoomName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   44
         Top             =   15
         Width           =   120
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   675
      TabIndex        =   2
      Top             =   1560
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   93782017
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picDept 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   675
      ScaleHeight     =   300
      ScaleWidth      =   3330
      TabIndex        =   41
      Top             =   1560
      Width           =   3390
      Begin VB.Label lblDeptName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   42
         Top             =   15
         Width           =   120
      End
   End
   Begin VB.Label lbl�� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   135
      TabIndex        =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblPayMode 
      AutoSize        =   -1  'True
      Caption         =   "֧����ʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2970
      TabIndex        =   23
      Top             =   4995
      Width           =   1320
   End
   Begin VB.Label lblSum 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   22
      Top             =   4995
      Width           =   660
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "����Ԥ�����:0.00     "
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3615
      TabIndex        =   17
      Top             =   645
      Width           =   2880
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   14
      Top             =   645
      Width           =   480
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "���ݺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5310
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�:     ����:       �����:              �ѱ�: "
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   135
      TabIndex        =   38
      Top             =   1110
      Width           =   5880
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   25
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4875
      TabIndex        =   26
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4875
      TabIndex        =   20
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   21
      Top             =   1620
      Width           =   480
   End
End
Attribute VB_Name = "frmStationRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset, mblnStartFactUseType As Boolean
Private mblnCard As Boolean, mintSysAppLimit As Integer
Private mfrmPatiInfo As frmPatiInfo
Private mstrYBPati As String, mlng�Һ�ID As Long, mlng����ID As Long
Private mblnOlnyBJYB As Boolean, mblnSharedInvoice As Boolean
Private mstr���� As String, mblnChangeFeeType As Boolean
Private mstrAge As String, mstrFeeType As String, mstrGender As String, mstrClinic As String
Private mstr������� As String
Private mstrPassWord As String, mblnUnload As Boolean, mstrInsure As String
Private mlngDept As Long
Private mblnAppointment As Boolean 'ԤԼ�Һ�
Private Const SNCOLS = 10
Private Const SnArgCols = 7


Private mrsPlan As ADODB.Recordset, mblnInit As Boolean
Private mrsSNState As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsItems As ADODB.Recordset
Private mrsInComes As ADODB.Recordset
Private mrsExpenses As ADODB.Recordset '��¼���ӷ���Ŀ(����������Ϣ)
Private mrsʱ��� As ADODB.Recordset

Public mlngNewPatiID As Long
Private mcolCardPayMode As Collection
Private mcolArrangeNo As Collection
Private mlng����ID As Long, mintIDKind As Integer
Private mcur������� As Currency
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mlngSN As Long
Private mintInsure As Integer
Private mdatLast As Date, mblnNewPati As Boolean
Private mblnChangeByCode As Boolean
Private mstrCardNO As String
Private mcur����͸֧ As Currency
Private mblnAppointPrice As Boolean

'����������Ϣ
Private mstrDef������� As String
Private mstrDef���ʽ As String
Private mstrDef�ѱ� As String

Private Enum EM_REGISTFEE_MODE  '�Һŷ�����ȡ��ʽ
        EM_RG_���� = 0
        EM_RG_���� = 1
        EM_RG_���� = 2
End Enum
Private Enum EM_PATI_CHARGE_MODE    '�����շ�ģʽ
    EM_�Ƚ�������� = 0
    EM_�����ƺ���� = 1
End Enum
Private mRegistFeeMode As EM_REGISTFEE_MODE '�Һŷ�����ȡ��ʽ
Private mPatiChargeMode As EM_PATI_CHARGE_MODE    '�����շ�ģʽ

Private Type TYPE_MedicarePAR
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ʹ�ø����ʻ�   As Boolean  'support�Һ�ʹ�ø����ʻ�
    �����Һ�  As Boolean    'support�����Һ�
    ���ղ����� As Boolean   'support�ҺŲ���ȡ������
End Type
Private MCPAR As TYPE_MedicarePAR

Private Enum ViewMode
     V_��ͨ��
     v_ר�Һ�
     v_ר�Һŷ�ʱ��
     V_��ͨ�ŷ�ʱ��
End Enum
Private mViewMode As ViewMode

Private Type ty_ModulePara
    bln����ģ������ As Boolean
    lng������������ As Long
    blnĬ�Ϲ����� As Boolean
    blnĬ������ժҪ As Boolean
    byt�Һ�ģʽ As Byte
    bln�Һű���ˢ�� As Boolean
    bln����ʹ��Ԥ�� As Boolean
    blnסԺ���˹Һ� As Boolean
    bln�������Ұ��� As Boolean
    int�Һŷ�Ʊ��ӡ As Integer
    int�Һ�ƾ����ӡ As Integer
    intԤԼ�ҺŴ�ӡ As Integer
    bln������ѡ�� As Boolean
    lngԤԼ��Чʱ�� As Long
    bln�����շ�Ʊ�� As Boolean
    bln�˺����� As Boolean
    blnԤԼʱ�տ� As Boolean
    bln������֤ As Boolean
    bln����ҽ�� As Boolean
End Type

Private mty_Para As ty_ModulePara

Public Sub zlShowMe(ByVal frmMain As Object, ByVal lngModul As Long, ByVal strDeptIDs As String, _
                    ByVal blnAppointment As Boolean, ByVal lng����ID As Long, ByRef strOutNO As String)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    mblnAppointment = blnAppointment
    mlngModul = lngModul
    mlng����ID = lng����ID
    If frmMain Is Nothing Then
        Me.Show
    Else
        Me.Show 1, frmMain
    End If
    If mblnOK = True Then
        strOutNO = mstrNO
        Unload Me
    End If
End Sub

Private Sub InitData()
    '��ʼ�����õĻ�������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo Errhand
    strSQL = "Select ����, ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName)
    If Not rsTmp.EOF Then
        mstrDef������� = Nvl(rsTmp!����)
        mstrDef���ʽ = Nvl(rsTmp!����)
    End If
    
    strSQL = "Select ���� From �ѱ� Where ȱʡ��־ = 1"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        mstrDef�ѱ� = Nvl(rsTmp!����)
    End If
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub InitPara()
    Dim strValue As String
    On Error GoTo errH
    With mty_Para
        .bln����ģ������ = Val(gobjDatabase.GetPara("����ģ������", glngSys, 9000, "0")) = 1
        .lng������������ = Val(gobjDatabase.GetPara("������������", glngSys, 9000, 0))
        .blnĬ�Ϲ����� = Val(gobjDatabase.GetPara("Ĭ�Ϲ�����", glngSys, 9000, "0")) = 1
        .blnĬ������ժҪ = Val(gobjDatabase.GetPara("Ĭ������ժҪ", glngSys, 9000, "1")) = 1
        .byt�Һ�ģʽ = 0
        .bln����ʹ��Ԥ�� = Val(gobjDatabase.GetPara("����ʹ��Ԥ��", glngSys, 9000, "0")) = 1
        .blnסԺ���˹Һ� = Val(gobjDatabase.GetPara("����סԺ���˹Һ�", glngSys, 9000, "0")) = 1
        .int�Һŷ�Ʊ��ӡ = Val(gobjDatabase.GetPara("�Һŷ�Ʊ��ӡ��ʽ", glngSys, 9000, "0"))
        .int�Һ�ƾ����ӡ = Val(gobjDatabase.GetPara("�Һ�ƾ����ӡ��ʽ", glngSys, 9000, "0"))
        .intԤԼ�ҺŴ�ӡ = Val(gobjDatabase.GetPara("ԤԼ�Һŵ���ӡ��ʽ", glngSys, 9000, "0"))
        .bln������ѡ�� = Val(gobjDatabase.GetPara("������ѡ��", glngSys, 9000, "0")) = 1
        .bln�����շ�Ʊ�� = Val(gobjDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121)) = 1
        .bln�˺����� = Val(gobjDatabase.GetPara("�����������Һ�", glngSys, 1111)) = 1
        .blnԤԼʱ�տ� = Val(gobjDatabase.GetPara("ԤԼʱ�տ�", glngSys, 9000, "0")) = 1
        .bln�������Ұ��� = Val(gobjDatabase.GetPara("�������Ұ���", glngSys, 9000, "0")) = 1
        .bln�Һű���ˢ�� = Val(gobjDatabase.GetPara("�Һű���ˢ��", glngSys, 9000)) = 1
        .bln������֤ = Val(gobjDatabase.GetPara(28, glngSys)) <> 0
        .bln����ҽ�� = Val(gobjDatabase.GetPara("����ҽ��", glngSys, 9000)) = 1
        If .blnĬ������ժҪ Then
            cboRemark.TabStop = True
        Else
            cboRemark.TabStop = False
        End If
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_����
        Else
            If .byt�Һ�ģʽ = 0 Then
                mRegistFeeMode = EM_RG_����
            Else
                mRegistFeeMode = EM_RG_����
            End If
        End If
    End With
    'ˢ��Ҫ����������
    mstrCardPass = gobjDatabase.GetPara(46, glngSys, , "0000000000")
    Call gobjControl.PicShowFlat(picInfo)
    '�շѺ͹ҺŹ���Ʊ��
    mblnSharedInvoice = gobjDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121) = "1"
    '���ع��ùҺ�����ID
    If mblnSharedInvoice Then
        mlng�Һ�ID = Val(gobjDatabase.GetPara("�����շ�Ʊ������", glngSys, 1121, ""))
    Else
        mlng�Һ�ID = Val(gobjDatabase.GetPara("���ùҺ�Ʊ������", glngSys, mlngModul, ""))
    End If
    mlngDept = Val(gobjDatabase.GetPara("�������", glngSys, 1260, ""))
    If mlng�Һ�ID > 0 Then
        If Not ExistShareBill(mlng�Һ�ID, IIf(mblnSharedInvoice, 1, 4)) Then
            If mblnSharedInvoice Then
                gobjDatabase.SetPara "�����շ�Ʊ������", "0", glngSys, 1121
            Else
                gobjDatabase.SetPara "���ùҺ�Ʊ������", "0", glngSys, mlngModul
            End If
            mlng�Һ�ID = 0
        End If
    End If
    'Ʊ���ϸ����
    strValue = gobjDatabase.GetPara(24, glngSys, , "00000")
    gblnBill�Һ� = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    mintSysAppLimit = Val(gobjDatabase.GetPara("�Һ�����ԤԼ����", glngSys))
    If mblnSharedInvoice Then
        '�Һ�������Ʊ��:42703
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function RefreshFact(Optional ByRef strFact As String) As Boolean
'������blnNew=�Ƿ��µ�����ʱ����,��ʱ���ڷ��ϸ���Ƶ�Ʊ���Ǳ��浱ǰ��
    Dim strUseType As String
    If mblnStartFactUseType Then
        strUseType = zl_GetInvoiceUserType(Val(mrsInfo!����ID), 0, mintInsure)
    End If
    If gblnBill�Һ� Then
        mlng����ID = CheckUsedBill(IIf(mblnSharedInvoice, 1, 4), IIf(mlng����ID > 0, mlng����ID, mlng�Һ�ID), , strUseType)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õĹҺ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            strFact = "": Exit Function
        Else
            '�ϸ�ȡ��һ������
            strFact = GetNextBill(mlng����ID)
        End If
    Else
        If mblnSharedInvoice Then
            strFact = gobjDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121)
        Else
            strFact = gobjDatabase.GetPara("��ǰ�Һ�Ʊ�ݺ�", glngSys, 1111)
        End If
        strFact = IncStr(strFact)
        If mblnSharedInvoice Then
            gobjDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", strFact, glngSys, 1121
        Else
            gobjDatabase.SetPara "��ǰ�Һ�Ʊ�ݺ�", strFact, glngSys, 1111
        End If
    End If
    RefreshFact = True
End Function

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long, strTemp As String
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "��|����|0;ҽ|ҽ����|0;��|���֤��|0;��|�����|0", txtPatient)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If

    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function

Private Sub cboAppointStyle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboArrangeNo_Click()
    Call ReadLimit
    Call LoadDoctor
    Call LoadFeeItem(Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1)
    Call GetActiveView
    If mblnAppointment Then
        Select Case mViewMode
            Case V_��ͨ�ŷ�ʱ��, v_ר�Һŷ�ʱ��
                cmdTime.Visible = True
            Case Else
                cmdTime.Visible = False
        End Select
        Call InitRegTime
    Else
        cmdTime.Visible = False
    End If
    lblDeptName.Caption = Nvl(mrsPlan!����)
End Sub

Private Sub InitRegTime()
    Dim dateCur As Date, strNO As String, strDay As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, rsTime As ADODB.Recordset
    On Error GoTo errH
    strDay = zlGet��ǰ���ڼ�(dtpDate.Value)
    strSQL = "Select ʱ���,��ʼʱ��,ȱʡʱ�� From ʱ��� Where ���� Is Null And վ�� Is Null"
    Set rsTime = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If IsNull(mrsPlan.Fields(strDay).Value) Then
        If Format(dtpDate.Value, "yyyy-mm-dd") = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") Then
            '���첻����,ȡ��ǰʱ��
            dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
        Else
            'δ��������,ȡĬ��ʱ��
            rsTime.Filter = "ʱ���='����'"
            If rsTime.EOF Then
                dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
            Else
                If IsNull(rsTime!ȱʡʱ��) Then
                    dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
                Else
                    dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
                End If
            End If
        End If
    Else
        Select Case mViewMode
            Case V_��ͨ�ŷ�ʱ��, v_ר�Һŷ�ʱ��
            strSQL = "Select Distinct a.��� As ID, To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
                    "From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
                    "Where a.����id = b.Id And b.���� = [1] And" & vbNewLine & _
                    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                    "      Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����'," & vbNewLine & _
                    "             Null) = a.����(+) And Not Exists" & vbNewLine & _
                    " (Select Count(1)" & vbNewLine & _
                    "       From �Һ����״̬" & vbNewLine & _
                    "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having" & vbNewLine & _
                    "        Count(1) - a.�������� >= 0) And Not Exists" & vbNewLine & _
                    " (Select 1" & vbNewLine & _
                    "       From �ҺŰ��żƻ� E" & vbNewLine & _
                    "       Where e.����id = b.Id And e.���ʱ�� Is Not Null And" & vbNewLine & _
                    "             [2] Between Nvl(e.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                    "             Nvl(e.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')))"
            strSQL = strSQL & " Union " & _
                    "Select Distinct a.��� As ID, To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
                    "From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C," & vbNewLine & _
                    "     (Select Max(a.��Чʱ��) ��Ч" & vbNewLine & _
                    "       From �ҺŰ��żƻ� A, �ҺŰ��� B" & vbNewLine & _
                    "       Where a.����id = b.Id And b.���� = [1] And a.���ʱ�� Is Not Null And" & vbNewLine & _
                    "             [2] Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                    "             Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'))) D" & vbNewLine & _
                    "Where a.�ƻ�id = b.Id And b.����id = c.Id And c.���� = [1] And b.��Чʱ�� = d.��Ч And b.���ʱ�� Is Not Null And" & vbNewLine & _
                    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                    "      [2] Between Nvl(b.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                    "      Nvl(b.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And Not Exists" & vbNewLine & _
                    " (Select Count(1)" & vbNewLine & _
                    "       From �Һ����״̬" & vbNewLine & _
                    "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having" & vbNewLine & _
                    "        Count(1) - a.�������� >= 0) And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5'," & vbNewLine & _
                    "                                           '����', '6', '����', '7', '����', Null) = a.����(+)" & vbNewLine & _
                    "Order By ��ʼʱ��"
        
            dateCur = Format(dtpDate, "yyyy-mm-dd")
            strNO = Get�ű�
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, dateCur)
            If Not rsTmp.EOF Then
                'ʱ�ε�����ʱ��,ȡ��Сʱ��
                dtpTime.Value = Format(Nvl(rsTmp!��ʼʱ��), "hh:mm:ss")
            Else
                If Format(dtpDate.Value, "yyyy-mm-dd") = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") Then
                    dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                Else
                    'ʱ�ε�����ʱ��,ȡ��ʼʱ��
                    rsTime.Filter = "ʱ���='" & Nvl(mrsPlan.Fields(strDay).Value) & "'"
                    If rsTime.EOF Then
                        dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                    Else
                        If IsNull(rsTime!ȱʡʱ��) Then
                            dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
                        Else
                            dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
                        End If
                    End If
                End If
            End If
            Case Else
                If Format(dtpDate.Value, "yyyy-mm-dd") = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") Then
                    dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                Else
                    '������ʱ��,ȡ��ʼʱ��
                    rsTime.Filter = "ʱ���='" & Nvl(mrsPlan.Fields(strDay).Value) & "'"
                    If rsTime.EOF Then
                        dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                    Else
                        If IsNull(rsTime!ȱʡʱ��) Then
                            dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
                        Else
                            dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub GetAllҽ��()
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.����, Upper(a.����) As ����,b.����id,a.���" & _
            " From ��Ա�� a, ������Ա b, ��Ա����˵�� c" & _
            " Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order By a.���� Desc"
    Set mrsDoctor = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, "ҽ��")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub cboArrangeNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub LoadDoctor()
    With cboDoctor
        .Clear
        If Nvl(mrsPlan!ҽ��) = "" Then
            If mty_Para.bln����ҽ�� Then
                mrsDoctor.Filter = "����id=" & Val(Nvl(mrsPlan!����ID))
                
                Do While Not mrsDoctor.EOF
                    .AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    If Nvl(mrsDoctor!����) = UserInfo.���� Then .ListIndex = .NewIndex
                    mrsDoctor.MoveNext
                Loop
                If .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                .Enabled = True
                lblDoctor.Enabled = True
            Else
                mrsDoctor.Filter = "����='" & UserInfo.���� & "'"
                If mrsDoctor.RecordCount <> 0 Then
                    .AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    .ListIndex = 0
                End If
                .Enabled = False
                lblDoctor.Enabled = False
            End If
        Else
            mrsDoctor.Filter = "����='" & Nvl(mrsPlan!ҽ��) & "'"
            If mrsDoctor.RecordCount <> 0 Then
                .AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
                .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                .ListIndex = 0
            End If
            .Enabled = False
            lblDoctor.Enabled = False
        End If
    End With
End Sub

Private Sub cboPayMode_Click()
    If MCPAR.���ղ����� And cboPayMode.Text = mstrInsure Then
        chkBook.Enabled = False
        chkBook.Value = 0
    Else
        chkBook.Enabled = True
    End If
End Sub

Private Sub cboPayMode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboRemark_Change()
    cboRemark.Tag = ""
End Sub

Private Sub cboRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRemark.Tag <> "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(cboRemark.Text) = "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If SelectMemo(Trim(cboRemark.Text)) = False Then
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
End Sub

Private Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String

    If Val(gobjDatabase.GetPara("����ƥ��")) = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Private Function SelectMemo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ����ժҪ
    '���:strInput-���봮;Ϊ��ʱ,��ʾȫ��
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If gobjCommFun.IsCharChinese(cboRemark.Text) Then
             strWhere = " And  ���� like [1] "
        ElseIf gobjCommFun.IsNumOrChar(cboRemark.Text) Then
             strWhere = " And (���� like upper([1]) or ���� like upper([1]))"
        End If
    End If
    
    strSQL = "" & _
     "   Select RowNum AS ID,����,����,����  " & _
     "   From ���ùҺ�ժҪ " & _
     "   Where 1=1 " & strWhere & _
     "   Order by ȱʡ��־"
     vRect = GetControlRect(cboRemark.hWnd)
     On Error GoTo Hd
     Set rsInfo = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "���ùҺ�ժҪ", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cboRemark.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "û�����ó��ùҺ�ժҪ,�����ֵ����������", vbOKOnly + vbInformation, gstrSysName
        End If
        gobjCommFun.PressKey vbKeyTab: Exit Function
     End If
     gobjControl.CboSetText Me.cboRemark, Nvl(rsInfo!����)
     cboRemark.Tag = Nvl(rsInfo!����)
     gobjCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then Resume
    gobjComlib.SaveErrLog
End Function

Private Sub chkBook_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cmdNewPati_Click()
    If Not mrsInfo Is Nothing Then
        If Val(Nvl(mrsInfo!����ID)) <> 0 Then
            Call ViewPatiInfo
            Exit Sub
        End If
    End If
    Call CreateNewPati
End Sub

Private Sub ViewPatiInfo()
    '�鿴������Ϣ
    Dim bln���� As Boolean, lng����id As Long
    On Error GoTo errH
    If mrsPlan Is Nothing Then
        lng����id = 0
    Else
        If mrsPlan.RecordCount = 0 Then
            lng����id = 0
        Else
            lng����id = Val(Nvl(mrsPlan!����ID))
        End If
    End If
    
    bln���� = Check����(Val(Nvl(mrsInfo!����ID)), lng����id)
    With mfrmPatiInfo
        Set .mfrmMain = Me
        .mbytFun = 0
        .mlng����ID = Val(Nvl(mrsInfo!����ID))
        .mbln���� = bln����
        .Show 1, Me
    End With
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CreateNewPati()
    '�½�������Ϣ
    On Error GoTo errH

    With mfrmPatiInfo
        Set .mfrmMain = Me
        .mbytFun = 2
        .Show 1, Me
        If mlngNewPatiID <> 0 Then mblnNewPati = True
    End With
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdTime_Click()
    If InitTimePlan = False Then Exit Sub
    If mrsʱ���.RecordCount <> 0 Then
        dtpTime.Value = Format(mrsʱ���!��ʼʱ��, "hh:mm:ss")
    End If
End Sub

Private Sub dtpDate_Change()
    Call LoadRegPlans(False)
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub dtpTime_GotFocus()
    Call cmdTime_Click
End Sub

Private Sub dtpTime_Validate(Cancel As Boolean)
    If Format(dtpDate.Value, "YYYY-MM-DD") = Format(gobjDatabase.CurrentDate, "YYYY-MM-DD") Then
        If Format(dtpTime.Value, "hh:mm:ss") < Format(gobjDatabase.CurrentDate, "hh:mm:ss") Then
            MsgBox "ԤԼʱ�䲻��С�ڵ�ǰʱ��!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnload Then mblnUnload = False: Unload Me: Exit Sub
    If mblnInit And Not mrsInfo Is Nothing Then
        If cboArrangeNo.Enabled And cboArrangeNo.Visible Then cboArrangeNo.SetFocus
        If cboArrangeNo.ListCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
    End If
    mblnInit = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsInfo = Nothing
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    Set mrsExpenses = Nothing
    
    mstrDef������� = ""
    mstrDef���ʽ = ""
    mstrDef�ѱ� = ""
    mstrCardNO = ""
    If Not mobjIDCard Is Nothing Then
         Call mobjIDCard.SetEnabled(False)
         Set mobjIDCard = Nothing
     End If
     If Not mobjICCard Is Nothing Then
         Call mobjICCard.SetEnabled(False)
         Set mobjICCard = Nothing
     End If
     mintIDKind = IDKind.IDKind
     Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strExpand As String
    Dim strOutCardNO As String, strOutPatiInforXML As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        'ϵͳIC��
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(objCard, txtPatient.Text, True)
            End If
        End If
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    
    If txtPatient.Text <> "" Then
        Call GetPatient(objCard, txtPatient.Text, True)
    End If
End Sub

Private Sub txtRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub chkBook_Click()
    If mrsPlan Is Nothing Then Exit Sub
    Call LoadFeeItem(Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1)
End Sub

Private Sub cmdCancel_Click()
    If txtPatient.Text <> "" Then
        If MsgBox("�Ƿ���յ�ǰ������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ClearPatient
        End If
        Exit Sub
    End If
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp "zl9RegEvent", Me.hWnd, "frmRegistEdit"
    Exit Sub
End Sub

Private Function CheckBrushCard(ByVal dblMoney As Double, ByVal lngҽ�ƿ����ID As Long, ByVal bln���ѿ� As Boolean, _
                                ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, _
                                ByVal rsExpenses As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str���� As String
    On Error GoTo errHandle
    '68991
    If mRegistFeeMode <> EM_RG_���� Then CheckBrushCard = True: Exit Function
    If dblMoney = 0 Then
        CheckBrushCard = True: Exit Function
    End If
    If Not (cboPayMode.Visible And cboPayMode.Enabled) Then
        CheckBrushCard = True: Exit Function
    End If
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then
        CheckBrushCard = True: Exit Function
    End If
    If lngҽ�ƿ����ID = 0 Then
        MsgBox cboPayMode.Text & "�쳣,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "ʹ��" & cboPayMode.Text & "֧�������ȳ�ʼ���ӿڲ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call zlGetClassMoney(rsMoney, rsItems, rsIncomes, rsExpenses)
    
     '����ˢ������
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln���ѿ� As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl��� As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String) As Boolean
    str���� = Trim(mstrAge)

   If gobjSquare.objSquareCard.zlBrushCard(Me, glngModul, rsMoney, lngҽ�ƿ����ID, bln���ѿ�, _
    txtPatient.Text, NeedName(mstrGender), str����, dblMoney, mstrCardNO, mstrPassWord, _
    False, True, False, True, Nothing, False, True) = False Then Exit Function
    
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, glngModul, lngҽ�ƿ����ID, _
        bln���ѿ�, mstrCardNO, dblMoney, "", "") = False Then Exit Function

    CheckBrushCard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, _
                                ByVal rsExpenses As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    
    Err = 0: On Error GoTo Errhand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic

        rsItems.Filter = 0
        If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
        Do While Not rsItems.EOF
            rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
            rsMoney.Filter = "�շ����='" & Nvl(rsItems!���, "��") & "'"
            If rsMoney.EOF Then
                .AddNew
            End If
            !�շ���� = Nvl(rsItems!���, "��")
            !��� = Val(Nvl(!���)) + Val(Nvl(rsIncomes!ʵ��))
            .Update
            rsItems.MoveNext
        Loop
        
        If Not rsExpenses Is Nothing Then
            If rsExpenses.RecordCount > 0 Then rsExpenses.MoveFirst
            Do While Not rsExpenses.EOF
                rsMoney.Filter = "�շ����='" & Nvl(rsExpenses!���, "��") & "'"
                If rsMoney.EOF Then
                    .AddNew
                End If
                !�շ���� = Nvl(rsExpenses!���, "��")
                !��� = Val(Nvl(!���)) + Val(Nvl(rsExpenses!ʵ��))
                .Update
                rsExpenses.MoveNext
            Loop
        End If
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim blnSlipPrint As Boolean, blnInvoicePrint As Boolean, int�۸񸸺� As Integer, blnBalance As Boolean
    Dim k As Integer, i As Integer, j As Integer, strNO As String, strFactNO As String
    Dim strSQL As String, str�Ǽ�ʱ�� As String, str����ʱ�� As String
    Dim curԤ�� As Currency, cur���� As Currency, cur�ֽ� As Currency, str����NO As String
    Dim lngSN As Long, lng�Һſ���ID As Long, lng����ID As Long, byt���� As Byte
    Dim lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean, blnNoDoc As Boolean, strBalanceStyle As String
    Dim blnTrans As Boolean, blnNotCommit As Boolean, strAdvance As String, lng����ID As Long
    Dim lngҽ��ID As Long, blnOneCard As Boolean, rsTmp As ADODB.Recordset
    Dim strDay As String, blnAppointPrint As Boolean, strNotValiedNos As String
    Dim rs���ʽ As ADODB.Recordset, strҽ�� As String, blnAdd As Boolean
    Dim cllPro As New Collection, cllCardPro As Collection, cllTheeSwap As Collection, cllProAfter As New Collection
    
    If CheckValied = False Then Exit Sub
    
    strSQL = "Select ���,����,ҽԺ����,���㷽ʽ From һ��ͨĿ¼ Where ���� = 1 And ���㷽ʽ = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, cboPayMode.Text)
    blnOneCard = rsTmp.RecordCount <> 0
    
    If mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then
        blnSlipPrint = False
    Else
        Select Case Val(mty_Para.int�Һ�ƾ����ӡ)
            Case 0    '����ӡ
                blnSlipPrint = False
            Case 1    '�Զ���ӡ
                If InStr(gstrPrivs, ";���˹Һ�ƾ��;") > 0 Then
                    blnSlipPrint = True
                Else
                    blnSlipPrint = False
                    MsgBox "��û�йҺ�ƾ����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                End If
            Case 2    'ѡ���ӡ
                If MsgBox("Ҫ��ӡ�Һ�ƾ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If InStr(gstrPrivs, ";���˹Һ�ƾ��;") > 0 Then
                        blnSlipPrint = True
                    Else
                        blnSlipPrint = False
                        MsgBox "��û�йҺ�ƾ����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                    End If
                Else
                    blnSlipPrint = False
                End If
        End Select
    End If
    
    If mRegistFeeMode = EM_RG_���� Or mRegistFeeMode = EM_RG_���� Or (mblnAppointment And mty_Para.blnԤԼʱ�տ� = False) Then
        blnInvoicePrint = False
    Else
        If Not (mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
            Select Case Val(mty_Para.int�Һŷ�Ʊ��ӡ)
                Case 0    '����ӡ
                    blnInvoicePrint = False
                Case 1    '�Զ���ӡ
                    If InStr(gstrPrivs, ";�Һŷ�Ʊ��ӡ;") > 0 Then
                        blnInvoicePrint = True
                    Else
                        blnInvoicePrint = False
                        MsgBox "��û�йҺŷ�Ʊ��ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                    End If
                Case 2    'ѡ���ӡ
                    If MsgBox("Ҫ��ӡ�Һŷ�Ʊ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        If InStr(gstrPrivs, ";�Һŷ�Ʊ��ӡ;") > 0 Then
                            blnInvoicePrint = True
                        Else
                            blnInvoicePrint = False
                            MsgBox "��û�йҺŷ�Ʊ��ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                        End If
                    Else
                        blnInvoicePrint = False
                    End If
            End Select
        End If
    End If
    
    If mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then
        Select Case Val(mty_Para.intԤԼ�ҺŴ�ӡ)
            Case 0
                blnAppointPrint = False
            Case 1
                If InStr(gstrPrivs, ";ԤԼ�Һŵ�;") > 0 Then
                    blnAppointPrint = True
                Else
                    blnAppointPrint = False
                    MsgBox "��û��ԤԼ�Һŵ���ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                End If
            Case 2
                If InStr(gstrPrivs, ";ԤԼ�Һŵ�;") > 0 Then
                    If MsgBox("Ҫ��ӡԤԼ�Һŵ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        blnAppointPrint = True
                    Else
                        blnAppointPrint = False
                    End If
                Else
                    MsgBox "��û��ԤԼ�Һŵ���ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                    blnAppointPrint = False
                End If
        End Select
    Else
        blnAppointPrint = False
    End If
    
    If blnInvoicePrint Or (mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
        If RefreshFact(strFactNO) = False Then Exit Sub
    End If
    
    If mblnAppointment Then
        If mRegistFeeMode = EM_RG_���� And mty_Para.blnԤԼʱ�տ� Then
            MsgBox "��֧�������ƺ���㲡�˵�ԤԼ�տ�Һţ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If mty_Para.blnԤԼʱ�տ� Then
            If Not mRegistFeeMode = EM_RG_���� Then
                If cboPayMode.Text = "Ԥ����" Then
                    curԤ�� = Val(lblTotal.Caption)
                Else
                    If cboPayMode.Text = mstrInsure Then
                        cur���� = Val(lblTotal.Caption)
                    Else
                        blnBalance = True
                        cur�ֽ� = Val(lblTotal.Caption)
                    End If
                End If
            End If
        Else
            blnBalance = False
        End If
    Else
        If Not mRegistFeeMode = EM_RG_���� Then
            If cboPayMode.Text = "Ԥ����" Then
                curԤ�� = Val(lblTotal.Caption)
            Else
                If cboPayMode.Text = "�����ʻ�" Then
                    cur���� = Val(lblTotal.Caption)
                Else
                    blnBalance = True
                    cur�ֽ� = Val(lblTotal.Caption)
                End If
            End If
        End If
    End If
    
    If frmPatiInfo.SaveAfterArrList(mblnNewPati, lng����ID) = False Then
        MsgBox "���没����Ϣʧ�ܣ����飡", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnNewPati = False
    mlngNewPatiID = 0
    txtPatient.Text = "-" & lng����ID
    GetPatient IDKind.GetCurCard, txtPatient.Text, False
    
    If Val(curԤ��) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!����ID), Val(curԤ��), mlngModul, 1, , mty_Para.bln������֤) Then Exit Sub
    End If
    
    If mblnAppointment = False Or (mblnAppointment = True And mty_Para.blnԤԼʱ�տ�) Then
        If zlIsAllowPatiChargeFeeMode(ZVal(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!����ģʽ))) = False Then Exit Sub
    End If

    '126802:���ϴ�,2018/6/8,�����һ��֧������
    mstrCardNO = ""
    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lngҽ�ƿ����ID = mcolCardPayMode.Item(i)(3)
                bln���ѿ� = Val(mcolCardPayMode.Item(i)(5)) = 1
                strBalanceStyle = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur�ֽ�), lngҽ�ƿ����ID, bln���ѿ�, mrsItems, mrsInComes, mrsExpenses) = False Then Exit Sub
    End If
    
    str�Ǽ�ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
    
    If mblnAppointment Then
        strDay = zlGet��ǰ���ڼ�(dtpDate.Value)
    Else
        strDay = zlGet��ǰ���ڼ�
    End If
    
    '��ȡ����ʱ��
    blnAdd = False
    If mblnAppointment Then
        mlngSN = 0
        str����ʱ�� = "To_Date('" & Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
        If mViewMode = v_ר�Һŷ�ʱ�� Then
            If Val(Nvl(mrsPlan!�ƻ�ID)) <> 0 Then
                strSQL = "Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
                        "       Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
                        "From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����," & vbNewLine & _
                        "              To_Date(To_Char(" & str����ʱ�� & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
                        "              To_Date(To_Char(" & str����ʱ�� & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
                        "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
                        "       From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd" & vbNewLine & _
                        "       Where Jh.Id = Sd.�ƻ�id And Jh.Id = [1] And" & vbNewLine & _
                        "             Sd.���� =" & vbNewLine & _
                        "             Decode(To_Char(" & str����ʱ�� & ", 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)) Jh," & vbNewLine & _
                        "     �Һ����״̬ Zt" & vbNewLine & _
                        "Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Jh.��ʼʱ�� = " & str����ʱ�� & " And Zt.���(+) = Jh.��� And Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
                        "Order By ���"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!�ƻ�ID)))
                If rsTmp.RecordCount <> 0 Then
                    mlngSN = Val(Nvl(rsTmp!���))
                Else
                    strSQL = "Select Max(���) As ��� From �Һ����״̬ Where ���� = [1] And Trunc(����) = Trunc(" & str����ʱ�� & ")"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(mrsPlan!�ű�))
                    If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!���))
                    strSQL = "Select Max(���) As ��� From �Һżƻ�ʱ�� Where �ƻ�ID = [1] And ���� = Decode(To_Char(" & str����ʱ�� & ", 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!�ƻ�ID)))
                    If mlngSN = 0 Then
                        If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!���))
                    Else
                        If Val(Nvl(rsTmp!���)) > mlngSN Then mlngSN = Val(Nvl(rsTmp!���))
                    End If
                    mlngSN = mlngSN + 1
                    blnAdd = True
                End If
            Else
                strSQL = "Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
                        "       Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
                        "From (Select Sd.����id, Sd.���, Sd.����, Ap.����," & vbNewLine & _
                        "              To_Date(To_Char(" & str����ʱ�� & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
                        "              To_Date(To_Char(" & str����ʱ�� & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
                        "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
                        "       From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd" & vbNewLine & _
                        "       Where Ap.Id = Sd.����id And Ap.Id = [1] And" & vbNewLine & _
                        "             Sd.���� = Decode(To_Char(" & str����ʱ�� & ", 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7'," & vbNewLine & _
                        "                            '����', Null)) Ap, �Һ����״̬ Zt" & vbNewLine & _
                        "Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Ap.��ʼʱ�� = " & str����ʱ�� & " And Zt.���(+) = Ap.��� And Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
                        "Order By ���"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!ID)))
                If rsTmp.RecordCount <> 0 Then
                    mlngSN = Val(Nvl(rsTmp!���))
                Else
                    strSQL = "Select Max(���) As ��� From �Һ����״̬ Where ���� = [1] And Trunc(����) = Trunc(" & str����ʱ�� & ")"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(mrsPlan!�ű�))
                    If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!���))
                    strSQL = "Select Max(���) As ��� From �ҺŰ���ʱ�� Where ����ID = [1] And ���� = Decode(To_Char(" & str����ʱ�� & ", 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!ID)))
                    If mlngSN = 0 Then
                        If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!���))
                    Else
                        If Val(Nvl(rsTmp!���)) > mlngSN Then mlngSN = Val(Nvl(rsTmp!���))
                    End If
                    mlngSN = mlngSN + 1
                    blnAdd = True
                End If
            End If
        End If
        If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
    Else
        Select Case mViewMode
            Case V_��ͨ��
                str����ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
            Case V_��ͨ�ŷ�ʱ��
                str����ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
            Case v_ר�Һ�
                str����ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
            Case v_ר�Һŷ�ʱ��
                If IsNull(mrsPlan.Fields(strDay).Value) Then
                    str����ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                    blnAdd = True
                Else
                    strSQL = "Select 1" & vbNewLine & _
                            "From ʱ���" & vbNewLine & _
                            "Where ���� Is Null And վ�� Is Null And (('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between" & vbNewLine & _
                            "      Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-09 ' || To_Char(Nvl(��ǰʱ��, ��ʼʱ��), 'HH24:MI:SS')," & vbNewLine & _
                            "               '3000-01-10 ' || To_Char(Nvl(��ǰʱ��, ��ʼʱ��), 'HH24:MI:SS')) And" & vbNewLine & _
                            "      '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')) Or" & vbNewLine & _
                            "      ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between '3000-01-10 ' || To_Char(Nvl(��ǰʱ��, ��ʼʱ��), 'HH24:MI:SS') And" & vbNewLine & _
                            "      Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')," & vbNewLine & _
                            "               '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')))) And ʱ��� = [1]"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsPlan.Fields(strDay).Value)
                    '������
                    If rsTmp.RecordCount = 0 Then
                        str����ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                        blnAdd = True
                    Else
                        'ȡ��С����ʱ���
                        If Val(Nvl(mrsPlan!�ƻ�ID)) <> 0 Then
                            strSQL = "Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
                                    "       Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
                                    "From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
                                    "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
                                    "       From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd" & vbNewLine & _
                                    "       Where Jh.Id = Sd.�ƻ�id And Jh.Id = [1] And" & vbNewLine & _
                                    "             Sd.���� =" & vbNewLine & _
                                    "             Decode(To_Char(Sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)) Jh," & vbNewLine & _
                                    "     �Һ����״̬ Zt" & vbNewLine & _
                                    "Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
                                    "Order By ���"
                            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!�ƻ�ID)))
                            If rsTmp.RecordCount <> 0 Then
                                mlngSN = Val(Nvl(rsTmp!���))
                                str����ʱ�� = "To_Date('" & Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") & " " & Format(Nvl(rsTmp!��ʼʱ��), "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                            Else
                                str����ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                            End If
                        Else
                            strSQL = "Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
                                    "       Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
                                    "From (Select Sd.����id, Sd.���, Sd.����, Ap.����," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
                                    "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
                                    "       From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd" & vbNewLine & _
                                    "       Where Ap.Id = Sd.����id And Ap.Id = [1] And" & vbNewLine & _
                                    "             Sd.���� = Decode(To_Char(Sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7'," & vbNewLine & _
                                    "                            '����', Null)) Ap, �Һ����״̬ Zt" & vbNewLine & _
                                    "Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
                                    "Order By ���"
    
                            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!ID)))
                            If rsTmp.RecordCount <> 0 Then
                                mlngSN = Val(Nvl(rsTmp!���))
                                str����ʱ�� = "To_Date('" & Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") & " " & Format(Nvl(rsTmp!��ʼʱ��), "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                            Else
                                str����ʱ�� = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    
    lng�Һſ���ID = Val(Nvl(mrsPlan!����ID))
    lng����ID = gobjDatabase.GetNextId("���˽��ʼ�¼")
    byt���� = IIf(Check����(Val(mrsInfo!����ID), lng�Һſ���ID), 1, 0)
    
    'Ʊ�ݴ���
    If mRegistFeeMode = EM_RG_���� Then
        str����NO = gobjDatabase.GetNextNo(13)
    End If
    lngSN = mlngSN
    strNO = gobjDatabase.GetNextNo(12)
    
    mrsItems.Filter = ""
    strҽ�� = NeedName(cboDoctor.Text)
    If cboDoctor.ListCount = 0 Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    strSQL = "Select ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rs���ʽ = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, Nvl(mrsInfo!ҽ�Ƹ��ʽ))
    If rs���ʽ.RecordCount <> 0 Then
        mstr������� = Nvl(rs���ʽ!����)
    Else
        mstr������� = mstrDef�������
    End If
    
    k = 1: mrsItems.MoveFirst
    For i = 1 To mrsItems.RecordCount
        int�۸񸸺� = k
        mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
        For j = 1 To mrsInComes.RecordCount
            strSQL = _
            "zl_���˹Һż�¼_INSERT(" & ZVal(Nvl(mrsInfo!����ID)) & "," & IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & txtPatient.Text & "','" & mstrGender & "'," & _
                     "'" & mstrAge & "','" & mstr������� & "','" & mstrFeeType & "','" & strNO & "'," & _
                     "'" & IIf(blnInvoicePrint = False, "", strFactNO) & "'," & k & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & IIf(mrsItems!���� = 2, 1, "NULL") & "," & _
                     "'" & mrsItems!��� & "'," & mrsItems!��ĿID & "," & mrsItems!���� & "," & mrsInComes!���� & "," & _
                     mrsInComes!������ĿID & ",'" & mrsInComes!�վݷ�Ŀ & "','" & IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), "") & "'," & _
                     IIf(mRegistFeeMode = EM_RG_����, 0, mrsInComes!Ӧ��) & "," & IIf(mRegistFeeMode = EM_RG_����, 0, mrsInComes!ʵ��) & "," & _
                     lng�Һſ���ID & "," & UserInfo.����ID & "," & IIf(mrsItems!ִ�п���ID = 0, lng�Һſ���ID, mrsItems!ִ�п���ID) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                     str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                     "'" & strҽ�� & "'," & ZVal(lngҽ��ID) & "," & IIf(mrsItems!���� = 3, 1, IIf(mrsItems!���� = 4, 2, 0)) & "," & IIf(lbl��.Visible, 1, 0) & "," & _
                     "'" & Get�ű� & "','" & IIf(strҽ�� = UserInfo.����, lblRoomName.Caption, "") & "'," & ZVal(lng����ID) & "," & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng����ID)) & "," & _
                     ZVal(IIf(k = 1, curԤ��, 0)) & "," & ZVal(IIf(k = 1, cur�ֽ�, 0)) & "," & _
                     ZVal(IIf(k = 1, cur����, 0)) & "," & ZVal(Nvl(mrsItems!���մ���ID, 0)) & "," & _
                     ZVal(Nvl(mrsItems!������Ŀ��, 0)) & "," & ZVal(Nvl(mrsInComes!ͳ����, 0)) & "," & _
                     "'" & IIf(str����NO <> "", "����:" & str����NO, Me.cboRemark.Text) & "'," & IIf(mblnAppointment, IIf(mty_Para.blnԤԼʱ�տ�, 0, 1), 0) & "," & IIf(mty_Para.bln�����շ�Ʊ��, 1, 0) & ",'" & mrsItems!���ձ��� & "'," & byt���� & "," & ZVal(lngSN) & ",Null," & _
                     IIf(mblnAppointment, 1, 0) & ",'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "'," & _
                     0 & ","
            '�����id_In   ����Ԥ����¼.�����id%Type := Null,
            strSQL = strSQL & "" & IIf(lngҽ�ƿ����ID <> 0 And bln���ѿ� = False, lngҽ�ƿ����ID, "NULL") & ","
            '���㿨���_In ����Ԥ����¼.���㿨���%Type := Null,
            strSQL = strSQL & "" & IIf(lngҽ�ƿ����ID <> 0 And bln���ѿ�, lngҽ�ƿ����ID, "NULL") & ","
            '����_In       ����Ԥ����¼.����%Type := Null,
            strSQL = strSQL & "'" & mstrCardNO & "',"
            '������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSQL = strSQL & " NULL,"
            '����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
            strSQL = strSQL & " NULL,"
            '������λ_In   ����Ԥ����¼.������λ%Type := Null
            strSQL = strSQL & " NULL,"
            '  ��������_In   Number:=0
            strSQL = strSQL & IIf(blnAdd, 1, 0) & ","
            '  ����_IN       ���˹Һż�¼.����%type:=null,
            strSQL = strSQL & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  ����ģʽ_IN   NUMBER :=0,
            strSQL = strSQL & IIf(mPatiChargeMode = EM_�����ƺ����, 1, 0) & ","
            '  ���ʷ���_IN Number:=0
            strSQL = strSQL & IIf(mRegistFeeMode = EM_RG_����, 1, 0) & ","
            '  �˺�����_IN Number:=1
            strSQL = strSQL & IIf(mty_Para.bln�˺�����, 1, 0) & ")"
            
            Call zlAddArray(cllPro, strSQL)
            '����:31187:���ҺŻ��ܵ�������
            If Get�ű� <> "" And k = 1 Then
                If Nvl(mrsPlan!ҽ��) = "" Then blnNoDoc = True
                strSQL = "zl_���˹ҺŻ���_Update("
                '  ҽ������_In   �ҺŰ���.ҽ������%Type,
                strSQL = strSQL & IIf(blnNoDoc, "Null,", "'" & strҽ�� & "',")
                '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
                strSQL = strSQL & "" & IIf(blnNoDoc, "0,", ZVal(lngҽ��ID) & ",")
                '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
                strSQL = strSQL & "" & Val(Nvl(mrsItems!��ĿID)) & ","
                '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
                strSQL = strSQL & "" & IIf(Val(Nvl(mrsItems!ִ�п���ID)) = 0, lng�Һſ���ID, Val(Nvl(mrsItems!ִ�п���ID))) & ","
                '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                strSQL = strSQL & "" & str����ʱ�� & ","
                '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����,3-�շ�ԤԼ
                strSQL = strSQL & IIf(mblnAppointment, IIf(mty_Para.blnԤԼʱ�տ�, 3, 1), 0) & ","
                '  ����_In       �ҺŰ���.����%Type := Null
                strSQL = strSQL & "'" & Get�ű� & "')"
                Call zlAddArray(cllProAfter, strSQL)
            End If
            
            If mRegistFeeMode = EM_RG_���� Then
                strSQL = _
                "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & k & "," & ZVal(Nvl(mrsInfo!����ID)) & ",NULL," & _
                         IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & mstr������� & "'," & _
                         "'" & txtPatient.Text & "','" & mstrGender & "','" & mstrAge & "'," & _
                         "'" & mstrFeeType & "',NULL," & lng�Һſ���ID & "," & _
                         IIf(lng�Һſ���ID <> 0, lng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & IIf(mrsItems!���� = 2, 1, "NULL") & "," & _
                         mrsItems!��ĿID & ",'" & mrsItems!��� & "','" & mrsItems!���㵥λ & "'," & _
                         "NULL,1," & mrsItems!���� & ",NULL," & IIf(mrsItems!ִ�п���ID = 0, lng�Һſ���ID, mrsItems!ִ�п���ID) & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & _
                         mrsInComes!������ĿID & ",'" & mrsInComes!�վݷ�Ŀ & "'," & mrsInComes!���� & "," & _
                         mrsInComes!Ӧ�� & "," & mrsInComes!ʵ�� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "',NULL,'�Һ�:" & strNO & "')"
                Call zlAddArray(cllPro, strSQL)
            End If
            k = k + 1
            mrsInComes.MoveNext
            Next j
        mrsItems.MoveNext
    Next i
    
    If GetSqlExpenses(cllPro, mrsExpenses, strNO, k, mRegistFeeMode = EM_RG_����, str����NO, ZVal(Nvl(mrsInfo!����ID)), _
                mstrClinic, lngSN, str����ʱ��, str�Ǽ�ʱ��, Not blnInvoicePrint, IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), ""), _
                IIf(strҽ�� = UserInfo.����, lblRoomName.Caption, ""), lng����ID, blnBalance, blnAdd, byt����) = False Then

        Exit Sub
    End If
    
    If Not mblnAppointment Then
        If strҽ�� = UserInfo.���� Then
            strSQL = "ZL_���˹Һż�¼_��������('" & strNO & "'," & Nvl(mrsInfo!����ID) & ",'" & lblRoomName.Caption & "','" & UserInfo.���� & "','','','" & zl_GetԤԼ��ʽByNo(strNO) & "')"    '�����:48350
            Call zlAddArray(cllPro, strSQL)
            strSQL = "zl_���˽���(" & Nvl(mrsInfo!����ID) & ",'" & strNO & "',NULL,'" & UserInfo.���� & "','" & lblRoomName.Caption & "')"
            Call zlAddArray(cllPro, strSQL)
        End If
    End If
    
    Err = 0: On Error GoTo ErrFirt:
    
    If cllPro.Count > 0 Then
        Err = 0: On Error GoTo ErrFirt:
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, False

        Err = 0: On Error GoTo errH:
        blnTrans = True
        If blnOneCard And lngҽ�ƿ����ID <> 0 And mRegistFeeMode = EM_RG_���� And cur�ֽ� <> 0 Then
            If Not mobjICCard.PaymentSwap(Val(cur�ֽ�), Val(cur�ֽ�), Val(lngҽ�ƿ����ID), 0, mstrCardNO, "", lng����ID, Nvl(mrsInfo!����ID)) Then
                gcnOracle.RollbackTrans
                MsgBox "һ��ͨ����Һŷ�ʧ��", vbInformation, gstrSysName
                Exit Sub
            Else
                strSQL = "zl_һ��ͨ����_Update(" & lng����ID & ",'" & cboPayMode.Text & "','" & mstrCardNO & "','" & lngҽ�ƿ����ID & "','" & "" & "'," & cur�ֽ� & ")"
                Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If

        'ҽ���Ķ�
        blnNotCommit = False
        If mintInsure <> 0 And mstrYBPati <> "" And cur���� <> 0 Then
            '68991:strAdvance:����ģʽ(0��1)|�Һŷ���ȡ��ʽ(0��1) |�Һŵ���
            strAdvance = ""
            If mRegistFeeMode = EM_RG_���� Or mPatiChargeMode = EM_�����ƺ���� Then
                strAdvance = IIf(mPatiChargeMode = EM_�����ƺ����, "1", "0")
                strAdvance = strAdvance & "|" & IIf(mRegistFeeMode = EM_RG_����, "1", "0")
                strAdvance = strAdvance & "|" & strNO
            End If
            If Not gclsInsure.RegistSwap(lng����ID, cur����, mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
            End If
            blnNotCommit = True
        End If
        '����:31187 ����ҽ���ɹ���,�����һЩ���ݸ���:�ڲ������������ύ���,���Բ�����д
        zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
        Set cllCardPro = New Collection: Set cllTheeSwap = New Collection
        If mRegistFeeMode = EM_RG_���� And Not blnOneCard And Not mPatiChargeMode = EM_�����ƺ���� And cur�ֽ� <> 0 Then
            If zlInterfacePrayMoney(lng����ID, cllCardPro, cllTheeSwap, Val(cur�ֽ�), lngҽ�ƿ����ID, bln���ѿ�) = False Then
                gcnOracle.RollbackTrans: If cmdOK.Enabled = False Then cmdOK.Enabled = True
                Exit Sub
            End If
            '������������
            zlExecuteProcedureArrAy cllCardPro, Me.Caption, False, False
        End If
        
        Err = 0: On Error GoTo OthersCommit:
        zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, False, False
OthersCommit:
        gcnOracle.CommitTrans

        If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, True, mintInsure)
        
        blnTrans = False
        On Error GoTo 0
    End If
    '��ӡ����
    If blnInvoicePrint Then
RePrint:
        If Not (mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) And mRegistFeeMode = EM_RG_���� Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me, "NO=" & strNO, 2)
            If gblnBill�Һ� Then
                If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
                    If MsgBox("�Һŵ���Ϊ[" & strNotValiedNos & "]Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����½���Ʊ�ݴ�ӡ!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                End If
            End If
        End If
    End If
    
    If blnSlipPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
    End If
    
    If blnAppointPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    End If
    
    If blnSlipPrint Or blnInvoicePrint Then
        '��¼��ӡ��ƾ��
        gstrSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    Call ReloadPage
    mstrNO = strNO
    mblnOK = True
    Unload Me
    Exit Sub
ErrFirt:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Sub
errH:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Sub
ErrGo:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetSqlExpenses(ByRef cllPro As Collection, ByVal rsExpenses As ADODB.Recordset, _
                    ByVal strRegNo As String, ByVal lngNoSort As Long, _
                    Optional ByVal bln��Ϊ���۵� As Boolean, Optional ByVal str����NO As String, _
                    Optional ByVal lng����ID As Long, Optional ByVal str����� As String, Optional ByVal lngSN As Long, _
                    Optional ByVal str����ʱ�� As String, Optional ByVal str�Ǽ�ʱ�� As String, Optional ByVal blnNoPrint As Boolean, _
                    Optional ByVal str���㷽ʽ As String, Optional ByVal strRoom As String, Optional ByVal lng����ID As Long, _
                    Optional ByVal blnBalance As Boolean, Optional ByVal blnAdd As Boolean, Optional ByVal byt���� As Byte) As Boolean
    '��ȡ���ӷѼ�¼sql
    Dim str�ѱ� As String, str���� As String, strSQL As String, strҽ�� As String
    Dim lng�Һſ���ID As Long, lngҽ��ID As Long
    Dim i As Long, lngPre��ĿID As Long
    Dim int�۸񸸺� As Integer
    Dim lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean, strBalanceStyle As String
    
    If rsExpenses Is Nothing Then GetSqlExpenses = True: Exit Function
    rsExpenses.Filter = ""
    If rsExpenses.RecordCount = 0 Then GetSqlExpenses = True: Exit Function
    On Error GoTo Errhand

    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lngҽ�ƿ����ID = mcolCardPayMode.Item(i)(3)
                bln���ѿ� = Val(mcolCardPayMode.Item(i)(5)) = 1
                strBalanceStyle = mcolCardPayMode.Item(i)(6)
            End If
        Next i
    End If
        
    lng�Һſ���ID = Val(Nvl(mrsPlan!����ID))
    strҽ�� = NeedName(cboDoctor.Text)
    If cboDoctor.ListCount = 0 Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    If lngNoSort <> 0 Then lngNoSort = lngNoSort - 1
    
    For i = 1 To rsExpenses.RecordCount
        lngNoSort = lngNoSort + 1
        
        strSQL = _
        "zl_���˹Һż�¼_INSERT(" & ZVal(lng����ID) & "," & IIf(str����� = "", "NULL", str�����) & ",'" & txtPatient.Text & "','" & mstrGender & "'," & _
        "'" & mstrAge & "','" & mstr������� & "','" & mstrFeeType & "','" & strRegNo & "'," & _
        "''," & lngNoSort & "," & IIf(lngPre��ĿID = rsExpenses!��ĿID, int�۸񸸺�, "NULL") & ",NULL," & _
        "'" & rsExpenses!��� & "'," & rsExpenses!��ĿID & "," & rsExpenses!���� & "," & rsExpenses!���� & "," & _
        rsExpenses!������ĿID & ",'" & rsExpenses!�վݷ�Ŀ & "','" & str���㷽ʽ & "'," & _
        IIf(bln��Ϊ���۵�, 0, rsExpenses!Ӧ��) & "," & IIf(bln��Ϊ���۵�, 0, rsExpenses!ʵ��) & "," & _
        lng�Һſ���ID & "," & UserInfo.����ID & "," & IIf(rsExpenses!ִ�п���ID = 0, lng�Һſ���ID, rsExpenses!ִ�п���ID) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
        str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
        "'" & strҽ�� & "'," & ZVal(lngҽ��ID) & ",0," & IIf(lbl��.Visible, 1, 0) & "," & _
        "'" & Get�ű� & "','" & strRoom & "'," & ZVal(lng����ID) & "," & IIf(blnNoPrint, "NULL", ZVal(mlng����ID)) & "," & _
        "0, 0, 0," & ZVal(Nvl(rsExpenses!���մ���ID, 0)) & "," & _
        ZVal(Nvl(rsExpenses!������Ŀ��, 0)) & "," & ZVal(Nvl(rsExpenses!ͳ����, 0)) & "," & _
        "'" & IIf(str����NO <> "", "����:" & str����NO, Me.cboRemark.Text) & "'," & IIf(mblnAppointment, IIf(mty_Para.blnԤԼʱ�տ�, 0, 1), 0) & "," & IIf(mty_Para.bln�����շ�Ʊ��, 1, 0) & ",'" & rsExpenses!���ձ��� & "'," & byt���� & "," & ZVal(lngSN) & ",Null," & _
        IIf(mblnAppointment, 1, 0) & ",'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "'," & _
        0 & ","
        '�����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & IIf(lngҽ�ƿ����ID <> 0 And bln���ѿ� = False, lngҽ�ƿ����ID, "NULL") & ","
        '���㿨���_In ����Ԥ����¼.���㿨���%Type := Null,
        strSQL = strSQL & "" & IIf(lngҽ�ƿ����ID <> 0 And bln���ѿ�, lngҽ�ƿ����ID, "NULL") & ","
        '����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "" & IIf(mstrCardNO <> "", "'" & mstrCardNO & "'", "NULL") & ","
        '������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & " NULL,"
        '����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & " NULL,"
        '������λ_In   ����Ԥ����¼.������λ%Type := Null
        strSQL = strSQL & " NULL,"
        '  ��������_In   Number:=0
        strSQL = strSQL & IIf(blnAdd, "1", "0") & ","
        '  ����_IN       ���˹Һż�¼.����%type:=null,
        strSQL = strSQL & IIf(mintInsure = 0, "NULL", mintInsure) & ","
        '  ����ģʽ_IN   NUMBER :=0,
        strSQL = strSQL & IIf(mPatiChargeMode = EM_�����ƺ����, 1, 0) & ","
        '  ���ʷ���_IN Number:=0,
        strSQL = strSQL & IIf(mRegistFeeMode = EM_RG_����, 1, 0) & ","
        '  �˺�����_IN Number:=1,
        strSQL = strSQL & IIf(mty_Para.bln�˺�����, 1, 0) & ")"
        Call zlAddArray(cllPro, strSQL)
        
        If bln��Ϊ���۵� Then
            strSQL = _
            "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & lngNoSort & "," & lng����ID & ",NULL," & _
                     IIf(str����� = "", "NULL", str�����) & ",'" & mstr������� & "'," & _
                     "'" & txtPatient.Text & "','" & mstrGender & "','" & mstrAge & "'," & _
                     "'" & mstrFeeType & "',NULL," & lng�Һſ���ID & "," & _
                     IIf(lng�Һſ���ID <> 0, lng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & "NULL" & "," & _
                     rsExpenses!��ĿID & ",'" & rsExpenses!��� & "','" & rsExpenses!���㵥λ & "'," & _
                     "NULL,1," & Val(Nvl(rsExpenses!����)) & ",NULL," & IIf(rsExpenses!ִ�п���ID = 0, lng�Һſ���ID, rsExpenses!ִ�п���ID) & "," & _
                     IIf(lngPre��ĿID = rsExpenses!��ĿID, int�۸񸸺�, "NULL") & "," & _
                     Val(Nvl(rsExpenses!������ĿID)) & ",'" & Trim(Nvl(rsExpenses!�վݷ�Ŀ)) & "'," & Val(Nvl(rsExpenses!����)) & "," & _
                     Val(rsExpenses!Ӧ��) & "," & Val(rsExpenses!ʵ��) & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "',NULL,'�Һ�:" & strRegNo & "')"
            Call zlAddArray(cllPro, strSQL)
        End If
        If lngPre��ĿID <> rsExpenses!��ĿID Then int�۸񸸺� = lngNoSort
        lngPre��ĿID = rsExpenses!��ĿID
        rsExpenses.MoveNext
    Next
    rsExpenses.MoveFirst
    GetSqlExpenses = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ReloadPage()
    On Error GoTo errHandle
    Call LoadRegPlans(False)
    Call ClearPatient
    Call ClearRegInfo
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function Check����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As Boolean
'����:�жϲ����Ƿ��ٴε�����ͬ�ٴ����ʵ��ٴ����ҡ��Һ�
'     �����ҹ��ŵ�,��ס��Ժ��,���ﲻ��ȷ��ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select a.�ٴ�����id" & vbNewLine & _
    "       From (Select ִ�в���id �ٴ�����id From ���˹Һż�¼ Where ����id = [1] and ��¼����=1 and ��¼״̬=1 " & vbNewLine & _
    "             Union All" & vbNewLine & _
    "             Select ��Ժ����id �ٴ�����id From ������ҳ Where ����id = [1]) a" & vbNewLine & _
    "       Where Exists (Select 1" & vbNewLine & _
    "                    From �ٴ����� b" & vbNewLine & _
    "                    Where b.����id = a.�ٴ�����id And b.�������� = (Select �������� From �ٴ����� Where ����id = [2] And Rownum=1))"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngִ�в���ID)
    Check���� = Not rsTmp.EOF
End Function

Private Function zlIsAllowPatiChargeFeeMode(ByVal lng����ID As Long, ByVal intԭ����ģʽ As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�����ı䲡���շ�ģʽ
    '���:lng����ID-����ID
    '       intԭ����ģʽ-0��ʾ�Ƚ��������;1��ʾ�����ƺ����
    '����:��������շ�ģʽ,����true,���򷵻�False
    '����:���˺�
    '����:2013-12-25 10:06:49
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim dtDate As Date, intDay As Integer
    On Error GoTo errHandle
    
'    If mbytMode = 1 Then zlIsAllowPatiChargeFeeMode = True: Exit Function 'ԤԼ������
    'ģʽδ������ֱ�ӷ���true
    If intԭ����ģʽ = mPatiChargeMode Then zlIsAllowPatiChargeFeeMode = True: Exit Function
    
      
    If intԭ����ģʽ = 1 Then
        'ԭΪ�����ƺ�����Ҵ���δ����õ�,�������ü���ģʽ
        strSQL = "" & _
        "   Select 1 " & _
        "   From ����δ����� " & _
        "   Where ����id = [1] And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If rsTemp.EOF = False Then
            MsgBox "ע��:" & vbCrLf & "  ��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�" & _
                                          vbCrLf & "����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ�" & _
                                          vbCrLf & "�ٹҺŻ򲻵������˵ľ���ģʽ", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        intDay = -1 * Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        dtDate = DateAdd("d", intDay, gobjDatabase.CurrentDate)
        ' �ϴ�Ϊ"�����ƺ����",����Ϊ"�Ƚ��������"��,ͬʱ����δ����ҽ��ҵ�����ݵ� ,
        '   ��������ľ���ģʽ
        strSQL = "Select 1 " & _
        " From ���˹Һż�¼ A, ����ҽ����¼ B " & _
        " Where a.����id + 0 = b.����id And a.No || '' = b.�Һŵ�  " & _
        "               And a.��¼״̬ = 1 And a.��¼���� = 1 And a.�Ǽ�ʱ�� - 0 >= [2] " & _
        "               And  a.����id = [1] And rownum<2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, dtDate)
        If rsTemp.EOF Then
            'δ����ҽ������
            MsgBox "ע��:" & vbCrLf & "  ��ǰ���˵ľ���ģʽΪ�����ƺ����," & vbCrLf & "  ����������ò��˵ľ���ģʽ!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    zlIsAllowPatiChargeFeeMode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub ClearRegInfo()
    If cboArrangeNo.ListCount <> 0 Then cboArrangeNo.ListIndex = 0
    lblDeptName.Caption = ""
    lblRoomName.Caption = ""
    cboRemark.Text = ""
    chkBook.Value = IIf(mty_Para.blnĬ�Ϲ�����, 1, 0)
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = "0.00"
    lblPayMoney.Caption = "0.00"
    txtPatient.SetFocus
    lbl��.Visible = False
End Sub

Private Function zlIsNotSucceedPrintBill(ByVal bytType As Byte, ByVal strNos As String, ByRef strOutValidNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ��Ѿ�������ӡ
    '���:bytType-1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '       strNos-���δ�ӡƱ�ݵĵ���,�ö��ŷ���
    '����:strOutValidNos-��ӡʧ�ܵĵ��ݺ�
    '����:���ڲ��湦Ʊ�ݵĴ�ӡ,����true,���򷵻�False
    '����:���˺�
    '����:2012-01-16 18:06:01
    '����:44322,44326,44332,44330
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempNos As String, rsTemp As ADODB.Recordset
    Dim strSQL As String, strBillNos As String
    Dim bytBill As Byte
    On Error GoTo errHandle
    strBillNos = Replace(Replace(strNos, "'", ""), " ", "")
    strSQL = "" & _
        "Select  /*+ rule */ distinct  B.NO " & _
        " From Ʊ��ʹ����ϸ A,Ʊ�ݴ�ӡ���� B,Table( f_Str2list([2])) J" & _
        " Where A.��ӡID =b.ID And B.��������=[1] And B.No=J.Column_value "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���Ʊ���Ƿ��ӡ", bytType, strBillNos)
    
    strTempNos = ""
    With rsTemp
        Do While Not .EOF
            If InStr(1, "," & strBillNos & ",", "," & !NO & ",") = 0 Then
                strTempNos = strTempNos & "," & !NO
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then strTempNos = "," & strBillNos
    End With
    If strTempNos <> "" Then strTempNos = Mid(strTempNos, 2)
    rsTemp.Close: Set rsTemp = Nothing
    strOutValidNos = strTempNos
    zlIsNotSucceedPrintBill = strTempNos <> ""
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckValied() As Boolean
    Dim i As Integer
    '����ǰ���
    If mrsInfo Is Nothing Then
        MsgBox "�޷�ȷ��������Ϣ,����ѡ��һ�����ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "�޷�ȷ��������Ϣ,����ѡ��һ�����ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan Is Nothing Then
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.State = 0 Then
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.RecordCount = 0 Then
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If GetRegistMoney < 0 Then
        MsgBox "�Һŷ��ò���Ϊ����������Һ���Ŀ��", vbInformation, gstrSysName
        If cboArrangeNo.Visible And cboArrangeNo.Enabled Then cboArrangeNo.SetFocus
        Exit Function
    End If
    
    If cboPayMode.Text = "" And cboPayMode.Visible And Val(lblTotal.Caption) <> 0 Then
        MsgBox "û��ȷ�����õĽ��㷽ʽ,������ɹҺ�!", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then
        If IsNull(mrsPlan!�Ű�) Then
            MsgBox "ԤԼ���տ�ģʽ��,���ܹҲ�����ĺű�!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Nvl(mrsInfo!����) <> txtPatient.Text Then
        If MsgBox("��ǰ���������Ѿ������仯,�Ƿ����¶�ȡ������Ϣ?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            Call GetPatient(IDKind.GetCurCard, txtPatient.Text, False)
            Exit Function
        Else
            txtPatient.Text = Nvl(mrsInfo!����)
        End If
    End If
    
    If InStr(gstrPrivs, ";�Һŷѱ����;") = 0 Then
        For i = 1 To vsfMoney.Rows - 1
            If Val(vsfMoney.TextMatrix(i, 2)) <> Val(vsfMoney.TextMatrix(i, 1)) Then
                MsgBox "��û��Ȩ�޸�����ʹ�ô��۷ѱ�,������ɹҺ�", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    CheckValied = True
End Function

Private Function zlInterfacePrayMoney(ByVal lng�ҺŽ���ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double, lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:cllPro-�޸�������������
    '        cll��������-����������������
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lngҽ�ƿ����ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, lng�ҺŽ���ID, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '����������������
     If lng�ҺŽ���ID <> 0 Then
        '����:58322
        'mbytMode As Integer '0-�Һ�,1-ԤԼ,2-����,3-ȡ��ԤԼ ,4-�˺� ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
        If Not bln���ѿ� Then
            '���ѿ��Ѿ��ڲ���Һż�¼ʱ,�Ѿ��ۿ�
            Call zlAddUpdateSwapSQL(False, lng�ҺŽ���ID, lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lng�ҺŽ���ID, lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddThreeSwapSQLToCollection(ByVal blnԤ���� As Boolean, _
    ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���� As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    ' ����:cllPro-����SQL��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
    varData = Split(strExpend, "||")
    Dim str������Ϣ As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                    str������Ϣ = Mid(str������Ϣ, 3)
                    'Zl_�������㽻��_Insert
                    strSQL = "Zl_�������㽻��_Insert("
                    '�����id_In ����Ԥ����¼.�����id%Type,
                    strSQL = strSQL & "" & lng�����ID & ","
                    '���ѿ�_In   Number,
                    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                    '����_In     ����Ԥ����¼.����%Type,
                    strSQL = strSQL & "'" & str���� & "',"
                    '����ids_In  Varchar2,
                    strSQL = strSQL & "'" & strIDs & "',"
                    '������Ϣ_In Varchar2:������Ŀ|��������||...
                    strSQL = strSQL & "'" & str������Ϣ & "',"
                    'Ԥ����ɿ�_In Number := 0
                    strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
                    zlAddArray cllPro, strSQL
                    str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & "||" & strTemp
            End If
        End If
    Next
    If str������Ϣ <> "" Then
        str������Ϣ = Mid(str������Ϣ, 3)
        'Zl_�������㽻��_Insert
        strSQL = "Zl_�������㽻��_Insert("
        '�����id_In ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '���ѿ�_In   Number,
        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
        '����_In     ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '����ids_In  Varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '������Ϣ_In Varchar2:������Ŀ|��������||...
        strSQL = strSQL & "'" & str������Ϣ & "',"
        'Ԥ����ɿ�_In Number := 0
        strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
        zlAddArray cllPro, strSQL
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddUpdateSwapSQL(ByVal blnԤ�� As Boolean, ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    str���� As String, str������ˮ�� As String, str����˵�� As String, _
    ByRef cllPro As Collection, Optional intУ�Ա�־ As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������ˮ�ź���ˮ˵��
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    '����:cllPro-����SQL��
    '����:���˺�
    '����:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "Zl_�����ӿڸ���_Update("
    '  �����id_In   ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '  ���ѿ�_In     Number,
    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ����ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type
    strSQL = strSQL & "'" & str����˵�� & "',"
    'Ԥ����ɿ�_In Number := 0
    strSQL = strSQL & "" & IIf(blnԤ��, 1, 0) & ","
    '�˷ѱ�־ :1-�˷�;0-����
    strSQL = strSQL & "0,"
    'У�Ա�־
    strSQL = strSQL & "" & IIf(intУ�Ա�־ = 0, "NULL", intУ�Ա�־) & ")"
    zlAddArray cllPro, strSQL
End Function

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
        If cboArrangeNo.ListCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub SetControl()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTemp As String, i As Integer
    If mblnAppointment Then
        lblRoom.Visible = False
        picRoom.Visible = False
        lblDept.Visible = False
        picDept.Visible = False
        lblDept.Left = lblLimit.Left
        picDept.Left = lblDept.Left + lblDept.Width + 30
        picDept.Width = Me.Width - 240 - picDept.Left
        chkBook.Value = 0
        chkBook.Visible = False
        cboRemark.Width = 7170
        If mty_Para.blnԤԼʱ�տ� Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
        End If
        cboAppointStyle.Clear
        strSQL = "Select ����,ȱʡ��־ From ԤԼ��ʽ"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Do While Not rsTmp.EOF
            cboAppointStyle.AddItem Nvl(rsTmp!����)
            If Val(Nvl(rsTmp!ȱʡ��־)) = 1 Then cboAppointStyle.ListIndex = cboAppointStyle.NewIndex
            rsTmp.MoveNext
        Loop
        strTemp = gobjDatabase.GetPara("ȱʡԤԼ��ʽ", glngSys, 9000, "")
        For i = 0 To cboAppointStyle.ListCount - 1
            If Mid(cboAppointStyle.List(i), InStr(cboAppointStyle.List(i), ".") + 1) = strTemp Then
                cboAppointStyle.ListIndex = i
            End If
        Next i
    Else
        lblDate.Visible = False
        lblTime.Visible = False
        dtpDate.Visible = False
        dtpTime.Visible = False
        cmdTime.Visible = False
        If mty_Para.byt�Һ�ģʽ = 0 Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Err = 0
    mblnInit = True
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    If mblnAppointment Then
        Me.Caption = "ҽ��վԤԼ"
        lblAppointStyle.Visible = True
        cboAppointStyle.Visible = True
    Else
        Me.Caption = "ҽ��վ�Һ�"
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
    End If
    
    gobjDatabase.ExecuteProcedure "Zl_�ҺŰ���_Autoupdate", Me.Caption
    Call InitData
    Call InitPara
    chkBook.Value = IIf(mty_Para.blnĬ�Ϲ�����, 1, 0)
    Call InitIDKind
    Call InitTime
    Call GetAllҽ��
    If LoadRegPlans(False) = False Then
        mblnUnload = True
    End If
    Call LoadPayMode
    Call SetControl
    If mblnAppointment And mlng����ID <> 0 Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng����ID, False)
    End If
    cmdNewPati.Enabled = InStr(gstrPrivs, ";�ҺŲ��˽���;") > 0
End Sub

Private Sub InitTime()
    dtpDate.Value = Format(gobjDatabase.CurrentDate + mintSysAppLimit, "yyyy-mm-dd")
    dtpDate.minDate = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd")
    dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.����
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ҽ������鿨
    '���ƣ����˺�
    '���ڣ�2010-07-14 11:32:08
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim str�������� As String
    Dim rsTmp As ADODB.Recordset
    Dim cur��� As Currency
    Dim curMoney As Currency
    Dim blnDeposit As Boolean, blnInsure As Boolean
    If mrsInfo Is Nothing Then
        lng����ID = 0
        str�������� = ""
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        str�������� = Nvl(mrsInfo!��������)
    End If

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False

    Dim strAdvance As String    '����ģʽ(0-�Ƚ�������ƻ�1-�����ƺ����)|�Һŷ���ȡ��ʽ(0-���ջ�1-����)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng����ID, mintInsure, strAdvance)
    mRegistFeeMode = EM_RG_����: mPatiChargeMode = EM_�Ƚ��������
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng����ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
    
    txtPatient.Text = "-" & lng����ID
    Call txtPatient_Validate(False)    '���е�Setfocus����ʹ���¼�(txtPatient_KeyPress)ִ�����,�����ٴ��Զ�ִ��txtPatient_Validate
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
    Call SetPatiColor(txtPatient, str��������, vbRed)
    txtPatient.BackColor = &HE0E0E0
    txtPatient.Locked = True

    If strAdvance <> "" Then
        varData = Split(strAdvance & "|", "|")
        mPatiChargeMode = IIf(Val(varData(0)) = 1, EM_�����ƺ����, EM_�Ƚ��������)
        mRegistFeeMode = IIf(Val(varData(1)) = 1, EM_RG_����, EM_RG_����)
    End If
    MCPAR.���ղ����� = gclsInsure.GetCapability(support�ҺŲ���ȡ������, lng����ID, mintInsure)
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure)
    mlng����ID = 0
    curMoney = GetRegistMoney
    Set rsTmp = GetMoneyInfoRegist(lng����ID, , , 1)

    cur��� = 0
    Do While Not rsTmp.EOF
        cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
        cur��� = cur��� - Val(Nvl(rsTmp!�������))

        rsTmp.MoveNext
    Loop
    If cur��� > 0 Then
        lblMoney.Caption = "����Ԥ�����:" & Format(cur���, "0.00")
        If cur��� >= curMoney Then
            blnDeposit = True
        Else
            blnDeposit = False
        End If
    End If
    
    mcur������� = gclsInsure.SelfBalance(lng����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
    lblMoney.Caption = lblMoney.Caption & "/�����ʻ����:" & Format(mcur�������, "0.00")
    Call GetYBInfo
    If gclsInsure.GetCapability(support�Һ�ʹ�ø����ʻ�, lng����ID, mintInsure) = False Then
        blnInsure = False
    Else
        If mcur������� + mcur����͸֧ >= curMoney Then
            blnInsure = True
        Else
            blnInsure = False
        End If
    End If
    Call LoadPayMode(blnDeposit, blnInsure)
    If mRegistFeeMode = EM_RG_���� Then
        lblSum.Caption = "����"
        picPayMoney.Visible = False
        cboPayMode.Visible = False
        lblPayMode.Visible = False
    Else
        lblSum.Caption = "�ϼ�"
    End If
    If mRegistFeeMode = EM_RG_���� Then
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_����
        Else
            If mty_Para.byt�Һ�ģʽ = 0 Then
                mRegistFeeMode = EM_RG_����
                picPayMoney.Visible = True
                cboPayMode.Visible = True
                lblPayMode.Visible = True
            Else
                mRegistFeeMode = EM_RG_����
                picPayMoney.Visible = False
                cboPayMode.Visible = False
                lblPayMode.Visible = False
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '0-�����,1-����,2-�Һŵ�,3-���￨��,4-ҽ����
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    Static sngBegin As Single
    Dim sngNow As Single
    'ҽ����֤
    If txtPatient.Text = "" And KeyAscii = 13 Then
        KeyAscii = 0
        Call zlInusreIdentify
    End If
    
    If KeyAscii <> 0 And KeyAscii > 32 And mty_Para.bln�Һű���ˢ�� Then
        sngNow = Timer
        If txtPatient.Text = "" Then
            sngBegin = sngNow
        ElseIf Format((sngNow - sngBegin) / (Len(txtPatient.Text) + 1), "0.000") >= 0.04 Then    '>0.007>=0.01
            txtPatient.Text = Chr(KeyAscii)
            txtPatient.SelStart = 1
            KeyAscii = 0
            sngBegin = sngNow
        End If
    End If
    
    strKind = IDKind.GetCurCard.����
    txtPatient.PasswordChar = IIf(IDKind.GetCurCard.�������Ĺ��� <> "", "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.blnȱʡ��������)
        intLen = gobjSquare.intȱʡ���ų���
    Case "�����"
        If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "���֤"
    Case "ҽ����"
    Case Else
            If IDKind.GetCurCard.�ӿ���� <> 0 Then
                blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.GetCurCard.�������Ĺ��� <> "")
                intLen = IDKind.GetCurCard.���ų���
            End If
    End Select
    
    'ˢ����ϻ���������س�
    If (blnCard And Len(txtPatient.Text) = intLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13) Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0: mblnCard = True
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCard)
        mblnCard = False
        gobjControl.TxtSelAll txtPatient
   End If
End Sub

Private Function CheckNoValied(ByVal lngRow As Long) As Boolean
    CheckNoValied = True
End Function

Private Function zl_GetԤԼ��ʽByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݹҺŵ��ݺŻ�ȡ����ԤԼ��ʽ
    '���:strNo-�Һŵ��ݺ�
    '����:ԤԼ��ʽ
    '����:����
    '����:2012-07-03
    '�����:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strԤԼ��ʽ As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select ԤԼ��ʽ From ���˹Һż�¼ Where ��¼״̬=1 And No=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡԤԼ��ʽ", strNO)
    If rsTemp Is Nothing Then zl_GetԤԼ��ʽByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_GetԤԼ��ʽByNo = "": Exit Function
    While rsTemp.EOF = False
        strԤԼ��ʽ = Nvl(rsTemp!ԤԼ��ʽ)
        rsTemp.MoveNext
    Wend
    zl_GetԤԼ��ʽByNo = strԤԼ��ʽ
End Function

Public Function GetʧԼ��(ByVal str�ű� As String, ByVal datThis As Date) As Long
   '��ȡ������ĳһ��.ԤԼʧԼ��
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strDat  As String
'    If mty_Para.blnʧԼ���ڹҺ� = False Or mty_Para.lngԤԼ��Чʱ�� <= 0 Then Exit Function
    strSQL = "                " & " SELECT count(1) AS ʧԼ�� "
    strSQL = strSQL & vbNewLine & " FROM �Һ����״̬ "
    strSQL = strSQL & vbNewLine & " WHERE ����=[1] AND ״̬=2 AND ����-[3]/24/60 <SYSDATE AND To_Char(����,'yyyy-MM-dd')=[2]"
    strDat = Format(datThis, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�, strDat, mty_Para.lngԤԼ��Чʱ��)
    If rsTmp.EOF Then
        GetʧԼ�� = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    GetʧԼ�� = Val(Nvl(rsTmp!ʧԼ��, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
'        gobjControl.TxtSelAll txtPatient
'    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strID:
        If txtPatient.Text = "" Then
            Call mobjIDCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        DoEvents
        If txtPatient.Visible = True And txtPatient.Enabled Then
            Call txtPatient.SetFocus
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdNewPati_Click
    Else
        IDKind.ActiveFastKey
    End If
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strNO
        If txtPatient.Text = "" Then
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC����")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnInputIDCard As Boolean = False, Optional ByRef Cancel As Boolean)
    '���ܣ���ȡ������Ϣ
    '������blnCard=�Ƿ���￨ˢ��
    '
    '         blnInputIDCard-�Ƿ����֤ˢ��
    '����:Cancel-Ϊtrue��ʾ���صķ�����ȡ������Ϣ
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim strInputInfo As String '���洫��������ı� ������ʹ�����֤�� �Բ��˽��в��Һ� ���滻��"-" ����ID�����
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim blnҽ���� As Boolean, rsFeeType As ADODB.Recordset
    Dim IntMsg As VbMsgBoxResult
    Dim blnOtherType As Boolean '�Ƿ������

    strInputInfo = strInput
    
    On Error GoTo errH
    blnҽ���� = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard
    
    If objCard.���� Like "IC����" Or objCard.���� Like "IC��" Then '����IC�������Ӧ��ȡIC������
        strSQL = "Select  A.����ID,A.�����,A.סԺ��,A.���￨��,A.�ѱ�,A.ҽ�Ƹ��ʽ,A.����,A.�Ա�,A.����,A.��������,A.�����ص�,A.���֤��,A.����֤��,A.���,A.ְҵ,A.����,A.��������, " & _
                 "A.����,A.����,A.����,A.ѧ��,A.����״��,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�໤��,A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.���ڵ�ַ, " & _
                 "A.���ڵ�ַ�ʱ�,A.Email,A.QQ,A.��ͬ��λid,A.������λ,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.������,A.������,A.��������,A.����ʱ��,A.����״̬, " & _
                 "A.��������,A.סԺ����,A.��ǰ����id,A.��ǰ����id,A.��ǰ����,A.��Ժʱ��,A.��Ժʱ��,A.��Ժ,A.IC����,A.������,A.ҽ����,A.����,A.��ѯ����,A.�Ǽ�ʱ��,A.ͣ��ʱ��,A.����,A.��ϵ�����֤��, " & _
                 "B.���� ��������,C.���� As ����֤��,A.����ģʽ From ������Ϣ A,������� B,����ҽ�ƿ���Ϣ C Where A.���� = B.���(+) And A.ͣ��ʱ�� is NULL And A.����ID= C.����ID(+) And C.����= '" & UCase(strInput) & "'"
    Else
        strSQL = "Select A.*,B.���� �������� From ������Ϣ A,������� B Where A.���� = B.���(+) And A.ͣ��ʱ�� is NULL "
    End If
    If mty_Para.blnסԺ���˹Һ� = False Then
        str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID=A.����ID   And ��ҳID<>0 And ��ҳID=A.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    End If
   
    If blnCard And objCard.���� Like "����*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        ElseIf IDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
'        Else
'            lng�����ID = gCurSendCard.lng�����ID
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0

        If lng����ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And A.�����=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And A.����ID=[2]" & _
        IIf(mstrYBPati <> "", "", str����Ժ)
    ElseIf blnInputIDCard Then  '���������֤ʶ��
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        strInput = "-" & lng����ID
        strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                If Not mty_Para.bln����ģ������ Or mty_Para.bln����ģ������ And Len(txtPatient.Text) < 2 Then
                    Set mrsInfo = Nothing: Exit Sub
                End If
                strPati = _
                    " Select distinct 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                    " From ������Ϣ A " & _
                    " Where Rownum <101 And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & str����Ժ & _
                    IIf(mty_Para.lng������������ = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                    
'                strPati = strPati & " Union ALL " & _
'                        "Select 0,0 as ID,-NULL,'[�²���]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                strPati = strPati & " Order by ����ID,����"
                    
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTmp = gobjDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", mty_Para.lng������������)
                If Not rsTmp Is Nothing Then
                    If rsTmp!ID = 0 Then '�����²���
                        txtPatient.Text = ""
                        MsgBox "û���ҵ���Ӧ�Ĳ�����Ϣ������������Ϣ�Ƿ���ȷ���߲����Ƿ񽨵���", vbInformation, gstrSysName
                        Set mrsInfo = Nothing: Exit Sub
                    Else '�Բ���ID��ȡ
                        strInput = rsTmp!����ID
                        strSQL = strSQL & " And A.����ID=[1]"
                    End If
                Else 'ȡ��ѡ��
                    txtPatient.Text = ""
                    Set mrsInfo = Nothing: Exit Sub
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                blnҽ���� = True
                If mblnOlnyBJYB And gobjCommFun.ActualLen(strInput) >= 9 Then
                    strSQL = strSQL & " And A.ҽ���� like [3] " & str����Ժ
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And A.ҽ����=[1]" & str����Ժ
                End If
                
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
                 
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[1]" & str����Ժ
             Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                strSQL = strSQL & " And A.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
ReadPati:
    If Mid(mstrCardPass, 1, 1) = "1" And strPassWord <> "" Then
        If Not gobjCommFun.VerifyPassWord(Me, "" & strPassWord) Then
            MsgBox "���������֤ʧ�ܣ�", vbInformation, gstrSysName
            ClearPatient
            Exit Sub
        End If
    End If
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Val(Mid(strInput, 2)), strTemp)
    strInput = strInputInfo
    If Not mrsInfo.EOF Then
        Set mfrmPatiInfo = New frmPatiInfo
        txtPatient.Text = Nvl(mrsInfo!����) '�����Change�¼�
        txtPatient.BackColor = &H80000005
        lblSum.Caption = "�ϼ�"
        Call SetControl
        '�ڵ���txtPatient_Change�¼���������źͲ���������Ϊ�յ������ �޷�ʶ��ò�����Ϣ ���ִ���
        '���������ݿ����ݴ����ٽ��к����Ĵ���
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(Trim(mstr����) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        mstrGender = Nvl(mrsInfo!�Ա�)
        txtPatient.PasswordChar = ""
        
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        mstrFeeType = Nvl(mrsInfo!�ѱ�)
        If mstrFeeType = "" Then mstrFeeType = mstrDef�ѱ�
        mstrAge = Nvl(mrsInfo!����)
        mstrClinic = Nvl(mrsInfo!�����)
        If mstrClinic = "" Then
            mstrClinic = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        lblInfo.Caption = "�Ա�:" & mstrGender & "   ����:" & mstrAge & "   �����:" & mstrClinic & "   �ѱ�:" & mstrFeeType
        
        '����Ԥ������Ϣ
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!����ID, , , 1)
        cur��� = 0
        Do While Not rsTmp.EOF
            cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
            cur��� = cur��� - Val(Nvl(rsTmp!�������))
            rsTmp.MoveNext
        Loop
        If cur��� > 0 Then
            lblMoney.Caption = "����Ԥ�����:" & Format(cur���, "0.00")
            curMoney = GetRegistMoney
            If cur��� >= curMoney Then
                Call LoadPayMode(True)
            Else
                Call LoadPayMode
            End If
        Else
            lblMoney.Caption = "����Ԥ�����:0.00"
            Call LoadPayMode
        End If
        Call LoadFeeItem(Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1)
        cmdNewPati.ToolTipText = "��ϸ��Ϣ"
        cmdNewPati.Enabled = True
        If cboArrangeNo.Enabled And cboArrangeNo.Visible Then cboArrangeNo.SetFocus
        If cboArrangeNo.ListCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
    Else
NewPati:
        MsgBox "û���ҵ���Ӧ�Ĳ�����Ϣ������������Ϣ�Ƿ���ȷ���߲����Ƿ񽨵���", vbInformation, gstrSysName
        ClearPatient
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ClearPatient()
    Set mfrmPatiInfo = New frmPatiInfo
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    mlngNewPatiID = 0
    mstrGender = ""
    mstrAge = ""
    cmdNewPati.ToolTipText = "��������"
    cmdNewPati.Enabled = InStr(gstrPrivs, ";�ҺŲ��˽���;") > 0
    mstrClinic = ""
    mblnNewPati = False
    mstrFeeType = ""
    lblInfo.Caption = "�Ա�:     ����:       �����:              �ѱ�:  "
    lblMoney.Caption = "����Ԥ�����:0.00  "
    lblSum.Caption = "�ϼ�"
    mintInsure = 0
    mlng����ID = 0
    chkBook.Enabled = True
    LoadPayMode False, False
    Set mrsInfo = Nothing
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_����
    Else
        If mty_Para.byt�Һ�ģʽ = 0 Then
            mRegistFeeMode = EM_RG_����
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
        Else
            mRegistFeeMode = EM_RG_����
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
        End If
    End If
End Sub

Private Function GetRegistMoney(Optional blnOnlyReg As Boolean) As Currency
    '���ܣ���ȡ��ǰ�Һŵ��ĺϼƽ��
    'blnOnlyReg-�Ƿ������ȡ�Һŷ���
    Dim cur�ϼ� As Currency, i As Integer
    Dim curӦ�� As Currency, j As Integer
    Dim k As Integer
    If Not blnOnlyReg Then
        For i = 1 To vsfMoney.Rows - 1
            cur�ϼ� = cur�ϼ� + Val(vsfMoney.TextMatrix(i, 2))
        Next
    Else
        For i = 1 To vsfMoney.Rows - 1
            cur�ϼ� = cur�ϼ� + Val(vsfMoney.TextMatrix(i, 2))
        Next
    End If
    GetRegistMoney = cur�ϼ�
End Function

Private Sub LoadPayMode(Optional ByVal blnPrepay As Boolean = False, Optional ByVal blnInsure As Boolean = False)
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, str���� As String
    
    strSQL = _
        "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ And Instr([2] ,','||B.����||',')>0" & _
        " Order by B.����"
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, "�Һ�", ",3,7,8,")
    
    Set mcolCardPayMode = New Collection
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    
    With cboPayMode
        .Clear: j = 0
'        Do While Not rsTemp.EOF
'            blnFind = False
'            For i = 0 To UBound(varData)
'                varTemp = Split(varData(i) & "|||||", "|")
'                If varTemp(6) = Nvl(rsTemp!����) Then
'                    blnFind = True
'                    Exit For
'                End If
'            Next
'
'            If Not blnFind Then
'                .AddItem Nvl(rsTemp!����)
'                mcolCardPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
'                If Val(Nvl(rsTemp!ȱʡ)) = 1 Then
'                    If .ListIndex = -1 Then
'                         .ItemData(.NewIndex) = 1: .ListIndex = .NewIndex
'                    End If
'                End If
'                j = j + 1
'            End If
'            rsTemp.MoveNext
'        Loop
     
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                rsTemp.Filter = "����='" & varTemp(6) & "'"
                If Not rsTemp.EOF Then
                    mcolCardPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
    End With
    
    If blnPrepay Then
        cboPayMode.AddItem "Ԥ����"
        If mty_Para.bln����ʹ��Ԥ�� Then
            cboPayMode.ListIndex = cboPayMode.NewIndex
        End If
    End If
    
    If blnInsure Then
        rsTemp.Filter = "���� = 3"
        If rsTemp.EOF Then
            mstrInsure = ""
            MsgBox "���ܼ���ҽ�����㷽ʽ,����!", vbInformation, gstrSysName
        Else
            cboPayMode.AddItem Nvl(rsTemp!����)
            mstrInsure = Nvl(rsTemp!����)
            If Not mty_Para.bln����ʹ��Ԥ�� Or blnPrepay = False Then
                cboPayMode.ListIndex = cboPayMode.NewIndex
            End If
            If (mintInsure <> 0 And MCPAR.���ղ�����) And cboPayMode.Text = mstrInsure And cboPayMode.Visible Then
                chkBook.Enabled = False
                chkBook.Value = 0
            Else
                chkBook.Enabled = True
            End If
        End If
    End If
    
    If cboPayMode.ListCount > 0 And cboPayMode.ListIndex = -1 Then
        cboPayMode.ListIndex = 0
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function LoadRegPlans(ByVal blnCache As Boolean) As Boolean
    Dim strTime As String, strState As String, strWhere As String
    Dim strSQL As String, strIF As String
    Dim i As Integer, k As Integer
    Dim DateThis As Date, strZero As String
    Dim str�ҺŰ��� As String
    Dim str�ҺŰ��żƻ� As String
    Dim str����         As String
    On Error GoTo errH
    
    str���� = "Decode(ҽ��,Null,3,Decode(����ID," & mlngDept & ",1,2)),ҽ��,����,�ű�,��Ŀ,�ѹ�"
    
    If Not blnCache Then
        If gstrDeptIDs <> "" Then strIF = " And Instr(','||[4]||',',','||P.����ID||',')>0"
        If mty_Para.bln�������Ұ��� Then
            strIF = strIF & " And (P.ҽ������ = [1] or P.ҽ������ Is Null)"
        Else
            strIF = strIF & " And (P.ҽ������ = [1])"
        End If
        
        str�ҺŰ��� = "" & _
                "            Select A.ID, A.����, A.����, A.����id, A.��Ŀid, A.ҽ��id, A.ҽ������, A.��������, A. ����, A.��һ, A.�ܶ�, A.����, " & _
                "                   A.���� , A.����, A.����, A.���﷽ʽ,a.��ʼʱ��,a.��ֹʱ��, A.��ſ���, B.�޺���, B.��Լ��,a.ͣ������ " & vbNewLine & _
                "            From �ҺŰ��� A, �ҺŰ������� B " & vbNewLine & _
                "            Where a.ͣ������ Is Null And " & "[5] Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
                "                 Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                "                  And a.ID = B.����id(+) And Trunc(Sysdate)+Nvl(A.ԤԼ����," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5] And Decode(To_Char([5], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) = B.������Ŀ(+)" & vbNewLine
      
        If mblnAppointment Then
            DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
        Else
            DateThis = gobjDatabase.CurrentDate
        End If
        'ȡ��Ӧ���ڰ��ŵ�ʱ���
        strSQL = "Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL)"
        
        '�ò������ȡ��������Ӧ��ʱ���
        strTime = _
            "Select ʱ��� From ʱ��� Where ���� Is Null And վ�� Is Null And " & _
            "    ('3000-01-10 '||To_Char([5],'HH24:MI:SS') Between" & _
            "               Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-09 '||To_Char(Nvl(��ǰʱ��,��ʼʱ��),'HH24:MI:SS'),'3000-01-10 '||To_Char(Nvl(��ǰʱ��,��ʼʱ��),'HH24:MI:SS'))" & _
            "               And '3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS'))" & _
            " Or" & _
            " ('3000-01-10 '||To_Char([5],'HH24:MI:SS')  Between" & _
            "   '3000-01-10 '||To_Char(Nvl(��ǰʱ��,��ʼʱ��),'HH24:MI:SS') And" & _
            "     Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-11 '||To_Char(��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS')))"
            
        '�ò�����䵱ʱ��ȡ���ְ��ŵĹҺ����
        strState = _
        "   Select A.ID as ����ID,B.�ѹ���,B.��Լ��" & _
        "   From (" & str�ҺŰ��� & ") A,���˹ҺŻ��� B" & _
        "   Where A.����ID = B.����ID And A.��ĿID = B.��ĿID" & _
        "               And Nvl(A.ҽ��ID,0)=Nvl(B.ҽ��ID,0) " & _
        "               And Nvl(A.ҽ������,'ҽ��')=Nvl(B.ҽ������,'ҽ��') " & _
        "               And (A.����=B.���� or B.���� is Null )  And B.����=[6]"
        
        If mblnAppointment Then
            str�ҺŰ��żƻ� = " " & _
                "             Select A.ID,A.ID as �ƻ�ID, A.����id, A.����, A.��Ŀid, A.������, A.����ʱ��, A. ����, A.��һ, A.�ܶ�, A.����, A.����, A.����," & _
                "                    A.���� , A.���﷽ʽ, A.��ſ���, B.�޺���, B.��Լ��, A.��Чʱ��, A.ʧЧʱ�� ,A.ҽ������,A.ҽ��ID " & _
                "             From �ҺŰ��żƻ� A, �Һżƻ����� B," & vbNewLine & _
                "                  (" & vbNewLine & _
                "                      Select Max(��Чʱ��) As ��Чʱ��, ����id" & _
                "                      From �ҺŰ��żƻ� " & vbNewLine & _
                "                      Where ���ʱ�� Is Not Null And  [5] Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
                "                          Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))  " & vbNewLine & _
                "                       Group By ����id" & vbNewLine & _
                "                   ) C" & _
                "             Where A.���ʱ�� Is Not Null And ([5] Between  A.��Чʱ��  And A.ʧЧʱ��)" & _
                "                   And A.ID = B.�ƻ�id(+) And " & vbNewLine & _
                "                   Decode(To_Char([5], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6'," & _
                "                  '����', '7', '����', Null) = B.������Ŀ(+) And A.��Чʱ�� = C.��Чʱ�� And A.����id = C.����id"

            strSQL = _
            " Select P.ID,0 as �ƻ�ID,P.���� ,P.����,P.����ID,P.��ĿID," & _
            "       P.ҽ��ID,P.ҽ������,P.�޺���,P.��Լ��,Nvl(P.��������,0) as ��������," & _
            "       P.����,P.��һ ,P.�ܶ� ,P.���� ,P.���� ,P.���� ,P.����,P.���﷽ʽ,P.��ſ���," & _
            "       Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL)  as �Ű� " & _
            " From (" & str�ҺŰ��� & ") P" & _
            " Where    Not Exists(Select 1 From �ҺŰ��żƻ� where ����ID=P.id And ([5] BETWEEN ��Чʱ��  and ʧЧʱ��)  And ���ʱ�� is not NULL  ) " & _
            "          And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=P.ID and [5] between ��ʼֹͣʱ�� and ����ֹͣʱ�� )" & _
            " Union ALL " & _
            " Select   C.ID,P.�ƻ�ID,C.����,C.����,C.����ID,P.��ĿID," & _
            "       P.ҽ��ID,P.ҽ������,P.�޺���,P.��Լ��,Nvl(C.��������,0) as ��������," & _
            "       P.����,P.��һ ,P.�ܶ� ,P.���� ,P.���� ,P.���� ,P.����,P.���﷽ʽ,P.��ſ���," & _
            "       Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL)  as �Ű� " & _
            " From (" & str�ҺŰ��żƻ� & ") P, �ҺŰ��� C" & _
            " Where P.����ID=C.ID  And C.ͣ������ Is  NULL  And Trunc(Sysdate)+Nvl(C.ԤԼ����," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]  " & _
            "           And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=C.ID and [5] between ��ʼֹͣʱ�� and ����ֹͣʱ�� )"
            strSQL = "(" & strSQL & ") P"
        Else
            strSQL = _
                        " (Select P.ID,0 as �ƻ�ID,P.���� ,P.����,P.����ID,P.��ĿID," & _
                        "       P.ҽ��ID,P.ҽ������,P.�޺���,P.��Լ��,Nvl(P.��������,0) as ��������," & _
                        "       P.����,P.��һ ,P.�ܶ� ,P.���� ,P.���� ,P.���� ,P.����,P.���﷽ʽ,P.��ſ���," & _
                        "       Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL) as �Ű� " & _
                        " From (" & str�ҺŰ��� & ") P "
            strSQL = strSQL & vbNewLine & "  ) P"
        End If
        
        strSQL = _
                    "Select Distinct " & _
                    "       P.ID,p.�ƻ�ID,P.���� as �ű�,P.����,P.����ID,B.���� As ����,P.��ĿID,C.���� As ��Ŀ," & _
                    "       P.ҽ��ID,P.ҽ������ as ҽ��,Nvl(A.�ѹ���,0) as �ѹ�,Nvl(A.��Լ��,0) as ��Լ," & _
                    "       P.�޺��� as �޺�,P.��Լ�� as ��Լ,Nvl(P.��������,0) as ����,Nvl(C.��Ŀ����,0) as ����," & _
                    "       P.���� as ��,P.��һ as һ,P.�ܶ� as ��,P.���� as ��,P.���� as ��,P.���� as ��,P.���� as ��," & _
                    "       Decode(P.���﷽ʽ,1,'ָ��',2,'��̬',3,'ƽ��',NULL) as ����,P.��ſ���,P.�Ű�" & _
                    " From " & strSQL & "," & vbCrLf & _
                    "           (" & strState & ") A,���ű� B,�շ���ĿĿ¼ C" & _
                    " Where P.ID=A.����ID(+) And Nvl(B.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And P.����ID=B.ID And P.��ĿID=C.ID" & strIF & strZero & _
                    "           And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & strWhere & _
                    "           And (Nvl(P.ҽ��ID,0)=0 Or Exists(Select 1 From ��Ա�� Q Where P.ҽ��ID=Q.ID And (Q.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.����ʱ�� Is Null)))" & _
                    " Order by " & str����
                    
        Set mrsPlan = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                UserInfo.����, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
    Else
        Exit Function
    End If
    If mrsPlan.RecordCount = 0 And mblnAppointment Then
        cboArrangeNo.Clear
        lblDeptName.Caption = ""
        If mblnInit Then MsgBox "��ǰû�п��õĹҺŰ��ţ����ڹҺŰ��Ź��������ú����ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    Set mcolArrangeNo = New Collection
    With cboArrangeNo
        .Clear
        Do While Not mrsPlan.EOF
            If Nvl(mrsPlan!ҽ��) = "" Then
                .AddItem "[" & Nvl(mrsPlan!�ű�) & "]" & Nvl(mrsPlan!��Ŀ)
            Else
                .AddItem "[" & Nvl(mrsPlan!�ű�) & "]" & Nvl(mrsPlan!��Ŀ) & "(" & Nvl(mrsPlan!ҽ��) & ")"
            End If
            mcolArrangeNo.Add Nvl(mrsPlan!�ű�)
            mrsPlan.MoveNext
        Loop
        If .ListCount <> 0 Then
            .ListIndex = 0
        Else
            MsgBox "��ǰû�п��õĹҺŰ��ţ����ڹҺŰ��Ź��������ú����ԣ�", vbInformation, gstrSysName
            Exit Function
        End If
'        Call GetActiveView
        Call ReadLimit
        Call LoadFeeItem(Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1)
'        If mblnAppointment Then
'            Select Case mViewMode
'                Case V_��ͨ�ŷ�ʱ��, v_ר�Һŷ�ʱ��
'                    cmdTime.Visible = True
'                Case Else
'                    cmdTime.Visible = False
'            End Select
'            Call InitRegTime
'        Else
'            cmdTime.Visible = False
'        End If

        lblDeptName.Caption = Nvl(mrsPlan!����)
    End With
    LoadRegPlans = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub ReadLimit()
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    mrsPlan.Filter = "�ű�='" & Get�ű� & "'"
    If mrsPlan.RecordCount = 0 Then Exit Sub
    If mblnAppointment Then
        If Nvl(mrsPlan!��Լ) = "" Then
            lblLimit.Caption = "��Լ:" & Nvl(mrsPlan!��Լ, 0)
        Else
            lblLimit.Caption = "��Լ:" & Nvl(mrsPlan!��Լ) & "  ��Լ:" & Nvl(mrsPlan!��Լ, 0)
        End If
    Else
        If Nvl(mrsPlan!�޺�) = "" Then
            lblLimit.Caption = "�ѹ�:" & Nvl(mrsPlan!�ѹ�, 0)
        Else
            lblLimit.Caption = "�޺�:" & Nvl(mrsPlan!�޺�) & "  �ѹ�:" & Nvl(mrsPlan!�ѹ�, 0)
        End If
    End If
    If Val(Nvl(mrsPlan!����)) = 0 Then
        lbl��.Visible = False
    Else
        lbl��.Visible = True
    End If
    Call GetYBInfo
End Sub

Private Function Get�ű�() As String
    If cboArrangeNo.Text = "" Then Exit Function
    Get�ű� = Mid(cboArrangeNo.Text, 2, InStr(cboArrangeNo.Text, "]") - 2)
End Function

Private Function GetActiveView()
    '�õ���ǰ�Һ�ҵ��  ��ȡ�������͵�����
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    Dim str����         As String
    Dim dat            As Date
    
    On Error GoTo errH
    str���� = Get�ű�
    If mblnAppointment Then
        dat = dtpDate.Value
    Else
        dat = gobjDatabase.CurrentDate
    End If
    
    strSQL = _
    "       Select   Havedata, ����id" & vbNewLine & _
    "       From (" & vbNewLine & _
    "               Select 1 As Havedata, b.Id As ����id " & vbNewLine & _
    "               From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
    "               Where B.����=[1] And A.����id = b.ID " & _
    "                And   Decode(To_Char([2], 'D'), '1', '����', '2'," & _
    "                   '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =a.���� " & vbNewLine & _
    "                       And Not Exists" & vbNewLine & _
    "                     (Select 1 From �ҺŰ��żƻ� C " & vbNewLine & _
    "                         Where c.����id = b.Id And c.���ʱ�� Is Not Null And [2] Between " & _
    "                               Nvl(c.��Чʱ��, [2]) And" & _
    "                          Nvl(c.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-MM-dd')))" & vbNewLine & _
    "               Union All " & vbNewLine & _
    "               Select 1 As Havedata, c.Id As ����id" & vbNewLine & _
    "               From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C,(" & vbNewLine & _
    "                   SELECT MAX(a.��Чʱ�� ) ��Ч FROM �ҺŰ��żƻ� a,�ҺŰ��� B  WHERE a.����Id=b.ID AND b.����=[1] AND a.���ʱ�� IS NOT NULL" & vbNewLine & _
    "             And [2] Between nvl(a.��Чʱ��,to_date('1900-01-01','yyyy-mm-dd')) And nvl(a.ʧЧʱ��,to_date('3000-01-01','yyyy-mm-dd'))" & vbNewLine & _
    "           ) D  " & vbNewLine & _
    "               Where  C.����=[1] And c.Id = b.����id And b.Id = a.�ƻ�id And b.��Чʱ��=d.��Ч And b.���ʱ�� Is Not Null" & _
    "                    And   Decode(To_Char([2], 'D'), '1', '����', '2'," & _
    "                   '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =a.���� " & vbNewLine & _
    "                       And [2] Between Nvl(b.��Чʱ��,[2]) And nvl(b.ʧЧʱ��,To_Date('3000-01-01', 'yyyy-MM-dd'))) B"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, str����, dat)
    If rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!��ſ���)) = 1 Then
       '*********************
       'ר�Һŷ�ʱ��
       '*********************
       mViewMode = v_ר�Һŷ�ʱ��

    ElseIf rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!��ſ���)) = 0 Then
       '*********************
       '��ͨ�ŷ�ʱ��
       '*********************
       mViewMode = V_��ͨ�ŷ�ʱ��

    ElseIf Val(Nvl(mrsPlan!��ſ���)) = 1 And Nvl(mrsPlan!�޺�) <> "" Then
       '*********************
       'ר�ҺŲ���ʱ��
       '*********************
       mViewMode = v_ר�Һ�

     Else
       '*********************
       '��ͨ��
       '*********************
       mViewMode = V_��ͨ��

    End If
    Set rsTmp = Nothing
Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
         Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function InitTimePlan() As Boolean
    '**************************************
    '����ʱ��
    '����ʱ���Ƿ���سɹ����Ƿ��з�ʱ��
    '**************************************
     Dim strSQL         As String
     Dim dateCur        As Date
     Dim strNO          As String
     Dim vRect          As RECT
    If Not mblnAppointment Then Exit Function
    strSQL = "Select Distinct a.��� As ID, To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
            "From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
            "Where a.����id = b.Id And b.���� = [1] And" & vbNewLine & _
            " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
            "      Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����'," & vbNewLine & _
            "             Null) = a.����(+) And Not Exists" & vbNewLine & _
            " (Select Count(1)" & vbNewLine & _
            "       From �Һ����״̬" & vbNewLine & _
            "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having" & vbNewLine & _
            "        Count(1) - a.�������� >= 0) And Not Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From �ҺŰ��żƻ� E" & vbNewLine & _
            "       Where e.����id = b.Id And e.���ʱ�� Is Not Null And" & vbNewLine & _
            "             [2] Between Nvl(e.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "             Nvl(e.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')))"
    strSQL = strSQL & " Union " & _
            "Select Distinct a.��� As ID, To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
            "From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C," & vbNewLine & _
            "     (Select Max(a.��Чʱ��) ��Ч" & vbNewLine & _
            "       From �ҺŰ��żƻ� A, �ҺŰ��� B" & vbNewLine & _
            "       Where a.����id = b.Id And b.���� = [1] And a.���ʱ�� Is Not Null And" & vbNewLine & _
            "             [2] Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "             Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'))) D" & vbNewLine & _
            "Where a.�ƻ�id = b.Id And b.����id = c.Id And c.���� = [1] And b.��Чʱ�� = d.��Ч And b.���ʱ�� Is Not Null And" & vbNewLine & _
            " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
            "      [2] Between Nvl(b.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "      Nvl(b.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And Not Exists" & vbNewLine & _
            " (Select Count(1)" & vbNewLine & _
            "       From �Һ����״̬" & vbNewLine & _
            "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having" & vbNewLine & _
            "        Count(1) - a.�������� >= 0) And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5'," & vbNewLine & _
            "                                           '����', '6', '����', '7', '����', Null) = a.����(+)" & vbNewLine & _
            "Order By ��ʼʱ��"


    dateCur = Format(dtpDate, "yyyy-mm-dd")
    If strSQL = "" Then Exit Function
    strNO = Get�ű�
    vRect = GetControlRect(dtpTime.hWnd)
    
    On Error GoTo errH
    
    Set mrsʱ��� = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "ԤԼʱ��ѡ��", False, "", "ԤԼʱ��ѡ��", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, strNO, dateCur)
    If mrsʱ��� Is Nothing Then Exit Function
    If mrsʱ���.EOF Then Exit Function
    InitTimePlan = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadFeeItem(ByVal lngItemID As Long, ByVal blnBook As Boolean)
    Dim strSQL As String, str���ʽ As String
    Dim i As Integer, j As Integer, dblTotal As Double
    Dim curӦ�� As Currency, curʵ�� As Currency
    
    If lngItemID = 0 Then Exit Sub
    If mrsPlan Is Nothing Then Exit Sub
    '����:1-���Һŷ��� 2-������� 3-������
    ReadRegistPrice lngItemID, blnBook, False, mstrFeeType, mrsItems, mrsInComes
    
    '126802�����ϴ���2018/6/7��ԤԼ�������ӷ�
    If Not mrsInfo Is Nothing And (mblnAppointment = False Or mty_Para.blnԤԼʱ�տ�) Then
        If mrsInfo.State = 1 Then
            If Not mrsInfo.EOF Then
                str���ʽ = Nvl(mrsInfo!ҽ�Ƹ��ʽ)
                If str���ʽ = "" Then str���ʽ = mstrDef���ʽ
                
                Call ReadExRegistPrice(mrsExpenses, mblnAppointPrice, Val(Nvl(mrsInfo!����ID)), mintInsure, Get�ű�, _
                        Nvl(mrsInfo!����), mstrGender, mstrAge, Nvl(mrsInfo!���֤��), mstrFeeType, str���ʽ)
            End If
        End If
    End If
    
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = "0.00"
    lblPayMoney.Caption = "0.00"
    dblTotal = 0
    With vsfMoney
        If mrsItems.RecordCount = 0 Then Exit Sub
        mrsItems.MoveFirst
        For i = 1 To mrsItems.RecordCount
            .RowData(.Rows - 1) = Nvl(mrsItems!��ĿID)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Nvl(mrsItems!��Ŀ����)
            mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
            curӦ�� = 0: curʵ�� = 0
            For j = 1 To mrsInComes.RecordCount
                curӦ�� = curӦ�� + mrsInComes!Ӧ��
                curʵ�� = curʵ�� + mrsInComes!ʵ��
                mrsInComes.MoveNext
            Next j
            .TextMatrix(.Rows - 1, .ColIndex("Ӧ�ս��")) = Format(curӦ��, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("ʵ�ս��")) = Format(curʵ��, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(mrsItems!����)
            
            dblTotal = dblTotal + Val(.TextMatrix(.Rows - 1, vsfMoney.ColIndex("ʵ�ս��")))
            .Rows = .Rows + 1
            mrsItems.MoveNext
        Next i
        
        If Not mrsExpenses Is Nothing Then
            If mrsExpenses.RecordCount > 0 Then mrsExpenses.MoveFirst
            Do While Not mrsExpenses.EOF
                .RowData(.Rows - 1) = Nvl(mrsExpenses!��ĿID)
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Nvl(mrsExpenses!��Ŀ����)
                .TextMatrix(.Rows - 1, .ColIndex("Ӧ�ս��")) = Format(mrsExpenses!Ӧ��, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("ʵ�ս��")) = Format(mrsExpenses!ʵ��, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(mrsExpenses!����)
                
                dblTotal = dblTotal + Val(.TextMatrix(.Rows - 1, vsfMoney.ColIndex("ʵ�ս��")))
                .Rows = .Rows + 1
                mrsExpenses.MoveNext
            Loop
        End If
    End With
    If vsfMoney.Rows > 2 Then vsfMoney.Rows = vsfMoney.Rows - 1
    vsfMoney.RowHeightMin = 350
    lblTotal.Caption = Format(dblTotal, "0.00")
    lblPayMoney.Caption = Format(dblTotal, "0.00")
    lblRoomName.Caption = gstrRooms
End Sub


Private Function GetSNState(str�ű� As String, datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    Dim strSQL           As String
    Dim datStart         As Date
    Dim datEnd           As Date
    On Error GoTo errH
    datStart = CDate(Format(datThis, "yyyy-MM-dd"))
    datEnd = DateAdd("s", -1, DateAdd("d", 1, datStart))
    strSQL = "    " & vbNewLine & " Select ���,״̬,����Ա����,Nvl(ԤԼ,0) as ԤԼ,TO_Char(����,'hh24:mi:ss') as ����  "
    strSQL = strSQL & vbNewLine & " From �Һ����״̬ "
    strSQL = strSQL & vbNewLine & " Where ����=[1]"
    strSQL = strSQL & vbNewLine & IIf(datThis = CDate(0), " And ���� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 ", " And ���� Between  [2] And [3]")
    strSQL = strSQL & vbNewLine & IIf(lngSN > 0, " And ���=[4]", "")
    Set GetSNState = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�, datStart, datEnd, lngSN)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function zlGet��ǰ���ڼ�(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���������ڼ�
    '����:���˺�
    '����:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bln��ǰ���� As Boolean, strTemp As String
    If strDate = "" Then
        strSQL = "Select Decode(To_Char(Sysdate,'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��',NULL) as ����  From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "Select Decode(To_Char([1],'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��','') As ���� From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strDate))
    End If
    
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!����)
    zlGet��ǰ���ڼ� = strTemp
End Function

Private Sub GetYBInfo()
'���ܣ�'��ȡҽ��ͳ���������
    Dim strInfo As String, i As Long, j As Long, lng����ID As Long
    
    If mRegistFeeMode = EM_RG_���� Then Exit Sub
    If mstrYBPati <> "" Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    
    If mintInsure <> 0 And mstrYBPati <> "" Then
        If Not mrsItems Is Nothing Then
            mrsItems.MoveFirst
            For i = 1 To mrsItems.RecordCount
                mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                For j = 1 To mrsInComes.RecordCount
                    strInfo = gclsInsure.GetItemInsure(lng����ID, mrsItems!��ĿID, mrsInComes!ʵ��, True, mintInsure)
                    If strInfo <> "" Then
                        mrsItems!������Ŀ�� = Val(Split(strInfo, ";")(0))
                        mrsItems!���մ���ID = Val(Split(strInfo, ";")(1))
                        mrsItems!���ձ��� = CStr(Split(strInfo, ";")(3))
                        mrsInComes!ͳ���� = Format(Val(Split(strInfo, ";")(2)), "0.00")
                    End If
                    mrsInComes.MoveNext
                Next
                mrsItems.MoveNext
            Next
        End If
        
        If Not mrsExpenses Is Nothing Then
            mrsExpenses.MoveFirst
            For j = 1 To mrsExpenses.RecordCount
                strInfo = gclsInsure.GetItemInsure(lng����ID, mrsExpenses!��ĿID, mrsExpenses!ʵ��, True, mintInsure)
                If strInfo <> "" Then
                    mrsExpenses!������Ŀ�� = Val(Split(strInfo, ";")(0))
                    mrsExpenses!���մ���ID = Val(Split(strInfo, ";")(1))
                    mrsExpenses!���ձ��� = CStr(Split(strInfo, ";")(3))
                    mrsExpenses!ͳ���� = Format(Val(Split(strInfo, ";")(2)), "0.00")
                End If
                mrsExpenses.MoveNext
            Next
        End If
    End If
End Sub
