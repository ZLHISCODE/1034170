VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmDistRoomRegist 
   Caption         =   "�������Һ�"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10440
   Icon            =   "frmDistRoomRegist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10440
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtGender 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   3345
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   525
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   9090
      TabIndex        =   10
      Top             =   6420
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7890
      TabIndex        =   9
      Top             =   6420
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5805
      TabIndex        =   44
      Top             =   6420
      Width           =   1100
   End
   Begin VB.Frame fraPay 
      Height          =   750
      Left            =   5625
      TabIndex        =   41
      Top             =   5550
      Width           =   4740
      Begin VB.TextBox txtPayMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   210
         Width           =   1635
      End
      Begin VB.ComboBox cboPayMode 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   1500
      End
      Begin VB.Label lblPayMode 
         AutoSize        =   -1  'True
         Caption         =   "֧����ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   270
         Width           =   1200
      End
   End
   Begin VB.Frame fraTotal 
      Height          =   1095
      Left            =   5625
      TabIndex        =   38
      Top             =   4470
      Width           =   4740
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   3465
         TabIndex        =   40
         Top             =   330
         Width           =   960
      End
      Begin VB.Label lblSum 
         Caption         =   "�� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   15
         TabIndex        =   39
         Top             =   135
         Width           =   645
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   3165
      Left            =   5625
      TabIndex        =   30
      Top             =   1305
      Width           =   4740
      Begin VB.ComboBox cboAppointStyle 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   2760
         Width           =   1665
      End
      Begin VB.ComboBox cboRemark 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3165
         TabIndex        =   49
         Top             =   2760
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   330
         Left            =   3165
         TabIndex        =   5
         Top             =   615
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         Format          =   93782018
         CurrentDate     =   42121
      End
      Begin VB.CheckBox chkBook 
         Caption         =   " ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   7
         Top             =   1065
         Width           =   1485
      End
      Begin VB.TextBox txtRegistTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   4
         EndProperty
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3165
         TabIndex        =   36
         Top             =   615
         Width           =   1500
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1500
      End
      Begin VB.ComboBox cboRoom 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   1500
      End
      Begin VB.ComboBox cboDoctor 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   615
         Width           =   1500
      End
      Begin VB.TextBox txtArrangeNO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   570
         TabIndex        =   2
         Top             =   180
         Width           =   1500
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   1290
         Left            =   75
         TabIndex        =   37
         Top             =   1425
         Width           =   4575
         _cx             =   8070
         _cy             =   2275
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDistRoomRegist.frx":058A
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
      Begin VB.Label lblAppointStyle 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   52
         Top             =   2820
         Width           =   840
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2730
         TabIndex        =   50
         Top             =   2820
         Width           =   420
      End
      Begin VB.Label lblRegistTime 
         AutoSize        =   -1  'True
         Caption         =   "�Һ�ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2220
         TabIndex        =   35
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2640
         TabIndex        =   34
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblRoom 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   33
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   32
         Top             =   675
         Width           =   420
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "�ű�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   31
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.Frame fraTime 
      Height          =   5805
      Left            =   30
      TabIndex        =   23
      Top             =   975
      Width           =   5520
      Begin VB.ComboBox cboDoctorFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmDistRoomRegist.frx":0625
         Left            =   4185
         List            =   "frmDistRoomRegist.frx":0627
         TabIndex        =   27
         Top             =   165
         Width           =   1275
      End
      Begin VB.ComboBox cboDeptFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2445
         TabIndex        =   26
         Top             =   165
         Width           =   1275
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetailTime 
         Height          =   2370
         Left            =   60
         TabIndex        =   28
         Top             =   3360
         Width           =   5385
         _cx             =   9499
         _cy             =   4180
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   18
         FixedRows       =   0
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
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Left            =   510
         TabIndex        =   47
         Top             =   165
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93782019
         CurrentDate     =   42335
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfArrange 
         Height          =   2700
         Left            =   60
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   570
         Width           =   5385
         _cx             =   9499
         _cy             =   4762
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDistRoomRegist.frx":0629
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
         ExplorerBar     =   1
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
      Begin VB.Label lblDoctorFilter 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3765
         TabIndex        =   25
         Top             =   225
         Width           =   420
      End
      Begin VB.Label lblDeptFilter 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2010
         TabIndex        =   24
         Top             =   225
         Width           =   420
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   48
         Top             =   225
         Width           =   420
      End
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Left            =   -60
      TabIndex        =   22
      Top             =   945
      Width           =   11000
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   330
      Left            =   600
      TabIndex        =   21
      Top             =   525
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      Appearance      =   2
      IDKindStr       =   "��|��������￨|0|0|0|0|0|;ҽ|ҽ����|0|0|0|0|0|;��|���֤��|1|0|0|0|0|;��|�����|0|0|0|0|0|"
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   11.25
      FontName        =   "����"
      IDKind          =   -1
      DefaultCardType =   "0"
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4665
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   525
      Width           =   705
   End
   Begin VB.TextBox txtClinic 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   525
      Width           =   1470
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   525
      Width           =   1500
   End
   Begin VB.TextBox txtFeeType 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   525
      Width           =   1185
   End
   Begin VB.ComboBox cboNO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8625
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   90
      Width           =   1755
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
      TabIndex        =   46
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "����Ԥ�����:0.00     "
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5625
      TabIndex        =   29
      Top             =   1065
      Width           =   2310
   End
   Begin VB.Label lblFeeType 
      AutoSize        =   -1  'True
      Caption         =   "�ѱ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8130
      TabIndex        =   17
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   16
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblGender 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2835
      TabIndex        =   15
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4170
      TabIndex        =   14
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblClinic 
      AutoSize        =   -1  'True
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5685
      TabIndex        =   13
      Top             =   585
      Width           =   630
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "���ݺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7935
      TabIndex        =   0
      Top             =   150
      Width           =   630
   End
End
Attribute VB_Name = "frmDistRoomRegist"
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

Private mblnCard As Boolean, mblnStartFactUseType As Boolean
Private mstrYBPati As String, mlng�Һ�ID As Long, mlng����ID As Long
Private mblnOlnyBJYB As Boolean, mblnSharedInvoice As Boolean
Private mstr���� As String, mblnChangeFeeType As Boolean, mblnInit As Boolean
Private mstrPassWord As String, mstrInsure As String, mintSysAppLimit As Integer
Private mstrDeptIDs As String, mlngRow As Long, msngTime As Single
Private Const SNCOLS = 10
Private Const SnArgCols = 7

Private mrsInfo As ADODB.Recordset
Private mrsPlan As ADODB.Recordset
Private mrsSNState As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsItems As ADODB.Recordset
Private mrsInComes As ADODB.Recordset
Private mrsExpenses As ADODB.Recordset '��¼���ӷ���Ŀ(����������Ϣ)
Private mrsʱ��� As ADODB.Recordset
Private mcolCardPayMode As Collection
Private mcur������� As Currency
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mintIDKind As Integer
Private mintInsure As Integer
Private mdatLast As Date
Private mblnChangeByCode As Boolean, mblnFilterChange As Boolean
Private mstrCardNO As String
Private mcur����͸֧ As Currency
Private mblnAppointPrice As Boolean
Private mblnAppointment As Boolean  'ԤԼ�Һ�
Private mstr������� As String

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
    'bln�������Ұ��� As Boolean
    int�Һŷ�Ʊ��ӡ As Integer
    int�Һ�ƾ����ӡ As Integer
    intԤԼ�ҺŴ�ӡ As Integer
    bln������ѡ�� As Boolean
    lngԤԼ��Чʱ�� As Long
    bln�����շ�Ʊ�� As Boolean
    blnԤԼʱ�տ� As Boolean
    bln�˺����� As Boolean
    bln������֤ As Boolean
End Type

Private mty_Para As ty_ModulePara
Private mstr����IDs As String

Public Sub zlShowMe(ByVal frmMain As Object, ByVal lngModul As Long, ByVal strDeptIDs As String, ByRef strOutNO As String, ByVal blnAppointment As Boolean)
    mlngModul = lngModul
    mstrDeptIDs = strDeptIDs
    mblnAppointment = blnAppointment
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
        .blnԤԼʱ�տ� = Val(gobjDatabase.GetPara("ԤԼʱ�տ�", glngSys, 9000, "0")) = 1
        .bln�����շ�Ʊ�� = Val(gobjDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121)) = 1
        .bln�˺����� = Val(gobjDatabase.GetPara("�����������Һ�", glngSys, 1111)) = 1
        .bln�Һű���ˢ�� = Val(gobjDatabase.GetPara("�Һű���ˢ��", glngSys, 9000)) = 1
        .bln������֤ = Val(gobjDatabase.GetPara(28, glngSys)) <> 0
        If .blnĬ������ժҪ Then
            cboRemark.TabStop = True
        Else
            cboRemark.TabStop = False
        End If
        If .blnĬ�Ϲ����� Then
            chkBook.Value = 1
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
    '�շѺ͹ҺŹ���Ʊ��
    mblnSharedInvoice = gobjDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121) = "1"
    mintSysAppLimit = Val(gobjDatabase.GetPara("�Һ�����ԤԼ����", glngSys))
    '���ع��ùҺ�����ID
    If mblnSharedInvoice Then
        mlng�Һ�ID = Val(gobjDatabase.GetPara("�����շ�Ʊ������", glngSys, 1121, ""))
    Else
        mlng�Һ�ID = Val(gobjDatabase.GetPara("���ùҺ�Ʊ������", glngSys, 1111, ""))
    End If
    If mlng�Һ�ID > 0 Then
        If Not ExistShareBill(mlng�Һ�ID, IIf(mblnSharedInvoice, 1, 4)) Then
            If mblnSharedInvoice Then
                gobjDatabase.SetPara "�����շ�Ʊ������", "0", glngSys, 1121
            Else
                gobjDatabase.SetPara "���ùҺ�Ʊ������", "0", glngSys, 1111
            End If
            mlng�Һ�ID = 0
        End If
    End If
    'Ʊ���ϸ����
    strValue = gobjDatabase.GetPara(24, glngSys, , "00000")
    gblnBill�Һ� = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    If mblnSharedInvoice Then
        '�Һ�������Ʊ��:42703
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    If mblnAppointment Then
        dtpDate.minDate = gobjDatabase.CurrentDate
        dtpDate.Value = gobjDatabase.CurrentDate
    End If
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
    Dim objCard As Card, strTemp As String
    Dim lngCardID As Long
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


Private Sub cboDeptFilter_Click()
    mblnFilterChange = True
    LoadRegPlans (True)
    mblnFilterChange = False
    If mrsPlan.RecordCount <> 0 Then Call vsfArrange_EnterCell
End Sub

Private Sub cboDeptFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intInputType As Integer, i As Integer
    If KeyCode = 13 Then
        If cboDeptFilter.Text = "" Then
            cboDeptFilter.ListIndex = 0
            Exit Sub
        End If
        If IsNumeric(cboDeptFilter.Text) Then
            intInputType = 0
        ElseIf gobjCommFun.IsCharAlpha(cboDeptFilter.Text) Then
            intInputType = 1
        Else
            intInputType = 2
        End If
        For i = 1 To cboDeptFilter.ListCount - 1
            Select Case intInputType
            Case 0, 2
                If cboDeptFilter.List(i) Like "*" & cboDeptFilter.Text & "*" Then
                    cboDeptFilter.ListIndex = i
                    Exit For
                End If
            Case 1  '�������ȫ��ĸ
                '�����:116582,����,2017/12/5,ͨ�������ȡ����ʱ����ʾ'����ʱ����'9'���±�Խ��'
                If UCase(gobjCommFun.zlGetSymbol(cboDeptFilter.List(i))) Like "*" & UCase(cboDeptFilter.Text) & "*" Then
                    cboDeptFilter.ListIndex = i
                    Exit For
                End If
            End Select
        Next i
    End If
End Sub

Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboDoctorFilter_Click()
    mblnFilterChange = True
    LoadRegPlans (True)
    mblnFilterChange = False
    If mrsPlan.RecordCount <> 0 Then Call vsfArrange_EnterCell
End Sub

Private Sub cboDoctorFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intInputType As Integer, i As Integer
    If KeyCode = 13 Then
        If cboDoctorFilter.Text = "" Then
            cboDoctorFilter.ListIndex = 0
            Exit Sub
        End If
        If IsNumeric(cboDoctorFilter.Text) Then
            intInputType = 0
        ElseIf gobjCommFun.IsCharAlpha(cboDoctorFilter.Text) Then
            intInputType = 1
        Else
            intInputType = 2
        End If
        For i = 1 To cboDoctorFilter.ListCount - 1
            Select Case intInputType
            Case 0, 2
                If cboDoctorFilter.List(i) Like "*" & cboDoctorFilter.Text & "*" Then
                    cboDoctorFilter.ListIndex = i
                    Exit For
                End If
            Case 1  '�������ȫ��ĸ
                If UCase(gobjCommFun.zlGetSymbol(cboDoctorFilter.List(i))) Like "*" & UCase(cboDoctorFilter.Text) & "*" Then
                    cboDoctorFilter.ListIndex = i
                    Exit For
                End If
            End Select
        Next i
    End If
End Sub

Private Sub cboPayMode_Click()
    If MCPAR.���ղ����� And cboPayMode.Text = mstrInsure Then
        chkBook.Enabled = False
        chkBook.Value = 0
    Else
        chkBook.Enabled = True
    End If
End Sub

Private Sub cboPayMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
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

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub chkBook_Click()
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")) = "" Then Exit Sub
    Call LoadFeeItem(Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(1)), chkBook.Value = 1)
End Sub

Private Sub chkBook_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
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
    str���� = Trim(txtAge.Text)

   If gobjSquare.objSquareCard.zlBrushCard(Me, glngModul, rsMoney, lngҽ�ƿ����ID, bln���ѿ�, _
    txtPatient.Text, NeedName(txtGender.Text), str����, dblMoney, mstrCardNO, mstrPassWord, _
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
    Dim lngSN As Long, lng�Һſ���ID As Long, lng����ID As Long, byt���� As Byte, blnAppointPrint As Boolean
    Dim lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean, blnNoDoc As Boolean, strBalanceStyle As String
    Dim blnTrans As Boolean, blnNotCommit As Boolean, strAdvance As String
    Dim lngҽ��ID As Long, blnOneCard As Boolean, rsTmp As ADODB.Recordset, str���ʽ As String
    Dim strNotValiedNos As String
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
    
    mstr����IDs = ""
    If Val(curԤ��) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!����ID), Val(curԤ��), mlngModul, 1, , mty_Para.bln������֤) Then Exit Sub
    End If
    
    strSQL = "Select ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, Nvl(mrsInfo!ҽ�Ƹ��ʽ))
    If rsTmp.RecordCount <> 0 Then
        mstr������� = Nvl(rsTmp!����)
    Else
        mstr������� = mstrDef�������
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
        str����ʱ�� = "To_Date('" & Format(dtpDate.Value, "YYYY-MM-DD") & " " & Format(dtpTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    Else
        str����ʱ�� = "To_Date('" & Format(gobjDatabase.CurrentDate, "YYYY-MM-DD") & " " & Format(dtpTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    End If
    lng�Һſ���ID = Val(vsfArrange.RowData(vsfArrange.Row))
    If mRegistFeeMode = EM_RG_���� Then
        lng����ID = gobjDatabase.GetNextId("���˽��ʼ�¼")
    End If
    byt���� = IIf(Check����(Val(mrsInfo!����ID), lng�Һſ���ID), 1, 0)
    
    'Ʊ�ݴ���
    If mRegistFeeMode = EM_RG_���� Then
        str����NO = gobjDatabase.GetNextNo(13)
    End If
    If vsfDetailTime.Visible Then
        If mViewMode = v_ר�Һŷ�ʱ�� Then
            lngSN = Val(Getʱ��(vsfDetailTime.Row, vsfDetailTime.Col))
        End If
        If mViewMode = v_ר�Һ� Then
            lngSN = Val(vsfDetailTime.TextMatrix(vsfDetailTime.Row, vsfDetailTime.Col))
        End If
    Else
        lngSN = 0
    End If
    strNO = gobjDatabase.GetNextNo(12)
    
    mrsItems.Filter = ""
    If cboDoctor.ListIndex = -1 Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(cboDoctor.ItemData(cboDoctor.ListIndex))
    End If
    
    k = 1: mrsItems.MoveFirst
    For i = 1 To mrsItems.RecordCount
        int�۸񸸺� = k
        mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
        For j = 1 To mrsInComes.RecordCount
            strSQL = _
            "zl_���˹Һż�¼_INSERT(" & ZVal(Nvl(mrsInfo!����ID)) & "," & IIf(txtClinic.Text = "", "NULL", txtClinic.Text) & ",'" & txtPatient.Text & "','" & NeedName(txtGender.Text) & "'," & _
                     "'" & txtAge.Text & "','" & mstr������� & "','" & txtFeeType.Text & "','" & strNO & "'," & _
                     "'" & IIf(blnInvoicePrint = False, "", strFactNO) & "'," & k & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & IIf(mrsItems!���� = 2, 1, "NULL") & "," & _
                     "'" & mrsItems!��� & "'," & mrsItems!��ĿID & "," & mrsItems!���� & "," & mrsInComes!���� & "," & _
                     mrsInComes!������ĿID & ",'" & mrsInComes!�վݷ�Ŀ & "','" & IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), "") & "'," & _
                     IIf(mRegistFeeMode = EM_RG_����, 0, mrsInComes!Ӧ��) & "," & IIf(mRegistFeeMode = EM_RG_����, 0, mrsInComes!ʵ��) & "," & _
                     lng�Һſ���ID & "," & UserInfo.����ID & "," & IIf(mrsItems!ִ�п���ID = 0, lng�Һſ���ID, mrsItems!ִ�п���ID) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                     str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                     "'" & NeedName(cboDoctor.Text) & "'," & ZVal(lngҽ��ID) & "," & IIf(mrsItems!���� = 3, 1, IIf(mrsItems!���� = 4, 2, 0)) & "," & IIf(lbl��.Visible, 1, 0) & "," & _
                     "'" & txtArrangeNO.Text & "','" & cboRoom.Text & "'," & ZVal(lng����ID) & "," & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng����ID)) & "," & _
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
            strSQL = strSQL & "0" & ","
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
            If txtArrangeNO.Text <> "" And k = 1 Then
                If Nvl(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("ҽ��"))) = "" Then blnNoDoc = True
                strSQL = "zl_���˹ҺŻ���_Update("
                '  ҽ������_In   �ҺŰ���.ҽ������%Type,
                strSQL = strSQL & IIf(blnNoDoc, "Null,", "'" & NeedName(cboDoctor.Text) & "',")
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
                strSQL = strSQL & "'" & txtArrangeNO.Text & "')"
                Call zlAddArray(cllProAfter, strSQL)
            End If
            If mRegistFeeMode = EM_RG_���� Then
                strSQL = _
                "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & k & "," & ZVal(Nvl(mrsInfo!����ID)) & ",NULL," & _
                         IIf(txtClinic.Text = "", "NULL", txtClinic.Text) & ",'" & mstr������� & "'," & _
                         "'" & txtPatient.Text & "','" & txtGender.Text & "','" & txtAge.Text & "'," & _
                         "'" & txtFeeType.Text & "',NULL," & lng�Һſ���ID & "," & _
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
                txtClinic.Text, lngSN, str����ʱ��, str�Ǽ�ʱ��, Not blnInvoicePrint, IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), ""), _
                cboRoom.Text, lng����ID, blnBalance, byt����) = False Then
        Exit Sub
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
    
    If blnSlipPrint Or blnInvoicePrint Then
        '��¼��ӡ��ƾ��
        gstrSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    Call ReloadPage
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
                    Optional ByVal blnBalance As Boolean, Optional ByVal byt���� As Byte) As Boolean
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
        
    lng�Һſ���ID = Val(vsfArrange.RowData(vsfArrange.Row))
    strҽ�� = NeedName(cboDoctor.Text)
    If cboDoctor.ListIndex = -1 Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(cboDoctor.ItemData(cboDoctor.ListIndex))
    End If
    If lngNoSort <> 0 Then lngNoSort = lngNoSort - 1
    
    For i = 1 To rsExpenses.RecordCount
        lngNoSort = lngNoSort + 1
        
        strSQL = _
            "zl_���˹Һż�¼_INSERT(" & ZVal(lng����ID) & "," & IIf(str����� = "", "NULL", str�����) & ",'" & txtPatient.Text & "','" & NeedName(txtGender.Text) & "'," & _
            "'" & txtAge.Text & "','" & mstr������� & "','" & txtFeeType.Text & "','" & strRegNo & "'," & _
            "''," & lngNoSort & "," & IIf(lngPre��ĿID = rsExpenses!��ĿID, int�۸񸸺�, "NULL") & ",NULL," & _
            "'" & rsExpenses!��� & "'," & rsExpenses!��ĿID & "," & rsExpenses!���� & "," & rsExpenses!���� & "," & _
            rsExpenses!������ĿID & ",'" & rsExpenses!�վݷ�Ŀ & "','" & str���㷽ʽ & "'," & _
            IIf(bln��Ϊ���۵�, 0, rsExpenses!Ӧ��) & "," & IIf(bln��Ϊ���۵�, 0, rsExpenses!ʵ��) & "," & _
            lng�Һſ���ID & "," & UserInfo.����ID & "," & IIf(rsExpenses!ִ�п���ID = 0, lng�Һſ���ID, rsExpenses!ִ�п���ID) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
            str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
            "'" & strҽ�� & "'," & ZVal(lngҽ��ID) & "," & 0 & "," & IIf(lbl��.Visible, 1, 0) & "," & _
            "'" & txtArrangeNO.Text & "','" & strRoom & "'," & ZVal(lng����ID) & "," & IIf(blnNoPrint, "NULL", ZVal(mlng����ID)) & "," & _
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
            strSQL = strSQL & "'" & mstrCardNO & "',"
            '������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSQL = strSQL & " NULL,"
            '����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
            strSQL = strSQL & " NULL,"
            '������λ_In   ����Ԥ����¼.������λ%Type := Null
            strSQL = strSQL & " NULL,"
            '  ��������_In   Number:=0
            strSQL = strSQL & "0" & ","
            '  ����_IN       ���˹Һż�¼.����%type:=null,
            strSQL = strSQL & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  ����ģʽ_IN   NUMBER :=0,
            strSQL = strSQL & IIf(mPatiChargeMode = EM_�����ƺ����, 1, 0) & ","
            '  ���ʷ���_IN Number:=0
            strSQL = strSQL & IIf(mRegistFeeMode = EM_RG_����, 1, 0) & ","
            '  �˺�����_IN Number:=1
            strSQL = strSQL & IIf(mty_Para.bln�˺�����, 1, 0) & ")"
            Call zlAddArray(cllPro, strSQL)
            
        
        If bln��Ϊ���۵� Then
            strSQL = _
            "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & lngNoSort & "," & ZVal(lng����ID) & ",NULL," & _
                     IIf(str����� = "", "NULL", str�����) & ",'" & mstr������� & "'," & _
                     "'" & txtPatient.Text & "','" & txtGender.Text & "','" & txtAge.Text & "'," & _
                     "'" & txtFeeType.Text & "',NULL," & lng�Һſ���ID & "," & _
                     IIf(lng�Һſ���ID <> 0, lng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "',NULL," & _
                     rsExpenses!��ĿID & ",'" & rsExpenses!��� & "','" & rsExpenses!���㵥λ & "'," & _
                     "NULL,1," & rsExpenses!���� & ",NULL," & IIf(rsExpenses!ִ�п���ID = 0, lng�Һſ���ID, rsExpenses!ִ�п���ID) & "," & _
                     IIf(lngPre��ĿID = rsExpenses!��ĿID, int�۸񸸺�, "NULL") & "," & _
                     rsExpenses!������ĿID & ",'" & rsExpenses!�վݷ�Ŀ & "'," & rsExpenses!���� & "," & _
                     rsExpenses!Ӧ�� & "," & rsExpenses!ʵ�� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "',NULL,'�Һ�:" & strRegNo & "')"
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
    Call LoadRegPlans(False)
    Call ClearPatient
    Call ClearRegInfo
End Sub

Private Sub ClearRegInfo()
    mblnChangeByCode = True
    txtArrangeNO.Text = ""
    mblnChangeByCode = False
    txtDept.Text = ""
    cboDoctor.Clear
    cboRoom.Clear
    cboRemark.Text = ""
    chkBook.Value = IIf(mty_Para.blnĬ�Ϲ�����, 1, 0)
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    vsfArrange.Height = vsfDetailTime.Top + vsfDetailTime.Height - vsfArrange.Top
    vsfDetailTime.Visible = False
    lbl��.Visible = False
    txtPatient.SetFocus
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
        txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "�޷�ȷ��������Ϣ,����ѡ��һ�����ˣ�", vbInformation, gstrSysName
        txtPatient.SetFocus
        Exit Function
    End If
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")) = "" Or txtArrangeNO.Text = "" Then
        txtArrangeNO.SetFocus
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If GetRegistMoney < 0 Then
        MsgBox "�Һŷ��ò���Ϊ����������Һ���Ŀ��", vbInformation, gstrSysName
        If txtArrangeNO.Visible And txtArrangeNO.Enabled Then txtArrangeNO.SetFocus
        Exit Function
    End If
    
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")) <> txtArrangeNO.Text Then
        txtArrangeNO.SetFocus
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheck��Լ���޺���(txtArrangeNO.Text) = False Then Exit Function
    If vsfDetailTime.Visible Then
        If vsfDetailTime.Row > vsfDetailTime.Rows - 1 Or vsfDetailTime.Col > vsfDetailTime.Cols - 1 Then
            MsgBox "ѡ������Ч��ţ����飡", vbInformation, gstrSysName
            Exit Function
        End If
        If vsfDetailTime.Cell(flexcpForeColor, vsfDetailTime.Row, vsfDetailTime.Col) <> vbBlack Or vsfDetailTime.Cell(flexcpBackColor, vsfDetailTime.Row, vsfDetailTime.Col) = -2147483633 Then
            MsgBox "ѡ������Ч��ţ����飡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If cboPayMode.Text = "" And cboPayMode.Visible Then
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

Private Sub SetControl()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTemp As String, i As Integer
    If mblnAppointment Then
        Me.Caption = "����ԤԼ"
        If mty_Para.blnԤԼʱ�տ� Then
            fraPay.Visible = True
        Else
            fraPay.Visible = False
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
        Me.Caption = "����Һ�"
        lblDeptFilter.Left = lblDate.Left
        cboDeptFilter.Left = dtpDate.Left
        cboDeptFilter.Width = 2055
        lblDoctorFilter.Left = 2805
        cboDoctorFilter.Left = 3315
        cboDoctorFilter.Width = 2055
        lblDate.Visible = False
        dtpDate.Visible = False
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
        lblRemark.Left = lblRoom.Left
        cboRemark.Left = 570
        cboRemark.Width = 4110
        If mty_Para.byt�Һ�ģʽ = 0 Then
            fraPay.Visible = True
        Else
            fraPay.Visible = False
        End If
        
    End If
End Sub

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

Private Sub dtpDate_Change()
    cboDeptFilter.Text = ""
    cboDoctorFilter.Text = ""
    cboDeptFilter.ListIndex = -1
    cboDoctorFilter.ListIndex = -1
    Call LoadRegPlans(False)
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnInit Then
        mblnInit = False
    End If
End Sub

Private Sub Form_Load()
    Err = 0
    mblnInit = True
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitData
    Call InitPara
    Call InitIDKind
    Call GetAllҽ��
    Call RestoreWinState(Me, App.ProductName)
    Call LoadRegPlans(False)
    Call InitFilter
    Call LoadPayMode
    Call SetControl
    glngFormW = 10680: glngFormH = 7425
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If

    vsfArrange.Height = vsfDetailTime.Top + vsfDetailTime.Height - vsfArrange.Top
    vsfDetailTime.Visible = False
End Sub

Private Sub InitFilter()
    Dim strExists
    On Error GoTo errH
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    If mrsPlan.RecordCount = 0 Then Exit Sub
    mrsPlan.MoveFirst
    strExists = ","
    cboDeptFilter.AddItem " "
    Do While Not mrsPlan.EOF
        If InStr(strExists, "," & Nvl(mrsPlan!����, "") & ",") = 0 Then
            cboDeptFilter.AddItem Nvl(mrsPlan!����)
            strExists = strExists & Nvl(mrsPlan!����) & ","
        End If
        mrsPlan.MoveNext
    Loop
    
    mrsPlan.MoveFirst
    strExists = ","
    cboDoctorFilter.AddItem ""
    Do While Not mrsPlan.EOF
        If InStr(strExists, "," & Nvl(mrsPlan!ҽ��, "") & ",") = 0 And Not IsNull(mrsPlan!ҽ��) Then
            cboDoctorFilter.AddItem Nvl(mrsPlan!ҽ��)
            strExists = strExists & Nvl(mrsPlan!ҽ��) & ","
        End If
        mrsPlan.MoveNext
    Loop
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraSplit.Width = Me.Width + 300
    fraInfo.Height = Me.Height / 2 - 500
    lblRemark.Top = fraInfo.Height - 400
    cboRemark.Top = fraInfo.Height - 450
    lblAppointStyle.Top = lblRemark.Top
    cboAppointStyle.Top = cboRemark.Top
    vsfMoney.Height = fraInfo.Height - vsfMoney.Top - 500
    cboNO.Left = Me.Width - 2100
    lblNO.Left = Me.Width - 3000
    fraPay.Left = Me.Width - 5060
    fraTotal.Left = Me.Width - 5060
    fraInfo.Left = Me.Width - 5060
    lblMoney.Left = fraInfo.Left
    fraTime.Width = fraInfo.Left - 90
    cmdHelp.Left = fraInfo.Left + 150
    cmdOK.Left = Me.Width - 2750
    cmdCancel.Left = Me.Width - 1500
    fraTime.Height = Me.Height - fraTime.Top - 600
    vsfArrange.Width = fraTime.Width - 150
    vsfDetailTime.Width = fraTime.Width - 150
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")) <> "" Then Call GetActiveView
    lblMoney.Top = (Me.Height - fraSplit.Top - fraInfo.Height - 1100) / 2
    fraInfo.Top = lblMoney.Top + lblMoney.Height + 30
    fraTotal.Top = fraInfo.Top + fraInfo.Height + 30
    fraPay.Top = fraTotal.Top + fraTotal.Height + 30
    cmdHelp.Top = fraPay.Top + fraPay.Height + 30
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Top = cmdHelp.Top
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
    Call SaveWinState(Me, App.ProductName)
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

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.����
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub txtArrangeNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If vsfArrange.Row - 1 >= vsfArrange.FixedRows Then
                KeyCode = 0
                vsfArrange.Row = vsfArrange.Row - 1
                vsfArrange_EnterCell
            End If
        Case vbKeyDown
            If vsfArrange.Row + 1 <= vsfArrange.Rows - 1 Then
                KeyCode = 0
                vsfArrange.Row = vsfArrange.Row + 1
                vsfArrange_EnterCell
            End If
        Case 13
            Call vsfArrange_KeyDown(13, 0)
    End Select
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
    If mRegistFeeMode = EM_RG_���� Then
        fraPay.Visible = False
    End If
    If mRegistFeeMode = EM_RG_���� Then
        If mty_Para.byt�Һ�ģʽ = 0 Then
            mRegistFeeMode = EM_RG_����
        Else
            mRegistFeeMode = EM_RG_����
        End If
    End If
    MCPAR.���ղ����� = gclsInsure.GetCapability(support�ҺŲ���ȡ������, lng����ID, mintInsure)
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure)
    mlng����ID = 0
    curMoney = GetRegistMoney
    Set rsTmp = GetMoneyInfoRegist(lng����ID, , , 1)
    If Not rsTmp Is Nothing Then cur��� = rsTmp!Ԥ����� - rsTmp!�������
    If cur��� > 0 Then
        lblMoney.Caption = "����Ԥ�����:" & Format(cur���, "0.00")
        If cur��� >= curMoney Then
            blnDeposit = True
        Else
            blnDeposit = False
        End If
    End If
    mcur������� = gclsInsure.SelfBalance(lng����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
    lblMoney.Caption = lblMoney.Caption & "  �����ʻ����:" & Format(mcur�������, "0.00")
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
        lblSum.Caption = "�� ��"
    End If
    If mRegistFeeMode = EM_RG_���� Then
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_����
        Else
            If mty_Para.byt�Һ�ģʽ = 0 Then
                mRegistFeeMode = EM_RG_����
                fraPay.Visible = True
            Else
                mRegistFeeMode = EM_RG_����
                fraPay.Visible = False
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
    Case "�Һŵ�"
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
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ָ���еĺű��Ƿ���Ч
    '���أ���Ч,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-17 16:00:11
    '˵����31922
    '------------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivs, ";��ʱ�Һ�;") > 0 Then
        CheckNoValied = True: Exit Function
    End If
    With vsfArrange
        If Val(.Cell(flexcpData, lngRow, .ColIndex("�ű�"))) = 1 Then
            MsgBox "�ű�" & .TextMatrix(lngRow, .ColIndex("�ű�")) & "��������Ч��Χ�ڻ���Ȩ�޲���,���ܹҺ�,����!", vbInformation + vbOKOnly + vbDefaultButton1
            Exit Function
        End If
    End With
    CheckNoValied = True
End Function

Private Sub SetGridTop(intRow As Integer)
    Dim intRows As Integer
    intRows = vsfArrange.Height \ vsfArrange.RowHeight(1) - 2
    If vsfArrange.TopRow + intRows > intRow Then Exit Sub
    vsfArrange.TopRow = intRow
End Sub

Private Sub txtArrangeNo_Change()
'���ܣ���������ű���ʾ����
    Dim strInfo As String, i As Integer
    Dim blnChkLimit As Boolean
    If mblnChangeByCode Then Exit Sub
    'ˢ�ºű�ֱ�Ӵӻ����ж�ȡ����
    If vsfArrange.Tag = "" Then
        Call LoadRegPlans(True)
    End If
    
    If (IsNumeric(Trim(txtArrangeNO.Text)) Or vsfArrange.Rows = 2) Or vsfArrange.Tag <> "" Then
        If vsfArrange.Tag = "" Then
            If vsfArrange.Rows <> 2 And Trim(txtArrangeNO.Text) <> vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")) Then
                '��ǰ�ű��б�ֻ��һ��ʱ�����û���������ű𣬲��Զ�ƥ�䣬���ǰ��س�
                Exit Sub
            End If
            '��λ����еĺű�
            For i = 1 To vsfArrange.Rows - 1
                If Trim(vsfArrange.TextMatrix(i, vsfArrange.ColIndex("�ű�"))) = Trim(txtArrangeNO.Text) Then
                    If CheckNoValied(i) = False Then
                         txtArrangeNO.Text = "": txtArrangeNO.SetFocus: Exit Sub
                    End If
                    vsfArrange.Row = i: vsfArrange.RowSel = i
                    vsfArrange.Col = 0: vsfArrange.ColSel = vsfArrange.Cols - 1
                    Call vsfArrange_EnterCell
                    SetGridTop i
                    Exit For
                End If
            Next
            '�ű����ް���ʱҪ������
            If i = vsfArrange.Rows And mrsPlan.RecordCount = 0 Then
                mblnChangeByCode = True
                txtArrangeNO.Text = ""
                mblnChangeByCode = False
                txtArrangeNO.SetFocus: Exit Sub
            End If
        End If
        

        blnChkLimit = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�޺�")) <> ""

        '�޺ſ���
        If blnChkLimit Then
            If zlCheck��Լ���޺���(txtArrangeNO.Text) = False Then Exit Sub
        End If
    End If
End Sub

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

Private Function zlCheck��Լ���޺���(ByVal str�ű� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Լ�����޺����Ƿ�Ϸ�
    '���:str�ű�-�ű�
    '����:
    '����:�Ϸ�,����ture,���򷵻�False
    '����:���˺�
    '����:2009-12-30 15:15:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, lngTemp As Long, strSQL As String, curDate As Date
    Dim lng��Լ�� As Long, lng�޺��� As Long, lng�ѹ��� As Long, lng��Լ�� As Long, lngʣ��ԤԼ�� As Long
    Dim lngʧԼ�� As Long
    Dim bln��ʱ�� As Boolean
    Dim strMsg As String
    Dim lng������λ���� As Long
    Dim blnHaveUnitreg As Boolean
    Dim i As Integer, j As Integer
    Err = 0: On Error GoTo Errhand:
    lng��Լ�� = 0: lng�޺��� = 0: lng�ѹ��� = 0: lng��Լ�� = 0: lngʣ��ԤԼ�� = 0

    curDate = CDate(Format(gobjDatabase.CurrentDate, "yyyy-MM-dd"))
    strSQL = _
      "Select Nvl(C.�޺���,0) as �޺���,Nvl(B.�ѹ���,0)  as �ѹ���,Nvl(C.��Լ��,0) as ��Լ��,Nvl(B.��Լ��,0) as ��Լ��,NVL(B.�����ѽ���,0) as �ѽ���" & _
      " From �ҺŰ��� A,���˹ҺŻ��� B,�ҺŰ������� C " & _
      " Where A.����ID=B.����ID(+) And A.��ĿID=B.��ĿID(+)  " & _
      "       And A.����=[1] And B.����(+)=[2] And A.����=B.����(+) " & _
      "       And Nvl(A.ҽ��ID,0)=Nvl(B.ҽ��ID(+),0) And Nvl(A.ҽ������,'ҽ��')=Nvl(B.ҽ������(+),'ҽ��') And  A.ID = C.����id(+)" & _
      "       And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����','7', '����', Null) = C.������Ŀ(+)"

   
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�, curDate, CDate(Format(curDate, "YYYY-MM-DD")))

    lngʧԼ�� = GetʧԼ��(str�ű�, curDate)

    If Not rsTmp.EOF Then
        lng��Լ�� = Val(Nvl(rsTmp!��Լ��)): lng�޺��� = Val(Nvl(rsTmp!�޺���))
        lng�ѹ��� = Val(Nvl(rsTmp!�ѹ���)): lng��Լ�� = Val(Nvl(rsTmp!��Լ��)) - Val(Nvl(rsTmp!�ѽ���))
        If lng��Լ�� < 0 Then lng��Լ�� = 0
        lngʣ��ԤԼ�� = IIf(lng�޺��� - lng�ѹ��� - lng��Լ�� <= 0, 0, lng��Լ�� - lng��Լ��): If lngʣ��ԤԼ�� < 0 Then lngʣ��ԤԼ�� = 0
        If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
        lng��Լ�� = lng��Լ�� - lngʧԼ��
    End If
    If lng�޺��� <= 0 Then
        '��������:����
        zlCheck��Լ���޺��� = True: Exit Function
    End If
    
    If lng�ѹ��� + lng��Լ�� >= lng�޺��� Then
        If InStr(gstrPrivs, ";�Ӻ�;") > 0 Then
            If MsgBox("�úű�����Ѵﵽ�޺��� " & lng�޺��� & "�����Ƿ�����Һ�?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                 mblnChangeByCode = True
                 txtArrangeNO = ""
                 mblnChangeByCode = False
                 If txtArrangeNO.Enabled And txtArrangeNO.Visible Then txtArrangeNO.SetFocus
                 Exit Function
            End If
            With vsfDetailTime
                For i = 0 To .Rows - 1
                    For j = 0 To .Cols - 1
                        If .Cell(flexcpData, i, j) Like "��*" Then .Select i, j
                    Next j
                Next i
            End With
        Else
            MsgBox "�úű�����Ѵﵽ�޺��� " & lng�޺��� & "�����ٹҺţ�", vbInformation, gstrSysName
            mblnChangeByCode = True
            txtArrangeNO = ""
            mblnChangeByCode = False
            If txtArrangeNO.Enabled And txtArrangeNO.Visible Then txtArrangeNO.SetFocus
            Exit Function
        End If
    End If
    
    zlCheck��Լ���޺��� = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function

Private Sub txtArrangeNo_GotFocus()
    Call gobjControl.TxtSelAll(txtArrangeNO)
End Sub

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
    Else
        If KeyCode = vbKeyEscape Then
            Call ReloadPage
        Else
            IDKind.ActiveFastKey
        End If
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
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub txtArrangeNo_KeyPress(KeyAscii As Integer)
    cboDeptFilter.ListIndex = 0
    cboDeptFilter.ListIndex = 0
    If KeyAscii = Asc(".") Then
        '����ڰ����˼�
        KeyAscii = 0: gobjCommFun.PressKey vbKeyBack
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If CheckNoValied(vsfArrange.Row) = False Then
             txtArrangeNO.Text = "": txtArrangeNO.SetFocus: Exit Sub
        End If
        
        vsfArrange.Tag = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�"))
        If vsfArrange.Tag <> "" Then
            If txtArrangeNO.Text <> vsfArrange.Tag Then
                txtArrangeNO.Text = vsfArrange.Tag  '�Զ�����change�¼�
            Else
                Call txtArrangeNo_Change
            End If
            vsfArrange.Tag = ""
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("1234567890+ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
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
    Dim rsTmp As ADODB.Recordset, strTemp As String, rsFeeType As ADODB.Recordset
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim strInputInfo As String '���洫��������ı� ������ʹ�����֤�� �Բ��˽��в��Һ� ���滻��"-" ����ID�����
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim blnҽ���� As Boolean
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
        lng����ID = GetPatiID(mlngModul, Me, strInput, txtPatient, , , blnCancel)
        If lng����ID = 0 And Not blnCancel Then
            If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        End If
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
                    
                strPati = strPati & " Union ALL " & _
                        "Select 0,0 as ID,-NULL,'[�²���]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
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
                lng����ID = GetPatiID(mlngModul, Me, strInput, txtPatient, , , blnCancel)
                If lng����ID = 0 And Not blnCancel Then
                    If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
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
        txtPatient.Text = Nvl(mrsInfo!����) '�����Change�¼�
        txtPatient.BackColor = &H80000005
        lblSum.Caption = "�� ��"
        If mblnAppointment Then
            fraPay.Visible = False
        Else
            If mty_Para.byt�Һ�ģʽ = 0 Then
                fraPay.Visible = True
            Else
                fraPay.Visible = False
            End If
        End If
        
        '�ڵ���txtPatient_Change�¼���������źͲ���������Ϊ�յ������ �޷�ʶ��ò�����Ϣ ���ִ���
        '���������ݿ����ݴ����ٽ��к����Ĵ���
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(Trim(mstr����) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        txtGender.Text = Nvl(mrsInfo!�Ա�)
        txtPatient.PasswordChar = ""
        
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        txtFeeType.Text = Nvl(mrsInfo!�ѱ�)
        If txtFeeType.Text = "" Then txtFeeType.Text = mstrDef�ѱ�
        txtAge.Text = Nvl(mrsInfo!����)
        txtClinic.Text = Nvl(mrsInfo!�����)
        If txtClinic.Text = "" Then
            txtClinic.Text = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        '����Ԥ������Ϣ
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!����ID, , , 1)
        If Not rsTmp Is Nothing Then cur��� = rsTmp!Ԥ����� - rsTmp!�������
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
        If txtArrangeNO.Enabled And txtArrangeNO.Visible Then txtArrangeNO.SetFocus
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
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    txtGender.Text = ""
    txtAge.Text = ""
    txtClinic.Text = ""
    txtFeeType.Text = ""
    lblMoney.Caption = ""
    chkBook.Enabled = True
    lblSum.Caption = "�ϼ�"
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_����
    Else
        If mty_Para.byt�Һ�ģʽ = 0 Then
            mRegistFeeMode = EM_RG_����
            fraPay.Visible = True
        Else
            mRegistFeeMode = EM_RG_����
            fraPay.Visible = False
        End If
    End If
    mintInsure = 0
    mlng����ID = 0
    Set mrsInfo = Nothing
    LoadPayMode False, False
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
            If (mintInsure <> 0 And MCPAR.���ղ�����) And cboPayMode.Text = "�����ʻ�" And cboPayMode.Visible Then
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

Private Sub LoadRegPlans(ByVal blnCache As Boolean)
    Dim strTime As String, strState As String, strWhere As String
    Dim strSQL As String, strIF As String
    Dim i As Integer, k As Integer
    Dim DateThis As Date, strZero As String
    Dim str�ҺŰ��� As String
    Dim str�ҺŰ��żƻ� As String
    Dim str����         As String
    Dim strFilter As String
    On Error GoTo errH
    
    str���� = "�ű�,����,��Ŀ,�ѹ�"
    
    If mblnAppointment Then
        DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
    Else
        DateThis = gobjDatabase.CurrentDate
    End If
    
    If Not blnCache Then
        strSQL = "Zl_�ҺŰ���_Autoupdate"
        gobjDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    If Not blnCache Then
    
        If gstrDeptIDs <> "" Then strIF = " And Instr(','||[4]||',',','||P.����ID||',')>0"
        
        '������ĺű���ˣ����ű���������вŹ���,��ʱ��ActiveControlһ����txtArrangeNo
        If Trim(txtArrangeNO.Text) <> "" And Trim(txtArrangeNO.Text) <> "+" And ActiveControl Is txtArrangeNO Then
            If IsNumeric(Trim(txtArrangeNO.Text)) Then
                strIF = strIF & " And P.���� Like [2]"
            Else
                strIF = strIF & " And (zlSpellCode(P.ҽ������) Like [2] or B.���� Like [2])"
            End If
        End If
        
        str�ҺŰ��� = "" & _
                "            Select A.ID, A.����, A.����, A.����id, A.��Ŀid, A.ҽ��id, A.ҽ������, A.��������, A. ����, A.��һ, A.�ܶ�, A.����, " & _
                "                   A.���� , A.����, A.����, A.���﷽ʽ,a.��ʼʱ��,a.��ֹʱ��, A.��ſ���, B.�޺���, B.��Լ��,a.ͣ������ " & vbNewLine & _
                "            From �ҺŰ��� A, �ҺŰ������� B " & vbNewLine & _
                "            Where a.ͣ������ Is Null And " & "[5] Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
                "                 Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                "                  And a.ID = B.����id(+) And Decode(To_Char([5], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) = B.������Ŀ(+)" '& vbNewLine & _

        str�ҺŰ��� = str�ҺŰ��� & " And Decode(To_Char([5], 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7',a.����, Null) Is Not Null"

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
        '�����ɸѡ
        If mrsPlan Is Nothing Then
            LoadRegPlans (False)
            Exit Sub
        End If
        If txtArrangeNO.Text <> "" Or cboDeptFilter.Text <> "" Or cboDoctorFilter.Text <> "" Then
            If txtArrangeNO.Text <> "" And mblnFilterChange = False Then
                strFilter = "�ű� like '" & txtArrangeNO.Text & "*'"
            End If
            If Trim(cboDeptFilter.Text) <> "" Then
                If strFilter <> "" Then
                    If InStr(cboDeptFilter.Text, "-") > 0 Then
                        strFilter = strFilter & " And ���� = '" & Split(cboDeptFilter.Text, "-")(1) & "'"
                    Else
                        strFilter = strFilter & " And ���� = '" & cboDeptFilter.Text & "'"
                    End If
                Else
                    If InStr(cboDeptFilter.Text, "-") > 0 Then
                        strFilter = "���� = '" & Split(cboDeptFilter.Text, "-")(1) & "'"
                    Else
                        strFilter = "���� = '" & cboDeptFilter.Text & "'"
                    End If
                End If
            Else
                If mblnFilterChange Then strFilter = ""
            End If
            If Trim(cboDoctorFilter.Text) <> "" Then
                If strFilter <> "" Then
                    strFilter = strFilter & " And ҽ�� = '" & cboDoctorFilter.Text & "'"
                Else
                    strFilter = "ҽ�� = '" & cboDoctorFilter.Text & "'"
                End If
            Else
                If mblnFilterChange And Trim(cboDeptFilter.Text) = "" Then strFilter = ""
            End If
            mrsPlan.Filter = strFilter
        Else
            LoadRegPlans (False)
            Exit Sub
        End If
        If mrsPlan.RecordCount <> 0 Then
            mrsPlan.MoveFirst
        Else
            vsfArrange.Clear 1
            vsfArrange.Rows = 2
            Exit Sub
        End If
    End If
    If mrsPlan.RecordCount = 0 And mblnAppointment Then
        vsfArrange.Clear 1
        If mblnInit Then MsgBox "��ǰû�п��õĹҺŰ��ţ����ڹҺŰ��Ź��������ú����ԣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsfArrange
        .Redraw = flexRDNone
        If Not mrsPlan.EOF Then
            mblnChangeByCode = True
            .ToolTipText = "�� " & mrsPlan.RecordCount & " ������"
            .Clear 1
            .Rows = 2
            .Rows = mrsPlan.RecordCount + 1
            mrsPlan.MoveFirst
            For i = 1 To mrsPlan.RecordCount
                .RowData(i) = Nvl(mrsPlan!����ID)
                .TextMatrix(i, .ColIndex("IDS")) = mrsPlan!ID & "," & mrsPlan!��ĿID & "," & IIf(IsNull(mrsPlan!ҽ��ID), 0, mrsPlan!ҽ��ID)
                .Cell(flexcpData, i, .ColIndex("IDS")) = mrsPlan!ID & "," & Val(Nvl(mrsPlan!�ƻ�ID))
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(mrsPlan!����), "", mrsPlan!����)
                .TextMatrix(i, .ColIndex("�ű�")) = mrsPlan!�ű�
                .TextMatrix(i, .ColIndex("����")) = mrsPlan!����
                .TextMatrix(i, .ColIndex("��Ŀ")) = mrsPlan!��Ŀ
                .Cell(flexcpData, i, .ColIndex("��Ŀ")) = Val(Nvl(mrsPlan!����))
                .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(mrsPlan!ҽ��)
                .TextMatrix(i, .ColIndex("��Լ")) = Nvl(mrsPlan!��Լ)
                .TextMatrix(i, .ColIndex("��Լ")) = Nvl(mrsPlan!��Լ)
                .TextMatrix(i, .ColIndex("�ѹ�")) = Nvl(mrsPlan!�ѹ�)
                .TextMatrix(i, .ColIndex("�޺�")) = Nvl(mrsPlan!�޺�)
                .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!��), 1)
                .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!��)
                .TextMatrix(i, .ColIndex("һ")) = Left(Nvl(mrsPlan!һ), 1)
                .Cell(flexcpData, i, .ColIndex("һ")) = Nvl(mrsPlan!һ)
                .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!��), 1)
                .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!��)
                .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!��), 1)
                .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!��)
                .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!��), 1)
                .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!��)
                .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!��), 1)
                .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!��)
                .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!��), 1)
                .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!��)
                .TextMatrix(i, .ColIndex("����")) = Nvl(mrsPlan!����)
                .TextMatrix(i, .ColIndex("��ſ���")) = IIf(mrsPlan!��ſ��� = 1, "��", "")
                .Cell(flexcpData, i, .ColIndex("�ű�")) = ""
                
                mrsPlan.MoveNext
            Next
            mblnChangeByCode = False
        Else
            Set mrsPlan = Nothing
            .Clear 1
            .Rows = 2
            .ToolTipText = ""
        End If

        Call SetvsfarrangeFiexBackColor
        If blnCache = True Then mblnChangeByCode = True
        Call vsfArrange_EnterCell
        mblnChangeByCode = False
        If txtArrangeNO.Visible And txtArrangeNO.Enabled And Not mblnFilterChange Then txtArrangeNO.SetFocus
'        If mrsPlan.RecordCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub SetvsfarrangeFiexBackColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ع̶��еı���ɫ
    '����:blnCurDate-�Ƿ�ǰ������,�������ԤԼ������
    '����:���˺�
    '����:2010-02-04 14:39:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim PreRedaw As RedrawSettings, i As Long, strSQL As String, strNow As String
    Dim strKey As String, rsTmp As ADODB.Recordset, strColor As String
    With vsfArrange
         .Redraw = flexRDNone
        strSQL = "Select ʱ���,��ʼʱ��,��ǰʱ��,��ǰ��ɫ From ʱ���"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
        strNow = Format(gobjDatabase.CurrentDate, "HH:MM:SS")
        strKey = zlGet��ǰ���ڼ�
        For i = 1 To .Rows - 1
            rsTmp.Filter = "ʱ���='" & .Cell(flexcpData, i, .ColIndex(strKey)) & "'"
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!��ǰʱ��) Then
                    strColor = Nvl(rsTmp!��ǰ��ɫ, "0")
                    If strNow < Format(Nvl(rsTmp!��ʼʱ��), "HH:MM:SS") Then
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = strColor
                    End If
                End If
            End If
        Next i
        .Redraw = flexRDBuffered
    End With
End Sub

Private Function GetActiveView()
    '�õ���ǰ�Һ�ҵ��  ��ȡ�������͵�����
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    Dim str����         As String
    Dim dat            As Date
    
    On Error GoTo errH
    str���� = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�"))
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
    On Error Resume Next
    If rsTmp.RecordCount > 0 And vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("��ſ���")) <> "" Then
       '*********************
       'ר�Һŷ�ʱ��
       '*********************
       mViewMode = v_ר�Һŷ�ʱ��
       vsfArrange.Height = fraTime.Height / 2 - 300
       vsfDetailTime.Top = vsfArrange.Top + vsfArrange.Height + 60
       vsfDetailTime.Height = fraTime.Height - vsfDetailTime.Top - 90
'       vsfArrange.Height = vsfDetailTime.Top - 45 - vsfArrange.Top
       vsfDetailTime.Visible = True
    ElseIf rsTmp.RecordCount > 0 And vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("��ſ���")) = "" Then
       '*********************
       '��ͨ�ŷ�ʱ��
       '*********************
       mViewMode = V_��ͨ�ŷ�ʱ��
       vsfArrange.Height = fraTime.Height - 660
'       vsfArrange.Height = vsfDetailTime.Top + vsfDetailTime.Height - vsfArrange.Top
       vsfDetailTime.Visible = False
    ElseIf vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("��ſ���")) <> "" And vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�޺�")) <> "" Then
       '*********************
       'ר�ҺŲ���ʱ��
       '*********************
       mViewMode = v_ר�Һ�
       vsfArrange.Height = fraTime.Height / 2 - 300
       vsfDetailTime.Top = vsfArrange.Top + vsfArrange.Height + 60
       vsfDetailTime.Height = fraTime.Height - vsfDetailTime.Top - 90
'       vsfArrange.Height = vsfDetailTime.Top - 45 - vsfArrange.Top
       vsfDetailTime.Visible = True
     Else
       '*********************
       '��ͨ��
       '*********************
       mViewMode = V_��ͨ��
       vsfArrange.Height = fraTime.Height - 660
'       vsfArrange.Height = vsfDetailTime.Top + vsfDetailTime.Height - vsfArrange.Top
       vsfDetailTime.Visible = False
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
    
    strSQL = "Select Distinct a.��� As ID, a.���, To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��, To_Char(A.��ʼʱ��, 'hh24') || ':00' As ʱ���" & vbNewLine & _
            "From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
            "Where a.����id = b.Id And b.���� = [1] And" & vbNewLine & _
            "      Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����'," & vbNewLine & _
            "             Null) = a.����(+) And Not Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From �ҺŰ��żƻ� E" & vbNewLine & _
            "       Where e.����id = b.Id And e.���ʱ�� Is Not Null And" & vbNewLine & _
            "             [2] Between Nvl(e.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "             Nvl(e.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')))"
    strSQL = strSQL & " Union " & _
            "Select Distinct a.��� As ID,  a.���,To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��, To_Char(A.��ʼʱ��, 'hh24') || ':00' As ʱ���" & vbNewLine & _
            "From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C," & vbNewLine & _
            "     (Select Max(a.��Чʱ��) ��Ч" & vbNewLine & _
            "       From �ҺŰ��żƻ� A, �ҺŰ��� B" & vbNewLine & _
            "       Where a.����id = b.Id And b.���� = [1] And a.���ʱ�� Is Not Null And" & vbNewLine & _
            "             [2] Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "             Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'))) D" & vbNewLine & _
            "Where a.�ƻ�id = b.Id And b.����id = c.Id And c.���� = [1] And b.��Чʱ�� = d.��Ч And b.���ʱ�� Is Not Null And" & vbNewLine & _
            "      [2] Between Nvl(b.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "      Nvl(b.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And " & vbNewLine & _
            "      Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5'," & vbNewLine & _
            "                                           '����', '6', '����', '7', '����', Null) = a.����(+)" & vbNewLine & _
            "Order By ��ʼʱ��"
    If strSQL = "" Then Exit Function
    
    strNO = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�"))
    '��ȡ���� �������Ҫ����
    If mblnAppointment Then
        dateCur = Format(dtpDate, "yyyy-mm-dd")
    Else
        dateCur = gobjDatabase.CurrentDate
    End If
    
    On Error GoTo errH
    Set mrsʱ��� = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, dateCur)
    If mrsʱ���.EOF Then Exit Function
    InitTimePlan = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadTimePlan()
    '***************************************
    '����ʱ���
    '***************************************
    Dim i               As Integer
    Dim j               As Integer
    Dim blnPre          As Boolean
    Dim lngThis         As Long
    Dim lngMax          As Long
    Dim datThis         As Date
    Dim lngCurrSn       As Long
    Dim lngMaxSn        As Long 'ԤԼ�����ʹ�ú�
    Dim strSQL          As String
    Dim rsʱ��ͳ��      As ADODB.Recordset
    Dim strʱ���       As String
    Dim lngԤԼ����     As Long
    Dim lngTatol        As Long '���ڷ�ʱ�� ������¼�������
    Dim strMaxDate      As String  '���ڷ�ʱ�α����ԤԼʱ��
    Dim lngCols         As Long
    Dim lngRows         As Long
    Dim strData         As String
    Dim strDate         As String
    Dim blnHave         As Boolean
    Dim datMax          As Date
    Dim Datsys          As Date
    Dim blnʧԼ���ڹҺ� As Boolean
    Dim blnInserted     As Boolean
    Dim lng������λ���� As Long
    Dim blnFindSN      As Boolean '�Ƿ���Ҫ���¶�λ���ϴκű�����,����ˢ���б�ʱ,���ݱ���
    Dim lngFindSN      As Long '��Ҫ���ҵ����
     
    vsfDetailTime.Redraw = False
    vsfDetailTime.Clear
    '***************************************
    '�����Ϣ����
    '***************************************
    lngMax = Val(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�޺�")))
    
    '1.����λ��
    If lngMax > 1000 Then
        vsfDetailTime.FontWidth = 4
    Else
        vsfDetailTime.FontWidth = 0 '�ָ�ȱʡ����
    End If
    '***************************************
    '��ʼ��ʱ���
    '***************************************
     If InitTimePlan() = False Then vsfDetailTime.Redraw = True: Exit Sub
     Datsys = gobjDatabase.CurrentDate
    '***************************************
    '��ʼ�����
    '***************************************
     
     If mrsʱ��� Is Nothing Then vsfDetailTime.Redraw = True: Exit Sub
     'If mrsʱ���.RecordCount = 0 Then Exit Sub
 
    '***************************************
    '������
    '***************************************
    With vsfDetailTime
       .Rows = 1
       .Cols = 1
       .Clear
    End With
    lngCurrSn = -1
    Select Case mViewMode
    Case V_��ͨ�ŷ�ʱ��:
        Exit Sub
        Set rsʱ��ͳ�� = Nothing
    Case v_ר�Һŷ�ʱ��:
     '*******************************
     'ר�Һŷ�ʱ��
     'ÿ����ʱ�������
     '*******************************
regHD:
        blnInserted = False
        strʱ��� = ""
        With mrsʱ���
            lngRows = -1: lngCols = 0
            datMax = CDate("00:00:00")
            Do While Not .EOF
                If datMax < CDate(Nvl(!��ʼʱ��, "00:00:00")) Then datMax = CDate(!��ʼʱ��)
                'ԤԼ״̬ ֻ�������ԤԼ��ʱ���
                '�Һ�ʱ�����ֶ����
                If blnFindSN Then
                    If Val(Nvl(!���)) = lngFindSN And lngFindSN > 0 Then
                          lngCurrSn = lngFindSN
                    End If
                End If


                If strʱ��� <> Nvl(!ʱ���) Then
                    lngRows = lngRows + 1
                    strʱ��� = Nvl(!ʱ���)
                    If lngRows > vsfDetailTime.Rows - 1 Then vsfDetailTime.Rows = vsfDetailTime.Rows + 1: lngCols = 0
                    If lngCols > vsfDetailTime.Cols - 1 Then vsfDetailTime.Cols = vsfDetailTime.Cols + 1
                    vsfDetailTime.TextMatrix(lngRows, 0) = strʱ���
                    vsfDetailTime.Cell(flexcpForeColor, lngRows, 0, lngRows, 0) = vsfArrange.Cell(flexcpForeColor, vsfArrange.Row, 0, vsfArrange.Row, 0)
                 End If
                lngCols = lngCols + 1
                  If lngCols > vsfDetailTime.Cols - 1 Then vsfDetailTime.Cols = vsfDetailTime.Cols + 1
                strData = !��� & vbCrLf & !��ʼʱ�� & "-" & !����ʱ��
                vsfDetailTime.TextMatrix(lngRows, lngCols) = strData
                If (Format(Datsys, "hh:mm:ss") > Format(!��ʼʱ��, "hh:mm:ss")) And (Format(Datsys, "yyyy-mm-dd") = Format(dtpDate.Value, "yyyy-mm-dd")) Then
                    vsfDetailTime.Cell(flexcpFontUnderline, lngRows, lngCols) = True
                    vsfDetailTime.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                End If
            .MoveNext
          Loop
          If blnHave = False And vsfDetailTime.Rows = 1 And vsfDetailTime.Cols = 1 And mrsʱ���.RecordCount > 0 Then blnHave = True: mrsʱ���.MoveFirst: GoTo regHD
          
'            .MoveLast
'            For i = 1 To vsfDetailTime.Cols - 1
'                If vsfDetailTime.TextMatrix(vsfDetailTime.Rows - 1, i) = "" Then
'                    If blnInserted = False Then
'                        vsfDetailTime.TextMatrix(vsfDetailTime.Rows - 1, i) = " " & vbCrLf & !����ʱ�� & "�Ժ�"
'                        vsfDetailTime.Cell(flexcpData, vsfDetailTime.Rows - 1, i) = "�Ӻ�"
'                        blnInserted = True
'                    End If
'                End If
'            Next i
'            If blnInserted = False Then
'                vsfDetailTime.Cols = vsfDetailTime.Cols + 1
'                vsfDetailTime.TextMatrix(vsfDetailTime.Rows - 1, vsfDetailTime.Cols - 1) = " " & vbCrLf & !����ʱ�� & "�Ժ�"
'                vsfDetailTime.Cell(flexcpData, vsfDetailTime.Rows - 1, vsfDetailTime.Cols - 1) = "�Ӻ�"
'            End If
        End With
    End Select
    dtpTime.Tag = Format(datMax, "hh:mm:ss")
    '***************************************
    '��ű��״̬����
    '***************************************
    Call SetSnStyle(True)
    '***************************************
    '���״̬ ���
    '���ڹҺ�״̬��Ҫ����ֻ��һ��״̬
    '***************************************
     If mViewMode = v_ר�Һŷ�ʱ�� Then
        datThis = dtpDate
        Set mrsSNState = GetSNState(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")), datThis)

        If mrsSNState.RecordCount > 0 Then
                For i = 0 To vsfDetailTime.Rows - 1
                   For j = 1 To vsfDetailTime.Cols - 1
                       If vsfDetailTime.TextMatrix(i, j) <> "" And Not vsfDetailTime.Cell(flexcpData, i, j) Like "��*" Then
                        '**********************************************
                        '
                        '**********************************************
                        On Error Resume Next
                        vsfDetailTime.Row = i: vsfDetailTime.Col = j
                        On Error GoTo Hd
                        lngFindSN = Val(Getʱ��(i, j, False))
                        mrsSNState.Filter = "���=" & lngFindSN
                        If mrsSNState.RecordCount > 0 Then
                            If lngCurrSn = lngFindSN Then lngCurrSn = -1
                                Select Case mrsSNState!״̬
                                Case 1  '�ѹ�
                                      If Nvl(mrsSNState!ԤԼ, "0") = "0" Then
                                        vsfDetailTime.Cell(flexcpForeColor, i, j) = vbRed
                                      Else
                                        vsfDetailTime.Cell(flexcpForeColor, i, j) = &HC000C0
                                      End If
                                      vsfDetailTime.Cell(flexcpFontStrikethru, i, j) = True
                                Case 2  '��Լ
                                    vsfDetailTime.Cell(flexcpForeColor, i, j) = vbGreen
                                If lngMaxSn < Val(Nvl(mrsSNState!���)) Then
                                    lngMaxSn = Val(Nvl(mrsSNState!���))
                                End If
                                Case 3  '����
                                  vsfDetailTime.Cell(flexcpForeColor, i, j) = vbBlue
                                Case 4  '�˺�
                                    vsfDetailTime.Cell(flexcpForeColor, i, j) = vbGrayText
                                    vsfDetailTime.Cell(flexcpFontStrikethru, i, j) = True
                                Case 5  '����
                                    vsfDetailTime.Cell(flexcpForeColor, i, j) = vbRed
                                End Select
                            End If
                        End If
                   Next
                Next
        End If
     End If
     '���п�����ŵ�����£����μӺ���
    If CheckAddAvailable = False Then
        For i = 0 To vsfDetailTime.Rows - 1
            For j = 1 To vsfDetailTime.Cols - 1
                If vsfDetailTime.Cell(flexcpData, i, j) Like "��*" Then
                    vsfDetailTime.Cell(flexcpData, i, j) = ""
                    vsfDetailTime.TextMatrix(i, j) = ""
                End If
            Next j
        Next i
    End If
    If vsfDetailTime.Rows > 1 Then
       vsfDetailTime.Cell(flexcpFontBold, 0, 0, vsfDetailTime.Rows - 1, 0) = True
    End If
    
    dtpTime.Value = Format(Me.dtpTime.Tag, "hh:mm")
    vsfDetailTime.Redraw = True
    locateSnByʱ�� lngCurrSn
    mblnChangeByCode = True
    txtArrangeNO.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�"))
    mblnChangeByCode = False
    txtDept.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("����"))
    cboDoctor.Clear
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("ҽ��")) = "" Then
        cboDoctor.Locked = False
        cboDoctor.Enabled = True
        Call LoadDoctor(vsfArrange.RowData(vsfArrange.Row))
    Else
        cboDoctor.AddItem vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("ҽ��"))
        cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(2))
        cboDoctor.ListIndex = cboDoctor.NewIndex
        cboDoctor.Locked = True
        cboDoctor.Enabled = False
    End If
    Call LoadFeeItem(Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(1)), chkBook.Value = 1)
    Call vsfDetailTime_DblClick
    Exit Sub
Hd:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub locateSnByʱ��(Optional ByVal lngSN As Long = -1, _
    Optional blnǿ�ƶ�λ As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��λ��ָ����ʱ��
    '���:lngSN:>0��Ҫ��λ�������,-1:��ʾ������ȡ��
    '����:blnǿ�ƶ�λ-ǿ�ƶ�λ��ָ������������
    '����:���˺�
    '����:2013-12-07 13:01:55
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngRow As Long, lngCol As Long
    Dim blnFind  As Boolean, blnExit As Boolean, blnMaxSn As Boolean
    Dim lngLastRow As Long, lngLastCol As Long
    lngRow = 0: lngCol = 1
    On Error GoTo errH
'    vsfDetailTime.HighLight = flexHighlightAlways
    Select Case mViewMode
    Case V_��ͨ�ŷ�ʱ��:
         Exit Sub
    Case v_ר�Һŷ�ʱ��:
        blnMaxSn = True
        With vsfDetailTime
            For i = 0 To .Rows - 1
                For j = 1 To .Cols - 1
                    If .TextMatrix(i, j) <> "" Then
                       
                         If .Cell(flexcpForeColor, i, j) <> vbRed _
                             And .Cell(flexcpForeColor, i, j) <> vbBlue _
                             And .Cell(flexcpForeColor, i, j) <> vbGrayText Then
                             
                            If blnMaxSn = True _
                                And .Cell(flexcpForeColor, i, j) <> vbGreen _
                                And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then
                                If Not mty_Para.bln������ѡ�� Or lngSN = -1 Then  '66788
                                    blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                    blnExit = True: Exit For
                                End If
                             End If
                             
                             If lngSN <> -1 Then
                                 If lngSN = Val(Getʱ��(i, j, False)) Then
                                    .Row = i: .Col = j
                                     blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                     dtpTime.Value = CDate(Getʱ��(i, j, True))
                                     blnExit = True: Exit For
                                 End If
                             End If
                         Else
                              blnMaxSn = True
                         End If
                    End If
                Next
                If blnExit Then Exit For '45768
            Next
        End With
        
        If blnFind And blnMaxSn = False Then
            mblnChangeByCode = True
            vsfDetailTime.Row = lngRow: vsfDetailTime.Col = lngCol
            mblnChangeByCode = False
'            vsfDetailTime.HighLight = flexHighlightAlways
        Else
            vsfDetailTime.Select 0, 0
            vsfDetailTime.HighLight = flexHighlightNever
        End If
        dtpTime.Value = IIf(blnFind = False And blnMaxSn, Format(CDate(gobjDatabase.CurrentDate), "hh:mm:ss"), Format(CDate(Getʱ��(lngRow, lngCol, True)), "hh:mm:ss"))
    Case Else: Exit Sub
    End Select
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub vsfArrange_DblClick()
    mblnChangeByCode = True
    If txtPatient.Text = "" Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Else
        Call vsfArrange_KeyDown(13, 0)
    End If
    mblnChangeByCode = False
End Sub

Private Sub vsfArrange_EnterCell()
    Dim i           As Integer
    Dim j           As Integer
    Dim blnPre      As Boolean
    Dim lngThis     As Long
    Dim lngMax      As Long
    Dim datThis     As Date
    Dim lngCurrSn   As Long
    Dim lngMaxSn    As Long 'ԤԼ�����ʹ�ú�
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim blnChk      As Boolean
    Dim varTemp     As Variant
    Dim sngTime     As Single
      
    '*****************************
    '��ȡʹ���������̴���Һ�
    '******************************
    If mblnChangeByCode Then Exit Sub
    sngTime = Timer
    If Format(sngTime, "0.000") - Format(msngTime, "0.000") < 0.1 Then
        mblnChangeByCode = True
        If mlngRow <> 0 Then vsfArrange.Select mlngRow, 0
        mblnChangeByCode = False
        Exit Sub
    End If
    If Val(vsfArrange.Cell(flexcpData, vsfArrange.Row, vsfArrange.ColIndex("��Ŀ"))) = 1 Then
        lbl��.Visible = True
    Else
        lbl��.Visible = False
    End If
    msngTime = Timer
    mlngRow = vsfArrange.Row
    GetActiveView
    If mViewMode = v_ר�Һŷ�ʱ�� Then
       '*************************************************
       '������ڷ�ʱ�ε���� ʹ�÷�ʱ�εĴ�����
       '*************************************************
       LoadTimePlan
       Call vsfDetailTime_AfterRowColChange(vsfDetailTime.Row, vsfDetailTime.Col, vsfDetailTime.Row, vsfDetailTime.Col)
       Call ReadRoom
       Exit Sub
    End If
    
    vsfDetailTime.Redraw = False
    vsfDetailTime.Clear

    lngMax = Val(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�޺�"))) '�ҽ����ĺŲ�����ԤԼ,��Ϊ�ѽ���,Ӧ���ɹҺ�

    If lngMax > 0 And vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("��ſ���")) <> "" Then
        If lngMax = 0 Then GoTo regTab
        '1.����λ��
        If lngMax > 1000 Then
            vsfDetailTime.FontWidth = 4
        Else
            vsfDetailTime.FontWidth = 0 '�ָ�ȱʡ����
        End If

        If (lngMax \ SNCOLS) * SNCOLS = lngMax Then
            vsfDetailTime.Rows = lngMax \ SNCOLS
        Else
            vsfDetailTime.Rows = lngMax \ SNCOLS + 1
        End If
        'mblnNotClick = False
        vsfDetailTime.Cols = SNCOLS
        If Not vsfDetailTime.Visible Then
            vsfDetailTime.Visible = True
'            picSplit.Visible = True
        End If
                                
        '������
        lngThis = 1
        For i = 0 To vsfDetailTime.Rows - 1
            For j = 0 To vsfDetailTime.Cols - 1
                vsfDetailTime.TextMatrix(i, j) = lngThis
                lngThis = lngThis + 1
                If lngThis > lngMax Then Exit For
            Next
            If lngThis > lngMax Then Exit For
        Next
        
        Set mrsSNState = GetSNState(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")), datThis)
        lngMaxSn = 0
       For i = 0 To mrsSNState.RecordCount - 1
            If mrsSNState!��� <= lngMax Then
                If (mrsSNState!��� \ SNCOLS) * SNCOLS = mrsSNState!��� Then
                   lngRow = (mrsSNState!��� \ SNCOLS) - 1
                   lngRow = IIf(lngRow < 0, 0, lngRow)
                Else
                    lngRow = (mrsSNState!��� \ SNCOLS)
                End If
                    lngCol = (mrsSNState!��� - 1) Mod SNCOLS
                    lngCol = IIf(lngCol < 0, 0, lngCol)
                Select Case mrsSNState!״̬
                    Case 1  '�ѹ�
                       If Nvl(mrsSNState!ԤԼ, "0") = "0" Then
                          vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                          '������Ŷ�λ������Ч�ź�
                          If lngMaxSn < Val(Nvl(mrsSNState!���)) Then
                            lngMaxSn = Val(Nvl(mrsSNState!���))
                          End If
                       Else
                          'ԤԼ����
                          vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = &HC000C0
                       End If
                    Case 2  '��Լ
                        vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbGreen
                        
                       
                    Case 3  '����
                        vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbBlue
                    Case 4  '�˺�
                        vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
                        vsfDetailTime.Cell(flexcpFontStrikethru, lngRow, lngCol) = True
                    Case 5  '����
                        vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                End Select
            End If
            mrsSNState.MoveNext
        Next
        lngCurrSn = GetCurrSN(lngMaxSn)
    Else
regTab:
        Set mrsSNState = Nothing
        vsfDetailTime.Visible = False
    End If
    vsfDetailTime.Redraw = True
    SetSnStyle
    vsfDetailTime.Select 0, 0
    Call LocateSN(lngCurrSn)
    If vsfDetailTime.Row <= vsfDetailTime.Rows - 1 And vsfDetailTime.Col <= vsfDetailTime.Cols - 1 And vsfDetailTime.Cell(flexcpForeColor, vsfDetailTime.Row, vsfDetailTime.Col) = vbBlack Then
        vsfDetailTime.Cell(flexcpBackColor, vsfDetailTime.Row, vsfDetailTime.Col) = &H8000000D
    End If
    txtDept.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("����"))
    Call ReadRoom
    cboDoctor.Clear
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("ҽ��")) = "" Then
        cboDoctor.Locked = False
        cboDoctor.Enabled = True
        Call LoadDoctor(vsfArrange.RowData(vsfArrange.Row))
    Else
        cboDoctor.AddItem vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("ҽ��"))
        cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(2))
        cboDoctor.ListIndex = cboDoctor.NewIndex
        cboDoctor.Locked = True
        cboDoctor.Enabled = False
    End If
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")) = "" Then Exit Sub
    varTemp = Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")
    Call LoadFeeItem(Val(varTemp(1)), chkBook.Value = 1)
    Call vsfDetailTime_DblClick
End Sub

Private Sub ReadRoom()
    Dim blnBusy As Boolean, strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    blnBusy = Val(gobjDatabase.GetPara("����æʱ�������", glngSys, 1113, 0)) = 1
    strSQL = _
        " Select b.����, b.����, b.λ��" & vbNewLine & _
        " From �ҺŰ������� a, �������� b, �ҺŰ��� c" & vbNewLine & _
        " Where a.�������� = b.���� And a.�ű�id = c.Id And c.���� = [1] " & _
        IIf(blnBusy, " ", " And b.ȱʡ��־=0 ")
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")))
    cboRoom.Clear
    Do While Not rsTmp.EOF
        cboRoom.AddItem Nvl(rsTmp!����)
        cboRoom.ItemData(cboRoom.NewIndex) = Nvl(rsTmp!����)
        rsTmp.MoveNext
    Loop
    If cboRoom.ListCount = 1 Then cboRoom.ListIndex = 0
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub LoadFeeItem(ByVal lngItemID As Long, ByVal blnBook As Boolean)
    Dim strSQL As String, str���ʽ As String
    Dim i As Integer, j As Integer, dblTotal As Double
    Dim curӦ�� As Currency, curʵ�� As Currency
    
    On Error GoTo errH
    ReadRegistPrice lngItemID, blnBook, False, txtFeeType.Text, mrsItems, mrsInComes
    
    '126802�����ϴ���2018/6/7��ԤԼ�������ӷ�
    If Not mrsInfo Is Nothing And (mblnAppointment = False Or mty_Para.blnԤԼʱ�տ�) Then
        If mrsInfo.State = 1 Then
            If Not mrsInfo.EOF Then
                str���ʽ = Nvl(mrsInfo!ҽ�Ƹ��ʽ)
                If str���ʽ = "" Then str���ʽ = mstrDef���ʽ
                
                Call ReadExRegistPrice(mrsExpenses, mblnAppointPrice, Val(Nvl(mrsInfo!����ID)), mintInsure, _
                        vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�")), _
                        Nvl(mrsInfo!����), NeedName(txtGender.Text), txtAge.Text, Nvl(mrsInfo!���֤��), txtFeeType.Text, str���ʽ)
            End If
        End If
    End If
    
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = "0.00"
    txtPayMoney.Text = "0.00"
    dblTotal = 0
    With vsfMoney
        If mrsItems.RecordCount = 0 Then Exit Sub
        mrsItems.MoveFirst
        Do While Not mrsItems.EOF
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
        Loop
        
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
    lblTotal.Caption = Format(dblTotal, "0.00")
    txtPayMoney.Text = Format(dblTotal, "0.00")
    Call GetYBInfo
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub LoadDoctor(ByVal lng����id As Long)
'���ܣ����ݿ��Ҷ�ȡ����ҽ�������б�
    Dim strSQL As String
        
    On Error GoTo errH
    If mrsDoctor Is Nothing Then Call GetAllҽ��
    If mrsDoctor.State = 1 Then
        mrsDoctor.Filter = "����id=" & lng����id
        
        Do While Not mrsDoctor.EOF
            cboDoctor.AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
            cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
            mrsDoctor.MoveNext
        Loop
        cboDoctor.ListIndex = -1
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Call gobjComlib.SaveErrLog
End Sub


Private Sub LocateSN(lngCurrSn As Long)
'����:��λ��ָ�������
'     �����������ű�����,����ű��ý���
    Dim lngRow          As Long
    Dim i               As Long
    Dim j               As Long
    Dim blnHave         As Boolean
    If lngCurrSn = 0 Then Exit Sub
   
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then
        '************************************************
        '����ʱ�� ��Ŷ�λ���ǰ�����ǰ�ķ�ʽ
        '************************************************
        If (lngCurrSn \ SNCOLS) * SNCOLS = lngCurrSn Then
            lngRow = (lngCurrSn - 1) \ SNCOLS
        Else
            lngRow = (lngCurrSn \ SNCOLS)
        End If
        If Not vsfDetailTime.RowIsVisible(lngRow) Then
            If lngRow >= 1 Then  '������һ�пɼ�
                vsfDetailTime.TopRow = lngRow - 1
            Else
                vsfDetailTime.TopRow = lngRow
            End If
        End If
        mblnChangeByCode = True
        vsfDetailTime.Select lngRow, (lngCurrSn - 1) Mod SNCOLS
        mblnChangeByCode = False
'        vsfDetailTime.Row = lngRow
'        vsfDetailTime.RowSel = vsfDetailTime.Row
'        vsfDetailTime.Col = (lngCurrSn - 1) Mod SNCOLS
'        vsfDetailTime.ColSel = vsfDetailTime.Col
     
    ElseIf mViewMode = v_ר�Һŷ�ʱ�� Then
        '*******************************************
        'ר�Һŷ�ʱ�� ��Ŷ�λ
        '*******************************************
        For i = 0 To vsfDetailTime.Rows - 1
            For j = 1 To vsfDetailTime.Cols - 1
               If vsfDetailTime.TextMatrix(i, j) <> "" Then
                    If lngCurrSn = Val(Getʱ��(i, j, False)) Then
                     If Not vsfDetailTime.RowIsVisible(i) Then
                        If lngRow >= 1 Then  '������һ�пɼ�
                             vsfDetailTime.TopRow = i - 1
                        Else
                             vsfDetailTime.TopRow = i
                        End If
                      End If
                      vsfDetailTime.Row = i
                      vsfDetailTime.Col = j
                     blnHave = True
                     Exit For
                    End If
                End If
            Next
            If blnHave Then Exit For
        Next
    End If
'    vsfDetailTime.HighLight = flexHighlightAlways
    If vsfDetailTime.Visible And vsfDetailTime.Enabled _
                And Not Me.ActiveControl Is txtArrangeNO _
                And Not Me.ActiveControl Is vsfArrange Then Call vsfDetailTime.SetFocus     '�����ںű�������������
End Sub

Private Function Getʱ��(ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal blnTime As Boolean = False, Optional ByVal blnLastTime As Boolean = False) As String
    '*****************************************************************
    '����˵��:�ڹҺ�ר�Һŷ�ʱʱ ��ȡ ���,���� ��ʼʱ��
    '����:  blntime �Ƿ��ȡʱ�� �����ȡʱ��  ���򷵻����
    '*****************************************************************
    Dim strResult       As String, i As Long
    On Error GoTo errH
    If lngRow > vsfDetailTime.Rows - 1 Or lngCol > vsfDetailTime.Cols - 1 Then
        Exit Function
    End If
    If vsfDetailTime.TextMatrix(lngRow, lngCol) = "" Then
        Exit Function
    End If
    
    If blnTime Then
        i = IIf(blnLastTime = False, 0, 1)
        If InStr(vsfDetailTime.TextMatrix(lngRow, lngCol), "-") > 0 Then
            Getʱ�� = Split(Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(1), "-")(i)
        Else
            If InStr(vsfDetailTime.TextMatrix(lngRow, lngCol), "��") = 0 Then Exit Function
            Getʱ�� = Split(Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(1), "��")(i)
        End If
        Exit Function
    End If
    If mViewMode = v_ר�Һŷ�ʱ�� Then
       strResult = Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(0)
    ElseIf mViewMode = V_��ͨ�ŷ�ʱ�� Then
       strResult = Replace(Replace(Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(0), "ԤԼ", ""), "����", "")
    End If
    Getʱ�� = strResult
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub SetSnStyle(Optional ByVal bln��ʱ�� As Boolean = False)
'****************************************
'�Ա����ʽ��������
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
    On Error GoTo errH
    Select Case bln��ʱ��
    Case False:
        With vsfDetailTime
            
            .FixedCols = 0
            lngWidth = 570
            lngHeight = 450
            For i = 0 To vsfDetailTime.Cols - 1
                .ColWidth(i) = lngWidth
                .ColAlignment(i) = 4
            Next
            For i = 0 To vsfDetailTime.Rows - 1
                 .RowHeight(i) = lngHeight
            Next
            
        End With
    
    Case True:
        With vsfDetailTime
             If .Cols <= 1 Then Exit Sub
             .FixedCols = 1
             .FixedAlignment(0) = flexAlignRightTop
             .ColAlignment(0) = flexAlignRightTop
            lngWidth = 1275
            lngHeight = 550
            For i = 1 To vsfDetailTime.Cols - 1
                .ColWidth(i) = lngWidth
                .ColAlignment(i) = 4
            Next
            .ColAlignment(0) = 3
            .ColWidth(0) = lngWidth
            For i = 0 To vsfDetailTime.Rows - 1
                 .RowHeight(i) = lngHeight
            Next
           If .Rows > 0 And .Cols > 0 Then
                .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = True
                .Cell(flexcpFontSize, 0, 1, .Rows - 1, .Cols - 1) = 9
                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 18
           End If
        End With
    End Select
    If vsfDetailTime.Rows >= 1 And vsfDetailTime.Cols > 0 Then
       vsfDetailTime.Cell(flexcpFontBold, 0, 0, vsfDetailTime.Rows - 1, vsfDetailTime.Cols - 1) = True
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Function GetCurrSN(Optional ByVal lngCurMaxSN As Long = -1, Optional ByVal blnGetLapseNO As Boolean = False) As Long
'����:��ȡ��ǰ�ű�����������
'     ȫ��������ʱ����0
'    blngetlapseNo:�Ƿ����Ч���Ժ�ʼ��
'     lngCurMaxSN-�������ʹ�ú�
    Dim i           As Integer
    Dim j           As Integer
    Dim lngMaxSn    As Long
    Dim lngSN       As Long
    Dim intStart    As Integer
    Dim lngTmp      As Long
    Dim blnUnitReg  As Boolean
    Dim lngMaxLapse As Long '�����Ч����
    On Error GoTo errH
    If Not mrsSNState Is Nothing Or blnUnitReg Then
ReGet:
        mrsSNState.Filter = ""
        If mrsSNState.RecordCount > 0 Or blnUnitReg Then
            If lngCurMaxSN = -1 And mViewMode = v_ר�Һŷ�ʱ�� Then
                With vsfDetailTime
                    i = vsfDetailTime.Row
                    j = vsfDetailTime.Col
                    If .TextMatrix(i, j) <> "" Then
                        If .Cell(flexcpForeColor, i, j) <> vbRed And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGreen And .Cell(flexcpForeColor, i, j) <> vbGrayText And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then
                           lngTmp = Val(Getʱ��(i, j, False))
                           mrsSNState.Filter = "���=" & lngTmp
                            If mrsSNState.RecordCount = 0 And lngTmp > lngMaxLapse Then
                                    GetCurrSN = lngTmp
                                    Exit Function
                            End If
                        End If
                    End If
                End With
            End If
            
            
           If lngCurMaxSN = -1 And mViewMode = v_ר�Һ� Then
               lngTmp = 0
               mrsSNState.Filter = "ԤԼ=0 and ״̬=1"
                Do While Not mrsSNState.EOF
                   If lngTmp < Val(mrsSNState!���) Then lngTmp = Val(mrsSNState!���)
                   mrsSNState.MoveNext
                Loop
                
                'mrsSNState.MoveFirst
                mrsSNState.Filter = 0
               If lngTmp <> 0 Then lngCurMaxSN = lngTmp
            End If
            
            intStart = IIf(mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = V_��ͨ�ŷ�ʱ��, 1, 0)
            For i = 0 To vsfDetailTime.Rows - 1
                For j = intStart To vsfDetailTime.Cols - 1
                    Select Case mViewMode
                    Case V_��ͨ��, v_ר�Һ�:
                        lngSN = Val(vsfDetailTime.TextMatrix(i, j))
                    Case v_ר�Һŷ�ʱ��:
                        With vsfDetailTime
                            If .Cell(flexcpForeColor, i, j) = vbGrayText Or .Cell(flexcpForeColor, i, j) = &HC000C0 Then
                                lngSN = -1
                            Else
                               lngSN = IIf(Trim(.TextMatrix(i, j)) = "", -1, Val(Getʱ��(i, j, False)))
                               If lngSN < lngMaxLapse And mty_Para.bln������ѡ�� = False Then lngSN = -1
                            End If
                        End With
                    Case Else
                       Exit Function
                    End Select
                    If lngSN > -1 Then
                        mrsSNState.Filter = "���=" & lngSN
                        If mrsSNState.RecordCount = 0 Then
                            lngMaxSn = lngSN
                            vsfDetailTime.Select i, j
                            Exit For
                        End If
                    End If
                    
                Next
                
                If lngMaxSn = lngSN Then Exit For
            Next
            If lngCurMaxSN > 0 And lngMaxSn = 0 Then
                '���˺�:???
                '��Ҫ�ǽ��ԤԼ���+1��,����ԤԼ�����,�����ִ�1��ʼ����Ƿ���δѡ���.
                '��:ԤԼ��5��ʼ;����7�Ѿ���������,����ٴ�1��ʼȡ.
               ' lngCurMaxSN = -1: GoTo ReGet:
            End If
            GetCurrSN = lngMaxSn
        Else
            Select Case mViewMode
                Case v_ר�Һŷ�ʱ��:
                     vsfDetailTime.Redraw = False
                    For i = 0 To vsfDetailTime.Rows - 1
                        For j = 1 To vsfDetailTime.Cols - 1
                            If vsfDetailTime.Cell(flexcpForeColor, i, j) <> vbGrayText And vsfDetailTime.Cell(flexcpForeColor, i, j) <> &HC000C0 And vsfDetailTime.TextMatrix(i, j) <> "" Then
                                GetCurrSN = Val(Getʱ��(i, j, False))
                                vsfDetailTime.Redraw = True
                                Exit Function
                            End If
                        Next
                    Next
                    vsfDetailTime.Redraw = True
                Case Else:
                    GetCurrSN = 1
            End Select
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

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

Private Sub vsfArrange_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call vsfArrange_EnterCell
        If cboDoctor.Enabled Then
            cboDoctor.SetFocus
        Else
            If cboRoom.Enabled Then
                cboRoom.SetFocus
            Else
                chkBook.SetFocus
            End If
        End If
    End If
End Sub

Private Sub vsfDetailTime_Click()
    Call vsfDetailTime_DblClick
End Sub

Private Sub vsfDetailTime_KeyDown(KeyCode As Integer, Shift As Integer)
     If mty_Para.bln������ѡ�� Then Exit Sub
     If KeyCode <> 13 Then KeyCode = 0
End Sub

Private Sub vsfDetailTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <= vsfDetailTime.Rows - 1 And OldCol <= vsfDetailTime.Cols - 1 Then
        vsfDetailTime.Cell(flexcpBackColor, OldRow, OldCol) = &H80000005
        If OldRow = 0 And OldCol = 0 And InStr(vsfDetailTime.TextMatrix(OldRow, OldCol), ":") > 0 Then
            vsfDetailTime.Cell(flexcpBackColor, OldRow, OldCol) = &H8000000F
        End If
    End If
    If NewRow <= vsfDetailTime.Rows - 1 And NewCol <= vsfDetailTime.Cols - 1 Then
        If vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) = vbBlack And vsfDetailTime.Cell(flexcpBackColor, NewRow, NewCol) <> -2147483633 Then
            vsfDetailTime.Cell(flexcpBackColor, NewRow, NewCol) = &H8000000D
        End If
    End If
End Sub

Private Sub vsfDetailTime_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnChangeByCode Then Exit Sub
    If (mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = v_ר�Һ�) And mty_Para.bln������ѡ�� = False _
        And vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then
        Cancel = True
        Exit Sub
    End If
    If vsfDetailTime.TextMatrix(NewRow, NewCol) = "" Then Cancel = True
    If vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlack And vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then Cancel = True
    If Not CheckAddAvailable Then
        If vsfDetailTime.Cell(flexcpData, NewRow, NewCol) Like "��*" Then Cancel = True
    End If
End Sub

Private Sub vsfDetailTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsfDetailTime_DblClick
        If cboPayMode.Visible And cboPayMode.Enabled Then
            cboPayMode.SetFocus
        Else
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Function CheckAddAvailable() As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'����:��鵱ǰѡ��ĺű�Ӻ��Ƿ����
'����:���÷���True,�����÷���False
'����:������
'����:2014-01-15
'��ע:
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim intTotal As Integer, intUse As Integer
    On Error GoTo errH
    intTotal = 0
    intUse = 0
    'ֻ�Է�ʱ�ν��д���
    If mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then
        With vsfDetailTime
            For j = 1 To .Cols - 1
                For i = 0 To .Rows - 1
                    If .TextMatrix(i, j) <> "" And Not .Cell(flexcpData, i, j) Like "��*" Then
                        intTotal = intTotal + 1
                        If .Cell(flexcpForeColor, i, j) <> vbBlack Then
                            intUse = intUse + 1
                        End If
                    End If
                Next i
            Next j
        End With
        If intUse = intTotal Then CheckAddAvailable = True: Exit Function
        CheckAddAvailable = False
        Exit Function
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function zlGet��ǰ���ڼ�(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���������ڼ�
    '����:���˺�
    '����:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bln��ǰ���� As Boolean, strTemp As String
    On Error GoTo errH
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
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub vsfDetailTime_DblClick()
    Dim lngSN       As Long
    Dim datThis     As Date
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Or V_��ͨ�ŷ�ʱ�� Then
        '*************************************************
        '��ͨ�ź�û�з�ʱ�ε�ר�Һ� ������ǰ������
        '*************************************************
        dtpTime.Enabled = True
        mblnChangeByCode = True
        txtArrangeNO.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�"))
        mblnChangeByCode = False
        strTmp = zlGet��ǰ���ڼ�
        strTmp = vsfArrange.Cell(flexcpData, vsfArrange.Row, vsfArrange.ColIndex(strTmp))
        strSQL = "Select ��ʼʱ�� From ʱ��� Where ʱ��� = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, strTmp)
        If rsTmp.RecordCount <> 0 Then
            dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
            dtpTime.minDate = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
        End If
        Exit Sub
    End If
    
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then Exit Sub
    '*************************************************
    '��ʱ�� �����µķ�ʽ������
    '*************************************************
    dtpTime.Enabled = False
    Select Case mViewMode
    Case V_��ͨ�ŷ�ʱ��:
         If vsfDetailTime.CellForeColor = vbGrayText Then Exit Sub
         If vsfDetailTime.TextMatrix(vsfDetailTime.Row, vsfDetailTime.Col) = "" Then Exit Sub
         If Val(Getʱ��(vsfDetailTime.Row, vsfDetailTime.Col, False)) = 0 Then Exit Sub
         strTmp = Getʱ��(vsfDetailTime.Row, vsfDetailTime.Col, True)
         datThis = CDate(Format(strTmp, "hh:mm"))
         dtpTime.Value = datThis
         If datThis < CDate(Format(gobjDatabase.CurrentDate, "hh:mm:ss")) Then
            dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
            dtpTime.minDate = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
         End If
         mblnChangeByCode = True
         txtArrangeNO.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�"))
         mblnChangeByCode = False
    Case v_ר�Һŷ�ʱ��:
        '**********************************************
        '������Ϊ�ѹһ�����Լ�Ĳ�����ѡ��
        '**********************************************
        If vsfDetailTime.Row > vsfDetailTime.Rows - 1 Or vsfDetailTime.Col > vsfDetailTime.Cols - 1 Then Exit Sub
        If vsfDetailTime.TextMatrix(vsfDetailTime.Row, vsfDetailTime.Col) = "" Then Exit Sub
        If vsfDetailTime.CellForeColor = vbRed Or vsfDetailTime.CellForeColor = vbGreen Or vsfDetailTime.CellForeColor = vbGrayText Or vsfDetailTime.CellForeColor = &HC000C0 Then Exit Sub  '--And .CellForeColor <> vbBlue
        strTmp = Getʱ��(vsfDetailTime.Row, vsfDetailTime.Col, True)
        If strTmp <> "" Then
            datThis = CDate(Format(strTmp, "hh:mm"))
            dtpTime.Value = datThis
        End If
        If datThis < CDate(Format(gobjDatabase.CurrentDate, "hh:mm:ss")) Then
            dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
            dtpTime.minDate = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
        End If
        mblnChangeByCode = True
        txtArrangeNO.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("�ű�"))
        mblnChangeByCode = False
    Case Else
        Exit Sub
    End Select
    
End Sub

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
