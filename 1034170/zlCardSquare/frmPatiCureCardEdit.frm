VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Object = "*\A..\ZlPatiAddress\ZlPatiAddress.vbp"
Begin VB.Form frmPatiCureCardEdit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˷�������"
   ClientHeight    =   9450
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11610
   Icon            =   "frmPatiCureCardEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDrugAllergy 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   11820
      ScaleHeight     =   3255
      ScaleWidth      =   6840
      TabIndex        =   155
      Top             =   1200
      Width           =   6840
      Begin VB.CommandButton cmdSelDrug 
         Caption         =   "��"
         Height          =   300
         Left            =   600
         TabIndex        =   156
         Top             =   540
         Visible         =   0   'False
         Width           =   300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDrug 
         Height          =   3060
         Left            =   -30
         TabIndex        =   157
         Top             =   240
         Width           =   5895
         _cx             =   1973037214
         _cy             =   1973032213
         Appearance      =   1
         BorderStyle     =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   End
   Begin VB.PictureBox picTaskPanelOther 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   10350
      ScaleHeight     =   1125
      ScaleWidth      =   1215
      TabIndex        =   153
      Top             =   2910
      Visible         =   0   'False
      Width           =   1215
      Begin XtremeSuiteControls.TaskPanel wndTaskPanelOther 
         Height          =   945
         Left            =   0
         TabIndex        =   154
         Top             =   0
         Width           =   1035
         _Version        =   589884
         _ExtentX        =   1826
         _ExtentY        =   1667
         _StockProps     =   64
         VisualTheme     =   7
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin VB.PictureBox picԤ����� 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   8000
      ScaleHeight     =   225
      ScaleWidth      =   2205
      TabIndex        =   151
      Top             =   7380
      Visible         =   0   'False
      Width           =   2200
      Begin VB.Label lblԤ����� 
         Caption         =   "Ԥ�����:0Ԫ"
         ForeColor       =   &H000000FF&
         Height          =   220
         Left            =   0
         TabIndex        =   152
         Top             =   0
         Width           =   2200
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   104
      Top             =   9090
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   176
            Picture         =   "frmPatiCureCardEdit.frx":000C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15505
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   10680
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOtherInfo 
      BorderStyle     =   0  'None
      Height          =   4080
      Left            =   10500
      ScaleHeight     =   4080
      ScaleWidth      =   10110
      TabIndex        =   127
      Top             =   4620
      Width           =   10110
      Begin VB.CommandButton cmdMedicalWarning 
         Caption         =   "��"
         Height          =   330
         Left            =   9465
         TabIndex        =   144
         Top             =   0
         Width           =   315
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   300
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   0
         Width           =   1410
      End
      Begin VB.ComboBox cboBH 
         Height          =   300
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   0
         Width           =   1410
      End
      Begin VB.TextBox txtMedicalWarning 
         Height          =   350
         Left            =   5535
         Locked          =   -1  'True
         TabIndex        =   131
         Top             =   0
         Width           =   4260
      End
      Begin VB.TextBox txtOtherWaring 
         Height          =   350
         Left            =   1275
         MaxLength       =   100
         TabIndex        =   130
         Top             =   375
         Width           =   8535
      End
      Begin VB.Frame frameLinkMan 
         BackColor       =   &H80000004&
         Height          =   105
         Left            =   1065
         TabIndex        =   129
         Top             =   840
         Width           =   8895
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         Height          =   105
         Left            =   885
         TabIndex        =   128
         Top             =   2160
         Width           =   9135
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLinkMan 
         Height          =   975
         Left            =   60
         TabIndex        =   136
         Top             =   1080
         Width           =   9705
         _cx             =   1973043935
         _cy             =   1973028536
         Appearance      =   1
         BorderStyle     =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483634
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsOtherInfo 
         Height          =   1380
         Left            =   60
         TabIndex        =   137
         Top             =   2400
         Width           =   9705
         _cx             =   1973043935
         _cy             =   1973029250
         Appearance      =   1
         BorderStyle     =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPatiCureCardEdit.frx":08A0
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
      Begin VB.Label lblBloodType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ѫ��"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   525
         TabIndex        =   143
         Top             =   45
         Width           =   1020
      End
      Begin VB.Label lblBH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2535
         TabIndex        =   142
         Top             =   45
         Width           =   885
      End
      Begin VB.Label lblMedicalWarning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽѧ��ʾ"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4215
         TabIndex        =   141
         Top             =   45
         Width           =   1860
      End
      Begin VB.Label lblOtherWaring 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽѧ��ʾ"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -270
         TabIndex        =   140
         Top             =   420
         Width           =   1860
      End
      Begin VB.Label lblLinkman 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ����Ϣ"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -360
         TabIndex        =   139
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label lblOtherInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   -450
         TabIndex        =   138
         Top             =   2145
         Width           =   1860
      End
   End
   Begin VB.PictureBox picInoculate 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   120
      ScaleHeight     =   3105
      ScaleWidth      =   5925
      TabIndex        =   125
      Top             =   9030
      Width           =   5925
      Begin VSFlex8Ctl.VSFlexGrid vsInoculate 
         Height          =   3015
         Left            =   540
         TabIndex        =   126
         Top             =   210
         Width           =   5895
         _cx             =   1973037214
         _cy             =   1973032134
         Appearance      =   1
         BorderStyle     =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   End
   Begin VB.CommandButton cmd����˿� 
      Caption         =   "�˿�(&B)"
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
      Left            =   10350
      TabIndex        =   116
      Top             =   1845
      Width           =   1100
   End
   Begin VB.CommandButton cmd��ֵ 
      Caption         =   "��ֵ(&Z)"
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
      Left            =   10350
      TabIndex        =   115
      Top             =   1425
      Width           =   1100
   End
   Begin VB.PictureBox picTittle 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   150
      ScaleHeight     =   495
      ScaleWidth      =   9945
      TabIndex        =   105
      Top             =   240
      Width           =   9945
      Begin VB.CheckBox chkCancel 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9405
         Style           =   1  'Graphical
         TabIndex        =   109
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F8"
         Top             =   15
         Width           =   405
      End
      Begin VB.Frame fraSplit 
         Caption         =   "Frame1"
         Height          =   150
         Left            =   -750
         TabIndex        =   106
         Top             =   345
         Width           =   12990
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
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   111
         ToolTipText     =   "�ȼ�:F12"
         Top             =   45
         Width           =   1620
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -645
         TabIndex        =   114
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7080
         TabIndex        =   113
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.PictureBox picCard 
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   90
      ScaleHeight     =   1635
      ScaleWidth      =   9975
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   7650
      Width           =   9975
      Begin VB.Frame fraCard 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1545
         Left            =   30
         TabIndex        =   150
         Top             =   30
         Width           =   9795
         Begin VB.TextBox txt�ϼ� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   8355
            MaxLength       =   16
            TabIndex        =   89
            Top             =   660
            Width           =   1260
         End
         Begin VB.TextBox txt��� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6420
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   161
            Top             =   1100
            Width           =   3210
         End
         Begin VB.CommandButton cmdReadCard 
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
            Height          =   330
            Left            =   4845
            TabIndex        =   75
            TabStop         =   0   'False
            Tag             =   "����"
            Top             =   215
            Width           =   615
         End
         Begin VB.TextBox txt���� 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1085
            MaxLength       =   20
            TabIndex        =   74
            Tag             =   "����"
            Top             =   205
            Width           =   3780
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1085
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   85
            TabStop         =   0   'False
            Tag             =   "����"
            Top             =   650
            Width           =   2165
         End
         Begin VB.TextBox txtAudi 
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
            IMEMode         =   3  'DISABLE
            Left            =   8355
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   81
            Tag             =   "��֤"
            Top             =   205
            Width           =   1260
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "���Ѽ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3560
            TabIndex        =   86
            Top             =   690
            Width           =   1215
         End
         Begin VB.TextBox txtPass 
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
            IMEMode         =   3  'DISABLE
            Left            =   6420
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   79
            Tag             =   "����"
            Top             =   205
            Width           =   1125
         End
         Begin VB.ComboBox cbo֧����ʽ 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   330
            Left            =   6420
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   660
            Width           =   1935
         End
         Begin VB.TextBox txt����Ա 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1085
            Locked          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   1100
            Width           =   1080
         End
         Begin VB.TextBox txt�䶯ԭ�� 
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
            Left            =   1100
            MaxLength       =   100
            TabIndex        =   83
            Tag             =   "�䶯ԭ��"
            Top             =   660
            Visible         =   0   'False
            Width           =   8550
         End
         Begin VB.TextBox txtԭ������ 
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
            IMEMode         =   3  'DISABLE
            Left            =   6420
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   133
            Tag             =   "����"
            Top             =   205
            Visible         =   0   'False
            Width           =   1125
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   3075
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   1100
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
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
         Begin zlIDKind.IDKindNew IDKindPay 
            Height          =   360
            Left            =   500
            TabIndex        =   148
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   $"frmPatiCureCardEdit.frx":0902
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
            OlnyReadCardNo  =   0   'False
            BackColor       =   -2147483633
         End
         Begin VB.TextBox txtˢ������ 
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
            Left            =   6420
            TabIndex        =   77
            Tag             =   "ˢ������"
            Top             =   205
            Width           =   3210
         End
         Begin VB.ComboBox cbo��ʧ��ʽ 
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
            Left            =   6420
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   215
            Visible         =   0   'False
            Width           =   3225
         End
         Begin zlIDKind.IDKindNew IDKindPayMode 
            Height          =   360
            Left            =   5535
            TabIndex        =   162
            Top             =   1095
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
            ShowSortName    =   0   'False
            IDKindStr       =   "Ӧ��|Ӧ��|0|0|0|0|0|0|0|0|0;��ֵ|��ֵ|0|0|0|0|0|0|0|0|0"
            CaptionAlignment=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   10.5
            FontName        =   "����"
            IDKind          =   -1
            DefaultCardType =   "0"
            NotAutoAppendKind=   -1  'True
            BackColor       =   -2147483633
         End
         Begin VB.Label lbl֧����ʽ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɿʽ"
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
            Left            =   5925
            TabIndex        =   87
            Top             =   720
            Width           =   420
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   70
            TabIndex        =   73
            Top             =   260
            Width           =   450
         End
         Begin VB.Label lbl��֤ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��֤"
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
            Left            =   7670
            TabIndex        =   80
            Top             =   275
            Width           =   420
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   590
            TabIndex        =   84
            Top             =   720
            Width           =   420
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   5940
            TabIndex        =   78
            Top             =   275
            Width           =   420
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
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
            Left            =   425
            TabIndex        =   112
            Top             =   1175
            Width           =   630
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
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
            Left            =   2205
            TabIndex        =   110
            Top             =   1175
            Width           =   840
         End
         Begin VB.Label lblˢ����֤ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " ˢ����֤"
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
            Left            =   5430
            TabIndex        =   76
            Top             =   275
            Width           =   945
         End
         Begin VB.Label lblԭ������ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ԭ������"
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
            Left            =   5520
            TabIndex        =   135
            Top             =   275
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin XtremeSuiteControls.TabControl tbPageDo 
         Height          =   240
         Left            =   180
         TabIndex        =   149
         Top             =   330
         Width           =   420
         _Version        =   589884
         _ExtentX        =   741
         _ExtentY        =   423
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBasePati 
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   90
      ScaleHeight     =   2235
      ScaleWidth      =   9990
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   765
      Width           =   9990
      Begin VB.Frame fra 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   60
         TabIndex        =   96
         Top             =   -15
         Width           =   9840
         Begin VB.TextBox txt�ֻ� 
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
            IMEMode         =   3  'DISABLE
            Left            =   6150
            MaxLength       =   20
            TabIndex        =   9
            Tag             =   "�ֻ���"
            Top             =   570
            Width           =   1575
         End
         Begin VB.TextBox txt���ڵ�ַ�ʱ� 
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
            IMEMode         =   3  'DISABLE
            Left            =   6945
            MaxLength       =   6
            TabIndex        =   21
            Tag             =   "���ڵ�ַ�ʱ�"
            Top             =   1755
            Width           =   780
         End
         Begin VB.TextBox txt���� 
            Height          =   330
            Left            =   3765
            TabIndex        =   13
            Text            =   "����"
            Top             =   975
            Width           =   555
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   615
            Width           =   1245
         End
         Begin VB.TextBox txt��ͥ�ʱ� 
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6945
            MaxLength       =   6
            TabIndex        =   18
            Tag             =   "��ͥ��ַ�ʱ�"
            Top             =   1380
            Width           =   780
         End
         Begin VB.PictureBox picPatient 
            Height          =   1500
            Left            =   7830
            ScaleHeight     =   1440
            ScaleWidth      =   1815
            TabIndex        =   124
            Top             =   180
            Width           =   1875
            Begin VB.Image imgPatient 
               Height          =   1425
               Left            =   15
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1800
            End
         End
         Begin VB.CommandButton cmdPicCollect 
            Caption         =   "�ɼ�"
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
            Left            =   8445
            TabIndex        =   122
            Top             =   1710
            Width           =   600
         End
         Begin VB.CommandButton cmdPicFile 
            Caption         =   "�ļ�"
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
            Left            =   7815
            TabIndex        =   121
            Top             =   1710
            Width           =   585
         End
         Begin VB.CommandButton cmdPicClear 
            Caption         =   "���"
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
            TabIndex        =   120
            Top             =   1710
            Width           =   600
         End
         Begin VB.TextBox txtPatient 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1140
            TabIndex        =   0
            Tag             =   "����"
            Top             =   180
            Width           =   1935
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   555
            TabIndex        =   118
            Top             =   180
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   $"frmPatiCureCardEdit.frx":0991
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
         Begin VB.TextBox txt����� 
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
            IMEMode         =   3  'DISABLE
            Left            =   6150
            MaxLength       =   18
            TabIndex        =   4
            Tag             =   "�����"
            Top             =   195
            Width           =   1575
         End
         Begin VB.ComboBox cbo���䵥λ 
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
            Left            =   4305
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   975
            Width           =   690
         End
         Begin VB.ComboBox cbo�Ա� 
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
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   210
            Width           =   1245
         End
         Begin VB.TextBox txt���֤�� 
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
            IMEMode         =   3  'DISABLE
            Left            =   1110
            MaxLength       =   18
            TabIndex        =   6
            Tag             =   "���֤��"
            Text            =   "012345678901234567"
            Top             =   600
            Width           =   1965
         End
         Begin VB.TextBox txt��ͥ�绰 
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
            IMEMode         =   3  'DISABLE
            Left            =   6150
            MaxLength       =   20
            TabIndex        =   15
            Tag             =   "��ͥ�绰"
            Top             =   975
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txt����ʱ�� 
            Height          =   345
            Left            =   2280
            TabIndex        =   12
            Tag             =   "����ʱ��"
            Top             =   1005
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt�������� 
            Height          =   345
            Left            =   1125
            TabIndex        =   11
            Tag             =   "��������"
            Top             =   1005
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "YYYY-MM-DD"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin ZlPatiAddress.PatiAddress padd��ͥ��ַ 
            Height          =   330
            Left            =   1125
            TabIndex        =   17
            Tag             =   "��סַ"
            Top             =   1395
            Visible         =   0   'False
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padd���ڵ�ַ 
            Height          =   330
            Left            =   1125
            TabIndex        =   20
            Tag             =   "���ڵ�ַ"
            Top             =   1770
            Visible         =   0   'False
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   100
         End
         Begin VB.CommandButton cmd��ͥ��ַ 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5610
            TabIndex        =   22
            TabStop         =   0   'False
            Tag             =   "��סַ"
            ToolTipText     =   "�ȼ���F3"
            Top             =   1410
            Width           =   300
         End
         Begin VB.CommandButton cmd���ڵ�ַ 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5610
            TabIndex        =   160
            TabStop         =   0   'False
            Tag             =   "���ڵ�ַ"
            Top             =   1785
            Width           =   300
         End
         Begin VB.TextBox txt��ͥ��ַ 
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
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   16
            Tag             =   "��סַ"
            Top             =   1395
            Width           =   4785
         End
         Begin VB.TextBox txt���ڵ�ַ 
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
            Left            =   1125
            TabIndex        =   19
            Tag             =   "���ڵ�ַ"
            Top             =   1770
            Width           =   4785
         End
         Begin VB.Label lbl�ֻ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ֻ���"
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
            Left            =   5430
            TabIndex        =   8
            Top             =   615
            Width           =   630
         End
         Begin VB.Label lbl���ڵ�ַ 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "���ڵ�ַ"
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
            Left            =   270
            TabIndex        =   159
            Top             =   1830
            Width           =   840
         End
         Begin VB.Label lbl���ڵ�ַ�ʱ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ʱ�"
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
            Left            =   6075
            TabIndex        =   158
            Top             =   1830
            Width           =   840
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   3285
            TabIndex        =   146
            Top             =   660
            Width           =   420
         End
         Begin VB.Label lbl��ͥ�ʱ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�"
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
            Left            =   6450
            TabIndex        =   145
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   123
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lbl����� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   5430
            TabIndex        =   3
            Top             =   255
            Width           =   630
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
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
            Left            =   240
            TabIndex        =   10
            Top             =   1050
            Width           =   840
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   3300
            TabIndex        =   23
            Top             =   1035
            Width           =   420
         End
         Begin VB.Label lbl�Ա� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   3195
            TabIndex        =   1
            Top             =   270
            Width           =   420
         End
         Begin VB.Label lbl���֤�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���֤��"
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
            Left            =   255
            TabIndex        =   5
            Top             =   660
            Width           =   840
         End
         Begin VB.Label lbl��ͥ�绰 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ͥ�绰"
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
            Left            =   5205
            TabIndex        =   25
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label lbl��ͥ��ַ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��סַ"
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
            Left            =   255
            TabIndex        =   24
            Top             =   1425
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox picExpend 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4250
      Left            =   75
      ScaleHeight     =   4245
      ScaleWidth      =   10005
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   3135
      Width           =   10005
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   390
         Left            =   30
         TabIndex        =   119
         Top             =   240
         Width           =   270
         _Version        =   589884
         _ExtentX        =   476
         _ExtentY        =   688
         _StockProps     =   64
      End
      Begin VB.Frame fraBase 
         Height          =   3825
         Left            =   90
         TabIndex        =   97
         Top             =   60
         Width           =   9855
         Begin VB.TextBox txt��ϵ�����֤�� 
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
            IMEMode         =   3  'DISABLE
            Left            =   1365
            MaxLength       =   18
            TabIndex        =   65
            Tag             =   "���֤��"
            Top             =   3075
            Width           =   3270
         End
         Begin VB.CommandButton cmd�����ص� 
            Caption         =   "��"
            Height          =   255
            Left            =   4320
            TabIndex        =   51
            TabStop         =   0   'False
            Tag             =   "�����ص�"
            ToolTipText     =   "�ȼ���F3"
            Top             =   1958
            Width           =   285
         End
         Begin VB.TextBox txt��λ�ʻ� 
            Height          =   300
            Left            =   1155
            MaxLength       =   100
            TabIndex        =   62
            Tag             =   "��λ�ʻ�"
            Top             =   2730
            Width           =   3480
         End
         Begin VB.TextBox txt��λ������ 
            Height          =   300
            Left            =   5835
            MaxLength       =   100
            TabIndex        =   60
            Tag             =   "��λ������"
            Top             =   2340
            Width           =   3885
         End
         Begin VB.CommandButton cmd���� 
            Caption         =   "��"
            Height          =   255
            Left            =   9420
            TabIndex        =   48
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "�ȼ���F3"
            Top             =   1545
            Width           =   285
         End
         Begin VB.TextBox txt����֤�� 
            Height          =   300
            Left            =   1155
            MaxLength       =   20
            TabIndex        =   45
            Tag             =   "����֤��"
            Top             =   1530
            Width           =   3480
         End
         Begin VB.ComboBox cbo�ѱ� 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   690
            Width           =   1485
         End
         Begin VB.ComboBox cbo��� 
            Height          =   300
            Left            =   8250
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   705
            Width           =   1470
         End
         Begin VB.ComboBox cboְҵ 
            Height          =   300
            Left            =   5835
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1125
            Width           =   3885
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   690
            Width           =   1485
         End
         Begin VB.ComboBox cboѧ�� 
            Height          =   300
            Left            =   3150
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1125
            Width           =   1485
         End
         Begin VB.ComboBox cbo����״�� 
            Height          =   300
            Left            =   5835
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   705
            Width           =   1485
         End
         Begin VB.ComboBox cbo��ϵ�˹�ϵ 
            Height          =   300
            Left            =   8250
            TabIndex        =   67
            Top             =   3120
            Width           =   1470
         End
         Begin VB.CommandButton cmd��ͬ��λ 
            Caption         =   "��"
            Height          =   255
            Left            =   9420
            TabIndex        =   54
            TabStop         =   0   'False
            Tag             =   "������λ"
            ToolTipText     =   "�ȼ���F3"
            Top             =   1950
            Width           =   285
         End
         Begin VB.CommandButton cmd��ϵ�˵�ַ 
            Caption         =   "��"
            Height          =   255
            Left            =   9405
            TabIndex        =   70
            TabStop         =   0   'False
            Tag             =   "��ϵ�˵�ַ"
            ToolTipText     =   "�ȼ���F3"
            Top             =   3480
            Width           =   285
         End
         Begin VB.ComboBox cboҽ�Ƹ��� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1125
            Width           =   1485
         End
         Begin VB.TextBox txt������λ 
            Height          =   300
            Left            =   5835
            MaxLength       =   100
            TabIndex        =   53
            Tag             =   "������λ"
            Top             =   1935
            Width           =   3885
         End
         Begin VB.TextBox txt�����ص� 
            Height          =   300
            Left            =   1155
            MaxLength       =   30
            TabIndex        =   50
            Tag             =   "�����ص�"
            Top             =   1935
            Width           =   3480
         End
         Begin VB.TextBox txt��λ�绰 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            MaxLength       =   20
            TabIndex        =   56
            Tag             =   "��λ�绰"
            Top             =   2340
            Width           =   1485
         End
         Begin VB.TextBox txt��ϵ�˵绰 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5835
            MaxLength       =   20
            TabIndex        =   66
            Tag             =   "��ϵ�˵绰"
            Top             =   3120
            Width           =   1365
         End
         Begin VB.TextBox txt��λ�ʱ� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3585
            MaxLength       =   6
            TabIndex        =   58
            Tag             =   "��λ�ʱ�"
            Top             =   2340
            Width           =   1035
         End
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   5835
            MaxLength       =   30
            TabIndex        =   47
            Tag             =   "����"
            Top             =   1530
            Width           =   3885
         End
         Begin VB.TextBox txtҽ���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            MaxLength       =   30
            TabIndex        =   27
            Tag             =   "ҽ����"
            Top             =   285
            Width           =   3480
         End
         Begin VB.TextBox txt��ϵ������ 
            Height          =   300
            Left            =   5835
            MaxLength       =   64
            TabIndex        =   64
            Tag             =   "��ϵ������"
            Top             =   2730
            Width           =   3870
         End
         Begin VB.TextBox txt��ϵ�˵�ַ 
            Height          =   300
            Left            =   1170
            MaxLength       =   64
            TabIndex        =   69
            Tag             =   "��ϵ�˵�ַ"
            Top             =   3465
            Width           =   8535
         End
         Begin VB.TextBox txt��֤ҽ���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5835
            MaxLength       =   30
            TabIndex        =   29
            Tag             =   "��֤ҽ����"
            Top             =   285
            Width           =   3870
         End
         Begin VB.Label lbl��ϵ�����֤�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�����֤��"
            Height          =   180
            Left            =   45
            TabIndex        =   117
            Top             =   3165
            Width           =   1260
         End
         Begin VB.Label lblҽ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��֤ҽ����"
            Height          =   180
            Index           =   1
            Left            =   4845
            TabIndex        =   28
            Top             =   345
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ�ʻ�"
            Height          =   180
            Left            =   390
            TabIndex        =   61
            Top             =   2790
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�˵�ַ"
            Height          =   180
            Left            =   210
            TabIndex        =   68
            Top             =   3525
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ������"
            Height          =   180
            Left            =   4860
            TabIndex        =   59
            Top             =   2400
            Width           =   900
         End
         Begin VB.Label lbl��ע 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ע"
            Height          =   180
            Left            =   5220
            TabIndex        =   102
            Top             =   3840
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lblPatiColor 
            Height          =   255
            Left            =   9060
            TabIndex        =   101
            Top             =   2700
            Width           =   105
         End
         Begin VB.Label lbl����֤�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����֤��"
            Height          =   180
            Left            =   390
            TabIndex        =   44
            Top             =   1590
            Width           =   720
         End
         Begin VB.Label lbl�ѱ� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�"
            Height          =   180
            Left            =   750
            TabIndex        =   30
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl�����ص� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ص�"
            Height          =   180
            Left            =   390
            TabIndex        =   49
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            Height          =   180
            Left            =   7830
            TabIndex        =   42
            Top             =   765
            Width           =   360
         End
         Begin VB.Label lblְҵ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ְҵ"
            Height          =   180
            Left            =   5400
            TabIndex        =   40
            Top             =   1185
            Width           =   360
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   2730
            TabIndex        =   32
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lblѧ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ѧ��"
            Height          =   180
            Left            =   2730
            TabIndex        =   38
            Top             =   1185
            Width           =   360
         End
         Begin VB.Label lbl����״�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   5385
            TabIndex        =   34
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl��ϵ������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ������"
            Height          =   180
            Left            =   4845
            TabIndex        =   63
            Top             =   2790
            Width           =   900
         End
         Begin VB.Label lbl��ϵ�˹�ϵ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�˹�ϵ"
            Height          =   180
            Left            =   7320
            TabIndex        =   72
            Top             =   3180
            Width           =   900
         End
         Begin VB.Label lbl��ϵ�˵�ַ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�˵�ַ"
            Height          =   180
            Left            =   135
            TabIndex        =   100
            Top             =   3840
            Width           =   900
         End
         Begin VB.Label lbl��ϵ�˵绰 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�˵绰"
            Height          =   180
            Left            =   4845
            TabIndex        =   71
            Top             =   3180
            Width           =   900
         End
         Begin VB.Label lbl������λ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������λ"
            Height          =   180
            Left            =   5025
            TabIndex        =   52
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label lbl��λ�绰 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ�绰"
            Height          =   180
            Left            =   390
            TabIndex        =   55
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label lbl��λ�ʱ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ�ʱ�"
            Height          =   180
            Left            =   2760
            TabIndex        =   57
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label lbl��λ������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ������"
            Height          =   180
            Left            =   135
            TabIndex        =   99
            Top             =   4200
            Width           =   900
         End
         Begin VB.Label lbl��λ�ʺ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ�ʺ�"
            Height          =   180
            Left            =   4860
            TabIndex        =   98
            Top             =   4200
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ�Ƹ���"
            Height          =   180
            Index           =   1
            Left            =   390
            TabIndex        =   36
            Top             =   1185
            Width           =   720
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   5385
            TabIndex        =   46
            Top             =   1590
            Width           =   360
         End
         Begin VB.Label lblҽ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ����"
            Height          =   180
            Index           =   0
            Left            =   570
            TabIndex        =   26
            Top             =   345
            Width           =   540
         End
      End
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
      Left            =   10350
      TabIndex        =   91
      Top             =   585
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
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
      Left            =   10350
      TabIndex        =   92
      Top             =   7590
      Width           =   1100
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   8925
      Left            =   180
      TabIndex        =   93
      Top             =   0
      Width           =   10125
      _Version        =   589884
      _ExtentX        =   17859
      _ExtentY        =   15743
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
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
      Left            =   10350
      TabIndex        =   90
      Top             =   150
      Width           =   1100
   End
   Begin VB.CommandButton cmdCreateCard 
      Caption         =   "�ƿ�(&M)"
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
      Left            =   10350
      TabIndex        =   147
      Top             =   1005
      Width           =   1100
   End
End
Attribute VB_Name = "frmPatiCureCardEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
'��ڲ���
Private mstrPrivs As String, mlngModule As Long
Private mlngCardTypeID As Long, mstrCardNo As String
Public Enum gCardType
    Cr_���� = 0
    Cr_�˿� = 1
    Cr_�󶨿� = 2
    Cr_ȡ���� = 3
    Cr_���� = 4
    Cr_���� = 5
    Cr_��ʧ = 6
    Cr_��ѯ = 7
    Cr_����������Ϣ = 8
End Enum
Private mEditType As gCardType
Private mEditTypeOld As gCardType
Private mstrBillNo  As String, mint��¼״̬   As Integer
Private mblnNOMoved As Boolean  '��ʷ����ת��
Private mblnNotClick As Boolean
Private mblnUnLoad As Boolean
Private mstrPrepayPrivs As String
Private mstrIDImageFile As String
'---------------------------------------------------------------------------------------
'ģ�����
Private mintSucces As Integer
Private Enum mTaskPancel_ID
      idx_TP_Tittle = 1
      Idx_TP_PatiBase = 2
      Idx_TP_PatiExpend = 3
      Idx_TP_PatiCard = 4
End Enum
Private Const mFormMaxHeight = 11330 '�����:51071;�����:56599
Private mblnChange As Boolean
Private Type Ty_ParaData
        blnSeekName As Boolean  '�Ƿ�ͨ����������ģ������
        intNameDays As Integer     'ģ�����ҵ�����
        blnShowExpend As Boolean '��ʾ���˵���չ��Ϣ
        int�˿�ģʽ As Integer  '0-������ˢ��;1-ˢ���˿�;2-���ݺź�����֤ˢ��;3-1��2�Ĺ���ģʽ
        bln���� As Boolean
End Type
Private mblnStructAdress As Boolean  '���˵�ַ�ṹ��¼��
Private mblnShowTown As Boolean      '�����ַ�ṹ��¼��
Private mParaData As Ty_ParaData
Private mrsInfo As ADODB.Recordset
Private WithEvents mobjIDCard As zlIDCard.clsIDCard   '���֤����
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard   'IC���ӿ�
Attribute mobjICCard.VB_VarHelpID = -1
Private mobjReadCard As Object    '���������ӿڻ�����ӿ�
Private mlngȱʡ���ų��� As Long
Private mblnICCard As Boolean
Private mlng����ID As Long
Private mblnNotChange As Boolean
Private mstrCboSplit As String
Private Type Ty_CardProperty
       lng�����ID As Long
       str������  As String
       lng���ų��� As Long
       lng���㷽ʽ As String
       bln���ƿ� As Boolean
       bln�ϸ���� As Boolean
       lng����ID As Long
       lng�������� As Long
       bln��� As Boolean
       blnˢ�� As Boolean
       bln���￨ As Boolean
       str�������� As String
       int���볤�� As Integer
       int���볤������ As Integer
       int������� As Integer
       blnOneCard As Boolean '�Ƿ�������һ��ͨ�ӿ�,��ģʽ��,Ʊ���ϸ����;Ʊ�ŷ�Χ��ķ����Ͱ󶨿����շ�
       rsҽ�ƿ��� As ADODB.Recordset
       dblӦ�ս�� As Double
       dblʵ�ս�� As Double
       bln�Ƿ��ƿ� As Boolean
       bln�Ƿ񷢿� As Boolean
       bln�Ƿ�д�� As Boolean
       bln�Ƿ�Ժ�ⷢ�� As Boolean
       lng�������� As Long '0-������,1-ͬһ������ֻ����һ�ſ�,2-ͬһ�����˿��Է����ſ�,����Ҫ���� �����:57326
       bln�Ƿ��ظ�ʹ�� As Boolean
       byt�������� As Byte '0-���ű���ﵽ���ų��ȣ������ֹ��1-������С�ڵ��ڿ��ų��ȣ�2-��������С�ڿ��ų���ʱ��鲢����
End Type
Private mCardType As Ty_CardProperty
Private mlngBillCardTypeID As Long

Private Type Ty_InsurePara
        lng���ʽҽ������ As Long   '���ʽҽ��������
End Type

Private Type TY_PayMoney
    lngҽ�ƿ����ID As Long
    bln���ѿ� As Boolean
    str���㷽ʽ As String
    str���� As String
    strˢ������ As String
    strˢ������ As String
    strNo As String
    lngID As Long 'Ԥ��ID
    lng����ID As Long
End Type

Private mCurPayMoney As TY_PayMoney
Private mInsurePara As Ty_InsurePara
Private mblnFirst As Boolean
Private mobjCardObject As clsCardObject
Private mcolPayMode As Collection
Private mstrBrushCardNo As String, mstrBrushPassWord As String
Private mcolBillBalance As Collection '�˺ŵ�����������Ϣ
Private mobjDelObject As clsCardObject
Private mintTabIndex���� As Integer '���ŵ�TabIndex
Private mintTabIndexˢ������ As Integer 'ˢ����֤��TabIndex
Private mobjKeyboard As Object '�����������

Private mblnPassInputCardNo As Boolean  '�Ƿ��������뿨��
Private mblnDefaultPassInputCardNo As Boolean 'ȱʡˢ���Ƿ��������뿨��
Private mlngҽ�ƿ�����  As Boolean
Private mblnNotCharge As Boolean
Private mstr���� As String ' ��¼�����Ƿ�仯
Private mstr���䵥λ As String 'ͬ��
'�����:56599
Private Enum mPageIndex
    ���� = 1
    ҩ����� = 2
    ������Ϣ = 3
    ������Ϣ = 4
    ������Ϣ = 5 '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
End Enum
Private mobjPlugIn As Object '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
Private mlngPlugInHwnd As Long
Private mblnPlugin As Boolean
Private mrsEMPIOut As ADODB.Recordset
Private mobjPubPatient As Object
'�����:56599
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const C_ColumHeader = "����ҩ��,1,1500,1;������ӳ,4,3000,1;����ҩ��ID,1,100,0" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_InoculateHeader = "��������,4,2000,1;��������,4,2700,1;��������,4,2000,1;��������,4,1900,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_LinkManColumHeader = "��ϵ������,4,2000,1;��ϵ�˹�ϵ,4,2700,1;��ϵ�����֤��,4,2000,1;��ϵ�˵绰,4,1900,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_OtherInfoColumHeader = "��Ϣ��,4,2000,1;��Ϣֵ,4,2700,1;��Ϣ��,4,2000,1;��Ϣֵ,4,1900,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
'Private Const C_Ѫ�� = "A��,B��,O��,AB��,����"
Private Const C_BH = "��,��,����,δ��"
Private mdicҽ�ƿ����� As New Dictionary
Private mstr�ɼ�ͼƬ As String '�ɼ�ͼƬ���ر���·��
Private mlngͼ����� As Long 'ָ����ǰ�Բ���ͼ�����������(1-�ļ� 2-�ɼ� 3-��� 4-���֤��ȡ)
Private mblnAddPage As Boolean '�Ƿ���ʾ����/�󶨿���ҳ�ؼ�
Private mblnFromCardMgr As Boolean '�Ƿ�ӷ����������
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����
Private mblnTab As Boolean
Private mbln�Զ������ As Boolean
'------------------------Ԥ������-------------------------------------
Private mFactProperty As Ty_FactProperty
Private mblnBillԤ�� As Boolean '�Ƿ��ϸ�Ʊ�ݹ���
Private mbytԤ�� As Byte 'Ʊ�ݺ��볤��
Private mstrRedFact As String 'Ԥ����Ʊ
Private mlng����ID As Long 'Ԥ������ID
Private mblnPrepayPrint As Boolean '�Ƿ��ӡԤ��Ʊ��
Private mstrPrepayInvioce As String 'Ԥ��Ʊ�ݺ�
Private mlngԤ��ID As Long '����Ԥ����¼��ID
Private mstrPrePayNo As String
Private mlngԤ������ID As Long
Private mdatԤ��ʱ�� As Date
Private mblnGetBirth As Boolean '�ж��Ƿ�ͨ�������������

'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function LoadCardData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؿ�Ƭ����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-12 11:03:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strCardPass As String, strErrMsg As String
    If mEditType = Cr_�󶨿� Or mEditType = Cr_���� Then LoadCardData = True: Exit Function
    If mlngCardTypeID = 0 Then Exit Function
    
    If mstrCardNo <> "" Then
        If GetPatiID(mlngCardTypeID, mstrCardNo, False, lng����ID, strCardPass, strErrMsg) = False Then Exit Function
        If lng����ID <= 0 Then
           Exit Function
        End If
    Else
        lng����ID = mlng����ID  '�޸Ĳ���
    End If
    If lng����ID = 0 Then LoadCardData = True: Exit Function
    If GetPatient("-" & lng����ID, False, True) = False Then Exit Function
    
    Call LoadPatiInfor: Call zlQueryEMPIPatiInfo
    If mEditType = Cr_�˿� Then
        Me.txt����.Text = GetCardNODencode(Trim(mstrCardNo), mlngCardTypeID)
        Me.lbl����.Tag = mstrCardNo
    End If

End Function
Private Function InitCardType() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������
    '����:��ʼ�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-06 17:03:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, i As Long
    Dim str���� As String, varData As Variant, varTemp As Variant, lng���￨ID As Long
    
    Err = 0: On Error GoTo errHandle
    '�����:57326
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    '76505,Ƚ����,2014-8-14,��ֹ��������޸�,ע���޸�ҽ�ƿ������������ϵͳ����Ч
    Set rsTemp = zlGetҽ�ƿ����
    
    rsTemp.Filter = "ID=" & mlngCardTypeID & " And �Ƿ�����=1"
    
    '74449,Ƚ����,2014-6-25,���ϴη�����𡱲����ڻ�ͣ��ʱ�޷���ȡ���������
    If rsTemp.EOF Then Exit Function
    
    With mCardType
        .str������ = Nvl(rsTemp!����)
        .lng���ų��� = Val(Nvl(rsTemp!���ų���))
        .lng���㷽ʽ = Trim(Nvl(rsTemp!���㷽ʽ))
        .bln���ƿ� = Val(Nvl(rsTemp!�Ƿ�����)) = 1
        .bln�ϸ���� = Val(Nvl(rsTemp!�Ƿ��ϸ����)) = 1
        .blnˢ�� = Val(Nvl(rsTemp!�Ƿ�ˢ��)) = 1
        .str�������� = Nvl(rsTemp!��������)
        .int���볤�� = Val(Nvl(rsTemp!���볤��))
        .int���볤������ = Val(Nvl(rsTemp!���볤������))
        .int������� = Val(Nvl(rsTemp!�������))
        .bln�Ƿ��ƿ� = Val(Nvl(rsTemp!�Ƿ��ƿ�)) = 1 '�����:56599
        .bln�Ƿ񷢿� = Val(Nvl(rsTemp!�Ƿ񷢿�)) = 1
        .bln�Ƿ�д�� = Val(Nvl(rsTemp!�Ƿ�д��)) = 1
        .bln�Ƿ�Ժ�ⷢ�� = (InStr(mstrPrivs, ";����;") > 0 And .bln���ƿ� = False And .bln�Ƿ񷢿� = True) '�����:56599
        .lng�������� = Val(Nvl(rsTemp!��������)) '�����:57326
        .lng�����ID = Val(Nvl(rsTemp!id)) '�����:57326
        .bln�Ƿ��ظ�ʹ�� = Val(Nvl(rsTemp!�Ƿ��ظ�ʹ��))
        .bln���￨ = .str������ = "���￨" And Val(Nvl(rsTemp!�Ƿ�̶�)) = 1 And Nvl(rsTemp!����) = "" '�����:57962
        .byt�������� = Val(Nvl(rsTemp!��������)) '104238
        .blnOneCard = False
        If Trim(Nvl(rsTemp!�ض���Ŀ)) <> "" Then
            Set .rsҽ�ƿ��� = zlGetSpecialItemFee(Trim(Nvl(rsTemp!�ض���Ŀ)))
            '����:48090
            '�����:56599
            If (.bln���ƿ� = True Or .bln�Ƿ�Ժ�ⷢ��) And .rsҽ�ƿ��� Is Nothing Then
                MsgBox .str������ & "δ���ö�Ӧ�Ŀ���,�뵽��ҽ�ƿ�������������!", vbInformation + vbOKOnly, gstrSysName
                mblnUnLoad = True
                mblnChange = False
                Exit Function
            End If
            If .bln���￨ Then .blnOneCard = GetOneCard.RecordCount > 0
        Else
            Set .rsҽ�ƿ��� = Nothing
        End If
        str���� = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModule, "0")
        '����ID,�����ID|...
        '120556,δ��չ���Ʊ�����Σ�����Ʊ������ID��ȡʧ�ܡ�
        .lng����ID = 0
        .lng�������� = 0
        varData = Split(str����, "|")
        For i = 0 To UBound(varData)
             varTemp = Split(varData(i), ",")
             If Val(varTemp(0)) <> 0 Then
                If ExistShareBill(Val(varTemp(0)), 5) Then
                    If Val(varTemp(1)) = mlngCardTypeID Then
                        .lng�������� = Val(varTemp(0)): Exit For
                    End If
                End If
             End If
        Next
        txtPass.MaxLength = .int���볤��
        txtAudi.MaxLength = .int���볤��
        txt����.PasswordChar = IIf(.str�������� <> "", "*", "")
        txtˢ������.PasswordChar = IIf(.str�������� <> "", "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End With
    InitCardType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitInsurePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2011-07-07 02:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, i As Long
    With mInsurePara
        .lng���ʽҽ������ = 0
        varTemp = Split(GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", ""), ",")
        For i = 0 To UBound(varTemp)
            If IsNumeric(varTemp(i)) Then
                If zlCheckMCOutMode(Val(varTemp(i))) Then .lng���ʽҽ������ = Val(varTemp(i)): Exit For
            End If
        Next
    End With
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2011-07-01 11:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mParaData
        .blnSeekName = zlDatabase.GetPara("����ģ������", glngSys, mlngModule) = "1"
        .intNameDays = Val(zlDatabase.GetPara("������������", glngSys, mlngModule))
        .blnShowExpend = Val(zlDatabase.GetPara("��ʾ��չ��Ϣ", glngSys, mlngModule))
        .int�˿�ģʽ = Val(zlDatabase.GetPara("�˿�ˢ��", glngSys, glngModul))
        '0-������ˢ��;1-ˢ���˿�;2-���ݺź�����֤ˢ��;3-1��2�Ĺ���ģʽ
        .bln���� = Val(zlDatabase.GetPara("���Ѽ���", glngSys, mlngModule, , Array(chk����), InStr(1, mstrPrivs, ";��������;"))) = "1"
    End With
    
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '���˵�ַ�ṹ��¼��
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '�����ַ�ṹ��¼��
    '94941:���ϴ�,2016/4/7,�Ƿ��Զ����������
    mbln�Զ������ = Val(zlDatabase.GetPara("�Զ������", glngSys, 1111)) = 1
End Sub
Private Sub SetDefaultLen()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ�ı༭����
    '����:���˺�
    '����:2011-06-28 02:50:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
        strSQL = " " & _
    "   Select A.����,A.�����,A.���֤��,A.����,A.��ͥ��ַ,A.��ͥ�绰,A.ҽ����,A.��ͥ��ַ, " & _
    "          A.����֤��,A.��ͥ��ַ�ʱ�,A.����,A.�����ص�,A.������λ,A.��λ�绰,A.���ڵ�ַ,A.���ڵ�ַ�ʱ�," & _
    "          A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.��ϵ������,A.��ϵ�˵�ַ,A.��ϵ�˵绰,B.����,B.����,A.�ֻ���" & _
    "   From ������Ϣ A,����ҽ�ƿ���Ϣ B" & _
    "   Where a.����id = 0 and a.����ID=b.����ID and B.�����ID=0 " & _
    " "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    txtPatient.MaxLength = rsTemp.Fields("����").DefinedSize
    txt���֤��.MaxLength = rsTemp.Fields("���֤��").DefinedSize
    txt�����.MaxLength = rsTemp.Fields("�����").DefinedSize - 1
    txt����.MaxLength = rsTemp.Fields("����").DefinedSize
    txt��ͥ��ַ.MaxLength = rsTemp.Fields("��ͥ��ַ").DefinedSize
    padd��ͥ��ַ.MaxLength = rsTemp.Fields("��ͥ��ַ").DefinedSize
    txt���ڵ�ַ.MaxLength = rsTemp.Fields("���ڵ�ַ").DefinedSize
    padd���ڵ�ַ.MaxLength = rsTemp.Fields("���ڵ�ַ").DefinedSize
    txt���ڵ�ַ�ʱ�.MaxLength = rsTemp.Fields("���ڵ�ַ�ʱ�").DefinedSize
    txt��ͥ�绰.MaxLength = rsTemp.Fields("��ͥ�绰").DefinedSize
    txtҽ����.MaxLength = rsTemp.Fields("ҽ����").DefinedSize
    txt��ͥ�ʱ�.MaxLength = rsTemp.Fields("��ͥ��ַ�ʱ�").DefinedSize
    txt����֤��.MaxLength = rsTemp.Fields("����֤��").DefinedSize
    txt����.MaxLength = rsTemp.Fields("����").DefinedSize
    txt�����ص�.MaxLength = rsTemp.Fields("�����ص�").DefinedSize
    txt������λ.MaxLength = rsTemp.Fields("������λ").DefinedSize
    txt��λ�绰.MaxLength = rsTemp.Fields("��λ�绰").DefinedSize
    txt��λ�ʱ�.MaxLength = rsTemp.Fields("��λ�ʱ�").DefinedSize
    txt��λ������.MaxLength = rsTemp.Fields("��λ������").DefinedSize
    txt��λ�ʻ�.MaxLength = rsTemp.Fields("��λ�ʺ�").DefinedSize
    txt��ϵ������.MaxLength = rsTemp.Fields("��ϵ������").DefinedSize
    txt��ϵ�˵�ַ.MaxLength = rsTemp.Fields("��ϵ�˵�ַ").DefinedSize
    txt��ϵ�˵绰.MaxLength = rsTemp.Fields("��ϵ�˵绰").DefinedSize
    txtPass.MaxLength = rsTemp.Fields("����").DefinedSize
    txtAudi.MaxLength = rsTemp.Fields("����").DefinedSize
    txt�ֻ�.MaxLength = rsTemp.Fields("�ֻ���").DefinedSize
    If mCardType.lng���ų��� = 0 Then mCardType.lng���ų��� = rsTemp.Fields("����").DefinedSize
    txt����.MaxLength = mCardType.lng���ų���
    If mInsurePara.lng���ʽҽ������ = 920 Then '�������
        txtҽ����.MaxLength = 12
    Else
        txtҽ����.MaxLength = 30
    End If
    txtҽ����.ToolTipText = "��󳤶�" & txtҽ����.MaxLength & "λ"
    txt��֤ҽ����.MaxLength = txtҽ����.MaxLength
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-01 10:35:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control, lngLen As Long, strMCAccount As String, lngTmp As Long
    Dim strTemp As String, i As Long
    Dim blnNewPati As Boolean, str����ʱ�� As String
    blnNewPati = False
    If mrsInfo Is Nothing Then
        blnNewPati = True
    ElseIf mrsInfo.State <> 1 Then
        blnNewPati = True
    End If
  
    For Each objCtl In Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")  '�ı�
            lngLen = objCtl.MaxLength
            If lngLen <> 0 Then
                If zlCommFun.ActualLen(objCtl.Text) > lngLen Then
                    MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "���ֻ������" & lngLen \ 2 & "�����ֻ�" & lngLen & "���ַ�,����", vbInformation + vbOKOnly, gstrSysName
                    If InStr(1, ",����,�����,���֤��,��סַ,���ڵ�ַ,��ͥ�绰,�ֻ���,����,����,��֤,", "," & objCtl.Tag & ",") > 0 Then
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    Else
                        If wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = False Then
                            wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = True
                        End If
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
                End If
            End If
            If Trim(objCtl.Text) = "" And InStr(1, ",����,�����,����,", "," & objCtl.Tag & ",") > 0 Then
                '������Ŀ
                If Not (mEditType = Cr_����������Ϣ And objCtl.Tag = "����") Then
                    If Not (blnNewPati = False And objCtl.Tag = "�����") Then
                        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "��������,����", vbInformation + vbOKOnly, gstrSysName
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Case UCase("Patiaddress")
            If mblnStructAdress And objCtl.Enabled Then
                lngLen = objCtl.MaxLength
                If lngLen <> 0 Then
                    If zlCommFun.ActualLen(objCtl.value) > lngLen Then
                        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "���ֻ������" & lngLen \ 2 & "������,���顣", vbInformation + vbOKOnly, gstrSysName
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
                End If
                If objCtl.CheckNullValue(, True, False) <> "" Then
                    MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "��" & objCtl.CheckNullValue(, True, False) & "��δ����,���顣", vbInformation + vbOKOnly, gstrSysName
                    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                    Exit Function
                End If
            End If
        Case Else
        End Select
    Next
    If Not blnNewPati Then isValied = True: Exit Function
    
        If Not IsNumeric(txt�����.Text) And txt�����.Text <> "" Then
            MsgBox "������һ����Ч������ţ�", vbInformation, gstrSysName
            If txt�����.Enabled And txt�����.Visible Then txt�����.SetFocus
            Exit Function
        End If
        If IsNumeric(txt�����.Text) Then
            If ExistClinicNO(txt�����.Text) Then
                If mbln�Զ������ Then
                    If txt�����.Text <> lbl�����.Tag Then
                        MsgBox "���ָò��˵Ĳ��������[" & txt�����.Text & "]�Ѿ�����������ʹ��,ϵͳ���Զ�����һ�����ظ��ĺ��룡", vbInformation, gstrSysName
                        txt�����.Text = zlGet�����: lbl�����.Tag = txt�����.Text
                        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus: Exit Function
                    Else
                        '�Զ������ĺ����û���޸ģ���ֱ���ٴ��Զ���������
                        txt�����.Text = zlGet�����: lbl�����.Tag = txt�����.Text
                    End If
                Else
                    MsgBox "���ָò��˵Ĳ��������[" & txt�����.Text & "]�Ѿ�����������ʹ��,�����һ�����ظ��ĺ��룡", vbInformation, gstrSysName
                    txt�����.Text = "": lbl�����.Tag = ""
                    If txt�����.Enabled And txt�����.Visible Then txt�����.SetFocus: Exit Function
                End If
            End If
        End If
        If txtҽ����.Text <> "" Or txt��֤ҽ����.Text <> "" Then
            If txtҽ����.Text <> txt��֤ҽ����.Text And txt��֤ҽ����.Visible Then
                MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
                If txtҽ����.Visible And txtҽ����.Enabled Then txtҽ����.SetFocus
                Exit Function
            End If
            If zlCommFun.ActualLen(txtҽ����.Text) > txtҽ����.MaxLength Then
                MsgBox "����,ҽ������󳤶Ȳ��ܳ���" & txtҽ����.MaxLength & "���ַ���", vbInformation, gstrSysName
                If txtҽ����.Visible And txtҽ����.Enabled Then txtҽ����.SetFocus
                Exit Function
            End If
        End If
        
    
    strMCAccount = Trim(txtҽ����.Text)
    If mInsurePara.lng���ʽҽ������ = 920 And strMCAccount <> lblҽ����(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
        If CheckExistsMCNO(strMCAccount) Then
            If txtҽ����.Visible And txtҽ����.Enabled Then txtҽ����.SetFocus
            Exit Function
        End If
    End If
    If Not IsDate(txt��������.Text) Then
        MsgBox "������ȷ����������ڣ�", vbInformation, gstrSysName
        If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
        Exit Function
    End If
    If Trim(txt����.Text) = "" Then
        MsgBox "��������[����]��", vbExclamation, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Function
    End If
    '69026,Ƚ����,2014-8-11,������Ч�Լ��
    '76703,Ƚ����,2014-8-15
    If txt����.Enabled And txt����.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), _
                IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
                IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
        End If
    End If
    If IsDate(txt��������.Text) Then
        '76669�����ϴ�,2014-8-15,������������ڼ��
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        If CDate(str����ʱ��) > zlDatabase.Currentdate Then
            If MsgBox("����ʱ�䣺" & str����ʱ�� & " �����˵�ǰϵͳʱ�䡣" & _
                vbCrLf & vbCrLf & "���������������ڵ���ȷ�� ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
                Exit Function
            End If
        End If
    End If
    If cbo�ѱ�.ListIndex = -1 Then
        MsgBox "����ȷ��[�ѱ�]��", vbExclamation, gstrSysName
        If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus
        Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ��[����]��", vbExclamation, gstrSysName
        If cbo����.Enabled And cbo����.Visible Then cbo����.SetFocus
        Exit Function
    End If
    
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ��[����]��", vbExclamation, gstrSysName
        If cbo����.Enabled And cbo����.Visible Then cbo����.SetFocus
        Exit Function
    End If
    
    '������ƵĲ���,�����ظ�
    If mrsInfo Is Nothing Then
        strTemp = SimilarIDs(NeedName(cbo����.Text), NeedName(cbo����), CDate(IIf(IsDate(txt��������.Text), txt��������.Text, #1/1/1900#)), NeedName(cbo�Ա�), txtPatient.Text, txt���֤��.Text)
        If strTemp <> "" Then
            i = UBound(Split(strTemp, "|")) + 1
            strTemp = Replace(strTemp, "|", vbCrLf)
            If i > 20 Then strTemp = Mid(strTemp, 1, 200) & "..."
            
            If MsgBox("�����еĲ�����Ϣ�з��� " & i & " ����Ϣ���ƵĲ���(����,����,�Ա�,����,����������ͬ�����֤����ͬ): " & vbCrLf & vbCrLf & _
                strTemp & vbCrLf & vbCrLf & "ȷʵҪ����ò��˵���Ϣ��", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            Else
                MsgBox "�ò��˵����Ƽ�¼����ʹ��""�ϲ�""���ܴ���", vbInformation, gstrSysName
            End If
        End If
    End If
    isValied = True
End Function
Public Function zlShowCard(ByVal frmMain As Object, lngModule As Long, strPrivs As String, _
     EditType As gCardType, lngCardTypeID As Long, _
     Optional strCardNo As String = "", _
     Optional lng����ID As Long, _
     Optional strBillNo As String, _
     Optional int��¼״̬ As Integer, _
     Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ƭ
    '���:frmMain-���õ�������
    '       EditType:=�༭����
    '       lngCardTypeID-�����Id
    '       strCardNo-����
    '����:
    '����:�ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2011-06-28 22:29:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditTypeOld = EditType
    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs
    mlngCardTypeID = lngCardTypeID: mstrCardNo = strCardNo: mintSucces = 0
    mstrBillNo = strBillNo: mint��¼״̬ = int��¼״̬: mblnNOMoved = blnNOMoved
    mlng����ID = lng����ID: mblnUnLoad = False
    mblnFromCardMgr = False
    If frmMain.Caption = "ҽ�ƿ����Ź���" Then mblnFromCardMgr = True
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
End Function
Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��InitTaskPancel
    '����:���˺�
    '����:2011-06-30 18:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    Call wndTaskPanel.SetGroupInnerMargins(2, 0, 2, 0)
    Call wndTaskPanel.SetGroupOuterMargins(2, -10, 2, -21)
    Call wndTaskPanel.SetMargins(2, 16, 2, 10, 30)
     
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    
    If Not (mEditType = Cr_���� Or mEditType = Cr_���� Or mEditType = Cr_��ʧ Or mEditType = Cr_����������Ϣ) Then
        Set tkpGroup = wndTaskPanel.Groups.Add(idx_TP_Tittle, "")
        Set Item = tkpGroup.Items.Add(idx_TP_Tittle, "", xtpTaskItemTypeControl)
        Set Item.Control = picTittle
        fraSplit.BackColor = Item.BackColor
        picTittle.BackColor = Item.BackColor
        tkpGroup.Expandable = False
        Call Item.SetMargins(0, -19, 0, 0)
    End If

    Set tkpGroup = wndTaskPanel.Groups.Add(Idx_TP_PatiBase, "���˻�����Ϣ")
    Set Item = tkpGroup.Items.Add(Idx_TP_PatiBase, "", xtpTaskItemTypeControl)
    Set Item.Control = picBasePati
    fra.BackColor = Item.BackColor
    picBasePati.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    
    Set tkpGroup = wndTaskPanel.Groups.Add(Idx_TP_PatiExpend, "���ಡ����Ϣ")
    tkpGroup.Tooltip = "��CTRL+E ��ʾ����Ĳ�����Ϣ"
    Set Item = tkpGroup.Items.Add(Idx_TP_PatiExpend, "", xtpTaskItemTypeControl)
    Set Item.Control = picExpend
    picExpend.BackColor = Item.BackColor
    fraBase.BackColor = picExpend.BackColor
    If mEditType = Cr_����������Ϣ Then
        tkpGroup.Expandable = False
        tkpGroup.Expanded = True
    Else
        tkpGroup.Expanded = mParaData.blnShowExpend
    End If
    
    If mEditType <> Cr_����������Ϣ Then
        Set tkpGroup = wndTaskPanel.Groups.Add(Idx_TP_PatiCard, IIf(mCardType.str������ = "", "ҽ�ƿ�", mCardType.str������))
        tkpGroup.Expandable = False
        tkpGroup.Expanded = True
        picCard.Height = 2205: If mEditType <> Cr_�󶨿� And mEditType <> Cr_���� Then picCard.Height = 1705
        Set Item = tkpGroup.Items.Add(Idx_TP_PatiCard, "", xtpTaskItemTypeControl)
        Set Item.Control = picCard
        picCard.BackColor = Item.BackColor
        fraCard.BackColor = Item.BackColor
        tkpGroup.Expanded = True
    End If
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub
Private Sub cboNO_KeyPress(KeyAscii As Integer)
     Dim strOper As String, vDate As Date
    If cboNO.Locked Then Exit Sub
    
    'ת���ɴ�д(���ֲ��ɴ���)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        'Call SetNOInputLimit(cboNO, KeyAscii)
        Exit Sub
    End If
    If Not (cboNO.Text <> "" And Not cboNO.Locked) Then Exit Sub
    
    cboNO.Text = GetFullNO(cboNO.Text, 16)
    '�Ƿ���ת������ݱ���
    If zlDatabase.NOMoved("סԺ���ü�¼", cboNO.Text, , "5") Then
        If Not ReturnMovedExes(cboNO.Text, 5, Me.Caption) Then Exit Sub
        mblnNOMoved = False
    End If
    '����Ȩ��
    If Not ReadBillInfo(2, cboNO.Text, 5, strOper, vDate) Then
        txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
    End If
    If Not BillOperCheck(8, strOper, vDate, "�˿�") Then
        txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
    End If
    '��ȡҪ�˿��ļ�¼(��NO)
    Select Case ReadBill(cboNO.Text)
        Case -1
            If InStr(1, "13", mParaData.int�˿�ģʽ) > 0 Then
                If txtˢ������.Visible And txtˢ������.Enabled Then txtˢ������.SetFocus
            Else
               If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
            End If
        Case 0
            If glngSys Like "8??" Then
                MsgBox "��ȡ�û�Ա�����ż�¼ʧ�ܣ�", vbExclamation, gstrSysName
            Else
                MsgBox "��ȡ��ҽ�ƿ����ż�¼ʧ�ܣ�", vbExclamation, gstrSysName
            End If
            txtPatient.Text = "": cboNO.SetFocus
        Case 1
            If glngSys Like "8??" Then
                MsgBox "�û�Ա�����ż�¼�����ڣ�", vbExclamation, gstrSysName
            Else
                MsgBox "��ҽ�ƿ����ż�¼�����ڣ�", vbExclamation, gstrSysName
            End If
            txtPatient.Text = "": cboNO.SetFocus
        Case 2
            If glngSys Like "8??" Then
                MsgBox "�û�Ա�����ż�¼�Ѿ��˳���", vbExclamation, gstrSysName
            Else
                MsgBox "��ҽ�ƿ����ż�¼�Ѿ��˳���", vbExclamation, gstrSysName
            End If
            txtPatient.Text = ""
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    End Select
End Sub

Private Sub cbo�ѱ�_Change()
    mblnChange = True
End Sub

Private Sub cbo�ѱ�_Click()
    mblnChange = True
    If mblnNotChange Then Exit Sub
    Call LoadCardFee
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    SearchCombox cbo�ѱ�, KeyAscii
End Sub

Private Sub cbo����_Change()
    mblnChange = True
End Sub

Private Sub cbo����_Click()
    mblnChange = True
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    SearchCombox cbo����, KeyAscii
End Sub

Private Sub cbo����״��_Change()
    mblnChange = True
End Sub

Private Sub cbo����״��_Click()
    mblnChange = True
End Sub

Private Sub cbo����״��_KeyPress(KeyAscii As Integer)
    SearchCombox cbo����״��, KeyAscii
End Sub

Private Sub cbo��ϵ�˹�ϵ_Change()
        mblnChange = True
End Sub

Private Sub cbo��ϵ�˹�ϵ_Click()
    mblnChange = True
End Sub

Private Sub cbo��ϵ�˹�ϵ_KeyPress(KeyAscii As Integer)

    SearchCombox cbo��ϵ�˹�ϵ, KeyAscii
End Sub

Private Sub cbo����_Change()
    mblnChange = True
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    SearchCombox cbo����, KeyAscii
End Sub

Private Sub cbo���䵥λ_Click()
    mblnChange = True
End Sub

Private Sub cbo���䵥λ_LostFocus()
    Dim strBirth As String
    '69026,Ƚ����,2014-8-8,�����������
    '76703,Ƚ����,2014-8-15
    '111836:���ϴ���2017/7/21�����䷴��
    If mobjPubPatient Is Nothing Then Exit Sub
    If cbo���䵥λ.Text <> mstr���䵥λ Then
        mblnNotChange = True
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt����.Text) & cbo���䵥λ.Text, strBirth) Then
                txt��������.Text = Format(strBirth, "yyyy-mm-dd")
                txt����ʱ��.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnNotChange = False
        mstr���䵥λ = cbo���䵥λ.Text
    End If
    If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & cbo���䵥λ.Text, _
            IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
            IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
        If txt����.Visible And txt����.Enabled Then txt����.SetFocus: Exit Sub
    End If
End Sub

Private Sub cbo���_Change()
    mblnChange = True
End Sub

Private Sub cbo���_Click()
    mblnChange = True
End Sub

Private Sub cbo���_KeyPress(KeyAscii As Integer)
    SearchCombox cbo���, KeyAscii
End Sub

Private Sub cboѧ��_Change()
    mblnChange = True
End Sub

Private Sub cboѧ��_Click()
    mblnChange = True
End Sub

Private Sub cboѧ��_KeyPress(KeyAscii As Integer)
  SearchCombox cboѧ��, KeyAscii
End Sub

Private Sub cboҽ�Ƹ���_Change()
    mblnChange = True
End Sub

Private Sub cboҽ�Ƹ���_KeyPress(KeyAscii As Integer)
     SearchCombox cboҽ�Ƹ���, KeyAscii
End Sub

Private Sub cbo֧����ʽ_Change()
    mblnChange = True
End Sub

Private Sub cbo֧����ʽ_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    
    mblnChange = True
    If mblnNotClick = True Then Exit Sub
    
    With mCurPayMoney
            .lngҽ�ƿ����ID = 0
            .bln���ѿ� = False
            .str���㷽ʽ = ""
            .str���� = ""
            .strˢ������ = ""
            .strˢ������ = ""
            .strNo = ""
            .lngID = 0
            .lng����ID = 0
     End With
     
    If Not (mEditType = Cr_���� Or mEditType = Cr_����) Then Exit Sub
    With cbo֧����ʽ
        If .ListIndex = -1 Then Exit Sub
        lngIndex = .ListIndex + 1
    End With
    '��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not mcolPayMode Is Nothing Then
        With mCurPayMoney
                .lngҽ�ƿ����ID = Val(mcolPayMode(lngIndex)(3))
                .bln���ѿ� = Val(mcolPayMode(lngIndex)(5)) = 1
                .str���㷽ʽ = Trim(mcolPayMode(lngIndex)(6))
                .str���� = Trim(mcolPayMode(lngIndex)(1))
         End With
    Else
        '86578:���ϴ�,2015/7/16,ˢ�����㷽ʽ
        mCurPayMoney.str���㷽ʽ = NeedName(cbo֧����ʽ.Text)
    End If
    Call txt���_Change
End Sub

Private Sub cbo֧����ʽ_KeyPress(KeyAscii As Integer)
     SearchCombox cbo֧����ʽ, KeyAscii
End Sub
Private Sub cboְҵ_Change()
    mblnChange = True
End Sub

Private Sub cboְҵ_Click()
    mblnChange = True
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    SearchCombox cboְҵ, KeyAscii
End Sub

Private Sub chkCancel_Click()
    stbThis.Panels(2).Text = ""
    If mEditType <> Cr_���� And mEditType <> Cr_�˿� Then Exit Sub
    If mblnNotClick Then Exit Sub
    mblnNotClick = True
    If mEditType = Cr_�˿� Then chkCancel.value = 1
    mblnNotClick = False
    Load֧����ʽ (chkCancel.value = 1)
    lblԤ�����.Caption = "Ԥ�����:0Ԫ"
    If mEditType <> Cr_�˿� Then picԤ�����.Top = wndTaskPanel.Height - picCard.Height - picԤ�����.Height - 180
    Call SetControlEnable: Call SetControlVisitble
    chkCancel.ForeColor = IIf(chkCancel.value = 1, &HFF, 0) '��Ϊ��ɫ
    Call ClearData
    If chkCancel.value = Checked Then
        '�������˿�ĵ��ݺ�
        cboNO.Text = "": cboNO.Tag = "": cboNO.Locked = False
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
        If txtˢ������.Visible And txtˢ������.Enabled Then txtˢ������.SetFocus
    Else
        Call LoadCardFee
        txtPatient.Text = ""
        txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm")
        cboNO.Locked = True
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    If chkCancel.value = 1 Then
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    End If
End Sub

Private Sub chk����_Click()
    mblnChange = True
    cbo֧����ʽ.Enabled = Not chk����.value = Checked
    cbo֧����ʽ.BackColor = IIf(cbo֧����ʽ.Enabled, &H80000005, &H8000000F)
    If mEditType = Cr_�˿� Or chkCancel.value = Checked Then Exit Sub
    txt�ϼ�.Enabled = Not chk����.value = Checked
    txt�ϼ�.BackColor = IIf(txt�ϼ�.Enabled, &H80000005, &H8000000F)
    txt���.Enabled = Not chk����.value = Checked
    txt���.BackColor = IIf(txt���.Enabled, &H80000005, &H8000000F)
    IDKindPayMode.Enabled = Not chk����.value = Checked
    'txt���.BackColor = IIf(IDKindPayMode.Enabled, &H80000005, &H8000000F)
    
    If chk����.value = Checked Then
        txt�ϼ�.Text = "": txt�ϼ�.Tag = ""
        txt���.Text = ""
    Else
        Call txt���_Change
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-07 03:47:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, blnNewPati As Boolean, Curdate As Date, lng����ID As Long
    Dim cllPro As Collection, cllUpdateSwap As Collection, cllThree As Collection
    Dim strErrMsg As String
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        blnNewPati = True
    ElseIf mrsInfo.State <> 1 Then
        blnNewPati = True
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    '�����:57326
    If Check��������(lng����ID, mCardType.lng�����ID) = False Then Exit Function
    
    Set cllPro = New Collection
    Set cllUpdateSwap = New Collection: Set cllThree = New Collection
    Curdate = zlDatabase.Currentdate
    If blnNewPati Then
        lng����ID = zlDatabase.GetNextNo(1)
        Call AddNewPatiSQL(lng����ID, Curdate, cllPro)
    End If
    Call AddCardDataSQL(lng����ID, Curdate, cllPro, lng����ID)
    If IDKindPayMode.IDKind = 2 And Val(txt���.Text) > 0 Then Call AddDepositSQL(lng����ID, Curdate, cllPro, lng����ID)
    '�����:56599
    If lng����ID > 0 Then Call Add�����������Ϣ(lng����ID, cllPro)
    
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    '101170:���ϴ�,2017/5/3,����HIS����Ҫ�ύEMPI���ݣ�ʧ�ܺ��������ݶ�Ҫ����
    If zlSaveEMPIPatiInfo(blnNewPati, lng����ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "��EMPIƽ̨�ϴ�������Ϣʧ�ܣ�"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    If zlInterfacePrayMoney(cllUpdateSwap, cllThree) = False Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    zlExecuteProcedureArrAy cllUpdateSwap, Me.Caption, False, True
    '�����:53408
    '�����:54172
    '�����:54208
    If mCardType.str������ = "�������֤" And Not mrsInfo Is Nothing Then
        If mrsInfo.State <> 0 Then
            SaveModifyPati '�޸Ĳ�����Ϣ����Ҫ��Ϊ�˸��������֤��
        End If
    End If
    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllThree, Me.Caption
    '��Ƭ
    Select Case mlngͼ�����
        Case 1 '�ļ�
            SavePatPicture lng����ID, cmdialog.FileName
        Case 2 '�ɼ�
            SavePatPicture lng����ID, mstr�ɼ�ͼƬ
            mstr�ɼ�ͼƬ = ""
        Case 4 '�������֤
            mstrIDImageFile = App.Path & "\SFZIMG.bmp"
            SavePicture imgPatient.Picture, mstrIDImageFile
            SavePatPicture lng����ID, mstrIDImageFile
        Case 3 '����
            DeletePatPicture lng����ID
    End Select
    '�����:56599
    'Ժ�ⷢ���������Ҫ����д������
    If mCardType.bln�Ƿ�д�� Then Call WriteCard(lng����ID)
        
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '������������Ϣ
        On Error Resume Next
        Call mobjPlugIn.PatiInfoSaveAfter(mlng����ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
        Err.Clear: On Error GoTo 0
    End If

    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    Call ErrCenter
   
    Exit Function
ErrOthers:
    If ErrCenter = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    gcnOracle.CommitTrans
End Function
 
Private Sub AddCardDataSQL(ByVal lng����ID As Long, ByVal dtCurdate As Date, _
    ByRef cllPro As Collection, ByRef lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���￨���Ŵ���
    '���:lng����ID
    '����:���˺�
    '����:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt�������� As Byte, strNo As String, strPassWord As String, strSQL As String
    Dim strԭ���� As String, str���� As String, strCard As String, str�䶯ԭ�� As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str���㷽ʽ As String, strBrushCardNo As String
    Dim bln���ѿ� As Boolean, blnInRange As Boolean   '��Χ�ڵĿ�
    Dim lngIndex As Long
    Dim byt�䶯���� As Byte
     blnInRange = True
     lng����ID = 0
    
    If mCardType.blnOneCard And mCardType.bln�ϸ���� Then blnInRange = mCardType.lng����ID > 0
    Select Case mEditType
    Case Cr_�󶨿�
         blnInRange = False: byt�������� = 0
         byt�䶯���� = 11
    Case Cr_����
         byt�������� = 0: byt�䶯���� = 1
         If mCardType.rsҽ�ƿ��� Is Nothing Then
             blnInRange = False
         End If
    Case Cr_����
         byt�������� = 1: byt�䶯���� = 3
    Case Cr_����
        byt�������� = 2: blnInRange = False: byt�䶯���� = 2
        '���ԭ��,�Ǵ��ڿ��ѵ�,�ڻ���ʱ,��Ҫ���õ��ù��̴�����Ӧ��,Ʊ����ϸ
        If Not mCardType.rsҽ�ƿ��� Is Nothing Then
            blnInRange = True
        End If
    Case Else
        Exit Sub
    End Select
    strCard = Trim(txt����.Text): strICCard = IIf(mblnICCard, strCard, "")
    strԭ���� = Trim(txtˢ������.Text)
    lblNo.Tag = ""
    strPassWord = zlCommFun.zlStringEncode(Trim(txtPass.Text))
    If blnInRange = False Then
          'Zl_ҽ�ƿ��䶯_Insert
           strSQL = "Zl_ҽ�ƿ��䶯_Insert("
          '      �䶯����_In   Number,
          '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
          strSQL = strSQL & "" & byt�䶯���� & ","
          '      ����id_In     סԺ���ü�¼.����id%Type,
          strSQL = strSQL & "" & lng����ID & ","
          '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
          strSQL = strSQL & "" & mlngCardTypeID & ","
          '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strԭ���� & "',"
          '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
          '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
          strSQL = strSQL & "'" & str�䶯ԭ�� & "',"
          '      ����_In       ������Ϣ.����֤��%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
          strSQL = strSQL & "'" & UserInfo.���� & "',"
          '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic����_In     ������Ϣ.Ic����%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
          strSQL = strSQL & "NULL)"
    Else
        strNo = zlDatabase.GetNextNo(16)  'ҽ�ƿ�
        lblNo.Tag = strNo
        If chk����.value = 0 Then
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        End If
        mCurPayMoney.lng����ID = lng����ID
        mCurPayMoney.strNo = strNo
        If cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) < 0 Then
            lngIndex = cbo֧����ʽ.ListIndex + 1
            lngBrushCardTypeID = mcolPayMode(lngIndex)(3)
            '��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
            lngBrushCardTypeID = Val(mcolPayMode(lngIndex)(3))
            bln���ѿ� = Val(mcolPayMode(lngIndex)(5)) = 1
        Else
            bln���ѿ� = False
        End If
        
        '103980:���ϴ�,2017/1/19,���淢����������
        str���� = Trim(txt����.Text)
        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    
        '86578:���ϴ�,2015/7/16,ˢ�����㷽ʽ
        str���㷽ʽ = mcolPayMode(cbo֧����ʽ.ListIndex + 1)(6)
        If str���㷽ʽ = "" Then str���㷽ʽ = NeedName(cbo֧����ʽ.Text)
        If Not cbo֧����ʽ.Enabled Then str���㷽ʽ = ""
        strSQL = zlGetSaveCardFeeSQL(mlngCardTypeID, byt��������, strNo, lng����ID, 0, UserInfo.����ID, UserInfo.����ID, 0, _
        NeedName(cbo�ѱ�.Text), strԭ����, Trim(txtPatient.Text), NeedName(cbo�Ա�.Text), str����, _
        strCard, strPassWord, str�䶯ԭ��, IIf(mCardType.bln��� = False, mCardType.dblӦ�ս��, Val(txt����.Text)), Val(txt����.Text), str���㷽ʽ, _
        dtCurdate, mCardType.lng����ID, mCardType.rsҽ�ƿ���, strICCard, mCurPayMoney.lngҽ�ƿ����ID, mCurPayMoney.bln���ѿ�, mCurPayMoney.strˢ������, mCurPayMoney.lng����ID)
    End If
    zlAddArray cllPro, strSQL
 End Sub
Private Function AddNewPatiSQL(ByVal lng����ID As Long, ByVal dtCurdate As Date, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����²�������
    '����:cllPro-���̼�
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-07 04:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str���� As String, str�������� As String
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
       
    '  Zl_������Ϣ_Insert
    strSQL = "Zl_������Ϣ_Insert("
    '  ����id_In     ������Ϣ.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �����_In     ������Ϣ.�����%Type,
    strSQL = strSQL & "" & IIf(Trim(txt�����.Text) <> "", Val(txt�����.Text), "NULL") & ","
    '  �ѱ�_In       ������Ϣ.�ѱ�%Type,
    strSQL = strSQL & "'" & NeedName(cbo�ѱ�.Text) & "',"
    '  ҽ�Ƹ���_In   ������Ϣ.ҽ�Ƹ��ʽ%Type,
    strSQL = strSQL & "'" & NeedName(cboҽ�Ƹ���.Text) & "',"
    '  ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & txtPatient.Text & "',"
    '  �Ա�_In       ������Ϣ.�Ա�%Type,
    strSQL = strSQL & "'" & NeedName(cbo�Ա�.Text) & "',"
    '  ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ��������_In   ������Ϣ.��������%Type,
    strSQL = strSQL & "" & str�������� & ","
    '  �����ص�_In   ������Ϣ.�����ص�%Type,
    strSQL = strSQL & "'" & txt�����ص�.Text & "',"
    '  ���֤��_In   ������Ϣ.���֤��%Type,
    strSQL = strSQL & "'" & txt���֤��.Text & "',"
    '  ���_In       ������Ϣ.���%Type,
    strSQL = strSQL & "'" & NeedName(cbo���.Text) & "',"
    '  ְҵ_In       ������Ϣ.ְҵ%Type,
    strSQL = strSQL & "'" & NeedName(cboְҵ.Text, mstrCboSplit) & "',"
    '  ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & NeedName(cbo����.Text) & "',"
    '  ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & NeedName(cbo����.Text) & "',"
    '  ѧ��_In       ������Ϣ.ѧ��%Type,
    strSQL = strSQL & "'" & NeedName(cboѧ��.Text) & "',"
    '  ����_In       ������Ϣ.����״��%Type,
    strSQL = strSQL & "'" & NeedName(cbo����״��.Text) & "',"
    '  ��ͥ��ַ_In   ������Ϣ.��ͥ��ַ%Type,
    strSQL = strSQL & "'" & IIf(mblnStructAdress, padd��ͥ��ַ.value, txt��ͥ��ַ.Text) & "',"
    '  ��ͥ�绰_In   ������Ϣ.��ͥ�绰%Type,
    strSQL = strSQL & "'" & txt��ͥ�绰.Text & "',"
    '  ��ͥ��ַ�ʱ�_In   ������Ϣ.��ͥ��ַ�ʱ�%Type,
    strSQL = strSQL & "'" & txt��ͥ�ʱ�.Text & "',"
    '  ��ϵ������_In ������Ϣ.��ϵ������%Type,
    strSQL = strSQL & "'" & txt��ϵ������.Text & "',"
    '  ��ϵ�˹�ϵ_In ������Ϣ.��ϵ�˹�ϵ%Type,
    strSQL = strSQL & "'" & NeedName(cbo��ϵ�˹�ϵ.Text) & "',"
    '  ��ϵ�˵�ַ_In ������Ϣ.��ϵ�˵�ַ%Type,
    strSQL = strSQL & "'" & txt��ϵ�˵�ַ.Text & "',"
    '  ��ϵ�˵绰_In ������Ϣ.��ϵ�˵绰%Type,
    strSQL = strSQL & "'" & txt��ϵ�˵绰.Text & "',"
    '  ��ͬ��λid_In ������Ϣ.��ͬ��λid%Type,
    strSQL = strSQL & "" & IIf(Val(lbl������λ.Tag) = 0, "NULL", Val(lbl������λ.Tag)) & ","
    '  ������λ_In   ������Ϣ.������λ%Type,
    strSQL = strSQL & "'" & txt������λ.Text & "',"
    '  ��λ�绰_In   ������Ϣ.��λ�绰%Type,
    strSQL = strSQL & "'" & txt��λ�绰.Text & "',"
    '  ��λ�ʱ�_In   ������Ϣ.��λ�ʱ�%Type,
    strSQL = strSQL & "'" & txt��λ�ʱ�.Text & "',"
    '  ��λ������_In ������Ϣ.��λ������%Type,
    strSQL = strSQL & "'" & txt��λ������.Text & "',"
    '  ��λ�ʺ�_In   ������Ϣ.��λ�ʺ�%Type,
    strSQL = strSQL & "'" & txt��λ�ʻ�.Text & "',"
    '  ������_In     ������Ϣ.������%Type,
    strSQL = strSQL & "null,"
    '  ������_In     ������Ϣ.������%Type,
    strSQL = strSQL & "null,"
    '  ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "null,"
    '  �Ǽ�ʱ��_In   ������Ϣ.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(dtCurdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  ����_In       ������Ϣ.����%Type := Null,
    strSQL = strSQL & "'" & NeedName(txt����.Text) & "',"
    '  ��������_In   ������Ϣ.��������%Type := Null,
    strSQL = strSQL & "null,"
    '  ����Ա���_In ���˵�����¼.����Ա���%Type := Null,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ���˵�����¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ҽ����_In     ������Ϣ.ҽ����%Type := Null,
    strSQL = strSQL & "" & IIf(Trim(txtҽ����.Text) = "", "NULL", "'" & Trim(txtҽ����.Text) & "'") & ","
    '  ����֤��_In   ������Ϣ.����֤��%Type := Null
    strSQL = strSQL & "'" & txt����֤��.Text & "',"
    '�����:51071
    '  ����_In   ������Ϣ.����%Type := Null
    strSQL = strSQL & "'',"
    '  ���ڵ�ַ_In   ������Ϣ.���ڵ�ַ%Type := Null
    strSQL = strSQL & "'" & IIf(mblnStructAdress, Trim(padd���ڵ�ַ.value), Trim(txt���ڵ�ַ.Text)) & "',"
    '  ���ڵ�ַ�ʱ�_In   ������Ϣ.���ڵ�ַ�ʱ�%Type := Null
    strSQL = strSQL & "'" & Trim(txt���ڵ�ַ�ʱ�.Text) & "',"
    '  ��ϵ�����֤��_In   ������Ϣ.��ϵ�����֤��%Type := Null
    strSQL = strSQL & "'" & Trim(txt��ϵ�����֤��.Text) & "',"
    '  ��������_In   ������Ϣ.��������%Type := Null
    strSQL = strSQL & "'',"
    '  �ֻ���_In     ������Ϣ.�ֻ���%Type := Null
    strSQL = strSQL & "'" & txt�ֻ�.Text & "')"
    zlAddArray cllPro, strSQL
    '89242:���ϴ�,2015/12/3,���²��˵�ַ��Ϣ
    If Not mblnStructAdress Then Exit Function
    If padd��ͥ��ַ.Enabled Then
        If padd��ͥ��ַ.value <> "" Then
           strSQL = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,3,'" & padd��ͥ��ַ.valueʡ & "','" & _
               padd��ͥ��ַ.value�� & "','" & padd��ͥ��ַ.value���� & "','" & padd��ͥ��ַ.value���� & "','" & _
               padd��ͥ��ַ.value��ϸ��ַ & "','" & padd��ͥ��ַ.Code & "')"
        Else
           strSQL = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,3)"
        End If
        zlAddArray cllPro, strSQL
    End If
    If padd���ڵ�ַ.Enabled Then
        If padd���ڵ�ַ.value <> "" Then
           strSQL = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,4,'" & padd���ڵ�ַ.valueʡ & "','" & _
               padd���ڵ�ַ.value�� & "','" & padd���ڵ�ַ.value���� & "','" & padd���ڵ�ַ.value���� & "','" & _
               padd���ڵ�ַ.value��ϸ��ַ & "','" & padd���ڵ�ַ.Code & "')"
        Else
           strSQL = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,4)"
        End If
        zlAddArray cllPro, strSQL
    End If
End Function
Private Function SaveModifyPati() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ĳ�����Ϣ
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-07 03:48:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str���� As String, str�������� As String
    Dim str��ͥ��ַ As String, str���ڵ�ַ As String
    Dim blnTrans As Boolean, strErrMsg As String
    On Error GoTo errHandle
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    
    '    Zl_������Ϣ_Update
    strSQL = "Zl_������Ϣ_Update("
    '      ����id_In     ������Ϣ.����id%Type,
    strSQL = strSQL & "" & mrsInfo!����ID & ","
    '      �����_In     ������Ϣ.�����%Type,
    strSQL = strSQL & "" & IIf(Trim(txt�����.Text) <> "", Val(txt�����.Text), "NULL") & ","
    '      סԺ��_In     ������Ϣ.סԺ��%Type,
    strSQL = strSQL & "" & IIf(Val(Nvl(mrsInfo!סԺ��)) = 0, "NULL", Val(Nvl(mrsInfo!סԺ��))) & ","
    '      �ѱ�_In       ������Ϣ.�ѱ�%Type,
    strSQL = strSQL & "'" & NeedName(cbo�ѱ�.Text) & "',"
    '      ҽ�Ƹ���_In   ������Ϣ.ҽ�Ƹ��ʽ%Type,
    strSQL = strSQL & "'" & NeedName(cboҽ�Ƹ���.Text) & "',"
    '      ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & txtPatient.Text & "',"
    '      �Ա�_In       ������Ϣ.�Ա�%Type,
    strSQL = strSQL & "'" & NeedName(cbo�Ա�.Text) & "',"
    '      ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '      ��������_In   ������Ϣ.��������%Type,
    strSQL = strSQL & "" & str�������� & ","
    '      �����ص�_In   ������Ϣ.�����ص�%Type,
    strSQL = strSQL & "'" & txt�����ص�.Text & "',"
    '      ���֤��_In   ������Ϣ.���֤��%Type,
    strSQL = strSQL & "'" & txt���֤��.Text & "',"
    '      ���_In       ������Ϣ.���%Type,
    strSQL = strSQL & "'" & NeedName(cbo���.Text) & "',"
    '      ְҵ_In       ������Ϣ.ְҵ%Type,
    strSQL = strSQL & "'" & NeedName(cboְҵ.Text, mstrCboSplit) & "',"
    '      ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & NeedName(cbo����.Text) & "',"
    '      ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "'" & NeedName(cbo����.Text) & "',"
    '      ѧ��_In       ������Ϣ.ѧ��%Type,
    strSQL = strSQL & "'" & NeedName(cboѧ��.Text) & "',"
    '      ����_In       ������Ϣ.����״��%Type,
    strSQL = strSQL & "'" & NeedName(cbo����״��.Text) & "',"
    '      ��ͥ��ַ_In   ������Ϣ.��ͥ��ַ%Type,
    strSQL = strSQL & "'" & IIf(mblnStructAdress, padd��ͥ��ַ.value, txt��ͥ��ַ.Text) & "',"
    '      ��ͥ�绰_In   ������Ϣ.��ͥ�绰%Type,
    strSQL = strSQL & "'" & txt��ͥ�绰.Text & "',"
    '      ��ͥ��ַ�ʱ�_In   ������Ϣ.��ͥ��ַ�ʱ�%Type,
    strSQL = strSQL & "'" & txt��ͥ�ʱ�.Text & "',"
    '      ��ϵ������_In ������Ϣ.��ϵ������%Type,
    strSQL = strSQL & "'" & txt��ϵ������.Text & "',"
    '      ��ϵ�˹�ϵ_In ������Ϣ.��ϵ�˹�ϵ%Type,
    strSQL = strSQL & "'" & NeedName(cbo��ϵ�˹�ϵ.Text) & "',"
    '      ��ϵ�˵�ַ_In ������Ϣ.��ϵ�˵�ַ%Type,
    strSQL = strSQL & "'" & txt��ϵ�˵�ַ.Text & "',"
    '      ��ϵ�˵绰_In ������Ϣ.��ϵ�˵绰%Type,
    strSQL = strSQL & "'" & txt��ϵ�˵绰.Text & "',"
    '      ��ͬ��λid_In ������Ϣ.��ͬ��λid%Type,
    strSQL = strSQL & "" & IIf(Val(lbl������λ.Tag) = 0, "NULL", Val(lbl������λ.Tag)) & ","
    '      ������λ_In   ������Ϣ.������λ%Type,
    strSQL = strSQL & "'" & txt������λ.Text & "',"
    '      ��λ�绰_In   ������Ϣ.��λ�绰%Type,
    strSQL = strSQL & "'" & txt��λ�绰.Text & "',"
    '      ��λ�ʱ�_In   ������Ϣ.��λ�ʱ�%Type,
    strSQL = strSQL & "'" & txt��λ�ʱ�.Text & "',"
    '      ��λ������_In ������Ϣ.��λ������%Type,
    strSQL = strSQL & "'" & txt��λ������.Text & "',"
    '      ��λ�ʺ�_In   ������Ϣ.��λ�ʺ�%Type,
    strSQL = strSQL & "'" & txt��λ�ʻ�.Text & "',"
    '      ������_In     ������Ϣ.������%Type,
    strSQL = strSQL & "'" & Nvl(mrsInfo!������) & "',"
    '      ������_In     ������Ϣ.������%Type,
    strSQL = strSQL & "" & Val(Nvl(mrsInfo!������)) & ","
    '      ����_In       ������Ϣ.����%Type,
    strSQL = strSQL & "" & IIf(Val(Nvl(mrsInfo!����)) = 0, "NULL", Val(Nvl(mrsInfo!����))) & ","
    '      סԺ�ѱ�_In   Number := 0, --�Ƿ��޸ĵ��ǲ��˵�סԺ�ѱ�
    strSQL = strSQL & "0,"
    '      ҽ����_In     �����ʻ�.ҽ����%Type := Null,
    strSQL = strSQL & "" & IIf(Trim(txtҽ����.Text) = "", "NULL", "'" & Trim(txtҽ����.Text) & "'") & ","
    '      ����_In       ������Ϣ.����%Type := Null,
    strSQL = strSQL & "'" & NeedName(txt����.Text) & "',"
    '      ��������_In   ������Ϣ.��������%Type := Null,
    strSQL = strSQL & "" & Val(Nvl(mrsInfo!��������)) & ","
    '      ����Ա���_In ���˵�����¼.����Ա���%Type := Null,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '      ����Ա����_In ���˵�����¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '      ����֤��_In   ������Ϣ.����֤��%Type := Null,
    strSQL = strSQL & "'" & txt����֤��.Text & "',"
    '      ��������_In   ������ҳ.��������%Type := Null,
    strSQL = strSQL & "'" & Nvl(mrsInfo!��������) & "',"
    '      ��ע_In       ������ҳ.��ע%Type := Null
     strSQL = strSQL & "'" & Nvl(mrsInfo!��ע) & "',"
    '�����:51071
    '  ����_In   ������Ϣ.����%Type := Null
    strSQL = strSQL & "'',"
    '  ���ڵ�ַ_In   ������Ϣ.���ڵ�ַ%Type := Null
    strSQL = strSQL & "'" & IIf(mblnStructAdress, Trim(padd���ڵ�ַ.value), Trim(txt���ڵ�ַ.Text)) & "',"
    '  ���ڵ�ַ�ʱ�_In   ������Ϣ.���ڵ�ַ�ʱ�%Type := Null
    strSQL = strSQL & "'" & Trim(txt���ڵ�ַ�ʱ�.Text) & "',"
     '     ��ϵ�����֤��_In       ������Ϣ.��ϵ�����֤��%Type := Null WJ
    strSQL = strSQL & "'" & Trim(txt��ϵ�����֤��.Text) & "',"
    '   ģ���_In         Number := 0 --�޸Ĳ����������Ա����䡢�������ڵ�ģ��
    strSQL = strSQL & "" & mlngModule & ","
    '  �ֻ���_In         ������Ϣ.�ֻ���%Type :=Null
    strSQL = strSQL & "'" & txt�ֻ�.Text & "')"
    
    '89242:���ϴ�,2015/12/10,���²��˵�ַ��Ϣ
    '��ͥ��ַ
     If mblnStructAdress And padd��ͥ��ַ.Enabled Then
        If padd��ͥ��ַ.value <> "" Then
           str��ͥ��ַ = "zl_���˵�ַ��Ϣ_update(1," & mrsInfo!����ID & ",NULL,3,'" & padd��ͥ��ַ.valueʡ & "','" & _
               padd��ͥ��ַ.value�� & "','" & padd��ͥ��ַ.value���� & "','" & padd��ͥ��ַ.value���� & "','" & _
               padd��ͥ��ַ.value��ϸ��ַ & "','" & padd��ͥ��ַ.Code & "')"
        Else
           str��ͥ��ַ = "zl_���˵�ַ��Ϣ_update(2," & mrsInfo!����ID & ",NULL,3)"
        End If
    End If
    '���ڵ�ַ
    If mblnStructAdress And padd���ڵ�ַ.Enabled Then
        If padd���ڵ�ַ.value <> "" Then
           str���ڵ�ַ = "zl_���˵�ַ��Ϣ_update(1," & mrsInfo!����ID & ",NULL,4,'" & padd���ڵ�ַ.valueʡ & "','" & _
               padd���ڵ�ַ.value�� & "','" & padd���ڵ�ַ.value���� & "','" & padd���ڵ�ַ.value���� & "','" & _
               padd���ڵ�ַ.value��ϸ��ַ & "','" & padd���ڵ�ַ.Code & "')"
        Else
           str���ڵ�ַ = "zl_���˵�ַ��Ϣ_update(2," & mrsInfo!����ID & ",NULL,4)"
        End If
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If str��ͥ��ַ <> "" Then zlDatabase.ExecuteProcedure str��ͥ��ַ, Me.Caption
    If str���ڵ�ַ <> "" Then zlDatabase.ExecuteProcedure str���ڵ�ַ, Me.Caption
    
    '101170:���ϴ�,2017/5/3,����HIS����Ҫ�ύEMPI���ݣ�ʧ�ܺ��������ݶ�Ҫ����
    If zlSaveEMPIPatiInfo(False, mrsInfo!����ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "��EMPIƽ̨�ϴ�������Ϣʧ�ܣ�"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    SaveModifyPati = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LoadSaveNotoCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص��žݺŸ�Combox
    '����:���˺�
    '����:2011-07-12 18:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTmp As String
    If Not (mEditType = Cr_���� And lblNo.Tag <> "") Then Exit Sub
    '���뵥����ʷ��¼(�������͵���)
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    strTmp = lblNo.Tag & strTmp
    stbThis.Panels(2).Text = "�ϴα��浥��:" & lblNo.Tag
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For 'ֻ��ʾ10��
    Next
End Sub
Private Function IsCheckCancelValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷�ʱ��������Ч��
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, bln���ѿ� As Boolean, lng�����ID As Long
    Dim str��֤����  As String
   '����:48249
    Dim strSQL As String, rsBill As Recordset, rsTemp As ADODB.Recordset, lngCardBill As Long
    
    strName = IIf(glngSys \ 100 = 8, "��Ա��", "ҽ�ƿ�")
    
    If cboNO.Tag = "" Then
        MsgBox "��" & strName & "���ż�¼δ��ȷ��ȡ,�����˿���", vbExclamation, gstrSysName
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Exit Function
    End If
    
    If InStr(1, "12", mParaData.int�˿�ģʽ) > 0 And txtˢ������.Visible Then
        str��֤���� = Trim(txt����.Text)
            If Trim(txtˢ������) = "" Or str��֤���� <> Trim(lblˢ����֤.Tag) Then
                If mParaData.int�˿�ģʽ = 1 Then
                    MsgBox "�˿���֤ʧ�ܣ�����ˢ����֤��", vbExclamation, gstrSysName
                    If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
                Else
                    MsgBox "�˿���֤ʧ�ܣ���˶�ʵ�ʿ����뵱ǰ���ݿ����Ƿ�һ�£�", vbExclamation, gstrSysName
                    If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
                End If
                Exit Function
            End If
    End If
    If mcolBillBalance Is Nothing Then IsCheckCancelValied = True: Exit Function
    '92895:���ϴ�,2016/1/21,���ѿ���־�ڵ�3λ
    bln���ѿ� = Val(mcolBillBalance(1)(2)) = 1
    lng�����ID = mcolBillBalance(1)(0)
    If lng�����ID = 0 Then IsCheckCancelValied = True: Exit Function
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    '��Ϊ��,��Ҫ��ȡ��Ӧ��֧������
    Set mobjDelObject = zlGetClsCardObject(lng�����ID, bln���ѿ�)
    '92895:���ϴ�,2016/1/21,δ���ö�����nothing
    If mobjDelObject Is Nothing Then
        MsgBox "��δ���÷���ʱʹ�õ�֧���ӿ� ,�����ڴ˹���վ�Ͻ����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not mobjDelObject.CardPreporty.���� Then
        MsgBox "��δ����" & mobjDelObject.CardPreporty.���� & "�ӿ� ,�����ڴ˹���վ�Ͻ����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelObject.CardObject Is Nothing Then
        If zlCreatePatiCardObject(mobjDelObject.CardPreporty, mobjDelObject.CardObject) = False Then
            Exit Function
        End If
    End If
    If Not mobjDelObject.InitCompents Then
        If mobjDelObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
              Exit Function
        End If
        mobjDelObject.InitCompents = True
    End If
     
    '4.3.3.2.6   zlReturnCheck:�ʻ����˽���ǰ�ļ��
    'zlPaymentCheck�ʻ��ۿ�׼��
    '������  ��������    ��/��   ��ע
    'frmMain Object  In  ���õ�������
    'lngModule   Long    In  ģ���
    'lngCardTypeID   Long    In  �����ID:ҽ�ƿ����.ID
    'strCardNo   String  IN  ����
    'strBalanceIDs:��ʽ:�շ�����( 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�)|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    'dblMoney    Double  IN  �˿���
    'strSwapNo   String  In  ������ˮ��(�˿�ʱ���)
    'strSwapMemo String  In  ����˵��(�˿�ʱ����)
    '    Boolean ��������    True:���óɹ�,False:����ʧ��
    '˵��:
    '�ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬��ˣ��ٵ��û��˽���ǰ���Ƚ������ݵĺϷ��Լ��,�Ա�������������
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID
    'mcolBillBalance.Add Array(Val(Nvl(rsTmp!�����ID)), Trim(Nvl(rsTmp!����)), IIf(Val(Nvl(rsTmp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTmp!������ˮ��)), Trim(Nvl(rsTmp!����˵��))), strNO
    Dim str���� As String, str������ˮ�� As String, str����˵�� As String, str������Ϣ As String
    Dim strXMLExpend As String
    str���� = mcolBillBalance(1)(1)
    str������ˮ�� = mcolBillBalance(1)(3)
    str����˵�� = mcolBillBalance(1)(4)
    str������Ϣ = "5|" & mcolBillBalance(1)(6)
    If mobjDelObject.CardObject.zlReturncheck(Me, mlngModule, lng�����ID, str����, str������Ϣ, Val(txt����.Text), str������ˮ��, str����˵��, strXMLExpend) = False Then
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus: Exit Function
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
        Exit Function
    End If
    '100610:���ϴ�,2016/10/13��Ԥ���˿������˿��Ƿ���֤ˢ��
    If mobjDelObject.CardPreporty.�Ƿ��˿��鿨 Then
    '   zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
        '       <IN>
        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        '       </IN>
        Err = 0: On Error Resume Next
        If mobjDelObject.CardObject.zlBrushCard(Me, mlngModule, lng�����ID, _
         Nvl(mrsInfo!����), Nvl(mrsInfo!�Ա�), Nvl(mrsInfo!����), Val(txt����.Text), _
         mCurPayMoney.strˢ������, mCurPayMoney.strˢ������, "<IN><CZLX>2</CZLX></IN>") = False Then
            If Err = 450 Then
                Err = 0: On Error GoTo errHandle
                If mobjDelObject.CardObject.zlBrushCard(Me, mlngModule, lng�����ID, _
                 Nvl(mrsInfo!����), Nvl(mrsInfo!�Ա�), Nvl(mrsInfo!����), Val(txt����.Text), mCurPayMoney.strˢ������, mCurPayMoney.strˢ������) = False Then Exit Function
            Else
                Exit Function
            End If
        End If
    End If
    IsCheckCancelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsCheckCancel��Ԥ��()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˿�ʱ��鲡���Ƿ���Ԥ����δ��
     '����:��Ч,����true,���򷵻�False
    '����:����
    '����:2012-07-16 18:50:36
    '�����:51537
    '�����:50891
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim msgBoxResult As String
    Dim strSQL As String
    Dim blnOneCard As Boolean   '�Ƿ���Ψһһ��ҽ�ƿ�
    Dim rsBill As Recordset, rsCard As Recordset
    '69483,������,2014-01-15,����ҽ�ƿ��˿��˿��
    strSQL = "Select Count(1) As ҽ�ƿ��� From ����ҽ�ƿ���Ϣ Where ״̬=0 And ����ID=[1]"
    Set rsCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    strSQL = _
            "Select Ԥ�����,������� From ������� Where ����=1 And ����=1 And ����ID=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    '����:48249
    If InStr(1, mstrPrepayPrivs, ";Ԥ���˿�;") > 0 Then
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!Ԥ�����, 0) - Nvl(rsBill!�������, 0), "0.00") > 0 Then
                '�����:51537
                '�����:50891
                '108836�����ϴ���2017/6/28�������˿�����
                msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, vbNewLine & "�ò�������Ԥ�����δ��,�Ƿ��Ƚ�������˿����˿�?" & vbNewLine, "��������˿�,���˿�,ȡ��", Me, vbQuestion)
                If msgBoxResult = "��������˿�" Then '��Ԥ��������
                   '���ÿ��Ƿ��Ǽ����շ�(�˿������ʱӦ�ðѼ��˵ķ����㵽���������һ���˸�����)
                    '��������˿�
                    '�����:112995,����,2017/10/13,�˿��˷�ʱ��ʾ�����˷ѽ��
                     blnOneCard = IIf(rsCard!ҽ�ƿ��� = 1, True, False)
                     IsCheckCancel��Ԥ�� = zlPrepayFunc(2, mlng����ID, blnOneCard)
                     Exit Function
                ElseIf msgBoxResult = "ȡ��" Or msgBoxResult = "" Then
                     IsCheckCancel��Ԥ�� = False
                     Exit Function
                ElseIf msgBoxResult = "���˿�" Then
                    If rsCard!ҽ�ƿ��� = 1 Then
                        MsgBox "�ò�������Ԥ�������ܶԲ���Ψһ��ҽ�ƿ������˿�����!", vbInformation, gstrSysName
                        IsCheckCancel��Ԥ�� = False
                        Exit Function
                    End If
                End If
            Else
            '�����:51537
            '�����:50891
                msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "��ȷ��Ҫ�����˿�������?", "�˿�,ȡ��", Me, vbQuestion)
                If msgBoxResult = "" Or msgBoxResult = "ȡ��" Then
                    IsCheckCancel��Ԥ�� = False
                    Exit Function
                End If
            End If
        Else
        '�����:51537
        '�����:50891
           msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "��ȷ��Ҫ�����˿�������?", "�˿�,ȡ��", Me, vbQuestion)
           If msgBoxResult = "" Or msgBoxResult = "ȡ��" Then
                IsCheckCancel��Ԥ�� = False
                Exit Function
           End If
        End If
    Else
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!Ԥ�����, 0) - Nvl(rsBill!�������, 0), "0.00") > 0 Then
                If rsCard!ҽ�ƿ��� = 1 Then
                    MsgBox "��û��Ԥ���˿�Ȩ�ޣ����ܶԲ���Ψһ��ҽ�ƿ��˿�����!", vbInformation, gstrSysName
                    IsCheckCancel��Ԥ�� = False
                    Exit Function
                End If
            End If
        End If
        If MsgBox("��û��Ԥ���˿�Ȩ��,�Ƿ���������˿�����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then IsCheckCancel��Ԥ�� = False: Exit Function
    End If
        IsCheckCancel��Ԥ�� = True
End Function

Private Function SaveDelete(strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˿���
    '���:strNO-����ĵ��ݺ�
    '����:�˺ųɹ�,����true,����ķ���False
    '����:���˺�
    '����:2011-07-12 18:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnTrans As Boolean, bln���ѿ� As Boolean, lng�����ID As Long
    Dim lng����ID As Long, blnOraclTrans As Boolean
    On Error GoTo errH
    strSQL = "zl_ҽ�ƿ���¼_DELETE('" & strNo & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If CallBackBalanceInterface(strNo, blnOraclTrans) = False Then
        If blnOraclTrans = False Then gcnOracle.RollbackTrans
        Exit Function
    End If
    If blnOraclTrans = False Then gcnOracle.CommitTrans
    blnTrans = False
    SaveDelete = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CallBackBalanceInterface(ByVal strNo As String, ByRef blnTrancs As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���û��˽ӿ�
    '���:blnTrancs-�Ƿ���������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, strSwapGlideNO As String, strSwapMemo As String, str������Ϣ As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, lng����ID As Long, cllPro As Collection
    Dim bln���ѿ� As Boolean, lng�����ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim str������Ϣ As String, strTemp As String
    
    On Error GoTo errHandle
    blnTrancs = False
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    'mcolBillBalance.Add Array(Val(Nvl(rsTmp!�����ID)), Trim(Nvl(rsTmp!����)), IIf(Val(Nvl(rsTmp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTmp!������ˮ��)), Trim(Nvl(rsTmp!����˵��))), strNO
    If mcolBillBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    '92895:���ϴ�,2016/1/21,���ѿ���־�ڵ�3λ
    bln���ѿ� = Val(mcolBillBalance(1)(2)) = 1
    lng�����ID = mcolBillBalance(1)(0)
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    If lng�����ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str���� = mcolBillBalance(1)(1)
    strSwapGlideNO = mcolBillBalance(1)(3)
    strSwapMemo = mcolBillBalance(1)(4)
    str������Ϣ = "5|" & mcolBillBalance(1)(6)
    strSQL = "Select ����ID,���ʷ��� From סԺ���ü�¼  Where ��¼����=5 and NO=[1] and ��¼״̬=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTemp.EOF Then
        gcnOracle.RollbackTrans: blnTrancs = True
        MsgBox "δ�ҵ��˿���Ϣ�����ܼ���", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng����ID = Val(Nvl(rsTemp!����id))
    '81489,Ƚ����,2015-4-29,�˷Ѵ������ID
    strSwapExtendInfor = "5|" & lng����ID: strTemp = strSwapExtendInfor
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ���˽���
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
    '       strCardNo-����
    '       strBalanceIDs-����֧�����漰�Ľ���ID(����ԭ����ID):
    '                           ��ʽ:�շ�����(|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
    '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
    '       strSwapExtendInfor-���׵���չ��Ϣ
    '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
    If mobjDelObject.CardObject.zlReturnMoney(Me, mlngModule, lng�����ID, str����, str������Ϣ, Val(txt����.Text), strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans: blnTrancs = True
        Exit Function
    End If
    
    '���½�����Ϣ
    '    Zl_�����ӿڸ���_Update
    strSQL = "Zl_�����ӿڸ���_Update("
    '  �����id_In   ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '  ���ѿ�_In     Number,
    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ����ids_In    Varchar2,
    strSQL = strSQL & "'" & lng����ID & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    strSQL = strSQL & "'" & strSwapGlideNO & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type
    strSQL = strSQL & "'" & strSwapMemo & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    If strTemp <> strSwapExtendInfor Then
        'strSwapExtendInfor:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
        varData = Split(strSwapExtendInfor, "||")
        Set cllPro = New Collection
        For i = 0 To UBound(varData)
            If Trim(varData(i)) <> "" Then
                varTemp = Split(varData(i) & "|", "|")
                If varTemp(0) <> "" Then
                    strTemp = varTemp(0) & "|" & varTemp(1)
                    If zlCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
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
                        strSQL = strSQL & "'" & lng����ID & "',"
                        '������Ϣ_In Varchar2:������Ŀ|��������||...
                        strSQL = strSQL & "'" & str������Ϣ & "')"
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
            strSQL = strSQL & "'" & lng����ID & "',"
            '������Ϣ_In Varchar2:������Ŀ|��������||...
            strSQL = strSQL & "'" & str������Ϣ & "')"
            zlAddArray cllPro, strSQL
        End If
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllPro, Me.Caption
    End If
    CallBackBalanceInterface = True: blnTrancs = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans: blnTrancs = True
    Call ErrCenter
    Exit Function
ErrOthers:
    '��չ��Ϣ,������һ����,�Ա��֤
    If ErrCenter() = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    CallBackBalanceInterface = True
    gcnOracle.CommitTrans: blnTrancs = True
End Function
Private Function IsCheckChangeCardValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��黻���������Ƿ�Ϸ�
    '����:���ݺϷ�,����True,���򷵻�False
    '����:���˺�
    '����:2011-07-14 11:06:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If lblˢ����֤.Tag = "" Then
        If Trim(txtˢ������.Text) = "" Then
            MsgBox "ԭʼ����δ����ˢ��ȷ��,���ܻ���!", vbInformation + vbOKOnly, gstrSysName
            If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
            Exit Function
        End If
        '-1-�ɹ�;0-ʧ��;1-�ü�¼������
        Select Case ReadCardNo(Trim(txtˢ������.Text), 2)
        Case 0
            Exit Function
        Case 1
            MsgBox "δ�ҵ�ԭʼ���ŵĳ�����,����!", vbInformation + vbOKOnly, gstrSysName
            If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
            Exit Function
        End Select
    End If
    If mrsInfo Is Nothing Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If IsCheckҽ�ƿ� = False Then Exit Function
     '�����:50893
    If CStr(txtԭ������.Tag) <> zlCommFun.zlStringEncode(Trim(txtԭ������.Text)) Then
        MsgBox "ԭ�������������,��������������!", vbInformation + vbOKOnly, gstrSysName
        If txtԭ������.Enabled And txtԭ������.Visible Then txtԭ������.SetFocus
        Exit Function
    End If
    IsCheckChangeCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function IsCheckFillCardValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��黻���������Ƿ�Ϸ�
    '����:���ݺϷ�,����True,���򷵻�False
    '����:���˺�
    '����:2011-07-14 11:06:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
     If mrsInfo Is Nothing Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
   If IsCheckҽ�ƿ� = False Then Exit Function
    IsCheckFillCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveChangeCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���滻��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-14 11:50:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, Curdate As Date, lng����ID As Long, cllPro As Collection
    On Error GoTo errHandle
    lng����ID = Val(Nvl(mrsInfo!����ID))
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
    Call AddCardDataSQL(lng����ID, Curdate, cllPro, lng����ID)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveChangeCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function SaveFillCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���油����Ϣ
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-14 11:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng����ID As Long, Curdate As Date, lng����ID As Long
   Dim cllPro As Collection, cllUpdateSwap As Collection, cllThree As Collection
    On Error GoTo errHandle
    lng����ID = Val(Nvl(mrsInfo!����ID))
    
    Set cllPro = New Collection
    Set cllUpdateSwap = New Collection: Set cllThree = New Collection
    Curdate = zlDatabase.Currentdate
    Call AddCardDataSQL(lng����ID, Curdate, cllPro, lng����ID)
    If IDKindPayMode.IDKind = 2 And Val(txt���.Text) > 0 Then Call AddDepositSQL(lng����ID, Curdate, cllPro, lng����ID)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If zlInterfacePrayMoney(cllUpdateSwap, cllThree) = False Then
        gcnOracle.RollbackTrans
    End If
    zlExecuteProcedureArrAy cllUpdateSwap, Me.Caption, False, True
    On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllThree, Me.Caption
    SaveFillCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    If ErrCenter = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    gcnOracle.CommitTrans
End Function
Private Function isCheckLossValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʧ���ݵĺϷ���
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-14 13:40:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
   If mrsInfo Is Nothing Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "������Ϣδ�ҵ�,����ȷ��������Ϣ!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    If lblˢ����֤.Tag = "" Then
        If Trim(txtˢ������.Text) = "" Then
            MsgBox "��ʧ�Ŀ���δ����ˢ��ȷ��,���ܹ�ʧ!", vbInformation + vbOKOnly, gstrSysName
            If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
            Exit Function
        End If
        
        '-1-�ɹ�;0-ʧ��;1-�ü�¼������
        Select Case ReadCardNo(Trim(txtˢ������.Text), 2)
        Case 0
            Exit Function
        Case 1
            MsgBox "δ�ҵ���ǰ���ŵĳ�����,����!", vbInformation + vbOKOnly, gstrSysName
            If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
            Exit Function
        End Select
    End If
    isCheckLossValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function SaveLossCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʧ��Ϣ
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-14 11:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng����ID As Long, Curdate As Date, lng����ID As Long, cllPro As Collection
   Dim strSQL As String
   On Error GoTo errHandle
    lng����ID = Val(Nvl(mrsInfo!����ID))
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
      'Zl_ҽ�ƿ��䶯_Insert
       strSQL = "Zl_ҽ�ƿ��䶯_Insert("
      '      �䶯����_In   Number,
      '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
      strSQL = strSQL & "" & 6 & ","
      '      ����id_In     סԺ���ü�¼.����id%Type,
      strSQL = strSQL & "" & lng����ID & ","
      '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
      strSQL = strSQL & "" & mlngCardTypeID & ","
      '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & "'" & lblˢ����֤.Tag & "',"
      '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & "'" & lblˢ����֤.Tag & "',"
      '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
      '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
      strSQL = strSQL & "'" & Trim(txt�䶯ԭ��.Text) & "',"
      '      ����_In       ������Ϣ.����֤��%Type,
      strSQL = strSQL & "NULL,"
      '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
      strSQL = strSQL & "'" & UserInfo.���� & "',"
      '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
      strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
      '      Ic����_In     ������Ϣ.Ic����%Type := Null,
      strSQL = strSQL & "NULL,"
      '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
      strSQL = strSQL & "'" & cbo��ʧ��ʽ.Text & "')"
     Call zlAddArray(cllPro, strSQL)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveLossCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdCreateCard_Click()
    '�����:56599
    Dim strExpend As String
    Dim blnFlag As Boolean
    Dim strOutPatiInforXml As String

    If mrsInfo Is Nothing Then
        MsgBox "������Ϣ�����ڻ���δ�ڱ�Ժ����,���ܽ����ƿ�������", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Sub
    End If
    If mobjReadCard Is Nothing Then Exit Sub
    If mobjReadCard.zlMakeCard(Me, mlngModule, mlngCardTypeID, Get�ƿ�XML(mrsInfo!����ID), mstr�ɼ�ͼƬ, strOutPatiInforXml, strExpend) = False Then Exit Sub
    If strOutPatiInforXml <> "" Then
        LoadPati strOutPatiInforXml
    End If
End Sub

 
Private Sub cmdOK_Click()
    Dim blnPrint As Boolean, blnPlugInCheck As Boolean
    
    Call txt���_Change
    If CheckDepositFactValied = False Then Exit Sub
    If mEditType = Cr_�˿� Or chkCancel.value = 1 Then
       If IsCheckCancelValied = False Then Exit Sub
       If IsCheckCancel��Ԥ�� = False Then Exit Sub '�����:51537
       If SaveDelete(cboNO.Tag) = False Then Exit Sub
        mintSucces = 1
        If mEditType = Cr_�˿� Then
            mblnChange = False
            Unload Me: Exit Sub
        End If
        chkCancel.value = 0
        If Me.txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call ClearData
        mblnChange = False
        Exit Sub
    End If
    If mEditType = Cr_���� Then
        If IsCheckFillCardValied = False Then Exit Sub
        'ˢ������
        If CheckBrushCard = False Then Exit Sub
        If SaveFillCard = False Then Exit Sub
        Call PrintBill(mCurPayMoney.strNo, Trim(txt����.Text))
        mintSucces = 1
        Call ClearData(True)
        Unload Me: Exit Sub
    End If
    If mEditType = Cr_��ʧ Then
        If isCheckLossValied = False Then Exit Sub
        If SaveLossCard = False Then Exit Sub
        Call ClearData
        mintSucces = 1: Unload Me: Exit Sub
    End If
    If Not isValied Then Exit Sub
    
    If mEditType = Cr_����������Ϣ Then
        If mrsInfo Is Nothing Then Exit Sub
        If mrsInfo.State <> 1 Then Exit Sub
        '������Ĳ��˵Ļ�����Ϣ,����ҵ���,���ܽ��е���
        If Nvl(mrsInfo!����) <> txtPatient.Text _
            Or Nvl(mrsInfo!�Ա�) <> NeedName(cbo�Ա�.Text) _
            Or Nvl(mrsInfo!����) <> txt����.Text & cbo���䵥λ _
            Or Format(mrsInfo!��������, "yyyy-mm-dd") <> txt��������.Text Then
            If zlExistOperationData(mrsInfo!����ID, "") Then
                MsgBox "�ò����Ѿ�������ҽ��ҵ������,���ܽ��в��˵Ļ�����Ϣ����(����,�Ա�,����,�������ڵ�),���ڡ�������Ϣ�����н��е���!", vbOKOnly + vbInformation, gstrSysName
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                zlControl.TxtSelAll txtPatient
                Exit Sub
            End If
        End If
        If SaveModifyPati = False Then Exit Sub
        mintSucces = 1
        Call ClearData
        Unload Me: Exit Sub
    End If
    
    If mEditType = Cr_���� Then
        If IsCheckChangeCardValied = False Then Exit Sub
        If SaveChangeCard = False Then Exit Sub
        mintSucces = 1
        Call ClearData
        Unload Me: Exit Sub
    End If
           
    If mEditType = Cr_�󶨿� Or mEditType = Cr_���� Then
    
        If IsCheckҽ�ƿ� = False Then Exit Sub
        
        'ˢ������
        If CheckBrushCard = False Then Exit Sub
        '�����:51072
        If Len(Trim(txtPass.Text)) = 0 Then 'û�����뿨��������
           If zl_Get����Ĭ�Ϸ������� = False Then Exit Sub
        End If
        
        '�����56599
        If InoculateValid = False Then Exit Sub

        '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
        If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '������������Ϣǰ��������Ч�Լ��
            On Error Resume Next
            blnPlugInCheck = mobjPlugIn.PatiInfoSaveBefore(mlng����ID)
            Call zlPlugInErrH(Err, "PatiInfoSaveBefore")
            If Err = 0 And blnPlugInCheck = False Then
                Exit Sub '���δͨ����ֹ����
            End If
            Err.Clear
        End If

        If SaveData = False Then Exit Sub
        If mEditType = Cr_���� Then
            Call PrintBill(mCurPayMoney.strNo, Trim(txt����.Text))
        End If
        
        mintSucces = mintSucces + 1
        Call LoadSaveNotoCombox: Call ClearData(True)
        Call CheckBILL("")
        If txtPatient.Enabled And txtPass.Visible Then txtPatient.SetFocus
        mintSucces = mintSucces + 1
        Exit Sub
    End If
    mintSucces = mintSucces + 1
    Call ClearData
    Unload Me
End Sub

Private Sub PrintBill(ByVal strNo As String, ByVal strCardNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�з���Ʊ�ݴ�ӡ
    '������strNO ���￨NO��
    '      strCardNo ����
    '����:���ϴ�
    '����:2014-04-10 13:41:24
    '�����:57950
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    On Error GoTo Errhand
    
    If mblnPrepayPrint Then
        '��ӡԤ��Ʊ��
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mstrPrePayNo, "����ID=" & mlngԤ������ID, "�տ�ʱ��=" & Format(mdatԤ��ʱ��, "yyyy-mm-dd HH:MM:SS"), 2)
    End If
    blnPrint = False
    Select Case Val(zlDatabase.GetPara("������ӡ��ʽ", glngSys, mlngModule))
     Case 0 '����ӡ
     Case 2 'ѡ���Ƿ��ӡ
         If MsgBox("���Ƿ����Ҫ��ӡ����ƾ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
             blnPrint = True
         End If
     Case Else
          blnPrint = True
    End Select
    If Not blnPrint Then Exit Sub
    
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me, "NO=" & strNo, "����=" & strCardNo, "PrintEmpty=0", 2)
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Sub

Private Function zl_Get����Ĭ�Ϸ�������() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ�Ϸ�������
    '����:�Ƿ������������
    '����:����
    '����:2012-07-06 15:53:14
    '�����:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardType As clsCard
    Dim msgResult As VbMsgBoxResult
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    '59760
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    Set objCardType = objYLCards.Item("K" & mlngCardTypeID)
    If objCardType.�Ƿ�ȱʡ���� = False Then '������
        Select Case objCardType.������������
            Case 0 '������
                zl_Get����Ĭ�Ϸ������� = True
                Exit Function
            Case 1 'δ��������
               msgResult = MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName)
               zl_Get����Ĭ�Ϸ������� = IIf(msgResult = vbYes, True, False)
               Exit Function
            Case 2 'Ϊ�����ֹ
                 MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
                zl_Get����Ĭ�Ϸ������� = False
                Exit Function
        End Select
    ElseIf objCardType.�Ƿ�ȱʡ���� Then 'ȱʡ���֤��Nλ
        If Len(Trim(txt���֤��.Text)) > 0 Or Len(Trim(txt��ϵ�����֤��.Text)) > 0 Then '���������֤����ϵ�����֤��
            If Len(Trim(txt���֤��.Text)) > 0 Then '�����֤���������֤
                   txtPass.Text = Right(Trim(txt���֤��.Text), objCardType.���볤��)
            Else '������ô��������֤��Ϊ����
                   txtPass.Text = Right(Trim(txt��ϵ�����֤��.Text), objCardType.���볤��)
            End If
            txtAudi.Text = txtPass.Text
        Else '���֤����ϵ�����֤��û����
            Select Case objCardType.������������
                Case 0 '������
                    zl_Get����Ĭ�Ϸ������� = True
                    Exit Function
                Case 1 'δ��������
                    msgResult = MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName)
                    zl_Get����Ĭ�Ϸ������� = IIf(msgResult = vbYes, True, False)
                    Exit Function
                Case 2 'Ϊ�����ֹ
                    MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
                    zl_Get����Ĭ�Ϸ������� = False
                    Exit Function
            End Select
        End If
    End If
    zl_Get����Ĭ�Ϸ������� = True
End Function

Private Function CheckBILL(Optional strCardNo As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ʊ���Ƿ������
    '����:���˺�
    '����:2011-07-12 15:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '106010:���ϴ���2017/3/10�����ϸ���Ʒ����������ID
    Dim strSQL As String
    Dim rsTemp As Recordset
    If Not mCardType.bln�ϸ���� Then mCardType.lng����ID = 0: CheckBILL = True: Exit Function
    
    If mCardType.bln�Ƿ��ظ�ʹ�� Then
        mCardType.lng����ID = 0
        strSQL = "Select b.����Id" & vbNewLine & _
             "From Ʊ�����ü�¼ A, Ʊ��ʹ����ϸ B" & vbNewLine & _
             "Where a.Id = b.����id And a.Ʊ�� = 5 And (Nvl(a.ʹ�����, 'LXH') = [1] Or a.ʹ����� Is Null) And b.���� = [2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, strCardNo)
        If rsTemp.RecordCount > 0 Then
            mCardType.lng����ID = Nvl(rsTemp!����Id, 0)
        Else
            mCardType.lng����ID = CheckUsedBill(5, IIf(mCardType.lng����ID > 0, mCardType.lng����ID, mCardType.lng��������), strCardNo, mlngCardTypeID)
        End If
    Else
        mCardType.lng����ID = CheckUsedBill(5, IIf(mCardType.lng����ID > 0, mCardType.lng����ID, mCardType.lng��������), strCardNo, mlngCardTypeID)
    End If
    
    If mCardType.lng����ID <= 0 Then
        Select Case mCardType.lng����ID
            Case 0 '����ʧ��
            Case -1
                If txt����.Text <> "" Then MsgBox "����û�����ü����õ�" & mCardType.str������ & ",���ܷ��ţ�" & vbCrLf & _
                    "�����ڱ������ù������λ�����һ���¿�! ", vbExclamation, gstrSysName
                Exit Function
            Case -2
                If txt����.Text <> "" Then MsgBox "���ع��õ�" & mCardType.str������ & "������,���ܷ��ţ�" & vbCrLf & _
                    "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
                Exit Function
            Case -3
                MsgBox "���ſ��Ų�����Ч��Χ��,�����Ƿ���ȷˢ����", vbExclamation, gstrSysName
                If txt����.Enabled And txt����.Enabled Then txt����.SetFocus
                Exit Function
        End Select
    End If
    CheckBILL = True
End Function

Private Sub cmdPicClear_Click()
    '�����:56599
    imgPatient.Picture = Nothing
    mlngͼ����� = 3
End Sub

Private Sub cmdPicCollect_Click()
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.PatiImageGatherer(Me, mstr�ɼ�ͼƬ) = False Then Exit Sub
    imgPatient.Picture = LoadPicture(mstr�ɼ�ͼƬ)
    mlngͼ����� = 2
End Sub

Private Sub cmdPicFile_Click()
    '�����:56599
    Dim strFileDir As String
On Error GoTo ErrHanl:
    With cmdialog
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .Filter = "(*.bmp)|*.bmp"
        .FilterIndex = 2
        .ShowOpen
        strFileDir = .FileName
        If strFileDir = "" Then Exit Sub
        imgPatient.Picture = LoadPicture(strFileDir)
    End With
    mlngͼ����� = 1
    Exit Sub
ErrHanl:
    
End Sub

Private Sub cmdReadCard_Click()
    Call ReReadCard("")
End Sub

Private Function LoadPati(ByVal strPatiXML As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ,��ȡ������Ϣ
    '����:���˺�
    '����:2011-09-08 21:52:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '�����:56599
    Dim str����ҩ�� As String, str������Ӧ As String '�����:56599
    Dim str�������� As String, str�������� As String '�����:56599
    Dim strABOѪ�� As String '�����:56599
    Dim str��Ϣ�� As String, str��Ϣֵ As String '�����:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '�����:56599
    Dim str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, str��ַ As String '�����:56599
    Dim byt��Ϣ����ģʽ As Byte, lng����ID As Long
    On Error GoTo errHandle
    If Not (mEditType = Cr_�󶨿� Or mEditType = Cr_����) Then Exit Function
    If strPatiXML = "" Then Exit Function
    '92235:���ϴ�,2016/1/13,����ģʽ����
    If mrsInfo Is Nothing Then
    ElseIf mrsInfo.State <> 1 Then
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ��Ϣ����ģʽ Integer 1 '0-ǿ�Ƹ��£�1-�������˲����£�2-����������Ϣ��ȱ
    Call zlXML_GetNodeValue("��Ϣ����ģʽ", , strValue)
    byt��Ϣ����ģʽ = 0
    byt��Ϣ����ģʽ = Val(strValue)
    If byt��Ϣ����ģʽ = 1 And lng����ID <> 0 Then LoadPati = True: Exit Function
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue)
    '    ����    Varchar2    100
    Call zlXML_GetNodeValue("����", , strValue)
    '1-�������
    '2-�²���
    '3-�ϲ�����ϢΪ�յ������ȱ
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txtPatient.Text = "") Then txtPatient.Text = strValue
    '    �Ա�    Varchar2    4
    Call zlXML_GetNodeValue("�Ա�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo�Ա�.Text = "") Then
        If strValue <> "" Then
            Call zlControl.CboLocate(cbo�Ա�, strValue)
            If cbo�Ա�.ListIndex = -1 Then
                cbo�Ա�.AddItem strValue
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
        End If
    End If
    '    ����    Varchar2    10
    Call zlXML_GetNodeValue("����", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt����.Text = "") Then
        If strValue <> "" Then
            Call LoadOldData(strValue, txt����, cbo���䵥λ)
        End If
    End If
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("��������", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��������.Text = "") Then
        txt��������.Text = Format(IIf(strValue = "", "____-__-__", strValue), "YYYY-MM-DD")
        If strValue <> "" Then
             txt����.Text = ReCalcOld(CDate(txt��������.Text), cbo���䵥λ)      '�޸ĵ�ʱ��,���ݳ���������������
             If CDate(txt��������.Text) - CDate(strValue) <> 0 Then txt����ʱ��.Text = Format(strValue, "HH:MM")
        Else
            txt����ʱ��.Text = "__:__"
            txt��������.Text = ReCalcBirth(Val(txt����.Text), cbo���䵥λ.Text)
        End If
    End If
    '    �����ص�    Varchar2    50
    Call zlXML_GetNodeValue("�����ص�", , strValue)
    '    ���֤��    VARCHAR2    18
    Call zlXML_GetNodeValue("���֤��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt���֤��.Text = "") Then
        If strValue <> "" Then txt���֤��.Text = strValue
    End If
    '    ����֤��    Varchar2    20
    Call zlXML_GetNodeValue("����֤��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt����֤��.Text = "") Then
        If strValue <> "" Then txt����֤��.Text = strValue
    End If
    '    ְҵ    Varchar2    80
    Call zlXML_GetNodeValue("ְҵ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboְҵ.Text = "") Then
        If strValue <> "" Then
            cboְҵ.ListIndex = GetCboIndex(cboְҵ, strValue, , , mstrCboSplit)
            If cboְҵ.ListIndex = -1 Then
                cboְҵ.AddItem strValue, 0
                cboְҵ.ListIndex = cboְҵ.NewIndex
            End If
        End If
    End If
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo����.Text = "") Then
        cbo����.ListIndex = GetCboIndex(cbo����, strValue)
        If cbo����.ListIndex = -1 And strValue <> "" Then
            cbo����.AddItem strValue, 0
            cbo����.ListIndex = cbo����.NewIndex
        End If
    End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo����.Text = "") Then
        cbo����.ListIndex = GetCboIndex(cbo����, strValue)
        If cbo����.ListIndex = -1 And strValue <> "" Then
            cbo����.AddItem strValue, 0
            cbo����.ListIndex = cbo����.NewIndex
        End If
    End If
    '    ѧ��    Varchar2    10
    Call zlXML_GetNodeValue("ѧ��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboѧ��.Text = "") Then
        cboѧ��.ListIndex = GetCboIndex(cboѧ��, strValue)
        If cboѧ��.ListIndex = -1 And strValue <> "" Then
            cboѧ��.AddItem strValue, 0
            cboѧ��.ListIndex = cboѧ��.NewIndex
        End If
    End If
    '    ����״��    Varchar2    4
    Call zlXML_GetNodeValue("����״��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo����״��.Text = "") Then
        cbo����״��.ListIndex = GetCboIndex(cbo����״��, strValue)
        If cbo����״��.ListIndex = -1 And strValue <> "" Then
            cbo����״��.AddItem strValue, 0
            cbo����״��.ListIndex = cbo����״��.NewIndex
        End If
    End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt����.Text = "") Then txt����.Text = strValue
    '    ��ͥ��ַ    Varchar2    50
    Call zlXML_GetNodeValue("��ͥ��ַ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ͥ��ַ.Text = "") Then
        txt��ͥ��ַ.Text = strValue
        padd��ͥ��ַ.value = strValue
    End If
    '    ��ͥ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��ͥ�绰", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ͥ�绰.Text = "") Then txt��ͥ�绰.Text = strValue
    '    ��ͥ��ַ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��ͥ��ַ�ʱ�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ͥ�ʱ�.Text = "") Then txt��ͥ�ʱ�.Text = strValue
    '    ���ڵ�ַ    Varchar2    50
    Call zlXML_GetNodeValue("���ڵ�ַ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt���ڵ�ַ.Text = "") Then
        txt���ڵ�ַ.Text = strValue
        padd���ڵ�ַ.value = strValue
    End If
    '    �ֻ���    Varchar2    20
    Call zlXML_GetNodeValue("�ֻ���", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt�ֻ�.Text = "") Then txt�ֻ�.Text = strValue
    '    �໤��  Varchar2    64
    Call zlXML_GetNodeValue("�໤��", , strValue)
   'txt�໤��.Text = strValue
'    '    ��ϵ������  Varchar2    64
'    Call zlXML_GetNodeValue("��ϵ������", , strValue)
'    '    ��ϵ�˹�ϵ  Varchar2    30
'    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , strValue)
'    '    ��ϵ�˵�ַ  Varchar2    50
'    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , strValue)
'    txt��ϵ������.Text = strValue
'    '    ��ϵ�˵绰  Varchar2    20
'    Call zlXML_GetNodeValue("��ϵ�˵绰", , strValue)
'    txt��ϵ�˵绰.Text = strValue
    '    ������λ    Varchar2    100
    Call zlXML_GetNodeValue("������λ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt������λ.Text = "") Then
        txt������λ.Text = strValue
        lbl������λ.Tag = ""
    End If
    '    ��λ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��λ�绰", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ�绰.Text = "") Then txt��λ�绰.Text = strValue
    '    ��λ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��λ�ʱ�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ�ʱ�.Text = "") Then txt��λ�ʱ�.Text = strValue
    '    ��λ������  Varchar2    50
    Call zlXML_GetNodeValue("��λ������", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ������.Text = "") Then txt��λ������.Text = strValue
    '    ��λ�ʺ�    Varchar2    20
    Call zlXML_GetNodeValue("��λ�ʺ�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ�ʻ�.Text = "") Then txt��λ�ʻ�.Text = strValue
    '�����:56599
    '�������
    Call zlXML_GetRows("ҩ������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("ҩ������", i, str����ҩ��)
        Call zlXML_GetNodeValue("ҩ�ﷴӦ", i, str������Ӧ)
        SetDrugAllergy str����ҩ��, str������Ӧ, , lng����ID, byt��Ϣ����ģʽ
    Next
    lngCount = 0
    '���߼�¼
    Call zlXML_GetRows("��������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("��������", i, str��������)
        Call zlXML_GetNodeValue("����ʱ��", i, str��������)
        SetInoculate str��������, str��������
    Next
    lngCount = 0
    'ABOѪ��
    Call zlXML_GetNodeValue("ABOѪ��", , strABOѪ��)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboBloodType.Text = "") Then
        If strABOѪ�� <> "" Then
            For i = 0 To cboBloodType.ListCount - 1
                '76314,���ϴ���2014-08-06��������Ϣ��ȷ��ȡ
                If NeedName(cboBloodType.List(i), ".") = NeedName(strABOѪ��) Then cboBloodType.ListIndex = i
            Next
        End If
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboBH.Text = "") Then
        If strValue <> "" Then
            For i = 0 To cboBH.ListCount - 1
                If cboBH.List(i) = strValue Then cboBH.ListIndex = i
            Next
        End If
    End If
    'ҽѧ��ʾ
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("�ٴ�������Ϣ")
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txtMedicalWarning.Text = "") Then
        If Not xmlChildNodes Is Nothing Then
            If xmlChildNodes.length > 0 Then
                For i = 0 To xmlChildNodes.length - 1
                    Set xmlChildNode = xmlChildNodes(i)
                    If xmlChildNode.Text = "1" Then
                        strValue = strValue & ";" & Replace(xmlChildNode.nodeName, "��־", "")
                    End If
                Next
            End If
        End If
        If strValue <> "" Then txtMedicalWarning.Text = Mid(strValue, 2)
    End If
    '����ҽѧ��ʾ
    Call zlXML_GetNodeValue("����ҽѧ��ʾ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txtOtherWaring.Text = "") Then
        If strValue <> "" Then txtOtherWaring.Text = strValue
    End If
    '��ϵ��Ϣ
    '    ��ϵ�˵�ַ  Varchar2    50
    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , str��ַ)
    If byt��Ϣ����ģʽ = 0 Or lng����ID = 0 Or (lng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ϵ�˵�ַ.Text = "") Then txt��ϵ�˵�ַ.Text = str��ַ
     '    ��ϵ������  Varchar2    64
    Call zlXML_GetNodeValue("��ϵ������", , str����)
    '    ��ϵ�˹�ϵ  Varchar2    30
    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , str��ϵ)
    '    ��ϵ�˵绰  Varchar2    20
    Call zlXML_GetNodeValue("��ϵ�˵绰", , str�绰)
    '    ��ϵ�����֤ Varchar2   20
    Call zlXML_GetNodeValue("��ϵ�����֤��", , str���֤��)
    SetLinkInfo str����, str��ϵ, str�绰, str���֤��, lng����ID, byt��Ϣ����ģʽ
    
    Call zlXML_GetRows("��ϵ��Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("��ϵ��Ϣ", "����", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "����", i, j, str����)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "��ϵ", i, j, str��ϵ)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "�绰", i, j, str�绰)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "���֤��", i, j, str���֤��)
                SetLinkInfo str����, str��ϵ, str�绰, str���֤��, lng����ID, byt��Ϣ����ģʽ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '������Ϣ
    '�����������
    Call zlXML_GetNodeValue("�����������", , strValue)
    SetOtherInfo "�����������", strValue, lng����ID, byt��Ϣ����ģʽ
    
    '��ũ��֤��
    Call zlXML_GetNodeValue("��ũ��֤��", , strValue)
    SetOtherInfo "��ũ��֤��", strValue, lng����ID, byt��Ϣ����ģʽ

    '����֤��
    Call zlXML_GetRows("����֤��", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("����֤��", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣֵ", i, j, str��Ϣֵ)
                SetOtherInfo str��Ϣ��, str��Ϣֵ, lng����ID, byt��Ϣ����ģʽ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '������Ϣ
    Call zlXML_GetRows("������Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("������Ϣ", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣֵ", i, j, str��Ϣֵ)
                SetOtherInfo str��Ϣ��, str��Ϣֵ, lng����ID, byt��Ϣ����ģʽ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    'ҽ�ƿ�����
    Call zlXML_GetRows("ҽ�ƿ�����", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("ҽ�ƿ�����", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣֵ", i, j, str��Ϣֵ)
                If mdicҽ�ƿ�����.Exists(str��Ϣ��) Then
                    If Not (lng����ID <> 0 And byt��Ϣ����ģʽ = 2) Then mdicҽ�ƿ�����.Item(str��Ϣ��) = str��Ϣֵ
                Else
                    mdicҽ�ƿ�����.Add str��Ϣ��, str��Ϣֵ
                End If
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    
    LoadPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdSelDrug_Click()
    '�����:56599
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select ID,nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
        " NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ��" & _
        " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����," & _
        " A.����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"

    '��ȡ��ǰ�������ֵ
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (vsDrug.Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "����ҩ��ѡ����", "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ��", False, False, True, vRect.Left, vRect.Top, 0, True, False, True)

    If Not rsTemp Is Nothing Then
        vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) = rsTemp!����
        vsDrug.TextMatrix(vsDrug.Row, 2) = rsTemp!id
        If vsDrug.Rows - 1 = vsDrug.Row Then vsDrug.Rows = vsDrug.Rows + 1
    End If
    If vsDrug.Visible = True And vsDrug.Enabled = True Then vsDrug.SetFocus
    Exit Sub
ErrHandl:
    MsgBox Err.Description
End Sub

Private Sub cmd��ֵ_Click()
    '�����:54208
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State <> 0 Then
            Call zlPrepayFunc(1, mrsInfo!����ID)
        End If
    Else
        Call zlPrepayFunc(1, 0)
    End If
End Sub

Private Sub cmd���ڵ�ַ_Click()
    Call SearchAddress("", txt���ڵ�ַ)
End Sub

Private Sub cmd����˿�_Click()
    '�����:50891
    Call zlPrepayFunc(2, mlng����ID)
End Sub
Private Function zlPrepayFunc(ByVal intFunc As Integer, ByVal lng����ID As Long, Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ԥ���
    '���:intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
    '����:���˺�
    '����:2011-07-24 18:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun As Object, intԤ������ As Integer
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Function
    'bytԤ������: 0-��Ԥ����(ȱʡ,���л�����),1-�������(1),2-����״̬(1); 3-����˿�(37770), 4-����תסԺ;5-סԺת����
    Select Case intFunc
    Case 1  '1.��Ԥ��
        intԤ������ = 0
    Case 2 '�˿�
        intԤ������ = 3
    Case 3: intԤ������ = 2
    Case 4: intԤ������ = 4
    Case 5: intԤ������ = 5
    End Select
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ� ����Ԥ�����տ��
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ���ݿ��¼�û���
    '  bytCallObject:���˺����(0-Ԥ�������(ȱʡ��);1-���˷��ò�ѯ����,2-ҽ�ƿ�����)
    '  lng����ID-ȱʡ�Ĳ���ID
    '  lng��ҳID-ȱʡ����ҳID
    '  dblDefPrePayMoney-ȱʡ��Ԥ�����
    Set gfrmCardMgr = Me
    '����:48249
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng����ID, 0, 0, intԤ������, blnOneCard) = False Then
        zlPrepayFunc = False
        Set gfrmCardMgr = Nothing
        Exit Function
    End If
    Set gfrmCardMgr = Nothing
    zlPrepayFunc = True
End Function
Private Sub cmd�����ص�_Click()
    If Select����(txt�����ص�, lbl�����ص�, "") = False Then Exit Sub
End Sub
Private Sub cmd��ͬ��λ_Click()
    If Select��Լ��λ("") = False Then Exit Sub
End Sub

Private Sub cmd��ͥ��ַ_Click()
    If Select����(txt��ͥ��ַ, lbl��ͥ��ַ, "") = False Then Exit Sub
End Sub

Private Sub cmd��ϵ�˵�ַ_Click()
    If Select����(txt��ϵ�˵�ַ, lbl��ϵ�˵�ַ, "") = False Then Exit Sub
End Sub

Private Sub cmd����_Click()
    If Select����("") = False Then Exit Sub
End Sub

Private Sub SetȨ��()
    Dim strValue As String
    'Ʊ���ϸ����
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mblnBillԤ�� = Mid(strValue, 2, 1) = "1"
    
    'Ʊ�ݺ��볤�ȡ����￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    mbytԤ�� = Val(Split(strValue, "|")(1))
    
    cmd����˿�.Visible = zlCheckPrivs(mstrPrepayPrivs, "Ԥ���˿�")
    cmd��ֵ.Visible = zlCheckPrivs(mstrPrepayPrivs, "Ԥ���˿�")
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    
    Call LoadCardFee: Call SetCtrlMove
    Call SetControlEnable
    Call SetCardEditEnabled
    '�޸���:56599
    Call InitTabPage
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    Call InitTaskPanelOther
    If mstrCardNo <> "" Then
        If mEditType = Cr_��ѯ Then
            mint��¼״̬ = 1
            Call ReadCardNo(mstrCardNo, 2)
        Else
            Call ReReadCard(mstrCardNo)
        End If
    End If
    
    If mlng����ID <> 0 Then
        If GetPatient("-" & mlng����ID) Then
            Call LoadPatiInfor: Call zlQueryEMPIPatiInfo
        End If
        If mEditType = Cr_��ʧ Then
            txtˢ������.Text = mstrCardNo
            If txtˢ������.Text = "" Then
                If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
            Else
                If cbo��ʧ��ʽ.Enabled And cbo��ʧ��ʽ.Visible Then cbo��ʧ��ʽ.SetFocus
            End If
        End If
    Else
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    If mEditType = Cr_���� Then
         If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
    End If
    If mEditType = Cr_�˿� Then
        '����:47772
         chkCancel.value = 1
        '����:48249
         mblnNotClick = True
         '0-������ˢ��;1-ˢ���˿�;2-���ݺź�����֤ˢ��;3-1��2�Ĺ���ģʽ
         Select Case mParaData.int�˿�ģʽ
         Case 0
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
         Case 1
             If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
         Case 2
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
         Case Else
             If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
         End Select
        mblnNotClick = False
    End If
    wndTaskPanel.Reposition
    mblnChange = False
    
       '�����:56599
'    If mEditType <> Cr_�󶨿� And mEditType <> Cr_���� Then
'        NotVisibleImage
'    End If
End Sub
Private Sub BackCardReadCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˿�����
    '����:���˺�
    '����:2011-12-25 14:04:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutPut As String, strExpand As String, strOutXml As String, strCardNo As String
    If Not (mEditType = Cr_�˿� Or chkCancel.value = 1) Then Exit Sub
    If mCardType.bln���￨ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtˢ������.Text = mobjICCard.Read_Card()
            If Trim(txtˢ������.Text) = "" Then Exit Sub
            If ReadCardNo(Trim(txtˢ������.Text), 2) = -1 Then
                If mEditType <> Cr_���� Then
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
                End If
            Else
                Call zlControl.TxtSelAll(txtˢ������)
                stbThis.Panels(2) = "û�з��ָ�" & mCardType.str������ & "����Ϣ,����δ����,���飡"
                txtˢ������.Text = ""
                Exit Sub
            End If
        End If
        Exit Sub
    End If
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Sub
    End If
    If mobjReadCard Is Nothing Then Exit Sub
    If mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txtˢ������.Text = strCardNo
    If Trim(txtˢ������.Text) = "" Then Exit Sub
    If ReadCardNo(Trim(txtˢ������.Text), 2) = -1 Then
        '�ɹ�
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
    Else
        Call zlControl.TxtSelAll(txtˢ������)
        stbThis.Panels(2) = "û�з��ָ�" & mCardType.str������ & "����Ϣ,���飡"
        txtˢ������.Text = ""
        Exit Sub
    End If
End Sub

Private Function ReReadCard(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¶���
    '����:���˺�
    '����:2011-09-08 22:20:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPhotoFile As String
    Dim strOutPut As String, strExpand As String, strOutXml As String
    '����:48249
    If (mEditType = Cr_�˿� Or chkCancel.value = 1) And strCardNo = "" Then
        '�˿�����
        Call BackCardReadCard: Exit Function
    End If
    '�����:57962
    If mEditType = Cr_���� Then
        txtˢ������.Text = strCardNo '����ʱ�����Text����������ԭ���������
    End If
    
    '����:47914
    '����:48079
    If Not (mEditType = Cr_���� Or mEditType = Cr_�󶨿� _
                                Or (mEditType = Cr_���� And Not mCardType.blnˢ��) _
                                Or (mEditType = Cr_���� And Not mCardType.blnˢ��) _
                                    ) Then Exit Function
   ' If mCardType.bln���ƿ� Then Exit Sub
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Function
    End If
    
    If mobjReadCard Is Nothing Then Exit Function
    strExpand = mlngCardTypeID
    On Error Resume Next
    ReReadCard = mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml, strPhotoFile)
    If Err <> 0 Then
        If Err <> 450 Then GoTo errHandle:
        '450-����Ĳ����Ż���Ч�����Ը�ֵ
        '��Ҫ��Ǹ����ǰ��
         Err = 0: On Error GoTo errHandle
         ReReadCard = mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml)
    End If
    If Not ReReadCard Then Exit Function
    
    txt����.Text = Trim(strCardNo)
    If txt����.Text <> "" Then
        Call CheckFreeCard(txt����.Text)
        '����:62821
        If strPhotoFile <> "" Then imgPatient.Picture = LoadPicture(strPhotoFile)
        Call LoadPati(strOutXml)
    End If
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is txt���� Or Me.ActiveControl Is txtAudi Or Me.ActiveControl Is txtPass Then Exit Sub
        If Me.ActiveControl Is txtˢ������ Then Exit Sub
        If Me.ActiveControl Is txt��ϵ�˵�ַ Then Exit Sub
        If Me.ActiveControl Is txt���� Then Exit Sub
        If Me.ActiveControl Is txt��ͥ��ַ Then Exit Sub
        If Me.ActiveControl Is txt������λ Then Exit Sub
        If Me.ActiveControl Is txt�����ص� Then Exit Sub
        If Me.ActiveControl Is txt���� Then Exit Sub
        '76609,Ƚ����,2014-8-14,���㶨λ����
        If Me.ActiveControl Is txtPatient Then Exit Sub
        '78408:���ϴ�,2014/10/9,�����ת
        If Me.ActiveControl Is vsDrug Then Exit Sub
        If Me.ActiveControl Is vsInoculate Then Exit Sub
        '89242:���ϴ�,2015/12/3,PatiAddress�ؼ��ڲ���������ת���ⲿ���ٴ���
        If UCase(TypeName(Me.ActiveControl)) = UCase("PatiAddress") Then Exit Sub
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyE
        If Shift = vbCtrlMask Then
            If wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded Then
                wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = False
            Else
                wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = True
            End If
        End If
    Case vbKeyF2
        If cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus: Call cmdOK_Click
        End If
    Case vbKeyF6
        If txtPatient.Enabled And txtPatient.Visible Then
            txtPatient.SetFocus
        End If
    Case vbKeyF8
        If mEditType = Cr_���� Then
            chkCancel.value = 1
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        End If
    Case vbKeyF12
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    Case vbKeyEscape
        If cmdCancel.Enabled And cmdCancel.Visible Then
            cmdCancel.SetFocus: Call cmdCancel_Click
        End If
    Case Else
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub
Private Sub Form_Load()
    Dim intKind As Integer, strKey As String
    mstrCboSplit = "-" & Chr(30)
    mblnFirst = True
    mstrPrepayPrivs = ";" & GetPrivFunc(glngSys, 1103) & ";"
    mstrTitle = "���˷�������"
    RestoreWinState Me, App.ProductName, mstrTitle
    Call CreateObjectPlugIn '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    Call CreateObjectKeyboard
    '69026,Ƚ����,2014-8-8,�����������
    If CreatePublicPatient = False Then
        mblnUnLoad = True: Exit Sub
    End If
    
    Call InitFace
    Call InitTaskPancel '��ʼ�����
    Call SetControlVisitble: Call SetȨ��
    
    HookDefend txtPass.hWnd
    HookDefend txtAudi.hWnd
    HookDefend txtԭ������.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strName As String
    
    '115193:���ϴ�,2017/10/13,ж�ش���ʱ�����ģ�����
    '�����:56599
    strName = IIf(glngSys \ 100 = 8, "�ͻ��Ļ�Ա��", "���˵�ҽ�ƿ�")
    If Not mblnUnLoad Then
        If mEditType = Cr_��ѯ Then
        ElseIf chkCancel.value = Checked Then
            If mblnChange Then
                If MsgBox("ȷ��Ҫ�����˿�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
            End If
        ElseIf Not mrsInfo Is Nothing And (mEditType = Cr_���� Or mEditType = Cr_�󶨿�) Then
            If mrsInfo.State = adStateOpen Then
                If MsgBox("��" & strName & "��δ����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
            End If
        End If
        If mblnChange Then
             If MsgBox("��Ƭ��Ϣ�Ѿ������ı䣬���㻹δȷ�ϣ��Ƿ����Ҫ�˳���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
        End If
    End If
    
    mlngͼ����� = 0: mstr�ɼ�ͼƬ = "": Set mdicҽ�ƿ����� = Nothing
    If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", IDKind.IDKind)
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled False
        Set mobjICCard = Nothing
    End If
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    
    Set mobjReadCard = Nothing
    
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing Then Set mobjPlugIn = Nothing
    mlngPlugInHwnd = 0: mblnPlugin = False
    
    zlDatabase.SetPara "��ʾ��չ��Ϣ", IIf(mParaData.blnShowExpend, 1, 0), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    If mEditType = Cr_���� Or mEditType = Cr_�󶨿� Then
        '�������
        zlDatabase.SetPara "�ϴη������", mCardType.lng�����ID, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    SaveWinState Me, App.ProductName, mstrTitle
    Call UnHookKBD
    mblnGetBirth = False
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String
    If IsCardType(IDKind, "IC����") Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    lng�����ID = IDKind.GetCurCard.�ӿ����
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
    'Call InitInterFacel(Me, mlngModule, lng�����ID, False, mobjCardObject)
    strExpand = lng�����ID
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNo
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
    Exit Sub
 
End Sub

 
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
     '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    Set gobjSquare.objCurCard = objCard
    mlngҽ�ƿ����� = objCard.���ų���
    '105667:���ϴ���2017/5/23�����ż��ܵ��µ�һ������ƴ�����ܴ������뷨
    txtPatient.PasswordChar = ""
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Then Exit Sub  'Or Not Me.ActiveControl Is txtPatient Or txtPatient.Text <> ""
    mblnNotClick = True
    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtPatient.Text = objPatiInfor.����
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub IDKindPayMode_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnNotChange Then Exit Sub
    mblnNotChange = True
    If Val(txt�ϼ�.Text) - Val(txt�ϼ�.Tag) < 0 Then
        IDKindPayMode.IDKind = 1 'Ϊ����ʱ���ܳ�ֵ
    ElseIf cbo֧����ʽ.ListIndex >= 0 Then
        If cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) = -1 Then IDKindPayMode.IDKind = 2  '�����������ѿ������Ҳ�
    End If
    mblnNotChange = False
End Sub

Private Sub lblˢ����֤_Click()
    Dim strOutPut As String, strExpand As String, strOutXml As String, strCardNo As String
    If mCardType.bln���￨ = False Then Exit Sub
    If Not (mEditType = Cr_�˿� Or chkCancel.value = 1) Then Exit Sub
    If mCardType.bln���￨ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtˢ������.Text = mobjICCard.Read_Card()
            If Trim(txtˢ������.Text) = "" Then Exit Sub
            If ReadCardNo(Trim(txtˢ������.Text), 2) = -1 Then
                If mEditType <> Cr_���� Then
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
                End If
            Else
                Call zlControl.TxtSelAll(txtˢ������)
                stbThis.Panels(2) = "û�з��ָ�" & mCardType.str������ & "����Ϣ,����δ����,���飡"
                txtˢ������.Text = ""
                Exit Sub
            End If
        End If
        Exit Sub
    End If
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Sub
    End If
    If mobjReadCard Is Nothing Then Exit Sub
    If mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txtˢ������.Text = strCardNo
    If Trim(txtˢ������.Text) = "" Then Exit Sub
    If ReadCardNo(Trim(txtˢ������.Text), 2) = -1 Then
        If mEditType <> Cr_���� Then
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
        End If
    Else
        Call zlControl.TxtSelAll(txtˢ������)
        stbThis.Panels(2) = "û�з��ָ�" & mCardType.str������ & "����Ϣ,����δ����,���飡"
        txtˢ������.Text = ""
        Exit Sub
    End If
End Sub

Private Sub picCard_Resize()
    Err = 0: On Error Resume Next
    With picCard
        If mEditType = Cr_�󶨿� Or mEditType = Cr_���� Then
            tbPageDo.Move 0, 0, .ScaleWidth, .ScaleHeight
            fraCard.Move 0, 0, tbPageDo.ScaleWidth, tbPageDo.ScaleHeight
        Else
            fraCard.Move 0, 0, .ScaleWidth, .ScaleHeight
        End If
    End With
End Sub
Private Sub picDrugAllergy_Resize()
'�����:56599
    vsDrug.Left = picDrugAllergy.Left - 80
    vsDrug.Top = picDrugAllergy.Top - 380
    vsDrug.Height = picDrugAllergy.ScaleHeight
    vsDrug.Width = picDrugAllergy.ScaleWidth
End Sub

Private Sub picExpend_Resize()
'�޸���:56599
Err = 0: On Error Resume Next
    With picExpend
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ҳ�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Dim intEditType As Integer '���봰��ʱ�Ĳ�������
    
    Err = 0: On Error GoTo Errhand:
    If mEditType <> Cr_����������Ϣ Then
        Set objItem = tbPage.InsertItem(mPageIndex.����, "����", fraBase.hWnd, 0)
        objItem.Tag = mPageIndex.����
        
        If mEditType = Cr_�󶨿� Or mEditType = Cr_���� Then
            Set objItem = tbPage.InsertItem(mPageIndex.ҩ�����, "ҩ�����", picDrugAllergy.hWnd, 0)
            objItem.Tag = mPageIndex.ҩ�����
            Call InitvsDrug
            
            Set objItem = tbPage.InsertItem(mPageIndex.������Ϣ, "������Ϣ", picInoculate.hWnd, 0)
            objItem.Tag = mPageIndex.������Ϣ
            Call InitVsInoculate
            
            Set objItem = tbPage.InsertItem(mPageIndex.������Ϣ, "������Ϣ", picOtherInfo.hWnd, 0)
            objItem.Tag = mPageIndex.������Ϣ
            Call InitVsOtherInfo
            Call InitCombox
            
            '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
            If Not mobjPlugIn Is Nothing Then
                On Error Resume Next
                mlngPlugInHwnd = mobjPlugIn.GetFormHwnd
                Call zlPlugInErrH(Err, "GetFormHwnd")
                Err.Clear: On Error GoTo 0
                If mlngPlugInHwnd <> 0 Then
                    picTaskPanelOther.Visible = True
                    Set objItem = tbPage.InsertItem(mPageIndex.������Ϣ, "������Ϣ", picTaskPanelOther.hWnd, 0)
                    objItem.Tag = mPageIndex.������Ϣ
                End If
            End If
        Else
            picDrugAllergy.Visible = False
            picInoculate.Visible = False
            picOtherInfo.Visible = False
        End If
         
         With tbPage
            tbPage.Item(0).Selected = True
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        
        '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
        If mEditType = Cr_�󶨿� Or mEditType = Cr_���� Then
            intEditType = mEditType '��¼�������ͣ���ֹ����ҳ��ʱ������
            Set objItem = tbPageDo.InsertItem(0, "����", fraCard.hWnd, 0): objItem.Tag = Cr_����
            Set objItem = tbPageDo.InsertItem(1, "�󶨿�", fraCard.hWnd, 0): objItem.Tag = Cr_�󶨿�
            If intEditType = Cr_�󶨿� Then
                tbPageDo(1).Selected = True
            Else
                tbPageDo(1).Selected = True: tbPageDo(0).Selected = True
            End If
            With tbPageDo
                Call SetCardPayOrBound
                .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
                .PaintManager.BoldSelected = True
                .PaintManager.Layout = xtpTabLayoutAutoSize
                .PaintManager.StaticFrame = True
                .PaintManager.ClientFrame = xtpTabFrameSingleLine
            End With
        End If
    Else
        picDrugAllergy.Visible = False
        picInoculate.Visible = False
        picOtherInfo.Visible = False
        tbPage.Visible = False
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub
Private Sub picInoculate_Resize()
'�����:56599
    vsInoculate.Left = picInoculate.Left - 80
    vsInoculate.Top = picInoculate.Top - 380
    vsInoculate.Height = picInoculate.ScaleHeight
    vsInoculate.Width = picInoculate.ScaleWidth
End Sub

Private Sub picTaskPanelOther_Resize()
    wndTaskPanelOther.Move 0, 0, picTaskPanelOther.Width, picTaskPanelOther.Height
End Sub

Private Sub txtAudi_Change()
    mblnChange = True
End Sub

Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    zlCommFun.OpenIme False
    Call OpenPassKeyboard(txtAudi, True)
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)

    Call CheckInputPassWord(KeyAscii, mCardType.int������� = 1)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0

    If Not txt����.Locked And txt����.TabStop And txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Sub
    If chk����.Visible And chk����.Enabled Then chk����.SetFocus: Exit Sub
    If cbo֧����ʽ.Visible And cbo֧����ʽ.Enabled Then cbo֧����ʽ.SetFocus: Exit Sub
    Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub
Private Sub txtAudi_LostFocus()

    Call ClosePassKeyboard(txtAudi)
End Sub

Private Sub txtAudi_Validate(Cancel As Boolean)
    If txtPass.Text <> txtAudi.Text And txtAudi.Text <> "" Then
        MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
        Cancel = 1
        Call zlControl.TxtSelAll(txtAudi)
        If txtAudi.Enabled And txtAudi.Visible Then txtAudi.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtPass_Change()
    mblnChange = True
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    zlCommFun.OpenIme False
    txtPass.MaxLength = 0
    '108779:���ϴ�,2017/5/8,�������ƹ���ΪNλ����ʱ�����ܳ������볤��
    Select Case mCardType.int���볤������
        Case 0
        Case Else
            txtPass.MaxLength = mCardType.int���볤��
            txtAudi.MaxLength = mCardType.int���볤��
    End Select
    Call OpenPassKeyboard(txtPass, False)
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCardType.int������� = 1)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    If Not (txtPass.Text = "" And txtAudi.Text = "") Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If Not txt����.Locked And txt����.TabStop And txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Sub
    If chk����.Visible And chk����.Enabled Then chk����.SetFocus: Exit Sub
    If cbo֧����ʽ.Visible And cbo֧����ʽ.Enabled Then cbo֧����ʽ.SetFocus: Exit Sub
    Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub
Private Sub txtPass_LostFocus()
    Call ClosePassKeyboard(txtPass)
End Sub

Private Sub txtPatient_Change()
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    Call AutoBrushSet(IDKind, txtPatient.Text = "")
    If mblnNotCharge Then Exit Sub
    mblnChange = True
End Sub

Private Sub txt�����ص�_Change()
    mblnChange = True: lbl�����ص�.Tag = ""
End Sub

Private Sub txt�����ص�_GotFocus()
    zlControl.TxtSelAll txt�����ص�
    zlCommFun.OpenIme True
End Sub

Private Sub txt�����ص�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl�����ص�.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt�����ص�) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select����(txt�����ص�, lbl�����ص�, Trim(txt�����ص�)) = False Then Exit Sub
End Sub

Private Sub txt�����ص�_LostFocus()
      zlCommFun.OpenIme False
End Sub

Private Sub txt��������_Change()
    Dim str����ʱ�� As String
    If IsDate(txt��������.Text) And Not mblnNotChange Then
        mblnNotChange = True
        txt��������.Text = Format(CDate(txt��������.Text), "yyyy-mm-dd")
        mblnNotChange = False
        
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        mstr���� = txt����.Text: mstr���䵥λ = cbo���䵥λ.Text
        '111836:���ϴ���2017/7/21������ؼ�λ�ü���
        If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
            cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False: txt����.Width = 1220
        Else
            cbo���䵥λ.Visible = True: txt����.Width = 550
            If cbo���䵥λ.ListIndex = -1 Then cbo���䵥λ.ListIndex = 0
        End If
        mblnGetBirth = False
    End If
End Sub

Private Sub txt��������_LostFocus()
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
        If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
    End If
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt��������) Then
        KeyAscii = 0
        txt����ʱ��.Text = "__:__"
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    Dim str����ʱ�� As String
    '76669�����ϴ�,2014-8-18,�����������
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        If txt����ʱ��.Enabled And txt����ʱ��.Visible Then txt����ʱ��.SetFocus
        Cancel = True
    ElseIf IsDate(txt��������.Text) Then
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        mstr���� = txt����.Text: mstr���䵥λ = cbo���䵥λ.Text
    End If
End Sub

Private Sub txt��λ�绰_Change()
    mblnChange = True
End Sub

Private Sub txt��λ�绰_GotFocus()
    zlControl.TxtSelAll txt��λ�绰
    zlCommFun.OpenIme False
End Sub

Private Sub txt��λ�绰_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt��λ������_Change()
    mblnChange = True
End Sub

Private Sub txt��λ������_GotFocus()
    zlControl.TxtSelAll txt��λ������
    zlCommFun.OpenIme True
End Sub
Private Sub txt��λ������_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt��λ�ʱ�_Change()
    mblnChange = True
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��λ�ʱ�
    zlCommFun.OpenIme False
End Sub

Private Sub txt��λ�ʻ�_Change()
    mblnChange = True
End Sub

Private Sub txt��λ�ʻ�_GotFocus()
    zlControl.TxtSelAll txt��λ�ʻ�
    zlCommFun.OpenIme False
End Sub

Private Sub txt������λ_Change()
    mblnChange = True: lbl������λ.Tag = ""
End Sub

Private Sub txt������λ_GotFocus()
    zlControl.TxtSelAll txt������λ
    zlCommFun.OpenIme True
End Sub

Private Sub txt������λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl������λ.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt������λ) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select��Լ��λ(Trim(txt������λ.Text)) = False Then Exit Sub
End Sub

Private Sub txt������λ_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt�ϼ�_Change()
    Call txt���_Change
End Sub

Private Sub txt�ϼ�_GotFocus()
    zlControl.TxtSelAll txt�ϼ�
    zlCommFun.OpenIme False
End Sub
Private Sub txt�ϼ�_KeyPress(KeyAscii As Integer)
    If txt�ϼ�.Locked Or txt�ϼ�.Enabled = False Then Exit Sub
    zlControl.TxtCheckKeyPress txt�ϼ�, KeyAscii, m���ʽ
End Sub

Private Sub txt�ϼ�_Validate(Cancel As Boolean)
    txt�ϼ�.Text = Format(txt�ϼ�.Text, "0.00")
End Sub

Private Sub txt���ڵ�ַ_Change()
    mblnChange = True
    txt���ڵ�ַ.Tag = ""
End Sub

Private Sub txt���ڵ�ַ_GotFocus()
    Call zlControl.TxtSelAll(txt���ڵ�ַ)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt���ڵ�ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Trim(txt���ڵ�ַ.Text) <> "" Then
        Call SearchAddress(Trim(txt���ڵ�ַ.Text), txt���ڵ�ַ)
    End If
End Sub

Private Sub txt���ڵ�ַ_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub SearchAddress(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '����:ģ�����ң���������ѡ���б�
    '����:Ƚ����
    '����:2014-5-23
    '����:
    '   strInput:�����ı�����Ϊ�ձ�ʾ�����ť����
    '   txtInput:�ı������
    '--------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
    vRect = GetControlRect(txtInput.hWnd)
    If strInput = "" Then '�����ť
        strSQL = "" & _
            "Select ID, �ϼ�id, ����, ����, ĩ�� " & _
            "From (With ����_t As" & _
            "    (Select Rownum As �к�, ID, �ϼ�id, ĩ��, ����, ����" & _
            "     From (Select Distinct Substr(����, 1, 2) As ID, Null As �ϼ�id, 0 As ĩ��, Null As ����, Substr(����, 1, 2) As ����" & _
            "            From ����" & _
            "            Union All" & _
            "            Select ���� As ID, Substr(����, 1, 2) As �ϼ�id, 1 As ĩ��, ����, ���� From ����))" & _
            "   Select �к� As ID, To_Number(�ϼ�id) As �ϼ�id, ����, ����, ĩ�� From ����_t Where �ϼ�id Is Null" & _
            "   Union All" & _
            "   Select b.�к�, a.�к�, b.����, b.����, b.ĩ�� From ����_t A, ����_t B Where a.Id = b.�ϼ�id Order By ����)"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        'ȥ��"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '����ȫ������ʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then '����ȫ����ĸʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            Else
                strWhere = " Where ���� Like Upper([1]) Or ���� Like [1] Or ���� Like Upper([1])"
            End If
        End If
        
        strSQL = "" & _
            "Select Rownum As ID, ����, ���� From ���� " & strWhere & " Order By ����"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, _
                       "", "", False, False, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False, strKey)
    End If
    If blnCancel Then txtInput.SetFocus: Exit Sub

    If rsTemp Is Nothing Then txtInput.SetFocus: Exit Sub
    If rsTemp.State <> 1 Then txtInput.SetFocus: Exit Sub
    
    txtInput.Text = Nvl(rsTemp!����)
    txtInput.Tag = Nvl(rsTemp!id)
    txtInput.SelStart = Len(txtInput.Text)
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt���ڵ�ַ�ʱ�_Change()
    mblnChange = True
End Sub

Private Sub txt���ڵ�ַ�ʱ�_GotFocus()
    Call zlControl.TxtSelAll(txt���ڵ�ַ�ʱ�)
End Sub

Private Sub txt���ڵ�ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt��ͥ�ʱ�_Change()
    mblnChange = True
End Sub

Private Sub txt��ͥ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��ͥ�ʱ�
    zlCommFun.OpenIme False
End Sub

Private Sub txt��ͥ��ַ_Change()
    mblnChange = True
    lbl��ͥ��ַ.Tag = ""
End Sub

Private Sub txt��ͥ��ַ_GotFocus()
    zlControl.TxtSelAll txt��ͥ��ַ
    zlCommFun.OpenIme True
End Sub

Private Sub txt��ͥ��ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl��ͥ��ַ.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt��ͥ��ַ) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select����(txt��ͥ��ַ, lbl��ͥ��ַ, Trim(txt��ͥ��ַ)) = False Then Exit Sub
End Sub
 

Private Sub txt��ͥ��ַ_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt��ͥ�绰_Change()
    mblnChange = True
End Sub

Private Sub txt��ͥ�绰_GotFocus()
    zlControl.TxtSelAll txt��ͥ�绰
    zlCommFun.OpenIme False
End Sub

Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    zlCommFun.OpenIme False
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
    If txt����.Locked Then Exit Sub
    zlControl.TxtCheckKeyPress txt����, KeyAscii, m���ʽ
    If KeyAscii <> vbKeyReturn Then Exit Sub
    KeyAscii = 0
    If mCardType.bln��� Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If mCardType.rsҽ�ƿ��� Is Nothing Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If mCardType.rsҽ�ƿ���!�ּ� <> 0 And Abs(CCur(txt����.Text)) > Abs(mCardType.rsҽ�ƿ���!�ּ�) Then
        MsgBox mCardType.str������ & "������ֵ���ܴ�������޼ۣ�" & Format(Abs(mCardType.rsҽ�ƿ���!�ּ�), "0.00"), vbExclamation, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
    End If
    If mCardType.rsҽ�ƿ���!ԭ�� <> 0 And Abs(CCur(txt����.Text)) < Abs(mCardType.rsҽ�ƿ���!ԭ��) Then
        MsgBox mCardType.str������ & "��������ֵ����С������޼ۣ�" & Format(Abs(mCardType.rsҽ�ƿ���!ԭ��), "0.00"), vbExclamation, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
    End If
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Call txt���_Change
End Sub

Private Sub txt����_Change()
    Dim rsTemp As Recordset

    mblnChange = True
    Call SetCardEditEnabled
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    Call AutoBrushSet(IDKindPay, txt����.Text = "")
    '�����:53408
    If mCardType.str������ = "�������֤" Then
        Call OpenIDCard(txt����.Text = "")
        If Len(txt����.Text) = mCardType.lng���ų��� Then
            Set rsTemp = zl�Ƿ��Ѱ�(Trim(txt����.Text))
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount <= 0 Then Exit Sub
            If MsgBox("����Ϊ:" & txt���֤��.Text & "�Ѿ�������:" & rsTemp!���� & "��,�Ƿ�Ҫȡ���Ѱ󶨵����֤��", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
                frmPaticurCardCancelBound.zlCancelBand Me, mlngModule, mlngCardTypeID, rsTemp!����ID, txt����.Text, False
            End If
        End If

    End If
End Sub

Private Sub txt����_GotFocus()
    '76609,Ƚ����,2014-8-14,ˢ����ˢ��ĩβ���ܴ����лس������㶨λ����
    mblnTab = False
    If Not txt����.Enabled Then Exit Sub
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    Call AutoBrushSet(IDKindPay, txt����.Text = "")
    zlControl.TxtSelAll txt����
    zlCommFun.OpenIme False
    '�����:53408
    If mCardType.str������ = "�������֤" Then
        Call OpenIDCard(txt����.Text = "")
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    '�����:53408
    If mCardType.str������ = "�������֤" Or mCardType.str������ = "IC��" Then
        KeyAscii = 0
    End If

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> 13 Then
        If Len(txt����.Text) = mCardType.lng���ų��� - 1 And KeyAscii <> 8 Then
            '76609,Ƚ����,2014-8-14,ˢ����ˢ��ĩβ���ܴ����лس������㶨λ����
            mblnTab = True
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Call EnableKBDHook
        End If
    ElseIf txt����.Text = "" Then
        KeyAscii = 0: cmdOK.SetFocus  '������,ֱ������
    Else
        KeyAscii = 0: If Not mblnTab Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����֤������
    '����:����
    '����:2012-08-31 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (blnEnabled)
End Sub

Private Sub txt����_LostFocus()
    '76609,Ƚ����,2014-8-14,ˢ����ˢ��ĩβ���ܴ����лس������㶨λ����
    mblnTab = False
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    '97702,���ϴ�,2016/6/28,�����Ƴ���ر��Զ�����
    Call AutoBrushSet(IDKindPay, False)
    Call zlCommFun.OpenIme(False)
    If mCardType.str������ = "�������֤" Then
        Call OpenIDCard(False)
    End If
    Call ReLoadCardFee(True)
End Sub

Private Sub ReLoadCardFee(Optional ByVal blnFeedName As Boolean)
    'blnFeedName-�Ƿ���������飬���ٽ��������޸�������Ϣ�����ĵ���
    '�뿪��鿨��
    Dim lng����ID As Long, lng�շ�ϸĿID As Long
    Dim strSQL As String, str���� As String
    Dim rsTmp As ADODB.Recordset
    
    If (mEditType <> Cr_���� And mEditType <> Cr_�󶨿� And mEditType <> Cr_����) Or chkCancel.value = 1 Then Exit Sub
    If mCardType.rsҽ�ƿ��� Is Nothing Then Exit Sub
    If mCardType.rsҽ�ƿ���.RecordCount = 0 Then Exit Sub
    If mCardType.lng�����ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt����.Text) = "" Then Exit Sub
    If Trim(txt����.Text) = "" Then Exit Sub
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = mrsInfo!����ID
    End If
    If blnFeedName = False And lng����ID <> 0 Then Exit Sub
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    mCardType.rsҽ�ƿ���.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as �շ�ϸĿID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����", mlngModule, mCardType.lng�����ID, Trim(txt����.Text), lng����ID, _
                Trim(txtPatient.Text), NeedName(cbo�Ա�.Text), str����, txt���֤��.Text, Val(Nvl(mCardType.rsҽ�ƿ���!�շ�ϸĿID)))
    If rsTmp.EOF Then Exit Sub
    
    lng�շ�ϸĿID = Val(Nvl(rsTmp!�շ�ϸĿID))
    Set rsTmp = zlGetSpecialItemFee("", lng�շ�ϸĿID)
    If Not rsTmp Is Nothing Then Set mCardType.rsҽ�ƿ��� = rsTmp
    Call LoadCardFee
End Sub


Private Sub txt��ϵ�˵�ַ_Change()
    mblnChange = True
End Sub

Private Sub txt��ϵ�˵�ַ_GotFocus()
    zlControl.TxtSelAll txt��ϵ�˵�ַ
    zlCommFun.OpenIme True
End Sub
 

Private Sub txt��ϵ�˵�ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl��ϵ�˵�ַ.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt��ϵ�˵�ַ) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select����(txt��ϵ�˵�ַ, lbl��ϵ�˵�ַ, Trim(txt��ϵ�˵�ַ)) = False Then Exit Sub
End Sub

Private Sub txt��ϵ�˵�ַ_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt��ϵ�˵绰_Change()
    mblnChange = True
End Sub

Private Sub txt��ϵ�˵绰_GotFocus()
    zlControl.TxtSelAll txt��ϵ�˵绰
    zlCommFun.OpenIme False
End Sub

Private Sub txt��ϵ�����֤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt��ϵ������_Change()
    mblnChange = True
End Sub

Private Sub txt��ϵ������_GotFocus()
    zlControl.TxtSelAll txt��ϵ������
    zlCommFun.OpenIme True
End Sub

Private Sub txt��ϵ������_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt�����_Change()
    mblnChange = True
End Sub

Private Sub txt�����_GotFocus()
    '94941:���ϴ�,2016/4/7,�޸������Ȩ��
    If InStr(";" & mstrPrivs & ";", ";�����޸������;") > 0 Then
        zlControl.TxtSelAll txt�����
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    '94941:���ϴ�,2016/4/7,�޸������Ȩ��
    If KeyAscii = vbKeySpace Then
        txt�����.Text = zlGet�����: KeyAscii = 0: Exit Sub
    End If
    If InStr(";" & mstrPrivs & ";", ";�����޸������;") <= 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt�����, KeyAscii, m����ʽ
End Sub
Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    Call zlCommFun.OpenIme
    zlControl.TxtSelAll txt����
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            Call cbo���䵥λ.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) And cbo���䵥λ.Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    '111836:���ϴ���2017/7/21�����䷴��
    Dim strBirth As String
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False: txt����.Width = 1220
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.Visible = True: txt����.Width = 550
        If cbo���䵥λ.ListIndex = -1 Then cbo���䵥λ.ListIndex = 0
    End If
    '69026,Ƚ����,2014-8-8,�����������
    '76703,Ƚ����,2014-8-15
    If mobjPubPatient Is Nothing Then Exit Sub
    If txt����.Text <> mstr���� Then
        mstr���� = txt����.Text
        If Not IsDate(txt��������.Text) Then mblnGetBirth = True
        If cbo���䵥λ.Visible Then mstr���䵥λ = "": Exit Sub
        mblnNotChange = True
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), strBirth) Then
                txt��������.Text = Format(strBirth, "yyyy-mm-dd")
                txt����ʱ��.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnNotChange = False
    End If
    If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), _
            IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
            IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then Cancel = True: Exit Sub

End Sub
Private Sub txt����֤��_Change()
    mblnChange = True
End Sub
Private Sub txt����֤��_GotFocus()
    zlControl.TxtSelAll txt����֤��
    zlCommFun.OpenIme False
End Sub

Private Sub txt����_Change()
    mblnChange = True: lbl����.Tag = ""
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    zlCommFun.OpenIme True
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl����.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt����) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select����(Trim(txt����)) = False Then Exit Sub
End Sub

Private Sub txt����_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt���֤��_Change()
    Dim strDate As String
    mblnChange = True
    '�����ܸ��Ĳ��˻�����Ϣʱ,�������ڲ��ܷ���67184
    If Not mblnNotChange And txt��������.Enabled Then
        strDate = zlCommFun.GetIDCardDate(txt���֤��.Text)
        If IsDate(strDate) Then txt��������.Text = strDate
    End If
End Sub
Private Sub txt���֤��_GotFocus()
    zlControl.TxtSelAll txt���֤��
    zlCommFun.OpenIme False
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt�ֻ�_Change()
    mblnChange = True
End Sub

Private Sub txt�ֻ�_GotFocus()
    zlControl.TxtSelAll txt�ֻ�
    zlCommFun.OpenIme False
End Sub

Private Sub txt�ֻ�_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt�ֻ�, KeyAscii, m����ʽ)
End Sub

Private Sub txt�ֻ�_Validate(Cancel As Boolean)
    
    If CheckMobile(txt�ֻ�.Text) = False Then Cancel = True
End Sub

Private Sub txtˢ������_Change()
    lblˢ����֤.Tag = ""
End Sub

Private Sub txtˢ������_GotFocus()
   zlControl.TxtSelAll txtˢ������
End Sub

Private Sub txtˢ������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng����ID As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    lng����ID = Val(Nvl(mrsInfo!����ID))
    If txtˢ������.Text = "" Then
        If zlShowSelectCardNo(lng����ID, "") = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtˢ������_KeyPress(KeyAscii As Integer)
   Dim strCardNo As String
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii <> 13 Then
        If Len(txtˢ������.Text) = mCardType.lng���ų��� - 1 And KeyAscii <> 8 Then
            stbThis.Panels(2) = ""
            txtˢ������.Text = txtˢ������.Text & Chr(KeyAscii)
             strCardNo = Trim(txtˢ������)
             KeyAscii = 0:
             
            If ReadCardNo(strCardNo, 2) = -1 Then
                If mEditType <> Cr_���� Then
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
                End If
            Else
                Call zlControl.TxtSelAll(txtˢ������)
                stbThis.Panels(2) = "û�з��ָþ��￨����Ϣ,����δ����,���飡"
                txtˢ������.Text = ""
                Exit Sub
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        Exit Sub
    End If
    stbThis.Panels(2) = ""
    If lblˢ����֤.Tag = Trim(txtˢ������.Text) Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    strCardNo = Trim(txtˢ������)
    If ReadCardNo(strCardNo, 2) = -1 Then
        If mEditType <> Cr_���� Then
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        End If
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab: Exit Sub
    Else
        If (chkCancel.value = 1 Or mEditType = Cr_�˿�) And mParaData.int�˿�ģʽ = 2 And Trim(cboNO.Text) = "" Then
            Call zlControl.TxtSelAll(cboNO)
           If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Else
            Call zlControl.TxtSelAll(txtˢ������)
        End If
        stbThis.Panels(2) = "û�з��ָþ��￨����Ϣ,����δ����,���飡"
    End If
End Sub

Private Sub txt��֤ҽ����_Validate(Cancel As Boolean)

    txt��֤ҽ����.Text = UCase(Trim(txt��֤ҽ����.Text))
    If cboҽ�Ƹ���.ListCount > 0 And txt��֤ҽ����.Text <> "" Then cboҽ�Ƹ���.ListIndex = 0
    If txt��֤ҽ����.Text <> txtҽ����.Text Then
        MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
        Exit Sub
    End If
    If mInsurePara.lng���ʽҽ������ = 920 And txtҽ����.Text <> lblҽ����(0).Tag And txtҽ����.Text <> "" Then
        If CheckExistsMCNO(txtҽ����.Text) Then
             'Cancel = True
        End If
    End If
End Sub

Private Sub txtҽ����_Change()
    mblnChange = True
End Sub
Private Sub txtҽ����_GotFocus()
    zlControl.TxtSelAll txtҽ����
    zlCommFun.OpenIme False
End Sub

Private Sub txtҽ����_Validate(Cancel As Boolean)
    txtҽ����.Text = UCase(Trim(txtҽ����.Text))
    If cboҽ�Ƹ���.ListCount > 0 And txtҽ����.Text <> "" Then cboҽ�Ƹ���.ListIndex = 0
    If mInsurePara.lng���ʽҽ������ = 920 And txtҽ����.Text <> lblҽ����(0).Tag And txtҽ����.Text <> "" Then
        If CheckExistsMCNO(txtҽ����.Text) Then
             'Cancel = True
        End If
    End If
End Sub

Private Sub txt���_Change()
    If mblnNotChange = True Then Exit Sub
    If chk����.value = Checked Then txt���.Text = "": Exit Sub
    mblnNotChange = True
    txt�ϼ�.Tag = IIf(txt����.Visible, Val(txt����.Text), 0)
    If mEditType = Cr_�˿� Or chkCancel.value = 1 Then txt�ϼ�.Text = Format(txt�ϼ�.Tag, "0.00")
    txt���.Text = Format(Val(txt�ϼ�.Text) - Val(txt�ϼ�.Tag), "0.00")
    
    txt���.ForeColor = IIf(Val(txt���.Text) < 0, vbRed, &H80000008)
    If cbo֧����ʽ.ListIndex < 0 Then mblnNotChange = False: Exit Sub
    If Val(txt���.Text) < 0 Then
        IDKindPayMode.IDKind = 1
        IDKindPayMode.GetCurCard.���� = "Ӧ��"
        txt���.Text = Format(-1 * Val(txt���.Text), "0.00")
    Else
        If cbo֧����ʽ.ListIndex >= 0 And IDKindPayMode.IDKind = 1 Then
            If cbo֧����ʽ.Text = "֧Ʊ" And cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) = 2 Then
                IDKindPayMode.GetCurCard.���� = "��֧Ʊ"
            Else
                If cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) = -1 Then
                    IDKindPayMode.IDKind = 2
                Else
                    IDKindPayMode.GetCurCard.���� = "�Ҳ�"
                End If
            End If
        End If
    End If
    If Not IDKindPayMode.GetCurCard Is Nothing Then IDKindPayMode.IDKind = IDKindPayMode.GetKindIndex(IDKindPayMode.GetCurCard.����)
    mblnNotChange = False
End Sub

Private Sub wndTaskPanel_GroupExpanded(ByVal Group As XtremeSuiteControls.ITaskPanelGroup)
        If Group.id = Idx_TP_PatiExpend Then
            mParaData.blnShowExpend = Group.Expanded
            Call SetCtrlMove
        End If
End Sub
Private Sub SetCtrlMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���ȱʡλ��
    '����:���˺�
    '����:2011-07-12 08:45:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTaskHeight As Single, sngWinHeight As Single
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    Dim vRectForm As RECT, vRect As RECT
    Dim sinW As Single, sinH As Single
    
    Err = 0: On Error Resume Next
    If mParaData.blnShowExpend Then
        sngTaskHeight = mFormMaxHeight - 100 - stbThis.Height
        sngWinHeight = mFormMaxHeight + 400
    Else
        sngTaskHeight = mFormMaxHeight - 100 - picExpend.Height - stbThis.Height
        sngWinHeight = mFormMaxHeight - picExpend.Height + 400
    End If
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    picCard.Height = 2155
    If mEditType <> Cr_�󶨿� And mEditType <> Cr_���� Then
        picCard.Height = 1550
        sngTaskHeight = sngTaskHeight - 1050
        sngWinHeight = sngWinHeight - 1100
    Else
        If mEditType = Cr_���� Then
            picCard.Height = picCard.Height - cbo֧����ʽ.Height * 2 + 420
            sngTaskHeight = sngTaskHeight - cbo֧����ʽ.Height
            sngWinHeight = sngWinHeight - cbo֧����ʽ.Height
        Else
            '����صĿ�����Ϣ
            picCard.Height = picCard.Height - cbo֧����ʽ.Height * 2 - 320
            sngTaskHeight = sngTaskHeight - cbo֧����ʽ.Height - 750
            sngWinHeight = sngWinHeight - cbo֧����ʽ.Height - 750
        End If
        If Not mblnAddPage Then
            picCard.Height = picCard.Height - 350
            sngTaskHeight = sngTaskHeight - 350: sngWinHeight = sngWinHeight - 350
        End If
    End If
    '���¼���һ�η���ҳ��
    wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Clear
    Set Item = wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Add(Idx_TP_PatiCard, "", xtpTaskItemTypeControl)
    Set Item.Control = picCard: tkpGroup.Expanded = True
    wndTaskPanel.Reposition
    
    If mEditType = Cr_���� Then
        lbl����.Top = lbl����.Top: lbl����.Top = lbl����.Top: lbl��֤.Top = lbl����.Top
        txt����.Top = txt����.Top: txtAudi.Top = txt����.Top: txtPass.Top = txt����.Top
        txtˢ������.Left = txt����.Left: lblˢ����֤.Left = txtˢ������.Left - lblˢ����֤.Width - 20
        txtˢ������.Width = txt����.Width
        '�����:50893
        lblԭ������.Top = lblˢ����֤.Top: txtԭ������.Top = lblԭ������.Top - (txtԭ������.Height - lblԭ������.Height) / 2
        lblԭ������.Left = txtԭ������.Left - lblԭ������.Width - 50
    End If
    If mEditType = Cr_��ʧ Then
        txtˢ������.Left = txt����.Left: lblˢ����֤.Left = txtˢ������.Left - lblˢ����֤.Width - 50
        txtˢ������.Width = txt����.Width
    End If
    If mEditType <> Cr_�˿� Then
        sngTaskHeight = sngTaskHeight - picTittle.Height + 150
        sngWinHeight = sngWinHeight - picTittle.Height + 150
    End If
    If mEditType = Cr_����������Ϣ Then
        sngTaskHeight = sngTaskHeight - picCard.Height - picTittle.Height
        sngWinHeight = sngWinHeight - picCard.Height - picTittle.Height
    End If
    If mEditType = Cr_��ѯ Then
        txt����Ա.Top = txt�䶯ԭ��.Top: txtDate.Top = txt����Ա.Top
        lbl������.Top = lbl����.Top: lblDate.Top = lbl����.Top
        picCard.Height = picCard.Height - txt����Ա.Height - 50
        '���¼���һ�η���ҳ��
        wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Clear
        Set Item = wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Add(Idx_TP_PatiCard, "", xtpTaskItemTypeControl)
        Set Item.Control = picCard: tkpGroup.Expanded = True
        wndTaskPanel.Reposition
        sngTaskHeight = sngTaskHeight - 50
        sngWinHeight = sngWinHeight - 50

    End If
    '�����:56599

    wndTaskPanel.Height = sngTaskHeight
    Me.Height = sngWinHeight
 
    cmdHelp.Top = ScaleHeight - cmdHelp.Height - 100 - stbThis.Height
    
    '73063,Ƚ����,2014-5-20
    vRectForm = GetControlRect(Me.hWnd)
    vRect = GetControlRect(fraCard.hWnd)
    '����߿���
    sinW = (vRectForm.Right - vRectForm.Left - Me.ScaleWidth) / 2
    '�������߶�
    sinH = vRectForm.Bottom - vRectForm.Top - Me.ScaleHeight - sinW
    '��λ
    picԤ�����.Top = vRect.Top - vRectForm.Top - sinH - IIf(mEditType = Cr_�˿�, 120, 0)
'    picԤ�����.Top = wndTaskPanel.Height - picCard.Height - picԤ�����.Height - IIf(mEditType = Cr_�˿�, 80, 180)
End Sub


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-06-21 13:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKindStr As String, blnVisible As Boolean
    Dim intKind As Integer, strKey As String
    
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitPara: Call ClearData: Call InitData:  Call InitDicts
    Call InitInsurePara
    '74449,Ƚ����,2014-6-25,���ϴη�����𡱲����ڻ�ͣ��ʱ�޷���ȡ���������
    Call InitIDKind
    Call InitCardType
    '74539,Ƚ����,2014-6-27,���շѴ���Ժ�ڿ����ڲ��˱䶯��¼����ı䶯����Ϊ11���󶨿�����Ӧ��Ϊ1��������
    Call SetCardPayOrBound '���õ�ǰ���Ĳ�������
    Call SetDefaultLen
    'IDKind.IDKindStr = GetIDKindStr(IDKind.IDKindStr)
    
    mlngȱʡ���ų��� = IDKind.GetDefaultCardNoLen
    mintTabIndex���� = txt����.TabIndex: mintTabIndexˢ������ = txtˢ������.TabIndex
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strKey)
    intKind = Val(strKey)
     If intKind > 0 And intKind <= IDKind.ListCount Then IDKind.IDKind = intKind
     
    'ȡȱʡ��ˢ����ʽ
    '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    '��7λ��,��ֻ��������,��Ȼȡ������
    mblnDefaultPassInputCardNo = IDKind.ShowPassText
    Call SetBrushCardObject
    
    '94941:���ϴ�,2016/4/7,�޸������Ȩ��
    txt�����.Locked = InStr(";" & mstrPrivs & ";", ";�����޸������;") <= 0
    
    '��ʼ����ַ�ؼ�
    If Not mblnStructAdress Then Exit Sub
    padd��ͥ��ַ.Visible = mblnStructAdress: padd���ڵ�ַ.Visible = mblnStructAdress
    padd��ͥ��ַ.ShowTown = mblnShowTown: padd���ڵ�ַ.ShowTown = mblnShowTown
    txt��ͥ��ַ.Visible = False: cmd��ͥ��ַ.Visible = False
    padd��ͥ��ַ.Top = txt��ͥ��ַ.Top: padd��ͥ��ַ.Left = txt��ͥ��ַ.Left
    txt���ڵ�ַ.Visible = False: cmd���ڵ�ַ.Visible = False
    padd���ڵ�ַ.Top = txt���ڵ�ַ.Top: padd���ڵ�ַ.Left = txt���ڵ�ַ.Left
End Sub
Private Function SetBrushCardObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '����:���˺�
    '����:2011-07-08 11:06:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjReadCard Is Nothing Then
        Set mobjReadCard = zlGetComponentObject(mlngCardTypeID, False)
    End If
    If mobjReadCard Is Nothing Then Exit Function
    'zlInitComponents(ByVal frmMain As Object, _
    '    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '    ByVal cnOracle As ADODB.Connection, _
    '    Optional blnDeviceSet As Boolean = False, _
    '    Optional strExpand As String
    If Not mobjReadCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") Then
        Set mobjReadCard = Nothing: Exit Function
    End If
    SetBrushCardObject = True
End Function
Private Function InitCompoent(ByVal lngCardTypeID As Long, bln���ѿ� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ָ������
    '���:lngCardTypeID-��ʼ�������ID
    '        bln���ѿ�-���ѿ�
    '����:
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-09 23:50:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Object
    Set objCard = zlGetComponentObject(lngCardTypeID, bln���ѿ�)
    If objCard Is Nothing Then Exit Function
    'zlInitComponents(ByVal frmMain As Object, _
    '    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '    ByVal cnOracle As ADODB.Connection, _
    '    Optional blnDeviceSet As Boolean = False, _
    '    Optional strExpand As String
    If objCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") Then
         Exit Function
    End If
    InitCompoent = True
End Function
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-07-05 10:14:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    IDKind.Font = lbl����.Font
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    mblnChange = False: cbo���䵥λ.ListIndex = 0: mblnChange = True
    '������Ч��֧�����
    Call Load֧����ʽ
    If mEditType = Cr_��ʧ Then
        strSQL = "Select ����,����,����,��Ч����,ȱʡ��־ From ҽ�ƿ���ʧ��ʽ"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With cbo��ʧ��ʽ
            .Clear
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Int(Val(Nvl(rsTemp!��Ч����)) * 100)
                If Val(Nvl(rsTemp!ȱʡ��־)) = 1 Then
                    .ListIndex = .NewIndex
                End If
                rsTemp.MoveNext
            Loop
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Load֧����ʽ(Optional ByVal blnDel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    '128370:���ϴ�,2018/7/9,����Ϊ7��8�Ľ��㷽ʽ�����������ҽ�ƿ��ж�
    strSQL = "" & _
    "   Select Nvl(A.ȱʡ��־,0) as ȱʡ,B.����,B.����,B.����" & _
    "   From ���㷽ʽӦ�� A,���㷽ʽ B" & _
    "   Where A.���㷽ʽ=B.���� And A.Ӧ�ó���=[1]" & _
    "           And Nvl(B.����,1) IN(1,2,7,8)  " & _
    "   Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "���￨")
    Set mcolPayMode = New Collection
    '��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not blnDel Then strPayType = GetAvailabilityCardType
    varData = Split(strPayType, ";")
    With cbo֧����ʽ
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind And rsTemp!���� <> 7 And rsTemp!���� <> 8 Then
                .AddItem Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                mcolPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0, 1, 0), "K" & j
                If rsTemp!ȱʡ = 1 Then
                    .ListIndex = .NewIndex
                End If
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        For i = 0 To UBound(varData)
            '�����:116175��������2017/12/8����ҽ�ƿ��Ľɿʽ���Ƶ���Ϊ�ܽ��㷽ʽ������豸���ù�ͬ����
            rsTemp.Filter = "���� ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                If InStr(1, varData(i), "|") <> 0 Then
                    varTemp = Split(varData(i), "|")
                    mcolPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1)
                    .ItemData(.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
    End With
    If cbo֧����ʽ.ListCount > 0 And cbo֧����ʽ.ListIndex < 0 Then cbo֧����ʽ.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetControlVisitble()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Visible����
    '����:���˺�
    '����:2011-07-07 00:20:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    
    If mEditType = Cr_����������Ϣ Then
        picTittle.Visible = False
        picCard.Visible = False: Exit Sub
    End If
    '�����:56599
    cmdCreateCard.Visible = (mEditType = Cr_���� Or mEditType = Cr_�󶨿�) And InStr(1, mstrPrivs, ";�ƿ�;") > 0 And mCardType.bln�Ƿ��ƿ�
    
    If mEditType <> Cr_���� And mEditType <> Cr_�˿� Then picTittle.Visible = False
    
    blnVisible = mEditType = Cr_���� Or mEditType = Cr_���� Or mEditType = Cr_�˿� Or chkCancel.value = 1
    txt����.Visible = blnVisible: cbo֧����ʽ.Visible = blnVisible
    chk����.Visible = blnVisible: lbl����.Visible = blnVisible
    lbl֧����ʽ.Visible = blnVisible: txt�ϼ�.Visible = blnVisible
    IDKindPayMode.Visible = blnVisible: txt���.Visible = blnVisible
    blnVisible = Not mCardType.blnˢ�� And (blnVisible Or mEditType = Cr_�󶨿� Or mEditType = Cr_����)
    If mCardType.blnOneCard Or mCardType.str������ = "�������֤" Then  '�����:53408
        cmdReadCard.Visible = False '������һ��ͨ
    Else
        blnVisible = blnVisible And Not mCardType.bln���￨
        cmdReadCard.Visible = blnVisible And Not mCardType.bln���￨
        lbl����.BorderStyle = IIf(mCardType.bln���￨ And mEditType <> Cr_�˿�, 1, 0) '����� ��57962
    End If
    txtˢ������.TabIndex = mintTabIndexˢ������: txt����.TabIndex = mintTabIndex����
    '�˿���һЩ����
    If (mEditType = Cr_�˿� Or chkCancel.value = 1) _
        And InStr(1, "123", mParaData.int�˿�ģʽ) > 0 Then
        '0-������ˢ��;1-ˢ���˿�;2-���ݺź�����֤ˢ��;3-1��2�Ĺ���ģʽ
        cmdReadCard.Left = txtˢ������.Left + txtˢ������.Width - cmdReadCard.Width
        lbl����.Visible = False: lbl��֤.Visible = False
        txtPass.Visible = False: txtAudi.Visible = False
        lblˢ����֤.Visible = True: txtˢ������.Visible = True
        lblˢ����֤.BorderStyle = IIf(mCardType.bln���￨, 1, 0)
        'lblˢ����֤.Caption = "ˢ����֤"
    ElseIf mEditType = Cr_���� Then
        lblˢ����֤.Visible = True: txtˢ������.Visible = True
        lblˢ����֤.Caption = "ԭ����"
        txtˢ������.TabIndex = mintTabIndex����: txt����.TabIndex = mintTabIndexˢ������
        '50893
        lblԭ������.Visible = True: txtԭ������.Visible = True: txtԭ������.TabIndex = txtˢ������.TabIndex + 1
        txt����.TabIndex = txtԭ������.TabIndex + 1
    ElseIf mEditType = Cr_��ʧ Then
        lbl����.Visible = True: cbo��ʧ��ʽ.Visible = True
        lbl����.Caption = "��ʧ��ʽ"
        lblˢ����֤.Visible = True: txtˢ������.Visible = True: txt����.Visible = False
        lblˢ����֤.Caption = "��ʧ����"
        lbl����.Visible = False: txtPass.Visible = False: txtAudi.Visible = False
        lbl����.Visible = True: txt�䶯ԭ��.Visible = True: lbl����.Caption = "��ʧԭ��"
        txt�䶯ԭ��.Tag = "��ʧԭ��"
        lbl������.Caption = "��ʧ��": lblDate.Caption = "��ʧʱ��"
    Else
        cmdReadCard.Left = txt����.Left + txt����.Width
        lbl����.Visible = True: lbl��֤.Visible = True
        txtPass.Visible = True: txtAudi.Visible = True
        lblˢ����֤.Visible = False: txtˢ������.Visible = False
        If mEditType = Cr_��ѯ Then
            cmdOK.Visible = False: cmdCancel.Top = cmdOK.Top
            cmdCancel.Caption = "�˳�(&C)"
        End If
    End If

    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    If (mEditType = Cr_���� Or mEditType = Cr_�󶨿�) And chkCancel.value = 0 Then
        IDKindPay.Visible = True: IDKindPay.Enabled = True: lbl����.BorderStyle = 0
        lbl����.Left = IDKindPay.Left - lbl����.Width
        IDKindPay.Top = txt����.Top
        cmdReadCard.Visible = False: fraCard.BorderStyle = 0
    Else
        IDKindPay.Visible = False: IDKindPay.Enabled = False
        lbl����.Left = txt����.Left - lbl����.Width - 60
        fraCard.BorderStyle = IIf(mEditType = Cr_���� Or mEditType = Cr_�󶨿�, 0, 1)
    End If
    
    '�����:73063
    picԤ�����.Visible = mEditType = Cr_�˿� Or chkCancel.value = 1

    If mEditType = Cr_�˿� Or chkCancel.value = 1 Then
        IDKindPayMode.Visible = False: txt���.Visible = False
    End If
End Sub

Private Sub SetControlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enable����
    '����:���˺�
    '����:2011-07-05 10:19:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim objCtl As Control
   Dim blnEdit As Boolean
   '�����:56599
   If mEditType <> Cr_���� And mEditType <> Cr_�󶨿� Then
        cmdPicFile.Enabled = False: cmdPicCollect.Enabled = False: cmdPicClear.Enabled = False
   End If
   blnEdit = ((mEditType = Cr_����) Or (mEditType = Cr_�󶨿�)) And chkCancel.value = 0
    
   blnEdit = blnEdit And mrsInfo Is Nothing
   For Each objCtl In Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")  '�ı�
                If objCtl.Tag = "����" Then
                    objCtl.Enabled = (mEditType = Cr_���� Or mEditType = Cr_�󶨿� Or mEditType = Cr_���� Or mEditType = Cr_���� Or mEditType = Cr_����������Ϣ Or mEditType = Cr_��ʧ) And chkCancel.value = 0
                ElseIf InStr(1, ",��סַ,���ڵ�ַ,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = (blnEdit Or mEditType = Cr_����������Ϣ) And Not mblnStructAdress
                Else
                    objCtl.Enabled = (blnEdit Or mEditType = Cr_����������Ϣ)
                End If
                If InStr(1, ",����,����,��֤,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = (mEditType = Cr_���� Or mEditType = Cr_���� Or mEditType = Cr_�󶨿� Or mEditType = Cr_����) And chkCancel.value = 0
                End If
                If "����" = objCtl.Tag Then
                    objCtl.Enabled = (mEditType = Cr_���� Or mEditType = Cr_����) And chkCancel.value = 0
                    If mCardType.rsҽ�ƿ��� Is Nothing Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rsҽ�ƿ���.State <> 1 Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rsҽ�ƿ���.RecordCount = 0 Then
                        objCtl.Enabled = False
                    End If
                End If
                If InStr(1, ",������λ,��λ�绰,��λ�ʱ�,��λ������,��λ�ʺ�,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = (blnEdit Or mEditType = Cr_����������Ϣ) And InStr(mstrPrivs, ";��Լ���˵Ǽ�;") > 0
                End If
                If InStr(1, ",ˢ������,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = mEditType = Cr_�˿� Or mEditType = Cr_���� Or chkCancel.value = 1 Or mEditType = Cr_��ʧ
                End If
                If InStr(1, ",�䶯ԭ��,��ʧԭ��,", "," & objCtl.Tag & ",") > 0 Then
                      '�䶯ԭ��͹�ʧԭ����һ���ؼ�txt�䶯ԭ��.tag
                      objCtl.Enabled = mEditType = Cr_��ʧ
                End If
                '�����:56599
                If objCtl Is txtOtherWaring Then
                    objCtl.Enabled = True
                End If
                objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, &H8000000F)
        Case UCase("ComboBox")
                If Not objCtl Is cbo֧����ʽ Then
                    If objCtl Is cboNO Then
                        objCtl.Enabled = mEditType <> Cr_��ѯ
                    ElseIf objCtl Is cbo��ʧ��ʽ Then
                        objCtl.Enabled = mEditType = Cr_��ʧ
                    Else
                        objCtl.Enabled = (blnEdit Or mEditType = Cr_����������Ϣ)
                    End If
                    objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, &H8000000F)
                Else
                    objCtl.Enabled = chk����.value = 1 And (mEditType = Cr_���� Or mEditType = Cr_����)
                    If mCardType.rsҽ�ƿ��� Is Nothing Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rsҽ�ƿ���.State <> 1 Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rsҽ�ƿ���.RecordCount = 0 Then
                        objCtl.Enabled = False
                    End If
                End If
                '�����:56599
                If objCtl Is cboBloodType Or objCtl Is cboBH Then
                    objCtl.Enabled = True
                End If
        Case UCase("MaskEdBox")
                objCtl.Enabled = (blnEdit Or mEditType = Cr_����������Ϣ)
                objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, &H8000000F)
        Case UCase("CommandButton")
            If InStr(1, ",�����ص�,����,������λ,��סַ,���ڵ�ַ,��ϵ�˵�ַ,", "," & objCtl.Tag & ",") > 0 Then
                objCtl.Visible = (blnEdit Or mEditType = Cr_����������Ϣ)
                If objCtl.Tag = "��סַ" Then objCtl.Visible = objCtl.Visible And Not mblnStructAdress
                If objCtl.Tag = "���ڵ�ַ" Then objCtl.Visible = objCtl.Visible And Not mblnStructAdress
                If objCtl.Tag = "������λ" Then
                    objCtl.Visible = InStr(mstrPrivs, ";��Լ���˵Ǽ�;") > 0 And blnEdit
                End If
            End If
        Case UCase("CheckBox")
            If chkCancel Is objCtl Then
                objCtl.Enabled = mEditType <> Cr_��ѯ
            Else
                objCtl.Enabled = (mEditType = Cr_���� Or mEditType = Cr_����) And chkCancel.value = 0
                If mCardType.rsҽ�ƿ��� Is Nothing Then
                    objCtl.Enabled = False
                ElseIf mCardType.rsҽ�ƿ���.State <> 1 Then
                    objCtl.Enabled = False
                ElseIf mCardType.rsҽ�ƿ���.RecordCount = 0 Then
                    objCtl.Enabled = False
                End If
            End If
        Case UCase("PatiAddress")
            objCtl.Enabled = (blnEdit Or mEditType = Cr_����������Ϣ) And mblnStructAdress
            objCtl.ControlLock = Not objCtl.Enabled
        End Select
    Next
    txtDate.Enabled = False
    If mEditType = Cr_����������Ϣ Then
    
        '���ܸ��Ĳ������� 67184
        blnEdit = mlng����ID = 0
        If Not blnEdit Then
            blnEdit = Not zlExistOperationData(mlng����ID, "")
        End If
        cbo�Ա�.Enabled = blnEdit
        txt����.Enabled = blnEdit
        cbo���䵥λ.Enabled = blnEdit
        txt��������.Enabled = blnEdit
        txt����ʱ��.Enabled = blnEdit
        cbo�Ա�.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
        txt����.BackColor = cbo�Ա�.BackColor
        cbo���䵥λ.BackColor = cbo�Ա�.BackColor
        txt��������.BackColor = cbo�Ա�.BackColor
        txt����ʱ��.BackColor = cbo�Ա�.BackColor
    End If
    Call SetCardEditEnabled
End Sub
Public Sub ClearData(Optional ByVal blnSave As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
   '����:���˺�
    '����:2011-07-03 10:14:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control
    Set mrsInfo = Nothing
    For Each objCtl In Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")  '�ı�
            objCtl.Text = ""
        Case UCase("ComboBox")
            objCtl.ListIndex = -1
        Case UCase("MaskEdBox")
            If InStr(1, ",��������,����ʱ��,", "," & objCtl.Tag & ",") > 0 Then
                 objCtl.Text = IIf(objCtl.Tag = "��������", "____-__-__", "__:__")
            End If
        Case UCase("Command")
        Case UCase("Image") '�����:56599
            objCtl.Picture = Nothing
        Case UCase("VSFlexGrid") '�����:56599
            objCtl.Rows = 1
            objCtl.Rows = 2
        Case UCase("Patiaddress")
            objCtl.value = ""
        End Select
    Next
    Call SetDefaultValue
    chk����.value = IIf(mParaData.bln����, 1, 0)
    If mEditType = Cr_�˿� Or chkCancel.value = 1 Then
        lbl֧����ʽ.Caption = "�˿�"
    Else
        lbl֧����ʽ.Caption = "�ɿ�"
    End If
    If blnSave Then Call setFact
    mblnChange = False
    mstr���� = ""
    mstr���䵥λ = ""
End Sub
Private Sub SetDefaultValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡֵ
    '����:���˺�
    '����:2011-07-28 09:00:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call SetCboDefault(cbo�Ա�)
    Call SetCboDefault(cbo�ѱ�)
    Call SetCboDefault(cboҽ�Ƹ���)
    Call SetCboDefault(cbo����)
    Call SetCboDefault(cbo����)
    Call SetCboDefault(cboѧ��)
    Call SetCboDefault(cbo����״��)
    Call SetCboDefault(cboְҵ)
    Call SetCboDefault(cbo���)
    Call SetCboDefault(cbo��ϵ�˹�ϵ)
    Call SetCboDefault(cbo֧����ʽ)
    Call SetCboDefault(cbo���䵥λ)
    If cbo���䵥λ.ListIndex < 0 And cbo���䵥λ.ListCount > 0 Then cbo���䵥λ.ListIndex = 0
    'Call SetCboDefault(cbo��������)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:MM")
    txt�����.Text = zlGet�����
    txt����Ա.Text = UserInfo.����
    '�����:56599
    Set mdicҽ�ƿ����� = Nothing
    mstr�ɼ�ͼƬ = ""
    mlngͼ����� = 0
    '��ʼ����ַ��Ϣ
    Call zlLoadDefaultAddr(padd��ͥ��ַ)
    Call zlLoadDefaultAddr(padd���ڵ�ַ)
End Sub

Private Sub AutoBrushSet(ByVal objIDKind As IDKindNew, blnAutoRefrsh As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ�ˢ������
    '����:���˺�
    '����:2011-06-20 13:31:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(blnAutoRefrsh)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(blnAutoRefrsh)
    Call objIDKind.SetAutoReadCard(blnAutoRefrsh)
End Sub

Private Sub txtPatient_GotFocus()
    If Not txtPatient.Enabled Or txtPatient.Locked Then Exit Sub
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    Call AutoBrushSet(IDKind, txtPatient.Text = "")
    zlControl.TxtSelAll txtPatient
    If IsCardType(IDKind, "����") Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub
Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False
    Call zlCommFun.OpenIme(False)
End Sub
Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim strCardNo As String, blnNotMsg As Boolean
    Dim blnPass As Boolean
    On Error GoTo errH
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
'    If Not mrsInfo Is Nothing And mEditType = Cr_����������Ϣ And KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If IsCardType(IDKind, "����") Then
        '105567:���ϴ�,2017/5/24,���ż��ܵ��µ�һ������ƴ�����ܴ������뷨
        blnPass = txtPatient.PasswordChar <> ""
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, mblnDefaultPassInputCardNo)
        txtPatient.IMEMode = 0
        blnPass = txtPatient.PasswordChar = "" And blnPass
        If blnPass Then
            If txtPatient.SelLength = Len(txtPatient.Text) Then
                mblnNotCharge = True
                txtPatient.Text = ""
                mblnNotCharge = False
            End If
            SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
        End If
    ElseIf IsCardType(IDKind, "�����") Or IsCardType(IDKind, "סԺ��") Or IsCardType(IDKind, "�ֻ���") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
         txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    txtPatient.IMEMode = 0
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then
        '����ˢ���ͻس�,���˳�
        Exit Sub
    End If
    
    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
    End If
    blnNotMsg = mEditType = Cr_���� Or mEditType = Cr_�󶨿�
    
    KeyAscii = 0
    strCardNo = Trim(txtPatient.Text)
    If Not GetPatient(txtPatient.Text, blnCard, blnNotMsg) Then
        '�������˻�����Ϣʱ,����Ҳ���ܱ�����,���Բ������������Ϣ
        If Not mrsInfo Is Nothing And mEditType = Cr_����������Ϣ Then
            If mrsInfo.State = 1 Then Exit Sub
        End If
        strCardNo = Trim(txtPatient.Text): Call ClearData
        '101170:���ϴ�,2017/5/3��������Ϣ����
        If IDKind.IDKind = IDKind.GetKindIndex("����") Or blnCard Then
            Call zlQueryEMPIPatiInfo(strCardNo)
            If Not blnCard And Trim(txtPatient.Text) <> "" Then strCardNo = Trim(txtPatient.Text)
        End If
        If blnCard Then
            If mEditType = Cr_���� Or mEditType = Cr_�󶨿� Then
               If IDKindDefaultKind = mlngCardTypeID Then
                   txt����.Text = strCardNo
               End If
            End If
            txtPatient.PasswordChar = ""
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
            zlControl.TxtSelAll txtPatient
        Else
            txtPatient.Text = strCardNo: zlControl.TxtSelAll txtPatient
        End If
        Call SetControlEnable
        lblҽ����(1).Visible = True: txt��֤ҽ����.Visible = True
        If mInsurePara.lng���ʽҽ������ = 0 Or Not (mEditType = Cr_���� Or mEditType = Cr_�󶨿�) Then
            lblҽ����(1).Visible = False
            txt��֤ҽ����.Visible = False
        End If
        
        If InStr(1, "+*-", Left(txtPatient.Text & " ", 1)) > 0 Then
            KeyAscii = 0
            DoEvents
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            zlControl.TxtSelAll txtPatient
            
            Exit Sub
        End If
        '76609,Ƚ����,2014-8-14,���㶨λ����
        If IDKind.GetCurCard.�ӿ���� = 0 And Not blnCard Then zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    If mEditType = Cr_���� Or mEditType = Cr_��ʧ Then
        If blnCard Then txtˢ������.Text = strCardNo
    End If
    txtPatient.Text = Nvl(mrsInfo!����)
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0

    Call LoadPatiInfor: SetControlEnable: Call zlQueryEMPIPatiInfo
    lblҽ����(1).Visible = True: txt��֤ҽ����.Visible = True
    If mInsurePara.lng���ʽҽ������ = 0 Or mEditType <> Cr_����������Ϣ Then
        lblҽ����(1).Visible = False
        txt��֤ҽ����.Visible = False
    End If
    '76609,Ƚ����,2014-8-14,���㶨λ����
'    If blnCard Then
        zlCommFun.PressKey vbKeyTab
'    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function LoadPatiInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-04 11:51:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�ѱ� As String
    On Error GoTo errHandle
    Call LoadCardFee
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    Call zlControl.CboLocate(cbo����, Nvl(mrsInfo!����))
    txt�����.Text = Nvl(mrsInfo!�����)
    lbl�����.Tag = Nvl(mrsInfo!�����)
    txtPatient.Text = mrsInfo!����
    txtҽ����.Text = Nvl(mrsInfo!ҽ����)
    '�����:51071
    txt��ϵ�����֤��.Text = Nvl(mrsInfo!��ϵ�����֤��)
    If mEditType = Cr_����������Ϣ Then
        '���ҽ��,����Ժ����ʵҽ�����˿����޸�ҽ����
        txtҽ����.Enabled = mInsurePara.lng���ʽҽ������ > 0 Or Not IsNull(mrsInfo!סԺ����) And IsNull(mrsInfo!����)
        lblҽ����(0).Tag = txtҽ����.Text
        If mInsurePara.lng���ʽҽ������ > 0 Then txt��֤ҽ����.Text = txtҽ����.Text
    End If
    
    
    Call zlControl.CboLocate(cbo�Ա�, Nvl(mrsInfo!�Ա�))
    If cbo�Ա�.ListIndex = -1 And Not IsNull(mrsInfo!�Ա�) Then
        cbo�Ա�.AddItem mrsInfo!�Ա�, 0
        cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
    End If
    Call LoadOldData("" & mrsInfo!����, txt����, cbo���䵥λ)
    mblnNotChange = True
    txt��������.Text = Format(IIf(IsNull(mrsInfo!��������), "____-__-__", mrsInfo!��������), "YYYY-MM-DD")
    If Not IsNull(mrsInfo!��������) Then
         'txt����.Text = ReCalcOld(CDate(txt��������.Text), cbo���䵥λ, Val(Nvl(mrsInfo!����ID)))   '�޸ĵ�ʱ��,���ݳ���������������
         'If CDate(txt��������.Text) - CDate(mrsInfo!��������) <> 0 Then txt����ʱ��.Text = Format(mrsInfo!��������, "HH:MM")
     Else
         txt����ʱ��.Text = "__:__"
         txt��������.Text = ReCalcBirth(Val(txt����.Text), cbo���䵥λ.Text)
     End If
    txt���֤��.Text = Nvl(mrsInfo!���֤��)

    '���ݲ�ͬ�鿴��ʽ��ȡ��ͬ�ķѱ�
    str�ѱ� = Nvl(mrsInfo!�ѱ�)
    cbo�ѱ�.ListIndex = GetCboIndex(cbo�ѱ�, str�ѱ�)
    If cbo�ѱ�.ListIndex = -1 And str�ѱ� <> "" Then
        cbo�ѱ�.AddItem str�ѱ�, 0
        cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
    End If
    mblnNotChange = False
        
    cboҽ�Ƹ���.ListIndex = GetCboIndex(cboҽ�Ƹ���, Nvl(mrsInfo!ҽ�Ƹ��ʽ))
    If cboҽ�Ƹ���.ListIndex = -1 And Not IsNull(mrsInfo!ҽ�Ƹ��ʽ) Then
        cboҽ�Ƹ���.AddItem mrsInfo!ҽ�Ƹ��ʽ, 0
        cboҽ�Ƹ���.ListIndex = cboҽ�Ƹ���.NewIndex
    End If
       
   cbo����.ListIndex = GetCboIndex(cbo����, Nvl(mrsInfo!����))
   If cbo����.ListIndex = -1 And Not IsNull(mrsInfo!����) Then
       cbo����.AddItem mrsInfo!����, 0
       cbo����.ListIndex = cbo����.NewIndex
   End If
   
   cbo����.ListIndex = GetCboIndex(cbo����, Nvl(mrsInfo!����))
   If cbo����.ListIndex = -1 And Not IsNull(mrsInfo!����) Then
       cbo����.AddItem mrsInfo!����, 0
       cbo����.ListIndex = cbo����.NewIndex
   End If
   
   txt����.Text = Nvl(mrsInfo!����)
   
   cboѧ��.ListIndex = GetCboIndex(cboѧ��, Nvl(mrsInfo!ѧ��))
   If cboѧ��.ListIndex = -1 And Not IsNull(mrsInfo!ѧ��) Then
       cboѧ��.AddItem mrsInfo!ѧ��, 0
       cboѧ��.ListIndex = cboѧ��.NewIndex
   End If
   
   cbo����״��.ListIndex = GetCboIndex(cbo����״��, Nvl(mrsInfo!����״��))
   If cbo����״��.ListIndex = -1 And Not IsNull(mrsInfo!����״��) Then
       cbo����״��.AddItem mrsInfo!����״��, 0
       cbo����״��.ListIndex = cbo����״��.NewIndex
   End If
   
   cboְҵ.ListIndex = GetCboIndex(cboְҵ, Nvl(mrsInfo!ְҵ), , , mstrCboSplit)
   If cboְҵ.ListIndex = -1 And Not IsNull(mrsInfo!ְҵ) Then
       cboְҵ.AddItem mrsInfo!ְҵ, 0
       cboְҵ.ListIndex = cboְҵ.NewIndex
   End If
   
   cbo���.ListIndex = GetCboIndex(cbo���, Nvl(mrsInfo!���))
   If cbo���.ListIndex = -1 And Not IsNull(mrsInfo!���) Then
       cbo���.AddItem mrsInfo!���, 0
       cbo���.ListIndex = cbo���.NewIndex
   End If
        
   txt�����ص�.Text = Nvl(mrsInfo!�����ص�)
   txt��ͥ��ַ.Text = Nvl(mrsInfo!��ͥ��ַ)
   '��ȡ���˵�ַ��Ϣ
    Call zlReadAddrInfo(padd��ͥ��ַ, Val(Nvl(mrsInfo!����ID)), 0, 3, txt��ͥ��ַ.Text)
   txt��ͥ�绰.Text = Nvl(mrsInfo!��ͥ�绰)
   txt�ֻ�.Text = Nvl(mrsInfo!�ֻ���)
   txt��ͥ�ʱ�.Text = Nvl(mrsInfo!��ͥ��ַ�ʱ�)
   txt���ڵ�ַ.Text = Nvl(mrsInfo!���ڵ�ַ)
   Call zlReadAddrInfo(padd���ڵ�ַ, Val(Nvl(mrsInfo!����ID)), 0, 4, txt���ڵ�ַ.Text)
   txt���ڵ�ַ�ʱ�.Text = Nvl(mrsInfo!���ڵ�ַ�ʱ�)
   txt��ϵ������.Text = Nvl(mrsInfo!��ϵ������)
   cbo��ϵ�˹�ϵ.ListIndex = GetCboIndex(cbo��ϵ�˹�ϵ, Nvl(mrsInfo!��ϵ�˹�ϵ))
   If cbo��ϵ�˹�ϵ.ListIndex = -1 And Not IsNull(mrsInfo!��ϵ�˹�ϵ) Then
       cbo��ϵ�˹�ϵ.AddItem mrsInfo!��ϵ�˹�ϵ, 0
       cbo��ϵ�˹�ϵ.ListIndex = cbo��ϵ�˹�ϵ.NewIndex
   End If
   txt��ϵ�˵�ַ.Text = Nvl(mrsInfo!��ϵ�˵�ַ)
   txt��ϵ�˵绰.Text = Nvl(mrsInfo!��ϵ�˵绰)
   txt������λ.Text = Nvl(mrsInfo!������λ)
   lbl������λ.Tag = Nvl(mrsInfo!��ͬ��λid)
   txt��λ�绰.Text = Nvl(mrsInfo!��λ�绰)
   txt��λ�ʱ�.Text = Nvl(mrsInfo!��λ�ʱ�)
   txt��λ������.Text = Nvl(mrsInfo!��λ������)
   txt��λ�ʻ�.Text = Nvl(mrsInfo!��λ�ʺ�)
   txt����֤��.Text = "" & mrsInfo!����֤��
   'txt��ע.Text = IIf(IsNull(mrsInfo!��ע), "", mrsInfo!��ע)
   '�����:56599
    Load�����������Ϣ Nvl(mrsInfo!����ID)
    
    mblnChange = False
    LoadPatiInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim IDkindIndex As Integer
    Dim blnǩԼ As Boolean
    Dim strErrMsg As String
    Dim bln����ǩԼ As Boolean '�Ƿ�����ǩԼ,�����֤����Ϣ���ȡ���Ĳ�����Ϣ �Ƿ�һ���ж�
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        '���Ҳ���
        mblnNotClick = True
        IDkindIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        blnǩԼ = �Ƿ��Ѿ�ǩԼ(strID)
        If mCardType.str������ = "�������֤" Then
            '���������֤
            If blnǩԼ Then
                MsgBox "�����֤�Ѿ�ǩԼ,�����ٴ�ǩԼ!", vbInformation, Me.Caption
                Set mrsInfo = Nothing
                Call txtPatient_GotFocus
                Exit Sub
            End If
        End If
        If GetPatient(strID, False, True) Then
            If Not mrsInfo Is Nothing Then
                If mCardType.str������ = "�������֤" Then
                    '������֤�Ƿ�һֱ12-10-29 lgf
                    bln����ǩԼ = Not (Nvl(mrsInfo!����) <> Trim(strName) Or Nvl(mrsInfo!�Ա�) <> strSex _
                                      Or Format(Nvl(mrsInfo!��������, "00-00-00"), "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd"))

                    If Not bln����ǩԼ Then
                         If Nvl(mrsInfo!����) <> Trim(strName) Then
                             strErrMsg = strErrMsg & "," & "����"
                        End If

                        If Nvl(mrsInfo!�Ա�) <> strSex Then

                             strErrMsg = strErrMsg & "," & "�Ա�"
                        End If

                        If Format(txt��������.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                             strErrMsg = strErrMsg & "," & "��������"
                        End If

                        strErrMsg = Mid(strErrMsg, 2)
                        strErrMsg = "��ǰ������Ϣ�����֤�ϵ�[" & strErrMsg & "]����Ϣ��һ��," & vbCrLf & "���ܽ������֤ǩԼ!"
                        Call MsgBox(strErrMsg, vbInformation, Me.Caption)
                        Set mrsInfo = Nothing
                        Call txtPatient_GotFocus
                        Exit Sub
                    End If
                    txt����.Text = strID
                End If
                Call LoadPatiInfor: SetControlEnable: Call zlQueryEMPIPatiInfo
                '75717,Ƚ����,2014-7-22,�Һ�ԤԼʱ��ȡ�²������֤��Ƭ
                If imgPatient.Picture = 0 Then Call LoadIDImage
                txt���ڵ�ַ.Text = IIf(Trim(txt���ڵ�ַ.Text) = "", strAddress, txt���ڵ�ַ.Text)
                padd���ڵ�ַ.value = txt���ڵ�ַ.Text
                txtPatient.PasswordChar = ""
            End If
        Else
            '�²���
             txtPatient.Text = strName
             txt���֤��.Text = strID
             Call zlControl.CboLocate(cbo�Ա�, strSex)
             Call zlControl.CboLocate(cbo����, strNation)
             txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
             '�����:57817
             txt��ͥ��ַ.Text = IIf(Trim(txt��ͥ��ַ.Text) = "", strAddress, txt��ͥ��ַ.Text)
             padd��ͥ��ַ.value = txt��ͥ��ַ.Text
             
             txt���ڵ�ַ.Text = strAddress
             padd���ڵ�ַ.value = txt���ڵ�ַ.Text
             
             If mCardType.str������ = "�������֤" Then
                txt����.Text = strID
             End If
             Call LoadIDImage
             '�²���,����������ʾ
             txtPatient.PasswordChar = ""
             Call zlQueryEMPIPatiInfo
        End If
        IDKind.IDKind = IDkindIndex
        mblnNotClick = False
        
         '�����53408
        If mCardType.str������ = "�������֤" Then
            txt���֤��.PasswordChar = IIf(mCardType.str�������� <> "", "*", "")
        Else
            txt���֤��.PasswordChar = ""
        End If
        
        '�����:58072
        'Call SetControlEnable
        zlCommFun.PressKey vbKeyTab
    End If
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
End Sub
Private Sub txtԭ������_Change()
'�����:50893
    mblnChange = True
    Call SetCardEditEnabled
End Sub

Private Sub txtԭ������_GotFocus()
'�����:50893
    zlControl.TxtSelAll txtԭ������
    zlCommFun.OpenIme False
End Sub

Private Sub txtԭ������_KeyPress(KeyAscii As Integer)
'�����:50893
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Function GetPatient(ByVal strInput As String, Optional ByVal blnCard As Boolean, Optional blnNotMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=��ʾ�Ƿ���￨ˢ��
    '����:
    '����:���˶�ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-03 10:46:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String, blnHavePass As Boolean
    Dim lng����ID As Long, blnCancel As Boolean, blnICCard As Boolean
    Dim strPassWord As String, bln�����ʻ� As Boolean, strErrMsg As String
    Dim strCardNo As String, lng�����ID As Long, blnIsMobileNO As Boolean
    
    txtPatient.ForeColor = &HFF0000
    strErrMsg = ""
    blnIsMobileNO = IDKind.IsMobileNO(strInput)
    If IsCardType(IDKind, "IC����") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   'ˢ����ȱʡ�Ŀ�
        
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        ElseIf IDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
        Else
            lng�����ID = -1
        End If
        
        If InStr("," & "���֤��,�������֤��,�������֤,���֤" & ",", "," & IDKind.GetCurCard.���� & ",") > 0 Then
             lng����ID = GetPatiIDFromProcedure(mlngModule, Me, strInput, txtPatient, , , blnCancel)
        End If
        If lng����ID = 0 And Not blnCancel Then
            '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
            If GetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then
                If blnIsMobileNO Then
                    '�ֻ��Ų���
                    If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                Else
                    GoTo NotFoundPati:
                End If
            End If
        End If
        
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strCardNo = strInput
        strInput = "-" & lng����ID
        blnHavePass = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strWhere = strWhere & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strWhere = strWhere & " And A.סԺ��=[1]"
    ElseIf IsCardType(IDKind, "����") And blnIsMobileNO Then
        '103000�����ϴ���2017/2/7�����ֻ��Ų���
        If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strInput = "-" & lng����ID
    Else
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                If mrsInfo!���� = strInput Then
                    '74309:���ϴ���2014-7-7������������ʾ��ɫ����
                    Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), txtPatient.ForeColor)
                    GetPatient = True: Exit Function
                    End If
            End If
        End If
        Select Case IDKind.GetCurCard.����
            Case "����", "��������￨"
                'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
                If Not mParaData.blnSeekName Or mEditType = Cr_����������Ϣ Then
                    If Not mEditType = Cr_����������Ϣ Then
                        Set mrsInfo = New ADODB.Recordset
                    End If
                    Exit Function
                End If
                strPati = _
                " Select  1 as ����ID,0 as ID,0 as ����ID,'[�²���]'  as ����,'' as �Ա�,'' as ����," & _
                "            0*NULL as �����,0*NULl as סԺ��,to_date(null) as ��������,null as ���֤��,null as ��ͥ��ַ, null as  ������λ" & _
                " From dual " & _
                " Union ALL " & _
                " Select  2 as ����ID ,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
                "            A.�����,A.סԺ��,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                " From ������Ϣ A,���ű� B" & _
                " Where A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And Rownum <101 And A.���� Like [1]" & _
                IIf(mParaData.intNameDays = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                strPati = strPati & " Order by  ����ID,����"
                
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTemp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "����ѡ��", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", mParaData.intNameDays)
                If blnCancel Then
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
                If rsTemp Is Nothing Then GoTo NotFoundPati:
                If rsTemp.State <> 1 Then GoTo NotFoundPati:
                If rsTemp.RecordCount = 0 Then GoTo NotFoundPati:
                If Val(Nvl(rsTemp!����ID)) = 0 Then GoTo NotFoundPati:
                
                strInput = "-" & rsTemp!����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.ҽ����=[2]"
             Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                lng����ID = GetPatiIDFromProcedure(mlngModule, Me, strInput, txtPatient, , , blnCancel)
                If lng����ID = 0 And Not blnCancel Then
                    '�����:54197
                     If GetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg, , , , InStr(mstrPrivs, ";�ϲ�������Ϣ;") > 0) = False Then lng����ID = 0
                End If
                 strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "��ϵ�����֤��", "��ϵ�����֤" '�����:51071
                strInput = UCase(strInput)
                 If GetPatiID("��ϵ�����֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                 strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If GetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case Else
                '�������ĺ���
                If Val(IDKind.GetCurCard.�ӿ����) > 0 Then
                    lng�����ID = IDKind.GetCurCard.�ӿ����
                    bln�����ʻ� = IDKind.GetCurCard.�Ƿ�����ʻ�
                    If GetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                    strCardNo = strInput
                    blnHavePass = True
                Else
                    If GetPatiID(IDKind.GetCurCard.����, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
        End Select
    End If
    On Error GoTo errH
    '��ȡ������Ϣ
   strSQL = "" & _
    "   Select Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����," & _
    "        A.����id,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.�ѱ�,A.ҽ�Ƹ��ʽ," & _
    "        A.����,A.�Ա�,A.����,A.��������,A.�����ص�,A.���֤��," & _
    "        A.����֤��,A.���,A.ְҵ,A.����,A.����,A.����,A.ѧ��,A.����״��,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�໤��,A.��ϵ������," & _
    "        A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.��ͬ��λid,A.������λ,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.������," & _
    "        A.������,A.��������,A.����ʱ��,A.����״̬,A.��������,A.סԺ����,A.��ǰ����id,A.��ǰ����id,A.��ǰ����,A.��Ժʱ��,A.��Ժʱ��," & _
    "        A.��Ժ,A.Ic����,A.������,A.ҽ����,A.����,A.��ѯ����,A.�Ǽ�ʱ��,A.ͣ��ʱ��,A.����,A.��ϵ�����֤��,A.���ڵ�ַ,A.���ڵ�ַ�ʱ�," & _
    "        M.���� as ���ʽ����, decode(B1.��������,NULL,0,1,1,0) as ����,B1.��ע, " & _
    "        Nvl(Nvl(A.��������,B1.��������),Decode(Nvl(A.����,B1.����),Null,'��ͨ����','ҽ������')) ��������,B1.��Ժ����, C.���� ��������," & _
    "        A.�ֻ���" & _
    "   From ������Ϣ A,������ҳ B1,������� C ,ҽ�Ƹ��ʽ M" & _
    "   Where A.���� = C.���(+) And A.ҽ�Ƹ��ʽ=M.����(+) " & _
    "               And A.����ID=B1.����ID(+) And A.��ҳID=B1.��ҳID(+) " & _
    "               And A.ͣ��ʱ�� is NULL" & strWhere
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput)
    If mrsInfo.EOF Then GoTo NotFoundPati:
    If Not blnHavePass Then
        strPassWord = Nvl(mrsInfo!����֤��)
    End If
    '74309:���ϴ���2014-7-7������������ʾ��ɫ����
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), txtPatient.ForeColor)
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    If strErrMsg <> "" Then Exit Function
    
    If (IDKind.IDKind = IDKind.GetKindIndex("����") Or blnCard) And blnNotMsg Then
        txt�����.Text = zlGet�����
        Exit Function
    Else
        If blnCard Then
            MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����    ", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Else
            MsgBox "������Ϣδ�ҵ�,�����Ƿ�������ȷ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        End If
    End If
End Function


Private Sub InitDicts()
    Call ReadDict("�Ա�", cbo�Ա�)
    Call ReadDict("�ѱ�", cbo�ѱ�)
    Call ReadDict("ҽ�Ƹ��ʽ", cboҽ�Ƹ���)
    Call ReadDict("����", cbo����)
    Call ReadDict("����", cbo����)
    Call ReadDict("ѧ��", cboѧ��)
    Call ReadDict("����״��", cbo����״��)
    Call ReadDict("ְҵ", cboְҵ, , mstrCboSplit)
    Call ReadDict("���", cbo���)
    Call ReadDict("����ϵ", cbo��ϵ�˹�ϵ)
 
End Sub

Private Function ReadDict(strDict As String, cbo As ComboBox, _
    Optional strClass As String, Optional strSplit As String = "-") As Boolean
'���ܣ���ʼ��ָ���ʵ�
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long
    On Error GoTo errHandle
     If strDict = "���㷽ʽ" Then
        strSQL = "" & _
        "   Select Nvl(A.ȱʡ��־,0) as ȱʡ,B.����,B.����,B.����" & _
        "   From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        "   Where A.���㷽ʽ=B.���� And A.Ӧ�ó���=[1]" & _
        "           And Nvl(B.����,1) IN(1,2) Order by B.����"
    ElseIf strDict = "���" Then
        strSQL = "Select ����,����,����,Nvl(���ȼ�,0) as ȱʡ From " & strDict & " Order by ����"
    ElseIf strDict = "�ѱ�" Then
        strSQL = _
        "   Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ " & _
        "   From �ѱ�" & _
        "   Where ����=1 And Nvl(���޳���,0)=0 And Nvl(�������,3) IN(1,3)" & _
        "               And  Sysdate Between NVL(��Ч��ʼ,Sysdate-1) and NVL(��Ч����,Sysdate+1)" & _
        "   Order by ����"
    ElseIf strDict = "��������" Then
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ,��ɫ From �������� Order by ����"
    Else
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strClass)
    cbo.Clear
    If Not rsTemp.EOF Then
        For i = 1 To rsTemp.RecordCount
            cbo.AddItem rsTemp!���� & strSplit & rsTemp!����
            If rsTemp!ȱʡ = 1 Then
                cbo.ListIndex = cbo.NewIndex
                cbo.ItemData(cbo.NewIndex) = 1
            End If
            If strDict = "���㷽ʽ" And strClass = "Ԥ����" Then
                   cbo.ItemData(cbo.NewIndex) = Val(Nvl(rsTemp!����))
                   cbo.Tag = cbo.NewIndex   '��������Ϊȱʡ����������
            End If
            If TextWidth(cbo.List(cbo.NewIndex) & "�˺�") > lngMaxW Then lngMaxW = TextWidth(cbo.List(cbo.NewIndex) & "�˺�")
            rsTemp.MoveNext
        Next
        If strDict = "���㷽ʽ" And strClass <> "Ԥ����" Then cbo.Tag = cbo.Text
        
    ElseIf strDict = "���㷽ʽ" Then
        If glngSys Like "8??" Then
            MsgBox "��Ա������û�п��õĽ��㷽ʽ�����ܷ�����" & vbCrLf & _
                "���ȵ����㷽ʽ���������û�Ա���Ľ��㷽ʽ��", vbInformation, gstrSysName
        Else
            MsgBox "ҽ�ƿ�����û�п��õĽ��㷽ʽ��ֻ��ʹ�ü��ʷ�ʽ������" & vbCrLf & _
                "Ҫʹ�ý��㷢��,���ȵ����㷽ʽ���������þ��￨���㷽ʽ��", vbInformation, gstrSysName
            chk����.value = 1: chk����.Enabled = False: cbo.Enabled = False
            chk����.Tag = 1
        End If
    End If
    ReadDict = True
    If GetWidth(cbo.hWnd) * Screen.TwipsPerPixelX < lngMaxW Then zlControl.CboSetWidth cbo.hWnd, lngMaxW / Screen.TwipsPerPixelX
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub lbl����_Click()
    If mCardType.bln���￨ = False Then Exit Sub
    If mEditType = Cr_�˿� Then Exit Sub '�����:57962
    
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    If mEditType = Cr_���� Or mEditType = Cr_�󶨿� Then Exit Sub
    
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = gcnOracle
    End If

    If Not mobjICCard Is Nothing Then
        txt����.Text = mobjICCard.Read_Card()
        If txt����.Text <> "" Then
            mblnICCard = True
            Call CheckFreeCard(txt����.Text)
        End If
    End If
End Sub
Private Sub CheckFreeCard(ByVal strCard As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��һ��ͨģʽ�µĿ��ţ��ϸ����Ʊ��ʱ�� ����Ƿ���Ʊ�����÷�Χ�ڣ���Χ֮��Ŀ����շ�
    '����:���˺�
    '����:2011-07-05 08:53:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If txt����.Visible = False Then Exit Sub
    If Not mCardType.rsҽ�ƿ��� Is Nothing And Val(txt����.Text) = 0 Then  '�Ȼָ�
        txt����.Text = Format(IIf(mCardType.bln���, mCardType.rsҽ�ƿ���!ȱʡ�۸�, mCardType.rsҽ�ƿ���!�ּ�), "0.00")
        lbl����.Tag = txt����.Text
    End If
    '142204:���ϴ���2020/6/18���������IDʱ��Ҫ���뿨���ID�������Ƿ���ݷѱ���۽����Ƿ����ηѱ��йأ����Ƿ����޹�
    If mCardType.blnOneCard And mCardType.bln�ϸ���� Then
        mCardType.lng����ID = CheckUsedBill(5, IIf(mCardType.lng����ID > 0, mCardType.lng����ID, mCardType.lng��������), strCard, mlngCardTypeID)
        If mCardType.lng����ID <= 0 Then txt����.Text = "0.00": lbl����.Tag = txt����.Text
    End If
    If Not mCardType.rsҽ�ƿ��� Is Nothing And Val(txt����.Text) <> 0 Then
        If Nvl(mCardType.rsҽ�ƿ���!���ηѱ�, 0) <> 1 Then
            txt����.Text = Format(GetActualMoney(NeedName(cbo�ѱ�.Text), mCardType.rsҽ�ƿ���!������ĿID, IIf(mCardType.bln���, mCardType.rsҽ�ƿ���!ȱʡ�۸�, mCardType.rsҽ�ƿ���!�ּ�), mCardType.rsҽ�ƿ���!�շ�ϸĿID), "0.00")
            lbl����.Tag = txt����.Text
        End If
    End If
End Sub
Private Function Select��Լ��λ(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ���Լ��λ
    '����:���˺�
    '����:2011-07-05 09:34:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, vRect As RECT, bytStyle As Byte
    Dim strWhere As String, strKey As String, blnCancel As Boolean
    
    bytStyle = 1: strWhere = "": strKey = GetMatchingSting(strInput)
    If strInput <> "" Then
        bytStyle = 0
        strWhere = " And ĩ��=1 and (���� like upper([1]) or ���� like [1] or ���� like upper([1]) )"
    End If
    strSQL = "" & _
    "   Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��  " & _
    "   From  ��Լ��λ" & _
    "   Where (����ʱ�� IS NULL OR TO_CHAR(����ʱ��, 'yyyy-MM-dd') = '3000-01-01') " & _
        strWhere & _
    "       Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
    vRect = GetControlRect(txt������λ.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "��Լ��λѡ��", 1, "", "��ѡ���˵ĺ�Լ��λ", False, False, True, vRect.Left, vRect.Top, txt������λ.Height, blnCancel, False, True, strKey)
    If blnCancel Then
        If txt������λ.Enabled And txt������λ.Visible Then txt������λ.SetFocus
        zlControl.TxtSelAll txt������λ
        Set rsTemp = Nothing: Exit Function
    End If
    
    lbl������λ.Tag = ""
    If Not rsTemp Is Nothing Then
        txt������λ.Text = rsTemp!����
        lbl������λ.Tag = rsTemp!id
        txt��λ�绰.Text = Trim(rsTemp!�绰 & "")
        txt��λ������.Text = Trim(rsTemp!�������� & "")
        txt��λ�ʻ�.Text = Trim(rsTemp!�ʺ� & "")
    End If
    If txt������λ.Enabled And txt������λ.Visible Then txt������λ.SetFocus
    zlCommFun.PressKey vbKeyTab
    Select��Լ��λ = True
End Function
Private Function Select����(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ������
    '����:���˺�
    '����:2011-07-05 09:34:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, vRect As RECT, bytStyle As Byte
    Dim strWhere As String, strKey As String, blnCancel As Boolean
    
    bytStyle = 0: strWhere = "": strKey = GetMatchingSting(strInput)
    If strInput <> "" Then
        strWhere = "  And  (���� like upper([1]) or ���� like [1] or ���� like upper([1]))  "
    End If
    strSQL = "" & _
    "   Select ���� as ID,����,����,���� " & _
    "   From ����" & _
    "   Where Nvl(����,0)<3 " & strWhere
    vRect = GetControlRect(txt����.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "����ѡ��", 1, "", "��ѡ���˵�����", False, False, True, vRect.Left, vRect.Top, txt����.Height, blnCancel, False, True, strKey)
    If blnCancel Then
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        zlControl.TxtSelAll txt����
        Set rsTemp = Nothing: Exit Function
    End If
    lbl����.Tag = ""
    If Not rsTemp Is Nothing Then
        txt����.Text = rsTemp!����
        lbl����.Tag = rsTemp!����
    End If
    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    zlCommFun.PressKey vbKeyTab
    Select���� = True
End Function

Private Function Select����(ByVal objCtrl As Control, ByVal objCtrlTag As Control, _
    ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�����
    '����:���˺�
    '����:2011-07-05 09:34:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, vRect As RECT, bytStyle As Byte
    Dim strWhere As String, strKey As String, blnCancel As Boolean
    bytStyle = 0: strWhere = "": strKey = GetMatchingSting(strInput)
    
    If strInput <> "" Then
        strSQL = "" & _
        "   Select ���� as ID,����,����,���� " & _
        "   From ����" & _
        "   Where     (���� like upper([1]) or ���� like [1] or ���� like upper([1]) )"
    Else
        bytStyle = 2
        strSQL = "" & _
        "   Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
        "           Substr(����,1,2) as ����  " & _
        "   From ����" & _
        "   Union All" & _
        "   Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
        "   From ����  " & _
        "   Order by ����"
    End If
    vRect = GetControlRect(objCtrl.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "�����ص�ѡ��", 1, "", "��ѡ���˵ĳ����ص�", False, False, True, vRect.Left, vRect.Top, objCtrl.Height, blnCancel, False, True, strKey)
    If blnCancel Then
        If objCtrl.Enabled And objCtrl.Visible Then objCtrl.SetFocus
        zlControl.TxtSelAll objCtrl
        Set rsTemp = Nothing: Exit Function
    End If
    objCtrlTag.Tag = ""
    If Not rsTemp Is Nothing Then
        objCtrl.Text = rsTemp!����
        objCtrlTag.Tag = rsTemp!����
    End If
    If objCtrl.Enabled And objCtrl.Visible Then objCtrl.SetFocus
    zlCommFun.PressKey vbKeyTab
    Select���� = True
End Function
Private Sub LoadCardFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؿ���
    '����:���˺�
    '����:2011-07-06 17:24:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mCardType.rsҽ�ƿ��� Is Nothing Then
        txt����.Text = ""
        Exit Sub
    End If
    If mCardType.rsҽ�ƿ���.RecordCount = 0 Then
        txt����.Text = ""
        Exit Sub
    End If
    With mCardType.rsҽ�ƿ���
        mCardType.bln��� = Val(Nvl(!�Ƿ���)) = 1
        mCardType.dblӦ�ս�� = Format(IIf(mCardType.bln���, !ȱʡ�۸�, !�ּ�), "0.00")
        mCardType.dblʵ�ս�� = mCardType.dblӦ�ս��
        If Nvl(!���ηѱ�, 0) <> 1 And mCardType.dblӦ�ս�� <> 0 Then
            mCardType.dblʵ�ս�� = Format(GetActualMoney(NeedName(cbo�ѱ�.Text), !������ĿID, mCardType.dblӦ�ս��, !�շ�ϸĿID), "0.00")
        End If
        txt����.Locked = Not mCardType.bln���
        txt����.TabStop = mCardType.bln���
        If mCardType.bln��� And Val(txt����.Text) = 0 Or Not mCardType.bln��� Then
            txt����.Text = Format(mCardType.dblʵ�ս��, "0.00")
            Call txt���_Change
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function IsCheckҽ�ƿ�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ�ƿ��������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-06 17:44:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCard As String, strICCard As String
    
    strCard = UCase(txt����.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    '-----------------------------------------------------------------------------------------------------------------
    '1.��۽����
    If (mEditType = Cr_���� Or mEditType = Cr_����) And mCardType.bln���ƿ� = True Then
        If Not mCardType.rsҽ�ƿ��� Is Nothing Then
            If mCardType.rsҽ�ƿ���!�Ƿ��� = 1 Then
                '70595:������,2014-03-04,����δ����ʱ��������
                If mCardType.rsҽ�ƿ���!�ּ� <> 0 And Abs(CCur(Val(txt����.Text))) > Abs(mCardType.rsҽ�ƿ���!�ּ�) Then
                    MsgBox mCardType.str������ & "�Ŀ��Ѿ���ֵ���ܴ�������޼ۣ�" & Format(Abs(mCardType.rsҽ�ƿ���!�ּ�), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus:  Exit Function
                End If
                
                If mCardType.rsҽ�ƿ���!ԭ�� <> 0 And Abs(CCur(Val(txt����.Text))) < Abs(mCardType.rsҽ�ƿ���!ԭ��) Then
                    MsgBox mCardType.str������ & "�Ŀ��Ѿ���ֵ����С������޼ۣ�" & Format(Abs(mCardType.rsҽ�ƿ���!ԭ��), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
                End If
            End If
        End If
        If mCardType.bln���ƿ� = True Then
            If cbo֧����ʽ.Visible And txt����.Text <> "" And cbo֧����ʽ.Enabled And cbo֧����ʽ.ListIndex = -1 Then
                MsgBox "��ȷ��" & mCardType.str������ & "�Ľɿ���㷽ʽ��", vbExclamation, gstrSysName
                If cbo֧����ʽ.Enabled And cbo֧����ʽ.Visible Then cbo֧����ʽ.SetFocus: Exit Function
            End If
        End If
    
    End If
    
    '104238:���ϴ���2017/2/15����鿨���Ƿ����㷢����������
    If txt����.Text <> "" And Len(txt����.Text) <> mCardType.lng���ų��� And Not mCardType.bln�ϸ���� Then
        Select Case mCardType.byt��������
            Case 0
                MsgBox "����Ŀ���С��" & mCardType.str������ & "�趨�Ŀ��ų��ȣ����������룡", vbExclamation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                    Exit Function
            Case 2
                If MsgBox("����Ŀ���С��" & mCardType.str������ & "�趨�Ŀ��ų��ȣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                    Exit Function
                End If
        End Select
    End If
    '108779:���ϴ�,2017/5/8,�������Ϊ�գ�ֻ��������������
    If txt����.Text <> "" And txtPass.Text <> "" And txtPass.Visible Then
        Select Case mCardType.int���볤������
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCardType.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "�����������" & mCardType.int���볤�� & "λ", vbOKOnly + vbInformation
                txtPass.Text = "": txtAudi.Text = ""
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCardType.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "�����������" & Abs(mCardType.int���볤������) & "λ����.", vbOKOnly + vbInformation
                txtPass.Text = "": txtAudi.Text = ""
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        End Select
    End If
    
    If txtPass.Text <> txtAudi.Text And txt����.Text <> "" Then
        MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus: Exit Function
    End If

    If mEditType = Cr_�󶨿� Or mEditType = Cr_���� Or mEditType = Cr_���� Or mEditType = Cr_���� Then
        If Trim(txt����.Text) = "" Then
            MsgBox "��ˢ����������￨�ţ�", vbExclamation, gstrSysName
            If txt����.Enabled And txt����.Enabled Then txt����.SetFocus
            Exit Function
        End If
    End If
    
    If txt����.Text <> "" And (mEditType = Cr_���� Or mEditType = Cr_���� Or mEditType = Cr_����) Then
        '����ǰ�����￨�Ƿ��У��Ƿ��ڷ�Χ��
        If CheckBILL(txt����.Text) = False Then Exit Function
    End If
    
    IsCheckҽ�ƿ� = True
End Function
 Private Sub SetCardEditEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ�����ؿؼ���Enable����
    '����:���˺�
    '����:2011-07-07 00:12:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    Select Case mEditType
    Case Cr_����, Cr_����, Cr_����, Cr_�󶨿�
        blnEdit = Trim(txt����.Text) <> ""
        If chkCancel.value = 1 Then Exit Sub
    Case Else
        Exit Sub
    End Select
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    lbl����.Enabled = txtPass.Enabled: lbl��֤.Enabled = blnEdit
    txtPass.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    txtAudi.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    If Not (mEditType = Cr_���� Or mEditType = Cr_����) Then Exit Sub
    blnEdit = True
    If mCardType.rsҽ�ƿ��� Is Nothing Then
        blnEdit = False
    ElseIf mCardType.rsҽ�ƿ���.State <> 1 Then
        blnEdit = False
    ElseIf mCardType.rsҽ�ƿ���.RecordCount = 0 Then
        blnEdit = False
    End If
    'ֻ�з����Ͳ����Ŵ��ڿ���
    txt����.Enabled = blnEdit: cbo֧����ʽ.Enabled = blnEdit And chk����.value = 0
    chk����.Enabled = blnEdit
    txt����.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    cbo֧����ʽ.BackColor = IIf(cbo֧����ʽ.Enabled, &H80000005, &H8000000F)
    txt�ϼ�.Enabled = blnEdit And chk����.value = 0
    txt�ϼ�.BackColor = IIf(txt�ϼ�.Enabled, &H80000005, &H8000000F)
    txt���.Enabled = blnEdit And chk����.value = 0
    txt���.BackColor = IIf(txt���.Enabled, &H80000005, &H8000000F)
End Sub
Private Sub CheckInputPassWord(KeyAscii As Integer, Optional ByVal blnOnlyNum As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2011-07-07 00:40:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If blnOnlyNum Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
       If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
            If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                 If InStr(1, "!@#$%^&*()_+-=><?,:;~`./", Asc(KeyAscii)) = 0 Then KeyAscii = 0
            End If
       End If
    End If
End Sub
Private Sub SearchCombox(cbo As ComboBox, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ�����ָ������Ŀֵ
    '����:���˺�
    '����:2011-07-07 00:53:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngIdx As Long
    lngIdx = zlControl.CboMatchIndex(cbo.hWnd, KeyAscii)
    If lngIdx = -1 And cbo.ListCount > 0 Then lngIdx = 0
    cbo.ListIndex = lngIdx
End Sub
Private Function CheckExistsMCNO(ByVal strMCNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ���Ƿ����
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-07 03:08:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 1 From ������Ϣ Where ҽ���� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strMCNO)
    If rsTemp.RecordCount > 0 Then
        MsgBox "����,�����ҽ�����Ѵ���!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlCheckMCOutMode(ByVal lng���� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���������Ƿ����ҽ��
    '���:lng����
    '����:�����ҽ��,����True
    '����:���˺�
    '����:2011-07-07 02:35:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    strSQL = "Select 1 From ������� Where ���=1 And ���=[1]"
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����)
    zlCheckMCOutMode = rsTemp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOldAcademic(ByVal dt�������� As Date, ByVal str���䵥λ As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ�ĳ������ں����䵥λ�����������ϵ�����ֵ
    '����:����
    '����:���˺�
    '����:2011-07-07 03:21:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim dtCurrDate As Date, lngOld As Long, strInterval As String
    If dt�������� = CDate(0) Or InStr(" ������", str���䵥λ) < 2 Then Exit Function
    dtCurrDate = zlDatabase.Currentdate
    strInterval = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
    lngOld = DateDiff(strInterval, dt��������, dtCurrDate)
    If DateAdd(strInterval, lngOld, dt��������) > dtCurrDate Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function
Private Function SimilarIDs(str���� As String, str���� As String, dat�������� As Date, str�Ա� As String, str���� As String, str���֤�� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ����������Ϣ
    '���:
    '����:
    '����:���Ƽ�¼�Ĳ���ID��,��"234,235,236"
    '����:���˺�
    '����:2011-07-07 03:34:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String, i As Integer
    On Error GoTo errH
    strSQL = _
        " Select ����ID,�����,סԺ��,Nvl(���֤��,'δ�Ǽ�') ���֤��,Nvl(��ͥ��ַ,'δ�Ǽ�') ��ַ,To_Char(�Ǽ�ʱ��,'YYYY-MM-DD') �Ǽ�ʱ�� " & _
        " From ������Ϣ Where (����=[1] And ����=[2] And �Ա�=[3] And ����=[4]" & _
        " And ��������=[6]) Or ���֤��=[5] " & _
        " Order by ����ID Desc"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", str����, str����, str�Ա�, str����, str���֤��, CDate(Format(dat��������, "YYYY-MM-DD")))
    For i = 1 To rsTemp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTemp!����ID & ",�����:" & Nvl(rsTemp!�����, "��") & ",סԺ��:" & Nvl(rsTemp!סԺ��, "��") & ",���֤��:" & rsTemp!���֤�� & ",��ַ:" & rsTemp!��ַ & ",�Ǽ�����:" & rsTemp!�Ǽ�ʱ��
        rsTemp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ExistClinicNO(str����� As String, Optional lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ��������Ƿ��Ѿ����������ݿ���
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-07 03:40:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    On Error GoTo errH
    strSQL = "Select ����ID,����� From ������Ϣ Where �����=[1] And ����ID<>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���������Ƿ����", Val(str�����), lng����ID)
    If rsTemp.RecordCount > 0 Then ExistClinicNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function
Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
         If Not mCardType.rsҽ�ƿ��� Is Nothing Then
            .AddNew
            !�շ���� = mCardType.rsҽ�ƿ���!�շ����
            !��� = StrToNum(txt����.Text)
            .Update
        End If
        If Val(txt���.Text) > 0 And IDKindPayMode.IDKind = 2 Then
            .AddNew
            !�շ���� = "Ԥ��"
            !��� = StrToNum(txt���.Text)
            .Update
        End If
    End With
    zlGetClassMoney = True
End Function

Private Function SetBrushObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-10 13:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, bln���ѿ� As Boolean, lngIndex As Long
    If mCurPayMoney.lngҽ�ƿ����ID = 0 Then SetBrushObject = True: Exit Function
    
    Set mobjCardObject = zlGetClsCardObject(mCurPayMoney.lngҽ�ƿ����ID, mCurPayMoney.bln���ѿ�)
    If mobjCardObject Is Nothing Then
        MsgBox "ע��:" & vbCrLf & "   δ�ҵ���ص������ӿ�,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not mobjCardObject.InitCompents Then
        If mobjCardObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
              Exit Function
        End If
        mobjCardObject.InitCompents = True
    End If
    SetBrushObject = True
End Function
Private Function ReadCardNo(ByVal strCardNo As String, ByVal intFlag As Integer) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����֤���￨�˿�����һ���Լ�ˢ��ȡ��
    '���:strCardNo-����
    '        intFlag ��־ 1 ��֤ 2 ȡ��
    '����:
    '����:-1-�ɹ�;0-ʧ��;1-�ü�¼������
    '����:���˺�
    '����:2011-07-12 17:08:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim lng����ID As Long, str���ݺ� As String, strPassWord As String, strErrMsg As String
    Dim lng�����ID As String
    
    Err = 0: On Error GoTo errH:
    ReadCardNo = 0
    If GetPatiID(mlngCardTypeID, strCardNo, False, lng����ID, strPassWord, strErrMsg) = False Then
        If lng����ID = 0 Then ReadCardNo = 1
        Exit Function
    End If
    If lng����ID = 0 Then ReadCardNo = 1: Exit Function
    lblˢ����֤.Tag = strCardNo
    If intFlag = 1 Then
        ReadCardNo = -1
        rsTmp.Close
        Exit Function
    End If
    If mEditType = Cr_���� Then
        If Not mrsInfo Is Nothing Then
            If Val(Nvl(mrsInfo!����ID)) <> lng����ID Then
                If GetPatient("-" & lng����ID) = False Then
                    ReadCardNo = 1: Exit Function
                End If
            End If
        Else
            If GetPatient("-" & lng����ID) = False Then
                ReadCardNo = 1: Exit Function
            End If
        End If
        Call LoadPatiInfor
        txtˢ������.Text = strCardNo: lblˢ����֤.Tag = strCardNo
        '�����:50893
        txtԭ������.Tag = strPassWord
        ReadCardNo = -1
        Exit Function
    End If
     If mEditType = Cr_��ʧ Then
        txtˢ������.Text = strCardNo: lblˢ����֤.Tag = strCardNo
        ReadCardNo = -1
        Exit Function
     End If
     
    If mCardType.str������ = "���￨" Then
        lng�����ID = mlngCardTypeID
    End If
    '��ȡ���￨�ڷ����е�No
    strSQL = _
    " Select A.NO" & _
    " From סԺ���ü�¼ A" & _
    " Where A.��¼����=5   And A.ʵ��Ʊ��=[1] " & _
    "           And A.����ID = [2]  And A.��¼״̬=1 And nvl(A.����,[3])=[4] "
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, lng����ID, CStr(lng�����ID), CStr(mlngCardTypeID))
    If rsTmp.EOF Then ReadCardNo = 1: Exit Function
    str���ݺ� = IIf(IsNull(rsTmp!NO), "", rsTmp!NO)
    If mEditType = Cr_�˿� Or chkCancel.value = 1 Then
        If mParaData.int�˿�ģʽ = 2 And Trim(cboNO.Text) = "" Then
            MsgBox "ע��:" & vbCrLf & "  �˿�ʱ,���������뵥��,��ˢ��!", vbInformation + vbOKOnly, gstrSysName
            
            Exit Function
        Else
            If str���ݺ� <> Trim(cboNO.Text) And Trim(cboNO.Text) <> "" Then
                MsgBox "��ǰˢ���ĵ��ݺ���ָ���ĵ��ݺŲ���,�����˿�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                If Nvl(mrsInfo!����ID, 0) <> lng����ID Then
                    MsgBox "��ǰ���������еĿ�����,�����˿�", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If

    If ReadBill(str���ݺ�) = -1 Then
        ReadCardNo = -1
        rsTmp.Close
        Exit Function
    End If
    rsTmp.Close
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBill(strNo As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ɵ��ݺŶ�ȡ����ʾ���￨���ż�¼
    '���:strNO-���ݺ�
    '����:
    '����:-1-�ɹ�;0-ʧ��;1-�ü�¼������;2-�ü�¼�Ѿ�����(��mblnViewCancel=Falseʱ��Ч)
    '����:���˺�
    '����:2011-07-12 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strFullNO As String
    Dim strSQL As String, str���㷽ʽ As String, intIndex As Integer
    Dim strժҪ As String
    On Error GoTo errH
    cmdOK.Enabled = True
    strFullNO = GetFullNO(strNo, 16)
    '��Ϊ���￨���õĽ���ID�����Ǽ��ʷ��������ʱ������ID,
    '������Ԥ����¼����ʱһ��Ҫ�Ӽ�¼����=5����
    '�����:50891
    gstrSQL = _
        "   Select A.NO,A.����ID,A.����,A.�Ա�,A.����,A.ʵ��Ʊ��,A.���ӱ�־,A.��¼״̬,A.���� , " & _
        "           A.ʵ�ս��,A.����Ա����,A.����ʱ��,B.����֤��,A.����ID,A.ժҪ, " & _
        "          C.���㷽ʽ,C.�����ID,C.����,C.����˵��,C.�������,C.���㿨���,C.������ˮ��,D.Ԥ����� " & _
        "   From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A ," & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ C,������Ϣ B,������� D " & _
        "   Where A.����ID=C.����ID(+) And C.��¼����(+)=5 And A.����ID=D.����ID(+) and C.NO(+)=[1] And A.��¼����=5 And A.����ID=B.����ID And A.NO=[1] And D.����(+)=1 " & _
            IIf(mEditType = Cr_��ѯ, "And A.��¼״̬=[2] ", "")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO, mint��¼״̬)
    If rsTemp.EOF Then ReadBill = 1: Exit Function
    If mEditType <> Cr_��ѯ And (rsTemp!��¼״̬ = 3 Or rsTemp!��¼״̬ = 2) Then
        ReadBill = 2: Exit Function
    End If
    Call GetPatient("-" & rsTemp!����ID)
    Call LoadPatiInfor
    '�����:73063
    lblԤ�����.Caption = "Ԥ�����:" & Nvl(rsTemp!Ԥ�����, "0") & "Ԫ"
    Call SetCtrlMove '���²��ֵ�ǰ����ؼ�
    
    cboNO.Text = rsTemp!NO
    cboNO.Tag = rsTemp!NO
    txtPatient.Text = rsTemp!����
    txtPatient.PasswordChar = ""
    strժҪ = Nvl(rsTemp!ժҪ)
    
    Call zlControl.CboLocate(cbo�Ա�, Nvl(mrsInfo!�Ա�))
    If cbo�Ա�.ListIndex = -1 And Not IsNull(rsTemp!�Ա�) Then
        cbo�Ա�.AddItem mrsInfo!�Ա�, 0
        cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
    End If
    Call LoadOldData("" & rsTemp!����, txt����, cbo���䵥λ)
    mlngBillCardTypeID = Val(Nvl(rsTemp!����))
    Set mcolBillBalance = New Collection
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID
    mcolBillBalance.Add Array(Val(Nvl(rsTemp!�����id)), Trim(Nvl(rsTemp!����)), IIf(Val(Nvl(rsTemp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTemp!������ˮ��)), Trim(Nvl(rsTemp!����˵��)), strNo, Val(Nvl(rsTemp!����id)), Nvl(rsTemp!���㷽ʽ)), strNo
    If IsNull(rsTemp!���㷽ʽ) Then
        chk����.value = Checked
    Else
        str���㷽ʽ = zlGet֧����ʽ(Nvl(rsTemp!�����id, Nvl(rsTemp!���㿨���, 0)), rsTemp!���㷽ʽ)
        
        chk����.value = Unchecked
        cbo֧����ʽ.ListIndex = GetCboIndex(cbo֧����ʽ, rsTemp!���㷽ʽ)
        
        If cbo֧����ʽ.ListIndex = -1 Then
            mcolPayMode.Add Array("", Split(str���㷽ʽ, "|")(0), 0, 0, 0, 0, Split(str���㷽ʽ, "|")(1), 0, 0, Split(str���㷽ʽ, "|")(2), Split(str���㷽ʽ, "|")(3))
            cbo֧����ʽ.AddItem Split(str���㷽ʽ, "|")(0)
            cbo֧����ʽ.ItemData(cbo֧����ʽ.NewIndex) = Val(Split(str���㷽ʽ, "|")(4))
            cbo֧����ʽ.ListIndex = cbo֧����ʽ.NewIndex
            intIndex = cbo֧����ʽ.NewIndex + 1
        Else
            intIndex = cbo֧����ʽ.ListIndex + 1
        End If
        cbo֧����ʽ.Tag = ""
    End If
    
    txt����.Text = IIf(IsNull(rsTemp!ʵ��Ʊ��), "", rsTemp!ʵ��Ʊ��)
    txtPass.Text = IIf(IsNull(rsTemp!����֤��), "", rsTemp!����֤��)
    txtAudi.Text = txtPass.Text
    txt����.Text = Format(rsTemp!ʵ�ս��, "0.00")
    txt����Ա.Text = rsTemp!����Ա����
    txtDate.Text = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm")
    
    '����:48249
    If mEditType = Cr_�˿� Or chkCancel.value = 1 Then
        mlng����ID = 0
        mlng����ID = rsTemp!����ID
        '116278:���ϴ�,2017/12/15����֧�ֲ����˵����������˺ű���ͬʱ�˿�,��ʱ�������ѿ�
        If str���㷽ʽ <> "" And Nvl(rsTemp!�����id) <> 0 And Nvl(rsTemp!���㿨���, 0) = 0 Then
            If Val(Split(str���㷽ʽ & "||||", "|")(2)) = 0 Then
                strSQL = "Select 1 From ������ü�¼ Where ��¼����=4 And ��¼״̬=1 And (����ID,�Ǽ�ʱ��) = " & _
                        " (Select ����ID,�Ǽ�ʱ�� From סԺ���ü�¼ Where ��¼����=5 And NO=[1] And Rownum=1)"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", cboNO.Text)
                If Not rsTemp.EOF Then
                    MsgBox "��ǰ������Һŷ�һ����ȡ�ģ��뵽�ҺŴ�����Һŷ�һ���ˡ�", vbInformation + vbOKOnly, gstrSysName
                    cmdOK.Enabled = False: ReadBill = 3: Exit Function
                End If
            End If
        End If
        
        '90143,���ϴ�,2016/1/25,�������ժҪ,��Ҫ��ȡ������ü�¼
        If strժҪ <> "" Then
            strSQL = "Select NO,��¼״̬ from " & IIf(mblnNOMoved, "H", "") & "������ü�¼ where ����ID=[1] and ��¼����=1 and ժҪ=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, cboNO.Text)
            If rsTemp.RecordCount > 0 Then
                If Nvl(rsTemp!��¼״̬, 0) = 1 Then
                    MsgBox "��ǰ���ѻ����շѣ����˿����շѴ����˷ѡ�", vbInformation + vbOKOnly, gstrSysName
                    cmdOK.Enabled = False: ReadBill = 3: Exit Function
                End If
            End If
        End If
    End If
    txt�ϼ�.Text = Format(IIf(txt����.Visible, Val(txt����.Text), 0), "0.00")
    txt�ϼ�.Tag = txt�ϼ�.Text
    ReadBill = -1
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlGet�����() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����
    '����:�����
    '����:���˺�
    '����:2011-07-28 08:39:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbln�Զ������ Then zlGet����� = zlDatabase.GetNextNo(3)
End Function
Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str���� As String
    Dim dblMoney As Double, cllBalance As Collection
    Dim blnErrCount As Boolean
    Dim frmInput As frmInputPass
    
    If Not (mEditType = Cr_���� Or mEditType = Cr_����) Then CheckBrushCard = True: Exit Function
    If SetBrushObject = False Then Exit Function
    
    On Error GoTo errHandle
    If mCurPayMoney.lngҽ�ƿ����ID = 0 Then CheckBrushCard = True: Exit Function
    dblMoney = IIf(IDKindPayMode.IDKind = 2, StrToNum(txt�ϼ�.Text), StrToNum(txt�ϼ�.Tag))
    Call zlGetClassMoney(rsMoney)
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text

    Set cllBalance = Nothing '57682
    
    '���ѿ�
    If mCurPayMoney.bln���ѿ� And mobjCardObject.���ƿ� Then
        Err = 0: On Error Resume Next
        If IsEmpty(cllBalance) Then   '57682
            Set cllBalance = Nothing
        End If
        blnErrCount = cllBalance.count
        If Err <> 0 Then
            Set cllBalance = Nothing
            Err = 0: On Error GoTo 0
        End If
        '����:����ָ��֧�����,����ˢ������
        '���:rsClassMoney:�շ����,���
        '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
        '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
        '58322
        '115668:���ϴ�,2017/10/25,�½�ʵ���������ѿ�֧��
        Set frmInput = New frmInputPass
        CheckBrushCard = frmInput.zlBrushPay(Me, mlngModule, mobjCardObject, rsMoney, _
                mCurPayMoney.lngҽ�ƿ����ID, True, txtPatient.Text, NeedName(cbo�Ա�.Text), str����, _
                dblMoney, mCurPayMoney.strˢ������, mCurPayMoney.strˢ������, False, True, False, True, cllBalance, Val(txt���.Text) > 0 And IDKindPayMode.IDKind = 2)
        Unload frmInput
        Set frmInput = Nothing
        Exit Function
    End If
    
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
    If mobjCardObject.CardObject.zlBrushCard(Me, mlngModule, mCurPayMoney.lngҽ�ƿ����ID, _
            txtPatient.Text, NeedName(cbo�Ա�.Text), str����, dblMoney, mCurPayMoney.strˢ������, mCurPayMoney.strˢ������) = False Then Exit Function
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If mobjCardObject.CardObject.zlPaymentCheck(Me, mlngModule, mCurPayMoney.lngҽ�ƿ����ID, _
         mCurPayMoney.strˢ������, dblMoney, "", "") = False Then Exit Function
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlInterfacePrayMoney(ByRef cllPro As Collection, ByRef cllThreeSwap As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:cllPro-�޸�������������
    '        cll��������-����������������
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim dblMoney As Double
    If mCurPayMoney.lngҽ�ƿ����ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    If Val(txt�ϼ�.Tag) <= 0 Then zlInterfacePrayMoney = True: Exit Function
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
    dblMoney = IIf(IDKindPayMode.IDKind = 2, StrToNum(txt�ϼ�.Text), StrToNum(txt�ϼ�.Tag))
    If mobjCardObject.CardObject.zlPaymentMoney(Me, mlngModule, mCurPayMoney.lngҽ�ƿ����ID, mCurPayMoney.strˢ������, mCurPayMoney.lng����ID, mCurPayMoney.strNo, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    
    '����������������
    '�����:58536
    If Not mCurPayMoney.bln���ѿ� Then
        Call zlAddUpdateSwapSQL(False, mCurPayMoney.lng����ID, mCurPayMoney.lngҽ�ƿ����ID, mCurPayMoney.bln���ѿ�, mCurPayMoney.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
    End If
    Call zlAddThreeSwapSQLToCollection(False, mCurPayMoney.lng����ID, mCurPayMoney.lngҽ�ƿ����ID, mCurPayMoney.bln���ѿ�, mCurPayMoney.strˢ������, strSwapExtendInfor, cllThreeSwap)
    If IDKindPayMode.IDKind = 2 And Val(StrToNum(txt���.Text)) > 0 Then
        Call zlAddUpdateSwapSQL(True, mlngԤ��ID, mCurPayMoney.lngҽ�ƿ����ID, mCurPayMoney.bln���ѿ�, mCurPayMoney.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
        Call zlAddThreeSwapSQLToCollection(True, mlngԤ��ID, mCurPayMoney.lngҽ�ƿ����ID, mCurPayMoney.bln���ѿ�, mCurPayMoney.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlShowSelectCardNo(Optional ByVal lng����ID As Long = 0, _
    Optional str���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ�����˵Ŀ���
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-16 17:12:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, vRect As RECT, rsTemp As ADODB.Recordset, blnCancel As Boolean
    strSQL = "" & _
    "   Select RowNum as ID, A.����, A.��������, A.������,B.����, B.����, B.���֤��, B.��������, B.�ֻ���, B.��ͥ�绰,B.��ͥ��ַ,B.��ϵ������,B.��ϵ�˹�ϵ " & _
    "   From ����ҽ�ƿ���Ϣ A, ������Ϣ B " & _
    "   Where A.����id = B.����id And A.�����id = [1] and A.����ID=[2]  " & IIf(str���� = "", "", " And A.����=[3]") & _
    "   Order by A.����"
    
    vRect = GetControlRect(txtˢ������.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ѡ����Ҫ��ʧ�Ŀ���", 1, "", "ѡ����Ҫ��ʧ�Ŀ���", False, False, True, vRect.Left, vRect.Top, txtˢ������.Height, blnCancel, False, True, mlngCardTypeID, lng����ID, str����)
    If blnCancel = True Then GoTo GoCancel:
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ����������Ŀ��Ż�ò���δ���п�!", vbOKOnly + vbInformation, gstrSysName
        GoTo GoCancel:
        Exit Function
    End If
    If rsTemp.State <> 1 Then GoTo GoCancel:
    txtˢ������.Text = Nvl(rsTemp!����)
    lblˢ����֤.Tag = txtˢ������.Text
    
    zlShowSelectCardNo = True
    Exit Function
GoCancel:
    txtˢ������.Text = ""
    If txtˢ������.Enabled And txtˢ������.Visible Then txtˢ������.SetFocus
    zlControl.TxtSelAll txtˢ������
End Function

Private Function zl�Ƿ��Ѱ�(str���� As String) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鿨���Ƿ��Ѿ�����
    '���:��Ҫ���Ŀ���
    '����:�󶨵Ĳ�����Ϣ
    '����:����
    '����:2012-09-5 17:12:38
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo ErrHandl:
        strSQL = "" & _
        "   Select A.����ID,A.����,A.���֤�� From ������Ϣ A,����ҽ�ƿ���Ϣ B Where A.����ID = B.����ID And B.���� = [1]"
        Set zl�Ƿ��Ѱ� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Dim strKindStr As String, blnFindDefaultCard  As Boolean
    Dim lngCurCardTypeID As Long
    
    If gobjSquare Is Nothing Then Exit Function
    lngCurCardTypeID = mlngCardTypeID
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, IDKind.IDKindStr, txtPatient)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    
    '96809:���ϴ�,2016/9/29,Ԥ���շ�
    IDKindPayMode.IDKindStr = "Ӧ��|Ӧ��|0|0|0|0|0|0|0|0|0;��ֵ|��ֵ|0|0|0|0|0|0|0|0|0"
    IDKindPayMode.IDKind = 1
    
    '72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
    If mEditType <> Cr_���� And mEditType <> Cr_�󶨿� Then Exit Function
    
'    IDKindPay.NotAutoAppendKind = True '���Զ����뿨���
    Call IDKindPay.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txt����)
    
    blnFindDefaultCard = GetValidKindStr(mlngCardTypeID)
    If mblnFromCardMgr Then
        If blnFindDefaultCard = False Then
            MsgBox "�ÿ��豸δ���ã������ܽ���" & IIf(mEditType = Cr_����, "����", "�󶨿�") & "�������뵽����������>�豸���á������ã�", vbInformation, gstrSysName
            mblnUnLoad = True: Exit Function
        End If
    End If
    
    If IDKindPay.Cards.count = 0 Then
        MsgBox "û�п����ڷ�����󶨿�����Чҽ�ƿ�������飡", vbOKOnly + vbInformation, gstrSysName
        mblnUnLoad = True: Exit Function
    End If
    
    '��λȱʡĬ�Ͽ����
    If blnFindDefaultCard Then
        If lngCurCardTypeID <> 0 Then
            IDKindPay.DefaultCardType = lngCurCardTypeID
            IDKindPay.IDKind = IDKindPay.GetKindIndex(IDKindPay.GetfaultCard.����)
        End If
    Else
        mlngCardTypeID = IDKindPay.GetfaultCard.�ӿ����
        IDKindPay.DefaultCardType = mlngCardTypeID
        IDKindPay.IDKind = IDKindPay.GetKindIndex(IDKindPay.GetfaultCard.����)
    End If
    txt����.Locked = Not IDKindPay.GetCurCard.�Ƿ�ˢ��
End Function
'��ȡĬ��IDKind����
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

 
'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case "סԺ��"
          IsCardType = IDKindCtl.GetCurCard.���� = "סԺ��"
     Case "�ֻ���"
          IsCardType = IDKindCtl.GetCurCard.���� = "�ֻ���"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.����
            Else
                If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
            End If
     End Select
End Function
                
Private Function �Ƿ��Ѿ�ǩԼ(strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ҫ�󶨵Ŀ����Ƿ��Ѿ�ǩԼ
    '���:�󶨿���
    '����:����
    '����:2012-08-31 11:32:14
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand:
    strSQL = "" & _
    "   Select Count(1) as �Ƿ�ǩԼ From ����ҽ�ƿ���Ϣ Where ����=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ���", strCardNo)
    �Ƿ��Ѿ�ǩԼ = rsTemp!�Ƿ�ǩԼ > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub InitvsDrug()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsDrug
    '��ʼ���б�����
     vsDrug.Editable = flexEDKbdMouse
    '������ͷ
        SetColumHeader vsDrug, C_ColumHeader
    End With

End Sub

Private Sub SetColumHeader(objList As Object, strColumHeader As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ͷ
    '����:objList - ���ö���,strColumHeader - �б������ַ���
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varSet As Variant
    Dim varColum As Variant
    Dim i As Long
    varSet = Split(strColumHeader, ";")
    If UBound(varSet) = 0 Then Exit Sub
    
    For i = LBound(varSet) To UBound(varSet)
        varColum = Split(varSet(i), ",")
        Select Case TypeName(objList)
            Case "VSFlexGrid"
                With objList
                    .Cols = UBound(varSet) + 1
                    .Cell(flexcpText, 0, i) = varColum(0)
                    .ColAlignment(i) = varColum(1)
                    .ColWidth(i) = varColum(2)
                    .ColHidden(i) = Not (varColum(3) = 1)
                End With
            Case Else
            '�ݲ�����
        End Select
    Next
End Sub
Private Sub vsDrug_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�����:56599
    If vsDrug.Col = 1 Then  '������Ӧ�б༭ʱ���ж��Ƿ�����������200
        With vsDrug
           If Len(.TextMatrix(vsDrug.Row, vsDrug.Col)) > 200 Then
                MsgBox "������Ӧ�����ַ���������ַ���200,������ַ������Զ��س���", vbInformation, gstrSysName
                .TextMatrix(.Row, .Col) = Mid(.TextMatrix(.Row, .Col), 1, 200)
           End If
        End With
    End If
End Sub

Private Sub vsDrug_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    Dim strFliter As String
    On Error GoTo ErrHandl:
    If KeyAscii <> 13 Then
        If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf vsDrug.Col = 0 Then
        KeyAscii = 0
        datCurr = zlDatabase.Currentdate
        strSQL = _
        " Select Distinct A.ID,A.����," & _
        " A.����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
        " And (C.���� like [1] OR A.���� like [1] OR C.���� like [1])" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
        
        strFliter = gstrLike & UCase(vsDrug.EditText) & "%"
        'ת��һ�ν���(��¼��ֻ��һ��ʱ���Զ����أ���ʱ�Ե�Ԫ��ĸ�ֵ��Ч)
        cmdSelDrug.SetFocus
        '��ȡ��ǰ�������ֵ
        vRect = GetControlRect(vsDrug.hWnd)
        vRect.Top = vRect.Top + (vsDrug.Row - 1) * 300 + 150
        vRect.Left = vRect.Left + 30
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ��", False, "����ҩ��ѡ����", "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ��", False, False, True, vRect.Left, vRect.Top, 0, True, False, True, strFliter)
        vsDrug.SetFocus
        If Not rsTemp Is Nothing Then
            vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) = rsTemp!����
            vsDrug.TextMatrix(vsDrug.Row, 2) = rsTemp!id
            If vsDrug.Rows - 1 = vsDrug.Row Then vsDrug.Rows = vsDrug.Rows + 1
        End If
    End If
    Exit Sub
ErrHandl:
    MsgBox Err.Description
End Sub

Private Sub vsDrug_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetCmdCtrlMove
End Sub
Private Sub vsDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:56599
    If KeyCode = 27 And vsDrug.Rows = 2 Then vsDrug.TextMatrix(1, 0) = "": vsDrug.Cell(flexcpData, 1, 0) = "": vsDrug.TextMatrix(1, 1) = ""
    If KeyCode = 27 And vsDrug.Rows > 2 Then vsDrug.Rows = vsDrug.Rows - 1 'Esc

End Sub

Private Sub vsDrug_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cmdSelDrug.Visible = False
End Sub

Private Sub vsDrug_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then SetCmdCtrlMove
End Sub

Private Sub vsDrug_KeyPress(KeyAscii As Integer)
    '78408:���ϴ�,2014/10/9,�����ת
    If KeyAscii = 13 Then
        If vsDrug.Col = 0 Then
             zlCommFun.PressKey vbKeyRight
        ElseIf vsDrug.Rows > vsDrug.Row + 1 Then
            vsDrug.Row = vsDrug.Row + 1
            vsDrug.Col = 0
        End If
    End If
End Sub

Private Sub vsDrug_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsDrug.Col = 0 Then SetCmdCtrlMove
    End If
End Sub
Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
    '��ʼ���б�����
     vsInoculate.Editable = flexEDKbdMouse
    '������ͷ
        SetColumHeader vsInoculate, C_InoculateHeader
    '����ѡ��ť
        .ColDataType(0) = flexDTDate
        .ColEditMask(0) = "####-##-##"
        .ColDataType(2) = flexDTDate
        .ColEditMask(2) = "####-##-##"
    End With

End Sub
Private Sub VsInoculate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�����:56599
    If Col = 1 Or Col = 3 Then '���������б༭ʱ���ж��Ƿ�����������100
        With vsInoculate
           If Len(.TextMatrix(Row, Col)) > 100 Then
                MsgBox "�������������ַ���������ַ���100,������ַ������Զ��س���", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 100)
           End If
        End With
        If Col = 3 And vsInoculate.Rows - 1 = Row And vsInoculate.TextMatrix(Row, Col) <> "" Then
                vsInoculate.Rows = vsInoculate.Rows + 1
        End If
    Else
        With vsInoculate
           If IsDate(.TextMatrix(Row, Col)) = False And .TextMatrix(Row, Col) <> "    -  -  " Then
                MsgBox "��������ڸ�ʽ���Ի�����ȷ�����ڣ�", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = ""
           ElseIf .TextMatrix(Row, Col) = "    -  -  " Then
                .TextMatrix(Row, Col) = ""
           End If
        End With
    End If
End Sub
Private Sub VsInoculate_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:56599
    If KeyCode = 27 And vsInoculate.Rows = 2 Then
        If vsInoculate.TextMatrix(1, 2) <> "    -  -  " And vsInoculate.TextMatrix(1, 3) <> "" Then
            vsInoculate.TextMatrix(1, 2) = "": vsInoculate.TextMatrix(1, 3) = ""
        Else
            vsInoculate.TextMatrix(1, 0) = "": vsInoculate.TextMatrix(1, 1) = ""
        End If
    End If
    If KeyCode = 27 And vsInoculate.Rows > 2 Then 'Esc
        If vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) <> "    -  -  " And vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) <> "" Or vsInoculate.TextMatrix(vsInoculate.Rows - 1, 3) <> "" Then
            vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) = "": vsInoculate.TextMatrix(vsInoculate.Rows - 1, 3) = ""
        Else
            vsInoculate.Rows = vsInoculate.Rows - 1
        End If
    End If
End Sub

Private Sub vsInoculate_KeyPress(KeyAscii As Integer)
    '78408:���ϴ�,2014/10/9,�����ת
    If KeyAscii = 13 Then
        If vsInoculate.Col = 3 And vsInoculate.Rows - 1 = vsInoculate.Row Then
            zlCommFun.PressKey vbKeyTab
        ElseIf vsInoculate.Col = 3 Then
            vsInoculate.Col = 0: vsInoculate.Row = vsInoculate.Row + 1
            zlCommFun.PressKey vbKeyReturn
        Else
            zlCommFun.PressKey vbKeyRight
        End If
    End If
End Sub

Public Function InoculateValid() As Boolean
    '�����56599
    Dim i As Long
    
    With vsInoculate
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" And .TextMatrix(i, 1) = "" Then
                MsgBox "�������Ʊ������룡", vbInformation, gstrSysName
                .Select i, 1
                InoculateValid = False
                Exit Function
            End If
            If .TextMatrix(i, 0) = "" And .TextMatrix(i, 1) <> "" Then
                MsgBox "�������ڱ������룡", vbInformation, gstrSysName
                .Select i, 0
                InoculateValid = False
                Exit Function
            End If
            If .TextMatrix(i, 2) <> "" And .TextMatrix(i, 3) = "" Then
                MsgBox "�������Ʊ������룡", vbInformation, gstrSysName
                .Select i, 3
                InoculateValid = False
                Exit Function
            End If
            If .TextMatrix(i, 2) = "" And .TextMatrix(i, 3) <> "" Then
                MsgBox "�������ڱ������룡", vbInformation, gstrSysName
                .Select i, 2
                InoculateValid = False
                Exit Function
            End If
        Next
    End With
    InoculateValid = True
End Function
Private Sub ComboBox(objcbo As ComboBox, strSet As String)
    Dim varTemp As Variant
    Dim i As Long
    varTemp = Split(strSet, ",")
    With objcbo
        For i = LBound(varTemp) To UBound(varTemp)
            .AddItem varTemp(i)
        Next
    End With
    If objcbo.ListCount <> 0 Then objcbo.ListIndex = 0
End Sub
Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsLinkMan
    '��ʼ���б�����
        .Editable = flexEDNone
    '������ͷ
        SetColumHeader vsLinkMan, C_LinkManColumHeader
    End With
    With vsOtherInfo
         .Editable = flexEDNone
    '������ͷ
        SetColumHeader vsOtherInfo, C_OtherInfoColumHeader
    End With
End Sub
Private Sub InitCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ComBox�ؼ�
    '����:56599
    '����:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '66743:������,2013-11-25,Ѫ����RHĬ��ֵ������
    'ComboBox cboBloodType, C_Ѫ��
    zlComboxLoadFromSQL "Select ����,����,ȱʡ��־ From Ѫ��", cboBloodType
    ComboBox cboBH, C_BH
    If cboBH.ListCount <> 0 Then cboBH.ListIndex = -1
End Sub

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
Private Sub cmdMedicalWarning_Click()
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim strTemp As String
    
    vRect = GetControlRect(txtMedicalWarning.hWnd)
    strSQL = "" & _
    "       Select ���� as ID,����,���� From ҽѧ��ʾ Where ���� Not Like '����%'"
    Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ҽѧ��ʾ", False, "", "", False, False, False, vRect.Left, vRect.Top - 180, 500, True, False, True)
    If Not rsTemp Is Nothing Then
      While rsTemp.EOF = False
        strTemp = strTemp & ";" & rsTemp!����
        rsTemp.MoveNext
      Wend
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
End Sub
Private Sub SetDrugAllergy(str����ҩ�� As String, str������Ӧ As String, Optional lng����ID = 0, Optional ByVal lng����ID As Long, Optional ByVal byt��Ϣ����ģʽ As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���ҩ��
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str����ҩ�� Then
                    If lng����ID <> 0 And byt��Ϣ����ģʽ = 2 Then Exit Sub
                    .TextMatrix(i, 1) = str������Ӧ
                    If lng����ID <> 0 Then .TextMatrix(i, 2) = lng����ID
                    Exit Sub
                End If

            Next
        End If
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str����ҩ��
        .TextMatrix(.Rows - 1, 1) = str������Ӧ
        If lng����ID <> 0 Then .TextMatrix(.Rows - 1, 2) = lng����ID
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetInoculate(str�������� As String, str�������� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ý������
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If Format(.TextMatrix(i, j - 1), "YYYY-MM-DD") = Format(str��������, "YYYY-MM-DD") And .TextMatrix(i, j) = str�������� Then Exit Sub
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" Or .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str��������
                .TextMatrix(.Rows - 1, j + 1) = str��������
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" Or .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Private Sub SetLinkInfo(str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, Optional ByVal lng����ID As Long, Optional ByVal byt��Ϣ����ģʽ As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϵ�������Ϣ
    '����:56599
    '����:2012-12-12 09:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    With vsLinkMan
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str���� And .TextMatrix(i, 2) = str���֤�� Then
                    If lng����ID <> 0 And byt��Ϣ����ģʽ = 2 Then Exit Sub
                    .TextMatrix(i, 1) = str��ϵ: .TextMatrix(i, 3) = str�绰
                    If i = 1 Then
                        txt��ϵ�����֤��.Text = str���֤��
                        txt��ϵ������.Text = str����
                        For j = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                            If NeedName(cbo��ϵ�˹�ϵ.List(j)) = str��ϵ Then cbo��ϵ�˹�ϵ.ListIndex = j
                        Next
                        txt��ϵ�˵绰.Text = str�绰
                    End If
                    Exit Sub
                End If
            Next
        End If
        
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str����
        .TextMatrix(.Rows - 1, 1) = str��ϵ
        .TextMatrix(.Rows - 1, 3) = str�绰
        .TextMatrix(.Rows - 1, 2) = str���֤��
        If .Rows - 1 = 1 Then
            txt��ϵ�����֤��.Text = str���֤��
            txt��ϵ������.Text = str����
            For j = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                If NeedName(cbo��ϵ�˹�ϵ.List(j)) = str��ϵ Then cbo��ϵ�˹�ϵ.ListIndex = j
            Next
            txt��ϵ�˵绰.Text = str�绰
        End If
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetOtherInfo(str��Ϣ�� As String, str��Ϣֵ As String, Optional ByVal lng����ID As Long, Optional ByVal byt��Ϣ����ģʽ As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str��Ϣ�� Then
                        If lng����ID <> 0 And byt��Ϣ����ģʽ = 2 Then Exit Sub
                        .TextMatrix(i, j + 1) = str��Ϣֵ
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" Or .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str��Ϣ��
                .TextMatrix(.Rows - 1, j + 1) = str��Ϣֵ
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" Or .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Private Sub Load�����������Ϣ(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˽�������Ϣ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-12 14:55:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs����ҩ�� As Recordset
    Dim rs���߼�¼ As Recordset
    Dim rsABOѪ�� As Recordset
    Dim rsRH As Recordset
    Dim rsҽѧ��ʾ As Recordset
    Dim rs����ҽѧ��ʾ As Recordset
    Dim rs������Ϣ As Recordset
    Dim rs��ϵ�� As Recordset
    Dim rs������Ϣ As Recordset
    Dim strҽѧ��ʾ As String
    Dim str��ϵ������ As String
    Dim str��ϵ�˹�ϵ As String
    Dim str��ϵ�˵绰 As String
    Dim str��ϵ�����֤�� As String
    Dim lng��ϵ������ As Long
    Dim i As Long
    On Error GoTo ErrHandl:
    '��ȡ��Ƭ
    ReadPatPricture lng����ID
    
    If mEditType = Cr_�󶨿� Or mEditType = Cr_���� Then
        '��ȡ����ҩ��
        strSQL = "" & _
        "   Select ����ID,����ҩ��ID,����ҩ��,������Ӧ From ���˹���ҩ�� Where ����ID=[1]"
        Set rs����ҩ�� = zlDatabase.OpenSQLRecord(strSQL, "���˹���ҩ��", lng����ID)
        While rs����ҩ��.EOF = False
            SetDrugAllergy Nvl(rs����ҩ��!����ҩ��), Nvl(rs����ҩ��!������Ӧ), Nvl(rs����ҩ��!����ҩ��ID, 0)
            rs����ҩ��.MoveNext
        Wend
        '��ȡ���߼�¼
        strSQL = "" & _
        "   Select ����ID,����ʱ��,�������� From �������߼�¼ Where ����ID=[1]"
        Set rs���߼�¼ = zlDatabase.OpenSQLRecord(strSQL, "�������߼�¼", lng����ID)
        While rs���߼�¼.EOF = False
            SetInoculate Nvl(rs���߼�¼!����ʱ��), Nvl(rs���߼�¼!��������)
            rs���߼�¼.MoveNext
        Wend
        '82072:���ϴ�,2015/1/23,Ѫ�ͺ�RHȡ����ID Ϊnull�ļ�¼
        'Ѫ��
        strSQL = "" & _
        "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='Ѫ��' And ����ID Is NULL "
        Set rsABOѪ�� = zlDatabase.OpenSQLRecord(strSQL, "ABOѪ��", lng����ID)
        While rsABOѪ��.EOF = False
            For i = 0 To cboBloodType.ListCount - 1
                '76314,���ϴ���2014-08-06��������Ϣ��ȷ��ȡ
                If NeedName(cboBloodType.List(i), ".") = NeedName(Nvl(rsABOѪ��!��Ϣֵ)) Then cboBloodType.ListIndex = i
            Next
            rsABOѪ��.MoveNext
        Wend
        'RH
        strSQL = "" & _
        "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='RH' And ����ID Is NULL "
        Set rsRH = zlDatabase.OpenSQLRecord(strSQL, "RH", lng����ID)
        While rsRH.EOF = False
            For i = 0 To cboBH.ListCount - 1
                If cboBH.List(i) = Nvl(rsRH!��Ϣֵ) Then cboBH.ListIndex = i
            Next
            rsRH.MoveNext
        Wend
        'ҽѧ��ʾ
        strSQL = "" & _
        "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='ҽѧ��ʾ' "
        Set rsҽѧ��ʾ = zlDatabase.OpenSQLRecord(strSQL, "ҽѧ��ʾ", lng����ID)
        While rsҽѧ��ʾ.EOF = False
            strҽѧ��ʾ = strҽѧ��ʾ & ";" & Nvl(rsҽѧ��ʾ!��Ϣֵ)
            rsҽѧ��ʾ.MoveNext
        Wend
        If strҽѧ��ʾ <> "" Then strҽѧ��ʾ = Mid(strҽѧ��ʾ, 2)
        txtMedicalWarning.Text = strҽѧ��ʾ
        '����ҽѧ��ʾ
        strSQL = "" & _
        "  Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='����ҽѧ��ʾ' "
        Set rs����ҽѧ��ʾ = zlDatabase.OpenSQLRecord(strSQL, "����ҽѧ��ʾ", lng����ID)
        While rs����ҽѧ��ʾ.EOF = False
            txtOtherWaring.Text = Nvl(rs����ҽѧ��ʾ!��Ϣֵ)
            rs����ҽѧ��ʾ.MoveNext
        Wend
        '��ϵ�������Ϣ
        'ȡ������Ϣ���е���ϵ����Ϣ
        strSQL = "" & _
        "   Select  ��ϵ������,��ϵ�˹�ϵ,��ϵ�˵绰,��ϵ�����֤�� From ������Ϣ Where ����ID=[1] And Not ��ϵ������ is Null"
        Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ��ϵ����Ϣ", lng����ID)
            If rs������Ϣ.EOF = False Then
            txt��ϵ�����֤��.Text = Nvl(rs������Ϣ!��ϵ�����֤��)
            txt��ϵ������.Text = Nvl(rs������Ϣ!��ϵ������)
            For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                If NeedName(cbo��ϵ�˹�ϵ.List(i)) = Nvl(rs������Ϣ!��ϵ�˹�ϵ) Then cbo��ϵ�˹�ϵ.ListIndex = i
            Next
            txt��ϵ�˵绰.Text = Nvl(rs������Ϣ!��ϵ�˵绰)
            
            SetLinkInfo Nvl(rs������Ϣ!��ϵ������), Nvl(rs������Ϣ!��ϵ�˹�ϵ), Nvl(rs������Ϣ!��ϵ�˵绰), Nvl(rs������Ϣ!��ϵ�����֤��)
        End If
        'ȡ������Ϣ�ӱ��е���ϵ����Ϣ
        strSQL = "" & _
        "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� like '��ϵ��%' order by ��Ϣ�� Asc"
        Set rs��ϵ�� = zlDatabase.OpenSQLRecord(strSQL, "��ϵ�������Ϣ", lng����ID)
        If rs��ϵ��.EOF = False Then
            rs��ϵ��.Filter = "��Ϣ�� like '��ϵ������%'"
            lng��ϵ������ = rs��ϵ��.RecordCount
            rs��ϵ��.Filter = ""
            For i = 2 To lng��ϵ������ + 1
                While rs��ϵ��.EOF = False
                    Select Case Nvl(rs��ϵ��!��Ϣ��)
                        Case "��ϵ������" & i
                            str��ϵ������ = Nvl(rs��ϵ��!��Ϣֵ)
                        Case "��ϵ�˹�ϵ" & i
                            str��ϵ�˹�ϵ = Nvl(rs��ϵ��!��Ϣֵ)
                        Case "��ϵ�˵绰" & i
                            str��ϵ�˵绰 = Nvl(rs��ϵ��!��Ϣֵ)
                        Case "��ϵ�����֤��" & i
                            str��ϵ�����֤�� = Nvl(rs��ϵ��!��Ϣֵ)
                    End Select
                    rs��ϵ��.MoveNext
                Wend
                SetLinkInfo str��ϵ������, str��ϵ�˹�ϵ, str��ϵ�˵绰, str��ϵ�����֤��
                rs��ϵ��.MoveFirst
            Next
        End If
        '������Ϣ
        strSQL = "" & _
        "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� Not in ('Ѫ��','ABO','RH','ҽѧ��ʾ','����ҽѧ��ʾ') And ��Ϣ�� Not like '��ϵ��%'"
        Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ϵ��������Ϣ", lng����ID)
        While rs������Ϣ.EOF = False
                SetOtherInfo rs������Ϣ!��Ϣ��, rs������Ϣ!��Ϣֵ
            rs������Ϣ.MoveNext
        Wend
        'ҽ�ƿ�����
        Set mdicҽ�ƿ����� = Nothing
    End If
    
    Exit Sub
ErrHandl:
     If ErrCenter() = 1 Then Resume
End Sub

Private Sub Add�����������Ϣ(ByVal lng����ID As Long, ByRef colPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ݴ���
    '���:
    '����:56599
    '����:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim varKey As Variant
    '����ҩ��
    With vsDrug
        If .Rows > 1 Then
            '����ò������м�¼
            strSQL = " Zl_���˹���ҩ��_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_���˹���ҩ��_Update("
                    '����ID_In ���˹���ҩ��.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ҩ��ID_In ���˹���ҩ��.����ҩ��ID%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 2) = "", "", .TextMatrix(i, 2)) & "',"
                    '����ҩ��_In  ���˹���ҩ��.����ҩ��%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '������Ӧ_In ���˹�����Ӧ.������Ӧ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '������Ϣ
    With vsInoculate
        If .Rows > 1 Then
            '����ò������м�¼
            strSQL = " Zl_�������߼�¼_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSQL
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '������Ϣ
    'ABOѪ��
    '������Ϣ�ӱ�
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'Ѫ��',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & NeedName(cboBloodType.Text, ".") & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'RH',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'ҽѧ��ʾ
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & txtMedicalWarning.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '����ҽѧ��ʾ
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'����ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & txtOtherWaring.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    
    '��ϵ�������Ϣ
    With vsLinkMan
        If .Rows > 1 And .TextMatrix(1, 0) <> "" And Not mrsInfo Is Nothing Then
            SaveModifyPati
        End If
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '��ϵ����������Ϊ��
                    For j = 0 To .Cols - 1
                        strSQL = "Zl_������Ϣ�ӱ�_Update("
                        '����ID_In ������Ϣ�ӱ�.����Id%Type
                        strSQL = strSQL & "" & lng����ID & ","
                        '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                        strSQL = strSQL & "'" & .TextMatrix(0, j) & i & "',"
                        '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                        strSQL = strSQL & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '����Id_In ������Ϣ�ӱ�.����Id%Type
                        strSQL = strSQL & "'')"
                        
                        zlAddArray colPro, strSQL
                    Next
                End If
            Next
        End If
    End With
    '������Ϣ
     With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    strSQL = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 0) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSQL = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 2) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
     End With
     'ҽ�ƿ�����
     If Not mdicҽ�ƿ����� Is Nothing Then
        For Each varKey In mdicҽ�ƿ�����.Keys
            strSQL = "Zl_����ҽ�ƿ�����_Update("
            strSQL = strSQL & lng����ID & ","
            strSQL = strSQL & mlngCardTypeID & ","
            strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdicҽ�ƿ�����(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
End Sub
Private Sub DeletePatPicture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��������Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-14 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo Errhand:
    strSQL = strSQL & "Zl_������Ƭ_Delete("
    strSQL = strSQL & lng����ID & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SavePatPicture(lng����ID As Long, strFile As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���没����Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
        gcnOracle.Execute "Delete From ������Ƭ Where ����id=" & lng����ID
        gstrSQL = "Select ����id,��Ƭ From ������Ƭ where ����id=" & lng����ID
        
        If strFile = "" Then Exit Sub
        rs.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
        If rs.BOF Then
    
            If rs.EOF Then rs.AddNew
    
            rs("����id").value = lng����ID
            rs("��Ƭ").value = Null
            rs.Update
            
            If zlDatabase.SavePicture(strFile, rs, "��Ƭ") = False Then
    
            ShowMsgbox "������Ƭ����,��ȷ���ļ��Ƿ�ɾ��!"
    
            Exit Sub
    
            End If
    
            rs.Close
        End If
End Sub
Private Function ReadPatPricture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-13 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTmp As String
    Dim rsData As Recordset
    
    '67776:������,2013-11-20,��ȡ����Ƭ�Ĳ�����Ϣ����Ƭû�����
    Set imgPatient.Picture = Nothing
    strSQL = "Select ����id,��Ƭ From ������Ƭ Where ����id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsData.BOF = False Then
        strTmp = zlDatabase.ReadPicture(rsData, "��Ƭ", strTmp)
        mstr�ɼ�ͼƬ = strTmp
        imgPatient.Picture = LoadPicture(strTmp)
    
       If strTmp <> "" Then Kill strTmp
    End If
End Function

Private Function Get�ƿ�XML(lng����ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ƿ�XML��(���ڴ����ƿ������ѽ����ƿ�����)
    '���:
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    
    strXML = strXML & "<����>" & Trim(txt����.Text) & "</����>"  'Varchar2 20
    strXML = strXML & "<����>" & Trim(txtPatient.Text) & "</����>"  'Varchar2 100
    strXML = strXML & "<�Ա�>" & NeedName(cbo�Ա�) & "</�Ա�>"  'Varchar2 4
    strXML = strXML & "<����>" & txt����.Text & cbo���䵥λ.Text & "</����>"  'Varchar2 10
    strXML = strXML & "<��������>" & Format(txt��������.Text & " " & txt����ʱ��.Text, "yyyy-mm-dd hh24:mi:ss") & "</��������>" 'Varchar2 20 yyyy-mm-dd hh24:mi:ss
    strXML = strXML & "<�����ص�>" & Trim(txt�����ص�.Text) & "</�����ص�>"  'Varchar2 50
    strXML = strXML & "<���֤��>" & Trim(txt���֤��.Text) & "</���֤��>"  'Varchar2 18
    strXML = strXML & "<����֤��>" & Trim(txt����֤��.Text) & "</����֤��>" 'Varchar2 20
    strXML = strXML & "<ְҵ>" & NeedName(cboְҵ.Text, mstrCboSplit) & "</ְҵ>" 'Varchar2 80
    strXML = strXML & "<����>" & NeedName(cbo����.Text) & "</����>" 'Varchar2 20
    strXML = strXML & "<����>" & NeedName(cbo����.Text) & "</����>" 'Varchar2 30
    strXML = strXML & "<ѧ��>" & NeedName(cboѧ��.Text) & "</ѧ��>" 'Varchar2 10
    strXML = strXML & "<����״��>" & NeedName(cbo����״��.Text) & "</����״��>" 'Varchar2 4
    strXML = strXML & "<����>" & NeedName(txt����.Text) & "</����>" 'Varchar2 30
    strXML = strXML & "<��ͥ��ַ>" & IIf(mblnStructAdress, Trim(padd��ͥ��ַ.value), Trim(txt��ͥ��ַ.Text)) & "</��ͥ��ַ>" 'Varchar2 50
    strXML = strXML & "<��ͥ�绰>" & Trim(txt��ͥ�绰.Text) & "</��ͥ�绰>" 'Varchar2 20
    strXML = strXML & "<�ֻ���>" & Trim(txt�ֻ�.Text) & "</�ֻ���>" 'Varchar2 20
    strXML = strXML & "<�����ʱ�>" & Trim(txt���ڵ�ַ�ʱ�.Text) & "</�����ʱ�>" 'Varchar2 6
    strXML = strXML & "<�໤��>" & "" & "</�໤��>" 'Varchar2 64
    strXML = strXML & "<��ϵ������>" & Trim(txt��ϵ������.Text) & "</��ϵ������>" 'Varchar2 64
    strXML = strXML & "<��ϵ�˹�ϵ>" & NeedName(cbo��ϵ�˹�ϵ.Text) & "</��ϵ�˹�ϵ>" 'Varchar2 30
    strXML = strXML & "<��ϵ�˵�ַ>" & Trim(txt��ϵ�˵�ַ.Text) & "</��ϵ�˵�ַ>" 'Varchar2 50
    strXML = strXML & "<��ϵ�˵绰>" & Trim(txt��ϵ�˵绰.Text) & "</��ϵ�˵绰>" 'Varchar2 20
    strXML = strXML & "<������λ>" & Trim(txt������λ.Text) & "</������λ>" 'Varchar2 100
    strXML = strXML & "<��λ�绰>" & Trim(txt��λ�绰.Text) & "</��λ�绰>" 'Varchar2 20
    strXML = strXML & "<��λ�ʱ�>" & Trim(txt��λ�ʱ�.Text) & "</��λ�ʱ�>" 'Varchar2 6
    strXML = strXML & "<��λ������>" & Trim(txt��λ������.Text) & "</��λ������>" 'Varchar2 50
    strXML = strXML & "<��λ�ʺ�>" & Trim(txt��λ�ʻ�.Text) & "</��λ�ʺ�>" 'Varchar2 20
    strXML = strXML & "<����ID>" & lng����ID & "</����ID>" 'Varchar2 18
    strXML = strXML & "<ABOѪ��>" & cboBloodType.Text & "</ABOѪ��>"  'Varchar2 10
    strXML = strXML & "<RH>" & cboBH.Text & "</RH>"  'Varchar2 10
    'ҽѧ��ʾ
    strXML = strXML & "<������־>" & Getҽѧ��ʾ("������־") & "</������־>" 'Varchar2 2
    strXML = strXML & "<���ಡ��־>" & Getҽѧ��ʾ("���ಡ��־") & "</���ಡ��־>" 'Varchar2 2
    strXML = strXML & "<����Ѫ�ܲ���־>" & Getҽѧ��ʾ("����Ѫ�ܲ���־") & "</����Ѫ�ܲ���־>" 'Varchar2 2
    strXML = strXML & "<��ﲡ��־>" & Getҽѧ��ʾ("��ﲡ��־") & "</��ﲡ��־>" 'Varchar2 2
    strXML = strXML & "<��Ѫ���ұ�־>" & Getҽѧ��ʾ("��Ѫ���ұ�־") & "</��Ѫ���ұ�־>" 'Varchar2 2
    strXML = strXML & "<���򲡱�־>" & Getҽѧ��ʾ("���򲡱�־") & "</���򲡱�־>" 'Varchar2 2
    strXML = strXML & "<����۱�־>" & Getҽѧ��ʾ("����۱�־") & "</����۱�־>" 'Varchar2 2
    strXML = strXML & "<͸����־>" & Getҽѧ��ʾ("͸����־") & "</͸����־>" 'Varchar2 2
    strXML = strXML & "<������ֲ��־>" & Getҽѧ��ʾ("������ֲ��־") & "</������ֲ��־>" 'Varchar2 2
    strXML = strXML & "<����ȱʧ��־>" & Getҽѧ��ʾ("����ȱʧ��־") & "</����ȱʧ��־>" 'Varchar2 2
    strXML = strXML & "<��װж��֫��־>" & Getҽѧ��ʾ("��װж��֫��־") & "</��װж��֫��־>" 'Varchar2 2
    strXML = strXML & "<����������־>" & Getҽѧ��ʾ("����������־") & "</����������־>" 'Varchar2 2
    '����ҽѧ��ʾ
    strXML = strXML & "<����ҽѧ��ʾ>" & Trim(txtOtherWaring.Text) & "</����ҽѧ��ʾ>" 'Varchar2 100
    '��ϵ�������Ϣ
    strXML = strXML & GetLinkXML
    '�����������
    strXML = strXML & "<�����������>" & GetOtherInfo("�����������") & "</�����������>" 'Varchar2 18
    '��ũ��֤��
    strXML = strXML & "<��ũ��֤��>" & GetOtherInfo("��ũ��֤��") & "</��ũ��֤��>" 'Varchar2 18
    '����֤��
    strXML = strXML & Get����֤��
    '������Ϣ
    strXML = strXML & Get������Ϣ
    '�������
    strXML = strXML & Get����ҩ��
    '���߼�¼
    strXML = strXML & Get���߼�¼
    'ҽ�ƿ�����
    strXML = strXML & Getҽ�ƿ�����
    
    Get�ƿ�XML = strXML
End Function
Private Function Getҽ�ƿ�����() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ�����XML
    '���:
    '����:56599
    '����:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varKey As Variant
    Dim strXML As String
    strXML = "<ҽ�ƿ�����>"
    For Each varKey In mdicҽ�ƿ�����
        strXML = strXML & "<��Ϣ��>" & varKey & "</��Ϣ��>"
        strXML = strXML & "<��Ϣֵ>" & mdicҽ�ƿ�����.Item(varKey) & "</��Ϣֵ>"
    Next
    strXML = strXML & "</ҽ�ƿ�����>"
End Function
Private Function Get���߼�¼() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���߼�¼XML
    '���:
    '����:56599
    '����:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim i As Long
    
    strXML = "<���߼�¼>"
    With vsInoculate
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                strXML = strXML & "<��������>" & .TextMatrix(i, 1) & "</��������>"
                strXML = strXML & "<����ʱ��>" & .TextMatrix(i, 0) & "</����ʱ��>"
            End If
            If .TextMatrix(i, 3) <> "" Then
                strXML = strXML & "<��������>" & .TextMatrix(i, 3) & "</��������>"
                strXML = strXML & "<����ʱ��>" & .TextMatrix(i, 2) & "</����ʱ��>"
            End If
        Next
    End With
    strXML = strXML & "</���߼�¼>"
End Function
Private Function Get����ҩ��() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ҩ��
    '���:
    '����:56599
    '����:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim i As Long
    
    strXML = "<�������>"
    With vsDrug
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                strXML = strXML & "<ҩ������>" & .TextMatrix(i, 0) & "</ҩ������>"
                strXML = strXML & "<ҩ�ﷴӦ>" & .TextMatrix(i, 1) & "</ҩ�ﷴӦ>"
            End If
        Next
    End With
    strXML = strXML & "</�������>"
    
    Get����ҩ�� = strXML
End Function
Private Function Get������Ϣ() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:
    '����:56599
    '����:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim strSQL As String
    Dim rs֤������ As Recordset
    Dim str֤������ As String
    Dim i As Long
    On Error GoTo Errhand
    strSQL = "Select ���� From ֤������"
    Set rs֤������ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs֤������.EOF Then Get������Ϣ = "": Exit Function
    While rs֤������.EOF = False
        str֤������ = str֤������ & ";" & Nvl(rs֤������!����)
        rs֤������.MoveNext
    Wend
    str֤������ = str֤������ & ";"
    strXML = "<������Ϣ>"
    With vsOtherInfo
        For i = 1 To .Rows - 1
            If InStr(str֤������, ";" & .TextMatrix(i, 0) & ";") < 0 Then
                strXML = strXML & "<��Ϣ��>" & .TextMatrix(i, 0) & "</��Ϣ��>"
                strXML = strXML & "<��Ϣֵ>" & .TextMatrix(i, 1) & "</��Ϣֵ>"
            End If
            If InStr(str֤������, ";" & .TextMatrix(i, 2) & ";") < 0 Then
                strXML = strXML & "<��Ϣ��>" & .TextMatrix(i, 2) & "</��Ϣ��>"
                strXML = strXML & "<��Ϣֵ>" & .TextMatrix(i, 3) & "</��Ϣֵ>"
            End If
        Next
    End With
    strXML = strXML & "</������Ϣ>"
    Get������Ϣ = strXML
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function Get����֤��() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����֤��
    '���:
    '����:56599
    '����:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim strSQL As String
    Dim rs֤������ As Recordset
    Dim str֤������ As String
    Dim i As Long
    On Error GoTo Errhand
    strSQL = "Select ���� From ֤������"
    Set rs֤������ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs֤������.EOF Then Get����֤�� = "": Exit Function
    While rs֤������.EOF = False
        str֤������ = str֤������ & ";" & Nvl(rs֤������!����)
        rs֤������.MoveNext
    Wend
    str֤������ = str֤������ & ";"
    strXML = "<����֤��>"
    With vsOtherInfo
        For i = 1 To .Rows - 1
            If InStr(str֤������, ";" & .TextMatrix(i, 0) & ";") > 0 Then
                strXML = strXML & "<��Ϣ��>" & .TextMatrix(i, 0) & "</��Ϣ��>"
                strXML = strXML & "<��Ϣֵ>" & .TextMatrix(i, 1) & "</��Ϣֵ>"
            End If
            If InStr(str֤������, ";" & .TextMatrix(i, 2) & ";") > 0 Then
                strXML = strXML & "<��Ϣ��>" & .TextMatrix(i, 2) & "</��Ϣ��>"
                strXML = strXML & "<��Ϣֵ>" & .TextMatrix(i, 3) & "</��Ϣֵ>"
            End If
        Next
    End With
    strXML = strXML & "</����֤��>"
    Get����֤�� = strXML
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Function Getҽѧ��ʾ(str��־ As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽѧ��ʾ
    '���:str��־ - ��Ҫ���ҵı�־
    '����:56599
    '����:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Getҽѧ��ʾ = IIf(InStr(";" & txtMedicalWarning.Text & ";", Replace(str��־, "��־", "")) > 0, 1, 0)
End Function
Private Function GetLinkXML() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ϵ����ϢXML�ַ���
    '���:
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim i As Long

    strXML = "<��ϵ��Ϣ>"
    With vsLinkMan
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then '��ϵ������������Ϊ��
                strXML = strXML & "<����>" & .TextMatrix(i, 0) & "</����>"
                strXML = strXML & "<��ϵ>" & .TextMatrix(i, 1) & "</��ϵ>"
                strXML = strXML & "<�绰>" & .TextMatrix(i, 3) & "</�绰>"
                strXML = strXML & "<���֤��>" & .TextMatrix(i, 2) & "</���֤��>"
            End If
        Next
    End With
    strXML = strXML & "</��ϵ��Ϣ>"
    GetLinkXML = strXML
End Function
Private Function GetOtherInfo(str��Ϣ�� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ڵ��ȡ������Ϣ��ָ��������
    '���:
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strFind As String
    With vsOtherInfo
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = str��Ϣ�� Then
                strFind = .TextMatrix(i, 1)
            End If
            If .TextMatrix(i, 2) = str��Ϣ�� Then
                strFind = .TextMatrix(i, 3)
            End If
        Next
    End With
    GetOtherInfo = strFind
End Function

Private Function WriteCard(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:д��
    '���:lng����ID - ����ID
    '����:����
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Function
    End If
    If mobjReadCard Is Nothing Then Exit Function
    '84196:���ϴ�,2015/4/22���ӿڲ�֧��д������Ϣ��ʾ
    On Error Resume Next
    WriteCard = mobjReadCard.zlBandCardArfter(Me, mlngModule, mlngCardTypeID, lng����ID, strExpend)
    'vbʵʱ����438����֧�ָ����Ի򷽷�
    If Err = 438 Then
        MsgBox mCardType.str������ & "�ӿڲ�֧��д��", vbInformation, gstrSysName
        Err.Clear
    End If
    If Err = 0 Then Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function Check��������(lng����ID As Long, lng�����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ����Ƿ����Ʋ��˵ķ�������
    '���:lng����ID - ����ID;lng�����ID  - ҽ�ƿ������ID
    '����:����
    '����:57326
    '����:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
        strSQL = "Select Count(1) as ���� From ����ҽ�ƿ���Ϣ Where ״̬=0 And ����ID=[1] And �����ID=[2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng�����ID)
        If Val(Nvl(rsTemp!����)) <= 0 Then Check�������� = True: Exit Function
        Select Case mCardType.lng��������
            Case 0 '������
                Check�������� = True
            Case 1 'ͬһ������ֻ����һ�ſ�
                MsgBox "�ò����Ѿ�����" & mCardType.str������ & ",�����ڽ��з�������!", vbInformation + vbOKOnly
                Check�������� = False
            Case 2 'ͬһ�������������ſ�,����Ҫ����
               Check�������� = MsgBox("�ò����Ѿ�����" & mCardType.str������ & ",�Ƿ�Ҫ���з�������?", vbQuestion + vbYesNo) = vbYes
        End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

'72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
Private Sub IDKindPay_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String
    
    If IsCardType(IDKindPay, "IC����") Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt����.Text = mobjICCard.Read_Card()
            If txt����.Text <> "" Then
                mblnICCard = True
                Call CheckFreeCard(txt����.Text)
                Call txt����_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    
    lng�����ID = IDKindPay.GetCurCard.�ӿ����
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
    'Call InitInterFacel(Me, mlngModule, lng�����ID, False, mobjCardObject)
    strExpand = lng�����ID
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txt����.Text = strOutCardNo
    If txt����.Text <> "" Then
        Call CheckFreeCard(txt����.Text)
        Call txt����_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub IDKindPay_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If wndTaskPanel.Groups.count = 0 Or IDKindPay.Enabled = False Then Exit Sub
    wndTaskPanel.Groups.Item(Idx_TP_PatiCard).Caption = objCard.����
    mlngCardTypeID = objCard.�ӿ����
    '���³�ʼ�������Ͷ�Ӧ����
    Call InitCardType: Call LoadCardFee
    txt����.MaxLength = mCardType.lng���ų���
    txt����.Locked = Not objCard.�Ƿ�ˢ��
    Call SetCardPayOrBound

     '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    Set gobjSquare.objCurCard = objCard
    
    mlngҽ�ƿ����� = objCard.���ų���
    txt����.PasswordChar = IIf(IDKindPay.ShowPassText, "*", "")
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txt����.Text <> "" Then txt����.Text = ""
    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txt����.IMEMode = 0
End Sub

Private Sub IDKindPay_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    
    If IDKindPay.Enabled = False Then Exit Sub
    If txt����.Visible = False Then Exit Sub
    'ֻ�ܶ�ѡ�����Ŀ�
    If mCardType.lng�����ID <> objCard.�ӿ���� Then Exit Sub
    
    txt����.Text = objPatiInfor.����
'    Call CheckAvailableCard(objPatiInfor)
    If txt����.Text <> "" Then
        Call ReLoadCardFee(True)
        Call CheckFreeCard(txt����.Text)
        Call LoadPati(objPatiInfor.PatiXML)
        Call zlQueryEMPIPatiInfo
    End If
    '76505,Ƚ����,2014-8-11,����ʱ����,���潹�㶪ʧ
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

'72541,Ƚ����,2014-5-7,�շѴ��ķ���ֻ�ܷ��ž��￨������
Private Sub tbPageDo_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If chkCancel.Visible = True Then chkCancel.value = 0
    Select Case Item.Index
    Case 0
        mEditType = Cr_����
    Case 1
        mEditType = Cr_�󶨿�
    End Select
    
    txt����.Text = ""
    txtPass.Text = ""
    txtAudi.Text = ""
    
    Call SetCardView
End Sub

Private Sub SetCardPayOrBound()
    '-------------------------------------------------------------------------------------
    '���ܣ��ڷ�����󶨿�֮���л�ʱ�����õ�ǰ��������
    '���ƣ�Ƚ����
    '���ڣ�2014-5-7
    '����ţ�72541
    '˵����
    '-------------------------------------------------------------------------------------
    Dim i As Integer
    Dim blnPay As Boolean, blnBound As Boolean
    Dim objItem As TabControlItem
    
    If mblnFromCardMgr Then mblnAddPage = False: tbPageDo.RemoveAll: Exit Sub '��������ģ�����ֱ����Ĭ�ϲ���
    '�Ƿ�ɷ���
    blnPay = zlCheckPrivs(mstrPrivs, "����") And (mCardType.bln���ƿ� Or (Not mCardType.bln���ƿ� And mCardType.bln�Ƿ񷢿�))
    '�Ƿ�ɰ󶨿�
    blnBound = zlCheckPrivs(mstrPrivs, "�󶨿�") And (Not mCardType.bln���ƿ� Or (mCardType.bln���ƿ� And mCardType.bln�Ƿ��ظ�ʹ��))
    
    If blnPay And blnBound Then '��ǰ�����ɷ�����Ҳ�ɰ󶨿�
        If tbPageDo.ItemCount <> 0 Then tbPageDo.RemoveAll
        Set objItem = tbPageDo.InsertItem(0, "����", fraCard.hWnd, 0): objItem.Tag = Cr_����
        Set objItem = tbPageDo.InsertItem(1, "�󶨿�", fraCard.hWnd, 0): objItem.Tag = Cr_�󶨿�
        If mEditType = Cr_�󶨿� Then
            tbPageDo(1).Selected = True
        Else
            tbPageDo(1).Selected = True: tbPageDo(0).Selected = True
        End If
        mblnAddPage = True
    ElseIf blnPay And Not blnBound Then '��ǰ�������ɷ���
        mEditType = Cr_����
        mblnAddPage = False: tbPageDo.RemoveAll
    ElseIf Not blnPay And blnBound Then
        mEditType = Cr_�󶨿�
        mblnAddPage = False: tbPageDo.RemoveAll
    End If
    Call SetCardView
End Sub

Private Sub SetCardView()
    '-------------------------------------------------------------------------------------
    '���ܣ��ڷ�����󶨿�֮���л�ʱ������������ʾ
    '���ƣ�Ƚ����
    '���ڣ�2014-5-7
    '����ţ�72541
    '˵����
    '-------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    Dim sngTaskHeight As Single, sngWinHeight As Single
    
    '������ʾ��Ϣ
    cmdCreateCard.Visible = (mEditType = Cr_���� Or mEditType = Cr_�󶨿�) And InStr(1, mstrPrivs, ";�ƿ�;") > 0 And mCardType.bln�Ƿ��ƿ�

    blnVisible = mEditType = Cr_���� Or mEditType = Cr_���� Or mEditType = Cr_�˿� Or chkCancel.value = 1
    lbl����.Visible = blnVisible: txt����.Visible = blnVisible
    chk����.Visible = blnVisible
    lbl֧����ʽ.Visible = blnVisible: cbo֧����ʽ.Visible = blnVisible
    '���²��ֵ�ǰ����ؼ�
    Call SetCtrlMove
End Sub

Private Function FromKindLoadPati(ByVal objPatiInfor As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����IDKind���ز�����Ϣ,��ȡ������Ϣ
    '����:Ƚ����
    '����:2014-05-08
    '����ţ�72541,75848
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '�����:56599
    Dim str����ҩ�� As String, str������Ӧ As String '�����:56599
    Dim str�������� As String, str�������� As String '�����:56599
    Dim strABOѪ�� As String '�����:56599
    Dim str��Ϣ�� As String, str��Ϣֵ As String '�����:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '�����:56599
    Dim str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, str��ַ As String '�����:56599
    
    On Error GoTo errHandle
    If Not (mEditType = Cr_�󶨿� Or mEditType = Cr_����) Then Exit Function
    If objPatiInfor Is Nothing Then Exit Function
    
    With objPatiInfor
        If .���� = "" Then Exit Function '�������Ϊ�գ�����Ϊû����ȡ������
        Call ClearData
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ����    Varchar2    20
        txt����.Text = .����
        '    ����    Varchar2    64
        txtPatient.Text = .����
        '    �Ա�    Varchar2    4
        If .�Ա� <> "" Then
            Call zlControl.CboLocate(cbo�Ա�, .�Ա�)
            If cbo�Ա�.ListIndex = -1 Then
                cbo�Ա�.AddItem .�Ա�
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
        End If
        '    ����    Varchar2    10
        If .���� <> "" Then
            Call LoadOldData(.����, txt����, cbo���䵥λ)
        End If
        '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
        txt��������.Text = Format(IIf(.�������� = "", "____-__-__", .��������), "YYYY-MM-DD")
        If .�������� <> "" Then
             txt����.Text = ReCalcOld(CDate(txt��������.Text), cbo���䵥λ)      '�޸ĵ�ʱ��,���ݳ���������������
             If CDate(txt��������.Text) - CDate(.��������) <> 0 Then txt����ʱ��.Text = Format(.��������, "HH:MM")
         Else
             txt����ʱ��.Text = "__:__"
             txt��������.Text = ReCalcBirth(Val(txt����.Text), cbo���䵥λ.Text)
         End If
        '    �����ص�    Varchar2    50
        txt�����ص�.Text = .������ַ
        '    ���֤��    VARCHAR2    18
        txt���֤��.Text = .���֤��
        '    ����֤��    Varchar2    20
        txt����֤��.Text = .����֤��
        '    ְҵ    Varchar2    80
        cboְҵ.ListIndex = GetCboIndex(cboְҵ, .ְҵ, , , mstrCboSplit)
        If cboְҵ.ListIndex = -1 Then
            cboְҵ.AddItem .ְҵ, 0
            cboְҵ.ListIndex = cboְҵ.NewIndex
        End If
        '    ����    Varchar2    20
        cbo����.ListIndex = GetCboIndex(cbo����, .����)
        If cbo����.ListIndex = -1 And .���� <> "" Then
            cbo����.AddItem .����, 0
            cbo����.ListIndex = cbo����.NewIndex
        End If
        '    ����    Varchar2    30
        cbo����.ListIndex = GetCboIndex(cbo����, .����)
        If cbo����.ListIndex = -1 And .���� <> "" Then
            cbo����.AddItem .����, 0
            cbo����.ListIndex = cbo����.NewIndex
        End If
        '    ѧ��    Varchar2    10
        cboѧ��.ListIndex = GetCboIndex(cboѧ��, .ѧ��)
        If cboѧ��.ListIndex = -1 And .ѧ�� <> "" Then
            cboѧ��.AddItem .ѧ��, 0
            cboѧ��.ListIndex = cboѧ��.NewIndex
        End If
        '    ����״��    Varchar2    4
        cbo����״��.ListIndex = GetCboIndex(cbo����״��, .����״��)
        If cbo����״��.ListIndex = -1 And .����״�� <> "" Then
            cbo����״��.AddItem .����״��, 0
            cbo����״��.ListIndex = cbo����״��.NewIndex
        End If
        '    ����    Varchar2    30
        txt����.Text = .����
        '    ��ͥ��ַ    Varchar2    50
        txt��ͥ��ַ.Text = .��ͥ��ַ
        Call zlReadAddrInfo(padd��ͥ��ַ, .����ID, 0, 3, .��ͥ��ַ)
        '    ��ͥ�绰    Varchar2    20
        txt��ͥ�绰.Text = .��ͥ�绰
        '    ��ͥ��ַ�ʱ�    Varchar2    6
        txt��ͥ�ʱ�.Text = .��ͥ�ʱ�
        '    �໤��  Varchar2    64
'        txt�໤��.Text = .�໤��
        '    ��ϵ������  Varchar2    64
        txt��ϵ������.Text = .��ϵ��
        '    ��ϵ�˹�ϵ  Varchar2    30
        cbo��ϵ�˹�ϵ.ListIndex = GetCboIndex(cbo��ϵ�˹�ϵ, .��ϵ�˹�ϵ)
        If cbo��ϵ�˹�ϵ.ListIndex = -1 And .��ϵ�˹�ϵ <> "" Then
            cbo��ϵ�˹�ϵ.AddItem .��ϵ�˹�ϵ, 0
            cbo��ϵ�˹�ϵ.ListIndex = cbo��ϵ�˹�ϵ.NewIndex
        End If
        '    ��ϵ�˵�ַ  Varchar2    50
        txt��ϵ�˵�ַ.Text = .��ϵ�˵�ַ
        '    ��ϵ�˵绰  Varchar2    20
        txt��ϵ�˵绰.Text = .��ϵ�˵绰
        '    ������λ    Varchar2    100
        txt������λ.Text = .������λ
        lbl������λ.Tag = ""
        '    ��λ�绰    Varchar2    20
        txt��λ�绰.Text = .������λ��ַ
        '    ��λ�ʱ�    Varchar2    6
        txt��λ�ʱ�.Text = .������λ�ʱ�
        '    ��λ������  Varchar2    50
        txt��λ������.Text = .������λ������
        '    ��λ�ʺ�    Varchar2    20
        txt��λ�ʻ�.Text = .������λ�������ʻ�
        '    �ֻ���      Varchar2    20
        txt�ֻ�.Text = .�ֻ���
    End With
    FromKindLoadPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetValidKindStr(ByVal lngCardTypeID As Long) As Boolean
    '----------------------------------------------
    '���ܣ�������Чҽ�ƿ�IDKind�����ַ���,���жϴ���ҽ�ƿ�����Ƿ�����Чҽ�ƿ���
    '�������ж��Ƿ���ڵ�ҽ�ƿ����ID
    '���أ�
    '   1:����ҽ�ƿ�������
    '   2:����ҽ�ƿ������IDKind�ؼ����������á���δ����
    '   0:����ҽ�ƿ�����ڲ���������δ����
    '���ƣ�Ƚ����
    'ʱ�䣺2014-5-16
    '���⣺72541
    '˵����
    '
    '----------------------------------------------
    Dim objCard As Card, i As Long, blnDelete As Boolean
    Dim objCards As Cards
    
    On Error GoTo errHandle
    If Not IDKindPay.Cards Is Nothing Then
        Set objCards = IDKindPay.Cards
        For Each objCard In objCards
            blnDelete = False
            With objCard
                If Not zlCheckPrivs(mstrPrivs, "����") And .���ƿ� = 1 Then blnDelete = True '�޷���Ȩ�޲��ܷ���
                If Not zlCheckPrivs(mstrPrivs, "�󶨿�") And .���ƿ� = 0 Then blnDelete = True '�ް󶨿�Ȩ�޲��ܰ󶨿�
                If mblnFromCardMgr And .�ӿ���� <> lngCardTypeID Then blnDelete = True '�����������ֻ�ܶԵ�ǰ�����в���
                If .�ӿ���� = 0 Then blnDelete = True 'ɾ��Ĭ�Ϸ������
                '�Ƴ�
                If blnDelete Then
                    If .�ӿ���� = 0 Then
                        objCards.Remove "M" & .����
                    Else
                        objCards.Remove "K" & .�ӿ����
                    End If
                Else
                    If .�ӿ���� = lngCardTypeID Then GetValidKindStr = True
                End If
            End With
        Next

    End If
    Set IDKindPay.Cards = objCards
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������֤ͼ��
    '����:���˺�
    '����:2014-06-30 16:20:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    mlngͼ����� = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CreateObjectPlugIn() As Boolean
    '����:��������������Ϣ���
    '����:�����ɹ�,����True,���򷵻�False
    '�����:73935
    '����:Ƚ����
    '����:2014-07-3
    mblnPlugin = False
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        Call mobjPlugIn.Initialize(gcnOracle, glngSys, mlngModule)
        mblnPlugin = Err = 0
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
    End If
    CreateObjectPlugIn = True
End Function

Private Function InitTaskPanelOther() As Boolean
    '����:���ظ�����Ϣҳ��
    '����:
    '�����:73935
    '����:Ƚ����
    '����:2014-07-3
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    
    Err = 0: On Error GoTo Errhand
    If Not mobjPlugIn Is Nothing Then
        If mlngPlugInHwnd <> 0 Then
            With wndTaskPanelOther
                Call .SetGroupInnerMargins(0, 0, 0, 0)
                Call .SetGroupOuterMargins(-1, -24, -1, -1)
                
                Set tkpGroup = .Groups.Add(1, "������Ϣ")
                tkpGroup.CaptionVisible = False
                tkpGroup.Expandable = False
                tkpGroup.Expanded = True
                
                Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
                Call HideFormCaption(mlngPlugInHwnd, False) '���ش���߿�
                Item.Handle = mlngPlugInHwnd
                
                .HotTrackStyle = xtpTaskPanelHighlightItem
                .Reposition
                .DrawFocusRect = True
            End With
        End If
    End If
    InitTaskPanelOther = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Sub HideFormCaption(ByVal lngHwnd As Long, Optional ByVal blnBorder As Boolean = True)
'���ܣ�����һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(lngHwnd, vRect)
    lngStyle = GetWindowLong(lngHwnd, GWL_STYLE)

    If blnBorder Then
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
    Else
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    End If
    SetWindowLong lngHwnd, GWL_STYLE, lngStyle
    SetWindowPos lngHwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����zlPublicPatient����
    '����:�����ɹ�,����True,���򷵻�False
    '����:Ƚ����
    '����:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "������Ϣ����������zlPublicPatient������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "������Ϣ����������zlPublicPatient����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CreatePublicPatient = True
End Function

Private Sub AddDepositSQL(ByVal lng����ID As Long, ByVal dtCurdate As Date, _
    ByRef cllPro As Collection, ByRef lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�����SQL
    '����:���˺�
    '����:2011-07-26 18:26:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, strSQL As String, i As Integer
    Dim dblMoney As Double, str���㷽ʽ As String
     
    '����Ԥ�����¼
    str���㷽ʽ = mcolPayMode(cbo֧����ʽ.ListIndex + 1)(6)
    If str���㷽ʽ = "" Then str���㷽ʽ = NeedName(cbo֧����ʽ.Text)
    If Not cbo֧����ʽ.Enabled Then str���㷽ʽ = ""
        
    mstrPrePayNo = zlDatabase.GetNextNo(11)
    mlngԤ��ID = zlDatabase.GetNextId("����Ԥ����¼")
    mlngԤ������ID = lng����ID
    mdatԤ��ʱ�� = dtCurdate
    dblMoney = StrToNum(txt���.Text)
    'Zl_����Ԥ����¼_Insert
    strSQL = "Zl_����Ԥ����¼_Insert("
    '  Id_In         ����Ԥ����¼.ID%Type,
    strSQL = strSQL & "" & mlngԤ��ID & ","
    '  ���ݺ�_In     ����Ԥ����¼.NO%Type,
    strSQL = strSQL & "'" & mstrPrePayNo & "',"
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "" & IIf(mblnPrepayPrint, "'" & mstrPrepayInvioce & "'", "Null") & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
    strSQL = strSQL & "NULL,"
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "NULL,"
    '  ���_In       ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  �������_In   ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  �ɿλ_In   ����Ԥ����¼.�ɿλ%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  ��λ������_In ����Ԥ����¼.��λ������%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  ��λ�ʺ�_In   ����Ԥ����¼.��λ�ʺ�%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    strSQL = strSQL & "'ҽ�ƿ�:" & mCurPayMoney.strNo & "',"
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(mlng����ID = 0, "NULL", mlng����ID) & ","
    '  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type := Null,
    strSQL = strSQL & "" & 1 & ","
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPayMoney.lngҽ�ƿ����ID = 0 Or mCurPayMoney.bln���ѿ�, "NULL", mCurPayMoney.lngҽ�ƿ����ID) & ","
   '  ���㿨���_in ����Ԥ����¼.���㿨���%type:=NULL,
    strSQL = strSQL & "" & IIf(mCurPayMoney.lngҽ�ƿ����ID = 0 Or Not mCurPayMoney.bln���ѿ�, "NULL", mCurPayMoney.lngҽ�ƿ����ID) & ","
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPayMoney.strˢ������ = "", "NULL", "'" & mCurPayMoney.strˢ������ & "'") & ","
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  ������λ_In   ����Ԥ����¼.������λ%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null
    strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '   ��������_In Integer:=0 :0-������Ԥ��;1-��Ϊ���۵�
    strSQL = strSQL & "0 )"
    zlAddArray cllPro, strSQL
End Sub

Private Function CheckDepositFactValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤ����Ʊ��
    '����:������ȡ,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-30 11:14:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng����ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    
    On Error GoTo errHandle
    mlng����ID = 0
    
    mstrPrepayInvioce = "": mblnPrepayPrint = False
    
    '��鱾�γ�ֵ���
    '114422:���ϴ���2017/11/14,���жϽ��㷽ʽ
    If IDKindPayMode.IDKind = 2 And cbo֧����ʽ.Enabled And cbo֧����ʽ.Visible Then
        If cbo֧����ʽ.ListIndex = -1 Then
            MsgBox "��ȷ��" & mCardType.str������ & "�Ľɿ���㷽ʽ��", vbExclamation, gstrSysName
            If cbo֧����ʽ.Enabled And cbo֧����ʽ.Visible Then cbo֧����ʽ.SetFocus: Exit Function
        ElseIf cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> 1 And Val(txt���.Text) < 0 Then
            MsgBox "��Ԥ������Ϊ���������ٴ�ȷ�Ͻɿ��", vbExclamation, gstrSysName
            If txt�ϼ�.Enabled And txt�ϼ�.Visible Then txt�ϼ�.SetFocus: Exit Function
        End If
    End If
    
    '�����ڳ�Ԥ��
    If Not (Val(txt���.Text) > 0 And IDKindPayMode.IDKind = 2) Then CheckDepositFactValied = True: Exit Function

    mFactProperty = zl_GetInvoicePreperty(mlngModule, 2, 1)
    
    Select Case mFactProperty.intInvoicePrint
    Case 0 '����ӡ
        CheckDepositFactValied = True: Exit Function
    Case 1 '�Զ���ӡ
        mblnPrepayPrint = True
    Case 2 'ѡ���Ƿ��ӡ
        If MsgBox("�Ƿ��ӡԤ��Ʊ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) <> vbYes Then CheckDepositFactValied = True: Exit Function
        mblnPrepayPrint = True
    End Select
    
    If mblnBillԤ�� = False Then
        '�п����ǵ�һ��ʹ��
        Do
            blnInput = False
            '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
            strInvoice = UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngModule, ""))
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("û���ҵ����õ�Ԥ��Ʊ�ݵ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                vbCrLf & "�����뽫Ҫʹ�õ�Ԥ��Ʊ�ݵĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", Me.Left + 3000, Me.Top + 3000))
                blnInput = True
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                strInvoice = UCase(InputBox("��ȷ��ʹ�õ�Ԥ��Ʊ�ݵĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                strInvoice, Me.Left + 3000, Me.Top + 3000))
                blnInput = True
            End If
                
            '�û�ȡ������,�����ӡ
            If strInvoice = "" Then
                If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnValid = True
            Else
                '���������Ч��
                If blnInput Then
                    If zlCommFun.ActualLen(strInvoice) <> mbytԤ�� Then
                        MsgBox "����Ԥ����Ʊ�ݺ��볤��Ӧ��Ϊ " & mbytԤ�� & " λ��", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
        mstrPrepayInvioce = strInvoice
        CheckDepositFactValied = True: Exit Function
    End If
    
    Do
        '����Ʊ�����ö�ȡ
        blnInput = False
        mlng����ID = CheckUsedBill(2, IIf(mlng����ID > 0, mlng����ID, mFactProperty.lngShareUseID), "", 1)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õ�����Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Exit Function
                Case -2
                    MsgBox "���صĹ���Ԥ��Ʊ�ݵ�����Ԥ��Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Exit Function
                    strInvoice = ""
            End Select
        End If
        strInvoice = GetNextBill(mlng����ID)

        If strInvoice = "" Then
            '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
            strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ��Ԥ��Ʊ�ݵĿ�ʼƱ�ݺţ�" & _
                            vbCrLf & "�������뽫Ҫʹ�õ�Ʊ�ݺ��룺", gstrSysName, _
                            "", Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        Else
            strInvoice = UCase(InputBox("��ȷ��ʹ��ʹ��Ԥ��Ʊ�ݵ�Ʊ�ݺ��룺", gstrSysName, _
                            strInvoice, Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        End If
        
        '�û�ȡ������,����ӡ
        If strInvoice = "" Then Exit Function
        
        '���������Ч��
        If blnInput Then
            mlng����ID = GetInvoiceGroupID(2, 1, mlng����ID, mFactProperty.lngShareUseID, strInvoice, 1)
            If lng����ID < 0 Then
                MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
            Else
                blnValid = True
            End If
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    mstrPrepayInvioce = strInvoice
    CheckDepositFactValied = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub setFact()
    If Not mblnBillԤ�� And mstrPrepayInvioce <> "" Then
        zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", mstrPrepayInvioce, glngSys, mlngModule
    End If
End Sub

Private Sub SetCmdCtrlMove()
    '78408:���ϴ�,2014/10/9,����ҩ��ѡ��ʽ
    With vsDrug
        If .Row >= 1 And .Col = 0 And .Visible = True And .Enabled = True Then
            cmdSelDrug.Left = .CellLeft + .CellWidth - cmdSelDrug.Width
            cmdSelDrug.Top = .CellTop + 15
            cmdSelDrug.Visible = True
        Else
            cmdSelDrug.Visible = False
        End If
    End With
End Sub

Private Function CheckMobile(strMobile As String) As Boolean
    '����Ƿ����������˵��ֻ����ظ�
    Dim strSQL As String, lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State = 0 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    strSQL = "Select 1 From ������Ϣ Where �ֻ���=[1] And ����ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˷�������", strMobile, lng����ID)
    If rsTmp.RecordCount > 0 Then
        If MsgBox("������ֻ��������������ظ����Ƿ�ȷ��¼�룿", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Function
    End If
    CheckMobile = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlInitMEPIPati(ByRef rsPati As ADODB.Recordset) As Boolean
    Set rsPati = New ADODB.Recordset
    With rsPati
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "����ID", adBigInt, , adFldIsNullable
            .Append "��ҳID", adBigInt, , adFldIsNullable
            .Append "�Һ�ID", adBigInt, , adFldIsNullable
            .Append "�����", adVarChar, 18, adFldIsNullable
            .Append "סԺ��", adVarChar, 18, adFldIsNullable
            .Append "ҽ����", adVarChar, 30, adFldIsNullable
            .Append "���֤��", adVarChar, 18, adFldIsNullable
            .Append "����֤��", adVarChar, 20, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "�Ա�", adVarChar, 4, adFldIsNullable
            .Append "��������", adVarChar, 20, adFldIsNullable
            .Append "�����ص�", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "ѧ��", adVarChar, 10, adFldIsNullable
            .Append "ְҵ", adVarChar, 80, adFldIsNullable
            .Append "������λ", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����״��", adVarChar, 4, adFldIsNullable
            .Append "��ͥ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ϵ�˵绰", adVarChar, 20, adFldIsNullable
            .Append "��λ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ͥ��ַ", adVarChar, 100, adFldIsNullable
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "���ڵ�ַ", adVarChar, 100, adFldIsNullable
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��λ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��ϵ�˵�ַ", adVarChar, 100, adFldIsNullable
            .Append "��ϵ�˹�ϵ", adVarChar, 30, adFldIsNullable
            .Append "��ϵ������", adVarChar, 64, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    zlInitMEPIPati = True
End Function

Public Sub zlQueryEMPIPatiInfo(Optional ByVal strPatiName As String)
    '���ܣ���EMPIƽ̨��ȡ������Ϣ
    '���ڣ�2016/10/9 10:47:13
    '���ƣ����ϴ�
    '˵����101170
    Dim rsTmp As ADODB.Recordset, lng����ID As Long, strDiff As String, strMsgInfo As String
    Dim strSQL As String
    If mblnPlugin = False Then Exit Sub
    If mobjPlugIn Is Nothing Then Exit Sub
    If mEditType <> Cr_���� And mEditType <> Cr_�󶨿� And mEditType <> Cr_����������Ϣ Or chkCancel.value = 1 Then Exit Sub
    
    On Error GoTo Errhand
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State = 0 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    '���������ڷ���ʱ�������������Ϣ
    If lng����ID <> 0 And mEditType <> Cr_����������Ϣ Then Exit Sub
    If zlInitMEPIPati(rsTmp) = False Then Exit Sub
    
    With rsTmp
        .AddNew
        !����ID = lng����ID
        !����� = txt�����.Text
        !ҽ���� = txtҽ����.Text
        !���֤�� = txt���֤��.Text
        !���� = IIf(strPatiName = "", txtPatient.Text, strPatiName)
        !�Ա� = NeedName(cbo�Ա�.Text)
        If IsDate(txt��������) Then
            !�������� = Format(txt�������� & " " & IIf(IsDate(txt����ʱ��), txt����ʱ��, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !�������� = ""
        End If
        !�����ص� = txt�����ص�.Text
        !���� = NeedName(cbo����.Text)
        !���� = NeedName(cbo����.Text)
        !ְҵ = NeedName(cboְҵ.Text)
        !������λ = txt������λ.Text
        !��ͥ�绰 = txt��ͥ�绰.Text
        !��ϵ�˵绰 = txt��ϵ�˵绰.Text
        !��λ�绰 = txt��λ�绰.Text
        !��ͥ��ַ = txt��ͥ��ַ.Text
        !��ͥ��ַ�ʱ� = txt��ͥ�ʱ�.Text
        !���ڵ�ַ = txt���ڵ�ַ.Text
        !���ڵ�ַ�ʱ� = txt���ڵ�ַ�ʱ�.Text
        !��λ�ʱ� = txt��λ�ʱ�.Text
        !��ϵ������ = txt��ϵ������.Text
        !��ϵ�˹�ϵ = NeedName(cbo��ϵ�˹�ϵ.Text)
        .Update
    End With
    'EMPIû���ҵ�������Ϣ,ֱ�ӷ���
    Dim rsOut As New ADODB.Recordset
    On Error Resume Next
    If mobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsTmp, rsOut) = False Then
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: Set mrsEMPIOut = Nothing: Exit Sub
    End If
    Err.Clear: On Error GoTo Errhand
    Set mrsEMPIOut = rsOut
    If mrsEMPIOut Is Nothing Then Exit Sub
    If mrsEMPIOut.RecordCount = 0 Then Exit Sub
    mrsEMPIOut.MoveFirst
    On Error Resume Next
    With mrsEMPIOut
        '104905:���ϴ���2017/1/12������EMPI���صĲ���ID�����Ҳ���
        If lng����ID <> Val(Nvl(!����ID)) And Val(Nvl(!����ID)) <> 0 Then
            strSQL = "" & _
            "   Select Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����," & _
            "        A.����id,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.�ѱ�,A.ҽ�Ƹ��ʽ," & _
            "        A.����,A.�Ա�,A.����,A.��������,A.�����ص�,A.���֤��," & _
            "        A.����֤��,A.���,A.ְҵ,A.����,A.����,A.����,A.ѧ��,A.����״��,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�໤��,A.��ϵ������," & _
            "        A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.��ͬ��λid,A.������λ,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.������," & _
            "        A.������,A.��������,A.����ʱ��,A.����״̬,A.��������,A.סԺ����,A.��ǰ����id,A.��ǰ����id,A.��ǰ����,A.��Ժʱ��,A.��Ժʱ��," & _
            "        A.��Ժ,A.Ic����,A.������,A.ҽ����,A.����,A.��ѯ����,A.�Ǽ�ʱ��,A.ͣ��ʱ��,A.����,A.��ϵ�����֤��,A.���ڵ�ַ,A.���ڵ�ַ�ʱ�," & _
            "        M.���� as ���ʽ����, decode(B1.��������,NULL,0,1,1,0) as ����,B1.��ע, " & _
            "        Nvl(Nvl(A.��������,B1.��������),Decode(Nvl(A.����,B1.����),Null,'��ͨ����','ҽ������')) ��������,B1.��Ժ����, C.���� ��������," & _
            "        A.�ֻ���" & _
            "   From ������Ϣ A,������ҳ B1,������� C ,ҽ�Ƹ��ʽ M" & _
            "   Where A.���� = C.���(+) And A.ҽ�Ƹ��ʽ=M.����(+) " & _
            "               And A.����ID=B1.����ID(+) And A.��ҳID=B1.��ҳID(+) " & _
            "               And A.ͣ��ʱ�� is NULL And A.����ID = [1]"
            Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(!����ID)))
            If mrsInfo.EOF Then
                lng����ID = 0: Call ClearData
            Else
                lng����ID = Val(Nvl(!����ID))
                Call LoadPatiInfor: SetControlEnable
                '������ǵ���������Ϣ�����˳�����
                If mEditType <> Cr_����������Ϣ Then Exit Sub
            End If
        End If
        
        If Nvl(!ҽ����) <> "" Then txtҽ����.Text = Nvl(!ҽ����): txt��֤ҽ����.Text = txtҽ����.Text
        If zlExistOperationData(mrsInfo!����ID, "") = False Or lng����ID = 0 Then
            If Nvl(!���֤��) <> "" Then txt���֤��.Text = Nvl(!���֤��)
            If Nvl(!����) <> "" Then txtPatient.Text = Nvl(!����)
            If Nvl(!�Ա�) <> "" Then cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, Nvl(!�Ա�))
            If Nvl(!��������) <> "" Then
                txt��������.Text = Format(Nvl(!��������), "YYYY-MM-DD")
                txt����ʱ��.Text = Format(Nvl(!��������), "HH:MM")
            End If
        Else
            If Nvl(!����) <> "" And txtPatient.Text <> Nvl(!����) Then strDiff = ",����"
            If Nvl(!�Ա�) <> "" And cbo�Ա�.ListIndex <> GetCboIndex(cbo�Ա�, Nvl(!�Ա�)) Then strDiff = strDiff & ",�Ա�"
            If Nvl(!��������) <> "" And Format(Nvl(!��������), "YYYY-MM-DD HH:MM:SS") <> Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",��������"
        End If
        If Not txt�����.Locked And ExistClinicNO(Nvl(!�����), lng����ID) = False Then
            If Nvl(!�����) <> "" Then txt�����.Text = Nvl(!�����): lbl�����.Tag = Nvl(!�����)
        Else
            If Nvl(!����) <> "" And txt�����.Text <> Nvl(!�����) Then strDiff = strDiff & ",�����"
        End If
        If Nvl(!�����ص�) <> "" Then txt�����ص�.Text = Nvl(!�����ص�)
        If Nvl(!����) <> "" Then cbo����.ListIndex = GetCboIndex(cbo����, Nvl(!����))
        If Nvl(!����) <> "" Then cbo����.ListIndex = GetCboIndex(cbo����, Nvl(!����))
        If Nvl(!ְҵ) <> "" Then cboְҵ.ListIndex = GetCboIndex(cboְҵ, Nvl(!ְҵ), , , mstrCboSplit)
        If Nvl(!������λ) <> "" Then txt������λ.Text = Nvl(!������λ)
        If Nvl(!��ͥ�绰) <> "" Then txt��ͥ�绰.Text = Nvl(!��ͥ�绰)
        If Nvl(!��ϵ�˵绰) <> "" Then txt��ϵ�˵绰.Text = Nvl(!��ϵ�˵绰)
        If Nvl(!��λ�绰) <> "" Then txt��λ�绰.Text = Nvl(!��λ�绰)
        If Nvl(!��ͥ��ַ) <> "" Then txt��ͥ��ַ.Text = Nvl(!��ͥ��ַ): padd��ͥ��ַ.value = Nvl(!��ͥ��ַ)
        If Nvl(!��ͥ��ַ�ʱ�) <> "" Then txt��ͥ�ʱ�.Text = Nvl(!��ͥ��ַ�ʱ�)
        If Nvl(!���ڵ�ַ) <> "" Then txt���ڵ�ַ.Text = Nvl(!���ڵ�ַ): padd���ڵ�ַ.value = Nvl(!���ڵ�ַ)
        If Nvl(!���ڵ�ַ�ʱ�) <> "" Then txt���ڵ�ַ�ʱ�.Text = Nvl(!���ڵ�ַ�ʱ�)
        If Nvl(!��λ�ʱ�) <> "" Then txt��λ�ʱ�.Text = Nvl(!��λ�ʱ�)
        If Nvl(!��ϵ������) <> "" Then txt��ϵ������.Text = Nvl(!��ϵ������)
        If Nvl(!��ϵ�˹�ϵ) <> "" Then cbo��ϵ�˹�ϵ.ListIndex = GetCboIndex(cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ))
    End With
    Err = 0: On Error GoTo 0
    If lng����ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If strDiff <> "" Then
            strMsgInfo = "���˵� " & strDiff & " ��EMPI��Ϣ��һ�£�������ĳ��ԭ��" & vbNewLine & _
                        "     ���˷���ҽ��ҵ��;" & vbNewLine & _
                        "     ������������Ϣ��ͻ;" & vbNewLine & _
                        "     ����������Ӧ��Ȩ�ޡ�" & vbNewLine & _
                        "���β�����и��¡� "
        End If
        If strMsgInfo <> "" Then MsgBox strMsgInfo, vbInformation, gstrSysName
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function zlSaveEMPIPatiInfo(ByVal blnNewPati As Boolean, ByVal lngPatiID As Long, ByVal lngClinicID As Long, ByRef strErrMsg As String) As Boolean
    '����:�ϴ�������Ϣ��EMPIƽ̨,���ƽ̨��Ϣ����ʧ�ܣ���ͬHIS����һ�����
    '����: In-lngPatiID ����ID,lngClinicID �Һ�ID
    '      Out-strErrMsg ������Ϣ����������ʧ����Ч
    '����:True-EMPIƽ̨������Ϣ�ɹ�,False-����ʧ��
    '����:���ϴ�
    '˵��:101170
    Dim blnCharge As Boolean, lngRet As Long
    If mblnPlugin = False Then zlSaveEMPIPatiInfo = True: Exit Function
    If mobjPlugIn Is Nothing Then zlSaveEMPIPatiInfo = True: Exit Function
    If mEditType <> Cr_���� And mEditType <> Cr_�󶨿� And mEditType <> Cr_����������Ϣ Or chkCancel.value = 1 Then zlSaveEMPIPatiInfo = True: Exit Function
    If mEditType <> Cr_����������Ϣ And Not blnNewPati Then zlSaveEMPIPatiInfo = True: Exit Function
    On Error GoTo Errhand
    If mrsEMPIOut Is Nothing Then
        'EMPIû�в�����Ϣ����Ҫ�½�
        On Error Resume Next
        lngRet = mobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '�ж�ƽ̨�ش�����Ϣ�Ƿ����ı�
        With mrsEMPIOut
            If Not txt�����.Locked And ExistClinicNO(Nvl(!�����), lngPatiID) = False Then
                If txt�����.Text <> Nvl(!�����) Then blnCharge = True: GoTo EMPIModify
            End If
            If txtҽ����.Text <> Nvl(!ҽ����) Then blnCharge = True: GoTo EMPIModify
            If txt���֤��.Text <> Nvl(!���֤��) Then blnCharge = True: GoTo EMPIModify
            If zlExistOperationData(lngPatiID, "") = False Or blnNewPati Then
                If txtPatient.Text <> Nvl(!����) Then blnCharge = True: GoTo EMPIModify
                If cbo�Ա�.ListIndex <> GetCboIndex(cbo�Ա�, Nvl(!�Ա�)) Then blnCharge = True: GoTo EMPIModify
                If Format(txt��������.Text, "YYYY-MM-DD") <> Format(Nvl(!��������), "YYYY-MM-DD") Then blnCharge = True: GoTo EMPIModify
                If Format(txt����ʱ��.Text, "HH:MM") <> Format(Nvl(!��������), "HH:MM") Then blnCharge = True: GoTo EMPIModify
            End If
            If txt�����ص�.Text <> Nvl(!�����ص�) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> GetCboIndex(cbo����, Nvl(!����)) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> GetCboIndex(cbo����, Nvl(!����)) Then blnCharge = True: GoTo EMPIModify
            If cboְҵ.ListIndex <> GetCboIndex(cboְҵ, Nvl(!ְҵ), , , mstrCboSplit) Then blnCharge = True: GoTo EMPIModify
            If txt������λ.Text <> Nvl(!������λ) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ�绰.Text <> Nvl(!��ͥ�绰) Then blnCharge = True: GoTo EMPIModify
            If txt��ϵ�˵绰.Text <> Nvl(!��ϵ�˵绰) Then blnCharge = True: GoTo EMPIModify
            If txt��λ�绰.Text <> Nvl(!��λ�绰) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ��ַ.Text <> Nvl(!��ͥ��ַ) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ�ʱ�.Text <> Nvl(!��ͥ��ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txt���ڵ�ַ.Text <> Nvl(!���ڵ�ַ) Then blnCharge = True: GoTo EMPIModify
            If txt���ڵ�ַ�ʱ�.Text <> Nvl(!���ڵ�ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txt��λ�ʱ�.Text <> Nvl(!��λ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txt��ϵ������.Text <> Nvl(!��ϵ������) Then blnCharge = True: GoTo EMPIModify
            If cbo��ϵ�˹�ϵ.ListIndex <> GetCboIndex(cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ)) Then blnCharge = True: GoTo EMPIModify
        End With
    End If
EMPIModify:
    If blnCharge Then
        On Error Resume Next
        lngRet = mobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    End If
    zlSaveEMPIPatiInfo = True
    Exit Function
Errhand:
    strErrMsg = Err.Description
    Call zlPlugInErrH(Err, "zlSaveEMPIPatiInfo")
    Call SaveErrLog
End Function

