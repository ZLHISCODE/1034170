VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmReCharge 
   Caption         =   "���˷�����������"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14805
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14805
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTop 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   0
      Left            =   45
      TabIndex        =   26
      Top             =   45
      Visible         =   0   'False
      Width           =   12195
      Begin zl9InExse.ComboxExpend cboKind 
         Height          =   360
         Left            =   3915
         TabIndex        =   50
         Top             =   660
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "����"
         FontSize        =   9
      End
      Begin VB.CheckBox chk��Ŀ 
         Caption         =   "��ִ����Ŀ"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   683
         Width           =   1365
      End
      Begin VB.CheckBox chk��Ŀ 
         Caption         =   "δִ����Ŀ"
         Height          =   315
         Index           =   1
         Left            =   1590
         TabIndex        =   9
         Top             =   683
         Width           =   1350
      End
      Begin VB.Frame fraPatiInfor 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   135
         TabIndex        =   42
         Tag             =   "2700"
         Top             =   255
         Width           =   2865
         Begin VB.TextBox txtPatient 
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1155
            MaxLength       =   100
            TabIndex        =   1
            Tag             =   "1580"
            Top             =   0
            Width           =   1605
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   495
            TabIndex        =   43
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   $"frmReCharge.frx":038A
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
            ShowPropertySet =   -1  'True
            BackColor       =   -2147483633
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            Caption         =   "����"
            ForeColor       =   &H80000007&
            Height          =   210
            Index           =   7
            Left            =   0
            TabIndex        =   44
            Top             =   75
            Width           =   420
         End
      End
      Begin VB.ComboBox cbo���� 
         Height          =   330
         Left            =   10215
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   285
         Width           =   1845
      End
      Begin VB.CheckBox chkShowOthers 
         Caption         =   "��ʾ����ִ�з���"
         Height          =   315
         Left            =   6885
         TabIndex        =   10
         Top             =   683
         Width           =   2070
      End
      Begin VB.ComboBox cboBaby 
         Height          =   330
         Left            =   10215
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   675
         Width           =   1845
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "�����ڼ�"
         Height          =   255
         Left            =   9000
         TabIndex        =   6
         Top             =   323
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpApplyE 
         Height          =   360
         Left            =   6885
         TabIndex        =   5
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   94306307
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpApplyB 
         Height          =   360
         Left            =   3915
         TabIndex        =   3
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   94306307
         CurrentDate     =   36257
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�շ����"
         Height          =   210
         Left            =   3015
         TabIndex        =   49
         Top             =   735
         Width           =   840
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "�����ڼ�"
         Height          =   210
         Left            =   3015
         TabIndex        =   48
         Top             =   315
         Width           =   855
      End
      Begin VB.Label lblTo 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   210
         Left            =   6360
         TabIndex        =   47
         Top             =   315
         Width           =   255
      End
      Begin VB.Label lblPatiInfo 
         Caption         =   "�Ա�     ���䣺        סԺ�ţ�             ���ţ�         ���ң�       ������      ���ʽ�� "
         Height          =   210
         Left            =   120
         TabIndex        =   46
         Top             =   1065
         Width           =   11895
      End
      Begin VB.Label lblShowBabyFee 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������ʾ"
         Height          =   210
         Left            =   9060
         TabIndex        =   45
         Top             =   735
         Width           =   1050
      End
   End
   Begin VB.Frame fraTop 
      Height          =   1140
      Index           =   1
      Left            =   210
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   10470
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   7545
         TabIndex        =   13
         Top             =   665
         Width           =   1380
      End
      Begin VB.CheckBox chkDateAudit 
         Caption         =   "�����ڼ�"
         Height          =   255
         Left            =   6090
         TabIndex        =   35
         Top             =   705
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpAuditE 
         Height          =   360
         Left            =   3915
         TabIndex        =   12
         Top             =   660
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   94306307
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpAuditB 
         Height          =   360
         Left            =   1305
         TabIndex        =   2
         Top             =   660
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   94306307
         CurrentDate     =   36257
      End
      Begin VB.Label lblAuditDate 
         BackStyle       =   0  'Transparent
         Caption         =   "�����ڼ�"
         Height          =   210
         Left            =   420
         TabIndex        =   0
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   210
         Left            =   3555
         TabIndex        =   4
         Top             =   735
         Width           =   255
      End
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   435
      ScaleWidth      =   11295
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5070
      Width           =   11295
   End
   Begin VB.CheckBox chkVerfy 
      Caption         =   "��������ͬʱ������"
      Height          =   420
      Left            =   10650
      TabIndex        =   39
      Top             =   1590
      Width           =   2505
   End
   Begin VB.CommandButton cmdAudit 
      Caption         =   "���(&A)"
      Height          =   350
      Left            =   12945
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   33
      Tag             =   "�Ѵ���"
      Top             =   4440
      Width           =   10905
      _cx             =   19235
      _cy             =   1508
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":0420
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
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
   Begin VB.Frame fraCmd 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdCancelRefuse 
         Caption         =   "ȡ���ܾ�(&C)"
         Height          =   350
         Left            =   4440
         TabIndex        =   41
         Top             =   30
         Width           =   1350
      End
      Begin VB.CommandButton cmdOKAudit 
         Caption         =   "ȷ�����(&S)"
         Height          =   350
         Left            =   2940
         TabIndex        =   20
         Top             =   30
         Width           =   1350
      End
   End
   Begin VB.Frame fraCmd 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   6570
      Begin VB.CheckBox chkOtherOperator 
         Caption         =   "��ʾ������������"
         Height          =   315
         Left            =   3015
         TabIndex        =   40
         Top             =   30
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.CommandButton cmdCancelApply 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ������(&C)"
         Height          =   350
         Left            =   5130
         TabIndex        =   19
         ToolTipText     =   "�ȼ���F2"
         Top             =   0
         Width           =   1350
      End
      Begin VB.ComboBox cboState 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   15
         Width           =   1815
      End
      Begin VB.Label lblState 
         Caption         =   "���״̬"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.Frame fraCmd 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2535
      TabIndex        =   23
      Top             =   1620
      Visible         =   0   'False
      Width           =   8070
      Begin VB.CommandButton cmdSeleItem 
         Caption         =   "��"
         Height          =   300
         Left            =   4605
         TabIndex        =   37
         Top             =   15
         Width           =   300
      End
      Begin VB.TextBox txtFeeItem 
         Height          =   350
         Left            =   1080
         TabIndex        =   15
         Top             =   0
         Width           =   3855
      End
      Begin VB.CommandButton cmdAllDetail 
         Caption         =   "���з���(&A)"
         Height          =   350
         Left            =   5175
         TabIndex        =   16
         Top             =   0
         Width           =   1350
      End
      Begin VB.CommandButton cmdOKApply 
         Caption         =   "ȷ������(&S)"
         Height          =   350
         Left            =   6585
         TabIndex        =   17
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label lblItem 
         Caption         =   "������Ŀ"
         Height          =   255
         Left            =   195
         TabIndex        =   24
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   12945
      TabIndex        =   22
      Top             =   600
      Width           =   1350
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   12945
      TabIndex        =   21
      Top             =   120
      Width           =   1350
   End
   Begin MSComctlLib.TabStrip tbsType 
      Height          =   375
      Left            =   45
      TabIndex        =   14
      Top             =   1620
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "T1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��������ϸ"
            Key             =   "T2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   1695
      Left            =   0
      TabIndex        =   31
      Tag             =   "��ϸ"
      Top             =   5610
      Width           =   7245
      _cx             =   12779
      _cy             =   2990
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":0495
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   3495
      Index           =   0
      Left            =   0
      TabIndex        =   30
      Tag             =   "������"
      Top             =   2085
      Width           =   10905
      _cx             =   19235
      _cy             =   6165
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
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":050A
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfTogether 
      Height          =   1695
      Left            =   7320
      TabIndex        =   36
      Tag             =   "��ϸ"
      ToolTipText     =   "һ����ҩҩƷ"
      Top             =   5520
      Visible         =   0   'False
      Width           =   3570
      _cx             =   6297
      _cy             =   2990
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":057F
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   38
      Top             =   7875
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReCharge.frx":05A9
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21034
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
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
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
End
Attribute VB_Name = "frmReCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mbytUseType As Byte  '0-��������,1-ҽ�����ҵ���,2-ҽ��վ����(ֻ������ҩƷ��������˹���)
Public mbytFun As Byte      '0-����,1-���
Public mlngDeptID As Long   '��������ʱ���뵱ǰ�����Ĳ��˲���ID,ҽ�����ҵ���ʱ����ҽ������ID
Public mstrPrivs As String
Public mlngPatientID As Long '���벡��ID
Public mstrInNO As String
Public mlngAdviceID As Long
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private Const mlngModul = 1150
Private Const HeadApply = "���,4,850|��Ŀ����,1,3500|���,1,2500|����,1,2500|ҩƷ��Դ,1,1200|��λ,1,550|����,7,850|��������,7,1000|���ʽ��,7,1000|ԭʼ��������,7,0|ԭʼ���ʽ��,7,0"
Private Const HeadApplied = "ѡ��,4,850|����,1,850|�Ա�,1,550|���,1,850|��Ŀ����,1,2500|���,1,2000|����,1,2500|ҩƷ��Դ,1,1200|��λ,1,550|��������,7,1000|���ʽ��,7,1000|������,1,850|����ʱ��,1,2100"
Private Const HeadAudit = "���,4,550|����,1,850|�Ա�,1,550|���˲���,1,1100|����,1,650|���,1,850|��Ŀ����,1,2500|���,1,2000|����,1,2500|ҩƷ��Դ,1,1200|��λ,1,550|��������,7,1000|���ʽ��,7,1000|������,1,850|����ʱ��,1,2100"
Private Const HeadAudited = "״̬,4,550|����,1,850|�Ա�,1,550|���˲���,1,1200|����,1,650|���,1,850|��Ŀ����,1,2500|���,1,2000|����,1,2500|ҩƷ��Դ,1,1200|��λ,1,550|��������,7,1000|������,1,850|����ʱ��,1,2100"
Private Const HeadApplyDetail = "ִ��״̬,4,1000|Ӥ����,4,600|NO,4,1000|����ʱ��,1,2100|ִ�п���,1,1200|��������,1,1200|����,7,1250|����,7,850|����,7,850|Ӧ�ս��,7,1050|ʵ�ս��,7,1050|��������,7,1000|���ʽ��,7,1000|ԭʼ��������,7,0|ԭʼ���ʽ��,7,0"
Private Const HeadAppliedDetail = "NO,4,1000|����ʱ��,1,2100|ִ�п���,1,1200|��������,1,1200|��������,7,1000|���ʽ��,7,1000"
Private Const HeadAuditDetail = "NO,4,1000|����ʱ��,1,2100|��������,1,1200|��������,7,1000"
Private Const HeadAuditedDetail = "NO,4,1000|����ʱ��,1,2100|��������,1,1200|��������,7,1000"
Private mrsApplyDept As ADODB.Recordset
Private mblnInit As Boolean
Private mblnOperatorICU As Boolean  '��ǰ����Ա��ICU���ҵ�
Private mblnPatiDeptICU As Boolean '���˵�ǰ�����Ƿ�ΪICU����
Private mrsOperatorDept As ADODB.Recordset '����Ա����ID
Private mblnOperatorNurse As Boolean '��ǰ����Ա�Ƿ�ʿ
Private mstrOperatorDeptIDs As String  '����Ա��������ID(����Ϊ"��ʿ"��)
'���Ʊ���
Private Enum EFun
    E���� = 0
    E��� = 1
End Enum
Private Enum ESTATE
    Eȫ�� = 0
    Eδ��� = 1
    E���ͨ�� = 2
    E���δͨ�� = 3
End Enum
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
    ���ֳ�����ϸ As Boolean
    �����ѽ��ʵ��� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private ColApply As Collection
Private ColApplied As Collection
Private ColAudit As Collection
Private ColAudited As Collection

Private mbonNotEnter As Boolean
Private mlngPreFeeItemID As Long '����ʱ��¼��ǰ��
Private mstrUnitIDs As String   '����Ա��Ȩ�޵Ĳ�������ID��
Private mblnUnChange As Boolean
Private mblnNotClick As Boolean

'���ݱ���
Private mrsInfo As ADODB.Recordset
Private mrsApply As ADODB.Recordset     '������ϸ
Private mrsApplied As ADODB.Recordset   '��������ϸ
Private mrsAudit As ADODB.Recordset     '�������ϸ
Private mrsAudited As ADODB.Recordset   '�������ϸ
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mlngOldY As Long
'��Ϣ��ض������
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

'-----------------------------------------------------------------------------------
'���㿨���
Private Sub cbo����_Click()
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If cbo����.ListIndex < 0 Then Exit Sub
    If cbo����.ItemData(cbo����.ListIndex) = 0 Then Exit Sub
    If zlIsAllowFeeChange(Nvl(Val(mrsInfo!����ID)), cbo����.ItemData(cbo����.ListIndex)) = False Then Exit Sub
End Sub

'-----------------------------------------------------------------------------------
Private Sub chkDateAudit_Click()
    dtpAuditB.Enabled = chkDateAudit.Value = 0
    dtpAuditE.Enabled = dtpAuditB.Enabled
End Sub

Private Sub chkOtherOperator_Click()
    Call cboState_Click
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk��Ŀ_Click(Index As Integer)
        Dim i As Integer
        i = IIf(Index = 0, 1, 0)
        If chk��Ŀ(Index).Value = 0 Then    '����ѡһ��
            If chk��Ŀ(i).Value = 0 Then chk��Ŀ(i).Value = 1
        End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdAudit_Click()
    Dim frmTmp As New frmReCharge
    
    With frmTmp
        .mlngDeptID = mlngDeptID
        .mbytUseType = 0
        .mbytFun = 1
        .mstrPrivs = mstrPrivs
        .Show GetModuleType, Me
    End With
End Sub

Private Sub chkShowOthers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub chk��Ŀ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpApplyB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpApplyE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdSeleItem_Click()
    If zlSelectItem("") = False Then Exit Sub
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If Shift = vbCtrlMask Then
            Dim intIndex As Integer
            intIndex = IDKind.GetKindIndex("IC����")
            If intIndex < 0 Then Exit Sub
            IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        ElseIf Me.ActiveControl Is txtPatient Then
            If IDKind.Enabled Then
                If Shift = vbShiftMask Then
                    IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
                Else
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyF5 Then
        If cmdRefresh.Visible Then Call cmdRefresh_Click
        If cmdAllDetail.Visible And cmdAllDetail.Enabled Then Call cmdAllDetail_Click
    ElseIf KeyCode = vbKeyF6 Then  '��λ�����������
        txtPatient.SetFocus
        Call zlControl.TxtSelAll(txtPatient)
    End If
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXML As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
        Exit Sub
    End If
   lng�����ID = objCard.�ӿ����
    If lng�����ID = 0 Then Exit Sub
    
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
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    txtPatient.Text = strID
    Dim objCard  As Card
    Set objCard = IDKind.GetIDKindCard("����֤")
    If objCard Is Nothing Then Exit Sub
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub chkDate_KeyPress(KeyAscii As Integer)
    SendKeys "{Tab}"
End Sub


Private Sub cmdAllDetail_Click()
    If mrsApply.State = 1 Then
        If mrsApply.RecordCount > 0 Then
            mrsApply.Filter = "��������<>0"
            If mrsApply.RecordCount > 0 Then
                If MsgBox("���¶�ȡ��¼��,��ǰ���������Ϣ����ʧ,��ȷ��Ҫ������?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    If dtpApplyB.Value > dtpApplyE.Value Then
        MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ��.", vbInformation, gstrSysName
        If dtpApplyB.Visible And dtpApplyB.Enabled Then dtpApplyB.SetFocus
        Exit Sub
    End If
    Call LoadMainData(0)
    vsfMain(0).SetFocus
    Call ShowSumMoney
End Sub

Private Sub cmdRefresh_Click()
    If dtpAuditB.Value > dtpAuditE.Value Then
        MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ��.", vbInformation, gstrSysName
        If dtpAuditB.Visible And dtpAuditB.Enabled Then dtpAuditB.SetFocus
        Exit Sub
    End If
    If mbytFun = E���� Then
        Call cboState_Click
    Else
        Call LoadMainData(0)
    End If
End Sub
Private Sub cboState_Click()
    Dim strFirstCol As String, lngWidth As Long
    Dim intState As Integer
    
    If Not Visible Or cboState.ListIndex = -1 Then Exit Sub
    
    intState = Val(cboState.ItemData(cboState.ListIndex))
    
    cmdCancelApply.Visible = intState = Eδ���
        
    Call LoadMainData(0)
    
    strFirstCol = "״̬"
    chkOtherOperator.Visible = False
    Select Case intState
        Case ESTATE.Eȫ��
            lngWidth = 550
        Case ESTATE.Eδ���
            strFirstCol = "ѡ��"
            lngWidth = 550
            chkOtherOperator.Visible = InStr(1, mstrPrivsOpt, ";ȡ����������;") > 0
        Case ESTATE.E���ͨ��
            lngWidth = 0
        Case ESTATE.E���δͨ��
            lngWidth = 0
    End Select
    vsfMain(1).TextMatrix(0, ColApplied("ѡ��")) = strFirstCol
    vsfMain(1).ColWidth(ColApplied("ѡ��")) = lngWidth
    Call ShowSumMoney
End Sub

Private Sub chkDate_Click()
    
    dtpApplyB.Enabled = chkDate.Value = 0
    dtpApplyE.Enabled = chkDate.Value = 0
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub AdjustFace()
    On Error Resume Next
    
    fraTop(1).Top = fraTop(0).Top
    fraTop(1).Left = fraTop(0).Left
    
    If mbytFun = E���� Then
        Me.Caption = "���˷�����������"
        If tbsType.SelectedItem.Key = "T1" Then
            fraTop(0).Visible = True
            fraTop(1).Visible = False
            tbsType.Top = fraTop(0).Top + fraTop(0).Height + 100
            Set fraPatiInfor.Container = fraTop(0)
            fraPatiInfor.Width = Val(fraPatiInfor.Tag)
            txtPatient.Width = Val(txtPatient.Tag)
           ' fraPatiInfor.Top = dtpApplyE.Top - 10
        Else
            fraTop(0).Visible = False
            fraTop(1).Visible = True
            tbsType.Top = fraTop(1).Top + fraTop(1).Height + 100
            Set fraPatiInfor.Container = fraTop(1)
            fraPatiInfor.Width = Val(fraPatiInfor.Tag) + 520
            txtPatient.Width = Val(txtPatient.Tag) + 520
            chkDateAudit.Visible = True
        End If
        fraCmd(0).Left = tbsType.Left + tbsType.Width + 50
    Else
        Me.Caption = "���˷����������"
        fraTop(0).Visible = False
        fraTop(1).Visible = True
        tbsType.Top = fraTop(1).Top + fraTop(1).Height + 100
        Set fraPatiInfor.Container = fraTop(1)
        fraPatiInfor.Width = Val(fraPatiInfor.Tag) + 520
        txtPatient.Width = Val(txtPatient.Tag) + 520
    End If
    
    fraCmd(0).Top = tbsType.Top + (fraCmd(0).Height - tbsType.Height) \ 2
    fraCmd(0).Left = tbsType.Left + tbsType.Width + 100
    fraCmd(1).Top = fraCmd(0).Top: fraCmd(1).Left = fraCmd(0).Left
    fraCmd(2).Top = fraCmd(0).Top: fraCmd(2).Left = fraCmd(0).Left
            
    vsfMain(0).Top = tbsType.Top + tbsType.Height + 100
    If picHsc.Top - vsfMain(0).Top - 20 < 500 Then
        vsfMain(0).Height = 500
        picHsc.Top = vsfMain(0).Top + vsfMain(0).Height + 10
        vsfDetail.Top = picHsc.Top + picHsc.Height + 10
        vsfDetail.Height = stbThis.Top - vsfDetail.Top - 20
    Else
        vsfMain(0).Height = picHsc.Top - vsfMain(0).Top - 20
    End If
    vsfMain(1).Top = vsfMain(0).Top
    vsfMain(1).Height = vsfMain(0).Height
    vsfMain(1).Left = vsfMain(0).Left
    vsfMain(1).Width = vsfMain(0).Width
End Sub

Private Sub InitFace()
    Dim i As Integer
    
    Call AdjustFace
        
    tbsType.Tabs("T1").Selected = True
    
    Call InitMainHead(True)
    Call InitDetailHead(True)
    If mbytFun = E���� Then
        chkDateAudit.Visible = False
        txtPatient.ToolTipText = "��λ��ݼ�F6"
        tbsType.Tabs("T1").Caption = "��������"
        tbsType.Tabs("T2").Caption = "��������ϸ"
        
        Set ColApply = New Collection
        Set ColApplied = New Collection
        For i = 0 To vsfMain(0).Cols - 1
            ColApply.Add i, vsfMain(0).TextMatrix(0, i)
        Next
        For i = 0 To vsfMain(1).Cols - 1
            ColApplied.Add i, vsfMain(1).TextMatrix(0, i)
        Next
        
        chkVerfy.Visible = InStr(1, mstrPrivsOpt, ";�������;") > 0  '34994
        chkVerfy.Value = IIf(zlDatabase.GetPara("��������ͬʱ���", glngSys, Enum_Inside_Program.p���ʲ���, "0", Array(chkVerfy), InStr(1, mstrPrivsOpt, ";����ѡ������;") > 0) = "1", 1, 0)
        chkShowOthers.Value = IIf(zlDatabase.GetPara("��ʾ����ִ�з���", glngSys, Enum_Inside_Program.p���ʲ���, "1", Array(chkShowOthers), InStr(1, mstrPrivsOpt, ";����ѡ������;") > 0) = "1", 1, 0)
    Else
        chkVerfy.Visible = False  '34994
        chkDateAudit.Value = 1
        tbsType.Tabs("T1").Caption = "�������"
        tbsType.Tabs("T2").Caption = "�������ϸ"
        
        Set ColAudit = New Collection
        Set ColAudited = New Collection
        For i = 0 To vsfMain(0).Cols - 1
            ColAudit.Add i, vsfMain(0).TextMatrix(0, i)
        Next
        For i = 0 To vsfMain(1).Cols - 1
            ColAudited.Add i, vsfMain(1).TextMatrix(0, i)
        Next
    End If
End Sub

Private Sub InitData()
    Dim DatSys As Date
    Dim rsOperator As ADODB.Recordset
    Dim i As Long, strTmp As String, arrTmp As Variant
    
    Set mrsInfo = New ADODB.Recordset
            
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    '60679
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    strSql = "" & _
    "   Select 1 From ��Ա�� a,��Ա����˵�� b" & _
    "   Where a.ID = b.��ԱID And b.��Ա����='��ʿ'  and A.id=[1] " & _
    "           And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
    "           And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
    ""
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    mblnOperatorNurse = rsTemp.RecordCount <> 0
    rsTemp.Close
    Set rsTemp = Nothing
    Set rsOperator = GetOperatorDept
    With rsOperator
        If .RecordCount <> 0 Then .MoveFirst
        mstrOperatorDeptIDs = ""
        Do While Not .EOF
            mstrOperatorDeptIDs = mstrOperatorDeptIDs & "," & Nvl(!ID)
            .MoveNext
        Loop
        mstrOperatorDeptIDs = mstrOperatorDeptIDs & ","
    End With
    
    mstrUnitIDs = GetUserUnits
    DatSys = zlDatabase.Currentdate
    If mbytFun = E���� Then
        Set mrsApply = New ADODB.Recordset
        Set mrsApplied = New ADODB.Recordset
            
        dtpApplyB.Value = DateAdd("D", -5, DatSys)
        dtpApplyE.Value = DatSys
    
        strTmp = "0-ȫ��,1-δ���,2-���ͨ��,3-���δͨ��"
        arrTmp = Split(strTmp, ",")
        cboState.Clear
        For i = 0 To UBound(arrTmp)
            cboState.AddItem arrTmp(i)
            cboState.ItemData(cboState.NewIndex) = i
        Next
        cboState.ListIndex = 0
        cmdCancelApply.Visible = False
        
        If InStr(mstrPrivsOpt, "�������") > 0 And mbytUseType <> 2 Then
            cmdAudit.Visible = True
            cmdAudit.Top = cmdHelp.Top
            cmdHelp.Top = cmdHelp.Top + cmdHelp.Height + 100
        End If
    Else
        Set mrsAudit = New ADODB.Recordset
        Set mrsAudited = New ADODB.Recordset
    End If
    
    dtpAuditB.Value = DateAdd("D", -5, DatSys)
    dtpAuditE.Value = CDate(Format(DatSys, "yyyy-MM-dd 23:59:59"))
    
    cboKind.Clear: cboKind.AddItem "0", "�����շ����", True, True, True
    strSql = "Select ����,��� From �շ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Do While Not rsTemp.EOF
        cboKind.AddItem "" & rsTemp!����, "" & rsTemp!���, False, True, True
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If mbytFun = E���� Then
        dtpApplyB.Value = CDate(zlDatabase.GetPara("���ÿ�ʼʱ��", glngSys, mlngModul, Format(dtpApplyB.Value, "YYYY-MM-DD HH:MM:SS"), Array(dtpApplyB, dtpApplyE), zlCheckPrivs(mstrPrivsOpt, "����ѡ������")))
        chkDate.Value = IIf(zlDatabase.GetPara("�����ڼ�", glngSys, mlngModul, "0", Array(chkDate), zlCheckPrivs(mstrPrivsOpt, "����ѡ������")) = "0", 0, 1)
        i = Val(zlDatabase.GetPara("��Ŀ��ʾ��ʽ", glngSys, mlngModul, "0", Array(chk��Ŀ(0), chk��Ŀ(1)), InStr(mstrPrivsOpt, "����ѡ������")))
        Select Case i
        Case 1
            chk��Ŀ(0).Value = 1: chk��Ŀ(1).Value = 0
        Case 2
            chk��Ŀ(0).Value = 0: chk��Ŀ(1).Value = 1
        Case Else
            chk��Ŀ(0).Value = 1: chk��Ŀ(1).Value = 1
        End Select
        
        fraCmd(0).Enabled = False
        txtFeeItem.Enabled = False
        cmdAllDetail.Enabled = False
        cmdOKApply.Enabled = False
        '59051
        chkDateAudit.Value = IIf(zlDatabase.GetPara("������ϸ�����ڼ�", glngSys, Enum_Inside_Program.p���ʲ���, "0", Array(chkDateAudit), zlCheckPrivs(mstrPrivsOpt, "����ѡ������")) = "0", 0, 1)
    Else
        dtpAuditB.Value = zlDatabase.GetPara("��˿�ʼʱ��", glngSys, mlngModul, Format(dtpAuditB.Value, "YYYY-MM-DD HH:MM:SS"), Array(dtpAuditB, dtpApplyE), zlCheckPrivs(mstrPrivsOpt, "����ѡ������"))
        cmdOKAudit.Enabled = False
    End If
    
    If mlngPatientID <> 0 Then      ' And mbytFun = E����
        txtPatient.Text = "-" & mlngPatientID
        Call txtPatient_KeyPress(13)
    End If
End Sub

Private Sub InitMainHead(Optional blnSetWidth As Boolean, Optional bytScope As Byte)
'����:
'   bytScope=0-��ʼ�����ű�,1-��ʼ����һ�ű�,2-��ʼ���ڶ��ű�
    Dim i As Long, ArrTmp0 As Variant, ArrTmp1 As Variant, arrTmp As Variant
    
    If mbytFun = E���� Then
        ArrTmp0 = Split(HeadApply, "|")
        ArrTmp1 = Split(HeadApplied, "|")
    Else
        ArrTmp0 = Split(HeadAudit, "|")
        ArrTmp1 = Split(HeadAudited, "|")
    End If
    If bytScope = 0 Or bytScope = 1 Then
        With vsfMain(0)
            .Redraw = flexRDNone
            .Clear
            .RowHeightMin = 320: .Rows = 2
            .Cols = UBound(ArrTmp0) + 1
            For i = 0 To .Cols - 1
                arrTmp = Split(ArrTmp0(i), ",")
                .TextMatrix(0, i) = arrTmp(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
                If blnSetWidth Then
                    .FixedAlignment(i) = flexAlignCenterCenter
                    .ColAlignment(i) = arrTmp(1)
                    .ColWidth(i) = arrTmp(2)
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
    
    If bytScope = 0 Or bytScope = 2 Then
        With vsfMain(1)
            .Redraw = flexRDNone
            .Clear
            .RowHeightMin = 320
            .Rows = 2
            .Cols = UBound(ArrTmp1) + 1
            For i = 0 To .Cols - 1
                arrTmp = Split(ArrTmp1(i), ",")
                .TextMatrix(0, i) = arrTmp(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
                If blnSetWidth Then
                    .FixedAlignment(i) = flexAlignCenterCenter
                    .ColAlignment(i) = arrTmp(1)
                    .ColWidth(i) = arrTmp(2)
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
End Sub

Private Sub InitDetailHead(Optional blnSetWidth As Boolean)
    Dim ArrTmpDetail As Variant, arrTmp As Variant
    Dim i As Long
    Dim strHead As String
    
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            strHead = HeadApplyDetail
        Else
            strHead = HeadAppliedDetail
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            strHead = HeadAuditDetail
        Else
            strHead = HeadAuditedDetail
        End If
    End If
    
    vsfDetail.Clear
    vsfDetail.Rows = 2
    vsfDetail.RowHeightMin = 320
    ArrTmpDetail = Split(strHead, "|")
    vsfDetail.Cols = UBound(ArrTmpDetail) + 1
    
     
    With vsfDetail
        For i = 0 To .Cols - 1
            arrTmp = Split(ArrTmpDetail(i), ",")
            .TextMatrix(0, i) = arrTmp(0)
            .ColKey(i) = .TextMatrix(0, i)
             
            
            If blnSetWidth Then
                .FixedAlignment(i) = flexAlignCenterCenter
                .ColAlignment(i) = arrTmp(1)
                .ColWidth(i) = arrTmp(2)
            End If
        Next
    End With
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    gblnOK = False
    '55368
    Call LoadBabyCombox(0)
    mblnOperatorICU = zlisCheckOperatorICU
    Call initCardSquareData
    
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���) & ";" & mstrPrivs
    Call RestoreWinState(Me, App.ProductName)
    Call InitFace
     '����:39373
     '55368
    Call RestoreFlexState(vsfMain(0), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call RestoreFlexState(vsfMain(1), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Me.WindowState = vbMaximized
     '����:47798
    Call GetRegisterItem(g˽��ģ��, Me.Name, "idkind", strTmp)
    Err = 0: On Error Resume Next
    IDKind.IDKind = Val(strTmp)
    Err = 0: On Error GoTo 0
    
    Call InitData
    Call zlMsgModule_Init
    If mstrInNO <> "" Or mlngAdviceID <> 0 Then
        mblnInit = True
        Call LoadMainData(0, mstrInNO, mlngAdviceID)
        mblnInit = False
        mstrInNO = ""
        mlngAdviceID = 0
    End If
End Sub



Private Sub Form_Resize()
    Dim lngTmp As Long
    
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
        
    vsfMain(0).Left = Me.ScaleLeft + 20
    vsfMain(0).Width = Me.ScaleLeft + Me.ScaleWidth - vsfMain(0).Left - 20
    vsfMain(1).Left = vsfMain(0).Left
    vsfMain(1).Width = vsfMain(0).Width
    vsfDetail.Left = vsfMain(0).Left
    vsfDetail.Width = vsfMain(0).Width - IIf(vsfTogether.Visible, vsfTogether.Width + 50, 0)
    picHsc.Width = vsfMain(0).Width
    
    If vsfMain(0).Visible Then
        lngTmp = Me.ScaleTop + Me.ScaleHeight - (picHsc.Height + vsfDetail.Height + stbThis.Height + 30) - vsfMain(0).Top
        
        If lngTmp > 500 Then
            vsfMain(0).Height = lngTmp
            picHsc.Top = vsfMain(0).Top + vsfMain(0).Height + 10
            vsfDetail.Top = picHsc.Top + picHsc.Height + 10
        End If
    ElseIf vsfMain(1).Visible Then
        lngTmp = Me.ScaleTop + Me.ScaleHeight - (picHsc.Height + vsfDetail.Height + stbThis.Height + 30) - vsfMain(1).Top
        If lngTmp > 500 Then
            vsfMain(1).Height = lngTmp
            picHsc.Top = vsfMain(0).Top + vsfMain(1).Height + 10
            vsfDetail.Top = picHsc.Top + picHsc.Height + 10
        End If
    End If
    
    If mbytFun = EFun.E���� Then
        If vsfTogether.Visible Then
            vsfTogether.Top = vsfDetail.Top
            vsfTogether.Height = vsfDetail.Height
            vsfTogether.Left = vsfDetail.Left + vsfDetail.Width + 50
        End If
        chkVerfy.Top = fraCmd(0).Top + 15
        chkVerfy.Width = IIf(Me.ScaleWidth - chkVerfy.Left > 2555, 2505, Me.ScaleWidth - chkVerfy.Left)
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveFlexState(vsfMain(0), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call SaveFlexState(vsfMain(1), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call SaveFlexState(vsfDetail, App.ProductName & "\" & Me.Name & "-" & mbytFun)
    '55368
    Call zlDatabase.SetPara("��������Ӥ������ʾ����", cboBaby.ItemData(cboBaby.ListIndex), glngSys, Enum_Inside_Program.p���ʲ���, InStr(mstrPrivsOpt, ";����ѡ������;") > 0)
    If mbytFun = E���� Then
        zlDatabase.SetPara "���ÿ�ʼʱ��", Format(dtpApplyB.Value, "YYYY-MM-DD HH:MM:SS"), glngSys, mlngModul, zlCheckPrivs(mstrPrivsOpt, "����ѡ������")
        zlDatabase.SetPara "�����ڼ�", chkDate.Value, glngSys, mlngModul
        zlDatabase.SetPara "��Ŀ��ʾ��ʽ", IIf(chk��Ŀ(0).Value = 1 And chk��Ŀ(1).Value = 0, 1, IIf(chk��Ŀ(0).Value = 0 And chk��Ŀ(1).Value = 1, 2, 0)), glngSys, mlngModul, zlCheckPrivs(mstrPrivsOpt, "����ѡ������")
        zlDatabase.SetPara "��������ͬʱ���", IIf(chkVerfy.Value = 1, 1, 0), glngSys, Enum_Inside_Program.p���ʲ���, zlCheckPrivs(mstrPrivsOpt, "����ѡ������")
        zlDatabase.SetPara "������ϸ�����ڼ�", chkDateAudit.Value, glngSys, Enum_Inside_Program.p���ʲ���, zlCheckPrivs(mstrPrivsOpt, "����ѡ������")
        zlDatabase.SetPara "��ʾ����ִ�з���", IIf(chkShowOthers.Value = 1, 1, 0), glngSys, Enum_Inside_Program.p���ʲ���, zlCheckPrivs(mstrPrivsOpt, "����ѡ������")
    Else
        zlDatabase.SetPara "��˿�ʼʱ��", Format(dtpAuditB.Value, "YYYY-MM-DD HH:MM:SS"), glngSys, mlngModul, zlCheckPrivs(mstrPrivsOpt, "����ѡ������")
    End If
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    
    mlngDeptID = 0
    mlngPatientID = 0
     '����:47798
    Call SaveRegisterItem(g˽��ģ��, Me.Name, "idkind", IDKind.IDKind)
    Set mrsOperatorDept = Nothing
    mblnOperatorNurse = False
    mstrOperatorDeptIDs = ""
    
    '��Ϣ��ж
    zlMsgModule_Unload
End Sub

Private Sub picHsc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngOldY = Y
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsfMain(0).Height + Y - mlngOldY <= 500 Or vsfDetail.Height - Y + mlngOldY <= 500 Then Exit Sub
        
        picHsc.Top = picHsc.Top + Y - mlngOldY
        If vsfMain(0).Visible Then
            vsfMain(0).Height = picHsc.Top - vsfMain(0).Top ' vsfMain(0).vsfMain(0).Height + Y
            vsfMain(1).Height = vsfMain(0).Height
        Else
            vsfMain(1).Height = picHsc.Top - vsfMain(1).Top ' vsfMain(1).Height + Y
            vsfMain(0).Height = vsfMain(1).Height
        End If
        
        vsfDetail.Top = picHsc.Top + picHsc.Height ' vsfDetail.Top + Y
        vsfDetail.Height = IIf(ScaleHeight - vsfDetail.Top - stbThis.Height < 0, 0, ScaleHeight - vsfDetail.Top - stbThis.Height) ' vsfDetail.Height - Y
        
        Me.Refresh
    End If
End Sub



Private Sub txtFeeItem_Change()
    txtFeeItem.Tag = ""
End Sub

Private Sub txtFeeItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtFeeItem.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txtFeeItem.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zlSelectItem(Trim(txtFeeItem.Text)) = False Then Exit Sub
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    Call IDKind.SetAutoReadCard(False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If Trim(txtPatient.Text) = "" Then
         If mbytFun = E���� Then
            If tbsType.SelectedItem.Key <> "T1" Then
                Call ClearPatientInfo
            End If
        Else
                Call ClearPatientInfo
        End If
    End If
    
    If mrsInfo.State = 0 And Trim(txtPatient.Text) <> "" Then txtPatient.Text = ""
    If mrsInfo.State = 1 Then
        If txtPatient.Text <> mrsInfo!���� Then txtPatient.Text = mrsInfo!����
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    Dim blnOutMsg  As Boolean
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            .mlngUnitID = mlngDeptID
            .mbytUseType = 4
            .mstrPrivs = mstrPrivs
            Set .mfrmParent = Me
            .Show 1, Me
        End With
    Else
        If IDKind.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
            End If
        Else
            If IDKind.GetCurCard.�ӿ���� <> 0 Then
                blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
            End If
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    'ˢ����ϻ���������س�
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
      End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOutMsg As Boolean
    Call ClearPatientInfo
    If Not GetPatient(objCard, strInput, blnCard, blnOutMsg) Then
        If Not blnOutMsg Then stbThis.Panels(2).Text = "û���ҵ��ò���,������������!"
        Call zlControl.TxtSelAll(txtPatient)
        Exit Sub
    End If
    If Not IsNull(mrsInfo!����) Then
        MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(mrsInfo!����))
        MCPAR.���ֳ�����ϸ = gclsInsure.GetCapability(support�������ֳ�����ϸ, , Val(mrsInfo!����))
        MCPAR.�����ѽ��ʵ��� = gclsInsure.GetCapability(support���������ѽ��ʵļ��ʵ���, , Val(mrsInfo!����))
        If MCPAR.���������ϴ� Then
            If Not gclsInsure.GetCapability(support�������ݳ�������, , Val(mrsInfo!����)) Then  '���ܲ�������
                MsgBox "��ǰҽ�����������ֳ������ݣ���֧�ֲ����������ģʽ���ʣ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    stbThis.Panels(2).Text = ""
    Call LoadPatientInfo
     zlCommFun.PressKey vbKeyTab
 End Sub


Private Sub ClearPatientInfo()
    Set mrsInfo = New ADODB.Recordset
    
    txtPatient.ForeColor = Me.ForeColor
    lblPatiInfo.Caption = "�Ա�     ���䣺        סԺ�ţ�             ���ţ�         ���ң�       ������        ���ʽ�� "

    
    fraCmd(0).Enabled = False
    txtFeeItem.Enabled = False
    cmdAllDetail.Enabled = False
    cmdOKApply.Enabled = False
    
    If vsfMain(0).Rows >= 2 Then
        If Val(vsfMain(0).RowData(1)) <> 0 Then
            Call InitMainHead(False, 1)
            Call InitDetailHead(False)
            Set mrsApply = New ADODB.Recordset
        End If
    End If
End Sub

Private Sub LoadPatientInfo()
    txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
    mblnNotClick = True
    txtPatient.Text = mrsInfo!����
    lblPatiInfo.Caption = "�Ա�" & mrsInfo!�Ա� & "   ���䣺" & mrsInfo!���� & "   סԺ�ţ�" & mrsInfo!סԺ�� & _
                          "   ���ţ�" & mrsInfo!���� & "   ���ң�" & mrsInfo!���� & "   ������" & mrsInfo!���� & "   ���ʽ��" & mrsInfo!ҽ�Ƹ��ʽ

    fraCmd(0).Enabled = True
    fraCmd(0).Enabled = True
    txtFeeItem.Enabled = True
    cmdAllDetail.Enabled = True
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    mblnNotClick = False
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:blnOutMsg-true�����Ѿ���ʾ,����δ��ʾ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 16:53:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strIF As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    '�Ƿ����ǿ�Ƽ���Ȩ��
    If InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
    Else
        strIF = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
    End If
    
    strSql = _
    "   Select A.����ID,B.��ҳID,B.��Ժ����ID," & _
    "          Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, A.����, A.סԺ��, B.��Ժ���� ����, " & _
    "          C.���� ����, D.���� ����, A.ҽ�Ƹ��ʽ, B.����,B.��������,a.�����,A.����֤��" & vbNewLine & _
    "   From ������Ϣ A, ������ҳ B, ���ű� C, ���ű� D,������� X" & vbNewLine & _
    "   Where A.����id = B.����id And A.��ҳID = B.��ҳID And B.��Ժ����ID = C.ID And B.��ǰ����id = D.ID" & _
    "       And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2  And A.ͣ��ʱ�� Is Null" & strIF
        
        '����:38332:ȡ��վ������,��Ϊ���ܴ��ڶ�ת�����˵Ĵ���
'        " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _

    If blnCard = True And objCard.���� Like "����*" Then   'ˢ��
    
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSql = strSql & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSql = strSql & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSql = strSql & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSql = strSql & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "/" And mbytUseType <> 1 And mlngDeptID <> 0 Then   '��λ��,ҽ�����ҵ���ʱ��ʹ�ô���,�������ý���ʱѡ���в���ʱ��ʹ�ô���
        '41654 And IsNumeric(Mid(strInput, 2))
        strSql = strSql & " And B.��ǰ����ID=[3] And B.��Ժ����=[1]"
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                strPati = "" & _
                "   Select A.����ID as ID,A.����ID,A.סԺ��, A.�����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, A.����, A.סԺ����, A.��ͥ��ַ, A.������λ," & vbNewLine & _
                "       To_Char(A.��������,'YYYY-MM-DD') as ��������,  To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����, To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����" & vbNewLine & _
                "   From ������Ϣ A, ������ҳ B,������� X" & vbNewLine & _
                "   Where A.����id = B.����id(+) And A.��ҳID = B.��ҳid(+) And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2 And A.ͣ��ʱ�� Is Null And A.���� = [1]" & strIF & vbNewLine & _
                "   Order By Decode(סԺ��, Null, 1, 0), ��Ժ���� Desc"
                        
                vRect = GetControlRect(txtPatient.hWnd)
                
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!����ID)
                    strSql = strSql & " And A.����ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset:  Exit Function
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSql = strSql & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.סԺ��=[2]"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSql = strSql & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Mid(strInput, 2), strInput, mlngDeptID)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: Exit Function
    If zlPatiIS�����ѱ�Ŀ(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = True Then    '����:28725
        Set mrsInfo = New ADODB.Recordset
        blnOutMsg = True
        Exit Function
    End If
    mblnPatiDeptICU = zlisCheckDeptICU(Val(Nvl(mrsInfo!��Ժ����ID)))
    Call LoadסԺ����
    
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function
Private Sub LoadסԺ����()
    Dim rsTemp As ADODB.Recordset
    With cbo����
        .Clear
        .AddItem "����סԺ"
        .ListIndex = 0
    End With
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    gstrSQL = "select ��ҳID From ������ҳ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Nvl(mrsInfo!����ID)))
    With cbo����
        Do While Not rsTemp.EOF
             .AddItem "��" & Val(Nvl(rsTemp!��ҳID)) & "��סԺ"
             .ItemData(.NewIndex) = Val(Nvl(rsTemp!��ҳID))
             If Val(Nvl(mrsInfo!��ҳID)) = Val(Nvl(rsTemp!��ҳID)) Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub
Private Sub SetWindowsTittle()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô������
    '����:���˺�
    '����:2009-10-26 15:21:22
    '����:25850
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case mbytFun
    Case E����
        
        'mbytFun As Byte      '0-����,1-���
        'mlngDeptID As Long   '��������ʱ���뵱ǰ�����Ĳ��˲���ID,ҽ�����ҵ���ʱ����ҽ������ID
        If mlngDeptID = 0 Then
            fraTop(0).ForeColor = vbRed
            If tbsType.SelectedItem.Key = "T1" Then
               fraTop(0).Caption = "���벿�ţ�" & "���벿��δѡ��!"
            Else
                fraTop(0).Caption = ""
            End If
        Else
            fraTop(0).ForeColor = vbRed
            If tbsType.SelectedItem.Key = "T1" Then
               fraTop(0).Caption = "���벿�ţ�" & "���벿��δѡ��!"
                If mrsApplyDept Is Nothing Then
                    GoTo GetApplyDept:
                ElseIf mrsApplyDept.State <> 1 Then
                    GoTo GetApplyDept:
                Else
                    fraTop(0).Caption = "���벿�ţ�" & "���벿��δѡ��!"
                    If mrsApplyDept.EOF = False Then
                        fraTop(0).Caption = "���벿�ţ�" & Nvl(mrsApplyDept!����)
                        fraTop(0).ForeColor = &H80000012
                    End If
                End If
            Else
                fraTop(0).Caption = ""
            End If
        End If
    Case Else
        fraTop(0).Caption = ""
        fraTop(0).ForeColor = &H80000012
    End Select
    Exit Sub
GetApplyDept:

    On Error GoTo errHandle
    gstrSQL = "Select ����  From ���ű� where id=[1]"
    Set mrsApplyDept = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptID)
    fraTop(0).Caption = "���벿�ţ�" & "���벿��δѡ��!"
    If mrsApplyDept.EOF = False Then
        fraTop(0).Caption = "���벿�ţ�" & Nvl(mrsApplyDept!����)
        fraTop(0).ForeColor = &H80000012
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub tbsType_Click()
    Dim lngFeeItemID As Long
    
    Me.AutoRedraw = False
    Call AdjustFace
    
    Call SetWindowsTittle
    
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            vsfMain(0).Visible = True
            vsfMain(1).Visible = False
            fraCmd(0).Visible = True
            fraCmd(1).Visible = False
            chkVerfy.Visible = InStr(1, mstrPrivsOpt, ";�������;") > 0  '34994
            If Visible Then
                vsfMain(0).SetFocus
                lngFeeItemID = vsfMain(0).RowData(vsfMain(0).Row)
                Call ShowDetail(lngFeeItemID)
            End If
        Else
            chkVerfy.Visible = False '34994
            vsfMain(0).Visible = False
            vsfMain(1).Visible = True
            fraCmd(0).Visible = False
            fraCmd(1).Visible = True
            If Visible Then
                vsfMain(1).SetFocus
                Call cmdRefresh_Click
            End If
        End If
        Call Form_Resize
        Call ShowSumMoney
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            lblAuditDate.Caption = "�����ڼ�"
            vsfMain(0).Visible = True
            vsfMain(1).Visible = False
            fraCmd(2).Visible = True
            cmdOKAudit.Caption = "ȷ�����(&S)"
            cmdCancelRefuse.Visible = False
            chkDateAudit.Visible = True
            Call chkDateAudit_Click
            
            If Visible Then
                vsfMain(0).SetFocus
                lngFeeItemID = vsfMain(0).RowData(vsfMain(0).Row)
                Call ShowDetail(lngFeeItemID)
            End If
        Else
            lblAuditDate.Caption = "����ڼ�"
            vsfMain(0).Visible = False
            vsfMain(1).Visible = True
            fraCmd(2).Visible = True
            cmdOKAudit.Caption = "����ܾ�(&S)"
            cmdCancelRefuse.Visible = True
            
            chkDateAudit.Visible = False
            dtpAuditB.Enabled = True
            dtpAuditE.Enabled = dtpAuditB.Enabled
            
            
            If Visible Then
                vsfMain(1).SetFocus
                Call cmdRefresh_Click
            End If
        End If
        Call ShowSumMoney
    End If
    Me.AutoRedraw = True
End Sub


Private Sub txtFeeItem_GotFocus()
    zlControl.TxtSelAll txtFeeItem
End Sub

Private Sub txtFeeItem_KeyPress(KeyAscii As Integer)

'
'    Dim rsTmp As ADODB.Recordset, strSQL As String
'    Dim strIF As String, strInput As String, strMatch As String
'    Dim vRect As RECT
'    Dim blnCancel As Boolean
'
'    If KeyAscii <> vbKeyReturn Then Exit Sub
'
'    strInput = UCase(Trim(txtFeeItem.Text))
'    If strInput = "" Then Exit Sub
'    strMatch = IIf(Len(strInput) < 3, "", gstrLike)
'
'    If zlCommFun.IsNumOrChar(strInput) Then
'        strIF = " And (A.���� like [1] Or B.���� like [1] And B.���� in(3," & gbytCode + 1 & "))"
'    Else
'        strIF = " And B.���� like [1]"
'    End If
'    strSQL = "Select Distinct A.ID,A.���� ,B.���� ,A.��� " & _
'          " From �շ���ĿĿ¼ A,�շ���Ŀ���� B Where A.id=B.�շ�ϸĿID " & strIF & _
'          " And rownum<101 Order by ����"
'
'    vRect = GetControlRect(txtFeeItem.hwnd)
'    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ʷ�����Ŀ", 1, "", "��ѡ��", False, False, True, vRect.Left, vRect.Top, txtFeeItem.Height, blnCancel, False, True, strMatch & strInput & "%")
'
'    If Not rsTmp Is Nothing Then
'        txtFeeItem.Text = rsTmp!����
'        Call LoadMainData(rsTmp!ID)
'        stbThis.Panels(2).Text = ""
'    Else
'        stbThis.Panels(2).Text = "δ�ҵ�����Ŀ!"
'    End If
'    txtFeeItem.SelStart = 0
'    txtFeeItem.SelLength = Len(txtFeeItem.Text)
End Sub
Private Function zlSelectItem(ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ����������Ŀ
    '���:strKey-��������
    '����:
    '����:ѡ��ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2009-09-21 14:23:25
    '����:25182
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strIF As String, strSql As String, DatBegin As Date, DatEnd As Date, strWhere As String
    Dim strSQLDetail As String, strSearch As String, vRect As RECT, blnCancel As Boolean
    Dim strDosage As String '��ҩ��������
    Dim lng��ҳID As Long
    Dim intBaby As Integer
    '59220
    strIF = " And A.����id = [1] And A.��¼״̬ > 0"
    '����:39373
    '55368
    intBaby = cboBaby.ItemData(cboBaby.ListIndex)
    Select Case intBaby
    Case 0  '����Ӥ����
        strIF = strIF & " And nvl(A.Ӥ����,0)= 0"
    Case 1  '��Ӥ����
    Case Else '��ʾ�ڼ���Ӥ����
        strIF = strIF & " And nvl(A.Ӥ����,0)= [9]"
    End Select
    '����:40304
    lng��ҳID = 0
    If cbo����.ListIndex >= 0 Then
         lng��ҳID = cbo����.ItemData(cbo����.ListIndex)
    End If
    strIF = strIF & IIf(lng��ҳID = 0, "", " And nvl(A.��ҳID,0)= [8]")
        
    If mlngDeptID <> 0 Then
        If mbytUseType <> 1 Then
            If Not mblnOperatorICU Then
                strIF = strIF & " And Instr(','||[6]||',',','||A.���˲���id||',')>0"
                '����:43940:���ڻ���ҽ��Ҳ���ڿ�������<>���˿��ҵ����,���, _
                '       �����������,ֱ���Կ�������ID�Ƿ�Ϊ�ٴ������ж�, '
                '       �����ò��˿���ID=��������ID���ж��Ƿ�Ϊ�ٴ����ĵ���
                 'exists(select 1 From ��������˵�� where A.��������id=����ID And ��������='�ٴ�')
                
                '����:36462
                strIF = strIF & " And (exists(select 1 From ��������˵�� where A.��������id=����ID And ��������='�ٴ�') And " & _
                         "       (Instr(',5,6,7,', ',' || A.�շ���� || ',') > 0 Or (A.�շ���� = '4' And Nvl(C.��������, 0) = 1)) Or " & _
                         "       (Instr(',5,6,7,', ',' || A.�շ���� || ',') = 0 Or A.�շ���� = '4' And Nvl(C.��������, 0) = 0))"
            ElseIf Not mblnPatiDeptICU Then
                '�Ե�ʱ���˿����Ƿ�ΪICU����:42526
                strIF = strIF & "  And (  exists(Select 1 From  ��������˵�� J1  Where A.���˿���ID=J1.����ID And J1.��������='ICU') "
                strIF = strIF & "  or  (exists(select 1 From ��������˵�� where A.��������id=����ID And ��������='�ٴ�') And " & _
                         "       (Instr(',5,6,7,', ',' || A.�շ���� || ',') > 0 Or (A.�շ���� = '4' And Nvl(C.��������, 0) = 1)) Or " & _
                         "       (Instr(',5,6,7,', ',' || A.�շ���� || ',') = 0 Or A.�շ���� = '4' And Nvl(C.��������, 0) = 0)) )"
            End If
        Else
            strIF = strIF & " And A.��������id+0 = [2]"
        End If
    End If
      If chkDate.Value = 0 Then
        If dtpApplyB.Value <= dtpApplyE.Value Then
            DatBegin = dtpApplyB.Value
            DatEnd = dtpApplyE.Value
        Else
            DatBegin = dtpApplyE.Value
            DatEnd = dtpApplyB.Value
        End If
        '59220
        strIF = strIF & " And A.����ʱ��+0 Between [4] And [5]"
    End If
    '36391:��1�滻ΪRowNum,����Oracle��ͼ�Զ��ϲ�:42333
    '77686,���ϴ�,2014/9/18,�����������
    strDosage = " And Not Exists (Select  Rownum as ��� From סԺ���ü�¼ J, ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1 Where j.NO = a.NO And a.��¼���� = j.��¼���� And  nvl(A.�۸񸸺�, A.���) = Nvl(J.�۸񸸺�, J.���) And B1.����id = j.ID And B1.ID = C1.�շ�id And instr( ',8,9,10,21,24,25,26,',','||B1.����||',')>0)  "
    
'    '����:29887,55380
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    blnYP = zlCheckPrivs(mstrPrivsOpt, "ҩƷ��������")
    blnZL = zlCheckPrivs(mstrPrivsOpt, "������������")
    blnWC = zlCheckPrivs(mstrPrivsOpt, "������������")
    
  If blnYP And blnWC And blnZL Then
        'ȫ��,������
    ElseIf blnYP And blnWC And blnZL = False Then
        strIF = strIF & "  And  A.�շ���� In('4','5','6','7')"
    ElseIf blnYP And blnWC = False And blnZL Then
        strIF = strIF & "  And  A.�շ���� <>'4'"
    ElseIf blnYP And blnWC = False And blnZL = False Then
        strIF = strIF & "  And  A.�շ���� In('5','6','7')"
    ElseIf blnYP = False And blnWC And blnZL = False Then
        strIF = strIF & "  And  A.�շ���� ='4'"
    ElseIf blnYP = False And blnWC And blnZL Then
        strIF = strIF & "  And instr( '5,6,7',  A.�շ����)=0 "
    ElseIf blnYP = False And blnWC = False And blnZL Then
        strIF = strIF & "  And instr( '4,5,6,7',  A.�շ����)=0 "
    Else
        MsgBox "ע��:" & vbCrLf & "  �㲻�߱�ҩƷ�����ļ��������������Ȩ��,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strWhere = ""
    '����:30523
    '������ǰ��δִ�е�ҩƷ������,ֻ�ܶԲ�����������ҩƷ�����Ľ�������,����ҩ����ҩ�����ĵ�,���ܹ����봦��,��˳��ֳ����ڴ˻����ϵ����̴���©��.
    '��������ȡ���˸�����,���ڵĴ�����ʽ�����ҩ������δִ�е�,������ʱ,��˲���ֻ��Ϊ����(����ڲ������ǰ,��ҩƷ��ҩƷִ��,���ֹ���),ִ���˵�,��Ϊִ�в���.
    If chk��Ŀ(0).Value = 1 And chk��Ŀ(1).Value = 0 Then 'ֻ��ʾ��ִ�е�
        strWhere = "     And Exists (  Select 1 From סԺ���ü�¼ B  Where A.NO = B.NO And A.��¼���� = B.��¼���� And Nvl(A.�۸񸸺�, A.���) = Nvl(B.�۸񸸺�, B.���)  And B.ִ��״̬ <> 0 )" & vbNewLine
    ElseIf chk��Ŀ(0).Value = 0 And chk��Ŀ(1).Value = 1 Then 'ֻ��ʾδִ�е�
        strWhere = "     And Exists (  Select 1 From סԺ���ü�¼ B  Where A.NO = B.NO And A.��¼���� = B.��¼���� And Nvl(A.�۸񸸺�, A.���) = Nvl(B.�۸񸸺�, B.���)  And B.ִ��״̬ = 0 )" & vbNewLine
    ElseIf chk��Ŀ(0).Value = 0 And chk��Ŀ(1).Value = 0 Then 'δѡ��ִ����Ŀ��,ȱʡΪȫѡ
    Else
    End If
    
    If strKey <> "" Then
        strSearch = IIf(Len(strKey) < 3, "", gstrLike) & strKey & "%"
        If zlCommFun.IsNumOrChar(strKey) Then
            strIF = strIF & vbCrLf & " And  exists(Select 1 From  �շ���ĿĿ¼ Q1,�շ���Ŀ���� Q2 " & vbCrLf & _
            "                                      where Q1.ID=Q2.�շ�ϸĿID and A.�շ�ϸĿid=Q1.id And (Q1.���� like upper([7]) or ( Q2.���� like upper([7]) and Q2.���� in (3," & gbytCode + 1 & "))))  "
        Else
            strIF = strIF & vbCrLf & " And  exists(Select 1 From  �շ���ĿĿ¼ Q1,�շ���Ŀ���� Q2 " & vbCrLf & _
            "                                      where Q1.ID=Q2.�շ�ϸĿID and A.�շ�ϸĿid=Q1.id And Q2.���� like upper([7]))"
        End If
    End If

    'δ���ʵ�(���ʲ����ϵ�δ����)
    strIF = strIF & " And (A.NO, Nvl(A.�۸񸸺�, A.���)) In (Select A.No ,Nvl(A.�۸񸸺�, A.���)" & vbNewLine & _
            "From סԺ���ü�¼ A" & vbNewLine & _
            "Where Mod(A.��¼����, 10) = 2 " & strIF & vbNewLine & _
            "Group By A.NO, Mod(A.��¼����, 10), Nvl(A.�۸񸸺�, A.���)" & vbNewLine & _
            "Having Nvl(Sum(���ʽ��),0) = 0)"

    'δ���������������
    '�˹�ҩ��,��Ϊ�˵�ʱ��ֻ������,���Ը�����׼,��ȡ1
    '�����ҩƷ,û�з�ҩ�Ĳ���������,���ܷ�ҩ������ҩ��,����Ҫ��Exists�Ӳ�ѯ�ж�,����ֱ����ִ��״̬<>0
    strSQLDetail = "Select Max(ID) ID, NO, ����ʱ��, ���, ִ�в���id,��������id, �շ����, �շ�ϸĿid, Avg(����) ����," & vbNewLine & _
            "       Decode(Sign(Min(ִ��״̬)), -1, 1, Sum(����)) ����," & vbNewLine & _
            "       Decode(Sign(Min(ִ��״̬)), -1, Sum(���� * ����), Sum(����)) ����, Sum(Ӧ�ս��) Ӧ�ս��, Sum(ʵ�ս��) ʵ�ս��, ����ID, ҽ�����" & vbNewLine & _
            "From (Select Max(Decode(Sign(A.ִ��״̬), -1, 0, Decode(A.�۸񸸺�, Null, A.ID, 0))) ID, A.ִ��״̬, A.����ʱ��, A.NO," & vbNewLine & _
            "              Nvl(A.�۸񸸺�, A.���) As ���, A.ִ�в���id,A.��������id, A.�շ����, A.�շ�ϸĿid, Avg(A.��׼����) ����," & vbNewLine & _
            "              Avg(A.����) ����, Avg(A.����) ����, Sum(A.Ӧ�ս��) Ӧ�ս��, Sum(A.ʵ�ս��) ʵ�ս��, A.����ID, A.ҽ�����" & vbNewLine & _
            "       From סԺ���ü�¼ A, �������� C" & vbNewLine & _
            "       Where A.�շ�ϸĿid = C.����id(+)  " & strDosage & vbNewLine & strWhere & _
            "             And A.��¼���� = 2 " & vbCrLf & strIF & vbNewLine & _
            "       Group By A.NO, A.ִ��״̬, Nvl(A.�۸񸸺�, A.���), A.����ʱ��, A.ִ�в���id,A.��������id, A.�շ����, A.�շ�ϸĿid, A.����ID, A.ҽ�����)" & vbNewLine & _
            "Group By NO, ����ʱ��, ���, ִ�в���id,��������id, �շ����, �շ�ϸĿid, ����ID, ҽ�����" & vbNewLine & _
            "Having Sum(���� * ����) <> 0 "
    
    '���������ʵ���ϸ
    strSql = "" & _
    "   Select Distinct  M1.ID,M1.����,M2.����,M1.��� " & _
    "   From (" & strSQLDetail & ") A, ���˷������� C,�շ���ĿĿ¼ M1,�շ���Ŀ���� M2" & vbNewLine & _
    "   Where   A.�շ�ϸĿID=M1.ID And A.�շ�ϸĿID=M2.�շ�ϸĿID and A.ID = C.����id(+) And  C.״̬(+) = 0"
    
    
    vRect = GetControlRect(txtFeeItem.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "���ʷ�����Ŀ", 1, " ", "��ѡ��", False, False, True, vRect.Left, vRect.Top, txtFeeItem.Height, blnCancel, False, True, _
    Val(mrsInfo!����ID), mlngDeptID, 0, DatBegin, DatEnd, mstrUnitIDs, strSearch, lng��ҳID, intBaby - 1)
    If blnCancel Then
        zlControl.TxtSelAll txtFeeItem
        If txtFeeItem.Enabled And txtFeeItem.Visible Then txtFeeItem.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ���Ŀ,���ܴ˲���δ�����˷��ã�����!", vbInformation + vbDefaultButton1, gstrSysName
        zlControl.TxtSelAll txtFeeItem
        If txtFeeItem.Enabled And txtFeeItem.Visible Then txtFeeItem.SetFocus
        Exit Function
    End If
        
    '������ط�����Ϣ����
    txtFeeItem.Text = Nvl(rsTemp!����): txtFeeItem.Tag = Nvl(rsTemp!ID)
    Call LoadMainData(rsTemp!ID)
    stbThis.Panels(2).Text = ""
    zlControl.TxtSelAll txtFeeItem
    DoEvents
    If txtFeeItem.Enabled And txtFeeItem.Visible Then txtFeeItem.SetFocus
    zlSelectItem = True
End Function

Private Sub LoadMainData(ByVal lngFeeItemID As Long, Optional ByVal strNo As String, Optional ByVal lngAdviceID As Long)
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            If mrsInfo.State = 0 Then Exit Sub
            Call LoadApplyData(lngFeeItemID, lngAdviceID, strNo)
        Else
            Call LoadAppliedData
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            Call LoadAuditData(0)
        Else
            Call LoadAuditData(1)
        End If
    End If
End Sub
Private Function zlGetVarBoundSQL(ByVal strVars As String, ByVal lngStep As Long, ByRef strSql As String) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�󶨱�����SQL,��Ҫ����Oracle����
    '���:strVars -���봮(�ö��ŷ���)
    '       lngStep-����(���󶨱����Ӻö࿪ʼ)
    '����:strSQL-���ص�SQL
    '����:���ظ��󶨱���,��Ҫ��10������
    '����:���˺�
    '����:2010-12-27 15:37:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intR As Long, strItems As String, strSubTable As String
    Dim varData As Variant, i As Long, strValues(0 To 10) As String
    strItems = "": strSubTable = ""
    intR = 0:
    varData = Split(strVars, ",")
    For i = 0 To UBound(varData)
        If Len(strItems) > 2000 Then
            If intR <= 10 Then
                strValues(intR) = Mid(strItems, 2)
                strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list([" & intR + lngStep & "]) As ZLTOOLS.t_numlist))"
            Else
                strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list('" & Mid(strItems, 2) & "')  As ZLTOOLS.t_numlist))"
            End If
            strItems = "": intR = intR + 1
        End If
        strItems = strItems & "," & varData(i)
    Next
    
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list([" & intR + lngStep & "]) As ZLTOOLS.t_numlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list('" & Mid(strItems, 2) & "')  As ZLTOOLS.t_numlist))"
        End If
    End If
    If strSubTable <> "" Then strSubTable = Mid(strSubTable, 11)
    strSql = strSubTable: zlGetVarBoundSQL = strValues
End Function
Private Function zlApplyToVerify(ByRef str����ID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-27 14:53:02
    '����:34994
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strMCNO As String, strTable As String
    Dim cllPro As New Collection, i As Long, varValue As Variant, strSql As String
    Dim strIF As String, strNos As String, arrNO As Variant, arrMCRec As Variant, arrMCPar As Variant
    On Error GoTo errHandle
     strIF = " And Instr(','||[1]||',',','||A.��˲���ID||',')>0 and A.��˲���ID=A.���벿��ID and a.���벿��ID=[2] and    A.״̬ = 0"
    '�Ƿ����ǿ�Ƽ���Ȩ��
    If Not (InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0) Then
        If InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 Then
            strIF = strIF & " And ((G.��Ժ���� is NULL And Nvl(G.״̬,0)<>3) Or Nvl(Y.�������,0)<>0)"
        ElseIf InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0 Then
            strIF = strIF & " And ((G.��Ժ���� is NULL And Nvl(G.״̬,0)<>3) Or Nvl(Y.�������,0)=0)"
        Else
            strIF = strIF & " And G.��Ժ���� is NULL And Nvl(G.״̬,0)<>3"
        End If
    End If
    strTable = ""
    varValue = zlGetVarBoundSQL(str����ID, 3, strTable)
   ' strTable = "select * From (" & strTable & ")"
    '��ϸ��¼
    strSql = "" & _
    "   With C1 as  (" & strTable & ")" & _
    "   Select    /*+ RULE */  A.����ID ID,A.��˲���ID,A.�������, To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') ����ʱ��,A.����," & _
    "               B.NO,  B.���, B.��¼����, G.����, A.״̬,  B.����, B.�Ա�,b.����Ա����,B.�Ǽ�ʱ��" & _
    "   From סԺ���ü�¼ B, ������ҳ G, ������� Y, ���˷������� A ,C1" & vbNewLine & _
    "   Where   A.����id = B.ID and A.����ID=C1.ID " & vbNewLine & _
    "               And B.����id = G.����id And B.��ҳid = G.��ҳid   " & _
    "               And B.����ID=Y.����ID(+) And Y.����(+)=1 And Y.����(+)=2 " & strIF
    strSql = "Select * From (" & strSql & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrUnitIDs, mlngDeptID, _
        CStr(varValue(0)), CStr(varValue(1)), CStr(varValue(2)), CStr(varValue(3)), CStr(varValue(4)), CStr(varValue(5)), _
        CStr(varValue(6)), CStr(varValue(7)), CStr(varValue(8)), CStr(varValue(9)), CStr(varValue(10)))
        
    Do While Not rsTemp.EOF
            If zlCheckFeeIsValied(Val(Nvl(rsTemp!ID)), Val(Nvl(rsTemp!��˲���id)), Val(Nvl(rsTemp!�������))) = False Then Exit Function
            ' �������_In ���˷�������.�������%Type := 1 --��ҩƷ��������Ч,ȱʡΪ��ִ�е�ҩƷ������
            gstrSQL = "zl_���˷�������_Audit(" & rsTemp!ID & ",To_Date('" & rsTemp!����ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                UserInfo.���� & "',to_date('" & rsTemp!����ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),1,1," & Val(Nvl(rsTemp!�������)) & ")"
            zlAddArray cllPro, gstrSQL
            '����:49206 --����״̬_In:0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������)
            gstrSQL = "ZL_סԺ���ʼ�¼_Delete('" & rsTemp!NO & "','" & rsTemp!��� & ":" & rsTemp!���� & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & rsTemp!��¼���� & ",1)"
            zlAddArray cllPro, gstrSQL
            
            If Not IsNull(rsTemp!����) And InStr(1, strMCNO, rsTemp!NO) = 0 Then
                    MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val("" & rsTemp!����))
                    MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val("" & rsTemp!����))
                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & Nvl(rsTemp!NO) & "," & Val(Nvl(rsTemp!����)) & _
                            "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
            End If
            
            If InStr("," & strNos, "," & Nvl(rsTemp!NO)) = 0 Then
                strNos = strNos & "," & Nvl(rsTemp!NO) & "|" & Format(rsTemp!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS") & "|" & rsTemp!����Ա����
            End If
        rsTemp.MoveNext
    Loop
    If strNos <> "" Then    '�������ʱ
        arrNO = Split(Mid(strNos, 2), ",")
        For i = 0 To UBound(arrNO)
            If Not BillOperCheck(5, CStr(Split(arrNO(i), "|")(2)), CDate(Split(arrNO(i), "|")(1)), "�������", _
                Split(arrNO(i), "|")(0), , 2, , False, False) Then Exit Function
        Next
    End If
    If cllPro Is Nothing Then
        zlApplyToVerify = True: Exit Function
    End If
    If cllPro.Count = 0 Then
        zlApplyToVerify = True: Exit Function
    End If
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    'ҽ�������������ϴ�������ʱ�ϴ�
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans: Exit Function
                End If
            End If
        Next
    End If
    gcnOracle.CommitTrans
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
    stbThis.Panels(2).Text = "����������˳ɹ�!"
    zlApplyToVerify = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LoadAuditData(ByVal BytType As Byte)
'����:bytType=0-����˼�¼,1-����˼�¼
    Dim strSql As String, strDetail As String, strDosage As String
    Dim rsTmp As ADODB.Recordset
    Dim strIF As String, strFirstCol As String
    Dim DatBegin As Date, DatEnd As Date
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    
    If dtpAuditB.Value <= dtpAuditE.Value Then
        DatBegin = dtpAuditB.Value
        DatEnd = dtpAuditE.Value
    Else
        DatBegin = dtpAuditE.Value
        DatEnd = dtpAuditB.Value
    End If
        
    strIF = " And Instr(','||[3]||',',','||A.��˲���ID||',')>0"
    
    If BytType = 0 Then
        If chkDateAudit.Value = 0 Then
            strIF = strIF & " And A.����ʱ�� Between [1] And [2]"
        End If
        strIF = strIF & " And A.״̬ = 0"
        strFirstCol = "' ' ���, "
    Else
        strIF = strIF & " And A.���ʱ�� Between [1] And [2]"
        strIF = strIF & " And A.״̬ IN(1,2)"
        strFirstCol = "Decode(״̬,1,'��','��') ״̬, "
    End If
    
    If BytType = 0 Then
        '�Ƿ����ǿ�Ƽ���Ȩ��
        If Not (InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0) Then
            If InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 Then
                strIF = strIF & " And ((G.��Ժ���� is NULL And Nvl(G.״̬,0)<>3) Or Nvl(Y.�������,0)<>0)"
            ElseIf InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0 Then
                strIF = strIF & " And ((G.��Ժ���� is NULL And Nvl(G.״̬,0)<>3) Or Nvl(Y.�������,0)=0)"
            Else
                strIF = strIF & " And G.��Ժ���� is NULL And Nvl(G.״̬,0)<>3"
            End If
        End If
    End If
    '����:42827,42837
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    If lng����ID <> 0 Then
        strIF = strIF & " And B.����id+0 = [4] "
    End If
    '����59958,������:��ʾ��������Ϣ,Ӧ���ų�������Һ��ҩ���ĵ�ҩƷ
    '77686,���ϴ�,2014/9/18,�����������
    strDosage = " And Not Exists (Select  RowNum as ��� From ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1 Where B1.����id = B.ID And B1.ID = C1.�շ�id And instr( ',8,9,10,21,24,25,26,',','||B1.����||',')>0) "
    '��ϸ��¼
    strDetail = "Select A.����ID ID, B.���, B.��¼����, A.״̬,  B.����, B.�Ա�, G.����, F.���� ���, D.���� ���˲���, G.��Ժ���� ����, E.���� ��������, A.�շ�ϸĿid," & vbNewLine & _
                "       C.���� As ��Ŀ����, C.���, Nvl(X.סԺ��λ,C.���㵥λ) as ��λ,B.NO, B.����ʱ��, B.�Ǽ�ʱ��, B.����Ա����, A.����/Nvl(X.סԺ��װ,1) ��������,A.���� �ۼ���������,A.����*Nvl(B.ʵ�ս��,0)/B.����/B.���� As ���ʽ�� ," & vbNewLine & _
                " A.������, To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') ����ʱ��, C.����, X.ҩƷ��Դ,B.����ID,B.ִ��״̬,B.ִ�в���ID,A.��˲���ID,A.�������" & vbNewLine & _
                "From סԺ���ü�¼ B, ������ҳ G, ������� Y, ���˷������� A, �շ���ĿĿ¼ C, ҩƷ��� X, ���ű� D, ���ű� E, �շ���Ŀ��� F" & vbNewLine & _
                "Where A.����id = B.ID And B.�շ�ϸĿid = C.ID And B.���˲���id = D.ID And B.��������ID = E.ID " & vbNewLine & _
                "      And B.�շ���� = F.���� And B.����id = G.����id And B.��ҳid = G.��ҳid And A.�շ�ϸĿID=X.ҩƷID(+) And B.����ID=Y.����ID(+) And Y.����(+)=1 And Y.����(+)=2  " & strDosage & strIF
    Set rsTmp = zlDatabase.OpenSQLRecord(strDetail, Me.Caption, DatBegin, DatEnd, mstrUnitIDs, lng����ID)
    If BytType = 0 Then
        Set mrsAudit = rsTmp
    Else
        Set mrsAudited = rsTmp
    End If
     
    strSql = "Select " & strFirstCol & "����, �Ա�, ���˲���, ����, ���, ��Ŀ����,�շ�ϸĿID, ���, ��λ,�������, Sum(��������) ��������, Sum(���ʽ��) ���ʽ��, ������, ����ʱ��, ����, ҩƷ��Դ" & vbNewLine & _
            "From (" & strDetail & ")" & vbNewLine & _
            "Group by �շ�ϸĿid,�������, ״̬, ����, �Ա�, ���˲���, ����, ���, ��Ŀ����, ���, ��λ, ������, ����ʱ��, ����, ҩƷ��Դ" & vbNewLine & _
            "Order by ����ʱ�� Desc,����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, DatBegin, DatEnd, mstrUnitIDs, lng����ID)
    Call ShowMainData(rsTmp)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadAppliedData()
'���ܣ���ȡ����������ʵ�
    Dim strSql As String, strDetail As String
    Dim rsTmp As ADODB.Recordset
    Dim strIF As String, strFirstCol As String, strDosage As String
    Dim DatBegin As Date, DatEnd As Date
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    If Not chkDateAudit.Value = 1 Then
        If dtpAuditB.Value <= dtpAuditE.Value Then
            DatBegin = dtpAuditB.Value
            DatEnd = dtpAuditE.Value
        Else
            DatBegin = dtpAuditE.Value
            DatEnd = dtpAuditB.Value
        End If
        strIF = " And A.����ʱ�� Between [1] And [2]"
    End If
    If mlngDeptID <> 0 Then
        If mbytUseType <> 1 Then
            strIF = strIF & " And Instr(','||[4]||',',','||A.���벿��ID||',')>0"
        Else
            strIF = strIF & " And A.���벿��ID = [3]"
        End If
    End If
    '����:42827,42837
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    If lng����ID <> 0 Or chkDateAudit.Value = 1 Then
        strIF = Replace(strIF, "A.����ʱ��", "A.����ʱ��+0") & "  And B.����id  = [6] "
    End If
    
    '0-ȫ��,1-δ���,2-���ͨ��,3-���δͨ��
    Select Case Val(cboState.ItemData(cboState.ListIndex))
        Case ESTATE.Eȫ��
            strFirstCol = "Decode(״̬,0,' ',1,'��','��') ״̬,"
        Case ESTATE.Eδ���
            strFirstCol = "' ' ״̬,"
            '����:42716
            strIF = strIF & " And A.״̬ = 0 " & IIf(chkOtherOperator.Value = 1 And chkOtherOperator.Visible, "", " And A.������ = [5]")
        Case ESTATE.E���ͨ��
            strFirstCol = "'��' ״̬,"
            strIF = strIF & " And A.״̬ = 1"
        Case ESTATE.E���δͨ��
            strFirstCol = "' ' ״̬,"
            strIF = strIF & " And A.״̬ = 2 And A.������ = [5]"
    End Select
    '����59958,������:��ʾ��������Ϣ,Ӧ���ų�������Һ��ҩ���ĵ�ҩƷ
    '77686,���ϴ�,2014/9/18,�����������
    strDosage = " And Not Exists (Select  RowNum as ��� From ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1 Where B1.����id = B.ID And B1.ID = C1.�շ�id And instr( ',8,9,10,21,24,25,26,',','||B1.����||',')>0)  "
    '��ϸ��¼
    strDetail = "Select A.����ID ID, A.״̬, B.����, B.�Ա�, F.���� ���,B.�շ����, A.�շ�ϸĿid, C.���� As ��Ŀ����, C.���," & vbNewLine & _
            "       Nvl(X.סԺ��λ,C.���㵥λ) as ��λ,B.NO, B.����ʱ��, D.���� ִ�п���,E.���� ��������," & vbNewLine & _
            "       A.����/Nvl(X.סԺ��װ,1) ��������,A.����*nvl(B.��׼����,0) as ���ʽ��, A.������, To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') ����ʱ��,C.����, X.ҩƷ��Դ,B.ҽ�����" & vbNewLine & _
            "From ���˷������� A, סԺ���ü�¼ B, �շ���ĿĿ¼ C, �շ���Ŀ��� F, ҩƷ��� X, ���ű� D, ���ű� E" & vbNewLine & _
            "Where A.����id = B.ID And A.�շ�ϸĿid = C.ID And B.ִ�в���id = D.ID And B.��������id = E.ID And B.�շ���� = F.���� And A.�շ�ϸĿID=X.ҩƷID(+)" & strDosage & strIF
    Set mrsApplied = zlDatabase.OpenSQLRecord(strDetail, Me.Caption, DatBegin, DatEnd, mlngDeptID, mstrUnitIDs, UserInfo.����, lng����ID)
     
    strSql = "Select " & strFirstCol & " ����, �Ա�, ���, ��Ŀ����,�շ�ϸĿID, ���, ��λ, Sum(��������) ��������,sum(���ʽ��) as ���ʽ��, ������, ����ʱ��, ����, ҩƷ��Դ" & vbNewLine & _
            "From (" & strDetail & ")" & vbNewLine & _
            "Group by �շ�ϸĿid, ״̬, ����, �Ա�, ���, ��Ŀ����, ���, ��λ, ������, ����ʱ��, ����, ҩƷ��Դ" & vbNewLine & _
            "Order by ����ʱ�� Desc,����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, DatBegin, DatEnd, mlngDeptID, mstrUnitIDs, UserInfo.����, lng����ID)
    Call ShowMainData(rsTmp)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadApplyData(ByVal lngFeeItemID As Long, Optional ByVal lngAdviceID As Long, _
                        Optional ByVal strNo As String, Optional lngSerial As Long)
    '����:��ȡ�������ʼ�¼
    Dim strSql As String, strSQLDetail As String
    Dim rsTmp As ADODB.Recordset
    Dim strIF As String, blnAppend As Boolean, blnVsfEmpt As Boolean
    Dim DatBegin As Date, DatEnd As Date
    Dim strWhere As String, strWhereExists As String
    Dim strTable As String, bln��Ӥ���� As Boolean
    Dim strDosage As String '��ҩ������ҩ����
    Dim strWhereOthers As String
    Dim lng��ҳID As Long, str�շ���� As String
    Dim intBaby As Integer
    On Error GoTo errHandle
     
    If lngFeeItemID <> 0 Then
        If vsfMain(0).Rows > 1 Then
            blnVsfEmpt = Val(vsfMain(0).RowData(1)) = 0
        Else
            blnVsfEmpt = True
        End If
        
        If Not blnVsfEmpt Then
            If CheckExistFeeItem(lngFeeItemID) Then
                If MsgBox("�����������Ŀ�Ѵ������б���,��Ҫ����б��е�����,ֻ��ʾ����Ŀ��?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnAppend = False
                Else
                    Exit Sub
                End If
            Else
                blnAppend = True
            End If
        End If
    End If
    
    'strIF = " And A.����id" & IIf(chkDate.Value = 0, "+0", "") & " = [1] And A.��¼״̬ > 0"
    
    'ȫ���߲�������:����:29176
    strIF = " And A.����id = [1] And A.��¼״̬ > 0  "
    '����:39373
    '55368
    intBaby = cboBaby.ItemData(cboBaby.ListIndex)
    Select Case intBaby
    Case 0  '����Ӥ����
        strIF = strIF & " And nvl(A.Ӥ����,0)= 0"
    Case 1  '��Ӥ����
    Case Else '��ʾ�ڼ���Ӥ����
        strIF = strIF & " And nvl(A.Ӥ����,0)= [8]"
    End Select

    '����:40304
    lng��ҳID = 0
    If cbo����.ListIndex >= 0 Then
         lng��ҳID = cbo����.ItemData(cbo����.ListIndex)
    End If
    
    str�շ���� = cboKind.GetNodesCheckedDatas
    If str�շ���� = "" And cboKind.GetNodesCheckedDatas(False) = "" Then
        MsgBox "��ѡ��һ���շ����!", vbInformation, gstrSysName
        Exit Sub
    End If
    strIF = strIF & IIf(Replace(str�շ����, ",", "") = "", "", " And Instr('," & str�շ���� & ",',',' || A.�շ���� || ',') > 0")
    
    strIF = strIF & IIf(lng��ҳID = 0, "", " And nvl(A.��ҳID,0)= [7]")
    If mlngDeptID <> 0 Then
        '0-��������,1-ҽ�����ҵ���,2-ҽ��վ����(ֻ������ҩƷ��������˹���)
        If mbytUseType <> 1 Then
            '38463
            If Not mblnOperatorICU Then
                strIF = strIF & " And Instr(','||[6]||',',','||A.���˲���id||',')>0"
                ' ����:36462
                '�����ҽ�����ҿ���, ��ʿ��ҽ��վ(�ٴ�)��������ʾҩƷ������
                '�����ҽ��վ����,Ҳ���ܿ����ٴ����ĵ�
                '����:43940:���ڻ���ҽ��Ҳ���ڿ�������<>���˿��ҵ����,���, _
                '       �����������,ֱ���Կ�������ID�Ƿ�Ϊ�ٴ������ж�, '
                '       �����ò��˿���ID=��������ID���ж��Ƿ�Ϊ�ٴ����ĵ���
                 'exists(select 1 From ��������˵�� where A.��������id=����ID And ��������='�ٴ�')
                strIF = strIF & _
                    " And ( exists(select 1 From ��������˵�� where A.��������id=����ID And ��������='�ٴ�') And  (Instr(',5,6,7,', ',' || A.�շ���� || ',') > 0 Or (A.�շ���� = '4' And Nvl(C.��������, 0) = 1))  " & _
                         "      Or  (Instr(',5,6,7,', ',' || A.�շ���� || ',') = 0 Or A.�շ���� = '4' And Nvl(C.��������, 0) = 0))"
            ElseIf Not mblnPatiDeptICU Then
                '�Ե�ʱ���˿����Ƿ�ΪICU����:42526
                strIF = strIF & "  And (  exists(Select 1 From  ��������˵�� J1  Where A.���˿���ID=J1.����ID And J1.��������='ICU') "
                strIF = strIF & "  or  ( exists(select 1 From ��������˵�� where A.��������id=����ID And ��������='�ٴ�') And " & _
                         "       (Instr(',5,6,7,', ',' || A.�շ���� || ',') > 0 Or (A.�շ���� = '4' And Nvl(C.��������, 0) = 1)) Or " & _
                         "       (Instr(',5,6,7,', ',' || A.�շ���� || ',') = 0 Or A.�շ���� = '4' And Nvl(C.��������, 0) = 0)) )"
            End If
        Else
            strIF = strIF & " And A.��������id+0 = [2]"
        End If
    End If
    If lngFeeItemID <> 0 Then
        strIF = strIF & " And A.�շ�ϸĿID+0 = [3]"
    End If
    If lngAdviceID <> 0 Then
        strIF = strIF & " And A.NO In (Select Distinct a.No From ����ҽ������ A, ����ҽ����¼ B Where a.ҽ��id = b.Id And (b.Id = [10] Or b.���id = [10]))"
    End If
    If strNo <> "" Then
        strIF = strIF & " And A.NO = [11]"
    End If
    If lngSerial <> 0 Then
        strIF = strIF & " And A.��� = [12]"
    End If
    If chkDate.Value = 0 Then
        If dtpApplyB.Value <= dtpApplyE.Value Then
            DatBegin = dtpApplyB.Value
            DatEnd = dtpApplyE.Value
        Else
            DatBegin = dtpApplyE.Value
            DatEnd = dtpApplyB.Value
        End If
        strIF = strIF & " And A.����ʱ��+0 Between [4] And [5]"
    End If
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    blnYP = zlCheckPrivs(mstrPrivsOpt, "ҩƷ��������")
    blnZL = zlCheckPrivs(mstrPrivsOpt, "������������")
    blnWC = zlCheckPrivs(mstrPrivsOpt, "������������")
    
    If blnYP And blnWC And blnZL Then
        'ȫ��,������
    ElseIf blnYP And blnWC And blnZL = False Then
        strIF = strIF & "  And  �շ���� In('4','5','6','7')"
    ElseIf blnYP And blnWC = False And blnZL Then
        strIF = strIF & "  And  �շ���� <>'4'"
    ElseIf blnYP And blnWC = False And blnZL = False Then
        strIF = strIF & "  And  �շ���� In('5','6','7')"
    ElseIf blnYP = False And blnWC And blnZL = False Then
        strIF = strIF & "  And  �շ���� ='4'"
    ElseIf blnYP = False And blnWC And blnZL Then
        strIF = strIF & "  And instr( '5,6,7',  �շ����)=0 "
    ElseIf blnYP = False And blnWC = False And blnZL Then
        strIF = strIF & "  And instr( '4,5,6,7',  �շ����)=0 "
    Else
        MsgBox "ע��:" & vbCrLf & "  �㲻�߱�ҩƷ�����ļ��������������Ȩ��,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    '69899:������,2014-02-09,�����������ҵ�ִ�з���
    strWhereOthers = IIf(chkShowOthers.Value = 0, " And Exists (  Select 1 From ������Ա Where A.ִ�в���ID =����ID And ��ԱID= [9]) ", " ")

    '36391:��1�滻ΪRowNum,����Oracle��ͼ�Զ��ϲ�:42333
    '77686,���ϴ�,2014/9/18,�����������
    strDosage = " And Not Exists (Select  RowNum as ��� From סԺ���ü�¼ J, ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1 Where j.NO = a.NO And a.��¼���� = j.��¼���� and nvl(A.�۸񸸺�, A.���) = Nvl(J.�۸񸸺�, J.���) And B1.����id = j.ID And B1.ID = C1.�շ�id And instr( ',8,9,10,21,24,25,26,',','||B1.����||',')>0)  "
    strWhere = ""
    '����:30523
    '������ǰ��δִ�е�ҩƷ������,ֻ�ܶԲ�����������ҩƷ�����Ľ�������,����ҩ����ҩ�����ĵ�,���ܹ����봦��,��˳��ֳ����ڴ˻����ϵ����̴���©��.
    '��������ȡ���˸�����,���ڵĴ�����ʽ�����ҩ������δִ�е�,������ʱ,��˲���ֻ��Ϊ����(����ڲ������ǰ,��ҩƷ��ҩƷִ��,���ֹ���),ִ���˵�,��Ϊִ�в���.
    If chk��Ŀ(0).Value = 1 And chk��Ŀ(1).Value = 0 Then 'ֻ��ʾ��ִ�е�
        strWhere = "     And Exists (  Select 1 From סԺ���ü�¼ B  Where A.NO = B.NO And A.��¼���� = B.��¼���� And Nvl(A.�۸񸸺�, A.���) = Nvl(B.�۸񸸺�, B.���)  And B.ִ��״̬ <> 0 )" & vbNewLine
    ElseIf chk��Ŀ(0).Value = 0 And chk��Ŀ(1).Value = 1 Then 'ֻ��ʾδִ�е�
        strWhere = "     And Exists (  Select 1 From סԺ���ü�¼ B  Where A.NO = B.NO And A.��¼���� = B.��¼���� And Nvl(A.�۸񸸺�, A.���) = Nvl(B.�۸񸸺�, B.���)  And B.ִ��״̬ = 0 )" & vbNewLine
    ElseIf chk��Ŀ(0).Value = 0 And chk��Ŀ(1).Value = 0 Then 'δѡ��ִ����Ŀ��,ȱʡΪȫѡ
    Else
    End If
    
    'δ���ʵ�(���ʲ����ϵ�δ����)
    strWhereExists = "" & _
    "   And exists( Select 1 From סԺ���ü�¼ A1" & vbNewLine & _
    "                      Where Mod(A1.��¼����, 10) = 2 And A.NO=A1.NO and Nvl(A.�۸񸸺�, A.���)=Nvl(A1.�۸񸸺�, A1.���)" & vbNewLine & _
    "                      Group By A1.NO, Mod(A1.��¼����, 10), Nvl(A1.�۸񸸺�, A1.���)" & vbNewLine & _
    "                       Having Nvl(Sum(A1.���ʽ��),0) = 0) "
    
    strTable = " " & _
    "       Select Max(Decode(Sign(A.ִ��״̬), -1, 0, Decode(A.�۸񸸺�, Null, A.ID, 0))) ID, " & _
    "               A.ִ��״̬,nvl(A.Ӥ����,0) as Ӥ����, A.����ʱ��, Max(A.�Ǽ�ʱ��) �Ǽ�ʱ��, " & vbNewLine & _
    "              Max(Decode(Sign(A.ִ��״̬), -1, Null,A.����Ա����)) ����Ա����, A.NO," & vbNewLine & _
    "              Nvl(A.�۸񸸺�, A.���) As ���, A.ִ�в���id,A.��������id, A.�շ����, A.�շ�ϸĿid, Avg(A.��׼����) ����," & vbNewLine & _
    "              Avg(A.����) ����, Avg(A.����) ����, Sum(A.Ӧ�ս��) Ӧ�ս��, Sum(A.ʵ�ս��) ʵ�ս��, A.����ID, A.ҽ�����  " & vbNewLine & _
    "       From סԺ���ü�¼ A, �������� C " & vbNewLine & _
    "       Where A.�շ�ϸĿid = C.����id(+) And  A.��¼���� = 2  " & strDosage & _
    "                   And Exists (  Select 1 From סԺ���ü�¼ B  Where A.NO = B.NO And A.��¼���� = B.��¼���� And Nvl(A.�۸񸸺�, A.���) = Nvl(B.�۸񸸺�, B.���)  And B.ִ��״̬ <> 0 )" & vbNewLine & _
    "                   And ( A.�շ���� in ('5','6','7') or (A.�շ����='4' and  nvl(C.��������,0) = 1)) " & vbCrLf & strWhereExists & _
                         strIF & vbNewLine & _
    "       Group By A.NO,A.ִ��״̬,nvl(A.Ӥ����,0),  Nvl(A.�۸񸸺�, A.���),A.����ʱ��, A.ִ�в���id,A.��������id, A.�շ����, A.�շ�ϸĿid,A.����ID,  A.ҽ����� "
    
    strTable = "" & _
    " Select Max(ID) ID, NO, ����ʱ��, Max(�Ǽ�ʱ��) �Ǽ�ʱ��, Max(����Ա����) as ����Ա����,max(Ӥ����) as Ӥ����, ���,  " & _
    "           ִ�в���id,��������id, �շ����, �շ�ϸĿid, Avg(����) ����," & vbNewLine & _
    "           Sum(���� * ����)  ����, Sum(Ӧ�ս��) Ӧ�ս��, Sum(ʵ�ս��) ʵ�ս��, Max(����ID) as ����ID, ҽ�����" & vbNewLine & _
    " From (" & strTable & ") " & _
    " Group By NO, ����ʱ��, ���, ִ�в���id,��������id, �շ����, �շ�ϸĿid, ҽ�����" & vbNewLine & _
    " Having Sum(���� * ����) <> 0 "
    '����:38388
    strTable = " With ����  as ( " & strTable & ") "
    strSql = ""
    If chk��Ŀ(0).Value = 1 Or (chk��Ŀ(0).Value = 0 And chk��Ŀ(1).Value = 0) Then
            '��ִ����Ŀ,��Ҫ��������ҩ����
            strSql = strSql & " UNION ALL " & _
            "      Select C1.ID,-1 as ִ��״̬,max(C1.Ӥ����) as Ӥ����,C1.����ʱ��,C1.�Ǽ�ʱ��,C1.����Ա����,C1.NO,C1.���,C1.ִ�в���id,C1.��������id, C1.�շ����, C1.�շ�ϸĿid, max(C1.����) ����," & _
            "                1 as ����, -1* Sum(B.ʵ������)  as ���� ," & _
            "               -1*Sum(C1.Ӧ�ս��)*Round(Sum(Nvl(B.����,1)*B.ʵ������) /  sum(C1.����),5) as Ӧ�ս��," & _
            "               -1*Sum(C1.ʵ�ս��)*Round(Sum(Nvl(B.����,1)*B.ʵ������) / sum(C1.����),5) as ʵ�ս��," & _
            "               C1.����ID, C1.ҽ�����,1 as ��ִ��״̬,1 as ҩƷ����" & _
            "      From  ���� C1,ҩƷ�շ���¼  B " & _
            "      Where C1.ID=B.����ID And MOD(B.��¼״̬,3)=1 And    B.����  In (24,25,26,8,9,10)  And B.����� is NULL " & _
            "      Group By C1.ID,C1.����ʱ��,C1.�Ǽ�ʱ��,C1.����Ա����,C1.NO,C1.���,C1.ִ�в���id,C1.��������id, C1.�շ����, C1.�շ�ϸĿid,C1.����ID, C1.ҽ�����"
     End If
    If chk��Ŀ(1).Value = 1 Or (chk��Ŀ(0).Value = 0 And chk��Ŀ(1).Value = 0) Then
        strSql = strSql & " Union ALL " & _
        "      Select C1.ID,0 as ִ��״̬,max(C1.Ӥ����) as Ӥ���� ,C1.����ʱ��,C1.�Ǽ�ʱ��,C1.����Ա����,C1.NO,C1.���,C1.ִ�в���id,C1.��������id, C1.�շ����, C1.�շ�ϸĿid, max(C1.����) ����," & _
        "               1 as ����,  Sum(Nvl(B.����,1)*B.ʵ������)  as ���� ," & _
        "               Sum(C1.Ӧ�ս��)*Round(Sum(Nvl(B.����,1)*B.ʵ������) /  sum(C1.����),5) as Ӧ�ս��," & _
        "               Sum(C1.ʵ�ս��)*Round(Sum(Nvl(B.����,1)*B.ʵ������) / sum(C1.����),5) as ʵ�ս��," & _
        "               C1.����ID, C1.ҽ�����,0 as ��ִ��״̬,1 as ҩƷ����" & _
        "      From  ���� C1,ҩƷ�շ���¼  B " & _
        "      Where C1.ID=B.����ID And MOD(B.��¼״̬,3)=1  And    B.����  In (24,25,26,8,9,10)  And B.����� is NULL " & _
        "      Group By C1.ID,C1.����ʱ��,C1.�Ǽ�ʱ��,C1.����Ա����,C1.NO,C1.���,C1.ִ�в���id,C1.��������id, C1.�շ����, C1.�շ�ϸĿid,C1.����ID, C1.ҽ�����"
    End If
    
    'δ���������������
    '�˹�ҩ��,��Ϊ�˵�ʱ��ֻ������,���Ը�����׼,��ȡ1
    '�����ҩƷ, ���ܷ�ҩ������ҩ��,����Ҫ��Exists�Ӳ�ѯ�ж�,����ֱ����ִ��״̬<>0
    '31313:Max(����ID):��Ҫ�ǽ���Ƚ��ʺ�,�ٶԼ��ʵ��������ʵ����
    strSQLDetail = "" & _
            " Select Max(ID) ID,��ִ��״̬ as ִ��״̬,max(Ӥ����) as Ӥ����,ҩƷ����, NO, ����ʱ��, Max(�Ǽ�ʱ��) �Ǽ�ʱ��, Max(����Ա����) as ����Ա����, ���, ִ�в���id,��������id, �շ����, �շ�ϸĿid, Avg(����) ����," & vbNewLine & _
            "       Decode(Sign(Min(ִ��״̬)), -1, 1, Sum(����)) ����," & vbNewLine & _
            "       Decode(Sign(Min(ִ��״̬)), -1, Sum(���� * ����), Sum(����)) ����, Sum(Ӧ�ս��) Ӧ�ս��, Sum(ʵ�ս��) ʵ�ս��, Max(����ID) as ����ID, ҽ�����" & vbNewLine & _
            " From (Select Max(Decode(Sign(A.ִ��״̬), -1, 0, Decode(A.�۸񸸺�, Null, A.ID, 0))) ID, A.ִ��״̬ ,max(nvl(A.Ӥ����,0)) as Ӥ����, A.����ʱ��, Max(A.�Ǽ�ʱ��) �Ǽ�ʱ��, " & vbNewLine & _
            "              Decode(Sign(A.ִ��״̬), -1, Null,A.����Ա����) ����Ա����, A.NO," & vbNewLine & _
            "              Nvl(A.�۸񸸺�, A.���) As ���, A.ִ�в���id,A.��������id, A.�շ����, A.�շ�ϸĿid, Avg(A.��׼����) ����," & vbNewLine & _
            "              Avg(A.����) ����, Avg(A.����) ����, Sum(A.Ӧ�ս��) Ӧ�ս��, Sum(A.ʵ�ս��) ʵ�ս��, A.����ID, A.ҽ�����, " & vbNewLine & _
            "              Max(Decode(Sign(A.ִ��״̬), -1, 1, Decode(A.�۸񸸺�, Null, decode(A.ִ��״̬,2,1,1,1,decode(A.��¼״̬,1,0,1)), 1))) ��ִ��״̬," & _
            "             decode(A.�շ����,'5',1,'6',1,'7',1,'4',decode(Max(nvl(C.��������,0)),1,1,0),0) as ҩƷ����  " & _
            "       From סԺ���ü�¼ A, �������� C" & vbNewLine & _
            "       Where A.�շ�ϸĿid = C.����id(+) " & vbNewLine & strDosage & strWhere & _
            "                   And A.��¼���� = 2 " & strWhereExists & strIF & vbNewLine & _
            "       Group By A.NO, A.ִ��״̬,Decode(Sign(A.ִ��״̬), -1, Null,A.����Ա����), Nvl(A.�۸񸸺�, A.���), " & _
            "                       A.����ʱ��, A.ִ�в���id,A.��������id, A.�շ����, A.�շ�ϸĿid, A.����ID, A.ҽ����� " & _
                    strSql & _
            "           )" & vbNewLine & _
            " Group By NO, ����ʱ��, ���, ��ִ��״̬,ִ�в���id,��������id, �շ����,ҩƷ����, �շ�ϸĿid, ҽ�����" & vbNewLine & _
            " Having Sum(���� * ����) <> 0 "
            
       strSQLDetail = strTable & vbCrLf & strSQLDetail
            '30523:����
'            "            And   (A.�շ���� Not In ('4', '5', '6', '7') Or A.�շ���� = '4' And C.�������� = 0 Or" & vbNewLine & _
            "                     (A.�շ���� In ('5', '6', '7') Or A.�շ���� = '4' And C.�������� = 1) And Exists" & vbNewLine & _
            "                        (Select 1" & vbNewLine & _
            "              From סԺ���ü�¼ B" & vbNewLine & _
            "              Where A.NO = B.NO And A.��¼���� = B.��¼���� And Nvl(A.�۸񸸺�, A.���) = Nvl(B.�۸񸸺�, B.���)  " & vbNewLine & _
            "                         And (B.ִ��״̬ <> 0 " & strWhere & ")))" & vbNewLine
    '���������ʵ���ϸ
    'A.����*Nvl(X.סԺ��װ,1) as ����,:����:42823
    strSql = "" & _
            "    Select A.ID,A.ִ��״̬,Ӥ����, A.NO, A.���, A.����ʱ��, A.�Ǽ�ʱ��, A.����Ա����,  " & _
            "           B.���� ִ�п���,A.ִ�в���ID,D.���� ��������,A.��������ID, A.�շ����, A.�շ�ϸĿID," & _
            "           A.����*Nvl(X.סԺ��װ,1) as ����, A.����, A.���� as �ۼ�����,A.����/Nvl(X.סԺ��װ,1) ����," & vbNewLine & _
            "       A.Ӧ�ս��, A.ʵ�ս��, Nvl(C.����, 0)/Nvl(X.סԺ��װ,1) ��������,nvl(C.����,0)*A.���� as ���ʽ��,Nvl(X.סԺ��װ,1) סԺ��װ, A.����ID, A.ҽ�����" & vbNewLine & _
            "   From (" & strSQLDetail & ") A, ���˷������� C,ҩƷ��� X, ���ű� B, ���ű� D" & vbNewLine & _
            "   Where A.ִ�в���id = B.ID And A.��������id = D.ID And A.ID = C.����id(+) and decode(A.ҩƷ����,1,A.ִ��״̬,0)=C.�������(+) And A.�շ�ϸĿID=X.ҩƷID(+) And" & vbNewLine & _
            "           C.״̬(+) = 0" & strWhereOthers
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsInfo!����ID), mlngDeptID, lngFeeItemID, DatBegin, DatEnd, mstrUnitIDs, lng��ҳID, intBaby - 1, UserInfo.ID, lngAdviceID, strNo, lngSerial)
    Call MakeApplyRecordSet(rsTmp, blnAppend) 'Ϊ���޸�����,תΪ���޸ĵļ�¼��
    '��ϸ���շ�ϸĿ����
    strSql = "" & _
            "   Select A.�շ�ϸĿID, C.���� ���, C.���� �շ����, B.���� ��Ŀ����, B.���, Nvl(X.סԺ��λ,B.���㵥λ) as ��λ,B.����, X.ҩƷ��Դ," & vbNewLine & _
            "           Sum(A.���� * A.����/Nvl(X.סԺ��װ,1)) ����, Sum(Nvl(D.����/Nvl(X.סԺ��װ,1), 0)) ��������,sum(Nvl(D.����,0)*nvl(A.����,0)) as ���ʽ�� " & vbNewLine & _
            "   From (" & strSQLDetail & ") A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���˷������� D, ҩƷ��� X" & vbNewLine & _
            "   Where A.�շ�ϸĿID = B.ID And A.�շ���� = C.���� And A.ID = D.����id(+) And D.״̬(+) = 0  and decode(A.ҩƷ����,1,A.ִ��״̬,0)=D.�������(+)  And A.�շ�ϸĿID=X.ҩƷID(+)" & strWhereOthers & vbNewLine & _
            "   Group By A.�շ�ϸĿID,A.�շ����,C.����, C.����, B.����, B.���, Nvl(X.סԺ��λ,B.���㵥λ),B.����, X.ҩƷ��Դ"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsInfo!����ID), mlngDeptID, lngFeeItemID, DatBegin, DatEnd, mstrUnitIDs, lng��ҳID, intBaby - 1, UserInfo.ID, lngAdviceID, strNo, lngSerial)
    Call ShowMainData(rsTmp, blnAppend)
    If rsTmp.RecordCount = 0 And mblnInit = True Then
        MsgBox "�޷��ҵ���������ĵ��ݣ���������������ԡ�", vbInformation, gstrSysName
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CheckExistFeeItem(ByRef lngFeeItemID As Long) As Boolean
    Dim i As Long
    
    For i = 1 To vsfMain(0).Rows - 1
        If lngFeeItemID = Val(vsfMain(0).RowData(i)) Then
            CheckExistFeeItem = True
            Exit For
        End If
    Next
End Function

Private Sub MakeApplyRecordSet(ByRef rsDetail As ADODB.Recordset, ByVal blnAppend As Boolean)
'���ܣ����������ʵļ�¼��ת��Ϊ���޸ĵļ�¼��
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim rsOperator As ADODB.Recordset
    Dim strOperatorIDs As String
    If Not blnAppend Then
        rsTmp.Fields.Append "ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "����", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "ִ��״̬", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "ҽ�����", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "�շ����", adVarChar, 20, adFldIsNullable
        rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
        rsTmp.Fields.Append "���", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
        rsTmp.Fields.Append "�Ǽ�ʱ��", adDBTimeStamp, , adFldIsNullable
        rsTmp.Fields.Append "����Ա����", adVarChar, 100, adFldIsNullable
        rsTmp.Fields.Append "ִ�п���", adVarChar, 100, adFldIsNullable
        rsTmp.Fields.Append "��������", adVarChar, 100, adFldIsNullable
        rsTmp.Fields.Append "��������ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "�ۼ�����", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "Ӧ�ս��", adCurrency, , adFldIsNullable
        rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
        rsTmp.Fields.Append "��������", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "���ʽ��", adDouble, , adFldIsNullable     '����:35595
        rsTmp.Fields.Append "ԭʼ���ʽ��", adDouble, , adFldIsNullable '����:35595
        rsTmp.Fields.Append "ԭʼ��������", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "סԺ��װ", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "Ӥ����", adBigInt, , adFldIsNullable
        
        rsTmp.CursorLocation = adUseClient
        rsTmp.LockType = adLockOptimistic
        rsTmp.CursorType = adOpenStatic
        rsTmp.Open
        
        Set mrsApply = rsTmp
    End If

    With mrsApply
        For i = 1 To rsDetail.RecordCount
            .AddNew
            !���� = 0
            If mblnOperatorNurse Then
                '60679
                '����ǻ�ʿ,������������������,�Ȱ�����Ա�������ҿ������з���
                If InStr(1, mstrOperatorDeptIDs, "," & Val(Nvl(rsDetail!��������ID)) & ",") > 0 Then
                    !���� = 1
                End If
            End If
            !ID = rsDetail!ID
            !����id = rsDetail!����id
            !ִ��״̬ = Val(Nvl(rsDetail!ִ��״̬))
            !ҽ����� = rsDetail!ҽ�����
            !�շ���� = rsDetail!�շ����
            !�շ�ϸĿID = rsDetail!�շ�ϸĿID
            !NO = rsDetail!NO
            !��� = rsDetail!���
            !����ʱ�� = rsDetail!����ʱ��
            !�Ǽ�ʱ�� = rsDetail!�Ǽ�ʱ��
            !����Ա���� = rsDetail!����Ա����
            !ִ�п��� = rsDetail!ִ�п���
            !�������� = rsDetail!��������
            !ִ�в���ID = rsDetail!ִ�в���ID
            !��������ID = rsDetail!��������ID
            !���� = rsDetail!����
            !���� = rsDetail!����
            !�ۼ����� = rsDetail!�ۼ�����
            !���� = rsDetail!����
            !Ӧ�ս�� = rsDetail!Ӧ�ս��
            !ʵ�ս�� = rsDetail!ʵ�ս��
            !�������� = rsDetail!��������
            !���ʽ�� = rsDetail!���ʽ�� '����:35595
            !ԭʼ���ʽ�� = rsDetail!���ʽ�� '����:35595
            !ԭʼ�������� = rsDetail!��������
            !סԺ��װ = rsDetail!סԺ��װ
            !Ӥ���� = rsDetail!Ӥ���� '39374
            .Update
            rsDetail.MoveNext
        Next
        If .RecordCount > 0 Then .MoveFirst
    End With
End Sub
Private Sub ShowMainData(ByRef rsTmp As ADODB.Recordset, Optional ByVal blnAppend As Boolean)
'����:blnAppend=True-׷��,False-���¼���
    Dim i As Long, j As Long, lngInitRows As Long
    Dim intColModify As Integer, intState As Integer
    Dim vsfCurrent As VSFlexGrid
    
    If tbsType.SelectedItem.Key = "T1" Then
        Set vsfCurrent = vsfMain(0)
        If mbytFun = E���� Then
            cmdOKApply.Enabled = rsTmp.RecordCount > 0
        Else
            cmdOKAudit.Enabled = rsTmp.RecordCount > 0
        End If
    Else
        Set vsfCurrent = vsfMain(1)
        If mbytFun = E���� Then
            intState = Val(cboState.ItemData(cboState.ListIndex))
            cmdCancelApply.Enabled = rsTmp.RecordCount > 0
        Else
            cmdCancelRefuse.Enabled = False
            cmdOKAudit.Enabled = False
        End If
    End If
    
    If blnAppend Then   'And mbytFun = E���� And tbsType.SelectedItem.Key = "T1"
        lngInitRows = vsfCurrent.Rows
        If vsfCurrent.Rows = 2 Then
            If Val(vsfCurrent.RowData(1)) = 0 Then lngInitRows = 1
        End If
    Else
        Call InitMainHead(False, IIf(tbsType.SelectedItem.Key = "T1", 1, 2))
        lngInitRows = 1
    End If
    
    
    With vsfCurrent
        If rsTmp.RecordCount <> 0 Then
            .Redraw = flexRDNone
            .Rows = rsTmp.RecordCount + lngInitRows
            For i = lngInitRows To .Rows - 1
                If mbytFun = E���� Then
                    If tbsType.SelectedItem.Key = "T1" Then
                        .TextMatrix(i, ColApply("���")) = rsTmp!���
                        .TextMatrix(i, ColApply("��Ŀ����")) = rsTmp!��Ŀ����
                        .TextMatrix(i, ColApply("���")) = "" & rsTmp!���
                        .TextMatrix(i, ColApply("��λ")) = "" & rsTmp!��λ
                        .TextMatrix(i, ColApply("����")) = "" & rsTmp!����
                        '.TextMatrix(i, ColApply("Ӥ����")) = IIf(Val(Nvl(rsTmp!Ӥ����)) <> 0, "��", "")
                        .TextMatrix(i, ColApply("ҩƷ��Դ")) = "" & rsTmp!ҩƷ��Դ
                        .TextMatrix(i, ColApply("����")) = FormatEx(rsTmp!����, 5)
                        .TextMatrix(i, ColApply("��������")) = FormatEx(rsTmp!��������, 5)
                        .TextMatrix(i, ColApply("���ʽ��")) = FormatEx(rsTmp!���ʽ��, 5)
                        .TextMatrix(i, ColApply("ԭʼ��������")) = FormatEx(rsTmp!��������, 5)
                        .TextMatrix(i, ColApply("ԭʼ���ʽ��")) = FormatEx(rsTmp!���ʽ��, 5)
                        .RowData(i) = Val(rsTmp!�շ�ϸĿID)
                                    
                        '���ÿ��޸��е���ɫ
                        mbonNotEnter = True
                        .Row = i
                        .Col = ColApply("��������")
                        .CellBackColor = &HE7CFBA    '��ɫ
                        mbonNotEnter = False
                    Else
                        .TextMatrix(i, ColApplied("ѡ��")) = rsTmp!״̬
                        .TextMatrix(i, ColApplied("����")) = rsTmp!����
                        .TextMatrix(i, ColApplied("�Ա�")) = "" & rsTmp!�Ա�
                        .TextMatrix(i, ColApplied("���")) = rsTmp!���
                        .TextMatrix(i, ColApplied("��Ŀ����")) = rsTmp!��Ŀ����
                        .TextMatrix(i, ColApplied("���")) = "" & rsTmp!���
                        .TextMatrix(i, ColApplied("��λ")) = "" & rsTmp!��λ
                        .TextMatrix(i, ColApplied("����")) = "" & rsTmp!����
                        .TextMatrix(i, ColApplied("ҩƷ��Դ")) = "" & rsTmp!ҩƷ��Դ
                        .TextMatrix(i, ColApplied("��������")) = FormatEx(rsTmp!��������, 5)
                        .TextMatrix(i, ColApplied("���ʽ��")) = FormatEx(rsTmp!���ʽ��, 5)
                        .TextMatrix(i, ColApplied("������")) = rsTmp!������
                        .TextMatrix(i, ColApplied("����ʱ��")) = rsTmp!����ʱ��
                        .RowData(i) = Val(rsTmp!�շ�ϸĿID)
                        
                        mbonNotEnter = True
                        .Row = i
                        If intState = ESTATE.Eδ��� Then
                            .Col = ColApplied("ѡ��")
                            .CellBackColor = &HE7CFBA    '��ɫ
                        ElseIf intState = ESTATE.Eȫ�� Then
                            For j = 0 To .Cols - 1
                                .Col = j
                                If rsTmp!״̬ = "��" Then
                                    .CellForeColor = &HC00000
                                ElseIf rsTmp!״̬ = "��" Then
                                    .CellForeColor = &HC0&
                                End If
                            Next
                        ElseIf intState = ESTATE.E���ͨ�� Then
                            For j = 0 To .Cols - 1
                                .Col = j
                                .CellForeColor = &HC00000
                            Next
                        ElseIf intState = ESTATE.E���δͨ�� Then
                            For j = 0 To .Cols - 1
                                .Col = j
                                .CellForeColor = &HC0&
                            Next
                        End If
                        mbonNotEnter = False
                    End If
                Else
                    If tbsType.SelectedItem.Key = "T1" Then
                        .TextMatrix(i, ColAudit("���")) = rsTmp!���
                        .Cell(flexcpData, i, ColAudit("���")) = Val(Nvl(rsTmp!�������))
                        
                        .TextMatrix(i, ColAudit("����")) = rsTmp!����
                        .TextMatrix(i, ColAudit("�Ա�")) = "" & rsTmp!�Ա�
                        .TextMatrix(i, ColAudit("���˲���")) = "" & rsTmp!���˲���
                        .TextMatrix(i, ColAudit("����")) = "" & rsTmp!����
                        .TextMatrix(i, ColAudit("���")) = rsTmp!���
                        .TextMatrix(i, ColAudit("��Ŀ����")) = rsTmp!��Ŀ����
                        .TextMatrix(i, ColAudit("���")) = "" & rsTmp!���
                        .TextMatrix(i, ColAudit("����")) = "" & rsTmp!����
                        .TextMatrix(i, ColAudit("ҩƷ��Դ")) = "" & rsTmp!ҩƷ��Դ
                        .TextMatrix(i, ColAudit("��λ")) = "" & rsTmp!��λ
                        .TextMatrix(i, ColAudit("��������")) = FormatEx(rsTmp!��������, 5)
                        .TextMatrix(i, ColAudit("���ʽ��")) = FormatEx(rsTmp!���ʽ��, 5)
                        .TextMatrix(i, ColAudit("������")) = rsTmp!������
                        .TextMatrix(i, ColAudit("����ʱ��")) = rsTmp!����ʱ��
                        .RowData(i) = Val(rsTmp!�շ�ϸĿID)
                        
                        mbonNotEnter = True
                        .Row = i
                        .Col = ColAudit("���")
                        .CellBackColor = &HE7CFBA    '��ɫ
                        mbonNotEnter = False
                    Else
                        .Cell(flexcpData, i, ColAudited("״̬")) = Val(Nvl(rsTmp!�������))
                        .TextMatrix(i, ColAudited("״̬")) = rsTmp!״̬
                        .TextMatrix(i, ColAudited("����")) = rsTmp!����
                        .TextMatrix(i, ColAudited("�Ա�")) = "" & rsTmp!�Ա�
                        .TextMatrix(i, ColAudited("���˲���")) = "" & rsTmp!���˲���
                        .TextMatrix(i, ColAudited("����")) = "" & rsTmp!����
                        .TextMatrix(i, ColAudited("���")) = rsTmp!���
                        .TextMatrix(i, ColAudited("��Ŀ����")) = rsTmp!��Ŀ����
                        .TextMatrix(i, ColAudited("���")) = "" & rsTmp!���
                        .TextMatrix(i, ColAudited("����")) = "" & rsTmp!����
                        .TextMatrix(i, ColAudited("ҩƷ��Դ")) = "" & rsTmp!ҩƷ��Դ
                        .TextMatrix(i, ColAudited("��λ")) = "" & rsTmp!��λ
                        .TextMatrix(i, ColAudited("��������")) = FormatEx(rsTmp!��������, 5)
                        .TextMatrix(i, ColAudited("������")) = rsTmp!������
                        .TextMatrix(i, ColAudited("����ʱ��")) = rsTmp!����ʱ��
                        .RowData(i) = Val(rsTmp!�շ�ϸĿID)
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            .Row = 1: .Col = 0
            If mbytFun = E���� Then
                If tbsType.SelectedItem.Key = "T1" Then
                    .Col = ColApply("��������") '�����¼�AfterRowColChange
                End If
            Else
                If tbsType.SelectedItem.Key = "T1" Then
                    .Row = 0: .Col = ColAudit("���")
                    .CellBackColor = &HE7CFBA    '��ɫ
                    .Row = 1
                End If
            End If
            
            .Redraw = flexRDDirect
        End If
        Call ShowDetail(.RowData(.Row))
    End With
End Sub


Private Sub ShowDetail(ByVal lngFeeItem As Long)
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            Set rsTmp = mrsApply
            rsTmp.Filter = "�շ�ϸĿID=" & lngFeeItem   'ע��,����ı�ԭ��¼����Filter
        Else
            Set rsTmp = mrsApplied
            With vsfMain(1)
                rsTmp.Filter = "�շ�ϸĿID=" & lngFeeItem & " And ������='" & .TextMatrix(.Row, ColApplied("������")) & _
                            "' And ����ʱ��='" & .TextMatrix(.Row, ColApplied("����ʱ��")) & "'"
            End With
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            Set rsTmp = mrsAudit
            With vsfMain(0)
                rsTmp.Filter = "�շ�ϸĿID=" & lngFeeItem & " And �������=" & Val(.Cell(flexcpData, .Row, ColAudit("���"))) & " And ������='" & .TextMatrix(.Row, ColAudit("������")) & _
                            "' And ����ʱ��='" & .TextMatrix(.Row, ColAudit("����ʱ��")) & "'"
            End With
        Else
            Set rsTmp = mrsAudited
            With vsfMain(1)
                rsTmp.Filter = "�շ�ϸĿID=" & lngFeeItem & " And ������='" & .TextMatrix(.Row, ColAudited("������")) & _
                            "' And ����ʱ��='" & .TextMatrix(.Row, ColAudited("����ʱ��")) & "'"
            End With
        End If
    End If
    
    Call InitDetailHead(True)   '����ʾ���в�ͬ,Ҫ�������
       
    If rsTmp.State = 0 Then Exit Sub
    If rsTmp.RecordCount = 0 Then Exit Sub
    rsTmp.Sort = IIf(tbsType.SelectedItem.Key = "T1", "ִ��״̬,", "") & "����ʱ�� Desc"
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    With vsfDetail
        .Redraw = flexRDNone
        .Rows = rsTmp.RecordCount + 1
        mblnUnChange = True
        For i = 1 To .Rows - 1
            If mbytFun = E���� Then
                If tbsType.SelectedItem.Key = "T1" Then
                    If InStr(1, "5,6,7", Nvl(rsTmp!�շ����)) > 0 Then
                        .TextMatrix(i, .ColIndex("ִ��״̬")) = IIf(Val(Nvl(rsTmp!ִ��״̬)) = 0, "δ��ҩ", "�ѷ�ҩ")
                    ElseIf Nvl(rsTmp!�շ����) = "4" Then
                        .TextMatrix(i, .ColIndex("ִ��״̬")) = IIf(Val(Nvl(rsTmp!ִ��״̬)) = 0, "δ����", "�ѷ���")
                    Else
                         .TextMatrix(i, .ColIndex("ִ��״̬")) = IIf(Val(Nvl(rsTmp!ִ��״̬)) = 0, "δִ��", "��ִ��")
                    End If
                    
                    .Cell(flexcpData, i, .ColIndex("ִ��״̬")) = Nvl(rsTmp!ִ��״̬)
                    .TextMatrix(i, .ColIndex("NO")) = rsTmp!NO
                    .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:MM:SS")
                    .TextMatrix(i, .ColIndex("Ӥ����")) = IIf(Val(Nvl(rsTmp!Ӥ����)) >= 1, "��", "")
                    .TextMatrix(i, .ColIndex("ִ�п���")) = rsTmp!ִ�п���
                    .TextMatrix(i, .ColIndex("��������")) = rsTmp!��������
                    .TextMatrix(i, .ColIndex("����")) = Format(rsTmp!����, "######" & gstrFeePrecisionFmt)
                    .TextMatrix(i, .ColIndex("����")) = rsTmp!����
                    .TextMatrix(i, .ColIndex("����")) = FormatEx(rsTmp!����, 5)
                    .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(rsTmp!Ӧ�ս��, "#######" & gstrDec)
                    .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(rsTmp!ʵ�ս��, "#######" & gstrDec)
                    .TextMatrix(i, .ColIndex("��������")) = FormatEx(rsTmp!��������, 5)
                    .TextMatrix(i, .ColIndex("���ʽ��")) = FormatEx(rsTmp!���ʽ��, 5)
                    .TextMatrix(i, .ColIndex("ԭʼ��������")) = FormatEx(rsTmp!��������, 5)
                    .TextMatrix(i, .ColIndex("ԭʼ���ʽ��")) = FormatEx(rsTmp!���ʽ��, 5)
                    .RowData(i) = Val(rsTmp!ID)
                    .Cell(flexcpBackColor, i, .ColIndex("��������")) = &HE7CFBA    '��ɫ
                    .Cell(flexcpBackColor, i, .ColIndex("ִ��״̬")) = Me.BackColor     '��ɫ
                    
                Else
                    .TextMatrix(i, .ColIndex("NO")) = rsTmp!NO
                    .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:MM:SS")
                    .TextMatrix(i, .ColIndex("ִ�п���")) = rsTmp!ִ�п���
                    .TextMatrix(i, .ColIndex("��������")) = rsTmp!��������
                    .TextMatrix(i, .ColIndex("��������")) = FormatEx(rsTmp!��������, 5)
                    .TextMatrix(i, .ColIndex("���ʽ��")) = FormatEx(rsTmp!���ʽ��, 5)
                    .RowData(i) = Val(rsTmp!ID)
                End If
            Else
                .TextMatrix(i, .ColIndex("NO")) = rsTmp!NO
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:MM:SS")
                .TextMatrix(i, .ColIndex("��������")) = rsTmp!��������
                .TextMatrix(i, .ColIndex("��������")) = FormatEx(rsTmp!��������, 5)
                .RowData(i) = Val(rsTmp!ID)
            End If
            rsTmp.MoveNext
        Next
        mblnUnChange = False
        .Row = 0: .Col = 0
        .Row = 1: .Col = 0
        If mbytFun = E���� And tbsType.SelectedItem.Key = "T1" Then
            .Col = .ColIndex("��������") '�����¼�AfterRowColChange
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub vsfMain_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim dblTotalNum As Double, i As Long, lngFeeItem As Long, j As Long, blnDo As Boolean
    Dim dbl���ʽ�� As Double
    dblTotalNum = Val(vsfMain(Index).EditText)
    lngFeeItem = Val(vsfMain(Index).RowData(Row))
    
    '������ȳ�������ϸ
    With mrsApply
        .Filter = "�շ�ϸĿID=" & lngFeeItem
        dbl���ʽ�� = 0
        If .RecordCount = 0 Then
            MsgBox "�����쳣,δ���޸���ϸ��¼������!", vbInformation, gstrSysName
            Exit Sub
        End If
        .Sort = "ִ��״̬,���� Desc,����ʱ�� Desc"
        For i = 1 To .RecordCount
            If dblTotalNum = 0 Then
                !�������� = 0
                !���ʽ�� = 0
                .Update
            Else
                If Not MCPAR.���ֳ�����ϸ And Not IsNull(mrsInfo!����) And dblTotalNum < !���� * !���� Then
                    If Val(vsfMain(Index).EditText) = dblTotalNum Then
                        MsgBox "��������ҽ�����˽��в��ֳ�����ϸ", vbInformation, gstrSysName
                        vsfMain(Index).TextMatrix(Row, Col) = 0
                        dblTotalNum = 0 'Ҫ����ѭ��,�������еĳ�����ϸ��Ϊ0
                    Else
                        MsgBox "��������ҽ�����˽��в��ֳ�����ϸ,����[" & !NO & "]���ܳ���.", vbInformation, gstrSysName
                        '��ǰ���ݲ��ܳ���,������ĵ��ݿ��ܿ�����ȫ����.
                    End If
                    !�������� = 0
                    !���ʽ�� = 0
                    .Update
                Else
                    blnDo = True
                    If Not IsNull(!����id) Then
                        If CheckBalance(!NO, !���) Then    'Ŀǰδ���ʵ�û����ȡ����,����ĳ�����ʱû��ʹ��
                            If Not IsNull(mrsInfo!����) Then
                                If Not MCPAR.�����ѽ��ʵ��� Then blnDo = False
                            Else
                                Select Case gbytBillOpt
                                    Case 1
                                        If MsgBox("����[" & !NO & "]�еĵ�ǰ������Ŀ�Ѿ�����,ȷ��Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnDo = False
                                    Case 2
                                        MsgBox "����[" & !NO & "]�еĵ�ǰ������Ŀ�Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                                        blnDo = False
                                End Select
                            End If
                        End If
                    End If
                    '�����Һ��ҩ�����Ƿ�����δ��ҩ����
                    '����:?????
                    If InStr(1, "4,5,6,7", Nvl(!�շ����)) > 0 And Val(Nvl(!ҽ�����)) <> 0 Then
                        If Val(Nvl(!ִ��״̬)) = 0 Then  'ֻ��δִ�в��ֲŻ���ڼ��
                            If ҩƷ������ҩ����(Val(Nvl(mrsApply!ҽ�����))) Then
                                MsgBox "����[" & !NO & "]�еĵ�ǰ������Ŀ����Һ��ҩ�����Ѿ�ʹ���˸�ҩƷ������,��������", vbExclamation, gstrSysName
                               blnDo = False
                            End If
                        End If
                    End If
                    If blnDo Then
                        !�������� = IIf(dblTotalNum <= !���� * !����, dblTotalNum, !���� * !����)
                        !���ʽ�� = Nvl(!��������, 0) * Nvl(!����, 0)
                        .Update
                        dblTotalNum = dblTotalNum - !��������
                        dbl���ʽ�� = dbl���ʽ�� + Nvl(!��������, 0) * Nvl(!����, 0)
                    Else
                        !�������� = 0
                        !���ʽ�� = 0
                        .Update
                    End If
                End If
            End If
            .MoveNext
        Next
        
        If dblTotalNum <> 0 Then
            vsfMain(Index).TextMatrix(Row, Col) = Val(vsfMain(Index).EditText) - dblTotalNum
        End If
        vsfMain(Index).TextMatrix(Row, ColApply("���ʽ��")) = FormatEx(dbl���ʽ��, 5)
    End With
    Call ShowDetail(lngFeeItem)
    Call ShowSumMoney
End Sub

Private Function CheckBalance(ByVal strNo As String, ByVal lngRow As Long) As Boolean
'����:����н���ID��ĳ��������ϸ,�Ƿ��ѽ���(��������Ҫ����û�н���)
'����:True:�ѽ���,False:δ����
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1" & vbNewLine & _
            "From סԺ���ü�¼ A" & vbNewLine & _
            "Where Mod(A.��¼����, 10) = 2 And NO = [1] And Nvl(A.�۸񸸺�, A.���) = [2]" & vbNewLine & _
            "Group By A.NO, Mod(A.��¼����, 10), Nvl(A.�۸񸸺�, A.���)" & vbNewLine & _
            "Having Nvl(Sum(���ʽ��),0) = 0"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, lngRow)
    
    CheckBalance = rsTmp.RecordCount = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfMain_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngFeeItem As Long
    
    If mbonNotEnter Or NewRow = 0 Then Exit Sub
    
   With vsfMain(Index)
        If OldRow <> NewRow Then
            lngFeeItem = Val(.RowData(NewRow))
            If lngFeeItem = 0 Then Exit Sub '�쳣
            Call ShowDetail(lngFeeItem)
        End If
            
        If OldCol <> NewCol Then
            If mbytFun = E���� And tbsType.SelectedItem.Key = "T1" Then
                If NewCol = ColApply("��������") And Val(.RowData(NewRow)) <> 0 Then
                    .Editable = flexEDKbdMouse
                Else
                    .Editable = flexEDNone
                End If
            End If
        End If
        If mbytFun = E��� And tbsType.SelectedItem.Key = "T2" Then
            If .TextMatrix(NewRow, ColAudited("״̬")) = "��" Then
                cmdOKAudit.Enabled = True
                cmdCancelRefuse.Enabled = True
            Else
                cmdOKAudit.Enabled = False
                cmdCancelRefuse.Enabled = False
            End If
        End If
    End With
End Sub

Private Function SaveRefuse(blnCancel As Boolean) As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'����:ִ��ȡ���ܾ�������˾ܾ�����
'���:blnCancel-True��ʾȡ���ܾ� False��ʾ��˾ܾ�
'����:�ɹ�����True,ʧ�ܷ���False
'����:������
'����:2014-4-15
'��ע:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strDate As String, strNos As String
    Dim intRow As Integer
    Dim arrSQL As Variant, i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant, arrNO As Variant
    arrSQL = Array()
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    '84026:���ϴ�,2015/4/20���ݴ�����
    On Error GoTo ErrHand
    With mrsAudited
        intRow = vsfMain(1).Row
        .Filter = "�շ�ϸĿID=" & vsfMain(1).RowData(intRow) & " And �������=" & Val(vsfMain(1).Cell(flexcpData, intRow, ColAudited("״̬"))) & " And ������='" & vsfMain(1).TextMatrix(intRow, ColAudited("������")) & _
                "' And ����ʱ��='" & vsfMain(1).TextMatrix(intRow, ColAudited("����ʱ��")) & "'"
        If blnCancel = True Then
            If .RecordCount <> 0 Then
                strSql = "zl_���˷�������_Cancel(" & !ID & ",To_Date('" & !����ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                    UserInfo.���� & "'," & strDate & "," & IIf(blnCancel, "1", "0") & ",1," & Val(vsfMain(1).Cell(flexcpData, intRow, ColAudited("״̬"))) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Else
                SaveRefuse = False
                Exit Function
            End If
        Else
            If .RecordCount <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_���˷�������_Cancel(" & !ID & ",To_Date('" & !����ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                    UserInfo.���� & "'," & strDate & "," & IIf(blnCancel, "1", "0") & ",1," & Val(vsfMain(1).Cell(flexcpData, intRow, ColAudited("״̬"))) & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_סԺ���ʼ�¼_Delete('" & !NO & "','" & !��� & ":" & !�ۼ��������� & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & !��¼���� & ",1)"
                
                If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                    MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                    MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                            "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                End If
                
                If InStr("," & strNos, "," & !NO) = 0 Then
                    strNos = strNos & "," & !NO & "|" & Format(!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS") & "|" & !����Ա����
                End If
                
                If strNos <> "" Then    '������������ʱ
                    arrNO = Split(Mid(strNos, 2), ",")
                    For i = 0 To UBound(arrNO)
                        If Not BillOperCheck(5, CStr(Split(arrNO(i), "|")(2)), CDate(Split(arrNO(i), "|")(1)), IIf(mbytFun = E����, "��������", "�������"), _
                            Split(arrNO(i), "|")(0), , 2, , False, False) Then Exit Function
                    Next
                End If
                
                 On Error GoTo errH
                 Screen.MousePointer = 11
                 gcnOracle.BeginTrans
                     For i = 0 To UBound(arrSQL)
                         Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
                     Next
                     
                     'ҽ�������������ϴ�������ʱ�ϴ�
                     If strMCNO <> "" Then
                         arrMCRec = Split(strMCNO, "|")
                         For i = 0 To UBound(arrMCRec)
                             arrMCPar = Split(arrMCRec(i), ",")
                             If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                                 If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                     gcnOracle.RollbackTrans: Exit Function
                                 End If
                             End If
                         Next
                     End If
                 gcnOracle.CommitTrans
                 
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
                 Screen.MousePointer = 0
            Else
                SaveRefuse = False
                Exit Function
            End If
        End If
    End With
    Call cmdRefresh_Click
    SaveRefuse = True
    Exit Function
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ҩƷ������ҩ����(ByVal lngҽ��ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ҽ������ҩƷ�Ƿ��Ѿ�����������ʹ����
    '���أ����ڷ���True,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-29 14:55:19
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 1 " & _
             "   From ����ҽ����¼ A, ����ҽ������ B, ��Һ��ҩ��¼ D " & _
             "   Where A.���id = B.ҽ��id And B.ҽ��id = D.ҽ��id And B.���ͺ� = D.���ͺ� And A.ID = [1] And Rownum =1"
     Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҽ��ID)
    ҩƷ������ҩ���� = rsTemp.RecordCount <> 0
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub vsfMain_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    For i = 1 To vsfMain(Index).Rows - 1
        If mlngPreFeeItemID = vsfMain(Index).RowData(i) Then vsfMain(Index).Row = i
    Next
End Sub

Private Sub vsfMain_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mlngPreFeeItemID = vsfMain(Index).RowData(vsfMain(Index).Row)
End Sub

Private Sub vsfMain_DblClick(Index As Integer)
    Dim i As Long, strResult As String, intState As Integer
    
    If mbytFun = E���� And tbsType.SelectedItem.Key = "T2" Then
        With vsfMain(Index)
            If .Col = 0 Then
                intState = Val(cboState.ItemData(cboState.ListIndex))
                If intState = ESTATE.Eδ��� Then
                    If .MouseRow = 0 Then
                        If .ColData(ColApplied("ѡ��")) = "" Then
                            .ColData(ColApplied("ѡ��")) = "��"
                        Else
                            .ColData(ColApplied("ѡ��")) = ""
                        End If
                        strResult = .ColData(ColApplied("ѡ��"))
                        For i = 1 To .Rows - 1
                            .TextMatrix(i, ColApplied("ѡ��")) = strResult
                        Next
                    Else
                        If .TextMatrix(.Row, ColApplied("ѡ��")) = "��" Then
                            .TextMatrix(.Row, ColApplied("ѡ��")) = ""
                        Else
                            .TextMatrix(.Row, ColApplied("ѡ��")) = "��"
                        End If
                    End If
                End If
            End If
        End With
        
    ElseIf mbytFun = E��� And tbsType.SelectedItem.Key = "T1" Then
        With vsfMain(Index)
            If .Col = 0 Then
                If .MouseRow = 0 Then
                    If .ColData(ColAudit("���")) = "" Then
                        .ColData(ColAudit("���")) = "��"
                    Else
                        .ColData(ColAudit("���")) = ""
                    End If
                    strResult = .ColData(ColAudit("���"))
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, ColAudit("���")) = strResult
                        If strResult = "��" Then
                            If Not CheckCanAudit(.RowData(i), .TextMatrix(i, ColAudited("����")), .TextMatrix(i, ColAudited("������")), .TextMatrix(i, ColAudited("����ʱ��"))) Then .TextMatrix(i, ColAudit("���")) = ""
                        End If
                    Next
                Else
                    Select Case Trim(.TextMatrix(.Row, ColAudit("���")))
                        Case "��"
                            .TextMatrix(.Row, ColAudit("���")) = "��"
                        Case "��"
                            .TextMatrix(.Row, ColAudit("���")) = ""
                        Case ""
                            If CheckCanAudit(.RowData(.Row), .TextMatrix(.Row, ColAudited("����")), .TextMatrix(.Row, ColAudited("������")), .TextMatrix(.Row, ColAudited("����ʱ��"))) Then
                                .TextMatrix(.Row, ColAudit("���")) = "��"
                            Else
                                .TextMatrix(.Row, ColAudit("���")) = ""
                            End If
                    End Select
                End If
            End If
        End With
    End If
End Sub

Private Function CheckCanAudit(ByVal lngFeeItemID As Long, ByVal strPatient As String, ByVal strOperater As String, ByVal strDate As String) As Boolean
'����:������˵ķ�����Ŀ�ĵ�����ϸ���Ƿ��ѽ���,�ѽ��ʵĲ��������(����)
    Dim i As Long
    
    '����:29613
    If mrsAudit Is Nothing Then Exit Function
    If mrsAudit.State <> 1 Then Exit Function
    
    CheckCanAudit = True
    With mrsAudit
        .Filter = "�շ�ϸĿid=" & lngFeeItemID & " And ����='" & strPatient & "' And ������='" & strOperater & "' And ����ʱ��='" & strDate & "'"
        For i = 1 To .RecordCount
            If Not IsNull(!����id) Then
                If CheckBalance(!NO, !���) Then
                    If Not IsNull(!����) Then
                        If Not gclsInsure.GetCapability(support���������ѽ��ʵļ��ʵ���, , Val(!����)) Then
                            MsgBox "����������ҽ�������ѽ��ʵĵ���[" & !NO & "]."
                            CheckCanAudit = False
                            Exit For
                        End If
                    Else
                        Select Case gbytBillOpt
                            Case 1
                                If MsgBox("����[" & !NO & "]�еĵ�ǰ������Ŀ�Ѿ�����,ȷ��Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then CheckCanAudit = False: Exit For
                            Case 2
                                MsgBox "����[" & !NO & "]�еĵ�ǰ������Ŀ�Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                                CheckCanAudit = False
                                Exit For
                        End Select
                    End If
                End If
            End If
            .MoveNext
        Next
    End With
End Function

Private Sub vsfMain_EnterCell(Index As Integer)
    With vsfMain(Index)
        .BackColorSel = .CellBackColor
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub vsfMain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsfMain(Index)
            If Val(.RowData(.Row)) = 0 Or Not (mbytFun = E���� And tbsType.SelectedItem.Key = "T1") Then Exit Sub
                        
            If .Col = ColApply("��������") Then
                .TextMatrix(.Row, .Col) = 0
            End If
        End With
    End If
End Sub

Private Sub vsfMain_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       With vsfMain(Index)
            KeyAscii = 0
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If cmdOKApply.Visible And cmdOKApply.Enabled Then cmdOKApply.SetFocus
            End If
       End With
    End If
End Sub


Private Sub vsfMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsfMain(Index)
        If mbytFun = E��� And tbsType.SelectedItem.Key = "T1" Then
            If .MouseCol = 0 And .MouseRow = 0 Then
                .ToolTipText = "˫��ȫѡ,�ٴ�˫��ȫ��ȡ��."
            Else
                .ToolTipText = ""
            End If
        End If
    End With
End Sub

Private Sub vsfMain_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsfMain(Index).EditSelStart = 0
    vsfMain(Index).EditSelLength = Len(vsfMain(Index).EditText)
End Sub

Private Sub vsfMain_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsfMain(Index)
        If Not IsNumeric(.EditText) Then Cancel = True: Exit Sub
        If Val(.EditText) > Val(.TextMatrix(Row, ColApply("����"))) Then
            stbThis.Panels(2).Text = "�����������ܴ��ڿ���������!"
            Cancel = True
        Else
            stbThis.Panels(2).Text = ""
        End If
    End With
End Sub

Private Sub vsfMain_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub


Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset
    
    If mblnUnChange Then Exit Sub
    With vsfDetail
        If OldCol <> NewCol Then
            .Editable = flexEDNone
            If Val(.RowData(NewRow)) = 0 Or Not (mbytFun = E���� And tbsType.SelectedItem.Key = "T1") Then Exit Sub
            
            If NewCol = .ColIndex("��������") Then .Editable = flexEDKbdMouse
        End If
        If OldRow <> NewRow Then
            vsfTogether.Visible = False
            If Val(.RowData(NewRow)) <> 0 And mbytFun = E���� Then
                If tbsType.SelectedItem.Key = "T1" Then
                    Set rsTmp = mrsApply
                Else
                    Set rsTmp = mrsApplied
                End If
                rsTmp.Filter = "ID=" & Val(.RowData(NewRow))    'ע��,����ı�ԭ��¼����Filter
                If InStr(1, ",5,6,7,", "," & rsTmp!�շ���� & ",") > 0 And Not IsNull(rsTmp!ҽ�����) Then
                    '��ʾһ����ҩ���
                    Call ShowTogetherMedi(Val(rsTmp!ҽ�����), Val(.RowData(NewRow)))
                End If
            End If
            Call Form_Resize
        End If
    End With
End Sub

Private Sub ShowTogetherMedi(ByVal lngAdviceID As Long, ByVal lngFeeItemID As Long)
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
        
    vsfTogether.Clear
    vsfTogether.Rows = 1
    vsfTogether.TextMatrix(0, 0) = "һ����ҩҩƷ"
 
    strSql = "Select 1" & vbNewLine & _
            "From סԺ���ü�¼ A, סԺ���ü�¼ B" & vbNewLine & _
            "Where A.ID = [1] And A.ҽ����� is Not Null And A.NO = B.NO And A.��¼���� = B.��¼���� And A.��¼״̬ = B.��¼״̬ And A.ִ��״̬ = B.ִ��״̬ And" & vbNewLine & _
            "      A.�շ�ϸĿid = B.�շ�ϸĿid And A.�Ǽ�ʱ�� = B.�Ǽ�ʱ�� Having Count(A.ID) > 1"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngFeeItemID)
    If rsTmp.RecordCount > 0 Then
        strSql = "Select B.ҽ������ From ����ҽ����¼ A, ����ҽ����¼ B" & vbNewLine & _
                "Where A.ID = [1] And A.���id = B.���id And A.ID <> B.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
        If rsTmp.RecordCount > 0 Then
            Set vsfTogether.DataSource = rsTmp
            vsfTogether.TextMatrix(0, 0) = "һ����ҩҩƷ"
        End If
    End If
    vsfTogether.Visible = vsfTogether.Rows > 1
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetail_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsfDetail.EditSelStart = 0
    vsfDetail.EditSelLength = Len(vsfDetail.EditText)
End Sub

Private Sub vsfDetail_EnterCell()
    With vsfDetail
        .BackColorSel = .CellBackColor
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub vsfDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblTotal As Double
    
    With vsfDetail
        If Not IsNumeric(.EditText) Then Cancel = True: Exit Sub
        dblTotal = Val(.TextMatrix(Row, .ColIndex("����")) * .TextMatrix(Row, .ColIndex("����")))
        If Val(.EditText) > dblTotal Then
            stbThis.Panels(2).Text = "�����������ܴ��ڿ���������!"
            Cancel = True
        Else
            stbThis.Panels(2).Text = ""
            If Val(.EditText) < dblTotal And Val(.EditText) <> 0 Then
                If Not MCPAR.���ֳ�����ϸ And Not IsNull(mrsInfo!����) Then
                    stbThis.Panels(2).Text = "��������ҽ�����˽��в��ֳ�����ϸ."
                    Cancel = True
                    Exit Sub
                End If
            End If
            If .ColIndex("ִ��״̬") < 0 Then
                mrsApply.Filter = "ID=" & .RowData(Row)
            Else
                mrsApply.Filter = "ID=" & .RowData(Row) & " And ִ��״̬=" & Val(.Cell(flexcpData, .Row, .ColIndex("ִ��״̬")))
            End If
            If mrsApply.RecordCount > 0 Then
                '�����Һ��ҩ�����Ƿ�����δ��ҩ����
                '����:?????
                If InStr(1, "4,5,6,7", Nvl(mrsApply!�շ����)) > 0 And Val(Nvl(mrsApply!ҽ�����)) <> 0 And .ColIndex("ִ��״̬") >= 0 Then
                    If Val(.Cell(flexcpData, .Row, .ColIndex("ִ��״̬"))) = 0 Then 'ֻ��δִ�в��ֲŻ���ڼ��
                        If ҩƷ������ҩ����(Val(Nvl(mrsApply!ҽ�����))) Then
                            stbThis.Panels(2).Text = "��Һ��ҩ�����Ѿ�ʹ���˸�ҩƷ������."
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                If Not IsNull(mrsApply!����id) Then 'Ŀǰδ���ʵ�û����ȡ����,����ĳ�����ʱû��ʹ��
                    If CheckBalance(mrsApply!NO, mrsApply!���) Then
                        If Not IsNull(mrsInfo!����) Then
                            If Not MCPAR.�����ѽ��ʵ��� Then
                                stbThis.Panels(2).Text = "����������ҽ�������ѽ��ʵĵ���."
                                Cancel = True
                            End If
                        Else
                            Select Case gbytBillOpt
                                Case 1
                                    If MsgBox("����[" & mrsApply!NO & "]�еĵ�ǰ������Ŀ�Ѿ�����,ȷ��Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
                                Case 2
                                    stbThis.Panels(2).Text = "����[" & mrsApply!NO & "]�еĵ�ǰ������Ŀ�Ѿ�����,�������ʣ�"
                                    Cancel = True
                            End Select
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       With vsfDetail
            If .Row < .Rows - 1 Then KeyAscii = 0: .Row = .Row + 1
       End With
    End If
End Sub


Private Sub vsfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsfDetail
            If Val(.RowData(.Row)) = 0 Or Not (mbytFun = E���� And tbsType.SelectedItem.Key = "T1") Then Exit Sub
            
            If .Col = .ColIndex("��������") Then
                .TextMatrix(.Row, .Col) = "0"
            End If
        End With
    End If
End Sub

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblLack As Double
    Dim lngCol As Long
    Dim dblMny As Double
    
    lngCol = vsfDetail.ColIndex("ִ��״̬")
    mrsApply.Filter = "ID=" & vsfDetail.RowData(Row) & IIf(lngCol >= 0, " And ִ��״̬=" & Val(vsfDetail.Cell(flexcpData, Row, lngCol)), "")
    If mrsApply.RecordCount > 0 Then
        dblLack = Val(vsfDetail.EditText) - mrsApply!��������
        dblMny = (Val(vsfDetail.EditText) - mrsApply!��������) * mrsApply!����
        
        mrsApply!�������� = vsfDetail.EditText
        mrsApply!���ʽ�� = mrsApply!�������� * mrsApply!����
        mrsApply.Update
        vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("���ʽ��")) = FormatEx(mrsApply!���ʽ��, 5)
        vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("��������")) = vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("��������")) + dblLack
        vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("���ʽ��")) = vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("���ʽ��")) + dblMny
        Call ShowSumMoney
    End If
End Sub
Private Sub cmdCancelApply_Click()
    Call SaveData
    gblnOK = True
End Sub

Private Sub cmdOKApply_Click()
    If mlngDeptID = 0 Then
        MsgBox "û��ѡ�����벿��, ����ȷ������!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If SaveData = False Then Exit Sub
    '����:26551
    'gblnOK = True
End Sub

Private Sub cmdOKAudit_Click()
    If mbytFun = E��� And tbsType.SelectedItem.Key = "T2" Then
        If SaveRefuse(False) = False Then Exit Sub
    Else
        If SaveData = False Then Exit Sub
        gblnOK = True
    End If
End Sub

Private Sub cmdCancelRefuse_Click()
    If SaveRefuse(True) = False Then Exit Sub
End Sub

Private Function SaveData() As Boolean

    Dim arrSQL As Variant
    Dim i As Long, j As Long, lngTmp As Long, bytState As Byte
    Dim strDate As String, str����IDs As String, str����ID As String, strTmp As String
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant, arrNO As Variant
    Dim dbl���� As Double, strNos As String
    Dim str��˷���ID As String
    Dim strMsgDate As String
    
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            With mrsApply
                If .State = 0 Then Exit Function
                .Filter = ""
                For i = 1 To .RecordCount
                    If !�������� <> !ԭʼ�������� Then
                        str����IDs = str����IDs & IIf(str����IDs = "", "", ",") & !ID
                        str��˷���ID = str��˷���ID & "," & !ID
                    End If
                    .MoveNext
                Next
                
                If str����IDs = "" Then
                    stbThis.Panels(2).Text = "���м�¼��û����д��������!"
                    Exit Function
                End If
            End With
        Else
            For i = 1 To vsfMain(1).Rows - 1
                If vsfMain(1).TextMatrix(i, ColApplied("ѡ��")) = "��" And Val(vsfMain(1).RowData(i)) <> 0 Then Exit For
            Next
            If i > vsfMain(1).Rows - 1 Then
                stbThis.Panels(2).Text = "û��ѡ��ȡ������ļ�¼!����Ҫȡ������ļ�¼��""ѡ��""����˫����"
                Exit Function
            End If
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            For i = 1 To vsfMain(0).Rows - 1
                If Val(vsfMain(0).RowData(i)) <> 0 Then
                    If vsfMain(0).TextMatrix(i, ColAudit("���")) = "��" Or vsfMain(0).TextMatrix(i, ColAudit("���")) = "��" Then Exit For
                End If
            Next
            If i > vsfMain(0).Rows - 1 Then
                stbThis.Panels(2).Text = "û��ѡ����˵ļ�¼!����Ҫ��˵ļ�¼��""���""����˫����"
                Exit Function
            End If
        End If
    End If
    
    
    arrSQL = Array()
    If str��˷���ID <> "" Then str��˷���ID = Mid(str��˷���ID, 2)
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            strMsgDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            strDate = "To_Date('" & strMsgDate & "','YYYY-MM-DD HH24:MI:SS')"
            Dim strKey����IDs As String
            
            With mrsApply
                .MoveFirst
                strKey����IDs = ""
                For i = 1 To .RecordCount
                    If InStr(1, "," & str����IDs & ",", "," & !ID & ",") > 0 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    
                        dbl���� = !�������� * !סԺ��װ
                        If !�������� = !���� Then dbl���� = !�ۼ�����
                        '  Id_In         ���˷�������.����id%Type,
                        '  �շ�ϸĿid_In ���˷�������.�շ�ϸĿid%Type,
                        '  ���벿��id_In ���˷�������.���벿��id%Type,
                        '  ����_In       ���˷�������.����%Type,
                        '  ������_In     ���˷�������.������%Type,
                        '  ����ʱ��_In   ���˷�������.����ʱ��%Type,
                        '  ״̬_In       סԺ���ü�¼.ִ��״̬%Type := Null --��ҩƷ��������Ч:0-δ��ҩ(��);1-�ѷ�ҩ(��);NULL
                        '   ɾ����־_In   INTEGER :=0 --ɾ�����˷�������ʱ������:1-ɾ��ʱ�����������,0-ɾ��ʱ,�����������������ɾ��(��Ϊ���ܳ�������������ʱ,������ִ�к�δִ������״̬)
                        arrSQL(UBound(arrSQL)) = "zl_���˷�������_Insert(" & !ID & "," & !�շ�ϸĿID & "," & _
                                mlngDeptID & "," & dbl���� & ",'" & UserInfo.���� & "'," & strDate & "," & Val(Nvl(!ִ��״̬)) & "," & _
                              IIf(InStr(1, "," & strKey����IDs & ",", "," & Nvl(!ID) & ",") > 0, 0, 1) & ")"
                        strKey����IDs = strKey����IDs & "," & !ID
                        If InStr("," & strNos, "," & !NO) = 0 Then
                            strNos = strNos & "," & !NO & "|" & Format(!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS") & "|" & !����Ա����
                        End If
                    End If
                    .MoveNext
                Next
            End With
        Else
            With mrsApplied
                For i = 1 To vsfMain(1).Rows - 1
                    If vsfMain(1).TextMatrix(i, ColApplied("ѡ��")) = "��" Then
                        .Filter = "�շ�ϸĿID=" & vsfMain(1).RowData(i) & " And ������='" & vsfMain(1).TextMatrix(i, ColApplied("������")) & _
                                "' And ����ʱ��='" & vsfMain(1).TextMatrix(i, ColApplied("����ʱ��")) & "'"
                        For j = 1 To .RecordCount
                            str����IDs = str����IDs & IIf(str����IDs = "", "", ",") & !ID
                            .MoveNext
                        Next
                    End If
                Next
            End With
            While str����IDs <> ""
                str����IDs = str����IDs & ","
                If Len(str����IDs) > 3998 Then
                    lngTmp = InStrRev(Mid(str����IDs, 1, 3998), ",")
                    str����ID = Mid(str����IDs, 1, lngTmp - 1)
                    str����IDs = Mid(str����IDs, lngTmp + 1)
                Else
                    str����ID = Mid(str����IDs, 1, Len(str����IDs) - 1)
                    str����IDs = ""
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_���˷�������_Delete('" & str����ID & "')"
            Wend
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            With mrsAudit
                For i = 1 To vsfMain(0).Rows - 1
                    strTmp = vsfMain(0).TextMatrix(i, ColAudit("���"))
                    If strTmp = "��" Or strTmp = "��" Then
                        .Filter = "�շ�ϸĿID=" & vsfMain(0).RowData(i) & " And �������=" & Val(vsfMain(0).Cell(flexcpData, i, ColAudit("���"))) & " And ������='" & vsfMain(0).TextMatrix(i, ColAudit("������")) & _
                                "' And ����ʱ��='" & vsfMain(0).TextMatrix(i, ColAudit("����ʱ��")) & "'"
                                
                        For j = 1 To .RecordCount
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            If zlCheckFeeIsValied(Val(Nvl(!ID)), Val(Nvl(!��˲���id)), Val(vsfMain(0).Cell(flexcpData, i, ColAudit("���")))) = False Then Exit Function
                            ' �������_In ���˷�������.�������%Type := 1 --��ҩƷ��������Ч,ȱʡΪ��ִ�е�ҩƷ������
                            arrSQL(UBound(arrSQL)) = "zl_���˷�������_Audit(" & !ID & ",To_Date('" & !����ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                                UserInfo.���� & "'," & strDate & "," & IIf(strTmp = "��", "1", "2") & ",1," & Val(vsfMain(0).Cell(flexcpData, i, ColAudit("���"))) & ")"
                                    
                            If strTmp = "��" Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                '����:49206 --����״̬_In:0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������)
                                arrSQL(UBound(arrSQL)) = "ZL_סԺ���ʼ�¼_Delete('" & !NO & "','" & !��� & ":" & !�ۼ��������� & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & !��¼���� & ",1)"
                            End If
                                    
                            If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                                MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                                MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                                strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                                        "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                            End If
                            If InStr("," & strNos, "," & !NO) = 0 Then
                                strNos = strNos & "," & !NO & "|" & Format(!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS") & "|" & !����Ա����
                            End If
                            
                            .MoveNext
                        Next
                    End If
                Next
            End With
        End If
    End If
    
    If strNos <> "" Then    '������������ʱ
        arrNO = Split(Mid(strNos, 2), ",")
        For i = 0 To UBound(arrNO)
            If Not BillOperCheck(5, CStr(Split(arrNO(i), "|")(2)), CDate(Split(arrNO(i), "|")(1)), IIf(mbytFun = E����, "��������", "�������"), _
                Split(arrNO(i), "|")(0), , 2, , False, False) Then Exit Function
        Next
    End If
   
    On Error GoTo errH
    Screen.MousePointer = 11
    gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        
        'ҽ�������������ϴ�������ʱ�ϴ�
        If strMCNO <> "" Then
            arrMCRec = Split(strMCNO, "|")
            For i = 0 To UBound(arrMCRec)
                arrMCPar = Split(arrMCRec(i), ",")
                If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                    If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            Next
        End If
    gcnOracle.CommitTrans
    
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
    '����:34994
    '   ������˲���
    If mbytFun = E���� And chkVerfy.Visible And chkVerfy.Value = 1 Then
        If zlApplyToVerify(str��˷���ID) = False Then
            MsgBox "ע��:" & vbCrLf & "    ���ڲ�����˵�����,��ͨ��������˽�����˲���!", vbInformation + vbOKOnly, gstrSysName
        End If
    End If
    Screen.MousePointer = 0
    
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            '��Ϣ����
            If Not (chkVerfy.Visible And chkVerfy.Value = 1) Then
                Call SendMsgModule(str����IDs, strMsgDate)
            End If
            txtPatient.Text = "": txtPatient.SetFocus
            Call ClearPatientInfo
        Else
            Call cmdRefresh_Click
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            Call cmdRefresh_Click
        End If
    End If
    
    stbThis.Panels(2).Text = "�������ݳɹ�!"
    SaveData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function
Private Function zlCheckFeeIsValied(ByVal lng����ID As Long, ByVal lng��˲���ID As Long, Optional int������� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������Ƿ���Ч
    '���:int�������-1-��ִ��;0-δִ��
    '����:
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2009-07-28 09:48:59
    '����:24597
    '����:1.�����������δ��ִ�У�����ԭ���Ĺ��򲻱䣬˭������˭�Ϳ��Խ�������
    '     2.����������ñ�ִ��,����Ҫ�ж��������:
    '        a.�����˿�����ִ�п������,���������ȷ��
    '        b.�����˿�����ִ���Ҳ���ȣ�����Ҫ���ִ�п����Ƿ��ڵ�ǰ����Ա��Ա�������������,����ǣ���������ˣ������������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lngִ�в���ID As Long, strNo As String, str��Ŀ���� As String, str�������� As String
    
    On Error GoTo errHandle
    
    '֮����Ҫ���¶�ȡִ��״̬�����ǲ��������������
    gstrSQL = "Select A.NO,a.ִ��״̬,a.ִ�в���ID,a.�շ�ϸĿID,a.�շ����, nvl(b.��������,0) as �������� From סԺ���ü�¼ a ,�������� B Where a.�շ�ϸĿID=b.����ID(+) and a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    If rsTemp.EOF Then
        MsgBox "ע��:" & vbCrLf & _
               "    ���������һ����ϸ���ò����ڣ����ܱ�����ɾ������ˢ�º����ԣ�", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    lngִ�в���ID = Val(Nvl(rsTemp!ִ�в���ID))
    '1.�����������δ��ִ�У�����ԭ���Ĺ��򲻱䣬˭������˭�Ϳ��Խ�������
    '��¼״̬=1,3ʱ��0:δִ��;1:��ȫִ��;2:����ִ�У���¼״̬=2ʱ��-x:��x���˷�
    If Val(Nvl(rsTemp!ִ��״̬)) = 0 Then zlCheckFeeIsValied = True: Exit Function
    
    '����������ñ�ִ��,����Ҫ�ж��������:
    '1. �����˿�����ִ�п������,���������ȷ��
    If lng��˲���ID = lngִ�в���ID Then zlCheckFeeIsValied = True: Exit Function
    '2  �����˿�����ִ���Ҳ���ȣ�����Ҫ���ִ�п����Ƿ��ڵ�ǰ����Ա��Ա�������������,����ǣ���������ˣ������������
    If InStr(1, "," & mstrUnitIDs & ",", "," & lngִ�в���ID & ",") > 0 Then zlCheckFeeIsValied = True: Exit Function
    strNo = Nvl(rsTemp!NO)
    
    '3.�����ҩƷ,����,��Ҫ���
    If InStr(1, "5,6,7", Nvl(rsTemp!�շ����)) > 0 Or (Nvl(rsTemp!�շ����) = 4 And Nvl(rsTemp!��������) = "1") Then
            If int������� = 0 Then
                zlCheckFeeIsValied = True: Exit Function
            End If
    End If
    
    gstrSQL = "Select ����,���� From �շ���ĿĿ¼ a Where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Nvl(rsTemp!�շ�ϸĿID)))
    If Not rsTemp.EOF Then
        str��Ŀ���� = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
    End If
    
    
    gstrSQL = "Select ����,���� From ���ű� a Where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngִ�в���ID)
    If Not rsTemp.EOF Then
        str�������� = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
    End If
    
    MsgBox "ע��:" & vbCrLf & _
           "    ���ݺ�Ϊ��" & strNo & "��" & vbCrLf & _
           "    �շ���ĿΪ��" & str��Ŀ���� & "��" & vbCrLf & _
           "    �Ѿ�����" & str�������� & "�� ִ�У�����ȷ�����ʣ�", vbInformation + vbDefaultButton1, gstrSysName

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ShowSumMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���������ܶ�
    '����:���˺�
    '����:2011-02-15 16:57:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, vsGrid As VSFlexGrid, lngRow As Long
    Dim lngCol As Long
    Err = 0: On Error Resume Next
    If mbytFun = E���� Then
        If tbsType.SelectedItem.Key = "T1" Then
            Set vsGrid = vsfMain(0): lngCol = ColApply("���ʽ��")
        Else
            Set vsGrid = vsfMain(1): lngCol = ColApplied("���ʽ��")
        End If
        With vsGrid
            For lngRow = .FixedRows To .Rows - 1
                dblMoney = dblMoney + Val(.TextMatrix(lngRow, lngCol))
            Next
        End With
        picHsc.Height = 435
        picHsc.Cls
        picHsc.CurrentY = 100: picHsc.CurrentX = 50
        picHsc.FontBold = True
        picHsc.Print "���ʽ��ϼ�:" & FormatEx(dblMoney, 5)
    Else
        picHsc.Height = 30
    End If
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKind.Cards.��ȱʡ������
End Sub
Private Sub LoadBabyCombox(ByVal lng����ID As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ӥ���ѵ������Ϣ��Combox����
    '����:���˺�
    '����:2013-04-10 17:36:17
    '˵��:
    '����:55368
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, intCount As Integer
    intCount = Val(zlDatabase.GetPara("��������Ӥ������ʾ����", glngSys, Enum_Inside_Program.p���ʲ���, "1", Array(cboBaby), InStr(mstrPrivsOpt, ";����ѡ������;") > 0))
    On Error GoTo errHandle
    With cboBaby
        .Clear
        .AddItem "������Ӥ������"
        .ItemData(.NewIndex) = 0
        If intCount = 0 Then .ListIndex = .NewIndex
        .AddItem "����Ӥ������"
        .ItemData(.NewIndex) = 1
        If intCount = 1 Then .ListIndex = .NewIndex
        For i = 1 To 5
            .AddItem "����ʾ��" & i & "��Ӥ������"
            .ItemData(.NewIndex) = i + 1
            If intCount = i + 1 Then .ListIndex = .NewIndex
        Next
        If .ListIndex < 0 Then .ListIndex = 0
    End With
     Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetOperatorDept() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ա����������(����Աֻ��Ϊ��ʿʱ)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-04-24 11:33:20
    '����:60679
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mrsOperatorDept Is Nothing Then
       Set GetOperatorDept = mrsOperatorDept
       Exit Function
    End If
    Set mrsOperatorDept = GetDepartments("", "1,2,3", True, True)
    Set GetOperatorDept = mrsOperatorDept
 End Function



Private Function zlMsgModule_Init() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣģ��
    '���:lngModule -ģ���
    '     strPivs-Ȩ�޴�
    '����:objMsgModule-������Ϣ����
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Err = 0: On Error GoTo ErrHand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    zlMsgModule_Init = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlMsgModule_Unload() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ж��Ϣģ��
    '���:objMsgModule-��Ϣ����
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    
    If mobjMsgModule Is Nothing Then Exit Function
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    zlMsgModule_Unload = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Sub SendMsgModule(ByVal str����IDs As String, ByVal strDate As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ϣ���ʹ���
    '���:
    '����:���˺�
    '����:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSelIDs As String
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mbytFun <> 0 Then Exit Sub
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    
    strSql = "" & _
    "   Select A.����id, A.�������, A.�շ�ϸĿid,B.���� as ������Ŀ, B.���㵥λ," & _
    "       A.��˲���id,C.���� as ��˲���, A.���벿��id,D.���� as ���벿��, " & _
    "       A.����, A.������, A.����ʱ��, A.״̬ " & _
    "   From ���˷������� A,�շ���ĿĿ¼ B,���ű� C,���ű� D ,Table(f_Num2List([1])) M" & _
    "   where A.�շ�ϸĿID=B.ID and A.��˲���ID=C.ID(+) and A.���벿��ID=D.ID(+)" & _
    "         And A.����ID=M.Column_value And A.����ʱ��=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str����IDs, CDate(strDate))
    If rsTemp.EOF Then Exit Sub
                        
    zlXML.ClearXmlText
        'ZLHIS_CHARGE_001 ������������֪ͨ
    '�ڵ�����    ����    ����    �ظ�    ����    ȱʡֵ  ֵ������
    'patient_info        ������Ϣ    1
    '   patient_id      ����id  1   N
    '   page_id     ��ҳid  1   N
    '   patient_name        ����    1   S
    '   patient_sex     �Ա�    1   S
    '   patient_age     ����    1   S
    '   identity_card       ����֤��    0..1    S
    '   in_number       סԺ��  0..1    S
    '   out_number      �����  0..1    S
    'cancel_reqeust      ��������    1
    '   cancel_charge           1..*
    '       charge_id       ����id  1   N
    '       request_kind        �������    1   N
    '       request_time        ����ʱ��    1   S
    '       request_person      ������Ա    1   S
    '       cancel_item_id      ������Ŀid  1   N
    '       cancel_item_title       ������Ŀ    1   S
    '       calcel_num      ��������    1   N
    '       charge_unit     ���õ�λ    1   S
    '       audit_dept_id       ��˲���id  1   N
    '       audit_dept_title        ��˲���    1   S
    Call zlXML.AppendNode("patient_info")
        Call zlXML.appendData("patient_id", Val(Nvl(mrsInfo!����ID)))
        Call zlXML.appendData("page_id", Val(Nvl(mrsInfo!��ҳID)))
        Call zlXML.appendData("patient_name", Nvl(mrsInfo!����))
        Call zlXML.appendData("patient_sex", Nvl(mrsInfo!�Ա�))
        Call zlXML.appendData("patient_age", Nvl(mrsInfo!����))
        Call zlXML.appendData("identity_card", Nvl(mrsInfo!����֤��))
        Call zlXML.appendData("in_number", Nvl(mrsInfo!סԺ��))
        Call zlXML.appendData("out_number", Nvl(mrsInfo!�����))
    Call zlXML.AppendNode("patient_info", True)
    
    Call zlXML.AppendNode("cancel_reqeust")
        
    With rsTemp
        .MoveFirst
        Do While Not .EOF
            Call zlXML.AppendNode("cancel_charge")
            '       charge_id       ����id  1   N
                Call zlXML.appendData("charge_id", Val(Nvl(!����ID)))
            '       request_kind        �������    1   N
                Call zlXML.appendData("request_kind", Val(Nvl(!�������)))
            '       request_time        ����ʱ��    1   D
                Call zlXML.appendData("request_time", Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS"))
            '       request_person      ������Ա    1   S
                Call zlXML.appendData("request_person", Nvl(!������))
            '       cancel_item_id      ������Ŀid  1   N
                Call zlXML.appendData("cancel_item_id", Val(Nvl(!�շ�ϸĿID)))
            '       cancel_item_title       ������Ŀ    1   S
                Call zlXML.appendData("cancel_item_title", Trim(Nvl(!������Ŀ)))
            '       calcel_num      ��������    1   N
                Call zlXML.appendData("calcel_num", Val(Nvl(!����)))
            '       charge_unit     ���õ�λ    1   S
                Call zlXML.appendData("charge_unit", Trim(Nvl(!���㵥λ)))
            '       audit_dept_id       ��˲���id  1   N
                Call zlXML.appendData("audit_dept_id", Val(Nvl(!��˲���id)))
            '       audit_dept_title        ��˲���    1   S
                Call zlXML.appendData("audit_dept_title", Trim(Nvl(!��˲���)))
            Call zlXML.AppendNode("cancel_charge", True)
            .MoveNext
        Loop
    End With
    Call zlXML.AppendNode("cancel_reqeust", True)
    
    If Not mobjMsgModule Is Nothing Then
        If mobjMsgModule.IsConnect = True Then
        '�������Ϣ
            Call mobjMsgModule.CommitMessage("ZLHIS_CHARGE_001", zlXML.XmlText)
        End If
    End If
    
    Call zlDatabase.SendMsg("ZLHIS_CHARGE_001", zlXML.XmlText)
    zlXML.ClearXmlText
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub


