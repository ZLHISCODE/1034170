VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.3#0"; "ZlPatiAddress.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPatient 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人登记"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "frmPatient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9434.273
   ScaleMode       =   0  'User
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicHealth 
      BorderStyle     =   0  'None
      Height          =   8280
      Left            =   12240
      ScaleHeight     =   8280
      ScaleMode       =   0  'User
      ScaleWidth      =   14791.25
      TabIndex        =   136
      Top             =   120
      Width           =   11805
      Begin VB.Frame Frame3 
         Height          =   105
         Left            =   915
         TabIndex        =   156
         Top             =   915
         Width           =   10575
      End
      Begin VB.CommandButton cmdMedicalWarning 
         Caption         =   "…"
         Height          =   255
         Left            =   11249
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   68
         Width           =   239
      End
      Begin VB.Frame Frame2 
         Height          =   105
         Left            =   915
         TabIndex        =   148
         Top             =   2280
         Width           =   10575
      End
      Begin VB.Frame Frame1 
         Height          =   105
         Left            =   945
         TabIndex        =   147
         Top             =   5250
         Width           =   10545
      End
      Begin VB.Frame frameLinkMan 
         Height          =   105
         Left            =   1125
         TabIndex        =   146
         Top             =   3690
         Width           =   10380
      End
      Begin VB.TextBox txtOtherWaring 
         Height          =   300
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   141
         Top             =   420
         Width           =   10280
      End
      Begin VB.TextBox txtMedicalWarning 
         Height          =   300
         Left            =   7355
         Locked          =   -1  'True
         TabIndex        =   139
         Top             =   45
         Width           =   4155
      End
      Begin VB.ComboBox cboBH 
         Height          =   300
         Left            =   4095
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   45
         Width           =   1770
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   45
         Width           =   1770
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLinkMan 
         Height          =   1020
         Left            =   150
         TabIndex        =   144
         Top             =   3930
         Width           =   11385
         _cx             =   1973046898
         _cy             =   1973028615
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
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
         ExtendLastCol   =   0   'False
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsOtherInfo 
         Height          =   2460
         Left            =   150
         TabIndex        =   145
         Top             =   5535
         Width           =   11385
         _cx             =   1973046898
         _cy             =   1973031155
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInoculate 
         Height          =   975
         Left            =   150
         TabIndex        =   143
         Top             =   2520
         Width           =   11385
         _cx             =   1973046898
         _cy             =   1973028536
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
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
         ColWidthMax     =   2287
         ExtendLastCol   =   0   'False
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDrug 
         Height          =   975
         Left            =   150
         TabIndex        =   142
         Top             =   1155
         Width           =   11385
         _cx             =   1973046898
         _cy             =   1973028536
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "过敏反应"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   150
         TabIndex        =   157
         Top             =   915
         Width           =   720
      End
      Begin VB.Label lblInoculate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "接种情况"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   150
         TabIndex        =   155
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lblOtherInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "其他信息"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   150
         TabIndex        =   154
         Top             =   5265
         Width           =   720
      End
      Begin VB.Label lblLinkman 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "联系人信息"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   165
         TabIndex        =   153
         Top             =   3675
         Width           =   900
      End
      Begin VB.Label lblOtherWaring 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "其他医学警示"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   152
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label lblMedicalWarning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "医学警示"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6600
         TabIndex        =   151
         Top             =   105
         Width           =   720
      End
      Begin VB.Label lblRH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RH"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3810
         TabIndex        =   150
         Top             =   105
         Width           =   195
      End
      Begin VB.Label lblBloodType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "血型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   149
         Top             =   105
         Width           =   360
      End
   End
   Begin VB.PictureBox PicBaseInfo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8160
      Left            =   60
      ScaleHeight     =   8160
      ScaleWidth      =   11730
      TabIndex        =   77
      Top             =   360
      Width           =   11730
      Begin VB.Frame fraCard 
         Caption         =   "【就诊卡信息】"
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   45
         TabIndex        =   129
         Top             =   7200
         Width           =   11640
         Begin VB.ComboBox cbo结算方式 
            Height          =   300
            Left            =   8925
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   370
            Width           =   2550
         End
         Begin VB.TextBox txt卡号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   720
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   66
            Top             =   360
            Width           =   1425
         End
         Begin VB.TextBox txtPass 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2640
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   67
            Top             =   370
            Width           =   1335
         End
         Begin VB.TextBox txt卡额 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00C00000&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   370
            Width           =   1335
         End
         Begin VB.CheckBox chk记帐 
            Caption         =   "记帐"
            Height          =   195
            Left            =   8280
            TabIndex        =   70
            Top             =   435
            Width           =   675
         End
         Begin VB.TextBox txtAudi 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4455
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   68
            Top             =   370
            Width           =   1335
         End
         Begin MSComctlLib.TabStrip tabCardMode 
            Height          =   315
            Left            =   120
            TabIndex        =   89
            Top             =   0
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            Style           =   2
            TabFixedHeight  =   526
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            TabMinWidth     =   882
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "发卡收费(&1)"
                  Key             =   "CardFee"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "绑定卡号(&2)"
                  Key             =   "CardBind"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl卡名称 
            Height          =   255
            Left            =   8925
            TabIndex        =   15
            Top             =   0
            Width           =   1590
         End
         Begin VB.Label lbl就诊卡号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "卡号"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   134
            Top             =   400
            Width           =   420
         End
         Begin VB.Label lbl密码 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "密码"
            Height          =   180
            Left            =   2250
            TabIndex        =   133
            Top             =   435
            Width           =   360
         End
         Begin VB.Label lbl金额 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "金额"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5870
            TabIndex        =   132
            Top             =   435
            Width           =   360
         End
         Begin VB.Label lbl验证 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "验证"
            Height          =   180
            Left            =   4065
            TabIndex        =   131
            Top             =   435
            Width           =   360
         End
         Begin VB.Label lbl结算方式 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算方式"
            Height          =   180
            Left            =   8115
            TabIndex        =   130
            Top             =   430
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "【预交款信息】"
         ForeColor       =   &H00C00000&
         Height          =   1230
         Left            =   45
         TabIndex        =   120
         Top             =   5880
         Width           =   11640
         Begin VB.ComboBox cbo预交结算 
            Height          =   300
            Left            =   5055
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   390
            Width           =   2550
         End
         Begin VB.TextBox txt预交额 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   2715
            MaxLength       =   12
            TabIndex        =   60
            Top             =   390
            Width           =   1050
         End
         Begin VB.TextBox txt结算号码 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   8925
            MaxLength       =   30
            TabIndex        =   62
            Top             =   390
            Width           =   2550
         End
         Begin VB.TextBox txtFact 
            Height          =   300
            Left            =   1100
            MaxLength       =   50
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   390
            Width           =   1110
         End
         Begin VB.TextBox txt缴款单位 
            Height          =   300
            Left            =   1100
            MaxLength       =   50
            TabIndex        =   63
            Top             =   780
            Width           =   2670
         End
         Begin VB.TextBox txt开户行 
            Height          =   300
            Left            =   5055
            MaxLength       =   50
            TabIndex        =   64
            Top             =   780
            Width           =   2550
         End
         Begin VB.TextBox txt帐号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   8925
            MaxLength       =   50
            TabIndex        =   65
            Top             =   780
            Width           =   2550
         End
         Begin MSComctlLib.TabStrip tbDeposit 
            Height          =   270
            Left            =   90
            TabIndex        =   86
            Top             =   0
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   476
            Style           =   2
            TabFixedHeight  =   526
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            TabMinWidth     =   882
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "门诊预交(&M)"
                  Key             =   "K1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "住院预交(&Z)"
                  Key             =   "K2"
                  ImageVarType    =   2
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
         Begin VB.Label lblMoney 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "金额"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2280
            TabIndex        =   128
            Top             =   450
            Width           =   360
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算号码"
            Height          =   180
            Left            =   8115
            TabIndex        =   127
            Top             =   450
            Width           =   720
         End
         Begin VB.Label lblStyle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "缴款方式"
            Height          =   180
            Left            =   4290
            TabIndex        =   126
            Top             =   450
            Width           =   720
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "摘要"
            Height          =   240
            Left            =   825
            TabIndex        =   125
            Top             =   1605
            Width           =   480
         End
         Begin VB.Label lblFact 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "实际票号"
            Height          =   180
            Left            =   315
            TabIndex        =   124
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "缴款单位"
            Height          =   180
            Left            =   315
            TabIndex        =   123
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开户行"
            Height          =   180
            Left            =   4470
            TabIndex        =   122
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "帐号"
            Height          =   180
            Left            =   8415
            TabIndex        =   121
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lblYBMoney 
            AutoSize        =   -1  'True
            Caption         =   "个人帐户余额:"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   2805
            TabIndex        =   87
            Top             =   45
            Visible         =   0   'False
            Width           =   1170
         End
      End
      Begin VB.Frame fraInfo 
         Height          =   5775
         Left            =   45
         TabIndex        =   78
         Top             =   0
         Width           =   11640
         Begin VB.Frame fraBase 
            BorderStyle     =   0  'None
            Height          =   5565
            Left            =   15
            TabIndex        =   79
            Top             =   120
            Width           =   11520
            Begin VB.ComboBox cboIDNumber 
               Height          =   300
               Left            =   3240
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   1215
               Width           =   1335
            End
            Begin VB.TextBox txtMobile 
               Height          =   300
               Left            =   1110
               MaxLength       =   20
               TabIndex        =   57
               Top             =   5235
               Width           =   4800
            End
            Begin VB.CommandButton cmd联系人地址 
               Caption         =   "…"
               Height          =   255
               Left            =   5595
               TabIndex        =   54
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   4898
               Width           =   285
            End
            Begin VB.TextBox txt验证密码 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   10085
               MaxLength       =   20
               PasswordChar    =   "*"
               TabIndex        =   25
               Top             =   2340
               Width           =   1410
            End
            Begin VB.CommandButton cmdPicClear 
               Caption         =   "清除"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   10815
               TabIndex        =   161
               TabStop         =   0   'False
               Top             =   1530
               Width           =   600
            End
            Begin VB.CommandButton cmdPicCollect 
               Caption         =   "采集"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   10170
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   1530
               Width           =   600
            End
            Begin VB.CommandButton cmdPicFile 
               Caption         =   "文件"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   9540
               TabIndex        =   159
               TabStop         =   0   'False
               Top             =   1530
               Width           =   585
            End
            Begin VB.PictureBox picPatient 
               Height          =   1500
               Left            =   9540
               ScaleHeight     =   1440
               ScaleWidth      =   1815
               TabIndex        =   158
               Top             =   30
               Width           =   1875
               Begin VB.Image imgPatient 
                  Height          =   1425
                  Left            =   15
                  Stretch         =   -1  'True
                  Top             =   15
                  Width           =   1800
               End
            End
            Begin VB.CommandButton cmd区域 
               Caption         =   "…"
               Height          =   255
               Left            =   8370
               TabIndex        =   42
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   3458
               Width           =   285
            End
            Begin VB.CommandButton cmd家庭地址 
               Caption         =   "…"
               Height          =   255
               Left            =   5595
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   2745
               Width           =   285
            End
            Begin VB.CommandButton cmd籍贯 
               Caption         =   "…"
               Height          =   255
               Left            =   11180
               TabIndex        =   36
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   3098
               Width           =   285
            End
            Begin VB.CommandButton cmd户口地址 
               Caption         =   "…"
               Height          =   255
               Left            =   5595
               TabIndex        =   32
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   3105
               Width           =   285
            End
            Begin VB.CommandButton cmd合同单位 
               Caption         =   "…"
               Height          =   255
               Left            =   5595
               TabIndex        =   45
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   3825
               Width           =   285
            End
            Begin VB.CommandButton cmd出生地点 
               Caption         =   "…"
               Height          =   255
               Left            =   5595
               TabIndex        =   39
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   3450
               Width           =   285
            End
            Begin VB.TextBox txt籍贯 
               Height          =   300
               Left            =   9395
               MaxLength       =   30
               TabIndex        =   35
               Top             =   3075
               Width           =   2100
            End
            Begin VB.TextBox txt区域 
               Height          =   300
               Left            =   7200
               MaxLength       =   30
               TabIndex        =   41
               Top             =   3435
               Width           =   1485
            End
            Begin VB.TextBox txt单位邮编 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   10085
               MaxLength       =   6
               TabIndex        =   47
               Top             =   3795
               Width           =   1410
            End
            Begin VB.TextBox txt联系人电话 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   4080
               MaxLength       =   20
               TabIndex        =   51
               Top             =   4515
               Width           =   1815
            End
            Begin VB.TextBox txt联系人姓名 
               Height          =   300
               Left            =   1110
               MaxLength       =   64
               TabIndex        =   50
               Top             =   4515
               Width           =   1815
            End
            Begin VB.TextBox txt单位开户行 
               Height          =   300
               Left            =   1110
               MaxLength       =   50
               TabIndex        =   48
               Top             =   4155
               Width           =   4800
            End
            Begin VB.TextBox txt单位电话 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   20
               TabIndex        =   46
               Top             =   3795
               Width           =   1485
            End
            Begin VB.TextBox txt家庭电话 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   20
               TabIndex        =   29
               Top             =   2715
               Width           =   1485
            End
            Begin VB.TextBox txt家庭地址邮编 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   10085
               MaxLength       =   6
               TabIndex        =   30
               Top             =   2715
               Width           =   1410
            End
            Begin VB.TextBox txt工作单位 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   44
               Top             =   3795
               Width           =   4785
            End
            Begin VB.ComboBox cbo年龄单位 
               Height          =   300
               Left            =   4005
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   855
               Width           =   580
            End
            Begin VB.TextBox txtPatiMCNO 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   1110
               MaxLength       =   30
               TabIndex        =   20
               Top             =   1950
               Width           =   4785
            End
            Begin VB.ComboBox cbo医疗付款 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   4080
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   2340
               Width           =   1815
            End
            Begin VB.TextBox txt住院号 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   5655
               MaxLength       =   18
               TabIndex        =   2
               Top             =   120
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.TextBox txt门诊号 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3045
               MaxLength       =   18
               TabIndex        =   1
               Top             =   120
               Width           =   1545
            End
            Begin VB.TextBox txt年龄 
               Height          =   300
               IMEMode         =   2  'OFF
               Left            =   3180
               TabIndex        =   9
               Top             =   855
               Width           =   800
            End
            Begin VB.ComboBox cbo联系人关系 
               Height          =   300
               Left            =   7200
               TabIndex        =   56
               Text            =   "cbo联系人关系"
               Top             =   4875
               Width           =   4295
            End
            Begin VB.ComboBox cbo婚姻状况 
               Height          =   300
               Left            =   5655
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1560
               Width           =   1485
            End
            Begin VB.ComboBox cbo学历 
               Height          =   300
               Left            =   7905
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   870
               Width           =   1485
            End
            Begin VB.ComboBox cbo国籍 
               Height          =   300
               Left            =   5655
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1230
               Width           =   1485
            End
            Begin VB.ComboBox cbo民族 
               Height          =   300
               Left            =   5655
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   870
               Width           =   1485
            End
            Begin VB.ComboBox cbo职业 
               Height          =   300
               Left            =   7905
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   1230
               Width           =   1485
            End
            Begin VB.ComboBox cbo身份 
               Height          =   300
               Left            =   7905
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   495
               Width           =   1485
            End
            Begin VB.ComboBox cbo费别 
               Height          =   300
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   2325
               Width           =   1815
            End
            Begin VB.ComboBox cbo性别 
               Height          =   300
               Left            =   5655
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   495
               Width           =   1485
            End
            Begin VB.TextBox txtPatient 
               Height          =   300
               Left            =   1110
               TabIndex        =   4
               Top             =   495
               Width           =   3480
            End
            Begin VB.TextBox txt病人ID 
               ForeColor       =   &H00C00000&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   1110
               Locked          =   -1  'True
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   120
               Width           =   1170
            End
            Begin VB.TextBox txtPatiMCNO 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   7200
               MaxLength       =   30
               TabIndex        =   21
               Top             =   1950
               Width           =   4295
            End
            Begin VB.ComboBox cbo病人类型 
               Height          =   300
               Left            =   10085
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   3435
               Width           =   1410
            End
            Begin VB.TextBox txt单位帐号 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   50
               TabIndex        =   49
               Top             =   4155
               Width           =   4295
            End
            Begin VB.TextBox txt备注 
               Height          =   300
               Left            =   7200
               MaxLength       =   100
               TabIndex        =   58
               Top             =   5235
               Visible         =   0   'False
               Width           =   4295
            End
            Begin VB.CommandButton cmdYB 
               Caption         =   "验证"
               Height          =   345
               Left            =   7155
               TabIndex        =   3
               Top             =   98
               Width           =   600
            End
            Begin VB.TextBox txt户口地址邮编 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   6
               TabIndex        =   34
               Top             =   3075
               Width           =   1485
            End
            Begin VB.TextBox txt联系人身份证 
               Height          =   300
               Left            =   7200
               MaxLength       =   18
               TabIndex        =   52
               Top             =   4515
               Width           =   4295
            End
            Begin VB.TextBox txt支付密码 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   20
               PasswordChar    =   "*"
               TabIndex        =   24
               Top             =   2340
               Width           =   1485
            End
            Begin zlIDKind.IDKindNew IDKind 
               Height          =   300
               Left            =   435
               TabIndex        =   80
               ToolTipText     =   "快捷键F4"
               Top             =   495
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   529
               Appearance      =   2
               IDKindStr       =   $"frmPatient.frx":0E42
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   9
               FontName        =   "宋体"
               IDKind          =   -1
               DefaultCardType =   "0"
               BackColor       =   -2147483633
            End
            Begin MSMask.MaskEdBox txt出生时间 
               Height          =   300
               Left            =   2145
               TabIndex        =   8
               Top             =   855
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txt出生日期 
               Height          =   300
               Left            =   1110
               TabIndex        =   7
               Top             =   855
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Format          =   "YYYY-MM-DD"
               Mask            =   "####-##-##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txt身份证号 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   1110
               TabIndex        =   13
               Top             =   1215
               Width           =   2160
            End
            Begin VB.TextBox txt其他证件 
               Height          =   300
               Left            =   1110
               MaxLength       =   20
               TabIndex        =   18
               Top             =   1560
               Width           =   3480
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   300
               Index           =   2
               Left            =   9395
               TabIndex        =   37
               Tag             =   "籍贯"
               Top             =   3075
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Items           =   2
               MaxLength       =   100
            End
            Begin VB.TextBox txt户口地址 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   31
               Top             =   3075
               Width           =   4785
            End
            Begin VB.TextBox txt出生地点 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   38
               Top             =   3435
               Width           =   4785
            End
            Begin VB.TextBox txt家庭地址 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   26
               Top             =   2715
               Width           =   4785
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   300
               Index           =   1
               Left            =   1110
               TabIndex        =   40
               Tag             =   "出生地点"
               Top             =   3435
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Items           =   3
               MaxLength       =   100
            End
            Begin VB.TextBox txt联系人地址 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   53
               Top             =   4875
               Width           =   4785
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   300
               Index           =   3
               Left            =   1110
               TabIndex        =   28
               Tag             =   "现住址"
               Top             =   2715
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   100
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   300
               Index           =   4
               Left            =   1110
               TabIndex        =   33
               Tag             =   "户口地址"
               Top             =   3075
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   100
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   300
               Index           =   5
               Left            =   1110
               TabIndex        =   55
               Tag             =   "联系人地址"
               Top             =   4875
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   100
            End
            Begin VB.Label lblMobile 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "手机号"
               Height          =   180
               Left            =   540
               TabIndex        =   169
               Top             =   5295
               Width           =   540
            End
            Begin VB.Label lbl验证密码 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "验证密码"
               Height          =   180
               Left            =   9270
               TabIndex        =   167
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lbl支付密码 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "支付密码"
               Height          =   180
               Left            =   6420
               TabIndex        =   166
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lbl学历 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "学历"
               Height          =   180
               Left            =   7500
               TabIndex        =   165
               Top             =   930
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付费方式"
               Height          =   180
               Left            =   3300
               TabIndex        =   164
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lblPatiMCNO 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "验证医保号"
               Height          =   180
               Index           =   1
               Left            =   6240
               TabIndex        =   163
               Top             =   2010
               Width           =   900
            End
            Begin VB.Label lbl职业 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "职业"
               Height          =   180
               Left            =   7500
               TabIndex        =   162
               Top             =   1290
               Width           =   360
            End
            Begin VB.Label lbl区域 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "区域"
               Height          =   180
               Left            =   6780
               TabIndex        =   119
               Top             =   3495
               Width           =   360
            End
            Begin VB.Label lblPatiMCNO 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医保号"
               Height          =   180
               Index           =   0
               Left            =   540
               TabIndex        =   118
               Top             =   2010
               Width           =   540
            End
            Begin VB.Label lbl住院号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号"
               Height          =   180
               Left            =   5070
               TabIndex        =   117
               Top             =   180
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lbl门诊号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊号"
               Height          =   180
               Left            =   2490
               TabIndex        =   116
               Top             =   180
               Width           =   540
            End
            Begin VB.Label lbl单位帐号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位帐号"
               Height          =   180
               Left            =   6420
               TabIndex        =   115
               Top             =   4215
               Width           =   720
            End
            Begin VB.Label lbl单位开户行 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位开户行"
               Height          =   180
               Left            =   180
               TabIndex        =   114
               Top             =   4215
               Width           =   900
            End
            Begin VB.Label lbl单位邮编 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位邮编"
               Height          =   180
               Left            =   9270
               TabIndex        =   113
               Top             =   3855
               Width           =   720
            End
            Begin VB.Label lbl单位电话 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位电话"
               Height          =   180
               Left            =   6420
               TabIndex        =   112
               Top             =   3855
               Width           =   720
            End
            Begin VB.Label lbl工作单位 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "工作单位"
               Height          =   180
               Left            =   360
               TabIndex        =   111
               Top             =   3855
               Width           =   720
            End
            Begin VB.Label lbl联系人电话 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人电话"
               Height          =   180
               Left            =   3120
               TabIndex        =   110
               Top             =   4575
               Width           =   900
            End
            Begin VB.Label lbl联系人地址 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人地址"
               Height          =   180
               Left            =   180
               TabIndex        =   109
               Top             =   4935
               Width           =   900
            End
            Begin VB.Label lbl联系人关系 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人关系"
               Height          =   180
               Left            =   6240
               TabIndex        =   108
               Top             =   4935
               Width           =   900
            End
            Begin VB.Label lbl联系人姓名 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人姓名"
               Height          =   180
               Left            =   180
               TabIndex        =   107
               Top             =   4575
               Width           =   900
            End
            Begin VB.Label lbl家庭地址邮编 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭地址邮编"
               Height          =   180
               Left            =   8910
               TabIndex        =   106
               Top             =   2775
               Width           =   1080
            End
            Begin VB.Label lbl家庭电话 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭电话"
               Height          =   180
               Left            =   6420
               TabIndex        =   105
               Top             =   2775
               Width           =   720
            End
            Begin VB.Label lbl家庭地址 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "现住址"
               Height          =   180
               Left            =   540
               TabIndex        =   104
               Top             =   2775
               Width           =   540
            End
            Begin VB.Label lbl婚姻状况 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "婚姻"
               Height          =   180
               Left            =   5250
               TabIndex        =   103
               Top             =   1620
               Width           =   360
            End
            Begin VB.Label lbl国籍 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "国籍"
               Height          =   180
               Left            =   5250
               TabIndex        =   102
               Top             =   1290
               Width           =   360
            End
            Begin VB.Label lbl民族 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "民族"
               Height          =   180
               Left            =   5250
               TabIndex        =   101
               Top             =   930
               Width           =   360
            End
            Begin VB.Label lbl身份 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份"
               Height          =   180
               Left            =   7500
               TabIndex        =   100
               Top             =   555
               Width           =   360
            End
            Begin VB.Label lbl身份证号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证号"
               Height          =   180
               Left            =   360
               TabIndex        =   99
               Top             =   1275
               Width           =   720
            End
            Begin VB.Label lbl出生地点 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生地点"
               Height          =   180
               Left            =   360
               TabIndex        =   98
               Top             =   3495
               Width           =   720
            End
            Begin VB.Label lbl出生日期 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生日期"
               Height          =   180
               Left            =   360
               TabIndex        =   97
               Top             =   915
               Width           =   720
            End
            Begin VB.Label lbl费别 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "费别"
               Height          =   180
               Left            =   720
               TabIndex        =   96
               Top             =   2385
               Width           =   360
            End
            Begin VB.Label lbl年龄 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄"
               Height          =   180
               Left            =   2790
               TabIndex        =   95
               Top             =   915
               Width           =   360
            End
            Begin VB.Label lbl性别 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别"
               Height          =   180
               Left            =   5250
               TabIndex        =   94
               Top             =   555
               Width           =   360
            End
            Begin VB.Label lbl姓名 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               Height          =   180
               Left            =   30
               TabIndex        =   93
               Top             =   555
               Width           =   360
            End
            Begin VB.Label lbl病人ID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病人ID"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   540
               TabIndex        =   92
               Top             =   180
               Width           =   540
            End
            Begin VB.Label lbl其他证件 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他证件"
               Height          =   180
               Left            =   360
               TabIndex        =   91
               Top             =   1620
               Width           =   720
            End
            Begin VB.Label lblPatiColor 
               Height          =   255
               Left            =   11270
               TabIndex        =   90
               Top             =   3440
               Width           =   210
            End
            Begin VB.Label lblPatiType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病人类型"
               Height          =   180
               Left            =   9270
               TabIndex        =   88
               Top             =   3495
               Width           =   720
            End
            Begin VB.Label lbl备注 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "备注"
               Height          =   180
               Left            =   6780
               TabIndex        =   85
               Top             =   5295
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label lbl籍贯 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "籍贯"
               Height          =   180
               Left            =   8910
               TabIndex        =   168
               Top             =   3135
               Width           =   360
            End
            Begin VB.Label lbl户口地址邮编 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址邮编"
               Height          =   180
               Left            =   6060
               TabIndex        =   83
               Top             =   3135
               Width           =   1080
            End
            Begin VB.Label lbl户口地址 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址"
               Height          =   180
               Left            =   360
               TabIndex        =   82
               Top             =   3135
               Width           =   720
            End
            Begin VB.Label lbl联系人身份证 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人身份证"
               Height          =   180
               Left            =   6060
               TabIndex        =   81
               Top             =   4575
               Width           =   1080
            End
         End
      End
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "医疗卡(&2)"
      Height          =   350
      Index           =   1
      Left            =   2760
      TabIndex        =   76
      ToolTipText     =   "补发就诊卡"
      Top             =   8460
      Width           =   1100
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "预交款(&1)"
      Height          =   350
      Index           =   0
      Left            =   1440
      TabIndex        =   74
      ToolTipText     =   "补交病人预交款"
      Top             =   8460
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   135
      TabIndex        =   75
      Top             =   8475
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10455
      TabIndex        =   73
      Top             =   8460
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9240
      TabIndex        =   72
      Top             =   8460
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   555
      Left            =   0
      TabIndex        =   135
      Top             =   0
      Width           =   1350
      _Version        =   589884
      _ExtentX        =   2381
      _ExtentY        =   979
      _StockProps     =   64
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   165
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   84
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit '要求变量声明
Public mlngModul As String
Public mstrPrivs As String
Public mbytInState As Byte '入：0=新增,1=修改,2=查看
Public mbytView As Byte '入：0-所有,1-在院,2-出院,3-门诊
Public mlng病人ID As Long '要修改或查看的病人ID
Public mlng主页ID As Long
Private mlng预交领用ID As Long '预交款票据领用ID
Private mlng险类 As Long
Private mlngOutModeMC As Long '外挂式医保的险类
Private mblnUnLoad As Boolean
Private mblnICCard As Boolean 'IC卡发卡,要同时填写病人信息的IC卡字段
Private mblnChange As Boolean
Private mblnSel As Boolean
Private mblnCheckPatiCard As Boolean
Private mstrYBPati As String
Private mblnPrepayPrint As Boolean    '是否打印预交款
Private mstr采集图片 As String '采集图片本地保存路径
Private mlng图像操作 As Long '指明当前对病人图像操作的类型(1-文件 2-采集 3-清除)
Private mobjPublicPatient As Object
Private mstrPatiPlus    As String     '从表信息:信息名1:信息值1,信息名2:信息值2
Private mblnEMPI As Boolean             'T-找到EMPI病人；F-未找到EMPI病人
Private Enum OPT
    C0预交款 = 0
    C1就诊卡 = 1
End Enum
Private mlngPatientID As Long '新增时提取病人身份时才有
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'关于结算卡的的处理变量
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '安装了结算卡的的
    dbl刷卡总额 As Double
    bln卡结算 As Boolean '当前读取的单据是卡结算
End Type

Private mtySquareCard As Ty_SquareCard
Private mobjKeyboard As Object

'Private mobjSquareCard As Object
Private mblnClickSquareCtrl As Boolean
Private mFactProperty As Ty_FactProperty
Private mblnStartFactUseType As Boolean '是否启用的相关的门诊类别的
Private mbytPrepayType As Byte '0-门诊住院;1-门诊;2-住院
Private mblnNotClick As Boolean
Private mobjSquare As Object '医疗卡部件
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private Type Ty_CardProperty
       lng卡类别ID As Long
       str卡名称  As String
       lng卡号长度 As Long
       lng结算方式 As String
       bln自制卡 As Boolean
       bln严格控制 As Boolean
       lng领用ID As Long
       lng共用批次 As Long
       bln变价 As Boolean
       bln刷卡 As Boolean
       int密码长度 As Integer
       int密码长度限制 As Integer
       int密码规则 As Integer
       bln就诊卡 As Boolean
       str卡号密文 As String
       blnOneCard As Boolean '  '是否启用了一卡通接口,此模式下，票号严格管理，票号范围外的发卡或绑定卡不收费
       rs卡费 As ADODB.Recordset
       dbl应收金额 As Double
       dbl实收金额 As Double
       bln是否制卡 As Boolean
       bln是否发卡 As Boolean
       bln是否写卡 As Boolean
       bln是否院外发卡  As Boolean
       lng发卡性质 As Long '0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0 为题号:57326
       bln重复使用 As Boolean
       byt发卡控制 As Byte
End Type
Private mCurSendCard As Ty_CardProperty
Private mcolPrepayPayMode As Collection   '预交款支付方式
Private mcolCardPayMode As Collection   '就诊卡支付方式
Private Type Ty_PayMoney
    lng医疗卡类别ID As Long
    bln消费卡 As Boolean
    str结算方式 As String
    str名称 As String
    str刷卡卡号 As String
    str刷卡密码 As String
    strNO As String
    lngID As Long '预交ID
    lng结帐ID As Long
End Type
Private mCurPrepay As Ty_PayMoney
Private mCurCardPay As Ty_PayMoney
Private mbln是否扫描身份证 As Boolean '是否是执行的扫描身份证操作
Private mbln扫描身份证签约 As Boolean '根据参数设置中的“扫描身份证签约”来取值
Private marrAddress(0 To 4) As String  '结构化地址缺省值
'问题号 :56599
Private Type ty_PageHeight
    基本 As Long
    健康档案 As Long
    附加信息 As Long
End Type
Private mPageHeight As ty_PageHeight

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Enum EState
    E新增 = 0
    E修改 = 1
    E查阅 = 2
End Enum

Private mstrCboSplit As String
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const C_ColumHeader = "过敏药物,1,3000,1;过敏反映,4,3000,1;过敏药物ID,1,100,0" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_InoculateHeader = "接种日期,4,2100,1;接种名称,4,2100,1;接种日期,4,2100,1;接种名称,4,2100,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_LinkManColumHeader = "联系人姓名,4,2100,1;联系人关系,4,2100,1;联系人身份证号,4,2100,1;联系人电话,4,2100,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_OtherInfoColumHeader = "信息名,4,2288,1;信息值,4,2288,1;信息名,4,2287,1;信息值,4,2287,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
'Private Const C_血型 = "A型,B型,O型,AB型,不详"
Private Const C_BH = "阴,阳,不详,未查"
Private mdic医疗卡属性 As New Dictionary
Private mobjHealthCard As Object '制卡接口对象
Private mbln发卡或绑定卡 As Boolean '标识是否进行了发卡或绑定卡操作
Private mbln基本  As Boolean '标识当前选中页
Private mlngPlugInHwnd As Long
'Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:创建或关闭结算卡对象
'    '入参:blnClosed:关闭对象
'    '编制:刘兴洪
'    '日期:2010-01-05 14:51:23
'    '问题:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strExpend As String
'    '0=新增,1=修改,2=查看
'   If mbytInState = E查阅 Then Exit Sub
'
'    '只有:执行或退费时,才可能管结算卡的
'    If blnClosed Then
'       If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.CloseWindows

'            Set mobjSquareCard = Nothing
'        End If
'        Exit Sub
'    End If
'
'    '创建对象
'    '刘兴洪:增加结算卡的结算:执行或退费时
'    Err = 0: On Error Resume Next
'    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
'    If Err <> 0 Then
'        Err = 0: On Error GoTo 0:      Exit Sub
'    End If
'
'    '安装了结算卡的部件
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    '功能:zlInitComponents (初始化接口部件)
'    '    ByVal frmMain As Object, _
'    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
'    '        ByVal cnOracle As ADODB.Connection, _
'    '        Optional blnDeviceSet As Boolean = False, _
'    '        Optional strExpand As String
'    '出参:
'    '返回:   True:调用成功,False:调用失败
'    '编制:刘兴洪
'    '日期:2009-12-15 15:16:22
'    'HIS调用说明.
'    '   1.进入门诊收费时调用本接口
'    '   2.进入住院结帐时调用本接口
'    '   3.进入预交款时
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    If mobjSquareCard.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
'         '初始部件不成功,则作为不存在处理
'         Exit Sub
'    End If
'    '初始成功,则证明此窗口存在相关的结算卡
'     mtySquareCard.blnExistsObjects = True
'End Sub


Private Sub InitSendCardPreperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化刷卡属性
    '编制:刘兴洪
    '日期:2011-07-25 11:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strSQL As String, blnBoundCard As Boolean
    Dim rsTemp As ADODB.Recordset, str批次 As String, varData As Variant, i As Long
    Dim varTemp  As Variant, blnNotBind As Boolean
    '76824，李南春，2014/8/19，医疗卡类别处理
    lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, 0))
    If InStr(mstrPrivs, ";发卡事务;") = 0 Or lngCardTypeID = 0 Then '无发卡权限
NotCard:
        fraCard.Visible = False: cmdOperation(OPT.C1就诊卡).Visible = False
        Me.Height = Me.Height - fraCard.Height
        mPageHeight.基本 = Me.Height
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    '问题号:57326
    strSQL = "" & _
    "   Select Id, 编码, 名称, 短名, 前缀文本, 卡号长度, 缺省标志, 是否固定, 是否严格控制,nvl( 是否刷卡,0) as 是否刷卡, " & _
    "           nvl(是否自制,0) as 是否自制, nvl(是否存在帐户,0) as 是否存在帐户, " & _
    "           nvl(是否全退,0) as 是否全退,nvl(是否重复使用,0) as 是否重复使用 , " & _
    "           nvl(密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    "           nvl(是否退现,0) as 是否退现,部件, 备注, 特定项目, 结算方式, 是否启用, 卡号密文," & _
    "           nvl(是否制卡,0) as 是否制卡,nvl(是否发卡,0) as 是否发卡, nvl(是否写卡,0) as 是否写卡, " & _
    "           nvl(发卡性质,0) as 发卡性质,nvl(发卡控制,0) as 发卡控制 " & _
    "    From 医疗卡类别 A" & _
    "    Where nvl(是否启用,0)=1 And (ID=[1] or nvl(缺省标志,0)=1)" & _
    "    Order by 编码"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardTypeID)
    If rsTemp.EOF Then GoTo NotCard:
    If rsTemp.RecordCount >= 2 Then
        rsTemp.Filter = "ID=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0
    End If
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        With mCurSendCard
            .lng卡类别ID = Val(Nvl(rsTemp!ID))
            .str卡名称 = Nvl(rsTemp!名称)
            .lng卡号长度 = Val(Nvl(rsTemp!卡号长度))
            .lng结算方式 = Trim(Nvl(rsTemp!结算方式))
            .bln自制卡 = Val(Nvl(rsTemp!是否自制)) = 1
            .bln严格控制 = Val(Nvl(rsTemp!是否严格控制)) = 1
            .bln刷卡 = Val(Nvl(rsTemp!是否刷卡)) = 1
            .str卡号密文 = Nvl(rsTemp!卡号密文)
            .int密码长度 = Val(Nvl(rsTemp!密码长度))
            .int密码长度限制 = Val(Nvl(rsTemp!密码长度限制))
            .int密码规则 = Val(Nvl(rsTemp!密码规则))
            .bln就诊卡 = .str卡名称 = "就诊卡" And Val(Nvl(rsTemp!是否固定)) = 1
            '问题号:56599
            .bln是否制卡 = Val(Nvl(rsTemp!是否制卡)) = 1
            .bln是否发卡 = Val(Nvl(rsTemp!是否发卡)) = 1
            .bln是否写卡 = Val(Nvl(rsTemp!是否写卡)) = 1
            .bln重复使用 = Val(Nvl(rsTemp!是否重复使用)) = 1
            .bln是否院外发卡 = (InStr(mstrPrivs, ";发卡事务;") > 0 And .bln自制卡 = False And .bln是否发卡 = True) '问题号:56599
            .lng发卡性质 = Val(Nvl(rsTemp!发卡性质)) '问题号:57326
            .byt发卡控制 = Val(Nvl(rsTemp!发卡控制))
            '76824，李南春，2014/8/19，医疗卡类别处理
            lbl卡名称.Caption = .str卡名称
            lbl卡名称.Width = LenB(lbl卡名称.Caption) * 100
            .blnOneCard = False
            If Trim(Nvl(rsTemp!特定项目)) <> "" Then
                Set .rs卡费 = GetSpecialInfo(Trim(Nvl(rsTemp!特定项目)))
                If .bln就诊卡 Then .blnOneCard = GetOneCard.RecordCount > 0
                '142204:李南春,2020/6/16,修改发卡无应收金额
                If Not .rs卡费.EOF Then
                    .bln变价 = Val(Nvl(.rs卡费!是否变价)) = 1
                    .dbl应收金额 = Format(IIf(.bln变价, .rs卡费!缺省价格, .rs卡费!现价), "0.00")
                End If
            Else
                Set .rs卡费 = Nothing
            End If
            str批次 = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModul, "0")
            '领用ID,卡类别ID|...
             .lng共用批次 = 0
            varData = Split(str批次, "|")
            For i = 0 To UBound(varData)
                 varTemp = Split(varData(i), ",")
                 If Val(varTemp(0)) <> 0 Then
                    If ExistShareBill(Val(varTemp(0)), 5) Then
                        If Val(varTemp(1)) = .lng卡类别ID Then
                            .lng共用批次 = Val(varTemp(0)): Exit For
                        End If
                    End If
                 End If
            Next
           txt卡号.PasswordChar = IIf(.str卡号密文 <> "", "*", "")
           txt卡号.MaxLength = .lng卡号长度
        End With
    End If

    If mCurSendCard.rs卡费 Is Nothing Then
    
        cmdOperation(OPT.C1就诊卡).Visible = False
        tabCardMode.Tabs.Remove ("CardFee")
        blnBoundCard = InStr(mstrPrivs, ";绑定卡号;") > 0
        '无绑定卡权限
          fraCard.Visible = blnBoundCard: cmdOperation(OPT.C1就诊卡).Visible = blnBoundCard
        If Not blnBoundCard Then
            Me.Height = Me.Height - fraCard.Height
            mPageHeight.基本 = Me.Height
        Else
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "绑定卡号"
            tabCardMode.Width = tabCardMode.Width / 2
        End If
        Exit Sub
    End If
     
    
    With mCurSendCard.rs卡费
        txt卡额.Text = Format(IIf(Val(Nvl(!是否变价)) = 1, Val(Nvl(!缺省价格)), Val(Nvl(!现价))), "0.00")
        txt卡额.Tag = txt卡额.Text  '保持不变
        txt卡额.Locked = Not (Val(Nvl(!是否变价)) = 1)
        txt卡额.TabStop = (Val(Nvl(!是否变价)) = 1)
    End With
     
     
    '自制卡,在卡号不重复使用 或者严格控制时,不能进行绑定卡操作
    blnNotBind = mCurSendCard.bln自制卡 And (Not mCurSendCard.bln重复使用 Or mCurSendCard.bln严格控制)
    
    '如果没有绑定卡权限,加载窗体时,已经移除了绑定卡号
    blnBoundCard = Not InStr(mstrPrivs, ";绑定卡号;") > 0
    If Not blnBoundCard Then
        If zlDatabase.GetPara("发卡模式", glngSys, mlngModul, "CardFee") = "CardFee" Then
            tabCardMode.Tabs("CardFee").Selected = True
        ElseIf Not blnNotBind Then
            tabCardMode.Tabs("CardBind").Selected = True
        End If
    End If
    
    '问题号:56599
    If (mCurSendCard.bln是否院外发卡 Or blnNotBind) And Not blnBoundCard Then
       '1.如果院外卡进行发卡 2.院内卡.严格控制或者不重复利用   以上这2种情况但是同时拥有绑定卡权限 都不能进行绑定卡操作,无绑定卡权限,在窗体加载时,便删除了绑定卡
        tabCardMode.Tabs.Remove ("CardBind")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardFee").Selected = True
            tabCardMode.Tabs("CardFee").Caption = "收费发卡"
            tabCardMode.Width = tabCardMode.Width / 2
        Else
            fraCard.Visible = False
            Me.Height = Me.Height - fraCard.Height
            mPageHeight.基本 = Me.Height
        End If
    ElseIf mCurSendCard.bln自制卡 = False And mCurSendCard.bln是否发卡 = False Then
        tabCardMode.Tabs.Remove ("CardFee")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "绑定卡号"
            tabCardMode.Width = tabCardMode.Width / 2
        Else
            fraCard.Visible = False
            Me.Height = Me.Height - fraCard.Height
            mPageHeight.基本 = Me.Height
        End If
    End If
    
    If mCurSendCard.bln严格控制 Then
        '就诊卡领用检查
        mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), , mCurSendCard.lng卡类别ID)
        If mCurSendCard.lng领用ID <= 0 Then
            Select Case mCurSendCard.lng领用ID
                Case 0 '操作失败
                Case -1
'                    MsgBox "你没有自用或共用的就诊卡,不能发放！" & vbCrLf & _
'                        "请先在本地设置共用批次或领用一批新卡! ", vbExclamation, gstrSysName
                Case -2
'                    MsgBox "本地共用的就诊卡已用完,不能发放！" & vbCrLf & _
'                        "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
            End Select
            cmdOperation(OPT.C1就诊卡).Visible = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo病人类型_Click()
    If cbo病人类型.ListCount > 0 And cbo病人类型.ListIndex <> -1 Then
        lblPatiColor.BackColor = zlDatabase.GetPatiColor(NeedName(cbo病人类型.Text))
        txtPatient.ForeColor = lblPatiColor.BackColor
    End If
End Sub
Private Sub cbo结算方式_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    With mCurCardPay
            .lng医疗卡类别ID = 0
            .bln消费卡 = False
            .str结算方式 = ""
            .str名称 = ""
     End With
    '0=新增,1=修改,2=查看
    If mbytInState = E查阅 Then Exit Sub
    Call SetCardVaribles(False)
    '130245,结算方式切换，同步更新卡类别ID
    If mblnNotClick = True Then Exit Sub
    Call Local结算方式(mCurCardPay.lng医疗卡类别ID, True)
End Sub

Private Sub cbo联系人关系_Click()
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人关系") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人关系")) = NeedName(cbo联系人关系.Text)
    End If
End Sub

Private Sub cbo性别_Change()
    Call ReLoadCardFee
End Sub

Private Sub cbo预交结算_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    
    With mCurPrepay
            .lng医疗卡类别ID = 0
            .bln消费卡 = False
            .str结算方式 = ""
            .str名称 = ""
     End With
    '0=新增,1=修改,2=查看
    If mbytInState = E查阅 Then Exit Sub
    Call SetCardVaribles(True)
    '130245,结算方式切换，同步更新卡类别ID
    If mblnNotClick = True Then Exit Sub
    Call Local结算方式(mCurPrepay.lng医疗卡类别ID, False)
End Sub

Private Sub cmdPicClear_Click()
    '问题号:74421
    imgPatient.Picture = Nothing
    mlng图像操作 = 3
End Sub

Private Sub cmdPicCollect_Click()
    If mobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set mobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPublicPatient Is Nothing Then
        MsgBox "创建病人信息公共部件(zlPublicPatient.clsPublicPatient)失败!", vbInformation, gstrSysName
        Exit Sub
    End If
    If mobjPublicPatient.PatiImageGatherer(Me, mstr采集图片) = False Then Exit Sub
    Set imgPatient.Picture = LoadPicture(mstr采集图片)
    mlng图像操作 = 2
End Sub

Private Sub cmdPicFile_Click()
    '问题号:74421
    Dim strFileDir As String
    On Error GoTo ErrHand:
    With cmdialog
        .CancelError = False
        .flags = cdlOFNHideReadOnly
        .Filter = "(*.bmp)|*.bmp"
        .FilterIndex = 2
        .ShowOpen
        strFileDir = .FileName
        If strFileDir = "" Then Exit Sub
        imgPatient.Picture = LoadPicture(strFileDir)
    End With
    mlng图像操作 = 1
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdYB_Click()
    Dim lng病人ID As Long, lng病种ID As Long
    Dim objCurrent As Object, strTxt As String, arrTxt As Variant
    Dim i As Long, blnDo As Boolean, arrPati As Variant
    Dim objcbo As ComboBox
    Dim strYBPati As String, strYBPatiBak As String
    Dim intInsure As Integer
    
    '医保改动
    lng病人ID = mlngPatientID
    strYBPati = gclsInsure.Identify(1, lng病人ID, intInsure, 1)
    mstrYBPati = strYBPati
    If strYBPati <> "" Then
        arrPati = Split(strYBPati, ";")
        '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,...
        If UBound(arrPati) >= 8 Then
            If Val(arrPati(8)) > 0 Then
               txtPatient.Text = "-" & Val(arrPati(8))
                blnDo = txtPatient.Locked
                txtPatient.Locked = False
                Call txtPatient_KeyPress(13)
                txtPatient.Locked = blnDo
                If strYBPati = "" Then txtPatient.SetFocus: Exit Sub  '可能因为余额不足提醒选择了退出等,调用了clearcard
            End If
        End If
        
        
        '医保号
        txtPatiMCNO(0).Text = arrPati(1)
        txtPatiMCNO(0).Locked = True
        
        '姓名
        txtPatient.Text = arrPati(3)
        
        '性别
        cbo性别.ListIndex = GetCboIndex(cbo性别, CStr(arrPati(4)))
        
        '出生日期
        If IsDate(arrPati(5)) Then
            txt出生日期.Text = Format(arrPati(5), "yyyy-MM-dd")
            Call txt出生日期_LostFocus
        End If
        
        '身份证号
        txt身份证号.Text = arrPati(6)
        
        '工作单位
        txt工作单位.Text = arrPati(7)
        
        If txt门诊号.Text = "" Then txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
        
        If cbo国籍.ListIndex = -1 Then Call ReadDict("国籍", cbo国籍)
        If cbo民族.ListIndex = -1 Then Call ReadDict("民族", cbo民族)
        If cbo学历.ListIndex = -1 Then Call ReadDict("学历", cbo学历)
        If cbo婚姻状况.ListIndex = -1 Then Call ReadDict("婚姻状况", cbo婚姻状况)
        If cbo职业.ListIndex = -1 Then Call ReadDict("职业", cbo职业)
        If cbo身份.ListIndex = -1 Then Call ReadDict("身份", cbo身份)
        
        '新增时病人类型不可见
        'lblPatiType.Visible = False: cbo病人类型.Visible = False: lblPatiColor.Visible = False
       
        If Not IsDate(txt出生日期.Text) Then
            txt出生日期.SetFocus
        Else
            strTxt = "txt年龄,cbo性别,cbo费别,cbo国藉,cbo民族,cbo学历,cbo婚姻状况,cbo职业,cbo身份," & _
                     "txt身份证号,txt出生地点,txt家庭地址,txt家庭地址邮编,txt家庭电话,txt工作单位,txt单位电话,txt单位邮编," & _
                     "txt单位开户行,txt单位帐号,txt联系人姓名,cbo联系人关系,txt联系人地址,txt联系人电话,txt联系人身份证"
            arrTxt = Split(strTxt, ",")
            i = 0
            For i = 0 To UBound(arrTxt)
                For Each objCurrent In Me.Controls
                    If objCurrent.Name = arrTxt(i) Then
                        blnDo = False
                        If TypeOf objCurrent Is TextBox Then
                            If Trim(objCurrent.Text) = "" And objCurrent.Enabled = True Then blnDo = True
                        ElseIf TypeOf objCurrent Is ComboBox Then
                            Set objcbo = objCurrent
                            If objcbo.ListIndex = -1 And objCurrent.Enabled = True Then blnDo = True
                        End If
                        If blnDo Then
                            If objCurrent.TabStop Then
                                If objCurrent.Visible Then objCurrent.SetFocus
                                Exit Sub
                            End If
                        End If
                        GoTo exitHandle
                    End If
                Next
exitHandle:
            Next
        End If
        txtPatient.SetFocus
    Else
        txtPatient.SetFocus
    End If
End Sub

Private Sub cmd户口地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
            " Substr(名称,1,2) as 名称 From 地区" & _
            " Union All" & _
            " Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
            " From 地区 Order by 编码", 2, "地区", , txt出生地点.Text)
    If Not rsTmp Is Nothing Then
        txt户口地址.Text = rsTmp!名称
        txt户口地址.SelStart = Len(txt户口地址.Text)
        txt户口地址.SetFocus
    End If
End Sub

Private Sub cmd籍贯_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt籍贯, True)
    If Not rsTmp Is Nothing Then
        txt籍贯.Text = rsTmp!名称
        txt籍贯.SelStart = Len(txt籍贯.Text)
        txt籍贯.SetFocus
    Else
        SelAll txt籍贯
        txt籍贯.SetFocus
    End If
End Sub

Private Sub cmd区域_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt区域, True)
    If Not rsTmp Is Nothing Then
        txt区域.Text = rsTmp!名称
        txt区域.SelStart = Len(txt区域.Text)
        txt区域.SetFocus
    Else
        SelAll txt区域
        txt区域.SetFocus
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXml As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
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
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    '问题号:56599
    If strOutPatiInforXml <> "" Then LoadPati strOutPatiInforXml
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    
    Set gobjSquare.objCurCard = objCard
    '是否密文显示
    'txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55571:刘鹏飞,2012-011-12
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then
        txtPatient.Text = ""
        '69200:刘鹏飞,2013-12-31,新增提取现有病人,切换输入方式表示要开始录入新病人。
        If mbytInState = E新增 And mlngPatientID <> 0 Then
            Call ClearCard
            mblnICCard = False
            txt病人ID.Text = zlDatabase.GetNextNo(1): lbl病人ID.Tag = txt病人ID.Text
            txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
        End If
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub lbl就诊卡号_Click()
    Dim strExpand As String, strOutCardNO As String, strOutPatiInforXml As String

    If mCurSendCard.bln就诊卡 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        
        If Not mobjICCard Is Nothing Then
            txt卡号.Text = mobjICCard.Read_Card()
            If txt卡号.Text <> "" Then
                mblnICCard = True
                Call CheckFreeCard(txt卡号.Text)
            End If
        End If
        Exit Sub
    End If
    If mCurSendCard.bln刷卡 Or mCurSendCard.lng卡类别ID = 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\

    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, mCurSendCard.lng卡类别ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txt卡号.Text = strOutCardNO
    If txt卡号.Text <> "" Then
        '问题号:56599
        If strOutPatiInforXml <> "" Then Call LoadPati(strOutPatiInforXml)
        Call CheckFreeCard(txt卡号.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt卡号.Text = strCardNo
    If txt卡号.Text <> "" Then
        '问题号:56599
        If strXmlCardInfor <> "" Then Call LoadPati(strXmlCardInfor)
        Call CheckFreeCard(txt卡号.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim objCard As Card
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
        If objCard Is Nothing Then Exit Sub
        txtPatient.Text = strCardNo
        Call FindPati(objCard, True, strCardNo)
        
        If txtPatient.Text <> "" Then
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
        End If
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    
    Dim lngIndex As Long, lngPatientID As Long
    Dim objCard As Card
    Dim bln签约 As Boolean
    Dim strErrMsg As String
    Dim str年龄 As String
    '57945:刘鹏飞,2013-10-30,读取身份证中的地址应该放到户口地址而不是家庭地址
    '55218:刘鹏飞,2012-10-25
'    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
'
'        txtPatient.Text = strName
'        Call zlControl.CboLocate(cbo性别, strSex)
'        Call zlControl.CboLocate(cbo民族, strNation)
'        txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
'        txt出生时间.Text = "00:00"
'        txt户口地址.Text = strAddress
'        txt身份证号.Text = strID
'    End If
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
        If objCard Is Nothing Then Exit Sub
        txtPatient.Text = strID
        Call FindPati(objCard, False, strID, lngPatientID)
        '刷身份证时校验姓名
        If strName <> Trim(txtPatient.Text) And strName <> "" And mlngPatientID <> 0 Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIn(mlngPatientID) Then
                    '正在就诊
                    MsgBox "病人原来的姓名【" & Trim(txtPatient.Text) & "】与身份证上面的姓名【" & strName & "】不一致，请先修改成正确的姓名后再进行登记。", vbInformation, gstrSysName
                    Call ClearCard:  Exit Sub
                Else
                    MsgBox "病人原来的姓名【" & Trim(txtPatient.Text) & "】与身份证上面的姓名【" & strName & "】不一致，系统将更新为身份证上面的姓名。", vbInformation, gstrSysName
                    str年龄 = Trim(txt年龄.Text)
                    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
                    If gobjPublicPatient.SavePatiBaseInfo(mlngPatientID, 0, strName, strSex, str年龄, Format(datBirthDay, "yyyy-MM-dd"), Me.Caption, 1, "", True, True) Then
                        txtPatient.Text = "-" & mlngPatientID
                        txtPatient.SetFocus
                        Call txtPatient_KeyPress(vbKeyReturn)
                    End If
                End If
            End If
        End If
        mbln是否扫描身份证 = False
        If (mCurSendCard.str卡名称 = "二代身份证" Or mbln扫描身份证签约) Then bln签约 = 是否已经签约(Trim(strID))
        If lngPatientID <> 0 And Not bln签约 And (mCurSendCard.str卡名称 = "二代身份证" Or mbln扫描身份证签约) Then
            '现有病人，身份证没签约,检查身份证信息信息是否和身份证卡片上的信息一致 2012-10-26 lgf
            If Trim(txtPatient.Text) <> Trim(strName) Or NeedName(cbo性别.Text) <> strSex Or Format(txt出生日期.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                      If Trim(txtPatient.Text) <> Trim(strName) Then
                           strErrMsg = strErrMsg & "," & "姓名"
                      End If
                      If NeedName(Me.cbo性别.Text) <> strSex Then
                           strErrMsg = strErrMsg & "," & "性别"
                      End If
                      If Format(txt出生日期.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                          strErrMsg = strErrMsg & "," & "出生日期"
                      End If
                      strErrMsg = Mid(strErrMsg, 2)
                      strErrMsg = "当前病人信息与身份证上的[" & strErrMsg & "]等信息不一致!" & vbCrLf & "不能进行身份证签约操作!"
                      Call MsgBox(strErrMsg, vbQuestion, Me.Caption)
                       mbln是否扫描身份证 = False
            Else
                 mbln是否扫描身份证 = True
            End If
        End If
        
        If lngPatientID = 0 Then '新病人
            lngIndex = IDKind.GetKindIndex("姓名")
            If lngIndex >= 0 Then IDKind.IDKind = lngIndex
            txtPatient.Text = "": txtPatient.PasswordChar = ""
            '55571:刘鹏飞,2012-011-12
            txtPatient.IMEMode = 0
            txtPatient.Text = strName
            Call zlControl.CboLocate(cbo性别, strSex)
            Call zlControl.CboLocate(cbo民族, strNation)
            txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt出生时间.Text = "00:00"
            txt身份证号.Text = strID
            '74421,刘鹏飞,2014-07-04,读取病人照片信息
            Call LoadIDImage
            mbln是否扫描身份证 = Not bln签约
        End If
        '101692新病人直接提取；已经建档病人当户口地址为空时自动更新
        If lngPatientID = 0 Or (lngPatientID <> 0 And Trim(txt户口地址.Text) = "") Then
            txt户口地址.Text = strAddress
            If gbln启用结构化地址 Then
                PatiAddress(E_IX_户口地址).Value = strAddress
            End If
        End If
        mblnNotClick = False
    End If
'   55240 2012-10-26 lgf
'    '问题号:53408
'    mbln是否扫描身份证 = False
'    If mbln扫描身份证签约 Then
'         mbln是否扫描身份证 = Not 是否已经签约(strID)
'    End If
''    If mCurSendCard.str卡名称 = "二代身份证" And Me.ActiveControl Is txt卡号 Then
'
'        If txtPatient.Text <> "" And cbo性别.ListCount <> 0 And txt出生日期.Text <> "" Then
'            If strName <> txtPatient.Text Or strSex <> Split(cbo性别.Text, "-")(1) Or txt出生日期.Text <> Format(datBirthDay, "yyyy-MM-dd") Then
'                    MsgBox "身份证信息与挂号病人信息不一致,不能进行签约操作！", vbInformation, gstrSysName
'                    Exit Sub
'            End If
'        Else
'             MsgBox "绑定二代身份证时,病人信息不允许为空！", vbInformation, gstrSysName
'             Exit Sub
'        End If
'
'        If 是否已经签约(Trim(strID)) Then
'            MsgBox "身份证号码为:" & strID & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
'            txt卡号.SetFocus
'            Exit Sub
'        Else
'            txt身份证号.Text = strID
'            txt卡号.Text = strID
'            mbln是否扫描身份证 = True
'        End If
'
'    End If
    If Me.ActiveControl Is txt身份证号 Then
        
        If txtPatient.Text <> "" And cbo性别.ListCount <> 0 And txt出生日期.Text <> "" Then
            If strName <> txtPatient.Text Or strSex <> Split(cbo性别.Text, "-")(1) Or txt出生日期.Text <> Format(datBirthDay, "yyyy-MM-dd") Then
                    MsgBox "身份证信息与病人信息不一致,不能进行签约操作！", vbInformation, gstrSysName
                    Exit Sub
            End If
        Else
             MsgBox "绑定二代身份证时,病人信息不允许为空！", vbInformation, gstrSysName
             Exit Sub
        End If
        
        If 是否已经签约(Trim(strID)) Then
            MsgBox "身份证号码为:" & strID & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
            txt身份证号.SetFocus
            Exit Sub
        Else
            txt身份证号.Text = strID
            mbln是否扫描身份证 = True
        End If
        
    End If
    
    Call Show绑定控件(mbln是否扫描身份证 And mbln扫描身份证签约)
End Sub

Private Sub cbo年龄单位_LostFocus()
    '68489:刘鹏飞,2013-12-06,没有输入年龄则不反算出生日期
    If Trim(txt年龄.Text) = "" Then Exit Sub
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Sub
    
    If Not IsDate(txt出生日期.Text) Then
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    Call ReLoadCardFee
End Sub

Private Sub cbo预交结算_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo预交结算.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo预交结算.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo预交结算.ListIndex = lngIdx
End Sub

Private Sub chk记帐_Click()
    If chk记帐.Value = Checked Then
        cbo结算方式.Enabled = False
        If Visible Then cmdOK.SetFocus
    Else
        cbo结算方式.Enabled = True
        cbo结算方式.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    If mbytInState = E新增 And mlngPatientID <> 0 Then
        If MsgBox("你确定要清除当前病人信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ClearCard
            mblnICCard = False  '不能放在clearcard中,因为可能先读卡再查出病人
            '问题27207 by lesfeng 2010-1-4
            txt病人ID.Text = zlDatabase.GetNextNo(1): lbl病人ID.Tag = txt病人ID.Text
            txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
        End If
    ElseIf mbytInState = E新增 And gblnOK Then
        If txtPatient.Text <> "" Then
            If glngSys Like "8??" Then
                If MsgBox("当前客户信息尚未保存,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("当前病人信息尚未保存,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        Else
            If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function IsCheck就诊卡() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的输入是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-27 10:21:41
    '问题:25302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCard As String, strICCard As String
    strCard = UCase(txt卡号.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    '-----------------------------------------------------------------------------------------------------------------
    '1.就诊卡的检查
    '变价金额检查
    '刘兴洪:And tabCardMode.SelectedItem.Key = "CardFee"
    '29134
    '82401:李南春,2015/3/11,判断对象是否存在
    If mbytInState = E新增 And fraCard.Visible = True Then
        If Trim(txt卡号.Text) <> "" And tabCardMode.SelectedItem.Key = "CardFee" Then
            If Not mCurSendCard.rs卡费 Is Nothing Then
                If mCurSendCard.rs卡费!是否变价 = 1 Then
                    If mCurSendCard.rs卡费!现价 <> 0 And Abs(CCur(txt卡额.Text)) > Abs(mCurSendCard.rs卡费!现价) Then
                        MsgBox IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡金额绝对值不能大于最高限价：" & Format(Abs(mCurSendCard.rs卡费!现价), "0.00"), vbExclamation, gstrSysName
                        If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus:  Exit Function
                    End If
                    If mCurSendCard.rs卡费!原价 <> 0 And Abs(CCur(txt卡额.Text)) < Abs(mCurSendCard.rs卡费!原价) Then
                        MsgBox IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡金额绝对值不能小于最低限价：" & Format(Abs(mCurSendCard.rs卡费!原价), "0.00"), vbExclamation, gstrSysName
                        If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus: Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    If fraCard.Visible = True Then
        If tabCardMode.SelectedItem.Key = "CardFee" Then
            If cbo结算方式.Visible And txt卡号.Text <> "" And cbo结算方式.Enabled And cbo结算方式.ListIndex = -1 Then
                MsgBox "请确定" & IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡的缴款结算方式！", vbExclamation, gstrSysName
                If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus: Exit Function
            End If
        End If
    End If
    
    If txtPass.Text <> txtAudi.Text And fraCard.Visible = True And txt卡号.Text <> "" Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus: Exit Function
    End If
    
    If Trim(txt卡号.Text) = "" And txt卡号.Visible And mbytInState = E新增 And gblnMustCard Then
        MsgBox "请刷卡或输入" & IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡号！", vbExclamation, gstrSysName
        If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
        Exit Function
    End If
    If txt卡号.Text <> "" And mbytInState = E新增 Then
        '保存前检查就诊卡是否有，是否在范围内
        If mCurSendCard.bln严格控制 Then
            mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), txt卡号.Text, mCurSendCard.lng卡类别ID)
     
            If mCurSendCard.lng领用ID <= 0 And Not mCurSendCard.blnOneCard Then
                Select Case mCurSendCard.lng领用ID
                    Case 0 '操作失败
                    Case -1
'                        If txt卡号.Text <> "" Then MsgBox "你已没有自用及共用的" & IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡,不能发放！" & vbCrLf & _
'                            "请先在本地设置共用批次或领用一批新卡! ", vbExclamation, gstrSysName
                    Case -2
'                        If txt卡号.Text <> "" Then MsgBox "本地共用的" & IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡已用完,不能发放！" & vbCrLf & _
'                            "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
                    Case -3
                        MsgBox "该张卡号不在有效范围内,请检查是否正确刷卡！", vbExclamation, gstrSysName
                        If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
                End Select
                Exit Function
            End If
        End If
    End If
    '保存前,需要检查支付金额
    
    
    IsCheck就诊卡 = True
End Function
Private Sub SetCardEditEnabled()
    '设置就诊卡编辑属性
    Dim blnEdit As Boolean
    If Not (mbytInState = E新增 Or mbytInState = E修改) Then Exit Sub
    blnEdit = Trim(txt卡号.Text) <> ""
    
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    lbl密码.Enabled = txtPass.Enabled: lbl验证.Enabled = blnEdit
    
    txt卡额.Enabled = blnEdit: lbl金额.Enabled = blnEdit
    chk记帐.Enabled = blnEdit
    cbo结算方式.Enabled = chk记帐.Value = 0 And blnEdit
End Sub

Private Function CanFocus(ctlError As Control) As Boolean
    CanFocus = ctlError.Enabled And ctlError.Visible
End Function

Private Function IsValied(Optional blnModify As Boolean, Optional strBirthDay As String, Optional strAge As String, Optional strSex As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '返回:数据合法,返回true,否则返回False
    '   出参： blnModify =True时 病人出生日期和性别和年龄会根据身份证信息同步调整（与 基本信息调整 权限有关） =false 只保存身份证号,病人信息不同步做调整
    '          blnModify=True时 返回 strBirthday,strAge,strSex
    '编制:刘兴洪
    '日期:2011-07-26 16:40:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSimilar As String, i As Long, str称呼 As String, lngTmp As Long
    Dim str出生日期 As String, str年龄 As String, strInfo As String
    Dim blnMod As Boolean, bln基本信息调整 As Boolean
    Dim strMsg As String
    
    On Error GoTo errHandle
    
    str称呼 = IIf(glngSys Like "8??", "客户", "病人")
    
    '65965:刘鹏飞,2013-09-24,处理预交显示千位位格式
    If Not CheckFormInput(Me, "txt预交额") Then Exit Function
    
    '合法性检查
    If Not IsNumeric(txt门诊号.Text) And txt门诊号.Text <> "" Then
        MsgBox "请输入一个有效的门诊号！", vbInformation, gstrSysName
        If txt门诊号.Enabled And txt门诊号.Visible Then txt门诊号.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txt住院号.Text) And txt住院号.Text <> "" Then
        MsgBox "请输入一个有效的住院号！", vbInformation, gstrSysName
        If txt住院号.Enabled And txt住院号.Visible Then txt住院号.SetFocus: Exit Function
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text And txtPatiMCNO(1).Visible Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "请检查,医保号最大长度不能超过" & txtPatiMCNO(0).MaxLength & "个字符！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If

    If Trim(txtPatient.Text) = "" Then
        MsgBox "必须输入[姓名]！", vbExclamation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
    End If
    If cbo性别.ListIndex = -1 Then
        MsgBox "必须确定[性别]！", vbExclamation, gstrSysName
        If cbo性别.Enabled And cbo性别.Visible Then cbo性别.SetFocus: Exit Function
    End If
    If Not IsDate(txt出生日期.Text) Then
        MsgBox "必须正确输入[出生日期]！", vbInformation, gstrSysName
        If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus: Exit Function
    End If
    If Trim(txt年龄.Text) = "" Then
        MsgBox "必须输入[年龄]！", vbExclamation, gstrSysName
        If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus: Exit Function
    End If
    
    '76409,刘鹏飞,2014-08-06,年龄合法性检查
    If txt年龄.Locked = False Then
        str年龄 = txt年龄.Text
        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
        If IsDate(txt出生日期.Text) Then
            If txt出生时间.Text = "__:__" Then
                str出生日期 = Format(txt出生日期.Text, "YYYY-MM-DD")
            Else
                str出生日期 = Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS")
            End If
            strInfo = CheckAge(str年龄, str出生日期, CDate(txt出生日期.Tag))
        Else
            strInfo = CheckAge(str年龄)
        End If
        If InStr(1, strInfo, "|") > 0 Then
            lngTmp = Val(Split(strInfo, "|")(0)) '1禁止,0提示
            strInfo = Split(strInfo, "|")(1)
            If lngTmp = 1 Then
                MsgBox strInfo, vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus: Exit Function
            Else
                If MsgBox(strInfo & vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    str出生日期 = ""
    '--46119,刘鹏飞,2012-08-16,根据身份证对出生日期和年龄的检查
    '身份证长度检查
    '--81012,余伟节,2014-12-22,根据身份证对出生日期\年龄\性别 的同步调整
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) = "中国" Then
        If Not CheckLen(txt身份证号, 18) Then Exit Function
        lngTmp = LenB(StrConv(Trim(txt身份证号.Text), vbFromUnicode))
        If lngTmp > 0 Then
            If CreatePublicPatient() Then
                strInfo = ""
                If gobjPublicPatient.CheckPatiIdcard(Trim(txt身份证号.Text), strBirthDay, strAge, strSex, strInfo, CDate(txt出生日期.Tag)) Then
                    '同一身份证只能对应一个建档病人非新建档病人时,检查身份证号
                    If gblnPatiByID And Trim(txt身份证号.Text) <> txt身份证号.Tag And mlng病人ID <> 0 Then
                       If gobjPublicPatient.CheckPatiExistByID(Trim(txt身份证号.Text), mlng病人ID) Then
                            MsgBox "已存在身份证号为【" & Trim(txt身份证号.Text) & "】的建档病人，禁止修改身份证号！", vbInformation, gstrSysName
                            If CanFocus(txt身份证号) Then txt身份证号.SetFocus: Exit Function
                        End If
                    End If
                    '有无基本信息调整权限
                    bln基本信息调整 = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;") > 0 And ((mlngPatientID > 0 And mbytInState = 0) Or mbytInState = 1)
                    '出生日期
                    strMsg = ""
                    If CDate(Format(strBirthDay, "YYYY-MM-DD")) <> CDate(Format(txt出生日期.Text, "YYYY-MM-DD")) Then
                        strMsg = "身份证号码中出生日期[" & strBirthDay & "]与病人出生日期[" & Format(txt出生日期.Text, "YYYY-MM-DD") & "]不一致"
                        '年龄 带单位
                        str年龄 = txt年龄.Text
                        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
                        If str年龄 <> strAge Then
                            strMsg = strMsg & vbCrLf & "身份证号码中年龄[" & strAge & "]与病人年龄[" & str年龄 & "]不一致"
                            If str年龄 Like "*小时*分钟" Or str年龄 Like "*分钟" Or str年龄 Like "*天*小时" Or str年龄 Like "*小时" Then
                                strAge = str年龄
                            End If
                        End If
                    End If
                    If txt出生时间.Text <> "__:__" Then
                        strBirthDay = strBirthDay & " " & Format(txt出生时间.Text, "HH:MM")
                    End If
                    '性别
                    If InStr(cbo性别.Text, strSex) = 0 Then
                        strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & "身份证号码中性别[" & strSex & "]与病人性别[" & NeedName(cbo性别.Text) & "]不一致"
                    End If
                    
                    If ((mlngPatientID > 0 And mbytInState = 0) Or mbytInState = 1) Then
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",是否继续？" & vbCrLf & IIf(bln基本信息调整, "选【是】,用身份证的信息替换病人的信息及相关业务数据。", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Function
                            Else
                                blnMod = True
                            End If
                        End If
                    Else
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",是否继续？" & vbCrLf & "选【是】,用身份证的信息替换病人的信息。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Function
                            Else
                                If CDate(Format(strBirthDay, "YYYY-MM-DD")) <> CDate(Format(txt出生日期.Text, "YYYY-MM-DD")) Then
                                    txt出生日期.Text = strBirthDay
                                    If mblnChange = False Then
                                        Call LoadOldData(strAge, txt年龄, cbo年龄单位)
                                    End If
                                End If
                                Call zlControl.CboLocate(cbo性别, strSex, False)
                            End If
                        End If
                    End If
                Else
                    MsgBox strInfo, vbInformation + vbOKOnly, gstrSysName
                    If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    If cbo费别.ListIndex = -1 Then
        MsgBox "必须确定[费别]！", vbExclamation, gstrSysName
        If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus: Exit Function
    End If
    If cbo国籍.ListIndex = -1 Then
        MsgBox "必须确定[国籍]！", vbExclamation, gstrSysName
        If cbo国籍.Enabled And cbo国籍.Visible Then cbo国籍.SetFocus: Exit Function
    End If
    If cbo民族.ListIndex = -1 Then
        MsgBox "必须确定[民族]！", vbExclamation, gstrSysName
        If cbo民族.Enabled And cbo民族.Visible Then cbo民族.SetFocus: Exit Function
    End If
    
    If gbln启用结构化地址 And mbytInState <> E查阅 Then
        For i = PatiAddress.LBound To PatiAddress.UBound
            If Trim(PatiAddress(i).Value) <> "" And PatiAddress(i).CheckNullValue() <> "" Then
                MsgBox "病人的" & PatiAddress(i).Tag & "录入不完整,请重新录入。", vbInformation, gstrSysName
                If CanFocus(PatiAddress(i)) = True Then PatiAddress(i).SetFocus
                Exit Function
            End If
        Next
    End If
    '手机号合法性检查
    If Trim(txtMobile.Text) <> "" Then
        If Not IsNumeric(Trim(txtMobile.Text)) Or Len(Trim(txtMobile.Text)) <> 11 Then
            MsgBox "输入的手机号格式不正确，请重新录入！", vbInformation, gstrSysName
            If txtMobile.Enabled And txtMobile.Visible Then txtMobile.SetFocus: Exit Function
        Else
            If CheckMobile(Trim(txtMobile.Text), Val(txt病人ID.Text)) Then
                If MsgBox("在已有的病人信息中存在相同的手机号:" & Trim(txtMobile.Text) & "是否重新录入？", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If txtMobile.Enabled And txtMobile.Visible Then txtMobile.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    '长度检查
    If Not CheckTextLength("姓名", txtPatient) Then Exit Function
    If Not CheckTextLength("年龄", txt年龄) Then Exit Function
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Function
    
    '64701:刘鹏飞,2013-10-31,修改出生地址支持100个字符或50个汉字
    If Not CheckLen(txt出生地点, 100) Then Exit Function
    If Not CheckLen(txt户口地址, 100) Then Exit Function
    If Not CheckLen(txt户口地址邮编, 6) Then Exit Function
    If Not CheckLen(txt家庭地址, 100) Then Exit Function
    If Not CheckLen(txt家庭地址邮编, 6) Then Exit Function
    If Not CheckLen(txt家庭电话, 20) Then Exit Function
    If Not CheckLen(txt联系人姓名, 64) Then Exit Function
    If Not CheckLen(txt联系人地址, 100) Then Exit Function
    If Not CheckLen(txt联系人电话, 20) Then Exit Function
    If Not CheckLen(txt联系人身份证, 18) Then Exit Function
    If Not CheckLen(txt工作单位, txt工作单位.MaxLength) Then Exit Function
    If Not CheckLen(txt单位电话, 20) Then Exit Function
    If Not CheckLen(txtMobile, 20) Then Exit Function
    If Not CheckLen(txt单位邮编, 6) Then Exit Function
    If Not CheckLen(txt单位开户行, 50) Then Exit Function
    If Not CheckLen(txt单位帐号, 50) Then Exit Function
    If Not CheckLen(txt卡号, CInt(mCurSendCard.lng卡号长度)) Then Exit Function
    If Not CheckLen(txtPass, 10) Then Exit Function
    If Not CheckLen(txt缴款单位, 50) Then Exit Function
    If Not CheckLen(txt开户行, 50) Then Exit Function
    If Not CheckLen(txt帐号, 50) Then Exit Function
    If Not CheckLen(txt结算号码, 30) Then Exit Function
    '问题27351 by lesfeng 2010-01-12
    If Not CheckLen(txt备注, txt备注.MaxLength) Then Exit Function
    
    '104238:李南春，2017/2/15，检查卡号是否满足发卡控制限制
    If txt卡号.Text <> "" And Len(txt卡号.Text) <> mCurSendCard.lng卡号长度 And Not mCurSendCard.bln严格控制 Then
        Select Case mCurSendCard.byt发卡控制
            Case 0
                MsgBox "输入的卡号小于" & mCurSendCard.str卡名称 & "设定的卡号长度，请重新输入！", vbExclamation, gstrSysName
                If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                Exit Function
            Case 2
                If MsgBox("输入的卡号小于" & mCurSendCard.str卡名称 & "设定的卡号长度，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                    Exit Function
                End If
        End Select
    End If
    
    If IsCheck就诊卡 = False Then Exit Function
    '结算方式
    If IsNumeric(txt预交额.Text) And cbo预交结算.Visible And cbo预交结算.Enabled And cbo预交结算.ListIndex = -1 Then
        MsgBox "请确定病人预交款结算方式！", vbInformation, gstrSysName
        cbo预交结算.SetFocus: Exit Function
    End If
    
    '问题号:53408
'    If IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, glngModul) = "1", 1, 0) = 0 And ((mCurSendCard.str卡名称 = "二代身份证" And Trim(txt卡号.Text) <> "") Or Trim(txt支付密码.Text) <> "") Then
'         MsgBox "您没有权限进行签约操作,请到参数设置中设置【扫描身份证签约】！", vbOKOnly + vbInformation, gstrSysName
'         txt卡号.Text = ""
'         txtPass.Text = ""
'         txtAudi.Text = ""
'         If txt卡号.Visible = True Then txt卡号.SetFocus
'         Exit Function
'    End If
    
    If Trim(txt支付密码.Text) <> "" And Trim(txt身份证号.Text) <> "" Then
           If 是否已经签约(txt身份证号.Text) Then
                 MsgBox "身份证号码为:" & txt身份证号.Text & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
                 txt支付密码.Text = ""
                 If txt支付密码.Visible = True Then txt支付密码.SetFocus
                 Exit Function
           End If
    End If
    
    If mbln是否扫描身份证 = False And mCurSendCard.str卡名称 = "二代身份证" And txt卡号.Text <> "" Then
            MsgBox "绑定身份证只能以刷卡的方式进行，不允许手动输入身份证进行绑定!", vbOKOnly + vbInformation, gstrSysName
            txt卡号.Text = ""
            txtPass.Text = ""
            txtAudi.Text = ""
            txt支付密码.Text = ""
            txt验证密码.Text = ""
            If txt卡号.Visible = True Then txt卡号.SetFocus
            Exit Function
    End If
    
    If mbln是否扫描身份证 = False And mCurSendCard.str卡名称 <> "二代身份证" And txt支付密码.Text <> "" Then
            MsgBox "绑定身份证只能以刷卡的方式进行，不允许手动输入身份证进行绑定!", vbOKOnly + vbInformation, gstrSysName
            txt身份证号.Text = ""
            txt支付密码.Text = ""
            txt验证密码.Text = ""
            If txt身份证号.Visible = True Then txt身份证号.SetFocus
        Exit Function
    End If
    
    If Trim(txt支付密码.Text) <> Trim(txt验证密码.Text) And (Trim(txt支付密码.Text) <> "" Or Trim(txt验证密码.Text) <> "") Then
        MsgBox "两次输入的密码不一致,请重新输入", vbOKOnly + vbInformation, gstrSysName
        txt支付密码.Text = "": txt验证密码.Text = ""
        If txt支付密码.Visible = True Then txt支付密码.SetFocus
        Exit Function
    End If
    
    blnModify = blnMod And bln基本信息调整
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckNewPati() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查新病人
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSimilar As String, strMCAccount As String, strNote As String
    Dim i As Long, lng接口编号 As Long, strBalanceInfor As String
    Dim str称呼 As String
    Dim lngTmp As Long
    Dim rsSimilar As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If Trim(txt卡号.Text) <> "" And txtPass.Visible Then
        Select Case mCurSendCard.int密码长度限制
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & mCurSendCard.int密码长度 & "位", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(mCurSendCard.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        End Select
    End If
    If mlngPatientID <> 0 Then CheckNewPati = True: Exit Function
    If Not CheckPatiIsExist() Then Exit Function

    '病人ID检查
    '问题27207 by lesfeng 2010-1-4
    If ExistInPatiID(CLng(txt病人ID.Text)) Then
        If txt病人ID.Text <> lbl病人ID.Tag Then
            MsgBox "该" & str称呼 & "的标识 " & txt病人ID.Text & " 已经被使用，" & vbCrLf & _
                "系统将自动更换一个不重复的标识！", vbInformation, gstrSysName
            txt病人ID.Text = zlDatabase.GetNextNo(1): lbl病人ID.Tag = txt病人ID.Text
            cmdOK.SetFocus: Exit Function
        Else
            '自动产生的号如果没有修改，则直接再次自动产生即可
            txt病人ID.Text = zlDatabase.GetNextNo(1): lbl病人ID.Tag = txt病人ID.Text
        End If
    End If
    
    '门诊号检查
    If IsNumeric(txt门诊号.Text) Then
        '问题27207 by lesfeng 2010-1-4
        If ExistClinicNO(txt门诊号.Text) Then
            If txt门诊号.Text <> lbl门诊号.Tag Then
                MsgBox "发现该病人的病人门诊号[" & txt门诊号.Text & "]已经被其它病人使用,系统将自动更换一个不重复的号码！", vbInformation, gstrSysName
                txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
                cmdOK.SetFocus: Exit Function
            Else
                '自动产生的号如果没有修改，则直接再次自动产生即可
                txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
            End If
        End If
    End If
    

    CheckNewPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetCardVaribles(ByVal blnPrepay As Boolean)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置结算对象数据
    '入参:blnPrepay-是否预交结算对象
    '编制:刘尔旋
    '日期:2014-01-07
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    
    If blnPrepay = True Then
        With cbo预交结算
            If .ListIndex = -1 Then Exit Sub
            lngIndex = .ListIndex + 1
        End With
        '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
        If Not mcolPrepayPayMode Is Nothing Then
            With mCurPrepay
                    .lng医疗卡类别ID = Val(mcolPrepayPayMode(lngIndex)(3))
                    .bln消费卡 = Val(mcolPrepayPayMode(lngIndex)(5)) = 1
                    .str结算方式 = Trim(mcolPrepayPayMode(lngIndex)(6))
                    .str名称 = Trim(mcolPrepayPayMode(lngIndex)(1))
             End With
        End If
    Else
        With cbo结算方式
            If .ListIndex = -1 Then Exit Sub
            lngIndex = .ListIndex + 1
        End With
        If Not mcolCardPayMode Is Nothing Then
            With mCurCardPay
                .lng医疗卡类别ID = Val(mcolCardPayMode(lngIndex)(3))
                .bln消费卡 = Val(mcolCardPayMode(lngIndex)(5)) = 1
                .str结算方式 = Trim(mcolCardPayMode(lngIndex)(6))
                .str名称 = Trim(mcolCardPayMode(lngIndex)(1))
             End With
         End If
     End If
End Sub

Private Sub cmdOK_Click()
    Dim strMCAccount As String, str称呼 As String
    Dim blnOK As Boolean
    Dim blnModify As Boolean
    Dim strErrInfo As String
    Dim str性别 As String, str年龄 As String, str出生日期 As String
    
    '问题号:56599
    tbcPage.Item(0).Selected = True
    
    str称呼 = IIf(glngSys Like "8??", "客户", "病人")
    If IsValied(blnModify, str出生日期, str年龄, str性别) = False Then Exit Sub
    
    '69231,刘尔旋,2014-01-07 14:42:55,保存时强制更新卡对象数据
    Call SetCardVaribles(False)
    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If mlngOutModeMC = 920 And strMCAccount <> txtPatiMCNO(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
        If CheckExistsMCNO(strMCAccount) Then
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Sub
        End If
    End If
    
    If CheckBrushCard = False Then Exit Sub
    mblnPrepayPrint = False
    
    If IsNumeric(txt预交额.Text) Then
        mblnPrepayPrint = True
        '检查是否打印票据
        '78751:李南春,2015/08/24,增加预交票据打印格式
        Select Case mFactProperty.intInvoicePrint
            Case "0" '不打印预交发票
               mblnPrepayPrint = False
            Case "1" '自动打印
               mblnPrepayPrint = True
            Case "2" '打印提醒
                mblnPrepayPrint = MsgBox("是否打印预交款票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
        End Select
        
        If mblnPrepayPrint Then
            If gblnBill预交 Then
                If Trim(txtFact.Text) = "" Then
                    MsgBox "必须输入一个有效的预交票据号码！", vbInformation, gstrSysName
                    txtFact.SetFocus: Exit Sub
                End If
                
                mlng预交领用ID = CheckUsedBill(2, IIf(mlng预交领用ID > 0, mlng预交领用ID, mFactProperty.lngShareUseID), txtFact.Text, Val(Mid(tbDeposit.SelectedItem.Key, 2)))
                If mlng预交领用ID <= 0 Then
                    Select Case mlng预交领用ID
                        Case 0 '操作失败
                        Case -1
                            MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                        Case -2
                            MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                        Case -3
                            MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
                            txtFact.SetFocus
                    End Select
                    Exit Sub
                End If
            Else
                If Len(txtFact.Text) <> gbyt预交 And txtFact.Text <> "" Then
                    MsgBox "预交票据号码长度应该为 " & gbyt预交 & " 位！", vbInformation, gstrSysName
                    txtFact.SetFocus: Exit Sub
                End If
            End If
        End If
    End If
    
    '63246:刘鹏飞,2013-07-03
    If CheckPatiCard = False Then Exit Sub
    
    '73937:刘鹏飞,2013-07-03
    If CreatePlugInOK(glngModul) Then
        blnOK = True
        On Error Resume Next
        blnOK = gobjPlugIn.PatiInfoSaveBefore(Val(txt病人ID.Text))
        If blnOK = False Then
            If tbcPage.Item(tbcPage.ItemCount).Caption = "附加信息" Then tbcPage.Item(tbcPage.ItemCount).Selected = True
            Err.Clear
            Exit Sub
        End If
        Err.Clear: On Error GoTo 0
    End If
    '病人信息从表\病案主页从表处理
    mstrPatiPlus = ""
    '身份证号未录入时附加消息
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) = "中国" Then
        mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
        mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:"
    Else
        If Trim(txt身份证号.Text) <> "" Then
            mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:"
            mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:" & Trim(txt身份证号.Text)
            txt身份证号.Text = ""
        Else
            mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
            mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:"
        End If
    End If
    If mstrPatiPlus <> "" Then mstrPatiPlus = Mid(mstrPatiPlus, 2)
    
    If mbytInState = E新增 Then
         If CheckNewPati = False Then Exit Sub
        '保存新卡
        '--------------------------------------------------------------
        If Not SaveNewCard(strMCAccount) Then
            MsgBox str称呼 & "身份登记失败,请重试该操作！", vbExclamation, gstrSysName
            Exit Sub
        End If
        '病人信息保存成功,根据身份证信息同步调整病人信息的性别,年龄和日期
        If blnModify Then
            strErrInfo = ""
            Call gobjPublicPatient.SavePatiBaseInfo(mlng病人ID, mlng主页ID, Trim(txtPatient.Text), str性别, str年龄, str出生日期, Me.Caption, IIf(mlng病人ID = 0, 1, 2), strErrInfo, False, True)
            If strErrInfo <> "" Then
                MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
            End If
        End If
        '打印预交款收据
        '78751:李南春,2015/08/24,增加预交票据打印格式
        If mblnPrepayPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mCurPrepay.strNO, "收款时间=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
                            "病人ID=" & Val(txt病人ID), IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
        End If
        
        '打印病案主页
        If InStr(mstrPrivs, "首页打印") > 0 Then
            If MsgBox("病人信息保存成功，要打印病案首页吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1101", Me, "病人ID=" & Val(txt病人ID.Text), 2)
            End If
        End If
        
        gblnOK = True
        
        '保存后继续下一个病人信息
        Call ClearCard
        mblnICCard = False  '不能放在clearcard中,因为可能先读卡再查出病人
        '问题27207 by lesfeng 2010-1-4
        txt病人ID.Text = zlDatabase.GetNextNo(1): lbl病人ID.Tag = txt病人ID.Text
        txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
        
        If Not mCurSendCard.rs卡费 Is Nothing Then txt卡额.Text = Format(IIf(mCurSendCard.rs卡费!是否变价 = 1, mCurSendCard.rs卡费!缺省价格, mCurSendCard.rs卡费!现价), "0.00"): txt卡额.Tag = txt卡额.Text
        
        '预交款检查
        If mblnPrepayPrint Then
            If Not gblnBill预交 Then
                zlDatabase.SetPara "当前预交票据号", txtFact.Text, glngSys, mlngModul
            End If
            Call GetFact(False)
        End If
        
        '就诊卡领用检查
        If mCurSendCard.bln严格控制 Then
            mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), , mCurSendCard.lng卡类别ID)
            If mCurSendCard.lng领用ID <= 0 Then
                Select Case mCurSendCard.lng领用ID
                    Case 0 '操作失败
                    Case -1
                        If txt卡号.Text <> "" Then MsgBox "你已没有自用及共用的" & IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡,不能再发放！" & vbCrLf & _
                            "请先在本地设置共用批次或领用一批新卡！", vbExclamation, gstrSysName
                    Case -2
                        If txt卡号.Text <> "" Then MsgBox "本地共用的" & IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡已用完,你不能再发放！" & vbCrLf & _
                            "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
                End Select
            End If
        End If
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    ElseIf mbytInState = E修改 Then
        '门诊号检查
        If IsNumeric(txt门诊号.Text) Then
            If ExistClinicNO(txt门诊号.Text, CLng(txt病人ID.Text)) Then
                '问题27207 by lesfeng 2010-1-4
                If txt门诊号.Text <> lbl门诊号.Tag Then
                    MsgBox "发现该病人的病人门诊号[" & txt门诊号.Text & "]已经被其它病人使用,系统将自动更换一个不重复的号码！", vbInformation, gstrSysName
                    txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
                    cmdOK.SetFocus: Exit Sub
                Else
                    '自动产生的号如果没有修改，则直接再次自动产生即可
                    txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
                End If
            End If
        End If
        
        '住院号检查
        If IsNumeric(txt住院号.Text) Then
            If ExistInPatiNO(Trim(txt住院号.Text), Val(txt病人ID.Text)) Then
                MsgBox "发现该病人的病人住院号[" & txt住院号.Text & "]已经被其它病人使用,系统将自动更换一个不重复的号码！", vbInformation, gstrSysName
                txt住院号.Text = zlDatabase.GetNextNo(2)
                cmdOK.SetFocus: Exit Sub
            End If
        End If
        '保存修改
        '--------------------------------------------------------------------
        If Not SaveModiCard(strMCAccount) Then
            MsgBox "保存失败,请重试该操作！", vbExclamation, gstrSysName
            Exit Sub
        End If
        '病人信息保存成功,根据身份证信息同步调整病人信息的性别,年龄和日期
        If blnModify Then
            strErrInfo = ""
            Call gobjPublicPatient.SavePatiBaseInfo(mlng病人ID, mlng主页ID, Trim(txtPatient.Text), str性别, str年龄, str出生日期, Me.Caption, IIf(mlng主页ID = 0, 1, 2), strErrInfo, True, True)
            If strErrInfo <> "" Then
                MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
            End If
        End If
        '修改后退出
        gblnOK = True
        Unload Me: Exit Sub
    End If
End Sub

Private Sub cmdOperation_Click(Index As Integer)
    Dim bln缴预交 As Boolean, bln退预交 As Boolean
    Dim lng病人ID As Long
    
    Dim strPrivs As String
    On Error Resume Next
    Select Case Index
    Case 0
        Call InitLocPar(1103)
        strPrivs = ";" & GetPrivFunc(glngSys, 1103) & ";"
        bln缴预交 = InStr(1, strPrivs, ";门诊预交;") > 0 Or InStr(1, strPrivs, ";住院预交;") > 0 Or InStr(1, strPrivs, ";共用预交;") > 0
        bln退预交 = InStr(1, strPrivs, ";预交退款;") > 0
        If bln退预交 = False And bln缴预交 = False Then Exit Sub
        Call frmDeposit.zlShowEdit(Me, 0, IIf(bln缴预交, 0, 2), strPrivs, 1103)
        Call InitLocPar(mlngModul)
    Case 1
        '调用就诊卡发卡管理
        strPrivs = ";" & GetPrivFunc(glngSys, 1107) & ";"
        If gobjSquare.objSquareCard.zlSendCard(Me, mlngModul, mCurSendCard.lng卡类别ID, lng病人ID, strPrivs) = False Then Exit Sub
        'frmIDCard.mbytInState = E新增
       ' frmIDCard.Show 1, Me
    End Select
    Err.Clear
End Sub

Private Sub cmd出生地点_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
            " Substr(名称,1,2) as 名称 From 地区" & _
            " Union All" & _
            " Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
            " From 地区 Order by 编码", 2, "地区", , txt出生地点.Text)
    If Not rsTmp Is Nothing Then
        txt出生地点.Text = rsTmp!名称
        txt出生地点.SelStart = Len(txt出生地点.Text)
        txt出生地点.SetFocus
    End If
End Sub

Private Sub cmd合同单位_Click()
    Dim rsTmp As ADODB.Recordset
    '问题27040 by lesfeng 对合约单位加上撤档时间的处理
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & _
            "  Where (撤档时间 IS NULL OR TO_CHAR(撤档时间, 'yyyy-MM-dd') = '3000-01-01') " & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID", _
            2, "单位", , txt工作单位.Text)
    If Not rsTmp Is Nothing Then
        txt工作单位.Tag = rsTmp!ID
        txt工作单位.Text = rsTmp!名称
        txt工作单位.SelStart = Len(txt工作单位.Text)
        txt单位电话.Text = Trim(rsTmp!电话 & "")
        txt单位开户行.Text = Trim(rsTmp!开户银行 & "")
        txt单位帐号.Text = Trim(rsTmp!帐号 & "")
        txt工作单位.SetFocus
    End If
End Sub

Private Sub cmd家庭地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
            " Substr(名称,1,2) as 名称 From 地区" & _
            " Union All" & _
            " Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
            " From 地区 Order by 编码", 2, "地区", , txt出生地点.Text)
    If Not rsTmp Is Nothing Then
        txt家庭地址.Text = rsTmp!名称
        txt家庭地址.SelStart = Len(txt家庭地址.Text)
        txt家庭地址.SetFocus
    End If
End Sub

Private Sub cmd联系人地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
            " Substr(名称,1,2) as 名称 From 地区" & _
            " Union All" & _
            " Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
            " From 地区 Order by 编码", 2, "地区", , txt出生地点.Text)
    If Not rsTmp Is Nothing Then
        txt联系人地址.Text = rsTmp!名称
        txt联系人地址.SelStart = Len(txt联系人地址.Text)
        txt联系人地址.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If mbytInState = E新增 And mblnSel = False Then txtPatient.SetFocus
    '问题号:53408
    mbln扫描身份证签约 = IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, glngModul) = "1", 1, 0) = "1"
    If mCurSendCard.str卡名称 Like "*二代身份证*" Then
        lbl就诊卡号.Enabled = False: txt卡号.Enabled = False
        lbl密码.Enabled = False: txtPass.Enabled = False
        lbl验证.Enabled = False: txtAudi.Enabled = False
    End If
    mblnSel = True
    Call SetCardEditEnabled
    Call Show绑定控件(False)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim obj As Control
    
    Select Case KeyCode
        Case vbKeyF2
            If cmdOK.Visible And cmdOK.Enabled Then
                cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF3
            If Me.ActiveControl.Name = txt出生地点.Name _
                And cmd出生地点.Enabled And cmd出生地点.Visible Then
                cmd出生地点_Click
            ElseIf Me.ActiveControl.Name = txt家庭地址.Name _
                And cmd家庭地址.Enabled And cmd家庭地址.Visible Then
                cmd家庭地址_Click
            ElseIf Me.ActiveControl.Name = txt联系人地址.Name _
                And cmd联系人地址.Enabled And cmd联系人地址.Visible Then
                cmd联系人地址_Click
            ElseIf Me.ActiveControl.Name = txt工作单位.Name _
                And cmd合同单位.Enabled And cmd合同单位.Visible Then
                cmd合同单位_Click
            ElseIf Me.ActiveControl.Name = txt区域.Name And cmd区域.Enabled And cmd区域.Visible Then
                cmd区域_Click
            End If
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC卡号")
                If intIndex < 0 Then Exit Sub
                IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyReturn
            Set obj = Me.ActiveControl
            If obj.Name = "txtPatient" Then
                If txtPatient.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo性别" Then
                If cbo性别.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo费别" Then
                If cbo费别.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo结算方式" Then
                If cbo结算方式.ListIndex <> -1 Then cmdOK.SetFocus
            '问题 25458 增加 txtPatiMCNO判断 实现单个 vbKeyTab
            ElseIf InStr(1, ",txt卡号,txt出生地点,txt家庭地址,txt户口地址,txt联系人地址,txt工作单位,txtPass,txtAudi,txt卡额,txt年龄,txt预交额,txtPatiMCNO,vsDrug,vsInoculate,vsLinkMan,vsOtherInfo,PatiAddress,", "," & obj.Name & ",") <= 0 Then
                If Not obj Is txtPass Then
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
        End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    mlng图像操作 = 0: mstr采集图片 = "":
    With mPageHeight
        .基本 = Me.Height
        .健康档案 = Me.Height
        .附加信息 = Me.Height
    End With
    '上次默认预交类型
    mbytPrepayType = Val(zlDatabase.GetPara("上次预交类型", glngSys, mlngModul, "0"))
    '初始化
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
     '初始成功,则证明此窗口存在相关的结算卡
     mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
    'Call zlCardSquareObject:
    Call CreateObjectKeyboard
    
     
    If glngSys Like "8??" Then
        Me.Caption = "客户信息卡片"
        lbl病人ID.Caption = "客户ID"
        lbl门诊号.Visible = False
        txt门诊号.Visible = False
        txt门诊号.Text = ""
        
        lbl住院号.Visible = False
        txt住院号.Visible = False
        txt住院号.Text = ""
        '问题27351 by lesfeng 2010-01-12
        txt备注.Visible = False
        lbl备注.Visible = False
        txt备注.Text = ""
        txt备注.Tag = "0"
        
        chk记帐.Visible = False
        lbl结算方式.Visible = True
        
        lbl费别.Caption = "会员等级"
    Else
        Me.Caption = "病人信息" & Choose(mbytInState + 1, "登记", "修改", "卡片")
        If mbytInState = E新增 Then
            lbl费别.Caption = "门诊费别" '新增时不可能为住院费别
        Else
            If mbytView = 1 Or mbytView = 2 Then
                lbl费别.Caption = "住院费别"
            Else
                lbl费别.Caption = "门诊费别"
            End If
        End If
    End If
    
    '问题27356 by lesfeng 2010-01-13
    If InStr(mstrPrivs, "绑定卡号") = 0 Then
        tabCardMode.Tabs.Remove ("CardBind")
        tabCardMode.Width = tabCardMode.Width / 2
    End If
    
    mblnChange = True
    gblnOK = False
    mblnUnLoad = False
    mstrYBPati = ""
    txt出生日期.Tag = "0"
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    mblnChange = False: cbo年龄单位.ListIndex = 0: mblnChange = True
    '问题号:56599
    Call InitCard
    Call InitTabPage
    
    'SetCreateCardObject '创建写卡对象
    Call zlCreateSquare
    Call HookDefend(txt支付密码.hWnd)
    Call HookDefend(txt验证密码.hWnd)
    Call HookDefend(txtPass.hWnd)
    Call HookDefend(txtAudi.hWnd)
    
    If mblnUnLoad Then Unload Me: Exit Sub
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Integer
    
    If Not Visible Then Exit Sub
    
'    If tbcPage.Selected.Index = 0 Then
'        If fraDeposit.Visible = False Then
'            fraCard.Top = fraInfo.Top + fraInfo.Height + 30
'            cmdOK.Top = fraCard.Top + fraCard.Height + 500
'        End If
'
'        If fraCard.Visible = False Then
'            cmdOK.Top = IIf(fraDeposit.Visible = True, fraDeposit.Top + fraDeposit.Height, fraInfo.Top + fraInfo.Height) + 500
'        End If
'    Else
'        cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 140
'    End If
    cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 140
    cmdHelp.Top = cmdOK.Top
    cmdCancel.Top = cmdOK.Top
    If cmdOperation(OPT.C0预交款).Visible Then cmdOperation(OPT.C0预交款).Top = cmdHelp.Top
    If cmdOperation(OPT.C1就诊卡).Visible Then cmdOperation(OPT.C1就诊卡).Top = cmdHelp.Top
    If cmdOperation(OPT.C0预交款).Visible = False Then cmdOperation(OPT.C1就诊卡).Left = cmdOperation(OPT.C0预交款).Left
    tbcPage.Height = cmdOK.Top - 120
    tbcPage.Width = Me.ScaleWidth - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mlng图像操作 = 0: mstr采集图片 = ""

    '问题号:565999
    If Not mobjHealthCard Is Nothing Then
        Set mobjHealthCard = Nothing
    End If
    
    If Not mobjKeyboard Is Nothing Then
        Set mobjKeyboard = Nothing
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.Terminate(glngSys, 1101)
        Err.Clear: On Error GoTo 0
        Set gobjPlugIn = Nothing
    End If
    
    If Not mobjPublicPatient Is Nothing Then
        Set mobjPublicPatient = Nothing
    End If
    If Not mobjSquare Is Nothing Then Set mobjSquare = Nothing
    If Not mobjCommEvents Is Nothing Then Set mobjCommEvents = Nothing
    mbln发卡或绑定卡 = False
    
    '82401:李南春,2015/3/11,检查对象是否存在
    If mbytInState = E新增 And fraCard.Visible = True Then
        zlDatabase.SetPara "发卡模式", tabCardMode.SelectedItem.Key, glngSys, mlngModul
    End If
    
    mblnICCard = False: mbytInState = E新增
    mblnUnLoad = False: mlng病人ID = 0: mlng主页ID = 0
    mCurSendCard.lng领用ID = 0: mlng预交领用ID = 0: mstrPrivs = ""
    Call ClearCard: mblnSel = False
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    If Not mdic医疗卡属性 Is Nothing Then
        Set mdic医疗卡属性 = Nothing
    End If
    Err = 0: On Error Resume Next
    'Call zlCardSquareObject(True)
End Sub

Private Sub InitDicts()
    Call ReadDict("性别", cbo性别)
    Call ReadDict("费别", cbo费别)
    Call ReadDict("医疗付款方式", cbo医疗付款)
    Call ReadDict("国籍", cbo国籍)
    Call ReadDict("民族", cbo民族)
    Call ReadDict("学历", cbo学历)
    Call ReadDict("婚姻状况", cbo婚姻状况)
    Call ReadDict("职业", cbo职业)
    Call ReadDict("身份", cbo身份)
    Call ReadDict("社会关系", cbo联系人关系)
    Call ReadDict("病人类型", cbo病人类型, "病人类型")
    Call ReadDict("身份证未录原因", cboIDNumber)
    If mbytInState = E新增 Then
        'Call ReadDict("结算方式", cbo结算方式, "就诊卡")
        'Call ReadDict("结算方式", cbo预交结算, "预交款")
    End If
End Sub

Private Function ReadDict(strDict As String, cbo As ComboBox, Optional strClass As String) As Boolean
'功能：初始化指定词典
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long

    On Error GoTo errH
    'by lesfeng 2010-03-08 性能优化
    If strDict = "结算方式" Then
        If strClass = "预交款" Then
            strSQL = "1,2,5,8"
        Else
            strSQL = "1,2"
        End If
        strSQL = "Select Nvl(A.缺省标志,0) as 缺省,B.编码,B.名称,B.性质" & _
            " From 结算方式应用 A,结算方式 B" & _
            " Where A.结算方式=B.名称 And A.应用场合=[1]" & _
            " And Nvl(B.性质,1) IN(" & strSQL & ") Order by B.编码"
    ElseIf strDict = "身份" Then
        strSQL = "Select 编码,名称,简码,Nvl(优先级,0) as 缺省 From " & strDict & " Order by 编码"
    ElseIf strDict = "费别" Then
        '根据视图性质,配合过程参数,决定费别服务对象
        'mbytView:0-所有,1-在院,2-出院,3-门诊
        If glngSys Like "8??" Then
            strSQL = "1,3" '药店系统使用门诊费别
        ElseIf mbytInState = E新增 Then
            strSQL = "1,3" '新增时使用门诊费别
        Else
            If mbytView = 1 Or mbytView = 2 Then
                strSQL = "2,3" '查看/修改时使用住院费别
            Else
                strSQL = "1,3" '查看/修改时使用门诊费别
            End If
        End If
        
        '不是仅限初诊身份唯一性项目(包含了缺省费别),不管有效期间及科室
        strSQL = _
            " Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别" & _
            " Where 属性=1 And Nvl(仅限初诊,0)=0 And Nvl(服务对象,3) IN(" & strSQL & ")" & _
                " And  Sysdate Between NVL(有效开始,Sysdate-1) and NVL(有效结束,Sysdate+1)" & _
            " Order by 编码"
    ElseIf strDict = "病人类型" Then
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省,颜色 From 病人类型 Order by 编码"
    Else
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strClass)
    cbo.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If strDict = "职业" Then
                cbo.AddItem rsTmp!编码 & "-" & Chr(&HA) & rsTmp!名称
            Else
                cbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
            End If
            If rsTmp!缺省 = 1 Then
                cbo.ListIndex = cbo.NewIndex
                cbo.ItemData(cbo.NewIndex) = 1
            End If
            If strDict = "结算方式" And strClass = "预交款" Then
                   cbo.ItemData(cbo.NewIndex) = Val(Nvl(rsTmp!性质))
                   cbo.Tag = cbo.NewIndex   '单独保存为缺省的性质索引
            End If
            
            If TextWidth(cbo.List(cbo.NewIndex) & "两个") > lngMaxW Then lngMaxW = TextWidth(cbo.List(cbo.NewIndex) & "两个")
            rsTmp.MoveNext
        Next
        If strDict = "结算方式" And strClass <> "预交款" Then cbo.Tag = cbo.Text
        
    ElseIf strDict = "结算方式" Then
        If mbytInState = E新增 Then
            If glngSys Like "8??" Then
                MsgBox "会员卡场合没有可用的结算方式，不能发卡！" & vbCrLf & _
                    "请先到结算方式管理中设置会员卡的结算方式。", vbInformation, gstrSysName
                fraCard.Visible = False: cmdOperation(OPT.C1就诊卡).Visible = False
                Me.Height = Me.Height - fraCard.Height
                mPageHeight.基本 = Me.Height
            Else
                MsgBox "就诊卡场合没有可用的结算方式，只能使用记帐方式发卡！" & vbCrLf & _
                    "要使用结算发卡,请先到结算方式管理中设置就诊卡结算方式。", vbInformation, gstrSysName
                chk记帐.Value = 1: chk记帐.Enabled = False: cbo.Enabled = False
                chk记帐.Tag = 1
            End If
        End If
    End If
    ReadDict = True
    If GetWidth(cbo.hWnd) * Screen.TwipsPerPixelX < lngMaxW Then SetWidth cbo.hWnd, lngMaxW / Screen.TwipsPerPixelX
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub PatiAddress_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True) '打开中文输入法
End Sub

Private Sub PatiAddress_LostFocus(Index As Integer)
'功能:
    Select Case Index
    
    Case E_IX_现住址
        txt家庭地址.Text = PatiAddress(Index).Value
    Case E_IX_出生地点
        txt出生地点.Text = PatiAddress(Index).Value
    Case E_IX_户口地址
        txt户口地址.Text = PatiAddress(Index).Value
    Case E_IX_籍贯
        txt籍贯.Text = PatiAddress(Index).Value
    Case E_IX_联系人地址
        txt联系人地址.Text = PatiAddress(Index).Value
    End Select
    Call zlCommFun.OpenIme '关闭中文输入法
End Sub

Private Sub PatiAddress_Validate(Index As Integer, Cancel As Boolean)
    Dim lngLen As Long
    
    lngLen = PatiAddress(Index).MaxLength
    If LenB(StrConv(PatiAddress(Index).Value, vbFromUnicode)) > lngLen Then
        MsgBox PatiAddress(Index).Tag & "只允许输入 " & lngLen & " 个字符或 " & lngLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub tabCardMode_Click()
    If tabCardMode.SelectedItem.Key = "CardFee" Then
        lbl金额.Visible = True
        txt卡额.Visible = True
        chk记帐.Visible = True
        cbo结算方式.Visible = True
    Else
        lbl金额.Visible = False
        txt卡额.Visible = False
        chk记帐.Visible = False
        cbo结算方式.Visible = False
    End If
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '问题号:56599
    Dim intIndex As Integer, objItem As TabControlItem
    mbln基本 = IIf(Item.Caption = "基本", True, False)
    Select Case Item.Caption
        Case "基本"
            Me.Height = mPageHeight.基本
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Case "健康档案"
            Me.Height = mPageHeight.健康档案
            If cboBloodType.Enabled And cboBloodType.Visible Then cboBloodType.SetFocus
        Case "附加信息"
            Me.Height = mPageHeight.附加信息
            If Item.Handle = picTmp.hWnd Then
                intIndex = Item.Index
                Call HideFormCaption(mlngPlugInHwnd, False)
                Set objItem = tbcPage.InsertItem(intIndex, "附加信息", mlngPlugInHwnd, 0)
                objItem.Tag = mPageHeight.附加信息
                Call tbcPage.RemoveItem(intIndex + 1)
                objItem.Selected = True
                picTmp.Visible = False
            End If
    End Select
End Sub

Private Sub tbDeposit_Click()
    If mblnNotClick Then Exit Sub
     
    'If fraDeposit.Visible = False Then Exit Sub
    If tbDeposit.SelectedItem Is Nothing Then Exit Sub
    mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, Val(Mid(tbDeposit.SelectedItem.Key, 2)))
    mlng预交领用ID = 0
    Call GetFact(False)
End Sub

Private Sub GetFact(Optional blnFirst As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取不同类别的发票
    '编制:刘兴洪
    '日期:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnBill预交 = False Then
        '松散：取下一个号码
        txtFact.Text = IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModul, "")))
        Exit Sub
    End If
    '严格:     取下一个号码
    mlng预交领用ID = CheckUsedBill(2, IIf(mlng预交领用ID > 0, mlng预交领用ID, mFactProperty.lngShareUseID), , Val(Mid(tbDeposit.SelectedItem.Key, 2)))
    If mlng预交领用ID <= 0 Then
        Select Case mlng预交领用ID
            Case 0 '操作失败
'            Case -1
'                MsgBox "你没有自用或共用的预交票据,登记病人信息时不能同时缴预交款！" & _
'                    "请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
'            Case -2
'                MsgBox "本地的共用票据已经用完,登记病人信息时不能同时缴预交款！" & _
'                    "请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
        txtFact.Text = ""
        'fraDeposit.Visible = False
      '  Me.Height = Me.Height - fraDeposit.Height
    Else
        txtFact.Text = GetNextBill(mlng预交领用ID)
    End If
End Sub
Private Sub txtAudi_GotFocus()
    SelAll txtAudi
    OpenPassKeyboard txtAudi, True
End Sub
Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mCurSendCard.int密码规则 = 1 Then
            Call zlControl.TxtCheckKeyPress(txtAudi, KeyAscii, m数字式)
        End If
    End If
    
    If KeyAscii = 13 Then
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txtAudi)
            If txtAudi.Enabled And txtAudi.Visible Then txtAudi.SetFocus
        Else
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub
Private Sub txtAudi_LostFocus()
    Call ClosePassKeyboard(txtAudi)
End Sub

Private Sub txtFact_GotFocus()
    zlControl.TxtSelAll txtFact
End Sub

Private Sub txtFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtFact.Text) = txtFact.MaxLength And KeyAscii <> 8 And txtFact.SelLength <> Len(txtFact) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtMedicalWarning_GotFocus()
    zlControl.TxtSelAll txtMedicalWarning
End Sub

Private Sub txtMobile_GotFocus()
    zlControl.TxtSelAll txtMobile
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOtherWaring_GotFocus()
    zlControl.TxtSelAll txtOtherWaring
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtOtherWaring_KeyPress(KeyAscii As Integer)
    If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    CheckInputLen txtOtherWaring, KeyAscii
End Sub

Private Sub txtOtherWaring_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mCurSendCard.int密码规则 = 1 Then
            Call zlControl.TxtCheckKeyPress(txtPass, KeyAscii, m数字式)
        End If
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            If Not txt卡额.Locked And txt卡额.TabStop And txt卡额.Enabled Then
                    txt卡额.SetFocus
            ElseIf chk记帐.Visible And chk记帐.Enabled Then
                chk记帐.SetFocus
            ElseIf Me.cbo结算方式.Enabled And cbo结算方式.Visible Then
                cbo结算方式.SetFocus
            Else
                Debug.Print "A1"
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
           Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
    'mlngPatientID = 0
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.ActiveFastKey
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
'问题27351 by lesfeng 2010-01-12  b
Private Sub txt备注_GotFocus()
    Call zlControl.TxtSelAll(txt备注)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    CheckInputLen txt备注, KeyAscii
End Sub

Private Sub txt备注_LostFocus()
    Call zlCommFun.OpenIme
End Sub
'问题27351 by lesfeng 2010-01-12 e
Private Sub txt病人ID_Change()
    '问题27207 by lesfeng 2010-1-4
    lbl病人ID.Tag = "" '记录自动编号是否被人工修改
End Sub

Private Sub txt病人ID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
        glngTXTProc = GetWindowLong(txt病人ID.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt病人ID.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt病人ID_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
        Call SetWindowLong(txt病人ID.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt出生地点_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt出生地点.Text <> "" Then
            '问题32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select 编码 as ID,编码,名称,简码 From 地区" & _
                    " Where 编码 Like '" & gstrLike & txt出生地点.Text & "%'" & _
                    " Or 简码 Like '" & gstrLike & txt出生地点.Text & "%'" & _
                    " Or 名称 Like '" & gstrLike & txt出生地点.Text & "%'", _
                    0, "地区", , txt出生地点.Text)
            If Not rsTmp Is Nothing Then
                txt出生地点.Text = rsTmp!名称
                mblnSel = True
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt出生地点, KeyAscii
End If
End Sub

Private Sub txt出生地点_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt出生日期_Change()
    Dim str出生日期 As String
    If IsDate(txt出生日期.Text) And mblnChange Then
        mblnChange = False
        txt出生日期.Text = Format(CDate(txt出生日期.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        If txt出生时间.Text = "__:__" Then
            str出生日期 = Format(txt出生日期.Text, "YYYY-MM-DD")
        Else
            str出生日期 = Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        txt年龄.Text = ReCalcOld(CDate(str出生日期), cbo年龄单位, , , CDate(txt出生日期.Tag))
    End If
End Sub

Private Sub txt出生日期_LostFocus()
    If txt出生日期.Text <> "____-__-__" And Not IsDate(txt出生日期.Text) Then
        txt出生日期.SetFocus
    End If
End Sub

Private Sub txt出生时间_Change()
    Dim str出生日期 As String
    
    If IsDate(txt出生时间.Text) And IsDate(txt出生日期.Text) And mblnChange Then
        str出生日期 = Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS")
        txt年龄.Text = ReCalcOld(CDate(str出生日期), cbo年龄单位, , , CDate(txt出生日期.Tag))
    End If
End Sub

Private Sub txt出生时间_GotFocus()
    Call OpenIme
    SelAll txt出生时间
End Sub

Private Sub txt出生时间_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt出生日期) Then
        KeyAscii = 0
        txt出生时间.Text = "__:__"
    End If
End Sub

Private Sub txt出生时间_Validate(Cancel As Boolean)
    If txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        txt出生时间.SetFocus
        Cancel = True
    End If
End Sub


Private Sub txt单位电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt单位开户行_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt单位开户行, KeyAscii
End Sub

Private Sub txt单位开户行_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt单位邮编_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt单位帐号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt单位帐号, KeyAscii
End Sub

Private Sub txt工作单位_Change()
    If txt工作单位.Text = "" Then txt工作单位.Tag = ""
End Sub

Private Sub txt工作单位_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt工作单位.Text <> "" Then
            '问题27040 by lesfeng 对合约单位加上撤档时间的处理 '问题32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                    " Where 末级=1 And (编码 Like '" & gstrLike & txt工作单位.Text & "%'" & _
                    " Or 简码 Like '" & gstrLike & txt工作单位.Text & "%'" & _
                    " Or 名称 Like '" & gstrLike & txt工作单位.Text & "%')" & _
                    " and (撤档时间 IS NULL OR TO_CHAR(撤档时间, 'yyyy-MM-dd') = '3000-01-01') ", _
                    0, "单位", , txt工作单位.Text)
            If Not rsTmp Is Nothing Then
                txt工作单位.Text = rsTmp!名称
                txt工作单位.Tag = rsTmp!ID
                txt单位电话.Text = Trim(rsTmp!电话 & "")
                txt单位开户行.Text = Trim(rsTmp!开户银行 & "")
                txt单位帐号.Text = Trim(rsTmp!帐号 & "")
            Else
                txt工作单位.Tag = ""
            End If
        Else
            txt工作单位.Tag = ""
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt工作单位, KeyAscii
    End If
End Sub

Private Sub txt工作单位_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt户口地址_GotFocus()
    SelAll txt家庭地址
    Call OpenIme(gstrIme)
End Sub

Private Sub txt户口地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt户口地址.Text <> "" Then
            '问题32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select 编码 as ID,编码,名称,简码 From 地区" & _
                    " Where 编码 Like '" & gstrLike & txt户口地址.Text & "%'" & _
                    " Or 简码 Like '" & gstrLike & txt户口地址.Text & "%'" & _
                    " Or 名称 Like '" & gstrLike & txt户口地址.Text & "%'", _
                    0, "地区", , txt户口地址.Text)
            If Not rsTmp Is Nothing Then
                txt户口地址.Text = rsTmp!名称
                mblnSel = True
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt户口地址, KeyAscii
    End If
End Sub

Private Sub txt户口地址_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt户口地址邮编_GotFocus()
    SelAll txt户口地址邮编
End Sub

Private Sub txt户口地址邮编_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt籍贯_GotFocus()
    SelAll txt籍贯
    Call OpenIme(gstrIme)
End Sub

Private Sub txt籍贯_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt籍贯.Text <> "" Then
            Set rsTmp = GetArea(Me, txt籍贯)
            If Not rsTmp Is Nothing Then
                txt籍贯.Text = rsTmp!名称
                '问题27390 by lesfeng 2010-02-25
'                Call zlCommFun.PressKey(vbKeyTab)
            Else
                SelAll txt籍贯
                txt籍贯.SetFocus
            End If
        Else
            '问题27390 by lesfeng 2010-02-25
'            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt籍贯, KeyAscii
    End If
End Sub

Private Sub txt籍贯_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt家庭地址邮编_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt家庭地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt家庭地址.Text <> "" Then
            '问题32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select 编码 as ID,编码,名称,简码 From 地区" & _
                    " Where 编码 Like '" & gstrLike & txt家庭地址.Text & "%'" & _
                    " Or 简码 Like '" & gstrLike & txt家庭地址.Text & "%'" & _
                    " Or 名称 Like '" & gstrLike & txt家庭地址.Text & "%'", _
                    0, "地区", , txt家庭地址.Text)
            If Not rsTmp Is Nothing Then
                txt家庭地址.Text = rsTmp!名称
                mblnSel = True
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt家庭地址, KeyAscii
    End If
End Sub

Private Sub txt家庭地址_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt家庭电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub



Private Sub txt缴款单位_GotFocus()
    If IsNumeric(txt预交额.Text) And txt缴款单位.Text = "" Then
        txt缴款单位.Text = txt工作单位.Text
    End If
    zlControl.TxtSelAll txt缴款单位
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt缴款单位_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt缴款单位, KeyAscii
End Sub

Private Sub txt缴款单位_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt结算号码_GotFocus()
    zlControl.TxtSelAll txt结算号码
End Sub

Private Sub txt结算号码_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt结算号码, KeyAscii
End Sub

Private Sub txt卡额_KeyPress(KeyAscii As Integer)
    If txt卡额.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not mCurSendCard.rs卡费 Is Nothing Then
            If mCurSendCard.rs卡费!是否变价 = 1 Then
                If mCurSendCard.rs卡费!现价 <> 0 And Abs(CCur(txt卡额.Text)) > Abs(mCurSendCard.rs卡费!现价) Then
                    MsgBox IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡金额绝对值不能大于最高限价：" & Format(Abs(mCurSendCard.rs卡费!现价), "0.00"), vbExclamation, gstrSysName
                    If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus: Call zlControl.TxtSelAll(txt卡额): Exit Sub
                End If
                If mCurSendCard.rs卡费!原价 <> 0 And Abs(CCur(txt卡额.Text)) < Abs(mCurSendCard.rs卡费!原价) Then
                    MsgBox IIf(glngSys Like "8??", "会员", mCurSendCard.str卡名称) & "卡金额绝对值不能小于最低限价：" & Format(Abs(mCurSendCard.rs卡费!原价), "0.00"), vbExclamation, gstrSysName
                    If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus: Call zlControl.TxtSelAll(txt卡额): Exit Sub
                End If
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(txt卡额.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0:  Exit Sub
    End If
End Sub

Private Sub txt卡号_Change()
    Call SetCardEditEnabled
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    
    mbln是否扫描身份证 = False
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> 13 Then
        If Len(txt卡号.Text) = mCurSendCard.lng卡号长度 - 1 And KeyAscii <> 8 Then
            txt卡号.Text = txt卡号.Text & Chr(KeyAscii)
            
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf txt卡号.Text = "" Then
        KeyAscii = 0: cmdOK.SetFocus  '不发卡,直接跳过
    Else
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    
End Sub

Private Sub txt卡号_LostFocus()
    Call ReLoadCardFee
    Call CheckFreeCard(txt卡号.Text)
    Call SetBrushCardObject(False)
End Sub

Private Sub txt卡号_Validate(Cancel As Boolean)
    Dim lngPatientID As Long
    Dim lng变动类型 As Long
    Dim blnCardBind As Boolean  '卡是否进行绑定

    txt卡号.Text = Trim(txt卡号.Text)
    If mCurSendCard.lng卡号长度 = Len(Trim(txt卡号.Text)) Then
        If WhetherTheCardBinding(txt卡号.Text, mCurSendCard.lng卡类别ID, lngPatientID) Then
            If mCurSendCard.bln自制卡 And mCurSendCard.bln重复使用 And lngPatientID > 0 Then
                lng变动类型 = GetCardLastChangeType(txt卡号.Text, mCurSendCard.lng卡类别ID, lngPatientID)
                If lng变动类型 = 11 Then
                    '如果是绑定
                    If MsgBox("卡号为【" & txt卡号.Text & "】的{" & mCurSendCard.str卡名称 & "}的卡已经与病人标识为【" & lngPatientID & "】的进行了绑定！" & vbCrLf & "是否取消该卡的绑定?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Cancel = True
                        txt卡号.Text = ""
                        Exit Sub
                    End If
                    If BlandCancel(mCurSendCard.lng卡类别ID, Trim(txt卡号.Text), lngPatientID) Then
                        Exit Sub
                    End If
                End If
            End If

            MsgBox "该卡号已经被绑定,不能绑定该卡号.", vbInformation, gstrSysName
            Cancel = True
            txt卡号.Text = ""
            Exit Sub
        End If
    End If

End Sub

Private Sub CheckFreeCard(ByVal strCard As String)
'功能：对一卡通模式下的卡号，严格控制票号时，检查是否在票据领用范围内，范围之外的卡不收费
    
    If txt卡额.Visible = False Then Exit Sub
    
    If Not mCurSendCard.rs卡费 Is Nothing And Val(txt卡额.Text) = 0 Then  '先恢复
        txt卡额.Text = Format(IIf(mCurSendCard.rs卡费!是否变价 = 1, mCurSendCard.rs卡费!缺省价格, mCurSendCard.rs卡费!现价), "0.00")
        txt卡额.Tag = txt卡额.Text
    End If
    '142204:李南春，2020/6/18，检查领用ID时需要传入卡类别ID，并且是否根据费别打折仅跟是否屏蔽费别有关，与是否变价无关
    If mCurSendCard.blnOneCard And mCurSendCard.bln严格控制 Then
        mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), strCard, mCurSendCard.lng卡类别ID)
        If mCurSendCard.lng领用ID <= 0 Then txt卡额.Text = "0.00": txt卡额.Tag = txt卡额.Text
    End If

    If Not mCurSendCard.rs卡费 Is Nothing And Val(txt卡额.Text) <> 0 Then
        If Val(Nvl(mCurSendCard.rs卡费!屏蔽费别)) <> 1 Then
            txt卡额.Text = Format(GetActualMoney(NeedName(cbo费别.Text), mCurSendCard.rs卡费!收入项目ID, IIf(mCurSendCard.rs卡费!是否变价 = 1, mCurSendCard.rs卡费!缺省价格, mCurSendCard.rs卡费!现价), mCurSendCard.rs卡费!收费细目ID), "0.00")
            txt卡额.Tag = txt卡额.Text
        End If
    End If
End Sub

Private Sub cbo费别_Click()
     
    If Not mCurSendCard.rs卡费 Is Nothing And Val(txt卡额.Text) <> 0 And txt卡额.Visible And Trim(txt卡号.Text) <> "" Then
        If Val(Nvl(mCurSendCard.rs卡费!屏蔽费别)) <> 1 Then
            txt卡额.Text = Format(GetActualMoney(NeedName(cbo费别.Text), mCurSendCard.rs卡费!收入项目ID, IIf(mCurSendCard.rs卡费!是否变价 = 1, mCurSendCard.rs卡费!缺省价格, mCurSendCard.rs卡费!现价), mCurSendCard.rs卡费!收费细目ID), "0.00")
            txt卡额.Tag = txt卡额.Text
        End If
    End If
End Sub

Private Sub txt开户行_GotFocus()
    If IsNumeric(txt预交额.Text) And txt开户行.Text = "" Then
        txt开户行.Text = txt单位开户行.Text
    End If
    zlControl.TxtSelAll txt开户行
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt开户行_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt缴款单位, KeyAscii
End Sub

Private Sub txt开户行_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt联系人地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt联系人地址.Text <> "" Then
            '问题32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select 编码 as ID,编码,名称,简码 From 地区" & _
                    " Where 编码 Like '" & gstrLike & txt联系人地址.Text & "%'" & _
                    " Or 简码 Like '" & gstrLike & txt联系人地址.Text & "%'" & _
                    " Or 名称 Like '" & gstrLike & txt联系人地址.Text & "%'", _
                    0, "地区", , txt联系人地址.Text)
            If Not rsTmp Is Nothing Then
                txt联系人地址.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt联系人地址, KeyAscii
    End If
End Sub

Private Sub txt联系人地址_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt联系人电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt联系人电话_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人电话") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人电话")) = txt联系人电话.Text
    End If
End Sub

Private Sub txt联系人身份证_GotFocus()
    SelAll txt联系人身份证
End Sub

Private Sub txt联系人身份证_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt联系人身份证_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人身份证号") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人身份证号")) = txt联系人身份证.Text
    End If
End Sub

Private Sub txt联系人姓名_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt联系人姓名, KeyAscii
End Sub

Private Sub txt联系人姓名_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt联系人姓名_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人姓名") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人姓名")) = txt联系人姓名.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt联系人姓名.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt门诊号_Change()
    '问题27207 by lesfeng 2010-1-4
    lbl门诊号.Tag = "" '记录自动编号是否被人工修改
End Sub

Private Sub txt门诊号_GotFocus()
    SelAll txt门诊号
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8) & Chr(22), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt门诊号_Validate(Cancel As Boolean)
    If Val(txt门诊号.Text) = 0 And Val(txt门诊号.Tag) <> 0 Then txt门诊号.Text = txt门诊号.Tag
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            Call cbo年龄单位.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt年龄.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt年龄_Validate(Cancel As Boolean)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt区域_GotFocus()
    SelAll txt区域
    Call OpenIme(gstrIme)
End Sub

Private Sub txt区域_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt区域.Text <> "" Then
            Set rsTmp = GetArea(Me, txt区域)
            If Not rsTmp Is Nothing Then
                txt区域.Text = rsTmp!名称
                '问题27390 by lesfeng 2010-02-25
'                Call zlCommFun.PressKey(vbKeyTab)
            Else
                SelAll txt区域
                txt区域.SetFocus
            End If
        Else
            '问题27390 by lesfeng 2010-02-25
'            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt区域, KeyAscii
    End If
End Sub

Private Sub txt区域_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt身份证号_Change()
    Dim strTmp As String
    
    If mblnChange Then
        strTmp = GetIDDate(txt身份证号.Text)
        If IsDate(strTmp) And txt出生日期.Enabled = True Then txt出生日期.Text = strTmp
    End If
    If mbln扫描身份证签约 Then
        OpenIDCard txt身份证号.Text = ""
    End If
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    '问题号:53408
    mbln是否扫描身份证 = False

    Call Show绑定控件(mbln是否扫描身份证 And mbln扫描身份证签约)
    
    If zl当前用户身份证是否绑定(Val(IIf(Trim(txt病人ID.Text) = "", "0", Trim(txt病人ID.Text)))) = True Then
            MsgBox "当前用户的身份证号已经绑定，不允许修改其身份证号", vbInformation, gstrSysName
            KeyAscii = 0
    End If
    
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPatient_GotFocus()
    SelAll txtPatient
    Call OpenIme(gstrIme)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjICCard.SetEnabled (True)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txt年龄_GotFocus()
    Call zlCommFun.OpenIme
    SelAll txt年龄
End Sub

Private Sub txt出生日期_GotFocus()
    Call OpenIme
    SelAll txt出生日期
End Sub

Private Sub txt身份证号_GotFocus()
    SelAll txt身份证号
    '问题号:53408
    If mbln扫描身份证签约 = True Then
        Call OpenIDCard(txt身份证号.Text = "")
    End If
End Sub

Private Sub txt出生地点_GotFocus()
    SelAll txt出生地点
    Call OpenIme(gstrIme)
End Sub

Private Sub txt家庭地址_GotFocus()
    SelAll txt家庭地址
    Call OpenIme(gstrIme)
End Sub

Private Sub txt家庭地址邮编_GotFocus()
    SelAll txt家庭地址邮编
End Sub

Private Sub txt家庭电话_GotFocus()
    SelAll txt家庭电话
End Sub

Private Sub txt联系人姓名_GotFocus()
    SelAll txt联系人姓名
    Call OpenIme(gstrIme)
End Sub

Private Sub txt联系人地址_GotFocus()
    SelAll txt联系人地址
    Call OpenIme(gstrIme)
End Sub

Private Sub txt联系人电话_GotFocus()
    SelAll txt联系人电话
End Sub

Private Sub txt工作单位_GotFocus()
    SelAll txt工作单位
    Call OpenIme(gstrIme)
End Sub

Private Sub txt单位电话_GotFocus()
    SelAll txt单位电话
End Sub

Private Sub txt单位邮编_GotFocus()
    SelAll txt单位邮编
End Sub

Private Sub txt单位开户行_GotFocus()
    SelAll txt单位开户行
    Call OpenIme(gstrIme)
End Sub

Private Sub txt卡号_GotFocus()
    SelAll txt卡号
    Call SetBrushCardObject(True)
End Sub
Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开身份证读卡器
    '编制:王吉
    '日期:2012-08-31 16:28:23
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '初始化对卡对象
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '打开读卡器
    mobjIDCard.SetEnabled (blnEnabled)
End Sub
Private Sub txtPass_GotFocus()
    SelAll txtPass
    OpenPassKeyboard txtPass, False
End Sub

Private Sub txt卡额_GotFocus()
    SelAll txt卡额
End Sub

Private Sub txt单位帐号_GotFocus()
    SelAll txt单位帐号
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo性别.Locked = True Then Exit Sub
    If SendMessage(cbo性别.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo性别.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo性别.ListIndex = lngIdx
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo费别.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo费别.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo费别.ListIndex = lngIdx
End Sub

Private Sub cbo国籍_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo国籍.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo国籍.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo国籍.ListIndex = lngIdx
End Sub

Private Sub cbo民族_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo民族.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo民族.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo民族.ListIndex = lngIdx
End Sub

Private Sub cbo学历_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo学历.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo学历.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo学历.ListIndex = lngIdx
End Sub

Private Sub cbo婚姻状况_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo婚姻状况.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo婚姻状况.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo婚姻状况.ListIndex = lngIdx
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo职业.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo职业.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo职业.ListIndex = lngIdx
End Sub

Private Sub cbo身份_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo身份.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo身份.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo身份.ListIndex = lngIdx
End Sub

Private Sub cbo联系人关系_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo联系人关系.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo联系人关系.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo联系人关系.ListIndex = lngIdx
End Sub

Private Sub cbo结算方式_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo结算方式.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo结算方式.hWnd, KeyAscii)
    If lngIdx <> -2 Then
        cbo结算方式.ListIndex = lngIdx
        Call cbo联系人关系_Click
    End If
End Sub

Private Function CheckMCOutMode(ByVal strMCCode As String) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From 保险类别 Where 外挂=1 And 序号=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMCCode)

    CheckMCOutMode = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'问题27390 by lesfeng 2010-01-25
Private Sub InitInputTabStop()
'功能：根据本地设置光标要定位的输入项目
    Dim i As Integer
    cbo国籍.TabStop = zlDatabase.GetPara("国籍", glngSys, mlngModul, 1)
    cbo民族.TabStop = zlDatabase.GetPara("民族", glngSys, mlngModul, 1)
    cbo学历.TabStop = zlDatabase.GetPara("学历", glngSys, mlngModul, 1)
    cbo婚姻状况.TabStop = zlDatabase.GetPara("婚姻状况", glngSys, mlngModul, 1)
    cbo职业.TabStop = zlDatabase.GetPara("职业", glngSys, mlngModul, 1)
    cbo身份.TabStop = zlDatabase.GetPara("身份", glngSys, mlngModul, 1)
    txt出生日期.TabStop = zlDatabase.GetPara("出生日期", glngSys, mlngModul, 1)
    txt出生时间.TabStop = zlDatabase.GetPara("出生日期", glngSys, mlngModul, 1)
    txt身份证号.TabStop = zlDatabase.GetPara("身份证号", glngSys, mlngModul, 1)
    cboIDNumber.TabStop = zlDatabase.GetPara("身份证号", glngSys, mlngModul, 1)
    txt家庭地址邮编.TabStop = zlDatabase.GetPara("家庭地址邮编", glngSys, mlngModul, 1)
    txt家庭电话.TabStop = zlDatabase.GetPara("家庭电话", glngSys, mlngModul, 1)
    txt户口地址邮编.TabStop = zlDatabase.GetPara("户口地址邮编", glngSys, mlngModul, 1)
    txt联系人姓名.TabStop = zlDatabase.GetPara("联系人姓名", glngSys, mlngModul, 1)
    cbo联系人关系.TabStop = zlDatabase.GetPara("联系人关系", glngSys, mlngModul, 1)
    txt联系人电话.TabStop = zlDatabase.GetPara("联系人电话", glngSys, mlngModul, 1)
    txt联系人身份证.TabStop = zlDatabase.GetPara("联系人身份证号", glngSys, mlngModul, 1)
    txt工作单位.TabStop = zlDatabase.GetPara("工作单位", glngSys, mlngModul, 1)
    txt单位电话.TabStop = zlDatabase.GetPara("单位电话", glngSys, mlngModul, 1)
    txt单位邮编.TabStop = zlDatabase.GetPara("单位邮编", glngSys, mlngModul, 1)
    txt单位开户行.TabStop = zlDatabase.GetPara("单位开户行", glngSys, mlngModul, 1)
    txt单位帐号.TabStop = zlDatabase.GetPara("单位帐号", glngSys, mlngModul, 1)
    txt其他证件.TabStop = zlDatabase.GetPara("其他证件", glngSys, mlngModul, 1)
    txt区域.TabStop = zlDatabase.GetPara("区域", glngSys, mlngModul, 1)
    
    If gbln启用结构化地址 Then
        PatiAddress(E_IX_出生地点).TabStop = zlDatabase.GetPara("出生地点", glngSys, mlngModul, 1)
        PatiAddress(E_IX_籍贯).TabStop = zlDatabase.GetPara("籍贯", glngSys, mlngModul, 1)
        PatiAddress(E_IX_现住址).TabStop = zlDatabase.GetPara("现住址", glngSys, mlngModul, 1)
        PatiAddress(E_IX_户口地址).TabStop = zlDatabase.GetPara("户口地址", glngSys, mlngModul, 1)
        PatiAddress(E_IX_联系人地址).TabStop = zlDatabase.GetPara("联系人地址", glngSys, mlngModul, 1)
    Else
        txt出生地点.TabStop = zlDatabase.GetPara("出生地点", glngSys, mlngModul, 1)
        txt籍贯.TabStop = zlDatabase.GetPara("籍贯", glngSys, mlngModul, 1)
        txt家庭地址.TabStop = zlDatabase.GetPara("现住址", glngSys, mlngModul, 1)
        txt户口地址.TabStop = zlDatabase.GetPara("户口地址", glngSys, mlngModul, 1)
        txt联系人地址.TabStop = zlDatabase.GetPara("联系人地址", glngSys, mlngModul, 1)
    End If
End Sub

Private Sub InitCard()
'功能：根据入口参数设置卡片状态
    Dim i As Long, arrTmp As Variant
    
    Call InitvsDrug
    Call InitVsInoculate
    Call InitVsOtherInfo
    Call InitCombox
    '结构化地址
    Call InitStructAddress
    
    If mbytInState <> 2 Then
        txtPatient.MaxLength = GetColumnLength("病人信息", "姓名")
        txt年龄.MaxLength = GetColumnLength("病人信息", "年龄")
        txt门诊号.MaxLength = GetColumnLength("病人信息", "门诊号")
        txt住院号.MaxLength = GetColumnLength("病人信息", "住院号")
    End If
    
    '问题27390 by lesfeng 2010-01-25
    Call InitInputTabStop
    
    If InStr(mstrPrivs, "合约病人登记") = 0 Then
        txt工作单位.Enabled = False
        txt工作单位.BackColor = &H8000000F
        txt单位电话.Enabled = False
        txt单位电话.BackColor = &H8000000F
        txt单位邮编.Enabled = False
        txt单位邮编.BackColor = &H8000000F
        txt单位开户行.Enabled = False
        txt单位开户行.BackColor = &H8000000F
        txt单位帐号.Enabled = False
        txt单位帐号.BackColor = &H8000000F
        cmd合同单位.Visible = False
    End If
    
    cbo病人类型.Enabled = InStr(mstrPrivs, "调整病人类型") > 0
    txt门诊号.Enabled = InStr(mstrPrivs, ";允许修改门诊号;") > 0

    mlngOutModeMC = 0
    arrTmp = Split(GetSetting("ZLSOFT", "公共全局", "本地支持的医保", ""), ",")
    For i = 0 To UBound(arrTmp)
        If IsNumeric(arrTmp(i)) Then
            If CheckMCOutMode(arrTmp(i)) Then mlngOutModeMC = Val(arrTmp(i)): Exit For
        End If
    Next
    
    If mlngOutModeMC = 920 Then
        txtPatiMCNO(0).MaxLength = 12
    Else
        txtPatiMCNO(0).MaxLength = 30
    End If
    txtPatiMCNO(0).ToolTipText = "最大长度" & txtPatiMCNO(0).MaxLength & "位"
    txtPatiMCNO(1).MaxLength = txtPatiMCNO(0).MaxLength
    If mlngOutModeMC = 0 Or mbytInState = E查阅 Then
        txtPatiMCNO(1).Visible = False
        lblPatiMCNO(1).Visible = False
    End If
    
    Call InitDicts
    If cbo费别.ListCount = 0 Then
        MsgBox "没有设置费别信息,请先到费别等级设置中设置！", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    IDKind.Enabled = mbytInState = E新增
    Select Case mbytInState
        Case E新增   '新增
            If Not gobjSquare.objSquareCard Is Nothing Then
                IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
            Set mobjIDCard = New clsIDCard
            Set mobjICCard = New clsICCard
            Call mobjIDCard.SetParent(Me.hWnd)
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
            Call InitPrepayType: Call InitSendCardPreperty
            chk记帐.Value = IIf(gbln记账 = True, 1, 0)
            chk记帐.Tag = IIf(chk记帐.Value = 1, 1, 0)
            '问题27207 by lesfeng 2010-1-4
            txt病人ID.Text = zlDatabase.GetNextNo(1): lbl病人ID.Tag = txt病人ID.Text
            
            cmdYB.Left = lbl性别.Left - lbl性别.Width
            If Not glngSys Like "8??" Then txt门诊号.Text = zlDatabase.GetNextNo(3): lbl门诊号.Tag = txt门诊号.Text
            '74299:刘鹏飞,2014-07-03,病人信息也可以进行病人类型设置
            '新增时病人类型不可见
            'lblPatiType.Visible = False: cbo病人类型.Visible = False: lblPatiColor.Visible = False
            Call Load支付方式
            '89980病人结构化 新增病人设置缺省值
            If gbln启用结构化地址 Then
                Call LoadStructAddressDef(marrAddress)
                Call SetStrutAddress(2)
            End If
        Case E修改  '修改
            If Not gobjSquare.objSquareCard Is Nothing Then
                IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
        
            If Not glngSys Like "8??" Then
                lbl住院号.Visible = True
                txt住院号.Visible = True
                '问题27351 by lesfeng 2010-01-12
                txt备注.Visible = True
                lbl备注.Visible = True
                txt备注.Tag = "1"  '标记是否显示备注
                cmdYB.Visible = False
            End If
            
            If Not ReadPatiCard(mlng病人ID) Then
                If glngSys Like "8??" Then
                    MsgBox "客户信息读取失败！", vbExclamation, gstrSysName
                Else
                    MsgBox "病人信息读取失败！", vbExclamation, gstrSysName
                End If
                mblnUnLoad = True: Exit Sub
            End If
            Call EMPI_LoadPati
        Case E查阅  '查看
            fraInfo.Enabled = False
            PicHealth.Enabled = False
            cmdOK.Visible = False
            cboIDNumber.Locked = True
            cmdCancel.Caption = "退出(&X)"
            
            If Not ReadPatiCard(mlng病人ID) Then
                If glngSys Like "8??" Then
                    MsgBox "客户信息读取失败！", vbExclamation, gstrSysName
                Else
                    MsgBox "病人信息读取失败！", vbExclamation, gstrSysName
                End If
                mblnUnLoad = True: Exit Sub
            End If
    End Select
    
    '界面调整
    If mbytInState <> E新增 Then '修改和查看都不显示预交款和发卡界面
        fraDeposit.Visible = False: cmdOperation(OPT.C0预交款).Visible = False
        fraCard.Visible = False: cmdOperation(OPT.C1就诊卡).Visible = False
        Me.Height = Me.Height - fraDeposit.Height
        Me.Height = Me.Height - fraCard.Height
        mPageHeight.基本 = Me.Height
    End If
    '是否显示备注
    txt备注.Visible = (Val(txt备注.Tag) = 1)
    lbl备注.Visible = (Val(txt备注.Tag) = 1)
End Sub

Private Sub ClearCard()
    mblnEMPI = False
    mlngPatientID = 0
    '55251:刘鹏飞,2012-10-26
    mlng病人ID = 0: mlng主页ID = 0
    mblnICCard = False
    mstrYBPati = ""
    
    txt门诊号.Text = ""
    txt住院号.Text = ""
    txtPatient.Text = ""
    '对病人姓名、性别、出生日期、年龄的解锁
    txtPatient.Locked = False
    txtPatient.BackColor = &H80000005
    cbo性别.Locked = False
    cbo性别.BackColor = txtPatient.BackColor
    txt出生日期.Enabled = True
    txt出生日期.BackColor = txtPatient.BackColor
    txt出生日期.Tag = "0"
    txt出生时间.Enabled = True
    txt出生时间.BackColor = txtPatient.BackColor
    txt年龄.Locked = False
    txt年龄.BackColor = txtPatient.BackColor
    cbo年龄单位.Locked = False
    cbo年龄单位.BackColor = txtPatient.BackColor
    txtPatiMCNO(0).Text = "": txtPatiMCNO(0).Tag = "": txtPatiMCNO(1).Text = ""
    
    txt年龄.Text = "": Call txt年龄_Validate(False)
    txt出生日期.Text = "____-__-__"
    txt出生时间.Text = "__:__"
    txt身份证号.Text = ""
    txt其他证件.Text = ""
    txt出生地点.Text = ""
    txt家庭地址.Text = ""
    txt家庭地址邮编.Text = ""
    txt家庭电话.Text = ""
    txt户口地址.Text = ""
    txt户口地址邮编.Text = ""
    txt籍贯.Text = ""
    txt区域.Text = ""
    txt联系人姓名.Text = ""
    txt联系人地址.Text = ""
    txt联系人电话.Text = ""
    txt联系人身份证.Text = ""
    txt工作单位.Text = "": txt工作单位.Tag = ""
    txt工作单位.Text = ""
    txt单位电话.Text = ""
    txt单位邮编.Text = ""
    txt单位开户行.Text = ""
    txt单位帐号.Text = ""
    txt卡号.Text = ""
    txtPass.Text = ""
    txtAudi.Text = ""
    txtMobile.Text = ""
    '问题27351 by lesfeng 2010-01-12
    txt备注.Text = ""
    
    If gbln启用结构化地址 Then
        Call SetStrutAddress(1)
        Call SetStrutAddress(2)
    End If
    
    chk记帐.Value = IIf(gbln记账 = True, 1, 0)
    cboIDNumber.ListIndex = -1 '缺省
    cboIDNumber.Enabled = True
    Call SetCboDefault(cbo性别)
    Call SetCboDefault(cbo费别)
    Call SetCboDefault(cbo医疗付款)
    Call SetCboDefault(cbo国籍)
    Call SetCboDefault(cbo民族)
    Call SetCboDefault(cbo学历)
    Call SetCboDefault(cbo婚姻状况)
    Call SetCboDefault(cbo职业)
    Call SetCboDefault(cbo身份)
    Call SetCboDefault(cbo联系人关系)
    Call SetCboDefault(cbo结算方式)
    Call SetCboDefault(cbo病人类型)
    '74299:刘鹏飞,2014-07-03,病人信息也可以进行病人类型设置
    '新增病人时不可见
    'If mbytInState = E新增 Then lblPatiType.Visible = False: cbo病人类型.Visible = False: lblPatiColor.Visible = False
    '预交信息
    txt预交额.Text = ""
    txt缴款单位.Text = ""
    txt帐号.Text = ""
    txt开户行.Text = ""
    txt结算号码.Text = ""
    '问题号:51072
    txt联系人身份证.Text = ""
    '问题号:53408
    txt支付密码.Text = ""
    txt验证密码.Text = ""
    txt验证密码.Tag = ""
    txt支付密码.Enabled = False
    txt验证密码.Enabled = False
    lbl支付密码.Enabled = False
    lbl验证密码.Enabled = False
    
    mlng图像操作 = 0: mstr采集图片 = ""
    imgPatient.Picture = Nothing
    '问题号:56599
    Call Clear健康档案
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    Dim lngIndex As Long
    
    If IDKind.GetCurCard.名称 = "门诊号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    If mlngPatientID <> 0 Then Exit Sub
        
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Or IDKind.IDKind = IDKind.GetKindIndex("住院号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    '55571:刘鹏飞,2012-11-12
    txtPatient.IMEMode = 0
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        strInput = txtPatient.Text
        
        Call FindPati(IDKind.GetCurCard, blnCard, strInput)
        Call EMPI_LoadPati
        Call ReLoadCardFee(True)
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String, Optional lngPatientIDRef As Long = 0)
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:查找病人
'编制:刘鹏飞
'日期:2012-10-25
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean
    Dim lngPatientID As Long, lngIndex As Long
    
    If objCard.名称 Like "IC卡*" And objCard.系统 = True Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    '读取病人信息
    lngPatientID = GetPatient(objCard, strInput, blnCard)
    lngPatientIDRef = lngPatientID
    If lngPatientID <> 0 Then
        Call ClearCard
        mlngPatientID = lngPatientID
        txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
        Call ReadPatiCard(mlngPatientID)
    Else
        If (blnICCard Or blnCard) And fraCard.Visible Then '发新卡
            MsgBox "该卡没有建档,将作为新卡登记,请输入病人姓名。", vbInformation, gstrSysName
            txt卡号.Text = strInput
            lngIndex = IDKind.GetKindIndex("姓名")
            txtPatient.Text = "": txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            If lngIndex >= 0 Then IDKind.IDKind = lngIndex
            Call CheckFreeCard(txt卡号.Text)
            
        ElseIf Not (IDKind.GetCurCard.名称 Like "姓名*" And InStr("+-*", Left(strInput, 1)) = 0) Then
           txtPatient.Text = "": txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
           MsgBox "没有找到指定的病人。", vbInformation, gstrSysName
        End If
    End If
    Call zlControl.TxtSelAll(txtPatient)
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    strSQL = "Select A.病人ID From 病人信息 A Where A.停用时间 is NULL "
    
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then
            '按手机号查询
            If IDKind.IsMobileNo(strInput) = False Then GoTo NotFoundPati:
            If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Function
        End If
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strSQL = strSQL & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
    Else
        Select Case objCard.名称
            Case "姓名"
                '输入姓名当成新病人
                Exit Function
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
            Case Else
                If InStr(",身份证号,身份证,二代身份证,", "," & objCard.名称 & ",") > 0 Then
                    If Not CreatePublicPatient() Then Exit Function
                    lng病人ID = gobjPublicPatient.GetPatiByID(glngModul, strInput)
                End If
                If lng病人ID = 0 Then
                    '其他类别的,获取相关的病人ID
                    If Val(objCard.接口序号) > 0 Then
                        lng卡类别ID = Val(objCard.接口序号)
                        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                        If lng病人ID = 0 Then GoTo NotFoundPati:
                    Else
                        If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                            strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    End If
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If rsTmp.RecordCount > 0 Then GetPatient = rsTmp!病人ID
    mblnICCard = IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Function
NotFoundPati:
End Function

Private Function ReadPatiCard(ByVal lng病人ID As Long) As Integer
'功能：修改或查看时,读取指定病人信息,并显示在界面上
'返回：
'     -1=成功
'      0=失败
'      1=该病人不存在
    Dim rsTmp As New ADODB.Recordset
    Dim str费别 As String
    '问题27351 by lesfeng 2010-01-12
    On Error GoTo errH
    '问题号:51071
    gstrSQL = "Select A.门诊号,A.住院号,A.主页ID 就诊次数,A.姓名,A.性别,A.费别,A.医疗付款方式,A.国籍,A.民族,A.区域,A.学历,A.婚姻状况," & _
        " A.职业,A.身份,Decode(nvl(A.在院,0),0,A.年龄,B.年龄) as 年龄,A.出生日期,A.身份证号,A.手机号,A.出生地点,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.户口地址,A.户口地址邮编,A.籍贯,A.担保人,A.担保额,A.担保性质," & _
        " A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.工作单位,A.合同单位ID,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.联系人身份证号," & _
        " B.病人ID,B.费别 as 住院费别,Nvl(B.险类,A.险类) as 险类,Nvl(A.医保号,D.信息值) as 医保号,A.其他证件," & _
        IIf(mstrYBPati = "", " NVL(Decode(B.病人ID,Null,A.病人类型,B.病人类型),Decode(A.险类,Null,'普通病人','医保病人'))", "zl_PatiType(A.病人ID)") & " 病人类型,B.备注,B.入院日期,B.出院日期 " & _
        " From 病人信息 A,病案主页 B,病案主页从表 D" & _
        " Where A.病人ID=B.病人ID(+) And Nvl(A.主页ID,0)=B.主页ID(+)" & _
        " And A.病人ID=D.病人ID(+) And Nvl(A.主页ID,0)=D.主页ID(+) And D.信息名(+)='医保号' And A.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID)
    
    If rsTmp.RecordCount = 0 Then ReadPatiCard = 1: Exit Function
    
    mlngPatientID = lng病人ID
    
    mlng险类 = Nvl(rsTmp!险类, 0)
    txt病人ID.Text = lng病人ID
    '55251,刘鹏飞，2012-10-26
    mlng病人ID = Val(txt病人ID.Text)
    txt门诊号.Text = Nvl(rsTmp!门诊号)
    txt门诊号.Tag = Nvl(rsTmp!门诊号)
    txt住院号.Text = Nvl(rsTmp!住院号)
    txt住院号.Tag = Nvl(rsTmp!住院号)
    txtPatient.Text = rsTmp!姓名
    '问题号:51071
    txt联系人身份证.Text = Nvl(rsTmp!联系人身份证号)
    If mbytInState = E修改 Then
        '外挂医保,或在院非真实医保病人可以修改医保号
        txtPatiMCNO(0).Enabled = mlngOutModeMC > 0 Or Not IsNull(rsTmp!就诊次数) And IsNull(rsTmp!险类)
        
        txtPatiMCNO(0).Text = "" & rsTmp!医保号 '最大长度自动截断超长字符S
        txtPatiMCNO(0).Tag = txtPatiMCNO(0).Text
        If mlngOutModeMC > 0 Then txtPatiMCNO(1).Text = txtPatiMCNO(0).Text
    Else
        txtPatiMCNO(0).Text = Nvl(rsTmp!医保号)
    End If
    
    cbo性别.ListIndex = GetCboIndex(cbo性别, Nvl(rsTmp!性别))
    If cbo性别.ListIndex = -1 And Not IsNull(rsTmp!性别) Then
        cbo性别.AddItem rsTmp!性别, 0
        cbo性别.ListIndex = cbo性别.NewIndex
    End If
       
    Call LoadOldData("" & rsTmp!年龄, txt年龄, cbo年龄单位)
    mblnChange = False
    txt出生日期.Text = Format(IIf(IsNull(rsTmp!出生日期), "____-__-__", rsTmp!出生日期), "YYYY-MM-DD")
    mblnChange = True
    
    If rsTmp!出院日期 & "" = "" And rsTmp!入院日期 & "" <> "" Then
        txt出生日期.Tag = rsTmp!入院日期 & ""
    Else
        txt出生日期.Tag = "0"
    End If
    
    If Not IsNull(rsTmp!出生日期) Then
        If mbytInState <> 2 And mbytInState <> 1 Then txt年龄.Text = ReCalcOld(CDate(Format(rsTmp!出生日期, "YYYY-MM-DD HH:MM:SS")), cbo年龄单位, lng病人ID, , CDate(txt出生日期.Tag)) '修改的时候,根据出生日期重算年龄
        
        If CDate(txt出生日期.Text) - CDate(rsTmp!出生日期) <> 0 Then
            mblnChange = False
            txt出生时间.Text = Format(rsTmp!出生日期, "HH:MM")
            mblnChange = True
        End If
    Else
        txt出生时间.Text = "__:__"
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
        
    mblnChange = False          '修改和查看时,身份证号与出生日期独立
    txt身份证号.Text = Nvl(rsTmp!身份证号)
    mblnChange = True
    cboIDNumber.Enabled = txt身份证号.Text = ""
    '根据不同查看方式读取不同的费别
    If mbytInState = E新增 Then
        str费别 = Nvl(rsTmp!费别)
    Else
        If mbytView = 1 Or mbytView = 2 Then
            str费别 = Nvl(rsTmp!住院费别)
        Else
            str费别 = Nvl(rsTmp!费别)
        End If
    End If
    
    cbo费别.ListIndex = GetCboIndex(cbo费别, str费别)
    If cbo费别.ListIndex = -1 And str费别 <> "" Then
        cbo费别.AddItem str费别, 0
        cbo费别.ListIndex = cbo费别.NewIndex
    End If
    
    cbo医疗付款.ListIndex = GetCboIndex(cbo医疗付款, Nvl(rsTmp!医疗付款方式))
    If cbo医疗付款.ListIndex = -1 And Not IsNull(rsTmp!医疗付款方式) Then
        cbo医疗付款.AddItem rsTmp!医疗付款方式, 0
        cbo医疗付款.ListIndex = cbo医疗付款.NewIndex
    End If
    
    cbo国籍.ListIndex = GetCboIndex(cbo国籍, Nvl(rsTmp!国籍))
    If cbo国籍.ListIndex = -1 And Not IsNull(rsTmp!国籍) Then
        cbo国籍.AddItem rsTmp!国籍, 0
        cbo国籍.ListIndex = cbo国籍.NewIndex
    End If
    
    cbo民族.ListIndex = GetCboIndex(cbo民族, Nvl(rsTmp!民族))
    If cbo民族.ListIndex = -1 And Not IsNull(rsTmp!民族) Then
        cbo民族.AddItem rsTmp!民族, 0
        cbo民族.ListIndex = cbo民族.NewIndex
    End If
    
    txt区域.Text = Nvl(rsTmp!区域)
    
    cbo病人类型.ListIndex = GetCboIndex(cbo病人类型, Nvl(rsTmp!病人类型, "普通病人"))
    cbo病人类型.Enabled = InStr(mstrPrivs, "调整病人类型") > 0
    lblPatiType.Visible = True: cbo病人类型.Visible = True: lblPatiColor.Visible = True
    
    cbo学历.ListIndex = GetCboIndex(cbo学历, Nvl(rsTmp!学历))
    If cbo学历.ListIndex = -1 And Not IsNull(rsTmp!学历) Then
        cbo学历.AddItem rsTmp!学历, 0
        cbo学历.ListIndex = cbo学历.NewIndex
    End If
    
    cbo婚姻状况.ListIndex = GetCboIndex(cbo婚姻状况, Nvl(rsTmp!婚姻状况))
    If cbo婚姻状况.ListIndex = -1 And Not IsNull(rsTmp!婚姻状况) Then
        cbo婚姻状况.AddItem rsTmp!婚姻状况, 0
        cbo婚姻状况.ListIndex = cbo婚姻状况.NewIndex
    End If
    
    cbo职业.ListIndex = GetCboIndex(cbo职业, Nvl(rsTmp!职业))
    If cbo职业.ListIndex = -1 And Not IsNull(rsTmp!职业) Then
        cbo职业.AddItem rsTmp!职业, 0
        cbo职业.ListIndex = cbo职业.NewIndex
    End If
    
    cbo身份.ListIndex = GetCboIndex(cbo身份, Nvl(rsTmp!身份))
    If cbo身份.ListIndex = -1 And Not IsNull(rsTmp!身份) Then
        cbo身份.AddItem rsTmp!身份, 0
        cbo身份.ListIndex = cbo身份.NewIndex
    End If
         
    txt家庭电话.Text = Nvl(rsTmp!家庭电话)
    txt家庭地址邮编.Text = Nvl(rsTmp!家庭地址邮编)
    txt户口地址邮编.Text = Nvl(rsTmp!户口地址邮编)
    
    '担保信息暂存于此，界面不显示，但修改保存时需要
    txt联系人姓名.Tag = Nvl(rsTmp!担保人)
    txt联系人电话.Tag = Nvl(rsTmp!担保额, 0)
    txt联系人地址.Tag = Nvl(rsTmp!担保性质, 0)
    txt联系人姓名.Text = Nvl(rsTmp!联系人姓名)
    
    cbo联系人关系.ListIndex = GetCboIndex(cbo联系人关系, Nvl(rsTmp!联系人关系))
    If cbo联系人关系.ListIndex = -1 And Not IsNull(rsTmp!联系人关系) Then
        cbo联系人关系.AddItem rsTmp!联系人关系, 0
        cbo联系人关系.ListIndex = cbo联系人关系.NewIndex
    End If

    txt联系人电话.Text = Nvl(rsTmp!联系人电话)
    txt联系人身份证.Text = Nvl(rsTmp!联系人身份证号)
    txt工作单位.Text = Nvl(rsTmp!工作单位)
    txt工作单位.Tag = Nvl(rsTmp!合同单位ID)
    txt单位电话.Text = Nvl(rsTmp!单位电话)
    txt单位邮编.Text = Nvl(rsTmp!单位邮编)
    txt单位开户行.Text = Nvl(rsTmp!单位开户行)
    txt单位帐号.Text = Nvl(rsTmp!单位帐号)
    txt其他证件.Text = "" & rsTmp!其他证件
    txtMobile.Text = Nvl(rsTmp!手机号)
    '问题27351 by lesfeng 2010-01-12
    If Nvl(rsTmp!就诊次数, 0) = 0 Then
        txt备注.Visible = False
        lbl备注.Visible = False
        txt备注.Tag = "0"
    Else
        mlng主页ID = rsTmp!就诊次数
        txt备注.Tag = "1"
    End If
    txt备注.Text = IIf(IsNull(rsTmp!备注), "", rsTmp!备注)
    
    If gbln启用结构化地址 Then
        Call ReadStructAddress(mlng病人ID, mlng主页ID, PatiAddress)
        txt出生地点.Text = PatiAddress(E_IX_出生地点).Value
        txt籍贯.Text = PatiAddress(E_IX_籍贯).Value
        txt家庭地址.Text = PatiAddress(E_IX_现住址).Value
        txt户口地址.Text = PatiAddress(E_IX_户口地址).Value
        txt联系人地址.Text = PatiAddress(E_IX_联系人地址).Value
    Else
        txt出生地点.Text = Nvl(rsTmp!出生地点)
        txt籍贯.Text = Nvl(rsTmp!籍贯)
        txt家庭地址.Text = Nvl(rsTmp!家庭地址)
        txt户口地址.Text = Nvl(rsTmp!户口地址)
        txt联系人地址.Text = Nvl(rsTmp!联系人地址)
    End If
    '74299:
'    If IsNull(rsTmp!病人ID) Then
'         lblPatiType.Visible = False: cbo病人类型.Visible = False: lblPatiColor.Visible = False
'    End If
    '74421,刘鹏飞,2014-07-04,读取病人照片信息
    Call ReadPatPricture(lng病人ID)
    '问题号:56599
    Call Load健康卡相关信息(lng病人ID)
    
    '不允许修改病人姓名、性别、出生日期、年龄
    txtPatient.Locked = True
    txtPatient.BackColor = &H80000016
    cbo性别.Locked = True
    cbo性别.BackColor = txtPatient.BackColor
    txt出生日期.Enabled = False
    txt出生日期.BackColor = txtPatient.BackColor
    txt出生时间.Enabled = False
    txt出生时间.BackColor = txtPatient.BackColor
    txt年龄.Locked = True
    txt年龄.BackColor = txtPatient.BackColor
    cbo年龄单位.Locked = True
    cbo年龄单位.BackColor = txtPatient.BackColor
    ' 读取从表信息
    Set rsTmp = Get病人信息从表(lng病人ID, "身份证号状态")
    rsTmp.Filter = "信息名='身份证号状态'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!信息值) Then
            Call cbo.Locate(cboIDNumber, zlCommFun.GetNeedName(rsTmp!信息值 & ""))
        End If
    End If
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) <> "中国" And Trim(txt身份证号.Text) = "" Then
        If Trim(zlCommFun.GetNeedName(cboIDNumber.Text)) = "" Then
            Set rsTmp = Get病人信息从表(lng病人ID, "外籍身份证号")
            rsTmp.Filter = "信息名='外籍身份证号'"
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!信息值) Then
                    txt身份证号.Text = "" & rsTmp!信息值
                End If
            End If
        End If
    End If
    ReadPatiCard = -1
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadPatPricture(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '74421,刘鹏飞,2014-07-04,读取病人照片信息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTmp As String
    Dim rsData As Recordset
    
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    mstr采集图片 = ""
    strSQL = "Select 病人id,照片 From 病人照片 Where 病人id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rsData.BOF = False Then
        strTmp = zlDatabase.ReadPicture(rsData, "照片", strTmp)
        mstr采集图片 = strTmp
        imgPatient.Picture = LoadPicture(strTmp)
        
        ReadPatPricture = True
        If strTmp <> "" Then Kill strTmp
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载身份证图像
    '编制:刘鹏飞
    '日期:2014-07-04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    mlng图像操作 = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SavePatPicture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:lng病人ID - 病人ID
    '74421,刘鹏飞,2014-07-04,读取病人照片信息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
    Dim strFile As String, strSQL As String
    
    On Error GoTo ErrHand
    Select Case mlng图像操作
        Case 1 '文件
            strFile = cmdialog.FileName
        Case 2 '采集
            strFile = mstr采集图片
            mstr采集图片 = ""
        Case 4 '二代身份证
            strFile = App.Path & "\SFZIMG.bmp"
            SavePicture imgPatient.Picture, strFile
    End Select
    If InStr(1, ",1,2,4,", "," & mlng图像操作 & ",") <> 0 Then
        gcnOracle.Execute "Delete From 病人照片 Where 病人id=" & lng病人ID
        strSQL = "Select 病人id,照片 From 病人照片 where 病人id=" & lng病人ID
        
        If strFile = "" Then Exit Sub
        rs.Open strSQL, gcnOracle, adOpenStatic, adLockOptimistic
        If rs.BOF Then
            If rs.EOF Then rs.AddNew
            rs("病人id").Value = lng病人ID
            rs("照片").Value = Null
            rs.Update
            
            If zlDatabase.SavePicture(strFile, rs, "照片") = False Then
                MsgBox "保存照片失败,文件可能被删除!", vbInformation, gstrSysName
                Exit Sub
            End If
            rs.Close
        End If
    ElseIf mlng图像操作 = 3 Then
        strSQL = strSQL & "Zl_病人照片_Delete("
        strSQL = strSQL & lng病人ID & ")"
        
        zlDatabase.ExecuteProcedure strSQL, "Zl_病人照片_Delete"
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddCardDataSQL(ByVal lng病人ID As Long, ByVal dtCurdate As Date, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:就诊卡发放处理
    '入参:lng病人ID
    '编制:刘兴洪
    '日期:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt操作类型 As Byte, strNO As String, strPassWord As String, strSQL As String
    Dim str原卡号 As String, str年龄 As String, strCard As String, str变动原因 As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str结算方式 As String, strBrushCardNo As String
    Dim bln消费卡 As Boolean, blnInRange As Boolean   '范围内的卡
    Dim lngIndex As Long, byt变动类型 As Byte, lng结帐ID As Long
    
    strCard = UCase(txt卡号.Text): strICCard = IIf(mblnICCard, strCard, "")
    If Not ((strCard <> "" Or strICCard <> "") And (fraCard.Visible = True Or mbln基本 = False)) Then Exit Sub
    '问题号:56599
    mbln发卡或绑定卡 = True
    
    lng结帐ID = 0: blnInRange = True
    If mCurSendCard.blnOneCard And mCurSendCard.bln严格控制 Then blnInRange = mCurSendCard.lng领用ID > 0
    
    If blnInRange And tabCardMode.SelectedItem.Key = "CardFee" Then
        blnInRange = True
        byt操作类型 = 0: byt变动类型 = 1
    Else
        blnInRange = False
        byt变动类型 = 11: byt操作类型 = 0
    End If
    str变动原因 = "病人信息登记发卡"
    strPassWord = zlCommFun.zlStringEncode(Trim(txtPass.Text))
    If blnInRange = False Then
          'Zl_医疗卡变动_Insert
           strSQL = "Zl_医疗卡变动_Insert("
          '      变动类型_In   Number,
          '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
          strSQL = strSQL & "" & byt变动类型 & ","
          '      病人id_In     住院费用记录.病人id%Type,
          strSQL = strSQL & "" & lng病人ID & ","
          '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
          strSQL = strSQL & "" & mCurSendCard.lng卡类别ID & ","
          '      原卡号_In     病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & str原卡号 & "',"
          '      医疗卡号_In   病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      变动原因_In   病人医疗卡变动.变动原因%Type,
          '      --变动原因_In:如果密码调整，变动原因为密码.加密的
          strSQL = strSQL & "'" & str变动原因 & "',"
          '      密码_In       病人信息.卡验证码%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      操作员姓名_In 住院费用记录.操作员姓名%Type,
          strSQL = strSQL & "'" & UserInfo.姓名 & "',"
          '      变动时间_In   住院费用记录.登记时间%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
          strSQL = strSQL & "NULL)"
    Else
        '103980:李南春,2017/1/19,保存发卡病人年龄
        str年龄 = Trim(txt年龄.Text)
        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
        
        strNO = zlDatabase.GetNextNo(16)  '医疗卡
        If chk记帐.Value = 0 Then
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
        End If
        mCurCardPay.strNO = strNO
        mCurCardPay.lng结帐ID = lng结帐ID
        strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng卡类别ID, byt操作类型, strNO, lng病人ID, 0, UserInfo.部门ID, UserInfo.部门ID, 0, _
         NeedName(cbo费别.Text), "", Trim(txtPatient.Text), NeedName(cbo性别.Text), str年龄, _
        strCard, strPassWord, str变动原因, IIf(mCurSendCard.bln变价 = False, mCurSendCard.dbl应收金额, Val(txt卡额.Text)), Val(txt卡额.Text), IIf(cbo结算方式.Enabled, mCurCardPay.str结算方式, ""), _
        dtCurdate, mCurSendCard.lng领用ID, mCurSendCard.rs卡费, strICCard, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, lng结帐ID)
    End If
    
    zlAddArray cllPro, strSQL
 End Sub
 Private Sub AddDepositSQL(ByVal cllPro As Collection, ByVal dtDate As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加预交款的SQL
    '编制:刘兴洪
    '日期:2011-07-26 18:26:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, strSQL As String, i As Integer, lng预交ID As Long
    Dim dblMoney As Double
    
    If Not (IsNumeric(txt预交额.Text) And fraDeposit.Visible) Then Exit Sub
     
    '病人预交款记录
    strNO = zlDatabase.GetNextNo(11)
    lng预交ID = zlDatabase.GetNextId("病人预交记录")
    mCurPrepay.strNO = strNO
    mCurPrepay.lngID = lng预交ID
    dblMoney = StrToNum(txt预交额.Text)
    'Zl_病人预交记录_Insert
    strSQL = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "" & IIf(mblnPrepayPrint, "'" & txtFact.Text & "'", "Null") & ","
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & "" & Val(txt病人ID.Text) & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSQL = strSQL & "NULL,"
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "NULL,"
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & mCurPrepay.str结算方式 & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "'" & txt结算号码.Text & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strSQL = strSQL & "'" & Trim(txt缴款单位.Text) & "',"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strSQL = strSQL & "'" & Trim(txt开户行.Text) & "',"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strSQL = strSQL & "'" & Trim(txt帐号.Text) & "',"
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'入院预交',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(mlng预交领用ID = 0, "NULL", mlng预交领用ID) & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & "" & Val(Mid(tbDeposit.SelectedItem.Key, 2)) & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.lng医疗卡类别ID = 0 Or mCurPrepay.bln消费卡, "NULL", mCurPrepay.lng医疗卡类别ID) & ","
   '  结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSQL = strSQL & "" & IIf(mCurPrepay.lng医疗卡类别ID = 0 Or Not mCurPrepay.bln消费卡, "NULL", mCurPrepay.lng医疗卡类别ID) & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.str刷卡卡号 = "", "NULL", "'" & mCurPrepay.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  合作单位_In   病人预交记录.合作单位%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  收款时间_In   病人预交记录.收款时间%Type := Null
    '108001:李南春，2017/5/8，格式化预交时间为24小时制
    strSQL = strSQL & "to_date('" & Format(dtDate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '   操作类型_In Integer:=0 :0-正常缴预交;1-存为划价单
    strSQL = strSQL & "0 )"
    zlAddArray cllPro, strSQL
End Sub
Private Function SaveNewCard(strMCAccount As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人病人保存
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 16:57:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPati As String, strSQLCard As String, strCard As String, strICCard As String
    Dim strNO As String, curDate As Date, strSQL As String
    Dim str出生日期 As String, str年龄 As String
    Dim strDepositNO As String, strDeposit As String
    Dim lng预交ID As Long, blnInRange As Boolean
    Dim blnTrans As Boolean, strOut As String, strErr As String
    Dim cllPro As Collection, cllUpdate As Collection, cllThreeInsert As Collection
    Dim i As Long
    Dim arrTmp As Variant
    
    '身份登记
    
    Set cllPro = New Collection
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    
    strCard = UCase(txt卡号.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    curDate = zlDatabase.Currentdate
    '问题号:51071
    If mlngPatientID <> 0 Then
        strPati = "zl_病人信息_UPDATE(" & txt病人ID.Text & "," & _
            IIf(Trim(txt门诊号.Text) <> "", Trim(txt门诊号.Text), "NULL") & "," & _
            IIf(Trim(txt住院号.Text) <> "", Trim(txt住院号.Text), "NULL") & "," & _
            "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo医疗付款.Text) & "','" & txtPatient.Text & "'," & _
            "'" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
            str出生日期 & "," & _
            "'" & txt出生地点.Text & "','" & txt身份证号.Text & "','" & NeedName(cbo身份.Text) & "'," & _
            "'" & NeedName(cbo职业.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo国籍.Text) & "'," & _
            "'" & NeedName(cbo学历.Text) & "','" & NeedName(cbo婚姻状况.Text) & "','" & txt家庭地址.Text & "'," & _
            "'" & txt家庭电话.Text & "','" & txt家庭地址邮编.Text & "','" & txt联系人姓名.Text & "'," & _
            "'" & NeedName(cbo联系人关系.Text) & "','" & txt联系人地址.Text & "','" & txt联系人电话.Text & "'," & _
            Val(txt工作单位.Tag) & ",'" & txt工作单位.Text & "','" & txt单位电话.Text & "','" & txt单位邮编.Text & "'," & _
            "'" & txt单位开户行.Text & "','" & txt单位帐号.Text & "','" & txt联系人姓名.Tag & "'," & Val(txt联系人电话.Tag) & "," & _
            IIf(mlng险类 = 0, "NULL", mlng险类) & "," & IIf(mbytInState = 0, 0, IIf(mbytView = 1 Or mbytView = 2, 1, 0)) & "," & _
            "'" & strMCAccount & "','" & NeedName(txt区域.Text) & "'," & Val(txt联系人地址.Tag) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
            "'" & Trim(txt其他证件.Text) & "','" & NeedName(cbo病人类型.Text) & "'," & _
            IIf(Trim(txt备注.Text) = "", "Null", "'" & Trim(txt备注.Text) & "'") & ",'" & NeedName(txt籍贯.Text) & "','" & txt户口地址.Text & "','" & txt户口地址邮编.Text & "'," & _
            "'" & txt联系人身份证.Text & "',0,'" & Trim(txtMobile.Text) & "')"
        zlAddArray cllPro, strPati
    Else
        strPati = "zl_病人信息_INSERT(" & txt病人ID.Text & "," & _
            IIf(Trim(txt门诊号.Text) <> "", Trim(txt门诊号.Text), "NULL") & "," & _
            "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo医疗付款.Text) & "','" & Trim(txtPatient.Text) & "'," & _
            "'" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
            str出生日期 & "," & _
            "'" & txt出生地点.Text & "','" & txt身份证号.Text & "','" & NeedName(cbo身份.Text) & "'," & _
            "'" & NeedName(cbo职业.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo国籍.Text) & "'," & _
            "'" & NeedName(cbo学历.Text) & "','" & NeedName(cbo婚姻状况.Text) & "','" & txt家庭地址.Text & "'," & _
            "'" & txt家庭电话.Text & "','" & txt家庭地址邮编.Text & "','" & txt联系人姓名.Text & "'," & _
            "'" & NeedName(cbo联系人关系.Text) & "','" & txt联系人地址.Text & "','" & txt联系人电话.Text & "'," & _
            Val(txt工作单位.Tag) & ",'" & txt工作单位.Text & "','" & txt单位电话.Text & "','" & txt单位邮编.Text & "'," & _
            "'" & txt单位开户行.Text & "','" & txt单位帐号.Text & "',null,null," & _
            "NULL,To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & NeedName(txt区域.Text) & "',null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & strMCAccount & "'," & _
            "'" & Trim(txt其他证件.Text) & "','" & NeedName(txt籍贯.Text) & "','" & txt户口地址.Text & "','" & txt户口地址邮编.Text & "'," & _
            "'" & txt联系人身份证.Text & "','" & NeedName(cbo病人类型.Text) & "','" & Trim(txtMobile.Text) & "')"
        zlAddArray cllPro, strPati
    End If
    '从表信息保存
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",联系人附加信息,身份证号状态,外籍身份证号,", Split(arrTmp(i), ":")(0)) > 0 Then
                strPati = "Zl_病人信息从表_Update(" & txt病人ID.Text & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
                zlAddArray cllPro, strPati
            End If
        Next
    End If
    '问题号:53408
    If Trim(txt支付密码.Text) <> "" And Trim(txt身份证号.Text) <> "" Then
        If zl绑定身份证(cllPro) = False Then Exit Function
    End If
    
    '结构化地址 89980
    If gbln启用结构化地址 Then
        Call CreateStructAddressSQL(CLng(txt病人ID.Text), mlng主页ID, cllPro, PatiAddress)
    End If
    
    '医疗卡处理
    '问题号:51072
    If Len(Trim(txtPass.Text)) <= 0 And Len(Trim(txt卡号.Text)) > 0 Then '没有输入密码
        If zl_Get设置默认发卡密码 = False Then Exit Function
    End If
    
    Call AddCardDataSQL(Val(txt病人ID.Text), curDate, cllPro) '加入医疗卡
    '问题号:57326
    If mbln发卡或绑定卡 Then
        If Check发卡性质(Val(txt病人ID.Text), mCurSendCard.lng卡类别ID) = False Then
            txt卡号.Text = "": txtPass.Text = "": txtAudi.Text = "": txt卡额.Text = ""
            Exit Function
        End If
        '检查结算方式信息是否合法
        If (cbo结算方式.ItemData(cbo结算方式.ListIndex) = 8 Or cbo结算方式.ItemData(cbo结算方式.ListIndex) = 7) And mCurCardPay.lng医疗卡类别ID = 0 Then
            MsgBox "当前发卡结算方式存在异常，无法使用该结算方式，请检查是否启用相应设备或与管理员联系!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    Call AddDepositSQL(cllPro, curDate)  '加入预交款
    '检查预交结算方式信息是否合法
    If IsNumeric(txt预交额.Text) And fraDeposit.Visible Then
        If (cbo预交结算.ItemData(cbo预交结算.ListIndex) = 8 Or cbo预交结算.ItemData(cbo预交结算.ListIndex) = 7) And mCurPrepay.lng医疗卡类别ID = 0 Then
            MsgBox "当前预交结算方式存在异常，无法使用该结算方式，请检查是否启用相应设备或与管理员联系!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    '问题号:56599
    If Val(Trim(txt病人ID.Text)) > 0 Then Call Add健康卡相关信息(Val(Trim(txt病人ID.Text)), cllPro)
    
    On Error GoTo errH
    
    Set cllUpdate = New Collection
    Set cllThreeInsert = New Collection
    
    Err = 0: On Error GoTo ErrHand:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    '支付交易
    If Not zlInterfacePrayMoney(cllUpdate, cllThreeInsert) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    '修正三方交易
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    '74421,病人照片添加
    Call SavePatPicture(Val(txt病人ID.Text))
    '101160EMPI
    If Not EMPI_AddORUpdatePati(CLng(txt病人ID.Text), mlng主页ID, strErr) Then
        gcnOracle.RollbackTrans
        MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    '问题号:56599
    '写卡
    If mbln发卡或绑定卡 And mCurSendCard.bln是否写卡 Then WriteCard (Val(txt病人ID.Text))
    
    Err = 0: On Error GoTo OthersCommit:
    zlExecuteProcedureArrAy cllThreeInsert, Me.Caption
    Call zlExcuteUploadSwap(txt病人ID.Text, strOut, mobjICCard) '调用宁波一卡通上传功能
    
    '73937:刘鹏飞,2013-07-03
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.PatiInfoSaveAfter(Val(txt病人ID.Text))
        Err.Clear
    End If
    SaveNewCard = True
    Exit Function
OthersCommit:
      If ErrCenter = 1 Then
            gcnOracle.RollbackTrans
            Resume
      End If
      Call SaveErrLog
      gcnOracle.CommitTrans
      SaveNewCard = True
      Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveModiCard(strMCAccount As String) As Boolean
'功能：对被修改的病人信息卡片进行保存
    Dim strSQL As String
    Dim str出生日期 As String, str年龄 As String
    Dim blnTrans As Boolean
    Dim cllPro As New Collection  '问题号:56599
    Dim arrTmp As Variant
    Dim i As Long
    Dim strErr As String
    
    On Error GoTo errH
    
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    
    '问题号:51071
    '问题27351 by lesfeng 2010-01-12
    strSQL = "zl_病人信息_UPDATE(" & txt病人ID.Text & "," & _
        IIf(Trim(txt门诊号.Text) <> "", Trim(txt门诊号.Text), "NULL") & "," & _
        IIf(Trim(txt住院号.Text) <> "", Trim(txt住院号.Text), "NULL") & "," & _
        "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo医疗付款.Text) & "','" & txtPatient.Text & "'," & _
        "'" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
        str出生日期 & "," & _
        "'" & txt出生地点.Text & "','" & txt身份证号.Text & "','" & NeedName(cbo身份.Text) & "'," & _
        "'" & NeedName(cbo职业.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo国籍.Text) & "'," & _
        "'" & NeedName(cbo学历.Text) & "','" & NeedName(cbo婚姻状况.Text) & "','" & txt家庭地址.Text & "'," & _
        "'" & txt家庭电话.Text & "','" & txt家庭地址邮编.Text & "','" & txt联系人姓名.Text & "'," & _
        "'" & NeedName(cbo联系人关系.Text) & "','" & txt联系人地址.Text & "','" & txt联系人电话.Text & "'," & _
        Val(txt工作单位.Tag) & ",'" & txt工作单位.Text & "','" & txt单位电话.Text & "','" & txt单位邮编.Text & "'," & _
        "'" & txt单位开户行.Text & "','" & txt单位帐号.Text & "','" & txt联系人姓名.Tag & "'," & Val(txt联系人电话.Tag) & "," & _
        IIf(mlng险类 = 0, "NULL", mlng险类) & "," & IIf(mbytView = 1 Or mbytView = 2, 1, 0) & "," & _
        "'" & strMCAccount & "','" & NeedName(txt区域.Text) & "'," & Val(txt联系人地址.Tag) & ",'" & UserInfo.编号 & "','" & _
        UserInfo.姓名 & "','" & Trim(txt其他证件.Text) & "','" & NeedName(cbo病人类型.Text) & "'," & _
        IIf(Trim(txt备注.Text) = "", "Null", "'" & Trim(txt备注.Text) & "'") & ",'" & NeedName(txt籍贯.Text) & "','" & txt户口地址.Text & "','" & txt户口地址邮编.Text & "'," & _
        "'" & Trim(txt联系人身份证.Text) & "',0,'" & Trim(txtMobile.Text) & "')"
    '结构化地址
    If gbln启用结构化地址 Then
        Call CreateStructAddressSQL(CLng(txt病人ID.Text), mlng主页ID, cllPro, PatiAddress, 1)
    End If
    '病案主页从表信息保存
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",联系人附加信息,身份证号状态,外籍身份证号,", Split(arrTmp(i), ":")(0)) > 0 Then
                zlAddArray cllPro, "Zl_病人信息从表_Update(" & txt病人ID.Text & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
        Next
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    '74421
    Call SavePatPicture(Val(txt病人ID.Text))
    '问题号:56599
    If mlng病人ID > 0 Then Call Add健康卡相关信息(mlng病人ID, cllPro)
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
    '101160 EMPI
    If Not EMPI_AddORUpdatePati(CLng(txt病人ID.Text), mlng主页ID, strErr) Then
        gcnOracle.RollbackTrans
        MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    '新网118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(IIf(mlng主页ID <> 0, 2, 1), CLng(txt病人ID.Text), mlng主页ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "当前启用了影像信息系统接口，但于由RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    '问题号:56599
    '写卡
    If mbln发卡或绑定卡 And mCurSendCard.bln是否发卡 Then WriteCard (Val(txt病人ID.Text))
    
    blnTrans = False
    
    '73937:刘鹏飞,2013-07-03
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.PatiInfoSaveAfter(Val(txt病人ID.Text))
        Err.Clear
    End If
    
    SaveModiCard = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatient_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatiMCNO_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtPatiMCNO(Index))
End Sub

Private Sub txtPatiMCNO_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
Public Function CheckExistsMCNO(ByVal strMCNO As String) As Boolean
'功能:检查医保号是否已存在
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 1 From 病人信息 Where 医保号 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strMCNO)
    If rsTmp.RecordCount > 0 Then
        MsgBox "请检查,输入的医保号已存在!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatiMCNO_Validate(Index As Integer, Cancel As Boolean)
    txtPatiMCNO(Index).Text = UCase(Trim(txtPatiMCNO(Index).Text))
    '问题28474 by lesfeng 2010-03-16 取消不能退出医保号及验证医保号输入
    If Index = 1 Then
        If txtPatiMCNO(1).Text <> txtPatiMCNO(0).Text Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
'            Cancel = True
            Exit Sub
        End If
    End If
    
    If mlngOutModeMC = 920 And txtPatiMCNO(0).Text <> txtPatiMCNO(0).Tag And txtPatiMCNO(0).Text <> "" Then
        If CheckExistsMCNO(txtPatiMCNO(0).Text) Then
'            Cancel = True
        End If
    End If
End Sub

Private Sub txt身份证号_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
    If Trim(txt身份证号.Text) = "" Then
        cboIDNumber.Enabled = True
        If cboIDNumber.Enabled And cboIDNumber.Visible Then cboIDNumber.SetFocus
    Else
        cboIDNumber.Enabled = False
        cboIDNumber.ListIndex = -1
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt验证密码_GotFocus()
    Call zlControl.TxtSelAll(txt验证密码)
    Call OpenPassKeyboard(txt验证密码, False)
End Sub

Private Sub txt验证密码_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int密码规则 = 1)
End Sub

Private Sub txt验证密码_LostFocus()
    Call ClosePassKeyboard(txt验证密码)
End Sub

Private Sub txt预交额_GotFocus()
    If IsNumeric(txt预交额.Text) Then
        txt预交额.Text = StrToNum(txt预交额.Text)
    Else
        txt预交额.Text = ""
    End If
    txt预交额.SelStart = 0: txt预交额.SelLength = Len(txt预交额.Text)
End Sub
Private Sub CheckInputPassWord(KeyAscii As Integer, Optional ByVal blnOnlyNum As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查密码输入
    '编制:刘兴洪
    '日期:2011-07-07 00:40:53
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
Private Sub txt预交额_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then
        If InStr(txt预交额.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        '65965:刘鹏飞,2013-09-24,处理预交显示千位位格式
        If (txt预交额.Text <> "" And txt预交额.SelLength <> Len(Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;")) >= txt预交额.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txt预交额.SelLength > 0 And txt预交额.SelLength <= txt预交额.MaxLength Then
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf IsNumeric(txt预交额.Text) Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '不收取预交款,直接跳过
        txt预交额.Text = ""
        If fraCard.Visible Then
            txt卡号.SetFocus
        Else
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txt预交额_LostFocus()
    '65965:刘鹏飞,2013-09-24,处理预交显示千位位格式
    If IsNumeric(txt预交额.Text) Then
        txt预交额.Text = Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;")
    Else
        txt预交额.Text = ""
    End If
    If txt预交额.MaxLength > 12 Then txt预交额.MaxLength = 12
End Sub

Private Sub txt帐号_GotFocus()
    If IsNumeric(txt预交额.Text) And txt帐号.Text = "" Then
        txt帐号.Text = txt单位帐号.Text
    End If
    zlControl.TxtSelAll txt帐号
End Sub

Private Sub txt帐号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt缴款单位, KeyAscii
End Sub

Private Sub txt帐号_LostFocus()
    Call zlCommFun.OpenIme
End Sub
Private Sub txt支付密码_GotFocus()
    Call zlControl.TxtSelAll(txt支付密码)
    Call OpenPassKeyboard(txt支付密码, False)
End Sub

Private Sub txt支付密码_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int密码规则 = 1)
End Sub

Private Sub txt支付密码_LostFocus()
    Call ClosePassKeyboard(txt支付密码)
End Sub

Private Sub txt住院号_GotFocus()
    SelAll txt住院号
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt住院号_Validate(Cancel As Boolean)
    If Val(txt住院号.Text) = 0 And Val(txt住院号.Tag) <> 0 Then txt住院号.Text = txt住院号.Tag
End Sub
 
Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化预交类型
    '编制:刘兴洪
    '日期:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With tbDeposit
        mblnNotClick = True
        .Tabs.Clear
        If InStr(1, mstrPrivs, ";门诊预交;") > 0 Then
            .Tabs.Add(, "K1", "门诊预交(&M)").Selected = IIf(mbytPrepayType = 1, True, False)
        End If
        If InStr(1, mstrPrivs, ";住院预交;") > 0 Then
            .Tabs.Add(, "K2", "住院预交(&Z)").Selected = IIf(mbytPrepayType = 2, True, False)
        End If
         If .Tabs.Count > 0 And .SelectedItem Is Nothing Then
            .Tabs(0).Selected = True
         End If
         mblnNotClick = False
         Call tbDeposit_Click
         If .Tabs.Count = 0 Then
            fraDeposit.Visible = False
            Me.Height = Me.Height - fraDeposit.Height
            mPageHeight.基本 = Me.Height
            If InStr(mstrPrivs, ";预交退款;") = 0 Then cmdOperation(OPT.C0预交款).Visible = False
         Else
            Call GetFact(True)
         End If
     End With
End Sub



Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
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

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    '结算方式:费用查询和医疗卡调用时，一般只支付预交款,不存在代收的情况
    strSQL = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合 ='预交款'  And B.名称=A.结算方式  " & _
        "           And Nvl(B.性质,1) In(1,2,7,8)" & _
        " Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolPrepayPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cbo预交结算
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!性质) & ",") = 0 Then
                .AddItem Nvl(rsTemp!名称)
                mcolPrepayPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                'If mstr缺省结算方式 = Nvl(rsTemp!名称) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '结算方式中设置且设备配置启用了的结算方式才有效
            rsTemp.Filter = "名称 ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                                If InStr(1, varData(i), "|") <> 0 Then
                                        varTemp = Split(varData(i), "|")
                                        mcolPrepayPayMode.Add varTemp, "K" & j
                                        .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                                        'If mstr缺省结算方式 = varTemp(1) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                                        j = j + 1
                                End If
                        End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cbo预交结算.ListCount = 0 Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    strSQL = _
    "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
    " From 结算方式应用 A,结算方式 B" & _
    " Where A.应用场合 ='就诊卡'  And B.名称=A.结算方式  " & _
    "           And Nvl(B.性质,1) In(1,2,7,8)" & _
    " Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolCardPayMode = New Collection
    With cbo结算方式
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!性质) & ",") = 0 Then
                .AddItem Nvl(rsTemp!名称)
                mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '结算方式中设置且设备配置启用了的结算方式才有效
            rsTemp.Filter = "名称 ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                                If InStr(1, varData(i), "|") <> 0 Then
                                        varTemp = Split(varData(i), "|")
                                        mcolCardPayMode.Add varTemp, "K" & j
                                        .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                                        j = j + 1
                                End If
                        End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Local结算方式(ByVal lng卡类别ID As Long, Optional bln预交 As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定位结算方式
    '编制:刘兴洪
    '日期:2011-07-26 15:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPayMoney As Collection, cboPay As ComboBox
    Dim i As Long
    If mblnNotClick Then Exit Sub
    
    If bln预交 Then
       Set cllPayMoney = mcolPrepayPayMode
        Set cboPay = cbo预交结算
    Else
       Set cllPayMoney = mcolCardPayMode
        Set cboPay = cbo结算方式
    End If
    If cllPayMoney Is Nothing Then Exit Sub
    With cboPay
        If .ListIndex >= 0 Then
            If bln预交 Then
                If .ItemData(.ListIndex) >= 0 Then Exit Sub
            End If
        End If
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            ''短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
            If Val(cllPayMoney(i + 1)(3)) = lng卡类别ID Then
                .ListIndex = i: Exit For
            End If
        Next
        mblnNotClick = False
    End With
End Sub
Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        '58322
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        .ActiveConnection = Nothing
        If mCurPrepay.lng医疗卡类别ID <> 0 Then
            .AddNew
            !收费类别 = "预交"
            !金额 = StrToNum(txt预交额.Text)
            .Update
        End If
        If mCurCardPay.lng医疗卡类别ID <> 0 And cbo结算方式.Enabled And cbo结算方式.Visible Then
            .AddNew
            !收费类别 = mCurSendCard.rs卡费!收费类别
            !金额 = StrToNum(txt卡额.Text)
            .Update
        End If
    End With
    zlGetClassMoney = True
End Function

Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str年龄 As String
    Dim dblMoney As Double, bln三方结算 As Boolean
    Dim dblThreeMoney As Double, tyCurThreePay As Ty_PayMoney
    Dim blnTemp As Boolean
    
    On Error GoTo errHandle
    '58322
    dblMoney = 0: dblThreeMoney = 0
    If cbo预交结算.Visible Then
        If cbo预交结算.ListIndex >= 0 Then
            bln三方结算 = cbo预交结算.ItemData(cbo预交结算.ListIndex) = -1
            If bln三方结算 Then dblThreeMoney = dblThreeMoney + StrToNum(txt预交额.Text)
        End If
        dblMoney = dblMoney + StrToNum(txt预交额.Text)
    End If
    
    If cbo结算方式.Visible And cbo结算方式.Enabled Then
        If cbo结算方式.ListIndex >= 0 Then
            blnTemp = cbo结算方式.ItemData(cbo结算方式.ListIndex) = -1
            If blnTemp Then dblThreeMoney = dblThreeMoney + StrToNum(txt卡额.Text)
            If blnTemp Then bln三方结算 = bln三方结算 Or blnTemp
        End If
        dblMoney = dblMoney + StrToNum(txt卡额.Text)
    End If
    If Not bln三方结算 Then CheckBrushCard = True: Exit Function
    '问题：126188，2018-06-19,入院登记失败
    If dblThreeMoney = 0 Then CheckBrushCard = True: Exit Function
    If mCurPrepay.lng医疗卡类别ID <> 0 Then
       tyCurThreePay = mCurPrepay
    Else
       tyCurThreePay = mCurCardPay
    End If
    
    
    If (mCurCardPay.lng医疗卡类别ID <> mCurCardPay.lng医疗卡类别ID Or _
        mCurPrepay.bln消费卡 <> mCurCardPay.bln消费卡) _
        And mCurCardPay.lng医疗卡类别ID <> 0 And mCurPrepay.lng医疗卡类别ID <> 0 Then
        MsgBox "不能同时使用两种不同类别的支付方式,不能继续?", vbOKOnly + vbInformation, gstrSysName
        If cbo预交结算.Enabled And cbo预交结算.Visible Then cbo预交结算.SetFocus: Exit Function
        If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus
        Exit Function
    End If
    Call zlGetClassMoney(rsMoney)
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String) As Boolean
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
   '58322
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, tyCurThreePay.lng医疗卡类别ID, tyCurThreePay.bln消费卡, _
    txtPatient.Text, NeedName(cbo性别.Text), str年龄, dblThreeMoney, tyCurThreePay.str刷卡卡号, tyCurThreePay.str刷卡密码, False, True, False) = False Then Exit Function
    
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, tyCurThreePay.lng医疗卡类别ID, _
        tyCurThreePay.bln消费卡, tyCurThreePay.str刷卡卡号, dblThreeMoney, "", "") = False Then Exit Function
    mCurCardPay.str刷卡卡号 = tyCurThreePay.str刷卡卡号
    mCurCardPay.str刷卡密码 = tyCurThreePay.str刷卡密码
    mCurPrepay.str刷卡卡号 = tyCurThreePay.str刷卡卡号
    mCurPrepay.str刷卡密码 = tyCurThreePay.str刷卡密码
    
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlInterfacePrayMoney(ByRef cllPro As Collection, ByRef cllThreeSwap As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim dblMoney As Double
    If mCurCardPay.lng医疗卡类别ID = 0 And mCurPrepay.lng医疗卡类别ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cbo预交结算.ItemData(cbo预交结算.ListIndex) <> -1 _
        And cbo结算方式.ItemData(cbo结算方式.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dblMoney = 0
    If mCurCardPay.lng医疗卡类别ID <> 0 And cbo结算方式.Enabled And cbo结算方式.Visible Then
        dblMoney = Val(txt卡额.Text)
    End If
    If mCurPrepay.lng医疗卡类别ID <> 0 And cbo预交结算.Visible Then
        dblMoney = dblMoney + Val(StrToNum(txt预交额.Text))
    End If
    '问题：126188，2018-06-19,入院登记失败
    If dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, mCurCardPay.lng结帐ID, mCurPrepay.strNO, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.lng结帐ID <> 0 And cbo结算方式.Visible Then
     
        If Not mCurCardPay.bln消费卡 Then
            Call zlAddUpdateSwapSQL(False, mCurCardPay.lng结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mCurCardPay.lng结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    If mCurPrepay.lng医疗卡类别ID <> 0 And cbo预交结算.Visible And Val(StrToNum(txt预交额.Text)) <> 0 Then
        Call zlAddUpdateSwapSQL(True, mCurPrepay.lngID, mCurPrepay.lng医疗卡类别ID, mCurPrepay.bln消费卡, mCurPrepay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
        Call zlAddThreeSwapSQLToCollection(True, mCurPrepay.lngID, mCurPrepay.lng医疗卡类别ID, mCurPrepay.bln消费卡, mCurPrepay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtAudi_Validate(Cancel As Boolean)
    Select Case mCurSendCard.int密码长度限制
        Case 0
        Case 1
            If Len(txtAudi.Text) <> mCurSendCard.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "确认密码必须输入" & mCurSendCard.int密码长度 & "位", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtAudi.Text) < Abs(mCurSendCard.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "确密码必须输入" & Abs(mCurSendCard.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Exit Sub
             End If
        End Select
End Sub

Private Sub txtPass_Validate(Cancel As Boolean)
   Select Case mCurSendCard.int密码长度限制
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & mCurSendCard.int密码长度 & "位", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(mCurSendCard.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        End Select
End Sub

Private Function zl_Get设置默认发卡密码() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置默认发卡密码
    '返回:是否继续发卡操作
    '编制:王吉
    '日期:2012-07-06 15:53:14
    '问题号:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardType As clsCard
    Dim msgResult As VbMsgBoxResult
    Dim arr() As String
    arr = zl_Get医疗卡类型(mCurSendCard.lng卡类别ID)
    If Val(arr(2)) = 0 Then '无限制
        Select Case Val(arr(1))
            Case 0 '无限制
                zl_Get设置默认发卡密码 = True
                Exit Function
            Case 1 '未输入提醒
               msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续？", vbQuestion + vbYesNo, gstrSysName)
               zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
               Exit Function
            Case 2 '为输入禁止
                 MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                zl_Get设置默认发卡密码 = False
                Exit Function
        End Select
    ElseIf Val(arr(2)) = 1 Then '缺省身份证后N位
        If Len(Trim(txt身份证号.Text)) > 0 Or Len(Trim(txt联系人身份证.Text)) > 0 Then '输入了身份证或联系人身份证号
            If Len(Trim(txt身份证号.Text)) > 0 Then '有身份证优先用身份证
                   txtPass.Text = Right(Trim(txt身份证号.Text), Val(arr(0)))
            Else '否则就用代办人身份证作为密码
                   txtPass.Text = Right(Trim(txt联系人身份证.Text), Val(arr(0)))
            End If
        Else '身份证与联系人身份证都没输入
            Select Case Val(arr(1))
                Case 0 '无限制
                    zl_Get设置默认发卡密码 = True
                    Exit Function
                Case 1 '未输入提醒
                    msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续！", vbQuestion + vbYesNo, gstrSysName)
                    zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
                    Exit Function
                Case 2 '为输入禁止
                    MsgBox "未输入卡密码,不能进行发卡？", vbExclamation, gstrSysName
                    zl_Get设置默认发卡密码 = False
                    Exit Function
            End Select
        End If
    End If
    zl_Get设置默认发卡密码 = True
End Function


Public Sub Show绑定控件(blnShow As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否显示绑定密码
    '入参:blnShow 是否显示绑定密码
    '编制:王吉
    '日期:2012-09-04 15:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    lbl支付密码.Enabled = blnShow: txt支付密码.Enabled = blnShow
    lbl验证密码.Enabled = blnShow: txt验证密码.Enabled = blnShow
    If blnShow = False Then
        txt支付密码.Text = "": txt验证密码.Text = ""
    End If
    
End Sub
Private Function zl绑定身份证(colPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定二代身份证
    '入参:blnShow 是否显示绑定密码
    '编制:王吉
    '日期:2012-09-04 15:53:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txt支付密码.Text) <> Trim(txt验证密码.Text) Then
        MsgBox "两次输入的密码不一致,请重新输入", vbOKOnly + vbInformation, gstrSysName
        txt支付密码.Text = "": txt验证密码.Text = ""
        If txt支付密码.Visible = True Then txt支付密码.SetFocus
        Exit Function
    End If
    If Trim(txt支付密码.Text) <> "" Then
       If 是否已经签约(Trim(txt身份证号.Text)) Then
             MsgBox "身份证号码为:" & txt身份证号.Text & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
             txt支付密码.Text = "": txt验证密码.Text = ""
             If txt支付密码.Visible = True Then txt支付密码.SetFocus
             Exit Function
       End If
    End If
    AddSQL绑定卡 Trim(txt病人ID.Text), Get医疗卡类别ID("二代身份证"), Trim(txt身份证号.Text), zlCommFun.zlStringEncode(Trim(txt支付密码.Text)), zlDatabase.Currentdate, False, colPro
    
    zl绑定身份证 = True
End Function
Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分页控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    
    mlngPlugInHwnd = 0
    
    picTmp.Visible = False
    
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        mlngPlugInHwnd = gobjPlugIn.GetFormHwnd

        Err.Clear: On Error GoTo 0
    End If
    
    Err = 0: On Error GoTo ErrHand:
        
    Set objItem = tbcPage.InsertItem(1, "基本", PicBaseInfo.hWnd, 0)
    objItem.Tag = mPageHeight.基本
    
    Set objItem = tbcPage.InsertItem(2, "健康档案", PicHealth.hWnd, 0)
    objItem.Tag = mPageHeight.健康档案
    If mlngPlugInHwnd <> 0 Then
        picTmp.Visible = True
        Set objItem = tbcPage.InsertItem(3, "附加信息", picTmp.hWnd, 0)
        objItem.Tag = mPageHeight.附加信息
    End If
    
    PicBaseInfo.Enabled = False
    PicHealth.Enabled = False
    With tbcPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .Item(0).Selected = True
    End With
    PicBaseInfo.Enabled = True
    PicHealth.Enabled = True
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub SetColumHeader(ByRef vsGrid As VSFlexGrid, ByVal strHead As String, Optional ByVal lngNO As Long = 0)
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1
        If lngNO = 0 Then
            .FixedCols = 0
            .Cols = .FixedCols + UBound(arrHead) + 1
            .Rows = .FixedRows + 1
        Else
            .FixedCols = 1
            .Cols = .FixedCols + UBound(arrHead)
            .Rows = .FixedRows + 1
        End If

        For i = 0 To UBound(arrHead)
            If .FixedCols > 0 Then
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            Else
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            End If
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               '为了支持zl9PrintMode
                If .FixedCols > 0 Then
                    .ColHidden(i) = Val(Split(arrHead(i), ",")(3)) = 0
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(2))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(1))
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(1))
                Else
                    .ColHidden(.FixedCols + i) = Val(Split(arrHead(i), ",")(3)) = 0
                    .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                    .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
'                    .ColData
                    '为了支持zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  '为了支持zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
                End If
            End If
        Next
        
        '固定行文字居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub

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
Private Sub InitCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化ComBox控件
    '编制:56599
    '日期:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '66743:刘尔旋,2013-11-25,血型与RH默认值的问题
    'ComboBox cboBloodType, C_血型
    Call ReadDict("血型", cboBloodType)
    ComboBox cboBH, C_BH
    If cboBH.ListCount <> 0 Then cboBH.ListIndex = -1
End Sub
Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo ErrHand
    
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 社会关系 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "社会关系")
    With rsTemp
        Do While Not rsTemp.EOF
            strTmp = strTmp & "|" & Nvl(rsTemp!名称)
        rsTemp.MoveNext
        Loop
    End With
    If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
    
    With vsLinkMan
        '初始化列表属性
        SetColumHeader vsLinkMan, C_LinkManColumHeader
        .Editable = IIf(mbytInState = E查阅, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        If strTmp <> "" Then .ColComboList(.ColIndex("联系人关系")) = strTmp
    End With
    
    With vsOtherInfo
        '设置列头
        SetColumHeader vsOtherInfo, C_OtherInfoColumHeader
        .Editable = IIf(mbytInState = E查阅, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InitvsDrug()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsDrug
        '初始化列表属性
        SetColumHeader vsDrug, C_ColumHeader
        .Editable = IIf(mbytInState = E查阅, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        .ColComboList(0) = "..."
    End With
End Sub

Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
        '初始化列表属性
        SetColumHeader vsInoculate, C_InoculateHeader
         vsInoculate.Editable = IIf(mbytInState = E查阅, flexEDNone, flexEDKbdMouse)
        '设置选择按钮
        .ColDataType(0) = flexDTDate
        .ColEditMask(0) = "####-##-##"
        .ColDataType(2) = flexDTDate
        .ColEditMask(2) = "####-##-##"
        .SelectionMode = flexSelectionFree
    End With

End Sub

Private Sub vsDrug_EnterCell()
    If vsDrug.Col = vsDrug.FixedCols Then
        vsDrug.ColComboList(vsDrug.Col) = "..."
    End If
End Sub

Private Sub vsDrug_GotFocus()
    If mblnCheckPatiCard = False Then
        vsDrug.Col = vsDrug.FixedCols
        vsDrug.Row = vsDrug.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub vsDrug_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsDrug_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim strInput As String, strFilter As String
    Dim rsTemp As Recordset
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandl
    
    If Not Col = vsDrug.FixedCols Then Exit Sub

    strInput = Trim(vsDrug.EditText)
    
    If strInput <> "" Then
        If zlCommFun.IsCharAlpha(strInput) Then
            strFilter = " And zlspellcode(A.名称) like [1]"
            strInput = UCase(strInput)
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strFilter = " And A.名称 like [1]"
        Else
            strFilter = " And A.编码 like [1]"
        End If
    End If
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select Distinct A.ID,A.编码," & _
        " A.名称,zlspellcode(A.名称) 简码,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & strFilter

    '获取当前鼠标坐标值
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    strInput = gstrLike & Trim(strInput) & "%"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏药物", False, "过敏药物选择器", "请从下面的药品中选择一项作为病人过敏药物", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True, strInput)

    If Not rsTemp Is Nothing And blnCancel = False Then
        If rsTemp.RecordCount > 0 Then
            vsDrug.EditText = Nvl(rsTemp!名称)
            vsDrug.TextMatrix(Row, Col) = Nvl(rsTemp!名称)
            vsDrug.TextMatrix(Row, 2) = Nvl(rsTemp!ID)
            If vsDrug.Rows - 1 = Row Then vsDrug.Rows = vsDrug.Rows + 1
        End If
    Else
        vsDrug.EditText = vsDrug.TextMatrix(Row, Col)
    End If
    
    Exit Sub
ErrHandl:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsInoculate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col <> 1 And Col <> 3 Then
        If vsInoculate.TextMatrix(Row, Col) = "____-__-__" Then vsInoculate.TextMatrix(Row, Col) = ""
    End If
End Sub

Private Sub vsInoculate_GotFocus()
    If mblnCheckPatiCard = False Then
        vsInoculate.Col = vsInoculate.FixedCols
        vsInoculate.Row = vsInoculate.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub VsInoculate_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:56599
    Dim intRow As Integer
    
    With vsInoculate
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Col > .FixedCols + 1 Then
                .TextMatrix(intRow, .FixedCols + 2) = ""
                .TextMatrix(intRow, .FixedCols + 3) = ""
            Else
                If .Rows = .FixedRows + 1 Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    Call .RemoveItem(.Row)
                    If intRow >= .Rows Then
                        .Row = .Rows - 1
                    Else
                        .Row = intRow
                    End If
                    .Col = .FixedCols
                End If
            End If
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If ((.TextMatrix(.Row, .FixedCols) = "" And .Col = .FixedCols) Or (.Col = .FixedCols + 2 And .TextMatrix(.Row, .FixedCols + 2) = "") Or .Col = .FixedCols + 3) And .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
               Call MoveNextCell(vsInoculate)
            End If
        End If
    End With
End Sub
Private Sub vsDrug_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '问题号:56599
    If Col = 1 Then  '过敏反应列编辑时需判断是否字数超过了100
        With vsDrug
           If LenB(StrConv(.TextMatrix(Row, Col), vbFromUnicode)) > 100 Then
                MsgBox "过敏反应输入字符超出最大字符数100,多出的字符将被自动截除！", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = StrConv(MidB(StrConv(.TextMatrix(Row, Col), vbFromUnicode), 1, 100), vbUnicode)
           End If
        End With
    End If
End Sub

Private Sub vsDrug_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '问题号:56599
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandl:
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL 简码,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL 简码,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL 简码,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select ID,nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称,NULL 简码," & _
        " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试" & _
        " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码," & _
        " A.名称,zlspellcode(A.名称) 简码,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"

    '获取当前鼠标坐标值
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "过敏药物选择器", "请从下面的药品中选择一项作为病人过敏药物", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True)

    If Not rsTemp Is Nothing And blnCancel = False Then
        vsDrug.TextMatrix(Row, Col) = rsTemp!名称
        vsDrug.TextMatrix(Row, 2) = rsTemp!ID
        If vsDrug.Rows - 1 = Row Then vsDrug.Rows = vsDrug.Rows + 1
    End If
    Exit Sub
ErrHandl:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:56599
    Dim intRow As Integer
    With vsDrug
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete And .ColComboList(.Col) = "..." Then
            .ColComboList(.Col) = ""
        End If
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Rows = .FixedRows + 1 Then
                vsDrug.TextMatrix(1, 0) = "": vsDrug.Cell(flexcpData, 1, 0) = "": vsDrug.TextMatrix(1, 1) = ""
            Else
                Call vsDrug.RemoveItem(.Row)
                If intRow >= .Rows Then
                    .Row = .Rows - 1
                Else
                    .Row = intRow
                End If
            End If
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If .TextMatrix(.Row, .FixedCols) = "" And .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                Call MoveNextCell(vsDrug)
            End If
        End If
    End With
End Sub

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
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
    "       Select 编码 as ID,名称,简码 From 医学警示 Where 名称 Not Like '其他%'"
    Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "医学警示", False, "", "", False, False, False, vRect.Left, vRect.Top - 180, 500, True, False, True)
    If Not rsTemp Is Nothing Then
        While rsTemp.EOF = False
          strTemp = strTemp & ";" & rsTemp!名称
          rsTemp.MoveNext
        Wend
    Else
        If cmdMedicalWarning.Enabled And cmdMedicalWarning.Visible Then cmdMedicalWarning.SetFocus: Exit Sub
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
    If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
End Sub
Private Sub SetDrugAllergy(str过敏药物 As String, str过敏反应 As String, Optional lng过敏ID = 0, Optional ByVal byt信息更新模式 As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置过敏药物
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str过敏药物 Then
                    If mlng病人ID <> 0 And byt信息更新模式 = 2 Then Exit Sub
                    .TextMatrix(i, 1) = str过敏反应
                    If lng过敏ID <> 0 Then .TextMatrix(i, 2) = lng过敏ID
                    Exit Sub
                End If
            Next
        End If
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str过敏药物
        .TextMatrix(.Rows - 1, 1) = str过敏反应
        If lng过敏ID <> 0 Then .TextMatrix(.Rows - 1, 2) = lng过敏ID
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetInoculate(str接种日期 As String, str接种名称 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置接种情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    '68192:刘鹏飞,2013-12-02
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If Format(.TextMatrix(i, j - 1), "YYYY-MM-DD") = Format(str接种日期, "YYYY-MM-DD") And .TextMatrix(i, j) = str接种名称 Then Exit Sub
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" Or .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = Format(str接种日期, "YYYY-MM-DD")
                .TextMatrix(.Rows - 1, j + 1) = str接种名称
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
        
    End With
End Sub

Private Sub SetLinkInfo(str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, Optional ByVal byt信息更新模式 As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置联系人相关信息
    '编制:56599
    '日期:2012-12-12 09:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    With vsLinkMan
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str姓名 And .TextMatrix(i, 2) = str身份证号 Then
                    If mlng病人ID <> 0 And byt信息更新模式 = 2 Then Exit Sub
                    .TextMatrix(i, 1) = str关系: .TextMatrix(i, 3) = str电话
                    If i = 1 Then
                        txt联系人身份证.Text = str身份证号
                        txt联系人姓名.Text = str姓名
                        For j = 0 To cbo联系人关系.ListCount - 1
                            If NeedName(cbo联系人关系.List(j)) = str关系 Then cbo联系人关系.ListIndex = j
                        Next
                        txt联系人电话.Text = str电话
                    End If
                    Exit Sub
                End If
            Next
        End If
        
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str姓名
        .TextMatrix(.Rows - 1, 1) = str关系
        .TextMatrix(.Rows - 1, 3) = str电话
        .TextMatrix(.Rows - 1, 2) = str身份证号
        If .Rows - 1 = 1 Then
            txt联系人身份证.Text = str身份证号
            txt联系人姓名.Text = str姓名
            For j = 0 To cbo联系人关系.ListCount - 1
                If NeedName(cbo联系人关系.List(j)) = str关系 Then cbo联系人关系.ListIndex = j
            Next
            txt联系人电话.Text = str电话
        End If
        .Rows = .Rows + 1
    End With
End Sub

Private Sub SetOtherInfo(str信息名 As String, str信息值 As String, Optional ByVal byt信息更新模式 As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置其他情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str信息名 Then
                        If mlng病人ID <> 0 And byt信息更新模式 = 2 Then Exit Sub
                        .TextMatrix(i, j + 1) = str信息值
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" Or .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str信息名
                .TextMatrix(.Rows - 1, j + 1) = str信息值
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
        
    End With
End Sub
Private Sub Load健康卡相关信息(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人健康卡信息
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-12 14:55:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs过敏药物 As Recordset
    Dim rs免疫记录 As Recordset
    Dim rsABO血型 As Recordset
    Dim rsRH As Recordset
    Dim rs医学警示 As Recordset
    Dim rs其他医学警示 As Recordset
    Dim rs病人信息 As Recordset
    Dim rs联系人 As Recordset
    Dim rs其他信息 As Recordset
    Dim str医学警示 As String
    Dim str联系人姓名 As String
    Dim str联系人关系 As String
    Dim str联系人电话 As String
    Dim str联系人身份证号 As String
    Dim lng联系人数量 As Long
    Dim i As Long
    On Error GoTo ErrHandl:

    '获取过敏药物
    strSQL = "" & _
    "   Select 病人ID,过敏药物ID,过敏药物,过敏反应 From 病人过敏药物 Where 病人ID=[1]"
    Set rs过敏药物 = zlDatabase.OpenSQLRecord(strSQL, "病人过敏药物", lng病人ID)
    While rs过敏药物.EOF = False
        SetDrugAllergy Nvl(rs过敏药物!过敏药物), Nvl(rs过敏药物!过敏反应), Nvl(rs过敏药物!过敏药物ID, 0)
        rs过敏药物.MoveNext
    Wend
    '获取免疫记录
    strSQL = "" & _
    "   Select 病人ID,接种时间,接种名称 From 病人免疫记录 Where 病人ID=[1]"
    Set rs免疫记录 = zlDatabase.OpenSQLRecord(strSQL, "病人免疫记录", lng病人ID)
    While rs免疫记录.EOF = False
        SetInoculate Format(Nvl(rs免疫记录!接种时间), "YYYY-MM-DD"), Nvl(rs免疫记录!接种名称)
        rs免疫记录.MoveNext
    Wend
    '血型
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='血型'"
    Set rsABO血型 = zlDatabase.OpenSQLRecord(strSQL, "ABO血型", lng病人ID)
    While rsABO血型.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            '76314,李南春，2014-08-06，正确获取病人信息
            If NeedName(cboBloodType.List(i)) = NeedName(Nvl(rsABO血型!信息值)) Then cboBloodType.ListIndex = i
        Next
        rsABO血型.MoveNext
    Wend
    'RH
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='RH'"
    Set rsRH = zlDatabase.OpenSQLRecord(strSQL, "RH", lng病人ID)
    While rsRH.EOF = False
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = Nvl(rsRH!信息值) Then cboBH.ListIndex = i
        Next
        rsRH.MoveNext
    Wend
    '医学警示
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='医学警示' "
    Set rs医学警示 = zlDatabase.OpenSQLRecord(strSQL, "医学警示", lng病人ID)
    While rs医学警示.EOF = False
        str医学警示 = str医学警示 & ";" & Nvl(rs医学警示!信息值)
        rs医学警示.MoveNext
    Wend
    If str医学警示 <> "" Then str医学警示 = Mid(str医学警示, 2)
    txtMedicalWarning.Text = str医学警示
    txtMedicalWarning.Tag = txtMedicalWarning.Text
    '其他医学警示
    strSQL = "" & _
    "  Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='其他医学警示' "
    Set rs其他医学警示 = zlDatabase.OpenSQLRecord(strSQL, "其他医学警示", lng病人ID)
    While rs其他医学警示.EOF = False
        txtOtherWaring.Text = Nvl(rs其他医学警示!信息值)
        rs其他医学警示.MoveNext
    Wend
    '联系人相关信息
    '取病人信息表中的联系人信息
    strSQL = "" & _
    "   Select  联系人姓名,联系人关系,联系人电话,联系人身份证号 From 病人信息 Where 病人ID=[1] And Not 联系人姓名 is Null"
    Set rs病人信息 = zlDatabase.OpenSQLRecord(strSQL, "病人信息联系人信息", lng病人ID)
        If rs病人信息.EOF = False Then
        txt联系人身份证.Text = Nvl(rs病人信息!联系人身份证号)
        txt联系人姓名.Text = Nvl(rs病人信息!联系人姓名)
        For i = 0 To cbo联系人关系.ListCount - 1
            If NeedName(cbo联系人关系.List(i)) = Nvl(rs病人信息!联系人关系) Then cbo联系人关系.ListIndex = i
        Next
        txt联系人电话.Text = Nvl(rs病人信息!联系人电话)
        
        SetLinkInfo Nvl(rs病人信息!联系人姓名), Nvl(rs病人信息!联系人关系), Nvl(rs病人信息!联系人电话), Nvl(rs病人信息!联系人身份证号)
    End If
    '取病人信息从表中的联系人信息
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 like '联系人%' order by 信息名 Asc"
    Set rs联系人 = zlDatabase.OpenSQLRecord(strSQL, "联系人相关信息", lng病人ID)
    If rs联系人.EOF = False Then
        rs联系人.Filter = "信息名 like '联系人姓名%'"
        lng联系人数量 = rs联系人.RecordCount
        rs联系人.Filter = ""
        For i = 2 To lng联系人数量 + 1
            While rs联系人.EOF = False
                Select Case Nvl(rs联系人!信息名)
                    Case "联系人姓名" & i
                        str联系人姓名 = Nvl(rs联系人!信息值)
                    Case "联系人关系" & i
                        str联系人关系 = Nvl(rs联系人!信息值)
                    Case "联系人电话" & i
                        str联系人电话 = Nvl(rs联系人!信息值)
                    Case "联系人身份证号" & i
                        str联系人身份证号 = Nvl(rs联系人!信息值)
                End Select
                rs联系人.MoveNext
            Wend
            SetLinkInfo str联系人姓名, str联系人关系, str联系人电话, str联系人身份证号
            rs联系人.MoveFirst
        Next
    End If
    '其他信息
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 Not in ('血型','RH','医学警示','其他医学警示','身份证号状态','外籍身份证号') And 信息名 Not like '联系人%'"
    Set rs其他信息 = zlDatabase.OpenSQLRecord(strSQL, "联系人其他信息", lng病人ID)
    While rs其他信息.EOF = False
            SetOtherInfo rs其他信息!信息名, rs其他信息!信息值
        rs其他信息.MoveNext
    Wend

    Exit Sub
ErrHandl:
    If ErrCenter() = 1 Then
       Resume
    End If
End Sub

Private Sub Add健康卡相关信息(ByVal lng病人ID As Long, ByRef colPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:健康卡数据处理
    '入参:
    '编制:56599
    '日期:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim varKey As Variant
    '过敏药物
    With vsDrug
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人过敏药物_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人过敏药物_Update("
                    '病人ID_In 病人过敏药物.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '过敏药物ID_In 病人过敏药物.过敏药物ID%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 2) = "", "", .TextMatrix(i, 2)) & "',"
                    '过敏药物_In  病人过敏药物.过敏药物%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '过敏反应_In 病人过敏反应.过敏反应%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '接种信息
    With vsInoculate
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人免疫记录_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '其他信息
    'ABO血型
    '病人信息从表
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'血型',"
    '信息值_In 病人信息从表.信息值%Type
    '76314,李南春，2014-08-06，正确获取病人信息
    strSQL = strSQL & "'" & NeedName(cboBloodType.Text) & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'RH',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtMedicalWarning.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '其他医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'其他医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtOtherWaring.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    
    '联系人相关信息
    With vsLinkMan
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '联系人姓名不能为空
                    For j = 0 To .Cols - 1
                        strSQL = "Zl_病人信息从表_Update("
                        '病人ID_In 病人信息从表.病人Id%Type
                        strSQL = strSQL & "" & lng病人ID & ","
                        '信息名_In 病人信息从表.信息名%Type
                        strSQL = strSQL & "'" & .TextMatrix(0, j) & i & "',"
                        '信息值_In 病人信息从表.信息值%Type
                        strSQL = strSQL & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '就诊Id_In 病人信息从表.就诊Id%Type
                        strSQL = strSQL & "'')"

                        zlAddArray colPro, strSQL
                    Next
                End If
            Next
        End If
    End With
    '其他信息
     With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 0) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 2) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
     End With
     '医疗卡属性
     If Not mdic医疗卡属性 Is Nothing And txt卡号.Text <> "" Then
        For Each varKey In mdic医疗卡属性.Keys
            strSQL = "Zl_病人医疗卡属性_Update("
            strSQL = strSQL & lng病人ID & ","
            strSQL = strSQL & mCurSendCard.lng卡类别ID & ","
            strSQL = strSQL & "'" & Trim(txt卡号.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdic医疗卡属性(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
End Sub

Private Function CheckPatiCard() As Boolean
'功能：检查病人健康卡片录入的内容是否合法
'63246:刘鹏飞,2013-07-03
    Dim intLen As Integer, i As Integer, j As Integer
    
    intLen = 100
    If LenB(StrConv(txtMedicalWarning.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "医学警示只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        If txtMedicalWarning.Enabled And txtMedicalWarning.Visible Then txtMedicalWarning.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtOtherWaring.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "其他医学警示只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
        Exit Function
    End If
    
    mblnCheckPatiCard = True
    '过敏药物
    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 60
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "过敏药物只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第1列", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "过敏反应只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第2列", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '接种信息
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    If Not IsDate(.TextMatrix(i, 0)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种时间不是有效的日期格式！" & vbCrLf & "错误行:第" & i & "行、第1列", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种名称只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第2列", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
                If .TextMatrix(i, 3) <> "" Then
                    If Not IsDate(.TextMatrix(i, 2)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种时间不是有效的日期格式！" & vbCrLf & "错误行:第" & i & "行、第3列", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种名称只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第4列", vbInformation, gstrSysName
                        Call .Select(i, 3, i, 3)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '联系人地址
    With vsLinkMan
        intLen = 100
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    For j = 0 To .Cols - 1
                        If .ColIndex("联系人姓名") = j Then
                            intLen = 64
                        ElseIf .ColIndex("联系人身份证号") = j Then
                            intLen = 18
                        ElseIf .ColIndex("联系人电话") = j Then
                            intLen = 20
                        Else
                            intLen = 100
                        End If
                        If LenB(StrConv(.TextMatrix(i, j), vbFromUnicode)) > intLen Then
                            tbcPage.Item(1).Selected = True
                            MsgBox .TextMatrix(0, j) & "只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行", vbInformation, gstrSysName
                            Call .Select(i, j, i, j)
                            .TopRow = i
                            If .Enabled = True And .Visible = True Then .SetFocus
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If
    End With
    
    '其他信息
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 20
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "信息名只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第1列", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "信息值只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第2列", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
                If .TextMatrix(i, 2) <> "" Then
                    intLen = 20
                    If LenB(StrConv(.TextMatrix(i, 2), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "信息名只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第3列", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "信息值只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第4列", vbInformation, gstrSysName
                        Call .Select(i, 3, i, 3)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
     End With
     
     mblnCheckPatiCard = False
     tbcPage.Item(0).Selected = True
     CheckPatiCard = True
End Function

Private Function LoadPati(ByVal strPatiXML As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息,读取病人信息
    '编制:刘兴洪
    '日期:2011-09-08 21:52:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '问题号:56599
    Dim str过敏药物 As String, str过敏反应 As String '问题号:56599
    Dim str接种日期 As String, str接种名称 As String '问题号:56599
    Dim strABO血型 As String '问题号:56599
    Dim str信息名 As String, str信息值 As String '问题号:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '问题号:56599
    Dim str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str地址 As String '问题号:56599
    Dim byt信息更新模式 As Byte
    On Error GoTo errHandle

    If strPatiXML = "" Then Exit Function
    
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    '    标识    数据类型    长度    精度    说明
    '    信息更新模式 Integer 1 '0-强制更新，1-建档病人不更新，2-建档病人信息补缺
    Call zlXML_GetNodeValue("信息更新模式", , strValue)
    byt信息更新模式 = 0
    byt信息更新模式 = Val(strValue)
    If byt信息更新模式 = 1 And mlng病人ID <> 0 Then LoadPati = True: Exit Function
    '    卡号    Varchar2    20
    Call zlXML_GetNodeValue("卡号", , strValue)
    '    姓名    Varchar2    64
    Call zlXML_GetNodeValue("姓名", , strValue)
    '1-必须更新
    '2-新病人
    '3-老病人信息为空的情况补缺
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txtPatient.Text = "") Then txtPatient.Text = strValue
    '    性别    Varchar2    4
    Call zlXML_GetNodeValue("性别", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cbo性别.Text = "") Then
        If strValue <> "" Then
            Call zlControl.CboLocate(cbo性别, strValue)
            If cbo性别.ListIndex = -1 Then
                cbo性别.AddItem strValue
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
        End If
    End If
    '    年龄    Varchar2    10
    Call zlXML_GetNodeValue("年龄", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt年龄.Text = "") Then
        If strValue <> "" Then
            Call LoadOldData(strValue, txt年龄, cbo年龄单位)
        End If
    End If
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("出生日期", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt出生日期.Text = "") Then
        mblnChange = False
        txt出生日期.Text = Format(IIf(strValue = "", "____-__-__", strValue), "YYYY-MM-DD")
        mblnChange = True
        If strValue <> "" Then
            txt年龄.Text = ReCalcOld(CDate(Format(strValue, "YYYY-MM-DD HH:MM:SS")), cbo年龄单位, , , CDate(txt出生日期.Tag))   '修改的时候,根据出生日期重算年龄
            If CDate(txt出生日期.Text) - CDate(strValue) <> 0 Then
                mblnChange = False
                txt出生时间.Text = Format(strValue, "HH:MM")
                mblnChange = True
            End If
        Else
            txt出生时间.Text = "__:__"
            mblnChange = False
            Call ReCalcBirthDay
            mblnChange = True
        End If
    End If
    '    出生地点    Varchar2    50
    Call zlXML_GetNodeValue("出生地点", , strValue)
    '    身份证号    VARCHAR2    18
    Call zlXML_GetNodeValue("身份证号", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt身份证号.Text = "") Then
        If strValue <> "" Then txt身份证号.Text = strValue
    End If
    '    其他证件    Varchar2    20
    Call zlXML_GetNodeValue("其他证件", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt其他证件.Text = "") Then
        If strValue <> "" Then txt其他证件.Text = strValue
    End If
    '    职业    Varchar2    80
    Call zlXML_GetNodeValue("职业", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cbo职业.Text = "") Then
        If strValue <> "" Then
            cbo职业.ListIndex = GetCboIndex(cbo职业, strValue, , , mstrCboSplit)
            If cbo职业.ListIndex = -1 Then
                cbo职业.AddItem strValue, 0
                cbo职业.ListIndex = cbo职业.NewIndex
            End If
        End If
    End If
    '    民族    Varchar2    20
    Call zlXML_GetNodeValue("民族", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cbo民族.Text = "") Then
        cbo民族.ListIndex = GetCboIndex(cbo民族, strValue)
        If cbo民族.ListIndex = -1 And strValue <> "" Then
            cbo民族.AddItem strValue, 0
            cbo民族.ListIndex = cbo民族.NewIndex
        End If
    End If
    '    国籍    Varchar2    30
    Call zlXML_GetNodeValue("国籍", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cbo国籍.Text = "") Then
        cbo国籍.ListIndex = GetCboIndex(cbo国籍, strValue)
        If cbo国籍.ListIndex = -1 And strValue <> "" Then
            cbo国籍.AddItem strValue, 0
            cbo国籍.ListIndex = cbo国籍.NewIndex
        End If
    End If
    '    学历    Varchar2    10
    Call zlXML_GetNodeValue("学历", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cbo学历.Text = "") Then
        cbo学历.ListIndex = GetCboIndex(cbo学历, strValue)
        If cbo学历.ListIndex = -1 And strValue <> "" Then
            cbo学历.AddItem strValue, 0
            cbo学历.ListIndex = cbo学历.NewIndex
        End If
    End If
    '    婚姻状况    Varchar2    4
    Call zlXML_GetNodeValue("婚姻状况", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cbo婚姻状况.Text = "") Then
        cbo婚姻状况.ListIndex = GetCboIndex(cbo婚姻状况, strValue)
        If cbo婚姻状况.ListIndex = -1 And strValue <> "" Then
            cbo婚姻状况.AddItem strValue, 0
            cbo婚姻状况.ListIndex = cbo婚姻状况.NewIndex
        End If
    End If
    '    区域    Varchar2    30
    Call zlXML_GetNodeValue("区域", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt区域.Text = "") Then txt区域.Text = strValue
    '    家庭地址    Varchar2    50
    Call zlXML_GetNodeValue("家庭地址", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt家庭地址.Text = "") Then
        txt家庭地址.Text = strValue
        If gbln启用结构化地址 Then PatiAddress(E_IX_现住址).Value = strValue
    End If
    '    家庭电话    Varchar2    20
    Call zlXML_GetNodeValue("家庭电话", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt家庭电话.Text = "") Then txt家庭电话.Text = strValue
    '    家庭地址邮编    Varchar2    6
    Call zlXML_GetNodeValue("家庭地址邮编", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt家庭地址邮编.Text = "") Then txt家庭地址邮编.Text = strValue
    '    户口地址    Varchar2    50
    Call zlXML_GetNodeValue("户口地址", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt户口地址.Text = "") Then txt户口地址.Text = strValue
    '    户口地址    Varchar2    50
    Call zlXML_GetNodeValue("户口地址", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt户口地址.Text = "") Then
        txt户口地址.Text = strValue
        If gbln启用结构化地址 Then PatiAddress(E_IX_户口地址).Value = strValue
    End If
    '    监护人  Varchar2    64
    Call zlXML_GetNodeValue("监护人", , strValue)
   'txt监护人.Text = strValue
    '    工作单位    Varchar2    100
    Call zlXML_GetNodeValue("工作单位", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt工作单位.Text = "") Then
        txt工作单位.Text = strValue
        lbl工作单位.Tag = ""
    End If
    '    单位电话    Varchar2    20
    Call zlXML_GetNodeValue("单位电话", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt单位电话.Text = "") Then txt单位电话.Text = strValue
    '手机号   Varchar2    20
    Call zlXML_GetNodeValue("手机号", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txtMobile.Text = "") Then txtMobile.Text = strValue
    '    单位邮编    Varchar2    6
    Call zlXML_GetNodeValue("单位邮编", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt单位邮编.Text = "") Then txt单位邮编.Text = strValue
    '    单位开户行  Varchar2    50
    Call zlXML_GetNodeValue("单位开户行", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt单位开户行.Text = "") Then txt单位开户行.Text = strValue
    '    单位帐号    Varchar2    20
    Call zlXML_GetNodeValue("单位帐号", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt单位帐号.Text = "") Then txt单位帐号.Text = strValue
    '问题号:56599
    '过敏情况
    Call zlXML_GetRows("药物名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("药物名称", i, str过敏药物)
        Call zlXML_GetNodeValue("药物反应", i, str过敏反应)
        SetDrugAllergy str过敏药物, str过敏反应, , byt信息更新模式
    Next
    lngCount = 0
    '免疫记录
    Call zlXML_GetRows("疫苗名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("疫苗名称", i, str接种名称)
        Call zlXML_GetNodeValue("接种时间", i, str接种日期)
        SetInoculate str接种日期, str接种名称
    Next
    lngCount = 0
    'ABO血型
    Call zlXML_GetNodeValue("ABO血型", , strABO血型)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cboBloodType.Text = "") Then
        If strABO血型 <> "" Then
            For i = 0 To cboBloodType.ListCount - 1
                '76314,李南春，2014-08-06，正确获取病人信息
                If NeedName(cboBloodType.List(i)) = NeedName(strABO血型) Then cboBloodType.ListIndex = i
            Next
        End If
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And cboBH.Text = "") Then
        If strValue <> "" Then
            For i = 0 To cboBH.ListCount - 1
                If cboBH.List(i) = strValue Then cboBH.ListIndex = i
            Next
        End If
    End If
    '医学警示
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("临床基本信息")
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txtMedicalWarning.Text = "") Then
        If Not xmlChildNodes Is Nothing Then
            If xmlChildNodes.length > 0 Then
                For i = 0 To xmlChildNodes.length - 1
                    Set xmlChildNode = xmlChildNodes(i)
                    If xmlChildNode.Text = "1" Then
                        strValue = strValue & ";" & Replace(xmlChildNode.nodeName, "标志", "")
                    End If
                Next
            End If
        End If
        If strValue <> "" Then txtMedicalWarning.Text = Mid(strValue, 2)
   End If
    
    '其他医学警示
    Call zlXML_GetNodeValue("其他医学警示", , strValue)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txtOtherWaring.Text = "") Then
        If strValue <> "" Then txtOtherWaring.Text = strValue
    End If
    '联系信息
    '    联系人地址  Varchar2    50
    Call zlXML_GetNodeValue("联系人地址", , str地址)
    If byt信息更新模式 = 0 Or mlng病人ID = 0 Or (mlng病人ID <> 0 And byt信息更新模式 = 2 And txt联系人地址.Text = "") Then
        txt联系人地址.Text = str地址
        If gbln启用结构化地址 Then PatiAddress(E_IX_联系人地址).Value = str地址
    End If
     '    联系人姓名  Varchar2    64
    Call zlXML_GetNodeValue("联系人姓名", , str姓名)
    '    联系人关系  Varchar2    30
    Call zlXML_GetNodeValue("联系人关系", , str关系)
    '    联系人电话  Varchar2    20
    Call zlXML_GetNodeValue("联系人电话", , str电话)
    '    联系人身份证 Varchar2   20
    Call zlXML_GetNodeValue("联系人身份证号", , str身份证号)
    SetLinkInfo str姓名, str关系, str电话, str身份证号, byt信息更新模式
    
    Call zlXML_GetRows("联系信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("联系信息", "姓名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("联系信息", "姓名", i, j, str姓名)
                Call zlXML_GetChildNodeValue("联系信息", "关系", i, j, str关系)
                Call zlXML_GetChildNodeValue("联系信息", "电话", i, j, str电话)
                Call zlXML_GetChildNodeValue("联系信息", "身份证号", i, j, str身份证号)
                SetLinkInfo str姓名, str关系, str电话, str身份证号, byt信息更新模式
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '其他信息
    '健康档案编号
    Call zlXML_GetNodeValue("健康档案编号", , strValue)
    SetOtherInfo "健康档案编号", strValue, byt信息更新模式
    
    '新农合证号
    Call zlXML_GetNodeValue("新农合证号", , strValue)
    SetOtherInfo "新农合证号", strValue, byt信息更新模式

    '其他证件
    Call zlXML_GetRows("其他证件", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他证件", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他证件", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他证件", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值, byt信息更新模式
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '其他信息
    Call zlXML_GetRows("其他信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他信息", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他信息", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他信息", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值, byt信息更新模式
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '医疗卡属性
    Call zlXML_GetRows("医疗卡属性", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("医疗卡属性", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息值", i, j, str信息值)
                If mdic医疗卡属性.Exists(str信息名) Then
                    If Not (mlng病人ID <> 0 And byt信息更新模式 = 2) Then mdic医疗卡属性.Item(str信息名) = str信息值
                Else
                    mdic医疗卡属性.Add str信息名, str信息值
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
Public Function GetCboIndex(cbo As ComboBox, strFind As String, _
    Optional blnKeep As Boolean, _
    Optional blnLike As Boolean, Optional strSplit As String = "-") As Long
'功能：由字符串在ComboBox中查找索引
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '先精确查找
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), strSplit) > 0 Then
            If NeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '最后模糊查找
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Sub Clear健康档案()
    '---------------------------------------------------------------------------------------------------------------------------------------------
'功能:判断当前是否为卡发操作 (不是发卡操作既是绑定卡操作)
'入参:
'编制:56599
'日期:2012-12-25 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    '血型
    Call SetCboDefault(cboBloodType)
    'RH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
    '医学警示
    txtMedicalWarning.Text = ""
    '其他医学警示
    txtOtherWaring.Text = ""
    '联系人信息
    With vsLinkMan
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '接种情况
    With vsInoculate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '其他信息
    With vsOtherInfo
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '过敏反应
    With vsDrug
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
    End With
End Sub
Private Function SetCreateCardObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置制卡对象
    '编制:王吉
    '日期:2012-12-17 11:06:41
    '问题号:56599
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand:
    If mobjHealthCard Is Nothing Then
        Set mobjHealthCard = CreateObject("zl9Card_HealthCard.clsHealthCard")
    End If
    SetCreateCardObject = True
    Exit Function
ErrHand:
    SetCreateCardObject = False
End Function

Public Function zlCreateSquare() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建医疗卡对象
    '编制:李南春
    '日期:2016/6/21 11:57:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not mobjSquare Is Nothing Then zlCreateSquare = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjSquare = CreateObject("zl9CardSquare.clsCardsquare")
    If Err <> 0 Then Err = 0: Exit Function
    Call mobjSquare.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend)
    '初始部件不成功,则作为不存在处理
    zlCreateSquare = True
End Function

Private Function WriteCard(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写卡
    '入参:lng病人ID - 病人ID
    '编制:王吉
    '问题:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    WriteCard = mobjSquare.zlBandCardArfter(Me, mlngModul, mCurSendCard.lng卡类别ID, lng病人ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
    If ErrCenter() = 1 Then Resume
End Function

Private Function Check发卡性质(lng病人ID As Long, lng卡类别ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发卡时检查是否限制病人的发卡张数
    '入参:lng病人ID - 病人ID;lng卡类别ID  - 医疗卡的类别ID
    '编制:王吉
    '问题:57326
    '日期:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
        strSQL = "Select Count(1) as 存在 From 病人医疗卡信息 Where 状态=0 And 病人ID=[1] And 卡类别ID=[2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng卡类别ID)
        If Val(Nvl(rsTemp!存在)) <= 0 Then Check发卡性质 = True: Exit Function
        Select Case mCurSendCard.lng发卡性质
            Case 0 '不限制
                Check发卡性质 = True
            Case 1 '同一个病人只允许发一张卡
                MsgBox "该病人已经发过" & mCurSendCard.str卡名称 & ",不能在进行发卡操作!", vbInformation + vbOKOnly
                Check发卡性质 = False
            Case 2 '同一个病人允许发多张卡,但需要提醒
               Check发卡性质 = MsgBox("该病人已经发过" & mCurSendCard.str卡名称 & ",是否要进行发卡操作?", vbQuestion + vbYesNo) = vbYes
        End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function WhetherTheCardBinding(ByVal str卡号 As String, ByVal lng卡类别 As Long, Optional ByRef lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取指定卡号是否已经发卡
'入参:str卡号：卡号 ，lng卡类别：卡类别 , lngPatientID :病人ID
'返回:True :已经发卡;False:未发卡
'编制:
'日期:
'问题号:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl
    strSQL = "" & _
           "   Select 病人ID From 病人医疗卡信息 Where 卡号=[1]  And 卡类别ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "门诊挂号", str卡号, lng卡类别)
    WhetherTheCardBinding = rsTemp.RecordCount > 0

    If rsTemp.RecordCount > 0 Then
        lngPatientID = Val(Nvl(rsTemp!病人ID))
    End If

    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function GetCardLastChangeType(ByVal str卡号 As String, ByVal lng卡类别 As Long, ByVal lngPaitentID As Long) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取卡最后的变动类型
'入参:str卡号：卡号 ，lng卡类别：卡类别 , lngPatientID :病人ID
'返回:0-未找到相关信息   1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
'编制:李光福
'日期:2013-2-4 17:36:33
'问题号:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "     Select 变动类别" & vbNewLine & _
           "    From (With 医疗卡变动 As (Select 病人id, ID, 变动类别, 变动时间 " & vbNewLine & _
           "                              From 病人医疗卡变动 Bd" & vbNewLine & _
           "                              Where Bd.卡号 = [2] And 卡类别id = [1] And 病人id = [3])" & vbNewLine & _
           "           Select A.变动类别" & vbNewLine & _
           "           From 医疗卡变动 A, (Select Max(变动时间) As 变动时间 From 医疗卡变动 C) B" & vbNewLine & _
           "           Where A.变动时间 = B.变动时间) A"
    On Error GoTo ErrHand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取卡最后变动信息", lng卡类别, str卡号, lngPaitentID)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            GetCardLastChangeType = Val(Nvl(rsTmp!变动类别))
        End If
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function BlandCancel(ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:取消绑定卡
'入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
'返回:取消成功,返回true,否则返回False
'编制:刘兴洪
'日期:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curDate As Date
    Dim strSQL As String, strPassWord As String

    On Error GoTo errHandle

    curDate = zlDatabase.Currentdate

    'Zl_医疗卡变动_Insert
    strSQL = "Zl_医疗卡变动_Insert("
    '      变动类型_In   Number,
    '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
    strSQL = strSQL & "" & 14 & ","
    '      病人id_In     住院费用记录.病人id%Type,
    strSQL = strSQL & "" & lngPatientID & ","
    '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '      原卡号_In     病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "NULL,"
    '      医疗卡号_In   病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "'" & strCardNo & "'" & ","
    '      变动原因_In   病人医疗卡变动.变动原因%Type,
    strSQL = strSQL & "'卡重复利用自动取消原卡绑定信息',"
    '      密码_In       病人信息.卡验证码%Type,
    strSQL = strSQL & "NULL,"
    '      操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSQL = strSQL & "NULL,"
    '      变动时间_In   住院费用记录.登记时间%Type,
    strSQL = strSQL & "to_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
    strSQL = strSQL & "NULL)"

     
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    BlandCancel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub MoveNextCell(ByVal VsfData As VSFlexGrid, Optional ByVal blnNext As Boolean = True)

    Dim intRow As Integer
    If blnNext = True Then
toMoveNextCol:
        If VsfData.Col < VsfData.Cols - 1 Then
            VsfData.Col = VsfData.Col + 1
            If VsfData.ColWidth(VsfData.Col) = 0 Or VsfData.ColHidden(VsfData.Col) Then
                GoTo toMoveNextCol
            End If
        Else
toMoveNextRow:
            '跳到下一行
            intRow = 1
            If VsfData.Row + intRow < VsfData.Rows Then
                VsfData.Row = VsfData.Row + intRow
            End If
            If VsfData.RowHidden(VsfData.Row) Then
                If VsfData.Row < VsfData.Rows - 1 Then
                    GoTo toMoveNextRow
                Else
                    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
                        If VsfData.RowHidden(intRow) = False Then
                            VsfData.Row = intRow
                            Exit For
                        End If
                    Next intRow
                End If
            End If
            VsfData.Col = VsfData.FixedCols
        End If
    Else
toMovePrevCol:
        If VsfData.Col > VsfData.FixedCols Then
            VsfData.Col = VsfData.Col - 1
            If VsfData.ColWidth(VsfData.Col) = 0 Or VsfData.ColHidden(VsfData.Col) Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '跳到上一行
            intRow = 1
            If VsfData.Row - intRow >= VsfData.FixedRows Then
                VsfData.Row = VsfData.Row - intRow
            End If
            If VsfData.RowHidden(VsfData.Row) Then
                If VsfData.Row > VsfData.FixedRows Then
                    GoTo toMovePrevRow
                Else
                    For intRow = VsfData.FixedRows To VsfData.Rows - 1
                        If VsfData.RowHidden(intRow) = False Then
                            VsfData.Row = intRow
                            Exit For
                        End If
                    Next intRow
                End If
            End If
            VsfData.Col = VsfData.FixedCols
        End If
    End If

    If VsfData.ColIsVisible(VsfData.Col) = False Then
        VsfData.LeftCol = VsfData.Col
    End If
    If VsfData.RowIsVisible(VsfData.Row) = False Then
        VsfData.TopRow = VsfData.Row
    End If
End Sub

Private Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If objCtl.Enabled And objCtl.Visible = True Then: objCtl.SetFocus
End Sub

Private Sub vsInoculate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsInoculate_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strCurDate As String
    
    If Col = 1 Or Col = 3 Then '接种名称列编辑时需判断是否字数超过了200
        With vsInoculate
           If LenB(StrConv(.EditText, vbFromUnicode)) > 200 Then
                If MsgBox("接种名称输入字符超出最大字符数200,请问是否将多出的字符将被自动截除？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    .EditText = StrConv(MidB(StrConv(.EditText, vbFromUnicode), 1, 200), vbUnicode)
                Else
                    Cancel = True
                End If
           End If
        End With
    Else
        With vsInoculate
            If IsDate(Format(.EditText, "YYYY-MM-DD")) = False And .EditText <> "    -  -  " Then
                 MsgBox "输入的日期格式不对或不是正确的日期，请检查！", vbInformation, gstrSysName
                 Cancel = True
            ElseIf .EditText = "    -  -  " Then
                 .EditText = ""
            Else
                If .EditText <> "" Then
                    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
                    If Format(.EditText, "YYYY-MM-DD") > strCurDate Then
                        MsgBox "接种日期不能大于服务器系统时间[" & strCurDate & "],请检查！", vbInformation, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = Format(.EditText, "YYYY-MM-DD")
                    
                    If Col = 2 And vsInoculate.Rows - 1 = Row Then
                        vsInoculate.Rows = vsInoculate.Rows + 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsLinkMan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsLinkMan
        If Row = .Rows - 1 And Col = .FixedCols And .TextMatrix(Row, Col) <> "" Then
            .Rows = .Rows + 1
        End If
    End With
End Sub

Private Sub vsLinkMan_GotFocus()
    If mblnCheckPatiCard = False Then
        vsLinkMan.Col = vsLinkMan.FixedCols
        vsLinkMan.Row = vsLinkMan.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub vsLinkMan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer, i As Integer
    With vsLinkMan
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Rows = .FixedRows + 1 Then
                .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
            Else
                Call .RemoveItem(.Row)
                If intRow >= .Rows Then
                    .Row = .Rows - 1
                Else
                    .Row = intRow
                End If
                .Col = .FixedCols
            End If
            If .Rows <= .FixedRows Then .Rows = .FixedRows + 1
            txt联系人姓名.Text = .TextMatrix(.FixedRows, .ColIndex("联系人姓名"))
            For i = 0 To cbo联系人关系.ListCount - 1
                If NeedName(cbo联系人关系.List(i)) = .TextMatrix(.FixedRows, .ColIndex("联系人关系")) Then
                    Exit For
                End If
            Next
            If i < cbo联系人关系.ListCount Then
                cbo联系人关系.ListIndex = i
            Else
                cbo联系人关系.ListIndex = -1
            End If
            
            txt联系人身份证.Text = .TextMatrix(.FixedRows, .ColIndex("联系人身份证号"))
            txt联系人电话.Text = .TextMatrix(.FixedRows, .ColIndex("联系人电话"))
            
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If .TextMatrix(.Row, .FixedCols) = "" And .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                Call MoveNextCell(vsLinkMan)
            End If
        End If
    End With
End Sub

Private Sub vsLinkMan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsLinkMan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsLinkMan
        If KeyAscii = vbKeyReturn Then Exit Sub
        If Col = .ColIndex("联系人身份证号") Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        ElseIf Col = .ColIndex("联系人电话") Then
            If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End With
End Sub

Private Sub vsLinkMan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Integer
    
    With vsLinkMan
        If Not Row = .FixedRows Then Exit Sub
        Select Case Col
            Case .ColIndex("联系人姓名")
                txt联系人姓名.Text = Trim(.EditText)
            Case .ColIndex("联系人关系")
                For i = 0 To cbo联系人关系.ListCount - 1
                    If NeedName(cbo联系人关系.List(i)) = Trim(.EditText) Then Exit For
                Next
                If i < cbo联系人关系.ListCount Then
                    cbo联系人关系.ListIndex = i
                Else
                    cbo联系人关系.ListIndex = -1
                End If
                
            Case .ColIndex("联系人身份证号")
                txt联系人身份证.Text = Trim(.EditText)
            Case .ColIndex("联系人电话")
                txt联系人电话.Text = Trim(.EditText)
        End Select
    End With
End Sub

Private Sub vsOtherInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 And vsOtherInfo.Rows - 1 = Row And vsOtherInfo.TextMatrix(Row, Col) <> "" Then
        vsOtherInfo.Rows = vsOtherInfo.Rows + 1
    End If
End Sub

Private Sub vsOtherInfo_GotFocus()
    If mblnCheckPatiCard = False Then
        vsOtherInfo.Col = vsOtherInfo.FixedCols
        vsOtherInfo.Row = vsOtherInfo.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub vsOtherInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    
    With vsOtherInfo
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Col > .FixedCols + 1 Then
                .TextMatrix(intRow, .FixedCols + 2) = ""
                .TextMatrix(intRow, .FixedCols + 3) = ""
            Else
                If .Rows = .FixedRows + 1 Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    Call .RemoveItem(.Row)
                    If intRow >= .Rows Then
                        .Row = .Rows - 1
                    Else
                        .Row = intRow
                    End If
                    .Col = .FixedCols
                End If
            End If
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If ((.TextMatrix(.Row, .FixedCols) = "" And .Col = .FixedCols) Or (.Col = .FixedCols + 2 And .TextMatrix(.Row, .FixedCols + 2) = "") Or .Col = .FixedCols + 3) And .Row = .Rows - 1 Then
                If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
            Else
               Call MoveNextCell(vsOtherInfo)
            End If
        End If
    End With
End Sub

Private Sub vsOtherInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub InitStructAddress()
'功能:根据是否启用结构化地址调整界面
    Dim i As Long
    
    If gbln启用结构化地址 Then
        For i = PatiAddress.LBound To PatiAddress.UBound
            If i = E_IX_现住址 Or i = E_IX_户口地址 Or i = E_IX_联系人地址 Then
                PatiAddress(i).Items = Five
            End If
            PatiAddress(i).TextBackColor = &H80000005
            PatiAddress(i).Visible = True
            PatiAddress(i).ShowTown = gbln显示乡镇
        Next
        For i = LBound(marrAddress) To UBound(marrAddress)
            marrAddress(i) = ""
        Next
        txt家庭地址.Visible = False
        cmd家庭地址.Visible = False
        txt出生地点.Visible = False
        cmd出生地点.Visible = False
        txt户口地址.Visible = False
        cmd户口地址.Visible = False
        txt籍贯.Visible = False
        cmd籍贯.Visible = False
        txt联系人地址.Visible = False
        cmd联系人地址.Visible = False
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = False
        Next
        
        txt家庭地址.Visible = True
        cmd家庭地址.Visible = True
        txt出生地点.Visible = True
        cmd出生地点.Visible = True
        txt户口地址.Visible = True
        cmd户口地址.Visible = True
        txt籍贯.Visible = True
        cmd籍贯.Visible = True
        txt联系人地址.Visible = True
        cmd联系人地址.Visible = True
    End If
End Sub

Private Sub SetStrutAddress(Optional ByVal bytFunc As Byte)
'功能:89980病人结构化
'bytFunc=1 清空数据
'       =2 设置户口地址和家庭地址的缺省值
    Dim i As Long
    
    If bytFunc = 2 Then
        txt家庭地址.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        txt户口地址.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        Call PatiAddress(E_IX_现住址).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
        Call PatiAddress(E_IX_户口地址).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
            If bytFunc = 1 Then
                PatiAddress(i).Value = ""
            Else
                PatiAddress(i).Enabled = (mbytInState <> EState.E查阅)
            End If
        Next
    End If
End Sub

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡接口
    '返回: true-成功，false-失败
    '编制:李南春
    '日期:2016/6/20 13:54:56
    '问题:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    If txt卡号.Locked Then Exit Function
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    '不是医疗卡或者不是刷卡
    If mCurSendCard.lng卡类别ID = 0 Or Not mCurSendCard.bln刷卡 Then Exit Function
    If mobjSquare.zlSetBrushCardObject(mCurSendCard.lng卡类别ID, IIf(blnComm, txt卡号, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call mobjSquare.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function

Private Sub ReCalcBirthDay()
    Dim strBirth As String

    If CreatePublicPatient() Then
        If gobjPublicPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, Trim(cbo年龄单位.Text), ""), strBirth) Then
            If txt出生日期.Enabled Then txt出生日期.Text = Format(strBirth, "YYYY-MM-DD")
            
            If txt出生时间.Enabled Then
                strBirth = Format(strBirth, "HH:MM")
                txt出生时间.Text = IIf(strBirth = "00:00", "__:__", strBirth)
            End If
        End If
    End If
End Sub

Private Function CheckMobile(ByVal strMobile As String, ByVal lngPatiID As Long) As Boolean
'功能:检查当前手机号是否存在
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM 病人信息 Where 手机号 = [1] And 病人ID <> [2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查手机号", strMobile, lngPatiID)
    If Not rsTemp Is Nothing Then
        CheckMobile = rsTemp.EOF = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub EMPI_LoadPati()
'功能:将EMPI返回来的病人信息更新到界面
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim str出生日期 As String
    Dim blnRet As Boolean

    If CreatePlugInOK(glngModul) Then
        '组织病人基本信息
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "病人ID", adBigInt
            .Append "主页ID", adBigInt
            .Append "挂号ID", adBigInt
            '-------------------------------
            .Append "门诊号", adVarChar, 18
            .Append "住院号", adVarChar, 18
            .Append "医保号", adVarChar, 30
            .Append "身份证号", adVarChar, 18
            .Append "其他证件", adVarChar, 20
            .Append "姓名", adVarChar, 100
            .Append "性别", adVarChar, 4
            .Append "出生日期", adVarChar, 20 '日期格式：YYYY-MM-DD HH:MM:SS
            .Append "出生地点", adVarChar, 100
            .Append "国籍", adVarChar, 30
            .Append "民族", adVarChar, 20
            .Append "学历", adVarChar, 10
            .Append "职业", adVarChar, 80
            .Append "工作单位", adVarChar, 100
            .Append "邮箱", adVarChar, 30
            .Append "婚姻状况", adVarChar, 4
            .Append "家庭电话", adVarChar, 20
            .Append "联系人电话", adVarChar, 20
            .Append "单位电话", adVarChar, 20
            .Append "家庭地址", adVarChar, 100
            .Append "家庭地址邮编", adVarChar, 6
            .Append "户口地址", adVarChar, 100
            .Append "户口地址邮编", adVarChar, 6
            .Append "单位邮编", adVarChar, 6
            .Append "联系人地址", adVarChar, 100
            .Append "联系人关系", adVarChar, 30
            .Append "联系人姓名", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open
        
        If txt出生时间 = "__:__" Then
            str出生日期 = IIf(IsDate(txt出生日期.Text), Format(txt出生日期.Text, "YYYY-MM-DD"), "")
        Else
            str出生日期 = IIf(IsDate(txt出生日期.Text), Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS"), "")
        End If
 
        With rsPatiIn
            .AddNew
            !病人ID = CLng(txt病人ID.Text)
            !主页ID = mlng主页ID
            !住院号 = IIf(txt住院号.Text <> "", txt住院号.Text, "")
            !门诊号 = IIf(Trim(txt门诊号.Text) <> "", Trim(txt门诊号.Text), "")
            !医保号 = txtPatiMCNO(0).Text
            '-要更新的字段--------------------------------------------
            !身份证号 = Trim(txt身份证号.Text)
            !其他证件 = Trim(txt其他证件.Text)
            !姓名 = Trim(txtPatient.Text)
            !性别 = zlCommFun.GetNeedName(cbo性别.Text)
            !出生日期 = str出生日期 '日期格式：YYYY-MM-DD HH:MM:SS
            !出生地点 = Trim(txt出生地点.Text)
            !国籍 = zlCommFun.GetNeedName(cbo国籍.Text)
            !民族 = zlCommFun.GetNeedName(cbo民族.Text)
            !学历 = zlCommFun.GetNeedName(cbo学历.Text)
            !职业 = zlCommFun.GetNeedName(cbo职业.Text)
            !工作单位 = Trim(txt工作单位.Text)
            !婚姻状况 = zlCommFun.GetNeedName(cbo婚姻状况.Text)
            !家庭电话 = Trim(txt家庭电话.Text)
            !联系人电话 = Trim(txt联系人电话.Text)
            !单位电话 = Trim(txt单位电话.Text)
            !家庭地址 = Trim(txt家庭地址.Text)
            !家庭地址邮编 = Trim(txt家庭地址邮编.Text)
            !户口地址 = Trim(txt户口地址.Text)
            !户口地址邮编 = Trim(txt户口地址邮编.Text)
            !单位邮编 = Trim(txt单位邮编.Text)
            !联系人地址 = Trim(txt联系人地址.Text)
            !联系人关系 = zlCommFun.GetNeedName(cbo联系人关系.Text)
            !联系人姓名 = Trim(txt联系人姓名.Text)
            .Update
            '-------------------------------------------------------
        End With
        
        '调用查询接口
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsPatiIn, rsPatiOut)
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Sub
        If rsPatiOut Is Nothing Then Exit Sub
        If rsPatiOut.RecordCount = 0 Then Exit Sub
        '找到病人，将病人最新的信息更新到界面
        With rsPatiOut
            mblnEMPI = True     '用于标记找到建档病人
            '104916 只输入姓名,接口弹出界面输入更多信息找到HIS病人ID时无需再新建病人
            If mbytInState = E新增 And CLng(txt病人ID.Text) <> CLng(!病人ID & "") And CLng(!病人ID & "") <> 0 Then
                ClearCard
                txt病人ID.Text = !病人ID
                Call ReadPatiCard(CLng(txt病人ID.Text))
            End If
            Call cbo.Locate(cbo民族, !民族 & "")
            Call cbo.Locate(cbo国籍, !国籍 & "")
            Call cbo.Locate(cbo学历, !学历 & "")
            Call cbo.SeekIndex(cbo职业, !职业 & "")
            Call cbo.Locate(cbo婚姻状况, !婚姻状况 & "")
            Call cbo.Locate(cbo联系人关系, !联系人关系 & "")
            
            If mbytInState = EState.E新增 Then
                '修改时不允许EMPI直接更新病人的基本信息
                txtPatient.Text = !姓名 & ""
                Call cbo.Locate(cbo性别, !性别 & "")
                If IsDate(!出生日期 & "") Then
                    txt出生日期.Text = Format(CDate(!出生日期 & ""), "YYYY-MM-DD")
                    txt出生时间.Text = IIf(Format(CDate(!出生日期 & ""), "HH:MM") = "00:00", "__:__", Format(CDate(!出生日期 & ""), "HH:MM"))
                End If
            End If
            
            If gbln启用结构化地址 Then
                PatiAddress(E_IX_出生地点).Value = !出生地点 & ""
                PatiAddress(E_IX_现住址).Value = !家庭地址 & ""
                PatiAddress(E_IX_户口地址).Value = !户口地址 & ""
                PatiAddress(E_IX_联系人地址).Value = !联系人地址 & ""
            End If
            txtPatiMCNO(0).Text = !医保号 & ""
            txt出生地点.Text = !出生地点 & ""
            txt家庭地址.Text = !家庭地址 & ""
            txt户口地址.Text = !户口地址 & ""
            txt联系人地址.Text = !联系人地址 & ""
            txt身份证号.Text = !身份证号 & ""
            txt其他证件.Text = !其他证件 & ""
            txt工作单位.Text = !工作单位 & ""
            txt家庭电话.Text = !家庭电话 & ""
            txt联系人电话.Text = !联系人电话 & ""
            txt单位电话.Text = !单位电话 & ""
            txt家庭地址邮编.Text = !家庭地址邮编 & ""
            txt户口地址邮编.Text = !户口地址邮编 & ""
            txt单位邮编.Text = !单位邮编 & ""
            txt联系人姓名.Text = !联系人姓名 & ""
        End With
    End If
End Sub

Private Function EMPI_AddORUpdatePati(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strErr As String) As Boolean
'功能:增加或更新EMPI病人信息
    Dim lngRet  As Long
    Dim strPlugErr As String
    Dim strTmp As String
    
    lngRet = 1 '默认成功 兼容 老版zlPlug当不支持此接口错误号:438
    If CreatePlugInOK(glngModul) Then
        
        If Not mblnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "向EMPI平台新增病人信息失败！"
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "向EMPI平台更新病人信息失败！"
        End If
        If strPlugErr <> "" Then
            strErr = strTmp & vbCrLf & strPlugErr
             Exit Function
        ElseIf lngRet = 0 Then
            strErr = strTmp & vbCrLf & strErr
            Exit Function
        End If
    End If
    
    EMPI_AddORUpdatePati = True
End Function


Private Sub ReLoadCardFee(Optional ByVal blnFeedName As Boolean)
    '离开检查卡费
    Dim lng病人ID As Long, lng收费细目ID As Long
    Dim strSQL As String, str年龄 As String
    Dim rsTmp As ADODB.Recordset
    
    If mCurSendCard.rs卡费 Is Nothing Then Exit Sub
    If mCurSendCard.rs卡费.RecordCount = 0 Then Exit Sub
    If mCurSendCard.lng卡类别ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt卡号.Text) = "" Then Exit Sub
    If mbytInState = E新增 Then
        lng病人ID = mlngPatientID
    Else
        lng病人ID = mlng病人ID
    End If
    If blnFeedName = False And lng病人ID <> 0 Then Exit Sub
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    mCurSendCard.rs卡费.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 收费细目ID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "卡费", mlngModul, mCurSendCard.lng卡类别ID, Trim(txt卡号.Text), lng病人ID, _
                Trim(txtPatient.Text), NeedName(cbo性别.Text), str年龄, Trim(txt身份证号.Text), Val(Nvl(mCurSendCard.rs卡费!收费细目ID)))
    If rsTmp.EOF Then Exit Sub
    
    lng收费细目ID = Val(Nvl(rsTmp!收费细目ID))
    Set rsTmp = GetSpecialInfo("", lng收费细目ID)
    If Not rsTmp Is Nothing Then Set mCurSendCard.rs卡费 = rsTmp
    
    With mCurSendCard.rs卡费
        txt卡额.Text = Format(IIf(Val(Nvl(!是否变价)) = 1, Val(Nvl(!缺省价格)), Val(Nvl(!现价))), "0.00")
        txt卡额.Tag = txt卡额.Text  '保持不变
        txt卡额.Locked = Not (Val(Nvl(!是否变价)) = 1)
        txt卡额.TabStop = (Val(Nvl(!是否变价)) = 1)
        
        If Val(Nvl(mCurSendCard.rs卡费!屏蔽费别)) <> 1 And Val(txt卡额.Text) <> 0 Then
            txt卡额.Text = Format(GetActualMoney(NeedName(cbo费别.Text), mCurSendCard.rs卡费!收入项目ID, Val(txt卡额.Text), mCurSendCard.rs卡费!收费细目ID), "0.00")
        End If
    End With
End Sub

Private Function CheckPatiIsExist() As Boolean
'功能:检查相似病人信息(新增之前检查,以免加入了重复信息！！！)
'返回值:
'   True-继续保存;
'   False-中断保存
    Dim strSimilar As String
    Dim strType As String
    Dim strNote As String
    Dim lngPatiID As Long
    Dim rsSimilar As ADODB.Recordset
    Dim i As Long
    
    If Not CreatePublicPatient() Then Exit Function
    strType = IIf(glngSys Like "8??", "客户", "病人")
    '检查相似病人信息(新增之前检查,以免加入了重复信息！！！)
    strSimilar = SimilarIDs(zlCommFun.GetNeedName(cbo国籍.Text), zlCommFun.GetNeedName(cbo民族), CDate(IIf(IsDate(txt出生日期.Text), txt出生日期.Text, #1/1/1900#)), zlCommFun.GetNeedName(cbo性别), txtPatient.Text, txt身份证号.Text, rsSimilar)
    If strSimilar <> "" Then
        If gblnPatiByID And Trim(txt身份证号.Text) <> "" Then
            '110541 同一身份证只能对应一个建档病人;启用该参数且通过身份证号找到已建档病人时弹出选择框
            rsSimilar.Filter = "身份证号 ='" & Trim(txt身份证号.Text) & "'"
            If rsSimilar.RecordCount > 0 Then
                strNote = "在已有的病人信息中发现" & rsSimilar.RecordCount & "个身份证号相同的的病人。" & vbCrLf & vbCrLf & _
                    "提取已有的病人信息请选择病人后[双击]或点击[确定]。"
                If gobjPublicPatient.ShowSelect(rsSimilar, "ID", "病人选择", strNote, , , "0|800|1200|800|800|1500|1000", True) Then
                    lngPatiID = Val(rsSimilar!病人ID & "")
                    GoTo FindPati:
                End If
            End If
        End If
                    
        i = UBound(Split(strSimilar, "|")) + 1
        strSimilar = Replace(strSimilar, "|", vbCrLf)
        If i > 20 Then strSimilar = Mid(strSimilar, 1, 200) & "..."
        If MsgBox("在已有的" & strType & "信息中发现 " & i & " 个信息相似的" & strType & "(国籍,民族,性别,姓名,出生日期相同或身份证号相同): " & vbCrLf & vbCrLf & _
            strSimilar & vbCrLf & vbCrLf & "确实要保存该" & strType & "的信息吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        Else
            lngPatiID = -1
        End If
    End If
FindPati:
    If lngPatiID > 0 Then
        txtPatient.Text = "-" & lngPatiID
        txtPatient.SetFocus
        Call txtPatient_KeyPress(vbKeyReturn)
        Exit Function
    ElseIf lngPatiID = -1 Then
        MsgBox "该" & strType & "的相似记录可以使用""合并""功能处理！", vbInformation, gstrSysName
    End If
    
    CheckPatiIsExist = True
End Function
