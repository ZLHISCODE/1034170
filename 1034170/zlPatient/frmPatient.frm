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
   Caption         =   "���˵Ǽ�"
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
   StartUpPosition =   1  '����������
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
         Caption         =   "��"
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
         Caption         =   "������Ӧ"
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
         Caption         =   "�������"
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
         Caption         =   "������Ϣ"
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
         Caption         =   "��ϵ����Ϣ"
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
         Caption         =   "����ҽѧ��ʾ"
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
         Caption         =   "ҽѧ��ʾ"
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
         Caption         =   "Ѫ��"
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
         Caption         =   "�����￨��Ϣ��"
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   45
         TabIndex        =   129
         Top             =   7200
         Width           =   11640
         Begin VB.ComboBox cbo���㷽ʽ 
            Height          =   300
            Left            =   8925
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   370
            Width           =   2550
         End
         Begin VB.TextBox txt���� 
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
         Begin VB.TextBox txt���� 
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
         Begin VB.CheckBox chk���� 
            Caption         =   "����"
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
                  Caption         =   "�����շ�(&1)"
                  Key             =   "CardFee"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "�󶨿���(&2)"
                  Key             =   "CardBind"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl������ 
            Height          =   255
            Left            =   8925
            TabIndex        =   15
            Top             =   0
            Width           =   1590
         End
         Begin VB.Label lbl���￨�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   134
            Top             =   400
            Width           =   420
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   2250
            TabIndex        =   133
            Top             =   435
            Width           =   360
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5870
            TabIndex        =   132
            Top             =   435
            Width           =   360
         End
         Begin VB.Label lbl��֤ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��֤"
            Height          =   180
            Left            =   4065
            TabIndex        =   131
            Top             =   435
            Width           =   360
         End
         Begin VB.Label lbl���㷽ʽ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���㷽ʽ"
            Height          =   180
            Left            =   8115
            TabIndex        =   130
            Top             =   430
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "��Ԥ������Ϣ��"
         ForeColor       =   &H00C00000&
         Height          =   1230
         Left            =   45
         TabIndex        =   120
         Top             =   5880
         Width           =   11640
         Begin VB.ComboBox cboԤ������ 
            Height          =   300
            Left            =   5055
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   390
            Width           =   2550
         End
         Begin VB.TextBox txtԤ���� 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   2715
            MaxLength       =   12
            TabIndex        =   60
            Top             =   390
            Width           =   1050
         End
         Begin VB.TextBox txt������� 
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
         Begin VB.TextBox txt�ɿλ 
            Height          =   300
            Left            =   1100
            MaxLength       =   50
            TabIndex        =   63
            Top             =   780
            Width           =   2670
         End
         Begin VB.TextBox txt������ 
            Height          =   300
            Left            =   5055
            MaxLength       =   50
            TabIndex        =   64
            Top             =   780
            Width           =   2550
         End
         Begin VB.TextBox txt�ʺ� 
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
                  Caption         =   "����Ԥ��(&M)"
                  Key             =   "K1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "סԺԤ��(&Z)"
                  Key             =   "K2"
                  ImageVarType    =   2
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
         Begin VB.Label lblMoney 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
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
            Caption         =   "�������"
            Height          =   180
            Left            =   8115
            TabIndex        =   127
            Top             =   450
            Width           =   720
         End
         Begin VB.Label lblStyle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɿʽ"
            Height          =   180
            Left            =   4290
            TabIndex        =   126
            Top             =   450
            Width           =   720
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ժҪ"
            Height          =   240
            Left            =   825
            TabIndex        =   125
            Top             =   1605
            Width           =   480
         End
         Begin VB.Label lblFact 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʵ��Ʊ��"
            Height          =   180
            Left            =   315
            TabIndex        =   124
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɿλ"
            Height          =   180
            Left            =   315
            TabIndex        =   123
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   4470
            TabIndex        =   122
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʺ�"
            Height          =   180
            Left            =   8415
            TabIndex        =   121
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lblYBMoney 
            AutoSize        =   -1  'True
            Caption         =   "�����ʻ����:"
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
            Begin VB.CommandButton cmd��ϵ�˵�ַ 
               Caption         =   "��"
               Height          =   255
               Left            =   5595
               TabIndex        =   54
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   4898
               Width           =   285
            End
            Begin VB.TextBox txt��֤���� 
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
               Left            =   10815
               TabIndex        =   161
               TabStop         =   0   'False
               Top             =   1530
               Width           =   600
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
               Left            =   10170
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   1530
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
            Begin VB.CommandButton cmd���� 
               Caption         =   "��"
               Height          =   255
               Left            =   8370
               TabIndex        =   42
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   3458
               Width           =   285
            End
            Begin VB.CommandButton cmd��ͥ��ַ 
               Caption         =   "��"
               Height          =   255
               Left            =   5595
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   2745
               Width           =   285
            End
            Begin VB.CommandButton cmd���� 
               Caption         =   "��"
               Height          =   255
               Left            =   11180
               TabIndex        =   36
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   3098
               Width           =   285
            End
            Begin VB.CommandButton cmd���ڵ�ַ 
               Caption         =   "��"
               Height          =   255
               Left            =   5595
               TabIndex        =   32
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   3105
               Width           =   285
            End
            Begin VB.CommandButton cmd��ͬ��λ 
               Caption         =   "��"
               Height          =   255
               Left            =   5595
               TabIndex        =   45
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   3825
               Width           =   285
            End
            Begin VB.CommandButton cmd�����ص� 
               Caption         =   "��"
               Height          =   255
               Left            =   5595
               TabIndex        =   39
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   3450
               Width           =   285
            End
            Begin VB.TextBox txt���� 
               Height          =   300
               Left            =   9395
               MaxLength       =   30
               TabIndex        =   35
               Top             =   3075
               Width           =   2100
            End
            Begin VB.TextBox txt���� 
               Height          =   300
               Left            =   7200
               MaxLength       =   30
               TabIndex        =   41
               Top             =   3435
               Width           =   1485
            End
            Begin VB.TextBox txt��λ�ʱ� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   10085
               MaxLength       =   6
               TabIndex        =   47
               Top             =   3795
               Width           =   1410
            End
            Begin VB.TextBox txt��ϵ�˵绰 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   4080
               MaxLength       =   20
               TabIndex        =   51
               Top             =   4515
               Width           =   1815
            End
            Begin VB.TextBox txt��ϵ������ 
               Height          =   300
               Left            =   1110
               MaxLength       =   64
               TabIndex        =   50
               Top             =   4515
               Width           =   1815
            End
            Begin VB.TextBox txt��λ������ 
               Height          =   300
               Left            =   1110
               MaxLength       =   50
               TabIndex        =   48
               Top             =   4155
               Width           =   4800
            End
            Begin VB.TextBox txt��λ�绰 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   20
               TabIndex        =   46
               Top             =   3795
               Width           =   1485
            End
            Begin VB.TextBox txt��ͥ�绰 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   20
               TabIndex        =   29
               Top             =   2715
               Width           =   1485
            End
            Begin VB.TextBox txt��ͥ��ַ�ʱ� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   10085
               MaxLength       =   6
               TabIndex        =   30
               Top             =   2715
               Width           =   1410
            End
            Begin VB.TextBox txt������λ 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   44
               Top             =   3795
               Width           =   4785
            End
            Begin VB.ComboBox cbo���䵥λ 
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
            Begin VB.ComboBox cboҽ�Ƹ��� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   4080
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   2340
               Width           =   1815
            End
            Begin VB.TextBox txtסԺ�� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   5655
               MaxLength       =   18
               TabIndex        =   2
               Top             =   120
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.TextBox txt����� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3045
               MaxLength       =   18
               TabIndex        =   1
               Top             =   120
               Width           =   1545
            End
            Begin VB.TextBox txt���� 
               Height          =   300
               IMEMode         =   2  'OFF
               Left            =   3180
               TabIndex        =   9
               Top             =   855
               Width           =   800
            End
            Begin VB.ComboBox cbo��ϵ�˹�ϵ 
               Height          =   300
               Left            =   7200
               TabIndex        =   56
               Text            =   "cbo��ϵ�˹�ϵ"
               Top             =   4875
               Width           =   4295
            End
            Begin VB.ComboBox cbo����״�� 
               Height          =   300
               Left            =   5655
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1560
               Width           =   1485
            End
            Begin VB.ComboBox cboѧ�� 
               Height          =   300
               Left            =   7905
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   870
               Width           =   1485
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Left            =   5655
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1230
               Width           =   1485
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Left            =   5655
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   870
               Width           =   1485
            End
            Begin VB.ComboBox cboְҵ 
               Height          =   300
               Left            =   7905
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   1230
               Width           =   1485
            End
            Begin VB.ComboBox cbo��� 
               Height          =   300
               Left            =   7905
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   495
               Width           =   1485
            End
            Begin VB.ComboBox cbo�ѱ� 
               Height          =   300
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   2325
               Width           =   1815
            End
            Begin VB.ComboBox cbo�Ա� 
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
            Begin VB.TextBox txt����ID 
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
            Begin VB.ComboBox cbo�������� 
               Height          =   300
               Left            =   10085
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   3435
               Width           =   1410
            End
            Begin VB.TextBox txt��λ�ʺ� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   50
               TabIndex        =   49
               Top             =   4155
               Width           =   4295
            End
            Begin VB.TextBox txt��ע 
               Height          =   300
               Left            =   7200
               MaxLength       =   100
               TabIndex        =   58
               Top             =   5235
               Visible         =   0   'False
               Width           =   4295
            End
            Begin VB.CommandButton cmdYB 
               Caption         =   "��֤"
               Height          =   345
               Left            =   7155
               TabIndex        =   3
               Top             =   98
               Width           =   600
            End
            Begin VB.TextBox txt���ڵ�ַ�ʱ� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   7200
               MaxLength       =   6
               TabIndex        =   34
               Top             =   3075
               Width           =   1485
            End
            Begin VB.TextBox txt��ϵ�����֤ 
               Height          =   300
               Left            =   7200
               MaxLength       =   18
               TabIndex        =   52
               Top             =   4515
               Width           =   4295
            End
            Begin VB.TextBox txt֧������ 
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
               ToolTipText     =   "��ݼ�F4"
               Top             =   495
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   529
               Appearance      =   2
               IDKindStr       =   $"frmPatient.frx":0E42
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   9
               FontName        =   "����"
               IDKind          =   -1
               DefaultCardType =   "0"
               BackColor       =   -2147483633
            End
            Begin MSMask.MaskEdBox txt����ʱ�� 
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
            Begin MSMask.MaskEdBox txt�������� 
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
            Begin VB.TextBox txt���֤�� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   1110
               TabIndex        =   13
               Top             =   1215
               Width           =   2160
            End
            Begin VB.TextBox txt����֤�� 
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
               Tag             =   "����"
               Top             =   3075
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
            Begin VB.TextBox txt���ڵ�ַ 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   31
               Top             =   3075
               Width           =   4785
            End
            Begin VB.TextBox txt�����ص� 
               Height          =   300
               Left            =   1110
               MaxLength       =   100
               TabIndex        =   38
               Top             =   3435
               Width           =   4785
            End
            Begin VB.TextBox txt��ͥ��ַ 
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
               Tag             =   "�����ص�"
               Top             =   3435
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
            Begin VB.TextBox txt��ϵ�˵�ַ 
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
               Tag             =   "��סַ"
               Top             =   2715
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
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
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   300
               Index           =   4
               Left            =   1110
               TabIndex        =   33
               Tag             =   "���ڵ�ַ"
               Top             =   3075
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
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
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   300
               Index           =   5
               Left            =   1110
               TabIndex        =   55
               Tag             =   "��ϵ�˵�ַ"
               Top             =   4875
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   529
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
            Begin VB.Label lblMobile 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ֻ���"
               Height          =   180
               Left            =   540
               TabIndex        =   169
               Top             =   5295
               Width           =   540
            End
            Begin VB.Label lbl��֤���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��֤����"
               Height          =   180
               Left            =   9270
               TabIndex        =   167
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lbl֧������ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "֧������"
               Height          =   180
               Left            =   6420
               TabIndex        =   166
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lblѧ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѧ��"
               Height          =   180
               Left            =   7500
               TabIndex        =   165
               Top             =   930
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ѷ�ʽ"
               Height          =   180
               Left            =   3300
               TabIndex        =   164
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lblPatiMCNO 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��֤ҽ����"
               Height          =   180
               Index           =   1
               Left            =   6240
               TabIndex        =   163
               Top             =   2010
               Width           =   900
            End
            Begin VB.Label lblְҵ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ְҵ"
               Height          =   180
               Left            =   7500
               TabIndex        =   162
               Top             =   1290
               Width           =   360
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Left            =   6780
               TabIndex        =   119
               Top             =   3495
               Width           =   360
            End
            Begin VB.Label lblPatiMCNO 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ����"
               Height          =   180
               Index           =   0
               Left            =   540
               TabIndex        =   118
               Top             =   2010
               Width           =   540
            End
            Begin VB.Label lblסԺ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ��"
               Height          =   180
               Left            =   5070
               TabIndex        =   117
               Top             =   180
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lbl����� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����"
               Height          =   180
               Left            =   2490
               TabIndex        =   116
               Top             =   180
               Width           =   540
            End
            Begin VB.Label lbl��λ�ʺ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ�ʺ�"
               Height          =   180
               Left            =   6420
               TabIndex        =   115
               Top             =   4215
               Width           =   720
            End
            Begin VB.Label lbl��λ������ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ������"
               Height          =   180
               Left            =   180
               TabIndex        =   114
               Top             =   4215
               Width           =   900
            End
            Begin VB.Label lbl��λ�ʱ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ�ʱ�"
               Height          =   180
               Left            =   9270
               TabIndex        =   113
               Top             =   3855
               Width           =   720
            End
            Begin VB.Label lbl��λ�绰 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ�绰"
               Height          =   180
               Left            =   6420
               TabIndex        =   112
               Top             =   3855
               Width           =   720
            End
            Begin VB.Label lbl������λ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������λ"
               Height          =   180
               Left            =   360
               TabIndex        =   111
               Top             =   3855
               Width           =   720
            End
            Begin VB.Label lbl��ϵ�˵绰 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ�˵绰"
               Height          =   180
               Left            =   3120
               TabIndex        =   110
               Top             =   4575
               Width           =   900
            End
            Begin VB.Label lbl��ϵ�˵�ַ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ�˵�ַ"
               Height          =   180
               Left            =   180
               TabIndex        =   109
               Top             =   4935
               Width           =   900
            End
            Begin VB.Label lbl��ϵ�˹�ϵ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ�˹�ϵ"
               Height          =   180
               Left            =   6240
               TabIndex        =   108
               Top             =   4935
               Width           =   900
            End
            Begin VB.Label lbl��ϵ������ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ������"
               Height          =   180
               Left            =   180
               TabIndex        =   107
               Top             =   4575
               Width           =   900
            End
            Begin VB.Label lbl��ͥ��ַ�ʱ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ͥ��ַ�ʱ�"
               Height          =   180
               Left            =   8910
               TabIndex        =   106
               Top             =   2775
               Width           =   1080
            End
            Begin VB.Label lbl��ͥ�绰 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ͥ�绰"
               Height          =   180
               Left            =   6420
               TabIndex        =   105
               Top             =   2775
               Width           =   720
            End
            Begin VB.Label lbl��ͥ��ַ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��סַ"
               Height          =   180
               Left            =   540
               TabIndex        =   104
               Top             =   2775
               Width           =   540
            End
            Begin VB.Label lbl����״�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Left            =   5250
               TabIndex        =   103
               Top             =   1620
               Width           =   360
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Left            =   5250
               TabIndex        =   102
               Top             =   1290
               Width           =   360
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Left            =   5250
               TabIndex        =   101
               Top             =   930
               Width           =   360
            End
            Begin VB.Label lbl��� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���"
               Height          =   180
               Left            =   7500
               TabIndex        =   100
               Top             =   555
               Width           =   360
            End
            Begin VB.Label lbl���֤�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���֤��"
               Height          =   180
               Left            =   360
               TabIndex        =   99
               Top             =   1275
               Width           =   720
            End
            Begin VB.Label lbl�����ص� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����ص�"
               Height          =   180
               Left            =   360
               TabIndex        =   98
               Top             =   3495
               Width           =   720
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               Height          =   180
               Left            =   360
               TabIndex        =   97
               Top             =   915
               Width           =   720
            End
            Begin VB.Label lbl�ѱ� 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ѱ�"
               Height          =   180
               Left            =   720
               TabIndex        =   96
               Top             =   2385
               Width           =   360
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Left            =   2790
               TabIndex        =   95
               Top             =   915
               Width           =   360
            End
            Begin VB.Label lbl�Ա� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ա�"
               Height          =   180
               Left            =   5250
               TabIndex        =   94
               Top             =   555
               Width           =   360
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Left            =   30
               TabIndex        =   93
               Top             =   555
               Width           =   360
            End
            Begin VB.Label lbl����ID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ID"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   540
               TabIndex        =   92
               Top             =   180
               Width           =   540
            End
            Begin VB.Label lbl����֤�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����֤��"
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
               Caption         =   "��������"
               Height          =   180
               Left            =   9270
               TabIndex        =   88
               Top             =   3495
               Width           =   720
            End
            Begin VB.Label lbl��ע 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ע"
               Height          =   180
               Left            =   6780
               TabIndex        =   85
               Top             =   5295
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Left            =   8910
               TabIndex        =   168
               Top             =   3135
               Width           =   360
            End
            Begin VB.Label lbl���ڵ�ַ�ʱ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ڵ�ַ�ʱ�"
               Height          =   180
               Left            =   6060
               TabIndex        =   83
               Top             =   3135
               Width           =   1080
            End
            Begin VB.Label lbl���ڵ�ַ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ڵ�ַ"
               Height          =   180
               Left            =   360
               TabIndex        =   82
               Top             =   3135
               Width           =   720
            End
            Begin VB.Label lbl��ϵ�����֤ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ�����֤"
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
      Caption         =   "ҽ�ƿ�(&2)"
      Height          =   350
      Index           =   1
      Left            =   2760
      TabIndex        =   76
      ToolTipText     =   "�������￨"
      Top             =   8460
      Width           =   1100
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "Ԥ����(&1)"
      Height          =   350
      Index           =   0
      Left            =   1440
      TabIndex        =   74
      ToolTipText     =   "��������Ԥ����"
      Top             =   8460
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   135
      TabIndex        =   75
      Top             =   8475
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10455
      TabIndex        =   73
      Top             =   8460
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
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

Option Explicit 'Ҫ���������
Public mlngModul As String
Public mstrPrivs As String
Public mbytInState As Byte '�룺0=����,1=�޸�,2=�鿴
Public mbytView As Byte '�룺0-����,1-��Ժ,2-��Ժ,3-����
Public mlng����ID As Long 'Ҫ�޸Ļ�鿴�Ĳ���ID
Public mlng��ҳID As Long
Private mlngԤ������ID As Long 'Ԥ����Ʊ������ID
Private mlng���� As Long
Private mlngOutModeMC As Long '���ʽҽ��������
Private mblnUnLoad As Boolean
Private mblnICCard As Boolean 'IC������,Ҫͬʱ��д������Ϣ��IC���ֶ�
Private mblnChange As Boolean
Private mblnSel As Boolean
Private mblnCheckPatiCard As Boolean
Private mstrYBPati As String
Private mblnPrepayPrint As Boolean    '�Ƿ��ӡԤ����
Private mstr�ɼ�ͼƬ As String '�ɼ�ͼƬ���ر���·��
Private mlngͼ����� As Long 'ָ����ǰ�Բ���ͼ�����������(1-�ļ� 2-�ɼ� 3-���)
Private mobjPublicPatient As Object
Private mstrPatiPlus    As String     '�ӱ���Ϣ:��Ϣ��1:��Ϣֵ1,��Ϣ��2:��Ϣֵ2
Private mblnEMPI As Boolean             'T-�ҵ�EMPI���ˣ�F-δ�ҵ�EMPI����
Private Enum OPT
    C0Ԥ���� = 0
    C1���￨ = 1
End Enum
Private mlngPatientID As Long '����ʱ��ȡ�������ʱ����
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'���ڽ��㿨�ĵĴ������
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '��װ�˽��㿨�ĵ�
    dblˢ���ܶ� As Double
    bln������ As Boolean '��ǰ��ȡ�ĵ����ǿ�����
End Type

Private mtySquareCard As Ty_SquareCard
Private mobjKeyboard As Object

'Private mobjSquareCard As Object
Private mblnClickSquareCtrl As Boolean
Private mFactProperty As Ty_FactProperty
Private mblnStartFactUseType As Boolean '�Ƿ����õ���ص���������
Private mbytPrepayType As Byte '0-����סԺ;1-����;2-סԺ
Private mblnNotClick As Boolean
Private mobjSquare As Object 'ҽ�ƿ�����
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
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
       int���볤�� As Integer
       int���볤������ As Integer
       int������� As Integer
       bln���￨ As Boolean
       str�������� As String
       blnOneCard As Boolean '  '�Ƿ�������һ��ͨ�ӿ�,��ģʽ�£�Ʊ���ϸ����Ʊ�ŷ�Χ��ķ�����󶨿����շ�
       rs���� As ADODB.Recordset
       dblӦ�ս�� As Double
       dblʵ�ս�� As Double
       bln�Ƿ��ƿ� As Boolean
       bln�Ƿ񷢿� As Boolean
       bln�Ƿ�д�� As Boolean
       bln�Ƿ�Ժ�ⷢ��  As Boolean
       lng�������� As Long '0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0 Ϊ���:57326
       bln�ظ�ʹ�� As Boolean
       byt�������� As Byte
End Type
Private mCurSendCard As Ty_CardProperty
Private mcolPrepayPayMode As Collection   'Ԥ����֧����ʽ
Private mcolCardPayMode As Collection   '���￨֧����ʽ
Private Type Ty_PayMoney
    lngҽ�ƿ����ID As Long
    bln���ѿ� As Boolean
    str���㷽ʽ As String
    str���� As String
    strˢ������ As String
    strˢ������ As String
    strNO As String
    lngID As Long 'Ԥ��ID
    lng����ID As Long
End Type
Private mCurPrepay As Ty_PayMoney
Private mCurCardPay As Ty_PayMoney
Private mbln�Ƿ�ɨ�����֤ As Boolean '�Ƿ���ִ�е�ɨ�����֤����
Private mblnɨ�����֤ǩԼ As Boolean '���ݲ��������еġ�ɨ�����֤ǩԼ����ȡֵ
Private marrAddress(0 To 4) As String  '�ṹ����ַȱʡֵ
'����� :56599
Private Type ty_PageHeight
    ���� As Long
    �������� As Long
    ������Ϣ As Long
End Type
Private mPageHeight As ty_PageHeight

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Enum EState
    E���� = 0
    E�޸� = 1
    E���� = 2
End Enum

Private mstrCboSplit As String
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const C_ColumHeader = "����ҩ��,1,3000,1;������ӳ,4,3000,1;����ҩ��ID,1,100,0" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_InoculateHeader = "��������,4,2100,1;��������,4,2100,1;��������,4,2100,1;��������,4,2100,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_LinkManColumHeader = "��ϵ������,4,2100,1;��ϵ�˹�ϵ,4,2100,1;��ϵ�����֤��,4,2100,1;��ϵ�˵绰,4,2100,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_OtherInfoColumHeader = "��Ϣ��,4,2288,1;��Ϣֵ,4,2288,1;��Ϣ��,4,2287,1;��Ϣֵ,4,2287,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
'Private Const C_Ѫ�� = "A��,B��,O��,AB��,����"
Private Const C_BH = "��,��,����,δ��"
Private mdicҽ�ƿ����� As New Dictionary
Private mobjHealthCard As Object '�ƿ��ӿڶ���
Private mbln������󶨿� As Boolean '��ʶ�Ƿ�����˷�����󶨿�����
Private mbln����  As Boolean '��ʶ��ǰѡ��ҳ
Private mlngPlugInHwnd As Long
'Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:������رս��㿨����
'    '���:blnClosed:�رն���
'    '����:���˺�
'    '����:2010-01-05 14:51:23
'    '����:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strExpend As String
'    '0=����,1=�޸�,2=�鿴
'   If mbytInState = E���� Then Exit Sub
'
'    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
'    If blnClosed Then
'       If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.CloseWindows

'            Set mobjSquareCard = Nothing
'        End If
'        Exit Sub
'    End If
'
'    '��������
'    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
'    Err = 0: On Error Resume Next
'    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
'    If Err <> 0 Then
'        Err = 0: On Error GoTo 0:      Exit Sub
'    End If
'
'    '��װ�˽��㿨�Ĳ���
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    '����:zlInitComponents (��ʼ���ӿڲ���)
'    '    ByVal frmMain As Object, _
'    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
'    '        ByVal cnOracle As ADODB.Connection, _
'    '        Optional blnDeviceSet As Boolean = False, _
'    '        Optional strExpand As String
'    '����:
'    '����:   True:���óɹ�,False:����ʧ��
'    '����:���˺�
'    '����:2009-12-15 15:16:22
'    'HIS����˵��.
'    '   1.���������շ�ʱ���ñ��ӿ�
'    '   2.����סԺ����ʱ���ñ��ӿ�
'    '   3.����Ԥ����ʱ
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    If mobjSquareCard.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
'         '��ʼ�������ɹ�,����Ϊ�����ڴ���
'         Exit Sub
'    End If
'    '��ʼ�ɹ�,��֤���˴��ڴ�����صĽ��㿨
'     mtySquareCard.blnExistsObjects = True
'End Sub


Private Sub InitSendCardPreperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ˢ������
    '����:���˺�
    '����:2011-07-25 11:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strSQL As String, blnBoundCard As Boolean
    Dim rsTemp As ADODB.Recordset, str���� As String, varData As Variant, i As Long
    Dim varTemp  As Variant, blnNotBind As Boolean
    '76824�����ϴ���2014/8/19��ҽ�ƿ������
    lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, 0))
    If InStr(mstrPrivs, ";��������;") = 0 Or lngCardTypeID = 0 Then '�޷���Ȩ��
NotCard:
        fraCard.Visible = False: cmdOperation(OPT.C1���￨).Visible = False
        Me.Height = Me.Height - fraCard.Height
        mPageHeight.���� = Me.Height
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    '�����:57326
    strSQL = "" & _
    "   Select Id, ����, ����, ����, ǰ׺�ı�, ���ų���, ȱʡ��־, �Ƿ�̶�, �Ƿ��ϸ����,nvl( �Ƿ�ˢ��,0) as �Ƿ�ˢ��, " & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����, nvl(�Ƿ�����ʻ�,0) as �Ƿ�����ʻ�, " & _
    "           nvl(�Ƿ�ȫ��,0) as �Ƿ�ȫ��,nvl(�Ƿ��ظ�ʹ��,0) as �Ƿ��ظ�ʹ�� , " & _
    "           nvl(���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����,����, ��ע, �ض���Ŀ, ���㷽ʽ, �Ƿ�����, ��������," & _
    "           nvl(�Ƿ��ƿ�,0) as �Ƿ��ƿ�,nvl(�Ƿ񷢿�,0) as �Ƿ񷢿�, nvl(�Ƿ�д��,0) as �Ƿ�д��, " & _
    "           nvl(��������,0) as ��������,nvl(��������,0) as �������� " & _
    "    From ҽ�ƿ���� A" & _
    "    Where nvl(�Ƿ�����,0)=1 And (ID=[1] or nvl(ȱʡ��־,0)=1)" & _
    "    Order by ����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardTypeID)
    If rsTemp.EOF Then GoTo NotCard:
    If rsTemp.RecordCount >= 2 Then
        rsTemp.Filter = "ID=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0
    End If
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        With mCurSendCard
            .lng�����ID = Val(Nvl(rsTemp!ID))
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
            .bln���￨ = .str������ = "���￨" And Val(Nvl(rsTemp!�Ƿ�̶�)) = 1
            '�����:56599
            .bln�Ƿ��ƿ� = Val(Nvl(rsTemp!�Ƿ��ƿ�)) = 1
            .bln�Ƿ񷢿� = Val(Nvl(rsTemp!�Ƿ񷢿�)) = 1
            .bln�Ƿ�д�� = Val(Nvl(rsTemp!�Ƿ�д��)) = 1
            .bln�ظ�ʹ�� = Val(Nvl(rsTemp!�Ƿ��ظ�ʹ��)) = 1
            .bln�Ƿ�Ժ�ⷢ�� = (InStr(mstrPrivs, ";��������;") > 0 And .bln���ƿ� = False And .bln�Ƿ񷢿� = True) '�����:56599
            .lng�������� = Val(Nvl(rsTemp!��������)) '�����:57326
            .byt�������� = Val(Nvl(rsTemp!��������))
            '76824�����ϴ���2014/8/19��ҽ�ƿ������
            lbl������.Caption = .str������
            lbl������.Width = LenB(lbl������.Caption) * 100
            .blnOneCard = False
            If Trim(Nvl(rsTemp!�ض���Ŀ)) <> "" Then
                Set .rs���� = GetSpecialInfo(Trim(Nvl(rsTemp!�ض���Ŀ)))
                If .bln���￨ Then .blnOneCard = GetOneCard.RecordCount > 0
                '142204:���ϴ�,2020/6/16,�޸ķ�����Ӧ�ս��
                If Not .rs����.EOF Then
                    .bln��� = Val(Nvl(.rs����!�Ƿ���)) = 1
                    .dblӦ�ս�� = Format(IIf(.bln���, .rs����!ȱʡ�۸�, .rs����!�ּ�), "0.00")
                End If
            Else
                Set .rs���� = Nothing
            End If
            str���� = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModul, "0")
            '����ID,�����ID|...
             .lng�������� = 0
            varData = Split(str����, "|")
            For i = 0 To UBound(varData)
                 varTemp = Split(varData(i), ",")
                 If Val(varTemp(0)) <> 0 Then
                    If ExistShareBill(Val(varTemp(0)), 5) Then
                        If Val(varTemp(1)) = .lng�����ID Then
                            .lng�������� = Val(varTemp(0)): Exit For
                        End If
                    End If
                 End If
            Next
           txt����.PasswordChar = IIf(.str�������� <> "", "*", "")
           txt����.MaxLength = .lng���ų���
        End With
    End If

    If mCurSendCard.rs���� Is Nothing Then
    
        cmdOperation(OPT.C1���￨).Visible = False
        tabCardMode.Tabs.Remove ("CardFee")
        blnBoundCard = InStr(mstrPrivs, ";�󶨿���;") > 0
        '�ް󶨿�Ȩ��
          fraCard.Visible = blnBoundCard: cmdOperation(OPT.C1���￨).Visible = blnBoundCard
        If Not blnBoundCard Then
            Me.Height = Me.Height - fraCard.Height
            mPageHeight.���� = Me.Height
        Else
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "�󶨿���"
            tabCardMode.Width = tabCardMode.Width / 2
        End If
        Exit Sub
    End If
     
    
    With mCurSendCard.rs����
        txt����.Text = Format(IIf(Val(Nvl(!�Ƿ���)) = 1, Val(Nvl(!ȱʡ�۸�)), Val(Nvl(!�ּ�))), "0.00")
        txt����.Tag = txt����.Text  '���ֲ���
        txt����.Locked = Not (Val(Nvl(!�Ƿ���)) = 1)
        txt����.TabStop = (Val(Nvl(!�Ƿ���)) = 1)
    End With
     
     
    '���ƿ�,�ڿ��Ų��ظ�ʹ�� �����ϸ����ʱ,���ܽ��а󶨿�����
    blnNotBind = mCurSendCard.bln���ƿ� And (Not mCurSendCard.bln�ظ�ʹ�� Or mCurSendCard.bln�ϸ����)
    
    '���û�а󶨿�Ȩ��,���ش���ʱ,�Ѿ��Ƴ��˰󶨿���
    blnBoundCard = Not InStr(mstrPrivs, ";�󶨿���;") > 0
    If Not blnBoundCard Then
        If zlDatabase.GetPara("����ģʽ", glngSys, mlngModul, "CardFee") = "CardFee" Then
            tabCardMode.Tabs("CardFee").Selected = True
        ElseIf Not blnNotBind Then
            tabCardMode.Tabs("CardBind").Selected = True
        End If
    End If
    
    '�����:56599
    If (mCurSendCard.bln�Ƿ�Ժ�ⷢ�� Or blnNotBind) And Not blnBoundCard Then
       '1.���Ժ�⿨���з��� 2.Ժ�ڿ�.�ϸ���ƻ��߲��ظ�����   ������2���������ͬʱӵ�а󶨿�Ȩ�� �����ܽ��а󶨿�����,�ް󶨿�Ȩ��,�ڴ������ʱ,��ɾ���˰󶨿�
        tabCardMode.Tabs.Remove ("CardBind")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardFee").Selected = True
            tabCardMode.Tabs("CardFee").Caption = "�շѷ���"
            tabCardMode.Width = tabCardMode.Width / 2
        Else
            fraCard.Visible = False
            Me.Height = Me.Height - fraCard.Height
            mPageHeight.���� = Me.Height
        End If
    ElseIf mCurSendCard.bln���ƿ� = False And mCurSendCard.bln�Ƿ񷢿� = False Then
        tabCardMode.Tabs.Remove ("CardFee")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "�󶨿���"
            tabCardMode.Width = tabCardMode.Width / 2
        Else
            fraCard.Visible = False
            Me.Height = Me.Height - fraCard.Height
            mPageHeight.���� = Me.Height
        End If
    End If
    
    If mCurSendCard.bln�ϸ���� Then
        '���￨���ü��
        mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), , mCurSendCard.lng�����ID)
        If mCurSendCard.lng����ID <= 0 Then
            Select Case mCurSendCard.lng����ID
                Case 0 '����ʧ��
                Case -1
'                    MsgBox "��û�����û��õľ��￨,���ܷ��ţ�" & vbCrLf & _
'                        "�����ڱ������ù������λ�����һ���¿�! ", vbExclamation, gstrSysName
                Case -2
'                    MsgBox "���ع��õľ��￨������,���ܷ��ţ�" & vbCrLf & _
'                        "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
            End Select
            cmdOperation(OPT.C1���￨).Visible = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo��������_Click()
    If cbo��������.ListCount > 0 And cbo��������.ListIndex <> -1 Then
        lblPatiColor.BackColor = zlDatabase.GetPatiColor(NeedName(cbo��������.Text))
        txtPatient.ForeColor = lblPatiColor.BackColor
    End If
End Sub
Private Sub cbo���㷽ʽ_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    With mCurCardPay
            .lngҽ�ƿ����ID = 0
            .bln���ѿ� = False
            .str���㷽ʽ = ""
            .str���� = ""
     End With
    '0=����,1=�޸�,2=�鿴
    If mbytInState = E���� Then Exit Sub
    Call SetCardVaribles(False)
    '130245,���㷽ʽ�л���ͬ�����¿����ID
    If mblnNotClick = True Then Exit Sub
    Call Local���㷽ʽ(mCurCardPay.lngҽ�ƿ����ID, True)
End Sub

Private Sub cbo��ϵ�˹�ϵ_Click()
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ�˹�ϵ") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ�˹�ϵ")) = NeedName(cbo��ϵ�˹�ϵ.Text)
    End If
End Sub

Private Sub cbo�Ա�_Change()
    Call ReLoadCardFee
End Sub

Private Sub cboԤ������_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    
    With mCurPrepay
            .lngҽ�ƿ����ID = 0
            .bln���ѿ� = False
            .str���㷽ʽ = ""
            .str���� = ""
     End With
    '0=����,1=�޸�,2=�鿴
    If mbytInState = E���� Then Exit Sub
    Call SetCardVaribles(True)
    '130245,���㷽ʽ�л���ͬ�����¿����ID
    If mblnNotClick = True Then Exit Sub
    Call Local���㷽ʽ(mCurPrepay.lngҽ�ƿ����ID, False)
End Sub

Private Sub cmdPicClear_Click()
    '�����:74421
    imgPatient.Picture = Nothing
    mlngͼ����� = 3
End Sub

Private Sub cmdPicCollect_Click()
    If mobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set mobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPublicPatient Is Nothing Then
        MsgBox "����������Ϣ��������(zlPublicPatient.clsPublicPatient)ʧ��!", vbInformation, gstrSysName
        Exit Sub
    End If
    If mobjPublicPatient.PatiImageGatherer(Me, mstr�ɼ�ͼƬ) = False Then Exit Sub
    Set imgPatient.Picture = LoadPicture(mstr�ɼ�ͼƬ)
    mlngͼ����� = 2
End Sub

Private Sub cmdPicFile_Click()
    '�����:74421
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
    mlngͼ����� = 1
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdYB_Click()
    Dim lng����ID As Long, lng����ID As Long
    Dim objCurrent As Object, strTxt As String, arrTxt As Variant
    Dim i As Long, blnDo As Boolean, arrPati As Variant
    Dim objcbo As ComboBox
    Dim strYBPati As String, strYBPatiBak As String
    Dim intInsure As Integer
    
    'ҽ���Ķ�
    lng����ID = mlngPatientID
    strYBPati = gclsInsure.Identify(1, lng����ID, intInsure, 1)
    mstrYBPati = strYBPati
    If strYBPati <> "" Then
        arrPati = Split(strYBPati, ";")
        '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,...
        If UBound(arrPati) >= 8 Then
            If Val(arrPati(8)) > 0 Then
               txtPatient.Text = "-" & Val(arrPati(8))
                blnDo = txtPatient.Locked
                txtPatient.Locked = False
                Call txtPatient_KeyPress(13)
                txtPatient.Locked = blnDo
                If strYBPati = "" Then txtPatient.SetFocus: Exit Sub  '������Ϊ��������ѡ�����˳���,������clearcard
            End If
        End If
        
        
        'ҽ����
        txtPatiMCNO(0).Text = arrPati(1)
        txtPatiMCNO(0).Locked = True
        
        '����
        txtPatient.Text = arrPati(3)
        
        '�Ա�
        cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, CStr(arrPati(4)))
        
        '��������
        If IsDate(arrPati(5)) Then
            txt��������.Text = Format(arrPati(5), "yyyy-MM-dd")
            Call txt��������_LostFocus
        End If
        
        '���֤��
        txt���֤��.Text = arrPati(6)
        
        '������λ
        txt������λ.Text = arrPati(7)
        
        If txt�����.Text = "" Then txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
        
        If cbo����.ListIndex = -1 Then Call ReadDict("����", cbo����)
        If cbo����.ListIndex = -1 Then Call ReadDict("����", cbo����)
        If cboѧ��.ListIndex = -1 Then Call ReadDict("ѧ��", cboѧ��)
        If cbo����״��.ListIndex = -1 Then Call ReadDict("����״��", cbo����״��)
        If cboְҵ.ListIndex = -1 Then Call ReadDict("ְҵ", cboְҵ)
        If cbo���.ListIndex = -1 Then Call ReadDict("���", cbo���)
        
        '����ʱ�������Ͳ��ɼ�
        'lblPatiType.Visible = False: cbo��������.Visible = False: lblPatiColor.Visible = False
       
        If Not IsDate(txt��������.Text) Then
            txt��������.SetFocus
        Else
            strTxt = "txt����,cbo�Ա�,cbo�ѱ�,cbo����,cbo����,cboѧ��,cbo����״��,cboְҵ,cbo���," & _
                     "txt���֤��,txt�����ص�,txt��ͥ��ַ,txt��ͥ��ַ�ʱ�,txt��ͥ�绰,txt������λ,txt��λ�绰,txt��λ�ʱ�," & _
                     "txt��λ������,txt��λ�ʺ�,txt��ϵ������,cbo��ϵ�˹�ϵ,txt��ϵ�˵�ַ,txt��ϵ�˵绰,txt��ϵ�����֤"
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

Private Sub cmd���ڵ�ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
            " Substr(����,1,2) as ���� From ����" & _
            " Union All" & _
            " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
            " From ���� Order by ����", 2, "����", , txt�����ص�.Text)
    If Not rsTmp Is Nothing Then
        txt���ڵ�ַ.Text = rsTmp!����
        txt���ڵ�ַ.SelStart = Len(txt���ڵ�ַ.Text)
        txt���ڵ�ַ.SetFocus
    End If
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        SelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        SelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXml As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    '�����:56599
    If strOutPatiInforXml <> "" Then LoadPati strOutPatiInforXml
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    
    Set gobjSquare.objCurCard = objCard
    '�Ƿ�������ʾ
    'txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55571:������,2012-011-12
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then
        txtPatient.Text = ""
        '69200:������,2013-12-31,������ȡ���в���,�л����뷽ʽ��ʾҪ��ʼ¼���²��ˡ�
        If mbytInState = E���� And mlngPatientID <> 0 Then
            Call ClearCard
            mblnICCard = False
            txt����ID.Text = zlDatabase.GetNextNo(1): lbl����ID.Tag = txt����ID.Text
            txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
        End If
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub lbl���￨��_Click()
    Dim strExpand As String, strOutCardNO As String, strOutPatiInforXml As String

    If mCurSendCard.bln���￨ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        
        If Not mobjICCard Is Nothing Then
            txt����.Text = mobjICCard.Read_Card()
            If txt����.Text <> "" Then
                mblnICCard = True
                Call CheckFreeCard(txt����.Text)
            End If
        End If
        Exit Sub
    End If
    If mCurSendCard.blnˢ�� Or mCurSendCard.lng�����ID = 0 Then Exit Sub
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

    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, mCurSendCard.lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txt����.Text = strOutCardNO
    If txt����.Text <> "" Then
        '�����:56599
        If strOutPatiInforXml <> "" Then Call LoadPati(strOutPatiInforXml)
        Call CheckFreeCard(txt����.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt����.Text = strCardNo
    If txt����.Text <> "" Then
        '�����:56599
        If strXmlCardInfor <> "" Then Call LoadPati(strXmlCardInfor)
        Call CheckFreeCard(txt����.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim objCard As Card
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
        If objCard Is Nothing Then Exit Sub
        txtPatient.Text = strCardNo
        Call FindPati(objCard, True, strCardNo)
        
        If txtPatient.Text <> "" Then
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
        End If
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    
    Dim lngIndex As Long, lngPatientID As Long
    Dim objCard As Card
    Dim blnǩԼ As Boolean
    Dim strErrMsg As String
    Dim str���� As String
    '57945:������,2013-10-30,��ȡ���֤�еĵ�ַӦ�÷ŵ����ڵ�ַ�����Ǽ�ͥ��ַ
    '55218:������,2012-10-25
'    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
'
'        txtPatient.Text = strName
'        Call zlControl.CboLocate(cbo�Ա�, strSex)
'        Call zlControl.CboLocate(cbo����, strNation)
'        txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
'        txt����ʱ��.Text = "00:00"
'        txt���ڵ�ַ.Text = strAddress
'        txt���֤��.Text = strID
'    End If
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
        If objCard Is Nothing Then Exit Sub
        txtPatient.Text = strID
        Call FindPati(objCard, False, strID, lngPatientID)
        'ˢ���֤ʱУ������
        If strName <> Trim(txtPatient.Text) And strName <> "" And mlngPatientID <> 0 Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIn(mlngPatientID) Then
                    '���ھ���
                    MsgBox "����ԭ����������" & Trim(txtPatient.Text) & "�������֤�����������" & strName & "����һ�£������޸ĳ���ȷ���������ٽ��еǼǡ�", vbInformation, gstrSysName
                    Call ClearCard:  Exit Sub
                Else
                    MsgBox "����ԭ����������" & Trim(txtPatient.Text) & "�������֤�����������" & strName & "����һ�£�ϵͳ������Ϊ���֤�����������", vbInformation, gstrSysName
                    str���� = Trim(txt����.Text)
                    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
                    If gobjPublicPatient.SavePatiBaseInfo(mlngPatientID, 0, strName, strSex, str����, Format(datBirthDay, "yyyy-MM-dd"), Me.Caption, 1, "", True, True) Then
                        txtPatient.Text = "-" & mlngPatientID
                        txtPatient.SetFocus
                        Call txtPatient_KeyPress(vbKeyReturn)
                    End If
                End If
            End If
        End If
        mbln�Ƿ�ɨ�����֤ = False
        If (mCurSendCard.str������ = "�������֤" Or mblnɨ�����֤ǩԼ) Then blnǩԼ = �Ƿ��Ѿ�ǩԼ(Trim(strID))
        If lngPatientID <> 0 And Not blnǩԼ And (mCurSendCard.str������ = "�������֤" Or mblnɨ�����֤ǩԼ) Then
            '���в��ˣ����֤ûǩԼ,������֤��Ϣ��Ϣ�Ƿ�����֤��Ƭ�ϵ���Ϣһ�� 2012-10-26 lgf
            If Trim(txtPatient.Text) <> Trim(strName) Or NeedName(cbo�Ա�.Text) <> strSex Or Format(txt��������.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                      If Trim(txtPatient.Text) <> Trim(strName) Then
                           strErrMsg = strErrMsg & "," & "����"
                      End If
                      If NeedName(Me.cbo�Ա�.Text) <> strSex Then
                           strErrMsg = strErrMsg & "," & "�Ա�"
                      End If
                      If Format(txt��������.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                          strErrMsg = strErrMsg & "," & "��������"
                      End If
                      strErrMsg = Mid(strErrMsg, 2)
                      strErrMsg = "��ǰ������Ϣ�����֤�ϵ�[" & strErrMsg & "]����Ϣ��һ��!" & vbCrLf & "���ܽ������֤ǩԼ����!"
                      Call MsgBox(strErrMsg, vbQuestion, Me.Caption)
                       mbln�Ƿ�ɨ�����֤ = False
            Else
                 mbln�Ƿ�ɨ�����֤ = True
            End If
        End If
        
        If lngPatientID = 0 Then '�²���
            lngIndex = IDKind.GetKindIndex("����")
            If lngIndex >= 0 Then IDKind.IDKind = lngIndex
            txtPatient.Text = "": txtPatient.PasswordChar = ""
            '55571:������,2012-011-12
            txtPatient.IMEMode = 0
            txtPatient.Text = strName
            Call zlControl.CboLocate(cbo�Ա�, strSex)
            Call zlControl.CboLocate(cbo����, strNation)
            txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt����ʱ��.Text = "00:00"
            txt���֤��.Text = strID
            '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
            Call LoadIDImage
            mbln�Ƿ�ɨ�����֤ = Not blnǩԼ
        End If
        '101692�²���ֱ����ȡ���Ѿ��������˵����ڵ�ַΪ��ʱ�Զ�����
        If lngPatientID = 0 Or (lngPatientID <> 0 And Trim(txt���ڵ�ַ.Text) = "") Then
            txt���ڵ�ַ.Text = strAddress
            If gbln���ýṹ����ַ Then
                PatiAddress(E_IX_���ڵ�ַ).Value = strAddress
            End If
        End If
        mblnNotClick = False
    End If
'   55240 2012-10-26 lgf
'    '�����:53408
'    mbln�Ƿ�ɨ�����֤ = False
'    If mblnɨ�����֤ǩԼ Then
'         mbln�Ƿ�ɨ�����֤ = Not �Ƿ��Ѿ�ǩԼ(strID)
'    End If
''    If mCurSendCard.str������ = "�������֤" And Me.ActiveControl Is txt���� Then
'
'        If txtPatient.Text <> "" And cbo�Ա�.ListCount <> 0 And txt��������.Text <> "" Then
'            If strName <> txtPatient.Text Or strSex <> Split(cbo�Ա�.Text, "-")(1) Or txt��������.Text <> Format(datBirthDay, "yyyy-MM-dd") Then
'                    MsgBox "���֤��Ϣ��ҺŲ�����Ϣ��һ��,���ܽ���ǩԼ������", vbInformation, gstrSysName
'                    Exit Sub
'            End If
'        Else
'             MsgBox "�󶨶������֤ʱ,������Ϣ������Ϊ�գ�", vbInformation, gstrSysName
'             Exit Sub
'        End If
'
'        If �Ƿ��Ѿ�ǩԼ(Trim(strID)) Then
'            MsgBox "���֤����Ϊ:" & strID & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
'            txt����.SetFocus
'            Exit Sub
'        Else
'            txt���֤��.Text = strID
'            txt����.Text = strID
'            mbln�Ƿ�ɨ�����֤ = True
'        End If
'
'    End If
    If Me.ActiveControl Is txt���֤�� Then
        
        If txtPatient.Text <> "" And cbo�Ա�.ListCount <> 0 And txt��������.Text <> "" Then
            If strName <> txtPatient.Text Or strSex <> Split(cbo�Ա�.Text, "-")(1) Or txt��������.Text <> Format(datBirthDay, "yyyy-MM-dd") Then
                    MsgBox "���֤��Ϣ�벡����Ϣ��һ��,���ܽ���ǩԼ������", vbInformation, gstrSysName
                    Exit Sub
            End If
        Else
             MsgBox "�󶨶������֤ʱ,������Ϣ������Ϊ�գ�", vbInformation, gstrSysName
             Exit Sub
        End If
        
        If �Ƿ��Ѿ�ǩԼ(Trim(strID)) Then
            MsgBox "���֤����Ϊ:" & strID & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
            txt���֤��.SetFocus
            Exit Sub
        Else
            txt���֤��.Text = strID
            mbln�Ƿ�ɨ�����֤ = True
        End If
        
    End If
    
    Call Show�󶨿ؼ�(mbln�Ƿ�ɨ�����֤ And mblnɨ�����֤ǩԼ)
End Sub

Private Sub cbo���䵥λ_LostFocus()
    '68489:������,2013-12-06,û�����������򲻷����������
    If Trim(txt����.Text) = "" Then Exit Sub
    If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Sub
    
    If Not IsDate(txt��������.Text) Then
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    Call ReLoadCardFee
End Sub

Private Sub cboԤ������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboԤ������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboԤ������.hWnd, KeyAscii)
    If lngIdx <> -2 Then cboԤ������.ListIndex = lngIdx
End Sub

Private Sub chk����_Click()
    If chk����.Value = Checked Then
        cbo���㷽ʽ.Enabled = False
        If Visible Then cmdOK.SetFocus
    Else
        cbo���㷽ʽ.Enabled = True
        cbo���㷽ʽ.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    If mbytInState = E���� And mlngPatientID <> 0 Then
        If MsgBox("��ȷ��Ҫ�����ǰ������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ClearCard
            mblnICCard = False  '���ܷ���clearcard��,��Ϊ�����ȶ����ٲ������
            '����27207 by lesfeng 2010-1-4
            txt����ID.Text = zlDatabase.GetNextNo(1): lbl����ID.Tag = txt����ID.Text
            txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
        End If
    ElseIf mbytInState = E���� And gblnOK Then
        If txtPatient.Text <> "" Then
            If glngSys Like "8??" Then
                If MsgBox("��ǰ�ͻ���Ϣ��δ����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("��ǰ������Ϣ��δ����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        Else
            If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function IsCheck���￨() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-27 10:21:41
    '����:25302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCard As String, strICCard As String
    strCard = UCase(txt����.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    '-----------------------------------------------------------------------------------------------------------------
    '1.���￨�ļ��
    '��۽����
    '���˺�:And tabCardMode.SelectedItem.Key = "CardFee"
    '29134
    '82401:���ϴ�,2015/3/11,�ж϶����Ƿ����
    If mbytInState = E���� And fraCard.Visible = True Then
        If Trim(txt����.Text) <> "" And tabCardMode.SelectedItem.Key = "CardFee" Then
            If Not mCurSendCard.rs���� Is Nothing Then
                If mCurSendCard.rs����!�Ƿ��� = 1 Then
                    If mCurSendCard.rs����!�ּ� <> 0 And Abs(CCur(txt����.Text)) > Abs(mCurSendCard.rs����!�ּ�) Then
                        MsgBox IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��������ֵ���ܴ�������޼ۣ�" & Format(Abs(mCurSendCard.rs����!�ּ�), "0.00"), vbExclamation, gstrSysName
                        If txt����.Enabled And txt����.Visible Then txt����.SetFocus:  Exit Function
                    End If
                    If mCurSendCard.rs����!ԭ�� <> 0 And Abs(CCur(txt����.Text)) < Abs(mCurSendCard.rs����!ԭ��) Then
                        MsgBox IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��������ֵ����С������޼ۣ�" & Format(Abs(mCurSendCard.rs����!ԭ��), "0.00"), vbExclamation, gstrSysName
                        If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    If fraCard.Visible = True Then
        If tabCardMode.SelectedItem.Key = "CardFee" Then
            If cbo���㷽ʽ.Visible And txt����.Text <> "" And cbo���㷽ʽ.Enabled And cbo���㷽ʽ.ListIndex = -1 Then
                MsgBox "��ȷ��" & IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "���Ľɿ���㷽ʽ��", vbExclamation, gstrSysName
                If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus: Exit Function
            End If
        End If
    End If
    
    If txtPass.Text <> txtAudi.Text And fraCard.Visible = True And txt����.Text <> "" Then
        MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus: Exit Function
    End If
    
    If Trim(txt����.Text) = "" And txt����.Visible And mbytInState = E���� And gblnMustCard Then
        MsgBox "��ˢ��������" & IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "���ţ�", vbExclamation, gstrSysName
        If txt����.Enabled And txt����.Enabled Then txt����.SetFocus
        Exit Function
    End If
    If txt����.Text <> "" And mbytInState = E���� Then
        '����ǰ�����￨�Ƿ��У��Ƿ��ڷ�Χ��
        If mCurSendCard.bln�ϸ���� Then
            mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), txt����.Text, mCurSendCard.lng�����ID)
     
            If mCurSendCard.lng����ID <= 0 And Not mCurSendCard.blnOneCard Then
                Select Case mCurSendCard.lng����ID
                    Case 0 '����ʧ��
                    Case -1
'                        If txt����.Text <> "" Then MsgBox "����û�����ü����õ�" & IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��,���ܷ��ţ�" & vbCrLf & _
'                            "�����ڱ������ù������λ�����һ���¿�! ", vbExclamation, gstrSysName
                    Case -2
'                        If txt����.Text <> "" Then MsgBox "���ع��õ�" & IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��������,���ܷ��ţ�" & vbCrLf & _
'                            "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
                    Case -3
                        MsgBox "���ſ��Ų�����Ч��Χ��,�����Ƿ���ȷˢ����", vbExclamation, gstrSysName
                        If txt����.Enabled And txt����.Enabled Then txt����.SetFocus
                End Select
                Exit Function
            End If
        End If
    End If
    '����ǰ,��Ҫ���֧�����
    
    
    IsCheck���￨ = True
End Function
Private Sub SetCardEditEnabled()
    '���þ��￨�༭����
    Dim blnEdit As Boolean
    If Not (mbytInState = E���� Or mbytInState = E�޸�) Then Exit Sub
    blnEdit = Trim(txt����.Text) <> ""
    
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    lbl����.Enabled = txtPass.Enabled: lbl��֤.Enabled = blnEdit
    
    txt����.Enabled = blnEdit: lbl���.Enabled = blnEdit
    chk����.Enabled = blnEdit
    cbo���㷽ʽ.Enabled = chk����.Value = 0 And blnEdit
End Sub

Private Function CanFocus(ctlError As Control) As Boolean
    CanFocus = ctlError.Enabled And ctlError.Visible
End Function

Private Function IsValied(Optional blnModify As Boolean, Optional strBirthDay As String, Optional strAge As String, Optional strSex As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '����:���ݺϷ�,����true,���򷵻�False
    '   ���Σ� blnModify =Trueʱ ���˳������ں��Ա�������������֤��Ϣͬ���������� ������Ϣ���� Ȩ���йأ� =false ֻ�������֤��,������Ϣ��ͬ��������
    '          blnModify=Trueʱ ���� strBirthday,strAge,strSex
    '����:���˺�
    '����:2011-07-26 16:40:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSimilar As String, i As Long, str�ƺ� As String, lngTmp As Long
    Dim str�������� As String, str���� As String, strInfo As String
    Dim blnMod As Boolean, bln������Ϣ���� As Boolean
    Dim strMsg As String
    
    On Error GoTo errHandle
    
    str�ƺ� = IIf(glngSys Like "8??", "�ͻ�", "����")
    
    '65965:������,2013-09-24,����Ԥ����ʾǧλλ��ʽ
    If Not CheckFormInput(Me, "txtԤ����") Then Exit Function
    
    '�Ϸ��Լ��
    If Not IsNumeric(txt�����.Text) And txt�����.Text <> "" Then
        MsgBox "������һ����Ч������ţ�", vbInformation, gstrSysName
        If txt�����.Enabled And txt�����.Visible Then txt�����.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtסԺ��.Text) And txtסԺ��.Text <> "" Then
        MsgBox "������һ����Ч��סԺ�ţ�", vbInformation, gstrSysName
        If txtסԺ��.Enabled And txtסԺ��.Visible Then txtסԺ��.SetFocus: Exit Function
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text And txtPatiMCNO(1).Visible Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "����,ҽ������󳤶Ȳ��ܳ���" & txtPatiMCNO(0).MaxLength & "���ַ���", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If

    If Trim(txtPatient.Text) = "" Then
        MsgBox "��������[����]��", vbExclamation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
    End If
    If cbo�Ա�.ListIndex = -1 Then
        MsgBox "����ȷ��[�Ա�]��", vbExclamation, gstrSysName
        If cbo�Ա�.Enabled And cbo�Ա�.Visible Then cbo�Ա�.SetFocus: Exit Function
    End If
    If Not IsDate(txt��������.Text) Then
        MsgBox "������ȷ����[��������]��", vbInformation, gstrSysName
        If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus: Exit Function
    End If
    If Trim(txt����.Text) = "" Then
        MsgBox "��������[����]��", vbExclamation, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
    End If
    
    '76409,������,2014-08-06,����Ϸ��Լ��
    If txt����.Locked = False Then
        str���� = txt����.Text
        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
        If IsDate(txt��������.Text) Then
            If txt����ʱ��.Text = "__:__" Then
                str�������� = Format(txt��������.Text, "YYYY-MM-DD")
            Else
                str�������� = Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS")
            End If
            strInfo = CheckAge(str����, str��������, CDate(txt��������.Tag))
        Else
            strInfo = CheckAge(str����)
        End If
        If InStr(1, strInfo, "|") > 0 Then
            lngTmp = Val(Split(strInfo, "|")(0)) '1��ֹ,0��ʾ
            strInfo = Split(strInfo, "|")(1)
            If lngTmp = 1 Then
                MsgBox strInfo, vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
            Else
                If MsgBox(strInfo & vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    str�������� = ""
    '--46119,������,2012-08-16,�������֤�Գ������ں�����ļ��
    '���֤���ȼ��
    '--81012,��ΰ��,2014-12-22,�������֤�Գ�������\����\�Ա� ��ͬ������
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) = "�й�" Then
        If Not CheckLen(txt���֤��, 18) Then Exit Function
        lngTmp = LenB(StrConv(Trim(txt���֤��.Text), vbFromUnicode))
        If lngTmp > 0 Then
            If CreatePublicPatient() Then
                strInfo = ""
                If gobjPublicPatient.CheckPatiIdcard(Trim(txt���֤��.Text), strBirthDay, strAge, strSex, strInfo, CDate(txt��������.Tag)) Then
                    'ͬһ���ֻ֤�ܶ�Ӧһ���������˷��½�������ʱ,������֤��
                    If gblnPatiByID And Trim(txt���֤��.Text) <> txt���֤��.Tag And mlng����ID <> 0 Then
                       If gobjPublicPatient.CheckPatiExistByID(Trim(txt���֤��.Text), mlng����ID) Then
                            MsgBox "�Ѵ������֤��Ϊ��" & Trim(txt���֤��.Text) & "���Ľ������ˣ���ֹ�޸����֤�ţ�", vbInformation, gstrSysName
                            If CanFocus(txt���֤��) Then txt���֤��.SetFocus: Exit Function
                        End If
                    End If
                    '���޻�����Ϣ����Ȩ��
                    bln������Ϣ���� = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;") > 0 And ((mlngPatientID > 0 And mbytInState = 0) Or mbytInState = 1)
                    '��������
                    strMsg = ""
                    If CDate(Format(strBirthDay, "YYYY-MM-DD")) <> CDate(Format(txt��������.Text, "YYYY-MM-DD")) Then
                        strMsg = "���֤�����г�������[" & strBirthDay & "]�벡�˳�������[" & Format(txt��������.Text, "YYYY-MM-DD") & "]��һ��"
                        '���� ����λ
                        str���� = txt����.Text
                        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
                        If str���� <> strAge Then
                            strMsg = strMsg & vbCrLf & "���֤����������[" & strAge & "]�벡������[" & str���� & "]��һ��"
                            If str���� Like "*Сʱ*����" Or str���� Like "*����" Or str���� Like "*��*Сʱ" Or str���� Like "*Сʱ" Then
                                strAge = str����
                            End If
                        End If
                    End If
                    If txt����ʱ��.Text <> "__:__" Then
                        strBirthDay = strBirthDay & " " & Format(txt����ʱ��.Text, "HH:MM")
                    End If
                    '�Ա�
                    If InStr(cbo�Ա�.Text, strSex) = 0 Then
                        strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & "���֤�������Ա�[" & strSex & "]�벡���Ա�[" & NeedName(cbo�Ա�.Text) & "]��һ��"
                    End If
                    
                    If ((mlngPatientID > 0 And mbytInState = 0) Or mbytInState = 1) Then
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",�Ƿ������" & vbCrLf & IIf(bln������Ϣ����, "ѡ���ǡ�,�����֤����Ϣ�滻���˵���Ϣ�����ҵ�����ݡ�", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Function
                            Else
                                blnMod = True
                            End If
                        End If
                    Else
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",�Ƿ������" & vbCrLf & "ѡ���ǡ�,�����֤����Ϣ�滻���˵���Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Function
                            Else
                                If CDate(Format(strBirthDay, "YYYY-MM-DD")) <> CDate(Format(txt��������.Text, "YYYY-MM-DD")) Then
                                    txt��������.Text = strBirthDay
                                    If mblnChange = False Then
                                        Call LoadOldData(strAge, txt����, cbo���䵥λ)
                                    End If
                                End If
                                Call zlControl.CboLocate(cbo�Ա�, strSex, False)
                            End If
                        End If
                    End If
                Else
                    MsgBox strInfo, vbInformation + vbOKOnly, gstrSysName
                    If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    If cbo�ѱ�.ListIndex = -1 Then
        MsgBox "����ȷ��[�ѱ�]��", vbExclamation, gstrSysName
        If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus: Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ��[����]��", vbExclamation, gstrSysName
        If cbo����.Enabled And cbo����.Visible Then cbo����.SetFocus: Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ��[����]��", vbExclamation, gstrSysName
        If cbo����.Enabled And cbo����.Visible Then cbo����.SetFocus: Exit Function
    End If
    
    If gbln���ýṹ����ַ And mbytInState <> E���� Then
        For i = PatiAddress.LBound To PatiAddress.UBound
            If Trim(PatiAddress(i).Value) <> "" And PatiAddress(i).CheckNullValue() <> "" Then
                MsgBox "���˵�" & PatiAddress(i).Tag & "¼�벻����,������¼�롣", vbInformation, gstrSysName
                If CanFocus(PatiAddress(i)) = True Then PatiAddress(i).SetFocus
                Exit Function
            End If
        Next
    End If
    '�ֻ��źϷ��Լ��
    If Trim(txtMobile.Text) <> "" Then
        If Not IsNumeric(Trim(txtMobile.Text)) Or Len(Trim(txtMobile.Text)) <> 11 Then
            MsgBox "������ֻ��Ÿ�ʽ����ȷ��������¼�룡", vbInformation, gstrSysName
            If txtMobile.Enabled And txtMobile.Visible Then txtMobile.SetFocus: Exit Function
        Else
            If CheckMobile(Trim(txtMobile.Text), Val(txt����ID.Text)) Then
                If MsgBox("�����еĲ�����Ϣ�д�����ͬ���ֻ���:" & Trim(txtMobile.Text) & "�Ƿ�����¼�룿", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If txtMobile.Enabled And txtMobile.Visible Then txtMobile.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    '���ȼ��
    If Not CheckTextLength("����", txtPatient) Then Exit Function
    If Not CheckTextLength("����", txt����) Then Exit Function
    If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Function
    
    '64701:������,2013-10-31,�޸ĳ�����ַ֧��100���ַ���50������
    If Not CheckLen(txt�����ص�, 100) Then Exit Function
    If Not CheckLen(txt���ڵ�ַ, 100) Then Exit Function
    If Not CheckLen(txt���ڵ�ַ�ʱ�, 6) Then Exit Function
    If Not CheckLen(txt��ͥ��ַ, 100) Then Exit Function
    If Not CheckLen(txt��ͥ��ַ�ʱ�, 6) Then Exit Function
    If Not CheckLen(txt��ͥ�绰, 20) Then Exit Function
    If Not CheckLen(txt��ϵ������, 64) Then Exit Function
    If Not CheckLen(txt��ϵ�˵�ַ, 100) Then Exit Function
    If Not CheckLen(txt��ϵ�˵绰, 20) Then Exit Function
    If Not CheckLen(txt��ϵ�����֤, 18) Then Exit Function
    If Not CheckLen(txt������λ, txt������λ.MaxLength) Then Exit Function
    If Not CheckLen(txt��λ�绰, 20) Then Exit Function
    If Not CheckLen(txtMobile, 20) Then Exit Function
    If Not CheckLen(txt��λ�ʱ�, 6) Then Exit Function
    If Not CheckLen(txt��λ������, 50) Then Exit Function
    If Not CheckLen(txt��λ�ʺ�, 50) Then Exit Function
    If Not CheckLen(txt����, CInt(mCurSendCard.lng���ų���)) Then Exit Function
    If Not CheckLen(txtPass, 10) Then Exit Function
    If Not CheckLen(txt�ɿλ, 50) Then Exit Function
    If Not CheckLen(txt������, 50) Then Exit Function
    If Not CheckLen(txt�ʺ�, 50) Then Exit Function
    If Not CheckLen(txt�������, 30) Then Exit Function
    '����27351 by lesfeng 2010-01-12
    If Not CheckLen(txt��ע, txt��ע.MaxLength) Then Exit Function
    
    '104238:���ϴ���2017/2/15����鿨���Ƿ����㷢����������
    If txt����.Text <> "" And Len(txt����.Text) <> mCurSendCard.lng���ų��� And Not mCurSendCard.bln�ϸ���� Then
        Select Case mCurSendCard.byt��������
            Case 0
                MsgBox "����Ŀ���С��" & mCurSendCard.str������ & "�趨�Ŀ��ų��ȣ����������룡", vbExclamation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                Exit Function
            Case 2
                If MsgBox("����Ŀ���С��" & mCurSendCard.str������ & "�趨�Ŀ��ų��ȣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                    Exit Function
                End If
        End Select
    End If
    
    If IsCheck���￨ = False Then Exit Function
    '���㷽ʽ
    If IsNumeric(txtԤ����.Text) And cboԤ������.Visible And cboԤ������.Enabled And cboԤ������.ListIndex = -1 Then
        MsgBox "��ȷ������Ԥ������㷽ʽ��", vbInformation, gstrSysName
        cboԤ������.SetFocus: Exit Function
    End If
    
    '�����:53408
'    If IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, glngModul) = "1", 1, 0) = 0 And ((mCurSendCard.str������ = "�������֤" And Trim(txt����.Text) <> "") Or Trim(txt֧������.Text) <> "") Then
'         MsgBox "��û��Ȩ�޽���ǩԼ����,�뵽�������������á�ɨ�����֤ǩԼ����", vbOKOnly + vbInformation, gstrSysName
'         txt����.Text = ""
'         txtPass.Text = ""
'         txtAudi.Text = ""
'         If txt����.Visible = True Then txt����.SetFocus
'         Exit Function
'    End If
    
    If Trim(txt֧������.Text) <> "" And Trim(txt���֤��.Text) <> "" Then
           If �Ƿ��Ѿ�ǩԼ(txt���֤��.Text) Then
                 MsgBox "���֤����Ϊ:" & txt���֤��.Text & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
                 txt֧������.Text = ""
                 If txt֧������.Visible = True Then txt֧������.SetFocus
                 Exit Function
           End If
    End If
    
    If mbln�Ƿ�ɨ�����֤ = False And mCurSendCard.str������ = "�������֤" And txt����.Text <> "" Then
            MsgBox "�����ֻ֤����ˢ���ķ�ʽ���У��������ֶ��������֤���а�!", vbOKOnly + vbInformation, gstrSysName
            txt����.Text = ""
            txtPass.Text = ""
            txtAudi.Text = ""
            txt֧������.Text = ""
            txt��֤����.Text = ""
            If txt����.Visible = True Then txt����.SetFocus
            Exit Function
    End If
    
    If mbln�Ƿ�ɨ�����֤ = False And mCurSendCard.str������ <> "�������֤" And txt֧������.Text <> "" Then
            MsgBox "�����ֻ֤����ˢ���ķ�ʽ���У��������ֶ��������֤���а�!", vbOKOnly + vbInformation, gstrSysName
            txt���֤��.Text = ""
            txt֧������.Text = ""
            txt��֤����.Text = ""
            If txt���֤��.Visible = True Then txt���֤��.SetFocus
        Exit Function
    End If
    
    If Trim(txt֧������.Text) <> Trim(txt��֤����.Text) And (Trim(txt֧������.Text) <> "" Or Trim(txt��֤����.Text) <> "") Then
        MsgBox "������������벻һ��,����������", vbOKOnly + vbInformation, gstrSysName
        txt֧������.Text = "": txt��֤����.Text = ""
        If txt֧������.Visible = True Then txt֧������.SetFocus
        Exit Function
    End If
    
    blnModify = blnMod And bln������Ϣ����
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckNewPati() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����²���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSimilar As String, strMCAccount As String, strNote As String
    Dim i As Long, lng�ӿڱ�� As Long, strBalanceInfor As String
    Dim str�ƺ� As String
    Dim lngTmp As Long
    Dim rsSimilar As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If Trim(txt����.Text) <> "" And txtPass.Visible Then
        Select Case mCurSendCard.int���볤������
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "�����������" & mCurSendCard.int���볤�� & "λ", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "�����������" & Abs(mCurSendCard.int���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        End Select
    End If
    If mlngPatientID <> 0 Then CheckNewPati = True: Exit Function
    If Not CheckPatiIsExist() Then Exit Function

    '����ID���
    '����27207 by lesfeng 2010-1-4
    If ExistInPatiID(CLng(txt����ID.Text)) Then
        If txt����ID.Text <> lbl����ID.Tag Then
            MsgBox "��" & str�ƺ� & "�ı�ʶ " & txt����ID.Text & " �Ѿ���ʹ�ã�" & vbCrLf & _
                "ϵͳ���Զ�����һ�����ظ��ı�ʶ��", vbInformation, gstrSysName
            txt����ID.Text = zlDatabase.GetNextNo(1): lbl����ID.Tag = txt����ID.Text
            cmdOK.SetFocus: Exit Function
        Else
            '�Զ������ĺ����û���޸ģ���ֱ���ٴ��Զ���������
            txt����ID.Text = zlDatabase.GetNextNo(1): lbl����ID.Tag = txt����ID.Text
        End If
    End If
    
    '����ż��
    If IsNumeric(txt�����.Text) Then
        '����27207 by lesfeng 2010-1-4
        If ExistClinicNO(txt�����.Text) Then
            If txt�����.Text <> lbl�����.Tag Then
                MsgBox "���ָò��˵Ĳ��������[" & txt�����.Text & "]�Ѿ�����������ʹ��,ϵͳ���Զ�����һ�����ظ��ĺ��룡", vbInformation, gstrSysName
                txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
                cmdOK.SetFocus: Exit Function
            Else
                '�Զ������ĺ����û���޸ģ���ֱ���ٴ��Զ���������
                txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
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
    '����:���ý����������
    '���:blnPrepay-�Ƿ�Ԥ���������
    '����:������
    '����:2014-01-07
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    
    If blnPrepay = True Then
        With cboԤ������
            If .ListIndex = -1 Then Exit Sub
            lngIndex = .ListIndex + 1
        End With
        '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
        If Not mcolPrepayPayMode Is Nothing Then
            With mCurPrepay
                    .lngҽ�ƿ����ID = Val(mcolPrepayPayMode(lngIndex)(3))
                    .bln���ѿ� = Val(mcolPrepayPayMode(lngIndex)(5)) = 1
                    .str���㷽ʽ = Trim(mcolPrepayPayMode(lngIndex)(6))
                    .str���� = Trim(mcolPrepayPayMode(lngIndex)(1))
             End With
        End If
    Else
        With cbo���㷽ʽ
            If .ListIndex = -1 Then Exit Sub
            lngIndex = .ListIndex + 1
        End With
        If Not mcolCardPayMode Is Nothing Then
            With mCurCardPay
                .lngҽ�ƿ����ID = Val(mcolCardPayMode(lngIndex)(3))
                .bln���ѿ� = Val(mcolCardPayMode(lngIndex)(5)) = 1
                .str���㷽ʽ = Trim(mcolCardPayMode(lngIndex)(6))
                .str���� = Trim(mcolCardPayMode(lngIndex)(1))
             End With
         End If
     End If
End Sub

Private Sub cmdOK_Click()
    Dim strMCAccount As String, str�ƺ� As String
    Dim blnOK As Boolean
    Dim blnModify As Boolean
    Dim strErrInfo As String
    Dim str�Ա� As String, str���� As String, str�������� As String
    
    '�����:56599
    tbcPage.Item(0).Selected = True
    
    str�ƺ� = IIf(glngSys Like "8??", "�ͻ�", "����")
    If IsValied(blnModify, str��������, str����, str�Ա�) = False Then Exit Sub
    
    '69231,������,2014-01-07 14:42:55,����ʱǿ�Ƹ��¿���������
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
    
    If IsNumeric(txtԤ����.Text) Then
        mblnPrepayPrint = True
        '����Ƿ��ӡƱ��
        '78751:���ϴ�,2015/08/24,����Ԥ��Ʊ�ݴ�ӡ��ʽ
        Select Case mFactProperty.intInvoicePrint
            Case "0" '����ӡԤ����Ʊ
               mblnPrepayPrint = False
            Case "1" '�Զ���ӡ
               mblnPrepayPrint = True
            Case "2" '��ӡ����
                mblnPrepayPrint = MsgBox("�Ƿ��ӡԤ����Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
        End Select
        
        If mblnPrepayPrint Then
            If gblnBillԤ�� Then
                If Trim(txtFact.Text) = "" Then
                    MsgBox "��������һ����Ч��Ԥ��Ʊ�ݺ��룡", vbInformation, gstrSysName
                    txtFact.SetFocus: Exit Sub
                End If
                
                mlngԤ������ID = CheckUsedBill(2, IIf(mlngԤ������ID > 0, mlngԤ������ID, mFactProperty.lngShareUseID), txtFact.Text, Val(Mid(tbDeposit.SelectedItem.Key, 2)))
                If mlngԤ������ID <= 0 Then
                    Select Case mlngԤ������ID
                        Case 0 '����ʧ��
                        Case -1
                            MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                        Case -2
                            MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                        Case -3
                            MsgBox "Ʊ�ݺ��벻�ڵ�ǰ��Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                            txtFact.SetFocus
                    End Select
                    Exit Sub
                End If
            Else
                If Len(txtFact.Text) <> gbytԤ�� And txtFact.Text <> "" Then
                    MsgBox "Ԥ��Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytԤ�� & " λ��", vbInformation, gstrSysName
                    txtFact.SetFocus: Exit Sub
                End If
            End If
        End If
    End If
    
    '63246:������,2013-07-03
    If CheckPatiCard = False Then Exit Sub
    
    '73937:������,2013-07-03
    If CreatePlugInOK(glngModul) Then
        blnOK = True
        On Error Resume Next
        blnOK = gobjPlugIn.PatiInfoSaveBefore(Val(txt����ID.Text))
        If blnOK = False Then
            If tbcPage.Item(tbcPage.ItemCount).Caption = "������Ϣ" Then tbcPage.Item(tbcPage.ItemCount).Selected = True
            Err.Clear
            Exit Sub
        End If
        Err.Clear: On Error GoTo 0
    End If
    '������Ϣ�ӱ�\������ҳ�ӱ���
    mstrPatiPlus = ""
    '���֤��δ¼��ʱ������Ϣ
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) = "�й�" Then
        mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
        mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:"
    Else
        If Trim(txt���֤��.Text) <> "" Then
            mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:"
            mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:" & Trim(txt���֤��.Text)
            txt���֤��.Text = ""
        Else
            mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
            mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:"
        End If
    End If
    If mstrPatiPlus <> "" Then mstrPatiPlus = Mid(mstrPatiPlus, 2)
    
    If mbytInState = E���� Then
         If CheckNewPati = False Then Exit Sub
        '�����¿�
        '--------------------------------------------------------------
        If Not SaveNewCard(strMCAccount) Then
            MsgBox str�ƺ� & "��ݵǼ�ʧ��,�����Ըò�����", vbExclamation, gstrSysName
            Exit Sub
        End If
        '������Ϣ����ɹ�,�������֤��Ϣͬ������������Ϣ���Ա�,���������
        If blnModify Then
            strErrInfo = ""
            Call gobjPublicPatient.SavePatiBaseInfo(mlng����ID, mlng��ҳID, Trim(txtPatient.Text), str�Ա�, str����, str��������, Me.Caption, IIf(mlng����ID = 0, 1, 2), strErrInfo, False, True)
            If strErrInfo <> "" Then
                MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
            End If
        End If
        '��ӡԤ�����վ�
        '78751:���ϴ�,2015/08/24,����Ԥ��Ʊ�ݴ�ӡ��ʽ
        If mblnPrepayPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mCurPrepay.strNO, "�տ�ʱ��=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
                            "����ID=" & Val(txt����ID), IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
        End If
        
        '��ӡ������ҳ
        If InStr(mstrPrivs, "��ҳ��ӡ") > 0 Then
            If MsgBox("������Ϣ����ɹ���Ҫ��ӡ������ҳ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1101", Me, "����ID=" & Val(txt����ID.Text), 2)
            End If
        End If
        
        gblnOK = True
        
        '����������һ��������Ϣ
        Call ClearCard
        mblnICCard = False  '���ܷ���clearcard��,��Ϊ�����ȶ����ٲ������
        '����27207 by lesfeng 2010-1-4
        txt����ID.Text = zlDatabase.GetNextNo(1): lbl����ID.Tag = txt����ID.Text
        txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
        
        If Not mCurSendCard.rs���� Is Nothing Then txt����.Text = Format(IIf(mCurSendCard.rs����!�Ƿ��� = 1, mCurSendCard.rs����!ȱʡ�۸�, mCurSendCard.rs����!�ּ�), "0.00"): txt����.Tag = txt����.Text
        
        'Ԥ������
        If mblnPrepayPrint Then
            If Not gblnBillԤ�� Then
                zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", txtFact.Text, glngSys, mlngModul
            End If
            Call GetFact(False)
        End If
        
        '���￨���ü��
        If mCurSendCard.bln�ϸ���� Then
            mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), , mCurSendCard.lng�����ID)
            If mCurSendCard.lng����ID <= 0 Then
                Select Case mCurSendCard.lng����ID
                    Case 0 '����ʧ��
                    Case -1
                        If txt����.Text <> "" Then MsgBox "����û�����ü����õ�" & IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��,�����ٷ��ţ�" & vbCrLf & _
                            "�����ڱ������ù������λ�����һ���¿���", vbExclamation, gstrSysName
                    Case -2
                        If txt����.Text <> "" Then MsgBox "���ع��õ�" & IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��������,�㲻���ٷ��ţ�" & vbCrLf & _
                            "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
                End Select
            End If
        End If
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    ElseIf mbytInState = E�޸� Then
        '����ż��
        If IsNumeric(txt�����.Text) Then
            If ExistClinicNO(txt�����.Text, CLng(txt����ID.Text)) Then
                '����27207 by lesfeng 2010-1-4
                If txt�����.Text <> lbl�����.Tag Then
                    MsgBox "���ָò��˵Ĳ��������[" & txt�����.Text & "]�Ѿ�����������ʹ��,ϵͳ���Զ�����һ�����ظ��ĺ��룡", vbInformation, gstrSysName
                    txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
                    cmdOK.SetFocus: Exit Sub
                Else
                    '�Զ������ĺ����û���޸ģ���ֱ���ٴ��Զ���������
                    txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
                End If
            End If
        End If
        
        'סԺ�ż��
        If IsNumeric(txtסԺ��.Text) Then
            If ExistInPatiNO(Trim(txtסԺ��.Text), Val(txt����ID.Text)) Then
                MsgBox "���ָò��˵Ĳ���סԺ��[" & txtסԺ��.Text & "]�Ѿ�����������ʹ��,ϵͳ���Զ�����һ�����ظ��ĺ��룡", vbInformation, gstrSysName
                txtסԺ��.Text = zlDatabase.GetNextNo(2)
                cmdOK.SetFocus: Exit Sub
            End If
        End If
        '�����޸�
        '--------------------------------------------------------------------
        If Not SaveModiCard(strMCAccount) Then
            MsgBox "����ʧ��,�����Ըò�����", vbExclamation, gstrSysName
            Exit Sub
        End If
        '������Ϣ����ɹ�,�������֤��Ϣͬ������������Ϣ���Ա�,���������
        If blnModify Then
            strErrInfo = ""
            Call gobjPublicPatient.SavePatiBaseInfo(mlng����ID, mlng��ҳID, Trim(txtPatient.Text), str�Ա�, str����, str��������, Me.Caption, IIf(mlng��ҳID = 0, 1, 2), strErrInfo, True, True)
            If strErrInfo <> "" Then
                MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
            End If
        End If
        '�޸ĺ��˳�
        gblnOK = True
        Unload Me: Exit Sub
    End If
End Sub

Private Sub cmdOperation_Click(Index As Integer)
    Dim bln��Ԥ�� As Boolean, bln��Ԥ�� As Boolean
    Dim lng����ID As Long
    
    Dim strPrivs As String
    On Error Resume Next
    Select Case Index
    Case 0
        Call InitLocPar(1103)
        strPrivs = ";" & GetPrivFunc(glngSys, 1103) & ";"
        bln��Ԥ�� = InStr(1, strPrivs, ";����Ԥ��;") > 0 Or InStr(1, strPrivs, ";סԺԤ��;") > 0 Or InStr(1, strPrivs, ";����Ԥ��;") > 0
        bln��Ԥ�� = InStr(1, strPrivs, ";Ԥ���˿�;") > 0
        If bln��Ԥ�� = False And bln��Ԥ�� = False Then Exit Sub
        Call frmDeposit.zlShowEdit(Me, 0, IIf(bln��Ԥ��, 0, 2), strPrivs, 1103)
        Call InitLocPar(mlngModul)
    Case 1
        '���þ��￨��������
        strPrivs = ";" & GetPrivFunc(glngSys, 1107) & ";"
        If gobjSquare.objSquareCard.zlSendCard(Me, mlngModul, mCurSendCard.lng�����ID, lng����ID, strPrivs) = False Then Exit Sub
        'frmIDCard.mbytInState = E����
       ' frmIDCard.Show 1, Me
    End Select
    Err.Clear
End Sub

Private Sub cmd�����ص�_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
            " Substr(����,1,2) as ���� From ����" & _
            " Union All" & _
            " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
            " From ���� Order by ����", 2, "����", , txt�����ص�.Text)
    If Not rsTmp Is Nothing Then
        txt�����ص�.Text = rsTmp!����
        txt�����ص�.SelStart = Len(txt�����ص�.Text)
        txt�����ص�.SetFocus
    End If
End Sub

Private Sub cmd��ͬ��λ_Click()
    Dim rsTmp As ADODB.Recordset
    '����27040 by lesfeng �Ժ�Լ��λ���ϳ���ʱ��Ĵ���
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & _
            "  Where (����ʱ�� IS NULL OR TO_CHAR(����ʱ��, 'yyyy-MM-dd') = '3000-01-01') " & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID", _
            2, "��λ", , txt������λ.Text)
    If Not rsTmp Is Nothing Then
        txt������λ.Tag = rsTmp!ID
        txt������λ.Text = rsTmp!����
        txt������λ.SelStart = Len(txt������λ.Text)
        txt��λ�绰.Text = Trim(rsTmp!�绰 & "")
        txt��λ������.Text = Trim(rsTmp!�������� & "")
        txt��λ�ʺ�.Text = Trim(rsTmp!�ʺ� & "")
        txt������λ.SetFocus
    End If
End Sub

Private Sub cmd��ͥ��ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
            " Substr(����,1,2) as ���� From ����" & _
            " Union All" & _
            " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
            " From ���� Order by ����", 2, "����", , txt�����ص�.Text)
    If Not rsTmp Is Nothing Then
        txt��ͥ��ַ.Text = rsTmp!����
        txt��ͥ��ַ.SelStart = Len(txt��ͥ��ַ.Text)
        txt��ͥ��ַ.SetFocus
    End If
End Sub

Private Sub cmd��ϵ�˵�ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
            " Substr(����,1,2) as ���� From ����" & _
            " Union All" & _
            " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
            " From ���� Order by ����", 2, "����", , txt�����ص�.Text)
    If Not rsTmp Is Nothing Then
        txt��ϵ�˵�ַ.Text = rsTmp!����
        txt��ϵ�˵�ַ.SelStart = Len(txt��ϵ�˵�ַ.Text)
        txt��ϵ�˵�ַ.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If mbytInState = E���� And mblnSel = False Then txtPatient.SetFocus
    '�����:53408
    mblnɨ�����֤ǩԼ = IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, glngModul) = "1", 1, 0) = "1"
    If mCurSendCard.str������ Like "*�������֤*" Then
        lbl���￨��.Enabled = False: txt����.Enabled = False
        lbl����.Enabled = False: txtPass.Enabled = False
        lbl��֤.Enabled = False: txtAudi.Enabled = False
    End If
    mblnSel = True
    Call SetCardEditEnabled
    Call Show�󶨿ؼ�(False)
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
            If Me.ActiveControl.Name = txt�����ص�.Name _
                And cmd�����ص�.Enabled And cmd�����ص�.Visible Then
                cmd�����ص�_Click
            ElseIf Me.ActiveControl.Name = txt��ͥ��ַ.Name _
                And cmd��ͥ��ַ.Enabled And cmd��ͥ��ַ.Visible Then
                cmd��ͥ��ַ_Click
            ElseIf Me.ActiveControl.Name = txt��ϵ�˵�ַ.Name _
                And cmd��ϵ�˵�ַ.Enabled And cmd��ϵ�˵�ַ.Visible Then
                cmd��ϵ�˵�ַ_Click
            ElseIf Me.ActiveControl.Name = txt������λ.Name _
                And cmd��ͬ��λ.Enabled And cmd��ͬ��λ.Visible Then
                cmd��ͬ��λ_Click
            ElseIf Me.ActiveControl.Name = txt����.Name And cmd����.Enabled And cmd����.Visible Then
                cmd����_Click
            End If
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC����")
                If intIndex < 0 Then Exit Sub
                IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyReturn
            Set obj = Me.ActiveControl
            If obj.Name = "txtPatient" Then
                If txtPatient.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo�Ա�" Then
                If cbo�Ա�.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo�ѱ�" Then
                If cbo�ѱ�.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo���㷽ʽ" Then
                If cbo���㷽ʽ.ListIndex <> -1 Then cmdOK.SetFocus
            '���� 25458 ���� txtPatiMCNO�ж� ʵ�ֵ��� vbKeyTab
            ElseIf InStr(1, ",txt����,txt�����ص�,txt��ͥ��ַ,txt���ڵ�ַ,txt��ϵ�˵�ַ,txt������λ,txtPass,txtAudi,txt����,txt����,txtԤ����,txtPatiMCNO,vsDrug,vsInoculate,vsLinkMan,vsOtherInfo,PatiAddress,", "," & obj.Name & ",") <= 0 Then
                If Not obj Is txtPass Then
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
        End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    mlngͼ����� = 0: mstr�ɼ�ͼƬ = "":
    With mPageHeight
        .���� = Me.Height
        .�������� = Me.Height
        .������Ϣ = Me.Height
    End With
    '�ϴ�Ĭ��Ԥ������
    mbytPrepayType = Val(zlDatabase.GetPara("�ϴ�Ԥ������", glngSys, mlngModul, "0"))
    '��ʼ��
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
     '��ʼ�ɹ�,��֤���˴��ڴ�����صĽ��㿨
     mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
    'Call zlCardSquareObject:
    Call CreateObjectKeyboard
    
     
    If glngSys Like "8??" Then
        Me.Caption = "�ͻ���Ϣ��Ƭ"
        lbl����ID.Caption = "�ͻ�ID"
        lbl�����.Visible = False
        txt�����.Visible = False
        txt�����.Text = ""
        
        lblסԺ��.Visible = False
        txtסԺ��.Visible = False
        txtסԺ��.Text = ""
        '����27351 by lesfeng 2010-01-12
        txt��ע.Visible = False
        lbl��ע.Visible = False
        txt��ע.Text = ""
        txt��ע.Tag = "0"
        
        chk����.Visible = False
        lbl���㷽ʽ.Visible = True
        
        lbl�ѱ�.Caption = "��Ա�ȼ�"
    Else
        Me.Caption = "������Ϣ" & Choose(mbytInState + 1, "�Ǽ�", "�޸�", "��Ƭ")
        If mbytInState = E���� Then
            lbl�ѱ�.Caption = "����ѱ�" '����ʱ������ΪסԺ�ѱ�
        Else
            If mbytView = 1 Or mbytView = 2 Then
                lbl�ѱ�.Caption = "סԺ�ѱ�"
            Else
                lbl�ѱ�.Caption = "����ѱ�"
            End If
        End If
    End If
    
    '����27356 by lesfeng 2010-01-13
    If InStr(mstrPrivs, "�󶨿���") = 0 Then
        tabCardMode.Tabs.Remove ("CardBind")
        tabCardMode.Width = tabCardMode.Width / 2
    End If
    
    mblnChange = True
    gblnOK = False
    mblnUnLoad = False
    mstrYBPati = ""
    txt��������.Tag = "0"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    mblnChange = False: cbo���䵥λ.ListIndex = 0: mblnChange = True
    '�����:56599
    Call InitCard
    Call InitTabPage
    
    'SetCreateCardObject '����д������
    Call zlCreateSquare
    Call HookDefend(txt֧������.hWnd)
    Call HookDefend(txt��֤����.hWnd)
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
    If cmdOperation(OPT.C0Ԥ����).Visible Then cmdOperation(OPT.C0Ԥ����).Top = cmdHelp.Top
    If cmdOperation(OPT.C1���￨).Visible Then cmdOperation(OPT.C1���￨).Top = cmdHelp.Top
    If cmdOperation(OPT.C0Ԥ����).Visible = False Then cmdOperation(OPT.C1���￨).Left = cmdOperation(OPT.C0Ԥ����).Left
    tbcPage.Height = cmdOK.Top - 120
    tbcPage.Width = Me.ScaleWidth - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mlngͼ����� = 0: mstr�ɼ�ͼƬ = ""

    '�����:565999
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
    mbln������󶨿� = False
    
    '82401:���ϴ�,2015/3/11,�������Ƿ����
    If mbytInState = E���� And fraCard.Visible = True Then
        zlDatabase.SetPara "����ģʽ", tabCardMode.SelectedItem.Key, glngSys, mlngModul
    End If
    
    mblnICCard = False: mbytInState = E����
    mblnUnLoad = False: mlng����ID = 0: mlng��ҳID = 0
    mCurSendCard.lng����ID = 0: mlngԤ������ID = 0: mstrPrivs = ""
    Call ClearCard: mblnSel = False
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    If Not mdicҽ�ƿ����� Is Nothing Then
        Set mdicҽ�ƿ����� = Nothing
    End If
    Err = 0: On Error Resume Next
    'Call zlCardSquareObject(True)
End Sub

Private Sub InitDicts()
    Call ReadDict("�Ա�", cbo�Ա�)
    Call ReadDict("�ѱ�", cbo�ѱ�)
    Call ReadDict("ҽ�Ƹ��ʽ", cboҽ�Ƹ���)
    Call ReadDict("����", cbo����)
    Call ReadDict("����", cbo����)
    Call ReadDict("ѧ��", cboѧ��)
    Call ReadDict("����״��", cbo����״��)
    Call ReadDict("ְҵ", cboְҵ)
    Call ReadDict("���", cbo���)
    Call ReadDict("����ϵ", cbo��ϵ�˹�ϵ)
    Call ReadDict("��������", cbo��������, "��������")
    Call ReadDict("���֤δ¼ԭ��", cboIDNumber)
    If mbytInState = E���� Then
        'Call ReadDict("���㷽ʽ", cbo���㷽ʽ, "���￨")
        'Call ReadDict("���㷽ʽ", cboԤ������, "Ԥ����")
    End If
End Sub

Private Function ReadDict(strDict As String, cbo As ComboBox, Optional strClass As String) As Boolean
'���ܣ���ʼ��ָ���ʵ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long

    On Error GoTo errH
    'by lesfeng 2010-03-08 �����Ż�
    If strDict = "���㷽ʽ" Then
        If strClass = "Ԥ����" Then
            strSQL = "1,2,5,8"
        Else
            strSQL = "1,2"
        End If
        strSQL = "Select Nvl(A.ȱʡ��־,0) as ȱʡ,B.����,B.����,B.����" & _
            " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
            " Where A.���㷽ʽ=B.���� And A.Ӧ�ó���=[1]" & _
            " And Nvl(B.����,1) IN(" & strSQL & ") Order by B.����"
    ElseIf strDict = "���" Then
        strSQL = "Select ����,����,����,Nvl(���ȼ�,0) as ȱʡ From " & strDict & " Order by ����"
    ElseIf strDict = "�ѱ�" Then
        '������ͼ����,��Ϲ��̲���,�����ѱ�������
        'mbytView:0-����,1-��Ժ,2-��Ժ,3-����
        If glngSys Like "8??" Then
            strSQL = "1,3" 'ҩ��ϵͳʹ������ѱ�
        ElseIf mbytInState = E���� Then
            strSQL = "1,3" '����ʱʹ������ѱ�
        Else
            If mbytView = 1 Or mbytView = 2 Then
                strSQL = "2,3" '�鿴/�޸�ʱʹ��סԺ�ѱ�
            Else
                strSQL = "1,3" '�鿴/�޸�ʱʹ������ѱ�
            End If
        End If
        
        '���ǽ��޳������Ψһ����Ŀ(������ȱʡ�ѱ�),������Ч�ڼ估����
        strSQL = _
            " Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ�" & _
            " Where ����=1 And Nvl(���޳���,0)=0 And Nvl(�������,3) IN(" & strSQL & ")" & _
                " And  Sysdate Between NVL(��Ч��ʼ,Sysdate-1) and NVL(��Ч����,Sysdate+1)" & _
            " Order by ����"
    ElseIf strDict = "��������" Then
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ,��ɫ From �������� Order by ����"
    Else
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strClass)
    cbo.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If strDict = "ְҵ" Then
                cbo.AddItem rsTmp!���� & "-" & Chr(&HA) & rsTmp!����
            Else
                cbo.AddItem rsTmp!���� & "-" & rsTmp!����
            End If
            If rsTmp!ȱʡ = 1 Then
                cbo.ListIndex = cbo.NewIndex
                cbo.ItemData(cbo.NewIndex) = 1
            End If
            If strDict = "���㷽ʽ" And strClass = "Ԥ����" Then
                   cbo.ItemData(cbo.NewIndex) = Val(Nvl(rsTmp!����))
                   cbo.Tag = cbo.NewIndex   '��������Ϊȱʡ����������
            End If
            
            If TextWidth(cbo.List(cbo.NewIndex) & "����") > lngMaxW Then lngMaxW = TextWidth(cbo.List(cbo.NewIndex) & "����")
            rsTmp.MoveNext
        Next
        If strDict = "���㷽ʽ" And strClass <> "Ԥ����" Then cbo.Tag = cbo.Text
        
    ElseIf strDict = "���㷽ʽ" Then
        If mbytInState = E���� Then
            If glngSys Like "8??" Then
                MsgBox "��Ա������û�п��õĽ��㷽ʽ�����ܷ�����" & vbCrLf & _
                    "���ȵ����㷽ʽ���������û�Ա���Ľ��㷽ʽ��", vbInformation, gstrSysName
                fraCard.Visible = False: cmdOperation(OPT.C1���￨).Visible = False
                Me.Height = Me.Height - fraCard.Height
                mPageHeight.���� = Me.Height
            Else
                MsgBox "���￨����û�п��õĽ��㷽ʽ��ֻ��ʹ�ü��ʷ�ʽ������" & vbCrLf & _
                    "Ҫʹ�ý��㷢��,���ȵ����㷽ʽ���������þ��￨���㷽ʽ��", vbInformation, gstrSysName
                chk����.Value = 1: chk����.Enabled = False: cbo.Enabled = False
                chk����.Tag = 1
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
    Call zlCommFun.OpenIme(True) '���������뷨
End Sub

Private Sub PatiAddress_LostFocus(Index As Integer)
'����:
    Select Case Index
    
    Case E_IX_��סַ
        txt��ͥ��ַ.Text = PatiAddress(Index).Value
    Case E_IX_�����ص�
        txt�����ص�.Text = PatiAddress(Index).Value
    Case E_IX_���ڵ�ַ
        txt���ڵ�ַ.Text = PatiAddress(Index).Value
    Case E_IX_����
        txt����.Text = PatiAddress(Index).Value
    Case E_IX_��ϵ�˵�ַ
        txt��ϵ�˵�ַ.Text = PatiAddress(Index).Value
    End Select
    Call zlCommFun.OpenIme '�ر��������뷨
End Sub

Private Sub PatiAddress_Validate(Index As Integer, Cancel As Boolean)
    Dim lngLen As Long
    
    lngLen = PatiAddress(Index).MaxLength
    If LenB(StrConv(PatiAddress(Index).Value, vbFromUnicode)) > lngLen Then
        MsgBox PatiAddress(Index).Tag & "ֻ�������� " & lngLen & " ���ַ��� " & lngLen \ 2 & " �����֣�", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub tabCardMode_Click()
    If tabCardMode.SelectedItem.Key = "CardFee" Then
        lbl���.Visible = True
        txt����.Visible = True
        chk����.Visible = True
        cbo���㷽ʽ.Visible = True
    Else
        lbl���.Visible = False
        txt����.Visible = False
        chk����.Visible = False
        cbo���㷽ʽ.Visible = False
    End If
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '�����:56599
    Dim intIndex As Integer, objItem As TabControlItem
    mbln���� = IIf(Item.Caption = "����", True, False)
    Select Case Item.Caption
        Case "����"
            Me.Height = mPageHeight.����
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Case "��������"
            Me.Height = mPageHeight.��������
            If cboBloodType.Enabled And cboBloodType.Visible Then cboBloodType.SetFocus
        Case "������Ϣ"
            Me.Height = mPageHeight.������Ϣ
            If Item.Handle = picTmp.hWnd Then
                intIndex = Item.Index
                Call HideFormCaption(mlngPlugInHwnd, False)
                Set objItem = tbcPage.InsertItem(intIndex, "������Ϣ", mlngPlugInHwnd, 0)
                objItem.Tag = mPageHeight.������Ϣ
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
    mlngԤ������ID = 0
    Call GetFact(False)
End Sub

Private Sub GetFact(Optional blnFirst As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͬ���ķ�Ʊ
    '����:���˺�
    '����:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnBillԤ�� = False Then
        '��ɢ��ȡ��һ������
        txtFact.Text = IncStr(UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngModul, "")))
        Exit Sub
    End If
    '�ϸ�:     ȡ��һ������
    mlngԤ������ID = CheckUsedBill(2, IIf(mlngԤ������ID > 0, mlngԤ������ID, mFactProperty.lngShareUseID), , Val(Mid(tbDeposit.SelectedItem.Key, 2)))
    If mlngԤ������ID <= 0 Then
        Select Case mlngԤ������ID
            Case 0 '����ʧ��
'            Case -1
'                MsgBox "��û�����û��õ�Ԥ��Ʊ��,�Ǽǲ�����Ϣʱ����ͬʱ��Ԥ���" & _
'                    "��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
'            Case -2
'                MsgBox "���صĹ���Ʊ���Ѿ�����,�Ǽǲ�����Ϣʱ����ͬʱ��Ԥ���" & _
'                    "��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
        txtFact.Text = ""
        'fraDeposit.Visible = False
      '  Me.Height = Me.Height - fraDeposit.Height
    Else
        txtFact.Text = GetNextBill(mlngԤ������ID)
    End If
End Sub
Private Sub txtAudi_GotFocus()
    SelAll txtAudi
    OpenPassKeyboard txtAudi, True
End Sub
Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mCurSendCard.int������� = 1 Then
            Call zlControl.TxtCheckKeyPress(txtAudi, KeyAscii, m����ʽ)
        End If
    End If
    
    If KeyAscii = 13 Then
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
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
        If mCurSendCard.int������� = 1 Then
            Call zlControl.TxtCheckKeyPress(txtPass, KeyAscii, m����ʽ)
        End If
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            If Not txt����.Locked And txt����.TabStop And txt����.Enabled Then
                    txt����.SetFocus
            ElseIf chk����.Visible And chk����.Enabled Then
                chk����.SetFocus
            ElseIf Me.cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
                cbo���㷽ʽ.SetFocus
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
'����27351 by lesfeng 2010-01-12  b
Private Sub txt��ע_GotFocus()
    Call zlControl.TxtSelAll(txt��ע)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    CheckInputLen txt��ע, KeyAscii
End Sub

Private Sub txt��ע_LostFocus()
    Call zlCommFun.OpenIme
End Sub
'����27351 by lesfeng 2010-01-12 e
Private Sub txt����ID_Change()
    '����27207 by lesfeng 2010-1-4
    lbl����ID.Tag = "" '��¼�Զ�����Ƿ��˹��޸�
End Sub

Private Sub txt����ID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
        glngTXTProc = GetWindowLong(txt����ID.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt����ID.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����ID_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
        Call SetWindowLong(txt����ID.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt�����ص�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt�����ص�.Text <> "" Then
            '����32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select ���� as ID,����,����,���� From ����" & _
                    " Where ���� Like '" & gstrLike & txt�����ص�.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt�����ص�.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt�����ص�.Text & "%'", _
                    0, "����", , txt�����ص�.Text)
            If Not rsTmp Is Nothing Then
                txt�����ص�.Text = rsTmp!����
                mblnSel = True
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt�����ص�, KeyAscii
End If
End Sub

Private Sub txt�����ص�_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt��������_Change()
    Dim str�������� As String
    If IsDate(txt��������.Text) And mblnChange Then
        mblnChange = False
        txt��������.Text = Format(CDate(txt��������.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        If txt����ʱ��.Text = "__:__" Then
            str�������� = Format(txt��������.Text, "YYYY-MM-DD")
        Else
            str�������� = Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        txt����.Text = ReCalcOld(CDate(str��������), cbo���䵥λ, , , CDate(txt��������.Tag))
    End If
End Sub

Private Sub txt��������_LostFocus()
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
        txt��������.SetFocus
    End If
End Sub

Private Sub txt����ʱ��_Change()
    Dim str�������� As String
    
    If IsDate(txt����ʱ��.Text) And IsDate(txt��������.Text) And mblnChange Then
        str�������� = Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS")
        txt����.Text = ReCalcOld(CDate(str��������), cbo���䵥λ, , , CDate(txt��������.Tag))
    End If
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call OpenIme
    SelAll txt����ʱ��
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt��������) Then
        KeyAscii = 0
        txt����ʱ��.Text = "__:__"
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.SetFocus
        Cancel = True
    End If
End Sub


Private Sub txt��λ�绰_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��λ������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt��λ������, KeyAscii
End Sub

Private Sub txt��λ������_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt��λ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��λ�ʺ�_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt��λ�ʺ�, KeyAscii
End Sub

Private Sub txt������λ_Change()
    If txt������λ.Text = "" Then txt������λ.Tag = ""
End Sub

Private Sub txt������λ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt������λ.Text <> "" Then
            '����27040 by lesfeng �Ժ�Լ��λ���ϳ���ʱ��Ĵ��� '����32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                    " Where ĩ��=1 And (���� Like '" & gstrLike & txt������λ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt������λ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt������λ.Text & "%')" & _
                    " and (����ʱ�� IS NULL OR TO_CHAR(����ʱ��, 'yyyy-MM-dd') = '3000-01-01') ", _
                    0, "��λ", , txt������λ.Text)
            If Not rsTmp Is Nothing Then
                txt������λ.Text = rsTmp!����
                txt������λ.Tag = rsTmp!ID
                txt��λ�绰.Text = Trim(rsTmp!�绰 & "")
                txt��λ������.Text = Trim(rsTmp!�������� & "")
                txt��λ�ʺ�.Text = Trim(rsTmp!�ʺ� & "")
            Else
                txt������λ.Tag = ""
            End If
        Else
            txt������λ.Tag = ""
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt������λ, KeyAscii
    End If
End Sub

Private Sub txt������λ_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt���ڵ�ַ_GotFocus()
    SelAll txt��ͥ��ַ
    Call OpenIme(gstrIme)
End Sub

Private Sub txt���ڵ�ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt���ڵ�ַ.Text <> "" Then
            '����32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select ���� as ID,����,����,���� From ����" & _
                    " Where ���� Like '" & gstrLike & txt���ڵ�ַ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt���ڵ�ַ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt���ڵ�ַ.Text & "%'", _
                    0, "����", , txt���ڵ�ַ.Text)
            If Not rsTmp Is Nothing Then
                txt���ڵ�ַ.Text = rsTmp!����
                mblnSel = True
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt���ڵ�ַ, KeyAscii
    End If
End Sub

Private Sub txt���ڵ�ַ_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt���ڵ�ַ�ʱ�_GotFocus()
    SelAll txt���ڵ�ַ�ʱ�
End Sub

Private Sub txt���ڵ�ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
    Call OpenIme(gstrIme)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                '����27390 by lesfeng 2010-02-25
'                Call zlCommFun.PressKey(vbKeyTab)
            Else
                SelAll txt����
                txt����.SetFocus
            End If
        Else
            '����27390 by lesfeng 2010-02-25
'            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt��ͥ��ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��ͥ��ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ͥ��ַ.Text <> "" Then
            '����32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select ���� as ID,����,����,���� From ����" & _
                    " Where ���� Like '" & gstrLike & txt��ͥ��ַ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt��ͥ��ַ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt��ͥ��ַ.Text & "%'", _
                    0, "����", , txt��ͥ��ַ.Text)
            If Not rsTmp Is Nothing Then
                txt��ͥ��ַ.Text = rsTmp!����
                mblnSel = True
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt��ͥ��ַ, KeyAscii
    End If
End Sub

Private Sub txt��ͥ��ַ_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub



Private Sub txt�ɿλ_GotFocus()
    If IsNumeric(txtԤ����.Text) And txt�ɿλ.Text = "" Then
        txt�ɿλ.Text = txt������λ.Text
    End If
    zlControl.TxtSelAll txt�ɿλ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt�ɿλ_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�ɿλ, KeyAscii
End Sub

Private Sub txt�ɿλ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt�������_GotFocus()
    zlControl.TxtSelAll txt�������
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�������, KeyAscii
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If txt����.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not mCurSendCard.rs���� Is Nothing Then
            If mCurSendCard.rs����!�Ƿ��� = 1 Then
                If mCurSendCard.rs����!�ּ� <> 0 And Abs(CCur(txt����.Text)) > Abs(mCurSendCard.rs����!�ּ�) Then
                    MsgBox IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��������ֵ���ܴ�������޼ۣ�" & Format(Abs(mCurSendCard.rs����!�ּ�), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
                End If
                If mCurSendCard.rs����!ԭ�� <> 0 And Abs(CCur(txt����.Text)) < Abs(mCurSendCard.rs����!ԭ��) Then
                    MsgBox IIf(glngSys Like "8??", "��Ա", mCurSendCard.str������) & "��������ֵ����С������޼ۣ�" & Format(Abs(mCurSendCard.rs����!ԭ��), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
                End If
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(txt����.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0:  Exit Sub
    End If
End Sub

Private Sub txt����_Change()
    Call SetCardEditEnabled
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    
    mbln�Ƿ�ɨ�����֤ = False
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> 13 Then
        If Len(txt����.Text) = mCurSendCard.lng���ų��� - 1 And KeyAscii <> 8 Then
            txt����.Text = txt����.Text & Chr(KeyAscii)
            
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf txt����.Text = "" Then
        KeyAscii = 0: cmdOK.SetFocus  '������,ֱ������
    Else
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    
End Sub

Private Sub txt����_LostFocus()
    Call ReLoadCardFee
    Call CheckFreeCard(txt����.Text)
    Call SetBrushCardObject(False)
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim lngPatientID As Long
    Dim lng�䶯���� As Long
    Dim blnCardBind As Boolean  '���Ƿ���а�

    txt����.Text = Trim(txt����.Text)
    If mCurSendCard.lng���ų��� = Len(Trim(txt����.Text)) Then
        If WhetherTheCardBinding(txt����.Text, mCurSendCard.lng�����ID, lngPatientID) Then
            If mCurSendCard.bln���ƿ� And mCurSendCard.bln�ظ�ʹ�� And lngPatientID > 0 Then
                lng�䶯���� = GetCardLastChangeType(txt����.Text, mCurSendCard.lng�����ID, lngPatientID)
                If lng�䶯���� = 11 Then
                    '����ǰ�
                    If MsgBox("����Ϊ��" & txt����.Text & "����{" & mCurSendCard.str������ & "}�Ŀ��Ѿ��벡�˱�ʶΪ��" & lngPatientID & "���Ľ����˰󶨣�" & vbCrLf & "�Ƿ�ȡ���ÿ��İ�?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Cancel = True
                        txt����.Text = ""
                        Exit Sub
                    End If
                    If BlandCancel(mCurSendCard.lng�����ID, Trim(txt����.Text), lngPatientID) Then
                        Exit Sub
                    End If
                End If
            End If

            MsgBox "�ÿ����Ѿ�����,���ܰ󶨸ÿ���.", vbInformation, gstrSysName
            Cancel = True
            txt����.Text = ""
            Exit Sub
        End If
    End If

End Sub

Private Sub CheckFreeCard(ByVal strCard As String)
'���ܣ���һ��ͨģʽ�µĿ��ţ��ϸ����Ʊ��ʱ������Ƿ���Ʊ�����÷�Χ�ڣ���Χ֮��Ŀ����շ�
    
    If txt����.Visible = False Then Exit Sub
    
    If Not mCurSendCard.rs���� Is Nothing And Val(txt����.Text) = 0 Then  '�Ȼָ�
        txt����.Text = Format(IIf(mCurSendCard.rs����!�Ƿ��� = 1, mCurSendCard.rs����!ȱʡ�۸�, mCurSendCard.rs����!�ּ�), "0.00")
        txt����.Tag = txt����.Text
    End If
    '142204:���ϴ���2020/6/18���������IDʱ��Ҫ���뿨���ID�������Ƿ���ݷѱ���۽����Ƿ����ηѱ��йأ����Ƿ����޹�
    If mCurSendCard.blnOneCard And mCurSendCard.bln�ϸ���� Then
        mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), strCard, mCurSendCard.lng�����ID)
        If mCurSendCard.lng����ID <= 0 Then txt����.Text = "0.00": txt����.Tag = txt����.Text
    End If

    If Not mCurSendCard.rs���� Is Nothing And Val(txt����.Text) <> 0 Then
        If Val(Nvl(mCurSendCard.rs����!���ηѱ�)) <> 1 Then
            txt����.Text = Format(GetActualMoney(NeedName(cbo�ѱ�.Text), mCurSendCard.rs����!������ĿID, IIf(mCurSendCard.rs����!�Ƿ��� = 1, mCurSendCard.rs����!ȱʡ�۸�, mCurSendCard.rs����!�ּ�), mCurSendCard.rs����!�շ�ϸĿID), "0.00")
            txt����.Tag = txt����.Text
        End If
    End If
End Sub

Private Sub cbo�ѱ�_Click()
     
    If Not mCurSendCard.rs���� Is Nothing And Val(txt����.Text) <> 0 And txt����.Visible And Trim(txt����.Text) <> "" Then
        If Val(Nvl(mCurSendCard.rs����!���ηѱ�)) <> 1 Then
            txt����.Text = Format(GetActualMoney(NeedName(cbo�ѱ�.Text), mCurSendCard.rs����!������ĿID, IIf(mCurSendCard.rs����!�Ƿ��� = 1, mCurSendCard.rs����!ȱʡ�۸�, mCurSendCard.rs����!�ּ�), mCurSendCard.rs����!�շ�ϸĿID), "0.00")
            txt����.Tag = txt����.Text
        End If
    End If
End Sub

Private Sub txt������_GotFocus()
    If IsNumeric(txtԤ����.Text) And txt������.Text = "" Then
        txt������.Text = txt��λ������.Text
    End If
    zlControl.TxtSelAll txt������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�ɿλ, KeyAscii
End Sub

Private Sub txt������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��ϵ�˵�ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ϵ�˵�ַ.Text <> "" Then
            '����32632 by lesfeng 2010-09-07
            Set rsTmp = frmPubSel.ShowSelect(Me, _
                    " Select ���� as ID,����,����,���� From ����" & _
                    " Where ���� Like '" & gstrLike & txt��ϵ�˵�ַ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt��ϵ�˵�ַ.Text & "%'" & _
                    " Or ���� Like '" & gstrLike & txt��ϵ�˵�ַ.Text & "%'", _
                    0, "����", , txt��ϵ�˵�ַ.Text)
            If Not rsTmp Is Nothing Then
                txt��ϵ�˵�ַ.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt��ϵ�˵�ַ, KeyAscii
    End If
End Sub

Private Sub txt��ϵ�˵�ַ_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt��ϵ�˵绰_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��ϵ�˵绰_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ�˵绰") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ�˵绰")) = txt��ϵ�˵绰.Text
    End If
End Sub

Private Sub txt��ϵ�����֤_GotFocus()
    SelAll txt��ϵ�����֤
End Sub

Private Sub txt��ϵ�����֤_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt��ϵ�����֤_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ�����֤��") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ�����֤��")) = txt��ϵ�����֤.Text
    End If
End Sub

Private Sub txt��ϵ������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt��ϵ������, KeyAscii
End Sub

Private Sub txt��ϵ������_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt��ϵ������_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ������") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ������")) = txt��ϵ������.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt��ϵ������.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt�����_Change()
    '����27207 by lesfeng 2010-1-4
    lbl�����.Tag = "" '��¼�Զ�����Ƿ��˹��޸�
End Sub

Private Sub txt�����_GotFocus()
    SelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8) & Chr(22), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�����_Validate(Cancel As Boolean)
    If Val(txt�����.Text) = 0 And Val(txt�����.Tag) <> 0 Then txt�����.Text = txt�����.Tag
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            Call cbo���䵥λ.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
    Call OpenIme(gstrIme)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                '����27390 by lesfeng 2010-02-25
'                Call zlCommFun.PressKey(vbKeyTab)
            Else
                SelAll txt����
                txt����.SetFocus
            End If
        Else
            '����27390 by lesfeng 2010-02-25
'            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt���֤��_Change()
    Dim strTmp As String
    
    If mblnChange Then
        strTmp = GetIDDate(txt���֤��.Text)
        If IsDate(strTmp) And txt��������.Enabled = True Then txt��������.Text = strTmp
    End If
    If mblnɨ�����֤ǩԼ Then
        OpenIDCard txt���֤��.Text = ""
    End If
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    '�����:53408
    mbln�Ƿ�ɨ�����֤ = False

    Call Show�󶨿ؼ�(mbln�Ƿ�ɨ�����֤ And mblnɨ�����֤ǩԼ)
    
    If zl��ǰ�û����֤�Ƿ��(Val(IIf(Trim(txt����ID.Text) = "", "0", Trim(txt����ID.Text)))) = True Then
            MsgBox "��ǰ�û������֤���Ѿ��󶨣��������޸������֤��", vbInformation, gstrSysName
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

Private Sub txt����_GotFocus()
    Call zlCommFun.OpenIme
    SelAll txt����
End Sub

Private Sub txt��������_GotFocus()
    Call OpenIme
    SelAll txt��������
End Sub

Private Sub txt���֤��_GotFocus()
    SelAll txt���֤��
    '�����:53408
    If mblnɨ�����֤ǩԼ = True Then
        Call OpenIDCard(txt���֤��.Text = "")
    End If
End Sub

Private Sub txt�����ص�_GotFocus()
    SelAll txt�����ص�
    Call OpenIme(gstrIme)
End Sub

Private Sub txt��ͥ��ַ_GotFocus()
    SelAll txt��ͥ��ַ
    Call OpenIme(gstrIme)
End Sub

Private Sub txt��ͥ��ַ�ʱ�_GotFocus()
    SelAll txt��ͥ��ַ�ʱ�
End Sub

Private Sub txt��ͥ�绰_GotFocus()
    SelAll txt��ͥ�绰
End Sub

Private Sub txt��ϵ������_GotFocus()
    SelAll txt��ϵ������
    Call OpenIme(gstrIme)
End Sub

Private Sub txt��ϵ�˵�ַ_GotFocus()
    SelAll txt��ϵ�˵�ַ
    Call OpenIme(gstrIme)
End Sub

Private Sub txt��ϵ�˵绰_GotFocus()
    SelAll txt��ϵ�˵绰
End Sub

Private Sub txt������λ_GotFocus()
    SelAll txt������λ
    Call OpenIme(gstrIme)
End Sub

Private Sub txt��λ�绰_GotFocus()
    SelAll txt��λ�绰
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    SelAll txt��λ�ʱ�
End Sub

Private Sub txt��λ������_GotFocus()
    SelAll txt��λ������
    Call OpenIme(gstrIme)
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
    Call SetBrushCardObject(True)
End Sub
Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����֤������
    '����:����
    '����:2012-08-31 16:28:23
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (blnEnabled)
End Sub
Private Sub txtPass_GotFocus()
    SelAll txtPass
    OpenPassKeyboard txtPass, False
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
End Sub

Private Sub txt��λ�ʺ�_GotFocus()
    SelAll txt��λ�ʺ�
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo�Ա�.Locked = True Then Exit Sub
    If SendMessage(cbo�Ա�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�Ա�.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo�ѱ�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�ѱ�.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo�ѱ�.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cboѧ��_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboѧ��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboѧ��.hWnd, KeyAscii)
    If lngIdx <> -2 Then cboѧ��.ListIndex = lngIdx
End Sub

Private Sub cbo����״��_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����״��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����״��.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo����״��.ListIndex = lngIdx
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboְҵ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboְҵ.hWnd, KeyAscii)
    If lngIdx <> -2 Then cboְҵ.ListIndex = lngIdx
End Sub

Private Sub cbo���_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo���.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo���.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo���.ListIndex = lngIdx
End Sub

Private Sub cbo��ϵ�˹�ϵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo��ϵ�˹�ϵ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo��ϵ�˹�ϵ.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo��ϵ�˹�ϵ.ListIndex = lngIdx
End Sub

Private Sub cbo���㷽ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo���㷽ʽ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo���㷽ʽ.hWnd, KeyAscii)
    If lngIdx <> -2 Then
        cbo���㷽ʽ.ListIndex = lngIdx
        Call cbo��ϵ�˹�ϵ_Click
    End If
End Sub

Private Function CheckMCOutMode(ByVal strMCCode As String) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From ������� Where ���=1 And ���=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMCCode)

    CheckMCOutMode = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'����27390 by lesfeng 2010-01-25
Private Sub InitInputTabStop()
'���ܣ����ݱ������ù��Ҫ��λ��������Ŀ
    Dim i As Integer
    cbo����.TabStop = zlDatabase.GetPara("����", glngSys, mlngModul, 1)
    cbo����.TabStop = zlDatabase.GetPara("����", glngSys, mlngModul, 1)
    cboѧ��.TabStop = zlDatabase.GetPara("ѧ��", glngSys, mlngModul, 1)
    cbo����״��.TabStop = zlDatabase.GetPara("����״��", glngSys, mlngModul, 1)
    cboְҵ.TabStop = zlDatabase.GetPara("ְҵ", glngSys, mlngModul, 1)
    cbo���.TabStop = zlDatabase.GetPara("���", glngSys, mlngModul, 1)
    txt��������.TabStop = zlDatabase.GetPara("��������", glngSys, mlngModul, 1)
    txt����ʱ��.TabStop = zlDatabase.GetPara("��������", glngSys, mlngModul, 1)
    txt���֤��.TabStop = zlDatabase.GetPara("���֤��", glngSys, mlngModul, 1)
    cboIDNumber.TabStop = zlDatabase.GetPara("���֤��", glngSys, mlngModul, 1)
    txt��ͥ��ַ�ʱ�.TabStop = zlDatabase.GetPara("��ͥ��ַ�ʱ�", glngSys, mlngModul, 1)
    txt��ͥ�绰.TabStop = zlDatabase.GetPara("��ͥ�绰", glngSys, mlngModul, 1)
    txt���ڵ�ַ�ʱ�.TabStop = zlDatabase.GetPara("���ڵ�ַ�ʱ�", glngSys, mlngModul, 1)
    txt��ϵ������.TabStop = zlDatabase.GetPara("��ϵ������", glngSys, mlngModul, 1)
    cbo��ϵ�˹�ϵ.TabStop = zlDatabase.GetPara("��ϵ�˹�ϵ", glngSys, mlngModul, 1)
    txt��ϵ�˵绰.TabStop = zlDatabase.GetPara("��ϵ�˵绰", glngSys, mlngModul, 1)
    txt��ϵ�����֤.TabStop = zlDatabase.GetPara("��ϵ�����֤��", glngSys, mlngModul, 1)
    txt������λ.TabStop = zlDatabase.GetPara("������λ", glngSys, mlngModul, 1)
    txt��λ�绰.TabStop = zlDatabase.GetPara("��λ�绰", glngSys, mlngModul, 1)
    txt��λ�ʱ�.TabStop = zlDatabase.GetPara("��λ�ʱ�", glngSys, mlngModul, 1)
    txt��λ������.TabStop = zlDatabase.GetPara("��λ������", glngSys, mlngModul, 1)
    txt��λ�ʺ�.TabStop = zlDatabase.GetPara("��λ�ʺ�", glngSys, mlngModul, 1)
    txt����֤��.TabStop = zlDatabase.GetPara("����֤��", glngSys, mlngModul, 1)
    txt����.TabStop = zlDatabase.GetPara("����", glngSys, mlngModul, 1)
    
    If gbln���ýṹ����ַ Then
        PatiAddress(E_IX_�����ص�).TabStop = zlDatabase.GetPara("�����ص�", glngSys, mlngModul, 1)
        PatiAddress(E_IX_����).TabStop = zlDatabase.GetPara("����", glngSys, mlngModul, 1)
        PatiAddress(E_IX_��סַ).TabStop = zlDatabase.GetPara("��סַ", glngSys, mlngModul, 1)
        PatiAddress(E_IX_���ڵ�ַ).TabStop = zlDatabase.GetPara("���ڵ�ַ", glngSys, mlngModul, 1)
        PatiAddress(E_IX_��ϵ�˵�ַ).TabStop = zlDatabase.GetPara("��ϵ�˵�ַ", glngSys, mlngModul, 1)
    Else
        txt�����ص�.TabStop = zlDatabase.GetPara("�����ص�", glngSys, mlngModul, 1)
        txt����.TabStop = zlDatabase.GetPara("����", glngSys, mlngModul, 1)
        txt��ͥ��ַ.TabStop = zlDatabase.GetPara("��סַ", glngSys, mlngModul, 1)
        txt���ڵ�ַ.TabStop = zlDatabase.GetPara("���ڵ�ַ", glngSys, mlngModul, 1)
        txt��ϵ�˵�ַ.TabStop = zlDatabase.GetPara("��ϵ�˵�ַ", glngSys, mlngModul, 1)
    End If
End Sub

Private Sub InitCard()
'���ܣ�������ڲ������ÿ�Ƭ״̬
    Dim i As Long, arrTmp As Variant
    
    Call InitvsDrug
    Call InitVsInoculate
    Call InitVsOtherInfo
    Call InitCombox
    '�ṹ����ַ
    Call InitStructAddress
    
    If mbytInState <> 2 Then
        txtPatient.MaxLength = GetColumnLength("������Ϣ", "����")
        txt����.MaxLength = GetColumnLength("������Ϣ", "����")
        txt�����.MaxLength = GetColumnLength("������Ϣ", "�����")
        txtסԺ��.MaxLength = GetColumnLength("������Ϣ", "סԺ��")
    End If
    
    '����27390 by lesfeng 2010-01-25
    Call InitInputTabStop
    
    If InStr(mstrPrivs, "��Լ���˵Ǽ�") = 0 Then
        txt������λ.Enabled = False
        txt������λ.BackColor = &H8000000F
        txt��λ�绰.Enabled = False
        txt��λ�绰.BackColor = &H8000000F
        txt��λ�ʱ�.Enabled = False
        txt��λ�ʱ�.BackColor = &H8000000F
        txt��λ������.Enabled = False
        txt��λ������.BackColor = &H8000000F
        txt��λ�ʺ�.Enabled = False
        txt��λ�ʺ�.BackColor = &H8000000F
        cmd��ͬ��λ.Visible = False
    End If
    
    cbo��������.Enabled = InStr(mstrPrivs, "������������") > 0
    txt�����.Enabled = InStr(mstrPrivs, ";�����޸������;") > 0

    mlngOutModeMC = 0
    arrTmp = Split(GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", ""), ",")
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
    txtPatiMCNO(0).ToolTipText = "��󳤶�" & txtPatiMCNO(0).MaxLength & "λ"
    txtPatiMCNO(1).MaxLength = txtPatiMCNO(0).MaxLength
    If mlngOutModeMC = 0 Or mbytInState = E���� Then
        txtPatiMCNO(1).Visible = False
        lblPatiMCNO(1).Visible = False
    End If
    
    Call InitDicts
    If cbo�ѱ�.ListCount = 0 Then
        MsgBox "û�����÷ѱ���Ϣ,���ȵ��ѱ�ȼ����������ã�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    IDKind.Enabled = mbytInState = E����
    Select Case mbytInState
        Case E����   '����
            If Not gobjSquare.objSquareCard Is Nothing Then
                IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
            Set mobjIDCard = New clsIDCard
            Set mobjICCard = New clsICCard
            Call mobjIDCard.SetParent(Me.hWnd)
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
            Call InitPrepayType: Call InitSendCardPreperty
            chk����.Value = IIf(gbln���� = True, 1, 0)
            chk����.Tag = IIf(chk����.Value = 1, 1, 0)
            '����27207 by lesfeng 2010-1-4
            txt����ID.Text = zlDatabase.GetNextNo(1): lbl����ID.Tag = txt����ID.Text
            
            cmdYB.Left = lbl�Ա�.Left - lbl�Ա�.Width
            If Not glngSys Like "8??" Then txt�����.Text = zlDatabase.GetNextNo(3): lbl�����.Tag = txt�����.Text
            '74299:������,2014-07-03,������ϢҲ���Խ��в�����������
            '����ʱ�������Ͳ��ɼ�
            'lblPatiType.Visible = False: cbo��������.Visible = False: lblPatiColor.Visible = False
            Call Load֧����ʽ
            '89980���˽ṹ�� ������������ȱʡֵ
            If gbln���ýṹ����ַ Then
                Call LoadStructAddressDef(marrAddress)
                Call SetStrutAddress(2)
            End If
        Case E�޸�  '�޸�
            If Not gobjSquare.objSquareCard Is Nothing Then
                IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
        
            If Not glngSys Like "8??" Then
                lblסԺ��.Visible = True
                txtסԺ��.Visible = True
                '����27351 by lesfeng 2010-01-12
                txt��ע.Visible = True
                lbl��ע.Visible = True
                txt��ע.Tag = "1"  '����Ƿ���ʾ��ע
                cmdYB.Visible = False
            End If
            
            If Not ReadPatiCard(mlng����ID) Then
                If glngSys Like "8??" Then
                    MsgBox "�ͻ���Ϣ��ȡʧ�ܣ�", vbExclamation, gstrSysName
                Else
                    MsgBox "������Ϣ��ȡʧ�ܣ�", vbExclamation, gstrSysName
                End If
                mblnUnLoad = True: Exit Sub
            End If
            Call EMPI_LoadPati
        Case E����  '�鿴
            fraInfo.Enabled = False
            PicHealth.Enabled = False
            cmdOK.Visible = False
            cboIDNumber.Locked = True
            cmdCancel.Caption = "�˳�(&X)"
            
            If Not ReadPatiCard(mlng����ID) Then
                If glngSys Like "8??" Then
                    MsgBox "�ͻ���Ϣ��ȡʧ�ܣ�", vbExclamation, gstrSysName
                Else
                    MsgBox "������Ϣ��ȡʧ�ܣ�", vbExclamation, gstrSysName
                End If
                mblnUnLoad = True: Exit Sub
            End If
    End Select
    
    '�������
    If mbytInState <> E���� Then '�޸ĺͲ鿴������ʾԤ����ͷ�������
        fraDeposit.Visible = False: cmdOperation(OPT.C0Ԥ����).Visible = False
        fraCard.Visible = False: cmdOperation(OPT.C1���￨).Visible = False
        Me.Height = Me.Height - fraDeposit.Height
        Me.Height = Me.Height - fraCard.Height
        mPageHeight.���� = Me.Height
    End If
    '�Ƿ���ʾ��ע
    txt��ע.Visible = (Val(txt��ע.Tag) = 1)
    lbl��ע.Visible = (Val(txt��ע.Tag) = 1)
End Sub

Private Sub ClearCard()
    mblnEMPI = False
    mlngPatientID = 0
    '55251:������,2012-10-26
    mlng����ID = 0: mlng��ҳID = 0
    mblnICCard = False
    mstrYBPati = ""
    
    txt�����.Text = ""
    txtסԺ��.Text = ""
    txtPatient.Text = ""
    '�Բ����������Ա𡢳������ڡ�����Ľ���
    txtPatient.Locked = False
    txtPatient.BackColor = &H80000005
    cbo�Ա�.Locked = False
    cbo�Ա�.BackColor = txtPatient.BackColor
    txt��������.Enabled = True
    txt��������.BackColor = txtPatient.BackColor
    txt��������.Tag = "0"
    txt����ʱ��.Enabled = True
    txt����ʱ��.BackColor = txtPatient.BackColor
    txt����.Locked = False
    txt����.BackColor = txtPatient.BackColor
    cbo���䵥λ.Locked = False
    cbo���䵥λ.BackColor = txtPatient.BackColor
    txtPatiMCNO(0).Text = "": txtPatiMCNO(0).Tag = "": txtPatiMCNO(1).Text = ""
    
    txt����.Text = "": Call txt����_Validate(False)
    txt��������.Text = "____-__-__"
    txt����ʱ��.Text = "__:__"
    txt���֤��.Text = ""
    txt����֤��.Text = ""
    txt�����ص�.Text = ""
    txt��ͥ��ַ.Text = ""
    txt��ͥ��ַ�ʱ�.Text = ""
    txt��ͥ�绰.Text = ""
    txt���ڵ�ַ.Text = ""
    txt���ڵ�ַ�ʱ�.Text = ""
    txt����.Text = ""
    txt����.Text = ""
    txt��ϵ������.Text = ""
    txt��ϵ�˵�ַ.Text = ""
    txt��ϵ�˵绰.Text = ""
    txt��ϵ�����֤.Text = ""
    txt������λ.Text = "": txt������λ.Tag = ""
    txt������λ.Text = ""
    txt��λ�绰.Text = ""
    txt��λ�ʱ�.Text = ""
    txt��λ������.Text = ""
    txt��λ�ʺ�.Text = ""
    txt����.Text = ""
    txtPass.Text = ""
    txtAudi.Text = ""
    txtMobile.Text = ""
    '����27351 by lesfeng 2010-01-12
    txt��ע.Text = ""
    
    If gbln���ýṹ����ַ Then
        Call SetStrutAddress(1)
        Call SetStrutAddress(2)
    End If
    
    chk����.Value = IIf(gbln���� = True, 1, 0)
    cboIDNumber.ListIndex = -1 'ȱʡ
    cboIDNumber.Enabled = True
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
    Call SetCboDefault(cbo���㷽ʽ)
    Call SetCboDefault(cbo��������)
    '74299:������,2014-07-03,������ϢҲ���Խ��в�����������
    '��������ʱ���ɼ�
    'If mbytInState = E���� Then lblPatiType.Visible = False: cbo��������.Visible = False: lblPatiColor.Visible = False
    'Ԥ����Ϣ
    txtԤ����.Text = ""
    txt�ɿλ.Text = ""
    txt�ʺ�.Text = ""
    txt������.Text = ""
    txt�������.Text = ""
    '�����:51072
    txt��ϵ�����֤.Text = ""
    '�����:53408
    txt֧������.Text = ""
    txt��֤����.Text = ""
    txt��֤����.Tag = ""
    txt֧������.Enabled = False
    txt��֤����.Enabled = False
    lbl֧������.Enabled = False
    lbl��֤����.Enabled = False
    
    mlngͼ����� = 0: mstr�ɼ�ͼƬ = ""
    imgPatient.Picture = Nothing
    '�����:56599
    Call Clear��������
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    Dim lngIndex As Long
    
    If IDKind.GetCurCard.���� = "�����" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    If mlngPatientID <> 0 Then Exit Sub
        
    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Or IDKind.IDKind = IDKind.GetKindIndex("סԺ��") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    '55571:������,2012-11-12
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
'����:���Ҳ���
'����:������
'����:2012-10-25
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean
    Dim lngPatientID As Long, lngIndex As Long
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ = True Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    '��ȡ������Ϣ
    lngPatientID = GetPatient(objCard, strInput, blnCard)
    lngPatientIDRef = lngPatientID
    If lngPatientID <> 0 Then
        Call ClearCard
        mlngPatientID = lngPatientID
        txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
        Call ReadPatiCard(mlngPatientID)
    Else
        If (blnICCard Or blnCard) And fraCard.Visible Then '���¿�
            MsgBox "�ÿ�û�н���,����Ϊ�¿��Ǽ�,�����벡��������", vbInformation, gstrSysName
            txt����.Text = strInput
            lngIndex = IDKind.GetKindIndex("����")
            txtPatient.Text = "": txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            If lngIndex >= 0 Then IDKind.IDKind = lngIndex
            Call CheckFreeCard(txt����.Text)
            
        ElseIf Not (IDKind.GetCurCard.���� Like "����*" And InStr("+-*", Left(strInput, 1)) = 0) Then
           txtPatient.Text = "": txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
           MsgBox "û���ҵ�ָ���Ĳ��ˡ�", vbInformation, gstrSysName
        End If
    End If
    Call zlControl.TxtSelAll(txtPatient)
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    strSQL = "Select A.����ID From ������Ϣ A Where A.ͣ��ʱ�� is NULL "
    
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then
            '���ֻ��Ų�ѯ
            If IDKind.IsMobileNo(strInput) = False Then GoTo NotFoundPati:
            If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Function
        End If
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSQL = strSQL & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    Else
        Select Case objCard.����
            Case "����"
                '�������������²���
                Exit Function
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
            Case Else
                If InStr(",���֤��,���֤,�������֤,", "," & objCard.���� & ",") > 0 Then
                    If Not CreatePublicPatient() Then Exit Function
                    lng����ID = gobjPublicPatient.GetPatiByID(glngModul, strInput)
                End If
                If lng����ID = 0 Then
                    '��������,��ȡ��صĲ���ID
                    If Val(objCard.�ӿ����) > 0 Then
                        lng�����ID = Val(objCard.�ӿ����)
                        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                        If lng����ID = 0 Then GoTo NotFoundPati:
                    Else
                        If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                            strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    End If
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If rsTmp.RecordCount > 0 Then GetPatient = rsTmp!����ID
    mblnICCard = IDKind.IDKind = IDKind.GetKindIndex("IC����")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Function
NotFoundPati:
End Function

Private Function ReadPatiCard(ByVal lng����ID As Long) As Integer
'���ܣ��޸Ļ�鿴ʱ,��ȡָ��������Ϣ,����ʾ�ڽ�����
'���أ�
'     -1=�ɹ�
'      0=ʧ��
'      1=�ò��˲�����
    Dim rsTmp As New ADODB.Recordset
    Dim str�ѱ� As String
    '����27351 by lesfeng 2010-01-12
    On Error GoTo errH
    '�����:51071
    gstrSQL = "Select A.�����,A.סԺ��,A.��ҳID �������,A.����,A.�Ա�,A.�ѱ�,A.ҽ�Ƹ��ʽ,A.����,A.����,A.����,A.ѧ��,A.����״��," & _
        " A.ְҵ,A.���,Decode(nvl(A.��Ժ,0),0,A.����,B.����) as ����,A.��������,A.���֤��,A.�ֻ���,A.�����ص�,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.���ڵ�ַ,A.���ڵ�ַ�ʱ�,A.����,A.������,A.������,A.��������," & _
        " A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.������λ,A.��ͬ��λID,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.��ϵ�����֤��," & _
        " B.����ID,B.�ѱ� as סԺ�ѱ�,Nvl(B.����,A.����) as ����,Nvl(A.ҽ����,D.��Ϣֵ) as ҽ����,A.����֤��," & _
        IIf(mstrYBPati = "", " NVL(Decode(B.����ID,Null,A.��������,B.��������),Decode(A.����,Null,'��ͨ����','ҽ������'))", "zl_PatiType(A.����ID)") & " ��������,B.��ע,B.��Ժ����,B.��Ժ���� " & _
        " From ������Ϣ A,������ҳ B,������ҳ�ӱ� D" & _
        " Where A.����ID=B.����ID(+) And Nvl(A.��ҳID,0)=B.��ҳID(+)" & _
        " And A.����ID=D.����ID(+) And Nvl(A.��ҳID,0)=D.��ҳID(+) And D.��Ϣ��(+)='ҽ����' And A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    
    If rsTmp.RecordCount = 0 Then ReadPatiCard = 1: Exit Function
    
    mlngPatientID = lng����ID
    
    mlng���� = Nvl(rsTmp!����, 0)
    txt����ID.Text = lng����ID
    '55251,�����ɣ�2012-10-26
    mlng����ID = Val(txt����ID.Text)
    txt�����.Text = Nvl(rsTmp!�����)
    txt�����.Tag = Nvl(rsTmp!�����)
    txtסԺ��.Text = Nvl(rsTmp!סԺ��)
    txtסԺ��.Tag = Nvl(rsTmp!סԺ��)
    txtPatient.Text = rsTmp!����
    '�����:51071
    txt��ϵ�����֤.Text = Nvl(rsTmp!��ϵ�����֤��)
    If mbytInState = E�޸� Then
        '���ҽ��,����Ժ����ʵҽ�����˿����޸�ҽ����
        txtPatiMCNO(0).Enabled = mlngOutModeMC > 0 Or Not IsNull(rsTmp!�������) And IsNull(rsTmp!����)
        
        txtPatiMCNO(0).Text = "" & rsTmp!ҽ���� '��󳤶��Զ��ضϳ����ַ�S
        txtPatiMCNO(0).Tag = txtPatiMCNO(0).Text
        If mlngOutModeMC > 0 Then txtPatiMCNO(1).Text = txtPatiMCNO(0).Text
    Else
        txtPatiMCNO(0).Text = Nvl(rsTmp!ҽ����)
    End If
    
    cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, Nvl(rsTmp!�Ա�))
    If cbo�Ա�.ListIndex = -1 And Not IsNull(rsTmp!�Ա�) Then
        cbo�Ա�.AddItem rsTmp!�Ա�, 0
        cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
    End If
       
    Call LoadOldData("" & rsTmp!����, txt����, cbo���䵥λ)
    mblnChange = False
    txt��������.Text = Format(IIf(IsNull(rsTmp!��������), "____-__-__", rsTmp!��������), "YYYY-MM-DD")
    mblnChange = True
    
    If rsTmp!��Ժ���� & "" = "" And rsTmp!��Ժ���� & "" <> "" Then
        txt��������.Tag = rsTmp!��Ժ���� & ""
    Else
        txt��������.Tag = "0"
    End If
    
    If Not IsNull(rsTmp!��������) Then
        If mbytInState <> 2 And mbytInState <> 1 Then txt����.Text = ReCalcOld(CDate(Format(rsTmp!��������, "YYYY-MM-DD HH:MM:SS")), cbo���䵥λ, lng����ID, , CDate(txt��������.Tag)) '�޸ĵ�ʱ��,���ݳ���������������
        
        If CDate(txt��������.Text) - CDate(rsTmp!��������) <> 0 Then
            mblnChange = False
            txt����ʱ��.Text = Format(rsTmp!��������, "HH:MM")
            mblnChange = True
        End If
    Else
        txt����ʱ��.Text = "__:__"
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
        
    mblnChange = False          '�޸ĺͲ鿴ʱ,���֤����������ڶ���
    txt���֤��.Text = Nvl(rsTmp!���֤��)
    mblnChange = True
    cboIDNumber.Enabled = txt���֤��.Text = ""
    '���ݲ�ͬ�鿴��ʽ��ȡ��ͬ�ķѱ�
    If mbytInState = E���� Then
        str�ѱ� = Nvl(rsTmp!�ѱ�)
    Else
        If mbytView = 1 Or mbytView = 2 Then
            str�ѱ� = Nvl(rsTmp!סԺ�ѱ�)
        Else
            str�ѱ� = Nvl(rsTmp!�ѱ�)
        End If
    End If
    
    cbo�ѱ�.ListIndex = GetCboIndex(cbo�ѱ�, str�ѱ�)
    If cbo�ѱ�.ListIndex = -1 And str�ѱ� <> "" Then
        cbo�ѱ�.AddItem str�ѱ�, 0
        cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
    End If
    
    cboҽ�Ƹ���.ListIndex = GetCboIndex(cboҽ�Ƹ���, Nvl(rsTmp!ҽ�Ƹ��ʽ))
    If cboҽ�Ƹ���.ListIndex = -1 And Not IsNull(rsTmp!ҽ�Ƹ��ʽ) Then
        cboҽ�Ƹ���.AddItem rsTmp!ҽ�Ƹ��ʽ, 0
        cboҽ�Ƹ���.ListIndex = cboҽ�Ƹ���.NewIndex
    End If
    
    cbo����.ListIndex = GetCboIndex(cbo����, Nvl(rsTmp!����))
    If cbo����.ListIndex = -1 And Not IsNull(rsTmp!����) Then
        cbo����.AddItem rsTmp!����, 0
        cbo����.ListIndex = cbo����.NewIndex
    End If
    
    cbo����.ListIndex = GetCboIndex(cbo����, Nvl(rsTmp!����))
    If cbo����.ListIndex = -1 And Not IsNull(rsTmp!����) Then
        cbo����.AddItem rsTmp!����, 0
        cbo����.ListIndex = cbo����.NewIndex
    End If
    
    txt����.Text = Nvl(rsTmp!����)
    
    cbo��������.ListIndex = GetCboIndex(cbo��������, Nvl(rsTmp!��������, "��ͨ����"))
    cbo��������.Enabled = InStr(mstrPrivs, "������������") > 0
    lblPatiType.Visible = True: cbo��������.Visible = True: lblPatiColor.Visible = True
    
    cboѧ��.ListIndex = GetCboIndex(cboѧ��, Nvl(rsTmp!ѧ��))
    If cboѧ��.ListIndex = -1 And Not IsNull(rsTmp!ѧ��) Then
        cboѧ��.AddItem rsTmp!ѧ��, 0
        cboѧ��.ListIndex = cboѧ��.NewIndex
    End If
    
    cbo����״��.ListIndex = GetCboIndex(cbo����״��, Nvl(rsTmp!����״��))
    If cbo����״��.ListIndex = -1 And Not IsNull(rsTmp!����״��) Then
        cbo����״��.AddItem rsTmp!����״��, 0
        cbo����״��.ListIndex = cbo����״��.NewIndex
    End If
    
    cboְҵ.ListIndex = GetCboIndex(cboְҵ, Nvl(rsTmp!ְҵ))
    If cboְҵ.ListIndex = -1 And Not IsNull(rsTmp!ְҵ) Then
        cboְҵ.AddItem rsTmp!ְҵ, 0
        cboְҵ.ListIndex = cboְҵ.NewIndex
    End If
    
    cbo���.ListIndex = GetCboIndex(cbo���, Nvl(rsTmp!���))
    If cbo���.ListIndex = -1 And Not IsNull(rsTmp!���) Then
        cbo���.AddItem rsTmp!���, 0
        cbo���.ListIndex = cbo���.NewIndex
    End If
         
    txt��ͥ�绰.Text = Nvl(rsTmp!��ͥ�绰)
    txt��ͥ��ַ�ʱ�.Text = Nvl(rsTmp!��ͥ��ַ�ʱ�)
    txt���ڵ�ַ�ʱ�.Text = Nvl(rsTmp!���ڵ�ַ�ʱ�)
    
    '������Ϣ�ݴ��ڴˣ����治��ʾ�����޸ı���ʱ��Ҫ
    txt��ϵ������.Tag = Nvl(rsTmp!������)
    txt��ϵ�˵绰.Tag = Nvl(rsTmp!������, 0)
    txt��ϵ�˵�ַ.Tag = Nvl(rsTmp!��������, 0)
    txt��ϵ������.Text = Nvl(rsTmp!��ϵ������)
    
    cbo��ϵ�˹�ϵ.ListIndex = GetCboIndex(cbo��ϵ�˹�ϵ, Nvl(rsTmp!��ϵ�˹�ϵ))
    If cbo��ϵ�˹�ϵ.ListIndex = -1 And Not IsNull(rsTmp!��ϵ�˹�ϵ) Then
        cbo��ϵ�˹�ϵ.AddItem rsTmp!��ϵ�˹�ϵ, 0
        cbo��ϵ�˹�ϵ.ListIndex = cbo��ϵ�˹�ϵ.NewIndex
    End If

    txt��ϵ�˵绰.Text = Nvl(rsTmp!��ϵ�˵绰)
    txt��ϵ�����֤.Text = Nvl(rsTmp!��ϵ�����֤��)
    txt������λ.Text = Nvl(rsTmp!������λ)
    txt������λ.Tag = Nvl(rsTmp!��ͬ��λID)
    txt��λ�绰.Text = Nvl(rsTmp!��λ�绰)
    txt��λ�ʱ�.Text = Nvl(rsTmp!��λ�ʱ�)
    txt��λ������.Text = Nvl(rsTmp!��λ������)
    txt��λ�ʺ�.Text = Nvl(rsTmp!��λ�ʺ�)
    txt����֤��.Text = "" & rsTmp!����֤��
    txtMobile.Text = Nvl(rsTmp!�ֻ���)
    '����27351 by lesfeng 2010-01-12
    If Nvl(rsTmp!�������, 0) = 0 Then
        txt��ע.Visible = False
        lbl��ע.Visible = False
        txt��ע.Tag = "0"
    Else
        mlng��ҳID = rsTmp!�������
        txt��ע.Tag = "1"
    End If
    txt��ע.Text = IIf(IsNull(rsTmp!��ע), "", rsTmp!��ע)
    
    If gbln���ýṹ����ַ Then
        Call ReadStructAddress(mlng����ID, mlng��ҳID, PatiAddress)
        txt�����ص�.Text = PatiAddress(E_IX_�����ص�).Value
        txt����.Text = PatiAddress(E_IX_����).Value
        txt��ͥ��ַ.Text = PatiAddress(E_IX_��סַ).Value
        txt���ڵ�ַ.Text = PatiAddress(E_IX_���ڵ�ַ).Value
        txt��ϵ�˵�ַ.Text = PatiAddress(E_IX_��ϵ�˵�ַ).Value
    Else
        txt�����ص�.Text = Nvl(rsTmp!�����ص�)
        txt����.Text = Nvl(rsTmp!����)
        txt��ͥ��ַ.Text = Nvl(rsTmp!��ͥ��ַ)
        txt���ڵ�ַ.Text = Nvl(rsTmp!���ڵ�ַ)
        txt��ϵ�˵�ַ.Text = Nvl(rsTmp!��ϵ�˵�ַ)
    End If
    '74299:
'    If IsNull(rsTmp!����ID) Then
'         lblPatiType.Visible = False: cbo��������.Visible = False: lblPatiColor.Visible = False
'    End If
    '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
    Call ReadPatPricture(lng����ID)
    '�����:56599
    Call Load�����������Ϣ(lng����ID)
    
    '�������޸Ĳ����������Ա𡢳������ڡ�����
    txtPatient.Locked = True
    txtPatient.BackColor = &H80000016
    cbo�Ա�.Locked = True
    cbo�Ա�.BackColor = txtPatient.BackColor
    txt��������.Enabled = False
    txt��������.BackColor = txtPatient.BackColor
    txt����ʱ��.Enabled = False
    txt����ʱ��.BackColor = txtPatient.BackColor
    txt����.Locked = True
    txt����.BackColor = txtPatient.BackColor
    cbo���䵥λ.Locked = True
    cbo���䵥λ.BackColor = txtPatient.BackColor
    ' ��ȡ�ӱ���Ϣ
    Set rsTmp = Get������Ϣ�ӱ�(lng����ID, "���֤��״̬")
    rsTmp.Filter = "��Ϣ��='���֤��״̬'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!��Ϣֵ) Then
            Call cbo.Locate(cboIDNumber, zlCommFun.GetNeedName(rsTmp!��Ϣֵ & ""))
        End If
    End If
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) <> "�й�" And Trim(txt���֤��.Text) = "" Then
        If Trim(zlCommFun.GetNeedName(cboIDNumber.Text)) = "" Then
            Set rsTmp = Get������Ϣ�ӱ�(lng����ID, "�⼮���֤��")
            rsTmp.Filter = "��Ϣ��='�⼮���֤��'"
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!��Ϣֵ) Then
                    txt���֤��.Text = "" & rsTmp!��Ϣֵ
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

Private Function ReadPatPricture(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTmp As String
    Dim rsData As Recordset
    
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    mstr�ɼ�ͼƬ = ""
    strSQL = "Select ����id,��Ƭ From ������Ƭ Where ����id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsData.BOF = False Then
        strTmp = zlDatabase.ReadPicture(rsData, "��Ƭ", strTmp)
        mstr�ɼ�ͼƬ = strTmp
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
    '����:�������֤ͼ��
    '����:������
    '����:2014-07-04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    mlngͼ����� = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SavePatPicture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���没����Ƭ
    '���:lng����ID - ����ID
    '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
    Dim strFile As String, strSQL As String
    
    On Error GoTo ErrHand
    Select Case mlngͼ�����
        Case 1 '�ļ�
            strFile = cmdialog.FileName
        Case 2 '�ɼ�
            strFile = mstr�ɼ�ͼƬ
            mstr�ɼ�ͼƬ = ""
        Case 4 '�������֤
            strFile = App.Path & "\SFZIMG.bmp"
            SavePicture imgPatient.Picture, strFile
    End Select
    If InStr(1, ",1,2,4,", "," & mlngͼ����� & ",") <> 0 Then
        gcnOracle.Execute "Delete From ������Ƭ Where ����id=" & lng����ID
        strSQL = "Select ����id,��Ƭ From ������Ƭ where ����id=" & lng����ID
        
        If strFile = "" Then Exit Sub
        rs.Open strSQL, gcnOracle, adOpenStatic, adLockOptimistic
        If rs.BOF Then
            If rs.EOF Then rs.AddNew
            rs("����id").Value = lng����ID
            rs("��Ƭ").Value = Null
            rs.Update
            
            If zlDatabase.SavePicture(strFile, rs, "��Ƭ") = False Then
                MsgBox "������Ƭʧ��,�ļ����ܱ�ɾ��!", vbInformation, gstrSysName
                Exit Sub
            End If
            rs.Close
        End If
    ElseIf mlngͼ����� = 3 Then
        strSQL = strSQL & "Zl_������Ƭ_Delete("
        strSQL = strSQL & lng����ID & ")"
        
        zlDatabase.ExecuteProcedure strSQL, "Zl_������Ƭ_Delete"
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddCardDataSQL(ByVal lng����ID As Long, ByVal dtCurdate As Date, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���￨���Ŵ���
    '���:lng����ID
    '����:���˺�
    '����:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt�������� As Byte, strNO As String, strPassWord As String, strSQL As String
    Dim strԭ���� As String, str���� As String, strCard As String, str�䶯ԭ�� As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str���㷽ʽ As String, strBrushCardNo As String
    Dim bln���ѿ� As Boolean, blnInRange As Boolean   '��Χ�ڵĿ�
    Dim lngIndex As Long, byt�䶯���� As Byte, lng����ID As Long
    
    strCard = UCase(txt����.Text): strICCard = IIf(mblnICCard, strCard, "")
    If Not ((strCard <> "" Or strICCard <> "") And (fraCard.Visible = True Or mbln���� = False)) Then Exit Sub
    '�����:56599
    mbln������󶨿� = True
    
    lng����ID = 0: blnInRange = True
    If mCurSendCard.blnOneCard And mCurSendCard.bln�ϸ���� Then blnInRange = mCurSendCard.lng����ID > 0
    
    If blnInRange And tabCardMode.SelectedItem.Key = "CardFee" Then
        blnInRange = True
        byt�������� = 0: byt�䶯���� = 1
    Else
        blnInRange = False
        byt�䶯���� = 11: byt�������� = 0
    End If
    str�䶯ԭ�� = "������Ϣ�ǼǷ���"
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
          strSQL = strSQL & "" & mCurSendCard.lng�����ID & ","
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
        '103980:���ϴ�,2017/1/19,���淢����������
        str���� = Trim(txt����.Text)
        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
        
        strNO = zlDatabase.GetNextNo(16)  'ҽ�ƿ�
        If chk����.Value = 0 Then
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        End If
        mCurCardPay.strNO = strNO
        mCurCardPay.lng����ID = lng����ID
        strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng�����ID, byt��������, strNO, lng����ID, 0, UserInfo.����ID, UserInfo.����ID, 0, _
         NeedName(cbo�ѱ�.Text), "", Trim(txtPatient.Text), NeedName(cbo�Ա�.Text), str����, _
        strCard, strPassWord, str�䶯ԭ��, IIf(mCurSendCard.bln��� = False, mCurSendCard.dblӦ�ս��, Val(txt����.Text)), Val(txt����.Text), IIf(cbo���㷽ʽ.Enabled, mCurCardPay.str���㷽ʽ, ""), _
        dtCurdate, mCurSendCard.lng����ID, mCurSendCard.rs����, strICCard, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, lng����ID)
    End If
    
    zlAddArray cllPro, strSQL
 End Sub
 Private Sub AddDepositSQL(ByVal cllPro As Collection, ByVal dtDate As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�����SQL
    '����:���˺�
    '����:2011-07-26 18:26:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, strSQL As String, i As Integer, lngԤ��ID As Long
    Dim dblMoney As Double
    
    If Not (IsNumeric(txtԤ����.Text) And fraDeposit.Visible) Then Exit Sub
     
    '����Ԥ�����¼
    strNO = zlDatabase.GetNextNo(11)
    lngԤ��ID = zlDatabase.GetNextId("����Ԥ����¼")
    mCurPrepay.strNO = strNO
    mCurPrepay.lngID = lngԤ��ID
    dblMoney = StrToNum(txtԤ����.Text)
    'Zl_����Ԥ����¼_Insert
    strSQL = "Zl_����Ԥ����¼_Insert("
    '  Id_In         ����Ԥ����¼.ID%Type,
    strSQL = strSQL & "" & lngԤ��ID & ","
    '  ���ݺ�_In     ����Ԥ����¼.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "" & IIf(mblnPrepayPrint, "'" & txtFact.Text & "'", "Null") & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & Val(txt����ID.Text) & ","
    '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
    strSQL = strSQL & "NULL,"
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "NULL,"
    '  ���_In       ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mCurPrepay.str���㷽ʽ & "',"
    '  �������_In   ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "'" & txt�������.Text & "',"
    '  �ɿλ_In   ����Ԥ����¼.�ɿλ%Type,
    strSQL = strSQL & "'" & Trim(txt�ɿλ.Text) & "',"
    '  ��λ������_In ����Ԥ����¼.��λ������%Type,
    strSQL = strSQL & "'" & Trim(txt������.Text) & "',"
    '  ��λ�ʺ�_In   ����Ԥ����¼.��λ�ʺ�%Type,
    strSQL = strSQL & "'" & Trim(txt�ʺ�.Text) & "',"
    '  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    strSQL = strSQL & "'��ԺԤ��',"
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(mlngԤ������ID = 0, "NULL", mlngԤ������ID) & ","
    '  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type := Null,
    strSQL = strSQL & "" & Val(Mid(tbDeposit.SelectedItem.Key, 2)) & ","
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.lngҽ�ƿ����ID = 0 Or mCurPrepay.bln���ѿ�, "NULL", mCurPrepay.lngҽ�ƿ����ID) & ","
   '  ���㿨���_in ����Ԥ����¼.���㿨���%type:=NULL,
    strSQL = strSQL & "" & IIf(mCurPrepay.lngҽ�ƿ����ID = 0 Or Not mCurPrepay.bln���ѿ�, "NULL", mCurPrepay.lngҽ�ƿ����ID) & ","
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.strˢ������ = "", "NULL", "'" & mCurPrepay.strˢ������ & "'") & ","
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  ������λ_In   ����Ԥ����¼.������λ%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null
    '108001:���ϴ���2017/5/8����ʽ��Ԥ��ʱ��Ϊ24Сʱ��
    strSQL = strSQL & "to_date('" & Format(dtDate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '   ��������_In Integer:=0 :0-������Ԥ��;1-��Ϊ���۵�
    strSQL = strSQL & "0 )"
    zlAddArray cllPro, strSQL
End Sub
Private Function SaveNewCard(strMCAccount As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˲��˱���
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 16:57:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPati As String, strSQLCard As String, strCard As String, strICCard As String
    Dim strNO As String, curDate As Date, strSQL As String
    Dim str�������� As String, str���� As String
    Dim strDepositNO As String, strDeposit As String
    Dim lngԤ��ID As Long, blnInRange As Boolean
    Dim blnTrans As Boolean, strOut As String, strErr As String
    Dim cllPro As Collection, cllUpdate As Collection, cllThreeInsert As Collection
    Dim i As Long
    Dim arrTmp As Variant
    
    '��ݵǼ�
    
    Set cllPro = New Collection
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    
    strCard = UCase(txt����.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    curDate = zlDatabase.Currentdate
    '�����:51071
    If mlngPatientID <> 0 Then
        strPati = "zl_������Ϣ_UPDATE(" & txt����ID.Text & "," & _
            IIf(Trim(txt�����.Text) <> "", Trim(txt�����.Text), "NULL") & "," & _
            IIf(Trim(txtסԺ��.Text) <> "", Trim(txtסԺ��.Text), "NULL") & "," & _
            "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cboҽ�Ƹ���.Text) & "','" & txtPatient.Text & "'," & _
            "'" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
            str�������� & "," & _
            "'" & txt�����ص�.Text & "','" & txt���֤��.Text & "','" & NeedName(cbo���.Text) & "'," & _
            "'" & NeedName(cboְҵ.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
            "'" & NeedName(cboѧ��.Text) & "','" & NeedName(cbo����״��.Text) & "','" & txt��ͥ��ַ.Text & "'," & _
            "'" & txt��ͥ�绰.Text & "','" & txt��ͥ��ַ�ʱ�.Text & "','" & txt��ϵ������.Text & "'," & _
            "'" & NeedName(cbo��ϵ�˹�ϵ.Text) & "','" & txt��ϵ�˵�ַ.Text & "','" & txt��ϵ�˵绰.Text & "'," & _
            Val(txt������λ.Tag) & ",'" & txt������λ.Text & "','" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "'," & _
            "'" & txt��λ������.Text & "','" & txt��λ�ʺ�.Text & "','" & txt��ϵ������.Tag & "'," & Val(txt��ϵ�˵绰.Tag) & "," & _
            IIf(mlng���� = 0, "NULL", mlng����) & "," & IIf(mbytInState = 0, 0, IIf(mbytView = 1 Or mbytView = 2, 1, 0)) & "," & _
            "'" & strMCAccount & "','" & NeedName(txt����.Text) & "'," & Val(txt��ϵ�˵�ַ.Tag) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
            "'" & Trim(txt����֤��.Text) & "','" & NeedName(cbo��������.Text) & "'," & _
            IIf(Trim(txt��ע.Text) = "", "Null", "'" & Trim(txt��ע.Text) & "'") & ",'" & NeedName(txt����.Text) & "','" & txt���ڵ�ַ.Text & "','" & txt���ڵ�ַ�ʱ�.Text & "'," & _
            "'" & txt��ϵ�����֤.Text & "',0,'" & Trim(txtMobile.Text) & "')"
        zlAddArray cllPro, strPati
    Else
        strPati = "zl_������Ϣ_INSERT(" & txt����ID.Text & "," & _
            IIf(Trim(txt�����.Text) <> "", Trim(txt�����.Text), "NULL") & "," & _
            "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cboҽ�Ƹ���.Text) & "','" & Trim(txtPatient.Text) & "'," & _
            "'" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
            str�������� & "," & _
            "'" & txt�����ص�.Text & "','" & txt���֤��.Text & "','" & NeedName(cbo���.Text) & "'," & _
            "'" & NeedName(cboְҵ.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
            "'" & NeedName(cboѧ��.Text) & "','" & NeedName(cbo����״��.Text) & "','" & txt��ͥ��ַ.Text & "'," & _
            "'" & txt��ͥ�绰.Text & "','" & txt��ͥ��ַ�ʱ�.Text & "','" & txt��ϵ������.Text & "'," & _
            "'" & NeedName(cbo��ϵ�˹�ϵ.Text) & "','" & txt��ϵ�˵�ַ.Text & "','" & txt��ϵ�˵绰.Text & "'," & _
            Val(txt������λ.Tag) & ",'" & txt������λ.Text & "','" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "'," & _
            "'" & txt��λ������.Text & "','" & txt��λ�ʺ�.Text & "',null,null," & _
            "NULL,To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & NeedName(txt����.Text) & "',null,'" & UserInfo.��� & "','" & UserInfo.���� & "','" & strMCAccount & "'," & _
            "'" & Trim(txt����֤��.Text) & "','" & NeedName(txt����.Text) & "','" & txt���ڵ�ַ.Text & "','" & txt���ڵ�ַ�ʱ�.Text & "'," & _
            "'" & txt��ϵ�����֤.Text & "','" & NeedName(cbo��������.Text) & "','" & Trim(txtMobile.Text) & "')"
        zlAddArray cllPro, strPati
    End If
    '�ӱ���Ϣ����
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",��ϵ�˸�����Ϣ,���֤��״̬,�⼮���֤��,", Split(arrTmp(i), ":")(0)) > 0 Then
                strPati = "Zl_������Ϣ�ӱ�_Update(" & txt����ID.Text & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
                zlAddArray cllPro, strPati
            End If
        Next
    End If
    '�����:53408
    If Trim(txt֧������.Text) <> "" And Trim(txt���֤��.Text) <> "" Then
        If zl�����֤(cllPro) = False Then Exit Function
    End If
    
    '�ṹ����ַ 89980
    If gbln���ýṹ����ַ Then
        Call CreateStructAddressSQL(CLng(txt����ID.Text), mlng��ҳID, cllPro, PatiAddress)
    End If
    
    'ҽ�ƿ�����
    '�����:51072
    If Len(Trim(txtPass.Text)) <= 0 And Len(Trim(txt����.Text)) > 0 Then 'û����������
        If zl_Get����Ĭ�Ϸ������� = False Then Exit Function
    End If
    
    Call AddCardDataSQL(Val(txt����ID.Text), curDate, cllPro) '����ҽ�ƿ�
    '�����:57326
    If mbln������󶨿� Then
        If Check��������(Val(txt����ID.Text), mCurSendCard.lng�����ID) = False Then
            txt����.Text = "": txtPass.Text = "": txtAudi.Text = "": txt����.Text = ""
            Exit Function
        End If
        '�����㷽ʽ��Ϣ�Ƿ�Ϸ�
        If (cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 8 Or cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 7) And mCurCardPay.lngҽ�ƿ����ID = 0 Then
            MsgBox "��ǰ�������㷽ʽ�����쳣���޷�ʹ�øý��㷽ʽ�������Ƿ�������Ӧ�豸�������Ա��ϵ!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    Call AddDepositSQL(cllPro, curDate)  '����Ԥ����
    '���Ԥ�����㷽ʽ��Ϣ�Ƿ�Ϸ�
    If IsNumeric(txtԤ����.Text) And fraDeposit.Visible Then
        If (cboԤ������.ItemData(cboԤ������.ListIndex) = 8 Or cboԤ������.ItemData(cboԤ������.ListIndex) = 7) And mCurPrepay.lngҽ�ƿ����ID = 0 Then
            MsgBox "��ǰԤ�����㷽ʽ�����쳣���޷�ʹ�øý��㷽ʽ�������Ƿ�������Ӧ�豸�������Ա��ϵ!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    '�����:56599
    If Val(Trim(txt����ID.Text)) > 0 Then Call Add�����������Ϣ(Val(Trim(txt����ID.Text)), cllPro)
    
    On Error GoTo errH
    
    Set cllUpdate = New Collection
    Set cllThreeInsert = New Collection
    
    Err = 0: On Error GoTo ErrHand:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    '֧������
    If Not zlInterfacePrayMoney(cllUpdate, cllThreeInsert) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    '������������
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    '74421,������Ƭ���
    Call SavePatPicture(Val(txt����ID.Text))
    '101160EMPI
    If Not EMPI_AddORUpdatePati(CLng(txt����ID.Text), mlng��ҳID, strErr) Then
        gcnOracle.RollbackTrans
        MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    '�����:56599
    'д��
    If mbln������󶨿� And mCurSendCard.bln�Ƿ�д�� Then WriteCard (Val(txt����ID.Text))
    
    Err = 0: On Error GoTo OthersCommit:
    zlExecuteProcedureArrAy cllThreeInsert, Me.Caption
    Call zlExcuteUploadSwap(txt����ID.Text, strOut, mobjICCard) '��������һ��ͨ�ϴ�����
    
    '73937:������,2013-07-03
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.PatiInfoSaveAfter(Val(txt����ID.Text))
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
'���ܣ��Ա��޸ĵĲ�����Ϣ��Ƭ���б���
    Dim strSQL As String
    Dim str�������� As String, str���� As String
    Dim blnTrans As Boolean
    Dim cllPro As New Collection  '�����:56599
    Dim arrTmp As Variant
    Dim i As Long
    Dim strErr As String
    
    On Error GoTo errH
    
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    
    '�����:51071
    '����27351 by lesfeng 2010-01-12
    strSQL = "zl_������Ϣ_UPDATE(" & txt����ID.Text & "," & _
        IIf(Trim(txt�����.Text) <> "", Trim(txt�����.Text), "NULL") & "," & _
        IIf(Trim(txtסԺ��.Text) <> "", Trim(txtסԺ��.Text), "NULL") & "," & _
        "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cboҽ�Ƹ���.Text) & "','" & txtPatient.Text & "'," & _
        "'" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
        str�������� & "," & _
        "'" & txt�����ص�.Text & "','" & txt���֤��.Text & "','" & NeedName(cbo���.Text) & "'," & _
        "'" & NeedName(cboְҵ.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
        "'" & NeedName(cboѧ��.Text) & "','" & NeedName(cbo����״��.Text) & "','" & txt��ͥ��ַ.Text & "'," & _
        "'" & txt��ͥ�绰.Text & "','" & txt��ͥ��ַ�ʱ�.Text & "','" & txt��ϵ������.Text & "'," & _
        "'" & NeedName(cbo��ϵ�˹�ϵ.Text) & "','" & txt��ϵ�˵�ַ.Text & "','" & txt��ϵ�˵绰.Text & "'," & _
        Val(txt������λ.Tag) & ",'" & txt������λ.Text & "','" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "'," & _
        "'" & txt��λ������.Text & "','" & txt��λ�ʺ�.Text & "','" & txt��ϵ������.Tag & "'," & Val(txt��ϵ�˵绰.Tag) & "," & _
        IIf(mlng���� = 0, "NULL", mlng����) & "," & IIf(mbytView = 1 Or mbytView = 2, 1, 0) & "," & _
        "'" & strMCAccount & "','" & NeedName(txt����.Text) & "'," & Val(txt��ϵ�˵�ַ.Tag) & ",'" & UserInfo.��� & "','" & _
        UserInfo.���� & "','" & Trim(txt����֤��.Text) & "','" & NeedName(cbo��������.Text) & "'," & _
        IIf(Trim(txt��ע.Text) = "", "Null", "'" & Trim(txt��ע.Text) & "'") & ",'" & NeedName(txt����.Text) & "','" & txt���ڵ�ַ.Text & "','" & txt���ڵ�ַ�ʱ�.Text & "'," & _
        "'" & Trim(txt��ϵ�����֤.Text) & "',0,'" & Trim(txtMobile.Text) & "')"
    '�ṹ����ַ
    If gbln���ýṹ����ַ Then
        Call CreateStructAddressSQL(CLng(txt����ID.Text), mlng��ҳID, cllPro, PatiAddress, 1)
    End If
    '������ҳ�ӱ���Ϣ����
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",��ϵ�˸�����Ϣ,���֤��״̬,�⼮���֤��,", Split(arrTmp(i), ":")(0)) > 0 Then
                zlAddArray cllPro, "Zl_������Ϣ�ӱ�_Update(" & txt����ID.Text & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
        Next
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    '74421
    Call SavePatPicture(Val(txt����ID.Text))
    '�����:56599
    If mlng����ID > 0 Then Call Add�����������Ϣ(mlng����ID, cllPro)
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
    '101160 EMPI
    If Not EMPI_AddORUpdatePati(CLng(txt����ID.Text), mlng��ҳID, strErr) Then
        gcnOracle.RollbackTrans
        MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    '����118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(IIf(mlng��ҳID <> 0, 2, 1), CLng(txt����ID.Text), mlng��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    '�����:56599
    'д��
    If mbln������󶨿� And mCurSendCard.bln�Ƿ񷢿� Then WriteCard (Val(txt����ID.Text))
    
    blnTrans = False
    
    '73937:������,2013-07-03
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.PatiInfoSaveAfter(Val(txt����ID.Text))
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
    If gstrIme <> "���Զ�����" Then Call OpenIme
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
'����:���ҽ�����Ƿ��Ѵ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 1 From ������Ϣ Where ҽ���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strMCNO)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����,�����ҽ�����Ѵ���!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatiMCNO_Validate(Index As Integer, Cancel As Boolean)
    txtPatiMCNO(Index).Text = UCase(Trim(txtPatiMCNO(Index).Text))
    '����28474 by lesfeng 2010-03-16 ȡ�������˳�ҽ���ż���֤ҽ��������
    If Index = 1 Then
        If txtPatiMCNO(1).Text <> txtPatiMCNO(0).Text Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
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

Private Sub txt���֤��_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
    If Trim(txt���֤��.Text) = "" Then
        cboIDNumber.Enabled = True
        If cboIDNumber.Enabled And cboIDNumber.Visible Then cboIDNumber.SetFocus
    Else
        cboIDNumber.Enabled = False
        cboIDNumber.ListIndex = -1
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt��֤����_GotFocus()
    Call zlControl.TxtSelAll(txt��֤����)
    Call OpenPassKeyboard(txt��֤����, False)
End Sub

Private Sub txt��֤����_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int������� = 1)
End Sub

Private Sub txt��֤����_LostFocus()
    Call ClosePassKeyboard(txt��֤����)
End Sub

Private Sub txtԤ����_GotFocus()
    If IsNumeric(txtԤ����.Text) Then
        txtԤ����.Text = StrToNum(txtԤ����.Text)
    Else
        txtԤ����.Text = ""
    End If
    txtԤ����.SelStart = 0: txtԤ����.SelLength = Len(txtԤ����.Text)
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
Private Sub txtԤ����_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then
        If InStr(txtԤ����.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        '65965:������,2013-09-24,����Ԥ����ʾǧλλ��ʽ
        If (txtԤ����.Text <> "" And txtԤ����.SelLength <> Len(Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;")) >= txtԤ����.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txtԤ����.SelLength > 0 And txtԤ����.SelLength <= txtԤ����.MaxLength Then
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf IsNumeric(txtԤ����.Text) Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '����ȡԤ����,ֱ������
        txtԤ����.Text = ""
        If fraCard.Visible Then
            txt����.SetFocus
        Else
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txtԤ����_LostFocus()
    '65965:������,2013-09-24,����Ԥ����ʾǧλλ��ʽ
    If IsNumeric(txtԤ����.Text) Then
        txtԤ����.Text = Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;")
    Else
        txtԤ����.Text = ""
    End If
    If txtԤ����.MaxLength > 12 Then txtԤ����.MaxLength = 12
End Sub

Private Sub txt�ʺ�_GotFocus()
    If IsNumeric(txtԤ����.Text) And txt�ʺ�.Text = "" Then
        txt�ʺ�.Text = txt��λ�ʺ�.Text
    End If
    zlControl.TxtSelAll txt�ʺ�
End Sub

Private Sub txt�ʺ�_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�ɿλ, KeyAscii
End Sub

Private Sub txt�ʺ�_LostFocus()
    Call zlCommFun.OpenIme
End Sub
Private Sub txt֧������_GotFocus()
    Call zlControl.TxtSelAll(txt֧������)
    Call OpenPassKeyboard(txt֧������, False)
End Sub

Private Sub txt֧������_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int������� = 1)
End Sub

Private Sub txt֧������_LostFocus()
    Call ClosePassKeyboard(txt֧������)
End Sub

Private Sub txtסԺ��_GotFocus()
    SelAll txtסԺ��
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtסԺ��_Validate(Cancel As Boolean)
    If Val(txtסԺ��.Text) = 0 And Val(txtסԺ��.Tag) <> 0 Then txtסԺ��.Text = txtסԺ��.Tag
End Sub
 
Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ԥ������
    '����:���˺�
    '����:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With tbDeposit
        mblnNotClick = True
        .Tabs.Clear
        If InStr(1, mstrPrivs, ";����Ԥ��;") > 0 Then
            .Tabs.Add(, "K1", "����Ԥ��(&M)").Selected = IIf(mbytPrepayType = 1, True, False)
        End If
        If InStr(1, mstrPrivs, ";סԺԤ��;") > 0 Then
            .Tabs.Add(, "K2", "סԺԤ��(&Z)").Selected = IIf(mbytPrepayType = 2, True, False)
        End If
         If .Tabs.Count > 0 And .SelectedItem Is Nothing Then
            .Tabs(0).Selected = True
         End If
         mblnNotClick = False
         Call tbDeposit_Click
         If .Tabs.Count = 0 Then
            fraDeposit.Visible = False
            Me.Height = Me.Height - fraDeposit.Height
            mPageHeight.���� = Me.Height
            If InStr(mstrPrivs, ";Ԥ���˿�;") = 0 Then cmdOperation(OPT.C0Ԥ����).Visible = False
         Else
            Call GetFact(True)
         End If
     End With
End Sub



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

Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    '���㷽ʽ:���ò�ѯ��ҽ�ƿ�����ʱ��һ��ֻ֧��Ԥ����,�����ڴ��յ����
    strSQL = _
        "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó��� ='Ԥ����'  And B.����=A.���㷽ʽ  " & _
        "           And Nvl(B.����,1) In(1,2,7,8)" & _
        " Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolPrepayPayMode = New Collection
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cboԤ������
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
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!����) & ",") = 0 Then
                .AddItem Nvl(rsTemp!����)
                mcolPrepayPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                'If mstrȱʡ���㷽ʽ = Nvl(rsTemp!����) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '���㷽ʽ���������豸���������˵Ľ��㷽ʽ����Ч
            rsTemp.Filter = "���� ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                                If InStr(1, varData(i), "|") <> 0 Then
                                        varTemp = Split(varData(i), "|")
                                        mcolPrepayPayMode.Add varTemp, "K" & j
                                        .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                                        'If mstrȱʡ���㷽ʽ = varTemp(1) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                                        j = j + 1
                                End If
                        End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cboԤ������.ListCount = 0 Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    strSQL = _
    "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
    " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
    " Where A.Ӧ�ó��� ='���￨'  And B.����=A.���㷽ʽ  " & _
    "           And Nvl(B.����,1) In(1,2,7,8)" & _
    " Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolCardPayMode = New Collection
    With cbo���㷽ʽ
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
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!����) & ",") = 0 Then
                .AddItem Nvl(rsTemp!����)
                mcolCardPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '���㷽ʽ���������豸���������˵Ľ��㷽ʽ����Ч
            rsTemp.Filter = "���� ='" & Split(varData(i), "|")(6) & "'"
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



Private Sub Local���㷽ʽ(ByVal lng�����ID As Long, Optional blnԤ�� As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��λ���㷽ʽ
    '����:���˺�
    '����:2011-07-26 15:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPayMoney As Collection, cboPay As ComboBox
    Dim i As Long
    If mblnNotClick Then Exit Sub
    
    If blnԤ�� Then
       Set cllPayMoney = mcolPrepayPayMode
        Set cboPay = cboԤ������
    Else
       Set cllPayMoney = mcolCardPayMode
        Set cboPay = cbo���㷽ʽ
    End If
    If cllPayMoney Is Nothing Then Exit Sub
    With cboPay
        If .ListIndex >= 0 Then
            If blnԤ�� Then
                If .ItemData(.ListIndex) >= 0 Then Exit Sub
            End If
        End If
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            ''��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
            If Val(cllPayMoney(i + 1)(3)) = lng�����ID Then
                .ListIndex = i: Exit For
            End If
        Next
        mblnNotClick = False
    End With
End Sub
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
        '58322
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        .ActiveConnection = Nothing
        If mCurPrepay.lngҽ�ƿ����ID <> 0 Then
            .AddNew
            !�շ���� = "Ԥ��"
            !��� = StrToNum(txtԤ����.Text)
            .Update
        End If
        If mCurCardPay.lngҽ�ƿ����ID <> 0 And cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
            .AddNew
            !�շ���� = mCurSendCard.rs����!�շ����
            !��� = StrToNum(txt����.Text)
            .Update
        End If
    End With
    zlGetClassMoney = True
End Function

Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str���� As String
    Dim dblMoney As Double, bln�������� As Boolean
    Dim dblThreeMoney As Double, tyCurThreePay As Ty_PayMoney
    Dim blnTemp As Boolean
    
    On Error GoTo errHandle
    '58322
    dblMoney = 0: dblThreeMoney = 0
    If cboԤ������.Visible Then
        If cboԤ������.ListIndex >= 0 Then
            bln�������� = cboԤ������.ItemData(cboԤ������.ListIndex) = -1
            If bln�������� Then dblThreeMoney = dblThreeMoney + StrToNum(txtԤ����.Text)
        End If
        dblMoney = dblMoney + StrToNum(txtԤ����.Text)
    End If
    
    If cbo���㷽ʽ.Visible And cbo���㷽ʽ.Enabled Then
        If cbo���㷽ʽ.ListIndex >= 0 Then
            blnTemp = cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = -1
            If blnTemp Then dblThreeMoney = dblThreeMoney + StrToNum(txt����.Text)
            If blnTemp Then bln�������� = bln�������� Or blnTemp
        End If
        dblMoney = dblMoney + StrToNum(txt����.Text)
    End If
    If Not bln�������� Then CheckBrushCard = True: Exit Function
    '���⣺126188��2018-06-19,��Ժ�Ǽ�ʧ��
    If dblThreeMoney = 0 Then CheckBrushCard = True: Exit Function
    If mCurPrepay.lngҽ�ƿ����ID <> 0 Then
       tyCurThreePay = mCurPrepay
    Else
       tyCurThreePay = mCurCardPay
    End If
    
    
    If (mCurCardPay.lngҽ�ƿ����ID <> mCurCardPay.lngҽ�ƿ����ID Or _
        mCurPrepay.bln���ѿ� <> mCurCardPay.bln���ѿ�) _
        And mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurPrepay.lngҽ�ƿ����ID <> 0 Then
        MsgBox "����ͬʱʹ�����ֲ�ͬ����֧����ʽ,���ܼ���?", vbOKOnly + vbInformation, gstrSysName
        If cboԤ������.Enabled And cboԤ������.Visible Then cboԤ������.SetFocus: Exit Function
        If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus
        Exit Function
    End If
    Call zlGetClassMoney(rsMoney)
    
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
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
   '58322
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, tyCurThreePay.lngҽ�ƿ����ID, tyCurThreePay.bln���ѿ�, _
    txtPatient.Text, NeedName(cbo�Ա�.Text), str����, dblThreeMoney, tyCurThreePay.strˢ������, tyCurThreePay.strˢ������, False, True, False) = False Then Exit Function
    
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, tyCurThreePay.lngҽ�ƿ����ID, _
        tyCurThreePay.bln���ѿ�, tyCurThreePay.strˢ������, dblThreeMoney, "", "") = False Then Exit Function
    mCurCardPay.strˢ������ = tyCurThreePay.strˢ������
    mCurCardPay.strˢ������ = tyCurThreePay.strˢ������
    mCurPrepay.strˢ������ = tyCurThreePay.strˢ������
    mCurPrepay.strˢ������ = tyCurThreePay.strˢ������
    
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
    If mCurCardPay.lngҽ�ƿ����ID = 0 And mCurPrepay.lngҽ�ƿ����ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cboԤ������.ItemData(cboԤ������.ListIndex) <> -1 _
        And cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
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
    dblMoney = 0
    If mCurCardPay.lngҽ�ƿ����ID <> 0 And cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
        dblMoney = Val(txt����.Text)
    End If
    If mCurPrepay.lngҽ�ƿ����ID <> 0 And cboԤ������.Visible Then
        dblMoney = dblMoney + Val(StrToNum(txtԤ����.Text))
    End If
    '���⣺126188��2018-06-19,��Ժ�Ǽ�ʧ��
    If dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, mCurCardPay.lng����ID, mCurPrepay.strNO, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '����������������
     If mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.lng����ID <> 0 And cbo���㷽ʽ.Visible Then
     
        If Not mCurCardPay.bln���ѿ� Then
            Call zlAddUpdateSwapSQL(False, mCurCardPay.lng����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mCurCardPay.lng����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    If mCurPrepay.lngҽ�ƿ����ID <> 0 And cboԤ������.Visible And Val(StrToNum(txtԤ����.Text)) <> 0 Then
        Call zlAddUpdateSwapSQL(True, mCurPrepay.lngID, mCurPrepay.lngҽ�ƿ����ID, mCurPrepay.bln���ѿ�, mCurPrepay.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
        Call zlAddThreeSwapSQLToCollection(True, mCurPrepay.lngID, mCurPrepay.lngҽ�ƿ����ID, mCurPrepay.bln���ѿ�, mCurPrepay.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtAudi_Validate(Cancel As Boolean)
    Select Case mCurSendCard.int���볤������
        Case 0
        Case 1
            If Len(txtAudi.Text) <> mCurSendCard.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "ȷ�������������" & mCurSendCard.int���볤�� & "λ", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtAudi.Text) < Abs(mCurSendCard.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "ȷ�����������" & Abs(mCurSendCard.int���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Exit Sub
             End If
        End Select
End Sub

Private Sub txtPass_Validate(Cancel As Boolean)
   Select Case mCurSendCard.int���볤������
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "�����������" & mCurSendCard.int���볤�� & "λ", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "�����������" & Abs(mCurSendCard.int���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        End Select
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
    Dim arr() As String
    arr = zl_Getҽ�ƿ�����(mCurSendCard.lng�����ID)
    If Val(arr(2)) = 0 Then '������
        Select Case Val(arr(1))
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
    ElseIf Val(arr(2)) = 1 Then 'ȱʡ���֤��Nλ
        If Len(Trim(txt���֤��.Text)) > 0 Or Len(Trim(txt��ϵ�����֤.Text)) > 0 Then '���������֤����ϵ�����֤��
            If Len(Trim(txt���֤��.Text)) > 0 Then '�����֤���������֤
                   txtPass.Text = Right(Trim(txt���֤��.Text), Val(arr(0)))
            Else '������ô��������֤��Ϊ����
                   txtPass.Text = Right(Trim(txt��ϵ�����֤.Text), Val(arr(0)))
            End If
        Else '���֤����ϵ�����֤��û����
            Select Case Val(arr(1))
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


Public Sub Show�󶨿ؼ�(blnShow As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ���ʾ������
    '���:blnShow �Ƿ���ʾ������
    '����:����
    '����:2012-09-04 15:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    lbl֧������.Enabled = blnShow: txt֧������.Enabled = blnShow
    lbl��֤����.Enabled = blnShow: txt��֤����.Enabled = blnShow
    If blnShow = False Then
        txt֧������.Text = "": txt��֤����.Text = ""
    End If
    
End Sub
Private Function zl�����֤(colPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�󶨶������֤
    '���:blnShow �Ƿ���ʾ������
    '����:����
    '����:2012-09-04 15:53:14
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txt֧������.Text) <> Trim(txt��֤����.Text) Then
        MsgBox "������������벻һ��,����������", vbOKOnly + vbInformation, gstrSysName
        txt֧������.Text = "": txt��֤����.Text = ""
        If txt֧������.Visible = True Then txt֧������.SetFocus
        Exit Function
    End If
    If Trim(txt֧������.Text) <> "" Then
       If �Ƿ��Ѿ�ǩԼ(Trim(txt���֤��.Text)) Then
             MsgBox "���֤����Ϊ:" & txt���֤��.Text & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
             txt֧������.Text = "": txt��֤����.Text = ""
             If txt֧������.Visible = True Then txt֧������.SetFocus
             Exit Function
       End If
    End If
    AddSQL�󶨿� Trim(txt����ID.Text), Getҽ�ƿ����ID("�������֤"), Trim(txt���֤��.Text), zlCommFun.zlStringEncode(Trim(txt֧������.Text)), zlDatabase.Currentdate, False, colPro
    
    zl�����֤ = True
End Function
Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ҳ�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
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
        
    Set objItem = tbcPage.InsertItem(1, "����", PicBaseInfo.hWnd, 0)
    objItem.Tag = mPageHeight.����
    
    Set objItem = tbcPage.InsertItem(2, "��������", PicHealth.hWnd, 0)
    objItem.Tag = mPageHeight.��������
    If mlngPlugInHwnd <> 0 Then
        picTmp.Visible = True
        Set objItem = tbcPage.InsertItem(3, "������Ϣ", picTmp.hWnd, 0)
        objItem.Tag = mPageHeight.������Ϣ
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
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

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
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               'Ϊ��֧��zl9PrintMode
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
                    'Ϊ��֧��zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  'Ϊ��֧��zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
                End If
            End If
        Next
        
        '�̶������־���
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
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
    '����:��ʼ��ComBox�ؼ�
    '����:56599
    '����:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '66743:������,2013-11-25,Ѫ����RHĬ��ֵ������
    'ComboBox cboBloodType, C_Ѫ��
    Call ReadDict("Ѫ��", cboBloodType)
    ComboBox cboBH, C_BH
    If cboBH.ListCount <> 0 Then cboBH.ListIndex = -1
End Sub
Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo ErrHand
    
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ����ϵ Order by ����"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "����ϵ")
    With rsTemp
        Do While Not rsTemp.EOF
            strTmp = strTmp & "|" & Nvl(rsTemp!����)
        rsTemp.MoveNext
        Loop
    End With
    If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
    
    With vsLinkMan
        '��ʼ���б�����
        SetColumHeader vsLinkMan, C_LinkManColumHeader
        .Editable = IIf(mbytInState = E����, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        If strTmp <> "" Then .ColComboList(.ColIndex("��ϵ�˹�ϵ")) = strTmp
    End With
    
    With vsOtherInfo
        '������ͷ
        SetColumHeader vsOtherInfo, C_OtherInfoColumHeader
        .Editable = IIf(mbytInState = E����, flexEDNone, flexEDKbdMouse)
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
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsDrug
        '��ʼ���б�����
        SetColumHeader vsDrug, C_ColumHeader
        .Editable = IIf(mbytInState = E����, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        .ColComboList(0) = "..."
    End With
End Sub

Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
        '��ʼ���б�����
        SetColumHeader vsInoculate, C_InoculateHeader
         vsInoculate.Editable = IIf(mbytInState = E����, flexEDNone, flexEDKbdMouse)
        '����ѡ��ť
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
            strFilter = " And zlspellcode(A.����) like [1]"
            strInput = UCase(strInput)
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strFilter = " And A.���� like [1]"
        Else
            strFilter = " And A.���� like [1]"
        End If
    End If
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select Distinct A.ID,A.����," & _
        " A.����,zlspellcode(A.����) ����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & strFilter

    '��ȡ��ǰ�������ֵ
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    strInput = gstrLike & Trim(strInput) & "%"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ��", False, "����ҩ��ѡ����", "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ��", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True, strInput)

    If Not rsTemp Is Nothing And blnCancel = False Then
        If rsTemp.RecordCount > 0 Then
            vsDrug.EditText = Nvl(rsTemp!����)
            vsDrug.TextMatrix(Row, Col) = Nvl(rsTemp!����)
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
    '�����:56599
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
    '�����:56599
    If Col = 1 Then  '������Ӧ�б༭ʱ���ж��Ƿ�����������100
        With vsDrug
           If LenB(StrConv(.TextMatrix(Row, Col), vbFromUnicode)) > 100 Then
                MsgBox "������Ӧ�����ַ���������ַ���100,������ַ������Զ��س���", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = StrConv(MidB(StrConv(.TextMatrix(Row, Col), vbFromUnicode), 1, 100), vbUnicode)
           End If
        End With
    End If
End Sub

Private Sub vsDrug_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '�����:56599
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandl:
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select ID,nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����,NULL ����," & _
        " NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ��" & _
        " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����," & _
        " A.����,zlspellcode(A.����) ����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"

    '��ȡ��ǰ�������ֵ
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "����ҩ��ѡ����", "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ��", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True)

    If Not rsTemp Is Nothing And blnCancel = False Then
        vsDrug.TextMatrix(Row, Col) = rsTemp!����
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
    '�����:56599
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
    Else
        If cmdMedicalWarning.Enabled And cmdMedicalWarning.Visible Then cmdMedicalWarning.SetFocus: Exit Sub
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
    If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
End Sub
Private Sub SetDrugAllergy(str����ҩ�� As String, str������Ӧ As String, Optional lng����ID = 0, Optional ByVal byt��Ϣ����ģʽ As Byte)
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
                    If mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 Then Exit Sub
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
    '68192:������,2013-12-02
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
                .TextMatrix(.Rows - 1, j) = Format(str��������, "YYYY-MM-DD")
                .TextMatrix(.Rows - 1, j + 1) = str��������
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
        
    End With
End Sub

Private Sub SetLinkInfo(str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, Optional ByVal byt��Ϣ����ģʽ As Byte)
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
                    If mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 Then Exit Sub
                    .TextMatrix(i, 1) = str��ϵ: .TextMatrix(i, 3) = str�绰
                    If i = 1 Then
                        txt��ϵ�����֤.Text = str���֤��
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
            txt��ϵ�����֤.Text = str���֤��
            txt��ϵ������.Text = str����
            For j = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                If NeedName(cbo��ϵ�˹�ϵ.List(j)) = str��ϵ Then cbo��ϵ�˹�ϵ.ListIndex = j
            Next
            txt��ϵ�˵绰.Text = str�绰
        End If
        .Rows = .Rows + 1
    End With
End Sub

Private Sub SetOtherInfo(str��Ϣ�� As String, str��Ϣֵ As String, Optional ByVal byt��Ϣ����ģʽ As Byte)
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
                        If mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 Then Exit Sub
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
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
        
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
        SetInoculate Format(Nvl(rs���߼�¼!����ʱ��), "YYYY-MM-DD"), Nvl(rs���߼�¼!��������)
        rs���߼�¼.MoveNext
    Wend
    'Ѫ��
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='Ѫ��'"
    Set rsABOѪ�� = zlDatabase.OpenSQLRecord(strSQL, "ABOѪ��", lng����ID)
    While rsABOѪ��.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            '76314,���ϴ���2014-08-06����ȷ��ȡ������Ϣ
            If NeedName(cboBloodType.List(i)) = NeedName(Nvl(rsABOѪ��!��Ϣֵ)) Then cboBloodType.ListIndex = i
        Next
        rsABOѪ��.MoveNext
    Wend
    'RH
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='RH'"
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
    txtMedicalWarning.Tag = txtMedicalWarning.Text
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
        txt��ϵ�����֤.Text = Nvl(rs������Ϣ!��ϵ�����֤��)
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
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� Not in ('Ѫ��','RH','ҽѧ��ʾ','����ҽѧ��ʾ','���֤��״̬','�⼮���֤��') And ��Ϣ�� Not like '��ϵ��%'"
    Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ϵ��������Ϣ", lng����ID)
    While rs������Ϣ.EOF = False
            SetOtherInfo rs������Ϣ!��Ϣ��, rs������Ϣ!��Ϣֵ
        rs������Ϣ.MoveNext
    Wend

    Exit Sub
ErrHandl:
    If ErrCenter() = 1 Then
       Resume
    End If
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
    '76314,���ϴ���2014-08-06����ȷ��ȡ������Ϣ
    strSQL = strSQL & "'" & NeedName(cboBloodType.Text) & "',"
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
     If Not mdicҽ�ƿ����� Is Nothing And txt����.Text <> "" Then
        For Each varKey In mdicҽ�ƿ�����.Keys
            strSQL = "Zl_����ҽ�ƿ�����_Update("
            strSQL = strSQL & lng����ID & ","
            strSQL = strSQL & mCurSendCard.lng�����ID & ","
            strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdicҽ�ƿ�����(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
End Sub

Private Function CheckPatiCard() As Boolean
'���ܣ���鲡�˽�����Ƭ¼��������Ƿ�Ϸ�
'63246:������,2013-07-03
    Dim intLen As Integer, i As Integer, j As Integer
    
    intLen = 100
    If LenB(StrConv(txtMedicalWarning.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "ҽѧ��ʾֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbInformation, gstrSysName
        If txtMedicalWarning.Enabled And txtMedicalWarning.Visible Then txtMedicalWarning.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtOtherWaring.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "����ҽѧ��ʾֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbInformation, gstrSysName
        If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
        Exit Function
    End If
    
    mblnCheckPatiCard = True
    '����ҩ��
    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 60
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "����ҩ��ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���1��", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "������Ӧֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���2��", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '������Ϣ
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    If Not IsDate(.TextMatrix(i, 0)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "����ʱ�䲻����Ч�����ڸ�ʽ��" & vbCrLf & "������:��" & i & "�С���1��", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��������ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���2��", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
                If .TextMatrix(i, 3) <> "" Then
                    If Not IsDate(.TextMatrix(i, 2)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "����ʱ�䲻����Ч�����ڸ�ʽ��" & vbCrLf & "������:��" & i & "�С���3��", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��������ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���4��", vbInformation, gstrSysName
                        Call .Select(i, 3, i, 3)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '��ϵ�˵�ַ
    With vsLinkMan
        intLen = 100
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    For j = 0 To .Cols - 1
                        If .ColIndex("��ϵ������") = j Then
                            intLen = 64
                        ElseIf .ColIndex("��ϵ�����֤��") = j Then
                            intLen = 18
                        ElseIf .ColIndex("��ϵ�˵绰") = j Then
                            intLen = 20
                        Else
                            intLen = 100
                        End If
                        If LenB(StrConv(.TextMatrix(i, j), vbFromUnicode)) > intLen Then
                            tbcPage.Item(1).Selected = True
                            MsgBox .TextMatrix(0, j) & "ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "��", vbInformation, gstrSysName
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
    
    '������Ϣ
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 20
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��Ϣ��ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���1��", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��Ϣֵֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���2��", vbInformation, gstrSysName
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
                        MsgBox "��Ϣ��ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���3��", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��Ϣֵֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���4��", vbInformation, gstrSysName
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
    Dim byt��Ϣ����ģʽ As Byte
    On Error GoTo errHandle

    If strPatiXML = "" Then Exit Function
    
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ��Ϣ����ģʽ Integer 1 '0-ǿ�Ƹ��£�1-�������˲����£�2-����������Ϣ��ȱ
    Call zlXML_GetNodeValue("��Ϣ����ģʽ", , strValue)
    byt��Ϣ����ģʽ = 0
    byt��Ϣ����ģʽ = Val(strValue)
    If byt��Ϣ����ģʽ = 1 And mlng����ID <> 0 Then LoadPati = True: Exit Function
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue)
    '    ����    Varchar2    64
    Call zlXML_GetNodeValue("����", , strValue)
    '1-�������
    '2-�²���
    '3-�ϲ�����ϢΪ�յ������ȱ
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txtPatient.Text = "") Then txtPatient.Text = strValue
    '    �Ա�    Varchar2    4
    Call zlXML_GetNodeValue("�Ա�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo�Ա�.Text = "") Then
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
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt����.Text = "") Then
        If strValue <> "" Then
            Call LoadOldData(strValue, txt����, cbo���䵥λ)
        End If
    End If
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("��������", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��������.Text = "") Then
        mblnChange = False
        txt��������.Text = Format(IIf(strValue = "", "____-__-__", strValue), "YYYY-MM-DD")
        mblnChange = True
        If strValue <> "" Then
            txt����.Text = ReCalcOld(CDate(Format(strValue, "YYYY-MM-DD HH:MM:SS")), cbo���䵥λ, , , CDate(txt��������.Tag))   '�޸ĵ�ʱ��,���ݳ���������������
            If CDate(txt��������.Text) - CDate(strValue) <> 0 Then
                mblnChange = False
                txt����ʱ��.Text = Format(strValue, "HH:MM")
                mblnChange = True
            End If
        Else
            txt����ʱ��.Text = "__:__"
            mblnChange = False
            Call ReCalcBirthDay
            mblnChange = True
        End If
    End If
    '    �����ص�    Varchar2    50
    Call zlXML_GetNodeValue("�����ص�", , strValue)
    '    ���֤��    VARCHAR2    18
    Call zlXML_GetNodeValue("���֤��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt���֤��.Text = "") Then
        If strValue <> "" Then txt���֤��.Text = strValue
    End If
    '    ����֤��    Varchar2    20
    Call zlXML_GetNodeValue("����֤��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt����֤��.Text = "") Then
        If strValue <> "" Then txt����֤��.Text = strValue
    End If
    '    ְҵ    Varchar2    80
    Call zlXML_GetNodeValue("ְҵ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboְҵ.Text = "") Then
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
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo����.Text = "") Then
        cbo����.ListIndex = GetCboIndex(cbo����, strValue)
        If cbo����.ListIndex = -1 And strValue <> "" Then
            cbo����.AddItem strValue, 0
            cbo����.ListIndex = cbo����.NewIndex
        End If
    End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo����.Text = "") Then
        cbo����.ListIndex = GetCboIndex(cbo����, strValue)
        If cbo����.ListIndex = -1 And strValue <> "" Then
            cbo����.AddItem strValue, 0
            cbo����.ListIndex = cbo����.NewIndex
        End If
    End If
    '    ѧ��    Varchar2    10
    Call zlXML_GetNodeValue("ѧ��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboѧ��.Text = "") Then
        cboѧ��.ListIndex = GetCboIndex(cboѧ��, strValue)
        If cboѧ��.ListIndex = -1 And strValue <> "" Then
            cboѧ��.AddItem strValue, 0
            cboѧ��.ListIndex = cboѧ��.NewIndex
        End If
    End If
    '    ����״��    Varchar2    4
    Call zlXML_GetNodeValue("����״��", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cbo����״��.Text = "") Then
        cbo����״��.ListIndex = GetCboIndex(cbo����״��, strValue)
        If cbo����״��.ListIndex = -1 And strValue <> "" Then
            cbo����״��.AddItem strValue, 0
            cbo����״��.ListIndex = cbo����״��.NewIndex
        End If
    End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt����.Text = "") Then txt����.Text = strValue
    '    ��ͥ��ַ    Varchar2    50
    Call zlXML_GetNodeValue("��ͥ��ַ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ͥ��ַ.Text = "") Then
        txt��ͥ��ַ.Text = strValue
        If gbln���ýṹ����ַ Then PatiAddress(E_IX_��סַ).Value = strValue
    End If
    '    ��ͥ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��ͥ�绰", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ͥ�绰.Text = "") Then txt��ͥ�绰.Text = strValue
    '    ��ͥ��ַ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��ͥ��ַ�ʱ�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ͥ��ַ�ʱ�.Text = "") Then txt��ͥ��ַ�ʱ�.Text = strValue
    '    ���ڵ�ַ    Varchar2    50
    Call zlXML_GetNodeValue("���ڵ�ַ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt���ڵ�ַ.Text = "") Then txt���ڵ�ַ.Text = strValue
    '    ���ڵ�ַ    Varchar2    50
    Call zlXML_GetNodeValue("���ڵ�ַ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt���ڵ�ַ.Text = "") Then
        txt���ڵ�ַ.Text = strValue
        If gbln���ýṹ����ַ Then PatiAddress(E_IX_���ڵ�ַ).Value = strValue
    End If
    '    �໤��  Varchar2    64
    Call zlXML_GetNodeValue("�໤��", , strValue)
   'txt�໤��.Text = strValue
    '    ������λ    Varchar2    100
    Call zlXML_GetNodeValue("������λ", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt������λ.Text = "") Then
        txt������λ.Text = strValue
        lbl������λ.Tag = ""
    End If
    '    ��λ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��λ�绰", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ�绰.Text = "") Then txt��λ�绰.Text = strValue
    '�ֻ���   Varchar2    20
    Call zlXML_GetNodeValue("�ֻ���", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txtMobile.Text = "") Then txtMobile.Text = strValue
    '    ��λ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��λ�ʱ�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ�ʱ�.Text = "") Then txt��λ�ʱ�.Text = strValue
    '    ��λ������  Varchar2    50
    Call zlXML_GetNodeValue("��λ������", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ������.Text = "") Then txt��λ������.Text = strValue
    '    ��λ�ʺ�    Varchar2    20
    Call zlXML_GetNodeValue("��λ�ʺ�", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��λ�ʺ�.Text = "") Then txt��λ�ʺ�.Text = strValue
    '�����:56599
    '�������
    Call zlXML_GetRows("ҩ������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("ҩ������", i, str����ҩ��)
        Call zlXML_GetNodeValue("ҩ�ﷴӦ", i, str������Ӧ)
        SetDrugAllergy str����ҩ��, str������Ӧ, , byt��Ϣ����ģʽ
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
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboBloodType.Text = "") Then
        If strABOѪ�� <> "" Then
            For i = 0 To cboBloodType.ListCount - 1
                '76314,���ϴ���2014-08-06����ȷ��ȡ������Ϣ
                If NeedName(cboBloodType.List(i)) = NeedName(strABOѪ��) Then cboBloodType.ListIndex = i
            Next
        End If
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And cboBH.Text = "") Then
        If strValue <> "" Then
            For i = 0 To cboBH.ListCount - 1
                If cboBH.List(i) = strValue Then cboBH.ListIndex = i
            Next
        End If
    End If
    'ҽѧ��ʾ
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("�ٴ�������Ϣ")
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txtMedicalWarning.Text = "") Then
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
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txtOtherWaring.Text = "") Then
        If strValue <> "" Then txtOtherWaring.Text = strValue
    End If
    '��ϵ��Ϣ
    '    ��ϵ�˵�ַ  Varchar2    50
    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , str��ַ)
    If byt��Ϣ����ģʽ = 0 Or mlng����ID = 0 Or (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2 And txt��ϵ�˵�ַ.Text = "") Then
        txt��ϵ�˵�ַ.Text = str��ַ
        If gbln���ýṹ����ַ Then PatiAddress(E_IX_��ϵ�˵�ַ).Value = str��ַ
    End If
     '    ��ϵ������  Varchar2    64
    Call zlXML_GetNodeValue("��ϵ������", , str����)
    '    ��ϵ�˹�ϵ  Varchar2    30
    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , str��ϵ)
    '    ��ϵ�˵绰  Varchar2    20
    Call zlXML_GetNodeValue("��ϵ�˵绰", , str�绰)
    '    ��ϵ�����֤ Varchar2   20
    Call zlXML_GetNodeValue("��ϵ�����֤��", , str���֤��)
    SetLinkInfo str����, str��ϵ, str�绰, str���֤��, byt��Ϣ����ģʽ
    
    Call zlXML_GetRows("��ϵ��Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("��ϵ��Ϣ", "����", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "����", i, j, str����)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "��ϵ", i, j, str��ϵ)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "�绰", i, j, str�绰)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "���֤��", i, j, str���֤��)
                SetLinkInfo str����, str��ϵ, str�绰, str���֤��, byt��Ϣ����ģʽ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '������Ϣ
    '�����������
    Call zlXML_GetNodeValue("�����������", , strValue)
    SetOtherInfo "�����������", strValue, byt��Ϣ����ģʽ
    
    '��ũ��֤��
    Call zlXML_GetNodeValue("��ũ��֤��", , strValue)
    SetOtherInfo "��ũ��֤��", strValue, byt��Ϣ����ģʽ

    '����֤��
    Call zlXML_GetRows("����֤��", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("����֤��", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣֵ", i, j, str��Ϣֵ)
                SetOtherInfo str��Ϣ��, str��Ϣֵ, byt��Ϣ����ģʽ
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
                SetOtherInfo str��Ϣ��, str��Ϣֵ, byt��Ϣ����ģʽ
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
                    If Not (mlng����ID <> 0 And byt��Ϣ����ģʽ = 2) Then mdicҽ�ƿ�����.Item(str��Ϣ��) = str��Ϣֵ
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
Public Function GetCboIndex(cbo As ComboBox, strFind As String, _
    Optional blnKeep As Boolean, _
    Optional blnLike As Boolean, Optional strSplit As String = "-") As Long
'���ܣ����ַ�����ComboBox�в�������
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '�Ⱦ�ȷ����
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), strSplit) > 0 Then
            If NeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '���ģ������
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Sub Clear��������()
    '---------------------------------------------------------------------------------------------------------------------------------------------
'����:�жϵ�ǰ�Ƿ�Ϊ�������� (���Ƿ����������ǰ󶨿�����)
'���:
'����:56599
'����:2012-12-25 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    'Ѫ��
    Call SetCboDefault(cboBloodType)
    'RH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
    'ҽѧ��ʾ
    txtMedicalWarning.Text = ""
    '����ҽѧ��ʾ
    txtOtherWaring.Text = ""
    '��ϵ����Ϣ
    With vsLinkMan
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '�������
    With vsInoculate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '������Ϣ
    With vsOtherInfo
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '������Ӧ
    With vsDrug
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
    End With
End Sub
Private Function SetCreateCardObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ƿ�����
    '����:����
    '����:2012-12-17 11:06:41
    '�����:56599
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
    '����:����ҽ�ƿ�����
    '����:���ϴ�
    '����:2016/6/21 11:57:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not mobjSquare Is Nothing Then zlCreateSquare = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjSquare = CreateObject("zl9CardSquare.clsCardsquare")
    If Err <> 0 Then Err = 0: Exit Function
    Call mobjSquare.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend)
    '��ʼ�������ɹ�,����Ϊ�����ڴ���
    zlCreateSquare = True
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
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    WriteCard = mobjSquare.zlBandCardArfter(Me, mlngModul, mCurSendCard.lng�����ID, lng����ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
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
        Select Case mCurSendCard.lng��������
            Case 0 '������
                Check�������� = True
            Case 1 'ͬһ������ֻ����һ�ſ�
                MsgBox "�ò����Ѿ�����" & mCurSendCard.str������ & ",�����ڽ��з�������!", vbInformation + vbOKOnly
                Check�������� = False
            Case 2 'ͬһ�������������ſ�,����Ҫ����
               Check�������� = MsgBox("�ò����Ѿ�����" & mCurSendCard.str������ & ",�Ƿ�Ҫ���з�������?", vbQuestion + vbYesNo) = vbYes
        End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function WhetherTheCardBinding(ByVal str���� As String, ByVal lng����� As Long, Optional ByRef lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��ȡָ�������Ƿ��Ѿ�����
'���:str���ţ����� ��lng����𣺿���� , lngPatientID :����ID
'����:True :�Ѿ�����;False:δ����
'����:
'����:
'�����:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl
    strSQL = "" & _
           "   Select ����ID From ����ҽ�ƿ���Ϣ Where ����=[1]  And �����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Һ�", str����, lng�����)
    WhetherTheCardBinding = rsTemp.RecordCount > 0

    If rsTemp.RecordCount > 0 Then
        lngPatientID = Val(Nvl(rsTemp!����ID))
    End If

    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function GetCardLastChangeType(ByVal str���� As String, ByVal lng����� As Long, ByVal lngPaitentID As Long) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��ȡ�����ı䶯����
'���:str���ţ����� ��lng����𣺿���� , lngPatientID :����ID
'����:0-δ�ҵ������Ϣ   1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
'����:��⸣
'����:2013-2-4 17:36:33
'�����:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "     Select �䶯���" & vbNewLine & _
           "    From (With ҽ�ƿ��䶯 As (Select ����id, ID, �䶯���, �䶯ʱ�� " & vbNewLine & _
           "                              From ����ҽ�ƿ��䶯 Bd" & vbNewLine & _
           "                              Where Bd.���� = [2] And �����id = [1] And ����id = [3])" & vbNewLine & _
           "           Select A.�䶯���" & vbNewLine & _
           "           From ҽ�ƿ��䶯 A, (Select Max(�䶯ʱ��) As �䶯ʱ�� From ҽ�ƿ��䶯 C) B" & vbNewLine & _
           "           Where A.�䶯ʱ�� = B.�䶯ʱ��) A"
    On Error GoTo ErrHand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����䶯��Ϣ", lng�����, str����, lngPaitentID)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            GetCardLastChangeType = Val(Nvl(rsTmp!�䶯���))
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
'����:ȡ���󶨿�
'���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
'����:ȡ���ɹ�,����true,���򷵻�False
'����:���˺�
'����:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curDate As Date
    Dim strSQL As String, strPassWord As String

    On Error GoTo errHandle

    curDate = zlDatabase.Currentdate

    'Zl_ҽ�ƿ��䶯_Insert
    strSQL = "Zl_ҽ�ƿ��䶯_Insert("
    '      �䶯����_In   Number,
    '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
    strSQL = strSQL & "" & 14 & ","
    '      ����id_In     סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "" & lngPatientID & ","
    '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "NULL,"
    '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & strCardNo & "'" & ","
    '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    strSQL = strSQL & "'���ظ������Զ�ȡ��ԭ������Ϣ',"
    '      ����_In       ������Ϣ.����֤��%Type,
    strSQL = strSQL & "NULL,"
    '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSQL = strSQL & "NULL,"
    '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic����_In     ������Ϣ.Ic����%Type := Null,
    strSQL = strSQL & "NULL,"
    '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
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
            '������һ��
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
'            '������һ��
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
    '����:�������ƶ��ؼ���
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If objCtl.Enabled And objCtl.Visible = True Then: objCtl.SetFocus
End Sub

Private Sub vsInoculate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsInoculate_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strCurDate As String
    
    If Col = 1 Or Col = 3 Then '���������б༭ʱ���ж��Ƿ�����������200
        With vsInoculate
           If LenB(StrConv(.EditText, vbFromUnicode)) > 200 Then
                If MsgBox("�������������ַ���������ַ���200,�����Ƿ񽫶�����ַ������Զ��س���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    .EditText = StrConv(MidB(StrConv(.EditText, vbFromUnicode), 1, 200), vbUnicode)
                Else
                    Cancel = True
                End If
           End If
        End With
    Else
        With vsInoculate
            If IsDate(Format(.EditText, "YYYY-MM-DD")) = False And .EditText <> "    -  -  " Then
                 MsgBox "��������ڸ�ʽ���Ի�����ȷ�����ڣ����飡", vbInformation, gstrSysName
                 Cancel = True
            ElseIf .EditText = "    -  -  " Then
                 .EditText = ""
            Else
                If .EditText <> "" Then
                    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
                    If Format(.EditText, "YYYY-MM-DD") > strCurDate Then
                        MsgBox "�������ڲ��ܴ��ڷ�����ϵͳʱ��[" & strCurDate & "],���飡", vbInformation, gstrSysName
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
            txt��ϵ������.Text = .TextMatrix(.FixedRows, .ColIndex("��ϵ������"))
            For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                If NeedName(cbo��ϵ�˹�ϵ.List(i)) = .TextMatrix(.FixedRows, .ColIndex("��ϵ�˹�ϵ")) Then
                    Exit For
                End If
            Next
            If i < cbo��ϵ�˹�ϵ.ListCount Then
                cbo��ϵ�˹�ϵ.ListIndex = i
            Else
                cbo��ϵ�˹�ϵ.ListIndex = -1
            End If
            
            txt��ϵ�����֤.Text = .TextMatrix(.FixedRows, .ColIndex("��ϵ�����֤��"))
            txt��ϵ�˵绰.Text = .TextMatrix(.FixedRows, .ColIndex("��ϵ�˵绰"))
            
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
        If Col = .ColIndex("��ϵ�����֤��") Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        ElseIf Col = .ColIndex("��ϵ�˵绰") Then
            If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End With
End Sub

Private Sub vsLinkMan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Integer
    
    With vsLinkMan
        If Not Row = .FixedRows Then Exit Sub
        Select Case Col
            Case .ColIndex("��ϵ������")
                txt��ϵ������.Text = Trim(.EditText)
            Case .ColIndex("��ϵ�˹�ϵ")
                For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                    If NeedName(cbo��ϵ�˹�ϵ.List(i)) = Trim(.EditText) Then Exit For
                Next
                If i < cbo��ϵ�˹�ϵ.ListCount Then
                    cbo��ϵ�˹�ϵ.ListIndex = i
                Else
                    cbo��ϵ�˹�ϵ.ListIndex = -1
                End If
                
            Case .ColIndex("��ϵ�����֤��")
                txt��ϵ�����֤.Text = Trim(.EditText)
            Case .ColIndex("��ϵ�˵绰")
                txt��ϵ�˵绰.Text = Trim(.EditText)
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
'����:�����Ƿ����ýṹ����ַ��������
    Dim i As Long
    
    If gbln���ýṹ����ַ Then
        For i = PatiAddress.LBound To PatiAddress.UBound
            If i = E_IX_��סַ Or i = E_IX_���ڵ�ַ Or i = E_IX_��ϵ�˵�ַ Then
                PatiAddress(i).Items = Five
            End If
            PatiAddress(i).TextBackColor = &H80000005
            PatiAddress(i).Visible = True
            PatiAddress(i).ShowTown = gbln��ʾ����
        Next
        For i = LBound(marrAddress) To UBound(marrAddress)
            marrAddress(i) = ""
        Next
        txt��ͥ��ַ.Visible = False
        cmd��ͥ��ַ.Visible = False
        txt�����ص�.Visible = False
        cmd�����ص�.Visible = False
        txt���ڵ�ַ.Visible = False
        cmd���ڵ�ַ.Visible = False
        txt����.Visible = False
        cmd����.Visible = False
        txt��ϵ�˵�ַ.Visible = False
        cmd��ϵ�˵�ַ.Visible = False
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = False
        Next
        
        txt��ͥ��ַ.Visible = True
        cmd��ͥ��ַ.Visible = True
        txt�����ص�.Visible = True
        cmd�����ص�.Visible = True
        txt���ڵ�ַ.Visible = True
        cmd���ڵ�ַ.Visible = True
        txt����.Visible = True
        cmd����.Visible = True
        txt��ϵ�˵�ַ.Visible = True
        cmd��ϵ�˵�ַ.Visible = True
    End If
End Sub

Private Sub SetStrutAddress(Optional ByVal bytFunc As Byte)
'����:89980���˽ṹ��
'bytFunc=1 �������
'       =2 ���û��ڵ�ַ�ͼ�ͥ��ַ��ȱʡֵ
    Dim i As Long
    
    If bytFunc = 2 Then
        txt��ͥ��ַ.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        txt���ڵ�ַ.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        Call PatiAddress(E_IX_��סַ).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
        Call PatiAddress(E_IX_���ڵ�ַ).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
            If bytFunc = 1 Then
                PatiAddress(i).Value = ""
            Else
                PatiAddress(i).Enabled = (mbytInState <> EState.E����)
            End If
        Next
    End If
End Sub

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ���ӿ�
    '����: true-�ɹ���false-ʧ��
    '����:���ϴ�
    '����:2016/6/20 13:54:56
    '����:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    If txt����.Locked Then Exit Function
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    '����ҽ�ƿ����߲���ˢ��
    If mCurSendCard.lng�����ID = 0 Or Not mCurSendCard.blnˢ�� Then Exit Function
    If mobjSquare.zlSetBrushCardObject(mCurSendCard.lng�����ID, IIf(blnComm, txt����, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call mobjSquare.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function

Private Sub ReCalcBirthDay()
    Dim strBirth As String

    If CreatePublicPatient() Then
        If gobjPublicPatient.ReCalcBirthDay(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, Trim(cbo���䵥λ.Text), ""), strBirth) Then
            If txt��������.Enabled Then txt��������.Text = Format(strBirth, "YYYY-MM-DD")
            
            If txt����ʱ��.Enabled Then
                strBirth = Format(strBirth, "HH:MM")
                txt����ʱ��.Text = IIf(strBirth = "00:00", "__:__", strBirth)
            End If
        End If
    End If
End Sub

Private Function CheckMobile(ByVal strMobile As String, ByVal lngPatiID As Long) As Boolean
'����:��鵱ǰ�ֻ����Ƿ����
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM ������Ϣ Where �ֻ��� = [1] And ����ID <> [2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ֻ���", strMobile, lngPatiID)
    If Not rsTemp Is Nothing Then
        CheckMobile = rsTemp.EOF = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub EMPI_LoadPati()
'����:��EMPI�������Ĳ�����Ϣ���µ�����
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim str�������� As String
    Dim blnRet As Boolean

    If CreatePlugInOK(glngModul) Then
        '��֯���˻�����Ϣ
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "����ID", adBigInt
            .Append "��ҳID", adBigInt
            .Append "�Һ�ID", adBigInt
            '-------------------------------
            .Append "�����", adVarChar, 18
            .Append "סԺ��", adVarChar, 18
            .Append "ҽ����", adVarChar, 30
            .Append "���֤��", adVarChar, 18
            .Append "����֤��", adVarChar, 20
            .Append "����", adVarChar, 100
            .Append "�Ա�", adVarChar, 4
            .Append "��������", adVarChar, 20 '���ڸ�ʽ��YYYY-MM-DD HH:MM:SS
            .Append "�����ص�", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����", adVarChar, 20
            .Append "ѧ��", adVarChar, 10
            .Append "ְҵ", adVarChar, 80
            .Append "������λ", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����״��", adVarChar, 4
            .Append "��ͥ�绰", adVarChar, 20
            .Append "��ϵ�˵绰", adVarChar, 20
            .Append "��λ�绰", adVarChar, 20
            .Append "��ͥ��ַ", adVarChar, 100
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6
            .Append "���ڵ�ַ", adVarChar, 100
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6
            .Append "��λ�ʱ�", adVarChar, 6
            .Append "��ϵ�˵�ַ", adVarChar, 100
            .Append "��ϵ�˹�ϵ", adVarChar, 30
            .Append "��ϵ������", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open
        
        If txt����ʱ�� = "__:__" Then
            str�������� = IIf(IsDate(txt��������.Text), Format(txt��������.Text, "YYYY-MM-DD"), "")
        Else
            str�������� = IIf(IsDate(txt��������.Text), Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS"), "")
        End If
 
        With rsPatiIn
            .AddNew
            !����ID = CLng(txt����ID.Text)
            !��ҳID = mlng��ҳID
            !סԺ�� = IIf(txtסԺ��.Text <> "", txtסԺ��.Text, "")
            !����� = IIf(Trim(txt�����.Text) <> "", Trim(txt�����.Text), "")
            !ҽ���� = txtPatiMCNO(0).Text
            '-Ҫ���µ��ֶ�--------------------------------------------
            !���֤�� = Trim(txt���֤��.Text)
            !����֤�� = Trim(txt����֤��.Text)
            !���� = Trim(txtPatient.Text)
            !�Ա� = zlCommFun.GetNeedName(cbo�Ա�.Text)
            !�������� = str�������� '���ڸ�ʽ��YYYY-MM-DD HH:MM:SS
            !�����ص� = Trim(txt�����ص�.Text)
            !���� = zlCommFun.GetNeedName(cbo����.Text)
            !���� = zlCommFun.GetNeedName(cbo����.Text)
            !ѧ�� = zlCommFun.GetNeedName(cboѧ��.Text)
            !ְҵ = zlCommFun.GetNeedName(cboְҵ.Text)
            !������λ = Trim(txt������λ.Text)
            !����״�� = zlCommFun.GetNeedName(cbo����״��.Text)
            !��ͥ�绰 = Trim(txt��ͥ�绰.Text)
            !��ϵ�˵绰 = Trim(txt��ϵ�˵绰.Text)
            !��λ�绰 = Trim(txt��λ�绰.Text)
            !��ͥ��ַ = Trim(txt��ͥ��ַ.Text)
            !��ͥ��ַ�ʱ� = Trim(txt��ͥ��ַ�ʱ�.Text)
            !���ڵ�ַ = Trim(txt���ڵ�ַ.Text)
            !���ڵ�ַ�ʱ� = Trim(txt���ڵ�ַ�ʱ�.Text)
            !��λ�ʱ� = Trim(txt��λ�ʱ�.Text)
            !��ϵ�˵�ַ = Trim(txt��ϵ�˵�ַ.Text)
            !��ϵ�˹�ϵ = zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text)
            !��ϵ������ = Trim(txt��ϵ������.Text)
            .Update
            '-------------------------------------------------------
        End With
        
        '���ò�ѯ�ӿ�
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsPatiIn, rsPatiOut)
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Sub
        If rsPatiOut Is Nothing Then Exit Sub
        If rsPatiOut.RecordCount = 0 Then Exit Sub
        '�ҵ����ˣ����������µ���Ϣ���µ�����
        With rsPatiOut
            mblnEMPI = True     '���ڱ���ҵ���������
            '104916 ֻ��������,�ӿڵ����������������Ϣ�ҵ�HIS����IDʱ�������½�����
            If mbytInState = E���� And CLng(txt����ID.Text) <> CLng(!����ID & "") And CLng(!����ID & "") <> 0 Then
                ClearCard
                txt����ID.Text = !����ID
                Call ReadPatiCard(CLng(txt����ID.Text))
            End If
            Call cbo.Locate(cbo����, !���� & "")
            Call cbo.Locate(cbo����, !���� & "")
            Call cbo.Locate(cboѧ��, !ѧ�� & "")
            Call cbo.SeekIndex(cboְҵ, !ְҵ & "")
            Call cbo.Locate(cbo����״��, !����״�� & "")
            Call cbo.Locate(cbo��ϵ�˹�ϵ, !��ϵ�˹�ϵ & "")
            
            If mbytInState = EState.E���� Then
                '�޸�ʱ������EMPIֱ�Ӹ��²��˵Ļ�����Ϣ
                txtPatient.Text = !���� & ""
                Call cbo.Locate(cbo�Ա�, !�Ա� & "")
                If IsDate(!�������� & "") Then
                    txt��������.Text = Format(CDate(!�������� & ""), "YYYY-MM-DD")
                    txt����ʱ��.Text = IIf(Format(CDate(!�������� & ""), "HH:MM") = "00:00", "__:__", Format(CDate(!�������� & ""), "HH:MM"))
                End If
            End If
            
            If gbln���ýṹ����ַ Then
                PatiAddress(E_IX_�����ص�).Value = !�����ص� & ""
                PatiAddress(E_IX_��סַ).Value = !��ͥ��ַ & ""
                PatiAddress(E_IX_���ڵ�ַ).Value = !���ڵ�ַ & ""
                PatiAddress(E_IX_��ϵ�˵�ַ).Value = !��ϵ�˵�ַ & ""
            End If
            txtPatiMCNO(0).Text = !ҽ���� & ""
            txt�����ص�.Text = !�����ص� & ""
            txt��ͥ��ַ.Text = !��ͥ��ַ & ""
            txt���ڵ�ַ.Text = !���ڵ�ַ & ""
            txt��ϵ�˵�ַ.Text = !��ϵ�˵�ַ & ""
            txt���֤��.Text = !���֤�� & ""
            txt����֤��.Text = !����֤�� & ""
            txt������λ.Text = !������λ & ""
            txt��ͥ�绰.Text = !��ͥ�绰 & ""
            txt��ϵ�˵绰.Text = !��ϵ�˵绰 & ""
            txt��λ�绰.Text = !��λ�绰 & ""
            txt��ͥ��ַ�ʱ�.Text = !��ͥ��ַ�ʱ� & ""
            txt���ڵ�ַ�ʱ�.Text = !���ڵ�ַ�ʱ� & ""
            txt��λ�ʱ�.Text = !��λ�ʱ� & ""
            txt��ϵ������.Text = !��ϵ������ & ""
        End With
    End If
End Sub

Private Function EMPI_AddORUpdatePati(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strErr As String) As Boolean
'����:���ӻ����EMPI������Ϣ
    Dim lngRet  As Long
    Dim strPlugErr As String
    Dim strTmp As String
    
    lngRet = 1 'Ĭ�ϳɹ� ���� �ϰ�zlPlug����֧�ִ˽ӿڴ����:438
    If CreatePlugInOK(glngModul) Then
        
        If Not mblnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "��EMPIƽ̨����������Ϣʧ�ܣ�"
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "��EMPIƽ̨���²�����Ϣʧ�ܣ�"
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
    '�뿪��鿨��
    Dim lng����ID As Long, lng�շ�ϸĿID As Long
    Dim strSQL As String, str���� As String
    Dim rsTmp As ADODB.Recordset
    
    If mCurSendCard.rs���� Is Nothing Then Exit Sub
    If mCurSendCard.rs����.RecordCount = 0 Then Exit Sub
    If mCurSendCard.lng�����ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt����.Text) = "" Then Exit Sub
    If mbytInState = E���� Then
        lng����ID = mlngPatientID
    Else
        lng����ID = mlng����ID
    End If
    If blnFeedName = False And lng����ID <> 0 Then Exit Sub
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    mCurSendCard.rs����.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as �շ�ϸĿID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����", mlngModul, mCurSendCard.lng�����ID, Trim(txt����.Text), lng����ID, _
                Trim(txtPatient.Text), NeedName(cbo�Ա�.Text), str����, Trim(txt���֤��.Text), Val(Nvl(mCurSendCard.rs����!�շ�ϸĿID)))
    If rsTmp.EOF Then Exit Sub
    
    lng�շ�ϸĿID = Val(Nvl(rsTmp!�շ�ϸĿID))
    Set rsTmp = GetSpecialInfo("", lng�շ�ϸĿID)
    If Not rsTmp Is Nothing Then Set mCurSendCard.rs���� = rsTmp
    
    With mCurSendCard.rs����
        txt����.Text = Format(IIf(Val(Nvl(!�Ƿ���)) = 1, Val(Nvl(!ȱʡ�۸�)), Val(Nvl(!�ּ�))), "0.00")
        txt����.Tag = txt����.Text  '���ֲ���
        txt����.Locked = Not (Val(Nvl(!�Ƿ���)) = 1)
        txt����.TabStop = (Val(Nvl(!�Ƿ���)) = 1)
        
        If Val(Nvl(mCurSendCard.rs����!���ηѱ�)) <> 1 And Val(txt����.Text) <> 0 Then
            txt����.Text = Format(GetActualMoney(NeedName(cbo�ѱ�.Text), mCurSendCard.rs����!������ĿID, Val(txt����.Text), mCurSendCard.rs����!�շ�ϸĿID), "0.00")
        End If
    End With
End Sub

Private Function CheckPatiIsExist() As Boolean
'����:������Ʋ�����Ϣ(����֮ǰ���,����������ظ���Ϣ������)
'����ֵ:
'   True-��������;
'   False-�жϱ���
    Dim strSimilar As String
    Dim strType As String
    Dim strNote As String
    Dim lngPatiID As Long
    Dim rsSimilar As ADODB.Recordset
    Dim i As Long
    
    If Not CreatePublicPatient() Then Exit Function
    strType = IIf(glngSys Like "8??", "�ͻ�", "����")
    '������Ʋ�����Ϣ(����֮ǰ���,����������ظ���Ϣ������)
    strSimilar = SimilarIDs(zlCommFun.GetNeedName(cbo����.Text), zlCommFun.GetNeedName(cbo����), CDate(IIf(IsDate(txt��������.Text), txt��������.Text, #1/1/1900#)), zlCommFun.GetNeedName(cbo�Ա�), txtPatient.Text, txt���֤��.Text, rsSimilar)
    If strSimilar <> "" Then
        If gblnPatiByID And Trim(txt���֤��.Text) <> "" Then
            '110541 ͬһ���ֻ֤�ܶ�Ӧһ����������;���øò�����ͨ�����֤���ҵ��ѽ�������ʱ����ѡ���
            rsSimilar.Filter = "���֤�� ='" & Trim(txt���֤��.Text) & "'"
            If rsSimilar.RecordCount > 0 Then
                strNote = "�����еĲ�����Ϣ�з���" & rsSimilar.RecordCount & "�����֤����ͬ�ĵĲ��ˡ�" & vbCrLf & vbCrLf & _
                    "��ȡ���еĲ�����Ϣ��ѡ���˺�[˫��]����[ȷ��]��"
                If gobjPublicPatient.ShowSelect(rsSimilar, "ID", "����ѡ��", strNote, , , "0|800|1200|800|800|1500|1000", True) Then
                    lngPatiID = Val(rsSimilar!����ID & "")
                    GoTo FindPati:
                End If
            End If
        End If
                    
        i = UBound(Split(strSimilar, "|")) + 1
        strSimilar = Replace(strSimilar, "|", vbCrLf)
        If i > 20 Then strSimilar = Mid(strSimilar, 1, 200) & "..."
        If MsgBox("�����е�" & strType & "��Ϣ�з��� " & i & " ����Ϣ���Ƶ�" & strType & "(����,����,�Ա�,����,����������ͬ�����֤����ͬ): " & vbCrLf & vbCrLf & _
            strSimilar & vbCrLf & vbCrLf & "ȷʵҪ�����" & strType & "����Ϣ��", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
        MsgBox "��" & strType & "�����Ƽ�¼����ʹ��""�ϲ�""���ܴ���", vbInformation, gstrSysName
    End If
    
    CheckPatiIsExist = True
End Function
