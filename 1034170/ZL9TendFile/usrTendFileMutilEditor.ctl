VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrTendFileMutilEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   11460
   Begin VB.PictureBox picNull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   1005
      ScaleHeight     =   1215
      ScaleWidth      =   7335
      TabIndex        =   34
      Top             =   1335
      Width           =   7365
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "然后点击刷新按钮装载数据..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   30
         TabIndex        =   36
         Top             =   540
         Width           =   8115
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "请选择一种护理文件格式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   30
         TabIndex        =   35
         Top             =   60
         Width           =   8145
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8415
      TabIndex        =   32
      Top             =   3840
      Width           =   8415
   End
   Begin VB.PictureBox pic过滤条件 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   450
      ScaleHeight     =   330
      ScaleWidth      =   10695
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   30
      Width           =   10695
      Begin VB.OptionButton optLevel 
         Caption         =   "纵向"
         Height          =   180
         Index           =   1
         Left            =   9720
         TabIndex        =   46
         Top             =   67
         Width           =   735
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "横向"
         Height          =   180
         Index           =   0
         Left            =   8640
         TabIndex        =   45
         Top             =   67
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmd刷新 
         Caption         =   "刷新(&R)"
         Height          =   315
         Left            =   6660
         TabIndex        =   31
         Top             =   0
         Width           =   885
      End
      Begin VB.ComboBox cbo科室 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin VB.CheckBox chk出院 
         Caption         =   "出院"
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   5940
         TabIndex        =   29
         ToolTipText     =   "勾选表示提取出院病人"
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox chk出科 
         Caption         =   "出科"
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   5190
         TabIndex        =   28
         ToolTipText     =   "勾选表示提取出科病人"
         Top             =   60
         Width           =   675
      End
      Begin VB.ComboBox cbo护理文件格式 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   0
         Width           =   2205
      End
      Begin VB.Label lblEntry 
         AutoSize        =   -1  'True
         Caption         =   "录入方式"
         Height          =   180
         Left            =   7800
         TabIndex        =   47
         Top             =   67
         Width           =   720
      End
      Begin VB.Label lbl科室 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3180
         TabIndex        =   26
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl文件格式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "文件格式"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   30
      Top             =   510
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
            Picture         =   "usrTendFileMutilEditor.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileMutilEditor.ctx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2550
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   60
      ScaleHeight     =   3315
      ScaleWidth      =   8385
      TabIndex        =   13
      Top             =   510
      Width           =   8385
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   930
         Index           =   2
         ItemData        =   "usrTendFileMutilEditor.ctx":0734
         Left            =   4200
         List            =   "usrTendFileMutilEditor.ctx":0747
         TabIndex        =   6
         Top             =   675
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.PictureBox PicLst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   2310
         Left            =   5310
         ScaleHeight     =   2280
         ScaleWidth      =   1185
         TabIndex        =   3
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   1470
            Index           =   0
            ItemData        =   "usrTendFileMutilEditor.ctx":076B
            Left            =   -10
            List            =   "usrTendFileMutilEditor.ctx":0781
            TabIndex        =   5
            Top             =   825
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtLst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   -10
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "录入："
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   44
            Top             =   30
            Width           =   540
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "选择："
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   15
            TabIndex        =   43
            Top             =   615
            Width           =   540
         End
      End
      Begin VB.PictureBox picDoubleChoose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4680
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picChooseRight 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   1
               ItemData        =   "usrTendFileMutilEditor.ctx":07B9
               Left            =   -30
               List            =   "usrTendFileMutilEditor.ctx":07C9
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.PictureBox picChooseLeft 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   0
               ItemData        =   "usrTendFileMutilEditor.ctx":07DB
               Left            =   -30
               List            =   "usrTendFileMutilEditor.ctx":07EB
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   42
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.CommandButton cmdWord 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6300
         Picture         =   "usrTendFileMutilEditor.ctx":07FD
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5640
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picDnInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.Label lblDnInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   5.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   60
               TabIndex        =   21
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.PictureBox picUpInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.Label lblUpInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   5.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   60
               TabIndex        =   20
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   10
            Top             =   30
            Width           =   345
         End
         Begin VB.TextBox txtUpInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   30
            MaxLength       =   12
            TabIndex        =   9
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   17
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "usrTendFileMutilEditor.ctx":0B3F
         Left            =   6540
         List            =   "usrTendFileMutilEditor.ctx":0B55
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   660
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5970
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   615
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "√"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   15
            Top             =   30
            Width           =   315
         End
      End
      Begin VB.PictureBox picMutilInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6810
         ScaleHeight     =   435
         ScaleWidth      =   1575
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1600
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   810
            TabIndex        =   12
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "体温体录"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   15
            TabIndex        =   14
            Top             =   112
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   390
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
         Appearance      =   0
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendFileMutilEditor.ctx":0B8D
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一般护理记录单"
         Height          =   180
         Left            =   3720
         TabIndex        =   30
         Top             =   90
         Width           =   1275
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfHistory 
      Height          =   1545
      Left            =   60
      TabIndex        =   33
      Top             =   3885
      Width           =   4305
      _cx             =   7594
      _cy             =   2725
      Appearance      =   0
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"usrTendFileMutilEditor.ctx":0BEF
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      AutoSizeMouse   =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrTendFileMutilEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnHistory As Boolean              '历史表格是否已初始化
Private mblnShow As Boolean                 '是否显示录入框
Private mblnBlowup As Boolean               '放大否？放大1/3，如字体9号放大为12号
Private mblnChange As Boolean               '是否修改数据
Private mblnSaved As Boolean                '是否已保存
Private mblnSigned As Boolean               '是否已签名
Private mstrData As String                  '进入编辑状态前保存之前的数据
Private mintPreDays As Long
Private mstrMaxDate As String
Private mlngSingerType As Long              '护士、签名列显示模式

Private mlng文件ID As Long
Private mlng格式ID As Long
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mint页码 As Integer
Private mstrPrivs As String

Private mdtOutEnd As Date
Private mdtOutbegin As Date
Private mintChange As Integer
Private mstrBPItem As String                '血压频次对应的诊疗项目ID,格式1,2

Private mintSymbol As Integer               '当前控件索引
Private mstrSymbol As String                '特殊字符
Private mstrCollectItems As String          '汇总项目集合
Private mstrColCollect As String            '汇总项目列集合:col;1|col;4,5
Private mstrColCorrelative As String        '汇总项目关联列集合:COl,3;COl,4|COl,5;COl,6(名称列号,项目序号;汇总列,项目序号),主要针对分类汇总
Private mstrColImCorrelative As String    '汇总项目关联列集合:COl,3;COl,4|COl,5;COl,6(名称列号,项目序号;汇总列,项目序号),主要针对入量导入
Private mstrCOLNothing As String            '未绑定的列集合+活动项目列(不管活动项目列是否绑定)
Private mstrCOLActive As String             '活动列集合
Private mstrCatercorner As String           '列对角线集合
Private mblnEditAssistant As Boolean        '当前选择的项目是否允许进行词句选择
Private mblnEditText As Boolean             '选择的项目是否是文本项目
Private mblnEditHistoryAssistant As Boolean
Private mlngPageRows As Long                '此文件格式一页所显示的数据行
Private mlngOverrunRows As Long             '超出数据行
Private mlngRowCount As Long                '当前记录总行数
Private mlngRowCurrent As Long              '当前记录在本页的实际行数
Private mlngDate As Long                    '日期
Private mlngTime As Long                    '时间
Private mlngOperator As Long                '护士
Private mlngSignLevel As Long               '签名级别
Private mlngSigner As Long                  '签名信息
Private mlngSignName As Long                '签名人
Private mlngSignTime As Long                '签名时间
Private mlngRecord As Long                  '记录ID
Private mlngNoEditor As Long                '禁止编辑列,存在护士列则以护士列为准,不存在护士列则以签名列为准
Private mlngActiveTime As Long              '发生时间


Private mintType As Integer                 '记录当前的编辑模式
Private mblnDateAd As Boolean               '日期缩写?
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
Private mrsTemperItems As New ADODB.Recordset       '所有体温记录项目
Private mrsSelItems As New ADODB.Recordset          '当前录入的护理记录项目清单
Private mrsDataMap As New ADODB.Recordset           '当前操作员录入的数据镜像,与记录单格式一致,相关行数据全部保存以便迅速恢复
Private mrsCellMap As New ADODB.Recordset           '编辑过的数据镜像,字段有:页号,行号,列号,记录ID,数据,部位,删除
Private mrsCopyMap As New ADODB.Recordset           '复制行数据
Private mrsUsual As New ADODB.Recordset             '常用体温说明

Private Enum ColIcon
    签名 = 1
    审签 = 2
End Enum
Private Enum SignLevel
    正高 = 1
    副高 = 2
    中级 = 3
    师级 = 4
    员士 = 5
    未定义 = 9
End Enum

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public WithEvents zlEvent_Print As zlTFPrintMethod
Attribute zlEvent_Print.VB_VarHelpID = -1
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean)
Public Event UsrHelp()
Public Event UsrExit()
'59118:刘鹏飞,2013-03-05
Public Event ShowTipInfo(ByVal vsfObj As Object, ByVal strInfo As String, ByVal blnMultiRow As Boolean)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'病历文件格式定义相关
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTagFont As New StdFont  '条件样式字体
Private mlngTagColor As Long        '条件样式颜色
Private mstrPaperSet As String      '格式
Private mstrPageHead As String      '页眉
Private mstrPageFoot As String      '页脚
Private mblnChildForm As Boolean
Private mstrSubhead As String       '表上标签
Private mstrTabHead As String       '表头单元
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'保存打开护理记录文件的SQL，在其它地方也有使用，不能修改
Private mstrSQL内 As String
Private mstrSQL中 As String
Private mstrSQL列 As String
Private mstrSQL条件 As String
Private mstrSQL As String

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与绘图相关,没事别动
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const WHITE_BRUSH = 0                   '白色画笔
Private Const cdblWidth As Double = 6           '一个英文字符的宽度
Private Const cHideCols = 8                    '前缀列:床号,姓名
Private Const cControlFields = 2                '记录集控制列:页号,行号
Private Const mlngDemo As Long = 1                  '分组列
Private Const c分组 As Integer = 1
Private Const c文件ID As Integer = 2
Private Const c床号 As Integer = 3
Private Const c姓名 As Integer = 4
Private Const c病人ID As Integer = 5
Private Const c主页ID As Integer = 6
Private Const c婴儿 As Integer = 7
Private Const c血压频次 As Integer = 8

Private Const p住院护士站 = 1262

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '将VB的颜色转换为RGB表示
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long

    OleTranslateColor dwOleColour, 0, clrref

    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF

    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    '缺省是宋体9号,按字体大小同比放大
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer

    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '绘图相关
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo ErrHand
    '******************************************
    '在此事件中不能对单元格的任何属性赋值,包括Celldata,否则会引起该事件的死循环,导致工具栏或计时器无法正常工作。
    '******************************************
    '使用匹配的背景色，前景色与字体进行文本输出。
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False

    strText = VsfData.TextMatrix(ROW, COL)
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '赋初值
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        '取字符宽度
        dblWidth = GetSymbolWidth(strRight)
        '设定客户区域大小
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With

        '1、清空内容
        '创建与背景色相同的刷子
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        '使用该刷子填充背景色
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '立即销毁临时使用的刷子并还原刷子
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)

        '2、准备画线
        '创建新画笔
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '画线
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '输出文本
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)

        '还原画笔并销毁
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)

        '已完成作图
        Done = True
    End If
'
'    '3、如果是汇总行，则进行特殊处理
'    If Val(VsfData.TextMatrix(Row, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(Row, mlngCollectStyle)) = 1 _
'        And (Col >= mlngDate And Col < mlngNoEditor) Then
'        Call DrawCollectCell(hDC, Left, Top, Right, Bottom)
'    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub DrawCellHistory(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '绘图相关
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo ErrHand
    '******************************************
    '在此事件中不能对单元格的任何属性赋值,包括Celldata,否则会引起该事件的死循环,导致工具栏或计时器无法正常工作。
    '******************************************
    '使用匹配的背景色，前景色与字体进行文本输出。
    If Not mblnInit Then Exit Sub
    If vsfHistory.RowHidden(ROW) Then Exit Sub
    Done = False

    strText = vsfHistory.TextMatrix(ROW, COL)
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '赋初值
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        '取字符宽度
        dblWidth = GetSymbolWidth(strRight)
        '设定客户区域大小
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With

        '1、清空内容
        '创建与背景色相同的刷子
        If ROW < vsfHistory.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(vsfHistory.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(vsfHistory.ForeColorFixed)
        Else
            If ROW = vsfHistory.RowSel Then
                lngBackColor = GetRBGFromOLEColor(vsfHistory.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(vsfHistory.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        '使用该刷子填充背景色
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '立即销毁临时使用的刷子并还原刷子
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)

        '2、准备画线
        '创建新画笔
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '画线
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '输出文本
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)

        '还原画笔并销毁
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)

        '已完成作图
        Done = True
    End If
'
'    '3、如果是汇总行，则进行特殊处理
'    If Val(vsfHistory.TextMatrix(Row, mlngCollectType)) < 0 And Val(vsfHistory.TextMatrix(Row, mlngCollectStyle)) = 1 _
'        And (Col >= mlngDate And Col < mlngNoEditor) Then
'        Call DrawCollectCell(hDC, Left, Top, Right, Bottom)
'    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与分行相关,没事别动
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '不为零,表示仅设置字符串结束符
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long

    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

'**********************************************************************************************************************
'######################################################################################################################

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    On Error GoTo ErrHand

    '读取文件属性
    mblnDateAd = False

    '提取活动项目并加入列定义(格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    mstrColCorrelative = ""
    mstrColImCorrelative = ""
    '读取病历文件格式定义
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlng格式ID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数":  VsfData.Cols = Val("" & !内容文本): vsfHistory.Cols = Val("" & !内容文本)
            Case "最小行高": VsfData.RowHeightMin = BlowUp(Val("" & !内容文本)): vsfHistory.RowHeightMin = BlowUp(Val("" & !内容文本))
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set vsfHistory.Font = objFont
                Set Font = objFont
            Case "文本颜色"
                VsfData.ForeColor = Val("" & !内容文本)
                vsfHistory.ForeColor = Val("" & !内容文本)
            Case "表格颜色"
                VsfData.GridColor = Val("" & !内容文本): VsfData.GridColorFixed = VsfData.GridColor
                vsfHistory.GridColor = Val("" & !内容文本): vsfHistory.GridColorFixed = vsfHistory.GridColor
            
            Case "标题文本"
                lblTitle.Caption = "" & !内容文本
                lblTitle.AutoSize = True
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False

            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "条件颜色": mlngTagColor = Val("" & !内容文本)
            Case "有效数据行"
                mlngOverrunRows = 0
                mlngPageRows = Val("" & !内容文本)
            End Select
            .MoveNext
        Loop
    End With
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表头单元定义", mlng格式ID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With

    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql外 As String, str格式 As String, strSqlNull As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim bln对角线 As Boolean, bln选择项 As Boolean          '如果上一列是对角线且选择项,则直接提取各项数据,拼列头时在数值间加上/
    Dim lngColumn As Long, blnAddCollect As Boolean
    Dim strColCorrelative  As String
    
    gstrSQL = "Select   d.对象序号,d.对象标记, d.对象属性, d.内容行次, d.内容文本, upper(d.要素名称) AS 要素名称, d.要素单位,d.要素表示,d.要素值域 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = "": strColCorrelative = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = "": strSqlNull = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            If lngColumn <> !对象序号 Then
                blnAddCollect = False
                If strColCorrelative <> "" Then
                    mstrColCorrelative = mstrColCorrelative & "|" & strColCorrelative
                End If
                strColCorrelative = ""
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) & "|" & !对象序号 & "'" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性 & "`" & !对象序号 & "`" & !要素表示
                If !要素表示 = 1 Then mstrCatercorner = mstrCatercorner & "," & !对象序号
                str格式 = ""
                If !要素名称 <> "" Then str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                If Mid(strSqlNull, 3) = "" Then
                    strSqlNull = "''"
                Else
                    strSqlNull = Mid(strSqlNull, 3)
                End If
                If InStr(1, strSql外, "脉搏") > 0 And InStr(1, strSql外, "心率") > 0 And InStr(1, strSql外, "脉搏") < InStr(1, strSql外, "心率") Then
                    mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",'',''||""脉搏""||''||'/',''||""脉搏""||''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
                Else
                    mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
                End If
 
                strSql外 = ""
                strSqlNull = ""
                lngColumn = !对象序号
                bln对角线 = (NVL(!要素表示, 0) = 1)
                bln选择项 = False
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    bln选择项 = (mrsItems!项目表示 = 5)
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                        If Val(NVL(!对象标记)) > 0 And Val(NVL(!对象序号)) <> Val(NVL(!对象标记)) Then
                            strColCorrelative = Val(NVL(!对象标记)) & ";" & !对象序号 & "," & mrsItems!项目序号
                        End If
                    End If
                    If NVL(!要素值域) <> "" Then
                        mstrColImCorrelative = mstrColImCorrelative & "|" & Val(!对象序号) & "," & mrsItems!项目序号 & "," & NVL(!要素值域)
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        If blnAddCollect Then
                            strColCorrelative = ""
                            mstrColCollect = mstrColCollect & "," & mrsItems!项目序号
                        Else    '有可能一列绑定两个项目,第一个项目不是汇总项目,第二个项目才是汇总项目,因此,下面的代码保证加上列序号
                            blnAddCollect = True
                            mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                            If Val(NVL(!对象标记)) > 0 And Val(NVL(!对象序号)) <> Val(NVL(!对象标记)) Then
                                strColCorrelative = Val(NVL(!对象标记)) & ";" & !对象序号 & "," & mrsItems!项目序号
                            End If
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            End If

            Select Case !要素名称
            Case "日期"
                bln日期 = True
                mblnDateAd = (NVL(!要素表示, 0) = 1)
                mstrSQL中 = mstrSQL中 & ",日期"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As 日期"
                strSql外 = strSql外 & "||" & !要素名称
            Case "时间"
                bln时间 = True
                mstrSQL中 = mstrSQL中 & ",时间"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称

            Case "签名人"
                bln签名人 = True
                mstrSQL中 = mstrSQL中 & ",签名人"
                mstrSQL内 = mstrSQL内 & ",l.签名人"
                strSql外 = strSql外 & "||" & !要素名称

            Case "签名时间"
                bln签名时间 = True
                mstrSQL中 = mstrSQL中 & ",签名时间"
                mstrSQL内 = mstrSQL内 & ",l.签名时间"
                strSql外 = strSql外 & "||" & !要素名称

            Case "护士"
                bln护士 = True
                mstrSQL中 = mstrSQL中 & ",护士"
                mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
                strSql外 = strSql外 & "||" & !要素名称
            Case Else
                If !要素名称 <> "" Then
                    mstrSQL中 = mstrSQL中 & ",Max(""" & !要素名称 & """) As """ & !要素名称 & """"
                    mstrSQL条件 = mstrSQL条件 & " Or """ & !要素名称 & """ Is Not Null"

                    strSql外 = strSql外 & "||'" & !内容文本 & "'||""" & !要素名称 & """||'" & !要素单位 & "'"
                    strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
                    mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', c.记录内容, '') As """ & !要素名称 & """"

'                    If bln对角线 And bln选择项 Then
'                        If strSql外 <> "" Then
'                            '第二项
'                            strSql外 = strSql外 & "||'/'||""" & !要素名称 & """"
'                        Else
'                            '第一项
'                            strSql外 = strSql外 & "||""" & !要素名称 & """"
'                        End If
'                    Else
'                        strSql外 = strSql外 & "||""" & !要素名称 & """"
'                        strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
'                    End If
'
'                    If (Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "") Or (bln对角线 And bln选择项) Then
'                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', c.记录内容, '') As """ & !要素名称 & """"
'                    Else
'                        'mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "'), '') As """ & !要素名称 & """"
'                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "'),  '" & !内容文本 & "'||'" & !要素单位 & "') As """ & !要素名称 & """"
'                    End If
                Else
'                    '为空表示未绑定列,强制加,后面进行替换
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!对象序号, "00"))
'                    mstrSQL中 = mstrSQL中 & ",Max(""" & "C" & Format(!对象序号, "00") & """) As C" & Format(!对象序号, "00")
'                    mstrSQL条件 = mstrSQL条件 & " Or """ & "C" & Format(!对象序号, "00") & """ Is Not Null"
'                    mstrSQL内 = mstrSQL内 & ", C" & Format(!对象序号, "00") & " AS C" & Format(!对象序号, "00")
                End If
            End Select
            .MoveNext
        Loop

        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        
        If mstrColImCorrelative <> "" Then
            mstrColImCorrelative = Mid(mstrColImCorrelative, 2)
        End If
        '在InitRecords中需要给汇总项目关列的名称列明添加项目序号
        If Left(mstrColCorrelative, 1) = "|" Then mstrColCorrelative = Mid(mstrColCorrelative, 2)
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) '& "|" & !对象序号 & "'" & !要素名称
        mstrColumns = Mid(mstrColumns, 2)     '格式如:列号;项目名称1,项目名称2|列号...,实例;1;体温|2;脉搏|3...

        If Mid(strSqlNull, 3) = "" Then
            strSqlNull = "''"
        Else
            strSqlNull = Mid(strSqlNull, 3)
        End If
        mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
 
        If mstrSQL条件 <> "" Then mstrSQL条件 = "(" & Mid(mstrSQL条件, 5) & ")"

        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        If bln护士 = False Then mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"

        If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",l.签名时间"

        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！", vbInformation, gstrSysName
            Exit Function
        End If

        '程序内部控制增加固定列
        mstrSQL中 = UCase(mstrSQL中 & ",MAX(签名级别) AS 签名级别,MAX(签名信息) AS 签名信息,MAX(记录ID) AS 记录ID,MAX(行数) AS 行数,MAX(实际行数) AS 实际行数")
        mstrSQL内 = UCase(mstrSQL内 & ",l.签名级别,l.签名人 AS 签名信息,C.记录ID,P.行数||'' AS 行数,1 AS 实际行数")
        mstrSQL列 = UCase(mstrSQL列 & ",签名级别,签名信息,记录ID,行数,实际行数")

        Call SQLCombination
    End With
    ReadStruDef = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SQLCombination(Optional ByVal lng记录ID As Long = 0)
    Dim str条件 As String
    str条件 = mstrSQL条件
    
    mstrSQL = "Select   '' AS 分组,0 AS 文件ID,'' AS 床号,'' AS 姓名,0 AS 病人ID,0 AS 主页ID,0 AS 婴儿,'' as 血压频次," & Mid(mstrSQL列, 12) & ",发生时间" & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,to_char(发生时间,'yyyy-MM-dd hh24:mi:ss') 发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select c.记录组号,l.发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f,病人护理打印 p " & vbCrLf & _
                "               Where l.ID=p.记录ID And l.Id = c.记录id And l.文件ID+0=f.ID+0 And f.ID=p.文件ID " & _
                "               And nvl(l.汇总类别,0)=0 And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And 1=2)" & vbCrLf & _
                IIf(str条件 <> "", "Where " & str条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号,护士,签名人,签名时间" & _
                                "       Order By 日期, 时间, 发生时间,记录组号,护士,签名人,签名时间)"
End Sub

Private Sub zlRefresh()
    Err = 0: On Error GoTo ErrHand
    Call InitCons

    '产生列记录集
    Call InitRecords

    '装入数据
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID)
    '清除并拷贝记录集结构
    Call DataMap_Init(rsTemp)
    '绑定数据并设置护理记录单的格式,同时实现一行数据分行显示的功能
    Call PreTendFormat(rsTemp)
    
    '初始化历史表格
    Call PreTendFormatHistory(rsTemp)
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '初始化内存数据集

    '数据记录集,用于快速恢复
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "页号,行号"
    '修改单元格记录,用于保存
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|页号," & adDouble & ",18|行号," & adDouble & ",18|" & _
            "列号," & adDouble & ",18|记录ID," & adDouble & ",18|数据," & adLongVarChar & ",4000|部位," & adLongVarChar & ",100|" & _
            "标记," & adLongVarChar & ",100|汇总," & adDouble & ",1|删除," & adDouble & ",1")
    mrsCellMap.Sort = "页号,行号,列号"
    '复制记录集
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
End Sub

Private Function DataMap_Save() As Boolean
    '将当前页面中用户编辑过的数据保存起来,页面切换或保存前触发
    Dim blnExit As Boolean, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim intCount As Integer
    Dim arrRows()
    On Error GoTo ErrHand
    
    '先删除指定页号的所有数据行
    If mrsDataMap.RecordCount <> 0 Then mrsDataMap.MoveFirst
    Do While True
        If mrsDataMap.EOF Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    
    '复制指定页号的所有数据行
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    
    arrRows = Array()
    '清除数据记录ID
    For lngRow = VsfData.FixedRows To lngRows
        blnNULL = True
        If VsfData.RowHidden(lngRow) = False Then
            For intCount = mlngTime + 1 To mlngNoEditor - 1
                If Not VsfData.ColHidden(intCount) Then
                    If VsfData.TextMatrix(lngRow, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(lngRow, intCount), "/") <> 0) Then
                        blnNULL = False
                        Exit For
                    End If
                End If
            Next
            If blnNULL Then
               VsfData.TextMatrix(lngRow, mlngRecord) = ""
            End If
        Else
            VsfData.TextMatrix(lngRow, mlngRecord) = ""
            ReDim Preserve arrRows(UBound(arrRows) + 1)
            arrRows(UBound(arrRows)) = lngRow
        End If
    Next
    '清除隐藏的列
    For lngRow = UBound(arrRows) To 0 Step -1
        If VsfData.ROW >= Val(arrRows(lngRow)) Then
            VsfData.ROW = VsfData.ROW - 1
        End If
        VsfData.RemoveItem Val(arrRows(lngRow))
    Next lngRow
    
    lngRows = VsfData.Rows - 1
    '保存列数据
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!页号 = mint页码
        mrsDataMap!行号 = lngRow
        mrsDataMap!删除 = IIf(VsfData.RowHidden(lngRow), 1, 0)
        For lngCol = 0 To lngCols - VsfData.FixedCols
            mrsDataMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
        Next
        mrsDataMap.Update
    Next
    
    DataMap_Save = True
    
    '刷新历史数据
    Call RefreshHistoryData(VsfData.ROW)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DataMap_Restore() As Boolean
    '将指定页面的数据恢复到表格中
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    On Error GoTo ErrHand
    
    mrsDataMap.MoveFirst
    lngCols = VsfData.Cols - 1
    lngRows = mrsDataMap.RecordCount
    VsfData.Rows = VsfData.FixedRows
    For lngRow = 0 To lngRows - 1
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCol = 0 To lngCols - VsfData.FixedCols
            VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCol).Value)
        Next
        If mrsDataMap!删除 = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        mrsDataMap.MoveNext
    Next
    
    DataMap_Restore = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CellMap_Update(ByVal lngStart As Long, ByVal lngDeff As Long, Optional ByVal blnBig As Boolean = True)
    Dim lngPos As Long
    Dim intCol As Integer
    
    '更新当前页面所有大于起始行的行号数据
    With mrsCellMap
        If lngDeff > 0 Then
            If .RecordCount = 0 Then Exit Sub
            If .RecordCount <> 0 Then .MoveLast
            If .BOF Then Exit Sub
            Do While Not mrsCellMap.BOF
                If !页号 = mint页码 And IIf(blnBig = True, !行号 > lngStart, !行号 = lngStart) Then
                    intCol = !列号
                    lngPos = .AbsolutePosition
                    !行号 = !行号 + lngDeff
                    !ID = mint页码 & "," & !行号 & "," & !列号
                    .Update
                    .MoveFirst
                    .Move lngPos - 2
                Else
                    .MovePrevious
                End If
            Loop
        ElseIf lngDeff < 0 Then
            If .RecordCount = 0 Then Exit Sub
            If .RecordCount <> 0 Then .MoveFirst
            If .EOF Then Exit Sub
            Do While Not mrsCellMap.EOF
                If !页号 = mint页码 And IIf(blnBig = True, !行号 > lngStart, !行号 = lngStart) Then
                    intCol = !列号
                    lngPos = .AbsolutePosition
                    !行号 = !行号 + lngDeff
                    !ID = mint页码 & "," & !行号 & "," & !列号
                    .Update
                    .MoveFirst
                    .Move lngPos
                Else
                    .MoveNext
                End If
            Loop
        End If
        If .RecordCount <> 0 Then .MoveFirst
        
    End With
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    '只拷贝记录集的结构,同时增加页号,行号字段
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer

    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "页号", adDouble, 18
            .Fields.Append "行号", adDouble, 18
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "汇总日期" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:表示新增
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '日期型处理为字符型
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
            End If
        Next
        If blnAddPage Then
            .Fields.Append "删除", adDouble, 1
        End If

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim lngRowCount As Long, lngRowCurrent As Long  '当前记录总行数,当前记录在本页的实际行数
    Dim lngCol As Long, lngMax As Long
    Dim lngRow As Long, lngLastRow As Long
    Dim str发生时间 As String, str发生时间_L As String
    Dim lngStart As Long, lngPrintedRow As Long
    Dim strSignName As String
    Dim blnClear As Boolean
    
    On Error GoTo ErrHand

    Dim arrData
    Dim intData As Integer, intDatas As Integer
    '如果一行显示不完则分行显示(根据当前数据占用行数先添加空白行并处理行坐标,然后再依次处理当前行的数据)
    '每页只显示实际的数据行,把'@处取消注释即可

    lngRow = vsfHistory.FixedRows
    Do While True
        If lngRow > vsfHistory.Rows - 1 Then Exit Do
        'If lngRow >= mlngPageRows + mlngOverrunRows + vsfHistory.FixedRows Then Exit Do
        If InStr(1, vsfHistory.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        lngRowCount = Val(vsfHistory.TextMatrix(lngRow, mlngRowCount))
        '@实际数据行
'        lngRowCurrent = Val(vsfhistory.TextMatrix(lngRow, mlngRowCurrent))
        str发生时间 = Format(vsfHistory.TextMatrix(lngRow, mlngActiveTime), "YYYY-MM-DD HH:mm:ss")
        If str发生时间_L <> "" And Mid(str发生时间_L, 1, 16) = Mid(str发生时间, 1, 16) Then
            '日期相同，秒数不同，且不是汇总数据行，则说明这些数据是一组，更新lngDemo列
            vsfHistory.TextMatrix(lngRow, mlngDate) = ""
            vsfHistory.TextMatrix(lngRow, mlngTime) = ""
            vsfHistory.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
            If lngRow - lngLastRow = Val(vsfHistory.TextMatrix(lngLastRow, mlngRowCount)) Then
                vsfHistory.TextMatrix(lngLastRow, mlngDemo) = 1
            End If
        Else
            lngLastRow = lngRow
        End If
        
        If lngRowCount > 1 Then
            '先增加空行
            vsfHistory.Rows = vsfHistory.Rows + lngRowCount - 1
            '从当前行的下一行开始，每行的位置+所增加的空白行数，保证新增的空白行从当前行的下一行开始
            For intData = vsfHistory.Rows - lngRowCount To lngRow + 1 Step -1
                vsfHistory.RowPosition(intData) = intData + lngRowCount - 1
            Next

            '循环处理当前行数据
            For lngCol = 0 To vsfHistory.Cols - 1
                If vsfHistory.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                    '循环赋值
                    For intData = 2 To lngRowCount
                        vsfHistory.TextMatrix(lngRow + intData - 1, lngCol) = vsfHistory.TextMatrix(lngRow, lngCol)
                    Next
                ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime) Then
                    '准备赋值
                    With txtLength
                        .Width = vsfHistory.ColWidth(lngCol)
                        .Text = Replace(Replace(Replace(vsfHistory.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        .FontName = vsfHistory.CellFontName
                        .FontSize = vsfHistory.CellFontSize
                        .FontBold = vsfHistory.CellFontBold
                        .FontItalic = vsfHistory.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)

                    If intDatas > 0 Then
                        '循环赋值
                        For intData = 0 To intDatas
                            vsfHistory.TextMatrix(lngRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        Next
                    End If
                ElseIf lngCol = mlngNoEditor Then
                        '将行值改为从1开始,比如有4行数据,就是4|1
                        For intData = 1 To lngRowCount
                            vsfHistory.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                        Next
                        '最后一行需要填写封闭签名
                        If mlngSignName > 0 Then vsfHistory.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = vsfHistory.TextMatrix(lngRow, mlngSignName)
                        If mlngSignTime > 0 Then vsfHistory.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = vsfHistory.TextMatrix(lngRow, mlngSignTime)
                        '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                        Call SingerShowType(vsfHistory, lngRow, lngRow + lngRowCount - 1)
                    Else
                End If
            Next
            '@实际数据行
'            '如果本页第一行的数据不全,则先将该记录第一行的主数据(日期,时间,签名)信息复制到
'            If lngRow = vsfhistory.FixedRows And lngRowCount <> lngRowCurrent Then
'                '固定复制显示日期时间与签名列
'                lngMax = lngRowCount - lngRowCurrent
'                If mlngDate > -1 Then vsfhistory.TextMatrix(lngRow + lngMax, mlngDate) = vsfhistory.TextMatrix(lngRow, mlngDate)
'                If mlngTime > -1 Then vsfhistory.TextMatrix(lngRow + lngMax, mlngTime) = vsfhistory.TextMatrix(lngRow, mlngTime)
'                if mlngOperator <>-1 then vsfhistory.TextMatrix(lngRow + lngMax, mlngOperator) = vsfhistory.TextMatrix(lngRow, mlngOperator)
'                if mlngOperator <>-1 then vsfhistory.TextMatrix(lngRow + lngMax, mlngsignname) = vsfhistory.TextMatrix(lngRow, mlngsignname)
'                '删除多余的行
'                For lngCol = 1 To lngMax
'                    vsfhistory.RemoveItem lngRow
'                Next
'            End If
'            lngRow = lngRow + lngRowCurrent - 1 '加上该记录在本页实际的行数
            '@实际数据行要注释下面这行代码
            lngRow = lngRow + lngRowCount - 1
        Else
            vsfHistory.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
        str发生时间_L = str发生时间
    Loop
    
    '63760:刘鹏飞,分组数据护士、签名人、签名时间的处理（同一个签名人始终显示一次）
    If mlngSingerType > 0 And vsfHistory.FixedRows <= vsfHistory.Rows - 1 Then
        lngPrintedRow = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        lngRow = vsfHistory.FixedRows
        Do While True
            lngStart = GetStartRowHistory(lngRow)
            lngRowCount = Val(vsfHistory.TextMatrix(lngStart, mlngRowCount))
            If lngRowCount <= 0 Then Exit Do
            
            If mlngSingerType = 3 Then '尾行签名
                strSignName = vsfHistory.TextMatrix(lngStart + lngRowCount - 1, lngPrintedRow)
            Else '首行签名或首尾签名
                strSignName = vsfHistory.TextMatrix(lngStart, lngPrintedRow)
            End If
            strSignName = FormatValue(strSignName)
            
            If Val(vsfHistory.TextMatrix(lngStart, mlngDemo)) = 1 And lngStart = lngRow And strSignName <> "" Then
                For lngRow = lngStart + lngRowCount To vsfHistory.Rows - 1
                    If lngRow = lngStart + lngRowCount Then
                    
                        If Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                        
                        lngRowCount = Val(vsfHistory.TextMatrix(lngRow, mlngRowCount))
                        If lngRowCount = 0 Then Exit For
                        
                        If mlngSingerType = 3 Then '尾行签名
                            If strSignName = FormatValue(vsfHistory.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) Then
                                If lngStart <= lngRow - 1 Then
                                    If mlngOperator <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(vsfHistory.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(vsfHistory.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow))
                                End If
                            End If
                        Else '首行签名或首尾签名
                            If strSignName = FormatValue(vsfHistory.TextMatrix(lngRow, lngPrintedRow)) Then
                                '首行签名或首尾签名都需要去掉下一条数据的首行,但首尾签名需要注意分组中的最后一条数据行数=1的情况
                                blnClear = True
                                If mlngSingerType = 2 And lngRowCount = 1 Then
                                    If lngRow + lngRowCount < vsfHistory.Rows Then
                                        If Val(vsfHistory.TextMatrix(lngRow + lngRowCount, mlngDemo)) <= 1 Then
                                            blnClear = False
                                        End If
                                    Else
                                        blnClear = False
                                    End If
                                End If
                                If blnClear Then
                                    If mlngOperator <> -1 Then vsfHistory.TextMatrix(lngRow, mlngOperator) = ""
                                    If mlngSignName <> -1 Then vsfHistory.TextMatrix(lngRow, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then vsfHistory.TextMatrix(lngRow, mlngSignTime) = ""
                                End If
                                '首尾签名还应该去掉上一条数据的尾行(上一行数据行数需要>1)
                                If mlngSingerType = 2 And lngStart < lngRow - 1 Then
                                    If mlngOperator <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(vsfHistory.TextMatrix(lngRow, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(vsfHistory.TextMatrix(lngRow, lngPrintedRow))
                                End If
                            End If
                        End If
                        
                        lngStart = lngRow
                    End If
                Next lngRow
            Else
                lngRow = lngStart + Val(vsfHistory.TextMatrix(lngStart, mlngRowCount))
            End If
            
            If lngRow > vsfHistory.Rows - 1 Then Exit Do
        Loop
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim stdPicFont As StdFont
    On Error GoTo ErrHand

    '设置护理记录单的格式
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '表头填写
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True

        '程序内部控制列隐藏\
        .ColHidden(c分组) = True
        .ColHidden(c文件ID) = True
        .ColHidden(c病人ID) = True
        .ColHidden(c主页ID) = True
        .ColHidden(c婴儿) = True
        .ColHidden(c血压频次) = (mstrBPItem = "")
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngActiveTime) = True
        .ColWidth(0) = 250
        .ColWidth(c姓名) = 1500
        .ColAlignment(c床号) = flexAlignRightCenter
        Set stdPicFont = picMain.Font
        Set picMain.Font = .Font
        .ColWidth(c血压频次) = (picMain.TextWidth("血") * 5)
        Set picMain.Font = stdPicFont
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&

        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        '设置固定列及选择列
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, c分组) = "分组"
        .TextMatrix(1, c分组) = "分组"
        .TextMatrix(2, c分组) = "分组"
        .TextMatrix(0, c文件ID) = "文件ID"
        .TextMatrix(1, c文件ID) = "文件ID"
        .TextMatrix(2, c文件ID) = "文件ID"
        .TextMatrix(0, c床号) = "床号"
        .TextMatrix(1, c床号) = "床号"
        .TextMatrix(2, c床号) = "床号"
        .TextMatrix(0, c姓名) = "姓名"
        .TextMatrix(1, c姓名) = "姓名"
        .TextMatrix(2, c姓名) = "姓名"
        .TextMatrix(0, c病人ID) = "病人ID"
        .TextMatrix(1, c病人ID) = "病人ID"
        .TextMatrix(2, c病人ID) = "病人ID"
        .TextMatrix(0, c主页ID) = "主页ID"
        .TextMatrix(1, c主页ID) = "主页ID"
        .TextMatrix(2, c主页ID) = "主页ID"
        .TextMatrix(0, c婴儿) = "婴儿"
        .TextMatrix(1, c婴儿) = "婴儿"
        .TextMatrix(2, c婴儿) = "婴儿"
        .TextMatrix(0, c血压频次) = "血压频次"
        .TextMatrix(1, c血压频次) = "血压频次"
        .TextMatrix(2, c血压频次) = "血压频次"

        '列宽设置
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next

        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next

        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
        Else
            '得到第一行的超出行
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            '加上最后一行的超出行
            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
        End If

        Call FillPage
        Call WriteColor
        
        '可能固定行的行高不正确需要自动调整下
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        '将非固定行的行高设置为最小行高
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .ROW = .FixedRows
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormatHistory(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    On Error GoTo ErrHand

    '设置护理记录单的格式
    With vsfHistory
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '表头填写
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True

        '程序内部控制列隐藏
        .ColHidden(c分组) = True
        .ColHidden(c文件ID) = True
        .ColHidden(c病人ID) = True
        .ColHidden(c主页ID) = True
        .ColHidden(c婴儿) = True
        .ColHidden(c姓名) = True
        .ColHidden(c床号) = True
        .ColHidden(c血压频次) = True
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngActiveTime) = True
        .ColWidth(0) = 250

        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&

        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        '设置固定列及选择列
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, c分组) = "分组"
        .TextMatrix(1, c分组) = "分组"
        .TextMatrix(2, c分组) = "分组"
        .TextMatrix(0, c文件ID) = "文件ID"
        .TextMatrix(1, c文件ID) = "文件ID"
        .TextMatrix(2, c文件ID) = "文件ID"
        .TextMatrix(0, c床号) = "床号"
        .TextMatrix(1, c床号) = "床号"
        .TextMatrix(2, c床号) = "床号"
        .TextMatrix(0, c姓名) = "姓名"
        .TextMatrix(1, c姓名) = "姓名"
        .TextMatrix(2, c姓名) = "姓名"
        .TextMatrix(0, c病人ID) = "病人ID"
        .TextMatrix(1, c病人ID) = "病人ID"
        .TextMatrix(2, c病人ID) = "病人ID"
        .TextMatrix(0, c主页ID) = "主页ID"
        .TextMatrix(1, c主页ID) = "主页ID"
        .TextMatrix(2, c主页ID) = "主页ID"
        .TextMatrix(0, c婴儿) = "婴儿"
        .TextMatrix(1, c婴儿) = "婴儿"
        .TextMatrix(2, c婴儿) = "婴儿"
        .TextMatrix(0, c血压频次) = "婴儿"
        .TextMatrix(1, c血压频次) = "婴儿"
        .TextMatrix(2, c血压频次) = "婴儿"

        '列宽设置
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next

        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next

        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
        Else
            '得到第一行的超出行
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            '加上最后一行的超出行
            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
        End If
        
        Call PreTendMutilRows
        Call WriteColorHistory
        
        '可能固定行的行高不正确需要自动调整下
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        '将非固定行的行高设置为最小行高
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    '晚班以红色显示，同时将非起始行设置为NoCheckBox，设置图标
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 0) <> "" Then
                '晚班以红色显示
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If

            '将非起始行设置为NoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                    '设置图标
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(审签).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(签名).Picture
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    Call SetActiveColColor
End Sub

Private Sub WriteColorHistory()
    Dim blnTag As Boolean
    Dim lngCount As Long
    '晚班以红色显示，同时将非起始行设置为NoCheckBox，设置图标
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 0) <> "" Then
                '晚班以红色显示
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If

            '将非起始行设置为NoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                    '设置图标
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(审签).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(签名).Picture
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    Call SetActiveColColor
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long

    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    VsfData.Move lngScaleLeft + 210, lblTitle.Top + lblTitle.Height + 300, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub InitEnv()
    Dim curDate As Date
    Dim intDay As Integer
    Dim rs As New ADODB.Recordset
    Dim blntype As Boolean
    On Error GoTo ErrHand
    
    glngHours = Val(zlDatabase.GetPara("数据补录时限", glngSys))
    mintChange = Val(zlDatabase.GetPara("最近转出天数", glngSys, p住院护士站, 7))
    '出院病人时间范围
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, p住院护士站, 7))
    mdtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, p住院护士站, 30))
    mdtOutbegin = Format(mdtOutEnd - intDay, "yyyy-MM-dd 00:00:00")
    
    blntype = Val(GetSetting("ZLSOFT", "私有模块\usrTendFileMutilEditor\" & gstrUserName, "Value")) = 0
    If blntype Then
        optLevel(0).Value = True
    Else
        optLevel(1).Value = True
    End If
    
    '打开现存在的所有护理记录项目
    gstrSQL = " Select   项目序号,upper(项目名称) AS 项目名称,分组名,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,缺省值,护理等级,应用方式,说明" & _
              " From 护理记录项目 B" & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    
    '打开现存在的所有护理记录项目
    gstrSQL = "Select   项目序号,记录法 from 体温记录项目 order by 项目序号 "
    Set mrsTemperItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有体温记录项目")
    
    '提取常用体温说明
    gstrSQL = "select 编码,名称,简码 from  常用体温说明 order  by 编码"
    Set mrsUsual = zlDatabase.OpenSQLRecord(gstrSQL, "提取常用体温说明")
    
    '提取除体温单和产程图外的护理文件清单
    gstrSQL = " Select  ID,名称 FROM 病历文件列表 " & vbNewLine & _
              " WHERE 种类=3 AND DECODE(保留,-1,0,1,0,1)=1 AND (通用 =1 OR (通用=2 And ID IN (Select 文件ID FROM 病历应用科室 Where 科室ID=[1])))" & vbNewLine & _
              " ORDER BY 编号 "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取除体温单外的护理文件清单", mlng病区ID)
    With rs
        cbo护理文件格式.Clear
        Do While Not .EOF
            cbo护理文件格式.AddItem !名称
            cbo护理文件格式.ItemData(cbo护理文件格式.NewIndex) = !ID
            .MoveNext
        Loop
        If .RecordCount <> 0 Then cbo护理文件格式.ListIndex = 0
    End With
    
    '提取当前病区下的所有科室
    gstrSQL = " Select distinct B.ID,B.编码||'-'||B.名称 AS 科室" & _
              " From 病区科室对应 A,部门表 B,部门人员 C,人员表 D" & _
              " Where A.科室ID = b.ID And A.科室ID=C.部门ID And C.人员ID=D.ID And A.病区ID = [1]" & _
              IIf(InStr(1, mstrPrivs, "当前病区") <> 0, "", " And D.ID=[2]") & _
              " Order by B.编码||'-'||B.名称"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前病区下的所有科室", mlng病区ID, glngUserId)
    With cbo科室
        .Clear
        If InStr(1, mstrPrivs, "当前病区") <> 0 Then
            .AddItem "所有科室"
            .ItemData(.NewIndex) = -1
        End If
        Do While Not rs.EOF
            .AddItem rs!科室
            .ItemData(.NewIndex) = rs!ID
            rs.MoveNext
        Loop
        If rs.RecordCount <> 0 Then .ListIndex = 0
    End With
    
    '读取绑定的测血压项目
    mstrBPItem = ""
    gstrSQL = "Select a.Xh 项目" & vbNewLine & _
        " From 病区公告栏样式 p, Xmltable('/ITEMLIST/ITEM/XH' Passing p.诊疗项目 Columns Xh Varchar2(256) Path '/XH') a" & vbNewLine & _
        " Where p.病区id = [1] And 名称 = '测血压列表'"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "病区公告栏样式", mlng病区ID)
    Do While Not rs.EOF
        If Not IsNull(rs!项目) Then
            mstrBPItem = mstrBPItem & "," & Val(rs!项目)
        End If
        rs.MoveNext
    Loop
    mstrBPItem = Mid(mstrBPItem, 2)
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCol As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, arrCorrelative(), strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand

    strColumns = mstrColumns
    If Not mblnInit Then
        '初始化内存记录集(未对应项目的列为活动项目,其它列均为固定项)
        strFields = "列," & adDouble & ",18|序号," & adDouble & ",2|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|固定," & adDouble & ",2|格式," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "列|序号|项目序号|项目名称|固定|格式"
    End If

    '加入列定义
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCol = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '如果已设置以传入值为准'否则找不到项目就是活动项目
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If

            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "项目名称='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!项目序号
                    If Not blnSet Then intImmovable = 1   '固定不允许修改
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0

                    '记录特殊列
                    Select Case strName
                    Case "日期"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "时间"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "护士"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "签名人"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "签名时间"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        '整理分类汇总关联列信息
        arrCorrelative = Array()
        arrColumn = Split(mstrColCorrelative, "|")
        For i = 0 To UBound(arrColumn)
            arrItem = Split(arrColumn(i), ";")
            If UBound(arrItem) = 1 Then
                mrsSelItems.Filter = "列=" & Val(arrItem(0))
                If mrsSelItems.RecordCount = 1 Then
                    ReDim Preserve arrCorrelative(UBound(arrCorrelative) + 1)
                    arrCorrelative(UBound(arrCorrelative)) = Val(arrItem(0)) & "," & mrsSelItems!项目序号 & ";" & CStr(arrItem(1))
                End If
            End If
        Next i
        If UBound(arrCorrelative) = -1 Then
            mstrColCorrelative = ""
        Else
            mstrColCorrelative = Join(arrCorrelative, "|")
        End If
'        mstrColImCorrelative = mstrColCorrelative
'        If mblnCorrelative = False Then mstrColCorrelative = ""
        mrsSelItems.Filter = ""
        '加入程序内部控制列(列是在读取数据后绑定时增加的,此时只有预处理下)
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '加上隐藏列
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngActiveTime = mlngRowCurrent + 1
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If

    mrsItems.Filter = 0
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function SignMe() As Boolean
    Dim blnSign As Boolean          '是否签名成功
    Dim blnRefresh As Boolean
    Dim strTime As String
    Dim strSignTime As String       '保证所有签名的签名时间一致,便于取消签名时按签名时间统一取消
    Dim str状态 As String           '保存签名选项,避免循环签名时不停的弹出签名窗口
    Dim str行错误 As String
    Dim str错误 As String
    Dim intRow As Integer, intRows As Integer
    On Error GoTo ErrHand
    
    '普签:对当前界面的所有数据进行签名
    '准备签名
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    intRows = VsfData.Rows - 1
    For intRow = VsfData.FixedRows To intRows
        If Val(VsfData.TextMatrix(intRow, mlngRecord)) > 0 And VsfData.TextMatrix(intRow, mlngSigner) = "" Then
            str行错误 = ""
'            If InStr(1, VsfData.TextMatrix(intRow, mlngDate), "/") <> 0 Then
'                strTime = Format(Now, "yyyy") & "-" & ToStandDate(VsfData.TextMatrix(intRow, mlngDate)) & " " & VsfData.TextMatrix(intRow, mlngTime) & ":00"
'            Else
'                strTime = VsfData.TextMatrix(intRow, mlngDate) & " " & VsfData.TextMatrix(intRow, mlngTime) & ":00"
'            End If
            strTime = VsfData.TextMatrix(intRow, mlngActiveTime)
            
            blnSign = SignName(intRow, strTime, strSignTime, str状态, str行错误)
            If Not blnSign Then Exit For
            If Not blnRefresh Then blnRefresh = blnSign
            If str行错误 <> "" Then
                str错误 = str错误 & vbCrLf & "发生时间=[" & strTime & "]" & str行错误
            End If
        End If
    Next
    
    SignMe = blnRefresh
    mblnSigned = blnRefresh
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe()
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strTime As String
    Dim blnTrans As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim clsSign As Object
    Dim strSQLTime() As String, intPos As Integer
    Dim arrRow
    
    On Error GoTo ErrHand
    ReDim Preserve strSQLTime(1 To 1)
    arrRow = Array()
    '首先最后一次是本人的签名，根据当前选择数据的签名时间，批量取消签名

    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 And VsfData.TextMatrix(lngRow, mlngSigner) <> "" Then
            If Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) > 0 Then
                '数字签名验证，只验证一次
                If clsSign Is Nothing Then
                    If gobjESign Is Nothing Then
                        On Error Resume Next
                        Set gobjESign = CreateObject("zl9ESign.clsESign")
                        If Err <> 0 Then Err.Clear
                        On Error GoTo 0
                        If Not gobjESign Is Nothing Then Call gobjESign.Initialize(gcnOracle, glngSys)
                    End If
                    Set clsSign = gobjESign
    
                    If Not clsSign Is Nothing Then
                        If Not clsSign.CheckCertificate(gstrDBUser) Then
                            Exit Sub
                        End If
                    Else
                        RaiseEvent AfterRowColChange("电子签名部件未能正确安装，回退操作不能继续！", True)
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
            '提取发生时间
'            If InStr(1, VsfData.TextMatrix(lngRow, mlngDate), "/") <> 0 Then
'                strTime = Format(Now, "yyyy") & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
'            Else
'                strTime = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
'            End If
            strTime = VsfData.TextMatrix(lngRow, mlngActiveTime)
            '取消签名
            gstrSQL = "ZL_病人护理数据_UNSIGNNAME("
            gstrSQL = gstrSQL & VsfData.TextMatrix(lngRow, c文件ID) & ","
            gstrSQL = gstrSQL & "To_Date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),1)"
            
            strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
            
            ReDim Preserve arrRow(UBound(arrRow) + 1)
            arrRow(UBound(arrRow)) = lngRow
            
        End If
    Next
    
    gcnOracle.BeginTrans
    blnTrans = True
    For intPos = 1 To UBound(strSQLTime)
        If strSQLTime(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "执行取消签名")
        End If
    Next intPos
    gcnOracle.CommitTrans
    blnTrans = False
    
    '更改图标
    ''更改图标
    For intPos = 0 To UBound(arrRow)
        lngRow = Val(arrRow(intPos))
        VsfData.Cell(flexcpPicture, lngRow, 0) = Nothing
        If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = ""
        VsfData.TextMatrix(lngRow, mlngSignLevel) = 0
        VsfData.TextMatrix(lngRow, mlngSigner) = ""
        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = ""
        '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
        Call SingerShowType(VsfData, lngRow, lngRow + Val(VsfData.TextMatrix(lngRow, mlngRowCount)) - 1, True)
    Next intPos
    mblnSigned = False
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal intRow As Integer, ByVal strStart As String, ByVal strSignTime As String, _
    str状态 As String, Optional str错误 As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cTendSign
    Dim strSource As String             '审签源数据串
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset

    On Error GoTo ErrHand

    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""

    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
   gstrSQL = " Select  a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.记录时间  " & _
              " From 病人护理明细 a,病人护理数据 b,病人护理文件 c " & _
              " Where a.记录id=b.ID And B.汇总类别=0 AND MOD(A.记录类型,10)<>5 And b.文件ID=c.ID And a.终止版本 Is Null And C.ID=[1] And b.发生时间=[2]" & _
              " Order by a.项目序号"
    Call SQLDIY(gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "获取要签名的内容", Val(VsfData.TextMatrix(intRow, c文件ID)), CDate(strStart))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    If strSource = "" Then
        RaiseEvent AfterRowColChange("当前没有需要签名的信息！", True)
        Exit Function
    End If

    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    '76223:刘鹏飞,2012-09-13,电子签名添加时间戳信息
    Set oSign = frmTendFileSign.ShowMe(Me, mstrPrivs, Val(VsfData.TextMatrix(intRow, c文件ID)), Val(VsfData.TextMatrix(intRow, c病人ID)), Val(VsfData.TextMatrix(intRow, c主页ID)), mlng病区ID, 未定义, strSource, False, str状态, str错误)
    On Error GoTo ErrHand

    If Not oSign Is Nothing Then
        gstrSQL = "ZL_病人护理数据_SIGNNAME("
        gstrSQL = gstrSQL & Val(VsfData.TextMatrix(intRow, c文件ID)) & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),0,"
        gstrSQL = gstrSQL & "'" & oSign.姓名 & "',"
        gstrSQL = gstrSQL & "'" & oSign.签名信息 & "'," & oSign.签名级别 & ","
        gstrSQL = gstrSQL & oSign.证书ID & ","
        gstrSQL = gstrSQL & oSign.签名方式 & ",'" & oSign.时间戳 & "',0,'" & oSign.时间戳信息 & "')"
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, "执行签名")
        SignName = True
        
        VsfData.TextMatrix(intRow, mlngSignLevel) = oSign.证书ID
        VsfData.TextMatrix(intRow, mlngSigner) = "SignName"
        '更新图标
        VsfData.Cell(flexcpPicture, intRow, 0) = imgRow.ListImages(签名).Picture
        If mlngSignName <> -1 Then VsfData.TextMatrix(intRow, mlngSignName) = gstrUserName
        If mlngSignTime <> -1 Then VsfData.TextMatrix(intRow, mlngSignTime) = oSign.时间戳
        '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
        Call SingerShowType(VsfData, intRow, intRow + Val(VsfData.TextMatrix(intRow, mlngRowCount)) - 1)
    End If

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    mblnChange = False
    mintType = -1
    
    '内存记录集清空
    mrsCellMap.Filter = 0
    If mrsCellMap.RecordCount <> 0 Then mrsCellMap.MoveFirst
    Do While True
        If mrsCellMap.EOF Then Exit Do
        mrsCellMap.Delete
        mrsCellMap.Update
        mrsCellMap.MoveNext
    Loop
    
    Call DataMap_Restore
    
    Call InitCons
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function

    mblnShow = False
    Call InitCons
    SaveME = True
    RaiseEvent AfterRowColChange("保存成功！", False)
    
    '--48659:刘鹏飞,2012-09-14,添加字段'说明'
    RaiseEvent ShowTipInfo(VsfData, "", True)
    
    If VsfData.ROW < VsfData.Rows And mlngDate < VsfData.Cols Then
        VsfData.Select VsfData.ROW, mlngDate
    End If
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngDeptID As Long, Optional ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngDeptID           要显示护理记录的科室
    '返回： 无
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    Err = 0

    mblnInit = False
    mblnHistory = False
    
    mint页码 = 1
    mlng病区ID = lngDeptID
    mstrPrivs = strPrivs
    mblnBlowup = (bytSize = 1) '(zlDatabase.GetPara("护理文件显示模式", glngSys, 1255, 0) = 1)
    Set mfrmParent = frmParent

    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm")
    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    mlngSingerType = Val(zlDatabase.GetPara("护士、签名列显示模式", glngSys, 1255, "2"))
    If InStr(1, ",0,1,2,3,", "," & mlngSingerType & ",") = 0 Then mlngSingerType = 2
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '初始化环境
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    
    Call InitVariable
    Call InitCons
    
     '--48659:刘鹏飞,2012-09-14,添加字段'说明'
    RaiseEvent ShowTipInfo(VsfData, "", True)
    
    If cbo科室.ListCount = 0 Then
        MsgBox "您不属于当前病区的任何科室，不能使用该功能！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ReSetFontSize
    ShowMe = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = IIf(mblnBlowup = True, 12, 9)
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "宋体"
    For Each objCtrl In UserControl.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            Select Case UCase(objCtrl.Name)
            Case UCase("lbl文件格式"), UCase("lbl科室")
                objCtrl.FontSize = bytFontSize
                objCtrl.Height = TextHeight("刘") + 20
            End Select
        Case UCase("ComboBox")
            Select Case UCase(objCtrl.Name)
            Case UCase("cbo护理文件格式"), UCase("cbo科室")
                objCtrl.FontSize = bytFontSize
            End Select
        Case UCase("CheckBox")
            Select Case UCase(objCtrl.Name)
            Case UCase("chk出科"), UCase("chk出院")
                objCtrl.FontSize = bytFontSize
                objCtrl.Width = TextWidth("刘鹏" & objCtrl.Caption) - TextWidth("刘") / 3
            End Select
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("CommandButton")
            If UCase(objCtrl.Name) = UCase("cmd刷新") Then
                objCtrl.FontSize = bytFontSize
                objCtrl.Width = TextWidth(" " & IIf(objCtrl.Caption = "", "  ", objCtrl.Caption) & " ")
            End If
        End Select
    Next
    
    '移动控件位置
    cbo护理文件格式.Top = (pic过滤条件.Height - cbo护理文件格式.Height) \ 2
    lbl文件格式.Left = 60
    cbo护理文件格式.Left = lbl文件格式.Left + lbl文件格式.Width + TextWidth("刘") / 2
    lbl文件格式.Top = cbo护理文件格式.Top + (cbo护理文件格式.Height - lbl文件格式.Height) \ 2
    lbl科室.Left = cbo护理文件格式.Left + cbo护理文件格式.Width + TextWidth("刘")
    lbl科室.Top = lbl文件格式.Top
    cbo科室.Left = lbl科室.Left + lbl科室.Width + TextWidth("刘") / 2
    cbo科室.Top = cbo护理文件格式.Top
    chk出科.Left = cbo科室.Left + cbo科室.Width + TextWidth("刘")
    chk出科.Top = lbl文件格式.Top
    chk出院.Left = chk出科.Left + chk出科.Width + TextWidth("刘") / 2
    chk出院.Top = chk出科.Top
    cmd刷新.Height = cbo科室.Height + 15
    cmd刷新.Left = chk出院.Left + chk出院.Width + TextWidth("刘")
    lblEntry.Left = cmd刷新.Left + cmd刷新.Width + TextWidth("刘")
    lblEntry.Top = cmd刷新.Top + (cmd刷新.Height - lblEntry.Height) \ 2
    lblEntry.Height = cmd刷新.Height
    optLevel(0).Left = lblEntry.Left + lblEntry.Width + TextWidth("刘") / 2
    optLevel(0).Top = cmd刷新.Top + 10
    optLevel(0).Height = cmd刷新.Height
    optLevel(1).Left = optLevel(0).Left + optLevel(0).Width + TextWidth("刘")
    optLevel(1).Top = optLevel(0).Top
    optLevel(1).Height = optLevel(0).Height
    
    
    pic过滤条件.Width = optLevel(1).Left + optLevel(1).Width + 50
End Sub

Private Function CheckFlip() As Boolean
    Dim blnExit As Boolean, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim lngOldRow As Long, lngOldCol As Long, lngEditCol As Long, blnShow As Boolean
    Dim strDate As String, strInfo As String
    
    On Error GoTo ErrHand
    
    '隐藏编辑控件
    lngOldRow = VsfData.ROW: lngOldCol = VsfData.COL
    
    blnShow = mblnShow
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
        lstSelect(2).Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
        If mintType = 1 Then
            txtLst.Visible = False
            PicLst.Visible = False
        End If
    Case 4, 5
        picDouble.Visible = False
        lstSelect(2).Visible = False
    Case 6
        picMutilInput.Visible = False
        lstSelect(2).Visible = False
    Case 7
        picDoubleChoose.Visible = False
    End Select
    
    cmdWord.Visible = False
    mintType = -1
    mblnShow = False
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) > 1 And Trim(VsfData.TextMatrix(lngRow, mlngSigner)) = "" And VsfData.RowHidden(lngRow) = False Then
            blnNULL = True
            For lngCol = mlngTime + 1 To lngCols - 1
                If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
                    If VsfData.TextMatrix(lngRow, lngCol) <> "" And Not (IsDiagonal(lngCol) And InStr(1, VsfData.TextMatrix(lngRow, lngCol), "/") <> 0) Then
                        blnNULL = False
                        Exit For
                    End If
                End If
            Next
            If blnNULL = True And cbsThis.FindControl(xtpControlButton, conMenu_Edit_Clear).Enabled = True Then
                VsfData.ROW = lngRow
                Call cbsThis_Execute(cbsThis.FindControl(xtpControlButton, conMenu_Edit_Clear))
            End If
        End If
    Next
    
    '页面切换前检查：日期时间正确才允许继续，这样在保存时就不必再检查其它页面的数据了（其它数据在录入时已经进行了检查，此处略过）
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>" & mlngTime
        If mrsCellMap.RecordCount = 0 And Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
            mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>=" & mlngDate
        End If
        If mrsCellMap.RecordCount <> 0 Then
            If Not VsfData.RowHidden(lngRow) And Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                blnExit = (VsfData.TextMatrix(lngRow, mlngDate) = "" Or VsfData.TextMatrix(lngRow, mlngTime) = "")
                If blnExit Then
                    mrsCellMap.Filter = 0
                    If VsfData.TextMatrix(lngRow, mlngDate) = "" Then
                        lngCol = mlngDate
                    Else
                        lngCol = mlngTime
                    End If
                    VsfData.ROW = lngRow: VsfData.COL = lngCol
                    mblnShow = True: Call VsfData_EnterCell
                    If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                    CheckFlip = False
                    RaiseEvent AfterRowColChange("请补充日期时间！", True)
                    Exit Function
                Else
                    '日期不为空将检查日期的合法性
                    If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
                        If mblnDateAd Then
                            strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 4) & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                        Else
                            strDate = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                        End If
                        strDate = Format(strDate, "YYYY-MM-DD HH:mm")
                        blnExit = (strDate = Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "YYYY-MM-DD HH:mm"))
                    End If
                    If blnExit = False Then
                        VsfData.ROW = lngRow: VsfData.COL = mlngTime
                        If Not CheckDateTime(VsfData.TextMatrix(VsfData.ROW, VsfData.COL), strInfo) Then
                            mrsCellMap.Filter = ""
                            mblnShow = True: Call VsfData_EnterCell
                            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                            RaiseEvent AfterRowColChange(strInfo, True)
                            CheckFlip = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    
     '检查新增的子分组数据，如果汇总行已经签名则提示（只处理在原有数据新增的分组数据，因新增的分组数据上面已经检查）
    strDate = ""
    For lngRow = VsfData.FixedRows To lngRows
        If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
            If Not Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) >= 1 Then strDate = ""
            If Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) = 1 And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRecord))) > 0 Then
                If mblnDateAd Then
                    strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 4) & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                Else
                    strDate = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                End If
                strDate = Format(strDate, "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(strDate) And Not VsfData.RowHidden(lngRow) And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) > 1 And _
                 Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRecord))) <= 0 Then
                mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>" & mlngTime
                If mrsCellMap.RecordCount > 0 Then
                    lngEditCol = 0
                    If CheckCollectIsData(lngRow, 1, lngEditCol) = True Then
                        If ISCollectSigned(Val(VsfData.TextMatrix(lngRow, c文件ID)), Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                            VsfData.ROW = lngRow: VsfData.COL = lngEditCol
                            strInfo = "您新增的分组数据所对应的汇总行数据已签名，不允许再添加新的汇总列数据！"
                            mrsCellMap.Filter = ""
                            mblnShow = True: Call VsfData_EnterCell
                            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                            RaiseEvent AfterRowColChange(strInfo, True)
                            CheckFlip = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    mblnShow = blnShow
    VsfData.Select lngOldRow, lngOldCol
    mrsCellMap.Filter = 0
    CheckFlip = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsCellMap.Filter = 0
End Function

Private Function CheckData() As Boolean
    Dim intLevel As Integer
    Dim lngPage As Long
    On Error GoTo ErrHand
    '检查数据

    '如果修改了数据而日期时间不全则提示（数据合法性在录入时已经检查）
'    Call OutputRsData(mrsCellMap)
'    Call OutputRsData(mrsDataMap)
    If Not CheckFlip Then Exit Function

    CheckData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim arrValue, arrOrder, arrPart, arrCollect
    Dim strSQL() As String, strSQLTime() As String, strCollectSQL() As String
    Dim intAllow As Integer
    Dim lngRecord As Long
    Dim blnTrans As Boolean, blnSaved As Boolean, blnDel As Boolean
    Dim intPos As Integer, intMax As Integer, intPage As Integer, intRow As Integer, intUsedRows As Integer
    Dim strReturn As String, strCellData As String, strPart As String
    Dim strMonth As String, strDay As String
    Dim strDate As String, strTime As String, strTemp As String
    Dim strDatetime As String, strCurrDate As String, strDays As String
    Dim strSaveRows As String
    Dim str发生时间_L As String, str发生时间 As String, str文件ID As String
    Dim lngLastRow As Long
    
    ReDim Preserve strSQL(1 To 1)
    ReDim Preserve strSQLTime(1 To 1) '发生时间变动SQL数组
    Dim rsTemp As New ADODB.Recordset, rsTime As New ADODB.Recordset, rsTimeCur As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    On Error GoTo ErrHand
    
    strFileds = "ID," & adDouble & ",18|文件ID," & adDouble & ",18|时间," & adDate & ",20|发生时间," & adDate & ",20|标记," & adInteger & ",1"
    Call Record_Init(rsTime, strFileds)
    Call Record_Init(rsTimeCur, strFileds)
    
    '同行多列循环调用：ZL_病人护理数据_UPDATE
    '下一行前调用：
    '   1、ZL_病人护理数据_SYNCHRO，同步数据到体温单与护理记录单中，需要记录删除的明细ID串
    '   2、ZL_病人护理打印_UPDATE，完成打印数据解析
    '删除项目需记录，删除行也需要记录
    '修改数据的同步就将该行数据对应的日期与时间保存到mrsCellMap中

'    objStream.WriteLine (Now & "产生保存SQL")
    intAllow = IIf(InStr(mstrPrivs, "他人护理记录") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")

    With mrsCellMap
        '将有效数据过滤出来:记录ID>0的历史数据+新增的有效数据
        .Filter = "记录ID>0 or (记录ID=0 And 删除=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If intRow <> !行号 Then
endWork:
                If intRow > 0 Then
                    blnDel = VsfData.RowHidden(intRow)
                    intUsedRows = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0))
                End If

                If blnSaved Then
                    strSaveRows = strSaveRows & "," & intRow
                     
                    '完成打印数据解析
'                    文件ID_IN IN 病人护理打印.文件ID%TYPE,
'                    发生时间_IN IN 病人护理打印.发生时间%TYPE,
'                    行数_IN IN 病人护理打印.行数%TYPE,
'                    删除_IN Number:=0
                    gstrSQL = "ZL_病人护理打印_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & intUsedRows & "," & IIf(blnDel, "1", "0") & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL

                    '只要修改过数据,必然会执行打印解析,因此在这里进行汇总日期的处理
                    If InStr(1, "," & strDays & ",", "," & Mid(strDatetime, 1, 10) & ",") = 0 Then
                        '同步更新明天的汇总(夜班,全天汇总跨天的处理)
                        strDays = strDays & "," & Mid(strDatetime, 1, 10)
                        gstrSQL = "ZL_汇总数据_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",'" & Mid(strDatetime, 1, 10) & "')"
                        strSQL(ReDimArray(strSQL)) = gstrSQL

                        strTemp = Format(DateAdd("d", 1, CDate(strDatetime)), "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                            strDays = strDays & "," & strTemp
                            gstrSQL = "ZL_汇总数据_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",'" & strTemp & "')"
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        End If
                    End If

                    blnSaved = False
                    If .EOF Then Exit Do
                End If

                '赋初值
                intPage = !页号
                intRow = !行号
                strDate = ""
                strDatetime = ""
                lngRecord = NVL(!记录ID, 0)
            End If

            If !列号 = mlngDate Then
                If NVL(!汇总, 0) = 1 Then
                    arrCollect = Split(!数据, ";")
                    strDatetime = arrCollect(3)
                '    文件ID_IN IN 病人护理数据.文件ID%TYPE,
                '    发生时间_IN IN 病人护理数据.发生时间%TYPE,
                '    汇总类别_IN IN 病人护理数据.汇总类别%TYPE,
                '    汇总文本_IN IN 病人护理数据.汇总文本%TYPE,
                '    汇总标记_IN IN 病人护理数据.汇总标记%TYPE,
                '    删除_IN Number:=0
                    gstrSQL = "ZL_病人护理数据_COLLECT(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",to_date('" & arrCollect(3) & "','yyyy-MM-dd hh24:mi:ss')," & _
                            Val(arrCollect(1)) & ",'" & arrCollect(0) & "'," & Val(arrCollect(2)) & "," & !删除 & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    blnSaved = True
                Else
                    strDate = NVL(!数据)
                    If strDate <> "" Then
                        If mblnDateAd Then
                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                            '检查是否翻年后编辑之前的时间(一个月的限制)
                            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                                strDate = DateAdd("yyyy", -1, CDate(strDate))
                            End If
                        Else
                            strDate = Format(strDate, "yyyy-MM-dd")
                        End If
                    End If
                End If
            ElseIf !列号 = mlngTime Then
                strTime = NVL(!数据)
                If strDatetime = "" Then
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                End If
                '处理分组数据，保存时与普通数据无区别，只是秒数+
                If Val(NVL(!部位)) >= 1 Then
                    'strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!部位), "0") & Val(!部位) - 1
                    strDatetime = DateAdd("S", Val(!部位) - 1, CDate(strDatetime))
                End If
                If lngRecord <> 0 Then
                    '更新发生时间
'                    gstrSQL = "Zl_病人护理数据_发生时间(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
'                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    strValues = lngRecord & "|" & Val(VsfData.TextMatrix(intRow, c文件ID)) & "|" & Format(strDatetime, "YYYY-MM-DD HH:mm:ss") & "|" & Format(VsfData.TextMatrix(intRow, mlngActiveTime), "YYYY-MM-DD HH:mm:ss") & "|0"
                    Call Record_Update(rsTime, "ID|文件ID|时间|发生时间|标记", strValues, "ID|" & lngRecord)
                    Call Record_Update(rsTimeCur, "ID|文件ID|时间|发生时间|标记", strValues, "ID|" & lngRecord)
                    blnSaved = True
                End If
            Else
                If !列号 > mlngTime Then
                    '取指定单元格的数据
                    strCellData = NVL(!数据)
                    strPart = NVL(!部位)
                    strReturn = ShowInput(!列号, strCellData, True)
                    'strOrders格式：项目序号,项目序号...
                    'strValues格式：值'值'值...
                    arrOrder = Split(Split(strReturn, "||")(0), ",")
                    arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                    arrPart = Split(strPart & "/////", "/")

                    intMax = UBound(arrOrder)
                    For intPos = 0 To intMax
                        If Not (Val(VsfData.TextMatrix(intRow, mlngRecord)) = 0 And arrValue(intPos) = "") Then
    '                    文件ID_IN IN 病人护理数据.文件ID%TYPE,
    '                    发生时间_IN IN 病人护理数据.发生时间%TYPE,
    '                    记录类型_IN IN 病人护理明细.记录类型%TYPE,          --护理项目=1，上标说明=2，手术日标记=4，签名记录=5，下标说明=6，入出量汇总=9
    '                    项目序号_IN IN 病人护理明细.项目序号%TYPE,          --护理项目的序号，非护理项目固定为0
    '                    记录内容_IN IN 病人护理明细.记录内容%TYPE := NULL,  --记录内容，如果内容为空，即清除以前的内容；37或38/37
    '                    体温部位_IN IN 病人护理明细.体温部位%TYPE := NULL,
    '                    他人记录_IN IN NUMBER := 1,
                        gstrSQL = "ZL_病人护理数据_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & intAllow & ",0,0,NULL,NULL,NULL,'" & NVL(!标记) & "')"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        blnSaved = True
                        End If
                    Next
                    mrsItems.Filter = 0
                End If
            End If

            .MoveNext
        Loop

        If blnSaved Then GoTo endWork
    End With
    
    '更新数据发生时间，对于分组数据中间某行数据行数变化会引起后面本分组数据的其他分组数据时间发生变化。如(增加数据行)：
    ',ID:403,时间:2012/5/8 18:23:00,发生时间:2012/5/8 18:23:00
    ',ID:407,时间:2012/5/8 18:23:02,发生时间:2012/5/8 18:23:01
    ',ID:517,时间:2012/5/8 18:23:03,发生时间:2012/5/8 18:23:02
    '需要先更新最后一行发生时间：如(减少数据行):
    ',ID:403,时间:2012/5/8 18:23:00,发生时间:2012/5/8 18:23:00
    ',ID:407,时间:2012/5/8 18:23:01,发生时间:2012/5/8 18:23:02
    ',ID:517,时间:2012/5/8 18:23:02,发生时间:2012/5/8 18:23:03
    '需要从前往后更新
    strDays = ""
    rsTime.Filter = ""
    'Call OutputRsData(rsTime)
    rsTime.Sort = "时间 DESC"
    Do While Not rsTime.EOF
        If InStr(1, "," & strDays & ",", "," & rsTime!ID & ",") = 0 Then
            rsTimeCur.Filter = "发生时间='" & Format(rsTime!时间, "YYYY-MM-DD HH:mm:ss") & "'And 标记=0 And ID<>" & Val(rsTime!ID)
            If rsTimeCur.RecordCount > 0 Then
                lngRecord = rsTimeCur!ID
                gstrSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!时间, "YYYY-MM-DD HH:mm:ss"), lngRecord, Val(rsTime!文件ID))
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                rsTimeCur.Filter = ""
                Call Record_Update(rsTimeCur, "标记", "1", "ID|" & lngRecord)
                strDays = IIf(strDays = "", "", ",") & lngRecord
                GoTo ErrLoop
            Else
                lngRecord = rsTime!ID
                gstrSQL = "Zl_病人护理数据_发生时间(" & rsTime!ID & ",to_date('" & rsTime!时间 & "','yyyy-MM-dd hh24:mi:ss'))"
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                rsTimeCur.Filter = ""
                Call Record_Update(rsTimeCur, "标记", "1", "ID|" & lngRecord)
                strDays = IIf(strDays = "", "", ",") & lngRecord
            End If
        End If
    rsTime.MoveNext
ErrLoop:
    Loop
    
    '循环执行SQL保存数据
    On Error Resume Next

    gcnOracle.BeginTrans
    blnTrans = True

    On Error GoTo ErrHand
    '先更新发生时间
    intMax = UBound(strSQLTime)
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strSQLTime(intPos) <> "" Then
                'Debug.Print strSQLTime(intPos)
    '            objStream.WriteLine (Now & "；SQL：" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "保存护理记录单数据")
            End If
        Next intPos
    End If
    
    intMax = UBound(strSQL)
    If intMax > 0 Then
'        objStream.WriteLine (Now & "准备保存数据")
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
    '            objStream.WriteLine (Now & "；SQL：" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "保存护理记录单数据")
            End If
        Next
    '    objStream.WriteLine (Now & "保存数据完成")
    End If

    gcnOracle.CommitTrans
    SaveData = True
    blnTrans = False
    mblnSaved = True
    mblnChange = False
    
    '更新数据行的记录ID列,表示该数据已保存
    strSaveRows = strSaveRows & ","
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        If InStr(1, strSaveRows, "," & intRow & ",") <> 0 Then
            strDatetime = ""
            If Val(VsfData.TextMatrix(intRow, mlngDemo)) > 0 Then
                If CheckGroupDate(intRow) = True Then
                    '保存后的修改才进入此流程，取该条记录的实际时间
                    If mblnDateAd Then
                        strDate = Format(VsfData.TextMatrix(intRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(intRow, mlngActiveTime), "MM")
                    Else
                        strDate = Mid(VsfData.TextMatrix(intRow, mlngActiveTime), 1, 10)
                    End If
                    strTime = Mid(VsfData.TextMatrix(intRow, mlngActiveTime), 12, 5)
                Else
                    strDate = VsfData.TextMatrix(intRow - Val(VsfData.TextMatrix(intRow, mlngDemo)) + 1, mlngDate)
                    strTime = VsfData.TextMatrix(intRow - Val(VsfData.TextMatrix(intRow, mlngDemo)) + 1, mlngTime)
                End If
            Else
                '普通数据
                strDate = VsfData.TextMatrix(intRow, mlngDate)
                strTime = VsfData.TextMatrix(intRow, mlngTime)
            End If

            If strDate <> "" Then
                If mblnDateAd Then
                    strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                Else
                    strDate = Format(strDate, "yyyy-MM-dd")
                End If
                strDatetime = strDate & " " & strTime & ":00"
                If Val(VsfData.TextMatrix(intRow, mlngDemo)) >= 1 Then
                    strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(VsfData.TextMatrix(intRow, mlngDemo)), "0") & Val(VsfData.TextMatrix(intRow, mlngDemo)) - 1
                End If
            End If
            
            If strDatetime <> "" Then
                gstrSQL = " Select A.ID,A.发生时间,A.保存人 From 病人护理数据 A,病人护理文件 B" & vbNewLine & _
                          " Where A.文件ID=B.ID And B.ID=[1] And A.发生时间=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取记录ID", Val(VsfData.TextMatrix(intRow, c文件ID)), CDate(strDatetime))
                If rsTemp.RecordCount <> 0 Then
                    VsfData.TextMatrix(intRow, mlngRecord) = rsTemp!ID
                    VsfData.TextMatrix(intRow, mlngActiveTime) = Format(rsTemp!发生时间, "YYYY-MM-DD HH:mm:ss")
                    If mlngOperator <> -1 Then VsfData.TextMatrix(intRow, mlngOperator) = NVL(rsTemp!保存人)
                    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                    Call SingerShowType(VsfData, intRow, intRow + Val(VsfData.TextMatrix(intRow, mlngRowCount)) - 1)
                End If
            End If
        End If
    Next
    
    '94211,2016-4-14,陈刘 批量录入
    ReDim Preserve strCollectSQL(1 To 1)
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        If VsfData.RowHidden(intRow) = True Then
            strCollectSQL(ReDimArray(strCollectSQL)) = intRow
        End If
    Next
    
    For intRow = 1 To UBound(strCollectSQL)
        If strCollectSQL(intRow) <> "" Then
            VsfData.RowPosition(Val(strCollectSQL(intRow))) = VsfData.Rows - 1
        End If
    Next
    
    str文件ID = ""
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        
        If VsfData.TextMatrix(intRow, mlngRowCount) Like "*|1" And VsfData.RowHidden(intRow) = False Then
            If str文件ID <> VsfData.TextMatrix(intRow, c文件ID) Then str发生时间_L = ""
            If VsfData.TextMatrix(intRow, mlngActiveTime) <> "" Then str发生时间 = VsfData.TextMatrix(intRow, mlngActiveTime)
        
            If VsfData.TextMatrix(intRow, c文件ID) = str文件ID And str发生时间_L <> "" And Mid(str发生时间_L, 1, 16) = Mid(str发生时间, 1, 16) And str发生时间_L <> str发生时间 Then
                '日期相同，秒数不同，且不是汇总数据行，则说明这些数据是一组，更新lngDemo列
                VsfData.TextMatrix(intRow, mlngDemo) = intRow - lngLastRow + 1
                If intRow - lngLastRow = Val(FormatValue(VsfData.TextMatrix(lngLastRow, mlngRowCount))) Then
                    VsfData.TextMatrix(lngLastRow, mlngDemo) = 1
                End If
            Else
                lngLastRow = intRow
                str发生时间_L = str发生时间
                str文件ID = VsfData.TextMatrix(intRow, c文件ID)
            End If
        End If
    Next
    
    '内存记录集清空
    mrsCellMap.Filter = 0
    If mrsCellMap.RecordCount <> 0 Then mrsCellMap.MoveFirst
    Do While True
        If mrsCellMap.EOF Then Exit Do
        mrsCellMap.Delete
        mrsCellMap.Update
        mrsCellMap.MoveNext
    Loop
    '保存当前数据
    Call InitCons
    mblnShow = False
    Call DataMap_Save
    
    If VsfData.Rows > VsfData.FixedRows Then
        VsfData.ROW = VsfData.FixedRows
    Else
        vsfHistory.Rows = vsfHistory.FixedRows
    End If
    
    Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UpdateTime(rsTimeCur As ADODB.Recordset, ByVal strTime As String, lngID As Long, lng文件ID As Long) As String
    Dim strSQL As String
    rsTimeCur.Filter = "发生时间='" & Format(strTime, "YYYY-MM-DD HH:mm:ss") & "' And 标记=0 And ID<>" & lngID & " and  文件ID=" & lng文件ID
    If rsTimeCur.RecordCount > 0 Then
        lngID = Val(rsTimeCur!ID)
        lng文件ID = Val(rsTimeCur!文件ID)
        strSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!时间, "YYYY-MM-DD HH:mm:ss"), lngID, lng文件ID)
    Else
        strSQL = "Zl_病人护理数据_发生时间(" & lngID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'))"
    End If
    UpdateTime = strSQL
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String
    Dim strLockItem As String                   '同步过来的数据,不允许修改或删除
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       '同步过来的数据占用的最大行数
    Dim intNULL As Integer, lngStartRow As Long, lngRowCount As Long, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    Dim strPart As String
    Dim lngOrder As Long, intGroupFirstRows As Integer
    Dim lngCol1 As Long, lngRow1 As Long, lngCurRow As Long, strText As String, lngCount As Long
    Dim blnTure As Boolean, blnShow As Boolean
    Dim varAssistant() As Variant, strAssistantCols As String
    Dim strCols As String
    '120694 批量录入增加打印功能
    Dim objPrint As New zlTFPrintTends, objAppRow As zlTFTabAppRow
    Dim lngWidth As Long
    Dim bytMode As Byte
    On Error GoTo err_exit
    
    Select Case Control.ID
    '数据分组，包括保存前的分组和保存后的追加分组
    Case conMenu_Edit_Group_Append
        '添加分组，在当前行(只有一行的数据行)的最后分组行增加空白行
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        '0)数据初始化
        '隐蔽已显示的录入控件
        cmdWord.Visible = False
        Select Case mintType
        Case 0, 3
            picInput.Visible = False
            lstSelect(2).Visible = False
        Case 1, 2
            lstSelect(mintType - 1).Visible = False
            If mintType = 1 Then
                txtLst.Visible = False
                PicLst.Visible = False
            End If
        Case 4, 5
            picDouble.Visible = False
            lstSelect(2).Visible = False
        Case 6
            picMutilInput.Visible = False
            lstSelect(2).Visible = False
        Case 7
            picDoubleChoose.Visible = False
        End Select
        blnShow = mblnShow
        mintType = -1: mblnShow = False
        
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
            lngRowCount = 1
        Else
            If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
                lngStartRow = GetStartRow(VsfData.ROW)
                lngRowCount = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            Else
                lngStartRow = VsfData.ROW
                lngRowCount = 1
            End If
        End If
        
        '确定分组起始行(取消此检查，保存时在进行检查)
        lngRow = lngStartRow
'        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) > 1 Then '分组数据行
'            lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
'            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
'                For lngStartRow = lngRow To VsfData.FixedRows Step -1
'                    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
'                        Exit For
'                    End If
'                Next lngStartRow
'                If lngStartRow < VsfData.FixedRows Then Exit Sub
'            End If
'        End If
'
'        If VsfData.TextMatrix(lngStartRow, mlngDate) = "" Or VsfData.TextMatrix(lngStartRow, mlngTime) = "" Then
'            VsfData.ROW = lngStartRow
'            VsfData.COL = mlngDate
'            RaiseEvent AfterRowColChange("进行数据分组时，分组起始行日期或时间不能为空。", True)
'            Exit Sub
'        End If
        
        lngStartRow = lngRow
        '追加数据时，需重新计算选中数据的行数，计算是不能包含大文本段信息。
        '如：一行数据5行，最大非大文本的内容只有3行，选中改行追加数据时，应该追加到第4行，demo=1的为3行，demo=4的为2行
        intNULL = lngStartRow + lngRowCount - 1
        For lngRow = lngRowCount To 1 Step -1
            blnNULL = True
            For lngCol = cHideCols + 1 To VsfData.Cols - 1
                If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
                    If VsfData.TextMatrix(lngRow + lngStartRow - 1, lngCol) <> "" And Not (IsDiagonal(lngCol) And InStr(1, VsfData.TextMatrix(lngRow + lngStartRow - 1, lngCol), "/") <> 0) Then
                        blnNULL = False
                        Exit For
                    End If
                End If
            Next
            
            If Not blnNULL Then Exit For
            intNULL = intNULL - 1
        Next
        '从新填写行序号
        If intNULL < lngStartRow Then intNULL = lngStartRow
        For lngRow = lngStartRow To intNULL
            VsfData.TextMatrix(lngRow, mlngRowCount) = (intNULL - lngStartRow + 1) & "|" & lngRow - lngStartRow + 1
            VsfData.TextMatrix(lngRow, mlngRowCurrent) = (intNULL - lngStartRow + 1)
        Next
        
        If mlngSignName <> -1 Then
            If Trim(VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignName)) <> "" Then
                VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignName)
                If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignTime)
            End If
        End If
        
        If intNULL + 1 <= lngStartRow + lngRowCount - 1 Then
            For lngRow = intNULL + 1 To lngStartRow + lngRowCount - 1
                '清空隐藏列数据
                For lngCol = cHideCols + 1 To VsfData.Cols - 1
                    If VsfData.ColHidden(lngCol) = True Then VsfData.TextMatrix(lngRow, lngCol) = ""
                Next lngCol
                VsfData.TextMatrix(lngRow, mlngRowCount) = (lngStartRow + lngRowCount - intNULL - 1) & "|" & (lngRow - intNULL)
                VsfData.TextMatrix(lngRow, mlngRowCurrent) = (lngStartRow + lngRowCount - intNULL - 1)
            Next
            lngRowCount = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            '更新本列大文本列数据集信息
            Call CellMap_UpdateAssistant(lngStartRow)
            blnTure = False
        Else
            '检查下一行数据是否为空,如果不是空行直接添加到下一行
            lngCurRow = lngStartRow + lngRowCount
            blnTure = False
            If lngCurRow >= VsfData.Rows Then
                blnTure = True
            ElseIf VsfData.TextMatrix(lngStartRow, c文件ID) <> VsfData.TextMatrix(lngCurRow, c文件ID) Then
                blnTure = True
            Else
                If Not VsfData.RowHidden(lngCurRow) Then
                    For lngCol = cHideCols + 1 To VsfData.Cols - 1
                        If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor Then
                            If VsfData.TextMatrix(lngCurRow, lngCol) <> "" Then
                                blnTure = True
                                Exit For
                            End If
                        End If
                    Next
                Else
                    blnTure = True
                End If
            End If
        End If
        
        If blnTure = True Then
            '1)先添加一个空行
            VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(VsfData.Rows - 1, c文件ID) = VsfData.TextMatrix(lngStartRow, c文件ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c病人ID) = VsfData.TextMatrix(lngStartRow, c病人ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c主页ID) = VsfData.TextMatrix(lngStartRow, c主页ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c婴儿) = VsfData.TextMatrix(lngStartRow, c婴儿)
            
            '从当前行记录的空白行开始，每行的位置+所增加的空白行数
            For lngRow = VsfData.Rows - 2 To lngStartRow + lngRowCount Step -1
                VsfData.RowPosition(lngRow) = lngRow + 1
            Next
            lngRow = lngStartRow + lngRowCount - 1
            '2)当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
            Call CellMap_Update(lngRow, 1)
        End If
        '3)更新分组相关控制
        Call AppendGroup(lngStartRow)
        lngRow1 = VsfData.ROW
        lngCol1 = VsfData.COL
        If InStr(1, VsfData.TextMatrix(lngRow1, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow1, mlngRowCount) = "1|1"
        '4)在原有分组数据上分组,需要处理分组序号
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then '分组数据行
            '确定分组起始行
            lngRow = lngStartRow
            lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
                For lngStartRow = lngRow To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                        lngRow = lngStartRow
                        Exit For
                    End If
                Next lngStartRow
                If lngStartRow < VsfData.FixedRows Then GoTo ErrNext
                lngStartRow = lngRow
            End If
            '重新组织分组序号
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow1, mlngRowCount), "|")(0))
            lngCurRow = lngRow1
            For lngRow = lngRow1 + intGroupFirstRows To VsfData.Rows - 1
                If lngRow = lngCurRow + intGroupFirstRows Then
                    If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                        Exit For
                    Else
                        VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - Val(lngStartRow) + 1
                    End If
                    If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                    intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
                    lngCurRow = lngRow
                End If
            Next
            blnTure = False
            mblnEditAssistant = False
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                '寻找大文本列
                mrsSelItems.Filter = "列=" & lngCol - cHideCols
                If mrsSelItems.RecordCount > 0 Then
                    lngOrder = Val(mrsSelItems!项目序号)
                    mrsItems.Filter = "项目序号=" & lngOrder
                    If mrsItems.RecordCount = 0 Then
                        mrsItems.Filter = 0
                        GoTo ErrNext
                    End If
                    mblnEditAssistant = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100) And Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <= 1
                    If Not mblnEditAssistant Then GoTo ErrNext
                        
                    If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
                    intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
                    '为分组行时，选择数据起始行，编辑内容显示所有大文本行
                    strText = ""
                    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                        For lngRow = 0 To intGroupFirstRows - 1
                            strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow + lngStartRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                        Next lngRow
                        lngCount = lngStartRow + intGroupFirstRows - 1
                        For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
                            If VsfData.RowHidden(lngRow) = False Then
                                 '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                                If lngRow > lngCount Then
                                    If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                                    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                                    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
                                End If
                                    
                                strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                            Else
                                lngCount = lngCount + 1
                            End If
                        Next lngRow
                        mintType = -1: mblnShow = False
                        If strText = "" Then GoTo ErrNext
                        VsfData.ROW = lngStartRow
                        VsfData.COL = lngCol
                        mintType = 0
                        'lngRow1 要追加的行
                        Call MoveNextCell(False, True, strText, lngRow1)
                        mintType = -1
                        blnTure = True
                    End If
ErrNext:
                End If
            Next lngCol
            'blnTrue=false 说明记录单没有大行文本(分组数据在起始行点击追加，在追加行录入数据，只能保存追加行和邻近下一行其中的一条数据)
            If blnTure = False Then
                '从起始行开始处理分组数据(防止已经保存的分组数据分组行和保存的时间不对应，导致存在两条相同时间的数据)
                '如：保存的数据起始行Demo=1，发生时间秒数为=01，此时追加一条新记录，Demo=2 保存时秒数也为01(如果修改了起始行数据就不存在这种情况)
                intGroupFirstRows = 0
                lngCurRow = lngStartRow
                For lngRow = lngStartRow To VsfData.Rows - 1
                    If lngRow = lngCurRow + intGroupFirstRows Then
                        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 And intGroupFirstRows > 0 Then Exit For
                        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
                        lngCurRow = lngRow
                        If CheckGroupDate(lngRow) = True Then
                            '保存后的修改才进入此流程，取该条记录的实际时间
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 12, 5)
                        Else
                            '新增时进入此流程
                            strDate = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngTime)
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngRow & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngRow, mlngDemo) & "|0"
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngRow & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngRow, mlngDemo) & "|0"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                Next lngRow
            Else
                '还原选择列
                mblnShow = False
                mintType = -1
                VsfData.ROW = lngRow1
                VsfData.COL = lngCol1
            End If
        End If
        Call SetActiveColColor
    '粘贴,清除时需要同步mrsCellMap数据
    Case conMenu_Edit_Copy
        '复制指定数据行的数据
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)

        '复制记录集
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)

        '得到指定数据行的起始行,结束行
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!页号 = mint页码
            mrsCopyMap!行号 = lngRow
            For lngCol = 0 To lngCols - VsfData.FixedCols    '多了一个固定列
                mrsCopyMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        '粘贴时，将目标行整体覆盖，同步过来的数据列，活动列除外
        '活动项目可能不同页面项目不同，部位不同，所以不考虑活动项目
        '同步行所占用的行数不变，如不够再添加空白行，再行粘贴
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        lngRow = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngRow, mlngDemo) <> "" Then
            RaiseEvent AfterRowColChange("要粘贴的数据行，不能是分组数据行。", True)
            Exit Sub
        End If
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
        '对于已经存在数据，如果汇总数据已经签名不能粘贴
        blnTure = False
        If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 And mstrColCollect <> "" Then
            '找出数据不为空的汇总列
            For lngRow = 0 To UBound(Split(mstrColCollect, "|"))
                strValue = GetRelatiionNo(CStr(Split(mstrColCollect, "|")(lngRow)), 2)
                strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(Split(mstrColCollect, "|")(lngRow), ";")(0)
            Next
            strCols = Mid(strCols, 2)
            If strCols <> "" Then
                If ISCollectSigned(Val(VsfData.TextMatrix(lngStartRow, c文件ID)), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "HH:MM")) Then
                    blnTure = True
                    If MsgBox("您要修改的数据所对应的汇总数据已签名，复制数据中所包含的汇总列数据将不能被粘贴，请问您是否继续。", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
        
        '得到目标数据行的起始行,结束行
        strField = "ID|页号|行号|列号|记录ID|数据|删除"
        lngCols = VsfData.Cols - 1
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
        Else
            '删除多余的数据行,仅留一行
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) - 1
            For intNULL = 1 To lngRows
                VsfData.RemoveItem lngRow + 1
            Next
            '更新数据行
            Call CellMap_Update(lngStartRow, -1 * lngRows)
        End If

        '往下搜索空行,如果有其它数据行则计算需增加的行数
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '保证当前输入的内容在一页中显示全
            If lngRow + lngStartRow > VsfData.Rows - 1 Then Exit For

            If Val(VsfData.TextMatrix(lngRow + lngStartRow, c病人ID)) = 0 And VsfData.TextMatrix(lngRow + lngStartRow, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '先增加空行
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '从当前行记录的空白行开始，每行的位置+所增加的空白行数
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If

        '还原日期，时间，强制不允许修改
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        '记录用户修改过的单元格
        If mlngDate <> -1 Then
            strKey = mint页码 & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\时间
        strKey = mint页码 & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '向表格填充数据
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCol = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCol + VsfData.FixedCols
                    Case 1, c文件ID, c床号, c姓名, c病人ID, c主页ID, c婴儿, c血压频次, _
                         mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord, mlngSignName
                    Case Else
                        If Not (blnTure = True And InStr(1, "," & strCols & ",", "," & lngCol - (cHideCols - 1) & ",") > 0) Then
                            If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 Then
                                VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCol + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCol).Value)
    
                                '修改标志
                                If .AbsolutePosition = .RecordCount Then
                                    strKey = mint页码 & "," & lngStartRow & "," & lngCol + VsfData.FixedCols
                                    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol + VsfData.FixedCols & "|" & _
                                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & GetMutilData(lngStartRow, lngCol + VsfData.FixedCols, lngTop, lngHeight) & "|0"
                                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                                End If
                            End If
                        End If
                    End Select
                Next
                .MoveNext
            Loop
        End With
        '当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
        Call CellMap_Update(lngStartRow, mrsCopyMap.RecordCount - 1)
        '表格上色
        Call SetActiveColColor
        mblnChange = True

    Case conMenu_Edit_Clear
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        
        lngRow = GetStartRow(VsfData.ROW)
        lngStartRow = lngRow
        lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
        
        '准备删除
        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
        
        If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("已签名的数据不允许删除！", True)
            Exit Sub
        End If
        
        blnTure = False
        '已经分组的数据不允许删除起始行
        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) = 1 And lngRow + lngRows < VsfData.Rows Then
            lngCount = lngRow + lngRows - 1
            For lngCurRow = lngRow + lngRows To VsfData.Rows - 1
                 '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                If lngCurRow > lngCount Then
                    If Val(VsfData.TextMatrix(lngCurRow, mlngDemo)) <= 1 Then Exit For
                    If VsfData.RowHidden(lngCurRow) = False Then blnTure = True: Exit For '只要存在一个没有隐藏的分组就退出
                    lngCount = Val(Split(VsfData.TextMatrix(lngCurRow, mlngRowCount), "|")(0)) + lngCurRow - 1
                End If
            Next lngCurRow
        End If
        
        If blnTure = True Then
            RaiseEvent AfterRowColChange("存在分组数据行时，不允许删除分组起始行。", True)
            Exit Sub
        End If
        
        '已有的数据存在汇总已签名的数据不允许删除
        If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 Then
            If CheckCollectIsData(lngStartRow, 1) = True Then
                If ISCollectSigned(Val(VsfData.TextMatrix(lngStartRow, c文件ID)), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("您要删除的数据存在汇总列数据，且本条数据所对应的汇总数据已签名，不允许删除。", True)
                    Exit Sub
                End If
            End If
        End If
        
        cmdWord.Visible = False
        Select Case mintType
        Case 0, 3
            picInput.Visible = False
            lstSelect(2).Visible = False
        Case 1, 2
            lstSelect(mintType - 1).Visible = False
            If mintType = 1 Then
                txtLst.Visible = False
                PicLst.Visible = False
            End If
        Case 4, 5
            picDouble.Visible = False
            lstSelect(2).Visible = False
        Case 6
            picMutilInput.Visible = False
            lstSelect(2).Visible = False
        Case 7
            picDoubleChoose.Visible = False
        End Select
        mintType = -1
        blnNULL = mblnShow
        mblnShow = False
        '删除所有数据行
        lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
        lngRowCount = lngRows
        strAssistantCols = ""
        '获取分组数据的大文本列数据内容
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 1 Then
            Call GetGroupAssistant(strAssistantCols, varAssistant)
        End If
        
        For intNULL = 2 To lngRows
            VsfData.RowHidden(lngRow + intNULL - 1) = True
        Next
        '清除非起始行分组数据，不清除大文本信息并取消该分组
        '如：本分组包含三组，清除第二组时，将第二组大文段内容累加在第3祖上
        '记录用户修改过的单元格
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
            If CheckGroupDate(lngStartRow) = True Then
                '保存后的修改才进入此流程，取该条记录的实际时间
                If mblnDateAd Then
                    strDate = Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "MM")
                Else
                    strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 1, 10)
                End If
                strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 12, 5)
            Else
                strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
                strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
            End If
        Else
            '普通数据
            strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
            strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
        End If
      
        strField = "ID|页号|行号|列号|记录ID|数据|部位|汇总|删除"
        '记录用户修改过的单元格
        strKey = mint页码 & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '2\时间
        strKey = mint页码 & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        strField = "ID|页号|行号|列号|记录ID|数据|汇总|删除"
        '删除启始行中非同步的数据
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            '填写修改标志
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                strKey = mint页码 & "," & lngStartRow & "," & lngCol
                strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Next
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        Else
            '填写修改标志(存在同步数据,日期与时间列不允许清除)``
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCol <> mlngDate And lngCol <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCol) = ""

                    strKey = mint页码 & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            Next
            VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        End If
        
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then
            For intNULL = 2 To lngRows
                For lngCol = 0 To VsfData.Cols - 1
                        VsfData.TextMatrix(lngRow + 1, lngCol) = ""
                Next lngCol
                VsfData.RowPosition(lngRow + 1) = VsfData.Rows - 1
            Next
            Call CellMap_Update(lngStartRow, -1 * (lngRows - 1))
            lngRowCount = 1
            
            '重新组织分组行号和大文本段内容
            If strAssistantCols <> "" Then
                Call ReSetGroupAssistant(True, False, strAssistantCols, varAssistant)
            Else
                Call ReSetGroupDemo(lngStartRow)
            End If
        End If

        mblnShow = False
        If lngStartRow + lngRowCount < VsfData.Rows - 1 Then
            lngRow1 = lngStartRow + lngRowCount
            If Val(VsfData.TextMatrix(lngRow1, mlngRowCount)) > 1 Then
                lngRow1 = GetStartRow(lngRow1)
                If lngRow1 + Val(Split(VsfData.TextMatrix(lngRow1, mlngRowCount), "|")(0)) < VsfData.Rows - 1 Then
                    lngRow1 = lngRow1 + Val(Split(VsfData.TextMatrix(lngRow1, mlngRowCount), "|")(0))
                End If
            End If
            
            If VsfData.RowHidden(lngRow1) = False Then
                VsfData.ROW = lngRow1
            Else
                For lngRow = lngRow1 + 1 To VsfData.Rows - 1
                    If VsfData.RowHidden(lngRow) = False Then VsfData.ROW = lngRow: Exit For
                Next lngRow
            End If
        End If
        mblnShow = blnNULL
        mblnChange = True
    Case conMenu_Edit_SPECIALCHAR

        '检查当前录入控件
        On Error Resume Next
        Dim objTXT As TextBox
        Dim intPos As Integer, intLen As Integer

        mstrSymbol = frmInsSymbol.ShowMe(False, 0)
        If mintSymbol = -1 Then
            Set objTXT = txtInput
        Else
            Set objTXT = txt(mintSymbol)
        End If
        strText = objTXT.Text
        intPos = objTXT.SelStart
        intLen = Len(objTXT)
        objTXT.Text = Mid(strText, 1, intPos) & mstrSymbol & Mid(strText, intPos + 1)
    
        If mintSymbol = -1 Then
            Call txtInput_KeyDown(vbKeyReturn, 0)
        Else
            Call txt_KeyDown(Val(txt(mintSymbol)), vbKeyReturn, 0)
        End If
    
    Case conMenu_Edit_Word
        Call cmdWord_Click
    Case conMenu_Edit_Import
        '导入入量
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub '其它条件已经在update中进行了判断
        Call ImportAmount
    Case conMenu_Edit_NewItem
        '在当前有效数据行（可能当前有效数据行是多行）之后增加一空白行
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        
        '确定分组起始行
        lngRow = lngStartRow
        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) > 1 Then '分组数据行
            lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
                For lngStartRow = lngRow To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngStartRow
                If lngStartRow < VsfData.FixedRows Then Exit Sub
            End If
        End If
        lngRow = lngStartRow
        
        VsfData.Rows = VsfData.Rows + 1
        VsfData.TextMatrix(VsfData.Rows - 1, c文件ID) = VsfData.TextMatrix(lngStartRow, c文件ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c床号) = VsfData.TextMatrix(lngStartRow, c床号)
        VsfData.TextMatrix(VsfData.Rows - 1, c姓名) = VsfData.TextMatrix(lngStartRow, c姓名)
        VsfData.TextMatrix(VsfData.Rows - 1, c病人ID) = VsfData.TextMatrix(lngStartRow, c病人ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c主页ID) = VsfData.TextMatrix(lngStartRow, c主页ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c婴儿) = VsfData.TextMatrix(lngStartRow, c婴儿)
        VsfData.TextMatrix(VsfData.Rows - 1, c血压频次) = VsfData.TextMatrix(lngStartRow, c血压频次)
        
        lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
        intNULL = lngStartRow + lngRows - 1
        '确定当前行的最后分组行或数据的最后一行
        For lngRow = lngStartRow + lngRows To VsfData.Rows - 1
            If lngRow > intNULL Then
                '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                       lngStartRow = lngRow - 1
                       Exit For
                End If
                intNULL = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
            End If
        Next lngRow
        
'        strKey = VsfData.TextMatrix(lngStartRow, mlngRowCount)
'        If InStr(1, strKey, "|") <> 0 And strKey <> "1|1" Then
'            strKey = Split(strKey, "|")(0)
'            strKey = strKey & "|" & strKey
'            For lngRow = VsfData.ROW + 1 To VsfData.Rows - 1
'                If VsfData.TextMatrix(lngRow, mlngRowCount) = strKey Then
'                    lngStartRow = lngRow + 1
'                    Exit For
'                End If
'            Next
'        Else
'            lngStartRow = VsfData.ROW + 1
'        End If
        
        For lngRow = VsfData.Rows - 2 To lngStartRow + 1 Step -1  '从倒数第二行开始
            VsfData.RowPosition(lngRow) = lngRow + 1
        Next
        VsfData.ROW = lngStartRow + 1
        Call CellMap_Update(VsfData.ROW, 1)
        Call SetActiveColColor
        mblnChange = True
    Case conMenu_Edit_Save
        Call SaveME
    Case conMenu_Edit_Transf_Cancle
        Call CancelMe
    Case conMenu_File_BatPrint
        If zlEvent_Print Is Nothing Then
            Set zlEvent_Print = New zlTFPrintMethod
        End If
        
        Call zlEvent_Print.InitPrint(gcnOracle, gstrDBUser)
        bytMode = zlEvent_Print.zlPrintAsk(VsfData.TextMatrix(VsfData.ROW, c病人ID), VsfData.TextMatrix(VsfData.ROW, c主页ID), VsfData.TextMatrix(VsfData.ROW, c婴儿), VsfData.TextMatrix(VsfData.ROW, c文件ID))

        If bytMode <> 0 Then zlEvent_Print.zlPrintOrViewTends (True), bytMode
    Case conMenu_Tool_Sign
        Call SignMe
    Case conMenu_Tool_SignEarse
        Call UnSignMe
    Case conMenu_Help_Help '帮助
        RaiseEvent UsrHelp
    Case conMenu_File_Exit '退出
        RaiseEvent UsrExit
    End Select

err_exit:
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrData
    Dim blnFind As Boolean, blnAllow As Boolean
    Dim strItem As String
    Dim intDo  As Integer, intCount As Integer
    
    blnAllow = (cmd刷新.Tag <> "")
    Select Case Control.ID
    Case conMenu_Edit_Group_Append
        Control.Enabled = False
        If Not mblnInit Then Exit Sub
        Control.Enabled = Not mblnSigned And blnAllow And VsfData.ROW >= VsfData.FixedRows And (ISGroupAppend = True)
    Case conMenu_Edit_Copy
        Control.Enabled = Not mblnShow And blnAllow
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If Not mblnInit Then Exit Sub
        If mblnSigned Then Exit Sub
        If mrsCopyMap.State = 0 Then Exit Sub
        Control.Enabled = Not mblnShow And mrsCopyMap.RecordCount And blnAllow
    Case conMenu_Edit_Clear
        Control.Enabled = Not mblnSigned And blnAllow
    Case conMenu_Edit_SPECIALCHAR
        Control.Enabled = mblnShow And (mintType = 0 Or mintType = 6) And blnAllow
    Case conMenu_Edit_Word
        '60291:刘鹏飞,2013-04-17,只要是文本项目都允许进行词句选择
        Control.Enabled = (mblnEditAssistant Or mblnEditText) And Not mblnSigned And blnAllow
    Case conMenu_Edit_NewItem
        Control.Enabled = Not mblnSigned And blnAllow
    Case conMenu_Edit_Save
        Control.Enabled = mblnChange And Not mblnSigned And blnAllow
    Case conMenu_Edit_Transf_Cancle
        Control.Enabled = mblnChange And blnAllow
    Case conMenu_File_BatPrint
        Control.Enabled = mblnInit
    Case conMenu_Tool_Sign
        Control.Enabled = mblnSaved And Not mblnSigned And Not mblnChange And blnAllow
    Case conMenu_Tool_SignEarse
        Control.Enabled = mblnSaved And mblnSigned And Not mblnChange And blnAllow
    Case conMenu_Edit_Import '入量导入
        Control.Enabled = False
        If Not mblnInit Then Exit Sub
        Control.Enabled = Not mblnSigned And blnAllow And VsfData.ROW >= VsfData.FixedRows And mblnShow And mstrColImCorrelative <> ""
    End Select
End Sub

Private Function ISActiveUsed(ByVal strTest As String) As Boolean
    Dim arrData, arrCol
    Dim lngCol As Long
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '检查某个活动项目是否已被其它列绑定
    ISActiveUsed = True

    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        lngCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            If strTest = arrCol(intIn) And VsfData.COL - (cHideCols + VsfData.FixedCols - 1) <> lngCol Then
                RaiseEvent AfterRowColChange(Split(strTest, ",")(1) & mrsItems!项目名称 & " 已经被绑定到" & lngCol & "列，不允许重复绑定！", True)
                Exit Function
            End If
        Next
    Next
    ISActiveUsed = False
End Function

Private Function GetActivePart(ByVal intFindCol As Integer, ByVal intItem As Integer) As String
    '获取指定列的活动项目
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strPart As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '将活动项目加入到查询SQL中，格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...
    '绑定多个项目，该列就自动转为对角线列

    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCol = intFindCol - cHideCols Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            strPart = Split(arrCol(intItem), ",")(1)
            Exit For
        End If
    Next
    GetActivePart = strPart
End Function



Private Sub cmdWord_Click()
    Dim strInput As String
    '弹出词句选择器

    If Val(cmdWord.Tag) = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, Val(VsfData.TextMatrix(VsfData.ROW, c病人ID)), Val(VsfData.TextMatrix(VsfData.ROW, c主页ID)), Val(VsfData.TextMatrix(VsfData.ROW, c婴儿)), strInput)
    
    If Val(cmdWord.Tag) = -1 Then
        txtInput.Text = strInput
        Call txtInput_KeyDown(vbKeyReturn, 0)
    Else
        txt(Val(cmdWord.Tag)).Text = strInput
        Call txt_KeyDown(Val(cmdWord.Tag), vbKeyReturn, 0)
    End If
End Sub

Private Sub cmd刷新_Click()
    '读取文件格式
    mblnInit = False
    mblnHistory = False
    picNull.Visible = False
    mlng格式ID = cbo护理文件格式.ItemData(cbo护理文件格式.ListIndex)
    mlng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    
    Call InitVariable
    Call InitCons
    If Not ReadStruDef Then Exit Sub
    Call zlRefresh
    cmd刷新.Tag = 1
    mblnInit = True
    
    '保存当前数据
    Call DataMap_Save
End Sub



Private Sub lstSelect_LostFocus(Index As Integer)
    If Index = 2 Then lstSelect(2).Visible = False
End Sub

Private Sub optLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optLevel(0).Value Then
         SaveSetting "ZLSOFT", "私有模块\usrTendFileMutilEditor\" & gstrUserName, "Value", 0
    Else
        SaveSetting "ZLSOFT", "私有模块\usrTendFileMutilEditor\" & gstrUserName, "Value", 1
    End If
End Sub

Private Sub PicLst_GotFocus()
    If PicLst.Visible = False Then Exit Sub
    If Trim(txtLst.Text) = "" Then
        PicLst.Tag = 0
        lstSelect(0).SetFocus
    Else
        PicLst.Tag = 1
        txtLst.SetFocus
    End If
End Sub

Private Sub txtLst_GotFocus()
    mblnEditAssistant = False
    mblnEditText = False
    PicLst.Tag = 1
    Call zlControl.TxtSelAll(txtLst)
    lstSelect(0).ListIndex = -1
End Sub

Private Sub txtLst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    If KeyCode = vbKeyReturn Or _
        (KeyCode = vbKeyRight And txtLst.SelStart = Len(txtLst.Text)) Or _
        (KeyCode = vbKeyLeft And txtLst.SelStart = 0) Then
        '移动到下一个单元格
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
    If Shift = vbShiftMask And KeyCode = vbKeyDown Then KeyCode = 0: lstSelect(0).SetFocus
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplit.Tag = 1
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplit.Tag) = 0 Then Exit Sub
    
    If picSplit.Top + Y < 4000 Then
        picSplit.Top = 4000
    ElseIf ScaleHeight - (picSplit.Top + Y) < 3000 Then
        picSplit.Top = ScaleHeight - 3000
    Else
        picSplit.Move picSplit.Left, picSplit.Top + Y
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplit.Tag) = 1 Then Call cbsThis_Resize

    picSplit.Tag = 0
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 58 And VsfData.COL = mlngTime Then KeyAscii = 0
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim lngRow As Long, lngCol As Long
    Dim dblHeight As Double, dblWidth As Double
    Dim strItemInfo As String
    
    If Not mblnInit Then Exit Sub
    Call InitCons
    
    '当列不可见时不显示说明信息
    If VsfData.RowIsVisible(VsfData.ROW) = True And VsfData.ColIsVisible(VsfData.COL) = True Then
         '显示当前项目的相关信息
        mrsSelItems.Filter = "列=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
        If mrsSelItems.RecordCount <> 0 Then
            mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
            If mrsItems.RecordCount <> 0 Then
                strItemInfo = Trim(NVL(mrsItems!说明, ""))
            End If
        End If
        mrsSelItems.Filter = 0
        mrsItems.Filter = 0
    End If
     '--48659:刘鹏飞,2012-09-14,添加字段'说明'
    RaiseEvent ShowTipInfo(VsfData, strItemInfo, True)
    
'    '计算固定行的高度
'    For lngRow = 0 To 2
'        If Not VsfData.RowHidden(lngRow) Then dblHeight = dblHeight + VsfData.ROWHEIGHT(lngRow)
'    Next
'    '从可见行开始向下查找最后一个可见行
'    For lngRow = NewTopRow To VsfData.Rows - 1
'        If Not VsfData.RowIsVisible(lngRow) Then
'            lngRow = lngRow - 1
'            Exit For
'        End If
'    Next
'    '从可见列开始查找最后一个可见列
'    For lngCol = NewLeftCol To VsfData.Cols - 1
'        If Not VsfData.ColIsVisible(lngCol) Then
'            lngCol = lngCol - 1
'            Exit For
'        Else
'            dblWidth = dblWidth + VsfData.ColWidth(lngCol)
'        End If
'    Next
'
'    If Not VsfData.RowIsVisible(VsfData.Row) Then
'        VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'    Else
'        '当前数据行的高度+固定行的高度如果大于表格控件的高度,说明当前选择的数据行存在遮住部分的情况
'        If VsfData.Row >= lngRow - 1 And CellRect.Bottom * (lngRow - NewTopRow + 1) + dblHeight >= VsfData.ClientHeight Then
'            '遮住部分的情况下
'            VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'        End If
'    End If
'
'    If Not VsfData.ColIsVisible(VsfData.Col) Then
'        VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'    Else
'        '当前数据行的高度+固定行的高度如果大于表格控件的高度,说明当前选择的数据行存在遮住部分的情况
'        If VsfData.Col = lngCol And dblWidth >= VsfData.ClientWidth Then
'            '遮住部分的情况下
'            VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'        End If
'    End If
'
'    Call VsfData_EnterCell
End Sub

Private Sub VsfData_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim blnResult As Boolean
    
    If Not mblnInit Then Exit Sub
    If mintType = -1 Then Exit Sub
    blnResult = MoveNextCell(True, True)
    Cancel = Not blnResult
End Sub

Private Sub VsfData_DblClick()
    Call vsfdata_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim strCols As String
    Dim intMax As Integer
    Dim lngStart As Long
    Dim blnCheck As Boolean
    
    On Error Resume Next
    
    If Not mblnInit Then Exit Sub
    If VsfData.ROW < VsfData.FixedRows Then Exit Sub
    '隐蔽已显示的录入控件
    cmdWord.Visible = False
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
        lstSelect(2).Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
        If mintType = 1 Then
            txtLst.Visible = False
            PicLst.Visible = False
        End If
    Case 4, 5
        picDouble.Visible = False
        lstSelect(2).Visible = False
    Case 6
        picMutilInput.Visible = False
        lstSelect(2).Visible = False
    Case 7
        picDoubleChoose.Visible = False
    End Select

    '未定义的列不允许录入数据
    mintType = -1
    If InStr(1, mstrPrivs, "护理记录登记") = 0 Then Exit Sub
    If mblnSigned Then Exit Sub
    If Not mblnShow Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    
    '如果是活动项目则不允许编辑
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then Exit Sub
    If VsfData.COL <= c血压频次 Then Exit Sub
    
    If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
        RaiseEvent AfterRowColChange("已签名的数据不允许再次编辑，请取消签名后再试！", True)
        Exit Sub
    End If
    
    If VsfData.TextMatrix(VsfData.ROW, mlngDemo) <> "" Then
        '只有新增的未保存的数据，才允许修改日期与时间
        If (VsfData.COL = mlngDate Or VsfData.COL = mlngTime) Then
            If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1 Then
                Exit Sub
            Else
                'If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then Exit Sub
            End If
        End If
    End If
    
    '汇总行涉及到的明细,如果汇总行已签名则其汇总列不允许修改
    If mstrColCollect <> "" Then
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) > 0 Then
            '不允许修改汇总列数据，也不允许修改日期与时间
            If InStr(1, "|" & mstrColCollect, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then
                blnCheck = True
            ElseIf InStr(1, "|" & mstrColCorrelative, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 And mstrColCorrelative <> "" Then
                blnCheck = True
            ElseIf VsfData.COL = mlngTime Or VsfData.COL = mlngDate Then
                blnCheck = True
            End If
            If blnCheck = True Then
                If ISCollectSigned(Val(VsfData.TextMatrix(lngStart, c文件ID)), Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10), Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("该条数据所对应的汇总行数据已签名，不允许修改当前汇总列或日期时间列！", True)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    '让控件获得焦点
    Select Case mintType
    Case 0, 3
        picInput.SetFocus
    Case 1, 2
        If mintType = 2 Then
            lstSelect(mintType - 1).SetFocus
        Else
            PicLst.SetFocus
        End If
    Case 4, 5
        picDouble.SetFocus
    Case 6
        picMutilInput.SetFocus
    End Select
End Sub

Private Sub vsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    Dim strStart As String, strEnd As String
    Dim strItemInfo As String
    
    If mblnInit = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    On Error GoTo ErrHand

    '选择列,同步数据列直接退出,避免此处清除提示信息
    '显示当前项目的相关信息
    mrsSelItems.Filter = "列=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!项目值域) <> "" Then
                If mrsItems!项目类型 = 0 Then
                    strInfo = "有效范围:" & Split(mrsItems!项目值域, ";")(0) & "～" & Split(mrsItems!项目值域, ";")(1)
                Else
                    strInfo = "有效范围:" & mrsItems!项目值域
                End If
            Else
                strInfo = ""
            End If
            strItemInfo = Trim(NVL(mrsItems!说明, ""))
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '--48659:刘鹏飞,2012-09-14,添加字段'说明'
    RaiseEvent ShowTipInfo(VsfData, strItemInfo, True)
    
    '提取该病人历史数据
    If OldRow <> NewRow Then
        Call RefreshHistoryData(NewRow)
    End If
    
    RaiseEvent AfterRowColChange(strInfo, False)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshHistoryData(ByVal lngRow As Long)
'刷新历史数据
    Dim strStart As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCurrDate As String
    
    On Error GoTo ErrHand
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    
    If lngRow > VsfData.Rows Then
        lngRow = VsfData.Rows - 1
        VsfData.ROW = lngRow
    End If
    If Val(VsfData.TextMatrix(lngRow, c文件ID)) > 0 Then
        mstrMaxDate = Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm")
        strStart = Format(DateAdd("d", -1 * (mintPreDays + 1), CDate(strCurrDate)), "yyyy-MM-dd") & " 23:59:59"  '倒算出昨天
        
        '装入数据
        Call SQLCombination
        gstrSQL = Replace(mstrSQL, " And 1=2", " And l.发生时间 between [2] and [3]")
        Call SQLDIY(gstrSQL)
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", Val(VsfData.TextMatrix(lngRow, c文件ID)), CDate(strStart), CDate(mstrMaxDate))
        '绑定数据并设置护理记录单的格式,同时实现一行数据分行显示的功能
        Call PreTendFormatHistory(rsTemp)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub vsfdata_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String
    On Error GoTo ErrHand
    
    If Not mblnInit Then Exit Sub
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitVariable()
    '清除常量
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignName = -1
    mlngSignTime = -1
    mlngRecord = -1
    mlngNoEditor = -1
    mlngActiveTime = -1
    mintType = -1
    
    mblnShow = False
    mblnSigned = False
    mblnSaved = False
    mblnChange = False
    mblnEditAssistant = False
    mblnEditText = False
    mblnEditHistoryAssistant = False
End Sub

Private Sub InitCons()
    '隐藏输入控件
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    lstSelect(2).Visible = False
    picDouble.Visible = False
    picDoubleChoose.Visible = False
    picMutilInput.Visible = False
    cmdWord.Visible = False
    txtLst.Visible = False
    PicLst.Visible = False
    mintType = -1
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar

    On Error GoTo ErrHand

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False

    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)

        '------------------------------------------------------------------------------------------------------------------
        '工具栏定义
        Set cbrToolBar = cbsThis.Add("标准工具", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        cbrToolBar.ShowTextBelowIcons = False
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Group_Append, "追加"): cbrControl.IconId = 3045: cbrControl.ToolTipText = "追加分组"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制"): cbrControl.ToolTipText = "复制(Ctrl+C)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "粘贴"):  cbrControl.ToolTipText = "粘贴(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"):   cbrControl.ToolTipText = "清除"

            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "特殊符号"):  cbrControl.ToolTipText = "插入特殊符号(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "词句选择"):  cbrControl.ToolTipText = "词句选择(Ctrl+W)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "入量导入"):  cbrControl.ToolTipText = "入量导入(Ctrl+I)"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "空行"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "增加空行"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "签名"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消"): cbrControl.IconId = 229
            
            Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "记录单输出"):  cbrControl.ToolTipText = "记录单打印预览"
            cbrControl.IconId = conMenu_File_Print: cbrControl.BeginGroup = True
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        End With
        
        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next

        '------------------------------------------------------------------------------------------------------------------
        '工具栏定义
        pic过滤条件.Height = IIf(mblnBlowup = True, 375, 300)
        Set cbrToolBar = cbsThis.Add("过滤条件", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.Flags = xtpFlagAlignLeft
            cbrCustom.Handle = pic过滤条件.hWnd
            cbrCustom.ToolTipText = "条件"
        End With

         '快键绑定
        With cbsThis.KeyBindings
            .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
            .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
            .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
            .Add FCONTROL, Asc("W"), conMenu_Edit_Word
            .Add FCONTROL, Asc("I"), conMenu_Edit_Import
            .Add 0, VK_F1, conMenu_Help_Help
        End With

    InitMenuBar = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim lng文件ID As Long, lng病人ID As Long, lng主页ID As Long, int婴儿 As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    On Error GoTo ErrHand
    '数据发生时间必须在当前科室的有效时间范围内
    lng文件ID = Val(VsfData.TextMatrix(lngRow, c文件ID))
    lng病人ID = Val(VsfData.TextMatrix(lngRow, c病人ID))
    lng主页ID = Val(VsfData.TextMatrix(lngRow, c主页ID))
    int婴儿 = Val(VsfData.TextMatrix(lngRow, c婴儿))
    
    blnMsg = (strMsg <> "")
    
    gstrSQL = "Select 开始时间,结束时间 From 病人护理文件 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取文件ID", lng文件ID)
    If rsTemp.RecordCount > 0 Then
        '检查文件开始,结束时间
        If Format(strTime, "yyyy-MM-dd HH:mm") < Format(NVL(rsTemp!开始时间), "yyyy-MM-dd HH:mm") Then
            strMsg = "发生时间不能小于文件开始时间[" & NVL(rsTemp!开始时间) & "]"
            GoTo exitHand
        End If
        If NVL(rsTemp!结束时间) <> "" Then
            If Format(strTime, "yyyy-MM-dd HH:mm") <= Format(NVL(rsTemp!结束时间), "yyyy-MM-dd HH:mm") Then
                strMsg = "发生时间不能大于文件结束时间[" & NVL(rsTemp!结束时间) & "]"
                GoTo exitHand
            End If
        End If
    End If
    
    '75760:刘鹏飞,处理婴儿存在出院医嘱的情况
    If int婴儿 <> 0 Then
        strBabyOutTime = GetAdviceOutTime(lng病人ID, lng主页ID, int婴儿)
        If strBabyOutTime <> "" Then
            If Format(strTime, "YYYY-MM-DD HH:mm") > Format(strBabyOutTime, "YYYY-MM-DD HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & Format(strBabyOutTime, "YYYY-MM-DD HH:mm") & "]"
                GoTo exitHand
            End If
            '补录小时检查
            If Format(DateAdd("H", glngHours, strBabyOutTime), "yyyy-MM-dd HH:mm") < Format(strCurTime, "yyyy-MM-dd HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]"
                GoTo exitHand
            End If
            CheckTime = True
            Exit Function
        End If
    End If
    
    '根据病人变动记录进行检查
    gstrSQL = " Select   开始原因,病区ID,to_char(开始时间,'yyyy-MM-dd hh24:mi') AS 开始时间,to_char(NVL(终止时间,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS 终止时间 " & _
              " From 病人变动记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]" & _
              " Order by 开始时间,开始原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前科室有效时间范围", lng病人ID, lng主页ID)
    With rsTemp
        .Filter = "病区ID=" & mlng病区ID
        Do While Not .EOF
            If Format(strTime, "yyyy-MM-dd HH:mm") >= Format(!开始时间, "yyyy-MM-dd HH:mm") And Format(strTime, "yyyy-MM-dd HH:mm") <= Format(!终止时间, "yyyy-MM-dd HH:mm") Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '找到了就退出
        If blnExist Then
            If Not IsAllowInput(lng病人ID, lng主页ID, strTime, strCurTime) Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]"
                GoTo exitHand
            End If

            CheckTime = True
            Exit Function
        End If

        '没找到,就整理原因进行准确性提示
        .Filter = "开始原因=1"
        If .RecordCount <> 0 Then
            If !开始原因 = 1 And Format(strTime, "yyyy-MM-dd HH:mm") < Format(!开始时间, "yyyy-MM-dd HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入院时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=2"
        If .RecordCount <> 0 Then
            If !开始原因 = 2 And Format(strTime, "yyyy-MM-dd HH:mm") < Format(!开始时间, "yyyy-MM-dd HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入科时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=10"
        If .RecordCount <> 0 Then
            If !开始原因 = 10 And Format(strTime, "yyyy-MM-dd HH:mm") > Format(!终止时间, "yyyy-MM-dd HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & !终止时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '其他情况说明
        strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[不在当前病区的有效时间范围内]"
        GoTo exitHand
    End With

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strOrders As String, strText As String
    '检查录入数据的合法性(中文也认为是一个字符,考虑到体温项目等存在不升\外出等信息)
    '返回的数据,如果一列绑定多个项目,以单引号做为分隔符

    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定N个项目,手工录入
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '免检
        If mintType = 1 Then
            If Val(PicLst.Tag) = 0 Then
                txtLst.Text = ""
                If InStr(1, lstSelect(mintType - 1).Text, "-") <> 0 Then
                    strText = Mid(lstSelect(mintType - 1).Text, InStr(1, lstSelect(mintType - 1).Text, "-") + 1)
                Else
                    strText = ""
                End If
            Else
                strText = Trim(txtLst.Text)
            End If
        Else
            j = lstSelect(mintType - 1).ListCount
            For i = 1 To j
                If lstSelect(mintType - 1).Selected(i - 1) Then
                    strText = strText & "," & Mid(lstSelect(mintType - 1).List(i - 1), InStr(1, lstSelect(mintType - 1).List(i - 1), "-") + 1)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOrders = lstSelect(mintType - 1).Tag
    Case 4
        strText = txtUpInput.Text & "'" & txtDnInput.Text
        strOrders = txtUpInput.Tag & "'" & txtDnInput.Tag
    Case 6
        j = txt.Count
        For i = 1 To j
            strText = strText & "'" & txt(i - 1).Text
            strOrders = strOrders & "'" & txt(i - 1).Tag
        Next
        If strText <> "" Then
            strText = Mid(strText, 2)
            strOrders = Mid(strOrders, 2)
        End If
    Case 3      '免检
        strText = lblInput.Caption
    Case 5      '免检
        strText = lblUpInput.Caption & "/" & lblDnInput.Caption
    Case 7
        strText = cboChoose(0).Text & "/" & cboChoose(1).Text
    End Select
    If Val(strOrders) <> 0 Then
        If Not CheckValid(strText, strOrders, strInfo) Then Exit Function
    ElseIf VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then
        If Not CheckDateTime(strText, strInfo) Then Exit Function
    End If

    strReturn = strText
    CheckInput = True
End Function

Private Function CheckDateTime(strText As String, strInfo As String) As Boolean
    Dim arrData
    Dim blnCheck As Boolean
    Dim strCurrDate As String
    Dim strDate As String, strMonth As String, strDay As String
    Dim rsCheck As New ADODB.Recordset
    Dim arrTime As Variant
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If VsfData.COL = mlngDate Then
        If mblnDateAd Then
            If Trim(strText) = "" Then
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "日期格式错误，如1月12日：12/01"
                Exit Function
            End If

            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
            '检查是否翻年后编辑之前的时间(一个月的限制)
            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                strDate = DateAdd("yyyy", -1, CDate(strDate))
            End If
            If Not IsDate(strDate) Then
                strInfo = "录入的数据不是合法的日期，如1月12日：12/01"
                Exit Function
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "录入的数据不是合法的日期，如1月12日：2011-01-12"
                Exit Function
            End If
            strDate = Format(strText, "yyyy-MM-dd")
        End If
        
        If Format(strDate, "YYYY-MM-DD") > Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "YYYY-MM-DD") Then
            strInfo = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
            Exit Function
        End If

        If VsfData.TextMatrix(VsfData.ROW, mlngTime) <> "" Then
            blnCheck = True
            strDate = strDate & " " & VsfData.TextMatrix(VsfData.ROW, mlngTime)
        End If
    Else
        If Trim(strText) = "" Then
            strInfo = "时间不能为空！"
            Exit Function
        End If
        If InStr(1, Trim(strText), ":") = 0 Then
            Select Case Len(strText)
            Case 3, 4
                strText = String(4 - Len(strText), "0") & strText
                strText = Mid(strText, 1, 2) & ":" & Mid(strText, 3)
            Case Is < 3
                strText = String(2 - Len(strText), "0") & strText
                strText = Format(Now, "HH") & ":" & strText
            End Select
        End If
        arrTime = Split(Trim(strText), ":")
        
        If UBound(arrTime) <> 1 Then
            strInfo = "录入的时点格式非法！[小时:分钟]"
            Exit Function
        Else
            If Len(Trim(arrTime(0))) < 2 Then arrTime(0) = String(2 - Len(Trim(arrTime(0))), "0") & Trim(arrTime(0))
            If Len(Trim(arrTime(1))) < 2 Then arrTime(1) = String(2 - Len(Trim(arrTime(1))), "0") & Trim(arrTime(1))
            strText = arrTime(0) & ":" & arrTime(1)
        End If
        
        '合法性检查
        If IsNumeric(arrTime(0)) = False Or IsNumeric(arrTime(1)) = False Or Len(Trim(arrTime(0))) > 2 Or Len(Trim(arrTime(1))) > 2 Then
            strInfo = "录入的时点格式非法！[小时:分钟]"
            Exit Function
        End If
        If Mid(strText, 3, 1) <> ":" Then
            strInfo = "录入的时点格式非法！[小时:分钟]"
            Exit Function
        End If
        If Val(arrTime(0)) < 0 Or Val(arrTime(0)) > 23 Then
            strInfo = "录入的时点格式非法！[小时应在0至23之间]"
            Exit Function
        End If
        If Val(arrTime(1)) < 0 Or Val(arrTime(1)) > 59 Then
            strInfo = "录入的时点格式非法！[分钟应在0至59之间]"
            Exit Function
        End If

        '进行合法性检查
        If VsfData.TextMatrix(VsfData.ROW, mlngDate) <> "" Then
            strDate = VsfData.TextMatrix(VsfData.ROW, mlngDate)
            If mblnDateAd Then
                strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            '检查是否翻年后编辑之前的时间(一个月的限制)
            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                strDate = DateAdd("yyyy", -1, CDate(strDate))
            End If
            strDate = Format(strDate & " " & strText, "YYYY-MM-DD HH:mm:ss")
            
            '70990:刘鹏飞,2014-03-13,超期补录天数控制修改
            If Format(strDate, "YYYY-MM-DD HH:mm") > Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "YYYY-MM-DD HH:mm") Then
                strInfo = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
                Exit Function
            End If
            
            blnCheck = True
        End If
    End If

    If blnCheck Then
        '不管是新录入还是修改的数据 如果存在历史数据都不允许修改
        gstrSQL = " Select 1 From 病人护理数据 Where 文件ID=[1] And 发生时间=[2] And ([3]=0 OR ID<>[3])"
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "检查发生时间", Val(VsfData.TextMatrix(VsfData.ROW, c文件ID)), CDate(strDate), Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)))
        If rsCheck.RecordCount > 0 Then
            strInfo = "您录入的时点已经存在历史数据！"
            Exit Function
        End If
        
        '从数据库中没有找到，开始从用户录入的数据寻找
        If Not CheckChangeDataTime(VsfData.ROW, strDate, strInfo) Then Exit Function
        
        '修改时间的对应的汇总列如果存在数据，则检查是否已经存在相应的小结并进行了签名
        '规则:新增的数据强制检查;已有的数据则只需要检查时间变化的数据(因可能存在A操作员在开始无汇总行或有但未签名进行了时间调整，B操作员签名了，A操作员在保存的情况)
        '说明：新增的数据是可以修改任何列；已有的数据如果汇总行已经签名是不允许修改日期时间列和汇总列的（未处理对修改的时间在同一个汇总范围的判断）
        blnCheck = True
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then
            If Format(VsfData.TextMatrix(VsfData.ROW, mlngActiveTime), "YYYY-MM-DD HH:mm") = Format(strDate, "YYYY-MM-DD HH:mm") Then
                blnCheck = False
            End If
        End If
        If blnCheck = True Then
            If CheckCollectIsData(VsfData.ROW) = True Then
                If ISCollectSigned(Val(VsfData.TextMatrix(VsfData.ROW, c文件ID)), Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                    strInfo = "您录入的时点所对应的汇总行数据已签名，不允许再添加新的汇总列数据！"
                    Exit Function
                End If
            End If
        End If
        
        '70990:刘鹏飞,2014-03-13
        '数据发生时间不能在当前操作员所属科室的有效时间以前
        If Not CheckTime(VsfData.ROW, strDate, strCurrDate, strInfo) Then
            Exit Function
        End If
    End If

    CheckDateTime = True
End Function

Private Function CheckChangeDataTime(ByVal lngRow As Long, ByVal strCurDate As String, ByRef strMsg As String) As Boolean
'检查新录入的时间，是否与现有的时间相同，如果相同则提示不能录入
    Dim strDateHistory As String, strTimeHistory As String, strDatetime As String '用户已经录入的日期和时间
    Dim lngCurRow As Long, intPage As Integer, blnDel As Boolean, blnTrue As Boolean
    Dim strCurrDate As String, lngRecord As Long, strActiveTime As String
    Dim strRows As String, strPages As String, strTimes As String, lngCol As Long
    Dim lng文件ID As Long
    Dim arrRows
    On Error GoTo ErrHand
    
    lng文件ID = Val(VsfData.TextMatrix(lngRow, c文件ID))
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With mrsCellMap
        .Filter = "列号=" & mlngDate & " OR 列号=" & mlngTime
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Not (lngCurRow = !行号 And intPage = !页号) Then
                blnDel = False
endWork:
                If Not (lng文件ID = Val(VsfData.TextMatrix(lngCurRow, c文件ID))) Then GoTo ErrNext
                If lngCurRow = lngRow And intPage = mint页码 Then GoTo ErrNext
                If lngCurRow > 0 Then
                    blnDel = VsfData.RowHidden(lngCurRow)
                    strActiveTime = Format(VsfData.TextMatrix(lngCurRow, mlngActiveTime), "YYYY-MM-DD HH:mm:ss")
                End If
                
                If blnTrue = True And strDatetime <> "" Then
                    If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strCurDate, "YYYY-MM-DD HH:mm:ss") Then
                        '存在相同时间的数据没有删除，直接进行提示
                        If blnDel = False Then
                            strMsg = "第" & lngCurRow & "行已经存在相同时点的数据，请检查！"
                            Exit Function
                        Else
                            If lngRecord > 0 Then '保存的数据删除，如果时间和原有时间相同直接提示，不相同恢复时间为原有时间
                                If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strActiveTime, "YYYY-MM-DD HH:mm:ss") Then
                                    strMsg = "您录入的时点已经存在历史数据！"
                                    Exit Function
                                Else '恢复时间为原有时间
                                    VsfData.TextMatrix(lngCurRow, mlngDate) = Format(strActiveTime, "YYYY-MM-DD")
                                    VsfData.TextMatrix(lngCurRow, mlngTime) = Mid(strActiveTime, 12, 5)
                                    '记录行号和页号
                                    strRows = strRows & "," & lngCurRow
                                    strPages = strPages & "," & intPage
                                    strTimes = strTimes & "," & strActiveTime
                                End If
                            Else '未保存的数据删除，直接清空记录集内容信息
                                For lngCol = mlngDate To VsfData.Cols - 1
                                    VsfData.TextMatrix(lngCurRow, lngCol) = ""
                                Next lngCol
                                '记录行号和页号
                                strRows = strRows & "," & lngCurRow
                                strPages = strPages & "," & intPage
                                strTimes = strTimes & "," & "[LPF]"
                            End If
                        End If
                    End If
ErrNext:
                    blnTrue = False
                    If .EOF Then Exit Do
                End If
                '赋初值
                intPage = !页号
                lngCurRow = !行号
                strDateHistory = ""
                strTimeHistory = ""
                strDatetime = ""
                lngRecord = NVL(!记录ID, 0)
                blnTrue = False
            End If
            
            If !列号 = mlngDate Then
                If NVL(!汇总, 0) <> 1 Then
                    strDateHistory = NVL(!数据)
                    If strDateHistory <> "" Then
                        If mblnDateAd Then
                            strDateHistory = Mid(strCurrDate, 1, 5) & ToStandDate(strDateHistory)
                            '检查是否翻年后编辑之前的时间(一个月的限制)
                            If CDate(strDateHistory) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDateHistory, 6, 2) = "12" Then
                                strDateHistory = DateAdd("yyyy", -1, CDate(strDateHistory))
                            End If
                        Else
                            strDateHistory = Format(strDateHistory, "yyyy-MM-dd")
                        End If
                    End If
                End If
            Else '时间列
                strTimeHistory = NVL(!数据, "00:00")
                If strDateHistory = "" Then strDateHistory = Mid(strCurrDate, 1, 10)
                strDatetime = strDateHistory & " " & strTimeHistory & ":00"

                '处理分组数据，保存时与普通数据无区别，只是秒数+
                If Val(NVL(!部位)) >= 1 Then
                    strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!部位), "0") & Val(!部位) - 1
                End If
                strDatetime = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
                blnTrue = True
            End If
        .MoveNext
        Loop
        
        If blnTrue Then GoTo endWork
        mrsDataMap.Filter = 0
    End With
    
    '更新mrsCellMap记录集
    If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
    If Left(strPages, 1) = "," Then strPages = Mid(strPages, 2)
    If Left(strTimes, 1) = "," Then strTimes = Mid(strTimes, 2)
    arrRows = Split(strRows, ",")
    For lngCurRow = 0 To UBound(arrRows)
        mrsCellMap.Filter = "页号=" & Val(Split(strPages, ",")(lngCurRow)) & " And 行号=" & Val(arrRows(lngCurRow))
        If CStr(Split(strTimes, ",")(lngCurRow)) = "[LPF]" Then
            Do While Not mrsCellMap.EOF
                mrsCellMap.Delete
                mrsCellMap.Update
                mrsCellMap.MoveNext
            Loop
        Else
            Do While Not mrsCellMap.EOF
                If mrsCellMap!列号 = mlngDate Then
                    mrsCellMap!数据 = Format(CStr(Split(strTimes, ",")(lngCurRow)), "YYYY-MM-DD")
                    mrsCellMap.Update
                ElseIf mrsCellMap!列号 = mlngTime Then
                    mrsCellMap!数据 = Mid(CStr(Split(strTimes, ",")(lngCurRow)), 12, 5)
                    mrsCellMap.Update
                End If
            mrsCellMap.MoveNext
            Loop
        End If
        
    Next lngCurRow
    
    mrsCellMap.Filter = 0
    strMsg = ""
    CheckChangeDataTime = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckValid(strReturn As String, ByVal strOrders As String, strInfo As String) As Boolean
    Dim arrData, arrOrder
    Dim i As Integer, j As Integer, blnNumber As Boolean
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String, strFormat As String, strFormat1 As String

    '按列格式组装数据
    mrsSelItems.Filter = "列=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '有此列但未进行定义
        strFormat = NVL(mrsSelItems!格式)   '{P[体温]C}{...}
        strFormat1 = strFormat
    End If
    mrsSelItems.Filter = 0

    '检查数据
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!项目名称
            blnNumber = False
            If strText <> "" Then
                If mrsItems!项目类型 = 0 And InStr(1, "0,4", mrsItems!项目表示) <> 0 Then
                    blnNumber = IIf(IsCanves(Val(strOrders)), False, True)
                    If blnNumber Then strText = Val(strText)
                    If NVL(mrsItems!项目小数, 0) <> 0 Then   '等于零是通过控件的MaxLength来控制的
                        If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                        If Len(strText) > mrsItems!项目长度 Then
                            mrsItems.Filter = 0
                            strInfo = "[" & strName & "]录入的数据超过了合法精度！"
                            Exit Function
                        End If

                        strText = Val(arrData(i))
                        If InStr(1, strText, ".") <> 0 Then
                            strText = Mid(strText, InStr(1, strText, ".") + 1)
                            If Len(strText) > mrsItems!项目小数 Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]录入的小数部分超过了合法精度！"
                                Exit Function
                            End If
                        End If
                        strText = IIf(IsNumeric(arrData(i)), Val(arrData(i)), arrData(i))
                    End If
                    If mrsItems!项目表示 = 0 Then
                        If Not IsNull(mrsItems!项目值域) Then
                            dblMin = Val(Split(mrsItems!项目值域, ";")(0))
                            dblMax = Val(Split(mrsItems!项目值域, ";")(1))
                            If blnNumber Then
                                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]录入的数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                                    Exit Function
                                End If
                            Else
                                If Not IsNumeric(strText) Then
                                    mrsUsual.Filter = "名称 = '" & strText & "'"
                                    If Not mrsUsual.RecordCount > 0 Then
                                        strInfo = "[" & strName & "]录入的未记说明不是常用体温说明！"
                                        Exit Function
                                    End If
                                Else
                                    If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                                        mrsItems.Filter = 0
                                        strInfo = "[" & strName & "]录入的数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!项目长度 Then
                        strInfo = "[" & strName & "]录入的数据超过了最大长度：" & mrsItems!项目长度 & "！"
                        mrsItems.Filter = 0
                        Exit Function
                    End If
                End If
                If IsNumeric(strText) And blnNumber = True Then
                    If Val(strText) < 1 And Val(strText) > 0 Then strText = "0" & Val(strText)
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                '删除该项目
                Call SubstrPro(strFormat, strName)
            End If
        Else
            strFormat = strReturn
        End If
    Next
    If j = -1 Then
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = mrsItems!项目名称
            strFormat = Replace(strFormat, "[" & strName & "]", strText)
        End If
    End If
    mrsItems.Filter = 0

    strFormat = Replace(strFormat, "{", "")
    strFormat = Replace(strFormat, "}", "")
    If strFormat = SubstrFormat(strFormat1, arrOrder) Then strFormat = ""
    If InStr(strFormat, SubstrFormat(strFormat1, arrOrder)) = Len(strFormat) And InStr(1, strFormat1, "心率") > 0 And InStr(1, strFormat1, "脉搏") > 0 Then strFormat = Replace(strFormat, SubstrFormat(strFormat1, arrOrder), "")
    
    strReturn = strFormat
    CheckValid = True
End Function

Public Function SubstrFormat(ByVal strData As String, ByVal arrOrder As Variant) As String
    '获取绑定项目的前后缀符号
    Dim i As Integer
    Dim strOrders As String, strName As String
    For i = 0 To UBound(arrOrder)
        strOrders = CStr(arrOrder(i))
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!项目名称
        End If
        strData = Replace(strData, "[" & strName & "]", "")
    Next i
    strData = Replace(strData, "{", "")
    strData = Replace(strData, "}", "")
    
    SubstrFormat = strData
End Function

Public Function SubstrVal(ByVal strData As String, ByVal strFormat As String, ByVal strName As String, intPos As Integer) As String
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    Dim strQZ As String, strHZ As String
    '返回前一个项目的后缀符号+当前项目的前缀符号的位置

    If strData = "" Then Exit Function
    strData = strData
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '得到前缀
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '找到该项目格式串中的结束符号
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '得到后缀
    strHZ = Mid(strFormat, i, r - i)
    '如果后缀为空,继续向后寻找下一个项目的前缀符号
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    '取出指定项目完整的数据串
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '因为是连续取数,考虑到分隔符可能相同的情况,记录上一次的最后位置,下次从这个位置往后取数据
        If j = 0 Then
            '有可能中间存在回车换行符
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 And Not (InStr(strFormat, "心率") > 0 And InStr(strFormat, "脉搏") > 0) Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    '前缀为空,继续向前寻找上一个项目的后缀符号
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If

    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '如果是数字型则去掉回车换行符返回,如果是字符型则原样返回
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '丢弃该项目后的数据
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '丢弃该项目后的数据
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '对比串
    Dim intLen As Integer, intActLen As Integer           '前缀/后缀的长度
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '遇到回车换行符忽略,空格重新比对

    strText = strData
    If strHZ <> "" Then
        '把后缀去掉
        strHZ = Replace(strHZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strHZ)
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strHZ Then
                        blnFind = True
                        intPos = intPos - intActLen
                    Else
                        strCompare = ""
                        intPos = intPos - intActLen + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        '肯定有
        strText = Mid(strText, 1, intPos)
    End If

    '再去掉前缀
    If strQZ <> "" Then
        If InStr(1, strText, strQZ) = 0 Then strText = strQZ & strText
        strQZ = Replace(strQZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strQZ)
        strCompare = ""
        intActLen = 0
        blnFind = False
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strQZ Then
                        blnFind = True
                        intPos = intPos + 1
                    Else
                        strCompare = ""
                        intPos = intPos + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        strText = Mid(strText, intPos)
    End If

    If IsNumeric(Replace(strText, vbCrLf, "")) Then
        SubstrAnaly = Replace(strText, vbCrLf, "")
    Else
        SubstrAnaly = strText
    End If
End Function

'Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
'    Dim i As Integer, j As Integer, l As Integer, r As Integer
'    'intType=0-删除指定格式串;1-得到指定格式串
'    j = Len(strFormat)
'    i = InStr(1, strFormat, "[" & strName & "]")
'    If i = 0 Then Exit Sub
'
'    For l = i To 1 Step -1
'        If Mid(strFormat, l, 1) = "{" Then Exit For
'    Next
'    For r = i To j
'        If Mid(strFormat, r, 1) = "}" Then Exit For
'    Next
'    If intType = 0 Then
'        strFormat = Mid(strFormat, 1, l - 1) & Mid(strFormat, r + 1)
'    Else
'        strFormat = Mid(strFormat, l, r - l + 1)
'    End If
'End Sub

Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
    Dim i As Integer, j As Integer, l As Integer, r As Integer, strHZ As String, strQZ As String
    'intType=0-删除指定格式串;1-得到指定格式串
    j = Len(strFormat)
    i = InStr(1, strFormat, "[" & strName & "]")
    If i = 0 Then Exit Sub
    
    For l = i To 1 Step -1
        If Mid(strFormat, l, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, l + 1, i - l - 1)
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    strHZ = Mid(strFormat, i + Len(strName) + 2, r - i - Len(strName) - 2)
    If intType = 0 Then
        'strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        If Mid(strFormat, 1, l - 1) = "" And Mid(strFormat, r + 1) = "" Then
            strFormat = ""
        Else
            strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        End If
    Else
        strFormat = Mid(strFormat, l, r - l + 1)
    End If
End Sub

Private Function MoveNextCell(Optional ByVal blnNext As Boolean = True, Optional ByVal blnNoMove As Boolean = False, Optional ByVal strText As String = "", Optional ByVal lngDemoRow As Long = 0) As Boolean
 '----------------------------------------------
    '修改人：LPF 2012-04-20
    '修改内容：允许非分组起始行，也可以录入多行数据
    '----------------------------------------------
    Dim arrData
    Dim blnNULL As Boolean                      '是否为空行
    Dim blnGroup As Boolean                     '分组行
    Dim strDate As String, strTime As String    '分组首记录的日期与时间
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngStartGroup As Long, lngMutilRows As Long, lngDeff As Long, intGroupFirstRows As Integer, intBound As Integer, intRowCount As Integer
    Dim intRow As Integer, intRowGroup As Integer, intCount As Integer, intNULL As Integer  '其后有多少空行
    Dim blnTrue As Boolean, blnDate As Boolean, strRows As String
    Dim lngDemo As Long, lngLastNull As Long, lngLastNoNull As Long
    '赋值然后移动到下一个有效单元格
    Dim strKey As String, strField As String, strValue As String, strAppend As String
    Dim blnCallback As Boolean, blnReseGroupAssistant As Boolean, blnGroupAddNum As Boolean '分组数据增加行
    '大文本列和内容信息
    Dim varAssistant() As Variant, strAssistantCols As String
    Dim blnNewRow As Boolean
    
    On Error GoTo ErrHand
    If VsfData.ROW >= VsfData.Rows Then Exit Function
    blnReseGroupAssistant = False
    blnNewRow = Val(GetSetting("ZLSOFT", "私有模块\usrTendFileMutilEditor\" & gstrUserName, "Value")) = 0
    '检查数据,不合格就再次弹出要求录入
    If mintType >= 0 Then
        If strText = "" Then
            strReturn = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
            If Not CheckInput(strReturn, strMsg) Then
                RaiseEvent AfterRowColChange(strMsg, True)
                Exit Function
            End If
            strText = strReturn
        Else
            strReturn = strText
            mstrData = strText
        End If
        '标记当前行为分组行
        blnDate = (InStr(1, "," & mlngDate & "," & mlngTime & ",", "," & VsfData.COL & ",") > 0)
        lngDemo = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo))
        blnGroup = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) >= 1
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
         '如果修改的是非大文本列或时间列的分组数据，检查修改内容行数是否发生变化，如果变化就当分组数据处理，否则以普通数据处理
        If blnGroup = True And Not (mblnEditAssistant = True Or blnDate = True) Then
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            lngStart = GetStartRow(VsfData.ROW)
            '如果编辑的是分组数据最后一个分组行，则当普通数据处理
            If lngStart + intGroupFirstRows < VsfData.Rows Then
                If Val(VsfData.TextMatrix(lngStart + intGroupFirstRows, mlngDemo)) <= 1 Then
                    blnGroup = False: GoTo ErrBegin
                End If
            ElseIf lngStart + intGroupFirstRows >= VsfData.Rows Then
                blnGroup = False
                GoTo ErrBegin
            End If
            
            With txtLength
                '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
                .Width = VsfData.CellWidth
                .Text = Replace(Replace(Replace(IIf(strReturn = "", "a", strReturn), Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = VsfData.CellFontName
                .FontSize = VsfData.CellFontSize
                .FontBold = VsfData.CellFontBold
                .FontItalic = VsfData.CellFontItalic
            End With
            arrData = GetData(txtLength.Text)
            
            blnGroup = False
            intBound = -1
            If (UBound(arrData) + 1) > intGroupFirstRows Then
                blnGroup = True
            ElseIf (UBound(arrData) + 1) < intGroupFirstRows Then
                '得到本条数据占用最大行的列(不包含大文本项目)
                blnNULL = True
                For intRow = lngStart + intGroupFirstRows - 1 To lngStart Step -1
                    For intCount = 0 To mlngNoEditor - 1
                        If VsfData.ColHidden(intCount) = False And ISEditAssistant(intCount) = False Then
                            If VsfData.TextMatrix(intRow, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(intRow, intCount), "/") <> 0) Then
                                blnNULL = False
                                If intCount = VsfData.COL Then
                                    intBound = intCount
                                Else
                                    intBound = intCount
                                    Exit For
                                End If
                            End If
                        End If
                    Next intCount
                    If blnNULL = False Then Exit For
                Next intRow
                
                If blnNULL = False Then
                    intNULL = intRow - lngStart + 1
                    If intBound = VsfData.COL Then
                        blnGroup = True
                    Else
                        blnGroup = (intNULL < intGroupFirstRows)
                    End If
                Else
                    blnGroup = True
                End If
            End If
        End If
ErrBegin:
        blnTrue = False
        lngMutilRows = 1
        intGroupFirstRows = 1
        If Not blnGroup Then
            If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                lngMutilRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            End If
            lngStart = GetStartRow(VsfData.ROW)
        Else
            lngMutilRows = 1
            If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = 1 And (mblnEditAssistant Or blnDate) Then
                '记录分组起始行的数据行数
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intBound = VsfData.ROW + intGroupFirstRows - 1
                For intCount = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                    '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                    If intCount > intBound Then
                        If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '不分组或遇新分组就退出
                        intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
                    End If
                    lngMutilRows = lngMutilRows + 1
                Next
                lngMutilRows = lngMutilRows + intGroupFirstRows - 1 '保证数据行数的准确性
            Else
                '记录分组起始行的数据行数
                If VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1 >= VsfData.FixedRows Then
                    intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngRowCount), "|")(0))
                End If
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                lngMutilRows = intGroupFirstRows
            End If
            lngStart = VsfData.ROW
        End If
        
        '准备赋值
        With txtLength
            '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
            .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
            .Text = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
            .FontName = VsfData.CellFontName
            .FontSize = VsfData.CellFontSize
            .FontBold = VsfData.CellFontBold
            .FontItalic = VsfData.CellFontItalic
        End With
        arrData = GetData(txtLength.Text)
        intCount = UBound(arrData)
        If intCount = -1 Then
            arrData = Array()
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = ""
            intCount = 1
        End If
        lngLastNull = VsfData.ROW + intGroupFirstRows - 1: lngLastNoNull = VsfData.ROW + intGroupFirstRows - 1
        '分组数据中可能存在隐藏的行(点击清除功能时),只有是选择大文本才处理
        If blnGroup = True And mblnEditAssistant = True And blnDate = False Then
            If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 Then
                intNULL = intGroupFirstRows
                lngDeff = VsfData.ROW + intGroupFirstRows - 1
                For intRow = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                    If intRow > lngDeff Then
                        If Val(VsfData.TextMatrix(intRow, mlngDemo)) <= 1 Or intNULL > intCount Then Exit For     '不分组或遇新分组就退出
                        lngDeff = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0)) + intRow - 1
                    End If
                    If VsfData.RowHidden(intRow) = True Then  '删除的分组列
                        '重新组织文本数组，保证分组数据正常复制
                        ReDim Preserve arrData(UBound(arrData) + 1)
                        For intBound = UBound(arrData) To intRow - VsfData.ROW + 1 Step -1
                            arrData(intBound) = arrData(intBound - 1)
                        Next intBound
                        arrData(intRow - VsfData.ROW) = ""
                        '记录最后一次隐藏的行
                        lngLastNull = intRow
                    Else
                        intNULL = intNULL + 1
                        '记录最后一次没有隐藏的行
                        lngLastNoNull = intRow
                    End If
                Next
            End If
        End If
        intCount = UBound(arrData)
        
        lngDeff = 0
        blnGroupAddNum = False
        blnTrue = blnGroup = True And mblnEditAssistant
        If intCount > lngMutilRows - 1 Then
            '对于新增分组数据时，必须要先录入完分组数据才能录入大文本段数据
'            If mblnEditAssistant = True And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 And lngMutilRows = intGroupFirstRows Then
'                strMsg = "新增分组数据时，请先完成数据的分组，最后在录入大文段项目内容！"
'                RaiseEvent AfterRowColChange(strMsg, True)
'                strMsg = ""
'                Exit Function
'            End If
            '往下搜索空行,如果有其它数据行则计算需增加的行数
            '20110830分组号算做同一数据行，将多行文本分解到各行，多余的文本放在统一放在最后一行上;在非首行按回车,只对现有数据进行修改,不对行发生变化
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '保证当前输入的内容在一页中显示全
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, c病人ID)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                    If VsfData.RowHidden(intRow + lngStart) = True Then VsfData.RowHidden(intRow + lngStart) = False
                Else
                    Exit For
                End If
            Next
            '先增加空行
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '从当前行记录的空白行开始，每行的位置+所增加的空白行数
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
                For intRow = lngMutilRows To intCount
                    VsfData.TextMatrix(intRow + lngStart, c文件ID) = VsfData.TextMatrix(lngStart, c文件ID)
                    VsfData.TextMatrix(intRow + lngStart, c病人ID) = VsfData.TextMatrix(lngStart, c病人ID)
                    VsfData.TextMatrix(intRow + lngStart, c主页ID) = VsfData.TextMatrix(lngStart, c主页ID)
                    VsfData.TextMatrix(intRow + lngStart, c婴儿) = VsfData.TextMatrix(lngStart, c婴儿)
                Next intRow

            End If
            
            '当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
            If lngDeff <> 0 Then
                If Not blnGroup Then
                    Call CellMap_Update(lngStart, lngDeff)
                Else
                    Call CellMap_Update(lngStart + lngMutilRows - 1, lngDeff)  '分组行数据从最大一条明细行之后开始处理
                End If
            End If
        
            '对分组数据最后一个分组行为隐藏行的处理
            '例如：该分组具有2个分组，第一组为一行大文本内容为A，第二组为一行大文本内容为C(改行隐藏).此时添加大文本内容为A、B占两行，此时组织得到本组的数据为A、C、B
            '计算方式为:第一行占用内容+隐藏行内容+多出的内容。此处就会把隐藏行放在最后，最后得到的内容为A、B、C，在下面循环赋值中，第一组就为占用2行内容为A、B
            '说明：如果中间存在隐藏行，最后一组没有隐藏，多出的数据就会追加在最后一组数据的后面
            If (lngLastNull - lngLastNoNull) > 0 Then
                For intRow = lngLastNoNull + 1 To lngLastNull
                    strValue = arrData(lngLastNoNull + 1 - VsfData.ROW)
                    For intBound = lngLastNoNull + 1 - VsfData.ROW To UBound(arrData) - 1
                        arrData(intBound) = arrData(intBound + 1)
                    Next intBound
                    arrData(UBound(arrData)) = strValue
                    VsfData.RowPosition(lngLastNoNull + 1) = lngLastNull + (intCount - (lngMutilRows - 1))
                Next intRow
                '更新记录
                For intRow = lngLastNull To lngLastNoNull + 1 Step -1
                    Call CellMap_Update(intRow, intCount - (lngMutilRows - 1), False)
                Next intRow
            End If
        
            '循环赋值
            intCount = UBound(arrData)
            intBound = 0
            blnReseGroupAssistant = (blnGroup = True And Not (mblnEditAssistant Or blnDate))
            blnGroupAddNum = blnReseGroupAssistant
            If blnGroup = True And blnDate = False Then strReturn = ""
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
                '修改非分组数据或非大文本或日期的分组数据
                '非分组数据：直接处理计算行数并更新数据
                '非大文本或日期的分组数据：1、直接处理计算行数并更新数据，2、需要重新处理大文段的内容显示位置
                If (Not blnGroup) Then
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intCount + 1
                    If intRow > 0 And intRow < intCount Then
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = ""
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = ""
                    ElseIf intRow = intCount Then
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = VsfData.TextMatrix(lngStart, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = VsfData.TextMatrix(lngStart, mlngSignTime)
                    End If
                Else
                    '修改分组大文段或日期，需从分组起始行到分组结束行从新整理文本内容显示或日期
                    '分组行的特殊处理,更新内部记录集的代码较多
                    '##########################################
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                        If (mblnEditAssistant = True Or blnDate) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        End If
                        intRowCount = 1
                        '获取该分组数据行的行数
                        For intBound = intRow + 1 To intCount
                             If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) >= 1 Then Exit For
                             intRowCount = intRowCount + 1
                        Next intBound
                        intBound = intRow
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intRowCount & "|1"
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intRowCount
                        If Not blnDate Then strReturn = ""
                    Else
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intRowCount & "|" & intRow - intBound + 1
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = ""
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = ""
                    End If
                    If Not blnDate Then
                        strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStart + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
                    End If
                    '到该分组数据行数最后一行才执行更新操作
                    If intRow = intBound + intRowCount - 1 Then
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = VsfData.TextMatrix(lngStart + intBound, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = VsfData.TextMatrix(lngStart + intBound, mlngSignTime)
                        '保存数据
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '保存后的修改才进入此流程，取该条记录的实际时间
                                If mblnDateAd Then
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                            Else
                                '新增时进入此流程
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                            End If
                        Else
                            '普通数据
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngStart + intBound & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mint页码 & "," & lngStart + intBound & "," & VsfData.COL
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                    '##########################################
                End If
            Next
            '所有隐蔽列进行赋值
            intBound = lngStart + intCount
            For intRow = lngStart + 1 To intBound
                For intCount = 0 To VsfData.Cols - 1
                    VsfData.Cell(flexcpForeColor, intRow, intCount) = VsfData.Cell(flexcpForeColor, lngStart, intCount)
                    If VsfData.ColHidden(intCount) And InStr(1, "," & mlngRowCount & "," & mlngRowCurrent & ",", "," & intCount & ",") = 0 Then
                        If blnGroup And InStr(1, "," & mlngDemo & "," & mlngRecord & "," & mlngActiveTime & ",", "," & intCount & ",") = 0 Then
                            VsfData.TextMatrix(intRow, intCount) = VsfData.TextMatrix(lngStart, intCount)
                        End If
                    End If
                Next
            Next
            lngMutilRows = lngStart + lngMutilRows - 1
        Else
            blnReseGroupAssistant = False
            If blnGroup = True And blnDate = False Then strReturn = ""
            '对该列重新赋值（当只输入一个数字时，不知为何会产生字符ASCII码为1的符号）
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
                If blnGroup = True Then
                    '分组行的特殊处理,更新内部记录集的代码较多
                    '##########################################
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                        If (mblnEditAssistant = True Or blnDate = True) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        End If
                        intRowCount = 1
                        '获取该分组数据行的行数
                        For intBound = intRow + 1 To intCount
                             If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) >= 1 Then Exit For
                             intRowCount = intRowCount + 1
                        Next intBound
                        intBound = intRow
                        If Not blnDate Then strReturn = ""
                    End If
                    If Not blnDate Then
                        strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStart + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
                    End If
                    '到该分组数据行数最后一行才执行更新操作
                    If intRow = intBound + intRowCount - 1 Then
                        '保存数据
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '保存后的修改才进入此流程，取该条记录的实际时间
                                If mblnDateAd Then
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                            Else
                                '新增时进入此流程
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                            End If
                        Else
                            '普通数据
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngStart + intBound & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mint页码 & "," & lngStart + intBound & "," & VsfData.COL
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                    '##########################################
                End If
            Next
            strRows = ""
            lngStartGroup = -1
            For intRow = intCount + 1 To lngMutilRows - 1
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
            Next intRow
            blnReseGroupAssistant = False
            If intCount < (lngMutilRows - 1) Then
                blnReseGroupAssistant = (blnGroup And Not (mblnEditAssistant Or blnDate))
            End If
            
            For intRow = intCount + 1 To lngMutilRows - 1
                '分组行的特殊处理,更新内部记录集的代码较多
                '##########################################
                '保存数据
                If (blnGroup And (mblnEditAssistant Or blnDate)) Then
                    '获取改行起始行
                    If lngStartGroup <> GetStartRow(lngStart + intRow) Then
                        intNULL = GetStartRow(lngStart + intRow)
                        '寻找的起始列mlngDemo肯定>0
                        If Val(VsfData.TextMatrix(intNULL, mlngDemo)) <= 0 Then
                            For intRowGroup = lngStart + intRow To lngStart Step -1
                                If Val(VsfData.TextMatrix(intRowGroup, mlngDemo)) > 0 Then
                                    intNULL = intRowGroup
                                    Exit For
                                End If
                            Next intRowGroup
                            If intNULL = lngStartGroup Then GoTo ErrDemo
                        End If
                        lngStartGroup = intNULL
                        '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
                        If VsfData.TextMatrix(lngStartGroup, mlngRowCount) = "" Then VsfData.TextMatrix(lngStartGroup, mlngRowCount) = "1|1"
                        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartGroup, mlngRowCount), "|")(0))
                        intNULL = lngStartGroup + intGroupFirstRows - 1
                        For intRowGroup = intNULL To lngStartGroup Step -1
                            blnNULL = True
                            For intBound = 0 To VsfData.Cols - 1
                                If Not VsfData.ColHidden(intBound) And intBound < mlngNoEditor Then
                                    If VsfData.TextMatrix(intRowGroup, intBound) <> "" And Not (IsDiagonal(intBound) And InStr(1, VsfData.TextMatrix(intRowGroup, intBound), "/") <> 0) Then
                                        blnNULL = False
                                        Exit For
                                    End If
                                End If
                            Next
                            If Not blnNULL Then Exit For
                            intNULL = intNULL - 1
                            If intRowGroup = lngStartGroup Then
                                 intNULL = intNULL + 1
                            Else
                                If InStr(1, strRows & ",", "," & intRowGroup & ",") = 0 Then strRows = strRows & "," & intRowGroup
                            End If
                        Next intRowGroup
                        
                        '重新填写数据行数
                        For intRowGroup = lngStartGroup To intNULL
                            VsfData.TextMatrix(intRowGroup, mlngRowCount) = intNULL - lngStartGroup + 1 & "|" & intRowGroup - lngStartGroup + 1
                            VsfData.TextMatrix(intRowGroup, mlngRowCurrent) = intNULL - lngStartGroup + 1
                        Next intRowGroup
                        If mlngSignName <> -1 Then
                            If Trim(VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignName)) <> "" Then
                                VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignName)
                                If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignTime)
                            End If
                        End If
                        For intRowGroup = intNULL + 1 To lngStartGroup + intGroupFirstRows - 1
                            VsfData.TextMatrix(intRowGroup, mlngRowCount) = ""
                            VsfData.TextMatrix(intRowGroup, mlngRowCurrent) = ""
                            VsfData.TextMatrix(intRowGroup, mlngRecord) = ""
                            If mlngSignName <> -1 Then VsfData.TextMatrix(intRowGroup, mlngSignName) = ""
                            If mlngOperator <> -1 Then VsfData.TextMatrix(intRowGroup, mlngOperator) = ""
                            If mlngSignTime <> -1 Then VsfData.TextMatrix(intRowGroup, mlngSignTime) = ""
                        Next
                    End If
ErrDemo:
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) > 0 And intRow > intCount Then
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        If CheckGroupDate(lngStart + intRow) = True Then
                            '保存后的修改才进入此流程，取该条记录的实际时间
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 12, 5)
                        Else
                            '新增时进入此流程
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    
                        '分组起始行的行数减少时，重新设置分组号
                        If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
                        If strRows <> "" Then
                            intNULL = 0
                            For intBound = 0 To UBound(Split(strRows, ","))
                                If Val(Split(strRows, ",")(intBound)) < (lngStart + intRow) Then
                                    intNULL = intNULL + 1
                                End If
                            Next intBound
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) - intNULL
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngStart + intRow & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngStart + intRow & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mint页码 & "," & lngStart + intRow & "," & VsfData.COL
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & "" & "|" & strPart & "|1"
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                End If
                '##########################################
            Next
            '修改分组非大文本或日期行时，需要获取分组数据大文本段内容信息，重新组织文本显示
            '如有3组数据，第二2行有3行，修改为1行，第3组数据应该紧接着显示在第2组下面(第二组此时只有1行)
            If blnReseGroupAssistant = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
            lngMutilRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
            '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
            intNULL = lngStart + lngMutilRows - 1
            For intRow = lngMutilRows To 1 Step -1
                blnNULL = True
                For intCount = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(intCount) And intCount < mlngNoEditor And IIf(blnReseGroupAssistant = True, ISEditAssistant(intCount) = False, True) Then
                        If VsfData.TextMatrix(intRow + lngStart - 1, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(intRow + lngStart - 1, intCount), "/") <> 0) Then
                            blnNULL = False
                            Exit For
                        End If
                    End If
                Next
                
                If Not blnNULL Then Exit For
                intNULL = intNULL - 1
            Next
            '从新填写行序号
            If Not blnGroup Then
                If intNULL < lngStart Then intNULL = lngStart
                For intRow = lngStart To intNULL
                    VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                    VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
                Next
                If mlngSignName <> -1 Then
                    If Trim(VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignName)) <> "" Then
                        VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignTime)
                    End If
                End If
                strRows = ""
            Else '分组行以保存的数据删除时，不清空行号
                For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                    If intRow = lngStart Then intNULL = intNULL + 1
                Next intRow
                
                For intRow = lngStart To intNULL
                    VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                    VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
                Next
                If mlngSignName <> -1 Then
                    If Trim(VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignName)) <> "" Then
                        VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignTime)
                    End If
                End If
            End If
            If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
            If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.TextMatrix(intRow, mlngRowCount) = ""
                VsfData.TextMatrix(intRow, mlngRowCurrent) = ""
                VsfData.TextMatrix(intRow, mlngRecord) = ""
                If mlngSignName <> -1 Then VsfData.TextMatrix(intRow, mlngSignName) = ""
                If mlngOperator <> -1 Then VsfData.TextMatrix(intRow, mlngOperator) = ""
                If mlngSignTime <> -1 Then VsfData.TextMatrix(intRow, mlngSignTime) = ""
                If blnReseGroupAssistant = True Then
                    If InStr(1, "," & strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
                ElseIf Not blnGroup Then
                    If InStr(1, "," & strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
                End If
            Next
            '更新记录集大文段信息
            If blnReseGroupAssistant = True Then Call CellMap_UpdateAssistant(lngStart)
        End If
        
        '获取分组起始行所有行信息
        If blnTrue = True Then 'blnTrue为真说明选择的是分组行的起始行，并且是大文本段
            strReturn = ""
            intCount = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
            For intRow = 0 To intCount - 1
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(VsfData.ROW + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next intRow
        End If
        
        If mstrData <> strReturn Or blnTrue = True Then
            If strText <> mstrData Then mblnChange = True
            '同步保存日期与时间列的数据
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 0 Then
                If CheckGroupDate(lngStart) = True Then
                    '保存后的修改才进入此流程，取该条记录的实际时间
                    If mblnDateAd Then
                        strDate = Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "MM")
                    Else
                        strDate = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10)
                    End If
                    strTime = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 12, 5)
                Else
                    '新增时进入此流程
                    strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                    strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                End If
            Else
                '普通数据
                strDate = VsfData.TextMatrix(lngStart, mlngDate)
                strTime = VsfData.TextMatrix(lngStart, mlngTime)
            End If
            
            '1\日期
            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
            If mlngDate <> -1 Then
                strKey = mint页码 & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\时间
            strKey = mint页码 & "," & lngStart & "," & mlngTime
            strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|" & _
                VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
         
            If Not blnGroup Or blnTrue Then
                '记录用户修改过的单元格
                If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    strPart = GetActivePart(VsfData.COL, 0)
                Else
                    strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                End If
                
                strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                strKey = mint页码 & "," & lngStart & "," & VsfData.COL
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & VsfData.COL & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            Call SetActiveColColor
        End If
    End If
    
    '数据行数减少时，将空白行移至到最后一行
    If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
    If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
    strRows = Replace("," & strRows & ",", ",,", "") '不能删除要追加的行
    strRows = Replace("," & strRows & ",", "," & lngDemoRow & ",", "") '不能删除要追加的行
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
            Call CellMap_Update(intRow, -1)
            VsfData.TextMatrix(intRow, mlngDemo) = ""
        End If
    Next intRow
    
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
            '清空改行所有信息
            For intBound = 0 To VsfData.Cols - 1
                VsfData.TextMatrix(intRow, intBound) = ""
            Next intBound
            VsfData.RowHidden(intRow) = True
            VsfData.RowPosition(intRow) = VsfData.Rows - 1
        End If
    Next intRow
   ' Call OutputRsData(mrsCellMap)

    '重新组织分组数据内容
    If blnReseGroupAssistant = True Then
        If blnGroupAddNum = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
        If strAssistantCols <> "" Then
            Call ReSetGroupAssistant(blnNoMove, blnNext, strAssistantCols, varAssistant)
        Else
            Call ReSetGroupDemo(lngStart)
        End If
        
    End If

    MoveNextCell = True
    
    If blnNoMove Then Exit Function
    If blnNext Then
        If blnNewRow = False Then
        '问题号：56592,李涛,批量录入的纵向跳转
              '跳到下一行
toMoveNextRow2:
            If VsfData.ROW < VsfData.Rows - 1 Then
                If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                    intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                    intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
                Else
                    intRow = 1
                End If
                If VsfData.ROW + intRow < VsfData.Rows Then
                    VsfData.ROW = VsfData.ROW + intRow
                Else
                    GoTo toMoveNextCol2
                End If
                If VsfData.RowHidden(VsfData.ROW) Then
                    If VsfData.ROW < VsfData.Rows - 1 Then
                        GoTo toMoveNextRow2
                    Else
                        For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
                            If VsfData.RowHidden(intRow) = False Then
                                VsfData.ROW = GetStartRow(intRow)
                                Exit For
                            End If
                        Next intRow
                    End If
                End If
            Else
toMoveNextCol2:
                If VsfData.COL < mlngNoEditor - 1 Then
                    VsfData.ROW = VsfData.FixedRows
                    VsfData.COL = VsfData.COL + 1
                    If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or _
                        InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then
                        GoTo toMoveNextCol2
                    End If
                Else
                    VsfData.ROW = VsfData.FixedRows
                    VsfData.COL = mlngNoEditor - 1
                End If
            End If
            
            
            If VsfData.ColIsVisible(VsfData.COL) = False Then
                VsfData.LeftCol = VsfData.COL
            End If
            If VsfData.RowIsVisible(VsfData.ROW) = False Then
                VsfData.TopRow = VsfData.ROW
            End If
        Else
toMoveNextCol:
            If VsfData.COL < mlngNoEditor - 1 Then       '护理记录单肯定有护士签名列
                VsfData.COL = VsfData.COL + 1
                If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or _
                    InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then
                    GoTo toMoveNextCol
                End If
            Else
toMoveNextRow:
                '跳到下一行
                If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                    intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                    intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
                Else
                    intRow = 1
                End If
                If VsfData.ROW + intRow < VsfData.Rows Then
                    VsfData.ROW = VsfData.ROW + intRow
                End If
                If VsfData.RowHidden(VsfData.ROW) Then
                    If VsfData.ROW < VsfData.Rows - 1 Then
                        GoTo toMoveNextRow
                    Else
                        For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
                            If VsfData.RowHidden(intRow) = False Then
                                VsfData.ROW = GetStartRow(intRow)
                                Exit For
                            End If
                        Next intRow
                    End If
                End If
          
                VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
            End If
        End If
    Else
toMovePrevCol:
        If VsfData.COL > mlngDate Then      '护理记录单肯定有护士签名列
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '跳到上一行
'            intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
'            intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
'            If VsfData.ROW + intRow < VsfData.Rows Then
'                VsfData.ROW = VsfData.ROW + intRow
'            End If
'            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMovePrevRow
'            VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
        End If
    End If

    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '提取数据起始行,超出本页则返回0
    '如果本页未显示全,则说明超出本页,也返回0
    '不允许在连续的数据行中插入新行

    If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '总行数
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '当前行
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If

    '寻找起始行
    For lngRow = lngRow To 3 Step -1
        If Format(VsfData.TextMatrix(lngRow, mlngRowCount)) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next

    GetStartRow = lngStart
End Function

Private Function GetStartRowHistory(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '提取数据起始行,超出本页则返回0
    '如果本页未显示全,则说明超出本页,也返回0
    '不允许在连续的数据行中插入新行
    
    If InStr(1, vsfHistory.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then
        GetStartRowHistory = lngRow
        Exit Function
    End If
    
    lngRows = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(0))    '总行数
    lngCurRows = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(1)) '当前行
    If lngCurRows = 1 Then
        GetStartRowHistory = lngRow
        Exit Function
    End If

    '寻找起始行
    For lngRow = lngRow To 3 Step -1
        If Format(vsfHistory.TextMatrix(lngRow, mlngRowCount)) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next

    GetStartRowHistory = lngStart
End Function

Private Function GetMutilData(ByVal lngRow As Long, ByVal lngCol As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '起始行
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '返回第一行的坐标
    '不分行直接取，分行时检查如果当页显示全就拼接，否则从库中读取

    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilData = VsfData.TextMatrix(lngRow, lngCol)
        Exit Function
    End If
    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
    lngCurRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1))

    If lngCount > 1 Then
        lngStart = GetStartRow(lngRow)
    Else
        lngStart = lngRow
    End If
    For lngRow = lngStart To lngStart + lngCount - 1
        strReturn = strReturn & VsfData.TextMatrix(lngRow, lngCol)
    Next
    
    '取行高
    VsfData.ROW = lngStart
    dblHeight = lngCount * VsfData.RowHeightMin + 20
    dblTop = VsfData.Top + VsfData.CellTop

    GetMutilData = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetMutilDataHistory(ByVal lngRow As Long, ByVal lngCol As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '起始行
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngOrder As Long
    Dim intCount As Integer, intBound As Integer
    On Error GoTo ErrHand
    '返回第一行的坐标
    '不分行直接取，分行时检查如果当页显示全就拼接，否则从库中读取
    mblnEditHistoryAssistant = False
    mblnEditHistoryAssistant = ISEditAssistant(lngCol)
   
    If vsfHistory.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilDataHistory = vsfHistory.TextMatrix(lngRow, lngCol)
        Exit Function
    End If
    
    '获取数据起始行
    lngRow = GetStartRowHistory(lngRow)
    lngStart = lngRow
    If Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) > 0 And mblnEditHistoryAssistant Then '选择的是分组起始行(大文本)
        '获取分组数据的第一行
        If Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) > 1 Then
            lngStart = lngRow - Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) + 1
            If Val(vsfHistory.TextMatrix(lngStart, mlngDemo)) <> 1 Then
                For lngStart = lngRow To vsfHistory.FixedRows Step -1
                    If Val(vsfHistory.TextMatrix(lngStart, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngStart
                If lngStart < vsfHistory.FixedRows Then lngStart = lngRow
            End If
        End If
        lngRow = lngStart
        lngCount = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(0))
        intBound = lngRow + lngCount - 1
        
        For intCount = lngRow + lngCount To vsfHistory.Rows - 1
            If intCount > intBound Then
                If Val(vsfHistory.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '不分组或遇新分组就退出
                intBound = Val(Split(vsfHistory.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
            End If
            lngCount = lngCount + 1
        Next
    Else
        lngCount = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(0))
    End If
    
    lngStart = lngRow
    strReturn = ""
    For lngRow = lngStart To lngStart + lngCount - 1
        strReturn = strReturn & vsfHistory.TextMatrix(lngRow, lngCol)
    Next
    strReturn = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
    '取行高
    vsfHistory.ROW = lngStart
    dblHeight = lngCount * vsfHistory.RowHeightMin + 20
    dblTop = vsfHistory.Top + vsfHistory.CellTop

    GetMutilDataHistory = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowInput(Optional ByVal intCol As Integer = -1, Optional ByVal strCellData As String = "", Optional ByVal blnAnalyse As Boolean = False) As String
    Dim arrData, arrValue
    Dim lngOrder As Long
    Dim blnMake As Boolean
    Dim i As Integer, j As Integer, intPos As Integer, intIndex As Integer
    Dim strFormat As String, strText As String, strValue As String  '格式串,数据串,数值串
    Dim strOrders As String, strTypes As String, strBounds As String
    Dim strLen As String, strName As String, strState As String
    Const txtHeight = 300
    Dim str缺省 As String
    Dim arr缺省() As String
    
    On Error GoTo ErrHand

    '病历文件构造管理模块需要处理:
    '1、一列绑定一个项目的不用管
    '2、一列绑定两个项目的，血压必须成对，要么都是录入，要么都是选择，不允许交叉出现，也不允许出现单选、复选
    '3、一列绑定多个项目的，只能是录入项目
    '由于以上条件限制，只取第一个项目的性质即可

    '如果是保存处调用则做如下处理
    If intCol = -1 Then intCol = VsfData.COL
    If blnAnalyse Then
        strText = strCellData
    Else
        '取当前单元格的属性
        CellRect.Left = VsfData.CellLeft + VsfData.Left
        CellRect.Top = VsfData.CellTop + VsfData.Top
        CellRect.Bottom = VsfData.CellHeight + 20
        CellRect.Right = VsfData.CellWidth + 20
        strText = GetMutilData(VsfData.ROW, intCol, CellRect.Top, CellRect.Bottom)
    End If
    strText = Replace(Replace(Replace(strText, Chr(10), ""), Chr(13), ""), Chr(1), "")
    mstrData = strText
    mintType = 0
    intIndex = 0

    '取当前列的绑定项目
    intPos = 1
    mrsSelItems.Filter = "列=" & intCol - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!项目序号
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If

        '项目表示:2单选;3-多选;4-汇总;5-选择
        '项目值域:项目表示为0-表示最小值;最大值;项目表示为2,3-表示项目A;项目B,前有勾的表示缺省项
        strFormat = NVL(mrsSelItems!格式)
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "项目序号=" & lngOrder
            strName = strName & "," & mrsItems!项目名称
            strLen = strLen & "," & mrsItems!项目长度 & ";" & NVL(mrsItems!项目小数)
            strState = strState & "," & mrsItems!项目类型
            strTypes = strTypes & "," & mrsItems!项目表示
            strBounds = strBounds & "," & mrsItems!项目值域
            str缺省 = str缺省 & "," & mrsItems!缺省值
            strValue = strValue & "'" & SubstrVal(strText, strFormat, GetActivePart(intCol, intIndex) & mrsItems!项目名称, intPos)

            Select Case mrsItems!项目表示
            Case 0  '文本录入项
                If mrsSelItems.RecordCount = 2 Then
                    If InStr(1, strState & ",", ",1,") = 0 Then
                        mintType = 4
                    Else
                        mintType = 6
                    End If
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '单选
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 1
                ElseIf mrsSelItems.RecordCount = 2 Then
                    mintType = 7
                End If
            Case 3  '多选
                mintType = 2
            Case 4  '汇总
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  '选择
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 3
                Else
                    mintType = 5
                End If
            End Select
        Else
            strState = strState & ","
            strTypes = strTypes & ","
            strBounds = strBounds & ","
            strLen = strLen & ","
            strName = strName & ","
        End If

        intIndex = intIndex + 1
        mrsSelItems.MoveNext
    Loop
    If strOrders <> "" Then
        strOrders = Mid(strOrders, 2)
        strName = Mid(strName, 2)
        strLen = Mid(strLen, 2)
        strState = Mid(strState, 2)
        str缺省 = Mid(str缺省, 2)
        strTypes = Mid(strTypes, 2)
        strBounds = Mid(strBounds, 2)
        strValue = Mid(strValue, 2)
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    strValue = Replace(Replace(Replace(strValue, Chr(10), ""), Chr(13), ""), Chr(1), "")
    If blnAnalyse Then
        ShowInput = strOrders & "||" & strValue
        Exit Function
    End If

    '针对4进行校对,如果表头文本不含/则处理为6
    If mintType = 4 Then
        If Not IsDiagonal(intCol) Then
            mintType = 6
        End If
    End If

    '判断当前列的性质
    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定2个及以上项目,手工录入
    arrValue = Split(strValue & "'", "'")
    Select Case mintType
    Case 0, 3
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                If ScaleHeight - picMain.Top - .Height < 0 Then
                    .Height = ScaleHeight - picMain.Top
                Else
                    .Top = ScaleHeight - picMain.Top - .Height
                End If
            End If
            
            .Visible = True
            .ZOrder 0
        End With
        If mintType = 0 Then
            txtInput.Visible = True
            If Val(strLen) <> 0 And Val(strOrders) <> 10 Then
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) '小数位数要加上小数点
            Else
                txtInput.MaxLength = 0
            End If
            txtInput.Tag = lngOrder
        Else
            txtInput.Visible = False
        End If
        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '宋体9号时减去90,字体越大扣除的边距越小,以保证文本框分行与实际一致
        End With
        With lblInput
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = CellRect.Bottom
            .Width = CellRect.Right
            .Top = 50
            .Tag = lngOrder
            .Caption = strValue
            .Visible = (mintType = 3)
        End With

        '如果是日期或时间列，设定固定值
        If mintType = 0 And txtInput.Text = "" Then
            Dim lngStart As Long
            blnMake = True
            If intCol = mlngDate Then
                If VsfData.ROW > VsfData.FixedRows Then
                    lngStart = GetStartRow(VsfData.ROW - 1)
                    blnMake = (VsfData.TextMatrix(lngStart, mlngDate) = "")
                End If
                If blnMake Then
                    If mblnDateAd Then
                        txtInput.Text = Format(zlDatabase.Currentdate, "d-M")
                        txtInput.Text = Replace(txtInput.Text, "-", "/")
                    Else
                        txtInput.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                    End If
                Else
                    txtInput.Text = VsfData.TextMatrix(lngStart, mlngDate)
                End If
            ElseIf intCol = mlngTime Then
                If VsfData.ROW > VsfData.FixedRows Then
                    lngStart = GetStartRow(VsfData.ROW - 1)
                    blnMake = (VsfData.TextMatrix(lngStart, mlngTime) = "")
                End If
                If blnMake Then
                    txtInput.Text = Format(zlDatabase.Currentdate, "HH:mm")
                Else
                    txtInput.Text = VsfData.TextMatrix(lngStart, mlngTime)
                End If
            End If
        End If
        '95807,CL
        '加载未记说明
        lstSelect(2).Clear
        mrsUsual.Filter = ""
        Do While Not mrsUsual.EOF
            lstSelect(2).AddItem mrsUsual!名称
            mrsUsual.MoveNext
        Loop
        With lstSelect(2)
            .Left = CellRect.Left
            .Top = CellRect.Top + picInput.Height - 20
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * (PicLst.TextHeight("刘")) + PicLst.TextHeight("刘") \ 3
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                If ScaleHeight - picMain.Top - .Height < 0 Then
                    .Top = 10
                    .Height = IIf(ScaleHeight - picMain.Top - 10 < PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, ScaleHeight - picMain.Top - 10)
                Else
                    .Top = ScaleHeight - picMain.Top - .Height
                End If
            End If
        End With
    Case 1, 2
        '56439:刘鹏飞,2012-11-30,单选项目如果未设置缺省想，默认定位到清除选择，以前的方式是定位到
        '实际数据项，对于有些项目不需要录入，就会要操作员手工选择到清除选择项，在操作上很麻烦。
        '加载数据
        lstSelect(mintType - 1).Clear
        If mintType = 1 Then lstSelect(mintType - 1).AddItem "清除选择"
        If strBounds = "" Then strBounds = ";"
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If str缺省 = arrData(i) Then
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & arrData(i)
                    If strText = "" Then lstSelect(mintType - 1).ListIndex = lstSelect(mintType - 1).NewIndex
                Else
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        '多选且已录入数据的情况下
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            txtLst.Text = strValue
            PicLst.Tag = "1"
            j = lstSelect(mintType - 1).ListCount - 1
            For i = 0 To j
                '单选的第一个项目是清除选择，需要跳过此项,多选项目则直接进入
                If Not (mintType = 1 And i = 0) Then
                    If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(mintType - 1).List(i), InStr(1, lstSelect(mintType - 1).List(i), "-") + 1) & ",") <> 0 Then
                        lstSelect(mintType - 1).Selected(i) = True
                        txtLst.Text = ""
                        PicLst.Tag = "0"
                    End If
                End If
            Next
        Else
            txtLst.Text = ""
            PicLst.Tag = "0"
        End If
        '控件显示
        '51134,刘鹏飞,2012-07-11,单选提供文本录入
        PicLst.FontName = VsfData.FontName
        PicLst.FontSize = VsfData.FontSize
        If mintType = 1 Then
        
            With PicLst
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = LenB(StrConv(lstSelect(mintType - 1).List(lstSelect(mintType - 1).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                If .Width < CellRect.Right Then .Width = CellRect.Right
            End With
            
            With lbllst(0)
                .Left = 20
                .Top = 20
                If .Width > PicLst.Width Then
                    PicLst.Width = .Width + PicLst.TextWidth("刘")
                End If
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Visible = True
            End With
            
            With txtLst
                .Top = lbllst(0).Top + lbllst(0).Height + 20
                .Left = -10
                .Width = PicLst.Width
                
                If .Text <> "" Then
                    txtLength.Width = .Width
                    txtLength.Text = Replace(Replace(Replace(.Text, Chr(10), ""), Chr(13), ""), Chr(1), "")
                    txtLength.FontName = VsfData.CellFontName
                    txtLength.FontSize = VsfData.CellFontSize
                    txtLength.FontBold = VsfData.CellFontBold
                    txtLength.FontItalic = VsfData.CellFontItalic
                    arrData = GetData(txtLength.Text)
                    .Text = Join(arrData, "")
                    .Height = (UBound(arrData) + 1) * (VsfData.CellHeight + 20)
                Else
                    .Height = VsfData.CellHeight + 20
                End If
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                If strLen <> "" Then
                    .MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) '小数位数要加上小数点
                End If
                .Visible = True
            End With
            
            With lbllst(1)
                .Left = 20
                .Top = txtLst.Top + txtLst.Height + 20
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Visible = True
            End With
            
            With PicLst
                .Height = lbllst(1).Top + lbllst(1).Height + 20 + lstSelect(mintType - 1).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3)
                If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
                If .Height + .Top + picMain.Top > ScaleHeight Then
                    If ScaleHeight - picMain.Top - .Height < 0 Then
                        .Top = 10
                        .Height = ScaleHeight - picMain.Top - 10
                    Else
                        .Top = ScaleHeight - picMain.Top - .Height
                    End If
                End If
                
                .Visible = True
                .ZOrder 0
            End With
            
            With lstSelect(mintType - 1)
                .Top = lbllst(1).Top + lbllst(1).Height + 20
                .Left = -10
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Width = PicLst.Width
                .Height = PicLst.Height - .Top
                .Tag = lngOrder
                If .Top + .Height <> PicLst.Height Then
                    PicLst.Height = .Top + .Height
                End If
                .Visible = True
            End With
        Else
            '显示
            With lstSelect(mintType - 1)
                .Left = CellRect.Left
                .Top = CellRect.Top
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Height = .ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3)
                If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
                .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                If .Width < CellRect.Right Then .Width = CellRect.Right
                If .Height + .Top + picMain.Top > ScaleHeight Then
                    If ScaleHeight - picMain.Top - .Height < 0 Then
                        .Top = 10
                        .Height = ScaleHeight - picMain.Top - 10
                    Else
                        .Top = ScaleHeight - picMain.Top - .Height
                    End If
                End If
                .Tag = lngOrder
                .Visible = True
                .ZOrder 0
            End With
        End If
    Case 4, 5
        With picDouble
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
            .ZOrder 0
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If

        With txtUpInput
            .Text = arrValue(0)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = (picDouble.Width - lblSplit.Width) * 0.4
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(0)
        End With
        With picUpInput
            .Left = txtUpInput.Left
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(0)
        End With
        With lblUpInput
            .Alignment = 2
            .Caption = arrValue(0)
            .Left = 0
            .Top = 50
            .FontName = VsfData.FontName
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .Tag = Split(strOrders, ",")(0)
        End With
        With txtDnInput
            .Text = arrValue(1)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Left = lblSplit.Left + lblSplit.Width
            .Width = picDouble.Width - .Left
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(1)
        End With
        With picDnInput
            .Left = txtDnInput.Left
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(1)
        End With
        With lblDnInput
            .Alignment = 2
            .Caption = arrValue(1)
            .Left = 0
            .Top = 50
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .FontName = VsfData.FontName
            .Tag = Split(strOrders, ",")(1)
        End With

        If mintType = 4 Then
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
        End If
        
        '加载未记说明
        lstSelect(2).Clear
        mrsUsual.Filter = ""
        Do While Not mrsUsual.EOF
            lstSelect(2).AddItem mrsUsual!名称
            mrsUsual.MoveNext
        Loop
        With lstSelect(2)
            .Left = CellRect.Left
            .Top = CellRect.Top + picDouble.Height - 20
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * (PicLst.TextHeight("刘")) + PicLst.TextHeight("刘") \ 3
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                If ScaleHeight - picMain.Top - .Height < 0 Then
                    .Top = 10
                    .Height = IIf(ScaleHeight - picMain.Top - 10 < PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, ScaleHeight - picMain.Top - 10)
                Else
                    .Top = ScaleHeight - picMain.Top - .Height
                End If
            End If
        End With
    Case 6
        '先删除以前的控件
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '设定坐标
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '对缺省控件赋值
        arrData = Split(strOrders, ",")
        j = UBound(arrData)
        lbl(0).Top = 130
        lbl(0).Caption = Split(strName, ",")(0)
        lbl(0).FontName = VsfData.FontName
        lbl(0).FontSize = VsfData.FontSize
        txt(0).Tag = arrData(0)
        txt(0).FontName = VsfData.FontName
        txt(0).FontSize = VsfData.FontSize
        txt(0).Width = picMutilInput.Width - txt(0).Left - 100
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  '小数位数要加上小数点
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If

        '加载控件
        For i = 1 To j
            Load lbl(i)
            With lbl(i)
                .Caption = Split(strName, ",")(i)
                .Left = lbl(0).Left + lbl(0).Width - .Width
                .Top = lbl(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Visible = True
            End With
            Load txt(i)
            With txt(i)
                .TabIndex = txt(i - 1).TabIndex + 1
                .Left = txt(0).Left
                .Top = txt(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Tag = arrData(i)
                If strLen <> "" Then
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
                End If
                .Text = arrValue(i)
                .Visible = True
            End With
        Next

        With picMutilInput
            .Height = txt(j).Top + txt(j).Height + 120
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            If .Top < 0 Then .Top = 0
            .Visible = True
            .ZOrder 0
        End With
        
        '加载未记说明
        lstSelect(2).Clear
        mrsUsual.Filter = ""
        Do While Not mrsUsual.EOF
            lstSelect(2).AddItem mrsUsual!名称
            mrsUsual.MoveNext
        Loop
        With lstSelect(2)
            .Left = picMutilInput.Left + picMutilInput.Width
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * (PicLst.TextHeight("刘")) + PicLst.TextHeight("刘") \ 3
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                If ScaleHeight - picMain.Top - .Height < 0 Then
                    .Top = 10
                    .Height = IIf(ScaleHeight - picMain.Top - 10 < PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, ScaleHeight - picMain.Top - 10)
                Else
                    .Top = ScaleHeight - picMain.Top - .Height
                End If
            End If
        End With
    Case 7
        cboChoose(0).Clear
        cboChoose(0).FontName = VsfData.FontName
        cboChoose(0).FontSize = VsfData.FontSize
        cboChoose(0).Tag = Split(strOrders, ",")(0)
        arr缺省 = Split(str缺省, ",")
        arrData = Split(Split(strBounds, ",")(0), ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) = arr缺省(0) Then
                cboChoose(0).AddItem arrData(i)
                If strValue = "'" Then
                    cboChoose(0).ListIndex = i
                Else
                    If arrData(i) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            Else
                cboChoose(0).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            End If
        Next
        
        cboChoose(1).Clear
        cboChoose(1).FontName = VsfData.FontName
        cboChoose(1).FontSize = VsfData.FontSize
        cboChoose(1).Tag = Split(strOrders, ",")(1)
        arrData = Split(Split(strBounds, ",")(1), ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) = arr缺省(1) Then
                cboChoose(1).AddItem arrData(i)
                If strValue = "'" Then
                    cboChoose(1).ListIndex = i
                Else
                    If arrData(i) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            Else
                cboChoose(1).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            End If
        Next
        
        With picDoubleChoose
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
            .ZOrder 0
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDoubleChoose.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        picChooseRight.Left = lblSplit.Left + 150
        cboChoose(0).SetFocus
    End Select
    Exit Function

ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CheckFormat(ByVal strNames As String, ByVal strFormat As String)
    '如果格式与血压的方式不同,则将样式处理为6

    '去掉前缀后进行对比
    strFormat = Mid(strFormat, InStr(1, strFormat, "["))
    strFormat = Replace(strFormat, "[", "")
    strFormat = Replace(strFormat, "]", "")
    If Not (strFormat Like Split(strNames, ",")(0) & "/}*" Or strFormat Like "{/*" & Split(strNames, ",")(1)) Then
        mintType = 6
    End If
End Sub

Private Function IsDiagonal(ByVal intCol As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '判断指定列是否设置了列对角线（mstrColWidth的格式：765`11`1`1,765`11`2`1,...，对象属性`对象序号`列对角线）

    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer, intType As Integer
    Dim objParent As Object
    Dim intRow As Integer, intCount As Integer, i As Integer, intGroupFirstRows As Integer, intHidden As Integer
    Dim strText As String, lngCount As Long
    Dim arrData, lngStartRow As Long
    '根据项目的长度决定是否允许进行词句选择
    mblnEditAssistant = False
    mblnEditText = False
    cmdWord.Visible = mblnEditAssistant
    
    mrsItems.Filter = "项目序号=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    intType = mintType
    mblnEditAssistant = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100) 'And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) <= "1"
    mblnEditText = (mrsItems!项目类型 = 1 And NVL(mrsItems!项目表示, 0) = 0)
    If mblnEditText = True And mblnEditAssistant = False Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            cmdWord.Tag = -1  '表示txtInput
        Else
            cmdWord.Tag = objTXT.Index
        End If
    End If
    mrsItems.Filter = 0
    lngStartRow = VsfData.ROW
    '获取分组数据的第一行
    If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1 Then
        lngStartRow = VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
            For lngStartRow = VsfData.ROW To VsfData.FixedRows Step -1
                If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                    Exit For
                End If
            Next lngStartRow
            If lngStartRow < VsfData.FixedRows Then Exit Sub
        End If
    End If
    
    '如果允许词句选择,显示并定位
    If mblnEditAssistant Then
        mintType = -1
        VsfData.ROW = lngStartRow
        mintType = intType
        
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '表示txtInput
            Set objParent = picInput
        Else
            intIndex = objTXT.Index
            Set objParent = picMutilInput
        End If
        With cmdWord
            .Tag = intIndex
            .Top = objParent.Top + objTXT.Top + 25
            .Left = objParent.Left + objTXT.Left + objTXT.Width - .Width + 25
            .Visible = True
            .ZOrder 0
        End With
        strText = ""
        intCount = 0
        intHidden = 0
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        '为分组行时，选择数据起始行，编辑内容显示所有大文本行
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 Then
            For intRow = 0 To intGroupFirstRows - 1
                strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(intRow + VsfData.ROW, VsfData.COL), Chr(13), ""), Chr(10), ""), Chr(1), "")
            Next intRow
            lngCount = VsfData.ROW + intGroupFirstRows - 1
            For intRow = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                If VsfData.RowHidden(intRow) = False Then
                    '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                    If intRow > lngCount Then
                        If Val(VsfData.TextMatrix(intRow, mlngDemo)) <= 1 Then
                            'If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 2)
                            Exit For
                        End If
                        lngCount = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0)) + intRow - 1
                    End If
                    strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(intRow, VsfData.COL), Chr(13), ""), Chr(10), ""), Chr(1), "")
                    'strText = strText & IIf(intRow > VsfData.ROW And strText <> "", vbCrLf, "") & Replace(VsfData.TextMatrix(intRow, VsfData.COL), vbCrLf, "")
                Else
                    lngCount = lngCount + 1
                    intHidden = intHidden + 1
                End If
            Next intRow
            '准备赋值
            With txtLength
                '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
                .Width = VsfData.CellWidth
                .Text = Replace(Replace(Replace(strText, Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = VsfData.CellFontName
                .FontSize = VsfData.CellFontSize
                .FontBold = VsfData.CellFontBold
                .FontItalic = VsfData.CellFontItalic
            End With
            arrData = GetData(txtLength.Text)
            intCount = UBound(arrData)
            strText = ""
            For i = 0 To intCount
                strText = strText & CStr(arrData(i))
            Next i
            intRow = intRow - VsfData.ROW - intHidden
            picInput.Height = intRow * VsfData.RowHeightMin + 20
            If picInput.Height + picInput.Top + picMain.Top > ScaleHeight Then
                picInput.Top = ScaleHeight - picMain.Top - picInput.Height
            End If
            txtInput.Height = picInput.Height
            txtInput.Text = strText
            mstrData = Replace(Replace(Replace(strText, Chr(10), ""), Chr(13), ""), Chr(1), "")
            Call zlControl.TxtSelAll(txtInput)
            lblInput.Height = picInput.Height
        End If
    End If
End Sub

Private Sub FillPage()
    Dim strPatient As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCSQL As String
    On Error GoTo ErrHand
    '读取符合条件的病人清单(在院病人+最近几天转科病人+指定时间范围内出院病人),病人清单决定了行数
    
    '58890:刘鹏飞,2013-02-26,在院病人读取性能优化(关联在院病人表进行查询)
    '在院病人清单
    strPatient = "" & _
        " SELECT 1 AS 性质,B.病人ID, B.主页ID, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, B.住院号, lpad(B.出院病床,10,' ') AS 床号,0 AS 婴儿" & _
        " FROM 病人信息 A,病案主页 B,在院病人 C" & _
        " Where A.病人ID = B.病人ID And A.主页ID=B.主页ID And NVL(b.主页ID, 0) <> 0 " & _
        " And Nvl(B.状态,0)<>1 AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL And A.病人ID=C.病人ID And C.病区ID=[3] " & _
        IIf(mlng科室ID = -1, "", " And C.科室ID=[4]")
    If chk出院.Value = 1 Then
        '最近几天出院病人清单
        strPatient = strPatient & _
            " UNION " & _
            " SELECT 3 AS 性质,B.病人ID, B.主页ID, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, B.住院号, lpad(B.出院病床,10,' ') AS 床号,0 AS 婴儿" & _
            " FROM 病人信息 A,病案主页 B" & _
            " Where A.病人ID = b.病人ID And NVL(b.主页ID, 0) <> 0 And b.当前病区ID + 0 = [3]" & _
            " AND B.出院日期 BETWEEN [1] AND [2] AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL" & _
            IIf(mlng科室ID = -1, "", " And B.出院科室ID+0=[4]")
    End If
    If chk出科.Value = 1 Then
        '最近几天转科病人清单
        strPatient = strPatient & _
            " UNION " & _
            " Select 2 AS 性质,B.病人ID, B.主页ID, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, B.住院号, lpad(c.床号,10,' ') AS 床号,0 AS 婴儿" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 " & _
            " And Nvl(B.状态,0)<>2 And Nvl(C.附加床位,0)=0 " & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.病区ID+0=[3]" & IIf(mlng科室ID = -1, "", " And B.出院科室ID<>[4] And C.科室ID+0=[4]") & _
            " And C.终止原因=3 And C.终止时间 Between Sysdate-" & mintChange & " And Sysdate" & _
            " And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    End If
    '提取新生儿列表
    strPatient = strPatient & _
              " UNION " & _
              " Select B.性质,B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,B.住院号,lpad(b.床号,10,' ') AS 床号,A.序号 AS 婴儿" & _
              " From 病人新生儿记录 A,(" & strPatient & ") B" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    
    If mstrBPItem = "" Then
        gstrSQL = " SELECT  A.性质,A.病人ID,A.主页ID,A.婴儿,A.姓名,lpad(a.床号,10,' ') AS 床号,MAX(B.ID) AS 文件ID,'' 血压频次" & _
                  " FROM (" & strPatient & ") A,病人护理文件 B" & _
                  " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.婴儿=B.婴儿 " & _
                  " And B.归档人 is null And B.结束时间 is null And B.格式ID=[5]" & _
                  " GROUP BY A.性质,A.病人ID,A.主页ID,A.婴儿,A.姓名 ,A.床号" & _
                  " Order by A.性质,A.床号"
    Else
        strCSQL = ",(Select 血压频次" & vbNewLine & _
                    "From (Select Distinct 病人id, 主页id, NVL(婴儿,0) 婴儿, First_Value(b.英文名称) Over(Partition By 病人id, 主页id, NVL(婴儿,0) Order By 开始执行时间 Desc) 血压频次" & vbNewLine & _
                    "       From 病人医嘱记录 a, 诊疗频率项目 b, (Select Column_Value From Table(f_Num2list('" & mstrBPItem & "'))) c" & vbNewLine & _
                    "       Where ((a.医嘱期效 = 0 And a.医嘱状态 In (3, 5, 6, 7,8,9) And ((a.执行终止时间 Is Null Or a.执行终止时间 >= Sysdate))) Or (a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8)" & vbNewLine & _
                    "       And  a.开始执行时间 Between Sysdate-1 And Sysdate)) And a.执行频次 = b.名称 And" & vbNewLine & _
                    "             a.诊疗项目id = c.Column_Value) C Where A.病人ID=C.病人ID And A.主页ID=C.主页ID And A.婴儿=C.婴儿) 血压频次"

        gstrSQL = " SELECT /*+ Rule */  A.性质,A.病人ID,A.主页ID,A.婴儿,A.姓名,lpad(a.床号,10,' ') AS 床号,MAX(B.ID) AS 文件ID" & strCSQL & _
                  " FROM (" & strPatient & ") A,病人护理文件 B" & _
                  " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.婴儿=B.婴儿 " & _
                  " And B.归档人 is null And B.结束时间 is null And B.格式ID=[5]" & _
                  " GROUP BY A.性质,A.病人ID,A.主页ID,A.婴儿,A.姓名 ,A.床号" & _
                  " Order by A.性质,A.床号"
    End If
    
    
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人清单", mdtOutbegin, mdtOutEnd, mlng病区ID, mlng科室ID, mlng格式ID, mstrBPItem)
    
    '填充数据到表格
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > VsfData.Rows - VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
            
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c文件ID) = !文件ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c床号) = Trim(NVL(!床号))
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c姓名) = IIf(!婴儿 > 0, Space(4), "") & !姓名
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c病人ID) = !病人ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c主页ID) = !主页ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c婴儿) = !婴儿
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c血压频次) = NVL(!血压频次)
            If mlngRowCount < VsfData.Cols Then VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, mlngRowCount) = ""
            .MoveNext
        Loop
    End With
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'以下是基础函数或过程
Private Sub lblDnInput_Click()
    txtDnInput.SetFocus
End Sub

Private Sub lblUpInput_Click()
    txtUpInput.SetFocus
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    Dim i As Integer, j As Integer
    mblnEditAssistant = False
    mblnEditText = False
    PicLst.Tag = 0
    j = lstSelect(Index).ListCount - 1
    If Index = 0 And j >= 0 Then
        If lstSelect(Index).ListIndex < 0 Then lstSelect(Index).ListIndex = 0
    End If
End Sub

Private Sub txtDnInput_GotFocus()
    txtDnInput.SelStart = 0
    txtDnInput.SelLength = 100
    Call ISAssistant(Val(txtDnInput.Tag), txtDnInput)
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    Call zlControl.TxtSelAll(txtInput)
    mintSymbol = -1
    Call ISAssistant(Val(txtInput.Tag), txtInput)
End Sub

Private Sub txtUpInput_GotFocus()
    txtUpInput.SelStart = 0
    txtUpInput.SelLength = 100
    Call ISAssistant(Val(txtUpInput.Tag), txtUpInput)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = 100
    mintSymbol = Index
    Call ISAssistant(Val(txt(Index).Tag), txt(Index))
End Sub

Private Sub lblUpInput_DblClick()
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "√", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyLeft And txtUpInput.SelStart = 0 Then
        Call MoveNextCell(False)
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    End If
    
    If KeyCode = vbKeyDown And IsCanves(Val(txtUpInput.Tag)) Then
        lstSelect(2).Visible = True
        lstSelect(2).Tag = 2
        Call LstChoose(txtUpInput.Text)
        lstSelect(2).SetFocus
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyRight And txtDnInput.SelStart = Len(txtDnInput.Text)) Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
    End If
    
    If KeyCode = vbKeyDown And IsCanves(Val(txtUpInput.Tag)) Then
        lstSelect(2).Visible = True
        lstSelect(2).Tag = 2
        Call LstChoose(txtUpInput.Text)
        lstSelect(2).SetFocus
    End If
End Sub

Private Sub picMutilInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picDouble_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not txtInput.Visible Then
        If KeyCode = vbKeySpace Then
            Call lblInput_DblClick
        End If
    End If

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        '移动到下一个单元格
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Index <> 2 Then Call MoveNextCell
    If Index = 0 And Shift = vbShiftMask And KeyCode = vbKeyUp Then
        KeyCode = 0
        txtLst.SetFocus
    End If
    
    If Index = 2 And (KeyCode = vbKeyUp Or KeyCode = vbKeyReturn) Then
        KeyCode = 0
        If lstSelect(2).ListIndex >= 0 Then
            Select Case Val(lstSelect(2).Tag)
            Case 1
                txtInput.Text = lstSelect(2).Text
                txtInput.SetFocus
                lstSelect(2).Visible = False
                
            Case 2
                txtUpInput.Text = lstSelect(2).Text
                txtUpInput.SetFocus
                lstSelect(2).Visible = False
            Case 3
                txtDnInput.Text = lstSelect(2).Text
                txtDnInput.SetFocus
                lstSelect(2).Visible = False
            Case 4
                txt(picMutilInput.Tag).Text = lstSelect(2).Text
                txt(picMutilInput.Tag).SetFocus
                lstSelect(2).Visible = False
            End Select
        End If
    End If
End Sub

Private Sub picMutilInput_GotFocus()
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txt.Count - 1 Then
            txt(Index + 1).SetFocus
        Else
            Call picMutilInput_KeyDown(KeyCode, Shift)
        End If
    End If
    
    If KeyCode = vbKeyDown Then
        If IsCanves(txt(Index).Tag) Then
            lstSelect(2).Visible = True
            lstSelect(2).Tag = 4
            picMutilInput.Tag = Index
            Call LstChoose(txt(Index).Text)
            lstSelect(2).SetFocus
        End If
    End If
End Sub

Private Sub picDouble_GotFocus()
    If txtUpInput.Visible Then
        txtUpInput.SetFocus
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    VsfData.Width = picMain.Width
    VsfData.Height = IIf(picMain.Height - VsfData.Top < 0, 0, picMain.Height - VsfData.Top)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    If KeyCode = vbKeyReturn Or _
        (KeyCode = vbKeyRight And txtInput.SelStart = Len(txtInput.Text)) Or _
        (KeyCode = vbKeyLeft And txtInput.SelStart = 0) Then
        Call picInput_KeyDown(KeyCode, Shift)
    End If
    
    If KeyCode = vbKeyDown And IsCanves(Val(txtInput.Tag)) Then
        KeyCode = 0
        lstSelect(2).Visible = True
        lstSelect(2).Tag = 1
        Call LstChoose(txtInput.Text)
        lstSelect(2).SetFocus
    End If
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Then
        KeyAscii = 0
        txtDnInput.SetFocus
    End If
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnSigned = False
    mblnSaved = False
    mblnChange = False
    mblnInit = False

'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '以下字符做为数据分隔符或更新记录集的分隔符，因此不允许录入
    If KeyAscii = 39 Or KeyAscii = 13 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)

    Err = 0: On Error Resume Next
    If Not mblnInit Then picSplit.Top = ScaleHeight - 3000
    picSplit.Left = VsfData.Left
    picSplit.Width = VsfData.Width
    
    lblTitle.Move lngScaleLeft, 120, lngScaleRight - lngScaleLeft
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, picSplit.Top - lngScaleTop
    VsfData.Move lngScaleLeft + 210, lblTitle.Top + lblTitle.Height + 300, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
    
    vsfHistory.Left = VsfData.Left
    vsfHistory.Top = picSplit.Top + picSplit.Height
    vsfHistory.Height = lngScaleBottom - picSplit.Top - 50
    vsfHistory.Width = VsfData.Width
    
    picNull.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom
    lblInfo(0).Move lngScaleLeft, (lngScaleBottom - lngScaleTop) / 2 - lblInfo(0).Height, lngScaleRight
    lblInfo(1).Move lngScaleLeft, (lngScaleBottom - lngScaleTop) / 2 + 100, lngScaleRight
    
    '表上标签分散处理
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub cboChoose_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            cboChoose(1).SetFocus
        Else
            Call MoveNextCell
        End If
    End If
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long

    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

Private Sub SetActiveColColor()
    '活动列的背景色设置为灰色,表示不允许编辑
    Dim aryItem, lngRow As Long
    aryItem = Split(mstrCOLNothing, ",")
    For lngRow = 0 To UBound(aryItem)
        VsfData.Cell(flexcpBackColor, VsfData.FixedRows, Val(aryItem(lngRow)) + cHideCols, VsfData.Rows - 1, Val(aryItem(lngRow)) + cHideCols) = &H8000000F
        '.ColHidden(Val(aryItem(lngCount)) + cHideCols) = True
    Next
End Sub

Private Sub vsfHistory_DblClick()
    Call vsfHistory_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub vsfHistory_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCellHistory(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub vsfHistory_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------
'修改人：LPF 2012-04-20
'修改内容：允许非分组起始行，也可以录入多行数据
'----------------------------------------------
    Dim arrData, i As Integer
    Dim blnNULL As Boolean                      '是否为空行
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngMutilRows As Long, lngDeff As Long, lngStartGroup As Long
    Dim intRow As Integer, intCount As Integer, intNULL As Integer, intBound As Integer, intRowCount As Integer  '其后有多少空行
    Dim dblTop As Long, dblHeight As Long, intRowGroup As Integer
    Dim blnGroup As Boolean, blnDate As Boolean, strRows As String, lngDemo As Long, lngLastNull As Long, lngLastNoNull As Long
    Dim strDate As String, strTime As String
    Dim strKey As String, strField As String, strValue As String, strCols As String
    Dim intGroupFirstRows As Integer, blnTrue As Boolean
    Dim blnReseGroupAssistant As Boolean, blnGroupAddNum As Boolean '分组数据增加行
    '大文本列和内容信息
    Dim varAssistant() As Variant, strAssistantCols As String
    On Error GoTo ErrHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Not mblnInit Then Exit Sub
    '将历史表格当前单元格的数据复制到当前表格中
    '赋值然后移动到下一个有效单元格
    '获取历史数据
    If vsfHistory.COL <= mlngTime Or vsfHistory.COL >= mlngNoEditor Then Exit Sub
    If Val(vsfHistory.TextMatrix(vsfHistory.ROW, mlngRecord)) = 0 Then Exit Sub
    
    cmdWord.Visible = False
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
        lstSelect(2).Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
        If mintType = 1 Then
            txtLst.Visible = False
            PicLst.Visible = False
        End If
    Case 4, 5
        picDouble.Visible = False
        lstSelect(2).Visible = False
    Case 6
        picMutilInput.Visible = False
        lstSelect(2).Visible = False
    Case 7
        picDoubleChoose.Visible = False
    End Select
    mintType = -1: mblnShow = False: blnReseGroupAssistant = False
    
    strReturn = GetMutilDataHistory(vsfHistory.ROW, vsfHistory.COL, dblTop, dblHeight)
    If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then
        lngStart = GetStartRow(VsfData.ROW)
        For i = 0 To UBound(Split(mstrColCollect, "|"))
            strValue = GetRelatiionNo(CStr(Split(mstrColCollect, "|")(i)), 2)
            strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(Split(mstrColCollect, "|")(i), ";")(0)
        Next
        strCols = Mid(strCols, 2)
        If InStr(1, "," & strCols & ",", "," & vsfHistory.COL - (cHideCols + vsfHistory.FixedCols - 1) & ",") > 0 Then
            If ISCollectSigned(Val(VsfData.TextMatrix(lngStart, c文件ID)), Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "HH:MM")) Then
                RaiseEvent AfterRowColChange("您要粘贴的数据为汇总列数据，且目标数据所对应的汇总数据已签名，数据将不能被粘贴。", True)
                Exit Sub
            End If
        End If
    End If
    mblnEditAssistant = mblnEditHistoryAssistant
    VsfData.COL = vsfHistory.COL
    If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 0 And mblnEditHistoryAssistant Then '选择的是分组起始行(大文本)
        '获取分组数据的第一行
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1 Then
            lngStart = VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then
                For lngStart = VsfData.ROW To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngStart
                If lngStart < VsfData.FixedRows Then lngStart = VsfData.ROW
            End If
            VsfData.ROW = lngStart
        End If
    End If
    blnDate = (InStr(1, "," & mlngDate & "," & mlngTime & ",", "," & VsfData.COL & ",") > 0)
    lngDemo = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo))
    blnGroup = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) >= 1
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
     '如果修改的是非大文本列或时间列的分组数据，检查修改内容行数是否发生变化，如果变化就当分组数据处理，否则以普通数据处理
    If blnGroup = True And Not (mblnEditAssistant = True Or blnDate = True) Then
        intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngStart = GetStartRow(VsfData.ROW)
        '如果编辑的是分组数据最后一个分组行，则当普通数据处理
        If lngStart + intGroupFirstRows < VsfData.Rows Then
            If Val(VsfData.TextMatrix(lngStart + intGroupFirstRows, mlngDemo)) <= 1 Then
                blnGroup = False: GoTo ErrBegin
            End If
        ElseIf lngStart + intGroupFirstRows >= VsfData.Rows Then
            blnGroup = False
            GoTo ErrBegin
        End If
        
        With txtLength
            '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
            .Width = VsfData.CellWidth
            .Text = Replace(Replace(Replace(IIf(strReturn = "", "a", strReturn), Chr(10), ""), Chr(13), ""), Chr(1), "")
            .FontName = VsfData.CellFontName
            .FontSize = VsfData.CellFontSize
            .FontBold = VsfData.CellFontBold
            .FontItalic = VsfData.CellFontItalic
        End With
        arrData = GetData(txtLength.Text)
        
        blnGroup = False
        intBound = -1
        If (UBound(arrData) + 1) > intGroupFirstRows Then
            blnGroup = True
        ElseIf (UBound(arrData) + 1) < intGroupFirstRows Then
            '得到本条数据占用最大行的列(不包含大文本项目)
            blnNULL = True
            For intRow = lngStart + intGroupFirstRows - 1 To lngStart Step -1
                For intCount = 0 To mlngNoEditor - 1
                    If VsfData.ColHidden(intCount) = False And ISEditAssistant(intCount) = False Then
                        If VsfData.TextMatrix(intRow, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(intRow, intCount), "/") <> 0) Then
                            blnNULL = False
                            If intCount = VsfData.COL Then
                                intBound = intCount
                            Else
                                intBound = intCount
                                Exit For
                            End If
                        End If
                    End If
                Next intCount
                If blnNULL = False Then Exit For
            Next intRow
            
            If blnNULL = False Then
                intNULL = intRow - lngStart + 1
                If intBound = VsfData.COL Then
                    blnGroup = True
                Else
                    blnGroup = (intNULL < intGroupFirstRows)
                End If
            Else
                blnGroup = True
            End If
        End If
    End If
ErrBegin:
    blnTrue = False
    lngMutilRows = 1
    intGroupFirstRows = 1
    If Not blnGroup Then
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
            lngMutilRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        End If
        lngStart = GetStartRow(VsfData.ROW)
    Else
        lngMutilRows = 1
        If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = 1 And (mblnEditAssistant Or blnDate) Then
            '记录分组起始行的数据行数
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            intBound = VsfData.ROW + intGroupFirstRows - 1
            For intCount = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                If intCount > intBound Then
                    If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '不分组或遇新分组就退出
                    intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
                End If
                lngMutilRows = lngMutilRows + 1
            Next
            lngMutilRows = lngMutilRows + intGroupFirstRows - 1 '保证数据行数的准确性
        Else
            '记录分组起始行的数据行数
            If VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1 >= VsfData.FixedRows Then
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngRowCount), "|")(0))
            End If
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            lngMutilRows = intGroupFirstRows
        End If
        lngStart = VsfData.ROW
    End If
    
    '准备赋值
    With txtLength
        '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
        .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
        .Text = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
        .FontName = VsfData.CellFontName
        .FontSize = VsfData.CellFontSize
        .FontBold = VsfData.CellFontBold
        .FontItalic = VsfData.CellFontItalic
    End With
    arrData = GetData(txtLength.Text)
    intCount = UBound(arrData)
    If intCount = -1 Then
        arrData = Array()
        ReDim Preserve arrData(UBound(arrData) + 1)
        arrData(UBound(arrData)) = ""
        intCount = 1
    End If
    lngLastNull = VsfData.ROW + intGroupFirstRows - 1: lngLastNoNull = VsfData.ROW + intGroupFirstRows - 1
    '分组数据中可能存在隐藏的行(点击清除功能时),只有是选择大文本才处理
    If blnGroup = True And mblnEditAssistant = True And blnDate = False Then
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 Then
            intNULL = intGroupFirstRows
            lngDeff = VsfData.ROW + intGroupFirstRows - 1
            For intRow = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                If intRow > lngDeff Then
                    If Val(VsfData.TextMatrix(intRow, mlngDemo)) <= 1 Or intNULL > intCount Then Exit For     '不分组或遇新分组就退出
                    lngDeff = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0)) + intRow - 1
                End If
                If VsfData.RowHidden(intRow) = True Then  '删除的分组列
                    '重新组织文本数组，保证分组数据正常复制
                    ReDim Preserve arrData(UBound(arrData) + 1)
                    For intBound = UBound(arrData) To intRow - VsfData.ROW + 1 Step -1
                        arrData(intBound) = arrData(intBound - 1)
                    Next intBound
                    arrData(intRow - VsfData.ROW) = ""
                    '记录最后一次隐藏的行
                    lngLastNull = intRow
                Else
                    intNULL = intNULL + 1
                    '记录最后一次没有隐藏的行
                    lngLastNoNull = intRow
                End If
            Next
        End If
    End If
    intCount = UBound(arrData)
    
    lngDeff = 0
    blnGroupAddNum = False
    blnTrue = blnGroup = True And mblnEditAssistant
    If intCount > lngMutilRows - 1 Then
        '对于新增分组数据时，必须要先录入完分组数据才能录入大文本段数据
'        If mblnEditAssistant = True And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 And lngMutilRows = intGroupFirstRows Then
'            strMsg = "新增分组数据时，请先完成数据的分组，最后在录入大文段项目内容！"
'            RaiseEvent AfterRowColChange(strMsg, True)
'            strMsg = ""
'            Exit Sub
'        End If
        '往下搜索空行,如果有其它数据行则计算需增加的行数
        '20110830分组号算做同一数据行，将多行文本分解到各行，多余的文本放在统一放在最后一行上;在非首行按回车,只对现有数据进行修改,不对行发生变化
        intNULL = intCount - (lngMutilRows - 1)
        For intRow = lngMutilRows To intCount
            '保证当前输入的内容在一页中显示全
            If intRow + lngStart > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(intRow + lngStart, c病人ID)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                intNULL = intNULL - 1
                If VsfData.RowHidden(intRow + lngStart) = True Then VsfData.RowHidden(intRow + lngStart) = False
            Else
                Exit For
            End If
        Next
        '先增加空行
        If intNULL > 0 Then
            lngDeff = intNULL
            VsfData.Rows = VsfData.Rows + intNULL
            '从当前行记录的空白行开始，每行的位置+所增加的空白行数
            For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                VsfData.RowPosition(intRow) = intRow + intNULL
            Next
            For intRow = lngMutilRows To intCount
                VsfData.TextMatrix(intRow + lngStart, c文件ID) = VsfData.TextMatrix(lngStart, c文件ID)
                VsfData.TextMatrix(intRow + lngStart, c病人ID) = VsfData.TextMatrix(lngStart, c病人ID)
                VsfData.TextMatrix(intRow + lngStart, c主页ID) = VsfData.TextMatrix(lngStart, c主页ID)
                VsfData.TextMatrix(intRow + lngStart, c婴儿) = VsfData.TextMatrix(lngStart, c婴儿)
            Next intRow
        End If
        
        '当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
        If lngDeff <> 0 Then
            If Not blnGroup Then
                Call CellMap_Update(lngStart, lngDeff)
            Else
                Call CellMap_Update(lngStart + lngMutilRows - 1, lngDeff)  '分组行数据从最大一条明细行之后开始处理
            End If
        End If
        
        '对分组数据最后一个分组行为隐藏行的处理
        '例如：该分组具有2个分组，第一组为一行大文本内容为A，第二组为一行大文本内容为C(改行隐藏).此时添加大文本内容为A、B占两行，此时组织得到本组的数据为A、C、B
        '计算方式为:第一行占用内容+隐藏行内容+多出的内容。此处就会把隐藏行放在最后，最后得到的内容为A、B、C，在下面循环赋值中，第一组就为占用2行内容为A、B
        '说明：如果中间存在隐藏行，最后一组没有隐藏，多出的数据就会追加在最后一组数据的后面
        If (lngLastNull - lngLastNoNull) > 0 Then
            For intRow = lngLastNoNull + 1 To lngLastNull
                strValue = arrData(lngLastNoNull + 1 - VsfData.ROW)
                For intBound = lngLastNoNull + 1 - VsfData.ROW To UBound(arrData) - 1
                    arrData(intBound) = arrData(intBound + 1)
                Next intBound
                arrData(UBound(arrData)) = strValue
                VsfData.RowPosition(lngLastNoNull + 1) = lngLastNull + (intCount - (lngMutilRows - 1))
            Next intRow
            '更新记录
            For intRow = lngLastNull To lngLastNoNull + 1 Step -1
                Call CellMap_Update(intRow, intCount - (lngMutilRows - 1), False)
            Next intRow
        End If
        
        '循环赋值
        intCount = UBound(arrData)
        intBound = 0
        blnReseGroupAssistant = (blnGroup = True And Not (mblnEditAssistant Or blnDate))
        blnGroupAddNum = blnReseGroupAssistant
        If blnGroup = True And blnDate = False Then strReturn = ""
        For intRow = 0 To intCount
            VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
            '修改非分组数据或非大文本或日期的分组数据
            '非分组数据：直接处理计算行数并更新数据
            '非大文本或日期的分组数据：1、直接处理计算行数并更新数据，2、需要重新处理大文段的内容显示位置
            If (Not blnGroup) Then
                VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
                VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intCount + 1
                If intRow > 0 And intRow < intCount Then
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = ""
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = ""
                ElseIf intRow = intCount Then
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = VsfData.TextMatrix(lngStart, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = VsfData.TextMatrix(lngStart, mlngSignTime)
                End If
            Else
                '修改分组大文段或日期，需从分组起始行到分组结束行从新整理文本内容显示或日期
                '分组行的特殊处理,更新内部记录集的代码较多
                '##########################################
                If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                    If (mblnEditAssistant = True Or blnDate) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                    End If
                    intRowCount = 1
                    '获取该分组数据行的行数
                    For intBound = intRow + 1 To intCount
                         If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) >= 1 Then Exit For
                         intRowCount = intRowCount + 1
                    Next intBound
                    intBound = intRow
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intRowCount & "|1"
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intRowCount
                    If Not blnDate Then strReturn = ""
                Else
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intRowCount & "|" & intRow - intBound + 1
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = ""
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = ""
                End If
                If Not blnDate Then
                    strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStart + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
                End If
                '到该分组数据行数最后一行才执行更新操作
                If intRow = intBound + intRowCount - 1 Then
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = VsfData.TextMatrix(lngStart + intBound, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = VsfData.TextMatrix(lngStart + intBound, mlngSignTime)
                    '保存数据
                    If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                        If CheckGroupDate(lngStart + intBound) = True Then
                            '保存后的修改才进入此流程，取该条记录的实际时间
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                        Else
                            '新增时进入此流程
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    Else
                        '普通数据
                        strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                        strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                    End If
                    
                    '1\日期
                    strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                    If mlngDate <> -1 Then
                        strKey = mint页码 & "," & lngStart + intBound & "," & mlngDate
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                    '2\时间
                    strKey = mint页码 & "," & lngStart + intBound & "," & mlngTime
                    strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                        VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    
                    If Not blnDate Then
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                            strPart = GetActivePart(VsfData.COL, 0)
                        Else
                            strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                        End If
                        strKey = mint页码 & "," & lngStart + intBound & "," & VsfData.COL
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                End If
                '##########################################
            End If
        Next
        '所有隐蔽列进行赋值
        intBound = lngStart + intCount
        For intRow = lngStart + 1 To intBound
            For intCount = 0 To VsfData.Cols - 1
                VsfData.Cell(flexcpForeColor, intRow, intCount) = VsfData.Cell(flexcpForeColor, lngStart, intCount)
                If VsfData.ColHidden(intCount) And InStr(1, "," & mlngRowCount & "," & mlngRowCurrent & ",", "," & intCount & ",") = 0 Then
                    If blnGroup And InStr(1, "," & mlngDemo & "," & mlngRecord & "," & mlngActiveTime & ",", "," & intCount & ",") = 0 Then
                        VsfData.TextMatrix(intRow, intCount) = VsfData.TextMatrix(lngStart, intCount)
                    End If
                End If
            Next
        Next
        lngMutilRows = lngStart + lngMutilRows - 1
    Else
        blnReseGroupAssistant = False
        If blnGroup = True And blnDate = False Then strReturn = ""
        '对该列重新赋值（当只输入一个数字时，不知为何会产生字符ASCII码为1的符号）
        For intRow = 0 To intCount
            VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
            If blnGroup = True Then
                '分组行的特殊处理,更新内部记录集的代码较多
                '##########################################
                If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                    If (mblnEditAssistant = True Or blnDate = True) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                    End If
                    intRowCount = 1
                    '获取该分组数据行的行数
                    For intBound = intRow + 1 To intCount
                         If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) >= 1 Then Exit For
                         intRowCount = intRowCount + 1
                    Next intBound
                    intBound = intRow
                    If Not blnDate Then strReturn = ""
                End If
                If Not blnDate Then
                    strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStart + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
                End If
                '到该分组数据行数最后一行才执行更新操作
                If intRow = intBound + intRowCount - 1 Then
                    '保存数据
                    If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                        If CheckGroupDate(lngStart + intBound) = True Then
                            '保存后的修改才进入此流程，取该条记录的实际时间
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                        Else
                            '新增时进入此流程
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    Else
                        '普通数据
                        strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                        strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                    End If
                    
                    '1\日期
                    strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                    If mlngDate <> -1 Then
                        strKey = mint页码 & "," & lngStart + intBound & "," & mlngDate
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                    '2\时间
                    strKey = mint页码 & "," & lngStart + intBound & "," & mlngTime
                    strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                        VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    
                    If Not blnDate Then
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                            strPart = GetActivePart(VsfData.COL, 0)
                        Else
                            strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                        End If
                        strKey = mint页码 & "," & lngStart + intBound & "," & VsfData.COL
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                End If
                '##########################################
            End If
        Next
        strRows = ""
        lngStartGroup = -1
        For intRow = intCount + 1 To lngMutilRows - 1
            VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
        Next intRow
        blnReseGroupAssistant = False
        If intCount < (lngMutilRows - 1) Then
            blnReseGroupAssistant = (blnGroup And Not (mblnEditAssistant Or blnDate))
        End If
        For intRow = intCount + 1 To lngMutilRows - 1
            '分组行的特殊处理,更新内部记录集的代码较多
            '##########################################
            '保存数据
            If (blnGroup And (mblnEditAssistant Or blnDate)) Then
                '获取改行起始行
                If lngStartGroup <> GetStartRow(lngStart + intRow) Then
                    intNULL = GetStartRow(lngStart + intRow)
                    '寻找的起始列mlngDemo肯定>0
                    If Val(VsfData.TextMatrix(intNULL, mlngDemo)) <= 0 Then
                        For intRowGroup = lngStart + intRow To lngStart Step -1
                            If Val(VsfData.TextMatrix(intRowGroup, mlngDemo)) >= 0 Then
                                intNULL = intRowGroup
                                Exit For
                            End If
                        Next intRowGroup
                        If intNULL = lngStartGroup Then GoTo ErrDemo
                    End If
                    lngStartGroup = intNULL
                    '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
                    If VsfData.TextMatrix(lngStartGroup, mlngRowCount) = "" Then VsfData.TextMatrix(lngStartGroup, mlngRowCount) = "1|1"
                    intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartGroup, mlngRowCount), "|")(0))
                    intNULL = lngStartGroup + intGroupFirstRows - 1
                    For intRowGroup = intNULL To lngStartGroup Step -1
                        blnNULL = True
                        For intBound = 0 To VsfData.Cols - 1
                            If Not VsfData.ColHidden(intBound) And intBound < mlngNoEditor Then
                                If VsfData.TextMatrix(intRowGroup, intBound) <> "" And Not (IsDiagonal(intBound) And InStr(1, VsfData.TextMatrix(intRowGroup, intBound), "/") <> 0) Then
                                    blnNULL = False
                                    Exit For
                                End If
                            End If
                        Next
                        If Not blnNULL Then Exit For
                        intNULL = intNULL - 1
                        If intRowGroup = lngStartGroup Then
                             intNULL = intNULL + 1
                        Else
                            If InStr(1, strRows & ",", "," & intRowGroup & ",") = 0 Then strRows = strRows & "," & intRowGroup
                        End If
                    Next intRowGroup
                    
                    '重新填写数据行数
                    For intRowGroup = lngStartGroup To intNULL
                        VsfData.TextMatrix(intRowGroup, mlngRowCount) = intNULL - lngStartGroup + 1 & "|" & intRowGroup - lngStartGroup + 1
                        VsfData.TextMatrix(intRowGroup, mlngRowCurrent) = intNULL - lngStartGroup + 1
                    Next intRowGroup
                    If mlngSignName <> -1 Then
                        If Trim(VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignName)) <> "" Then
                            VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignName)
                            If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignTime)
                        End If
                    End If
                    For intRowGroup = intNULL + 1 To lngStartGroup + intGroupFirstRows - 1
                        VsfData.TextMatrix(intRowGroup, mlngRowCount) = ""
                        VsfData.TextMatrix(intRowGroup, mlngRowCurrent) = ""
                        VsfData.TextMatrix(intRowGroup, mlngRecord) = ""
                        If mlngSignName <> -1 Then VsfData.TextMatrix(intRowGroup, mlngSignName) = ""
                        If mlngOperator <> -1 Then VsfData.TextMatrix(intRowGroup, mlngOperator) = ""
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(intRowGroup, mlngSignTime) = ""
                    Next
                End If
ErrDemo:
                If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) > 0 And intRow > intCount Then
                    VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                    If CheckGroupDate(lngStart + intRow) = True Then
                        '保存后的修改才进入此流程，取该条记录的实际时间
                        If mblnDateAd Then
                            strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "MM")
                        Else
                            strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 1, 10)
                        End If
                        strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 12, 5)
                    Else
                        '新增时进入此流程
                        strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                        strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                    End If
                
                    '分组起始行的行数减少时，重新设置分组号
                    If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
                    If strRows <> "" Then
                        intNULL = 0
                        For intBound = 0 To UBound(Split(strRows, ","))
                            If Val(Split(strRows, ",")(intBound)) < (lngStart + intRow) Then
                                intNULL = intNULL + 1
                            End If
                        Next intBound
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) = Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) - intNULL
                    End If
                    
                    '1\日期
                    strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                    If mlngDate <> -1 Then
                        strKey = mint页码 & "," & lngStart + intRow & "," & mlngDate
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                    '2\时间
                    strKey = mint页码 & "," & lngStart + intRow & "," & mlngTime
                    strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & mlngTime & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strTime & "|" & _
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    
                    If Not blnDate Then
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                            strPart = GetActivePart(VsfData.COL, 0)
                        Else
                            strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                        End If
                        strKey = mint页码 & "," & lngStart + intRow & "," & VsfData.COL
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & VsfData.COL & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & "" & "|" & strPart & "|1"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                End If
            End If
            '##########################################
        Next
        '修改分组非大文本或日期行时，需要获取分组数据大文本段内容信息，重新组织文本显示
        '如有3组数据，第二2行有3行，修改为1行，第3组数据应该紧接着显示在第2组下面(第二组此时只有1行)
        If blnReseGroupAssistant = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
        lngMutilRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
        intNULL = lngStart + lngMutilRows - 1
        For intRow = lngMutilRows To 1 Step -1
            blnNULL = True
            For intCount = 0 To VsfData.Cols - 1
                If Not VsfData.ColHidden(intCount) And intCount < mlngNoEditor And IIf(blnReseGroupAssistant = True, ISEditAssistant(intCount) = False, True) Then
                    If VsfData.TextMatrix(intRow + lngStart - 1, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(intRow + lngStart - 1, intCount), "/") <> 0) Then
                        blnNULL = False
                        Exit For
                    End If
                End If
            Next
            
            If Not blnNULL Then Exit For
            intNULL = intNULL - 1
        Next
        '从新填写行序号
        If Not blnGroup Then
            If intNULL < lngStart Then intNULL = lngStart
            For intRow = lngStart To intNULL
                VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
            Next
            If mlngSignName <> -1 Then
                If Trim(VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignName)) <> "" Then
                    VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignTime)
                End If
            End If
            strRows = ""
        Else '分组行以保存的数据删除时，不清空行号
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                If intRow = lngStart Then intNULL = intNULL + 1
            Next intRow
            
            For intRow = lngStart To intNULL
                VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
            Next
            If mlngSignName <> -1 Then
                If Trim(VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignName)) <> "" Then
                    VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignTime)
                End If
            End If
        End If
        If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
        If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
        For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
            VsfData.TextMatrix(intRow, mlngRowCount) = ""
            VsfData.TextMatrix(intRow, mlngRowCurrent) = ""
            VsfData.TextMatrix(intRow, mlngRecord) = ""
            If mlngSignName <> -1 Then VsfData.TextMatrix(intRow, mlngSignName) = ""
            If mlngOperator <> -1 Then VsfData.TextMatrix(intRow, mlngOperator) = ""
            If mlngSignTime <> -1 Then VsfData.TextMatrix(intRow, mlngSignTime) = ""
            If blnReseGroupAssistant = True Then
                If InStr(1, strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
            ElseIf Not blnGroup Then
                If InStr(1, strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
            End If
        Next
        '更新记录集大文段信息
        If blnReseGroupAssistant = True Then Call CellMap_UpdateAssistant(lngStart)
    End If
    
    '获取分组起始行所有行信息
    If blnTrue = True Then 'blnTrue为真说明选择的是分组行的起始行，并且是大文本段
        strReturn = ""
        intCount = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        For intRow = 0 To intCount - 1
            strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(VsfData.ROW + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
        Next intRow
    End If
    mblnChange = True
           
    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 0 Then
        If CheckGroupDate(lngStart) = True Then
            '保存后的修改才进入此流程，取该条记录的实际时间
            If mblnDateAd Then
                strDate = Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "MM")
            Else
                strDate = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10)
            End If
            strTime = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 12, 5)
        Else
            '新增时进入此流程
            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
        End If
    Else
        '普通数据
        strDate = VsfData.TextMatrix(lngStart, mlngDate)
        strTime = VsfData.TextMatrix(lngStart, mlngTime)
    End If
    
    '1\日期
    strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
    If mlngDate <> -1 Then
        strKey = mint页码 & "," & lngStart & "," & mlngDate
        strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    End If
    '2\时间
    strKey = mint页码 & "," & lngStart & "," & mlngTime
    strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngTime & "|" & _
        Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|" & _
        VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
          
    If Not blnGroup Or blnTrue Then
        '记录用户修改过的单元格
        If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
            strPart = GetActivePart(VsfData.COL, 0)
        Else
            strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
        End If
        
        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
        strKey = mint页码 & "," & lngStart & "," & VsfData.COL
        strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & VsfData.COL & "|" & _
            Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    End If
    Call SetActiveColColor
    
    '数据行数减少时，将空白行移至到最后一行
    If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
    If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
    strRows = Replace("," & strRows & ",", ",,", "")
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
            Call CellMap_Update(intRow, -1)
            VsfData.TextMatrix(intRow, mlngDemo) = ""
        End If
    Next intRow
    
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
           '清空改行所有信息
            For intBound = 0 To VsfData.Cols - 1
                VsfData.TextMatrix(intRow, intBound) = ""
            Next intBound
            VsfData.RowHidden(intRow) = True
            VsfData.RowPosition(intRow) = VsfData.Rows - 1
        End If
    Next intRow

    '重新组织分组数据内容
    If blnReseGroupAssistant = True Then
        If blnGroupAddNum = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
        If strAssistantCols <> "" Then
            Call ReSetGroupAssistant(True, False, strAssistantCols, varAssistant)
        Else
            Call ReSetGroupDemo(lngStart)
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AppendGroup(ByVal lngStartRow As Long)
    Dim lngDemo As Long, lngStart As Long, lngRows As Long
    Dim blnGroup As Boolean
    '追加分组行(只能在单数据行后追加分组行)
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
    lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
    '先检查当前行是否为分组行
    blnGroup = (VsfData.TextMatrix(lngStartRow, mlngDemo) <> "")
    If Not blnGroup Then
        lngDemo = 1
        VsfData.TextMatrix(lngStartRow, mlngDemo) = 1
    Else
        lngDemo = VsfData.TextMatrix(lngStartRow, mlngDemo)
    End If
    VsfData.TextMatrix(lngStartRow + lngRows, mlngDemo) = lngDemo + lngRows
    VsfData.ROW = lngStartRow + lngRows
    lngStart = VsfData.ROW - VsfData.TextMatrix(lngStartRow + lngRows, mlngDemo) + 1
End Sub

Private Function ISEditAssistant(ByVal lngCol As Long) As Boolean
'是否编辑的是大文本项目
    Dim blnTrue As Boolean, lngOrder As Long
    
    mrsSelItems.Filter = "列=" & lngCol - cHideCols
    If mrsSelItems.RecordCount > 0 Then
        lngOrder = Val(mrsSelItems!项目序号)
        mrsItems.Filter = "项目序号=" & lngOrder
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            Exit Function
        End If
        blnTrue = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100)
    End If
    ISEditAssistant = blnTrue
End Function

Private Sub ReSetGroupAssistant(blnNoMove As Boolean, blnNext As Boolean, ByVal strAssistantCols As String, varAssistantText() As Variant)
'功能：重新排列大文本列在每一行的数据
'说明：对于修改非大文段或日期时间列的分组数据时才调用(先调用GetGroupAssistant方法在调用此方法)
    Dim lngCol As Long, lngRow As Long, lngStartRow As Long, varCol
    Dim lngOldRow As Long, lngOldCol As Long, intType As Integer, blnTrue As Boolean
    Dim strText As String, blnOldNoMove As Boolean, blnOldNext As Boolean
    
    lngOldRow = VsfData.ROW
    lngOldCol = VsfData.COL
    intType = mintType
    blnOldNoMove = blnNoMove
    blnOldNext = blnNext
    
    '获取编辑当前行的起始行行
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
        lngRow = GetStartRow(VsfData.ROW)
    Else
        lngRow = VsfData.ROW
    End If
    '获取分组数据的第一行
    lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
        For lngStartRow = lngRow To VsfData.FixedRows Step -1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                lngRow = lngStartRow
                Exit For
            End If
        Next lngStartRow
        If lngRow < VsfData.FixedRows Then Exit Sub
        lngStartRow = lngRow
    End If
    
    If Left(strAssistantCols, 1) = "," Then strAssistantCols = Mid(strAssistantCols, 2)
    varCol = Split(strAssistantCols, ",")
    For lngCol = 0 To UBound(varAssistantText)
        strText = CStr(varAssistantText(lngCol))
        mintType = -1: mblnShow = False
        VsfData.ROW = lngStartRow
        VsfData.COL = Val(varCol(lngCol))
        mblnEditAssistant = True
        blnTrue = True
        mintType = 0
        Call MoveNextCell(False, True, strText)
        mintType = -1
    Next lngCol
    
    '恢复列
    If blnTrue = True Then
        VsfData.ROW = lngOldRow
        VsfData.COL = lngOldCol
        mintType = intType
    End If
    mblnEditAssistant = False
    mblnShow = True
    
    blnNoMove = blnOldNoMove
    blnNext = blnOldNext
End Sub

Private Sub GetGroupAssistant(strAssistantCols As String, varAssistantText() As Variant)
'功能：获取大文本段信息
'说明：对于修改非大文段或日期时间列的分组数据时才调用
    Dim lngRow As Long, lngCol As Long, lngOrder As Long, intGroupFirstRows As Integer, lngCount As Long
    Dim lngStartRow As Long
    Dim strText As String
    
    strAssistantCols = ""
    varAssistantText = Array()
    
    '获取编辑当前行的起始行行
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
        lngRow = GetStartRow(VsfData.ROW)
    Else
        lngRow = VsfData.ROW
    End If
    '获取分组数据的第一行
    lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
        For lngStartRow = lngRow To VsfData.FixedRows Step -1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                lngRow = lngStartRow
                Exit For
            End If
        Next lngStartRow
        If lngStartRow < VsfData.FixedRows Then Exit Sub
        lngStartRow = lngRow
    End If
    
    For lngCol = mlngTime + 1 To mlngNoEditor - 1
        '寻找大文本列
        mrsSelItems.Filter = "列=" & lngCol - cHideCols
        If mrsSelItems.RecordCount > 0 Then
            lngOrder = Val(mrsSelItems!项目序号)
            mrsItems.Filter = "项目序号=" & lngOrder
            If mrsItems.RecordCount = 0 Then
                mrsItems.Filter = 0
                GoTo ErrNext
            End If
             
            mblnEditAssistant = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100) And Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <= 1
            If Not mblnEditAssistant Then GoTo ErrNext
                
            If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            '为分组行时，选择数据起始行，编辑内容显示所有大文本行
            strText = ""
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                For lngRow = 0 To intGroupFirstRows - 1
                    strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow + lngStartRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                Next lngRow
                lngCount = lngStartRow + intGroupFirstRows - 1
                For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
                    If VsfData.RowHidden(lngRow) = False Then
                        '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                        If lngRow > lngCount Then
                            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For  '不分组或遇新分组就退出
                            If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                            lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
                        End If
                        strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                    Else
                        lngCount = lngCount + 1
                    End If
                Next lngRow
                
                If strText = "" Then GoTo ErrNext
                strAssistantCols = strAssistantCols & "," & lngCol
                ReDim Preserve varAssistantText(UBound(varAssistantText) + 1)
                varAssistantText(UBound(varAssistantText)) = strText
            End If
ErrNext:
        End If
    Next lngCol
End Sub

Private Sub ReSetGroupDemo(ByVal lngRow As Long)
'功能：设置分组行的行号和记录集信息
'在修改分组行数据时，如果包含大文本切文本内容不为空通过GetGroupAssistant和ReSetGroupAssistant完成设置，如果没有则调用此函数完成设置
    Dim strDate As String, strTime As String
    Dim intNULL As Integer, lngStartRow As Long, lngRowCount As Long, blnNULL As Boolean
    Dim lngCurRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    Dim intGroupFirstRows As Integer
    Dim varAssistant() As Variant, strAssistantCols As String
    
    If Val(VsfData.TextMatrix(lngRow, mlngRowCount)) > 1 Then
        lngStartRow = GetStartRow(lngRow)
    Else
        lngStartRow = lngRow
    End If
    '确定分组起始行
    lngRow = lngStartRow
    lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
        For lngStartRow = lngRow To VsfData.FixedRows Step -1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                lngRow = lngStartRow
                Exit For
            End If
        Next lngStartRow
        lngStartRow = lngRow
    End If
    '重新组织分组序号
    VsfData.TextMatrix(lngStartRow, mlngDemo) = 1
    intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
    lngCurRow = lngStartRow
    For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
        If lngRow = lngCurRow + intGroupFirstRows Then
            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                Exit For
            Else
                VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - Val(lngStartRow) + 1
            End If
            If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
            lngCurRow = lngRow
        End If
    Next

    '从起始行开始处理分组数据记录集
    intGroupFirstRows = 0
    lngCurRow = lngStartRow
    For lngRow = lngStartRow To VsfData.Rows - 1
        If lngRow = lngCurRow + intGroupFirstRows Then
            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 And intGroupFirstRows > 0 Then Exit For
            If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
            lngCurRow = lngRow
            If CheckGroupDate(lngRow) = True Then
                '保存后的修改才进入此流程，取该条记录的实际时间
                If mblnDateAd Then
                    strDate = Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "MM")
                Else
                    strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 10)
                End If
                strTime = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 12, 5)
            Else
                '新增时进入此流程
                strDate = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngDate)
                strTime = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngTime)
            End If
            
            '1\日期
            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
            If mlngDate <> -1 Then
                strKey = mint页码 & "," & lngRow & "," & mlngDate
                strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngRow) = True, 1, 0)
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\时间
            strKey = mint页码 & "," & lngRow & "," & mlngTime
            strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strTime & "|" & _
                VsfData.TextMatrix(lngRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngRow) = True, 1, 0)
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
    Next lngRow
End Sub

Private Sub CellMap_UpdateAssistant(ByVal lngStartRow As Long)
'功能：更新记录集大文段信息
    Dim strDate As String, strTime As String
    Dim strKey As String, strField As String, strValue As String, strPart As String
    Dim lngCol As Long, lngRow As Long, lngRowCount As Long, strReturn As String
    
    On Error GoTo ErrHand
    
    If VsfData.TextMatrix(lngStartRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
    lngRowCount = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
    
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
        If CheckGroupDate(lngStartRow) = True Then
            '保存后的修改才进入此流程，取该条记录的实际时间
            If mblnDateAd Then
                strDate = Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "MM")
            Else
                strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 1, 10)
            End If
            strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 12, 5)
        Else
            '新增时进入此流程
            strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
            strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
        End If
    Else
        '普通数据
        strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
        strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
    End If
    
    '1\日期
    strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
    If mlngDate <> -1 Then
        strKey = mint页码 & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    End If
    '2\时间
    strKey = mint页码 & "," & lngStartRow & "," & mlngTime
    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & _
        VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    
    For lngCol = mlngTime + 1 To mlngNoEditor - 1
        If ISEditAssistant(lngCol) Then
            strReturn = ""
            For lngRow = 0 To lngRowCount - 1
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStartRow + lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next lngRow
            '记录用户修改过的单元格
            If InStr(1, "," & mstrCatercorner & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                strPart = GetActivePart(lngCol, 0)
            Else
                strPart = GetActivePart(lngCol, 0) & "/" & GetActivePart(lngCol, 1)
            End If
            
            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
            strKey = mint页码 & "," & lngStartRow & "," & lngCol
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
    Next lngCol
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckGroupDate(ByVal lngRow As Long) As Boolean
'--功能：检查分组数据起始行时间和保存时间是否相等
    Dim strDate As String, strTime As String
    Dim strDate1 As String, strTime1 As String
    Dim lngStart As Long
    
    lngStart = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    
    If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
        If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then CheckGroupDate = True: Exit Function
        strDate = VsfData.TextMatrix(lngStart, mlngDate)
        strTime = VsfData.TextMatrix(lngStart, mlngTime)
        If mblnDateAd Then
            strDate1 = Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "MM")
        Else
            strDate1 = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10)
        End If
        strTime1 = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 12, 5)
        If strDate <> strDate1 Or strTime <> strTime1 Then
            CheckGroupDate = False
        Else
            CheckGroupDate = True
        End If
    Else
        CheckGroupDate = False
    End If
End Function

Private Function ISGroupAppend() As Boolean
'追加分组数据，在选择的行有数据才能追加（不包含大文本项目）
    Dim lngCol As Long, lngRow As Long
    Dim blnNULL As Boolean
    
    lngRow = VsfData.ROW
    If lngRow > VsfData.Rows - 1 Then lngRow = VsfData.Rows - 1
    blnNULL = True
    For lngCol = mlngTime + 1 To VsfData.Cols - 1
        If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
            If VsfData.TextMatrix(lngRow, lngCol) <> "" And Not (IsDiagonal(lngCol) And InStr(1, VsfData.TextMatrix(lngRow, lngCol), "/") <> 0) Then
                blnNULL = False
                Exit For
            End If
        End If
    Next
    
    ISGroupAppend = Not blnNULL
End Function

Private Sub SingerShowType(ByVal vsfObj As VSFlexGrid, ByVal lngStartRow As Long, ByVal lngEndRow As Long, Optional ByVal blnUnSingMe As Boolean = False)
'-------------------------------------------------
'功能：护士签名人显示方式
''--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
'-------------------------------------------------
    Dim lngRow As Integer
    '取消签名
    If blnUnSingMe = True Then
        For lngRow = lngStartRow To lngEndRow
            If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
            If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
        Next
        Exit Sub
    End If
    Select Case mlngSingerType
        Case 0 '所有行显示
            For lngRow = lngStartRow To lngEndRow
                If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
            Next
        Case 1 '首行显示
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case 3 '尾行显示
            If mlngOperator > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngOperator) = "" Then vsfObj.TextMatrix(lngStartRow, mlngOperator) = vsfObj.TextMatrix(lngEndRow, mlngOperator)
            End If
            If mlngSignName > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignName) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignName) = vsfObj.TextMatrix(lngEndRow, mlngSignName)
            End If
            If mlngSignTime > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignTime) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignTime) = vsfObj.TextMatrix(lngEndRow, mlngSignTime)
            End If
            For lngRow = lngEndRow To lngStartRow Step -1
                If lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case Else '首尾显示
            '最后一行需要填写封闭签名
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Or lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
    End Select
End Sub

Private Function GetRelatiionNo(ByVal strKey As String, Optional ByVal bytType As Byte = 1, Optional ByVal blnCorrelative As Boolean = True) As String
'---------------------------------------------------
'功能:获取汇总项目关联的名称列的项目序号或列号(分类汇总)
'strKey 汇总项目的列号和序号,格式:列号,序号
'bytType 1:项目序号,2:列号
'blnCorrelative TRUE:分类汇总,FALSE:入量导入
'返回值:为空表示汇总项目没有设置关联列
'---------------------------------------------------
    Dim arrItem, arrCorrelative, i As Long
    Dim strValue As String
    
    If blnCorrelative = True Then
        arrItem = Split(mstrColCorrelative, "|")
    Else
        arrItem = Split(mstrColImCorrelative, "|")
    End If
    arrItem = Split(mstrColCorrelative, "|")
    For i = 0 To UBound(arrItem)
        arrCorrelative = Split(arrItem(i), ";")
        If InStr(1, strKey, ";") <> 0 Then
            strKey = Split(strKey, ";")(0) & "," & Split(strKey, ";")(1)
        End If
        If blnCorrelative = True Then
            If strKey = arrCorrelative(1) Then
                If bytType = 1 Then
                    strValue = Split(arrCorrelative(0), ",")(1)
                Else
                    strValue = Split(arrCorrelative(0), ",")(0)
                End If
                Exit For
            End If
        Else
            If strKey = arrCorrelative(0) Then
                If bytType = 1 Then
                    strValue = Split(arrCorrelative(0), ",")(1)
                Else
                    strValue = Split(arrCorrelative(0), ",")(0)
                End If
                Exit For
            End If
        End If
    Next i
    
    GetRelatiionNo = strValue
End Function

Private Function CheckCollectIsData(ByVal lngStartRow As Long, Optional ByVal bytMode As Byte = 0, Optional ByRef lngEditCol As Long = 0) As Boolean
'功能:检查汇总列及关联列是否存在数据，只要有一列存在就退出
'入参：bytMode：主要针对分组数据，是检查整个分组数据还是值检查子数据：0-整个都检查,1- 只检查子数据
'出参：汇总列不为空则返回行号
    Dim strCols As String, strValue As String
    Dim i As Integer, arrCol
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngRowCount As Long
    If mstrColCollect <> "" Then
        '1、获取汇总相关列号
        arrCol = Split(mstrColCollect, "|")
        For i = 0 To UBound(arrCol)
            strValue = GetRelatiionNo(CStr(arrCol(i)), 2)
            strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(arrCol(i), ";")(0)
        Next
        strCols = Mid(strCols, 2)
        
        lngStartRow = GetStartRow(lngStartRow)
        '2、检查对应的列是否存在汇总数据
         '如果lngStartRow不是分组起始行，首先获取分组数据的第一行
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 1 And bytMode = 0 Then
            lngRow = lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1
            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <> 1 Then
                For lngRow = lngStartRow To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngRow, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngRow
                If lngRow >= VsfData.FixedRows Then lngStartRow = lngRow
            End If
        End If
        
        '获取数据的总行数
        lngRows = lngStartRow
        lngRowCount = Val(VsfData.TextMatrix(lngStartRow, mlngRowCount))
        If lngRowCount <= 0 Then lngRowCount = 1
        lngRows = lngRows + lngRowCount - 1
        
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 And bytMode = 0 Then
            For lngRow = lngStartRow + lngRowCount To VsfData.Rows - 1
                If Not VsfData.RowHidden(lngRow) Then
                    If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                        Exit For
                    End If
                    lngRows = lngRows + 1
                End If
            Next lngRow
        End If
        
        For lngRow = lngStartRow To lngRows
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strCols & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") > 0 And Trim(FormatValue(VsfData.TextMatrix(lngRow, lngCol))) <> "" Then
                    lngEditCol = lngCol
                    CheckCollectIsData = True
                    Exit Function
                End If
            Next lngCol
        Next lngRow
    End If
    CheckCollectIsData = False
End Function

Private Sub ImportAmount()
'从医嘱导入入量，导入的入量当分组数据处理
    Dim cbrControl As CommandBarControl
    Dim rsImpAmount As ADODB.Recordset
    Dim strDate As String, strValue As String
    Dim intCount As Integer, i As Integer, lngCurRow As Long
    Dim lngNameCol As Long, lngNameOrder As Long, strName As String '记录导入列的列号，项目序号、名称
    Dim lngCheckOrder As Long
    Dim arrColImCorrelative() As String, strImportOrders As String '可导入的项目序号及医嘱列名称
    
    On Error GoTo ErrHand
    If mstrColImCorrelative = "" Then Exit Sub
    arrColImCorrelative = Split(mstrColImCorrelative, "|")
    For intCount = 0 To UBound(arrColImCorrelative)
        strImportOrders = strImportOrders & ";" & Split(arrColImCorrelative(intCount), ",")(1) & "," & Split(arrColImCorrelative(intCount), ",")(2)
        '记录当前选择的导入列项目序号，取消恢复焦点使用
        If VsfData.COL - (cHideCols + VsfData.FixedCols - 1) = Split(arrColImCorrelative(intCount), ",")(0) Then
            lngCheckOrder = Split(arrColImCorrelative(intCount), ",")(1)
        End If
    Next
    If Left(strImportOrders, 1) = ";" Then strImportOrders = Mid(strImportOrders, 2)
    
    '返回记录集内容包含:key,名称,用量
    Set rsImpAmount = frmImportOrder.ShowMe(Me, Val(VsfData.TextMatrix(VsfData.ROW, c文件ID)), Val(VsfData.TextMatrix(VsfData.ROW, c病人ID)), Val(VsfData.TextMatrix(VsfData.ROW, c主页ID)), Val(VsfData.TextMatrix(VsfData.ROW, c婴儿)), strImportOrders, strDate)
    If rsImpAmount Is Nothing Then Call SetControlValue(lngCheckOrder, "", False): Exit Sub
    If rsImpAmount.RecordCount = 0 Then Call SetControlValue(lngCheckOrder, "", False): Exit Sub
    '导入入量
    If rsImpAmount.RecordCount > 0 Then rsImpAmount.MoveFirst
    For intCount = 1 To rsImpAmount.RecordCount
        If mblnShow = False Then
            VsfData.COL = Split(arrColImCorrelative(0), ",")(0) + (cHideCols + VsfData.FixedCols - 1)  '默认定位到第一个导入列
            Call VsfData_DblClick '追加后会取消编辑，此处需要重新设置
        End If
        lngCurRow = GetStartRow(VsfData.ROW)
        For i = 0 To UBound(arrColImCorrelative)
            lngNameCol = Split(arrColImCorrelative(i), ",")(0) + (cHideCols + VsfData.FixedCols - 1)
            lngNameOrder = Split(arrColImCorrelative(i), ",")(1)
            strName = Split(arrColImCorrelative(i), ",")(2)
            VsfData.COL = lngNameCol
            If mblnShow = False Then mintType = 1
            If SetControlValue(lngNameOrder, NVL(rsImpAmount(strName).Value)) = True Then '完成编辑控件赋值
                If intCount = rsImpAmount.RecordCount Then
                    Call MoveNextCell(True, True)
                Else
                    Call MoveNextCell(VsfData.COL < mlngNoEditor - 1, True)
                End If
                '完成医嘱信息赋值(入量列)
                If Record_Locate(mrsCellMap, "ID|" & mint页码 & "," & lngCurRow & "," & lngNameCol) = True Then
                    mrsCellMap.Fields("标记").Value = NVL(rsImpAmount("key").Value)
                    mrsCellMap.Update
                End If
            End If
        Next i
        If intCount < rsImpAmount.RecordCount Then
            '使用追加功能，追加一行(直接在当前行下面追加)
            Set cbrControl = cbsThis.FindControl(, conMenu_Edit_Group_Append)
            If Not cbrControl Is Nothing Then
                Call cbsThis_Execute(cbrControl)
            Else
                Exit For
            End If
        End If
        rsImpAmount.MoveNext
    Next intCount
    
    '隐蔽已显示的录入控件
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
        lstSelect(2).Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
        If mintType = 1 Then
            txtLst.Visible = False
            PicLst.Visible = False
        End If
    Case 4, 5
        picDouble.Visible = False
        lstSelect(2).Visible = False
    Case 6
        picMutilInput.Visible = False
        lstSelect(2).Visible = False
    Case 7
        picDoubleChoose.Visible = False
    End Select
    cmdWord.Visible = False
    mintType = -1
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SetControlValue(ByVal lngOrder As Long, ByVal strValue As String, Optional ByVal blnMode As Boolean = True) As Boolean
'功能：根据项目序号完成相应编辑控件的赋值(必须在编辑状态下)
'       blnMode:True 赋值,False 设置焦点
    Dim i As Integer, j As Integer
    Dim objControl As Object
    If mintType = -1 Then Exit Function
    On Error Resume Next
    If blnMode = True Then
        Select Case mintType
            Case 0
                txtInput.Text = strValue
            Case 1, 2
                If strValue <> "" Then
                    strValue = Replace(strValue, vbCrLf, "")
                    txtLst.Text = strValue
                    PicLst.Tag = "1"
                    j = lstSelect(mintType - 1).ListCount - 1
                    For i = 0 To j
                        '单选的第一个项目是清除选择，需要跳过此项，多选项目则直接进入
                        If Not (mintType = 1 And i = 0) Then
                            If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(mintType - 1).List(i), InStr(1, lstSelect(mintType - 1).List(i), "-") + 1) & ",") <> 0 Then
                                lstSelect(mintType - 1).Selected(i) = True
                                txtLst.Text = ""
                                PicLst.Tag = "0"
                            End If
                        End If
                    Next
                Else
                    txtLst.Text = ""
                    PicLst.Tag = "0"
                End If
            Case 3
                If strValue = "√" Or strValue = "" Then lblInput.Caption = strValue
            Case 4
                If lngOrder = Val(txtUpInput.Tag) Then
                    txtUpInput.Text = strValue
                Else
                    txtDnInput.Text = strValue
                End If
            Case 5
                If lngOrder = Val(lblUpInput.Tag) Then
                     If strValue = "√" Or strValue = "" Then lblUpInput.Caption = strValue
                Else
                     If strValue = "√" Or strValue = "" Then lblDnInput.Caption = strValue
                End If
            Case 6
                For i = 0 To txt.Count - 1
                    If lngOrder = Val(txt(i).Tag) Then
                        txt(i).Text = strValue
                    End If
                Next
            Case 7
                If lngOrder = Val(cboChoose(0).Tag) Then
                    j = 0
                Else
                    j = 1
                End If
                For i = 0 To cboChoose(j).ListCount - 1
                    If strValue = cboChoose(j).List(i) Then
                        cboChoose(j).ListIndex = i
                    End If
                Next
            Case Else
                SetControlValue = False
                Exit Function
        End Select
    Else
        Select Case mintType
            Case 0
                Set objControl = txtInput
            Case 1, 2
                Set objControl = lstSelect(mintType - 1)
            Case 3
                Set objControl = lblInput
            Case 4
                If lngOrder = Val(txtUpInput.Tag) Then
                    Set objControl = txtUpInput
                Else
                    Set objControl = txtDnInput
                End If
            Case 5
                If lngOrder = Val(lblUpInput.Tag) Then
                     Set objControl = lblUpInput
                Else
                     Set objControl = lblDnInput
                End If
            Case 6
                For i = 0 To txt.Count - 1
                    If lngOrder = Val(txt(i).Tag) Then
                        Set objControl = txt(i)
                    End If
                Next
            Case 7
                If lngOrder = Val(cboChoose(0).Tag) Then
                    j = 0
                Else
                    j = 1
                End If
                Set objControl = cboChoose(j)
            Case Else
                Set objControl = Nothing
        End Select
    End If
    If Not objControl Is Nothing Then
        If objControl.Visible And objControl.Enabled Then objControl.SetFocus
    End If
    SetControlValue = True
    If Err <> 0 Then Err.Clear
End Function

Private Function IsCanves(ByVal lngNO As Long) As Boolean
    Dim blnNULL As Boolean
    blnNULL = False
    mrsTemperItems.Filter = "项目序号 = " & lngNO & " and 记录法=1 "
    If mrsTemperItems.RecordCount > 0 Then blnNULL = True
    IsCanves = blnNULL
End Function

Private Sub LstChoose(ByVal strText As String)
    Dim i As Long
    If lstSelect(2).ListCount > 0 Then
        For i = 0 To lstSelect(2).ListCount
            If strText = lstSelect(2).List(i) And strText <> "" Then
                lstSelect(2).ListIndex = i
                Exit Sub
            End If
        Next
        lstSelect(2).ListIndex = 0
    Else
        lstSelect(2).ListIndex = 0
    End If

End Sub
