VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutAdviceEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "����ҽ���༭"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11130
   Icon            =   "frmOutAdviceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   1035
      Left            =   10845
      TabIndex        =   74
      Top             =   225
      Width           =   150
      _Version        =   589884
      _ExtentX        =   265
      _ExtentY        =   1826
      _StockProps     =   64
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   0
      ScaleHeight     =   8340
      ScaleWidth      =   11085
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   0
      Width           =   11085
      Begin VB.PictureBox pictmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9720
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   69
         Top             =   2160
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox pic���� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         ScaleHeight     =   270
         ScaleWidth      =   11070
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   5130
         Visible         =   0   'False
         Width           =   11070
         Begin VB.Label lbl���� 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   495
            TabIndex        =   61
            Top             =   45
            Width           =   1725
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   150
            Picture         =   "frmOutAdviceEdit.frx":058A
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Frame fra��� 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   0
         TabIndex        =   56
         Top             =   960
         Width           =   11060
         Begin VB.CommandButton cmdLastDiag 
            Caption         =   "�ϴ����"
            Height          =   300
            Left            =   50
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   320
            Width           =   900
         End
         Begin VB.OptionButton opt��� 
            Caption         =   "����ϱ�׼"
            Height          =   180
            Index           =   0
            Left            =   9840
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   100
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton opt��� 
            Caption         =   "����������"
            Height          =   180
            Index           =   1
            Left            =   9820
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   380
            Width           =   1200
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiag 
            Height          =   630
            Left            =   975
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   0
            Width           =   8760
            _cx             =   15452
            _cy             =   1111
            Appearance      =   2
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
            BackColorSel    =   13684944
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOutAdviceEdit.frx":0B14
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   115
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
            Editable        =   2
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
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
            Height          =   180
            Left            =   120
            TabIndex        =   33
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.Frame fraPati 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   0
         TabIndex        =   37
         Top             =   510
         Width           =   10995
         Begin VB.CommandButton cmdAlley 
            Caption         =   "����ʷ/����״̬"
            Height          =   350
            Left            =   9240
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.ComboBox cboӤ�� 
            Height          =   300
            ItemData        =   "frmOutAdviceEdit.frx":0CDD
            Left            =   9555
            List            =   "frmOutAdviceEdit.frx":0CF3
            Style           =   2  'Dropdown List
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   75
            Width           =   1395
         End
         Begin VB.Label lblӤ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ӥ��"
            Height          =   180
            Left            =   9135
            TabIndex        =   30
            Top             =   135
            Width           =   360
         End
         Begin VB.Label lblPati 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����: �Ա�: ����: �����: �ѱ�: ҽ�Ƹ��ʽ:"
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   210
            TabIndex        =   38
            Top             =   135
            Width           =   4050
         End
      End
      Begin MSComCtl2.MonthView dtpDate 
         Height          =   2220
         Left            =   1725
         TabIndex        =   1
         Top             =   2505
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   179830785
         TitleBackColor  =   -2147483636
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483637
         CurrentDate     =   37904
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   3645
         Left            =   60
         TabIndex        =   0
         Top             =   1650
         Width           =   10995
         _cx             =   19394
         _cy             =   6429
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
         BackColorSel    =   16444122
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   18
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOutAdviceEdit.frx":0D42
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
         WordWrap        =   -1  'True
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
         Begin MSComctlLib.ImageList img16 
            Left            =   1965
            Top             =   450
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":0E2A
                  Key             =   "��ʾ"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":13C4
                  Key             =   "���"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":195E
                  Key             =   "���_��ǰ"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":1EF8
                  Key             =   "���_����"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fraAdvice 
         Height          =   2685
         Left            =   45
         TabIndex        =   39
         Top             =   5340
         Width           =   11040
         Begin VB.ComboBox cboDruPur 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2220
            Width           =   1815
         End
         Begin VB.CommandButton cmdComExcReason 
            Height          =   300
            Left            =   10260
            Picture         =   "frmOutAdviceEdit.frx":2492
            Style           =   1  'Graphical
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "����ǰ˵������Ϊ����˵��"
            Top             =   1860
            Width           =   315
         End
         Begin VB.CommandButton cmdExcReason 
            Caption         =   "��"
            Height          =   265
            Left            =   9930
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1875
            Width           =   285
         End
         Begin VB.TextBox txt����˵�� 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Left            =   6255
            MaxLength       =   500
            TabIndex        =   18
            Top             =   1875
            Width           =   3945
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1200
            Visible         =   0   'False
            Width           =   4260
         End
         Begin VB.CommandButton cmdҽ������ 
            Caption         =   "��"
            Height          =   265
            Left            =   9960
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   573
            Width           =   285
         End
         Begin VB.TextBox cboҽ������ 
            Height          =   300
            Left            =   6255
            MaxLength       =   100
            TabIndex        =   24
            Top             =   555
            Width           =   4000
         End
         Begin VB.CommandButton cmd�ղ���ҩ���� 
            Height          =   300
            Left            =   10275
            Picture         =   "frmOutAdviceEdit.frx":2A1C
            Style           =   1  'Graphical
            TabIndex        =   64
            TabStop         =   0   'False
            ToolTipText     =   "����ǰ��������Ϊ�������ɡ�"
            Top             =   2250
            Width           =   315
         End
         Begin VB.CommandButton cmdReason 
            Caption         =   "��"
            Height          =   265
            Left            =   9930
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   2250
            Width           =   285
         End
         Begin VB.PictureBox picHelp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5200
            Picture         =   "frmOutAdviceEdit.frx":2FA6
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   62
            Top             =   900
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "����"
            Height          =   180
            Left            =   4560
            TabIndex        =   5
            Top             =   255
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.TextBox txt��ҩ���� 
            Height          =   300
            Left            =   4860
            MaxLength       =   1000
            TabIndex        =   21
            Top             =   2250
            Width           =   5385
         End
         Begin VB.CheckBox chkZeroBilling 
            Caption         =   "����Ϊ���շѵļ��ʵ�(&F)"
            Height          =   225
            Left            =   8280
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   233
            Width           =   2370
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   6255
            TabIndex        =   22
            Top             =   195
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CommandButton cmd����ʱ�� 
            Height          =   240
            Left            =   2490
            Picture         =   "frmOutAdviceEdit.frx":97F8
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "ѡ������(F4)"
            Top             =   1545
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txt����ʱ�� 
            Height          =   300
            Left            =   960
            TabIndex        =   9
            Top             =   1515
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmd�������� 
            Height          =   300
            Left            =   10275
            Picture         =   "frmOutAdviceEdit.frx":98EE
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "����ǰ��������Ϊ��������(Ctrl+D)"
            Top             =   555
            Width           =   315
         End
         Begin VB.ComboBox cbo����ִ�� 
            Height          =   300
            Left            =   6255
            TabIndex        =   29
            Text            =   "cbo����ִ��"
            Top             =   1515
            Width           =   1860
         End
         Begin VB.TextBox txt���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2385
            MaxLength       =   3
            TabIndex        =   16
            Top             =   1875
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.CommandButton cmdƵ�� 
            Height          =   240
            Left            =   4920
            Picture         =   "frmOutAdviceEdit.frx":9E78
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(F4)"
            Top             =   1545
            Width           =   270
         End
         Begin VB.TextBox txtƵ�� 
            Height          =   300
            Left            =   3540
            TabIndex        =   13
            Top             =   1515
            Width           =   1665
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3540
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1875
            Width           =   1365
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   960
            MaxLength       =   10
            TabIndex        =   15
            Top             =   1875
            Width           =   1530
         End
         Begin VB.CommandButton cmd�÷� 
            Height          =   240
            Left            =   2445
            Picture         =   "frmOutAdviceEdit.frx":9F6E
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(F4)"
            Top             =   1545
            Width           =   270
         End
         Begin VB.TextBox txt�÷� 
            Height          =   300
            Left            =   960
            TabIndex        =   11
            Top             =   1515
            Width           =   1815
         End
         Begin VB.CommandButton cmd��ʼʱ�� 
            Height          =   240
            Left            =   2460
            Picture         =   "frmOutAdviceEdit.frx":A064
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "ѡ������(F4)"
            Top             =   225
            Width           =   255
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "����ҽ��(&E)"
            Height          =   225
            Left            =   3000
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   233
            Width           =   1290
         End
         Begin VB.CommandButton cmdExt 
            Height          =   285
            Left            =   4920
            Picture         =   "frmOutAdviceEdit.frx":A15A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "�༭(F4)"
            Top             =   552
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   285
            Left            =   4920
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   870
            Width           =   285
         End
         Begin VB.ComboBox cboִ������ 
            Height          =   300
            ItemData        =   "frmOutAdviceEdit.frx":A250
            Left            =   9015
            List            =   "frmOutAdviceEdit.frx":A25D
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1200
            Width           =   1590
         End
         Begin VB.ComboBox cboִ�п��� 
            Height          =   300
            Left            =   6255
            TabIndex        =   27
            Text            =   "cboִ�п���"
            Top             =   1200
            Width           =   1860
         End
         Begin VB.TextBox txtҽ������ 
            Height          =   900
            Left            =   960
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   6
            ToolTipText     =   "�� ~ ����ʾ���׷���ѡ����,Ctrl+F1����ҽ����Ϣ"
            Top             =   552
            Width           =   3945
         End
         Begin VB.TextBox txt��ʼʱ�� 
            Height          =   300
            Left            =   960
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   195
            Width           =   1815
         End
         Begin VB.ComboBox cboִ��ʱ�� 
            Height          =   300
            Left            =   6255
            TabIndex        =   26
            Top             =   877
            Width           =   4350
         End
         Begin VB.Label lbl����˵�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����˵��"
            Height          =   180
            Left            =   5490
            TabIndex        =   20
            Top             =   1935
            Width           =   720
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���㷽ʽ"
            Height          =   180
            Left            =   180
            TabIndex        =   68
            Top             =   1260
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblҽ������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ������"
            Height          =   180
            Left            =   180
            TabIndex        =   67
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lbl��ҩĿ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ҩĿ��"
            Height          =   180
            Left            =   165
            TabIndex        =   59
            Top             =   2265
            Width           =   720
         End
         Begin VB.Label lbl��ҩ���� 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ����"
            Height          =   180
            Left            =   4080
            TabIndex        =   58
            Top             =   2280
            Width           =   720
         End
         Begin VB.Label lbl���ٵ�λ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��/����"
            Height          =   180
            Left            =   7335
            TabIndex        =   55
            Top             =   255
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   5670
            TabIndex        =   54
            Top             =   255
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lbl����ʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   165
            TabIndex        =   53
            Top             =   1575
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl����ִ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ִ��"
            Height          =   180
            Left            =   5490
            TabIndex        =   52
            Top             =   1575
            Width           =   720
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ��"
            Height          =   180
            Left            =   2205
            TabIndex        =   51
            Top             =   1935
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblƵ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ƶ��"
            Height          =   180
            Left            =   3135
            TabIndex        =   46
            Top             =   1575
            Width           =   360
         End
         Begin VB.Label lbl������λ 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "��λ"
            Height          =   180
            Left            =   4905
            TabIndex        =   42
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2955
            TabIndex        =   41
            ToolTipText     =   "���ü�������(Ctrl+~)"
            Top             =   1935
            Width           =   540
         End
         Begin VB.Label lbl������λ 
            BackStyle       =   0  'Transparent
            Caption         =   "��λ"
            Height          =   180
            Left            =   2505
            TabIndex        =   44
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   540
            TabIndex        =   43
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lblҽ������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ������"
            Height          =   180
            Left            =   5490
            TabIndex        =   50
            Top             =   615
            Width           =   720
         End
         Begin VB.Label lblִ������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ������"
            Height          =   180
            Left            =   8250
            TabIndex        =   49
            Top             =   1260
            Width           =   720
         End
         Begin VB.Label lblִ�п��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ�п���"
            Height          =   180
            Left            =   5490
            TabIndex        =   48
            Top             =   1260
            Width           =   720
         End
         Begin VB.Label lbl�÷� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�÷�"
            Height          =   180
            Left            =   540
            TabIndex        =   45
            Top             =   1575
            Width           =   360
         End
         Begin VB.Label lbl��ʼʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼʱ��"
            Height          =   180
            Left            =   180
            TabIndex        =   40
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblִ��ʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ��ʱ��"
            Height          =   180
            Left            =   5490
            TabIndex        =   47
            Top             =   937
            Width           =   720
         End
      End
      Begin MSComctlLib.StatusBar stbThis 
         Height          =   360
         Left            =   0
         TabIndex        =   57
         Top             =   8025
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   9
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Bevel           =   2
               Object.Width           =   2355
               MinWidth        =   882
               Picture         =   "frmOutAdviceEdit.frx":A27F
               Text            =   "�������"
               TextSave        =   "�������"
               Key             =   "ZLFLAG"
               Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   12912
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   318
               MinWidth        =   2
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   318
               MinWidth        =   2
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   318
               MinWidth        =   2
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   970
               MinWidth        =   970
               Picture         =   "frmOutAdviceEdit.frx":AB13
               Key             =   "KB"
            EndProperty
            BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   617
               MinWidth        =   617
               Picture         =   "frmOutAdviceEdit.frx":B879
               Key             =   "PY"
               Object.ToolTipText     =   "ƴ��(F7)"
            EndProperty
            BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   617
               MinWidth        =   617
               Picture         =   "frmOutAdviceEdit.frx":BEB3
               Key             =   "WB"
               Object.ToolTipText     =   "���(F7)"
            EndProperty
            BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
               Bevel           =   0
               Object.Width           =   953
               MinWidth        =   25
               Text            =   "�Ƽ�"
               TextSave        =   "�Ƽ�"
               Key             =   "Price"
               Object.ToolTipText     =   "��ʾ���ƼƼ����(F8)"
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
      Begin VB.Image imgButtonDel 
         Height          =   240
         Left            =   1680
         Picture         =   "frmOutAdviceEdit.frx":C4ED
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgButtonNew 
         Height          =   240
         Left            =   1005
         Picture         =   "frmOutAdviceEdit.frx":12D3F
         Top             =   105
         Visible         =   0   'False
         Width           =   240
      End
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   75
         Top             =   75
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
End
Attribute VB_Name = "frmOutAdviceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event FormUnload(Cancel As Integer)
Public Event EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean) '�༭�������
Public Event CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���Id As String, ByRef blnNo As Boolean) '������ϼ���Ƿ���д��Ⱦ�����濨

Public mblnOK As Boolean
'��ڲ���
Private mint���� As Integer '���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
Private mblnModal As Boolean
Private mfrmParent As Object
Private mMainPrivs As String
Private mlng����ID As Long
Private mstr�Һŵ� As String '���˹Һŵ��ݺ�
Private mlng�Һ�ID As Long
Private mlng��ͬ��λID As Long
Private mlngǰ��ID As Long 'ҽ������վ��ҽ��ʱ��
Private mstrǰ��IDs As String
Private mlngҽ������ID As Long
Private mintӤ�� As Integer '�޸�ʱ��
Private mlngҽ��ID As Long '�޸�ʱ��
Private mblnCancle As Boolean   '����ʱ��֤

'�������
Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As ADODB.Recordset 'ҽ�����ݶ���
Private mrsDrugScale As ADODB.Recordset '���ü�������
Private mrsPrice As ADODB.Recordset 'ҽ����Ӧ���շ���Ŀ��Ϣ��

Private WithEvents mfrmSend As frmOutAdviceSend
Attribute mfrmSend.VB_VarHelpID = -1
Private WithEvents mfrmShortCut As frmClinicShortCut
Attribute mfrmShortCut.VB_VarHelpID = -1
Private WithEvents mfrmPrice As frmAdvicePrice
Attribute mfrmPrice.VB_VarHelpID = -1
Private mobjKeyBoard As Object '��Ļ���̶���̬����
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀⷿ�ĳ����鷽ʽ
Private mstrDelIDs As String '��¼��Ҫ��ɾ����ҽ��ID
Private mstrAduitDelIDs As String '������������ļ�ɾ��
Private mstr�Ա� As String '������Ŀ���������ж�
Private mint���� As Integer '���˵���������
Private mDat�������� As Date '���˳�������
Private mdbl����� As Double   '��������� PASS =3-̫Ԫͨ
Private mstr���� As String
Private mstr���֤�� As String
Private mstr�ѱ� As String
Private mdat�Һ�ʱ�� As Date '��������ж�
Private mlng���˿���id As Long '����(�Һ�)����ID
Private mint���� As Integer '��ǰ��������
Private mstr������ As String '��ǰ����ҽ�Ƹ��ʽ����
Private mbln��ҽ As Boolean
Private mblnReturn As Boolean
Private mblnIsInHelp As Boolean
Private mrs��� As ADODB.Recordset
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln�������� As Boolean
Private msngPre���� As Single
Private mbln���� As Boolean  '��ǰ�����Ƿ���
Private mstr�Զ������뵥IDs As String 'ID1,����1|ID2,����2������

Private mbln��չҳǩ As Boolean '�Ƿ�֧����չҳǩ
Private mblnNoRefresh As Boolean '��ˢ�½���
Private mcolSubForm As Collection '��չҳǩ���󼯺�
Private mlngID���� As Long '��ҽ��ID���ø�����ʾ
Private mstrDel��Ѫ As String '��¼��Ҫ��ɾ������Ѫҽ��ID������¼��ҽ��ID
Private mlngΣ��ֵID As Long

'���ڲ���
Private mobjPassMap As Object  ' PASS  ���ڶ���ӳ��
Private mblnPass As Boolean
'���ز���
Private mint���� As Integer
Private mstrLike As String
Private mbln�Զ�Ƥ�� As Boolean
Private mbln���� As Boolean
Private msng���� As Single
Private mbln���� As Boolean
Private mblnStaKB As Boolean '�Ƿ��Զ�������Ļ����
Private mbln���Ѷ��� As Boolean
Private mblnAutoClose As Boolean
Private mblnAddAgent As Boolean '�Ƿ�Ҫ��ǼǶ���ҩƷ��������Ϣ
Private mblnNewLIS As Boolean
Private mstrPurMed As String  '����ҩ��ȱʡ��ҩĿ�� "1"-Ԥ����"2"-���ƣ�"0"-�´�ʱȷ��
Private mbytSize As Byte
Private mrsDiag  As ADODB.Recordset '��ҽ��ϼ�¼��
Private mblnFreeInput As Boolean

'�¼�״̬���Ʊ���
Private mblnRowMerge As Boolean
Private mblnNoSave As Boolean
Private mblnRunFirst As Boolean
Private mblnRowChange As Boolean
Private mblnDoCheck As Boolean
Private mbytPatiType As Byte  '�������ͣ�1 ��ͨ��2 ����
'�����Ƴ�����������3�죬��ͨ7��
Private Const conEmergency = 3
Private Const conOrdinary = 7

'����������
Private Const conMenu_New = 100
Private Const conMenu_Insert = 101
Private Const conMenu_Delete = 102
Private Const conMenu_Merge = 104
Private Const conMenu_Copy = 105
Private Const conMenu_Scheme = 106
Private Const conMenu_Save = 107
Private Const conMenu_Sign = 108
Private Const conMenu_Reference = 109
Private Const conMenu_Help = 110
Private Const conMenu_Exit = 111
Private Const conMenu_Agent = 112
Private Const conMenu_Send = 205
Private Const conMenu_DrugScale = 300
Private Const conMenu_AdvicePay = 3006

'ִ��ʱ��ʾ��
Private Const COL_����ִ�� = _
    "ÿ������ 1/8-3/8-5/8 �� 1/8:00-3/8:00-5/8:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ������һ��8:00,��������8:00,�������8:00�⼸��ʱ��ִ��"
Private Const COL_����ִ�� = _
    "ÿ������ 8-12-16 �� 8:00-12:00-16:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ��8:00,12:00,16:00�⼸��ʱ��ִ��" & vbCrLf & _
    "����һ�� 1/8 �� 1/8:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ�����еĵ�1��8:00���ʱ��ִ��"
Private Const COL_��ʱִ�� = _
    "ÿСʱ���� 1:20-1:40" & vbCrLf & _
        vbTab & "��ʾ��ÿСʱ�ڵ�20��40����������ʱ��ִ��" & vbCrLf & _
    "��Сʱһ�� 2:30 �� 1:30 �� 1:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ��Сʱ�ڵĵ�2�ĸ�Сʱ��30�������ʱ��ִ��" & vbCrLf & _
        vbTab & "������ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ��30�������ʱ��ִ��" & vbCrLf & _
        vbTab & "������ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ���ʱ��ִ��"

'�̶���
Private Const COL_F��־ = 0 '����¼�룬��������¼,������ҩ���״̬
'�ɼ�������
Private Const COL_��ʾ = 1 'Pass:���ַ������ʹ���,�ձ�ʾû�������
Private Const COL_��� = 2
Private Const COL_��ʼʱ�� = 3
Private Const col_ҽ������ = 4
Private Const COL_���� = 5
Private Const COL_������λ = 6
Private Const COL_���� = 7
Private Const COL_������λ = 8
Private Const COL_���� = 9
Private Const COL_Ƶ�� = 10
Private Const COL_�÷� = 11
Private Const COL_ҽ������ = 12 'Data���ڴ��ժҪ(ҽ��)
Private Const COL_ִ��ʱ�� = 13
Private Const COL_����ҽ�� = 14
Private Const COL_����˵�� = 15  'ҩƷ����˵��
Private Const COL_����ҩ�� = 16

'����������
Private Const COL_EDIT = 17 '�༭��־��0-ԭʼ��,1-������,2-�޸�������,3-�޸������,����Dataֵ=���µĳ��׷���ID
Private Const COL_���ID = COL_EDIT + 1
Private Const COL_Ӥ�� = COL_EDIT + 2
Private Const COL_��� = COL_EDIT + 3  'Pass:Dataֵ���ڼ�¼�Ƿ��������˽��
Private Const COL_״̬ = COL_EDIT + 4  '����ҽ����¼.״̬��Dataֵ��¼��ҽ���Ƿ��ѽ�����ҽ���ܿؼ��
Private Const COL_��� = COL_EDIT + 5  '����û������¼��ҽ��(*)��AdviceSet������Ŀ��Nvl(���,'*')ֻ��Ϊ��Ԥ����һ
Private Const COL_������ĿID = COL_EDIT + 6
Private Const COL_���� = COL_EDIT + 7
Private Const COL_�걾��λ = COL_EDIT + 8
    Private Const COL_����ʱ�� = COL_EDIT + 8
    Private Const COL_��Ѫʱ�� = COL_EDIT + 8
Private Const COL_��鷽�� = COL_EDIT + 9 '�� ���=K ��Ѫҽ��ʱ���ֶ��������"1"��ʾ��Ѫҽ��
    Private Const COL_��ҩ��̬ = COL_EDIT + 9 '0=ɢװ��1=��ҩ��Ƭ��2=����
Private Const COL_ִ�б�� = COL_EDIT + 10
Private Const COL_�շ�ϸĿID = COL_EDIT + 11
Private Const COL_Ƶ�ʴ��� = COL_EDIT + 12
Private Const COL_Ƶ�ʼ�� = COL_EDIT + 13
Private Const COL_�����λ = COL_EDIT + 14
Private Const COL_�Ƽ����� = COL_EDIT + 15
Private Const COL_ִ�п���ID = COL_EDIT + 16
Private Const COL_ִ������ = COL_EDIT + 17 '����ҽ����¼.ִ������=������ĿĿ¼.ִ�п���
Private Const COL_��������ID = COL_EDIT + 18
Private Const COL_����ʱ�� = COL_EDIT + 19
Private Const COL_��־ = COL_EDIT + 20     '0-��ͨ,1-������2-��¼

Private Const COL_���㷽ʽ = COL_��־ + 1 '������ĿĿ¼.���㷽ʽ
Private Const COL_Ƶ������ = COL_��־ + 2 '������ĿĿ¼.ִ��Ƶ��
Private Const COL_�������� = COL_��־ + 3 '������ĿĿ¼.��������
Private Const COL_ִ�з��� = COL_��־ + 4 '������ĿĿ¼.ִ�з���
Private Const COL_��� = COL_��־ + 5 '�������װ��ŵĿ��ÿ��
Private Const COL_�ɷ���� = COL_��־ + 6 '�������ڴ���Ƿ��������
    Private Const COL_�������� = COL_��־ + 6
Private Const COL_����ϵ�� = COL_��־ + 7
Private Const COL_���ﵥλ = COL_��־ + 8
Private Const COL_�����װ = COL_��־ + 9
Private Const COL_�������� = COL_��־ + 10 '��ҩ������ĿΪ¼������
Private Const COL_����ְ�� = COL_��־ + 11
Private Const COL_������� = COL_��־ + 12
Private Const COL_ҩƷ���� = COL_��־ + 13
Private Const COL_���� = COL_��־ + 14
Private Const COL_ǩ���� = COL_��־ + 15
Private Const COL_���� = COL_��־ + 16
Private Const COL_��Ѽ��� = COL_��־ + 17
Private Const COL_�����ȼ� = COL_��־ + 18 '����ҩ��ȼ�:0-�ǿ���ҩ,1-�����Ƽ�,2-���Ƽ�,3-����ʹ�ü�
Private Const COL_��ҩĿ�� = COL_��־ + 19
Private Const COL_��ҩ���� = COL_��־ + 20
Private Const COL_���״̬ = COL_��־ + 21 '����ҩ�����״̬��Null-������ˣ�1-����ˣ�2-���ͨ����3-���δͨ��
Private Const COL_���� = COL_��־ + 22    '�Ƿ�����   1-���ԣ�0-������
Private Const COL_�Ƿ��� = COL_��־ + 23   'ҩƷ�Ƿ���
Private Const COL_�Ƿ��� = COL_��־ + 24
Private Const COL_�䷽ID = COL_��־ + 25
Private Const COL_�ٴ��Թ�ҩ = COL_��־ + 26
Private Const COL_��ΣҩƷ = COL_��־ + 27
Private Const COL_�����ĿID = COL_��־ + 28
Private Const COL_�Ƿ�ͣ�� = COL_��־ + 29 '=1��ʶ��ͣ�ã�=0��NULL��ʶδͣ��
Private Const COL_�Ƿ���ý = COL_��־ + 30 '=1��ý
Private Const COL_������� = COL_��־ + 31 '
Private Const COL_����Ӧ�� = COL_��־ + 32
Private Const COL_�������״̬ = COL_��־ + 33
Private Const COL_��������� = COL_��־ + 34
Private Const COL_������� = COL_��־ + 35    '������ҩ�����
Private Const COL_Σ��ֵID = COL_��־ + 36

Private Const M_LNG_DIAGCOUNT = 10 '��������

Private Type AGENT_INFO
    ����������      As String
    ���������֤��  As String
    ���ξ�����¼��  As Boolean
End Type
Private AgentInfo As AGENT_INFO

Private Enum COL_ENUM_���
    col��־ = 0
    col��ҽ = 1
    COL��ҽ = 2
    col���� = 3
    col��� = 4
    col��ҽ֤�� = 5
    col����ʱ�� = 6
    col���� = 7
    col���� = 8
    col���ID = 9
    col����ID = 10
    col֤��ID = 11
    colICD�� = 12
    colҽ��ID = 13
    COLDEL = 14
    col��ϱ��� = 15
    col�������� = 16
    col������� = 17
    col�������� = 18
    col֤����� = 19
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal int���� As Integer, ByVal MainPrivs As String, ByVal lng����ID As Long, ByVal str�Һŵ� As String, _
    Optional ByVal lngǰ��ID As Long, Optional ByVal intӤ�� As Integer, Optional ByVal lngҽ��ID As Long, Optional ByVal blnModal As Boolean, _
     Optional ByVal lng�������ID As Long, Optional ByVal strǰ��IDs As String, Optional ByRef objMip As Object, Optional ByVal lngΣ��ֵID As Long) As Boolean
    
    Set mfrmParent = frmParent
    mint���� = int����
    mblnModal = blnModal
    mMainPrivs = MainPrivs
    mlng����ID = lng����ID
    mstr�Һŵ� = str�Һŵ�
    mlngǰ��ID = lngǰ��ID
    mstrǰ��IDs = strǰ��IDs
    mlngҽ������ID = IIF(mlngǰ��ID <> 0, lng�������ID, 0)
    mintӤ�� = intӤ��
    mlngҽ��ID = lngҽ��ID
    mlngΣ��ֵID = lngΣ��ֵID
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Me.Show IIF(blnModal, 1, 0), frmParent
    ShowMe = mblnOK
End Function

Private Sub InitCommandBar()
'���ܣ���ʼ��������
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objMenu As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    Dim i As Long
    
    Dim blnTwo As Boolean, lng��ͬ��λID As Long, bln��λ���� As Boolean
    Dim strInsidePrivs As String
    
    strInsidePrivs = GetInsidePrivs(p����ҽ���´�)
    lng��ͬ��λID = Get��ͬ��λID
    bln��λ���� = Val(zlDatabase.GetPara("��λ����", glngSys, p����ҽ���´�)) <> 0
    
    blnTwo = Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) <> 2 Or _
             Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) = 2 And _
             (InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ���ʵ�") = 0 Or _
            InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ�շѵ�") = 0 Or _
            bln��λ���� And lng��ͬ��λID = 0)
            
    If bln��λ���� And lng��ͬ��λID = 0 Or _
        InStr(GetInsidePrivs(p����ҽ���´�), "��Ѽ���") = 0 Or _
        InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ���ʵ�") = 0 Or _
        Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) = 0 Then
        '�Ƿ���ʾ��Ѽ���
        chkZeroBilling.Visible = False
    End If
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
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
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = frmIcons.imgMain.Icons
    
    '�˵�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҩ�����", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        '�����չ����
        Call CreatePlugInOK(p����ҽ���´�, mint����)
        If Not gobjPlugIn Is Nothing Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����")
            objPopup.BeginGroup = True
        End If
                If Not gobjDrugExplain Is Nothing Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewDrugExplain, "�鿴ҩƷ˵����")
            objControl.IconId = 3205
        End If
    End With
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_New, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Insert, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "ɾ��")
        
        If mint���� = 0 Then 'ֻ������ҽ������վ����ʱ�����⼸����ť
            intTmp = Val(Mid(gstrOutUseApp, 1, 1))
            If intTmp = 1 Then strTmp = strTmp & ",�������:" & conMenu_Edit_PacsApply
            intTmp = Val(Mid(gstrOutUseApp, 2, 1))
            If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",��������:" & conMenu_Edit_LISApply
            intTmp = Val(Mid(gstrOutUseApp, 3, 1))
            If intTmp = 1 Then strTmp = strTmp & ",��Ѫ����:" & conMenu_Edit_BloodApply
            

                        Get�Զ������뵥 1, mstr�Զ������뵥IDs
            If mstr�Զ������뵥IDs <> "" Then
                For i = 0 To UBound(Split(mstr�Զ������뵥IDs, "|"))
                    strTmp = strTmp & "," & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(0)
                Next
            End If
            strTmp = Mid(strTmp, 2)
            If strTmp <> "" Then
                If InStr(strTmp, ",") = 0 Then
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    Set objControl = .Add(xtpControlButton, lngID, strName)
                        objControl.IconId = conMenu_Edit_PacsApply
                        objControl.ToolTipText = strName
                        objControl.BeginGroup = True
                                            If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Else
                    varArr = Split(strTmp, ",")
                    For i = 0 To UBound(varArr)
                        strTmp = varArr(i)
                        strName = Split(strTmp, ":")(0)
                        lngID = Val(Split(strTmp, ":")(1))
                        
                        If i = 0 Then
                            Set objPopup = .Add(xtpControlSplitButtonPopup, lngID, strName)
                                objPopup.IconId = conMenu_Edit_PacsApply
                                objPopup.BeginGroup = True
                                objPopup.ToolTipText = strName
                                With objPopup.CommandBar.Controls
                                    Set objControl = .Add(xtpControlButton, lngID * 10# + 1, strName)
                                End With
                        Else
                            Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                        End If
                                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    Next
                End If
            End If
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Merge, "һ����ҩ"): objControl.BeginGroup = True
        If InStr(strInsidePrivs, ";����ҽ��;") > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Copy, "����ҽ��")
        End If
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Scheme, "���׷���")
        objPopup.ToolTipText = "����Ϊ���׷���"
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Scheme * 10# + 1, "����Ϊ���׷���"
            .Add xtpControlButton, conMenu_Scheme * 10# + 2, "��ʾ���׷���ѡ����"
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_Agent, "������")
        Set objControl = .Add(xtpControlButton, conMenu_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Sign, "ǩ��")
        
        If blnTwo Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Send, "����")
            objPopup.ToolTipText = "�Զ���ɷ���(F3)"
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Send * 10# + 1, "�Զ���ɷ���"
                .Add xtpControlButton, conMenu_Send * 10# + 2, "ҽ�����ʹ���"
            End With
        Else
            Set objControl = .Add(xtpControlButton, conMenu_Send, "����")
        End If
        If InStr(GetInsidePrivs(p����ҽ���´�), ";����޿�֧��;") > 0 And Not mblnAutoClose Then
            Set objControl = .Add(xtpControlButton, conMenu_AdvicePay, "���֧��"): objControl.BeginGroup = True
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Reference, "�ο�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.ID = conMenu_Help Or objControl.ID = conMenu_Exit Or objControl.ID = conMenu_Reference Then
            objControl.Style = xtpButtonIcon
        Else
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
   
    
    '�ȼ���:ע�ⲻ�ܺ�ϵͳ���ı��༭�ȼ���ͻ���Լ�Form_keydown�еĳ�ͻ
    With cbsMain.KeyBindings
        .Add 0, vbKeyF2, conMenu_Save
        If blnTwo Then
            .Add 0, vbKeyF3, conMenu_Send * 10# + 1
            .Add FCONTROL, vbKeyG, conMenu_Send * 10# + 2
        Else
            .Add 0, vbKeyF3, conMenu_Send
        End If
        .Add 0, vbKeyF6, conMenu_Reference
        .Add 0, vbKeyF9, conMenu_Merge
        .Add 0, vbKeyF1, conMenu_Help
        .Add FCONTROL, vbKeyA, conMenu_New
        .Add FCONTROL, vbKeyI, conMenu_Insert
        .Add FCONTROL, vbKeyK, conMenu_Merge
        .Add FCONTROL, vbKeyY, conMenu_Copy
        .Add FCONTROL, vbKeyT, conMenu_Scheme
        .Add FCONTROL, vbKeyS, conMenu_Save
        .Add FALT, vbKeyX, conMenu_Exit
    End With
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��������ݣ����ڴ�����Ի����ûָ�֮ǰ
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant

    strHead = _
        ",240,4;,270,4;��ʼʱ��,1530,1;ҽ������,3500,1;����,600,7;��λ,450,1;����,600,7;��λ,450,1;" & _
        "����,450,1;Ƶ��,1200,1;�÷�,1200,1;ҽ������,1000,1;ִ��ʱ��;����ҽ��,850,1;����˵��,1000,1;����ҩ��,850,1;" & _
        "EDIT;���ID;Ӥ��;���;ҽ��״̬;�������;������ĿID;����;�걾��λ;��鷽��;ִ�б��;�շ�ϸĿID;" & _
        "Ƶ�ʴ���;Ƶ�ʼ��;�����λ;�Ƽ�����;ִ�п���ID;ִ������;��������ID;����ʱ��;��־;���㷽ʽ;" & _
        "Ƶ������;��������;ִ�з���;���;�ɷ����;����ϵ��;���ﵥλ;�����װ;��������;����ְ��;" & _
        "�������;ҩƷ����;����;ǩ����;����;��Ѽ���;�����ȼ�;��ҩĿ��;��ҩ����;���״̬;����;" & _
        "�Ƿ���;�Ƿ���;�䷽ID;�ٴ��Թ�ҩ;��ΣҩƷ;�����ĿID;�Ƿ�ͣ��;�Ƿ���ý;�������;����Ӧ��;�������״̬;���������;�������;Σ��ֵID"
        
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColHidden(COL_��ʾ) = True 'Pass
        '.FrozenCols = COL_ҽ������ + 1 - .FixedCols
        .ColWidth(0) = 14 * Screen.TwipsPerPixelX
        
        '��ͷͼ��
        Set .Cell(flexcpPicture, 0, COL_��ʾ) = img16.ListImages("��ʾ").Picture
        Set .Cell(flexcpPicture, 0, COL_���) = img16.ListImages("���").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 4
    End With
End Sub

Private Sub Set�÷�Input(rsInput As ADODB.Recordset, ByVal int���� As Integer)
'���ܣ������ҩ;������ҩ�÷������
'������rsInput=�����ѡ��ķ��ؼ�¼
'      int����=2-��ҩ;��,4-��ҩ�÷�
'˵���������ѡƵ��,����ϸ�ҩ;���������ִ��ʱ�䷽���ı仯
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnValid As Boolean, sng���� As Single
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim vMsg As VbMsgBoxResult, strMsg As String
    Dim bln�÷����� As Boolean, bln�÷�Ƶ�� As Boolean
    Dim strIDs1 As String, strIDs2 As String, strҽ������ As String
    
    On Error GoTo errH
    cmd�÷�.Tag = rsInput!ID
    txt�÷�.Text = rsInput!����
    txt�÷�.Tag = "1"
    
    With vsAdvice
        '����Һ���������
        If int���� = 2 Then
            If Nvl(rsInput!ִ�з���ID, 0) <> 1 And cbo����.Text <> "" Then
                cbo����.Text = ""
                cbo����.Tag = "1"
            End If
        End If
        
        '���»�ȡ���õ�ȱʡʱ�䷽��
        If cboִ��ʱ��.Enabled Then '"��ѡƵ��"��ҩƷʱ
            Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(.Row), .TextMatrix(.Row, COL_Ƶ��), rsInput!ID)
            If cboִ��ʱ��.ListCount > 0 Then
                cboִ��ʱ��.ListIndex = 0
                cboִ��ʱ��.Tag = "1"
            Else
                '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                If cboִ��ʱ��.Text <> "" Then
                    blnValid = ExeTimeValid(cboִ��ʱ��.Text, Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), .TextMatrix(.Row, COL_�����λ))
                    If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                        cboִ��ʱ��.Text = ""
                        cboִ��ʱ��.Tag = "1"
                    End If
                End If
            End If
        End If
        
        '���������÷�������ȱʡ����
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            strSQL = "Select Ƶ��,С������,���˼���,ҽ������,�Ƴ�" & _
                " From �����÷����� Where ����>0 And ��ĿID=[1] And �÷�ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_������ĿID)), Val(rsInput!ID))
            If rsTmp.EOF Then
                'ҩƷû�������÷����������Ҹ�ҩ;����ȱʡƵ�ʣ����ֻ������һ������Ƶ�ʵĻ���,֮ǰ����ҽ����Ŀʱ���õ�ȱʡƵ���ǰ���������ȡ�ĵ�һ��
                strSQL = "Select Ƶ�� From �����÷����� Where ����>0 And ��ĿID=[1] Order by Ƶ��"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!ID))
                bln�÷�Ƶ�� = rsTmp.RecordCount > 0
            Else
                bln�÷����� = True
            End If
            
            If bln�÷����� Or bln�÷�Ƶ�� Then
                If Not IsNull(rsTmp!Ƶ��) Then
                    Call GetƵ����Ϣ_����(rsTmp!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                    txtƵ��.Text = strƵ��
                    cmdƵ��.Tag = strƵ��
                    txtƵ��.Tag = "1"
                End If
                
                '�����µ�Ƶ����������ִ��ʱ��
                If cboִ��ʱ��.Enabled Then
                    Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(.Row), strƵ��, rsInput!ID)
                    If cboִ��ʱ��.ListCount > 0 Then
                        cboִ��ʱ��.ListIndex = 0
                        cboִ��ʱ��.Tag = "1"
                    Else
                        '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                        If cboִ��ʱ��.Text <> "" Then
                            blnValid = ExeTimeValid(cboִ��ʱ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                            If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                                cboִ��ʱ��.Text = ""
                                cboִ��ʱ��.Tag = "1"
                            End If
                        End If
                    End If
                End If
                
                If bln�÷����� Then
                    'ҩƷ����
                    If mint���� > 12 Then
                        If Nvl(rsTmp!���˼���, 0) <> 0 Then
                            txt����.Text = FormatEx(rsTmp!���˼���, 5)
                            txt����.Tag = "1"
                        End If
                    Else
                        If Nvl(rsTmp!С������, 0) <> 0 Then
                            txt����.Text = FormatEx(rsTmp!С������, 5)
                            txt����.Tag = "1"
                        ElseIf Nvl(rsTmp!���˼���, 0) <> 0 Then
                            txt����.Text = FormatEx(rsTmp!���˼��� * (mint���� + 2) * 5 / 100, 5)
                            txt����.Tag = "1"
                        End If
                    End If
                    
                    'ȡȱʡ������
                    sng���� = msng����
                    If mbln���� Then
                        If str�����λ = "��" Then
                            sng���� = IIF(7 > sng����, 7, sng����)
                        ElseIf str�����λ = "��" Then
                            sng���� = IIF(intƵ�ʼ�� > sng����, intƵ�ʼ��, sng����)
                        ElseIf str�����λ = "Сʱ" Then
                            sng���� = IIF(intƵ�ʼ�� \ 24 > sng����, intƵ�ʼ�� \ 24, sng����)
                        ElseIf str�����λ = "����" Then
                            If sng���� = 0 Then sng���� = 1
                        End If
                        If sng���� = 0 Then sng���� = 1
                    End If
                    If Nvl(rsTmp!�Ƴ�, 1) > sng���� Then
                        sng���� = Nvl(rsTmp!�Ƴ�, 1)
                    End If
                    If Val(.TextMatrix(.Row, COL_����)) > sng���� Then
                        sng���� = Val(.TextMatrix(.Row, COL_����))
                    End If
                    If Val(.TextMatrix(.Row, COL_����)) <> sng���� Then
                        txt����.Text = sng����
                        txt����.Tag = "1"
                    End If
                    
                    'ҩƷ��������:�����װ
                    If strƵ�� <> "" And Val(txt����.Text) <> 0 _
                        And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                        And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                        
                        txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                            Val(txt����.Text), sng����, _
                            intƵ�ʴ���, intƵ�ʼ��, str�����λ, _
                            .TextMatrix(.Row, COL_ִ��ʱ��), _
                            Val(.TextMatrix(.Row, COL_����ϵ��)), _
                            Val(.TextMatrix(.Row, COL_�����װ)), _
                            Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                        If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") = 0 Then
                            txt����.Text = IntEx(Val(txt����.Text))
                        ElseIf Val(.TextMatrix(.Row, COL_�ɷ����)) <> 0 Then
                            txt����.Text = IntEx(Val(txt����.Text))
                        End If
                        txt����.Tag = "1"
                    End If
                    
                    'ҽ������
                    If Not IsNull(rsTmp!ҽ������) Then
                        cboҽ������.Text = rsTmp!ҽ������
                        cboҽ������.Tag = "1"
                    End If
                End If
            End If
        End If
    End With
    
    '����ǰҽ����ҩ;��/�巨�ı仯
    Call AdviceChange
    
    '�Ա��ն�����м��
    Call GetInsureStr(strIDs1, strIDs2, strҽ������, vsAdvice.Row)
    strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, 1, strIDs1, strIDs2, strҽ������)
    If strMsg <> "" Then
        If gintҽ������ = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln���Ѷ��� = False
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetƵ��Input(rsInput As ADODB.Recordset, ByVal int��Χ As Integer)
'���ܣ�����ִ��Ƶ�ʺ����
'������rsInput=�����ѡ��ķ��ؼ�¼
'      int��Χ=1-��ҽ;2-��ҽ;-1-һ����;-2-������
'˵��������÷��������ִ��ʱ�䷽���ı仯
    Dim lng�÷�ID As Long, blnValid As Boolean
    Dim sng���� As Single, dbl���� As Double
    Dim lngRow As Long
    
    cmdƵ��.Tag = rsInput!����
    txtƵ��.Text = rsInput!����
    txtƵ��.Tag = "1"
        
    With vsAdvice
        '������ִ��ʱ��Ŀ�����:����Ƶ���л�
        lngRow = GetBaseRow(.Row)
        If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Or InStr(",5,6,7,", .TextMatrix(lngRow, COL_���)) > 0 Then
            If Not cboִ��ʱ��.Enabled Then SetItemEditable , , , , 1
        Else
            If cboִ��ʱ��.Enabled Then SetItemEditable , , , , -1
        End If
        
        If cboִ��ʱ��.Enabled Then '"��ѡƵ��"��ҩƷʱ
            If rsInput!�����λ = "����" Then
                If cboִ��ʱ��.Text <> "" Then cboִ��ʱ��.Tag = "1"
                cboִ��ʱ��.Text = ""
            Else
                '�������ִ��ʱ�䷽���ı仯
                If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                    '���Ҹ�ҩ;����Ӧ����
                    lng�÷�ID = .FindRow(CLng(.TextMatrix(.Row, COL_���ID)), .Row + 1)
                    If lng�÷�ID <> -1 Then 'δ�ҵ���ҩ;�������,Ӧ�ò�����
                        lng�÷�ID = Val(.TextMatrix(lng�÷�ID, COL_������ĿID))
                    Else
                        lng�÷�ID = 0
                    End If
                ElseIf RowIn�䷽��(.Row) Then
                    '�õ���Ӧ����ҩ�÷�ID
                    lng�÷�ID = Val(.TextMatrix(.Row, COL_������ĿID))
                End If
                
                Call Getʱ�䷽��(cboִ��ʱ��, int��Χ, txtƵ��.Text, lng�÷�ID)
                'ȡ�µ�Ƶ�ʵ�Ĭ��ִ��ʱ��
                If cboִ��ʱ��.ListCount > 0 Then
                    cboִ��ʱ��.ListIndex = 0
                    cboִ��ʱ��.Tag = "1"
                Else
                    '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                    If cboִ��ʱ��.Text <> "" Then
                        blnValid = ExeTimeValid(cboִ��ʱ��.Text, rsInput!Ƶ�ʴ���, rsInput!Ƶ�ʼ��, rsInput!�����λ)
                        If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                            cboִ��ʱ��.Text = ""
                            cboִ��ʱ��.Tag = "1"
                        End If
                    End If
                End If
            End If
            
            '���¼�������
            If mbln���� And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                sng���� = Val(txt����.Text)
                If sng���� = 0 Then sng���� = 1
                
                If txtƵ��.Text <> "" And Val(txt����.Text) <> 0 _
                    And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                    
                    txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                        Val(txt����.Text), sng����, rsInput!Ƶ�ʴ���, _
                        rsInput!Ƶ�ʼ��, rsInput!�����λ, cboִ��ʱ��.Text, _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), _
                        Val(.TextMatrix(.Row, COL_�����װ)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") = 0 Then
                        txt����.Text = IntEx(Val(txt����.Text))
                    ElseIf Val(.TextMatrix(.Row, COL_�ɷ����)) <> 0 Then
                        txt����.Text = IntEx(Val(txt����.Text))
                    End If
                    txt����.Tag = "1"
                End If
            End If
        End If
        If rsInput!�����λ = "����" Then
            If cboִ��ʱ��.Enabled Then SetItemEditable , , , , -1
        End If
    End With
        
    '����ǰҽ��ִ��Ƶ�ʵı仯
    Call AdviceChange
End Sub

Private Function GetBaseRow(ByVal lngRow As Long) As Long
'���ܣ��ɵ�ǰ�ɼ��л�ȡ����Ŀ����
    If RowIn�䷽��(lngRow) Then
        '��ȡ��ҩ�䷽��һζ��ҩ��
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
    ElseIf RowIn������(lngRow) Then
        '��ȡһ�������ĵ�һ����Ŀ��
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
    Else
        GetBaseRow = lngRow
    End If
End Function

Private Sub cbo����_Change()
    cbo����.Tag = "1"
End Sub

Private Sub cbo����_Click()
    cbo����.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo����_GotFocus()
    zlControl.TxtSelAll cbo����
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo����_Validate(False)
    ElseIf InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cbo����.Text) > 10 Then
        MsgBox "�����������ݹ��������������Ƿ���ȷ��", vbInformation, gstrSysName
        Call cbo����_GotFocus: Cancel = True: Exit Sub
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub cbo����ִ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo����ִ��.ListIndex = -1 Then Exit Sub
    
    If cbo����ִ��.ItemData(cbo����ִ��.ListIndex) = -1 Then
        strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And B.������� IN(1,3)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
        vRect = GetControlRect(cbo����ִ��.hWnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lbl����ִ��.Caption, , , , , , True, vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then
                cbo����ִ��.ListIndex = intIdx
            Else
                cbo����ִ��.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ִ��.ListCount - 1
                cbo����ִ��.ItemData(cbo����ִ��.NewIndex) = rsTmp!ID
                cbo����ִ��.ListIndex = cbo����ִ��.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = SeekCboIndex(cbo����ִ��, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
            Call zlControl.CboSetIndex(cbo����ִ��.hWnd, intIdx)
        End If
    Else
        cbo����ִ��.Tag = "1"
        lngRow = vsAdvice.Row
        
        '���¸����˵�ִ�п���ҽ������
       Call AdviceChange
    End If
End Sub

Private Sub cbo����ִ��_GotFocus()
    Call zlControl.TxtSelAll(cbo����ִ��)
End Sub

Private Sub cbo����ִ��_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo����ִ��.ListIndex = -1 Then
            Call cbo����ִ��_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo����ִ��_Validate(False)
        End If
    End If
End Sub

Private Sub cbo����ִ��_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo����ִ��.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cbo����ִ��.Text = "" Then '������
        With vsAdvice
            'ԭҺƤ��
            If .TextMatrix(.Row, COL_���) = "E" And .TextMatrix(.Row, COL_��������) = "1" And .TextMatrix(.Row, COL_ִ�з���) = "5" Then
                cbo����ִ��.Tag = "1"
                Call AdviceChange
                Exit Sub
            Else
                If cbo����ִ��.ListCount > 0 Then Cancel = True
                Exit Sub
            End If
        End With
    End If
    
    On Error GoTo errH
    
    '�Ƿ���������ѡ�����
    blnLimit = True
    If cbo����ִ��.ListCount > 0 Then
        If cbo����ִ��.ItemData(cbo����ִ��.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    strInput = UCase(NeedName(cbo����ִ��.Text))
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(1,3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then cbo����ִ��.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo����ִ��.ListIndex = -1 Then
            MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cbo����ִ��.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl����ִ��.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then
                cbo����ִ��.ListIndex = intIdx
            Else
                cbo����ִ��.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ִ��.ListCount - 1
                cbo����ִ��.ItemData(cbo����ִ��.NewIndex) = rsTmp!ID
                cbo����ִ��.ListIndex = cbo����ִ��.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboҽ������_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim strTmp As String, arrTmp As Variant
    Dim objControl As CommandBarControl
    Dim i As Long
    
    If CommandBar Is Nothing Then Exit Sub
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
    Case conMenu_Tool_PlugIn
        Call CreatePlugInOK(p����ҽ���´�, mint����)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            If mint���� = 0 Then 'ҽ��վ����
                strTmp = gobjPlugIn.GetFuncNames(glngSys, p����ҽ���´�, 0)
            ElseIf mint���� = 1 Then '��ʿվ����
                strTmp = gobjPlugIn.GetFuncNames(glngSys, p����ҽ���´�, 1)
            ElseIf mint���� = 2 Then 'ҽ��վ����
                strTmp = gobjPlugIn.GetFuncNames(glngSys, p����ҽ���´�, 2)
            End If
            Call zlPlugInErrH(err, "GetFuncNames")
            err.Clear: On Error GoTo 0
        End If
        If strTmp <> "" Then
            With CommandBar.Controls
                If .Count = 0 Then
                    strTmp = Replace(strTmp, "Auto:", "")
                    arrTmp = Split(strTmp, ",")
                    For i = 0 To UBound(arrTmp)
                        Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
                        If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIF(i = 9, 0, i + 1) & ")"
                        objControl.IconId = conMenu_Tool_PlugIn_Item
                        objControl.Parameter = arrTmp(i)
                    Next
                End If
            End With
        End If
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngColW As Long, i As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    fraPati.Left = lngLeft
    fraPati.Top = lngTop
    fraPati.Width = lngRight - lngLeft
    fraPati.Height = cmdAlley.Top + cmdAlley.Height + 60
    
    fra���.Left = lngLeft
    fra���.Top = fraPati.Top + fraPati.Height
    fra���.Width = lngRight - lngLeft
    
    opt���(1).Left = fra���.Width - opt���(1).Width - 10 * Screen.TwipsPerPixelX
    opt���(0).Left = opt���(1).Left
    vsDiag.Width = opt���(0).Left - vsDiag.Left - 100
    For i = 0 To vsDiag.Cols - 1
        If Not vsDiag.ColHidden(i) And i <> col��� Then
            lngColW = lngColW + vsDiag.ColWidth(i)
        End If
    Next
    vsDiag.ColWidth(col���) = vsDiag.Width - lngColW - 2 * Screen.TwipsPerPixelX
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = fraPati.Top + fraPati.Height + fra���.Height
    vsAdvice.Height = lngBottom - lngTop - fraPati.Height - fra���.Height - (fraAdvice.Height - 80) - IIF(pic����.Visible, pic����.Height, 0)
    vsAdvice.Width = lngRight - lngLeft
    
    pic����.Left = 0
    pic����.Top = vsAdvice.Top + vsAdvice.Height
    pic����.Width = vsAdvice.Width
    lbl����.Width = pic����.ScaleWidth - lbl����.Left - 45
    
    fraAdvice.Left = lngLeft
    fraAdvice.Top = vsAdvice.Top + vsAdvice.Height - 6 * Screen.TwipsPerPixelX + IIF(pic����.Visible, pic����.Height, 0)
    fraAdvice.Width = lngRight - lngLeft
    
    stbThis.Top = lngBottom - 10
    stbThis.Width = picSub.ScaleWidth
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim vRect As RECT, vPos As PointAPI
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    
    'PASS�����˵��Ŀɼ�����������ڶ������������ã���ʱ��Ҫ���Σ��粻���Σ�����������ֵ���ᱻ����
    If Between(Control.ID, conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99) Then
        Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_Delete
            With vsAdvice
                blnEnabled = True
                If .RowData(.Row) <> 0 Then
                    If Not fraAdvice.Enabled And Val(.TextMatrix(.Row, COL_���״̬)) <> 2 Then blnEnabled = False
                    If Not fraAdvice.Enabled And Val(.TextMatrix(.Row, COL_���״̬)) = 2 And .TextMatrix(.Row, COL_���) = "K" And gblnѪ��ϵͳ Then blnEnabled = False
                    If .TextMatrix(.Row, COL_���) = "K" And Val(.TextMatrix(.Row, COL_��鷽��)) = 1 And Val(.TextMatrix(.Row, COL_���״̬)) = 1 Then blnEnabled = False
                    If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then blnEnabled = False
                    '��ǩ��ҽ������ɾ��
                    If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then blnEnabled = False
                End If
                Control.Enabled = blnEnabled
            End With
        Case conMenu_Merge
            Control.Checked = mblnRowMerge
        Case conMenu_Scheme, conMenu_Scheme * 10# + 1, conMenu_Scheme * 10# + 2
                        If mblnModal Then
                Control.Visible = False
            Else
            If InStr(GetInsidePrivs(p����ҽ���´�), "������׷���") = 0 Then
                Control.Visible = False
            End If
            If Control.ID = conMenu_Scheme * 10# + 2 Then
                Control.Caption = IIF(mfrmShortCut.Visible, "����", "��ʾ") & "���׷���ѡ����"
            End If
                        End If
        Case conMenu_Agent
            If Not mblnAddAgent Then
                Control.Visible = False
            Else
                strTmp = GetTsPrivs(p����ҽ���´�)
                Control.Visible = InStr(strTmp, "�´�����ҩ��") > 0 Or InStr(strTmp, "�´ﶾ��ҩ��") > 0 Or InStr(strTmp, "�´ﾫ��ҩ��") > 0
            End If
        Case conMenu_Save
            Control.Enabled = mblnNoSave
        Case conMenu_Sign
            If mint���� = 1 Or InStr(UserInfo.����, "ҽ��") = 0 Or gobjESign Is Nothing _
                Or InStr(GetInsidePrivs(p����ҽ���´�), ";ҽ���´�;") = 0 Then
                Control.Visible = False
            ElseIf mint���� = 0 And Control.Category <> "���ж�" Then
                If CheckSign(0, 0, mlngҽ������ID, mlng���˿���id, 1, False, gobjESign) = False Then
                    Control.Visible = False '��ͬ����û������Ҫʹ��ǩ��
                End If
                Control.Category = "���ж�"
            ElseIf mint���� = 2 And Control.Category <> "���ж�" Then
                If CheckSign(3, 0, mlngҽ������ID, mlng���˿���id, 1, False, gobjESign) = False Then
                    Control.Visible = False '��ͬ����û������Ҫʹ��ǩ��
                End If
                Control.Category = "���ж�"
            End If
        Case conMenu_Send, conMenu_Send * 10# + 1, conMenu_Send * 10# + 2
            If InStr(GetInsidePrivs(p����ҽ���´�), "ҽ������") = 0 Then
                Control.Visible = False
            ElseIf InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ�շѵ�") = 0 And InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ���ʵ�") = 0 Then
                Control.Visible = False
            End If
            If Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) = 0 And InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ�շѵ�") = 0 Or _
               Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) = 1 And InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ���ʵ�") = 0 Then
                Control.Visible = False
            End If
                Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
                Control.Enabled = vsAdvice.RowData(vsAdvice.Row) <> 0 And InStr(",5,6,7,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) > 0
        Case conMenu_Reference
            If GetInsidePrivs(pҩƷ���Ʋο�) = "" Then
                Control.Visible = False
            End If
        Case Else
            '������ʾ״̬
            blnEnabled = False
            If txt����.Enabled Then
                If InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) > 0 Then
                    GetCursorPos vPos
                    GetWindowRect fraAdvice.hWnd, vRect
                    If Between(vPos.X * Screen.TwipsPerPixelX, vRect.Left * Screen.TwipsPerPixelX + lbl����.Left, vRect.Left * Screen.TwipsPerPixelX + lbl����.Left + lbl����.Width) Then
                        If Between(vPos.Y * Screen.TwipsPerPixelY, vRect.Top * Screen.TwipsPerPixelY + lbl����.Top, vRect.Top * Screen.TwipsPerPixelY + lbl����.Top + lbl����.Height) Then
                            blnEnabled = True
                        End If
                    End If
                End If
            End If
            Call Set����Face(blnEnabled)
    End Select
End Sub

Private Sub chk����_Click()
    If Not mblnDoCheck Then Exit Sub
    
    chk����.Tag = "1"
    '��������
    Call AdviceChange
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmdAlley_Click()
'���ܣ��Բ��˹���ʷ/����״̬���й���
    'Pass
    If mblnPass Then
        Call gobjPass.zlPassCmdAlleyManage(mobjPassMap)
    End If
End Sub

Private Sub cmdLastDiag_Click()
'���ܣ���ȡ�ϴξ���������Ϣ��׷�ӵ�������Ϻ�
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Long, strTmp As String
    
    If MsgBox("�Ƿ��ȡ�ϴ������Ϣ��ӵ��������֮��", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        On Error GoTo errH
        strSQL = "Select A.ID,A.��¼��Դ,A.�������,A.����ID,A.���ID,A.֤��ID,A.�������,A.�Ƿ�����,c.���� as ICD��,D.���� as ��ϱ���,E.���� as ֤�����,A.����ʱ��" & vbNewLine & _
            "From ������ϼ�¼ A, ���˹Һż�¼ B,��������Ŀ¼ C, �������Ŀ¼ D,��������Ŀ¼ E" & vbNewLine & _
            "Where a.����id = b.����id And a.��ҳid = b.Id And A.����ID=C.ID(+)  And a.���id = D.Id(+) And  a.֤��ID=E.ID(+) And a.����id = [1] And (������� = 1 Or ������� = 11) And" & vbNewLine & _
            "   (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And b.�Ǽ�ʱ�� =" & vbNewLine & _
            "      (Select Max(�Ǽ�ʱ��)" & vbNewLine & _
            "       From ���˹Һż�¼ C" & vbNewLine & _
            "       Where c.����id = [1] And ID <> [2] And Exists (Select 1 From ������ϼ�¼ D Where d.����id = [1] And ��ҳid = c.Id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
        If rsTmp.RecordCount = 0 Then
            MsgBox "�ò����ڱ��ξ���֮ǰδ��д��������ϻ�����д���������ͣ�á�", vbInformation
        Else
            With vsDiag
                Do While Not rsTmp.EOF
                    i = -1
                    If Val(rsTmp!���id & "") <> 0 Then i = .FindRow(Val(rsTmp!���id & ""), , col���ID)
                    If i = -1 And Val(rsTmp!����id & "") <> 0 Then i = .FindRow(Val(rsTmp!����id & ""), , col����ID)
                    '�����ǰ�Ѿ���д����ֻ���·���ʱ��
                    If i <> -1 Then
                        If .TextMatrix(i, col����ʱ��) = "" Then
                            .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                        End If
                    Else
                        If Not (.TextMatrix(1, col���) = "" And .Rows = 2) Then .AddItem ""
                        i = .Rows - 1
                        Call SetDiagType(i, rsTmp!�������)
                        
                        If IsNull(rsTmp!�������) Then
                            .TextMatrix(i, col����) = ""
                            .TextMatrix(i, col���) = ""
                        Else
                            If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���id & "") = 0 And Val(rsTmp!����id & "") = 0) Then '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                                '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                                If Val(rsTmp!����id & "") <> 0 Then
                                    .TextMatrix(i, col����) = Nvl(rsTmp!ICD��)
                                ElseIf Val(rsTmp!���id & "") <> 0 Then
                                    .TextMatrix(i, col����) = Nvl(rsTmp!��ϱ���)
                                Else
                                    .TextMatrix(i, col����) = ""
                                End If
                                .TextMatrix(i, col���) = rsTmp!�������
                            Else
                                .TextMatrix(i, col����) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                                .TextMatrix(i, col���) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                            End If
                        End If

                        .Cell(flexcpData, i, col����) = Val(Nvl(rsTmp!�Ƿ�����, 0))
                        .Cell(flexcpForeColor, i, col����) = IIF(Nvl(rsTmp!�Ƿ�����, 0) = 1, vbRed, .GridColor)
                        
                        .TextMatrix(i, col���ID) = Nvl(rsTmp!���id, 0)
                        .Cell(flexcpData, i, col���ID) = Nvl(rsTmp!ID, 0)
                        .TextMatrix(i, col����ID) = Nvl(rsTmp!����id, 0)
                        .TextMatrix(i, col֤��ID) = Nvl(rsTmp!֤��id, 0)
                        .TextMatrix(i, colICD��) = Nvl(rsTmp!ICD��)
                        .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                        'ȡ֤������
                        If InStr(.TextMatrix(i, col���), "(") > 0 And InStr(.TextMatrix(i, col���), ")") > 0 Then
                            strTmp = Mid(.TextMatrix(i, col���), InStrRev(.TextMatrix(i, col���), "(") + 1)
                            strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                            '��ȡ֤��
                            .TextMatrix(i, col��ҽ֤��) = strTmp
                            'ȥ�����������֤��
                            .TextMatrix(i, col���) = Mid(.TextMatrix(i, col���), 1, InStrRev(.TextMatrix(i, col���), "(") - 1)
                        Else
                           .TextMatrix(i, col��ҽ֤��) = ""
                        End If
                        '����¼����ϵ������������Ҫȥ��֤����˴˾�������
                        If Not IsNull(rsTmp!����id) Or Not IsNull(rsTmp!���id) Then
                            .Cell(flexcpData, i, col���) = Get�������(Val("" & rsTmp!���id), Val("" & rsTmp!����id))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                        Else
                            .Cell(flexcpData, i, col���) = .TextMatrix(i, col���)
                        End If
                    End If
                    
                    rsTmp.MoveNext
                Loop
                mblnNoSave = True: lbl���.Tag = "1"
                Call SetDiagHeight
            End With
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdReason_Click()
    Call ReasonSelect("", 1)
    Call AdviceChange
End Sub

Private Sub cmdExcReason_Click()
    Call ReasonSelect("", 3)
    Call AdviceChange
End Sub

Private Function ReasonSelect(ByVal strFind As String, ByVal intType As Integer) As Boolean
'�������кͿ�����ҩ���ɳ���˵��ѡ����
'intType  1-������ҩ���ɣ�2-�������У�3-����˵��
    Dim blnCancle As Boolean
    Dim strRetrun As String
    Dim lngLeft As Long, lngTop As Long
    Dim strName As String
    
    If intType = 1 Then
        lngLeft = txt��ҩ����.Left
        lngTop = txt��ҩ����.Top
        strName = "������ҩ���ɡ�"
    ElseIf intType = 2 Then
        lngLeft = cboҽ������.Left
        lngTop = cboҽ������.Top
        strName = "�������С�"
    ElseIf intType = 3 Then
        lngLeft = txt����˵��.Left
        lngTop = txt����˵��.Top
        strName = "����˵����"
    End If
    
    lngLeft = lngLeft + fraAdvice.Left + Me.Left
    lngTop = lngTop + fraAdvice.Top + Me.Top - 2600
    
    strRetrun = frmKssReasonSelect.ShowMe(Me, strFind, blnCancle, lngLeft, lngTop, intType)
    If Not blnCancle Then
        If strRetrun = "" Then
            If strFind = "" Then
                MsgBox "û���ҵ����õ�" & strName, vbInformation, Me.Caption
            End If
        Else
            If intType = 1 Then
                txt��ҩ����.Text = strRetrun
            ElseIf intType = 2 Then
                cboҽ������.Text = strRetrun
            ElseIf intType = 3 Then
                txt����˵��.Text = strRetrun
            End If
        End If
    End If
    ReasonSelect = blnCancle
End Function

Private Sub ReasonSave(ByVal intType As Integer)
'���ܣ�������ҩ���ɺͳ���˵������
'������intType  0-������ҩ���ɣ�1-����˵��
    Dim strSQL As String, rsTmp As Recordset
    Dim strTmp As String
    Dim strPar As String
    
    If txt��ҩ����.Text = "" And intType = 0 Then MsgBox "����������Ҫ�ղص���ҩ���ɡ�", vbInformation, Me.Caption: txt��ҩ����.SetFocus: Exit Sub
    If txt����˵��.Text = "" And intType = 1 Then MsgBox "����������Ҫ�ղصĳ���˵����", vbInformation, Me.Caption: txt����˵��.SetFocus: Exit Sub
    
    If intType = 0 Then
        strPar = txt��ҩ����.Text
        strTmp = "��ҩ����"
    ElseIf intType = 1 Then
        strPar = txt����˵��.Text
        strTmp = "����˵��"
    End If
    
    On Error GoTo errH
    strSQL = "Select 1 From ҽ������ԭ�� Where ����=[1]"
    If intType = 1 Then strSQL = strSQL & " And ����=1 And ��Ա=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPar, UserInfo.����)
    '����Ѿ����ˣ���ʾ�û��Ƿ������
    If rsTmp.RecordCount > 0 Then
        MsgBox "�Ѿ�������ͬ��" & strTmp & "��", vbInformation, Me.Caption
        Exit Sub
    End If
    strSQL = "zl_ҽ������ԭ��_Update(0,Null,'" & strPar & "',Null" & _
        IIF(intType = 1, ",1,'" & UserInfo.���� & "'", "") & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    MsgBox strTmp & "�ղسɹ���", vbInformation, Me.Caption
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd��������_Click()
    Dim strSQL As String, i As Integer
    Dim rsTmp As Recordset
    
    If Trim(cboҽ������.Text) = "" Then
        MsgBox "�������������ݡ�", vbInformation, gstrSysName
        If cboҽ������.Enabled Then cboҽ������.SetFocus
        Exit Sub
    End If
    On Error GoTo errH
    strSQL = "Select 1 From �������� Where ����=[1] And (��Ա=[2] Or ��Ա is null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(cboҽ������.Text), UserInfo.����)
    If rsTmp.RecordCount > 0 Then
        MsgBox "�����������Ѿ��ڳ��������С�", vbInformation, gstrSysName
        If cboҽ������.Enabled Then cboҽ������.SetFocus
        Exit Sub
    End If
    
    strSQL = zlCommFun.zlGetSymbol(cboҽ������.Text, CByte(mint����))
    strSQL = "zl_��������_Insert('" & Replace(cboҽ������.Text, "'", "''") & "','" & strSQL & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem cboҽ������.hWnd, CB_ADDSTRING, 0, cboҽ������.Text
    MsgBox "������Ϊ�������С�", vbInformation, gstrSysName
    If cboҽ������.Enabled Then cboҽ������.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdƵ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
    Dim strSeek As String, lng������ĿID As Long, lngFind As Long
    
    On Error GoTo errH
    
    With vsAdvice
        int��Χ = GetƵ�ʷ�Χ(.Row)
        
        If txtƵ��.Text <> "" Then
            strSQL = "Select ���� From ����Ƶ����Ŀ Where ����=[1] And ���÷�Χ=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtƵ��.Text, int��Χ)
            If Not rsTmp.EOF Then strSeek = rsTmp!����
        End If
        
        '��ѡ��Ƶ�ʵĳ���Ƶ��
        lng������ĿID = Val(.TextMatrix(.Row, COL_������ĿID))
        If RowIn������(.Row) Then
            lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_���ID)
            If lngFind <> -1 Then
                lng������ĿID = Val(.TextMatrix(lngFind, COL_������ĿID))
            End If
        ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            lngFind = .FindRow(CLng(.TextMatrix(.Row, COL_���ID)), .Row + 1)
            If lngFind <> -1 Then
                lng������ĿID = Val(.TextMatrix(lngFind, COL_������ĿID))
            End If
        End If
        strSQL = ""
        If int��Χ = 1 Then
            strSQL = " And (Exists(Select 1 From �����÷����� Where ��ĿID=[2] And �÷�ID is NULL And Ƶ��=A.���� And A.���÷�Χ=1)" & _
                " Or (Select Count(*) From �����÷����� Where ��ĿID=[2] And �÷�ID is NULL And Ƶ�� Is Not NULL)<=1)"
        End If
        strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
            " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ" & _
            " From ����Ƶ����Ŀ A Where A.���÷�Χ=[1]" & strSQL & _
            " Order by A.����"
        vRect = GetControlRect(txtƵ��.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����Ƶ��", False, strSeek, "", False, False, True, _
            vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, int��Χ, lng������ĿID)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�п��õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
            End If
            txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
            Call zlControl.TxtSelAll(txtƵ��)
            txtƵ��.SetFocus: Exit Sub
        End If
        Call SetƵ��Input(rsTmp, int��Χ)
        txtƵ��.SetFocus
        Call SeekNextControl
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd����ʱ��_Click()
    If IsDate(txt����ʱ��.Text) Then
        dtpDate.value = CDate(txt����ʱ��.Text)
    ElseIf IsDate(txt��ʼʱ��.Text) Then
        dtpDate.value = CDate(txt��ʼʱ��.Text)
    Else
        dtpDate.value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "����ʱ��"
    dtpDate.Left = txt����ʱ��.Left + fraAdvice.Left
    dtpDate.Top = txt����ʱ��.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub cmd�ղ���ҩ����_Click()
    Call ReasonSave(0)
End Sub

Private Sub cmdComExcReason_Click()
    Call ReasonSave(1)
End Sub

Private Sub cmdҽ������_Click()
    If ReasonSelect("", 2) Then Exit Sub
    cboҽ������.Tag = "1"
    Call AdviceChange
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnNoSave Then
        If MsgBox("��ǰҽ�����ݱ༭����δ���棬ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    If Not mfrmShortCut Is Nothing Then mfrmShortCut.SaveShowState 'ϵͳ�Զ�ж�ظ��Ӵ���
End Sub

Private Sub Set����Face(ByVal blnOver As Boolean)
    If blnOver Then
        If lbl����.BorderStyle = 0 Then
            lbl����.BorderStyle = 1
            lbl����.BackStyle = 1
        End If
    Else
        If lbl����.BorderStyle = 1 Then
            lbl����.BorderStyle = 0
            lbl����.BackStyle = 0
        End If
    End If
End Sub

Private Sub lbl����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vRect As RECT, strSQL As String
    Dim str��λ As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    
    If Not (InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) > 0 And txt����.Enabled) Then Exit Sub
    
    If mrsDrugScale Is Nothing Then
        strSQL = "Select ����,���� From ���ü������� Where ���� is Not NULL And ���� is Not NULL Order by ����"
        Set mrsDrugScale = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If mrsDrugScale.EOF Then
            MsgBox "û�����ó��ü������������ȵ��ֵ�����߽������á�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_�շ�ϸĿID) <> 0 Then
        Set rsTmp = Get�շ���Ŀ��¼(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�շ�ϸĿID)))
        str��λ = rsTmp!���㵥λ & ""
    End If
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        mrsDrugScale.MoveFirst
        Do While Not mrsDrugScale.EOF
            Set objControl = .Add(xtpControlButton, conMenu_DrugScale * 100# + .Count + 1, mrsDrugScale!���� & "[" & str��λ & "]")
            objControl.Parameter = mrsDrugScale!����
            mrsDrugScale.MoveNext
        Loop
    End With
    GetWindowRect fraAdvice.hWnd, vRect
    objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + lbl����.Left + lbl����.Width, vRect.Top * Screen.TwipsPerPixelY + lbl����.Top
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lbl����_Click()
    Call Load��Һ����(cbo����, lbl���ٵ�λ, True)
    cbo����.Tag = "1"
    Call AdviceChange
End Sub

Private Sub mfrmPrice_PanelHide()
    Call stbThis_PanelClick(stbThis.Panels("Price"))
End Sub

Private Sub mfrmSend_EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, �Һŵ�, Succeed)
End Sub

Private Sub mfrmShortCut_ItemClick(ByVal ���� As Integer, ByVal ����ID As Long)
    If cmdSel.Enabled And cmdSel.Visible Then
        Call ClinicSelecter(����, ����ID)
    End If
End Sub

Private Sub picHelp_Click()
    Dim strTip As String
    
    On Error Resume Next
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "��" Then
        strTip = COL_����ִ��
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "��" Then
        strTip = COL_����ִ��
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "Сʱ" Then
        strTip = COL_��ʱִ��
    End If
    MsgBox strTip, vbInformation, Me.Caption
    cboִ��ʱ��.SetFocus
    mblnIsInHelp = False
End Sub

Private Sub picHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    On Error Resume Next
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "��" Then
        strTip = COL_����ִ��
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "��" Then
        strTip = COL_����ִ��
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "Сʱ" Then
        strTip = COL_��ʱִ��
    End If
    
    zlCommFun.ShowTipInfo picHelp.hWnd, strTip, True
    
    If X >= 0 And X <= picHelp.Width And Y >= 0 And Y <= picHelp.Height Then
        mblnIsInHelp = True
        SetCapture picHelp.hWnd
    Else
        mblnIsInHelp = False
        ReleaseCapture
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "Price" Then
        If Panel.Bevel <> sbrNoBevel Then
            Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Panel.Tag = IIF(Panel.Bevel = sbrInset, "1", "")
            Call ShowPrice(vsAdvice.Row)
        End If
    ElseIf Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("���뷽ʽ", IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0)))
        mint���� = IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
    ElseIf Panel.Key = "KB" Then
        On Error Resume Next
        If mobjKeyBoard Is Nothing Then Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
        Call mobjKeyBoard.StartUp
        Call mobjKeyBoard.SetPos
        err.Clear: On Error GoTo 0
    End If
End Sub

Private Sub txt����˵��_GotFocus()
    Call zlControl.TxtSelAll(txt����˵��)
End Sub

Private Sub txt����˵��_Change()
    txt����˵��.Tag = "1"
End Sub

Private Sub txt����˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����˵��.Text <> "" Then
            If ReasonSelect(txt����˵��.Text, 3) Then Exit Sub
        End If
        If SeekNextControl Then Call txt����˵��_Validate(False)
    End If
End Sub

Private Sub txt����˵��_Validate(Cancel As Boolean)
    Call AdviceChange
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = 192 Then '~
        Call lbl����_MouseDown(1, 0, 0, 0)
    End If
End Sub

Private Sub txt����_LostFocus()
    mblnReturn = False
End Sub

Private Sub txtƵ��_GotFocus()
    Call zlControl.TxtSelAll(txtƵ��)
End Sub

Private Sub txtƵ��_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
    Dim lng������ĿID As Long, lngFind As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If cmdƵ��.Tag <> "" And txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��) And txtƵ��.Text <> "" Then
                Call SeekNextControl
            ElseIf txtƵ��.Text = "" Then
                If cmdƵ��.Enabled And cmdƵ��.Visible Then cmdƵ��_Click
            Else
                int��Χ = GetƵ�ʷ�Χ(.Row)
                
                '��ѡ��Ƶ�ʵĳ���Ƶ��
                lng������ĿID = Val(.TextMatrix(.Row, COL_������ĿID))
                If RowIn������(.Row) Then
                    lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_���ID)
                    If lngFind <> -1 Then
                        lng������ĿID = Val(.TextMatrix(lngFind, COL_������ĿID))
                    End If
                ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(.Row, COL_���ID)), .Row + 1)
                    If lngFind <> -1 Then
                        lng������ĿID = Val(.TextMatrix(lngFind, COL_������ĿID))
                    End If
                End If
                strSQL = ""
                If int��Χ = 1 Then
                    strSQL = " And (Exists(Select 1 From �����÷����� Where ��ĿID=[4] And �÷�ID is NULL And Ƶ��=A.���� And A.���÷�Χ=1)" & _
                        " Or (Select Count(*) From �����÷����� Where ��ĿID=[4] And �÷�ID is NULL And Ƶ�� Is Not NULL)<=1)"
                End If
                strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
                    " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ" & _
                    " From ����Ƶ����Ŀ A Where A.���÷�Χ=[3]" & strSQL & _
                    " And (A.���� Like [1] Or Upper(A.����) Like [2]" & _
                    " Or Upper(A.����) Like [2] Or Upper(A.Ӣ������) Like [2])" & _
                    " Order by A.����"
                vRect = GetControlRect(txtƵ��.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����Ƶ��", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, UCase(txtƵ��.Text) & "%", mstrLike & UCase(txtƵ��.Text) & "%", int��Χ, lng������ĿID)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ�������Ƶ����Ŀ��", vbInformation, gstrSysName
                    End If
                    txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
                    Call zlControl.TxtSelAll(txtƵ��)
                    txtƵ��.SetFocus: Exit Sub
                End If
                Call SetƵ��Input(rsTmp, int��Χ)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub txt����ʱ��_Change()
    txt����ʱ��.Tag = "1"
End Sub

Private Sub txt����ʱ��_GotFocus()
    If txt����ʱ��.Text = "" Then txt����ʱ��.Text = txt��ʼʱ��.Text
    zlControl.TxtSelAll txt����ʱ��
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����ʱ��.Text <> "" Then
            txt����ʱ��.Text = GetFullDate(txt����ʱ��.Text)
            If SeekNextControl Then Call txt����ʱ��_Validate(False)
        End If
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Locked Then Exit Sub
        
    If Not IsDate(txt����ʱ��.Text) Then
        If txt����ʱ��.Text <> "" Then
            Cancel = True
            txt����ʱ��_GotFocus
            Exit Sub
        ElseIf vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If IsDate(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��)) Then
                '�ָ���Ϊ�����ȱʡΪ��ʼʱ��
                txt����ʱ��.Text = vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��)
            End If
        End If
    Else
        '���ʱ��Ϸ���
        If Not Check����ʱ��(txt����ʱ��.Text, vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��), vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) Then
            Cancel = True
            txt����ʱ��_GotFocus
            Exit Sub
        End If
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub txt����_Change()
    With vsAdvice
        If .RowData(.Row) <> 0 Then
            If Val(.TextMatrix(.Row, COL_����)) <> Val(txt����.Text) Then
                txt����.Tag = "1"
            End If
        Else
            txt����.Tag = "1"
        End If
    End With
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txt����.Text) > 0 Then
            Call txt����_Validate(blnCancel)
            If Not blnCancel Then mblnReturn = True: Call SeekNextControl
        Else
            If Val(txt����.Text) = 0 Then txt����.Text = ""
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_LostFocus()
    mblnReturn = False
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim sng���� As Single, i As Long
    Dim strSame As String, strMsg As String
    Dim dbl���� As Double
    Dim strTmpTag As String
    Dim bln�������� As Boolean
    
    With vsAdvice
        If Val(txt����.Text) = 0 Then txt����.Text = ""
        If mblnReturn Then mblnReturn = False: Exit Sub
        If Val(txt����.Text) <= 0 Then
            Cancel = True: txt����_GotFocus: Exit Sub
        End If
        
        '����������Ҫһ��Ƶ��ͬ�ڵ�����
        If Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 Then
            If .TextMatrix(.Row, COL_�����λ) = "��" Then
                sng���� = 7
            ElseIf .TextMatrix(.Row, COL_�����λ) = "��" Then
                sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��))
            ElseIf .TextMatrix(.Row, COL_�����λ) = "Сʱ" Then
                sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) \ 24
            ElseIf .TextMatrix(.Row, COL_�����λ) = "����" Then
                sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) \ (24 * 60)
            End If
            If Val(txt����.Text) < sng���� Then
                If MsgBox("��""" & .TextMatrix(.Row, COL_Ƶ��) & """ִ��ʱ��������Ҫ " & sng���� & " �����ҩ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Cancel = True: txt����_GotFocus: Exit Sub
                End If
            End If
        End If
        
        dbl���� = Val(txt����.Text)
        If mbln�������� Then
            '�������Ϊ0����˵�������������������
            If dbl���� = 0 Or Val(txt����.Text) <> msngPre���� Then
                bln�������� = True
            End If
        Else
            bln�������� = True
        End If
        
        If bln�������� Then
            txt����.Text = ReGetҩƷ����(dbl����, Val(txt����.Text), Val(txt����.Text), .Row) '��ʽ����Change�¼�

            Call txt����_Validate(Cancel)
            If Cancel Then
                txt����.Text = dbl����
                Exit Sub
            End If
        End If
                         
        Call CheckDrugOutOfRange(.Row, Val(txt����.Text))
        
        msng���� = Val(txt����.Text)

    End With
    msngPre���� = Val(txt����.Text)
    Call AdviceChange
    
    '���׷�����������
    With vsAdvice
        If CStr(.Cell(flexcpData, .Row, COL_EDIT)) <> "" Then
            strSame = CStr(.Cell(flexcpData, .Row, COL_EDIT))
            If InStr(strSame, ",") > 0 Then
                strMsg = "�ôθ��Ƶ����е�ҩƷ�����������ִ����"
            Else
                strMsg = "�ó��׷���������ҩƷ�����������ִ����"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                For i = .FixedRows To .Rows - 1
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        If Not (Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) _
                            Or .RowData(i) = Val(.TextMatrix(.Row, COL_���ID)) Or i = .Row) _
                                And CStr(.Cell(flexcpData, i, COL_EDIT)) = strSame Then
                                
                            If .TextMatrix(i, COL_Ƶ��) <> "" And Val(.TextMatrix(i, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(i, COL_Ƶ�ʼ��)) <> 0 _
                                And Val(.TextMatrix(i, COL_����)) <> 0 And Val(.TextMatrix(i, COL_����ϵ��)) <> 0 And Val(.TextMatrix(i, COL_�����װ)) <> 0 Then
                                
                                .TextMatrix(i, COL_����) = txt����.Text
                                .TextMatrix(i, COL_����) = ReGetҩƷ����(Val(.TextMatrix(i, COL_����)), Val(.TextMatrix(i, COL_����)), Val(txt����.Text), i)
                                If Val(.TextMatrix(i, COL_���ID)) = Val(.RowData(i + 1)) Then
                                    .TextMatrix(i + 1, COL_����) = txt����.Text
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub txt�÷�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int���� As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long
    Dim strLike As String, i As Long, strWhere As String
        
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�)) And txt�÷�.Text <> "" Then
                Call SeekNextControl
            ElseIf txt�÷�.Text = "" Then
                If cmd�÷�.Enabled And cmd�÷�.Visible Then cmd�÷�_Click
            Else
                If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                    int���� = 2 '��ҩ;��
                ElseIf RowIn������(vsAdvice.Row) Then
                    int���� = 6 '�ɼ�����
                ElseIf .TextMatrix(.Row, COL_���) = "K" Then
                    If gblnѪ��ϵͳ = True Then
                        If Val(.TextMatrix(.Row, COL_��鷽��)) = 0 Then
                            int���� = 9 '�ɼ���Ѫ;��
                        Else
                            int���� = 8 '��Ѫ;��
                            strWhere = " And nvl(A.ִ�з���,0)=1 "
                        End If
                    Else
                        int���� = 8 '��Ѫ;��
                    End If
                Else
                    int���� = 4 '��ҩ�÷�
                End If
                If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
                    strSQL = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[4] And ����>0)" & _
                        " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                            " Where A.�÷�ID=B.ID And B.������� IN(1,3) And A.��ĿID=[4] And A.����>0)<=1)"
                End If
                
                '�Ż�
                strLike = mstrLike
                If Len(txt�÷�.Text) < 2 Then strLike = ""
                
                strSQL = "Select Distinct A.ID,A.����,A.����,A.ִ�з��� as ִ�з���ID" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.ID=B.������ĿID" & _
                    " And A.���='E' And A.��������=[3] And A.������� IN(1,3)" & strWhere & strSQL & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])" & _
                    " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[6])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                    Decode(mint����, 0, " And B.���� IN([5],3)", 1, " And B.���� IN([5],3)", "") & _
                    " Order by A.����"
                vRect = GetControlRect(txt�÷�.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl�÷�.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, UCase(txt�÷�.Text) & "%", _
                    strLike & UCase(txt�÷�.Text) & "%", CStr(int����), Val(.TextMatrix(.Row, COL_������ĿID)), mint���� + 1, mlng���˿���id)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ���" & lbl�÷�.Caption & "��", vbInformation, gstrSysName
                    End If
                    txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
                    Call zlControl.TxtSelAll(txt�÷�)
                    txt�÷�.SetFocus: Exit Sub
                End If
                
                '��һ����ҩ������ҩƷ�Ŀ��ø�ҩ;�����м��
                If int���� = 2 Then
                    Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
                    For i = lngBegin To lngEnd
                        If i <> .Row And .RowData(i) <> 0 Then
                            If Not Check�����÷�(rsTmp!ID, Val(.TextMatrix(i, COL_������ĿID)), 1) Then
                                .Refresh
                                MsgBox """" & rsTmp!���� & """���������뵱ǰҩƷһ����ҩ��""" & .TextMatrix(i, col_ҽ������) & """��", vbInformation, gstrSysName
                                .Refresh
                                txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
                                Call zlControl.TxtSelAll(txt�÷�)
                                txt�÷�.SetFocus: Exit Sub
                            End If
                        End If
                    Next
                End If
                If Val(cmd�÷�.Tag) <> Val(rsTmp!ID & "") Then .TextMatrix(.Row, COL_����) = ""
                Call Set�÷�Input(rsTmp, int����)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub cmd�÷�_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int���� As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim strSeek As String
    Dim strWhere As String
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            int���� = 2 '��ҩ;��
            lngBegin = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))), .Row + 1)
        ElseIf RowIn������(vsAdvice.Row) Then
            int���� = 6 '�ɼ�����
            lngBegin = .Row
        ElseIf .TextMatrix(.Row, COL_���) = "K" Then
            If gblnѪ��ϵͳ = True Then
                If Val(.TextMatrix(.Row, COL_��鷽��)) = 0 Then
                    int���� = 9 '�ɼ���Ѫ;��
                Else
                    int���� = 8 '��Ѫ;��
                    strWhere = " And nvl(A.ִ�з���,0)=1 "
                End If
            Else
                int���� = 8 '��Ѫ;��
            End If
            lngBegin = .FindRow(CStr(.RowData(.Row)), .Row + 1, COL_���ID)
        Else
            int���� = 4 '��ҩ�÷�
            lngBegin = .Row
        End If
        If txt�÷�.Text <> "" And lngBegin <> -1 Then
            strSeek = GetItemField("������ĿĿ¼", Val(.TextMatrix(lngBegin, COL_������ĿID)), "����")
        End If
        
        If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
            strSQL = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[2] And ����>0)" & _
                " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                    " Where A.�÷�ID=B.ID And B.������� IN(1,3) And A.��ĿID=[2] And A.����>0)<=1)"
        End If
        strSQL = "Select Distinct A.ID,A.����,A.����,C.���� as ����,A.ִ�з��� as ִ�з���ID" & _
            " From ������ĿĿ¼ A,���Ʒ���Ŀ¼ C" & _
            " Where A.����ID=C.ID(+) And A.���='E' And A.��������=[1] And A.������� IN(1,3)" & strWhere & strSQL & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[3])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
            " Order by A.����"
        vRect = GetControlRect(txt�÷�.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl�÷�.Caption, False, strSeek, "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, CStr(int����), Val(.TextMatrix(.Row, COL_������ĿID)), mlng���˿���id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�п��õ�" & lbl�÷�.Caption & "�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
            Call zlControl.TxtSelAll(txt�÷�)
            txt�÷�.SetFocus: Exit Sub
        End If
        
        '��һ����ҩ������ҩƷ�Ŀ��ø�ҩ;�����м��
        If int���� = 2 Then
            Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                If i <> .Row And .RowData(i) <> 0 Then
                    If Not Check�����÷�(rsTmp!ID, Val(.TextMatrix(i, COL_������ĿID)), 1) Then
                        .Refresh
                        MsgBox """" & rsTmp!���� & """���������뵱ǰҩƷһ����ҩ��""" & .TextMatrix(i, col_ҽ������) & """��", vbInformation, gstrSysName
                        .Refresh
                        txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
                        Call zlControl.TxtSelAll(txt�÷�)
                        txt�÷�.SetFocus: Exit Sub
                    End If
                End If
            Next
        End If
        
        Call Set�÷�Input(rsTmp, int����)
        txt�÷�.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub txt�÷�_GotFocus()
    Call zlControl.TxtSelAll(txt�÷�)
End Sub

Private Sub txt�÷�_LostFocus()
    'PASS
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            Call gobjPass.zlPassCloseDrugHint(mobjPassMap)
        End If
    End If
End Sub

Private Sub txt�÷�_Validate(Cancel As Boolean)
    With vsAdvice
        '�ָ���Ϊ�����
        If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Text <> IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�)) Then
            txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
        End If
    End With
End Sub

Private Sub txtƵ��_Validate(Cancel As Boolean)
    With vsAdvice
        '�ָ���Ϊ�����
        If cmdƵ��.Tag <> "" And txtƵ��.Text <> .TextMatrix(.Row, COL_Ƶ��) Then
            txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
        End If
    End With
End Sub

Private Sub cboӤ��_Click()
    If Not Visible Then Exit Sub
    If cboӤ��.ListIndex = -1 Then Exit Sub
    
    If cboӤ��.ListIndex = Val(cboӤ��.Tag) Then Exit Sub
    cboӤ��.Tag = cboӤ��.ListIndex
    
    Call ShowAdvice
    'PASS Ӥ�������ı�
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            mobjPassMap.PassPati.intӤ�� = cboӤ��.ListIndex
        End If
    End If
    vsAdvice.SetFocus
End Sub

Private Sub cboִ�п���_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim lng�������� As Long, strҩ��IDs As String
    Dim lngBegin As Long, lngEnd As Long, blnNode As Boolean, bln��Ժ As Boolean, bln���� As Boolean
        
    If cboִ�п���.ListIndex = -1 Then Exit Sub
    
    If cboִ�п���.ItemData(cboִ�п���.ListIndex) = -1 Then
        
        blnNode = True
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_���) = "Z" Then   '��Ժ������
            bln��Ժ = vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "2"
            bln���� = vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "1" '����Ϊ�����סԺ����
            blnNode = Not (bln��Ժ Or bln����)
        End If
    
        '����ִ�У�����ѡ��ִ�п���
        strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And B.������� IN(" & IIF(bln��Ժ, "2", IIF(bln����, "1,2", "1")) & ",3)" & _
            IIF(blnNode, " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
            IIF(bln��Ժ Or bln����, " And B.��������='�ٴ�'", "") & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by A.����"
        vRect = GetControlRect(cboִ�п���.hWnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lblִ�п���.Caption, , , , , , True, vRect.Left, vRect.Top, cboִ�п���.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then
                cboִ�п���.ListIndex = intIdx
            Else
                cboִ�п���.AddItem rsTmp!���� & "-" & rsTmp!����, cboִ�п���.ListCount - 1
                cboִ�п���.ItemData(cboִ�п���.NewIndex) = rsTmp!ID
                cboִ�п���.ListIndex = cboִ�п���.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = SeekCboIndex(cboִ�п���, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
            Call zlControl.CboSetIndex(cboִ�п���.hWnd, intIdx)
        End If
    Else
        lngRow = vsAdvice.Row
        
        '���һ����ҩ����������
        With vsAdvice
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And RowInһ����ҩ(lngRow) Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow, COL_���ID)), lngBegin, lngEnd)
                
                '��ǰ������ͨҩ���������������ĸ�Ϊ��������
                If Have��������(cboִ�п���.ItemData(cboִ�п���.ListIndex), "��������") Then
                    lng�������� = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                End If
                '��ǰ�����������Ļ��Ϊ��ͨҩ��
                If lng�������� = 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                            '�Ա�ҩ������
                            If Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                                If Have��������(Val(.TextMatrix(i, COL_ִ�п���ID)), "��������") Then
                                    lng�������� = Val(.TextMatrix(i, COL_ִ�п���ID)): Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                '�������������ҩƷ��ִ�п�����ͬ�����洢�趨
                If lng�������� <> 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                            '�Ա�ҩ������
                            If Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                                strҩ��IDs = Get����ҩ��IDs(.TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), Val(.TextMatrix(i, COL_�շ�ϸĿID)), mlng���˿���id, 1)
                                If InStr("," & strҩ��IDs & ",", "," & cboִ�п���.ItemData(cboִ�п���.ListIndex) & ",") = 0 Then
                                    MsgBox "һ����ҩ��ҩƷ�У�""" & .TextMatrix(i, col_ҽ������) & """��""" & NeedName(cboִ�п���.Text) & """��û�д洢��", vbInformation, gstrSysName
                                    '�ָ������еĿ���(������Click)
                                    intIdx = SeekCboIndex(cboִ�п���, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
                                    Call zlControl.CboSetIndex(cboִ�п���.hWnd, intIdx)
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End With
        
        cboִ�п���.Tag = "1"
        
        '���¸����˵�ִ�п���ҽ������
        Call AdviceChange
        
        '���»�ȡ��沢��ʾ�������ﵥλ����ҩ�䷽����ʾ
        With vsAdvice
            If (.TextMatrix(lngRow, COL_���) = "4" And Val(.TextMatrix(lngRow, COL_��������)) = 1 _
                Or InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0) And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                Call GetDrugStock(lngRow)
                If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                    stbThis.Panels(3).Text = IIF(Val(.TextMatrix(lngRow, COL_���)) > 0, "�п��", "�޿��")
                Else
                    stbThis.Panels(3).Text = "���: " & FormatEx(Val(.TextMatrix(lngRow, COL_���)), 5) & .TextMatrix(lngRow, COL_���ﵥλ)
                End If
            ElseIf RowIn�䷽��(lngRow) Then
                Call GetDrugStock(lngRow)
            End If
        End With
    End If
End Sub

Private Sub cboִ�п���_GotFocus()
    Call zlControl.TxtSelAll(cboִ�п���)
End Sub

Private Sub cboִ�п���_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboִ�п���.ListIndex = -1 Then
            Call cboִ�п���_Validate(blnCancel)
            cboִ�п���.SetFocus
        Else
            If SeekNextControl Then Call cboִ�п���_Validate(False)
        End If
    End If
End Sub

Private Sub cboִ�п���_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String, strIDs As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim blnNode As Boolean, bln�������� As Boolean, bln��Ժ As Boolean
    
    If cboִ�п���.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cboִ�п���.Text = "" Then '������
        If cboִ�п���.ListCount > 0 Then Cancel = True
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '�Ƿ���������ѡ�����
    blnLimit = True
    If cboִ�п���.ListCount > 0 Then
        If cboִ�п���.ItemData(cboִ�п���.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    blnNode = True
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_���) = "Z" Then
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "1" Then
            blnNode = False
            bln�������� = True
        ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "2" Then
            blnNode = False
            bln��Ժ = True
        End If
    End If
    
    strInput = UCase(NeedName(cboִ�п���.Text))
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(" & IIF(bln��������, "1,2", IIF(bln��Ժ, "2", "1")) & ",3)" & _
        IIF(blnNode, " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (A.���� Like [1] Or A.���� Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then strIDs = strIDs & "," & rsTmp!ID
            rsTmp.MoveNext
        Next
        
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
            If InStr(strIDs, ",") = 0 Then
                intIdx = SeekCboIndex(cboִ�п���, CLng(strIDs))
                If intIdx <> -1 Then cboִ�п���.ListIndex = intIdx
            Else
                strSQL = "Select /*+ rule*/ A.ID,A.����,A.����,A.���� From ���ű� A,Table(f_num2list([1])) B Where A.ID = B.Column_Value"
        
                vRect = GetControlRect(cboִ�п���.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lblִ�п���.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, strIDs)
                If Not rsTmp Is Nothing Then
                    intIdx = SeekCboIndex(cboִ�п���, rsTmp!ID)
                    If intIdx <> -1 Then cboִ�п���.ListIndex = intIdx
                End If
            End If
        End If
        If cboִ�п���.ListIndex = -1 Then
            MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cboִ�п���.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lblִ�п���.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then
                cboִ�п���.ListIndex = intIdx
            Else
                cboִ�п���.AddItem rsTmp!���� & "-" & rsTmp!����, cboִ�п���.ListCount - 1
                cboִ�п���.ItemData(cboִ�п���.NewIndex) = rsTmp!ID
                cboִ�п���.ListIndex = cboִ�п���.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboִ��ʱ��_Change()
    cboִ��ʱ��.Tag = "1"
End Sub

Private Sub cboִ��ʱ��_LostFocus()
    If Not mblnIsInHelp Then picHelp.Visible = False
    mblnIsInHelp = False
End Sub

Private Sub cboִ��ʱ��_Click()
    'cboִ��ʱ��_Change
    '��������
    cboִ��ʱ��.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cboִ��ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cboִ��ʱ��_Validate(False)
    Else
        If InStr("0123456789:-/" & Chr(8) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cboִ��ʱ��_Validate(Cancel As Boolean)
    Dim blnValid As Boolean, lngRow As Long, strTmp As String
    
    lngRow = vsAdvice.Row
        
    With vsAdvice
        If cboִ��ʱ��.Text <> "" Then
            '��鳤��
            If Len(cboִ��ʱ��.Text) > 50 Then
                MsgBox "�������ݲ��ܳ��� 50 ���ַ���", vbInformation, gstrSysName
                Call cboִ��ʱ��_GotFocus
                Cancel = True: Exit Sub
            End If
            '���Ϸ���
            If .RowData(lngRow) <> 0 Then
                blnValid = ExeTimeValid(cboִ��ʱ��.Text, Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), .TextMatrix(lngRow, COL_�����λ))
                If Not blnValid Then
                    If .TextMatrix(lngRow, COL_�����λ) = "��" Then
                        strTmp = COL_����ִ��
                    ElseIf .TextMatrix(lngRow, COL_�����λ) = "��" Then
                        strTmp = COL_����ִ��
                    ElseIf .TextMatrix(lngRow, COL_�����λ) = "Сʱ" Then
                        strTmp = COL_��ʱִ��
                    End If
                    MsgBox "�����ִ��ʱ�䷽����ʽ����ȷ�����顣" & vbCrLf & vbCrLf & "����" & vbCrLf & strTmp, vbInformation, gstrSysName
                    Call cboִ��ʱ��_GotFocus
                    Cancel = True: Exit Sub
                End If
            End If
        End If
    End With
    
    '��������
    Call AdviceChange
End Sub

Private Sub cboִ������_Click()
    cboִ������.Tag = "1"
    '��������
    Call AdviceChange
End Sub

Private Sub cboִ������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboִ������.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cboִ������.hWnd, KeyAscii)
        If lngIdx = -1 And cboִ������.ListCount > 0 Then lngIdx = 0
        cboִ������.ListIndex = lngIdx
    End If
End Sub

Private Sub chk����_Click()
    If Not mblnDoCheck Then Exit Sub
    
    chk����.Tag = "1"
    '��������
    Call AdviceChange
    
    If txt��ҩ����.Enabled And Trim(txt��ҩ����.Text) = "" Then
        txt��ҩ����.SetFocus
    End If
End Sub

Private Sub chkZeroBilling_click()
    If Not mblnDoCheck Then Exit Sub
    
    chkZeroBilling.Tag = "1"
    '��������
    Call AdviceChange
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmdExt_Click()
'���ܣ��޸�����ҽ������������
    Dim rsCurr As New ADODB.Recordset
    Dim strExtData As String, strAppend As String
    Dim lngRow As Long, lngFirstRow As Long
    Dim lng������ĿID As Long, lng�÷�ID As Long
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim strTmp As String, lngDiag As Long
    Dim lng�䷽ID As Long
    Dim t_Pati As TYPE_PatiInfoEx
    Dim lng��Ŀid As Long, intType As Integer
    Dim blnOK As Boolean
    Dim str������λ As String
    Dim strIDs1 As String, strIDs2 As String, strҽ������ As String
    Dim lngAppType As Long '���뵥Ӧ��
    Dim objAppPages()  As clsApplicationData
    Dim rsCard As ADODB.Recordset
    Dim lngNo As Long
    Dim lngTmp As Long
    Dim strժҪ As String 'ҽ��ժҪ GetItemInfo
        Dim strSQL As String, rsTmp As Recordset
    
    lngRow = vsAdvice.Row
    '��ȡ���븽�����ݣ�������¼ҽ������¼�롢�����ס�����ʱ�Ѷ�ȡ
    If vsAdvice.TextMatrix(lngRow, COL_����) = "" And Val(vsAdvice.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
        If Not RowIn�䷽��(lngRow) Then
            vsAdvice.TextMatrix(lngRow, COL_����) = Get����ҽ������(vsAdvice.RowData(vsAdvice.Row))
        End If
    End If
    strAppend = vsAdvice.TextMatrix(lngRow, COL_����)
    
    lngNo = Val(vsAdvice.TextMatrix(lngRow, COL_�������) & "")
    intType = -1
    lngAppType = -1
        If lngNo <> 0 Then
        strSQL = "Select �ļ�ID From ҽ�����뵥�ļ� Where ҽ��ID=[1] And RowNum<2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(vsAdvice.TextMatrix(lngRow, COL_���ID)) = 0, Val(vsAdvice.RowData(lngRow)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID))))
        If rsTmp.RecordCount > 0 Then Call FuncApplyCustom(1, Val(rsTmp!�ļ�ID), lngNo): Exit Sub
    End If
    If vsAdvice.TextMatrix(lngRow, COL_���) = "D" Then
        If lngNo <> 0 Then
            Call GetData�������(lngRow, objAppPages())
            lngAppType = 0
        Else
            strExtData = Get��鲿λ����(lngRow)
            If strExtData = "" Then
                MsgBox "�ü��ҽ����ϵͳ������ǰ�´�ģ������з�ʽ�����ݡ��������´�ü��ҽ����", vbInformation, gstrSysName
                Exit Sub
            End If
            intType = 0
        End If
    ElseIf vsAdvice.TextMatrix(lngRow, COL_���) = "F" Then
        strExtData = Get��������IDs(lngRow)
        intType = 1
    ElseIf RowIn�䷽��(lngRow) Then
        strExtData = Get��ҩ�䷽IDs(lngRow)
        intType = 2
    ElseIf RowIn������(lngRow) Then
        If lngNo <> 0 Then
            lngAppType = 3
            Call GetData��������(lngRow, rsCard)
        Else
            strExtData = Get�������IDs(lngRow)
            intType = 4
        End If
    ElseIf vsAdvice.TextMatrix(lngRow, COL_���) = "E" Or vsAdvice.TextMatrix(lngRow, COL_���) = "K" Or vsAdvice.TextMatrix(lngRow, COL_���) = "Z" Then
        If CanUseApply(vsAdvice.TextMatrix(lngRow, COL_���)) Then
            Call GetData��Ѫ����(lngRow, rsCard)
            lngAppType = 1
        Else
            intType = 5
            If vsAdvice.TextMatrix(lngRow, COL_���) = "K" And Val(vsAdvice.TextMatrix(lngRow, COL_�������) & "") <> 0 Then
                Call frmBloodApply.ShowMe(Me, mlng����ID, 0, 1, 3, Val(vsAdvice.RowData(lngRow)), mlng���˿���id, , Val(vsAdvice.TextMatrix(lngRow, COL_��������ID)), , , mrsDefine, mclsMipModule, 1, mstr�Һŵ�)
                Exit Sub
            End If
        End If
    Else
        Exit Sub '������ǰ�ļ�����Ŀ
    End If
    
    If intType = 4 Then
        lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
        lng��Ŀid = Val(vsAdvice.TextMatrix(lngFirstRow, COL_������ĿID))
    Else
        lng��Ŀid = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
    End If
    With t_Pati
        .blnҽ�� = InStr(",1,2,", mstr������) > 0 And mstr������ <> ""
        .int���� = mint����
        .intӤ�� = mintӤ��
        .lng����ID = mlng����ID
        .lng���˿���ID = mlng���˿���id
        .str�Һŵ� = mstr�Һŵ�
        .str�Ա� = mstr�Ա�
    End With

    On Error Resume Next
    '����ӿڣ���ǰint���ϴ�δ�������ڴ�0��bytUseType��ǰδ�������ڴ�0
    If intType = 2 Then
        blnOK = frmAdviceFormula.ShowMe(Me, gclsInsure, txtҽ������.hWnd, t_Pati, 0, 0, 1, 1, 1, _
                    lng��Ŀid, strExtData, strժҪ)
    ElseIf intType <> -1 Then
        blnOK = frmAdviceEditEx.ShowMe(Me, txtҽ������.hWnd, t_Pati, 0, intType, 0, 1, 1, 1, mblnNewLIS, False, _
                    lng��Ŀid, strExtData, strAppend, , GetAdviceDiagnosis, str������λ)
    End If
    '���뵥
    If lngAppType = 0 Then
        blnOK = ApplyNew�������(1, "", objAppPages())
    ElseIf lngAppType = 1 Then
        blnOK = ApplyNew��Ѫ����(1, "", rsCard)
    ElseIf lngAppType = 3 Then
        blnOK = ApplyNew��������(1, "", rsCard)
    End If
    On Error GoTo 0
    
    '���������������
    If blnOK Then
        '��ȡ��ǰ����Ϲ�����
        lngDiag = AdviceHaveDiag(lngRow)
        '���¿���ʱ��
        vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        vsAdvice.TextMatrix(lngRow, COL_����) = "" '������¼���
        
        If vsAdvice.TextMatrix(lngRow, COL_���) = "D" Then
            If lngAppType = 0 Then
                Call Delete���������Ѫ(lngRow, True, lngTmp)
                lngRow = lngTmp
                Call AdviceSet�������(lngRow, objAppPages())
                strAppend = vsAdvice.TextMatrix(lngRow, COL_����)
            Else
                '������
                Call AdviceSet������(lngRow, strExtData)
                vsAdvice.TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
            End If
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, col_ҽ������)
        ElseIf vsAdvice.TextMatrix(lngRow, COL_���) = "F" Then
            'һ������
            Call AdviceSet�������(lngRow, strExtData)
            vsAdvice.Cell(flexcpData, lngRow, COL_�걾��λ) = str������λ
            vsAdvice.TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, col_ҽ������)
        ElseIf lngAppType = 1 Then
            Call Delete���������Ѫ(lngRow)
            Call DeleteRow(lngRow, True)
            Call AdviceSet��Ѫ����(lngRow, rsCard)
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, col_ҽ������)
            strAppend = ""
        ElseIf lngAppType = 3 Then
            Call Delete�������뵥(lngRow, True, lngTmp)
            lngRow = lngTmp
            Call AdviceSet��������(lngRow, rsCard)
            strAppend = vsAdvice.TextMatrix(lngRow, COL_����)
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, col_ҽ������)
        ElseIf RowIn������(lngRow) Then
            '�������
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
            lng�÷�ID = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
            
            '�Ȼ�ȡ��ǰ�Ѿ����ú�ֵ
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʴ���", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʼ��", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "�����λ", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "����", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "ִ��ʱ��", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "��ʼʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "����ҽ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "��������ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "����ʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "ҽ������", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "��־", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
                        
            '�ɼ�������ִ�п��ҿ����������Ŀ��ͬ
            If Val(vsAdvice.TextMatrix(lngFirstRow, COL_ִ�п���ID)) <> 0 Then
                rsCurr!ִ�п���ID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_ִ�п���ID))
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_����)) <> 0 Then
                rsCurr!���� = Val(vsAdvice.TextMatrix(lngRow, COL_����))
            End If
            rsCurr!ִ��ʱ�� = vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��)
            rsCurr!Ƶ�� = vsAdvice.TextMatrix(lngRow, COL_Ƶ��)
            rsCurr!Ƶ�ʴ��� = Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���))
            rsCurr!Ƶ�ʼ�� = Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��))
            rsCurr!�����λ = vsAdvice.TextMatrix(lngRow, COL_�����λ)
            rsCurr!��ʼʱ�� = vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��)
            rsCurr!����ҽ�� = vsAdvice.TextMatrix(lngRow, COL_����ҽ��)
            rsCurr!��������id = Val(vsAdvice.TextMatrix(lngRow, COL_��������ID))
            rsCurr!����ʱ�� = vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��)
            rsCurr!ҽ������ = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
            rsCurr!��־ = vsAdvice.TextMatrix(lngRow, COL_��־)
            '�޸��˼����������,�ɼ�������Ӧ���Ϊ�޸�
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!ҽ��ID = vsAdvice.RowData(lngRow)
            rsCurr.Update
            
            '��ȫ�������øü������
            '------------------------
            'ɾ��������Ŀ��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete�������(lngRow)
            '�����ǰ��(�ɼ�������)
            Call DeleteRow(lngRow, True, False)
            '���²���:����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet�������(lngRow, lng�÷�ID, strExtData, rsCurr)
        ElseIf RowIn�䷽��(lngRow) Then
            '��ҩ�䷽
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
            lng������ĿID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_������ĿID))
            lng�÷�ID = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
            
            '�Ȼ�ȡ��ǰ�Ѿ����ú�ֵ
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "ִ������", adVarChar, 10, adFldIsNullable
            rsCurr.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʴ���", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʼ��", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "�����λ", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "����", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "ִ��ʱ��", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "��ʼʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "����ҽ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "��������ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "����ʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "ҽ������", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "��־", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
            
            rsCurr!ִ������ = NeedName(cboִ������.Text) '����,�Ա�ҩ,��Ժ��ҩ
            
            'ȡ�䷽����ѡ���ҩ��
            rsCurr!ִ�п���ID = Val(Split(strExtData, "|")(4))
            
            rsCurr!Ƶ�� = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
            rsCurr!Ƶ�ʴ��� = Val(vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���))
            rsCurr!Ƶ�ʼ�� = Val(vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��))
            rsCurr!�����λ = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
            
            'ȡ�䷽����ѡ��ĸ���
            rsCurr!���� = Val(Split(strExtData, "|")(3))
            
            rsCurr!ִ��ʱ�� = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
            rsCurr!��ʼʱ�� = vsAdvice.Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
            rsCurr!����ҽ�� = vsAdvice.TextMatrix(lngFirstRow, COL_����ҽ��)
            rsCurr!��������id = Val(vsAdvice.TextMatrix(lngFirstRow, COL_��������ID))
            rsCurr!����ʱ�� = vsAdvice.Cell(flexcpData, lngFirstRow, COL_����ʱ��)
            rsCurr!ҽ������ = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
            rsCurr!��־ = vsAdvice.TextMatrix(lngRow, COL_��־)
            '�޸����䷽����,�÷���Ӧ���Ϊ�޸�
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!ҽ��ID = vsAdvice.RowData(lngRow)
            
            rsCurr.Update
            
            '��ȫ�������ø���ҩ�䷽��
            '------------------------
            'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete��ҩ�䷽(lngRow)
            '�����ǰ�÷����䷽ID��Ϊ�գ������䷽ID
            lng�䷽ID = Val(vsAdvice.TextMatrix(lngRow, COL_�䷽ID))
            '�����ǰ��(��ҩ�÷���)
            Call DeleteRow(lngRow, True, False)
            '�����䷽:����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet��ҩ�䷽(lng������ĿID, lngRow, lng�÷�ID, strExtData, rsCurr, strժҪ, lng�䷽ID)
        End If
        
        '���¸�������:�Ե�ǰ�ɼ���Ϊ׼
        If strAppend <> "" Then
            vsAdvice.TextMatrix(lngRow, COL_����) = strAppend
            vsAdvice.Cell(flexcpData, lngRow, COL_����) = 1 '������Ҫ����д��(�������޸�)
            Call ReplaceAdviceAppend(lngRow) 'ȱʡ�滻����ҽ�������븽��
        End If
        
        'ǿ����ʾ��ǰҽ����Ƭ
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
                
        Call CalcAdviceMoney '��ʾ�¿�ҽ�����
        
        If InStr(",0,3,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
            vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            Call ReSetColor(lngRow)
        End If
        
        '�޸ĺ������ϵı�Ǵ���
        If lngDiag <> -1 Then
            Call SetDiagFlag(vsAdvice.Row, 1, lngDiag)
        End If
        
        mblnNoSave = True '���Ϊδ����
    End If
    
    Call vsAdvice.AutoSize(col_ҽ������)
    
    '�Ա��ն�����м��
    Call GetInsureStr(strIDs1, strIDs2, strҽ������, vsAdvice.Row)
    strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, 1, strIDs1, strIDs2, strҽ������)
    If strMsg <> "" Then
        If gintҽ������ = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln���Ѷ��� = False
    End If
    
    txtҽ������.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReplaceAdviceAppend(ByVal lngRow As Long)
'���ܣ�����ָ���е����븽�������������������¼ҽ�������븽�����ȱʡ�滻
'������lngRow=����������޸ĵĿɼ�ҽ����
    Dim strAppend As String, i As Long
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_����) = "" Then Exit Sub
        
        For i = .FixedRows To .Rows - 1
            'ֻ�����¼���ҽ�����޸ĵ�ҽ���޸�ʱ�Ѽ��
            '".Cell(flexcpData, i, COL_����) = 1"�Ŀ��ܻ�û������������м�飬ֻ���Զ��滻��
            If .RowData(i) <> 0 And Not .RowHidden(i) And i <> lngRow And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                If .TextMatrix(i, COL_����) <> "" Then
                    strAppend = ReplaceAppend(.TextMatrix(i, COL_����), .TextMatrix(lngRow, COL_����))
                    If .TextMatrix(i, COL_����) <> strAppend Then
                        .TextMatrix(i, COL_����) = strAppend
                        .Cell(flexcpData, i, COL_����) = 1
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub ClinicSelecter(Optional ByVal ���� As Integer, Optional ByVal lng����id As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If ���� = 8 And lng����id <> 0 Then
        'ֱ�Ӷ�ȡѡ��ĳ�����Ŀ
        On Error GoTo errH
        strSQL = "Select A.��� As ���ID,A.ID as ������ĿID,Null as �շ�ϸĿID,B.���� As ���,A.����,A.����,A.���㵥λ,A.�걾��λ,NULL as ��Ŀ����" & _
            " From ������ĿĿ¼ A,������Ŀ��� B Where A.���=B.���� And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id)
    Else
        '��ѡ����������ָ���˳�ʼ����Ŀ¼
        Set rsTmp = frmClinicSelect.ShowSelect(Me, IIF(mlngǰ��ID <> 0, 2, 0), 0, mlng���˿���id, 1, mstr�Ա�, , , 1, lng����id, mint����)
        If rsTmp Is Nothing Then 'ȡ����������
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
    End If
    
    '����ѡ����Ŀ����ȱʡҽ����Ϣ
    If AdviceInput(rsTmp, vsAdvice.Row) Then
        '��ʾ��ȱʡ���õ�ֵ
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        
        Call CalcAdviceMoney '��ʾ�¿�ҽ�����
        
        'ҽ���ܿ�ʵʱ���
        If mint���� <> 0 And Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_EDIT)) = 0 Then
            '�����������룺ȱʡ���̶�������ҽ�����Լ�����
            '����ҽ������������
            If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) And Not txt����.Enabled Then
                If MakePriceRecord(vsAdvice.Row) Then
                    If Not gclsInsure.CheckItem(mint����, 0, 0, mrsPrice) Then
                        Call AdviceCurRowClear: Exit Sub
                    End If
                End If
                '���Ϊ�Ѿ����˼��
                vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_״̬) = 1
            End If
        End If
                 
        txtҽ������.SetFocus: Call SeekNextControl '�����ȶ�λ
    Else
        '�ָ�ԭֵ(AdviceInput�����п��ܴ�����һ��)
        txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
        txtҽ������.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
    Call ClinicSelecter
End Sub

Private Sub cmd��ʼʱ��_Click()
    If IsDate(txt��ʼʱ��.Text) Then
        dtpDate.value = CDate(txt��ʼʱ��.Text)
    Else
        dtpDate.value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "��ʼʱ��"
    dtpDate.Left = txt��ʼʱ��.Left + fraAdvice.Left
    dtpDate.Top = txt��ʼʱ��.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If dtpDate.Tag = "��ʼʱ��" Then
        'ȡֵ
        If IsDate(txt��ʼʱ��.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�ж�ʱ��Ϸ���
        If Not Check��ʼʱ��(strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt��ʼʱ��.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt��ʼʱ��_Validate(False) '��������
        txt��ʼʱ��.SetFocus
    ElseIf dtpDate.Tag = "����ʱ��" Then
        'ȡֵ
        If IsDate(txt����ʱ��.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt����ʱ��.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�ж�ʱ��Ϸ���
        If Not Check����ʱ��(strDate, txt��ʼʱ��.Text, vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt����ʱ��.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt����ʱ��_Validate(False) '��������
        txt����ʱ��.SetFocus
    End If
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call dtpDate_DateClick(dtpDate.value)
    End If
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    dtpDate.Visible = False
    dtpDate.Tag = ""
End Sub

Private Sub Form_Activate()
    If mblnRunFirst Then
        mblnRunFirst = False
        If vsDiag.Rows = 2 And vsDiag.TextMatrix(1, col���) = "" Then
            If vsDiag.Enabled Then
                Call vsDiag_AfterRowColChange(vsDiag.Row, vsDiag.Col, vsDiag.Row, vsDiag.Col)
                vsDiag.SetFocus
            End If
        Else
            If txtҽ������.Enabled Then txtҽ������.SetFocus
        End If
        '��һ�ν���ʱ���������Ƶ���ǰ�У���ΪLoad�в����ƶ�������,�������β���Ч��
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strժҪ As String
    Dim str��ҩIDs As String
    Dim lngBaseRow As Long, i As Long
    Dim lng�շ�ϸĿID As Long, str������ĿID As String
    
    If Shift = vbAltMask Then
        If Between(Chr(KeyCode), "1", "9") And Not mfrmShortCut Is Nothing Then
            Call mfrmShortCut.ShowShortCut(Val(Chr(KeyCode)))
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        '���泣������
        If cmd��������.Enabled And cmd��������.Visible Then
            Call cmd��������_Click
        End If
    ElseIf KeyCode = vbKeyF1 And Shift = vbCtrlMask Then
        '����ҽ����ʾ
        With vsAdvice
            If .RowData(.Row) <> 0 Then
                lng�շ�ϸĿID = Val(.TextMatrix(.Row, COL_�շ�ϸĿID))
                str������ĿID = Val(.TextMatrix(.Row, COL_������ĿID))
                If RowIn�䷽��(.Row) Then
                    '��ȡ��ҩ�䷽��һζ��ҩ��
                    lngBaseRow = .FindRow(CStr(.RowData(.Row)), , COL_���ID)
                    For i = lngBaseRow To .Row
                        If i = lngBaseRow Then lng�շ�ϸĿID = Val(.TextMatrix(i, COL_�շ�ϸĿID))
                        If .TextMatrix(i, COL_���) = "7" Then
                            str��ҩIDs = str��ҩIDs & "," & .TextMatrix(i, COL_������ĿID)
                        End If
                    Next
                    str��ҩIDs = Mid(str��ҩIDs, 2)
                    If UBound(Split(str��ҩIDs, ",")) <> 0 Then
                        lng�շ�ϸĿID = 0
                    End If
                    str������ĿID = str��ҩIDs
                End If
                'ҽ��������������ʱ����ʾ����ҽ������ҲҪ��(Or And mint���� <> 0)
                strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, lng�շ�ϸĿID, CStr(.Cell(flexcpData, .Row, COL_ҽ������)), 0, "", str������ĿID)
                .Cell(flexcpData, .Row, COL_ҽ������) = strժҪ
            End If
        End With
    Else
        Select Case KeyCode
            Case vbKeyEscape
                If dtpDate.Visible Then
                    dtpDate.Visible = False
                    dtpDate.Tag = ""
                End If
            Case vbKeyF4, vbKeyUp, vbKeyDown
                If Me.ActiveControl Is txt��ʼʱ�� Then
                    If cmd��ʼʱ��.Visible And cmd��ʼʱ��.Enabled Then cmd��ʼʱ��_Click
                ElseIf Me.ActiveControl Is txt����ʱ�� Then
                    If cmd����ʱ��.Enabled And cmd����ʱ��.Visible Then cmd����ʱ��_Click
                ElseIf Me.ActiveControl Is txtҽ������ Then
                    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
                ElseIf Me.ActiveControl Is txt�÷� Then
                    If cmd�÷�.Visible And cmd�÷�.Enabled Then cmd�÷�_Click
                ElseIf Me.ActiveControl Is txtƵ�� Then
                    If cmdƵ��.Visible And cmdƵ��.Enabled Then cmdƵ��_Click
                End If
            Case vbKeyF7 '�л����뷨
                If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
                    If stbThis.Panels("WB").Bevel = sbrRaised Then
                        Call stbThis_PanelClick(stbThis.Panels("WB"))
                    Else
                        Call stbThis_PanelClick(stbThis.Panels("PY"))
                    End If
                End If
            Case vbKeyF8 '�л���ʾ�Ƽ���Ŀ
                If stbThis.Panels("Price").Visible Then
                    Call stbThis_PanelClick(stbThis.Panels("Price"))
                End If
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
        If Not mfrmShortCut Is Nothing Then Call mfrmShortCut.ShowMe(Me, mint����, 1, 0, mlng���˿���id)
    ElseIf KeyAscii = vbKeySpace Then
        If Me.ActiveControl Is txt��ʼʱ�� And txt��ʼʱ��.SelLength = Len(txt��ʼʱ��.Text) Then
            KeyAscii = 0
            If cmd��ʼʱ��.Visible And cmd��ʼʱ��.Enabled Then cmd��ʼʱ��_Click
        ElseIf Me.ActiveControl Is txt����ʱ�� And txt����ʱ��.SelLength = Len(txt����ʱ��.Text) Then
            KeyAscii = 0
            If cmd����ʱ��.Enabled And cmd����ʱ��.Visible Then cmd����ʱ��_Click
        ElseIf Me.ActiveControl Is txtҽ������ Then
            KeyAscii = 0
            If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
        ElseIf Me.ActiveControl Is txt�÷� Then
            KeyAscii = 0
            If cmd�÷�.Visible And cmd�÷�.Enabled Then cmd�÷�_Click
        ElseIf Me.ActiveControl Is txtƵ�� Then
            KeyAscii = 0
            If cmdƵ��.Visible And cmdƵ��.Enabled Then cmdƵ��_Click
        ElseIf Me.ActiveControl Is cbo���� Then
            KeyAscii = 0
            If cbo����.Visible And cbo����.Enabled Then zlCommFun.PressKey (vbKeyF4)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lngTmp As Long
    Dim strErr As String
    
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim objForm As Object
    Dim strNames As String
    Dim i As Long
    
    mbln��չҳǩ = False
    
    '����ṩ�Ŀ�Ƭ
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    If Not gobjPlugIn Is Nothing Then
        Set mcolSubForm = New Collection
        On Error Resume Next
        strTmp = gobjPlugIn.GetFormCaption(glngSys, p����ҽ���´�)
        Call zlPlugInErrH(err, "GetFormCaption")
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = 0 To UBound(arrTmp)
                strTmp = arrTmp(i)
                Set objForm = gobjPlugIn.GetForm(glngSys, p����ҽ���´�, strTmp)
                Call zlPlugInErrH(err, "GetForm")
                If Not objForm Is Nothing Then
                    mcolSubForm.Add objForm, "_" & strTmp
                    strNames = strNames & "," & strTmp
                End If
                Set objForm = Nothing
            Next
        End If
        err.Clear: On Error GoTo 0
    End If

    If strNames <> "" Then
        mbln��չҳǩ = True
        strNames = Mid(strNames, 2)
    End If
    
    If mbln��չҳǩ Then
        arrTmp = Split(strNames, ",")
        With tbcSub
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .ClientFrame = xtpTabFrameSingleLine
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
            
            .InsertItem(0, "ҽ���༭", picSub.hWnd, 0).Tag = "ҽ���༭"
            For i = 0 To UBound(arrTmp)
                strTmp = arrTmp(i)
                .InsertItem(i + 1, strTmp, mcolSubForm("_" & strTmp).hWnd, 0).Tag = strTmp
            Next
        End With
    Else
        tbcSub.Visible = False
    End If
    
    Call InitObjLis(p����ҽ��վ)
    If gobjLIS Is Nothing Then
        mblnNewLIS = False
    Else
        On Error Resume Next
        mblnNewLIS = gobjLIS.GetApplicationFormShowType
        err.Clear: On Error GoTo 0
    End If
    Call InitCommandBar
    Call InitAdviceTable
    If mint���� = 0 Then
        '��������
        mbytSize = zlDatabase.GetPara("����", glngSys, p����ҽ��վ, "0")
    ElseIf mint���� = 2 Then
        mbytSize = zlDatabase.GetPara("����", glngSys, pҽ������վ, "0")
    End If
    Call SetFontSize(mbytSize)
    Call RestoreWinState(Me, App.ProductName)
    vsAdvice.ColWidth(0) = 14 * Screen.TwipsPerPixelX

    Call zlControl.CboSetHeight(cbo����, Me.Height)
    Call zlControl.CboSetHeight(cboִ�п���, Me.Height)
    Call zlControl.CboSetWidth(cboִ�п���.hWnd, cboִ�п���.Width * 1.3)
    fraPati.BackColor = Me.BackColor
    fra���.BackColor = Me.BackColor

    mblnOK = False
    mblnNoSave = False
    mblnRowMerge = False
    mblnRunFirst = True
    mblnRowChange = True
    mblnDoCheck = True
    mstrDelIDs = ""
    mstrAduitDelIDs = ""
    mstrDel��Ѫ = ""
    mlngID���� = 0
    mblnAddAgent = zlDatabase.GetPara("Ҫ��ǼǴ�����", glngSys, p����ҽ���´�, "1") = "1"
    mblnFreeInput = Val(zlDatabase.GetPara("��������������ɵ���", glngSys, 0)) = 1
    '����ƥ��
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    '����ƥ�䷽ʽ��0-ƴ��,1-���
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    '������Ļ����
    mblnStaKB = Val(zlDatabase.GetPara("������Ļ����", glngSys, p����ҽ��վ)) <> 0
    If Not mblnStaKB Then
        stbThis.Panels("KB").Visible = False
    End If
    stbThis.Panels("KB").ToolTipText = "���������Ļ����"
    Select Case mint����
    Case 0
        stbThis.Panels("PY").Bevel = sbrInset
        stbThis.Panels("WB").Bevel = sbrRaised
    Case 1
        stbThis.Panels("PY").Bevel = sbrRaised
        stbThis.Panels("WB").Bevel = sbrInset
    Case Else
        stbThis.Panels("PY").Bevel = sbrInset
        stbThis.Panels("WB").Bevel = sbrInset
    End Select
    If Not gbln����ƥ�䷽ʽ�л� Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If

    'PASS�ӿڳ�ʼ��
    Call zlPASSMap
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then        'Pass
            '���˹���ʷ/����״̬���ü��
            Call gobjPass.zlPassCmdAlleyEnable(mobjPassMap)
        End If
    End If
    '���������Դ
    opt���(Val(zlDatabase.GetPara("�����������", glngSys, p����ҽ��վ, 0, Array(opt���(0), opt���(1)), InStr(GetInsidePrivs(p����ҽ���´�), "ҽ��ѡ������") > 0))).value = True
    If gint�����Դ > 1 Then
        opt���(0).Enabled = False
        opt���(1).Enabled = False
        If gint�����Դ = 2 Then
            opt���(0).value = True
        ElseIf gint�����Դ = 3 Then
            opt���(1).value = True
        End If
    End If
    opt���(0).TabStop = False
    opt���(1).TabStop = False

    '�Ƽ����״̬
    If mblnModal Then
        stbThis.Panels("Price").Visible = False
    Else
        Set mfrmPrice = New frmAdvicePrice
        stbThis.Panels("Price").Tag = zlDatabase.GetPara("��ʾҽ���Ƽ����", glngSys, p����ҽ���´�)
    End If

    '����¼��ҩƷ����
    mbln���� = Val(zlDatabase.GetPara("����¼��ҩƷ����", glngSys, p����ҽ���´�)) <> 0

    'ִ������
    mbln���� = Val(zlDatabase.GetPara("ҽ��ִ������", glngSys, p����ҽ���´�)) <> 0
    vsAdvice.ColHidden(COL_����) = Not mbln����
    
    mbln�������� = (gbln�������� And mbln����)

    '����ҩ��ȱʡ��ҩĿ��
    mstrPurMed = zlDatabase.GetPara("����ҩ��ȱʡ��ҩĿ��", glngSys, p����ҽ���´�, "0")

    '�Զ�����Ƥ��
    mbln�Զ�Ƥ�� = Val(zlDatabase.GetPara("�Զ�����Ƥ��", glngSys, p����ҽ���´�)) <> 0 And mlngǰ��ID = 0

    '�Զ��رմ���
    mblnAutoClose = Val(zlDatabase.GetPara("������ɺ�ر�ҽ������", glngSys, p����ҽ���´�)) <> 0

    'ҩƷ�����ĳ����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    With cboDruPur
        .Clear
        .AddItem " "
        .AddItem "Ԥ��"
        .AddItem "����"
    End With
    
    '��������
    Call ReadEnjoin
    'ҽ�����ݶ���
    If CreateScript(mobjVBA, mobjScript) Then
        Set mrsDefine = InitAdviceDefine
    End If
    '--------------------------------------------------
    '��ȡ������Ϣ
    Call GetPatiInfo
    Call SetBabyVisible(mlng���˿���id)
    '��ȡ��������Ϣ
    Call GetAgentInfo

    '�޸�ʱǿ�ж�λӤ��
    If mlngҽ��ID = 0 Then    '����
        cboӤ��.ListIndex = 0    'ȱʡ�������˵�ҽ��
    Else    '�޸�
        cboӤ��.ListIndex = mintӤ��
    End If
    cboӤ��.Tag = cboӤ��.ListIndex

    '��ȡ����ʾ����ҽ��
    Call ReLoadAdvice(mlngҽ��ID)

    '���������봰��
        If mblnModal = False Then
        Set mfrmShortCut = New frmClinicShortCut
        mfrmShortCut.ShowMe Me, mint����, 1, 0, mlng���˿���id, True    '�����ϴ��Ϸ���ʾ
        End If
    vsDiag.AllowUserResizing = flexResizeNone

    If mblnStaKB Then
        On Error Resume Next
        Set mobjKeyBoard = Nothing
        Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
        err.Clear: On Error GoTo 0
        If Not mobjKeyBoard Is Nothing Then
            Call mobjKeyBoard.StartUp
        Else
            stbThis.Panels("KB").Visible = False
            MsgBox "��Ļ���̲���δ����ȷ��װ������ʹ�ã�", vbInformation, gstrSysName
        End If
    End If
    stbThis.Visible = True
End Sub

Private Function TheStockCheck(ByVal lng�ⷿID As Long, ByVal str��� As String) As Integer
'���ܣ���ȡָ���ⷿ�ĳ������鷽ʽ
    Dim intStyle As Integer
    
    On Error Resume Next
    If InStr(",5,6,7,", str���) > 0 Then
        intStyle = mcolStock1("_" & lng�ⷿID)
    ElseIf str��� = "4" Then
        intStyle = mcolStock2("_" & lng�ⷿID)
    End If
    err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Sub ReLoadAdvice(Optional ByVal lngҽ��ID As Long)
'���ܣ����¶�ȡ����ʾ���˵ĵ�ǰҽ���嵥
'������lngҽ��ID=���ڶ�λ
    Dim lngRow As Long
    
    If LoadAdvice Then
        '��ʾ����ҽ����ʶ
        Call ShowDiagFlag(vsDiag.Row)
    
        '��ʾ�ɼ���ҽ��
        Call ShowAdvice
        
        If lngҽ��ID = 0 Then
            If vsAdvice.RowData(vsAdvice.Row) <> 0 Then
                cbsMain.FindControl(, conMenu_New, True, True).Execute
            End If
        Else
            '�޸ĵ�ҽ��IDӦ������ʾ��
            lngRow = vsAdvice.FindRow(lngҽ��ID)
            If lngRow <> -1 Then
                If Not vsAdvice.RowHidden(lngRow) Then
                    mblnRowChange = False
                    vsAdvice.Col = col_ҽ������: vsAdvice.Row = lngRow
                    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
                    mblnRowChange = True
                End If
            End If
        End If
        '����ʱ������ShowAdvice�еĵ���,ǿ�н���
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Function ReadEnjoin() As Boolean
'���ܣ���ȡ�����볣�õ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
        
    On Error GoTo errH
    
    '���õ���
    strPre = cbo����.Text '����󱣳�ԭ��ֵ
    Call Load��Һ����(cbo����, lbl���ٵ�λ, False)
    cbo����.Text = strPre
    
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    If dtpDate.Visible Then
        dtpDate.Visible = False
        dtpDate.Tag = ""
    End If
    
    If mbln��չҳǩ Then
        tbcSub.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Else
        picSub.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Call cbsMain_Resize
    
    'Pass
    cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 2 * Screen.TwipsPerPixelX
    cboӤ��.Left = Me.ScaleWidth - IIF(cmdAlley.Visible, cmdAlley.Width + 30, 0) - cboӤ��.Width - 2 * Screen.TwipsPerPixelX
    lblӤ��.Left = cboӤ��.Left - lblӤ��.Width - 2 * Screen.TwipsPerPixelX
    
    If cmdAlley.Visible Or lblӤ��.Visible Then
        lblPati.Width = IIF(lblӤ��.Visible, lblӤ��.Left, cmdAlley.Left) - lblPati.Left - 6 * Screen.TwipsPerPixelX
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mdat�Һ�ʱ�� = Empty
    msng���� = 0
    
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mrsDefine = Nothing
    Set mrsDrugScale = Nothing
    Set mfrmSend = Nothing
    Set mrsPrice = Nothing
    Set mobjKeyBoard = Nothing
    Set mclsMipModule = Nothing
    '�Ƽ����״̬
    If Not mfrmPrice Is Nothing Then
        Unload mfrmPrice
        Set mfrmPrice = Nothing
        Call zlDatabase.SetPara("��ʾҽ���Ƽ����", IIF(Val(stbThis.Panels("Price").Tag) <> 0, 1, 0), glngSys, p����ҽ���´�, InStr(GetInsidePrivs(p����ҽ���´�), ";ҽ��ѡ������;") > 0)
    End If
    
    If mbln��չҳǩ Then
        For i = 1 To mcolSubForm.Count
            Unload mcolSubForm(i)
        Next
        Set mcolSubForm = Nothing
    End If
    
    If mblnPass Then
        Call gobjPass.zlPassClearLight(mobjPassMap)
        Set mobjPassMap = Nothing
    End If
     
    Call zlDatabase.SetPara("�����������", IIF(opt���(0).value, 0, 1), glngSys, p����ҽ��վ, InStr(GetInsidePrivs(p����ҽ��վ), ";��������;") > 0)
    Call SaveWinState(Me, App.ProductName)
    mblnNoSave = False
    RaiseEvent FormUnload(Cancel)
    Set mrs��� = Nothing
    mlngΣ��ֵID = 0
End Sub

Private Function RowCanMerge(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional strMsg As String) As Boolean
'���ܣ��ж������Ƿ����һ����ҩ
'������lngRow1=ǰ��һ���Ѿ������ҩƷ��
'      lngRow2=��ǰ��(�������δ����)
'���أ���������ԣ���strMsg������ʾ��Ϣ
    Dim lngFind As Long, lngRxCount As Long
    Dim lng�������� As Long, strҩ��IDs As String
    
    With vsAdvice
        strMsg = ""
        If Not Between(lngRow1, .FixedRows, .Rows - 1) Then Exit Function
        If Not Between(lngRow2, .FixedRows, .Rows - 1) Then Exit Function
        If .RowHidden(lngRow1) Or .RowHidden(lngRow2) Then Exit Function
        If .RowData(lngRow1) = 0 Then Exit Function
        
        If .RowData(lngRow2) = 0 Then
            '����ȫ��Ϊ��ҩ�������ͬ
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_���)) = 0 Then
                strMsg = "һ����ҩ��ҩƷ���붼Ϊ����ҩ��Ϊ�г�ҩ��"
                Exit Function
            End If
            
            '���ܰ����ѷ��͵�ҽ��
            If Val(.TextMatrix(lngRow1, COL_״̬)) <> 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ����͵�ҽ����"
                Exit Function
            End If
            '���ܰ�����ǩ����ҽ��
            If Val(.TextMatrix(lngRow1, COL_ǩ����)) = 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ�ǩ����ҽ����"
                Exit Function
            End If
        ElseIf .RowData(lngRow2) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_���)) = 0 _
                Or InStr(",5,6,", .TextMatrix(lngRow2, COL_���)) = 0 Then
                strMsg = "һ����ҩ��ҩƷ���붼Ϊ����ҩ��Ϊ�г�ҩ��"
                Exit Function
            End If
            
            '���ܰ����ѷ��͵�ҽ��
            If Val(.TextMatrix(lngRow1, COL_״̬)) <> 1 Or Val(.TextMatrix(lngRow2, COL_״̬)) <> 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ����͵�ҽ����"
                Exit Function
            End If
            '���ܰ�����ǩ����ҽ��
            If Val(.TextMatrix(lngRow1, COL_ǩ����)) = 1 Or Val(.TextMatrix(lngRow2, COL_ǩ����)) = 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ�ǩ����ҽ����"
                Exit Function
            End If
            
            'һ����ҩ(ǰ��ҩƷ)�ĸ�ҩ;���Ƿ������ڵ�ǰҩƷ
            lngFind = .FindRow(CLng(.TextMatrix(lngRow1, COL_���ID)), lngRow1 + 1)
            If lngFind <> -1 Then
                If Not Check�����÷�(Val(.TextMatrix(lngFind, COL_������ĿID)), Val(.TextMatrix(lngRow2, COL_������ĿID)), 1) Then
                    strMsg = """" & .TextMatrix(lngRow2, col_ҽ������) & """����ʹ��""" & .TextMatrix(lngFind, col_ҽ������) & """��ҩ;����" & _
                    vbCrLf & "������""" & .TextMatrix(lngRow1, col_ҽ������) & """����Ϊһ����ҩ��"
                    Exit Function
                End If
            End If
            
            '���������������ģ��Ƿ񶼿��Դ洢���Ա�ҩ������
            If Not (Val(.TextMatrix(lngRow1, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow1, COL_ִ������)) = 5) Then
                If Have��������(Val(.TextMatrix(lngRow1, COL_ִ�п���ID)), "��������") Then
                    lng�������� = Val(.TextMatrix(lngRow1, COL_ִ�п���ID))
                End If
            End If
            If lng�������� = 0 Then
                If Not (Val(.TextMatrix(lngRow2, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow2, COL_ִ������)) = 5) Then
                    If Have��������(Val(.TextMatrix(lngRow2, COL_ִ�п���ID)), "��������") Then
                        lng�������� = Val(.TextMatrix(lngRow2, COL_ִ�п���ID))
                    End If
                End If
            End If
            If lng�������� <> 0 Then
                If Not (Val(.TextMatrix(lngRow1, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow1, COL_ִ������)) = 5) Then
                    strҩ��IDs = Get����ҩ��IDs(.TextMatrix(lngRow1, COL_���), Val(.TextMatrix(lngRow1, COL_������ĿID)), Val(.TextMatrix(lngRow1, COL_�շ�ϸĿID)), mlng���˿���id, 1)
                    If InStr("," & strҩ��IDs & ",", "," & lng�������� & ",") = 0 Then
                        strMsg = "ҩƷ""" & .TextMatrix(lngRow1, col_ҽ������) & """����������""" & Get��������(lng��������) & """û�д洢��"
                        Exit Function
                    End If
                End If
                If Not (Val(.TextMatrix(lngRow2, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow2, COL_ִ������)) = 5) Then
                    strҩ��IDs = Get����ҩ��IDs(.TextMatrix(lngRow2, COL_���), Val(.TextMatrix(lngRow2, COL_������ĿID)), Val(.TextMatrix(lngRow2, COL_�շ�ϸĿID)), mlng���˿���id, 1)
                    If InStr("," & strҩ��IDs & ",", "," & lng�������� & ",") = 0 Then
                        strMsg = "ҩƷ""" & .TextMatrix(lngRow2, col_ҽ������) & """����������""" & Get��������(lng��������) & """û�д洢��"
                        Exit Function
                    End If
                End If
            End If
        End If
        
        '��鴦��ҩƷ��������
        If gintRXCount > 0 Then
            lngFind = .FindRow(.TextMatrix(lngRow1, COL_���ID), , COL_���ID)
            lngRxCount = GetMergeCount(vsAdvice, lngFind, COL_���ID, COL_�շ�ϸĿID)
            If lngRxCount >= gintRXCount Then
                strMsg = "һ����ҩ��ҩƷ���� " & lngRxCount & " ���Ѵﵽ�򳬹�ҩƷ���������������� " & gintRXCount & " �֡�"
                Exit Function
            End If
        End If
    End With
    RowCanMerge = True
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngҽ��ID As Long, lng���ID As Long
    Dim str��� As String, lngTmp As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long, strMsg As String
    Dim lng������ĿID As Long, i As Long, j As Long
    Dim lngDiag As Long, lngSeek As Long
    Dim intLoop As Integer
    Dim blnTag As Boolean
    
    Dim lng����ID As Long, str�Һŵ� As String, blnMoved As Boolean
    
    Call AdviceChange 'ǿ�Ƹ���ҽ������
    
    With vsAdvice
        Select Case Control.ID
            Case conMenu_New
                If .RowData(.Row) = 0 Then
'                    If .Row <> .Rows - 1 Then
'                        MsgBox "��ǰ�������ݣ������ڵ�ǰ��¼����Чҽ����ɾ����ǰ�С�", vbInformation, gstrSysName
'                    Else
'                        MsgBox "��ǰ�������ݣ������ڵ�ǰ��¼����Чҽ����", vbInformation, gstrSysName
'                    End If
'                    Exit Sub
                ElseIf .RowData(.Rows - 1) = 0 Then
                    .Row = .Rows - 1
                Else
                    '��ɾ���м����Ŀ���
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                End If
                Call .ShowCell(.Row, .Col)
                If Visible And txtҽ������.Enabled Then txtҽ������.SetFocus
            Case conMenu_Insert
                If .RowData(.Row) = 0 Then
                    MsgBox "��ǰ�������ݣ������ڵ�ǰ��¼����Чҽ����", vbInformation, gstrSysName
                    Exit Sub
                End If
                            
                lngPreRow = GetPreRow(.Row)
                            
                '�������Զ���Ϊһ����ҩ:������һ����ҩ���м����
                If lngPreRow <> -1 Then
                    If Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) _
                        And Val(.TextMatrix(lngPreRow, COL_���ID)) <> 0 And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                        
                        '�������ѷ��͵�һ����ҩ�в���
                        If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then
                            MsgBox "����һ����ҩ��ҽ���Ѿ����ͣ������ٲ��롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '��������ǩ����һ����ҩ�в���
                        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
                            MsgBox "����һ����ҩ��ҽ���Ѿ�ǩ���������ٲ��롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        lng���ID = Val(.TextMatrix(lngPreRow, COL_���ID))
                    End If
                End If
                
                '��ɾ���м����Ŀ���
                mblnRowChange = False
                lngҽ��ID = .RowData(.Row)
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                .Row = .FindRow(lngҽ��ID)
                mblnRowChange = True
                            
                '��ǰ��֮ǰ��������
                '--------------------------------------------------------------
                If RowIn�䷽��(.Row) Or RowIn������(.Row) Then
                    '��ҩ�䷽�������������ǰ���������
                    lngBegin = .FindRow(CStr(.RowData(.Row)), , COL_���ID)
                Else
                    lngBegin = .Row
                End If
                
                mblnRowChange = False
                .AddItem "", lngBegin
                .Row = lngBegin
                .Col = .FixedCols
                mblnRowChange = True
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
                Call .ShowCell(.Row, .Col)
                
                txtҽ������.SetFocus '�ȶ�λ�������
                        Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
                Call FuncViewDrugExplain(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�շ�ϸĿID)), Me)
            Case conMenu_Merge 'һ����ҩ
                If Not Control.Checked Then '�밴��
                    lngBegin = GetPreRow(.Row)
                    'ǰ��û����
                    If lngBegin = -1 Then
                        MsgBox "ǰ��û�п���һ����ҩ��ҽ���С�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '���в���������
                    If Not RowCanMerge(lngBegin, .Row, strMsg) Then
                        MsgBox strMsg, vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If .RowData(.Row) = 0 Then
                        '��ǰ����δ�������ݵ����
                        If DateDiff("n", CDate(.Cell(flexcpData, lngBegin, COL_��ʼʱ��)), zlDatabase.Currentdate) <= gint�����¿�ҽ����� Then
                            txt��ʼʱ��.Text = .Cell(flexcpData, lngBegin, COL_��ʼʱ��)
                        End If
                        mblnRowMerge = True: cbsMain.RecalcLayout '*������
                        txtҽ������.SetFocus: Exit Sub
                    Else
                        'Ҫ�ѵ�ǰ����ǰ����һ��һ����ҩ
                        Call MergeRow(lngBegin, .Row, False)
                        Call ReSetColor(.Row) 'һ��֮����һ������
                    End If
                Else '�뵯��
                    If .RowData(.Row) = 0 Then
                        '��ǰ����δ�������ݵ����
                        If Not RowInһ����ҩ(.Row) Then
                            mblnRowMerge = False '*������
                            cbsMain.RecalcLayout
                        End If
                        Exit Sub
                    Else
                        '��ǰ����һ����ҩ�е���
                        Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
                                                
                        '���жϿɷ�ȡ��һ����ҩ
                        '���ܰ����ѷ��͵�ҽ��
                        If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then
                            MsgBox "��ǰҽ���Ѿ����͡�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '���ܰ�����ǩ����ҽ��
                        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
                            MsgBox "��ǰҽ���Ѿ�ǩ����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                                                
                        '����ʾ
                        If Not (.Row = lngEnd And lngEnd - lngBegin > 1) Then
                            '����һ����ҩȡ��Ϊ������ҩ
                            If MsgBox("Ҫ������һ����ҩ��ҩƷȫ��ȡ��Ϊ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        'ɾ���м�Ŀ���
                        lngTmp = .RowData(.Row)
                        For i = lngEnd To lngBegin Step -1
                            If .RowData(i) = 0 Then
                                .RemoveItem i
                                lngEnd = lngEnd - 1
                            End If
                        Next
                        .Row = .FindRow(lngTmp, lngBegin)
                        
                        '��¼��ǰһ��ʱ����Ϲ����Ա�ָ�
                        lngDiag = AdviceHaveDiag(.Row)
                        lngSeek = .RowData(lngEnd)
                        
                        If .Row = lngEnd And lngEnd - lngBegin > 1 Then
                            '��һ����ҩ�з������
                            Call ReSetColor(.Row) '��ȡ��֮ǰһ������
                            Call SplitRow(.Row)
                        Else
                            'ȡ��һ����ҩ
                            Call ReSetColor(.Row) '��ȡ��֮ǰһ������
                            lngTmp = .RowData(.Row) '��¼���ڻָ��ж�λ
                            Call AdviceSet������ҩ(lngBegin, lngEnd)
                            .Row = .FindRow(lngTmp)
                        End If
                        
                        '���ݼ�¼����Ϲ������лָ�
                        If lngDiag <> -1 Then
                            lngSeek = .FindRow(lngSeek, lngBegin)
                            For i = lngBegin To lngSeek
                                If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 _
                                    And Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) _
                                    And Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                                    Call SetDiagFlag(i, 1, lngDiag)
                                End If
                            Next
                        End If
                    End If
                End If
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
            Case conMenu_Delete
                If .RowSel <> .Row Then
                    MsgBox "һ��ֻ��ɾ��һ��ҽ������ѡ��Ҫɾ����ҽ���С�", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .RowData(.Row) <> 0 Then
                    '�ѷ��͵�ҽ������ɾ��
                    If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then
                        MsgBox "����ҽ���Ѿ����ͣ�����ɾ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '��ǩ����ҽ������ɾ��
                    If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
                        MsgBox "����ҽ���Ѿ�ǩ��������ɾ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                Call AdviceDelete(.Row) 'ɾ����ǰ��
                Call CalcAdviceMoney '��ʾ�¿�ҽ�����
                
                vsAdvice.SetFocus
            Case conMenu_Edit_PacsApply, conMenu_Edit_PacsApply * 10# + 1 '�������
                Call AdviceInput���뵥(1)
            Case conMenu_Edit_LISApply, conMenu_Edit_LISApply * 10# + 1 '��������
                Call AdviceInput���뵥(2)
            Case conMenu_Edit_BloodApply, conMenu_Edit_BloodApply * 10# + 1 '��Ѫ����
                Call AdviceInput���뵥(3)
                        Case conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101# '�Զ������뵥
                FuncApplyCustom 0, Control.Parameter
            Case conMenu_AdvicePay
                Call FuncClinicPay(Me, mlng����ID, mstr�Һŵ�)
            Case conMenu_Reference
                If Val(.TextMatrix(.Row, COL_������ĿID)) <> 0 Then
                    If RowIn�䷽��(.Row) Or RowIn������(.Row) Then
                        i = .FindRow(CStr(.RowData(.Row)), , COL_���ID)
                        If i <> -1 Then
                            lng������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                        End If
                    Else
                        lng������ĿID = Val(.TextMatrix(.Row, COL_������ĿID))
                    End If
                End If
                '���ƴ�ʩ�ο�
                Call frmClinicHelp.ShowMe(IIF(mblnModal, 1, 0), mfrmParent, lng������ĿID)
            Case conMenu_Copy
                lng����ID = mlng����ID: str�Һŵ� = mstr�Һŵ�: blnMoved = False
                strMsg = frmAdviceCopy.ShowMe(Me, mMainPrivs, lng����ID, str�Һŵ�, blnMoved, False, mlngǰ��ID, , mlng���˿���id, , , mstr�Ա�)
                If strMsg <> "" Then
                    cbsMain.FindControl(, conMenu_New, True, True).Execute
                    Call AdviceSet����ҽ��(lng����ID, str�Һŵ�, strMsg, blnMoved)
                End If
            Case conMenu_Scheme, conMenu_Scheme * 10# + 1
                Call frmAdviceScheme.ShowMe(IIF(mlngǰ��ID <> 0, 2, 0), 1, mlng����ID, 0, mstr�Һŵ�, cboӤ��.ListIndex, Me)
            Case conMenu_Scheme * 10# + 2
                Call mfrmShortCut.ShowMe(Me, mint����, 1, 0, mlng���˿���id)
            Case conMenu_Agent
                
                For intLoop = 1 To vsAdvice.Rows - 1
                    If (Val(vsAdvice.TextMatrix(intLoop, COL_״̬)) = 1 And InStr(",����ҩ,����ҩ,����I��,", "," & Trim(vsAdvice.TextMatrix(intLoop, COL_�������)) & ",") > 0) Then
                        blnTag = True: Exit For
                    End If
                Next
                If Not blnTag Then MsgBox "����δִ�е�ҽ���в���������ҩ������ҩ������I��ҩƷҽ��������Ҫ��д��������Ϣ��", vbInformation, gstrSysName: Exit Sub
                Call GetPatiInfo
                Call frmAgentInfo.ShowMe(Me, mlng����ID, mlng�Һ�ID, mstr����, mstr���֤��, AgentInfo.����������, AgentInfo.���������֤��)
                Call GetAgentInfo
                Screen.MousePointer = 0
            Case conMenu_Save
                If vsDiag.EditText <> "" Then
                    mblnCancle = False
                    Me.SetFocus
                    If mblnCancle = True Then mblnCancle = False: Exit Sub
                End If
                If Not CheckAdvice Then Exit Sub '����д����˹�궨λ
                If Not SaveAdvice Then .SetFocus: Exit Sub
            Case conMenu_Send, conMenu_Send * 10# + 1, conMenu_Send * 10# + 2
                '����֮ǰ�Զ�����
                If mblnNoSave Then
                    If Not CheckAdvice Then Exit Sub
                    If Not SaveAdvice Then .SetFocus: Exit Sub
                End If
                If mfrmSend Is Nothing Then Set mfrmSend = New frmOutAdviceSend
                If mfrmSend.ShowMe(Me, mMainPrivs, mlng����ID, mstr�Һŵ�, mstrǰ��IDs, _
                    Control.ID = conMenu_Send And Control.Type = xtpControlSplitButtonPopup Or Control.ID = conMenu_Send * 10# + 1, mlngҽ������ID, mint����, mclsMipModule) Then
                    '����ҽ��������ɺ��Զ��رմ���
                    If mblnAutoClose Then
                        If Not ExistNoSendAdvice(mlng����ID, mstr�Һŵ�) Then
                            mblnOK = True: Unload Me: Exit Sub
                        End If
                    End If
                    
                    '���¶�ȡ��ʾҽ��
                    Call ReLoadAdvice
                    mblnOK = True 'ǿ��
                    If txtҽ������.Enabled Then
                        txtҽ������.SetFocus
                    Else
                        .SetFocus
                    End If
                End If
            Case conMenu_Sign
                Call AdviceSign
            Case conMenu_Help
                ShowHelp App.ProductName, Me.hWnd, Me.Name
            Case conMenu_Exit
                Unload Me
            Case conMenu_DrugScale * 100# + 1 To conMenu_DrugScale * 100# + 99
                With vsAdvice
                    txt����.Text = FormatEx(Val(.TextMatrix(.Row, COL_����ϵ��)) * Val(Control.Parameter), 5)
                    Call zlControl.TxtSelAll(txt����)
                End With
            Case conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99  'PASS������ҩ
                If mblnPass Then
                    Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#, mblnNoSave)
                End If
            Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
                Call ExePlugIn(Control.Parameter)
        End Select
    End With
End Sub

Private Function ExistNoSendAdvice(ByVal lng����ID As Long, ByVal str�Һŵ� As String) As Boolean
'���ܣ���鵱ǰ���˵�ҽ���Ƿ��Ѿ�������ɡ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '����������Ƥ�ԡ�����ȼ�������
    strSQL = "Select 1 From ����ҽ����¼ A,������ĿĿ¼ B" & _
        " Where Nvl(A.Ƥ�Խ��,'��')<>'����' And Not (A.�������='H' And B.��������='1')" & _
        " And Nvl(A.ִ������,0)<>0 And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & _
        " And A.������ĿID=B.ID And A.ҽ��״̬=1 And A.����ID=[1] And A.�Һŵ�=[2] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistNoSendAdvice", lng����ID, str�Һŵ�)
    If Not rsTmp.EOF Then ExistNoSendAdvice = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Getһ����ҩ��Χ(ByVal lng���ID As Long, lngBegin As Long, lngEnd As Long)
'���ܣ�������صĸ�ҩ;��ҽ��ID,ȷ��һ����ҩ��һ��ҩƷ����ֹ�к�
'˵�����м���ܰ����п���
    Dim i As Long
    lngBegin = vsAdvice.FindRow(CStr(lng���ID), , COL_���ID)
    For i = lngBegin To vsAdvice.Rows - 1
        If Not vsAdvice.RowHidden(i) And vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lng���ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Sub txt����_Change()
    With vsAdvice
        If .RowData(.Row) <> 0 Then
            If Val(.TextMatrix(.Row, COL_����)) <> Val(txt����.Text) Then
                txt����.Tag = "1"
            End If
        Else
            txt����.Tag = "1"
        End If
    End With
    
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean

    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt����.Text) And Val(txt����.Text) > 0 _
            Or txt����.Text = "" And (Not mbln���� Or InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) = 0) Then
            Call txt����_Validate(blnCancel)
            If Not blnCancel Then mblnReturn = True: Call SeekNextControl
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim strMsg As String, dbl���� As Double, sng���� As Single
    Dim lngFind As Long, blnDo As Boolean
    Dim sng���� As Single
    
    With vsAdvice
        If Val(txt����.Text) = 0 Then txt����.Text = ""
        If mblnReturn Then mblnReturn = False: Exit Sub
        If Not IsNumeric(txt����.Text) Then
            If txt����.Text <> "" Then
                Cancel = True: txt����_GotFocus: Exit Sub
            ElseIf txt����.Text = "" And mbln���� And InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) > 0 Then
                Cancel = True: txt����_GotFocus: Exit Sub
            End If
        ElseIf CDbl(txt����.Text) <= 0 Then
            Cancel = True: txt����_GotFocus: Exit Sub
        ElseIf CDbl(txt����.Text) > LONG_MAX Then
            Cancel = True: txt����_GotFocus: Exit Sub
        ElseIf txt����.Text <> "" Then
            '�����Ϸ��Լ��
            If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 And Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) <> 0 Then
                blnDo = Not txt����.Visible '��ҩ����Ϊ����������������ʱ�����
                If blnDo Then
                    lngFind = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))), .Row + 1)
                    If lngFind <> -1 Then blnDo = blnDo And Val(.TextMatrix(lngFind, COL_ִ������)) <> 0
                End If
                If blnDo Then
                    dbl���� = IIF(Val(.TextMatrix(.Row, COL_����)) = 0, 1, Val(.TextMatrix(.Row, COL_����))) * _
                        Val(.TextMatrix(.Row, COL_�����װ)) * Val(.TextMatrix(.Row, COL_����ϵ��)) / Val(txt����.Text)
                    If dbl���� > 200 Then
                        If MsgBox("��ҩƷ��ÿ�� " & FormatEx(txt����.Text, 5) & .TextMatrix(.Row, COL_������λ) & " ʹ�ã�" & _
                            IIF(Val(.TextMatrix(.Row, COL_����)) = 0, "ÿ", Val(.TextMatrix(.Row, COL_����))) & _
                            .TextMatrix(.Row, COL_���ﵥλ) & "����ʹ�� " & FormatEx(dbl����, 5) & " �Ρ�" & _
                            vbCrLf & vbCrLf & "��ȷ�ϵ���������ȷ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Cancel = True: txt����_GotFocus: Exit Sub
                        End If
                    End If
                End If
            End If
            
            txt����.Text = FormatEx(txt����.Text, 5)
            
            '���¼���ҩƷ����(�����뵥��ʱ)
            If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                If mbln���� Then
                    If mbln�������� Then
                        If Val(txt����.Text) <> 0 Then
                            '��ʼ��������
                            '��������ʱ������������������Ƶ�ʼ�������������Ƿ���
                            If Val(txt����.Text) <> 0 And Val(txt����.Text) <> 0 <> 0 _
                                And .TextMatrix(.Row, COL_Ƶ��) <> "" And Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 _
                                And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                                
                                sng���� = CalcȱʡҩƷ����(Val(txt����.Text), Val(txt����.Text), _
                                    Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), .TextMatrix(.Row, COL_�����λ), _
                                    Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_�����װ)), _
                                    Val(.TextMatrix(.Row, COL_�ɷ����)))
                                Call CheckDrugOutOfRange(.Row, sng����)
                            End If
                        End If
                        If sng���� = 0 Then sng���� = 1
                        msngPre���� = sng����
                        txt����.Text = sng����
                    Else
                        sng���� = Val(.TextMatrix(.Row, COL_����))
                        If sng���� = 0 Then sng���� = 1
                        sng���� = Val(txt����.Text)
                        
                        txt����.Text = ReGetҩƷ����(sng����, Val(txt����.Text), sng����, .Row)
                        '��ʽ������Change�¼�
                                               
                        Call txt����_Validate(Cancel)
                        If Cancel Then
                            txt����.Text = sng����
                            Exit Sub
                        End If
                    End If
                Else
                    '��������ʱ������������������Ƶ�ʼ�������������Ƿ���
                    If Val(txt����.Text) <> 0 And Val(txt����.Text) <> 0 <> 0 _
                        And .TextMatrix(.Row, COL_Ƶ��) <> "" And Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 _
                        And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                        
                        sng���� = CalcȱʡҩƷ����(Val(txt����.Text), Val(txt����.Text), _
                            Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), .TextMatrix(.Row, COL_�����λ), _
                            Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_�����װ)), _
                            Val(.TextMatrix(.Row, COL_�ɷ����)))
                            
                        Call CheckDrugOutOfRange(.Row, sng����)
                    End If
                End If
            End If
            
            '��������
            Call AdviceChange
        End If
    End With
End Sub

Private Sub txt��ʼʱ��_Change()
    txt��ʼʱ��.Tag = "1"
End Sub

Private Sub txt��ʼʱ��_GotFocus()
    If txt��ʼʱ��.Text = "" Then txt��ʼʱ��.Text = GetDefaultTime(vsAdvice.Row)
    zlControl.TxtSelAll txt��ʼʱ��
End Sub

Private Sub txt��ʼʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ʼʱ��.Text <> "" Then
            txt��ʼʱ��.Text = GetFullDate(txt��ʼʱ��.Text)
            If SeekNextControl Then Call txt��ʼʱ��_Validate(False)
        End If
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt��ʼʱ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt��ʼʱ��.Locked Then
        glngTXTProc = GetWindowLong(txt��ʼʱ��.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt��ʼʱ��.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt��ʼʱ��_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt��ʼʱ��.Locked Then
        Call SetWindowLong(txt��ʼʱ��.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt��ʼʱ��_Validate(Cancel As Boolean)
    If txt��ʼʱ��.Locked Then Exit Sub
        
    If Not IsDate(txt��ʼʱ��.Text) Then
        If txt��ʼʱ��.Text <> "" Then
            Cancel = True
            txt��ʼʱ��_GotFocus
            Exit Sub
        ElseIf vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If IsDate(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��)) Then
                '�ָ���Ϊ�����
                txt��ʼʱ��.Text = vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��)
            End If
        End If
    Else
        '���ʱ��Ϸ���
        If Not Check��ʼʱ��(txt��ʼʱ��.Text) Then
            Cancel = True
            txt��ʼʱ��_GotFocus
            Exit Sub
        End If
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub cboҽ������_Change()
    cboҽ������.Tag = "1"
End Sub

Private Sub cboҽ������_GotFocus()
    zlControl.TxtSelAll cboҽ������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cboҽ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboҽ������.Text <> "" Then
            If ReasonSelect(cboҽ������.Text, 2) Then Exit Sub
        End If
        If SeekNextControl Then Call cboҽ������_Validate(False)
    End If
End Sub

Private Sub cboҽ������_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cboҽ������.Text) > 100 Then
        MsgBox "�������ݲ������� 50 �����ֻ� 100 ���ַ���", vbInformation, gstrSysName
        cboҽ������_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub txtҽ������_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txtҽ������_GotFocus()
    If txt��ʼʱ��.Text = "" Then txt��ʼʱ��_GotFocus
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

Private Sub txtҽ������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intTmp As Integer
    
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txtҽ������)
    End If
    If KeyCode = vbKeySpace And txtҽ������.Text = "" And gblnOut���� Then
        intTmp = ApplySelect
        If intTmp <> 0 Then Call AdviceInput���뵥(intTmp)
    End If
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As ADODB.Recordset
    Dim str���� As String
    Dim blnBarcode As Boolean '�Ƿ��ÿ����ĵ� ���� �ֶ�

    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtҽ������.Text = "" Then Exit Sub
        If txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������) Then
            Call SeekNextControl
            Exit Sub
        End If
        
        Set rsTmp = frmClinicSelect.ShowSelect(Me, IIF(mlngǰ��ID <> 0, 2, 0), 0, mlng���˿���id, 1, mstr�Ա�, txtҽ������.Text, txtҽ������, 1, , mint����)
        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            'txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������)
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
        '����Ŀ��¼��
        '������Ŀ�����������ҩ,���ܰ������ҽ��
        
        If Val(rsTmp!���ID & "") = 4 And rsTmp!���� & "" <> "" Then
            str���� = txtҽ������.Text
            'ʹ������ƥ��Ĺ���1��ȫ���ֻ�������+��ĸ������10λ�����ϣ�
            If (Not zlCommFun.IsCharChinese(str����)) And Len(str����) >= 10 Then
                If InStr("," & rsTmp!����, "," & str����) > 0 Then
                    '��ƥ����Ǳ��룬�ÿ�����
                    blnBarcode = True
                End If
            End If
            If IsNumeric(str����) Then
                '1X.����ȫ������ʱֻƥ�����
                If Mid(gstrMatchMode, 1, 1) = "1" Then
                    If Len(str����) >= 10 Then
                        If InStr("," & rsTmp!����, "," & str����) > 0 Then
                            '��ƥ����Ǳ��룬�ÿ�����
                            blnBarcode = True
                        End If
                    End If
                End If
            ElseIf zlCommFun.IsCharAlpha(str����) Then
                'X1.����ȫ����ĸʱֻƥ�����
                If Mid(gstrMatchMode, 2, 1) = "1" Then
                    If Len(str����) >= 10 Then
                        If InStr("," & rsTmp!����, "," & str����) > 0 Then
                            '��ƥ����Ǳ��룬�ÿ�����
                            blnBarcode = True
                        End If
                    End If
                End If
            End If
        End If
   
        If blnBarcode Then
            Set rsTmpOther = zlDatabase.CopyNewRec(rsTmp)
            rsTmpOther!���� = Null
            Set rsTmp = Nothing
            Set rsTmp = rsTmpOther
        End If
        
        '����ѡ����Ŀ����ȱʡҽ����Ϣ
        Me.Refresh
        If AdviceInput(rsTmp, vsAdvice.Row) Then
            '��ʾ��ȱʡ���õ�ֵ
            Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
                        
            Call CalcAdviceMoney '��ʾ�¿�ҽ�����
            
            'ҽ���ܿ�ʵʱ���
            If mint���� <> 0 And Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_EDIT)) = 0 Then
                '�����������룺ȱʡ���̶�������ҽ�����Լ�����
                '����ҽ������������
                If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) And Not txt����.Enabled Then
                    If MakePriceRecord(vsAdvice.Row) Then
                        If Not gclsInsure.CheckItem(mint����, 0, 0, mrsPrice) Then
                            Call AdviceCurRowClear: Exit Sub
                        End If
                    End If
                    '���Ϊ�Ѿ����˼��
                    vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_״̬) = 1
                End If
            End If
            
            Call SeekNextControl
        Else
            '�ָ�ԭֵ
            'txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������)
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
    End If
End Sub

Private Sub AdviceCurRowClear()
'���ܣ������ǰҽ���е����ݣ�����������������ǰ�����е�������״̬
    Dim str��ʼʱ�� As String
    Dim lngPre As Long
    
    LockWindowUpdate Me.hWnd
    
    '��¼֮ǰ��������������
    Call GetRowScope(vsAdvice.Row, lngPre, 0)
    str��ʼʱ�� = txt��ʼʱ��.Text
    
    'ɾ����
    Call AdviceDelete(vsAdvice.Row)
    
    '��ԭλ�ò�������
    mblnRowChange = False
    vsAdvice.AddItem "", lngPre
    vsAdvice.Row = lngPre: vsAdvice.Col = col_ҽ������
    mblnRowChange = True
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    txt��ʼʱ��.Text = str��ʼʱ��
    txtҽ������.SetFocus
    
    LockWindowUpdate 0
End Sub

Private Sub cboִ��ʱ��_GotFocus()
    zlControl.TxtSelAll cboִ��ʱ��
    If vsAdvice.Row < 1 Then Exit Sub
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "��" Or vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "��" Or vsAdvice.TextMatrix(vsAdvice.Row, COL_�����λ) = "Сʱ" Then
        picHelp.Visible = True
    End If
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txtҽ������.Text <> vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������) Then
        txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
    End If
End Sub

Private Sub txt����_Change()
    With vsAdvice
        If .RowData(.Row) <> 0 Then
            If Val(.TextMatrix(.Row, COL_����)) <> Val(txt����.Text) Then
                txt����.Tag = "1"
            End If
        Else
            txt����.Tag = "1"
        End If
    End With
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim strMask As String
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt����.Text) Or (txt����.Text = "" And vsAdvice.TextMatrix(vsAdvice.Row, COL_���) = "K") Then
            Call txt����_Validate(blnCancel)
            If Not blnCancel Then mblnReturn = True: Call SeekNextControl
        End If
    Else
        If RowIn�䷽��(vsAdvice.Row) Then
            strMask = "0123456789" '��ҩ�䷽ֻ����������
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) > 0 Then
            If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") > 0 _
                And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�ɷ����)) = 0 Then
                strMask = "0123456789."
            Else
                strMask = "0123456789"
            End If
        Else
            strMask = "0123456789."
        End If
        If InStr(strMask & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_LostFocus()
    mblnReturn = False
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim bln�䷽�� As Boolean
    Dim strMsg, strTmp As String
    Dim dbl���� As Double, sng���� As Single
    Dim blnOutTotal As Boolean 'ҩƷ����
    Dim blnTmp As Boolean
    Dim blnTag As Boolean
 
    With vsAdvice
        If Val(txt����.Text) = 0 Then txt����.Text = ""
        If mblnReturn Then mblnReturn = False: Exit Sub
        If Not IsNumeric(txt����.Text) Then
            If txt����.Text <> "" Then
                Cancel = True: txt����_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 Then
                '�ָ���Ϊ���������Ѫ����������������
                If .TextMatrix(.Row, COL_���) <> "K" Then
                    If IsNumeric(.TextMatrix(.Row, COL_����)) Then
                        txt����.Text = .TextMatrix(.Row, COL_����)
                    End If
                End If
            End If
        ElseIf CDbl(txt����.Text) <= 0 Then
            Cancel = True: txt����_GotFocus: Exit Sub
        ElseIf CDbl(txt����.Text) > LONG_MAX Then
            Cancel = True: txt����_GotFocus: Exit Sub
        Else
            txt����.Text = FormatEx(txt����.Text, 5)
        End If
        
        bln�䷽�� = RowIn�䷽��(.Row)
        If IsNumeric(txt����.Text) Then
            If bln�䷽�� Then
                txt����.Text = CInt(txt����.Text)
            ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") = 0 Then
                    txt����.Text = IntEx(Val(txt����.Text))
                ElseIf Val(.TextMatrix(.Row, COL_�ɷ����)) <> 0 Then
                    txt����.Text = IntEx(Val(txt����.Text))
                End If
            ElseIf Val(.TextMatrix(.Row, COL_���㷽ʽ)) = 3 Then
                '�ƴ���Ŀ��������Ϊ�������ƴ���Ŀ�����뵥��,��˵�������
                'txt����.Text = IntEx(Val(txt����.Text))
            End If
        End If
        
        '�����������
        If InStr(",4,5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            If .TextMatrix(.Row, COL_Ƶ��) <> "" And Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 _
                And Val(txt����.Text) <> 0 And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                
                sng���� = Val(txt����.Text)
                If sng���� = 0 Then sng���� = 1
                
                dbl���� = FormatEx(CalcȱʡҩƷ����( _
                    Val(txt����.Text), sng����, _
                    Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), _
                    .TextMatrix(.Row, COL_�����λ), .TextMatrix(.Row, COL_ִ��ʱ��), _
                    Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_�����װ)), _
                    Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    
                If Val(txt����.Text) < dbl���� Then
                    If MsgBox(.TextMatrix(.Row, COL_����) & "��ÿ�� " & _
                        txt����.Text & .TextMatrix(.Row, COL_������λ) & "," & _
                        .TextMatrix(.Row, COL_Ƶ��) & IIF(mbln���� And .TextMatrix(.Row, COL_���) <> "4", ",��ҩ " & sng���� & " ��", "") & _
                        "ִ��ʱ,������Ҫ " & FormatEx(dbl����, 5) & .TextMatrix(.Row, COL_������λ) & ",Ҫ������", _
                        vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        '��鴦������
        .TextMatrix(.Row, COL_�Ƿ���) = ""
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 And Val(.TextMatrix(.Row, COL_��������)) <> 0 Then
           dbl���� = Val(txt����.Text) * Val(.TextMatrix(.Row, COL_�����װ)) * Val(.TextMatrix(.Row, COL_����ϵ��))
           If dbl���� > Val(.TextMatrix(.Row, COL_��������)) And Val(.TextMatrix(.Row, COL_��������)) > 0 Then
               .TextMatrix(.Row, COL_�Ƿ���) = "1"
           End If
    
        ElseIf bln�䷽�� Then
            txt����.Tag = "1" '��ҩ�䷽ҩƷ�б����أ�ͨ������AdviceChange������ı��������
            blnTmp = CheckCHLimited(.Row, Val(txt����.Text), blnOutTotal, vsAdvice, COL_���ID, COL_������ĿID, COL_���, COL_����)
            If blnOutTotal Then .TextMatrix(.Row, COL_�Ƿ���) = "1"
            
            'ͬʱ����ҩ����ҩ�Ƿ���
            If Val(txt����.Text) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                .TextMatrix(.Row, COL_�Ƿ���) = "1"
            End If
            
        ElseIf InStr(",5,6,7,", .TextMatrix(.Row, COL_���)) = 0 And Val(.TextMatrix(.Row, COL_��������)) > 0 Then
            If Val(txt����.Text) > Val(.TextMatrix(.Row, COL_��������)) And Val(.TextMatrix(.Row, COL_��������)) > 0 Then
                .TextMatrix(.Row, COL_�Ƿ���) = "1"
            End If
        End If
                
        '�������ʾ
        If gcurMaxMoney > 0 Then
            If .TextMatrix(.Row, COL_����) = "" Then .TextMatrix(.Row, COL_����) = GetItemPrice(.Row) '��""������
            If Val(.TextMatrix(.Row, COL_����)) * Val(txt����.Text) > gcurMaxMoney Then
                If MsgBox("��ǰҽ�� " & txt����.Text & lbl������λ.Caption & " �Ľ��ﵽ�ˣ�" & Format(Val(.TextMatrix(.Row, COL_����)) * Val(txt����.Text), "0.00") & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: txt����_GotFocus: Exit Sub
                End If
            End If
        End If
        
        '����������û�����������ŷ��㣬��ҩ������ʾ������
        If mbln�������� And txt����.Tag <> "" Then
            If .TextMatrix(.Row, COL_Ƶ��) <> "" And Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 _
                And Val(txt����.Text) <> 0 And Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 _
                And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                
                sng���� = CalcȱʡҩƷ����(Val(txt����.Text), Val(txt����.Text), _
                    Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), .TextMatrix(.Row, COL_�����λ), _
                    Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_�����װ)), _
                    Val(.TextMatrix(.Row, COL_�ɷ����)))
                
                Call CheckDrugOutOfRange(.Row, sng����)
                If sng���� = 0 Then sng���� = 1
                msngPre���� = sng����
                If sng���� <> Val(txt����.Text) Then
                    txt����.Text = sng����
                    txt����.Tag = "1"
                    msng���� = sng����
                End If
            End If
        End If
        
        '��������
        blnTag = (txt����.Tag <> "")
        Call AdviceChange
        'ҽ���ܿ�ʵʱ��⣺�״�����(����)���߸���ʱ���
        If mint���� <> 0 And (.Cell(flexcpData, .Row, COL_״̬) = 0 Or blnTag) Then
            If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) Then
                If MakePriceRecord(.Row) Then
                    If Not gclsInsure.CheckItem(mint����, 0, 0, mrsPrice) Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
                '���Ϊ�Ѿ����˼��
                .Cell(flexcpData, .Row, COL_״̬) = 1
            End If
        End If

        Call CalcAdviceMoney '��ʾ�¿�ҽ�����
        
        'ҩƷ�����:ֻ����,�޸��˲�����
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Or bln�䷽�� _
            Or .TextMatrix(.Row, COL_���) = "4" And Val(.TextMatrix(.Row, COL_��������)) = 1 Then
            strMsg = CheckStock(.Row)
            If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
        End If
        
    End With
End Sub
 
Private Sub ClearAdviceCard()
'���ܣ����ҽ����ʾ��Ƭ��ص�����
'������bln��ʼʱ��=�Ƿ������ʼʱ��
    Call SetCardEditable(True)
    
    txt��ʼʱ��.Text = ""
    txt����ʱ��.Text = ""
    txtҽ������.Text = ""
    cboҽ������.Text = ""
    cboִ�п���.Clear
    cbo����ִ��.Clear
    chk����.value = 0
    chk����.Visible = False '��ҽ�����ݺ�ſ���
    cbo����.Text = ""
    txt����˵��.Text = ""
    txt��ҩ����.Text = ""
    
    mblnDoCheck = False
    chk����.value = 0
    chkZeroBilling.value = 0
    mblnDoCheck = True
    
    cmdExt.Enabled = False
    Call SetDayState(-1, -1)
    Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
    Call SetStartTime(True)
    
    stbThis.Panels(3).Text = ""
    stbThis.Panels(4).Text = ""
End Sub

Private Sub SetCardEditable(ByVal Editable As Boolean)
'���ܣ�����ɫ��ʶ��ǰҽ���Ƿ���Ա༭
    Dim obj As Object
    
    For Each obj In Controls
        If InStr("Label;TextBox;ComboBox;CheckBox;OptionButton", TypeName(obj)) > 0 Then
            If Not obj.Container Is Nothing Then
                If obj.Container Is fraAdvice Then
                    If Editable Then
                        obj.ForeColor = Me.ForeColor
                    Else
                        obj.ForeColor = &H808080
                    End If
                End If
            End If
        End If
    Next
    fraAdvice.Enabled = Editable
    cmdSel.Enabled = fraAdvice.Enabled
    cmd��������.Enabled = fraAdvice.Enabled
    cmdҽ������.Enabled = fraAdvice.Enabled
End Sub

Private Function GetƵ�ʷ�Χ(ByVal lngRow As Long) As Integer
    Dim lngFind As Long
    
    With vsAdvice
        If RowIn�䷽��(lngRow) Then
            GetƵ�ʷ�Χ = 2 '��ҽ
        Else
            If RowIn������(lngRow) Then '�Լ�����Ŀ��Ϊ׼
                lngFind = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                If lngFind <> -1 Then lngRow = lngFind
            End If
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Or Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Then
                GetƵ�ʷ�Χ = 1 '��ҩ���ѡƵ�ʵ���Ŀʹ����ҽƵ����Ŀ
            ElseIf Val(.TextMatrix(lngRow, COL_Ƶ������)) = 1 Then
                GetƵ�ʷ�Χ = -1 'һ����
            ElseIf Val(.TextMatrix(lngRow, COL_Ƶ������)) = 2 Then
                GetƵ�ʷ�Χ = -2 '������
            End If
        End If
    End With
End Function

Private Function SeekVisibleRow() As Boolean
'���ܣ���ǰ��Ϊ������ʱ����λ���������Ŀɼ���
    Dim lngRow As Long
    
    With vsAdvice
        If Not .RowHidden(.Row) Then Exit Function
        If InStr(",F,G,C,D,E,", .TextMatrix(.Row, COL_���)) > 0 And Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))))
        ElseIf .TextMatrix(.Row, COL_���) = "7" Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))))
        ElseIf .TextMatrix(.Row, COL_���) = "E" And Val(.TextMatrix(.Row, COL_���ID)) = 0 Then
            lngRow = .Row - 1
        End If
        If lngRow <> -1 Then
            If .RowData(lngRow) <> 0 Then
                .Row = lngRow: SeekVisibleRow = True
            End If
        End If
    End With
End Function

Private Sub FuncApplyCustom(ByVal intType As Integer, ByVal lng�ļ�ID As Long, Optional ByVal lng������� As Long, Optional ByVal lng��Ŀid As Long)
'���ܣ��Զ������뵥
    Dim objApplyCustom As New frmApplyCustom
    Dim lngOutҽ��ID As Long

    If mblnNoSave Then
        If Not CheckAdvice Then Exit Sub
        If Not SaveAdvice Then vsAdvice.SetFocus: Exit Sub
    End If
    
    If objApplyCustom.ShowMe(Me, 1, intType, mlng����ID, mstr�Һŵ�, 1, lng�ļ�ID, lng�������, mlng���˿���id, mlng���˿���id, , mrsDefine, , , 1, mclsMipModule, mlngǰ��ID, , mint����, lngOutҽ��ID, lng��Ŀid) Then
         '���¶�ȡ��ʾҽ��
        Call ReLoadAdvice(lngOutҽ��ID)
        mblnOK = True 'ǿ��
        If txtҽ������.Enabled Then
            txtҽ������.SetFocus
        Else
            vsAdvice.SetFocus
        End If
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'���ܣ����иı�ʱ�����¿�Ƭ����
    Dim rsItem As New ADODB.Recordset
    Dim rsBlood As New ADODB.Recordset
    Dim strSQL As String, lngRow As Long
    Dim lng�÷�ID As Long, blnEditable As Boolean
    Dim lngҩƷID As Long, lngBaseRow As Long '��ҩ�䷽�ĵ�һζ���ҩ��
    Dim dblPrice As Double, strTmp As String, i As Long
    Dim bln��ʾ���� As Boolean, bln��ʾ����ҩ�� As Boolean
    Dim blnEditableTmp As Boolean
    
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_��ʼʱ��)
    End If
    
    If NewRow = OldRow Then Exit Sub
    If Not mblnRowChange Then Exit Sub
    If SeekVisibleRow Then Exit Sub
    'Pass
    If mblnPass And Me.Visible Then
        If NewRow <> OldRow Then
            If gobjPass.zlPassCheck(mobjPassMap) Then
                Call gobjPass.zlPassSetDrug(mobjPassMap)
            End If
        End If
    End If
    
    lngRow = NewRow
     '��ǰ���ǿ���ʱ�����ǰһ����һ����ҩ�У���ȱʡ���¡�һ������ť
    If vsAdvice.RowData(lngRow) = 0 Then
        i = GetPreRow(lngRow)
        If i = -1 Then
            mblnRowMerge = False
        Else
            mblnRowMerge = RowInһ����ҩ(i)
        End If
    Else
        mblnRowMerge = RowInһ����ҩ(lngRow)
    End If
    cbsMain.RecalcLayout '��ʱˢ��
        
    '��ʾ�������������˵��
    Call ShowOrHideQuestion
        
    Me.Refresh
    LockWindowUpdate Me.hWnd
    
    chk����.Visible = False
        
    On Error GoTo errH
    
    With vsAdvice
        If Val(.RowData(lngRow)) = 0 Then
            '��Ч�������Ƭ����
            Call ClearAdviceCard
            
            'ȱʡ��ʼʱ��
            Call txt��ʼʱ��_GotFocus
        Else
            '��Ƭ�༭
            blnEditable = True
            '�ѷ��͵�ҽ�������޸�
            If Val(.TextMatrix(lngRow, COL_״̬)) <> 1 Then blnEditable = False
            '��ǩ����ҽ�������޸�
            If Val(.TextMatrix(lngRow, COL_ǩ����)) = 1 Then blnEditable = False
            
            If .TextMatrix(lngRow, COL_���) = "K" And gblnѪ��ϵͳ Then
                'Ѫ�⻷��
                If Not (Val(.TextMatrix(lngRow, COL_��鷽��)) = 1 And Val(.TextMatrix(lngRow, COL_���״̬)) = 2) Then
                    '�ɳ� 5��Ѫ����Ѫ��
                    If Val(.TextMatrix(lngRow, COL_���״̬)) = 5 Or Val(.TextMatrix(lngRow, COL_���״̬)) = 2 Then blnEditable = False
                End If
            Else
                '���ͨ���Ĳ������޸�
                If Val(.TextMatrix(lngRow, COL_���״̬)) = 2 Then blnEditable = False
            End If
            
            '����Ѫҽ��ʱ���¿���ֱ�ӽ���4���״̬����ʱ��������༭�ģ������ֹ
            If Val(.TextMatrix(lngRow, COL_���״̬)) = 4 And .TextMatrix(lngRow, COL_���) = "K" Then
                If CanEditBloodAdvice(Val(.RowData(lngRow)), Val(.TextMatrix(lngRow, COL_���״̬)), Val(.TextMatrix(lngRow, COL_��־)) = 1, Val(.TextMatrix(lngRow, COL_��鷽��)) = 1, False) = False Then blnEditable = False
            End If
            
            '��ʾ���������Ա��
            If Val(.TextMatrix(lngRow, COL_��������)) = 1 And .TextMatrix(lngRow, COL_���) = "E" Then chk����.Visible = True
            mblnDoCheck = False
            chk����.value = Val(.TextMatrix(lngRow, COL_����))
            mblnDoCheck = True
            Call SetCardEditable(blnEditable)
            
            '��ȡ������Ŀ������Ϣ
            '---------------------
            If InStr(",4,5,6,7,", Val(.TextMatrix(lngRow, COL_���))) > 0 Then
                lngҩƷID = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID))
            End If
            
            
            If RowIn�䷽��(lngRow) Then
                txt����.MaxLength = 3
                '��ȡ��ҩ�䷽��һζ��ҩ��
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                lngҩƷID = Val(.TextMatrix(lngBaseRow, COL_�շ�ϸĿID))
            ElseIf RowIn������(lngRow) Then
                '��ȡһ�������ĵ�һ����Ŀ��
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                txt����.MaxLength = txt����.MaxLength
            Else
                lngBaseRow = lngRow
                txt����.MaxLength = txt����.MaxLength
            End If
            Set rsItem = Get������Ŀ��¼(Val(.TextMatrix(lngBaseRow, COL_������ĿID)))
            
            '��չ��ť����״̬(������,�������,����,��ҩ�䷽)
            cmdExt.Enabled = InStr(",7,C,F,D,", rsItem!���) > 0
            If rsItem!��� = "E" Or rsItem!��� = "K" Or rsItem!��� = "Z" Then
                If rsItem!��� = "K" And Val(.TextMatrix(lngRow, COL_�������) & "") <> 0 Then
                    cmdExt.Enabled = True
                Else
                    cmdExt.Enabled = CheckApplication(Val(.TextMatrix(lngBaseRow, COL_������ĿID)), 1)
                End If
            End If
            
            '��ʾ��ǰҽ����Ƭ����
            '--------------------------------------------------------------------------------------------
            '��ʼʱ�䣺ֻ������ҽ��ʱ�����޸Ŀ�ʼʱ��
            txt��ʼʱ��.Text = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
            Call SetStartTime(.TextMatrix(lngRow, COL_EDIT) = "1")
            
            'ҽ������
            txtҽ������.Text = .TextMatrix(lngRow, col_ҽ������)
            
            
            '����˵��
            txt����˵��.Text = .TextMatrix(lngRow, COL_����˵��)
            SetItemEditable , , , , , , , , , , , , IIF(.TextMatrix(lngRow, COL_�Ƿ���) = "1" Or .TextMatrix(lngRow, COL_�Ƿ���) = "1", 1, -1)
            
            If txt����˵��.Text <> "" Then
                cmdExcReason.Enabled = blnEditable
                cmdComExcReason.Enabled = blnEditable
            End If
                        
            '����������,��ҩ���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����¼��
            '----------------------
            If rsItem!��� = "7" Then '��ҩ�䷽(�в�ҩ)��Ȼ�е���,������������д
                SetItemEditable -1
            ElseIf (Nvl(rsItem!ִ��Ƶ��, 0) = 0 And InStr(",1,2,", Nvl(rsItem!���㷽ʽ, 0)) > 0) _
                    Or InStr(",5,6,", rsItem!���) > 0 Then
                SetItemEditable 1
                bln��ʾ���� = True
                txt����.Text = .TextMatrix(lngRow, COL_����)
                lbl������λ.Caption = .TextMatrix(lngRow, COL_������λ)
            Else
                SetItemEditable -1
            End If
            
            '��������ҩ���г�ҩ������ʹ�ã����ڼ�������
            'һ�㣺������ҩƷ(����ҩ)���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����ʹ���������Զ���������
            blnEditableTmp = False
            If InStr(",5,6,", rsItem!���) > 0 Then
                If mbln���� Then blnEditableTmp = True
            End If
            If blnEditableTmp Then
                SetDayState 1, 1
            Else
                SetDayState -1, -1
            End If
            txt����.Text = Val(.TextMatrix(lngRow, COL_����))
            If Val(txt����.Text) = 0 Then txt����.Text = ""
            
            '����
            '--------------------
            If rsItem!��� = "7" Then
                '��ҩ�䷽(�в�ҩ)��дΪ����
                SetItemEditable , 1
                lbl������λ.Caption = "��"
                txt����.Text = .TextMatrix(lngRow, COL_����) '����
                If Val(txt����.Text) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                    SetItemEditable , , , , , , , , , , , , 1
                End If
                bln��ʾ���� = True
                
                '��ɢװ��̬��ֻ�������䷽�����丶��
                If Val(.TextMatrix(lngRow, COL_��ҩ��̬)) <> 0 Then
                    txt����.Enabled = False
                    txt����.BackColor = Me.BackColor
                End If
            Else
                '��������Ҫ��д����:��������������Ϊ׼
                If rsItem!��� = "Z" And Nvl(rsItem!��������) <> "0" Then
                    SetItemEditable , -1 '����ҽ���������޸�����(�̶�Ϊ1��)
                ElseIf Nvl(rsItem!ִ��Ƶ��, 0) = 1 And Nvl(rsItem!���㷽ʽ, 0) = 3 Then
                    SetItemEditable , -1 'һ���Լƴ���Ŀ����������
                Else
                    SetItemEditable , 1
                    bln��ʾ���� = True
                End If
                lbl������λ.Caption = .TextMatrix(lngRow, COL_������λ)
                txt����.Text = .TextMatrix(lngRow, COL_����)
            End If
            
            '��ҩ;������ҩ�÷�
            '--------------
            If InStr(",5,6,", rsItem!���) > 0 Then
                SetItemEditable , , 1
                lbl�÷�.Caption = "��ҩ;��"
                '���Ҹ�ҩ;����Ӧ����:���ҵ�Rowdata(Variant)����ҪתΪLong��,���ܾ�ȷƥ��
                lng�÷�ID = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                lng�÷�ID = Val(.TextMatrix(lng�÷�ID, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = Get��Ŀ����(lng�÷�ID)
            ElseIf rsItem!��� = "K" Then
                '��Ѫҽ����Ҫ������ǰû����Ѫ;�������
                lng�÷�ID = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
                If lng�÷�ID <> -1 Then
                    SetItemEditable , , 1
                     If Val(.TextMatrix(lngRow, COL_��鷽��)) = 0 And gblnѪ��ϵͳ = True Then
                        lbl�÷�.Caption = "�ɼ�����"
                    Else
                        lbl�÷�.Caption = "��Ѫ;��"
                    End If
                    
                    lng�÷�ID = Val(.TextMatrix(lng�÷�ID, COL_������ĿID))
                    cmd�÷�.Tag = lng�÷�ID
                    txt�÷�.Text = Get��Ŀ����(lng�÷�ID)
                Else
                    SetItemEditable , , -1
                End If
            ElseIf rsItem!��� = "7" Then
                SetItemEditable , , 1
                lbl�÷�.Caption = "��ҩ�÷�"
                
                '��ҩ�䷽��ʾ�о�����ҩ�÷���
                lng�÷�ID = Val(.TextMatrix(lngRow, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = Get��Ŀ����(lng�÷�ID)
            ElseIf RowIn������(lngRow) Then '��������ж�,������ǰ�ļ���
                '�������
                SetItemEditable , , 1
                lbl�÷�.Caption = "�ɼ�����"
                
                '���������ʾ�о��ǲɼ�������
                lng�÷�ID = Val(.TextMatrix(lngRow, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = Get��Ŀ����(lng�÷�ID)
            Else
                SetItemEditable , , -1
            End If
            
            '����ʱ��/��Ѫʱ�䣺ֻ������/��Ѫ����(�������÷�λ��)
            If rsItem!��� = "F" Or rsItem!��� = "K" Then
                SetItemEditable , , , , , , , , 1
                If IsDate(.TextMatrix(lngRow, COL_����ʱ��)) Then
                    txt����ʱ��.Text = .TextMatrix(lngRow, COL_����ʱ��)
                Else
                    txt����ʱ��.Text = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                End If
                Call Set����ʱ��(rsItem!���)
            Else
                SetItemEditable , , , , , , , , -1
            End If
            
            'Ƶ��(������������ָ��ʹ��)
            If InStr("F,G,H,I", rsItem!���) > 0 Or rsItem!��� = "Z" And InStr(",1,2,3,4,5,6,7,8,9,10,11,12,14,", "," & rsItem!�������� & ",") > 0 Then
                SetItemEditable , , , -1
            Else
                SetItemEditable , , , 1
            End If
            cmdƵ��.Tag = .TextMatrix(lngRow, COL_Ƶ��)
            txtƵ��.Text = .TextMatrix(lngRow, COL_Ƶ��)
                    
            'ִ��ʱ�䣺"��ѡƵ��"��ҩƷ��
            If (Nvl(rsItem!ִ��Ƶ��, 0) = 0 Or InStr(",5,6,7,", rsItem!���) > 0) And .TextMatrix(lngRow, COL_�����λ) <> "����" Then
                SetItemEditable , , , , 1
                Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(lngRow), .TextMatrix(lngRow, COL_Ƶ��), lng�÷�ID)
                cboִ��ʱ��.Text = .TextMatrix(lngRow, COL_ִ��ʱ��)
            Else
                SetItemEditable , , , , -1
            End If
                    
            '���٣���Һ���ҩ;����ҩƷ��������
            If InStr(",5,6,", rsItem!���) > 0 Then
                i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                If Val(.TextMatrix(i, COL_ִ�з���)) = 1 Then
                    SetItemEditable , , , , , , , , , 1
                    If InStr(.TextMatrix(i, COL_ҽ������), "��/����") > 0 Then
                        lbl���ٵ�λ.Caption = "��/����"
                    ElseIf InStr(.TextMatrix(i, COL_ҽ������), "����/Сʱ") > 0 Then
                        lbl���ٵ�λ.Caption = "����/Сʱ"
                    End If
                    Call Load��Һ����(cbo����, lbl���ٵ�λ, False)
                    cbo����.Text = Replace(.TextMatrix(i, COL_ҽ������), lbl���ٵ�λ.Caption, "")
                Else
                    SetItemEditable , , , , , , , , , -1
                End If
            Else
                SetItemEditable , , , , , , , , , -1
            End If
            
            
            '��ҩĿ�ĺ���ҩ����
            If Val(.TextMatrix(lngRow, COL_�����ȼ�)) = 0 Then
                SetItemEditable , , , , , , , , , , -1, -1
            Else
                SetItemEditable , , , , , , , , , , 1, 1
                bln��ʾ����ҩ�� = True
                
                If .TextMatrix(lngRow, COL_��ҩĿ��) = "1" Then
                    Call zlControl.CboSetIndex(cboDruPur.hWnd, 1)
                ElseIf .TextMatrix(lngRow, COL_��ҩĿ��) = "2" Then
                    Call zlControl.CboSetIndex(cboDruPur.hWnd, 2)
                End If
                
                txt��ҩ����.Text = .TextMatrix(lngRow, COL_��ҩ����)
            End If
            cboDruPur.Enabled = blnEditable
            cmdReason.Enabled = blnEditable
            cmd�ղ���ҩ����.Enabled = blnEditable
            
            '��Ѫҽ����ʾ��ҩĿ�ģ�����ҩĿ����ʾΪ��Ѫԭ��
            If (gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ) And .TextMatrix(lngRow, COL_���) = "K" Then
                bln��ʾ����ҩ�� = True
                SetItemEditable , , , , , , , , , , , , , 1
                txt��ҩ����.Width = cmdExcReason.Left + cmdExcReason.Width - txt��ҩ����.Left
                If .TextMatrix(lngRow, COL_��־) = "1" Then
                    SetItemEditable , , , , , , , , , , , 1
                    txt��ҩ����.Text = .TextMatrix(lngRow, COL_��ҩ����)
                Else
                    SetItemEditable , , , , , , , , , , , -1
                End If
            Else
                SetItemEditable , , , , , , , , , , , , , -1
                txt��ҩ����.Width = cmdExcReason.Left + cmdExcReason.Width - txt��ҩ����.Left
                cmd�ղ���ҩ����.Left = txt��ҩ����.Left + txt��ҩ����.Width + 30
                cmdReason.Left = txt��ҩ����.Left + txt��ҩ����.Width - cmdReason.Width
            End If
            
            'ҽ������
            cboҽ������.Text = .TextMatrix(lngRow, COL_ҽ������)
                    
            'ִ������
            If InStr(",5,6,7,", rsItem!���) > 0 Then
                '������Թ�ҩ��̶�ѡ���Ա�ҩ
                If Val(.TextMatrix(lngRow, COL_�ٴ��Թ�ҩ)) = 1 And InStr(",5,6,", rsItem!���) > 0 Then
                    strTmp = "�Ա�ҩ"
                Else
                    If rsItem!��� = "7" Then
                        '������ҩ�䷽,����������Ŀ���������Ƽ���������,�������÷��ͼ巨һ��ΪԺ��ִ��,һ����Ϊ
                        If Val(.TextMatrix(lngBaseRow, COL_ִ������)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������)) <> 5 Then
                            strTmp = "�Ա�ҩ"
                        ElseIf Val(.TextMatrix(lngBaseRow, COL_ִ������)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                            strTmp = "��Ժ��ҩ"
                        Else
                            strTmp = "����"
                        End If
                    Else
                        i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                        If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                            strTmp = "�Ա�ҩ"
                        ElseIf Val(.TextMatrix(lngRow, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                            strTmp = "��Ժ��ҩ"
                        Else
                            strTmp = "����"
                        End If
                    End If
                End If
                Call SetCboִ������(gbln����ҩ��ʹ���Ա�ҩ Or Not gblnKSSStrict Or Val(.TextMatrix(lngRow, COL_�����ȼ�)) = 0, Val(.TextMatrix(lngRow, COL_�ٴ��Թ�ҩ)) = 1 And rsItem!��� & "" <> "7")
                SetItemEditable , , , , , , 1
                Call SeekIndex(cboִ������, strTmp)
            Else
                SetItemEditable , , , , , , -1
            End If
            
            lblִ�п���.Caption = "ִ�п���"
            'ִ�п���:���ۻ�סԺҽ�����ٴ�����
            If rsItem!��� = "Z" And InStr(",1,2,", Nvl(rsItem!��������, 0)) > 0 Then
                SetItemEditable , , , , , 1
                If Nvl(rsItem!��������, 0) = 1 Then
                    lblִ�п���.Caption = "���ۿ���"
                    '����:���������סԺ�ٴ�����,�ɷ������������������ۻ�סԺ����
                    Call Get�ٴ�����(3, , Val(.TextMatrix(lngRow, COL_ִ�п���ID)), cboִ�п���, True, False, True)
                ElseIf Nvl(rsItem!��������, 0) = 2 Then
                    lblִ�п���.Caption = "סԺ����"
                    'סԺ:����סԺ�ٴ�����
                    Call Get�ٴ�����(2, , Val(.TextMatrix(lngRow, COL_ִ�п���ID)), cboִ�п���, True, False, True, 1)
                End If
                If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) <> 0 And cboִ�п���.ListIndex = -1 Then .TextMatrix(lngRow, COL_ִ�п���ID) = 0
            Else
                '��ҩƷ����ҩƷ��Ϊ׼��ʾ,��������Լ�����ĿΪ׼��ʾ
                i = lngRow
                If rsItem!��� = "7" Then
                    i = lngBaseRow
                ElseIf RowIn������(lngRow) Then '��������ж�,������ǰ�ļ���
                    i = lngBaseRow
                End If
                
                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                    '�Ƕ�����Ժ��ִ��ʱ����ʾ�Ϳ���ѡ��(����ҩƷ)
                    SetItemEditable , , , , , 1
                    Call Get����ִ�п���(mlng����ID, 0, cboִ�п���, rsItem!���, rsItem!ID, lngҩƷID, Nvl(rsItem!ִ�п���, 0), _
                        mlng���˿���id, Val(.TextMatrix(i, COL_��������ID)), Val(.TextMatrix(i, COL_ִ�п���ID)), 1, 1, , blnEditable)
                        
                     
                    '��ɢװ��̬��ֻ�������䷽����ѡҩ��
                    If rsItem!��� = "7" Then
                        If Val(.TextMatrix(lngRow, COL_��ҩ��̬)) <> 0 Then
                            cboִ�п���.Enabled = False
                            cboִ�п���.BackColor = Me.BackColor
                        End If
                    End If
                ElseIf InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                    SetItemEditable , , , , , -1
                    If Val(.TextMatrix(i, COL_ִ������)) = 0 Then
                        cboִ�п���.AddItem "<��ִ�ж���>"
                    Else
                        cboִ�п���.AddItem "-"
                    End If
                    Call zlControl.CboSetIndex(cboִ�п���.hWnd, 0)
                End If
                If Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 And cboִ�п���.ListIndex = -1 Then .TextMatrix(i, COL_ִ�п���ID) = 0
                If InStr("5,6,7", rsItem!���) > 0 Then lblִ�п���.Caption = "��ҩҩ��"
            End If

            If cboִ�п���.ListIndex = -1 And cboִ�п���.ListCount = 1 Then
                If Val(.TextMatrix(i, COL_״̬)) < 3 Then
                    cboִ�п���.ListIndex = 0
                Else
                    Call zlControl.CboSetIndex(cboִ�п���.hWnd, 0)
                End If
            End If
            
            If cboִ�п���.ListCount = 1 Then
                If cboִ�п���.List(cboִ�п���.ListIndex) <> "[����...]" Then
                    cboִ�п���.Enabled = False
                Else
                    cboִ�п���.Enabled = True
                End If
            Else
                cboִ�п���.Enabled = True
            End If
            
            '����ִ��:ָ��ҩ;��,��ҩ�÷�,��������,�ɼ���ʽ��ִ�п��ң�ԭҺƤ����Ŀ
            If Should����ִ��(lngRow, i, strTmp) Then
                If .TextMatrix(lngRow, COL_���) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5" Then
                    '����ԭҺƤ�Լ���ҩ��
                    lngҩƷID = GetԭҺƤ��ҩƷ(Val(.TextMatrix(i, COL_������ĿID)))
                    If lngҩƷID <> 0 Then
                        SetItemEditable , , , , , , , 1
                        Call Get����ִ�п���(mlng����ID, 0, cbo����ִ��, "5", 0, lngҩƷID, 0, mlng���˿���id, 0, Val(.TextMatrix(i, COL_��ҩ����)), 1, 1, , blnEditable)
                        '���û��ѡҩ���򲻼����κ�Ĭ��ֵ
                        If Val(.TextMatrix(i, COL_��ҩ����)) = 0 Then cbo����ִ��.ListIndex = -1
                    Else
                        cbo����ִ��.Clear
                        SetItemEditable , , , , , , , -1
                    End If
                Else
                    'ִ�п���:��ҩ��ҩƷ���������ĵ�ר��ȡ
                    SetItemEditable , , , , , , , 1
                    Call Get����ִ�п���(mlng����ID, 0, cbo����ִ��, .TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), lngҩƷID, _
                        Val(.TextMatrix(i, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(i, COL_��������ID)), Val(.TextMatrix(i, COL_ִ�п���ID)), 1, 1, , blnEditable)
                        
                    If Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 And cbo����ִ��.ListIndex = -1 Then .TextMatrix(i, COL_ִ�п���ID) = 0
                    If cbo����ִ��.ListIndex = -1 And cbo����ִ��.ListCount = 1 Then cbo����ִ��.ListIndex = 0
                End If
            Else
                SetItemEditable , , , , , , , -1
                If i <> -1 Then
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                        If Val(.TextMatrix(i, COL_ִ������)) = 0 Then
                            cbo����ִ��.AddItem "<��ִ�ж���>"
                        ElseIf Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                            cbo����ִ��.AddItem "-"
                        End If
                        Call zlControl.CboSetIndex(cbo����ִ��.hWnd, 0)
                    End If
                End If
            End If
            lbl����ִ��.Caption = strTmp
            
            '������־
            chk����.Visible = True
            mblnDoCheck = False
            chk����.value = Val(.TextMatrix(lngRow, COL_��־))
            chkZeroBilling.value = Val(.TextMatrix(lngRow, COL_��Ѽ���))
            mblnDoCheck = True
                        
            
            '��ʾҩƷ��棺�����ﵥλ����ҩ�䷽����ʾ
            '----------------------------------------
            If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 _
                And (InStr(",5,6,", rsItem!���) > 0 Or rsItem!��� = "4" And Val(.TextMatrix(lngRow, COL_��������)) = 1) Then
                If .TextMatrix(lngRow, COL_���) = "" And Val(.TextMatrix(lngRow, COL_ִ�п���ID)) <> 0 Then Call GetDrugStock(lngRow)
                If .TextMatrix(lngRow, COL_���) <> "" And Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then .TextMatrix(lngRow, COL_���) = ""
                If .TextMatrix(lngRow, COL_���) <> "" Then
                    If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                        stbThis.Panels(3).Text = IIF(Val(.TextMatrix(lngRow, COL_���)) > 0, "�п��", "�޿��")
                    Else
                        stbThis.Panels(3).Text = "���:" & FormatEx(Val(.TextMatrix(lngRow, COL_���)), 5) & .TextMatrix(lngRow, COL_���ﵥλ)
                    End If
                Else
                    stbThis.Panels(3).Text = ""
                End If
            Else
                If rsItem!��� = "7" And Val(.TextMatrix(lngRow, COL_״̬)) = 1 Then
                    Call GetDrugStock(lngRow)
                End If
                stbThis.Panels(3).Text = ""
            End If
            
            '��ʾҽ�����ۺͷ�������
            If .TextMatrix(lngRow, COL_����) = "" Then '��""������
                .TextMatrix(lngRow, COL_����) = GetItemPrice(lngRow)
            End If
            dblPrice = Val(.TextMatrix(lngRow, COL_����))
            If dblPrice <> 0 Then
                If InStr(",4,5,6,", rsItem!���) > 0 Then
                    stbThis.Panels(4).Text = "ÿ" & .TextMatrix(lngRow, COL_���ﵥλ) & ":" & FormatEx(dblPrice, 5) & "Ԫ"
                ElseIf rsItem!��� = "7" Then
                    stbThis.Panels(4).Text = "ÿ��:" & FormatEx(dblPrice, 5) & "Ԫ"
                Else
                    stbThis.Panels(4).Text = IIF(IsNull(rsItem!���㵥λ), "�۸�:", "ÿ" & Nvl(rsItem!���㵥λ) & ":") & FormatEx(dblPrice, 5) & "Ԫ"
                End If
            Else
                stbThis.Panels(4).Text = ""
            End If
            
            '��ʾ��������
            strTmp = Get��������(lngRow)
            If strTmp <> "" Then
                stbThis.Panels(4).Text = stbThis.Panels(4).Text & IIF(stbThis.Panels(4).Text <> "", ",", "") & strTmp
            End If
            
            '����˵���Ѫҽ����������:��Ѫ�ɷ֡�ִ�п��ҡ�Ԥ����Ѫ��
            If .TextMatrix(lngRow, COL_���) = "K" And Val(.TextMatrix(lngRow, COL_��鷽��)) = 1 And gblnѪ��ϵͳ = True And blnEditable Then
                If InitObjBlood = True Then
                    If gobjPublicBlood.GetPrepareBloodRs(Val(.RowData(lngRow)), rsBlood) = True Then
                        If Val(rsBlood!��¼���� & "") = 2 And Val(rsBlood!��¼״̬ & "") = 1 Then
                            cmdSel.Enabled = False
                            txtҽ������.Enabled = False
                            txt����.Enabled = False
                            cboִ�п���.Enabled = False
                        End If
                    End If
                End If
            End If
            
        End If
    End With
    
    '����༭��־
    Call ClearItemTag
    
    '��ʾ������ҩƷ��ص�������Ŀ
    Call SetMediInfoItem(bln��ʾ����, bln��ʾ����ҩ��)
    
    '��ʾ�Ƽ۴���
    Call ShowPrice(lngRow)
    
    '������ҽӿ�
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        If OldRow <> NewRow Then
            Call zlPluginAdviceRowChange(NewRow)
            With vsAdvice
                If OldRow <> -1 Then
                    If Val(.RowData(OldRow)) <> 0 Then
                        If Val(.TextMatrix(OldRow, COL_EDIT)) = 1 Or Val(.TextMatrix(OldRow, COL_EDIT)) = 2 Then
                            Call zlPluginAdviceRowChange(OldRow, 1)
                        End If
                    End If
                End If
            End With
        End If
    End If
    
    cbsMain.RecalcLayout '��ʱˢ��,��Lock�ɲ�Ҫ
    LockWindowUpdate 0
    Exit Sub
errH:
    LockWindowUpdate 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowPrice(ByVal lngRow As Long)
'���ݵ�ǰ�е������ʾ�Ƽ۴���
    If mblnModal Then Exit Sub
    
    If vsAdvice.RowData(lngRow) = 0 Or Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_״̬))) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",4,5,6,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf RowIn�䷽��(lngRow) Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf stbThis.Panels("Price").Bevel = sbrNoBevel Then
        stbThis.Panels("Price").Visible = True
        If Val(stbThis.Panels("Price").Tag) <> 0 Then
            stbThis.Panels("Price").Bevel = sbrInset
        Else
            stbThis.Panels("Price").Bevel = sbrRaised
        End If
    End If
    
    If stbThis.Panels("Price").Bevel <> sbrInset Then
        '�رռƼ۴���
        mfrmPrice.HideMe
    Else
        Call mfrmPrice.ShowMe(Me, vsAdvice, mlng����ID, 0, mlng���˿���id, 1, mint����, _
            COL_��� & "," & COL_���ID & "," & COL_״̬ & "," & COL_��� & "," & COL_������ĿID & "," & _
            COL_�շ�ϸĿID & "," & COL_�걾��λ & "," & COL_��鷽�� & "," & COL_ִ�б�� & "," & COL_�Ƽ����� & "," & COL_ִ������ & "," & COL_ִ�п���ID)
    End If
End Sub

Private Function Get��������(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еķ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str���� As String, str���� As String, lng�շ�ϸĿID As Long
    
    lng�շ�ϸĿID = Val(vsAdvice.TextMatrix(lngRow, COL_�շ�ϸĿID))
    If lng�շ�ϸĿID <> 0 Then
        'ȡҽ���ķ�������
        If mint���� <> 0 Then
            str���� = gclsInsure.GetItemInsure(mlng����ID, lng�շ�ϸĿID, 0, True, mint����)
            If str���� <> "" Then
                If UBound(Split(str����, ";")) >= 5 Then
                    str���� = Split(str����, ";")(5)
                Else
                    str���� = ""
                End If
            End If
        End If
        'û����ȡHIS�ķ�������
        strSQL = "Select A.��������,N.���� as ҽ������ From �շ���ĿĿ¼ A,����֧����Ŀ M,����֧������ N" & _
            " Where A.ID=[1] And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�շ�ϸĿID, mint����)
        If Not rsTmp.EOF Then
            If str���� = "" Then str���� = Nvl(rsTmp!��������)
            str���� = Nvl(rsTmp!ҽ������)
        End If
    End If
        
    Get�������� = Mid(IIF(str���� <> "", ",����:" & str����, "") & IIF(str���� <> "", ",����:" & str����, ""), 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Should����ִ��(ByVal lngRow As Long, lngRow2 As Long, strִ�п��� As String) As Boolean
'���ܣ��ж�ָ����ҽ����(�ɼ���)�Ƿ�������ø��ӵ�ִ�п���
'������lngRow2=���ظ����е�ҽ���к�
'      strִ�п���=����ִ�п�������
    Dim i As Long
    
    lngRow2 = -1
    strִ�п��� = "����ִ��"
    With vsAdvice
        If lngRow = 0 Or .RowData(lngRow) = 0 Then Exit Function
        
        If RowIn�䷽��(lngRow) Then
            '��ҩ�÷�
            lngRow2 = lngRow
            strִ�п��� = "�÷�ִ��"
            Should����ִ�� = True
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            '��ҩ;��
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
            strִ�п��� = "��ҩִ��"
            Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "F" Then
            '��������
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            strִ�п��� = "����ִ��"
            If lngRow2 <> -1 Then Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "K" Then
            '��Ѫ;��
            If Val(.TextMatrix(lngRow, COL_��鷽��)) = 0 And gblnѪ��ϵͳ = True Then
                strִ�п��� = "�ɼ�ִ��"
            Else
                strִ�п��� = "��Ѫִ��"
            End If
            
            lngRow2 = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
            If lngRow2 <> -1 Then Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "E" _
            And .TextMatrix(lngRow - 1, COL_���) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
            '�ɼ���ʽ
            lngRow2 = lngRow
            strִ�п��� = "�ɼ�ִ��"
            Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5" Then
            'ԭҺƤ��ҩ��
            lngRow2 = lngRow
            strִ�п��� = "ԭҺҩ��"
            Should����ִ�� = True
        End If
        
        '������Ժ��ִ��
        If Should����ִ�� Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_ִ������))) > 0 Then
                Should����ִ�� = False
            End If
        End If
    End With
End Function

Private Function GetItemPrice(ByVal lngRow As Long) As Double
'���ܣ���ȡ��ǰҽ���еļ۸�(ҩƷΪһ��ҩ����װ�ĵ���,���������շѶ���)
'˵����ҩƷ��������ҩ;������ҩ�÷��巨
    Dim rsTmp As New ADODB.Recordset
    Dim strҽ��IDs As String, str����s As String, str�����շ� As String
    Dim str��ĿIDs As String, strҽ�� As String, str��Ŀ���� As String, strTmp As String
    Dim strAdviceIDs As String, lngִ�п���ID As Long
    Dim dblPrice As Double, dbl���� As Double
    Dim blnҩƷ As Boolean, strSQL As String
    Dim str������Ŀ As String, i As Long
    
    With vsAdvice
        blnҩƷ = True
        If InStr(",4,5,6,", .TextMatrix(lngRow, COL_���)) > 0 And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
            '��ҩ���г�ҩ������²��ܼ���۸�
            If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(lngRow, COL_�շ�ϸĿID))
            End If
            lngִ�п���ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        ElseIf RowIn�䷽��(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" And Val(.TextMatrix(i, COL_�շ�ϸĿID)) <> 0 Then
                        If lngִ�п���ID = 0 Then
                            lngִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                        End If
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COL_�շ�ϸĿID))
                        str����s = str����s & ";" & Val(.TextMatrix(i, COL_����))
                    End If
                Else
                    Exit For
                End If
            Next
        Else
            blnҩƷ = False
            '����ҽ��,δУ��(�Ƽ�)�İ��շѶ��ռ���,����ֱ��ȡҽ���Ƽ�
            '���������Ƽۺ��ֹ��Ƽ۵���Ŀ,������������Ժ��ִ�е���Ŀ
            If Val(.TextMatrix(lngRow, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                If InStr(",1,2,", .TextMatrix(lngRow, COL_״̬)) > 0 Then
                    str������Ŀ = Val(.TextMatrix(lngRow, COL_������ĿID))
                    If RowIn������(lngRow) Then
                        i = .FindRow(CStr(.RowData(lngRow)), .FixedRows, COL_���ID)
                        If i <> -1 Then
                            str������Ŀ = Val(.TextMatrix(i, COL_������ĿID))
                        End If
                    End If
                    
                    str��Ŀ���� = str��Ŀ���� & "," & Val(.TextMatrix(lngRow, COL_������ĿID)) & ":" & Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(lngRow, COL_������ĿID))
                    strҽ�� = strҽ�� & " Union ALL Select " & _
                        IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, "-NULL", Val(.TextMatrix(lngRow, COL_���ID))) & " as ���ID," & _
                        Val(.TextMatrix(lngRow, COL_������ĿID)) & " as ������ĿID," & Val(.TextMatrix(lngRow, COL_ִ�б��)) & " as ִ�б��," & _
                        IIF(.TextMatrix(lngRow, COL_�걾��λ) = "", "NULL", "'" & .TextMatrix(lngRow, COL_�걾��λ) & "'") & " as �걾��λ," & _
                        IIF(.TextMatrix(lngRow, COL_��鷽��) = "", "NULL", "'" & .TextMatrix(lngRow, COL_��鷽��) & "'") & " as ��鷽��," & _
                        str������Ŀ & " as ������ĿID From Dual"
                Else
                    strҽ��IDs = strҽ��IDs & "," & .RowData(lngRow)
                End If
            End If
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_״̬)) > 0 Then
                            strTmp = Val(.TextMatrix(i, COL_������ĿID)) & ":" & Val(.TextMatrix(i, COL_ִ�п���ID))
                            If InStr("," & str��Ŀ���� & ",", "," & strTmp & ",") = 0 Then str��Ŀ���� = str��Ŀ���� & "," & strTmp
                            
                            strҽ�� = strҽ�� & " Union ALL Select " & _
                                IIF(Val(.TextMatrix(i, COL_���ID)) = 0, "-NULL", Val(.TextMatrix(i, COL_���ID))) & " as ���ID," & _
                                Val(.TextMatrix(i, COL_������ĿID)) & " as ������ĿID," & Val(.TextMatrix(i, COL_ִ�б��)) & " as ִ�б��," & _
                                IIF(.TextMatrix(i, COL_�걾��λ) = "", "NULL", "'" & .TextMatrix(i, COL_�걾��λ) & "'") & " as �걾��λ," & _
                                IIF(.TextMatrix(i, COL_��鷽��) = "", "NULL", "'" & .TextMatrix(i, COL_��鷽��) & "'") & " as ��鷽��," & _
                                Val(.TextMatrix(i, COL_������ĿID)) & " as ������ĿID From Dual"
                        Else
                            strҽ��IDs = strҽ��IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1 '�������
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_״̬)) > 0 Then
                            strTmp = Val(.TextMatrix(i, COL_������ĿID)) & ":" & Val(.TextMatrix(i, COL_ִ�п���ID))
                            If InStr("," & str��Ŀ���� & ",", "," & strTmp & ",") = 0 Then str��Ŀ���� = str��Ŀ���� & "," & strTmp
                            
                            strҽ�� = strҽ�� & " Union ALL Select " & _
                                IIF(Val(.TextMatrix(i, COL_���ID)) = 0, "-NULL", Val(.TextMatrix(i, COL_���ID))) & " as ���ID," & _
                                Val(.TextMatrix(i, COL_������ĿID)) & " as ������ĿID," & Val(.TextMatrix(i, COL_ִ�б��)) & " as ִ�б��," & _
                                IIF(.TextMatrix(i, COL_�걾��λ) = "", "NULL", "'" & .TextMatrix(i, COL_�걾��λ) & "'") & " as �걾��λ," & _
                                IIF(.TextMatrix(i, COL_��鷽��) = "", "NULL", "'" & .TextMatrix(i, COL_��鷽��) & "'") & " as ��鷽��," & _
                                Val(.TextMatrix(i, COL_������ĿID)) & " as ������ĿID From Dual"
                        Else
                            strҽ��IDs = strҽ��IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    strҽ��IDs = Mid(strҽ��IDs, 2)
    str����s = Mid(str����s, 2)
    str��ĿIDs = Mid(str��ĿIDs, 2)
    str��Ŀ���� = Mid(str��Ŀ����, 2)
    strҽ�� = Mid(strҽ��, 12)
    
    On Error GoTo errH
    
    If blnҩƷ Then
        If str��ĿIDs = "" Then Exit Function
    
        strSQL = "Select Rownum As ���,Column_Value As ID From Table(f_Num2list([1]))"
        strSQL = "Select /*+ RULE */ A.ID,A.���,A.�Ƿ���,D.��������,Nvl(B.�����װ,1) as �����װ,Nvl(B.����ϵ��,1) as ����ϵ��,B.����ɷ���� As �ɷ����" & _
            " From �շ���ĿĿ¼ A,ҩƷ��� B,�������� D,(" & strSQL & ") C" & _
            " Where A.ID=B.ҩƷID(+) And A.ID=D.����ID(+) And A.ID=C.ID Order By C.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��ĿIDs)
        For i = 1 To rsTmp.RecordCount
            '�ۼ�����
            If str����s <> "" Then '��ҩ�䷽�Ź�ÿζ����
                dbl���� = Val(Split(str����s, ";")(i - 1))
                
                '�ۼ���������ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Nvl(rsTmp!�ɷ����, 0) = 0 Then
                    dbl���� = Format(dbl���� / Nvl(rsTmp!����ϵ��, 1), "0.00000")
                Else
                    dbl���� = Format(IntEx(dbl���� / Nvl(rsTmp!����ϵ��, 1) / Nvl(rsTmp!�����װ, 1)) * Nvl(rsTmp!�����װ, 1), "0.00000")
                End If
            Else
                dbl���� = Nvl(rsTmp!�����װ, 1) '1��ҩ����λ���ۼ�����
            End If
            If Nvl(rsTmp!�Ƿ���, 0) = 1 And (rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1 Or InStr(",5,6,7,", rsTmp!���) > 0) Then
                dblPrice = dblPrice + Format(Format(CalcDrugPrice(rsTmp!ID, lngִ�п���ID, dbl����), "0.00000") * dbl����, "0.00000")
            Else
                dblPrice = dblPrice + Format(Format(CalcPrice(rsTmp!ID), "0.00000") * dbl����, "0.00000")
            End If
            
            rsTmp.MoveNext
        Next
    Else
        If strҽ�� = "" And strҽ��IDs = "" Then Exit Function
    
        If strҽ��IDs <> "" Then
            strSQL = _
                " Select /*+ RULE */B.����,Decode(C.�Ƿ���,1,B.����,Sum(D.�ּ�)) as ����" & _
                " From ����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where B.�շ�ϸĿID=C.ID And B.�շ�ϸĿID=D.�շ�ϸĿID" & _
                " And ((Sysdate Between D.ִ������ And D.��ֹ����) Or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                " And B.ҽ��ID IN(Select Column_Value From Table(f_Num2list([1])))" & _
                " Group by B.����,C.�Ƿ���,B.����"
        End If
        If strҽ�� <> "" Then
            '����û�мӲ�λ������������Ҫ��Distinct
            str�����շ� = "Select * From (" & _
                "Select Distinct C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,C.���ÿ���id" & _
                " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                " From �����շѹ�ϵ C,Table(f_Num2list2([2])) D Where C.������ĿID=D.c1" & _
                "      And (C.���ÿ���ID is Null or C.���ÿ���ID = D.c2 And C.������Դ = 1)" & _
                " ) Where Nvl(���ÿ���id, 0) = Top"
                
            strSQL = IIF(strSQL = "", "", strSQL & " Union ALL") & _
                " Select " & IIF(strSQL = "", "/*+ RULE */", "") & "B.�շ����� as ����,Decode(C.�Ƿ���,1,Sum(D.ȱʡ�۸�),Sum(D.�ּ�)) as ����" & _
                " From (" & strҽ�� & ") A,(" & str�����շ� & ") B,�շ���ĿĿ¼ C,�շѼ�Ŀ D,������ĿĿ¼ E,��Ѫ������ F" & _
                " Where A.������ĿID=B.������ĿID And B.�շ���ĿID=C.ID And B.�շ���ĿID=D.�շ�ϸĿID" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) And A.������ĿID=E.ID And E.�Թܱ���=F.����(+)" & _
                " And (Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And B.�շ���ĿID=F.����ID Or Not(Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And F.����ID Is Not NULL))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                " And (A.���ID is Null And A.ִ�б�� IN(1,2) And B.��������=1" & _
                "       Or A.�걾��λ=B.��鲿λ And A.��鷽��=B.��鷽�� And Nvl(B.��������,0)=0" & _
                "       Or (A.��鷽�� is Null Or e.��� = 'E' And e.��������='4') And Nvl(B.��������,0)=0 And B.��鲿λ is Null And B.��鷽�� is Null)" & _
                " Group by B.�շ�����,C.�Ƿ���"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strҽ��IDs, str��Ŀ����)
        For i = 1 To rsTmp.RecordCount
            dblPrice = dblPrice + Format(Nvl(rsTmp!����, 0) * Nvl(rsTmp!����, 0), "0.00000")
            rsTmp.MoveNext
        Next
    End If
    
    GetItemPrice = Format(dblPrice, gstrDecPrice)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function MakePriceRecord(ByVal lngRow As Long) As Boolean
'���ܣ����ݵ�ǰ�¿�ҽ�������ݣ����ɶ�Ӧ������ҽ���ķ�����ϸ��¼��
'������lngRow=��ǰҽ����
'���أ��мƼ����ݼ�¼�����ݲŷ���True
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strAdvice As String, str��Ŀ���� As String, str�����շ� As String
    Dim lngBegin As Long, lngEnd As Long
    Dim lngִ�п���ID As Long, blnLoad As Boolean
    Dim dbl���� As Double, dbl��� As Double, dblʵ�� As Double
    Dim str������Ŀ As String, i As Long
    Dim str��Ŀ As String, blnDo As Boolean
    Dim lng��ID As Long, lng���ID As Long, lng����ID As Long
    
    On Error GoTo errH
        
    With vsAdvice
        If .RowData(lngRow) = 0 Then Exit Function
        
        If RowIn������(lngRow) Then
            i = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If i <> -1 Then str������Ŀ = Val(.TextMatrix(i, COL_������ĿID))
        End If
        
        '���ɲ���ҽ����¼��ʱ��
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                str��Ŀ���� = str��Ŀ���� & "," & Val(.TextMatrix(i, COL_������ĿID)) & ":" & Val(.TextMatrix(i, COL_ִ�п���ID))
                strAdvice = strAdvice & " Union ALL " & _
                    "Select " & .RowData(i) & " as ID," & Val(.TextMatrix(i, COL_���)) & " as ���," & _
                    ZVal(.TextMatrix(i, COL_���ID), True) & " as ���ID,'" & .TextMatrix(i, COL_���) & "' as �������," & _
                    IIF(str������Ŀ = "", Val(.TextMatrix(i, COL_������ĿID)), str������Ŀ) & " as ������ĿID," & _
                    Val(.TextMatrix(i, COL_������ĿID)) & " as ������ĿID," & ZVal(.TextMatrix(i, COL_�շ�ϸĿID), True) & " as �շ�ϸĿID," & _
                    Val(.TextMatrix(i, COL_����)) & " as ����," & Val(.TextMatrix(i, COL_����)) & " as ����," & _
                    "'" & .TextMatrix(i, COL_�걾��λ) & "' as �걾��λ,'" & .TextMatrix(i, COL_��鷽��) & "' as ��鷽��," & _
                    Val(.TextMatrix(i, COL_ִ�б��)) & " as ִ�б��," & Val(.TextMatrix(i, COL_�Ƽ�����)) & " as �Ƽ�����," & _
                    IIF(.TextMatrix(i, COL_���) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0, 1, 0) & " as ��������," & _
                    Val(.TextMatrix(i, COL_ִ������)) & " as ִ������," & ZVal(.TextMatrix(i, COL_ִ�п���ID), True) & " as ִ�п���ID From Dual"
            End If
        Next
        strAdvice = Mid(strAdvice, 12)
        str��Ŀ���� = Mid(str��Ŀ����, 2)
    End With
    
    blnLoad = True
    
    'ҩƷ�����ĵļƼۣ����ۼ����������ۼ���
    If vsAdvice.TextMatrix(lngRow, COL_���) = "4" Then
        '���ģ��̶�������´�
        strSQL = "Select A.���,A.�������,C.��� as �շ����,A.�շ�ϸĿID,D.������ĿID," & _
            " Decode(A.����,0,1,A.����) as ����,Decode(Nvl(C.�Ƿ���,0),1,D.ȱʡ�۸�,D.�ּ�) as ����," & _
            " C.�Ƿ���,C.���ηѱ�,B.��������,A.ִ�п���ID,A.��������,D.�����շ���" & _
            " From (" & strAdvice & ") A,�������� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
            " Where A.ID=[1] And Nvl(A.ִ������,0)<>5 And A.�շ�ϸĿID=B.����ID" & _
            " And A.�շ�ϸĿID=C.ID And C.������� IN(1,3) And D.�շ�ϸĿID=C.ID" & _
            " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        blnLoad = False
    ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
        '��,����ҩ:���ܰ������ҽ��,����1��ҩ����װ�ĵ���
        strSQL = "Select A.���,A.�������,C.��� as �շ����,C.ID as �շ�ϸĿID,D.������ĿID," & _
            " Decode(A.����,0,1,A.����)/Decode(1,1,B.�����װ,B.סԺ��װ) as ����," & _
            " Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�) as ����,C.�Ƿ���,C.���ηѱ�," & _
            " 0 as ��������,A.ִ�п���ID,A.��������,D.�����շ���" & _
            " From (" & strAdvice & ") A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
            " Where A.ID=[1] And Nvl(A.ִ������,0)<>5 And A.�շ�ϸĿID=B.ҩƷID" & _
            " And B.ҩƷID=C.ID And C.������� IN(1,3) And D.�շ�ϸĿID=C.ID" & _
            " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
            
        '��һ����ҩ(�����)�ĵ�һ��ҩ�в���ʾ��ҩ;���ļƼ�
        blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_���ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
    ElseIf RowIn�䷽��(lngRow) Then
        '�в�ҩ:һ����Ӧ�й���¼����д���շ�ϸĿID
        strSQL = "Select A.���,A.�������,C.��� as �շ����,C.ID as �շ�ϸĿID,D.������ĿID," & _
            " Decode(A.����,0,1,A.����)*A.����/Nvl(B.����ϵ��,1) as ����," & _
            " Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�) as ����,C.�Ƿ���,C.���ηѱ�," & _
            " 0 as ��������,A.ִ�п���ID,A.��������,D.�����շ���" & _
            " From (" & strAdvice & ") A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
            " Where A.�������='7' And A.���ID=[1] And A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID" & _
            " And C.������� IN(1,3) And D.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0)<>5" & _
            " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
    End If
    
    '��ȡ�Ƽ۹�ϵ����ҩƷ������ҽ����ļƼ�,�������ҽ���Ƽۣ����Ƽ�,�ֹ��Ƽ۵�ҽ������ȡ
    If blnLoad Then
        '����û�мӲ�λ������������Ҫ��Distinct
        str�����շ� = "Select * From (" & _
                "Select Distinct C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,C.���ÿ���id" & _
                " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                " From �����շѹ�ϵ C,Table(f_Num2list2([3])) D Where C.������ĿID=D.c1" & _
                "      And (C.���ÿ���ID is Null or C.���ÿ���ID = D.c2 And C.������Դ = 1)" & _
                " ) Where Nvl(���ÿ���id, 0) = Top"
                
        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
            " Select A.���,A.�������,C.��� as �շ����,B.�շ���ĿID as �շ�ϸĿID,D.������ĿID," & _
            " Decode(A.����,0,1,A.����)*B.�շ����� as ����,Decode(C.�Ƿ���,1,D.ȱʡ�۸�,D.�ּ�) as ����," & _
            " C.�Ƿ���,C.���ηѱ�,0 as ��������,A.ִ�п���ID,A.��������,D.�����շ���" & _
            " From (" & strAdvice & ") A,(" & str�����շ� & ") B,�շ���ĿĿ¼ C,�շѼ�Ŀ D,������ĿĿ¼ E,��Ѫ������ F" & _
            " Where A.������� Not IN('4','5','6','7') And A.������ĿID=B.������ĿID" & _
            " And (A.���ID is Null And A.ִ�б�� IN(1,2) And B.��������=1" & _
            "       Or A.�걾��λ=B.��鲿λ And A.��鷽��=B.��鷽�� And Nvl(B.��������,0)=0" & _
            "       Or (A.��鷽�� is Null Or e.��� = 'E' And e.��������='4') And Nvl(B.��������,0)=0 And B.��鲿λ is Null And B.��鷽�� is Null)" & _
            " And A.������ĿID=E.ID And E.�Թܱ���=F.����(+)" & _
            "   And (Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And B.�շ���ĿID=F.����ID" & _
            "       Or Not(Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And F.����ID Is Not NULL))" & _
            " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ���ĿID=C.ID And B.�շ���ĿID=D.�շ�ϸĿID" & _
            " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
            " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) And C.������� IN(1,3)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) And (A.ID=[1] Or A.ID=[2] Or A.���ID=[1])"
    End If
    
    strSQL = "Select /*+ RULE */ A.* From (" & strSQL & ") A Order by ���,������ĿID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.RowData(lngRow)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)), str��Ŀ����)
    If Not rsTmp.EOF Then
        '��ȡ����ID�����ID
        With vsDiag
            lng��ID = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_���ID)) = 0, vsAdvice.RowData(lngRow), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)))
            For i = 1 To .Rows - 1
                If InStr("," & .TextMatrix(i, colҽ��ID) & ",", "," & lng��ID & ",") > 0 Then
                    lng����ID = Val(.TextMatrix(i, col����ID))
                    lng���ID = Val(.TextMatrix(i, col���ID))
                    Exit For
                End If
            Next
        End With
    
        '��ʼ����¼��
        Set mrsPrice = New ADODB.Recordset
        mrsPrice.Fields.Append "����ID", adBigInt
        mrsPrice.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
        mrsPrice.Fields.Append "�շ����", adVarChar, 1
        mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt
        mrsPrice.Fields.Append "����", adDouble
        mrsPrice.Fields.Append "����", adDouble
        mrsPrice.Fields.Append "ʵ�ս��", adDouble
        mrsPrice.Fields.Append "������", adVarChar, 100, adFldIsNullable
        mrsPrice.Fields.Append "��������", adVarChar, 100, adFldIsNullable
        mrsPrice.Fields.Append "����ID", adBigInt, , adFldIsNullable
        mrsPrice.Fields.Append "���ID", adBigInt, , adFldIsNullable
        mrsPrice.CursorLocation = adUseClient
        mrsPrice.LockType = adLockOptimistic
        mrsPrice.CursorType = adOpenStatic
        mrsPrice.Open
        
        '���������ϸ
        dblʵ�� = 0: blnDo = True
        Do While Not rsTmp.EOF
            'ִ�п���
            If blnDo Then
                lngִ�п���ID = Nvl(rsTmp!ִ�п���ID, 0)
                If rsTmp!�շ���� = "4" And Nvl(rsTmp!��������, 0) = 1 Or InStr(",5,6,7,", rsTmp!�շ����) > 0 And InStr(",5,6,7,", rsTmp!�������) = 0 Then
                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!�շ����, rsTmp!�շ�ϸĿID, 4, mlng���˿���id, 0, 2, lngִ�п���ID)
                End If
            End If
            
            '����
            If InStr(",5,6,7,", rsTmp!�շ����) > 0 And Nvl(rsTmp!�Ƿ���, 0) = 1 Then
                'ҩƷʱ��
                dbl���� = Format(CalcDrugPrice(rsTmp!�շ�ϸĿID, lngִ�п���ID, Nvl(rsTmp!����, 0)), gstrDecPrice)
            ElseIf rsTmp!�շ���� = "4" And Nvl(rsTmp!��������, 0) = 1 And Nvl(rsTmp!�Ƿ���, 0) = 1 Then
                '������������ʱ��
                dbl���� = Format(CalcDrugPrice(rsTmp!�շ�ϸĿID, lngִ�п���ID, Nvl(rsTmp!����, 0)), gstrDecPrice)
            Else
                dbl���� = Format(Nvl(rsTmp!����, 0), gstrDecPrice) '�������ȡ��ȱʡ�۸�
            End If
            
            '���
            dbl��� = CCur(Nvl(rsTmp!����, 0) * dbl����)
            If Nvl(rsTmp!��������, 0) = 1 Then
                dbl��� = dbl��� * Nvl(rsTmp!�����շ���, 100) / 100
            End If
            dbl��� = Format(dbl���, gstrDec)
            
            If Nvl(rsTmp!���ηѱ�, 0) = 0 And mstr�ѱ� <> "" Then
                dbl��� = ActualMoney(mstr�ѱ�, rsTmp!������ĿID, dbl���, rsTmp!�շ�ϸĿID, lngִ�п���ID, Nvl(rsTmp!����, 0))
            End If
            
            dblʵ�� = dblʵ�� + dbl���
            
            '��Ŀ�仯ʱ����
            str��Ŀ = rsTmp!��� & "," & rsTmp!�շ�ϸĿID
            blnDo = False: rsTmp.MoveNext
            If Not rsTmp.EOF Then
                If rsTmp!��� & "," & rsTmp!�շ�ϸĿID <> str��Ŀ Then blnDo = True
            Else
                blnDo = True
            End If
            rsTmp.MovePrevious
            
            If blnDo Then
                With vsAdvice
                    mrsPrice.AddNew
                    mrsPrice!����ID = mlng����ID
                    mrsPrice!��ҳID = Null
                    mrsPrice!�շ���� = rsTmp!�շ����
                    mrsPrice!�շ�ϸĿID = rsTmp!�շ�ϸĿID
                    mrsPrice!���� = Nvl(rsTmp!����, 0)
                    mrsPrice!���� = dbl����
                    mrsPrice!ʵ�ս�� = dbl���
                    If .TextMatrix(lngRow, COL_����ҽ��) <> "" Then
                        mrsPrice!������ = .TextMatrix(lngRow, COL_����ҽ��)
                    End If
                    If Val(.TextMatrix(lngRow, COL_��������ID)) <> 0 Then
                        mrsPrice!�������� = CStr(GetItemField("���ű�", Val(.TextMatrix(lngRow, COL_��������ID)), "����"))
                    End If
                    If lng����ID <> 0 Then mrsPrice!����id = lng����ID
                    If lng���ID <> 0 Then mrsPrice!���id = lng���ID
                    mrsPrice.Update
                End With
                dblʵ�� = 0
            End If
            
            rsTmp.MoveNext
        Loop
        
        mrsPrice.MoveFirst
        MakePriceRecord = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDrugStock(ByVal lngRow As Long)
'���ܣ����»�ȡָ��ҩƷ�е�ҩƷ���
'������lngRow=���ġ���ҩ�л���ҩ�÷���
'˵�����������ҩ�䷽��,һ���Ի�ȡ�����䷽�е�������ҩ�Ŀ��
    Dim i As Long
    
    With vsAdvice
        If InStr(",4,5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Or Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) = 0 _
                Or .TextMatrix(lngRow, COL_���) = "4" And Val(.TextMatrix(lngRow, COL_��������)) = 0 Then
                .TextMatrix(lngRow, COL_���) = ""
            Else
                .TextMatrix(lngRow, COL_���) = GetStock(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), 1)
            End If
        ElseIf RowIn�䷽��(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" Then
                        If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Or Val(.TextMatrix(i, COL_�շ�ϸĿID)) = 0 Then
                            .TextMatrix(i, COL_���) = ""
                        Else
                            .TextMatrix(i, COL_���) = GetStock(Val(.TextMatrix(i, COL_�շ�ϸĿID)), Val(.TextMatrix(i, COL_ִ�п���ID)), 1)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(0, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
        
        If Col = col_ҽ������ Then Call vsAdvice.AutoSize(col_ҽ������)
    End If
End Sub

Private Sub vsAdvice_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If dtpDate.Visible Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = True
    End If
    If fraAdvice.Tag <> "" Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_��ʾ Then 'Pass
            Cancel = True
        ElseIf Col = COL_��� Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_Click()
    'PASS
    With vsAdvice
        If mblnPass Then
            If gobjPass.zlPassCheck(mobjPassMap) Then
                Call gobjPass.zlPassAdviceMainPoint(mobjPassMap, 1)
            End If
        End If
    
    End With
End Sub

Private Sub vsAdvice_DblClick()
    Dim lngRow As Long, lngCol As Long
    
    With vsAdvice
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
            If lngCol = COL_��� Then
                Call vsAdvice_KeyPress(32)    '�л���Ϲ�����ʾ
            ElseIf lngCol >= .FixedCols And lngCol <= .Cols - 1 Then
                Call vsAdvice_KeyPress(13)    '��λ����Ӧ�ı༭�ؼ�
                'PASS������ҩ���
                If mblnPass Then
                    If gobjPass.zlPassCheck(mobjPassMap) Then
                        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
                    End If
                End If
            ElseIf .MouseCol = COL_F��־ Then
                '��д����
                '##
            End If
        End If
    End With
End Sub

Private Function RowIsLastVisible(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����һ�ɼ���
    Dim i As Long
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            RowIsLastVisible = lngRow = i
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If RowIsLastVisible(Row) Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            lngLeft = COL_���: lngRight = COL_��ʼʱ��
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_����: lngRight = COL_�÷�
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowInһ����ҩ(Row) Then Exit Sub
            If .RowData(Row) = 0 Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(Row - 1, COL_���ID)), lngBegin, lngEnd)
            Else
                Call Getһ����ҩ��Χ(Val(.TextMatrix(Row, COL_���ID)), lngBegin, lngEnd)
            End If
            
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cbsMain.FindControl(, conMenu_Delete, True, True).Execute
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim objEdit As Object
    Dim lngDiag As Long
    
    If KeyAscii = 13 Then
        '��λ����Ӧ�ı༭�ؼ�
        KeyAscii = 0
        Select Case vsAdvice.Col
            Case COL_��ʼʱ��
                If txt��ʼʱ��.TabStop Then
                    Set objEdit = txt��ʼʱ�� 'ȱʡ����λ����ʼʱ��
                Else
                    Set objEdit = txtҽ������
                End If
            Case col_ҽ������
                Set objEdit = txtҽ������
            Case COL_�걾��λ
                Set objEdit = txt����ʱ��
            Case COL_����
                Set objEdit = txt����
            Case COL_����
                Set objEdit = txt����
            Case COL_����
                Set objEdit = txt����
            Case COL_�÷�
                Set objEdit = txt�÷�
            Case COL_Ƶ��
                Set objEdit = txtƵ��
            Case COL_ִ��ʱ��
                Set objEdit = cboִ��ʱ��
            Case COL_ִ�п���ID
                Set objEdit = cboִ�п���
            Case COL_ҽ������
                Set objEdit = cboҽ������
            Case COL_��־
                Set objEdit = chk����
            Case COL_����˵��
                Set objEdit = txt����˵��
            Case COL_��ҩĿ��
                Set objEdit = cboDruPur
            Case COL_��ҩ����
                Set objEdit = txt��ҩ����
        End Select
        If Not objEdit Is Nothing Then
            If objEdit.Enabled And objEdit.Visible Then objEdit.SetFocus
        End If
    ElseIf KeyAscii = 32 Then
        '�л���Ϲ���
        With vsAdvice
            If .Col = COL_��� And .RowData(.Row) <> 0 And vsDiag.TextMatrix(vsDiag.Row, col���) <> "" Then
                KeyAscii = 0
                Call SetDiagFlag(.Row, IIF(AdviceHaveDiag(.Row) = vsDiag.Row, 0, 1))
            End If
        End With
    End If
End Sub

Private Sub ClearItemTag()
'���ܣ�����ؼ��༭��־
    txt��ʼʱ��.Tag = ""
    txt����ʱ��.Tag = ""
    txt����.Tag = ""
    txt����.Tag = ""
    txt����.Tag = ""
    txt�÷�.Tag = ""
    txtƵ��.Tag = ""
    cboִ��ʱ��.Tag = ""
    cboҽ������.Tag = ""
    cboִ�п���.Tag = ""
    cboִ������.Tag = ""
    cbo����ִ��.Tag = ""
    chk����.Tag = ""
    chk����.Tag = ""
    cbo����.Tag = ""
    chkZeroBilling.Tag = ""
    lbl��ҩĿ��.Tag = ""
    txt��ҩ����.Tag = ""
    txt����˵��.Tag = ""
    lbl����˵��.Tag = ""
End Sub

Private Sub SetStartTime(ByVal Editable As Boolean)
'���ܣ����ÿ�ʼʱ���Ƿ�����༭
    'txt��ʼʱ��.TabStop = Editable 'ȱʡ����λ����ʼʱ��
    txt��ʼʱ��.Locked = Not Editable
    cmd��ʼʱ��.Enabled = Editable
    If Editable Then
        txt��ʼʱ��.BackColor = vsAdvice.BackColor
    Else
        txt��ʼʱ��.BackColor = &HE0E0E0
    End If
End Sub

Private Sub SetDayState(Optional ByVal intVisible As Integer, Optional ByVal intEnabled As Integer)
'���ܣ�����ִ���������úͻ��״̬
'������0-���ֲ���,-1-��ֹ,1-����
    If intEnabled = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
    ElseIf intEnabled = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If
    
    If intVisible = -1 Then
        lbl����.Visible = False
        txt����.Visible = False
        txt����.Text = ""
        
        lbl����.Left = lbl�÷�.Left + lbl�÷�.Width - lbl����.Width
        txt����.Left = txt�÷�.Left
        txt����.Width = txt�÷�.Width - cmd�÷�.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        lbl����.Left = lblƵ��.Left + lblƵ��.Width - lbl����.Width
        txt����.Left = txtƵ��.Left
        txt����.Width = txtƵ��.Width - cmdƵ��.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        txt����.TabIndex = cmdƵ��.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
    ElseIf intVisible = 1 Then
        lbl����.Visible = True
        txt����.Visible = True
        
        lbl����.Left = lbl�÷�.Left + lbl�÷�.Width - lbl����.Width
        txt����.Left = txt�÷�.Left
        txt����.Width = txt�÷�.Width - txt����.Width - Me.TextWidth("������!") - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        lbl����.Left = lblƵ��.Left + lblƵ��.Width - lbl����.Width
        txt����.Left = txtƵ��.Left
        txt����.Width = txtƵ��.Width - cmdƵ��.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        txt����.TabIndex = cmdƵ��.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
    End If
End Sub

Private Sub SetItemEditable(Optional int���� As Integer, Optional int���� As Integer, _
    Optional int�÷� As Integer, Optional intƵ�� As Integer, _
    Optional intִ��ʱ�� As Integer, Optional intִ�п��� As Integer, _
    Optional intִ������ As Integer, Optional int����ִ�� As Integer, _
    Optional int����ʱ�� As Integer, Optional int���� As Integer, _
    Optional int��ҩĿ�� As Integer, Optional int��ҩ���� As Integer, _
    Optional int����˵�� As Integer, Optional int��Ѫԭ�� As Integer)
'���ܣ�����ָ���༭��Ŀ���״̬
'������0-���ֲ���,-1-��ֹ,1-����,2-����
'˵������ֹʱ,ͬʱ�������Ŀ����(����ȫ��)

    '��������Ϊ��ֹʱ,����������ı�,�Ӷ���������Validate�¼�,�����Ƚ�ֹ����˳��
    If int���� = -1 Then txt����.TabStop = False
    If int���� = -1 Then txt����.TabStop = False
    If int�÷� = -1 Then txt�÷�.TabStop = False
    If intƵ�� = -1 Then txtƵ��.TabStop = False
    If intִ��ʱ�� = -1 Then cboִ��ʱ��.TabStop = False
    If intִ�п��� = -1 Then cboִ�п���.TabStop = False
    If intִ������ = -1 Then cboִ������.TabStop = False
    If int����ִ�� = -1 Then cbo����ִ��.TabStop = False
     
    If int��ҩĿ�� = -1 Then cboDruPur.TabStop = False
    If int��ҩ���� = -1 Then txt��ҩ����.TabStop = False
    
    If int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
        lbl������λ.Caption = "" '"��λ"
    ElseIf int���� = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If

    If int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
        lbl������λ.Caption = "" '"��λ"
    ElseIf int���� = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If
    
    If int�÷� = -1 Then
        txt�÷�.Enabled = False
        txt�÷�.BackColor = Me.BackColor
        txt�÷�.Text = ""
        cmd�÷�.Enabled = False
        lbl�÷�.Caption = "�÷�"
    ElseIf int�÷� = 1 Then
        txt�÷�.TabStop = True
        txt�÷�.Enabled = True
        cmd�÷�.Enabled = True
        txt�÷�.BackColor = vsAdvice.BackColor
    End If

    If intƵ�� = -1 Then
        txtƵ��.Enabled = False
        cmdƵ��.Enabled = False
        txtƵ��.BackColor = Me.BackColor
        txtƵ��.Text = ""
    ElseIf intƵ�� = 1 Then
        txtƵ��.TabStop = True
        txtƵ��.Enabled = True
        cmdƵ��.Enabled = True
        txtƵ��.BackColor = vsAdvice.BackColor
    End If

    If intִ��ʱ�� = -1 Then
        cboִ��ʱ��.Text = ""
        cboִ��ʱ��.Enabled = False
        cboִ��ʱ��.BackColor = Me.BackColor
        cboִ��ʱ��.Clear
    ElseIf intִ��ʱ�� = 1 Then
        cboִ��ʱ��.TabStop = True
        cboִ��ʱ��.Enabled = True
        cboִ��ʱ��.BackColor = vsAdvice.BackColor
    End If

    If intִ�п��� = -1 Then
        lblִ�п���.Caption = "ִ�п���"
        cboִ�п���.Enabled = False
        cboִ�п���.BackColor = Me.BackColor
        cboִ�п���.Clear
    ElseIf intִ�п��� = 1 Then
        lblִ�п���.Caption = "ִ�п���"
        cboִ�п���.TabStop = True
        cboִ�п���.Enabled = True
        cboִ�п���.BackColor = vsAdvice.BackColor
    End If

    If intִ������ = -1 Then
        cboִ������.Enabled = False
        cboִ������.BackColor = Me.BackColor
        Call zlControl.CboSetIndex(cboִ������.hWnd, -1) '�����
    ElseIf intִ������ = 1 Then
        cboִ������.TabStop = True
        cboִ������.Enabled = True
        cboִ������.BackColor = vsAdvice.BackColor
    End If
    
    If int����ִ�� = -1 Then
        lbl����ִ��.Caption = "����ִ��"
        cbo����ִ��.Enabled = False
        cbo����ִ��.BackColor = Me.BackColor
        cbo����ִ��.Clear
    ElseIf int����ִ�� = 1 Then
        lbl����ִ��.Caption = "����ִ��"
        cbo����ִ��.TabStop = True
        cbo����ִ��.Enabled = True
        cbo����ִ��.BackColor = vsAdvice.BackColor
    End If
    
    If int����ʱ�� = -1 Then
        txt����ʱ��.Text = ""
        lbl����ʱ��.Visible = False
        txt����ʱ��.Visible = False
        cmd����ʱ��.Visible = False
        lbl�÷�.Visible = True
        txt�÷�.Visible = True
        cmd�÷�.Visible = True
    ElseIf int����ʱ�� = 1 Then
        lbl����ʱ��.Visible = True
        txt����ʱ��.Visible = True
        cmd����ʱ��.Visible = True
        lbl�÷�.Visible = False
        txt�÷�.Visible = False
        cmd�÷�.Visible = False
    End If
    
    If int���� = -1 Then
        cbo����.Text = ""
        lbl����.Visible = False
        cbo����.Visible = False
        lbl���ٵ�λ.Visible = False
    ElseIf int���� = 1 Then
        lbl����.Visible = True
        cbo����.Visible = True
        lbl���ٵ�λ.Visible = True
    End If
    
    'ȱʡ��ѡ��
    If int��ҩĿ�� = -1 Then
        Call zlControl.CboSetIndex(cboDruPur.hWnd, 0)
        cboDruPur.Enabled = True
    ElseIf int��ҩĿ�� = 1 Then
        Call zlControl.CboSetIndex(cboDruPur.hWnd, 0)
        cboDruPur.Enabled = True
        cboDruPur.TabStop = True
    End If
        
    If int��ҩ���� = -1 Then
        txt��ҩ����.Text = ""   'û�����أ�������Ҫ���
        txt��ҩ����.Enabled = False
        txt��ҩ����.BackColor = Me.BackColor
        cmd�ղ���ҩ����.Enabled = False
        cmdReason.Enabled = False
    ElseIf int��ҩ���� = 1 Then
        txt��ҩ����.Enabled = True
        txt��ҩ����.TabStop = True
        txt��ҩ����.BackColor = vsAdvice.BackColor
        cmd�ղ���ҩ����.Enabled = True
        cmdReason.Enabled = True
    End If
    
    If int����˵�� = -1 Then
        txt����˵��.Text = ""   'û�����أ�������Ҫ���
        txt����˵��.Enabled = False
        txt����˵��.BackColor = Me.BackColor
        cmdExcReason.Enabled = False
        cmdComExcReason.Enabled = False
    ElseIf int����˵�� = 1 Then
        txt����˵��.Enabled = True
        txt����˵��.TabStop = True
        txt����˵��.BackColor = vsAdvice.BackColor
        cmdExcReason.Enabled = True
        cmdComExcReason.Enabled = True
    End If
    
    '=1��Ѫҽ������ʾ��Ѫԭ��
    If int��Ѫԭ�� = -1 Then
        lbl��ҩĿ��.Visible = True
        cboDruPur.Visible = True
        cmdReason.Visible = True
        cmd�ղ���ҩ����.Visible = True
        lbl��ҩ����.Caption = "��ҩ����"
    ElseIf int��Ѫԭ�� = 1 Then
        lbl��ҩĿ��.Visible = False
        cboDruPur.Visible = False
        cmdReason.Visible = False
        cmd�ղ���ҩ����.Visible = False
        lbl��ҩ����.Caption = "��Ѫԭ��"
    End If
End Sub

Private Function Get��ͬ��λID() As Long
'���ܣ���ȡ��ǰ���˵ĺ�ͬ��λID
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(��ͬ��λID,0) as ��ͬ��λID From ������Ϣ Where ����ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    
    If rsTmp.RecordCount > 0 Then Get��ͬ��λID = rsTmp!��ͬ��λID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Get�������(ByVal lng���ID As Long, ByVal lng����ID As Long) As String
'���ܣ��������ID�򼲲�ID��ȡ�ֵ���е����ƣ�������ϼ�¼�е����ƿ������޸ĺ��,�����ǰ׺���׺�����Ա��ٴ��޸�ʱ�ж�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If lng���ID <> 0 Then
        strSQL = "Select ���� From �������Ŀ¼ Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng���ID)
        If rsTmp.RecordCount > 0 Then Get������� = "" & rsTmp!����
    ElseIf lng����ID <> 0 Then
        strSQL = "Select ���� From ��������Ŀ¼ Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng����ID)
        If rsTmp.RecordCount > 0 Then Get������� = "" & rsTmp!����
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiInfo() As Boolean
'���ܣ���ȡ������Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim rsSub As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    Dim blnMsgOk As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select A.ID,a.����,a.����,A.����,A.�Ա�,A.����,B.��������,B.�����,B.�ѱ�,B.ҽ�Ƹ��ʽ," & _
        " Nvl(D.Ԥ�����,0)-Nvl(D.�������,0) as Ԥ����,B.����,B.��������,A.�Ǽ�ʱ��," & _
        " A.ִ�в���ID as ���˿���ID,b.���֤��" & _
        " From ���˹Һż�¼ A,������Ϣ B,������� D" & _
        " Where A.NO=[1] And a.��¼����=1 And a.��¼״̬=1 And A.����ID+0=[2]" & _
        " And A.����ID=B.����ID And B.����ID=D.����ID(+) And D.����(+)=1 And D.����(+) = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�, mlng����ID)
    
    lblPati.Caption = _
        "������" & rsTmp!���� & "���Ա�" & Nvl(rsTmp!�Ա�) & "�����䣺" & Nvl(rsTmp!����) & _
        "���ѱ�" & Nvl(rsTmp!�ѱ�) & "�� ҽ�Ƹ��ʽ��" & Nvl(rsTmp!ҽ�Ƹ��ʽ) & "��Ԥ���" & Format(Nvl(rsTmp!Ԥ����, 0), "0.00")
    lblPati.Tag = rsTmp!���� '����ҽ��������ʾ
    
    '���˵�׼ȷ����:�����ж�
    mlng�Һ�ID = rsTmp!ID
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            mdbl����� = Val(rsTmp!����� & "")
        End If
    End If
    mint���� = GetPatiYear(mlng����ID)
    If IsNull(rsTmp!��������) Then
        mDat�������� = DateAdd("yyyy", -mint����, zlDatabase.Currentdate)
    Else
        mDat�������� = rsTmp!��������
    End If
    mstr���� = rsTmp!����
    mstr���֤�� = "" & rsTmp!���֤��
    
    mstr�Ա� = Nvl(rsTmp!�Ա�)
    mstr�ѱ� = Nvl(rsTmp!�ѱ�)
    mdat�Һ�ʱ�� = rsTmp!�Ǽ�ʱ��
    mlng���˿���id = rsTmp!���˿���id
    mstr������ = Getҽ�Ƹ�����(Nvl(rsTmp!ҽ�Ƹ��ʽ))
    mbln��ҽ = Have��������(rsTmp!���˿���id, "��ҽ��")
    mbytPatiType = IIF(Val(Nvl(rsTmp!����)) = 0, 1, 2)
    mbln���� = Val(rsTmp!���� & "") <> 0
    
    'PASS ���˲�����Ϣ
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            Call zlPASSPati
        End If
    End If
    '���ղ����ú�ɫ��ʾ
    mint���� = 0
    If Not IsNull(rsTmp!����) Then
        mint���� = rsTmp!����
        lblPati.ForeColor = vbRed
    End If

    mbln���Ѷ��� = True
    
    '��϶�Ӧҽ��
    strSQL = "Select B.���ID,B.ҽ��ID From ������ϼ�¼ A,�������ҽ�� B" & _
        " Where A.ID=B.���ID And A.��¼��Դ=3 And A.������� IN(1,11) And A.ȡ��ʱ�� is Null And A.����ID=[1] And A.��ҳID=[2]"
    Set rsSub = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    
    '���������Ϣ
    Set vsDiag.Cell(flexcpPicture, 0, col��־) = img16.ListImages("���_��ǰ").Picture
    vsDiag.Cell(flexcpPictureAlignment, 0, col��־) = 4
    
    strSQL = "Select A.ID,A.��¼��Դ,A.��ϴ���,A.�������,A.����ID,A.���ID,A.֤��ID,A.�������,A.�Ƿ�����, a.��¼����, a.��¼��,B.���� as ICD��,c.���� as ��ϱ���,d.���� as ֤�����,A.����ʱ��," & _
        " b.���� As ��������, b.��� As �������,d.���� As ֤������ From ������ϼ�¼ A,��������Ŀ¼ B, �������Ŀ¼ C,��������Ŀ¼ D" & _
        " Where A.����ID=B.ID(+)  And a.���id = c.Id(+) And  a.֤��ID=d.ID(+)  And A.��¼��Դ IN(1,3) And A.������� IN(1,11)" & _
        " And A.ȡ��ʱ�� is Null And A.����ID=[1] And A.��ҳID=[2]" & _
        " Order by A.�������,A.��ϴ���,a.�������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    If Not mclsMipModule Is Nothing Then
        blnMsgOk = mclsMipModule.IsConnect
    End If
    If blnMsgOk Then
        Set mrs��� = New ADODB.Recordset
        mrs���.Fields.Append "id", adBigInt
        mrs���.Fields.Append "��ʾ����", adVarChar, 120
        mrs���.Fields.Append "��ϱ���", adVarChar, 120
        mrs���.Fields.Append "��������", adVarChar, 120
        mrs���.Fields.Append "״̬", adBigInt '0-ԭʼ��¼��1-���޸ģ�2-�����ļ�¼
        mrs���.CursorLocation = adUseClient
        mrs���.LockType = adLockOptimistic
        mrs���.CursorType = adOpenStatic
        mrs���.Open
    End If
    
    If Not rsTmp.EOF Then
        '��ҽ���
        rsTmp.Filter = "��¼��Դ=3 " & IIF(Not mbln��ҽ, " And �������=1", "") '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" & IIF(Not mbln��ҽ, " And �������=1", "") '������Դ����Ϊȱʡ��ʾ
        With vsDiag
            Set mrsDiag = zlDatabase.CopyNewRec(rsTmp)
            If Not rsTmp.EOF Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    Call SetDiagType(i, rsTmp!�������)
                    
                    If IsNull(rsTmp!�������) Then
                        .TextMatrix(i, col����) = ""
                        .TextMatrix(i, col���) = ""
                    Else
                        If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���id & "") = 0 And Val(rsTmp!����id & "") = 0) Then '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                            '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                            If Val(rsTmp!����id & "") <> 0 Then
                                .TextMatrix(i, col����) = Nvl(rsTmp!ICD��)
                            ElseIf Val(rsTmp!���id & "") <> 0 Then
                                .TextMatrix(i, col����) = Nvl(rsTmp!��ϱ���)
                            Else
                                .TextMatrix(i, col����) = ""
                            End If
                            .TextMatrix(i, col���) = rsTmp!�������
                        Else
                            .TextMatrix(i, col����) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                            .TextMatrix(i, col���) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                        End If
                    End If
                    
                    'ȡ֤������
                    If InStr(.TextMatrix(i, col���), "(") > 0 And InStr(.TextMatrix(i, col���), ")") > 0 And Val(rsTmp!������� & "") = 11 Then
                        strTmp = Mid(.TextMatrix(i, col���), InStrRev(.TextMatrix(i, col���), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '��ȡ֤��
                        .TextMatrix(i, col��ҽ֤��) = strTmp
                        'ȥ�����������֤��
                        .TextMatrix(i, col���) = Mid(.TextMatrix(i, col���), 1, InStrRev(.TextMatrix(i, col���), "(") - 1)
                    Else
                       .TextMatrix(i, col��ҽ֤��) = ""
                    End If
                    If Not IsNull(rsTmp!����id) Or Not IsNull(rsTmp!���id) Then
                        .Cell(flexcpData, i, col���) = Get�������(Val("" & rsTmp!���id), Val("" & rsTmp!����id))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                    Else
                        .Cell(flexcpData, i, col���) = .TextMatrix(i, col���)
                    End If
                    
                    .Cell(flexcpData, i, col����) = Val(Nvl(rsTmp!�Ƿ�����, 0))
                    .Cell(flexcpForeColor, i, col����) = IIF(Nvl(rsTmp!�Ƿ�����, 0) = 1, vbRed, .GridColor)
                    .TextMatrix(i, col���ID) = Nvl(rsTmp!���id, 0)
                    .Cell(flexcpData, i, col���ID) = Nvl(rsTmp!ID, 0)
                    .TextMatrix(i, col����ID) = Nvl(rsTmp!����id, 0)
                    .TextMatrix(i, col֤��ID) = Nvl(rsTmp!֤��id, 0)
                    .TextMatrix(i, colICD��) = Nvl(rsTmp!ICD��)
                    .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                  
                    .TextMatrix(i, col��ϱ���) = Nvl(rsTmp!��ϱ���)
                    .TextMatrix(i, col��������) = Nvl(rsTmp!ICD��)
                    .TextMatrix(i, col�������) = Nvl(rsTmp!�������)
                    .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                    .TextMatrix(i, col֤�����) = Nvl(rsTmp!֤�����)
                    
                    If blnMsgOk Then
                        mrs���.AddNew
                        mrs���!ID = rsTmp!ID
                        mrs���!��ʾ���� = .TextMatrix(i, col����)
                        mrs���!��ϱ��� = .TextMatrix(i, col��ϱ���)
                        mrs���!�������� = .TextMatrix(i, col��������)
                        mrs���!״̬ = 0
                        mrs���.Update
                    End If
                    
                    '��д��Ӧ�Ĺ���ҽ��
                    strSQL = ""
                    rsSub.Filter = "���ID=" & rsTmp!ID
                    Do While Not rsSub.EOF
                        strSQL = strSQL & "," & rsSub!ҽ��ID
                        rsSub.MoveNext
                    Loop
                    .TextMatrix(i, colҽ��ID) = Mid(strSQL, 2)
                    rsTmp.MoveNext
                Next
            End If
        End With
    Else
        Call SetDiagType(1, IIF(mbln��ҽ, 11, 1))
    End If
    
    '��ҽʱֻ������ҽ���
    vsDiag.ColHidden(col��ҽ) = Not mbln��ҽ
    vsDiag.ColHidden(COL��ҽ) = Not mbln��ҽ
    vsDiag.ColHidden(col��ҽ֤��) = Not mbln��ҽ
    vsDiag.ColWidth(col���) = IIF(mbln��ҽ, 2760, 4360)
    vsDiag.ColHidden(COLDEL) = False
    vsDiag.ColWidth(COLDEL) = vsDiag.ColWidth(col����)
 
    vsDiag.Col = col���: vsDiag.Row = vsDiag.Rows - 1
    Call vsDiag_AfterRowColChange(-1, -1, vsDiag.Row, vsDiag.Col)
    Call SetDiagHeight
    'PASS��ϴ���
    If mblnPass Then
        zlPassDrags
    End If
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetDiagType(ByVal lngRow As Long, ByVal int������� As Integer)
'���ܣ�����ĳһ����е��������
    With vsDiag
        .TextMatrix(lngRow, col��ҽ) = "��"
        .TextMatrix(lngRow, COL��ҽ) = "��"
        
        .Cell(flexcpData, lngRow, col��ҽ) = IIF(int������� = 11, 1, 0)
        .Cell(flexcpForeColor, lngRow, col��ҽ) = IIF(int������� = 11, .ForeColor, .GridColor)
        .Cell(flexcpFontBold, lngRow, col��ҽ) = IIF(int������� = 11, True, False)
        
        .Cell(flexcpData, lngRow, COL��ҽ) = IIF(int������� = 1, 1, 0)
        .Cell(flexcpForeColor, lngRow, COL��ҽ) = IIF(int������� = 1, .ForeColor, .GridColor)
        .Cell(flexcpFontBold, lngRow, COL��ҽ) = IIF(int������� = 1, True, False)
    End With
End Sub

Private Function GetPreRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ�����Ч�ɼ���
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ�����Ч�ɼ���
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetNextRow = lngTmp
End Function

Private Function GetDefaultTime(lngRow As Long) As String
'���ܣ���ȡ�¿�ҽ����ȱʡ��ʼʱ��
'˵����
'      ���һ����Чʱ��Ϊ���죬�Ҽ�������ڲ�¼�������
'      ���û��,��ȡ����¿�һ����ʱ��
'      ���û��,��ȡ��ǰʱ��
    Dim curDate As Date, strDate As String, i As Long
    
    curDate = zlDatabase.Currentdate
    
    With vsAdvice
        '�ȴӵ�ǰ�������
        For i = lngRow - 1 To .FixedRows Step -1
            If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) Then
                If Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                    If DateAdd("n", gint�����¿�ҽ�����, CDate(.Cell(flexcpData, i, COL_��ʼʱ��))) >= curDate Then
                        strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                        Exit For
                    End If
                End If
            End If
        Next
            
        '�ٴ�����������
        If strDate = "" Then
            For i = .Rows - 1 To lngRow + 1 Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) Then
                    If Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                        If DateAdd("n", gint�����¿�ҽ�����, CDate(.Cell(flexcpData, i, COL_��ʼʱ��))) >= curDate Then
                            strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        
        If strDate = "" Then
            '�ȴӵ�ǰ�������
            For i = lngRow - 1 To .FixedRows Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) _
                    And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                    strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                    Exit For
                End If
            Next
            '�ٴ�����������
            If strDate = "" Then
                For i = .Rows - 1 To lngRow + 1 Step -1
                    If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) _
                        And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    If strDate = "" Then strDate = Format(curDate, "yyyy-MM-dd HH:mm")
    GetDefaultTime = strDate
End Function

Private Function GetCurRow���(lngRow As Long) As Long
'���ܣ���ȡָ���п��õĵ����
'������lngRow=Ҫȡ��ŵ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng��� As Long, i As Long
    Dim lng���1 As Long, lng���2 As Long
            
    'ȡ֮�����һ����Ч���,ֱ��ʹ��
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                lng��� = Val(vsAdvice.TextMatrix(i, COL_���))
                Exit For
            End If
        End If
    Next
    If lng��� = 0 Then
        '����û��,��ȡ���ݿ�֮�е���������֮ǰ�������űȽ�
        On Error GoTo errH
        strSQL = "Select Max(���) as ��� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And Nvl(Ӥ��,0)=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, cboӤ��.ListIndex)
        If Not rsTmp.EOF Then lng���1 = Nvl(rsTmp!���, 0)
        On Error GoTo 0
        
        For i = lngRow - 1 To vsAdvice.FixedRows Step -1
            If vsAdvice.RowData(i) <> 0 Then
                If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex _
                    And IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                    lng���2 = Val(vsAdvice.TextMatrix(i, COL_���))
                    Exit For
                End If
            End If
        Next
        
        If lng���1 > lng���2 Then
            lng��� = lng���1
        Else
            lng��� = lng���2
        End If

        If lng��� <> 0 Then lng��� = lng��� + 1 '������+1
    End If
    If lng��� = 0 Then lng��� = 1
    GetCurRow��� = lng���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSetҽ�����(lngRow As Long, intStep As Integer)
'���ܣ�����ǰ����ҽ����¼�����ǰ�ƻ����
'������lngRow=��ʼ������,intStep=��������,��1��-1
    Dim i As Long
    
    For i = lngRow To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                vsAdvice.TextMatrix(i, COL_���) = Val(vsAdvice.TextMatrix(i, COL_���)) + intStep
                If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 0 Then
                    vsAdvice.TextMatrix(i, COL_EDIT) = 3 '��־�޸������
                End If
            End If
        End If
    Next
End Sub

Private Sub AdviceDelete(ByVal lngRow As Long)
'���ܣ�ָ����ҽ��ɾ������
    Dim lngBegin As Long, lngEnd As Long
    Dim lng���ID As Long, blnGroup As Boolean
    Dim lngҽ��ID As Long, i As Integer
    Dim lngDiag As Long, lng���״̬ As Long
    
    lngDiag = -1
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
    
    If vsAdvice.RowData(lngRow) <> 0 Then
        '����ɾ��ǰ��ҽӿ�
        On Error Resume Next
        If Val(vsAdvice.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
            If CreatePlugInOK(p����ҽ���´�, mint����) Then
                If gobjPlugIn.AdviceDeletBefor(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(vsAdvice.RowData(lngRow)), mint����) = False Then
                    If err.Number = 0 Then Exit Sub
                End If
                Call zlPlugInErrH(err, "AdviceDeletBefor")
            End If
        End If
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
        
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
            lngҽ��ID = vsAdvice.RowData(lngRow)
            lng���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
            blnGroup = RowInһ����ҩ(lngRow)
            If blnGroup Then
                '��ɾ��һ����ҩ�еĿ���(һ��Ҫɾ)
                Call Getһ����ҩ��Χ(lng���ID, lngBegin, lngEnd)
                For i = lngEnd To lngBegin Step -1 '���뷴��
                    If vsAdvice.RowData(i) = 0 Then Call DeleteRow(i)
                Next
                
                'ɾ��֮��ǰ�кſ��ܱ���
                lngRow = vsAdvice.FindRow(lngҽ��ID, lngBegin)
                    
                'һ����ҩֻɾ����ǰ��
                lngDiag = AdviceHaveDiag(lngRow) '��¼��һ�еĹ������
                lng���״̬ = Val(vsAdvice.TextMatrix(lngRow, COL_���״̬))
                Call DeleteRow(lngRow)
                
                If lng���״̬ <> 0 Then
                    lngRow = vsAdvice.FindRow(lng���ID, lngBegin)
                    Call ReSet���״̬ͼ��(lngRow)
                End If
            Else
                '�����ĳ�ҩ��ɾ����ҩ;���м���ǰ��
                i = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow)
            End If
        ElseIf InStr(",D,F,K,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
            Call Delete���������Ѫ(lngRow)
            Call DeleteRow(lngRow)
        ElseIf RowIn�䷽��(lngRow) Then
            'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete��ҩ�䷽(lngRow)
            'ɾ����ǰ��(��ҩ�÷���)
            Call DeleteRow(lngRow)
        ElseIf RowIn������(lngRow) Then
            lngRow = Delete�������(lngRow)
            Call DeleteRow(lngRow)
        Else
            Call DeleteRow(lngRow)
        End If
        
        mblnNoSave = True '���Ϊδ����
    Else
        '����ֱ��ɾ��
        Call DeleteRow(lngRow)
    End If
    
    '���¶�λ��
    If vsAdvice.RowHidden(vsAdvice.Row) Then
        i = GetPreRow(vsAdvice.Row)
        If i = -1 Then i = GetNextRow(vsAdvice.Row)
        If i <> -1 Then vsAdvice.Row = i
    End If
    
    '�ָ�һ����ҩ����Ϲ���
    If lngDiag <> -1 Then
        Call SetDiagFlag(vsAdvice.FindRow(lng���ID), 1, lngDiag) '��ǰ��ҲӦ�ǻָ���һ����ҩ�е�
    End If
    
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    mblnRowChange = True
    vsAdvice.Redraw = flexRDDirect
    Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub DeleteRow(ByVal lngRow As Long, Optional ByVal blnClear As Boolean, Optional blnDelID As Boolean = True)
'���ܣ�ɾ������е�һ��,�����ı䵱ǰ��
'������blnClear=�Ƿ�������������,��ɾ��
'      blnDelID=�Ƿ��¼Ҫɾ����ҽ��ID
    Dim lngCol As Long, blnDraw As Boolean, blnChange As Boolean
    
    With vsAdvice
        lngCol = .Col
        blnDraw = .Redraw
        blnChange = mblnRowChange
        
        mblnRowChange = False
        .Redraw = flexRDNone
        
        If .RowData(lngRow) <> 0 Then
            '�������
            Call AdviceSetҽ�����(lngRow + 1, -1)
            
            '��¼Ҫɾ����ID(���˲�������)
            If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 And blnDelID Then
                If .TextMatrix(lngRow, COL_�������״̬) = "1" Or .TextMatrix(lngRow, COL_�������״̬) = "2" Then
                    If InStr("," & mstrAduitDelIDs & ",", "," & .TextMatrix(lngRow, COL_���ID) & ",") = 0 And .TextMatrix(lngRow, COL_���ID) <> "" Then
                        mstrAduitDelIDs = mstrAduitDelIDs & "," & .TextMatrix(lngRow, COL_���ID)
                        mstrDelIDs = Replace(mstrDelIDs, "," & .TextMatrix(lngRow, COL_���ID), "")
                    End If
                Else
                    mstrDelIDs = mstrDelIDs & "," & .RowData(lngRow)
                End If
                If .TextMatrix(lngRow, COL_���) = "K" And gblnѪ��ϵͳ Then
                    mstrDel��Ѫ = mstrDel��Ѫ & "," & .RowData(lngRow)
                End If
            End If
            
            'ɾ���������ϵı�Ǵ�������������һ��ҽ��ʱ�Ĳ�����ɾ����û��ϵ�����ĺ��������ٱ����
            Call SetDiagFlag(lngRow, 0)
        End If
            
        '���Ϊ��1�ҽ�ʣ��1������,����
        If Not (lngRow = .FixedRows And .Rows = .FixedRows + 1) And Not blnClear Then
            .RemoveItem lngRow
        Else
            '�����������
            .RowData(lngRow) = Empty
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "" '����
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = Empty '����
            .Cell(flexcpFontBold, lngRow, .FixedCols, lngRow, .Cols - 1) = False '����
            .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = .ForeColor '����ɫ
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .FixedCols - 1) = .ForeColorFixed '�̶�������ɫ
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .FixedCols - 1) = .BackColorFixed '�̶��б���ɫ
            Set .Cell(flexcpPicture, lngRow, 0, lngRow, .Cols - 1) = Nothing '��ԪͼƬ
            Set .Cell(flexcpPicture, lngRow, COL_��ʾ) = Nothing 'Pass��ʾ��
            
            '��Ԫ��߿�
            .Select lngRow, .FixedCols, lngRow, COL_��־
            .CellBorder vbRed, 0, 0, 0, 0, 0, 0
        End If
        
        .Col = lngCol '��Ϊ��ɾ����,���Ե��ó���϶����ж�λ,���Բ��ػָ���
        .Redraw = blnDraw
        mblnRowChange = blnChange
    End With
End Sub

Private Sub Delete���������Ѫ(ByVal lngRow As Long, Optional ByVal bln������� As Boolean, Optional ByRef lngTmpRow As Long)
'���ܣ�1.ɾ����������Ŀ�Ĳ�λ��
'      2.ɾ��������Ŀ�ĸ��������м�������Ŀ��
'      3.ɾ����Ѫ��Ŀ����Ѫ;����
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim lngNo As Long
    Dim strRows As String
    Dim varArr As Variant
    Dim lngTmp As Long
    On Error GoTo errH
    With vsAdvice
        If bln������� Then
            lngNo = Val(.TextMatrix(lngRow, COL_�������))
        End If
        If lngNo = 0 Then
            i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_���ID) '��һ����,�����ò���
            If i <> -1 Then
                lngBegin = i
                For i = lngBegin To vsAdvice.Rows - 1
                    If Val(vsAdvice.TextMatrix(i, COL_���ID)) = vsAdvice.RowData(lngRow) Then
                        lngEnd = i
                    Else
                        Exit For
                    End If
                Next
                For i = lngEnd To lngBegin Step -1
                    Call DeleteRow(i)
                Next
            End If
        Else
            lngTmp = -1
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_״̬)) < 2 Then
                    If lngNo = Val(.TextMatrix(i, COL_�������)) Then
                        strRows = i & IIF(strRows = "", "", "," & strRows)
                    End If
                End If
                If i > lngRow Then
                    If lngNo <> Val(.TextMatrix(i, COL_�������)) Then
                        If lngTmp = -1 Then
                            lngTmp = vsAdvice.RowData(i)
                        End If
                    End If
                End If
            Next
            varArr = Split(strRows, ",")
            For i = 0 To UBound(varArr)
                Call DeleteRow(Val(varArr(i)))
            Next
            
            If lngTmp = -1 Then
                '��ɾ���м����Ŀ���
                mblnRowChange = False
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                mblnRowChange = True
                .AddItem "", .Rows
                .Row = .Rows - 1
                .Col = .FixedCols
                lngTmpRow = .Row
            Else
                For i = .FixedRows To .Rows - 1
                    If lngTmp = Val(vsAdvice.RowData(i)) Then
                        lngTmp = i
                        Exit For
                    End If
                Next
                i = lngTmp
                If i <> -1 Then
                    .AddItem "", i
                    .Row = i
                    lngTmpRow = i
                Else
                    '��ɾ���м����Ŀ���
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                    lngTmpRow = .Row
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Delete�������뵥(ByVal lngRow As Long, Optional ByVal bln������� As Boolean, Optional ByRef lngTmpRow As Long)
'���ܣ�1.ɾ����������Ŀ�Ĳ�λ��
'      2.ɾ��������Ŀ�ĸ��������м�������Ŀ��
'      3.ɾ����Ѫ��Ŀ����Ѫ;����
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim lngNo As Long
    Dim strRows As String
    Dim varArr As Variant
    Dim lngTmp As Long
    On Error GoTo errH
    With vsAdvice
        If bln������� Then
            lngNo = Val(.TextMatrix(lngRow, COL_�������))
        End If
        If lngNo = 0 Then
            lngTmpRow = Delete�������(lngRow)
        Else
            lngTmp = -1
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_״̬)) < 2 Then
                    If lngNo = Val(.TextMatrix(i, COL_�������)) Then
                        strRows = i & IIF(strRows = "", "", "," & strRows)
                    End If
                End If
                If i > lngRow Then
                    If lngNo <> Val(.TextMatrix(i, COL_�������)) Then
                        If lngTmp = -1 Then
                            lngTmp = vsAdvice.RowData(i)
                        End If
                    End If
                End If
            Next
            varArr = Split(strRows, ",")
            For i = 0 To UBound(varArr)
                Call DeleteRow(Val(varArr(i)))
            Next
            
            If lngTmp = -1 Then
                '��ɾ���м����Ŀ���
                mblnRowChange = False
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                mblnRowChange = True
                .AddItem "", .Rows
                .Row = .Rows - 1
                .Col = .FixedCols
                lngTmpRow = .Row
            Else
                For i = .FixedRows To .Rows - 1
                    If lngTmp = Val(vsAdvice.RowData(i)) Then
                        lngTmp = i
                        Exit For
                    End If
                Next
                i = lngTmp
                If i <> -1 Then
                    .AddItem "", i
                    .Row = i
                    lngTmpRow = i
                Else
                    '��ɾ���м����Ŀ���
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                    lngTmpRow = .Row
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Delete��ҩ�䷽(ByVal lngRow As Long) As Long
'���ܣ�ɾ����ҩ�䷽�����ζҩ���巨��
'������lngRow=��ҩ�䷽�÷���(�ɼ�)
'���أ�ɾ��֮�����¶�λ�ĵ�ǰ��(��ҩ�÷���)
    Dim lngBegin As Long, lngEnd As Long
    Dim lngҽ��ID As Long, i As Long
    
    lngҽ��ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '��Ϊ����ǰ��ɾ��,��Ҫ���¶�λ����ҩ�÷���
    i = vsAdvice.FindRow(lngҽ��ID)
    vsAdvice.Row = i '�������Ҳ���
    
    mblnRowChange = True
    
    Delete��ҩ�䷽ = vsAdvice.Row
End Function

Private Function Delete�������(ByVal lngRow As Long) As Long
'���ܣ�ɾ��һ���ɼ��Ķ��������Ŀ��
'������lngRow=�ɼ�������(�ɼ�)
'���أ�ɾ��֮�����¶�λ�ĵ�ǰ��(�ɼ�������)
    Dim lngBegin As Long, lngEnd As Long
    Dim lngҽ��ID As Long, i As Long
    
    lngҽ��ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '��Ϊ����ǰ��ɾ��,��Ҫ���¶�λ���ɼ�������
    i = vsAdvice.FindRow(lngҽ��ID)
    vsAdvice.Row = i '�������Ҳ���
    
    mblnRowChange = True
    
    Delete������� = vsAdvice.Row
End Function

Private Function Get��鲿λ����(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еļ�鲿λ������
'������lngRow=���ҽ���Ŀɼ���
'���أ�"��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
'      ������ϵļ����Ϸ�ʽ����������ǰ�ĵ���λ��飬�򷵻ؿ��Ա����ʶ��
    Dim str��λ As String, str��λLast As String
    Dim str���� As String, i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_������ĿID)) <> Val(.TextMatrix(lngRow, COL_������ĿID)) Then Exit Function '�ϵķ�ʽ
                
                If .TextMatrix(i, COL_�걾��λ) <> "" Then
                    If .TextMatrix(i, COL_�걾��λ) <> str��λLast And str��λLast <> "" Then
                        str��λ = str��λ & "|" & str��λLast & IIF(str���� <> "", ";" & Mid(str����, 2), "")
                        str���� = ""
                    End If
                    If .TextMatrix(i, COL_��鷽��) <> "" Then
                        str���� = str���� & "," & .TextMatrix(i, COL_��鷽��)
                    End If
                    
                    str��λLast = .TextMatrix(i, COL_�걾��λ)
                End If
            Else
                Exit For
            End If
        Next
        If str��λLast <> "" Then
            str��λ = str��λ & "|" & str��λLast & IIF(str���� <> "", ";" & Mid(str����, 2), "")
        End If
        Get��鲿λ���� = Mid(str��λ, 2) & vbTab & Val(.TextMatrix(lngRow, COL_ִ�б��))
    End With
End Function

Private Function Get��������IDs(ByVal lngRow As Long) As String
'���ܣ���ȡָ�������еĸ���������������ĿID��
'���أ�"����ID1,����ID2,...;����ID",���п���û�и�������������
    Dim strTmp As String, lng����ID As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_���ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = vsAdvice.RowData(lngRow) Then
                If vsAdvice.TextMatrix(i, COL_���) = "G" Then
                    lng����ID = Val(vsAdvice.TextMatrix(i, COL_������ĿID))
                Else
                    strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_������ĿID))
                End If
            Else
                Exit For
            End If
        Next
    End If
    Get��������IDs = Mid(strTmp, 2) & ";" & IIF(lng����ID = 0, "", lng����ID)
End Function

Private Function Get��ҩ�䷽IDs(ByVal lngRow As Long) As String
'���ܣ���ȡ��ҩ�䷽�����ζҩ���巨ID��
'���أ�"��ҩID1,����1,��ע1;��ҩID2,����2,��ע2;...|�巨ID"
    Dim lng�巨ID As Long, str��ҩIDs As String, i As Long, lng��̬ As Long
    Dim lng���� As Long, lngҩ��ID As Long
    Dim strTmp As String
    
    With vsAdvice
        lng��̬ = Val(.TextMatrix(lngRow, COL_��ҩ��̬))    '�÷���
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_���) = "E" Then
                    lng�巨ID = Val(.TextMatrix(i, COL_������ĿID))
                    strTmp = .TextMatrix(i, COL_�걾��λ) '������ҩ�� ����
                ElseIf .TextMatrix(i, COL_���) = "7" Then
                    str��ҩIDs = Val(.TextMatrix(i, COL_�շ�ϸĿID)) & "," & _
                        .TextMatrix(i, COL_����) & "," & .TextMatrix(i, COL_ҽ������) & _
                        ";" & str��ҩIDs
                    If lngҩ��ID = 0 Then
                        lngҩ��ID = .TextMatrix(i, COL_ִ�п���ID)
                        lng���� = .TextMatrix(i, COL_����)
                    End If
                End If
            Else
                Exit For
            End If
        Next
        Get��ҩ�䷽IDs = Mid(str��ҩIDs, 1, Len(str��ҩIDs) - 1) & "|" & lng�巨ID & "|" & lng��̬ & "|" & lng���� & "|" & lngҩ��ID & "|" & strTmp
    End With
End Function

Private Function Get�������IDs(ByVal lngRow As Long) As String
'���ܣ���ȡһ���ɼ��ļ��������ĿID���걾
'���أ�"'      �������="��ĿID1,��ĿID2,...;����걾" ������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾""
    Dim str��ĿIDs As String, str�걾 As String, i As Long
    Dim j As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_�����ĿID)) = 0 And mblnNewLIS Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_���ID)) = .RowData(lngRow) Then
                            If Val(.TextMatrix(j, COL_�����ĿID)) = Val(.TextMatrix(i, COL_������ĿID)) And Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                                str��ĿIDs = "|" & Val(.TextMatrix(j, COL_������ĿID)) & str��ĿIDs
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    str��ĿIDs = "," & Val(.TextMatrix(i, COL_������ĿID)) & str��ĿIDs
                Else
                    If Not mblnNewLIS Then
                        str��ĿIDs = "," & Val(.TextMatrix(i, COL_������ĿID)) & str��ĿIDs
                    End If
                End If
                str�걾 = .TextMatrix(i, COL_�걾��λ)
            Else
                Exit For
            End If
        Next
    End With
    Get�������IDs = Right(str��ĿIDs, Len(str��ĿIDs) - 1) & ";" & str�걾
End Function

Private Function RowIn������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����ڼ�������е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_���) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
            '�ɼ�������
            If .TextMatrix(lngRow - 1, COL_���) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                RowIn������ = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_���) = "C" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '������Ŀ��
            RowIn������ = True: Exit Function
        End If
    End With
End Function

Private Function RowIn�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�������ҩ�䷽�е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_���) = "E" Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
                '�÷���
                If Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_���) = "E" Then
                    RowIn�䷽�� = True: Exit Function
                End If
            Else
                '�巨��
                If .TextMatrix(lngRow - 1, COL_���) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    RowIn�䷽�� = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_���) = "7" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '��ҩ��
            RowIn�䷽�� = True: Exit Function
        End If
    End With
End Function

Private Function RowInһ����ҩ(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��
'������lngRow=�ɼ�����,�����ǿ���
'˵����һ����ҩ�ķ�Χ�п��ܴ��ڿ���
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lng���ID As Long, blnGroup As Boolean, i As Long
    
    lngPreRow = GetPreRow(lngRow)
    lngNextRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            If lngPreRow <> -1 And lngNextRow <> -1 Then
                If Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(lngNextRow, COL_���ID)) _
                    And Val(.TextMatrix(lngPreRow, COL_���ID)) <> 0 _
                    And InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                    And InStr(",5,6,", .TextMatrix(lngNextRow, COL_���)) > 0 Then
                    blnGroup = True
                End If
            End If
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 _
            And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            
            lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
            If lngPreRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_���ID)) = lng���ID Then blnGroup = True
            End If
            If Not blnGroup And lngNextRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngNextRow, COL_���)) > 0 _
                    And Val(.TextMatrix(lngNextRow, COL_���ID)) = lng���ID Then blnGroup = True
            End If
        End If
    End With
    RowInһ����ҩ = blnGroup
End Function

Private Function AdviceSet��������(ByVal lngRow As Long, ByVal lngƤ��ID As Long) As Boolean
'���ܣ��Զ�����Ƥ����
'������lngRow=��ǰ������,�Ѿ�������ҩ���г�ҩ
'      lngƤ��ID=Ҫ���ӵ�Ƥ����ĿID
'˵�����Զ�����֮��,��ǰ�м�����Զ�λ���Ѹ������ҩƷ��λ��
    Dim rsInput As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
        
    '���ID,����,������ĿID,�շ�ϸĿID,���,����
    strSQL = "Select ��� as ���ID,����,ID as ������ĿID,NULL as �շ�ϸĿID,NULL as ���,NULL as ����,NULL as ��Ŀ���� From ������ĿĿ¼ Where ID=[1]"
    Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngƤ��ID)
        
    'Ѱ��ʵ��Ҫ����Ƥ�Ե���:һ����ҩ�����
    With vsAdvice
        For i = lngRow - 1 To .FixedRows - 1 Step -1 '��������������ǰ��
            If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(lngRow, COL_���ID)) Then
                lngRow = i + 1: Exit For '�����е��к�
            End If
        Next
    End With
    
    '�������
    vsAdvice.AddItem "", lngRow
    
    '����Ƥ��
    Call AdviceInput(rsInput, lngRow, True)
    
    '���¶�λ�������ҩƷ��
    mblnRowChange = False
    vsAdvice.Row = vsAdvice.Row + 1
    mblnRowChange = True
    
    AdviceSet�������� = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset, ByVal lngRow As Long, Optional ByVal blnByƤ�� As Boolean) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��,lngRow=��ǰ������,blnByƤ��=�Ƿ��Զ�����Ƥ������
'���أ�����¼���Ƿ���Ч
    Dim str���� As String, blnGroup As Boolean, i As Long
    Dim lng�÷�ID As Long, lngGroupRow As Long, lngRowID As Long
    Dim lngƤ��ID As Long, blnƤ���� As Boolean
    Dim lngPreRow As Long, lngNextRow As Long
    Dim strExtData As String, strAppend As String
    Dim intType As Integer, strժҪ As String, lng���״̬ As Long, strժҪOut As String
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim objControl As CommandBarControl
    Dim lngPreRow��ҩ�� As Long, lngҩƷID As Long
    Dim blnOK As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    Dim str������λ As String
    Dim strIDs1 As String, strIDs2 As String, strҽ������ As String
    Dim lngBegin As Long, lngEnd As Long, sng���� As Single
    Dim lngAppType As Long '���뵥Ӧ��
    Dim objAppPages()  As clsApplicationData
    Dim rsCard As ADODB.Recordset
    Dim lngTmpRow As Long
    Dim bln��Ѫ As Boolean '�Ƿ�Ϊ��Ѫҽ�� ��Ѫ=0����Ѫ=1,����K���ҽ���е� ��鷽��  �ֶ�;��Ѫ-�ɼ���ʽ / ��Ѫ-��Ѫ;��
    Dim strWhere As String
        Dim lngApplyID As Long
    
    On Error GoTo errH
        
    lngPreRow = GetPreRow(lngRow) 'ȡ��һ��Ч��,ĳЩ����ȱʡ����һ����ͬ
    lngNextRow = GetNextRow(lngRow) 'ȡ��һ��Ч��
    
    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    txtҽ������.Text = rsInput!���� '��ʱ��ʾ
    
    'ҩƷ����ְ����(��ʿվ�ڱ���ʱ���)
    If InStr(",5,6,7,", rsInput!���ID) > 0 Then
        strMsg = CheckOneDuty(rsInput!����, Nvl(rsInput!����ְ��ID), UserInfo.����, InStr(",1,2,", mstr������) > 0 And mstr������ <> "")
        If strMsg <> "" Then
            vsAdvice.Refresh
            MsgBox strMsg, vbInformation, gstrSysName
            vsAdvice.Refresh: Exit Function
        End If
        If mblnPass Then
            If gobjPass.zlPassCheck(mobjPassMap) Then
                Call gobjPass.zlPassAdviceInput(mobjPassMap, Val(rsInput!������ĿID & ""), Val(rsInput!�շ�ϸĿID & ""), rsInput!���� & "")
            End If
        End If
    End If
    With vsAdvice
        '����Ƿ������Ч��סԺҽ��������ҽ��
        If rsInput!���ID = "Z" And InStr(",����,סԺ,", "," & rsInput!��Ŀ���� & ",") > 0 Then
            If CheckInHosAdvice Then
                Exit Function
            End If
        End If
        
        'ҽ��������������ʱ����ʾ����ҽ������ҲҪ��(Or And mint���� <> 0)
        If InStr(",7,8,9,", rsInput!���ID) = 0 Then '���ף��䷽����ζ��ҩ����������ʾ
            strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, Nvl(rsInput!�շ�ϸĿID, 0), "", 0, "", rsInput!������ĿID)
        End If
    
        
        '������Ŀ���ɼ������ж�
        If rsInput!���ID = "C" Then
            '����������ȡһ��ȱʡ�Ĳɼ�����,ͬʱ�ж��Ƿ��вɼ���������
            lng�÷�ID = Getȱʡ�÷�ID(6, 1)
            If lng�÷�ID = 0 Then
                .Refresh
                MsgBox "û�п��õı걾�ɼ�����,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            'ȱʡ����һ����ͬ
            If lngPreRow <> -1 Then
                If RowIn������(lngPreRow) Then
                    If Val(.TextMatrix(lngPreRow, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(lngPreRow, COL_������ĿID))
                End If
            End If
        End If
        
        '��ҩ�䷽����������ҩ�÷��ж�
        If InStr(",7,8,", rsInput!���ID) > 0 Then
            If rsInput!���ID = "8" Then
                If GetGroupCount(rsInput!������ĿID, 1, False) = 0 Then
                    .Refresh
                    MsgBox """" & rsInput!���� & """��һ����ҩ�䷽����û��������Ч�������ҩ��" & vbCrLf & "���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                    .Refresh: Exit Function
                End If
            
                '����ҩ��Ч����ʾ
                strMsg = GetGroupNone(rsInput!������ĿID, 1)
                If strMsg <> "" Then
                    .Refresh
                    MsgBox "�䷽""" & rsInput!���� & """������ҩƷ�ѳ�����������ƥ�䣺" & _
                        vbCrLf & vbCrLf & vbTab & strMsg & vbCrLf & vbCrLf & "��ЩҩƷ������������䷽�С�", vbInformation, gstrSysName
                    .Refresh
                End If
            End If
        
            '����������ȡһ��ȱʡ����ҩ�÷�,ͬʱ�ж��Ƿ�����ҩ�÷�����
            lng�÷�ID = Getȱʡ�÷�ID(4, 1)
            If lng�÷�ID = 0 Then
                .Refresh
                MsgBox "û�п��õ���ҩ��(��)��,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '��ҩ�÷�ȱʡ����һ����ͬ
            If RowIn�䷽��(lngPreRow) Then
                If Val(.TextMatrix(lngPreRow, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(lngPreRow, COL_������ĿID))
            End If
        End If
        
        '��Ѫҽ������Ѫ;���ж�
        If rsInput!���ID = "K" Then
            If gblnѪ��ϵͳ Then
                vMsg = frmMsgBox.ShowMsgBox("��ѡ����Ѫҽ�����͡�", Me, , 2)
                If vMsg = vbNo Then
                    bln��Ѫ = True
                ElseIf vMsg = vbCancel Then
                    Exit Function
                End If
            Else
                bln��Ѫ = True
            End If
            strWhere = ""
            If bln��Ѫ = False And gblnѪ��ϵͳ = True Then
                strWhere = " And NVL(ִ�з���,0)=1 "
            End If
            '����������ȡһ��ȱʡ����Ѫ;��
            lng�÷�ID = Getȱʡ�÷�ID(IIF(bln��Ѫ And gblnѪ��ϵͳ, 9, 8), 2, strWhere)
            If lng�÷�ID = 0 Then
                .Refresh
                MsgBox "û�п��õ���Ѫ" & IIF(bln��Ѫ And gblnѪ��ϵͳ, "�ɼ�����", ";��") & ",���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            'ȱʡ����һ����ͬ
            If lngPreRow <> -1 Then
                If .TextMatrix(lngPreRow, COL_���) = "K" And Val(.TextMatrix(lngPreRow, COL_��鷽��)) = IIF(bln��Ѫ, "0", "1") Then
                    i = .FindRow(CStr(.RowData(lngPreRow)), lngPreRow + 1, COL_���ID)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID))
                    End If
                End If
            End If
        End If
        
        '������ҩ����ҩ;���ж�
        If InStr(",5,6,", rsInput!���ID) > 0 Then
'            '����������ȡһ��ȱʡ�ĸ�ҩ;��,ͬʱ�ж��Ƿ��и�ҩ;������
'            lng�÷�ID = Getȱʡ�÷�ID(2, 1)
'            If lng�÷�ID = 0 Then
'                .Refresh
'                MsgBox "û�п��õĸ�ҩ;��,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
'                .Refresh: Exit Function
'            End If
            '��ҩ;��ȱʡ����һ������ͬ���͵���ͬ
            If lngPreRow <> -1 And Not IsNull(rsInput!ҩƷ����) Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 And .TextMatrix(lngPreRow, COL_ҩƷ����) = Nvl(rsInput!ҩƷ����) Then
                    i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_���ID)), lngPreRow + 1)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID))
                    End If
                End If
            End If
        End If
        
        '������ҩ������������
        If InStr(",5,6,", rsInput!���ID) > 0 And gint�����Ǽ���Ч���� <> 0 Then
            str���� = Check��������(Me, txtҽ������, mlng����ID, rsInput!������ĿID, rsInput!����, mbln�Զ�Ƥ��, lngƤ��ID)
            If str���� <> "" Then
                .Refresh
                If MsgBox(str����, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    .Refresh: Exit Function
                End If
            End If
            
            '�Զ����Ƥ��
            If lngƤ��ID <> 0 Then
                '�ȼ���Ƿ����и�Ƥ��(����Чʱ�����ֹ����Զ���ӵ�,��������Ƥ��)
                i = .FixedRows - 1
                Do
                    i = .FindRow(CStr(lngƤ��ID), i + 1, COL_������ĿID)
                    If i <> -1 Then
                        If Not .RowHidden(i) Then
                            If Int(CDate(.Cell(flexcpData, i, COL_��ʼʱ��))) >= Int(zlDatabase.Currentdate - gint�����Ǽ���Ч����) Then
                                blnƤ���� = True: Exit Do '��¼������־,��ǰ��������ɺ�������
                            End If
                        End If
                    End If
                Loop Until i = -1
            End If
        End If
        
        '������ҩ��һ����ҩ���ж�,����ȱʡ�ǰ���һ���ģ������һ����һ����
        blnGroup = RowInһ����ҩ(lngRow) And Not blnByƤ��
        If blnGroup Then
            If rsInput!���ID = "9" Then
                .Refresh
                MsgBox "������һ����ҩ��ҩƷ��ֱ��������׷�����", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
        
            If .RowData(lngRow) = 0 Then
                'һ����ҩ�еĴ�������У�ֻ�в�����һ����ҩ���м�,�����Զ���Ϊһ����ҩ
                lngGroupRow = lngPreRow
            Else
                'һ����ҩ�е�ҩƷ�У������ǵ�һ�л����һ��'ȡ��ǰ�е���һ�У������ڲ�������ҽ��ʱ��ѡ������Ŀ����ʱ����ǰ�е����ݱ�ɾ�������������޷�ȡ�����е�ֵ
                If lngPreRow = -1 Then
                    lngGroupRow = vsAdvice.FindRow(.TextMatrix(lngRow, COL_���ID), lngRow + 1, COL_���ID)
                Else
                    If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                        And Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                        lngGroupRow = lngPreRow
                    Else
                        lngGroupRow = lngNextRow
                    End If
                End If
            End If
            
            'һ����ҩ��,��𣬱�����ͬ
            If Decode(rsInput!���ID, "5", "Y", "6", "Y", "N") <> Decode(.TextMatrix(lngGroupRow, COL_���), "5", "Y", "6", "Y", "N") Then
                .Refresh
                MsgBox "����һ����ҩ��ҩƷ���붼Ϊ����ҩ���г�ҩ��", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            i = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_���ID)), lngGroupRow + 1)
            lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID)) 'һ����ҩ�ĸ�ҩ;����ͬ
            
            '���һ����ҩ�ĵĸ�ҩ;���Ƿ��ʺ��ڵ�ǰ����ҩƷ(��һ����ҩ��ȱʡ�÷������뺯���������жϴ���)
            If Not Check�����÷�(lng�÷�ID, rsInput!������ĿID, 1) Then
                .Refresh
                MsgBox "һ���ĸ�ҩ;��Ϊ""" & .TextMatrix(i, col_ҽ������) & """���������ڵ�ǰ����ҩƷ��", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
        End If
    
        '������Ŀ
        If rsInput!���ID = "9" Then
            If GetGroupCount(rsInput!������ĿID, 1) = 0 Then
                .Refresh
                MsgBox """" & rsInput!���� & """��һ�����׷�������û��������Ч�������Ŀ��" & vbCrLf & "���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            strExtData = frmSchemeSelect.ShowMe(Me, rsInput!������ĿID, 1, mlng���˿���id, mstr�Ա�)
            If strExtData = "" Then .Refresh: Exit Function
        End If
    
        '��Ҫ����������ݵ�һЩ��Ŀ
        '---------------------------------------------------------------------------------------------------------------
        intType = -1
        lngAppType = -1
        If rsInput!���ID <> "9" Then strExtData = ""
        If rsInput!���ID = "D" Then
            If gblnOut���� And CanUseApply("D", Val(rsInput!������ĿID & "")) Then
                lngAppType = 0
            Else
                '�����Ŀ����Ҫ��չ�༭�ˣ�������ǰ���е���λ��Ŀ
                intType = 0
            End If
        ElseIf rsInput!���ID = "F" Then
            '��������Ҫ����������Ŀ������ѡ�񸽼�����
            intType = 1
        ElseIf InStr(",7,8,", rsInput!���ID) > 0 Then
            '��ҩ�䷽(��ζ��ҩ���䷽����)
            intType = 2
        ElseIf rsInput!���ID = "C" Then
            If gblnOut���� And CanUseApply("C", Val(rsInput!������ĿID & ""), rsInput!���� & "") Then
                lngAppType = 3
            Else
                '����һ���ɼ��Ķ��������Ŀ������걾
                intType = 4
                strExtData = rsInput!������ĿID & ";" & Nvl(rsInput!���) '��Ŀ;�걾
            End If
        ElseIf rsInput!���ID = "K" Or rsInput!���ID = "E" Or rsInput!���ID = "Z" Then
            If gblnOut���� And CanUseApply(rsInput!���ID) Then
                lngAppType = 1
            ElseIf CheckApplication(Val(rsInput!������ĿID & ""), 1) Then
                '���ƺ���Ѫ�����������븽������д
                intType = 5
            End If
        End If
        '�жϵ�ǰ��Ŀ�Ƿ�����Զ������뵥��������򵯳�
        lngApplyID = GetApplyCustom(Val(rsInput!������ĿID & ""))
        If intType <> -1 And lngApplyID = 0 Then
            With t_Pati
                .blnҽ�� = InStr(",1,2,", mstr������) > 0 And mstr������ <> ""
                .int���� = mint����
                .intӤ�� = mintӤ��
                .lng����ID = mlng����ID
                .lng���˿���ID = mlng���˿���id
                .str�Һŵ� = mstr�Һŵ�
                .str�Ա� = mstr�Ա�
            End With
            If intType = 2 Then
                lngPreRow��ҩ�� = GetPreRow��ҩ��(lngRow)
                lngҩƷID = Val("" & rsInput!�շ�ϸĿID)   'һ���䷽ʱΪ��
            End If
            On Error Resume Next
            '����ӿڣ���ǰint���ϴ�δ�������ڴ�0��bytUseType��ǰδ�������ڴ�0
            If intType = 2 Then
                blnOK = frmAdviceFormula.ShowMe(Me, gclsInsure, txtҽ������.hWnd, t_Pati, 0, 0, 1, 1, 1, rsInput!������ĿID, strExtData, _
                             strժҪOut, lngҩƷID, lngPreRow��ҩ��)
            Else
                blnOK = frmAdviceEditEx.ShowMe(Me, txtҽ������.hWnd, t_Pati, 0, intType, 0, 1, 1, 1, mblnNewLIS, True, rsInput!������ĿID, strExtData, _
                            strAppend, GetAdviceAppendItem, GetAdviceDiagnosis, str������λ)
            End If
            
            On Error GoTo errH
            If intType = 2 Then strժҪ = strժҪOut
            If Not blnOK Then Exit Function
        End If
        
        If lngAppType <> -1 Or lngApplyID <> 0 Then
            On Error Resume Next
            If lngAppType = 0 Then
                blnOK = ApplyNew�������(0, rsInput!������ĿID & "", objAppPages())
            ElseIf lngAppType = 1 Then
                blnOK = ApplyNew��Ѫ����(0, rsInput!������ĿID & "", rsCard, bln��Ѫ)
            ElseIf lngAppType = 3 Then
                blnOK = ApplyNew��������(0, rsInput!���� & "", rsCard)
                        ElseIf lngApplyID <> 0 Then
                FuncApplyCustom 0, lngApplyID, , Val(rsInput!������ĿID & "")
            End If
            On Error GoTo errH
            If Not blnOK Then Exit Function
        End If
        
        '�޸�������Ŀʱ,��ɾ����ǰҽ��������
        '---------------------------------------------------------------------------------------------------------------
        If .RowData(lngRow) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '����ҩ���г�ҩ
                If Not blnGroup Then
                    '������ҩɾ����ҩ;����,�������ǰ��
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    Call DeleteRow(i)
                    Call DeleteRow(lngRow, True)
                Else
                    'һ���ҩʱ,ֻ�����ǰ��
                    lng���״̬ = Val(vsAdvice.TextMatrix(lngRow, COL_���״̬))
                    Call DeleteRow(lngRow, True)
                    If lng���״̬ <> 0 Then Call ReSet���״̬ͼ��(lngGroupRow)
                End If
            ElseIf InStr(",D,F,K,", .TextMatrix(lngRow, COL_���)) > 0 Then
                If .TextMatrix(lngRow, COL_���) = "D" And 0 <> Val(.TextMatrix(lngRow, COL_�������)) Then
                    Call Delete���������Ѫ(lngRow, True, lngTmpRow)
                    lngRow = lngTmpRow
                Else
                    '��������Ŀ��������Ŀ����Ѫҽ��
                    'ɾ����λ�У�������������(��������,������Ŀ)������Ѫ;��
                    Call Delete���������Ѫ(lngRow)
                    '�����ǰ��
                    Call DeleteRow(lngRow, True)
                End If
            ElseIf RowIn�䷽��(lngRow) Then
                '��ҩ�䷽��˳��(���)Ҫ������ϸ����
                'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
                lngRow = Delete��ҩ�䷽(lngRow)
                '�����ǰ��(��ҩ�÷���)
                Call DeleteRow(lngRow, True)
            ElseIf RowIn������(lngRow) Then
                If lngAppType = 3 Then
                    Call Delete�������뵥(lngRow, True, lngTmpRow)
                    lngRow = lngTmpRow
                Else
                    'ɾ��������Ŀ��:ɾ��֮�����¶�λ�ĵ�ǰ��
                    lngRow = Delete�������(lngRow)
                    '�����ǰ��(�ɼ�������)
                    Call DeleteRow(lngRow, True)
                End If
            Else
                '������Ŀֱ�������ǰ������
                Call DeleteRow(lngRow, True)
            End If
        End If
        
        '��ǰ������ҽ��
        '---------------------------------------------------------------------------------------------------------------
        If InStr(",7,8,", rsInput!���ID) > 0 Then
            '��ҩ�䷽(��ζ��ҩ���䷽����):����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet��ҩ�䷽(rsInput!������ĿID, lngRow, lng�÷�ID, strExtData, , strժҪ)
    
            '�����������ϵı�Ǵ���
            Call SetDiagFlag(vsAdvice.Row, 1)
        ElseIf rsInput!���ID = "9" Then
            '����ҽ����Ҫ�ֽ�Ϊ�����Ŀ����
            Call AdviceSet������Ŀ(rsInput!������ĿID, lngRow, strExtData)
            
            '�����������ϵı�Ǵ���
            '���׺͸���ʱ���ӹ���������������
        ElseIf rsInput!���ID = "C" Then
            If lngAppType = 3 Then
                Call AdviceSet��������(lngRow, rsCard)
            Else
                '�������
                lngRow = AdviceSet�������(lngRow, lng�÷�ID, strExtData, , strժҪ)
            
                '�����������ϵı�Ǵ���
                Call SetDiagFlag(vsAdvice.Row, 1)
            End If
        ElseIf lngAppType = 0 Then
            '���������뵥
            lngRow = AdviceSet�������(lngRow, objAppPages())
            Call SetDiagFlag(vsAdvice.Row, 1)
        ElseIf lngAppType = 1 Then
            lngRow = AdviceSet��Ѫ����(lngRow, rsCard)
            Call SetDiagFlag(vsAdvice.Row, 1)
        Else
            '��Ѫҽ����飬�����м�������רҵ����ְ���ҽʦ�������´�
            If rsInput!���ID & "" = "K" And gbln��Ѫ�����м����� Then
                If UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "������ҽʦ" Then
                    MsgBox "��������Ѫ�ּ��������Ѫҽ��ֻ���м�������רҵ����ְ��ҽʦ�����´", vbInformation, Me.Caption
                    Exit Function
                End If
            End If
            '�С�����ҩ�����ģ����(���)������(���)����Ѫ��������������Ŀ
            Call AdviceSet������Ŀ(rsInput, lngRow, lng�÷�ID, lngGroupRow, strExtData, strժҪ, str������λ, bln��Ѫ)
            
            '�����������ϵı�Ǵ���
            Call SetDiagFlag(vsAdvice.Row, 1)
            
            '�Զ�����һ����ҩ
            If InStr(",5,6,", rsInput!���ID) > 0 Then
                i = CheckAutoMerge(lngRow)
                If i = 1 Then
                    mblnRowMerge = True
                ElseIf i = 2 Then
                    mblnRowMerge = False
                    Set objControl = cbsMain.FindControl(, conMenu_Merge, , True)
                    objControl.Checked = False
                End If
                If Not RowInһ����ҩ(lngRow) Then
                    If mblnRowMerge Then
                        '�ֹ�ʹһ����ҩ
                        lngRowID = .RowData(lngRow)
                        Call MergeRow(lngPreRow, lngRow) '����������ʾ��ǰ�е�����,������ǿ��RowChange
                        lngRow = .FindRow(lngRowID, lngPreRow + 1)
                        '����ҩ������һ��ʱһ����ҩ�ģ�Ӧ��ȱʡ����һ�еġ���ҩĿ�ġ��͡���ҩ���ɡ���ͬ��
                        If Val(.TextMatrix(lngRow, COL_�����ȼ�)) <> 0 Then
                            If lngRow > 1 Then
                                txt��ҩ����.Text = .TextMatrix(lngRow - 1, COL_��ҩ����)
                                If Val(.TextMatrix(lngRow - 1, COL_��ҩĿ��)) <> 0 Then
                                    cboDruPur.ListIndex = Val(.TextMatrix(lngRow - 1, COL_��ҩĿ��))
                                End If
                            End If
                        End If
                        '��¼���ҩƷһ����ҩʱ�����û����������¼����������÷�����ָ���˵������Զ��������������ݵ�һ��ҽ��������������
                        If Val(.TextMatrix(lngRow, COL_����)) > 0 And mbln���� = False Then
                            Call GetRowScope(lngRow, lngBegin, lngEnd)
                            If Val(.TextMatrix(lngBegin, COL_����)) > 0 And Val(.TextMatrix(lngBegin, COL_����)) > 0 And .TextMatrix(lngBegin, COL_Ƶ��) <> "" _
                                And Val(.TextMatrix(lngBegin, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(lngBegin, COL_Ƶ�ʼ��)) <> 0 _
                                And Val(.TextMatrix(lngBegin, COL_����ϵ��)) <> 0 And Val(.TextMatrix(lngBegin, COL_�����װ)) <> 0 Then
                                sng���� = CalcȱʡҩƷ����(Val(.TextMatrix(lngBegin, COL_����)), Val(.TextMatrix(lngBegin, COL_����)), _
                                                Val(.TextMatrix(lngBegin, COL_Ƶ�ʴ���)), Val(.TextMatrix(lngBegin, COL_Ƶ�ʼ��)), .TextMatrix(lngBegin, COL_�����λ), _
                                                Val(.TextMatrix(lngBegin, COL_����ϵ��)), Val(.TextMatrix(lngBegin, COL_�����װ)), _
                                                Val(.TextMatrix(lngBegin, COL_�ɷ����)))
                                .TextMatrix(lngRow, COL_����) = ReGetҩƷ����(Val(.TextMatrix(lngRow, COL_����)), Val(.TextMatrix(lngRow, COL_����)), sng����, lngRow)
                            End If
                        End If
                    ElseIf lngPreRow <> -1 Then
                        '�Զ�ʹһ����ҩ
                        Set objControl = cbsMain.FindControl(, conMenu_Merge, , True)
                        If objControl.Checked = True Then
                            If .TextMatrix(lngPreRow, COL_���) = rsInput!���ID Then
                                If RowInһ����ҩ(lngPreRow) And RowCanMerge(lngPreRow, lngRow) And GetNextRow(lngRow) = -1 Then
                                    mblnRowMerge = True
                                    cbsMain.RecalcLayout '��ʱˢ��
                                    lngRowID = .RowData(lngRow)
                                    Call MergeRow(lngPreRow, lngRow, False)
                                    lngRow = .FindRow(lngRowID, lngPreRow + 1)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        '���¸�������:�Ե�ǰ�ɼ���Ϊ׼
        If strAppend <> "" Then
            .TextMatrix(.Row, COL_����) = strAppend
            .Cell(flexcpData, .Row, COL_����) = 1 '������Ҫ����д��(�������޸�)
            Call ReplaceAdviceAppend(.Row) 'ȱʡ�滻����ҽ�������븽��
        End If
        
        '������ҩ�ɳ�ҩʱ�Զ�����Ƥ����:����֮���Զ�λ�ڵ�ǰҩƷ
        If lngƤ��ID <> 0 And Not blnƤ���� Then
            Call AdviceSet��������(.Row, lngƤ��ID) 'ע���õ�ǰ��,��Ϊһ��֮��λ�ı�
        End If
        
        '�����Զ������и�
        Call .AutoSize(col_ҽ������)
    End With
    
    mblnNoSave = True '���Ϊδ����
    
    '�Ա��ն�����м��
    Call GetInsureStr(strIDs1, strIDs2, strҽ������, vsAdvice.Row)
    strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, 1, strIDs1, strIDs2, strҽ������)
    If strMsg <> "" Then
        If gintҽ������ = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln���Ѷ��� = False
    End If
    
    '������ҽӿ�
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        If zlPluginAdviceEnter(vsAdvice.Row) = False Then
            vsAdvice.Refresh: Exit Function
        End If
    End If
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlPASSMap()
'����:����Pass VsAdvie����ӳ��
'ע��:ɾ�����޸�������������ʱ�����������ҩ�����еĹ�������
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "������ҩ���", True)
        mblnPass = Not mobjPassMap Is Nothing And Not gobjPass Is Nothing
    End If
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_����༭
            .int���� = mint����
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .objCmdBar = cmdAlley
            
            Set .diags = .GetDiags
            
            Set .VSCOL = .GetVSCOL( _
                , COL_���ID, COL_���, COL_������ĿID, COL_�շ�ϸĿID, col_ҽ������, , COL_����, COL_������λ, COL_�÷�, _
                COL_����, COL_Ӥ��, COL_����ʱ��, COL_����ҽ��, COL_��ʼʱ��, COL_��������ID, , COL_Ƶ��, COL_Ƶ�ʴ���, COL_Ƶ�ʼ��, _
                COL_�����λ, COL_��ʾ, COL_���, COL_״̬, COL_EDIT, , , , COL_ִ������, COL_�걾��λ, _
                , , , , , COL_����, COL_������λ, COL_ҽ������, COL_��ҩĿ��, COL_��������, , _
                COL_��ҩ����, COL_��־, COL_�������, COL_ִ�з���)
            Set .PassPati = .GetPatient()
            Call zlPASSPati
        End With
    End If
End Sub

Private Sub zlPASSPati()
'����:���ò�����Ϣ
    If Not mobjPassMap Is Nothing Then
        With mobjPassMap.PassPati
            .intӤ�� = IIF(cboӤ��.ListIndex = -1, 0, cboӤ��.ListIndex)  'ȱʡ��������Ϊ0
            .dbl��ʶ�� = mdbl�����
            .Dat�������� = mDat��������
            .lng����ID = mlng����ID
            .lng�Һ�ID = mlng�Һ�ID
            .str�Һŵ� = mstr�Һŵ�
            .str�Ա� = mstr�Ա�
            .str���� = mstr����
        End With
    End If
End Sub

Private Sub zlPassDrags()
'����:���ò��������Ϣ
    Dim i As Long
    
    If Not mobjPassMap Is Nothing Then
        Set mobjPassMap.diags = Nothing '������¸�ֵ
        Set mobjPassMap.diags = mobjPassMap.GetDiags
        With vsDiag
            For i = .FixedRows To .Rows - 1
                mobjPassMap.diags.Add .TextMatrix(i, col���), .TextMatrix(i, col��ϱ���), .TextMatrix(i, col��������), "_" & i
            Next
        End With
    End If
End Sub

Private Sub MergeRow(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional ByVal blnCheck As Boolean = True)
'���ܣ�����������Ϊһ����ҩ
'������lngRow1=ǰ����,���ܱ����Ѿ�����һ����ҩ
'      lngRow2=��ǰ��
'˵����������ɺ�,����Զ�λ��ԭlngRow2�ĵ�ǰ��
    Dim lngBegin As Long, lngEnd As Long
    Dim blnDo As Boolean, lngTmp As Long
    Dim lngDiag As Long
    
    With vsAdvice
        If blnCheck Then
            blnDo = RowCanMerge(lngRow1, lngRow2)
        Else
            blnDo = True
        End If
        If blnDo Then
            mblnRowChange = False: .Redraw = flexRDNone
            
            '��¼��ǰ�еĹ������,һ�����Դ�Ϊ׼
            lngDiag = AdviceHaveDiag(lngRow2)
            
            lngTmp = .RowData(lngRow2) '��¼���ٶ�λ����ǰ��
            '��ȡ��֮ǰ��һ����ҩ
            If RowInһ����ҩ(lngRow1) Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow1, COL_���ID)), lngBegin, lngEnd)
                Call AdviceSet������ҩ(lngBegin, lngEnd)
                lngRow1 = lngBegin
                lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            End If
            Call AdviceSetһ����ҩ(lngRow1, lngRow2)
            lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            .Row = lngRow2
            
            '��һ��֮ǰ�ĵ�ǰ��Ϊ׼�ָ��������
            'һ��������ǰ���ҩƷ��Ϲ����ѱ�DeleteRow�Ƴ�
            If lngDiag <> -1 Then
                Call SetDiagFlag(.Row, 1, lngDiag)
            End If
            
            mblnRowChange = True: .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub SplitRow(ByVal lngRow As Long)
'���ܣ���ָ���д�һ����ҩ�ж�������(����һ����ҩ�������ٰ�������)
'������lngRow=��ǰ��,��Ϊһ����ҩ�е����һҩƷ��
'˵����������ɺ�,����Զ�λ��ԭlngRow�ĵ�ǰ��
    Dim lngBegin As Long, lngEnd As Long, lngTmp As Long
    
    With vsAdvice
        mblnRowChange = False: .Redraw = flexRDNone
        lngTmp = .RowData(lngRow) '��¼���ڻָ���λ��ǰ��
        Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow, COL_���ID)), lngBegin, lngEnd)
        
        '��ȡ��������һ����ҩ
        Call AdviceSet������ҩ(lngBegin, lngEnd)
        
        '�����ó�����������Ϊһ����ҩ
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        lngEnd = GetPreRow(lngRow)
        Call AdviceSetһ����ҩ(lngBegin, lngEnd)
        
        '�ָ���ǰ��
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        .Row = lngRow
        mblnRowChange = True: .Redraw = flexRDDirect
    End With
End Sub

Private Sub AdviceSet����ҽ��(ByVal lng����ID As Long, ByVal str�Һŵ� As String, ByVal strIDs As String, Optional ByVal blnHistory As Boolean)
'���ܣ�����ָ�����˵�ָ��ҽ��������Ϊ��ҽ��
'˵�����ɹ��ⲿ����,����֮ǰ��������ҽ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln�䷽ As Boolean
    Dim lngBegin As Long, lngEnd As Long
    Dim curDate As Date, blnHide As Boolean
    Dim lng��������ID As Long, dbl���ID As Double
    Dim lng��� As Long, intCount As Integer
    Dim lngRow As Long
    Dim i As Long, j As Long
    
    Dim lng��ҩ��ID As Long, lng��ҩ��ID As Long, lng��ҩ��ID As Long, lng���ϲ���ID As Long
    Dim strҩ��IDs As String, str���� As String
    Dim blnҩƷС������ As Boolean
    Dim str��ΣҩƷ As String
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    strSQL = _
        " Select a.ID+0.1 As ID,Decode(a.���id,Null,Null,a.���id+0.1) As ���id,Nvl(A.Ӥ��,0) as Ӥ��,A.���,A.ҽ����Ч," & _
        " A.ҽ��״̬,A.�������,A.������ĿID,B.����,A.�걾��λ,A.��鷽��,A.ִ�б��,A.�շ�ϸĿID," & _
        " A.��ʼִ��ʱ��,Nvl(B.����,A.ҽ������) ҽ������,A.ҽ������,A.��������,A.����,A.�ܸ�����,B.���㵥λ," & _
        " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,B.���㷽ʽ,B.ִ��Ƶ��,B.��������,B.����Ӧ��,B.ִ�з���," & _
        " B.�Ƽ�����,A.ִ��ʱ�䷽��,Decode(nvl(Instr(',5,6,7,',a.�������),0),0,b.ִ�п���,a.ִ������) as ִ������,A.ִ�п���ID,A.��������ID,A.����ҽ��,A.����ʱ��," & _
        " A.������־,C.�������,C.������,C.ҩƷ����,B.¼������,C.��������,C.����ְ��," & _
        " D.����ϵ��,D.�����װ,D.���ﵥλ,F.���㵥λ as ɢװ��λ,E.��������,D.����ɷ���� As �ɷ����,a.�䷽ID,c.�ٴ��Թ�ҩ,d.��ΣҩƷ,a.�����ĿID,c.��ý,d.����ҩ��" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,ҩƷ���� C,ҩƷ��� D,�������� E,�շ���ĿĿ¼ F" & _
        " Where Nvl(A.ҽ����Ч,0)=1 And A.������ĿID=B.ID" & _
        " And A.������ĿID=C.ҩ��ID(+) And A.�շ�ϸĿID=D.ҩƷID(+)" & _
        " And A.�շ�ϸĿID=E.����ID(+) And E.����ID=F.ID(+)" & _
        " And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
        " And Instr([3],','||A.ID||',')>0" & _
        " Order by Ӥ��,���"
    If blnHistory Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str�Һŵ�, "," & strIDs & ",")
    On Error GoTo 0
    
    If Not rsTmp.EOF Then
        If rsTmp!������� = "Z" And (rsTmp!�������� = "1" Or rsTmp!�������� = "2") Then
            If CheckInHosAdvice Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        
    
        lngBegin = vsAdvice.Row '��ʼ������
        lng��� = GetCurRow���(lngBegin) '��ʼ���
        intCount = 0 '�Ѿ����õ�����
        curDate = zlDatabase.Currentdate
        lng��������ID = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
        'ҩƷ�����Ƿ��������С������ҩ����������С�������ֻ�����ҩ���г�ҩ
        blnҩƷС������ = InStr(GetInsidePrivs(pסԺҽ���´�), "ҩƷС������") > 0
            
        mblnRowChange = False
        With vsAdvice
            .Redraw = flexRDNone
            For i = lngBegin To rsTmp.RecordCount + lngBegin - 1
                If i > lngBegin Then .AddItem "", i

                bln�䷽ = False
                
                .RowData(i) = -1 * rsTmp!ID
                If Not IsNull(rsTmp!���ID) Then
                    .TextMatrix(i, COL_���ID) = -1 * rsTmp!���ID
                End If
                .TextMatrix(i, COL_���) = lng��� + intCount
                
                .TextMatrix(i, COL_EDIT) = 1 '����
                .Cell(flexcpData, i, COL_EDIT) = CStr(lng����ID & "," & str�Һŵ�) '��¼��صĸ�����Ŀ
                .TextMatrix(i, COL_״̬) = 1 '�¿�
                .TextMatrix(i, COL_Ӥ��) = cboӤ��.ListIndex
                .TextMatrix(i, COL_���) = rsTmp!�������
                .TextMatrix(i, COL_������ĿID) = rsTmp!������ĿID
                .TextMatrix(i, COL_����) = rsTmp!����
                .TextMatrix(i, COL_�걾��λ) = Nvl(rsTmp!�걾��λ)
                .TextMatrix(i, COL_��鷽��) = Nvl(rsTmp!��鷽��)
                .TextMatrix(i, COL_ִ�б��) = Nvl(rsTmp!ִ�б��, 0)
                .TextMatrix(i, COL_�շ�ϸĿID) = Nvl(rsTmp!�շ�ϸĿID)
                .TextMatrix(i, col_ҽ������) = Nvl(rsTmp!ҽ������)
                .TextMatrix(i, COL_ҽ������) = Nvl(rsTmp!ҽ������)
                .Cell(flexcpData, i, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, Val(.TextMatrix(i, COL_�շ�ϸĿID)), "", 0, "", .TextMatrix(i, COL_������ĿID))
                
                .TextMatrix(i, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
                .TextMatrix(i, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
                .TextMatrix(i, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
                .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, COL_����Ӧ��) = Nvl(rsTmp!����Ӧ��)
                .TextMatrix(i, COL_ִ�з���) = Nvl(rsTmp!ִ�з���, 0)
                .TextMatrix(i, COL_�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, COL_�����ȼ�) = Val("" & rsTmp!������)
                .TextMatrix(i, COL_�䷽ID) = Nvl(rsTmp!�䷽ID)
                .TextMatrix(i, COL_�ٴ��Թ�ҩ) = rsTmp!�ٴ��Թ�ҩ & ""
                .TextMatrix(i, COL_�����ĿID) = rsTmp!�����ĿID & ""
                .TextMatrix(i, COL_��ΣҩƷ) = Val(rsTmp!��ΣҩƷ & "")
                .TextMatrix(i, COL_�Ƿ���ý) = Val(rsTmp!��ý & "")
                .TextMatrix(i, COL_����ҩ��) = rsTmp!����ҩ�� & ""
                If Val(.TextMatrix(i, COL_��ΣҩƷ)) <> 0 Then
                    str��ΣҩƷ = str��ΣҩƷ & vbCrLf & .TextMatrix(i, col_ҽ������) & ":" & Decode(Val(.TextMatrix(i, COL_��ΣҩƷ)), 1, "A", 2, "B", 3, "C", "") & "����"
                End If
                
                .TextMatrix(i, COL_ҩƷ����) = Nvl(rsTmp!ҩƷ����)
                If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                Else
                    .TextMatrix(i, COL_��������) = Nvl(rsTmp!¼������)
                End If
                .TextMatrix(i, COL_����ְ��) = Nvl(rsTmp!����ְ��)
                
                If InStr(",5,6,7,", .TextMatrix(i, COL_���)) > 0 Then
                    .TextMatrix(i, COL_����ϵ��) = Nvl(rsTmp!����ϵ��)
                    .TextMatrix(i, COL_�����װ) = Nvl(rsTmp!�����װ)
                    .TextMatrix(i, COL_���ﵥλ) = Nvl(rsTmp!���ﵥλ)
                    If Not IsNull(rsTmp!����ϵ��) Then
                        .TextMatrix(i, COL_�ɷ����) = Nvl(rsTmp!�ɷ����, 0)
                    End If
                ElseIf .TextMatrix(i, COL_���) = "4" Then
                    .TextMatrix(i, COL_����ϵ��) = 1
                    .TextMatrix(i, COL_�����װ) = 1
                    .TextMatrix(i, COL_���ﵥλ) = Nvl(rsTmp!ɢװ��λ)
                    .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                End If
                
                If IsDate(txt��ʼʱ��.Text) Then
                    .TextMatrix(i, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_��ʼʱ��) = txt��ʼʱ��.Text
                    
                    '����/��Ѫʱ�䣺����ʱȱʡ�뿪ʼʱ����ͬ,�ڱ걾��λ�����
                    If rsTmp!������� = "K" Or rsTmp!������� = "F" Or rsTmp!������� = "G" _
                        And Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                        .TextMatrix(i, COL_����ʱ��) = txt��ʼʱ��.Text
                    End If
                End If
                
                .TextMatrix(i, COL_Ƶ��) = Nvl(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, COL_Ƶ�ʴ���) = Nvl(rsTmp!Ƶ�ʴ���)
                .TextMatrix(i, COL_Ƶ�ʼ��) = Nvl(rsTmp!Ƶ�ʼ��)
                .TextMatrix(i, COL_�����λ) = Nvl(rsTmp!�����λ)
                .TextMatrix(i, COL_ִ��ʱ��) = Nvl(rsTmp!ִ��ʱ�䷽��)
                
                .TextMatrix(i, COL_ִ������) = Nvl(rsTmp!ִ������, 0)
                
                '����ִ�п���
                If rsTmp!������� = "Z" Then
                    .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID)
                ElseIf InStr(",0,5,", Nvl(rsTmp!ִ������, 0)) = 0 Then
                    If Nvl(rsTmp!ִ�п���ID, 0) <> 0 Then
                        If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                            strҩ��IDs = Get����ҩ��IDs(rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), mlng���˿���id, 1)
                            If InStr("," & strҩ��IDs & ",", "," & rsTmp!ִ�п���ID & ",") > 0 Then
                                .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID, 0)
                            End If
                        ElseIf Nvl(rsTmp!�������) = "4" Then
                            strҩ��IDs = Get���÷��ϲ���IDs(Nvl(rsTmp!�շ�ϸĿID, 0), mlng���˿���id, 1)
                            If InStr("," & strҩ��IDs & ",", "," & rsTmp!ִ�п���ID & ",") > 0 Then
                                .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID, 0)
                            End If
                        ElseIf Val(.TextMatrix(i, COL_ִ������)) = 4 Then
                            '4-ָ������ʱ��ȡ,�����Ĺ̶�����
                            .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID, 0)
                            
                            '���ִ�п��ҵ���Ч��
                            If Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 Then
                                If CheckExecDeptValidate(Val(.TextMatrix(i, COL_ִ�п���ID)), mlng���˿���id, 1, Val(.TextMatrix(i, COL_������ĿID))) = False Then
                                    .TextMatrix(i, COL_ִ�п���ID) = 0
                                End If
                            End If
                        End If
                    End If
                    If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                        'ҩƷ�������������������ͬ
                        If rsTmp!������� = "5" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf rsTmp!������� = "6" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf rsTmp!������� = "7" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf Nvl(rsTmp!�������) = "4" Then
                            If lng���ϲ���ID = 0 Then
                                lng���ϲ���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!�������, Nvl(rsTmp!�շ�ϸĿID, 0), 4, mlng���˿���id, lng��������ID, 1, , 1)
                            End If
                            .TextMatrix(i, COL_ִ�п���ID) = lng���ϲ���ID
                        Else
                            .TextMatrix(i, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, 0, Nvl(rsTmp!ִ������, 0), mlng���˿���id, lng��������ID, 1, 1)
                        End If
                    End If
                End If
                
                If rsTmp!������� = "E" Then
                    If Nvl(rsTmp!���ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_���ID)) = -1 * rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼�ǳ�ҩ�ĸ�ҩ;��,������һ����ҩ��
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = -1 * rsTmp!ID Then
                                    '��ʾ��ҩ;��
                                    .TextMatrix(j, COL_�÷�) = rsTmp!���� & rsTmp!ҽ������
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼����ҩ�䷽���÷�,���䷽��ʾ��
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                            bln�䷽ = True
                        ElseIf .TextMatrix(i - 1, COL_���) = "C" Then
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                        End If
                    ElseIf Not IsNull(rsTmp!���ID) And .TextMatrix(i - 1, COL_���) = "K" And -1 * Nvl(rsTmp!���ID, 0) = .RowData(i - 1) Then
                        '��ǰ��¼����Ѫ;����
                        .TextMatrix(i - 1, COL_�÷�) = rsTmp!����
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        '��ǰ��¼����ҩ�䷽�巨��
                        bln�䷽ = True
                    End If
                ElseIf rsTmp!������� = "7" Then
                    bln�䷽ = True
                End If
                
                '����
                .TextMatrix(i, COL_����) = FormatEx(Nvl(rsTmp!��������), 5)
                If Nvl(rsTmp!�������) = "4" Then
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!ɢװ��λ)
                ElseIf InStr(",5,6,7,", rsTmp!�������) > 0 _
                    Or (Val(.TextMatrix(i, COL_Ƶ������)) = 0 And InStr(",1,2,", Nvl(rsTmp!���㷽ʽ, 0)) > 0) Then
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                End If
                
                '����
                If mbln���� Then
                    .TextMatrix(i, COL_����) = Nvl(rsTmp!����)
                End If

                '����
                If InStr(",5,6,", rsTmp!�������) > 0 Then
                    '��ҩ����������,�����۵�λ���,���ﵥλ��ʾ
                    If Not IsNull(rsTmp!�ܸ�����) And Not IsNull(rsTmp!�����װ) Then
                        .TextMatrix(i, COL_����) = FormatEx(rsTmp!�ܸ����� / rsTmp!�����װ, 5)
                    End If
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���ﵥλ)
                    
                    If Val(.TextMatrix(i, COL_�ɷ����)) = 0 And Not blnҩƷС������ And InStr(.TextMatrix(i, COL_����), ".") > 0 Then
                        .TextMatrix(i, COL_����) = IntEx(Val(.TextMatrix(i, COL_����)))
                    End If
                    
                    '����˵��������
                    Call Set��ҩ�����Ƿ���(i)
                ElseIf bln�䷽ Then
                    If Not IsNull(rsTmp!�ܸ�����) Then .TextMatrix(i, COL_����) = rsTmp!�ܸ�����
                    
                    .TextMatrix(i, COL_������λ) = "��" '��ҩ�䷽������λΪ"��"
                    
                    If rsTmp!������� = "E" And rsTmp!�������� = "4" Then   '��ҩ�÷�
                        Call Set��ҩ�����Ƿ���(i)
                    End If
                Else
                    '��������
                    If Not IsNull(rsTmp!�ܸ�����) Then .TextMatrix(i, COL_����) = rsTmp!�ܸ�����
                        
                    If Nvl(rsTmp!�������) = "4" Then
                        .TextMatrix(i, COL_������λ) = Nvl(rsTmp!ɢװ��λ)
                    Else
                        .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                    End If
                End If
                
                '����ҩ��ȱʡ��ҩĿ��
                If Val(.TextMatrix(i, COL_�����ȼ�)) > 0 Then .TextMatrix(i, COL_��ҩĿ��) = mstrPurMed
                
                .TextMatrix(i, COL_��־) = chk����.value '�����ڽ�����ͳһ����Ϊ����
                If gblnKSSStrict And UserInfo.��ҩ���� < Val(.TextMatrix(i, COL_�����ȼ�)) And .TextMatrix(i, COL_��־) <> "1" Then

                    .TextMatrix(i, COL_���״̬) = 1
                End If
                '������Ѫҽ�����״̬
                If .TextMatrix(i, COL_���) = "E" And .TextMatrix(i, COL_��������) = "8" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                    strSQL = ""
                    strSQL = GetBloodState(IIF(.TextMatrix(i, COL_��־) = "1", 1, 0), Val(.TextMatrix(i, COL_ִ�з���)))
                    .TextMatrix(i - 1, COL_���״̬) = strSQL
                    .TextMatrix(i, COL_���״̬) = strSQL
                End If
                
                .TextMatrix(i, COL_����ҽ��) = UserInfo.����
                .TextMatrix(i, COL_��������ID) = lng��������ID
                .TextMatrix(i, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
                
                Call SetRow��־ͼ��(i, 1)
                
                
                '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                If InStr(",5,6,", rsTmp!�������) > 0 And Not IsNull(rsTmp!�������) Then
                    If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", rsTmp!�������) > 0 Then
                        .Cell(flexcpFontBold, i, col_ҽ������) = True
                    End If
                End If
                
                lngEnd = i
                intCount = intCount + 1
                
                rsTmp.MoveNext
            Next
            
            '������ҽ��ID
            For i = lngBegin To lngEnd
                dbl���ID = .RowData(i)
                .RowData(i) = GetNextҽ��ID
                For j = i - 1 To lngBegin Step -1
                    If Val(.TextMatrix(j, COL_���ID)) = dbl���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To lngEnd
                    If Val(.TextMatrix(j, COL_���ID)) = dbl���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            Next
            If gblnOut���� Then Call MakeAppNo(2, lngBegin, lngEnd)
            Call Setҽ������(lngBegin, lngEnd)
            '������Ӱ���е����
            Call AdviceSetҽ�����(lngEnd + 1, intCount)
            
            '��ʾ/������
            lngRow = 0
            For i = lngBegin To lngEnd
                blnHide = False
                If .TextMatrix(i, COL_���) = "E" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                        blnHide = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_���)) > 0 _
                    And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                    blnHide = True
                End If
                                
                .RowHidden(i) = blnHide
                If Not blnHide And lngRow = 0 Then lngRow = i
                
                '����ҽ�����ݵı仯
                If Not .RowHidden(i) Then
                    .TextMatrix(i, col_ҽ������) = AdviceTextMake(i)
                End If
                
                'Ԥ�ȼ������Ƶ���
                If Not .RowHidden(i) And .TextMatrix(i, COL_����) = "" Then
                    .TextMatrix(i, COL_����) = GetItemPrice(i)
                End If
                
                '��������ʱ���ڿɼ��ж�ȡ������Ŀ��������ڼ�飬�����ڱ���
                If Not .RowHidden(i) Then
                    If Not RowIn�䷽��(i) Then
                        If RowIn������(i) Then
                            j = .FindRow(CStr(.RowData(i)), , COL_���ID)
                            If j <> -1 Then
                                .TextMatrix(i, COL_����) = Getҽ����Ŀ����(Val(.TextMatrix(j, COL_������ĿID)), 1)
                            End If
                        Else
                            .TextMatrix(i, COL_����) = Getҽ����Ŀ����(Val(.TextMatrix(i, COL_������ĿID)), 1)
                        End If
                        If .TextMatrix(i, COL_����) <> "" Then
                            str���� = str���� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������)
                        End If
                    End If
                End If
                
                '�����������ϵı�Ǵ���
                If Not .RowHidden(i) Then
                    Call SetDiagFlag(i, 1)
                End If
            Next
            
            'ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            
            .Row = lngRow: .Col = col_ҽ������
            
            Call .AutoSize(col_ҽ������)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
        mblnNoSave = True '���Ϊδ����
    End If

    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call CalcAdviceMoney '��ʾ�¿�ҽ�����

    Screen.MousePointer = 0
    
    If str���� <> "" Then
        MsgBox "����ҽ����Ҫ��д���븽���ע����д��" & vbCrLf & str����, vbInformation, gstrSysName
    End If
    If str��ΣҩƷ <> "" Then
        MsgBox "����ҽ���Ǹ�ΣҩƷ��" & str��ΣҩƷ, vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AdviceSet������Ŀ(ByVal lng����ID As Long, ByVal lngRow As Long, Optional ByVal str��� As String)
'���ܣ����������Ŀ(����һ����ҩ,������,��������,��ҩ�䷽)
'������lngRow=�յ�������(�����ǲ��������,����λ��һ����ҩ�м�)
    Dim rsItems As New ADODB.Recordset
    Dim rs��� As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim rs�Ƴ� As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    Dim lngCurRow As Long, intCount As Integer, lng��� As Long
    Dim lngPreRow As Long, vCurDate As Date, lngTmp As Long
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim bln��ҩ;�� As Boolean, bln�ɼ����� As Boolean, bln��Ѫ;�� As Boolean
    Dim bln��ҩ�÷� As Boolean, bln��ҩ�巨 As Boolean, bln�䷽ As Boolean
    Dim lng��ҩ��ID As Long, lng��ҩ��ID As Long, lng��ҩ��ID As Long
    Dim dbl���ID As Double, int���÷�Χ As Integer, strƵ�� As String
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim lng���� As Long, vBookMark As Variant, strҩ��IDs As String
    Dim sng���� As Single, strSQL��� As String, str���� As String
    Dim intƵ������ As Integer, lng���ϲ���ID As Long, blnAdd As Boolean
    Dim str��ΣҩƷ As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Me.Refresh
    
    '������Ź��˴�
    If str��� <> "" Then
        If Left(str���, 1) = "+" Then
            strSQL��� = " And Instr([2],','||A.���||',')>0"
        ElseIf Left(str���, 1) = "-" Then
            strSQL��� = " And Instr([2],','||A.���||',')=0"
        End If
    End If
    
    'ҩƷ�����Ϣ:��Ȼ�����շ�ϸĿID,����ǰ������û��
    strSQL = "Select A.���,B.ҩ��ID,B.ҩƷID,B.����ϵ��,B.�����װ,B.���ﵥλ,b.����ҩ��," & _
        " B.����ɷ���� As �ɷ����,C.����,Nvl(D.����,C.����) as ����,C.���,C.����,b.��ΣҩƷ" & _
        " From ������Ŀ��� A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
        " Where A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID" & _
        " And C.ID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=[3]" & _
        " And A.�������ID=[1]" & strSQL��� & _
        " Order by A.���,C.����"
    Set rs��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",", IIF(gbytҩƷ������ʾ = 0, 1, 3))
    
    '������Ϣ
    strSQL = "Select A.���,B.����ID,B.��������,C.����,C.���㵥λ" & _
        " From ������Ŀ��� A,�������� B,�շ���ĿĿ¼ C" & _
        " Where A.�շ�ϸĿID=B.����ID And B.����ID=C.ID" & _
        " And A.�������ID=[1]" & strSQL��� & _
        " Order by A.���,C.����"
    Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",")
    
    '��ҩ�Ƴ���Ϣ(���������ֱ�Ӷ�Ӧ�䷽,��ҩȡ�����Ƴ�)
    strSQL = "Select Distinct A.������ĿID,C.�Ƴ�" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,�����÷����� C" & _
        " Where A.������ĿID=B.ID And B.��� IN('5','6')" & _
        " And A.������ĿID=C.��ĿID And A.�������ID=[1]" & strSQL���
    Set rs�Ƴ� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",")
    
    '��������к�Ӧ����ҽ���༭ʱ�Ĵ���һ��
    '�ſ�ִ��Ƶ��Ϊ�����Եĳ���(���ַ�������,ֻȡ����),���з���תΪ��������
    strSQL = "Select 1 as ��Ч,a.���+0.1 As ���,Decode(a.������,Null,Null,a.������+0.1) As ������,A.������ĿID,A.�շ�ϸĿID,A.ҽ������,A.����,A.�ܸ�����,A.��������," & _
        " A.ҽ������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ�п���ID,Nvl(B.���,'*') ���,B.����," & _
        " B.���㵥λ,Decode(B.���,'D',A.�걾��λ,Nvl(A.�걾��λ,B.�걾��λ)) as �걾��λ,A.��鷽��," & _
        " A.ʱ�䷽��,Nvl(A.ִ������,B.ִ�п���) as ִ������,B.�Ƽ�����,B.����Ӧ��,B.��������,B.ִ�з���,B.���㷽ʽ," & _
        " B.ִ��Ƶ��,B.¼������,C.��������,C.����ְ��,C.�������,C.������,C.ҩƷ����,A.�䷽ID,C.�ٴ��Թ�ҩ,A.�����ĿID,C.��ý" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,ҩƷ���� C,�շ���ĿĿ¼ D" & _
        " Where A.������ĿID=B.ID(+) And A.������ĿID=C.ҩ��ID(+)  And d.id(+)=a.�շ�ϸĿID " & _
        " And A.��Ч=1 And A.�������ID=[1]" & strSQL��� & _
        " And (NVL(d.����ʱ��,b.����ʱ��) is null or NVL(d.����ʱ��,b.����ʱ��) = To_Date('3000/1/1', 'yyyy/mm/dd') Or Not (b.��� = '7' Or b.��� = 'E' And b.ִ�з��� = 0 And b.�������� = '3'))" & _
        " Order by A.���"
    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",")
    With vsAdvice
        mblnRowChange = False
        .Redraw = flexRDNone
        
        lngPreRow = GetPreRow(lngRow) 'ǰһ������
        intCount = 0 '�Ѿ����õ�����
        lng��� = GetCurRow���(lngRow) '��ʼ���
        vCurDate = zlDatabase.Currentdate
        
        For i = 1 To rsItems.RecordCount
            blnAdd = True
            
            '����Ƿ������Ч��סԺҽ��������ҽ��
            If rsItems!��� = "Z" And InStr(",1,2,", "," & rsItems!�������� & ",") > 0 Then
                If CheckInHosAdvice Then
                    blnAdd = False
                End If
            End If
            
            If blnAdd Then
                lngCurRow = lngRow + intCount
                If lngCurRow > lngRow Then .AddItem "", lngCurRow
                 
                '��¼���ID
                .RowData(lngCurRow) = -1 * rsItems!���
                If Not IsNull(rsItems!������) Then
                    .TextMatrix(lngCurRow, COL_���ID) = -1 * rsItems!������
                End If
                
                .TextMatrix(lngCurRow, COL_EDIT) = 1 '������
                .Cell(flexcpData, lngCurRow, COL_EDIT) = lng����ID '��¼��صĳ�����Ŀ
                
                .TextMatrix(lngCurRow, COL_Ӥ��) = cboӤ��.ListIndex
                .TextMatrix(lngCurRow, COL_���) = lng��� + intCount
                .TextMatrix(lngCurRow, COL_״̬) = 1 '�¿�
                .TextMatrix(lngCurRow, COL_���) = rsItems!���
                .TextMatrix(lngCurRow, COL_������ĿID) = Nvl(rsItems!������ĿID, 0)
                .TextMatrix(lngCurRow, COL_����) = Nvl(rsItems!����)
                .TextMatrix(lngCurRow, COL_�걾��λ) = Nvl(rsItems!�걾��λ)
                .TextMatrix(lngCurRow, COL_��鷽��) = Nvl(rsItems!��鷽��)
                
                If IsDate(txt��ʼʱ��.Text) Then
                    .TextMatrix(lngCurRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, lngCurRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                    
                    '����/��Ѫʱ�䣺����ʱȱʡ�뿪ʼʱ����ͬ,�ڱ걾��λ�����
                    If rsItems!��� = "K" Or rsItems!��� = "F" Or rsItems!��� = "G" _
                        And Val(.TextMatrix(lngCurRow, COL_���ID)) = Val(.TextMatrix(lngCurRow - 1, COL_���ID)) Then
                        .TextMatrix(lngCurRow, COL_����ʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                    End If
                End If
                
                '����
                .TextMatrix(lngCurRow, COL_�Ƽ�����) = Nvl(rsItems!�Ƽ�����, 0)
                .TextMatrix(lngCurRow, COL_���㷽ʽ) = Nvl(rsItems!���㷽ʽ, 0)
                .TextMatrix(lngCurRow, COL_��������) = Nvl(rsItems!��������)
                .TextMatrix(lngCurRow, COL_����Ӧ��) = Nvl(rsItems!����Ӧ��)
                .TextMatrix(lngCurRow, COL_ִ�з���) = Nvl(rsItems!ִ�з���, 0)
                .TextMatrix(lngCurRow, COL_�������) = Nvl(rsItems!�������)
                .TextMatrix(lngCurRow, COL_�����ȼ�) = Val("" & rsItems!������)
                .TextMatrix(lngCurRow, COL_�䷽ID) = Nvl(rsItems!�䷽ID)
                .TextMatrix(lngCurRow, COL_�ٴ��Թ�ҩ) = rsItems!�ٴ��Թ�ҩ & ""
                .TextMatrix(lngCurRow, COL_�����ĿID) = rsItems!�����ĿID & ""
                .TextMatrix(lngCurRow, COL_�Ƿ���ý) = Val(rsItems!��ý & "")
                
                .TextMatrix(lngCurRow, COL_ҩƷ����) = Nvl(rsItems!ҩƷ����)
                If InStr(",5,6,7,", rsItems!���) > 0 Then
                    .TextMatrix(lngCurRow, COL_��������) = Nvl(rsItems!��������)
                Else
                    .TextMatrix(lngCurRow, COL_��������) = Nvl(rsItems!¼������)
                End If
                .TextMatrix(lngCurRow, COL_����ְ��) = Nvl(rsItems!����ְ��)
                
                'ҩƷ�����Ϣ:�в�ҩ�϶���,��ҩ�������������λ�Զ�ƥ��
                lng���� = 0: vBookMark = 0
                If rsItems!��� = "7" Or (InStr(",5,6,", rsItems!���) > 0) Then
                    If Not IsNull(rsItems!�շ�ϸĿID) Then '������ǰδ�����շ�ϸĿID
                        rs���.Filter = "ҩƷID=" & rsItems!�շ�ϸĿID
                    Else
                        rs���.Filter = "ҩ��ID=" & rsItems!������ĿID
                    End If
                    If Not rs���.EOF Then
                        If IsNull(rsItems!�շ�ϸĿID) Then
                            'ȡ����ϵ��Ϊ��������С����������һ�����
                            If CInt(Nvl(rsItems!��������, 0)) <> 0 Then
                                Do While Not rs���.EOF
                                    If rs���!����ϵ�� / rsItems!�������� = Int(rs���!����ϵ�� / rsItems!��������) Then
                                        If rs���!����ϵ�� / rsItems!�������� < lng���� Or lng���� = 0 Then
                                            vBookMark = rs���.Bookmark
                                            lng���� = rs���!����ϵ�� / rsItems!��������
                                        End If
                                    End If
                                    rs���.MoveNext
                                Loop
                                If vBookMark <> 0 Then rs���.Bookmark = vBookMark
                            End If
                            If rs���.EOF Then rs���.MoveFirst
                        End If
                        .TextMatrix(lngCurRow, COL_����) = Nvl(rs���!����)
                        .TextMatrix(lngCurRow, COL_�շ�ϸĿID) = rs���!ҩƷID
                        .TextMatrix(lngCurRow, COL_����ϵ��) = Nvl(rs���!����ϵ��)
                        .TextMatrix(lngCurRow, COL_�����װ) = Nvl(rs���!�����װ)
                        .TextMatrix(lngCurRow, COL_���ﵥλ) = Nvl(rs���!���ﵥλ)
                        .TextMatrix(lngCurRow, COL_�ɷ����) = Nvl(rs���!�ɷ����, 0)
                        .TextMatrix(lngCurRow, COL_��ΣҩƷ) = Nvl(rs���!��ΣҩƷ, 0)
                        .TextMatrix(lngCurRow, COL_����ҩ��) = rs���!����ҩ�� & ""
                        If Val(.TextMatrix(lngCurRow, COL_��ΣҩƷ)) <> 0 Then
                            str��ΣҩƷ = str��ΣҩƷ & vbCrLf & rsItems!���� & ":" & Decode(Val(.TextMatrix(lngCurRow, COL_��ΣҩƷ)), 1, "A", 2, "B", 3, "C", "") & "����"
                        End If
                    End If
                ElseIf rsItems!��� = "4" Then
                    rs����.Filter = "����ID=" & Nvl(rsItems!�շ�ϸĿID, 0)
                    If Not rs����.EOF Then
                        .TextMatrix(lngCurRow, COL_����) = Nvl(rs����!����)
                        .TextMatrix(lngCurRow, COL_���ﵥλ) = Nvl(rs����!���㵥λ) 'ɢװ��λ
                        .TextMatrix(lngCurRow, COL_��������) = Nvl(rs����!��������, 0)
                    End If
                    .TextMatrix(lngCurRow, COL_����ϵ��) = 1
                    .TextMatrix(lngCurRow, COL_�����װ) = 1
                    .TextMatrix(lngCurRow, COL_�շ�ϸĿID) = Nvl(rsItems!�շ�ϸĿID, 0)
                End If
                                    
                '�ж��Ƿ��ض���
                bln��ҩ;�� = False: bln�ɼ����� = False: bln��Ѫ;�� = False
                bln��ҩ�÷� = False: bln��ҩ�巨 = False: bln�䷽ = False
                If rsItems!��� = "E" Then
                    If IsNull(rsItems!������) Then
                        If Val(.TextMatrix(lngCurRow - 1, COL_���ID)) = .RowData(lngCurRow) Then
                            If InStr(",5,6,", .TextMatrix(lngCurRow - 1, COL_���)) > 0 Then
                                bln��ҩ;�� = True
                            ElseIf .TextMatrix(lngCurRow - 1, COL_���) = "C" Then
                                bln�ɼ����� = True
                            Else
                                bln��ҩ�÷� = True
                            End If
                        End If
                    ElseIf .TextMatrix(lngCurRow - 1, COL_���) = "K" And .RowData(lngCurRow - 1) = Val(.TextMatrix(lngCurRow, COL_���ID)) Then
                        bln��Ѫ;�� = True
                    Else
                        bln��ҩ�巨 = True
                    End If
                End If
                If rsItems!��� = "7" Or bln��ҩ�巨 Or bln��ҩ�÷� Then bln�䷽ = True
                        
                '��ȡ��ǰ��Ŀ�����÷�Χ
                If bln�ɼ����� Then
                    '�ɼ������Լ�����Ŀ��Ϊ׼
                    lngTmp = .FindRow(CStr(.RowData(lngCurRow)), , COL_���ID)
                    intƵ������ = .TextMatrix(lngTmp, COL_Ƶ������)
                Else
                    intƵ������ = Nvl(rsItems!ִ��Ƶ��, 0)
                End If
                If bln�䷽ Then
                    int���÷�Χ = 2 '��ҩ�䷽(�����巨,�÷�)����ҽ
    '            ElseIf bln�ɼ����� Then
    '                int���÷�Χ = -1 '�����������Ŀ��ͬ:һ����
                ElseIf intƵ������ = 0 Or bln��ҩ;�� _
                    Or InStr(",5,6,", .TextMatrix(lngCurRow, COL_���)) > 0 Then
                    int���÷�Χ = 1 '"��ѡƵ��"���ҩ(������ҩ;��)����ҽ
                ElseIf intƵ������ = 1 Then
                    int���÷�Χ = -1 'һ����
                ElseIf intƵ������ = 2 Then
                    int���÷�Χ = -2 '������
                End If
                        
                'Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ
                .TextMatrix(lngCurRow, COL_Ƶ������) = intƵ������
                If Not IsNull(rsItems!ִ��Ƶ��) Then
                    If CheckƵ�ʿ���(Nvl(rsItems!������ĿID, 0), int���÷�Χ, Nvl(rsItems!ִ��Ƶ��)) Then
                        If GetƵ����Ϣ_����(rsItems!ִ��Ƶ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ, CStr(int���÷�Χ)) Then
                            .TextMatrix(lngCurRow, COL_Ƶ��) = rsItems!ִ��Ƶ��
                            .TextMatrix(lngCurRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                            .TextMatrix(lngCurRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                            .TextMatrix(lngCurRow, COL_�����λ) = str�����λ
                        End If
                    End If
                End If
                If .TextMatrix(lngCurRow, COL_Ƶ��) = "" And Not IsNull(rsItems!������ĿID) Then 'ȡȱʡ��
                    Call GetȱʡƵ��(Nvl(rsItems!������ĿID, 0), int���÷�Χ, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                    .TextMatrix(lngCurRow, COL_Ƶ��) = strƵ��
                    .TextMatrix(lngCurRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                    .TextMatrix(lngCurRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                    .TextMatrix(lngCurRow, COL_�����λ) = str�����λ
                End If
                
                '����
                .TextMatrix(lngCurRow, COL_����) = FormatEx(Nvl(rsItems!��������), 5)
                If rsItems!��� = "4" Then
                    .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow, COL_���ﵥλ) 'ɢװ��λ
                ElseIf bln��ҩ�÷� Then
                    .TextMatrix(lngCurRow, COL_������λ) = ""
                Else
                    If InStr(",5,6,7,", rsItems!���) > 0 Or (intƵ������ = 0 And InStr(",1,2,", Nvl(rsItems!���㷽ʽ, 0)) > 0) Then
                        .TextMatrix(lngCurRow, COL_������λ) = Nvl(rsItems!���㵥λ)
                    End If
                End If
                
                '����
                If InStr(",5,6,", rsItems!���) > 0 Then
                    '��ҩ����(�ж�Ӧ���)
                    .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow, COL_���ﵥλ)
                                        
                    sng���� = Nvl(rsItems!����, msng����)
                    If mbln���� Then
                        If .TextMatrix(lngCurRow, COL_�����λ) = "��" Then
                            If 7 > sng���� Then sng���� = 7
                        ElseIf .TextMatrix(lngCurRow, COL_�����λ) = "��" Then
                            If Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)) > sng���� Then
                                sng���� = Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��))
                            End If
                        ElseIf .TextMatrix(lngCurRow, COL_�����λ) = "Сʱ" Then
                            If Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)) \ 24 > sng���� Then
                                sng���� = Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)) \ 24
                            End If
                        ElseIf .TextMatrix(lngCurRow, COL_�����λ) = "����" Then
                            If sng���� = 0 Then sng���� = 1
                        End If
                        If sng���� = 0 Then sng���� = 1
                    End If
                    
                    If Not IsNull(rsItems!�ܸ�����) Then
                        'ת��Ϊ���ﵥλ
                        .TextMatrix(lngCurRow, COL_����) = FormatEx(rsItems!�ܸ����� / Val(.TextMatrix(lngCurRow, COL_�����װ)), 5)
                    ElseIf .TextMatrix(lngCurRow, COL_Ƶ��) <> "" Then
                        '����ȱʡ����
                        rs�Ƴ�.Filter = "������ĿID=" & rsItems!������ĿID
                        If Not rs�Ƴ�.EOF Then
                            If Nvl(rs�Ƴ�!�Ƴ�, 1) > sng���� Then sng���� = Nvl(rs�Ƴ�!�Ƴ�, 1)
                        End If
                    
                        If (Val(.TextMatrix(lngCurRow, COL_����)) <> 0 _
                            And Val(.TextMatrix(lngCurRow, COL_�����װ)) <> 0 _
                            And Val(.TextMatrix(lngCurRow, COL_����ϵ��)) <> 0) Then
                            
                            .TextMatrix(lngCurRow, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                    Val(.TextMatrix(lngCurRow, COL_����)), sng����, _
                                    Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                                    Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), _
                                    .TextMatrix(lngCurRow, COL_�����λ), _
                                    .TextMatrix(lngCurRow, COL_ִ��ʱ��), _
                                    Val(.TextMatrix(lngCurRow, COL_����ϵ��)), _
                                    Val(.TextMatrix(lngCurRow, COL_�����װ)), _
                                    Val(.TextMatrix(lngCurRow, COL_�ɷ����))), 5)
                            If Val(.TextMatrix(lngCurRow, COL_�ɷ����)) <> 0 Then
                                .TextMatrix(lngCurRow, COL_����) = IntEx(Val(.TextMatrix(lngCurRow, COL_����)))
                            End If
                        End If
                    End If
                    
                    If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") = 0 Then
                        .TextMatrix(lngCurRow, COL_����) = IntEx(Val(.TextMatrix(lngCurRow, COL_����)))
                    End If
                    
                    '��������
                    If Val(.TextMatrix(lngCurRow, COL_��������)) <> 0 Then
                        If Val(.TextMatrix(lngCurRow, COL_����)) > FormatEx(Val(.TextMatrix(lngCurRow, COL_��������)) / Val(.TextMatrix(lngCurRow, COL_����ϵ��)) / Val(.TextMatrix(lngCurRow, COL_�����װ)), 5) Then
                            .TextMatrix(lngCurRow, COL_�Ƿ���) = "1"
                        End If
                    End If
                     
                    If mbln���� Then
                        .TextMatrix(lngCurRow, COL_����) = IIF(sng���� = 0, "", sng����)
                    End If
                    Call Set��ҩ�����Ƿ���(lngCurRow)
                    
                ElseIf bln�䷽ Then
                    If rsItems!��� = "7" Then
                        .TextMatrix(lngCurRow, COL_������λ) = "��"
                                                
                        If Not IsNull(rsItems!�ܸ�����) Then
                            .TextMatrix(lngCurRow, COL_����) = rsItems!�ܸ�����
                        ElseIf .TextMatrix(lngCurRow, COL_Ƶ��) <> "" Then
                             .TextMatrix(lngCurRow, COL_����) = CalcȱʡҩƷ����(1, 1, Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                                        Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), .TextMatrix(lngCurRow, COL_�����λ))
                        End If
                    Else
                        '��ҩ�巨,�÷������������ҩ��ͬ(Ϊ����ʾ)
                        .TextMatrix(lngCurRow, COL_����) = .TextMatrix(lngCurRow - 1, COL_����)
                        .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow - 1, COL_������λ)
                         
                    End If
                Else
                    '������������Ҫ����
                    '���Ϊһ���Ի�ƴ�����ȱʡ����Ϊ1
                    If Not IsNull(rsItems!�ܸ�����) Then
                        vsAdvice.TextMatrix(lngCurRow, COL_����) = rsItems!�ܸ�����
                    ElseIf intƵ������ = 1 Or Nvl(rsItems!���㷽ʽ, 0) = 3 Then
                        vsAdvice.TextMatrix(lngCurRow, COL_����) = 1
                    End If
                    If rsItems!��� = "4" Then
                        .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow, COL_���ﵥλ) 'ɢװ��λ
                    Else
                        .TextMatrix(lngCurRow, COL_������λ) = Nvl(rsItems!���㵥λ)
                    End If
                End If
                
                '����ҩ��ȱʡ��ҩĿ��
                If Val(.TextMatrix(lngCurRow, COL_�����ȼ�)) > 0 Then .TextMatrix(lngCurRow, COL_��ҩĿ��) = mstrPurMed
                        
                'ִ��ʱ��(����,Ƶ��,ִ��ʱ��֮��)
                If .TextMatrix(lngCurRow, COL_Ƶ��) <> "" Then
                    '�������ȱʡִ��ʱ�䷽��
                    If bln��ҩ;�� Or bln��ҩ�÷� Then
                        If Not IsNull(rsItems!ʱ�䷽��) Then
                            If ExeTimeValid(rsItems!ʱ�䷽��, Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                                Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), .TextMatrix(lngCurRow, COL_�����λ)) Then
                                .TextMatrix(lngCurRow, COL_ִ��ʱ��) = rsItems!ʱ�䷽��
                            End If
                        End If
                        If .TextMatrix(lngCurRow, COL_ִ��ʱ��) = "" Then
                            .TextMatrix(lngCurRow, COL_ִ��ʱ��) = Getȱʡʱ��(int���÷�Χ, .TextMatrix(lngCurRow, COL_Ƶ��), rsItems!������ĿID)
                        End If
                    ElseIf intƵ������ = 0 Then
                        If Not IsNull(rsItems!ʱ�䷽��) Then
                            If ExeTimeValid(rsItems!ʱ�䷽��, Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                                Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), .TextMatrix(lngCurRow, COL_�����λ)) Then
                                .TextMatrix(lngCurRow, COL_ִ��ʱ��) = rsItems!ʱ�䷽��
                            End If
                        End If
                        If .TextMatrix(lngCurRow, COL_ִ��ʱ��) = "" Then
                            .TextMatrix(lngCurRow, COL_ִ��ʱ��) = Getȱʡʱ��(int���÷�Χ, .TextMatrix(lngCurRow, COL_Ƶ��))
                        End If
                    End If
                    If bln�ɼ����� Then
                        .TextMatrix(lngCurRow, COL_�÷�) = rsItems!����
                    ElseIf bln��ҩ;�� Or bln��ҩ�÷� Then
                        '��ҩ����ҩ�䷽���÷�,ִ��ʱ��
                        If bln��ҩ�÷� Then
                            .TextMatrix(lngCurRow, COL_�÷�) = rsItems!����
                        End If
                        For j = lngCurRow - 1 To lngRow Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = .RowData(lngCurRow) Then
                                If bln��ҩ;�� Then .TextMatrix(j, COL_�÷�) = rsItems!���� & rsItems!ҽ������
                                .TextMatrix(j, COL_ִ��ʱ��) = .TextMatrix(lngCurRow, COL_ִ��ʱ��)
                            Else
                                Exit For
                            End If
                        Next
                    ElseIf bln��Ѫ;�� Then
                        .TextMatrix(lngCurRow - 1, COL_�÷�) = rsItems!����
                    End If
                End If
                
                '����ҽ���Ϳ�������
                .TextMatrix(lngCurRow, COL_����ҽ��) = UserInfo.����
                .TextMatrix(lngCurRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
                                    
                'ִ������
                If InStr(",5,6,7,", rsItems!���) > 0 Then
                    If Nvl(rsItems!ִ������, 0) = 5 Then
                        .TextMatrix(lngCurRow, COL_ִ������) = 5
                    Else
                        .TextMatrix(lngCurRow, COL_ִ������) = 4
                    End If
                ElseIf rsItems!��� = "4" Then
                    .TextMatrix(lngCurRow, COL_ִ������) = 4
                ElseIf bln��ҩ;�� Or bln��ҩ�巨 Or bln��ҩ�÷� Or bln�ɼ����� Then
                    .TextMatrix(lngCurRow, COL_ִ������) = Nvl(rsItems!ִ������, 0)
                Else
                    .TextMatrix(lngCurRow, COL_ִ������) = Nvl(rsItems!ִ������, 0)
                End If
                
                'ִ�п���ID:Ϊ0-����,5-Ժ��ִ��ʱȡ��Ϊ0
                If rsItems!��� = "Z" And InStr(",1,2,", Nvl(rsItems!��������, 0)) > 0 Then
                    If Nvl(rsItems!ִ�п���ID, 0) <> 0 Then
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = Nvl(rsItems!ִ�п���ID, 0)
                    Else
                        '���ۻ�סԺҽ��ȡ�ٴ�����(����ִ������)
                        If Nvl(rsItems!��������, 0) = 1 Then
                            '����:���������סԺ�ٴ�����
                            Call Get�ٴ�����(3, , lngTmp, , True, False, True)
                        ElseIf Nvl(rsItems!��������, 0) = 2 Then
                            'סԺ:����סԺ�ٴ�����
                            Call Get�ٴ�����(2, , lngTmp, , True, False, True)
                        End If
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = lngTmp
                    End If
                ElseIf InStr(",0,5,", Val(.TextMatrix(lngCurRow, COL_ִ������))) = 0 Then
                    If Nvl(rsItems!ִ�п���ID, 0) <> 0 Then
                        If InStr(",5,6,7,", rsItems!���) > 0 Then
                            strҩ��IDs = Get����ҩ��IDs(rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), mlng���˿���id, 1)
                            If InStr("," & strҩ��IDs & ",", "," & rsItems!ִ�п���ID & ",") > 0 Then
                                .TextMatrix(lngCurRow, COL_ִ�п���ID) = Nvl(rsItems!ִ�п���ID, 0)
                            End If
                        ElseIf rsItems!��� = "4" Then
                            strҩ��IDs = Get���÷��ϲ���IDs(Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), mlng���˿���id, 1)
                            If InStr("," & strҩ��IDs & ",", "," & rsItems!ִ�п���ID & ",") > 0 Then
                                .TextMatrix(lngCurRow, COL_ִ�п���ID) = Nvl(rsItems!ִ�п���ID, 0)
                            End If
                        ElseIf Val(.TextMatrix(lngCurRow, COL_ִ������)) = 4 Then
                            '4-ָ������ʱ��ȡ,�����Ĺ̶�����
                            .TextMatrix(lngCurRow, COL_ִ�п���ID) = Nvl(rsItems!ִ�п���ID, 0)
                            
                            '���ִ�п��ҵ���Ч��
                            If Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)) <> 0 Then
                                If CheckExecDeptValidate(Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)), mlng���˿���id, 1, Val(.TextMatrix(lngCurRow, COL_������ĿID))) = False Then
                                    .TextMatrix(lngCurRow, COL_ִ�п���ID) = 0
                                End If
                            End If
                        End If
                    End If
                    If Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)) = 0 Then
                        'ҩƷ�������������ͬ
                        If rsItems!��� = "5" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(lngCurRow, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf rsItems!��� = "6" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(lngCurRow, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf rsItems!��� = "7" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(lngCurRow, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf rsItems!��� = "4" Then
                            If lng���ϲ���ID = 0 Then
                                lng���ϲ���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsItems!���, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, , 1)
                            End If
                            .TextMatrix(lngCurRow, COL_ִ�п���ID) = lng���ϲ���ID
                        Else
                            '֮ǰ����������ID
                            .TextMatrix(lngCurRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, 0, _
                                Val(.TextMatrix(lngCurRow, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(lngCurRow, COL_��������ID)), 1, 1)
                        End If
                    End If
                End If
                
                'ҽ������
                .TextMatrix(lngCurRow, COL_ҽ������) = Nvl(rsItems!ҽ������)
                .Cell(flexcpData, lngCurRow, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), "", 0, "", .TextMatrix(lngCurRow, COL_������ĿID) & "|1")
                
                
                '����ʱ��
                .TextMatrix(lngCurRow, COL_����ʱ��) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngCurRow, COL_����ʱ��) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                
                '������־
                .TextMatrix(lngCurRow, COL_��־) = chk����.value '�����ڽ�����ͳһ����Ϊ����
                If gblnKSSStrict And UserInfo.��ҩ���� < Val(.TextMatrix(lngCurRow, COL_�����ȼ�)) And .TextMatrix(lngCurRow, COL_��־) <> "1" Then

                    .TextMatrix(lngCurRow, COL_���״̬) = 1
                End If
                
                If .TextMatrix(lngCurRow, COL_���) = "E" And .TextMatrix(lngCurRow, COL_��������) = "8" And Val(.TextMatrix(lngCurRow, COL_���ID)) <> 0 Then
                    strSQL = ""
                    strSQL = GetBloodState(IIF(.TextMatrix(lngCurRow, COL_��־) = "1", 1, 0), Val(.TextMatrix(lngCurRow, COL_ִ�з���)))
                    .TextMatrix(lngCurRow - 1, COL_���״̬) = strSQL
                    .TextMatrix(lngCurRow, COL_���״̬) = strSQL
                End If
                Call SetRow��־ͼ��(lngCurRow, 1)
                
                '��ȡҩƷ���
                If InStr(",5,6,7,", .TextMatrix(lngCurRow, COL_���)) > 0 Or .TextMatrix(lngCurRow, COL_���) = "4" And Val(.TextMatrix(lngCurRow, COL_��������)) = 1 Then
                    If Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)) <> 0 And Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)) <> 0 Then
                        .TextMatrix(lngCurRow, COL_���) = GetStock(Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)), 1)
                    End If
                End If
                
                '----------------------
                '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                If InStr(",5,6,", .TextMatrix(lngCurRow, COL_���)) > 0 And .TextMatrix(lngCurRow, COL_�������) <> "" Then
                    If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(lngCurRow, COL_�������)) > 0 Then
                        .Cell(flexcpFontBold, lngCurRow, col_ҽ������) = True
                    End If
                End If
                
                '����һЩ������
                If (InStr(",F,G,D,7,E,C,", rsItems!���) > 0 And Not IsNull(rsItems!������)) Or bln��ҩ;�� Then
                    .RowHidden(lngCurRow) = True
                End If
                
                'ҽ������
                If Not .RowHidden(lngCurRow) Then
                    If IsNull(rsItems!������ĿID) Then
                        .TextMatrix(lngCurRow, col_ҽ������) = rsItems!ҽ������ '����¼��ҽ��
                    ElseIf InStr(",F,D,", rsItems!���) > 0 And IsNull(rsItems!������) Then
                        .TextMatrix(lngCurRow, col_ҽ������) = rsItems!���� '��ʱ
                    Else
                        .TextMatrix(lngCurRow, col_ҽ������) = AdviceTextMake(lngCurRow)
                    End If
                Else
                    .TextMatrix(lngCurRow, col_ҽ������) = rsItems!����
                End If
                
                '��������ʱ���ڿɼ��ж�ȡ������Ŀ��������ڼ�飬�����ڱ���
                If Not .RowHidden(lngCurRow) Then
                    If Not bln��ҩ�÷� Then
                        If bln�ɼ����� Then
                            j = .FindRow(CStr(.RowData(lngCurRow)), , COL_���ID)
                            If j <> -1 Then
                                .TextMatrix(lngCurRow, COL_����) = Getҽ����Ŀ����(Val(.TextMatrix(j, COL_������ĿID)), 1)
                            End If
                        Else
                            .TextMatrix(lngCurRow, COL_����) = Getҽ����Ŀ����(Val(.TextMatrix(lngCurRow, COL_������ĿID)), 1)
                        End If
                        If .TextMatrix(lngCurRow, COL_����) <> "" Then
                            str���� = str���� & vbCrLf & "��" & .TextMatrix(lngCurRow, col_ҽ������)
                        End If
                    End If
                End If
                
                If lngPreRow = -1 And Not .RowHidden(lngCurRow) Then lngPreRow = lngCurRow
                            
                '----------------------
                intCount = intCount + 1
            End If
            rsItems.MoveNext
        Next
        
        '--------------------------------------------------
        '�������Ӵ���
        For i = lngRow To lngCurRow
            'ȡ����������ҽ������
            If InStr(",F,D,", .TextMatrix(i, COL_���)) > 0 And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                .TextMatrix(i, col_ҽ������) = AdviceTextMake(i)
            End If
            
            '�������Ƶ���
            If Not .RowHidden(i) And .TextMatrix(i, COL_����) = "" Then
                .TextMatrix(i, COL_����) = GetItemPrice(i)
            End If
        Next
        
        '������Ӱ���е����
        Call AdviceSetҽ�����(lngCurRow + 1, intCount)
        '������ҽ��ID
        For i = lngRow To lngCurRow
            dbl���ID = .RowData(i)
            .RowData(i) = GetNextҽ��ID
            For j = i - 1 To lngRow Step -1
                If Val(.TextMatrix(j, COL_���ID)) = dbl���ID Then
                    .TextMatrix(j, COL_���ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            For j = i + 1 To lngCurRow
                If Val(.TextMatrix(j, COL_���ID)) = dbl���ID Then
                    .TextMatrix(j, COL_���ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            
            '�����������ϵı�Ǵ���
            If Val(.TextMatrix(i, COL_���ID)) = 0 Then '�����д����Ϊ׼
                Call SetDiagFlag(i, 1)
            End If
        Next
        If gblnOut���� Then Call MakeAppNo(2, lngRow, lngCurRow)
        Call Setҽ������(lngRow, lngCurRow)
        
        '--------------------------------------------------
        If .RowHidden(lngRow) Then 'Ѱ�ҿɼ���(���䷽�ͼ���֮��)
            For i = lngRow + 1 To .Rows - 1
                If Not .RowHidden(i) And .RowData(i) <> 0 Then
                    lngRow = i: Exit For
                End If
            Next
        End If
        '��λ�����׷����ĵ�һ�У�����Ϊҽ�����ó��׺������޸�ҽ������һ��֮���ҽ��Ϊ���׷������õ�ҽ��
        .Row = lngRow: .Col = col_ҽ������
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        mblnRowChange = True
    End With
    Screen.MousePointer = 0
    
    If str���� <> "" Then
        MsgBox "����ҽ����Ҫ��д���븽���ע����д��" & vbCrLf & str����, vbInformation, gstrSysName
    End If
    If str��ΣҩƷ <> "" Then
        MsgBox "����ҽ���Ǹ�ΣҩƷ��" & str��ΣҩƷ, vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPreRow��ҩ��(lngRow As Long) As Long
'���ܣ���ȡ��һ����һ��ҩ�е�ִ�п���
    Dim lngDrugRow As Long, lngCopyRow As Long
    
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn�䷽��(lngCopyRow) Then
            '�����һ��Ч������ҩ�䷽��,��ȡ���ĵ�һ��ҩ��
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_���ID)
        End If
    End If
    
    If lngDrugRow <> -1 Then 'ȱʡ����һ�䷽����ͬ
        GetPreRow��ҩ�� = Val("" & vsAdvice.TextMatrix(lngDrugRow, COL_ִ�п���ID))
    End If
End Function

Private Function AdviceSet��ҩ�䷽(lng������ĿID As Long, ByVal lngRow As Long, ByVal lng�÷�ID As Long, _
    ByVal strExtData As String, Optional rsCurr As ADODB.Recordset, Optional ByVal strժҪ As String, Optional ByVal lng�䷽ID As Long) As Long
'���ܣ�(����)������ҩ�䷽��ȱʡҽ������
'������lng������ĿID=�������ҩ�䷽ID��ζ��ҩID
'      lngRow=��ǰ������
'      lng�÷�ID=ȱʡ��ҩ�÷�ID
'      strExtData=�����䷽���ζҩ���巨����:���ID1,����,��ע;���ID2,����,��ע...|��ҩ�巨|��ҩ��̬|����|ҩ��ID|����"
'      rsCurr=������޸����䷽���ݺ����,�����Ҫ���ֵ�һЩ��ǰֵ
'      strժҪ=ҽ��ժҪ
'���أ���������ҩ�䷽�ĵ�ǰ��ʾ�к�
    Dim rsItems As New ADODB.Recordset '��ҩ��ϸ��Ϣ
    Dim rsUse As New ADODB.Recordset '��ҩ�÷���Ϣ
    Dim rs�巨 As New ADODB.Recordset '��ҩ�巨��Ŀ��Ϣ
    Dim rs�÷� As New ADODB.Recordset '��ҩ�÷���Ŀ��Ϣ
    Dim arr��ҩs As Variant, str��ҩIDs As String, lng���ID As Long
    Dim lngCopyRow As Long 'ȱʡ������
    Dim lngDrugRow As Long '���ȱʡ����������ҩ�䷽,��Ϊ���䷽�ĵ�һ����ҩ��
    Dim lngFirstRow As Long '��ǰ�䷽�ĵ�һ����ҩ��
    Dim strSQL As String, i As Long
    Dim blnOutOfRange As Boolean
    
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng�巨ID As Long, int�Ƴ� As Integer
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim lng��̬ As Long
    Dim str��ΣҩƷ As String, str���� As String
        
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn�䷽��(lngCopyRow) Then
            '�����һ��Ч������ҩ�䷽��,��ȡ���ĵ�һ��ҩ��
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_���ID)
        End If
    End If
    
    '��ȡ������ݿ���Ϣ
    '------------------
    arr��ҩs = Split(Split(strExtData, "|")(0), ";")
    For i = 0 To UBound(arr��ҩs)
        str��ҩIDs = str��ҩIDs & "," & CStr(Split(arr��ҩs(i), ",")(0))
    Next
    str��ҩIDs = Mid(str��ҩIDs, 2)
    lng�巨ID = Val(Split(strExtData, "|")(1))
    lng��̬ = Val(Split(strExtData, "|")(2))
    str���� = Split(strExtData, "|")(5)

    
    '�䷽�÷���Ϣ:ֱ�������䷽ʱ���п�����,���뵥ζ��ҩ��
    strSQL = "Select A.�÷�ID,A.Ƶ��,A.�Ƴ�,A.ҽ������" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And B.������� IN(1,3)" & _
        " And Nvl(A.����,0)=0 And A.��ĿID=[1]" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
        " And (b.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or b.����ʱ�� is NULL)"
    Set rsUse = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng������ĿID)
    If Not rsUse.EOF Then lng�÷�ID = rsUse!�÷�ID 'ȱʡ���õ���ҩ�䷽�÷�����
    
    '�䷽���ζ��ҩ��Ϣ:��ҩ�޹�����,��Ӧ�ĵĹ���¼һ������ֻ��һ��
    strSQL = "Select /*+ rule*/A.�������,A.վ��,A.���,A.����ID,A.ID,A.����,A.����,A.�걾��λ,A.���㵥λ,A.���㷽ʽ,A.ִ��Ƶ��,A.�����Ա�," & _
        "A.����Ӧ��,A.�����Ŀ,A.��������,A.ִ�а���,A.ִ�п���,A.�������,A.�Ƽ�����,A.�ο�Ŀ¼ID,A.��ԱID,A.����ʱ��,A.����ʱ��," & _
        "A.¼������,A.�Թܱ���,A.ִ�з���,A.ִ�б��,B.ҩƷID,B.����ϵ��,B.�����װ,B.���ﵥλ,B.����ɷ���� As �ɷ����,C.��������,C.����ְ��,c.�ٴ��Թ�ҩ,b.��ΣҩƷ" & _
        " From ������ĿĿ¼ A,ҩƷ��� B,ҩƷ���� C" & _
        " Where A.ID=B.ҩ��ID And A.ID=C.ҩ��ID And B.ҩƷID IN(Select Column_Value From Table(f_Num2list([1])))"
    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��ҩIDs)
    
    '�䷽�巨��Ŀ��Ϣ
    Set rs�巨 = Get������Ŀ��¼(lng�巨ID)
    
    '�䷽�÷���Ŀ��Ϣ
    Set rs�÷� = Get������Ŀ��¼(lng�÷�ID)
    
    '�����䷽���ζ��ҩ��:�����û�����˳��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    mblnRowChange = False
    
    '��ҩ�÷���ҽ��ID,ID˳������Ų�һ��һ��
    If Not rsCurr Is Nothing Then
        '�޸����䷽�е�����,�÷��б��Ϊ�޸�,ҽ��ID����
        lng���ID = rsCurr!ҽ��ID
    Else
        '���������ҩ�䷽
        lng���ID = GetNextҽ��ID
    End If
    
    For i = 0 To UBound(arr��ҩs)
        rsItems.Filter = "ҩƷID=" & CStr(Split(arr��ҩs(i), ",")(0)) 'Ӧ�ÿ϶���
        
        vsAdvice.AddItem "", lngRow
        
        vsAdvice.RowHidden(lngRow) = True
        vsAdvice.RowData(lngRow) = GetNextҽ��ID
        vsAdvice.TextMatrix(lngRow, COL_���ID) = lng���ID '��Ӧ���������ҩ�÷���
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '����
        vsAdvice.TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
        vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�¿�
        vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1) '�������
        
        vsAdvice.TextMatrix(lngRow, COL_���) = rsItems!���
        vsAdvice.TextMatrix(lngRow, col_ҽ������) = rsItems!����
        vsAdvice.TextMatrix(lngRow, COL_������ĿID) = rsItems!ID
        vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsItems!���㷽ʽ, 0)
        vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsItems!ִ��Ƶ��, 0)
        vsAdvice.TextMatrix(lngRow, COL_��������) = Nvl(rsItems!��������)
        vsAdvice.TextMatrix(lngRow, COL_����Ӧ��) = Nvl(rsItems!����Ӧ��)
        vsAdvice.TextMatrix(lngRow, COL_ִ�з���) = Nvl(rsItems!ִ�з���, 0)
        
        vsAdvice.TextMatrix(lngRow, COL_����) = FormatEx(Val(Split(arr��ҩs(i), ",")(1)), 5) '��ζҩ�ĵ�������
        vsAdvice.TextMatrix(lngRow, COL_������λ) = Nvl(rsItems!���㵥λ)
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = CStr(Split(arr��ҩs(i), ",")(2)) '��ζҩ�Ľ�ע
        vsAdvice.Cell(flexcpData, lngRow, COL_ҽ������) = strժҪ
        
        '�����Ϣ:��ҩ�����ڹ�����,һ����
        vsAdvice.TextMatrix(lngRow, COL_�շ�ϸĿID) = rsItems!ҩƷID
        vsAdvice.TextMatrix(lngRow, COL_��������) = Nvl(rsItems!��������)
        vsAdvice.TextMatrix(lngRow, COL_����ϵ��) = rsItems!����ϵ��
        vsAdvice.TextMatrix(lngRow, COL_���ﵥλ) = rsItems!���ﵥλ
        vsAdvice.TextMatrix(lngRow, COL_�����װ) = rsItems!�����װ
        vsAdvice.TextMatrix(lngRow, COL_�ɷ����) = Nvl(rsItems!�ɷ����, 0) '����ҩʵ��������
        vsAdvice.TextMatrix(lngRow, COL_����ְ��) = Nvl(rsItems!����ְ��)
        vsAdvice.TextMatrix(lngRow, COL_�ٴ��Թ�ҩ) = rsItems!�ٴ��Թ�ҩ & ""
        vsAdvice.TextMatrix(lngRow, COL_��ΣҩƷ) = rsItems!��ΣҩƷ & ""
        If Val(vsAdvice.TextMatrix(lngRow, COL_��ΣҩƷ)) <> 0 Then
            str��ΣҩƷ = str��ΣҩƷ & vbCrLf & vsAdvice.TextMatrix(lngRow, col_ҽ������) & ":" & Decode(Val(vsAdvice.TextMatrix(lngRow, COL_��ΣҩƷ)), 1, "A", 2, "B", 3, "C", "") & "����"
        End If
        vsAdvice.Cell(flexcpData, lngRow, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, Val(vsAdvice.TextMatrix(lngRow, COL_�շ�ϸĿID)), "", 0, "", vsAdvice.TextMatrix(lngRow, COL_������ĿID))

        '�Ƽ�����:���Զ���
        vsAdvice.TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsItems!�Ƽ�����, 0)
        
        If lngFirstRow <> 0 Then
            '����һ�������õ������ҩ��ͬ
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ������)
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ�п���ID)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
            vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
            vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
            vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_��ʼʱ��)
            vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ҽ��)
            vsAdvice.TextMatrix(lngRow, COL_��������ID) = vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)
            
            vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ʱ��)
            vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_����ʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_��־) = vsAdvice.TextMatrix(lngFirstRow, COL_��־)
        ElseIf Not rsCurr Is Nothing Then
            '�޸����䷽���ݺ���������,�����뵱ǰ��ֵ
            
            'ִ������:�޸�ʱ���ݵ�ǰ�������þ���
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = Decode(Nvl(rsCurr!ִ������), "�Ա�ҩ", 5, 4)
            'ִ�п���
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_ִ������)) = 5, 0, Val("" & rsCurr!ִ�п���ID))
            
            vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = Nvl(rsCurr!Ƶ��)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = Nvl(rsCurr!Ƶ�ʴ���)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = Nvl(rsCurr!Ƶ�ʼ��)
            vsAdvice.TextMatrix(lngRow, COL_�����λ) = Nvl(rsCurr!�����λ)
            vsAdvice.TextMatrix(lngRow, COL_����) = Nvl(rsCurr!����)
            vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = Nvl(rsCurr!ִ��ʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = Format(Nvl(rsCurr!��ʼʱ��), "yyyy-MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = CStr(Nvl(rsCurr!��ʼʱ��))
            
            vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = Nvl(rsCurr!����ҽ��)
            vsAdvice.TextMatrix(lngRow, COL_��������ID) = Nvl(rsCurr!��������id)
            
            vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsCurr!����ʱ��), "yyyy-MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = CStr(Nvl(rsCurr!����ʱ��))
            
            vsAdvice.TextMatrix(lngRow, COL_��־) = Nvl(rsCurr!��־)
        Else
            'ִ������:��ҩ�䷽�����ҩ��ͬ,ȱʡ=4-ָ������,5-�Ա�ҩ
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_�ٴ��Թ�ҩ)) = 1, 5, 4)
        
            'ִ�п���(�����䷽����ѡ��)
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_ִ������)) = 5, 0, Val(Split(strExtData, "|")(4)))
            
            'ִ��Ƶ��
            '�����÷��������õ�����
            If Not rsUse.EOF Then
                If Not IsNull(rsUse!Ƶ��) Then
                    Call GetƵ����Ϣ_����(rsUse!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                    vsAdvice.TextMatrix(lngRow, COL_�����λ) = str�����λ
                End If
            End If
            '��ȱʡ����һ����ͬ
            If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = "" And lngDrugRow <> -1 Then
                If Val(vsAdvice.TextMatrix(lngDrugRow, COL_EDIT)) = 1 And vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��) <> "" Then
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʴ���)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʼ��)
                    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngDrugRow, COL_�����λ)
                End If
            End If
            '��ȡȱʡֵ
            If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = "" Then
                Call GetȱʡƵ��(Nvl(rsItems!ID, 0), 2, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                vsAdvice.TextMatrix(lngRow, COL_�����λ) = str�����λ
            End If
            
            '����(����):��ɢװ��̬��ȷ������
            If Val(Split(strExtData, "|")(3)) > 1 Or lng��̬ <> 0 Then
                vsAdvice.TextMatrix(lngRow, COL_����) = Val(Split(strExtData, "|")(3))
            Else
                If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) <> "" Then
                    int�Ƴ� = 1
                    If Not rsUse.EOF Then int�Ƴ� = Nvl(rsUse!�Ƴ�, 1)
                    '�䷽����
                    vsAdvice.TextMatrix(lngRow, COL_����) = CalcȱʡҩƷ����(1, int�Ƴ�, _
                            Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���)), _
                            Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��)), _
                            vsAdvice.TextMatrix(lngRow, COL_�����λ))
                End If
            End If
            
            'ִ��ʱ��
            If lngDrugRow <> -1 Then 'ȱʡ����һ����ͬ
                If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��) Then
                    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngDrugRow, COL_ִ��ʱ��)
                End If
            End If
            If vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then 'ȱʡʱ�䷽��
                vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(2, vsAdvice.TextMatrix(lngRow, COL_Ƶ��), lng�÷�ID)
            End If
            
            '��ʼʱ��
            If IsDate(txt��ʼʱ��.Text) Then
                vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
            End If
            
            '����ҽ��
            vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
            vsAdvice.TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
            
            vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            vsAdvice.TextMatrix(lngRow, COL_��־) = chk����.value
        End If
        
        '---------------------------------------
        If lngFirstRow = 0 Then lngFirstRow = lngRow '����ҩ�䷽�ĵ�һ�������ҩ��
        
        lngRow = lngRow + 1 '���ֵ�ǰ������λ��
    Next
    
    '������ҩ�䷽�巨��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.AddItem "", lngRow
    vsAdvice.RowHidden(lngRow) = True
    vsAdvice.RowData(lngRow) = GetNextҽ��ID
    vsAdvice.TextMatrix(lngRow, COL_���ID) = lng���ID
    vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '����
    vsAdvice.TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�¿�
    vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
    Call AdviceSetҽ�����(lngRow + 1, 1) '�������
    vsAdvice.TextMatrix(lngRow, COL_���) = rs�巨!���
    vsAdvice.TextMatrix(lngRow, COL_������ĿID) = lng�巨ID
    vsAdvice.TextMatrix(lngRow, COL_�걾��λ) = str����
    vsAdvice.Cell(flexcpData, lngRow, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", vsAdvice.TextMatrix(lngRow, COL_������ĿID))
    
    vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rs�巨!���㷽ʽ, 0)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = Nvl(rs�巨!ִ��Ƶ��, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������) = Nvl(rs�巨!��������)
    vsAdvice.TextMatrix(lngRow, COL_����Ӧ��) = Nvl(rs�巨!����Ӧ��)
    vsAdvice.TextMatrix(lngRow, COL_ִ�з���) = Nvl(rs�巨!ִ�з���, 0)
    
    '!��ҩ�巨��Ҳ�����ҩ�ĸ���
    vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
    
    vsAdvice.TextMatrix(lngRow, col_ҽ������) = rs�巨!����
    
    vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_��ʼʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
    
    'ִ������:ȱʡ������Ŀ����(������ΪԺ��ִ��),�޸�ʱ���ݵ�ǰ��������
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Nvl(rs�巨!ִ�п���, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Decode(Nvl(rsCurr!ִ������), "��Ժ��ҩ", 5, Nvl(rs�巨!ִ�п���, 0))
    End If
    
    '��ҩ�巨���δ����ִ�п���,��ȱʡΪ�������ڲ���(����Ҫ��Ϊ�������ڿ���!!)
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_ִ������))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rs�巨!���, lng�巨ID, 0, _
            Nvl(rs�巨!ִ�п���, 0), mlng���˿���id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)), 1, 1)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rs�巨!�Ƽ�����, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������ID) = vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)
    vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ҽ��)
    
    vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_����ʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_��־) = vsAdvice.TextMatrix(lngFirstRow, COL_��־)
    
    '���ֵ�ǰ������λ��
    lngRow = lngRow + 1
    
    '������ҩ�䷽�÷���:��ҩ�䷽����ʾ��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.RowData(lngRow) = lng���ID
    If Get������Ŀ��¼(lng������ĿID)!��� & "" = "8" Then
        vsAdvice.TextMatrix(lngRow, COL_�䷽ID) = lng������ĿID
    End If
    If lng�䷽ID <> 0 Then
        vsAdvice.TextMatrix(lngRow, COL_�䷽ID) = lng�䷽ID
    End If
    
    If Not rsCurr Is Nothing Then
        '�޸����䷽����,���Ϊ�޸�
        If InStr(",0,3,", rsCurr!Edit) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
        Else
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '���������������޸�
        End If
    Else
        '���������ҩ�䷽,Ϊ����
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�¿�
    vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
    Call AdviceSetҽ�����(lngRow + 1, 1) '�������
    vsAdvice.TextMatrix(lngRow, COL_���) = rs�÷�!���
    vsAdvice.TextMatrix(lngRow, COL_������ĿID) = lng�÷�ID
    vsAdvice.Cell(flexcpData, lngRow, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", vsAdvice.TextMatrix(lngRow, COL_������ĿID))
    
    vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rs�÷�!���㷽ʽ, 0)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = Nvl(rs�÷�!ִ��Ƶ��, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������) = Nvl(rs�÷�!��������)
    vsAdvice.TextMatrix(lngRow, COL_����Ӧ��) = Nvl(rs�÷�!����Ӧ��)
    vsAdvice.TextMatrix(lngRow, COL_ִ�з���) = Nvl(rs�÷�!ִ�з���, 0)
    
    '!��ҩ�÷���Ҳ�����ҩ�ĸ���
    vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
    vsAdvice.TextMatrix(lngRow, COL_������λ) = "��"
    
    vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_��ʼʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_����) = rs�÷�!����
    vsAdvice.TextMatrix(lngRow, COL_�÷�) = rs�÷�!����
    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
    
    'ִ������:ȱʡ������Ŀ����(������ΪԺ��ִ��),�޸�ʱ���ݵ�ǰ��������
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Nvl(rs�÷�!ִ�п���, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Decode(Nvl(rsCurr!ִ������), "��Ժ��ҩ", 5, Nvl(rs�÷�!ִ�п���, 0))
    End If
    
    '��ҩ�÷����δ����ִ�п���,��ȱʡΪ�������ڲ���(����Ҫ��Ϊ�������ڿ���!!)
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_ִ������))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rs�÷�!���, lng�÷�ID, 0, _
            Nvl(rs�÷�!ִ�п���, 0), mlng���˿���id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)), 1, 1)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rs�÷�!�Ƽ�����, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������ID) = vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)
    vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ҽ��)
    
    vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_����ʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_��־) = vsAdvice.TextMatrix(lngFirstRow, COL_��־)
    Call SetRow��־ͼ��(lngRow)
        
    If Not rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = Nvl(rsCurr!ҽ������)
    ElseIf Not rsUse.EOF Then
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = Nvl(rsUse!ҽ������)
    End If
    '��ҩ��̬(����AdviceTextMake��)
    vsAdvice.TextMatrix(lngRow, COL_��ҩ��̬) = lng��̬
    '�����������
    Call CheckCHLimited(lngRow, vsAdvice.TextMatrix(lngRow, COL_����), blnOutOfRange, vsAdvice, COL_���ID, COL_������ĿID, COL_���, COL_����)
    If blnOutOfRange Then vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "1"
    '��ҩ�䷽����ҩ���
    Call GetDrugStock(lngRow)
    
    '��ҩ�䷽ҽ������
    vsAdvice.TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
    
    '-------------------
    vsAdvice.Row = lngRow
    mblnRowChange = True
    
    If str��ΣҩƷ <> "" Then
        MsgBox "����ҽ���Ǹ�ΣҩƷ��" & str��ΣҩƷ, vbInformation, gstrSysName
    End If
        
    AdviceSet��ҩ�䷽ = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet�������(ByVal lngRow As Long, ByVal lng�ɼ�����ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset, Optional ByVal strժҪ As String, Optional ByVal bln���뵥 As Boolean) As Long
'���ܣ����������ļ���(���)
'������rsItems=�����ѡ�񷵻صļ�¼��
'      lngRow=��ǰ������
'      lng�ɼ�����ID=ȱʡ�Ĳɼ�����
'      strExtData=���:"��ĿID1,��ĿID2,...;����걾"������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾"
'      rsCurr=�޸ļ�����Ŀʱ��
'      strժҪ=ҽ��ժҪ
'      bln���뵥 ���뵥��������
'���أ�����֮��ĵ�ǰ��ʾ�к�
    Dim rsMore As New ADODB.Recordset '�ɼ�������Ϣ
    Dim rsItems As New ADODB.Recordset '������Ŀ��Ϣ
    Dim arrItems As Variant, strItems As String
    Dim strSQL As String, curDate As Date
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strƵ�� As String, intƵ�ʴ��� As Integer
    Dim intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng���ID As Long, strҽ������ As String
    Dim lngCopyRow As Long, lngFirstRow As Long, i As Long
    Dim rsLIS As New ADODB.Recordset
    Dim strTmp As String
    Dim Y As Long
    Dim blnLis As Boolean
    
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    '��ǰʱ��
    curDate = zlDatabase.Currentdate
    
    '������Ŀ��Ϣ
    '----------------------------------------------------------------------------
    '����������Ŀ��Ϣ:������˳��
    arrItems = Split(Split(strExtData, ";")(0), ",")
    For i = UBound(arrItems) To 0 Step -1
        If mblnNewLIS Then
            strTmp = arrItems(i)
            If InStr(strTmp, "|") > 0 Then
                For Y = 0 To UBound(Split(strTmp, "|"))
                    strItems = strItems & "," & Val(Split(strTmp, "|")(Y))
                    If Y > 0 Then
                        strSQL = strSQL & " Union All " & " Select '" & Val(Split(strTmp, "|")(Y)) & "' as ����,'" & Val(Split(strTmp, "|")(0)) & "' as ���� From Dual "
                    End If
                Next
            Else
                strItems = strItems & "," & Val(strTmp)
            End If
        Else
            strItems = strItems & "," & Val(arrItems(i))
        End If
    Next
    If strSQL <> "" Then
        Set rsLIS = zlDatabase.OpenSQLRecord(Mid(strSQL, 11), Me.Caption)
        blnLis = True
    End If
    Set rsItems = Get������Ŀ��¼(0, Mid(strItems, 2))
    
        If Not bln���뵥 Then
    'ȡĳ��������Ŀ�Ĳɼ�����
    strSQL = "Select /*+ RULE */ A.��ĿID,Nvl(A.����,0) as ���,A.�÷�ID" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And B.������� IN(1,3)" & _
        " And A.��ĿID IN(Select Column_Value From Table(f_Num2list([1])))" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
        " And (b.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or b.����ʱ�� is NULL)" & _
        " Order by A.��ĿID,Nvl(A.����,0)"
    Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strItems, 2))
    If Not rsMore.EOF Then
        If rsCurr Is Nothing Or lng�ɼ�����ID = 0 Then
            lng�ɼ�����ID = rsMore!�÷�ID '�޸�ʱ����
        End If
    End If
        End If

    Set rsMore = Get������Ŀ��¼(lng�ɼ�����ID)
    
    mblnRowChange = False
    
    '���ø��м�����Ŀ
    '----------------------------------------------------------------------------
    '�ɼ�����ҽ��ID,ID˳������Ų�һ��һ��
    If Not rsCurr Is Nothing Then
        '�޸��˼�������е�����,�ɼ������б��Ϊ�޸�,ҽ��ID����
        lng���ID = rsCurr!ҽ��ID
    Else
        '�������
        lng���ID = GetNextҽ��ID
    End If
    With vsAdvice
        For i = 1 To rsItems.RecordCount
            .AddItem "", lngRow
            
            .RowHidden(lngRow) = True
            .RowData(lngRow) = GetNextҽ��ID
            .TextMatrix(lngRow, COL_���ID) = lng���ID '��Ӧ���ɼ�������
            .TextMatrix(lngRow, COL_EDIT) = 1 '����
            .TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
            .TextMatrix(lngRow, COL_״̬) = 1 '�¿�
            
            .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
            Call AdviceSetҽ�����(lngRow + 1, 1) '�������
            
            .TextMatrix(lngRow, COL_���) = rsItems!���
            .TextMatrix(lngRow, col_ҽ������) = rsItems!����
            .TextMatrix(lngRow, COL_������ĿID) = rsItems!ID
            .TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsItems!���㷽ʽ, 0)
            .TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsItems!ִ��Ƶ��, 0)
            .TextMatrix(lngRow, COL_��������) = Nvl(rsItems!��������)
            .TextMatrix(lngRow, COL_����Ӧ��) = Nvl(rsItems!����Ӧ��)
            .TextMatrix(lngRow, COL_ִ�з���) = Nvl(rsItems!ִ�з���, 0)
            .TextMatrix(lngRow, COL_��������) = Nvl(rsItems!¼������)
            .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsItems!�Ƽ�����, 0)
            .TextMatrix(lngRow, COL_ִ������) = Nvl(rsItems!ִ�п���, 0)
            '����걾
            .TextMatrix(lngRow, COL_�걾��λ) = Split(strExtData, ";")(1)
            If mblnNewLIS And rsItems!ID & "" <> "" And blnLis Then
                rsLIS.Filter = "����=" & rsItems!ID
                If rsLIS.EOF = False Then
                    .TextMatrix(lngRow, COL_�����ĿID) = rsLIS!���� & ""
                End If
            End If
            
            .Cell(flexcpData, lngRow, COL_ҽ������) = strժҪ
            
            '��������һ���ɼ��ļ�����Ŀ��ͬ
            If lngFirstRow <> 0 Then
                .TextMatrix(lngRow, COL_����) = .TextMatrix(lngFirstRow, COL_����)
                
                'һ���ɼ��ļ�����ĿӦ����ͬ
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngFirstRow, COL_ִ�п���ID)
                End If
                .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngFirstRow, COL_Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngFirstRow, COL_�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngFirstRow, COL_ִ��ʱ��)
                
                .TextMatrix(lngRow, COL_��ʼʱ��) = .TextMatrix(lngFirstRow, COL_��ʼʱ��)
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = .Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
                
                .TextMatrix(lngRow, COL_����ҽ��) = .TextMatrix(lngFirstRow, COL_����ҽ��)
                .TextMatrix(lngRow, COL_��������ID) = .TextMatrix(lngFirstRow, COL_��������ID)
                
                .TextMatrix(lngRow, COL_����ʱ��) = .TextMatrix(lngFirstRow, COL_����ʱ��)
                .Cell(flexcpData, lngRow, COL_����ʱ��) = .Cell(flexcpData, lngFirstRow, COL_����ʱ��)
                
                .TextMatrix(lngRow, COL_��־) = .TextMatrix(lngFirstRow, COL_��־)
            ElseIf Not rsCurr Is Nothing Then
                .TextMatrix(lngRow, COL_����) = Nvl(rsCurr!����, 1)
                
                'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    If Nvl(rsCurr!ִ�п���ID, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = rsCurr!ִ�п���ID
                    Else
                        .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!ID, 0, _
                            Nvl(rsItems!ִ�п���, 0), mlng���˿���id, Nvl(rsCurr!��������id, 0), 1, 1)
                    End If
                End If
                
                'ִ��Ƶ��
                .TextMatrix(lngRow, COL_Ƶ��) = Nvl(rsCurr!Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = Nvl(rsCurr!Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = Nvl(rsCurr!Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = Nvl(rsCurr!�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = Nvl(rsCurr!ִ��ʱ��)
                
                'ʱ��/����/ҽ��
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(Nvl(rsCurr!��ʼʱ��), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = CStr(Nvl(rsCurr!��ʼʱ��))
                
                .TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsCurr!����ʱ��), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_����ʱ��) = CStr(Nvl(rsCurr!����ʱ��))
                
                .TextMatrix(lngRow, COL_����ҽ��) = Nvl(rsCurr!����ҽ��)
                .TextMatrix(lngRow, COL_��������ID) = Nvl(rsCurr!��������id)
                
                .TextMatrix(lngRow, COL_��־) = Nvl(rsCurr!��־)
            Else
                .TextMatrix(lngRow, COL_����) = 1 'ȱʡΪ1(��)
                
                '����ҽ��
                .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
                .TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
                
                'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    '֮ǰҪ�����������ID
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!ID, 0, _
                        Nvl(rsItems!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                End If
                
                'ִ��Ƶ��
                Call GetȱʡƵ��(Nvl(rsItems!ID, 0), GetƵ�ʷ�Χ(lngRow), strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
                
                'ִ��ʱ��:"��ѡƵ��"(ҩƷ�ǿ�ѡƵ��,����������Ϊһ����)
                If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Then
                    If lngCopyRow <> -1 Then '����һ����ͬ
                        If .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��) Then
                            .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngCopyRow, COL_ִ��ʱ��)
                        End If
                    End If
                    If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then  'ȱʡʱ�䷽��
                        .TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(1, .TextMatrix(lngRow, COL_Ƶ��))
                    End If
                End If
            
                '��ʼʱ��
                If IsDate(txt��ʼʱ��.Text) Then
                    .TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
                End If
                
                '����ʱ��
                .TextMatrix(lngRow, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
                
                '������־
                .TextMatrix(lngRow, COL_��־) = chk����.value
            End If
            
            strҽ������ = strҽ������ & "," & rsItems!���� 'ҽ������
            If lngFirstRow = 0 Then lngFirstRow = lngRow '��һ��Ŀ��
            lngRow = lngRow + 1 '���ֵ�ǰ������λ��
            
            rsItems.MoveNext
        Next
        
        '���ñ걾�Ĳɼ�����
        '----------------------------------------------------------------------------
        rsItems.MoveFirst
        .RowData(lngRow) = lng���ID
        
        If Not rsCurr Is Nothing Then
            '�޸��˼����������,���Ϊ�޸�
            If InStr(",0,3,", rsCurr!Edit) > 0 Then
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
            Else
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '���������������޸�
            End If
        Else
            '������ļ������,Ϊ����
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
        End If
        
        .TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
        .TextMatrix(lngRow, COL_״̬) = 1 '�¿�
        
        .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1) '�������
        
        .TextMatrix(lngRow, COL_���) = rsMore!���
        .TextMatrix(lngRow, COL_����) = rsMore!����
        .TextMatrix(lngRow, COL_�÷�) = rsMore!����
        .TextMatrix(lngRow, COL_������ĿID) = rsMore!ID
        .TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsMore!���㷽ʽ, 0)
        .TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsMore!ִ��Ƶ��, 0)
        .TextMatrix(lngRow, COL_��������) = Nvl(rsMore!��������)
        .TextMatrix(lngRow, COL_����Ӧ��) = Nvl(rsMore!����Ӧ��)
        .TextMatrix(lngRow, COL_ִ�з���) = Nvl(rsMore!ִ�з���, 0)
        .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsMore!�Ƽ�����, 0)
        .TextMatrix(lngRow, COL_�걾��λ) = .TextMatrix(lngFirstRow, COL_�걾��λ)
        
        '����Ϊ������Ŀ��,�������Ŀ��ͬ
        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngFirstRow, COL_����)
        .TextMatrix(lngRow, COL_������λ) = Nvl(rsMore!���㵥λ)
        
        'ִ��Ƶ��
        .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngFirstRow, COL_Ƶ��)
        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngFirstRow, COL_�����λ)
        .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngFirstRow, COL_ִ��ʱ��)
        .TextMatrix(lngRow, COL_ִ������) = Nvl(rsMore!ִ�п���, 0)
        
        'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
        If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
            .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsMore!���, rsMore!ID, 0, _
                Nvl(rsMore!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngFirstRow, COL_��������ID)), 1, 1)
        End If
        
        'ʱ��/����/ҽ��
        .TextMatrix(lngRow, COL_��ʼʱ��) = .TextMatrix(lngFirstRow, COL_��ʼʱ��)
        .Cell(flexcpData, lngRow, COL_��ʼʱ��) = .Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
        .TextMatrix(lngRow, COL_����ʱ��) = .TextMatrix(lngFirstRow, COL_����ʱ��)
        .Cell(flexcpData, lngRow, COL_����ʱ��) = .Cell(flexcpData, lngFirstRow, COL_����ʱ��)
        .TextMatrix(lngRow, COL_��������ID) = .TextMatrix(lngFirstRow, COL_��������ID)
        .TextMatrix(lngRow, COL_����ҽ��) = .TextMatrix(lngFirstRow, COL_����ҽ��)
        
        '��ʾ������־
        .TextMatrix(lngRow, COL_��־) = .TextMatrix(lngFirstRow, COL_��־)
        Call SetRow��־ͼ��(lngRow)
                
        If Not rsCurr Is Nothing Then
            .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsCurr!ҽ������)
        End If
        .Cell(flexcpData, lngRow, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", .TextMatrix(lngRow, COL_������ĿID))
        
        'ҽ������:����1,����2(�걾 �ɼ�����)
        .TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
        
        .Row = lngRow
    End With
    mblnRowChange = True
    AdviceSet������� = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceSet������Ŀ(rsInput As ADODB.Recordset, ByVal lngRow As Long, ByVal lng��ҩ;��ID As Long, ByVal lngGroupRow As Long, _
        ByVal strExtData As String, ByVal strժҪ As String, Optional ByVal str������λ As String, Optional ByVal bln��Ѫ As Boolean = True)
'���ܣ���������(����)���С�����ҩ�����(���)������(���)�����ģ���Ѫ��������������Ŀ��ȱʡҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'      lngRow=��ǰ������
'      lng��ҩ;��ID=ȱʡ��ҩ;��ID,��һ����ҩʱ�ĸ�ҩ;��ID
'      lngGroupRow=��һ����ҩ��һ���ҩ�в����µĳ�ҩ��ʱ,��Ӧһ����ҩ��һ���к�
'      strExtData=���:������鲿λ��Ϣ,����:���������������������Ϣ,�����޸�������
'      strժҪ=ҽ��ժҪ
'      str������λ=���븽���е�������λ
'      bln��Ѫ ��ǰ����Ѫҽ��Ϊ��Ѫҽ�����������ΪK��������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim rsMore As New ADODB.Recordset '������Ŀ��ϸ��Ϣ
    Dim strSQL As String, lngCopyRow As Long
    Dim lngTmp As Long, i As Long
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strҩ��IDs As String, sng���� As Single
    Dim lngִ�п���ID As Long, vCurDate As Date
    
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim blnǿ��ȱʡ As Boolean, strĬ��ҩ�� As String

    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
                
    With vsAdvice
        '��ʼ����ҽ��ȱʡ����
        .RowData(lngRow) = GetNextҽ��ID
        .TextMatrix(lngRow, COL_EDIT) = 1 '����
        .TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
        .TextMatrix(lngRow, COL_״̬) = 1 '�¿�
        
        '���:��������,��ǰ��ռ������ź�,�������������
        .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1)
        
        .TextMatrix(lngRow, COL_���) = rsInput!���ID
        .TextMatrix(lngRow, COL_����) = rsInput!���� '�����ƿ����Ǳ���
        .TextMatrix(lngRow, COL_������ĿID) = rsInput!������ĿID
        .TextMatrix(lngRow, COL_�շ�ϸĿID) = Nvl(rsInput!�շ�ϸĿID)
        .Cell(flexcpData, lngRow, COL_ҽ������) = strժҪ
        'ҩƷ�����ĵĹ����Ϣ
        If Not IsNull(rsInput!�շ�ϸĿID) Then
            If InStr(",5,6,", rsInput!���ID) > 0 Then
                strSQL = "Select Nvl(C.����,A.����) as ����,b.����ҩ��," & _
                    " B.����ϵ��,B.���ﵥλ,B.�����װ,B.����ɷ���� As �ɷ����,b.��ΣҩƷ" & _
                    " From �շ���ĿĿ¼ A,ҩƷ��� B,�շ���Ŀ���� C" & _
                    " Where A.ID=B.ҩƷID And A.ID=[1]" & _
                    " And A.ID=C.�շ�ϸĿID(+) And C.����(+)=1 And C.����(+)=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!�շ�ϸĿID), IIF(gbytҩƷ������ʾ = 0, 1, 3))
                .TextMatrix(lngRow, COL_����) = rsTmp!���� '������������ʽ�������
                .TextMatrix(lngRow, COL_����ϵ��) = rsTmp!����ϵ��
                .TextMatrix(lngRow, COL_���ﵥλ) = rsTmp!���ﵥλ
                .TextMatrix(lngRow, COL_�����װ) = rsTmp!�����װ
                .TextMatrix(lngRow, COL_�ɷ����) = Nvl(rsTmp!�ɷ����, 0)
                .TextMatrix(lngRow, COL_��ΣҩƷ) = Nvl(rsTmp!��ΣҩƷ, 0)
                .TextMatrix(lngRow, COL_����ҩ��) = rsTmp!����ҩ�� & ""
                If Val(.TextMatrix(lngRow, COL_��ΣҩƷ)) <> 0 Then
                    MsgBox "��ǰ�¿�����" & Decode(Val(.TextMatrix(lngRow, COL_��ΣҩƷ)), 1, "A", 2, "B", 3, "C", "") & "����ΣҩƷ�������ʹ�á�", vbInformation, Me.Caption
                End If
            ElseIf rsInput!���ID = "4" Then
                strSQL = "Select A.��������,B.����,B.���㵥λ From �������� A,�շ���ĿĿ¼ B Where A.����ID=B.ID And A.����ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!�շ�ϸĿID))
                .TextMatrix(lngRow, COL_����) = rsTmp!���� '������������ʽ�������
                .TextMatrix(lngRow, COL_����ϵ��) = 1
                .TextMatrix(lngRow, COL_�����װ) = 1
                .TextMatrix(lngRow, COL_���ﵥλ) = Nvl(rsTmp!���㵥λ) 'ɢװ��λ
                .TextMatrix(lngRow, COL_��������) = Nvl(rsTmp!��������, 0)
                .TextMatrix(lngRow, COL_��鷽��) = rsInput!���� & ""
            End If
        End If
        
        'ҩƷ����
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            strSQL = "Select �������,������,ҩƷ����,��������,����ְ��,�ٴ��Թ�ҩ,��ý From ҩƷ���� Where ҩ��ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!������ĿID))
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COL_�������) = Nvl(rsTmp!�������)
                .TextMatrix(lngRow, COL_�����ȼ�) = Val("" & rsTmp!������)
                .TextMatrix(lngRow, COL_ҩƷ����) = Nvl(rsTmp!ҩƷ����)
                .TextMatrix(lngRow, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(lngRow, COL_����ְ��) = Nvl(rsTmp!����ְ��)
                .TextMatrix(lngRow, COL_�ٴ��Թ�ҩ) = rsTmp!�ٴ��Թ�ҩ & ""
                .TextMatrix(lngRow, COL_�Ƿ���ý) = Val(rsTmp!��ý & "")
                
                If gblnKSSStrict And UserInfo.��ҩ���� < Val("" & rsTmp!������) Then
                    .TextMatrix(lngRow, COL_���״̬) = 1
                End If
            End If
        End If
        
        If rsInput!���ID & "" <> "K" Then
            If chk����.value = 1 Then
                If Val(.TextMatrix(lngRow, COL_���״̬)) = 1 Then .TextMatrix(lngRow, COL_���״̬) = ""
            Else
                If gblnKSSStrict And UserInfo.��ҩ���� < Val(.TextMatrix(lngRow, COL_�����ȼ�)) Then .TextMatrix(lngRow, COL_���״̬) = 1
            End If
        End If
        
        '��ȡ����������Ŀ��Ϣ
        '----------------------------------------------------------------------------
        strSQL = "Select A.*" & _
            " From �����÷����� A,������ĿĿ¼ B" & _
            " Where A.�÷�ID=B.ID And (Nvl(A.����,0)=0 Or B.������� IN(1,3))" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
            " And A.��ĿID=[1]"
        strSQL = "Select A.*,Nvl(B.����,0) as ����,B.�÷�ID," & _
            " B.Ƶ��,B.���˼���,B.С������,B.ҽ������,B.�Ƴ�" & _
            " From ������ĿĿ¼ A,(" & strSQL & ") B" & _
            " Where A.ID=B.��ĿID(+) And A.ID=[1]" & _
            " Order by ����"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!������ĿID))
                
        If IsNull(rsInput!�շ�ϸĿID) Then '������������ʽ��������
            .TextMatrix(lngRow, COL_����) = rsMore!����
        End If
                
        If rsInput!���ID = "4" Then
            .TextMatrix(lngRow, COL_������λ) = .TextMatrix(lngRow, COL_���ﵥλ) 'ɢװ��λ
        ElseIf InStr(",5,6,", rsInput!���ID) > 0 Or (Nvl(rsMore!ִ��Ƶ��, 0) = 0 And InStr(",1,2,", Nvl(rsMore!���㷽ʽ, 0)) > 0) Then
            .TextMatrix(lngRow, COL_������λ) = Nvl(rsMore!���㵥λ) 'ҩƷΪ������λ
        End If
        
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            '�С�����ҩ������������λ�������ﵥλ
            .TextMatrix(lngRow, COL_������λ) = .TextMatrix(lngRow, COL_���ﵥλ)
        ElseIf rsInput!���ID = "4" Then
            .TextMatrix(lngRow, COL_������λ) = .TextMatrix(lngRow, COL_���ﵥλ) 'ɢװ��λ
        Else
            '��������Ҫ��������(���㵥λ)
            '���Ϊһ���Ի�ƴ�����ȱʡ����Ϊ1
            If Nvl(rsMore!ִ��Ƶ��, 0) = 1 Or Nvl(rsMore!���㷽ʽ, 0) = 3 Then
                .TextMatrix(lngRow, COL_����) = 1
            End If
            .TextMatrix(lngRow, COL_������λ) = Nvl(rsMore!���㵥λ)
        End If
        
        '����ҩ��ȱʡ��ҩĿ��
        If Val(.TextMatrix(lngRow, COL_�����ȼ�)) > 0 Then .TextMatrix(lngRow, COL_��ҩĿ��) = mstrPurMed
        
        .TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsMore!���㷽ʽ, 0)
        .TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsMore!ִ��Ƶ��, 0)
        .TextMatrix(lngRow, COL_��������) = Nvl(rsMore!��������)
        .TextMatrix(lngRow, COL_����Ӧ��) = Nvl(rsMore!����Ӧ��)
        .TextMatrix(lngRow, COL_ִ�з���) = Nvl(rsMore!ִ�з���, 0)
        If InStr(",5,6,7,", rsInput!���ID) = 0 Then
            .TextMatrix(lngRow, COL_��������) = Nvl(rsMore!¼������)
        End If
        
        '�걾��λ
        If InStr(",4,5,6,", rsInput!���ID) > 0 Then
            .TextMatrix(lngRow, COL_�걾��λ) = rsInput!���� '��¼ҩƷ����������ʱѡ�������
        ElseIf rsInput!���ID = "F" Or rsInput!���ID = "K" Then
            .TextMatrix(lngRow, COL_����ʱ��) = "" '��¼����/��Ѫʱ��
        ElseIf rsInput!���ID <> "D" Then
            .TextMatrix(lngRow, COL_�걾��λ) = Nvl(rsMore!�걾��λ)
        End If
        
        '�Ƽ�����
        .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsMore!�Ƽ�����, 0)
    
        'ִ������:������Ŀʱ������Ŀ����,ҩƷ������=4-ָ������,һ����ҩ����ͬ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_ִ������) = .TextMatrix(lngGroupRow, COL_ִ������)
            Else
                .TextMatrix(lngRow, COL_ִ������) = IIF(Val(.TextMatrix(lngRow, COL_�ٴ��Թ�ҩ)) = 1, 5, 4) '�Ա�ҩ��ΪԺ��ִ��
            End If
        ElseIf rsInput!���ID = "4" Then
            .TextMatrix(lngRow, COL_ִ������) = 4
        Else
            .TextMatrix(lngRow, COL_ִ������) = Nvl(rsMore!ִ�п���, 0)
        End If
            
        '����ҽ���Ϳ���
        If lngGroupRow = 0 Then
            .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
            .TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
        Else
            .TextMatrix(lngRow, COL_����ҽ��) = .TextMatrix(lngGroupRow, COL_����ҽ��)
            .TextMatrix(lngRow, COL_��������ID) = .TextMatrix(lngGroupRow, COL_��������ID)
        End If
    
        'ִ�п���:ҩƷȱʡ����һ����ͬ,һ����ҩ����ͬ
        lngִ�п���ID = 0
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                .TextMatrix(lngRow, COL_ִ�п���ID) = 0
            Else

                strҩ��IDs = Get����ҩ��IDs(rsInput!���ID, rsInput!������ĿID, Nvl(rsInput!�շ�ϸĿID, 0), mlng���˿���id, 1)
                If lngGroupRow <> 0 Then
                    If InStr("," & strҩ��IDs & ",", "," & .TextMatrix(lngGroupRow, COL_ִ�п���ID) & ",") > 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngGroupRow, COL_ִ�п���ID)
                    End If
                ElseIf lngCopyRow <> -1 Then
                    blnǿ��ȱʡ = Val(zlDatabase.GetPara("����ҽ���´�ǿ��ȱʡҩ��", glngSys, p����ҽ���´�, 1)) = 1
                    strĬ��ҩ�� = zlDatabase.GetPara("����ȱʡ" & IIF(Val(rsInput!���ID) = 5, "��", "��") & "ҩ��", glngSys, p����ҽ���´�, mlng���˿���id)
                    If blnǿ��ȱʡ And InStr(strҩ��IDs, strĬ��ҩ��) > 0 And strĬ��ҩ�� <> "" Then
                        lngִ�п���ID = 0
                    Else
                        If rsInput!���ID = .TextMatrix(lngCopyRow, COL_���) Then
                            lngִ�п���ID = Val(.TextMatrix(lngCopyRow, COL_ִ�п���ID))
                        End If
                    End If
                End If
            End If
        End If

        If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
            If rsInput!���ID = "Z" And InStr(",1,2,", Nvl(rsMore!��������, 0)) > 0 Then
                '���ۻ�סԺҽ����ȱʡ
                If Nvl(rsMore!��������, 0) = 1 Then
                    '����:���������סԺ���ٴ�����
                    Call Get�ٴ�����(3, , lngTmp, , True, False, True)
                ElseIf Nvl(rsMore!��������, 0) = 2 Then
                    'סԺ:סԺ�ٴ�����
                    Call Get�ٴ�����(2, , lngTmp, , True, False, True)
                End If
                .TextMatrix(lngRow, COL_ִ�п���ID) = lngTmp
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                'ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                '֮ǰ�������������ID
                If rsInput!���ID = "4" Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get�շ�ִ�п���ID(mlng����ID, 0, _
                        rsInput!���ID, Nvl(rsInput!�շ�ϸĿID, 0), 4, mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, , 1)
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsInput!���ID, rsInput!������ĿID, _
                        Nvl(rsInput!�շ�ϸĿID, 0), Nvl(rsMore!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1, InStr(",5,6,", rsInput!���ID) > 0, lngִ�п���ID)
                End If
            End If
        End If
        
        'ҩƷ���
        If (InStr(",5,6,", rsInput!���ID) > 0 Or rsInput!���ID = "4" And Val(.TextMatrix(lngRow, COL_��������)) = 1) And Nvl(rsInput!�շ�ϸĿID, 0) <> 0 Then
            Call GetDrugStock(lngRow)
        End If
        
        'ִ��Ƶ��:��ѡƵ��,һ���Ի������
        
        'ȱʡ����һ��������ͬ
        If lngCopyRow <> -1 Then
            If GetƵ�ʷ�Χ(lngRow) = GetƵ�ʷ�Χ(lngCopyRow) Then
                If Val(.TextMatrix(lngCopyRow, COL_EDIT)) = 1 And .TextMatrix(lngCopyRow, COL_Ƶ��) <> "" _
                    And Not (.TextMatrix(lngRow, COL_���) = "7" And Not RowIn�䷽��(lngCopyRow)) _
                    And Not (.TextMatrix(lngRow, COL_���) <> "7" And RowIn�䷽��(lngCopyRow)) _
                    And CheckƵ�ʿ���(Nvl(rsInput!������ĿID, 0), GetƵ�ʷ�Χ(lngRow), .TextMatrix(lngCopyRow, COL_Ƶ��)) Then
                    .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��)
                    .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngCopyRow, COL_Ƶ�ʴ���)
                    .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngCopyRow, COL_Ƶ�ʼ��)
                    .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngCopyRow, COL_�����λ)
                End If
            End If
        End If
        '��ȡȱʡƵ��
        If .TextMatrix(lngRow, COL_Ƶ��) = "" Then
            Call GetȱʡƵ��(Nvl(rsInput!������ĿID, 0), GetƵ�ʷ�Χ(lngRow), strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
            .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
            .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
            .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
            .TextMatrix(lngRow, COL_�����λ) = str�����λ
        End If
        
        
        '�У�����ҩ��һЩȱʡ��Ϣ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            'ִ��Ƶ��
            If lngGroupRow <> 0 Then
                'һ����ҩ����ͬ
                .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngGroupRow, COL_Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngGroupRow, COL_Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngGroupRow, COL_Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngGroupRow, COL_�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngGroupRow, COL_ִ��ʱ��)
                
                If Val(.TextMatrix(lngRow, COL_�����ȼ�)) > 0 Then
                    .TextMatrix(lngRow, COL_��ҩĿ��) = .TextMatrix(lngGroupRow, COL_��ҩĿ��)
                    .TextMatrix(lngRow, COL_��ҩ����) = .TextMatrix(lngGroupRow, COL_��ҩ����)
                End If
            End If
            
            'ȷ��������ҩ������
            '1.����Ϊһ��Ƶ����������
            '2-���Ƴ���Ϊ�Ƴ�����(Ӧ����һ��Ƶ����������)
            sng���� = msng����
            If mbln���� Then
                If .TextMatrix(lngRow, COL_�����λ) = "��" Then
                    If 7 > sng���� Then sng���� = 7
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "��" Then
                    If Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) > sng���� Then
                        sng���� = Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��))
                    End If
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "Сʱ" Then
                    If Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) \ 24 > sng���� Then
                        sng���� = Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) \ 24
                    End If
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "����" Then
                    If sng���� = 0 Then sng���� = 1
                End If
                If sng���� = 0 Then sng���� = 1
            End If
            
            rsMore.Filter = "����>0" 'ȡ��һ�ָ�ҩ;����Ϊȱʡ����
            If Not rsMore.EOF Then
                '����һ����ҩʱ,���õ�ȱʡ�÷�Ƶ������
                If lngGroupRow = 0 Then
                    If Not IsNull(rsMore!�÷�ID) Then lng��ҩ;��ID = rsMore!�÷�ID
                    If Not IsNull(rsMore!Ƶ��) Then
                        Call GetƵ����Ϣ_����(rsMore!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                        .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                        .TextMatrix(lngRow, COL_�����λ) = str�����λ
                    End If
                End If
                
                'ҽ������
                .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsMore!ҽ������) 'һ��Ϊ��ҩ;����˵��
                
                'ҩƷ����
                If mint���� > 12 Then
                    If Nvl(rsMore!���˼���, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_����) = FormatEx(rsMore!���˼���, 5)
                    End If
                Else
                    If Nvl(rsMore!С������, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_����) = FormatEx(rsMore!С������, 5)
                    ElseIf Nvl(rsMore!���˼���, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_����) = FormatEx(rsMore!���˼��� * (mint���� + 2) * 5 / 100, 5)
                    End If
                End If
                If Val(.TextMatrix(lngRow, COL_����)) = 0 Then .TextMatrix(lngRow, COL_����) = ""
                
                'ҩƷ��������:�����װ
                If Nvl(rsMore!�Ƴ�, 1) > sng���� Then sng���� = Nvl(rsMore!�Ƴ�, 1)
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" And Val(.TextMatrix(lngRow, COL_����)) <> 0 _
                    And Val(.TextMatrix(lngRow, COL_����ϵ��)) <> 0 And Val(.TextMatrix(lngRow, COL_�����װ)) <> 0 Then
                    '�����Ƴ����Ϊ��������ҩ������
                    .TextMatrix(lngRow, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                            Val(.TextMatrix(lngRow, COL_����)), sng����, _
                            Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), _
                            Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), _
                            .TextMatrix(lngRow, COL_�����λ), _
                            .TextMatrix(lngRow, COL_ִ��ʱ��), _
                            Val(.TextMatrix(lngRow, COL_����ϵ��)), _
                            Val(.TextMatrix(lngRow, COL_�����װ)), _
                            Val(.TextMatrix(lngRow, COL_�ɷ����))), 5)
                    If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") = 0 Then
                        .TextMatrix(lngRow, COL_����) = IntEx(Val(.TextMatrix(lngRow, COL_����)))
                    ElseIf Val(.TextMatrix(lngRow, COL_�ɷ����)) <> 0 Then
                        .TextMatrix(lngRow, COL_����) = IntEx(Val(.TextMatrix(lngRow, COL_����)))
                    End If
                End If
            End If
            
            '��¼ȱʡ����
            If mbln���� Then .TextMatrix(lngRow, COL_����) = IIF(sng���� = 0, "", sng����)
            '���������������������������ܱ������Զ��������ֶ�¼�룩��ֵ�����������÷��������򲻻�ִ�г������Ĵ��루�ؼ�Validate�¼���
            Call Set��ҩ�����Ƿ���(lngRow)
            If Val(.TextMatrix(lngRow, COL_��������)) <> 0 And Val(.TextMatrix(lngRow, COL_����)) <> 0 Then
                If Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_�����װ)) * Val(.TextMatrix(lngRow, COL_����ϵ��)) > Val(.TextMatrix(lngRow, COL_��������)) Then
                    .TextMatrix(lngRow, COL_�Ƿ���) = "1"
                End If
            End If
        End If
        
        If rsMore.Filter <> 0 Then rsMore.Filter = 0
        
        'ִ��ʱ��:"��ѡƵ��"��ҩƷ
        If Nvl(rsMore!ִ��Ƶ��, 0) = 0 Or InStr(",5,6,", rsInput!���ID) > 0 Then
            If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then
                If lngCopyRow <> -1 Then '����һ����ͬ
                    If .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��) Then
                        .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngCopyRow, COL_ִ��ʱ��)
                    End If
                End If
                If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then 'ȱʡʱ�䷽��
                    .TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(1, .TextMatrix(lngRow, COL_Ƶ��), lng��ҩ;��ID)
                End If
            End If
        End If
        
        '����(����Ŀ�޹�)
        '---------------------------------------------------------------------
        If lngGroupRow = 0 Then
            vCurDate = zlDatabase.Currentdate
            If IsDate(txt��ʼʱ��.Text) Then
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
            
                '����/��Ѫʱ��ȱʡΪ��ʼʱ��
                If rsInput!���ID = "F" Or rsInput!���ID = "K" Then
                    .TextMatrix(lngRow, COL_����ʱ��) = txt��ʼʱ��.Text
                End If
            End If
            
            .TextMatrix(lngRow, COL_����ʱ��) = Format(vCurDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, COL_����ʱ��) = Format(vCurDate, "yyyy-MM-dd HH:mm")
    
            .TextMatrix(lngRow, COL_��־) = chk����.value
        Else
            .TextMatrix(lngRow, COL_��ʼʱ��) = .TextMatrix(lngGroupRow, COL_��ʼʱ��)
            .Cell(flexcpData, lngRow, COL_��ʼʱ��) = .Cell(flexcpData, lngGroupRow, COL_��ʼʱ��)
            
            .TextMatrix(lngRow, COL_����ʱ��) = .TextMatrix(lngGroupRow, COL_����ʱ��)
            .Cell(flexcpData, lngRow, COL_����ʱ��) = .Cell(flexcpData, lngGroupRow, COL_����ʱ��)
            
            .TextMatrix(lngRow, COL_��־) = .TextMatrix(lngGroupRow, COL_��־)
        End If
                        
        
        '�����д������֮��������,�����ҽ������
        '-------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            '����һ����ҩ;����Ŀ,���������
            If lng��ҩ;��ID <> 0 Then
                .TextMatrix(lngRow, COL_�÷�) = Get��Ŀ����(lng��ҩ;��ID)
            End If
            If lngGroupRow <> 0 Then
                'һ����ҩ�Ĺ�����ͬ�ĸ�ҩ;����
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_���ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    .TextMatrix(lngRow, COL_���ID) = .TextMatrix(lngGroupRow, COL_���ID)
                Else
                    '��������ǽ�Ϊ��ʹ��һ����ҩ����ͬ����
                    .TextMatrix(lngRow, COL_���ID) = AdviceSet��ҩ;��(lngRow, lng��ҩ;��ID)
                End If
            Else '���������ĳ�ҩ���������ĸ�ҩ;����
                .TextMatrix(lngRow, COL_���ID) = AdviceSet��ҩ;��(lngRow, lng��ҩ;��ID)
            End If
            
            '���龫����ɫ��ʶ
            If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(lngRow, COL_�������)) > 0 _
                And .TextMatrix(lngRow, COL_�������) <> "" Then
                .Cell(flexcpFontBold, lngRow, col_ҽ������) = True
            End If
        ElseIf rsInput!���ID = "D" And strExtData <> "" Then
            '������ϲ�λ��
            Call AdviceSet������(lngRow, strExtData)
        ElseIf rsInput!���ID = "F" And strExtData <> "" Then
            '�����ĸ���������������Ŀ��
            Call AdviceSet�������(lngRow, strExtData)
            vsAdvice.Cell(flexcpData, lngRow, COL_�걾��λ) = str������λ
        ElseIf rsInput!���ID = "K" Then
            If bln��Ѫ Then
                .TextMatrix(lngRow, COL_��鷽��) = ""
            Else
                .TextMatrix(lngRow, COL_��鷽��) = 1
            End If
            '��Ѫ��;����
            If lng��ҩ;��ID <> 0 Then
                If gblnѪ��ϵͳ = True Then
                    strSQL = "Select a.����,a.��������,a.ִ�з��� From ������ĿĿ¼ A where a.id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ҩ;��ID)
                    .TextMatrix(lngRow, COL_�÷�) = rsTmp!���� & ""
                    If Val(rsTmp!�������� & "") = 8 And Val(rsTmp!ִ�з��� & "") = 1 Then '����Ǳ༭���������뵥ʱ��Ҫ����һ��
                        .TextMatrix(lngRow, COL_��鷽��) = 1
                    Else
                        .TextMatrix(lngRow, COL_��鷽��) = ""
                    End If
                Else
                    .TextMatrix(lngRow, COL_�÷�) = Get��Ŀ����(lng��ҩ;��ID)
                End If
                Call AdviceSet��Ѫ;��(lngRow, lng��ҩ;��ID)
            End If
        End If
        
        '������־
        If lngGroupRow <> 0 And .TextMatrix(lngRow, COL_���״̬) <> "" Then
            Call SetRow��־ͼ��(lngRow, 0)
        Else
            Call SetRow��־ͼ��(lngRow, 1)
        End If
        
        .TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AdviceSet������(ByVal lngRow As Long, ByVal strExData As String)
'���ܣ���������ָ����������Ŀ�Ĳ�λ������,�����������������Ŀ���޸Ĳ�λ����
'������lngRow=��ǰ������
'      strExData=������鲿λ��������Ϣ,��ʽΪ:"��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
    Dim arrItems As Variant, arrMethod As Variant
    Dim intִ�б�� As Integer, str��鲿λ As String
    Dim i As Integer, j As Integer, k As Integer
    
    'ɾ�����еļ�鲿λ������
    Call Delete���������Ѫ(lngRow)
    
    '���¼��벿λ������
    If strExData <> "" Then
        'ִ�б��
        If UBound(Split(strExData, vbTab)) >= 1 Then
            intִ�б�� = Val(Split(strExData, vbTab)(1))
        End If
        vsAdvice.TextMatrix(lngRow, COL_ִ�б��) = intִ�б�� '������ͳһ�������е�ִ�б��
        
        arrItems = Split(Split(strExData, vbTab)(0), "|")
        For i = 0 To UBound(arrItems)
            str��鲿λ = Split(arrItems(i), ";")(0)
            arrMethod = Split(Split(arrItems(i), ";")(1), ",")
            For j = 0 To UBound(arrMethod)
                k = k + 1
                With vsAdvice
                    .AddItem "", lngRow + k
                    .RowHidden(lngRow + k) = True
                    
                    .RowData(lngRow + k) = GetNextҽ��ID
                    .TextMatrix(lngRow + k, COL_���ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + k, COL_EDIT) = 1 '����
                    
                    .TextMatrix(lngRow + k, COL_Ӥ��) = cboӤ��.ListIndex
                    .TextMatrix(lngRow + k, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + k
                    .TextMatrix(lngRow + k, COL_״̬) = 1 '�¿�
                    
                    .TextMatrix(lngRow + k, COL_���) = .TextMatrix(lngRow, COL_���)
                    .TextMatrix(lngRow + k, COL_������ĿID) = .TextMatrix(lngRow, COL_������ĿID) 'Ϊͬһ�������Ŀ
                    
                    .TextMatrix(lngRow + k, COL_���㷽ʽ) = .TextMatrix(lngRow, COL_���㷽ʽ)
                    .TextMatrix(lngRow + k, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������)
                    .TextMatrix(lngRow + k, COL_��������) = .TextMatrix(lngRow, COL_��������)
                    .TextMatrix(lngRow + k, COL_����Ӧ��) = .TextMatrix(lngRow, COL_����Ӧ��)
                    .TextMatrix(lngRow + k, COL_ִ�з���) = .TextMatrix(lngRow, COL_ִ�з���)
                    .TextMatrix(lngRow + k, COL_��������) = .TextMatrix(lngRow, COL_��������)
                    
                    .TextMatrix(lngRow + k, col_ҽ������) = .TextMatrix(lngRow, COL_����) '��¼Ϊ�����Ŀ����
                    .TextMatrix(lngRow + k, COL_�걾��λ) = str��鲿λ
                    .TextMatrix(lngRow + k, COL_��鷽��) = arrMethod(j)
                    .TextMatrix(lngRow + k, COL_ִ�б��) = intִ�б��
                    
                    .TextMatrix(lngRow + k, COL_�Ƽ�����) = .TextMatrix(lngRow, COL_�Ƽ�����)
                    
                    .TextMatrix(lngRow + k, COL_����) = .TextMatrix(lngRow, COL_����)
                    .TextMatrix(lngRow + k, COL_����) = .TextMatrix(lngRow, COL_����)
                    
                    .TextMatrix(lngRow + k, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                    .TextMatrix(lngRow + k, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                    .TextMatrix(lngRow + k, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                    .TextMatrix(lngRow + k, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                    .TextMatrix(lngRow + k, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                    
                    .TextMatrix(lngRow + k, COL_ִ������) = .TextMatrix(lngRow, COL_ִ������)
                    .TextMatrix(lngRow + k, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                    
                    .TextMatrix(lngRow + k, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                    .Cell(flexcpData, lngRow + k, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                    
                    .TextMatrix(lngRow + k, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                    .TextMatrix(lngRow + k, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                    
                    .TextMatrix(lngRow + k, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                    .Cell(flexcpData, lngRow + k, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                    
                    .TextMatrix(lngRow + k, COL_��־) = .TextMatrix(lngRow, COL_��־)
                End With
            Next
        Next
                
        '��������ҽ�������
        Call AdviceSetҽ�����(lngRow + k + 1, k)
    End If
End Sub

Private Sub AdviceSet�������(ByVal lngRow As Long, ByVal strDataIDs As String)
'���ܣ���������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������lngRow=��ǰ������
'      strDataIDs=��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    'ɾ�����еĸ���������������Ŀ��
    Call Delete���������Ѫ(lngRow)
    
    '���¼��븽�������м�������Ŀ��
    strDataIDs = Trim(Replace(strDataIDs, ";", ","))
    If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
    If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    
    If strDataIDs <> "" Then
        Set rsTmp = Get������Ŀ��¼(0, strDataIDs)
        If Not rsTmp.EOF Then
            arrIDs = Split(strDataIDs, ",")
            For i = 0 To UBound(arrIDs) '���û�������Ŀ˳��
                rsTmp.Filter = "ID=" & CStr(arrIDs(i)) '������EOF
                
                With vsAdvice
                    .AddItem "", lngRow + i + 1
                    .RowHidden(lngRow + i + 1) = True
                    
                    .RowData(lngRow + i + 1) = GetNextҽ��ID
                    .TextMatrix(lngRow + i + 1, COL_���ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + i + 1, COL_EDIT) = 1 '����
                    
                    .TextMatrix(lngRow + i + 1, COL_Ӥ��) = cboӤ��.ListIndex
                    .TextMatrix(lngRow + i + 1, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + i + 1
                    .TextMatrix(lngRow + i + 1, COL_״̬) = 1 '�¿�
                    
                    .TextMatrix(lngRow + i + 1, COL_���) = rsTmp!���
                    .TextMatrix(lngRow + i + 1, COL_������ĿID) = rsTmp!ID
                    .Cell(flexcpData, lngRow + i + 1, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", .TextMatrix(lngRow + i + 1, COL_������ĿID))
                    
                    .TextMatrix(lngRow + i + 1, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
                    .TextMatrix(lngRow + i + 1, COL_��������) = Nvl(rsTmp!��������)
                    .TextMatrix(lngRow + i + 1, COL_����Ӧ��) = Nvl(rsTmp!����Ӧ��)
                    .TextMatrix(lngRow + i + 1, COL_ִ�з���) = Nvl(rsTmp!ִ�з���, 0)
                    .TextMatrix(lngRow + i + 1, COL_��������) = Nvl(rsTmp!¼������)
                    
                    .TextMatrix(lngRow + i + 1, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��) '����/��Ѫʱ��
                    .TextMatrix(lngRow + i + 1, col_ҽ������) = rsTmp!����
                    
                    .TextMatrix(lngRow + i + 1, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
                    
                    .TextMatrix(lngRow + i + 1, COL_����) = .TextMatrix(lngRow, COL_����)
                    .TextMatrix(lngRow + i + 1, COL_����) = .TextMatrix(lngRow, COL_����)
    
                    .TextMatrix(lngRow + i + 1, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                    .TextMatrix(lngRow + i + 1, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                    
                    'ִ������:������Ŀ��������
                    .TextMatrix(lngRow + i + 1, COL_ִ������) = Nvl(rsTmp!ִ�п���, 0)
                    
                    '������Ժ��ִ����ִ�п���,����������ִ�п���
                    '���򲻹���ִ�п�������,һ���������Ӧ����ͬ
                    If InStr(",0,5,", Nvl(rsTmp!ִ�п���, 0)) > 0 Then
                        .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = 0
                    Else
                        If rsTmp!��� = "G" Then
                            .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!ID, 0, _
                                Nvl(rsTmp!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                        Else
                            .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                        End If
                    End If
                    
                    .TextMatrix(lngRow + i + 1, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                    .Cell(flexcpData, lngRow + i + 1, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                    
                    .TextMatrix(lngRow + i + 1, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                    .TextMatrix(lngRow + i + 1, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                    
                    .TextMatrix(lngRow + i + 1, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                    .Cell(flexcpData, lngRow + i + 1, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                    
                    .TextMatrix(lngRow + i + 1, COL_��־) = .TextMatrix(lngRow, COL_��־)
                End With
            Next
                
            '�������
            Call AdviceSetҽ�����(lngRow + UBound(arrIDs) + 2, UBound(arrIDs) + 1)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSet��ҩ;��(ByVal lngRow As Long, ByVal lng��ҩ;��ID As Long, _
    Optional ByVal strִ������ As String, Optional ByVal lng��ҩִ��ID As Long, Optional ByVal str���� As String) As Long
'���ܣ�Ϊ¼����У�����ҩ���ö�Ӧ�ĸ�ҩ;����(�������޸�)
'������lngRow=Ҫ�����ҩ;����ҩƷ��
'      lng��ҩ;��ID=��ҩ;��ID
'      strִ������=�޸ĸ�ҩ;��ʱ,��ǰ�������õ�ִ������
'      lng��ҩִ��ID=�޸ĸ�ҩ;��ʱ,��ǰ�������õ�ִ�п���
'      str����=�޸ĸ�ҩ;��ʱ,��ǰ�������õĵ���
'���أ������õĸ�ҩ;���е�ҽ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get������Ŀ��¼(lng��ҩ;��ID)
    If rsTmp.EOF Then lng��ҩ;��ID = 0 'û�����ݣ��������Ա��ֹ�ϵ
        
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then 'δ����"���ID"ʱ
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        Else
            '�޸�ҽ��������ʱ�������ø�ҩ;������(���Ǹ���������Ŀ)
            blnNew = False
            lngNewRow = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
        End If
        
        '��Ч���ݣ�����,�շ�ϸĿID,����ϵ��,���ﵥλ,�����װ,�걾��λ,ҽ������,����,����,�÷�
        If blnNew Then
            .RowData(lngNewRow) = GetNextҽ��ID
            .TextMatrix(lngNewRow, COL_EDIT) = 1 '����
            .TextMatrix(lngNewRow, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + 1
        Else
            'ҽ��ID(RowData),���:���ֲ���
            If InStr(",0,3,", .TextMatrix(lngNewRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngNewRow, COL_EDIT) = 2 '��־Ϊ�����޸�
                .TextMatrix(lngNewRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_Ӥ��) = cboӤ��.ListIndex
        .TextMatrix(lngNewRow, COL_״̬) = 1 '�¿�
        
        .TextMatrix(lngNewRow, COL_���) = "E" '��ҩ;����������
        .TextMatrix(lngNewRow, COL_������ĿID) = lng��ҩ;��ID
        '���û��ȷ����ҩ;������ʱ�����õ�����
        If Not rsTmp.EOF Then
            .TextMatrix(lngNewRow, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
            .TextMatrix(lngNewRow, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
            .TextMatrix(lngNewRow, COL_��������) = Nvl(rsTmp!��������)
            .TextMatrix(lngNewRow, COL_����Ӧ��) = Nvl(rsTmp!����Ӧ��)
            .TextMatrix(lngNewRow, COL_ִ�з���) = Nvl(rsTmp!ִ�з���, 0)
            .TextMatrix(lngNewRow, col_ҽ������) = rsTmp!����
            
            .TextMatrix(lngNewRow, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
            
            '����
            If str���� <> "" Then
                .TextMatrix(lngNewRow, COL_ҽ������) = str����
            End If
            .Cell(flexcpData, lngNewRow, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", .TextMatrix(lngNewRow, COL_������ĿID))
            
            'ִ������:ȱʡ������Ŀ����,�޸�ʱ���ݵ�ǰ��������
            If strִ������ = "" Then
                .TextMatrix(lngNewRow, COL_ִ������) = Nvl(rsTmp!ִ�п���, 0)
            Else
                .TextMatrix(lngNewRow, COL_ִ������) = Decode(strִ������, "��Ժ��ҩ", 5, Nvl(rsTmp!ִ�п���, 0))
            End If
            
            '��ҩ;�����δ����ִ�п���,��ȱʡΪ�������ڲ���(����Ҫ��Ϊ�������ڿ���!!)
            If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_ִ������))) = 0 Then
                If lng��ҩִ��ID <> 0 Then
                    .TextMatrix(lngNewRow, COL_ִ�п���ID) = lng��ҩִ��ID
                Else
                    .TextMatrix(lngNewRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, "E", lng��ҩ;��ID, 0, _
                        Nvl(rsTmp!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                End If
            Else
                .TextMatrix(lngNewRow, COL_ִ�п���ID) = 0
            End If
        End If
        
        '��ҩ;��������ҩƷ��ͬ
        .TextMatrix(lngNewRow, COL_����) = .TextMatrix(lngRow, COL_����)
        
        .TextMatrix(lngNewRow, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
        .TextMatrix(lngNewRow, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngNewRow, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngNewRow, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
        .TextMatrix(lngNewRow, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
        
        .TextMatrix(lngNewRow, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
        .Cell(flexcpData, lngNewRow, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
        
        .TextMatrix(lngNewRow, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
        .TextMatrix(lngNewRow, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
        
        .TextMatrix(lngNewRow, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
        .Cell(flexcpData, lngNewRow, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
        
        .TextMatrix(lngNewRow, COL_��־) = .TextMatrix(lngRow, COL_��־)
        .TextMatrix(lngNewRow, COL_���״̬) = .TextMatrix(lngRow, COL_���״̬)
            
        '����������
        If blnNew Then Call AdviceSetҽ�����(lngNewRow + 1, 1)
        
        AdviceSet��ҩ;�� = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet��Ѫ;��(ByVal lngRow As Long, ByVal lng��Ѫ;��ID As Long, Optional ByVal lng��Ѫִ��ID As Long) As Long
'���ܣ�Ϊ¼����У�����ҩ���ö�Ӧ�ĸ�ҩ;����(�������޸�)
'������lngRow=Ҫ������Ѫ;������Ѫҽ����
'      lng��Ѫ;��ID=��Ѫ;��ID
'      lng��Ѫִ��ID=�޸���Ѫ;��ʱ,��ǰ�������õ�ִ�п���
'���أ������õ���Ѫ;���е�ҽ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get������Ŀ��¼(lng��Ѫ;��ID)
    
    With vsAdvice
        lngNewRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
        If lngNewRow = -1 Then '��δ������Ѫ;��ʱ
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        End If
        
        '��Ч���ݣ�����,�շ�ϸĿID,����ϵ��,���ﵥλ,�����װ,�걾��λ,ҽ������,����,����,�÷�
        If blnNew Then
            .RowData(lngNewRow) = GetNextҽ��ID
            .TextMatrix(lngNewRow, COL_���ID) = .RowData(lngRow)
            .TextMatrix(lngNewRow, COL_EDIT) = 1 '����
            .TextMatrix(lngNewRow, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + 1
        Else
            'ҽ��ID(RowData),���:���ֲ���
            If InStr(",0,3,", .TextMatrix(lngNewRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngNewRow, COL_EDIT) = 2 '��־Ϊ�����޸�
                .TextMatrix(lngNewRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_Ӥ��) = cboӤ��.ListIndex
        .TextMatrix(lngNewRow, COL_״̬) = 1 '�¿�
        
        .TextMatrix(lngNewRow, COL_���) = "E" '��Ѫ;����������
        .TextMatrix(lngNewRow, COL_������ĿID) = lng��Ѫ;��ID
        .Cell(flexcpData, lngNewRow, COL_ҽ������) = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", .TextMatrix(lngNewRow, COL_������ĿID))
        
        .TextMatrix(lngNewRow, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
        .TextMatrix(lngNewRow, COL_��������) = Nvl(rsTmp!��������)
        .TextMatrix(lngNewRow, COL_����Ӧ��) = Nvl(rsTmp!����Ӧ��)
        .TextMatrix(lngNewRow, COL_ִ�з���) = Nvl(rsTmp!ִ�з���, 0)
        .TextMatrix(lngNewRow, col_ҽ������) = rsTmp!����
        
        .TextMatrix(lngNewRow, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
        .TextMatrix(lngNewRow, COL_ִ������) = Nvl(rsTmp!ִ�п���, 0)
        
        If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_ִ������))) = 0 Then
            If lng��Ѫִ��ID <> 0 Then
                .TextMatrix(lngNewRow, COL_ִ�п���ID) = lng��Ѫִ��ID
            Else
                .TextMatrix(lngNewRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, "E", lng��Ѫ;��ID, 0, _
                    Nvl(rsTmp!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
            End If
        Else
            .TextMatrix(lngNewRow, COL_ִ�п���ID) = 0
        End If
        
        .TextMatrix(lngNewRow, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������) '��ҩƷ��Ϊ׼
        .TextMatrix(lngNewRow, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
        .TextMatrix(lngNewRow, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngNewRow, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngNewRow, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
        .TextMatrix(lngNewRow, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
        .TextMatrix(lngNewRow, COL_���״̬) = .TextMatrix(lngRow, COL_���״̬)
        
        .TextMatrix(lngNewRow, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
        .Cell(flexcpData, lngNewRow, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
        
        .TextMatrix(lngNewRow, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
        .TextMatrix(lngNewRow, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
        
        .TextMatrix(lngNewRow, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
        .Cell(flexcpData, lngNewRow, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
        
        .TextMatrix(lngNewRow, COL_��־) = .TextMatrix(lngRow, COL_��־)
            
        '����������
        If blnNew Then Call AdviceSetҽ�����(lngNewRow + 1, 1)
        
        AdviceSet��Ѫ;�� = .RowData(lngNewRow)
        
        '������Ѫ;����������Ѫҽ�������״̬
        strTmp = GetBloodState(IIF(Val(.TextMatrix(lngNewRow, COL_��־)) = 1, 1, 0), Val(.TextMatrix(lngNewRow, COL_ִ�з���)))
        
        .TextMatrix(lngNewRow, COL_���״̬) = strTmp
        .TextMatrix(lngRow, COL_���״̬) = strTmp
        Call SetRow��־ͼ��(lngRow, 2)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceChange()
'���ܣ����ݵ�ǰҽ����Ƭ�е����ݣ����µ�ǰҽ������
'˵��������ListIndex=-1����Ӧҽ�����������ݵģ�����ԭ���ݲ�����
    Dim lngRow As Long, lngBeginRow As Long, lngEndRow As Long
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim blnCurDo As Boolean, blnOtherDo As Boolean
    Dim lngTmp As Long, strTmp As String, blnTmp As Boolean
    Dim strCurDate As String, lng��������ID As Long
    Dim blnReInRow As Boolean, i As Long, j As Long
    Dim lngִ�п���ID As Long, lngBegin As Long, lngEnd As Long
    Dim blnReSet����˵�� As Boolean
    Dim dbl���� As Double
    
    With vsAdvice
        lngRow = .Row
        
        If .RowData(lngRow) = 0 Then Call ClearItemTag: Exit Sub '����༭��־
        
        If RowIn�䷽��(lngRow) Then
            '��ҩ�䷽
            lngBeginRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            For i = lngBeginRow To lngRow
                '�޸Ĵ����䷽������������(�����巨���÷�)
                If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" Then
                    .TextMatrix(i, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_��ʼʱ��) = txt��ʼʱ��.Text
                    blnCurDo = True
                End If
                If chk����.Visible And chk����.Tag <> "" Then
                    .TextMatrix(i, COL_��־) = chk����.value
                    If i = lngRow Then '�÷�����ʾ������־
                         Call SetRow��־ͼ��(i, 0)
                    End If
                    blnCurDo = True
                End If
                If txt����.Enabled And txt����.Tag <> "" Then
                    .TextMatrix(i, COL_����) = FormatEx(IIF(Val(txt����.Text) = 0, "", Val(txt����.Text)), 5)
                    
                    '�������ˣ���Ҫ�����Ƿ������ó���˵���Ŀ�����
                    blnReSet����˵�� = True
                    blnCurDo = True
                End If
                If txtƵ��.Enabled And cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                    .TextMatrix(i, COL_Ƶ��) = txtƵ��.Text
                    Call GetƵ����Ϣ_����(txtƵ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ, 2) '��ҽ��Χ
                    .TextMatrix(i, COL_Ƶ�ʴ���) = intƵ�ʴ���
                    .TextMatrix(i, COL_Ƶ�ʼ��) = intƵ�ʼ��
                    .TextMatrix(i, COL_�����λ) = str�����λ
                    blnCurDo = True
                End If
                If cboִ��ʱ��.Tag <> "" Then
                    .TextMatrix(i, COL_ִ��ʱ��) = cboִ��ʱ��.Text
                    blnCurDo = True
                End If
                
                If .TextMatrix(i, COL_���) = "7" Then
                    '���ĵ��������ҩ��ִ�п���(�÷��巨�ĸĲ���)
                    If cboִ�п���.ListIndex <> -1 And cboִ�п���.Tag <> "" Then
                        .TextMatrix(i, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                        blnCurDo = True
                    End If
                    
                    'ִ������:�䷽��������ɵ���ҩ��ͬ
                    If cboִ������.Tag <> "" Then
                        .TextMatrix(i, COL_ִ������) = Decode(NeedName(cboִ������.Text), "�Ա�ҩ", 5, 4)
                        If Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                            .TextMatrix(i, COL_ִ�п���ID) = 0
                        ElseIf Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                            '�ָ�ȱʡִ�п���,ȱʡ��ǰ����ͬ
                            If i = lngBeginRow Then
                                For j = i - 1 To .FixedRows Step -1
                                    If .TextMatrix(j, COL_���) = "7" And Val(.TextMatrix(j, COL_ִ�п���ID)) <> 0 Then
                                        .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(j, COL_ִ�п���ID)
                                        Exit For
                                    End If
                                Next
                                If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), Val(.TextMatrix(i, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                                End If
                            Else
                                .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngBeginRow, COL_ִ�п���ID)
                            End If
                        End If
                        blnReInRow = True '����ִ�п��ұ༭�Ա仯
                        blnCurDo = True
                    End If
                End If
                
                '�޸�ʱ�Զ����²�������
                blnTmp = False
                If cboҽ������.Tag <> "" Or cboִ������.Tag <> "" _
                    Or (Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "") Then
                    blnTmp = True
                End If
                If blnCurDo Or blnTmp Then
                    '�޸�����������¿���ʱ��
                    If strCurDate = "" Then
                        strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                    End If
                    .TextMatrix(i, COL_����ʱ��) = Format(strCurDate, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_����ʱ��) = strCurDate
                    
                    '��鿪��ҽ��
                    If .TextMatrix(i, COL_����ҽ��) <> UserInfo.���� Then
                        .TextMatrix(i, COL_����ҽ��) = UserInfo.����
                        If lng��������ID = 0 Then
                            lng��������ID = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
                        End If
                        .TextMatrix(i, COL_��������ID) = lng��������ID
                    End If
                End If
                                                    
                If .TextMatrix(i, COL_���) = "E" And i <> lngRow Then lngTmp = i '�巨�к�
                                                    
                '---------------
                If blnCurDo Then '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2
                        .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                        If Not .RowHidden(i) Then Call ReSetColor(i) '�÷��в�����
                    End If
                    mblnNoSave = True '���Ϊδ����
                End If
            Next
            
            '�漰��ҩ�÷��е�����:ֱ�Ӹ��ĵ�ǰ�е�����(�巨�����䷽�༭�в��ܸ�)
            '-----------------------------------------------------------
            blnCurDo = False
                    
            'ҽ������:�Ƿ�����ҩ�÷���(��ʾ��)�е�
            If cboҽ������.Tag <> "" Then
                .TextMatrix(lngRow, COL_ҽ������) = cboҽ������.Text
                blnCurDo = True
            End If
        
            '��ҩ�÷�
            If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                .TextMatrix(lngRow, COL_������ĿID) = Val(cmd�÷�.Tag)
                .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                
                'ͬʱ���ļƼ����ʺ�ִ������
                .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "�Ƽ�����"), 0)
                i = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "ִ�п���"), 0)
                .TextMatrix(lngRow, COL_ִ������) = Decode(NeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, "E", Val(cmd�÷�.Tag), 0, _
                        Val(.TextMatrix(lngRow, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                End If
                
                blnReInRow = True '��Ҫˢ����ҩ�÷�ִ�п���
                blnCurDo = True
            End If
            
            '�÷��ͼ巨��ִ������
            If cboִ������.Tag <> "" Then
                '�÷�
                i = Nvl(GetItemField("������ĿĿ¼", Val(.TextMatrix(lngRow, COL_������ĿID)), "ִ�п���"), 0)
                .TextMatrix(lngRow, COL_ִ������) = Decode(NeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(lngRow, COL_���), _
                        Val(.TextMatrix(lngRow, COL_������ĿID)), 0, Val(.TextMatrix(lngRow, COL_ִ������)), _
                        mlng���˿���id, Val(Val(.TextMatrix(lngRow, COL_��������ID))), 1, 1)
                End If
                
                '�巨
                i = Nvl(GetItemField("������ĿĿ¼", Val(.TextMatrix(lngTmp, COL_������ĿID)), "ִ�п���"), 0)
                .TextMatrix(lngTmp, COL_ִ������) = Decode(NeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngTmp, COL_ִ������)) = 5 Then
                    .TextMatrix(lngTmp, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngTmp, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(lngTmp, COL_���), _
                        Val(.TextMatrix(lngTmp, COL_������ĿID)), 0, Val(.TextMatrix(lngTmp, COL_ִ������)), _
                        mlng���˿���id, Val(.TextMatrix(lngTmp, COL_��������ID)), 1, 1)
                End If
                
                If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                    .TextMatrix(lngTmp, COL_EDIT) = 2
                    .TextMatrix(lngTmp, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                End If
                mblnNoSave = True '���Ϊδ����
                
                blnCurDo = True
            End If
            
            '��ҩ�÷�ִ�п���:���䷽��ǰ��ʾ�е�ִ�п���
            If cbo����ִ��.ListIndex <> -1 And cbo����ִ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_ִ�п���ID) = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                blnCurDo = True
            End If
            
            Call Set��ҩ�����Ƿ���(lngRow)
            
        Else '����������Ŀ
            If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
                blnCurDo = True
            End If
            If IsDate(txt����ʱ��.Text) And txt����ʱ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_����ʱ��) = txt����ʱ��.Text
                blnCurDo = True
            End If
            '���Ա��
            If chk����.Visible And chk����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = chk����.value
                blnCurDo = True
            End If
            
            If chk����.Visible And chk����.Tag <> "" Then
                .TextMatrix(lngRow, COL_��־) = chk����.value
                
                '���洦��һ����ҩ�е������У�����˵ĸ�Ϊ�������
                If Not (.TextMatrix(lngRow, COL_���) = "K") Then
                    If chk����.value = 1 Then
                        If Val(.TextMatrix(lngRow, COL_���״̬)) = 1 Then .TextMatrix(lngRow, COL_���״̬) = ""
                    Else
                        If gblnKSSStrict And UserInfo.��ҩ���� < Val(.TextMatrix(lngRow, COL_�����ȼ�)) Then .TextMatrix(lngRow, COL_���״̬) = 1
                    End If
                End If
                If (gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ) And .TextMatrix(lngRow, COL_���) = "K" Then
                    blnReInRow = True
                End If
                '��ʾ������־,һ����ҩ��ʾ�ڵ�һ��
                Call SetRow��־ͼ��(lngRow, 0)
                                
                blnCurDo = True
            End If
            If txt����.Enabled And (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = FormatEx(txt����.Text, 5)
                If Not mbln���� Then Call Set��ҩ�����Ƿ���(lngRow): blnReSet����˵�� = True
                               
                blnCurDo = True
            End If
            
            If txt����.Visible And txt����.Enabled And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = txt����.Text
                
                If Val(txt����.Text) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                    .TextMatrix(lngRow, COL_�Ƿ���) = "1"
                Else
                    .TextMatrix(lngRow, COL_�Ƿ���) = ""
                End If
                blnReSet����˵�� = True
                
                blnCurDo = True
            End If
            
            If txt����.Enabled And (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = FormatEx(txt����.Text, 5)
                If Not mbln���� Then Call Set��ҩ�����Ƿ���(lngRow)
                
                '�����仯�����ó���˵���Ŀ�����
                blnReSet����˵�� = True
                               
                blnCurDo = True
            End If
            
            If txtƵ��.Enabled And cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_Ƶ��) = txtƵ��.Text
                Call GetƵ����Ϣ_����(txtƵ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ, GetƵ�ʷ�Χ(lngRow))
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
                
                If Not mbln���� Then Call Set��ҩ�����Ƿ���(lngRow): blnReSet����˵�� = True
                
                blnCurDo = True
            End If
            
            If cboִ��ʱ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_ִ��ʱ��) = cboִ��ʱ��.Text
                blnCurDo = True
            End If
            If cboҽ������.Tag <> "" Then
                .TextMatrix(lngRow, COL_ҽ������) = cboҽ������.Text
                blnCurDo = True
            End If
            
            If cboִ�п���.ListIndex <> -1 And cboִ�п���.Tag <> "" Then
                If Not RowIn������(lngRow) Then '�ɼ�������ִ�п��Ҳ�ͬ
                    .TextMatrix(lngRow, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                End If
                blnCurDo = True
            End If
            
            
            '��ҩĿ�ĺ�����
            If lbl��ҩĿ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_��ҩĿ��) = cboDruPur.ListIndex
                blnCurDo = True
            End If
            If txt��ҩ����.Tag <> "" Then
                .TextMatrix(lngRow, COL_��ҩ����) = Trim(txt��ҩ����.Text)
                blnCurDo = True
            End If
            
            '���٣���ҺҩƷ
            If cbo����.Tag <> "" Then
                lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                If lngTmp <> -1 Then
                    If cbo����.Text <> "" Then
                        .TextMatrix(lngTmp, COL_ҽ������) = cbo����.Text & lbl���ٵ�λ.Caption
                    Else
                        .TextMatrix(lngTmp, COL_ҽ������) = ""
                    End If
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    End If
                    'mblnNoSave = True '���Ϊδ����
                    blnCurDo = True
                End If
                '��ʾ��ҩ;��
                If cbo����.Text <> "" Then
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text & cbo����.Text & lbl���ٵ�λ.Caption
                Else
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                End If
            End If
            
            '����ִ�п��ң���ҩ;��,��������,�ɼ�������ԭҺƤ����Ŀ
            If cbo����ִ��.ListIndex <> -1 And cbo����ִ��.Tag <> "" Then
                lngTmp = -1
                If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                ElseIf .TextMatrix(lngRow, COL_���) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5" Then
                    lngTmp = lngRow 'ԭҺƤ����Ŀ
                ElseIf .TextMatrix(lngRow, COL_���) = "F" Then
                    For i = lngRow + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If .TextMatrix(i, COL_���) = "G" Then
                                lngTmp = i: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf .TextMatrix(lngRow, COL_���) = "E" _
                    And .TextMatrix(lngRow - 1, COL_���) = "C" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow
                ElseIf .TextMatrix(lngRow, COL_���) = "K" _
                    And .TextMatrix(lngRow + 1, COL_���) = "E" _
                    And Val(.TextMatrix(lngRow + 1, COL_���ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow + 1
                End If
                
                'ֻ���¶�Ӧ��,��Ӱ��������
                If lngTmp <> -1 Then
                    'ԭҺƤ��
                    If Not (.TextMatrix(lngRow, COL_���) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5") Then
                        .TextMatrix(lngTmp, COL_ִ�п���ID) = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                    End If
                    
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    End If
                    'mblnNoSave = True '���Ϊδ����
                    blnCurDo = True
                End If
            End If
            
            'ԭҺƤ����Ŀ���Բ�ָ�����ӵ�ҩ������
            If cbo����ִ��.Tag <> "" Then
                If .TextMatrix(lngRow, COL_���) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5" Then
                    lngTmp = lngRow
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    End If
                    'mblnNoSave = True '���Ϊδ����
                    blnCurDo = True
                End If
            End If
            
            'ִ������,��ҩ;��:Ϊ���¿���ʱ��(������ҩ;����ͬ������),���ж��Ƿ�ı�
            If InStr(",5,6,K,", .TextMatrix(lngRow, COL_���)) > 0 Then
                If cboִ������.Tag <> "" Then blnCurDo = True
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then blnCurDo = True
            End If
                                    
            '�޸�ʱ�Զ����²�������
            blnTmp = False
            If cboִ������.Tag <> "" Or (Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "") Then
                blnReInRow = True '��Ҫˢ�¸�ҩ;��,�ɼ���ʽ��ִ�п���
                blnTmp = True
            End If
            If blnCurDo Or blnTmp Then
                '�޸�����������¿���ʱ��
                strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                .TextMatrix(lngRow, COL_����ʱ��) = Format(strCurDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_����ʱ��) = strCurDate
                
                '��鿪��ҽ��
                If .TextMatrix(lngRow, COL_����ҽ��) <> UserInfo.���� Then
                    .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
                    If lng��������ID = 0 Then
                        lng��������ID = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
                    End If
                    .TextMatrix(lngRow, COL_��������ID) = lng��������ID
                End If
            End If
                                    
            '������Ҫͬ������Ĺ�����
            '----------------------------------------------------------------
            If RowIn������(lngRow) Then
                '�ɼ�����
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_������ĿID) = Val(cmd�÷�.Tag)
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                    .TextMatrix(lngRow, COL_����) = txt�÷�.Text
                    
                    'ͬʱ���ļƼ����ʺ�ִ������
                    .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "�Ƽ�����"), 0)
                    .TextMatrix(lngRow, COL_ִ������) = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "ִ�п���"), 0)
                    If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, "E", Val(cmd�÷�.Tag), 0, _
                            Val(.TextMatrix(lngRow, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                    Else
                        .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                    End If

                    blnCurDo = True
                End If
                
                '����һ���ɼ��ĸ���������Ŀ
                If blnCurDo Then
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                blnOtherDo = True
                            End If
                            If txtƵ��.Tag <> "" Then
                                .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                                blnOtherDo = True
                            End If
                            If cboִ�п���.Tag <> "" And cboִ�п���.ListIndex <> -1 Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = 0
                                Else
                                    .TextMatrix(i, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                                End If
                                blnOtherDo = True
                            End If
                            If cboִ��ʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                blnOtherDo = True
                            End If
                            If txt��ʼʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                                .Cell(flexcpData, i, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                                blnOtherDo = True
                            End If
                            If chk����.Tag <> "" Then
                                .TextMatrix(i, COL_��־) = .TextMatrix(lngRow, COL_��־)
                                blnOtherDo = True
                            End If
                                            
                            '����ʱ��
                            If .TextMatrix(i, COL_����ʱ��) <> .TextMatrix(lngRow, COL_����ʱ��) Then
                                .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                                .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                                blnOtherDo = True
                            End If
                            
                            '����ҽ��
                            If .TextMatrix(i, COL_����ҽ��) <> .TextMatrix(lngRow, COL_����ҽ��) Then
                                .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                                blnOtherDo = True
                            End If
                                            
                            '��������ID
                            If .TextMatrix(i, COL_��������ID) <> .TextMatrix(lngRow, COL_��������ID) Then
                                .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                                blnOtherDo = True
                            End If
                            
                            '���Ϊ�޸�
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '�С�����ҩ�����ҩ;����һ����ҩ�����
                
                'ִ������
                If cboִ������.Tag <> "" Then
                    .TextMatrix(lngRow, COL_ִ������) = Decode(NeedName(cboִ������.Text), "�Ա�ҩ", 5, 4)
                    If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                    ElseIf Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                        '�ָ�ȱʡҩ��,ȱʡ��ǰ��ĳ�ҩ��ͬ
                        strTmp = Get����ҩ��IDs(.TextMatrix(lngRow, COL_���), Val(.TextMatrix(lngRow, COL_������ĿID)), Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), mlng���˿���id, 1)
                        For i = lngRow - 1 To .FixedRows Step -1
                            '����ҩ���г�ҩ��ҩ�����ܲ�ͬ,�������Ҫ��ͬ
                            If .TextMatrix(i, COL_���) = .TextMatrix(lngRow, COL_���) And Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 Then
                                If InStr("," & strTmp & ",", "," & Val(.TextMatrix(i, COL_ִ�п���ID)) & ",") > 0 Then
                                    .TextMatrix(lngRow, COL_ִ�п���ID) = Val(.TextMatrix(i, COL_ִ�п���ID))
                                    Exit For
                                End If
                            End If
                        Next
                        If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                            .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(lngRow, COL_���), Val(.TextMatrix(lngRow, COL_������ĿID)), Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                        End If
                    End If
                    cboִ�п���.Tag = "1" '����ִ�п���һ����ҩ��Ҫͬ����
                    blnReInRow = True '����ִ�п��ұ༭�Ա仯
                End If
                
                '��ҩ;�����������������ͬ������
                lngִ�п���ID = 0
                If cbo����ִ��.ListIndex <> -1 Then
                    lngִ�п���ID = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                End If
                strTmp = ""
                If Trim(cbo����.Text) <> "" Then
                    strTmp = cbo����.Text & lbl���ٵ�λ.Caption
                End If
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text & strTmp
                    
                    If CheckExecDeptValidate(lngִ�п���ID, mlng���˿���id, 1, Val(cmd�÷�.Tag)) = False Then
                        lngִ�п���ID = 0
                    End If
                    
                    Call AdviceSet��ҩ;��(lngRow, Val(cmd�÷�.Tag), NeedName(cboִ������.Text), lngִ�п���ID, strTmp)
                ElseIf blnCurDo Then 'cboִ������.Tag <> "" Then
                    '���ִ�����ʸ�����,��Ҫǿ���޸Ķ�Ӧ�ĸ�ҩ;����ִ�����ʺ�ִ�п���
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    Call AdviceSet��ҩ;��(lngRow, Val(.TextMatrix(lngTmp, COL_������ĿID)), NeedName(cboִ������.Text), lngִ�п���ID, strTmp)
                End If
                
                'һ����ҩ:�������ҩ;��,ǰ���ѵ�������
                If blnCurDo Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_���ID), , COL_���ID)
                    If cboִ�п���.Tag <> "" Then
                        For i = lngBeginRow To .Rows - 1
                            If .TextMatrix(i, COL_���ID) = "" Then
                                lngTmp = i: Exit For
                            End If
                        Next
                    End If
                    For i = lngBeginRow To .Rows - 1
                        If i <> lngRow And .RowData(i) <> 0 _
                            And Val(.TextMatrix(i, COL_Ӥ��)) = Val(.TextMatrix(lngRow, COL_Ӥ��)) Then '���������м��п���
                            If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                                If txt��ʼʱ��.Tag <> "" Then
                                    .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                                    .Cell(flexcpData, i, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                                    blnOtherDo = True
                                End If
                                If txt�÷�.Tag <> "" Then
                                    .TextMatrix(i, COL_�÷�) = .TextMatrix(lngRow, COL_�÷�)
                                    blnOtherDo = True
                                End If
                                If txtƵ��.Tag <> "" Then
                                    .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                    .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                    .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                    .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                                    blnOtherDo = True
                                End If
                                
                                '�����ٲ��һ����ҩ����������
                                If cbo����.Tag <> "" Then
                                    .TextMatrix(i, COL_�÷�) = txt�÷�.Text & strTmp
                                    blnOtherDo = True
                                End If
                                
                                'һ����ҩ��,������ͬ�仯,�������¼���
                                If txt����.Tag <> "" Then
                                    .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                    If .TextMatrix(i, COL_Ƶ��) <> "" _
                                        And Val(.TextMatrix(i, COL_����)) <> 0 _
                                        And Val(.TextMatrix(i, COL_����ϵ��)) <> 0 _
                                        And Val(.TextMatrix(i, COL_�����װ)) <> 0 Then
                                        
                                        .TextMatrix(i, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                            Val(.TextMatrix(i, COL_����)), Val(.TextMatrix(i, COL_����)), _
                                            Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), _
                                            .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), _
                                            Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_�����װ)), _
                                            Val(.TextMatrix(i, COL_�ɷ����))), 5)
                                        If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") = 0 Then
                                            .TextMatrix(i, COL_����) = IntEx(Val(.TextMatrix(i, COL_����)))
                                        ElseIf Val(.TextMatrix(i, COL_�ɷ����)) <> 0 Then
                                            .TextMatrix(i, COL_����) = IntEx(Val(.TextMatrix(i, COL_����)))
                                        End If
                                    End If
                                    '���¼�鴦����������
                                    .TextMatrix(i, COL_�Ƿ���) = "": .TextMatrix(i, COL_�Ƿ���) = ""
                                    If Val(.TextMatrix(i, COL_��������)) <> 0 Then
                                        dbl���� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_�����װ)) * Val(.TextMatrix(i, COL_����ϵ��))
                                        If dbl���� > Val(.TextMatrix(i, COL_��������)) Then .TextMatrix(i, COL_�Ƿ���) = "1"
                                    End If
                                    
                                    If Val(.TextMatrix(i, COL_����)) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                                        .TextMatrix(i, COL_�Ƿ���) = "1"
                                    End If
                                    
                                    If .TextMatrix(i, COL_�Ƿ���) = "" And .TextMatrix(i, COL_�Ƿ���) = "" Then .TextMatrix(i, COL_����˵��) = ""
                                    
                                    blnOtherDo = True
                                End If
                                
                                If cboִ��ʱ��.Tag <> "" Then
                                    .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                    blnOtherDo = True
                                End If
                                
                                'ִ������:��Ժ��ҩ��һ����ҩ����һ�£������ɵ�������
                                If cboִ������.Tag <> "" And NeedName(cboִ������.Text) = "��Ժ��ҩ" Then
                                    .TextMatrix(i, COL_ִ������) = .TextMatrix(lngRow, COL_ִ������)
                                    '���Ա�ҩת����ʱ��Ҫ��������ִ�п���
                                    If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                                        .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                    End If
                                    blnOtherDo = True
                                End If
                                
                                'ִ�п���:ִ�п���(ҩ��)���Բ�ͬ,��������������
                                If cboִ�п���.Tag <> "" Then
                                    '�����и�Ϊ�Ա�ҩ����ĳ��Ϊ�Ա�ҩ�����������
                                    If Not (Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow, COL_ִ������)) = 5) _
                                        And Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                                        If .TextMatrix(lngTmp, COL_���) = "E" And .TextMatrix(lngTmp, COL_��������) = "2" And .TextMatrix(lngTmp, COL_ִ�з���) = "1" Then
                                            If Have��������(Val(.TextMatrix(lngRow, COL_ִ�п���ID)), "��������") Then
                                                '������ҩƷ����ͨҩ���������������ĸ�Ϊ�µ���������,�����ҩ����Ϊ����������
                                                .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                                blnOtherDo = True
                                            ElseIf Have��������(Val(.TextMatrix(i, COL_ִ�п���ID)), "��������") Then
                                                '������ҩƷ���������ĸĳ���ͨҩ��,�����ҩ����Ϊ����ͨҩ��
                                                .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                                blnOtherDo = True
                                            End If
                                        End If
                                    End If
                                End If
                                
                                '������־
                                If chk����.Tag <> "" Then
                                    .TextMatrix(i, COL_��־) = .TextMatrix(lngRow, COL_��־)
                                    .TextMatrix(i, COL_���״̬) = .TextMatrix(lngRow, COL_���״̬)
                                    blnOtherDo = True
                                End If
                                
                                '����ʱ��
                                If .TextMatrix(i, COL_����ʱ��) <> .TextMatrix(lngRow, COL_����ʱ��) Then
                                    .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                                    .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                                    blnOtherDo = True
                                End If
                                
                                '����ҽ��
                                If .TextMatrix(i, COL_����ҽ��) <> .TextMatrix(lngRow, COL_����ҽ��) Then
                                    .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                                    blnOtherDo = True
                                End If
                                
                                '��������ID
                                If .TextMatrix(i, COL_��������ID) <> .TextMatrix(lngRow, COL_��������ID) Then
                                    .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                                    blnOtherDo = True
                                End If
                                
                                '���Ϊ�޸�
                                If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                    .TextMatrix(i, COL_EDIT) = 2
                                    .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    Next
                End If
            ElseIf .TextMatrix(lngRow, COL_���) = "K" Then
                '��Ѫҽ���Ĵ���(ǰ���Ѵ�����Ѫʱ��(����ʱ��)���޸�)
                
                lngִ�п���ID = 0
                If cbo����ִ��.ListIndex <> -1 Then
                    lngִ�п���ID = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                End If
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                    Call AdviceSet��Ѫ;��(lngRow, Val(cmd�÷�.Tag), lngִ�п���ID)
                ElseIf blnCurDo Then
                    lngTmp = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
                    If lngTmp <> -1 Then
                        Call AdviceSet��Ѫ;��(lngRow, Val(.TextMatrix(lngTmp, COL_������ĿID)), lngִ�п���ID)
                    End If
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_���)) > 0 And blnCurDo Then
                '��������Ŀ�л�����������
                lngBeginRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
                If lngBeginRow <> -1 Then
                    For i = lngBeginRow To .Rows - 1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                blnOtherDo = True
                            End If
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                blnOtherDo = True
                            End If
                            
                            If cboִ��ʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                blnOtherDo = True
                            End If
                            If txtƵ��.Tag <> "" Then
                                .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                                blnOtherDo = True
                            End If
                            If cboִ�п���.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = 0
                                ElseIf .TextMatrix(i, COL_���) <> "G" Then '���������ִ�п���Ϊ����
                                    .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                End If
                                blnOtherDo = True
                            End If
                            If txt��ʼʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                                .Cell(flexcpData, i, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                                blnOtherDo = True
                            End If
                            If txt����ʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                                blnOtherDo = True
                            End If
                            If chk����.Tag <> "" Then
                                .TextMatrix(i, COL_��־) = .TextMatrix(lngRow, COL_��־)
                                blnOtherDo = True
                            End If
                            
                            '����ʱ��
                            If .TextMatrix(i, COL_����ʱ��) <> .TextMatrix(lngRow, COL_����ʱ��) Then
                                .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                                .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                                blnOtherDo = True
                            End If
                            
                            '����ҽ��
                            If .TextMatrix(i, COL_����ҽ��) <> .TextMatrix(lngRow, COL_����ҽ��) Then
                                .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                                blnOtherDo = True
                            End If
                            
                            '��������ID
                            If .TextMatrix(i, COL_��������ID) <> .TextMatrix(lngRow, COL_��������ID) Then
                                .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                                blnOtherDo = True
                            End If
                            
                            '���Ϊ�޸�
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf .TextMatrix(lngRow, COL_���) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5" Then
                'ԭҺƤ��
                If cbo����ִ��.Tag <> "" Then
                    lngִ�п���ID = 0
                    If cbo����ִ��.ListIndex <> -1 Then
                        lngִ�п���ID = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                    End If
                    .TextMatrix(lngRow, COL_��ҩ����) = lngִ�п���ID
                End If
            End If
        End If
        
        '�䷽������������Ŀ�����ܴ��ڳ���
        If txt����˵��.Enabled And txt����˵��.Tag <> "" Then
            .TextMatrix(lngRow, COL_����˵��) = txt����˵��.Text
            blnCurDo = True
        End If
        If lbl����˵��.Tag <> "" Then
            blnReSet����˵�� = True
        End If
        'ͳһ��������е����Ա仯
        If chkZeroBilling.Tag <> "" Then
            Call GetRowScope(lngRow, lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                .TextMatrix(i, COL_��Ѽ���) = chkZeroBilling.value
            Next
            blnCurDo = True
        End If
                    
        '---------------
        If blnCurDo Then '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
            If InStr(",0,2,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                 '���δͨ������һ��ҩƷ,����Ҫ�ı����������е����״̬Ϊδ��˻��������
                If Val(.TextMatrix(lngRow, COL_���״̬)) <> 2 And .TextMatrix(lngRow, COL_���) <> "K" Then
                    Call GetRowScope(lngRow, lngBeginRow, lngEndRow)
                    For i = lngBeginRow To lngEndRow
                        '����ǽ���ҽ�������������
                        If gblnKSSStrict And UserInfo.��ҩ���� < Val(.TextMatrix(i, COL_�����ȼ�)) And .TextMatrix(i, COL_��־) <> 1 Then
                            .TextMatrix(i, COL_���״̬) = 1
                        Else
                            .TextMatrix(i, COL_���״̬) = ""
                        End If
                        Call SetRow��־ͼ��(i, 2)
                    Next
                End If
                
                .TextMatrix(lngRow, COL_EDIT) = 2
                .TextMatrix(lngRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                Call ReSetColor(lngRow)
            End If
            mblnNoSave = True '���Ϊδ����
        End If
        
        '����ҽ������
        If AdviceTextChange(lngRow) Then
            .TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = .TextMatrix(lngRow, col_ҽ������)
        End If
    End With
        
    '����༭��־
    Call ClearItemTag
    
    'ĳЩ�������Ҫ�������ÿ�Ƭ����Ŀ�༭��(���޸���ִ������ʱ)
    If blnReInRow Then
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        If vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "" And vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "" Then vsAdvice.TextMatrix(lngRow, COL_����˵��) = ""
    ElseIf blnReSet����˵�� Then
        SetItemEditable , , , , , , , , , , , , IIF(vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "1" Or vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "1", 1, -1)
        If vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "" And vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "" Then vsAdvice.TextMatrix(lngRow, COL_����˵��) = ""
    End If
End Sub

Private Sub ReSetColor(ByVal lngRow As Long)
'���ܣ���������ָ���е���ɫ
'˵����
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    With vsAdvice
        'һ����ҩ��Χ
        lngBegin = lngRow: lngEnd = lngRow
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            If RowInһ����ҩ(lngRow) Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow, COL_���ID)), lngBegin, lngEnd)
            End If
        End If
        '�ָ�������ɫ
        For i = lngBegin To lngEnd
            .Cell(flexcpForeColor, i, .FixedCols, i, COL_����ҽ��) = .ForeColor
            '���龫����ɫ��ʶ
            If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(i, COL_�������)) > 0 _
                And .TextMatrix(i, COL_�������) <> "" Then
                .Cell(flexcpFontBold, i, col_ҽ������) = True
            End If
        Next
        .ForeColorSel = .Cell(flexcpForeColor, lngRow, COL_��ʼʱ��)
    End With
End Sub

Private Sub AdviceSetһ����ҩ(ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ���ѡ��Χ�ڵ�ҩƷ����Ϊһ����ҩ
'��������ֹ�к�,�м䲻��������,���������һ��ҩƷ�ĸ�ҩ;����
'˵�����Ե�һ��ҩƷ�ĸ�ҩ;��Ϊ׼,��λ�÷������һ��ҩƷ֮��
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lngRow1 As Long, lngRow2 As Long
    Dim lng���ID As Long, i As Long
    Dim strStart As String, curDate As Date
    Dim lng�������� As Long
        
    With vsAdvice
        lngRow1 = .FindRow(CLng(.TextMatrix(lngBegin, COL_���ID)), lngBegin + 1) '��һ��ҩ;����
        lngRow2 = .FindRow(CLng(.TextMatrix(lngEnd, COL_���ID)), lngEnd + 1) '����ҩ;����
        
        
        'ɾ����ҩ;����֮ǰ��¼ִ������,�Ա�������ж�
        For i = lngRow2 To lngRow1 Step -1
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex And .RowHidden(i) Then
                .Cell(flexcpData, i - 1, COL_ִ������) = Val(.TextMatrix(i, COL_ִ������))
            End If
        Next
        
        '���Ƶ�һ�еĸ�ҩ;�������һ�еĸ�ҩ;��
        For i = .FixedCols To .Cols - 1
            If i <> COL_EDIT And i <> COL_���ID And i <> COL_��� And i <> COL_״̬ Then
                .TextMatrix(lngRow2, i) = .TextMatrix(lngRow1, i)
            End If
        Next
        .Cell(flexcpData, lngRow2, COL_��ʼʱ��) = .TextMatrix(lngRow2, COL_��ʼʱ��)
        
        '�༭��־��0-ԭʼ��,1-������,2-�޸�������,3-�޸������
        If InStr(",0,3,", .TextMatrix(lngRow2, COL_EDIT)) > 0 Then
            .TextMatrix(lngRow2, COL_EDIT) = 2 '���Ϊ���޸�
            .TextMatrix(lngRow2, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
        End If
        lng���ID = .RowData(lngRow2)
        
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        'ɾ�������һ�и�ҩ;�����������ҩ;��
        For i = lngEnd To lngBegin Step -1
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                If .RowHidden(i) Then
                    Call DeleteRow(i)
                Else
                    .TextMatrix(i, COL_���ID) = lng���ID
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2 '���Ϊ���޸�
                        .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    End If
                End If
            End If
        Next
        
        '�к��ѱ��
        lngRow1 = lngBegin '��ʼһ����ҩ��
        curDate = zlDatabase.Currentdate
        
        '���ҽ���Ƿ���
        If .TextMatrix(lngRow1, COL_����ҽ��) <> UserInfo.���� Then
            '���������Ϣ:ǰ���ѱ��Ϊ�޸�,���ֹ��������ʱ���н������ˢ��
            .TextMatrix(lngRow1, COL_����ҽ��) = UserInfo.����
            .TextMatrix(lngRow1, COL_��������ID) = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
            
            .TextMatrix(lngRow1, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow1, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
        
        '����һ����ҩ�����е���ͬ��Ϣ
        For i = lngRow1 + 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                    lngRow2 = i '��¼�µĽ����к�
                    
                    'һ����ҩ�Ĳ�����Ϣ��ͬ
                    .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow1, COL_��ʼʱ��)
                    .Cell(flexcpData, i, COL_��ʼʱ��) = .Cell(flexcpData, lngRow1, COL_��ʼʱ��)
                    
                    .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow1, COL_����ҽ��)
                    .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow1, COL_��������ID)
                    
                    .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow1, COL_����ʱ��) 'һ����ҩ�Ŀ���ʱ����ͬ
                    .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow1, COL_����ʱ��)
                    
                    .TextMatrix(i, COL_����) = .TextMatrix(lngRow1, COL_����)
                    
                    .TextMatrix(i, COL_�÷�) = .TextMatrix(lngRow1, COL_�÷�)
                    
                    .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow1, COL_Ƶ��)
                    .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow1, COL_Ƶ�ʴ���)
                    .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow1, COL_Ƶ�ʼ��)
                    .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow1, COL_�����λ)
                    .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow1, COL_ִ��ʱ��)
                    
                    '�������������Ʋŷ��㣨71152��
                    If mbln���� Then .TextMatrix(i, COL_����) = ReGetҩƷ����(Val(.TextMatrix(i, COL_����)), Val(.TextMatrix(i, COL_����)), Val(.TextMatrix(i, COL_����)), i)
                    
                    
                    '��������
                    .TextMatrix(i, COL_�Ƿ���) = ""
                    If Val(.TextMatrix(i, COL_��������)) <> 0 Then
                        If Val(.TextMatrix(i, COL_����)) > FormatEx(Val(.TextMatrix(i, COL_��������)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ)), 5) Then
                            .TextMatrix(i, COL_�Ƿ���) = "1"
                        End If
                    End If
                    
                    .TextMatrix(i, COL_�Ƿ���) = ""
                    Call Set��ҩ�����Ƿ���(i)
                    
                    .TextMatrix(i, COL_��־) = .TextMatrix(lngRow1, COL_��־)
                    Set .Cell(flexcpPicture, i, COL_F��־) = Nothing '�ڿ�ʼ����ʾ
                    
                    '��Ժ��ҩһ����ͬ
                    If Val(.TextMatrix(lngRow1, COL_ִ������)) <> 5 And Val(.Cell(flexcpData, lngRow1, COL_ִ������)) = 5 Then
                        '��һ������Ժ��ҩ,ȫ������Ϊ��Ժ��ҩ
                        .TextMatrix(i, COL_ִ������) = .TextMatrix(lngRow1, COL_ִ������)
                        If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then 'ִ�п��ҿ��Բ�ͬ,û��ʱ��ȱʡ��ͬ
                            .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow1, COL_ִ�п���ID)
                        End If
                    ElseIf Val(.TextMatrix(i, COL_ִ������)) <> 5 And Val(.Cell(flexcpData, i, COL_ִ������)) = 5 Then
                        '��ǰ������Ժ��ҩ,������Ϊ���һ����ͬ
                        .TextMatrix(i, COL_ִ������) = .TextMatrix(lngRow1, COL_ִ������)
                        If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                            .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow1, COL_ִ�п���ID)
                        End If
                    Else
                        '���򱣳ֲ���
                    End If
                    
                    '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2
                        .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    End If
                Else
                    Exit For
                End If
            End If
        Next
    
        '�����ЩҩƷ���Ƿ��������������ҩ�ģ��Ե�һ��Ϊ׼
        For i = lngRow1 To .Rows - 1
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                    '�Ա�ҩ�����������
                    If Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                        If Have��������(Val(.TextMatrix(i, COL_ִ�п���ID)), "��������") Then
                            lng�������� = Val(.TextMatrix(i, COL_ִ�п���ID)): Exit For
                        End If
                    End If
                Else
                    Exit For
                End If
            End If
        Next
        '��������һ����ͬ
        If lng�������� <> 0 Then
            For i = lngRow1 To .Rows - 1
                If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                    If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                        '�Ա�ҩ�����������
                        If Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                            .TextMatrix(i, COL_ִ�п���ID) = lng��������
                            
                            '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
                            If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                            End If
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
        
        '��ʼִ��ʱ�䴦��(�¿��Ĳ���̫��)
        strStart = ""
        For i = lngRow1 To lngRow2
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                If Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                    If DateDiff("n", CDate(.Cell(flexcpData, i, COL_��ʼʱ��)), curDate) > gint�����¿�ҽ����� Then
                        strStart = GetDefaultTime(i): Exit For
                    End If
                End If
            End If
        Next
        If strStart <> "" Then
            For i = lngRow1 To lngRow2 + 1
                If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                    .Cell(flexcpData, i, COL_��ʼʱ��) = strStart
                    .TextMatrix(i, COL_��ʼʱ��) = Format(strStart, "yyyy-MM-dd HH:mm")
                End If
            Next
        End If
        
        Call ReSet���״̬ͼ��(lngBegin)
    
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '���Ϊδ����
    End With
End Sub

Private Sub AdviceSet������ҩ(ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ�ȡ��һ��ҩƷ��һ����ҩ
'��������ֹ�к�,�м䲻��������,���������һ��ҩƷ�ĸ�ҩ;����
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lng��ҩ;��ID As Long, lng��ҩִ��ID As Long, i As Long
    Dim intִ������ As Integer, strִ������ As String, str���� As String
    Dim lngRow As Long, curDate As Date, blnUpdate As Boolean
    
    With vsAdvice
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        'һ����ҩ;��
        lngRow = .FindRow(CLng(.TextMatrix(lngEnd, COL_���ID)), lngEnd + 1)
        lng��ҩ;��ID = Val(.TextMatrix(lngRow, COL_������ĿID))
        lng��ҩִ��ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        intִ������ = Val(.TextMatrix(lngRow, COL_ִ������))
        str���� = .TextMatrix(lngRow, COL_ҽ������)
                
        '���ҽ�����:�Ը�ҩ;����Ϊ׼�仯
        If .TextMatrix(lngRow, COL_����ҽ��) <> UserInfo.���� Then
            '���������Ϣ:�ֹ��������ʱ�н������ˢ��
            .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
            .TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 1)
            curDate = zlDatabase.Currentdate
            .TextMatrix(lngRow, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
            
            If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
                .TextMatrix(lngRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            End If
            blnUpdate = True
        End If
                
        '��ʾ������־:ÿһ��
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                'ҩƷ����Ӧ�仯
                If blnUpdate Then
                    .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                    .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                    .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                    .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2 '���Ϊ���޸�
                        .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    End If
                End If
                If gblnKSSStrict And UserInfo.��ҩ���� < Val(.TextMatrix(i, COL_�����ȼ�)) And .TextMatrix(i, COL_��־) <> "1" Then
                    .TextMatrix(i, COL_���״̬) = 1
                Else
                    .TextMatrix(i, COL_���״̬) = ""
                End If
                Call SetRow��־ͼ��(i, 2)
            End If
        Next
        
        For i = lngEnd - 1 To lngBegin Step -1 '���뷴��
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                '���ø�ҩ;����
                If Val(.TextMatrix(i, COL_ִ������)) = 5 And intִ������ <> 5 Then
                    strִ������ = "�Ա�ҩ"
                ElseIf Val(.TextMatrix(i, COL_ִ������)) <> 5 And intִ������ = 5 Then
                    strִ������ = "��Ժ��ҩ"
                Else
                    strִ������ = ""
                End If
                .TextMatrix(i, COL_���ID) = "" '���������Ϊ��־
                .TextMatrix(i, COL_���ID) = AdviceSet��ҩ;��(i, lng��ҩ;��ID, strִ������, lng��ҩִ��ID, str����)
                
                If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    .TextMatrix(i, COL_EDIT) = 2 '���Ϊ���޸�
                    .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                End If
            End If
        Next
        
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '���Ϊδ����
    End With
End Sub

Private Sub ShowAdvice()
'���ܣ���ʾ��ǰ���������µ�ҽ����¼
'˵����1.���ݳ���༭��ʽ,��ص��������ǰ�����ϸ�������һ�ڵġ�
'      2.���ﲻ����һ����ҩ�ı߿��䷽�иߣ�״̬��ɫ�ȸ�ʽ����,�������ڶ�ȡ��༭ʱ����
    Dim lngRow As Long, blnHide As Boolean, i As Long
    
    Screen.MousePointer = 11
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
        
    '��ɾ����Ч��
    For i = vsAdvice.Rows - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) = 0 Then vsAdvice.RemoveItem i
    Next
    
    '���ݵ�ǰ��Ч,Ӥ����ʾ
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                blnHide = False
                '�������������У�
                '1.��ҩ�ĸ�ҩ;����
                '2.�����ĸ���������������Ŀ��
                '3.�����ϵĲ�λ��
                '4.��ҩ�䷽�����ζ��ҩ����ҩ�巨��
                '5.(һ���ɼ���)������Ŀ
                '6.��Ѫ��Ŀ����Ѫ;��
                If .TextMatrix(i, COL_���) = "E" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                        blnHide = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_���)) > 0 _
                    And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                    blnHide = True
                End If
                                
                .RowHidden(i) = blnHide
                If Not blnHide And lngRow = 0 Then lngRow = i
                
                '�������Ƶ���:����Ϊ�ӿ��ٶ�,ֻ��ȡ�¿���,�����Ľ����ٶ�
                If Not .RowHidden(i) _
                    And Val(.TextMatrix(i, COL_״̬)) = 1 And .TextMatrix(i, COL_����) = "" Then
                    .TextMatrix(i, COL_����) = GetItemPrice(i)
                End If
            Else
                .RowHidden(i) = True
            End If
        Next
    End With
    
    'û��������,���һ�п�
    If lngRow = 0 Then
        vsAdvice.AddItem ""
        lngRow = vsAdvice.Rows - 1
    End If
    
    vsAdvice.Row = lngRow
    If vsAdvice.RowData(lngRow) = 0 Then
        vsAdvice.Col = vsAdvice.FixedCols
    Else
        vsAdvice.Col = col_ҽ������
    End If
    vsAdvice.Redraw = flexRDDirect
    mblnRowChange = True
    
    '��ʾ��ǰ��:����ʱ��FormLoad�д���,�Լӿ��ٶ�
    If Me.Visible Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    Call CalcAdviceMoney '��ʾ�¿�ҽ�����
    
    Screen.MousePointer = 0
End Sub

Private Function SaveAdvice() As Boolean
'���ܣ����浱ǰ���˵�ҽ����¼
    Dim arrSQL As Variant, arrDelID() As String
    Dim strSQL As String, dbl���� As Double
    Dim arrAppend As Variant, i As Long, j As Long
    Dim blnChecked As Boolean
    Dim curDate As Date
    Dim blnTrans As Boolean
    Dim blnDiagChange As Boolean, blnRecipeNo As Boolean
    Dim strFilter As String, strTmp As String
    Dim rsMsg As ADODB.Recordset
    Dim blnMsgOk As Boolean
    Dim lng��ϼ�¼id As Long
    Dim str��¼�� As String
    Dim str��ҩIDs As String
    Dim lng���ID As Long
    Dim str������ҩIDs As String
    Dim rsCard As ADODB.Recordset
    Dim rsBlood As ADODB.Recordset, intBloodState As Integer
    Dim varTmp As Variant
    Dim blnRIS As Boolean 'RIS�ӿ�
    Dim strNewAdvice As String '�������޸ĺ��ҽ�� ��ʽ ҽ��ID:������ĿID,....
    Dim str���α䶯 As String '����ɾ�����޸ĵ�ҽ���������ж�RISԤԼ��ҽ���Ƿ��Ѿ���ԤԼ���Ѿ���ԤԼ��ҽ������ɾ
    Dim rsTmp As ADODB.Recordset
    Dim rs��Ѫ As ADODB.Recordset '��Ѫҽ��Ѫ��ӿڵ���ʱ�Ĳ�����Ϣ
    Dim blnѪ�� As Boolean '�Ƿ���Ҫ����Ѫ��ӿڵ���
    Dim strAdvices��Ѫ As String
    Dim strErr As String
    Dim bln��Ѫ��� As Boolean
    
    If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
        blnRIS = True
    ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True And gbln����Ӱ����ϢϵͳԤԼ = True Then
        MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�滻����ʵ��ҽ��ID���������ҽ����Ӧ
    Call MakeRealID
    
    'Pass�Զ���ҩ���
    If mblnPass Then
        If gobjPass.HaveRecipNo() Then
            Call UpdateRecipeNo '��������
            blnRecipeNo = True
        End If
        If gobjPass.zlPassCheck(mobjPassMap) Then
            If Not gobjPass.zlPassAdviceSave(mobjPassMap, mblnNoSave) Then Exit Function
        End If
    End If

    '������ҽӿ�
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        If zlPluginAdviceSave = False Then Exit Function
    End If

    Screen.MousePointer = 11
    
    If gblnѪ��ϵͳ Then blnѪ�� = InitObjBlood()
    Set rs��Ѫ = New ADODB.Recordset
    With rs��Ѫ
        .Fields.Append "ҽ��ID", adBigInt
        .Fields.Append "����", adInteger '0���¿���1���޸ģ�2��ɾ��
        .Fields.Append "״̬", adInteger '0-����Ѫ��ӿ�,1-������Ѫ��ӿ�(��Ѫҽ�����)
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    If mstrDel��Ѫ <> "" Then
        arrDelID = Split(mstrDel��Ѫ, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                Call rs��Ѫ.AddNew(Array("ҽ��ID", "����"), Array(Val(arrDelID(i)), 2))
            End If
        Next
    End If
    
    If Not (mclsMipModule Is Nothing) Then
        blnMsgOk = mclsMipModule.IsConnect
    Else
        blnMsgOk = False
    End If
    
    '����SQL
    arrSQL = Array()
    curDate = zlDatabase.Currentdate

        '�������ϵͳ
    If InitObjRecipeAudit(p����ҽ���´�) Then
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 2 Then
                If vsAdvice.TextMatrix(i, COL_���) = "5" Or vsAdvice.TextMatrix(i, COL_���) = "6" Or vsAdvice.TextMatrix(i, COL_���) = "7" Then
                    If vsAdvice.TextMatrix(i, COL_�������״̬) = "1" Or vsAdvice.TextMatrix(i, COL_�������״̬) = "2" Then
                        If InStr("," & mstrAduitDelIDs & ",", "," & vsAdvice.TextMatrix(i, COL_���ID) & ",") = 0 And vsAdvice.TextMatrix(i, COL_���ID) <> "" Then
                            mstrAduitDelIDs = mstrAduitDelIDs & "," & vsAdvice.TextMatrix(i, COL_���ID)
                        End If
                    End If
                    If vsAdvice.TextMatrix(i, COL_�������״̬) <> "" Then
                        If InStr("," & str������ҩIDs & ",", "," & vsAdvice.TextMatrix(i, COL_���ID) & ",") = 0 And vsAdvice.TextMatrix(i, COL_���ID) <> "" Then
                            str������ҩIDs = str������ҩIDs & "," & vsAdvice.TextMatrix(i, COL_���ID)
                        End If
                    End If
                End If
            End If
        Next
        mstrAduitDelIDs = Mid(mstrAduitDelIDs, 2)
        For i = 0 To UBound(Split(mstrAduitDelIDs, ","))
            If InStr("," & str������ҩIDs & ",", "," & Split(mstrAduitDelIDs, ",")(i) & ",") = 0 Then
                str������ҩIDs = str������ҩIDs & "," & Split(mstrAduitDelIDs, ",")(i)
            End If
        Next
        str������ҩIDs = Mid(str������ҩIDs, 2)
        '�����һ������ɾ�����������ɾ��
        If mstrAduitDelIDs <> "" Then
            For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
                If InStr("," & mstrAduitDelIDs & ",", "," & vsAdvice.RowData(i) & ",") > 0 Then
                    lng���ID = vsAdvice.RowData(i)
                    vsAdvice.RowData(i) = zlDatabase.GetNextID("����ҽ����¼")
                    vsAdvice.TextMatrix(i, COL_EDIT) = 1
                    For j = i - 1 To vsAdvice.FixedRows Step -1
                        If Val(vsAdvice.TextMatrix(j, COL_���ID)) <> lng���ID Then Exit For
                        vsAdvice.TextMatrix(j, COL_���ID) = vsAdvice.RowData(i)
                        vsAdvice.RowData(j) = zlDatabase.GetNextID("����ҽ����¼")
                        vsAdvice.TextMatrix(j, COL_EDIT) = 1
                    Next
                End If
            Next

            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ����¼_�������ɾ��('" & mstrAduitDelIDs & "')"
        End If
        If str������ҩIDs <> "" Then
            '�����˴���������ݵĶ�Ҫ����
            Call gobjRecipeAudit.CancelData(str������ҩIDs, "")
        End If
    End If
    
    'ɾ���˵ļ�¼
    arrDelID = Split(mstrDelIDs, ",")
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & Val(arrDelID(i)) & ")"
            If blnRIS Then str���α䶯 = str���α䶯 & "," & Val(arrDelID(i))
        End If
    Next

    '�༭��־��0-ԭʼ��,1-������,2-�޸�������,3-�޸������
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then    '����ҽ����¼
                '����ת��
                dbl���� = 0
                If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    If Val(.TextMatrix(i, COL_����)) <> 0 Then
                        If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                            '��ҩת�������۵�λ
                            dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_�����װ)), "0.00000")
                        Else
                            '��ҩ�䷽�������ҩ��������,��ת��
                            dbl���� = Val(.TextMatrix(i, COL_����))
                        End If
                    End If
                End If

                '33629
                '����ҩƷ�͵�һ�ྫ��ҩƷ������Ӧ�������������֤����ţ����������������֤����š�
                If blnChecked = False And mblnAddAgent Then
                    If .RowHidden(i) = False And AgentInfo.���ξ�����¼�� = False Then
                        If Val(.TextMatrix(i, COL_״̬)) = 1 And InStr(",����ҩ,����ҩ,����I��,", "," & Trim(.TextMatrix(i, COL_�������)) & ",") > 0 Then
                            blnChecked = frmAgentInfo.ShowMe(Me, mlng����ID, mlng�Һ�ID, mstr����, mstr���֤��, AgentInfo.����������, AgentInfo.���������֤��)
                            If blnChecked Then
                                Call GetAgentInfo
                            Else
                                Screen.MousePointer = 0
                                Exit Function
                            End If
                        End If
                    End If
                End If
                
                If .TextMatrix(i, COL_���) = "K" And blnѪ�� Then
                    If Val(.TextMatrix(i, COL_EDIT)) = 2 Then
                        intBloodState = 0
                        '��Ѫҽ���Ѿ���Ѫ�����޸�ҽ�����״̬Ϊ2
                        If Val(.TextMatrix(i, COL_��鷽��)) = 1 Then
                            If gobjPublicBlood.GetPrepareBloodRs(Val(.RowData(i)), rsBlood) = True Then
                                If Val(rsBlood!��¼���� & "") = 2 And Val(rsBlood!��¼״̬ & "") = 1 Then
                                    .TextMatrix(i, COL_���״̬) = 2
                                    intBloodState = 1
                                    bln��Ѫ��� = True
                                End If
                            End If
                        End If
                        Call rs��Ѫ.AddNew(Array("ҽ��ID", "����", "״̬"), Array(Val(.RowData(i)), 1, intBloodState))
                    ElseIf Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        Call rs��Ѫ.AddNew(Array("ҽ��ID", "����"), Array(Val(.RowData(i)), 0))
                    End If
                End If
                
                If Val(.TextMatrix(i, COL_EDIT)) = 3 Then    '�޸�����ŵļ�¼
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & .RowData(i) & "," & Val(.TextMatrix(i, COL_���)) & ")"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 2 Then    '�޸������ݵļ�¼
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Update(" & _
                                             .RowData(i) & "," & ZVal(.TextMatrix(i, COL_���ID)) & "," & _
                                             Val(.TextMatrix(i, COL_���)) & "," & Val(.TextMatrix(i, COL_״̬)) & ",1," & _
                                             Val(.TextMatrix(i, COL_������ĿID)) & "," & ZVal(.TextMatrix(i, COL_�շ�ϸĿID)) & "," & _
                                             ZVal(.TextMatrix(i, COL_����)) & "," & ZVal(.TextMatrix(i, COL_����)) & "," & ZVal(dbl����) & "," & _
                                             "'" & Replace(.TextMatrix(i, col_ҽ������), "'", "''") & "','" & Replace(.TextMatrix(i, COL_ҽ������), "'", "''") & "'," & _
                                             "'" & .TextMatrix(i, COL_�걾��λ) & "','" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                                             ZVal(.TextMatrix(i, COL_Ƶ�ʴ���)) & "," & ZVal(.TextMatrix(i, COL_Ƶ�ʼ��)) & "," & _
                                             "'" & .TextMatrix(i, COL_�����λ) & "','" & .TextMatrix(i, COL_ִ��ʱ��) & "'," & _
                                             Val(.TextMatrix(i, COL_�Ƽ�����)) & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & _
                                             Val(.TextMatrix(i, COL_ִ������)) & "," & Val(.TextMatrix(i, COL_��־)) & "," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
                                             mlng���˿���id & "," & Val(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                             "'" & .TextMatrix(i, COL_��鷽��) & "'," & Val(.TextMatrix(i, COL_ִ�б��)) & "," & _
                                             "NULL,'" & .Cell(flexcpData, i, COL_ҽ������) & "','" & UserInfo.���� & "'," & ZVal(.TextMatrix(i, COL_��Ѽ���)) & "," & _
                                             ZVal(Val(.TextMatrix(i, COL_��ҩĿ��))) & ",'" & .TextMatrix(i, COL_��ҩ����) & "'," & ZVal(Val(.TextMatrix(i, COL_���״̬))) & ",'" & .TextMatrix(i, COL_����˵��) & "'" & _
                                             ",'',Null," & ZVal(Val(.TextMatrix(i, COL_�����ĿID))) & ",NULL," & IIF(blnRecipeNo, Val(.TextMatrix(i, COL_�������)), "NULL") & ")"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 1 Then    '�����ļ�¼
                
                    If .TextMatrix(i, COL_���) = "K" Then
                        '��Ѫ��ҽ�����»�ȡҽ�����״̬
                        If TypeName(.Cell(flexcpData, i, COL_�������)) = "Recordset" Then
                            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, i, COL_�������))
                            If Not rsCard.EOF Then
                                rsCard.MoveFirst
                                strTmp = Nvl(rsCard!������Ŀ & "", .TextMatrix(i, COL_������ĿID) & "," & dbl����)
                                .TextMatrix(i, COL_���״̬) = GetBloodVerifyState(1, mlng����ID, mlng�Һ�ID, .TextMatrix(i, COL_�걾��λ), GetBloodTotalByML(strTmp), Val(.TextMatrix(i, COL_��־)), Val(.TextMatrix(i, COL_��鷽��)), Val(.TextMatrix(i, COL_Ӥ��)), .RowData(i), strTmp)
                                If i < .Rows - 1 Then
                                    If Val(.TextMatrix(i + 1, COL_���ID)) = .RowData(i) Then
                                        .TextMatrix(i + 1, COL_���״̬) = .TextMatrix(i, COL_���״̬)
                                    End If
                                End If
                                Call SetRow��־ͼ��(i, 2)
                            End If
                        Else
                            '�����뵥�´��ҽ��
                            If blnѪ�� Then
                                strTmp = .TextMatrix(i, COL_������ĿID) & "," & dbl����
                                .TextMatrix(i, COL_���״̬) = GetBloodVerifyState(1, mlng����ID, mlng�Һ�ID, .TextMatrix(i, COL_�걾��λ), GetBloodTotalByML(strTmp), Val(.TextMatrix(i, COL_��־)), Val(.TextMatrix(i, COL_��鷽��)), Val(.TextMatrix(i, COL_Ӥ��)), .RowData(i), strTmp)
                                If i < .Rows - 1 Then
                                    If Val(.TextMatrix(i + 1, COL_���ID)) = .RowData(i) Then
                                        .TextMatrix(i + 1, COL_���״̬) = .TextMatrix(i, COL_���״̬)
                                    End If
                                End If
                                Call SetRow��־ͼ��(i, 2)
                            End If
                        End If
                    End If
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & _
                                             .RowData(i) & "," & ZVal(.TextMatrix(i, COL_���ID)) & "," & _
                                             Val(.TextMatrix(i, COL_���)) & ",1," & mlng����ID & ",NULL," & _
                                             Val(.TextMatrix(i, COL_Ӥ��)) & "," & Val(.TextMatrix(i, COL_״̬)) & ",1," & _
                                             "'" & IIF(.TextMatrix(i, COL_���) = "*", "", .TextMatrix(i, COL_���)) & "'," & Val(.TextMatrix(i, COL_������ĿID)) & "," & _
                                             ZVal(.TextMatrix(i, COL_�շ�ϸĿID)) & "," & _
                                             ZVal(.TextMatrix(i, COL_����)) & "," & ZVal(.TextMatrix(i, COL_����)) & "," & ZVal(dbl����) & "," & _
                                             "'" & Replace(.TextMatrix(i, col_ҽ������), "'", "''") & "','" & Replace(.TextMatrix(i, COL_ҽ������), "'", "''") & "'," & _
                                             "'" & .TextMatrix(i, COL_�걾��λ) & "','" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                                             ZVal(.TextMatrix(i, COL_Ƶ�ʴ���)) & "," & ZVal(.TextMatrix(i, COL_Ƶ�ʼ��)) & "," & _
                                             "'" & .TextMatrix(i, COL_�����λ) & "','" & .TextMatrix(i, COL_ִ��ʱ��) & "'," & _
                                             Val(.TextMatrix(i, COL_�Ƽ�����)) & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & _
                                             Val(.TextMatrix(i, COL_ִ������)) & "," & Val(.TextMatrix(i, COL_��־)) & "," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
                                             mlng���˿���id & "," & Val(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                             "'" & mstr�Һŵ� & "'," & ZVal(mlngǰ��ID) & ",'" & .TextMatrix(i, COL_��鷽��) & "'," & _
                                             Val(.TextMatrix(i, COL_ִ�б��)) & ",NULL,'" & .Cell(flexcpData, i, COL_ҽ������) & "','" & UserInfo.���� & "'," & ZVal(.TextMatrix(i, COL_��Ѽ���)) & "," & _
                                             ZVal(Val(.TextMatrix(i, COL_��ҩĿ��))) & ",'" & .TextMatrix(i, COL_��ҩ����) & "'," & ZVal(Val(.TextMatrix(i, COL_���״̬))) & "," & ZVal(Val(.TextMatrix(i, COL_�������))) & ",'" & .TextMatrix(i, COL_����˵��) & _
                                             "',Null," & ZVal(Val(.TextMatrix(i, COL_�䷽ID))) & ",Null," & ZVal(Val(.TextMatrix(i, COL_�����ĿID))) & ",NULL," & IIF(blnRecipeNo, Val(.TextMatrix(i, COL_�������)), "NULL") & ")"
                    
                    
                    '�ռ�����������ҽ�����ڹ�������Σ��ֵ��¼
                    If mlngΣ��ֵID <> 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_����Σ��ֵҽ��_Update(1," & mlngΣ��ֵID & "," & .RowData(i) & ")"
                        End If
                    End If
                    
                    '��Ѫ���뵥�Ķ������ݱ���
                    If .TextMatrix(i, COL_���) = "K" Then
                        If TypeName(.Cell(flexcpData, i, COL_�������)) = "Recordset" Then
                            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, i, COL_�������))
                            If Not rsCard.EOF Then
                                rsCard.MoveFirst
                                strTmp = rsCard!����������ĿSQL & ""
                                If strTmp <> "" Then
                                    strTmp = Replace(strTmp, "[���ID]", .RowData(i))
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = strTmp
                                End If
                                strTmp = rsCard!������ĿSQL & ""
                                If strTmp <> "" Then
                                    strTmp = Replace(strTmp, "[���ID]", .RowData(i))
                                    varTmp = Split(strTmp, "<splitSQL>")
                                    For j = 0 To UBound(varTmp)
                                        If varTmp(j) <> "" Then
                                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                            arrSQL(UBound(arrSQL)) = varTmp(j)
                                        End If
                                    Next
                                End If
                                strTmp = rsCard!��Ϲ�����ϢSQL & ""
                                If strTmp <> "" Then
                                    strTmp = Replace(strTmp, "[���ID]", .RowData(i))
                                    varTmp = Split(strTmp, "<splitSQL>")
                                    For j = 0 To UBound(varTmp)
                                        If j = 1 Then
                                            If varTmp(j) <> "" Then
                                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                                arrSQL(UBound(arrSQL)) = varTmp(j)
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
                
                If blnRIS Then
                    If InStr(",F,D,", .TextMatrix(i, COL_���)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_��������))) > 0 And .TextMatrix(i, COL_���) = "E" Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            If Val(.TextMatrix(i, COL_EDIT)) = 2 Then
                                str���α䶯 = str���α䶯 & "," & .RowData(i)
                            End If
                            If InStr(",1,2,", Val(.TextMatrix(i, COL_EDIT))) > 0 And Val(.TextMatrix(i, COL_��־)) <> 1 Then
                                strNewAdvice = strNewAdvice & "," & .RowData(i) & ":" & Val(.TextMatrix(i, COL_������ĿID))
                            End If
                        End If
                    End If
                End If
                
                '������Ե�
                If .Cell(flexcpData, i, COL_����) & "" <> "" & .TextMatrix(i, COL_����) Then
                    If .TextMatrix(i, COL_����) = "1" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Ƥ��(" & Val(.RowData(i)) & ",'����',NULL)"
                    ElseIf .TextMatrix(i, COL_����) = "0" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Ƥ��(" & Val(.RowData(i)) & ",'',NULL)"
                    End If
                End If
                '�������븽��
                If .TextMatrix(i, COL_����) <> "" And .Cell(flexcpData, i, COL_����) = 1 Then
                    arrAppend = Split(.TextMatrix(i, COL_����), "<Split1>")
                    For j = 0 To UBound(arrAppend)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & .RowData(i) & "," & _
                                                 "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                                                 j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                                                 IIF(j = 0, ",1", "") & ")"
                    Next
                End If

                'Pass:���������
                If Val(.Cell(flexcpData, i, COL_���)) = 1 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & .RowData(i) & "," & _
                                             IIF(CStr(.Cell(flexcpData, i, COL_��ʾ)) = "", "NULL", Val(.Cell(flexcpData, i, COL_��ʾ))) & ")"
                End If
            End If
        Next
    End With

    '��ϲ�������(��Ҫ����ҽ��ID����)
    If lbl���.Tag = "1" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3,Null,'1,11')"
        If blnMsgOk Then Call InitRsMsg(rsMsg)
        With vsDiag
            j = 0
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, col���)) <> "" Then
                    blnDiagChange = True
                    If Val(.Cell(flexcpData, i, col���ID)) > 0 And Not mrsDiag Is Nothing Then
                        strFilter = "�������=" & IIF(.Cell(flexcpData, i, col��ҽ) = 1, "11", "1") & " And ��¼��Դ=3 And ����id=" & ZVal(.TextMatrix(i, col����ID)) & " And ���id=" & ZVal(.TextMatrix(i, col���ID))
                        strTmp = IIF(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & IIF(.TextMatrix(i, col��ҽ֤��) <> "", "(" & .TextMatrix(i, col��ҽ֤��) & ")", "")
                        strFilter = strFilter & " And �������= '" & strTmp & "'"
                        If IsDate(.TextMatrix(i, col����ʱ��)) Then
                            strFilter = strFilter & " And  ����ʱ��= '" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "'"
                        Else
                            strFilter = strFilter & " And  ����ʱ��= Null "
                        End If

                        strFilter = strFilter & " And  ֤��ID= " & ZVal(.TextMatrix(i, col֤��ID))

                        strFilter = strFilter & " And �Ƿ�����=" & Val(.Cell(flexcpData, i, col����))
                        mrsDiag.Filter = strFilter
                        blnDiagChange = mrsDiag.EOF
                    End If
                    lng��ϼ�¼id = zlDatabase.GetNextID("������ϼ�¼")
                    strTmp = .TextMatrix(i, colҽ��ID)
                    If Len(strTmp) > 4000 Then
                        strTmp = ""
                    End If
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): j = j + 1
                    If blnDiagChange Then
                        str��¼�� = UserInfo.����
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                                                 "Null," & IIF(.Cell(flexcpData, i, col��ҽ) = 1, "11", "1") & "," & ZVal(.TextMatrix(i, col����ID)) & "," & _
                                                 ZVal(.TextMatrix(i, col���ID)) & "," & ZVal(.TextMatrix(i, col֤��ID)) & "," & _
                                                 "'" & IIF(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & IIF(.TextMatrix(i, col��ҽ֤��) <> "", "(" & .TextMatrix(i, col��ҽ֤��) & ")", "") & "',Null,Null," & Val(.Cell(flexcpData, i, col����)) & "," & _
                                                 "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                                 IIF(strTmp = "", "null", "'" & strTmp & "'") & "," & j & ",Null,Null,To_date('" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & UserInfo.���� & "'," & lng��ϼ�¼id & ")"
                    Else
                        str��¼�� = mrsDiag!��¼��
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                                                 "Null," & IIF(.Cell(flexcpData, i, col��ҽ) = 1, "11", "1") & "," & ZVal(.TextMatrix(i, col����ID)) & "," & _
                                                 ZVal(.TextMatrix(i, col���ID)) & "," & ZVal(.TextMatrix(i, col֤��ID)) & "," & _
                                                 "'" & IIF(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & IIF(.TextMatrix(i, col��ҽ֤��) <> "", "(" & .TextMatrix(i, col��ҽ֤��) & ")", "") & "',Null,Null," & Val(.Cell(flexcpData, i, col����)) & "," & _
                                                 "To_Date('" & Format(CDate(mrsDiag!��¼����), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                                 IIF(strTmp = "", "null", "'" & strTmp & "'") & "," & j & ",Null,Null,To_date('" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & mrsDiag!��¼�� & "'," & lng��ϼ�¼id & ")"
                    End If
                    strTmp = .TextMatrix(i, colҽ��ID)
                    If Len(strTmp) > 4000 Then Call Make���ҽ����Ӧ(arrSQL, strTmp, lng��ϼ�¼id)
                    If blnMsgOk Then
                        Call SendMsg���(i, j, lng��ϼ�¼id, str��¼��, rsMsg)
                    End If
                End If
            Next
        End With
    End If
    
    If blnѪ�� Then
        blnѪ�� = rs��Ѫ.RecordCount > 0
        If blnѪ�� Then rs��Ѫ.MoveFirst
    End If
    
    If blnRIS Then
        If str���α䶯 <> "" Then
            str���α䶯 = Mid(str���α䶯, 2)
            Set rsTmp = GetDataRISԤԼ(str���α䶯)
            If rsTmp.RecordCount > 0 Then
                str���α䶯 = "������"
            Else
                str���α䶯 = ""
            End If
        End If
    Else
        str���α䶯 = ""
    End If
    '����RIS��ȡ��ԤԼ
    If str���α䶯 <> "" Then
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!ԤԼid & "")) Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ����β���ɾ�����޸����Ѿ�ԤԼҽ����������Ӱ����Ϣϵͳ�ӿ�(HISSchedulingEx)ȡ��ϢԤԼδ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            rsTmp.MoveNext
        Next
        err.Clear: On Error GoTo 0
    End If
    '�ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    '����Ѫ��Ľӿ�
    If blnѪ�� Then
        For i = 1 To rs��Ѫ.RecordCount
            If Val("" & rs��Ѫ!״̬) <> 1 Then
                If gobjPublicBlood.AdviceOperation(p����ҽ���´�, Val(rs��Ѫ!ҽ��ID & ""), Val(rs��Ѫ!���� & ""), , strErr) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False
                    Screen.MousePointer = 0
                    MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            rs��Ѫ.MoveNext
        Next
    End If
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    mlngID���� = 0
    'Pass�Զ���ҩ����ϴ�����
    If mblnPass Then
        Call gobjPass.zlPassUpLoad(mobjPassMap)
    End If
    
    If blnMsgOk Then
        If Not (mrs��� Is Nothing) Then
            mrs���.Filter = "״̬ = 0"
            If Not mrs���.EOF Then
                For i = 1 To mrs���.RecordCount
                    Call ZLHIS_CIS_011(mclsMipModule, mlng����ID, mstr����, 1, mlng�Һ�ID, mlng���˿���id, mrs���!ID, mrs���!��ϱ���, mrs���!��������)
                    mrs���.Delete
                    mrs���.MoveNext
                Next
            End If
            mrs���.Filter = "״̬ <> 0"
            If Not mrs���.EOF Then
                For i = 1 To mrs���.RecordCount
                    mrs���!״̬ = 0
                    mrs���.MoveNext
                Next
            End If
            mrs���.Filter = 0
        End If
        If Not (rsMsg Is Nothing) Then
            rsMsg.Filter = "״̬ = 1"
            If Not rsMsg.EOF Then
                For i = 1 To rsMsg.RecordCount
                    Call ZLHIS_CIS_010(mclsMipModule, mlng����ID, mstr����, 1, mlng�Һ�ID, mlng���˿���id, _
                        rsMsg!���id, rsMsg!�������, rsMsg!�Ƿ�����, rsMsg!��ϴ���, rsMsg!��ϱ���, rsMsg!��������, rsMsg!��������, rsMsg!�������, rsMsg!֤�����, rsMsg!֤������, rsMsg!��¼����, rsMsg!��¼��Ա)
                    rsMsg.MoveNext
                Next
            End If
        End If
    End If
    
    If bln��Ѫ��� Then Call ReadMsg
    
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    '����ɾ������ҽӿ�
    On Error Resume Next
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            If Not gobjPlugIn Is Nothing Then
                Call gobjPlugIn.AdviceDeleted(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(arrDelID(i)), mint����)
                Call zlPlugInErrH(err, "AdviceDeleted")
            End If
        End If
    Next
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0

    '���ô������ϵͳ���
    If InitObjRecipeAudit(p����ҽ���´�) And mblnNoSave Then
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If vsAdvice.TextMatrix(i, COL_���) = "E" And vsAdvice.TextMatrix(i, COL_��������) = "2" Then
                '��ǰ�¿�״̬��ҽ������Ҫ����
                If Val(vsAdvice.TextMatrix(i, COL_״̬)) = 1 Then
                    str��ҩIDs = str��ҩIDs & "," & vsAdvice.RowData(i)
                End If
            End If
        Next
        If Mid(str��ҩIDs, 2) <> "" Then
            Call gobjRecipeAudit.AutoAudit(Me, 1, Mid(str��ҩIDs, 2), mlng���˿���id, 0, mlng����ID, mlng�Һ�ID)
        End If
    End If

    '����ɹ���,���м�¼���ԭʼ��¼
    With vsAdvice
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If .RowData(i) <> 0 Then
                .TextMatrix(i, COL_EDIT) = 0
                .Cell(flexcpData, i, COL_���) = Empty    'Pass:����������־
                .Cell(flexcpData, i, COL_����) = 0    '����:����������־
            End If
        Next
    End With

    '��������½�����(���翪ʼʱ�䲻׼����)
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    Screen.MousePointer = 0
    lbl���.Tag = ""
    mblnNoSave = False
    mstrDelIDs = ""
    mstrAduitDelIDs = ""
    mstrDel��Ѫ = ""
    SaveAdvice = True
    mblnOK = True
    mlngΣ��ֵID = 0
    
    'ҽ�����ݱ����ύ����RISԤԼ��ֻ�����ܲ��˲�ԤԼ
    If blnRIS And mbytPatiType = 1 Then
        If strNewAdvice <> "" Then
            strTmp = Mid(strNewAdvice, 2)
            varTmp = Split(strTmp, ",")
            On Error Resume Next
            For i = 0 To UBound(varTmp)
                strTmp = varTmp(i)
                Call gobjRis.HISScheduling(1, Val(Split(strTmp, ":")(0)), Val(Split(strTmp, ":")(1)))
            Next
            err.Clear: On Error GoTo 0
        End If
    End If
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SendMsg���(ByVal lngRow As Long, ByVal lng���� As Long, ByVal lngID As Long, ByVal str��¼�� As String, ByRef rsMsg As ADODB.Recordset)
    Dim i As Long
    With vsDiag
        rsMsg.AddNew
        rsMsg!���id = lngID
        rsMsg!������� = IIF(.Cell(flexcpData, lngRow, col��ҽ) = 1, "11", "1")
        rsMsg!�Ƿ����� = Val(.Cell(flexcpData, lngRow, col����))
        rsMsg!��ϴ��� = lng����
        rsMsg!��ϱ��� = .TextMatrix(lngRow, col��ϱ���)
        rsMsg!�������� = .TextMatrix(lngRow, col��������)
        rsMsg!�������� = .TextMatrix(lngRow, col��������)
        rsMsg!������� = .TextMatrix(lngRow, col�������)
        rsMsg!֤����� = .TextMatrix(lngRow, col֤�����)
        rsMsg!֤������ = .TextMatrix(lngRow, col��ҽ֤��)
        rsMsg!��¼���� = Format(.TextMatrix(lngRow, col����ʱ��), "yyyy-MM-dd HH:mm:ss")
        rsMsg!��¼��Ա = str��¼��
        rsMsg!״̬ = 1
        
        mrs���.Filter = "��ʾ����='" & .TextMatrix(lngRow, col����) & "'"
        
        If mrs���.EOF Then
            mrs���.AddNew
            mrs���!ID = lngID
            mrs���!��ʾ���� = .TextMatrix(lngRow, col����)
            mrs���!��ϱ��� = .TextMatrix(lngRow, col��ϱ���)
            mrs���!�������� = .TextMatrix(lngRow, col��������)
            mrs���!״̬ = 2
            mrs���.Update
        Else
            rsMsg!״̬ = 0
            mrs���!״̬ = 1
            mrs���!ID = lngID
        End If
        rsMsg.Update
    End With
End Sub

Private Sub InitRsMsg(ByRef rsMsg As ADODB.Recordset)
'���ܣ���ʼ����Ϣ��¼��
    Set rsMsg = New ADODB.Recordset
    rsMsg.Fields.Append "���id", adBigInt
    rsMsg.Fields.Append "�������", adVarChar, 6
    rsMsg.Fields.Append "�Ƿ�����", adVarChar, 6
    rsMsg.Fields.Append "��ϴ���", adBigInt
    rsMsg.Fields.Append "��ϱ���", adVarChar, 60
    rsMsg.Fields.Append "��������", adVarChar, 60
    rsMsg.Fields.Append "��������", adVarChar, 60
    rsMsg.Fields.Append "�������", adVarChar, 60
    rsMsg.Fields.Append "֤�����", adVarChar, 60
    rsMsg.Fields.Append "֤������", adVarChar, 120
    rsMsg.Fields.Append "��¼����", adVarChar, 60
    rsMsg.Fields.Append "��¼��Ա", adVarChar, 120
    rsMsg.Fields.Append "״̬", adBigInt '1 ������
    rsMsg.CursorLocation = adUseClient
    rsMsg.LockType = adLockOptimistic
    rsMsg.CursorType = adOpenStatic
    rsMsg.Open
End Sub

Private Function LoadAdvice() As Boolean
''���ܣ���ȡ��ǰ���˵�ҽ����¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln�䷽ As Boolean
    Dim i As Long, j As Long, lng���ID As Long

    Screen.MousePointer = 11

    On Error GoTo errH

    '��ҽ��ȱʡ������
    If msng���� = 0 Then msng���� = 1

    '�ٴ���ҽ��������༭
    strSQL = " And Nvl(A.ǰ��ID,0) in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) X)"
    
    '��ȡ"1-�¿�,8-��ֹͣ(�ѷ���)"��ҽ��
    strSQL = _
    " Select A.ID,A.���ID,Nvl(A.Ӥ��,0) as Ӥ��,A.���,A.ҽ����Ч,A.ҽ��״̬,A.�������,A.������ĿID,B.����,A.�걾��λ,A.��鷽��," & _
             " A.ִ�б��,A.�շ�ϸĿID,A.��ʼִ��ʱ��,A.ҽ������,A.ҽ������,A.��������,A.����,A.�ܸ�����,B.���㵥λ,A.ִ��Ƶ��," & _
             " A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,B.���㷽ʽ,B.ִ��Ƶ��,B.��������,B.����Ӧ��,B.ִ�з���,A.�Ƽ�����,A.ִ��ʱ�䷽��,A.ִ������," & _
             " A.ִ�п���ID,A.��������ID,A.����ҽ��,A.����ʱ��,A.������־,Decode(Nvl(c.��������, 0), 0, b.¼������, c.��������) As ��������,C.����ְ��,C.�������,C.������,C.ҩƷ����," & _
             " D.����ϵ��,D.�����װ,D.���ﵥλ,F.���㵥λ as ɢװ��λ,E.��������,D.����ɷ���� As �ɷ����,A.�����," & _
             " Decode(A.�¿�ǩ��ID,NULL,0,1) as ǩ����,A.ժҪ,A.��Ѽ���,A.��ҩĿ��,A.��ҩ����,A.���״̬,A.Ƥ�Խ��,A.����˵��," & _
             " a.�䷽ID,c.�ٴ��Թ�ҩ,d.��ΣҩƷ,a.�����ĿID,b.����ʱ��,C.��ý,a.�������,J.״̬ as �������״̬,J.����� as ���������,A.�������,d.����ҩ��,Nvl(Max(g.Σ��ֵid), Max(h.Σ��ֵid)) As Σ��ֵid" & _
             " From ����ҽ����¼ A,������ĿĿ¼ B,ҩƷ���� C,ҩƷ��� D,�������� E,�շ���ĿĿ¼ F, ���������ϸ I, ��������¼ J, ����Σ��ֵҽ�� G,����Σ��ֵҽ�� H" & _
             " Where Nvl(A.ҽ����Ч,0)=1 And A.������ĿID=B.ID And A.������ĿID=C.ҩ��ID(+)  And a.ID = i.ҽ��ID(+) And I.��ID = J.ID(+) and (I.����ύ =1 Or I.��ID is NULL) And Nvl(A.ִ�б��,0)<>-1 And a.ID = h.ҽ��ID(+) and a.���ID=g.ҽ��ID(+)  " & _
             " And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=E.����ID(+) And E.����ID=F.ID(+) And A.ҽ��״̬ IN(1,8)" & strSQL & _
             " And A.����ID+0=[1] And A.�Һŵ�=[2] And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3"
    strSQL = strSQL & " " & _
            "Group By a.Id, a.���id, a.Ӥ��, a.���, a.ҽ����Ч, a.ҽ��״̬, a.�������, a.������Ŀid, b.����, a.�걾��λ, a.��鷽��, a.ִ�б��, a.�շ�ϸĿid, a.��ʼִ��ʱ��," & vbNewLine & _
            "         a.ҽ������, a.ҽ������, a.��������, a.����, a.�ܸ�����, b.���㵥λ, a.ִ��Ƶ��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, b.���㷽ʽ, b.ִ��Ƶ��, b.��������,b.����Ӧ��, b.ִ�з���," & vbNewLine & _
            "         a.�Ƽ�����, a.ִ��ʱ�䷽��, a.ִ������, a.ִ�п���id, a.��������id, a.����ҽ��, a.����ʱ��, a.������־, c.��������, c.����ְ��, c.�������, c.������, c.ҩƷ����," & vbNewLine & _
            "         d.����ϵ��, d.�����װ, d.���ﵥλ, f.���㵥λ, e.��������, d.����ɷ����, a.�����, a.�¿�ǩ��id, a.ժҪ, a.��Ѽ���, a.��ҩĿ��, a.��ҩ����, a.���״̬," & vbNewLine & _
            "         a.Ƥ�Խ��, a.����˵��, a.�䷽id, c.�ٴ��Թ�ҩ, d.��ΣҩƷ, a.�����Ŀid, b.����ʱ��, c.��ý, a.�������, j.״̬, j.�����,d.����ҩ��, a.�������,b.¼������" & vbNewLine & _
            "Order By a.Ӥ��, a.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, IIF(mstrǰ��IDs = "", "0", mstrǰ��IDs))
    On Error GoTo 0

    If Not rsTmp.EOF Then
        mblnRowChange = False
        With vsAdvice
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                bln�䷽ = False

                .RowData(i) = CLng(rsTmp!ID)
                .TextMatrix(i, COL_EDIT) = 0    'ԭʼ��¼
                .TextMatrix(i, COL_���ID) = Nvl(rsTmp!���ID)
                .TextMatrix(i, COL_Ӥ��) = Nvl(rsTmp!Ӥ��, 0)
                .TextMatrix(i, COL_���) = rsTmp!���
                .TextMatrix(i, COL_״̬) = Nvl(rsTmp!ҽ��״̬, 0)

                .TextMatrix(i, COL_���) = rsTmp!�������
                .TextMatrix(i, COL_������ĿID) = rsTmp!������ĿID
                .TextMatrix(i, COL_����) = rsTmp!����
                .TextMatrix(i, COL_�걾��λ) = Nvl(rsTmp!�걾��λ)
                .TextMatrix(i, COL_��鷽��) = Nvl(rsTmp!��鷽��)
                .TextMatrix(i, COL_ִ�б��) = Nvl(rsTmp!ִ�б��, 0)
                .TextMatrix(i, COL_�շ�ϸĿID) = Nvl(rsTmp!�շ�ϸĿID)
                .TextMatrix(i, col_ҽ������) = Nvl(rsTmp!ҽ������)
                .TextMatrix(i, COL_ҽ������) = Nvl(rsTmp!ҽ������)
                .Cell(flexcpData, i, COL_ҽ������) = CStr(Nvl(rsTmp!ժҪ))

                .TextMatrix(i, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
                .TextMatrix(i, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
                .TextMatrix(i, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
                .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, COL_����Ӧ��) = Nvl(rsTmp!����Ӧ��)
                .TextMatrix(i, COL_ִ�з���) = Nvl(rsTmp!ִ�з���, 0)
                .TextMatrix(i, COL_�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, COL_�����ȼ�) = Val("" & rsTmp!������)
                .TextMatrix(i, COL_ҩƷ����) = Nvl(rsTmp!ҩƷ����)
                .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, COL_����ְ��) = Nvl(rsTmp!����ְ��)
                .TextMatrix(i, COL_����ҩ��) = rsTmp!����ҩ�� & ""
                If gblnKSSStrict Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ Then
                    .TextMatrix(i, COL_���״̬) = Val("" & rsTmp!���״̬)
                End If
                .TextMatrix(i, COL_��ҩĿ��) = "" & rsTmp!��ҩĿ��
                .TextMatrix(i, COL_��ҩ����) = "" & rsTmp!��ҩ����
                .TextMatrix(i, COL_�䷽ID) = Nvl(rsTmp!�䷽ID)
                .TextMatrix(i, COL_�ٴ��Թ�ҩ) = rsTmp!�ٴ��Թ�ҩ & ""
                .TextMatrix(i, COL_��ΣҩƷ) = Nvl(rsTmp!��ΣҩƷ, 0)
                .TextMatrix(i, COL_�����ĿID) = "" & rsTmp!�����ĿID
                If Format(Nvl(rsTmp!����ʱ��, "3000/1/1"), "yyyy-MM-dd") <> "3000-01-01" Then
                    .TextMatrix(i, COL_�Ƿ�ͣ��) = 1
                End If
                If InStr(",5,6,7,", .TextMatrix(i, COL_���)) > 0 Then
                    .TextMatrix(i, COL_����ϵ��) = Nvl(rsTmp!����ϵ��)
                    .TextMatrix(i, COL_�����װ) = Nvl(rsTmp!�����װ)
                    .TextMatrix(i, COL_���ﵥλ) = Nvl(rsTmp!���ﵥλ)
                    If Not IsNull(rsTmp!����ϵ��) Then
                        .TextMatrix(i, COL_�ɷ����) = Nvl(rsTmp!�ɷ����, 0)
                    End If
                ElseIf .TextMatrix(i, COL_���) = "4" Then
                    .TextMatrix(i, COL_����ϵ��) = 1
                    .TextMatrix(i, COL_�����װ) = 1
                    .TextMatrix(i, COL_���ﵥλ) = Nvl(rsTmp!ɢװ��λ)
                    .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������, 0)
                End If

                .TextMatrix(i, COL_��ʼʱ��) = Format(rsTmp!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COL_��ʼʱ��) = Format(rsTmp!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm")

                .TextMatrix(i, COL_Ƶ��) = Nvl(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, COL_Ƶ�ʴ���) = Nvl(rsTmp!Ƶ�ʴ���)
                .TextMatrix(i, COL_Ƶ�ʼ��) = Nvl(rsTmp!Ƶ�ʼ��)
                .TextMatrix(i, COL_�����λ) = Nvl(rsTmp!�����λ)
                .TextMatrix(i, COL_ִ��ʱ��) = Nvl(rsTmp!ִ��ʱ�䷽��)

                .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID)
                .TextMatrix(i, COL_ִ������) = Nvl(rsTmp!ִ������, 0)
                .TextMatrix(i, COL_����) = IIF(Nvl(rsTmp!Ƥ�Խ��, "") = "����", "1", "0")
                .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����)
                .TextMatrix(i, COL_�Ƿ���ý) = Val(rsTmp!��ý & "")
                .TextMatrix(i, COL_�������) = rsTmp!������� & ""
                .TextMatrix(i, COL_�������״̬) = rsTmp!�������״̬ & ""
                .TextMatrix(i, COL_���������) = rsTmp!��������� & ""
                .TextMatrix(i, COL_�������) = Val("" & rsTmp!�������)
                .TextMatrix(i, COL_Σ��ֵID) = Val("" & rsTmp!Σ��ֵID)
                
                If rsTmp!������� = "E" Then
                    If Nvl(rsTmp!���ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_���ID)) = rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼�ǳ�ҩ�ĸ�ҩ;��,������һ����ҩ��
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = rsTmp!ID Then
                                    '��ʾ��ҩ;��
                                    .TextMatrix(j, COL_�÷�) = rsTmp!���� & rsTmp!ҽ������
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼����ҩ�䷽���÷�,���䷽��ʾ��
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                            .TextMatrix(i, COL_��ҩ��̬) = Val("" & rsTmp!��鷽��)    '��ҩ�÷��еļ�鷽���ֶδ洢����ҩ��̬
                            bln�䷽ = True
                        ElseIf .TextMatrix(i - 1, COL_���) = "C" Then
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                        End If
                    ElseIf Not IsNull(rsTmp!���ID) And .TextMatrix(i - 1, COL_���) = "K" And Nvl(rsTmp!���ID, 0) = .RowData(i - 1) Then
                        '��ǰ��¼����Ѫ;����
                        .TextMatrix(i - 1, COL_�÷�) = rsTmp!����
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        '��ǰ��¼����ҩ�䷽�巨��
                        bln�䷽ = True
                    End If
                ElseIf rsTmp!������� = "7" Then
                    bln�䷽ = True
                End If

                '����
                .TextMatrix(i, COL_����) = FormatEx(Nvl(rsTmp!��������), 5)
                If .TextMatrix(i, COL_���) = "4" Then
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!ɢװ��λ)
                ElseIf InStr(",5,6,7,", rsTmp!�������) > 0 _
                       Or (Val(.TextMatrix(i, COL_Ƶ������)) = 0 And InStr(",1,2,", Nvl(rsTmp!���㷽ʽ, 0)) > 0) Then
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                End If

                '����
                .TextMatrix(i, COL_����) = Nvl(rsTmp!����)
                'ȡ����¿�ҽ���Ŀ�����Ϊȱʡ����
                If InStr(",1,2,", Nvl(rsTmp!ҽ��״̬, 0)) > 0 _
                   And InStr(",5,6,", rsTmp!�������) > 0 And Nvl(rsTmp!����, 0) <> 0 Then
                    msng���� = Nvl(rsTmp!����, 1)
                End If

                '����
                If InStr(",5,6,", rsTmp!�������) > 0 Then
                    '��ҩ����������,�����۵�λ���,���ﵥλ��ʾ
                    If Not IsNull(rsTmp!�ܸ�����) And Not IsNull(rsTmp!�����װ) Then
                        .TextMatrix(i, COL_����) = FormatEx(rsTmp!�ܸ����� / rsTmp!�����װ, 5)
                    End If
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���ﵥλ)

                    Call Set��ҩ�����Ƿ���(i)

                ElseIf bln�䷽ Then
                    If Not IsNull(rsTmp!�ܸ�����) Then .TextMatrix(i, COL_����) = rsTmp!�ܸ�����
                    .TextMatrix(i, COL_������λ) = "��"    '��ҩ�䷽������λΪ"��"
                    Call Set��ҩ�����Ƿ���(i)
                Else
                    '�����������ҩ����������
                    If Not IsNull(rsTmp!�ܸ�����) Then .TextMatrix(i, COL_����) = rsTmp!�ܸ�����

                    If .TextMatrix(i, COL_���) = "4" Then
                        .TextMatrix(i, COL_������λ) = Nvl(rsTmp!ɢװ��λ)
                    Else
                        .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                    End If
                End If

                .TextMatrix(i, COL_����˵��) = rsTmp!����˵�� & ""

                .TextMatrix(i, COL_��������ID) = rsTmp!��������id
                .TextMatrix(i, COL_����ҽ��) = rsTmp!����ҽ��

                .TextMatrix(i, COL_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COL_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")

                .TextMatrix(i, COL_��Ѽ���) = Val("" & rsTmp!��Ѽ���)

                '��ʾ������־:һ����ҩֻ��ʾ�ڵ�һ��
                .TextMatrix(i, COL_��־) = Nvl(rsTmp!������־, 0)
                .TextMatrix(i, COL_ǩ����) = Nvl(rsTmp!ǩ����)
                Call SetRow��־ͼ��(i, 1)


                '����ҽ��״̬,ҩƷ����������ɫ
                '-------------------------------------------------------------------
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = .ForeColor
                If rsTmp!ҽ��״̬ = 8 Then
                    '��ֹͣ(�ѷ���)
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000    '����
                End If

                '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                If InStr(",5,6,", rsTmp!�������) > 0 And Not IsNull(rsTmp!�������) Then
                    If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", rsTmp!�������) > 0 Then
                        .Cell(flexcpFontBold, i, col_ҽ������) = True
                    End If
                End If

                'Pass�����������ʾ��ʾ��
                If mblnPass Then
                    If gobjPass.zlPassCheck(mobjPassMap) And Not IsNull(rsTmp!�����) Then
                        Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, rsTmp!�����)
                    End If
                End If

                Call Setҽ������(i, i)

                rsTmp.MoveNext
            Next

            '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '����ǩ��ͼ�����
            .Cell(flexcpPictureAlignment, .FixedRows, col_ҽ������, .Rows - 1, col_ҽ������) = 0

            Call .AutoSize(col_ҽ������)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
    Else
        mblnRowChange = False
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        mblnRowChange = True
    End If

    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check��ʼʱ��(ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ��������Ŀ�ʼʱ���Ƿ�Ϸ�
'˵����
'1.��ʼʱ�䲻��С�ڲ��˵ĹҺ�ʱ��
'2.����¼��ʱ,��ʼʱ�䲻��С�ڵ�ǰʱ��֮ǰ30����(�Ӷ�������ɿ���ʱ����ڿ�ʼʱ��30����)
    If Not IsDate(strStart) Then
        MsgBox "�����ҽ����ʼִ��ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If
        
    If Format(strStart, "yyyy-MM-dd HH:mm") < Format(mdat�Һ�ʱ��, "yyyy-MM-dd HH:mm") Then
        strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��˵ĹҺ�ʱ�� " & Format(mdat�Һ�ʱ��, "yyyy-MM-dd HH:mm") & " ��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check��ʼʱ�� = True
End Function

Private Function Check����ʱ��(ByVal strDate As String, ByVal strStart As String, ByVal strType As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ�������������/��Ѫʱ���Ƿ�Ϸ�
'˵����
'1.����/��Ѫʱ�䲻��С��ҽ���Ŀ�ʼʱ��
    Dim strInDate As String, strDateType As String
    
    If strType = "F" Then
        strDateType = "����"
    ElseIf strType = "K" Then
        strDateType = "��Ѫ"
    End If
    
    If Not IsDate(strDate) Then
        strMsg = "�����" & strDateType & "ʱ����Ч��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = strDateType & "ʱ�䲻��С��ҽ����ʼʱ�䡣"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check����ʱ�� = True
End Function

Private Function Check����ʱ��(ByVal strDate As String, ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ���鿪��ʱ���Ƿ���Ч
'˵������ӦС�ڲ��˹Һ�ʱ��
    If Not IsDate(strDate) Then
        strMsg = "����Ŀ���ʱ����Ч��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
        
'    If Format(strDate, "yyyy-MM-dd HH:mm") < Format(mdat�Һ�ʱ��, "yyyy-MM-dd HH:mm") Then
'        strMsg = "����ʱ�䲻��С�ڲ��˵ĹҺ�ʱ�� " & Format(mdat�Һ�ʱ��, "yyyy-MM-dd HH:mm") & " ��"
'        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'        Exit Function
'    End If
    Check����ʱ�� = True
End Function

Private Function Check�������(ByVal strҩƷIDs As String) As Boolean
'���ܣ��������ҩ,�г�ҩ���������;��ҩ�䷽����������
'������strҩƷIDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr���� As Variant, arr���� As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng��Ŀid As Long, str���� As String, blnδ�༭ As Boolean
    Dim lng���� As Long, lngRow As Long, lngSeekRow As Long
    
    On Error GoTo errH
        
    arr���� = Array(): arr���� = Array()
        
    strSQL = "Select /*+ rule*/ ���� From ���ƻ�����Ŀ" & _
        " Where ��ĿID IN(Select Column_Value From Table(f_Num2list([1]))) Group by ���� Having Count(*)>1"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҩƷIDs)
    For k = 1 To rsMain.RecordCount
        strSQL = "Select /*+ RULE */ A.����,A.����,A.��ĿID,B.����" & _
            " From ���ƻ�����Ŀ A,������ĿĿ¼ B" & _
            " Where A.��ĿID=B.ID And A.����=[2]" & _
            " And A.��ĿID IN(Select Column_Value From Table(f_Num2list([1])))" & _
            " Order by A.����,B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҩƷIDs, Val(rsMain!����))
        For i = 1 To rsTmp.RecordCount
            If rsTmp!���� <> lng���� Then
                If rsTmp!���� = 1 Then
                    ReDim Preserve arr����(UBound(arr����) + 1)
                Else
                    ReDim Preserve arr����(UBound(arr����) + 1)
                End If
                lng���� = rsTmp!����
            End If
            If rsTmp!���� = 1 Then
                arr����(UBound(arr����)) = arr����(UBound(arr����)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            Else
                arr����(UBound(arr����)) = arr����(UBound(arr����)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    
    '�ȼ����ò���(��ֹ����)
    If UBound(arr����) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr����) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(Mid(arr����(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) 'ÿ��Ŀ
                lng��Ŀid = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "��" & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��Ŀid), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & "�� " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "�ڲ���ҽ���з�������ҩƷ������ã�" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ټ�����ò���(�����Ƿ����)
    If UBound(arr����) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr����) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(Mid(arr����(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) 'ÿ��Ŀ
                lng��Ŀid = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "��" & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��Ŀid), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & "�� " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("�ڲ���ҽ���з�������ҩƷ�������ã�" & strMsg & vbCrLf & vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check���ƻ���(ByVal str����IDs As String) As Boolean
'���ܣ�����ҩƷ(��ҩ,��ҩ)�Ļ���
'������str����IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr���� As Variant, arr��ֹ As Variant, arrֹͣ As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng��Ŀid As Long, str���� As String, blnδ�༭ As Boolean
    Dim lng���� As Long, lngRow As Long, lngSeekRow As Long
    
    On Error GoTo errH
        
    arr���� = Array(): arr��ֹ = Array(): arrֹͣ = Array()
    
    strSQL = "Select /*+ rule*/ ���� From ���ƻ�����Ŀ" & _
        " Where ��ĿID IN(Select Column_Value From Table(f_Num2list([1]))) Group by ���� Having Count(*)>1"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs)
    For k = 1 To rsMain.RecordCount
        strSQL = "Select /*+ RULE */ A.����,A.������,A.����,A.��ĿID,B.����" & _
            " From ���ƻ�����Ŀ A,������ĿĿ¼ B" & _
            " Where A.��ĿID=B.ID And A.����=[2]" & _
            " And A.��ĿID IN(Select Column_Value From Table(f_Num2list([1])))" & _
            " Order by A.����,B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs, Val(rsMain!����))
        For i = 1 To rsTmp.RecordCount
            If rsTmp!���� <> lng���� Then
                If rsTmp!���� = 1 Then
                    ReDim Preserve arr����(UBound(arr����) + 1)
                    arr����(UBound(arr����)) = rsTmp!������
                ElseIf rsTmp!���� = 2 Then
                    ReDim Preserve arr��ֹ(UBound(arr��ֹ) + 1)
                    arr��ֹ(UBound(arr��ֹ)) = rsTmp!������
                Else
                    ReDim Preserve arrֹͣ(UBound(arrֹͣ) + 1)
                    arrֹͣ(UBound(arrֹͣ)) = rsTmp!������
                End If
                lng���� = rsTmp!����
            End If
            If rsTmp!���� = 1 Then
                arr����(UBound(arr����)) = arr����(UBound(arr����)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            ElseIf rsTmp!���� = 2 Then
                arr��ֹ(UBound(arr��ֹ)) = arr��ֹ(UBound(arr��ֹ)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            Else
                arrֹͣ(UBound(arrֹͣ)) = arrֹͣ(UBound(arrֹͣ)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    '�ȼ���ֹ��������
    If UBound(arr��ֹ) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr��ֹ) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(arr��ֹ(i), Chr(234))
            For j = 1 To UBound(arrItems) 'ÿ��Ŀ
                lng��Ŀid = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��Ŀid), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_EDIT)) = 1 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "��" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "�ڲ���ҽ���з����������ݻ����ų⣺" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ټ���Զ�ֹͣ����,���ﴦ��Ϊ��ֹ(����)
    If UBound(arrֹͣ) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arrֹͣ) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(arrֹͣ(i), Chr(234))
            For j = 1 To UBound(arrItems) 'ÿ��Ŀ
                lng��Ŀid = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��Ŀid), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False ': Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "��" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "�ڲ���ҽ���з����������ݻ����ų⣺" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ټ�������Ƿ��������
    If UBound(arr����) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr����) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(arr����(i), Chr(234))
            For j = 1 To UBound(arrItems) 'ÿ��Ŀ
                lng��Ŀid = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��Ŀid), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "��" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("�ڲ���ҽ���з����������ݻ����ų⣺" & strMsg & vbCrLf & vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check���ƻ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckStock(ByVal lngRow As Long) As String
'���ܣ����ָ��ҩƷ�еĿ�����
'���أ���=��ʾͨ��
    Dim dbl���� As Double, strMsg As String
    Dim lngִ�п���ID As Long, i As Integer
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Or .TextMatrix(lngRow, COL_���) = "4" And Val(.TextMatrix(lngRow, COL_��������)) = 1 Then
            If TheStockCheck(Val(.TextMatrix(lngRow, COL_ִ�п���ID)), .TextMatrix(lngRow, COL_���)) <> 0 Then
                If .TextMatrix(lngRow, COL_���) <> "" Then
                    '��ҩ����ֱ�Ӽ������
                    dbl���� = Val(.TextMatrix(lngRow, COL_����))
                    If dbl���� > 0 Then
                        If dbl���� > Val(.TextMatrix(lngRow, COL_���)) Then
                            strMsg = """" & .TextMatrix(lngRow, col_ҽ������) & """������ѣ�" & _
                                vbCrLf & vbCrLf & Get��������(Val(.TextMatrix(lngRow, COL_ִ�п���ID))) & _
                                IIF(InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0, _
                                "��ǰ���ÿ�治�� " & FormatEx(dbl����, 5) & .TextMatrix(lngRow, COL_���ﵥλ) & "��", _
                                "��ǰ���ÿ��Ϊ " & FormatEx(Val(.TextMatrix(lngRow, COL_���)), 5) & .TextMatrix(lngRow, COL_���ﵥλ) & "������ " & FormatEx(dbl����, 5) & .TextMatrix(lngRow, COL_���ﵥλ) & "��")
                        End If
                    End If
                End If
            End If
        ElseIf RowIn�䷽��(lngRow) And Val(.TextMatrix(lngRow, COL_����)) <> 0 Then
            '���ݸ�����������
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" And .TextMatrix(i, COL_���) <> "" Then
                        '����=�����װ(��ζ����*����)
                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                            dbl���� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ))
                        Else
                            dbl���� = Val(.TextMatrix(i, COL_����)) * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ)))
                        End If
                        If dbl���� > Val(.TextMatrix(i, COL_���)) Then
                            lngִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                            If TheStockCheck(lngִ�п���ID, .TextMatrix(i, COL_���)) = 0 Then Exit For
                            
                            strMsg = strMsg & vbCrLf & .TextMatrix(i, col_ҽ������) & _
                                "���������� " & FormatEx(dbl����, 5) & .TextMatrix(i, COL_���ﵥλ) & _
                                "�����ÿ��" & IIF(InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0, _
                                    "����", " " & FormatEx(Val(.TextMatrix(i, COL_���)), 5) & .TextMatrix(i, COL_���ﵥλ))
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            If strMsg <> "" Then
                strMsg = "��ҩ�䷽������ѣ�" & Get��������(lngִ�п���ID) & "������ζҩ��治�㣺" & vbCrLf & strMsg
            End If
        End If
    End With
    CheckStock = strMsg
End Function

Private Function CheckMoney() As Boolean
'���ܣ����ñ������
'˵�����������ۼƷ��ñ�����ʽʱ,ֻ���ѡ�
    Dim rsTmp As New ADODB.Recordset
    Dim str���ò��� As String, strSQL As String
    Dim curԤ�� As Currency, cur��� As Currency
    Dim cur������ As Currency
    
    On Error GoTo errH
    '�������
    strSQL = "Select Ԥ�����,Nvl(Ԥ�����,0)-Nvl(�������,0) as ��� From ������� Where ����=1 And ���� = 1 And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If Not rsTmp.EOF Then
        curԤ�� = Nvl(rsTmp!Ԥ�����, 0)
        cur��� = Nvl(rsTmp!���, 0)
    End If
    
    '������
    strSQL = "Select ������ From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If Not rsTmp.EOF Then cur������ = Nvl(rsTmp!������, 0)
    
    '��Ԥ����Ĳ��˲��ж�
    If curԤ�� <> 0 Then
        '�Ƿ�ҽ��
        strSQL = "Select zl_PatiWarnScheme([1]) as ���ò��� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If Not rsTmp.EOF Then str���ò��� = Nvl(rsTmp!���ò���)
            
        '����ֵ:NULL��0������ͬ���崦��
        strSQL = "Select ����ֵ From ���ʱ����� Where ��������=1 And Nvl(����ID,0)=0 And ����ֵ is Not NULL And ���ò���=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���ò���)
        If Not rsTmp.EOF Then
            If cur��� + cur������ < Nvl(rsTmp!����ֵ, 0) Then
                If MsgBox("���˵�ǰ����ʣ��� " & FormatEx(cur��� + cur������, 2) & IIF(cur������ <> 0, "(��������:" & FormatEx(cur������, 2) & ")", "") & " ���ڱ���ֵ " & FormatEx(Nvl(rsTmp!����ֵ, 0), 2) & "��Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    CheckMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'���ܣ���ȡ��ID��ͬ��һ��ҽ���кŷ�Χ(ע�⿼��һ����ҩ�еĿ���)
    Dim lngS��ID As Long, lngO��ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, .RowData(lngRow), Val(.TextMatrix(lngRow, COL_���ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_���ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '��������
                If lngO��ID = lngS��ID Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_���ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '��������
                If lngO��ID = lngS��ID Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function CheckAdvice() As Boolean
'���ܣ���鵱ǰ����(Ӥ��)��ҽ�������Ƿ�Ϸ�
'˵��������в��Ϸ��ĵط����ڱ���������ʾ����λ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim bln�䷽�� As Boolean, bln������ As Boolean
    Dim dbl���� As Double, strMsg As String
    Dim strҩƷIDs As String, str����IDs As String
    Dim lngCount As Long, lngRow As Long
    Dim blnSkipStock As Boolean, blnSkipTotal As Boolean
    Dim vMsg As VbMsgBoxResult, sng���� As Single
    Dim blnValid As Boolean, lngRxCount As Long
    Dim blnAppend As Boolean, i As Long, j As Long, k As Long
    Dim str����IDs As String, str���IDs As String
    Dim str��ҩ�� As String, strExtra As String, lng��ID As Long
    Dim lng��������ID As Long, lng��ҩִ������ As Long, str��λ���� As String
    Dim lngBegin As Long, lngEnd As Long
    Dim blnExists As Boolean
    Dim dblOneDay As Double
    Dim datCur As Date
    Dim blnNo As Boolean, blnCheck���� As Boolean, blnOut As Boolean
    Dim strIDs1 As String, strIDs2 As String, strҽ������ As String
    Dim lngSame As Long
    
    On Error GoTo errH
    If vsAdvice.Enabled = True Then
        If Me.ActiveControl.Name = "txt����" Then
            Call txt����_Validate(False)
        ElseIf Me.ActiveControl.Name = "txt����" Then
            Call txt����_Validate(False)
        ElseIf Me.ActiveControl.Name = "txt����" Then
            Call txt����_Validate(False)
        End If
        vsAdvice.SetFocus  '�����궨λ�������껹û�뿪�������δ�������ݡ�
    End If
    If Not CheckApply Then Exit Function
    datCur = zlDatabase.Currentdate
    '��ϵļ��
    '-----------------------------------------------------------------------------------------
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col���)) <> "" Then
                If mint���� = 920 Then '����ҽ������Ҫ��
                    If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 82 Then
                        .Row = i: .Col = col���
                        MsgBox "�������̫����ֻ����82���ַ���41�����֡�", vbInformation, gstrSysName
                        vsDiag.SetFocus: Exit Function
                    End If
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 200 Then
                    .Row = i: .Col = col���
                    MsgBox "�������̫����ֻ����200���ַ���100�����֡�", vbInformation, gstrSysName
                    vsDiag.SetFocus: Exit Function
                End If
                If .TextMatrix(i, col����ʱ��) <> "" Then
                    If Format(datCur, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col����ʱ��), "YYYY-MM-DD HH:mm") Then
                         .Row = i: .Col = col����ʱ��
                        MsgBox "����ʱ��Ӧ�����ڵ�ǰʱ�䡣", vbInformation, gstrSysName
                        vsDiag.SetFocus: Exit Function
                    End If
                End If
                lngSame = 0
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, col���)) <> "" And .Cell(flexcpData, i, col��ҽ) = .Cell(flexcpData, j, col��ҽ) Then 'ͬ�������
                        If .TextMatrix(j, col���) & "|" & .TextMatrix(j, col��ҽ֤��) = .TextMatrix(i, col���) & "|" & .TextMatrix(i, col��ҽ֤��) Then
                            .Row = i: .Col = col���
                            MsgBox "���ִ���������ͬ�������Ϣ��", vbInformation, gstrSysName
                            vsDiag.SetFocus: Exit Function
                        ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                            If Val(.TextMatrix(j, col����ID)) & "|" & .TextMatrix(j, col��ҽ֤��) = Val(.TextMatrix(i, col����ID)) & "|" & .TextMatrix(i, col��ҽ֤��) Then
                                .Row = i: .Col = col���
                                MsgBox "���ִ���������ͬ�ļ�����Ϣ��", vbInformation, gstrSysName
                                vsDiag.SetFocus: Exit Function
                            End If
                        ElseIf Val(.TextMatrix(i, col���ID)) <> 0 And .Cell(flexcpData, i, col��ҽ) = 0 Then
                            '����ҽ��ϴ�֤��,�����޶�Ӧ֤��ID,���ID����ͬ
                            If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, col���ID)) Then
                                .Row = i: .Col = col���
                                MsgBox "���ִ���������ͬ�������Ϣ��", vbInformation, gstrSysName
                                vsDiag.SetFocus: Exit Function
                            End If
                        End If
                        If .Cell(flexcpData, i, col��ҽ) <> 0 Then '��ҽ���
                             If .TextMatrix(j, col���) = .TextMatrix(i, col���) Then
                                lngSame = lngSame + 1
                             ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                                If Val(.TextMatrix(j, col����ID)) = Val(.TextMatrix(i, col����ID)) Then
                                    lngSame = lngSame + 1
                                End If
                             End If
                            If lngSame >= 2 Then
                                .Row = i: .Col = col���
                                MsgBox "�����������ϵ������ͬ��֤��ͬ����ϣ���ϲ���ȷ��", vbInformation, gstrSysName
                                vsDiag.SetFocus: Exit Function
                            End If
                        End If
                    End If
                Next
                If Val(.TextMatrix(i, col����ID)) <> 0 Then str����IDs = str����IDs & "," & Val(.TextMatrix(i, col����ID))
                If Val(.TextMatrix(i, col���ID)) <> 0 Then str���IDs = str���IDs & "," & Val(.TextMatrix(i, col���ID))
            End If
        Next
    End With
        
    '-----------------------------------------------------
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            bln�䷽�� = False: bln������ = False
            '�����������޸�ҩƷ�еĴ���ְ����
            If .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, COL_���)) > 0 And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                strMsg = CheckOneDuty(.TextMatrix(i, col_ҽ������), .TextMatrix(i, COL_����ְ��), .TextMatrix(i, COL_����ҽ��), InStr(",1,2,", mstr������) > 0 And mstr������ <> "")
                If strMsg <> "" Then
                    .Col = col_ҽ������
                    If .TextMatrix(i, COL_���) = "7" Then
                        lngRow = .FindRow(CLng(.TextMatrix(i, COL_���ID)), i + 1)
                        If lngRow <> -1 Then .Row = lngRow
                    Else
                        .Row = i
                    End If
                    Call .ShowCell(.Row, .Col)
                    MsgBox strMsg, vbInformation, gstrSysName
                    .Refresh
                    Call vsAdvice_KeyPress(13)
                    Exit Function
                End If
                
                '������ҩ���
                If gblnKSSStrict Then
                    If Val(.TextMatrix(i, COL_�����ȼ�)) > 0 Then
                        If Val(.TextMatrix(i, COL_��ҩĿ��)) = 0 Then
                            strMsg = ",������ҩҪ��Ǽ���ҩĿ�ġ�"
                            .Col = COL_��ҩĿ��: Exit For
                        End If
                        
                        If Val(.TextMatrix(i, COL_�����ȼ�)) = 3 And mbytPatiType <> 2 Then
                            strMsg = ",����Ǽ���ҺŲ���ʹ������ʹ�ü��Ŀ���ҩ�"
                            .Col = col_ҽ������: Exit For
                        End If
                        
                        '�������Աû�п�����ҩ����Ȩ�����ֹ����
                        If UserInfo.��ҩ���� = 0 Then
                            strMsg = ",��û�п�����ҩȨ�ޣ�����ϵ����Ա��"
                            .Col = col_ҽ������: Exit For
                        End If
                        
                        If Val(.TextMatrix(i, COL_EDIT)) = 2 Then
                            If UserInfo.��ҩ���� < Val(.TextMatrix(i, COL_�����ȼ�)) And Val(.TextMatrix(i, COL_��־)) <> 1 Then
                                .TextMatrix(i, COL_���״̬) = 1
                            End If
                        End If
                        
                        'һ��ҩƷ�У�ֻҪ��һ��Ϊ����˻�δ���ͨ�������������(û������ʱ��������Ϊ���ܸ�ҩ;���Ǻ����ż��ϵģ�����ĵ�϶�)
                        If Val(.TextMatrix(i, COL_���״̬)) > 0 Then
                            Call GetRowScope(i, lngBegin, lngEnd)
                            For j = lngBegin To lngEnd
                                If j <> i Then
                                    .TextMatrix(j, COL_���״̬) = .TextMatrix(i, COL_���״̬)
                                End If
                            Next
                        End If
                        
                        
                        '����ҽ�����
                        If Val(.TextMatrix(i, COL_��־)) = 1 Then
                            If .TextMatrix(i, COL_��ҩ����) = "" Then
                                strMsg = ",����ʹ�õĿ�����ҩҪ��������ҩ���ɡ�"
                                .Col = COL_��ҩ����: Exit For
                            End If
                            If Val(.TextMatrix(i, COL_����)) > 1 Then
                                strMsg = ",����ʹ�õĿ�����ҩҪ�����һ�졣"
                                .Col = COL_����: Exit For
                            ElseIf Val(.TextMatrix(i, COL_����)) = 0 Then
                                'δ��������ʱ���з�������������/����/Ƶ��
                                If .TextMatrix(i, COL_����ϵ��) <> "" And .TextMatrix(i, COL_�����װ) <> "" Then
                                    '�����һ�������
                                    dblOneDay = FormatEx(CalcȱʡҩƷ����( _
                                    Val(.TextMatrix(i, COL_����)), 1, _
                                    Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), _
                                    .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), _
                                    Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_�����װ)), _
                                    Val(.TextMatrix(i, COL_�ɷ����))), 5)
                                    If Val(.TextMatrix(i, COL_����)) > dblOneDay Then
                                        strMsg = ",����ʹ�õĿ�����ҩ����������һ���ʹ������" & dblOneDay & "��"
                                        .Col = COL_����: Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            
            '�����Զ����ҽ����麯��
            If .RowData(i) <> 0 And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 And Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                
                strExtra = CStr(.Cell(flexcpData, i, COL_ҽ������))
                lng��������ID = 0
                lng��ҩִ������ = 0
                str��λ���� = ""
                If .TextMatrix(i, COL_���) = "F" Then
                    lng��ID = Val(.TextMatrix(i, COL_���ID))
                    If lng��ID = 0 Then lng��ID = .RowData(i)
                    For k = i + 1 To .Rows - 1
                        If Val(.TextMatrix(k, COL_���ID)) <> lng��ID Then Exit For
                        If .TextMatrix(k, COL_���) = "G" Then
                            lng��������ID = .TextMatrix(k, COL_������ĿID)
                        End If
                    Next
                ElseIf InStr("4,5,6,7", .TextMatrix(i, COL_���)) > 0 Then
                    k = .FindRow(CLng(Val(.TextMatrix(i, COL_���ID))), i + 1)
                    If k > 0 Then lng��ҩִ������ = Val(.TextMatrix(k, COL_ִ������))
                
                ElseIf .TextMatrix(i, COL_���) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    lng��ID = .RowData(i)
                    For k = i + 1 To .Rows - 1
                        If Val(.TextMatrix(k, COL_���ID)) <> lng��ID Then Exit For
                        str��λ���� = str��λ���� & "," & .TextMatrix(k, COL_�걾��λ) & ":" & .TextMatrix(k, COL_��鷽��)
                    Next
                    str��λ���� = Mid(str��λ����, 2)
                End If
                
                strExtra = strExtra & "||" & lng��ҩִ������ & "||" & lng��������ID & "||" & IIF(str��λ���� = "", " ", str��λ����) & "||" & Val(.TextMatrix(i, COL_�շ�ϸĿID))
                
                strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ��� From Dual"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 1, mlng����ID, mlng�Һ�ID, mint����, 1, _
                     .TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), _
                    Val(.TextMatrix(i, COL_��������ID)), CStr(.TextMatrix(i, COL_����ҽ��)), _
                    Val(.TextMatrix(i, COL_ִ�п���ID)), Val(.TextMatrix(i, COL_ִ������)), Val(.TextMatrix(i, COL_ִ�б��)), _
                    Val(.TextMatrix(i, COL_����)), strExtra)
                
                If Not rsTmp.EOF Then
                    strMsg = Nvl(rsTmp!���)
                    If strMsg <> "" Then
                        Select Case Val(Split(strMsg, "|")(0))
                        Case 1 '��ʾ
                            If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                strMsg = "": Exit For
                            End If
                        Case 2 '��ֹ
                            MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                            strMsg = "": Exit For
                        End Select
                        strMsg = ""
                    End If
                End If
            End If
            
            '��������Ϸ��Լ��
            If .RowData(i) <> 0 Then
            
                '���Ա�������Ŀ�ļ�飨ֻ��ԣ����顢��顢������
                If InStr(",C,D,F,5,6,7,", .TextMatrix(i, COL_���)) > 0 And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 And (mstr�Ա� = "��" Or mstr�Ա� = "Ů") And Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                    If .TextMatrix(i, COL_���) = "F" Or .TextMatrix(i, COL_���) = "D" And Not .RowHidden(i) Or .TextMatrix(i, COL_���) = "C" And .RowHidden(i) Or InStr(",5,6,7,", .TextMatrix(i, COL_���)) > 0 Then
                        strSQL = "Select Decode(a.�����Ա�, 1, '��', 2, 'Ů', 'δ֪') As �Ա� From ������ĿĿ¼ A Where a.Id = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_������ĿID)))
                        
                        If rsTmp!�Ա� <> mstr�Ա� And rsTmp!�Ա� <> "δ֪" Then
                            If .TextMatrix(i, COL_���) = "D" Then
                                strMsg = "���������Ա�""" & mstr�Ա� & """�����á�"
                            Else
                                strMsg = Decode(.TextMatrix(i, COL_���), "C", "������Ŀ", "F", "��������", "D", "�����Ŀ", "ҩƷ") & "�����Ա�""" & mstr�Ա� & """�����á�"
                            End If
                            .Col = col_ҽ������: Exit For
                        End If
                    End If
                End If
                
                If .RowHidden(i) Then
                    
                    '�����������޸ĵ���
                    '---------------------------------------------------
                    If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        If Not Check����Ӧ��(i, strMsg) Then Exit For
                        '���ִ�п���
                        If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                            strMsg = "û��ȷ��ִ�п��ҡ�"
                            .Col = COL_ִ�п���ID: Exit For
                        End If
                    
                        If .TextMatrix(i, COL_���) = "D" And .TextMatrix(i, COL_�걾��λ) <> "" Then
                            If Check��鲿λEnable(.TextMatrix(i, COL_������ĿID), .TextMatrix(i, COL_�걾��λ), mstr�Ա�, .TextMatrix(i, COL_��鷽��), blnExists) = False Then
                                If blnExists = True Then
                                    strMsg = "�еĲ�λ��" & .TextMatrix(i, COL_�걾��λ) & "�����Ա�""" & mstr�Ա� & """�����á�"
                                Else
                                    Call cmdExt_Click
                                End If
                                .Col = col_ҽ������: Exit For
                            End If
                        End If
                    End If
                Else
                    bln�䷽�� = RowIn�䷽��(i)
                    bln������ = RowIn������(i)
                    lngRow = i
                    If bln�䷽�� Then '�õ��䷽�ĵ�һҩƷ��
                        lngRow = .FindRow(CStr(.RowData(i)), , COL_���ID)
                        '�Ա�ҩ�����ҩ��
                        If Not (Val(.TextMatrix(lngRow, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5) And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                            
                            str��ҩ�� = ""
                            If Check��ҩ�洢�ⷿ(lngRow, i, str��ҩ��, vsAdvice, 1, mlng���˿���id, COL_���, col_ҽ������, COL_�շ�ϸĿID, COL_ִ�п���ID) = False Then
                                strMsg = "�е�[" & str��ҩ�� & "]û�д洢�ڵ�ǰѡ���ҩ�������ҩ�����Ƿ����ڵ�ǰ���˿��ҵģ�����ʹ�ø�ҩƷ��"
                                Exit For
                            End If
                        End If
                    ElseIf bln������ Then '�õ�����ҽ����
                        lngRow = .FindRow(CStr(.RowData(i)), , COL_���ID)
                    End If
                    
                    'δ���͵�ҽ����
                    '------------------------------------
                    If Val(.TextMatrix(i, COL_״̬)) = 1 Then
                        lngCount = lngCount + 1
                        
                        '����¼�뵥��:����:��ҩ���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����¼��(Ҳ�ɲ�¼)
                        If (Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 And InStr(",1,2,", Val(.TextMatrix(lngRow, COL_���㷽ʽ))) > 0) _
                            Or InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                            'ҩƷ����¼�뵥��
                            If .TextMatrix(lngRow, COL_����) <> "" Or mbln���� And InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                                If Not IsNumeric(.TextMatrix(lngRow, COL_����)) Or Val(.TextMatrix(lngRow, COL_����)) <= 0 Then
                                    strMsg = "û��¼����ȷ�ĵ���������"
                                    .Col = COL_����: Exit For
                                End If
                            End If
                        End If
                        
                        '����¼��������������ҩƷ�����ָ����Ҫ¼��
                        If mbln���� And InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                            If Val(.TextMatrix(i, COL_����)) <= 0 Then
                                strMsg = "��¼����ȷ����ҩ������"
                                .Col = COL_����: Exit For
                            End If
                        End If
                        
                        '����¼������:�䷽,����(ҩƷ������)
                        If Not IsNumeric(.TextMatrix(i, COL_����)) Or Val(.TextMatrix(i, COL_����)) <= 0 Then
                            '��Ѫҽ��������������
                            If Not (.TextMatrix(i, COL_���) = "K" And .TextMatrix(i, COL_����) = "") Then
                                If bln�䷽�� Then
                                    strMsg = "û��¼����ȷ����ҩ�䷽������"
                                ElseIf InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                                    strMsg = "û��¼����ȷ��ҩƷ�ܸ�������"
                                Else
                                    strMsg = "û��¼����ȷ��������"
                                End If
                                .Col = COL_����: Exit For
                            End If
                        End If
                                            
                        '����¼��Ƶ��:����ҲҪ���,����ָ��ʹ��,���Բ�¼��ִ��ʱ��
                        If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Or InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Or bln�䷽�� Then
                            If .TextMatrix(lngRow, COL_Ƶ��) = "" Then
                                strMsg = "û��ȷ��ִ��Ƶ�ʡ�"
                                .Col = COL_Ƶ��: Exit For
                            End If
                            
                            'ִ��ʱ���ж�:��ѡƵ�ʵı�������(������������������¼��,Ҫע�ⷢ�͵ȵط��Ĵ���)
                            If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" And .TextMatrix(lngRow, COL_�����λ) <> "����" And .TextMatrix(lngRow, COL_Ƶ��) <> "��Ҫʱ" And .TextMatrix(lngRow, COL_Ƶ��) <> "��Ҫʱ" Then
                                If Not bln������ Then '���������ʾ�еĲɼ�����Ϊ��ѡƵ��,��������ĿΪһ����
                                    If Val(.TextMatrix(lngRow, COL_Ƶ������)) <> 1 Then
                                        strMsg = "û��¼��ִ��ʱ�䷽����"
                                        .Col = COL_ִ��ʱ��: Exit For
                                    End If
                                End If
                            End If
                        End If
                        
                        '����¼��ִ�п���:�Ƕ�����Ժ��ִ��ʱ(�䷽��ҩƷ�н����ж�)
                        If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                            If .TextMatrix(lngRow, COL_���) = "Z" And InStr(",1,2,", Val(.TextMatrix(lngRow, COL_��������))) > 0 Then
                                If Val(.TextMatrix(lngRow, COL_��������)) = 1 Then
                                    strMsg = "û��ȷ������ҽ�������ۿ��ҡ�"
                                ElseIf Val(.TextMatrix(lngRow, COL_��������)) = 2 Then
                                    strMsg = "û��ȷ��סԺҽ����סԺ���ҡ�"
                                End If
                                .Col = COL_ִ�п���ID: Exit For
                            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                                strMsg = "û��ȷ��ִ�п��ҡ�"
                                .Col = COL_ִ�п���ID: Exit For
                            End If
                        End If
                        If lngRow <> i And Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                            If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                                strMsg = "û��ȷ��ִ�п��ҡ�"
                                .Col = COL_ִ�п���ID: Exit For
                            End If
                        End If
                        
                        '����ʱ���ж�
                        If Not Check����ʱ��(.Cell(flexcpData, i, COL_����ʱ��), .Cell(flexcpData, i, COL_��ʼʱ��), False, strMsg) Then
                            .Col = COL_����ʱ��: Exit For
                        End If
                        
                        '�����������Ƽ��
                        If gintRXCount > 0 And InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 _
                            And Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                            lngRxCount = GetMergeCount(vsAdvice, i, COL_���ID, COL_�շ�ϸĿID)
                            If lngRxCount > gintRXCount Then
                                strMsg = "һ����ҩ��ҩƷ���� " & lngRxCount & " ���Ѵﵽ�򳬹�ҩƷ���������������� " & gintRXCount & " �֡�"
                                .Col = col_ҽ������: Exit For
                            End If
                        End If
                    End If
                    
                    '�����������޸ĵ���
                    '---------------------------------------------------
                    If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        If Not Check����Ӧ��(i, strMsg) Then Exit For
                        '��ʼʱ���ж�:ֻ��������ҽ�������ж�,��Ϊ�����ǲ�׼�޸Ŀ�ʼʱ���(�����жϱ��޸ĵ�ҽ����ʼʱ��������Ч��)
                        If .TextMatrix(i, COL_EDIT) = "1" Then
                            If Not Check��ʼʱ��(.Cell(flexcpData, i, COL_��ʼʱ��), False, strMsg) Then
                                .Col = COL_��ʼʱ��: Exit For
                            End If
                        End If
                        '����/��Ѫҽ��������/��Ѫʱ���ж�
                        If .TextMatrix(i, COL_���) = "F" Or .TextMatrix(i, COL_���) = "K" Then
                            If Not Check����ʱ��(.TextMatrix(i, COL_����ʱ��), .Cell(flexcpData, i, COL_��ʼʱ��), .TextMatrix(.Row, COL_���), False, strMsg) Then
                                .Col = COL_����ʱ��: Exit For
                            End If
                            If .TextMatrix(i, COL_���) = "K" Then
                                'ֻ�н���ҽ��������Ѫԭ��
                                If .TextMatrix(i, COL_��ҩ����) <> "" And .TextMatrix(i, COL_��־) <> "1" Then
                                    .TextMatrix(i, COL_��ҩ����) = ""
                                Else
                                    If gbln��Ѫ�ּ����� And .TextMatrix(i, COL_��־) = "1" And .TextMatrix(i, COL_��ҩ����) = "" Then
                                        strMsg = "��������Ѫ�ּ�����󣬽�����Ѫҽ��������д��Ѫԭ��"
                                        .Col = COL_��ҩ����: Exit For
                                    End If
                                End If
                            End If
                        End If
                        
                        '��ҩ;������ҩ�÷����ɼ��������ü��
                        If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                            If Val(.TextMatrix(i, COL_���ID)) = .RowData(i + 1) And Val(.TextMatrix(i + 1, COL_������ĿID)) = 0 Then
                                strMsg = "û�����ö�Ӧ�ĸ�ҩ;����"
                                .Col = COL_�÷�: Exit For
                            End If
                        End If
                        If .TextMatrix(i, COL_���) = "E" And Val(.TextMatrix(i, COL_������ĿID)) = 0 Then
                            If .RowData(i) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                                If InStr(",7,E,", .TextMatrix(i - 1, COL_���)) > 0 Then
                                    strMsg = "��ҩ�䷽û�����ö�Ӧ���÷���"
                                ElseIf .TextMatrix(i - 1, COL_���) = "C" Then
                                    strMsg = "û�����ö�Ӧ�ı걾�ɼ�������"
                                End If
                                .Col = COL_�÷�: Exit For
                            End If
                        End If
                                                
                        '�����������:����Ҫ����һ��Ƶ�����ڵ�����
                        If InStr(",4,5,6,", .TextMatrix(i, COL_���)) > 0 Or bln�䷽�� Then
                            If Not blnSkipTotal And .TextMatrix(i, COL_Ƶ��) <> "" And Val(.TextMatrix(i, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(i, COL_Ƶ�ʼ��)) <> 0 Then
                                strMsg = ""
                                If bln�䷽�� Then '�ж�
                                    dbl���� = CalcȱʡҩƷ����(1, 1, Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ))
                                    If Val(.TextMatrix(i, COL_����)) < dbl���� Then
                                        strMsg = .TextMatrix(i, col_ҽ������) & vbCrLf & vbCrLf & _
                                            "�ڰ�""" & .TextMatrix(i, COL_Ƶ��) & """ִ��ʱ,������Ҫ " & dbl���� & "����"
                                    End If
                                ElseIf Val(.TextMatrix(i, COL_����ϵ��)) <> 0 And Val(.TextMatrix(i, COL_����)) <> 0 Then
                                    sng���� = Val(.TextMatrix(i, COL_����))
                                    If sng���� = 0 Then sng���� = 1
                                    dbl���� = CalcȱʡҩƷ����(Val(.TextMatrix(i, COL_����)), sng����, Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_�����װ)), Val(.TextMatrix(i, COL_�ɷ����)))
                                    If Val(.TextMatrix(i, COL_����)) < dbl���� Then
                                        strMsg = .TextMatrix(i, col_ҽ������) & vbCrLf & vbCrLf & _
                                            "�ڰ�ÿ�� " & .TextMatrix(i, COL_����) & .TextMatrix(i, COL_������λ) & "," & _
                                            .TextMatrix(i, COL_Ƶ��) & IIF(mbln���� And .TextMatrix(i, COL_���) <> "4", ",��ҩ " & sng���� & " ��", "") & _
                                            "ִ��ʱ,������Ҫ " & dbl���� & .TextMatrix(i, COL_������λ) & "��"
                                    End If
                                End If
                                If strMsg <> "" And False Then '��ʾ
                                    .Row = i: .Col = COL_����: Call .ShowCell(.Row, .Col)
                                    vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^Ҫ������", Me)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                                        Exit Function
                                    ElseIf vMsg = vbIgnore Then
                                        blnSkipTotal = True
                                    End If
                                End If
                            End If
                        End If
                            
                            '���ҩƷ�Ƿ����ͳ��ڣ�����ʱӦ�ӵ�һ��û����д����˵�����п�ʼ
                            If (gbyt����ԭ�� = 1 And InStr(gstr��¼��������, "," & mlng���˿���id & ",") = 0) And Not blnCheck���� _
                                And (.TextMatrix(i, COL_�Ƿ���) = "1" Or .TextMatrix(i, COL_�Ƿ���) = "1") And .TextMatrix(i, COL_����˵��) = "" Then
                                blnCheck���� = SetAll����˵��(i, blnOut) '�����ļ��ֻҪ�������ִ����һ�ξ��������ж��Ѿ�����˲����ٱ�����
                                If blnOut Then
                                    strMsg = "Not Null" '����Ӧ�ø�ֵȷ�� strMsg ��Ϊ�ռ���
                                    .Col = COL_����˵��: Exit For
                                End If
                            End If
                        
                        
                        'ҩƷ�����:ֻ����,����Ҳֻ�Ա��α༭�Ĳ��ж�
                        If (InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Or bln�䷽�� _
                            Or .TextMatrix(i, COL_���) = "4" And Val(.TextMatrix(i, COL_��������)) = 1) And Not blnSkipStock Then
                            strMsg = CheckStock(i)
                            If strMsg <> "" Then
                                .Row = i: .Col = col_ҽ������: Call .ShowCell(.Row, .Col)
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^Ҫ������", Me)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    Exit Function
                                ElseIf vMsg = vbIgnore Then
                                    blnSkipStock = True
                                End If
                            End If
                        End If
                        
                        'ִ��ʱ��Ϸ��Լ��
                        If .TextMatrix(i, COL_ִ��ʱ��) <> "" And .TextMatrix(i, COL_Ƶ��) <> "" Then
                            blnValid = ExeTimeValid(.TextMatrix(i, COL_ִ��ʱ��), Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ))
                            If Not blnValid Then
                                If .TextMatrix(i, COL_�����λ) = "��" Then
                                    strMsg = COL_����ִ��
                                ElseIf .TextMatrix(i, COL_�����λ) = "��" Then
                                    strMsg = COL_����ִ��
                                ElseIf .TextMatrix(i, COL_�����λ) = "Сʱ" Then
                                    strMsg = COL_��ʱִ��
                                End If
                                strMsg = "¼���ִ��ʱ�䷽����ʽ����ȷ�����顣" & vbCrLf & vbCrLf & "����" & vbCrLf & strMsg
                                .Col = COL_ִ��ʱ��: Exit For
                            End If
                        End If
                        
                        'ҽ��������:��һ��ҽ����һ�ɼ���Ϊ׼
                        If InStr(",5,6,", .TextMatrix(i, COL_���)) = 0 _
                            Or Val(.TextMatrix(i - 1, COL_���ID)) <> Val(.TextMatrix(i, COL_���ID)) Then
                            If gintҽ������ = 2 Then mbln���Ѷ��� = True
                            Call GetInsureStr(strIDs1, strIDs2, strҽ������, i)
                            strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, 1, strIDs1, strIDs2, strҽ������)
                            If strMsg <> "" Then
                                .Row = i: .Col = col_ҽ������: Call .ShowCell(.Row, .Col)
                                If gintҽ������ = 1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", Me)
                                    If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                                    If vMsg = vbIgnore Then mbln���Ѷ��� = False
                                ElseIf gintҽ������ = 2 Then
                                    MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName
                                    Exit Function
                                End If
                                strMsg = "" '��ֹ������������
                            End If
                        End If
                        
                        'ҽ���ܿ�ʵʱ��⣺�״�����(����)���߸���ʱ���
                        If mint���� <> 0 And .Cell(flexcpData, i, COL_״̬) = 0 Then
                            If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) Then
                                If MakePriceRecord(i) Then
                                    If Not gclsInsure.CheckItem(mint����, 0, 0, mrsPrice) Then
                                        .Row = i: .Col = col_ҽ������
                                        Call .ShowCell(.Row, .Col)
                                        If txt����.Enabled Then
                                            txt����.SetFocus
                                        ElseIf txtҽ������.Enabled Then
                                            txtҽ������.SetFocus
                                        End If
                                        Exit Function
                                    End If
                                End If
                                '���Ϊ�Ѿ����˼��
                                .Cell(flexcpData, .Row, COL_״̬) = 1
                            End If
                        End If
                    End If
                                    
                    'ҽ�����븽����д��飺
                    'ֻ�����¼���ҽ�����޸ĵ�ҽ���޸�ʱ�Ѽ��
                    '".Cell(flexcpData, i, COL_����) = 1"�Ŀ��ܻ�û������������м�飬ֻ���Զ��滻��
                    If Val(.TextMatrix(i, COL_EDIT)) = 1 And .TextMatrix(i, COL_����) <> "" Then
                        strMsg = CheckAdviceAppend(.TextMatrix(i, COL_����))
                        If strMsg <> "" Then
                            strMsg = "���븽��""" & strMsg & """û��¼�룬��ȷ��ϵͳĬ����д����Ϣ�Ƿ���ȷ��"
                            .Col = col_ҽ������: blnAppend = True: Exit For
                        End If
                    End If
                                    
                    '���������ռ�:��������Чҽ����,��Ϊ�����ѷ��͵���δ���͵Ļ���
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        '����ҩƷ������ɼ��:������Ч
                        strҩƷIDs = strҩƷIDs & "," & Val(.TextMatrix(i, COL_������ĿID))
                    ElseIf Not bln�䷽�� Then
                        '���ܼ����������������ڲ�֮�估�ڲ���������Ŀ֮��
                        str����IDs = str����IDs & "," & Val(.TextMatrix(i, COL_������ĿID))
                    End If
                End If
            End If
        Next
        
        '--------------------------------------------------------------------------
        '�м��˳��Ĵ�����ʾ
        If i <= .Rows - 1 Then
            .Row = i: Call .ShowCell(.Row, .Col)
            If strMsg <> "" Then
                If bln�䷽�� Then
                    strMsg = "����ҩ�䷽" & strMsg
                Else
                    strMsg = """" & .TextMatrix(i, col_ҽ������) & """" & strMsg
                End If
                If Not blnOut Then '����˵����ʾ���⴦��
                    MsgBox strMsg, vbInformation, gstrSysName
                End If
                blnOut = False
                .Refresh
            End If
            Call vsAdvice_KeyPress(13)
            If blnAppend Then '�Ƿ񵯳����븽��༭
                If cmdExt.Enabled And cmdExt.Visible Then Call cmdExt_Click
            End If
            Exit Function
        End If
        
        '���ҩƷ�������
        If strҩƷIDs <> "" Then
            If Not Check�������(Mid(strҩƷIDs, 2)) Then Exit Function
        End If
        '���������Ŀ����
        If str����IDs <> "" Then
            If Not Check���ƻ���(Mid(str����IDs, 2)) Then Exit Function
        End If
    End With
    
    '���ñ���:��δ����ҽ��ʱ
    If lngCount > 0 Then
        If Not CheckMoney Then Exit Function
    End If
    '���ﲻ������Ⱦ�����濨
    If Not mbln���� Then
        '��������ж��Ƿ�Ӧ����д��Ⱦ�����濨
        RaiseEvent CheckInfectDisease(False, Mid(str����IDs, 2), Mid(str���IDs, 2), blnNo)
    End If
        If blnNo Then Exit Function
    '--��鲡���Ƿ��˺ţ��Ƿ�ȡ���˽���
    If Not CheckBackNo(mstr�Һŵ�) Then Exit Function
    
    CheckAdvice = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SeekNextControl() As Boolean
'���ܣ���λ����һ������Ŀؼ���,��������������Ƿ��Զ�����һ��ҽ��
'���أ����ͨ��SetFocusǿ�ƶ�λ��,�򷵻�True
    Dim objActive As Object, objNext As Object
    Dim blnDo As Boolean, i As Long
    Dim strSkip As String
    
    Set objActive = Me.ActiveControl
    
    If Not objActive Is Nothing Then
        If TypeName(objActive) = "TextBox" Or TypeName(objActive) = "ComboBox" Then
            If objActive.Container Is fraAdvice Then
                strSkip = GetInputSkip(vsAdvice.Row)
                Set objNext = GetNextControl(objActive.TabIndex, Me, strSkip)
                If Not objNext Is Nothing Then
                    If objNext Is vsAdvice Then
                        For i = vsAdvice.Row + 1 To vsAdvice.Rows - 1
                            If Not vsAdvice.RowHidden(i) Then
                                Call AdviceChange 'ǿ�Ƹ���ҽ������
                                vsAdvice.Row = i
                                
                                '�������������¼��б���λ�ˣ������ظ��ƶ���λ���
                                If Not Me.ActiveControl Is Nothing Then
                                    If Not Me.ActiveControl Is vsAdvice Then
                                        Call zlCommFun.PressKey(vbKeyTab)
                                    End If
                                End If
                                '����������������ƶ���λ��ҽ�������
                                blnDo = vsAdvice.RowData(i) <> 0
                                
                                Exit For
                            End If
                        Next
                        If i > vsAdvice.Rows - 1 Then
                            blnDo = True
                            cbsMain.FindControl(, conMenu_New, True, True).Execute
                        End If
                    ElseIf strSkip <> "" And InStr(";" & strSkip & ";", objNext.Name) = 0 Then
                        If objNext.Enabled And objNext.Visible Then
                            blnDo = True
                            objNext.SetFocus
                        End If
                    End If
                End If
            End If
        End If
    End If
    If Not blnDo Then
        Call zlCommFun.PressKey(vbKeyTab) '��Ȼ��λ
    Else
        SeekNextControl = True
    End If
End Function

Private Function GetInputSkip(ByVal lngRow As Long) As String
'���ܣ���ȡ����ҽ�������У��س����Ӧ�����Ŀؼ�
    Dim strSkip As String, lngFind As Long
    
    With vsAdvice
        'һ����ҩ�е�ҩƷ����ʱӦ����������
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And .RowData(lngRow) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                '��ҩ;��,����ִ��
                If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    If lngFind <> -1 Then
                        If Val(.TextMatrix(lngFind, COL_������ĿID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.txt�÷�.Name
                        End If
                        If Val(.TextMatrix(lngFind, COL_ִ�п���ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.cbo����ִ��.Name
                        End If
                    End If
                End If
                'Ƶ��
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" Then strSkip = strSkip & ";" & Me.txtƵ��.Name
                'ִ��ʱ��
                If .TextMatrix(lngRow, COL_ִ��ʱ��) <> "" Then strSkip = strSkip & ";" & Me.cboִ��ʱ��.Name
            End If
        ElseIf InStr(",C,D,F,G,Z", .TextMatrix(lngRow, COL_���)) > 0 And .RowData(lngRow) <> 0 And .TextMatrix(lngRow, COL_Ƶ��) = "һ����" Then
            strSkip = strSkip & ";" & Me.txtƵ��.Name
        End If
    End With
    GetInputSkip = Mid(strSkip, 2)
End Function

Private Sub SetBabyVisible(ByVal lng����id As Long)
'���ܣ����ݿ�����������Ӥ��ҽ���Ƿ����ѡ��
'˵�������Ʋ���Ӥ��ҽ��
    If DeptIsWoman(lng����id) Then
        lblӤ��.Visible = True
        cboӤ��.Visible = True
    Else
        Call zlControl.CboSetIndex(cboӤ��.hWnd, 0)
        cboӤ��.Tag = 0
        lblӤ��.Visible = False
        cboӤ��.Visible = False
    End If
End Sub

Private Sub CalcAdviceMoney()
'���ܣ������¿�ҽ�����
'˵����ֻ�ܵ�ǰ��ʾ���Ĳ����¿�ҽ��
    Dim dblMoney As Double, i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) And Val(.TextMatrix(i, COL_״̬)) = 1 Then
                dblMoney = dblMoney + Format(CCur(Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����))), gstrDec)
            End If
        Next
        stbThis.Panels(5).Text = "�¿�:" & FormatEx(dblMoney, 5) & "Ԫ"
    End With
End Sub

Private Sub AdviceSign()
'���ܣ���ҽ�����е���ǩ��
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lngǩ��id As Long, lng֤��ID As Long
    Dim intRule As Integer, strTimeStamp As String, strTimeStampCode As String
    Dim ColIDs As Collection, ColSource As Collection
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.����) Then
        MsgBox "����ǩ��֤���ѱ�ͣ�ã�����ϵ��Ϣ�ơ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Զ�����
    If mblnNoSave Then
        If Not CheckAdvice Then Exit Sub
        If Not SaveAdvice Then vsAdvice.SetFocus: Exit Sub
    End If
    
    '��ȡǩ��ҽ��Դ��
    intRule = ReadAdviceSignSource(1, mlng����ID, mstr�Һŵ�, strIDs, 0, False, strSource, mstrǰ��IDs, , ColIDs, ColSource)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "�ò���Ŀǰû�п���ǩ����ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    For i = 1 To ColIDs.Count
        strSign = gobjESign.Signature(ColSource(i), gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
            strSQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & ColIDs(i) & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
        End If
    Next
    If strSign <> "" Then
        '���¶�ȡ��ʾҽ��
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
        mblnOK = True
        If txtҽ������.Enabled Then
            txtҽ������.SetFocus
        Else
            vsAdvice.SetFocus
        End If

        MsgBox "����ɵ���ǩ����", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceTextChange(ByVal lngRow As Long) As Boolean
'���ܣ���ҽ����Ƭ�������ݱ仯ʱ���ж�ҽ�������ı��Ƿ�Ӧ��������֯
    Dim str��� As String, strText As String, blnDefine As Boolean
    
    With vsAdvice
        'ȷ��ҽ�����
        str��� = .TextMatrix(lngRow, COL_���)
        If str��� = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then '��ҩ�䷽��һ�����
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If lngRow <> -1 Then str��� = .TextMatrix(lngRow, COL_���)
        End If
        If str��� = "7" Then str��� = "8"
                
        'ȷ���Ƿ���
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "�������='" & str��� & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!ҽ������)) = "" Then
                blnDefine = False
            End If
        End If
        If blnDefine Then strText = mrsDefine!ҽ������
        
        '������ݱ䶯
        If blnDefine Then '�����ֶβ��ݻ���Թ�������Ĳ���
            If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" And InStr(strText, "[��ʼʱ��]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsDate(txt����ʱ��.Text) And txt����ʱ��.Tag <> "" Then
                If InStr(strText, "[����ʱ��]") > 0 Or InStr(strText, "[��Ѫʱ��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cboҽ������.Tag <> "" And InStr(strText, "[ҽ������]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                If InStr(strText, "[����Ƶ��]") > 0 Or InStr(strText, "[Ӣ��Ƶ��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cboִ��ʱ��.Tag <> "" And InStr(strText, "[ִ��ʱ��]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsNumeric(txt����.Text) And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
        End If
        
        Select Case str��� '��ͬ�������
        Case "5", "6" '������ҩ
            If Not blnDefine Then
                
            Else
                '[������][ͨ����][��Ʒ��][Ӣ����][���][����]��������޸�����ҩƷʱ�仯
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[��ҩ;��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "8" '��ҩ�䷽
            If Not blnDefine Then
                If IsNumeric(txt����.Text) And txt����.Tag <> "" Then AdviceTextChange = True: Exit Function
                If cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[�䷽���][�巨]��������޸������䷽ʱ�仯
                If IsNumeric(txt����.Text) And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[�÷�]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "C" '����
            If Not blnDefine Then
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[������Ŀ][����걾]��������޸�������Ŀʱ�仯
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[�ɼ�����]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "D" '���
            If Not blnDefine Then
                
            Else
                '[�����Ŀ][��鲿λ]��������޸�������Ŀʱ�仯
            End If
        Case "F" '����
            If Not blnDefine Then
                If IsDate(txt����ʱ��.Text) And txt����ʱ��.Tag <> "" Then AdviceTextChange = True: Exit Function
                If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[��Ҫ����][��������][������]��������޸�������Ŀʱ�仯
            End If
        Case "K" '��Ѫ
            If Not blnDefine Then
                If IsDate(txt����ʱ��.Text) And txt����ʱ��.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[��Ѫ;��]
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[��Ѫ;��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case Else '����
            If Not blnDefine Then
                
            Else
                '[������Ŀ]��������޸�������Ŀʱ�仯
            End If
        End Select
    End With
End Function

Private Function AdviceTextMake(ByVal lngRow As Long) As String
'���ܣ���ȡҽ�������ı�
'������lngRow=����ҽ�����ݵĿɼ���
    Dim rsTmp As New ADODB.Recordset
    Dim rsCard As New ADODB.Recordset
    Dim blnDefine As Boolean, str��� As String
    Dim strText As String, strSQL As String
    Dim strField As String, intƵ�ʷ�Χ As Integer
    Dim i As Long, k As Long
    Dim blnDo As Boolean
    Dim str��ҩ���� As String
    
    Dim str��ҩ As String, str�巨 As String, str��̬ As String
    Dim str���� As String, str���� As String
    Dim str���� As String, str�걾 As String
    Dim str��λ As String, str��λLast As String, str���� As String
    Dim dbl���� As Double
    Dim str��ҩ������ĿIDS As String, strSame As String
    
    On Error GoTo errH
    
    With vsAdvice
        'ȷ��ҽ�����
        str��� = .TextMatrix(lngRow, COL_���)
        If str��� = "E" Then '��ҩ�䷽��һ�����
            k = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If k <> -1 Then str��� = .TextMatrix(k, COL_���)
        End If
        If str��� = "7" Then str��� = "8"
                
        'ȷ���Ƿ���
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "�������='" & str��� & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!ҽ������)) = "" Then
                blnDefine = False
            End If
        End If
        
ReDoDefault: '���ڰ����幫ʽ����ʧ�ܣ����°�ȱʡ���������֯
        strText = ""
        If blnDefine Then strText = mrsDefine!ҽ������
        
        '����ҽ������
        Select Case str���
        Case "C" '����-------------------------------------------------------------
            str���� = "": str�걾 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_�����ĿID)) = 0 And mblnNewLIS Or Not mblnNewLIS Then
                        str���� = .TextMatrix(i, col_ҽ������) & "," & str����
                    End If
                    str�걾 = .TextMatrix(i, COL_�걾��λ)
                Else
                    Exit For
                End If
            Next
            If str���� = "" Then '�ϵķ�ʽ
                str���� = .TextMatrix(lngRow, COL_����)
            Else
                str���� = Left(str����, Len(str����) - 1)
            End If
            
            If Not blnDefine Then
                strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
            Else
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[����걾]") > 0 Then
                    strField = str�걾
                    strText = Replace(strText, "[����걾]", """" & strField & """")
                End If
                If InStr(strText, "[�ɼ�����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[�ɼ�����]", """" & strField & """")
                End If
            End If
        Case "D" '���-------------------------------------------------------------
            str��λ = "": str���� = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_�걾��λ) <> "" Then
                        If .TextMatrix(i, COL_�걾��λ) <> str��λLast And str��λLast <> "" Then
                            str��λ = str��λ & "," & str��λLast & IIF(str���� <> "", "(" & Mid(str����, 2) & ")", "")
                            str���� = ""
                        End If
                        If .TextMatrix(i, COL_��鷽��) <> "" Then
                            str���� = str���� & "," & .TextMatrix(i, COL_��鷽��)
                        End If
                        
                        str��λLast = .TextMatrix(i, COL_�걾��λ)
                    End If
                Else
                    Exit For
                End If
            Next
            If str��λLast <> "" Then
                str��λ = str��λ & "," & str��λLast & IIF(str���� <> "", "(" & Mid(str����, 2) & ")", "")
            End If
            str��λ = Mid(str��λ, 2) '��������Ŀ�Ĳ�λ
            
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_����) & _
                    Decode(Val(.TextMatrix(lngRow, COL_ִ�б��)), 1, ",����ִ��", 2, ",����ִ��", "") & IIF(str��λ <> "", ":" & str��λ, "")
            Else
                If InStr(strText, "[�����Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����) & _
                        Decode(Val(.TextMatrix(lngRow, COL_ִ�б��)), 1, ",����ִ��", 2, ",����ִ��", "")
                    strText = Replace(strText, "[�����Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[��鲿λ]") > 0 Then
                    strField = str��λ
                    strText = Replace(strText, "[��鲿λ]", """" & strField & """")
                End If
            End If
        Case "F" '����-------------------------------------------------------------
            str���� = "": str���� = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "G" Then
                        str���� = .TextMatrix(i, col_ҽ������)
                    Else
                        str���� = str���� & "," & .TextMatrix(i, col_ҽ������)
                    End If
                Else
                    Exit For
                End If
            Next
            str���� = Mid(str����, 2)
            
            If Not blnDefine Then
                If IsDate(.TextMatrix(lngRow, COL_�걾��λ)) Then
                    strText = Format(.TextMatrix(lngRow, COL_�걾��λ), "MM��dd��HH:mm")
                Else
                    strText = Format(.Cell(flexcpData, lngRow, COL_��ʼʱ��), "MM��dd��HH:mm")
                End If
                If str���� <> "" Then
                    strText = strText & IIF(str���� <> "", " �� " & str���� & " ���� ", " �� ")
                End If
                strText = strText & .TextMatrix(lngRow, COL_����) & IIF(.Cell(flexcpData, lngRow, COL_�걾��λ) = "", "", "(��λ:" & .Cell(flexcpData, lngRow, COL_�걾��λ) & ")")
                If str���� <> "" Then
                    strText = strText & " �� " & str����
                End If
            Else
                If InStr(strText, "[����ʱ��]") > 0 Then
                    If IsDate(.TextMatrix(lngRow, COL_����ʱ��)) Then
                        strField = .TextMatrix(lngRow, COL_����ʱ��)
                    Else
                        strField = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                    End If
                    strText = Replace(strText, "[����ʱ��]", """" & strField & """")
                End If
                If InStr(strText, "[��Ҫ����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����) & IIF(.Cell(flexcpData, lngRow, COL_�걾��λ) = "", "", "(��λ:" & .Cell(flexcpData, lngRow, COL_�걾��λ) & ")")
                    strText = Replace(strText, "[��Ҫ����]", """" & strField & """")
                End If
                If InStr(strText, "[��������]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[��������]", """" & strField & """")
                End If
                If InStr(strText, "[������]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[������]", """" & strField & """")
                End If
            End If
        Case "8" '��ҩ�䷽---------------------------------------------------------
            str��ҩ = "": str�巨 = "": str��ҩ������ĿIDS = "": strSame = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" Then
                        If InStr("," & str��ҩ������ĿIDS & ",", "," & Val(.TextMatrix(i, COL_������ĿID)) & ",") > 0 Then
                            strSame = strSame & "," & Val(.TextMatrix(i, COL_������ĿID))
                        End If
                        str��ҩ������ĿIDS = str��ҩ������ĿIDS & "," & Val(.TextMatrix(i, COL_������ĿID))
                    End If
                Else
                    Exit For
                End If
            Next
            strSame = Mid(strSame, 2)
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" Then
                        dbl���� = dbl���� + Val(.TextMatrix(i, COL_����))
                        
                        If Val(.TextMatrix(lngRow, COL_��ҩ��̬)) = 0 Then
                            blnDo = .TextMatrix(i, COL_�շ�ϸĿID) <> .TextMatrix(i - 1, COL_�շ�ϸĿID)
                        Else
                            blnDo = .TextMatrix(i, COL_������ĿID) <> .TextMatrix(i - 1, COL_������ĿID)
                        End If
                        
                        If blnDo Then
                            str��ҩ���� = .TextMatrix(i, col_ҽ������)
                            
                            If Val(.TextMatrix(lngRow, COL_��ҩ��̬)) = 0 And InStr("," & strSame & ",", "," & Val(.TextMatrix(i, COL_������ĿID)) & ",") > 0 Then
                                strSQL = "Select ��� as ���� From �շ���ĿĿ¼ Where ID=[1] And Exists(Select 1 From ҩƷ��� Where ҩƷID<>[1] And ҩ��ID=[2])"
                                Set rsTmp = New ADODB.Recordset '���Filter
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_�շ�ϸĿID)), Val(.TextMatrix(i, COL_������ĿID)))
                                If rsTmp.RecordCount > 0 Then
                                    If Not IsNull(rsTmp!����) Then str��ҩ���� = str��ҩ���� & "(" & rsTmp!���� & ")"
                                End If
                            End If
                        
                            str��ҩ = RTrim(str��ҩ���� & _
                                " " & FormatEx(dbl����, 5) & .TextMatrix(i, COL_������λ) & _
                                " " & .TextMatrix(i, COL_ҽ������)) & "," & str��ҩ
                            dbl���� = 0
                        End If
                    ElseIf .TextMatrix(i, COL_���) = "E" Then
                        str�巨 = .TextMatrix(i, col_ҽ������) & .TextMatrix(i, COL_�걾��λ)
                    End If
                Else
                    Exit For
                End If
            Next
            If str��ҩ <> "" Then
                str��ҩ = Mid(str��ҩ, 1, Len(str��ҩ) - 1)
            End If
            If Not blnDefine Then
                If .TextMatrix(lngRow, COL_��ҩ��̬) = "1" Then
                    str��̬ = "[��Ƭ]"
                ElseIf .TextMatrix(lngRow, COL_��ҩ��̬) = "2" Then
                    str��̬ = "[����]"
                End If
                '���ֺ���˿ո����ı����л��Զ�����
                strText = "��ҩ" & str��̬ & .TextMatrix(lngRow, COL_����) & "��," & _
                    .TextMatrix(lngRow, COL_Ƶ��) & "," & str�巨 & "," & _
                    .TextMatrix(lngRow, COL_�÷�) & ":" & str��ҩ
            Else
                If InStr(strText, "[����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[����]", """" & strField & """")
                End If
                If InStr(strText, "[�䷽���]") > 0 Then
                    strField = str��ҩ
                    strText = Replace(strText, "[�䷽���]", """" & strField & """")
                End If
                If InStr(strText, "[�÷�]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[�÷�]", """" & strField & """")
                End If
                If InStr(strText, "[�巨]") > 0 Then
                    strField = str�巨
                    strText = Replace(strText, "[�巨]", """" & strField & """")
                End If
            End If
        Case "4" '����------------------------------------------------------------
                strSQL = "Select ����,���,���� From �շ���ĿĿ¼ Where ID=[1]"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)))
                
                If Not blnDefine Then
                    strText = .TextMatrix(lngRow, COL_����)
                    If Not IsNull(rsTmp!���) Then
                        strText = strText & " " & rsTmp!���
                    End If
                Else
                    If InStr(strText, "[��������]") > 0 Then
                        strField = rsTmp!����
                        strText = Replace(strText, "[��������]", """" & strField & """")
                    End If
                    If InStr(strText, "[���]") > 0 Then
                        strField = Nvl(rsTmp!���)
                        strText = Replace(strText, "[���]", """" & strField & """")
                    End If
                    If InStr(strText, "[����]") > 0 Then
                        strField = Nvl(rsTmp!����)
                        strText = Replace(strText, "[����]", """" & strField & """")
                    End If
                End If
        Case "5", "6" '����ҩ���г�ҩ---------------------------------------------
            If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                '����:0-����,1-Ӣ����,3-��Ʒ��
                strSQL = "Select Nvl(B.����,A.����) as ����,A.���,A.����,B.����" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ���� B Where A.ID=B.�շ�ϸĿID(+) And A.ID=[1] Order by B.����,B.����"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)))
            ElseIf blnDefine Then
                '����:0-����,1-Ӣ����
                strSQL = "Select Nvl(B.����,A.����) as ����,Null as ���,Null as ����,B.����" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B Where A.ID=B.������ĿID(+) And A.ID=[1] Order by B.����,B.����"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_������ĿID)))
            End If
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_�걾��λ)
                If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                    If strText = "" Then
                        If gbytҩƷ������ʾ <> 0 Then rsTmp.Filter = "����=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strText = rsTmp!����
                    End If
                    If Not IsNull(rsTmp!����) Then
                        strText = strText & "(" & rsTmp!���� & ")"
                    End If
                    If Not IsNull(rsTmp!���) Then
                        strText = strText & " " & rsTmp!���
                    End If
                Else
                    If strText = "" Then
                        strText = .TextMatrix(lngRow, COL_����)
                    End If
                End If
            Else
                If InStr(strText, "[������]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�걾��λ)
                    If strField = "" Then
                        If gbytҩƷ������ʾ <> 0 Then rsTmp.Filter = "����=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strField = rsTmp!����
                    End If
                    strText = Replace(strText, "[������]", """" & strField & """")
                End If
                If InStr(strText, "[ͨ����]") > 0 Then
                    rsTmp.Filter = 0
                    strField = rsTmp!����
                    strText = Replace(strText, "[ͨ����]", """" & strField & """")
                End If
                If InStr(strText, "[��Ʒ��]") > 0 Then
                    rsTmp.Filter = "����=3"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!����
                    End If
                    strText = Replace(strText, "[��Ʒ��]", """" & strField & """")
                End If
                If InStr(strText, "[Ӣ����]") > 0 Then
                    rsTmp.Filter = "����=2"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!����
                    End If
                    strText = Replace(strText, "[Ӣ����]", """" & strField & """")
                End If
                If InStr(strText, "[���]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!���)
                    strText = Replace(strText, "[���]", """" & strField & """")
                End If
                If InStr(strText, "[����]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!����)
                    strText = Replace(strText, "[����]", """" & strField & """")
                End If
                If InStr(strText, "[��ҩ;��]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[��ҩ;��]", """" & strField & """")
                End If
            End If
        Case "K" '��Ѫҽ��
            If Not blnDefine Then
                If IsDate(.TextMatrix(lngRow, COL_��Ѫʱ��)) Then
                    strText = Format(.TextMatrix(lngRow, COL_��Ѫʱ��), "MM��dd��HH:mm")
                Else
                    strText = Format(.Cell(flexcpData, lngRow, COL_��ʼʱ��), "MM��dd��HH:mm")
                End If
            
                strText = "��" & strText & "��" & .TextMatrix(lngRow, COL_����)
                If .TextMatrix(lngRow, COL_�÷�) <> "" Then
                    strText = strText & "(" & .TextMatrix(lngRow, COL_�÷�) & ")"
                End If
            Else
                If InStr(strText, "[��Ѫʱ��]") > 0 Then
                    If IsDate(.TextMatrix(lngRow, COL_��Ѫʱ��)) Then
                        strField = .TextMatrix(lngRow, COL_��Ѫʱ��)
                    Else
                        strField = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                    End If
                    strText = Replace(strText, "[��Ѫʱ��]", """" & strField & """")
                End If
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[��Ѫ��Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[��Ѫ��Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[��Ѫ;��]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[��Ѫ;��]", """" & strField & """")
                End If
                If TypeName(.Cell(flexcpData, lngRow, COL_�������)) = "Recordset" Then
                    Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, lngRow, COL_�������))
                    If InStr(strText, "[Ѫ��]") > 0 Then
                        If rsCard.EOF Then
                            strField = ""
                        Else
                            strField = Decode(Val("" & rsCard!Ѫ��), 1, "A", 2, "B", 3, "O", 4, "AB", "")
                        End If
                        strText = Replace(strText, "[Ѫ��]", """" & strField & """")
                    End If
                    If InStr(strText, "[RH]") > 0 Then
                        If rsCard.EOF Then
                            strField = ""
                        Else
                            strField = Decode(Val("" & rsCard!RHD), 1, "-", 2, "+", "")
                        End If
                        strText = Replace(strText, "[RH]", """" & strField & """")
                    End If
                End If
                If InStr(strText, "[ִ�з���]") > 0 Then
                    strField = Val(.TextMatrix(lngRow + 1, COL_ִ�з���))
                    strText = Replace(strText, "[ִ�з���]", """" & strField & """")
                End If
            End If
        Case Else '�����������-----------------------------------------------------
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_����)
            Else
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
            End If
            '����ҽ��������ʾ
            If .TextMatrix(lngRow, COL_���) = "Z" And (Val(.TextMatrix(lngRow, COL_��������)) = 4 Or Val(.TextMatrix(lngRow, COL_��������)) = 14) Then
                strText = "������" & strText & "������"
            End If
        End Select
        
        '�����ֶλ���Թ���������ֶ�-------------------------------------------
        If blnDefine Then
            If InStr(strText, "[��ʼʱ��]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                strText = Replace(strText, "[��ʼʱ��]", """" & strField & """")
            End If
            If InStr(strText, "[ҽ������]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_ҽ������)
                If .TextMatrix(lngRow, COL_ҽ������) <> "" Then
                    If strField <> "" Then
                        strField = strField & "," & .TextMatrix(lngRow, COL_ҽ������)
                    Else
                        strField = .TextMatrix(lngRow, COL_ҽ������)
                    End If
                End If
                strText = Replace(strText, "[ҽ������]", """" & strField & """")
            End If
            If InStr(strText, "[����Ƶ��]") > 0 Then
                strField = .TextMatrix(lngRow, COL_Ƶ��)
                strText = Replace(strText, "[����Ƶ��]", """" & strField & """")
            End If
            If InStr(strText, "[Ӣ��Ƶ��]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" Then
                    intƵ�ʷ�Χ = GetƵ�ʷ�Χ(lngRow)
                    strSQL = "Select Ӣ������ From ����Ƶ����Ŀ Where ����=[1] And ���÷�Χ=[2]"
                    Set rsTmp = New ADODB.Recordset '���Filter
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, COL_Ƶ��), intƵ�ʷ�Χ)
                    If Not rsTmp.EOF Then strField = Nvl(rsTmp!Ӣ������)
                End If
                strText = Replace(strText, "[Ӣ��Ƶ��]", """" & strField & """")
            End If
            If InStr(strText, "[����]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_����) <> "" Then
                    strField = .TextMatrix(lngRow, COL_����) & .TextMatrix(lngRow, COL_������λ)
                End If
                strText = Replace(strText, "[����]", """" & strField & """")
            End If
            If InStr(strText, "[����]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_����) <> "" Then
                    strField = .TextMatrix(lngRow, COL_����) & .TextMatrix(lngRow, COL_������λ)
                End If
                strText = Replace(strText, "[����]", """" & strField & """")
            End If
            If InStr(strText, "[ִ��ʱ��]") > 0 Then
                strField = .TextMatrix(lngRow, COL_ִ��ʱ��)
                strText = Replace(strText, "[ִ��ʱ��]", """" & strField & """")
            End If
        End If
                
        '����ҽ������
        If blnDefine Then
            On Error Resume Next
            strText = mobjVBA.Eval(strText)
            If mobjVBA.Error.Number <> 0 Then
                err.Clear: On Error GoTo errH
                blnDefine = False: GoTo ReDoDefault
            End If
        End If
    End With
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetAdviceAppendItem() As String
'���ܣ���ȡδ����ҽ��(�¿����޸ĵ�)�����µ��ݸ���
'���أ���Ŀ��1<Split2>����1<Split1>��Ŀ��2<Split2>����2<Split1>...
    Dim arrAppend As Variant, i As Long, j As Long
    Dim strName As String, strText As String
    Dim strResult As String
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) <> 0 And .TextMatrix(i, COL_����) <> "" And .Cell(flexcpData, i, COL_����) = 1 Then
                arrAppend = Split(.TextMatrix(i, COL_����), "<Split1>")
                For j = 0 To UBound(arrAppend)
                    strName = Split(arrAppend(j), "<Split2>")(0)
                    strText = Split(arrAppend(j), "<Split2>")(3)
                    
                    If InStr(strResult, "<Split1>" & strName & "<Split2>") = 0 Then
                        strResult = strResult & "<Split1>" & strName & "<Split2>" & strText
                    End If
                Next
            End If
        Next
    End With
    
    GetAdviceAppendItem = Mid(strResult, Len("<Split1>") + 1)
End Function

Private Function GetAdviceDiagnosis() As String
'���ܣ���ȡ��ǰ��¼���δ��������
'���أ�"1�������һ 2���������"
    Dim strText As String, i As Long, j As Long
    
    With vsDiag
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, col���)) <> "" Then
                j = j + 1
                strText = strText & "  " & j & "��" & .TextMatrix(i, col���) & IIF(Val(.Cell(flexcpData, i, col����)) = 1, "(��)", "")
            End If
        Next
    End With
    
    strText = Mid(strText, 3)
    If j = 1 Then strText = Mid(strText, 3)
    GetAdviceDiagnosis = strText
End Function

Private Sub vsDiag_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiag
        If Col = col��� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '�ڵ���vsDiagXY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiag_KeyDown(vbKeyDelete, 0)
            End If
        End If
        If .Col = Col Then Call vsDiag_AfterRowColChange(-1, -1, Row, Col)
  
    End With
End Sub

Private Sub vsDiag_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsDiag
        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, col����) Is Nothing Then
                Set .Cell(flexcpPicture, i, col����) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, COLDEL) Is Nothing Then
               Set .Cell(flexcpPicture, i, COLDEL) = Nothing
            End If
        Next
        '���ñ༭�ɼ�����
        If Not DiagCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
            If NewCol = col��� Then
                .ComboList = "..."
            ElseIf NewCol = col���� Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = COLDEL Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            ElseIf NewCol = col��ҽ֤�� Then
                If .TextMatrix(NewRow, col���) = "" Then
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                Else
                    .ComboList = "..."
                End If
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '��ʾͼƬ
            If NewCol <> col���� And .TextMatrix(NewRow, col���) <> "" Then
                Set .Cell(flexcpPicture, NewRow, col����) = imgButtonNew.Picture
            End If
            '��ʾͼƬ
            If NewCol <> COLDEL Then
                Set .Cell(flexcpPicture, NewRow, COLDEL) = imgButtonDel.Picture
            End If
        End If
        
        If NewRow <> OldRow Then
            '��ǰ�б�־��ʾ
            Set .Cell(flexcpPicture, 0, col��־, .Rows - 1, col��־) = Nothing
            Set .Cell(flexcpPicture, NewRow, col��־) = img16.ListImages("���_��ǰ").Picture
            .Cell(flexcpPictureAlignment, NewRow, col��־) = 4
                    
            '��ʾ����ҽ����ʶ
            Call ShowDiagFlag(NewRow)
        End If
    End With
End Sub

Private Sub vsDiag_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '��StartEdit�¼��д������ʾ��ǰһ�еİ�ť
    If Col = col���� Then Cancel = True
End Sub

Private Sub vsDiag_Click()
    With vsDiag
        If (.MouseCol = col���� Or .MouseCol = COLDEL) And .MouseRow >= .FixedRows Then
            If .MouseCol = col���� Then
                If .TextMatrix(.MouseRow, col���) = "" Then Exit Sub
            End If
            
            .Select .MouseRow, .MouseCol
            Call vsDiag_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiag_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = col��־ Then
        Cancel = True
    ElseIf (NewCol = col��ҽ Or NewCol = COL��ҽ) And vsDiag.TextMatrix(NewRow, col���) <> "" Then
        Cancel = True
    End If
End Sub

Private Sub vsDiag_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str�Ա� As String
    Dim lngRow As Long, int������� As Integer
    Dim blnCancle As Boolean
    Dim str��� As String
    
    With vsDiag
        If Col = col��� Then
            If .Cell(flexcpData, Row, col��ҽ) = 1 Then
                If opt���(0).value Then
                    '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", mlng���˿���id, , True, False, , , 1)
                    str��� = "2"
                Else
                    'B-��ҽ��������
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", mlng���˿���id, mstr�Ա�, True, , , , 1)
                    str��� = "B"
                End If
            Else
                If opt���(0).value Then
                    '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", mlng���˿���id, , True, False, , , 1)
                    str��� = "1"
                Else
                    'D-ICD-10��������
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "D", mlng���˿���id, mstr�Ա�, True, , , , 1)
                    str��� = "D"
                End If
            End If
            If rsTmp Is Nothing Then
                If opt���(0).value Then
                    MsgBox "û�м���������ݿ���ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call SetDiagInput(Row, rsTmp, str���)
                Call DiagEnterNextCell
            End If
        ElseIf Col = col��ҽ֤�� Then
            If opt���(0).value Then
                '���������:�Ȳ��Ƿ��ж�Ӧ
                If Not Set��ҽ֤��(Row, Val(.TextMatrix(Row, col���ID))) Then
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng���˿���id, mstr�Ա�, True, , , , 1)
                Else
                    Exit Sub
                End If
            Else
                'Z-��ҽ��������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng���˿���id, mstr�Ա�, True, , , , 1)
            End If
            If Not rsTmp Is Nothing Then
                Call Set��ҽ֤��(Row, 0, rsTmp)
                Call DiagEnterNextCell
            End If
        ElseIf Col = col���� Then
            If .Rows < M_LNG_DIAGCOUNT Then
                lngRow = Row + 1: .AddItem "", lngRow
                .Row = lngRow: .Col = col���
                
                int������� = IIF(mbln��ҽ, 11, 1)
                If lngRow - 1 >= .FixedRows Then
                    int������� = IIF(.Cell(flexcpData, lngRow - 1, col��ҽ) = 1, 11, 1)
                End If
                Call SetDiagType(lngRow, int�������)
                
                Call SetDiagHeight
            End If
        ElseIf Col = COLDEL Then
            Call vsDiag_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiag_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsDiag
        If Col = col��� Then
            .TextMatrix(Row, col����) = IIF(.TextMatrix(Row, Col) <> "", "��", "")
        End If
    End With
End Sub

Private Sub vsDiag_DblClick()
    Call vsDiag_KeyPress(32)
End Sub

Private Sub vsDiag_GotFocus()
    If Me.Visible Then vsDiag_AfterRowColChange -1, -1, vsDiag.Row, vsDiag.Col
End Sub

Private Sub vsDiag_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnDo As Boolean, i As Long
    Dim int������� As Integer
    
    With vsDiag
        If KeyCode = vbKeyF4 Then
            If .Col = col��� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            '����Ƿ�����ɾ��
            For i = 1 To vsAdvice.Rows - 1
                If InStr("," & .TextMatrix(.Row, colҽ��ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_״̬) = "8" Then
                    MsgBox "����϶�Ӧ�Ĵ����ѷ��ͣ�����ɾ����", vbInformation, Me.Caption
                    Exit Sub
                End If
                'ҽ������վ����,����Ϲ���ҽ��,����ҽ���Ŀ���ҽ���ǵ�ǰ����Ա,������ɾ�����
                If mint���� = 2 And InStr("," & .TextMatrix(.Row, colҽ��ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_����ҽ��) <> UserInfo.���� Then
                    MsgBox "����ϴ��ڹ���ҽ��,�Ҹ�ҽ�������´����ɾ����", vbInformation, Me.Caption
                    Exit Sub
                End If
            Next
            blnDo = True
            If .TextMatrix(.Row, col���) <> "" Then
                If MsgBox("ȷʵҪɾ�����������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnDo = False
            End If
            If blnDo Then
                If .TextMatrix(.Row, col���) <> "" Then
                    lbl���.Tag = "1"
                    mblnNoSave = True
                    'ɾ����/��Ҫ��Ϻ������ҽӿ�
                    If CreatePlugInOK(p����ҽ���´�, mint����) Then
                        On Error Resume Next
                        Call gobjPlugIn.DiagnosisDeleted(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(.TextMatrix(.Row, col���ID)), .TextMatrix(.Row, col���), mint����)
                        Call zlPlugInErrH(err, "DiagnosisDeleted")
                        err.Clear: On Error GoTo 0
                    End If
                    If mblnPass Then
                        zlPassDrags
                    End If
                End If
                
                If .Rows = 2 And .Row = 1 Then
                    int������� = IIF(.Cell(flexcpData, .Row, col��ҽ) = 1, 11, 1)
                    .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, 0, .Row, .Cols - 1) = Empty
                    Set .Cell(flexcpPicture, .Row, 0, .Row, .Cols - 1) = Nothing
                    Call SetDiagType(.Row, int�������)
                Else
                    .RemoveItem .Row
                    Call SetDiagHeight
                End If
            End If
            .SetFocus
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiag_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiag_KeyPress(KeyAscii As Integer)
    With vsDiag
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = col���� Or .Col = col��ҽ Or .Col = COL��ҽ) Then
            KeyAscii = 0
            If .Col = col��ҽ Then
                If .Cell(flexcpData, .Row, col��ҽ) = 0 Then
                    Call SetDiagType(.Row, 11): .Col = col���
                End If
            ElseIf .Col = COL��ҽ Then
                If .Cell(flexcpData, .Row, COL��ҽ) = 0 Then
                    Call SetDiagType(.Row, 1): .Col = col���
                End If
            ElseIf .Col = col���� Then
                If DiagCellEditable(.Row, .Col) Then
                    KeyAscii = 0
                    .Cell(flexcpData, .Row, .Col) = IIF(.Cell(flexcpData, .Row, .Col) = 1, 0, 1)
                    .Cell(flexcpForeColor, .Row, .Col) = IIF(.Cell(flexcpData, .Row, .Col) = 1, vbRed, .GridColor)
                    
                    lbl���.Tag = "1"
                    mblnNoSave = True
                End If
            End If
        Else
            If .Col = col��� Or .Col = col��ҽ֤�� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiag_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiag_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiag_LostFocus()
    If vsDiag.Col <> col����ʱ�� Then vsDiag.Col = IIF(vsDiag.Col = col��ҽ֤��, col��ҽ֤��, col���)
End Sub

Private Sub vsDiag_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiag.EditSelStart = 0
    vsDiag.EditSelLength = zlCommFun.ActualLen(vsDiag.EditText)
End Sub

Private Sub vsDiag_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = col���� Or Col = col��ҽ Or Col = COL��ҽ Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Function GetDiagSQL(ByVal Row As Long, ByRef strInput As String, ByRef strSQL As String, ByRef str�Ա� As String, Optional ByVal strType As String) As String
'���ܣ���ò�ѯ��ϵ�SQL
'������strInput-��ѯ����,strsql--���ص�SQL��str�Ա�--���˵��Ա�  ,strType�����������ࡣ
'���أ�strsql--��ѯ��ҽ��ϵ�SQL
    If vsDiag.Cell(flexcpData, Row, col��ҽ) = 1 Then
        If opt���(0).value And strType <> "Z" Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
            Else
                strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
            End If
           strSQL = _
                " Select Distinct A.ID,A.ID as ��ĿID,A.����,Null as ���,A.����,A.˵��,A.����," & vbNewLine & _
                " Decode(b.����, [5], 1, Decode(b.����,[5],1,decode(a.����,[5],1,NULL))) As ����1ID,Decode(d.���id, Null, Decode(c.���id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                " Decode(Substr(b.����, 1, Length([5])), [5], 1, Decode(Substr(b.����, 1, Length([5])),[5],1,decode(Substr(a.����, 1, Length([5])),[5],1,NULL))) As ����3ID" & _
                " From �������Ŀ¼ A,������ϱ��� B, ������Ͽ��� C, ������Ͽ��� D" & _
                " Where A.ID=B.���ID And c.���id(+) = a.Id And d.���id(+) = a.Id And A.���=2" & _
                " And B.����=[4] And d.��Աid(+) = [6] And c.����id(+)=[7] And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                " And ( Nvl(A.���÷�Χ,0) = 0 or  A.���÷�Χ = 1) " & _
                " And (" & strSQL & ")" & _
                " Order by ����1ID, ����2ID, ����3ID,A.����"
                '����˳��������ȫƥ��(���ơ����롢���룩�������ղء�����ǿ����ղء�Ȼ������ƥ��(���ơ����롢���룩�������˫��ƥ��
        Else
            'B-��ҽ��������
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "A.���� Like [2]" '���뺺��ʱֻƥ������
            Else
                strSQL = "A.���� Like [1] Or A.���� Like [2] Or " & IIF(mint���� = 0, "A.����", "A.�����") & " Like [2]"
            End If
            strSQL = _
                "Select Distinct a.Id, a.Id As ��Ŀid, a.����, a.���, a.����, a.����," & IIF(mint���� = 0, "A.����", "A.����� as ����") & ", a.˵��," & _
                " Decode(a.����, [5], 1, Decode(" & IIF(mint���� = 0, "A.����", "A.�����") & ",[5],1,decode(a.����,[5],1,NULL))) As ����1ID," & vbNewLine & _
                "                Decode(d.����id, Null, Decode(c.����id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                "                Decode(Substr(a.����, 1, Length([5])), [5], 1, Decode(Substr(" & IIF(mint���� = 0, "A.����", "A.�����") & ", 1, Length([5])),[5],1,decode(Substr(a.����, 1, Length([5])),[5],1,NULL))) As ����3ID" & vbNewLine & _
                "From ��������Ŀ¼ A, ����������� C, ����������� D" & vbNewLine & _
                "Where a.��� = '" & IIF(strType = "", "B", strType) & "' And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.����id(+) = a.Id And" & vbNewLine & _
                "      d.����id(+) = a.Id And c.����id(+)=[7] And d.��Աid(+) = [6]" & vbNewLine & _
                IIF(str�Ա� <> "", " And (A.�Ա�����=[3] Or A.�Ա����� is NULL)", "") & _
                " And ( Nvl(A.���÷�Χ,0) = 0 or  A.���÷�Χ = 1) " & _
                " And (" & strSQL & ")" & _
                "Order By ����1ID, ����2ID, ����3ID, ����"
        End If
    Else
        If opt���(0).value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "B.���� Like [2]" '���뺺��ʱ,ֻƥ������
            Else
                strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
            End If
            strSQL = _
                " Select Distinct A.ID,A.ID as ��ĿID,A.����,Null as ���,A.����,A.˵��,A.����," & vbNewLine & _
                " Decode(b.����, [5], 1, Decode(b.����,[5],1,decode(a.����,[5],1,NULL))) As ����1ID,Decode(d.���id, Null, Decode(c.���id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                " Decode(Substr(b.����, 1, Length([5])), [5], 1, Decode(Substr(b.����, 1, Length([5])),[5],1,decode(Substr(a.����, 1, Length([5])),[5],1,NULL))) As ����3ID" & _
                " From �������Ŀ¼ A,������ϱ��� B, ������Ͽ��� C, ������Ͽ��� D" & _
                " Where A.ID=B.���ID And c.���id(+) = a.Id And d.���id(+) = a.Id And A.���=1" & _
                " And B.����=[4] And d.��Աid(+) = [6] And c.����id(+)=[7] And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                " And ( Nvl(A.���÷�Χ,0) = 0 or  A.���÷�Χ = 1) " & _
                " And (" & strSQL & ")" & _
                " Order by ����1ID, ����2ID, ����3ID,A.����"
                '����˳��������ȫƥ��(���ơ����롢���룩�������ղء�����ǿ����ղء�Ȼ������ƥ��(���ơ����롢���룩�������˫��ƥ��
        Else
            'D-ICD-10��������
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "A.���� Like [2]" '���뺺��ʱ,ֻƥ������
            Else
                strSQL = "A.���� Like [1] Or A.���� Like [2] Or " & IIF(mint���� = 0, "A.����", "A.�����") & " Like [2]"
            End If
            strSQL = _
                "Select Distinct a.Id, a.Id As ��Ŀid, a.����, a.���, a.����, a.����," & IIF(mint���� = 0, "A.����", "A.����� as ����") & ", a.˵��," & _
                " Decode(a.����, [5], 1, Decode(" & IIF(mint���� = 0, "A.����", "A.�����") & ",[5],1,decode(a.����,[5],1,NULL))) As ����1ID," & vbNewLine & _
                "                Decode(d.����id, Null, Decode(c.����id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                "                Decode(Substr(a.����, 1, Length([5])), [5], 1, Decode(Substr(" & IIF(mint���� = 0, "A.����", "A.�����") & ", 1, Length([5])),[5],1,decode(Substr(a.����, 1, Length([5])),[5],1,NULL))) As ����3ID" & vbNewLine & _
                "From ��������Ŀ¼ A, ����������� C, ����������� D" & vbNewLine & _
                "Where a.��� = 'D' And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.����id(+) = a.Id And" & vbNewLine & _
                "      d.����id(+) = a.Id And c.����id(+)=[7] And d.��Աid(+) = [6]" & vbNewLine & _
                IIF(str�Ա� <> "", " And (A.�Ա�����=[3] Or A.�Ա����� is NULL)", "") & _
                " And ( Nvl(A.���÷�Χ,0) = 0 or  A.���÷�Χ = 1) " & _
                " And (" & strSQL & ")" & _
                "Order By ����1ID, ����2ID, ����3ID, ����"

        End If
    End If
    GetDiagSQL = strSQL
End Function

Private Sub vsDiag_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As PointAPI
    Dim str�Ա� As String, int������� As Integer
    Dim str��� As String
    
    On Error GoTo errH
    
    With vsDiag
        If Col = col��� Or Col = col��ҽ֤�� Then
            If .EditText = "" Then
                If .TextMatrix(Row, col����) <> "" And Col = col��� Then
                    .EditText = .Cell(flexcpData, Row, Col)
                Else
                    '��ҽ֤���������������
                    If Col = col��ҽ֤�� Then
                        .Cell(flexcpData, Row, Col) = ""
                    End If
                End If
                If mblnReturn Then Call DiagEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiagEnterNextCell
            ElseIf Col = col��� And .TextMatrix(Row, col����) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                strSQL = GetDiagSQL(Row, strInput, strSQL, str�Ա�)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng���˿���id)
                If rsTmp.RecordCount = 1 Then
                    Call SetDiagInput(Row, rsTmp, rsTmp!��� & ""): .EditText = .Text
                Else
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, col���) = .EditText
                    lbl���.Tag = "1"
                    mblnNoSave = True
                End If
            ElseIf Col = col��� And .TextMatrix(Row, col����) <> "" And .Cell(flexcpData, Row, Col) <> "" And mblnFreeInput Then
                strInput = UCase(.EditText)
                strSQL = GetDiagSQL(Row, strInput, strSQL, str�Ա�)
                On Error GoTo errH
                vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(opt���(0).value, "�������", "��������"), False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng���˿���id, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    If rsTmp Is Nothing Then
                        .TextMatrix(Row, col���) = .EditText
                        lbl���.Tag = "1"
                        mblnNoSave = True
                    Else
                         Call SetDiagInput(Row, rsTmp, rsTmp!��� & ""): .EditText = .Text
                    End If
                End If
            Else
                int������� = Val(Mid(gstr�������, 1, 1))
                If int������� = 0 Then int������� = 1
                
                If mstr�Ա� Like "*��*" Then
                    str�Ա� = "��"
                ElseIf mstr�Ա� Like "*Ů*" Then
                    str�Ա� = "Ů"
                End If
                                
                strInput = UCase(.EditText)
                
                strSQL = GetDiagSQL(Row, strInput, strSQL, str�Ա�, IIF(Col = col���, "B", "Z"))
                If Col = col��� Then
                    If int������� = 1 And zlCommFun.IsCharChinese(strInput) Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng���˿���id)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                        End If
                        If Not rsTmp Is Nothing Then str��� = rsTmp!��� & ""
                        Call SetDiagInput(Row, rsTmp, str���): .EditText = .Text
                        If mblnReturn Then Call DiagEnterNextCell
                    Else
                        vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(opt���(0).value, "�������", "��������"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng���˿���id, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                        If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                            Cancel = True
                        Else
                            '���������뷽ʽ
                            If rsTmp Is Nothing And (int������� = 2 Or int������� = 3 And mint���� <> 0) Then
                                MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                                Cancel = True
                            ElseIf Not (rsTmp Is Nothing) Then
                                Call SetDiagInput(Row, rsTmp, rsTmp!��� & ""): .EditText = .Text
                                If mblnReturn Then Call DiagEnterNextCell
                            Else
                                'û��ƥ��ɹ��ٴε�������¼��
                                If int������� = 1 Or (int������� = 3 And (rsTmp Is Nothing) And mint���� = 0) Then
                                    Call SetDiagInput(Row, Nothing, str���): .EditText = .Text
                                    If mblnReturn Then Call DiagEnterNextCell
                                Else
                                    Cancel = True
                                End If
                            End If
                        End If
                    End If
                ElseIf Col = col��ҽ֤�� Then
                    If opt���(0).value Then
                        '���������:�Ȳ��Ƿ��ж�Ӧ
                        If Set��ҽ֤��(Row, Val(.TextMatrix(Row, col���ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng���˿���id, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (int������� = 2 Or int������� = 3 And mint���� <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call Set��ҽ֤��(Row, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col����ʱ�� Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    mblnCancle = False
                Else
                    MsgBox "��������ȷ�ķ���ʱ�䣬���磺""2012-12-21 00:00""��"
                    Cancel = True
                End If
            End If
            If .EditText <> .TextMatrix(Row, Col) Then mblnNoSave = True: lbl���.Tag = "1"
        End If
        mblnCancle = Cancel
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowDiagFlag(ByVal lngDiag As Long)
'���ܣ���ʾ���������Ӧ������ҽ�����
'������lngDiag=��ϱ��ǰ��
    Dim strALL As String, strCurr As String
    Dim lng��ID As Long, i As Long
        
    With vsDiag
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colҽ��ID) <> "" Then
                strALL = strALL & "," & .TextMatrix(i, colҽ��ID)
                If i = lngDiag Then
                    strCurr = .TextMatrix(i, colҽ��ID)
                End If
            End If
        Next
        strALL = Mid(strALL, 2)
    End With
    
    With vsAdvice
        .Redraw = flexRDNone
        Set .Cell(flexcpPicture, .FixedRows, COL_���, .Rows - 1, COL_���) = Nothing
        
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                'ÿ�ж����������в�����ʾ��һ����ҩ�����б�OwnerDraw
                lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), .RowData(i))
                If InStr("," & strCurr & ",", "," & lng��ID & ",") > 0 Then
                    Set .Cell(flexcpPicture, i, COL_���) = img16.ListImages("���_��ǰ").Picture
                ElseIf InStr("," & strALL & ",", "," & lng��ID & ",") > 0 Then
                    Set .Cell(flexcpPicture, i, COL_���) = img16.ListImages("���_����").Picture
                End If
            End If
        Next
        
        .Cell(flexcpPictureAlignment, .FixedRows, COL_���, .Rows - 1, COL_���) = 4
        .Redraw = flexRDDirect
    End With
End Sub

Private Function GetDiagRow(ByVal lngҽ��ID As Long) As Long
'���ܣ�����ָ����ҽ�����������������
'������lngҽ��ID=ҽ������ID
'���أ����û�й������򷵻�-1
    Dim i As Long
    
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col���) <> "" Then
                If InStr("," & .TextMatrix(i, colҽ��ID) & ",", "," & lngҽ��ID & ",") > 0 Then
                    GetDiagRow = i: Exit Function
                End If
            End If
        Next
    End With
    
    GetDiagRow = -1
End Function

Private Sub SetDiagFlag(ByVal lngAdvice As Long, ByVal intFlag As Integer, Optional ByVal lngDiag As Long = -1)
'���ܣ�����ָ��ҽ�����뵱ǰ����й���������ȡ��������еĹ���
'������lngAdvice=ҽ�����ǰ��
'      blnFlag=0-���,1-���
'      lngDiag=�Ƿ�����Ϊ��ָ������й�����û��ָ��ʱΪ-1��ʾ�뵱ǰ����й���
    Dim lngBegin As Long, lngEnd As Long
    Dim strҽ��ID As String, lng��ID As Long
    Dim blnDo As Boolean, i As Long
    
    If vsAdvice.RowData(lngAdvice) = 0 Then Exit Sub
    '������Ѿ����͵�ҽ����ֻ׼��������׼ȡ������
    If (intFlag = 0 Or intFlag = 1 And Not vsAdvice.Cell(flexcpPicture, lngAdvice, COL_���) Is Nothing) And vsAdvice.TextMatrix(lngAdvice, COL_״̬) = "8" Then Exit Sub
    
    With vsAdvice
        '������������
        lng��ID = IIF(Val(.TextMatrix(lngAdvice, COL_���ID)) <> 0, Val(.TextMatrix(lngAdvice, COL_���ID)), .RowData(lngAdvice))
        With vsDiag
            '����ȡ����ǰҽ�������κ�����еĹ���
            For i = 1 To .Rows - 1
                strҽ��ID = .TextMatrix(i, colҽ��ID)
                If InStr("," & strҽ��ID & ",", "," & lng��ID & ",") > 0 Then
                    strҽ��ID = Replace("," & strҽ��ID & ",", "," & lng��ID & ",", ",")
                    If Left(strҽ��ID, 1) = "," Then strҽ��ID = Mid(strҽ��ID, 2)
                    If Right(strҽ��ID, 1) = "," Then strҽ��ID = Mid(strҽ��ID, 1, Len(strҽ��ID) - 1)
                    .TextMatrix(i, colҽ��ID) = strҽ��ID
                    
                    If intFlag = 0 Then blnDo = True
                End If
            Next
            
            '������Ϊ�뵱ǰ����й���
            If intFlag = 1 Then
                If lngDiag = -1 Then lngDiag = .Row
                
                If .TextMatrix(lngDiag, col���) <> "" Then
                    strҽ��ID = .TextMatrix(lngDiag, colҽ��ID)
                    If InStr("," & strҽ��ID & ",", "," & lng��ID & ",") = 0 Then
                        strҽ��ID = strҽ��ID & "," & lng��ID
                        blnDo = True
                    End If
                    If Left(strҽ��ID, 1) = "," Then strҽ��ID = Mid(strҽ��ID, 2)
                    .TextMatrix(lngDiag, colҽ��ID) = strҽ��ID
                Else
                    'ָ��Ҫ����������������ʱ������Ϊ�޹�����Ч��
                    intFlag = 0
                End If
            End If
        End With
        
        '������ʾ�л�
        Call GetRowScope(lngAdvice, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If .RowData(i) <> 0 And Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                If intFlag = 1 Then
                    If lngDiag = vsDiag.Row Then
                        Set .Cell(flexcpPicture, i, COL_���) = img16.ListImages("���_��ǰ").Picture
                    Else
                        Set .Cell(flexcpPicture, i, COL_���) = img16.ListImages("���_����").Picture
                    End If
                    .Cell(flexcpPictureAlignment, i, COL_���) = 4
                ElseIf intFlag = 0 Then
                    Set .Cell(flexcpPicture, i, COL_���) = Nothing
                End If
            End If
        Next
        
        If blnDo Then
            lbl���.Tag = "1"
            mblnNoSave = True
        End If
    End With
End Sub

Private Function AdviceHaveDiag(ByVal lngAdvice As Long) As Long
'���ܣ��ж�ָ���е�ҽ���Ƿ��ѹ��������
'������lngAdvice=ҽ�����ǰ��
'���أ�ָ��ҽ�������������������У�����-1��ʾ�޹���
    Dim lng��ID As Long, i As Long
    
    AdviceHaveDiag = -1
    If vsAdvice.RowData(lngAdvice) = 0 Then Exit Function
    
    With vsAdvice
        lng��ID = IIF(Val(.TextMatrix(lngAdvice, COL_���ID)) <> 0, Val(.TextMatrix(lngAdvice, COL_���ID)), .RowData(lngAdvice))
    End With
    
    With vsDiag
        For i = 1 To .Rows - 1
            If InStr("," & .TextMatrix(i, colҽ��ID) & ",", "," & lng��ID & ",") > 0 Then
                AdviceHaveDiag = i: Exit Function
            End If
        Next
    End With
End Function

Private Sub SetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset, ByVal str��� As String)
'���ܣ����������Ŀ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI, i As Long
    Dim blnDo As Boolean
    Dim strTmp As String
    Dim lngԭ���id As Long '0 ��ʾ����ӵ���ϣ� ��Ϊ0��ʾ�޸���ϣ�lngԭ���id ��ֵ�����޸�ǰ�� ���ID�򼲲�ID
    Dim int������� As Integer
    
    On Error GoTo errH
    With vsDiag
        '����Ƿ������޸�
        For i = 1 To vsAdvice.Rows - 1
            If InStr("," & .TextMatrix(.Row, colҽ��ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_״̬) = "8" Then
                MsgBox "����϶�Ӧ�Ĵ����ѷ��ͣ������޸ġ�", vbInformation, Me.Caption
                Exit Sub
            End If
            
            'ҽ������վ����,����Ϲ���ҽ��,����ҽ���Ŀ���ҽ���ǵ�ǰ����Ա,�������޸����
            If mint���� = 2 And InStr("," & .TextMatrix(.Row, colҽ��ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_����ҽ��) <> UserInfo.���� Then
                MsgBox "����ϴ��ڹ���ҽ��,�Ҹ�ҽ�������´�����޸ġ�", vbInformation, Me.Caption
                Exit Sub
            End If
        Next
        If Not rsInput Is Nothing Then
            int������� = IIF(.Cell(flexcpFontBold, lngRow, COL��ҽ), 1, 11)
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    If .Rows > M_LNG_DIAGCOUNT Then Exit For
                    .AddItem "", lngRow + 1: lngRow = lngRow + 1
                    lngԭ���id = 0
                Else
                    lngԭ���id = Val(.TextMatrix(lngRow, col���ID))
                End If
                
                If InStr(.TextMatrix(lngRow, col���), "(") > 0 And InStr(.TextMatrix(lngRow, col���), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, col���), InStrRev(.TextMatrix(lngRow, col���), "("))
                End If
                .TextMatrix(lngRow, col���) = Nvl(rsInput!����) & strTmp
                
                '�������ȷ������,����ݼ���ȷ�����
                If opt���(0).value Then
                    .TextMatrix(lngRow, col���ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col��ϱ���) = rsInput!���� & ""
                    .TextMatrix(lngRow, col����ID) = ""
                    strSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
                Else
                    .TextMatrix(lngRow, col����ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col��������) = rsInput!���� & ""
                    .TextMatrix(lngRow, col�������) = str���
                    .TextMatrix(lngRow, col��������) = rsInput!���� & ""
                    .TextMatrix(lngRow, col���ID) = ""
                    strSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!��ĿID))
                If Not rsTmp.EOF Then
                    If opt���(0).value Then
                        .TextMatrix(lngRow, col����ID) = Nvl(rsTmp!ID)
                        strSQL = "Select ����,����,��� From ��������Ŀ¼ where id=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!ID)))
                        If Not rsTmp.EOF Then
                            .TextMatrix(lngRow, col��������) = rsTmp!���� & ""
                            .TextMatrix(lngRow, col�������) = rsTmp!��� & ""
                            .TextMatrix(lngRow, col��������) = rsTmp!���� & ""
                        End If
                    Else
                        .TextMatrix(lngRow, col���ID) = Nvl(rsTmp!ID)
                        strSQL = "Select ���� From �������Ŀ¼ where id=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!ID)))
                        If Not rsTmp.EOF Then .TextMatrix(lngRow, col��ϱ���) = rsTmp!���� & ""
                    End If
                End If
                
                '��ҽ���ݼ�����ϲο�ȡ֤��
                If .Cell(flexcpData, lngRow, col��ҽ) = 1 Then
                    Call Set��ҽ֤��(lngRow, Val(.TextMatrix(lngRow, col���ID)))
                End If
                
                .TextMatrix(lngRow, col����) = IIF(Not IsNull(rsInput!����), rsInput!����, "")
                .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
                
                .Cell(flexcpData, lngRow, col����) = 0
                .Cell(flexcpForeColor, lngRow, col����) = .GridColor
                
                .TextMatrix(lngRow, colICD��) = Nvl(rsInput!����)
                
                '������/��Ҫ��Ϻ������ҽӿ�
                If CreatePlugInOK(p����ҽ���´�, mint����) Then
                    On Error Resume Next
                    If lngRow = .FixedRows Then
                        Call gobjPlugIn.DiagnosisEnter(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(.TextMatrix(lngRow, col���ID)), .TextMatrix(lngRow, col���), lngԭ���id, mint����)
                        Call zlPlugInErrH(err, "DiagnosisEnter")
                    Else
                        Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(.TextMatrix(lngRow, col���ID)), .TextMatrix(lngRow, col���), lngԭ���id, mint����)
                        Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End If
                    err.Clear: On Error GoTo 0
                End If
                Call SetDiagType(lngRow, int�������)
                rsInput.MoveNext
            Next
            
            Call SetDiagHeight
        Else
            .TextMatrix(lngRow, col���) = .EditText
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
            
            .Cell(flexcpData, lngRow, col����) = 0
            .Cell(flexcpForeColor, lngRow, col����) = .GridColor
            
            .TextMatrix(lngRow, col����) = ""
            .TextMatrix(lngRow, col���ID) = ""
            .TextMatrix(lngRow, col����ID) = ""
            .TextMatrix(lngRow, colICD��) = ""
        End If
        
        lbl���.Tag = "1"
        mblnNoSave = True
    End With
    'PASS ���¼��
    If mblnPass Then
        zlPassDrags
    End If
    '������¼����δ������ϵ�ҽ���뵱ǰ��Ϲ���
    If vsDiag.TextMatrix(lngRow, col���) <> "" Then
        blnDo = False
        
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .RowData(i) <> 0 And Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                    If Val(.TextMatrix(i, COL_״̬)) = 1 And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                        If GetDiagRow(.RowData(i)) = -1 Then
                            Call SetDiagFlag(i, 1, lngRow)
                            blnDo = True
                        End If
                    End If
                End If
            Next
        End With
        
        If blnDo Then
            Call ShowDiagFlag(lngRow)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Set��ҽ֤��(ByVal lngRow As Long, ByVal lng���ID As Long, Optional ByVal rsInput As Recordset, Optional ByVal blnFreeInput As Boolean) As Boolean
'���ܣ���ҽ���ݼ�����ϲο�ȡ֤��
'������rsInput-�����Ϊ�գ������ָ������ҩ֤���¼��
'      blnFreeInput  true - ����¼��
'���أ��Ƿ��ж�Ӧ��ϵ
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strTmp As String
    
    With vsDiag
        'ȥ�����е�֤��
        If InStr(.TextMatrix(lngRow, col���), "(") > 0 And InStr(.TextMatrix(lngRow, col���), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, col���), 1, InStrRev(.TextMatrix(lngRow, col���), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, col���)
        End If
        If blnFreeInput Then
            .TextMatrix(lngRow, col֤��ID) = ""
            .TextMatrix(lngRow, col֤�����) = ""
            .TextMatrix(lngRow, col��ҽ֤��) = .EditText
            .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
            mblnNoSave = True
            Exit Function
        Else
            If rsInput Is Nothing Then
                If lng���ID <> 0 Then
                    strSQL = "Select Distinct a.֤����� as ID,a.֤��ID,a.֤������,b.���� as ֤�����" & _
                        " From ������ϲο� A,��������Ŀ¼ B" & _
                        " Where a.֤��ID=b.ID(+) And a.���ID=[1] And a.֤������ is Not NULL" & _
                        " Order by a.֤�����"
                    vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = Nothing
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng���ID)
                    If Not rsTmp Is Nothing Then
                        .TextMatrix(lngRow, col֤��ID) = Nvl(rsTmp!֤��id)
                        .TextMatrix(lngRow, col֤�����) = Nvl(rsTmp!֤�����)
                        If Not IsNull(rsTmp!֤������) Then
                            .TextMatrix(lngRow, col���) = strTmp
                            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
                            .TextMatrix(lngRow, col��ҽ֤��) = Nvl(rsTmp!֤������)
                            .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
                            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
                            mblnNoSave = True
                        End If
                        Set��ҽ֤�� = True
                    Else
                        If blnCancel Then
                            Set��ҽ֤�� = True
                            If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col��ҽ֤��)
                        Else
                            Set��ҽ֤�� = False
                        End If
                    End If
                Else
                    Set��ҽ֤�� = False
                End If
            Else
                .TextMatrix(lngRow, col֤��ID) = Nvl(rsInput!��ĿID)
                .TextMatrix(lngRow, col֤�����) = Nvl(rsInput!����)
                .TextMatrix(lngRow, col���) = strTmp
                .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
                .TextMatrix(lngRow, col��ҽ֤��) = Nvl(rsInput!����)
                .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
                If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
                mblnNoSave = True
            End If
        End If
    End With
End Function

Private Sub DiagEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiag
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, col���) To col����
                If DiagCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col���� Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            If txtҽ������.Enabled And txtҽ������.Visible Then
                txtҽ������.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Function DiagCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    Dim i As Long
    
    With vsDiag
        If .ColHidden(lngCol) Then Exit Function
        If .TextMatrix(lngRow, colҽ��ID) <> "" Then
            If lngCol = col��� Then
                For i = 1 To vsAdvice.Rows - 1
                    If InStr("," & .TextMatrix(lngRow, colҽ��ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_״̬) = "8" Then
                        Exit Function
                    End If
                    'ҽ������վ����,����Ϲ���ҽ��,����ҽ���Ŀ���ҽ���ǵ�ǰ����Ա,������ɾ�����
                    If mint���� = 2 And InStr("," & .TextMatrix(lngRow, colҽ��ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_����ҽ��) <> UserInfo.���� Then
                        Exit Function
                    End If
                Next
            End If
        End If
        '�������������
        If .TextMatrix(lngRow, col���) = "" Then
            If lngCol = col���� Or lngCol = col���� Or lngCol = col����ʱ�� Then
                Exit Function
            End If
        End If
        If lngCol = col���� Then Exit Function
        '���������������֤��
        If lngCol = col��ҽ֤�� Then
            If .TextMatrix(lngRow, col���) = "" Then Exit Function
            If .Cell(flexcpData, lngRow, col��ҽ) <> 1 Then Exit Function
        End If
    End With
    DiagCellEditable = True
End Function

Private Sub GetAgentInfo()
'���ܣ���ȡ��������Ϣ
    Dim rsTmp As ADODB.Recordset
    
    gstrSQL = "Select c.��Ϣ��, c.��Ϣֵ" & vbNewLine & _
                "  From ������Ϣ�ӱ� C" & vbNewLine & _
                "  Where c.����id = [2] And c.����id = [1] And Instr(',����������,���������֤��,�������֤��,',','||c.��Ϣ��||',')>0"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    
    AgentInfo.���ξ�����¼�� = False
    AgentInfo.���������� = ""
    AgentInfo.���������֤�� = ""
    If rsTmp.EOF Then Exit Sub
    
    AgentInfo.���ξ�����¼�� = True
    While Not rsTmp.EOF
        Select Case Nvl(rsTmp!��Ϣ��)
            Case "����������"
                AgentInfo.���������� = Nvl(rsTmp!��Ϣֵ)
            Case "���������֤��"
                AgentInfo.���������֤�� = Nvl(rsTmp!��Ϣֵ)
        End Select
        rsTmp.MoveNext
    Wend
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDiagHeight()
    vsDiag.Height = vsDiag.Rows * vsDiag.RowHeightMin + IIF(mbytSize = 0, 2, 12) * Screen.TwipsPerPixelY
    fra���.Height = vsDiag.Height + 4 * Screen.TwipsPerPixelY + IIF(mbytSize = 0, 50, 0)
    Call Form_Resize
End Sub



Private Sub vsAdvice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    'Pass
    If Button = 2 Then
        With vsAdvice
            lngRow = .MouseRow
            If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
                If Not .RowHidden(lngRow) Then .Row = lngRow
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim blnDo As Boolean
    
    If Button = 2 Then
        If cbsMain Is Nothing Then Exit Sub
        Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If mblnPass = False Then
            blnDo = True
        Else
            blnDo = gobjPass.PassType = 2 Or gobjPass.PassType = 4 Or (gobjPass.PassType = 1 And gobjPass.PassVersion = "4.0")
            '����༭����˵�������ҽ��վ��һ��
            If Not blnDo Then
                If gobjPass.zlPassCheck(mobjPassMap) Then
                    Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, objPopup.CommandBar, conMenu_Edit_MediAudit)
                End If
            End If
        End If
        If gobjPlugIn Is Nothing And blnDo And gobjDrugExplain Is Nothing Then Exit Sub '������û�в˵���Ŀʱ����ʾһ���հ�С����
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
    
End Sub

Private Sub ExePlugIn(ByVal strName As String)
'���ܣ�ִ����ҹ���
    Dim lngID As String
    
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        With vsAdvice
            lngID = .RowData(.Row)
            If InStr(",1,2,", Val(.TextMatrix(.Row, COL_EDIT))) > 0 Then
                If MsgBox("��ǰѡ�е�ҽ��δ���棬�Ƿ��ȱ��棿", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                cbsMain.FindControl(, conMenu_Save, True, True).Execute: Exit Sub
            End If
        End With
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, p����ҽ���´�, strName, mlng����ID, mlng�Һ�ID, lngID, mlngǰ��ID, mint����)
        Call zlPlugInErrH(err, "ExecuteFunc")
        err.Clear: On Error GoTo 0
    End If
End Sub

Private Function CheckInHosAdvice() As Boolean
'���ܣ���鵱ǰ�����Ƿ������Ч�����ۻ�סԺҽ��
'������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, intDays As Integer
    
    On Error GoTo errH
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 And Not .RowHidden(i) Then
                If .TextMatrix(i, COL_EDIT) = "1" And i <> .Row Then
                    If .TextMatrix(i, COL_��������) = "1" Or .TextMatrix(i, COL_��������) = "2" Then
                        CheckInHosAdvice = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
     
    strSQL = "Select Trunc(Sysdate - ��Ժ����) ���� From ������ҳ Where ����id = [1] And ��ҳid = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTmp.RecordCount > 0 Then
        intDays = IIF(gint��ͨ�Һ����� > gint����Һ�����, gint��ͨ�Һ�����, gint����Һ�����)
        If intDays = 0 Then intDays = 1
        If rsTmp!���� <= intDays Then   'ҽ������ʱ��ɾ��ԤԼ�Ǽǣ�71009��
            If MsgBox("���ھɵ�ԤԼ���룬�����µ�ԤԼ����ʱ����Ҫɾ���ɵģ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                CheckInHosAdvice = True
                Exit Function
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceRSAddRow(ByRef rsAdvice As ADODB.Recordset, ByVal i As Long)
'���ܣ�����ָ���е���Ϣ���Ʋ���ҽ����¼����һ������
    Dim lng��ID As Long, lng����ID As Long, j As Long
    
    '��ȡ����ID
    With vsDiag
        lng��ID = IIF(Val(vsAdvice.TextMatrix(i, COL_���ID)) = 0, vsAdvice.RowData(i), Val(vsAdvice.TextMatrix(i, COL_���ID)))
        For j = .FixedRows To .Rows - 1
            If InStr("," & .TextMatrix(j, colҽ��ID) & ",", "," & lng��ID & ",") > 0 Then
                lng����ID = Val(.TextMatrix(j, col����ID))
                Exit For
            End If
        Next
    End With
    
    With vsAdvice
        rsAdvice.AddNew
        rsAdvice!ID = .RowData(i)
        rsAdvice!���ID = Val("" & .TextMatrix(i, COL_���ID))
        rsAdvice!ǰ��ID = mlngǰ��ID
        rsAdvice!������Դ = 1
        rsAdvice!����ID = mlng����ID
        rsAdvice!�Һŵ� = mstr�Һŵ�
        
        rsAdvice!Ӥ�� = Val("" & .TextMatrix(i, COL_Ӥ��))
        rsAdvice!���� = mstr����
        rsAdvice!�Ա� = mstr�Ա�
        rsAdvice!���� = mint����
        rsAdvice!���˿���id = mlng���˿���id
        
        rsAdvice!��� = .TextMatrix(i, COL_���)
        rsAdvice!ҽ��״̬ = .TextMatrix(i, COL_״̬)
        rsAdvice!ҽ����Ч = 1
        rsAdvice!������� = .TextMatrix(i, COL_���)
        rsAdvice!������ĿID = Val("" & .TextMatrix(i, COL_������ĿID))
        rsAdvice!�걾��λ = .TextMatrix(i, COL_�걾��λ)
        rsAdvice!��鷽�� = .TextMatrix(i, COL_��鷽��)
        
        rsAdvice!�շ�ϸĿID = Val("" & .TextMatrix(i, COL_�շ�ϸĿID))
        rsAdvice!���� = Val("" & .TextMatrix(i, COL_����))
        rsAdvice!�������� = Val("" & .TextMatrix(i, COL_����))
        rsAdvice!�ܸ����� = Val("" & .TextMatrix(i, COL_����))
        rsAdvice!ҽ������ = .TextMatrix(i, col_ҽ������)
        rsAdvice!ҽ������ = .TextMatrix(i, COL_ҽ������)
        rsAdvice!ִ�п���ID = Val("" & .TextMatrix(i, COL_ִ�п���ID))
        rsAdvice!ִ��Ƶ�� = .TextMatrix(i, COL_Ƶ��)
        rsAdvice!Ƶ�ʴ��� = Val("" & .TextMatrix(i, COL_Ƶ�ʴ���))
        rsAdvice!Ƶ�ʼ�� = Val("" & .TextMatrix(i, COL_Ƶ�ʼ��))
        rsAdvice!�����λ = .TextMatrix(i, COL_�����λ)
        rsAdvice!ִ��ʱ�䷽�� = .TextMatrix(i, COL_ִ��ʱ��)
        rsAdvice!�Ƽ����� = Val("" & .TextMatrix(i, COL_�Ƽ�����))
        rsAdvice!ִ������ = Val("" & .TextMatrix(i, COL_ִ������))
        rsAdvice!ִ�б�� = Val("" & .TextMatrix(i, COL_ִ�б��))
                    
        rsAdvice!�ɷ���� = Val("" & .TextMatrix(i, COL_�ɷ����))
        rsAdvice!������־ = Val("" & .TextMatrix(i, COL_��־))
        rsAdvice!��ʼִ��ʱ�� = .TextMatrix(i, COL_��ʼʱ��)
        rsAdvice!��������id = Val("" & .TextMatrix(i, COL_��������ID))
        rsAdvice!����ҽ�� = .TextMatrix(i, COL_����ҽ��)
        rsAdvice!����ʱ�� = CDate(.TextMatrix(i, COL_����ʱ��))
        rsAdvice!ժҪ = .Cell(flexcpData, i, COL_ҽ������)
        rsAdvice!����id = lng����ID
        rsAdvice!EditState = Val(.TextMatrix(i, COL_EDIT)) '1-������2���޸�
        rsAdvice!��ҩĿ�� = Val("" & .TextMatrix(i, COL_��ҩĿ��))     '1-Ԥ��,2-����
        rsAdvice!��ҩ���� = .TextMatrix(i, COL_��ҩ����)
        rsAdvice.Update
    End With
End Sub

Private Function zlPluginAdviceEnter(ByVal lngRow As Long) As Boolean
'���ܣ�����ҽ����ɺ������ҽӿ�
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim rsAdvice As ADODB.Recordset
    
    Set rsAdvice = GetAdviceRs
    
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    For i = lngBegin To lngEnd
        Call AdviceRSAddRow(rsAdvice, i)
    Next
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    On Error Resume Next
    zlPluginAdviceEnter = gobjPlugIn.AdviceEnter(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, rsAdvice, mint����)
    If err.Number <> 0 And zlPluginAdviceEnter = False Then zlPluginAdviceEnter = True
    Call zlPlugInErrH(err, "AdviceEnter")
    err.Clear: On Error GoTo 0
End Function

Private Function zlPluginAdviceSave() As Boolean
'���ܣ�ҽ������ǰ������ҽӿ�
    Dim i As Long
    Dim rsAdvice As ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim rsTmp As ADODB.Recordset
    Set rsAdvice = GetAdviceRs
    With vsAdvice
        'ҽ��¼�����֮��ֱ�ӱ�����ܵ�����ǰ����ҽ��û�е��õ���� AdviceEditAfter�˴��ٵ���һ�Ρ�
        i = .Row
        If .RowData(i) <> 0 And (Val(.TextMatrix(i, COL_EDIT)) = 1 Or Val(.TextMatrix(i, COL_EDIT)) = 2) Then
            Set rsTmp = zlDatabase.zlCopyDataStructure(rsAdvice)
            Call GetRowScope(i, lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                Call AdviceRSAddRow(rsTmp, i)
            Next
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        For i = .FixedRows To .Rows - 1
            '�������޸ĵ���Ч��(�ǿ���)
            If .RowData(i) <> 0 And (.TextMatrix(i, COL_EDIT) = "2" Or .TextMatrix(i, COL_EDIT) = "1") Then
                Call AdviceRSAddRow(rsAdvice, i)
            End If
        Next
    End With
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    On Error Resume Next
    If Not (rsTmp Is Nothing) Then
        If rsTmp.RecordCount > 0 Then
            Call gobjPlugIn.AdviceEditAfter(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, rsTmp, mint����)
            Call zlPlugInErrH(err, "AdviceEditAfter")
        End If
    End If
    zlPluginAdviceSave = gobjPlugIn.AdviceSave(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, rsAdvice, mint����)
    If err.Number <> 0 And zlPluginAdviceSave = False Then zlPluginAdviceSave = True
    Call zlPlugInErrH(err, "AdviceSave")
    err.Clear: On Error GoTo 0
End Function

Private Sub zlPluginAdviceRowChange(ByVal lngRow As Long, Optional ByVal intType As Integer)
'���ܣ�ҽ���л��к������ҽӿ�
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim rsAdvice As ADODB.Recordset
    
    If Val(vsAdvice.RowData(lngRow)) = 0 Then Exit Sub
    Set rsAdvice = GetAdviceRs
    
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    For i = lngBegin To lngEnd
        Call AdviceRSAddRow(rsAdvice, i)
    Next
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    On Error Resume Next
    If intType = 0 Then
        Call gobjPlugIn.AdviceRowChange(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, rsAdvice, mint����)
        Call zlPlugInErrH(err, "AdviceRowChange")
    Else
        Call gobjPlugIn.AdviceEditAfter(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, rsAdvice, mint����)
        Call zlPlugInErrH(err, "AdviceEditAfter")
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Function CheckBackNo(ByVal str�Һŵ� As String) As Boolean
'���ܣ�����Ƿ��Ѿ��˺ţ������Ƿ��Ѿ�ȡ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From ���˹Һż�¼ Where ִ��״̬ In (0, -1) And NO = [1] And ��¼����=1 And ��¼״̬=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceSendCheck", str�Һŵ�)
    If rsTmp.RecordCount > 0 Then
        MsgBox "���ܲ���û�о���Ĳ��ˡ�", vbInformation, "����ҽ���༭"
        Exit Function
    End If
    
    strSQL = "Select ID From ������ü�¼ Where ��¼����=4 And ��¼״̬=2 And NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceSendCheck", str�Һŵ�)
    If rsTmp.RecordCount > 0 Then
        MsgBox "���ܲ����Ѿ��˺ŵĲ��ˡ�", vbInformation, "����ҽ���༭"
        Exit Function
    End If
    CheckBackNo = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Set����ʱ��(ByVal strType As String)
'���ܣ�������������Ѫҽ���İ���ʱ��λ��
'������strType=F-����,K-��Ѫ

    If strType = "F" Then
        lbl����ʱ��.Caption = "����ʱ��"
        lbl����ʱ��.Top = lbl�÷�.Top
        txt����ʱ��.Top = txt�÷�.Top
    
        lbl����ʱ��.Left = lblҽ������.Left
        txt����ʱ��.Left = txt�÷�.Left
    Else
        lbl����ʱ��.Caption = "��Ѫʱ��"
        lbl����ʱ��.Top = lbl��ʼʱ��.Top
        txt����ʱ��.Top = txt��ʼʱ��.Top
    
        lbl����ʱ��.Left = lblҽ������.Left
        txt����ʱ��.Left = cboҽ������.Left
        
        'SetItemEditable����������ʾ������ʱ�䡱�Ͳ���ʾ���÷���
        lbl�÷�.Visible = True
        txt�÷�.Visible = True
        cmd�÷�.Visible = True
    End If
    
    cmd����ʱ��.Top = txt����ʱ��.Top + 30
    cmd����ʱ��.Left = txt����ʱ��.Left + txt����ʱ��.Width - cmd����ʱ��.Width - 30
End Sub


Private Sub SetCboִ������(ByVal bln���Ա�ҩ As Boolean, ByVal bln�ٴ��Թ�ҩ As Boolean)
    cboִ������.Clear
    
    If bln�ٴ��Թ�ҩ Then
        cboִ������.AddItem "1-�Ա�ҩ"
    Else
        cboִ������.AddItem "0-����"
        If bln���Ա�ҩ Then cboִ������.AddItem "1-�Ա�ҩ"
        cboִ������.AddItem "2-��Ժ��ҩ"
    End If
End Sub


Private Sub txt��ҩ����_GotFocus()
    zlControl.TxtSelAll txt��ҩ����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ҩ����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ҩ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ҩ����.Text <> "" And vsAdvice.TextMatrix(vsAdvice.Row, COL_���) <> "K" Then
            If ReasonSelect(txt��ҩ����.Text, 1) Then Exit Sub
        End If
        If SeekNextControl Then Call txt��ҩ����_Validate(False)
    End If
End Sub

Private Sub txt��ҩ����_Change()
    txt��ҩ����.Tag = "1"
End Sub

Private Sub txt��ҩ����_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(txt��ҩ����.Text) > 1000 Then
        MsgBox "�������ݲ������� 500 �����ֻ� 1000 ���ַ���", vbInformation, gstrSysName
        txt��ҩ����_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub cboDruPur_Click()
    lbl��ҩĿ��.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cboDruPur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SeekNextControl
    End If
End Sub

Private Sub SetRow��־ͼ��(ByVal i As Long, Optional ByVal bytMode As Byte)
'���ܣ����ݵ�ǰ�е�״̬�����ñ�־�е�ͼ����ʾ
'������i=��ǰ��
'      bytMode=һ����ҩʱ���룬0-���ݴ����е�״̬�����������У�1-����������������ʱ�Ŵ���,2-ֻ��������
    Dim blnFirst As Boolean, lngRow As Long
    Dim intͼ���� As Integer 'ҽ�����������ͼ�����
    
    With vsAdvice
        '����¼��
        If Val(.TextMatrix(i, COL_������ĿID)) = 0 Then
             Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("����").Picture
             .Cell(flexcpPictureAlignment, i, COL_F��־) = 4
        Else
            blnFirst = True
            lngRow = i
            'һ����ҩ��ͼ��ֻ��ʾ�ڵ�һ��(���״̬����)
            If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                If bytMode > 0 Then
                    If bytMode = 1 Then
                        '�жϴ���������������ʱ�Ŵ���(���ܻ�û�����ø�ҩ;��)
                        If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(i - 1, COL_���ID)) And Val(.TextMatrix(i, COL_���ID)) <> 0 Then blnFirst = False
                    Else
                        'ֻ���������(ȡ��һ��ʱ)
                    End If
                Else
                    '���ݴ����е�״̬������������
                    lngRow = .FindRow(.TextMatrix(i, COL_���ID), , COL_���ID)
                End If
            End If
        
            If blnFirst Then
                If .TextMatrix(i, COL_��־) = "2" Then
                    Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("��¼").Picture
                ElseIf .TextMatrix(i, COL_��־) = "1" Then
                    Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, COL_F��־) = Nothing
                End If
            
                'һ����ҩ��ʾ�ڵ�һ��
                If Val(.TextMatrix(i, COL_״̬)) < 2 Then   '�¿����ݴ��ҽ��
                    Select Case Val(.TextMatrix(i, COL_���״̬))
                    '0-������ˣ�1-����ˣ�2-���ͨ����3-���δͨ��
                        Case 1
                            If .TextMatrix(i, COL_���) = "K" And Val(.TextMatrix(i, COL_��鷽��)) = 1 Then
                                '��Ѫҽ�����ͼ�굥����ʾ(��������ҽ���˶�)
                                Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("�˶�").Picture
                            Else
                                Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                            End If
                        Case 2
                            If Not (.TextMatrix(i, COL_���) = "K" And Val(.TextMatrix(i, COL_��鷽��)) = 1) Then
                                Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("���ͨ��").Picture
                            End If
                        Case 3
                            Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("���δͨ��").Picture
                        Case 4, 5
                            If gblnѪ��ϵͳ = False Then Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                        Case 7
                            Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("��ǩ��").Picture
                    End Select
                End If
                                '�������ϵͳ
                If .TextMatrix(i, COL_�������״̬) = "0" Then
                    Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                ElseIf .TextMatrix(i, COL_�������״̬) = "2" Or .TextMatrix(i, COL_���������) = "1" Then
                    '��ʱ�������ϸ���
                    Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("���ͨ��").Picture
                ElseIf .TextMatrix(i, COL_���������) = "2" Then
                    ' ���ϸ�
                    Set .Cell(flexcpPicture, lngRow, COL_F��־) = frmIcons.imgFlag.ListImages("���δͨ��").Picture
                End If
                .Cell(flexcpPictureAlignment, lngRow, COL_F��־) = 4
            End If
            If Val(.TextMatrix(i, COL_ǩ����)) = 1 Then
                Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgSign.ListImages("ǩ��").Picture
                intͼ���� = 1
            End If
            
            If Val(.TextMatrix(i, COL_��ΣҩƷ)) > 0 Then
                If .Cell(flexcpPicture, i, col_ҽ������) Is Nothing Then
                    Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture
                    intͼ���� = 1
                Else
                    If .Cell(flexcpPicture, i, col_ҽ������) <> frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture Then
                        pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, i, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                        pictmp.PaintPicture frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                        Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                        intͼ���� = 2
                    End If
                End If
            End If
            'Σ��ֵͼ��
            If mlngΣ��ֵID > 0 And Val(.TextMatrix(i, COL_EDIT)) = 1 Or Val(.TextMatrix(i, COL_Σ��ֵID)) > 0 Then
                If intͼ���� = 0 Then
                    Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture
                ElseIf intͼ���� = 1 Then
                    pictmp.Cls
                    pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                    Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                    intͼ���� = 2
                ElseIf intͼ���� = 2 Then
                    pictmp.Cls
                    pictmp.Width = 720
                    pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, 480, pictmp.Height
                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture, 480, 0, 240, pictmp.Height
                    Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                    pictmp.Width = 480
                    intͼ���� = 3
                End If
            End If
        End If
    End With
End Sub

Private Sub ShowOrHideQuestion()
'���ܣ���ʾ�����ؿ�����ҩ���δͨ����˵��
    Dim strMsg As String
    
    If lbl����.Caption <> "" Then lbl����.Caption = ""
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)) = 3 Then
        strMsg = GetKSSAuditQuestion(Val(vsAdvice.RowData(vsAdvice.Row)))
        If strMsg <> "" Then lbl����.Caption = "��˷�����" & strMsg
        
    End If
    pic����.Visible = lbl����.Caption <> ""
    
    Call Form_Resize
End Sub

Private Sub ReSet���״̬ͼ��(ByVal lngRow As Long)
'���ܣ�ɾ�����޸�һ����ҩ�е�һ��ҽ���������������е�ͼ��Ϳ���ҩ���е����״̬
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim blnDo As Boolean
    
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    
    '��������δͨ���ģ�ɾ�����Ϊ�����
    With vsAdvice
        For i = lngBegin To lngEnd
            If gblnKSSStrict And UserInfo.��ҩ���� < Val(.TextMatrix(i, COL_�����ȼ�)) And .TextMatrix(i, COL_��־) <> "1" Then
                .TextMatrix(i, COL_���״̬) = 1
            Else
                .TextMatrix(i, COL_���״̬) = ""   'CheckAdvice�лὫһ�������Ϊ��ͬ
            End If
            If .TextMatrix(i, COL_���״̬) <> "" Then
                Set .Cell(flexcpPicture, lngBegin, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                .Cell(flexcpPictureAlignment, lngBegin, COL_F��־) = 4
                blnDo = True
            End If
        Next
        If blnDo = False Then Set .Cell(flexcpPicture, lngBegin, COL_F��־) = Nothing
    End With
End Sub


Private Sub SetMediInfoItem(ByVal bln��ʾ���� As Boolean, ByVal bln��ʾ����ҩ�� As Boolean)
'���ܣ����ػ���ʾҩƷ��ص���Ϣ��Ŀ���״�����������˵������ҩĿ�ģ���ҩ����
    Dim lngHeight As Long, lngHeightOld As Long
    Dim bytHideType As Byte
        
    lngHeightOld = fraAdvice.Height
    lngHeight = cbo����ִ��.Top + cbo����ִ��.Height + 60
    
    If bln��ʾ���� Or bln��ʾ����ҩ�� Then
        If bln��ʾ����ҩ�� = False Then
            lngHeight = lngHeight + txt����˵��.Height + 90
        Else
            lngHeight = lngHeight + txt����˵��.Height + 90 + txt��ҩ����.Height + 90
        End If
    End If
    
    If lngHeightOld <> lngHeight Then
        fraAdvice.Height = lngHeight
        '���������bug
        fraAdvice.Tag = "������"
        Call cbsMain_Resize
        '��Ҫ�������β�����Ч
        fraAdvice.Tag = ""
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub SetFontSize(ByVal bytSize As Byte)
'���ܣ����н��������ͳһ����
'������bytSize  0-9�����壬1-12������
    Call SetPublicFontSize(Me, bytSize)
    Call zlControl.VSFSetFontSize(vsAdvice, IIF(bytSize = 0, 9, 12))
    Call zlControl.VSFSetFontSize(vsDiag, IIF(bytSize = 0, 9, 12))
    Call SetCtrlPos
End Sub

Private Sub SetCtrlPos()
'���ܣ����ÿؼ�λ��,��ע�����пؼ�λ�õ����þ������ڴ˺���������
    Dim lngDistance1 As Long, lngDistance2 As Long
    Dim lngHeight As Long
    
    lngDistance1 = 30: lngDistance2 = 180
    cmdLastDiag.Top = lbl���.Top + lbl���.Height + 60
    Call SetCtrlPosOnLine(True, -1, lbl���, 60, cmdLastDiag)
    vsDiag.Left = cmdLastDiag.Left + cmdLastDiag.Width + lngDistance1
    opt���(0).Left = vsDiag.Left + vsDiag.Width + lngDistance1 * 2
    opt���(1).Left = opt���(0).Left
    fra���.Height = vsDiag.Top + vsDiag.Height + 120
    
    '��ֱ���ÿؼ�λ��
    lbl��ʼʱ��.Left = 120
    '��������ߵ�һ�пؼ�����������ұ߿ؼ�λ�ã����������߿ؼ�λ�ã�Ŀ��Ϊ������ҽ�����ݸ߶�
    '��ߵ�һ�ţ�Ϊ�˵�λ��һ���ұ߿ؼ�λ��
    Call SetCtrlPosOnLine(False, 0, lbl��ʼʱ��, lngDistance1, txt��ʼʱ��, -1 * cmd��ʼʱ��.Width, cmd��ʼʱ��, lngDistance2, chk����, lngDistance2, chk����)
    lbl����.Left = chk����.Left + chk����.Width + 5 * lngDistance2
    
    '�����ұ߿ؼ�λ��
    Call SetCtrlPosOnLine(True, 1, lbl����, lngDistance2, lblҽ������, lngDistance2, lblִ��ʱ��, lngDistance2, lblִ�п���, lngDistance2, lbl����ִ��, lngDistance2, lbl����˵��)
    Call SetCtrlPosOnLine(False, 0, lbl����, lngDistance1, cbo����, lngDistance1, lbl���ٵ�λ, lngDistance2, chkZeroBilling)
    Call SetCtrlPosOnLine(False, 0, lblҽ������, lngDistance1, cboҽ������, -1 * cmdҽ������.Width, cmdҽ������, lngDistance1, cmd��������)
    Call SetCtrlPosOnLine(False, 0, lblִ��ʱ��, lngDistance1, cboִ��ʱ��)
    Call SetCtrlPosOnLine(False, 0, lblִ�п���, lngDistance1, cboִ�п���, lngDistance2, lblִ������, lngDistance1, cboִ������)
    Call SetCtrlPosOnLine(False, 0, lbl����ִ��, lngDistance1, cbo����ִ��)
    Call SetCtrlPosOnLine(False, 0, lbl����˵��, lngDistance1, txt����˵��, -1 * cmdExcReason.Width, cmdExcReason, lngDistance1, cmdComExcReason)
    cboִ��ʱ��.Width = cmd��������.Left + cmd��������.Width - cboִ��ʱ��.Left
    cboִ������.Width = cmd��������.Left + cmd��������.Width - cboִ������.Left
    txt����˵��.Width = cmdҽ������.Left + cmdҽ������.Width - txt����˵��.Left - 80
    
    '������߿ؼ�λ��
    txtҽ������.Height = txt��ʼʱ��.Height 'Ϊ���в�����
    Call SetCtrlPosOnLine(True, 1, lbl��ʼʱ��, lngDistance2, lblҽ������, lblִ�п���.Top - lblҽ������.Top - lblҽ������.Height, lbl����, lngDistance2, lbl�÷�, -1 * lbl����ʱ��.Height, lbl����ʱ��, lngDistance2, lbl����, lngDistance2 + 10, lbl��ҩĿ��)
    Call SetCtrlPosOnLine(False, 0, lblҽ������, lngDistance1, txtҽ������, lngDistance1, cmdExt)
    cmdSel.Top = cmdExt.Top + cmdExt.Height + 120
    cmdSel.Left = cmdExt.Left
    Call SetCtrlPosOnLine(False, 0, cmdSel, lngDistance1, picHelp)
    txtҽ������.Height = cboִ�п���.Top + cboִ�п���.Height - txtҽ������.Top
    Call SetCtrlPosOnLine(False, 0, lbl����, lngDistance1, cbo����)
    Call SetCtrlPosOnLine(False, 0, lbl����ʱ��, -1 * lbl�÷�.Width, lbl�÷�, lngDistance1, txt����ʱ��, -1 * cmd����ʱ��.Width, cmd����ʱ��, -1 * txt�÷�.Width, txt�÷�, -1 * cmd�÷�.Width, cmd�÷�, lngDistance2, lblƵ��, lngDistance1, txtƵ��, -1 * cmdƵ��.Width, cmdƵ��)
    txtƵ��.Width = txtҽ������.Width + txtҽ������.Left - txtƵ��.Left
    cmdƵ��.Left = txtҽ������.Width + txtҽ������.Left - cmdƵ��.Width
    Call SetCtrlPosOnLine(False, 0, lbl����, lngDistance1, txt����, lngDistance1, lbl������λ, -1 * lbl����.Width, lbl����, -1 * (txt����.Width + Me.TextWidth("��")), txt����, lngDistance2, lbl����, lngDistance1, txt����, lngDistance1, lbl������λ)
    lbl��ҩĿ��.Top = IIF(mbytSize = 0, lbl��ҩĿ��.Top - 50, lbl��ҩĿ��.Top)
    Call SetCtrlPosOnLine(False, 0, lbl��ҩĿ��, lngDistance1, cboDruPur, lngDistance2, lbl��ҩ����, lngDistance1, txt��ҩ����, -1 * cmdReason.Width, cmdReason, 0.5 * lngDistance1, cmd�ղ���ҩ����)
    lbl������λ.Left = txtҽ������.Width + txtҽ������.Left - lbl������λ.Width
    txt����.Width = lbl������λ.Left - lngDistance1 - txt����.Left
    
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        fraAdvice.Width = cmd��������.Left + cmd��������.Width + 500
        Me.Width = fraAdvice.Width + fraAdvice.Left
    End If
End Sub

Private Sub CheckDrugOutOfRange(ByVal lngRow As Long, ByVal sngDays As Single)
'���ܣ����ҩƷ�����������Ƿ񳬹�����ķ�Χ,���ҵ���ѡ����ʾ
'���أ�ѡ���Ƿ����
    Dim blnReturn As Boolean
    Dim strOld As String
    
    strOld = vsAdvice.TextMatrix(lngRow, COL_�Ƿ���)

    If sngDays > IIF(mbytPatiType = 1, conOrdinary, conEmergency) And vsAdvice.TextMatrix(lngRow, COL_����˵��) = "" Then
        vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = "1"
    Else
        vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) = ""
    End If
    If strOld <> vsAdvice.TextMatrix(lngRow, COL_�Ƿ���) Then
        lbl����˵��.Tag = "1"
    End If
End Sub


Private Sub Set��ҩ�����Ƿ���(ByVal lngRow As Long)
'���ܣ����ݸ���ҽ���е�������������Ƶ����Ϣ�����������������á��Ƿ��ڡ��е�ֵ
    Dim sng���� As Single
    
    With vsAdvice
        If RowIn�䷽��(lngRow) And Val(.TextMatrix(lngRow, COL_����)) > 0 Then
            sng���� = Val(.TextMatrix(lngRow, COL_����))
        Else
            If mbln���� Then
                sng���� = Val(.TextMatrix(lngRow, COL_����))
            ElseIf Val(.TextMatrix(lngRow, COL_����)) <> 0 And Val(.TextMatrix(lngRow, COL_����)) <> 0 _
                And .TextMatrix(lngRow, COL_Ƶ��) <> "" And Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) <> 0 _
                And Val(.TextMatrix(lngRow, COL_����ϵ��)) <> 0 And Val(.TextMatrix(lngRow, COL_�����װ)) <> 0 Then
                
                sng���� = CalcȱʡҩƷ����(Val(.TextMatrix(lngRow, COL_����)), Val(.TextMatrix(lngRow, COL_����)), _
                    Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), .TextMatrix(lngRow, COL_�����λ), _
                    Val(.TextMatrix(lngRow, COL_����ϵ��)), Val(.TextMatrix(lngRow, COL_�����װ)), _
                    Val(.TextMatrix(lngRow, COL_�ɷ����)))
            End If
        End If
        
        If sng���� > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
            .TextMatrix(lngRow, COL_�Ƿ���) = "1"
        Else
            .TextMatrix(lngRow, COL_�Ƿ���) = ""
        End If
        
    End With
End Sub

Private Function ReGetҩƷ����(ByVal dblԭ���� As Double, ByVal dbl���� As Double, ByVal sng���� As Long, ByVal lngRow As Long) As Double
'���ܣ����¸��ݵ�������������ǰ�е�Ƶ����Ϣ�ȼ�������
'���أ����������
    Dim dbl���� As Double
    
    ReGetҩƷ���� = dblԭ����
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_Ƶ��) <> "" And Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)) <> 0 And Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) <> 0 _
            And dbl���� <> 0 And Val(.TextMatrix(lngRow, COL_����ϵ��)) <> 0 And Val(.TextMatrix(lngRow, COL_�����װ)) <> 0 Then
            
            dbl���� = FormatEx(CalcȱʡҩƷ����(dbl����, sng����, _
                Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), _
                .TextMatrix(lngRow, COL_�����λ), .TextMatrix(lngRow, COL_ִ��ʱ��), _
                Val(.TextMatrix(lngRow, COL_����ϵ��)), Val(.TextMatrix(lngRow, COL_�����װ)), _
                Val(.TextMatrix(lngRow, COL_�ɷ����))), 5)
                
            
            If InStr(GetInsidePrivs(p����ҽ���´�), "ҩƷС������") = 0 Then
                dbl���� = IntEx(dbl����)
            ElseIf Val(.TextMatrix(lngRow, COL_�ɷ����)) <> 0 Then
                dbl���� = IntEx(dbl����)
            End If
        End If
    End With
    ReGetҩƷ���� = dbl����
End Function

Private Sub Setҽ������(ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ�Ϊָ����Χ�ڵ�ҽ�����ó���� .TextMatrix(i, COL_�Ƿ���) .TextMatrix(i, COL_����˵��)
    Dim dbl���� As Double
    Dim lng���ID As Long
    Dim i As Long
    Dim j As Long
    
    With vsAdvice
        For i = lngBegin To lngEnd
            If .RowData(i) = 0 Then Exit Sub
            .TextMatrix(i, COL_�Ƿ���) = ""
            If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 And Val(.TextMatrix(i, COL_��������)) > 0 Then
                dbl���� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_�����װ)) * Val(.TextMatrix(i, COL_����ϵ��))
                If dbl���� > Val(.TextMatrix(i, COL_��������)) Then .TextMatrix(i, COL_�Ƿ���) = "1"
                If .TextMatrix(i, COL_�Ƿ���) = "" And .TextMatrix(i, COL_�Ƿ���) = "" Then .TextMatrix(i, COL_����˵��) = ""
            ElseIf Val(.TextMatrix(i, COL_���)) = 7 And Val(.TextMatrix(i, COL_��������)) > 0 Then  '��ҩ
                dbl���� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����))
                If dbl���� > Val(.TextMatrix(i, COL_��������)) Then .TextMatrix(i, COL_�Ƿ���) = "1"
            ElseIf .TextMatrix(i, COL_���) = "E" And .TextMatrix(i, COL_��������) = "4" Then
                lng���ID = .RowData(i)
                For j = i - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        If .TextMatrix(j, COL_�Ƿ���) = "1" Then
                            .TextMatrix(i, COL_�Ƿ���) = "1"
                            .TextMatrix(j, COL_�Ƿ���) = ""
                        End If
                    Else
                        Exit For
                    End If
                Next
            ElseIf Val(.TextMatrix(i, COL_��������)) > 0 Then '����������Ŀ
                If Val(.TextMatrix(i, COL_����)) > Val(.TextMatrix(i, COL_��������)) Then .TextMatrix(i, COL_�Ƿ���) = "1"
            End If
            If .TextMatrix(i, COL_�Ƿ���) = "" And .TextMatrix(i, COL_�Ƿ���) = "" Then .TextMatrix(i, COL_����˵��) = ""
        Next
    End With
End Sub

Private Function GetInsureStr(ByRef strIDs1 As String, ByRef strIDs2 As String, ByRef strҽ������ As String, ByVal lngRow As Long) As Boolean
'���ܣ���ȡҽ��������ַ���
'   strIDs1:ҩƷ���ĵ��շ�ϸĿID�ַ�����strIDs2 ������������Ŀ��������ĿID:ִ�п����ַ���
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    With vsAdvice
        'Ϊ��������,��Union��ʽ
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                If InStr(",4,5,6,7,", .TextMatrix(i, COL_���)) > 0 Then
                    'ҩƷ�������޶�Ӧ��ϵ,ҩƷֻ��������´�ʱ
                    If Val(.TextMatrix(i, COL_�շ�ϸĿID)) <> 0 And InStr("," & strIDs1 & ",", "," & Val(.TextMatrix(i, COL_�շ�ϸĿID)) & ",") = 0 Then
                        strIDs1 = strIDs1 & "," & .TextMatrix(i, COL_�շ�ϸĿID)
                    End If
                ElseIf InStr("," & strIDs2 & ",", "," & Val(.TextMatrix(i, COL_������ĿID)) & ",") = 0 Then
                    '�������շ�����Ϊ0��
                    strIDs2 = strIDs2 & "," & Val(.TextMatrix(i, COL_������ĿID)) & ":" & Val(.TextMatrix(i, COL_ִ�п���ID))
                End If
            End If
        Next
        strҽ������ = Left(vsAdvice.TextMatrix(lngRow, col_ҽ������), 50)
    End With
End Function

Private Function SetAll����˵��(ByVal lngRow As Long, ByRef blnOut As Boolean) As Boolean
'���ܣ����û��д����˵����ҽ��Ȼ��Ϊ����ӳ���˵��
'������lngRow��ʼ�к�
'      blnOut ��ʾ������������浫û��д�κ�����
    Dim i As Long
    Dim strTmp As String
    Dim str�䷽ As String
    Dim str��IDs As String '��Ҫ��д����˵����ҽ���к�
    Dim str����˵�� As String
    Dim strMsg As String
    Dim varArr As Variant
    
    With vsAdvice
        For i = lngRow To .Rows - 1
            If (.TextMatrix(i, COL_�Ƿ���) = "1" Or .TextMatrix(i, COL_�Ƿ���) = "1") And .TextMatrix(i, COL_����˵��) = "" Then
                Select Case (Val(.TextMatrix(i, COL_�Ƿ���)) - Val(.TextMatrix(i, COL_�Ƿ���)))
                    Case 1
                        strTmp = "�����˴�������������д����˵����"
                    Case -1
                        strTmp = "��������ҩ�Ƴ�(" & IIF(mbytPatiType = 1, conOrdinary, conEmergency) & "��)������д����˵����"
                    Case 0
                        strTmp = "�����˴�����������ҩ�Ƴ�(" & IIF(mbytPatiType = 1, conOrdinary, conEmergency) & "��)������д����˵����"
                End Select
                If .TextMatrix(i, COL_���) = "E" And .TextMatrix(i, COL_��������) = "4" Then '�в�ҩ�ĳ�����־�ŵ�����ʾ�У�ֻ�ܴ����ʾ
                    str�䷽ = .TextMatrix(i, col_ҽ������)
                    str�䷽ = "�в�ҩ�䷽��" & Mid(str�䷽, InStr(str�䷽, ":") + 1)
                    strMsg = strMsg & """" & str�䷽ & """" & strTmp
                Else
                    strMsg = strMsg & """" & .TextMatrix(i, col_ҽ������) & """" & strTmp
                End If
                str��IDs = str��IDs & "," & i
            End If
        Next
    End With
    str��IDs = Mid(str��IDs, 2)
    
    If strMsg <> "" And str��IDs <> "" Then
        Call frmMsgDruExcess.ShowMe(Me, strMsg, str����˵��)
    Else
        SetAll����˵�� = True
        Exit Function
    End If
    
    If str����˵�� = "*NULL*" Then 'str����˵�� = "*NULL*" ��ʾû����д�κζ���
        blnOut = True
    ElseIf str����˵�� <> "" Then  'str����˵�� = "" ��ʾû��ִ�� frmMsgDruExcess.ShowMe(Me, strMsg, str����˵��)
        varArr = Split(str��IDs, ",")
        For i = 0 To UBound(varArr)
            vsAdvice.TextMatrix(Val(varArr(i)), COL_����˵��) = str����˵��
        Next
    End If
    SetAll����˵�� = True
End Function

Private Function CheckAutoMerge(ByVal lngRow As Long) As Integer
'���ܣ��Զ��ж���ý�����Զ�һ����ҩ��ȡ��һ��
        '�ж���ý���Զ�һ��/ȡ��һ��
        '�����һ��ҽ������ҺҩƷ����ǰ�����Ҳ����ҺҩƷʱ��
        '1�������ǰ¼�����ҩƷ�����ǰ��һ��Һ���У���ý�ǵ�һ������ô�Զ�Ϊ��һ����״̬��
        '2�������ǰ¼�����ҩƷ�����ǰ��һ��Һ���У���ý�����һ������һ��ҽ��������ô�Զ�ȡ����һ����״̬����������ý���������
        '3�������ǰ��¼�����ý��ҩƷ����.��ý�������ж����һ��Һ�����Ѿ�����һ����ýʱ�����Զ�ȡ����һ����״̬����������ý���������
'������lngRow=��ǰ��
'���أ�0-��������1-һ����ҩ��2-ȡ��һ����ҩ
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long
    
    With vsAdvice
        lngPreRow = GetPreRow(lngRow) 'ȡ��һ��Ч��,ĳЩ����ȱʡ����һ����ͬ
        If lngPreRow <> -1 Then
            If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 Then
                i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_���ID)), lngPreRow + 1)
                If .TextMatrix(i, COL_���) = "E" And .TextMatrix(i, COL_��������) = "2" And .TextMatrix(i, COL_ִ�з���) = "1" Then
                    j = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    If .TextMatrix(j, COL_���) = "E" And .TextMatrix(j, COL_��������) = "2" And .TextMatrix(j, COL_ִ�з���) = "1" Then
                        If RowCanMerge(lngPreRow, lngRow) Then
                            '��ȡ��һ��ҽ���Ŀ�ʼ�к�
                            Call GetRowScope(i, lngBegin, lngEnd)
                            If mblnRowMerge = False Then
                                '��һ������ý(1)
                                If Val(.TextMatrix(lngBegin, COL_�Ƿ���ý)) = 1 And Val(.TextMatrix(lngRow, COL_�Ƿ���ý)) <> 1 Then
                                    CheckAutoMerge = 1
                                End If
                            Else
                                '��ǰ������ý��ǰ���Ѿ�������ý��(3)
                                If Val(.TextMatrix(lngRow, COL_�Ƿ���ý)) = 1 Then
                                    For i = lngBegin To lngEnd
                                        If Val(.TextMatrix(i, COL_�Ƿ���ý)) = 1 Then
                                            CheckAutoMerge = 2
                                            Exit For
                                        End If
                                    Next
                                Else
                                    '��ǰ��ҩƷ����һ������ý�����ֲ��ǵ�һ��ʱ(2)
                                    If Val(.TextMatrix(lngPreRow, COL_�Ƿ���ý)) = 1 And lngPreRow <> lngBegin Then
                                        CheckAutoMerge = 2
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'
    If mblnNoRefresh Then Exit Sub
    If mblnNoSave Then
        mblnNoRefresh = True
        tbcSub.Item(0).Selected = True
        MsgBox "��ǰҽ�����ݱ༭����δ���棬���ȱ��档", vbInformation, gstrSysName
        mblnNoRefresh = False
        Exit Sub
    End If
    Call SubWinRefreshData(Item)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'���ܣ�ˢ������
    If objItem.Tag = "ҽ���༭" Then
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
    Else
        If objItem.Tag <> "" Then
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, pסԺҽ���´�, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng����ID, mstr�Һŵ�, 0, False, _
                    0, 0, 0, mlng���˿���id)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End If
    End If
End Sub

Private Function CanUseApply(ByVal str��� As String, Optional ByVal lng��Ŀid As Long, Optional ByVal str��Ŀ���� As String) As Boolean
'���ܣ��Ƿ����ʹ����Ӧ�����뵥
    'ֻ����ҽ������վ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
        Dim str���� As String
    Dim strEx As String
 
    If mint���� <> 0 Then Exit Function
    
    If str��� = "D" And Val(Mid(gstrOutUseApp, 1, 1)) = 1 Or _
        str��� = "C" And Val(Mid(gstrOutUseApp, 2, 1)) = 1 Or _
        str��� = "K" And Val(Mid(gstrOutUseApp, 3, 1)) = 1 Then
        blnTmp = True
    End If
    
    If lng��Ŀid <> 0 And blnTmp Then
        If str��� = "D" Then
            strSQL = "select 1 from ��������Ӧ�� where ������ĿID=[1] and Ӧ�ó��� = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, 1)
            If rsTmp.EOF Then blnTmp = False
        ElseIf str��� = "C" Then
            blnTmp = False
            If Not gobjLIS Is Nothing Then
                If str��Ŀ���� <> "" Then
                    str���� = str��Ŀ����
                Else
                    strSQL = "select b.���� from ������ĿĿ¼ b where b.id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid)
                    str���� = rsTmp!���� & ""
                End If
                On Error Resume Next
                blnTmp = gobjLIS.CanUseLISApp(str����, strEx)
                err.Clear: On Error GoTo 0
            End If
        End If
    End If
    CanUseApply = blnTmp
End Function

Private Function CheckApply() As Boolean
'���ܣ��ж����뵥����д���
    Dim i As Long
    If Val(gstrOutUseApp) > 0 Then
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    If Val(.TextMatrix(i, COL_�������)) < 0 And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        .Row = i
                        Call cmdExt_Click
                    End If
                End If
            Next
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    If Val(.TextMatrix(i, COL_�������)) < 0 And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        Exit Function
                    End If
                End If
            Next
        End With
    End If
    CheckApply = True
End Function

Private Function ApplySelect() As Integer
'���ܣ��������뵥ѡ����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intTmp As Integer, blnCancel As Boolean
    Dim vRect As RECT
    
    If mint���� <> 0 Then Exit Function
 
    intTmp = Val(Mid(gstrOutUseApp, 1, 1))
    If intTmp = 1 Then strSQL = "select 1 as id,'������뵥' as ���뵥 from dual"
    intTmp = Val(Mid(gstrOutUseApp, 2, 1))
    If intTmp = 1 And Not gobjLIS Is Nothing Then strSQL = IIF(strSQL = "", "", strSQL & " union all ") & "select 2 as id,'�������뵥' as ���뵥 from dual"
    intTmp = Val(Mid(gstrOutUseApp, 3, 1))
    If intTmp = 1 Then strSQL = IIF(strSQL = "", "", strSQL & " union all ") & "select 3 as id,'��Ѫ���뵥' as ���뵥 from dual"
    
    If strSQL = "" Then Exit Function
    
    strSQL = strSQL & " order by id"

    On Error GoTo errH
    vRect = GetControlRect(txtҽ������.hWnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, Me.Caption, , , , , , True, vRect.Left, vRect.Top, txtҽ������.Height, blnCancel, , True)
    If Not rsTmp Is Nothing Then ApplySelect = Val(rsTmp!ID & "")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceInput���뵥(ByVal intType As Integer)
'���ܣ�����������Ŀ�󵯳����뵥
    Dim str��� As String
    Dim lngRow As Long
    Dim blnOK As Boolean
    Dim objAppPages() As clsApplicationData
    Dim rsCard As ADODB.Recordset
 
    On Error Resume Next
'    '����
    cbsMain.FindControl(, conMenu_New, True, True).Execute
    lngRow = vsAdvice.Row
    Select Case intType
    Case 1 '�������
        blnOK = ApplyNew�������(0, "", objAppPages())
    Case 2 '��������
        blnOK = ApplyNew��������(0, "", rsCard)
    Case 3 '��Ѫ����
        blnOK = ApplyNew��Ѫ����(0, "", rsCard)
    End Select
    err.Clear

    On Error GoTo errH
    '����ѡ����Ŀ����ȱʡҽ����Ϣ
    If blnOK Then
        Select Case intType
        Case 1 '�������
            Call AdviceSet�������(lngRow, objAppPages())
        Case 2 '��������
            Call AdviceSet��������(lngRow, rsCard)
        Case 3 '��Ѫ����
            Call AdviceSet��Ѫ����(lngRow, rsCard)
        End Select

        '��ʾ��ȱʡ���õ�ֵ
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        Call CalcAdviceMoney '��ʾ�¿�ҽ�����
        'ҽ���ܿ�ʵʱ���
        If mint���� <> 0 And Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_EDIT)) = 0 Then
            '�����������룺ȱʡ���̶�������ҽ�����Լ�����
            '����ҽ������������
            If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) And Not txt����.Enabled Then
                If MakePriceRecord(vsAdvice.Row) Then
                    If Not gclsInsure.CheckItem(mint����, 0, 0, mrsPrice) Then
                        Call AdviceCurRowClear: Exit Sub
                    End If
                End If
                '���Ϊ�Ѿ����˼��
                vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_״̬) = 1
            End If
        End If
        txtҽ������.SetFocus: Call SeekNextControl '�����ȶ�λ
        mblnNoSave = True
    Else
        '�ָ�ԭֵ(AdviceInput�����п��ܴ�����һ��)
        txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
        txtҽ������.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetData�������(ByVal lngRow As Long, ByRef objAppPages() As clsApplicationData)
'���ܣ���ҽ������л�ȡ�������ɶ���ֻ���¿�״̬��ҽ����δ�����ҽ������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strAppend As String, strExtData As String
    Dim lng������� As Long
    Dim lng���ID As Long
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim objTmp As clsApplicationData
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngObjIndex As Long
    Dim str��� As String, str���IDs As String
 
    On Error GoTo errH
    With vsAdvice
        lngEnd = -1
        lngObjIndex = -1
        lng������� = Val(.TextMatrix(lngRow, COL_�������))
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_�������)) = lng������� And i > lngEnd Then
                lng���ID = Val(.RowData(i))
                lngBegin = i
                lngRow = i
                Set objTmp = New clsApplicationData
                objTmp.blnIsModify = True
                objTmp.blnAllowUpdate = True
                objTmp.blnIsAdditionalRec = False
                objTmp.lngProjectId = Val(.TextMatrix(lngRow, COL_������ĿID))
                objTmp.blnIsPriority = Val(.TextMatrix(lngRow, COL_��־)) = 1
                objTmp.strStartExeTime = .TextMatrix(lngRow, COL_��ʼʱ��)
                objTmp.lngExeRoomId = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                objTmp.strExeRoomName = Get��������(objTmp.lngExeRoomId)
                objTmp.lngExeRoomType = Val(.TextMatrix(lngRow, COL_ִ������))
                objTmp.strRequestTime = .TextMatrix(lngRow, COL_����ʱ��)
                objTmp.lngRequestRoomId = Val(.TextMatrix(lngRow, COL_��������ID))
                
                If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
                    str���IDs = GetAdviceDiag(.RowData(lngRow), str���)
                    objTmp.strDiagnoseId = str���IDs
                End If
                
                strTmp = objTmp.Get���뵥��Ϣ(objTmp.lngProjectId, 1)
                If strTmp <> "" Then
                    objTmp.lngApplicationPageId = Val(Split(strTmp, "<Split>")(0))
                    objTmp.strApplicationPageName = Split(strTmp, "<Split>")(1)
                    objTmp.strRequestAffixCfg = objTmp.Get���븽��Ŀ����(objTmp.lngApplicationPageId)
                End If
                
                strExtData = Get��鲿λ����(lngRow)
                If InStr(strExtData, vbTab) > 0 Then
                    objTmp.strPartMethod = Split(strExtData, vbTab)(0)
                    objTmp.lngExeType = Val(Split(strExtData, vbTab)(1))
                End If
                
                strAppend = .TextMatrix(lngRow, COL_����)
                If strAppend <> "" Then
                    varArr = Split(strAppend, "<Split1>")
                    strAppend = ""
                    For j = 0 To UBound(varArr)
                        strTmp = varArr(j)
                        If strTmp <> "" Then
                            strAppend = strAppend & "|" & Split(strTmp, "<Split2>")(0) & ":" & Split(strTmp, "<Split2>")(3)
                        End If
                    Next
                    objTmp.strRequestAffix = Mid(strAppend, 2)
                End If
                
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        lngEnd = j
                    Else
                        Exit For
                    End If
                Next
                lngObjIndex = lngObjIndex + 1
                ReDim Preserve objAppPages(lngObjIndex)
                Set objAppPages(lngObjIndex) = objTmp
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ApplyNew�������(ByVal intType As Integer, ByVal str��ĿIDs As String, ByRef objAppPages() As clsApplicationData) As Boolean
'���ܣ��������뵥
'������intType 0-����,1-�޸�
    Dim objPacspplication As New clsPacsApplication
    Dim lng��Ŀid As Long
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    On Error GoTo errH
    lng��Ŀid = Val(str��ĿIDs)
    '��ʼ��������뵥����
    Call objPacspplication.InitComponents(mlng���˿���id, Me)
    ApplyNew������� = objPacspplication.ShowApplicationForm(mlng����ID, 1, mlng�Һ�ID, 0, intType, objAppPages(), , , lng��Ŀid)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceSet�������(ByVal lngRow As Long, ByRef objAppPages() As clsApplicationData) As Long
'���ܣ��Ѽ�����뵥���ݼ��ص������
    Dim objTmp As clsApplicationData
    Dim i As Long, j As Long, k As Long
    Dim lngCnt As Long, lng������� As Long, lngCurRow As Long
    Dim strSQL As String, strAppend As String
    Dim varTmp As Variant, strTmp As String
    Dim rsInput As ADODB.Recordset
    Dim rsAppend As ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim str���IDs As String
    Dim str�������� As String
    
    On Error GoTo errH
    lng������� = Get�������
    mblnRowChange = False
    For i = 0 To UBound(objAppPages)
        Set objTmp = objAppPages(i)
        strSQL = "Select a.����,a.Id As ������Ŀid, Null As �շ�ϸĿid,a.��� As ���ID From ������ĿĿ¼ A where a.id=[1]"
        Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objTmp.lngProjectId)
        strAppend = ""
        If objTmp.strRequestAffix <> "" Then
            strSQL = "Select C.��Ŀ,C.����,C.Ҫ��ID,C.����,d.������,decode(D.��ʾ��,4,D.��ֵ��,NULL) as ��ֵ��" & _
                " From ��������Ӧ�� A,�����ļ��б� B,�������ݸ��� C,����������Ŀ D" & _
                " Where A.������ĿID=[1] And A.Ӧ�ó���=[2]" & _
                " And A.�����ļ�ID=B.ID And B.����=7 And B.ID=C.�ļ�ID And c.Ҫ��id=d.id(+)" & _
                " Order by C.����"
            Set rsAppend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objTmp.lngProjectId, 1)
            varTmp = Split(objTmp.strRequestAffix, "|")
            For j = 0 To UBound(varTmp)
                If InStr(varTmp(j), ":") > 0 Then
                    rsAppend.Filter = "��Ŀ='" & Split(varTmp(j), ":")(0) & "'"
                    If Not rsAppend.EOF Then
                        strSQL = varTmp(j)
                        str�������� = Replace(strSQL, Mid(strSQL, 1, InStr(strSQL, ":")), "")
                        strAppend = IIF(strAppend = "", "", strAppend & "<Split1>") & rsAppend!��Ŀ & "<Split2>" & Val(rsAppend!���� & "") & _
                            "<Split2>" & rsAppend!Ҫ��ID & "<Split2>" & str��������
                    End If
                End If
            Next
        End If
        If i <> 0 Then
            vsAdvice.AddItem "", lngEnd + 1
            lngRow = lngEnd + 1
        End If
        str���IDs = objTmp.strDiagnoseId
        strTmp = objTmp.strPartMethod & vbTab & objTmp.lngExeType
        Call AdviceSet������Ŀ(rsInput, lngRow, 0, 0, strTmp, "")
        lngBegin = lngRow
        With vsAdvice
            For j = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(j, COL_���ID)) = Val(.RowData(lngRow)) Then
                    lngEnd = j
                Else
                    Exit For
                End If
            Next
            If lngEnd < lngBegin Then lngEnd = lngBegin
            For j = lngBegin To lngEnd
                '����һЩ��Ŀ
                .TextMatrix(j, COL_��־) = IIF(objTmp.blnIsPriority, 1, 0) '�����޲�¼
                .TextMatrix(j, COL_��ʼʱ��) = Format(objTmp.strStartExeTime, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, j, COL_��ʼʱ��) = .TextMatrix(j, COL_��ʼʱ��)
                .TextMatrix(j, COL_ִ�п���ID) = IIF(objTmp.lngExeRoomId <= 0, 0, objTmp.lngExeRoomId)
                .TextMatrix(j, COL_ִ������) = objTmp.lngExeRoomType
                .TextMatrix(j, COL_����ʱ��) = Format(objTmp.strRequestTime, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, j, COL_����ʱ��) = .TextMatrix(j, COL_����ʱ��)
                .TextMatrix(j, COL_��������ID) = objTmp.lngRequestRoomId
                .TextMatrix(j, COL_�������) = lng�������
            Next
            .Cell(flexcpData, lngRow, COL_�������) = str���IDs
            '�����������ϵı�Ǵ���
            Call SetDiagFlag(lngRow, 1)
            '���¸�������:�Ե�ǰ�ɼ���Ϊ׼
            If strAppend <> "" Then
                .TextMatrix(lngRow, COL_����) = strAppend
                .Cell(flexcpData, lngRow, COL_����) = 1 '������Ҫ����д��(�������޸�)
                Call ReplaceAdviceAppend(lngRow) 'ȱʡ�滻����ҽ�������븽��
            End If
            '�����Զ������и�
            Call .AutoSize(col_ҽ������)
            Call SetRow��־ͼ��(lngRow, 0)
        End With
    Next
    vsAdvice.Row = lngRow
    mblnRowChange = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetData��Ѫ����(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset)
'���ܣ���ҽ������л�ȡ�������ɶ���ֻ���¿�״̬��ҽ����δ�����ҽ������
    Dim rsCardBak As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As New ADODB.Recordset
    Dim blnTmp As Boolean
    Dim lngҽ��ID As Long
    Dim strSQL As String
    Dim str��� As String
    Dim str���IDs As String
    Dim strTmp As String
    Dim var1 As Variant
    Dim var2 As Variant
    Dim i As Long, j As Long
    Dim str������Ŀ As String, str��ѪĿ�� As String
    
    On Error GoTo errH
    With vsAdvice
        If TypeName(.Cell(flexcpData, lngRow, COL_�������)) = "Recordset" Then
            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, lngRow, COL_�������))
        Else
            Call InitCardRsBlood(rsCard)
            rsCard.AddNew
        End If
        If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
            lngҽ��ID = Val(.RowData(lngRow))
            
            strSQL = "Select ������Ŀid, ������, ����Ѫ��, ����rh,ѪҺ��Ϣ From ��Ѫ������Ŀ Where ҽ��id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��Ѫ������Ŀ", lngҽ��ID)
            Do While Not rsTmp.EOF
                str������Ŀ = str������Ŀ & ";" & rsTmp!������ĿID & "," & rsTmp!������ & "," & rsTmp!����Ѫ�� & "," & rsTmp!����rh & IIF(rsTmp!ѪҺ��Ϣ & "" <> "", "," & rsTmp!ѪҺ��Ϣ, "")
            rsTmp.MoveNext
            Loop
            If Left(str������Ŀ, 1) = ";" Then str������Ŀ = Mid(str������Ŀ, 2)
            
            strSQL = "Select �Ƿ���� as ����,��Ѫ����,��ѪĿ��, ��Ѫ����, ������Ѫʷ, ������Ѫ��Ӧʷ, ��Ѫ���ɼ�����ʷ, �в����, ��Ѫ������, ��ѪѪ�� as Ѫ��, RHD" & vbNewLine & _
                " From ��Ѫ�����¼" & vbNewLine & _
                " Where ҽ��id = [1]"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            If Not rsTmp.EOF Then
                If Val(rsTmp!���� & "") = 0 Then
                   str���IDs = GetAdviceDiag(lngҽ��ID, str���)
                    '�Ӹ����л�ȡ���������������Ը���Ϊ׼
                    strSQL = "select ���� from ����ҽ������ where ҽ��ID=[1] and ��Ŀ='���뵥���'"
                    Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                    If Not rsTmpOther.EOF Then str��� = rsTmpOther!���� & ""
                End If
                rsCard!�ٴ����IDs = str���IDs
                rsCard!�ٴ�������� = str���
                rsCard!Ѫ�� = Val(rsTmp!Ѫ�� & "")
                rsCard!RHD = Val(rsTmp!RHD & "")
                rsCard!���� = Val(rsTmp!���� & "")
                rsCard!��Ѫ���� = Val(rsTmp!��Ѫ���� & "")
                rsCard!������Ѫʷ = Val(rsTmp!������Ѫʷ & "")
                rsCard!������Ѫ��Ӧʷ = Val(rsTmp!������Ѫ��Ӧʷ & "")
                rsCard!��Ѫ���ɼ�����ʷ = Val(rsTmp!��Ѫ���ɼ�����ʷ & "")
                rsCard!��Ѫ���� = rsTmp!��Ѫ���� & ""
                rsCard!��ѪĿ�� = rsTmp!��ѪĿ�� & ""
                str��ѪĿ�� = rsTmp!��ѪĿ�� & ""
                rsCard!�в���� = Val(rsTmp!�в���� & "")
                rsCard!��Ѫ������ = Val(rsTmp!��Ѫ������ & "")
            End If
            var2 = Array()
            strSQL = "select ���,������ĿID,ָ�����,ָ��������,ָ��Ӣ����,ָ����,�����λ,�����־,����ο�,ȡֵ����,�Ƿ��˹���д from ��Ѫ������ Where ҽ��ID=[1] order by ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            For i = 1 To rsTmp.RecordCount
                var1 = Array()
                For j = 0 To rsTmp.Fields.Count - 1
                    ReDim Preserve var1(j)
                    var1(j) = rsTmp.Fields(j).value & ""
                Next
                strTmp = Join(var1, "<SplitCol>")
                ReDim Preserve var2(UBound(var2) + 1)
                var2(UBound(var2)) = strTmp
                rsTmp.MoveNext
            Next
            rsCard!����� = Join(var2, "<SplitRow>")
        End If
        
        rsCard!������Ŀ = str������Ŀ
        If str��ѪĿ�� = "" Then rsCard!��ѪĿ�� = .TextMatrix(lngRow, COL_��ҩ����)
        rsCard!��Ѫ���� = Val(.TextMatrix(lngRow, COL_��־))
        rsCard!Ԥ����Ѫ���� = .TextMatrix(lngRow, COL_��Ѫʱ��)
        rsCard!��Ѫ��ĿID = Val(.TextMatrix(lngRow, COL_������ĿID))
        rsCard!��Ѫִ�п���ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        rsCard!Ԥ����Ѫ�� = Val(.TextMatrix(lngRow, COL_����))
        rsCard!��Ѫ;����ĿID = Val(.TextMatrix(lngRow + 1, COL_������ĿID))
        rsCard!��Ѫ;��ִ�п���ID = Val(.TextMatrix(lngRow + 1, COL_ִ�п���ID))
        rsCard!��ע = .TextMatrix(lngRow, COL_ҽ������)
        rsCard!��Ѫ�������� = .TextMatrix(lngRow, COL_��ʼʱ��)
        rsCard!�������ID = .TextMatrix(lngRow, COL_��������ID)
        rsCard!���� = .TextMatrix(lngRow + 1, COL_ҽ������)
        rsCard.Update
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ApplyNew��Ѫ����(ByVal intType As Integer, ByVal str��ĿIDs As String, ByRef rsCard As ADODB.Recordset, Optional ByVal bln��Ѫ As Boolean = True) As Boolean
'���ܣ��������뵥
'������intType 0-����,1-�޸�
    Dim lng��Ŀid As Long
    Dim lng��������ID As Long
    
    On Error GoTo errH
    lng��Ŀid = Val(str��ĿIDs)
 
    '��Ѫҽ����飬�����м�������רҵ����ְ���ҽʦ�������´�
    If gbln��Ѫ�����м����� Then
        If UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "������ҽʦ" Then
            MsgBox "��������Ѫ�ּ��������Ѫҽ��ֻ���м�������רҵ����ְ��ҽʦ�����´", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    
    lng��������ID = Get��������ID(UserInfo.ID, mlngҽ������ID, mlng���˿���id, 2)
    
    If Not rsCard Is Nothing Then
        If Not rsCard.EOF Then
            If Val(rsCard!�������ID & "") <> 0 Then
                lng��������ID = Val(rsCard!�������ID & "")
            End If
        End If
    End If
    
    If gblnѪ��ϵͳ = True Then
        ApplyNew��Ѫ���� = frmBloodApplyNew.ShowMe(Me, mlng����ID, 0, 1, intType, 0, mlng���˿���id, _
             , lng��������ID, , , mrsDefine, , 1, mstr�Һŵ�, , lng��Ŀid, rsCard, mintӤ��, 1, IIF(bln��Ѫ = True, 0, 1))
    Else
        ApplyNew��Ѫ���� = frmBloodApply.ShowMe(Me, mlng����ID, 0, 1, intType, 0, mlng���˿���id, _
             , lng��������ID, , , mrsDefine, , 1, mstr�Һŵ�, , lng��Ŀid, rsCard, mintӤ��, 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceSet��Ѫ����(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset) As Long
'���ܣ��Ѽ�����뵥���ݼ��ص������
    Dim objTmp As clsApplicationData
    Dim i As Long, j As Long, k As Long
    Dim lngCnt As Long, lng������� As Long, lngCurRow As Long
    Dim strSQL As String, strAppend As String
    Dim varTmp As Variant, strTmp As String
    Dim rsInput As ADODB.Recordset
    Dim rsAppend As ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim arrItem, strIDs As String
    
    On Error GoTo errH
    lng������� = Get�������
    mblnRowChange = False
    
    strSQL = "Select a.����,a.Id As ������Ŀid, Null As �շ�ϸĿid,a.��� As ���ID From ������ĿĿ¼ A where a.id=[1]"
    Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsCard!��Ѫ��ĿID & ""))
    Call AdviceSet������Ŀ(rsInput, lngRow, Val(rsCard!��Ѫ;����ĿID & ""), 0, "", "")
    
    '��Ѫ�������������Ʒ�֣��˴���������������
    arrItem = Split(rsCard!������Ŀ & "", ";")
    strIDs = ""
    If UBound(arrItem) > 0 Then
        For i = 0 To UBound(arrItem)
            strIDs = strIDs & "," & Val(arrItem(i))
        Next
        strIDs = Mid(strIDs, 2)
        strSQL = "Select /*+ CARDINALITY(C 10) */" & vbNewLine & _
            "  f_List2str(Cast(Collect(a.����) As t_Strlist)) ����" & vbNewLine & _
            " From ������ĿĿ¼ a, Table(f_Num2list([1])) b" & vbNewLine & _
            " Where a.Id = b.Column_Value"
        Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
        If rsInput.EOF = False Then
            vsAdvice.TextMatrix(lngRow, COL_����) = rsInput!����
        End If
    End If
    
    With vsAdvice
        For j = lngRow To lngRow + 1
            '����һЩ��Ŀ
            .TextMatrix(j, COL_��־) = rsCard!��Ѫ����
            .TextMatrix(j, COL_����ʱ��) = rsCard!��Ѫ�������� '", adVarChar, 500
            .TextMatrix(j, COL_��ʼʱ��) = rsCard!��Ѫ��������
            .Cell(flexcpData, j, COL_��ʼʱ��) = .TextMatrix(j, COL_��ʼʱ��)
            .Cell(flexcpData, j, COL_����ʱ��) = .TextMatrix(j, COL_����ʱ��)
            .TextMatrix(j, COL_��������ID) = Val(rsCard!�������ID & "")
            .TextMatrix(j, COL_�������) = lng�������
        Next
        .TextMatrix(lngRow, COL_��Ѫʱ��) = rsCard!Ԥ����Ѫ����
        .TextMatrix(lngRow, COL_����) = rsCard!Ԥ����Ѫ��
        .TextMatrix(lngRow, COL_ҽ������) = rsCard!��ע
        .TextMatrix(lngRow, COL_��ҩ����) = rsCard!��ѪĿ��
        .TextMatrix(lngRow, COL_ִ�п���ID) = Val(rsCard!��Ѫִ�п���ID & "")
        .TextMatrix(lngRow + 1, COL_ִ�п���ID) = Val(rsCard!��Ѫ;��ִ�п���ID & "")
        .TextMatrix(lngRow + 1, COL_ҽ������) = rsCard!���� & ""
        '�����������ϵı�Ǵ���
        Call SetDiagFlag(lngRow, 1)
        Set .Cell(flexcpData, lngRow, COL_�������) = zlDatabase.CopyNewRec(rsCard)
        .TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
        '�����Զ������и�
        Call .AutoSize(col_ҽ������)
    End With
    vsAdvice.Row = lngRow
    mblnRowChange = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetData��������(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset)
'���ܣ���ҽ������л�ȡ�������ɶ���ֻ���¿�״̬��ҽ����δ�����ҽ������
    Dim rsCardBak As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
    Dim strSQL As String
    Dim str��� As String
    Dim str���IDs As String
    Dim strTmp As String
    Dim var1 As Variant
    Dim var2 As Variant
    Dim i As Long, j As Long
    Dim strLIS As String
    Dim strResult As String
    Dim lng������� As Long
    Dim strAppend As String
    
    On Error GoTo errH
    With vsAdvice
    
        If TypeName(.Cell(flexcpData, lngRow, COL_�������)) = "Recordset" Then
            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, lngRow, COL_�������))
        Else
            Call InitCardRsLIS(rsCard)
            rsCard.AddNew
        End If
        If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
            str���IDs = GetAdviceDiag(Val(.RowData(lngRow)), str���)
            rsCard!�ٴ����IDs = str���IDs
        End If
     
        lng������� = Val(.TextMatrix(lngRow, COL_�������))
        For i = .FixedRows To .Rows - 1
            '����ҽ����ʾ���ǲɼ���ʽ
            If Val(.TextMatrix(i, COL_�������)) = lng������� And Not .RowHidden(i) Then
         
                lngRow = i
                '�������id , ִ�п���id, ����ʱ��1, �걾1, ����, ����, �Ƿ�֢, �ɼ�id, ������Ŀid1
                strLIS = .TextMatrix(lngRow, COL_ִ�п���ID) & "<Split A>" & .TextMatrix(lngRow - 1, COL_ִ�п���ID) & "<Split A>" & .TextMatrix(lngRow, COL_��ʼʱ��) & _
                      "<Split A>" & .TextMatrix(lngRow, COL_�걾��λ)
                strAppend = .TextMatrix(lngRow, COL_����)
                If strAppend = "" And Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
                     strAppend = Get����ҽ������(.RowData(lngRow))
                End If
                strLIS = strLIS & "<Split A>" & strAppend & "<Split A>" & .TextMatrix(lngRow, COL_ҽ������) & "<Split A>" & Val(.TextMatrix(lngRow, COL_��־)) & "<Split A>" & .TextMatrix(lngRow, COL_������ĿID) & "<Split A>" & .TextMatrix(lngRow - 1, COL_������ĿID)
                If strLIS <> "" Then
                    strResult = strLIS & IIF(strResult = "", "", "<Split B>" & strResult)
                End If
            End If
        Next
        rsCard!������Ϣ = strResult
        rsCard.Update
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ApplyNew��������(ByVal intType As Integer, ByVal str��Ŀ���� As String, ByRef rsCard As ADODB.Recordset) As Boolean
'���ܣ��������뵥
'������intType 0-����,1-�޸�
    Dim lng��������ID As Long
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    Dim strResult As String
    Dim strDept As String
    Dim lng������� As String
    Dim strDiag As String
    Dim blnCancel As Boolean
    Dim strErr As String
    Dim strLIS As String

    On Error GoTo errH
    
    'ִ�в���(�ű����)�����˿���
    strSQL = "Select A.����,A.�Ա�,A.����,B.�����,B.סԺ��,B.������,a.ID as �Һ�ID," & _
        " B.����,B.��������,C.���� as ִ�в���,A.�Ǽ�ʱ��" & _
        " From ���˹Һż�¼ A,������Ϣ B,���ű� C" & _
        " Where A.NO(+)=[2] And a.��¼����(+)=1 And a.��¼״̬(+)=1 And B.����ID=[1]" & _
        " And A.����ID(+)=B.����ID And A.ִ�в���ID=C.ID(+)"
 
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    
    If rsPati.RecordCount = 0 Then
        MsgBox "δ����ȷ��ȡ������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    strDept = Get��������(mlng���˿���id)
    Call InitObjLis(p����ҽ���´�)
    If gobjLIS Is Nothing Then Exit Function
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    On Error GoTo errH
    If Not rsCard Is Nothing Then
        If Not rsCard.EOF Then
            strLIS = rsCard!������Ϣ & ""
            strDiag = rsCard!�ٴ����IDs & ""
        End If
    End If
    strResult = gobjLIS.ShowLisApplicationForm(mfrmParent, 0, mlng����ID, 0, Val("" & rsPati!�Һ�ID), rsPati!����, "" & rsPati!�Ա�, "" & rsPati!����, 1, _
        Val("" & rsPati!�����), Val("" & rsPati!סԺ��), Val("" & rsPati!������), strDiag, UserInfo.����, UserInfo.����ID, UserInfo.������, mlng���˿���id, strDept, blnCancel, strErr, strLIS, str��Ŀ����)
    
    If strErr <> "" Then
        MsgBox "����ӿ��ڲ�����" & strErr, vbInformation, gstrSysName
        Exit Function
    ElseIf blnCancel Then
        Exit Function
    End If
    
    Call InitCardRsLIS(rsCard)
    rsCard.AddNew Array("�ٴ����IDs", "������Ϣ"), Array(strDiag, strResult)
    ApplyNew�������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceSet��������(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset) As Long
'���ܣ��Ѽ�����뵥���ݼ��ص������
    Dim i As Long, j As Long
    Dim varTmp As Variant, strTmp As String
    Dim varArr As Variant
    Dim lng������� As Long
    Dim strLIS As String
    Dim str��� As String
    Dim str���� As String
    
    On Error GoTo errH
    mblnRowChange = False
    strLIS = rsCard!������Ϣ & ""
    str��� = rsCard!�ٴ����IDs & ""
    If str��� <> "" Then
        str��� = GetDiag�������(str���)
        If str��� <> "" Then
            str��� = "���뵥���<Split2>0<Split2><Split2>" & str���
        End If
    End If
    varTmp = Split(strLIS, "<Split B>")
    For i = 0 To UBound(varTmp)
        If i <> 0 Then
            vsAdvice.AddItem "", lngRow + 1
            lngRow = lngRow + 1
        End If
        lng������� = Get�������
        varArr = Split(varTmp(i), "<Split A>")
        Call AdviceSet�������(lngRow, Val(varArr(7)), varArr(8) & ";" & varArr(3), , , True)
        lngRow = lngRow + 1
        With vsAdvice
            For j = lngRow - 1 To lngRow
                .TextMatrix(j, COL_��־) = Val(varArr(6))
                .TextMatrix(j, COL_����ʱ��) = varArr(2)
                    .Cell(flexcpData, j, COL_����ʱ��) = .TextMatrix(j, COL_����ʱ��)
                .TextMatrix(j, COL_�������) = lng�������
            Next
            .TextMatrix(lngRow, COL_ִ�п���ID) = Val(varArr(0))
            .TextMatrix(lngRow, COL_ҽ������) = varArr(5)
            .TextMatrix(lngRow, COL_������ĿID) = Val(varArr(7))
            .TextMatrix(lngRow - 1, COL_ִ�п���ID) = Val(varArr(1))
            
            str���� = varArr(4)
            If str���� <> "" And str��� <> "" Then
                str���� = str���� & "<Split1>" & str���
            ElseIf str���� = "" And str��� <> "" Then
                str���� = str���
            End If
            If str���� <> "" Then
                .TextMatrix(lngRow, COL_����) = str����
                .Cell(flexcpData, lngRow, COL_����) = 1
            End If
            
            Set .Cell(flexcpData, lngRow, COL_�������) = zlDatabase.CopyNewRec(rsCard)
            Call .AutoSize(col_ҽ������)
        End With
        Call SetDiagFlag(lngRow, 1)
        Call SetRow��־ͼ��(lngRow, 0)
    Next
    vsAdvice.Row = lngRow
    mblnRowChange = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check����Ӧ��(ByVal lngRow As Long, ByRef strMsg As String) As Boolean
    '����Ƿ���ҽ����������Ŀδ��ѡ�����Ե���Ӧ�á���
    With vsAdvice
        If Not (InStr(",4,5,6,7,G,", .TextMatrix(lngRow, COL_���)) > 0) Then
            If Not (.TextMatrix(lngRow, COL_���) = "E" And InStr(",2,3,4,6,7,8,9,", .TextMatrix(lngRow, COL_��������)) > 0) Then
                If Val(.TextMatrix(lngRow, COL_����Ӧ��)) = 0 And Val(.TextMatrix(lngRow, COL_������ĿID)) <> 0 Then
                        strMsg = "����Ӧ��������Ŀ���ܵ���Ӧ�ã�"
                        Check����Ӧ�� = False
                        Exit Function
                End If
            End If
        End If
    End With
    Check����Ӧ�� = True
End Function

Private Sub MakeAppNo(ByVal intType As Integer, ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ�������ϵ�ҽ�����ݲ����������
'������lngBegin - lngEnd ����еķ�Χ��intType ��1 ����ҽ����2������ҽ��
    Dim str��� As String
    Dim lngNo As Long
    Dim lng��Ŀid As Long
    Dim i As Long, j As Long
    Dim lng��ID As Long
    Dim lngPre��ID As Long
    Dim lngStart As Long
    Dim lngStop As Long
    
    On Error GoTo errH
    With vsAdvice
        If intType = 1 Then
            str��� = .TextMatrix(lngBegin, COL_���)
            lng��Ŀid = Val(.TextMatrix(lngBegin, COL_������ĿID))
            Select Case str���
            Case "D", "C"
                If CanUseApply(str���, lng��Ŀid) Then
                    lngNo = -1 * Get�������
                End If
            Case "K"
                If CanUseApply("K", lng��Ŀid) Then
                    lngNo = -1 * Get�������
                End If
            End Select
            If lngNo <> 0 Then
                For i = lngBegin To lngEnd
                    .TextMatrix(i, COL_�������) = lngNo
                Next
            End If
        Else
            For i = lngBegin To lngEnd
                If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    lng��ID = .RowData(i)
                Else
                    lng��ID = Val(.TextMatrix(i, COL_���ID))
                End If
                
                lngStop = -1
                lngStart = i
                For j = i + 1 To lngEnd
                    If lng��ID = IIF(Val(.TextMatrix(j, COL_���ID)) = 0, Val(.RowData(j)), Val(.TextMatrix(j, COL_���ID))) Then
                        lngStop = j
                    Else
                        Exit For
                    End If
                Next
                If lngStop = -1 Then lngStop = lngStart
                
                Call MakeAppNo(1, lngStart, lngStop)
                
                If lngStop = lngEnd Then
                    Exit Sub
                Else
                    i = lngStop
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Make���ҽ����Ӧ(ByRef arrSQL As Variant, ByVal strҽ��IDs As String, ByVal lng���ID As Long)
'���ܣ��������ҽ����Ӧ�Ŀ�ִ��SQL�����102666
    Dim varTmp As Variant
    Dim i As Long
    Dim strTmp As String
    Dim strIDs As String
    varTmp = Array()
    strIDs = strҽ��IDs
    
    Do While Len(strIDs) > 4000
        strTmp = Mid(strIDs, 1, 3980)
        strTmp = Mid(strTmp, 1, InStrRev(strTmp, ",") - 1)
        ReDim Preserve varTmp(UBound(varTmp) + 1)
        varTmp(UBound(varTmp)) = strTmp
        strIDs = Replace(strIDs, strTmp & ",", "")
    Loop
    If strIDs <> "" Then
        ReDim Preserve varTmp(UBound(varTmp) + 1)
        varTmp(UBound(varTmp)) = strIDs
    End If
    For i = 0 To UBound(varTmp)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(NULL,NULL," & lng���ID & ",'" & varTmp(i) & "')"
    Next
End Sub

Private Function GetNextҽ��ID() As Long
'���ܣ�������ʱ��ҽ��ID
    mlngID���� = mlngID���� - 1
    GetNextҽ��ID = mlngID����
End Function

Private Function GetIDָ��ֵ(ByRef colIn As Collection, ByVal strKey As String) As Long
'���ܣ���ȡָ������ʵҽ��ID���ü��Ϸ�ʽ���ɼ�ֵ��
    Dim strID As String
    
    On Error Resume Next
    
    strID = colIn(strKey)
    If err.Number <> 0 Then
        strID = zlDatabase.GetNextID("����ҽ����¼")
        colIn.Add strID, strKey
    End If
    err.Clear
    
    GetIDָ��ֵ = Val(strID)
End Function

Private Sub MakeRealID()
'���ܣ���ҽ������е�ҽ��ID����Ϊ��ʵ��ҽ��ID���������ҽ����Ӧ�е�ҽ��ID
'˵�����ύ����ʱ����ʧ�ܣ����ܻᱻ�ظ�ִ�У������õ�ǰֵ�Ƿ����0�����жϣ�ֻ�б����µ�ҽ������Ҫ���²���
    Dim i As Long, j As Long
    Dim colID As New Collection
    Dim varTmp As Variant
    Dim strҽ��IDs As String
    Screen.MousePointer = 11
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If InStr("1,2", Val(.TextMatrix(i, COL_EDIT))) > 0 Then   '����ҽ����¼
                If Val(.RowData(i)) < 0 Then
                    .RowData(i) = GetIDָ��ֵ(colID, CStr(.RowData(i)))
                End If
                If Val(.TextMatrix(i, COL_���ID)) < 0 Then
                    .TextMatrix(i, COL_���ID) = GetIDָ��ֵ(colID, CStr(Val(.TextMatrix(i, COL_���ID))))
                End If
            End If
        Next
    End With
    
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col���)) <> "" Then
                If "" <> .TextMatrix(i, colҽ��ID) Then
                    varTmp = Split(.TextMatrix(i, colҽ��ID), ",")
                    strҽ��IDs = ""
                    For j = 0 To UBound(varTmp)
                        If Val(varTmp(j)) < 0 Then
                            strҽ��IDs = strҽ��IDs & "," & GetIDָ��ֵ(colID, CStr(Val(varTmp(j))))
                        Else
                            strҽ��IDs = strҽ��IDs & "," & Val(varTmp(j))
                        End If
                    Next
                    .TextMatrix(i, colҽ��ID) = Mid(strҽ��IDs, 2)
                End If
            End If
        Next
    End With
    Screen.MousePointer = 0
End Sub


Private Sub UpdateRecipeNo()
'����:��ȡ�������
'˵��:����������ɹ���
'1-��ҩ���г�ҩ����һ��������ţ���ҩ��ƬӦ����������������š�
'2-��ҩ���г�ҩÿ�Ŵ������ó���5��ҩƷ��һ����ҩ��һ��ҩƷ����
'
    Dim i               As Long
    Dim j               As Long
    Dim rsRecipe        As ADODB.Recordset
    Dim lngҽ��ID       As Long
    Dim lngSumCount     As Long
    Dim lngCount        As Long
    Dim lngNo           As Long
    Dim lngTemp         As Long
    Dim lngRecipeCount  As Long    '�������� =5
    
    lngRecipeCount = 5
    '���컺��ҽ��ID�봦����ŵļ�¼������
    Set rsRecipe = New ADODB.Recordset
    With rsRecipe
        .Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
        .Fields.Append "RecipeNo", adBigInt
        .Fields.Append "Type", adInteger    '1-��ҩ���г�ҩ;2-��ҩ��Ƭ
        .Fields.Append "Tag", adInteger    '1-�Ѿ������������;2-��Ҫ�����������
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_״̬)) = 1 And InStr(",5,6,7,", "," & .TextMatrix(i, COL_���) & ",") > 0 And lngҽ��ID <> Val(.TextMatrix(i, COL_���ID)) Then
                lngҽ��ID = Val(.TextMatrix(i, COL_���ID))
                rsRecipe.AddNew
                rsRecipe!ҽ��ID = Val(.TextMatrix(i, COL_���ID))
                rsRecipe!RecipeNo = Val(.TextMatrix(i, COL_�������))
                rsRecipe!Tag = IIF(Val(.TextMatrix(i, COL_�������)) = 0, 2, 1)
                rsRecipe!Type = IIF(InStr(",5,6,", "," & .TextMatrix(i, COL_���) & ",") > 0, 1, 2)
            End If
        Next
        If rsRecipe.RecordCount > 0 Then rsRecipe.UpdateBatch
        '��ҩ�г�ҩ�������
        rsRecipe.Filter = "Type = 1"
        lngSumCount = rsRecipe.RecordCount
        If lngSumCount > 0 Then
            rsRecipe.Filter = "Type = 1 And RecipeNo = 0"
            lngCount = rsRecipe.RecordCount
        End If
        If lngSumCount = lngCount And lngCount > 0 Then
            For i = 1 To rsRecipe.RecordCount
                If i Mod lngRecipeCount = 1 Then
                    lngNo = GetRecipeNo()
                End If
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        ElseIf lngCount > 0 And lngSumCount - lngCount > 0 Then
            rsRecipe.Filter = "Type = 1 And RecipeNo > 0"
            rsRecipe.Sort = "RecipeNo"
            '������ȡ������ż���Ҫ����ӵ��ô�����ŵ�ҽ������
            lngTemp = 0: lngNo = 0
            For i = 1 To rsRecipe.RecordCount
                If lngNo <> rsRecipe!RecipeNo Then
                    If lngTemp > 0 Then Exit For
                    lngNo = rsRecipe!RecipeNo
                    lngTemp = lngRecipeCount
                End If
                If lngNo = rsRecipe!RecipeNo Then lngTemp = lngTemp - 1
                rsRecipe.MoveNext
            Next
            rsRecipe.Filter = "Type = 1 And RecipeNo = 0"
            rsRecipe.Sort = ""
            For i = 1 To rsRecipe.RecordCount
                If i <= lngTemp Then
                    rsRecipe!RecipeNo = lngNo
                Else
                    Exit For
                End If
                rsRecipe.MoveNext
            Next
            '��Ҫ�²����������
            rsRecipe.Filter = "Type = 1 And RecipeNo = 0"
            For i = 1 To rsRecipe.RecordCount
                If i Mod lngRecipeCount = 1 Then lngNo = GetRecipeNo()
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        End If
        
        '��ҩ��Ƭ���ɴ������
        rsRecipe.Filter = "Type = 2"
        lngSumCount = rsRecipe.RecordCount
        If lngSumCount > 0 Then
            rsRecipe.Filter = "Type = 2 And RecipeNo = 0"
            lngCount = rsRecipe.RecordCount
        End If
        If lngSumCount = lngCount And lngCount > 0 Then
            lngNo = GetRecipeNo()
            For i = 1 To rsRecipe.RecordCount
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        ElseIf lngCount > 0 And lngSumCount - lngCount > 0 Then
            rsRecipe.Filter = "Type = 2 And RecipeNo > 0"
            If Not rsRecipe.EOF Then lngNo = rsRecipe!RecipeNo
            rsRecipe.Filter = "Type = 2 And RecipeNo = 0"
            For i = 1 To rsRecipe.RecordCount
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        End If
        rsRecipe.Filter = "Tag=2"
        rsRecipe.Sort = ""
        '���������׷�ӵ����
        For i = 1 To rsRecipe.RecordCount
            lngҽ��ID = rsRecipe!ҽ��ID
            lngCount = .FindRow(lngҽ��ID, , COL_���ID)
            If lngCount <> -1 Then
                For j = lngCount To .Rows - 1
                    If Val(.TextMatrix(j, COL_���ID)) = lngҽ��ID Or CLng(.RowData(j)) = lngҽ��ID Then
                        .TextMatrix(j, COL_�������) = rsRecipe!RecipeNo & ""
                    Else
                        Exit For
                    End If
                Next
            End If
            rsRecipe.MoveNext
        Next
    End With
End Sub

Private Function GetRecipeNo() As Long
'����:��ȡ�������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    GetRecipeNo = Val(rsTmp!�������)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReadMsg()
'���ܣ���Ϣ�Ķ� Ŀǰ��ʱ����ZLHIS_BLOOD_004��Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
        
    strSQL = "select 1 from ����ҽ����¼ a where a.�Һŵ�=[1] and a.ҽ��״̬=1 and a.�������='K' and a.��鷽��='1' and a.���״̬=1 and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    
    If rsTmp.EOF Then 'û�����������ˣ�����Ϣ��Ϊ����
        strSQL = "select 1 from ҵ����Ϣ�嵥 a where a.����id=[1] and a.����id=[2] and a.���ͱ���='ZLHIS_BLOOD_004' and nvl(a.�Ƿ�����,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
        If Not rsTmp.EOF Then
            strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & mlng����ID & "," & mlng�Һ�ID & ",'ZLHIS_BLOOD_004',1,'" & UserInfo.���� & "'," & mlng���˿���id & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
