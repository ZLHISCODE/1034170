VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocalSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab sTab 
      Height          =   5925
      Left            =   90
      TabIndex        =   63
      Top             =   60
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10451
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "�������(&1)"
      TabPicture(0)   =   "frmLocalSet.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraInput"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkAutoRefresh"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "ҽ�ƿ�Ʊ�ݿ���(&2)"
      TabPicture(1)   =   "frmLocalSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDefaultPayCard"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "img16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkMustCard"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdDeviceSetup(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkCardFeeCharge"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraTitle"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkBruhCardBackCard"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkBrushCardVerfy"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cboType"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkScanIDPatiVisa"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Ԥ��������(&3)"
      TabPicture(2)   =   "frmLocalSet.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblEdit"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "vs����"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboDefaultBalance"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fra�˿�����"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkNotClearPatiInfor"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "chkNotInDeptNotJk"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "chkAdvance"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkSeekName"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkVeryfyInDeposit"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Ԥ��Ʊ�ݿ���(&4)"
      TabPicture(3)   =   "frmLocalSet.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fra��Ʊ��ʽ"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fraPrepay"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "chkAllowDept"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "chkHave"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdDeviceSetup(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "chkLedWelcome"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "chkCheckBillNum"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtƱ������"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "updƱ������"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "fraƱ�ݸ�ʽ"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).ControlCount=   10
      Begin VB.Frame fraƱ�ݸ�ʽ 
         Caption         =   "Ԥ��Ʊ�ݸ�ʽ"
         Height          =   1305
         Left            =   90
         TabIndex        =   53
         Top             =   2205
         Width           =   6615
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1005
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   6435
            _cx             =   11351
            _cy             =   1773
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
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmLocalSet.frx":0070
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
            ExplorerBar     =   2
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
      Begin VB.CheckBox chkVeryfyInDeposit 
         Caption         =   "��סԺԤ��ˢ����֤"
         Height          =   300
         Left            =   -71235
         TabIndex        =   65
         Top             =   5445
         Width           =   2340
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "�Ƿ�����ͨ������ģ�����Ҳ���"
         Height          =   180
         Left            =   -74760
         TabIndex        =   64
         ToolTipText     =   "��Ԥ��ʱ���������Ƿ�ģ�����Ҳ���"
         Top             =   5505
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.CheckBox chkAdvance 
         Caption         =   "�����Ժ���˽�סԺԤ��"
         Height          =   300
         Left            =   -71235
         TabIndex        =   50
         Top             =   5145
         Width           =   2340
      End
      Begin VB.CheckBox chkNotInDeptNotJk 
         Caption         =   "��Ժ����δ��Ʋ�׼��Ԥ��"
         Height          =   300
         Left            =   -74745
         TabIndex        =   49
         Top             =   5145
         Width           =   2475
      End
      Begin VB.CheckBox chkAutoRefresh 
         Caption         =   "�л���������ѡ�ʱ���Զ�ˢ�²�������"
         Height          =   180
         Left            =   -73590
         TabIndex        =   31
         Top             =   4170
         Width           =   3840
      End
      Begin MSComCtl2.UpDown updƱ������ 
         Height          =   300
         Left            =   1755
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   4935
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtƱ������"
         BuddyDispid     =   196615
         OrigLeft        =   1500
         OrigTop         =   3285
         OrigRight       =   1755
         OrigBottom      =   3570
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtƱ������ 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1275
         TabIndex        =   57
         Text            =   "10"
         Top             =   4935
         Width           =   480
      End
      Begin VB.CheckBox chkCheckBillNum 
         Caption         =   "Ʊ��ʣ��         ��ʱ��ʼ�����շ�Ա"
         Height          =   285
         Left            =   255
         TabIndex        =   56
         Top             =   4950
         Width           =   3450
      End
      Begin VB.CheckBox chkScanIDPatiVisa 
         Caption         =   "ɨ�����֤ǩԼ"
         Height          =   180
         Left            =   -74640
         TabIndex        =   38
         Top             =   5025
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkNotClearPatiInfor 
         Caption         =   "��Ԥ�������������Ϣ"
         Height          =   300
         Left            =   -71235
         TabIndex        =   48
         Top             =   4785
         Width           =   2340
      End
      Begin VB.Frame fra�˿����� 
         Caption         =   "�˿�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74865
         TabIndex        =   43
         Top             =   3930
         Width           =   6510
         Begin VB.OptionButton optCheck 
            Caption         =   "����ʱ��ֹ�˿�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3600
            TabIndex        =   45
            Top             =   315
            Value           =   -1  'True
            Width           =   2220
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "����ʱ�����Ƿ��˿�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   495
            TabIndex        =   44
            Top             =   315
            Width           =   2625
         End
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73455
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   5355
         Width           =   2580
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Left            =   4785
         TabIndex        =   62
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   5580
         Value           =   1  'Checked
         Width           =   1710
      End
      Begin VB.CheckBox chkBrushCardVerfy 
         Caption         =   "�˿���ȡ���ݺź�ˢ����֤�˿�"
         Height          =   180
         Left            =   -71235
         TabIndex        =   35
         Top             =   4455
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.CheckBox chkBruhCardBackCard 
         Caption         =   "���������ˡ�ˢ���˿�"
         Height          =   240
         Left            =   -71250
         TabIndex        =   37
         Top             =   4725
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Index           =   1
         Left            =   4980
         TabIndex        =   61
         Top             =   4995
         Width           =   1500
      End
      Begin VB.CheckBox chkHave 
         Caption         =   "ֻ��ʾ��ʣ�����ʷ�ɿ�"
         Height          =   195
         Left            =   255
         TabIndex        =   59
         Top             =   5310
         Width           =   3120
      End
      Begin VB.CheckBox chkAllowDept 
         Caption         =   "������Ĳ��˵Ľɿ����"
         Height          =   195
         Left            =   255
         TabIndex        =   60
         Top             =   5595
         Value           =   1  'Checked
         Width           =   2280
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "���ع���Ʊ��"
         Height          =   1740
         Left            =   90
         TabIndex        =   51
         Top             =   435
         Width           =   6615
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   1455
            Left            =   60
            TabIndex        =   52
            Top             =   225
            Width           =   6480
            _cx             =   11430
            _cy             =   2566
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
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmLocalSet.frx":00FE
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
            ExplorerBar     =   2
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
      Begin VB.ComboBox cboDefaultBalance 
         Height          =   300
         Left            =   -73635
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   4770
         Width           =   1875
      End
      Begin VB.Frame fraInput 
         Caption         =   "�����꾭����Ŀ"
         Height          =   3270
         Left            =   -74265
         TabIndex        =   3
         Top             =   750
         Width           =   5190
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�����֤��"
            Height          =   195
            Index           =   26
            Left            =   3450
            TabIndex        =   24
            Top             =   930
            Value           =   1  'Checked
            Width           =   1560
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   25
            Left            =   3450
            TabIndex        =   30
            Top             =   2820
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���ڵ�ַ�ʱ�"
            Height          =   195
            Index           =   24
            Left            =   1785
            TabIndex        =   20
            Top             =   2475
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���ڵ�ַ"
            Height          =   195
            Index           =   23
            Left            =   1785
            TabIndex        =   19
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   22
            Left            =   1785
            TabIndex        =   21
            Top             =   2820
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����֤��"
            Height          =   195
            Index           =   21
            Left            =   285
            TabIndex        =   11
            Top             =   2475
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ������"
            Height          =   195
            Index           =   20
            Left            =   3450
            TabIndex        =   28
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ�ʱ�"
            Height          =   195
            Index           =   19
            Left            =   3450
            TabIndex        =   27
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ�绰"
            Height          =   195
            Index           =   18
            Left            =   3450
            TabIndex        =   26
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "������λ"
            Height          =   195
            Index           =   17
            Left            =   3450
            TabIndex        =   25
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�˵绰"
            Height          =   195
            Index           =   16
            Left            =   3450
            TabIndex        =   23
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�˵�ַ"
            Height          =   195
            Index           =   15
            Left            =   3450
            TabIndex        =   22
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�˹�ϵ"
            Height          =   195
            Index           =   14
            Left            =   1785
            TabIndex        =   18
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ������"
            Height          =   195
            Index           =   13
            Left            =   1785
            TabIndex        =   17
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ͥ�绰"
            Height          =   195
            Index           =   12
            Left            =   1785
            TabIndex        =   16
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ͥ��ַ�ʱ�"
            Height          =   195
            Index           =   11
            Left            =   1785
            TabIndex        =   15
            Top             =   930
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��סַ"
            Height          =   195
            Index           =   10
            Left            =   1785
            TabIndex        =   14
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "�����ص�"
            Height          =   195
            Index           =   9
            Left            =   1785
            TabIndex        =   13
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���֤��"
            Height          =   195
            Index           =   8
            Left            =   285
            TabIndex        =   12
            Top             =   2820
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��������"
            Height          =   195
            Index           =   7
            Left            =   285
            TabIndex        =   10
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���"
            Height          =   195
            Index           =   6
            Left            =   285
            TabIndex        =   9
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "ְҵ"
            Height          =   195
            Index           =   5
            Left            =   285
            TabIndex        =   8
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����״��"
            Height          =   195
            Index           =   4
            Left            =   285
            TabIndex        =   7
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "ѧ��"
            Height          =   195
            Index           =   3
            Left            =   285
            TabIndex        =   6
            Top             =   930
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   5
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   4
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ�ʺ�"
            Height          =   195
            Index           =   0
            Left            =   3450
            TabIndex        =   29
            Top             =   2475
            Value           =   1  'Checked
            Width           =   1200
         End
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع���..."
         Height          =   3825
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   6435
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   3405
            Left            =   60
            TabIndex        =   33
            Top             =   300
            Width           =   6300
            _cx             =   11112
            _cy             =   6006
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
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmLocalSet.frx":01DB
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
            ExplorerBar     =   2
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
      Begin VB.CheckBox chkCardFeeCharge 
         Caption         =   "���￨�����Լ��˷�ʽ��ȡ"
         Height          =   180
         Left            =   -74640
         TabIndex        =   36
         Top             =   4755
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Index           =   0
         Left            =   -70290
         TabIndex        =   41
         Top             =   5325
         Width           =   1500
      End
      Begin VB.CheckBox chkMustCard 
         Caption         =   "����ͬʱ���뷢��"
         Height          =   255
         Left            =   -74640
         TabIndex        =   34
         Top             =   4425
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   -71145
         Top             =   1155
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLocalSet.frx":02BC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vs���� 
         Height          =   3315
         Left            =   -74865
         TabIndex        =   42
         Top             =   465
         Width           =   6525
         _cx             =   11509
         _cy             =   5847
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLocalSet.frx":039E
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
         ExplorerBar     =   3
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
      Begin VB.Frame fra��Ʊ��ʽ 
         Caption         =   "Ԥ����Ʊ��ʽ"
         Height          =   1305
         Left            =   90
         TabIndex        =   66
         Top             =   3570
         Width           =   6615
         Begin VSFlex8Ctl.VSFlexGrid vsRedBillFormat 
            Height          =   1005
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   6435
            _cx             =   11351
            _cy             =   1773
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
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmLocalSet.frx":03FF
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
            ExplorerBar     =   2
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
      Begin VB.Label lblDefaultPayCard 
         Caption         =   "ȱʡ��������"
         Height          =   210
         Left            =   -74625
         TabIndex        =   39
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡ���㷽ʽ"
         Height          =   180
         Left            =   -74775
         TabIndex        =   46
         Top             =   4830
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7065
      TabIndex        =   2
      Top             =   4410
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7050
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7050
      TabIndex        =   0
      Top             =   360
      Width           =   1100
   End
End
Attribute VB_Name = "frmLocalSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlngModul As Long, mstrPrivs As String, mbln���� As Boolean
Private mstrClass As String, mstrDeposit As String
Private mblnOK As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal strPrivs As String, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:mlngModul-1101-������Ϣ����,1102-���￨����,1103-Ԥ�������
    '����:
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 14:22:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mstrPrivs = strPrivs: mlngModul = lngModule
    mbln���� = InStr(mstrPrivs, ";������Ϣ;") > 0 And mlngModul = 1101
    Me.Show 1, frmMain
    zlSetPara = True
End Function
Private Sub cboType_Click()
    chkScanIDPatiVisa.Enabled = Not (cboType.Text = "�������֤")
    If cboType.Text = "�������֤" Then
        chkScanIDPatiVisa.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDeviceSetup_Click(Index As Integer)
    Call zlCommFun.DeviceSetup(Me, 100, mlngModul)
End Sub

Private Sub cmdHelp_Click()
    Select Case mlngModul
        Case 1101 '������Ϣ
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet1"
        Case 1102 '���￨
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet2"
        Case 1103 'Ԥ����
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet3"
    End Select
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
    IsValied = False
    
    On Error GoTo errHandle
    If mlngModul <> 1103 Then
        '���ÿ��ʹ����ʽֻ��һ��ѡ��
        With vsBill
            str��� = "-"
            For i = 1 To vsBill.Rows - 1
                If str��� <> Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����"))) Then
                   str��� = Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����")))
                   lngSelCount = 0
                    For j = 1 To vsBill.Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����"))) = Trim(.TextMatrix(j, .ColIndex("ҽ�ƿ����"))) Then
                            If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                                lngSelCount = lngSelCount + 1
                            End If
                        End If
                    Next
                    If lngSelCount > 1 Then
                        MsgBox "ע��:" & vbCrLf & "    ҽ�ƿ����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                        Exit Function
                    End If
                End If
            Next
        End With
    End If
    If mlngModul = 1102 Then IsValied = True: Exit Function
  '���ÿ��ʹ��Ԥ��ֻ��һ��ѡ��
    With vsPrepay
        str��� = "-"
        For i = 1 To .Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("Ԥ������")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) = Trim(.TextMatrix(j, .ColIndex("Ԥ������"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    Ԥ������Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:���˺�
    '����:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    If mlngModul <> 1103 Then
        '���湲��Ʊ��
        strValue = ""
        With vsBill
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                    strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����")))
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "����ҽ�ƿ�����", strValue, glngSys, mlngModul, blnHavePrivs
    End If
    If mlngModul = 1102 Then Exit Sub
    
    
    '����Ԥ��Ʊ��
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Val(.Cell(flexcpData, i, .ColIndex("Ԥ������")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "����Ԥ��Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
    
    '61808:������,2013-05-21,ֻ��Ԥ���������������Ч
    '78751:���ϴ�,2015/08/24,����Ԥ��Ʊ�ݴ�ӡ��ʽ
    If mlngModul = 1103 Or mlngModul = 1101 Then
        '�����:50656
        Dim strPrintMode As String
        '�����շѸ�ʽ
        strValue = "": strPrintMode = ""
        With vsBillFormat
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("Ԥ����ӡ��ʽ")), 1))
                End If
            Next
            If strValue <> "" Then strValue = Mid(strValue, 2)
            If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
            zlDatabase.SetPara "Ԥ����Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "Ԥ����Ʊ��ӡ��ʽ", strPrintMode, glngSys, mlngModul, blnHavePrivs
        End With
        '��Ʊ��ʽ
        strValue = "": strPrintMode = ""
        With vsRedBillFormat
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("Ԥ����ӡ��ʽ")), 1))
                End If
            Next
            If strValue <> "" Then strValue = Mid(strValue, 2)
            If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
            zlDatabase.SetPara "�˿Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "Ԥ���˿��ӡ��ʽ", strPrintMode, glngSys, mlngModul, blnHavePrivs
        End With
    End If
End Sub
Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String, rsҽ�ƿ���� As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim strȱʡҽ�ƿ� As String, lngȱʡҽ�ƿ� As Long
    Dim strBillFormat As String
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    On Error GoTo errHandle
    '�ָ��п��
    If mlngModul <> 1103 Then
            lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, , , True, intType))
            
            gstrSQL = "Select ID,����,����, nvl(�Ƿ�̶�,0) as �Ƿ�̶�  from ҽ�ƿ����  Where nvl(�Ƿ�����,0)=1"
            Set rsҽ�ƿ���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            rsҽ�ƿ����.Filter = "����='���￨' and �Ƿ�̶�=1"
            If rsҽ�ƿ����.EOF = False Then
                strȱʡҽ�ƿ� = rsҽ�ƿ����!����: lngȱʡҽ�ƿ� = Val(rsҽ�ƿ����!ID)
            End If
            With rsҽ�ƿ����
                cboType.Clear
                rsҽ�ƿ����.Filter = 0
                If rsҽ�ƿ����.RecordCount <> 0 Then rsҽ�ƿ����.MoveFirst
                Do While Not .EOF
                    cboType.AddItem Nvl(!����)
                    cboType.ItemData(cboType.NewIndex) = Nvl(!ID)
                    If Nvl(!����) = "���￨" Then cboType.ListIndex = cboType.NewIndex
                    If lngCardTypeID = Val(Nvl(!ID)) Then
                        cboType.ListIndex = cboType.NewIndex
                    End If
                    .MoveNext
                Loop
            End With
            
            zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
            strShareInvoice = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModul, , , True, intType)
            '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
            vsBill.Tag = ""
            Select Case intType
            Case 1, 3, 5, 15
                vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
                fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
                If intType = 5 Then vsBill.Tag = ""
            Case Else
                vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
                fraTitle.ForeColor = &H80000008
            End Select
            With vsBill
                .Editable = flexEDKbdMouse
                If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
            End With
            
            '��ʽ:����ID1,ҽ�ƿ����ID1|����IDn,ҽ�ƿ����IDn|...
            varData = Split(strShareInvoice, "|")
    
            '1.���ù���Ʊ��
            Set rsTemp = GetShareInvoiceGroupID(5)
            With vsBill
                .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
                lngRow = 1
                .MergeCells = flexMergeRestrictRows
                .MergeCellsFixed = flexMergeFixedOnly
                .MergeCol(0) = True
                Do While Not rsTemp.EOF
                    .RowData(lngRow) = Val(Nvl(rsTemp!ID))
                    '105985:���ϴ�,2017/4/10,��ҽ�ƿ���������Ʊ��
                    If Val(Nvl(rsTemp!ʹ�����ID)) = 0 Then
                        .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = strȱʡҽ�ƿ�
                        .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = lngȱʡҽ�ƿ�
                    Else
                        rsҽ�ƿ����.Filter = "ID=" & Val(Nvl(rsTemp!ʹ�����ID))
                        If Not rsҽ�ƿ����.EOF Then
                            .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsҽ�ƿ����!����)
                        Else
                            .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsTemp!ʹ�����)
                        End If
                        .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = Val(Nvl(rsTemp!ʹ�����ID))
                    End If
                    .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
                    .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
                    .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
                    .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
                    For i = 0 To UBound(varData)
                        varTemp = Split(varData(i) & ",", ",")
                        lngTemp = Val(varTemp(0))
                        If Val(.RowData(lngRow)) = lngTemp _
                            And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����"))) Then
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                        End If
                    Next
                    .MergeRow(lngRow) = True
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
            End With
    End If
    If mlngModul = 1102 Then Exit Sub
    '����Ԥ��Ʊ������
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
    
    strShareInvoice = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, mlngModul, , , True, intType)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,Ԥ�����ID1|����IDn,Ԥ�����IDn|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            '58071
            Select Case Val(Nvl(rsTemp!ʹ�����, ""))
            Case 0 '�����������סԺƱ��
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = 0
            Case 1  '����Ʊ��
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ������Ʊ��"
                .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = 1
            Case Else   'סԺƱ��
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ��סԺƱ��"
                .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = 2
            End Select
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ԥ������"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    '78751:���ϴ�,2015/08/24,����Ԥ��Ʊ�ݴ�ӡ��ʽ
    If mlngModul = 1103 Or mlngModul = 1101 Then
        'Ʊ�ݸ�ʽ����
        Dim strReport As String
        
        zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "Ԥ����Ʊ��ӡ��ʽ", False, False
        strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
        Set rsTemp = zlReadBillFormat(strReport)
        With vsBillFormat
            .Clear 1
            .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
            .ColComboList(.ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
        End With
        
        '��ȡ����ֵ
        strBillFormat = zlDatabase.GetPara("Ԥ����Ʊ��ʽ", glngSys, mlngModul, , , True, intType)
        strPrintMode = zlDatabase.GetPara("Ԥ����Ʊ��ӡ��ʽ", glngSys, mlngModul, , , True, intType1)
        '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
        With vsBillFormat
            .TextMatrix(1, 0) = "����Ԥ��"
            .Cell(flexcpData, 1, 0) = 1
            .TextMatrix(2, 0) = "סԺԤ��"
            .Cell(flexcpData, 2, 0) = 2
            .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
            .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = "0"
            .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
            Select Case intType
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
            End Select
            Select Case intType1
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
            End Select
            
            If (Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Or _
                Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1) Then
                .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
            Else
                .Editable = flexEDKbdMouse
            End If
        End With
        
        vsBillFormat.Tag = ""
        varData = Split(strBillFormat, "|")
        VarType = Split(strPrintMode, "|")
        
        With vsBillFormat
            .Clear 1
            .Rows = 3
            For lngRow = 1 To .Cols - 1
                .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��"
                .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
                For i = 0 To UBound(varData)
                    varTemp = Split(varData(i) & "," & ",", ",")
                    If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                    End If
                Next
                For i = 0 To UBound(VarType)
                    varTemp1 = Split(VarType(i) & "," & ",", ",")
                    If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                        Exit For
                    End If
                Next
            Next
            If Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("Ԥ����ӡ��ʽ"), .Rows - 1, .ColIndex("Ԥ����ӡ��ʽ")) = vbBlue
            End If
            
            If Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("Ʊ�ݸ�ʽ")) = vbBlue
            End If
        End With
        '��Ʊ
        zl_vsGrid_Para_Restore mlngModul, vsRedBillFormat, Me.Name, "Ԥ���˿��ӡ��ʽ", False, False
        strReport = "ZL" & glngSys \ 100 & "_BILL_1103_1"
        Set rsTemp = zlReadBillFormat(strReport)
        With vsRedBillFormat
            .Clear 1
            .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
            .ColComboList(.ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
        End With
        
        '��ȡ����ֵ
        strBillFormat = zlDatabase.GetPara("�˿Ʊ��ʽ", glngSys, mlngModul, , , True, intType)
        strPrintMode = zlDatabase.GetPara("Ԥ���˿��ӡ��ʽ", glngSys, mlngModul, , , True, intType1)
        '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
        With vsRedBillFormat
            .TextMatrix(1, 0) = "����Ԥ��"
            .Cell(flexcpData, 1, 0) = 1
            .TextMatrix(2, 0) = "סԺԤ��"
            .Cell(flexcpData, 2, 0) = 2
            .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
            .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = "0"
            .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
            Select Case intType
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
            End Select
            Select Case intType1
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
            End Select
            
            If (Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Or _
                Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1) Then
                .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
            Else
                .Editable = flexEDKbdMouse
            End If
        End With
        
        vsRedBillFormat.Tag = ""
        varData = Split(strBillFormat, "|")
        VarType = Split(strPrintMode, "|")
        
        With vsRedBillFormat
            .Clear 1
            .Rows = 3
            For lngRow = 1 To .Cols - 1
                .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��"
                .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
                For i = 0 To UBound(varData)
                    varTemp = Split(varData(i) & "," & ",", ",")
                    If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                    End If
                Next
                For i = 0 To UBound(VarType)
                    varTemp1 = Split(VarType(i) & "," & ",", ",")
                    If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                        Exit For
                    End If
                Next
            Next
            If Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("Ԥ����ӡ��ʽ"), .Rows - 1, .ColIndex("Ԥ����ӡ��ʽ")) = vbBlue
            End If
            
            If Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("Ʊ�ݸ�ʽ")) = vbBlue
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, strTmp As String
    
    '���ع��þ��￨
    If IsValied = False Then Exit Sub
    Call SaveInvoice
    
    zlDatabase.SetPara "���Ѽ���", chkCardFeeCharge.Value, glngSys, glngModul, IIf(chkCardFeeCharge.Enabled = True, True, False)
    Select Case mlngModul
    Case 1101 '������Ϣ
        zlDatabase.SetPara "����ͬʱ���뷢��", chkMustCard.Value, glngSys, mlngModul, IIf(chkMustCard.Enabled = True, True, False)
        '����27390  ��꾭����Ŀ
        For i = 0 To chkItem.UBound
            zlDatabase.SetPara chkItem(i).Caption, chkItem(i).Value, glngSys, mlngModul, IIf(chkItem(i).Enabled = True, True, False)
        Next
        '76824�����ϴ���2014/8/19��ҽ�ƿ������
        If cboType.ListIndex >= 0 Then
            zlDatabase.SetPara "ȱʡҽ�ƿ����", cboType.ItemData(cboType.ListIndex), glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        Else
            zlDatabase.SetPara "ȱʡҽ�ƿ����", 0, glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        End If
        '54701:������,2012-09-19
        zlDatabase.SetPara "�Զ�ˢ������", chkAutoRefresh.Value, glngSys, mlngModul, IIf(chkAutoRefresh.Enabled = True, True, False)
    Case 1102   '���￨
        '����28130��27929
        If chkBruhCardBackCard.Value And chkBrushCardVerfy.Value Then
            strTmp = "3"
        ElseIf chkBruhCardBackCard.Value Then
            strTmp = "1"
        ElseIf chkBrushCardVerfy.Value Then
            strTmp = "2"
        Else
            strTmp = "0"
        End If
        Call zlDatabase.SetPara("�˿�ˢ��", strTmp, glngSys, mlngModul, IIf(chkBruhCardBackCard.Enabled = True, True, False))
    Case 1103
        zlDatabase.SetPara "������Ľɿ����", chkAllowDept.Value, glngSys, glngModul, IIf(chkAllowDept.Enabled = True, True, False)
        zlDatabase.SetPara "���������Ľɿ", chkHave.Value, glngSys, glngModul, IIf(chkHave.Enabled = True, True, False)
        zlDatabase.SetPara "�˿��ֹ��ʽ", IIf(optCheck(1).Value, 1, 0), glngSys, glngModul, IIf(optCheck(0).Enabled = True, True, False)
        '�����:51628 �޸���:���˺�,�޸�ʱ��:2012-12-11 11:56:43
        zlDatabase.SetPara "����δ��Ʋ�׼��Ԥ��", IIf(chkNotInDeptNotJk.Value, 1, 0), glngSys, glngModul, InStr(1, mstrPrivs, ";��������;") > 0
        
        zlDatabase.SetPara "�����Ժ���˽�Ԥ��", IIf(chkAdvance, 1, 0), glngSys, glngModul, InStr(1, mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "����ģ������", chkSeekName.Value, glngSys, glngModul, InStr(1, mstrPrivs, ";��������;") > 0
        '63113:������,2013-10-29,���Ӳ���,סԺ��Ԥ��������֤
        zlDatabase.SetPara "סԺ��Ԥ����֤", IIf(chkVeryfyInDeposit, "1", "0"), glngSys, glngModul, InStr(1, mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", IIf(chkCheckBillNum.Value = 1, "1", "0") & "|" & Val(txtƱ������.Text), glngSys, mlngModul, IIf(chkCheckBillNum.Enabled = True, True, False)
    End Select
    If mlngModul = 1101 Then '������Ϣ
    ElseIf mlngModul = 1102 Then '
    ElseIf mlngModul = 1103 Then 'Ԥ����
        
        With vs����
            strTmp = ""
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("�̶����"))) <> 0 And Trim(.TextMatrix(i, .ColIndex("���տ���"))) <> "" Then
                    strTmp = strTmp & "|" & Trim(.TextMatrix(i, .ColIndex("���տ���"))) & ":" & Val(.TextMatrix(i, .ColIndex("�̶����")))
                End If
            Next
            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        End With
        
        zlDatabase.SetPara "ȱʡԤ�����㷽ʽ", Trim(cboDefaultBalance.Text), glngSys, glngModul, IIf(cboDefaultBalance.Enabled = True, True, False)
        zlDatabase.SetPara "���տ�����", strTmp, glngSys, glngModul, IIf(vs����.Editable = flexEDKbdMouse, True, False)
        zlDatabase.SetPara "��Ԥ���������Ϣ", IIf(chkNotClearPatiInfor.Value = 1, 1, 0), glngSys, glngModul, IIf(chkNotClearPatiInfor.Enabled, True, False)
    End If
    'LED�豸
    zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chkLedWelcome.Value, glngSys, mlngModul, IIf(chkLedWelcome.Enabled = True, True, False)
    '�����:53408
    'ɨ�����֤ǩԼ
    zlDatabase.SetPara "ɨ�����֤ǩԼ", IIf(chkScanIDPatiVisa.Value = 1, 1, 0), glngSys, glngModul, InStr(1, mstrPrivs, ";��������;")
    
    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

 Private Sub Load���տ�()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ش��տ�
    '����:���˺�
    '����:2011-07-19 15:13:59
    '����:  34705
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, varData As Variant, varTemp As Variant, j As Long, strTmp As String
    
     str���㷽ʽ = zlDatabase.GetPara("ȱʡԤ�����㷽ʽ", glngSys, glngModul, , Array(cboDefaultBalance), InStr(mstrPrivs, ";��������;") > 0)
     
     On Error GoTo errHandle
    '���㷽ʽ
    strSQL = _
    " Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
    " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
    " Where A.Ӧ�ó���='Ԥ����' And B.����=A.���㷽ʽ And Nvl(B.����,1) In(1,2,3,5,8)" & _
    " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboDefaultBalance
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp!����)
            If .ListIndex < 0 And Val(Nvl(rsTmp!ȱʡ)) = 1 Then .ListIndex = .NewIndex
            If str���㷽ʽ = Nvl(rsTmp!����) Then .ListIndex = .NewIndex
            rsTmp.MoveNext
        Loop
    End With
    '���㷽ʽ:���|���㷽ʽ:���....
    strTmp = zlDatabase.GetPara("���տ�����", glngSys, glngModul, , Array(vs����), InStr(mstrPrivs, ";��������;") > 0)
    varData = Split(strTmp, "|")
    vs����.Tag = "1"
    If vs����.Enabled = False Then vs����.Tag = "0"
    If vs����.Tag = "1" Then vs����.Editable = flexEDKbdMouse
    vs����.Enabled = True
    rsTmp.Filter = "����=5" '������տ�
    With vs����
        If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
        i = 1
        .Rows = IIf(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        Do While rsTmp.EOF = False
            .TextMatrix(i, .ColIndex("���տ���")) = Nvl(rsTmp!����)
            For j = 0 To UBound(varData)
                varTemp = Split(varData(j) & ":", ":")
                If Nvl(rsTmp!����) = varTemp(0) Then
                    .TextMatrix(i, .ColIndex("�̶����")) = Format(Val(varTemp(1)), "###0.00;-###0.00;;")
                    Exit For
                End If
            Next
            i = i + 1
            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
 End Sub

Private Sub Form_Load()
    Dim i As Long, lngCardTypeID As Long
    Dim strPrintMode As String '�����:50656
    Dim strArr��ӡ��ʽ() As String '�����:50656
    Dim strTmp As String
    gblnOK = False
    Me.sTab.TabVisible(2) = False   '34705
    sTab.TabVisible(0) = mlngModul = 1101
    sTab.TabVisible(2) = mlngModul = 1103    '34705
    sTab.TabVisible(1) = mlngModul <> 1103    '34705
    If mlngModul = 1103 Then Call Load���տ�
    Call InitShareInvoice   '���ع�����Ʊ����Ϣ
    
    'LED�豸
    chkLedWelcome.Value = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, 1, Array(chkLedWelcome), InStr(mstrPrivs, ";��������;") > 0)
    chkCardFeeCharge.Value = IIf(zlDatabase.GetPara("���Ѽ���", glngSys, glngModul, , Array(chkCardFeeCharge), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
    '�����:53408
    chkScanIDPatiVisa.Value = IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, glngModul, , Array(chkScanIDPatiVisa), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
    
    Select Case mlngModul
    Case 1101 ''������Ϣ
        chkMustCard.Value = IIf(zlDatabase.GetPara("����ͬʱ���뷢��", glngSys, glngModul, , Array(chkMustCard), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
        '����27390 ��꾭����Ŀ
        For i = 0 To chkItem.UBound
            chkItem(i).Value = zlDatabase.GetPara(chkItem(i).Caption, glngSys, mlngModul, 1, Array(chkItem(i)), InStr(mstrPrivs, ";��������;") > 0)
        Next
        lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, , Array(cboType), InStr(mstrPrivs, ";��������;") > 0))
        For i = 0 To cboType.ListCount - 1
            If cboType.ItemData(i) = lngCardTypeID Then cboType.ListIndex = i: Exit For
        Next
        
        '54701:������,2012-09-19
        chkAutoRefresh.Value = zlDatabase.GetPara("�Զ�ˢ������", glngSys, mlngModul, 1, Array(chkAutoRefresh), InStr(mstrPrivs, ";��������;") > 0)
    
    Case 1102   '���￨
        '����28130
        Select Case Val(zlDatabase.GetPara("�˿�ˢ��", glngSys, mlngModul, "0", Array(chkBruhCardBackCard, chkBrushCardVerfy), InStr(mstrPrivs, ";��������;") > 0))
        Case 0: chkBruhCardBackCard.Value = 0: chkBrushCardVerfy.Value = 0
        Case 1: chkBruhCardBackCard.Value = 1
        Case 2: chkBrushCardVerfy.Value = 1
        Case 3: chkBruhCardBackCard.Value = 1: chkBrushCardVerfy.Value = 1
        End Select
        chkBruhCardBackCard.Visible = True: chkBrushCardVerfy.Visible = True
    Case 1103  'Ԥ����
        chkHave.Value = IIf(zlDatabase.GetPara("���������Ľɿ", glngSys, glngModul, , Array(chkHave), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
        chkAllowDept.Value = IIf(zlDatabase.GetPara("������Ľɿ����", glngSys, glngModul, , Array(chkAllowDept), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
        chkAdvance.Value = IIf(zlDatabase.GetPara("�����Ժ���˽�Ԥ��", glngSys, glngModul, , Array(chkAdvance), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
        chkSeekName.Value = IIf(zlDatabase.GetPara("����ģ������", glngSys, glngModul, , Array(chkSeekName), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
        '63113:������,2013-10-29,���Ӳ���,סԺ��Ԥ����֤
        If gbln������֤ = False Then chkVeryfyInDeposit.Visible = False
        chkVeryfyInDeposit.Value = Val(zlDatabase.GetPara("סԺ��Ԥ����֤", glngSys, glngModul, "0", Array(chkVeryfyInDeposit), InStr(mstrPrivs, ";��������;") > 0))

        
        If zlDatabase.GetPara("�˿��ֹ��ʽ", glngSys, glngModul, , Array(optCheck(0), optCheck(1), fra�˿�����), InStr(mstrPrivs, ";��������;") > 0) = "1" Then
            optCheck(1).Value = True
        Else
            optCheck(0).Value = True
        End If
        '����:43061
        chkNotClearPatiInfor.Value = IIf(zlDatabase.GetPara("��Ԥ���������Ϣ", glngSys, glngModul, , Array(chkNotClearPatiInfor), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
        '�����:51628 �޸���:���˺�,�޸�ʱ��:2012-12-11 11:56:43
        chkNotInDeptNotJk.Value = IIf(zlDatabase.GetPara("����δ��Ʋ�׼��Ԥ��", glngSys, mlngModul, , Array(chkNotInDeptNotJk), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
        '����:50656
        '78410,Ƚ����,2014-10-8,�����ͼ�Ȩ�����ÿؼ�������ɫ�Ϳɱ༭״̬
        '37372
        strTmp = zlDatabase.GetPara("Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", glngSys, mlngModul, "0|10", Array(txtƱ������, updƱ������, chkCheckBillNum), InStr(mstrPrivs, ";��������;") > 0)
        updƱ������.Value = Val(Split(strTmp & "|", "|")(1))
        txtƱ������.Text = updƱ������.Value
        chkCheckBillNum.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
        txtƱ������.Enabled = chkCheckBillNum.Enabled And chkCheckBillNum.Value = 1
        updƱ������.Enabled = txtƱ������.Enabled
    End Select
    
    '61808:������,2013-05-21
    '78751:���ϴ�,2015/08/24,����Ԥ��Ʊ�ݴ�ӡ��ʽ
    If Not (mlngModul = 1103 Or mlngModul = 1101) Then
        fraƱ�ݸ�ʽ.Visible = False
        fraƱ�ݸ�ʽ.Enabled = False
        fraPrepay.Height = fraƱ�ݸ�ʽ.Top + fraƱ�ݸ�ʽ.Height - fraPrepay.Top
        vsPrepay.Height = fraPrepay.Height - vsPrepay.Top - 90
    End If
    chkCheckBillNum.Visible = mlngModul = 1103
    txtƱ������.Visible = mlngModul = 1103
    updƱ������.Visible = mlngModul = 1103
    
    '����28130��27929
    chkHave.Visible = mlngModul = 1103
    'chkAllowOut.Visible = mlngModul = 1103
    chkAllowDept.Visible = mlngModul = 1103
    chkLedWelcome.Visible = mlngModul = 1103
    chkMustCard.Visible = mlngModul = 1101
    chkCardFeeCharge.Visible = mlngModul = 1102
    Exit Sub
errH:
    If ErrCenter() = 1 Then
         Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln���� = False
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
    zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

 
'����27390
Private Sub sTab_Click(PreviousTab As Integer)
    If sTab.Tab = 0 And chkItem(1).Enabled And chkItem(1).Visible Then chkItem(1).SetFocus
End Sub

Private Sub vs����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        With vs����
            Select Case Col
            Case .ColIndex("�̶����")
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "###0.00;-###0.00;;")
            Case .ColIndex("ѡ��")
            Case Else
            End Select
        End With
End Sub

Private Sub vs����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs����
        Select Case Col
        Case .ColIndex("�̶����")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vs����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs����
        If .Col >= .ColIndex("�̶����") And .Row = .Rows - 1 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
End Sub

Private Sub vs����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs����
        Select Case Col
        Case .ColIndex("�̶����")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vs����
        Select Case .Col
            Case .ColIndex("�̶����")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
        End Select
    End With
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub
 
Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub
Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("ҽ�ƿ����"))) = Trim(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub
Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("Ԥ������"))) = Trim(.Cell(flexcpData, i, .ColIndex("Ԥ������"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub
Private Sub chkCheckBillNum_Click()
    txtƱ������.Enabled = chkCheckBillNum.Enabled And chkCheckBillNum.Value = 1
    updƱ������.Enabled = txtƱ������.Enabled
End Sub
