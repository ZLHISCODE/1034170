VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPIVAParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Һ�������Ĳ�������"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11715
   Icon            =   "frmPIVAParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPRI 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   3120
      ScaleHeight     =   2055
      ScaleWidth      =   2535
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdYes 
         Height          =   360
         Left            =   720
         Picture         =   "frmPIVAParaSet.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1560
         Width           =   810
      End
      Begin VB.CommandButton cmdNO 
         Height          =   360
         Left            =   1560
         Picture         =   "frmPIVAParaSet.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1560
         Width           =   810
      End
      Begin MSComctlLib.ListView lvwPRI 
         Height          =   1305
         Left            =   120
         TabIndex        =   47
         ToolTipText     =   "˫���򰴻س���ȷ��"
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2302
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.PictureBox pic���� 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   2535
      TabIndex        =   40
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmd����Cancel 
         Height          =   360
         Left            =   1560
         Picture         =   "frmPIVAParaSet.frx":6F26
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1560
         Width           =   810
      End
      Begin VB.CommandButton cmd����Ok 
         Height          =   360
         Left            =   720
         Picture         =   "frmPIVAParaSet.frx":7070
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1560
         Width           =   810
      End
      Begin MSComctlLib.ListView LvwҩƷ���� 
         Height          =   1305
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "˫���򰴻س���ȷ��"
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2302
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
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "��������(&0)"
      TabPicture(0)   =   "frmPIVAParaSet.frx":D8C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��ӡ����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra��������"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "��������(&1)"
      TabPicture(1)   =   "frmPIVAParaSet.frx":D8DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAdd"
      Tab(1).Control(1)=   "cmdDel"
      Tab(1).Control(2)=   "vsfBatch"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "��������(&2)"
      TabPicture(2)   =   "frmPIVAParaSet.frx":D8FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabPrice"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdNext"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdLast"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "picprice"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "fraҽ������"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fra��Һ��"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "fra��Һ������"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblprice"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "��ҩ;��(&3)"
      TabPicture(3)   =   "frmPIVAParaSet.frx":D916
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chk��ҩ;��"
      Tab(3).Control(1)=   "Lvw��ҩ;��"
      Tab(3).Control(2)=   "lbl��ҩ;��"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "��Դ����(&4)"
      TabPicture(4)   =   "frmPIVAParaSet.frx":D932
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chk��Դ����"
      Tab(4).Control(1)=   "lvw��Դ����"
      Tab(4).Control(2)=   "lbl��Դ����"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "���ȼ�����(&5)"
      TabPicture(5)   =   "frmPIVAParaSet.frx":D94E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblpritip"
      Tab(5).Control(1)=   "vsfPri"
      Tab(5).Control(2)=   "vsfDept"
      Tab(5).Control(3)=   "cmdAddPri"
      Tab(5).Control(4)=   "cmdDelPri"
      Tab(5).Control(5)=   "cmdIN"
      Tab(5).Control(6)=   "chkAll"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "��������(&6)"
      TabPicture(6)   =   "frmPIVAParaSet.frx":D96A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblvoltip"
      Tab(6).Control(1)=   "vsfVolume"
      Tab(6).Control(2)=   "cmdVolDel"
      Tab(6).Control(3)=   "cmdVolAdd"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "����ҩƷ����(&7)"
      TabPicture(7)   =   "frmPIVAParaSet.frx":D986
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblMedi"
      Tab(7).Control(1)=   "vsfPrint"
      Tab(7).Control(2)=   "chkByMedi"
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "������ҩƷ(&8)"
      TabPicture(8)   =   "frmPIVAParaSet.frx":D9A2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblNoMedi"
      Tab(8).Control(1)=   "vsfNoMedi"
      Tab(8).ControlCount=   2
      Begin TabDlg.SSTab tabPrice 
         Height          =   1935
         Left            =   -68760
         TabIndex        =   109
         Top             =   3600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3413
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "��ҩ����"
         TabPicture(0)   =   "frmPIVAParaSet.frx":D9BE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "VSFPrice"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "��ҩ;��(ֻ֧�־���Ӫ������)"
         TabPicture(1)   =   "frmPIVAParaSet.frx":D9DA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "VSFPrice_��ҩ;��"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VSFlex8Ctl.VSFlexGrid VSFPrice 
            Height          =   1245
            Left            =   360
            TabIndex        =   110
            Top             =   480
            Width           =   3960
            _cx             =   6985
            _cy             =   2196
            Appearance      =   0
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
            BackColorSel    =   16771280
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   10329501
            GridColorFixed  =   10329501
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPIVAParaSet.frx":D9F6
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
            AccessibleDescription=   "200"
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFPrice_��ҩ;�� 
            Height          =   1245
            Left            =   -74520
            TabIndex        =   111
            Top             =   480
            Width           =   3600
            _cx             =   6350
            _cy             =   2196
            Appearance      =   0
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
            BackColorSel    =   16771280
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   10329501
            GridColorFixed  =   10329501
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPIVAParaSet.frx":DAA2
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
            AccessibleDescription=   "200"
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "����(&N)"
         Height          =   350
         Left            =   -67080
         TabIndex        =   108
         Top             =   5640
         Width           =   1100
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "����(&S)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -68760
         TabIndex        =   107
         Top             =   5640
         Width           =   1100
      End
      Begin VB.CheckBox chkByMedi 
         Caption         =   "�Ƿ�������õĳ���ҩƷ����ҩƷ���˲���"
         Height          =   255
         Left            =   -74880
         TabIndex        =   101
         Top             =   360
         Width           =   3855
      End
      Begin VB.PictureBox picprice 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -68760
         Picture         =   "frmPIVAParaSet.frx":DB4B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   90
         Top             =   3360
         Width           =   240
      End
      Begin VB.Frame fra�������� 
         Caption         =   "  �������� "
         Height          =   1695
         Left            =   120
         TabIndex        =   76
         Top             =   1320
         Width           =   11295
         Begin VB.CheckBox chkPeople 
            Caption         =   "��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա"
            Height          =   255
            Left            =   7320
            TabIndex        =   106
            Top             =   600
            Width           =   3495
         End
         Begin VB.CheckBox chkPacket 
            Caption         =   "���ҩƷ�ڷ��ͻ�����ȡ���÷�"
            Height          =   255
            Left            =   7320
            TabIndex        =   105
            Top             =   300
            Width           =   2895
         End
         Begin VB.CheckBox chkBeach 
            Caption         =   "���췢�͵�ҽ����������Һ��ȫ������������"
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Top             =   1080
            Width           =   3975
         End
         Begin VB.CheckBox chkOutPai 
            Caption         =   "��Ժ���˲������÷�"
            Height          =   255
            Left            =   4320
            TabIndex        =   99
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkMedi 
            Caption         =   "����ҩƷ��ҩƷ����ָ������"
            Height          =   255
            Left            =   4320
            TabIndex        =   93
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CheckBox chkSort 
            Caption         =   "��Һ�������Σ�ҩƷ��������"
            Height          =   255
            Left            =   4320
            TabIndex        =   92
            Top             =   840
            Width           =   2775
         End
         Begin VB.CheckBox chkMoney 
            Caption         =   "���÷Ѱ�������ȡ"
            Height          =   255
            Left            =   4320
            TabIndex        =   89
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chksend 
            Caption         =   "����ɨ��һ���Զ�����"
            Height          =   255
            Left            =   4320
            TabIndex        =   88
            Top             =   300
            Width           =   2895
         End
         Begin VB.CheckBox chkPackage 
            Caption         =   "��Һ��Һ����ҩ��������������"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   840
            Width           =   3975
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "����������״̬����ҩӡǩ����ҩ���ڣ�"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   560
            Width           =   3855
         End
         Begin VB.CheckBox chk�������� 
            Caption         =   "�����ֹ��������Σ���ҩӡǩ���ڣ�"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   300
            Width           =   3855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  �������Ŀⷿѡ�� "
         Height          =   615
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   11295
         Begin VB.CheckBox chkCheck 
            Caption         =   "��˸�ҩ��������ҽ��"
            Height          =   255
            Left            =   4680
            TabIndex        =   98
            Top             =   240
            Width           =   3855
         End
         Begin VB.ComboBox CboStore 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label lblStore 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Left            =   360
            TabIndex        =   73
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " ��Ƭ���� "
         Height          =   735
         Left            =   -68760
         TabIndex        =   67
         Top             =   1800
         Width           =   4575
         Begin VB.ComboBox cbo���� 
            Height          =   300
            ItemData        =   "frmPIVAParaSet.frx":1439D
            Left            =   1200
            List            =   "frmPIVAParaSet.frx":143AA
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl����1 
            Caption         =   "������ʾ"
            Height          =   195
            Left            =   360
            TabIndex        =   70
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lbl����2 
            Caption         =   "�ſ�Ƭ"
            Height          =   195
            Left            =   2160
            TabIndex        =   69
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Ӧ�������п��ҵ����ȼ�����"
         Height          =   250
         Left            =   -74880
         TabIndex        =   55
         Top             =   720
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CommandButton cmdIN 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   54
         Top             =   1800
         Width           =   1100
      End
      Begin VB.CommandButton cmdVolAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   51
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdVolDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   50
         Top             =   1560
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelPri 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   49
         Top             =   2400
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddPri 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   48
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CheckBox chk��Դ���� 
         Caption         =   "������Դ��������"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox chk��ҩ;�� 
         Caption         =   "������Һ��ҩ;������"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame fraҽ������ 
         Caption         =   "  ҽ������ѡ��  "
         Height          =   615
         Left            =   -68760
         TabIndex        =   22
         Top             =   2640
         Width           =   4575
         Begin VB.CheckBox chkҽ������ 
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkҽ������ 
            Caption         =   "����"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fra��Һ�� 
         Caption         =   "  ��ҩ���� "
         Height          =   4215
         Left            =   -74880
         TabIndex        =   20
         Top             =   1800
         Width           =   5895
         Begin VB.CheckBox chkAutoMode 
            Caption         =   "�Զ�����ʱ��Һ��������ֻ���������α䶯"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   3840
            Width           =   5535
         End
         Begin VB.CheckBox chkChangeDrug 
            Caption         =   "�������û�ҩ������Һ��������"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   3260
            Width           =   5535
         End
         Begin VB.CheckBox chkAutoBatch 
            Caption         =   "�����Զ������������Զ������󣬽����ٱ����ϴ����Σ�"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   2680
            Width           =   5535
         End
         Begin VB.CheckBox chkLastBatch 
            Caption         =   "�����ϴ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkTpn 
            Caption         =   "�������Ĳ����յľ���Ӫ��ҽ���ڲ�������"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   820
            Width           =   4335
         End
         Begin VB.CheckBox chkSpecial 
            Caption         =   "�Ա�ҩ����ȡҩ����Ժ��ҩ�����͵���������"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1400
            Width           =   4215
         End
         Begin VB.CheckBox chkBag 
            Caption         =   "����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����"
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   1980
            Width           =   5655
         End
      End
      Begin VB.Frame fra��Һ������ 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   10695
         Begin MSComCtl2.UpDown updDeff 
            Height          =   270
            Left            =   3600
            TabIndex        =   65
            Top             =   795
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDeff 
            Enabled         =   0   'False
            Height          =   270
            Left            =   3240
            TabIndex        =   64
            Text            =   "0"
            Top             =   795
            Width           =   375
         End
         Begin VB.CheckBox chkOpen 
            Caption         =   "���ý���ʱ��ο���"
            Height          =   180
            Left            =   360
            TabIndex        =   59
            Top             =   0
            Width           =   1935
         End
         Begin VB.PictureBox Picture5 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   4680
            Picture         =   "frmPIVAParaSet.frx":143B7
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
            Top             =   795
            Width           =   240
         End
         Begin VB.CheckBox chk����ҽ�� 
            Caption         =   "���յ��ռ���ǰ��ҽ��"
            Height          =   180
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   2175
         End
         Begin VB.PictureBox picHelpIcon 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   4680
            Picture         =   "frmPIVAParaSet.frx":1AC09
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   25
            Top             =   240
            Width           =   240
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   315
            Left            =   960
            TabIndex        =   61
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   80871426
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   315
            Left            =   3240
            TabIndex        =   63
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   80871426
            CurrentDate     =   36985
         End
         Begin VB.Label lblDeff 
            Caption         =   "Сʱ��"
            Height          =   255
            Left            =   2595
            TabIndex        =   66
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblEnd 
            Caption         =   "����ʱ��"
            Height          =   255
            Left            =   2400
            TabIndex        =   62
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblBegin 
            Caption         =   "��ʼʱ��"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbl����ҽ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ѡʱ�������Ľ���������ʱ��������ĵ���ִ�е�ҽ����"
            Height          =   180
            Left            =   5040
            TabIndex        =   39
            Top             =   795
            Width           =   4680
         End
         Begin VB.Label lblʱ����� 
            AutoSize        =   -1  'True
            Caption         =   "ҽ�����Ͳ��ڸ�ʱ�����Һҽ�������ٲ�����Һ����"
            Height          =   180
            Left            =   5040
            TabIndex        =   21
            Top             =   240
            Width           =   4140
         End
      End
      Begin VB.Frame fra��ӡ���� 
         Caption         =   "  ��ӡ���� "
         Height          =   2840
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   11295
         Begin VB.ComboBox cboSum 
            Height          =   300
            Left            =   6120
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   885
            Width           =   2415
         End
         Begin VB.ComboBox cboNum 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   2460
            Width           =   2415
         End
         Begin VB.ComboBox cbo��ǩ��ӡ 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   540
            Width           =   2415
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "��ҩ��"
            Height          =   180
            Index           =   1
            Left            =   1440
            TabIndex        =   34
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "��ҩ��"
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cbo��ǩ��ӡ 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   180
            Width           =   2415
         End
         Begin VB.CommandButton cmd��ӡ���� 
            Caption         =   "��ӡ����(&P)"
            Height          =   345
            Left            =   3960
            TabIndex        =   18
            Top             =   2055
            Width           =   1155
         End
         Begin VB.ComboBox cboƱ������ 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2085
            Width           =   2415
         End
         Begin VB.CheckBox chkManPrint 
            Caption         =   "�����ֹ����ƴ�ӡƿǩ���ɽ��в���"
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   915
            Width           =   3375
         End
         Begin VB.ComboBox cbo���͵� 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1695
            Width           =   2415
         End
         Begin VB.ComboBox cbo��ҩ�� 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�����ڰ�ҩ����ҩ���ӡ"
            Height          =   180
            Left            =   3960
            TabIndex        =   87
            Top             =   2520
            Width           =   1980
         End
         Begin VB.Label lblSumPrint 
            AutoSize        =   -1  'True
            Caption         =   "��ӡ��ǩ��"
            Height          =   180
            Left            =   5040
            TabIndex        =   86
            Top             =   952
            Width           =   900
         End
         Begin VB.Label lblSum 
            AutoSize        =   -1  'True
            Caption         =   "���ܱ���"
            Height          =   180
            Left            =   8640
            TabIndex        =   85
            Top             =   952
            Width           =   720
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "ƿǩ��ӡ����"
            Height          =   180
            Left            =   180
            TabIndex        =   74
            Top             =   2520
            Width           =   1080
         End
         Begin VB.Label lblPrintLabel 
            AutoSize        =   -1  'True
            Caption         =   "ƿǩ��ӡ��ʽ"
            Height          =   180
            Left            =   180
            TabIndex        =   36
            Top             =   405
            Width           =   1080
         End
         Begin VB.Label lblƱ�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ�ݺͱ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   17
            Top             =   2145
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "ƿǩ�ֹ���ӡ"
            Height          =   180
            Left            =   180
            TabIndex        =   15
            Top             =   952
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "���ܷ����嵥"
            Height          =   180
            Left            =   3960
            TabIndex        =   13
            Top             =   1755
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ�����嵥"
            Height          =   180
            Left            =   3960
            TabIndex        =   12
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "����ȷ�Ϻ�"
            Height          =   180
            Left            =   360
            TabIndex        =   10
            Top             =   1755
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��ҩȷ�Ϻ�"
            Height          =   180
            Left            =   360
            TabIndex        =   8
            Top             =   1380
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   -64800
         TabIndex        =   5
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   -64800
         TabIndex        =   4
         Top             =   1560
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   3
         Top             =   840
         Width           =   9960
         _cx             =   17568
         _cy             =   8864
         Appearance      =   0
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
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":2145B
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
      Begin MSComctlLib.ListView Lvw��ҩ;�� 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   8387
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
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lvw��Դ���� 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   29
         Top             =   1200
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   8387
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
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   4785
         Left            =   -74880
         TabIndex        =   52
         Top             =   1080
         Width           =   2400
         _cx             =   4233
         _cy             =   8440
         Appearance      =   0
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":215EF
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPri 
         Height          =   4785
         Left            =   -72360
         TabIndex        =   53
         Top             =   1080
         Width           =   7560
         _cx             =   13335
         _cy             =   8440
         Appearance      =   0
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":21685
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
      Begin VSFlex8Ctl.VSFlexGrid vsfVolume 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   58
         Top             =   840
         Width           =   10080
         _cx             =   17780
         _cy             =   8864
         Appearance      =   0
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":2173E
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   80
         Top             =   960
         Width           =   10080
         _cx             =   17780
         _cy             =   8864
         Appearance      =   0
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":217E2
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
      Begin VSFlex8Ctl.VSFlexGrid vsfNoMedi 
         Height          =   5145
         Left            =   -74880
         TabIndex        =   82
         Top             =   840
         Width           =   10080
         _cx             =   17780
         _cy             =   9075
         Appearance      =   0
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":2184B
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
      Begin VB.Label lblprice 
         AutoSize        =   -1  'True
         Caption         =   "�����շ�����"
         Height          =   180
         Left            =   -68400
         TabIndex        =   91
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label lblNoMedi 
         Caption         =   "�����������Ĳ��������õ�ҩƷ"
         Height          =   255
         Left            =   -74880
         TabIndex        =   83
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblMedi 
         Caption         =   "���ó���ҩƷ������Һ��������԰�ҩƷ���й��˺�����"
         Height          =   255
         Left            =   -74880
         TabIndex        =   81
         Top             =   720
         Width           =   7935
      End
      Begin VB.Label lblvoltip 
         AutoSize        =   -1  'True
         Caption         =   "����ĳ�����ҵ�������ĳ�����ο�����ҩ������"
         Height          =   180
         Left            =   -74880
         TabIndex        =   57
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label lblpritip 
         AutoSize        =   -1  'True
         Caption         =   "��������ͬ��������ͬ��ҩƷ�����ȼ�"
         Height          =   180
         Left            =   -74880
         TabIndex        =   56
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label lbl��Դ���� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��ѡ��������Һҽ������ʱ������˵����ڲ���û��ѡ���򲻻������Һ���ݡ�"
         Height          =   180
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   7200
      End
      Begin VB.Label lbl��ҩ;�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��ѡ��������Һ��ĸ�ҩ;������Һҽ������ʱ���ҽ���ĸ�ҩ;��û��ѡ���򲻻������Һ���ݡ�"
         Height          =   180
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   8640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�����������Ĺ�������(0������Ϊ�������δ��ڣ�������ҩʱ�䷶Χ����)"
         Height          =   180
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   5850
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10320
      TabIndex        =   1
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9000
      TabIndex        =   0
      Top             =   6360
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgLvwSel 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmPIVAParaSet.frx":218B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":21BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":21EE8
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":2223A
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   5880
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPIVAParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String                              'Ȩ�޴�
Public mlng�ⷿid As Long
Private mblnSetPara As Boolean
Private mRsDept As Recordset
Private mRsPC As Recordset
Private mRsWay As Recordset
Private mRsType As Recordset
Private mRsPrice As Recordset
Private mintRow As Integer
Private mintCol As Integer
Private mintPri As Integer
Private mblnPrice As Boolean
Private mblnEdit As Boolean     '�Ƿ�༭���ȼ�
Private mrs�շ���Ŀ As Recordset
Private Sub LoadStore()
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.id,B.���� From ��������˵�� A,���ű� B" & _
    " Where A.����ID=B.ID And A.��������='��������' And B.Id In (Select ����id From ������Ա Where ��Աid = [1])"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������ĵĲ���", glngUserId)
    
    With Me.CboStore
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadParams()
    Dim int��ҩ�� As Integer
    Dim int��ҩ�� As Integer
    Dim int�������� As Integer
    Dim int�ϴ����� As Integer
    Dim int������� As Integer
    Dim strAutoPrint As String
    Dim intManPrint As Integer
    Dim str��ֹʱ�� As String
    Dim intҽ������ As Integer
    Dim dbl��Һ�� As Double
    Dim str����ҺҩƷ���� As String
    Dim str��Һ��ҩ;�� As String
    Dim str��Դ���� As String
    Dim rsData As ADODB.Recordset
    Dim str����ҽ�� As String
    Dim IntCount As Integer
    Dim intOpen As Integer
    Dim lng����ID As Long
    Dim IntLocate As Integer
    Dim dateNow As Date
    Dim intNum As Integer
    Dim int��ҩ���� As Integer
    Dim i As Integer
    Dim int���� As Integer
    Dim intTPN As Integer
    Dim intSpecial As Integer
    
    On Error GoTo errHandle
    '����
    int��ҩ�� = Val(zlDatabase.GetPara("��ҩ���ӡ", glngSys, 1345, 0, Array(Label3, cbo��ҩ��, Label5), mblnSetPara))
    int��ҩ�� = Val(zlDatabase.GetPara("���ͺ��ӡ", glngSys, 1345, 0, Array(Label4, cbo���͵�, Label6), mblnSetPara))
    int�������� = Val(zlDatabase.GetPara("��������", glngSys, 1345, 0, Array(chk��������), mblnSetPara))
    int������� = Val(zlDatabase.GetPara("�������", glngSys, 1345, 0, Array(chk�������), mblnSetPara))
    strAutoPrint = zlDatabase.GetPara("ƿǩ�Զ���ӡ", glngSys, 1345, "00|00", Array(lblPrintLabel, chkPrintLabelStep(0), chkPrintLabelStep(1)), mblnSetPara)
    intManPrint = Val(zlDatabase.GetPara("ƿǩ�ֹ���ӡ", glngSys, 1345, "0", Array(Label8, chkManPrint), mblnSetPara))
    IntCount = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\��Һ��Ƭ", "��Ƭ����", 3))
    int��ҩ���� = Val(zlDatabase.GetPara("��Һ��Һ����ҩ��������������", glngSys, 1345, 0, Array(chkPackage), mblnSetPara))
    int���� = Val(zlDatabase.GetPara("��ӡ��ǩ���Ƿ��ӡ���ܱ���", glngSys, 1345, 0, Array(lblSumPrint, cboSum, lblSum), mblnSetPara))
    
    '��������
    str��ֹʱ�� = zlDatabase.GetPara("������ֹʱ��", glngSys, 1345, "", Array(lblBegin, dtpBegin, lblEnd, dtpEnd), mblnSetPara)
    str����ҽ�� = zlDatabase.GetPara("�����յ��ռ���ǰҽ��", glngSys, 1345, 0, Array(chk����ҽ��, txtDeff, updDeff, lblDeff), mblnSetPara)
    intҽ������ = Val(zlDatabase.GetPara("ҽ������", glngSys, 1345, 1, Array(chkҽ������(0), chkҽ������(1)), mblnSetPara))
'    dbl��Һ�� = Val(zlDatabase.GetPara("ͬ������Һ����", glngSys, 1345, "", Array(chk��Һ������, txt��Һ����, lbl��������˵��), mblnSetPara))
'    str����ҺҩƷ���� = zldatabase.GetPara("����ҺҩƷ����", glngSys, 1345, "", Array(chkҩƷ����, txtҩƷ����, cmdҩƷ����), mblnSetPara)
    int�ϴ����� = Val(zlDatabase.GetPara("�����ϴ�����", glngSys, 1345, 0, Array(chkLastBatch), mblnSetPara))
    intOpen = Val(zlDatabase.GetPara("���ý���ʱ�����", glngSys, 1345, 0, Array(chkOpen), mblnSetPara))
    lng����ID = Val(zlDatabase.GetPara("��������", glngSys, 1345, 0, Array(CboStore, lblStore), mblnSetPara))
    intNum = Val(zlDatabase.GetPara("ƿǩ��ӡ����", glngSys, 1345, 1, Array(lblNum, cboNum), mblnSetPara))
    intTPN = Val(zlDatabase.GetPara("�������Ĳ����յľ���Ӫ��ҽ���ڲ�������", glngSys, 1345, 0, Array(chkTpn), mblnSetPara))
    intSpecial = Val(zlDatabase.GetPara("��������ҩƷ�����͵���������", glngSys, 1345, 0, Array(chkSpecial), mblnSetPara))
    Me.chksend.Value = Val(zlDatabase.GetPara("ɨ����ƿǩ���Զ�����", glngSys, 1345, 0, Array(chksend), mblnSetPara))
    Me.chkMoney.Value = Val(zlDatabase.GetPara("���÷Ѱ�������ȡ", glngSys, 1345, 0, Array(chkMoney), mblnSetPara))
    Me.chkBag.Value = Val(zlDatabase.GetPara("����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����", glngSys, 1345, 0, Array(chkBag), mblnSetPara))
    Me.chkSort.Value = Val(zlDatabase.GetPara("�����Σ�ҩƷ����", glngSys, 1345, 0, Array(chkSort), mblnSetPara))
    Me.chkMedi.Value = Val(zlDatabase.GetPara("����ҩƷ��ҩƷ����ָ������", glngSys, 1345, 0, Array(chkMedi), mblnSetPara))
    Me.chkCheck.Value = Val(zlDatabase.GetPara("��˸�ҩ������������", glngSys, 1345, 0, Array(chkCheck), mblnSetPara))
    Me.chkOutPai.Value = Val(zlDatabase.GetPara("��Ժ���˲������÷�", glngSys, 1345, 0, Array(chkOutPai), mblnSetPara))
    Me.chkAutoBatch.Value = Val(zlDatabase.GetPara("�����Զ�����", glngSys, 1345, 0, Array(chkAutoBatch), mblnSetPara))
    Me.chkByMedi.Value = Val(zlDatabase.GetPara("�Ƿ����õĳ���ҩƷ����ҩƷ���˲���", glngSys, 1345, 0, Array(chkByMedi), mblnSetPara))
    Me.chkChangeDrug.Value = Val(zlDatabase.GetPara("�������û�ҩ������Һ��������", glngSys, 1345, 0, Array(chkChangeDrug), mblnSetPara))
    Me.chkAutoMode.Value = Val(zlDatabase.GetPara("�Զ�����ʱ��Һ��������ֻ���������α䶯", glngSys, 1345, 0, Array(chkAutoMode), mblnSetPara))
    Me.chkBeach.Value = Val(zlDatabase.GetPara("���췢�͵�ҽ����������Һ��ȫ������������", glngSys, 1345, 0, Array(chkBeach), mblnSetPara))
    Me.chkPeople.Value = Val(zlDatabase.GetPara("��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա", glngSys, 1345, 0, Array(chkPeople), mblnSetPara))
    Me.chkPacket.Value = Val(zlDatabase.GetPara("���ҩƷ�ڷ��ͻ�����ȡ���÷�", glngSys, 1345, 0, Array(chkPacket), mblnSetPara))
    
    '��ҩ;��
    str��Һ��ҩ;�� = zlDatabase.GetPara("��Һ��ҩ;��", glngSys, 1345, "", Array(chk��ҩ;��, Lvw��ҩ;��), mblnSetPara)
    
    '��Դ����
    str��Դ���� = zlDatabase.GetPara("��Դ����", glngSys, 1345, "", Array(chk��Դ����, lvw��Դ����), mblnSetPara)
    
    If lng����ID <> 0 Then                                  '��λҩ��
        '�����ڸ�ҩ������ʾ
        For IntLocate = 0 To Me.CboStore.ListCount - 1
            If Me.CboStore.ItemData(IntLocate) = lng����ID Then
                Me.CboStore.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (CboStore.ListCount - 1) Then
            MsgBox "�����������������ģ�ԭ�����õ�����������ʧЧ����", vbInformation, gstrSysName
            If CboStore.ListCount >= 1 Then CboStore.ListIndex = 0
        End If
    Else
        MsgBox "�������������ģ�", vbInformation, gstrSysName
    End If
    
    Me.chkOpen.Value = intOpen
    
    If InStr(1, str��ֹʱ��, "|") > 0 Then
        Me.dtpBegin.Value = Mid(str��ֹʱ��, 1, InStr(1, str��ֹʱ��, "|") - 1)
        Me.dtpEnd.Value = Mid(str��ֹʱ��, InStr(1, str��ֹʱ��, "|") + 1)
    End If
    
    Me.chk����ҽ��.Value = Mid(str����ҽ��, 1, 1)
    If InStr(1, str����ҽ��, "|") > 1 Then
        Me.txtDeff.Text = Mid(str����ҽ��, 3)
    Else
        Me.txtDeff.Text = 0
    End If
    
    ''��������
    If int��ҩ�� >= 0 And int��ҩ�� <= cbo��ҩ��.ListCount - 1 Then
        cbo��ҩ��.ListIndex = int��ҩ��
    End If
    
    If int���� >= 0 And int���� <= cboSum.ListCount - 1 Then
        cboSum.ListIndex = int����
    End If
    
    If int��ҩ�� >= 0 And int��ҩ�� <= cbo��ҩ��.ListCount - 1 Then
        cbo���͵�.ListIndex = int��ҩ��
    End If
    
    If int�������� >= 0 And int�������� <= 1 Then
        chk��������.Value = int��������
    End If
        
    If int������� >= 0 And int������� <= 1 Then
        chk�������.Value = int�������
    End If
    
    If int��ҩ���� >= 0 And int��ҩ���� <= 1 Then
        chkPackage.Value = int��ҩ����
    End If
    
    If InStr(1, strAutoPrint, "|") = 0 Or Len(strAutoPrint) <> 5 Then
        strAutoPrint = "00|00"
    End If
    
    If Mid(strAutoPrint, 1, 1) = 1 Then
        chkPrintLabelStep(0).Value = 1
        If Val(Mid(strAutoPrint, 2, 1)) = 1 Then
            cbo��ǩ��ӡ(0).ListIndex = 1
        Else
            cbo��ǩ��ӡ(0).ListIndex = 0
        End If
    End If
    
    If Mid(strAutoPrint, 4, 1) = 1 Then
        chkPrintLabelStep(1).Value = 1
        If Val(Mid(strAutoPrint, 5, 1)) = 1 Then
            cbo��ǩ��ӡ(1).ListIndex = 1
        Else
            cbo��ǩ��ӡ(1).ListIndex = 0
        End If
    End If
    
    cbo��ǩ��ӡ(0).Enabled = chkPrintLabelStep(0).Enabled And (chkPrintLabelStep(0).Value = 1)
    cbo��ǩ��ӡ(1).Enabled = chkPrintLabelStep(1).Enabled And (chkPrintLabelStep(1).Value = 1)
    
    vsfVolume.Enabled = mblnSetPara
    vsfPrint.Enabled = mblnSetPara
    vsfNoMedi.Enabled = mblnSetPara
    vsfPri.Enabled = mblnSetPara
    cmdAddPri.Enabled = mblnSetPara
    cmdIN.Enabled = mblnSetPara
    cmdDelPri.Enabled = mblnSetPara
    cmdVolAdd.Enabled = mblnSetPara
    cmdVolDel.Enabled = mblnSetPara
    
    If intManPrint < 0 Or intManPrint > 1 Then
        chkManPrint.Value = 0
    Else
        chkManPrint.Value = intManPrint
    End If
    
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
    
    '��Ƭ����
    Me.cbo����.Text = IIf(IntCount = 0, 3, IntCount)
    
    Me.cboNum.Text = IIf(intNum = 0, 3, intNum)
    
    ''��������
'    chk��ֹʱ��.Value = IIf(str��ֹʱ�� = "", 0, 1)
'    dtpTime.Enabled = (chk��ֹʱ��.Value = 1)
'    If str��ֹʱ�� <> "" Then
'        If IsDate(str��ֹʱ��) = True Then
'            dtpTime.Value = str��ֹʱ��
'        End If
'    End If
    
'    txt��Һ����.Text = ""
'    chk��Һ������.Value = IIf(dbl��Һ�� = 0, 0, 1)
'    txt��Һ����.Enabled = (chk��Һ������.Value = 1)
    If dbl��Һ�� > 0 Then
'        txt��Һ����.Text = dbl��Һ��
    End If
    
    If intҽ������ = 0 Then
        chkҽ������(0).Value = 1
        chkҽ������(1).Value = 1
    ElseIf intҽ������ = 1 Then
        chkҽ������(0).Value = 1
        chkҽ������(1).Value = 0
    ElseIf intҽ������ = 2 Then
        chkҽ������(0).Value = 0
        chkҽ������(1).Value = 1
    End If
    
'    chkҩƷ����.Value = IIf(str����ҺҩƷ���� = "", 0, 1)
'    txtҩƷ����.Text = str����ҺҩƷ����
'    txtҩƷ����.Enabled = (chkҩƷ����.Value = 1)
'    cmdҩƷ����.Enabled = (chkҩƷ����.Value = 1)
    
    chkLastBatch.Value = IIf(int�ϴ����� = 0, 0, 1)
    
    '����Ӫ��ҩ�ﴦ�÷�ʽ
    If intTPN >= 0 And intTPN <= 1 Then
        chkTpn.Value = intTPN
    End If
    
    '����ҩƷ����
    If intSpecial >= 0 And intSpecial <= 1 Then
        chkSpecial.Value = intSpecial
    End If
    
    ''��ҩ;��
    gstrSQL = "Select ID, ���� as �÷� ,�걾��λ As ���� From ������ĿĿ¼ Where ���='E' And ��������='2'And (�������=2 Or �������=3) And ִ�з��� = 1 " & _
            " And (����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or ����ʱ�� Is Null) Order by ���� "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ;��")
    
    With Lvw��ҩ;��
        .ListItems.Clear
        Do While Not rsData.EOF
            .ListItems.Add , "_" & rsData!Id, rsData!�÷�, 1, 1
            If InStr(1, "," & str��Һ��ҩ;�� & ",", "," & rsData!Id & ",") > 0 Then
                .ListItems(.ListItems.count).Checked = True
            End If
            rsData.MoveNext
        Loop
    End With
    
    If str��Һ��ҩ;�� <> "" Then
        chk��ҩ;��.Value = 1
    End If
    
    Lvw��ҩ;��.Enabled = chk��ҩ;��.Enabled And (chk��ҩ;��.Value = 1)
    Lvw��ҩ;��.BackColor = IIf(Lvw��ҩ;��.Enabled, &H80000005, &H8000000F)
    
    ''��Դ����
    gstrSQL = "Select ���� || '-' || ���� ����, Id " & _
            " From ���ű� " & _
            " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And Id In (Select ����id From ��������˵�� Where �������� = '����' And ������� In (2,3)) And " & _
            " (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By ���� || '-' || ���� "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load��Դ����")

    With rsData
        lvw��Դ����.ListItems.Clear
        Do While Not .EOF
            lvw��Դ����.ListItems.Add , "_" & !Id, !����, 1, 1
            If str��Դ���� <> "" Then
                If InStr("," & str��Դ���� & ",", "," & CStr(!Id) & ",") > 0 Then
                    lvw��Դ����.ListItems("_" & !Id).Checked = True
                End If
            End If
            .MoveNext
        Loop
    End With
    
    If str��Դ���� <> "" Then
        chk��Դ����.Value = 1
    End If
    lvw��Դ����.Enabled = chk��Դ����.Enabled And (chk��Դ����.Value = 1)
    lvw��Դ����.BackColor = IIf(lvw��Դ����.Enabled, &H80000005, &H8000000F)
    
    '����ҩƷ��ӡ����
    gstrSQL = "select ҩƷid,���� from ��Һ���ȴ�ӡҩƷ"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "LoadҩƷ")
    
    Me.vsfPrint.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("ҩƷid")) = rsData!ҩƷID
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("ҩƷ���������")) = rsData!����
       
       rsData.MoveNext
    Next
    
    
    '��Һ������ҩƷ
    gstrSQL = "select ҩƷid,���� from ��Һ������ҩƷ"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "LoadҩƷ")
    
    Me.vsfNoMedi.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("ҩƷid")) = rsData!ҩƷID
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("ҩƷ���������")) = rsData!����
       
       rsData.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub CboStore_Click()
    Call LoadBatchSet
    Call loadVolume
End Sub

Private Sub chkAll_Click()
    If mblnEdit Then
        If MsgBox("�뱣�����õ����ȼ����л����Һ����������ȼ����ý�ʧЧ���Ƿ��л���", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
           If Me.chkAll.Value = 0 Then
                Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
                Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
                Me.vsfDept.Visible = True
                Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
            Else
                Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
                Me.vsfPri.Left = Me.vsfDept.Left
                Me.vsfDept.Visible = False
                
                Call LoadVsfPRI(0)
            End If
            mblnEdit = False
            
        End If
    Else
        If Me.chkAll.Value = 0 Then
            Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
            Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
            Me.vsfDept.Visible = True
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
        Else
            Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
            Me.vsfPri.Left = Me.vsfDept.Left
            Me.vsfDept.Visible = False
            
            Call LoadVsfPRI(0)
        End If
    End If
End Sub

Private Sub chkManPrint_Click()
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
End Sub

Private Sub chkOpen_Click()
    Me.dtpBegin.Enabled = (Me.chkOpen.Value = 1)
    Me.dtpEnd.Enabled = (Me.chkOpen.Value = 1)
    Me.chk����ҽ��.Enabled = (Me.chkOpen.Value = 1)
    Me.updDeff.Enabled = (Me.chkOpen.Value = 1)
End Sub

Private Sub chkPrintLabelStep_Click(Index As Integer)
    cbo��ǩ��ӡ(Index).Enabled = (chkPrintLabelStep(Index).Value = 1)
End Sub


Private Sub chk��ҩ;��_Click()
    Lvw��ҩ;��.Enabled = (chk��ҩ;��.Value = 1)
    Lvw��ҩ;��.BackColor = IIf(Lvw��ҩ;��.Enabled, &H80000005, &H8000000F)
End Sub

Private Sub chk��Դ����_Click()
    lvw��Դ����.Enabled = (chk��Դ����.Value = 1)
    lvw��Դ����.BackColor = IIf(lvw��Դ����.Enabled, &H80000005, &H8000000F)
End Sub

'Private Sub chk��Һ������_Click()
'    If chk��Һ������.Value = 1 Then
'        txt��Һ����.Enabled = True
'    Else
'        txt��Һ����.Enabled = False
'    End If
'End Sub


'Private Sub chkҩƷ����_Click()
'    txtҩƷ����.Enabled = (chkҩƷ����.Value = 1)
'    cmdҩƷ����.Enabled = (chkҩƷ����.Value = 1)
'
'    If chkҩƷ����.Value = 0 And pic����.Visible = True Then
'        Call cmd����Cancel_Click
'    End If
'End Sub

Private Sub chkҽ������_Click(Index As Integer)
    If chkҽ������(0).Value = 0 And chkҽ������(1).Value = 0 Then
        chkҽ������(Index).Value = 1
    End If
End Sub

Private Sub cmdAdd_Click()
    With vsfBatch
        If .rows > 2 Then
            If Trim(.TextMatrix(.rows - 1, .ColIndex("����ʱ�俪ʼ"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("����ʱ�����"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("��ҩʱ�俪ʼ"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("��ҩʱ�����"))) = "" Then
                Exit Sub
            End If
        End If
        
        .rows = .rows + 1
        
        If .Row >= 2 Then
            .TextMatrix(.rows - 1, .ColIndex("����")) = Mid(.TextMatrix(.rows - 2, .ColIndex("����")), 1, Len(.TextMatrix(.rows - 2, .ColIndex("����"))) - 1) + 1 & "#"
        Else
            .TextMatrix(.rows - 1, .ColIndex("����")) = "0#"
        End If
        .TextMatrix(.rows - 1, .ColIndex("����")) = "��"
    End With
End Sub

Private Sub cmdAddPri_Click()
    If Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("��ҩ����")) <> "" And Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("Ƶ��")) <> "" Then
        Me.vsfPri.rows = Me.vsfPri.rows + 1
        Me.vsfPri.RowHeight(Me.vsfPri.rows - 1) = 250
        Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, vsfPri.ColIndex("���")) = Me.vsfPri.rows - 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim lngRow As Long
    
    With vsfBatch
        If .Row > 1 Then
            lngRow = .Row
            If MsgBox("�Ƿ�ɾ������(" & .TextMatrix(.Row, .ColIndex("����")) & ")��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            .Redraw = flexRDNone
            
            .RemoveItem .Row
            
            '�������κ�
            For lngRow = lngRow To .rows - 1
                .TextMatrix(lngRow, .ColIndex("����")) = Mid(.TextMatrix(lngRow, .ColIndex("����")), 1, Len(.TextMatrix(lngRow, .ColIndex("����"))) - 1) - 1 & "#"
            Next
            
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub cmdDelPri_Click()
    Dim i As Integer
    Dim intRow As Integer
    
    If Me.vsfPri.Row = 0 Then Exit Sub
    intRow = Me.vsfPri.Row
    Me.vsfPri.RemoveItem Me.vsfPri.Row
    
    '�������
    For i = intRow To Me.vsfPri.rows - 1
        Me.vsfPri.TextMatrix(i, Me.vsfPri.ColIndex("���")) = i
    Next
    
    mblnEdit = True
End Sub

Private Sub cmdIN_Click()
    Dim IntCount As Integer
    Dim lngRow As Long
    
    If mblnSetPara Then
         '�������ȼ�����
        With vsfPri
            IntCount = 1
            
            If .rows = 1 Then
                gstrSQL = "Zl_��ҺҩƷ���ȼ�_Save("
                '����id
                gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))) & "'"
                gstrSQL = gstrSQL & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�������ȼ�")
            End If
            
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("��ҩ����")) <> "" And .TextMatrix(lngRow, .ColIndex("Ƶ��")) <> "" Then
                    
                    gstrSQL = "Zl_��ҺҩƷ���ȼ�_Save("
                    '����id
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))) & "',"
                    '��������
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, "���п���", vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("��������"))) & "',"
                    '��ҩ����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("��ҩ����")) & "',"
                    'Ƶ��
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("Ƶ��")) & "',"
                    '��Ч
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("��Ч"))) & ","
                    '���ȼ�
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("���")))
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������ȼ�")
                    IntCount = IntCount + 1
                End If
            Next
        End With
    End If
    
    mblnEdit = False
End Sub

Private Sub cmdLast_Click()
    Dim intRow As Integer
    Dim str��ҩ���� As String
    Dim str�շ���Ŀ As String
    Dim lng��Ŀid As Long
    
    With VSFPrice
        intRow = .Row
        If intRow < 2 Then Exit Sub
        lng��Ŀid = .TextMatrix(.Row - 1, .ColIndex("��Ŀid"))
        str��ҩ���� = .TextMatrix(.Row - 1, .ColIndex("��ҩ����"))
        str�շ���Ŀ = .TextMatrix(.Row - 1, .ColIndex("�շ���Ŀ"))
        .TextMatrix(.Row - 1, .ColIndex("��Ŀid")) = .TextMatrix(.Row, .ColIndex("��Ŀid"))
        .TextMatrix(.Row - 1, .ColIndex("��ҩ����")) = .TextMatrix(.Row, .ColIndex("��ҩ����"))
        .TextMatrix(.Row - 1, .ColIndex("�շ���Ŀ")) = .TextMatrix(.Row, .ColIndex("�շ���Ŀ"))
        
        
        .TextMatrix(.Row, .ColIndex("��Ŀid")) = lng��Ŀid
        .TextMatrix(.Row, .ColIndex("��ҩ����")) = str��ҩ����
        .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) = str�շ���Ŀ
        
        .Row = intRow - 1
    End With
End Sub

Private Sub cmdNext_Click()
    Dim intRow As Integer
    Dim str��ҩ���� As String
    Dim str�շ���Ŀ As String
    Dim lng��Ŀid As Long
    
    With VSFPrice
        intRow = .Row
        If intRow = .rows - 1 Then Exit Sub
        lng��Ŀid = .TextMatrix(.Row + 1, .ColIndex("��Ŀid"))
        str��ҩ���� = .TextMatrix(.Row + 1, .ColIndex("��ҩ����"))
        str�շ���Ŀ = .TextMatrix(.Row + 1, .ColIndex("�շ���Ŀ"))
        .TextMatrix(.Row + 1, .ColIndex("��Ŀid")) = .TextMatrix(.Row, .ColIndex("��Ŀid"))
        .TextMatrix(.Row + 1, .ColIndex("��ҩ����")) = .TextMatrix(.Row, .ColIndex("��ҩ����"))
        .TextMatrix(.Row + 1, .ColIndex("�շ���Ŀ")) = .TextMatrix(.Row, .ColIndex("�շ���Ŀ"))
        
        
        .TextMatrix(.Row, .ColIndex("��Ŀid")) = lng��Ŀid
        .TextMatrix(.Row, .ColIndex("��ҩ����")) = str��ҩ����
        .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) = str�շ���Ŀ
        
        .Row = intRow + 1
    End With
End Sub

Private Sub cmdNo_Click()
    picPRI.Visible = False
    CmdOK.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Sub cmdOk_Click()
    Dim strInput As String
    Dim lngRow As Long
    Dim intҽ������ As Integer
    Dim str��ҩ;�� As String
    Dim str��Դ���� As String
    Dim strPrintLabel As String
    Dim IntCount As Integer
    Dim i As Integer
    Dim n As Integer
    
    On Error GoTo errHandle
    
    If chkҽ������(0).Value = 1 And chkҽ������(1).Value = 1 Then
        intҽ������ = 0
    ElseIf chkҽ������(0).Value = 1 Then
        intҽ������ = 1
    ElseIf chkҽ������(1).Value = 1 Then
        intҽ������ = 2
    End If
    
    '��Դ����
    With Me.Lvw��ҩ;��
        For lngRow = 1 To .ListItems.count
            If .ListItems(lngRow).Checked Then
                If str��ҩ;�� = "" Then
                    str��ҩ;�� = Mid(.ListItems(lngRow).Key, 2)
                Else
                    str��ҩ;�� = str��ҩ;�� & "," & Mid(.ListItems(lngRow).Key, 2)
                End If
            End If
        Next
    End With
    
    '��Դ����
    With Me.lvw��Դ����
        For lngRow = 1 To .ListItems.count
            If .ListItems(lngRow).Checked Then
                
                str��Դ���� = str��Դ���� & Mid(.ListItems(lngRow).Key, 2) & ","
            End If
        Next
    End With
    
    'ƿǩ��ӡ��ʽ
    If chkPrintLabelStep(0).Value = 0 Then
        strPrintLabel = "00"
    Else
        strPrintLabel = "1" & cbo��ǩ��ӡ(0).ListIndex
    End If
    strPrintLabel = strPrintLabel & "|"
    If chkPrintLabelStep(1).Value = 0 Then
        strPrintLabel = strPrintLabel & "00"
    Else
        strPrintLabel = strPrintLabel & "1" & cbo��ǩ��ӡ(1).ListIndex
    End If
    
    '���湫����˽�в���
    '��������
    zlDatabase.SetPara "��ҩ���ӡ", cbo��ҩ��.ListIndex, glngSys, 1345
    zlDatabase.SetPara "���ͺ��ӡ", cbo���͵�.ListIndex, glngSys, 1345
    zlDatabase.SetPara "��������", chk��������.Value, glngSys, 1345
    zlDatabase.SetPara "�������", chk�������.Value, glngSys, 1345
    zlDatabase.SetPara "ƿǩ�Զ���ӡ", strPrintLabel, glngSys, 1345
    zlDatabase.SetPara "ƿǩ�ֹ���ӡ", chkManPrint.Value, glngSys, 1345
    zlDatabase.SetPara "��Һ��Һ����ҩ��������������", chkPackage.Value, glngSys, 1345
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\��Һ��Ƭ", "��Ƭ����", Me.cbo����.Text
    zlDatabase.SetPara "��ӡ��ǩ���Ƿ��ӡ���ܱ���", cboSum.ListIndex, glngSys, 1345
    zlDatabase.SetPara "ɨ����ƿǩ���Զ�����", chksend.Value, glngSys, 1345
    zlDatabase.SetPara "���÷Ѱ�������ȡ", chkMoney.Value, glngSys, 1345
    zlDatabase.SetPara "����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����", chkBag.Value, glngSys, 1345
    zlDatabase.SetPara "�����Σ�ҩƷ����", chkSort.Value, glngSys, 1345
    zlDatabase.SetPara "����ҩƷ��ҩƷ����ָ������", chkMedi.Value, glngSys, 1345
    zlDatabase.SetPara "��˸�ҩ������������", chkCheck.Value, glngSys, 1345
     
    '��������
    zlDatabase.SetPara "������ֹʱ��", Format(dtpBegin.Value, "hh:mm:ss") & "|" & Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, 1345
    zlDatabase.SetPara "�����յ��ռ���ǰҽ��", chk����ҽ��.Value & "|" & Me.txtDeff.Text, glngSys, 1345
    zlDatabase.SetPara "ҽ������", intҽ������, glngSys, 1345
'    zlDatabase.SetPara "ͬ������Һ����", IIf(chk��Һ������.Value = 1, Val(txt��Һ����.Text), ""), glngSys, 1345
'    zldatabase.SetPara "����ҺҩƷ����", IIf(chkҩƷ����.Value = 1, txtҩƷ����.Text, ""), glngSys, 1345
    zlDatabase.SetPara "�����ϴ�����", chkLastBatch.Value, glngSys, 1345
    zlDatabase.SetPara "���ý���ʱ�����", chkOpen.Value, glngSys, 1345
    zlDatabase.SetPara "��������", Me.CboStore.ItemData(Me.CboStore.ListIndex), glngSys, 1345
    zlDatabase.SetPara "ƿǩ��ӡ����", Me.cboNum.Text, glngSys, 1345
    zlDatabase.SetPara "�������Ĳ����յľ���Ӫ��ҽ���ڲ�������", chkTpn.Value, glngSys, 1345
    zlDatabase.SetPara "��������ҩƷ�����͵���������", chkSpecial.Value, glngSys, 1345
    zlDatabase.SetPara "��Ժ���˲������÷�", chkOutPai.Value, glngSys, 1345
    zlDatabase.SetPara "�����Զ�����", chkAutoBatch.Value, glngSys, 1345
    zlDatabase.SetPara "�Ƿ����õĳ���ҩƷ����ҩƷ���˲���", chkByMedi.Value, glngSys, 1345
    zlDatabase.SetPara "�������û�ҩ������Һ��������", chkChangeDrug.Value, glngSys, 1345
    zlDatabase.SetPara "�Զ�����ʱ��Һ��������ֻ���������α䶯", chkAutoMode.Value, glngSys, 1345
    zlDatabase.SetPara "���췢�͵�ҽ����������Һ��ȫ������������", chkBeach.Value, glngSys, 1345
    zlDatabase.SetPara "��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա", chkPeople.Value, glngSys, 1345
    zlDatabase.SetPara "���ҩƷ�ڷ��ͻ�����ȡ���÷�", chkPacket.Value, glngSys, 1345
    


    '��ҩ;��
    zlDatabase.SetPara "��Һ��ҩ;��", IIf(chk��ҩ;��.Value = 1, str��ҩ;��, ""), glngSys, 1345
    
    '��Դ����
    zlDatabase.SetPara "��Դ����", IIf(chk��Դ����.Value = 1, str��Դ����, ""), glngSys, 1345
    
    If IsHavePrivs(mstrPrivs, "���ù�������") Then
        With vsfBatch
            For lngRow = 2 To .rows - 1
                If IsDate(.TextMatrix(lngRow, .ColIndex("����ʱ�俪ʼ"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("����ʱ�����"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("��ҩʱ�俪ʼ"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("��ҩʱ�����"))) Then
                    
                    strInput = IIf(strInput = "", "", strInput & "|") & _
                        Mid(.TextMatrix(lngRow, .ColIndex("����")), 1, Len(.TextMatrix(lngRow, .ColIndex("����"))) - 1) & "," & _
                        .TextMatrix(lngRow, .ColIndex("����ʱ�俪ʼ")) & "-" & .TextMatrix(lngRow, .ColIndex("����ʱ�����")) & "," & _
                        .TextMatrix(lngRow, .ColIndex("��ҩʱ�俪ʼ")) & "-" & .TextMatrix(lngRow, .ColIndex("��ҩʱ�����")) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("���")) = "", 0, 1) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("����")) = "", 0, 1) & "," & _
                        .Cell(flexcpBackColor, lngRow, .ColIndex("��ɫ")) & "," & _
                        .TextMatrix(lngRow, .ColIndex("ҩƷ����"))
                End If
            Next
        End With
        
        '���strInputΪ�ձ�ʾɾ��������������
        gstrSQL = "Zl_��ҩ��������_Save("
        '������Ϣ
        gstrSQL = gstrSQL & "'" & strInput & "',"
        gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ҩ��������")
    End If
     
    If mblnSetPara Then
        '������������
        With Me.vsfVolume
            For lngRow = 0 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("��������")) <> "" And .TextMatrix(lngRow, .ColIndex("����")) <> "" Then
                    
                    gstrSQL = "Zl_������������_Save("
                    '����id
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����id")) & "',"
                    '��������
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("��������")) & "',"
                    '����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("��ҩ����")) & "',"
                    '����
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����"))) & ","
                    '���ȼ�
                    gstrSQL = gstrSQL & lngRow & ","
                    '��������ID
                    gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
                    gstrSQL = gstrSQL & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
                End If
            Next
        End With
        
        '���泣��ҩƷ
        With Me.vsfPrint
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("ҩƷid")) <> "" And .TextMatrix(i, .ColIndex("ҩƷ���������")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_��Һ���ȴ�ӡҩƷ_��ӡ����("
                    'ҩƷid
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ","
                    'ҩƷ����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("ҩƷ���������")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "���泣��ҩƷ")
                End If
            Next
        End With
        
        '���治����ҩƷ
        With Me.vsfNoMedi
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("ҩƷid")) <> "" And .TextMatrix(i, .ColIndex("ҩƷ���������")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_��Һ������ҩƷ_����("
                    'ҩƷid
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ","
                    'ҩƷ����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("ҩƷ���������")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "���治����ҩƷ")
                End If
            Next
        End With
    End If
    
    With Me.VSFPrice
        For i = 1 To .rows - 1
            If (.TextMatrix(i, .ColIndex("���ȼ�")) <> "" And .TextMatrix(i, .ColIndex("�շ���Ŀ")) <> "" And .TextMatrix(i, .ColIndex("��Ŀid")) <> "" And .TextMatrix(i, .ColIndex("��ҩ����")) <> "") Or i = 1 Then
                gstrSQL = "Zl_�����շѷ���_����("
                '���
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("���ȼ�"))) & ","
                '��ҩ����
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("��ҩ����")) & "',"
                '��Ŀid
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("��Ŀid"))) & ","
                '�շ���Ŀ
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("�շ���Ŀ")) & "',"
                '����id
                gstrSQL = gstrSQL & "NULL" & ","
                '�Ƿ��һ������
                gstrSQL = gstrSQL & i & ")"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�������÷�")
            End If
        Next
    End With
    
    n = i - 1
    
    With Me.VSFPrice_��ҩ;��
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("�շ���Ŀ")) <> "" And .TextMatrix(i, .ColIndex("��Ŀid")) <> "" And .TextMatrix(i, .ColIndex("��ҩ;��")) <> "" And .TextMatrix(i, .ColIndex("����id")) <> "" Then
                gstrSQL = "Zl_�����շѷ���_����("
                '���
                gstrSQL = gstrSQL & i + n & ","
                '��ҩ����
                gstrSQL = gstrSQL & "NULL" & ","
                '��Ŀid
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("��Ŀid"))) & ","
                '�շ���Ŀ
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("�շ���Ŀ")) & "',"
                '����id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("����id"))) & ","
                '�Ƿ��һ������
                gstrSQL = gstrSQL & i + n & ")"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�������÷�")
            End If
        Next
    End With
    
    frmPIVAMain.mblnParamsRefresh = True
    
    Unload Me
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPRI(ByVal str����id As String)
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ����id,��������,��ҩ����,Ƶ��,��Ч,���ȼ� from ��ҺҩƷ���ȼ� where (����id=[1] or ����id='0') order by ���ȼ�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ȼ�����", str����id)
    
    i = 1
    rsTemp.Filter = "����id='" & str����id & "'"
    If rsTemp.EOF Then rsTemp.Filter = ""
    With Me.vsfPri
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
            .TextMatrix(1, .ColIndex("���")) = 1
        Else
            .rows = rsTemp.RecordCount + 1
        End If
       
        Do While Not rsTemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("���")) = rsTemp!���ȼ�
            .TextMatrix(i, .ColIndex("��ҩ����")) = rsTemp!��ҩ����
            .TextMatrix(i, .ColIndex("Ƶ��")) = rsTemp!Ƶ��
            .TextMatrix(i, .ColIndex("��Ч")) = rsTemp!��Ч
            i = i + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPrice()
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ���,��ҩ����,��Ŀid,�շ���Ŀ from �����շѷ��� where nvl(����id,0) = 0 order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
            .TextMatrix(1, .ColIndex("���ȼ�")) = 1
        Else
            .rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            If NVL(rsTemp!��Ŀid) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("���ȼ�")) = i
                .TextMatrix(i, .ColIndex("��ҩ����")) = rsTemp!��ҩ����
                .TextMatrix(i, .ColIndex("��Ŀid")) = rsTemp!��Ŀid
                .TextMatrix(i, .ColIndex("�շ���Ŀ")) = rsTemp!�շ���Ŀ
                i = i + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPrice_��ҩ;��()
    Dim rsTemp As Recordset
    Dim rsData As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ����id,��Ŀid,�շ���Ŀ from �����շѷ��� where nvl(����id,0) <> 0 order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice_��ҩ;��
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
        Else
            .rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            '��ѯ������Ŀ����
            gstrSQL = "select ���� from ������ĿĿ¼ where id = [1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ������Ŀ����", rsTemp!����id)
            
            If NVL(rsTemp!��Ŀid) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("����id")) = rsTemp!����id
                .TextMatrix(i, .ColIndex("��ҩ;��")) = rsData!����
                .TextMatrix(i, .ColIndex("��Ŀid")) = rsTemp!��Ŀid
                .TextMatrix(i, .ColIndex("�շ���Ŀ")) = rsTemp!�շ���Ŀ
                i = i + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdVolAdd_Click()
    If Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("��������")) <> "" And Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("����")) <> "" Then
        Me.vsfVolume.rows = Me.vsfVolume.rows + 1
        Me.vsfVolume.RowHeight(Me.vsfVolume.rows - 1) = 250
    End If
End Sub

Private Sub cmdVolDel_Click()
    If Me.vsfVolume.Row = 0 Then Exit Sub
    Me.vsfVolume.RemoveItem Me.vsfVolume.Row
End Sub

Private Sub cmdYes_Click()
    Dim strIDS As String
    Dim strReturn As String
    Dim i As Integer
    
    strReturn = ReturnSelectedPri(1, strIDS)
    
    If mintPri = 1 Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    ElseIf mintPri = 2 Then
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    ElseIf mintPri = 3 Then
        With Me.VSFPrice
            If mintCol = .ColIndex("��ҩ����") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "����ҩ�����Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
        End With
    ElseIf mintPri = 4 Then
        Me.VSFPrice.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice.ColIndex("�շ���Ŀ") Then
            VSFPrice.TextMatrix(mintRow, VSFPrice.ColIndex("��Ŀid")) = strIDS
        End If
    ElseIf mintPri = 5 Then
        With Me.VSFPrice_��ҩ;��
            If mintCol = .ColIndex("��ҩ;��") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "�ø�ҩ;���Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
            .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
        End With
    ElseIf mintPri = 6 Then
        Me.VSFPrice_��ҩ;��.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice_��ҩ;��.ColIndex("�շ���Ŀ") Then
            VSFPrice_��ҩ;��.TextMatrix(mintRow, VSFPrice_��ҩ;��.ColIndex("��Ŀid")) = strIDS
        End If
    End If
    
End Sub

Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    Select Case cboƱ������.ListIndex
    Case 0
        '��Һƿ��ǩ
        strBill = "ZL1_BILL_1345_1"
    Case 1
        '��ҩҩƷ�����嵥
        strBill = "ZL1_INSIDE_1345_1"
    Case 2
        '����ҩƷ�����嵥
        strBill = "ZL1_INSIDE_1345_2"
    Case 3
        '��ҩ�����嵥
        strBill = "ZL1_BILL_1345_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

'Private Sub cmd����Cancel_Click()
'    pic����.Visible = False
'    CmdOK.Enabled = True
'    CmdCancel.Enabled = True
'End Sub
'
'Private Sub cmd����ok_Click()
'    ReturnSelected���� 0
'End Sub
'
'Private Sub cmdҩƷ����_Click()
'    Dim lngRow As Long
'    Dim str���� As String
'
'    On Error Resume Next
'
'    With pic����
'        .Visible = True
'
'        .Height = fraҩƷ����.Top - 50
'        .Top = sstMain.Top + fraҩƷ����.Top + txtҩƷ����.Top - .Height - 50
'        .Left = sstMain.Left + fraҩƷ����.Left
'        .Width = fraҩƷ����.Width - 15
'
'        txtҩƷ����.Text = Trim(txtҩƷ����.Text)
'        If txtҩƷ����.Text <> "" Then
'            With Me.LvwҩƷ����
'                For lngRow = 1 To .ListItems.count
'                    .ListItems(lngRow).Checked = False
'                    str���� = Mid(.ListItems(lngRow).Text, InStr(1, .ListItems(lngRow).Text, "-") + 1)
'                    If InStr(1, "," & txtҩƷ����.Text & ",", "," & str���� & ",") > 0 Then
'                        .ListItems(lngRow).Checked = True
'                    End If
'                Next
'            End With
'        End If
'
'        .SetFocus
'        .ZOrder 0
'
'        CmdOK.Enabled = False
'        CmdCancel.Enabled = False
'    End With
'End Sub

'Private Sub Command1_Click()
'    ReturnSelected���� 0
'End Sub

Private Sub Command2_Click()
    pic����.Visible = False
    CmdOK.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Sub Form_Load()
    mblnSetPara = IsHavePrivs(mstrPrivs, "��������")
    
    With cbo��ǩ��ӡ(0)
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .ListIndex = 0
    End With
    
    With cbo��ǩ��ӡ(1)
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .ListIndex = 0
    End With
    
    With cbo��ҩ��
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .AddItem "2-����ӡ"
    End With
    
    With cbo���͵�
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .AddItem "2-����ӡ"
    End With
    
    With cboSum
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .AddItem "2-����ӡ"
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-��Һƿ��ǩ"
        .AddItem "2-��ҩҩƷ�����嵥"
        .AddItem "3-����ҩƷ�����嵥"
        .AddItem "4-��ҩ�����嵥"

        .ListIndex = 0
    End With
    
    With VSFPrice
        .Left = 0
        .Top = tabPrice.TabHeight
        .Width = tabPrice.Width
        .Height = tabPrice.Height - tabPrice.TabHeight
    End With
    
    With VSFPrice_��ҩ;��
        .Left = 0
        .Top = tabPrice.TabHeight
        .Width = tabPrice.Width
        .Height = tabPrice.Height - tabPrice.TabHeight
    End With
    
'    With cboTPN
'        .Clear
'        .AddItem "0-�������Ĳ����գ�ͨ�����ŷ�ҩҵ����"
'        .AddItem "1-��������ʼ�ս��ղ����"
'        .AddItem "2-��������ʼ�ս��ղ�����"
'
'        .ListIndex = 0
'    End With
    
    With cboNum
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        
        .ListIndex = 0
    End With
        
    With vsfBatch
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(.ColIndex("����")) = True
        .MergeCol(.ColIndex("��ɫ")) = True
        .MergeCol(.ColIndex("����ʱ�俪ʼ")) = True
        .MergeCol(.ColIndex("����ʱ�����")) = True
        .MergeCol(.ColIndex("��ҩʱ�俪ʼ")) = True
        .MergeCol(.ColIndex("��ҩʱ�����")) = True
        .MergeCol(.ColIndex("���")) = True
        .MergeCol(.ColIndex("����")) = True
        .MergeCol(.ColIndex("ҩƷ����")) = True
        .MergeCells = flexMergeFixedOnly
    End With
    
    Call LoadStore
    Call LoadҩƷ����
    
    '��ȡ����
    Call LoadBatchSet
    Call LoadParams
    Call LoadPRI
    Call LoadVsfPrice
    Call LoadVsfPrice_��ҩ;��
    
    Call loadVolume
    Call LoadDept
    
    Call chkAll_Click
    
    Call chkOpen_Click
End Sub
Private Sub LoadBatchSet()
    '��ȡ��ҩ���Ĺ�������
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ����,��ɫ, ��ҩʱ��, ��ҩʱ��, ���, ����,ҩƷ���� From ��ҩ�������� where ��������ID=[1] Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ���Ĺ�������", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    With vsfBatch
        .rows = 2
        .ColComboList(.ColIndex("ҩƷ����")) = "����ҩ|Ӫ��ҩ|������"
        Do While Not rsTmp.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsTmp!���� & "#"
            .TextMatrix(.rows - 1, .ColIndex("����ʱ�俪ʼ")) = Mid(rsTmp!��ҩʱ��, 1, InStr(rsTmp!��ҩʱ��, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("����ʱ�����")) = Mid(rsTmp!��ҩʱ��, InStr(rsTmp!��ҩʱ��, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("��ҩʱ�俪ʼ")) = Mid(rsTmp!��ҩʱ��, 1, InStr(rsTmp!��ҩʱ��, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("��ҩʱ�����")) = Mid(rsTmp!��ҩʱ��, InStr(rsTmp!��ҩʱ��, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("���")) = IIf(rsTmp!��� = 0, "", "��")
            .TextMatrix(.rows - 1, .ColIndex("����")) = IIf(rsTmp!���� = 0, IIf(rsTmp!���� = 0, "��", ""), "��")
            .TextMatrix(.rows - 1, .ColIndex("ҩƷ����")) = NVL(rsTmp!ҩƷ����)
            
            If .TextMatrix(.rows - 1, .ColIndex("����")) = "" Then
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &HE0E0E0
            Else
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &H80000005
            End If
            
            .Cell(flexcpBackColor, .rows - 1, .ColIndex("��ɫ"), .rows - 1, .ColIndex("��ɫ")) = IIf(rsTmp!���� = 0, &H80000005, rsTmp!��ɫ)
            rsTmp.MoveNext
        Loop
        
        vsfBatch.Enabled = IsHavePrivs(mstrPrivs, "���ù�������")
        If vsfBatch.Enabled = False Then
            Label2.Caption = Label2.Caption & "(��Ȩ�޽����޸�)"
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnEdit = False
End Sub

Private Sub lvwPRI_DblClick()
    Dim strIDS As String
    Dim strReturn As String
    Dim i As Integer
    
    strReturn = ReturnSelectedPri(0, strIDS)
    
    If mintPri = 1 Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    ElseIf mintPri = 2 Then
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    ElseIf mintPri = 3 Then
        With Me.VSFPrice
            If mintCol = .ColIndex("��ҩ����") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "����ҩ�����Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
        End With
    ElseIf mintPri = 4 Then
        Me.VSFPrice.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice.ColIndex("�շ���Ŀ") Then
            VSFPrice.TextMatrix(mintRow, VSFPrice.ColIndex("��Ŀid")) = strIDS
        End If
    ElseIf mintPri = 5 Then
        With Me.VSFPrice_��ҩ;��
            If mintCol = .ColIndex("��ҩ;��") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "�ø�ҩ;���Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
            .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
        End With
    ElseIf mintPri = 6 Then
        Me.VSFPrice_��ҩ;��.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice_��ҩ;��.ColIndex("�շ���Ŀ") Then
            VSFPrice_��ҩ;��.TextMatrix(mintRow, VSFPrice_��ҩ;��.ColIndex("��Ŀid")) = strIDS
        End If
    End If
End Sub

Private Sub lvwPRI_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With lvwPRI
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        
        Item.Selected = True
        If Mid(Item.Text, 1, 2) = "����" Then
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

Private Sub lvwPRI_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strIDS As String
    Dim strReturn As String
    Dim i As Integer

    If KeyCode = vbKeyReturn Then
        strReturn = ReturnSelectedPri(1, strIDS)
        
        If mintPri = 1 Then
            With Me.vsfPri
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("��������") Then
                    .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
                End If
            End With
        ElseIf mintPri = 2 Then
            With Me.vsfVolume
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("��������") Then
                    .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
                End If
            End With
        ElseIf mintPri = 3 Then
            With Me.VSFPrice
                If mintCol = .ColIndex("��ҩ����") Then
                    For i = 1 To .rows - 1
                        If strReturn = .TextMatrix(i, mintCol) Then
                            MsgBox "����ҩ�����Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(mintRow, mintCol) = strReturn
            End With
        ElseIf mintPri = 4 Then
            Me.VSFPrice.TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = VSFPrice.ColIndex("�շ���Ŀ") Then
                VSFPrice.TextMatrix(mintRow, VSFPrice.ColIndex("��Ŀid")) = strIDS
            End If
        
        End If
    End If
End Sub

'Private Sub LvwҩƷ����_DblClick()
'    ReturnSelected���� 0
'End Sub
'
'
'Private Sub ReturnSelected����(ByVal intType As Integer)
'    'intType:0-˫�������б�ʱ��1-�����б��а��س�ʱ
'    Dim n As Integer
'
'    With LvwҩƷ����
'        If .SelectedItem Is Nothing Then Exit Sub
'        Me.txtҩƷ����.Text = ""
'
'        '���ѡ����ȫѡ������ȡ���и�ҩ;����
'        If .ListItems(1).Checked Then
'            Me.txtҩƷ����.Text = "����ҩƷ����"
'            pic����.Visible = False
'            Exit Sub
'        End If
'
'        For n = 1 To .ListItems.count
'            If .ListItems(n).Checked Then
'                Me.txtҩƷ����.Text = IIf(Me.txtҩƷ����.Text = "", Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1), Me.txtҩƷ����.Text & "," & Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1))
'            End If
'        Next
'
'        If intType = 0 Then
'            '�����ǰ˫���ĸ�ҩ;��δ��ѡ�ϣ�����ǰ˫���ĸ�ҩ;��Ҳ���뵽�༭����
'            If .SelectedItem.Checked = False Then
'                .SelectedItem.Checked = True
'                Me.txtҩƷ����.Text = IIf(Me.txtҩƷ����.Text = "", Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1), Me.txtҩƷ����.Text & "," & Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1))
'            End If
'
'            If .ListItems(1).Checked Then
'                 Me.txtҩƷ����.Text = "����ҩƷ����"
'                pic����.Visible = False
'                Exit Sub
'            End If
'        End If
'
'        pic����.Visible = False
'
'        CmdOK.Enabled = True
'        CmdCancel.Enabled = True
'    End With
'End Sub

Private Function ReturnSelectedPri(ByVal intType As Integer, ByRef strIDS As String) As String
    'intType:0-˫���б�ʱ��1-�б��а��س�ʱ
    Dim n As Integer
    Dim strReturn As String
    
    With lvwPRI
        If .SelectedItem Is Nothing Then Exit Function
        
        strReturn = .SelectedItem.Text
        strIDS = Mid(.SelectedItem.Key, 2)
        
'        '���ѡ����ȫѡ������ȡ����ѡ����
'        If .ListItems(1).Checked Then
'            strReturn = .ListItems(1).Text
'            ReturnSelectedPri = strReturn
'            picPRI.Visible = False
'            Exit Function
'        End If
'
'        For n = 1 To .ListItems.Count
'            If .ListItems(n).Checked Then
'                strReturn = IIf(strReturn = "", .ListItems(n).Text, strReturn & "," & .ListItems(n).Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'        Next
'
'        If intType = 0 Then
'            '�����ǰ˫����ѡ��δ��ѡ�ϣ�����ǰ˫����ѡ��Ҳ���뵽�༭����
'            If .SelectedItem.Checked = False Then
'                .SelectedItem.Checked = True
'                strReturn = IIf(strReturn = "", .SelectedItem.Text, strReturn & "," & .SelectedItem.Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'
'            If .ListItems(1).Checked Then
'                strReturn = .ListItems(1).Text
'                ReturnSelectedPri = strReturn
'                Exit Function
'            End If
'        End If
        
        picPRI.Visible = False
        
        CmdOK.Enabled = True
        CmdCancel.Enabled = True
        ReturnSelectedPri = strReturn
        mblnEdit = True
    End With
End Function

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

'Private Sub LvwҩƷ����_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        ReturnSelected���� 1
'    End If
'End Sub
'
'
'Private Sub LvwҩƷ����_LostFocus()
''    LvwҩƷ����.Visible = False
'End Sub


Private Sub picPRI_Resize()
    On Error Resume Next
    
    With lvwPRI
        .Top = 0
        .Left = 0
        .Width = picPRI.Width
        .Height = picPRI.Height - 200 - cmdNO.Height
    End With
    
    With cmdNO
        .Top = picPRI.Height - .Height - 50
        .Left = picPRI.Width - .Width - 50
    End With
    
    With cmdYes
        .Top = cmdNO.Top
        .Left = cmdNO.Left - .Width - 100
    End With
End Sub

Private Sub pic����_Resize()
    On Error Resume Next
    
    With LvwҩƷ����
        .Top = 0
        .Left = 0
        .Width = pic����.Width
        .Height = pic����.Height - cmd����Ok.Height - 100
    End With
    
    With cmd����Cancel
        .Top = pic����.Height - .Height - 50
        .Left = pic����.Width - .Width - 50
    End With
    
    With cmd����Ok
        .Top = cmd����Cancel.Top
        .Left = cmd����Cancel.Left - .Width - 100
    End With
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
    Dim i As Integer
    
    If PreviousTab = 2 And pic����.Visible = True Then
'        Call cmd����Cancel_Click
    ElseIf PreviousTab = 5 Then
        Me.vsfVolume.Row = Me.vsfVolume.rows - 1
        Me.vsfVolume.Col = Me.vsfVolume.ColIndex("��������")
    End If
End Sub

Private Sub txt��Һ����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub LoadҩƷ����()
    Dim rsData As ADODB.Recordset

    Set rsData = DeptSendWork_Get����(mlng�ⷿid)
    
    With LvwҩƷ����
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "����ҩƷ����", 1, 1
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, rsData!����, 1, 1
            rsData.MoveNext
        Loop
    End With
End Sub


Private Sub LoadPRI()

    Set mRsDept = DeptSendWork_Get��������
    
    Set mRsType = DeptSendWork_Get��ҩ����
    
    Set mRsPC = DeptSendWork_GetƵ��
    
    Set mRsPrice = DeptSendWork_Get�շ���Ŀ
        
    Set mRsWay = DeptSendWork_��ҩ;��
End Sub


Private Sub updDeff_DownClick()
    If Me.txtDeff.Text <> "0" Then
        Me.txtDeff.Text = Me.txtDeff.Text - 1
    End If
End Sub

Private Sub updDeff_UpClick()
    If Me.txtDeff.Text <> "24" Then
        Me.txtDeff.Text = Me.txtDeff.Text + 1
    End If
End Sub

Private Sub vsfBatch_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBatch
        Select Case Col
            Case .ColIndex("����ʱ�俪ʼ"), .ColIndex("����ʱ�����"), .ColIndex("��ҩʱ�俪ʼ"), .ColIndex("��ҩʱ�����")
                If .TextMatrix(Row, Col) = "" Then Exit Sub
                
                If IsDate(.TextMatrix(Row, Col)) = False Then
                    MsgBox "��¼��ʱ���ʽ������12:59����9:20�ȡ�", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = ""
                    Exit Sub
                End If
                
                If Col = .ColIndex("����ʱ�俪ʼ") And .TextMatrix(Row, .ColIndex("����ʱ�����")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("����ʱ�俪ʼ"))) >= CDate(.TextMatrix(Row, .ColIndex("����ʱ�����"))) Then
                        MsgBox "��ʼʱ�����С�ڽ���ʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("����ʱ�����") And .TextMatrix(Row, .ColIndex("����ʱ�俪ʼ")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("����ʱ�����"))) <= CDate(.TextMatrix(Row, .ColIndex("����ʱ�俪ʼ"))) Then
                        MsgBox "����ʱ�������ڿ�ʼʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("��ҩʱ�俪ʼ") And .TextMatrix(Row, .ColIndex("��ҩʱ�����")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�俪ʼ"))) >= CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�����"))) Then
                        MsgBox "��ʼʱ�����С�ڽ���ʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("��ҩʱ�����") And .TextMatrix(Row, .ColIndex("��ҩʱ�俪ʼ")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�����"))) <= CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�俪ʼ"))) Then
                        MsgBox "����ʱ�������ڿ�ʼʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfBatch_DblClick()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        If (.Col <> .ColIndex("���") And .Col <> .ColIndex("����")) And .Col <> .ColIndex("��ɫ") Then Exit Sub
        If (.MouseRow <> .Row Or .MouseCol <> .Col) And .Col <> .ColIndex("��ɫ") Then Exit Sub
        
        If .Col <> .ColIndex("��ɫ") Then
            If .TextMatrix(.Row, .Col) = "��" Then
                If .TextMatrix(.Row, .ColIndex("����")) = "0#" And .Col = .ColIndex("����") Then
                    MsgBox "0������Ϊ�������Σ��޷�����Ϊ�������á�״̬��"
                Else
                    .TextMatrix(.Row, .Col) = ""
                End If
            Else
                .TextMatrix(.Row, .Col) = "��"
            End If
            
            If .Col = .ColIndex("����") Then
                If .TextMatrix(.Row, .Col) = "" Then
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HE0E0E0
                Else
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &H80000005
                End If
            End If
        
        Else
            On Error GoTo errHandle
            cmdialog.CancelError = True
            cmdialog.ShowColor
            .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = cmdialog.Color
            
errHandle:
        End If
    End With
End Sub


Private Sub vsfBatch_EnterCell()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        .Editable = flexEDNone
        
        If .Col = .ColIndex("����ʱ�俪ʼ") Or .Col = .ColIndex("����ʱ�����") Or .Col = .ColIndex("��ҩʱ�俪ʼ") Or .Col = .ColIndex("��ҩʱ�����") Or .Col = .ColIndex("ҩƷ����") Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfBatch_KeyPress(KeyAscii As Integer)
    With vsfBatch
        If KeyAscii = 13 Then
            If .Col < .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row < .rows - 1 Then
                    .Row = .Row + 1
                    .Col = .ColIndex("����ʱ�俪ʼ")
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfBatch_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfBatch
        Select Case Col
            Case .ColIndex("����ʱ�俪ʼ"), .ColIndex("����ʱ�����"), .ColIndex("��ҩʱ�俪ʼ"), .ColIndex("��ҩʱ�����")
                If InStr("1234567890:" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(":") Then
                    If InStr(.EditText, ":") <> 0 Then
                        KeyAscii = 0
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row <> 1 Then Cancel = True
End Sub

Private Sub vsfDept_EnterCell()
    If mblnEdit Then
        If MsgBox("�뱣�����õ����ȼ����л����Һ����������ȼ����ý�ʧЧ���Ƿ��л���", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
            mblnEdit = False
            
        End If
    Else
        If Me.vsfDept.Row > 1 Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
        End If
    End If
    
End Sub

Private Sub vsfDept_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    With Me.vsfDept
        If KeyAscii <> 13 Or .TextMatrix(1, .ColIndex("��������")) = "" Or .Row <> 1 Then Exit Sub
        
        For intRow = 2 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("����")) = UCase(.TextMatrix(1, .ColIndex("��������"))) Then
                .Row = intRow
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub vsfNoMedi_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strkey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = GetControlRect(vsfNoMedi.hWnd)
        dblLeft = vRect.Left + vsfNoMedi.CellLeft
        dblTop = vRect.Top + vsfNoMedi.CellTop + vsfNoMedi.CellHeight + 3200
        
        With vsfNoMedi
            If Col = .ColIndex("ҩƷ���������") Then
                strkey = Trim(.EditText)
                If strkey = "" Then Exit Sub
                
                If IsNumeric(strkey) Then
                    '������
                    StrCode = " d.���� like [1] "
                ElseIf zlCommFun.IsCharAlpha(strkey) Then
                    '����ĸ
                    StrCode = " n.���� Like [1] "
                ElseIf zlCommFun.IsCharChinese(strkey) Then
                    '������
                    StrCode = " d.���� like [1] "
                Else
                    StrCode = " (n.���� Like [1] Or d.���� Like [1] Or n.���� Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'��' || d.���� || '��' || d.���� || '(' || d.��� || ')' As ͨ����" & vbNewLine & _
                    " From ҩƷ��� T, �շ���ĿĿ¼ D, �շ���Ŀ���� N" & vbNewLine & _
                    " Where t.ҩƷid = d.Id And t.ҩƷid = n.�շ�ϸĿid And D.��� In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.����ʱ�� Is Null Or To_Char(d.����ʱ��, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '��' || d.���� || '��' || d.���� || '(' || d.��� || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ���������", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("ҩƷID"))) Then
                            MsgBox rsRecord!ͨ���� & "�Ѿ�¼�룬������ѡ��", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("ҩƷ���������")) = rsRecord!ͨ����
                    .EditText = rsRecord!ͨ����
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfNoMedi_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfPRI_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    mintPri = 1
    mintRow = Row
    mintCol = Col
    With Me.picPRI
        .Visible = True
    
        .Height = vsfPri.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfPri.Left
        .Width = vsfPri.Width
    End With
            
    Select Case Col
        Case vsfPri.ColIndex("��������")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "���п���", 1, 1
                mRsDept.MoveFirst
                Do While Not mRsDept.EOF
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
                    mRsDept.MoveNext
                Loop
                .ListItems.Add , "_00", "��������", 1, 1
            End With
        Case vsfPri.ColIndex("��ҩ����")
            With Me.lvwPRI
                .ListItems.Clear
                mRsType.MoveFirst
                Do While Not mRsType.EOF
                    .ListItems.Add , "_" & mRsType!����, mRsType!����, 1, 1
                    mRsType.MoveNext
                Loop
                 .ListItems.Add , "_00", "��������", 1, 1
            End With
        Case vsfPri.ColIndex("Ƶ��")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "����Ƶ��", 1, 1
                mRsPC.MoveFirst
                Do While Not mRsPC.EOF
                    .ListItems.Add , "_" & mRsPC!����, mRsPC!���� & "(" & mRsPC!Ӣ������ & ")", 1, 1
                    mRsPC.MoveNext
                Loop
                .ListItems.Add , "_00", "����Ƶ��", 1, 1
            End With
    End Select
End Sub

Private Sub VSFPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_��ҩ;��_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.picPRI
        .Visible = True
    
        .Height = VSFPrice.Height
        .Top = sstMain.Top + tabPrice.Top + VSFPrice.Top
        .Left = sstMain.Left + tabPrice.Left + VSFPrice.Left
        .Width = VSFPrice.Width
    End With
    
    mintRow = Row
    mintCol = Col
    
    If Col = VSFPrice.ColIndex("��ҩ����") Then
        mintPri = 3
        With Me.lvwPRI
            .ListItems.Clear
            mRsType.MoveFirst
            Do While Not mRsType.EOF
                .ListItems.Add , "_" & mRsType!����, mRsType!����, 1, 1
                mRsType.MoveNext
            Loop
             .ListItems.Add , "_00", "��������", 1, 1
        End With
    ElseIf Col = VSFPrice.ColIndex("�շ���Ŀ") Then
        mintPri = 4
        With Me.lvwPRI
            .ListItems.Clear
            mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!Id, mRsPrice!����, 1, 1
                mRsPrice.MoveNext
            Loop
        End With
    End If
    
End Sub

Private Sub VSFPrice_��ҩ;��_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.picPRI
        .Visible = True
    
        .Height = VSFPrice_��ҩ;��.Height
        .Top = sstMain.Top + tabPrice.Top + VSFPrice_��ҩ;��.Top
        .Left = sstMain.Left + tabPrice.Left + VSFPrice_��ҩ;��.Left
        .Width = VSFPrice_��ҩ;��.Width
    End With
    
    mintRow = Row
    mintCol = Col
    
    If Col = VSFPrice_��ҩ;��.ColIndex("��ҩ;��") Then
        mintPri = 5
        With Me.lvwPRI
            .ListItems.Clear
            If mRsWay.RecordCount > 0 Then mRsWay.MoveFirst
            Do While Not mRsWay.EOF
                .ListItems.Add , "_" & mRsWay!Id, mRsWay!����, 1, 1
                mRsWay.MoveNext
            Loop
        End With
    ElseIf Col = VSFPrice_��ҩ;��.ColIndex("�շ���Ŀ") Then
        mintPri = 6
        With Me.lvwPRI
            .ListItems.Clear
            If mRsPrice.RecordCount > 0 Then mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!Id, mRsPrice!����, 1, 1
                mRsPrice.MoveNext
            Loop
        End With
    End If
    
End Sub

Private Sub VSFPrice_EnterCell()
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    If Me.VSFPrice.Row < 2 Then
        cmdLast.Enabled = False
    ElseIf Me.VSFPrice.Row = Me.VSFPrice.rows - 1 Then
        cmdNext.Enabled = False
    End If
    
    VSFPrice.Editable = flexEDNone
    
    If VSFPrice.ColSel <> VSFPrice.ColIndex("���ȼ�") Then
        VSFPrice.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub VSFPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim i As Integer
    
    If VSFPrice.Row = 0 Then Exit Sub
    If KeyCode = 13 And VSFPrice.Row = VSFPrice.rows - 1 Then
        With Me.VSFPrice
            If .TextMatrix(.Row, .ColIndex("��ҩ����")) <> "" And .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) <> "" Then
                .rows = .rows + 1
                .Row = .rows - 1
                .Col = .ColIndex("��ҩ����")
                .TextMatrix(.Row, .ColIndex("���ȼ�")) = .Row
            End If
        End With
    ElseIf KeyCode = 46 Then
        intRow = VSFPrice.Row
        If VSFPrice.rows = 2 Then
           VSFPrice.rows = 1
           VSFPrice.rows = 2
        Else
            Me.VSFPrice.RemoveItem VSFPrice.Row
        End If
        
        '�������
        For i = intRow To Me.VSFPrice.rows - 1
            Me.VSFPrice.TextMatrix(i, Me.VSFPrice.ColIndex("���ȼ�")) = i
        Next
    End If
    
End Sub

Private Sub VSFPrice_��ҩ;��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim i As Integer
    
    If VSFPrice_��ҩ;��.Row = 0 Then Exit Sub
    If KeyCode = 13 And VSFPrice_��ҩ;��.Row = VSFPrice_��ҩ;��.rows - 1 Then
        Me.VSFPrice_��ҩ;��.Editable = flexEDNone
        With Me.VSFPrice_��ҩ;��
            If .TextMatrix(.Row, .ColIndex("��ҩ;��")) <> "" And .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) <> "" Then
                .rows = .rows + 1
                .Row = .rows - 1
                .Col = .ColIndex("��ҩ;��")
            End If
        End With
    ElseIf KeyCode = 46 Then
        intRow = VSFPrice_��ҩ;��.Row
        If VSFPrice_��ҩ;��.rows = 2 Then
           VSFPrice_��ҩ;��.rows = 1
           VSFPrice_��ҩ;��.rows = 2
        Else
            Me.VSFPrice_��ҩ;��.RemoveItem VSFPrice_��ҩ;��.Row
        End If
    End If
    Me.VSFPrice_��ҩ;��.Editable = flexEDKbd
    
End Sub


Private Sub vsfPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If Me.vsfPrint.rows = 2 Then
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("ҩƷid")) = ""
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("ҩƷ���������")) = ""
        Else
            Me.vsfPrint.RemoveItem vsfPrint.Row
        End If
        
    End If
End Sub

Private Sub vsfPrint_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strkey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = GetControlRect(vsfPrint.hWnd)
        dblLeft = vRect.Left + vsfPrint.CellLeft
        dblTop = vRect.Top + vsfPrint.CellTop + vsfPrint.CellHeight + 3200
        
        With vsfPrint
            If Col = .ColIndex("ҩƷ���������") Then
                strkey = Trim(.EditText)
                If strkey = "" Then Exit Sub
                
                If IsNumeric(strkey) Then
                    '������
                    StrCode = " d.���� like [1] "
                ElseIf zlCommFun.IsCharAlpha(strkey) Then
                    '����ĸ
                    StrCode = " n.���� Like [1] "
                ElseIf zlCommFun.IsCharChinese(strkey) Then
                    '������
                    StrCode = " d.���� like [1] "
                Else
                    StrCode = " (n.���� Like [1] Or d.���� Like [1] Or n.���� Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'��' || d.���� || '��' || d.���� || '(' || d.��� || ')' As ͨ����" & vbNewLine & _
                    " From ҩƷ��� T, �շ���ĿĿ¼ D, �շ���Ŀ���� N" & vbNewLine & _
                    " Where t.ҩƷid = d.Id And t.ҩƷid = n.�շ�ϸĿid And D.��� In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.����ʱ�� Is Null Or To_Char(d.����ʱ��, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '��' || d.���� || '��' || d.���� || '(' || d.��� || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ���������", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("ҩƷID"))) Then
                            MsgBox rsRecord!ͨ���� & "�Ѿ�¼�룬������ѡ��", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("ҩƷ���������")) = rsRecord!ͨ����
                    .EditText = rsRecord!ͨ����
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPrint_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub vsfVolume_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = Me.vsfVolume.ColIndex("����") Then
        If Not IsNumeric(vsfVolume.TextMatrix(Row, Col)) Then
            MsgBox "������¼�����֣�", vbInformation + vbOKOnly, gstrSysName
            vsfVolume.Col = vsfVolume.ColIndex("����")
        End If
    End If
End Sub

Private Sub vsfVolume_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str���� As String
    Dim i As Integer
    
    If Col <> vsfVolume.ColIndex("��ҩ����") Then Exit Sub
    With Me.vsfBatch
        If .rows > 2 Then
            For i = 2 To .rows - 1
                If .TextMatrix(i, .ColIndex("����")) <> "" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                    str���� = IIf(str���� = "", "", str���� & "|") & .TextMatrix(i, .ColIndex("����"))
                End If
            Next
        End If
        If str���� <> "" Then Me.vsfVolume.ColComboList(vsfVolume.ColIndex("��ҩ����")) = str����
    End With
End Sub

Private Sub vsfVolume_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    mblnPri = False
'    mintRow = Row
'
'    mintCol = Col
''    With Me.picPRI
'        .Visible = True
'        .Height = vsfVolume.Height
'        .Top = sstMain.Top + vsfPri.Top
'        .Left = sstMain.Left + vsfVolume.Left
'        .Width = vsfVolume.Width
'    End With
'
'    With vsfVolume
'        If Col = .ColIndex("��������") Then
'            With Me.lvwPRI
'                .ListItems.Clear
'                .ListItems.Add , "_" & 0, "���п���", 1, 1
'                mRsDept.MoveFirst
'                Do While Not mRsDept.EOF
'                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
'                    mRsDept.MoveNext
'                Loop
'                .ListItems.Add , "_00", "��������", 1, 1
'            End With
'        End If
'    End With

    mintPri = 2
    mintRow = vsfVolume.Row
    mintCol = vsfVolume.Col

    With Me.lvwPRI
        .ListItems.Clear
        .ListItems.Add , "_" & 0, "���п���", 1, 1
        mRsDept.MoveFirst
        Do While Not mRsDept.EOF
            If vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) <> "" Then
                If mRsDept!���� = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or mRsDept!��ʼ��� = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!���� Then
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!����
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.ColIndex("����id")) = mRsDept!Id
                    Exit Sub

                ElseIf InStr(1, mRsDept!��ʼ���, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!����, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!����, vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) > 0 Then
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
                End If
            Else
                .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
            End If
            mRsDept.MoveNext
        Loop
        
        If .ListItems.count = 1 Then
            .ListItems.Clear
            MsgBox "������ļ���û����֮ƥ��Ŀ��ң�������¼�룡"
            Exit Sub
        End If
        
        .ListItems.Add , "_00", "��������", 1, 1
    End With
    

    With Me.picPRI
        .Visible = True
        .Height = vsfVolume.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfVolume.Left
        .Width = vsfVolume.Width
    End With
End Sub

Private Sub loadVolume()
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ����id,��������,����,��ҩ���� from ������������ where ��������ID=[1] order by ����id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������������", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    i = 1
    With Me.vsfVolume
        .RowHeight(0) = 250
        .rows = 1
        .rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        Do While Not rsTemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("����id")) = rsTemp!����ID
            .TextMatrix(i, .ColIndex("��������")) = rsTemp!��������
            .TextMatrix(i, .ColIndex("��ҩ����")) = NVL(rsTemp!��ҩ����)
            .TextMatrix(i, .ColIndex("����")) = rsTemp!����
            i = i + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfVolume_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 13 Then Exit Sub
'
'    With Me.vsfVolume
'        If .Row = .rows - 1 Then
'            If .Col = .Cols - 1 Then
'                Exit Sub
'            Else
'                .Col = .Col + 1
'            End If
'        Else
'            If .Col = .Cols - 1 Then
'                .Row = .Row + 1
'                .Col = .ColIndex("��������")
'            Else
'                .Col = .Col + 1
'            End If
'        End If
'    End With
End Sub

Private Sub vsfVolume_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    With Me.vsfVolume
        If .Row = .rows - 1 Then
            If .Col = .Cols - 1 Then
                Exit Sub
            Else
                .Col = .Col + 1
            End If
        Else
            If .Col = .Cols - 1 Then
                .Row = .Row + 1
                .Col = .ColIndex("��������")
            Else
                .Col = .Col + 1
            End If
        End If
    End With
End Sub

Private Sub vsfVolume_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfVolume
        If Col = .ColIndex("����") Then
            If InStr("1234567890-." & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End With

End Sub

Private Sub LoadDept()
    Dim i As Integer
    
    i = 1
    vsfDept.rows = mRsDept.RecordCount + 2
    Do While Not mRsDept.EOF
        With Me.vsfDept
            i = i + 1
            .TextMatrix(i, .ColIndex("���")) = i - 1
            .TextMatrix(i, .ColIndex("����id")) = mRsDept!Id
            .TextMatrix(i, .ColIndex("��������")) = mRsDept!����
            .TextMatrix(i, .ColIndex("����")) = mRsDept!����
        End With
        mRsDept.MoveNext
    Loop
End Sub

Private Sub vsfPrint_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfPrint.ColIndex("ҩƷ���������") Then Cancel = True
End Sub

Private Sub vsfNoMedi_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfNoMedi.ColIndex("ҩƷ���������") Then Cancel = True
End Sub

Private Sub vsfNoMedi_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = 46 Then
        If Me.vsfNoMedi.rows = 2 Then
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("ҩƷid")) = ""
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("ҩƷ���������")) = ""
        Else
            Me.vsfNoMedi.RemoveItem vsfNoMedi.Row
        End If
    End If
End Sub
