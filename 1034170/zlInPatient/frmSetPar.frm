VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetPar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frmSetPar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7170
      TabIndex        =   2
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7200
      TabIndex        =   0
      Top             =   450
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin TabDlg.SSTab sTab 
      Height          =   5550
      Left            =   135
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   9790
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "�������(&1)"
      TabPicture(0)   =   "frmSetPar.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFee"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDeviceSetup"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "optDept(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "optDept(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk����"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chk����"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkSeekName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtNameDays"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkLedWelcome"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDiagDays"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chk����"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboԤ������"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboFee"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Ԥ��Ʊ�ݿ���(&2)"
      TabPicture(1)   =   "frmSetPar.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrepay"
      Tab(1).Control(1)=   "fraWristlet"
      Tab(1).Control(2)=   "fraPatientPage"
      Tab(1).Control(3)=   "fraDeposit"
      Tab(1).Control(4)=   "img16"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "ҽ�ƿ�Ʊ�ݿ���(&3)"
      TabPicture(2)   =   "frmSetPar.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkɨ�����֤ǩԼ"
      Tab(2).Control(1)=   "fraTitle"
      Tab(2).Control(2)=   "cboType"
      Tab(2).Control(3)=   "lblȱʡ����"
      Tab(2).ControlCount=   4
      Begin VB.ComboBox cboFee 
         Height          =   300
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   4695
         Width           =   2580
      End
      Begin VB.ComboBox cboԤ������ 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   5070
         Width           =   2580
      End
      Begin VB.CheckBox chkɨ�����֤ǩԼ 
         Caption         =   "ɨ�����֤ǩԼ"
         Height          =   180
         Left            =   -74715
         TabIndex        =   66
         Top             =   4335
         Value           =   1  'Checked
         Width           =   2520
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "��Ժʱ�Զ�����һ�η���"
         Height          =   180
         Left            =   2985
         MaskColor       =   &H00000000&
         TabIndex        =   63
         Top             =   4455
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2040
         TabIndex        =   59
         Top             =   4350
         Width           =   285
      End
      Begin VB.TextBox txtDiagDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   180
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   57
         Text            =   "3"
         Top             =   4154
         Width           =   285
      End
      Begin VB.Frame Frame1 
         Caption         =   "�����꾭����Ŀ"
         Height          =   3030
         Left            =   150
         TabIndex        =   34
         Top             =   420
         Width           =   6405
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�����֤��"
            Height          =   195
            Index           =   26
            Left            =   3450
            TabIndex        =   55
            Top             =   930
            Value           =   1  'Checked
            Width           =   1800
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ������"
            Height          =   195
            Index           =   20
            Left            =   3450
            TabIndex        =   62
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ�ʱ�"
            Height          =   195
            Index           =   19
            Left            =   3450
            TabIndex        =   60
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ�绰"
            Height          =   195
            Index           =   18
            Left            =   3450
            TabIndex        =   58
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "������λ"
            Height          =   195
            Index           =   17
            Left            =   3450
            TabIndex        =   56
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�˵绰"
            Height          =   195
            Index           =   16
            Left            =   3450
            TabIndex        =   54
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�˵�ַ"
            Height          =   195
            Index           =   15
            Left            =   3450
            TabIndex        =   53
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ�˹�ϵ"
            Height          =   195
            Index           =   14
            Left            =   1785
            TabIndex        =   49
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ϵ������"
            Height          =   195
            Index           =   13
            Left            =   1785
            TabIndex        =   48
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ͥ�绰"
            Height          =   195
            Index           =   12
            Left            =   1785
            TabIndex        =   47
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��ͥ��ַ�ʱ�"
            Height          =   195
            Index           =   11
            Left            =   1785
            TabIndex        =   46
            Top             =   930
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��סַ"
            Height          =   195
            Index           =   10
            Left            =   1785
            TabIndex        =   45
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "�����ص�"
            Height          =   195
            Index           =   9
            Left            =   1785
            TabIndex        =   44
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���֤��"
            Height          =   195
            Index           =   8
            Left            =   285
            TabIndex        =   43
            Top             =   2745
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��������"
            Height          =   195
            Index           =   7
            Left            =   285
            TabIndex        =   41
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���"
            Height          =   195
            Index           =   6
            Left            =   285
            TabIndex        =   40
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "ְҵ"
            Height          =   195
            Index           =   5
            Left            =   285
            TabIndex        =   39
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����״��"
            Height          =   195
            Index           =   4
            Left            =   285
            TabIndex        =   38
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "ѧ��"
            Height          =   195
            Index           =   3
            Left            =   285
            TabIndex        =   37
            Top             =   930
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   36
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   35
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "��λ�ʺ�"
            Height          =   195
            Index           =   0
            Left            =   3450
            TabIndex        =   64
            Top             =   2445
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����֤��"
            Height          =   195
            Index           =   21
            Left            =   285
            TabIndex        =   42
            Top             =   2445
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���ڵ�ַ"
            Height          =   195
            Index           =   22
            Left            =   1785
            TabIndex        =   50
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "���ڵ�ַ�ʱ�"
            Height          =   195
            Index           =   23
            Left            =   1785
            TabIndex        =   51
            Top             =   2445
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   24
            Left            =   1785
            TabIndex        =   52
            Top             =   2745
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "����"
            Height          =   195
            Index           =   25
            Left            =   3450
            TabIndex        =   65
            Top             =   2745
            Value           =   1  'Checked
            Width           =   1200
         End
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Left            =   3000
         TabIndex        =   33
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   3600
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "���ع���Ԥ��Ʊ��"
         Height          =   2535
         Left            =   -74895
         TabIndex        =   31
         Top             =   420
         Width           =   6510
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   2145
            Left            =   60
            TabIndex        =   32
            Top             =   270
            Width           =   6285
            _cx             =   11086
            _cy             =   3784
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
            FormatString    =   $"frmSetPar.frx":0060
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
      Begin VB.Frame fraTitle 
         Caption         =   "���ع���ҽ�ƿ�"
         Height          =   3570
         Left            =   -74865
         TabIndex        =   28
         Top             =   555
         Width           =   6390
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   3150
            Left            =   60
            TabIndex        =   29
            Top             =   300
            Width           =   6150
            _cx             =   10848
            _cy             =   5556
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
            FormatString    =   $"frmSetPar.frx":013F
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
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73575
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4635
         Width           =   2580
      End
      Begin VB.Frame fraWristlet 
         Caption         =   "�������"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74865
         TabIndex        =   22
         Top             =   4365
         Width           =   6465
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2685
            TabIndex        =   26
            Top             =   285
            Width           =   1500
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   25
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   24
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   2
            Left            =   5370
            TabIndex        =   23
            Top             =   160
            Width           =   990
         End
      End
      Begin VB.Frame fraPatientPage 
         Caption         =   "������ҳ"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74865
         TabIndex        =   17
         Top             =   3690
         Width           =   6465
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   20
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2655
            TabIndex        =   19
            Top             =   285
            Width           =   1380
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   1
            Left            =   5370
            TabIndex        =   18
            Top             =   160
            Width           =   990
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "Ԥ����Ʊ��"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74865
         TabIndex        =   12
         Top             =   3000
         Width           =   6480
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   16
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   15
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   14
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   0
            Left            =   5355
            TabIndex        =   13
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.TextBox txtNameDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   180
         Left            =   3090
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "0��ʾ����ʱ������ʱ��"
         Top             =   3884
         Width           =   285
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   3040
         TabIndex        =   5
         Top             =   4080
         Width           =   285
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
         Height          =   195
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   3877
         Width           =   4620
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���벡�˵�����Ϣ"
         Height          =   195
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   3600
         Width           =   1740
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "ҽ�ƿ������Լ��˷�ʽ��ȡ"
         Height          =   180
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   4455
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.OptionButton optDept 
         Caption         =   "��ѡ����"
         Height          =   255
         Index           =   0
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   4725
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDept 
         Caption         =   "��ѡ����"
         Height          =   255
         Index           =   1
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   4740
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   4830
         TabIndex        =   6
         Top             =   5070
         Width           =   1500
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   -70560
         Top             =   1320
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
               Picture         =   "frmSetPar.frx":0221
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡ�ѱ�"
         Height          =   180
         Left            =   2985
         TabIndex        =   70
         Top             =   4770
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "ȱʡ�ɿʽ"
         Height          =   225
         Left            =   390
         TabIndex        =   68
         Top             =   5145
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "ԤԼ����ʱ��ȡ����    ���ڵ������Ϣ"
         Height          =   180
         Left            =   405
         TabIndex        =   61
         Top             =   4154
         Width           =   3855
      End
      Begin VB.Label lblȱʡ���� 
         Caption         =   "ȱʡ��������"
         Height          =   225
         Left            =   -74730
         TabIndex        =   30
         Top             =   4695
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmSetPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrPrivs As String
Public mlngModul As Long

 

Private Sub cboType_Click()
    chkɨ�����֤ǩԼ.Enabled = Not (cboType.Text = "�������֤")
    If cboType.Text = "�������֤" Then
        chkɨ�����֤ǩԼ.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1131)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    '��꾭����Ŀ
    For i = 0 To chkItem.UBound
        zlDatabase.SetPara chkItem(i).Caption, chkItem(i).Value, glngSys, mlngModul, IIf(chkItem(i).Enabled = True, True, False)
    Next
    
    Call SaveInvoice
    
    zlDatabase.SetPara "������Ϣ", chk����.Value, glngSys, mlngModul, IIf(chk����.Enabled = True, True, False)
    zlDatabase.SetPara "����ģ������", chkSeekName.Value, glngSys, mlngModul, IIf(chkSeekName.Enabled = True, True, False)
    zlDatabase.SetPara "������������", Val(txtNameDays.Text), glngSys, mlngModul, IIf(txtNameDays.Enabled = True, True, False)
    zlDatabase.SetPara "���Ѽ���", chk����.Value, glngSys, mlngModul, IIf(chk����.Enabled = True, True, False)
    zlDatabase.SetPara "��ѡ����", IIf(optDept(1).Value, 1, 0), glngSys, mlngModul, IIf(optDept(1).Enabled = True, True, False)
    zlDatabase.SetPara "��ϲ�������", Val(txtDiagDays.Text), glngSys, mlngModul, True
    '36454,������,2012-09-06
    zlDatabase.SetPara "���ü���ʱ��", chk����.Value, glngSys, mlngModul, IIf(chk����.Enabled = True, True, False)

    'LED�豸
    zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chkLedWelcome.Value, glngSys, mlngModul, IIf(chkLedWelcome.Enabled = True, True, False)
    'Ԥ����Ʊ�ݴ�ӡ
    For i = 0 To optPrepayPrint.UBound
        If optPrepayPrint(i).Value Then
            zlDatabase.SetPara "Ԥ����Ʊ�ݴ�ӡ", i, glngSys, mlngModul, IIf(optPrepayPrint(i).Enabled = True, True, False)
        End If
    Next
    
    '������ҳ��ӡ��ʽ
    For i = 0 To optFpagePrint.UBound
        If optFpagePrint(i).Value Then
            zlDatabase.SetPara "������ҳ��ӡ", i, glngSys, mlngModul, IIf(optFpagePrint(i).Enabled = True, True, False)
        End If
    Next
    
    '���������ӡ��ʽ
    For i = 0 To optWristletPrint.UBound
        If optWristletPrint(i).Value Then
            zlDatabase.SetPara "���������ӡ", i, glngSys, mlngModul, IIf(optWristletPrint(i).Enabled = True, True, False)
        End If
    Next
    '�����:53408
    zlDatabase.SetPara "ɨ�����֤ǩԼ", IIf(chkɨ�����֤ǩԼ.Value = 1, 1, 0), glngSys, glngModul, InStr(mstrPrivs, "��������") > 0
    
    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset, objItem As ListItem
    Dim blnBill As Boolean
    
    gblnOK = False
    On Error GoTo errH
    
    '��꾭����Ŀ
    For i = 0 To chkItem.UBound
        chkItem(i).Value = zlDatabase.GetPara(chkItem(i).Caption, glngSys, mlngModul, 1, Array(chkItem(i)), InStr(mstrPrivs, "��������") > 0)
    Next
    Call InitShareInvoice
    
    chk����.Value = IIf(zlDatabase.GetPara("������Ϣ", glngSys, mlngModul, , Array(chk����), InStr(mstrPrivs, "������Ϣ") > 0) = "1", 1, 0)
    chkSeekName.Value = IIf(zlDatabase.GetPara("����ģ������", glngSys, mlngModul, , Array(chkSeekName), InStr(mstrPrivs, "��������") > 0) = "1", 1, 0)
    txtNameDays.Text = Val(zlDatabase.GetPara("������������", glngSys, mlngModul, , Array(txtNameDays), InStr(mstrPrivs, "��������") > 0))
    txtDiagDays.Text = Val(zlDatabase.GetPara("��ϲ�������", glngSys, mlngModul, "3", Array(txtDiagDays, Label1), InStr(mstrPrivs, "��������") > 0))
     '�����:53408
    chkɨ�����֤ǩԼ.Value = IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, glngModul, , Array(chkɨ�����֤ǩԼ), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
    
    'LED�豸
    chkLedWelcome.Value = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, 1, Array(chkLedWelcome), InStr(mstrPrivs, "��������") > 0)
        
    i = Val(zlDatabase.GetPara("��ѡ����", glngSys, mlngModul, , Array(optDept(0), optDept(1)), InStr(mstrPrivs, "��������") > 0))
    optDept(1).Value = (i = 1)
    optDept(0).Value = Not optDept(1).Value
    
    chk����.Value = IIf(zlDatabase.GetPara("���Ѽ���", glngSys, mlngModul, , Array(chk����), InStr(mstrPrivs, "��������") > 0) = "1", 1, 0)
    
    i = Val(zlDatabase.GetPara("Ԥ����Ʊ�ݴ�ӡ", glngSys, mlngModul, , Array(fraDeposit), InStr(mstrPrivs, "��������") > 0))
    If i <= optPrepayPrint.UBound Then optPrepayPrint(i).Value = True
    
    i = Val(zlDatabase.GetPara("������ҳ��ӡ", glngSys, mlngModul, , Array(fraPatientPage), InStr(mstrPrivs, "��������") > 0))
    If i <= optFpagePrint.UBound Then optFpagePrint(i).Value = True
    
    i = Val(zlDatabase.GetPara("���������ӡ", glngSys, mlngModul, , Array(fraWristlet), InStr(mstrPrivs, "��������") > 0))
    If i <= optWristletPrint.UBound Then optWristletPrint(i).Value = True
    
    '36454,������,2012-09-06
    chk����.Value = Val(zlDatabase.GetPara("���ü���ʱ��", glngSys, mlngModul, "1", Array(chk����), InStr(mstrPrivs, "��������") > 0))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
  

Private Sub sTab_Click(PreviousTab As Integer)
    If sTab.Tab = 0 Then
        chkItem(1).SetFocus
    ElseIf sTab.Tab = 1 Then
        If vsPrepay.Enabled And vsPrepay.Visible Then vsPrepay.SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
     End If
End Sub

Private Sub txtDiagDays_GotFocus()
    Call SelAll(txtDiagDays)
End Sub

Private Sub txtDiagDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDiagDays_Validate(Cancel As Boolean)
    If Val(txtDiagDays.Text) <= 0 Then
        txtDiagDays.Text = 0
    ElseIf Val(txtDiagDays.Text) > 999 Then
        txtDiagDays.Text = 999
    End If
End Sub

Private Sub txtNameDays_GotFocus()
    Call SelAll(txtNameDays)
End Sub

Private Sub txtNameDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNameDays_Validate(Cancel As Boolean)
    If Val(txtNameDays.Text) <= 0 Then
        txtNameDays.Text = 0
    ElseIf Val(txtNameDays.Text) > 999 Then
        txtNameDays.Text = 999
    End If
End Sub
Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1
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
    Dim strȱʡ�ѱ� As String
    
    On Error GoTo ErrHand
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '�ָ��п��
    lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, , Array(cboType), blnHavePrivs, intType))
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
            .MoveNext
        Loop
    End With
    '�����:58776
    For i = 0 To cboType.ListCount - 1
        If Val(cboType.ItemData(i)) = lngCardTypeID Then
             cboType.ListIndex = i
        End If
    Next
    
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
        rsTemp.Filter = " ʹ�����<>1   "   '������Ԥ������Ʊ��
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            If Val(Nvl(rsTemp!ʹ�����, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "�����סԺ����"
            ElseIf Val(Nvl(rsTemp!ʹ�����, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ������Ʊ��"
            Else
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ��סԺƱ��"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = Val(Nvl(rsTemp!ʹ�����))
            
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
    '����ȱʡ�ɿʽ(Ԥ����)
    Load�ɿʽ
    '���طѱ�
    strSQL = "Select A.����,A.����,A.����,Nvl(A.ȱʡ��־,0) as ȱʡ From �ѱ� A,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) B " & _
             " Where (A.������� = B.Column_Value or A.������� is null) And A.����=1 And Nvl(A.���޳���,0)=0 And  " & _
             "        Sysdate Between NVL(A.��Ч��ʼ,Sysdate-1) and NVL(A.��Ч����,Sysdate+1) Order by A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "1,2,3")
    cboFee.Clear
    Do While Not rsTemp.EOF
        cboFee.AddItem rsTemp!����
        If rsTemp!ȱʡ = 1 Then cboFee.ListIndex = cboFee.NewIndex
    rsTemp.MoveNext
    Loop
    If cboFee.ListCount > 0 And cboFee.ListIndex < 0 Then cboFee.ListIndex = 0
    strȱʡ�ѱ� = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, mlngModul, , blnHavePrivs)
    If strȱʡ�ѱ� <> "" Then
        For i = 0 To cboFee.ListCount - 1
            If cboFee.List(i) = strȱʡ�ѱ� Then
                cboFee.ListIndex = i
            End If
        Next
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:���˺�
    '����:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    Dim lng�����ID As Long
    If cboType.ListIndex >= 0 Then
        lng�����ID = cboType.ItemData(cboType.ListIndex)
    End If
    zlDatabase.SetPara "ȱʡҽ�ƿ����", lng�����ID, glngSys, mlngModul, blnHavePrivs
        
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
    
    zlDatabase.SetPara "ȱʡ�ɿʽ", Trim(cboԤ������.Text), glngSys, mlngModul, blnHavePrivs
    '69489
    zlDatabase.SetPara "ȱʡ�ѱ�", Trim(cboFee.Text), glngSys, mlngModul, blnHavePrivs
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

Public Sub Load�ɿʽ()
    Dim strTemp As String, strȱʡԤ���ʽ As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim objSquareCard As Object
    Dim varData As Variant, varTemp As Variant
    Dim strPayType As String
    Dim j As Long, i As Long
    Dim blnFind As Boolean, blnHavePrivs As Boolean
    
    strTemp = "1,2,5,7,8" & IIf(InStr(mstrPrivs, ";���ղ��˵Ǽ�;") > 0, ",3", "")

    
    strSQL = _
        "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó��� ='Ԥ����'  And B.����=A.���㷽ʽ  " & _
        "           And Nvl(B.����,1) In(" & strTemp & ")" & _
        " Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    strPayType = objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
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
            
            If Not blnFind Then
                .AddItem Nvl(rsTemp!����)
                If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                j = j + 1
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
        strȱʡԤ���ʽ = zlDatabase.GetPara("ȱʡ�ɿʽ", glngSys, mlngModul, , blnHavePrivs)
        If strȱʡԤ���ʽ <> "" Then
            For i = 0 To cboԤ������.ListCount
                If cboԤ������.List(i) = strȱʡԤ���ʽ Then
                    cboԤ������.ListIndex = i
                End If
            Next
        End If
    End With
End Sub
