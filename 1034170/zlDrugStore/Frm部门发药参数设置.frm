VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Frm���ŷ�ҩ�������� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   5775
   ClientLeft      =   8805
   ClientTop       =   3960
   ClientWidth     =   6735
   Icon            =   "Frm���ŷ�ҩ��������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6735
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   0
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   1
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1100
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   5010
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "����(&1)"
      TabPicture(0)   =   "Frm���ŷ�ҩ��������.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl��ѯ����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblNote(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl��ҩҩ��"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblNote(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl����ģʽ"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Cbo������"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraǩ��"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt��ѯ����"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Cbo��ҩҩ��"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Cbo����ģʽ"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fra��ҩ����"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chk�Զ�ˢ��"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt�Զ�ˢ��ʱ��"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chk���ܷ�ҩ"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Chk�����һ�����ʾ"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chk��Ժ"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkReview"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chk�Ƿ�������ʾܾ�"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chk��ҩ״̬"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chk����������ʱ���ܽ�����ҩ����"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chk���ط�ҩʱ�����ҩ����"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "����(&2)"
      TabPicture(1)   =   "Frm���ŷ�ҩ��������.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboName"
      Tab(1).Control(1)=   "frm��ΣҩƷ����"
      Tab(1).Control(2)=   "cbo��ҩ�嵥"
      Tab(1).Control(3)=   "cbo��ҩ�嵥"
      Tab(1).Control(4)=   "fra�豸����"
      Tab(1).Control(5)=   "chkҩƷ����"
      Tab(1).Control(6)=   "Frame3"
      Tab(1).Control(7)=   "Chk�Ƿ��Զ�ȱҩ���"
      Tab(1).Control(8)=   "lblName"
      Tab(1).Control(9)=   "lbl��ҩ�嵥"
      Tab(1).Control(10)=   "lbl��ҩ�嵥"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "����(&3)"
      TabPicture(2)   =   "Frm���ŷ�ҩ��������.frx":1D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "��ҩ����(&4)"
      TabPicture(3)   =   "Frm���ŷ�ҩ��������.frx":1D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lvw��Դ����"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "��ҩ��(&5)"
      TabPicture(4)   =   "Frm���ŷ�ҩ��������.frx":1D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(1)=   "Frame5"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "��ӡ����(&6)"
      TabPicture(5)   =   "Frm���ŷ�ҩ��������.frx":1D86
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmd��ӡ����"
      Tab(5).Control(1)=   "cboƱ������"
      Tab(5).Control(2)=   "lblƱ��"
      Tab(5).ControlCount=   3
      Begin VB.CheckBox chk���ط�ҩʱ�����ҩ���� 
         Caption         =   "���ط�ҩʱ�����ҩ����������ʱ���浥�������Ӷ����ӣ�"
         Height          =   180
         Left            =   180
         TabIndex        =   75
         Top             =   3300
         Width           =   5385
      End
      Begin VB.CheckBox chk����������ʱ���ܽ�����ҩ���� 
         Caption         =   "����������ʱ���ܽ�����ҩ����"
         Height          =   180
         Left            =   3000
         TabIndex        =   74
         Top             =   3060
         Width           =   3105
      End
      Begin VB.CheckBox chk��ҩ״̬ 
         Caption         =   "��ҩ��������Ĭ��Ϊ��ҩ״̬"
         Height          =   180
         Left            =   180
         TabIndex        =   73
         Top             =   3060
         Width           =   4095
      End
      Begin VB.CheckBox chk�Ƿ�������ʾܾ� 
         Caption         =   "�Ƿ�������ʾܾ�"
         Height          =   180
         Left            =   4320
         TabIndex        =   72
         Top             =   2565
         Width           =   1785
      End
      Begin VB.CheckBox chkReview 
         Caption         =   "��ҩʱ���ҽ��"
         Height          =   180
         Left            =   3960
         TabIndex        =   71
         Top             =   2820
         Width           =   1785
      End
      Begin VB.CommandButton cmd��ӡ���� 
         Caption         =   "��ӡ����(&P)"
         Height          =   345
         Left            =   -74760
         TabIndex        =   69
         Top             =   1050
         Width           =   3315
      End
      Begin VB.ComboBox cboƱ������ 
         Height          =   300
         Left            =   -74010
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   600
         Width           =   2565
      End
      Begin VB.Frame Frame4 
         Caption         =   " ���Ϳ���  "
         Height          =   615
         Left            =   -74880
         TabIndex        =   66
         Top             =   360
         Width           =   4935
         Begin VB.CheckBox chkStopTrans 
            Caption         =   "��ͣ��ҩƷ��װ�����ͷ�ҩ����"
            Height          =   255
            Left            =   360
            TabIndex        =   67
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �����������ݿ���  "
         Height          =   3855
         Left            =   -74880
         TabIndex        =   60
         Top             =   1080
         Width           =   4935
         Begin VB.Frame Frame6 
            Caption         =   " ��������  "
            Height          =   615
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   4695
            Begin VB.CheckBox chkType 
               Caption         =   "����"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   65
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox chkType 
               Caption         =   "����"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   64
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   " ����ѡ��"
            Height          =   2775
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   4695
            Begin MSComctlLib.ListView LvwҩƷ���� 
               Height          =   2385
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   4425
               _ExtentX        =   7805
               _ExtentY        =   4207
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
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "����"
                  Object.Width           =   3528
               EndProperty
            End
         End
      End
      Begin VB.ComboBox cboName 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1560
         Width           =   2640
      End
      Begin VB.Frame frm��ΣҩƷ���� 
         Caption         =   "  ѡ���ΣҩƷ�������ŵ����"
         Height          =   580
         Left            =   -74880
         TabIndex        =   53
         Top             =   1960
         Width           =   6135
         Begin VB.CheckBox chk��Σ 
            Caption         =   "C��"
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   56
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chk��Σ 
            Caption         =   "B��"
            Height          =   375
            Index           =   1
            Left            =   1140
            TabIndex        =   55
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chk��Σ 
            Caption         =   "A��"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   54
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.ComboBox cbo��ҩ�嵥 
         Height          =   300
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1185
         Width           =   2655
      End
      Begin VB.ComboBox cbo��ҩ�嵥 
         Height          =   300
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   800
         Width           =   2655
      End
      Begin VB.CheckBox chk��Ժ 
         Caption         =   "��ҩ����ʱ������˳�Ժ���˵���������"
         Height          =   180
         Left            =   180
         TabIndex        =   46
         Top             =   2820
         Width           =   4095
      End
      Begin VB.Frame fra�豸���� 
         Caption         =   "  ���ܿ��������豸���� "
         Height          =   1095
         Left            =   -71280
         TabIndex        =   44
         Top             =   750
         Width           =   2415
         Begin VB.CommandButton cmdDeviceSetup 
            Caption         =   "�豸����(&S)"
            Height          =   350
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ѯ��ϸ��¼����������ʱ����"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   40
         Top             =   1920
         Width           =   4335
         Begin VB.TextBox txtMaxRecordCount 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            TabIndex        =   41
            Text            =   "3000"
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   2160
            TabIndex        =   43
            Top             =   480
            Width           =   180
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ��ϸ��¼"
            Height          =   180
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "���ò�ѯ��ҩ����ҩ����ʱ��ʱ�䷶Χ������ʱ����"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   33
         Top             =   480
         Width           =   4335
         Begin VB.TextBox txtTimeArea_Sended 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   37
            Text            =   "3"
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox txtTimeArea_Send 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   34
            Text            =   "7"
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   1920
            TabIndex        =   39
            Top             =   900
            Width           =   180
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ��ҩ����"
            Height          =   180
            Left            =   240
            TabIndex        =   38
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   1920
            TabIndex        =   36
            Top             =   420
            Width           =   180
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ��ҩ����"
            Height          =   180
            Left            =   240
            TabIndex        =   35
            Top             =   420
            Width           =   1080
         End
      End
      Begin VB.CheckBox Chk�����һ�����ʾ 
         Caption         =   "�����һ�����ʾ"
         Height          =   180
         Left            =   2640
         TabIndex        =   32
         Top             =   2565
         Width           =   1785
      End
      Begin VB.CheckBox chk���ܷ�ҩ 
         Caption         =   "��ҩʱ������ҩ���ʼ�¼"
         Height          =   180
         Left            =   180
         TabIndex        =   31
         Top             =   2565
         Width           =   3615
      End
      Begin VB.CheckBox chkҩƷ���� 
         Caption         =   "��ʾ�ⷿ��λ�����������ʾ"
         Height          =   180
         Left            =   -71880
         TabIndex        =   30
         Top             =   480
         Width           =   2745
      End
      Begin VB.TextBox txt�Զ�ˢ��ʱ�� 
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "5"
         Top             =   1710
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CheckBox chk�Զ�ˢ�� 
         Caption         =   "�Զ�ˢ��δ��ҩ�嵥"
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   1740
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "ѡ���ڷ�ҩʱ�Զ����Ϊ�������ҩƷ����"
         Height          =   2340
         Left            =   -74880
         TabIndex        =   22
         Top             =   2640
         Width           =   6135
         Begin MSComctlLib.ListView lvw��ֵ���� 
            Height          =   1755
            Left            =   2160
            TabIndex        =   26
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvw������� 
            Height          =   1750
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvw��Σ���� 
            Height          =   1755
            Left            =   4200
            TabIndex        =   51
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "��ΣҩƷ�ȼ�����"
            Height          =   180
            Left            =   4200
            TabIndex        =   52
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ��ֵ����"
            Height          =   180
            Left            =   2160
            TabIndex        =   25
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ�������"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame fra��ҩ���� 
         Caption         =   "��ҩ����"
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   3540
         Width           =   4485
         Begin VB.OptionButton opt��ҩ���� 
            Caption         =   "ȫ��ʵ��"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   21
            ToolTipText     =   "ҩ������ҩ����ȫ��ʵ��ҩƷ"
            Top             =   270
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opt��ҩ���� 
            Caption         =   "��ʵ��"
            Height          =   180
            Index           =   1
            Left            =   1605
            TabIndex        =   20
            ToolTipText     =   "ҩ��������ҩƷȫ��תΪ����"
            Top             =   270
            Width           =   960
         End
         Begin VB.OptionButton opt��ҩ���� 
            Caption         =   "�����������"
            Height          =   180
            Index           =   2
            Left            =   2790
            TabIndex        =   19
            ToolTipText     =   "ҩ���������������װҩƷ����������װҩƷת������"
            Top             =   285
            Width           =   1470
         End
      End
      Begin VB.ComboBox Cbo����ģʽ 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Cbo��ҩҩ�� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   690
         Width           =   1815
      End
      Begin VB.CheckBox Chk�Ƿ��Զ�ȱҩ��� 
         Caption         =   "�Ƿ��Զ�ȱҩ���"
         Height          =   180
         Left            =   -74880
         TabIndex        =   9
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txt��ѯ���� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1005
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "1"
         Top             =   1725
         Width           =   405
      End
      Begin VB.Frame fraǩ�� 
         Caption         =   "��ҩ��/��ҩ���Ƿ���Ҫǩ��"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   4260
         Width           =   4485
         Begin VB.CheckBox chk��ҩ��ǩ�� 
            Caption         =   "��ҩ��ǩ��"
            Height          =   255
            Left            =   150
            TabIndex        =   7
            Top             =   285
            Width           =   1485
         End
         Begin VB.CheckBox chk��ҩ��ǩ�� 
            Caption         =   "��ҩ��ǩ��"
            Height          =   255
            Left            =   2550
            TabIndex        =   6
            Top             =   285
            Width           =   1485
         End
      End
      Begin VB.ComboBox Cbo������ 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2130
         Width           =   1815
      End
      Begin MSComctlLib.ListView Lvw��Դ���� 
         Height          =   4605
         Left            =   -74880
         TabIndex        =   57
         Top             =   360
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   8123
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
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblƱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74730
         TabIndex        =   70
         Top             =   660
         Width           =   630
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "ҩ����ʾ"
         Height          =   180
         Left            =   -74880
         TabIndex        =   59
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl��ҩ�嵥 
         Caption         =   "��ҩ�嵥"
         Height          =   195
         Left            =   -74880
         TabIndex        =   50
         Top             =   1238
         Width           =   735
      End
      Begin VB.Label lbl��ҩ�嵥 
         Caption         =   "��ҩ�嵥"
         Height          =   195
         Left            =   -74880
         TabIndex        =   48
         Top             =   853
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "����"
         Height          =   255
         Left            =   4245
         TabIndex        =   29
         Top             =   1740
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Lbl����ģʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������ҩ��"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Lbl��ҩҩ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩҩ��"
         Height          =   180
         Left            =   180
         TabIndex        =   15
         Top             =   750
         Width           =   720
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ҫ�������Ǵ����������ʱ�������߼���"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   1080
         Width           =   4710
      End
      Begin VB.Label lbl��ѯ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ����"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�� �� ��"
         Height          =   180
         Left            =   180
         TabIndex        =   12
         Top             =   2205
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   5160
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
            Picture         =   "Frm���ŷ�ҩ��������.frx":1DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm���ŷ�ҩ��������.frx":20BC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm���ŷ�ҩ��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPrivs As String
Private mblnSetPara As Boolean                          '�Ƿ���в�������Ȩ��
Private BlnStart As Boolean
Private intDays As Integer
Private lngҩ��ID As Long
Private Lng����ģʽ As Long
Private Lng������ʾ As Long
Private Lng�Զ���ӡ As Long
Private Lngȱҩ��� As Long
Private Lng��ҩ��ǩ�� As Long
Private Lng��ҩ��ǩ�� As Long
Private str������� As String
Private str��ֵ���� As String
Private RecDrugStore As New ADODB.Recordset             'ҩ��
Private mstrSourceDep As String                         '��Դ����
Private mLng��ӡ��ҩ�嵥 As Long                        '��ҩ�嵥
Public blnStartPacker As Boolean                       '�Ƿ�����ҩƷ�ְ����ӿ�
Private Sub Get������()
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    '���ü�����
    On Error GoTo errHandle
    strSQL = "Select Distinct A.����" & _
             " From ��Ա�� A,������Ա B,��������˵�� C,��Ա����˵�� D " & _
             " Where A.Id=B.��Աid And B.����id=C.����Id And D.��Աid=A.Id And D.��Ա���� = 'ҩ����ҩ��' " & _
             " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) "
        
    If Cbo��ҩҩ��.ListIndex <> -1 Then
        strSQL = strSQL & " AND B.����id=[1] "
    End If
    
    strSQL = strSQL & " ORDER BY A.���� "

    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.ListIndex))
    
    Cbo������.Clear
    Cbo������.AddItem "���м�����"
    Do While Not rs.EOF
        Cbo������.AddItem rs!����
        rs.MoveNext
    Loop
    
    rs.Close
    
    Cbo������.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Bill_GotFocus()
    Me.KeyPreview = False
End Sub


Private Sub Bill_LostFocus()
    Me.KeyPreview = True
End Sub

Private Sub Cbo��ҩҩ��_Click()
    Call Get������
End Sub

Private Sub chk���ܷ�ҩ_Click()
    If chk���ܷ�ҩ.Value = 1 Then
        Chk�����һ�����ʾ.Value = 1
        Chk�����һ�����ʾ.Enabled = False
    Else
        Chk�����һ�����ʾ.Enabled = True
    End If
End Sub

Private Sub chk�Զ�ˢ��_Click()
    If chk�Զ�ˢ��.Value = 1 Then
        If mblnSetPara = True Then
            txt�Զ�ˢ��ʱ��.Enabled = True
        Else
            txt�Զ�ˢ��ʱ��.Enabled = False
        End If
    Else
        txt�Զ�ˢ��ʱ��.Enabled = False
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1342)
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim n As Integer
    Dim int��ҩ���� As Integer
    Dim str���� As String
    Dim str���� As String
    Dim i As Integer
    Dim str��Σ���� As String
    Dim str��Σ���� As String
    
    If Trim(txt��ѯ����.Text) = "" Then
        txt��ѯ����.Text = "1"
'        MsgBox "�������ѯ������1��-30�죩��", vbInformation, gstrSysName
'        tabShow.Tab = 0
'        txt��ѯ����.SetFocus
'        Exit Sub
    End If
    If Not IsNumeric(txt��ѯ����.Text) Then
        MsgBox "��ѯ�����к��зǷ��ַ���", vbInformation, gstrSysName
        tabShow.Tab = 0
        If txt��ѯ����.Enabled = True Then txt��ѯ����.SetFocus
        Exit Sub
    End If
    If Val(txt��ѯ����.Text) < 1 Or Val(txt��ѯ����.Text) > 30 Then
        MsgBox "��ѯ��������С��1������30�죡", vbInformation, gstrSysName
        tabShow.Tab = 0
        If txt��ѯ����.Enabled = True Then txt��ѯ����.SetFocus
        Exit Sub
    End If
    
    For n = 0 To opt��ҩ����.count - 1
        If opt��ҩ����(n).Value = True Then
            int��ҩ���� = n
            Exit For
        End If
    Next
    
    str������� = ""
    For n = 1 To lvw�������.ListItems.count
        If lvw�������.ListItems(n).Checked = True Then
            str������� = IIf(str������� = "", lvw�������.ListItems(n).Text, str������� & "," & lvw�������.ListItems(n).Text)
        End If
    Next
    
    str��ֵ���� = ""
    For n = 1 To lvw��ֵ����.ListItems.count
        If lvw��ֵ����.ListItems(n).Checked = True Then
            str��ֵ���� = IIf(str��ֵ���� = "", lvw��ֵ����.ListItems(n).Text, str��ֵ���� & "," & lvw��ֵ����.ListItems(n).Text)
        End If
    Next
    
    For n = 1 To lvw��Σ����.ListItems.count
        If lvw��Σ����.ListItems(n).Checked = True Then
            str��Σ���� = IIf(str��Σ���� = "", n, str��Σ���� & "," & n)
        End If
    Next
    
    If chk��Σ(0).Value = 1 Then str��Σ���� = IIf(str��Σ���� = "", 1, str��Σ���� & "," & 1)
    If chk��Σ(1).Value = 1 Then str��Σ���� = IIf(str��Σ���� = "", 2, str��Σ���� & "," & 2)
    If chk��Σ(2).Value = 1 Then str��Σ���� = IIf(str��Σ���� = "", 3, str��Σ���� & "," & 3)
    
    '���湫����˽�в���
    zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, 1342
    zlDatabase.SetPara "��ҩ����", int��ҩ����, glngSys, 1342
    zlDatabase.SetPara "��ҩ��ǩ��", chk��ҩ��ǩ��.Value, glngSys, 1342
    zlDatabase.SetPara "ȱҩ���", Chk�Ƿ��Զ�ȱҩ���.Value, glngSys, 1342
    zlDatabase.SetPara "��ҩ��ǩ��", chk��ҩ��ǩ��.Value, glngSys, 1342
    zlDatabase.SetPara "�ⷿ��λ�����������ʾ", chkҩƷ����.Value, glngSys, 1342
    zlDatabase.SetPara "��ҩʱ������ҩ���ʼ�¼", chk���ܷ�ҩ.Value, glngSys, 1342
    
    zlDatabase.SetPara "��˳�Ժ���˵���������", chk��Ժ.Value, glngSys, 1342
    
    If chk�Զ�ˢ��.Value = 1 And Val(txt�Զ�ˢ��ʱ��.Text) > 0 Then
        zlDatabase.SetPara "�Զ�ˢ��δ��ҩ�嵥", Val(txt�Զ�ˢ��ʱ��.Text), glngSys, 1342
    Else
        zlDatabase.SetPara "�Զ�ˢ��δ��ҩ�嵥", 0, glngSys, 1342
    End If

    zlDatabase.SetPara "�����һ�����ʾ�����嵥", Chk�����һ�����ʾ.Value, glngSys, 1342
    zlDatabase.SetPara "����ģʽ", Cbo����ģʽ.ListIndex, glngSys, 1342
    zlDatabase.SetPara "������", Cbo������.Text, glngSys, 1342
    zlDatabase.SetPara "�������", str�������, glngSys, 1342
    zlDatabase.SetPara "��ֵ����", str��ֵ����, glngSys, 1342
    zlDatabase.SetPara "��Σ����", str��Σ����, glngSys, 1342
    zlDatabase.SetPara "��ΣҩƷ����", str��Σ����, glngSys, 1342
    zlDatabase.SetPara "��ҩҩ��", Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.ListIndex), glngSys, 1342
    zlDatabase.SetPara "�Զ���ӡ", Me.cbo��ҩ�嵥.ListIndex, glngSys, 1342
    zlDatabase.SetPara "��ѯ��ҩ����", Val(txtTimeArea_Send.Text), glngSys, 1342
    zlDatabase.SetPara "��ѯ��ҩ����", Val(txtTimeArea_Sended.Text), glngSys, 1342
    zlDatabase.SetPara "��ѯ��ϸ��¼��", Val(txtMaxRecordCount.Text), glngSys, 1342
    zlDatabase.SetPara "��ӡ��ҩ�嵥", Me.cbo��ҩ�嵥.ListIndex, glngSys, 1342
    zlDatabase.SetPara "��ҩʱ���ҽ��", chkReview.Value, glngSys, 1342
    zlDatabase.SetPara "�Ƿ�������ʾܾ�", chk�Ƿ�������ʾܾ�.Value, glngSys, 1342
    zlDatabase.SetPara "��ҩ��������Ĭ��Ϊ��ҩ״̬", chk��ҩ״̬.Value, glngSys, 1342
    zlDatabase.SetPara "����������ʱ���ܽ�����ҩ����", chk����������ʱ���ܽ�����ҩ����.Value, glngSys, 1342
    zlDatabase.SetPara "���ط�ҩʱ�����ҩ����", Me.chk���ط�ҩʱ�����ҩ����.Value, glngSys, 1342
    
    '��Դ����
    mstrSourceDep = ""
    With Me.Lvw��Դ����
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                If mstrSourceDep = "" Then
                    mstrSourceDep = Mid(.ListItems(i).Key, 2)
                Else
                    mstrSourceDep = mstrSourceDep & "," & Mid(.ListItems(i).Key, 2)
                End If
            End If
        Next
    End With
    zlDatabase.SetPara "��Դ����", mstrSourceDep, glngSys, 1342
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ���ŷ�ҩ����", "ҩƷ������ʾ��ʽ", Me.cboName.ListIndex)
    
    '�����װ������
    If blnStartPacker = True Then
        SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��ͣ����", chkStopTrans.Value
        
        str���� = ""
        str���� = str���� & chkType(0).Value
        str���� = str���� & chkType(1).Value

        SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��������", str����
        
        
        If LvwҩƷ����.ListItems(1).Checked Then
             str���� = "����"
        Else
            For n = 1 To LvwҩƷ����.ListItems.count
                If LvwҩƷ����.ListItems(n).Checked Then
                    str���� = IIf(str���� = "", "", str���� & ",") & LvwҩƷ����.ListItems(n).Text
                End If
            Next
        End If
        
        SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "ѡ�����", str����
    End If
    
'    Frm���ŷ�ҩ����.BlnSetPara = True
    frm���ŷ�ҩ����New.BlnSetPara = True
    Unload Me
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Activate()
    If BlnStart = False Then
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim intTrans As Integer
    Dim str���� As String
    Dim str���� As String
    Dim n As Integer
    
    BlnStart = False
    On Error GoTo errHandle
    If IsHavePrivs(strPrivs, "�޸���������") = False Then
        fra��ҩ����.Enabled = False
        opt��ҩ����(0).Enabled = False
        opt��ҩ����(1).Enabled = False
        opt��ҩ����(2).Enabled = False
    End If
    
    If IsHavePrivs(strPrivs, "����ҩ��") Then
        strSQL = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��' And ������� IN (2,3))"
    Else
        strSQL = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��' And B.������� IN (2,3))"
    End If
    gstrSQL = " Select ID,����||'-'||���� ҩ�� From ���ű� Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID In " & strSQL & _
             " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
             " Order by ����||'-'||����"
    Set RecDrugStore = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With RecDrugStore
        If .EOF Then
            MsgBox "���ʼ��ҩ���������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Cbo��ҩҩ��.Clear
        Do While Not .EOF
            Cbo��ҩҩ��.AddItem !ҩ��
            Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.NewIndex) = !Id
            .MoveNext
        Loop
        Cbo��ҩҩ��.ListIndex = 0
    End With
    
    With Cbo����ģʽ
        .Clear
        .AddItem "0-�������е���"
        .AddItem "1-���������ʵ�"
        .AddItem "2-���������ʱ�"
        .ListIndex = 0
    End With
        
    With cbo��ҩ�嵥
        .Clear
        .AddItem "0_��ҩ�󲻴�ӡ"
        .AddItem "1-��ҩ���Զ���ӡ"
        .AddItem "2_��ҩ����ʾ�Ƿ��ӡ"
        .ListIndex = 0
    End With
    
    With cbo��ҩ�嵥
        .Clear
        .AddItem "0_��ҩ�󲻴�ӡ"
        .AddItem "1-��ҩ���Զ���ӡ"
        .AddItem "2_��ҩ����ʾ�Ƿ��ӡ"
        .ListIndex = 0
    End With
    
    With Me.cboName
        .Clear
        .AddItem "0-��ʾҩƷ����������"
        .AddItem "1-����ʾҩƷ����"
        .AddItem "2-����ʾҩƷ����"
        .ListIndex = 0
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-���ܷ�ҩ�嵥"
        .AddItem "2-��ҩ�嵥"
        .ListIndex = 0
    End With
    
    Call Get������
    
    '�������
    gstrSQL = "Select ���� From ҩƷ������� Order By ���� "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption & "-ȡ�������")
    
    With rsTmp
        Do While Not .EOF
            lvw�������.ListItems.Add , "_" & lvw�������.ListItems.count + 1, !����
            .MoveNext
        Loop
    End With
    
    '��ֵ����
    gstrSQL = "Select ���� From ҩƷ��ֵ���� Order By ���� "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption & "-ȡ��ֵ����")
    
    With rsTmp
        Do While Not .EOF
            lvw��ֵ����.ListItems.Add , "_" & lvw��ֵ����.ListItems.count + 1, !����
            .MoveNext
        Loop
    End With
    
    '��ΣҩƷ����
    With lvw��Σ����
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "A��"
        .ListItems.Add , "_" & .ListItems.count + 1, "B��"
        .ListItems.Add , "_" & .ListItems.count + 1, "C��"
    End With
    
    '�ָ�����
    WriteCons

    '��Դ����
    Call SetSourceDep
    
    '��װ���ӿ��������
    Call LoadҩƷ����(Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.ListIndex))
    
    tabShow.TabVisible(4) = blnStartPacker
    
    If blnStartPacker = True Then
        intTrans = Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��ͣ����", "0"))
        chkStopTrans.Value = IIf(intTrans = 1, 1, 0)
        
        str���� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "��������", "11")
        chkType(0).Value = Val(Mid(str����, 1, 1))
        chkType(1).Value = Val(Mid(str����, 2, 1))
        
        str���� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ŷ�ҩ����\��װ������", "ѡ�����", "����")
        
        For n = 1 To LvwҩƷ����.ListItems.count
            LvwҩƷ����.ListItems(n).Checked = False
            If str���� = "����" Then
                LvwҩƷ����.ListItems(n).Checked = True
            Else
                If InStr(1, "," & str���� & ",", "," & LvwҩƷ����.ListItems(n).Text & ",") > 0 Then
                    LvwҩƷ����.ListItems(n).Checked = True
                End If
            End If
        Next
    End If
    
    BlnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadҩƷ����(ByVal lngҩ��ID As Long)
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get����(lngҩ��ID)
    
    With LvwҩƷ����
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "����ҩƷ����", 1, 1
        .ListItems(.ListItems.count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, Mid(rsData!����, InStr(1, rsData!����, "-") + 1), 1, 1
            .ListItems(.ListItems.count).Checked = True
            rsData.MoveNext
        Loop
    End With
End Sub


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
Private Function WriteCons()
    Dim IntLocate As Integer
    Dim str������ As String
    Dim n As Integer
    Dim i As Integer
    Dim int��ҩ���� As Integer
    Dim int�Զ�ˢ�� As Integer
    Dim strArr
    Dim int��ѯ��ҩ���� As Integer
    Dim int��ѯ��ҩ���� As Integer
    Dim lng����¼�� As Long
    Dim int��˳�Ժ�������� As Integer
    Dim str��Σ���� As String
    Dim str��Σ���� As String
    Dim int���ط�ҩʱ�����ҩ���� As Integer
    
    mblnSetPara = IsHavePrivs(strPrivs, "��������")
    
    'ȡ������˽�в���
    intDays = Val(zlDatabase.GetPara("��ѯ����", glngSys, 1342, 1, Array(lbl��ѯ����, txt��ѯ����), mblnSetPara))
    int��ҩ���� = Val(zlDatabase.GetPara("��ҩ����", glngSys, 1342, 0, Array(fra��ҩ����, opt��ҩ����(0), opt��ҩ����(1), opt��ҩ����(2)), mblnSetPara))
    Lng��ҩ��ǩ�� = Val(zlDatabase.GetPara("��ҩ��ǩ��", glngSys, 1342, 0, Array(chk��ҩ��ǩ��), mblnSetPara))
    Lngȱҩ��� = Val(zlDatabase.GetPara("ȱҩ���", glngSys, 1342, 1, Array(Chk�Ƿ��Զ�ȱҩ���), mblnSetPara))
    Lng��ҩ��ǩ�� = Val(zlDatabase.GetPara("��ҩ��ǩ��", glngSys, 1342, 0, Array(chk��ҩ��ǩ��), mblnSetPara))
    int�Զ�ˢ�� = Val(zlDatabase.GetPara("�Զ�ˢ��δ��ҩ�嵥", glngSys, 1342, 0, Array(chk�Զ�ˢ��, txt�Զ�ˢ��ʱ��, Label4), mblnSetPara))
    chkҩƷ����.Value = Val(zlDatabase.GetPara("�ⷿ��λ�����������ʾ", glngSys, 1342, 0, Array(chkҩƷ����), mblnSetPara))
    chk���ܷ�ҩ.Value = Val(zlDatabase.GetPara("��ҩʱ������ҩ���ʼ�¼", glngSys, 1342, 0, Array(chk���ܷ�ҩ), mblnSetPara))

    Lng����ģʽ = Val(zlDatabase.GetPara("����ģʽ", glngSys, 1342, 0, Array(Cbo����ģʽ), mblnSetPara))
    Lng������ʾ = Val(zlDatabase.GetPara("�����һ�����ʾ�����嵥", glngSys, 1342, 0, Array(Chk�����һ�����ʾ), mblnSetPara))
    str������ = zlDatabase.GetPara("������", glngSys, 1342, "���м�����", Array(Label1, Cbo������), mblnSetPara)
    str������� = zlDatabase.GetPara("�������", glngSys, 1342, "", Array(Label2, lvw�������), mblnSetPara)
    str��ֵ���� = zlDatabase.GetPara("��ֵ����", glngSys, 1342, "", Array(Label3, lvw��ֵ����), mblnSetPara)
    str��Σ���� = zlDatabase.GetPara("��Σ����", glngSys, 1342, "", Array(Label11, lvw��Σ����), mblnSetPara)
    str��Σ���� = zlDatabase.GetPara("��ΣҩƷ����", glngSys, 1342, "", Array(frm��ΣҩƷ����), mblnSetPara)
    lngҩ��ID = Val(zlDatabase.GetPara("��ҩҩ��", glngSys, 1342, 0, Array(Lbl��ҩҩ��, Cbo��ҩҩ��), mblnSetPara))
    Lng�Զ���ӡ = Val(zlDatabase.GetPara("�Զ���ӡ", glngSys, 1342, 0, Array(Me.lbl��ҩ�嵥, Me.cbo��ҩ�嵥), mblnSetPara))
    int��ѯ��ҩ���� = Val(zlDatabase.GetPara("��ѯ��ҩ����", glngSys, 1342, 7, Array(txtTimeArea_Send), mblnSetPara))
    int��ѯ��ҩ���� = Val(zlDatabase.GetPara("��ѯ��ҩ����", glngSys, 1342, 3, Array(txtTimeArea_Sended), mblnSetPara))
    lng����¼�� = Val(zlDatabase.GetPara("��ѯ��ϸ��¼��", glngSys, 1342, 3000, Array(txtMaxRecordCount), mblnSetPara))
    int��˳�Ժ�������� = Val(zlDatabase.GetPara("��˳�Ժ���˵���������", glngSys, 1342, 0, Array(chk��Ժ), mblnSetPara))
    mstrSourceDep = zlDatabase.GetPara("��Դ����", glngSys, 1342, "", Array(Lvw��Դ����), mblnSetPara)
    mLng��ӡ��ҩ�嵥 = Val(zlDatabase.GetPara("��ӡ��ҩ�嵥", glngSys, 1342, 0, Array(lbl��ҩ�嵥, Me.cbo��ҩ�嵥), mblnSetPara))
    chkReview.Value = Val(zlDatabase.GetPara("��ҩʱ���ҽ��", glngSys, 1342, 0, Array(Me.chkReview), mblnSetPara))
    chk�Ƿ�������ʾܾ�.Value = Val(zlDatabase.GetPara("�Ƿ�������ʾܾ�", glngSys, 1342, 1, Array(Me.chk�Ƿ�������ʾܾ�), mblnSetPara))
    chk��ҩ״̬.Value = Val(zlDatabase.GetPara("��ҩ��������Ĭ��Ϊ��ҩ״̬", glngSys, 1342, 0, Array(Me.chk��ҩ״̬), mblnSetPara))
    chk����������ʱ���ܽ�����ҩ����.Value = Val(zlDatabase.GetPara("����������ʱ���ܽ�����ҩ����", glngSys, 1342, 0, Array(Me.chk����������ʱ���ܽ�����ҩ����), mblnSetPara))
    int���ط�ҩʱ�����ҩ���� = Val(zlDatabase.GetPara("���ط�ҩʱ�����ҩ����", glngSys, 1342, 0))
    
    '���ݲ���ֵ����
    opt��ҩ����(int��ҩ����).Value = True
    
    If lngҩ��ID <> 0 Then                                  '��λҩ��
        '�����ڸ�ҩ������ʾ
        For IntLocate = 0 To Me.Cbo��ҩҩ��.ListCount - 1
            If Me.Cbo��ҩҩ��.ItemData(IntLocate) = lngҩ��ID Then
                Me.Cbo��ҩҩ��.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (Cbo��ҩҩ��.ListCount - 1) Then
            MsgBox "����������ҩ����ԭ�����õ�ҩ����ʧЧ����", vbInformation, gstrSysName
            If Cbo��ҩҩ��.ListCount >= 1 Then Cbo��ҩҩ��.ListIndex = 0
        End If
    End If
    Me.Cbo����ģʽ.ListIndex = Lng����ģʽ
    Me.cbo��ҩ�嵥.ListIndex = Lng�Զ���ӡ
    Me.Chk�Ƿ��Զ�ȱҩ���.Value = Lngȱҩ���
    Me.Chk�����һ�����ʾ.Value = Lng������ʾ
    Me.chk��ҩ��ǩ��.Value = Lng��ҩ��ǩ��
    Me.chk��ҩ��ǩ��.Value = Lng��ҩ��ǩ��
    Me.txt��ѯ����.Text = intDays
    Me.chk��Ժ.Value = int��˳�Ժ��������
    Me.cbo��ҩ�嵥.ListIndex = mLng��ӡ��ҩ�嵥
    Me.cboName.ListIndex = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ŷ�ҩ����", "ҩƷ������ʾ��ʽ", 0))
    Me.chk���ط�ҩʱ�����ҩ����.Value = int���ط�ҩʱ�����ҩ����
    
    If chk���ܷ�ҩ.Value = 1 Then
        Me.Chk�����һ�����ʾ.Value = 1
        Me.Chk�����һ�����ʾ.Enabled = False
    End If

    For n = 0 To Cbo������.ListCount - 1
        If Cbo������.List(n) = str������ Then
            Cbo������.ListIndex = n
            Exit For
        End If
    Next
    
    If str������� <> "" Then
        For n = 1 To lvw�������.ListItems.count
            If InStr("," & str������� & ",", "," & lvw�������.ListItems(n).Text & ",") > 0 Then
                lvw�������.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str��ֵ���� <> "" Then
        For n = 1 To lvw��ֵ����.ListItems.count
            If InStr("," & str��ֵ���� & ",", "," & lvw��ֵ����.ListItems(n).Text & ",") > 0 Then
                lvw��ֵ����.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str��Σ���� <> "" Then
        For n = 1 To lvw��Σ����.ListItems.count
            If InStr("," & str��Σ���� & ",", "," & n & ",") > 0 Then
                lvw��Σ����.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str��Σ���� <> "" Then
        If InStr(1, str��Σ����, "1") Then chk��Σ(0).Value = 1
        If InStr(1, str��Σ����, "2") Then chk��Σ(1).Value = 1
        If InStr(1, str��Σ����, "3") Then chk��Σ(2).Value = 1
    End If
    
    '�Զ�ˢ��ʱ��
    If int�Զ�ˢ�� > 0 Then
        chk�Զ�ˢ��.Value = 1
        txt�Զ�ˢ��ʱ��.Text = int�Զ�ˢ��
    End If
    
    If int��ѯ��ҩ���� <= 0 Or int��ѯ��ҩ���� > 99 Then
        int��ѯ��ҩ���� = 7
    End If
    txtTimeArea_Send.Text = int��ѯ��ҩ����
        
    If int��ѯ��ҩ���� <= 0 Or int��ѯ��ҩ���� > 99 Then
        int��ѯ��ҩ���� = 3
    End If
    txtTimeArea_Sended.Text = int��ѯ��ҩ����
    
    If lng����¼�� <= 0 Then
        lng����¼�� = 3000
    End If
    txtMaxRecordCount.Text = lng����¼��
    
End Function

Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case tabShow.Tab
    Case 0
        If Cbo��ҩҩ��.Enabled = True Then Cbo��ҩҩ��.SetFocus
    Case 1
        If Chk�Ƿ��Զ�ȱҩ���.Enabled = True Then Chk�Ƿ��Զ�ȱҩ���.SetFocus
    End Select
End Sub

Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    
    Select Case cboƱ������.ListIndex
    Case 0
        '���ܷ�ҩ��
        strBill = "ZL1_BILL_1342"
    Case 1
        '��ҩ�嵥
        strBill = "ZL1_BILL_1342_1"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub
Private Sub txtMaxRecordCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMaxRecordCount_Validate(Cancel As Boolean)
    If Val(txtMaxRecordCount.Text) <= 0 Then
        txtMaxRecordCount.Text = 3000
    End If
End Sub


Private Sub txtTimeArea_Send_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeArea_Send_Validate(Cancel As Boolean)
    If Val(txtTimeArea_Send.Text) <= 0 Then
        txtTimeArea_Send.Text = 7
    End If
End Sub


Private Sub txtTimeArea_Sended_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeArea_Sended_Validate(Cancel As Boolean)
    If Val(txtTimeArea_Sended.Text) <= 0 Then
        txtTimeArea_Sended.Text = 3
    End If
End Sub


Private Sub txt��ѯ����_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    KeyAscii = 0
End Sub


Private Sub txt�Զ�ˢ��ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub SetSourceDep()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select distinct A.���� || '-' || A.���� ����, A.Id " & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.Id =B.����id and B.�������� in ('���','����','����','����','Ӫ��', '�ٴ�','����') And B.������� In (2,3)  And " & _
            " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By A.���� || '-' || A.����"

    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "SetSourceDep")
    Call SQLTest

    With rs
        If .EOF Then
            MsgBox "û�����ø��ಿ�ţ������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        Lvw��Դ����.ListItems.Clear
        Do While Not .EOF
            Lvw��Դ����.ListItems.Add , "_" & !Id, !����, 1, 1
            If mstrSourceDep <> "" Then
                If InStr("," & mstrSourceDep & ",", "," & CStr(!Id) & ",") > 0 Then
                    Lvw��Դ����.ListItems("_" & !Id).Checked = True
                End If
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



