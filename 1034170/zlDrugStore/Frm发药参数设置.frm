VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form Frm��ҩ�������� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "Frm��ҩ��������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   10605
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic�������� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   3000
      ScaleHeight     =   7905
      ScaleWidth      =   7425
      TabIndex        =   2
      Top             =   120
      Width           =   7455
      Begin TabDlg.SSTab sstMain 
         Height          =   7695
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   13573
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "����"
         TabPicture(0)   =   "Frm��ҩ��������.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "picPar(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "����"
         TabPicture(1)   =   "Frm��ҩ��������.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "picPar(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "��ӡ"
         TabPicture(2)   =   "Frm��ҩ��������.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "picPar(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Ʊ��"
         TabPicture(3)   =   "Frm��ҩ��������.frx":035E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "picPar(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "��Դ����"
         TabPicture(4)   =   "Frm��ҩ��������.frx":037A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "picPar(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "��������"
         TabPicture(5)   =   "Frm��ҩ��������.frx":0396
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "picPar(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "�Ŷӽк�"
         TabPicture(6)   =   "Frm��ҩ��������.frx":03B2
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "picPar(6)"
         Tab(6).ControlCount=   1
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   7095
            Index           =   0
            Left            =   120
            ScaleHeight     =   7095
            ScaleWidth      =   7095
            TabIndex        =   18
            Top             =   480
            Width           =   7095
            Begin VB.Frame frm���˲鿴 
               Caption         =   " ���˲鿴 "
               Height          =   615
               Left            =   120
               TabIndex        =   53
               Top             =   6240
               Width           =   6615
               Begin VB.TextBox txt��ѯ���� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   915
                  TabIndex        =   54
                  Text            =   "1"
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lbl���� 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��"
                  Height          =   180
                  Left            =   1920
                  TabIndex        =   56
                  Top             =   300
                  Width           =   180
               End
               Begin VB.Label lbl��ѯ���� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ѯ����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   55
                  Top             =   300
                  Width           =   720
               End
            End
            Begin VB.Frame frm���ڿ��� 
               Caption         =   " ���ڿ��� "
               Height          =   1215
               Left            =   120
               TabIndex        =   48
               Top             =   4920
               Width           =   6615
               Begin VB.CheckBox chkIsDosage 
                  Caption         =   "��ǰҩ����Ҫ��ҩ����"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   52
                  Top             =   480
                  Width           =   2100
               End
               Begin VB.CheckBox chkIsDosageOk 
                  Caption         =   "��ǰҩ����Ҫ��ҩȷ��(����ǩ��)����"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   51
                  Top             =   240
                  Width           =   3420
               End
               Begin VB.CheckBox chkSign 
                  Caption         =   "ǩ��ʱ�Զ�������ҩ(ҩ������ǩ����Ч)"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   50
                  Top             =   735
                  Width           =   3615
               End
               Begin VB.CheckBox chkCheckStuff 
                  Caption         =   "��ҩ�������ķ������"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   49
                  Top             =   960
                  Width           =   2295
               End
            End
            Begin VB.Frame frm������ʾ 
               Caption         =   " ������ʾ "
               Height          =   975
               Left            =   120
               TabIndex        =   41
               Top             =   3840
               Width           =   6615
               Begin VB.ComboBox cbo����ҩ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1095
                  Style           =   2  'Dropdown List
                  TabIndex        =   44
                  Top             =   600
                  Width           =   2280
               End
               Begin VB.ComboBox cbo���ʴ��� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1095
                  Style           =   2  'Dropdown List
                  TabIndex        =   43
                  Top             =   240
                  Width           =   2280
               End
               Begin VB.ComboBox cbo�շѴ��� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   4215
                  Style           =   2  'Dropdown List
                  TabIndex        =   42
                  Top             =   240
                  Width           =   2280
               End
               Begin VB.Label lbl��ҩ��ӡ״̬ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "����ҩ����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   47
                  Top             =   660
                  Width           =   900
               End
               Begin VB.Label lbl���ʴ��� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "���ʴ���"
                  Height          =   180
                  Left            =   300
                  TabIndex        =   46
                  Top             =   300
                  Width           =   720
               End
               Begin VB.Label lbl�շѴ��� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "�շѴ���"
                  Height          =   180
                  Left            =   3480
                  TabIndex        =   45
                  Top             =   300
                  Width           =   720
               End
            End
            Begin VB.Frame frm�Զ���ҩ 
               Caption         =   " �Զ���ҩ "
               Height          =   975
               Left            =   120
               TabIndex        =   34
               Top             =   2760
               Width           =   6615
               Begin VB.ComboBox cbo�Զ���ҩ���� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1275
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   585
                  Width           =   3360
               End
               Begin VB.TextBox txt��ҩʱ�� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   3675
                  TabIndex        =   36
                  Top             =   240
                  Width           =   525
               End
               Begin VB.CheckBox chk�Զ���ҩ 
                  Caption         =   "�Զ���ҩģʽ"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   35
                  Top             =   300
                  Width           =   1440
               End
               Begin VB.Label lbl�Զ���ҩ���� 
                  AutoSize        =   -1  'True
                  Caption         =   "�Զ���ҩ����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   40
                  Top             =   645
                  Width           =   1080
               End
               Begin VB.Label Label2 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "����"
                  Height          =   180
                  Left            =   4245
                  TabIndex        =   39
                  Top             =   300
                  Width           =   360
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "�Զ���ҩʱ��"
                  Height          =   180
                  Left            =   2520
                  TabIndex        =   38
                  Top             =   300
                  Width           =   1080
               End
            End
            Begin VB.Frame frm��Ա���� 
               Caption         =   " ��Ա���� "
               Height          =   975
               Left            =   120
               TabIndex        =   28
               Top             =   1680
               Width           =   6615
               Begin VB.ComboBox cboCheck 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   855
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   600
                  Width           =   2280
               End
               Begin VB.ComboBox Cbo��ҩ�� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   855
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   240
                  Width           =   2280
               End
               Begin VB.CheckBox chkSame 
                  Caption         =   "������ҩ�˺ͺ˲�����ͬ"
                  Height          =   180
                  Left            =   3360
                  TabIndex        =   29
                  Top             =   660
                  Width           =   2280
               End
               Begin VB.Label lblCheck 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "�˲���"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   33
                  Top             =   660
                  Width           =   540
               End
               Begin VB.Label Lbl��ҩ�� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ��"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   32
                  Top             =   300
                  Width           =   540
               End
            End
            Begin VB.Frame frmҩ������ 
               Caption         =   " ҩ������ "
               Height          =   1455
               Left            =   120
               TabIndex        =   19
               Top             =   120
               Width           =   6855
               Begin VB.ComboBox Cboҩ�� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1320
                  TabIndex        =   23
                  Text            =   "Cboҩ��"
                  Top             =   240
                  Width           =   2280
               End
               Begin VB.ListBox lst��ҩ���� 
                  Appearance      =   0  'Flat
                  Columns         =   3
                  ForeColor       =   &H80000012&
                  Height          =   1080
                  IMEMode         =   3  'DISABLE
                  Left            =   4440
                  Style           =   1  'Checkbox
                  TabIndex        =   22
                  Top             =   300
                  Width           =   2640
               End
               Begin VB.ComboBox cbo�������� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   21
                  Top             =   1080
                  Width           =   2280
               End
               Begin VB.ComboBox cbo��λ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   660
                  Width           =   2280
               End
               Begin VB.Label Lblҩ�� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩҩ��"
                  Height          =   180
                  Left            =   480
                  TabIndex        =   27
                  Top             =   300
                  Width           =   720
               End
               Begin VB.Label Lbl��ҩ���� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ����"
                  Height          =   180
                  Left            =   3720
                  TabIndex        =   26
                  Top             =   300
                  Width           =   720
               End
               Begin VB.Label lbl����סԺ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "����סԺ����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   25
                  Top             =   1140
                  Width           =   1080
               End
               Begin VB.Label lbl��λ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ҩ������"
                  Height          =   180
                  Left            =   480
                  TabIndex        =   24
                  Top             =   720
                  Width           =   720
               End
            End
         End
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   7215
            Index           =   6
            Left            =   -74880
            ScaleHeight     =   7215
            ScaleWidth      =   7095
            TabIndex        =   9
            Top             =   360
            Width           =   7095
            Begin VB.Frame Fra�����豸���� 
               Height          =   3735
               Left            =   120
               TabIndex        =   142
               Top             =   1440
               Width           =   6795
               Begin VB.OptionButton optCallWay 
                  Caption         =   "����Զ������"
                  Height          =   330
                  Index           =   1
                  Left            =   240
                  TabIndex        =   145
                  Top             =   2340
                  Width           =   1455
               End
               Begin VB.CheckBox chkUseSound 
                  Caption         =   "������������"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   144
                  Top             =   0
                  Width           =   1455
               End
               Begin VB.OptionButton optCallWay 
                  Caption         =   "���ñ�������"
                  Height          =   330
                  Index           =   0
                  Left            =   240
                  TabIndex        =   143
                  Top             =   320
                  Width           =   1455
               End
               Begin VB.Frame frm�����㲥���� 
                  Height          =   1935
                  Left            =   120
                  TabIndex        =   146
                  Top             =   360
                  Width           =   6750
                  Begin VB.TextBox txt�㲥ʱ�䳤�� 
                     Height          =   270
                     Left            =   1800
                     TabIndex        =   152
                     Text            =   "10"
                     Top             =   1040
                     Width           =   615
                  End
                  Begin VB.CommandButton cmdTestSound 
                     Caption         =   "��������"
                     Height          =   350
                     Left            =   4080
                     TabIndex        =   151
                     Top             =   645
                     Width           =   1100
                  End
                  Begin VB.TextBox txtSpeed 
                     Height          =   270
                     Left            =   1080
                     TabIndex        =   150
                     Text            =   "65"
                     Top             =   685
                     Width           =   495
                  End
                  Begin VB.OptionButton optSoundType 
                     Caption         =   "ϵͳ����"
                     Height          =   255
                     Index           =   0
                     Left            =   1200
                     TabIndex        =   149
                     Top             =   338
                     Value           =   -1  'True
                     Width           =   1095
                  End
                  Begin VB.OptionButton optSoundType 
                     Caption         =   "΢������"
                     Height          =   255
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   148
                     Top             =   338
                     Width           =   1095
                  End
                  Begin VB.TextBox txtPlayCount 
                     Height          =   270
                     Left            =   1080
                     Locked          =   -1  'True
                     TabIndex        =   147
                     Text            =   "1"
                     Top             =   1515
                     Width           =   495
                  End
                  Begin MSComCtl2.UpDown UpDown 
                     Height          =   390
                     Index           =   0
                     Left            =   1570
                     TabIndex        =   153
                     Top             =   1455
                     Width           =   255
                     _ExtentX        =   450
                     _ExtentY        =   688
                     _Version        =   393216
                     Value           =   1
                     BuddyControl    =   "txtPlayCount"
                     BuddyDispid     =   196658
                     OrigLeft        =   1800
                     OrigTop         =   960
                     OrigRight       =   2055
                     OrigBottom      =   1335
                     Max             =   5
                     Min             =   1
                     SyncBuddy       =   -1  'True
                     BuddyProperty   =   0
                     Enabled         =   -1  'True
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     Caption         =   "ÿ�������㲥����Ϊ        ��"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   157
                     Top             =   1080
                     Width           =   2520
                  End
                  Begin VB.Label Label10 
                     AutoSize        =   -1  'True
                     Caption         =   "�������٣�      (��Χ��0��100֮�䣬�Ƽ�65)"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   156
                     Top             =   730
                     Width           =   3780
                  End
                  Begin VB.Label Label11 
                     AutoSize        =   -1  'True
                     Caption         =   "��������"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   155
                     Top             =   375
                     Width           =   720
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     Caption         =   "���Ŵ���Ϊ          ��(ÿ�κ��в��ŵĴ�������Χ1~5��)"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   154
                     Top             =   1560
                     Width           =   4770
                  End
               End
               Begin VB.Frame FraԶ���������� 
                  Height          =   1215
                  Left            =   120
                  TabIndex        =   158
                  Top             =   2400
                  Width           =   6495
                  Begin VB.ComboBox cboWorkStation 
                     Height          =   300
                     Left            =   1200
                     TabIndex        =   160
                     Top             =   360
                     Width           =   3375
                  End
                  Begin VB.TextBox txtLoopQueryTime 
                     Height          =   270
                     Left            =   3750
                     Locked          =   -1  'True
                     MaxLength       =   3
                     TabIndex        =   159
                     Text            =   "10"
                     Top             =   780
                     Width           =   495
                  End
                  Begin MSComCtl2.UpDown UpDown 
                     Height          =   390
                     Index           =   1
                     Left            =   4250
                     TabIndex        =   161
                     Top             =   720
                     Width           =   255
                     _ExtentX        =   450
                     _ExtentY        =   688
                     _Version        =   393216
                     Value           =   10
                     BuddyControl    =   "txtLoopQueryTime"
                     BuddyDispid     =   196666
                     OrigLeft        =   1920
                     OrigTop         =   1560
                     OrigRight       =   2175
                     OrigBottom      =   2295
                     Max             =   60
                     Min             =   5
                     SyncBuddy       =   -1  'True
                     BuddyProperty   =   0
                     Enabled         =   -1  'True
                  End
                  Begin VB.Label labRemoteComputerName 
                     Caption         =   "Զ��վ������"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   163
                     Top             =   405
                     Width           =   1215
                  End
                  Begin VB.Label Label8 
                     AutoSize        =   -1  'True
                     Caption         =   "������ΪԶ�˺��л���ʱ������ѯ���ʱ��Ϊ         ��(��Χ5~60��)"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   162
                     Top             =   825
                     Width           =   5670
                  End
               End
            End
            Begin VB.CheckBox chk�����Ŷӽк� 
               Caption         =   "�����Ŷӽк�"
               Height          =   255
               Left            =   330
               TabIndex        =   137
               Top             =   120
               Width           =   1455
            End
            Begin VB.CheckBox chkUseDisplay 
               Caption         =   "��ʾ�ŶӶ���"
               Height          =   255
               Left            =   330
               TabIndex        =   136
               Top             =   480
               Width           =   1455
            End
            Begin VB.Frame frm��ʾ�豸���� 
               Height          =   855
               Left            =   120
               TabIndex        =   138
               Top             =   480
               Width           =   6795
               Begin VB.ComboBox cbo��ʾӲ����� 
                  Height          =   300
                  ItemData        =   "Frm��ҩ��������.frx":03CE
                  Left            =   1560
                  List            =   "Frm��ҩ��������.frx":03D0
                  Style           =   2  'Dropdown List
                  TabIndex        =   140
                  Top             =   300
                  Width           =   2535
               End
               Begin VB.CommandButton cmd��ʾ�豸���� 
                  Caption         =   "�豸����"
                  Height          =   300
                  Left            =   4320
                  TabIndex        =   139
                  Top             =   300
                  Width           =   1100
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "��ʾ�豸���"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   141
                  Top             =   360
                  Width           =   1080
               End
            End
         End
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6615
            Index           =   5
            Left            =   -74760
            ScaleHeight     =   6615
            ScaleWidth      =   6975
            TabIndex        =   8
            Top             =   480
            Width           =   6975
            Begin VB.CommandButton cmdDefaultColor 
               BackColor       =   &H00000000&
               Caption         =   "�ָ�Ĭ����ɫ(&R)"
               Height          =   300
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   133
               Top             =   4200
               Width           =   2175
            End
            Begin VB.CommandButton cmdDefaultPrinter 
               BackColor       =   &H00000000&
               Caption         =   "�ָ�Ĭ�ϴ�ӡ��(&P)"
               Height          =   300
               Left            =   4440
               MaskColor       =   &H00000000&
               TabIndex        =   132
               Top             =   4200
               Width           =   2175
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfPrinter 
               Height          =   3735
               Left            =   120
               TabIndex        =   131
               Top             =   360
               Width           =   6495
               _cx             =   11451
               _cy             =   6583
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
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
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
               Editable        =   1
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
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "����������Ͷ��崦����ɫ��"
               Height          =   180
               Left            =   120
               TabIndex        =   135
               Top             =   120
               Width           =   2340
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "ѡ�񴦷���Ӧ�Ĵ�ӡ����������ҩ����ǩ����ҩ������"
               Height          =   180
               Left            =   2520
               TabIndex        =   134
               Top             =   120
               Width           =   4320
            End
         End
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   7095
            Index           =   4
            Left            =   -74760
            ScaleHeight     =   7095
            ScaleWidth      =   6975
            TabIndex        =   7
            Top             =   480
            Width           =   6975
            Begin VB.CommandButton cmdClear 
               Caption         =   "ȫ��(&D)"
               Height          =   350
               Left            =   5640
               TabIndex        =   128
               Top             =   6600
               Width           =   1100
            End
            Begin VB.CommandButton cmdCheckAll 
               Caption         =   "ȫѡ(&A)"
               Height          =   350
               Left            =   4440
               TabIndex        =   127
               Top             =   6600
               Width           =   1100
            End
            Begin MSComctlLib.ListView Lvw��Դ���� 
               Height          =   5805
               Left            =   120
               TabIndex        =   129
               Top             =   480
               Width           =   6675
               _ExtentX        =   11774
               _ExtentY        =   10239
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
            Begin VB.Label lblFrom 
               AutoSize        =   -1  'True
               Caption         =   "������ʾ����Դ���ң���������ѡ����Ĭ����ʾ���п���"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   120
               TabIndex        =   130
               Top             =   120
               Width           =   4500
            End
         End
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6615
            Index           =   3
            Left            =   -74880
            ScaleHeight     =   6615
            ScaleWidth      =   6975
            TabIndex        =   6
            Top             =   480
            Width           =   6975
            Begin VB.ComboBox cboƱ������ 
               Height          =   300
               Left            =   870
               Style           =   2  'Dropdown List
               TabIndex        =   125
               Top             =   120
               Width           =   2565
            End
            Begin VB.CommandButton cmd��ӡ���� 
               Caption         =   "��ӡ����(&P)"
               Height          =   345
               Left            =   120
               TabIndex        =   124
               Top             =   570
               Width           =   3315
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
               Left            =   150
               TabIndex        =   126
               Top             =   180
               Width           =   630
            End
         End
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   7095
            Index           =   2
            Left            =   -74880
            ScaleHeight     =   7095
            ScaleWidth      =   6975
            TabIndex        =   5
            Top             =   480
            Width           =   6975
            Begin VB.Frame frm�Զ�ˢ�� 
               Caption         =   " �Զ�ˢ�� "
               Height          =   1680
               Left            =   120
               TabIndex        =   109
               Top             =   5280
               Width           =   6615
               Begin VB.TextBox Txtˢ�¼�� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1620
                  MaxLength       =   2
                  TabIndex        =   113
                  Top             =   600
                  Width           =   1125
               End
               Begin VB.TextBox Txt�ӳٴ�ӡ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1620
                  MaxLength       =   2
                  TabIndex        =   112
                  Top             =   960
                  Width           =   1125
               End
               Begin VB.TextBox Txt��ӡ�˷ѵ��� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1620
                  MaxLength       =   2
                  TabIndex        =   111
                  Top             =   1320
                  Width           =   1125
               End
               Begin VB.TextBox txt��ӡ��� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1620
                  MaxLength       =   2
                  TabIndex        =   110
                  Top             =   240
                  Width           =   1125
               End
               Begin VB.Label Lblˢ�¼�� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ˢ�¼��"
                  Height          =   180
                  Left            =   840
                  TabIndex        =   123
                  Top             =   660
                  Width           =   720
               End
               Begin VB.Label Lbl�ӳٴ�ӡ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "�ӳٴ�ӡ"
                  Height          =   180
                  Left            =   840
                  TabIndex        =   122
                  Top             =   1020
                  Width           =   720
               End
               Begin VB.Label Lbl��ӡ�˷ѵ��� 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ӡ�˷ѵ��ݼ��"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   121
                  Top             =   1380
                  Width           =   1695
               End
               Begin VB.Label LblNote 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��"
                  Height          =   180
                  Index           =   0
                  Left            =   2760
                  TabIndex        =   120
                  Top             =   660
                  Width           =   180
               End
               Begin VB.Label LblNote 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��"
                  Height          =   180
                  Index           =   1
                  Left            =   2760
                  TabIndex        =   119
                  Top             =   1020
                  Width           =   180
               End
               Begin VB.Label LblNote 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��"
                  Height          =   180
                  Index           =   2
                  Left            =   2760
                  TabIndex        =   118
                  Top             =   1380
                  Width           =   180
               End
               Begin VB.Label LblNote 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��"
                  Height          =   180
                  Index           =   4
                  Left            =   2760
                  TabIndex        =   117
                  Top             =   285
                  Width           =   180
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ӡ���"
                  Height          =   180
                  Left            =   840
                  TabIndex        =   116
                  Top             =   300
                  Width           =   720
               End
               Begin VB.Label lblPrintComment 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��������Ϣ�������"
                  Height          =   180
                  Left            =   3120
                  TabIndex        =   115
                  Top             =   300
                  Width           =   1620
               End
               Begin VB.Label lblRefreshComment 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��������Ϣ�������"
                  Height          =   180
                  Left            =   3120
                  TabIndex        =   114
                  Top             =   660
                  Width           =   1620
               End
            End
            Begin VB.Frame frm�Զ���ӡ 
               Caption         =   " �Զ���ӡ "
               Height          =   2535
               Left            =   120
               TabIndex        =   100
               Top             =   2640
               Width           =   6615
               Begin VB.OptionButton Opt��ӡ��ҩ��ѡ�� 
                  Caption         =   "��ӡָ��������ҩ��"
                  Enabled         =   0   'False
                  Height          =   180
                  Left            =   225
                  TabIndex        =   108
                  Top             =   1095
                  Width           =   2100
               End
               Begin VB.OptionButton Opt��ӡ��ҩ�������� 
                  Caption         =   "��ӡ�����ڵ���ҩ��"
                  Enabled         =   0   'False
                  Height          =   180
                  Left            =   225
                  TabIndex        =   107
                  Top             =   855
                  Width           =   2190
               End
               Begin VB.OptionButton Opt��ӡ��ҩ�������� 
                  Caption         =   "��ӡ�����ŵ���ҩ��"
                  Enabled         =   0   'False
                  Height          =   180
                  Left            =   225
                  TabIndex        =   106
                  Top             =   615
                  Width           =   1935
               End
               Begin VB.CheckBox Chk��ӡ��ҩ�� 
                  Caption         =   "��ӡ��ҩ��"
                  Height          =   210
                  Left            =   1320
                  TabIndex        =   105
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.ListBox lst��ӡ���� 
                  Appearance      =   0  'Flat
                  Columns         =   3
                  Enabled         =   0   'False
                  ForeColor       =   &H80000012&
                  Height          =   1080
                  IMEMode         =   3  'DISABLE
                  Left            =   240
                  Style           =   1  'Checkbox
                  TabIndex        =   104
                  Top             =   1320
                  Width           =   5880
               End
               Begin VB.CheckBox chk���ʵ� 
                  Caption         =   "��ӡʱ�������ʵ���"
                  Height          =   195
                  Left            =   225
                  TabIndex        =   103
                  Top             =   315
                  Width           =   1920
               End
               Begin VB.CheckBox chkҩƷ��ǩ 
                  Caption         =   "��ӡҩƷ��ǩ"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   2640
                  TabIndex        =   102
                  Top             =   0
                  Width           =   1440
               End
               Begin VB.CheckBox chkAllType 
                  Caption         =   "�Զ���ӡ��ҩ��ʱ��ӡƱ�ݵ����и�ʽ"
                  Height          =   195
                  Left            =   2760
                  TabIndex        =   101
                  Top             =   315
                  Width           =   3360
               End
            End
            Begin VB.Frame frm��ӡ���� 
               Caption         =   " ��ӡ���� "
               Height          =   1335
               Left            =   120
               TabIndex        =   92
               Top             =   1200
               Width           =   6615
               Begin VB.ComboBox Cbo��ҩ�� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   900
                  Style           =   2  'Dropdown List
                  TabIndex        =   96
                  Top             =   600
                  Width           =   2520
               End
               Begin VB.ComboBox cbo��ҩ�� 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   900
                  Style           =   2  'Dropdown List
                  TabIndex        =   95
                  Top             =   240
                  Width           =   2520
               End
               Begin VB.ComboBox cboҩƷ��ǩ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   900
                  Style           =   2  'Dropdown List
                  TabIndex        =   94
                  Top             =   960
                  Width           =   2520
               End
               Begin VB.CheckBox chkPreview 
                  Caption         =   "��ӡ����ǩʱ��Ԥ���ٴ�ӡ"
                  Height          =   300
                  Left            =   3540
                  TabIndex        =   93
                  Top             =   600
                  Width           =   2520
               End
               Begin VB.Label lbl��ҩ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ��"
                  Height          =   180
                  Left            =   300
                  TabIndex        =   99
                  Top             =   300
                  Width           =   540
               End
               Begin VB.Label Lbl��ҩ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "����ǩ"
                  Height          =   180
                  Left            =   300
                  TabIndex        =   98
                  Top             =   660
                  Width           =   540
               End
               Begin VB.Label lblҩƷ��ǩ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ҩƷ��ǩ"
                  Height          =   180
                  Left            =   -120
                  TabIndex        =   97
                  Top             =   1020
                  Width           =   975
               End
            End
            Begin VB.Frame frm��ӡ��ʽ 
               Caption         =   " ��ӡ��ʽ "
               Height          =   975
               Left            =   120
               TabIndex        =   83
               Top             =   120
               Width           =   6615
               Begin VB.ComboBox cbo��ҩ��ҩ��ʽ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1260
                  Style           =   2  'Dropdown List
                  TabIndex        =   87
                  Top             =   600
                  Width           =   2040
               End
               Begin VB.ComboBox cbo��ҩ������ʽ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   4500
                  Style           =   2  'Dropdown List
                  TabIndex        =   86
                  Top             =   600
                  Width           =   2040
               End
               Begin VB.ComboBox cbo��ҩ��ҩ��ʽ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1260
                  Style           =   2  'Dropdown List
                  TabIndex        =   85
                  Top             =   240
                  Width           =   2040
               End
               Begin VB.ComboBox cbo��ҩ������ʽ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   4500
                  Style           =   2  'Dropdown List
                  TabIndex        =   84
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.Label lbl��ҩ��ҩ��ʽ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ��ҩ��ʽ"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   91
                  Top             =   660
                  Width           =   1080
               End
               Begin VB.Label lbl��ҩ������ʽ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ������ʽ"
                  Height          =   180
                  Left            =   3360
                  TabIndex        =   90
                  Top             =   660
                  Width           =   1080
               End
               Begin VB.Label lbl��ҩ��ҩ��ʽ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ��ҩ��ʽ"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   89
                  Top             =   300
                  Width           =   1080
               End
               Begin VB.Label lbl��ҩ������ʽ 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ������ʽ"
                  Height          =   180
                  Left            =   3360
                  TabIndex        =   88
                  Top             =   300
                  Width           =   1080
               End
            End
         End
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   7095
            Index           =   1
            Left            =   -74880
            ScaleHeight     =   7095
            ScaleWidth      =   6975
            TabIndex        =   4
            Top             =   480
            Width           =   6975
            Begin VB.Frame frm���� 
               Caption         =   " ���� "
               Height          =   2055
               Left            =   120
               TabIndex        =   67
               Top             =   4200
               Width           =   6615
               Begin VB.CheckBox chkOverTime 
                  Caption         =   "������ʾ"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   80
                  Top             =   915
                  Width           =   1020
               End
               Begin VB.CheckBox Chk��ʾ������ 
                  Caption         =   "��ʾ������"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   79
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.CheckBox chk��С��λ 
                  Caption         =   "�����ֵ�λ��ʾҩƷ����"
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   2760
                  TabIndex        =   77
                  Top             =   240
                  Width           =   2400
               End
               Begin VB.Frame fraline1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   15
                  Left            =   1440
                  TabIndex        =   76
                  Top             =   1140
                  Width           =   650
               End
               Begin VB.CheckBox chkˢ�� 
                  Caption         =   "��ҩʱˢ���￨��֤"
                  Height          =   225
                  Left            =   2760
                  TabIndex        =   75
                  Top             =   465
                  Width           =   3540
               End
               Begin VB.CheckBox chkTakeDrug 
                  Caption         =   "���ò���ʵ��ȡҩȷ��ģʽ"
                  Height          =   225
                  Left            =   2760
                  TabIndex        =   73
                  Top             =   1155
                  Width           =   2460
               End
               Begin VB.CheckBox chksend 
                  Caption         =   "һ��ͨ�շ��뷢ҩ����"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   72
                  Top             =   675
                  Width           =   2295
               End
               Begin VB.CheckBox chkɨ������ 
                  Caption         =   "����ҩ����ɨ����Զ�����"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   71
                  Top             =   1395
                  Width           =   2460
               End
               Begin VB.CheckBox chk��ҩɨ�� 
                  Caption         =   "��ҩģʽ����ɨ����������ɨ��ȷ�ϣ�"
                  Height          =   225
                  Left            =   2760
                  TabIndex        =   70
                  Top             =   675
                  Width           =   3500
               End
               Begin VB.CheckBox chkDispensing 
                  Caption         =   "���������С����ܵ�ͬʱ֪ͨҩƷ�Զ����豸��׼����ҩ��"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   69
                  Top             =   1635
                  Width           =   6015
               End
               Begin VB.CheckBox chk��ҩ���θ��� 
                  Caption         =   "��ҩʱ���۷���ҩƷ�������θ���"
                  Height          =   225
                  Left            =   2760
                  TabIndex        =   68
                  Top             =   1395
                  Width           =   3540
               End
               Begin VB.CheckBox chk�Զ����� 
                  Caption         =   "��ҩʱ�Զ������ʷ�������"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   78
                  Top             =   465
                  Width           =   3540
               End
               Begin VB.CheckBox chk����ʱ�� 
                  Caption         =   "ҩƷҽ��������ʱ�����"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   74
                  Top             =   1155
                  Width           =   2340
               End
               Begin VB.TextBox txtOverTime 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   10.5
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   1545
                  TabIndex        =   81
                  Text            =   "1440"
                  Top             =   915
                  Width           =   460
               End
               Begin VB.Label lblOverTime 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "����       ����δ��ҩ��ҩƷ����"
                  Height          =   180
                  Left            =   1140
                  TabIndex        =   82
                  Top             =   930
                  Width           =   2790
               End
            End
            Begin VB.Frame fra��֤��ʽ 
               Caption         =   " ��֤��ʽ "
               Height          =   600
               Left            =   120
               TabIndex        =   62
               Top             =   840
               Width           =   6645
               Begin VB.CheckBox chkУ�鷢ҩ�� 
                  Caption         =   "У�鷢ҩ��"
                  Height          =   195
                  Left            =   2190
                  TabIndex        =   66
                  Top             =   300
                  Width           =   1200
               End
               Begin VB.CheckBox chkУ����ҩ�� 
                  Caption         =   "У����ҩ��"
                  Height          =   195
                  Left            =   420
                  TabIndex        =   65
                  Top             =   300
                  Width           =   1200
               End
               Begin VB.OptionButton Opt��֤��ʽ 
                  Caption         =   "�û�����֤"
                  Height          =   180
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   64
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1245
               End
               Begin VB.OptionButton Opt��֤��ʽ 
                  Caption         =   "������֤"
                  Height          =   180
                  Index           =   1
                  Left            =   2520
                  TabIndex        =   63
                  Top             =   0
                  Width           =   1095
               End
            End
            Begin VB.Frame frm�豸ģʽ 
               Caption         =   " �豸ģʽ "
               Height          =   1095
               Left            =   120
               TabIndex        =   57
               Top             =   3000
               Width           =   6615
               Begin VB.CommandButton cmdDeviceSetup 
                  Caption         =   "�豸����(&S)"
                  Height          =   350
                  Left            =   2160
                  TabIndex        =   59
                  Top             =   600
                  Width           =   1500
               End
               Begin VB.ComboBox cbo�س���ʽ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   58
                  Top             =   240
                  Width           =   3360
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "���ܿ��������豸����"
                  Height          =   180
                  Left            =   300
                  TabIndex        =   61
                  Top             =   690
                  Width           =   1800
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "����ʱϵͳ�Զ��س���ʽ"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   60
                  Top             =   300
                  Width           =   1980
               End
            End
            Begin VB.Frame frmˢ������ 
               Height          =   1335
               Left            =   120
               TabIndex        =   15
               Top             =   1560
               Width           =   6615
               Begin VB.ListBox lst������ 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  ForeColor       =   &H80000012&
                  Height          =   870
                  IMEMode         =   3  'DISABLE
                  Left            =   120
                  Style           =   1  'Checkbox
                  TabIndex        =   17
                  Top             =   360
                  Width           =   6360
               End
               Begin VB.CheckBox chk��ҩˢ�� 
                  Caption         =   "��ҩģʽ����ˢ����ҩ"
                  Height          =   225
                  Left            =   240
                  TabIndex        =   16
                  Top             =   0
                  Width           =   2100
               End
            End
            Begin VB.Frame frm��ʾ��ʽ 
               Caption         =   " ��ʾ��ʽ "
               Height          =   615
               Left            =   120
               TabIndex        =   10
               Top             =   120
               Width           =   6855
               Begin VB.ComboBox cboҩƷ������ʾ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   12
                  Top             =   240
                  Width           =   2160
               End
               Begin VB.ComboBox cbo�����ʾ 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   4560
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Top             =   240
                  Width           =   2160
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "ҩƷ������ʾ"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   14
                  Top             =   300
                  Width           =   1080
               End
               Begin VB.Label lbl�����ʾ 
                  AutoSize        =   -1  'True
                  Caption         =   "�����ʾ"
                  Height          =   180
                  Left            =   3720
                  TabIndex        =   13
                  Top             =   300
                  Width           =   720
               End
            End
         End
      End
   End
   Begin VB.PictureBox pic����ҳǩ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   1200
      ScaleHeight     =   7905
      ScaleWidth      =   1665
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin XtremeSuiteControls.TaskPanel tplFunc 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   5318
         _StockProps     =   64
         Behaviour       =   1
         ItemLayout      =   2
         HotTrackStyle   =   3
      End
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4560
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
            Picture         =   "Frm��ҩ��������.frx":03D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ҩ��������.frx":06EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   3480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "Frm��ҩ��������.frx":0A06
      Left            =   120
      Top             =   2520
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.ImageManager imgFunc 
      Left            =   120
      Top             =   1800
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "Frm��ҩ��������.frx":0A1A
   End
End
Attribute VB_Name = "Frm��ҩ��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--ע�����ر���--
Private intDays As Integer
Private intUnit As Integer                              'ȱʡ��λ��0-����Ӧ;1-����ҩ����λ;2-סԺҩ����λ��
Private intPrint As Integer                             '����ӡδ��ҩ����(0)
Private intУ�鷽ʽ As Integer                          'У�鷽ʽ
Private intУ����ҩ�� As Integer                        '��ҩʱ�Ƿ�У����ҩ��
Private intУ�鷢ҩ�� As Integer                        '��ҩʱ�Ƿ�У�鷢ҩ��
Private mint���ʵ� As Integer                           '��ӡ��ҩ��ʱ�Ƿ�������ʵ�
Private mintҩƷ��ǩ As Integer                         '��ӡҩƷ��ǩ
Private strPrintWindow As String                        '��ӡδ��ҩ����Ϊ3ʱ��Ч
'0-����ӡδ��ҩ����
'1-��ӡ����������δ��ҩ����
'2-��ӡ����������δ��ҩ����
'3-ѡ���ӡ(��ҩ����)

Private IntRefresh As Integer                           'ˢ�¼��(0)
Private intPrintDelay As Integer                        '�ӳٴ�ӡ(60)
Private intPrintHandbackNO As Integer                   '��ӡ�˷ѵ��ݺ�(0)
Private mintPrintInterval As Integer                    '��ӡ��ҩ�����(0)
Private lngҩ��ID As Long                               'ҩ��(���ñ�������Ӧ��ҩ��)
Private Str���� As String                               '��ҩ����(���ñ�������Ӧ�ķ�ҩ����)
Private str��ҩ�� As String                             '������ҩ��
Private mint�Զ���ҩ As Integer                         '�Ƿ�ʹ���Զ���ҩ���ܣ�0-��ʹ�ã�1-ʹ��
Private mint�Զ���ҩʱ�� As Integer                     '������ʱ�޾���Ҫ��֤��ҩ�ˣ�Ĭ��Ϊʼ�ղ���֤��ҩ��
Private mintˢ��֤ As Integer                           '��ҩ���Ƿ����ˢ����֤��0-��ˢ��;1-Ҫˢ��
Private mint��ҩɨ�� As Integer                         '��ҩģʽ����ɨ������0-������;1-����
Private mint�����Ŷӽк� As Integer                     '�Ƿ������ŶӽкŹ���
Private mintSign As Integer                             'ǩ��ʱ������ҩ
Private mblnLoadDrug As Boolean
Private mblnUseMsg As Boolean                           '�Ƿ���������Ϣ����
Private mstr����ˢ����ҩ As String                      '����ˢ����ҩ����ʽ�������1,�����2......����������ݱ�ʾ������
Private mint����ʱ����� As Integer                     'ҩƷҽ��������ʱ�䣨�״�ʱ�䣩���ˣ�0-������ʱ����ˣ�1-������ʱ�����
Private mint�����ʾ��ʽ As Integer                     '0-��ʾӦ�ս�1-��ʾʵ�ս�2-��ʾӦ�ս���ʵ�ս��
Private mint����ȡҩģʽ As Integer                     '����ȡҩģʽ��0-�����ã�1-����
Private mintһ��ͨ��ҩģʽ As Integer                   'һ��ͨ��ҩģʽ��0-�շ��뷢ҩͬʱ���У�1-��ҩ���շѷ���
Private mintɨ������ As Integer                       '0-���Զ�����,1-ɨ����Զ���������
Private mstr�˲��� As String
Private mint��ҩ���θ��� As Integer                     '0-�����ã�1-���á����ú󣬶��۷���ҩƷ������ҩʱ���Ϊ�ϸ��飬�ҿ��ʵ����������ʱ�����Զ�Ѱ�ҿ���㹻���������β��滻����
Private mintRowNum As Integer
Private mint��ҩ���� As Integer                       '��ҩ���Ƿ����б�ҩ���˲�����δ�������ĵ���

Private mintShowName As Integer                         'ҩƷ������ʾ��ʽ��0-���ƺͱ��룻1-�����룻2-������
Private mintType As Integer                             '�������ͣ�0-��ʾ�����סԺ������1-ֻ��ʾ���ﴦ����2-ֻ��ʾסԺ����

Private IntShowCol As Integer                           '�ڴ�����ϸ���Ƿ���ʾ����(0)
Private mintShowBill�շ� As Integer                     '�շѴ�����ʾ��Χ
Private mintShowBill���� As Integer                     '���ʴ�����ʾ��Χ
Private mintShowBill��ҩ As Integer                     '����ҩ����ӡ״̬��ʾ��Χ
Private IntAutoPrint As Integer                         '��ҩ���ӡ������(1)
Private mint��ҩ���Զ���ӡ As Integer                   '�Զ���ӡ��ҩ��
Private mint��ҩ���Զ���ӡҩƷ��ǩ As Integer           '��ҩ���Զ���ӡҩƷ��ǩ
Private mstrWin As String                               '��ҩ���ڴ�
Private mint�س���ʽ As Integer                         'ͨ��¼���ˢ������ʱϵͳ�Զ���ӻس�����ķ�ʽ��0-ϵͳ���Զ��س�,1-��¼��ﵽ��Ŀ�򿨺ų���ʱ�Զ��س�
Private mint�Զ���ҩ���� As Integer                     '0-ȫ�������Զ���ҩ;1-���Ӵ���(��ҽ���Ĵ���)�Զ���ҩ;2-�ֹ�����(��ҽ���Ĵ���)�Զ���ҩ

Private mIntCol���� As Integer
Private mintCol��ʽ As Integer
Private mintCol��ӡ�� As Integer

Private Const mconstr���� = "��ͨ;����;����;�������;�������;����"
Private Const mconlng��ɫ = "&HFFFFFF;&HC0FFC0;&HC0FFFF;&HFFFFFF;&HC0C0FF;&HC0C0FF"

Public mstrPrivs As String                              'Ȩ�޴�
Private mblnSetPara As Boolean                          '�Ƿ���в�������Ȩ��
Private mstrRPTDefaultScheme_Recipt As String           '����ǩ�����Ĭ�ϸ�ʽ

'�Ŷӽк�ʹ�õĲ���

Private Type Type_Call
    int�����Ŷӽк� As Integer
    int�������� As Integer
    int��ʾģʽ As Integer
    int��ʾ�ŶӶ��� As Integer
    int������������ As Integer
    int�кŷ�ʽ As Integer
    strԶ�˺���վ�� As String
    int�����㲥ʱ�䳤�� As Integer
    int�����㲥���� As Integer
    int�������Ŵ��� As Integer
    int��ѯʱ�� As Integer
End Type

Private mType_Call As Type_Call
'--��������ʹ�õĶ���--
Public RecPart As New ADODB.Recordset                   'ҩ��
Private RecPeople As New ADODB.Recordset                'ҩ����ҩ��
Private BlnStartUp  As Boolean                          '�Ƿ������ɹ�
Public strShow As String                                '��ʾ��
Private mstrSourceDep As String                         '��Դ���Ҵ�

Private mstrPrinters As String                          '���ش�ӡ���б���;�ָ�

'�������ͣ���ͨ��������ơ�������һ������
Private Enum ��������
    ��ͨ = 0
    ���� = 1
    ���� = 2
    ���� = 3
    ��һ = 4
    ���� = 5
End Enum

'Ĭ�ϴ�����ɫ����ͨ����ɫ���������ɫ�����ƣ�����ɫ��������һ������ɫ����������ɫ
Private Const mconlng��ͨ = &HFFFFFF
Private Const mconlng���� = &HC0FFC0
Private Const mconlng���� = &HC0FFFF
Private Const mconlng���� = &HFFFFFF
Private Const mconlng��һ = &HC0C0FF
Private Const mconlng���� = &HC0C0FF

Public Property Get In_���÷�ҩ() As Boolean
    In_���÷�ҩ = mblnLoadDrug
End Property

Public Property Let In_���÷�ҩ(ByVal vNewValue As Boolean)
    mblnLoadDrug = vNewValue
End Property

Public Property Get In_������Ϣ() As Boolean
    In_������Ϣ = mblnUseMsg
End Property

Public Property Let In_������Ϣ(ByVal vNewValue As Boolean)
    mblnUseMsg = vNewValue
End Property


Private Sub LoadList()
    Dim rs��ҩ��ʽ As New ADODB.Recordset
    Dim rs��ҩ��ʽ As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim str��ҩ��ʽ As String
    Dim str������ʽ As String
    Dim str��� As String
    Dim strPrinter As String
    Dim strPrinters As String
    Dim strColor As String
    Dim myPrinter As Printer
    Dim n As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    mIntCol���� = 0
    mintCol��ʽ = 1
    mintCol��ӡ�� = 2
    
    '��ȡ�����ʽ
    '--��ҩ
    str��� = "ZL1_BILL_1341_3"
    
    gstrSQL = "Select b.˵�� From zlReports A, zlRPTFMTs B Where a.Id = b.����id And a.��� = [1] order by b.���"
    
    Set rs��ҩ��ʽ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ�����ʽ", str���)
    
    '--��ҩ
    str��� = "ZL1_BILL_1341_4"
    
    gstrSQL = "Select b.˵�� From zlReports A, zlRPTFMTs B Where a.Id = b.����id And a.��� = [1] order by b.���"
    
    Set rs��ҩ��ʽ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ�����ʽ", str���)
    
    '��ȡ�������͵���ɫ����
    strColor = zlDatabase.GetPara("������ɫ", glngSys, 1341, "", , mblnSetPara)
    
    '��ȡ����Ĵ�ӡ����������
    strPrinter = zlDatabase.GetPara("������Ӧ�Ĵ�ӡ��", glngSys, 1341, "", , mblnSetPara)
    
    '��ȡ��Ӧ�Ĵ�ӡ��ʽ����
    str��ҩ��ʽ = zlDatabase.GetPara("��ҩ����ӡ��ʽ", glngSys, 1341, "2;2", , mblnSetPara)
    str������ʽ = zlDatabase.GetPara("����ǩ��ӡ��ʽ", glngSys, 1341, "1;1", , mblnSetPara)
    
    '��Ӵ�ӡ��ʽ�������б�
    With rs��ҩ��ʽ
        For n = 1 To .RecordCount
            cbo��ҩ��ҩ��ʽ.AddItem !˵��
            cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.NewIndex) = n
            cbo��ҩ������ʽ.AddItem !˵��
            cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.NewIndex) = n
            .MoveNext
        Next
    End With
    
    With rs��ҩ��ʽ
        For n = 1 To .RecordCount
            cbo��ҩ��ҩ��ʽ.AddItem !˵��
            cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.NewIndex) = n
            cbo��ҩ������ʽ.AddItem !˵��
            cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.NewIndex) = n
            .MoveNext
        Next
    End With
    
    '�����û����õĴ�ӡ��ʽ
    '--��ҩ
    For i = 0 To cbo��ҩ��ҩ��ʽ.ListCount - 1
        If Val(Split(str��ҩ��ʽ, ";")(0)) = cbo��ҩ��ҩ��ʽ.ItemData(i) Then
            cbo��ҩ��ҩ��ʽ.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To cbo��ҩ������ʽ.ListCount - 1
        If Val(Split(str������ʽ, ";")(0)) = cbo��ҩ������ʽ.ItemData(i) Then
            cbo��ҩ������ʽ.ListIndex = i
            Exit For
        End If
    Next
    '--��ҩ
    For i = 0 To cbo��ҩ��ҩ��ʽ.ListCount - 1
        If Val(Split(str��ҩ��ʽ, ";")(1)) = cbo��ҩ��ҩ��ʽ.ItemData(i) Then
            cbo��ҩ��ҩ��ʽ.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To cbo��ҩ������ʽ.ListCount - 1
        If Val(Split(str������ʽ, ";")(1)) = cbo��ҩ������ʽ.ItemData(i) Then
            cbo��ҩ������ʽ.ListIndex = i
            Exit For
        End If
    Next
    
    '���뱾�ش�ӡ���б�
    mstrPrinters = ""
    For Each myPrinter In Printers
        mstrPrinters = IIf(mstrPrinters = "", "", mstrPrinters & ";") & myPrinter.DeviceName
    Next
    
    For n = 0 To UBound(Split(mstrPrinters, ";"))
        If Split(mstrPrinters, ";")(n) <> "" Then
            strPrinters = strPrinters & "|" & Split(mstrPrinters, ";")(n)
        End If
    Next
    strPrinters = Mid(strPrinters, 2)
    
    'װ�ر��ؼ�¼��
    With rsData
        If .State = 1 Then .Close
        
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ʽ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ӡ��", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '�ж����ݵĺϷ��ԣ�����������
    If UBound(Split(strPrinter, ";")) <> UBound(Split(mconstr����, ";")) Then
        For n = 0 To UBound(Split(mconstr����, ";"))
            strPrinter = strPrinter & ";"
        Next
    End If
    
    '�򱾵ؼ�¼�������û�����Ĵ�ӡ������
    For n = 0 To UBound(Split(mconstr����, ";"))
        rs��ҩ��ʽ.MoveFirst
        If InStr(strPrinter, "?") = 0 Then
            For i = 1 To rs��ҩ��ʽ.RecordCount
                rsData.AddNew
                
                rsData!���� = Split(mconstr����, ";")(n)
                rsData!��ʽ = rs��ҩ��ʽ!˵��
                rsData!��ӡ�� = Split(strPrinter, ";")(n)
    
                rsData.Update
                rs��ҩ��ʽ.MoveNext
            Next
        Else
            For i = 0 To UBound(Split(Split(strPrinter, ";")(n), ","))
                rsData.AddNew
                
                rsData!���� = Split(mconstr����, ";")(n)
                rsData!��ʽ = Mid(Split(Split(strPrinter, ";")(n), ",")(i), 1, InStr(Split(Split(strPrinter, ";")(n), ",")(i), "?") - 1)
                rsData!��ӡ�� = Mid(Split(Split(strPrinter, ";")(n), ",")(i), InStr(Split(Split(strPrinter, ";")(n), ",")(i), "?") + 1)
             
            Next
        End If
        rsData.Update
    Next
        
    With vsfPrinter
        .rows = rs��ҩ��ʽ.RecordCount * 6
        .Cols = 3
        .AllowSelection = False
        .ColAlignment(mIntCol����) = flexAlignCenterCenter
        .RowHeight(-1) = 250
        .ColWidth(mIntCol����) = 900
        .ColWidth(mintCol��ʽ) = 1500
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(mIntCol����) = True
        
        '���ش�ӡ��ѡ�������
        .ColComboList(mintCol��ӡ��) = strPrinters
        .ColComboList(mIntCol����) = "..."
        
        '����[����&��ɫ]��[��ʽ]
        For n = 0 To UBound(Split(mconstr����, ";"))
            rs��ҩ��ʽ.MoveFirst
            For i = 1 To rs��ҩ��ʽ.RecordCount
                .TextMatrix(n * rs��ҩ��ʽ.RecordCount + i - 1, mIntCol����) = Split(mconstr����, ";")(n)
                
                If strColor <> "" Then
                    .Cell(flexcpBackColor, n * rs��ҩ��ʽ.RecordCount + i - 1, mIntCol����) = Val(Split(strColor, ";")(n))
                Else
                    .Cell(flexcpBackColor, n * rs��ҩ��ʽ.RecordCount + i - 1, mIntCol����) = Split(mconlng��ɫ, ";")(n)
                End If
                
                .TextMatrix(n * rs��ҩ��ʽ.RecordCount + i - 1, mintCol��ʽ) = rs��ҩ��ʽ!˵��
                
                rs��ҩ��ʽ.MoveNext
            Next
        Next
        
        '�����û�����Ĵ�ӡ������
        For n = 0 To .rows - 1
            rsData.Filter = "���� = '" & .TextMatrix(n, mIntCol����) & "' and ��ʽ = '" & .TextMatrix(n, mintCol��ʽ) & "'"
            If rsData.RecordCount > 0 Then
                If InStr(strPrinters & "|", rsData!��ӡ�� & "|") > 0 Then   '���ô�ӡ�������Ƿ����
                    .TextMatrix(n, mintCol��ӡ��) = rsData!��ӡ��
                End If
            End If
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub







Private Sub GetDefaultRecipeColor()
    Dim intTemp As Integer
    
    With vsfPrinter
        intTemp = .rows / 6
        
        .Cell(flexcpBackColor, intTemp * 0, mIntCol����) = mconlng��ͨ
        .Cell(flexcpBackColor, intTemp * 1, mIntCol����) = mconlng����
        .Cell(flexcpBackColor, intTemp * 2, mIntCol����) = mconlng����
        .Cell(flexcpBackColor, intTemp * 3, mIntCol����) = mconlng����
        .Cell(flexcpBackColor, intTemp * 4, mIntCol����) = mconlng��һ
        .Cell(flexcpBackColor, intTemp * 5, mIntCol����) = mconlng����
    End With
End Sub

Private Function ReadFromReg()
    Dim strTmp As String
    Dim intOverTime As Integer
    
    On Error Resume Next
    
    mblnSetPara = IsHavePrivs(mstrPrivs, "��������")
    
    'ȡ������˽�в���
    intУ�鷢ҩ�� = Val(zlDatabase.GetPara("У�鷢ҩ��", glngSys, 1341, 0, Array(chkУ�鷢ҩ��), mblnSetPara))
    intУ�鷽ʽ = Val(zlDatabase.GetPara("У�鷽ʽ", glngSys, 1341, 0, Array(fra��֤��ʽ, Opt��֤��ʽ(0), Opt��֤��ʽ(1)), mblnSetPara))
    intУ����ҩ�� = Val(zlDatabase.GetPara("У����ҩ��", glngSys, 1341, 0, Array(chkУ����ҩ��), mblnSetPara))
    chk�Զ�����.Value = Val(zlDatabase.GetPara("�Զ�����", glngSys, 1341, 0, Array(chk�Զ�����), mblnSetPara))

    mintShowBill�շ� = Val(zlDatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, 1341, 3, Array(lbl�շѴ���, cbo�շѴ���), mblnSetPara))
    mintShowBill���� = Val(zlDatabase.GetPara("���ʴ�����ʾ��ʽ", glngSys, 1341, 3, Array(lbl���ʴ���, cbo���ʴ���), mblnSetPara))
    mintShowBill��ҩ = Val(zlDatabase.GetPara("����ҩ���ݴ�ӡ��ʾ��ʽ", glngSys, 1341, 0, Array(lbl��ҩ��ӡ״̬, cbo����ҩ), mblnSetPara))
    intDays = Val(zlDatabase.GetPara("��ѯ����", glngSys, 1341, 1, Array(lbl��ѯ����, txt��ѯ����, lbl����), mblnSetPara))
    mint���ʵ� = Val(zlDatabase.GetPara("��ӡ�������ʵ�", glngSys, 1341, 0, Array(chk���ʵ�), mblnSetPara))
    intPrintHandbackNO = Val(zlDatabase.GetPara("��ӡ�˷ѵ��ݼ��", glngSys, 1341, 0, Array(Lbl��ӡ�˷ѵ���, Txt��ӡ�˷ѵ���, LblNote(2)), mblnSetPara))
    intPrintDelay = Val(zlDatabase.GetPara("��ӡ�ӳ�", glngSys, 1341, 60, Array(Lbl�ӳٴ�ӡ, Txt�ӳٴ�ӡ, LblNote(1)), mblnSetPara))
    IntRefresh = Val(zlDatabase.GetPara("ˢ�¼��", glngSys, 1341, 0, Array(Lblˢ�¼��, Txtˢ�¼��, LblNote(0)), mblnSetPara))
    mintPrintInterval = Val(zlDatabase.GetPara("��ӡ���", glngSys, 1341, 0, Array(Label3, txt��ӡ���, LblNote(4)), mblnSetPara))
    IntShowCol = Val(zlDatabase.GetPara("��ʾ����", glngSys, 1341, 0, Array(Chk��ʾ������), mblnSetPara))
    IntAutoPrint = Val(zlDatabase.GetPara("��ҩ���Զ���ӡ", glngSys, 1341, 0, Array(Lbl��ҩ, Cbo��ҩ��), mblnSetPara))
    intUnit = Val(zlDatabase.GetPara("ҩ������", glngSys, 1341, 0, Array(lbl��λ, cbo��λ), mblnSetPara))
    mint��ҩ���Զ���ӡ = Val(zlDatabase.GetPara("��ҩ���Զ���ӡ", glngSys, 1341, 2, Array(lbl��ҩ, cbo��ҩ��), mblnSetPara))
    mint��ҩ���Զ���ӡҩƷ��ǩ = Val(zlDatabase.GetPara("��ҩ���ӡҩƷ��ǩ", glngSys, 1341, 2, Array(lblҩƷ��ǩ, cboҩƷ��ǩ), mblnSetPara))
    mintˢ��֤ = Val(zlDatabase.GetPara("��ҩ��ˢ����֤", glngSys, 1341, 0, Array(chkˢ��), mblnSetPara))
    mint��ҩɨ�� = Val(zlDatabase.GetPara("��ҩģʽɨ����ȷ��", glngSys, 1341, 0, Array(chk��ҩɨ��), mblnSetPara))
    intOverTime = Val(zlDatabase.GetPara("��ʱδ��ҩƷ��ʾʱ����", glngSys, 1341, 0, Array(chkOverTime, lblOverTime, txtOverTime, fraline1), mblnSetPara))
    mintType = Val(zlDatabase.GetPara("������סԺ����", glngSys, 1341, 0, Array(lbl����סԺ, cbo��������), mblnSetPara))
    mintSign = Val(zlDatabase.GetPara("ǩ��ʱ������ҩ", glngSys, 1341, 0, Array(chkSign), mblnSetPara))
    mstr����ˢ����ҩ = zlDatabase.GetPara("����ˢ����ҩ", glngSys, 1341, "", Array(chk��ҩˢ��, lst������), mblnSetPara)
    mint����ʱ����� = zlDatabase.GetPara("ҩƷҽ��������ʱ�����", glngSys, 1341, 0, Array(chk����ʱ��), mblnSetPara)
    mint�����ʾ��ʽ = Val(zlDatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0, Array(lbl�����ʾ, cbo�����ʾ), mblnSetPara))
    mint����ȡҩģʽ = zlDatabase.GetPara("���ò���ʵ��ȡҩȷ��ģʽ", glngSys, 1341, 0, Array(chkTakeDrug), mblnSetPara)
    mintһ��ͨ��ҩģʽ = zlDatabase.GetPara("һ��ͨ�շ��뷢ҩ����", glngSys, 1341, 0, Array(chksend), mblnSetPara)
    mintɨ������ = Val(zlDatabase.GetPara("����ҩ����ɨ����Զ�����", glngSys, 1341, 0, Array(chkɨ������), mblnSetPara))
    mint�س���ʽ = Val(zlDatabase.GetPara("����ʱϵͳ�Զ��س���ʽ", glngSys, 1341, 0, Array(cbo�س���ʽ), mblnSetPara))
    mint��ҩ���θ��� = Val(zlDatabase.GetPara("��ҩ���θ���", glngSys, 1341, 0, Array(chk��ҩ���θ���), mblnSetPara))

    With mType_Call
        .int�кŷ�ʽ = Val(zlDatabase.GetPara("�кŷ�ʽ", glngSys, 1341, 0))
        .int�����Ŷӽк� = Val(zlDatabase.GetPara("�����Ŷӽк�", glngSys, 1341, 0, Array(chk�����Ŷӽк�), mblnSetPara))
        .int������������ = Val(zlDatabase.GetPara("������������", glngSys, 1341, 0))
        .int��ʾģʽ = Val(zlDatabase.GetPara("��ʾģʽ", glngSys, 1341, 0))
        .int��ʾ�ŶӶ��� = Val(zlDatabase.GetPara("��ʾ�ŶӶ���", glngSys, 1341, 0))
        .int�������Ŵ��� = Val(zlDatabase.GetPara("�������Ŵ���", glngSys, 1341, 0))
        .int�����㲥ʱ�䳤�� = Val(zlDatabase.GetPara("�����㲥ʱ�䳤��", glngSys, 1341, 0))
        .int�����㲥���� = Val(zlDatabase.GetPara("�����㲥����", glngSys, 1341, 0))
        .int�������� = Val(zlDatabase.GetPara("��������", glngSys, 1341, 0))
        .strԶ�˺���վ�� = zlDatabase.GetPara("Զ�˺���վ��", glngSys, 1341, "")
        .int��ѯʱ�� = Val(zlDatabase.GetPara("������ѯʱ��", glngSys, 1341, 10))
        
        '������ѯ�ķǷ�����ʱ�䡣35.110��,��ѯʱ��Ϊ5~60��
        If .int��ѯʱ�� < 5 Or .int��ѯʱ�� > 60 Then .int��ѯʱ�� = 10
    End With
    
    '0-����ӡδ��ҩ����
    '1-��ӡ����������δ��ҩ����
    '2-��ӡ����������δ��ҩ����
    '3-ѡ���ӡ(��ҩ����)
    intPrint = Val(zlDatabase.GetPara("�����µ����Ƿ��ӡ", glngSys, 1341, 0, Array(Chk��ӡ��ҩ��), mblnSetPara))
    
    mintҩƷ��ǩ = Val(zlDatabase.GetPara("��ӡҩƷ��ǩ", glngSys, 1341, 0, Array(chkҩƷ��ǩ), mblnSetPara))
    lngҩ��ID = Val(zlDatabase.GetPara("��ҩҩ��", glngSys, 1341, 0, Array(Lblҩ��, , Cboҩ��), mblnSetPara))
    Str���� = zlDatabase.GetPara("��ҩ����", glngSys, 1341, "", Array(Lbl��ҩ����, lst��ҩ����), mblnSetPara)
    str��ҩ�� = zlDatabase.GetPara("��ҩ��", glngSys, 1341, "", Array(Lbl��ҩ��, Cbo��ҩ��), mblnSetPara)
    mstr�˲��� = zlDatabase.GetPara("�˲���", glngSys, 1341, "", Array(lblCheck, cboCheck), mblnSetPara)
    strPrintWindow = zlDatabase.GetPara("��ӡָ����ҩ����", glngSys, 1341, "", Array(Opt��ӡ��ҩ��ѡ��, lst��ӡ����), mblnSetPara)
    mstrSourceDep = zlDatabase.GetPara("��Դ����", glngSys, 1341, "", Array(Lvw��Դ����), mblnSetPara)
    mint�Զ���ҩ = Val(zlDatabase.GetPara("�Զ���ҩ", glngSys, 1341, 0, Array(chk�Զ���ҩ), mblnSetPara))
    mint�Զ���ҩʱ�� = Val(zlDatabase.GetPara("�Զ���ҩʱ��", glngSys, 1341, 0, Array(Label1, txt��ҩʱ��, Label2), mblnSetPara))
    mint�Զ���ҩ���� = Val(zlDatabase.GetPara("�Զ���ҩ����", glngSys, 1341, 0, Array(lbl�Զ���ҩ����, cbo�Զ���ҩ����), mblnSetPara))
    
    chkAllType.Value = (zlDatabase.GetPara("��ӡƱ�ݵ����и�ʽ", glngSys, 1341, 0, Array(chkAllType), mblnSetPara))
    chkSame.Value = (zlDatabase.GetPara("����˲��˺���ҩ����ͬ", glngSys, 1341, 0, Array(chkSame), mblnSetPara))
    chkPreview.Value = zlDatabase.GetPara("��ӡ����ǩʱ��Ԥ���ٴ�ӡ", glngSys, 1341, 0, Array(chkPreview), mblnSetPara)
    mint��ҩ���� = zlDatabase.GetPara("��ҩ�������ķ������", glngSys, 1341, 0, Array(chkCheckStuff), mblnSetPara)
    
    If lngҩ��ID <> 0 Then
        Call SetDispense
    End If
    
    strTmp = zlDatabase.GetPara("������", glngSys, 1341, "0", Array(Label4, cboҩƷ������ʾ), mblnSetPara)
    If InStr(1, strTmp, "|") > 0 Then
        mintShowName = Val(Mid(strTmp, 1, 1))
    Else
        mintShowName = Val(strTmp)
    End If
    If mintShowName > 2 Or mintShowName < 0 Then mintShowName = 0
    
    chk��С��λ.Value = Val(zlDatabase.GetPara("��ʾ��С��λ", glngSys, 1341, 0, Array(chk��С��λ), mblnSetPara))
    
    If intOverTime < 0 Or intOverTime > 1440 Then
        intOverTime = 0
    End If
    intOverTime = Int(intOverTime)
    chkOverTime.Value = IIf(intOverTime = 0, 0, 1)
    If chkOverTime.Value = 0 Then
        txtOverTime.Text = ""
        txtOverTime.Enabled = False
    Else
        txtOverTime.Text = intOverTime
        txtOverTime.Enabled = True
    End If
    
    Call LoadList
    
    'ʹ���˵���ǩ���Ͳ�����ͨ��У�鷽ʽ
    If gblnESign������ҩ = True Then
        fra��֤��ʽ.Enabled = False
        Opt��֤��ʽ(0).Enabled = False
        Opt��֤��ʽ(1).Enabled = False
        chkУ�鷢ҩ��.Enabled = False
        chkУ����ҩ��.Enabled = False
    End If
End Function

Private Sub SetSourceDep()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select ���� || '-' || ���� ����, Id " & _
            " From ���ű� " & _
            " Where Id In (Select ����id From ��������˵�� Where �������� = '�ٴ�' And ������� In (1,2,3)) And " & _
            " (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By ���� || '-' || ���� "

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

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_File_RecipePar_Save             '����
            Call ����
        Case mconMenu_File_RecipePar_Cancel           '�˳�
            Call �˳�
        Case mconMenu_File_RecipePar_Help             '����
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
    End Select
End Sub

Private Sub chkIsDosage_Click()
    chkSign.Enabled = chkIsDosageOk.Value = 1 And chkIsDosage.Value = 1
    If chkSign.Enabled = False Then chkSign.Value = 0
    
    lblRefreshComment.Caption = IIf(chkIsDosage.Value = 0, "��������Ϣ�������", "����ҩ������������Ϣ��������Զ�ˢ��")
End Sub

Private Sub chkIsDosageOk_Click()
    chkSign.Enabled = chkIsDosageOk.Value = 1 And chkIsDosage.Value = 1
    If chkSign.Enabled = False Then chkSign.Value = 0
End Sub

Private Sub chkOverTime_Click()
    If chkOverTime.Value = 1 Then
        txtOverTime.Enabled = True
        If Int(Val(txtOverTime.Text)) = 0 Then
            txtOverTime.Text = "30"
        End If
    Else
        txtOverTime.Enabled = False
    End If
End Sub

Private Sub chkUseDisplay_Click()
    If Me.chkUseDisplay.Value = 0 Then
        frm��ʾ�豸����.Enabled = False
    Else
        frm��ʾ�豸����.Enabled = True
    End If
End Sub

Private Sub chkUseSound_Click()
    If Me.chkUseSound.Value = 1 Then
        frm�����㲥����.Enabled = True
        FraԶ����������.Enabled = True
        Me.optCallWay(0).Enabled = True
        Me.optCallWay(1).Enabled = True
    Else
        frm�����㲥����.Enabled = False
        FraԶ����������.Enabled = False
        Me.optCallWay(0).Enabled = False
        Me.optCallWay(1).Enabled = False
    End If
End Sub

Private Sub chk��ҩˢ��_Click()
    lst������.Enabled = (chk��ҩˢ��.Value = 1)
End Sub

Private Sub chk�Զ���ҩ_Click()
    If chk�Զ���ҩ.Value = 1 Then
        txt��ҩʱ��.Enabled = chk�Զ���ҩ.Enabled
    Else
        txt��ҩʱ��.Enabled = False
    End If
End Sub
Private Sub cmdDefaultColor_Click()
    Call GetDefaultRecipeColor
End Sub

Private Sub cmdCheckAll_Click()
    Dim i As Integer
    
    For i = 1 To Lvw��Դ����.ListItems.count
        Lvw��Դ����.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    
    For i = 1 To Lvw��Դ����.ListItems.count
        Lvw��Դ����.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdDefaultPrinter_Click()
    Dim strDefault As String
    Dim n As Integer
    Dim i As Integer
    Dim rsData As ADODB.Recordset
    
    'ȡ����ĸ�ʽ���ƣ�Ĭ��ȡ��һ����ʽ��
    If mstrRPTDefaultScheme_Recipt = "" Then
        Set rsData = DeptSendWork_Get��ҩ����ʽ("ZL1_BILL_1341_3")
        If Not rsData.EOF Then mstrRPTDefaultScheme_Recipt = rsData!��ʽ
    End If
    
    '������ǰ�İ汾�����δӲ�ͬ��λ��ȡֵ
'    If mstrRPTDefaultScheme_Recipt <> "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3\" & mstrRPTDefaultScheme_Recipt, "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3\���и�ʽ", "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
       
    If strDefault = "" Or InStr(1, ";" & mstrPrinters & ";", ";" & strDefault & ";") = 0 Then
        '���Ĭ�ϴ�ӡ��Ϊ�գ����߲��ڱ��ش�ӡ���б���ʱ
        MsgBox "û��������ҩ����ǩ��Ӧ�Ĵ�ӡ�������ڡ�Ʊ��(4)�������ã�", vbInformation, gstrSysName
        sstMain.Tab = 3
        Exit Sub
    Else
        '����Ĭ�ϵĴ�ӡ��
        For n = 0 To vsfPrinter.rows - 1
            vsfPrinter.TextMatrix(n, mintCol��ӡ��) = strDefault
        Next
    End If
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1341)
End Sub

Private Sub cmdTestSound_Click()
    On Error GoTo errHandle
    If optSoundType(1).Value = True Then
        '΢������
        Call zlCall_MsSoundPlay("�롢" & "��־�ܡ�" & "��־�ܡ�" & "����һ�Ŵ���", Val(txtSpeed.Text))
    Else
        'ϵͳ����
        Call zlCall_SystemSoundPlay("�롢" & "��־�ܡ�" & "��־�ܡ�" & "����һ�Ŵ���", Val(txtSpeed.Text))
    End If
    Exit Sub
errHandle:
    Call SaveErrLog
End Sub
Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    Select Case cboƱ������.ListIndex
    Case 0
        '��ҩ����ǩ
        strBill = "ZL1_BILL_1341_3"
    Case 1
        '��ҩ����ǩ
        strBill = "ZL1_BILL_1341_4"
    Case 2
        '������ҩ�嵥
        strBill = "ZL1_BILL_1341_2"
    Case 3
        '������ҩ֪ͨ��
        strBill = "ZL1_BILL_1341_1"
    Case 4
        '���ʴ���ͳ�Ʊ�
        strBill = "ZL1_INSIDE_1341"
    Case 5
        '��ҩҩƷ��ǩ
        strBill = "ZL1_BILL_1341_6"
    Case 6
        '�в�ҩҩƷ��ǩ
        strBill = "ZL1_BILL_1341_7"
    Case 7
        '���˷ѵ���
        strBill = "ZL1_BILL_1341_8"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub cmd��ʾ�豸����_Click()
    If gobjLEDShow Is Nothing Then
        If Not CreateObject_LED(Val(cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListIndex))) Then Exit Sub
    End If
        
    If Not gobjLEDShow Is Nothing Then
        Call gobjLEDShow.zlDrugSetup(Me, mstrWin)
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = pic����ҳǩ.hWnd
        Case 2
            Item.Handle = pic��������.hWnd
    End Select
End Sub

Private Sub lst��ӡ����_GotFocus()
    sstMain.Tab = 2
End Sub

Private Sub InitCommandBar()
    '��ʼ��������
    'CommandBar
    '-----------------------------------------------------
    Dim cbrToolBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003
    
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
    cbsMain.Icons = imgFunc.Icons                                  '���ù�����ͼ��ؼ�

    '���ز˵�
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False                          '���ز˵�
    
    '���ء��ļ����˵�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)")
    objMenu.Id = 1                                                  'Popup��ID�����¸�ֵ������Ч
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_File_RecipePar_Save, "����(&S)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_RecipePar_Cancel, "�˳�(&E)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_RecipePar_Help, "����(&H)")
    End With
    '���ء��ļ�����ť
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_File_RecipePar_Save, "����")
        Set objControl = .Add(xtpControlButton, mconMenu_File_RecipePar_Cancel, "�˳�")
        Set objControl = .Add(xtpControlButton, mconMenu_File_RecipePar_Help, "����")
        objControl.BeginGroup = True
    End With
    
    For Each objControl In cbrToolBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitPanes()
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Dim objPaneList As Pane
    Dim objPaneParams As Pane
    
    Set objPaneList = Me.dkpMain.CreatePane(1, 25, 100, DockLeftOf, Nothing)
    objPaneList.Title = "����ҳǩ"
    objPaneList.Options = PaneNoCaption
    objPaneList.MaxTrackSize.SetSize 200, 100
    
    Set objPaneParams = Me.dkpMain.CreatePane(2, 100, 100, DockRightOf, objPaneList)
    objPaneParams.Title = "��������"
    objPaneParams.Options = PaneNoCaption
    
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.UseSplitterTracker = False   'ʵʱ�϶�
    Me.dkpMain.Options.AlphaDockingContext = True
    Me.dkpMain.Options.CloseGroupOnButtonClick = False
    Me.dkpMain.Options.HideClient = True
    Me.dkpMain.Options.LunaColors = True
    Me.dkpMain.Options.LockSplitters = True        '�����϶�
    Me.dkpMain.PaintManager.DrawSingleTab = False
    Me.dkpMain.TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
End Sub

Private Sub InitTPLItem()
    '����:��ʼ���������ؼ�
    
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    Dim strTabTitle As String   '����ҳǩ�ı��⴮
    Dim i As Integer
        
    '���ӷ���
    Set tplGroup = tplFunc.Groups.Add(1, "��������")
    tplGroup.CaptionVisible = False       '�Ƿ���ʾ����
    tplGroup.Expanded = True            '��ʼ��ʱ�Ƿ���ʾ�ӽڵ�
        
    tplFunc.SetMargins 8, 8, 8, 8, 0    '���߷�Χ
    tplFunc.SetIconSize 24, 24
    
    '�����ӽڵ�
    For i = 1 To sstMain.Tabs
        If sstMain.TabVisible(i - 1) Then       '�ų���δ���ŵĹ���
            Set tplItem = tplGroup.Items.Add(i, sstMain.TabCaption(i - 1), xtpTaskItemTypeLink, i)
            tplItem.IconIndex = i   '������һ�д����ͼ�긳ֵδ�ɹ��������¸�ֵ
        End If
    Next
    
End Sub

Private Sub Cboҩ��_Click()
    Dim intDO As Integer
    Dim bln���� As Boolean, blnסԺ As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If BlnStartUp = False Then Exit Sub
    '�����ܣ����û������ҩ���������涼������
    If Me.Cboҩ��.ListCount = 0 Then Exit Sub
    
    Call ReadWindowsAndPeople
    intUnit = Val(zlDatabase.GetPara("ҩ������", glngSys, 1341))
    
    '������ҩ����
    SetDispense
    
    '����ҩ����ʾ��λ
    gstrSQL = " Select distinct ������� From ��������˵��" & _
              " Where ����ID=[1] And �������� like '%ҩ��'" & _
              " Order By ������� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡҩ���������]", Cboҩ��.ItemData(Cboҩ��.ListIndex))
    
    rsTemp.Filter = "�������=3"
    If rsTemp.RecordCount <> 0 Then bln���� = True: blnסԺ = True
    rsTemp.Filter = "�������=2"
    If rsTemp.RecordCount <> 0 Then blnסԺ = True
    rsTemp.Filter = "�������=1"
    If rsTemp.RecordCount <> 0 Then bln���� = True
    rsTemp.Filter = 0
    
    With cbo��λ
        .Clear
        .AddItem "1-����Ӧ"
        .ItemData(.NewIndex) = 0
        If bln���� Then
            .AddItem "2-����ҩ��"
            .ItemData(.NewIndex) = 1
        End If
        If blnסԺ Then
            .AddItem "3-סԺҩ��"
            .ItemData(.NewIndex) = 2
        End If
        .ListIndex = 0
        
        For intDO = 0 To .ListCount - 1
            If .ItemData(intDO) = intUnit Then
                .ListIndex = intDO
                Exit For
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk��ӡ��ҩ��_Click()
    Dim ConState As Boolean
    
    ConState = (Chk��ӡ��ҩ��.Value = 1 And Chk��ӡ��ҩ��.Enabled = True)
    Opt��ӡ��ҩ��������.Enabled = ConState
    Opt��ӡ��ҩ��������.Enabled = ConState
    Opt��ӡ��ҩ��ѡ��.Enabled = ConState
    If Not ConState Then lst��ӡ����.Enabled = False
    
    If BlnStartUp = False Then Exit Sub
    
    If ConState Then
        If Opt��ӡ��ҩ��������.Enabled = True Then Opt��ӡ��ҩ��������.SetFocus
    End If
End Sub

Private Sub Chk��ӡ��ҩ��_GotFocus()
    sstMain.Tab = 2
End Sub

Private Sub ����()
    Dim IntPrintStyle As Integer, i As Integer
    Dim strWin1 As String, strWin2 As String
    Dim strColor As String
    Dim intTemp As Integer
    Dim n As Integer
    Dim strPrinters As String
    Dim intSendCount As Integer
    Dim strCardType As String
    Dim str���� As String
    
    If Trim(txt��ѯ����.Text) = "" Then
        txt��ѯ����.Text = "1"
    End If
    If Not IsNumeric(txt��ѯ����.Text) Then
        MsgBox "��ѯ�����к��зǷ��ַ���", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txt��ѯ����.Text) < 1 Or Val(txt��ѯ����.Text) > 365 Then
        MsgBox "��ѯ��������С��1������365�죡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Trim(Txtˢ�¼��) <> "" Then
        If Not IsNumeric(Txtˢ�¼��) Then
            MsgBox "ˢ�¼���к��зǷ��ַ���", vbInformation, gstrSysName
            Exit Sub
        End If
        If Val(Txtˢ�¼��) < 0 Or Val(Txtˢ�¼��) > 60 Then
            MsgBox "ˢ�¼��ֵ������Χ��0��60����", vbInformation, gstrSysName
            Exit Sub
        End If
        Txtˢ�¼�� = CInt(Txtˢ�¼��)
    End If
    If Trim(txt��ӡ���) <> "" Then
        If Not IsNumeric(txt��ӡ���) Then
            MsgBox "��ӡ����к��зǷ��ַ���", vbInformation, gstrSysName
            Exit Sub
        End If
        If Val(txt��ӡ���) < 0 Or Val(txt��ӡ���) > 60 Then
            MsgBox "��ӡ���ֵ������Χ��0��60����", vbInformation, gstrSysName
            Exit Sub
        End If
        txt��ӡ��� = CInt(txt��ӡ���)
    End If
    If Trim(Txt�ӳٴ�ӡ) <> "" Then
        If Not IsNumeric(Txt�ӳٴ�ӡ) Then
            MsgBox "�ӳٴ�ӡ�к��зǷ��ַ���", vbInformation, gstrSysName
            Exit Sub
        End If
        If Val(Txt�ӳٴ�ӡ) < 0 Or Val(Txt�ӳٴ�ӡ) > 60 Then
            MsgBox "�ӳٴ�ӡֵ������Χ��0��60����", vbInformation, gstrSysName
            Exit Sub
        End If
        Txt�ӳٴ�ӡ = CInt(Txt�ӳٴ�ӡ)
    End If
    If Trim(Txt��ӡ�˷ѵ���) <> "" Then
        If Not IsNumeric(Txt��ӡ�˷ѵ���) Then
            MsgBox "�˷ѵ����к��зǷ��ַ���", vbInformation, gstrSysName
            Exit Sub
        End If
        If Val(Txt��ӡ�˷ѵ���) < 0 Or Val(Txt��ӡ�˷ѵ���) > 60 Then
            MsgBox "��ӡ�˷ѵ�ֵ������Χ��0��60����", vbInformation, gstrSysName
            Exit Sub
        End If
        Txt��ӡ�˷ѵ��� = CInt(Txt��ӡ�˷ѵ���)
    End If
    
    '��鱾�����ܴ���:�����,����Ҫѡ��һ��
    For i = 0 To lst��ҩ����.ListCount - 1
        If lst��ҩ����.Selected(i) Then
            strWin1 = strWin1 & ",'" & lst��ҩ����.List(i) & "'"
            intSendCount = intSendCount + 1
        End If
    Next
    
    '��������Ŷӽкţ��򱾻�ֻ������һ����ҩ����
    If intSendCount > 1 And chk�����Ŷӽк�.Value = 1 Then
        MsgBox "�������Ŷӽкţ�ֻ������һ����ҩ���ڣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnLoadDrug And intSendCount > 1 Then
        MsgBox "�����������Զ���ҩ��ֻ������һ����ҩ���ڣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strWin1 = Mid(strWin1, 2)
    If strWin1 = "" And lst��ҩ����.ListCount > 0 Then
        MsgBox "��ָ��������վ����Ӧ�ķ�ҩ���ڡ�", vbInformation, gstrSysName
        Exit Sub
    End If
'    If UBound(Split(strWin1, ",")) + 1 = lst��ҩ����.ListCount Then strWin1 = ""
       
    
    '����ӡ��ҩ����:�����Ƿ���,����Ҫѡ��һ��
    For i = 0 To lst��ӡ����.ListCount - 1
        If lst��ӡ����.Selected(i) Then
            strWin2 = strWin2 & ",'" & lst��ӡ����.List(i) & "'"
        End If
    Next
    strWin2 = Mid(strWin2, 2)
    If strWin2 = "" And Chk��ӡ��ҩ��.Value = 1 And Opt��ӡ��ҩ��ѡ��.Value Then
        MsgBox "ѡ���ӡָ�����ڵ���ҩ��ʱ����Ҫ���ö�Ӧ�ķ�ҩ���ڣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If UBound(Split(strWin2, ",")) + 1 = lst��ӡ����.ListCount Then strWin2 = ""
    
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
    
    With vsfPrinter
        intTemp = .rows / 6
        
        '������ɫ
        For n = 1 To 6
            strColor = IIf(strColor = "", "", strColor & ";") & CStr(.Cell(flexcpBackColor, (n - 1) * intTemp, mIntCol����))
        Next
        
        '������Ӧ�Ĵ�ӡ��������ʱ�á�;���̶���Ϊ6�ֲ�ͬ�Ĵ������ͣ���ʽ1?��ӡ��,��ʽ2?��ӡ��;��ʽ2?��ӡ��,��ʽ2?��ӡ��...
        For n = 0 To .rows - 1
            If str���� <> .TextMatrix(n, mIntCol����) Then
                If str���� = "" Then
                    strPrinters = .TextMatrix(n, mintCol��ʽ) & "?" & .TextMatrix(n, mintCol��ӡ��)
                Else
                    strPrinters = strPrinters & ";" & .TextMatrix(n, mintCol��ʽ) & "?" & .TextMatrix(n, mintCol��ӡ��)
                End If
                str���� = .TextMatrix(n, mIntCol����)
            Else
                strPrinters = strPrinters & "," & .TextMatrix(n, mintCol��ʽ) & "?" & .TextMatrix(n, mintCol��ӡ��)
            End If
        Next
    End With
    
    '����ˢ���Ŀ����
    If chk��ҩˢ��.Value = 1 Then
        If lst������.ListCount > 0 Then
            For i = 0 To lst������.ListCount - 1
                If lst������.Selected(i) Then
                    strCardType = IIf(strCardType = "", strCardType, strCardType & ",") & lst������.ItemData(i)
                End If
            Next
        End If
    End If
        
    On Error Resume Next
    
    '���湫����˽�в���
    zlDatabase.SetPara "������", Me.cboҩƷ������ʾ.ListIndex, glngSys, 1341
    zlDatabase.SetPara "������ɫ", strColor, glngSys, 1341
    zlDatabase.SetPara "У�鷢ҩ��", chkУ�鷢ҩ��.Value, glngSys, 1341
    zlDatabase.SetPara "У�鷽ʽ", IIf(Opt��֤��ʽ(0).Value, 0, 1), glngSys, 1341
    zlDatabase.SetPara "У����ҩ��", chkУ����ҩ��.Value, glngSys, 1341
    zlDatabase.SetPara "�Զ�����", chk�Զ�����.Value, glngSys, 1341

    zlDatabase.SetPara "�շѴ�����ʾ��ʽ", cbo�շѴ���.ListIndex, glngSys, 1341
    zlDatabase.SetPara "���ʴ�����ʾ��ʽ", cbo���ʴ���.ListIndex, glngSys, 1341
    zlDatabase.SetPara "����ҩ���ݴ�ӡ��ʾ��ʽ", cbo����ҩ.ListIndex, glngSys, 1341
    zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, 1341
    zlDatabase.SetPara "��ӡ�������ʵ�", IIf(chk���ʵ�.Value, 1, 0), glngSys, 1341
    zlDatabase.SetPara "��ӡ�˷ѵ��ݼ��", Val(Txt��ӡ�˷ѵ���), glngSys, 1341
    zlDatabase.SetPara "��ӡ�ӳ�", Val(Txt�ӳٴ�ӡ), glngSys, 1341
    zlDatabase.SetPara "ˢ�¼��", Val(Txtˢ�¼��), glngSys, 1341
    zlDatabase.SetPara "��ӡ���", Val(txt��ӡ���), glngSys, 1341
    
    zlDatabase.SetPara "ҩ������", cbo��λ.ListIndex, glngSys, 1341
    zlDatabase.SetPara "��ʾ����", Chk��ʾ������.Value, glngSys, 1341
    zlDatabase.SetPara "��ҩ���Զ���ӡ", Me.Cbo��ҩ��.ListIndex, glngSys, 1341
    zlDatabase.SetPara "��ҩ���Զ���ӡ", Me.cbo��ҩ��.ListIndex, glngSys, 1341
    zlDatabase.SetPara "��ҩ���ӡҩƷ��ǩ", Me.cboҩƷ��ǩ.ListIndex, glngSys, 1341
    zlDatabase.SetPara "��ʾ��С��λ", chk��С��λ.Value, glngSys, 1341
    zlDatabase.SetPara "��ҩ��ˢ����֤", chkˢ��.Value, glngSys, 1341
    zlDatabase.SetPara "��ҩģʽɨ����ȷ��", chk��ҩɨ��.Value, glngSys, 1341
    zlDatabase.SetPara "��ʱδ��ҩƷ��ʾʱ����", IIf(chkOverTime.Value = 0, 0, Int(Val(txtOverTime.Text))), glngSys, 1341
    zlDatabase.SetPara "������סԺ����", Me.cbo��������.ListIndex, glngSys, 1341
    zlDatabase.SetPara "����ˢ����ҩ", strCardType, glngSys, 1341
    zlDatabase.SetPara "ҩƷҽ��������ʱ�����", chk����ʱ��.Value, glngSys, 1341
    zlDatabase.SetPara "�����ʾ��ʽ", cbo�����ʾ.ListIndex, glngSys, 1341
    zlDatabase.SetPara "���ò���ʵ��ȡҩȷ��ģʽ", chkTakeDrug.Value, glngSys, 1341
    zlDatabase.SetPara "һ��ͨ�շ��뷢ҩ����", chksend.Value, glngSys, 1341
    zlDatabase.SetPara "����ҩ����ɨ����Զ�����", chkɨ������.Value, glngSys, 1341
    zlDatabase.SetPara "����ʱϵͳ�Զ��س���ʽ", Me.cbo�س���ʽ.ListIndex, glngSys, 1341
    zlDatabase.SetPara "�Զ���ҩ����", Me.cbo�Զ���ҩ����.ListIndex, glngSys, 1341
    zlDatabase.SetPara "��ҩ���θ���", chk��ҩ���θ���.Value, glngSys, 1341
    zlDatabase.SetPara "��ҩ�������ķ������", chkCheckStuff.Value, glngSys, 1341
    
    If chkDispensing.Visible Then
        zlDatabase.SetPara "����ʱ֪ͨ��ʼ��ҩ", Me.chkDispensing.Value, glngSys, 1341
    Else
        zlDatabase.SetPara "����ʱ֪ͨ��ʼ��ҩ", "0", glngSys, 1341
    End If
    
    '��ӡ
    IntPrintStyle = Chk��ӡ��ҩ��.Value
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt��ӡ��ҩ��������.Value, 1, 1)
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt��ӡ��ҩ��������.Value, 2, 1)
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt��ӡ��ҩ��ѡ��.Value, 3, 1)
    zlDatabase.SetPara "�����µ����Ƿ��ӡ", IntPrintStyle, glngSys, 1341
    zlDatabase.SetPara "��ӡָ����ҩ����", strWin2, glngSys, 1341
    zlDatabase.SetPara "��ӡҩƷ��ǩ", IIf(chkҩƷ��ǩ.Value, 1, 0), glngSys, 1341
            
    '��ҩ
    zlDatabase.SetPara "��ҩҩ��", Cboҩ��.ItemData(Cboҩ��.ListIndex), glngSys, 1341
    zlDatabase.SetPara "��ҩ����", strWin1, glngSys, 1341
    zlDatabase.SetPara "��ҩ��", IIf(Cbo��ҩ��.Text <> "��ǰ����Ա", Cbo��ҩ��.Text, "|��ǰ����Ա|"), glngSys, 1341
    zlDatabase.SetPara "�˲���", IIf(cboCheck.Text <> "��ǰ����Ա", cboCheck.Text, "|��ǰ����Ա|"), glngSys, 1341
    zlDatabase.SetPara "�Զ���ҩ", IIf(chk�Զ���ҩ.Value = 1, 1, 0), glngSys, 1341
    zlDatabase.SetPara "�Զ���ҩʱ��", Val(txt��ҩʱ��.Text), glngSys, 1341
    zlDatabase.SetPara "��ӡƱ�ݵ����и�ʽ", IIf(chkAllType.Value = 1, 1, 0), glngSys, 1341
    zlDatabase.SetPara "����˲��˺���ҩ����ͬ", IIf(chkSame.Value = 1, 1, 0), glngSys, 1341
    zlDatabase.SetPara "��ӡ����ǩʱ��Ԥ���ٴ�ӡ", chkPreview.Value, glngSys, 1341
    
    
    '�����ŶӽкŵĲ���
    zlDatabase.SetPara "�кŷ�ʽ", IIf(Me.optCallWay(0).Value = True, 0, 1), glngSys, 1341
    zlDatabase.SetPara "�����Ŷӽк�", Me.chk�����Ŷӽк�.Value, glngSys, 1341
    zlDatabase.SetPara "������������", Me.chkUseSound.Value, glngSys, 1341
    zlDatabase.SetPara "��ʾ�ŶӶ���", chkUseDisplay.Value, glngSys, 1341
    zlDatabase.SetPara "�������Ŵ���", Val(txtPlayCount.Text), glngSys, 1341
    zlDatabase.SetPara "�����㲥ʱ�䳤��", Val(txt�㲥ʱ�䳤��.Text), glngSys, 1341
    zlDatabase.SetPara "�����㲥����", Val(txtSpeed.Text), glngSys, 1341
    zlDatabase.SetPara "��������", IIf(optSoundType(0).Value = True, 0, 1), glngSys, 1341
    zlDatabase.SetPara "Զ�˺���վ��", Me.cboWorkStation.Text, glngSys, 1341
    zlDatabase.SetPara "������ѯʱ��", Val(Me.txtLoopQueryTime.Text), glngSys, 1341
    zlDatabase.SetPara "��ʾ�豸���", cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListIndex), glngSys, 1341
    zlDatabase.SetPara "ǩ��ʱ������ҩ", chkSign.Value, glngSys, 1341
    
    '��Դ����
    zlDatabase.SetPara "��Դ����", mstrSourceDep, glngSys, 1341
    
    '��ҩ��&����ǩ��ӡ��ʽ
    zlDatabase.SetPara "��ҩ����ӡ��ʽ", cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.ListIndex) & ";" & cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.ListIndex), glngSys, 1341
    zlDatabase.SetPara "����ǩ��ӡ��ʽ", cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.ListIndex) & ";" & cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.ListIndex), glngSys, 1341
    
    '������Ӧ�Ĵ�ӡ��
    zlDatabase.SetPara "������Ӧ�Ĵ�ӡ��", strPrinters, glngSys, 1341
    
    frmҩƷ������ҩNew.BlnSetParaSuccess = True
    
    '������ҩ����ҩȷ�ϻ���
    gstrSQL = "Zl_ҩ����ҩ����_Update("
    gstrSQL = gstrSQL & Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex)
    gstrSQL = gstrSQL & "," & Me.chkIsDosage.Value
    gstrSQL = gstrSQL & "," & Me.chkIsDosageOk.Value
    gstrSQL = gstrSQL & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, "cmdOK_Click")
    Unload Me
    Exit Sub
End Sub

Private Sub �˳�()
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim objPic As PictureBox

    BlnStartUp = False
    
    For Each objPic In picPar
        Set objPic.Container = pic��������          '��������
        objPic.Visible = False                      '��ʼ�ر�������ʾ
    Next
    
    '��ʼ�����岼��
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)      '��ͼ��ؼ������ͼƬ���뵽taskpanel����
    Call InitCommandBar
    Call InitPanes
    Call InitTPLItem
    
    '��ʼ��chkDispensing
    Call InitDispensing
    
    picPar(0).Visible = True          'Ĭ����ʾ��һ��ҳ��
    sstMain.Visible = False
    
    For Each objPic In picPar
        objPic.BackColor = &H8000000F
    Next
    
    '��ȡע���
    Call ReadFromReg
    '����������ʾ
    Call WriteCons
    '��Դ����
    Call SetSourceDep
    
    BlnStartUp = True
    RestoreWinState Me, App.ProductName
End Sub

Private Function ReadWindowsAndPeople()
    '--��ȡ��ҩ���ķ�ҩ���ڼ���ҩ��--
    
    
        '��ҩ���ڣ�Ҫ��ӡ�ķ�ҩ�����������в�����"���з�ҩ����"��
'        If .State = 1 Then .Close
'        gstrSQL = " Select ���� From ��ҩ���� Where ҩ��ID=" & Cboҩ��.ItemData(Cboҩ��.ListIndex)
'        Call SQLTest(App.Title, Me.Caption, gstrSQL)
'        .Open gstrSQL, gcnOracle
'        Call SQLTest

    Dim lngLEDModal As Long
    
    On Error GoTo errHandle
    gstrSQL = " Select ���� From ��ҩ���� Where ҩ��ID=[1]"
    Set RecPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Cboҩ��.ItemData(Cboҩ��.ListIndex))
    
    mstrWin = ""
    
    With RecPeople
        Me.lst��ҩ����.Clear
        Me.lst��ӡ����.Clear

        Do While Not .EOF
            lst��ҩ����.AddItem !����
            lst��ӡ����.AddItem !����
            
            lst��ҩ����.Selected(lst��ҩ����.NewIndex) = True
            If Opt��ӡ��ҩ��ѡ��.Value Then
                lst��ӡ����.Selected(lst��ӡ����.NewIndex) = True
            End If
            
            mstrWin = IIf(mstrWin = "", "", mstrWin & ",") & !����
            
            .MoveNext
        Loop

        If lst��ҩ����.ListCount > 0 Then lst��ҩ����.ListIndex = 0
        If lst��ӡ����.ListCount > 0 Then lst��ӡ����.ListIndex = 0
    End With
    '��ҩ��
    gstrSQL = " Select ���� From ��Ա��  Where ID in " & _
             " (Select Distinct ��ԱID From ��Ա����˵�� Where ��Ա����='ҩ����ҩ��' " & _
             " And ��ԱID IN (Select ��ԱID From ������Ա Where ����ID=[1]))" & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
    Set RecPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Cboҩ��.ItemData(Cboҩ��.ListIndex))
    
    With RecPeople
        Me.Cbo��ҩ��.Clear
        Me.Cbo��ҩ��.AddItem "��ǰ����Ա"
        Do While Not .EOF
            Cbo��ҩ��.AddItem !����
            .MoveNext
        Loop
        Cbo��ҩ��.ListIndex = 0
    End With
    
    With RecPeople
        If .RecordCount <> 0 Then
            .MoveFirst
        End If
        Me.cboCheck.Clear
        Me.cboCheck.AddItem "��ǰ����Ա"
        Do While Not .EOF
            cboCheck.AddItem !����
            .MoveNext
        Loop
        cboCheck.ListIndex = 0
    End With
    
    
    
    lngLEDModal = zlDatabase.GetPara("��ʾ�豸���", glngSys, 1341, "101")
    cbo��ʾӲ�����.Clear
    
    gstrSQL = "Select ��������,������,Nvl(����,0) AS ����,˵�� From �Ŷ�LED��ʾ����  "
    Set RecPeople = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��LED��ʾ�ӿڵ�ע����Ϣ")
    
    While RecPeople.EOF = False
        cbo��ʾӲ�����.AddItem NVL(RecPeople!˵��)
        cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListCount - 1) = NVL(RecPeople!��������, 0)
        If lngLEDModal = NVL(RecPeople!��������, 0) Then
            cbo��ʾӲ�����.ListIndex = cbo��ʾӲ�����.ListCount - 1
        End If
        RecPeople.MoveNext
    Wend
    
    If cbo��ʾӲ�����.ListCount > 0 And cbo��ʾӲ�����.ListIndex = -1 Then
        cbo��ʾӲ�����.ListIndex = 0
    End If
    
    '���վ���б�
    ReadWorkStationInf
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteCons()
    Dim IntLocate As Integer
    Dim rsData As ADODB.Recordset
    
    '�����û�������ʾ
    
    RecPart.MoveFirst               '������Ϊ�գ������������涼�����ˣ�
    
    txt��ѯ����.Text = intDays
    'װ������������
    With Me.Cboҩ��
        Do While Not RecPart.EOF
            .AddItem RecPart!����
            .ItemData(.NewIndex) = RecPart!Id
            RecPart.MoveNext
        Loop
        .ListIndex = 0
    End With
    With Me.Cbo��ҩ��
        .AddItem "1-��ҩ����ʾ�Ƿ��ӡ"
        .AddItem "2-��ҩ���Զ���ӡ"
        .AddItem "3-��ҩ�󲻴�ӡ"
        .ListIndex = IntAutoPrint
    End With
    
    With Me.cbo��ҩ��
        .AddItem "1-��ҩ����ʾ�Ƿ��ӡ"
        .AddItem "2-��ҩ���Զ���ӡ"
        .AddItem "3-��ҩ�󲻴�ӡ"
        .ListIndex = mint��ҩ���Զ���ӡ
    End With
    
    With Me.cboҩƷ��ǩ
        .AddItem "1-��ҩ����ʾ�Ƿ��ӡ"
        .AddItem "2-��ҩ���Զ���ӡ"
        .AddItem "3-��ҩ�󲻴�ӡ"
        .ListIndex = mint��ҩ���Զ���ӡҩƷ��ǩ
    End With
    
    With cbo�շѴ���
        .Clear
        .AddItem "1-����ʾ�κδ���"
        .AddItem "2-��ʾδ�շѴ���"
        .AddItem "3-��ʾ���շѴ���"
        .AddItem "4-��ʾ���еĴ���"
        .ListIndex = 0
    End With
    With cbo���ʴ���
        .Clear
        .AddItem "1-����ʾ�κδ���"
        .AddItem "2-��ʾδ��˴���"
        .AddItem "3-��ʾ����˴���"
        .AddItem "4-��ʾ���еĴ���"
        .ListIndex = 0
    End With
    
    With cbo����ҩ
        .Clear
        .AddItem "0-��ʾ������ҩ��"
        .AddItem "1-��ʾδ��ӡ��ҩ��"
        .AddItem "2-��ʾ�Ѵ�ӡ��ҩ��"
        .ListIndex = 0
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-��ҩ����ǩ"
        .AddItem "2-��ҩ����ǩ"
        .AddItem "3-������ҩ�嵥"
        .AddItem "4-������ҩ֪ͨ��"
        .AddItem "5-���ʴ���ͳ�Ʊ�"
        .AddItem "6_��ҩҩƷ��ǩ"
        .AddItem "7_�в�ҩҩƷ��ǩ"
        .AddItem "8_���˷ѵ���"
        .ListIndex = 0
    End With
    
    With Me.cboҩƷ������ʾ
        .Clear
        .AddItem "0-��ʾҩƷ����������"
        .AddItem "1-����ʾҩƷ����"
        .AddItem "2-����ʾҩƷ����"
        .ListIndex = 0
    End With
    
    With Me.cbo�����ʾ
        .Clear
        .AddItem "0-��ʾӦ�ս��"
        .AddItem "1-��ʾʵ�ս��"
        .AddItem "2-��ʾӦ�պ�ʵ�ս��"
        .ListIndex = 0
    End With
    
    With Me.cbo��������
        .Clear
        .AddItem "0-��ʾ�����סԺ����"
        .AddItem "1-ֻ��ʾ���ﴦ��"
        .AddItem "2-ֻ��ʾסԺ����"
        .ListIndex = mintType
    End With
    
    With Me.cbo�س���ʽ
        .Clear
        .AddItem "0-ϵͳ���Զ��س�"
        .AddItem "1-��¼��ﵽ��Ŀ�򿨺ų���ʱ�Զ��س�"
    End With
    
    With Me.cbo�Զ���ҩ����
        .Clear
        .AddItem "0-ȫ�������Զ���ҩ"
        .AddItem "1-���Ӵ���(��ҽ���Ĵ���)�Զ���ҩ"
        .AddItem "2-�ֹ�����(��ҽ���Ĵ���)�Զ���ҩ"
    End With
    
    'װ���������
    cbo�շѴ���.ListIndex = mintShowBill�շ�
    cbo���ʴ���.ListIndex = mintShowBill����
    cbo����ҩ.ListIndex = mintShowBill��ҩ
    If intУ�鷽ʽ = 0 Then
        Opt��֤��ʽ(0).Value = True
    Else
        Opt��֤��ʽ(1).Value = True
    End If
    chkУ����ҩ��.Value = intУ����ҩ��
    chkУ�鷢ҩ��.Value = intУ�鷢ҩ��
    Chk��ʾ������.Value = IntShowCol
    
    Chk��ӡ��ҩ��.Value = IIf(intPrint = 0, 0, 1)
    
    cbo�����ʾ.ListIndex = mint�����ʾ��ʽ

    Opt��ӡ��ҩ��������.Value = IIf(intPrint = 1, True, False)
    Opt��ӡ��ҩ��������.Value = IIf(intPrint = 2, True, False)
    Opt��ӡ��ҩ��ѡ��.Value = IIf(intPrint = 3, True, False)
    
    Txtˢ�¼�� = Format(IntRefresh, "#####;-#####; ;")
    txt��ӡ��� = Format(mintPrintInterval, "#####;-#####; ;")
    Txt�ӳٴ�ӡ = Format(intPrintDelay, "#####;-#####; ;")
    Txt��ӡ�˷ѵ��� = Format(intPrintHandbackNO, "#####;-#####; ;")
    
    If txt��ӡ���.Enabled = True Then txt��ӡ���.Enabled = Not mblnUseMsg
    lblPrintComment.Visible = mblnUseMsg
    
    If Txtˢ�¼��.Enabled = True Then Txtˢ�¼��.Enabled = Not mblnUseMsg And chkIsDosage.Value = 0
    lblRefreshComment.Visible = mblnUseMsg
    lblRefreshComment.Caption = IIf(chkIsDosage.Value = 0, "��������Ϣ�������", "����ҩ������������Ϣ��������Զ�ˢ��")
    
    If lngҩ��ID <> 0 Then                                  '��λҩ��
        '�����ڸ�ҩ������ʾ
        For IntLocate = 0 To Me.Cboҩ��.ListCount - 1
            If Me.Cboҩ��.ItemData(IntLocate) = lngҩ��ID Then
                Me.Cboҩ��.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (Cboҩ��.ListCount - 1) Then
            MsgBox "����������ҩ����ԭ�����õ�ҩ����ʧЧ����", vbInformation, gstrSysName
            If Cboҩ��.ListCount >= 1 Then Cboҩ��.ListIndex = 0
        End If
    End If
    BlnStartUp = True
    Cboҩ��_Click                                           '��������ҩ���񣬾���ȡ��ҩ��������ҩ���ڼ���ҩ��
    BlnStartUp = False
    
    '��λ��ҩ����
    If Str���� <> "" Then
        For IntLocate = 0 To lst��ҩ����.ListCount - 1
            If InStr(Str����, "'" & lst��ҩ����.List(IntLocate) & "'") > 0 Then
                lst��ҩ����.Selected(IntLocate) = True
            Else
                lst��ҩ����.Selected(IntLocate) = False
            End If
        Next
        If lst��ҩ����.ListCount > 0 Then lst��ҩ����.ListIndex = 0
    End If
    
    If str��ҩ�� <> "" Then                                 '��ʾ
        '�����ڸ���ҩ������ʾ
        If str��ҩ�� = "|��ǰ����Ա|" Then
            Cbo��ҩ��.ListIndex = 0
        Else
            For IntLocate = 1 To Cbo��ҩ��.ListCount - 1
                If Cbo��ҩ��.List(IntLocate) = str��ҩ�� Then
                    Cbo��ҩ��.ListIndex = IntLocate
                    Exit For
                End If
            Next
            If IntLocate > (Cbo��ҩ��.ListCount - 1) Then
                MsgBox "������������ҩ�ˣ�ԭ�����õ���ҩ���Ѳ��ڱ����ţ���", vbInformation, gstrSysName
                If Cbo��ҩ��.ListCount >= 1 Then Cbo��ҩ��.ListIndex = 0
            End If
        End If
    End If
    
    If mstr�˲��� <> "" Then
        '�����ڸú˲�������ʾ
        If mstr�˲��� = "|��ǰ����Ա|" Then
            cboCheck.ListIndex = 0
        Else
            For IntLocate = 1 To cboCheck.ListCount - 1
                If cboCheck.List(IntLocate) = mstr�˲��� Then
                    cboCheck.ListIndex = IntLocate
                    Exit For
                End If
            Next
            If IntLocate > (cboCheck.ListCount - 1) Then
                MsgBox "���������ú˲��ˣ�ԭ�����õĺ˲����Ѳ��ڱ����ţ���", vbInformation, gstrSysName
                If cboCheck.ListCount >= 1 Then cboCheck.ListIndex = 0
            End If
        End If
    End If
    
    '��λ��ӡ��ҩ����
    If strPrintWindow <> "" Then
        For IntLocate = 0 To lst��ӡ����.ListCount - 1
            If InStr(strPrintWindow, "'" & lst��ӡ����.List(IntLocate) & "'") > 0 Then
                lst��ӡ����.Selected(IntLocate) = True
            Else
                lst��ӡ����.Selected(IntLocate) = False
            End If
        Next
        If lst��ӡ����.ListCount > 0 Then lst��ӡ����.ListIndex = 0
    End If
    
    Me.cboҩƷ������ʾ.ListIndex = mintShowName
    
    chk�Զ���ҩ.Value = IIf(mint�Զ���ҩ = 1, 1, 0)
    chk���ʵ�.Value = IIf(mint���ʵ� = 1, 1, 0)
    chkҩƷ��ǩ.Value = IIf(mintҩƷ��ǩ = 1, 1, 0)
    txt��ҩʱ��.Text = mint�Զ���ҩʱ��
    txt��ҩʱ��.Enabled = (mint�Զ���ҩ = 1 And chk�Զ���ҩ.Enabled = True)
    chkˢ��.Value = IIf(mintˢ��֤ = 1, 1, 0)
    chk��ҩɨ��.Value = IIf(mint��ҩɨ�� = 1, 1, 0)
    chkSign.Value = IIf(mintSign = 1, 1, 0)
    Me.chksend.Value = IIf(mintһ��ͨ��ҩģʽ = 1, 1, 0)
    Me.chkɨ������.Value = IIf(mintɨ������ = 1, 1, 0)
    chk��ҩ���θ���.Value = IIf(mint��ҩ���θ��� = 1, 1, 0)
    
    If mint�س���ʽ >= 0 And mint�س���ʽ <= 1 Then
        cbo�س���ʽ.ListIndex = mint�س���ʽ
    Else
        cbo�س���ʽ.ListIndex = 0
    End If
    
    Me.cbo�Զ���ҩ����.ListIndex = mint�Զ���ҩ����
    
    '�����ŶӽкŵĲ���
    With mType_Call
        chk�����Ŷӽк�.Value = .int�����Ŷӽк�
        chkUseDisplay.Value = .int��ʾ�ŶӶ���
        chkUseSound.Value = .int������������
        
        If .int�кŷ�ʽ = 0 Then
            optCallWay(0).Value = True
        Else
            optCallWay(1).Value = True
        End If
        
        optSoundType(.int��������).Value = 1
        txtSpeed.Text = .int�����㲥����
        txt�㲥ʱ�䳤��.Text = .int�����㲥ʱ�䳤��
        txtPlayCount.Text = .int�������Ŵ���
        Me.cboWorkStation.Text = .strԶ�˺���վ��
        txtLoopQueryTime.Text = .int��ѯʱ��
    End With
    
    chkUseDisplay_Click
    chkUseSound_Click
    
    If Me.optCallWay(0).Value = True Then
        optCallWay_Click 0
    Else
        optCallWay_Click 1
    End If
    
    '����ˢ��ģʽ�Ϳ����
    chk��ҩˢ��.Value = IIf(mstr����ˢ����ҩ = "", 0, 1)
    lst������.Enabled = (chk��ҩˢ��.Value = 1)
    
    gstrSQL = "Select ID, ����, ���� From ҽ�ƿ���� Order By ����"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "WriteCons")
    If rsData.RecordCount > 0 Then
        lst������.Clear

        Do While Not rsData.EOF
            lst������.AddItem rsData!����
            lst������.ItemData(lst������.NewIndex) = rsData!Id
            
            If mstr����ˢ����ҩ <> "" Then
                If InStr(1, "," & mstr����ˢ����ҩ & ",", "," & rsData!Id & ",") > 0 Then
                    lst������.Selected(lst������.NewIndex) = True
                End If
            End If
            
            rsData.MoveNext
        Loop

        If lst������.ListCount > 0 Then lst������.ListIndex = 0
    Else
        chk��ҩˢ��.Enabled = False
        lst������.Enabled = False
    End If
    
    chk����ʱ��.Value = IIf(mint����ʱ����� = 1, 1, 0)
    chkTakeDrug.Value = IIf(mint����ȡҩģʽ = 1, 1, 0)
    chkCheckStuff.Value = IIf(mint��ҩ���� = 1, 1, 0)
 End Function

Private Sub optCallWay_Click(index As Integer)
    If index = 0 Then
        FraԶ����������.Enabled = False
        frm�����㲥����.Enabled = True
    Else
        FraԶ����������.Enabled = True
        frm�����㲥����.Enabled = False
    End If
End Sub

Private Sub optSoundType_Click(index As Integer)
    If optSoundType(0).Value = True Then
        Label10.Caption = "�������٣�      (��Χ��0��100֮�䣬�Ƽ�65)"
        txtSpeed.Text = "65"
    Else
        Label10.Caption = "�������٣�      (��Χ��-10��10֮�䣬�Ƽ�-4)"
        txtSpeed.Text = "-4"
    End If
End Sub

Private Sub Opt��ӡ��ҩ��������_Click()
    lst��ӡ����.Enabled = False
End Sub

Private Sub Opt��ӡ��ҩ��������_GotFocus()
    sstMain.Tab = 2
End Sub

Private Sub Opt��ӡ��ҩ��������_Click()
    lst��ӡ����.Enabled = False
End Sub

Private Sub Opt��ӡ��ҩ��������_GotFocus()
    sstMain.Tab = 2
End Sub

Private Sub Opt��ӡ��ҩ��ѡ��_Click()
    lst��ӡ����.Enabled = Opt��ӡ��ҩ��ѡ��.Enabled
    If BlnStartUp = False Then Exit Sub
    
    If Opt��ӡ��ҩ��ѡ��.Value Then
        If lst��ӡ����.Enabled = True Then lst��ӡ����.SetFocus
    End If
End Sub
Private Sub Opt��ӡ��ҩ��ѡ��_GotFocus()
    sstMain.Tab = 2
End Sub

Private Sub pic��������_Resize()
    Dim objPic As PictureBox
    
    On Error Resume Next
    
    With sstMain
        .Left = 0
        .Top = 0
        .Height = pic��������.Height
        .Width = pic��������.Width
    End With
    
    For Each objPic In picPar
        objPic.Left = 0
        objPic.Top = 0
        objPic.Height = sstMain.Height
        objPic.Width = sstMain.Width
    Next
    
    '��ÿ��ҳǩ�еĲ����������й���
    '--------------------------------------------------
    '��������
    With frmҩ������
        .Top = 100
        .Left = 100
        .Width = sstMain.Width - .Left - 100
    End With
    
    With lst��ҩ����
        .Width = frmҩ������.Width - .Left - 100
    End With
    
    With frm��Ա����
        .Top = frmҩ������.Top + frmҩ������.Height + 100
        .Left = frmҩ������.Left
        .Width = frmҩ������.Width
    End With
    
    With frm�Զ���ҩ
        .Top = frm��Ա����.Top + frm��Ա����.Height + 100
        .Left = frmҩ������.Left
        .Width = frmҩ������.Width
    End With
    
    With frm������ʾ
        .Top = frm�Զ���ҩ.Top + frm�Զ���ҩ.Height + 100
        .Left = frmҩ������.Left
        .Width = frmҩ������.Width
    End With
    
    With frm���ڿ���
        .Top = frm������ʾ.Top + frm������ʾ.Height + 100
        .Left = frmҩ������.Left
        .Width = frmҩ������.Width
    End With
    
    With frm���˲鿴
        .Top = frm���ڿ���.Top + frm���ڿ���.Height + 100
        .Left = frmҩ������.Left
        .Width = frmҩ������.Width
    End With
    
    '����
    With frm��ʾ��ʽ
        .Top = 100
        .Left = 100
        .Width = sstMain.Width - .Left - 100
    End With
    
    With fra��֤��ʽ
        .Top = frm��ʾ��ʽ.Top + frm��ʾ��ʽ.Height + 100
        .Left = frm��ʾ��ʽ.Left
        .Width = frm��ʾ��ʽ.Width
    End With
    
    With frmˢ������
        .Top = fra��֤��ʽ.Top + fra��֤��ʽ.Height + 100
        .Left = frm��ʾ��ʽ.Left
        .Width = frm��ʾ��ʽ.Width
    End With
    
    With frm�豸ģʽ
        .Top = frmˢ������.Top + frmˢ������.Height + 100
        .Left = frm��ʾ��ʽ.Left
        .Width = frm��ʾ��ʽ.Width
    End With
    
    With frm����
        .Top = frm�豸ģʽ.Top + frm�豸ģʽ.Height + 100
        .Left = frm��ʾ��ʽ.Left
        .Width = frm��ʾ��ʽ.Width
    End With
    
    '��ӡ
    With frm��ӡ��ʽ
        .Top = 100
        .Left = 100
        .Width = sstMain.Width - .Left - 100
    End With
    
    With frm��ӡ����
        .Top = frm��ӡ��ʽ.Top + frm��ӡ��ʽ.Height + 100
        .Left = frm��ӡ��ʽ.Left
        .Width = frm��ӡ��ʽ.Width
    End With
    
    With frm�Զ���ӡ
        .Top = frm��ӡ����.Top + frm��ӡ����.Height + 100
        .Left = frm��ӡ��ʽ.Left
        .Width = frm��ӡ��ʽ.Width
    End With
    
    With lst��ӡ����
        .Width = frm�Զ���ӡ.Width - .Left - .Left
    End With
    
    With frm�Զ�ˢ��
        .Top = frm�Զ���ӡ.Top + frm�Զ���ӡ.Height + 100
        .Left = frm��ӡ��ʽ.Left
        .Width = frm��ӡ��ʽ.Width
    End With
    
    'Ʊ��
    With lblƱ��
        .Left = 150
        .Top = 180
    End With
    
    With cboƱ������
        .Left = lblƱ��.Left + lblƱ��.Width + 50
        .Top = lblƱ��.Top - (.Height - lblƱ��.Height) / 2
    End With
    
    With cmd��ӡ����
        .Left = lblƱ��.Left
        .Top = lblƱ��.Top + lblƱ��.Height + 200
    End With
    
    '��Դ����
    With lblFrom
        .Left = 100
        .Top = 100
    End With
    
    With Lvw��Դ����
        .Left = lblFrom.Left
        .Top = lblFrom.Top + lblFrom.Height + 100
        .Height = sstMain.Height - .Top - cmdClear.Height - 200
        .Width = sstMain.Width - .Left - 100
    End With
    
    With cmdClear
        .Left = sstMain.Width - .Width - 100
        .Top = Lvw��Դ����.Top + Lvw��Դ����.Height + 100
    End With
    
    With cmdCheckAll
        .Left = cmdClear.Left - .Width - 50
        .Top = cmdClear.Top
    End With
    
    '��������
    With Label5
        .Left = 100
        .Top = 100
    End With
    
    With Label6
        .Left = Label5.Left + Label5.Width + 50
        .Top = Label5.Top
    End With
    
    With vsfPrinter
        .Left = Label5.Left
        .Top = Label5.Top + Label5.Height + 100
        .Height = sstMain.Height - .Top - cmdDefaultPrinter.Height - 200
        .Width = sstMain.Width - .Left - 100
    End With
    
    With cmdDefaultPrinter
        .Left = sstMain.Width - .Width - 100
        .Top = vsfPrinter.Top + vsfPrinter.Height + 100
    End With
    
    With cmdDefaultColor
        .Left = vsfPrinter.Left
        .Top = cmdDefaultPrinter.Top
    End With
    
    '�Ŷӽк�
    With chk�����Ŷӽк�
        .Left = 100
        .Top = 100
    End With
    
    With frm��ʾ�豸����
        .Left = chk�����Ŷӽк�.Left
        .Top = chk�����Ŷӽк�.Top + chk�����Ŷӽк�.Height + 100
        .Width = sstMain.Width - .Left - 100
    End With
    
    With chkUseDisplay
        .Left = frm��ʾ�豸����.Left + chkUseSound.Left
        .Top = frm��ʾ�豸����.Top
    End With
    
    With Fra�����豸����
        .Left = chk�����Ŷӽк�.Left
        .Top = frm��ʾ�豸����.Top + frm��ʾ�豸����.Height + 100
        .Width = frm��ʾ�豸����.Width
    End With
    
    With frm�����㲥����
        .Width = Fra�����豸����.Width - .Left - 100
    End With
    
    With FraԶ����������
        .Width = frm�����㲥����.Width
    End With
End Sub

Private Sub pic����ҳǩ_Resize()
    With tplFunc
        .Left = 0
        .Top = 0
        .Width = pic����ҳǩ.Width
        .Height = pic����ҳǩ.Height
    End With
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
    Select Case sstMain.Tab
    Case 0
        If Me.Cboҩ��.Enabled = True Then Me.Cboҩ��.SetFocus
    Case 2
        If Me.Chk��ӡ��ҩ��.Enabled = True Then Me.Chk��ӡ��ҩ��.SetFocus
    Case 3
        If Me.cboƱ������.Enabled = True Then Me.cboƱ������.SetFocus
    End Select
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    '����:���taskpanel����Ľڵ㣬�ұ߲����������л�����Ӧ��ҳǩ
    
    Dim i As Integer
    Dim n As Integer
    Dim objPic As PictureBox
    
    '���ñ�ѡ��ʱ����ʾ״̬
    For i = 1 To tplFunc.Groups.count
        For n = 1 To tplFunc.Groups.Item(i).Items.count
            tplFunc.Groups.Item(i).Items.Item(n).Selected = False
        Next
    Next
    
    '�����ӽڵ㱻ѡ��ʱ����ʾ״̬
    Item.Selected = True
    
    '�����Ӧҳǩ�Ĳ�������
    For Each objPic In picPar
        objPic.Visible = (objPic.index + 1 = Item.Id)           '�����Ǵ�0��ʼ�ģ��ڵ�ID�Ǵ�1��ʼ��
    Next
End Sub

Private Sub txtOverTime_Change()
    txtOverTime.Text = Int(Val(txtOverTime.Text))
    If Val(txtOverTime.Text) > 1440 Then
        txtOverTime.Text = "1440"
    End If
End Sub

Private Sub txtOverTime_KeyPress(KeyAscii As Integer)
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

Private Sub txt��ӡ���_GotFocus()
    GetFocus txt��ӡ���
End Sub


Private Sub txt��ҩʱ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub


Private Sub Txt��ӡ�˷ѵ���_GotFocus()
    GetFocus Txt��ӡ�˷ѵ���
End Sub

Private Sub Txtˢ�¼��_GotFocus()
    GetFocus Txtˢ�¼��
End Sub

Private Sub Txt�ӳٴ�ӡ_GotFocus()
    GetFocus Txt�ӳٴ�ӡ
End Sub

Private Sub SetDispense()
'--------------------------------------
'������ҩ���Ƶ���ز���
'--------------------------------------
    Dim bln��ҩȷ�� As Boolean
    
    Me.chkIsDosage.Value = IIf(RecipeSendWork_DispensingMedi(Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex), bln��ҩȷ��) = True, 1, 0)
    
    Me.chkIsDosageOk.Value = IIf(bln��ҩȷ�� = True, 1, 0)
End Sub

Private Sub ReadWorkStationInf()
'*****************************************************
'��ȡվ����Ϣ
'*****************************************************

    Dim strsql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select ����վ from zlClients where ��ֹʹ��<>1 order by ����վ"
    Set rsTemp = zlDatabase.OpenSQLRecord(strsql, "��ȡվ����Ϣ")
    
    If rsTemp.EOF Then Exit Sub
    
    cboWorkStation.Clear
    
    While Not rsTemp.EOF
        Call cboWorkStation.AddItem(rsTemp("����վ"))
        rsTemp.MoveNext
    Wend
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function NOCheck() As Boolean
    Dim strsql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select 1 from δ��ҩƷ��¼ where �ⷿid=[1] and (����=8 or ����=9 or ����=10)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strsql, "NOCheck", Val(Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex)))
    
    If rsTemp.EOF Then
        NOCheck = True
    Else
        NOCheck = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDispensing()
'���ܣ���ʼ��chkDispensing�ؼ�

    Dim objMachine As Object
    
    err.Clear
    On Error Resume Next
    If Val(zlDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        '�����½ӿ�
        Set objMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If err.Number <> 0 Then
            '��ξɽӿ�
            Set objMachine = CreateObject("zlDrugPacker.clsDrugPacker")
        End If
    Else
        '�ɽӿ�
        Set objMachine = CreateObject("zlDrugPacker.clsDrugPacker")
    End If
    On Error GoTo 0
    
    If objMachine Is Nothing Then
        'ҩƷ�Զ����豸�ӿڲ�����
        chkDispensing.Visible = False
        chkDispensing.Value = 0
    Else
        'ҩƷ�Զ����豸�ӿڴ���
        chkDispensing.Visible = True
        chkDispensing.Value = Val(zlDatabase.GetPara("����ʱ֪ͨ��ʼ��ҩ", glngSys, 1341))
    End If
    
End Sub

Private Sub vsfPrinter_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPrinter
        '��ֹ�û��ԡ���ʽ���н�������
        If Col = mintCol��ʽ Then Cancel = True
    End With
End Sub

Private Sub vsfPrinter_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    If Col = mIntCol���� Then
        cmdialog.CancelError = True
        cmdialog.ShowColor
        
        vsfPrinter.Cell(flexcpBackColor, Row, mIntCol����) = cmdialog.Color
    End If
    
    Exit Sub
errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

