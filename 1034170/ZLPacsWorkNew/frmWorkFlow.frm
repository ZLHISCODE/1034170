VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmWorkFlow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   Icon            =   "frmWorkFlow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ӧ��(&A)"
      Height          =   350
      Left            =   4320
      TabIndex        =   66
      Top             =   7755
      Width           =   1100
   End
   Begin VB.Frame fraStudySetup 
      Height          =   2895
      Left            =   120
      TabIndex        =   44
      Top             =   8280
      Width           =   7575
      Begin VB.Frame Frame6 
         Caption         =   "��������"
         Height          =   2535
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   7335
         Begin VB.CheckBox chkUsePatient 
            Caption         =   "ʹ�û��ߺ�"
            Height          =   180
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   1288
         End
         Begin VB.CheckBox chkUseAdvice 
            Caption         =   "ʹ��ҽ����"
            Height          =   180
            Left            =   240
            TabIndex        =   67
            Top             =   750
            Width           =   1288
         End
         Begin VB.CheckBox chkAutoInc 
            Caption         =   "�Զ���������"
            Height          =   180
            Left            =   2160
            TabIndex        =   57
            Top             =   360
            Width           =   1635
         End
         Begin VB.OptionButton OptBuildcode 
            Caption         =   "���������Զ�����"
            Height          =   210
            Index           =   1
            Left            =   2520
            TabIndex        =   56
            ToolTipText     =   "�����Կ���Ϊ�������Զ�������"
            Top             =   840
            Width           =   1740
         End
         Begin VB.OptionButton OptBuildcode 
            Caption         =   "��ͬ�������Զ�����"
            Height          =   210
            Index           =   0
            Left            =   2520
            TabIndex        =   55
            ToolTipText     =   "�����Լ�����Ϊ�������Զ�������"
            Top             =   600
            Value           =   -1  'True
            Width           =   2130
         End
         Begin VB.Frame Frame7 
            Caption         =   "����һ����"
            Height          =   1920
            Left            =   4680
            TabIndex        =   49
            Top             =   240
            Width           =   2532
            Begin VB.Frame Frame10 
               Height          =   855
               Left            =   480
               TabIndex        =   52
               Top             =   930
               Width           =   1935
               Begin VB.OptionButton OptUnicode 
                  Caption         =   "������ͳһ"
                  Height          =   210
                  Index           =   1
                  Left            =   240
                  TabIndex        =   54
                  ToolTipText     =   "������ͬ�����ּ��Ų��䡣"
                  Top             =   520
                  Width           =   1290
               End
               Begin VB.OptionButton OptUnicode 
                  Caption         =   "��������ͳһ"
                  Height          =   210
                  Index           =   0
                  Left            =   240
                  TabIndex        =   53
                  ToolTipText     =   "��������ͬ�����ּ��Ų��䡣"
                  Top             =   220
                  Width           =   1590
               End
            End
            Begin VB.OptionButton OptCode 
               Caption         =   "���߼��ű��ֲ���"
               Height          =   180
               Index           =   1
               Left            =   360
               TabIndex        =   51
               ToolTipText     =   "ͬһ�����ߣ�����ʱ���ּ��Ų��䡣"
               Top             =   660
               Width           =   1935
            End
            Begin VB.OptionButton OptCode 
               Caption         =   "ÿ�μ�����¼���"
               Height          =   180
               Index           =   0
               Left            =   360
               TabIndex        =   50
               ToolTipText     =   "����ʱ�����µļ��š�"
               Top             =   345
               Value           =   -1  'True
               Width           =   1920
            End
         End
         Begin VB.CheckBox chkCanOverWrite 
            Caption         =   "��������ظ�"
            Height          =   180
            Left            =   240
            TabIndex        =   48
            ToolTipText     =   "����Ǽǲ��˵ļ��ų����ظ���"
            Top             =   1140
            Width           =   1935
         End
         Begin VB.CheckBox chkChangeNO 
            Caption         =   "�����ֹ���������"
            Height          =   180
            Left            =   240
            TabIndex        =   47
            ToolTipText     =   "�������ʵ����Ҫ�ֶ��޸ļ��š�"
            Top             =   1530
            Width           =   1935
         End
         Begin VB.CheckBox chkCheckMaxNo 
            Caption         =   "��ȡʵ��������"
            Height          =   180
            Left            =   240
            TabIndex        =   46
            ToolTipText     =   "��ʵ��������Ϊ����˳���ţ�����ѡ�����Ե�ǰ���õ�������˳���š�"
            Top             =   1920
            Width           =   1935
         End
      End
   End
   Begin VB.Frame framWorkFlow 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7815
      Begin VB.CheckBox chkPreView 
         Caption         =   "��������ͼԤ��"
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Top             =   5080
         Width           =   1575
      End
      Begin VB.Frame fra 
         Height          =   1455
         Index           =   27
         Left            =   120
         TabIndex        =   69
         Top             =   5160
         Width           =   4695
         Begin VB.TextBox txtDelayTime 
            Height          =   270
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   72
            ToolTipText     =   "0��ʾ���Զ��ر�"
            Top             =   652
            Width           =   495
         End
         Begin VB.OptionButton optMovePreview 
            Caption         =   "����ƶ�ʱԤ��ͼ��"
            Height          =   375
            Left            =   240
            TabIndex        =   71
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optClickPreview 
            Caption         =   "��굥��ʱԤ��ͼ��"
            Height          =   375
            Left            =   240
            TabIndex        =   70
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblDelayTime 
            Caption         =   "�ƶ�Ԥ��ʱ�Զ��ر���ʱʱ��       ��"
            Height          =   180
            Left            =   480
            TabIndex        =   73
            Top             =   697
            Width           =   3240
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "�ȼ��󱨵���ͼ��ƥ��"
         Height          =   1460
         Left            =   5160
         TabIndex        =   62
         Top             =   5160
         Width           =   2535
         Begin VB.OptionButton optMatch 
            Caption         =   "ҽ��ID"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   65
            ToolTipText     =   "����ʱͨ��ҽ��ID��ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "����"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   64
            ToolTipText     =   "����ʱͨ�����ź�ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "����/סԺ��"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   63
            ToolTipText     =   "����ʱͨ������/סԺ�ź�ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ƴ����"
         Height          =   1665
         Left            =   5160
         TabIndex        =   26
         Top             =   3320
         Width           =   2535
         Begin VB.OptionButton optCapital 
            Caption         =   "��д"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   32
            ToolTipText     =   "ѡ���ƴ������ʾȫΪ��д��ĸ��"
            Top             =   260
            Width           =   735
         End
         Begin VB.OptionButton optCapital 
            Caption         =   "Сд"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   31
            ToolTipText     =   "ѡ���ƴ������ʾȫΪСд��ĸ��"
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optCapital 
            Caption         =   "����ĸ��д"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   30
            ToolTipText     =   "ѡ���ƴ��������ĸ��д��"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Frame Frame9 
            Caption         =   "���"
            Height          =   540
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   2175
            Begin VB.OptionButton optSplitter 
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   29
               ToolTipText     =   "ƴ����֮���޼����"
               Top             =   200
               Width           =   495
            End
            Begin VB.OptionButton optSplitter 
               Caption         =   "�ո�"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   28
               ToolTipText     =   "ƴ����֮��ʹ�ÿո�Ϊ�������"
               Top             =   200
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "��������"
         Height          =   1665
         Left            =   120
         TabIndex        =   23
         Top             =   3320
         Width           =   4695
         Begin VB.CheckBox chkImgShowDesc 
            Caption         =   "ͼ������ʾ"
            Height          =   180
            Left            =   240
            TabIndex        =   78
            ToolTipText     =   "����ͼ�Ƿ�ͼ��ɼ�ʱ�䵹����ʾ��"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Frame Frame12 
            Caption         =   "���ٹ�������"
            Height          =   780
            Left            =   2280
            TabIndex        =   75
            Top             =   840
            Width           =   2175
            Begin VB.CheckBox chkNameQueryTimeLimit 
               Caption         =   "������ѯʱ������"
               Height          =   255
               Left            =   240
               TabIndex        =   77
               ToolTipText     =   "��������ѯʱ���Ƿ��в�ѯʱ������"
               Top             =   480
               Width           =   1800
            End
            Begin VB.CheckBox chkNameFuzzySearch 
               Caption         =   "����Ĭ��ģ����ѯ"
               Height          =   255
               Left            =   240
               TabIndex        =   76
               ToolTipText     =   "��������ѯʱʹ��ģ����ѯ��û�й�ѡʱ��ֻ������*��Ž���ģ����ѯ"
               Top             =   240
               Width           =   1800
            End
         End
         Begin VB.CheckBox chkSwitchUser 
            Caption         =   "�����л��û�"
            Height          =   180
            Left            =   240
            TabIndex        =   38
            ToolTipText     =   "�����л��û����ܣ����Խ����û��л���������Ӱ����վ��"
            Top             =   600
            Width           =   1455
         End
         Begin VB.Frame Frame2 
            Height          =   600
            Left            =   2280
            TabIndex        =   35
            ToolTipText     =   "ѡ��ɼ�ͼ���ɨ�����뵥��ʹ�õĴ洢�豸��"
            Top             =   160
            Width           =   2175
            Begin VB.ComboBox cboSaveDevice 
               Height          =   300
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   240
               Width           =   1725
            End
            Begin VB.CheckBox chkPetitionCapture 
               Caption         =   "�������뵥ɨ��"
               Height          =   180
               Left            =   120
               TabIndex        =   36
               ToolTipText     =   "������˺󣬸ü���Զ���ɡ�"
               Top             =   0
               Value           =   1  'Checked
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkUseReferencePatient 
            Caption         =   "���ù�������"
            Height          =   180
            Left            =   240
            TabIndex        =   25
            ToolTipText     =   "֧�ֶ����������ͬһ��������Ϣ��"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox chkChangeUser 
            Caption         =   "���ý����û�"
            Height          =   180
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "������û����ܣ����Խ������ҽ���ͱ���ҽ����������Ӱ��ɼ�վ��"
            Top             =   315
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����������"
         Height          =   3105
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   7600
         Begin VB.CheckBox chkEmergencyRequestNotExecuteMoney 
            Caption         =   "���ﲡ�˱�����ִ�з���"
            Height          =   180
            Left            =   120
            TabIndex        =   80
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CheckBox chkNoSignFinish 
            Caption         =   "����δǩ�������ӡ���"
            Height          =   180
            Left            =   5040
            TabIndex        =   79
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Frame Frame11 
            Caption         =   "ҽ��վ�鿴����"
            Height          =   615
            Left            =   5040
            TabIndex        =   60
            ToolTipText     =   "�������ڱ����ĵ��༭����"
            Top             =   2360
            Width           =   2415
            Begin VB.ComboBox cboViewReport 
               Height          =   300
               ItemData        =   "frmWorkFlow.frx":000C
               Left            =   240
               List            =   "frmWorkFlow.frx":0016
               Style           =   2  'Dropdown List
               TabIndex        =   61
               ToolTipText     =   "�������ڱ����ĵ��༭����"
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.CheckBox chkSetFocusWithReport 
            Caption         =   "����л�ʱ��λ����༭"
            Height          =   180
            Left            =   5040
            TabIndex        =   59
            ToolTipText     =   "�л�������ҳ��ʱ�Ƿ�λ����༭"
            Top             =   1707
            Width           =   2415
         End
         Begin VB.CheckBox chkFinallyCompleteCommit 
            Caption         =   "�����ֱ�����"
            Height          =   180
            Left            =   2640
            TabIndex        =   58
            ToolTipText     =   "��������󣬸ü���Զ���ɣ��������ڱ����ĵ��༭����"
            Top             =   1728
            Width           =   1935
         End
         Begin VB.TextBox txtViewHistoryImageDays 
            Height          =   270
            Left            =   6960
            MaxLength       =   2
            TabIndex        =   42
            Text            =   "1"
            Top             =   640
            Width           =   345
         End
         Begin VB.CheckBox chkAutoSendWorkList 
            Caption         =   "����ʱ�Զ�����WorkList"
            Height          =   252
            Left            =   120
            TabIndex        =   41
            Top             =   2020
            Value           =   1  'Checked
            Width           =   2412
         End
         Begin VB.CheckBox chkCompletePrint 
            Caption         =   "�����ֱ�Ӵ�ӡ"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            ToolTipText     =   "����ǩ����ֱ�Ӵ�ӡ���棬�������ڱ����ĵ��༭����"
            Top             =   2424
            Width           =   2040
         End
         Begin VB.CheckBox chkCanViewImage 
            Caption         =   "��ͼ��ҽ��վ���ɹ�Ƭ"
            Height          =   180
            Left            =   2640
            TabIndex        =   39
            ToolTipText     =   "�ɼ�ͼ�����û�м����ɵ�����£�ҽ��վҲ�ɽ��й�Ƭ��"
            Top             =   2760
            Width           =   2160
         End
         Begin VB.TextBox txtRefreshInterval 
            Enabled         =   0   'False
            Height          =   270
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   34
            Text            =   "1"
            Top             =   1340
            Width           =   390
         End
         Begin VB.TextBox TxtLike 
            Enabled         =   0   'False
            Height          =   270
            Left            =   7020
            MaxLength       =   2
            TabIndex        =   33
            ToolTipText     =   "0������ʱ������,ģ���������в���"
            Top             =   980
            Width           =   270
         End
         Begin VB.CheckBox ChkFinishCommit 
            Caption         =   "�ޱ�����ɺ�ֱ�����"
            Height          =   180
            Left            =   2640
            TabIndex        =   21
            ToolTipText     =   "����ޱ�����ɺ󣬸ü���Զ���ɡ�"
            Top             =   2412
            Width           =   2160
         End
         Begin VB.CheckBox chkPrintCommit 
            Caption         =   "��ӡ��ֱ�����"
            Height          =   180
            Left            =   2640
            TabIndex        =   20
            ToolTipText     =   "��ӡ����󣬸ü���Զ���ɡ�"
            Top             =   1044
            Width           =   1815
         End
         Begin VB.CheckBox ChkCompleteCommit 
            Caption         =   "��˺�ֱ�����"
            Height          =   180
            Left            =   2640
            TabIndex        =   19
            ToolTipText     =   "������˺󣬸ü���Զ���ɡ�"
            Top             =   1386
            Width           =   1935
         End
         Begin VB.CheckBox chkSample 
            Caption         =   "����ǼǺ�ֱ�ӱ���"
            Height          =   180
            Left            =   2640
            TabIndex        =   18
            ToolTipText     =   "�Ǽ��뱨��ͬʱ���С�"
            Top             =   2070
            Width           =   1935
         End
         Begin VB.TextBox TxtĬ������ 
            Height          =   270
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   17
            Text            =   "2"
            Top             =   320
            Width           =   585
         End
         Begin VB.CheckBox chkReportAfterImging 
            Caption         =   "��ͼ�����д����"
            Height          =   180
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "����ɼ�ͼ�����ܱ�дӰ�񱨸档"
            Top             =   360
            Width           =   2040
         End
         Begin VB.CheckBox chkPrintNeedComplete 
            Caption         =   "ƽ��������˲��ܴ򱨸�"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "ƽ������뾭����˺���ܴ�ӡ���档"
            Top             =   1024
            Width           =   2505
         End
         Begin VB.CheckBox chkTechReportSame 
            Caption         =   "ֻ����д�Լ����ı���"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "ֻ���Լ��ɼ�ͼ��ļ�飬������д���档"
            Top             =   692
            Width           =   2295
         End
         Begin VB.CheckBox chkWriteCapDoctor 
            Caption         =   "�ɼ�ͼ����Ϊ��鼼ʦ"
            Height          =   180
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "�ɼ�ͼ��֮���Զ�����ǰ�û���¼�ɼ�鼼ʦ��"
            Top             =   1356
            Width           =   2400
         End
         Begin VB.CheckBox chkLocalizerBackward 
            Caption         =   "��λƬ����"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "����λƬ�ŵ����һ��������ʾ��"
            Top             =   1688
            Width           =   1320
         End
         Begin VB.CheckBox chkRefreshInterval 
            Caption         =   "�����Զ�ˢ�¼��      ��"
            Height          =   180
            Left            =   5040
            TabIndex        =   11
            ToolTipText     =   "���˼���б������10-600�����Զ�ˢ�¡�"
            Top             =   1374
            Width           =   2500
         End
         Begin VB.CheckBox ChkLike 
            Caption         =   "�Ǽ�ʱ����ģ������    ��"
            Height          =   195
            Left            =   5040
            TabIndex        =   10
            ToolTipText     =   "�Ǽ�ʱ֧�ֶ���������ģ�����ң����Բ��ҵ�N���ڵ���Ϣ��"
            Top             =   1026
            Width           =   2520
         End
         Begin VB.CheckBox ChkReportFilmSameTime 
            Caption         =   "����ͽ�Ƭͬʱ����"
            Height          =   180
            Left            =   2640
            TabIndex        =   9
            ToolTipText     =   "�ڵ�����Ű�ťʱ����ͬʱ���ű���ͽ�Ƭ����������Ӱ��ҽ������վ��"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkAllPatientIsOutside 
            Caption         =   "���еǼǲ��˱��Ϊ����"
            Height          =   180
            Left            =   2640
            TabIndex        =   8
            ToolTipText     =   "���ڸù���վ�еǼǵĲ��˾����Ϊ�������ˡ�"
            Top             =   702
            Width           =   2295
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ�����ʷͼ������"
            Height          =   180
            Left            =   5040
            TabIndex        =   43
            ToolTipText     =   "�����ǰ���û��ͼ�����Զ���ָ��ʱ��Σ�1-15�죩�ڵ���ʷͼ��"
            Top             =   693
            Width           =   1800
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ĭ�ϼ�¼��ѯ����"
            Height          =   180
            Left            =   5040
            TabIndex        =   22
            ToolTipText     =   "����б���Ĭ����ʾ��Ӧ������1-15�죩�ڵļ���¼��"
            Top             =   360
            Width           =   1440
         End
      End
   End
   Begin VB.ComboBox cmbDept 
      Height          =   300
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   75
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6705
      TabIndex        =   3
      Top             =   7755
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5512
      TabIndex        =   2
      Top             =   7755
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7770
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl TabWindow 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _Version        =   589884
      _ExtentX        =   13996
      _ExtentY        =   12726
      _StockProps     =   64
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ӱ�����"
      Height          =   180
      Left            =   165
      TabIndex        =   5
      Top             =   135
      Width           =   735
   End
End
Attribute VB_Name = "frmWorkFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String         '��ģ���Ȩ��
Public mlng����ID As Long 'IN:��ǰִ�п���ID
Private mlngCur����ID As Long       '��ǰ����ID
Private mstrCur���� As String      '��ǰ���� ����-����
Private mstrCanUse���� As String    '��ǰ���ÿ���  ID_����-����
Private mobjfrmTabPass As New FrmReqInput     '��꾭������
Private mobjfrmEnableCtr As New FrmReqInput  '�������������
Private mobjFrmReportSetup As New frmReportSetup '��������
Private mobjFrmStudyListCfg As New frmStudyListCfg '����б�����
Private mobjfrmTechnicGroupCfg As New frmTechnicQueueCfg 'ҽ��ִ�м��������


Private Sub chkAutoInc_Click()
On Error Resume Next
    If chkAutoInc.value = 0 Then
        OptBuildcode(0).Enabled = False
        OptBuildcode(1).Enabled = False
        
        chkChangeNO.value = 1
        chkChangeNO.Enabled = False
        
        chkCheckMaxNo.value = 0
        chkCheckMaxNo.Enabled = False
    Else
        OptBuildcode(0).Enabled = True
        OptBuildcode(1).Enabled = True
        
        chkChangeNO.Enabled = True
        chkCheckMaxNo.Enabled = True
    End If
err.Clear
End Sub

Private Sub ChkCompleteCommit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ChkCompleteCommit.value = 1 Then chkFinallyCompleteCommit.value = 0
End Sub

Private Sub chkFinallyCompleteCommit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkFinallyCompleteCommit.value = 1 Then ChkCompleteCommit.value = 0
End Sub

Private Sub ChkLike_Click()
    TxtLike.Enabled = IIf(ChkLike.value, True, False)
End Sub

Private Sub chkPetitionCapture_Click()
    cboSaveDevice.Enabled = IIf(chkPetitionCapture.value, True, False)
End Sub

Private Sub chkPreView_Click()
    If chkPreView.value = 1 Then
        optMovePreview.Enabled = True
        lblDelayTime.Enabled = True
        txtDelayTime.Enabled = True
        optClickPreview.Enabled = True
    Else
        optMovePreview.Enabled = False
        lblDelayTime.Enabled = False
        txtDelayTime.Enabled = False
        optClickPreview.Enabled = False
    End If
End Sub

Private Sub chkRefreshInterval_Click()
    txtRefreshInterval.Enabled = IIf(chkRefreshInterval.value, True, False)
End Sub

Private Sub ConfigChkState()
    If chkUseAdvice.value = 0 And chkUsePatient.value = 0 Then
        OptCode(0).Enabled = True
        OptCode(1).Enabled = True
        If chkAutoInc.value = 0 Then
            OptBuildcode(0).Enabled = False
            OptBuildcode(1).Enabled = False
            
            chkChangeNO.value = 1
            chkChangeNO.Enabled = False
            
            chkCheckMaxNo.value = 0
            chkCheckMaxNo.Enabled = False
        Else
            OptBuildcode(0).Enabled = True
            OptBuildcode(1).Enabled = True
            
            chkChangeNO.Enabled = True
            chkCheckMaxNo.Enabled = True
        End If
        chkAutoInc.Enabled = True
        chkCanOverWrite.Enabled = True
    Else
        OptCode(0).value = True
        OptCode(0).Enabled = False
        OptCode(1).Enabled = False
  
        chkAutoInc.Enabled = False
        chkAutoInc.value = 0
        
        OptBuildcode(0).Enabled = False
        OptBuildcode(1).Enabled = False
        
        chkChangeNO.value = 0
        chkChangeNO.Enabled = False
        
        chkCheckMaxNo.value = 0
        chkCheckMaxNo.Enabled = False
        
        chkCanOverWrite.value = 1
        chkCanOverWrite.Enabled = False
    End If
End Sub

Private Sub chkUseAdvice_Click()
    If chkUseAdvice.value <> 0 Then
        chkUsePatient.value = 0
    End If
    
    Call ConfigChkState
End Sub

Private Sub chkUsePatient_Click()
    If chkUsePatient.value <> 0 Then
        chkUseAdvice.value = 0
    End If
    
    Call ConfigChkState
End Sub

Private Sub cmbDept_Click()
    mlng����ID = cmbDept.ItemData(cmbDept.ListIndex)
    If TabWindow.ItemCount = IIf(InStr(";" & GetPrivFunc(glngSys, 1160) & ";", ";����;") > 0, 8, 7) Then  '�ж�tab����=5��Ŀ����Ϊ��ȷ����װ����tab֮��Ŵ������е����
        'ˢ�¹������̲�������,�������ý���
        Call frmWorkFlowRefresh
        'ˢ��ִ�м����
        Call frmTechRoomRefresh
        'ˢ���������ý���
        Call frmReqInputRefresh(0)
        '���������
        Call frmReqInputRefresh(1)
        'ˢ�±�������
        Call frmReportRefresh
        'ˢ����ɫ����
        Call frmStudyListCfgRefresh
        'ˢ���Ŷӽк�����
        RefreshTechnicRoomGroupCfg
    End If
End Sub

Private Sub cmdApply_Click()
    Call SaveData
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub CmdOK_Click()

    Call SaveData
    
    Unload Me
End Sub

Private Sub SaveData()

    Call SaveWorkFlow
    Call mobjfrmTabPass.zlSave
    Call mobjfrmEnableCtr.zlSave
    Call mobjFrmReportSetup.zlSave
    Call mobjFrmStudyListCfg.zlSave
    Call mobjfrmTechnicGroupCfg.zlSave
End Sub

Private Sub Form_Load()
    '��ʼ��ģ�鼶����
    mstrPrivs = gstrPrivs
    mlng����ID = 0
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    
    mobjfrmTabPass.mintType = 0
    mobjfrmEnableCtr.mintType = 1
    
    'û�ж�Ӧ�Ŀ��ң����˳�
    If InitDepts = False Then
        Unload Me
        Exit Sub
    End If
    
    'װ���Ӵ���
    Call InitFaceScheme
    
    '��ʼ���Ӵ���
    'ˢ�¹������̲�������
    Call frmWorkFlowRefresh
    'ˢ��ִ�м����
    Call frmTechRoomRefresh
    'ˢ���������ý���
    Call frmReqInputRefresh(0)
    '���������
    Call frmReqInputRefresh(1)
    'ˢ�±�������
    Call frmReportRefresh
    'ˢ�¼���б�����
    Call frmStudyListCfgRefresh
    'ˢ���Ŷӽк�����
    Call RefreshTechnicRoomGroupCfg
End Sub

Private Sub Form_Resize()
    TabWindow.Left = 1
    TabWindow.Top = 480
    TabWindow.Width = Me.ScaleWidth
    TabWindow.Height = Me.ScaleHeight - 480
End Sub

Private Sub InitFaceScheme()
    Dim Item As TabControlItem
    
    mobjfrmTabPass.mlngDeptId = mlng����ID
    mobjfrmEnableCtr.mlngDeptId = mlng����ID
    frmTechnicRoom.mlngDept = mlng����ID
    
    With TabWindow
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 1, "����������", framWorkFlow.hWnd, 0
        .InsertItem 2, "��������", fraStudySetup.hWnd, 0
        .InsertItem 3, "ִ�м�����", frmTechnicRoom.hWnd, 0
        
        '��1160��Ȩ��ʱ�������������
        If CheckPopedom(";" & GetPrivFunc(glngSys, 1160) & ";", "����") Then
            .InsertItem 4, "�����Ŷ�����", mobjfrmTechnicGroupCfg.hWnd, 0
        End If
        
        .InsertItem 5, "���뾭������", mobjfrmTabPass.hWnd, 0
        .InsertItem 6, "�����¼����", mobjfrmEnableCtr.hWnd, 0
        .InsertItem 7, "PACS��������", mobjFrmReportSetup.hWnd, 0
        .InsertItem 8, "����б�����", mobjFrmStudyListCfg.hWnd, 0
        
        framWorkFlow.BorderStyle = 0
        .Item(0).Selected = True
    End With
    framWorkFlow.Width = Me.ScaleWidth
    framWorkFlow.Height = Me.ScaleHeight
    frmTechnicRoom.Width = Me.ScaleWidth
    frmTechnicRoom.Height = Me.ScaleHeight
    mobjfrmTabPass.Width = Me.ScaleWidth
    mobjfrmTabPass.Height = Me.ScaleHeight
    mobjfrmEnableCtr.Width = Me.ScaleWidth
    mobjfrmEnableCtr.Height = Me.ScaleHeight
    mobjFrmReportSetup.Width = Me.ScaleWidth
    mobjFrmReportSetup.Height = Me.ScaleHeight
    mobjFrmStudyListCfg.Width = Me.ScaleWidth
    mobjFrmStudyListCfg.Height = Me.ScaleHeight
    mobjfrmTechnicGroupCfg.Width = Me.ScaleWidth
    mobjfrmTechnicGroupCfg.Height = Me.ScaleHeight
End Sub

Private Sub frmTechRoomRefresh()
    'ˢ��ִ�м�ҳ��
    frmTechnicRoom.mlngDept = mlng����ID
    frmTechnicRoom.zlRoomRef
End Sub

Private Sub frmReqInputRefresh(ByVal intType As Integer)
    If intType = 0 Then
        mobjfrmTabPass.mlngDeptId = mlng����ID
        mobjfrmTabPass.zlRefresh
    ElseIf intType = 1 Then
        mobjfrmEnableCtr.mlngDeptId = mlng����ID
        mobjfrmEnableCtr.zlRefresh
    End If
End Sub

Private Sub frmStudyListCfgRefresh()
    Call mobjFrmStudyListCfg.zlRefresh(mlng����ID)
End Sub


Private Sub RefreshTechnicRoomGroupCfg()
'ˢ��ִ�м��������
    Call mobjfrmTechnicGroupCfg.zlRefresh(mlng����ID)
End Sub


Private Sub frmWorkFlowRefresh()
    Dim rsTemp As ADODB.Recordset
        
    '��ʼ��Ĭ��ֵ,Ӧ����һ��ͳһ�ĵط�����Ĭ��ֵ������������ʾ�����ն�ȡ
    
    ChkFinishCommit.value = 0   '�ޱ�����ɺ�ֱ�����
    chkReportAfterImging.value = 0  '��ͼ�񲻿ɱ༭����
    chkLocalizerBackward.value = 0  '��λƬ����
    chkChangeUser.value = 0         '�������û�
    chkSwitchUser.value = 0         '�����л��û�
    chkTechReportSame.value = 0     'ֻ����д�Լ����ı���
    chkWriteCapDoctor.value = 0     '�ɼ�ͼ����Ϊ��鼼ʦ
    ChkCompleteCommit.value = 0     '��˺�ֱ�����
    chkFinallyCompleteCommit.value = 0  '�����ֱ�����
    optMatch(0).value = True        'ƥ�����ݿ���Ŀ
    
    ChkLike.value = 0               '���õǼ�ʱ����ģ������
    TxtLike.Text = 0                '�Ǽ�ʱ����ģ����������
    TxtĬ������.Text = 2            'Ĭ�Ϲ�������
    txtViewHistoryImageDays.Text = 1 'Ĭ���Զ�����ʷͼ������
    chkRefreshInterval.value = 0    '���ò����б��Զ�ˢ��
    txtRefreshInterval.Text = 0     'Ĭ�ϲ����б��Զ�ˢ�¼��Ϊ0�룬��ˢ��
    cboSaveDevice.Clear                 '�洢�豸
    chkPrintCommit.value = 0        '��ӡ��ֱ�����
    chkCompletePrint.value = 0      '�����ֱ�Ӵ�ӡ
    chkUseReferencePatient.value = 0  'Ĭ�ϲ����ù�������
    chkImgShowDesc.value = 0
    optCapital(0).value = True      'Ĭ��ƴ��ʹ�ô�д
    optCapital(1).value = True      'Ĭ��ƴ������ÿո�
    chkCheckMaxNo.value = 1         'Ĭ����ȡʵ��������
    
    ChkReportFilmSameTime.value = 1 '����ͽ�Ƭͬʱ����Ĭ��Ϊѡ��
    
    chkPetitionCapture.value = 1     'Ĭ�Ϲ�ѡ�������뵥ɨ��
    If cboViewReport.ListCount > 0 Then cboViewReport.ListIndex = 0
    
    On Error GoTo err
    
    chkPetitionCapture.value = Val(GetDeptPara(mlng����ID, "�������뵥ɨ��", 1))    '��ȡ�������뵥ɨ�����

    ChkReportFilmSameTime.value = Val(GetDeptPara(mlng����ID, "����ͽ�Ƭͬʱ����", 1))  '��ȡ����ͽ�Ƭͬʱ���Ų���
    ChkFinishCommit.value = Val(GetDeptPara(mlng����ID, "�ޱ�����ɺ�ֱ�����", 0))
    chkCanViewImage.value = Val(GetDeptPara(mlng����ID, "��ͼ��ҽ��վ���ɹ�Ƭ", 0))
    chkReportAfterImging.value = Val(GetDeptPara(mlng����ID, "��ͼ�����д����", 0))
    chkNoSignFinish.value = Val(GetDeptPara(mlng����ID, "����δǩ�������ӡ���", 0))
    chkEmergencyRequestNotExecuteMoney.value = Val(GetDeptPara(mlng����ID, "���ﲡ�˱���ʱ��ִ�з���", 0))
    chkCanOverWrite.value = Val(GetDeptPara(mlng����ID, "��������ظ�", 0))
    chkCheckMaxNo.value = Val(GetDeptPara(mlng����ID, "��ȡʵ��������", 1))
    chkChangeNO.value = Val(GetDeptPara(mlng����ID, "�ֹ���������", 0))
    chkLocalizerBackward.value = Val(GetDeptPara(mlng����ID, "��λƬ����", 0))
    chkChangeUser.value = Val(GetDeptPara(mlng����ID, "�������û�", 0))
    chkSwitchUser.value = Val(GetDeptPara(mlng����ID, "�����л��û�", 0))
    chkTechReportSame.value = Val(GetDeptPara(mlng����ID, "ֻ����д�Լ����ı���", 0))
    chkWriteCapDoctor.value = Val(GetDeptPara(mlng����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", 0))
    ChkCompleteCommit.value = Val(GetDeptPara(mlng����ID, "��˺�ֱ�����", 0))
    chkFinallyCompleteCommit.value = Val(GetDeptPara(mlng����ID, "�����ֱ�����", 0))
    chkPrintCommit.value = Val(GetDeptPara(mlng����ID, "��ӡ��ֱ�����", 0))
    chkCompletePrint.value = Val(GetDeptPara(mlng����ID, "�����ֱ�Ӵ�ӡ", 0))
    
    TxtLike.Text = Val(GetDeptPara(mlng����ID, "�Ǽ�ʱ����ģ����������", 0))
    chkSample.value = Val(GetDeptPara(mlng����ID, "�ǼǺ�ֱ�Ӽ��", 0))
    ChkLike.value = IIf(Val(TxtLike.Text) <> 0, 1, 0)
    chkAllPatientIsOutside.value = Val(GetDeptPara(mlng����ID, "���еǼǲ��˱��Ϊ����", 0))
    
    TxtĬ������.Text = Val(GetDeptPara(mlng����ID, "Ĭ�Ϲ�������", 2))
    
    If Val(TxtĬ������.Text) > 15 Or Val(TxtĬ������.Text) <= 0 Then
        TxtĬ������.Text = 2
    End If
    
    txtViewHistoryImageDays.Text = Val(GetDeptPara(mlng����ID, "�Զ�����ʷͼ������", 1))
    If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
        txtViewHistoryImageDays.Text = 1
    End If
    
    txtRefreshInterval.Text = Val(GetDeptPara(mlng����ID, "�Զ�ˢ�¼��", 0))
    chkRefreshInterval.value = IIf(Val(txtRefreshInterval.Text) <> 0, 1, 0)
    optMatch(Val(GetDeptPara(mlng����ID, "ƥ�����ݿ���Ŀ", 0))).value = True
    
    OptBuildcode(Val(GetDeptPara(mlng����ID, "�������ɷ�ʽ", 0))).value = True
    chkAutoInc.value = Val(GetDeptPara(mlng����ID, "�Զ���������"))
    chkUseAdvice.value = Val(GetDeptPara(mlng����ID, "ʹ��ҽ����", 0))
    chkUsePatient.value = Val(GetDeptPara(mlng����ID, "ʹ�û��ߺ�", 0))
    chkAutoSendWorkList.value = Val(GetDeptPara(mlng����ID, "����ʱ�Զ�����WorkList", "1"))
    chkSetFocusWithReport.value = Val(GetDeptPara(mlng����ID, "����л�ʱ��λ����༭", "1"))
    chkNameFuzzySearch.value = Val(GetDeptPara(mlng����ID, "����Ĭ��ģ����ѯ", "1"))
    chkNameQueryTimeLimit.value = Val(GetDeptPara(mlng����ID, "������ѯʱ������", "1"))
    
    If Val(GetDeptPara(mlng����ID, "ҽ��վ�鿴����", "1")) = 0 Then
        cboViewReport.ListIndex = 0
    Else
        cboViewReport.ListIndex = 1
    End If
    
    OptCode(Val(GetDeptPara(mlng����ID, "���߼��ű��ֲ���", 0))).value = True
    If OptCode(1).value = True Then
        OptUnicode(0).Enabled = True
        OptUnicode(1).Enabled = True
        OptUnicode(Val(GetDeptPara(mlng����ID, "���ű��ֲ������", 0))).value = True
    Else
        OptUnicode(0).Enabled = False: OptUnicode(0).value = False
        OptUnicode(1).Enabled = False: OptUnicode(1).value = False
    End If
    
    If chkUseAdvice.value = 0 And chkUsePatient.value = 0 Then
        OptCode(0).Enabled = True
        OptCode(1).Enabled = True
        
        If chkAutoInc.value = 0 Then
            OptBuildcode(0).Enabled = False
            OptBuildcode(1).Enabled = False
            
            chkChangeNO.value = 1
            chkChangeNO.Enabled = False
            
            chkCheckMaxNo.value = 0
            chkCheckMaxNo.Enabled = False
        Else
            OptBuildcode(0).Enabled = True
            OptBuildcode(1).Enabled = True
            
            chkChangeNO.Enabled = True
            chkCheckMaxNo.Enabled = True
        End If
        
        chkAutoInc.Enabled = True
        chkCanOverWrite.Enabled = True
    Else
        chkAutoInc.value = 0
        chkAutoInc.Enabled = False
        
        OptBuildcode(0).Enabled = False
        OptBuildcode(1).Enabled = False
        
        chkChangeNO.value = 0
        chkChangeNO.Enabled = False
        
        chkCheckMaxNo.value = 0
        chkCheckMaxNo.Enabled = False
        
        chkCanOverWrite.value = 1
        chkCanOverWrite.Enabled = False
        
        OptCode(0).Enabled = True
        OptCode(1).Enabled = True
    End If
    
    chkPreView.value = IIf(Val(GetDeptPara(mlng����ID, "����ͼԤ����ʽ", "0")) > 0, 1, 0)
        
    If chkPreView.value = 1 Then
        optMovePreview.Enabled = True
        lblDelayTime.Enabled = True
        txtDelayTime.Enabled = True
        optClickPreview.Enabled = True
    Else
        optMovePreview.Enabled = False
        lblDelayTime.Enabled = False
        txtDelayTime.Enabled = False
        optClickPreview.Enabled = False
    End If
    
    optMovePreview.value = Val(GetDeptPara(mlng����ID, "����ͼԤ����ʽ", "0")) = 1
    optClickPreview.value = Val(GetDeptPara(mlng����ID, "����ͼԤ����ʽ", "0")) = 2
    txtDelayTime.Text = Val(GetDeptPara(mlng����ID, "�ƶ�Ԥ����ʱ", "2"))
    
    
    chkUseReferencePatient.value = Val(GetDeptPara(mlng����ID, "������������", 0))
    chkImgShowDesc.value = Val(GetDeptPara(mlng����ID, "ͼ������ʾ", 0))
    chkPrintNeedComplete.value = Val(GetDeptPara(mlng����ID, "ƽ������˲��ܴ򱨸�", 0))
    
    'ƴ��������
    optCapital(Val(GetDeptPara(mlng����ID, "ƴ������Сд", 0))).value = True
    optSplitter(Val(GetDeptPara(mlng����ID, "ƴ�����ָ���", 0))).value = True
    
    
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and NVL(״̬,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTemp.EOF Then
        MsgBoxD Me, "δ�������뵥�洢�豸���뵽Ӱ���豸Ŀ¼�����ã�", vbInformation, gstrSysName
        Exit Sub
    Else
        cboSaveDevice.AddItem ""
        
        Do While Not rsTemp.EOF
            cboSaveDevice.AddItem rsTemp!�豸�� & "-" & Nvl(rsTemp!�豸��)
            
            If GetDeptPara(mlng����ID, "���뵥�洢�豸��", "") = rsTemp!�豸�� Then
                cboSaveDevice.ListIndex = cboSaveDevice.NewIndex
            End If
            
            rsTemp.MoveNext
        Loop
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub SaveWorkFlow()
    Dim lngTemp As Long
    
    On Error GoTo errHand

    SetDeptPara mlng����ID, "�������뵥ɨ��", chkPetitionCapture.value        '�������뵥ɨ�� ��������
    SetDeptPara mlng����ID, "����ͽ�Ƭͬʱ����", ChkReportFilmSameTime.value '����ͽ�Ƭͬʱ���� ��������
    
    
    
    SetDeptPara mlng����ID, "�ޱ�����ɺ�ֱ�����", ChkFinishCommit.value
    SetDeptPara mlng����ID, "��ͼ��ҽ��վ���ɹ�Ƭ", chkCanViewImage.value     '��ͼ��ҽ��վ���ɹ�Ƭ
    SetDeptPara mlng����ID, "��ͼ�����д����", chkReportAfterImging.value
    SetDeptPara mlng����ID, "����δǩ�������ӡ���", chkNoSignFinish.value     'δǩ�������ӡ���
    SetDeptPara mlng����ID, "���ﲡ�˱���ʱ��ִ�з���", chkEmergencyRequestNotExecuteMoney.value     '���ﲡ�˱���ʱ��ִ�з���
    SetDeptPara mlng����ID, "���߼��ű��ֲ���", IIf(OptCode(1).value, 1, 0)
    SetDeptPara mlng����ID, "���ű��ֲ������", IIf(OptUnicode(1).value, 1, 0)
    SetDeptPara mlng����ID, "�������ɷ�ʽ", IIf(OptBuildcode(1).value, 1, 0)
    SetDeptPara mlng����ID, "ʹ��ҽ����", chkUseAdvice.value
    SetDeptPara mlng����ID, "ʹ�û��ߺ�", chkUsePatient.value
    SetDeptPara mlng����ID, "�Զ���������", chkAutoInc.value
    SetDeptPara mlng����ID, "�ֹ���������", chkChangeNO.value
    SetDeptPara mlng����ID, "��������ظ�", chkCanOverWrite.value
    SetDeptPara mlng����ID, "��ȡʵ��������", chkCheckMaxNo.value
    SetDeptPara mlng����ID, "��λƬ����", chkLocalizerBackward.value
    SetDeptPara mlng����ID, "�������û�", chkChangeUser.value
    SetDeptPara mlng����ID, "�����л��û�", chkSwitchUser.value
    SetDeptPara mlng����ID, "ֻ����д�Լ����ı���", chkTechReportSame.value
    SetDeptPara mlng����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", chkWriteCapDoctor.value
    SetDeptPara mlng����ID, "��˺�ֱ�����", ChkCompleteCommit.value
    SetDeptPara mlng����ID, "�����ֱ�����", chkFinallyCompleteCommit.value
    SetDeptPara mlng����ID, "��ӡ��ֱ�����", chkPrintCommit.value
    SetDeptPara mlng����ID, "�����ֱ�Ӵ�ӡ", chkCompletePrint.value
    SetDeptPara mlng����ID, "�ǼǺ�ֱ�Ӽ��", chkSample.value
    SetDeptPara mlng����ID, "ƥ�����ݿ���Ŀ", IIf(optMatch(0).value, 0, IIf(optMatch(1), 1, 2))
    
    SetDeptPara mlng����ID, "�Ǽ�ʱ����ģ����������", IIf(ChkLike.value = 1, Abs(Val(TxtLike.Text)), 0)
    SetDeptPara mlng����ID, "���еǼǲ��˱��Ϊ����", chkAllPatientIsOutside
    
    If Val(TxtĬ������.Text) > 15 Or Val(TxtĬ������.Text) <= 0 Then
        TxtĬ������.Text = 2
    End If
    SetDeptPara mlng����ID, "Ĭ�Ϲ�������", Val(TxtĬ������.Text)
    
    If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
        txtViewHistoryImageDays.Text = 1
    End If
    SetDeptPara mlng����ID, "�Զ�����ʷͼ������", Val(txtViewHistoryImageDays.Text)
    
    SetDeptPara mlng����ID, "������������", chkUseReferencePatient.value
    SetDeptPara mlng����ID, "ƽ������˲��ܴ򱨸�", chkPrintNeedComplete.value
    SetDeptPara mlng����ID, "ͼ������ʾ", chkImgShowDesc.value
    
    SetDeptPara mlng����ID, "ƴ������Сд", IIf(optCapital(0).value, 0, IIf(optCapital(1), 1, 2))
    SetDeptPara mlng����ID, "ƴ�����ָ���", IIf(optSplitter(0).value, 0, 1)
    
    If cboSaveDevice.Text <> "" Then
        SetDeptPara mlng����ID, "���뵥�洢�豸��", Split(cboSaveDevice.Text, "-")(0)
    Else
        SetDeptPara mlng����ID, "���뵥�洢�豸��", ""
    End If
    
    If Abs(Val(txtRefreshInterval.Text)) = 0 Or Abs(Val(txtRefreshInterval.Text)) > 65 Then
        txtRefreshInterval.Text = 10
    End If
    SetDeptPara mlng����ID, "�Զ�ˢ�¼��", IIf(chkRefreshInterval.value = 1, Abs(Val(txtRefreshInterval.Text)), 0)
    SetDeptPara mlng����ID, "����ʱ�Զ�����WorkList", chkAutoSendWorkList.value
    SetDeptPara mlng����ID, "ҽ��վ�鿴����", cboViewReport.ListIndex
    SetDeptPara mlng����ID, "����л�ʱ��λ����༭", chkSetFocusWithReport.value
    SetDeptPara mlng����ID, "����Ĭ��ģ����ѯ", chkNameFuzzySearch.value
    SetDeptPara mlng����ID, "������ѯʱ������", chkNameQueryTimeLimit.value
    
    If chkPreView.value = 1 Then
        If optMovePreview.value Then
            lngTemp = 1
        ElseIf optClickPreview.value Then
            lngTemp = 2
        End If
    Else
        lngTemp = 0
    End If
    
    SetDeptPara mlng����ID, "����ͼԤ����ʽ", lngTemp
    SetDeptPara mlng����ID, "�ƶ�Ԥ����ʱ", Val(txtDelayTime.Text)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Private Function InitDepts() As Boolean
'���ܣ���ʼ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str����IDs As String, str��Դ As String
    Dim strDepartment() As String
    Dim intCurDept As Integer
    
    On Error GoTo errH
    
    If CheckPopedom(mstrPrivs, "���п���") Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.�������� IN('���')  Order by A.����"
    Else
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.�������� IN('���')  Order by A.����"
    End If
     
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
            rsTmp.MoveNext
        Loop
        
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
            rsTmp.MoveNext
        Loop
        mstrCanUse���� = Mid(mstrCanUse����, 2)
        If InStr(mstrPrivs, "���п���") > 0 And mlngCur����ID = 0 Then
            mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
            mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
        End If
        
        If mlngCur����ID = 0 And InStr(mstrPrivs, "���п���") <= 0 Then 'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
            MsgBoxD Me, "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        '���cmbDept
        cmbDept.Clear
        intCurDept = -1
        strDepartment = Split(mstrCanUse����, "|")
        For i = 0 To UBound(strDepartment)
            cmbDept.AddItem Split(strDepartment(i), "_")(1)
            cmbDept.ItemData(cmbDept.ListCount - 1) = Split(strDepartment(i), "_")(0)
            If Split(strDepartment(i), "_")(0) = mlngCur����ID Then
                intCurDept = i
            End If
        Next i
        If intCurDept <> -1 Then
            cmbDept.ListIndex = intCurDept
        Else
            cmbDept.ListIndex = 0
        End If
        mlng����ID = cmbDept.ItemData(cmbDept.ListIndex)
        InitDepts = True
    End If
    
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTechnicRoom
    Unload mobjfrmEnableCtr
    Unload mobjfrmTabPass
    Unload mobjFrmReportSetup
    Unload mobjFrmStudyListCfg
    Unload mobjfrmTechnicGroupCfg
End Sub


Private Sub optClickPreview_Click()
    If optMovePreview.value = False Then
        txtDelayTime.Enabled = False
        lblDelayTime.Enabled = False
    End If
End Sub

Private Sub OptCode_Click(Index As Integer)
    OptUnicode(0).Enabled = Index = 1
    OptUnicode(1).Enabled = Index = 1
End Sub

Private Sub frmReportRefresh()
    mobjFrmReportSetup.zlRefresh (mlng����ID)
End Sub

Private Sub optMovePreview_Click()
    If optMovePreview.value = True Then
        txtDelayTime.Enabled = True
        lblDelayTime.Enabled = True
    End If
End Sub

Private Sub txtDelayTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtLike_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRefreshInterval_Change()
    If Val(txtRefreshInterval.Text) > 600 Then
        txtRefreshInterval.Text = 600
    End If
End Sub

Private Sub txtRefreshInterval_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRefreshInterval_LostFocus()
    If Val(txtRefreshInterval.Text) < 10 Then
        txtRefreshInterval.Text = 10
    End If
End Sub

Private Sub txtViewHistoryImageDays_Change()
    If Val(txtViewHistoryImageDays.Text) > 15 Then
        txtViewHistoryImageDays.Text = 15
    End If
End Sub

Private Sub txtViewHistoryImageDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtViewHistoryImageDays_LostFocus()
    If Val(txtViewHistoryImageDays.Text) < 1 Then
        txtViewHistoryImageDays.Text = 1
    End If
End Sub

Private Sub TxtĬ������_Change()
    If Val(TxtĬ������.Text) > 15 Then
        TxtĬ������.Text = 15
    End If
End Sub

Private Sub TxtĬ������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtĬ������_LostFocus()
    If Val(TxtĬ������.Text) <= 0 Then
        TxtĬ������.Text = 1
    End If
End Sub
