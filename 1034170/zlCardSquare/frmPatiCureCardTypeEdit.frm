VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatiCureCardTypeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�ƿ����༭"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   Icon            =   "frmPatiCureCardTypeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410.12
   ScaleMode       =   0  'User
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame7 
      Caption         =   "����ˢ������"
      Height          =   1875
      Left            =   5445
      TabIndex        =   79
      Top             =   3360
      Width           =   2565
      Begin VB.CheckBox chkȱʡ���� 
         Caption         =   "ȱʡ����"
         Height          =   180
         Left            =   1455
         TabIndex        =   42
         Top             =   270
         Width           =   1020
      End
      Begin VB.CheckBox chk�˿��鿨 
         Caption         =   "�˿�ʱ��Ҫ�鿨(&D)"
         Height          =   180
         Left            =   150
         TabIndex        =   47
         Top             =   1620
         Width           =   2010
      End
      Begin VB.CheckBox chk���͵��ýӿ� 
         Caption         =   "ҽ����������֧������(&A)"
         Height          =   180
         Left            =   150
         TabIndex        =   46
         Top             =   1350
         Width           =   2370
      End
      Begin VB.CheckBox chkת�ʼ����� 
         Caption         =   "֧��ת�ʼ�����(&H)"
         Height          =   180
         Left            =   150
         TabIndex        =   44
         Top             =   810
         Width           =   1875
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "�����˿�(&S)"
         Height          =   180
         Index           =   5
         Left            =   150
         TabIndex        =   43
         Top             =   540
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "��������(&R)"
         Height          =   180
         Index           =   7
         Left            =   150
         TabIndex        =   41
         Top             =   270
         Width           =   1455
      End
      Begin VB.CheckBox chk�ֿ����� 
         Caption         =   "����ֿ�����(&P)"
         Height          =   180
         Left            =   150
         TabIndex        =   45
         Top             =   1080
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8130
      TabIndex        =   60
      Top             =   885
      Width           =   1100
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   7305
      Index           =   0
      Left            =   135
      TabIndex        =   61
      Top             =   90
      Width           =   7935
      Begin VB.Frame frm���� 
         Caption         =   "��������"
         Height          =   795
         Left            =   0
         TabIndex        =   80
         Top             =   3660
         Width           =   5235
         Begin VB.OptionButton OptSendCard 
            Caption         =   "ͬһ�����˷����ſ�ʱ����"
            Height          =   285
            Index           =   2
            Left            =   300
            TabIndex        =   27
            Top             =   480
            Width           =   3495
         End
         Begin VB.OptionButton OptSendCard 
            Caption         =   "������"
            Height          =   285
            Index           =   0
            Left            =   300
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton OptSendCard 
            Caption         =   "ͬһ������ֻ��һ�ſ�"
            Height          =   285
            Index           =   1
            Left            =   2760
            TabIndex        =   25
            Top             =   240
            Width           =   2115
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "�������ų��Ȳ���ʱ"
         Height          =   600
         Left            =   0
         TabIndex        =   78
         Top             =   4545
         Width           =   5235
         Begin VB.OptionButton OptSendCardLen 
            Caption         =   "��ֹ����"
            Height          =   285
            Index           =   0
            Left            =   1860
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton OptSendCardLen 
            Caption         =   "������"
            Height          =   285
            Index           =   1
            Left            =   300
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton OptSendCardLen 
            Caption         =   "���ѷ���"
            Height          =   285
            Index           =   2
            Left            =   3615
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdInsureSel 
         Caption         =   "&P"
         Height          =   270
         Left            =   4920
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2925
         Width           =   270
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "��ע"
         Top             =   2910
         Width           =   4200
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   1410
      End
      Begin VB.Frame fra���� 
         Caption         =   "������������"
         Height          =   2100
         Left            =   0
         TabIndex        =   65
         Top             =   5220
         Width           =   7875
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1455
            TabIndex        =   75
            Top             =   690
            Width           =   5430
            Begin VB.OptionButton optPassConfine 
               Caption         =   "������"
               Height          =   210
               Index           =   0
               Left            =   -30
               TabIndex        =   51
               Top             =   105
               Width           =   1200
            End
            Begin VB.OptionButton optPassConfine 
               Caption         =   "��������������"
               Height          =   210
               Index           =   1
               Left            =   1370
               TabIndex        =   52
               Top             =   120
               Width           =   1890
            End
            Begin VB.OptionButton optPassConfine 
               Caption         =   "�����������ֹ"
               Height          =   210
               Index           =   2
               Left            =   3720
               TabIndex        =   53
               Top             =   105
               Value           =   -1  'True
               Width           =   1890
            End
         End
         Begin VB.Frame Frame4 
            Height          =   30
            Left            =   0
            TabIndex        =   74
            Top             =   1590
            Width           =   7875
         End
         Begin VB.TextBox txtPassByIDCard 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1710
            TabIndex        =   73
            Text            =   "0"
            Top             =   1740
            Width           =   300
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   0
            TabIndex        =   71
            Top             =   1080
            Width           =   7890
         End
         Begin VB.Frame fraSplit 
            Height          =   30
            Left            =   0
            TabIndex        =   69
            Top             =   585
            Width           =   7875
         End
         Begin VB.Frame fraRule 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1425
            TabIndex        =   68
            Top             =   270
            Width           =   3495
            Begin VB.OptionButton optRule 
               Caption         =   "�����ַ�ֻ��Ϊ����"
               Height          =   180
               Index           =   1
               Left            =   1395
               TabIndex        =   49
               Top             =   45
               Width           =   2070
            End
            Begin VB.OptionButton optRule 
               Caption         =   "�����ַ�"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   48
               Top             =   45
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin VB.TextBox txtPassInput 
            Height          =   270
            Left            =   6195
            TabIndex        =   57
            Text            =   "0"
            Top             =   1275
            Width           =   300
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "��������    λ��������"
            Height          =   210
            Index           =   2
            Left            =   5175
            TabIndex        =   56
            Top             =   1305
            Width           =   2295
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "���벻�̶�"
            Height          =   210
            Index           =   0
            Left            =   1425
            TabIndex        =   54
            Top             =   1290
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.TextBox txtPasLen 
            Height          =   270
            Left            =   6285
            MaxLength       =   2
            TabIndex        =   50
            Text            =   "10"
            Top             =   270
            Width           =   300
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "�̶�����10λ"
            Height          =   210
            Index           =   1
            Left            =   2825
            TabIndex        =   55
            Top             =   1305
            Width           =   1665
         End
         Begin VB.CheckBox chkPassByIDCard 
            Caption         =   "ȱʡ�����֤��    λΪȱʡ����:��ʾȱʡ����λ���������볤���Զ���ȡ"
            Height          =   360
            Left            =   150
            TabIndex        =   58
            Top             =   1680
            Width           =   6900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "�����������"
            Height          =   180
            Left            =   150
            TabIndex        =   81
            Top             =   1260
            Width           =   1080
         End
         Begin VB.Label lbl������������ 
            AutoSize        =   -1  'True
            Caption         =   "������������"
            Height          =   180
            Left            =   150
            TabIndex        =   72
            Top             =   765
            Width           =   1080
         End
         Begin VB.Label lbl������� 
            AutoSize        =   -1  'True
            Caption         =   "���빹�ɹ���"
            Height          =   180
            Left            =   150
            TabIndex        =   67
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "������󳤶�    λ"
            Height          =   180
            Left            =   5190
            TabIndex        =   66
            Top             =   315
            Width           =   1620
         End
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "&P"
         Height          =   270
         Left            =   4920
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1740
         Width           =   270
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   1005
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "ҽ�ƿ���"
         Top             =   1725
         Width           =   4200
      End
      Begin VB.TextBox txt����λ�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "1"
         Top             =   3330
         Width           =   360
      End
      Begin MSComCtl2.UpDown upd��ʼλ�� 
         Height          =   300
         Left            =   2235
         TabIndex        =   22
         Top             =   3300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt��ʼλ��"
         BuddyDispid     =   196642
         OrigLeft        =   1455
         OrigTop         =   2550
         OrigRight       =   1710
         OrigBottom      =   2940
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.TextBox txt��ʼλ�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "1"
         Top             =   3330
         Width           =   315
      End
      Begin VB.Frame Frame1 
         Caption         =   "ҽ�ƿ�����"
         Height          =   3165
         Left            =   5325
         TabIndex        =   62
         Top             =   60
         Width           =   2550
         Begin VB.CheckBox chkOpenEnter 
            Caption         =   "�豸���ûس�(&0)"
            Height          =   180
            Left            =   150
            TabIndex        =   40
            Top             =   2775
            Width           =   1650
         End
         Begin VB.CheckBox chkWriteCard 
            Caption         =   "����д��(&6)"
            Height          =   195
            Left            =   150
            TabIndex        =   36
            Top             =   1635
            Width           =   1695
         End
         Begin VB.CheckBox chkSendCard 
            Caption         =   "������(&5)"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   1365
            Width           =   1695
         End
         Begin VB.CheckBox chkMakeCard 
            Caption         =   "�����ƿ�(&4)"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   1095
            Width           =   1695
         End
         Begin VB.CheckBox chkģ������ 
            Caption         =   "֧��ģ������(&9)"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   2505
            Width           =   1695
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "�����ظ�ʹ��(&7)"
            Height          =   270
            Index           =   6
            Left            =   150
            TabIndex        =   37
            Top             =   1905
            Width           =   1665
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "ȱʡˢ�����(&8)"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   38
            Top             =   2175
            Width           =   1695
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "�����ʻ�(&3)"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   33
            Top             =   825
            Width           =   1305
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "�š�����(&2)"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   32
            ToolTipText     =   "���ϴ���ſ�ˢ������IC���������"
            Top             =   555
            Width           =   1320
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "�ϸ����(&1)"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   31
            Top             =   285
            Width           =   1320
         End
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   1005
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "����"
         Top             =   2130
         Width           =   4200
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   300
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1315
         Width           =   1515
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3675
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "ǰ׺�ı�"
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1005
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "����"
         Top             =   915
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1005
         MaxLength       =   100
         TabIndex        =   4
         Tag             =   "����"
         Top             =   525
         Width           =   4200
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "��ע"
         Top             =   2520
         Width           =   4200
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   3675
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "����"
         Top             =   150
         Width           =   1515
      End
      Begin MSComCtl2.UpDown upd����λ�� 
         Height          =   300
         Left            =   3360
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt����λ��"
         BuddyDispid     =   196641
         OrigLeft        =   1455
         OrigTop         =   2550
         OrigRight       =   1710
         OrigBottom      =   2940
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���Ŵ�        λ��        λ������ʾ(&M)"
         Height          =   180
         Left            =   1020
         TabIndex        =   20
         Top             =   3390
         Width           =   4665
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   1005
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "���ų���"
         Top             =   1315
         Width           =   1395
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&X)"
         Height          =   180
         Index           =   9
         Left            =   330
         TabIndex        =   76
         Top             =   2970
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�����(&G)"
         Height          =   180
         Left            =   165
         TabIndex        =   70
         Top             =   210
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ҽ�ƿ���(&F)"
         Height          =   180
         Index           =   5
         Left            =   -15
         TabIndex        =   13
         Top             =   1785
         Width           =   990
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(M)"
         Height          =   180
         Index           =   8
         Left            =   345
         TabIndex        =   15
         Top             =   2190
         Width           =   630
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���㷽ʽ(&J)"
         Height          =   180
         Index           =   7
         Left            =   2670
         TabIndex        =   11
         Top             =   1375
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���ų���(&L)"
         Height          =   180
         Index           =   6
         Left            =   -15
         TabIndex        =   9
         Top             =   1375
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ǰ׺�ı�(&T)"
         Height          =   180
         Index           =   3
         Left            =   2670
         TabIndex        =   7
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Index           =   1
         Left            =   345
         TabIndex        =   5
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   345
         TabIndex        =   3
         Top             =   585
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   3030
         TabIndex        =   1
         Top             =   210
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "˵��(&E)"
         Height          =   180
         Index           =   4
         Left            =   345
         TabIndex        =   17
         Top             =   2580
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   1
      Left            =   90
      TabIndex        =   63
      Top             =   165
      Width           =   7575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8130
      TabIndex        =   59
      Top             =   360
      Width           =   1100
   End
End
Attribute VB_Name = "frmPatiCureCardTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------
'��ڲ���
Public Enum gCardTypeEdit
    edT_���� = 0
    edt_�޸� = 1
    edt_ɾ�� = 2
    edt_ͣ�� = 3
    edt_���� = 4
    dt_�鿴 = 5
End Enum
Private mlngModule As Long
Private mEditType As gCardTypeEdit
Private mlngCardTypeID As Long
'-----------------------------------------------------------------------------------------
Private mintSucces As Integer
Private mstrPrivs As String
Private mblnFirst As Boolean
Private Enum mtxtIdx
     idx_���� = 0
     idx_���� = 1
     idx_���� = 2
     idx_ǰ׺�ı� = 3
     idx_���ų��� = 4
     idx_���� = 6
     idx_��ע = 7
     idx_ҽ�ƿ��� = 8
     idx_���� = 9
End Enum
Private Enum mchkIdx
    idx_ȱʡ = 0
    idx_�ϸ���� = 1
    idx_ˢ����ʽ = 2
    'idx_���ƿ� = 3
    idx_�����ʻ� = 4
    idx_�����˿� = 5
    idx_�����ظ�ʹ�� = 6
    idx_�������� = 7
 End Enum
 Private Enum mlblIdx
    idx_lbl���㷽ʽ = 7
 End Enum
 '�����:57326
 Private Enum moptIdx
    idx_������ = 0
    idx_ֻ��һ�ſ� = 1
    idx_�����ſ������� = 2
 End Enum
 
 Private Enum moptLenIdx
   idx_���Ų����ֹ = 0
   idx_���Ų����� = 1
   idx_���Ų������� = 2
End Enum
 Private mbln�̶� As Boolean
 Private Sub SetCtrlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ��ı༭����
    '����:���˺�
    '����:2011-06-28 03:50:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnEdit As Boolean
    Dim blnModify As Boolean
    Dim blnԺ�ڿ� As Boolean
    
    blnԺ�ڿ� = cbo���.Text = "Ժ�ڿ�"
    blnModify = mEditType = edt_�޸� And mbln�̶�
    blnEdit = mEditType = edT_���� Or (mEditType = edt_�޸�)
    
    For i = 0 To txtEdit.UBound
        If i <> 5 Then
            txtEdit(i).Enabled = blnEdit And Not mbln�̶�
            Select Case i
            Case mtxtIdx.idx_���ų���, mtxtIdx.idx_ҽ�ƿ���, mtxtIdx.idx_����
                txtEdit(i).Enabled = IIf(mEditType = edt_�޸�, blnEdit Or blnModify, blnEdit)
             
            End Select
            txtEdit(i).BackColor = IIf(txtEdit(i).Enabled, -2147483643, Me.BackColor)
        End If
    Next
    
    chkEdit(mchkIdx.idx_�����ʻ�).Enabled = blnEdit And Not blnԺ�ڿ�
    chkEdit(mchkIdx.idx_�ϸ����).Enabled = blnEdit And blnԺ�ڿ�
    chkEdit(mchkIdx.idx_�����ظ�ʹ��).Enabled = blnEdit And blnԺ�ڿ�
    chkEdit(mchkIdx.idx_��������).Enabled = blnEdit And Not blnԺ�ڿ�
    chkȱʡ����.Enabled = chkEdit(mchkIdx.idx_��������).Enabled
    '96938:���ϴ�,2016/5/30,ҽ�ƿ��Ƿ񲿷��˿�
    chkEdit(mchkIdx.idx_�����˿�).Enabled = blnEdit And Not blnԺ�ڿ�
    chkEdit(mchkIdx.idx_ˢ����ʽ).Enabled = blnEdit
    chkEdit(mchkIdx.idx_ȱʡ).Enabled = blnEdit
    chk����.Enabled = blnEdit Or blnModify
    '105718:���ϴ���2017/8/16����������ǰ����0����ʾ����
    txt��ʼλ��.Enabled = (blnEdit Or blnModify) And chk����.value = 1
    txt����λ��.Enabled = (blnEdit Or blnModify) And chk����.value = 1
    upd����λ��.Enabled = (blnEdit Or blnModify) And chk����.value = 1
    upd��ʼλ��.Enabled = (blnEdit Or blnModify) And chk����.value = 1
    txt��ʼλ��.BackColor = IIf(upd����λ��.Enabled, -2147483643, Me.BackColor)
    txt����λ��.BackColor = IIf(upd����λ��.Enabled, -2147483643, Me.BackColor)
    cbo���㷽ʽ.Enabled = blnEdit And Not blnԺ�ڿ�
    lblEdit(idx_lbl���㷽ʽ).Enabled = blnEdit And Not blnԺ�ڿ�
    txtPasLen.Enabled = blnEdit Or blnModify
    optPassInput(0).Enabled = blnEdit Or blnModify
    optPassInput(1).Enabled = blnEdit Or blnModify
    optPassInput(2).Enabled = blnEdit Or blnModify
    optRule(0).Enabled = blnEdit Or blnModify
    optRule(1).Enabled = blnEdit Or blnModify
    txtPassInput.Enabled = blnEdit Or blnModify
    cbo���.Enabled = Not mbln�̶� And blnEdit
    chkģ������.Enabled = blnEdit Or blnModify '47522
    '�����;56508
    chkMakeCard.Enabled = chkEdit(2).value <> 1
    chkSendCard.Enabled = Not blnԺ�ڿ�
    chkOpenEnter.Enabled = chkEdit(2).value = 1 And blnEdit
    
    txtPasLen.BackColor = IIf(txtPasLen.Enabled, -2147483643, Me.BackColor)
    txtPassInput.BackColor = IIf(txtPassInput.Enabled, -2147483643, Me.BackColor)
    cbo���㷽ʽ.BackColor = IIf(cbo���㷽ʽ.Enabled, -2147483643, Me.BackColor)
    cbo���.BackColor = IIf(cbo���.Enabled, -2147483643, Me.BackColor)
    cmdSel.Enabled = blnEdit Or blnModify
    chkת�ʼ�����.Enabled = chkEdit(mchkIdx.idx_�����ʻ�).value = 1 And chkEdit(mchkIdx.idx_�����ʻ�).Enabled
    
    '104238:���ϴ���2017/2/15��ҽ�ƿ�������ӷ������ſ���
    OptSendCardLen(0).Enabled = blnEdit Or blnModify
    OptSendCardLen(1).Enabled = blnEdit Or blnModify
    OptSendCardLen(2).Enabled = blnEdit Or blnModify
 End Sub
 Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ���Ч��
    '����:������Ч������true,���򷵻�False
    '����:���˺�
    '����:2011-06-28 03:58:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    For i = 0 To txtEdit.UBound
        If i <> 5 Then
            If i <> mtxtIdx.idx_ҽ�ƿ��� Then
                If i = mtxtIdx.idx_���� Or i = mtxtIdx.idx_���� Then
                    If Trim(txtEdit(i).Text) = "" Then
                        MsgBox txtEdit(i).Tag & " ��������,����", vbOKOnly + vbInformation, gstrSysName
                        If txtEdit(i).Enabled And txtEdit(i).Visible Then txtEdit(i).SetFocus
                        Exit Function
                    End If
                End If
                If zlCommFun.ActualLen(Trim(txtEdit(i).Text)) > txtEdit(i).MaxLength And txtEdit(i).MaxLength <> 0 Then
                    MsgBox txtEdit(i).Tag & " ���������" & txtEdit(i).MaxLength \ 2 & "�����ֻ�" & txtEdit(i).MaxLength & "���ַ�,����", vbOKOnly + vbInformation, gstrSysName
                    If txtEdit(i).Enabled And txtEdit(i).Visible Then txtEdit(i).SetFocus
                    Exit Function
                End If
                If InStr(1, Trim(txtEdit(i).Text), "'") > 0 Then
                    MsgBox txtEdit(i).Tag & " �������뵥����,����", vbOKOnly + vbInformation, gstrSysName
                    If txtEdit(i).Enabled And txtEdit(i).Visible Then txtEdit(i).SetFocus
                    Exit Function
                End If
            End If
        End If
    Next
    If cbo���.Text <> "Ժ�ڿ�" Then
        '������
        If Trim(cbo���㷽ʽ.Text) = "" And chkEdit(mchkIdx.idx_�����ʻ�).value = 1 Then
            MsgBox "ע��:" & vbCrLf & "    �����Ժ�⿨�Ҵ����ʻ���,�������ý��㷽ʽ!", vbInformation + vbOKOnly, gstrSysName
            If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus
            Exit Function
        End If
        '99858:���ϴ�,2016/9/2,�����˻��������ýӿڲ���
        If Trim(txtEdit(mtxtIdx.idx_����).Text) = "" And chkEdit(mchkIdx.idx_�����ʻ�).value = 1 Then
            MsgBox "ע��:" & vbCrLf & "    �����Ժ�⿨�Ҵ����ʻ���,�������ýӿڲ���!", vbInformation + vbOKOnly, gstrSysName
            If txtEdit(mtxtIdx.idx_����).Enabled And txtEdit(mtxtIdx.idx_����).Visible Then txtEdit(mtxtIdx.idx_����).SetFocus
            Exit Function
        End If
     Else
        '����:48090
        If Trim(txtEdit(mtxtIdx.idx_ҽ�ƿ���).Text) = "" Then
           MsgBox "ע��:" & vbCrLf & "    �����Ժ�ڿ�,��������ҽ�ƿ���!", vbInformation + vbOKOnly, gstrSysName
           txtEdit(mtxtIdx.idx_ҽ�ƿ���).SetFocus
           Exit Function
        End If
    End If
    
    If Val(txtPasLen.Text) = 0 Then
        MsgBox "ע��:" & vbCrLf & "    ���볤�Ȳ�������Ϊ��!", vbInformation + vbOKOnly, gstrSysName
        If txtPasLen.Enabled And txtPasLen.Visible Then txtPasLen.SetFocus
        Exit Function
    End If
    If Val(txtPasLen.Text) > 50 Then
        MsgBox "ע��:" & vbCrLf & "    ���볤�Ȳ��ܴ���50!", vbInformation + vbOKOnly, gstrSysName
        If txtPasLen.Enabled And txtPasLen.Visible Then txtPasLen.SetFocus
        Exit Function
    End If
    If optPassInput(2).value Then
        If Val(txtPasLen.Text) < Val(txtPassInput.Text) Then
            MsgBox "ע��:" & vbCrLf & "    ������������볤�Ȳ��ܴ����ܵ����볤��(" & Val(txtPasLen.Text) & ")λ!", vbInformation + vbOKOnly, gstrSysName
            If txtPassInput.Enabled And txtPassInput.Visible Then txtPassInput.SetFocus
            Exit Function
        End If
    End If
    '����:46851
    If Val(txtEdit(mtxtIdx.idx_���ų���).Text) > 50 Then
            MsgBox "ע��:" & vbCrLf & "    �����ֻ������50λ!", vbInformation + vbOKOnly, gstrSysName
            If txtEdit(mtxtIdx.idx_���ų���).Enabled And txtEdit(mtxtIdx.idx_���ų���).Visible Then txtEdit(mtxtIdx.idx_���ų���).SetFocus
            Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 Private Function CheckDepent() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĹ�����
    '����:���ݴ��ڹ���������true,���򷵻�False
    '����:���˺�
    '����:2011-06-28 03:43:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select ���� From ���㷽ʽ Where ���� =7 and nvl(Ӧ����,0)=0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cbo���㷽ʽ
        .Clear
        .AddItem "   "
        .ListIndex = .NewIndex
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!����)
            rsTemp.MoveNext
        Loop
    End With
    CheckDepent = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Function LoadCardData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؿ�Ƭ����
    '����:���سɹ�������true�����򷵻�False
    '����:���˺�
    '����:2011-06-28 02:57:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long, varTemp As Variant
    On Error GoTo errHandle
    Call ClearCtrlData
    mbln�̶� = False
    
    If mEditType = edT_���� Then
        txtEdit(mtxtIdx.idx_����).Text = zlDatabase.GetMax("ҽ�ƿ����", "����", txtEdit(mtxtIdx.idx_����).MaxLength)
        If txtEdit(mtxtIdx.idx_����).Enabled And txtEdit(mtxtIdx.idx_����).Visible Then txtEdit(mtxtIdx.idx_����).SetFocus
        '�����:50172
        txtPassByIDCard.Text = txtPasLen.Text
        '�����;56508
        chkSendCard.value = IIf(chkSendCard.Enabled, 0, 1)
        '�����:57326
        OptSendCard(moptIdx.idx_������).value = 1
        
        OptSendCardLen(moptLenIdx.idx_���Ų����ֹ).value = 1
        
        LoadCardData = True
        Exit Function
    End If
    '�����:57326
    '�����:57697
    '�����:51072
    '�����:56508
    '77872,���ϴ�,2014/10/28:�Ƿ�֧��ת�ʼ�����
    strSQL = "" & _
    "   Select A.Id, A.����, A.����, A.����, A.ǰ׺�ı�, A.���ų���,  nvl(A.ȱʡ��־,0) as ȱʡ��־,  " & _
    "            nvl(A.�Ƿ�̶�,0) as �Ƿ�̶�,  nvl(A.�Ƿ��ϸ����,0)  as  �Ƿ��ϸ����, " & _
    "            nvl(A.�Ƿ�ˢ��,0)  as   �Ƿ�ˢ��,  nvl(A.�Ƿ�����,0)  as    �Ƿ�����," & _
    "            nvl(A.�Ƿ�����ʻ�,0) as   �Ƿ�����ʻ�,  nvl(A.�Ƿ�����,0)  as    �Ƿ�����, " & _
    "           nvl(A.�Ƿ�ȫ��,0)  as    �Ƿ�ȫ��," & _
    "           A.����,A.�ض���Ŀ, A.���㷽ʽ,A.��������,nvl(A.�Ƿ��ظ�ʹ��,0)  as �Ƿ��ظ�ʹ��,  " & _
    "           nvl(A.���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    "           nvl(A.�Ƿ�����,0)  as �Ƿ�����, A.��ע,C.���� as ����,C.ID as ϸĿID,nvl(�Ƿ�ģ������,0) as �Ƿ�ģ������," & _
    "           nvl(A.������������,0) as ������������,nvl(A.�Ƿ�ȱʡ����,0) as �Ƿ�ȱʡ����,nvl(A.�Ƿ��ƿ�,0) as �Ƿ��ƿ�,nvl(A.�Ƿ񷢿�,0) as �Ƿ񷢿�,nvl(A.�Ƿ�д��,0) as �Ƿ�д��, " & _
    "           nvl(A.����,0) as ����,nvl(A.��������,0) as ��������, " & _
    "           nvl(A.�Ƿ�ת�ʼ�����,0) as �Ƿ�ת�ʼ�����, " & _
    "           nvl(A.�Ƿ�ֿ�����,0) as �Ƿ�ֿ�����, " & _
    "           nvl(A.���͵��ýӿ�,0) as ���͵��ýӿ�, " & _
    "           A.�豸�Ƿ����ûس� as ���ûس�,nvl(A.��������,0) as ��������, " & _
    "           Nvl(a.�Ƿ��˿��鿨,0) As �Ƿ��˿��鿨," & _
    "           Nvl(a.�Ƿ�ȱʡ����,0) As �Ƿ�ȱʡ����" & _
    "    From ҽ�ƿ���� A,�շ��ض���Ŀ B,�շ���ĿĿ¼ C" & _
    "    Where  A.ID=[1]  And A.�ض���Ŀ=B.�ض���Ŀ(+) and B.�շ�ϸĿID=C.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ�ҽ�ƿ������Ϣ�������Ѿ�������ɾ����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    With cbo���㷽ʽ
        For i = 0 To .ListCount - 1
            If Trim(.List(i)) = Trim(rsTemp!���㷽ʽ) Then
                .ListIndex = i: i = -1: Exit For
            End If
        Next
        If i >= 0 Then
            .AddItem Nvl(rsTemp!���㷽ʽ): .ListIndex = .NewIndex
        End If
    End With
    
    txtEdit(mtxtIdx.idx_����).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_����).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_����).Text = Nvl(rsTemp!����)
    upd��ʼλ��.Max = IIf(Val(Nvl(rsTemp!���ų���)) = 0, 1, Val(Nvl(rsTemp!���ų���)))
    upd����λ��.Max = upd��ʼλ��.Max
    txtEdit(mtxtIdx.idx_���ų���).Text = IIf(Val(Nvl(rsTemp!���ų���)) = 0, 1, Val(Nvl(rsTemp!���ų���)))
    txtEdit(mtxtIdx.idx_ǰ׺�ı�) = Nvl(rsTemp!ǰ׺�ı�)
    txtEdit(mtxtIdx.idx_��ע) = Nvl(rsTemp!��ע)
    txtEdit(mtxtIdx.idx_����) = Nvl(rsTemp!����)
    'txtEdit(mtxtIdx.idx_�ض���Ŀ) = Nvl(rsTemp!�ض���Ŀ)
    txtEdit(mtxtIdx.idx_ҽ�ƿ���) = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_ҽ�ƿ���).Tag = Val(Nvl(rsTemp!ϸĿID))
    varTemp = Split(Nvl(rsTemp!��������) & "-", "-")
    If Val(varTemp(0)) = 0 Or Val(varTemp(1)) = 0 Then
        upd��ʼλ��.value = IIf(Val(varTemp(0)) = 0, IIf(Val(varTemp(1)) = 0, 1, Val(varTemp(1))), Val(varTemp(0)))
        upd����λ��.value = upd����λ��.Max
        chk����.value = IIf(Val(varTemp(0)) = 0 And Val(varTemp(1)) = 0, 0, 1)
    Else
        upd��ʼλ��.value = Val(varTemp(0))
        upd����λ��.value = Val(varTemp(1))
        chk����.value = 1
    End If
    chkEdit(mchkIdx.idx_�ϸ����).value = IIf(Val(Nvl(rsTemp!�Ƿ��ϸ����)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_ˢ����ʽ).value = IIf(Val(Nvl(rsTemp!�Ƿ�ˢ��)) = 1, 1, 0)
    'chkEdit(mchkIdx.idx_���ƿ�).Value = IIf(Val(Nvl(rsTemp!�Ƿ�����)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_�����˿�).value = IIf(Val(Nvl(rsTemp!�Ƿ�ȫ��)) = 1, 0, 1)
    chkEdit(mchkIdx.idx_�����ʻ�).value = IIf(Val(Nvl(rsTemp!�Ƿ�����ʻ�)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_�����ظ�ʹ��).value = IIf(Val(Nvl(rsTemp!�Ƿ��ظ�ʹ��)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_ȱʡ).value = IIf(Val(Nvl(rsTemp!ȱʡ��־)) = 1, 1, 0)
    chkģ������.value = IIf(Val(Nvl(rsTemp!�Ƿ�ģ������)) = 1, 1, 0)
    '�����;56508
    chkMakeCard.value = IIf(Val(Nvl(rsTemp!�Ƿ��ƿ�)) = 1, 1, 0)
    chkSendCard.value = IIf(Val(Nvl(rsTemp!�Ƿ񷢿�)) = 1, 1, 0)
    chkWriteCard.value = IIf(Val(Nvl(rsTemp!�Ƿ�д��)) = 1, 1, 0)
    
    chkEdit(mchkIdx.idx_��������).value = IIf(Val(Nvl(rsTemp!�Ƿ�����)) = 1, 1, 0)
    If chkEdit(mchkIdx.idx_��������).value = 1 Then
        chkȱʡ����.value = IIf(Val(Nvl(rsTemp!�Ƿ�ȱʡ����)) = 1, 1, 0)
        chkȱʡ����.Enabled = True
    Else
        chkȱʡ����.value = 0
        chkȱʡ����.Enabled = False
    End If
    txtPasLen.Text = Val(Nvl(rsTemp!���볤��))
    For i = 0 To cbo���.ListCount - 1
        If cbo���.List(i) = IIf(Val(Nvl(rsTemp!�Ƿ�����)) = 0, "Ժ�⿨", "Ժ�ڿ�") Then
            cbo���.ListIndex = i: Exit For
        End If
    Next

    Select Case Val(Nvl(rsTemp!���볤������))
    Case 0
            optPassInput(0).value = True
    Case 1
            optPassInput(1).value = True
    Case Else
            optPassInput(2).value = True
            txtPassInput.Text = Abs(Val(Nvl(rsTemp!���볤������)))
    End Select
     optRule(0).value = IIf(Val(Nvl(rsTemp!�������)) = 0, True, False)
     optRule(1).value = IIf(Val(Nvl(rsTemp!�������)) = 1, True, False)
    '�����:51072
    Select Case Val(Nvl(rsTemp!������������))
    Case 0
            optPassConfine(0).value = True
    Case 1
            optPassConfine(1).value = True
    Case Else
            optPassConfine(2).value = True
    End Select
    '�����:50172
    chkPassByIDCard.value = rsTemp!�Ƿ�ȱʡ����
    txtPassByIDCard.Text = txtPasLen.Text
    
    If Val(Nvl(rsTemp!�Ƿ�̶�)) = 1 Then
        '�̶���ֻ�ܲ鿴
        mbln�̶� = True
    End If
    
    '�����:57697
    txtEdit(mtxtIdx.idx_����).Tag = Nvl(rsTemp!����, 0)
    txtEdit(mtxtIdx.idx_����).Text = Get��������(CStr(txtEdit(mtxtIdx.idx_����).Tag))
    
    '�����:57326
    OptSendCard(Val(Nvl(rsTemp!��������))).value = 1
    
    '77872,���ϴ�,2014/9/15:�Ƿ�֧��ת�ʼ�����
    chkת�ʼ�����.Enabled = chkEdit(mchkIdx.idx_�����ʻ�).value = 1
    If chkת�ʼ�����.Enabled Then chkת�ʼ�����.value = IIf(Val(Nvl(rsTemp!�Ƿ�ת�ʼ�����)) = 1, 1, 0)
    chk�ֿ�����.value = IIf(Val(Nvl(rsTemp!�Ƿ�ֿ�����)) = 1, 1, 0)
    chk���͵��ýӿ�.value = IIf(Val(Nvl(rsTemp!���͵��ýӿ�)) = 1, 1, 0)
    chk�˿��鿨.value = IIf(Val(Nvl(rsTemp!�Ƿ��˿��鿨)) = 1, 1, 0)
    
    '103310:���ϴ�,2016/12/6,���ûس������ӿ��ų���
    chkOpenEnter.Enabled = chkEdit(mchkIdx.idx_ˢ����ʽ).value = 1
    chkOpenEnter.value = IIf(Val(Nvl(rsTemp!���ûس�)) = 1, 1, 0)
    
    '104238:���ϴ���2017/2/15��ҽ�ƿ�������ӷ������ſ���
    OptSendCardLen(Val(Nvl(rsTemp!��������))).value = 1
    
    If mEditType = dt_�鿴 Then
        cmdOK.Visible = False
    End If
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function Get��������(str��� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����Ӧ����
    '����:����
    '����:2013-01-29 02:54:36
    '�����:57697
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo Errhand:
        strSQL = "Select ���� From ������� Where ���=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���)
        If rsTemp.EOF = False Then
            Get�������� = rsTemp!����
        End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ClearCtrlData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ�����
    '����:���˺�
    '����:2011-06-28 02:54:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If i <> 5 Then '5-��û��
            txtEdit(i).Text = ""
        End If
    Next
    For i = 0 To chkEdit.UBound
        If i <> 3 Then
            chkEdit(i).value = 0
        End If
    Next
    cbo���㷽ʽ.ListIndex = 0
    chk����.value = 0
End Sub

Private Sub SetDefaultLen()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ�ı༭����
    '����:���˺�
    '����:2011-06-28 02:50:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select  ����, ����, ����, ǰ׺�ı�, ���ų��� ,����,�ض���Ŀ,���㷽ʽ,��ע" & _
    "    From ҽ�ƿ����" & _
    "    Where ID=-1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    txtEdit(mtxtIdx.idx_����).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(mtxtIdx.idx_����).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(mtxtIdx.idx_����).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(mtxtIdx.idx_ǰ׺�ı�).MaxLength = 2 '  rsTemp.Fields("ǰ׺�ı�").DefinedSize
    txtEdit(mtxtIdx.idx_����).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(mtxtIdx.idx_��ע).MaxLength = rsTemp.Fields("��ע").DefinedSize
    txtEdit(mtxtIdx.idx_���ų���).MaxLength = 2
   ' txtEdit(mtxtIdx.idx_�ض���Ŀ).MaxLength = rsTemp.Fields("�ض���Ŀ").DefinedSize
   

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlEditCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As gCardTypeEdit, Optional lngCardTypeID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ�ƿ����༭
    '���:EditType-�༭����
    '        lngCardTypeID-����ʱΪ0
    '����:
    '����:ֻҪ�ɹ�һ��,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-06-27 20:43:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mlngModule = lngModule: mlngCardTypeID = lngCardTypeID
    mintSucces = 0: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    zlEditCard = mintSucces > 0
End Function

Private Sub cbo���_Click()
    Call SetCtrlEnable
End Sub

Private Sub chkEdit_Click(Index As Integer)
'�޸�����:2012-11-29
'�����:56508
    Select Case Index
        Case 2 '�ſ�
            If chkEdit(Index).value = 1 Then
                chkMakeCard.value = 0
                chkMakeCard.Enabled = False
                
                chkOpenEnter.Enabled = True
            Else
                chkMakeCard.Enabled = True
                
                chkOpenEnter.value = 0
                chkOpenEnter.Enabled = False
            End If
        Case 4 '�����ʻ�
            If chkEdit(Index).value = 0 Then
                chkת�ʼ�����.value = 0
                chkת�ʼ�����.Enabled = False
            Else
                chkת�ʼ�����.Enabled = True
            End If
        Case 7
            If chkEdit(Index).value = 0 Then
                chkȱʡ����.value = 0
                chkȱʡ����.Enabled = False
            Else
                chkȱʡ����.Enabled = True
            End If
    End Select
End Sub

''Private Sub chkEdit_Click(Index As Integer)
''    If Index = mchkIdx.idx_���ƿ� Then
''        chkEdit(mchkIdx.idx_�����ʻ�).Enabled = chkEdit(mchkIdx.idx_���ƿ�).Value = 0
''        chkEdit(mchkIdx.idx_��������).Enabled = chkEdit(mchkIdx.idx_�����ʻ�).Enabled
''    End If
''End Sub

Private Sub chk����_Click()
    Dim blnEnable As Boolean
    blnEnable = chk����.Enabled And chk����.value = 1
    txt��ʼλ��.Enabled = blnEnable
    txt����λ��.Enabled = blnEnable
    upd����λ��.Enabled = blnEnable
    upd��ʼλ��.Enabled = blnEnable
    '105718:���ϴ���2017/8/16����������ǰ����0����ʾ����
    txt��ʼλ��.BackColor = IIf(blnEnable, -2147483643, Me.BackColor)
    txt����λ��.BackColor = IIf(blnEnable, -2147483643, Me.BackColor)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsureSel_Click()
    '�����:57697
     If Select���� = False Then Exit Sub
End Sub

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mintSucces = mintSucces + 1
    If mEditType = edT_���� Then
        Call LoadCardData: Exit Sub
    End If
    Unload Me
End Sub
Private Function Select����() As Boolean
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset, strSQL As String
    '�����:57697
    On Error GoTo ErrHandl:
    strSQL = "Select ��� as Id,����,˵��,ҽԺ����,�Ƿ�̶�,�Ƿ��ֹ,��������,ҽ������,���,��Ŀ��ʾ,ҽ���� From �������"
    vRect = GetControlRect(txtEdit(mtxtIdx.idx_����).hWnd)
    lngH = txtEdit(mtxtIdx.idx_����).Height
    sngX = vRect.Left - 15
    sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҽ�Ʊ������", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False)
    If blnCancel = False Then
            txtEdit(mtxtIdx.idx_����).Text = Nvl(rsTemp!����, "")
            txtEdit(mtxtIdx.idx_����).Tag = Nvl(rsTemp!id, "")
    End If
    Select���� = Not blnCancel
    
    Exit Function
ErrHandl:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Select����(ByVal strInput As String) As Boolean
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset, strSQL As String
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
 
    strSQL = "select id,����,����,���㵥λ,˵�� from �շ���ĿĿ¼ where ���='Z' Order by ����"
    vRect = GetControlRect(txtEdit(mtxtIdx.idx_ҽ�ƿ���).hWnd)
    lngH = txtEdit(mtxtIdx.idx_ҽ�ƿ���).Height
    sngX = vRect.Left - 15
    sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҽ�ƿ�������Ŀѡ��", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strInput)
    If blnCancel = True Then
        If txtEdit(mtxtIdx.idx_ҽ�ƿ���).Enabled Then txtEdit(mtxtIdx.idx_ҽ�ƿ���).SetFocus
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_ҽ�ƿ���)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "û���ҵ����������Ŀ�����Ŀ,����!"
        If txtEdit(mtxtIdx.idx_ҽ�ƿ���).Enabled Then txtEdit(mtxtIdx.idx_ҽ�ƿ���).SetFocus
        If UCase(TypeName(txtEdit(mtxtIdx.idx_ҽ�ƿ���))) = UCase("TextBox") Then zlControl.TxtSelAll txtEdit(mtxtIdx.idx_ҽ�ƿ���)
        Exit Function
    End If
    If IsCtrlSetFocus(txtEdit(mtxtIdx.idx_ҽ�ƿ���)) Then txtEdit(mtxtIdx.idx_ҽ�ƿ���).SetFocus
    txtEdit(mtxtIdx.idx_ҽ�ƿ���).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_ҽ�ƿ���).Tag = Nvl(rsTemp!id)
    zlCommFun.PressKey vbKeyTab
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmdSel_Click()
    If Select����("") = False Then Exit Sub
    
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If CheckDepent = False Then Unload Me: Exit Sub
    If LoadCardData = False Then Unload Me: Exit Sub
    Call SetCtrlEnable
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFirst = True
    With cbo���
        .Clear
        .AddItem "Ժ�ڿ�": .ListIndex = .NewIndex
        .AddItem "Ժ�⿨"
    End With
    If mEditType = edT_���� Then chk�ֿ�����.value = 1
    Call SetDefaultLen
End Sub

Private Sub optPassInput_Click(Index As Integer)
    txtPassInput.Enabled = optPassInput(2).value
    txtPassInput.BackColor = IIf(txtPassInput.Enabled, -2147483643, Me.BackColor)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = mtxtIdx.idx_���ų��� Then
        upd����λ��.Max = Val(txtEdit(Index))
        upd��ʼλ��.Max = Val(txtEdit(Index))
        If upd����λ��.value > Val(txtEdit(Index)) Then upd����λ��.value = Val(txtEdit(Index))
        If upd��ʼλ��.value > Val(txtEdit(Index)) Then upd��ʼλ��.value = Val(txtEdit(Index))
    End If
    If Index = mtxtIdx.idx_���� Then
        If Trim(txtEdit(mtxtIdx.idx_����)) = "" And txtEdit(Index).Text <> "" Then txtEdit(mtxtIdx.idx_����) = Left(txtEdit(Index), 1)
    End If
    If Index = mtxtIdx.idx_ҽ�ƿ��� Then
        txtEdit(Index).Tag = ""
    End If
    '�����:57697
    If Index = mtxtIdx.idx_���� Then
        If txtEdit(Index).Text = "" Then
            txtEdit(Index).Tag = ""
        End If
    End If
    
    '�����:54259
    If Index = mtxtIdx.idx_���ų��� Then
       If Val(Trim(txtEdit(mtxtIdx.idx_���ų���).Text)) > 25 Then txtEdit(mtxtIdx.idx_���ų���).Text = 25
    End If
End Sub
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-28 04:13:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngID As Long
    If mEditType = edT_���� Then
        lngID = zlDatabase.GetNextId("ҽ�ƿ����")
    Else
        lngID = mlngCardTypeID
    End If
    
    On Error GoTo errHandle
    ' Zl_ҽ�ƿ����_Update
    strSQL = "Zl_ҽ�ƿ����_Update("
    '  Id_In           In ҽ�ƿ����.ID%Type,
    strSQL = strSQL & "" & lngID & ","
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_����).Text) & "',"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_����).Text) & "',"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_����).Text) & "',"
    '  ǰ׺�ı�_In     In ҽ�ƿ����.ǰ׺�ı�%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_ǰ׺�ı�).Text) & "',"
    '  ���ų���_In     In ҽ�ƿ����.���ų���%Type,
    strSQL = strSQL & "" & Val(txtEdit(mtxtIdx.idx_���ų���).Text) & ","
    '  ȱʡ��־_In     In ҽ�ƿ����.ȱʡ��־%Type,
    strSQL = strSQL & "" & chkEdit(mchkIdx.idx_ȱʡ).value & ","
    '  �Ƿ�̶�_In     In ҽ�ƿ����.�Ƿ�̶�%Type,
    strSQL = strSQL & "0,"
    '  �Ƿ��ϸ����_In In ҽ�ƿ����.�Ƿ��ϸ����%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_�ϸ����).value = 1 And chkEdit(mchkIdx.idx_�ϸ����).Enabled, 1, 0) & ","
    '  �Ƿ�ˢ��_In     In ҽ�ƿ����.�Ƿ�ˢ��%Type,
    strSQL = strSQL & "" & chkEdit(mchkIdx.idx_ˢ����ʽ).value & ","
    '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
    strSQL = strSQL & "" & IIf(cbo���.Text = "Ժ�ڿ�", 1, 0) & ","
    '  �Ƿ�����ʻ�_In In ҽ�ƿ����.�Ƿ�����ʻ�%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_�����ʻ�).Enabled And chkEdit(mchkIdx.idx_�����ʻ�).value = 1, 1, 0) & ","
    '  �Ƿ�ȫ��_In     In ҽ�ƿ����.�Ƿ�ȫ��%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_�����˿�).Enabled And chkEdit(mchkIdx.idx_�����˿�).value = 1, 0, 1) & ","
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_����).Text) & "',"
    '  ��ע_In         In ҽ�ƿ����.��ע%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_��ע).Text) & "',"
    '  �ض���Ŀ_In     In ҽ�ƿ����.�ض���Ŀ%Type,
    If Trim(txtEdit(mtxtIdx.idx_����).Text) = "���￨" Then
        strSQL = strSQL & "'���￨',"
    Else
        strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_����).Text) & "',"
    End If
    '    �շ�ϸĿid_In   In �շ���ĿĿ¼.ID%Type,
    strSQL = strSQL & "" & IIf(Val(txtEdit(mtxtIdx.idx_ҽ�ƿ���).Tag) = 0, "NULL", Val(txtEdit(mtxtIdx.idx_ҽ�ƿ���).Tag)) & ","
    '  ���㷽ʽ_In     In ҽ�ƿ����.���㷽ʽ%Type,
    strSQL = strSQL & "'" & cbo���㷽ʽ.Text & "',"
    '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
    strSQL = strSQL & "1,"
    '  ��������_In     In ҽ�ƿ����.��������%Type,
    strSQL = strSQL & "" & IIf(chk����.value = 1, "'" & upd��ʼλ��.value & "-" & upd����λ��.value & "'", "NULL") & ","
    '  �Ƿ��ظ�ʹ��_In In ҽ�ƿ����.�Ƿ��ظ�ʹ��%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_�����ظ�ʹ��).Enabled And chkEdit(mchkIdx.idx_�����ظ�ʹ��).value = 1, 1, 0) & ","
    '���볤��_In     In ҽ�ƿ����.���볤��%Type,
    strSQL = strSQL & "" & Val(txtPasLen.Text) & ","
    '���볤������_In In ҽ�ƿ����.���볤������%Type,
    If optPassInput(0).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf optPassInput(1).value Then
        strSQL = strSQL & "" & 1 & ","
    Else
        strSQL = strSQL & "" & -1 * Val(txtPassInput.Text) & ","
    End If
    '�������_In     In ҽ�ƿ����.�������%Type,
    If optRule(0).value Then
        strSQL = strSQL & "" & 0 & ","
    Else
        strSQL = strSQL & "" & 1 & ","
    End If
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_��������).Enabled And chkEdit(mchkIdx.idx_��������).value = 1, 1, 0) & ","
    '  ������ʽ_In     In Integer := 0
    strSQL = strSQL & "" & IIf(mEditType = edT_����, 0, 1) & ","
    '�Ƿ�ģ������_In     In ҽ�ƿ����.�Ƿ�ģ������%Type:=0
    strSQL = strSQL & "" & IIf(chkģ������.value = 1, 1, 0) & ","
    '�����:51072
    '������������_In     In ҽ�ƿ����.������������%Type:=0
    If optPassConfine(0).value Then
         strSQL = strSQL & "" & 0 & ","
    ElseIf optPassConfine(1) Then
         strSQL = strSQL & "" & 1 & ","
    ElseIf optPassConfine(2) Then
         strSQL = strSQL & "" & 2 & ","
    End If
    '�Ƿ�ȱʡ����_In     In ҽ�ƿ����.�Ƿ�ȱʡ����%Type:=0
    strSQL = strSQL & "" & IIf(chkPassByIDCard.value, 1, 0) & ","
    '�����:56508
    '�Ƿ��ƿ�_In
    strSQL = strSQL & "" & chkMakeCard & ","
    '�Ƿ񷢿�_In
    strSQL = strSQL & "" & chkSendCard & ","
    '�Ƿ�д��_In
    strSQL = strSQL & "" & chkWriteCard & ","
    '�����:57697
    '����_In
    strSQL = strSQL & "" & IIf(CStr(txtEdit(mtxtIdx.idx_����).Tag) = "", 0, Val(txtEdit(mtxtIdx.idx_����).Tag)) & ","
    '�����:57326
    If OptSendCard(moptIdx.idx_������).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf OptSendCard(moptIdx.idx_ֻ��һ�ſ�).value Then
        strSQL = strSQL & "" & 1 & ","
    ElseIf OptSendCard(moptIdx.idx_�����ſ�������).value Then
        strSQL = strSQL & "" & 2 & ","
    End If
    '77872,���ϴ�,2014/10/28:�Ƿ�֧��ת�ʼ�����
    '�Ƿ�ת�ʼ�����_In  In ҽ�ƿ����.�Ƿ�ת�ʼ�����%Type:=0
    strSQL = strSQL & "" & IIf(chkת�ʼ�����.Enabled And chkת�ʼ�����.value = 1, 1, 0) & ","
    '�Ƿ�ֿ�����_In  In ҽ�ƿ����.�Ƿ�ֿ�����%Type:=0
    strSQL = strSQL & "" & IIf(chk�ֿ�����.Enabled And chk�ֿ�����.value = 1, 1, 0) & ","
    '���͵��ýӿ�_In  In ҽ�ƿ����.���͵��ýӿ�%Type:=0
    strSQL = strSQL & "" & IIf(chk���͵��ýӿ�.Enabled And chk���͵��ýӿ�.value = 1, 1, 0) & ","
    '�豸�Ƿ����ûس�_In  In ҽ�ƿ����.�豸�Ƿ����ûس�%Type:=0
    strSQL = strSQL & "" & IIf(chkOpenEnter.Enabled And chkOpenEnter.value = 1, 1, 0) & ","
    '�������ſ���_In   In ҽ�ƿ����.��������%Type := 0
    If OptSendCardLen(moptLenIdx.idx_���Ų����ֹ).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf OptSendCardLen(moptLenIdx.idx_���Ų�����).value Then
        strSQL = strSQL & "" & 1 & ","
    ElseIf OptSendCardLen(moptLenIdx.idx_���Ų�������).value Then
        strSQL = strSQL & "" & 2 & ","
    End If
    '�Ƿ��˿��鿨_In   In ҽ�ƿ����.�Ƿ��˿��鿨%Type := 0
    strSQL = strSQL & "" & IIf(chk�˿��鿨.Enabled And chk�˿��鿨.value = 1, 1, 0) & ","
    '�Ƿ�ȱʡ����_In   In ҽ�ƿ����.�Ƿ�ȱʡ����%Type := 0
    strSQL = strSQL & "" & IIf(chkȱʡ����.Enabled And chkȱʡ����.value = 1, 1, 0) & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
    Case mtxtIdx.idx_����, mtxtIdx.idx_��ע, mtxtIdx.idx_����
        zlCommFun.OpenIme True
    Case Else
    End Select
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> mtxtIdx.idx_ҽ�ƿ��� Then Exit Sub
    If KeyCode = vbKeyDelete Then txtEdit(Index).Text = ""
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = mtxtIdx.idx_���ų��� Or Index = mtxtIdx.idx_���� Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m����ʽ
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
        zlCommFun.OpenIme False
End Sub
Private Sub txtPasLen_Change()
    optPassInput(1).Caption = "�̶�����" & Val(txtPasLen.Text) & "λ"
    '�����:51072
    txtPassByIDCard.Text = txtPasLen.Text
End Sub
Private Sub upd����λ��_Change()
     If upd����λ��.value < upd��ʼλ��.value Then upd��ʼλ��.value = upd����λ��.value
     If upd��ʼλ��.value = 0 And upd����λ��.value = 0 Then chk����.value = 0
End Sub

Private Sub upd��ʼλ��_Change()
     If upd����λ��.value < upd��ʼλ��.value Then upd����λ��.value = upd��ʼλ��.value
     If upd��ʼλ��.value = 0 And upd����λ��.value = 0 Then chk����.value = 0
End Sub
