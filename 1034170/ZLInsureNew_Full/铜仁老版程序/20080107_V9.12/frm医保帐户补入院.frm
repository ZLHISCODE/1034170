VERSION 5.00
Begin VB.Form frmҽ���ʻ�����Ժ 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˲���ҽ����Ժ�Ǽ�"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmҽ���ʻ�����Ժ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "�Ǽ�(&X)"
      Height          =   350
      Left            =   6090
      TabIndex        =   8
      Top             =   6015
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "��������Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   705
      Left            =   75
      TabIndex        =   27
      Top             =   1815
      Width           =   8745
      Begin VB.TextBox txt������� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtԤ����� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
         Height          =   180
         Left            =   2370
         TabIndex        =   30
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
         Height          =   180
         Left            =   375
         TabIndex        =   28
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   6765
         TabIndex        =   34
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   4695
         TabIndex        =   32
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "��������Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   3345
      Left            =   75
      TabIndex        =   58
      Top             =   2580
      Width           =   8745
      Begin VB.TextBox txtҽ�Ƹ��� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt�������� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   570
         Width           =   1140
      End
      Begin VB.TextBox txt��ϵ�˹�ϵ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1890
         Width           =   2000
      End
      Begin VB.TextBox txt��� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txtְҵ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt����״�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtѧ�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt�����ص� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txt��ͥ��ַ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1230
         Width           =   3150
      End
      Begin VB.TextBox txt�����ʱ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1170
      End
      Begin VB.TextBox txt��ϵ������ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1170
      End
      Begin VB.TextBox txt��ϵ�˵�ַ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1890
         Width           =   3225
      End
      Begin VB.TextBox txt��ϵ�˵绰 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2220
         Width           =   2000
      End
      Begin VB.TextBox txt������λ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2220
         Width           =   3225
      End
      Begin VB.TextBox txt��λ�绰 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2550
         Width           =   2000
      End
      Begin VB.TextBox txt��λ�ʱ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1170
      End
      Begin VB.TextBox txt��λ������ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txt��λ�ʺ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3225
      End
      Begin VB.TextBox txt��ͥ�绰 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txt���֤�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   900
         Width           =   3150
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ���"
         Height          =   180
         Left            =   345
         TabIndex        =   80
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   6570
         TabIndex        =   79
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Left            =   4470
         TabIndex        =   78
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   345
         TabIndex        =   77
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   4830
         TabIndex        =   76
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   2685
         TabIndex        =   75
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4830
         TabIndex        =   74
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2685
         TabIndex        =   73
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��"
         Height          =   180
         Left            =   6930
         TabIndex        =   72
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lvl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Left            =   345
         TabIndex        =   71
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   180
         Left            =   345
         TabIndex        =   70
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl��ͥ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Left            =   345
         TabIndex        =   69
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl�����ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʱ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   68
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl��ϵ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ������"
         Height          =   180
         Left            =   4290
         TabIndex        =   67
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˹�ϵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˹�ϵ"
         Height          =   180
         Left            =   165
         TabIndex        =   66
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ"
         Height          =   180
         Left            =   4290
         TabIndex        =   65
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵绰"
         Height          =   180
         Left            =   165
         TabIndex        =   64
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         Height          =   180
         Left            =   4470
         TabIndex        =   63
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lbl��λ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   180
         Left            =   345
         TabIndex        =   62
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl��λ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   61
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl��λ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ������"
         Height          =   180
         Left            =   165
         TabIndex        =   60
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label lbl��λ�ʺ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʺ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   59
         Top             =   2940
         Width           =   720
      End
   End
   Begin VB.Frame fra��Ժ��Ϣ 
      Caption         =   "��סԺ��Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   8730
      Begin VB.TextBox txt��� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1260
         Width           =   7320
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "��֤(&V)"
         Height          =   285
         Left            =   6330
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F12(ҽ��������֤)"
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt��Ժʱ�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txtҽ���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt�ѱ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         Left            =   1125
         TabIndex        =   2
         Top             =   225
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   330
         TabIndex        =   82
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4800
         TabIndex        =   23
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2715
         TabIndex        =   21
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   6540
         TabIndex        =   25
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   690
         TabIndex        =   19
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Height          =   180
         Left            =   4620
         TabIndex        =   5
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2535
         TabIndex        =   3
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   690
         TabIndex        =   11
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2715
         TabIndex        =   13
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4800
         TabIndex        =   15
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   6900
         TabIndex        =   17
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl����ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   510
         TabIndex        =   1
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   780
      TabIndex        =   10
      Top             =   6015
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7380
      TabIndex        =   9
      Top             =   6015
      Width           =   1100
   End
End
Attribute VB_Name = "frmҽ���ʻ�����Ժ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long 'Ҫ�޸Ļ�鿴�Ĳ���ID
Private mlng��ҳID As Long 'Ҫ�޸Ļ�鿴����ҳID
Private mstrҽ���� As String

Private Function ReadCard() As Boolean
'���ܣ���ȡָ��������Ϣ,����ʾ�ڽ�����
    Dim rstmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
        
    gstrSQL = "Select * From ������Ϣ Where ����ID=" & mlng����ID
    rstmp.CursorLocation = adUseClient
    
    Call OpenRecordset(rstmp, Me.Caption)
    
    If rstmp.EOF Then Exit Function
    If rstmp.RecordCount <> 1 Then Exit Function
    
    'סԺ��Ϣ
    txt����ID.Locked = True
    txt����ID.Text = mlng����ID
    txt����ID.Locked = False
    
    txt����.Text = rstmp!����
    txtסԺ��.Text = IIf(IsNull(rstmp!סԺ��), "", rstmp!סԺ��)
    
    '������Ϣ
    txt�Ա�.Text = IIf(IsNull(rstmp!�Ա�), "", rstmp!�Ա�)
    txt����.Text = IIf(IsNull(rstmp!����), "", rstmp!����)
    txt�ѱ�.Text = IIf(IsNull(rstmp!�ѱ�), "", rstmp!�ѱ�)
    txtҽ�Ƹ���.Text = IIf(IsNull(rstmp!ҽ�Ƹ��ʽ), "", rstmp!ҽ�Ƹ��ʽ)
    txt����.Text = IIf(IsNull(rstmp!����), "", rstmp!����)
    txt����.Text = IIf(IsNull(rstmp!����), "", rstmp!����)
    txtѧ��.Text = IIf(IsNull(rstmp!ѧ��), "", rstmp!ѧ��)
    txt����״��.Text = IIf(IsNull(rstmp!����״��), "", rstmp!����״��)
    txtְҵ.Text = IIf(IsNull(rstmp!ְҵ), "", rstmp!ְҵ)
    txt���.Text = IIf(IsNull(rstmp!���), "", rstmp!���)
    txt��������.Text = Format(IIf(IsNull(rstmp!��������), "", rstmp!��������), "yyyy-MM-dd")
    txt���֤��.Text = IIf(IsNull(rstmp!���֤��), "", rstmp!���֤��)
    txt�����ص�.Text = IIf(IsNull(rstmp!�����ص�), "", rstmp!�����ص�)
    txt��ͥ��ַ.Text = IIf(IsNull(rstmp!��ͥ��ַ), "", rstmp!��ͥ��ַ)
    txt��ͥ�绰.Text = IIf(IsNull(rstmp!��ͥ�绰), "", rstmp!��ͥ�绰)
    txt�����ʱ�.Text = IIf(IsNull(rstmp!�����ʱ�), "", rstmp!�����ʱ�)
    txt��ϵ������.Text = IIf(IsNull(rstmp!��ϵ������), "", rstmp!��ϵ������)
    txt��ϵ�˹�ϵ.Text = IIf(IsNull(rstmp!��ϵ�˹�ϵ), "", rstmp!��ϵ�˹�ϵ)
    txt��ϵ�˵�ַ.Text = IIf(IsNull(rstmp!��ϵ�˵�ַ), "", rstmp!��ϵ�˵�ַ)
    txt��ϵ�˵绰.Text = IIf(IsNull(rstmp!��ϵ�˵绰), "", rstmp!��ϵ�˵绰)
    txt������λ.Text = IIf(IsNull(rstmp!������λ), "", rstmp!������λ)
    txt��λ�绰.Text = IIf(IsNull(rstmp!��λ�绰), "", rstmp!��λ�绰)
    txt��λ�ʱ�.Text = IIf(IsNull(rstmp!��λ�ʱ�), "", rstmp!��λ�ʱ�)
    txt��λ������.Text = IIf(IsNull(rstmp!��λ������), "", rstmp!��λ������)
    txt��λ�ʺ�.Text = IIf(IsNull(rstmp!��λ�ʺ�), "", rstmp!��λ�ʺ�)
        
    '������Ϣ
    txt������.Text = IIf(IsNull(rstmp!������), "", rstmp!������)
    txt������.Text = Format(IIf(IsNull(rstmp!������), "", rstmp!������), "0.00")
    
    gstrSQL = "Select * From ������� Where ����=1 And ����ID=" & mlng����ID
    Call OpenRecordset(rstmp, Me.Caption)
    
    If Not rstmp.EOF Then
        txt�������.Text = Format(IIf(IsNull(rstmp!�������), 0, rstmp!�������), "0.00")
        txtԤ�����.Text = Format(IIf(IsNull(rstmp!Ԥ�����), 0, rstmp!Ԥ�����), "0.00")
    End If
    
    
    '����ҽ����Ϣ
    txtҽ����.Text = ""
    mstrҽ���� = ""
    
    
    '������ҳ��Ϣ
    gstrSQL = "Select A.��Ժ����,A.��Ժ����,b.���� as ��Ժ����,C.���� as ����ȼ�" & _
              " From ������ҳ A,���ű� B,����ȼ� C" & _
              " Where A.����ID=" & mlng����ID & " And A.��ҳID=" & mlng��ҳID & _
              "       and A.��Ժ����ID=B.ID and A.����ȼ�ID=C.���(+) "
    Call OpenRecordset(rstmp, Me.Caption)
    
    txt����.Text = rstmp!��Ժ����
    txt����.Text = IIf(IsNull(rstmp!����ȼ�), "��", rstmp!����ȼ�)
    txt����.Text = IIf(IsNull(rstmp!��Ժ����), "", rstmp!��Ժ����)
    txt��Ժʱ��.Text = Format(rstmp!��Ժ����, "yyyy-MM-dd HH:mm")
    
    '��Ժ���
    gstrSQL = "Select ������Ϣ" & _
              " From ������" & _
              " Where ����ID=" & mlng����ID & " And ��ҳID=" & mlng��ҳID & " and �������=1 "
    Call OpenRecordset(rstmp, Me.Caption)
    If rstmp.EOF = False Then
        txt���.Text = NVL(rstmp("������Ϣ"))
    End If
    
    Dim objInsure As New clsInsure
    If objInsure.GetCapability(support����¼��������) = True Then
        txt���.Locked = False
        txt���.BackColor = txt����ID.BackColor
    Else
        txt���.Locked = True
        txt���.BackColor = txtסԺ��.BackColor
    End If
    
    ReadCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
'������Ժ�Ǽ�
    Dim clsInsure As New clsInsure
    
    If mlng����ID = 0 Then
        MsgBox "����ȷ���Ȳ�����Ժ�ǼǵĲ��ˡ�", vbInformation, gstrSysName
        txt����ID.SetFocus
        Exit Sub
    End If
    If mstrҽ���� = "" Then
        MsgBox "������֤�ò����Ƿ���Խ���ҽ����Ժ��", vbInformation, gstrSysName
        cmdYB.SetFocus
        Exit Sub
    End If
    If txt���.Locked = False And txt���.Text = "" Then
        MsgBox "����д��Ժ��ϡ�", vbInformation, gstrSysName
        txt���.SetFocus
        Exit Sub
    End If
    If zlCommFun.StrIsValid(txt���.Text, txt���.MaxLength, txt���.hwnd, "��Ժ���") = False Then
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    gstrSQL = "zl_������ҳ_����ҽ����Ժ(" & mlng����ID & "," & mlng��ҳID & "," & gintInsure & ",'" & txt��� & "')"
    ExecuteProcedure Me.Caption
    
    If clsInsure.ComeInSwap(mlng����ID, mlng��ҳID, mstrҽ����) = False Then
        '�Ǽ�ʧ��
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    
    gcnOracle.CommitTrans
    MsgBox "����" & txt����.Text & "����ҽ����Ժ�ɹ���" & IIf(gintInsure > 900, vbCrLf & "���˷�����ϸ��ҽ�������Ѿ���ҽ���������㡣", "") _
        , vbInformation, gstrSysName
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdYB_Click()
'��֤ҽ���������
    Dim lng����ID As Long
    Dim strYBPati As String
    Dim clsInsure As New clsInsure
    Dim arr��Ϣ As Variant
    
    If mlng����ID = 0 Then
        MsgBox "����ȷ���Ȳ�����Ժ�ǼǵĲ��ˡ�", vbInformation, gstrSysName
        txt����ID.SetFocus
        Exit Sub
    End If
    lng����ID = mlng����ID
    strYBPati = clsInsure.Identify(1, lng����ID)
    If lng����ID <> 0 Then mlng����ID = lng����ID
    
    arr��Ϣ = Split(strYBPati, ";")
    '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,...
    If UBound(arr��Ϣ) >= 8 Then
        txtҽ����.Text = arr��Ϣ(1)
        mstrҽ���� = txtҽ����.Text
        
        txt����.Text = arr��Ϣ(3)
        txt�Ա�.Text = arr��Ϣ(4)
        txt��������.Text = arr��Ϣ(5)
        txt���֤��.Text = arr��Ϣ(6)
        
        cmdOK.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdYB_Click
    End If
End Sub

Private Sub Form_Load()
    mlng����ID = 0
    mlng��ҳID = 0
End Sub

Private Sub txt����ID_Change()
    If txt����ID.Locked = False Then
        mlng����ID = 0
        mlng��ҳID = 0
    End If
End Sub

Private Sub txt����ID_GotFocus()
    zlControl.TxtSelAll txt����ID
End Sub

Private Sub txt����ID_KeyPress(KeyAscii As Integer)
    Dim lng����ID  As Long
    
    'ת���ɴ�д(���ֲ��ɴ���)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 And (txt����ID.Text = "" Or txt����ID.SelLength = Len(txt����ID.Text)) Then
        txt����ID.MaxLength = 15
    End If
    
    If Len(Trim(Me.txt����ID.Text)) = 0 And KeyAscii = 13 Then
        If frmҽ������ѡ��.Get����(lng����ID) = True Then
            txt����ID.Text = "A" & lng����ID
        End If
    End If
    Me.Refresh
    
    'ˢ����ϻ���������س�
    If (KeyAscii = 13 And Trim(txt����ID.Text) <> "") Then
        If Val(txt����ID.Text) = mlng����ID And mlng����ID > 0 Then
            If mstrҽ���� = "" Then
                cmdYB.SetFocus
            Else
                cmdOK.SetFocus
            End If
            Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txt����ID.Text = txt����ID.Text & Chr(KeyAscii)
            txt����ID.SelStart = Len(txt����ID.Text)
        End If
        KeyAscii = 0
        
        If Not GetPatient() Then
            MsgBox "û�з��ָò��˵�סԺ��Ϣ,���������룡", vbInformation, gstrSysName
            txt����ID.Text = ""
            txt����ID.SetFocus
            Exit Sub
        Else
            Call ReadCard
            cmdYB.SetFocus
        End If
    End If

End Sub

Private Function GetPatient() As Boolean
'���ܣ���ȡ������Ϣ
'����:�Ƿ��ȡ�ɹ�,�ɹ�ʱrsInfo�а���������Ϣ,ʧ��ʱrsInfo=Close
    Dim rsInfo As New ADODB.Recordset
    Dim strCode As String
    
    strCode = Trim(txt����ID.Text)
    On Error GoTo errH
    
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID And Nvl(A.סԺ����,0)=C.��ҳID And A.����ID=" & Val(Mid(strCode, 2)) & _
            "       and C.���� is null and C.��Ժ���� is null"
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID And Nvl(A.סԺ����,0)=C.��ҳID And A.סԺ��=" & Val(Mid(strCode, 2)) & _
            "       and C.���� is null and C.��Ժ���� is null"
    Else '��������
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID And Nvl(A.סԺ����,0)=C.��ҳID And A.����='" & strCode & _
            "'       and C.���� is null and C.��Ժ���� is null"
    End If
    
    rsInfo.CursorLocation = adUseClient
    Call OpenRecordset(rsInfo, Me.Caption)
    
    '��ȡʧ��
    If rsInfo.EOF Then
        Exit Function
    End If
        
    mlng����ID = rsInfo("����ID")
    mlng��ҳID = rsInfo("��ҳID")
    
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


