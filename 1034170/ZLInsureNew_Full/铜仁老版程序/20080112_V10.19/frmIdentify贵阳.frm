VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt��ע 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1380
      TabIndex        =   62
      Top             =   7080
      Width           =   7995
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "������(&M)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   64
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   63
      Top             =   510
      Width           =   1335
   End
   Begin VB.TextBox txt������Ϣ 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      TabIndex        =   60
      Top             =   6690
      Width           =   7995
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ۼ���Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   180
      TabIndex        =   34
      Top             =   3720
      Width           =   9195
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ�����ת��ʹ�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   58
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ������� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   56
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ����𸶱�׼ 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   54
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ����ۼ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   52
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ����޶� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   50
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txt���֧���ۼ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   48
         Top             =   330
         Width           =   1965
      End
      Begin VB.TextBox txt���ͳ���޶� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   46
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txtͳ��֧���ۼ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   44
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt����ͳ���޶� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   42
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt��֧������ 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   40
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   38
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txtסԺ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   36
         Top             =   330
         Width           =   1965
      End
      Begin VB.Label lbl��ͨ����ҽ�Ʋ�����ת��ʹ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ�����ת��ʹ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4200
         TabIndex        =   57
         Top             =   2340
         Width           =   2730
      End
      Begin VB.Label lbl����Ա���ﲹ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4620
         TabIndex        =   55
         Top             =   1950
         Width           =   2310
      End
      Begin VB.Label lbl����Ա���ﲹ���𸶱�׼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ����𸶱�׼"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4410
         TabIndex        =   53
         Top             =   1560
         Width           =   2520
      End
      Begin VB.Label lbl����Ա���ﲹ���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ����ۼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4830
         TabIndex        =   51
         Top             =   1170
         Width           =   2100
      End
      Begin VB.Label lbl����Ա���ﲹ���޶� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ����޶�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4830
         TabIndex        =   49
         Top             =   780
         Width           =   2100
      End
      Begin VB.Label lbl���֧���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֧���ۼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5670
         TabIndex        =   47
         Top             =   390
         Width           =   1260
      End
      Begin VB.Label lbl���ͳ���޶� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ͳ���޶�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   45
         Top             =   2340
         Width           =   1260
      End
      Begin VB.Label lblͳ��֧���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��֧���ۼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   43
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label lbl����ͳ���޶� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ͳ���޶�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   41
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label lbl��֧������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��֧������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   39
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1020
         TabIndex        =   37
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblסԺ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   35
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   9195
      Begin VB.ComboBox cbo������� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1905
      End
      Begin VB.CheckBox chk������־ 
         Caption         =   "������־"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txt�ɷ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   33
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt�ʻ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   31
         Top             =   2670
         Width           =   2595
      End
      Begin VB.TextBox txt��λ���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   29
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox txt��λ���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   27
         Top             =   1890
         Width           =   1335
      End
      Begin VB.TextBox txt�������� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   25
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox txt���֤�� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   23
         Top             =   1110
         Width           =   2595
      End
      Begin VB.TextBox txt�Ա� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   19
         Top             =   330
         Width           =   2595
      End
      Begin VB.TextBox txt��Ա��� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   15
         Top             =   2670
         Width           =   2595
      End
      Begin VB.TextBox txtҽ���չ���Ⱥ 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   17
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt�����ı�� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox txtҽ���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1890
         Width           =   2595
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1500
         Width           =   1245
      End
      Begin VB.ComboBox cbo֧����� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   1905
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1110
         Width           =   2595
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   1
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lbl�ɷ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɷ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   32
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lbl�ʻ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   30
         Top             =   2730
         Width           =   840
      End
      Begin VB.Label lbl��λ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   28
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label lbl��λ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   26
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   24
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lbl���֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   22
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label lbl�Ա� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5880
         TabIndex        =   20
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5880
         TabIndex        =   18
         Top             =   390
         Width           =   420
      End
      Begin VB.Label lbl��Ա��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   14
         Top             =   2730
         Width           =   840
      End
      Begin VB.Label lblҽ���չ���Ⱥ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ���չ���Ⱥ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   16
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label lbl�����ı��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ı���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   600
         TabIndex        =   12
         Top             =   2340
         Width           =   1050
      End
      Begin VB.Label lblҽ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���˱��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   10
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   7
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label lbl֧����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "֧�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   3
         Top             =   780
         Width           =   840
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   5
         Top             =   1170
         Width           =   420
      End
   End
   Begin VB.Label lbl��ע 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ע"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   750
      TabIndex        =   61
      Top             =   7140
      Width           =   420
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   330
      TabIndex        =   59
      Top             =   6750
      Width           =   840
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstr���� As String
Private mstrҽ���� As String
Private mstr�����ı�� As String
Private mstr������� As String
Private mstr���� As String
Private mstr������ As String
Private mbln������־ As Boolean
Private mblnOK As Boolean
Private int����סԺ��־ As Integer   '����-0,סԺ-1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePassword_Click()
    Dim strNewPass As String
    strNewPass = frm�޸�����.ChangePassword("", Me.txt����.Text, 40)
    If strNewPass <> "" Then mstr������ = strNewPass
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If cmdOK.Enabled = False Then Exit Sub
    If Trim(txt����.Text) = "" Then
        MsgBox "δ��ȷ��ˢ��,����ͨ����֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txtҽ����.Text) = "" Then
        MsgBox "δ��ȷ��ˢ��,����ͨ����֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(mstr������) <> "" Then
        If ��������_������(txt����.Tag, txt����.Text, mstr������) = False Then Exit Sub
        mstr���� = mstr������
        mstr������ = ""
        txt����.Text = mstr����
    End If
    
    '2005.11.22,int����סԺ��־,סԺǿ��ѡ�������
    If (int����סԺ��־ = 1 And cbo�������.ListIndex = 0) Then
       MsgBox "��ѡ�������", vbInformation, gstrSysName
       cbo�������.SetFocus
       Exit Sub
    End If
    
    '�п����޸������룬��ɲ��������֤�󷵻ص�XML���ƻ����ٴε��ö���
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt����.Tag)            ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt����.Text)            ' ����
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo֧�����.ItemData(Me.cbo֧�����.ListIndex))            ' ֧�����
 
    '2005.11.22,int����סԺ��־,ҽ������
    If int����סԺ��־ = 0 Then
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex + 1)
    Else
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex)
    End If
    
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss"))           ' ��ʼʱ��
    '���ýӿ�
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    'mstr���� = Trim(txt����.Text)
    'ҽ���ӿ�δ���ؿ��ţ���ǰ�Ŀ����ֶθ�Ϊ����ſ����ݣ������������Ҫ��
    mstr���� = Me.txt����.Tag
    mstrҽ���� = Trim(txtҽ����.Text)
    mstr�����ı�� = Trim(txt�����ı��.Text)
    mstr���� = Trim(txt����.Text)
    mstr������� = cbo�������.ListIndex + 1
    mbln������־ = (chk������־.Value = 1)
    
    '����˲��˵�ҽ������
'    ҽ����_IN IN ҽ�����˵���.ҽ����%TYPE,
'    סԺ����_IN IN ҽ�����˵���.סԺ����%TYPE,
'    ����_IN IN ҽ�����˵���.����%TYPE,
'    ��֧������_IN IN ҽ�����˵���.��֧������%TYPE,
'    ����ͳ���޶�_IN IN ҽ�����˵���.����ͳ���޶�%TYPE,
'    ͳ��֧���ۼ�_IN IN ҽ�����˵���.ͳ��֧���ۼ�%TYPE,
'    ���ͳ���޶�_IN IN ҽ�����˵���.���ͳ���޶�%TYPE,
'    ���֧���ۼ�_IN IN ҽ�����˵���.���֧���ۼ�%TYPE,
'    ����Ա�����޶�_IN IN ҽ�����˵���.����Ա�����޶�%TYPE,
'    ����Ա�����ۼ�_IN IN ҽ�����˵���.����Ա�����ۼ�%TYPE,
'    ����Ա�𸶱�׼_IN IN ҽ�����˵���.����Ա�𸶱�׼%TYPE,
'    ����Ա��������_IN IN ҽ�����˵���.����Ա��������%TYPE,
'    �μ�75����Ա����_IN IN ҽ�����˵���.�μ�75����Ա����%TYPE)
    On Error GoTo errHand
    gstrSQL = "zl_ҽ�����˵���_INSERT(" & _
        "'" & mstrҽ���� & "'," & Val(txtסԺ����.Text) & "," & Val(txt����.Text) & "," & Val(txt��֧������.Text) & "," & _
        "" & Val(txt����ͳ���޶�.Text) & "," & Val(txtͳ��֧���ۼ�.Text) & "," & Val(txt���ͳ���޶�.Text) & "," & Val(txt���֧���ۼ�.Text) & "," & _
        "" & Val(txt��ͨ����ҽ�Ʋ����޶�.Text) & "," & Val(txt��ͨ����ҽ�Ʋ����ۼ�.Text) & "," & Val(txt��ͨ����ҽ�Ʋ����𸶱�׼.Text) & "," & _
        "" & Val(txt��ͨ����ҽ�Ʋ�������.Text) & ",'" & txt��ͨ����ҽ�Ʋ�����ת��ʹ��.Text & "','" & txt��ע.Text & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetIdentify(ByVal bytType As Byte, str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, _
    Optional ByRef bln������־ As Boolean = False) As Boolean
    mblnOK = False
    mstr���� = ""
    mstr������ = ""
    mbytType = bytType
    
    frmIdentify����.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        str���� = mstr���� & "^" & mstr�������
        strҽ���� = mstrҽ����
        str�����ı�� = mstr�����ı��
        str���� = mstr����
        bln������־ = mbln������־
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '2005.11.22,int����סԺ��־,����cbo�������itemdata
    With cbo֧�����
        .Clear
        If mbytType = 0 Or mbytType = 3 Then
            int����סԺ��־ = 0
            .AddItem "��ͨ����"
            .ItemData(.NewIndex) = 11
            .AddItem "��������"
            .ItemData(.NewIndex) = 18
            With cbo�������
                .Clear
                .AddItem "��ҵְ������ҽ�Ʊ���"
                .AddItem "��ҵ����ҽ�Ʊ���"
                .AddItem "������ҵ��λҽ�Ʊ���"
                .AddItem "��������"
                .ListIndex = 0
            End With
            .ListIndex = 0
         Else
            int����סԺ��־ = 1
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = 31
            With cbo�������
                .Clear
                .AddItem ""
                .AddItem "��ҵְ������ҽ�Ʊ���"
                .AddItem "��ҵ����ҽ�Ʊ���"
                .AddItem "������ҵ��λҽ�Ʊ���"
                .AddItem "��������"
                .ListIndex = 0
            End With
         End If
        .ListIndex = 0
    End With
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim str������Ϣ As String
    
    If Trim(txt����.Text) = "" Then
        MsgBox "��ˢ����", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    '2005.11.22,int����סԺ��־,סԺǿ��ѡ�������
    If (int����סԺ��־ = 1 And cbo�������.ListIndex = 0) Then
       MsgBox "��ѡ�������", vbInformation, gstrSysName
       cbo�������.SetFocus
       Exit Sub
    End If
    
    If InitXML = False Then Exit Sub
    
    '�������޸�����
    If Trim(mstr������) <> "" Then
        If ��������_������(txt����.Text, mstr����, mstr������) = False Then Exit Sub
        mstr���� = mstr������
        mstr������ = ""
        txt����.Text = mstr����
    End If

    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt����.Text)            ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt����.Text)            ' ����
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo֧�����.ItemData(Me.cbo֧�����.ListIndex))            ' ֧�����
        
    '2005.11.22,int����סԺ��־,ҽ������
    If int����סԺ��־ = 0 Then
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex + 1)
    Else
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex)
    End If
    
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss"))           ' ��ʼʱ��
    '���ýӿ�
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    'ȡ�÷���ֵ
    '������Ϣ
    txt����.Tag = txt����.Text                    '���濨�����ݣ��Ա��������ʱʹ��
    'txt����.Text = GetElemnetValue("CARDID")
    txtҽ����.Text = GetElemnetValue("PERSONCODE")
    txt�����ı��.Text = GetElemnetValue("CENTERCODE")
    txtҽ���չ���Ⱥ.Text = IIf(Val(GetElemnetValue("CAREPSNFLAG")) = 0, "��", "��")
    
    '2005.11.22,int����סԺ��־,סԺ��������ѡ�������,Ĭ���ÿ�
    If int����סԺ��־ = 0 Then
        cbo�������.ListIndex = GetElemnetValue("INSURETYPE") - 1
    Else
'        cbo�������.ListIndex = 0
    End If
    
    txt��Ա���.Text = GetElemnetValue("PERSONTYPE")
    txt��Ա���.Text = Switch(txt��Ա���.Text = "11", "��ְ", txt��Ա���.Text = "21", "����" _
                      , txt��Ա���.Text = "32", "ʡ������", txt��Ա���.Text = "34", "��������", True, "����")
    txt����.Text = GetElemnetValue("PERSONNAME")
    txt�Ա�.Text = GetElemnetValue("SEX")
    txt�Ա�.Text = Switch(txt�Ա�.Text = "1", "��", txt�Ա�.Text = "2", "Ů", txt�Ա�.Text = "9", "����", True, txt�Ա�.Text)
    txt���֤��.Text = GetElemnetValue("PID")
    txt��������.Text = GetElemnetValue("BIRTHDAY")
    txt��λ����.Text = GetElemnetValue("DEPTCODE")
    txt��λ����.Text = GetElemnetValue("DEPTNAME")
    txt�ʻ����.Text = GetElemnetValue("ACCTBALANCE")
    '�ۼ���Ϣ
    txtסԺ����.Text = GetElemnetValue("HOSPTIMES")
    txt����.Text = GetElemnetValue("STARTFEE")
    txt��֧������.Text = GetElemnetValue("STARTFEEPAID")
    txt����ͳ���޶�.Text = GetElemnetValue("FUND1LMT")
    txtͳ��֧���ۼ�.Text = GetElemnetValue("FUND1PAID")
    txt���ͳ���޶�.Text = GetElemnetValue("FUND2LMT")
    txt���֧���ۼ�.Text = GetElemnetValue("FUND2PAID")
    txt��ͨ����ҽ�Ʋ����޶�.Text = GetElemnetValue("FUND3LMT")
    txt��ͨ����ҽ�Ʋ����ۼ�.Text = GetElemnetValue("FUND3PAID")
    txt��ͨ����ҽ�Ʋ����𸶱�׼.Text = GetElemnetValue("STARTFEE2STD")
    txt��ͨ����ҽ�Ʋ�������.Text = GetElemnetValue("STARTFEE2")
    txt��ͨ����ҽ�Ʋ�����ת��ʹ��.Text = GetElemnetValue("FUND75BALANCE")
    txt��ע.Text = GetElemnetValue("NOTE")
    txt������Ϣ.Text = GetElemnetValue("LOCKINFO")

    cmdOK.Enabled = True
End Sub
