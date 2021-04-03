VERSION 5.00
Begin VB.Form frmOutStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   Icon            =   "frmOutStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkȱʡҩ�� 
      Caption         =   "����ҽ���´�ǿ��ȱʡҩ��"
      Height          =   240
      Left            =   4350
      TabIndex        =   24
      Top             =   3180
      Width           =   2580
   End
   Begin VB.Frame fraEPR 
      Caption         =   "��������"
      Height          =   1905
      Left            =   4350
      TabIndex        =   53
      Top             =   3495
      Width           =   4455
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ��Ӧ"
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   33
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   32
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   31
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�������"
         Height          =   195
         Index           =   2
         Left            =   2970
         TabIndex        =   30
         Top             =   885
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "����������ʾ"
         Height          =   195
         Left            =   105
         TabIndex        =   34
         Top             =   1545
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   1650
         TabIndex        =   35
         Top             =   1455
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ⱦ��"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   29
         Top             =   885
         Width           =   840
      End
      Begin VB.TextBox txtNotifyEPR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "10"
         Top             =   330
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   585
         TabIndex        =   55
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPRDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   585
         TabIndex        =   54
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtNotifyEPRDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "1"
         Top             =   600
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Σ��ֵ"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   28
         Top             =   885
         Width           =   840
      End
      Begin VB.CheckBox chkNotifyEPR 
         Caption         =   "ÿ    �����Զ�ˢ�����������е�����"
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   345
         Width           =   3900
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   360
         TabIndex        =   57
         Top             =   885
         Width           =   810
      End
      Begin VB.Label lblNotifyEPRDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ������ɵ�������ʾ����������"
         Height          =   180
         Left            =   375
         TabIndex        =   56
         Top             =   615
         Width           =   3060
      End
   End
   Begin VB.CheckBox chkStaKB 
      Caption         =   "������Ļ����"
      Height          =   255
      Left            =   330
      TabIndex        =   21
      Top             =   5070
      Width           =   1665
   End
   Begin VB.Frame fraBespeak 
      Caption         =   "ԤԼ�Һŵ���ӡ��ʽ"
      Height          =   2160
      Left            =   4350
      TabIndex        =   52
      Top             =   135
      Width           =   1920
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   2
         Left            =   315
         TabIndex        =   8
         Top             =   1575
         Width           =   1380
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   6
         Top             =   450
         Width           =   900
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   7
         Top             =   1020
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.Frame fraReception 
      Caption         =   "���˽������"
      Height          =   2145
      Left            =   6420
      TabIndex        =   49
      Top             =   150
      Width           =   2265
      Begin VB.OptionButton optMode 
         Caption         =   "����ֹ"
         Height          =   240
         Index           =   0
         Left            =   645
         TabIndex        =   9
         Top             =   585
         Width           =   1005
      End
      Begin VB.OptionButton optMode 
         Caption         =   "��ֹ"
         Height          =   240
         Index           =   1
         Left            =   645
         TabIndex        =   10
         Top             =   900
         Width           =   855
      End
      Begin VB.OptionButton optMode 
         Caption         =   "��ʾ"
         Height          =   240
         Index           =   2
         Left            =   645
         TabIndex        =   11
         Top             =   1230
         Width           =   750
      End
      Begin VB.TextBox txtReceptionTime 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   930
         TabIndex        =   12
         Text            =   "0"
         Top             =   1575
         Width           =   525
      End
      Begin VB.Label lblReceptionMode 
         Caption         =   "���Ʒ�ʽ"
         Height          =   270
         Left            =   135
         TabIndex        =   51
         Top             =   330
         Width           =   825
      End
      Begin VB.Line line 
         X1              =   840
         X2              =   1545
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Label lblReceptionTime 
         AutoSize        =   -1  'True
         Caption         =   "������ǰ       ���ӽ���"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   1590
         Width           =   2070
      End
   End
   Begin VB.OptionButton optAdd 
      Caption         =   "��������,�л���ҽ��ʱ����ҽ��"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   615
      TabIndex        =   18
      Top             =   4035
      Width           =   3015
   End
   Begin VB.CheckBox chk���к���� 
      Caption         =   "ҽ���������к�������ڶ����н���"
      Height          =   195
      Left            =   330
      TabIndex        =   20
      Top             =   4785
      Value           =   1  'Checked
      Width           =   3360
   End
   Begin VB.CheckBox chk�������ﲡ�� 
      Caption         =   "ҽ�������������ƺ����ﲡ��"
      Height          =   180
      Left            =   4350
      TabIndex        =   22
      Top             =   2580
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CheckBox chk�Һ�ˢ�� 
      Caption         =   "�Һű���ˢ����ȡ����"
      Height          =   255
      Left            =   330
      TabIndex        =   19
      Top             =   4410
      Width           =   2640
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   1605
      TabIndex        =   48
      Top             =   3045
      Width           =   465
   End
   Begin VB.TextBox txtQueuePatis 
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
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "3"
      ToolTipText     =   "��ʾ����ҽ������ܺ��ж��ٸ�����������,�����󣬾Ͳ����ٴκ���;�˲�����Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪҽ������������Ч"
      Top             =   2880
      Width           =   465
   End
   Begin VB.CheckBox chkAutoAdd 
      Caption         =   "���˽�����Զ�����"
      Height          =   195
      Left            =   330
      TabIndex        =   16
      Top             =   3525
      Width           =   2640
   End
   Begin VB.CheckBox chk�Զ����� 
      Caption         =   "���ҵ����ﲡ��֮���Զ�����"
      Height          =   195
      Left            =   4350
      TabIndex        =   23
      Top             =   2880
      Width           =   2640
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   300
      TabIndex        =   36
      Top             =   5550
      Width           =   1500
   End
   Begin VB.CheckBox chkPrice 
      Caption         =   "����Һŷ���ͨ�����۵��շ�"
      Height          =   195
      Left            =   330
      TabIndex        =   15
      Top             =   3180
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.Frame fraLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   735
      TabIndex        =   46
      Top             =   2685
      Width           =   465
   End
   Begin VB.TextBox txtRefresh 
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
      IMEMode         =   3  'DISABLE
      Left            =   735
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "180"
      ToolTipText     =   "���ˢ�¼��Ϊ 30 �룬����Ϊ 0 ��ʾ���Զ�ˢ��"
      Top             =   2505
      Width           =   465
   End
   Begin VB.Frame Frame2 
      Caption         =   " ������� "
      Height          =   2190
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   4155
      Begin VB.CommandButton cmdYS 
         Caption         =   "��"
         Height          =   255
         Left            =   3645
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   1755
         Width           =   255
      End
      Begin VB.TextBox txt����ҽ�� 
         Height          =   300
         Left            =   1020
         TabIndex        =   5
         Top             =   1725
         Width           =   2910
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "ֻ�����Ѿ�����Ĳ���"
         Height          =   195
         Left            =   1020
         TabIndex        =   4
         Top             =   1365
         Width           =   2100
      End
      Begin VB.ComboBox cbo���� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2910
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   255
         Left            =   3645
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   690
         Width           =   255
      End
      Begin VB.ComboBox cbo��Χ 
         ForeColor       =   &H80000012&
         Height          =   300
         ItemData        =   "frmOutStationSetup.frx":000C
         Left            =   1020
         List            =   "frmOutStationSetup.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "����Ĳ��˷�Χ"
         Top             =   1005
         Width           =   2910
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   2
         Top             =   660
         Width           =   2910
      End
      Begin VB.Label lblEditDept 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   255
         TabIndex        =   0
         Top             =   360
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   45
         X2              =   4090
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label lblҽ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ��"
         Height          =   180
         Left            =   240
         TabIndex        =   42
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   255
         TabIndex        =   39
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl��Χ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ﷶΧ"
         Height          =   180
         Left            =   225
         TabIndex        =   41
         Top             =   1065
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   45
         X2              =   4090
         Y1              =   1635
         Y2              =   1635
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   38
      Top             =   5550
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   37
      Top             =   5550
      Width           =   1100
   End
   Begin VB.OptionButton optAdd 
      Caption         =   "����ҽ��,�л�������ʱ��������"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   615
      TabIndex        =   17
      Top             =   3765
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.Label lblQueuePatis 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������ܺ���      ��"
      Height          =   180
      Left            =   330
      TabIndex        =   47
      ToolTipText     =   "��ʾ����ҽ������ܺ��ж��ٸ�����������,�����󣬾Ͳ����ٴκ���;�˲�����Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪҽ������������Ч"
      Top             =   2880
      Width           =   1980
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   -15
      X2              =   9000
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6780
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Label lblRefresh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ��      ���Զ�ˢ�º���/ת�ﲡ���嵥"
      Height          =   180
      Left            =   345
      TabIndex        =   44
      ToolTipText     =   "���ˢ�¼��Ϊ 30 �룬����Ϊ 0 ��ʾ���Զ�ˢ��"
      Top             =   2520
      Width           =   3330
   End
End
Attribute VB_Name = "frmOutStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Private mstrLike As String
Private mbln�ҺŰ��� As Boolean '���ݲ������Һ��Ű�ģʽ  ȷ������ѡ��Χ��true�°棬false�ϰ�

Private Enum Enum_chkWarn
    chkDΣ��ֵ = 0
    chkD��Ⱦ�� = 1
    chkD������� = 2
    chkD��Ѫ��� = 3
    chkD��Ѫ��� = 4
    chkD��Ѫ��Ӧ = 5
End Enum

Private Sub cbo��Χ_Click()
    '���˺Ż򱾿���ʱ
    chk����.Visible = cbo��Χ.ListIndex = 0 Or cbo��Χ.ListIndex = 2
End Sub


Private Sub chkAutoAdd_Click()
    If chkAutoAdd.Value = 1 Then
        optAdd(0).Enabled = True
        optAdd(1).Enabled = True
    Else
        optAdd(0).Enabled = False
        optAdd(1).Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdOK_Click()
    Dim str���˽������ As String '�����:57566
    Dim blnHavePara As Boolean  '�Ƿ��в�������Ȩ��
    Dim i As Integer
    Dim strTmp As String
    
    If txt����.Text = "" Then
        MsgBox "������ҽ�������ҡ�", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If txt����ҽ��.Text = "" Then
        MsgBox "�����ҽ����", vbInformation, gstrSysName
        txt����ҽ��.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex < 0 Then
        MsgBox "������ұ���ѡ��,����", vbInformation + vbOKOnly, gstrSysName
        cbo����.SetFocus
        Exit Sub
    End If
    blnHavePara = InStr(1, ";" & mstrPrivs & ";", ";��������;") > 0
    
    Call zlDatabase.SetPara("��������", Me.txt����.Text, glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("���ﷶΧ", Me.cbo��Χ.ItemData(Me.cbo��Χ.ListIndex), glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("����ҽ��", Me.txt����ҽ��.Text, glngSys, p����ҽ��վ, blnHavePara)
    '����:38603
    Call zlDatabase.SetPara("�Һű���ˢ��", chk�Һ�ˢ��.Value, glngSys, p����ҽ��վ, blnHavePara)
    
    '���˺�:Ӧ�����Ŷӽкŵĺ����˴�:��Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪ���������ŶӺ��ж���=2ʱ��Ч
    If txtQueuePatis.Enabled Then
        Call zlDatabase.SetPara("ҽ����������", Val(Me.txtQueuePatis.Text), glngSys, p����ҽ��վ, blnHavePara)
    End If
    '�������
    Call zlDatabase.SetPara("�������", cbo����.ItemData(cbo����.ListIndex), glngSys, p����ҽ��վ, blnHavePara)
    
    'ֻ�����Ѿ�����Ĳ���
    Call zlDatabase.SetPara("ֻ�����Ѿ�����Ĳ���", chk����.Value, glngSys, p����ҽ��վ, blnHavePara)

    '���ﲡ��ˢ�¼��
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
    Call zlDatabase.SetPara("����ˢ�¼��", Val(txtRefresh.Text), glngSys, p����ҽ��վ, blnHavePara)
    
    '�Һŷ��ò�ͨ�����۵��շ�
    Call zlDatabase.SetPara("����ҺŻ��۵�", chkPrice.Value, glngSys, p����ҽ��վ, blnHavePara)
    
    '�ҵ����˺��Զ�����
    Call zlDatabase.SetPara("�ҵ����˺��Զ�����", chk�Զ�����.Value, glngSys, p����ҽ��վ, blnHavePara)
    '������Զ�����
    If optAdd(1).Value And optAdd(1).Enabled Then
        Call zlDatabase.SetPara("������Զ�����", 2, glngSys, p����ҽ��վ, blnHavePara)
    Else
        Call zlDatabase.SetPara("������Զ�����", chkAutoAdd.Value, glngSys, p����ҽ��վ, blnHavePara)
    End If
    '����:44250
    Call zlDatabase.SetPara("��������������", chk�������ﲡ��.Value, glngSys, p����ҽ��վ, blnHavePara)
    'ҽ���������к���������
    Call zlDatabase.SetPara("ҽ���������к���������", chk���к����.Value, glngSys, p����ҽ��վ, blnHavePara)
    '������Ļ����
    Call zlDatabase.SetPara("������Ļ����", chkStaKB.Value, glngSys, p����ҽ��վ, blnHavePara)
    '�����:57566
    If optMode(0).Value = True Then
        str���˽������ = "0|0"
    ElseIf optMode(1).Value = True Then
        str���˽������ = "1|" & Nvl(txtReceptionTime.Text, "0")
    ElseIf optMode(2).Value = True Then
        str���˽������ = "2|" & Nvl(txtReceptionTime.Text, "0")
    End If
    zlDatabase.SetPara "���˽������", str���˽������, glngSys, p����ҽ��վ, blnHavePara
    
    '56274
    For i = 0 To optPrintBespeak.UBound
        If optPrintBespeak(i).Value Then
            zlDatabase.SetPara "ԤԼ�Һŵ���ӡ��ʽ", i, glngSys, p����ҽ��վ, blnHavePara
            Exit For
        End If
    Next
    
    Call zlDatabase.SetPara("����ҽ���´�ǿ��ȱʡҩ��", chkȱʡҩ��.Value, glngSys, p����ҽ���´�, blnHavePara)
    
    Call zlDatabase.SetPara("�Զ�ˢ�²������ļ��", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("�Զ�ˢ�²�����������", Val(txtNotifyEPRDay.Text), glngSys, p����ҽ��վ, blnHavePara)
    strTmp = ""
    For i = chkDΣ��ֵ To chkD��Ѫ��Ӧ
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zlDatabase.SetPara("�Զ�ˢ������", strTmp, glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("����������ʾ", chkSound.Value, glngSys, p����ҽ��վ, blnHavePara)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    If txt����.Tag <> txt���� Then Exit Sub '��txt���ҵ�Validate�¼�����
    
    If mbln�ҺŰ��� Then
        strSQL = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
            " From �������� A, �����������ÿ��� B, ������Ա C, �ϻ���Ա�� D" & vbNewLine & _
            " Where a.Id = b.����id And b.����id = c.����id And c.��Աid = d.��Աid" & vbNewLine & _
            "       And d.�û��� = User And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)"
    Else
        strSQL = "Select Distinct e.���� As ID,e.����,e.����" & vbNewLine & _
               "From �������� E, �ҺŰ������� D, �ҺŰ��� C, ������Ա A, �ϻ���Ա�� B" & vbNewLine & _
               "Where a.��Աid = b.��Աid And b.�û��� = User And c.����id = a.����id And c.Id = d.�ű�id And e.���� = d.�������� " & _
               " And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null)"
    End If
    '���û�в��ҵ����ݣ����ȡ�����е��������ҹ�ѡ��
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select a.Id, a.����, a.���� From �������� A Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)"
    End If

    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��������", , , , , , , txt����.Left, txt����.Top, txt����.Height, , , True)
    If Not rsTmp Is Nothing Then
        txt����.Tag = rsTmp("����"): txt���� = txt����.Tag
        If cbo��Χ.Enabled And cbo��Χ.Visible Then cbo��Χ.SetFocus
    End If
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 0)
End Sub

Private Sub cmdYS_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnCanle As Boolean
    If txt����ҽ��.Tag <> txt����ҽ�� Then Exit Sub '��txtҽ����Validate�¼�����
            
    strSQL = "Select Distinct A.��� as ID,A.���� as ����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And B.����ID=D.����ID" & _
        " And C.��Ա����||''='ҽ��' And D.������� IN(1,3) And D.��������||''='�ٴ�'" & _
        " And B.����ID In (Select ����ID From ������Ա Where ��ԱID=[1])" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, False, 0, 0, txt����ҽ��.Height, blnCanle, False, True, UserInfo.ID)
    If blnCanle Then Exit Sub
    If Not rsTmp Is Nothing Then txt����ҽ��.Tag = rsTmp("����"): txt����ҽ�� = txt����ҽ��.Tag: Me.cmdOK.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim blnSetup As Boolean
    Dim i As Long
    Dim str���˽������ As String  '�����:57566
    Dim intType As Integer
    Dim strNotify As String
    Dim str���� As String
    
    blnSetup = InStr(1, ";" & mstrPrivs & ";", ";��������;") > 0
    gblnOK = False
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    
    On Error Resume Next
    str���� = zlDatabase.GetPara("��������", glngSys, p����ҽ��վ, "", Array(lbl����, txt����, cmdSel), blnSetup)
    On Error GoTo 0
    
    On Error GoTo errH
    '��ȡ����ȱʡ���ҷ�Χ
    strPar = zlDatabase.GetPara("�������", glngSys, p����ҽ��վ, "", Array(lblEditDept, cbo����), blnSetup)
    
    strSQL = "Select Distinct B.ID,B.����,B.����,A.ȱʡ" & _
        " From ������Ա A,���ű� B,��������˵�� C" & _
        " Where A.����ID=B.ID And B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And A.��ԱID=[1]" & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        cbo����.AddItem rsTmp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
        If rsTmp!ID = Val(strPar) Then
            cbo����.ListIndex = cbo����.NewIndex
        ElseIf Nvl(rsTmp!ȱʡ, 0) = 1 And cbo����.ListIndex = -1 Then
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTmp.MoveNext
    Next
    Me.cbo��Χ.ListIndex = Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2", Array(lbl��Χ, cbo��Χ), blnSetup)) - 1
    
    strSQL = "Select 1 From �������� E where e.����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    If Not rsTmp.EOF Then
        txt����.Text = str����
        txt����.Tag = str����
    End If
    
    '����ѡ�������ҽ������Ĳ��˽��о���
    If InStr(mstrPrivs, "���ﲡ��") > 0 Then
        '����ѡ�񱾿����µ�ҽ��
        cmdYS.Enabled = True
        txt����ҽ��.Enabled = True
    Else
        cmdYS.Enabled = False
        txt����ҽ��.Enabled = False
    End If
    txt����ҽ��.Tag = zlDatabase.GetPara("����ҽ��", glngSys, p����ҽ��վ, UserInfo.����, Array(lblҽ��, txt����ҽ��, cmdYS), blnSetup)
    txt����ҽ��.Text = txt����ҽ��.Tag
    
    '����:38603
    chk�Һ�ˢ��.Value = IIf(Val(zlDatabase.GetPara("�Һű���ˢ��", glngSys, p����ҽ��վ, "0", Array(chk�Һ�ˢ��), blnSetup)) = 1, 1, 0)
    '���˺�:Ӧ�����Ŷӽкŵĺ����˴�:��Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪ���������ŶӺ���վ��=1ʱ��Ч
    txtQueuePatis.Text = Val(zlDatabase.GetPara("ҽ����������", glngSys, p����ҽ��վ, 3, Array(lblQueuePatis, txtQueuePatis), blnSetup))
    If txtQueuePatis.Enabled Then
        txtQueuePatis.Enabled = CheckDoctorPatisIsValid
    End If
    
    'ֻ�����Ѿ�����Ĳ���
    chk����.Value = Val(zlDatabase.GetPara("ֻ�����Ѿ�����Ĳ���", glngSys, p����ҽ��վ, , Array(chk����), blnSetup))
    
    '���ﲡ��ˢ�¼��
    txtRefresh.Text = Val(zlDatabase.GetPara("����ˢ�¼��", glngSys, p����ҽ��վ, 180, Array(lblRefresh, txtRefresh), blnSetup))
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
    
    '�Һŷ��ò�ͨ�����۵��շ�
    chkPrice.Value = Val(zlDatabase.GetPara("����ҺŻ��۵�", glngSys, p����ҽ��վ, 1, Array(chkPrice), blnSetup))
    
    '�ҵ����˺��Զ�����
    chk�Զ�����.Value = Val(zlDatabase.GetPara("�ҵ����˺��Զ�����", glngSys, p����ҽ��վ, , Array(chk�Զ�����), blnSetup))
    
    '������Զ�����
    strPar = Val(zlDatabase.GetPara("������Զ�����", glngSys, p����ҽ��վ, , Array(chkAutoAdd, optAdd(0), optAdd(1)), blnSetup))
    If strPar = 2 Then
        chkAutoAdd.Value = 1
        optAdd(1).Value = True
    Else
        chkAutoAdd.Value = strPar
    End If
    '����:44250
    chk�������ﲡ��.Value = Val(zlDatabase.GetPara("��������������", glngSys, p����ҽ��վ, 1, Array(chk�������ﲡ��), blnSetup))
    'ҽ���������к���������
    chk���к����.Value = Val(zlDatabase.GetPara("ҽ���������к���������", glngSys, p����ҽ��վ, 1, Array(chk���к����), blnSetup))
    '������Ļ����
    chkStaKB.Value = Val(zlDatabase.GetPara("������Ļ����", glngSys, p����ҽ��վ, , Array(chkStaKB), blnSetup))
    
    '�����:57566
    '���˽������
    str���˽������ = zlDatabase.GetPara("���˽������", glngSys, p����ҽ��վ, , Array(optMode(0), optMode(1), optMode(2), txtReceptionTime, lblReceptionMode, lblReceptionTime), blnSetup)
    If str���˽������ <> "" Then
        If Split(str���˽������, "|")(0) = "0" Then
            optMode(0).Value = True
        ElseIf Split(str���˽������, "|")(0) = "1" Then
            optMode(1).Value = True
        ElseIf Split(str���˽������, "|")(0) = "2" Then
            optMode(2).Value = True
        End If
        txtReceptionTime.Text = Split(str���˽������ & "|", "|")(1)
    End If
    '����:56274
    i = Val(zlDatabase.GetPara("ԤԼ�Һŵ���ӡ��ʽ", glngSys, p����ҽ��վ, 1, Array(optPrintBespeak(0), optPrintBespeak(1), optPrintBespeak(2)), blnSetup))
    If i <= optPrintBespeak.UBound Then optPrintBespeak(i).Value = True
    
    '��Ϣ����ˢ��
    strPar = zlDatabase.GetPara("�Զ�ˢ�²������ļ��", glngSys, p����ҽ��վ, , Array(chkNotifyEPR), blnSetup, intType)
    If Val(strPar) > 0 Then
        chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
    End If
 
    If (intType = 3 Or intType = 15) And Not blnSetup Then
        txtNotifyEPR.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("�Զ�ˢ�²�����������", glngSys, p����ҽ��վ, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), blnSetup)
    txtNotifyEPRDay.Text = Val(strPar)
        
    strNotify = zlDatabase.GetPara("�Զ�ˢ������", glngSys, p����ҽ��վ, , Array(chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), lblArea), blnSetup)
    chkWarn(chkDΣ��ֵ).Value = Val(Mid(strNotify, 1, 1))
    chkWarn(chkD��Ⱦ��).Value = Val(Mid(strNotify, 2, 1))
    chkWarn(chkD�������).Value = Val(Mid(strNotify, 3, 1))
    chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 4, 1))
    chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
    chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 5, 1))
    chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
    chkWarn(chkD��Ѫ��Ӧ).Value = Val(Mid(strNotify, 6, 1))
    chkWarn(chkD��Ѫ��Ӧ).Visible = gblnѪ��ϵͳ
    If InStr(mstrPrivs, "��������") = 0 Then
        chkWarn(chkDΣ��ֵ).Enabled = False
        chkWarn(chkD��Ⱦ��).Enabled = False
        chkWarn(chkD�������).Enabled = False
        chkWarn(chkD��Ѫ���).Enabled = False
        chkWarn(chkD��Ѫ���).Enabled = False
        chkWarn(chkD��Ѫ��Ӧ).Enabled = False
    End If
    chkSound.Value = Val(zlDatabase.GetPara("����������ʾ", glngSys, p����ҽ��վ, , Array(chkSound, cmdSoundSet), blnSetup))

    '����ҽ���´�ǿ��ȱʡҩ��
    chkȱʡҩ��.Value = Val(zlDatabase.GetPara("����ҽ���´�ǿ��ȱʡҩ��", glngSys, p����ҽ���´�, "1", Array(chkȱʡҩ��), blnSetup))

    strPar = ""
    mbln�ҺŰ��� = False
    strPar = zlDatabase.GetPara(256, glngSys) & "|"
    If 0 <> Val(Split(strPar, "|")(0)) Then
        If Split(strPar, "|")(1) <> "" Then
            strPar = Format(Split(strPar, "|")(1), "YYYY-MM-DD")
            If Format(zlDatabase.Currentdate, "YYYY-MM-DD") >= strPar Then
                mbln�ҺŰ��� = True
            End If
        End If
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optMode_Click(Index As Integer)
    '�����:57566
    Select Case Index
        Case 0
            txtReceptionTime.Text = 0: txtReceptionTime.Enabled = False
        Case Else
            txtReceptionTime.Enabled = True And InStr(mstrPrivs, ";��������;") > 0
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
End Sub

Private Sub txtRefresh_GotFocus()
    Call zlControl.TxtSelAll(txtRefresh)
End Sub

Private Sub txtRefresh_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRefresh_Validate(Cancel As Boolean)
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
End Sub

Private Sub txt����ҽ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean

    If txt����ҽ��.Tag = txt����ҽ�� Then Exit Sub

    strSQL = "Select Distinct A.��� as ID,A.���� as ����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And B.����ID=D.����ID" & _
        " And C.��Ա����||''='ҽ��' And D.������� IN(1,3) And D.��������||''='�ٴ�'" & _
        " And B.����ID In(Select ����ID From ������Ա Where ��ԱID=[1])" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (Upper(A.���) Like [2] Or Upper(A.����) Like [3] Or Upper(A.����) Like [3])" & _
        " Order by A.����"
        
    vRect = GetControlRect(txt����ҽ��.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt����ҽ��.Height, blnCancel, False, True, UserInfo.ID, UCase(txt����ҽ��.Text) & "%", mstrLike & UCase(txt����ҽ��.Text) & "%")
    If Not rsTmp Is Nothing Then
        txt����ҽ��.Tag = rsTmp("����")
        txt����ҽ�� = txt����ҽ��.Tag
    Else
        txt����ҽ��.Tag = ""
        txt����ҽ�� = ""
        Cancel = blnCancel
    End If
End Sub

Private Sub txt����ҽ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ҽ��)
End Sub

Private Sub txt����ҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt����ҽ�� = "" Then txt����ҽ��.Tag = "1"
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt���� = "" Then txt����.Tag = "1"
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If txt����.Tag = txt���� Then Exit Sub
    
    If mbln�ҺŰ��� Then
        strSQL = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
            " From �������� A, �����������ÿ��� B, ������Ա C, �ϻ���Ա�� D" & vbNewLine & _
            " Where a.Id = b.����id And b.����id = c.����id And c.��Աid = d.��Աid" & vbNewLine & _
            " And (Upper(a.����) Like [1] Or Upper(a.����) Like [2] Or Upper(a.����) Like [2])" & _
            "       And d.�û��� = User And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)"
    Else
        strSQL = "Select Distinct e.���� As ID,e.����,e.����" & vbNewLine & _
                "From �������� E, �ҺŰ������� D, �ҺŰ��� C, ������Ա A, �ϻ���Ա�� B" & vbNewLine & _
                "Where a.��Աid = b.��Աid And b.�û��� = User And c.����id = a.����id And c.Id = d.�ű�id And e.���� = d.�������� " & _
                " And (Upper(E.����) Like [1] Or Upper(E.����) Like [2] Or Upper(E.����) Like [2])" & _
                " And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) "
    End If
        
    '���û�в��ҵ����ݣ����ȡ�����е��������ҹ�ѡ��
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(txt����.Text) & "%", mstrLike & UCase(txt����.Text) & "%")
    If rsTmp.EOF Then
        strSQL = "Select a.Id, a.����, a.���� From �������� A Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)" & _
            " And (Upper(a.����) Like [1] Or Upper(a.����) Like [2] Or Upper(a.����) Like [2])"
    End If

    vRect = GetControlRect(txt����.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt����.Height, blnCancel, False, True, UCase(txt����.Text) & "%", mstrLike & UCase(txt����.Text) & "%")
    If Not rsTmp Is Nothing Then
        txt����.Tag = rsTmp("����")
        txt���� = txt����.Tag
    Else
        txt����.Tag = ""
        txt���� = ""
        Cancel = blnCancel
    End If
End Sub
Private Sub txtReceptionTime_GotFocus()
    zlControl.TxtSelAll txtReceptionTime
End Sub
Private Sub txtReceptionTime_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtReceptionTime_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtReceptionTime, KeyAscii, m����ʽ
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
