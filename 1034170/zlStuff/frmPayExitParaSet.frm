VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayExitParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   Icon            =   "frmPayExitParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8100
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Caption         =   "  ���� "
      Height          =   1272
      Index           =   2
      Left            =   5160
      TabIndex        =   18
      Top             =   960
      Width           =   2850
      Begin VB.CheckBox chk����ʱ�� 
         Caption         =   "����ҽ��������ʱ�����"
         Height          =   345
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   2340
      End
      Begin VB.CheckBox chkSign 
         Caption         =   "������ǩ��"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   1485
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����ʱ�Զ������ʷ�������"
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   255
         Width           =   2520
      End
   End
   Begin VB.Frame fra�豸���� 
      Caption         =   "  ���ܿ��������豸���� "
      Height          =   735
      Left            =   75
      TabIndex        =   23
      Top             =   5520
      Width           =   7950
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   240
         TabIndex        =   24
         Top             =   250
         Width           =   1500
      End
   End
   Begin VB.Frame fra 
      Caption         =   "  ���Ͽ��� "
      Height          =   1545
      Index           =   3
      Left            =   75
      TabIndex        =   19
      Top             =   2400
      Width           =   7950
      Begin VB.CheckBox chk���ܷ��� 
         Caption         =   "����ʱ�������������¼"
         Height          =   180
         Left            =   2400
         TabIndex        =   36
         Top             =   275
         Width           =   2655
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "Ӫ��"
         Enabled         =   0   'False
         Height          =   180
         Index           =   6
         Left            =   2280
         TabIndex        =   31
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   5
         Left            =   1440
         TabIndex        =   30
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   4
         Left            =   600
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   3
         Left            =   3240
         TabIndex        =   28
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "���"
         Enabled         =   0   'False
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   27
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   26
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "�ٴ�"
         Enabled         =   0   'False
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   25
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkSendByNo 
         Caption         =   "�����ݺŷ���"
         Height          =   180
         Left            =   5160
         TabIndex        =   22
         Top             =   240
         Width           =   2130
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����������ʱ�����ǲ��˿��ҿ����ļ�¼"
         Height          =   420
         Left            =   240
         TabIndex        =   21
         Top             =   500
         Width           =   4650
      End
      Begin VB.CheckBox Chk�Ƿ��Զ�ȱ�ϼ�� 
         Caption         =   "�Ƿ��Զ�ȱ�ϼ��"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   275
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  ҵ������ "
      Height          =   1272
      Left            =   75
      TabIndex        =   17
      Top             =   960
      Width           =   3090
      Begin VB.ComboBox cbo�շѴ��� 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   840
         Width           =   2280
      End
      Begin VB.CheckBox chkҵ�� 
         Caption         =   "�շѵ�(&S)"
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chkҵ�� 
         Caption         =   "���ʵ�(&J)"
         Height          =   285
         Index           =   1
         Left            =   1850
         TabIndex        =   1
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chkҵ�� 
         Caption         =   "���ʱ�(&B)"
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1150
      End
      Begin VB.Label lbl�շѴ��� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�շѴ���"
         Height          =   420
         Left            =   120
         TabIndex        =   34
         Top             =   825
         Width           =   465
      End
      Begin VB.Label lbl�������� 
         Caption         =   "��������"
         Height          =   420
         Left            =   120
         TabIndex        =   32
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   6480
      Width           =   8775
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   705
      Width           =   8775
   End
   Begin VB.Frame fra 
      Caption         =   "  ��ӡ��Ʊ������ "
      Height          =   1305
      Index           =   1
      Left            =   75
      TabIndex        =   14
      Top             =   4080
      Width           =   7950
      Begin VB.ComboBox cbo���Ϻ� 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   2280
      End
      Begin VB.ComboBox cbo���Ϻ� 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Width           =   2280
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "Ʊ�ݴ�ӡ����"
         Height          =   360
         Left            =   3360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   750
         Width           =   1875
      End
      Begin VB.ComboBox cboƱ������ 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   780
         Width           =   2280
      End
      Begin VB.Label lbl���ϵ� 
         AutoSize        =   -1  'True
         Caption         =   "���ϵ�"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lbl���ϵ� 
         AutoSize        =   -1  'True
         Caption         =   "���ϵ�"
         Height          =   180
         Left            =   3360
         TabIndex        =   40
         Top             =   420
         Width           =   540
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
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "  ȱʡ��λ "
      Height          =   1272
      Index           =   0
      Left            =   3240
      TabIndex        =   13
      Top             =   960
      Width           =   1770
      Begin VB.CheckBox chk��λ 
         Caption         =   "��װ��λ(&2)"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   660
         Width           =   1452
      End
      Begin VB.CheckBox chk��λ 
         Caption         =   "ɢװ��λ(&1)"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   348
         Width           =   1452
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   11
      Top             =   6735
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6570
      TabIndex        =   10
      Top             =   6735
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5385
      TabIndex        =   9
      Top             =   6735
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7344
      Top             =   -48
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
            Picture         =   "frmPayExitParaSet.frx":030A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayExitParaSet.frx":08A4
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayExitParaSet.frx":0E3E
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayExitParaSet.frx":13D8
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmPayExitParaSet.frx":1972
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "��������ѡ��Ŀ,������صĴ�ӡ�����ϵ�λ�����Ʊ�ݵ����á�"
      Height          =   390
      Index           =   0
      Left            =   735
      TabIndex        =   12
      Top             =   390
      Width           =   7125
   End
End
Attribute VB_Name = "frmPayExitParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mblnExit As Boolean
Private mlngModule As Long
Private mstrPrivs As String
Private mblnHavePriv As Boolean
Private Const mstrAllType As String = "�ٴ�,����,���,����,����,����,Ӫ��"


Private Sub cboƱ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkDeptType_Click(Index As Integer)
    Dim n As Integer
    Dim blnAllUnselect As Boolean
    
    '����Ҫѡ��һ��
    blnAllUnselect = True
    For n = 0 To chkDeptType.Count - 1
        If chkDeptType(n).Value = 1 Then
            blnAllUnselect = False
            Exit For
        End If
    Next
    If blnAllUnselect = True Then
        chkDeptType(Index).Value = 1
    End If
End Sub

Private Sub chk����_Click()
    Dim n As Integer

    For n = 0 To chkDeptType.Count - 1
        chkDeptType(n).Enabled = (chk����.Value = 1)
        If chk����.Tag = "0" Then
            chkDeptType(n).Value = 1
        End If
    Next
End Sub



Private Sub chk��ӡ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk��λ_Click(Index As Integer)
    
    If chk��λ(Index).Value = 1 Then
        chk��λ(IIf(Index = 1, 0, 1)).Value = 0
    Else
        chk��λ(IIf(Index = 1, 0, 1)).Value = 1
    End If
End Sub

Private Sub chk��λ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub



 
 
Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkҵ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub CmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1723)
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2007/12/24
    '------------------------------------------------------------------------------------------
    Dim strҵ������ As String
    Dim str�������� As String
    Dim n As Integer
    
    strҵ������ = IIf(chkҵ��(0).Value = 1, "24", "0")
    strҵ������ = strҵ������ & IIf(chkҵ��(1).Value = 1, ",25", ",0")
    strҵ������ = strҵ������ & IIf(chkҵ��(2).Value = 1, ",26", ",0")
    
    '������ҩ
    If chk����.Value = 0 Then
        str�������� = ""
    Else
        For n = 0 To chkDeptType.Count - 1
            If chkDeptType(n).Value = 0 Then
                str�������� = IIf(str�������� = "", "", str�������� & ",") & chkDeptType(n).Caption
            End If
        Next
        If str�������� = "" Then
            str�������� = mstrAllType
        End If
    End If
    
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
   
    Call zlDatabase.SetPara("���ϴ�ӡ���ѷ�ʽ", cbo���Ϻ�.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("���ϴ�ӡ���ѷ�ʽ", cbo���Ϻ�.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("��ѯҵ������", strҵ������, glngSys, mlngModule)
    Call zlDatabase.SetPara("���ĵ�λ", IIf(chk��λ(1).Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("�Զ�����", IIf(chk����.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("ȱ�ϼ��", IIf(Chk�Ƿ��Զ�ȱ�ϼ��.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("�������Ϸ�ʽ", str��������, glngSys, mlngModule)
    Call zlDatabase.SetPara("�����ݺŷ���", chkSendByNo.Value, glngSys, mlngModule)
    Call zlDatabase.SetPara("�շѴ�����ʾ��ʽ", cbo�շѴ���.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("����ʱ�����������ʼ�¼", chk���ܷ���.Value, glngSys, mlngModule)
    '59655
    Call zlDatabase.SetPara("������ǩ��", IIf(chkSign.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("����ҽ��������ʱ�����", IIf(chk����ʱ��.Value = 1, 1, 0), glngSys, mlngModule)

    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub CmdOK_Click()
    If SaveSet = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    
    If cboƱ������.ListIndex < 0 Then
        ShowMsgBox "�����ú�Ʊ��!"
        cboƱ������.SetFocus
    End If
    Select Case cboƱ������.ListIndex
    Case 0
        '���ݴ�ӡ
        strBill = "ZL1_BILL_1723"
    Case 1
        '�嵥��ӡ
        strBill = "ZL1_BILL_1723_1"
    Case 2
        '��������֪ͨ��
        strBill = "ZL1_BILL_1723_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Long
    Dim strArr As Variant
    Dim str�������� As String
    Dim BlnSelect As Boolean
    Dim n As Integer
    Dim int�շѴ��� As Integer
    
    mblnHavePriv = IsHavePrivs(mstrPrivs, "��������")
    
    With cbo�շѴ���
        .Clear
        .AddItem "1-��ʾ���еĴ���"
        .AddItem "2-����ʾ���շѴ���"
        .AddItem "3-����ʾδ�շѴ���"
        .ListIndex = 0
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-���Ĵ�����"
        .AddItem "2-��ӡ�ѷ����嵥"
        .AddItem "3-����֪ͨ����ӡ"
        .ListIndex = 0
    End With
    
    With Me.cbo���Ϻ�
        .AddItem "1-���Ϻ���ʾ�Ƿ��ӡ"
        .AddItem "2-���Ϻ��Զ���ӡ"
        .AddItem "3-���Ϻ󲻴�ӡ"
        .ListIndex = 0
    End With
    
    With Me.cbo���Ϻ�
        .AddItem "1-���Ϻ���ʾ�Ƿ��ӡ"
        .AddItem "2-���Ϻ��Զ���ӡ"
        .AddItem "3-���Ϻ󲻴�ӡ"
        .ListIndex = 0
    End With
    
    chk����.Value = IIf(Val(zlDatabase.GetPara("�Զ�����", glngSys, mlngModule, , Array(chk����), mblnHavePriv)) = 1, 1, 0)
  
    Chk�Ƿ��Զ�ȱ�ϼ��.Value = IIf(Val(zlDatabase.GetPara("ȱ�ϼ��", glngSys, mlngModule, , Array(Chk�Ƿ��Զ�ȱ�ϼ��), mblnHavePriv)) = 1, 1, 0)
    str�������� = zlDatabase.GetPara("�������Ϸ�ʽ", glngSys, mlngModule, "", Array(chk����, chkDeptType(0), chkDeptType(1), chkDeptType(2), chkDeptType(3), chkDeptType(4), chkDeptType(5), chkDeptType(6)), mblnHavePriv)
        
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0", Array(chk��λ(0), chk��λ(1), fra(0)), mblnHavePriv))
    chk��λ(0).Value = 0
    chk��λ(1).Value = 0
    If Val(strReg) = 0 Then
        chk��λ(0).Value = 1
    Else
        chk��λ(1).Value = 1
    End If
      
    strReg = Trim(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0", Array(cbo���Ϻ�), mblnHavePriv))
    If Val(strReg) >= 0 And Val(strReg) <= 2 Then
        cbo���Ϻ�.ListIndex = Val(strReg)
    Else
        cbo���Ϻ�.ListIndex = 0
    End If
    
    strReg = Trim(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0", Array(cbo���Ϻ�), mblnHavePriv))
    If Val(strReg) >= 0 And Val(strReg) <= 2 Then
        cbo���Ϻ�.ListIndex = Val(strReg)
    Else
        cbo���Ϻ�.ListIndex = 0
    End If
    
    strReg = Trim(zlDatabase.GetPara("��ѯҵ������", glngSys, mlngModule, "", Array(lbl��������, chkҵ��(0), chkҵ��(1), chkҵ��(2), Frame3), mblnHavePriv))
    If strReg = "" Then strReg = "24,25,26"
    strArr = Split(strReg & "," & "," & ",", ",")
    For i = 0 To UBound(strArr)
        If i > 2 Then Exit For
        chkҵ��(i).Value = IIf(Val(strArr(i)) > 0, 1, 0)
    Next
    
    chkSendByNo.Value = IIf(Val(zlDatabase.GetPara("�����ݺŷ���", glngSys, mlngModule, , Array(chkSendByNo), mblnHavePriv)) = 1, 1, 0)
    
    '������ҩ
    BlnSelect = False
    If str�������� = "" Then
        BlnSelect = False
    ElseIf str�������� = mstrAllType Then
        BlnSelect = True
        For n = 0 To chkDeptType.Count - 1
            chkDeptType(n).Value = 1
        Next
    Else
        str�������� = str�������� & ","
        strArr = Split(str��������, ",")
        
        For n = 0 To chkDeptType.Count - 1
            chkDeptType(n).Value = 1
        Next
        
        For i = 0 To UBound(strArr)
            For n = 0 To chkDeptType.Count - 1
                If strArr(i) = chkDeptType(n).Caption Then
                    chkDeptType(n).Value = 0
                    BlnSelect = True
                    Exit For
                End If
            Next
        Next
    End If
    If BlnSelect = True Then
        chk����.Value = 1
        chk����.Tag = 1
    Else
        chk����.Value = 0
        chk����.Tag = 0
    End If
    For n = 0 To chkDeptType.Count - 1
        chkDeptType(n).Enabled = BlnSelect
    Next
    
    int�շѴ��� = Val(zlDatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, mlngModule, 0, Array(lbl�շѴ���, cbo�շѴ���), mblnHavePriv))
    If int�շѴ��� >= 0 And int�շѴ��� <= 2 Then
        cbo�շѴ���.ListIndex = int�շѴ���
    Else
        cbo�շѴ���.ListIndex = 0
    End If
    
    '59655
    chkSign.Value = IIf(Val(zlDatabase.GetPara("������ǩ��", glngSys, mlngModule, , Array(chkSign), mblnHavePriv)) = 1, 1, 0)
    
    chk���ܷ���.Value = IIf(Val(zlDatabase.GetPara("����ʱ�����������ʼ�¼", glngSys, mlngModule, , Array(chk���ܷ���), mblnHavePriv)) = 1, 1, 0)
    
    chk����ʱ��.Value = IIf(Val(zlDatabase.GetPara("����ҽ��������ʱ�����", glngSys, mlngModule, 0, Array(chk����ʱ��), mblnHavePriv)) = 1, 1, 0)
End Sub
 
Public Function ShowSetPara(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:���ò������
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '�޸�:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    '���ز�������
     Me.Show 1, frmMain
    ShowSetPara = mblnOk
End Function
