VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDistPara 
   Caption         =   "�����������"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   ControlBox      =   0   'False
   Icon            =   "frmDistPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra���� 
      Caption         =   "�����ӡ"
      Height          =   735
      Left            =   135
      TabIndex        =   30
      Top             =   7260
      Width           =   6525
      Begin VB.CommandButton cmdBarcodeSet 
         Caption         =   "�����ӡ����"
         Height          =   375
         Left            =   4680
         TabIndex        =   34
         Top             =   240
         Width           =   1620
      End
      Begin VB.OptionButton optBarcode 
         Caption         =   "��ʾѡ���ӡ"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   33
         Top             =   360
         Width           =   1770
      End
      Begin VB.OptionButton optBarcode 
         Caption         =   "�Զ���ӡ"
         Height          =   195
         Index           =   1
         Left            =   1530
         TabIndex        =   32
         Top             =   360
         Width           =   1170
      End
      Begin VB.OptionButton optBarcode 
         Caption         =   "����ӡ"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   990
      End
   End
   Begin VB.CheckBox chkBusy 
      Caption         =   "ҽ������æʱ�������"
      Height          =   300
      Left            =   165
      TabIndex        =   29
      Top             =   9460
      Width           =   4620
   End
   Begin VB.TextBox txt��ǰ���� 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   585
      TabIndex        =   27
      Text            =   "0"
      Top             =   9115
      Width           =   375
   End
   Begin VB.Frame fra���� 
      Caption         =   "���ﲡ������ʽ"
      Height          =   1020
      Left            =   135
      TabIndex        =   23
      Top             =   8070
      Width           =   6540
      Begin VB.OptionButton optSort 
         Caption         =   "���ұ���,����,����ʱ��,�Ǽ�ʱ��"
         Height          =   210
         Index           =   2
         Left            =   390
         TabIndex        =   26
         Top             =   675
         Width           =   3555
      End
      Begin VB.OptionButton optSort 
         Caption         =   "���ұ���,����,�Һ�ʱ��"
         Height          =   210
         Index           =   1
         Left            =   2775
         TabIndex        =   25
         Top             =   360
         Width           =   2280
      End
      Begin VB.OptionButton optSort 
         Caption         =   "���ұ���,����,���ݺ�"
         Height          =   210
         Index           =   0
         Left            =   390
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   2280
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7995
      Left            =   7050
      TabIndex        =   21
      Top             =   -120
      Width           =   45
   End
   Begin VB.Frame fra�Ŷӵ� 
      Caption         =   "�Ŷӵ���ӡ"
      Height          =   735
      Left            =   135
      TabIndex        =   17
      Top             =   6420
      Width           =   6525
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "�Ŷӵ���ӡ����"
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   270
         Width           =   1620
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "��ʾѡ���ӡ"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   20
         Top             =   375
         Width           =   1455
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "�Զ���ӡ"
         Height          =   195
         Index           =   1
         Left            =   1530
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "����ӡ"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   990
      End
   End
   Begin VB.Frame fra�Ŷӽк� 
      Caption         =   "�Ŷӽк�����"
      Height          =   1875
      Left            =   150
      TabIndex        =   7
      Top             =   4440
      Width           =   6540
      Begin VB.CheckBox chkԤԼ�Ŷ� 
         Caption         =   "ԤԼ�ҺŽ������"
         Height          =   270
         Left            =   4200
         TabIndex        =   16
         Top             =   1410
         Width           =   1905
      End
      Begin VB.CheckBox chkǩ���Ŷ� 
         Caption         =   "����̨ǩ����ʼ�Ŷ�"
         Height          =   330
         Left            =   2130
         TabIndex        =   15
         Top             =   1380
         Width           =   1935
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "�������������"
         Height          =   300
         Left            =   180
         TabIndex        =   14
         Top             =   1395
         Width           =   2340
      End
      Begin VB.OptionButton opt�Ŷ�ģʽ 
         Caption         =   "�ȷ������,��ҽ�����о���"
         Height          =   240
         Index           =   2
         Left            =   3870
         TabIndex        =   10
         Top             =   405
         Width           =   2625
      End
      Begin VB.OptionButton opt�Ŷ�ģʽ 
         Caption         =   "����̨������л�ҽ����������"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   720
         Width           =   3045
      End
      Begin VB.OptionButton opt�Ŷ�ģʽ 
         Caption         =   "��ֹȫԺ�Ŷӽк�"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   390
         Width           =   1770
      End
      Begin VB.Frame fra���ж��� 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         TabIndex        =   11
         Top             =   720
         Width           =   6330
         Begin VB.OptionButton opt���ж��� 
            Caption         =   "ҽ����������"
            Height          =   240
            Index           =   1
            Left            =   2070
            TabIndex        =   13
            Top             =   315
            Width           =   1725
         End
         Begin VB.OptionButton opt���ж��� 
            Caption         =   "����̨�������"
            Height          =   240
            Index           =   0
            Left            =   330
            TabIndex        =   12
            Top             =   315
            Width           =   1725
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7320
      TabIndex        =   2
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   1
      Top             =   270
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   5790
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistPara.frx":058A
            Key             =   "bm"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3360
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Ctrl+Aȫѡ,Ctrl+Cȫ��"
      Top             =   540
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5927
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imglst"
      SmallIcons      =   "imglst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   7937
      EndProperty
   End
   Begin MSComCtl2.UpDown upd���� 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4005
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtUD(1)"
      BuddyDispid     =   196632
      BuddyIndex      =   1
      OrigLeft        =   2625
      OrigTop         =   3990
      OrigRight       =   2865
      OrigBottom      =   4290
      Max             =   7
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtUD 
      Alignment       =   1  'Right Justify
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1005
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "0"
      Top             =   4005
      Width           =   675
   End
   Begin VB.Label lbl��ǰ���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ     Сʱ����,����Ϊ��,��ʾ�ڵ�ǰϵͳʱ���ڹҺŲ��˽��з���"
      Height          =   180
      Left            =   165
      TabIndex        =   28
      Top             =   9145
      Width           =   5670
   End
   Begin VB.Label lbl��Ч���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Զ�ˢ��            ���ڵĹҺŲ���,����Ϊ��,��ʾֻˢ�µ�ǰ�ĹҺŲ���"
      Height          =   180
      Left            =   225
      TabIndex        =   6
      Top             =   4065
      Width           =   6120
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   285
      Picture         =   "frmDistPara.frx":0B24
      Top             =   30
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    һ������̨����ͬʱ�����������ٴ����ҹҺŲ��ˣ����з�����ش�����ѡ���ɱ�����̨���з�����ٴ�����(Ctrl+Aȫѡ,Ctrl+Cȫ��)"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   930
      TabIndex        =   3
      Top             =   90
      Width           =   5805
   End
End
Attribute VB_Name = "frmDistPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public mlngModul As Long
Private mblnNotClick As Boolean
 

Private Sub cmdBarcodeSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & Int(glngSys \ 100) & "_BILL_1113_1", Me)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ObjItem As ListItem, strTmp As String
    
    For Each ObjItem In Me.lvwMain.ListItems
        If ObjItem.Checked Then
            strTmp = strTmp & "," & Mid(ObjItem.Key, 2)
        End If
    Next
    If strTmp = "" Then
        If MsgBox("��û�����ö��κο��ҷ���÷���̨�����ܽ��з��������" & vbCrLf & "�����ʱ��������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
        strTmp = "0"
    Else
        strTmp = Mid(strTmp, 2)
        If UBound(Split(strTmp, ",")) + 1 = lvwMain.ListItems.Count Then strTmp = ""
    End If
    zlDatabase.SetPara "�������", strTmp, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0 '�ձ�ʾȫ������
    zlDatabase.SetPara "������Ч����", Val(txtUD(1).Text), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0  '�ձ�ʾȫ������
    
    '1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.3-���Ŷӽк�
    zlDatabase.SetPara "�Ŷӽк�ģʽ", IIf(opt�Ŷ�ģʽ(0).Value, 0, IIf(opt�Ŷ�ģʽ(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0 '�ձ�ʾȫ������
    If opt�Ŷ�ģʽ(1).Value Then
        zlDatabase.SetPara "�ŶӺ���վ��", IIf(opt���ж���(0).Value, 0, 1), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0 '�ձ�ʾȫ������
    End If
    zlDatabase.SetPara "�������������", IIf(chk�������.Enabled = False, 0, chk�������.Value), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0 '�ձ�ʾȫ������
    zlDatabase.SetPara "����̨ǩ���Ŷ�", chkǩ���Ŷ�.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    '����:44621
    zlDatabase.SetPara "ԤԼ���ɶ���", chkԤԼ�Ŷ�.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0

    '38165
    zlDatabase.SetPara "�Ŷӵ���ӡ", IIf(optPrint(0).Value, 0, IIf(optPrint(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��������ʽ", IIf(optSort(0).Value, 0, IIf(optSort(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    
    '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
    zlDatabase.SetPara "�����ӡ��ʽ", IIf(optBarcode(0).Value, 0, IIf(optBarcode(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    
    '51223
    zlDatabase.SetPara "��ǰNСʱ����", Val(txt��ǰ����.Text), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    
    zlDatabase.SetPara "����æʱ�������", chkBusy.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0

    Unload Me
End Sub

 
Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & Int(glngSys \ 100) & "_BILL_1113", Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Dim i As Integer
        If UCase(Chr(KeyCode)) = "A" Then
            For i = 1 To lvwMain.ListItems.Count
                lvwMain.ListItems(i).Checked = True
            Next
        ElseIf UCase(Chr(KeyCode)) = "C" Then
            For i = 1 To lvwMain.ListItems.Count
                lvwMain.ListItems(i).Checked = False
            Next
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim ObjItem As ListItem
    Dim blnEnabled As Boolean
    
    Call RestoreWinState(Me, App.ProductName)
    mblnNotClick = True
    Select Case Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, mlngModul, , Array(opt�Ŷ�ģʽ(0), opt�Ŷ�ģʽ(1), opt�Ŷ�ģʽ(2)), InStr(1, mstrPrivs, ";��������;") > 0))
    Case 1 '1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
        opt�Ŷ�ģʽ(1).Value = True
        blnEnabled = True
    Case 2
        opt�Ŷ�ģʽ(2).Value = True
        blnEnabled = False
    Case Else
        opt�Ŷ�ģʽ(0).Value = True
        blnEnabled = False
    End Select
    
    Select Case Val(zlDatabase.GetPara("�ŶӺ���վ��", glngSys, mlngModul, , Array(opt���ж���(0), opt���ж���(1)), InStr(1, mstrPrivs, ";��������;") > 0))
    Case 0  '0-�������̨�������;1-����ҽ����������
        opt���ж���(0).Value = True
    Case Else
        opt���ж���(1).Value = True
    End Select
    opt�Ŷ�ģʽ(1).Tag = IIf(opt���ж���(1).Enabled, 1, 0)
    opt���ж���(1).Enabled = opt���ж���(1).Enabled And blnEnabled
    opt���ж���(0).Enabled = opt���ж���(0).Enabled And blnEnabled
    
    chk�������.Value = IIf(Val(zlDatabase.GetPara("�������������", glngSys, mlngModul, , Array(chk�������), InStr(1, mstrPrivs, ";��������;") > 0)) = 1, 1, 0)
    chkǩ���Ŷ�.Value = IIf(Val(zlDatabase.GetPara("����̨ǩ���Ŷ�", glngSys, mlngModul, , Array(chkǩ���Ŷ�), InStr(1, mstrPrivs, ";��������;") > 0)) = 1, 1, 0)
    '����:44621
    chkԤԼ�Ŷ�.Value = IIf(Val(zlDatabase.GetPara("ԤԼ���ɶ���", glngSys, mlngModul, , Array(chkԤԼ�Ŷ�), InStr(1, mstrPrivs, ";��������;") > 0)) = 1, 1, 0)
    
    chk�������.Tag = IIf(chk�������.Enabled, 1, 0)
    'chk�������.Enabled = Not opt���ж���(1).Value And chk�������.Enabled
    '����:43012
    Select Case Val(zlDatabase.GetPara("��������ʽ", glngSys, mlngModul, , Array(fra����, optSort(0), optSort(1), optSort(2)), InStr(1, mstrPrivs, ";��������;") > 0))
    Case 0
        optSort(0).Value = True
        optSort(1).Value = False
        optSort(2).Value = False
    Case 1
        optSort(1).Value = True
        optSort(0).Value = False
        optSort(2).Value = False
    Case 2
        optSort(2).Value = True
        optSort(0).Value = False
        optSort(1).Value = False
    End Select
    
    '38165
    Select Case Val(zlDatabase.GetPara("�Ŷӵ���ӡ", glngSys, mlngModul, , Array(optPrint(0), optPrint(1), optPrint(2), fra�Ŷӵ�), InStr(1, mstrPrivs, ";��������;") > 0))
    Case 0
        optPrint(0).Value = True
    Case 1
        optPrint(1).Value = True
    Case Else
        optPrint(2).Value = True
    End Select
    '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
    Select Case Val(zlDatabase.GetPara("�����ӡ��ʽ", glngSys, mlngModul, , Array(optBarcode(0), optBarcode(1), optBarcode(2), fra����), InStr(1, mstrPrivs, ";��������;") > 0))
    Case 0
        optBarcode(0).Value = True
    Case 1
        optBarcode(1).Value = True
    Case Else
        optBarcode(2).Value = True
    End Select
    strTmp = zlDatabase.GetPara("������Ч����", glngSys, mlngModul, , Array(txtUD(1), lbl��Ч����), InStr(1, mstrPrivs, ";��������;") > 0)
    upd����.Value = Val(strTmp): txtUD(1).Text = Val(strTmp)
    upd����.Enabled = txtUD(1).Enabled
    mblnNotClick = False
    
    '�ȵõ���ǰ���õķ������ID,�ձ�ʾ��������
    strTmp = zlDatabase.GetPara("�������", glngSys, mlngModul, , Array(lvwMain), InStr(1, mstrPrivs, ";��������;") > 0)
    Me.lvwMain.ListItems.Clear
    On Error GoTo errH
    
    If InStr(mstrPrivs, "���п���") > 0 Then
        Set rsTmp = GetDepartments("'�ٴ�'", "1,3")
    Else
        strSQL = _
            " Select A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And B.��������='�ٴ�' And B.������� IN(1,3)" & _
            " And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    End If
    
    With rsTmp
        Do While Not .EOF
            Set ObjItem = Me.lvwMain.ListItems.Add(, "K" & !ID, !����, "bm", "bm")
            ObjItem.SubItems(1) = Nvl(!����)
            If InStr("," & strTmp & ",", "," & !ID & ",") > 0 Or strTmp = "" Then ObjItem.Checked = True
            .MoveNext
        Loop
    End With
    
    '�����:51223
     strTmp = zlDatabase.GetPara("��ǰNСʱ����", glngSys, mlngModul, , Array(txt��ǰ����, lbl��ǰ����), InStr(1, mstrPrivs, ";��������;") > 0)
     If strTmp = "" Then
        txt��ǰ����.Text = "0"
     Else
        txt��ǰ����.Text = strTmp
     End If
     
     chkBusy.Value = Val(zlDatabase.GetPara("����æʱ�������", glngSys, mlngModul, , Array(chkBusy), InStr(1, mstrPrivs, ";��������;") > 0))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Long
    
    'Me.cmdCancel.Top = ScaleTop + ScaleHeight - cmdCancel.Height - 100   '   lvwMain.Top + lvwMain.Height + 120
    'Me.cmdOK.Top = Me.cmdCancel.Top
    Me.cmdCancel.Left = ScaleWidth - (Me.cmdCancel.Width + 90)
    Me.cmdOK.Left = Me.cmdCancel.Left 'Me.cmdCancel.Left - (Me.cmdOK.Width + 20)
    Me.Frame1.Left = Me.cmdOK.Left - Frame1.Width - 50
    Me.Frame1.Height = ScaleHeight + 100
    '�����:51223
    txt��ǰ����.Top = Me.ScaleHeight - txt��ǰ����.Height - chkBusy.Height - 100
    lbl��ǰ����.Top = txt��ǰ����.Top + (txt��ǰ����.Height - lbl��ǰ����.Height) / 2
    chkBusy.Top = txt��ǰ����.Top + txt��ǰ����.Height + 50
    fra����.Top = txt��ǰ����.Top - fra����.Height - 50
    '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
    Me.fra����.Top = fra����.Top - fra����.Height - 50
    Me.fra�Ŷӵ�.Top = fra����.Top - fra�Ŷӵ�.Height - 50
    
    txtUD(1).Top = fra�Ŷӵ�.Top - txtUD(1).Height - 50: upd����.Top = txtUD(1).Top
    lbl��Ч����.Top = txtUD(1).Top + (txtUD(1).Height - lbl��Ч����.Height) \ 2
    fra�Ŷӽк�.Top = txtUD(1).Top - fra�Ŷӽк�.Height - 50
    fra�Ŷӽк�.Width = Frame1.Left - fra�Ŷӽк�.Left * 2
    fra�Ŷӵ�.Width = Frame1.Left - fra�Ŷӵ�.Left * 2
    fra����.Width = Frame1.Left - fra����.Left * 2
    fra����.Width = Frame1.Left - fra����.Left * 2
    i = Frame1.Left - (lvwMain.Left * 2)
    lvwMain.Width = IIf(i > Screen.TwipsPerPixelX, i, Screen.TwipsPerPixelX)
    lvwMain.Height = fra�Ŷӽк�.Top - 50 - lvwMain.Top 'IIf(i > Screen.TwipsPerPixelY, i, Screen.TwipsPerPixelY)
    
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain.Sorted = True
    If lvwMain.SortKey = ColumnHeader.Index - 1 Then
        If lvwMain.SortOrder = lvwAscending Then
            lvwMain.SortOrder = lvwDescending
        Else
            lvwMain.SortOrder = lvwAscending
        End If
    Else
        lvwMain.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub opt���ж���_Click(Index As Integer)
        If mblnNotClick Then Exit Sub
        
'        If opt���ж���(1).Value = False Then
'            chk�������.Enabled = IIf(Val(chk�������.Tag) = 1, True, False)
'        Else
'            chk�������.Enabled = False
'        End If
'        If chk�������.Enabled = False Then chk�������.Value = 0
End Sub

Private Sub opt�Ŷ�ģʽ_Click(Index As Integer)
        If mblnNotClick Then Exit Sub
        If opt�Ŷ�ģʽ(1).Value Then
                opt���ж���(0).Enabled = IIf(Val(opt�Ŷ�ģʽ(1).Tag) = 1, True, False)
                opt���ж���(1).Enabled = opt���ж���(0).Enabled
        Else
                opt���ж���(0).Enabled = False
                opt���ж���(1).Enabled = opt���ж���(0).Enabled
        End If
        chkǩ���Ŷ�.Enabled = opt�Ŷ�ģʽ(0).Value = False
End Sub
Private Sub txt��ǰ����_KeyPress(KeyAscii As Integer)
    '�����:51223
     zlControl.TxtCheckKeyPress txt��ǰ����, KeyAscii, m����ʽ
End Sub
