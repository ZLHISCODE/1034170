VERSION 5.00
Begin VB.Form frmTechnicSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frmTechnicSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkNewNurse 
      Caption         =   "Ѫ͸����д�°滤���¼"
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   2880
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   2535
      Begin VB.OptionButton opt���˹��� 
         Caption         =   "����ʱ��"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   34
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton opt���˹��� 
         Caption         =   " ִ��ʱ��"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkִ�б������� 
      Caption         =   "ִ�б���ʱ�շѻ�������"
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   2040
      Width           =   2520
   End
   Begin VB.Frame fraNotify 
      Caption         =   "��������"
      Height          =   1230
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   6270
      Begin VB.CheckBox chkWarn 
         Caption         =   "Ѫ������"
         Height          =   195
         Index           =   2
         Left            =   3290
         TabIndex        =   38
         Top             =   885
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "������"
         Height          =   195
         Index           =   1
         Left            =   2450
         TabIndex        =   19
         Top             =   885
         Width           =   900
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "����������ʾ"
         Height          =   195
         Left            =   4320
         TabIndex        =   36
         Top             =   330
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   4320
         TabIndex        =   35
         Top             =   630
         Width           =   1410
      End
      Begin VB.Frame fraLinM 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   28
         Top             =   525
         Width           =   300
      End
      Begin VB.TextBox txtMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "10"
         Top             =   330
         Width           =   300
      End
      Begin VB.CheckBox chkNotify 
         Caption         =   "ÿ    �����Զ�ˢ�����������е�����"
         Height          =   195
         Left            =   345
         TabIndex        =   25
         Top             =   345
         Width           =   3390
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   23
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame fraLinD 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   22
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "1"
         Top             =   600
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��������"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   20
         Top             =   885
         Width           =   1065
      End
      Begin VB.Label lblNotifyArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   600
         TabIndex        =   27
         Top             =   880
         Width           =   810
      End
      Begin VB.Label lblNotifyDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ������ɵ�������ʾ����������"
         Height          =   180
         Left            =   615
         TabIndex        =   26
         Top             =   615
         Width           =   3060
      End
   End
   Begin VB.ListBox lst������� 
      Columns         =   3
      ForeColor       =   &H80000012&
      Height          =   1110
      IMEMode         =   3  'DISABLE
      Left            =   2760
      Style           =   1  'Checkbox
      TabIndex        =   8
      ToolTipText     =   "��Ctrl+Aȫѡ����Ctrl+Cȫ��"
      Top             =   3480
      Width           =   3615
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6525
      TabIndex        =   14
      Top             =   5955
      Width           =   6525
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5310
         TabIndex        =   10
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4215
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1500
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   7080
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   7080
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.ListBox lst������� 
      Columns         =   3
      ForeColor       =   &H80000012&
      Height          =   2370
      IMEMode         =   3  'DISABLE
      Left            =   2760
      Style           =   1  'Checkbox
      TabIndex        =   7
      ToolTipText     =   "��Ctrl+Aȫѡ����Ctrl+Cȫ��"
      Top             =   450
      Width           =   3615
   End
   Begin VB.CheckBox chkExeLog 
      Caption         =   "�ϸ�Ҫ���¼ִ�е����"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1281
      Width           =   2280
   End
   Begin VB.CheckBox chkRoom 
      Caption         =   "ֻ��ʾָ����ִ�м䷶Χ"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2280
   End
   Begin VB.CheckBox chkActLog 
      Caption         =   "���������˴���ִ�м�¼"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   2280
   End
   Begin VB.CheckBox chkFinish 
      Caption         =   "����δ�շѲ������ִ��"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   727
      Width           =   2280
   End
   Begin VB.CheckBox chkƤ�� 
      Caption         =   "��дƤ�Խ��ʱ��֤���"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1004
      Width           =   2280
   End
   Begin VB.Frame fraLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   600
      TabIndex        =   12
      Top             =   345
      Width           =   465
   End
   Begin VB.Frame fraRoom 
      Caption         =   " ִ�м� "
      Height          =   1440
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   2445
      Begin VB.ListBox lstRoom 
         Enabled         =   0   'False
         Height          =   1110
         ItemData        =   "frmTechnicSetup.frx":000C
         Left            =   120
         List            =   "frmTechnicSetup.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
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
      Left            =   600
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "���ˢ�¼��Ϊ 30 �룬����Ϊ 0 ��ʾ���Զ�ˢ��"
      Top             =   165
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "���˹��˷�ʽ����"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "����һ��ͨ��"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lbl������� 
      Caption         =   "�������"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lbl������� 
      Caption         =   "���ݹ������"
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label lblRefresh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ      ���Զ�ˢ���嵥"
      Height          =   180
      Left            =   390
      TabIndex        =   13
      ToolTipText     =   "���ˢ�¼��Ϊ 30 �룬����Ϊ 0 ��ʾ���Զ�ˢ��"
      Top             =   180
      Width           =   1980
   End
End
Attribute VB_Name = "frmTechnicSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mstrPrivs As String
Public mlng����ID As Long 'IN:��ǰִ�п���ID
Public mblnOK As Boolean

Private Sub chkRoom_Click()
    lstRoom.Enabled = chkRoom.Value = 1 And lstRoom.Tag = ""
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long, k As Long, bln���� As Boolean
    Dim blnSetup As Boolean
    
    'ִ�м䷶Χ
    strPar = ""
    If chkRoom.Value = 1 Then
        For i = 0 To lstRoom.ListCount - 1
            If lstRoom.Selected(i) Then
                strPar = strPar & "|" & lstRoom.List(i)
            End If
        Next
        If strPar = "" Then
            MsgBox "������ѡ��һ��ִ�м䡣", vbInformation, gstrSysName
            lstRoom.SetFocus: Exit Sub
        End If
    End If
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("ִ�м䷶Χ", Replace(Mid(strPar, 2), "'", "''"), glngSys, pҽ������վ, blnSetup)
        
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
    Call zlDatabase.SetPara("ҽ��ˢ�¼��", Val(txtRefresh.Text), glngSys, pҽ������վ, blnSetup)
    
    '�Ƿ��������ִ�м�¼
    Call zlDatabase.SetPara("����ִ�м�¼", chkActLog.Value, glngSys, pҽ������վ, blnSetup)

    '�Ƿ��������δ�շѲ��˵���Ŀ
    Call zlDatabase.SetPara("δ�շ����", chkFinish.Value, glngSys, pҽ������վ, blnSetup)
    
    '�ϸ�Ҫ���¼ִ�е����
    Call zlDatabase.SetPara("��¼ִ�����", chkExeLog.Value, glngSys, pҽ������վ, blnSetup)
    
    '��дƤ�Խ��ʱ��֤���
    Call zlDatabase.SetPara("Ƥ����֤���", chkƤ��.Value, glngSys, pҽ������վ, blnSetup)
    
    'ִ�б���ʱ�շѻ�������
    Call zlDatabase.SetPara("ִ�б���ʱ�շѻ�������", chkִ�б�������.Value, glngSys, pҽ������վ, blnSetup)
    
    '���˹��˷�ʽ
    Call zlDatabase.SetPara("���˹��˷�ʽ", IIf(opt���˹���(0) = True, 0, 1), glngSys, pҽ������վ, blnSetup)
    
    'Ѫ͸����д�°滤���¼
    Call zlDatabase.SetPara("Ѫ͸����д�°滤���¼", chkNewNurse.Value, glngSys, pҽ������վ, blnSetup)
    
    '�������
    k = 0
    strPar = ""
    For i = 0 To lst�������.ListCount - 1
        If lst�������.Selected(i) Then
            strPar = strPar & Chr(lst�������.ItemData(i))
            If Chr(lst�������.ItemData(i)) = "E" Then bln���� = True
            k = k + 1
        End If
    Next
    If strPar = "" Then
        MsgBox "������ѡ��һ��Ҫִ�е��������", vbInformation, gstrSysName
        lst�������.SetFocus: Exit Sub
    End If
    If k = lst�������.ListCount Then strPar = ""
    Call zlDatabase.SetPara("�������", strPar, glngSys, pҽ������վ, blnSetup)
    
    '�������
    If bln���� Then
        k = 0
        strPar = ""
        For i = 0 To lst�������.ListCount - 1
            If lst�������.Selected(i) Then
                strPar = strPar & "," & lst�������.ItemData(i)
                k = k + 1
            End If
        Next
        If strPar = "" Then
            MsgBox "������ѡ��һ��Ҫִ�е��������", vbInformation, gstrSysName
            lst�������.SetFocus: Exit Sub
        Else
            strPar = Mid(strPar, 2)
        End If
        If k = lst�������.ListCount Then strPar = ""
        Call zlDatabase.SetPara("�������", strPar, glngSys, pҽ������վ, blnSetup)
    End If
    
    Call zlDatabase.SetPara("�Զ�ˢ��ҽ�����", IIf(chkNotify.Value = 1, Val(txtMin.Text), ""), glngSys, pҽ������վ, blnSetup)
    Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", Val(txtDay.Text), glngSys, pҽ������վ, blnSetup)
    Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", "" & chkWarn(0).Value & chkWarn(1).Value & chkWarn(2).Value, glngSys, pҽ������վ, blnSetup)
    Call zlDatabase.SetPara("����������ʾ", chkSound.Value, glngSys, pҽ������վ, blnSetup)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 3)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        If KeyCode = vbKeyA Then
            SelAll������� (True)
        ElseIf KeyCode = vbKeyC Then
            SelAll������� (False)
        End If
    End If
End Sub

Private Sub SelAll�������(ByVal blnSel As Boolean)
    Dim i As Long
    
    For i = 0 To lst�������.ListCount - 1
        lst�������.Selected(i) = blnSel
    Next
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim blnSetup As Boolean, arrTmp As Variant, i As Long, bln���� As Boolean
    Dim intType As Integer
    
    mblnOK = False
    
    blnSetup = InStr(mstrPrivs, "��������") > 0
    txtRefresh.Text = Val(zlDatabase.GetPara("ҽ��ˢ�¼��", glngSys, pҽ������վ, , Array(lblRefresh, txtRefresh), blnSetup))
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
        
    '�Ƿ��������ִ�м�¼
    chkActLog.Value = Val(zlDatabase.GetPara("����ִ�м�¼", glngSys, pҽ������վ, , Array(chkActLog), blnSetup))
    
    '�Ƿ��������δ�շѲ��˵���Ŀ
    chkFinish.Value = Val(zlDatabase.GetPara("δ�շ����", glngSys, pҽ������վ, , Array(chkFinish), blnSetup))
    
    '�ϸ�Ҫ���¼ִ�е����
    chkExeLog.Value = Val(zlDatabase.GetPara("��¼ִ�����", glngSys, pҽ������վ, "0", Array(chkExeLog), blnSetup))
        
    '��дƤ�Խ��ʱ��֤���
    chkƤ��.Value = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, pҽ������վ, , Array(chkƤ��), blnSetup))
    
    'ִ�б���ʱ�շѻ�������
    chkִ�б�������.Value = Val(zlDatabase.GetPara("ִ�б���ʱ�շѻ�������", glngSys, pҽ������վ, "0", Array(chkִ�б�������), blnSetup))
    
    'Ѫ͸����д�°滤���¼
    chkNewNurse.Value = Val(zlDatabase.GetPara("Ѫ͸����д�°滤���¼", glngSys, pҽ������վ, , Array(chkNewNurse), blnSetup))
    
    '���˹��˷�ʽ
    If Val(zlDatabase.GetPara("���˹��˷�ʽ", glngSys, pҽ������վ, , Array(opt���˹���(0), opt���˹���(1)), blnSetup)) = 1 Then
        opt���˹���(1).Value = True
    Else
        opt���˹���(0).Value = True
    End If

    'ִ�з���
    strPar = zlDatabase.GetPara("ִ�м䷶Χ", glngSys, pҽ������վ, "", Array(chkRoom, fraRoom, lstRoom), blnSetup)
    If Not chkRoom.Enabled Then lstRoom.Tag = "1" '�̶����Ϊ������
    chkRoom.Value = IIf(strPar = "", 0, 1)
    strSQL = "Select ִ�м� From ҽ��ִ�з��� Where ����ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    Do While Not rsTmp.EOF
        lstRoom.AddItem rsTmp!ִ�м�
        If InStr("|" & strPar & "|", "|" & rsTmp!ִ�м� & "|") > 0 Then
            lstRoom.Selected(lstRoom.NewIndex) = True
        End If
        rsTmp.MoveNext
    Loop
    If lstRoom.ListCount > 0 Then
        lstRoom.TopIndex = 0
        lstRoom.ListIndex = 0
    ElseIf blnSetup Then
        chkRoom.Value = 0
        chkRoom.Enabled = False
    End If
    
    
    '�������
    strPar = zlDatabase.GetPara("�������", glngSys, pҽ������վ, , Array(lst�������), blnSetup)
        
    strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('5','6','7','8','9') Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    With lst�������
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            .ItemData(.NewIndex) = Asc(rsTmp!����)
            
            If strPar <> "" Then
                If InStr(strPar, rsTmp!����) > 0 Then
                    .Selected(.NewIndex) = True
                    If rsTmp!���� = "E" Then bln���� = True
                End If
            Else
                .Selected(.NewIndex) = True
                If Not bln���� Then bln���� = True
            End If
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    
    strPar = "0-��ͨ;1-��������;2-��ҩ����;3-��ҩ�巨;4-��ҩ�÷�;5-��������;6-�ɼ�����;7-��Ѫ����;8-��Ѫ;��"
    arrTmp = Split(strPar, ";")
    
    strPar = zlDatabase.GetPara("�������", glngSys, pҽ������վ, , Array(lst�������), blnSetup)
    If strPar <> "" Then
        strPar = "," & strPar & ","
    End If
    With lst�������
        For i = 0 To UBound(arrTmp)
            .AddItem arrTmp(i)
            .ItemData(.NewIndex) = Val(arrTmp(i))
            
            If strPar <> "" Then
                If InStr(strPar, "," & Val(arrTmp(i)) & ",") > 0 Then
                    .Selected(.NewIndex) = True
                End If
            Else
                .Selected(.NewIndex) = True
            End If
        Next
    End With
    lst�������.Enabled = bln����
    
    strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, pҽ������վ, , Array(chkNotify), InStr(mstrPrivs, "��������") > 0, intType)
    If Val(strPar) > 0 Then chkNotify.Value = 1: txtMin.Text = Val(strPar)
    'ǰ���¼��л��Զ����ã���˺���ǿ������
    If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0 Then
        txtMin.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pҽ������վ, 1, Array(lblNotifyDay, txtDay), InStr(mstrPrivs, "��������") > 0)
    txtDay.Text = Val(strPar)
    
    strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pҽ������վ, "000", Array(lblNotifyArea, chkWarn(0), chkWarn(1), chkWarn(2)), InStr(mstrPrivs, "��������") > 0)
    chkWarn(2).Visible = gblnѪ��ϵͳ
    For i = 1 To chkWarn.Count
        chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
    Next
    chkSound.Value = Val(zlDatabase.GetPara("����������ʾ", glngSys, pҽ������վ, "1", Array(chkSound, cmdSoundSet), blnSetup))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkNotify_Click()
    txtMin.Enabled = chkNotify.Value = 1
    If Visible And txtMin.Enabled Then txtMin.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    mstrPrivs = ""
End Sub

Private Sub lst�������_ItemCheck(Item As Integer)
    If Chr(lst�������.ItemData(Item)) = "E" Then
        lst�������.Enabled = lst�������.Selected(Item)
    End If
End Sub

Private Sub opt���˹���_Click(Index As Integer)
    opt���˹���(Index).Value = True
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
