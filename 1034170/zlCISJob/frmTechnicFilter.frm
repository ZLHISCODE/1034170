VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTechnicFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   Icon            =   "frmTechnicFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1260
      TabIndex        =   5
      Top             =   1560
      Width           =   1470
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   31
      Top             =   3570
      Width           =   1100
   End
   Begin VB.ComboBox cboDoctor 
      Height          =   300
      Left            =   4035
      TabIndex        =   19
      Text            =   "cboDoctor"
      Top             =   3075
      Width           =   1710
   End
   Begin VB.TextBox txt��ʶ�� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1260
      MaxLength       =   18
      TabIndex        =   1
      Top             =   855
      Width           =   1470
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1260
      TabIndex        =   3
      Top             =   1207
      Width           =   1470
   End
   Begin VB.CheckBox chk����סԺ 
      Caption         =   "ֻ��ʾ����סԺ����Ŀ"
      Height          =   195
      Left            =   3660
      TabIndex        =   13
      Top             =   2460
      Value           =   1  'Checked
      Width           =   2100
   End
   Begin VB.CheckBox chk��Դ 
      Caption         =   "���"
      Height          =   195
      Index           =   2
      Left            =   2820
      TabIndex        =   16
      Top             =   2760
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.CheckBox chk��Ч 
      Caption         =   "����"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   11
      Top             =   2460
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.CheckBox chk��Ч 
      Caption         =   "��ʱ"
      Height          =   195
      Index           =   1
      Left            =   1965
      TabIndex        =   12
      Top             =   2460
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4275
      TabIndex        =   4
      Top             =   1215
      Width           =   1470
   End
   Begin VB.TextBox txt���￨ 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4275
      TabIndex        =   2
      Top             =   855
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   28
      Top             =   1920
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Top             =   720
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   -105
      TabIndex        =   25
      Top             =   3480
      Width           =   6360
   End
   Begin VB.CommandButton cmdDefault 
      Cancel          =   -1  'True
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   330
      TabIndex        =   24
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CheckBox chk��Դ 
      Caption         =   "סԺ"
      Height          =   195
      Index           =   1
      Left            =   1965
      TabIndex        =   15
      Top             =   2760
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.CheckBox chk��Դ 
      Caption         =   "����"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   14
      Top             =   2760
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3075
      Width           =   2115
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3915
      TabIndex        =   7
      Top             =   2010
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   57802755
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1140
      TabIndex        =   6
      Top             =   2010
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   95289347
      CurrentDate     =   38082
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3420
      TabIndex        =   20
      Top             =   3570
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����(&5)"
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   1620
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   3375
      TabIndex        =   18
      Top             =   3135
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3330
      TabIndex        =   30
      Top             =   2070
      Width           =   180
   End
   Begin VB.Label lbl��Ч 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����Ч"
      Height          =   180
      Left            =   270
      TabIndex        =   29
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&3)"
      Height          =   180
      Left            =   480
      TabIndex        =   9
      Top             =   1275
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݺ�(&4)"
      Height          =   180
      Left            =   3315
      TabIndex        =   10
      Top             =   1275
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���￨(&2)"
      Height          =   180
      Left            =   3315
      TabIndex        =   8
      Top             =   915
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʶ��(&1)"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   915
      Width           =   810
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   270
      Picture         =   "frmTechnicFilter.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ���ù��������Ա�׼ȷ����ִ�м�¼������ʱ�䷶Χ������ȷ������߲����ٶȡ�"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   915
      TabIndex        =   27
      Top             =   180
      Width           =   3780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Դ"
      Height          =   180
      Left            =   270
      TabIndex        =   23
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���˿���"
      Height          =   180
      Left            =   270
      TabIndex        =   22
      Top             =   3135
      Width           =   720
   End
   Begin VB.Label lbl��ѯʱ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ��ʱ��"
      Height          =   180
      Left            =   240
      TabIndex        =   21
      Top             =   2070
      Width           =   720
   End
End
Attribute VB_Name = "frmTechnicFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public mblnOK As Boolean
Public mstrDeptNode As String   '��ǰҽ������������վ��

Private mblnLoad As Boolean
Private mstrDeptNodePre As String

Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboDoctor.Tag = "keypress"
        If SeekDoctor = False Then
            cboDoctor.Tag = ""
            cboDoctor.SetFocus
        End If
    End If
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
    If cboDoctor.Tag = "keypress" Then
        cboDoctor.Tag = ""
    ElseIf cboDoctor.ListIndex = -1 And cboDoctor.Text <> "" Then
        If SeekDoctor = False Then
            cboDoctor.Text = ""
        End If
    End If
End Sub

Private Function SeekDoctor() As Boolean
'���ܣ����ݵ�ǰ�������ݲ���ҽ���б�

    Dim strTxt As String, blnYes As Boolean
    Dim i As Long, bytKind As Byte
    
    strTxt = UCase(Trim(cboDoctor.Text))
    If strTxt = "����ҽ��" Then
        cboDoctor.ListIndex = 0
        SeekDoctor = True
        Exit Function
    End If
    
    If zlCommFun.IsCharAlpha(strTxt) Then
        bytKind = 0
    ElseIf InStr(strTxt, "-") > 0 Then
        bytKind = 1
    Else
        bytKind = 2
    End If
    
    'i=0�ǡ�����ҽ����
    For i = 1 To cboDoctor.ListCount - 1
            If bytKind = 0 Then
            If cboDoctor.List(i) Like "*/" & strTxt & "-*" Or cboDoctor.List(i) Like strTxt & "/*" Then
                blnYes = True
            End If
        ElseIf bytKind = 2 Then
            If cboDoctor.List(i) Like "*-" & strTxt Then
                blnYes = True
            End If
        Else
            If cboDoctor.List(i) = strTxt Then
                blnYes = True
            End If
        End If
        If blnYes Then
            cboDoctor.ListIndex = i
            SeekDoctor = True
            Exit Function
        End If
    Next
    If cboDoctor.ListCount > 0 Then
        cboDoctor.ListIndex = 0
        SeekDoctor = True
    End If
End Function

Private Sub chk��Դ_Click(Index As Integer)
    If chk��Դ(0).Value = 0 And chk��Դ(1).Value = 0 And chk��Դ(2).Value = 0 Then
        chk��Դ((Index + 1) Mod 3).Value = 1
    End If
    
    chk����סԺ.Enabled = chk��Դ(1).Value = 1
    
    If Me.Visible Then
        Call LoadDeptList
        Call LoadDoctorList
    End If
End Sub

Private Sub chk��Ч_Click(Index As Integer)
    If chk��Ч(0).Value = 0 And chk��Ч(1).Value = 0 Then
        chk��Ч((Index + 1) Mod 2).Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub

Private Sub cmdDefault_Click()
    Call Form_Load
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String

    Call txtNO_Validate(False)

    '�������
    Call zlDatabase.SetPara("������Դ", chk��Դ(0).Value & chk��Դ(1).Value & chk��Դ(2).Value, glngSys, pҽ������վ, InStr(mstrPrivs, "��������") > 0)
    Call zlDatabase.SetPara("ҽ����Ч", chk��Ч(0).Value & chk��Ч(1).Value, glngSys, pҽ������վ, InStr(mstrPrivs, "��������") > 0)
    Call zlDatabase.SetPara("ֻ��ʾ����סԺ��Ŀ", chk����סԺ.Value, glngSys, pҽ������վ, InStr(mstrPrivs, "��������") > 0)
    With cboDoctor
        If .ListIndex = 0 Or .ListIndex = -1 Then
            strTmp = ""
        Else
            strTmp = Split(.Text, "-")(1)
        End If
        Call zlDatabase.SetPara("������", strTmp, glngSys, pҽ������վ, InStr(mstrPrivs, "��������") > 0)
    End With
        
    mblnOK = True
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim curDate As Date
    
    '���˹��˷�ʽ
    lbl��ѯʱ��.Caption = IIf(Val(zlDatabase.GetPara("���˹��˷�ʽ", glngSys, pҽ������վ)) = 1, "����ʱ��", "ִ��ʱ��")
    
    '�����һ����ȡ�ĵ�ǰʱ��,����������ʱˢ�½��ʱ��Ϊ��ǰʱ��
    If Not mblnLoad Then
        If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            curDate = zlDatabase.Currentdate
            dtpBegin.MaxDate = curDate + 7
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59")
            dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
        End If
    End If
    If mblnLoad Then mblnLoad = False
    
    If mstrDeptNodePre <> mstrDeptNode Then
        mstrDeptNodePre = mstrDeptNode
        
        Call LoadDeptList
        Call LoadDoctorList
    End If
    
    '�Զ���λ
    dtpBegin.SetFocus
    If txtNO.Text <> "" Then
        txtNO.Text = "": txtNO.SetFocus
    End If
    If txt����.Text <> "" Then
        txt����.Text = "": txt����.SetFocus
    End If
    If txt���￨.Text <> "" Then
        txt���￨.Text = "": txt���￨.SetFocus
    End If
    If txt��ʶ��.Text <> "" Then
        txt��ʶ��.Text = "": txt��ʶ��.SetFocus
    End If
    If txt����.Text <> "" Then
        txt����.Text = "": txt����.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strKey As String
    
    mblnLoad = True
    
    mstrDeptNodePre = ""
    txtNO.Text = ""
    txt��ʶ��.Text = ""
    txt����.Text = ""
    txt����.Text = ""
    txt���￨.Text = ""
    txt���￨.PasswordChar = IIf(gblnCardHide, "*", "")
    
    '����סԺ
    chk����סԺ.Value = Val(zlDatabase.GetPara("ֻ��ʾ����סԺ��Ŀ", glngSys, pҽ������վ, "1", Array(chk����סԺ), InStr(mstrPrivs, "��������") > 0))
    
    '��Դ
    strKey = zlDatabase.GetPara("������Դ", glngSys, pҽ������վ, "111", Array(chk��Դ(0), chk��Դ(1), chk��Դ(2)), InStr(mstrPrivs, "��������") > 0)
    chk��Դ(0).Value = Val(Mid(strKey, 1, 1))
    chk��Դ(1).Value = Val(Mid(strKey, 2, 1))
    chk��Դ(2).Value = Val(Mid(strKey, 3, 1))
    
    '��Ч
    strKey = zlDatabase.GetPara("ҽ����Ч", glngSys, pҽ������վ, "11", Array(chk��Ч(0), chk��Ч(1)), InStr(mstrPrivs, "��������") > 0)
    chk��Ч(0).Value = Val(Mid(strKey, 1, 1))
    chk��Ч(1).Value = Val(Mid(strKey, 2, 1))
    
    '����ʱ��
    curDate = zlDatabase.Currentdate
    dtpBegin.MaxDate = curDate + 7
    dtpBegin.Value = Format(curDate - 1, "yyyy-MM-dd 00:00")
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59")
    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
            
    Call LoadDeptList
    Call LoadDoctorList
    
    strKey = zlDatabase.GetPara("������", glngSys, pҽ������վ, "", , InStr(mstrPrivs, "��������") > 0)
    Call zlControl.CboLocate(cboDoctor, IIf(strKey = "ALL", "����ҽ��", strKey))
    mblnOK = False
End Sub

Private Sub LoadDeptList()
'���ܣ����ݲ�����Դ��ȡ���˿���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngPre As Long
    
    If cboDept.ListIndex <> -1 Then
        lngPre = cboDept.ItemData(cboDept.ListIndex)
    End If
    strSQL = "Select Distinct A.ID,A.����,A.����,B.�������" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.�������� IN('�ٴ�','����')" & _
        " And B.������� IN(3,[1],[2])" & _
        IIf(mstrDeptNode <> "", " And (A.վ�� = [3] Or A.վ�� is Null)", "") & _
        " And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(chk��Դ(0).Value = 1 Or chk��Դ(2).Value = 1, 1, -1), IIf(chk��Դ(1).Value = 1, 2, -1), mstrDeptNode)
    On Error GoTo 0
    cboDept.Clear
    cboDept.AddItem "���п���"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDept.ListIndex = cboDept.NewIndex
        rsTmp.MoveNext
    Next
            
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    If IsNumeric(txtNO.Text) Then
        txtNO.Text = GetFullNO(txtNO.Text, 14)
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txtNO_GotFocus()
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txt���￨_GotFocus()
    Call zlControl.TxtSelAll(txt���￨)
End Sub

Private Sub txt���￨_KeyPress(KeyAscii As Integer)
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt��ʶ��_GotFocus()
    Call zlControl.TxtSelAll(txt��ʶ��)
End Sub

Private Sub txt��ʶ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub LoadDoctorList()
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim lngPre As Long
    
    If cboDoctor.ListIndex <> -1 Then
        lngPre = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    
    cboDoctor.Clear
    cboDoctor.AddItem "����ҽ��"
    cboDoctor.ListIndex = 0
    
    Set rsTmp = GetDoctorRs
    For i = 1 To rsTmp.RecordCount
        cboDoctor.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDoctor.ItemData(cboDoctor.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDoctor.ListIndex = cboDoctor.NewIndex
        rsTmp.MoveNext
    Next
    
End Sub

Private Function GetDoctorRs() As ADODB.Recordset
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(3,[1],[2])"
    strSQL = "Select Distinct A.ID,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        IIf(mstrDeptNode <> "", " And (A.վ�� = [3] Or A.վ�� is Null)", "") & _
        " And B.����ID IN(" & strSQL & ")" & _
        " Order by A.����"
        
    Set GetDoctorRs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(chk��Դ(0).Value = 1 Or chk��Դ(2).Value = 1, 1, -1), IIf(chk��Դ(1).Value = 1, 2, -1), mstrDeptNode)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

