VERSION 5.00
Begin VB.Form frm���մ���༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���մ���༭"
   ClientHeight    =   5970
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   4530
   Icon            =   "frm���մ���༭.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmb������� 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1335
      Width           =   1425
   End
   Begin VB.CheckBox chkҽ�� 
      Caption         =   "ҽ����Ŀ(&I)"
      Height          =   225
      Left            =   1170
      TabIndex        =   8
      Top             =   1770
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   255
      TabIndex        =   26
      Top             =   5490
      Width           =   1100
   End
   Begin VB.Frame frmRule 
      Caption         =   "ͳ��֧���������"
      Height          =   2535
      Left            =   285
      TabIndex        =   13
      Top             =   2805
      Width           =   4080
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   23
         Top             =   2115
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   21
         Top             =   1755
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   19
         Top             =   1245
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   16
         Top             =   585
         Width           =   630
      End
      Begin VB.OptionButton opt�㷨 
         Caption         =   "סԺ�ն�����㷨(&Z)"
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   17
         Top             =   975
         Width           =   2265
      End
      Begin VB.OptionButton opt�㷨 
         Caption         =   "�ܶ�������㷨(&B)"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��׼��������        ��"
         Height          =   180
         Index           =   6
         Left            =   465
         TabIndex        =   22
         Top             =   2175
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ÿ����׼����        Ԫ"
         Height          =   180
         Index           =   5
         Left            =   465
         TabIndex        =   20
         Top             =   1815
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ÿ�ջ�������        Ԫ"
         Height          =   180
         Index           =   4
         Left            =   465
         TabIndex        =   18
         Top             =   1305
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ͳ��֧������        %"
         Height          =   180
         Index           =   3
         Left            =   465
         TabIndex        =   15
         Top             =   645
         Width           =   1890
      End
   End
   Begin VB.Frame fraKind 
      Caption         =   "����"
      Height          =   630
      Left            =   285
      TabIndex        =   9
      Top             =   2070
      Width           =   4095
      Begin VB.OptionButton opt���� 
         Caption         =   "����(&F)"
         Height          =   180
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Top             =   315
         Width           =   945
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҽ��(&D)"
         Height          =   180
         Index           =   2
         Left            =   1425
         TabIndex        =   11
         Top             =   315
         Width           =   945
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҩƷ(&M)"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2070
      TabIndex        =   24
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3265
      TabIndex        =   25
      Top             =   5490
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   5
      Top             =   937
      Width           =   1425
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   3
      Top             =   536
      Width           =   3195
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   1
      Top             =   135
      Width           =   1425
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "�������(&F)"
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   1398
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&U)"
      Height          =   180
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Index           =   2
      Left            =   495
      TabIndex        =   4
      Top             =   997
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   1
      Left            =   495
      TabIndex        =   2
      Top             =   596
      Width           =   630
   End
End
Attribute VB_Name = "frm���մ���༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    text���� = 0
    Text���� = 1
    Text���� = 2
    Text���� = 3
    Text���� = 4
    Text��׼ = 5
    Text���� = 6

    CheckҩƷ = 1
    Checkҽ�� = 2
    Check���� = 3
    
    Check���� = 1
    CheckסԺ�� = 2
End Enum

Dim mlng���� As Long
Dim mstrID As String         '��ǰ�༭��ҽ������ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub chkҽ��_Click()
    mblnChange = True
End Sub

Private Sub chkҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub cmb�������_Click()
    mblnChange = True
End Sub

Private Sub cmb�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If IsValid() = False Then Exit Sub
    If Save��Ŀ() = False Then Exit Sub
    
    If mstrID = "" Then
        '��������
        txtEdit(text����).Text = zlDatabase.GetMax("����֧������", "����", 6, " where ����=" & mlng����)
        For lngIndex = Text���� To Text����
            txtEdit(lngIndex).Text = ""
        Next
        chkҽ��.Value = 1
        mblnChange = False
        txtEdit(text����).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function Save��Ŀ() As Boolean
    Dim lngID As Long, lng���� As Long, lng�㷨 As Long
    Dim dblͳ��ȶ� As Double, dbl��׼���� As Double, dbl��׼���� As Double
    Dim lngIndex As Long, lst As ListItem
    
    On Error GoTo errHandle
    
    For lngIndex = 1 To 3
        If opt����(lngIndex).Value = True Then
            lng���� = lngIndex
            Exit For
        End If
    Next
    If opt�㷨(1).Value = True Then
        '������
        lng�㷨 = 1
        dblͳ��ȶ� = Val(txtEdit(Text����).Text)
        
    Else
        '��סԺ��
        lng�㷨 = 2
        dblͳ��ȶ� = Val(txtEdit(Text����).Text)
        dbl��׼���� = Val(txtEdit(Text��׼).Text)
        dbl��׼���� = Val(txtEdit(Text����).Text)
    End If
    
    If mstrID = "" Then
        '����
        lngID = zlDatabase.GetNextId("����֧������")
        gstrSQL = "zl_����֧������_INSERT(" & lngID & "," & mlng���� & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng���� & "," & lng�㷨 & "," & _
                 dblͳ��ȶ� & "," & dbl��׼���� & "," & dbl��׼���� & "," & GetTextFromCombo(cmb�������, False) & "," & chkҽ��.Value & ")"
    Else
        gstrSQL = "zl_����֧������_Update(" & mstrID & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng���� & "," & lng�㷨 & "," & _
                 dblͳ��ȶ� & "," & dbl��׼���� & "," & dbl��׼���� & "," & GetTextFromCombo(cmb�������, False) & "," & chkҽ��.Value & ")"
    End If
    Call ExecuteProcedure(Me.Caption)
    
    '����������
    If mstrID = "" Then
        Set lst = frm���մ���.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text����), "Class", "Class")
    Else
        Set lst = frm���մ���.lvwItem.SelectedItem
        lst.Text = Trim(txtEdit(text����).Text)
    End If
    lst.SubItems(1) = Trim(txtEdit(Text����).Text)
    lst.SubItems(2) = Trim(txtEdit(Text����).Text)
    lst.SubItems(3) = Choose(lng����, "ҩƷ", "ҽ��", "����")
    lst.SubItems(4) = IIf(lng�㷨 = 1, "�ܶ����", "סԺ�պ˶�")
    lst.SubItems(5) = Mid(cmb�������.Text, 3)
    lst.SubItems(6) = IIf(chkҽ��.Value = 1, "��", "��")
    lst.Tag = dblͳ��ȶ� & ";" & dbl��׼���� & ";" & dbl��׼����
    
    Save��Ŀ = True
    mblnOK = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
'����:���������й�ҽ�����������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim lngIndex As Integer
    For lngIndex = text���� To Text����
        If txtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                zlControl.TxtSelAll txtEdit(lngIndex)
                Exit Function
            End If
            
            If lngIndex = text���� Or lngIndex = Text���� Then
                If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                    txtEdit(lngIndex).Text = ""
                    MsgBox "��������ƶ�����Ϊ�ա�", vbExclamation, gstrSysName
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
            End If
            
            If lngIndex >= Text���� Then
                If IsNumeric(txtEdit(lngIndex).Text) = False Then
                    MsgBox "������Ϸ�����ֵ��", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                        
                If Val(txtEdit(lngIndex).Text) < 0 Then
                    MsgBox "��ֵ����С��0��", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                
                If lngIndex = Text���� Then
                    If Val(txtEdit(Text����).Text) > 100 Then
                        MsgBox "ͳ��֧���������ܳ���100��", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(Text����)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                Else
                    If Val(txtEdit(lngIndex).Text) > 10000 Then
                        MsgBox "��ֵ���ܳ���10000��", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(lngIndex)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    '����׼��������׼��������
    If opt�㷨(CheckסԺ��).Value = True Then
        If Val(txtEdit(Text��׼).Text) = 0 And Val(txtEdit(Text����).Text) <> 0 Then
            MsgBox "��׼����Ϊ0����׼����Ҳ��Ϊ0��", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text����)
            txtEdit(Text����).SetFocus
            Exit Function
        End If
        If Val(txtEdit(Text��׼).Text) <> 0 And Val(txtEdit(Text����).Text) = 0 Then
            MsgBox "��׼����Ϊ0����׼����Ҳ��Ϊ0��", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text��׼)
            txtEdit(Text��׼).SetFocus
            Exit Function
        End If
        If Val(txtEdit(Text��׼).Text) <> 0 And Val(txtEdit(Text����).Text) > Val(txtEdit(Text��׼).Text) Then
            MsgBox "��������ܴ�����׼���", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text����)
            txtEdit(Text����).SetFocus
            Exit Function
        End If
    End If
    
    If chkҽ��.Value = 0 Then
        If MsgBox("���������������ҽ������Ӱ�쵽������������ҽ����Ŀ��" & vbCrLf & "�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            chkҽ��.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Sub opt�㷨_Click(Index As Integer)
    Dim bln���� As Boolean
    
    mblnChange = True
    txtEdit(Text����).Enabled = (opt�㷨(Check����).Value = True)
    lblEdit(Text����).Enabled = txtEdit(Text����).Enabled
    
    txtEdit(Text����).Enabled = (opt�㷨(CheckסԺ��).Value = True)
    txtEdit(Text��׼).Enabled = txtEdit(Text����).Enabled
    txtEdit(Text����).Enabled = txtEdit(Text����).Enabled
    lblEdit(Text����).Enabled = txtEdit(Text����).Enabled
    lblEdit(Text��׼).Enabled = txtEdit(Text����).Enabled
    lblEdit(Text����).Enabled = txtEdit(Text����).Enabled
End Sub

Private Sub opt�㷨_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub opt����_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}", 1
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text���� Then
        txtEdit(Text����).Text = zlCommFun.SpellCode(txtEdit(Text����).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text����
          zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        SendKeys "{Tab}", 1
    Else
        If Index = text���� Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
    If Index >= Text���� And Index <= Text��׼ Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
    End If
End Sub

Public Function �༭ҽ������(ByVal lng���� As Long, ByVal strID As String) As Boolean
'����:��������õ�ҽ���������ڽ���ͨѶ�ĳ���
'����:str���           ��ǰ�༭��ҽ�����ĵ����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnOK = False
    mlng���� = lng����
    mstrID = strID
    
    cmb�������.AddItem "1.���ﲡ��"
    cmb�������.AddItem "2.סԺ����"
    cmb�������.AddItem "3.���в���"
    cmb�������.ListIndex = 2
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '�޸�ҽ������
        gstrSQL = "select ����,����,����,nvl(����,1) as ����,nvl(�㷨,1) as �㷨 " & _
                  ",ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ��,nvl(�������,3) as ������� " & _
                  "from ����֧������ where ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "�ñ��մ����Ѿ���ɾ������ˢ�¡�", vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(text����).Text = rsTemp("����")
        txtEdit(Text����).Text = rsTemp("����")
        txtEdit(Text����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        Call SetComboByText(cmb�������, rsTemp("�������"), False)
        chkҽ��.Value = IIf(rsTemp("�Ƿ�ҽ��") = 1, 1, 0)
        opt����(rsTemp("����")).Value = True
        opt�㷨(rsTemp("�㷨")).Value = True
        Call opt�㷨_Click(rsTemp("�㷨"))
        If rsTemp("�㷨") = 1 Then
            '1-����������Ŀ
            txtEdit(Text����).Text = Format(rsTemp("ͳ��ȶ�"), "0.00")
        Else
            '2-סԺ�պ˶���Ŀ
            txtEdit(Text����).Text = Format(rsTemp("ͳ��ȶ�"), "0.00")
            txtEdit(Text��׼).Text = Format(rsTemp("��׼����"), "0.00")
            txtEdit(Text����).Text = Format(rsTemp("��׼����"), "0")
        End If
        
    Else
        '����ҽ������
        txtEdit(text����).Text = zlDatabase.GetMax("����֧������", "����", 6, " where ����=" & mlng����)
        opt�㷨(1).Value = True
        Call opt�㷨_Click(1)
    End If
    
    
    mblnChange = False
    frm���մ���༭.Show vbModal, frm���մ���
    �༭ҽ������ = mblnOK
End Function

