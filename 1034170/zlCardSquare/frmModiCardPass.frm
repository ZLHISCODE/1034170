VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmModiCardPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����޸�"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   Icon            =   "frmModiCardPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
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
      Height          =   420
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   6
      Top             =   870
      Width           =   1200
   End
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   285
      ScaleHeight     =   3735
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   285
      Width           =   5625
      Begin VB.TextBox txtAudi 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3030
         Width           =   4245
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2520
         Width           =   4245
      End
      Begin VB.TextBox txtOldPass 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2010
         Width           =   4245
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   5520
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�뽫[XX]��ˢ�����ϻ�����  Ȼ����������������ͬ�����룡"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   180
         Width           =   8550
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   120
         Picture         =   "frmModiCardPass.frx":058A
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LabelVeriPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   390
         TabIndex        =   10
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label lblNewPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   9
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label lblOldPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   8
         Top             =   2040
         Width           =   945
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
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   390
         TabIndex        =   7
         Top             =   1500
         Width           =   630
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4065
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5955
      _Version        =   589884
      _ExtentX        =   10504
      _ExtentY        =   7170
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiCardPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long, mlngCardTypeID As Long
Private mblnOk As Boolean, mblnCheckOldPass As Boolean
Private mobjKeyboard As Object, mblnTest As Boolean
Private mblnFirst As Boolean
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object
Private mrsInfo As New ADODB.Recordset, mrsCardType As New ADODB.Recordset

Public Function zlModifyPass(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional blnCheckOldPass As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ڲ���
    '���:frmMain-���õ�������
    '     lngModule -ģ���
    '     lngCardTypeID-���ѿ��ӿڱ��
    '����:�޸ĳɹ�,����true,���򷵻�false
    '����:������
    '����:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mblnOk = False: mlngCardTypeID = lngCardTypeID
    mblnCheckOldPass = blnCheckOldPass
    mblnTest = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard", "TestCardNO", 0)) = 1
    mblnTest = IsDesinMode Or mblnTest
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyPass = mblnOk
End Function

Private Function InitCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ƭ��Ϣ
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 14:25:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mrsCardType = zlGet���ѿ��ӿ�
    mrsCardType.Filter = "���=" & mlngCardTypeID
    If mrsCardType Is Nothing Then Exit Function
    
    lbl����.BorderStyle = Val(Nvl(mrsCardType!�Ƿ�ˢ��)): lbl����.Tag = Nvl(mrsCardType!�Ƿ�ˢ��)
    
    If Val(Nvl(mrsCardType!���볤��)) <> 0 Then
        txtOldPass.MaxLength = Val(Nvl(mrsCardType!���볤��))
        txtPass.MaxLength = Val(Nvl(mrsCardType!���볤��))
        txtAudi.MaxLength = Val(Nvl(mrsCardType!���볤��))
    Else
        txtOldPass.MaxLength = 10
        txtPass.MaxLength = 10
        txtAudi.MaxLength = 10
    End If
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & Nvl(mrsCardType!����) & "]")
    InitCardInfor = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ƿ���Ч
    '����:������Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 11:15:42
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    On Error GoTo errHandle
    If CheckCard(mlngCardTypeID, txt����.Text) = False Then Exit Function
    If mrsInfo Is Nothing Then
        MsgBox "���ܶ�ȡ����Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
        Call ClearFace: txt����.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "���ܶ�ȡ����Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
        Call ClearFace: txt����.SetFocus: Exit Function
    End If
    If txtPass.Text <> txtAudi.Text Then
        MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        Exit Function
    End If
    If txtPass.Text = "" Then
        If MsgBox("��ǰ���õ�����Ϊ�գ�ȷʵҪ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function ModifCardPass() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ŀ�Ƭ������
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:������
    '����:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, intForce As Integer
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then Exit Function
    strPassWord = zlCommFun.zlStringEncode(txtOldPass.Text)     '�������
    
    If strPassWord <> Nvl(mrsInfo!����) And mblnCheckOldPass = True Then
        MsgBox "��Ƭԭ�����������,��������������!", vbInformation, gstrSysName
        txtOldPass.SetFocus
        ModifCardPass = False
        Exit Function
    End If
    
    If mblnCheckOldPass = True Then
        intForce = 0
    Else
        intForce = 1
    End If
    
     'Zl_���ѿ�����_Update
    strSQL = "Zl_���ѿ�����_Update('" & Nvl(mrsInfo!����) & "'," & mlngCardTypeID & "," & _
             Val(Nvl(mrsInfo!���)) & ",'" & strPassWord & "','" & zlCommFun.zlStringEncode(txtPass.Text) & "'," & intForce & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    ModifCardPass = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If ModifCardPass = False Then Exit Sub
    MsgBox "�����޸ĳɹ�!", vbOKOnly + vbInformation, gstrSysName
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitCardInfor = False Then Unload Me: Exit Sub
    Call ClearFace
    Call txt����_Change
    txt����.SetFocus
End Sub

Private Sub Form_Load()
    mblnFirst = True
    Set mrsInfo = Nothing
    Call CreateObjectKeyboard

    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    
    If mblnCheckOldPass = False Then
        lblNotes.Top = 180
        lblNotes.Caption = "�뽫[XX]��ˢ�����ϻ�����" & vbCrLf & "��������������ͬ�������룡"
        txtOldPass.Enabled = False
        txtOldPass.BackColor = &H8000000F
    Else
        lblNotes.Top = 180
        lblNotes.Caption = "�뽫[XX]��ˢ�����ϻ�����" & vbCrLf & "�����������������ͬ�������룡"
    End If
    HookDefend txtOldPass.hWnd
    HookDefend txtPass.hWnd
    HookDefend txtAudi.hWnd
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    txt����.Text = ""
    Set mobjCommEvents = Nothing
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    If strCardNo = "" Then Exit Sub
    If Not GetCardPass(strCardType, strCardNo) Then
        Call ClearFace: If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    OpenPassKeyboard txtAudi, True
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: cmdOK_Click
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAudi_LostFocus()
   ClosePassKeyboard txtPass
End Sub

Private Sub txtOldPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    OpenPassKeyboard txtPass, False
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            cmdOk.SetFocus
        Else
            txtAudi.SetFocus
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub
Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtOldPass_GotFocus()
    zlControl.TxtSelAll txtOldPass
End Sub

Private Function CheckCard(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ǰ��鿨����Ч��
    '���:lngCardTypeID-���ѿ��ӿڱ��
    '     strCardNO-����
    '����:�ɹ�����True,ʧ�ܷ���False
    '����:������
    '����:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select a.Id,a.������,a.����,a.���,a.�ɷ��ֵ,a.�ӿڱ��,to_char(a.��Ч��,'yyyy-mm-dd hh24:mi:ss') as ��Ч��,  a.����," & _
    "          to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� , " & _
    "          decode(a.��ǰ״̬,2,'����',3,'�˿�','����') as ��ǰ״̬, " & _
    "          to_char(a.ͣ������,'yyyy-mm-dd hh24:mi:ss') as ͣ������," & _
    "          a.������� " & _
    "   From ���ѿ�Ŀ¼ A  " & _
    "   Where A.���� = [1] and A.�ӿڱ��=[2] And ��� = (Select Max(���) From ���ѿ�Ŀ¼ B Where ���� = A.���� and �ӿڱ��=A.�ӿڱ��)  " & _
    "   Order by a.���"
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, lngCardTypeID)
    If mrsInfo.EOF Then
        ShowMsgbox "δ�ҵ���ص�" & Nvl(mrsCardType!����) & "��Ϣ,����!"
        Exit Function
    End If
    
    '��鵱ǰˢ���ĺϷ���
    '�Ƿ����
    If Nvl(mrsInfo!����ʱ��, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & Nvl(mrsCardType!����) & "�Ѿ���" & Nvl(mrsInfo!��ǰ״̬) & ",������ˢ��"
        Exit Function
    End If
    
    '�Ƿ�ͣ��
    If Nvl(mrsInfo!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & Nvl(mrsCardType!����) & "�Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    '�Ƿ�ͣ��
    If Nvl(mrsInfo!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & Nvl(mrsCardType!����) & "�Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    
    CheckCard = True
    Exit Function
errH:
    CheckCard = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
End Function

Private Function GetCardPass(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������
    '���:lngCardTypeID-���ѿ��ӿڱ��
    '     strCardNO-����
    '����:�ɹ�����True,ʧ�ܷ���False
    '����:������
    '����:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    
    txtPass.Text = "": txtAudi.Text = "": txtOldPass.Text = ""
    
    strSQL = "" & _
    "   Select a.Id,a.������,a.����,a.���,a.�ɷ��ֵ,a.�ӿڱ��,to_char(a.��Ч��,'yyyy-mm-dd hh24:mi:ss') as ��Ч��,  a.����," & _
    "          to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� , " & _
    "          decode(a.��ǰ״̬,2,'����',3,'�˿�','����') as ��ǰ״̬, " & _
    "          to_char(a.ͣ������,'yyyy-mm-dd hh24:mi:ss') as ͣ������," & _
    "          a.������� " & _
    "   From ���ѿ�Ŀ¼ A  " & _
    "   Where A.���� = [1] and A.�ӿڱ��=[2] And ��� = (Select Max(���) From ���ѿ�Ŀ¼ B Where ���� = A.���� and �ӿڱ��=A.�ӿڱ��)  " & _
    "   Order by a.���"
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, lngCardTypeID)
    If mrsInfo.EOF Then
        ShowMsgbox "δ�ҵ���ص�" & Nvl(mrsCardType!����) & "��Ϣ,����!"
        Exit Function
    End If
    
    '��鵱ǰˢ���ĺϷ���
    '�Ƿ����
    If Nvl(mrsInfo!����ʱ��, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & Nvl(mrsCardType!����) & "�Ѿ���" & Nvl(mrsInfo!��ǰ״̬) & ",������ˢ��"
        Exit Function
    End If
    
    '�Ƿ�ͣ��
    If Nvl(mrsInfo!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & Nvl(mrsCardType!����) & "�Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    '�Ƿ�ͣ��
    If Nvl(mrsInfo!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & Nvl(mrsCardType!����) & "�Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    
    GetCardPass = True
    Exit Function
errH:
    GetCardPass = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
    Exit Function
End Function

Private Sub ClearFace()
    txt����.PasswordChar = IIf(Val(Nvl(mrsCardType!�Ƿ�����)) <> 0, "*", "")
    txt����.Text = ""
    txtPass.Text = "": txtAudi.Text = ""
End Sub

Private Sub txt����_Change()
    If mblnCheckOldPass = True Then txtOldPass.Enabled = Trim(txt����.Text) <> ""
    txtPass.Enabled = Trim(txt����.Text) <> ""
    txtAudi.Enabled = Trim(txt����.Text) <> ""
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    txt����.PasswordChar = IIf(Val(Nvl(mrsCardType!�Ƿ�����)) <> 0, "*", "")

    If mobjSquare Is Nothing Then Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    '��ʼ����Ƶ������
    mobjSquare.zlInitEvents Me.hWnd, mobjCommEvents
    mobjSquare.SetEnabled True
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean

    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(txt����.Text) = Val(Nvl(mrsCardType!���ų���)) - 1 And txt����.SelLength <> Len(txt����.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txt����.Text = txt����.Text & Chr(KeyAscii)
            txt����.SelStart = Len(txt����.Text)
        End If
        KeyAscii = 0
        If GetCardPass(mlngCardTypeID, Trim(txt����.Text)) = False Then
            If txt����.Enabled Then txt����.SetFocus
            zlControl.TxtSelAll txt����
            Exit Sub
        End If
        If mblnCheckOldPass Then
            If txtOldPass.Enabled Then txtOldPass.SetFocus
        Else
            If txtPass.Enabled Then txtPass.SetFocus: Exit Sub
        End If
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        If mblnTest Then Exit Sub
        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If txt����.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txt����.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txt����.Text = Chr(KeyAscii)
                txt����.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub

Private Sub txt����_LostFocus()
   If Not mobjSquare Is Nothing Then mobjSquare.SetEnabled False
End Sub

