VERSION 5.00
Begin VB.Form frmAgentInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������Ϣ"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "frmAgentInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   10
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1100
   End
   Begin VB.Frame fraAgent 
      Caption         =   "������"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3615
      Begin VB.TextBox txtAgentName 
         Height          =   300
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   1
         Top             =   240
         Width           =   2130
      End
      Begin VB.TextBox txtAgentIDNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   2
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label lblAgentName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   600
         TabIndex        =   8
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblAgentIDNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.Frame fraPati 
      Caption         =   "����"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtPatiIDNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   0
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label lblPatiName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����  ���"
         Height          =   180
         Left            =   600
         TabIndex        =   5
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblPatiIDNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmAgentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlng����ID As Long
Private mlng����ID As Long
Private mstr���� As String
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1


Public Function ShowMe(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str�������� As String, _
                ByVal str�������֤�� As String, ByVal str���������� As String, ByVal str���������֤�� As String) As Boolean
    Screen.MousePointer = 0
    mblnOK = False
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mstr���� = str��������
    
    lblPatiName.Caption = "����  " & str��������
    txtPatiIDNO.Text = str�������֤��
    txtAgentName.Text = str����������
    txtAgentIDNO.Text = str���������֤��
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandle
    
    If Trim(txtPatiIDNO.Text) = "" Then
        MsgBox "�����벡�����֤�ţ�", vbInformation, gstrSysName
        txtPatiIDNO.SetFocus: Exit Sub
    End If
    
    If Len(txtPatiIDNO.Text) <> 15 And Len(txtPatiIDNO.Text) <> 18 Then
        MsgBox "���֤�ų��Ȳ���ȷ��������15��18λ���֤�ţ�", vbInformation, gstrSysName
        txtPatiIDNO.SetFocus: Exit Sub
    End If
'    If Trim(txtAgentName.Text) = "" Then
'        MsgBox "�����������������", vbInformation, gstrSysName
'        txtAgentName.SetFocus: Exit Sub
'    End If
'
'    If Trim(txtAgentIDNO.Text) = "" Then
'        MsgBox "��������������֤�ţ�", vbInformation, gstrSysName
'        txtAgentIDNO.SetFocus: Exit Sub
'    End If

    If txtAgentIDNO.Text <> "" And Len(txtAgentIDNO.Text) <> 15 And Len(txtAgentIDNO.Text) <> 18 Then
        MsgBox "���֤�ų��Ȳ���ȷ��������15��18λ���֤�ţ�", vbInformation, gstrSysName
        txtAgentIDNO.SetFocus: Exit Sub
    End If
    
    If Trim(txtAgentIDNO.Text) = Trim(txtPatiIDNO.Text) Then
        MsgBox "���������֤���벡�����֤����ͬ�����������룡", vbInformation, gstrSysName
        txtAgentIDNO.SetFocus: Exit Sub
    End If
    
    Screen.MousePointer = 11
    gstrSQL = "Zl_��������Ϣ_Insert(" & mlng����ID & ",'" & Trim(txtPatiIDNO.Text) & "'," & IIF(Trim(txtAgentName.Text) = "", "Null", "'" & Trim(txtAgentName.Text) & "'") & ",'" & _
                Trim(txtAgentIDNO.Text) & "'," & mlng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Screen.MousePointer = 0
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If txtPatiIDNO.Text <> "" Then
        txtAgentName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If InStr(GetInsidePrivs(p����ҽ��վ), "��������Ϣ��������¼��") = 0 Then
        txtPatiIDNO.Locked = True
        txtAgentName.Locked = True
        txtAgentIDNO.Locked = True
    End If
    Me.Caption = Me.Caption & "  (���֤ˢ��¼��)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Screen.MousePointer = 11
End Sub

Private Sub txtAgentIDNO_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtAgentIDNO)
End Sub


Private Sub txtAgentIDNO_GotFocus()
    zlControl.TxtSelAll txtAgentIDNO
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtAgentIDNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789X" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtAgentIDNO_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtAgentIDNO.Text = Trim(txtAgentIDNO.Text)
End Sub

Private Sub txtAgentName_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtAgentName)
End Sub


Private Sub txtAgentName_GotFocus()
    zlControl.TxtSelAll txtAgentName
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub


Private Sub txtAgentName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtAgentName_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtAgentName.Text = Trim(txtAgentName.Text)
End Sub


Private Sub txtPatiIDNO_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtPatiIDNO)
End Sub


Private Sub txtPatiIDNO_GotFocus()
    zlControl.TxtSelAll txtPatiIDNO
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub


Private Sub txtPatiIDNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789X" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPatiIDNO_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtPatiIDNO.Text = Trim(txtPatiIDNO.Text)
End Sub



Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
On Error GoTo errH
    If Me.ActiveControl Is txtPatiIDNO Then
        If mstr���� = strName Then
            txtPatiIDNO.Text = strID
        Else
            MsgBox "�����Ϣ¼��ʧ��,��ʹ�õ�ǰ���˵����֤ˢ����", vbInformation, gstrSysName
        End If
    ElseIf Me.ActiveControl Is txtAgentName Or Me.ActiveControl Is txtAgentIDNO Then
        txtAgentName.Text = strName
        txtAgentIDNO.Text = strID
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetOldAcademic(ByVal DateBir As Date, ByVal str���䵥λ As String) As Long
'���ܣ����ݵ�ǰ�ĳ������ں����䵥λ�����������ϵ�����ֵ
'���أ�����
    Dim datCur As Date, lngOld As Long, strInterval As String
    If DateBir = CDate(0) Or InStr(" ������", str���䵥λ) < 2 Then Exit Function
    
    datCur = zlDatabase.Currentdate
    
    strInterval = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
    lngOld = DateDiff(strInterval, DateBir, datCur)
    If DateAdd(strInterval, lngOld, DateBir) > datCur Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function


Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����֤������
    '����:����
    '����:2012-08-31 16:28:23
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (blnEnabled)
End Sub

