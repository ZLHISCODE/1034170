VERSION 5.00
Begin VB.Form frmSet�ɶ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk���� 
      Caption         =   "��Ų�������(&L)"
      Height          =   270
      Left            =   1200
      TabIndex        =   8
      Top             =   3330
      Width           =   3120
   End
   Begin VB.ComboBox cbo�籣���� 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2955
      Width           =   3315
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   16
      Top             =   3720
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   660
      Width           =   7665
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1980
      Left            =   435
      TabIndex        =   12
      Top             =   825
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   2955
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   555
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1335
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1200
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   945
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   1
         Top             =   555
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   330
         TabIndex        =   4
         Top             =   1395
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   2
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   615
         Width           =   810
      End
   End
   Begin VB.CommandButton cmd�籣���� 
      Caption         =   "�����籣����(&D)"
      Height          =   350
      Left            =   135
      TabIndex        =   11
      Top             =   3870
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2385
      TabIndex        =   9
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   10
      Top             =   3870
      Width           =   1100
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet�ɶ�����.frx":0000
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�籣����"
      Height          =   180
      Index           =   1
      Left            =   390
      TabIndex        =   6
      Top             =   3030
      Width           =   720
   End
   Begin VB.Label lbl 
      Caption         =   "������صĲ���."
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   15
      Top             =   360
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet�ɶ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Enum enum�ı�
    Textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
End Enum



Public Function ��������() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    frmSet�ɶ�����.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo�籣����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(Textҽ���û�).Text, txtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub
Private Sub cmd�籣����_Click()
    Dim strOutPut As String, strInPut As String
    Dim strArr, strArr1
    Dim i As Long
    If mcnTest Is Nothing Then
        MsgBox "���Ȳ����м���Ƿ�ɳ�!"
        Exit Sub
    End If
    If mcnTest.State <> 1 Then
        MsgBox "���Ȳ����м���Ƿ�ɳ�!"
        Exit Sub
    End If
    If ҽ����ʼ��_�ɶ����� = False Then Exit Sub
    
    zlCommFun.ShowFlash "���������籣����,���Ժ�..."
    If ҵ������_�ɶ�����(����籣����, strInPut, strOutPut) = False Then
        zlCommFun.StopFlash
        Exit Sub
    End If
    If strOutPut = "" Then
        zlCommFun.StopFlash
        Exit Sub
    End If
    strArr = Split(strOutPut, "@$")
    For i = 0 To UBound(strArr)
        strArr1 = Split(strArr(i), "||")
        '�����籣����
        gstrSQL = "ZL_�籣����Ŀ¼_UPDATE("
        gstrSQL = gstrSQL & "'" & strArr1(0) & "',"
        gstrSQL = gstrSQL & "'" & strArr1(1) & "')"
        gcnOracle_�ɶ�����.Execute gstrSQL, , adCmdStoredProc
    Next
    '���¼�������
    Call LoadCbo
    zlCommFun.StopFlash
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    
    mblnFirst = False
    Call LoadCbo
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_�ɶ�����
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "ҽ���û���"
                  txtEdit(Textҽ���û�).Text = Nvl(!����ֵ)
            Case "ҽ���û�����"
                  txtEdit(Textҽ������).Text = Nvl(!����ֵ)
            Case "ҽ��������"
                  txtEdit(Textҽ��������).Text = Nvl(!����ֵ)
            Case "��鲦������"
                    chk����.Value = IIf(Nvl(!����ֵ, 1) = 1, 1, 0)
            End Select
            .MoveNext
        Loop
    End With
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Textҽ������ Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = Textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(Textҽ���û�).Text, txtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�ɶ����� & ",null)"
    Call ExecuteProcedure(Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ����� & ",null,'ҽ���û���','" & txtEdit(Textҽ���û�).Text & "',1)"
    Call ExecuteProcedure(Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ����� & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call ExecuteProcedure(Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ����� & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
    Call ExecuteProcedure(Me.Caption)
    
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ����� & ",null,'��鲦������','" & IIf(chk����.Value = 1, 1, 0) & "',4)"
    Call ExecuteProcedure(Me.Caption)
    
    gcnOracle.CommitTrans
    If cbo�籣����.ListIndex >= 0 Then
        SaveRegInFor g����ģ��, "ҽ��", "�籣��������", Split(cbo�籣����.Text, "--")(0)
    End If
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub LoadCbo()
        '����Grid����
        Err = 0
        On Error GoTo ErrHand:
        Dim rsTemp As New ADODB.Recordset
        Dim i As Long
        gstrSQL = "Select * From �籣����Ŀ¼"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�籣����Ŀ¼" '  rsTemp.Open gstrSQL, gcnOracle_�ɶ�����
        With rsTemp
            i = 1
            Me.cbo�籣����.Clear
            Do While Not .EOF
                cbo�籣����.AddItem Nvl(!����) & "--" & Nvl(!����)
                .MoveNext
            Loop
        End With
        SetDefaultSel
        Exit Sub
ErrHand:
        If ErrCenter = 1 Then Resume
End Sub
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo ErrHand:
    Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
    If cbo�籣����.ListCount = 0 Then Exit Function
    For i = 0 To cbo�籣����.ListCount
        If Split(cbo�籣����.List(i), "--")(0) = strReg Then
            cbo�籣����.ListIndex = i
            Exit For
        End If
    Next
    If cbo�籣����.ListIndex < 0 Then
        cbo�籣����.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
