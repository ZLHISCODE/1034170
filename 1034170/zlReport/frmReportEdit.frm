VERSION 5.00
Begin VB.Form frmReportEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmReportEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5595
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   4
      Top             =   810
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4335
      TabIndex        =   3
      Top             =   345
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   4050
      Begin VB.TextBox txt˵�� 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   735
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1125
         Width           =   3000
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   735
         MaxLength       =   40
         TabIndex        =   1
         Top             =   705
         Width           =   3000
      End
      Begin VB.TextBox txt��� 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   735
         MaxLength       =   20
         TabIndex        =   0
         Top             =   285
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "˵��"
         Height          =   180
         Left            =   285
         TabIndex        =   8
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   285
         TabIndex        =   7
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   285
         TabIndex        =   6
         Top             =   345
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmReportEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnGroupEdit As Boolean
Private mlngSys As Long
Private mlngReortID As Long
Private mlngGroupID As Long
Private mstr���� As String
Private mstrOld���� As String
Private mstr���� As String
Private mstr˵�� As String
Private mstrOld˵�� As String
Private mblnOK As Boolean
Private mlngModule As Long

Public Function ShowMe(ByVal frmParent As Object, ByVal lngSys As Long, ByVal blnGroupEdit As Boolean, ByVal lngModule As Long, Optional ByRef LngGroupID As Long, _
                                        Optional ByRef lngReortID As Long, Optional ByRef str���� As String, Optional ByRef str���� As String, Optional ByRef str˵�� As String) As Boolean
    mblnGroupEdit = blnGroupEdit
    mlngSys = lngSys
    mlngModule = lngModule
    mlngReortID = lngReortID
    mlngGroupID = LngGroupID
    mstr���� = str����: mstrOld���� = str����
    mstr���� = str����
    mstr˵�� = str˵��: mstrOld˵�� = str˵��
    Me.Show 1, frmParent
    str���� = mstr����
    str���� = mstr����
    str˵�� = mstr˵��
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsCheck As New ADODB.Recordset
    Dim strSQL As String
    Dim strOldName As String, strOld˵�� As String
    Dim intOrder As Integer
    Dim arrSQL() As Variant
    Dim i As Long, blnTrans As Boolean
    
    arrSQL = Array()
    If Not CheckFormInput(Me) Then Exit Sub
    
    If Trim(txt���.Text) = "" Then
        MsgBox "�����뱨��" & IIF(mblnGroupEdit, "��", "") & "�ı�ţ�", vbInformation, App.Title
        txt���.SetFocus: Exit Sub
    End If
    If Trim(txt����.Text) = "" Then
        MsgBox "�����뱨��" & IIF(mblnGroupEdit, "��", "") & "�����ƣ�", vbInformation, App.Title
        txt����.SetFocus: Exit Sub
    Else
        txt����.Text = ConvertSBC(txt����.Text)
    End If
    
    If Not CheckLen(txt���, 20, "���") Then Exit Sub
    If Not CheckLen(txt����, 30, "����") Then Exit Sub
    If Not CheckLen(txt˵��, 255, "˵��") Then Exit Sub
    
    '��Ų����ظ�(����������)
    If CheckExist("zlReports", "���", txt���.Text, mlngReortID) Then
        MsgBox "�ñ���Ѿ���ʹ��,���������룡", vbInformation, App.Title
        txt���.SetFocus: Exit Sub
    End If
    If CheckExist("zlRPTGroups", "���", txt���.Text, mlngGroupID) Then
        MsgBox "�ñ���Ѿ���ʹ��,���������룡", vbInformation, App.Title
        txt���.SetFocus: Exit Sub
    End If
    If mlngGroupID <> 0 And Not mblnGroupEdit Then
        strSQL = "Select 1 From zlRPTSubs A,zlReports B Where B.����=[1] And A.����ID=B.ID And A.��ID=[2]" & IIF(mlngReortID = 0, "", " And ����ID<>[3]")
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, txt����.Text, mlngGroupID, mlngReortID)
        If Not rsCheck.EOF Then
            MsgBox "�ñ��������Ѿ�������ͬ���Ƶı���", vbInformation, App.Title
            txt����.SetFocus: Exit Sub
        End If
    End If
    strOldName = mstrOld����: strOld˵�� = mstrOld˵��
    mstr���� = txt����.Text: mstr���� = txt���.Text: mstr˵�� = txt˵��.Text
    On Error GoTo errH
    If mblnGroupEdit Then
        If mlngGroupID = 0 Then
            mlngGroupID = GetNextID("zlRPTGroups")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Insert Into zlRPTGroups(ID,���,����,˵��) Values(" & mlngGroupID & ",'" & mstr���� & "','" & mstr���� & "','" & mstr˵�� & "')"
        ElseIf Not (strOldName = mstr���� And strOld˵�� = mstr˵��) Then '˵�������Ʒ����仯
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Update zlRPTGroups Set ���='" & mstr���� & "',����='" & mstr���� & "',˵��='" & mstr˵�� & "' Where ID=" & mlngGroupID
            '����������̨�˵��ı������
            If mlngModule <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlPrograms Set ����='" & mstr���� & "',˵��='" & mstr˵�� & "' Where ���=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlMenus Set ����='" & mstr���� & "',�̱���='" & mstr���� & "',˵��='" & mstr˵�� & "' Where ID=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys
            End If
        End If
    Else
        If mlngReortID = 0 Then
            If mlngSys <> 0 Then mlngSys = 0
            mlngReortID = GetNextID("zlReports")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Insert Into zlReports(ID,���,����,˵��,ϵͳ,�޸�ʱ��,����) Values(" & _
                                                        mlngReortID & ",'" & mstr���� & "','" & mstr���� & "','" & mstr˵�� & "'," & IIF(mlngSys = 0, "NULL", mlngSys) & ",Sysdate," & AdjustStr(GetPass(mstr����, mstr����)) & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(" & _
                                                        mlngReortID & ",1,'" & mstr���� & "1'," & INIT_WIDTH & "," & INIT_HEIGHT & ",9,1,0,0)"

            If mlngGroupID <> 0 Then
                intOrder = 1
                strSQL = "Select Count(*) Records From zlRPTSubs Where ��ID=[1]"
                Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
                If Not rsCheck.EOF Then intOrder = Nvl(rsCheck!Records, 0) + 1
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Insert Into zlRPTSubs(��ID,����ID,���,����) " & _
                                         "Values(" & mlngGroupID & "," & mlngReortID & "," & intOrder & ",'" & mstr���� & "')"
                If mlngModule <> 0 Then '����Ȩ�޼�¼
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(" & _
                                                                IIF(mlngSys = 0, "NULL", mlngSys) & "," & mlngModule & ",'" & mstr���� & "','" & mstr˵�� & "')"
                End If
            End If
        ElseIf Not (strOldName = mstr���� And strOld˵�� = mstr˵��) Then '˵�������Ʒ����仯
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Update zlReports Set ���='" & mstr���� & "',����='" & _
                                     mstr���� & "',˵��='" & mstr˵�� & "',����=" & AdjustStr(GetPass(mstr����, mstr����)) & " Where ID=" & mlngReortID
            If mlngModule <> 0 Then '����������̨�˵��ı������
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlPrograms Set ����='" & mstr���� & "',˵��='" & mstr˵�� & "'" & _
                                        " Where Upper(����)=Upper('zl9Report') And ���=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlMenus  Set ����='" & mstr���� & "',�̱���='" & mstr���� & "',˵��='" & mstr˵�� & "'" & _
                                        " Where ģ��=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys & _
                                        " And Exists(Select ���� From zlPrograms Where Upper(����)=Upper('zl9Report') And ���=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys & ")"
            
            End If
            '����������̨�ı������ӱ�Ĺ�����
            strSQL = "Select Distinct Nvl(B.ϵͳ, 0) ϵͳ, B.����id ���, a.��Id " & vbNewLine & _
                     "From Zlrptsubs a, Zlrptgroups b, Zlprograms c" & vbNewLine & _
                     "Where A.��id = B.Id And A.����id = [1]  And Nvl(B.ϵͳ, 0) = Nvl(C.ϵͳ, 0) And B.����id = C.��� And Upper(C.����) = Upper('zl9Report')"
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReortID)
            Do While Not rsCheck.EOF
                If strOldName <> mstr���� Then  '�������Ʒ����仯
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '�����ӱ�������
                        arrSQL(UBound(arrSQL)) = _
                            "Update zlRPTSubs " & vbNewLine & _
                            "Set ���� = '" & mstr���� & "' " & vbNewLine & _
                            "Where ��Id = " & Nvl(rsCheck!��Id) & _
                            "    And ����Id = " & mlngReortID & " And ���� = '" & strOldName & "'"
                            
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ϣ
                    arrSQL(UBound(arrSQL)) = "Insert Into Zlprogfuncs" & vbNewLine & _
                                            "  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)" & vbNewLine & _
                                            "  Select A.ϵͳ, A.���, '" & mstr���� & "', A.����, '" & mstr˵�� & "', A.ȱʡֵ" & vbNewLine & _
                                            "  From Zlprogfuncs a" & vbNewLine & _
                                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & strOldName & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ȩ��Ϣ
                    arrSQL(UBound(arrSQL)) = "Insert Into zlrolegrant" & vbNewLine & _
                                            "  (ϵͳ,���,��ɫ,����)" & vbNewLine & _
                                            "  Select A.ϵͳ,A.���,A.��ɫ, '" & mstr���� & "' " & vbNewLine & _
                                            "  From zlrolegrant a" & vbNewLine & _
                                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & strOldName & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ���ܶ���Ȩ����Ϣ
                    arrSQL(UBound(arrSQL)) = "Insert Into zlprogprivs" & vbNewLine & _
                                            "  (ϵͳ,���,����,����,������,Ȩ��)" & vbNewLine & _
                                            "  Select A.ϵͳ,A.���,'" & mstr���� & "',A.����,A.������,A.Ȩ��" & vbNewLine & _
                                            "  From zlprogprivs a" & vbNewLine & _
                                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & strOldName & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) 'ɾ��ԭʼ�����������ڴ��ڼ���ɾ����ϵ
                    arrSQL(UBound(arrSQL)) = "  Delete From Zlprogfuncs a" & vbNewLine & _
                                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & strOldName & "'"
                    'ϵͳ����š����� ����һ������Null������ɾ����ʧЧ
                    If Nvl(rsCheck!ϵͳ, 0) = 0 Or Nvl(rsCheck!���, 0) = 0 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From zlProgPrivs A " & vbNewLine & _
                            "Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "    And A.���� = '" & strOldName & "'"
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From zlRoleGrant A " & vbNewLine & _
                            "Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "    And A.���� = '" & strOldName & "'"
                    End If
                Else '��������δ�����仯,ֻ����¹���˵��
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '���¹���˵��
                    arrSQL(UBound(arrSQL)) = "Update Zlprogfuncs A" & vbNewLine & _
                                                                "  Set  A.˵��='" & mstr˵�� & "'" & vbNewLine & _
                                                                "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & mstr���� & "'"
                End If
                rsCheck.MoveNext
            Loop
            '������ģ��ı�������
            strSQL = "Select Nvl(B.ϵͳ, 0) ϵͳ, B.����id ���, B.����" & vbNewLine & _
                            "From Zlrptputs b, Zlprograms c, Zlprogfuncs d" & vbNewLine & _
                            "Where B.����id =[1] And Nvl(B.ϵͳ, 0) = Nvl(C.ϵͳ, 0) And B.����id = C.��� And" & vbNewLine & _
                            "      Upper(C.����) <> Upper('zl9Report') And Nvl(C.ϵͳ, 0) = Nvl(D.ϵͳ, 0) And C.��� = D.��� And D.���� = B.����"
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReortID)
            Do While Not rsCheck.EOF
                If strOldName <> mstr���� And mlngSys = 0 Then   '��ϵͳ�������Ʒ����仯�����Զ����¹�������
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����Zlrptputs
                    arrSQL(UBound(arrSQL)) = "Update Zlrptputs Set ���� = '" & mstr���� & "' Where ����id = " & mlngReortID & " And Nvl(ϵͳ, 0) = " & rsCheck!ϵͳ & " And ����id = " & rsCheck!���
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ϣ
                    arrSQL(UBound(arrSQL)) = "Insert Into Zlprogfuncs" & vbNewLine & _
                                                                "  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)" & vbNewLine & _
                                                                "  Select A.ϵͳ, A.���, '" & mstr���� & "', A.����, '" & mstr˵�� & "', A.ȱʡֵ" & vbNewLine & _
                                                                "  From Zlprogfuncs a" & vbNewLine & _
                                                                "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & rsCheck!���� & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ȩ��Ϣ
                    arrSQL(UBound(arrSQL)) = "Insert Into zlrolegrant" & vbNewLine & _
                                                                "  (ϵͳ,���,��ɫ,����)" & vbNewLine & _
                                                                "  Select A.ϵͳ,A.���,A.��ɫ, '" & mstr���� & "' " & vbNewLine & _
                                                                "  From zlrolegrant a" & vbNewLine & _
                                                                "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & rsCheck!���� & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ���ܶ���Ȩ����Ϣ
                    arrSQL(UBound(arrSQL)) = "Insert Into zlprogprivs" & vbNewLine & _
                                                                "  (ϵͳ,���,����,����,������,Ȩ��)" & vbNewLine & _
                                                                "  Select A.ϵͳ,A.���,'" & mstr���� & "',A.����,A.������,A.Ȩ��" & vbNewLine & _
                                                                "  From zlprogprivs a" & vbNewLine & _
                                                                "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & rsCheck!���� & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) 'ɾ��ԭʼ�����������ڴ��ڼ���ɾ����ϵ
                    arrSQL(UBound(arrSQL)) = "  Delete From Zlprogfuncs a" & vbNewLine & _
                                                                "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & rsCheck!���� & "'"
                Else '��ϵͳ����˵���仯���߹̶�����������ֻ���¹���˵��
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '���¹���˵��
                    arrSQL(UBound(arrSQL)) = "Update Zlprogfuncs A" & vbNewLine & _
                                                                "  Set  A.˵��='" & mstr˵�� & "'" & vbNewLine & _
                                                                "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & " And A.���� = '" & rsCheck!���� & "'"
                End If
                rsCheck.MoveNext
            Loop
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Debug.Print arrSQL(i)
        gcnOracle.Execute arrSQL(i)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Set grsReport = Nothing '�������
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnGroupEdit And mlngGroupID <> 0 Or Not mblnGroupEdit And mlngReortID <> 0 Then txt����.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    mblnOK = False
    txt���.Text = mstr����
    txt����.Text = mstr����
    txt˵��.Text = mstr˵��
    If mblnGroupEdit Then
        If mlngGroupID = 0 Then
            Caption = "����������"
            txt���.Text = GetNextNO(mblnGroupEdit)
        Else
            Caption = "�޸ı�����"
        End If
    Else
        If mlngReortID = 0 Then
            Caption = "��������"
            txt���.Text = GetNextNO(mblnGroupEdit)
        Else
            Caption = "�޸ı���"
        End If
    End If
    If mlngSys > 0 Then txt���.Enabled = False
End Sub

Private Sub txt���_GotFocus()
    SelAll txt���
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf InStr(GSTR_SBC, Chr(KeyAscii)) > 0 Then
        KeyAscii = Asc(Mid(GSTR_DBC, InStr(GSTR_SBC, Chr(KeyAscii)), 1))
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If txt����.Text <> "" Then
        txt����.Text = ConvertSBC(txt����.Text)
    End If
End Sub

Private Sub txt˵��_GotFocus()
    SelAll txt˵��
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


