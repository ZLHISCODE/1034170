VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicOfficeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicOfficeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdRemove 
      Caption         =   "�Ƴ�(&D)"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4260
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2298
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "��"
      Height          =   345
      Left            =   2910
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2298
      Width           =   345
   End
   Begin VB.TextBox txtSelect 
      Height          =   350
      Left            =   960
      TabIndex        =   12
      Top             =   2293
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwDept 
      Height          =   2145
      Left            =   120
      TabIndex        =   15
      Top             =   2670
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   3784
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���ÿ���"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Frame frmSplit 
      Height          =   5205
      Left            =   5220
      TabIndex        =   16
      Top             =   -150
      Width           =   30
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   5460
      TabIndex        =   17
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5460
      TabIndex        =   18
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   360
      Left            =   5460
      TabIndex        =   19
      Top             =   4290
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ"
      Height          =   2055
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5085
      Begin VB.TextBox txt���� 
         Height          =   350
         Left            =   660
         MaxLength       =   3
         TabIndex        =   2
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txtλ�� 
         Height          =   350
         Left            =   660
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1650
         Width           =   4335
      End
      Begin VB.TextBox txt���� 
         Height          =   350
         Left            =   660
         MaxLength       =   20
         TabIndex        =   4
         Top             =   765
         Width           =   4335
      End
      Begin VB.TextBox txt���� 
         Height          =   350
         Left            =   660
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1215
         Width           =   1245
      End
      Begin VB.ComboBox cboStationNo 
         Height          =   330
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1225
         Width           =   2205
      End
      Begin VB.Label lblλ�� 
         AutoSize        =   -1  'True
         Caption         =   "λ��"
         Height          =   210
         Left            =   210
         TabIndex        =   9
         Top             =   1720
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   210
         TabIndex        =   1
         Top             =   400
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   210
         TabIndex        =   3
         Top             =   835
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   210
         TabIndex        =   5
         Top             =   1285
         Width           =   420
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "վ��"
         Height          =   210
         Left            =   2340
         TabIndex        =   7
         Top             =   1285
         Width           =   420
      End
   End
   Begin VB.Label lbl���ÿ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ÿ���"
      Height          =   210
      Left            =   90
      TabIndex        =   11
      Top             =   2363
      Width           =   840
   End
End
Attribute VB_Name = "frmClinicOfficeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As G_Enum_Fun '0-�鿴,1-���,2-����,3-ɾ��
Private mlngID As Long '��������ID
Private mrs���� As ADODB.Recordset

Private mblnOK As Boolean
Private mstrAddNewItem As String

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal lngID As Long, Optional ByRef strAddNewItem As String) As Boolean
    '�������
    '��Σ�
    '   frmParent - ������
    '   bytFun - ��������, 0-�鿴��1-������2-�޸ģ�3-ɾ��
    '���Σ�
    '   strAddNewItem:������������
    mbytFun = bytFun: mlngID = lngID
    mstrAddNewItem = ""
    
    Err = 0: On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    If mblnOK Then strAddNewItem = mstrAddNewItem
    ShowMe = mblnOK
End Function

Private Sub cboStationNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdAdd_Click()
    Call SelectDept(True)
End Sub

Private Sub SelectDept(ByVal blnButton As Boolean, Optional strLike As String)
    '����ѡ������ѡ��ʹ�ÿ���
    Dim strSql As String, rsResult As ADODB.Recordset
    Dim strID As String, str���� As String
    Dim i As Integer, vRect As RECT
    Dim blnCancel As Boolean, strIDs As String
    Dim ObjItem As ListItem
    
    Err = 0: On Error GoTo errHandler
    For i = 1 To lvwDept.ListItems.Count
        strIDs = strIDs & "," & Val(Mid(lvwDept.ListItems(i).Key, 2))
    Next
    If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    
    strSql = "Select a.ID, a.����, a.����, Upper(a.����) as ����" & vbNewLine & _
            " From ���ű� A,��������˵�� B" & vbNewLine & _
            " Where a.ID=b.����ID " & vbNewLine & _
            "       And (b.�������=1 Or b.�������=3) And b.�������� = '�ٴ�'" & vbNewLine & _
            "       And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine
    If blnButton = False Then
        'ģ������
        strSql = strSql & _
            "       And (a.���� Like [1] Or a.���� Like [1] Or Upper(a.����) Like Upper([1]))" & vbNewLine
    End If
    If strIDs <> "" Then
        '�ų���ѡ�����
        strSql = strSql & _
            "       And a.ID Not In(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))" & vbNewLine
    End If
    strSql = strSql & " Order By a.����"
    vRect = GetControlRect(txtSelect.Hwnd)
    Set rsResult = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "����", False, "", "", False, False, IIf(blnButton = False, True, False), _
        vRect.Left, vRect.Top, txtSelect.Height, blnCancel, True, False, strLike & "%", strIDs)
    If blnCancel Then Exit Sub
    If rsResult Is Nothing Then Exit Sub
    If rsResult.EOF Then Exit Sub
    
    Do While Not rsResult.EOF
        strID = Nvl(rsResult!ID): str���� = Nvl(rsResult!����)
        For i = 1 To lvwDept.ListItems.Count
            If Mid(lvwDept.ListItems(i).Key, 2) = strID Then Exit Sub
        Next
        Set ObjItem = lvwDept.ListItems.Add(, "K" & strID, str����)
        rsResult.MoveNext
    Loop
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRemove_Click()
    Err = 0: On Error GoTo errHandler
    If lvwDept.SelectedItem Is Nothing Then Exit Sub
    
    lvwDept.ListItems.Remove lvwDept.SelectedItem.Key
    If lvwDept.ListItems.Count > 0 Then
        lvwDept.ListItems(1).Selected = True
    End If
    
    If lvwDept.SelectedItem Is Nothing Then cmdRemove.Enabled = False: Exit Sub
    Call lvwDept_ItemClick(lvwDept.SelectedItem)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If Me.ActiveControl Is txt���� And txt����.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    
    Me.Caption = Choose(mbytFun + 1, "�鿴", "����", "�޸�", "ɾ��") & "��������"
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If InitData() = False Then Unload Me: Exit Sub
    End If
    If mbytFun = Fun_Add Then
        txt����.Text = GetMaxLocalCode("��������")
        Exit Sub
    End If
    
    Select Case mbytFun
    Case Fun_View
        cmdCancel.Visible = False
        cmdOk.Left = cmdCancel.Left
        Call SetEnabled(Me.Controls, False)
    Case Fun_Update
        txt����.Enabled = False
    End Select
    If LoadData(mlngID) = False Then Unload Me: Exit Sub
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim i As Long, strSql As String, rsTemp As ADODB.Recordset
    Dim intRow As Integer, intCol As Integer
    
    Err = 0: On Error GoTo errHandler
    '����վ������
    strSql = "Select ���, ���� From Zlnodelist"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    cboStationNo.Clear
    cboStationNo.AddItem ""
    Do While Not rsTemp.EOF
        cboStationNo.AddItem Nvl(rsTemp!���) & "-" & Nvl(rsTemp!����)
        If gstrNodeNo = Nvl(rsTemp!���) Then cboStationNo.ListIndex = cboStationNo.NewIndex
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData(ByVal lngID As Long) As Boolean
    '��������
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strSql = "Select a.ID, a.����, a.����, a.����, a.λ��, a.վ��, b.���" & vbNewLine & _
            " From �������� A,Zlnodelist B" & vbNewLine & _
            " Where a.վ��=b.����(+) And ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    If rsTemp.EOF Then Exit Function
    
    txt����.Text = Nvl(rsTemp!����)
    txt����.Text = Nvl(rsTemp!����)
    txt����.Text = Nvl(rsTemp!����)
    txtλ��.Text = Nvl(rsTemp!λ��)
    zlControl.CboSetText cboStationNo, Nvl(rsTemp!վ��), False
    If cboStationNo.ListIndex = -1 Then
        cboStationNo.AddItem Nvl(rsTemp!���) & "-" & Nvl(rsTemp!վ��)
        cboStationNo.ListIndex = cboStationNo.NewIndex
    End If
    
    '���ÿ���
    lvwDept.ListItems.Clear
    strSql = "Select b.Id, b.����" & vbNewLine & _
            " From �����������ÿ��� A, ���ű� B" & vbNewLine & _
            " Where a.����id = b.Id And a.����id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    If rsTemp.EOF Then LoadData = True: Exit Function
    
    Do Until rsTemp.EOF
        lvwDept.ListItems.Add , "K" & Nvl(rsTemp!ID), Nvl(rsTemp!����)
        rsTemp.MoveNext
    Loop
        
    LoadData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    
    Err = 0: On Error GoTo errHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOk.Enabled = False
    If IsValied() = False Then cmdOk.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOk.Enabled = True: Exit Sub
    
    mblnOK = True
    mstrAddNewItem = Trim(txt����.Text)
    If mbytFun = Fun_Add Then
        Call ClearFaceInfor
        cmdOk.Enabled = True
        Exit Sub
    End If
    Unload Me
    Exit Sub
errHandler:
    cmdOk.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearFaceInfor()
    '����:���������Ϣ���Ա�������������
    On Error GoTo errHandle
    txt����.Text = GetMaxLocalCode("��������")
    txt����.Text = ""
    txt����.Text = ""
    txtλ��.Text = ""
    txtSelect.Text = ""
    
    lvwDept.ListItems.Clear
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSql As String, i As Long
    Dim strTemp As String, str���ÿ��� As String
    
    Err = 0: On Error GoTo errHandler
    
    For i = 1 To lvwDept.ListItems.Count
        strTemp = Val(Mid(lvwDept.ListItems(i).Key, 2))
        str���ÿ��� = str���ÿ��� & ";" & strTemp
    Next
    If str���ÿ��� <> "" Then str���ÿ��� = Mid(str���ÿ���, 2)
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_��������_Modify(
        strSql = "Zl_��������_Modify("
        '��������_In Number,--0-������1-�޸�
        strSql = strSql & "" & 0 & ","
        'Id_In       ��������.Id%Type,
        strSql = strSql & "" & "NULL" & ","
        '����_In     ��������.����%Type := Null,
        strSql = strSql & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSql = strSql & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSql = strSql & "'" & Trim(txt����.Text) & "',"
        'λ��_In     ��������.λ��%Type := Null,
        strSql = strSql & "'" & Trim(txtλ��.Text) & "',"
        'վ��_In     ��������.վ��%Type := Null,
        strSql = strSql & "'" & NeedCode(cboStationNo.Text) & "',"
        '���ÿ���_In Varchar2:=Null--��ʽ������1;����2;����3;...
        strSql = strSql & "'" & str���ÿ��� & "')"
    Case Fun_Update
        'Zl_��������_Modify(
        strSql = "Zl_��������_Modify("
        '��������_In Number,--0-������1-�޸�
        strSql = strSql & "" & 1 & ","
        'Id_In       ��������.Id%Type,
        strSql = strSql & "" & mlngID & ","
        '����_In     ��������.����%Type := Null,
        strSql = strSql & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSql = strSql & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSql = strSql & "'" & Trim(txt����.Text) & "',"
        'λ��_In     ��������.λ��%Type := Null,
        strSql = strSql & "'" & Trim(txtλ��.Text) & "',"
        'վ��_In     ��������.վ��%Type := Null,
        strSql = strSql & "'" & NeedCode(cboStationNo.Text) & "',"
        '���ÿ���_In Varchar2:=Null--��ʽ������1;����2;����3;...
        strSql = strSql & "'" & str���ÿ��� & "')"
    End Select
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    SaveData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValied() As Boolean
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    If TxtCheckInput(txt����, "����", 3, False) = False Then Exit Function
    If TxtCheckInput(txt����, "����", 20, False) = False Then Exit Function
    If TxtCheckInput(txt����, "����", 6, False) = False Then Exit Function
    If TxtCheckInput(txtλ��, "λ��", 40, True) = False Then Exit Function
    
    If mbytFun = Fun_Add Then
        strSql = "Select 1 From �������� Where ���� = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Trim(txt����.Text))
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt����.Text) & " �Ѵ��ڣ�", vbInformation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                zlControl.TxtSelAll txt����
                Exit Function
            End If
        End If
    ElseIf mbytFun = Fun_Update Then
        strSql = "Select 1 From �������� Where ���� = [1] And ID <> [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Trim(txt����.Text), mlngID)
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt����.Text) & " �Ѵ��ڣ�", vbInformation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                zlControl.TxtSelAll txt����
                Exit Function
            End If
        End If
    End If
    IsValied = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mrs���� Is Nothing Then Set mrs���� = Nothing
End Sub

Private Sub lvwDept_GotFocus()
    cmdRemove.Enabled = Not lvwDept.SelectedItem Is Nothing
    If lvwDept.ListItems.Count = 0 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub lvwDept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwDept.SelectedItem Is Nothing Then Exit Sub
    cmdRemove.Enabled = True
End Sub

Private Sub lvwDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtSelect_GotFocus()
    zlControl.TxtSelAll txtSelect
End Sub

Private Sub txtSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtSelect.Text) = "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        Call SelectDept(False, Trim(txtSelect.Text))
        zlControl.TxtSelAll txtSelect
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_Change()
    txt����.Text = zlCommFun.SpellCode(txt����.Text)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txt����.Text) = "" Then
            MsgBox "���Ʋ���Ϊ�գ�", vbInformation, gstrSysName
            txt����.SetFocus: Exit Sub
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtλ��_GotFocus()
    zlControl.TxtSelAll txtλ��
End Sub

Private Sub txtλ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

