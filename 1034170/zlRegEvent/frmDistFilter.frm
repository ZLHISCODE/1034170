VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDistFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4020
      TabIndex        =   18
      Top             =   2565
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5205
      TabIndex        =   19
      Top             =   2565
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      Begin VB.TextBox txtValue 
         Height          =   300
         Left            =   1005
         TabIndex        =   17
         ToolTipText     =   "��λF3"
         Top             =   1890
         Width           =   2085
      End
      Begin VB.TextBox txtFactEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         TabIndex        =   12
         Top             =   1094
         Width           =   2085
      End
      Begin VB.TextBox txtFactBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         TabIndex        =   10
         Top             =   1094
         Width           =   2085
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   8
         TabIndex        =   8
         Top             =   682
         Width           =   2085
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   6
         Top             =   682
         Width           =   2085
      End
      Begin VB.ComboBox cbo����Ա 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1485
         Width           =   2085
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1506
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3975
         TabIndex        =   4
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   63766531
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1005
         TabIndex        =   2
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   63766531
         CurrentDate     =   36588
      End
      Begin VB.Label lblKind 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����š�"
         Height          =   180
         Left            =   285
         TabIndex        =   21
         Top             =   1935
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
         Height          =   180
         Left            =   405
         TabIndex        =   9
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3420
         TabIndex        =   11
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label lbl����Ա 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һ�Ա"
         Height          =   180
         Left            =   3390
         TabIndex        =   15
         Top             =   1545
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   585
         TabIndex        =   13
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һ�ʱ��"
         Height          =   180
         Left            =   225
         TabIndex        =   1
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   405
         TabIndex        =   5
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3420
         TabIndex        =   7
         Top             =   735
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3420
         TabIndex        =   3
         Top             =   330
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   150
      TabIndex        =   20
      Top             =   2580
      Width           =   1100
   End
   Begin VB.Menu mnuIDKind 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu mnuIDKinds 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDistFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlngModul As Long
Public mstrFilter As String
Public mstrSectName As String   '����ָ����ǰĬ�ϵĿ���
Private mrsDept As ADODB.Recordset  '��¼�ٴ�����
Private mrs�Һ�Ա As ADODB.Recordset
Private mcllFiter As Variant       '������Ϣ
Private mblnOK As Boolean
'-----------------------------------------------------
'���㿨���
Private mcllBrushCard As Collection
Private Type Tp_CardSquare
    blnȱʡ�������� As Boolean
    lngȱʡ�����ID As Long
    intȱʡ���ų��� As Integer
End Type
Private mTyCard As Tp_CardSquare
'-----------------------------------------------------

Public Function zlShowMe(ByVal frmMain As Form, ByVal lngModule As Long, _
    ByRef cllFilter As Variant) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��������,��ȡ�����������
    '��Σ�frmMain-������
    '         lngModule-ģ���
    '���Σ�cllFilter-������ص�������Ϣ
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-02 15:25:35
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    mlngModul = lngModule: Set mcllFiter = cllFilter: mblnOK = False
    Me.Show 1, frmMain
    If mblnOK Then Set cllFilter = mcllFiter
    zlShowMe = mblnOK
End Function

Private Sub InitCllData()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ����������
    '���ƣ����˺�
    '���ڣ�2010-06-02 15:44:19
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If mcllFiter Is Nothing Then
        Set mcllFiter = New Collection
        mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "�Һ�ʱ��"
        mcllFiter.Add Array("", ""), "�Һ�NO"
        mcllFiter.Add Array("", ""), "��Ʊ��"
        mcllFiter.Add "", "�Һ�Ա"
        mcllFiter.Add "", "����"
        mcllFiter.Add "", "�����": mcllFiter.Add "", "���￨��"
        mcllFiter.Add "", "ҽ����": mcllFiter.Add "", "��������"
        mcllFiter.Add 0, "KIND": mnuIDKinds_Click (0)
        mcllFiter.Add mstrFilter, "����"
        Exit Sub
    End If
    '�ָ�Ĭ������
    txtNOBegin.Text = mcllFiter("�Һ�NO")(0):    txtNOEnd.Text = mcllFiter("�Һ�NO")(1)
    txtFactBegin.Text = mcllFiter("��Ʊ��")(0):    txtFactEnd.Text = mcllFiter("��Ʊ��")(1)
    dtpBegin.Value = CDate(mcllFiter("�Һ�ʱ��")(0)):    dtpEnd.Value = CDate(mcllFiter("�Һ�ʱ��")(1))
    mstrFilter = CStr(mcllFiter("����"))
    Call mnuIDKinds_Click(Val(mcllFiter("KIND")))
    '�����п��ܲ�����,���Բ�����ֵ
    Err = 0: On Error Resume Next
    If mcllFiter(Trim(mnuIDKinds(Val(lblKind.Tag)).Tag)) <> "" Then
        '��ʼ��
        txtValue.Text = mcllFiter("_" & Trim(mnuIDKinds(Val(lblKind.Tag)).Tag))
    End If
End Sub
Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ػ�������
    '���ƣ����˺�
    '���ڣ�2010-06-02 15:59:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim str�Һ�Ա As String, lng����ID As Long, i  As Long, strTmp As String
    
    If mrs�Һ�Ա Is Nothing Then
        Set mrs�Һ�Ա = GetPersonnel("����Һ�Ա", True)
    ElseIf mrs�Һ�Ա.State <> 1 Then
        Set mrs�Һ�Ա = GetPersonnel("����Һ�Ա", True)
    End If
    If Not mcllFiter Is Nothing Then
        str�Һ�Ա = Trim(mcllFiter("�Һ�Ա"))
        lng����ID = Val(mcllFiter("����"))
    End If
    '�Һ�Ա
    cbo����Ա.Clear
    cbo����Ա.AddItem "���йҺ�Ա"
    cbo����Ա.ListIndex = 0
    If mrs�Һ�Ա.RecordCount > 0 Then
        Call mrs�Һ�Ա.MoveFirst
        For i = 1 To mrs�Һ�Ա.RecordCount
            cbo����Ա.AddItem mrs�Һ�Ա!���� & "-" & mrs�Һ�Ա!����
            If str�Һ�Ա = Nvl(mrs�Һ�Ա!����) Then cbo����Ա.ListIndex = cbo����Ա.NewIndex
            mrs�Һ�Ա.MoveNext
        Next
    End If
    cbo.SetListWidthAuto cbo����Ա, zlControl.OneCharWidth(cbo����Ա.Font) * 70 / cbo����Ա.Width
   '��ȡ�����ٴ����ң�����Ѿ���ȡ�Ͳ��ٶ�ȡ
    strTmp = zlDatabase.GetPara("�������", glngSys, mlngModul)
    If strTmp = "" Then strTmp = UserInfo.����ID
    
    If mrsDept Is Nothing Then
        Set mrsDept = GetDepartments("'�ٴ�'", "1,3")
    ElseIf mrsDept.State <> 1 Then
        Set mrsDept = GetDepartments("'�ٴ�'", "1,3")
    End If
    
    cbo����.Clear
    cbo����.AddItem "���п���"
    cbo����.ListIndex = 0
    With mrsDept
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, "," & strTmp & ",", "," & !ID & ",") > 0 Then
                cbo����.AddItem !���� & "-" & !����
                cbo����.ItemData(cbo����.NewIndex) = !ID
                If lng����ID = Val(Nvl(!ID)) Then cbo����.ListIndex = cbo����.NewIndex
            End If
            .MoveNext
        Loop
    End With
    cbo.SetListWidthAuto cbo����, zlControl.OneCharWidth(cbo����.Font) * 70 / cbo����.Width
    LoadData = True
End Function
Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����Ա.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����Ա.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����Ա.ListIndex = lngIdx
    If cbo����Ա.ListIndex = -1 And cbo����Ա.ListCount <> 0 Then cbo����Ա.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount <> 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdDef_Click()
    Dim Curdate As Date
    txtNOBegin.Text = ""
    txtNOEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txtValue.Text = ""
    '������
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.����ʱ�� Between [1] And [2]"
    Set mcllFiter = Nothing
    Call InitCllData
    Call LoadData
End Sub

Private Sub cmdOK_Click()
    If Not IsNull(dtpEnd.Value) Then
        If dtpEnd.Value < dtpBegin.Value Then
            MsgBox "����ʱ�䲻��С�ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            dtpEnd.SetFocus: Exit Sub
        End If
    End If
    If txtNOBegin.Text <> "" And txtNOEnd.Text <> "" Then
        If txtNOEnd.Text < txtNOBegin.Text Then
            MsgBox "�������ݺŲ���С�ڿ�ʼ���ݺţ�", vbInformation, gstrSysName
            txtNOEnd.SetFocus: Exit Sub
        End If
    End If
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "����Ʊ�ݺŲ���С�ڿ�ʼƱ�ݺţ�", vbInformation, gstrSysName
            txtFactEnd.SetFocus: Exit Sub
        End If
    End If
    If MakeFilter = False Then Exit Sub
    mblnOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
    If KeyCode = vbKeyF3 Then Call txtValue.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����:30346
    If InStr(1, "������������|��������<>?:;|'{}[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Public Sub Form_Load()
    Dim Curdate As Date, i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    txtNOBegin.Text = ""
    txtNOEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txtValue.Text = ""
    
    '������
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.����ʱ�� Between [1] And [2]"
    Call InitMenus
    Call LoadData
    Call InitCllData
End Sub
Private Sub InitMenus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��̬������ص�ҽ�ƿ����˵�
    '����:���˺�
    '����:2011-10-21 15:29:07
    '����:42315
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, strKind As String
    Dim i As Long, ObjItem As Menu
    Set mcllBrushCard = New Collection
    strKind = "��|�����|0|0|18|0|0||"
    strKind = strKind & ";" & "��|����|0|0|" & zlGetPatiInforMaxLen.intPatiName & "|0|0||"
    strKind = strKind & ";" & "��|���￨|0|0|18|0|0||"
    strKind = strKind & ";" & "ҽ|ҽ����|0|0|20|0|0||"
    If Not gobjSquare.objSquareCard Is Nothing Then
        strKind = gobjSquare.objSquareCard.zlGetIDKindStr(strKind)
    End If
        
    varData = Split(strKind, ";")
    For i = 0 To UBound(varData)
        Set ObjItem = Me.mnuIDKinds(mnuIDKinds.UBound)
        If Not (ObjItem.Caption = "-" Or Trim(ObjItem.Caption) = "" Or Not ObjItem.Visible) Then
            Load mnuIDKinds(mnuIDKinds.UBound + 1)
            Set ObjItem = mnuIDKinds(mnuIDKinds.UBound)
        End If
        varTemp = Split(varData(i), "|")
        'ȡȱʡ��ˢ����ʽ
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
        '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
        '��7λ��,��ֻ��������,��Ȼȡ������
        mcllBrushCard.Add varTemp, varTemp(1)
        If Val(varTemp(5)) = 1 Then
            mTyCard.blnȱʡ�������� = Trim(varTemp(7)) <> ""
            mTyCard.lngȱʡ�����ID = Val(varTemp(3))
            mTyCard.intȱʡ���ų��� = Val(varTemp(4))
        End If
        If i > 9 Then
            ObjItem.Caption = varTemp(1) & IIf(i - 9 > 24, "", "(&" & Chr(64 + i) & ")")
        Else
            ObjItem.Caption = varTemp(1) & "(&" & i & ")"
        End If
        ObjItem.Tag = CStr(varTemp(1))
    Next
    '����ȱʡ���Ҷ���
    mnuIDKinds_Click (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Set mrsDept = Nothing
End Sub

Private Sub lblKind_Click()
    PopupMenu mnuIDKind, 2
End Sub

Private Sub txtFactBegin_GotFocus()
    SelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    SelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub

Private Sub txtNOBegin_Change()
    txtNOEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNOEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    SelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m�ı�ʽ

End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 12)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNOEnd.Text <> "" Then txtNOEnd.Text = GetFullNO(txtNOEnd.Text, 12)
End Sub

Private Sub txtNoEnd_GotFocus()
    SelAll txtNOEnd
End Sub


Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
   zlControl.TxtCheckKeyPress txtNOEnd, KeyAscii, m�ı�ʽ
End Sub

Private Function MakeFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ĺ�������
    '����:���˺�
    '����:2011-10-21 15:23:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTmp As String, strSQLtmp As String
    Dim lng����ID As Long, lng�����ID As Long, strErrMsg As String, strPassWord As String
    Dim strKind As String
    Dim blnCancel As Boolean
    Set mcllFiter = New Collection
    mstrFilter = " And A.����ʱ�� Between [1] And [2]"
    mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "�Һ�ʱ��"
    mcllFiter.Add Array(Trim(txtNOBegin.Text), Trim(txtNOEnd)), "�Һ�NO"
    mcllFiter.Add Array(Trim(txtFactBegin.Text), Trim(txtFactEnd)), "��Ʊ��"
    If cbo����Ա.ListIndex > 0 Then
        mcllFiter.Add NeedName(cbo����Ա.Text), "�Һ�Ա"
    Else
        mcllFiter.Add "", "�Һ�Ա"
    End If
    mcllFiter.Add "", "����"
    mcllFiter.Add "", "�����": mcllFiter.Add "", "���￨��"
    mcllFiter.Add "", "ҽ����": mcllFiter.Add "", "��������"
    mcllFiter.Add Val(lblKind.Tag), "KIND"
    mcllFiter.Add "", "����ID"
    
    strKind = mnuIDKinds(Val(lblKind.Tag)).Tag
    mcllFiter.Add Trim(txtValue.Text), "_" & strKind
    If txtNOBegin.Text <> "" And txtNOEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[3]"
    End If
    
    If cbo����Ա.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.����Ա����||''=[9]"
    If Trim(txtValue.Text) <> "" Then
        Select Case strKind
        Case "�����"
            mstrFilter = mstrFilter & " And A.����� = [11]"
            mcllFiter.Remove "�����": mcllFiter.Add Trim(txtValue.Text), "�����"
        Case "����"
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtValue.Text, 1))) > 0 Then
                mstrFilter = mstrFilter & " And Upper(A.����) Like [8]"
            Else
                mstrFilter = mstrFilter & " And A.���� Like [8]"
            End If
            mcllFiter.Remove "��������": mcllFiter.Add Trim(txtValue.Text), "��������"
        Case "ҽ����"
            mstrFilter = mstrFilter & " And B.ҽ����=[13]"
            mcllFiter.Remove "ҽ����": mcllFiter.Add Trim(txtValue.Text), "ҽ����"
        Case Else
            '��������,��ȡ��صĲ���ID
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
            lng�����ID = Val(mcllBrushCard(Val(lblKind.Tag) + 1)(3))
            If lng�����ID <> 0 Then
                If InStr("," & "���֤��,�������֤��,�������֤,���֤" & ",", "," & strKind & ",") > 0 Then
                     lng����ID = GetPatiID(mlngModul, Me, Trim(txtValue.Text), txtValue, , , blnCancel)
                End If
                If lng����ID = 0 And Not blnCancel Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, Trim(txtValue.Text), True, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(strKind, Trim(txtValue.Text), True, lng����ID, _
                    strPassWord, strErrMsg) = False Then lng����ID = 0
            End If
            If lng����ID = 0 Then
                If strErrMsg = "" Then
                    MsgBox "δ�ҵ����������Ĳ���", vbInformation + vbOKOnly, gstrSysName
                    If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
                    zlControl.TxtSelAll txtValue
                    Exit Function
                End If
            End If
            mstrFilter = mstrFilter & " And A.����ID=[12]"
            mcllFiter.Remove "����ID": mcllFiter.Add lng����ID, "����ID"
        End Select
    End If
    
    strSQL = ""
    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
        '�������Ʊ�ݺ��ж�,ֱ�Ӹ��ݵ��ݵķ���ʱ���ж�
        strSQLtmp = IIf(txtFactEnd.Text = "", " =[5] ", " Between [5] And [6] ")
        strSQL = "Select A.NO" & _
        " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
        " Where A.��������=4 And A.ID=B.��ӡID And B.����=1" & _
        " And B.���� " & strSQLtmp
    End If
    If strSQL <> "" Then mstrFilter = mstrFilter & " And A.NO IN(" & strSQL & ")"
    '�Һſ���(ִ�п���)
    If cbo����.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And A.ִ�в���ID+0=[7]"
        mcllFiter.Remove "����"
        mcllFiter.Add cbo����.ItemData(cbo����.ListIndex), "����"
    End If
    mcllFiter.Add mstrFilter, "����"
    MakeFilter = True
End Function

Private Sub mnuIDKinds_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuIDKinds.UBound
        mnuIDKinds(i).Checked = i = Index
    Next
    lblKind.Caption = mnuIDKinds(Index).Tag & "��"
    lblKind.Tag = Index
    lblKind.ToolTipText = mnuIDKinds(Index).Tag
    txtValue.ToolTipText = mnuIDKinds(Index).Tag
End Sub

Private Sub txtValue_GotFocus()
    If mnuIDKinds(1).Checked Then zlCommFun.OpenIme True
    
    SelAll txtValue
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = vbKeyF4 Then
        For i = 0 To mnuIDKinds.Count - 1
            If mnuIDKinds(i).Checked = True Then Exit For
        Next
        If i >= mnuIDKinds.Count - 1 Then
            i = 0
        Else
            i = i + 1
        End If
        Call mnuIDKinds_Click(i)
    End If
End Sub

Private Sub txtvalue_LostFocus()
    zlCommFun.OpenIme
End Sub
Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    Dim strKind As String, intKind As Integer, int���ų��� As Long
    Dim bln���� As Boolean
    
    strKind = mnuIDKinds(Val(lblKind.Tag)).Tag
    intKind = Val(lblKind.Tag) + 1
    bln���� = mcllBrushCard(intKind)(7) <> ""
    txtValue.PasswordChar = IIf(bln����, "*", "")
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
           blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, mTyCard.blnȱʡ��������)
           int���ų��� = mTyCard.intȱʡ���ų��� - 1
    Case "�����"
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            int���ų��� = 0
    Case "ҽ����"
            int���ų��� = 0
    Case Else
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, bln����)
        int���ų��� = mcllBrushCard(intKind)(4)
    End Select
    If int���ų��� > 0 Then
         'ˢ����ϻ���������س�
         If blnCard And Len(txtValue.Text) = int���ų��� - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtValue.Text) <> "" Then
             If KeyAscii <> 13 Then
                 txtValue.Text = txtValue.Text & Chr(KeyAscii)
                 txtValue.SelStart = Len(txtValue.Text)
             End If
             KeyAscii = 0
              If MakeFilter Then Unload Me: Exit Sub
              zlControl.TxtSelAll txtValue
        End If
    End If
End Sub



