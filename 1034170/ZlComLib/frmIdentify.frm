VERSION 5.00
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������֤"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "frmIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdReadIC 
      Caption         =   "����"
      Height          =   405
      Left            =   4500
      TabIndex        =   10
      Top             =   1230
      Width           =   585
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   2
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   3
      Top             =   2865
      Width           =   1100
   End
   Begin VB.Frame fraDown 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   2685
      Width           =   6900
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   5670
      TabIndex        =   4
      Top             =   0
      Width           =   5670
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ˣ����������У�30��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   255
         TabIndex        =   9
         Top             =   105
         Width           =   2640
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4845
         Picture         =   "frmIdentify.frx":058A
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʣ���1000.00�����ν�1000.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   465
         Width           =   4320
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   6000
         Y1              =   810
         Y2              =   810
      End
   End
   Begin VB.TextBox txtPass 
      Enabled         =   0   'False
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
      IMEMode         =   3  'DISABLE
      Left            =   1455
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1815
      Width           =   3015
   End
   Begin VB.TextBox txtCard 
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
      IMEMode         =   3  'DISABLE
      Left            =   1455
      TabIndex        =   0
      Top             =   1230
      Width           =   3015
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   420
      Left            =   810
      TabIndex        =   11
      Top             =   1245
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   741
      Appearance      =   2
      IDKindStr       =   "��|���￨|0|0|0|0|0|;IC|IC����|1|0|0|0|0|"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "����"
      IDKind          =   -1
      ShowPropertySet =   -1  'True
      NotContainFastKey=   ""
      BackColor       =   -2147483633
      SaveRegType     =   4
      ProductName     =   "һ��ͨ����֧��"
   End
   Begin VB.Label lblCardNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   555
      TabIndex        =   7
      Top             =   1890
      Width           =   870
   End
End
Attribute VB_Name = "frmIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintCount As Integer
Private mlng����ID As String
Private mlngSys As Long
Private mblnPreCard As Boolean
Private mobjCard As Card '��ǰ�����Ŀ�
'--------------------------------------------------
'�����:
Private mobjKeyboard As Object
Private mblnPassInputCardNo As Boolean  '�Ƿ��������뿨��
Private mobjSquareCard As Object
Private mlngҽ�ƿ����� As Long
Private mlngModul As Long
Private mstrPassWord As String
Private mlngDefaultCardTypeID As Long 'ȱʡ��ˢ�����ID
Private mblnBrushCard As Boolean
Private Const VK_RETURN = &HD
Private mblnCheckPassWord As Boolean
Private mblnReadIDCard As Boolean  '��ȡ��������֤
Private mblnReadICCard As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard  '����:47945
Attribute mobjICCard.VB_VarHelpID = -1
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mstrRegSection As String
Private mlngPreBrushCardTypeID As Long '�ϴ�ˢ�����
'--------------------------------------------------
Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng����ID As Long, _
    ByVal cur��� As Currency, Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0, _
    Optional lngDefaultCardTypeID As Long = 0, _
    Optional blnCheckPassWord As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�������
    '���:frmParent-���õ�������
    '       lngSys-ϵͳ��
    '       lng����ID-ָ���Ĳ���ID
    '       lngModul-ģ���
    '       bytOperationType-ҵ������(0-������;1-����;2-סԺ)
    '       mlngDefaultCardTypeID-ȱʡ��ˢ�����ID
    '       blnCheckPassWord-��֤����(true-��֤����,false-ֻˢ��,����������)
    '����:
    '����:��֤�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-10 16:35:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strSQL As String, intMouse As Integer
    mblnCheckPassWord = blnCheckPassWord
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngDefaultCardTypeID
    mblnOK = False: mintCount = 3: mlng����ID = lng����ID
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0

    '��ȡ���￨��Ϣ
    On Error GoTo ErrH
    strSQL = "" & _
    "   Select A.����,A.�Ա�,A.����,A.���￨��,A.����֤��, " & _
    "              nvl(B.���,0) as ���" & _
    "   From ������Ϣ A," & _
    "       (   Select ����ID,nvl(Sum(Ԥ�����),0)-nvl(sum(�������),0) as ��� " & _
    "           From  ������� " & _
    "           Where ����ID=[1] and ����=1 and decode([2],0,0,����)=[2]  Group by ����ID) B " & _
    "   Where A.����ID=[1] And A.����ID=B.����ID(+) "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����ID, bytOperationType)
    
    If rsTmp.EOF Then
        MsgBox "������Ϣ������,����!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If IIf(IsNull(rsTmp!���￨��), "", rsTmp!���￨�� & "") = "" Then
        '����:43449
        strSQL = "Select Count(Distinct �����ID) as ����� From ����ҽ�ƿ���Ϣ Where  ����ID=[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����ID, bytOperationType)
        If IIf(IsNull(rsTemp!�����), 0, Val(rsTemp!����� & "")) = 0 Then
            '--δ����,ֱ�ӷ���true,�����鿨
            ShowMe = True: Exit Function
        End If
    End If
    Me.lblPati.Caption = "���ˣ�" & gobjComLib.zlCommFun.NVL(rsTmp!����) & _
        IIf(Not IsNull(rsTmp!�Ա�), "��" & rsTmp!�Ա�, "") & _
        IIf(Not IsNull(rsTmp!����), "��" & rsTmp!����, "")
    Me.lblMoney.Caption = "ʣ���" & Format(rsTmp!���, "0.00") & "�����ν�" & Format(cur���, "0.00")
    Me.txtCard.Tag = gobjComLib.zlCommFun.NVL(rsTmp!���￨��)
    Me.txtPass.Tag = gobjComLib.zlCommFun.NVL(rsTmp!����֤��)
    On Error GoTo 0
    Me.Show 1, frmParent
    ShowMe = mblnOK
    
    Screen.MousePointer = intMouse
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function zlCheckICCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���IC���������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-28 18:04:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnSucces As Boolean, strPassWord As String
    
    If Not mblnReadICCard Then zlCheckICCard = True: Exit Function
    On Error GoTo errHandle
    If UCase(txtCard.Text) <> UCase(txtCard.Tag) Then
        MsgBox "��ǰIC�����벡�˵�IC���Ų�����", vbExclamation, gstrSysName
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Function '���Ų�ƥ�䣬��׼����
    End If
    If Val(lblPass.Tag) <> mlng����ID Or Val(lblPass.Tag) = 0 Then
        MsgBox "��ǰIC�����벡�˵�IC���Ų������", vbExclamation, gstrSysName
        Exit Function '���Ų�ƥ�䣬��׼����
    End If
    If Not mblnCheckPassWord Then zlCheckICCard = True: Exit Function
    
    strSQL = "Select ���� as IC����,���� From ����ҽ�ƿ���Ϣ Where ����ID=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTemp.EOF Then
        strSQL = "Select  IC����,����֤�� as ���� From ������Ϣ Where ����ID=[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ����֤�Ҳ���ָ���Ĳ���,��ȷ�ϸò����Ƿ���", vbExclamation, gstrSysName
            If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
    End If
    '�ȼ���Ƿ��иÿ�,���Ըÿ������������
    strPassWord = gobjComLib.zlCommFun.zlStringEncode(txtPass.Text)
    rsTemp.Filter = "IC����='" & UCase(txtCard.Text) & "'"
    If Not rsTemp.EOF Then
        If strPassWord = gobjComLib.zlCommFun.NVL(rsTemp!����) Then zlCheckICCard = True: Exit Function
        rsTemp.Filter = 0
    End If
    
    'ֻҪ��һ����������,������
    With rsTemp
        blnSucces = False
        Do While Not .EOF
            If strPassWord = gobjComLib.zlCommFun.NVL(rsTemp!����) Then blnSucces = True: Exit Do
            .MoveNext
        Loop
    End With
    If blnSucces Then zlCheckICCard = True: Exit Function
    
    If mintCount = 1 Then
        MsgBox "���������������,���������룡", vbExclamation, gstrSysName
    Else
        MsgBox "�����������", vbExclamation, gstrSysName
    End If
    txtPass.Text = "": mintCount = mintCount - 1
    If mintCount = 0 Then Unload Me: Exit Function   '������󣬿�����2��
    If txtPass.Enabled Then txtPass.SetFocus
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ������Ч��
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 17:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPassWord As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim blnSucces As Boolean '����ɹ�
    Dim str���� As String
    On Error GoTo errHandle
    
    If mobjCard Is Nothing Then Exit Function
    If mobjCard.���� Like "*����" Then
        str���� = mobjCard.����
    ElseIf mobjCard.���� Like "*����֤" Then
        str���� = "����֤��"
    ElseIf mobjCard.���� Like "*��" Then
        str���� = mobjCard.���� & "����"
    Else
        str���� = mobjCard.���� & "������"
    End If

    If UCase(Trim(txtCard.Text)) = "" Then Exit Function
    If Val(lblPass.Tag) <> mlng����ID Or Val(lblPass.Tag) = 0 Then
        MsgBox "��ǰ" & str���� & "�벡�˵�" & str���� & "�������", vbExclamation, gstrSysName
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Function
    End If
    
    If Not mblnCheckPassWord Then IsValied = True: Exit Function
    strPassWord = gobjComLib.zlCommFun.zlStringEncode(txtPass.Text)
    If strPassWord <> mstrPassWord Then
        If mintCount = 1 Then
            MsgBox "���������������,���������룡", vbExclamation, gstrSysName
        Else
            MsgBox "�����������", vbExclamation, gstrSysName
        End If
        txtPass.Text = "": mintCount = mintCount - 1
        If mintCount = 0 Then
            Unload Me '������󣬿�����2��
        ElseIf txtPass.Enabled Then
            txtPass.SetFocus
        End If
        Exit Function
    End If
    IsValied = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub cmdOK_Click()
    If IsValied = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub
Private Sub cmdReadIC_Click()
        Call IDKind_Click(IDKind.GetCurCard)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IDKind.ActiveFastKey
End Sub
Private Sub Form_Load()
    
    HookDefend txtPass.hwnd
    mstrRegSection = "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name & Me.Name
    mlngPreBrushCardTypeID = GetSetting("ZLSOFT", mstrRegSection, "ȱʡ�����ID", 0)

    Call CreateObjectKeyboard
    Call zlCardSquareObject
    Call SetCtrlVisible
    Call NewCardObject
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not IDKind.GetCurCard Is Nothing Then
         SaveSetting "ZLSOFT", mstrRegSection, "ȱʡ�����ID", IDKind.GetCurCard.�ӿ����
    End If
    
    Set mobjKeyboard = Nothing
    Set mobjCard = Nothing
    Call zlCardSquareObject(True)
    Call CloseIDCard
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard Is Nothing Then Exit Sub
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        
        If mobjICCard Is Nothing Then Exit Sub
        txtCard.MaxLength = 0
        txtCard.Text = mobjICCard.Read_Card()
        If txtCard.Text = "" Then
            If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
            Exit Sub
        End If
        
            '�����:42948
        If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            txtCard.Text = "": If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
        End If
        mblnReadICCard = True
        If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
        If txtCard.Text <> "" Then
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Exit Sub
        End If
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus: Exit Sub
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If mobjSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtCard.Text = strOutCardNO
    
    '�����:42948
    If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
    End If
    
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
     If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
     End If
     If mblnCheckPassWord Then txtPass.SetFocus: Exit Sub
     Call cmdOK_Click
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    txtCard.PasswordChar = IIf(objCard.�������Ĺ��� <> "", "*", "")
    mblnBrushCard = objCard.�Ƿ�ˢ��
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not mblnBrushCard
    cmdReadIC.Visible = Not mblnBrushCard
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)

    txtCard.Text = objPatiInfor.����
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
    End If
    
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
     If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
     End If
     If mblnCheckPassWord Then txtPass.SetFocus: Exit Sub
     Call cmdOK_Click
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    'IC����ȡ
    
    If strCardNO = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtCard.MaxLength = Len(strCardNO)
    txtCard.Text = strCardNO: mblnReadICCard = True
    If GetPatient(objCard, strCardNO) = False Then
         mblnReadICCard = False: Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    '��ʾ����Ϣ
    If strID = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("����֤��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtCard.Text = strID: mblnReadICCard = True
    txtCard.MaxLength = Len(strID)
    If GetPatient(objCard, strID) = False Then
         mblnReadICCard = False: Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub
Private Sub txtCard_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    IDKind.SetAutoReadCard (False)
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub
Private Sub txtCard_Change()
    lblPass.Tag = "": txtCard.Tag = ""
    txtPass.Enabled = txtCard.Text <> ""
    If Not txtPass.Enabled Then txtPass.Text = ""
    mblnReadIDCard = False
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtCard.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtCard.Text = "")
    IDKind.SetAutoReadCard (txtCard.Text = "")
End Sub

Private Sub txtCard_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtCard)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtCard.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtCard.Text = "")
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    mblnPreCard = False

    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = IDKind.GetCurCard.���ų��� - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        If GetPatient(IDKind.GetCurCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
        End If
        mblnPreCard = blnCard
        If mblnCheckPassWord Then
            If txtPass.Enabled Then txtPass.SetFocus
        Else
            Call cmdOK_Click: Exit Sub
        End If
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 And IDKind.GetCurCard.�Ƿ�ֿ����� = True Then
            sngNow = timer
            If txtCard.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txtCard.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txtCard.Text = Chr(KeyAscii)
                txtCard.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub
Private Sub txtCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtCard.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtCard.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtCard.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    glngTXTProc = GetWindowLong(txtPass.hwnd, GWL_WNDPROC)
    Call SetWindowLong(txtPass.hwnd, GWL_WNDPROC, AddressOf WndMessage)
End Sub

Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button <> 2 Then Exit Sub
    Call SetWindowLong(txtPass.hwnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPass_GotFocus()
    If txtCard.Text <> "" And mstrPassWord = "" Then Call cmdOK_Click: Exit Sub
    Call gobjComLib.zlControl.TxtSelAll(txtPass)
    OpenPassKeyboard txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mblnPreCard Then
            '60580
            mblnPreCard = False
             If (GetAsyncKeyState(VK_RETURN) And &H1) <> 0 Then
                txtPass.Text = ""
                Exit Sub
             End If
        End If
        mblnPreCard = False
        Call cmdOK_Click
    ElseIf KeyAscii = 22 Then
        KeyAscii = 0 '������ճ��
    Else
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        End If
    End If
    '60580
    mblnPreCard = False
End Sub

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
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
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function
 
 
Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String, i As Integer, intIdKind As Integer
    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
    If blnClosed Then
       If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.CloseWindows
            Set mobjSquareCard = Nothing
        End If
        Exit Sub
    End If
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitCompoent (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Call mobjSquareCard.zlInitComponents(Me, mlngModul, mlngSys, gstrDBUser, gcnOracle, False, strExpend)
    mobjSquareCard.mblnYLMgr = True
    Err = 0: On Error GoTo 0
    Call IDKind.zlInit(Me, mlngSys, mlngModul, gcnOracle, gstrDBUser, mobjSquareCard, "", txtCard)
    
    Err = 0: On Error Resume Next
     If mlngPreBrushCardTypeID <> 0 Then
        intIdKind = IDKind.GetKindIndex(mlngPreBrushCardTypeID)
        If intIdKind <> 0 Then
            IDKind.IDKind = intIdKind
        End If
     End If
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    Optional blnIDCard As Boolean = False, Optional blnICCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:objCard-��ָ���Ŀ������ж���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo ErrH
    
    mstrPassWord = ""
    Set mobjCard = Nothing
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Function
    '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
    If mobjSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID, Nothing, Me, False, True) = False Then
        '����ģ������:-1:ҽ�ƿ����(���������ǰ�Ŀ��ų��Ȳ����Ļ�,���������)
        If mobjSquareCard.zlGetPatiID(-1, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID, Nothing, Me, False, True) = False Then
            GoTo NotFoundPati:
        End If
    End If
    If lng����ID <= 0 Then GoTo NotFoundPati:
    If mlng����ID <> lng����ID Then
        If objCard.���� Like "*����" Then
            MsgBox "��ǰ" & objCard.���� & "�벡�������е�" & objCard.���� & "�����,���飡", vbExclamation, gstrSysName
        ElseIf objCard.���� Like "*����֤" Then
            MsgBox "��ǰ����֤���벡�������е�����֤�Ų����,���飡", vbExclamation, gstrSysName
        ElseIf objCard.���� Like "*��" Then
            MsgBox "��ǰ" & objCard.���� & "�����벡�������е�" & objCard.���� & "���Ų����,���飡", vbExclamation, gstrSysName
        Else
            MsgBox "��ǰ" & objCard.���� & "�������벡�������е�" & objCard.���� & "�����Ų����,���飡", vbExclamation, gstrSysName
        End If
        txtCard.Text = ""
        Exit Function '���Ų�ƥ�䣬��׼����
    End If
    txtCard.Tag = strInput
    lblPass.Tag = lng����ID
    mstrPassWord = strPassWord
    Set mobjCard = objCard
    GetPatient = True
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
    Else
        MsgBox "δ�ҵ���ǰ���ĳ��в���,����!", vbOKOnly + vbInformation, gstrSysName
        txtCard.Text = ""
    End If
    txtCard.Tag = "": lblPass.Tag = ""
End Function
Private Function IsDesinMode() As Boolean
      '���˺� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���visible����
    '����:���˺�
    '����:2012-03-13 11:28:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    lblPass.Visible = mblnCheckPassWord
    txtPass.Visible = mblnCheckPassWord
    If mblnCheckPassWord Then Exit Sub
    With txtCard
        .Top = picTop.Top + picTop.Height + (fraDown.Top - (picTop.Top + picTop.Height) - .Height) \ 2
        IDKind.Top = .Top
        cmdReadIC.Top = .Top
        lblCardNO.Top = .Top + (.Height - lblCardNO.Height) \ 2
    End With
    If Err <> 0 Then Err.Clear
End Sub

Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ر�������������
    '����:���˺�
    '����:2012-03-09 16:26:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���µĿ�����
    '����:���˺�
    '����:2012-03-09 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
    End If
    If mobjICCard Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Err = 0: On Error GoTo 0
    End If
End Sub


