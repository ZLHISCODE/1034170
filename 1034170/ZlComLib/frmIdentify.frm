VERSION 5.00
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人身份验证"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdReadIC 
      Caption         =   "读卡"
      Height          =   405
      Left            =   4500
      TabIndex        =   10
      Top             =   1230
      Width           =   585
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   2
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
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
         Caption         =   "病人：张永康，男，30岁"
         BeginProperty Font 
            Name            =   "宋体"
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
         Caption         =   "剩余款额：1000.00，本次金额：1000.00"
         BeginProperty Font 
            Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      IDKindStr       =   "就|就诊卡|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "宋体"
      IDKind          =   -1
      ShowPropertySet =   -1  'True
      NotContainFastKey=   ""
      BackColor       =   -2147483633
      SaveRegType     =   4
      ProductName     =   "一卡通消费支付"
   End
   Begin VB.Label lblCardNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
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
Private mlng病人ID As String
Private mlngSys As Long
Private mblnPreCard As Boolean
Private mobjCard As Card '当前处理的卡
'--------------------------------------------------
'卡相关:
Private mobjKeyboard As Object
Private mblnPassInputCardNo As Boolean  '是否密文输入卡号
Private mobjSquareCard As Object
Private mlng医疗卡长度 As Long
Private mlngModul As Long
Private mstrPassWord As String
Private mlngDefaultCardTypeID As Long '缺省的刷卡类别ID
Private mblnBrushCard As Boolean
Private Const VK_RETURN = &HD
Private mblnCheckPassWord As Boolean
Private mblnReadIDCard As Boolean  '读取的是身份证
Private mblnReadICCard As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard  '问题:47945
Attribute mobjICCard.VB_VarHelpID = -1
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mstrRegSection As String
Private mlngPreBrushCardTypeID As Long '上次刷卡类别
'--------------------------------------------------
Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng病人ID As Long, _
    ByVal cur金额 As Currency, Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0, _
    Optional lngDefaultCardTypeID As Long = 0, _
    Optional blnCheckPassWord As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证窗体入口
    '入参:frmParent-调用的主窗体
    '       lngSys-系统号
    '       lng病人ID-指定的病人ID
    '       lngModul-模块号
    '       bytOperationType-业务类型(0-不区分;1-门诊;2-住院)
    '       mlngDefaultCardTypeID-缺省的刷卡类别ID
    '       blnCheckPassWord-验证密码(true-验证密码,false-只刷卡,不输入密码)
    '出参:
    '返回:验证成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-10 16:35:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strSQL As String, intMouse As Integer
    mblnCheckPassWord = blnCheckPassWord
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngDefaultCardTypeID
    mblnOK = False: mintCount = 3: mlng病人ID = lng病人ID
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0

    '读取就诊卡信息
    On Error GoTo ErrH
    strSQL = "" & _
    "   Select A.姓名,A.性别,A.年龄,A.就诊卡号,A.卡验证码, " & _
    "              nvl(B.余额,0) as 余额" & _
    "   From 病人信息 A," & _
    "       (   Select 病人ID,nvl(Sum(预交余额),0)-nvl(sum(费用余额),0) as 余额 " & _
    "           From  病人余额 " & _
    "           Where 病人ID=[1] and 性质=1 and decode([2],0,0,类型)=[2]  Group by 病人ID) B " & _
    "   Where A.病人ID=[1] And A.病人ID=B.病人ID(+) "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng病人ID, bytOperationType)
    
    If rsTmp.EOF Then
        MsgBox "病人信息不存在,请检查!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If IIf(IsNull(rsTmp!就诊卡号), "", rsTmp!就诊卡号 & "") = "" Then
        '问题:43449
        strSQL = "Select Count(Distinct 卡类别ID) as 类别数 From 病人医疗卡信息 Where  病人ID=[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng病人ID, bytOperationType)
        If IIf(IsNull(rsTemp!类别数), 0, Val(rsTemp!类别数 & "")) = 0 Then
            '--未发卡,直接返回true,不用验卡
            ShowMe = True: Exit Function
        End If
    End If
    Me.lblPati.Caption = "病人：" & gobjComLib.zlCommFun.NVL(rsTmp!姓名) & _
        IIf(Not IsNull(rsTmp!性别), "，" & rsTmp!性别, "") & _
        IIf(Not IsNull(rsTmp!年龄), "，" & rsTmp!年龄, "")
    Me.lblMoney.Caption = "剩余款额：" & Format(rsTmp!余额, "0.00") & "，本次金额：" & Format(cur金额, "0.00")
    Me.txtCard.Tag = gobjComLib.zlCommFun.NVL(rsTmp!就诊卡号)
    Me.txtPass.Tag = gobjComLib.zlCommFun.NVL(rsTmp!卡验证码)
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
    '功能:检查IC卡的数据是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-28 18:04:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnSucces As Boolean, strPassWord As String
    
    If Not mblnReadICCard Then zlCheckICCard = True: Exit Function
    On Error GoTo errHandle
    If UCase(txtCard.Text) <> UCase(txtCard.Tag) Then
        MsgBox "当前IC卡号与病人的IC卡号不符！", vbExclamation, gstrSysName
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Function '卡号不匹配，不准重试
    End If
    If Val(lblPass.Tag) <> mlng病人ID Or Val(lblPass.Tag) = 0 Then
        MsgBox "当前IC卡号与病人的IC卡号不相符！", vbExclamation, gstrSysName
        Exit Function '卡号不匹配，不准重试
    End If
    If Not mblnCheckPassWord Then zlCheckICCard = True: Exit Function
    
    strSQL = "Select 卡号 as IC卡号,密码 From 病人医疗卡信息 Where 病人ID=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If rsTemp.EOF Then
        strSQL = "Select  IC卡号,卡验证码 as 密码 From 病人信息 Where 病人ID=[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        If rsTemp.EOF Then
            MsgBox "当前身份证找不到指定的病人,请确认该病人是否就诊！", vbExclamation, gstrSysName
            If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
    End If
    '先检查是否有该卡,就以该卡的密码就行了
    strPassWord = gobjComLib.zlCommFun.zlStringEncode(txtPass.Text)
    rsTemp.Filter = "IC卡号='" & UCase(txtCard.Text) & "'"
    If Not rsTemp.EOF Then
        If strPassWord = gobjComLib.zlCommFun.NVL(rsTemp!密码) Then zlCheckICCard = True: Exit Function
        rsTemp.Filter = 0
    End If
    
    '只要有一个密码区配,就行了
    With rsTemp
        blnSucces = False
        Do While Not .EOF
            If strPassWord = gobjComLib.zlCommFun.NVL(rsTemp!密码) Then blnSucces = True: Exit Do
            .MoveNext
        Loop
    End With
    If blnSucces Then zlCheckICCard = True: Exit Function
    
    If mintCount = 1 Then
        MsgBox "三次密码输入错误,不能再输入！", vbExclamation, gstrSysName
    Else
        MsgBox "密码输入错误！", vbExclamation, gstrSysName
    End If
    txtPass.Text = "": mintCount = mintCount - 1
    If mintCount = 0 Then Unload Me: Exit Function   '密码错误，可输入2次
    If txtPass.Enabled Then txtPass.SetFocus
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡的有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 17:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPassWord As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim blnSucces As Boolean '输入成功
    Dim str名称 As String
    On Error GoTo errHandle
    
    If mobjCard Is Nothing Then Exit Function
    If mobjCard.名称 Like "*卡号" Then
        str名称 = mobjCard.名称
    ElseIf mobjCard.名称 Like "*身份证" Then
        str名称 = "身份证号"
    ElseIf mobjCard.名称 Like "*卡" Then
        str名称 = mobjCard.名称 & "卡号"
    Else
        str名称 = mobjCard.名称 & "卡卡号"
    End If

    If UCase(Trim(txtCard.Text)) = "" Then Exit Function
    If Val(lblPass.Tag) <> mlng病人ID Or Val(lblPass.Tag) = 0 Then
        MsgBox "当前" & str名称 & "与病人的" & str名称 & "不相符！", vbExclamation, gstrSysName
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Function
    End If
    
    If Not mblnCheckPassWord Then IsValied = True: Exit Function
    strPassWord = gobjComLib.zlCommFun.zlStringEncode(txtPass.Text)
    If strPassWord <> mstrPassWord Then
        If mintCount = 1 Then
            MsgBox "三次密码输入错误,不能再输入！", vbExclamation, gstrSysName
        Else
            MsgBox "密码输入错误！", vbExclamation, gstrSysName
        End If
        txtPass.Text = "": mintCount = mintCount - 1
        If mintCount = 0 Then
            Unload Me '密码错误，可输入2次
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
    mstrRegSection = "私有模块\" & gstrDBUser & "\界面设置\" & Me.Name & Me.Name
    mlngPreBrushCardTypeID = GetSetting("ZLSOFT", mstrRegSection, "缺省卡类别ID", 0)

    Call CreateObjectKeyboard
    Call zlCardSquareObject
    Call SetCtrlVisible
    Call NewCardObject
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not IDKind.GetCurCard Is Nothing Then
         SaveSetting "ZLSOFT", mstrRegSection, "缺省卡类别ID", IDKind.GetCurCard.接口序号
    End If
    
    Set mobjKeyboard = Nothing
    Set mobjCard = Nothing
    Call zlCardSquareObject(True)
    Call CloseIDCard
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard Is Nothing Then Exit Sub
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
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
        
            '问题号:42948
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
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If mobjSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtCard.Text = strOutCardNO
    
    '问题号:42948
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
    txtCard.PasswordChar = IIf(objCard.卡号密文规则 <> "", "*", "")
    mblnBrushCard = objCard.是否刷卡
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not mblnBrushCard
    cmdReadIC.Visible = Not mblnBrushCard
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)

    txtCard.Text = objPatiInfor.卡号
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
    'IC卡读取
    
    If strCardNO = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
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
    '显示卡信息
    If strID = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证号", CardTypeName)
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

    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = IDKind.GetCurCard.卡号长度 - 1 And txtCard.SelLength <> Len(txtCard.Text)
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
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 And IDKind.GetCurCard.是否持卡消费 = True Then
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
        KeyAscii = 0 '不允许粘贴
    Else
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        End If
    End If
    '60580
    mblnPreCard = False
End Sub

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
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
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
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
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
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
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String, i As Integer, intIdKind As Integer
    '只有:执行或退费时,才可能管结算卡的
    If blnClosed Then
       If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.CloseWindows
            Set mobjSquareCard = Nothing
        End If
        Exit Sub
    End If
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitCompoent (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
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
    '功能:读取病人信息
    '入参:objCard-按指定的卡类别进行读卡
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo ErrH
    
    mstrPassWord = ""
    Set mobjCard = Nothing
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Function
    '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
    If mobjSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID, Nothing, Me, False, True) = False Then
        '进行模糊查找:-1:医疗卡类别(但是如果当前的卡号长度不够的话,会存在问题)
        If mobjSquareCard.zlGetPatiID(-1, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID, Nothing, Me, False, True) = False Then
            GoTo NotFoundPati:
        End If
    End If
    If lng病人ID <= 0 Then GoTo NotFoundPati:
    If mlng病人ID <> lng病人ID Then
        If objCard.名称 Like "*卡号" Then
            MsgBox "当前" & objCard.名称 & "与病人所持有的" & objCard.名称 & "不相符,请检查！", vbExclamation, gstrSysName
        ElseIf objCard.名称 Like "*身份证" Then
            MsgBox "当前身份证号与病人所持有的身份证号不相符,请检查！", vbExclamation, gstrSysName
        ElseIf objCard.名称 Like "*卡" Then
            MsgBox "当前" & objCard.名称 & "卡号与病人所持有的" & objCard.名称 & "卡号不相符,请检查！", vbExclamation, gstrSysName
        Else
            MsgBox "当前" & objCard.名称 & "卡卡号与病人所持有的" & objCard.名称 & "卡卡号不相符,请检查！", vbExclamation, gstrSysName
        End If
        txtCard.Text = ""
        Exit Function '卡号不匹配，不准重试
    End If
    txtCard.Tag = strInput
    lblPass.Tag = lng病人ID
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
        MsgBox "未找到当前卡的持有病人,请检查!", vbOKOnly + vbInformation, gstrSysName
        txtCard.Text = ""
    End If
    txtCard.Tag = "": lblPass.Tag = ""
End Function
Private Function IsDesinMode() As Boolean
      '刘兴洪 确定当前模式为设计模式
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
    '功能:设置控件的visible属性
    '编制:刘兴洪
    '日期:2012-03-13 11:28:12
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
    '功能:关闭自助读卡功能
    '编制:刘兴洪
    '日期:2012-03-09 16:26:40
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
    '功能:初始化新的卡对象
    '编制:刘兴洪
    '日期:2012-03-09 16:28:23
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



