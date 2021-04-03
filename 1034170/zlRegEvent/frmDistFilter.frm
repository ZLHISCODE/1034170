VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDistFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分诊过滤"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4020
      TabIndex        =   18
      Top             =   2565
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
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
         ToolTipText     =   "定位F3"
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
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1485
         Width           =   2085
      End
      Begin VB.ComboBox cbo科室 
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
         Caption         =   "门诊号↓"
         Height          =   180
         Left            =   285
         TabIndex        =   21
         Top             =   1935
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   180
         Left            =   405
         TabIndex        =   9
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3420
         TabIndex        =   11
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号员"
         Height          =   180
         Left            =   3390
         TabIndex        =   15
         Top             =   1545
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   585
         TabIndex        =   13
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号时间"
         Height          =   180
         Left            =   225
         TabIndex        =   1
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   405
         TabIndex        =   5
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3420
         TabIndex        =   7
         Top             =   735
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3420
         TabIndex        =   3
         Top             =   330
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   150
      TabIndex        =   20
      Top             =   2580
      Width           =   1100
   End
   Begin VB.Menu mnuIDKind 
      Caption         =   "身份类别"
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
Option Explicit '要求变量声明
Public mlngModul As Long
Public mstrFilter As String
Public mstrSectName As String   '用来指定当前默认的科室
Private mrsDept As ADODB.Recordset  '记录临床科室
Private mrs挂号员 As ADODB.Recordset
Private mcllFiter As Variant       '条件信息
Private mblnOK As Boolean
'-----------------------------------------------------
'结算卡相关
Private mcllBrushCard As Collection
Private Type Tp_CardSquare
    bln缺省卡号密文 As Boolean
    lng缺省卡类别ID As Long
    int缺省卡号长度 As Integer
End Type
Private mTyCard As Tp_CardSquare
'-----------------------------------------------------

Public Function zlShowMe(ByVal frmMain As Form, ByVal lngModule As Long, _
    ByRef cllFilter As Variant) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：程序入口,获取相关条件设置
    '入参：frmMain-主窗体
    '         lngModule-模块号
    '出参：cllFilter-返回相关的条件信息
    '返回：
    '编制：刘兴洪
    '日期：2010-06-02 15:25:35
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    mlngModul = lngModule: Set mcllFiter = cllFilter: mblnOK = False
    Me.Show 1, frmMain
    If mblnOK Then Set cllFilter = mcllFiter
    zlShowMe = mblnOK
End Function

Private Sub InitCllData()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化集合数据
    '编制：刘兴洪
    '日期：2010-06-02 15:44:19
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If mcllFiter Is Nothing Then
        Set mcllFiter = New Collection
        mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "挂号时间"
        mcllFiter.Add Array("", ""), "挂号NO"
        mcllFiter.Add Array("", ""), "发票号"
        mcllFiter.Add "", "挂号员"
        mcllFiter.Add "", "科室"
        mcllFiter.Add "", "门诊号": mcllFiter.Add "", "就诊卡号"
        mcllFiter.Add "", "医保号": mcllFiter.Add "", "病人姓名"
        mcllFiter.Add 0, "KIND": mnuIDKinds_Click (0)
        mcllFiter.Add mstrFilter, "条件"
        Exit Sub
    End If
    '恢复默认数据
    txtNOBegin.Text = mcllFiter("挂号NO")(0):    txtNOEnd.Text = mcllFiter("挂号NO")(1)
    txtFactBegin.Text = mcllFiter("发票号")(0):    txtFactEnd.Text = mcllFiter("发票号")(1)
    dtpBegin.Value = CDate(mcllFiter("挂号时间")(0)):    dtpEnd.Value = CDate(mcllFiter("挂号时间")(1))
    mstrFilter = CStr(mcllFiter("条件"))
    Call mnuIDKinds_Click(Val(mcllFiter("KIND")))
    '集何中可能不存在,所以不加载值
    Err = 0: On Error Resume Next
    If mcllFiter(Trim(mnuIDKinds(Val(lblKind.Tag)).Tag)) <> "" Then
        '初始化
        txtValue.Text = mcllFiter("_" & Trim(mnuIDKinds(Val(lblKind.Tag)).Tag))
    End If
End Sub
Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：加载基础数据
    '编制：刘兴洪
    '日期：2010-06-02 15:59:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim str挂号员 As String, lng科室ID As Long, i  As Long, strTmp As String
    
    If mrs挂号员 Is Nothing Then
        Set mrs挂号员 = GetPersonnel("门诊挂号员", True)
    ElseIf mrs挂号员.State <> 1 Then
        Set mrs挂号员 = GetPersonnel("门诊挂号员", True)
    End If
    If Not mcllFiter Is Nothing Then
        str挂号员 = Trim(mcllFiter("挂号员"))
        lng科室ID = Val(mcllFiter("科室"))
    End If
    '挂号员
    cbo操作员.Clear
    cbo操作员.AddItem "所有挂号员"
    cbo操作员.ListIndex = 0
    If mrs挂号员.RecordCount > 0 Then
        Call mrs挂号员.MoveFirst
        For i = 1 To mrs挂号员.RecordCount
            cbo操作员.AddItem mrs挂号员!简码 & "-" & mrs挂号员!姓名
            If str挂号员 = Nvl(mrs挂号员!姓名) Then cbo操作员.ListIndex = cbo操作员.NewIndex
            mrs挂号员.MoveNext
        Next
    End If
    cbo.SetListWidthAuto cbo操作员, zlControl.OneCharWidth(cbo操作员.Font) * 70 / cbo操作员.Width
   '读取门诊临床科室，如果已经读取就不再读取
    strTmp = zlDatabase.GetPara("分诊科室", glngSys, mlngModul)
    If strTmp = "" Then strTmp = UserInfo.部门ID
    
    If mrsDept Is Nothing Then
        Set mrsDept = GetDepartments("'临床'", "1,3")
    ElseIf mrsDept.State <> 1 Then
        Set mrsDept = GetDepartments("'临床'", "1,3")
    End If
    
    cbo科室.Clear
    cbo科室.AddItem "所有科室"
    cbo科室.ListIndex = 0
    With mrsDept
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, "," & strTmp & ",", "," & !ID & ",") > 0 Then
                cbo科室.AddItem !编码 & "-" & !名称
                cbo科室.ItemData(cbo科室.NewIndex) = !ID
                If lng科室ID = Val(Nvl(!ID)) Then cbo科室.ListIndex = cbo科室.NewIndex
            End If
            .MoveNext
        Loop
    End With
    cbo.SetListWidthAuto cbo科室, zlControl.OneCharWidth(cbo科室.Font) * 70 / cbo科室.Width
    LoadData = True
End Function
Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo操作员.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo操作员.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo操作员.ListIndex = lngIdx
    If cbo操作员.ListIndex = -1 And cbo操作员.ListCount <> 0 Then cbo操作员.ListIndex = 0
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo科室.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo科室.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo科室.ListIndex = lngIdx
    If cbo科室.ListIndex = -1 And cbo科室.ListCount <> 0 Then cbo科室.ListIndex = 0
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
    '当天内
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.发生时间 Between [1] And [2]"
    Set mcllFiter = Nothing
    Call InitCllData
    Call LoadData
End Sub

Private Sub cmdOK_Click()
    If Not IsNull(dtpEnd.Value) Then
        If dtpEnd.Value < dtpBegin.Value Then
            MsgBox "结束时间不能小于开始时间！", vbInformation, gstrSysName
            dtpEnd.SetFocus: Exit Sub
        End If
    End If
    If txtNOBegin.Text <> "" And txtNOEnd.Text <> "" Then
        If txtNOEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNOEnd.SetFocus: Exit Sub
        End If
    End If
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "结束票据号不能小于开始票据号！", vbInformation, gstrSysName
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
    '问题号:30346
    If InStr(1, "《》？；：‘|｛｝【】<>?:;|'{}[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
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
    
    '当天内
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.发生时间 Between [1] And [2]"
    Call InitMenus
    Call LoadData
    Call InitCllData
End Sub
Private Sub InitMenus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:动态加载相关的医疗卡类别菜单
    '编制:刘兴洪
    '日期:2011-10-21 15:29:07
    '问题:42315
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, strKind As String
    Dim i As Long, ObjItem As Menu
    Set mcllBrushCard = New Collection
    strKind = "门|门诊号|0|0|18|0|0||"
    strKind = strKind & ";" & "姓|姓名|0|0|" & zlGetPatiInforMaxLen.intPatiName & "|0|0||"
    strKind = strKind & ";" & "就|就诊卡|0|0|18|0|0||"
    strKind = strKind & ";" & "医|医保号|0|0|20|0|0||"
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
        '取缺省的刷卡方式
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
        '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
        '第7位后,就只能用索引,不然取不到数
        mcllBrushCard.Add varTemp, varTemp(1)
        If Val(varTemp(5)) = 1 Then
            mTyCard.bln缺省卡号密文 = Trim(varTemp(7)) <> ""
            mTyCard.lng缺省卡类别ID = Val(varTemp(3))
            mTyCard.int缺省卡号长度 = Val(varTemp(4))
        End If
        If i > 9 Then
            ObjItem.Caption = varTemp(1) & IIf(i - 9 > 24, "", "(&" & Chr(64 + i) & ")")
        Else
            ObjItem.Caption = varTemp(1) & "(&" & i & ")"
        End If
        ObjItem.Tag = CStr(varTemp(1))
    Next
    '设置缺省查找对象
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
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式

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
   zlControl.TxtCheckKeyPress txtNOEnd, KeyAscii, m文本式
End Sub

Private Function MakeFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的过滤条件
    '编制:刘兴洪
    '日期:2011-10-21 15:23:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTmp As String, strSQLtmp As String
    Dim lng病人ID As Long, lng卡类别ID As Long, strErrMsg As String, strPassWord As String
    Dim strKind As String
    Dim blnCancel As Boolean
    Set mcllFiter = New Collection
    mstrFilter = " And A.发生时间 Between [1] And [2]"
    mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "挂号时间"
    mcllFiter.Add Array(Trim(txtNOBegin.Text), Trim(txtNOEnd)), "挂号NO"
    mcllFiter.Add Array(Trim(txtFactBegin.Text), Trim(txtFactEnd)), "发票号"
    If cbo操作员.ListIndex > 0 Then
        mcllFiter.Add NeedName(cbo操作员.Text), "挂号员"
    Else
        mcllFiter.Add "", "挂号员"
    End If
    mcllFiter.Add "", "科室"
    mcllFiter.Add "", "门诊号": mcllFiter.Add "", "就诊卡号"
    mcllFiter.Add "", "医保号": mcllFiter.Add "", "病人姓名"
    mcllFiter.Add Val(lblKind.Tag), "KIND"
    mcllFiter.Add "", "病人ID"
    
    strKind = mnuIDKinds(Val(lblKind.Tag)).Tag
    mcllFiter.Add Trim(txtValue.Text), "_" & strKind
    If txtNOBegin.Text <> "" And txtNOEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[3]"
    End If
    
    If cbo操作员.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.操作员姓名||''=[9]"
    If Trim(txtValue.Text) <> "" Then
        Select Case strKind
        Case "门诊号"
            mstrFilter = mstrFilter & " And A.门诊号 = [11]"
            mcllFiter.Remove "门诊号": mcllFiter.Add Trim(txtValue.Text), "门诊号"
        Case "姓名"
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtValue.Text, 1))) > 0 Then
                mstrFilter = mstrFilter & " And Upper(A.姓名) Like [8]"
            Else
                mstrFilter = mstrFilter & " And A.姓名 Like [8]"
            End If
            mcllFiter.Remove "病人姓名": mcllFiter.Add Trim(txtValue.Text), "病人姓名"
        Case "医保号"
            mstrFilter = mstrFilter & " And B.医保号=[13]"
            mcllFiter.Remove "医保号": mcllFiter.Add Trim(txtValue.Text), "医保号"
        Case Else
            '其他类别的,获取相关的病人ID
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
            lng卡类别ID = Val(mcllBrushCard(Val(lblKind.Tag) + 1)(3))
            If lng卡类别ID <> 0 Then
                If InStr("," & "身份证号,二代身份证号,二代身份证,身份证" & ",", "," & strKind & ",") > 0 Then
                     lng病人ID = GetPatiID(mlngModul, Me, Trim(txtValue.Text), txtValue, , , blnCancel)
                End If
                If lng病人ID = 0 And Not blnCancel Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, Trim(txtValue.Text), True, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(strKind, Trim(txtValue.Text), True, lng病人ID, _
                    strPassWord, strErrMsg) = False Then lng病人ID = 0
            End If
            If lng病人ID = 0 Then
                If strErrMsg = "" Then
                    MsgBox "未找到满足条件的病人", vbInformation + vbOKOnly, gstrSysName
                    If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
                    zlControl.TxtSelAll txtValue
                    Exit Function
                End If
            End If
            mstrFilter = mstrFilter & " And A.病人ID=[12]"
            mcllFiter.Remove "病人ID": mcllFiter.Add lng病人ID, "病人ID"
        End Select
    End If
    
    strSQL = ""
    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
        '无需根据票据号判断,直接根据单据的发生时间判断
        strSQLtmp = IIf(txtFactEnd.Text = "", " =[5] ", " Between [5] And [6] ")
        strSQL = "Select A.NO" & _
        " From 票据打印内容 A,票据使用明细 B" & _
        " Where A.数据性质=4 And A.ID=B.打印ID And B.性质=1" & _
        " And B.号码 " & strSQLtmp
    End If
    If strSQL <> "" Then mstrFilter = mstrFilter & " And A.NO IN(" & strSQL & ")"
    '挂号科室(执行科室)
    If cbo科室.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And A.执行部门ID+0=[7]"
        mcllFiter.Remove "科室"
        mcllFiter.Add cbo科室.ItemData(cbo科室.ListIndex), "科室"
    End If
    mcllFiter.Add mstrFilter, "条件"
    MakeFilter = True
End Function

Private Sub mnuIDKinds_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuIDKinds.UBound
        mnuIDKinds(i).Checked = i = Index
    Next
    lblKind.Caption = mnuIDKinds(Index).Tag & "↓"
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
    Dim strKind As String, intKind As Integer, int卡号长度 As Long
    Dim bln密文 As Boolean
    
    strKind = mnuIDKinds(Val(lblKind.Tag)).Tag
    intKind = Val(lblKind.Tag) + 1
    bln密文 = mcllBrushCard(intKind)(7) <> ""
    txtValue.PasswordChar = IIf(bln密文, "*", "")
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
           blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, mTyCard.bln缺省卡号密文)
           int卡号长度 = mTyCard.int缺省卡号长度 - 1
    Case "门诊号"
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            int卡号长度 = 0
    Case "医保号"
            int卡号长度 = 0
    Case Else
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, bln密文)
        int卡号长度 = mcllBrushCard(intKind)(4)
    End Select
    If int卡号长度 > 0 Then
         '刷卡完毕或输入号码后回车
         If blnCard And Len(txtValue.Text) = int卡号长度 - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtValue.Text) <> "" Then
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



