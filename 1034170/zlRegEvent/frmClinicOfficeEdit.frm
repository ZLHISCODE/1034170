VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicOfficeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊诊室设置"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "宋体"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdRemove 
      Caption         =   "移除(&D)"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4260
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2298
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "…"
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
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "适用科室"
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
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   5460
      TabIndex        =   17
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   5460
      TabIndex        =   18
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   360
      Left            =   5460
      TabIndex        =   19
      Top             =   4290
      Width           =   1100
   End
   Begin VB.Frame fra基本信息 
      Caption         =   "基本信息"
      Height          =   2055
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5085
      Begin VB.TextBox txt编码 
         Height          =   350
         Left            =   660
         MaxLength       =   3
         TabIndex        =   2
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txt位置 
         Height          =   350
         Left            =   660
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1650
         Width           =   4335
      End
      Begin VB.TextBox txt名称 
         Height          =   350
         Left            =   660
         MaxLength       =   20
         TabIndex        =   4
         Top             =   765
         Width           =   4335
      End
      Begin VB.TextBox txt简码 
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
      Begin VB.Label lbl位置 
         AutoSize        =   -1  'True
         Caption         =   "位置"
         Height          =   210
         Left            =   210
         TabIndex        =   9
         Top             =   1720
         Width           =   420
      End
      Begin VB.Label lbl编码 
         AutoSize        =   -1  'True
         Caption         =   "编码"
         Height          =   210
         Left            =   210
         TabIndex        =   1
         Top             =   400
         Width           =   420
      End
      Begin VB.Label lbl名称 
         AutoSize        =   -1  'True
         Caption         =   "名称"
         Height          =   210
         Left            =   210
         TabIndex        =   3
         Top             =   835
         Width           =   420
      End
      Begin VB.Label lbl简码 
         AutoSize        =   -1  'True
         Caption         =   "简码"
         Height          =   210
         Left            =   210
         TabIndex        =   5
         Top             =   1285
         Width           =   420
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "站点"
         Height          =   210
         Left            =   2340
         TabIndex        =   7
         Top             =   1285
         Width           =   420
      End
   End
   Begin VB.Label lbl适用科室 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "适用科室"
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
Private mbytFun As G_Enum_Fun '0-查看,1-添加,2-调整,3-删除
Private mlngID As Long '门诊诊室ID
Private mrs科室 As ADODB.Recordset

Private mblnOK As Boolean
Private mstrAddNewItem As String

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal lngID As Long, Optional ByRef strAddNewItem As String) As Boolean
    '程序入口
    '入参：
    '   frmParent - 父窗口
    '   bytFun - 操作类型, 0-查看，1-新增，2-修改，3-删除
    '出参：
    '   strAddNewItem:新增诊室名称
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
    '弹出选择器，选择使用科室
    Dim strSql As String, rsResult As ADODB.Recordset
    Dim strID As String, str名称 As String
    Dim i As Integer, vRect As RECT
    Dim blnCancel As Boolean, strIDs As String
    Dim ObjItem As ListItem
    
    Err = 0: On Error GoTo errHandler
    For i = 1 To lvwDept.ListItems.Count
        strIDs = strIDs & "," & Val(Mid(lvwDept.ListItems(i).Key, 2))
    Next
    If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    
    strSql = "Select a.ID, a.编码, a.名称, Upper(a.简码) as 简码" & vbNewLine & _
            " From 部门表 A,部门性质说明 B" & vbNewLine & _
            " Where a.ID=b.部门ID " & vbNewLine & _
            "       And (b.服务对象=1 Or b.服务对象=3) And b.工作性质 = '临床'" & vbNewLine & _
            "       And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine
    If blnButton = False Then
        '模糊查找
        strSql = strSql & _
            "       And (a.编码 Like [1] Or a.名称 Like [1] Or Upper(a.简码) Like Upper([1]))" & vbNewLine
    End If
    If strIDs <> "" Then
        '排除已选择科室
        strSql = strSql & _
            "       And a.ID Not In(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))" & vbNewLine
    End If
    strSql = strSql & " Order By a.名称"
    vRect = GetControlRect(txtSelect.Hwnd)
    Set rsResult = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "科室", False, "", "", False, False, IIf(blnButton = False, True, False), _
        vRect.Left, vRect.Top, txtSelect.Height, blnCancel, True, False, strLike & "%", strIDs)
    If blnCancel Then Exit Sub
    If rsResult Is Nothing Then Exit Sub
    If rsResult.EOF Then Exit Sub
    
    Do While Not rsResult.EOF
        strID = Nvl(rsResult!ID): str名称 = Nvl(rsResult!名称)
        For i = 1 To lvwDept.ListItems.Count
            If Mid(lvwDept.ListItems(i).Key, 2) = strID Then Exit Sub
        Next
        Set ObjItem = lvwDept.ListItems.Add(, "K" & strID, str名称)
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
    If Me.ActiveControl Is txt编码 And txt编码.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    
    Me.Caption = Choose(mbytFun + 1, "查看", "新增", "修改", "删除") & "门诊诊室"
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If InitData() = False Then Unload Me: Exit Sub
    End If
    If mbytFun = Fun_Add Then
        txt编码.Text = GetMaxLocalCode("门诊诊室")
        Exit Sub
    End If
    
    Select Case mbytFun
    Case Fun_View
        cmdCancel.Visible = False
        cmdOk.Left = cmdCancel.Left
        Call SetEnabled(Me.Controls, False)
    Case Fun_Update
        txt编码.Enabled = False
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
    '加载站点数据
    strSql = "Select 编号, 名称 From Zlnodelist"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    cboStationNo.Clear
    cboStationNo.AddItem ""
    Do While Not rsTemp.EOF
        cboStationNo.AddItem Nvl(rsTemp!编号) & "-" & Nvl(rsTemp!名称)
        If gstrNodeNo = Nvl(rsTemp!编号) Then cboStationNo.ListIndex = cboStationNo.NewIndex
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
    '加载数据
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strSql = "Select a.ID, a.编码, a.名称, a.简码, a.位置, a.站点, b.编号" & vbNewLine & _
            " From 门诊诊室 A,Zlnodelist B" & vbNewLine & _
            " Where a.站点=b.名称(+) And ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    If rsTemp.EOF Then Exit Function
    
    txt编码.Text = Nvl(rsTemp!编码)
    txt名称.Text = Nvl(rsTemp!名称)
    txt简码.Text = Nvl(rsTemp!简码)
    txt位置.Text = Nvl(rsTemp!位置)
    zlControl.CboSetText cboStationNo, Nvl(rsTemp!站点), False
    If cboStationNo.ListIndex = -1 Then
        cboStationNo.AddItem Nvl(rsTemp!编号) & "-" & Nvl(rsTemp!站点)
        cboStationNo.ListIndex = cboStationNo.NewIndex
    End If
    
    '适用科室
    lvwDept.ListItems.Clear
    strSql = "Select b.Id, b.名称" & vbNewLine & _
            " From 门诊诊室适用科室 A, 部门表 B" & vbNewLine & _
            " Where a.科室id = b.Id And a.诊室id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    If rsTemp.EOF Then LoadData = True: Exit Function
    
    Do Until rsTemp.EOF
        lvwDept.ListItems.Add , "K" & Nvl(rsTemp!ID), Nvl(rsTemp!名称)
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
    mstrAddNewItem = Trim(txt名称.Text)
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
    '功能:清除界面信息，以便重新输入数据
    On Error GoTo errHandle
    txt编码.Text = GetMaxLocalCode("门诊诊室")
    txt名称.Text = ""
    txt简码.Text = ""
    txt位置.Text = ""
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
    Dim strTemp As String, str适用科室 As String
    
    Err = 0: On Error GoTo errHandler
    
    For i = 1 To lvwDept.ListItems.Count
        strTemp = Val(Mid(lvwDept.ListItems(i).Key, 2))
        str适用科室 = str适用科室 & ";" & strTemp
    Next
    If str适用科室 <> "" Then str适用科室 = Mid(str适用科室, 2)
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_门诊诊室_Modify(
        strSql = "Zl_门诊诊室_Modify("
        '操作类型_In Number,--0-新增，1-修改
        strSql = strSql & "" & 0 & ","
        'Id_In       门诊诊室.Id%Type,
        strSql = strSql & "" & "NULL" & ","
        '编码_In     门诊诊室.编码%Type := Null,
        strSql = strSql & "'" & Trim(txt编码.Text) & "',"
        '名称_In     门诊诊室.名称%Type := Null,
        strSql = strSql & "'" & Trim(txt名称.Text) & "',"
        '简码_In     门诊诊室.简码%Type := Null,
        strSql = strSql & "'" & Trim(txt简码.Text) & "',"
        '位置_In     门诊诊室.位置%Type := Null,
        strSql = strSql & "'" & Trim(txt位置.Text) & "',"
        '站点_In     门诊诊室.站点%Type := Null,
        strSql = strSql & "'" & NeedCode(cboStationNo.Text) & "',"
        '适用科室_In Varchar2:=Null--格式：科室1;科室2;科室3;...
        strSql = strSql & "'" & str适用科室 & "')"
    Case Fun_Update
        'Zl_门诊诊室_Modify(
        strSql = "Zl_门诊诊室_Modify("
        '操作类型_In Number,--0-新增，1-修改
        strSql = strSql & "" & 1 & ","
        'Id_In       门诊诊室.Id%Type,
        strSql = strSql & "" & mlngID & ","
        '编码_In     门诊诊室.编码%Type := Null,
        strSql = strSql & "'" & Trim(txt编码.Text) & "',"
        '名称_In     门诊诊室.名称%Type := Null,
        strSql = strSql & "'" & Trim(txt名称.Text) & "',"
        '简码_In     门诊诊室.简码%Type := Null,
        strSql = strSql & "'" & Trim(txt简码.Text) & "',"
        '位置_In     门诊诊室.位置%Type := Null,
        strSql = strSql & "'" & Trim(txt位置.Text) & "',"
        '站点_In     门诊诊室.站点%Type := Null,
        strSql = strSql & "'" & NeedCode(cboStationNo.Text) & "',"
        '适用科室_In Varchar2:=Null--格式：科室1;科室2;科室3;...
        strSql = strSql & "'" & str适用科室 & "')"
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
    If TxtCheckInput(txt编码, "编码", 3, False) = False Then Exit Function
    If TxtCheckInput(txt名称, "名称", 20, False) = False Then Exit Function
    If TxtCheckInput(txt简码, "简码", 6, False) = False Then Exit Function
    If TxtCheckInput(txt位置, "位置", 40, True) = False Then Exit Function
    
    If mbytFun = Fun_Add Then
        strSql = "Select 1 From 门诊诊室 Where 名称 = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Trim(txt名称.Text))
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt名称.Text) & " 已存在！", vbInformation, gstrSysName
                If txt名称.Visible And txt名称.Enabled Then txt名称.SetFocus
                zlControl.TxtSelAll txt名称
                Exit Function
            End If
        End If
    ElseIf mbytFun = Fun_Update Then
        strSql = "Select 1 From 门诊诊室 Where 名称 = [1] And ID <> [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Trim(txt名称.Text), mlngID)
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt名称.Text) & " 已存在！", vbInformation, gstrSysName
                If txt名称.Visible And txt名称.Enabled Then txt名称.SetFocus
                zlControl.TxtSelAll txt名称
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
    If Not mrs科室 Is Nothing Then Set mrs科室 = Nothing
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

Private Sub txt编码_GotFocus()
    zlControl.TxtSelAll txt编码
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt简码_GotFocus()
    zlControl.TxtSelAll txt简码
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt名称_Change()
    txt简码.Text = zlCommFun.SpellCode(txt名称.Text)
End Sub

Private Sub txt名称_GotFocus()
    zlControl.TxtSelAll txt名称
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txt名称.Text) = "" Then
            MsgBox "名称不能为空！", vbInformation, gstrSysName
            txt名称.SetFocus: Exit Sub
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt位置_GotFocus()
    zlControl.TxtSelAll txt位置
End Sub

Private Sub txt位置_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

