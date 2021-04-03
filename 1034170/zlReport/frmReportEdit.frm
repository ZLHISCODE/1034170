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
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   4
      Top             =   810
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
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
      Begin VB.TextBox txt说明 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   735
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1125
         Width           =   3000
      End
      Begin VB.TextBox txt名称 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   735
         MaxLength       =   40
         TabIndex        =   1
         Top             =   705
         Width           =   3000
      End
      Begin VB.TextBox txt编号 
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
         Caption         =   "说明"
         Height          =   180
         Left            =   285
         TabIndex        =   8
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   180
         Left            =   285
         TabIndex        =   7
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编号"
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
Private mstr名称 As String
Private mstrOld名称 As String
Private mstr编码 As String
Private mstr说明 As String
Private mstrOld说明 As String
Private mblnOK As Boolean
Private mlngModule As Long

Public Function ShowMe(ByVal frmParent As Object, ByVal lngSys As Long, ByVal blnGroupEdit As Boolean, ByVal lngModule As Long, Optional ByRef LngGroupID As Long, _
                                        Optional ByRef lngReortID As Long, Optional ByRef str名称 As String, Optional ByRef str编码 As String, Optional ByRef str说明 As String) As Boolean
    mblnGroupEdit = blnGroupEdit
    mlngSys = lngSys
    mlngModule = lngModule
    mlngReortID = lngReortID
    mlngGroupID = LngGroupID
    mstr名称 = str名称: mstrOld名称 = str名称
    mstr编码 = str编码
    mstr说明 = str说明: mstrOld说明 = str说明
    Me.Show 1, frmParent
    str名称 = mstr名称
    str编码 = mstr编码
    str说明 = mstr说明
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsCheck As New ADODB.Recordset
    Dim strSQL As String
    Dim strOldName As String, strOld说明 As String
    Dim intOrder As Integer
    Dim arrSQL() As Variant
    Dim i As Long, blnTrans As Boolean
    
    arrSQL = Array()
    If Not CheckFormInput(Me) Then Exit Sub
    
    If Trim(txt编号.Text) = "" Then
        MsgBox "请输入报表" & IIF(mblnGroupEdit, "组", "") & "的编号！", vbInformation, App.Title
        txt编号.SetFocus: Exit Sub
    End If
    If Trim(txt名称.Text) = "" Then
        MsgBox "请输入报表" & IIF(mblnGroupEdit, "组", "") & "的名称！", vbInformation, App.Title
        txt名称.SetFocus: Exit Sub
    Else
        txt名称.Text = ConvertSBC(txt名称.Text)
    End If
    
    If Not CheckLen(txt编号, 20, "编号") Then Exit Sub
    If Not CheckLen(txt名称, 30, "名称") Then Exit Sub
    If Not CheckLen(txt说明, 255, "说明") Then Exit Sub
    
    '编号不能重复(报表及报表组)
    If CheckExist("zlReports", "编号", txt编号.Text, mlngReortID) Then
        MsgBox "该编号已经被使用,请重新输入！", vbInformation, App.Title
        txt编号.SetFocus: Exit Sub
    End If
    If CheckExist("zlRPTGroups", "编号", txt编号.Text, mlngGroupID) Then
        MsgBox "该编号已经被使用,请重新输入！", vbInformation, App.Title
        txt编号.SetFocus: Exit Sub
    End If
    If mlngGroupID <> 0 And Not mblnGroupEdit Then
        strSQL = "Select 1 From zlRPTSubs A,zlReports B Where B.名称=[1] And A.报表ID=B.ID And A.组ID=[2]" & IIF(mlngReortID = 0, "", " And 报表ID<>[3]")
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, txt名称.Text, mlngGroupID, mlngReortID)
        If Not rsCheck.EOF Then
            MsgBox "该报表组中已经包含相同名称的报表！", vbInformation, App.Title
            txt名称.SetFocus: Exit Sub
        End If
    End If
    strOldName = mstrOld名称: strOld说明 = mstrOld说明
    mstr名称 = txt名称.Text: mstr编码 = txt编号.Text: mstr说明 = txt说明.Text
    On Error GoTo errH
    If mblnGroupEdit Then
        If mlngGroupID = 0 Then
            mlngGroupID = GetNextID("zlRPTGroups")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Insert Into zlRPTGroups(ID,编号,名称,说明) Values(" & mlngGroupID & ",'" & mstr编码 & "','" & mstr名称 & "','" & mstr说明 & "')"
        ElseIf Not (strOldName = mstr名称 And strOld说明 = mstr说明) Then '说明与名称发生变化
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Update zlRPTGroups Set 编号='" & mstr编码 & "',名称='" & mstr名称 & "',说明='" & mstr说明 & "' Where ID=" & mlngGroupID
            '发布到导航台菜单的报表标题
            If mlngModule <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlPrograms Set 标题='" & mstr名称 & "',说明='" & mstr说明 & "' Where 序号=" & mlngModule & " And Nvl(系统,0)=" & mlngSys
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlMenus Set 标题='" & mstr名称 & "',短标题='" & mstr名称 & "',说明='" & mstr说明 & "' Where ID=" & mlngModule & " And Nvl(系统,0)=" & mlngSys
            End If
        End If
    Else
        If mlngReortID = 0 Then
            If mlngSys <> 0 Then mlngSys = 0
            mlngReortID = GetNextID("zlReports")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Insert Into zlReports(ID,编号,名称,说明,系统,修改时间,密码) Values(" & _
                                                        mlngReortID & ",'" & mstr编码 & "','" & mstr名称 & "','" & mstr说明 & "'," & IIF(mlngSys = 0, "NULL", mlngSys) & ",Sysdate," & AdjustStr(GetPass(mstr编码, mstr名称)) & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(" & _
                                                        mlngReortID & ",1,'" & mstr名称 & "1'," & INIT_WIDTH & "," & INIT_HEIGHT & ",9,1,0,0)"

            If mlngGroupID <> 0 Then
                intOrder = 1
                strSQL = "Select Count(*) Records From zlRPTSubs Where 组ID=[1]"
                Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
                If Not rsCheck.EOF Then intOrder = Nvl(rsCheck!Records, 0) + 1
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Insert Into zlRPTSubs(组ID,报表ID,序号,功能) " & _
                                         "Values(" & mlngGroupID & "," & mlngReortID & "," & intOrder & ",'" & mstr名称 & "')"
                If mlngModule <> 0 Then '插入权限记录
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Insert Into zlProgFuncs(系统,序号,功能,说明) Values(" & _
                                                                IIF(mlngSys = 0, "NULL", mlngSys) & "," & mlngModule & ",'" & mstr名称 & "','" & mstr说明 & "')"
                End If
            End If
        ElseIf Not (strOldName = mstr名称 And strOld说明 = mstr说明) Then '说明与名称发生变化
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Update zlReports Set 编号='" & mstr编码 & "',名称='" & _
                                     mstr名称 & "',说明='" & mstr说明 & "',密码=" & AdjustStr(GetPass(mstr编码, mstr名称)) & " Where ID=" & mlngReortID
            If mlngModule <> 0 Then '发布到导航台菜单的报表标题
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlPrograms Set 标题='" & mstr名称 & "',说明='" & mstr说明 & "'" & _
                                        " Where Upper(部件)=Upper('zl9Report') And 序号=" & mlngModule & " And Nvl(系统,0)=" & mlngSys
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Update zlMenus  Set 标题='" & mstr名称 & "',短标题='" & mstr名称 & "',说明='" & mstr说明 & "'" & _
                                        " Where 模块=" & mlngModule & " And Nvl(系统,0)=" & mlngSys & _
                                        " And Exists(Select 标题 From zlPrograms Where Upper(部件)=Upper('zl9Report') And 序号=" & mlngModule & " And Nvl(系统,0)=" & mlngSys & ")"
            
            End If
            '发布到导航台的报表组子表的功能名
            strSQL = "Select Distinct Nvl(B.系统, 0) 系统, B.程序id 序号, a.组Id " & vbNewLine & _
                     "From Zlrptsubs a, Zlrptgroups b, Zlprograms c" & vbNewLine & _
                     "Where A.组id = B.Id And A.报表id = [1]  And Nvl(B.系统, 0) = Nvl(C.系统, 0) And B.程序id = C.序号 And Upper(C.部件) = Upper('zl9Report')"
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReortID)
            Do While Not rsCheck.EOF
                If strOldName <> mstr名称 Then  '报表名称发生变化
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新子报表名称
                        arrSQL(UBound(arrSQL)) = _
                            "Update zlRPTSubs " & vbNewLine & _
                            "Set 功能 = '" & mstr名称 & "' " & vbNewLine & _
                            "Where 组Id = " & Nvl(rsCheck!组Id) & _
                            "    And 报表Id = " & mlngReortID & " And 功能 = '" & strOldName & "'"
                            
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能信息
                    arrSQL(UBound(arrSQL)) = "Insert Into Zlprogfuncs" & vbNewLine & _
                                            "  (系统, 序号, 功能, 排列, 说明, 缺省值)" & vbNewLine & _
                                            "  Select A.系统, A.序号, '" & mstr名称 & "', A.排列, '" & mstr说明 & "', A.缺省值" & vbNewLine & _
                                            "  From Zlprogfuncs a" & vbNewLine & _
                                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & strOldName & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能授权信息
                    arrSQL(UBound(arrSQL)) = "Insert Into zlrolegrant" & vbNewLine & _
                                            "  (系统,序号,角色,功能)" & vbNewLine & _
                                            "  Select A.系统,A.序号,A.角色, '" & mstr名称 & "' " & vbNewLine & _
                                            "  From zlrolegrant a" & vbNewLine & _
                                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & strOldName & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能对象权限信息
                    arrSQL(UBound(arrSQL)) = "Insert Into zlprogprivs" & vbNewLine & _
                                            "  (系统,序号,功能,对象,所有者,权限)" & vbNewLine & _
                                            "  Select A.系统,A.序号,'" & mstr名称 & "',A.对象,A.所有者,A.权限" & vbNewLine & _
                                            "  From zlprogprivs a" & vbNewLine & _
                                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & strOldName & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '删除原始功能名，由于存在级联删除关系
                    arrSQL(UBound(arrSQL)) = "  Delete From Zlprogfuncs a" & vbNewLine & _
                                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & strOldName & "'"
                    '系统、序号、功能 任意一个存在Null，级联删除将失效
                    If Nvl(rsCheck!系统, 0) = 0 Or Nvl(rsCheck!序号, 0) = 0 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From zlProgPrivs A " & vbNewLine & _
                            "Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "    And A.功能 = '" & strOldName & "'"
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From zlRoleGrant A " & vbNewLine & _
                            "Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "    And A.功能 = '" & strOldName & "'"
                    End If
                Else '报表名称未发生变化,只许更新功能说明
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新功能说明
                    arrSQL(UBound(arrSQL)) = "Update Zlprogfuncs A" & vbNewLine & _
                                                                "  Set  A.说明='" & mstr说明 & "'" & vbNewLine & _
                                                                "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & mstr名称 & "'"
                End If
                rsCheck.MoveNext
            Loop
            '发布到模块的报表功能名
            strSQL = "Select Nvl(B.系统, 0) 系统, B.程序id 序号, B.功能" & vbNewLine & _
                            "From Zlrptputs b, Zlprograms c, Zlprogfuncs d" & vbNewLine & _
                            "Where B.报表id =[1] And Nvl(B.系统, 0) = Nvl(C.系统, 0) And B.程序id = C.序号 And" & vbNewLine & _
                            "      Upper(C.部件) <> Upper('zl9Report') And Nvl(C.系统, 0) = Nvl(D.系统, 0) And C.序号 = D.序号 And D.功能 = B.功能"
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReortID)
            Do While Not rsCheck.EOF
                If strOldName <> mstr名称 And mlngSys = 0 Then   '非系统报表名称发生变化，则自动更新功能名称
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新Zlrptputs
                    arrSQL(UBound(arrSQL)) = "Update Zlrptputs Set 功能 = '" & mstr名称 & "' Where 报表id = " & mlngReortID & " And Nvl(系统, 0) = " & rsCheck!系统 & " And 程序id = " & rsCheck!序号
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能信息
                    arrSQL(UBound(arrSQL)) = "Insert Into Zlprogfuncs" & vbNewLine & _
                                                                "  (系统, 序号, 功能, 排列, 说明, 缺省值)" & vbNewLine & _
                                                                "  Select A.系统, A.序号, '" & mstr名称 & "', A.排列, '" & mstr说明 & "', A.缺省值" & vbNewLine & _
                                                                "  From Zlprogfuncs a" & vbNewLine & _
                                                                "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & rsCheck!功能 & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能授权信息
                    arrSQL(UBound(arrSQL)) = "Insert Into zlrolegrant" & vbNewLine & _
                                                                "  (系统,序号,角色,功能)" & vbNewLine & _
                                                                "  Select A.系统,A.序号,A.角色, '" & mstr名称 & "' " & vbNewLine & _
                                                                "  From zlrolegrant a" & vbNewLine & _
                                                                "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & rsCheck!功能 & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能对象权限信息
                    arrSQL(UBound(arrSQL)) = "Insert Into zlprogprivs" & vbNewLine & _
                                                                "  (系统,序号,功能,对象,所有者,权限)" & vbNewLine & _
                                                                "  Select A.系统,A.序号,'" & mstr名称 & "',A.对象,A.所有者,A.权限" & vbNewLine & _
                                                                "  From zlprogprivs a" & vbNewLine & _
                                                                "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & rsCheck!功能 & "'"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '删除原始功能名，由于存在级联删除关系
                    arrSQL(UBound(arrSQL)) = "  Delete From Zlprogfuncs a" & vbNewLine & _
                                                                "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & rsCheck!功能 & "'"
                Else '非系统报表说明变化或者固定报表变更，则只更新功能说明
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新功能说明
                    arrSQL(UBound(arrSQL)) = "Update Zlprogfuncs A" & vbNewLine & _
                                                                "  Set  A.说明='" & mstr说明 & "'" & vbNewLine & _
                                                                "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & " And A.功能 = '" & rsCheck!功能 & "'"
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
    Set grsReport = Nothing '清除缓存
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnGroupEdit And mlngGroupID <> 0 Or Not mblnGroupEdit And mlngReortID <> 0 Then txt名称.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    mblnOK = False
    txt编号.Text = mstr编码
    txt名称.Text = mstr名称
    txt说明.Text = mstr说明
    If mblnGroupEdit Then
        If mlngGroupID = 0 Then
            Caption = "新增报表组"
            txt编号.Text = GetNextNO(mblnGroupEdit)
        Else
            Caption = "修改报表组"
        End If
    Else
        If mlngReortID = 0 Then
            Caption = "新增报表"
            txt编号.Text = GetNextNO(mblnGroupEdit)
        Else
            Caption = "修改报表"
        End If
    End If
    If mlngSys > 0 Then txt编号.Enabled = False
End Sub

Private Sub txt编号_GotFocus()
    SelAll txt编号
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt名称_GotFocus()
    SelAll txt名称
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf InStr(GSTR_SBC, Chr(KeyAscii)) > 0 Then
        KeyAscii = Asc(Mid(GSTR_DBC, InStr(GSTR_SBC, Chr(KeyAscii)), 1))
    End If
End Sub

Private Sub txt名称_Validate(Cancel As Boolean)
    If txt名称.Text <> "" Then
        txt名称.Text = ConvertSBC(txt名称.Text)
    End If
End Sub

Private Sub txt说明_GotFocus()
    SelAll txt说明
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


