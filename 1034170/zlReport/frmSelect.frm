VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "frmSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ListView lvw 
      Height          =   2850
      Left            =   2535
      TabIndex        =   2
      Top             =   555
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   5027
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6390
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6390
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   165
         Picture         =   "frmSelect.frx":014A
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   2355
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6390
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3660
      Width           =   6390
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4785
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3540
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   2745
         Top             =   15
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelect.frx":06D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   2760
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   4868
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4905
      TabIndex        =   8
      Top             =   315
      Width           =   435
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入：SQL及字段描述
Public strSQLList As String
Public strSQLTree As String
Public strFLDList As String
Public strFLDTree As String
Public strParName As String '参数名称
Public bytType As Byte      '参数数据类型
Public strMatch As String '输入匹配的内容
Public lngSeekHwnd As Long '用于定位窗体位置的控件
Public mintConnect As Integer           '数据连接编号

Public mblnMulti As Boolean '是否多选择
Public mblnOK As Boolean
Public mlngSel As Long  '绑定列的值等于这个值时选中

'出：未作格式处理的数据原始值
Public strOutBand As String '选择的绑定值,对应&B
Public strOutDisp As String '选择的显示值,对应&D

Private intPreNode As Long
Private blnItem As Boolean
Private blnSetFlex As Boolean, blnSetLvw As Boolean
Private rsList As ADODB.Recordset
Private strList As String
Private BlnSave As Boolean
Private rParent As RECT

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim strDisp As String, strBand As String
    
    On Error GoTo hErr
    
    strDisp = GetScript(strFLDList, "&D") '显示的字段名
    strBand = GetScript(strFLDList, "&B") '绑定的字段名
    
    If strDisp = "" Or strBand = "" Then
        MsgBox "选择器中没有定义条件的绑定及显示字段项目！", vbInformation, App.Title
        Exit Sub
    End If
    
    If mblnMulti Then
        '多选时自动返回的情况
        If Not lvw.Visible And lvw.ListItems.count = 1 Then
            lvw.ListItems(1).Checked = True
        End If
        
        For i = 1 To lvw.ListItems.count
            If lvw.ListItems(i).Checked Then
                If Split(lvw.ListItems(i).Tag, "|")(0) = "" Then
                    lvw.ListItems(i).Selected = True
                    lvw.ListItems(i).EnsureVisible
                    MsgBox "该行内容的""" & strDisp & """为空,不能在条件""" & strParName & """中显示！", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
                If Split(lvw.ListItems(i).Tag, "|")(1) = "" Then
                    lvw.ListItems(i).Selected = True
                    lvw.ListItems(i).EnsureVisible
                    MsgBox "该行内容的""" & strBand & """为空,不能与条件""" & strParName & """相绑定！", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
                '类型检查(无类型,无需检查)
            End If
        Next
        '返回显示串，绑定串
        strOutDisp = ""
        strOutBand = ""
        For i = 1 To lvw.ListItems.count
            If lvw.ListItems(i).Checked Then
                strOutDisp = strOutDisp & "," & Split(lvw.ListItems(i).Tag, "|")(0)
                strOutBand = strOutBand & "," & Split(lvw.ListItems(i).Tag, "|")(1)
            End If
        Next
        If strOutDisp = "" Or strOutBand = "" Then
            MsgBox "没有选择任何内容！", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        ElseIf UBound(Split(strOutBand, ",")) > 1000 Then
            MsgBox "选择的内容过多！", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        strOutDisp = Mid(strOutDisp, 2)
        strOutBand = " IN (" & Mid(strOutBand, 2) & ") "
    Else
        If lvw.SelectedItem Is Nothing Then
            MsgBox "没有选择任何内容！", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        If InStr(lvw.SelectedItem.Tag, "|") <= 0 Then
            MsgBox "该行内容的为空，请检查数据源！", vbInformation, App.Title
            Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(0) = "" Then
            MsgBox "该行内容的""" & strDisp & """为空,不能在条件""" & strParName & """中显示！", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(1) = "" Then
            MsgBox "该行内容的""" & strBand & """为空,不能与条件""" & strParName & """相绑定！", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        
        '类型检查
        Select Case bytType
            Case 1
                If Not IsNumeric(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "项目""" & strBand & """的内容非数字型,不能被选择！", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
            Case 2
                If Not IsDate(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "项目""" & strBand & """的内容非日期型,不能被选择！", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
        End Select
    
        strOutDisp = Split(lvw.SelectedItem.Tag, "|")(0)
        strOutBand = Split(lvw.SelectedItem.Tag, "|")(1)
    End If
    
    mblnOK = True
    
    On Error Resume Next
    Hide
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub Form_Activate()
    If tvw_s.Visible Then
        If Not tvw_s.SelectedItem Is Nothing Then
            If tvw_s.SelectedItem.Key = "ALL" Then lvw.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Not lvw.Visible Then Exit Sub
        
        For i = 1 To lvw.ListItems.count
            lvw.ListItems(i).Checked = True
        Next
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Not lvw.Visible Then Exit Sub
        
        For i = 1 To lvw.ListItems.count
            lvw.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim lngW As Long, i As Integer
    
    If Not InDesign Then
        glngSelProc = GetWindowLong(hwnd, GWL_WNDPROC)
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SelMessage)
    End If
    
    mblnOK = False
    BlnSave = True
    blnSetFlex = False '是否已经对表格恢复宽度
    blnSetLvw = False
    intPreNode = 0
    
    strOutBand = ""
    strOutDisp = ""
    
    lvw.Tag = strParName
    
    Me.Caption = strParName & "选择器"
    
    strSQLList = Replace(strSQLList, "[*]", strMatch)
    strSQLTree = Replace(strSQLTree, "[*]", strMatch)
    
    If strSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
        If Not FillList Then BlnSave = False: Unload Me: Exit Sub
    Else
        tvw_s.Visible = True
        If Not FillTree Then BlnSave = False: Unload Me: Exit Sub
        If tvw_s.Nodes.count > 0 Then
            tvw_s.Nodes(1).Selected = True
            If Not tvw_s.Nodes(1).Child Is Nothing And strMatch = "" Then
                tvw_s.Nodes(1).Child.Selected = True
            End If
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        End If
    End If
    
    lvw.Checkboxes = mblnMulti
    lvw.ToolTipText = "全选(Ctrl+A),全清(Ctrl+R)"
    
    '输入匹配自动返回
    If strMatch <> "" Then
        If rsList.RecordCount = 1 Then
            BlnSave = False
            Call cmdOK_Click
            Unload Me: Exit Sub
        ElseIf rsList.RecordCount = 0 Then
            MsgBox "没有找到相匹配的项目,请重新输入！", vbInformation, App.Title
            BlnSave = False
            Call cmdCancel_Click: Exit Sub
        End If
    End If
    
    Call Form_Resize
    
    '窗体及列表缺省宽度
    If lvw.ColumnHeaders.count = 1 Then
        lvw.ColumnHeaders(1).Width = 2500
        Me.Width = 3000 + IIF(strSQLTree = "", 0, tvw_s.Width + pic.Width)
    Else
        For i = 1 To lvw.ColumnHeaders.count
            lngW = lngW + lvw.ColumnHeaders(i).Width
        Next
        Me.Width = lngW + 500 + IIF(strSQLTree = "", 0, tvw_s.Width + pic.Width)
        If Me.Width < 3000 Then Me.Width = 3000
    End If
    
    If strSQLTree <> "" Then
        If Me.Width < (tvw_s.Width + pic.Width) * 2.2 Then Me.Width = (tvw_s.Width + pic.Width) * 2.2
    End If
    
    RestoreWinState Me, App.ProductName, strParName
    
    If strSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
    Else
        tvw_s.Visible = True
    End If
    
    '定位
    If lngSeekHwnd <> 0 Then
        Call Form_Resize
        GetWindowRect lngSeekHwnd, rParent
        If rParent.Top >= Me.Height / 15 Then
            Me.Top = rParent.Bottom * 15 - Me.Height + 30
        Else
            Me.Top = (rParent.Bottom - rParent.Top) * 15 + 30
        End If
        If rParent.Left >= Me.Width / 15 Then
            Me.Left = rParent.Right * 15 - Me.Width + 30
        Else
            Me.Left = (rParent.Right - rParent.Left) * 15 + 30
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim lngTVW As Long
    lngTVW = IIF(tvw_s.Visible, tvw_s.Width + pic.Width, 0)
    
    tvw_s.Left = Me.ScaleLeft
    tvw_s.Top = picInfo.Top + picInfo.Height + 15
    tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height - 15
    
    pic.Left = tvw_s.Left + tvw_s.Width
    pic.Top = tvw_s.Top
    pic.Height = tvw_s.Height
    
    lvw.Left = Me.ScaleLeft + lngTVW
    lvw.Top = tvw_s.Top
    lvw.Height = tvw_s.Height
    lvw.Width = Me.ScaleWidth - lngTVW
    
    lbl.Left = lvw.Left
    lbl.Top = lvw.Top
    lbl.Width = lvw.Width
    lbl.Height = lvw.Height
    
    If ScaleWidth - cmdCancel.Width - 300 >= 1445 Then
        cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 45
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strMatch = ""
    lngSeekHwnd = 0
    If BlnSave Then SaveWinState Me, App.ProductName, strParName
    If Not InDesign Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngSelProc)
End Sub

Private Sub lvw_DblClick()
    If blnItem Then Call cmdOK_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnItem = True
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub lvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnItem = False
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or lvw.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        lvw.Left = lvw.Left + X
        lvw.Width = lvw.Width - X
        
        lbl.Left = lbl.Left + X
        lbl.Width = lbl.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Index = intPreNode Then Exit Sub
    intPreNode = Node.Index
    DoEvents
    Call FillList(Node.Tag)
End Sub

Private Function FillTree() As Boolean
'功能：根据定义数据源及字段属性，将分类数据显示在TreeView中
'返回：操作是否成功(用户非正常定义)
    Dim rstmp As New ADODB.Recordset
    Dim i As Integer, objNode As Node
    Dim strSel As String, strRela As String
    
    On Error GoTo errH
    
    strSel = GetScript(strFLDTree, "&S")
    strRela = GetScript(strFLDTree, "&R")
    
    If strSel = "" Or strRela = "" Then
        MsgBox "未发现用于选择或与明细列表相关联的字段项目！", vbInformation, App.Title
        Exit Function
    End If
    Call OpenRecord(rstmp, RemoveNote(strSQLTree), Me.Caption & "_FillTree", mintConnect) 'SQL一般固定,[*]在SQL的''中,类型无法处理
    
    tvw_s.Nodes.Clear
        
    If InStr("|" & UCase(strFLDTree), "|ID,") > 0 And InStr("|" & UCase(strFLDTree), "|上级ID,") > 0 Then
        '采用树形列表显示
        Set objNode = tvw_s.Nodes.Add(, , "ALL", "所有项目", 1)
        objNode.Tag = "ALL"
        objNode.Expanded = True
        
        For i = 1 To rstmp.RecordCount
            If IsNull(rstmp!上级ID) Then
                Set objNode = tvw_s.Nodes.Add("ALL", 4, "_" & rstmp!id, IIF(IsNull(rstmp.Fields(strSel).Value), "", rstmp.Fields(strSel).Value), 1)
            Else
                Set objNode = tvw_s.Nodes.Add("_" & rstmp!上级ID, 4, "_" & rstmp!id, IIF(IsNull(rstmp.Fields(strSel).Value), "", rstmp.Fields(strSel).Value), 1)
            End If
            objNode.Tag = IIF(IsNull(rstmp.Fields(strRela).Value), "", rstmp.Fields(strRela).Value)
            rstmp.MoveNext
        Next
    Else
        '采用一般列表显示
        For i = 1 To rstmp.RecordCount
            Set objNode = tvw_s.Nodes.Add(, , , IIF(IsNull(rstmp.Fields(strSel).Value), "", rstmp.Fields(strSel).Value), 1)
            objNode.Tag = IIF(IsNull(rstmp.Fields(strRela).Value), "", rstmp.Fields(strRela).Value)
            rstmp.MoveNext
        Next
    End If

    FillTree = True
    Exit Function
errH:
    If Err.Number = 35601 Then
        MsgBox "不能正常处理树形列表，条件选择器不能使用！", vbExclamation, App.Title
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Private Function GetRelaSQL(ByVal strSQL As String, ByVal strFld As String, ByVal strKey As String) As String
'功能：处理关联的SQL
    Dim i As Integer, strRela As String
    
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), "&R") > 0 Then
            strRela = Split(Split(strFld, "|")(i), ",")(0)
            If strKey = "" Then
                GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & " is NULL"
            Else
                Select Case Split(Split(strFld, "|")(i), ",")(1)
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=" & strKey
                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "='" & strKey & "'"
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                        If Format(strKey, "hh:mm:ss") = "00:00:00" Then
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & ">=To_Date('" & Format(strKey, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & strRela & "<=To_Date('" & Format(strKey, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=To_Date('" & Format(strKey, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                End Select
            End If
            Exit Function
        End If
    Next
End Function

Private Function GetScript(strFld As String, strType As String) As String
'功能：根据指定的字段描述返回字段名
'参数：strType="&S &D &B &R"
'说明：适用于唯一性描述字段(如绑定字段)
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), strType) > 0 Then
            GetScript = Split(Split(strFld, "|")(i), ",")(0)
            Exit Function
        End If
    Next
End Function

Private Function HaveScript(strFld As String, strName As String, strType As String) As Boolean
'功能：判断在字段描述中，指定的字段是否具有指定的描述属性
'参数：strName=字段名,strFld=字段描述串,strType="&S &D &B &R"
'返回：False=未发现字段或字段不具有指定描述
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If Split(Split(strFld, "|")(i), ",")(0) = strName Then
            If InStr(Split(Split(strFld, "|")(i), ",")(2), strType) > 0 Then
                HaveScript = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function FillList(Optional strKey As String, Optional blnSort As Boolean) As Boolean
'功能：根据当前选择的分类或在无分类时处理对应的明细列表
'参数：strKey=分类列表中的当前关联值
'说明：根据数据量的多少，确定用ListView还是DataGrid
    Dim strSQL As String, i As Long, j As Integer
    Dim objitem As ListItem, strValue As String
    Dim strDisp As String, strBand As String
    
    On Error GoTo errH
    
    lvw.ListItems.Clear
    
    '可能为只处理排序
    If Not blnSort Then
        If strSQLTree = "" Then
            strSQL = strSQLList
        Else
            '动态将明细数据处理为只读取关联的分类部分(处理 Order by 子句)
            If strKey = "ALL" Then
                strSQL = strSQLList
            Else
                strSQL = GetRelaSQL(RemoveOrderBy(strSQLList), strFLDList, strKey)
            End If
            
            If strSQL = "" Then
                MsgBox "该类数据读取失败！", vbInformation, App.Title
                Exit Function
            End If
        End If
        
        Screen.MousePointer = 11
        Me.Refresh
        
        Set rsList = New ADODB.Recordset
        Call OpenRecord(rsList, RemoveNote(strSQL), Me.Caption & "_FillList", mintConnect) 'SQL一般固定,[*]在SQL的''中,类型无法处理
    End If
    
    If Not rsList.EOF Then
        If lvw.ColumnHeaders.count = 0 Then Call AddListCols
        
        strDisp = GetScript(strFLDList, "&D") '显示值项目
        strBand = GetScript(strFLDList, "&B") '绑定值项目
        
        For i = 1 To rsList.RecordCount
            strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(1).Text))
            If lvw.ColumnHeaders(1).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(1).Tag)
            Set objitem = lvw.ListItems.Add(, , strValue, , 1)
            For j = 2 To lvw.ColumnHeaders.count
                strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(j).Text))
                If lvw.ColumnHeaders(j).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(j).Tag)
                objitem.SubItems(j - 1) = strValue
            Next
            
            '将显示值及绑定值保存在TAG中,因为不一定这些字段会为选择字段
            '格式为"显示值|绑定值"
            If strDisp <> "" Then
                objitem.Tag = IIF(IsNull(rsList.Fields(strDisp).Value), "", rsList.Fields(strDisp).Value)
            End If
            objitem.Tag = objitem.Tag & "|"
            If strBand <> "" Then
                objitem.Tag = objitem.Tag & IIF(IsNull(rsList.Fields(strBand).Value), "", rsList.Fields(strBand).Value)
                If mlngSel <> 0 And Val(rsList.Fields(strBand).Value & "") = mlngSel Then objitem.Selected = True: Call objitem.EnsureVisible
            End If
                            
            rsList.MoveNext
        Next
        
        '自动调整列宽
        Call AutoSizeCol(lvw)
        
        If Not Visible Or Not blnSetLvw Then
            Call RestoreListViewState(lvw, App.ProductName & "\" & Me.name & strParName)
            blnSetLvw = True
        End If
        lblInfo.Caption = "共 " & rsList.RecordCount & " 个明细项目."
    Else
        '没有数据时，显示空的ListView(带列头)
        If lvw.ColumnHeaders.count = 0 Then Call AddListCols
        lblInfo.Caption = "没有明细项目."
    End If
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddListCols()
'功能：根据strFLDList字段描述值,为ListView增加列头
    Dim i As Integer, j As Integer, strFld As String
    Dim objCol As ColumnHeader
    
    For i = 0 To UBound(Split(strFLDList, "|"))
        strFld = Split(strFLDList, "|")(i)
        If strFld Like "*&S*" Then
            Set objCol = lvw.ColumnHeaders.Add(, "_" & Split(strFld, ",")(0), Split(strFld, ",")(0))
            
            objCol.Width = Me.TextWidth(Split(strFld, ",")(0) & "字")
            
            '根据字段名及类型设置对齐(列1只能左对齐)
            Select Case Split(strFld, ",")(1)
                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    If rsList.Fields(objCol.Text).NumericScale > 0 Then
                        j = rsList.Fields(objCol.Text).NumericScale
                        objCol.Tag = "0." & String(IIF(j > 2, 2, j), "0; ;")
                        If objCol.Index <> 1 Then objCol.Alignment = lvwColumnRight
                    ElseIf objCol.Index <> 1 Then
                        If rsList.Fields(objCol.Text).Precision < 3 Then
                            objCol.Alignment = lvwColumnCenter
                        Else
                            objCol.Alignment = lvwColumnLeft
                        End If
                    End If
                    If objCol.Text Like "*价" Then objCol.Tag = "0.000"
                    If objCol.Text Like "*额" Then objCol.Tag = "0.00"
                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
                Case Else
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
            End Select
            If objCol.Text Like "*单位*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
            If objCol.Text Like "*否*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
        End If
    Next
End Sub

Private Function GetValue(objFld As Field) As String
'功能:根据字段内容取合适的显示值
    Dim strValue As String
    Select Case objFld.type
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            strValue = IIF(IsNull(objFld.Value), 0, objFld.Value)
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
            If Format(strValue, "HH:mm:ss") = "00:00:00" Then
                strValue = Format(strValue, "yyyy-MM-dd")
            Else
                strValue = Format(strValue, "yyyy-MM-dd HH:mm:ss")
            End If
        Case Else
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
    End Select
    GetValue = strValue
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'功能：按列排序
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub
