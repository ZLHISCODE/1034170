VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientsParas 
   BackColor       =   &H80000005&
   Caption         =   "站点运行控制"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmClientsParas.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   12690
   WindowState     =   2  'Maximized
   Begin VB.Timer timerConnect 
      Interval        =   1000
      Left            =   10920
      Top             =   4920
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   10800
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRemote 
      Caption         =   "远程控制(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   13
      Top             =   1260
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearClients 
      Caption         =   "清理3个月未登录工作站"
      Height          =   350
      Left            =   9990
      TabIndex        =   12
      Top             =   1260
      Width           =   2400
   End
   Begin VB.CommandButton cmdStopAll 
      Caption         =   "全部禁用"
      Height          =   350
      Left            =   8660
      TabIndex        =   10
      Top             =   1260
      Width           =   1100
   End
   Begin VB.TextBox txtLocate 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1297
      Width           =   1785
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   270
      TabIndex        =   5
      Top             =   5715
      Width           =   1100
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "禁用(&S)"
      Height          =   350
      Left            =   7560
      TabIndex        =   3
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   6450
      TabIndex        =   2
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   350
      Left            =   4230
      TabIndex        =   0
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改(&M)"
      Height          =   350
      Left            =   5340
      TabIndex        =   1
      Top             =   1260
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3960
      Left            =   255
      TabIndex        =   4
      Top             =   1650
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   6985
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilsIcon"
      SmallIcons      =   "ilsIcon"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "工作站"
         Object.Tag             =   "工作站"
         Text            =   "工作站"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Ip"
         Object.Tag             =   "Ip"
         Text            =   "IP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "CPU"
         Object.Tag             =   "CPU"
         Text            =   "CPU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "内存"
         Object.Tag             =   "内存"
         Text            =   "内存"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "硬盘"
         Object.Tag             =   "硬盘"
         Text            =   "硬盘"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "操作系统"
         Object.Tag             =   "操作系统"
         Text            =   "操作系统"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "部门"
         Object.Tag             =   "部门"
         Text            =   "部门"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "用途"
         Object.Tag             =   "用途"
         Text            =   "用途"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "说明"
         Object.Tag             =   "说明"
         Text            =   "说明"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "允许连接数"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "状态"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "启用视频源"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "最近登陆"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Port"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsIcon 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmClientsParas.frx":04F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPrompt 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblLocate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "机器名或IP(&L)"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "站点运行控制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   7
      Top             =   105
      Width           =   1440
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "对各站点进行增加、删除、修改，同时可禁止指定工作站的运行及站点参数的置换。"
      Height          =   345
      Left            =   1215
      TabIndex        =   6
      Top             =   750
      Width           =   7365
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   225
      Picture         =   "frmClientsParas.frx":0FC3
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmClientsParas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const StopColor = vbRed '禁用时的颜色
Const StartColor = &H80000008 '启用时的颜色
Dim mintColumn As Integer '

Private mintLastTime  As Integer    '记录连接的持续时间,用于超时后断开连接
Private mstrConnStat As String  '记录连接状态,1.开始 2.停止



Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

Private Sub cmdAdd_Click()
    Dim blnReturn As Boolean
    Dim strKey As String
    frmClientsEdit.ShowEdit "", "", 新增, blnReturn
    If Not blnReturn Then Exit Sub
    If Me.lvwMain.ListItems.Count = 0 Then
        '初始化信息
        Call LoadClientsInfor
        SetCtlEnabled
        Exit Sub
    End If
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    strKey = Me.lvwMain.SelectedItem.Key
    '初始化信息
    Call LoadClientsInfor
    err = 0
    On Error Resume Next
    Me.lvwMain.ListItems(strKey).Selected = True
    Me.lvwMain.ListItems(strKey).EnsureVisible
    SetCtlEnabled
    err = 0
End Sub

Private Sub cmdClearClients_Click()
    Dim strSql As String
    
    On Error GoTo errH
    If MsgBox("使用此功能将会删除所有三个月内未登录的工作站，" & vbCrLf & "确定要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If Me.lvwMain.Enabled Then Me.lvwMain.SetFocus
        Exit Sub
    End If
    strSql = "Zl_Zlclients_Deletebatch()"
    ExecuteProcedure strSql, Me.Caption
    Call LoadClientsInfor
    SetCtlEnabled
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdDel_Click()
    Dim strKey As String
    Dim strIp As String
    Dim intIndex As Long
    If Me.lvwMain.ListItems.Count = 0 Then Exit Sub
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("你是否真要删除工作站为" & Me.lvwMain.SelectedItem & "的站点吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If Me.lvwMain.Enabled Then Me.lvwMain.SetFocus
        Exit Sub
    End If
    If Me.lvwMain.Enabled Then Me.lvwMain.SetFocus
    
    err = 0
    On Error Resume Next
    Call ExecuteProcedure("Zl_Zlclients_Delete('" & Me.lvwMain.SelectedItem.Text & "')", Me.Caption)
    lvwMain.Tag = ""
    strKey = Me.lvwMain.SelectedItem
    With lvwMain
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    SetCtlEnabled
End Sub

Private Sub cmdModify_Click()
    Dim blnReturn As Boolean
    Dim strKey As String
    Dim strIp As String
    Dim strName As String
    
    If Me.lvwMain.ListItems.Count = 0 Then Exit Sub
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    strKey = Me.lvwMain.SelectedItem.Key
    strName = Me.lvwMain.SelectedItem.Text
    strIp = Me.lvwMain.SelectedItem.SubItems(1)
    frmClientsEdit.ShowEdit strIp, strName, 修改, blnReturn
    If Not blnReturn Then Exit Sub
    '初始化信息
    Call LoadClientsInfor
    err = 0
    On Error Resume Next
    Me.lvwMain.ListItems(strKey).Selected = True
    Me.lvwMain.ListItems(strKey).EnsureVisible
    err = 0
    SetCtlEnabled
End Sub

Private Sub cmdRefresh_Click()
    Dim strTxt As String
    Dim itm As ListItem

    If Not Me.lvwMain.SelectedItem Is Nothing Then
        strTxt = lvwMain.SelectedItem.Text
    End If
    
    Call LoadClientsInfor
    
    If strTxt <> "" Then
        For Each itm In lvwMain.ListItems
            If itm.Text = strTxt Then
                itm.Selected = True
                Call itm.EnsureVisible
                lvwMain_ItemClick itm
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cmdStop_Click()
    Dim itm As ListItem
    Dim bytTmp As Byte
    
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    Set itm = lvwMain.SelectedItem
    err = 0
    On Error Resume Next
    Call ExecuteProcedure("Zl_Zlclients_Control(0,'" & UCase(Me.lvwMain.SelectedItem.Text) & "','" & lvwMain.SelectedItem.SubItems(1) & "',Null,Null,Null,Null,Null,Null, " & IIf(itm.Tag = 1, 0, 1) & ")", Me.Caption)
    
    If itm.Tag = "1" Then
        SetSelItemColor itm, StartColor
        itm.Tag = "0"
    Else
        SetSelItemColor itm, StopColor
        itm.Tag = "1"
    End If
    If itm.Tag = "1" Then
        Me.CmdStop.Caption = "启用(&S)"
        lblPrompt.Caption = lvwMain.SelectedItem.Text & " 已禁用"
    Else
        Me.CmdStop.Caption = "禁用(&S)"
        lblPrompt.Caption = lvwMain.SelectedItem.Text & " 已启用"
    End If
    
End Sub

Private Sub cmdStopAll_Click()
    Dim i As Long, lngCount As Long
    Dim strErr As String
    Dim itm As ListItem
    
    On Error Resume Next
    cmdStopAll.Enabled = False
    lngCount = lvwMain.ListItems.Count
    
    For Each itm In lvwMain.ListItems
        i = i + 1
        lblPrompt.Caption = "正在处理第" & i & "个，共" & lngCount & "个"
        lblPrompt.Refresh
        Call ExecuteProcedure("Zl_Zlclients_Control(0,'" & UCase(itm.Text) & "','" & itm.SubItems(1) & "',Null,Null,Null,Null,Null,Null, " & IIf(cmdStopAll.Tag = "1", 0, 1) & ")", Me.Caption)
        
        If cmdStopAll.Tag = "1" Then
            SetSelItemColor itm, StartColor
            itm.Tag = 0
        Else
            SetSelItemColor itm, StopColor
            itm.Tag = 1
        End If
    
        If err.Number <> 0 Then
            strErr = IIf(strErr = "", "", strErr & ",") & itm.Text
            err.Clear
        End If
    Next
    
    If cmdStopAll.Tag = "" Or cmdStopAll.Tag = "0" Then
        cmdStopAll.Caption = "全部启用"
        cmdStopAll.Tag = "1"
    Else
        cmdStopAll.Caption = "全部禁用"
        cmdStopAll.Tag = "0"
    End If
    
    lblPrompt.Caption = "操作完成"
    cmdStopAll.Enabled = True
    lvwMain.Refresh
    
    If strErr <> "" Then
        If Len(strErr) > 4000 Then strErr = Mid(strErr, 1, 4000) & "......"
        MsgBox "对以下客户端的操作出错：" & vbCrLf & strErr, vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    txtLocate.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    End If
    If KeyCode = vbKeyDelete Then
        cmdDel_Click
    End If
End Sub

Private Sub Form_Resize()
    Dim lngWdt As Single
    
    err = 0
    On Error Resume Next
    lblNote.Width = ScaleWidth - lblNote.Left
    With cmdRefresh
        .Top = ScaleHeight - .Height - 50
    End With
    
    With lvwMain
        lngWdt = ScaleWidth - .Left
        .Width = lngWdt
        .Height = cmdRefresh.Top - .Top - 50
    End With
        
    With cmdClearClients
        .Left = ScaleWidth - .Width
    End With
    With cmdRemote
        .Left = cmdClearClients.Left - .Width - 200
    End With
    With cmdStopAll
        .Left = cmdRemote.Left - .Width
    End With
    With CmdStop
        .Left = cmdStopAll.Left - .Width
    End With
    With cmdDel
        .Left = CmdStop.Left - .Width
    End With
    With cmdModify
        .Left = cmdDel.Left - .Width
    End With
    With cmdAdd
        .Left = cmdModify.Left - .Width
    End With
    
End Sub

Private Sub LoadClientsInfor()
    '---------------------------------------------------------------------------------------------
    '功能：加载站点信息
    '参数：
    '返回：
    '---------------------------------------------------------------------------------------------
    Dim RsClients As New ADODB.Recordset
    Dim strSql As String
    Dim itm As ListItem
    Dim strKey As String, strErr As String, lngCount As Long
    Dim dateNow As Date

    err = 0
    On Error GoTo ErrHand:
    dateNow = CurrentDate()
    Set RsClients = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client", "")
    With RsClients
        
        lvwMain.ListItems.Clear
        lvwMain.Tag = ""
        If Not .EOF Then
            strKey = "K" & Nvl(!工作站)
        End If
        On Error Resume Next
        
        Do While Not .EOF
           Set itm = lvwMain.ListItems.Add(, "K" & Nvl(!工作站), Nvl(!工作站), 1, 1)
           If err.Number = 0 Then
                itm.SubItems(1) = Nvl(!IP)
                itm.SubItems(2) = Nvl(!cpu)
                itm.SubItems(3) = Nvl(!内存)
                itm.SubItems(4) = Nvl(!硬盘)
                itm.SubItems(5) = Nvl(!操作系统)
                itm.SubItems(6) = Nvl(!部门)
                itm.SubItems(7) = Nvl(!用途)
                itm.SubItems(8) = Nvl(!说明)
                itm.SubItems(9) = IIf(Nvl(!连接数, 0) = 0, "无限制", Nvl(!连接数, 0) & "个连接")
                If !状态 = 1 Then itm.SubItems(10) = "正在使用"
                itm.Tag = Nvl(!禁止使用, 0)
                itm.SubItems(11) = IIf(Nvl(!启用视频源, 0) = 0, "未启用", "已启用")
                itm.SubItems(12) = TimeGraded(Nvl(!最近登陆时间, Format("3000-01-01 01:01:01", "YYYY-MM-DD HH:mm:ss")), dateNow)
                
                itm.SubItems(13) = IIf(!状态 = 1, "未加载", "不在线")
            
                If Nvl(!禁止使用, 0) = 1 Then
                   SetSelItemColor itm, StopColor
                   lngCount = lngCount + 1
                Else
                   SetSelItemColor itm, StartColor
                End If
            Else
                strErr = IIf(strErr = "", "", strErr & ",") & !工作站 & "(" & !部门 & ")"
                err.Clear
            End If
            .MoveNext
        Loop
        
    End With
    If Me.lvwMain.ListItems.Count <> 0 Then
        Me.lvwMain.ListItems(strKey).Selected = True
        Me.lvwMain.ListItems(strKey).EnsureVisible
        lvwMain_ItemClick Me.lvwMain.SelectedItem
    End If
    
    If lngCount = lvwMain.ListItems.Count And lngCount <> 0 Then
        cmdStopAll.Caption = "全部启用"
        cmdStopAll.Tag = "1"
    End If
    
    If strErr <> "" Then
        If Len(strErr) > 4000 Then strErr = Mid(strErr, 1, 4000) & "......"
        MsgBox "以下客户端与其他机器名重复，请检查并更改机器名:" & vbCrLf & strErr, vbInformation, gstrSysName
    End If
    
    Call SetCtlEnabled
    
    Exit Sub
ErrHand:
    MsgBox "系统出现错误,错误为:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
    SetCtlEnabled
End Sub

Private Function TimeGraded(ByVal dateRecentlyTime As Date, ByVal dateNow As Date) As String
'功能：根据传入的时间进行分级，即将其分为不同的时间段，比如1小时前，2月前
'入参：
'       dateRecentlyTime：需要进行分级的时间
'       dateNow         ：当前时间

    Dim lngHour As Long, lngDay As Long, lngMonth As Long
    Dim strNote As String

    '当最小时间大于当前时间时，则返回“未知”
    If dateRecentlyTime = Format("3000-01-01 01:01:01", "YYYY-MM-DD HH:mm:ss") Then
        TimeGraded = "未知"
        Exit Function
    End If
    lngHour = DateDiff("h", dateRecentlyTime, dateNow)
    If lngHour <= 23 Then
        If lngHour = 0 Then
            strNote = "刚刚"   '1小时内
        Else
            strNote = lngHour & "小时前"
        End If
    Else
        If dateRecentlyTime > DateAdd("m", -1, dateNow) Then
            '1个月以内，用天表示
            lngDay = DateDiff("d", dateRecentlyTime, dateNow)
            strNote = lngDay & "天前"
        Else
            '大于1个月，用月表示
            lngMonth = DateDiff("M", dateRecentlyTime, dateNow)
            If DateAdd("m", lngMonth, dateRecentlyTime) > dateNow Then
                strNote = lngMonth - 1 & "月前"
            Else
                strNote = lngMonth & "月前"
            End If
        End If
    End If
    TimeGraded = strNote
End Function

Private Sub SetSelItemColor(ByVal itm As ListItem, ByVal lngColor As Long)
    Dim i As Integer
        
    '设置被选择的颜色
    itm.ForeColor = lngColor
    For i = 1 To itm.ListSubItems.Count
        itm.ListSubItems(i).ForeColor = lngColor
    Next
End Sub

Private Sub Form_Load()
    '初始化信息
    Call LoadClientsInfor
End Sub

Private Sub SetCtlEnabled()
    Dim blnNoClients As Boolean '没有站点
    Dim blnSel As Boolean
    
    blnSel = Not Me.lvwMain.SelectedItem Is Nothing
    blnNoClients = Me.lvwMain.ListItems.Count = 0
    
    Me.cmdDel.Enabled = Not blnNoClients And blnSel
    Me.cmdModify.Enabled = Not blnNoClients And blnSel
    Me.CmdStop.Enabled = Not blnNoClients And blnSel
    Me.cmdStopAll.Enabled = Not blnNoClients
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    Call cmdModify_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strPort As String, strTerminal As String
    
    strTerminal = lvwMain.SelectedItem.Text
    strPort = lvwMain.SelectedItem.SubItems(13)
    If strPort = "未加载" Then
        lvwMain.SelectedItem.SubItems(13) = Val(gclsBase.GetPara("允许远程控制", , , , strTerminal, "1001"))
    End If
    
    If Item.Tag = 1 Then
        Me.CmdStop.Caption = "启用(&S)"
    Else
        Me.CmdStop.Caption = "禁用(&S)"
    End If
    If lvwMain.Tag <> "" Then
        Call SetSelItemBold(lvwMain.ListItems(lvwMain.Tag), False)
    End If
    Call SetSelItemBold(Item, True)
    lvwMain.Tag = Item.Key
End Sub
Private Sub SetSelItemBold(ByVal itm As ListItem, ByVal blnBold As Boolean)
    Dim i As Integer
        
    '设置被选择的颜色
    itm.Bold = blnBold
    For i = 1 To itm.ListSubItems.Count
        itm.ListSubItems(i).Bold = blnBold
    Next
End Sub

Private Sub txtLocate_GotFocus()
    txtLocate.SelStart = 0
    txtLocate.SelLength = Len(txtLocate.Text)
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        Dim strTxt As String
        Dim i As Long, lngStart As Long, lngP As Long
        
        strTxt = UCase(Trim(txtLocate.Text) & "*")
        
        '从上次找到的位置之后继续找
        If txtLocate.Tag = strTxt Then
            lngStart = Val(lblLocate.Tag) + 1
        Else
            lngStart = 1
        End If
        
        For i = lngStart To lvwMain.ListItems.Count
            If UCase(lvwMain.ListItems(i).Text) Like strTxt Or lvwMain.ListItems(i).SubItems(1) Like strTxt Then
                lvwMain.ListItems(i).Selected = True
                Call lvwMain.ListItems(i).EnsureVisible
                lvwMain_ItemClick Me.lvwMain.SelectedItem
                
                lngP = i
                Exit For
            End If
        Next
        
        txtLocate.Tag = strTxt
        lblLocate.Tag = lngP
    End If
End Sub

Private Sub InitConnect()
    With winSock
        If .State <> sckClosed Then .Close
        .RemoteHost = lvwMain.SelectedItem.SubItems(1)
        .RemotePort = lvwMain.SelectedItem.SubItems(13)
    End With
End Sub

Private Sub winSock_Connect()
    winSock.SendData "请求远程"
    mstrConnStat = "开始"
    mintLastTime = 0
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strMsg As String
    Dim strPort As String, strUser As String, strPwd As String
    Dim strName As String, strErr As String
    Dim rsTmp As New ADODB.Recordset
    
    winSock.GetData strData
    mstrConnStat = "停止"
    If strData = "YES" Then
        ShowFlash ""
        strPort = winSock.RemoteHost
        strName = lvwMain.SelectedItem.Text
        Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client", strName)  '获取用户跟密码
        
        strUser = Nvl(rsTmp!管理员用户)
        strPwd = Decipher(Nvl(rsTmp!管理员密码))
        
        If strUser = "" Or strPwd = "" Then
            strMsg = "当前客户端没有设置远程连接的帐号密码，是否进行设置？"
            If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbYes Then
                frmClientsEdit.ShowEdit strPort, strName, 1, True, strUser, strPwd
            End If
        End If
        RunCommand "cmdkey /generic:termsrv/" & strPort & " /user:" & strUser & " /pass:" & strPwd
        RunCommand "mstsc /v: " & strPort & "  /admin", , , 0
        RunCommand "cmdkey /delete:Termsrv/" & strPort
    End If
End Sub


Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
     mstrConnStat = "停止"
     ShowFlash ""
     
     Select Case Number
        Case 10061
            MsgBox "对方并没有运行远程监听程序。"
        Case Else
            MsgBox Description
     End Select
    
End Sub

Private Sub cmdRemote_Click()
    Dim strSql As String, rsData As ADODB.Recordset
    Dim strIp As String, strTerminal As String
    Dim strState As String, strPort As String
    
    On Error GoTo errH
    strPort = lvwMain.SelectedItem.SubItems(13)
    strTerminal = lvwMain.SelectedItem.Text
    strIp = lvwMain.SelectedItem.SubItems(1)
    
    If strPort = "不在线" Then
        '不在线的时候重新查一下状态和信息
        strSql = "Select 1 from " & IIf(gblnRac, "G", "") & "v$Session where Terminal=[1]"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "", strTerminal)
        
        If rsData.RecordCount > 0 Then
            strPort = gclsBase.GetPara("允许远程控制", , , , strTerminal, "1001")
            If Val(strPort) <= 0 Then
                MsgBox "当前客户端没有开启监听，无法发起远程申请。": Exit Sub
            Else
                lvwMain.SelectedItem.SubItems(13) = Val(strPort)
            End If
        Else
            MsgBox "当前客户端并没有处于运行状态,无法发起远程申请。": Exit Sub
        End If
        
    ElseIf strPort = "-1" Or Val(strPort) < 0 Then
        MsgBox "当前客户端没有开启监听，无法发起远程申请。": Exit Sub
    End If
    

    If MsgBox("是否对客户端" & strTerminal & "（IP：" & strIp & ":" & Val(strPort) & "）" & "申请远程控制?", vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then Exit Sub

    
    mstrConnStat = "开始"
    mintLastTime = 0
    InitConnect
    winSock.Connect
    Exit Sub
errH:
    ShowFlash ""
    mintLastTime = 0
    mstrConnStat = "停止"
    MsgBox err.Description
End Sub
Private Sub timerConnect_Timer()
    '每秒进行一次刷新
    
    DoEvents
    If mstrConnStat = "开始" Then
    
        ShowFlash "正在等待对方响应..."
        mintLastTime = mintLastTime + 1
        
        If mintLastTime > 19 Then
             If winSock.State <> sckClosed Then winSock.Close
             
            ShowFlash ""
            MsgBox "超过20秒未接收到响应,连接中断,请重试"
            mintLastTime = 0
            mstrConnStat = "停止"
            
        End If
    ElseIf mstrConnStat = "停止" Then
        ShowFlash ""
    End If
End Sub

