VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientsParas 
   BackColor       =   &H80000005&
   Caption         =   "վ�����п���"
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
      Caption         =   "Զ�̿���(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   13
      Top             =   1260
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearClients 
      Caption         =   "����3����δ��¼����վ"
      Height          =   350
      Left            =   9990
      TabIndex        =   12
      Top             =   1260
      Width           =   2400
   End
   Begin VB.CommandButton cmdStopAll 
      Caption         =   "ȫ������"
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
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   270
      TabIndex        =   5
      Top             =   5715
      Width           =   1100
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   7560
      TabIndex        =   3
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   6450
      TabIndex        =   2
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   4230
      TabIndex        =   0
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�(&M)"
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
         Key             =   "����վ"
         Object.Tag             =   "����վ"
         Text            =   "����վ"
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
         Key             =   "�ڴ�"
         Object.Tag             =   "�ڴ�"
         Text            =   "�ڴ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Ӳ��"
         Object.Tag             =   "Ӳ��"
         Text            =   "Ӳ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "����ϵͳ"
         Object.Tag             =   "����ϵͳ"
         Text            =   "����ϵͳ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "��;"
         Object.Tag             =   "��;"
         Text            =   "��;"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "˵��"
         Object.Tag             =   "˵��"
         Text            =   "˵��"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "����������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "״̬"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "������ƵԴ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "�����½"
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
      Caption         =   "��������IP(&L)"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "վ�����п���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�Ը�վ��������ӡ�ɾ�����޸ģ�ͬʱ�ɽ�ָֹ������վ�����м�վ��������û���"
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
Const StopColor = vbRed '����ʱ����ɫ
Const StartColor = &H80000008 '����ʱ����ɫ
Dim mintColumn As Integer '

Private mintLastTime  As Integer    '��¼���ӵĳ���ʱ��,���ڳ�ʱ��Ͽ�����
Private mstrConnStat As String  '��¼����״̬,1.��ʼ 2.ֹͣ



Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Private Sub cmdAdd_Click()
    Dim blnReturn As Boolean
    Dim strKey As String
    frmClientsEdit.ShowEdit "", "", ����, blnReturn
    If Not blnReturn Then Exit Sub
    If Me.lvwMain.ListItems.Count = 0 Then
        '��ʼ����Ϣ
        Call LoadClientsInfor
        SetCtlEnabled
        Exit Sub
    End If
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    strKey = Me.lvwMain.SelectedItem.Key
    '��ʼ����Ϣ
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
    If MsgBox("ʹ�ô˹��ܽ���ɾ��������������δ��¼�Ĺ���վ��" & vbCrLf & "ȷ��Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
    If MsgBox("���Ƿ���Ҫɾ������վΪ" & Me.lvwMain.SelectedItem & "��վ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
    frmClientsEdit.ShowEdit strIp, strName, �޸�, blnReturn
    If Not blnReturn Then Exit Sub
    '��ʼ����Ϣ
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
        Me.CmdStop.Caption = "����(&S)"
        lblPrompt.Caption = lvwMain.SelectedItem.Text & " �ѽ���"
    Else
        Me.CmdStop.Caption = "����(&S)"
        lblPrompt.Caption = lvwMain.SelectedItem.Text & " ������"
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
        lblPrompt.Caption = "���ڴ����" & i & "������" & lngCount & "��"
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
        cmdStopAll.Caption = "ȫ������"
        cmdStopAll.Tag = "1"
    Else
        cmdStopAll.Caption = "ȫ������"
        cmdStopAll.Tag = "0"
    End If
    
    lblPrompt.Caption = "�������"
    cmdStopAll.Enabled = True
    lvwMain.Refresh
    
    If strErr <> "" Then
        If Len(strErr) > 4000 Then strErr = Mid(strErr, 1, 4000) & "......"
        MsgBox "�����¿ͻ��˵Ĳ�������" & vbCrLf & strErr, vbInformation, gstrSysName
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
    '���ܣ�����վ����Ϣ
    '������
    '���أ�
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
            strKey = "K" & Nvl(!����վ)
        End If
        On Error Resume Next
        
        Do While Not .EOF
           Set itm = lvwMain.ListItems.Add(, "K" & Nvl(!����վ), Nvl(!����վ), 1, 1)
           If err.Number = 0 Then
                itm.SubItems(1) = Nvl(!IP)
                itm.SubItems(2) = Nvl(!cpu)
                itm.SubItems(3) = Nvl(!�ڴ�)
                itm.SubItems(4) = Nvl(!Ӳ��)
                itm.SubItems(5) = Nvl(!����ϵͳ)
                itm.SubItems(6) = Nvl(!����)
                itm.SubItems(7) = Nvl(!��;)
                itm.SubItems(8) = Nvl(!˵��)
                itm.SubItems(9) = IIf(Nvl(!������, 0) = 0, "������", Nvl(!������, 0) & "������")
                If !״̬ = 1 Then itm.SubItems(10) = "����ʹ��"
                itm.Tag = Nvl(!��ֹʹ��, 0)
                itm.SubItems(11) = IIf(Nvl(!������ƵԴ, 0) = 0, "δ����", "������")
                itm.SubItems(12) = TimeGraded(Nvl(!�����½ʱ��, Format("3000-01-01 01:01:01", "YYYY-MM-DD HH:mm:ss")), dateNow)
                
                itm.SubItems(13) = IIf(!״̬ = 1, "δ����", "������")
            
                If Nvl(!��ֹʹ��, 0) = 1 Then
                   SetSelItemColor itm, StopColor
                   lngCount = lngCount + 1
                Else
                   SetSelItemColor itm, StartColor
                End If
            Else
                strErr = IIf(strErr = "", "", strErr & ",") & !����վ & "(" & !���� & ")"
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
        cmdStopAll.Caption = "ȫ������"
        cmdStopAll.Tag = "1"
    End If
    
    If strErr <> "" Then
        If Len(strErr) > 4000 Then strErr = Mid(strErr, 1, 4000) & "......"
        MsgBox "���¿ͻ����������������ظ������鲢���Ļ�����:" & vbCrLf & strErr, vbInformation, gstrSysName
    End If
    
    Call SetCtlEnabled
    
    Exit Sub
ErrHand:
    MsgBox "ϵͳ���ִ���,����Ϊ:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
    SetCtlEnabled
End Sub

Private Function TimeGraded(ByVal dateRecentlyTime As Date, ByVal dateNow As Date) As String
'���ܣ����ݴ����ʱ����зּ����������Ϊ��ͬ��ʱ��Σ�����1Сʱǰ��2��ǰ
'��Σ�
'       dateRecentlyTime����Ҫ���зּ���ʱ��
'       dateNow         ����ǰʱ��

    Dim lngHour As Long, lngDay As Long, lngMonth As Long
    Dim strNote As String

    '����Сʱ����ڵ�ǰʱ��ʱ���򷵻ء�δ֪��
    If dateRecentlyTime = Format("3000-01-01 01:01:01", "YYYY-MM-DD HH:mm:ss") Then
        TimeGraded = "δ֪"
        Exit Function
    End If
    lngHour = DateDiff("h", dateRecentlyTime, dateNow)
    If lngHour <= 23 Then
        If lngHour = 0 Then
            strNote = "�ո�"   '1Сʱ��
        Else
            strNote = lngHour & "Сʱǰ"
        End If
    Else
        If dateRecentlyTime > DateAdd("m", -1, dateNow) Then
            '1�������ڣ������ʾ
            lngDay = DateDiff("d", dateRecentlyTime, dateNow)
            strNote = lngDay & "��ǰ"
        Else
            '����1���£����±�ʾ
            lngMonth = DateDiff("M", dateRecentlyTime, dateNow)
            If DateAdd("m", lngMonth, dateRecentlyTime) > dateNow Then
                strNote = lngMonth - 1 & "��ǰ"
            Else
                strNote = lngMonth & "��ǰ"
            End If
        End If
    End If
    TimeGraded = strNote
End Function

Private Sub SetSelItemColor(ByVal itm As ListItem, ByVal lngColor As Long)
    Dim i As Integer
        
    '���ñ�ѡ�����ɫ
    itm.ForeColor = lngColor
    For i = 1 To itm.ListSubItems.Count
        itm.ListSubItems(i).ForeColor = lngColor
    Next
End Sub

Private Sub Form_Load()
    '��ʼ����Ϣ
    Call LoadClientsInfor
End Sub

Private Sub SetCtlEnabled()
    Dim blnNoClients As Boolean 'û��վ��
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
    If strPort = "δ����" Then
        lvwMain.SelectedItem.SubItems(13) = Val(gclsBase.GetPara("����Զ�̿���", , , , strTerminal, "1001"))
    End If
    
    If Item.Tag = 1 Then
        Me.CmdStop.Caption = "����(&S)"
    Else
        Me.CmdStop.Caption = "����(&S)"
    End If
    If lvwMain.Tag <> "" Then
        Call SetSelItemBold(lvwMain.ListItems(lvwMain.Tag), False)
    End If
    Call SetSelItemBold(Item, True)
    lvwMain.Tag = Item.Key
End Sub
Private Sub SetSelItemBold(ByVal itm As ListItem, ByVal blnBold As Boolean)
    Dim i As Integer
        
    '���ñ�ѡ�����ɫ
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
        
        '���ϴ��ҵ���λ��֮�������
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
    winSock.SendData "����Զ��"
    mstrConnStat = "��ʼ"
    mintLastTime = 0
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strMsg As String
    Dim strPort As String, strUser As String, strPwd As String
    Dim strName As String, strErr As String
    Dim rsTmp As New ADODB.Recordset
    
    winSock.GetData strData
    mstrConnStat = "ֹͣ"
    If strData = "YES" Then
        ShowFlash ""
        strPort = winSock.RemoteHost
        strName = lvwMain.SelectedItem.Text
        Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client", strName)  '��ȡ�û�������
        
        strUser = Nvl(rsTmp!����Ա�û�)
        strPwd = Decipher(Nvl(rsTmp!����Ա����))
        
        If strUser = "" Or strPwd = "" Then
            strMsg = "��ǰ�ͻ���û������Զ�����ӵ��ʺ����룬�Ƿ�������ã�"
            If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbYes Then
                frmClientsEdit.ShowEdit strPort, strName, 1, True, strUser, strPwd
            End If
        End If
        RunCommand "cmdkey /generic:termsrv/" & strPort & " /user:" & strUser & " /pass:" & strPwd
        RunCommand "mstsc /v: " & strPort & "  /admin", , , 0
        RunCommand "cmdkey /delete:Termsrv/" & strPort
    End If
End Sub


Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
     mstrConnStat = "ֹͣ"
     ShowFlash ""
     
     Select Case Number
        Case 10061
            MsgBox "�Է���û������Զ�̼�������"
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
    
    If strPort = "������" Then
        '�����ߵ�ʱ�����²�һ��״̬����Ϣ
        strSql = "Select 1 from " & IIf(gblnRac, "G", "") & "v$Session where Terminal=[1]"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "", strTerminal)
        
        If rsData.RecordCount > 0 Then
            strPort = gclsBase.GetPara("����Զ�̿���", , , , strTerminal, "1001")
            If Val(strPort) <= 0 Then
                MsgBox "��ǰ�ͻ���û�п����������޷�����Զ�����롣": Exit Sub
            Else
                lvwMain.SelectedItem.SubItems(13) = Val(strPort)
            End If
        Else
            MsgBox "��ǰ�ͻ��˲�û�д�������״̬,�޷�����Զ�����롣": Exit Sub
        End If
        
    ElseIf strPort = "-1" Or Val(strPort) < 0 Then
        MsgBox "��ǰ�ͻ���û�п����������޷�����Զ�����롣": Exit Sub
    End If
    

    If MsgBox("�Ƿ�Կͻ���" & strTerminal & "��IP��" & strIp & ":" & Val(strPort) & "��" & "����Զ�̿���?", vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbNo Then Exit Sub

    
    mstrConnStat = "��ʼ"
    mintLastTime = 0
    InitConnect
    winSock.Connect
    Exit Sub
errH:
    ShowFlash ""
    mintLastTime = 0
    mstrConnStat = "ֹͣ"
    MsgBox err.Description
End Sub
Private Sub timerConnect_Timer()
    'ÿ�����һ��ˢ��
    
    DoEvents
    If mstrConnStat = "��ʼ" Then
    
        ShowFlash "���ڵȴ��Է���Ӧ..."
        mintLastTime = mintLastTime + 1
        
        If mintLastTime > 19 Then
             If winSock.State <> sckClosed Then winSock.Close
             
            ShowFlash ""
            MsgBox "����20��δ���յ���Ӧ,�����ж�,������"
            mintLastTime = 0
            mstrConnStat = "ֹͣ"
            
        End If
    ElseIf mstrConnStat = "ֹͣ" Then
        ShowFlash ""
    End If
End Sub

