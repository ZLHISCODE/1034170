VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClientsUpgrade 
   BackColor       =   &H80000005&
   Caption         =   "վ�㲿������"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmClientsUpgrade.frx":0000
   ScaleHeight     =   5910
   ScaleMode       =   0  'User
   ScaleWidth      =   11783.4
   WindowState     =   2  'Maximized
   Begin VB.Timer timerConnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton cmdkillProcess 
      Caption         =   "�ͻ��˽��̹���(&P)"
      Height          =   350
      Left            =   4410
      TabIndex        =   22
      Top             =   5535
      Width           =   1800
   End
   Begin VB.CommandButton cmdClientModify 
      Caption         =   "�ͻ��˿����޸�(&M)"
      Height          =   350
      Left            =   5520
      TabIndex        =   21
      Top             =   5115
      Width           =   1965
   End
   Begin VB.ComboBox cboUpResult 
      Appearance      =   0  'Flat
      Height          =   276
      Left            =   6315
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1182
      Width           =   1000
   End
   Begin VB.CommandButton cmdClearUpLog 
      Caption         =   "�������������־"
      Height          =   350
      Left            =   7320
      TabIndex        =   18
      ToolTipText     =   "��һ������ʱ,�����������ø�վ�������״̬Ϊ""δ����"""
      Top             =   1155
      Width           =   1668
   End
   Begin VB.PictureBox Piccmb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2520
      ScaleHeight     =   240
      ScaleWidth      =   915
      TabIndex        =   17
      Top             =   1185
      Width           =   945
      Begin VB.ComboBox cboFind 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1000
      End
   End
   Begin VB.CommandButton cmdUpdateS 
      Caption         =   "��������״̬"
      Height          =   350
      Left            =   9000
      TabIndex        =   5
      ToolTipText     =   "��һ������ʱ,�����������ø�վ�������״̬Ϊ""δ����"""
      Top             =   1155
      Width           =   1425
   End
   Begin VB.CommandButton cmd�û��������� 
      Caption         =   "��������(&J)"
      Height          =   350
      Left            =   7530
      TabIndex        =   10
      ToolTipText     =   "�ͻ���ΪUserȨ��,����ʱʹ�õĹ���Ա�û�����������"
      Top             =   5115
      Width           =   1200
   End
   Begin VB.CommandButton cmdԤ�������� 
      Caption         =   "Ԥ��������(&K)"
      Height          =   350
      Left            =   8715
      TabIndex        =   11
      Top             =   5115
      Width           =   1416
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H80000005&
      Caption         =   "FTP"
      Height          =   180
      Index           =   1
      Left            =   4725
      TabIndex        =   16
      Top             =   5200
      Width           =   810
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H80000005&
      Caption         =   "�ļ�����"
      Height          =   180
      Index           =   0
      Left            =   3705
      TabIndex        =   9
      Top             =   5200
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.CommandButton cmdӦ�� 
      Caption         =   "Ӧ���ڱ���(&P)��"
      Height          =   350
      Left            =   900
      TabIndex        =   7
      Top             =   5115
      Width           =   1575
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��������(&O)"
      Height          =   350
      Left            =   2475
      TabIndex        =   8
      Top             =   5115
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsClients 
      Height          =   3390
      Left            =   150
      TabIndex        =   0
      Top             =   1515
      Width           =   11460
      _cx             =   20214
      _cy             =   5980
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483643
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmClientsUpgrade.frx":04F9
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   810
      TabIndex        =   2
      Text            =   "255.255.255.255"
      Top             =   1185
      Width           =   1680
   End
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3705
      ScaleHeight     =   285
      ScaleWidth      =   1200
      TabIndex        =   15
      Top             =   75
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   105
      TabIndex        =   6
      Top             =   5115
      Width           =   795
   End
   Begin MSComctlLib.ImageList ilsIcon 
      Left            =   5565
      Top             =   -210
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
            Picture         =   "frmClientsUpgrade.frx":079F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "��������(&L)"
      Height          =   350
      Left            =   10140
      TabIndex        =   12
      Top             =   5115
      Width           =   1200
   End
   Begin VB.CheckBox chkAllSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��ǰȫ��վ������(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3480
      TabIndex        =   4
      Top             =   1230
      Width           =   2040
   End
   Begin VB.CommandButton cmdClearClients 
      Caption         =   "����3����δ��¼����վ"
      Height          =   350
      Left            =   9960
      TabIndex        =   23
      Top             =   1155
      Width           =   2400
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblUpResult 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����״̬"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5520
      TabIndex        =   19
      Top             =   1230
      Width           =   720
   End
   Begin VB.Label lblFind 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   1
      Top             =   1230
      Width           =   630
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   150
      Picture         =   "frmClientsUpgrade.frx":1269
      Top             =   615
      Width           =   480
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   2925
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "��վ�㲿�������������ú��ռ��ļ�����ϵͳ�����²�����Ϣ����ͨ��˫���ͻ��˲鿴���������"
      Height          =   348
      Left            =   828
      TabIndex        =   14
      Top             =   648
      Width           =   5112
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "վ�㲿������"
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
      TabIndex        =   13
      Top             =   105
      Width           =   1440
   End
End
Attribute VB_Name = "frmClientsUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const mMenu_Popu = 1
Private Const mMenu_Popu_ClientName = 11
Private Const mMenu_Popu_ClientIP = 12
Private Const mMenu_Popu_ClientDept = 13
Private Const mMenu_Popu_ClientUser = 14
Private mlngConnTimes       As Long
Private mcllAllConn         As New Collection '���еĿͻ�����Ϣ��WinSock�������Ӵ���
Private marrCurConn         As Variant      '��ǰ������Ϣ
Private Const M_EXPIRED_TIMES = 100
Dim mintColumn As Integer
Private mintType As Integer     '11-��վ�������й���,12-��IP����,13-�����Ź���,14-����;����
Private mrsClients As ADODB.Recordset
Private mrsFileServer As ADODB.Recordset
Private mrsFilePreUpgrade As ADODB.Recordset 'Ԥ������¼��
Private mblnChange As Boolean '�����˸ı�
Private mblnTypeChange As Boolean '������ʽ�����ı�
Private mintUpType     As Integer  '0 ������ʽ 1 FTP��ʽ'
Private mblnLoad       As Boolean '�Ƿ��Ѿ��������
Private Enum UpgradeState
    US_δ���� = 0
    US_�ɹ� = 1
    US_ʧ�� = 2
    US_������ = 3
    US_���� = 4
End Enum
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
End Sub

Private Sub cboFind_Click()
    Select Case cboFind.ListIndex
        Case 0
            mintType = 11
            txtSearch.Tag = "�����빤��վ����"
        Case 1
            mintType = 13
            txtSearch.Tag = "�����벿������"
        Case 2
            mintType = 12
            txtSearch.Tag = "������IP��ַ"
        Case 3
            mintType = 14
            txtSearch.Tag = "��������;"
    End Select
    txtSearch.Text = txtSearch.Tag
End Sub

Private Sub cboUpResult_Click()
    If mblnLoad Then
        LoadClientsInfor
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
        Dim objControl As CommandBarControl
        Dim objPopu As CommandBarPopup
        
        Select Case Control.id
        Case mMenu_Popu_ClientName 'վ������
            mintType = Control.id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        Case mMenu_Popu_ClientIP   'IP
            mintType = Control.id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        Case mMenu_Popu_ClientDept '��������
            mintType = Control.id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        Case mMenu_Popu_ClientUser  '��;
            mintType = Control.id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        End Select
End Sub

Private Sub chkAllSel_Click()
    Dim i As Long
    If chkAllSel.Tag = "T" Then chkAllSel.Tag = "": Exit Sub
    With vsClients
        .Cell(flexcpChecked, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = IIf(Me.chkAllSel.value = 1, flexChecked, flexUnchecked)
    End With
    mblnChange = True
    Call SetCtlEnabled
End Sub

Private Sub cmdClearClients_Click()
    Dim strSQL As String
    
    On Error GoTo errH
    If MsgBox("ʹ�ô˹��ܽ���ɾ��������������δ��¼�Ĺ���վ��" & vbCrLf & "ȷ��Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    strSQL = "Zl_Zlclients_Deletebatch()"
    ExecuteProcedure strSQL, Me.Caption
    Call cmdRefresh_Click
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdClearUpLog_Click()
    Dim strSQL As String
    If MsgBox("ȷ��Ҫ�������������־��?", vbYesNo + vbInformation, gstrSysName) = vbYes Then
        'ɾ��������־
        strSQL = "delete zltools.zlClientUpdatelog"
        gcnOracle.Execute strSQL
    End If
End Sub

Private Sub cmdClientModify_Click()
    Dim blnReturn   As Boolean
    Dim strIp       As String
    Dim strName     As String
    Dim lngRow      As Long
    
    With vsClients
        If .Row >= .FixedRows Then
            lngRow = .Row
            strIp = .TextMatrix(lngRow, .ColIndex("IP"))
            strName = .TextMatrix(lngRow, .ColIndex("����վ"))
            frmClientsEdit.ShowEdit strIp, strName, 1, blnReturn
            If Not blnReturn Then Exit Sub
            Call LoadClientsInfor(True)
            lngRow = .FindRow(strName, , .ColIndex("����վ"))
            If lngRow >= .FixedRows Then
                .SetFocus
                .Row = lngRow
                .ShowCell lngRow, .ColIndex("����վ")
            End If
        End If
    End With
End Sub

Private Sub cmdFile_Click()
    Dim blnReturn As Boolean
    If OptType(0).value Then
        Call frmFilesSet.ShowEdit(Me, blnReturn)
        If blnReturn = False Then Exit Sub
        '������������
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ like '������Ŀ¼%' or ��Ŀ like '�����û�%' or ��Ŀ like '��������%'"
        Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
        Call initVsGrid
        LoadClientsInfor (True)
    Else
        Call frmFilesFTPSet.ShowEdit(Me, blnReturn)
        If blnReturn = False Then Exit Sub
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ like 'FTP������%' or ��Ŀ like 'FTP�û�%' or ��Ŀ like 'FTP����%'"
        Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
        Call initVsGrid
        LoadClientsInfor (True)
    End If
End Sub

Private Sub cmdkillProcess_Click()
    frmKillProcessManage.ShowMe ("0307")
End Sub

Private Sub cmdRefresh_Click()
    '��ʼ����Ϣ
    Call LoadClientsInfor(True)
End Sub

Private Sub cmdUpdateS_Click()
    Dim lngRet As Long
    Dim i As Long
    Dim strName As String
    Dim strSQL As String
    
    lngRet = MsgBox("�µ�һ������ʱ,�����������ø�վ�������״̬Ϊ[δ����]" & vbNewLine & "ȷ��Ҫ����ѡ��վ�������״̬��?", vbYesNo + vbInformation, "��ʾ")
    If lngRet = vbYes Then
        With vsClients
            For i = .Row To .RowSel
                '���ݹ��ѡ������
                'If Val(.TextMatrix(i, .ColIndex("����"))) = -1 Then
                    strName = .TextMatrix(i, .ColIndex("����վ"))
                    strSQL = "Zl_Zlclients_Control(6,'" & strName & "')"
                    Call ExecuteProcedure(strSQL, Me.Caption)
                    
                    'ɾ��������־
                    strSQL = "delete zltools.zlClientUpdatelog where ����վ='" & UCase(strName) & "'"
                    gcnOracle.Execute strSQL
                'End If
            Next
            Call LoadClientsInfor(True)  'ˢ���б�
        End With
    End If
End Sub

Private Sub Cmd����_Click()
    If mblnChange Then
        If SaveData = False Then
            MsgBox "����վ������ʧ��!", vbInformation, gstrSysName
            Exit Sub
        Else
            MsgBox "����վ�����óɹ�!", vbInformation, gstrSysName
        End If
    End If
    
    If mblnTypeChange Then
        Call SaveUpType
        
        If mintUpType = 0 Then
            
            gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ like '������Ŀ¼%' or ��Ŀ like '�����û�%' or ��Ŀ like '��������%'"
            Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
            mrsFileServer.Filter = ""
            initVsGrid
            
        Else
            gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ like 'FTP������%' or ��Ŀ like 'FTP�û�%' or ��Ŀ like 'FTP����%'"
            Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
            mrsFileServer.Filter = ""
            
            initVsGrid
        End If
    End If
    Call LoadClientsInfor(mblnTypeChange Or mblnChange)
    mblnTypeChange = False
    mblnChange = False
    Call SetCtlEnabled
End Sub

Private Sub cmdӦ��_Click()
    
    Dim i As Long
    Dim strKey As String
    With vsClients
        
    
        If .Col = .ColIndex("������") Then
            .Redraw = flexRDNone
            strKey = Trim(.TextMatrix(.Row, .Col))
            For i = 1 To .Rows - 1
                .TextMatrix(i, .Col) = strKey
            Next
            .Redraw = flexRDBuffered
        End If
        
        If .Col = .ColIndex("Ԥ��ʱ��") Then
            .Redraw = flexRDNone
            strKey = Trim(.TextMatrix(.Row, .Col))
            For i = .Row To .RowSel
                .TextMatrix(i, .Col) = strKey
            Next
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("����")) = -1 Then
                    .TextMatrix(i, .Col) = strKey
                End If
            Next
            
            .Redraw = flexRDBuffered
        End If
        
        If .Col = .ColIndex("Ԥ�����") Then
            .Redraw = flexRDNone
            strKey = Trim(.TextMatrix(.Row, .Col))
            For i = .Row To .RowSel
                .TextMatrix(i, .Col) = strKey
                If strKey = "" Or strKey = "δ���" Then
                    .Cell(flexcpForeColor, i, .Col, i, .Col) = 0
                Else
                    .Cell(flexcpForeColor, i, .Col, i, .Col) = vbRed
                End If
            Next
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("����")) = -1 Then
                    .TextMatrix(i, .Col) = strKey
                    
                    If strKey = "" Or strKey = "δ���" Then
                        .Cell(flexcpForeColor, i, .Col, i, .Col) = 0
                    Else
                        .Cell(flexcpForeColor, i, .Col, i, .Col) = vbRed
                    End If
                End If
            Next
            .Redraw = flexRDBuffered
        End If
        If .Col = .ColIndex("�������") Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                If i <> .Row Then
                    .Cell(flexcpChecked, i, .Col) = .Cell(flexcpChecked, .Row, .Col)
                End If
            Next
            .Redraw = flexRDBuffered
        End If
    End With
End Sub

Private Sub cmd�û���������_Click()
    Load frmFilesUpgradeAdmin
    frmFilesUpgradeAdmin.Show 1, frmMDIMain
    If frmFilesUpgradeAdmin.mblnOK Then
    End If
    Exit Sub
End Sub

Private Sub cmdԤ��������_Click()
    Load frmFilesUpgradeTime
    frmFilesUpgradeTime.Show 1, frmMDIMain
    If frmFilesUpgradeTime.mblnOK Then
        '����Ԥ����ʱ������
        On Error GoTo errHandle
        Call ExecuteProcedure("Zl_Zlclients_Control(1)", Me.Caption)
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ = '�ͻ���Ԥ����ʱ���'"
        Set mrsFilePreUpgrade = New ADODB.Recordset
        Call OpenRecordset(mrsFilePreUpgrade, gstrSQL, Me.Caption)
    End If
    Call initVsGrid
    LoadClientsInfor (True)
    Exit Sub
errHandle:
    MsgBox "����ʧ�ܡ�" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    mblnLoad = False
    mintType = mMenu_Popu_ClientName
'    Call PrintSearch("������վ����", vbBlue, False)
'    txtSearch.Tag = "������վ����"
    '��ʼ������ʽ
    Call InitUpType
    
    'mintUpType =0 ������ʽ
    If mintUpType = 0 Then
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ like '������Ŀ¼%' or ��Ŀ like '�����û�%' or ��Ŀ like '��������%'"
    Else
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ like 'FTP������%' or ��Ŀ like 'FTP�û�%' or ��Ŀ like 'FTP����%'"
    End If
    Set mrsFileServer = New ADODB.Recordset
    Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
    
    
    gstrSQL = "Select ��Ŀ,���� From zlRegInfo where  ��Ŀ = '�ͻ���Ԥ����ʱ���'"
    Set mrsFilePreUpgrade = New ADODB.Recordset
    Call OpenRecordset(mrsFilePreUpgrade, gstrSQL, Me.Caption)
    '���ҹ��ܳ�ʼ��
    cboFind.AddItem "����վ", 0
    cboFind.AddItem "����", 1
    cboFind.AddItem "IP", 2
    cboFind.AddItem "��;", 3
    cboFind.ListIndex = 0
    
    cboUpResult.AddItem "δ����", US_δ����
    cboUpResult.AddItem "�ɹ�", US_�ɹ�
    cboUpResult.AddItem "ʧ��", US_ʧ��
    cboUpResult.AddItem "������", US_������
    cboUpResult.AddItem "����", US_����
    cboUpResult.ListIndex = US_����
    
    txtSearch.ForeColor = vbGrayText
    '��ʼ�˵�
    Call InitCommandBar
    '��ʼ��������������
    Call initVsGrid
    mblnLoad = True
    '��ʼ����Ϣ
    Call LoadClientsInfor(True)
    Call RestoreGridSet

    mblnChange = False
End Sub

Private Sub RestoreGridSet()
    '---------------------------------------------------------------------------------
    '����:�ָ����Ի�����
    '����:���˺�
    '����:2007/09/10
    '---------------------------------------------------------------------------------
    Dim i As Long
    Dim strColumns As String
    Dim arrColumn As Variant
    Dim arrValue As Variant
    err = 0: On Error GoTo ErrHand:
    '�ָ����Ի�����
    strColumns = ""
    strColumns = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Caption, "վ��", "")
    
    If strColumns <> "" Then
        arrColumn = Split(strColumns, "|")
        '�з����仯���򲻸��Ի�
        If UBound(arrColumn) <> vsClients.Cols - 1 Then Exit Sub
        With vsClients
            For i = 0 To UBound(arrColumn)
                arrValue = Split(arrColumn(i), ",")
                .ColWidth(.ColIndex(arrValue(0))) = Val(arrValue(1))
                .ColPosition(.ColIndex(arrValue(0))) = i
            Next
        End With
    End If
ErrHand:
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lblNote.Width = Me.ScaleWidth - Me.lblNote.Left
    
    With cmdClearClients
         .Left = ScaleWidth - .Width - 50
    End With
    
    With cmdUpdateS
         .Left = cmdClearClients.Left - .Width - 120
    End With
    
    With cmdClearUpLog
         .Left = cmdUpdateS.Left - .Width - 5
    End With
    
    Call SetCtrlPosOnLine(False, 0, chkAllSel, 30, lblUpResult, 15, cboUpResult)
    cmdRefresh.Top = Me.ScaleHeight - cmdRefresh.Height - 50
    vsClients.Height = cmdRefresh.Top - vsClients.Top - 90
    vsClients.Width = Me.ScaleWidth - vsClients.Left - 90
    picSel.Left = vsClients.Left
    picSel.Top = vsClients.Top
    Call SetCtrlPosOnLine(False, 0, cmdRefresh, 45, cmdӦ��, 45, cmd����, 60, OptType(0), 45, OptType(1))
    cmdFile.Top = cmdRefresh.Top
    cmdFile.Left = Me.ScaleWidth - cmdFile.Width - 120
    
    Call SetCtrlPosOnLine(False, 0, cmdFile, (45 + cmdԤ��������.Width + cmdFile.Width) * -1, cmdԤ��������, (45 + cmd�û���������.Width + cmdԤ��������.Width) * -1, cmd�û���������, (45 + cmd�û���������.Width + cmdClientModify.Width) * -1, cmdClientModify, (700 + cmdClientModify.Width + cmdkillProcess.Width) * -1, cmdkillProcess)
End Sub

Private Sub SetCtlEnabled()
    '---------------------------------------------------------------------------------------------
    '���ܣ����ÿؼ����������
    '������
    '���أ�
    '���ƣ����˺�
    '���ڣ�2007/09/07
    '---------------------------------------------------------------------------------------------
    
    Dim blnNoClients As Boolean 'û��վ��
    Dim i As Long, blnӦ�� As Boolean
    blnNoClients = True
    With vsClients
        blnӦ�� = (.Col = .ColIndex("������")) Or (.Col = .ColIndex("Ԥ��ʱ��")) Or (.Col = .ColIndex("Ԥ�����") Or (.Col = .ColIndex("�������")))
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("����վ"))) <> "" Then
                blnNoClients = False
                Exit For
            End If
        Next
    End With
    chkAllSel.Enabled = Not blnNoClients
    cmdӦ��.Enabled = blnӦ��
    cmd����.Enabled = mblnChange
End Sub

Private Sub LoadClientsInfor(Optional blnRefresh As Boolean)
    '---------------------------------------------------------------------------------------------
    '���ܣ�����վ����Ϣ
    '������blnFilter-�Ƿ�ͨ�����еļ�¼���й���
    '���أ�
    '���ƣ����˺�
    '����:2007/08/20
    '---------------------------------------------------------------------------------------------
    Dim strSQL As String, strFilter As String, strKey As String
    Dim i As Long
    Dim StrDate As String
    Dim lngColore   As Long
    
    err = 0: On Error GoTo ErrHand:
    
    strSQL = "Select a.����վ, a.Ip, a.����, zlSpellCode(a.����) As ���ż���, a.��;, a.Cpu, a.�ڴ�, a.Ӳ��, a.����ϵͳ, a.����������, a.Ftp������," & vbNewLine & _
            "       To_Char(a.Ԥ��ʱ��, 'hh24:mi') Ԥ��ʱ��, a.������־, a.�ռ���־," & vbNewLine & _
            "       Decode(a.�������, 0, 'δ����', 1, '���', 2, 'ʧ��', 3, '��������', Decode(Nvl(a.������־, 0), 0, Null, 'δ����')) �������," & vbNewLine & _
            "       Decode(a.Ԥ�����, 0, 'δ����', 1, '���', 2, 'ʧ��', 3, '��������', Decode(Nvl(a.������־, 0), 0, Null, 'δ����')) Ԥ�����," & vbNewLine & _
            "       Decode(a.�ռ�״̬, 0, 'δ���', 1, '��������', 2, '�����쳣', 3, '���ڼ��', Decode(Nvl(a.�ռ���־, 0), 0, Null, 'δ���')) ������," & vbNewLine & _
            "       a.������� As ����״̬, a.Ԥ�����, a.�ռ�״̬, a.˵��, a.��ֹʹ��, a.������, a.����Ա�û�, a.����Ա����, Decode(c.Terminal, Null, 0, 1) As ״̬, d.����ֵ �˿�" & vbNewLine & _
            "From zlClients A, (Select Distinct Terminal From GV$session) C," & vbNewLine & _
            "     (Select ������, Nvl(b.����ֵ, 0) ����ֵ" & vbNewLine & _
            "       From zlParameters A, zlUserParas B" & vbNewLine & _
            "       Where a.������ = '����Զ�̿���' And Nvl(a.ģ��, 0) = 0 And Nvl(a.ϵͳ, 0) = 0 And a.Id = b.����id) D" & vbNewLine & _
            "Where Upper(a.����վ) = Upper(c.Terminal(+)) And a.����վ = d.������(+)" & vbNewLine & _
            "Order By Ip"
    If blnRefresh = True Or mrsClients Is Nothing Then
        Set mrsClients = New ADODB.Recordset
        Call OpenRecordset(mrsClients, strSQL, Me.Caption)
        'Set rsClients = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client", "")
    ElseIf mrsClients.State <> 1 Then
        Call OpenRecordset(mrsClients, strSQL, Me.Caption)
    End If
    strKey = txtSearch.Text
    If strKey <> "" And strKey <> txtSearch.Tag Or cboUpResult.ListIndex <> US_���� Then
        If strKey <> "" And strKey <> txtSearch.Tag Then
            Select Case mintType
                Case 12     '-��IP����
                    strFilter = "IP like '" & strKey & "%'"
                Case 13     '-�����Ź���
                    strFilter = "���� like '" & strKey & "%' OR ���ż��� like '" & UCase(strKey) & "%'"
                Case 14     '����;����
                    strFilter = "��; like '" & strKey & "%'"
                Case Else           ' 11-��վ�������й���
                    strFilter = "����վ like '" & UCase(strKey) & "%'"
            End Select
        End If
        
        If cboUpResult.ListIndex <> US_���� Then
            If strFilter <> "" Then
                If mintType = 13 Then
                    strFilter = "(���� like '" & strKey & "%' And ����״̬=" & cboUpResult.ListIndex & ") OR (���ż��� like '" & UCase(strKey) & "%' And ����״̬=" & cboUpResult.ListIndex & ")"
                Else
                    strFilter = strFilter & " And ����״̬=" & cboUpResult.ListIndex
                End If
            Else
                strFilter = "����״̬=" & cboUpResult.ListIndex
            End If
        End If
        mrsClients.Filter = strFilter
    Else
        mrsClients.Filter = 0
    End If
    
    With vsClients
        .Redraw = flexRDNone
        .Rows = IIf(mrsClients.RecordCount = 0, 1, mrsClients.RecordCount) + 1
        If mrsClients.RecordCount <> 0 Then mrsClients.MoveFirst
        If mrsClients.EOF Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
            .Redraw = flexRDBuffered
            SetCtlEnabled
            mrsClients.Filter = 0
            Exit Sub
        End If
        '������
        i = 1
        Do While Not mrsClients.EOF
            lngColore = 0
            If mintUpType = 0 Then
                strKey = Val(Nvl(mrsClients!����������))
            Else
                strKey = Val(Nvl(mrsClients!FTP������))
            End If
'            strKey = IIf(Val(strKey) = 0, "", strKey)
            If mintUpType = 0 Then
                mrsFileServer.Find "��Ŀ='������Ŀ¼" & strKey & "'", , , 1
            Else
                If strKey = "" Then strKey = "0"
                mrsFileServer.Find "��Ŀ='FTP������" & strKey & "'", , , 1
            End If
            If mrsFileServer.EOF = False Then
                .TextMatrix(i, .ColIndex("������")) = Val(strKey) & ":" & Nvl(mrsFileServer!����)
            Else
                .TextMatrix(i, .ColIndex("������")) = Val(strKey) & ":"
            End If
            .Cell(flexcpData, i, .ColIndex("������")) = Val(strKey)
            
            If mrsFilePreUpgrade.EOF = False Then
                .TextMatrix(i, .ColIndex("Ԥ��ʱ��")) = Nvl(mrsClients!Ԥ��ʱ��)
            Else
                .TextMatrix(i, .ColIndex("Ԥ��ʱ��")) = ""
            End If
            .TextMatrix(i, .ColIndex("����վ")) = Nvl(mrsClients!����վ)
            .TextMatrix(i, .ColIndex("IP")) = Nvl(mrsClients!IP)
            .TextMatrix(i, .ColIndex("CPU")) = Nvl(mrsClients!cpu)
            .TextMatrix(i, .ColIndex("�ڴ�")) = Nvl(mrsClients!�ڴ�)
            .TextMatrix(i, .ColIndex("Ӳ��")) = Nvl(mrsClients!Ӳ��)
            .TextMatrix(i, .ColIndex("����ϵͳ")) = Nvl(mrsClients!����ϵͳ)
            .TextMatrix(i, .ColIndex("����")) = Nvl(mrsClients!����)
            .TextMatrix(i, .ColIndex("��;")) = Nvl(mrsClients!��;)
            .TextMatrix(i, .ColIndex("�������")) = Nvl(mrsClients!�������)
            .TextMatrix(i, .ColIndex("�����")) = Nvl(mrsClients!������)
            .TextMatrix(i, .ColIndex("����Ա")) = Nvl(mrsClients!����Ա�û�)
            .TextMatrix(i, .ColIndex("����")) = Decipher(Nvl(mrsClients!����Ա����))
            .TextMatrix(i, .ColIndex("״̬")) = Nvl(mrsClients!״̬)
            .TextMatrix(i, .ColIndex("�˿�")) = Nvl(mrsClients!�˿�)
            If Nvl(mrsClients!����״̬, 0) = 3 Then
                lngColore = vbGreen '��ɫ
            ElseIf Nvl(mrsClients!����״̬, 0) = 2 Then
                lngColore = vbRed '��ɫ
            ElseIf Nvl(mrsClients!����״̬, 0) = 1 Then
                lngColore = vbBlue '��ɫ
            End If
            .Cell(flexcpForeColor, i, .ColIndex("����վ"), i, .ColIndex("IP")) = lngColore
            'ʹ����ɫ��ʶԤ���Ƿ����!
            If Nvl(mrsClients!Ԥ�����, 0) = 1 Then
                .TextMatrix(i, .ColIndex("Ԥ�����")) = "���"
                .Cell(flexcpForeColor, i, .ColIndex("Ԥ�����"), i, .ColIndex("Ԥ�����")) = vbRed
            Else
                .TextMatrix(i, .ColIndex("Ԥ�����")) = "δ���"
                .Cell(flexcpForeColor, i, .ColIndex("Ԥ�����"), i, .ColIndex("Ԥ�����")) = 0
            End If
            
            .TextMatrix(i, .ColIndex("˵��")) = Nvl(mrsClients!˵��)
            If Val(Nvl(mrsClients!������־)) = 1 Then
                .Cell(flexcpChecked, i, .ColIndex("����")) = flexChecked
            Else
                .Cell(flexcpChecked, i, .ColIndex("����")) = flexUnchecked
            End If
            
            If Val(Nvl(mrsClients!�ռ���־)) = 1 Then
                .Cell(flexcpChecked, i, .ColIndex("�������")) = flexChecked
            Else
                .Cell(flexcpChecked, i, .ColIndex("�������")) = flexUnchecked
            End If
            i = i + 1
            mrsClients.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    mrsClients.Filter = 0
    SetCtlEnabled
    Exit Sub
ErrHand:
   ' Resume
    vsClients.Redraw = flexRDBuffered
    MsgBox "ϵͳ���ִ���,����Ϊ:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
    SetCtlEnabled
    Exit Sub
End Sub

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------
    '����:����������Ϣ
    '����:str����վ-����վ
    '     bln������־
    '     str��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2007/09/07
    '---------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim i As Long
    Dim str����վ As String, int���� As Integer, str�������� As String
    Dim strԤ��ʱ�� As String
    Dim intԤ����� As Integer, str������� As String
    Dim strIp As String
    Dim strIPPort As String, strJobs As String
    Dim blnStart    As Boolean
    err = 0: On Error GoTo ErrHand:
    
    blnStart = mcllAllConn.Count = 0
    timerConnect.Enabled = False '����ͣ
    With vsClients
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("����վ"))) <> "" Then
                str����վ = Trim(.TextMatrix(i, .ColIndex("����վ")))
                int���� = IIf(.Cell(flexcpChecked, i, .ColIndex("����")) = flexChecked, 1, 0)
                strIp = Trim(.TextMatrix(i, .ColIndex("IP")))
                If Trim(.TextMatrix(i, .ColIndex("������"))) = "" Then
                    str�������� = "0"
                Else
                    str�������� = Val(Split(Trim(.TextMatrix(i, .ColIndex("������"))), ":")(0))
                End If
                
                If Trim(.TextMatrix(i, .ColIndex("Ԥ��ʱ��"))) = "" Then
                    strԤ��ʱ�� = "NULL"
                Else
                    strԤ��ʱ�� = Trim(.TextMatrix(i, .ColIndex("Ԥ��ʱ��")))
                    strԤ��ʱ�� = Format(Now(), "yyyy-MM-dd") & " " & Format(strԤ��ʱ��, "hh:mm:00")
                    strԤ��ʱ�� = "to_date('" & strԤ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                
                If Trim(.TextMatrix(i, .ColIndex("Ԥ�����"))) = "" Then
                    intԤ����� = 0
                Else
                    If Trim(.TextMatrix(i, .ColIndex("Ԥ�����"))) = "δ���" Then
                        intԤ����� = 0
                    Else
                        intԤ����� = 1
                    End If
                End If
                str������� = IIf(.Cell(flexcpChecked, i, .ColIndex("�������")) = flexChecked, 1, "NULL")
                
                If .TextMatrix(i, .ColIndex("״̬")) = "1" Then
                    strIPPort = .TextMatrix(i, .ColIndex("IP"))
                    If Val(.TextMatrix(i, .ColIndex("IP"))) = 0 Then
                        strIPPort = strIPPort & ":1001"
                    ElseIf Val(.TextMatrix(i, .ColIndex("IP"))) = -1 Then
                        strIPPort = ""
                    Else
                        strIPPort = strIPPort & ":" & Val(.TextMatrix(i, .ColIndex("IP")))
                    End If
                    strJobs = IIf(int���� = 0, "", 1) & "," & IIf(str������� = "1", str�������, "")
                    If strJobs = "," Then strJobs = ""
                Else
                    strIPPort = ""
                    strJobs = ""
                End If
                If mintUpType = 0 Then
                    strSQL = "Zl_Zlclients_Control(2,Null,'" & strIp & "'," & int���� & "," & str�������� & "," & strԤ��ʱ�� & "," & intԤ����� & ",NULL," & str������� & ")"
                Else
                    strSQL = "Zl_Zlclients_Control(2,Null,'" & strIp & "'," & int���� & ",Null," & strԤ��ʱ�� & "," & intԤ����� & "," & str�������� & "," & str������� & ")"
                End If
                Call ExecuteProcedure(strSQL, Me.Caption)
                Call AddClientsJob(strIPPort, strJobs)
            End If
        Next
    End With
    If mcllAllConn.Count <> 0 Then
        timerConnect.Enabled = True
        If blnStart Then mlngConnTimes = 0
    End If
    SaveData = True
    Exit Function
ErrHand:
    MsgBox "����������Ϣʱ����,������Ϣ����:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description, vbInformation, gstrSysName
'''    Resume
End Function

Private Sub Form_Unload(Cancel As Integer)
    '������Ի�����
    Dim i As Long
    Dim strColumns As String
    strColumns = ""
    With vsClients
        For i = 0 To .Cols - 1
            strColumns = strColumns & "|" & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    If strColumns <> "" Then strColumns = Mid(strColumns, 2)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Caption, "վ��", strColumns
End Sub

Private Sub OptType_Click(Index As Integer)
    mblnTypeChange = True
    cmd����.Enabled = True
    mintUpType = Index
End Sub

Private Sub picSel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picSel.Tag = "In" Then
        If x < 0 Or y < 0 Or x > picSel.Width Or y > picSel.Height Then
            ReleaseCapture
            picSel.Tag = ""
            PrintSearch Me.txtSearch.Tag, vbBlue, False
        End If
    Else
        picSel.Tag = "In"
        SetCapture picSel.hwnd
        MousePointer = 99
        PrintSearch Me.txtSearch.Tag, vbRed, True
    End If
End Sub

Private Sub picSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsMain.FindControl(xtpControlPopup, mMenu_Popu, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    
    Call PrintSearch(Me.txtSearch.Tag, vbBlue, False)
    picSel.Tag = ""
End Sub

Private Sub PrintSearch(ByVal strTittle As String, ByVal lngColor As Long, ByVal blnBoderStyle As Boolean)
    '----------------------------------------------------------------------------------------
    '����:��ӡָ������������
    '����:strTittle-����
    '     lngColor-��ɫֵ
    '     lngBoderStyl-�Ƿ�ӱ߿���
    '----------------------------------------------------------------------------------------
 
    With picSel
        .Cls
        .ForeColor = lngColor
        .FontUnderline = True
        .CurrentX = 30 '(.ScaleWidth - .TextWidth(strTittle))
        .CurrentY = (.ScaleHeight - .TextHeight(strTittle)) / 2
        picSel.Print strTittle
        .ZOrder 1
    End With
End Sub

Private Sub timerConnect_Timer()
    DoEvents
    mlngConnTimes = mlngConnTimes + 1
    If mlngConnTimes >= M_EXPIRED_TIMES Then
        If winSock.State <> sckClosed Then winSock.Close
        Call AddClientsJob(marrCurConn(0) & ":" & marrCurConn(1), "")
        mlngConnTimes = 1
    End If
    If mlngConnTimes = 1 And winSock.State = sckClosed And mcllAllConn.Count <> 0 Then '��ʼ������Ϣ
        marrCurConn = Split(mcllAllConn(1), ":")
        winSock.RemoteHost = marrCurConn(0)
        winSock.RemotePort = marrCurConn(1)
        winSock.Connect
    ElseIf mcllAllConn.Count = 0 Then
        timerConnect.Enabled = False
    End If
End Sub

Private Sub txtSearch_Change()
    If mblnLoad Then
        If txtSearch.Text = txtSearch.Tag Then Exit Sub
        If mblnChange = True Then
            If MsgBox("վ��������Ϣ����༭��,�Ƿ񱣴���༭����Ϣ?", vbQuestion + vbYesNo + vbQuestion) = vbYes Then
                Call SaveData
            End If
            mblnChange = False
        End If
        LoadClientsInfor
    End If
End Sub

Private Sub txtSearch_GotFocus()
    If txtSearch.ForeColor = vbGrayText Then
        txtSearch.Text = ""
        txtSearch.ForeColor = vbBlack
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub InitCommandBar()
    '-------------------------------------------------------------------------------------------
    '����:��ʼ���˵�
    '����:
    '����:
    '����:���˺�
    '����:2007/08/07
    '-------------------------------------------------------------------------------------------
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objDeptBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�����˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mMenu_Popu, "�����˵�(&P)", -1, False)
    objMenu.id = mMenu_Popu
    objMenu.Visible = False
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientName, "������վ����(&0)"): objControl.id = mMenu_Popu_ClientName: objControl.IconId = 102: objControl.Checked = True
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientIP, "    ��IP����(&1)"): objControl.id = mMenu_Popu_ClientIP: objControl.IconId = 102
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientDept, "  ����������(&2)"): objControl.id = mMenu_Popu_ClientDept: objControl.IconId = 102
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientUser, "  ����;����(&3)"): objControl.id = mMenu_Popu_ClientUser: objControl.IconId = 102
    End With
 End Sub
 Private Sub initVsGrid()
    '----------------------------------------------------------------------------------------
    '����:��ʼ��վ��������������
    '----------------------------------------------------------------------------------------
    With vsClients
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("������")) = Get������
        .ColComboList(.ColIndex("Ԥ��ʱ��")) = GetԤ��ʱ��
        If .ColComboList(.ColIndex("Ԥ��ʱ��")) = "" Then
            .ColHidden(.ColIndex("Ԥ��ʱ��")) = True
        Else
            .ColHidden(.ColIndex("Ԥ��ʱ��")) = False
        End If
        .ColComboList(.ColIndex("Ԥ�����")) = GetԤ�����
    End With
 End Sub
 
 Private Function Get������() As String
    Dim strCombox As String
    Dim strTemp As String
    strCombox = ""
    With mrsFileServer
        If mintUpType = 0 Then
            .Filter = "��Ŀ like '������Ŀ¼%'"
            Do While Not .EOF
                strTemp = Replace(Nvl(!��Ŀ), "������Ŀ¼", "")
                strCombox = strCombox & "|" & Val(strTemp) & ":" & Nvl(!����)
                .MoveNext
            Loop
        Else
            .Filter = "��Ŀ like 'FTP������%'"
            Do While Not .EOF
                strTemp = Replace(Nvl(!��Ŀ), "FTP������", "")
                strCombox = strCombox & "|" & Val(strTemp) & ":" & Nvl(!����)
                .MoveNext
            Loop
        End If

    End With
    If strCombox <> "" Then strCombox = Mid(strCombox, 2)
    Get������ = strCombox
 End Function
 
 Private Function GetԤ��ʱ��() As String
    Dim strTemp As String
    If mrsFilePreUpgrade.RecordCount = 1 Then
        mrsFilePreUpgrade.MoveFirst
        strTemp = Replace(Nvl(mrsFilePreUpgrade!����), ",", "|")
    Else
        strTemp = ""
    End If
    
    GetԤ��ʱ�� = strTemp
 End Function
 
 Private Function GetԤ�����() As String
    GetԤ����� = "δ���|���"
 End Function
 
Private Sub txtSearch_LostFocus()
    If txtSearch.Text = "" Then
        txtSearch.Text = txtSearch.Tag
        txtSearch.ForeColor = vbGrayText
    End If
End Sub

Private Sub vsClients_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
      With vsClients
        Select Case Col
        Case vsClients.ColIndex("������")
            mblnChange = True
            Call SetCtlEnabled
        Case vsClients.ColIndex("Ԥ��ʱ��")
            mblnChange = True
            Call SetCtlEnabled
        Case vsClients.ColIndex("Ԥ�����")
            mblnChange = True
            If vsClients.TextMatrix(Row, Col) = "" Or vsClients.TextMatrix(Row, Col) = "δ���" Then
              vsClients.Cell(flexcpForeColor, Row, Col, Row, Col) = 0
            Else
              vsClients.Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
            Call SetCtlEnabled
        Case vsClients.ColIndex("����")
            chkAllSel.Tag = "T"
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("����")) = flexChecked Then
'                    MsgBox "��" & i & "��"
                    Exit For
                End If
            Next
            If i = .Rows Then
                chkAllSel.value = 0
            Else
                chkAllSel.value = 2
            End If
            mblnChange = True
            Call SetCtlEnabled
        Case vsClients.ColIndex("�������")
            mblnChange = True
            Call SetCtlEnabled
        Case Else
        End Select
    End With
End Sub

Private Sub vsClients_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case vsClients.ColIndex("������"), vsClients.ColIndex("����"), vsClients.ColIndex("Ԥ��ʱ��"), vsClients.ColIndex("Ԥ�����"), vsClients.ColIndex("�������")
        'ֻ�з������к������в��ܸ���
    Case Else
        '�����в��ܸ���
        Cancel = True
    End Select
End Sub
  
Private Sub vsClients_DblClick()
    '�鿴�������
    Dim strName As String, intType As Integer
    If vsClients.Row > 0 Then
        strName = vsClients.TextMatrix(vsClients.Row, vsClients.ColIndex("����վ"))
        intType = -1
        If vsClients.Col = vsClients.ColIndex("�������") And vsClients.TextMatrix(vsClients.Row, vsClients.ColIndex("�������")) = "������" Then
            intType = 1
        ElseIf vsClients.TextMatrix(vsClients.Row, vsClients.ColIndex("�������")) <> "δ����" And vsClients.Col = vsClients.ColIndex("�������") Then
            intType = 0
        ElseIf vsClients.TextMatrix(vsClients.Row, vsClients.ColIndex("�������")) <> "δ����" Then
            intType = 0
        ElseIf vsClients.TextMatrix(vsClients.Row, vsClients.ColIndex("�������")) = "������" Then
            intType = 1
        End If
        If intType <> -1 Then
            Call frmFilesUpgradeLogView.ShowMe(strName, 0)
        End If
    End If
    
    Exit Sub
errHandle:
        MsgBox "����ʧ�ܡ�" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub vsClients_RowColChange()
    Call SetCtlEnabled
End Sub

Private Sub winSock_Connect()
    winSock.SendData "CLIENT_JOB:" & marrCurConn(2)
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    winSock.GetData strData
    '��������Ϣ���Լ���Ӧ��״̬
    If strData = "MESSAGE:" & marrCurConn(2) & ",STATE:1" Then
        mlngConnTimes = M_EXPIRED_TIMES '��־�Ѿ�����
    End If
End Sub

Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number <> 0 Then
        mlngConnTimes = M_EXPIRED_TIMES '��־�Ѿ�����
    End If
End Sub

Private Sub SaveUpType()
'----------------------------------------------------------------------------------------
'����:�޸�������ʽ��Ϣ
'----------------------------------------------------------------------------------------
    On Error GoTo errH
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim str��Ŀ As String '��Ŀ
    Dim str���� As String '����
    Dim strSQLTemp As String
    str��Ŀ = "��������"
    If OptType(0).value Then
        str���� = "0"
    Else
        str���� = "1"
    End If
    strSQL = " Select ��Ŀ,���� From zlregInfo where ��Ŀ= '��������'"
    
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp.EOF = True Then
'        gcnOracle.BeginTrans
        strSQLTemp = "insert into zlregInfo(��Ŀ,����) values ('" & str��Ŀ & "','" & str���� & "')"
        gcnOracle.Execute strSQLTemp
'        gcnOracle.CommitTrans
    Else
'        gcnOracle.BeginTrans
        strSQLTemp = "delete zlRegInfo where ��Ŀ='" & str��Ŀ & "'"
        gcnOracle.Execute strSQLTemp
        strSQLTemp = "insert into zlregInfo(��Ŀ,����) values ('" & str��Ŀ & "','" & str���� & "')"
        gcnOracle.Execute strSQLTemp
'        gcnOracle.CommitTrans
    End If
    
    Exit Sub
errH:
    If err Then
        MsgBox "��������������Ϣʱ����,������Ϣ����:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub InitUpType()
'----------------------------------------------------------------------------------------
'����:��ʼ������ʽ��Ϣ
'----------------------------------------------------------------------------------------
    On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    strSQL = " Select ��Ŀ,���� From zlregInfo where ��Ŀ= '��������'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)

    If rsTemp.EOF = False Then
        strTemp = Nvl(rsTemp!����, "0")
        If strTemp = "1" Then
             OptType(1).value = True
             mintUpType = 1
        Else
             OptType(0).value = True
             mintUpType = 0
        End If
    Else
        OptType(0).value = True
        mintUpType = 0
    End If
    Exit Sub
errH:
    If err Then
        MsgBox "��ʼ��������ʽ����,������Ϣ����:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub AddClientsJob(ByVal strIPPortInfo As String, ByVal strJobs As String)

    If strIPPortInfo = "" Then Exit Sub
    On Error Resume Next
    If strJobs = "" Then
        mcllAllConn.Remove "K_" & strIPPortInfo
        If err.Number <> 0 Then err.Clear
        Exit Sub
    End If
    mcllAllConn.Add strIPPortInfo & ":" & strJobs, "K_" & strIPPortInfo
    If err.Number <> 0 Then err.Clear
End Sub
