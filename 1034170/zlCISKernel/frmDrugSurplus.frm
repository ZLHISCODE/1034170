VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugSurplus 
   AutoRedraw      =   -1  'True
   Caption         =   "ҩƷ����Ǽ�"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10830
   Icon            =   "frmDrugSurplus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic���� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   195
      ScaleHeight     =   2865
      ScaleWidth      =   2415
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2415
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmDrugSurplus.frx":058A
         Left            =   1380
         List            =   "frmDrugSurplus.frx":0597
         TabIndex        =   10
         Text            =   "50%"
         Top             =   1907
         Width           =   765
      End
      Begin VB.OptionButton optRule 
         Caption         =   "��������"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   9
         Top             =   1960
         Width           =   1380
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "��ȡ����ҩƷ(&L)"
         Height          =   350
         Left            =   870
         TabIndex        =   12
         Top             =   2535
         Width           =   1560
      End
      Begin VB.OptionButton optRule 
         Caption         =   "�����������"
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   11
         Top             =   2250
         Width           =   1380
      End
      Begin VB.OptionButton optRule 
         Caption         =   "ȫ������"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   8
         Top             =   1670
         Width           =   1380
      End
      Begin VB.OptionButton optRule 
         Caption         =   "������"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   7
         Top             =   1380
         Value           =   -1  'True
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   345
         TabIndex        =   5
         Top             =   255
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   102498307
         CurrentDate     =   39610
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   345
         TabIndex        =   6
         Top             =   615
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   102498307
         CurrentDate     =   39610
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ������㣺"
         Height          =   180
         Left            =   105
         TabIndex        =   20
         Top             =   1095
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����ʱ�䣺"
         Height          =   180
         Left            =   105
         TabIndex        =   19
         Top             =   15
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   105
         TabIndex        =   17
         Top             =   675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   105
         TabIndex        =   16
         Top             =   315
         Width           =   180
      End
   End
   Begin VB.PictureBox picWay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   300
      ScaleHeight     =   3330
      ScaleWidth      =   2415
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1260
      Width           =   2415
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   1335
         TabIndex        =   4
         ToolTipText     =   "Ctrl+R"
         Top             =   1560
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   255
         TabIndex        =   3
         ToolTipText     =   "Ctrl+A"
         Top             =   1560
         Width           =   1100
      End
      Begin MSComctlLib.ListView lvwWay 
         Height          =   1410
         Left            =   30
         TabIndex        =   2
         Top             =   45
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   2487
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��ҩ;��"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame fraLR 
      BorderStyle     =   0  'None
      Height          =   6345
      Left            =   3090
      MousePointer    =   9  'Size W E
      TabIndex        =   23
      Top             =   600
      Width           =   45
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   195
      ScaleHeight     =   990
      ScaleWidth      =   2415
      TabIndex        =   18
      Top             =   6150
      Width           =   2415
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   1335
         TabIndex        =   14
         ToolTipText     =   "F3"
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   105
         TabIndex        =   13
         Top             =   285
         Width           =   2310
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ���롢���롢���ƣ�"
         Height          =   180
         Left            =   105
         TabIndex        =   21
         Top             =   60
         Width           =   1980
      End
   End
   Begin VB.ComboBox cboҩ�� 
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   870
      Width           =   2100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDrug 
      Height          =   6300
      Left            =   3195
      TabIndex        =   0
      Top             =   600
      Width           =   7560
      _cx             =   13335
      _cy             =   11112
      Appearance      =   2
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugSurplus.frx":05AA
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   115
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
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   6600
      Left            =   90
      TabIndex        =   22
      Top             =   645
      Width           =   3000
      _Version        =   589884
      _ExtentX        =   5292
      _ExtentY        =   11642
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   405
      Top             =   210
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDrugSurplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long

Private mrsDrug As ADODB.Recordset '����ҩ����
Private mrsApply As ADODB.Recordset '��ҩ�������
Private mrsAdvice As ADODB.Recordset 'ҽ����Ϣ����

Private mblnReturn As Boolean
Private mlngPeҩ�� As Long
Private mstrLike As String
Private mint���� As Integer
Private mblnChange As Boolean
Private Enum COL_DRUG
    col���� = 0
    colҩƷ = 1
    col��� = 2
    col���� = 3
    col��λ = 4
    colӦ���� = 5
    col������ҩ�� = 6
    col�����ҩ�� = 7
    col������ = 8
    col��� = 9
    colסԺ��װ = 10
End Enum

Public Sub ShowMe(frmParent As Object, ByVal lng����ID As Long)
    mlng����ID = lng����ID
    
    On Error Resume Next
    Me.Show , frmParent
End Sub

Private Sub cbo����_GotFocus()
    Call zlControl.TxtSelAll(cbo����)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cbo����_Validate(blnCancel)
        If Not blnCancel Then Call cmdLoad_Click
    Else
        If InStr("0123456789%" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    If Val(cbo����.Text) < 0 Or Val(cbo����.Text) > 100 Then
        Cancel = True
    Else
        cbo����.Text = Val(cbo����.Text) & "%"
    End If
End Sub

Private Sub cboҩ��_Click()
    If cboҩ��.ListIndex <> -1 Then
        If cboҩ��.ListIndex = mlngPeҩ�� Then Exit Sub
        If mblnChange Then
            If MsgBox("��ǰ������û�б��棬ȷʵҪ�л�ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Call zlControl.CboSetIndex(cboҩ��.hWnd, mlngPeҩ��)
                Exit Sub
            End If
        End If
        If Not CheckDate Then
            Call zlControl.CboSetIndex(cboҩ��.hWnd, mlngPeҩ��)
            dtpBegin.SetFocus: Exit Sub
        End If
        
        mlngPeҩ�� = cboҩ��.ListIndex
        Call ReleaseRecord(mrsDrug)
        Call ReleaseRecord(mrsAdvice)
        Call ReleaseRecord(mrsApply)
        
        Call LoadSurplus
        
        mblnChange = False
        If Me.Visible Then vsDrug.SetFocus
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_Edit_Save
        Call SaveData
    Case conMenu_File_Print
        Call OutputList(1)
    Case conMenu_File_Preview
        Call OutputList(2)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    Me.tkpMain.Left = lngLeft
    Me.tkpMain.Top = lngTop
    Me.tkpMain.Height = lngBottom - lngTop
    
    Me.fraLR.Left = lngLeft + tkpMain.Width
    Me.fraLR.Top = lngTop
    Me.fraLR.Height = lngBottom - lngTop

    Me.vsDrug.Left = fraLR.Left + fraLR.Width
    Me.vsDrug.Top = lngTop
    Me.vsDrug.Width = lngRight - lngLeft - fraLR.Width - tkpMain.Width
    Me.vsDrug.Height = lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Save
        Control.Enabled = mblnChange
    End Select
End Sub

Private Sub cmdClear_Click()
    Call SelectLVW(lvwWay, False)
    Call ReleaseRecord(mrsDrug)
End Sub

Private Sub cmdFind_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInput As String, strMatch As String
    Dim strSQL As String, i As Long
    Dim blnFirst As Boolean
    
    If vsDrug.Rows = vsDrug.FixedRows + 1 And vsDrug.RowData(vsDrug.Row) = 0 Then
        MsgBox "�Ҳ���ƥ���ҩƷ��"
        txtFind.SetFocus: Exit Sub
    End If
    
    If txtFind.Tag = "" Then
        '��ͬ������ƥ�䷽ʽ
        strInput = UCase(txtFind.Text)
        strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
        If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
        ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
            If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strMatch = " And C.���� Like [2] And C.����=[3]"
        End If
        
        strSQL = _
            " Select Distinct A.ID" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ���� C" & _
            " Where (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And A.������� IN(2,3) And A.ID=C.�շ�ϸĿID And A.��� IN('5','6','7')" & strMatch
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint���� + 1)
        
        strSQL = "0"
        Do While Not rsTmp.EOF
            strSQL = strSQL & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        txtFind.Tag = strSQL
        
        blnFirst = True
    End If
    
    If txtFind.Tag = "0" Then
        MsgBox "�Ҳ���ƥ���ҩƷ��"
        txtFind.SetFocus: Exit Sub
    End If
    
    With vsDrug
        For i = IIF(blnFirst, 1, .Row + 1) To .Rows - 1
            If .RowData(i) <> 0 And InStr("," & txtFind.Tag & ",", "," & .RowData(i) & ",") > 0 Then
                .Row = i: Call .ShowCell(i, .Col): .SetFocus: Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "�Ҳ���ƥ���ҩƷ��"
            txtFind.SetFocus: Exit Sub
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveData
End Sub

Private Function CheckDate() As Boolean
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "��ʼʱ��Ӧ�ñȽ���ʱ��С��", vbInformation, gstrSysName
        Exit Function
    End If
    If DateDiff("d", dtpBegin.Value, dtpEnd.Value) >= 7 And (dtpBegin.Tag <> "" Or dtpEnd.Tag <> "") Then
        If MsgBox("���õ�ʱ�䷶Χ̫�󣬿�������ϵͳ��ѯ�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
        dtpBegin.Tag = "": dtpEnd.Tag = ""
    End If
    CheckDate = True
End Function

Private Sub cmdLoad_Click()
    Dim arrData As Variant, strData As String
    Dim lngRow As Long, i As Long
    Dim sng������ As Single, sng����� As Single
    
    If Not CheckDate Then dtpBegin.SetFocus: Exit Sub
    
    For i = 1 To lvwWay.ListItems.Count
        If lvwWay.ListItems(i).Checked Then Exit For
    Next
    If i > lvwWay.ListItems.Count Then
        MsgBox "û��ָ����ҩ;����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Not (vsDrug.RowData(vsDrug.Row) = 0 And vsDrug.Row = vsDrug.Rows - 1 And vsDrug.Rows = vsDrug.FixedRows + 1) Then
        If MsgBox("ȷʵҪ��ȡ���д�����ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    'If mrsDrug Is Nothing Then Call LoadDrugPut
    Call LoadDrugPut
    
    Screen.MousePointer = 11
    
    With vsDrug
        '��¼ԭ��������
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, col������)) > 0 Then
                strData = strData & ";" & .RowData(i) & "," & Val(.TextMatrix(i, col������))
            End If
        Next
        strData = Mid(strData, 2)
        
        'װ�����ҩƷ
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        mrsDrug.Filter = 0
        If Not mrsDrug.EOF Then
            .Rows = .FixedRows + mrsDrug.RecordCount
            For i = .FixedRows To .FixedRows + mrsDrug.RecordCount - 1
                .RowData(i) = Val(mrsDrug!ҩƷID)
                .TextMatrix(i, col����) = Nvl(mrsDrug!����)
                .TextMatrix(i, colҩƷ) = Nvl(mrsDrug!����)
                .TextMatrix(i, col���) = Nvl(mrsDrug!���)
                .TextMatrix(i, col����) = Nvl(mrsDrug!����)
                .TextMatrix(i, col��λ) = Nvl(mrsDrug!��λ)
                .TextMatrix(i, col���) = Nvl(mrsDrug!���)
                .TextMatrix(i, colסԺ��װ) = Nvl(mrsDrug!סԺ��װ, 0)
                
                .TextMatrix(i, colӦ����) = FormatEx(Nvl(mrsDrug!����, 0), 5)
                Call GetDrugApply(mrsDrug!ҩƷID, sng������, sng�����)
                .TextMatrix(i, col������ҩ��) = sng������
                .TextMatrix(i, col�����ҩ��) = sng�����
                .TextMatrix(i, col������) = GetSurplus(i)
                
                mrsDrug.MoveNext
            Next
        Else
            .Rows = .FixedRows + 1
        End If
        
        '����֮ǰ��������
        arrData = Split(strData, ";")
        For i = 0 To UBound(arrData)
            lngRow = .FindRow(Val(Split(arrData(i), ",")(0)))
            If lngRow <> -1 Then
                If Val(.TextMatrix(lngRow, col������)) = 0 Then
                    .TextMatrix(lngRow, col������) = Val(Split(arrData(i), ",")(1))
                    
                    'ԭ����������������ڵ�Ӧ�����󣬴�����ʾ
                    If Val(.TextMatrix(lngRow, col������)) > Val(.TextMatrix(lngRow, colӦ����)) Then
                        .TextMatrix(lngRow, col������) = Val(.TextMatrix(lngRow, colӦ����))
                        .Cell(flexcpFontBold, lngRow, col������) = True
                    End If
                End If
            End If
        Next
        
        .Row = .FixedRows
        .Col = IIF(.RowData(.Row) = 0, colҩƷ, col������)
        Call vsDrug_AfterRowColChange(-1, -1, .Row, .Col)
        Call .ShowCell(.Row, .Col)
        
        .Redraw = flexRDDirect
    End With
    
    Screen.MousePointer = 0
    
    mblnChange = True
    
    vsDrug.SetFocus
End Sub

Private Sub cmdSelALL_Click()
    Call SelectLVW(lvwWay, True)
    Call ReleaseRecord(mrsDrug)
End Sub

Private Sub dtpBegin_Change()
    dtpBegin.Tag = "1"
    Call ReleaseRecord(mrsDrug)
    Call ReleaseRecord(mrsAdvice)
    Call ReleaseRecord(mrsApply)
End Sub

Private Sub dtpEnd_Change()
    dtpEnd.Tag = "1"
    Call ReleaseRecord(mrsDrug)
    Call ReleaseRecord(mrsAdvice)
    Call ReleaseRecord(mrsApply)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If cmdFind.Enabled And cmdFind.Visible Then
            cmdFind_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPane As Pane
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngҩ��ID As Long
    Dim i As Long
    
    For i = 0 To vsDrug.Cols - 1
        If vsDrug.ColHidden(i) Then
            vsDrug.ColWidth(i) = 0 'Ϊ֧��PrintMode
        Else
            vsDrug.MergeCol(i) = True
        End If
    Next
    vsDrug.MergeRow(0) = True
    
    '������----------------------------------------------
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
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add FALT, vbKeyX, conMenu_File_Exit
    End With
    
    '����ؼ�------------------------------------------
    Call tkpMain.SetMargins(8, 8, 8, 8, 8)
    
    Set objGroup = tkpMain.Groups.Add(0, "����ҩ��")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = cboҩ��
    
    Set objGroup = tkpMain.Groups.Add(0, "��ҩ;��")
    objGroup.Expanded = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picWay
    picWay.BackColor = objItem.BackColor
    If InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") = 0 Then
        lvwWay.Enabled = False
        cmdSelALL.Visible = False
        cmdClear.Visible = False
    End If
    
    Set objGroup = tkpMain.Groups.Add(0, "����ҩƷ")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic����
    pic����.BackColor = objItem.BackColor
    optRule(0).BackColor = objItem.BackColor
    optRule(1).BackColor = objItem.BackColor
    optRule(2).BackColor = objItem.BackColor
    optRule(3).BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(0, "����")
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picFind
    picFind.BackColor = objItem.BackColor
    
    '���ݳ�ʼ-------------------------------------------
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    
    dtpEnd.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(dtpEnd.Value, "yyyy-MM-dd 00:00:00")
    
    cbo����.Text = Val(zlDatabase.GetPara("ȱʡ�������", glngSys, pסԺҽ������, "50", Array(cbo����), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)) & "%"
    If Not cbo����.Enabled Then
        cbo����.Tag = "1" '��ʶ�̶�������
    Else
        cbo����.Enabled = False 'ȱʡѡ��Ӧ�ǲ�����
    End If
    
    optRule(Val(zlDatabase.GetPara("ȱʡ�������", glngSys, pסԺҽ������, "0", Array(optRule(0), optRule(1), optRule(2), optRule(3)), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0))).Value = True
    
    'סԺҩ��
    lngҩ��ID = Val(zlDatabase.GetPara("ȱʡ����ҩ��", glngSys, pסԺҽ������, , Array(cboҩ��), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0))
    strSQL = _
        "Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " AND B.����ID=A.ID And B.������� IN(2,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " Order by A.����"
    On Error GoTo errH
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cboҩ��.AddItem rsTmp!���� & "-" & rsTmp!����
        cboҩ��.ItemData(cboҩ��.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngҩ��ID Then
            Call zlControl.CboSetIndex(cboҩ��.hWnd, cboҩ��.NewIndex)
        End If
        rsTmp.MoveNext
    Loop
    If cboҩ��.ListCount = 0 Then
        MsgBox "û�п��õ�סԺҩ�������ȵ����Ź����н������á�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    If cboҩ��.ListIndex = -1 Then
        Call zlControl.CboSetIndex(cboҩ��.hWnd, 0)
    End If
    
    '��ҩ;��
    Call LoadDrugWay
    
    mlngPeҩ�� = -1
    mblnChange = False
    Call cboҩ��_Click

    '-------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("��ǰ������û�б��棬ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    Call zlDatabase.SetPara("����ǼǸ�ҩ;������", GetSelDrugWay, glngSys, pסԺҽ������, InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
    
    If cboҩ��.ListIndex <> -1 Then
        Call zlDatabase.SetPara("ȱʡ����ҩ��", cboҩ��.ItemData(cboҩ��.ListIndex), glngSys, pסԺҽ������, InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
    End If
    Call zlDatabase.SetPara("ȱʡ�������", IIF(optRule(0).Value, 0, IIF(optRule(1).Value, 1, IIF(optRule(2).Value, 2, 3))), glngSys, pסԺҽ������, InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
    Call zlDatabase.SetPara("ȱʡ�������", Val(cbo����.Text), glngSys, pסԺҽ������, InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
    
    Call SaveWinState(Me, App.ProductName)
    
    Call ReleaseRecord(mrsDrug)
    Call ReleaseRecord(mrsAdvice)
    Call ReleaseRecord(mrsApply)
End Sub

Private Sub ReleaseRecord(rsData As ADODB.Recordset)
    If Not rsData Is Nothing Then
        If rsData.State = 1 Then rsData.Close
        Set rsData = Nothing
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tkpMain.Width + X < 2700 Or vsDrug.Width - X < 3000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tkpMain.Width = tkpMain.Width + X
        vsDrug.Left = vsDrug.Left + X
        vsDrug.Width = vsDrug.Width - X
        Me.Refresh
    End If
End Sub

Private Sub lvwWay_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ReleaseRecord(mrsDrug)
End Sub

Private Sub lvwWay_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        If cmdSelALL.Enabled And cmdSelALL.Visible Then Call cmdSelALL_Click
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        If cmdClear.Enabled And cmdClear.Visible Then Call cmdClear_Click
    End If
End Sub

Private Sub optRule_Click(Index As Integer)
    cbo����.Enabled = optRule(2).Value And cbo����.Tag = ""
End Sub

Private Sub picFind_Resize()
    On Error Resume Next
    
    txtFind.Width = picFind.ScaleWidth - txtFind.Left
    cmdFind.Left = picFind.ScaleWidth - cmdFind.Width + 15
End Sub

Private Sub picWay_Resize()
    On Error Resume Next
    
    lvwWay.Left = 0
    lvwWay.Top = 0
    lvwWay.Width = picWay.ScaleWidth
    lvwWay.Height = picWay.ScaleHeight - IIF(cmdSelALL.Visible, cmdSelALL.Height - 15, 0)
    
    cmdClear.Top = picWay.ScaleHeight - cmdClear.Height + 15
    cmdClear.Left = picWay.ScaleWidth - cmdClear.Width + 15
    
    cmdSelALL.Top = cmdClear.Top
    cmdSelALL.Left = cmdClear.Left - cmdSelALL.Width + 30
    
    lvwWay.ColumnHeaders(1).Width = lvwWay.Width - 360
End Sub

Private Sub pic����_Resize()
    On Error Resume Next
    
    dtpBegin.Width = pic����.ScaleWidth - dtpBegin.Left
    dtpEnd.Width = pic����.ScaleWidth - dtpEnd.Left
    cmdLoad.Left = pic����.ScaleWidth - cmdLoad.Width + 15
End Sub

Private Sub txtFind_Change()
    txtFind.Tag = ""
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlControl.TxtSelAll txtFind
        Call cmdFind_Click
    End If
End Sub

Private Sub vsDrug_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        If vsDrug.Redraw <> flexRDNone Then
            If OldRow <> -1 And OldRow <= vsDrug.Rows - 1 Then
                vsDrug.Cell(flexcpForeColor, OldRow, 0, OldRow, vsDrug.Cols - 1) = vsDrug.ForeColor
            End If
            If NewRow <> -1 Then
                vsDrug.Cell(flexcpForeColor, NewRow, 0, NewRow, vsDrug.Cols - 1) = vbBlue
            End If
        End If
    End If
    
    If CellEditable(NewRow, NewCol) Then
        vsDrug.FocusRect = flexFocusSolid
        If NewCol = colҩƷ Then
            vsDrug.ComboList = "..."
        Else
            vsDrug.ComboList = ""
        End If
    Else
        vsDrug.ComboList = ""
        vsDrug.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsDrug_AfterSort(ByVal Col As Long, Order As Integer)
    vsDrug.Cell(flexcpForeColor, vsDrug.Row, 0, vsDrug.Row, vsDrug.Cols - 1) = vbBlue
End Sub

Private Sub vsDrug_BeforeSort(ByVal Col As Long, Order As Integer)
    vsDrug.Cell(flexcpForeColor, vsDrug.Row, 0, vsDrug.Row, vsDrug.Cols - 1) = vsDrug.ForeColor
End Sub

Private Sub vsDrug_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnCancel As Boolean
    Dim str���� As String, str���� As String, str��� As String

    If Col = colҩƷ Then
        str���� = Get��������(cboҩ��.ItemData(cboҩ��.ListIndex))
        str���� = " And ���� IN(1,2,3)"
        str��� = " And A.��� IN('5','6','7')"
        If InStr(str����, "��ҩ��") = 0 Then
            str���� = Replace(str����, "1,", "")
            str��� = Replace(str���, "'5',", "")
        End If
        If InStr(str����, "��ҩ��") = 0 Then
            str���� = Replace(str����, "2,", "")
            str��� = Replace(str���, "'6',", "")
        End If
        If InStr(str����, "��ҩ��") = 0 Then
            str���� = Replace(str����, ",3", "")
            str��� = Replace(str���, ",'7'", "")
        End If
        
        strSQL = _
            " Select Distinct 0 as ĩ��,To_Number('999999999'||����) as ID,-NULL as �ϼ�ID," & _
            " CHR(13)||���� as ����,Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ') as ����," & _
            " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���ID,-NULL as ϵ��ID" & _
            " From ���Ʒ���Ŀ¼ Where (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & str����
        strSQL = strSQL & " Union ALL " & _
            " Select 0 as ĩ��,-1*ID as ID,Nvl(-1*�ϼ�ID,To_Number('999999999'||����)) as �ϼ�ID," & _
            " ����,����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���ID,-NULL as ϵ��ID" & _
            " From ���Ʒ���Ŀ¼ Where (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & str���� & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
        strSQL = strSQL & " Union ALL " & _
            " Select Distinct 1 as ĩ��,A.ID,-1*E.����ID as �ϼ�ID,A.����," & _
            " Nvl(F.����,A.����) as ����,D.סԺ��λ as ��λ,A.���,A.����,A.��� as ���ID,D.סԺ��װ as ϵ��ID" & _
            " From �շ���ĿĿ¼ A,ҩƷ��� D,������ĿĿ¼ E,�շ���Ŀ���� F" & _
            " Where (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And A.������� IN(2,3) And A.ID=D.ҩƷID And D.ҩ��ID=E.ID" & _
            " And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[1]" & str���
        strSQL = strSQL & " Order by ����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "ҩƷ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, IIF(gbytҩƷ������ʾ = 0, 1, 3))
        If Not rsTmp Is Nothing Then
            If SetItemInput(Row, rsTmp) Then Call EnterNextCell(Row, Col)
        Else
            If Not blnCancel Then
                MsgBox "û�п��õ�ҩƷ�����ȵ�ҩƷĿ¼���������ã�", vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function SetItemInput(ByVal lngRow As Long, rsTmp As ADODB.Recordset) As Boolean
    Dim lngFind As Long
    Dim sng������ As Single, sng����� As Single
    
    With vsDrug
        lngFind = .FindRow(Val(rsTmp!ID))
        If lngFind <> -1 And lngFind <> lngRow Then
            MsgBox "ҩƷ""" & rsTmp!���� & """�Ѿ�¼�롣", vbInformation, gstrSysName
            Exit Function
        End If
        
        '���ݸ�ҩ;�����ƴ���ҩƷ��������
        If GetSelDrugWay <> "" Then
            If mrsDrug Is Nothing Then Call LoadDrugPut
            mrsDrug.Filter = "ҩƷID=" & rsTmp!ID
            If mrsDrug.EOF Then
                MsgBox "��ǰָ����ҩ;����������ҩƷ""" & rsTmp!���� & """û�д���ҩ��¼��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        .RowData(lngRow) = Val(rsTmp!ID)
        .TextMatrix(lngRow, col����) = Nvl(rsTmp!����)
        .TextMatrix(lngRow, colҩƷ) = Nvl(rsTmp!����)
        .TextMatrix(lngRow, col���) = Nvl(rsTmp!���)
        .TextMatrix(lngRow, col����) = Nvl(rsTmp!����)
        .TextMatrix(lngRow, col��λ) = Nvl(rsTmp!��λ)
        .TextMatrix(lngRow, col���) = Nvl(rsTmp!���ID)
        .TextMatrix(lngRow, colסԺ��װ) = Nvl(rsTmp!ϵ��ID, 0)
        
        .TextMatrix(lngRow, colӦ����) = GetDrugPut(rsTmp!ID)
        Call GetDrugApply(rsTmp!ID, sng������, sng�����)
        .TextMatrix(lngRow, col������ҩ��) = sng������
        .TextMatrix(lngRow, col�����ҩ��) = sng�����
        .TextMatrix(lngRow, col������) = GetSurplus(lngRow)
        
        .Cell(flexcpFontBold, lngRow, col������) = False
        If Val(.TextMatrix(lngRow, col������)) > 0 Then mblnChange = True
    End With
    
    SetItemInput = True
End Function

Private Sub vsDrug_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsDrug
        '���������ɫ���
        If Col = col������ Then
            If Val(.TextMatrix(Row, Col)) > 0 Then
                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = &HC0FFFF
            Else
                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = .BackColor
            End If
        End If
    End With
End Sub

Private Sub vsDrug_DblClick()
    Call vsDrug_KeyPress(32)
End Sub

Private Sub vsDrug_GotFocus()
    '�ô����д������ؼ������л������Żἤ��
    If Not CheckDate Then dtpBegin.SetFocus
End Sub

Private Sub vsDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsDrug
        If KeyCode = vbKeyDelete Then
            If vsDrug.Col = col������ Then
                If vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) <> "" Then
                    vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) = ""
                    vsDrug.CellFontBold = False
                    mblnChange = True
                End If
            Else
                If .RowData(.Row) <> 0 Then
                    If MsgBox("ȷ��Ҫɾ����ǰҩƷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    If Val(.TextMatrix(.Row, col������)) <> 0 Then mblnChange = True
                End If
                
                .RemoveItem .Row
    
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = colҩƷ
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDrug_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDrug_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell(vsDrug.Row, vsDrug.Col)
    Else
        If vsDrug.Col = colҩƷ Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsDrug_CellButtonClick(vsDrug.Row, vsDrug.Col)
            Else
                vsDrug.ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End If
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    If lngCol < colҩƷ Then
        vsDrug.Col = colҩƷ
    ElseIf lngCol < col������ Then
        If vsDrug.RowData(lngRow) <> 0 Then
            vsDrug.Col = col������
        End If
    ElseIf vsDrug.RowData(lngRow) <> 0 Then
        If lngRow = vsDrug.Rows - 1 Then vsDrug.AddItem ""
        vsDrug.Row = vsDrug.Row + 1
        vsDrug.Col = colҩƷ
    End If
    vsDrug.ShowCell vsDrug.Row, vsDrug.Col
End Sub

Private Sub vsDrug_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
        If Col = col������ Then
            If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub vsDrug_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDrug.EditSelStart = 0
    vsDrug.EditSelLength = zlCommFun.ActualLen(vsDrug.EditText)
End Sub

Private Sub vsDrug_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not CellEditable(Row, Col) Then Cancel = True
    If Col = col������ Then
        vsDrug.EditMaxLength = 10
    Else
        vsDrug.EditMaxLength = 0
    End If
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    CellEditable = True
    If Not (lngCol = colҩƷ Or lngCol = col������) Then
        CellEditable = False
    ElseIf lngCol = col������ And vsDrug.RowData(lngRow) = 0 Then
        CellEditable = False
    End If
End Function

Private Function LoadSurplus() As Boolean
'���ܣ���ȡ��ǰҩ������д�������¼
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngҩ��ID As Long
    Dim sng������ As Single, sng����� As Single
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If cboҩ��.ListIndex <> -1 Then
        lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
    End If
    strSQL = "Select A.ҩƷID,C.����,Nvl(D.����,C.����) as ����,C.���,C.����," & _
        " B.סԺ��λ as ��λ,A.��������/Nvl(B.סԺ��װ,1) as ��������,C.���,B.סԺ��װ" & _
        " From ҩƷ����ƻ� A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
        " Where A.ҩƷID=B.ҩƷID And A.ҩƷID=C.ID And A.����ID=[1] And A.�ⷿID=[2]" & _
        " And A.״̬=0 And C.ID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=[3]" & _
        " Order by C.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, lngҩ��ID, IIF(gbytҩƷ������ʾ = 0, 1, 3))
    
    With vsDrug
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = .FixedRows To .FixedRows + rsTmp.RecordCount - 1
                .RowData(i) = Val(rsTmp!ҩƷID)
                .TextMatrix(i, col����) = Nvl(rsTmp!����)
                .TextMatrix(i, colҩƷ) = Nvl(rsTmp!����)
                .TextMatrix(i, col���) = Nvl(rsTmp!���)
                .TextMatrix(i, col����) = Nvl(rsTmp!����)
                .TextMatrix(i, col��λ) = Nvl(rsTmp!��λ)
                .TextMatrix(i, col���) = Nvl(rsTmp!���)
                .TextMatrix(i, colסԺ��װ) = Nvl(rsTmp!סԺ��װ, 0)
                
                .TextMatrix(i, colӦ����) = GetDrugPut(rsTmp!ҩƷID) '��ǰӦ����
                Call GetDrugApply(rsTmp!ҩƷID, sng������, sng�����)
                .TextMatrix(i, col������ҩ��) = sng������
                .TextMatrix(i, col�����ҩ��) = sng�����
                .TextMatrix(i, col������) = Nvl(rsTmp!��������)
                
                'ԭ����������������ڵ�Ӧ�����󣬴�����ʾ
                If Val(.TextMatrix(i, col������)) > Val(.TextMatrix(i, colӦ����)) Then
                    .TextMatrix(i, col������) = Val(.TextMatrix(i, colӦ����))
                    .Cell(flexcpFontBold, i, col������) = True
                End If
                
                rsTmp.MoveNext
            Next
        Else
            .Rows = .FixedRows + 1
        End If
        
        .Row = .FixedRows
        .Col = IIF(.RowData(.Row) = 0, colҩƷ, col������)
        Call vsDrug_AfterRowColChange(-1, -1, .Row, .Col)
        Call .ShowCell(.Row, .Col)
    End With
    
    Screen.MousePointer = 0
    LoadSurplus = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    Dim arrSQL As Variant, i As Long
    
    arrSQL = Array()
    With vsDrug
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҩƷ����ƻ�_Delete(" & mlng����ID & "," & cboҩ��.ItemData(cboҩ��.ListIndex) & ")"
        
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 And Val(.TextMatrix(i, col������)) > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_ҩƷ����ƻ�_Insert(" & mlng����ID & "," & cboҩ��.ItemData(cboҩ��.ListIndex) & "," & _
                    .RowData(i) & "," & Val(.TextMatrix(i, col������)) * Val(.TextMatrix(i, colסԺ��װ)) & ",'" & UserInfo.���� & "')"
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    vsDrug.Cell(flexcpFontBold, vsDrug.FixedRows, col������, vsDrug.Rows - 1, col������) = False
    mblnChange = False
    SaveData = True
    
    Screen.MousePointer = 0
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetSelDrugWay(Optional ByVal blnbyName As Boolean) As String
'���ܣ�����ûѡ��ĸ�ҩ;����
    Dim i As Long, strData As String
    
    For i = 1 To lvwWay.ListItems.Count
        If Not lvwWay.ListItems(i).Checked Then
            If blnbyName Then
                strData = strData & "," & lvwWay.ListItems(i).Text
            Else
                strData = strData & "," & Mid(lvwWay.ListItems(i).Key, 2)
            End If
        End If
    Next
    strData = Mid(strData, 2)
    If UBound(Split(strData, ",")) + 1 = lvwWay.ListItems.Count Then
        strData = "" 'ȫ���ų�Ҳ������
    End If
    GetSelDrugWay = strData
End Function

Private Function LoadDrugApply() As Boolean
'���ܣ���ȡ��ҩ�����������Ϣ
    Dim strSQL As String
    Dim intMouse As Integer
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.�շ�ϸĿID as ҩƷID," & _
        " Sum(A.����/Nvl(B.סԺ��װ,1)) as ������," & _
        " Sum(Decode(A.״̬,1,A.����/Nvl(B.סԺ��װ,1),0)) as �����" & _
        " From ���˷������� A,ҩƷ��� B" & _
        " Where A.�շ�ϸĿID=B.ҩƷID And A.����ʱ�� Between [1] And [2] And A.���벿��ID=[3] And A.��˲���ID=[4]" & _
        " Group by A.�շ�ϸĿID"
    Set mrsApply = New ADODB.Recordset
    Set mrsApply = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(dtpBegin.Value), CDate(dtpEnd.Value), mlng����ID, cboҩ��.ItemData(cboҩ��.ListIndex))
    
    Screen.MousePointer = intMouse
    LoadDrugApply = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ReleaseRecord(mrsApply)
End Function

Private Function LoadDrugPut() As Boolean
'���ܣ���ȡҩƷδ������
    Dim strSQL As String
    Dim str��ҩ As String
    Dim intMouse As Integer
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    str��ҩ = GetSelDrugWay(True)

    strSQL = _
        " Select /*+ Rule*/ A.ҩƷID,D.����,Nvl(D.����,E.����) As ����,C.סԺ��λ As ��λ," & _
        " D.���,D.����,Sum(A.��д����/Nvl(C.סԺ��װ,1)) as ����,D.���,C.סԺ��װ" & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,ҩƷ��� C,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
        " Where A.���� = 9 And A.NO = B.NO And B.��¼���� = 2 And A.����ID = B.ID And A.ҩƷID = C.ҩƷID" & _
        " And C.ҩƷID=D.ID And Mod(A.��¼״̬,3)=1 And A.����� is Null" & _
        " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=[5]" & _
        IIF(str��ҩ <> "", " And Nvl(A.�÷�,'Null') Not IN(Select Column_Value From Table(Cast(f_Str2list([6]) As zlTools.t_Strlist)))", "") & _
        " And A.�������� Between [1] And [2] And A.�ⷿID=[3] And B.���˲���ID=[4]" & _
        " Group By A.ҩƷID,D.����,Nvl(D.����,E.����),C.סԺ��λ,D.���,D.����,D.���,C.סԺ��װ" & _
        " Order By ����"
    Set mrsDrug = New ADODB.Recordset
    Set mrsDrug = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(dtpBegin.Value), CDate(dtpEnd.Value), cboҩ��.ItemData(cboҩ��.ListIndex), mlng����ID, IIF(gbytҩƷ������ʾ = 0, 1, 3), str��ҩ)
    
    Screen.MousePointer = intMouse
    LoadDrugPut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ReleaseRecord(mrsDrug)
End Function

Private Sub GetDrugApply(ByVal lngҩƷID As Long, sng������ As Single, sng����� As Single)
    sng������ = 0: sng����� = 0
    
    If mrsApply Is Nothing Then Call LoadDrugApply
    
    mrsApply.Filter = "ҩƷID=" & lngҩƷID
    If Not mrsApply.EOF Then
        sng������ = FormatEx(Nvl(mrsApply!������, 0), 5)
        sng����� = FormatEx(Nvl(mrsApply!�����, 0), 5)
    End If
End Sub

Private Function GetDrugPut(ByVal lngҩƷID As Long) As Single
    If mrsDrug Is Nothing Then Call LoadDrugPut
    
    mrsDrug.Filter = "ҩƷID=" & lngҩƷID
    If Not mrsDrug.EOF Then
        GetDrugPut = FormatEx(Nvl(mrsDrug!����, 0), 5)
    End If
End Function

Private Sub vsDrug_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strInput As String, strMatch As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    Dim str���� As String, str��� As String
    
    With vsDrug
        If Col = colҩƷ And .EditText <> "" Then
            str���� = Get��������(cboҩ��.ItemData(cboҩ��.ListIndex))
            str��� = " And A.��� IN('5','6','7')"
            If InStr(str����, "��ҩ��") = 0 Then str��� = Replace(str���, "'5',", "")
            If InStr(str����, "��ҩ��") = 0 Then str��� = Replace(str���, "'6',", "")
            If InStr(str����, "��ҩ��") = 0 Then str��� = Replace(str���, ",'7'", "")
            
            '��ͬ������ƥ�䷽ʽ
            strInput = UCase(.EditText)
            strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
            If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
            ElseIf zlCommFun.IsCharChinese(strInput) Then
                strMatch = " And C.���� Like [2] And C.����=[3]"
            End If
            
            strSQL = _
                " Select Distinct 1 as ĩ��,A.ID,A.����,C.����," & _
                " B.סԺ��λ as ��λ,A.���,A.����,A.��� as ���ID,B.סԺ��װ as ϵ��ID" & _
                " From �շ���ĿĿ¼ A,ҩƷ��� B,�շ���Ŀ���� C" & _
                " Where (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And A.������� IN(2,3) And A.ID=B.ҩƷID And A.ID=C.�շ�ϸĿID" & str��� & strMatch & _
                " Order by ����"
            vPoint = GetCoordPos(.hWnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩƷ", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                strInput & "%", mstrLike & strInput & "%", mint���� + 1)
            If Not rsTmp Is Nothing Then
                If Not SetItemInput(Row, rsTmp) Then
                    Cancel = True
                Else
                    .EditText = .Text
                    If mblnReturn Then
                        Call EnterNextCell(Row, Col)
                    End If
                End If
            Else
                If Not blnCancel Then
                    MsgBox "����""" & .EditText & """û���ҵ����õ�ҩƷ��", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
            mblnReturn = False
        ElseIf Col = col������ Then
            If Not IsNumeric(.EditText) And .EditText <> "" Or Val(.EditText) < 0 Or Val(.EditText) > LONG_MAX Then
                MsgBox "�������������""" & .EditText & """���󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                Cancel = True
            Else
                .CellFontBold = False
                If Val(.EditText) = 0 Then .EditText = ""
                mblnChange = True
                If mblnReturn Then
                    Call EnterNextCell(Row, Col)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Function LoadDrugAdvice() As Boolean
'���ܣ���ȡҩƷҽ����Ϣ�����ڼ���
'˵����
'   û�а�����ҩƷҽ����ҩƷ�Ƽ�
'   �����˳�����������´�����
    Dim strSQL As String
    Dim intMouse As Integer
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    strSQL = _
        "Select M.��Ժ����,A.��ʼִ��ʱ��,A.ҽ����Ч,D.ҩƷID,D.����ϵ��,D.סԺ��װ,Nvl(A.�ɷ����,D.סԺ�ɷ����) as �ɷ����," & _
        " B.�״�ʱ��,B.ĩ��ʱ��,A.����,A.ִ��ʱ�䷽��,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.��������,B.��������" & _
        " From ����ҽ����¼ A,����ҽ������ B,סԺ���ü�¼ C,ҩƷ��� D,������ҳ M" & _
        " Where A.������� IN('5','6') And A.ID=B.ҽ��ID And A.����ID=M.����ID And A.��ҳID=M.��ҳID" & _
        " And B.NO=C.NO And B.��¼����=C.��¼���� And B.ҽ��ID=C.ҽ����� And C.��¼״̬ IN(0,1,3)" & _
        " And C.�շ�ϸĿID=D.ҩƷID And B.����ʱ�� Between [1] And [2] And B.ִ�в���ID=[3] And C.���˲���ID=[4]"
    Set mrsAdvice = New ADODB.Recordset
    Set mrsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(dtpBegin.Value), CDate(dtpEnd.Value), cboҩ��.ItemData(cboҩ��.ListIndex), mlng����ID)
    
    Screen.MousePointer = intMouse
    LoadDrugAdvice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ReleaseRecord(mrsAdvice)
End Function

Private Function GetSurplus(ByVal lngRow As Long) As String
'���ܣ�����ָ����ҩƷ��ȱʡ��������
    Dim dbl���� As Double, dbl�ܵ��� As Double
    Dim lng���� As Long, strTime As String
    
    If optRule(0).Value Then
        GetSurplus = ""
    ElseIf optRule(1).Value Then
        GetSurplus = Val(vsDrug.TextMatrix(lngRow, colӦ����))
    ElseIf optRule(2).Value Then
        dbl���� = Val(cbo����.Text) / 100
        If dbl���� > 1 Then dbl���� = 1
        If dbl���� < 0 Then dbl���� = 0
        GetSurplus = IntEx(Val(vsDrug.TextMatrix(lngRow, colӦ����)) * dbl����)
    ElseIf optRule(3).Value Then
        '�����������
        If vsDrug.TextMatrix(lngRow, col���) = "7" Then
            GetSurplus = "" '��ҩӦ����������
        Else
            If mrsAdvice Is Nothing Then Call LoadDrugAdvice
            mrsAdvice.Filter = "ҩƷID=" & vsDrug.RowData(lngRow)
            If Not mrsAdvice.EOF Then
                Do While Not mrsAdvice.EOF
                    '�μ�ҩƷҽ�����ͣ�
                    '��ͬҩƷ����ͬҽ��������Ч��Ƶ�ʵȲ�ͬ
                    '�ܵ���=������ҩ����*�ڴε���
                    If Nvl(mrsAdvice!ҽ����Ч, 0) = 0 Then
                        If Not IsNull(mrsAdvice!�״�ʱ��) And Not IsNull(mrsAdvice!ĩ��ʱ��) And Not IsNull(mrsAdvice!ִ��ʱ�䷽��) Then
                            strTime = Calc���ڷֽ�ʱ��(mrsAdvice!�״�ʱ��, mrsAdvice!ĩ��ʱ��, "", _
                                mrsAdvice!ִ��ʱ�䷽��, mrsAdvice!Ƶ�ʴ���, mrsAdvice!Ƶ�ʼ��, mrsAdvice!�����λ, mrsAdvice!��ʼִ��ʱ��)
                            lng���� = UBound(Split(strTime, ",")) + 1
                            dbl�ܵ��� = dbl�ܵ��� + lng���� * mrsAdvice!��������
                        Else
                            dbl�ܵ��� = dbl�ܵ��� + mrsAdvice!�������� '�쳣�����ֱ��ȡ�����ܵ���
                        End If
                    ElseIf Not IsNull(mrsAdvice!��������) Then
                        If Nvl(mrsAdvice!Ƶ�ʴ���, 0) = 0 Or Nvl(mrsAdvice!Ƶ�ʼ��, 0) = 0 Then
                            lng���� = 1 '����Ϊһ���Ե�����ҩƷ
                        ElseIf Nvl(mrsAdvice!����, 0) <> 0 And Not IsNull(mrsAdvice!ִ��Ƶ��) Then
                            '��ҩ�����ڰ�Ƶ�����ڵĴ���
                            If mrsAdvice!�����λ = "��" Then
                                lng���� = IntEx(mrsAdvice!���� * (mrsAdvice!Ƶ�ʴ��� / 7))
                            ElseIf mrsAdvice!�����λ = "��" Then
                                lng���� = IntEx(mrsAdvice!���� * (mrsAdvice!Ƶ�ʴ��� / mrsAdvice!Ƶ�ʼ��))
                            ElseIf mrsAdvice!�����λ = "Сʱ" Then
                                lng���� = IntEx(mrsAdvice!���� * (mrsAdvice!Ƶ�ʴ��� / mrsAdvice!Ƶ�ʼ��) * 24)
                            ElseIf mrsAdvice!�����λ = "����" Then
                                lng���� = IntEx(mrsAdvice!���� * (mrsAdvice!Ƶ�ʴ��� / mrsAdvice!Ƶ�ʼ��) * (24 * 60))
                            End If
                        Else
                            '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,����һ��Ƶ�����ڵĴ�������
                            If Nvl(mrsAdvice!�ɷ����, Nvl(mrsAdvice!�ɷ����, 0)) = 0 And Nvl(mrsAdvice!��������, 0) <> 0 Then
                                lng���� = IntEx(mrsAdvice!�ܸ����� * mrsAdvice!����ϵ�� / mrsAdvice!��������)
                            Else
                                lng���� = Nvl(mrsAdvice!Ƶ�ʴ���, 0)
                            End If
                        End If
                        
                        dbl�ܵ��� = dbl�ܵ��� + lng���� * mrsAdvice!��������
                    End If
                    
                    mrsAdvice.MoveNext
                Loop
                
                mrsAdvice.MoveFirst 'ȡһЩҩƷ��Ϣ,תΪסԺ��λ
                dbl�ܵ��� = IntEx(dbl�ܵ��� / Nvl(mrsAdvice!����ϵ��, 1) / Nvl(mrsAdvice!סԺ��װ, 1))
                If dbl�ܵ��� > Val(vsDrug.TextMatrix(lngRow, colӦ����)) Then
                    dbl�ܵ��� = Val(vsDrug.TextMatrix(lngRow, colӦ����))
                End If
                '��������=Ӧ������-ʵ������
                GetSurplus = Val(vsDrug.TextMatrix(lngRow, colӦ����)) - dbl�ܵ���
            End If
        End If
    End If
    
    If Val(GetSurplus) = 0 Then GetSurplus = ""
End Function

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    
    '��ͷ
    objOut.Title.Text = Get��������(mlng����ID) & "ҩƷ�����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "ҩ����" & NeedName(cboҩ��.Text)
    objRow.Add "ʱ�䣺" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm") & " �� " & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsDrug
    
    '���
    vsDrug.Redraw = flexRDNone
    lngRow = vsDrug.Row: lngCol = vsDrug.Col
    vsDrug.Cell(flexcpForeColor, vsDrug.Row, 0, vsDrug.Row, vsDrug.Cols - 1) = vsDrug.ForeColor
    
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    vsDrug.Row = lngRow: vsDrug.Col = lngCol
    vsDrug.Redraw = flexRDDirect
    
    Call vsDrug_AfterRowColChange(-1, -1, vsDrug.Row, vsDrug.Col)
End Sub

Private Function LoadDrugWay() As Boolean
'���ܣ���ȡ���õĸ�ҩ;��(����ҩ���г�ҩ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem, str��ҩIDs As String
    
    On Error GoTo errH
    
    str��ҩIDs = zlDatabase.GetPara("����ǼǸ�ҩ;������", glngSys, pסԺҽ������, , Array(lvwWay, cmdSelALL, cmdClear), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
    
    strSQL = "Select ID,����,���� From ������ĿĿ¼" & _
        " Where ���='E' And ��������='2' And ������� IN(2,3) And (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwWay.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����)
        
        If str��ҩIDs <> "" Then
            If InStr("," & str��ҩIDs & ",", "," & rsTmp!ID & ",") = 0 Then
                objItem.Checked = True
            End If
        Else
            objItem.Checked = True
        End If
        rsTmp.MoveNext
    Next
    LoadDrugWay = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub
