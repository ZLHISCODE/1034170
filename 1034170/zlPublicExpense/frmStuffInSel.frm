VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffInSel 
   Caption         =   "������ⵥѡ��"
   ClientHeight    =   7296
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   11772
   Icon            =   "frmStuffInSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7296
   ScaleWidth      =   11772
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   708
      ScaleWidth      =   11772
      TabIndex        =   12
      Top             =   6588
      Width           =   11775
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "ȫѡ(&A)"
         Height          =   380
         Left            =   105
         TabIndex        =   15
         ToolTipText     =   "���:CTRL+A"
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdALLCls 
         Caption         =   "ȫ��(&S)"
         Height          =   380
         Left            =   1365
         TabIndex        =   14
         ToolTipText     =   "���:CTRL+C"
         Top             =   165
         Width           =   1250
      End
      Begin VB.Frame fraBottomSplit 
         Height          =   30
         Left            =   -210
         TabIndex        =   13
         Top             =   0
         Width           =   12405
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   380
         Left            =   8865
         TabIndex        =   7
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   380
         Left            =   10125
         TabIndex        =   8
         Top             =   165
         Width           =   1250
      End
   End
   Begin VB.PictureBox picSeach 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   11772
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   11775
      Begin VB.Frame fraSearch 
         Height          =   960
         Left            =   90
         TabIndex        =   10
         Top             =   -15
         Width           =   11235
         Begin VB.CommandButton cmdSel 
            Caption         =   "����(&F)"
            Height          =   380
            Left            =   8130
            TabIndex        =   5
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txtNO 
            Height          =   330
            Left            =   870
            TabIndex        =   1
            Top             =   355
            Width           =   1770
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   345
            Left            =   4830
            TabIndex        =   3
            Top             =   348
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   614
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   133038083
            CurrentDate     =   40528
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   345
            Left            =   6525
            TabIndex        =   4
            Top             =   348
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   614
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   133038083
            CurrentDate     =   40528
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "��������ڲ���"
            Height          =   330
            Left            =   3195
            TabIndex        =   2
            Top             =   355
            Width           =   1680
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            Caption         =   "��ⵥ��"
            Height          =   180
            Left            =   90
            TabIndex        =   0
            Top             =   430
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   6255
            TabIndex        =   11
            Top             =   430
            Width           =   180
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5490
      Left            =   90
      TabIndex        =   6
      Top             =   1065
      Width           =   11355
      _cx             =   20029
      _cy             =   9684
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.8
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffInSel.frx":06EA
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   3
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   75
         ScaleHeight     =   228
         ScaleWidth      =   216
         TabIndex        =   16
         Top             =   60
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   156
            Left            =   0
            Picture         =   "frmStuffInSel.frx":0717
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   156
         End
      End
   End
End
Attribute VB_Name = "frmStuffInSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String, mblnOK As Boolean
Private mrsSel As ADODB.Recordset, mlng����ⷿID As Long
Public Function zlSelect(ByVal frmMain As Form, ByVal lngMoudle As Long, _
    ByVal strPrivs As String, ByVal lng����ⷿID As Long, ByRef rsReturnSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�����
    '���:frmMain-���õĴ���
    '����:rsReturnSel-���ر�ѡ��Ľ����
    '����:���ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-16 10:28:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModule = lngMoudle: mstrPrivs = strPrivs: mblnOK = False
    Set mrsSel = Nothing: mlng����ⷿID = lng����ⷿID
    Me.Show 1, frmMain
    Set rsReturnSel = mrsSel
    zlSelect = mblnOK
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub chk���_Click()
        dtpStart.Enabled = chk���.Value = 1
        dtpEnd.Enabled = chk���.Value = 1
End Sub

Private Sub chk���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdALLCls_Click()
    Dim i As Long
    With vsItem
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
            .TextMatrix(i, .ColIndex("���μ�������")) = ""
        Next
    End With
End Sub

Private Sub cmdAllSel_Click()
    Dim i As Long
    With vsItem
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 1
            If Val(.TextMatrix(i, .ColIndex("���μ�������"))) = 0 Then
                .TextMatrix(i, .ColIndex("���μ�������")) = .TextMatrix(i, .ColIndex("��������"))
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If zlBuliedRec = False Then Exit Sub
    mblnOK = True
    Unload Me:
End Sub

Private Sub cmdSel_Click()
    Call FillData(False)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyA
            If Shift = vbCtrlMask Then cmdAllSel_Click
        Case vbKeyC
            If Shift = vbCtrlMask Then cmdALLCls_Click
        End Select
End Sub

Private Sub Form_Load()
    RestoreWinState Me, Me.Name
    dtpEnd.MaxDate = gobjDatabase.Currentdate
    dtpEnd.Value = dtpEnd.MaxDate
    dtpEnd.minDate = dtpEnd.MaxDate - 2 * 365
    dtpStart.MaxDate = dtpEnd.MaxDate
    dtpStart.minDate = dtpEnd.minDate
    dtpStart.Value = dtpEnd.Value   'Ĭ��Ϊ����
    Call FillData(True)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsItem
        .Left = Me.ScaleLeft + 50
        .Width = Me.ScaleWidth - 100
        .Top = picSeach.Top + picSeach.Height + 20
        .Height = Me.ScaleHeight - .Top - picDown.Height - 50
    End With
End Sub
 
Private Sub picSeach_Resize()
    Err = 0: On Error Resume Next
    With picSeach
        fraSearch.Left = .ScaleLeft + 50
        fraSearch.Top = .ScaleTop + 50
        fraSearch.Height = .ScaleHeight - 100
        fraSearch.Width = .ScaleWidth - 100
    End With
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        fraBottomSplit.Left = .ScaleLeft
        fraBottomSplit.Top = .ScaleTop
        fraBottomSplit.Width = .ScaleWidth
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - cmdCancel.Width / 2
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub
Private Function FillData(Optional blnDefault As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:blnDefault-�Ƿ����ȱʡ����,�����,�������һ�α������п�����ⵥ��Ϊ����ѡ��Ķ���,������ݽ�������������
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-16 10:40:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, rsTemp As ADODB.Recordset, i As Long
    
    On Error GoTo errHandle
    If blnDefault Then
            strWhere = "And A.No =( " & _
            "                                Select Max(No) As No From ҩƷ�շ���¼ A1, ҩƷ��� B1  " & _
            "                                Where a1.������� Between Sysdate-7 And Sysdate " & _
            "                                        And  A1.ҩƷid = B1.ҩƷid And A1.�ⷿid = B1.�ⷿid And Nvl(A1.����, 0) = Nvl(B1.����, 0)   " & _
            "                                        And Nvl(B1.��������, 0) > 0 And A1.�ⷿid = [1] ) "
    Else
        strWhere = ""
        If txtNO.Text <> "" Then strWhere = " And A.NO=[2] "
        If chk���.Value = 1 Then strWhere = strWhere & " and  (A.������� between [3] and [4] )"
        strWhere = strWhere & " and   A.�ⷿID=[1] "
        If strWhere = "" Then
            MsgBox "ע��:" & vbCrLf & "    ��ѯǰ�������뵥�ݺŻ����������!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    gstrSQL = " " & _
    " Select  a.Id,'' as ѡ��,A.����,A.�ⷿID,A.��ҩ��λID,a.�������,A.ҩƷID,A.����, " & _
    "            A.No,A.���,b.����||'-'||b.���� As ��������,b.���,b.���㵥λ,A.����,E.���� As ��Ӧ��,A.����, " & _
    "           to_char(A.��������,'yyyy-mm-dd') as ��������, to_char(A.Ч��,'yyyy-mm-dd') as Ч��, " & _
    "           to_char(A.���Ч��,'yyyy-mm-dd') as ���Ч�� , " & _
    "           A.ʵ������  as �������,LTrim(To_Char(A.���ۼ�,'999999" & gSysPara.Price_Decimal.strFormt_VB & "'))  as ������ۼ�,A.���۽�� as ������۽��, " & _
    "           to_char(nvl(D.��������,0),'9999990.00000') as ���μ�������,D.��Ʒ����,D.�ڲ�����, " & _
    "           Decode(B.�Ƿ���,1,'ʱ��',LTrim(To_Char(C1.�ּ�,'999999" & gSysPara.Price_Decimal.strFormt_VB & "'))) as ����," & _
            IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, " To_Char(D.��������,'9999990.00000')", "Decode(Sign(D.��������),1,'��','��')") & " as ���," & _
    "           D.��������" & _
    " From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B, �������� C, �շѼ�Ŀ C1, ҩƷ��� D,��Ӧ�� E " & _
    " Where   a.����=15  " & strWhere & _
    "              And a.ҩƷID=C1.�շ�ϸĿID And (Sysdate Between C1.ִ������ and Nvl(C1.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')))" & _
    "              And a.��ҩ��λID=e.Id  " & _
    "              And a.ҩƷID=b.Id And  a.ҩƷID=C.����ID " & _
    "              And a.�ⷿID=D.�ⷿID And nvl(a.����,0)=nvl(D.����,0)    " & _
    "  Order By No,���"
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ⷿID, CStr(Trim(txtNO.Text)), _
        CDate(Format(dtpStart.Value, "yyyy-mm-dd")), CDate(Format(dtpEnd.Value, "yyyy-mm-dd")) + 1 - 1 / 24 / 60 / 60)
        
    With vsItem
        .Clear 0: .Cols = 1
        .FixedCols = 1
       Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 1 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            .ColData(i) = "0||1"
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "����" Or .ColKey(i) = "����" Or Trim(.ColKey(i)) = "��������" Then
                .ColHidden(i) = True
                ' ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                .ColData(i) = "-1||1"
            ElseIf .ColKey(i) = "NO" Or .ColKey(i) = "ѡ��" Or .ColKey(i) = "���μ�������" Then
                   .ColData(i) = "1||0"
                   If .ColKey(i) = "ѡ��" Then .ColDataType(i) = flexDTBoolean
                   .ColAlignment(i) = flexAlignCenterCenter
            End If
            If .ColKey(i) Like "*��*" Or .ColKey(i) Like "*��*" Or .ColKey(i) Like "*���*" Then
                 .ColAlignment(i) = flexAlignRightCenter
            End If
        Next
        '�Զ��п�
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsItem, Me.Caption, "����ѡ���б�", False
        If .ColIndex("��־") >= 0 Then .ColWidth(.ColIndex("��־")) = 300
        .Cell(flexcpBackColor, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = &HE7CFBA
        .Cell(flexcpBackColor, 1, .ColIndex("���μ�������"), .Rows - 1, .ColIndex("���μ�������")) = &HE7CFBA
    End With
    
    FillData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, Me.Name
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "����ѡ���б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    Dim strNO As String
    If Len(txtNO) < 8 And Len(txtNO) > 0 Then
        strNO = txtNO.Text
        Call MakeNO(68, mlng����ⷿID, strNO)
        txtNO.Text = strNO
    End If
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        With vsItem
            Select Case Col
            Case .ColIndex("���μ�������")
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "0.00000")
                If Val(.TextMatrix(Row, Col)) <> 0 And GetVsGridBoolColVal(vsItem, Row, .ColIndex("ѡ��")) = False Then
                    vsItem.TextMatrix(Row, .ColIndex("ѡ��")) = 1
                ElseIf Val(.TextMatrix(Row, Col)) = 0 Then
                    vsItem.TextMatrix(Row, .ColIndex("ѡ��")) = 0
                End If
            Case .ColIndex("ѡ��")
                If GetVsGridBoolColVal(vsItem, Row, Col) Then
                    If Val(.TextMatrix(Row, .ColIndex("���μ�������"))) = 0 Then
                            .TextMatrix(Row, .ColIndex("���μ�������")) = Format(Val(.TextMatrix(Row, .ColIndex("��������"))), "0.00000")
                    End If
                End If
            Case Else
            End Select
        End With
End Sub


Private Sub vsItem_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "����ѡ���б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "����ѡ���б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsItem, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "����ѡ���б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub
Private Sub picImg_Click()
    Call imgCol_Click
End Sub
Private Function zlBuliedRec() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ�е�����
    '����:���˺�
    '����:2010-12-16 15:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, blnData As Boolean
    '�ȼ��ɼ�������
    With vsItem
        blnData = False
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsItem, i, .ColIndex("ѡ��")) Then
                If Val(.TextMatrix(i, .ColIndex("���μ�������"))) <= 0 Then
                    MsgBox "ע��:" & "    �ڵ�" & i & "���еı��μ����������������,����!", vbOKOnly + vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("���μ�������")
                    If .RowIsVisible(.Row) = False Or .ColIsVisible(.Col) = False Then
                        Call .ShowCell(.Row, .Col)
                    End If
                    Exit Function
                End If
                If Val(.TextMatrix(i, .ColIndex("���μ�������"))) > Val(.TextMatrix(i, .ColIndex("��������"))) Then
                    MsgBox "ע��:" & "    �ڵ�" & i & "���еı��μ������������˿������,����!", vbOKOnly + vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("���μ�������")
                    If .RowIsVisible(.Row) = False Or .ColIsVisible(.Col) = False Then
                        Call .ShowCell(.Row, .Col)
                    End If
                    Exit Function
                End If
                blnData = True
            End If
        Next
    End With
    If blnData = False Then
        MsgBox "ע��:" & "    δѡ��ָ���ļ�������,����!", vbOKOnly + vbInformation, gstrSysName
        vsItem.SetFocus
        Exit Function
    End If
    Set mrsSel = New ADODB.Recordset
    With mrsSel
        If .State = 1 Then .Close
        .Fields.Append "����ⷿID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�շ���ĿID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���Ϲ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��Ʒ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ڲ�����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���ÿ��", adDouble, , adFldIsNullable
        .Fields.Append "����", adDouble, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    With vsItem
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsItem, i, .ColIndex("ѡ��")) Then
                mrsSel.AddNew
                mrsSel!����ⷿid = mlng����ⷿID
                mrsSel!�շ���ĿID = Val(.TextMatrix(i, .ColIndex("ҩƷID")))
                mrsSel!���� = Val(.TextMatrix(i, .ColIndex("����")))
                mrsSel!�������� = Trim(.TextMatrix(i, .ColIndex("��������")))
                mrsSel!���Ϲ�� = Trim(.TextMatrix(i, .ColIndex("���")))
                mrsSel!�ڲ����� = Trim(.TextMatrix(i, .ColIndex("�ڲ�����")))
                mrsSel!��Ʒ���� = Trim(.TextMatrix(i, .ColIndex("��Ʒ����")))
                mrsSel!���� = Trim(.TextMatrix(i, .ColIndex("���μ�������")))
                mrsSel!���ÿ�� = Val(.TextMatrix(i, .ColIndex("��������")))
                mrsSel.Update
            End If
        Next
    End With
    zlBuliedRec = True
End Function
Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem
        Select Case Col
        Case .ColIndex("���μ�������"), .ColIndex("ѡ��")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem
        Select Case Col
        Case .ColIndex("��־")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsItem_EnterCell()
    '��δ����
    With vsItem
    End With
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsItem
        If Val(.TextMatrix(.Row, .ColIndex("ҩƷID"))) = 0 And .Col >= .ColIndex("���μ�������") And .Row = .Rows - 1 Then
            gobjCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
End Sub

Private Sub vsItem_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsItem
        Select Case Col
        Case .ColIndex("���μ�������")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsItem
        Select Case .Col
            Case .ColIndex("���μ�������")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
            
        End Select
    End With
End Sub
Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    '������֤
    With vsItem
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("���μ�������")
                If zlDblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                strKey = Format(Val(strKey), "0.00000")
                If Val(strKey) > Val(.TextMatrix(Row, .ColIndex("��������"))) Then
                    MsgBox "ע��:" & vbCrLf & "    ��������������˿������,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                    Cancel = True: Exit Sub
                End If
                .EditText = strKey
                .TextMatrix(Row, .Col) = strKey
        End Select
    End With
End Sub
 
Private Sub MakeNO(ByVal intBillID As Integer, ByVal lng����id As Long, ByRef strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ�Ź����Զ���������
    '���:
    '����:strNo-���ص��ݺ�
    '����:
    '����:���˺�
    '����:2010-12-17 14:28:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intYear As Integer, strYear As String
    Dim intMonth As Integer, strMonth As String
    Dim str��� As String
    Dim rsTemp As New ADODB.Recordset
    
    strNO = UCase(LTrim(strNO))
    intYear = Format(gobjDatabase.Currentdate, "YYYY") - 1990
    strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
    intMonth = Month(gobjDatabase.Currentdate())
    strMonth = intMonth
    strMonth = String(2 - Len(strMonth), "0") & strMonth
    
    If rsTemp.State = 1 Then rsTemp.Close
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord("Select ��Ź��� From ������Ʊ� Where ��Ŀ���=[1]", "��ȡ���ݹ���", intBillID)
    
    
    Dim bln��� As Boolean
    Dim rsTmp As New ADODB.Recordset
    If Nvl(rsTemp!��Ź���, 0) = 2 And lng����id <> 0 Then
        gstrSQL = "Select ��������, ����id, ������� from ��������˵�� where �������� in ( '���Ŀ�','�Ƽ���','����ⷿ') and ����ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", lng����id)
        If rsTmp.EOF Then
            bln��� = True
        Else
            bln��� = False
        End If
    Else
        bln��� = False
    End If
    If Nvl(rsTemp!��Ź���, 0) = 0 Or bln��� Then
        If Len(strNO) < 8 Then strNO = strYear & String(7 - Len(strNO), "0") & strNO
    ElseIf rsTemp!��Ź��� = 2 Then
        If rsTemp.State = 1 Then rsTemp.Close
        Set rsTemp = gobjDatabase.OpenSQLRecord("Select ��� From  ���Һ���� Where ��Ŀ���=[1] and nvl(����ID,0)=[2]", "��ȡ���ұ��", intBillID, lng����id)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "��δ���ÿ��ұ�ţ��޷��������룡", vbInformation, gstrSysName
            Exit Sub
        End If
        If Nvl(rsTemp!���) = "" Then
            MsgBox "��δ���ÿ��ұ�ţ��޷��������룡", vbInformation, gstrSysName
            Exit Sub
        End If
        str��� = Nvl(rsTemp!���)
        
        'С����λ�������²�������
        '��λ����λ������Ϊ��ָ���·ݵĺ���
        '��λ������Ϊ�ǲ�������ָ�����ҡ��·ݵĺ���
        '���ڵ��ڰ�λ��������
        If Len(strNO) <= 4 Then
            strNO = strYear & str��� & strMonth & String(4 - Len(strNO), "0") & strNO
        ElseIf Len(strNO) <= 6 Then
            strNO = String(6 - Len(strNO), "0") & strNO
            strNO = strYear & str��� & strNO
        ElseIf Len(strNO) = 7 Then
            strNO = strYear & strNO
        End If
    Else
        MsgBox "��֧�����ֱ�Ź���", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub