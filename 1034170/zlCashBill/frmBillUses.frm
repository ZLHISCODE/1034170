VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillUses 
   Caption         =   "Ʊ����ϸ"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   Icon            =   "frmBillUses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11805
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   105
      ScaleHeight     =   390
      ScaleWidth      =   8055
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   135
      Width           =   8055
      Begin VB.PictureBox picTimeRange 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   2145
         ScaleHeight     =   390
         ScaleWidth      =   6285
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   -30
         Visible         =   0   'False
         Width           =   6285
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "��ȡ����(&R)"
            Height          =   350
            Left            =   4680
            TabIndex        =   4
            Top             =   45
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   300
            Left            =   0
            TabIndex        =   2
            Top             =   60
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   127401987
            CurrentDate     =   41520
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   2550
            TabIndex        =   3
            Top             =   60
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   127401987
            CurrentDate     =   41520
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2265
            TabIndex        =   21
            Top             =   120
            Width           =   225
         End
      End
      Begin VB.ComboBox cboʹ������ 
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   30
         Width           =   1350
      End
      Begin VB.Label lblʹ������ 
         AutoSize        =   -1  'True
         Caption         =   "ʹ��ʱ��"
         Height          =   180
         Left            =   0
         TabIndex        =   0
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.Frame fraCMD 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   10320
      TabIndex        =   18
      Top             =   240
      Width           =   1455
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1200
      End
      Begin VB.CommandButton cmdDistant 
         Caption         =   "��λ�Ϻ�(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1200
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   150
         TabIndex        =   12
         Top             =   2700
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "��λƱ��(&F)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   13
         Top             =   3060
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "ȫ���˶�(&A)"
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "ȫ��ȡ��(&R)"
         Height          =   350
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   2430
         Width           =   630
      End
      Begin VB.Line linBlack 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   1300
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H00FFFFFF&
         X1              =   150
         X2              =   1300
         Y1              =   1815
         Y2              =   1815
      End
   End
   Begin VB.TextBox txtInput 
      Height          =   270
      Left            =   7520
      MaxLength       =   200
      TabIndex        =   17
      Top             =   740
      Visible         =   0   'False
      Width           =   1450
   End
   Begin VB.ComboBox cboResult 
      Height          =   300
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   6735
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorSel    =   12320767
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483644
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "  ����  |   ʹ��ʱ��   |ʹ����|    ʹ�����    |    �˶�ʱ��    |�˶���|   �˶Խ��  |      ��ע     |ID"
      MouseIcon       =   "frmBillUses.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Label lbl��ʾ 
      Caption         =   "ʹ����ϸ"
      Height          =   180
      Index           =   0
      Left            =   8160
      TabIndex        =   6
      Top             =   255
      Width           =   6210
   End
End
Attribute VB_Name = "frmBillUses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytInFun As Byte    '0-�鿴Ʊ��ʹ����ϸ,1-�˶�Ʊ����ϸ
Private mblnViewCheck As Boolean '��mbytInFun=0ʱ,�Ƿ���ʾ�˶�����ֶ�
Private mlngƱ�� As Long
Private mlng����ID As Long
Private mdblGiveCount As Double   '������Ʊ��������
Private mstrǰ׺�ı� As String
Private mblnUnClick As Boolean

Private Enum Col
    C0���� = 0
    C1ʹ��ʱ�� = 1
    C2ʹ���� = 2
    C3ʹ����� = 3
    C4�˶�ʱ�� = 4
    C5�˶��� = 5
    C6�˶Խ�� = 6
    C7��ע = 7
    C8ID = 8
End Enum
Private Sub SetUnChecked(ByVal lngRow As Long)
    With mshDetail
        .TextMatrix(lngRow, Col.C4�˶�ʱ��) = ""
        .TextMatrix(lngRow, Col.C5�˶���) = ""
        .TextMatrix(lngRow, Col.C6�˶Խ��) = ""
        .TextMatrix(lngRow, Col.C7��ע) = ""
        
        .RowData(lngRow) = 1  '���ڱ���ʱ�ж�
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub SetChecked(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strContent As String, Optional ByVal strDate As String)
    With mshDetail
        If lngCol = Col.C6�˶Խ�� Then
            .TextMatrix(lngRow, Col.C4�˶�ʱ��) = strDate
            .TextMatrix(lngRow, Col.C5�˶���) = UserInfo.����
            .TextMatrix(lngRow, lngCol) = strContent
        ElseIf lngCol = Col.C7��ע Then
            .TextMatrix(lngRow, lngCol) = strContent
        End If
        
        .RowData(lngRow) = 1 '���ڱ���ʱ�ж�
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub cboResult_LostFocus()
    If cboResult.Visible Then cboResult.Visible = False
End Sub
Private Sub RefreshCustomTime()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ��ʱ���������
    '����:���˺�
    '����:2013-11-01 10:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngCount As Long, dtDate As Date
    
    On Error GoTo errHandle
    If cboʹ������.Text <> "ʱ�䷶Χ" Then Exit Sub
    lngCount = 0
    With mshDetail
        For i = 1 To .Rows - 1
             'ָ��ʱ���
             If IsDate(.TextMatrix(i, Col.C1ʹ��ʱ��)) Then
                 dtDate = CDate(.TextMatrix(i, Col.C1ʹ��ʱ��))
                 If dtDate >= dtpStartDate.Value And dtDate <= dtpEndDate.Value Then
                     .RowHeight(i) = .RowHeight(0)
                 Else
                     .RowHeight(i) = 0
                 End If
             Else
                 .RowHeight(i) = 0
             End If
            If .RowHeight(i) <> 0 Then
                lngCount = lngCount + 1
            End If
        Next
    End With
    lbl��ʾ(0).Caption = lbl��ʾ(0).Tag & "���е�ǰѡ��" & lngCount & "��Ʊ��"
     Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub cboʹ������_Click()
    Dim i As Long, lngCount As Long, dtDate As Date
    If mblnUnClick = True Then Exit Sub
    On Error GoTo errHandle
    '����:29885
    picTimeRange.Visible = False
    With mshDetail
        For i = 1 To .Rows - 1
            If cboʹ������.Text = "����" Then
                .RowHeight(i) = .RowHeight(0)
            ElseIf cboʹ������.Text = "ʱ�䷶Χ" Then
                picTimeRange.Visible = True
                If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
                Call Form_Resize
                Exit Sub
            ElseIf InStr(1, .TextMatrix(i, Col.C1ʹ��ʱ��), cboʹ������.Text) > 0 Then
                .RowHeight(i) = .RowHeight(0)
            Else
                .RowHeight(i) = 0
            End If
            If .RowHeight(i) <> 0 Then
                lngCount = lngCount + 1
            End If
        Next
    End With
    If cboʹ������.Text <> "����" Then
        lbl��ʾ(0).Caption = lbl��ʾ(0).Tag & "���е�ǰѡ��" & lngCount & "��Ʊ��"
    Else
        lbl��ʾ(0).Caption = lbl��ʾ(0).Tag
    End If
    Call Form_Resize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboʹ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdAllDO_Click(Index As Integer)
    Dim i As Long, strDate As String
    Dim blnSel As Boolean '�Ƿ���ڶ���ѡ��
    Dim lngRows As Long
    Dim lngStart As Long
    
    With mshDetail
        blnSel = .Row <> .RowSel And .RowSel > .Row
        
        If Index = 0 Then
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        lngStart = IIf(blnSel, .Row, 1)
        lngRows = IIf(blnSel, .RowSel, .Rows - 1)
        For i = lngStart To lngRows
            
            If .RowHeight(i) <> 0 Then
                If Index = 0 Then
                   '��ʹ�Ѻ˶Ե�Ҳ���º˶�,��д�µĺ˶��˺ͺ˶�ʱ��,���ע,��ǰ���˵�Ҳ�������
                   Call SetChecked(i, Col.C6�˶Խ��, .TextMatrix(i, Col.C3ʹ�����), strDate)
                Else
                    'û�к˶Թ���,����ȡ���˶�
                    If Trim(.TextMatrix(i, Col.C6�˶Խ��)) <> "" Then Call SetUnChecked(i)
                End If
            End If
            
        Next
    End With
End Sub

Private Sub cboResult_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = vbKeyReturn Then
        With mshDetail
            If cboResult.ListIndex <= 0 Then
                Call SetUnChecked(.Row)
            Else
                Call SetChecked(.Row, Col.C6�˶Խ��, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
            End If
            .SetFocus    '����lostfocus
            .Col = .Col + 1
        End With
    ElseIf KeyAscii >= 32 Then
        If Chr(KeyAscii) > 5 Or Chr(KeyAscii) < 0 Then Exit Sub
        lngIdx = zlControl.CboMatchIndex(cboResult.hwnd, KeyAscii, 0.008)
        If lngIdx = -1 And cboResult.ListCount > 0 And cboResult.ListIndex = -1 Then lngIdx = 0
        cboResult.ListIndex = lngIdx
    End If
End Sub

Private Function SaveData() As Boolean
    Dim i As Long, arrSQL As Variant, blnTrans As Boolean, bytAllChecked As Byte, bytAllCheckOK As Byte
    Dim strDate As String, lngGiveCount As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    With mshDetail
        arrSQL = Array()
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                strDate = Trim(.TextMatrix(i, Col.C4�˶�ʱ��))
                If strDate = "" Then
                    strDate = "Null"
                Else
                    strDate = "To_Date('" & .TextMatrix(i, Col.C4�˶�ʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                arrSQL(UBound(arrSQL)) = "zl_Ʊ��ʹ����ϸ_check(" & .TextMatrix(i, Col.C8ID) & "," & ZVal(Val(.TextMatrix(i, Col.C6�˶Խ��))) & _
                                        ",'" & .TextMatrix(i, Col.C5�˶���) & "','" & .TextMatrix(i, Col.C7��ע) & "'," & strDate & ")"
            End If
        Next
    End With
    
    On Error GoTo errH
    If UBound(arrSQL) >= 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        
        '����Ƿ���Ҫ��д����˶Լ�¼
        strSQL = "Select Nvl(Sum(Decode(�˶Խ��, Null, 1, 0)), 0) As δ�˶���, Count(Distinct ����) As ��ʹ����" & vbNewLine & _
                "From Ʊ��ʹ����ϸ" & vbNewLine & _
                "Where ����id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If rsTmp!δ�˶��� = 0 And rsTmp!��ʹ���� = mdblGiveCount Then
            bytAllChecked = 1
            strSQL = "Select Count(ID) ������� From Ʊ��ʹ����ϸ Where ����id = [1] And �˶Խ�� <> ԭ��"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            If rsTmp!������� = 0 Then bytAllCheckOK = 1
        End If
                
        If bytAllChecked = 1 Then
            strSQL = "zl_Ʊ�����ü�¼_check(" & mlng����ID & "," & bytAllCheckOK & ",'" & UserInfo.���� & "',Null,1)"
        Else
            'ȡ������˶�
            strSQL = "zl_Ʊ�����ü�¼_check(" & mlng����ID & ",Null,Null,Null,Null)"
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
       gcnOracle.CommitTrans: blnTrans = False
    End If
    SaveData = True
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
      Call RefreshCustomTime
End Sub

Private Sub cmdSave_Click()
    Dim i As Long
    
    If SaveData Then
        With mshDetail
            For i = 1 To .Rows - 1
                If .RowData(i) = 1 Then .RowData(i) = 0
            Next
        End With
        cmdSave.Enabled = False
    End If
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub
Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Activate()
    If mshDetail.Rows > 1 Then Call SetRow(1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

 

Private Sub mshDetail_Click()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        Select Case .Col
            Case Col.C6�˶Խ��
                If .TextMatrix(.Row, .Col) <> "" Then
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, .Col)))
                Else
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, Col.C3ʹ�����)))
                End If
                Call SetCboResult
            Case Else
        End Select
    End With
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        With mshDetail
            Select Case .Col
                Case Col.C6�˶Խ��
                    Call SetUnChecked(.Row)
                Case Col.C7��ע
                    Call SetChecked(.Row, Col.C7��ע, "")
                Case Else
                
            End Select
        End With
    End If
End Sub

Private Sub mshDetail_KeyPress(KeyAscii As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            If .Row = .Rows - 1 And (.Col = Col.C7��ע Or .Col = Col.C6�˶Խ�� And .TextMatrix(.Row, .Col) = "") Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If .Col = Col.C7��ע Then
                    .Row = .Row + 1
                    .Col = Col.C6�˶Խ��
                Else
                    .Col = .Col + 1
                End If
            End If
        Else
            Select Case .Col
                Case Col.C6�˶Խ��
                    Call SetCboResult
                    Call cboResult_KeyPress(KeyAscii)
                Case Col.C7��ע
                    If .TextMatrix(.Row, Col.C6�˶Խ��) <> "" Then
                        txtInput.Text = Chr(KeyAscii)
                        txtInput.SelStart = 2
                        Call SetTxtInput
                    End If
                Case Else
                
            End Select
        End If
    End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
            MsgBox "���ֻ��������" & txtInput.MaxLength & "���ַ�!", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(1, txtInput.Text, "'") > 0 Then
            'MsgBox "ע��:��������ϵͳ��ֹ����������ַ�!", vbInformation, gstrSysName
            Beep
            Beep
            Exit Sub
        End If
        
        With mshDetail
            Call SetChecked(.Row, Col.C7��ע, Trim(txtInput.Text))
            txtInput.Visible = False
            .SetFocus  '����lostfocus
            If .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
                .Col = Col.C6�˶Խ��
            End If
        End With
    End If
End Sub

Private Sub cmdDistant_Click()
    Dim lngRow As Long, bln���� As Boolean
    Dim lngǰ׺ As Long
    
    MousePointer = vbHourglass
    lngǰ׺ = Len(mstrǰ׺�ı�) + 1
    With mshDetail
        lngRow = .Row + 1
        
        While True
            If lngRow > .Rows - 1 Then
                '���һ��
                If bln���� = False Then
                    If .Row = 1 Then
                        MsgBox "����δ���ֶϺ������", vbInformation, gstrSysName
                        MousePointer = vbDefault
                        Exit Sub
                    Else
                        If MsgBox("����δ���ֶϺŵ�������Ƿ��ͷ��ʼ��", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    bln���� = True
                    lngRow = 1
                Else
                    MsgBox "����δ���ֶϺ������", vbInformation, gstrSysName
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
            If lngRow > 1 Then
                If Val(Mid(.TextMatrix(lngRow - 1, 0), lngǰ׺)) < Val(Mid(.TextMatrix(lngRow, 0), lngǰ׺)) - 1 Then
                    '���ֶϺ�
                    If .RowHeight(lngRow) = 0 Then
                        If MsgBox("ע��:" & vbCrLf & "   �Ѿ����ҵ��˶Ϻ�,�����ڵ�ǰʱ�䷶Χ��,�Ƿ���ж�λ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                             If cboʹ������.Visible Then cboʹ������.ListIndex = 0:
                        Else
                            Exit Sub
                        End If
                    End If
                    Call SetRow(lngRow)
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            lngRow = lngRow + 1
        Wend
     End With
End Sub

Private Sub cmdFind_Click()
'����ָ������
    Dim strFind As String
    Dim lngRow As Long
    
    If txt����.Text = "" Then Exit Sub
    If Len(txt����.Text) > Len(mshDetail.TextMatrix(1, 0)) Then Exit Sub
    
    '�ѳ��Ȳ���
    strFind = UCase(Mid(mshDetail.TextMatrix(1, 0), 1, Len(mshDetail.TextMatrix(1, 0)) - Len(txt����.Text)) & txt����.Text)
    With mshDetail
        For lngRow = 1 To mshDetail.Rows - 1
            If mshDetail.TextMatrix(lngRow, 0) = strFind Then
                If .RowHeight(lngRow) = 0 Then
                    If MsgBox("ע��:" & vbCrLf & "   �������ҵĺ��벻�ڵ�ǰʱ�䷶Χ��,�Ƿ�Ҫ���ж�λ!", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                         If cboʹ������.Visible Then cboʹ������.ListIndex = 0:
                    Else
                        Exit Sub
                    End If
                End If
                Call SetRow(lngRow)
                Exit Sub
            End If
        Next
    End With
    MsgBox "δ�ҵ�����Ϊ " & strFind & " ��ʹ�ü�¼��", vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    If mbytInFun = 1 Then
        Call SaveData
    End If
    Unload Me
End Sub

Private Sub SetHeader()
    Dim strHead As String, arrTmp As Variant, i As Long
    
    With mshDetail
        If mbytInFun = 0 Then
            .SelectionMode = flexSelectionByRow
            .Row = 0: .Col = 0: .RowSel = 0: .ColSel = .Cols - 1
        Else
            .SelectionMode = flexSelectionFree
            .BackColorSel = &HE7CFBA
        End If
                
        If mbytInFun = 0 And Not mblnViewCheck Then
            strHead = "����,1,1000|ʹ��ʱ��,1,1800|ʹ����,4,800|ʹ�����,1,1000"
        Else
            strHead = "����,1,1000|ʹ��ʱ��,1,1800|ʹ����,4,800|ʹ�����,1,1000|�˶�ʱ��,1,1800|�˶���,4,800|�˶Խ��,1,1000|��ע,1,2000|ID,1,0"
        End If
        arrTmp = Split(strHead, "|")
        
        .Cols = UBound(arrTmp) + 1
        For i = 0 To UBound(arrTmp)
            .TextMatrix(0, i) = Split(arrTmp(i), ",")(0)
            .ColAlignment(i) = Split(arrTmp(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(arrTmp(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
    End With
End Sub

Private Sub Form_Load()

    If mbytInFun = 0 And Not mblnViewCheck Then Me.Width = 7000
    RestoreWinState Me, App.ProductName
    
    Me.Caption = IIf(mbytInFun = 0, "Ʊ����ϸ�嵥", "�˶�Ʊ����ϸ")
    Call SetHeader
    Call RestoreFlexState(mshDetail, Me.Caption)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call SaveFlexState(mshDetail, Me.Caption)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picFilter.Width = cboʹ������.Left + cboʹ������.Width + IIf(picTimeRange.Visible, picTimeRange.Width, 50) + 50
    lbl��ʾ(0).Top = IIf(picTimeRange.Visible, picFilter.Height + picFilter.Top + 50, picFilter.Top + (picFilter.Height - lbl��ʾ(0).Height) \ 2)
    lbl��ʾ(0).Left = IIf(picTimeRange.Visible, picFilter.Left, picFilter.Left + picFilter.Width + 50)
    If picTimeRange.Visible Then
        mshDetail.Top = lbl��ʾ(0).Height + lbl��ʾ(0).Top + 50
    Else
        mshDetail.Top = picFilter.Height + picFilter.Top + 50
    End If
    mshDetail.Height = Me.ScaleHeight - mshDetail.Top - 120
    If Me.ScaleWidth > 3000 Then
        fraCMD.Left = Me.ScaleWidth - fraCMD.Width - 120
        mshDetail.Width = fraCMD.Left - mshDetail.Left - 120
    End If
End Sub


Public Sub ShowMe(ByVal frmOwner As Form, ByVal bytInFun As Byte, ByVal blnViewCheck As Boolean, ByVal blnNOMoved As Boolean, _
    ByVal lngƱ�� As Long, ByVal lng����ID As Long, ByVal strǰ׺ As String, _
    Optional strCondition As String, Optional lngԭ�� As Long, Optional lng���� As Long, Optional strʹ���� As String, Optional str��ʾ As String)
    '����:bytInFun:0-�鿴Ʊ����ϸ,1-�˶�Ʊ����ϸ
    '   blnViewCheck:��bytInFun=0ʱ,�Ƿ���ʾ�˶�����ֶ�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, strNOs As String, strResult As String, arrTmp As Variant
    Dim strMinDate As String, strMaxDate As String
    mbytInFun = bytInFun
    mblnViewCheck = blnViewCheck
    mlngƱ�� = lngƱ��
    mlng����ID = lng����ID
    mstrǰ׺�ı� = strǰ׺
    
    strSQL = "Select ����, To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��, ʹ����," & vbNewLine & _
            "       Decode(ԭ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-�ش򷢳�', 4, '4-�ش��ջ�', '5-����') As ʹ�����," & vbNewLine & _
            "       To_Char(�˶�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �˶�ʱ��, �˶���, Decode(�˶Խ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-�ش򷢳�', 4, '4-�ش��ջ�', 5,'5-����','') as �˶Խ��, ��ע, ID" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ" & vbNewLine & _
            "Where ����id = [1] " & strCondition & vbNewLine & _
            "Order By ����"
    If mbytInFun = 0 And Not mblnViewCheck Then
        strSQL = "Select ����, To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��, ʹ����," & vbNewLine & _
            "       Decode(ԭ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-�ش򷢳�', 4, '4-�ش��ջ�', '5-����') As ʹ�����" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ" & vbNewLine & _
            "Where ����id = [1] " & strCondition & vbNewLine & _
            "Order By ����"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngԭ��, lng����, strʹ����)
    
    Dim strTemp As String, strʹ������ As String
    
    '��ʵ,���û��ʹ����ϸ,�˵��ѽ���,������ô˹���
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If InStr(1, strNOs & ",", "," & rsTmp!���� & ",") = 0 Then strNOs = strNOs & "," & rsTmp!����
            strTemp = "|" & Format(rsTmp!ʹ��ʱ��, "yyyy-MM-DD")
            If InStr(1, strʹ������ & "|", strTemp & "|") = 0 Then strʹ������ = strʹ������ & strTemp
            rsTmp.MoveNext
        Next
        i = 0
        If strNOs <> "" Then
            strNOs = Mid(strNOs, 2)
            i = UBound(Split(strNOs, ",")) + 1
        End If
        lbl��ʾ(0).Caption = str��ʾ & IIf(str��ʾ = "", "", ",") & "����" & i & "��Ʊ��."
        lbl��ʾ(0).Tag = lbl��ʾ(0).Caption
    End If
    Set mshDetail.DataSource = rsTmp
    
    Dim varData As Variant
    Dim j As Long
    If strʹ������ <> "" Then strʹ������ = Mid(strʹ������, 2)
    varData = Split(strʹ������, "|")
    
    mblnUnClick = True
    '��������С��������
    cboʹ������.AddItem "����": cboʹ������.ListIndex = cboʹ������.NewIndex
    cboʹ������.AddItem "ʱ�䷶Χ"
    mblnUnClick = False
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            For j = i + 1 To UBound(varData)
                If varData(j) < varData(i) Then
                    strTemp = varData(i)
                    varData(i) = varData(j)
                     varData(j) = strTemp
                End If
            Next
            If varData(i) < strMinDate Or strMinDate = "" Then strMinDate = varData(i)
            If varData(i) > strMaxDate Then strMaxDate = varData(i)
            cboʹ������.AddItem varData(i)
        End If
    Next
    dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpEndDate.MaxDate = dtpStartDate.MaxDate
    If strMinDate <> "" And IsDate(strMinDate) Then
        dtpStartDate.MinDate = Format(CDate(strMinDate), "yyyy-mm-dd 00:00:00")
        dtpStartDate.Value = dtpStartDate.MinDate
        dtpEndDate.MinDate = dtpStartDate.MinDate
        If IsDate(strMaxDate) Then
            dtpEndDate.Value = Format(CDate(strMaxDate), "yyyy-mm-dd 23:59:59")
        Else
            dtpEndDate.Value = Format(dtpStartDate.MinDate, "yyyy-mm-dd 23:59:59")
        End If
    End If
    cboResult.Visible = False
    txtInput.Visible = False
    If mbytInFun = 0 Then
        cmdOK.Caption = "�˳�(&X)"
        cmdOK.Cancel = True
        cmdCancel.Visible = False
        cmdSave.Visible = False
        cmdAllDO(0).Visible = False
        cmdAllDO(1).Visible = False
        picFilter.Visible = False
        lbl��ʾ(0).Left = picFilter.Left
    Else
        strResult = " ,1-����ʹ��,2-�����ջ�,3-�ش򷢳�,4-�ش��ջ�,5-����"
        arrTmp = Split(strResult, ",")
        For i = 0 To UBound(arrTmp)
            cboResult.AddItem arrTmp(i)
        Next
        Call zlControl.CboSetWidth(cboResult.hwnd, 800)
        
        mdblGiveCount = 0
        strSQL = "Select To_Number(Replace(��ֹ����, ǰ׺�ı�)) - To_Number(Replace(��ʼ����, ǰ׺�ı�))+1 ����" & vbNewLine & _
                "From Ʊ�����ü�¼" & vbNewLine & _
                "Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If rsTmp.RecordCount > 0 Then mdblGiveCount = rsTmp!����
        picFilter.Visible = True
    End If
    frmBillUses.Show vbModal, frmOwner
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub mshDetail_DblClick()
    Dim strReportNO As String, strInvoiceNO As String
    
    With mshDetail
        Select Case .Col
            Case Col.C7��ע
                If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
                If .TextMatrix(.Row, Col.C6�˶Խ��) = "" Then Exit Sub
                
                Call SetTxtInput
                txtInput.Text = .TextMatrix(.Row, .Col)
                Call SelAll(txtInput)
            Case Else
                strReportNO = "ZL" & glngSys \ 100 & "_INSIDE_1501"
                strInvoiceNO = .TextMatrix(.Row, Col.C0����)
                Call ReportOpen(gcnOracle, glngSys, strReportNO, Me, "Ʊ�ݺ�=" & strInvoiceNO & "", "Ʊ��=" & mlngƱ��, "ReportFormat=" & mlngƱ��, 1)
        End Select
    End With
End Sub

Private Sub SetCboResult()
    With mshDetail
        cboResult.Left = .Left + .CellLeft - 15
        cboResult.Top = .Top + .CellTop - 15
        cboResult.Width = .CellWidth + 15
        cboResult.Visible = True
        cboResult.SetFocus
    End With
End Sub

Private Sub SetTxtInput()
    With mshDetail
        txtInput.Left = .Left + .CellLeft - 15
        txtInput.Top = .Top + .CellTop - 15
        txtInput.Width = .CellWidth + 15
        txtInput.Height = .CellHeight
        txtInput.Visible = True
        txtInput.SetFocus
    End With
End Sub

Private Sub mshDetail_LeaveCell()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If cboResult.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(cboResult.Text) Then
                If cboResult.ListIndex <= 0 Then
                    Call SetUnChecked(.Row)
                Else
                    Call SetChecked(.Row, Col.C6�˶Խ��, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
                End If
            End If
        ElseIf txtInput.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(txtInput.Text) Then
                If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
                    MsgBox "���ֻ��������" & txtInput.MaxLength & "���ַ�!", vbInformation, gstrSysName
                    Exit Sub
                End If
                If InStr(1, txtInput.Text, "'") > 0 Then
                    'MsgBox "ע��:��������ϵͳ��ֹ����������ַ�!", vbInformation, gstrSysName
                    Beep
                    Beep
                    Exit Sub
                End If
                Call SetChecked(.Row, Col.C7��ע, Trim(txtInput.Text))
            End If
        End If
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngCol As Long
    
    With mshDetail
        If Button = 1 And .MousePointer = 99 Then
            lngCol = .MouseCol
            If .TextMatrix(0, lngCol) = "" Then Exit Sub
            
            .ColData(lngCol) = (.ColData(lngCol) + 1) Mod 2
            
            .Redraw = False
            .Col = lngCol: .ColSel = lngCol   '��������
            .Sort = IIf(.ColData(lngCol) = 1, 6, 5)
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
        End If
    End With
End Sub


Private Sub txtInput_LostFocus()
    If txtInput.Visible Then txtInput.Visible = False
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
        SelAll txt����
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub SetRow(ByVal lngRow As Long)
    Dim lngTop As Long
    With mshDetail
        .Row = lngRow
        lngTop = lngRow - 1
        If lngTop < 1 Then lngTop = 1
        If .RowIsVisible(lngTop) = False Then
            .TopRow = lngTop
        End If
        If mbytInFun = 0 Then
            .Col = 0
            .ColSel = .Cols - 1
        Else
            .Col = Col.C6�˶Խ��
        End If
    End With
End Sub


