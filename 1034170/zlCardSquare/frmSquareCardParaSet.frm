VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6792
   Icon            =   "frmSquareCardParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4884
   ScaleWidth      =   6792
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk������ֵ 
      Caption         =   "��ֵ���˳���ֵ����(&N)"
      Height          =   240
      Left            =   4215
      TabIndex        =   10
      Top             =   2700
      Width           =   2400
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   4020
      Width           =   9600
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   9600
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   4290
      TabIndex        =   7
      Top             =   3345
      Width           =   1500
   End
   Begin VB.Frame fra 
      Caption         =   "���ѿ�����"
      Height          =   2220
      Left            =   90
      TabIndex        =   4
      Top             =   195
      Width           =   6525
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   1770
         Left            =   60
         TabIndex        =   5
         Top             =   315
         Width           =   6390
         _cx             =   11271
         _cy             =   3122
         Appearance      =   1
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   350
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   0
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ӡ����"
      Height          =   1245
      Left            =   75
      TabIndex        =   2
      Top             =   2535
      Width           =   4005
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "Ʊ�ݴ�ӡ����"
         Height          =   360
         Left            =   555
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   705
         Width           =   1875
      End
      Begin VB.CheckBox chk�ɿ 
         Caption         =   "�ɿ�������ӡ�ɿ"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   270
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4215
      TabIndex        =   1
      Top             =   4290
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   0
      Top             =   4290
      Width           =   1100
   End
End
Attribute VB_Name = "frmSquareCardParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs   As String, mblnFirst As Boolean, mblnChange As Boolean
Private Sub InitSqure()
    Dim strColHead As String, varData As Variant, i As Long
    strColHead = "���ѿ��ӿ�����|���㷽ʽ|����ǰ׺�ı�|���ų���|����������ʾ"
     varData = Split(strColHead, "|")
    With vsCardList
            .Clear
            .Cols = UBound(varData) + 1
            For i = 0 To UBound(varData)
                .FixedAlignment(i) = flexAlignCenterCenter
                .TextMatrix(0, i) = varData(i)
                If varData(i) = "����������ʾ" Then
                    .ColDataType(i) = flexDTBoolean
                End If
                .ColKey(i) = varData(i)
            Next
    End With
End Sub
Private Function LoadCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ѿ���Ϣ
    '����:
    '����:���˺�
    '����:2009-12-15 11:29:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lngRow As Long
    ' ����,����,���㷽ʽ,nvl(���ƿ�,0)  as ���ƿ�,ǰ׺�ı�,���ų���
    
    On Error GoTo errHandle
    
    Set rsTemp = zlGet���ѿ��ӿ�
    rsTemp.Filter = "���ƿ�=1"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "���ѿ��ӿڲ�����,����!"
        Exit Function
    End If
    With vsCardList
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("���ѿ��ӿ�����")) = Nvl(rsTemp!���) & "-" & Nvl(rsTemp!����)
            .Cell(flexcpData, lngRow, .ColIndex("���ѿ��ӿ�����")) = Nvl(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Nvl(rsTemp!���㷽ʽ)
            .Cell(flexcpData, lngRow, .ColIndex("���㷽ʽ")) = .TextMatrix(lngRow, .ColIndex("���㷽ʽ"))
            .TextMatrix(lngRow, .ColIndex("����ǰ׺�ı�")) = Nvl(rsTemp!ǰ׺�ı�)
            .Cell(flexcpData, lngRow, .ColIndex("����ǰ׺�ı�")) = .TextMatrix(lngRow, .ColIndex("����ǰ׺�ı�"))
            .TextMatrix(lngRow, .ColIndex("���ų���")) = Nvl(rsTemp!���ų���, 20)
            .Cell(flexcpData, lngRow, .ColIndex("���ų���")) = .TextMatrix(lngRow, .ColIndex("���ų���"))
            .TextMatrix(lngRow, .ColIndex("����������ʾ")) = Val(Nvl(rsTemp!�Ƿ�����))
            .Cell(flexcpData, lngRow, .ColIndex("����������ʾ")) = Val(Nvl(rsTemp!�Ƿ�����))
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If InStr(1, mstrPrivs, ";��������;") > 0 Then
            gstrSQL = "Select distinct a.���㷽ʽ From ���㷽ʽӦ�� A,���㷽ʽ b Where a.Ӧ�ó��� in ('�շ�','����') AND a.���㷽ʽ=b.���� and b.����=8 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            .ColComboList(.ColIndex("���㷽ʽ")) = .BuildComboList(rsTemp, "���㷽ʽ", "���㷽ʽ")
            .Editable = flexEDKbdMouse
        End If
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    End With
    LoadCardInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub ShowParaSet(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:frmMain-������
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�

    '����:���˺�
    '����:2009-11-19 15:29:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mblnFirst = True
    Me.Show 1, frmMain
End Sub
Private Sub LoadParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�������
    '����:���˺�
    '����:2009-12-10 17:03:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, varData As Variant
    Dim blnIsHavePriv As Boolean
    blnIsHavePriv = InStr(1, mstrPrivs, ";��������;") > 0
    chk�ɿ.Value = IIf(Val(zlDatabase.GetPara("�ɿ��ӡ", glngSys, mlngModule, , Array(chk�ɿ), blnIsHavePriv)) = 1, 1, 0)
    chk������ֵ.Value = IIf(Val(zlDatabase.GetPara("������ֵ", glngSys, mlngModule, , Array(chk������ֵ), blnIsHavePriv)) = 1, 1, 0)
    Call LoadCardInfor
End Sub

 

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ü��
    '����:���˺�
    '����:2009-12-10 17:15:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsHavePriv As Boolean, lngLen As Long, rsTemp As ADODB.Recordset, strTemp As String
    Dim lngRow  As Long
    
    On Error GoTo errHandle
    
    blnIsHavePriv = InStr(1, mstrPrivs, ";��������;") > 0

    If blnIsHavePriv Then
        With vsCardList
            For lngRow = 1 To .Rows - 1
                strTemp = Trim(.TextMatrix(lngRow, .ColIndex("����ǰ׺�ı�")))
                lngLen = zlCommFun.ActualLen(Trim(strTemp)) + Val(.TextMatrix(lngRow, .ColIndex("���ų���")))
                If lngLen > 20 Then
                    ShowMsgbox "���ѿ��ŵ���󳤶�(ǰ׺+���ų���)���ܴ���20λ,����"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                If InStr(1, Trim(strTemp), "|") > 0 Then
                    ShowMsgbox "���ѿ��ŵ�ǰ׺�ı��в��ܰ���:��|,'��~;��,����"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                
                If InStr(1, Trim(strTemp), ",") > 0 Then
                    ShowMsgbox "���ѿ��ŵ�ǰ׺�ı��в��ܰ���:��|,'��~;��,����"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                If InStr(1, Trim(strTemp), ";") > 0 Then
                    ShowMsgbox "���ѿ��ŵ�ǰ׺�ı��в��ܰ���:��|,';��,����"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                
                If InStr(1, Trim(strTemp), "'") > 0 Then
                    ShowMsgbox "���ѿ��ŵ�ǰ׺�ı��в��ܰ���:��|,'��~;��,����"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                If InStr(1, Trim(strTemp), "��") > 0 Or InStr(1, Trim(strTemp), "~") > 0 Then
                    ShowMsgbox "���ѿ��ŵ�ǰ׺�ı��в��ܰ���:��|,'��~;��,����"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                        
                If Val(.Cell(flexcpData, lngRow, .ColIndex("���ų���"))) <> Val(.TextMatrix(lngRow, .ColIndex("���ų���"))) Or Len(.Cell(flexcpData, lngRow, .ColIndex("����ǰ׺�ı�"))) <> Len(strTemp) Then
                    '�����˸���,������Ҫ��鳤���Ƿ��С
                    gstrSQL = "Select 1 From ���ѿ�Ŀ¼ where ID>0 and rownum =1 and �ӿڱ��=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.Cell(flexcpData, lngRow, .ColIndex("���ѿ��ӿ�����"))))
                    If Not rsTemp.EOF Then
                        If lngLen < zlCommFun.ActualLen(Trim(.Cell(flexcpData, lngRow, .ColIndex("����ǰ׺�ı�")))) + Val(.Cell(flexcpData, lngRow, .ColIndex("���ų���"))) Then
                            ShowMsgbox "���ڷ����˷�����Ϣ,�������ѿ��ŵĲ��ܵ���,����"
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
    End If
    
    IsValied = True
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
        
End Function


Private Function SaveSet() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2009-12-10 16:59:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsHavePriv As Boolean, lng�ӿ���� As Long, lngRow As Long
    Dim cllPro As Collection, blnTrans As Boolean
    
    blnIsHavePriv = InStr(1, mstrPrivs, ";��������;") > 0
    Err = 0: On Error GoTo ErrHand:
   
    With vsCardList
        If blnIsHavePriv Then
            Set cllPro = New Collection
            For lngRow = 1 To .Rows - 1
                 lng�ӿ���� = Val(.Cell(flexcpData, lngRow, .ColIndex("���ѿ��ӿ�����")))
                 If lng�ӿ���� <> 0 Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("���㷽ʽ"))) <> Trim(.Cell(flexcpData, lngRow, .ColIndex("���㷽ʽ"))) Or _
                       Trim(.TextMatrix(lngRow, .ColIndex("����ǰ׺�ı�"))) <> Trim(.Cell(flexcpData, lngRow, .ColIndex("����ǰ׺�ı�"))) Or _
                       Val(.TextMatrix(lngRow, .ColIndex("���ų���"))) <> Val(.Cell(flexcpData, lngRow, .ColIndex("���ų���"))) Or Abs(Val(.TextMatrix(lngRow, .ColIndex("����������ʾ")))) <> Val(.Cell(flexcpData, lngRow, .ColIndex("����������ʾ"))) Then
                           'ֻ�з����˸ı�Ĳ��ܸ���
                           ' Zl_�����ѽӿ�Ŀ¼_Update
                           gstrSQL = "Zl_�����ѽӿ�Ŀ¼_Update("
                           '  ���_In     In �����ѽӿ�Ŀ¼.���%Type,
                           gstrSQL = gstrSQL & "" & lng�ӿ���� & ","
                           '  ���㷽ʽ_In In �����ѽӿ�Ŀ¼.���㷽ʽ%Type,
                           gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) & "',"
                           '  ����ǰ׺_In In �����ѽӿ�Ŀ¼.����ǰ׺%Type,
                           gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����ǰ׺�ı�")) & "',"
                           '  ���ų���_In In �����ѽӿ�Ŀ¼.���ų���%Type
                           gstrSQL = gstrSQL & " " & Val(.TextMatrix(lngRow, .ColIndex("���ų���"))) & ","
                           '    �Ƿ�����_In In �����ѽӿ�Ŀ¼.�Ƿ�����%Type := 0
                           gstrSQL = gstrSQL & " " & IIf(Abs(Val(.TextMatrix(lngRow, .ColIndex("����������ʾ")))) = 0, 0, 1) & ")"
                           zlAddArray cllPro, gstrSQL
                    End If
                 End If
            Next
        End If
    End With
    gcnOracle.BeginTrans
    blnTrans = True
    If Not cllPro Is Nothing Then
        If cllPro.Count > 0 Then zlExecuteProcedureArrAy cllPro, Me.Caption, True, blnTrans
    End If
    Call zlDatabase.SetPara("�ɿ��ӡ", IIf(chk�ɿ.Value = 1, 1, 0), glngSys, mlngModule, blnIsHavePriv)
    Call zlDatabase.SetPara("������ֵ", IIf(chk������ֵ.Value = 1, 1, 0), glngSys, mlngModule, blnIsHavePriv)
    gcnOracle.CommitTrans: blnTrans = False
    Set grsStatic.rs���ѿ��ӿ� = Nothing
    Call zlGet���ѿ��ӿ�
    SaveSet = True
    Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    Call ErrCenter
    SaveErrLog
End Function

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Sub CmdOK_Click()
    If IsValied = False Then Exit Sub
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_1503"
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadCardInfor = False Then Unload Me: Exit Sub
    Call LoadParaSet
End Sub

 

Private Sub Form_Load()
    Call InitSqure
End Sub

Private Sub vsCardList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCardList
        Select Case Col
        Case .ColIndex("���㷽ʽ")
        Case .ColIndex("����ǰ׺�ı�")
        Case .ColIndex("���ų���")
        Case .ColIndex("����������ʾ")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsCardList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsCardList
        Select Case Col
        Case .ColIndex("���㷽ʽ"), .ColIndex("����ǰ׺�ı�"), .ColIndex("���ų���"), .ColIndex("����������ʾ")
        Case Else
            Exit Sub
        End Select
    End With
End Sub
Private Sub vsCardList_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsCardList_EnterCell()
    
    With vsCardList
        .EditMaxLength = 0
        Select Case .Col
        Case .ColIndex("����ǰ׺�ı�")
            .EditMaxLength = 4
        Case .ColIndex("���ų���")
            .EditMaxLength = 3
        End Select
    End With
End Sub

Private Sub vsCardList_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsCardList
        Select Case Col
        Case .ColIndex("���㷽ʽ"), .ColIndex("����ǰ׺�ı�"), .ColIndex("���ų���")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsCardList, 0, .Cols - 1, False)
    End With
End Sub

Private Sub vsCardList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsCardList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    
    With vsCardList
        Select Case Col
        Case .ColIndex("����ǰ׺�ı�")
            If InStr(1, "'~��|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Case .ColIndex("���ų���")
            '��Ҫ���ܴ����˿����
            Call VsFlxGridCheckKeyPress(vsCardList, Row, Col, KeyAscii, m���ʽ)
        Case Else
        End Select
    End With
End Sub
 
