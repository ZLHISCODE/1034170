VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.1#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSet���� 
   AutoRedraw      =   -1  'True
   Caption         =   "����"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7650
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   30
      ScaleHeight     =   375
      ScaleWidth      =   7575
      TabIndex        =   7
      Top             =   3960
      Width           =   7575
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "1"
         Top             =   30
         Width           =   360
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "������(&K)"
         Height          =   255
         Left            =   2220
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CheckBox chkʵʱ 
         Caption         =   "������ϸʱʵ�ϴ�(&M)"
         Height          =   285
         Index           =   0
         Left            =   3450
         TabIndex        =   9
         Top             =   45
         Width           =   2085
      End
      Begin VB.CheckBox chkʵʱ 
         Caption         =   "סԺ��ϸʱʵ�ϴ�(&Z)"
         Height          =   285
         Index           =   1
         Left            =   5520
         TabIndex        =   8
         Top             =   45
         Width           =   2085
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����(&D)"
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   90
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ŵ���"
         Height          =   180
         Index           =   4
         Left            =   1500
         TabIndex        =   12
         Top             =   90
         Width           =   540
      End
   End
   Begin ZL9BillEdit.BillEdit mshBill 
      Height          =   2685
      Left            =   120
      TabIndex        =   2
      Top             =   1170
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   4736
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   4380
      Width           =   7755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   4
      Top             =   4590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5100
      TabIndex        =   3
      Top             =   4590
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   690
      Width           =   7665
   End
   Begin MSComctlLib.TabStrip tabSel 
      Height          =   3105
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   5477
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmSet����.frx":000C
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "�����豸�Ĵ��ںż����ô����Ƿ�Ĭ��Ϊ������,�������շ���������ص�ҽ����Ŀ���Ӧ"
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   390
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlngҽ������ As Long
Private mlng���� As Long
Private Enum mColHead
    �շ���� = 0
    ������Ŀ
    ������Ŀ
End Enum
Private Sub chk������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}", 1
    End If
End Sub


Private Sub chkʵʱ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}", 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    
    If Trim(txtEdit) = "" Then Exit Sub
    SaveRegInFor g����ģ��, "����", "�˿ں�", Me.txtEdit
    SaveRegInFor g����ģ��, "����", "������", Me.chk������.Value
    If Val(txtEdit) = 0 Then
        gintComPort_���� = 1
    Else
        gintComPort_���� = Val(txtEdit)
    End If
    gblnKFQCom_���� = IIf(chk������.Value = 1, True, False)
    gintComPort = txtEdit.Text
        
    'ɾ���Ѿ�����
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",NUll)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    With MshBill
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, mColHead.�շ����) <> "" Then
                '������������
                gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'" & .TextMatrix(lngRow, mColHead.�շ����) & "' ,'" & .TextMatrix(lngRow, mColHead.������Ŀ) & ";" & .TextMatrix(lngRow, mColHead.������Ŀ) & "'," & lngRow + 2 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    '����
    
    ' gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'������ϸʱʵ�ϴ�' ,'" & IIf(chkʵʱ(0).Value = 1, "1", "0") & "'," & 1 & ")"
     
     gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'������ϸʱʵ�ϴ�' ,'" & IIf(chkʵʱ(0).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'     gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'סԺ��ϸʱʵ�ϴ�' ,'" & IIf(chkʵʱ(1).Value = 1, "1", "0") & "'," & 2 & ")"
    gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'סԺ��ϸʱʵ�ϴ�' ,'" & IIf(chkʵʱ(1).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    mblnReturn = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    Resume
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    mblnReturn = False
    
    Call GetRegInFor(g����ģ��, "����", "�˿ں�", strReg)
    
    If Val(strReg) = 0 Then
        txtEdit.Text = 1
    Else
        txtEdit.Text = Val(strReg)
    End If
    
    Call GetRegInFor(g����ģ��, "����", "������", strReg)
    
    If Val(strReg) = 1 Then
        Me.chk������.Value = 1
    Else
        Me.chk������.Value = 0
    End If
    RestoreWinState Me, App.ProductName
    
    '��ʼ����
    Call iniData
End Sub

Public Function ShowME(ByVal lng���� As Long, ByVal lngҽ������ As Long) As Boolean
    mlngҽ������ = lngҽ������
    mlng���� = lng����
    
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub Form_Resize()
   Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = 0
    sngBottom = 0
    
    fra(0).Width = ScaleWidth + 50
    With cmdCancel
        .Top = ScaleHeight - .Height - 100
        .Left = ScaleWidth - .Width - 50
    End With
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - 50 - .Width
    End With
    
    fra(1).Width = fra(0).Width
    fra(1).Top = cmdOK.Top - fra(1).Height - 50
    
    With pic
        .Top = fra(1).Top - .Height - 50
        .Width = ScaleWidth - 50
    End With
    With tabSel
        .Width = ScaleWidth - 50
        .Height = pic.Top - .Top - 20
    End With
    With MshBill
        .Top = tabSel.Top + tabSel.Tabs(1).Height + 100
        .Left = tabSel.Left + 100
        .Height = tabSel.Height - tabSel.Tabs(1).Height - 200
        .Width = tabSel.Width - 200
    End With
    MshBill.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshBill_EnterCell(Row As Long, Col As Long)
    With MshBill
        Select Case Col
            Case mColHead.������Ŀ
                MshBill.Clear
                MshBill.AddItem "����"
                MshBill.AddItem "��ҩ��"
                MshBill.AddItem "��ҩ��"
                MshBill.AddItem "��ҩ��"
                MshBill.AddItem "����"
                MshBill.AddItem "����"
                MshBill.AddItem "���Ʒ�"
                MshBill.AddItem "�������Ʒ�"
                If mlng���� = TYPE_���������� Then
                    'mshBill.AddItem "������"
                Else
                   ' mshBill.AddItem "������"
                End If
            Case mColHead.������Ŀ
                MshBill.Clear
                MshBill.AddItem "A�в�ҩ��"
                MshBill.AddItem "B�г�ҩ��"
                MshBill.AddItem "C��ҩ��"
                MshBill.AddItem "D����"
                MshBill.AddItem "E������"
                MshBill.AddItem "F�����"
                MshBill.AddItem "G������"
                MshBill.AddItem "H�����"
                MshBill.AddItem "I���Ʒ�"
                MshBill.AddItem "J�����"
                MshBill.AddItem "K��λ��"
                MshBill.AddItem "L�����"
                MshBill.AddItem "M��������"
        End Select
    End With
End Sub
Private Sub pic_Resize()
    Err = 0
    On Error Resume Next
    With chkʵʱ(1)
        .Left = pic.ScaleWidth - .Width - 50
        chkʵʱ(0).Left = .Left - chkʵʱ(1).Width - 50
    End With
End Sub

Private Sub tabSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
            SendKeys "{Tab}", 1
    End If
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}", 1
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
        zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m����ʽ
End Sub
Private Function iniData() As Boolean
    '��ʼ����
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    
    '����ҳͷ
    Err = 0
    On Error Resume Next
    strSql = "Select * from ��������Ŀ¼ where ����=" & mlng����
    zlDatabase.OpenRecordset rsTmp, strSql, Me.Caption
    If rsTmp.EOF Then
        tabSel.Tabs(1).Caption = "��"
    Else
        tabSel.Tabs(1).Caption = NVL(rsTmp!����)
    End If
    rsTmp.Close
  
    If mlng���� = TYPE_���������� Then
        Me.chk������.Value = 1
    Else
        Me.chk������.Value = 0
    End If
    
    '���ñ���ͷ
    Call initGrid
    strSql = "" & _
        "   Select A.���,b.����ֵ From �շ���� a,(Select * From ���ղ��� where ����=" & mlng���� & ") b " & _
        "   Where A.���=b.������(+) " & _
        "   order by A.���� "
    zlDatabase.OpenRecordset rsTmp, strSql, Me.Caption
    With MshBill
        .ClearBill
        If rsTmp.RecordCount = 0 Then
            .Rows = 2
        Else
            .Rows = rsTmp.RecordCount + 1
        End If
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, mColHead.�շ����) = NVL(rsTmp!���)
            strTmp = NVL(rsTmp!����ֵ)
            If InStr(1, strTmp, ";") <> 0 Then
                .TextMatrix(lngRow, mColHead.������Ŀ) = Split(strTmp, ";")(0)
                .TextMatrix(lngRow, mColHead.������Ŀ) = Split(strTmp, ";")(1)
            End If
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        
        strSql = "Select * From ���ղ��� where ������ in('������ϸʱʵ�ϴ�','סԺ��ϸʱʵ�ϴ�') and ����=" & mlng����
        zlDatabase.OpenRecordset rsTmp, strSql, Me.Caption
        chkʵʱ(0).Value = 1
        chkʵʱ(1).Value = 1
        Do While Not rsTmp.EOF
            Select Case NVL(rsTmp!������)
            Case "������ϸʱʵ�ϴ�"
                chkʵʱ(0).Value = IIf(Val(NVL(rsTmp!����ֵ)) = 1, 1, 0)
            Case "סԺ��ϸʱʵ�ϴ�"
                chkʵʱ(1).Value = IIf(Val(NVL(rsTmp!����ֵ)) = 1, 1, 0)
            End Select
            rsTmp.MoveNext
        Loop
        
    End With
    
End Function
Private Sub initGrid()
    With MshBill
        .Active = True
        .Cols = 3
        
        .msfObj.FixedCols = 1
        .AllowAddRow = False
        
        .TextMatrix(0, mColHead.�շ����) = "�շ����"
        .TextMatrix(0, mColHead.������Ŀ) = "������Ŀ"
        .TextMatrix(0, mColHead.������Ŀ) = "������Ŀ"
        
        
        .ColWidth(mColHead.�շ����) = 1500
        .ColWidth(mColHead.������Ŀ) = 2000
        .ColWidth(mColHead.������Ŀ) = 2000
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(mColHead.�շ����) = 5
        .ColData(mColHead.������Ŀ) = 3
        .ColData(mColHead.������Ŀ) = 3
        
        .ColAlignment(mColHead.�շ����) = flexAlignLeftCenter
        .ColAlignment(mColHead.������Ŀ) = flexAlignLeftCenter
        .ColAlignment(mColHead.������Ŀ) = flexAlignLeftCenter
        .PrimaryCol = mColHead.������Ŀ
        .LocateCol = mColHead.������Ŀ
    End With
End Sub



