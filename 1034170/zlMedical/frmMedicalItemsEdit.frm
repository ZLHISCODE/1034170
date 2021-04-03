VERSION 5.00
Begin VB.Form frmMedicalItemsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������"
   ClientHeight    =   5685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8865
   Icon            =   "frmMedicalItemsEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7635
      TabIndex        =   6
      Top             =   1380
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7635
      TabIndex        =   5
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7635
      TabIndex        =   4
      Top             =   45
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   555
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   165
         Width           =   6405
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   270
         Left            =   7170
         TabIndex        =   1
         Top             =   180
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ(&D)"
         Height          =   180
         Index           =   12
         Left            =   75
         TabIndex        =   3
         Top             =   225
         Width           =   630
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5160
      Left            =   15
      TabIndex        =   7
      Top             =   465
      Width           =   7515
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   1
         Left            =   6675
         Picture         =   "frmMedicalItemsEdit.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "����ƶ�"
         Top             =   150
         Width           =   345
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   0
         Left            =   7080
         Picture         =   "frmMedicalItemsEdit.frx":0159
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "��ǰ�ƶ�"
         Top             =   150
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   4545
         Left            =   75
         TabIndex        =   8
         Top             =   525
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   8017
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����Ŀ(&M)"
         Height          =   180
         Index           =   14
         Left            =   420
         TabIndex        =   11
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   0
         Left            =   75
         Picture         =   "frmMedicalItemsEdit.frx":02A6
         Top             =   195
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMedicalItemsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mstrName As String
Private Enum mCol
    ������ = 1
    ����
    Ӣ����
    ����
    ����
    С��
    ��λ
    
End Enum

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
            
    txt.Locked = False
    cmd.Enabled = True
    
    If vData = False Then
        cmdOK.Tag = ""
    Else
        cmdOK.Tag = "Changed"
        txt.Locked = True
        cmd.Enabled = False
    End If
End Property

Private Property Get EditChanged() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
            
    EditChanged = (cmdOK.Tag = "Changed")
    
End Property


Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
            
    mlngKey = lngKey
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If mlngKey > 0 Then
        Call ReadData(mlngKey)
    End If
            
    EditChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT ���� FROM ������ĿĿ¼ WHERE ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then txt.Text = rs("����").Value
    mstrName = txt.Text
    
    
    gstrSQL = "Select a.ID,a.����,a.������,a.Ӣ����,A.����,A.����,A.С��,A.��λ " & _
                    "From ����������Ŀ a,����Ԫ��Ŀ¼ b,���������� c where b.����=-1 and  b.id=c.Ԫ��id and c.������id=a.id and c.��=[1] Order By c.�ؼ��� "
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.������) = zlCommFun.NVL(rs("������"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.Ӣ����) = zlCommFun.NVL(rs("Ӣ����"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.С��) = zlCommFun.NVL(rs("С��"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.��λ) = zlCommFun.NVL(rs("��λ"))
                        
            rs.MoveNext
        Loop
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand

    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "������", 2700, 1, "...", 1
        .NewColumn "����", 1200, 1
        .NewColumn "Ӣ����", 900, 1
        .NewColumn "����", 900, 1
        .NewColumn "����", 600, 7
        .NewColumn "С��", 600, 7
        .NewColumn "��λ", 900, 1
        
        .Body.ColHidden(mCol.����) = True
        
        .FixedCols = 1
        
        .SelectMode = True
    End With
        
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim lngElementID As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_��������_DELETE(" & mlngKey & ")"
    
    gstrSQL = "Select ID From ����Ԫ��Ŀ¼ Where ����=-1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        lngElementID = rs("ID").Value
    Else
        lngElementID = zlDatabase.GetNextId("����Ԫ��Ŀ¼")
        strSQL(ReDimArray(strSQL)) = "ZL_����Ԫ��_INSERT(-1," & lngElementID & ",'000000','������Ӧ','','����,9',1,Null,'00001')"
        
    End If
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            gstrSQL = "ZL_������_SAVE("
            gstrSQL = gstrSQL & lngElementID & ","
            gstrSQL = gstrSQL & lngLoop & ","
            gstrSQL = gstrSQL & "'2',"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & mlngKey & ","
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & Val(vsf.RowData(lngLoop)) & ","
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "'" & vsf.TextMatrix(lngLoop, mCol.��λ) & "',"                   '��λ
            gstrSQL = gstrSQL & "NULL)"
            
            strSQL(ReDimArray(strSQL)) = gstrSQL
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        If vsf.Row > 1 Then
            
            Call MoveItem(vsf.Row, -1)
            vsf.Row = vsf.Row - 1
            cmdOK.Tag = "Changed"
            
        End If
    ElseIf vsf.Row < vsf.Rows - 1 Then
        
        Call MoveItem(vsf.Row, 1)
        vsf.Row = vsf.Row + 1
        cmdOK.Tag = "Changed"
        
    End If
    
    vsf.ShowCell vsf.Row, vsf.Col
    vsf.SetFocus
End Sub

Private Function MoveItem(ByVal intCurRow As Integer, Optional ByVal intMove As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim intCol As Integer
    
    On Error GoTo errHand
    
    strTmp = CStr(vsf.RowData(intCurRow))
            
    vsf.RowData(intCurRow) = vsf.RowData(intCurRow + intMove)
    vsf.RowData(intCurRow + intMove) = Val(strTmp)
    
    For intCol = 0 To vsf.Cols - 1
        
        strTmp = vsf.TextMatrix(intCurRow, intCol)
        
        vsf.TextMatrix(vsf.Row, intCol) = vsf.TextMatrix(intCurRow + intMove, intCol)
        
        vsf.TextMatrix(intCurRow + intMove, intCol) = strTmp
        
    Next
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
       
    If EditChanged Then
    
        If SaveEdit Then
            mblnOK = True
            
            EditChanged = False
        End If
        
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Select Case Col
        Case mCol.������
            
            
            gstrSQL = "Select ID,�ϼ�id,0 As ĩ��,���� ,���� ,'' As Ӣ����,0 as ����,0 As ����,0 As С��,'' As ��λ from ������������ where ����=4 Start With �ϼ�id is null connect by prior id =�ϼ�id "
            
            gstrSQL = gstrSQL & " Union All Select A.ID,A.����id As �ϼ�id,1 As ĩ��,A.����,A.������ As ����,a.Ӣ����,A.����,A.����,A.С��,A.��λ " & _
                    "From ����������Ŀ A  "
                    
            gstrSQL = "Select * From (" & gstrSQL & ") A ORDER BY A.ĩ��, A.����"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            If ShowGrdSelect(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;Ӣ����,900,0,0;����,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100) Then
                
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
                
                vsf.EditText = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.��λ) = zlCommFun.NVL(rs("��λ").Value)
                vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.С��) = zlCommFun.NVL(rs("С��").Value)
                vsf.TextMatrix(Row, mCol.������) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.Ӣ����) = zlCommFun.NVL(rs("Ӣ����").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                EditChanged = True
                
            End If

    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    Dim rsData As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.������
                    
                    gstrSQL = "Select A.ID,1 As ĩ��,A.����,A.������ As ����,a.Ӣ����,A.����,A.����,A.С��,A.��λ " & _
                    "From ����������Ŀ A Where ����id In (Select ID from ������������ where ����=4) And (���� Like [1] Or Upper(������) Like [2] Or Upper(Ӣ����) Like [2])"
                               
                    strText = UCase(vsf.EditText) & "%"
                    If ParamInfo.��Ŀ����ƥ�䷽ʽ = 0 Then strTmp = "%" & strText
                                
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
                    If ShowGrdFilter(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;Ӣ����,900,0,0;����,900,0,0", Me.Name & "\�����Ŀ����", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100) Then
                        
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                           
                        vsf.EditText = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.��λ) = zlCommFun.NVL(rs("��λ").Value)
                        vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.С��) = zlCommFun.NVL(rs("С��").Value)
                        vsf.TextMatrix(Row, mCol.������) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.Ӣ����) = zlCommFun.NVL(rs("Ӣ����").Value)
                        vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        EditChanged = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                        vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        EditChanged = True
    End If
End Sub