VERSION 5.00
Begin VB.Form frmSet���� 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdTrans 
      Caption         =   "�ϴ�"
      Height          =   350
      Left            =   90
      TabIndex        =   11
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CheckBox chk��λ 
      Caption         =   "�ϴ���λ��Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   10
      Top             =   1860
      Width           =   3375
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�ϴ�������Ŀ��Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   9
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CheckBox chkҩƷ 
      Caption         =   "�ϴ�ҩƷ������Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   8
      Top             =   1260
      Width           =   3375
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�ϴ�����������Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   7
      Top             =   960
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   750
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3810
      TabIndex        =   5
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2670
      TabIndex        =   4
      Top             =   2400
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2250
      Width           =   5265
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1545
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�Ŵ���"
      Height          =   180
      Index           =   4
      Left            =   1950
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ����(&D)"
      Height          =   180
      Index           =   3
      Left            =   450
      TabIndex        =   1
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng���� As Long
 
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtEdit) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call ExecuteProcedure(Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'�˿ں�','" & txtEdit.Text & "',1)"
    Call ExecuteProcedure(Me.Caption)
    
    gcnOracle.CommitTrans
    gintComPort = txtEdit.Text
    mblnReturn = True
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTrans_Click()
    Dim rsTemp As New ADODB.Recordset, iLoop As Long, strTemp As String
'    gstrҽ���������� = "500102"
'    gstrҽԺ���� = "5001020003"
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Sub
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
    If chk����.Value = 1 Then
        gstrSQL = "Select id as ����,���� From ���ղ���"
        Call OpenRecordset(rsTemp, gstrSysName)
        chk����.Caption = "�ϴ�����������Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstrҽ����������, gstrҽԺ����, "0", rsTemp!����, rsTemp!����, "", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk����.Caption = "�ϴ�����������Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk����.Value = 0
    End If
    If chkҩƷ.Value = 1 Then
        gstrSQL = "select a.��� as ���,a.id as ����,a.���� as ����,b.ҩƷ��Դ as ҩƷ��Դ from �շ�ϸĿ a,ҩƷĿ¼ b where a.��� In ('5','6','7') and a.����=b.����"
        Call OpenRecordset(rsTemp, gstrSysName)
        chkҩƷ.Caption = "�ϴ�ҩƷ������Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstrҽ����������, gstrҽԺ����, "1", rsTemp!��� & "_" & rsTemp!����, rsTemp!����, IIf(rsTemp!ҩƷ��Դ = "����", "03", "02"), gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chkҩƷ.Caption = "�ϴ�ҩƷ������Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chkҩƷ.Value = 0
    End If
    If chk����.Value = 1 Then
        gstrSQL = "select * from �շ�ϸĿ where ��� Not In ('J','5','6','7')"
        Call OpenRecordset(rsTemp, gstrSysName)
        chk����.Caption = "�ϴ�������Ŀ��Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstrҽ����������, gstrҽԺ����, "2", rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, "", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk����.Caption = "�ϴ�������Ŀ��Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk����.Value = 0
    End If
    If chk��λ.Value = 1 Then
        gstrSQL = "select * from �շ�ϸĿ where ���='J'"
        Call OpenRecordset(rsTemp, gstrSysName)
        chk��λ.Caption = "�ϴ���λ��Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstrҽ����������, gstrҽԺ����, "3", rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, " ", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk��λ.Caption = "�ϴ���λ��Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk��λ.Value = 0
    End If
    MsgBox "������Ŀ��Ϣ�ϴ����", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnReturn = False
    
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=" & mlng����
    Call OpenRecordset(rsTemp, "��ȡ����")
    
    If Not rsTemp.EOF Then txtEdit.Text = rsTemp!����ֵ
End Sub

Public Function ShowME(ByVal lng���� As Long) As Boolean
    mlng���� = lng����
    Me.Show 1
    ShowME = mblnReturn
End Function
