VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ������"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   2485
      TabIndex        =   5
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   1095
      TabIndex        =   4
      Top             =   1755
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -97
      TabIndex        =   3
      Top             =   1500
      Width           =   4875
   End
   Begin VB.CommandButton cmdBrower 
      Caption         =   "���(&B)"
      Height          =   400
      Left            =   3400
      TabIndex        =   2
      Top             =   885
      Width           =   1100
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ָ���ļ����λ��"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng���� As Long, mblnReturn As Boolean

Public Function ShowMe(ByVal lng���� As Long) As Boolean
    mlng���� = lng����
    Me.Show 1
    ShowMe = mblnReturn
End Function

Private Sub cmdBrower_Click()
    txtPath.Text = BrowPath(Me.hwnd, "��ѡ���ļ����λ�ã�")
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(txtPath.Text) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call ExecuteProcedure(Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'�ļ����λ��','" & txtPath.Text & "',1)"
    Call ExecuteProcedure(Me.Caption)
    
    mstrSavePath = txtPath.Text
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Me.Hide
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        If rsTemp!������ = "�ļ����λ��" Then txtPath.Text = rsTemp!����ֵ
        rsTemp.MoveNext
    Loop
End Sub
