VERSION 5.00
Begin VB.Form frmSet�ɶ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmSet�ɶ�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   90
      TabIndex        =   5
      Top             =   1230
      Width           =   4275
   End
   Begin VB.TextBox txtIC�˿ں� 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3810
      TabIndex        =   4
      Text            =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.ComboBox cbo������ 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   750
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   7
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1770
      TabIndex        =   6
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CheckBox Chk��Ժ��Ϣ 
      Caption         =   "��Ժ�Ǽǵ�ͬʱ���ϴ�ҽ��������Ժ��Ϣ(&1)"
      Height          =   345
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblIC�˿ں� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�˿ں�"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3180
      TabIndex        =   3
      Top             =   810
      Width           =   540
   End
   Begin VB.Label lbl������ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   1
      Top             =   810
      Width           =   540
   End
End
Attribute VB_Name = "frmSet�ɶ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnOK As Boolean

Public Function ShowSet() As Boolean
    blnOK = False
    
    Me.Show 1
    ShowSet = blnOK
End Function

Private Sub cbo������_Click()
    Me.lblIC�˿ں�.Enabled = (cbo������.ListIndex <> 0)
    Me.txtIC�˿ں�.Enabled = (cbo������.ListIndex <> 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "zl_���ղ���_Delete(" & gintInsure & ",NULL)", , adCmdStoredProc
    gcnOracle.Execute "zl_���ղ���_Insert(" & gintInsure & ",NULL,'�ϴ���Ժ��Ϣ'," & Chk��Ժ��Ϣ.Value & ",1)", , adCmdStoredProc
    gcnOracle.CommitTrans
    
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������", Me.cbo������.ListIndex)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", txtIC�˿ں�.Text)
    blnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    
    '���ӳ�ʼ������
    Me.cbo������.Clear
    Me.cbo������.AddItem "�ſ�"
    Me.cbo������.AddItem "IC��-JKP428"
    Me.cbo������.AddItem "IC��-ICIOX"
    Me.cbo������.ListIndex = 0
    
    '����ǰ�Ĳ���ȡ������ʾ�ڽ�����
    gstrSQL = "Select ������,Nvl(����ֵ,0) Value From ���ղ��� Where ����=22 "
    Call OpenRecordset(rsTmp, "��ȡ�ϴ���Ժ��Ϣ����ֵ")
    With rsTmp
        Do While Not rsTmp.EOF
            Select Case !������
            Case "�ϴ���Ժ��Ϣ"
                Chk��Ժ��Ϣ.Value = rsTmp!Value
            End Select
            .MoveNext
        Loop
    End With
    
    cbo������.ListIndex = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������", 0)
    txtIC�˿ں�.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", 1)
End Sub
