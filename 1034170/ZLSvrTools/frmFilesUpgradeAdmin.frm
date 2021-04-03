VERSION 5.00
Begin VB.Form frmFilesUpgradeAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ͻ��˹����û���������"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   Icon            =   "frmFilesUpgradeAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1155
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2850
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Left            =   1155
      TabIndex        =   0
      Text            =   "Administrator"
      Top             =   255
      Width           =   2850
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   1785
      TabIndex        =   2
      Top             =   1200
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   2895
      TabIndex        =   4
      Top             =   1200
      Width           =   1100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   225
      TabIndex        =   5
      Top             =   795
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����û�"
      Height          =   180
      Left            =   225
      TabIndex        =   3
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmFilesUpgradeAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean

'�ر�
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmd����_Click()
    Dim strUser As String
    Dim strPass As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strUser = txtUser.Text
    strPass = cipher(txtPass.Text)
    
    '������½��˺�
    gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ = '����Ա�˺�'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set ����='" & strUser & "' Where ��Ŀ='����Ա�˺�'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('����Ա�˺�',Null,'" & strUser & "')"
        gcnOracle.Execute strSQL
    End If
    
    '������½�����
    gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ = '����Ա����'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set ����='" & strPass & "' Where ��Ŀ='����Ա����'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('����Ա����',Null,'" & strPass & "')"
        gcnOracle.Execute strSQL
    End If
    
    mblnOK = True
    Unload Me
  Exit Sub
errHand:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub Form_Load()
    
    '��ȡ����Ա�û���������
    Call LoadReadAdmin
End Sub

Private Sub Form_Resize()
    With cmd����
        .Top = cmd����.Top
        .Left = cmdCancel.Left - .Width - 30
    End With
End Sub

Private Sub LoadReadAdmin()
    Dim rsTmp As New ADODB.Recordset
    Dim rsTmpPass As New ADODB.Recordset
    gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ like '����Ա�˺�'"
    Call OpenRecordset(rsTmp, gstrSQL, "����")
    
    If rsTmp.RecordCount = 1 Then
        txtUser.Text = Trim(Nvl(rsTmp!����))
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ like '����Ա����'"
        Call OpenRecordset(rsTmpPass, gstrSQL, Me.Caption)
        If rsTmpPass.RecordCount = 1 Then
            txtPass.Text = decipher(Trim(Nvl(rsTmpPass!����)))
        Else
            txtPass.Text = ""
        End If
    Else
        txtUser.Text = "Administrator"
        txtPass.Text = ""
    End If
    
End Sub


'������ܳ���
Private Function cipher(stext As String)
    Const min_asc = 32
    Const max_asc = 126
    Const num_asc = max_asc - min_asc + 1
    Dim offset As Long
    Dim strlen As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim ptext As String
    offset = 123
    Rnd (-1)
    Randomize (offset)
    strlen = Len(stext)
    For i = 1 To strlen
       ch = Asc(Mid(stext, i, 1))
       If ch >= min_asc And ch <= max_asc Then
           ch = ch - min_asc
           offset = Int((num_asc + 1) * Rnd())
           ch = ((ch + offset) Mod num_asc)
           ch = ch + min_asc
           ptext = ptext & Chr(ch)
       End If
    Next i
    cipher = ptext
End Function

'���ܳ���
Private Function decipher(stext As String)      '������ܳ���
    Const min_asc = 32 '��СASCII��
    Const max_asc = 126 '���ASCII�� �ַ�
    Const num_asc = max_asc - min_asc + 1
    Dim offset As Long
    Dim strlen As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim ptext As String
    offset = 123
    Rnd (-1)
    Randomize (offset)
    strlen = Len(stext)
    For i = 1 To strlen
       ch = Asc(Mid(stext, i, 1)) 'ȡ��ĸת���ASCII��
       If ch >= min_asc And ch <= max_asc Then
           ch = ch - min_asc
           offset = Int((num_asc + 1) * Rnd())
           ch = ((ch - offset) Mod num_asc)
           If ch < 0 Then
               ch = ch + num_asc
           End If
           ch = ch + min_asc
           ptext = ptext & Chr(ch)
       End If
    Next i
    decipher = ptext
End Function
