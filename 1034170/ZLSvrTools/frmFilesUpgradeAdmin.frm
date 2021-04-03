VERSION 5.00
Begin VB.Form frmFilesUpgradeAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "客户端管理用户密码设置"
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
   StartUpPosition =   2  '屏幕中心
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
   Begin VB.CommandButton cmd保存 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   1785
      TabIndex        =   2
      Top             =   1200
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   2895
      TabIndex        =   4
      Top             =   1200
      Width           =   1100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "管理密码"
      Height          =   180
      Left            =   225
      TabIndex        =   5
      Top             =   795
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管理用户"
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

'关闭
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmd保存_Click()
    Dim strUser As String
    Dim strPass As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strUser = txtUser.Text
    strPass = cipher(txtPass.Text)
    
    '保存或新建账号
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '管理员账号'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set 内容='" & strUser & "' Where 项目='管理员账号'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('管理员账号',Null,'" & strUser & "')"
        gcnOracle.Execute strSQL
    End If
    
    '保存或新建密码
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '管理员密码'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set 内容='" & strPass & "' Where 项目='管理员密码'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('管理员密码',Null,'" & strPass & "')"
        gcnOracle.Execute strSQL
    End If
    
    mblnOK = True
    Unload Me
  Exit Sub
errHand:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub Form_Load()
    
    '读取管理员用户名及密码
    Call LoadReadAdmin
End Sub

Private Sub Form_Resize()
    With cmd保存
        .Top = cmd保存.Top
        .Left = cmdCancel.Left - .Width - 30
    End With
End Sub

Private Sub LoadReadAdmin()
    Dim rsTmp As New ADODB.Recordset
    Dim rsTmpPass As New ADODB.Recordset
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 like '管理员账号'"
    Call OpenRecordset(rsTmp, gstrSQL, "管理")
    
    If rsTmp.RecordCount = 1 Then
        txtUser.Text = Trim(Nvl(rsTmp!内容))
        gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 like '管理员密码'"
        Call OpenRecordset(rsTmpPass, gstrSQL, Me.Caption)
        If rsTmpPass.RecordCount = 1 Then
            txtPass.Text = decipher(Trim(Nvl(rsTmpPass!内容)))
        Else
            txtPass.Text = ""
        End If
    Else
        txtUser.Text = "Administrator"
        txtPass.Text = ""
    End If
    
End Sub


'密码加密程序
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

'解密程序
Private Function decipher(stext As String)      '密码解密程序
    Const min_asc = 32 '最小ASCII码
    Const max_asc = 126 '最大ASCII码 字符
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
       ch = Asc(Mid(stext, i, 1)) '取字母转变成ASCII码
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
