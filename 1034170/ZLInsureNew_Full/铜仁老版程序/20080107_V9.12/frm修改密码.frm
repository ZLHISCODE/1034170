VERSION 5.00
Begin VB.Form frm�޸����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�޸�����"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frm�޸�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1680
      TabIndex        =   7
      Top             =   1860
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3210
      TabIndex        =   8
      Top             =   1860
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -210
      TabIndex        =   6
      Top             =   1680
      Width           =   5865
   End
   Begin VB.TextBox txtȷ�������� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2370
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1170
      Width           =   2265
   End
   Begin VB.TextBox txt������ 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2370
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2265
   End
   Begin VB.TextBox txtԭ���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2370
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   270
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frm�޸�����.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblȷ�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ��������(&V)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   1230
      Width           =   1680
   End
   Begin VB.Label lbl������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&N)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1110
      TabIndex        =   2
      Top             =   780
      Width           =   1170
   End
   Begin VB.Label lblԭ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ԭ����(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1110
      TabIndex        =   0
      Top             =   330
      Width           =   1170
   End
End
Attribute VB_Name = "frm�޸�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr������ As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txt������.Text) = "" Then
        MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
        txt������.SetFocus
        Exit Sub
    End If
    
    If txt������.Text <> txtȷ��������.Text Then
        MsgBox "��������������벻һ�£������䣡", vbInformation, gstrSysName
        txt������.SetFocus
        Exit Sub
    End If
    mstr������ = txt������.Text
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    txtԭ����.Text = mstr������
    mstr������ = ""
End Sub

Private Sub txtȷ��������_GotFocus()
    txtȷ��������.SelStart = 0
    txtȷ��������.SelLength = 8
End Sub

Private Sub txt������_GotFocus()
    txt������.SelStart = 0
    txt������.SelLength = 8
End Sub

Private Sub txtԭ����_Change()
    cmdOK.Enabled = (Len(txtԭ����.Text) <> 0)
End Sub

Private Sub txtԭ����_GotFocus()
    txtԭ����.SelStart = 0
    txtԭ����.SelLength = 8
End Sub

Public Function ChangePassword(ByVal strPass As String) As String
    mstr������ = strPass
    Me.Show 1
    ChangePassword = mstr������
End Function
