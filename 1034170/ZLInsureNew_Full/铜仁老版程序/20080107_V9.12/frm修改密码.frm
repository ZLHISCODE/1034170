VERSION 5.00
Begin VB.Form frm修改密码 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改密码"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frm修改密码.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.TextBox txt确认新密码 
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.TextBox txt新密码 
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.TextBox txt原密码 
      BeginProperty Font 
         Name            =   "宋体"
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
      Picture         =   "frm修改密码.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lbl确认新密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "确认新密码(&V)"
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.Label lbl新密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "新密码(&N)"
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.Label lbl原密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "原密码(&O)"
      BeginProperty Font 
         Name            =   "宋体"
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
Attribute VB_Name = "frm修改密码"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr新密码 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txt新密码.Text) = "" Then
        MsgBox "密码不能为空！", vbInformation, gstrSysName
        txt新密码.SetFocus
        Exit Sub
    End If
    
    If txt新密码.Text <> txt确认新密码.Text Then
        MsgBox "两次输入的新密码不一致，请重输！", vbInformation, gstrSysName
        txt新密码.SetFocus
        Exit Sub
    End If
    mstr新密码 = txt新密码.Text
    
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
    txt原密码.Text = mstr新密码
    mstr新密码 = ""
End Sub

Private Sub txt确认新密码_GotFocus()
    txt确认新密码.SelStart = 0
    txt确认新密码.SelLength = 8
End Sub

Private Sub txt新密码_GotFocus()
    txt新密码.SelStart = 0
    txt新密码.SelLength = 8
End Sub

Private Sub txt原密码_Change()
    cmdOK.Enabled = (Len(txt原密码.Text) <> 0)
End Sub

Private Sub txt原密码_GotFocus()
    txt原密码.SelStart = 0
    txt原密码.SelLength = 8
End Sub

Public Function ChangePassword(ByVal strPass As String) As String
    mstr新密码 = strPass
    Me.Show 1
    ChangePassword = mstr新密码
End Function
