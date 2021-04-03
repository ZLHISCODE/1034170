VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmPassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2445
      TabIndex        =   3
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1215
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   570
      Width           =   2850
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   90
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   -15
      X2              =   5000
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -15
      X2              =   5000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���������룺"
      Height          =   180
      Left            =   1230
      TabIndex        =   0
      Top             =   300
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   300
      Picture         =   "frmPassword.frx":058A
      Top             =   345
      Width           =   720
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrPass As String
Private mbytPWDMin As Byte
Private mbytPWDMax As Byte

Public Function ShowMe(strPass As String, Optional ByVal bytPwdMin As Byte, Optional ByVal bytPwdMax As Byte) As Boolean
    strPass = ""
    lblInfo.Caption = ""
    mbytPWDMin = bytPwdMin
    mbytPWDMax = bytPwdMax
    Me.Show 1
    
    If mblnOK Then
        strPass = mstrPass
        ShowMe = True
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrPass = txtPass.Text
    If mbytPWDMin < mbytPWDMax Then
        If Len(mstrPass) < mbytPWDMin Then
            lblInfo.Caption = "���벻�ܵ���" & mbytPWDMin & "λ!" & "ʵ��¼�볤��Ϊ:" & Len(mstrPass)
            Exit Sub
        ElseIf Len(mstrPass) > mbytPWDMax Then
            lblInfo.Caption = "���벻�ܳ���" & mbytPWDMax & "λ!" & "ʵ��¼�볤��Ϊ:" & Len(mstrPass)
            Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    Call HookDefend(txtPass.hwnd)
    mstrPass = ""
End Sub

Private Sub txtPass_Change()
    If lblInfo.Caption <> "" Then lblInfo.Caption = ""
End Sub
