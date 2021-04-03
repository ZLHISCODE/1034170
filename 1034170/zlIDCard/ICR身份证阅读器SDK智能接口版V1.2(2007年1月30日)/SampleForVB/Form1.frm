VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   10710
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   0
      Top             =   120
   End
   Begin VB.CommandButton EndCmd 
      Caption         =   "返  回"
      Height          =   372
      Left            =   8520
      TabIndex        =   6
      Top             =   5520
      Width           =   1452
   End
   Begin VB.CommandButton SearchCardCmd 
      Caption         =   "找  卡"
      Height          =   372
      Left            =   8520
      TabIndex        =   5
      Top             =   3120
      Width           =   1452
   End
   Begin VB.CommandButton ReadCardCmd 
      Caption         =   "读  卡"
      Height          =   372
      Left            =   8520
      TabIndex        =   4
      Top             =   4320
      Width           =   1452
   End
   Begin VB.CommandButton SelectCmd 
      Caption         =   " 选  卡"
      Height          =   372
      Left            =   8520
      TabIndex        =   3
      Top             =   3720
      Width           =   1452
   End
   Begin VB.CommandButton ResetCMD 
      Caption         =   "SAM复位"
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton ChipSNCmd 
      Caption         =   "卡芯片序列号"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox RecText 
      BackColor       =   &H80000014&
      Height          =   5535
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   6855
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   8280
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   8280
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'进入
Private Sub Form_Load()

   '显示二进制返回码
   'RecText.Text = LTrim(tmp)

End Sub


