VERSION 5.00
Begin VB.Form frmErr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "错误提示"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "frmErr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txt 
      BackColor       =   &H8000000F&
      Height          =   885
      Left            =   900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
      Width           =   3255
   End
   Begin VB.CommandButton CDM确认 
      Caption         =   "是(&Y)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1080
      TabIndex        =   0
      Top             =   1470
      Width           =   1100
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "否(&N)"
      Height          =   350
      Left            =   2460
      TabIndex        =   1
      Top             =   1470
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "发生如下错误，是否需要继续？"
      Height          =   180
      Left            =   930
      TabIndex        =   2
      Top             =   150
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmErr.frx":000C
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msgReturn As VbMsgBoxResult

Private Sub CDM确认_Click()
    msgReturn = vbYes
    Unload Me
End Sub

Private Sub CMD放弃_Click()
    Unload Me
End Sub

Public Function ShowErr(ByVal strErr As String) As VbMsgBoxResult
    msgReturn = vbNo
    
    txt.Text = strErr
    frmErr.Show vbModal
    ShowErr = msgReturn
End Function
