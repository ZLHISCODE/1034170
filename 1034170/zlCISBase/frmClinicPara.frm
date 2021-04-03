VERSION 5.00
Begin VB.Form frmClinicPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frmClinicPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraAddMode 
      Height          =   1365
      Left            =   150
      TabIndex        =   2
      Top             =   105
      Width           =   4155
      Begin VB.OptionButton opt增加模式 
         Caption         =   "连续增加(保存后自动增加项目)"
         Height          =   225
         Index           =   1
         Left            =   735
         TabIndex        =   5
         Top             =   810
         Width           =   3105
      End
      Begin VB.OptionButton opt增加模式 
         Caption         =   "单项增加(保存后关闭编辑)"
         Height          =   225
         Index           =   0
         Left            =   735
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   2580
      End
      Begin VB.Label lblAddMode 
         AutoSize        =   -1  'True
         Caption         =   "1、项目增加操作模式"
         Height          =   180
         Left            =   570
         TabIndex        =   3
         Top             =   15
         Width           =   1710
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   60
         Picture         =   "frmClinicPara.frx":000C
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4455
      TabIndex        =   0
      Top             =   390
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4455
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
End
Attribute VB_Name = "frmClinicPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowMe(ByVal frmParent As Object)
    Me.Show 1, frmParent
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If Me.opt增加模式(0).Value = True Then
        Call zlDatabase.SetPara("诊疗项目连续增加", 0, glngSys, 1054)
    Else
        Call zlDatabase.SetPara("诊疗项目连续增加", 1, glngSys, 1054)
    End If

    Unload Me
End Sub

Private Sub Form_Load()
    '根据用户权限，装入控件
    Dim lngValues As Long
    lngValues = Val(zlDatabase.GetPara("诊疗项目连续增加", glngSys, 1054, 0, Array(Me.opt增加模式(0), Me.opt增加模式(1)), True))
    
    If lngValues = 0 Then
        Me.opt增加模式(0).Value = True: Me.opt增加模式(1).Value = False
    Else
        Me.opt增加模式(0).Value = False: Me.opt增加模式(1).Value = True
    End If
End Sub
