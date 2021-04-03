VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPathDefinition 
   Caption         =   "查看路径表定义"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10050
   Icon            =   "frmPathDefinition.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleMode       =   0  'User
   ScaleWidth      =   7893.961
   StartUpPosition =   1  '所有者中心
   Begin XtremeSuiteControls.TabControl tbcPath 
      Height          =   3090
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      _Version        =   589884
      _ExtentX        =   9657
      _ExtentY        =   5450
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPathDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object
Private mlng路径ID As Long
Private mfrmPath As frmPathDesign

Public Sub ShowMe(frmParent As Object, ByVal lng路径ID As Long)
    mlng路径ID = lng路径ID
    Set mfrmParent = frmParent
    
    Me.Show 0, mfrmParent
End Sub

Private Sub Form_Load()
       
    If Me.WindowState = 1 Then Me.WindowState = 0
    Set mfrmPath = New frmPathDesign
    'TabControl
    '-----------------------------------------------------
    With Me.tbcPath
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        .InsertItem 0, "病人临床路径", mfrmPath.hWnd, 0
    End With
    
    Call mfrmPath.zlRefresh(mlng路径ID, "")
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    Me.tbcPath.Left = 0
    Me.tbcPath.Top = 0
    Me.tbcPath.Width = Me.ScaleWidth
    Me.tbcPath.Height = Me.ScaleHeight
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    Unload mfrmPath
    Set mfrmPath = Nothing
End Sub


