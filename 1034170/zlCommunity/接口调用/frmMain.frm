VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gobjCommunity As Object

Public Function GetColItem(colInfo As Collection, strItem As String) As String
    If colInfo Is Nothing Then Exit Function
    
    Err.Clear: On Error Resume Next
    GetColItem = colInfo("_" & strItem)
    Err.Clear: On Error GoTo 0
End Function

Private Sub Command1_Click()
    Dim int社区 As Integer
    Dim str社区号 As String
    Dim colInfo As Collection
    
    '功能：身份验证
    If Not gobjCommunity Is Nothing Then
        If gobjCommunity.Identify(100, 1111, int社区, str社区号, colInfo) Then
            Me.Command1.Caption = GetColItem(colInfo, "姓名") & "," & GetColItem(colInfo, "性别") & "," & GetColItem(colInfo, "年龄")
        End If
    End If
End Sub

Private Sub Form_Load()
    '初始化
    On Error Resume Next
    Set gobjCommunity = CreateObject("zlCommunity.clsCommunity")
    Err.Clear: On Error GoTo 0
    If Not gobjCommunity Is Nothing Then
        If Not gobjCommunity.Initialize(gcnOracle) Then
            Set gobjCommunity = Nothing
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止化
    If Not gobjCommunity Is Nothing Then
        Call gobjCommunity.Terminate
        Set gobjCommunity = Nothing
    End If
End Sub
