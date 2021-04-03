VERSION 5.00
Begin VB.Form frmUpdateInfo 
   Caption         =   "修改队列信息"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   Icon            =   "frmUpdateInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4365
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txt医生姓名 
         Height          =   350
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txt诊室 
         Height          =   350
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txt患者姓名 
         Height          =   350
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboQueueName 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "医生姓名"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "诊室 "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "患者姓名"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "队列名称"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   350
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   1100
   End
End
Attribute VB_Name = "frmUpdateInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr队列名称 As String
Private mstr患者姓名 As String
Private mstr诊室 As String
Private mstr医生姓名 As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub LoadQueueName(ByRef astr队列名称() As String)
    
End Sub

Public Function zlShowMe(frmParent As Form, ByRef astr队列名称() As String, ByRef str队列名称 As String, str患者姓名 As String, _
            ByRef str诊室 As String, ByRef str医生姓名 As String) As Boolean
    Dim i As Integer
    
    mstr队列名称 = str队列名称
    mstr患者姓名 = str患者姓名
    mstr诊室 = str诊室
    mstr医生姓名 = str医生姓名
    
    On Error GoTo err
    
    cboQueueName.Clear
    
    If SafeArrayGetDim(astr队列名称) <> 0 Then
        For i = 1 To UBound(astr队列名称)
            cboQueueName.AddItem astr队列名称(i)
            If astr队列名称(i) = str队列名称 Then cboQueueName.ListIndex = i - 1
        Next i
        
        If cboQueueName.ListIndex = -1 Then Exit Function
        
        txt患者姓名 = mstr患者姓名
        txt医生姓名 = mstr医生姓名
        txt诊室 = mstr诊室
        
        Me.Show 1, frmParent
        
        If mstr队列名称 <> str队列名称 Or mstr患者姓名 <> str患者姓名 Or _
            mstr医生姓名 <> str医生姓名 Or mstr诊室 <> str诊室 Then
            str队列名称 = mstr队列名称
            str患者姓名 = mstr患者姓名
            str医生姓名 = mstr医生姓名
            str诊室 = mstr诊室
            zlShowMe = True
        End If
    End If
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    mstr队列名称 = cboQueueName.Text
    mstr患者姓名 = txt患者姓名.Text
    mstr医生姓名 = txt医生姓名.Text
    mstr诊室 = txt诊室.Text
    
    Unload Me
End Sub

