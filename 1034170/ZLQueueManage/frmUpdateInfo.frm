VERSION 5.00
Begin VB.Form frmUpdateInfo 
   Caption         =   "�޸Ķ�����Ϣ"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   Icon            =   "frmUpdateInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4365
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtҽ������ 
         Height          =   350
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txt���� 
         Height          =   350
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txt�������� 
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
         Caption         =   "ҽ������"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "���� "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "��������"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "��������"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
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

Private mstr�������� As String
Private mstr�������� As String
Private mstr���� As String
Private mstrҽ������ As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub LoadQueueName(ByRef astr��������() As String)
    
End Sub

Public Function zlShowMe(frmParent As Form, ByRef astr��������() As String, ByRef str�������� As String, str�������� As String, _
            ByRef str���� As String, ByRef strҽ������ As String) As Boolean
    Dim i As Integer
    
    mstr�������� = str��������
    mstr�������� = str��������
    mstr���� = str����
    mstrҽ������ = strҽ������
    
    On Error GoTo err
    
    cboQueueName.Clear
    
    If SafeArrayGetDim(astr��������) <> 0 Then
        For i = 1 To UBound(astr��������)
            cboQueueName.AddItem astr��������(i)
            If astr��������(i) = str�������� Then cboQueueName.ListIndex = i - 1
        Next i
        
        If cboQueueName.ListIndex = -1 Then Exit Function
        
        txt�������� = mstr��������
        txtҽ������ = mstrҽ������
        txt���� = mstr����
        
        Me.Show 1, frmParent
        
        If mstr�������� <> str�������� Or mstr�������� <> str�������� Or _
            mstrҽ������ <> strҽ������ Or mstr���� <> str���� Then
            str�������� = mstr��������
            str�������� = mstr��������
            strҽ������ = mstrҽ������
            str���� = mstr����
            zlShowMe = True
        End If
    End If
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    mstr�������� = cboQueueName.Text
    mstr�������� = txt��������.Text
    mstrҽ������ = txtҽ������.Text
    mstr���� = txt����.Text
    
    Unload Me
End Sub

