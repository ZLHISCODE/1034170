VERSION 5.00
Begin VB.Form FrmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "示例"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6270
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox ChbIsPwd 
      Caption         =   "是否有密码器"
      Height          =   180
      Left            =   1200
      TabIndex        =   11
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox TxtAdditionData 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "发票操作"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "医疗服务门诊收费专用发票"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3255
      Begin VB.OptionButton Option3 
         Caption         =   "错票"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "废票"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "退票"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox TxtInvoice_NO 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "80000001"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "医疗服务收费专用发票"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label LblAdditionData 
      Caption         =   "原始票号:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "机打票号:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "FrmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tax1 As New Tax
  
Dim Invoice_Kind As Long
Dim Invoice_NO As String

Const S_Consumer_Name As String = "测试用消费者"
Const s_Oper_Name As String = "周兴华"

Dim InvoiceData As String
Dim errMessage As String


Dim Inv_Type As Long
Dim AdditionData As String

Dim ReturnValue As String
  
Private Sub Command1_Click()
On Error GoTo ErrDesc:
  Invoice_Kind = 1
  Invoice_NO = TxtInvoice_NO.Text
  InvoiceData = "2;0001;运输费;1;10;备注;0002;运输费;1;20;备注;住院号;支票号;"
  errMessage = String(255, " ")
  If ChbIsPwd.Value Then
    ReturnValue = Tax1.BJ_Normal_Invoice(Invoice_Kind, Invoice_NO, S_Consumer_Name, s_Oper_Name, InvoiceData, errMessage)
  Else
    ReturnValue = Tax1.BJ_Normal_Invoice_NoPwd(Invoice_Kind, Invoice_NO, S_Consumer_Name, s_Oper_Name, InvoiceData, errMessage)
  End If
  MsgBox "返回值：" + errMessage
  Exit Sub
ErrDesc:
  MsgBox Err.Description
End Sub

Private Sub Command2_Click()
On Error GoTo ErrDesc:
  Invoice_Kind = 2
  Invoice_NO = TxtInvoice_NO.Text
  InvoiceData = "1.00;0.01;0.02;0.03;0.04;0.05;0.06;0.07;0.08;0.09;0.10;0.11;0.12;0.13;0.14;0.15;0.16;0.17;0.18;" & _
                    "开发科技部;0.19;" & _
                    "开发科技部;0.20;" & _
                    "开发科技部;0.21;" & _
                    "开发科技部;0.22;" & _
                    "开发科技部;0.23;"

  errMessage = String(255, " ")
  If ChbIsPwd.Value Then
    ReturnValue = Tax1.BJ_Normal_Invoice(Invoice_Kind, Invoice_NO, S_Consumer_Name, s_Oper_Name, InvoiceData, errMessage)
  Else
    ReturnValue = Tax1.BJ_Normal_Invoice_NoPwd(Invoice_Kind, Invoice_NO, S_Consumer_Name, s_Oper_Name, InvoiceData, errMessage)
  End If
  MsgBox "返回值：" + errMessage
  Exit Sub
ErrDesc:
  MsgBox Err.Description
End Sub

Private Sub Command3_Click()
On Error GoTo ErrDesc:
  If Inv_Type = 1 Or Inv_Type = 2 Then
    AdditionData = ""
  Else
   AdditionData = Trim(TxtAdditionData.Text)
   If AdditionData = "" Then
     TxtAdditionData.SetFocus
     MsgBox "请输入原始票号!"
     Exit Sub
   End If
  End If
  errMessage = String(255, " ")
  'Tax1.test (errMessage)
  ReturnValue = Tax1.BJ_Other_Invoice(Inv_Type, 1, Invoice_NO, s_Oper_Name, AdditionData, errMessage)
  MsgBox "返回值：" + errMessage
  Exit Sub
ErrDesc:
  MsgBox Err.Description
End Sub

Private Sub Form_Load()
  Option1_Click
End Sub

Private Sub Option1_Click()
Inv_Type = 1
LblAdditionData.Visible = False
TxtAdditionData.Visible = False
End Sub

Private Sub Option2_Click()
Inv_Type = 2
LblAdditionData.Visible = False
TxtAdditionData.Visible = False
End Sub

Private Sub Option3_Click()
Inv_Type = 3
LblAdditionData.Visible = True
TxtAdditionData.Visible = True
End Sub
