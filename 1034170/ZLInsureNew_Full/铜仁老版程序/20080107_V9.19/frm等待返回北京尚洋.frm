VERSION 5.00
Begin VB.Form frm等待返回北京尚洋 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "等待返回......"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   2097
      TabIndex        =   0
      Top             =   630
      Width           =   1965
   End
   Begin VB.Timer TimeRead 
      Interval        =   1000
      Left            =   270
      Top             =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2940
      TabIndex        =   3
      Top             =   1320
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   23
      Picture         =   "frm等待返回北京尚洋.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1065
      Width           =   5325
   End
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   15
      Top             =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费号码"
      Height          =   180
      Left            =   1309
      TabIndex        =   5
      Top             =   690
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "等待医保返回结算数据......"
      Height          =   180
      Left            =   1309
      TabIndex        =   2
      Top             =   255
      Width           =   2340
   End
End
Attribute VB_Name = "frm等待返回北京尚洋"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFile As String

Public Function waitReturn(strFile As String) As String
    mstrFile = strFile
    Me.Show vbModal
    waitReturn = mstrFile
End Function

Private Sub cmdCancel_Click()
    mstrFile = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    mstrFile = Text1.Text
    Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & mstrFile & "'")
    If rsTemp.EOF Then
        MsgBox "中间数据库中没有找到指定收费号码的数据", vbInformation, gstrSysName
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
    Text1.Text = mstrFile
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

Private Sub TimeRead_Timer()
    Dim rsTemp As New ADODB.Recordset
    If cmdOK.Enabled = False Then
        mstrFile = Text1.Text
        Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where VISIT_NUMBER='" & mstrFile & "'")
        If rsTemp.EOF Then
            Exit Sub
        End If
        TimeRead.Enabled = False
        cmdOK.Enabled = True
        Text1.Text = rsTemp!CHARGE_NUMBER
        mstrFile = Text1.Text
    End If
End Sub
