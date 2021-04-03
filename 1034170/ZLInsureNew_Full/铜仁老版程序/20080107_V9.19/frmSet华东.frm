VERSION 5.00
Begin VB.Form frmSet华东 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保设置"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2485
      TabIndex        =   5
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   1095
      TabIndex        =   4
      Top             =   1755
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -97
      TabIndex        =   3
      Top             =   1500
      Width           =   4875
   End
   Begin VB.CommandButton cmdBrower 
      Caption         =   "浏览(&B)"
      Height          =   400
      Left            =   3400
      TabIndex        =   2
      Top             =   885
      Width           =   1100
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请指定文件存放位置"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "frmSet华东"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng险类 As Long, mblnReturn As Boolean

Public Function ShowMe(ByVal lng险类 As Long) As Boolean
    mlng险类 = lng险类
    Me.Show 1
    ShowMe = mblnReturn
End Function

Private Sub cmdBrower_Click()
    txtPath.Text = BrowPath(Me.hwnd, "请选择文件存放位置：")
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(txtPath.Text) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call ExecuteProcedure(Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'文件存放位置','" & txtPath.Text & "',1)"
    Call ExecuteProcedure(Me.Caption)
    
    mstrSavePath = txtPath.Text
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Me.Hide
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        If rsTemp!参数名 = "文件存放位置" Then txtPath.Text = rsTemp!参数值
        rsTemp.MoveNext
    Loop
End Sub
