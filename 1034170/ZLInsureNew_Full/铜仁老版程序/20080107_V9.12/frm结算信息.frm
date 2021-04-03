VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm结算信息 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结算信息"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ControlBox      =   0   'False
   Icon            =   "frm结算信息.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshBill 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4154
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorSel    =   4194304
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm结算信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng结帐ID As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

'Modified by 朱玉宝 20031218 地区：福州 新增窗体
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select Decode(A.记录性质,1,'冲预交',11,'冲预交',A.结算方式) 结算方式,Nvl(A.冲预交,0) 金额 " & _
                " From 病人预交记录 A,保险帐户 B " & _
                " Where A.病人ID=B.病人ID And B.险类=" & gintInsure & " And A.结帐ID=" & mlng结帐ID
    Call OpenRecordset(rsTemp, "获取本次交易结算信息")
    
    With MshBill
        .Clear
        .Rows = 2
        .Cols = 2
        .TextMatrix(0, 0) = "结算方式"
        .TextMatrix(0, 1) = "金额"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1200
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
    With rsTemp
        Do While Not .EOF
            MshBill.TextMatrix(.AbsolutePosition, 0) = !结算方式
            MshBill.TextMatrix(.AbsolutePosition, 1) = Format(!金额, "#####0.00;-#####0.00; ;")
            If MshBill.Rows - 1 = .AbsolutePosition Then MshBill.Rows = MshBill.Rows + 1
            .MoveNext
        Loop
        If Trim(MshBill.TextMatrix(MshBill.Rows - 1, 0)) = "" Then MshBill.Rows = MshBill.Rows - 1
    End With
End Sub

Public Sub ShowME(Optional ByVal lng结帐ID As Long = 0)
    mlng结帐ID = lng结帐ID
    Me.Show 1
End Sub
