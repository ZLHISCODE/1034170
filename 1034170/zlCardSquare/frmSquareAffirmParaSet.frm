VERSION 5.00
Begin VB.Form frmSquareAffirmParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   5250
      TabIndex        =   17
      Top             =   -135
      Width           =   30
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5475
      TabIndex        =   14
      Top             =   165
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5475
      TabIndex        =   15
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置(&S)"
      Height          =   345
      Left            =   3735
      TabIndex        =   16
      Top             =   3465
      Width           =   1425
   End
   Begin VB.Frame fraRecored 
      Caption         =   "消费记帐审核后票据设置"
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   90
      TabIndex        =   7
      Top             =   1935
      Width           =   4980
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   2250
         TabIndex        =   10
         Top             =   450
         Width           =   1020
      End
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3375
         TabIndex        =   11
         Top             =   450
         Width           =   1380
      End
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   1320
         TabIndex        =   9
         Top             =   450
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.ComboBox cboRecordFormat 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   3450
      End
      Begin VB.Label lblRecordMode 
         AutoSize        =   -1  'True
         Caption         =   "票据打印方式"
         Height          =   180
         Left            =   135
         TabIndex        =   8
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label lblRecordPrint 
         AutoSize        =   -1  'True
         Caption         =   "票据打印格式"
         Height          =   180
         Left            =   150
         TabIndex        =   12
         Top             =   900
         Width           =   1080
      End
   End
   Begin VB.Frame fraCharge 
      Caption         =   "消费确定后票据设置"
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   4995
      Begin VB.ComboBox cboPrintFormat 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   3450
      End
      Begin VB.OptionButton optChargePrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   450
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optChargePrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3375
         TabIndex        =   4
         Top             =   450
         Width           =   1380
      End
      Begin VB.OptionButton optChargePrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   2250
         TabIndex        =   3
         Top             =   450
         Width           =   1020
      End
      Begin VB.Label lblPrintFormat 
         AutoSize        =   -1  'True
         Caption         =   "票据打印格式"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lblChargePrintMode 
         AutoSize        =   -1  'True
         Caption         =   "票据打印方式"
         Height          =   180
         Left            =   135
         TabIndex        =   1
         Top             =   450
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSquareAffirmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String, mblnOk As Boolean
Private Const mlngModul = 1151
Public Function SetPara(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费的相关参数设置入口
    '返回:点击确定,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-11 00:16:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOk = False
    Me.Show 1, frmMain
    SetPara = mblnOk
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function SavePara() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数
    '编制:刘兴洪
    '日期:2011-08-10 23:37:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "收费收据格式", Val(Split(cboPrintFormat.Text, "-")(0)), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "收费打印方式", IIf(optChargePrint(0).Value, 0, IIf(optChargePrint(1).Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "审核收据格式", Val(Split(cboRecordFormat.Text, "-")(0)), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "审核打印方式", IIf(optRecordPrint(0).Value, 0, IIf(optRecordPrint(1).Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    'zlDatabase.setPara "药品单位", IIf(opt单位(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    
    SavePara = True
End Function
 Private Sub cmdOK_Click()
    If SavePara = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub
Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", Me)
End Sub
Private Sub InitBillFormat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化票据格式
    '编制:刘兴洪
    '日期:2011-08-10 23:57:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    '票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A, zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号='ZL" & glngSys \ 100 & "_BILL_1151'  " & _
    "   Order by  序号"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboPrintFormat
        .Clear: cboRecordFormat.Clear
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!序号) & "-" & Nvl(rsTemp!说明)
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!序号))
            cboRecordFormat.AddItem Nvl(rsTemp!序号) & "-" & Nvl(rsTemp!说明)
            cboRecordFormat.ItemData(cboRecordFormat.NewIndex) = Val(Nvl(rsTemp!序号))
            rsTemp.MoveNext
        Loop
        If .ListCount <> 0 Then .ListIndex = 0
        If cboRecordFormat.ListCount <> 0 Then cboRecordFormat.ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数
    '编制:刘兴洪
    '日期:2011-08-10 23:48:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim blnHavePrivs As Boolean, strValue As String
    Dim j As Long
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    i = Val(zlDatabase.getPara("收费收据格式", glngSys, mlngModul, , Array(cboPrintFormat), blnHavePrivs))
    With cboPrintFormat
        For j = 0 To .ListCount - 1
            If .ItemData(j) = i Then .ListIndex = j: Exit For
        Next
    End With
    i = Val(zlDatabase.getPara("收费打印方式", glngSys, mlngModul, , Array(optChargePrint(0), optChargePrint(1), optChargePrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optChargePrint(i).Value = True
    
    i = Val(zlDatabase.getPara("审核收据格式", glngSys, mlngModul, , Array(cboRecordFormat), blnHavePrivs))
    With cboRecordFormat
        For j = 0 To .ListCount - 1
            If .ItemData(j) = i Then .ListIndex = j: Exit For
        Next
    End With
    
    i = Val(zlDatabase.getPara("审核打印方式", glngSys, mlngModul, , Array(optRecordPrint(0), optRecordPrint(1), optRecordPrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optRecordPrint(i).Value = True
    'i=val(zlDatabase.setPara ("药品单位", glngSys, mlngModul, ,array(opt单位(0),opt单位(1)),blnHavePrivs))
    'opt单位(iif(i=0,0,1)).value=true
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mstrPrivs = ";" & GetPrivFunc(glngSys, mlngModul) & ";"
    Call InitBillFormat
    Call InitPara
End Sub

