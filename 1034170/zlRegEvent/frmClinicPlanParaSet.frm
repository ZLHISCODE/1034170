VERSION 5.00
Begin VB.Form frmClinicPlanParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6840
   Icon            =   "frmClinicPlanParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo号码比较方式 
      Height          =   300
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox cbo号源安排站点 
      Height          =   300
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox chk替诊医生级别检查 
      Caption         =   "替诊医生职务级别检查"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   2145
   End
   Begin VB.Frame fraSplit 
      Height          =   4485
      Left            =   5280
      TabIndex        =   22
      Top             =   -150
      Width           =   25
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "周出诊表打印设置(&4)"
      Height          =   405
      Index           =   3
      Left            =   2910
      TabIndex        =   21
      Top             =   3870
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "月出诊表打印设置(&3)"
      Height          =   405
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   3870
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "固定出诊表打印设置(&2)"
      Height          =   405
      Index           =   1
      Left            =   2910
      TabIndex        =   19
      Top             =   3420
      Width           =   2145
   End
   Begin VB.CheckBox chkReplaceDoctor 
      Caption         =   "按替诊医生同步更新预约挂号单"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   90
      Width           =   2835
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "预约清单打印设置(&1)"
      Height          =   405
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   3420
      Width           =   2145
   End
   Begin VB.Frame fraVisitTablePrintMode 
      Caption         =   "出诊表打印方式"
      Height          =   735
      Left            =   180
      TabIndex        =   14
      Top             =   2550
      Width           =   4875
      Begin VB.OptionButton optVisitTablePrintMode 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optVisitTablePrintMode 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1035
      End
      Begin VB.OptionButton optVisitTablePrintMode 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3090
         TabIndex        =   17
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Frame fraPrintMode 
      Caption         =   "预约清单打印方式"
      Height          =   1305
      Left            =   3060
      TabIndex        =   10
      Top             =   1110
      Width           =   1995
      Begin VB.OptionButton optPrintMode 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   13
         Top             =   930
         Width           =   1395
      End
      Begin VB.OptionButton optPrintMode 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   615
         Width           =   1035
      End
      Begin VB.OptionButton optPrintMode 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraToExcelMode 
      Caption         =   "预约清单控制方式"
      Height          =   1305
      Left            =   180
      TabIndex        =   6
      Top             =   1110
      Width           =   2715
      Begin VB.OptionButton optToExcelMode 
         Caption         =   "选择是否输出到Excel"
         Height          =   225
         Index           =   2
         Left            =   300
         TabIndex        =   9
         Top             =   930
         Width           =   2025
      End
      Begin VB.OptionButton optToExcelMode 
         Caption         =   "自动输出到Excel"
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   8
         Top             =   615
         Width           =   1665
      End
      Begin VB.OptionButton optToExcelMode 
         Caption         =   "不输出到Excel"
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   330
      Left            =   5550
      TabIndex        =   23
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   5550
      TabIndex        =   24
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   330
      Left            =   5580
      TabIndex        =   25
      Top             =   3390
      Width           =   1100
   End
   Begin VB.Label lbl号码比较方式 
      AutoSize        =   -1  'True
      Caption         =   "排序时号源号码的比较方式"
      Height          =   180
      Left            =   210
      TabIndex        =   4
      Top             =   780
      Width           =   2160
   End
   Begin VB.Label lbl号源安排站点 
      AutoSize        =   -1  'True
      Caption         =   "将未区分站点的号源分配给                   进行出诊安排"
      Height          =   180
      Left            =   210
      TabIndex        =   2
      Top             =   420
      Width           =   4950
   End
End
Attribute VB_Name = "frmClinicPlanParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String
Private mlngModul As Long
Private mblnOk As Boolean

Public Function ShowMe(frmParent As Form, ByVal lngModul As Long, _
    ByVal strPrivs As String) As Boolean
    '程序入口
    mstrPrivs = strPrivs: mlngModul = lngModul
    
    On Error Resume Next
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim blnHavePrivs As Boolean
    Dim strValue As String
    
    On Error GoTo ErrHandler
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "替诊医生级别检查", IIf(chk替诊医生级别检查.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "按替诊医生同步更新预约挂号单", IIf(chkReplaceDoctor.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "未区分站点的号源的维护站点", zlStr.NeedCode(cbo号源安排站点.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "号码排序比较方式", cbo号码比较方式.ItemData(cbo号码比较方式.ListIndex), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "预约清单控制方式", GetSelectedIndex(optToExcelMode), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "预约清单打印方式", GetSelectedIndex(optPrintMode), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "出诊表打印方式", GetSelectedIndex(optVisitTablePrintMode), glngSys, mlngModul, blnHavePrivs
    mblnOk = True
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub cmdPrintSet_Click(index As Integer)
    On Error GoTo ErrHandler
    Select Case index
    Case 0: '预约清单打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me)
    Case 1: '固定出诊表打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_1", Me)
    Case 2: '月出诊表打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_2", Me)
    Case 3: '周出诊表打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_3", Me)
    Case Else:
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub Form_Load()
    Dim index As Integer, blnHavePrivs As Boolean, strValue As String
    Dim strSQL As String, rsRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    blnHavePrivs = IsHavePrivs(mstrPrivs, "参数设置")
    
    chk替诊医生级别检查.Value = Val(zlDatabase.GetPara("替诊医生级别检查", glngSys, mlngModul, 0, Array(chk替诊医生级别检查), blnHavePrivs))
    chkReplaceDoctor.Value = Val(zlDatabase.GetPara("按替诊医生同步更新预约挂号单", glngSys, mlngModul, 0, Array(chkReplaceDoctor), blnHavePrivs))
    
    strValue = zlDatabase.GetPara("未区分站点的号源的维护站点", glngSys, mlngModul, "", Array(lbl号源安排站点, cbo号源安排站点), blnHavePrivs)
    strSQL = _
        "Select Distinct b.编号, b.名称" & vbNewLine & _
        "From 部门表 A, Zlnodelist B" & vbNewLine & _
        "Where a.站点 = b.编号" & vbNewLine & _
        "Order By b.编号"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "站点查询")
    With cbo号源安排站点
        .Clear
        .AddItem ""
        Do While Not rsRecord.EOF
            .AddItem rsRecord!编号 & "-" & rsRecord!名称
            If strValue = rsRecord!编号 Then .ListIndex = .NewIndex
            rsRecord.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    With cbo号码比较方式
        .Clear
        .AddItem "0-按字符比较":  .ItemData(.NewIndex) = 0
        .AddItem "1-按数值比较": .ItemData(.NewIndex) = 1
    End With
    index = Val(zlDatabase.GetPara("号码排序比较方式", glngSys, mlngModul, 0, Array(lbl号码比较方式, cbo号码比较方式), blnHavePrivs))
    zlControl.CboLocate cbo号码比较方式, index, True
    
    index = Val(zlDatabase.GetPara("预约清单控制方式", glngSys, mlngModul, 0, Array(optToExcelMode(0), optToExcelMode(1), optToExcelMode(2)), blnHavePrivs))
    If index <= optToExcelMode.UBound Then optToExcelMode(index).Value = True
    
    index = Val(zlDatabase.GetPara("预约清单打印方式", glngSys, mlngModul, 0, Array(optPrintMode(0), optPrintMode(1), optPrintMode(2)), blnHavePrivs))
    If index <= optPrintMode.UBound Then optPrintMode(index).Value = True
    
    index = Val(zlDatabase.GetPara("出诊表打印方式", glngSys, mlngModul, 0, Array(optVisitTablePrintMode(0), optVisitTablePrintMode(1), optVisitTablePrintMode(2)), blnHavePrivs))
    If index <= optVisitTablePrintMode.UBound Then optVisitTablePrintMode(index).Value = True
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

