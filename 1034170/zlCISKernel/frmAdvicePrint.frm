VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdvicePrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人医嘱单打印"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "frmAdvicePrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Tag             =   "打印选项"
      Top             =   960
      Visible         =   0   'False
      Width           =   6600
      Begin VB.Frame fraDrug 
         Caption         =   "药品医嘱打印方式"
         Height          =   735
         Left            =   285
         TabIndex        =   25
         Top             =   105
         Width           =   6225
         Begin VB.CheckBox chkDrugUseWayHaveRow 
            Caption         =   "西药、成药用法单独打印一行"
            Height          =   285
            Left            =   210
            TabIndex        =   26
            Top             =   300
            Width           =   3180
         End
      End
      Begin VB.Frame fraLongAdvice 
         Caption         =   "医嘱单样式"
         Height          =   1665
         Left            =   285
         TabIndex        =   22
         Top             =   1140
         Width           =   6225
         Begin VB.CheckBox chkCZCZHY 
            Caption         =   "重整"
            Height          =   195
            Left            =   3600
            TabIndex        =   43
            Top             =   360
            Width           =   690
         End
         Begin VB.CheckBox chkCQSHHY 
            Caption         =   "术后"
            Height          =   195
            Left            =   2760
            TabIndex        =   42
            Top             =   360
            Width           =   690
         End
         Begin VB.CheckBox chkLSZKHY 
            Caption         =   "转科"
            Height          =   195
            Left            =   2040
            TabIndex        =   39
            Top             =   720
            Width           =   690
         End
         Begin VB.CheckBox chkPrintTurnPage 
            Caption         =   "转科"
            Height          =   195
            Left            =   2040
            TabIndex        =   24
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkTurn 
            Caption         =   "转科换页后在长嘱单首行打印""重开医嘱"""
            Height          =   405
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   3720
         End
         Begin VB.Label lblLSYZ 
            AutoSize        =   -1  'True
            Caption         =   "临嘱单另起一页打印"
            Height          =   180
            Left            =   240
            TabIndex        =   41
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label lblCQYZ 
            AutoSize        =   -1  'True
            Caption         =   "长嘱单另起一页打印"
            Height          =   180
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   1620
         End
      End
      Begin VB.Frame fraPrintPos 
         Caption         =   "转科、出院、死亡医嘱打印位置"
         Height          =   750
         Left            =   255
         TabIndex        =   12
         Top             =   3105
         Width           =   6225
         Begin VB.OptionButton optPrintPos 
            Caption         =   "长期医嘱单"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   330
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton optPrintPos 
            Caption         =   "临时医嘱单"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   1815
            TabIndex        =   14
            Top             =   330
            Width           =   1395
         End
         Begin VB.OptionButton optPrintPos 
            Caption         =   "以上两者"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   13
            Top             =   345
            Width           =   1275
         End
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&F)"
      Height          =   315
      Left            =   5460
      TabIndex        =   38
      Top             =   3
      Width           =   800
   End
   Begin VB.CommandButton cmdClsLastPrint 
      Caption         =   "清除上次打印(&C)"
      Height          =   350
      Left            =   1890
      TabIndex        =   33
      Top             =   5340
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Frame fraPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4065
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Tag             =   "常规打印"
      Top             =   900
      Visible         =   0   'False
      Width           =   6600
      Begin VB.Frame fraClear 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3810
         TabIndex        =   34
         Top             =   3615
         Width           =   2385
         Begin VB.CommandButton cmdClear 
            Caption         =   "清除(&D)"
            Height          =   350
            Left            =   1485
            TabIndex        =   36
            Top             =   0
            Width           =   800
         End
         Begin VB.TextBox txtClearPage 
            Height          =   270
            Left            =   945
            MaxLength       =   3
            TabIndex        =   35
            Top             =   45
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "清除起始页"
            Height          =   180
            Left            =   0
            TabIndex        =   37
            Top             =   80
            Width           =   900
         End
      End
      Begin VB.CheckBox chkSeqPage 
         Caption         =   "重打“待续打”页"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   4020
         TabIndex        =   32
         Top             =   585
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   2
         Left            =   2655
         MouseIcon       =   "frmAdvicePrint.frx":058A
         Picture         =   "frmAdvicePrint.frx":0B14
         Top             =   450
         Width           =   360
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   1
         Left            =   1395
         MouseIcon       =   "frmAdvicePrint.frx":11FE
         Picture         =   "frmAdvicePrint.frx":1788
         Top             =   450
         Width           =   360
      End
      Begin VB.Image imgIcon 
         DragIcon        =   "frmAdvicePrint.frx":1E72
         Height          =   360
         Index           =   0
         Left            =   195
         MouseIcon       =   "frmAdvicePrint.frx":255C
         Picture         =   "frmAdvicePrint.frx":2AE6
         Top             =   450
         Width           =   360
      End
      Begin VB.Label lblPrint 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdvicePrint.frx":31D0
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   210
         TabIndex        =   17
         Top             =   90
         Width           =   3600
      End
      Begin VB.Label lblStopPrint 
         AutoSize        =   -1  'True
         Caption         =   "提醒：该病人有确认停止的医嘱需要打印。"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   3675
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.Label lblPrintIcoInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "已打印       待续打        未打印"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   27
         Top             =   585
         Width           =   2970
      End
   End
   Begin VB.Frame fraPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4080
      Index           =   1
      Left            =   765
      TabIndex        =   5
      Tag             =   "停嘱打印"
      Top             =   735
      Visible         =   0   'False
      Width           =   6600
      Begin VB.Image imgIcon 
         DragIcon        =   "frmAdvicePrint.frx":321A
         Height          =   360
         Index           =   3
         Left            =   210
         MouseIcon       =   "frmAdvicePrint.frx":3904
         Picture         =   "frmAdvicePrint.frx":3E8E
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lblPrintIcoInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "套打"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lblInSidePrint 
         BackStyle       =   0  'Transparent
         Caption         =   "停嘱打印指在已打印的医嘱单上对停止时间进行套打。请单击下面的图片选择要套打的医嘱单页号。"
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   210
         TabIndex        =   10
         Top             =   135
         Width           =   4320
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   4950
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   30
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar vsc 
      Height          =   1815
      Left            =   6585
      SmallChange     =   50
      TabIndex        =   29
      Top             =   375
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox picCHKH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   5535
      ScaleHeight     =   810
      ScaleWidth      =   1320
      TabIndex        =   18
      Top             =   -135
      Visible         =   0   'False
      Width           =   1320
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   0
         Left            =   1245
         ScaleHeight     =   900
         ScaleWidth      =   705
         TabIndex        =   20
         Top             =   525
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.PictureBox picPaperB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   0
         Left            =   1365
         ScaleHeight     =   900
         ScaleWidth      =   705
         TabIndex        =   19
         Top             =   585
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   495
         Picture         =   "frmAdvicePrint.frx":4578
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   0
         Left            =   150
         Picture         =   "frmAdvicePrint.frx":4F7A
         Top             =   465
         Width           =   240
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   21
         Top             =   855
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.OptionButton optReport 
      Caption         =   "长期医嘱单"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton optReport 
      Caption         =   "临时医嘱单"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   1
      Left            =   1455
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox cboBaby 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3510
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   5945
      TabIndex        =   8
      Top             =   5340
      Width           =   800
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   975
      TabIndex        =   7
      Top             =   5340
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   100
      TabIndex        =   6
      Top             =   5340
      Width           =   800
   End
   Begin MSComctlLib.TabStrip tbsMain 
      Height          =   4605
      Left            =   100
      TabIndex        =   3
      Top             =   400
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   8123
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "常规打印"
            Key             =   "常规打印"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "停嘱打印"
            Key             =   "停嘱打印"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "打印选项"
            Key             =   "打印选项"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "长期医嘱单：共13页。"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3630
      TabIndex        =   31
      Top             =   5430
      Width           =   1800
   End
   Begin VB.Label lblBaby 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医嘱"
      Height          =   180
      Left            =   3090
      TabIndex        =   9
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmAdvicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'接口参数：
Private mfrmParent As Object
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstrDefKey As String '缺省定位到的打印功能

'模块变量
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mrsPrint As ADODB.Recordset
Private mintPrintCount As Integer
Attribute mintPrintCount.VB_VarHelpID = -1
Private mblnTrans As Boolean '用于事件之件的事务嵌套控制
Private mlngPrintType As Long   '医嘱打印模式   1-新开后打印，0-校对后打印

Private mlngRows临嘱 As Long    '在1页纸上打印的总行数
Private mlngRows长嘱 As Long    '
Private mintMid As Integer      '重打页和应打页的临界页，mintMid  （未打页/续打页）
Private mbln续打页 As Boolean   '是否有待续打页
Private mstrTurnPages As String '所有的换页打印的页号，格式 "2,3,6,8,9"
Private mdat重整时间 As Date
Private mlngPrintedMaxPage As Long ' 已经打印过的医嘱的最大页号

Private mintPageCount As Integer        '总页数  常规打印的页数，只要进入窗体这个数字是固定的。
Private mintStopPageCount As Integer    '总页数  停嘱打印的页数，只变小，执行套打后会
Private mstrPrePars As String '上一次的参数值状态组合

Private Enum mCtlID
    opt医嘱_长嘱 = 0
    opt医嘱_临嘱 = 1
    
    fra界面_连打 = 0
    fra界面_套打 = 1
    fra界面_设置 = 2
    
    opt位置_长嘱 = 0
    opt位置_临嘱 = 1
    opt位置_两者 = 2
    
    lbl图标说明_连打 = 0
    lbl图标说明_套打 = 1
    
    img已打 = 0
    img续打 = 1
    img未打 = 2
    img套打 = 3
    
    pic连打_容器 = 1
    pic连打_纸面 = 2
    
    pic套打_容器 = 3
    pic套打_纸面 = 4
    
End Enum

Public Sub ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strDefKey As String)
    Set mfrmParent = frmParent
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrDefKey = strDefKey
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub

Private Sub cmdRefresh_Click()
'刷新
    Call RefreshPrintPage
    Call tbsMain_Click
End Sub

Private Sub Form_Load()
    Dim arrBaby As Variant, strBaby As String
    Dim blnPriv As Boolean, i As Long
 
    mblnTrans = False
    '设置报表权限
    '长期医嘱单
    blnPriv = False
    If InStr(UserInfo.性质, "医生") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱下达), "长期医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv And InStr(UserInfo.性质, "护士") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱发送), "长期医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv Then
        optReport(opt医嘱_临嘱).value = True
        optReport(opt医嘱_长嘱).Enabled = False
    End If
    
    '临时医嘱单
    blnPriv = False
    If InStr(UserInfo.性质, "医生") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱下达), "临时医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv And InStr(UserInfo.性质, "护士") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱发送), "临时医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv Then
        optReport(opt医嘱_长嘱).value = True
        optReport(opt医嘱_临嘱).Enabled = False
    End If
    
    '例外情况：两个报表应至少有一个有权限
    If Not optReport(opt医嘱_长嘱).Enabled And Not optReport(opt医嘱_临嘱).Enabled Then
        Unload Me: Exit Sub
    End If
    
    mdat重整时间 = GetRsRedoDate(mlng病人ID, mlng主页ID)
    
    '初始化婴儿选择
    cboBaby.AddItem "病人医嘱"
    Call zlControl.CboSetIndex(cboBaby.hWnd, 0)
    
    strBaby = GetBabyRegList(mlng病人ID, mlng主页ID)
    If strBaby <> "" Then
        arrBaby = Split(strBaby, "<Split>")
        For i = 0 To UBound(arrBaby)
            cboBaby.AddItem "婴儿 " & i + 1 & IIF(arrBaby(i) <> "", "：" & arrBaby(i), "")
        Next
    Else
        lblBaby.Visible = False
        cboBaby.Visible = False
    End If
    Call zlControl.CboSetWidth(cboBaby.hWnd, cboBaby.Width * 1.55)
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
        
    '打印选项
    chkPrintTurnPage.value = Val(zlDatabase.GetPara("长嘱单转科换页", glngSys, p住院医嘱发送, 0))
    chkCQSHHY.value = Val(zlDatabase.GetPara("长嘱单术后换页", glngSys, p住院医嘱发送, 0))
    chkCZCZHY.value = Val(zlDatabase.GetPara("长嘱单重整换页", glngSys, p住院医嘱发送, 0))
    
    '用法单独打印一行
    chkDrugUseWayHaveRow.value = Val(zlDatabase.GetPara("药品用法单独打印一行", glngSys, p住院医嘱发送, 0))
    
    '转科换页后在首行打印重开医嘱
    chkTurn.value = Val(zlDatabase.GetPara("转科换页后在首行打印重开医嘱", glngSys, p住院医嘱发送, 0))
    
    '医嘱打印模式
    mlngPrintType = Val(zlDatabase.GetPara("医嘱单打印模式", glngSys, p住院医嘱下达))
    
    chkLSZKHY.value = Val(zlDatabase.GetPara("临嘱单转科换页", glngSys, p住院医嘱发送, 0))
    
    If mlngPrintType = 1 Then
        lblStopPrint = "提醒：该病人有停止/确认停止的医嘱需要打印。"
    Else
        lblStopPrint = "提醒：该病人有确认停止的医嘱需要打印。"
    End If
    
    i = Val(zlDatabase.GetPara("转科和出院打印", glngSys, p住院医嘱发送, 1))
    optPrintPos(i).value = True
    
    mstrPrePars = chkPrintTurnPage.value & chkDrugUseWayHaveRow.value & chkTurn.value & mlngPrintType & chkLSZKHY.value & i
    
    If InStr(GetInsidePrivs(p住院医嘱发送), "医嘱选项设置") = 0 Then
        fraPrint(fra界面_设置).Enabled = False
        chkPrintTurnPage.Enabled = False
        optPrintPos(opt位置_长嘱).Enabled = False
        optPrintPos(opt位置_临嘱).Enabled = False
        optPrintPos(opt位置_两者).Enabled = False
    End If
    
    mlngRows临嘱 = GetReportRows(glngSys, "ZL1_INSIDE_1254_2")
    mlngRows长嘱 = GetReportRows(glngSys, "ZL1_INSIDE_1254_1")
    
    Call Insert打印记录
    
    Call LoadAllPaper
    
    '刷新界面数据
    If mstrDefKey <> "" And tbsMain.SelectedItem.Key <> mstrDefKey Then
        tbsMain.Tag = "NoneClick"
        For i = 1 To tbsMain.Tabs.Count
            If tbsMain.Tabs(i).Key = mstrDefKey Then
                tbsMain.Tabs(i).Selected = True
                Exit For
            End If
        Next
        tbsMain.Tag = ""
    End If
     
    Call tbsMain_Click
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    On Error Resume Next
    
    Me.Height = 6600
    
    With tbsMain
        .Left = 100
        .Top = 400
        .Width = Me.ScaleWidth - 200
        .Height = 5200
    End With
    
    cmdPreview.Top = Me.ScaleHeight - 470
    cmdPrint.Top = cmdPreview.Top
    cmdCancel.Top = cmdPreview.Top
    cmdCancel.Left = tbsMain.Width + tbsMain.Left - cmdCancel.Width
    cmdClsLastPrint.Top = cmdPreview.Top
    lblTotal.Top = cmdPreview.Top + 100
    fraClear.Visible = cmdClsLastPrint.Visible
    For i = 0 To 2
        fraPrint(i).Top = 750
        fraPrint(i).Left = 150
        fraPrint(i).Width = tbsMain.Width - 400
        fraPrint(i).Height = tbsMain.Height - 430
    Next
    
    fraPrint(fra界面_设置).Width = tbsMain.Width - 100
    
    For i = 0 To 3
        imgIcon(i).Top = 530
    Next
    
    imgIcon(img套打).Left = imgIcon(img已打).Left
    
    lblPrintIcoInfo(lbl图标说明_连打).Top = 670
    lblPrintIcoInfo(lbl图标说明_套打).Top = 670
    lblStopPrint.Top = 4500
    lblStopPrint.Left = 100
    fraClear.Top = lblStopPrint.Top - 80
    fraClear.Left = tbsMain.Width - fraClear.Width - 350
 
    lblPrint.Left = lblInSidePrint.Left
    lblInSidePrint.Top = lblPrint.Top
    
    fraDrug.Top = 250
    fraLongAdvice.Top = 1400
    fraPrintPos.Top = 3400
    fraDrug.Left = 200
    fraLongAdvice.Left = fraDrug.Left
    fraPrintPos.Left = fraDrug.Left
    cmdRefresh.Top = cboBaby.Top
    cmdRefresh.Left = cmdCancel.Left + cmdCancel.Width - cmdRefresh.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False '以防万一
    Set mrsPrint = Nothing
    mbln续打页 = False
    Call UnLoadPaper
    mintPageCount = 0
    mintStopPageCount = 0
End Sub

Private Sub chkTurn_Click()
    Call zlDatabase.SetPara("转科换页后在首行打印重开医嘱", chkTurn.value, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
End Sub

Private Sub chkLSZKHY_Click()
    Call zlDatabase.SetPara("临嘱单转科换页", chkLSZKHY.value, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
End Sub

Private Sub chkDrugUseWayHaveRow_Click()
    Call zlDatabase.SetPara("药品用法单独打印一行", chkDrugUseWayHaveRow.value, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
End Sub

Private Sub optPrintPos_Click(Index As Integer)
    Call zlDatabase.SetPara("转科和出院打印", Index, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
End Sub

Private Sub chkPrintTurnPage_Click()
    Call zlDatabase.SetPara("长嘱单转科换页", chkPrintTurnPage.value, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
    
    
    If chkPrintTurnPage.value = 1 Then
        chkTurn.Enabled = True
    Else
        chkTurn.value = 0
        chkTurn.Enabled = False
    End If
End Sub

Private Sub cboBaby_Click()
    Call RefreshFace
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Call AdvicePrint(1)
    Call RefreshFace
End Sub

Private Sub cmdPrint_Click()
    Call AdvicePrint(2)
    Call RefreshFace
End Sub

Private Sub optReport_Click(Index As Integer)
    Dim i As Integer
    
    If Not Visible Then Exit Sub '设置权限时不处理
    
    tbsMain.Tag = "NoneClick"
    If optReport(opt医嘱_长嘱).value Then
        tbsMain.Tabs.Add 2, "停嘱打印", "停嘱打印"
    Else
        For i = 1 To tbsMain.Tabs.Count
            If tbsMain.Tabs(i).Key = "停嘱打印" Then
                tbsMain.Tabs.Remove tbsMain.Tabs(i).Key
                Exit For
            End If
        Next
    End If
    tbsMain.Tag = ""
    
    Call tbsMain_Click
End Sub

Private Sub tbsMain_Click()
    Dim i As Long
    
    If tbsMain.Tag = "NoneClick" Then Exit Sub
    
    For i = 0 To fraPrint.UBound
        fraPrint(i).Visible = fraPrint(i).Tag = tbsMain.SelectedItem.Key
        If fraPrint(i).Tag = tbsMain.SelectedItem.Key Then
            fraPrint(i).ZOrder
            If i = fra界面_连打 Then picContainer(pic连打_容器).ZOrder
            If i = fra界面_套打 Then picContainer(pic套打_容器).ZOrder
        End If
    Next
    Call RefreshFace
    picContainer(pic套打_纸面).Top = 0
    picContainer(pic连打_纸面).Top = 0
    vsc.value = 0
    cmdRefresh.Enabled = tbsMain.SelectedItem.Key = "常规打印"
End Sub
 
Private Sub RefreshFace()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim i As Long, j As Long
    Dim intIco As Integer '0 - 已打印 1 - 待续打 2 － 待打印
    Dim intIndex As Integer
    Dim lngTmp As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    If tbsMain.SelectedItem.Key = "常规打印" Then
        '判断参数是否发生变化如果参数发生变要要重新生成打印数据
        If optPrintPos(opt位置_长嘱).value Then
            i = opt位置_长嘱
        ElseIf optPrintPos(opt位置_临嘱).value Then
            i = opt位置_临嘱
        ElseIf optPrintPos(opt位置_两者).value Then
            i = opt位置_两者
        End If
        strTmp = chkPrintTurnPage.value & chkDrugUseWayHaveRow.value & chkTurn.value & mlngPrintType & chkLSZKHY.value & i
        
        If mstrPrePars <> strTmp Then
            mstrPrePars = strTmp
            Call RefreshPrintPage(True)
        End If
        
        strSql = "select m.页号,sum(m.打印) as 打印,sum(m.未打印) as 未打印,count(1) as 行数" & vbNewLine & _
            "from (select a.页号,decode(a.打印时间,null,0,1) as 打印,decode(a.打印时间,null,1,0) as 未打印" & vbNewLine & _
            "from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2] and nvl(a.婴儿,0)=[3] and a.期效=[4] and 行号>0) m" & vbNewLine & _
            "group by m.页号 order by m.页号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(opt医嘱_长嘱).value, 0, 1))
        mintMid = 0
        mlngPrintedMaxPage = 0
        mstrTurnPages = ""
        mbln续打页 = False
        mintPageCount = rsTmp.RecordCount
        chkSeqPage.Visible = False
        For i = 1 To rsTmp.RecordCount

            If Val(rsTmp!打印 & "") > 0 And Val(rsTmp!未打印 & "") = 0 Then
                intIco = 0
                mlngPrintedMaxPage = Val(rsTmp!页号 & "")
            ElseIf Val(rsTmp!打印 & "") > 0 And Val(rsTmp!未打印 & "") > 0 Then
                If 0 = mintMid Then
                    mintMid = Val(rsTmp!页号 & "")
                    mbln续打页 = True
                    chkSeqPage.Visible = True
                End If
                intIco = 1
                mlngPrintedMaxPage = Val(rsTmp!页号 & "")
            ElseIf Val(rsTmp!打印 & "") = 0 And Val(rsTmp!未打印 & "") > 0 Then
                If 0 = mintMid Then mintMid = Val(rsTmp!页号 & "")
                intIco = 2
            End If
            
            '最后一页不算做换打页
            If i <> rsTmp.RecordCount Then
                If Val(rsTmp!行数 & "") < IIF(optReport(0).value, mlngRows长嘱, mlngRows临嘱) Then
                    mstrTurnPages = mstrTurnPages & "," & rsTmp!页号
                End If
            End If
      
            Set imgIco(i).Picture = imgIcon(intIco).Picture
            imgIco(i).ToolTipText = IIF(intIco = 0, "已打印", IIF(intIco = 1, "待续打", "未打印"))
            imgChk(i).Visible = IIF(intIco = 0, False, True)
            picPaper(i).Visible = True
            picPaperB(i).Visible = True
            rsTmp.MoveNext
        Next
        
        '判断是否应该显示清除上次打印按钮
        rsTmp.Filter = "打印>0"
        
        cmdClsLastPrint.Visible = Not rsTmp.EOF
        
        For i = mintPageCount + 1 To Val(picPaper(0).Tag)
            imgChk(i).Visible = False
            picPaper(i).Visible = False
            picPaperB(i).Visible = False
        Next
        
        If mstrTurnPages <> "" Then mstrTurnPages = Mid(mstrTurnPages, 2)
        
        Set rsTmp = GetStopedAdvice(True)
        lblStopPrint.Visible = rsTmp.RecordCount > 0
        
        cmdPreview.Enabled = mintPageCount > 0
        cmdPrint.Enabled = mintPageCount > 0
        
        lngTmp = IntEx(mintPageCount / 21)
        If lngTmp = 0 Then lngTmp = 1
 
        picContainer(pic连打_纸面).Height = lngTmp * 3450
        vsc.Visible = lngTmp > 1
        If lngTmp > 1 Then
            vsc.Max = (lngTmp - 1) * 3450 / Screen.TwipsPerPixelY
        End If
        
        If mintPageCount = 0 Then
            lblTotal.Caption = IIF(optReport(opt医嘱_长嘱).value, "长期", "临时") & "医嘱单：无。"
        Else
            lblTotal.Caption = IIF(optReport(opt医嘱_长嘱).value, "长期", "临时") & "医嘱单：共" & mintPageCount & "页。"
        End If
        lblTotal.Visible = True
        If mlngPrintedMaxPage <> 0 Then txtClearPage.Text = mlngPrintedMaxPage
    ElseIf tbsMain.SelectedItem.Key = "停嘱打印" Then
        
        Set rsTmp = GetStopedAdvice(False)
        
        mintStopPageCount = rsTmp.RecordCount
        
        For i = 1 To rsTmp.RecordCount
            lblNum(i + 1000).Caption = Val(rsTmp!页号 & "")
            lblNum(i + 1000).ToolTipText = "第" & Val(rsTmp!页号 & "") & "页"
            picPaper(i + 1000).Visible = True
            picPaperB(i + 1000).Visible = True
            imgChk(i + 1000).Visible = False
            rsTmp.MoveNext
        Next
        
        For i = mintStopPageCount + 1 To Val(picPaperB(0).Tag)
            imgChk(i + 1000).Visible = False
            picPaper(i + 1000).Visible = False
            picPaperB(i + 1000).Visible = False
        Next
        
        cmdPrint.Enabled = mintStopPageCount <> 0
        cmdPreview.Enabled = mintStopPageCount <> 0
        
        lngTmp = IntEx(mintStopPageCount / 21)
        If lngTmp = 0 Then lngTmp = 1
        picContainer(pic套打_纸面).Height = lngTmp * 3450
        
        vsc.Visible = lngTmp > 1
        If lngTmp > 1 Then
            vsc.Max = (lngTmp - 1) * 3450 / Screen.TwipsPerPixelY
        End If
        If mintStopPageCount = 0 Then
            lblTotal.Caption = "套打医嘱单：无。"
        Else
            lblTotal.Caption = "套打医嘱单：共" & mintStopPageCount & "页。"
        End If
        lblTotal.Visible = True
        cmdClsLastPrint.Visible = False
    ElseIf tbsMain.SelectedItem.Key = "打印选项" Then
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        vsc.Visible = False
        lblTotal.Visible = False
        cmdClsLastPrint.Visible = False
    End If
    fraClear.Visible = cmdClsLastPrint.Visible
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetStopedAdvice(ByVal blnOnlyCheckExists As Boolean) As ADODB.Recordset
'功能：获取当前病人需要停嘱打印的记录集
'参数：blnOnlyCheckExists-只检查是否存在停嘱打印
    Dim strSql As String
    
    If blnOnlyCheckExists Then
        strSql = _
            "Select 1" & vbNewLine & _
            "From 病人医嘱打印 A, 病人医嘱记录 B" & vbNewLine & _
            "Where A.医嘱id = B.ID And A.期效 = 0 And A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And a.打印时间 is not null and (B.确认停嘱时间 Is Not Null And" & vbNewLine & _
            "     Not Exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 = 2) " & _
            IIF(mlngPrintType = 1, "Or B.执行终止时间 Is Not Null And Not exists(Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 > 0)", "") & ") And Rownum<2"
    Else
        strSql = _
            "Select Distinct 页号" & vbNewLine & _
            "From (Select A.医嘱id, Max(A.页号) As 页号" & vbNewLine & _
            "       From 病人医嘱打印 A, 病人医嘱记录 B" & vbNewLine & _
            "       Where A.医嘱id = B.ID And A.期效 = 0 And A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And a.打印时间 is not null And (B.确认停嘱时间 Is Not Null And" & vbNewLine & _
            "             Not Exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 = 2) " & _
            IIF(mlngPrintType = 1, "Or B.执行终止时间 Is Not Null And Not exists(Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 > 0)", "") & ")" & vbNewLine & _
            "       Group By A.医嘱id)" & vbNewLine & _
            "Order By 页号"
    End If
    
    On Error GoTo errH
    
    Set GetStopedAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdvicePrint(ByVal intMode As Integer)
'功能：执行医嘱单打印或以预览
'参数：intMode=1-预览,2-打印
    Dim lngBegin As Long, lngEnd As Long
    Dim lng行号 As Long, strReport As String
    Dim colSegment As Collection
    Dim col常规打印 As Collection
    Dim strSql As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim intIndex As Long
    Dim varArr As Variant
    
    '确定具体的报表编号
    strReport = IIF(optReport(opt医嘱_长嘱).value, "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    
    On Error GoTo errH
    
    If tbsMain.SelectedItem.Key = "常规打印" Then '医嘱续打
        '只有在打印过的医嘱界面才能进行跳选，未打的只能连续选择
        '根据选择情况自动对页号分段
        Set colSegment = New Collection
        lngBegin = 0: lngEnd = 0
        
        For i = 1 To mintPageCount
            If imgChk(i).Visible Then
                If lngBegin = 0 Then
                    lngBegin = i: lngEnd = i
                ElseIf i = lngEnd + 1 Then
                    lngEnd = i
                Else
                    colSegment.Add lngBegin & "-" & lngEnd
                    lngBegin = i: lngEnd = i
                End If
            End If
        Next
        
        If lngBegin <> 0 And lngEnd <> 0 Then
            colSegment.Add lngBegin & "-" & lngEnd
        End If
        
        If colSegment.Count = 0 Then
            MsgBox "请选择需要打印的医嘱单页号范围。", vbInformation, gstrSysName
            Exit Sub
        ElseIf intMode = 1 And colSegment.Count > 1 Then
            MsgBox "请一次只选择一个或连续的一段页号范围进行预览。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '换页处理，可能存在换页打印的情况，再生成一次打印段
        If mstrTurnPages <> "" Then
            Set col常规打印 = New Collection
            For i = 1 To colSegment.Count '分段调用打印
            
                lngBegin = Split(colSegment(i), "-")(0)
                lngEnd = Split(colSegment(i), "-")(1)
                
                varArr = Split(mstrTurnPages, ",")
                For j = 0 To UBound(varArr)
                    If lngBegin <= Val(varArr(j)) And Val(varArr(j)) <= lngEnd Then
                        col常规打印.Add lngBegin & "-" & Val(varArr(j))
                        lngBegin = Val(varArr(j)) + 1
                    End If
                Next
                
                If lngBegin <= lngEnd Then col常规打印.Add lngBegin & "-" & lngEnd
            Next
            Set colSegment = col常规打印
        End If
        
        For i = 1 To colSegment.Count '分段调用打印
        
            mintPrintCount = 0 '用于防止预览时多次重复打印
            
            lng行号 = 0
            lngBegin = Split(colSegment(i), "-")(0)
            lngEnd = Split(colSegment(i), "-")(1)
            
            '续打处理，只会处理一次
            If mintMid = lngBegin Then
                If mbln续打页 Then '续打页，本次按续打规则处理，计算行号
                    strSql = "select max(行号)+1 as 行号 from 病人医嘱打印 where 打印时间 is not null and 病人id=[1] and 主页id=[2] and nvl(婴儿,0)=[3] and 期效=[4] and 页号=[5]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(0).value, 0, 1), lngBegin)
                    If Not rsTmp.EOF Then
                        lng行号 = Val(rsTmp!行号 & "")
                        If chkSeqPage.value = 1 And chkSeqPage.Visible Then lng行号 = 0
                    End If
                End If
            End If
            
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReport, mfrmParent, _
                "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "婴儿=" & cboBaby.ListIndex, "打印模式=" & mlngPrintType, "停嘱打印=0", "起始行号=" & lng行号, _
                "StartPageNum=" & lngBegin, "起始页号=" & lngBegin, "结束页号=" & lngEnd, "PressWorkFirst=" & IIF(lng行号 <> 0, 1, 0), intMode)
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        Next
    ElseIf tbsMain.SelectedItem.Key = "停嘱打印" Then
        '根据选择情况自动对页号分段
        Set colSegment = New Collection
        lngBegin = 0: lngEnd = 0
    
        For i = 1 To mintStopPageCount
            intIndex = 1000 + i
            If imgChk(intIndex).Visible Then
                If lngBegin = 0 Then
                    lngBegin = Val(lblNum(intIndex).Caption)
                    lngEnd = lngBegin
                ElseIf Val(lblNum(intIndex).Caption) = lngEnd + 1 Then
                    lngEnd = Val(lblNum(intIndex).Caption)
                Else
                    colSegment.Add lngBegin & "-" & lngEnd
                    lngBegin = Val(lblNum(intIndex).Caption)
                    lngEnd = lngBegin
                End If
            End If
        Next
        
        If lngBegin <> 0 Then colSegment.Add lngBegin & "-" & lngEnd

        If colSegment.Count = 0 Then
            MsgBox "请选择需要套打停止时间的医嘱单页号范围。", vbInformation, gstrSysName
            Exit Sub
        ElseIf intMode = 1 And colSegment.Count > 1 Then
            MsgBox "请一次只选择一个或连续的一段页号范围进行预览。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '调用报表输出，全部套打,并激活相应事件；在事务临时表中产生打印所需数据
        For i = 1 To colSegment.Count '分页号段调用套打
            
            mintPrintCount = 0 '用于防止预览时多次重复打印
            lngBegin = Split(colSegment(i), "-")(0): lngEnd = Split(colSegment(i), "-")(1)
            
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReport, mfrmParent, _
                "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "婴儿=" & cboBaby.ListIndex, "打印模式=" & mlngPrintType, "停嘱打印=1", "起始行号=1", _
                "StartPageNum=" & lngBegin, "起始页号=" & lngBegin, "结束页号=" & lngEnd, "PressWork=1", intMode)
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        Next
        
    End If
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
'功能：开始打印事件，初始化医嘱打印信息记录集
    
    If tbsMain.SelectedItem.Key = "常规打印" Then
        '预览时多次重复打印检查
        If mintPrintCount > 0 Then
            MsgBox "已经打印过了，要想重新打印，请使用重打功能。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        '预览时部份打印检查
        If TotalPages < 0 Then
            MsgBox "为保证有效进行续打，请选择对全部页面进行打印。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        mintPrintCount = mintPrintCount + 1
        
        Set mrsPrint = New ADODB.Recordset
        mrsPrint.Fields.Append "医嘱ID", adBigInt
        mrsPrint.Fields.Append "页号", adBigInt
        mrsPrint.Fields.Append "行号", adBigInt
        mrsPrint.CursorLocation = adUseClient
        mrsPrint.LockType = adLockOptimistic
        mrsPrint.CursorType = adOpenStatic
        mrsPrint.Open
    ElseIf tbsMain.SelectedItem.Key = "停嘱打印" Then
        '预览时多次重复打印检查
        If mintPrintCount > 0 Then
            MsgBox "已经打印过了，不能重复打印。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        '预览时部份打印检查
        If TotalPages < 0 Then
            MsgBox "为保证有效进行套打，请选择对全部页面进行打印。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        mintPrintCount = mintPrintCount + 1
        
        Set mrsPrint = New ADODB.Recordset
        mrsPrint.Fields.Append "医嘱ID", adBigInt
        mrsPrint.Fields.Append "页号", adBigInt
        mrsPrint.Fields.Append "行号", adBigInt
        mrsPrint.CursorLocation = adUseClient
        mrsPrint.LockType = adLockOptimistic
        mrsPrint.CursorType = adOpenStatic
        mrsPrint.Open
    End If
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'功能：结束打印事件，写入病人医嘱打印数据
    Dim curDate As Date, strSql As String
    
    If tbsMain.SelectedItem.Key = "常规打印" Then
        '产生医嘱打印位置记录
        curDate = zlDatabase.Currentdate
        mrsPrint.Filter = 0
        If Not mrsPrint.EOF Then
            On Error GoTo errH
            
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Do While Not mrsPrint.EOF
                strSql = "zl_病人医嘱打印_Update(" & ZVal(mrsPrint!医嘱ID) & "," & mrsPrint!页号 & "," & mrsPrint!行号 & "," & _
                    mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(0).value, 0, 1) & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.姓名 & "'," & mlngPrintType & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                
                mrsPrint.MoveNext
            Loop
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    ElseIf tbsMain.SelectedItem.Key = "停嘱打印" Then
        '标记医嘱停嘱时间已套打标志
        mrsPrint.Filter = 0
        If Not mrsPrint.EOF Then
            On Error GoTo errH
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Do While Not mrsPrint.EOF
                strSql = "Zl_病人医嘱打印_Update(" & mrsPrint!医嘱ID & "," & mrsPrint!页号 & "," & mrsPrint!行号 & ",null,null,null,null,null,null,null,1)"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                mrsPrint.MoveNext
            Loop
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    End If
 
    Set mrsPrint = Nothing
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
'功能：报表数据打印事件，记录医嘱打印行数据
'说明：当表格行无数据要打印时，是不会激活该事件的
    If tbsMain.SelectedItem.Key = "常规打印" Then
        If Page >= 1 And Row >= 1 Then
            mrsPrint.Filter = "医嘱ID=" & ID 'NULL会返回为0
            If mrsPrint.EOF Then
                mrsPrint.AddNew
                mrsPrint!医嘱ID = ID
                mrsPrint!页号 = Page
                mrsPrint!行号 = Row
            End If
            mrsPrint.Update
        End If
    ElseIf tbsMain.SelectedItem.Key = "停嘱打印" Then
        If ID > 0 And Page >= 1 And Row >= 1 Then
            mrsPrint.Filter = "医嘱ID=" & ID
            If mrsPrint.EOF Then
                mrsPrint.AddNew
                mrsPrint!医嘱ID = ID
                mrsPrint!页号 = Page
                mrsPrint!行号 = Row
                mrsPrint.Update
            End If
        End If
    End If
End Sub

Private Function GetReportRows(ByVal lngSys As Long, ByVal strReport As String, Optional ByVal intFormat As Integer = 1) As Long
'功能：获取指定报表中主要任意表格的可打印数据行数
'参数：lngSys=系统编号，为0表示共享报表
'      strReport=报表编号
'      intFormat=报表格式号,缺省为1
'返回：0表示没有任意表格
'说明：
'  1.如果报表中存在多个任意表格，则以最大的一个作为主要表格。
'  2.如果表格分栏，则可打印行数是指分栏之后的总行数。
    Dim rsTable As ADODB.Recordset
    Dim rsColumn As ADODB.Recordset
    Dim strSql As String, i As Long, j
    Dim blnHead As Boolean, blnBody As Boolean
    Dim lngBodyH As Long, lngHeadH As Long
    
    On Error GoTo errH
    
    strSql = "Select A.ID as 报表ID,B.ID,B.W,B.H,B.行高,B.分栏" & _
        " From zlReports A,zlRPTItems B" & _
        " Where A.ID=B.报表ID And B.类型=4 And Nvl(A.系统,0)=[1] And A.编号=[2] And B.格式号=[3]" & _
        " Order by B.W*B.H Desc"
    Set rsTable = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngSys, strReport, intFormat)
    If rsTable.EOF Then Exit Function
    
    strSql = "Select 序号,表头,内容 From zlRPTItems Where 报表ID=[1] And 格式号=[2] And 上级ID=[3] And 类型=6 Order by 序号"
    Set rsColumn = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(rsTable!报表ID), intFormat, Val(rsTable!ID))
    If rsColumn.EOF Then Exit Function
    
    '以下代码参照自定义报表中的方法编写
    '----------------------------------
    '求出表头高度:以第一列为准
    For i = 0 To UBound(Split(rsColumn!表头, "|"))
        lngHeadH = lngHeadH + Val(Split(Split(rsColumn!表头, "|")(i), "^")(1))
    Next
    
    '求出表体高度
    blnHead = False: blnBody = False
    rsColumn.MoveFirst
    Do While Not rsColumn.EOF
        i = UBound(Split(rsColumn!表头, "|"))
        If i > 0 Then
            blnHead = True
        ElseIf i = 0 Then
            blnHead = blnHead Or (Split(Split(rsColumn!表头, "|")(i), "^")(2) <> "#")
        End If
        blnBody = blnBody Or Not IsNull(rsColumn!内容)
        rsColumn.MoveNext
    Loop
    If Not blnHead And blnBody Then '仅有表体
        lngBodyH = rsTable!H
    Else
        If rsTable!H - lngHeadH + 15 < 0 Then
            lngBodyH = 0
        Else
            lngBodyH = rsTable!H - lngHeadH + 15
        End If
    End If
    
    '求出行数
    GetReportRows = Int(lngBodyH / rsTable!行高) * Nvl(rsTable!分栏, 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Insert打印记录()
'功能：生成将要打印的医嘱记录和要进行停嘱打印的医嘱。临嘱/长嘱，病人医嘱和婴儿是分开的，单独产生。
    Dim arrSQL As Variant
    Dim lngRows As Long
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errH
    arrSQL = Array()
    
    '病人和婴儿都要判断，在求医嘱最后序号时要考虑重整的情况，临时医嘱单不考虑转换页和打印重整标记的情况
    '判断是否要生成常规打印的记录，两层循环，j 表示期效，i 表示婴儿序号i=0时表示病人
    For j = 0 To 1
        lngRows = IIF(j = 0, mlngRows长嘱, mlngRows临嘱)
        For i = 0 To cboBaby.ListCount - 1
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & i & "," & j & ",null,null,3)"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & i & "," & j & "," & lngRows & ")"

        Next
    Next
    
    '提交数据
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub LoadAllPaper()
'功能：加载容器，所有图片纸张
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim intTmp As Integer, i As Integer
    
    On Error GoTo errH
    
    For i = 1 To 4
        Load picContainer(i)
        picContainer(i).Width = 6160
        picContainer(i).Height = 3400
    Next

    Set picContainer(1).Container = Me
    Set picContainer(3).Container = Me

    Set picContainer(2).Container = picContainer(1)
    Set picContainer(4).Container = picContainer(3)
    
    picContainer(1).Top = 1720
    picContainer(1).Left = 350
    picContainer(3).Top = 1720
    picContainer(3).Left = 350
    
    picContainer(2).Top = 0
    picContainer(4).Top = 0
    picContainer(2).Left = 0
    picContainer(4).Left = 0
    
    For i = 1 To 4
        picContainer(i).Visible = True
        picContainer(i).ZOrder 0
    Next
    
    vsc.Left = 6520
    vsc.Height = 3300
    vsc.Top = 1820
    vsc.Width = 200
    vsc.ZOrder 0
    
    strSql = "select max(a.页号) as 页数 from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID)
    
    intTmp = Val(rsTmp!页数 & "") * 2
    picPaper(0).Tag = intTmp
    picPaperB(0).Tag = intTmp
    For i = 1 To intTmp
        Call LoadPaper(0, i)
        Call LoadPaper(1, i)
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPaper(ByVal intCt As Integer, ByVal intNum As Integer)
'功能：加载图纸张，目前支持最多页数 999页
'参数：intCt容器，0－连续打印fraPrint(0)，2－停嘱打印fraPrint(1)；intNum 页号
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim intRow As Integer
    Dim intIndex As Integer
    
    On Error GoTo errH
    
    intIndex = intNum + 1000 * intCt
    
    intRow = 1 + (intNum - 1) \ 7
    
    lngLeft = ((intNum - 1) Mod 7) * (picPaper(0).Width + 200)
    lngTop = (intRow - 1) * (picPaper(0).Height + 250)
    
    '主图片
    Load picPaper(intIndex)
    Load picPaperB(intIndex)
 
    '背景图和容器图片
    Set picPaperB(intIndex).Container = picContainer(2 + intCt * 2)
    Set picPaper(intIndex).Container = picContainer(2 + intCt * 2)
 
    picPaper(intIndex).Left = lngLeft
    picPaper(intIndex).Top = lngTop
    picPaper(intIndex).Width = picPaper(0).Width
    picPaper(intIndex).Height = picPaper(0).Height
    picPaper(intIndex).BackColor = picPaper(0).BackColor
    picPaper(intIndex).Visible = False
    picPaper(intIndex).ZOrder 0

    picPaperB(intIndex).Left = picPaper(intIndex).Left + 50
    picPaperB(intIndex).Top = picPaper(intIndex).Top + 50
    picPaperB(intIndex).Width = picPaper(0).Width
    picPaperB(intIndex).Height = picPaper(0).Height
    picPaperB(intIndex).BackColor = picPaperB(0).BackColor
    picPaperB(intIndex).Visible = False
    
    '纸上的图标
    Load imgIco(intIndex)
    Set imgIco(intIndex).Container = picPaper(intIndex)
    Set imgIco(intIndex).Picture = imgIcon(0).Picture
    imgIco(intIndex).Left = (picPaper(intIndex).Width - imgIco(intIndex).Width) / 2
    imgIco(intIndex).Top = 260
    imgIco(intIndex).Visible = True
    imgIco(intIndex).ZOrder 1
    
    Load lblNum(intIndex)
    Set lblNum(intIndex).Container = picPaper(intIndex)
    lblNum(intIndex).Visible = True
    lblNum(intIndex).Caption = intNum
    lblNum(intIndex).ToolTipText = "第" & intNum & "页"
    lblNum(intIndex).FontSize = lblNum(0).FontSize
    lblNum(intIndex).Left = (picPaper(intIndex).Width - lblNum(intIndex).Width) / 2
    lblNum(intIndex).Top = imgIco(intIndex).Height + imgIco(intIndex).Top + 10
    lblNum(intIndex).BackColor = picPaper(0).BackColor
    
    '勾选图片，程序中控件可见性
    Load imgChk(intIndex)
    Set imgChk(intIndex).Container = picPaper(intIndex)
    Set imgChk(intIndex).Picture = imgChk(0).Picture '固定
    imgChk(intIndex).Width = 240
    imgChk(intIndex).Height = 240
    imgChk(intIndex).Left = picPaper(0).Width - imgChk(intIndex).Width
    imgChk(intIndex).Top = -10
    imgChk(intIndex).Visible = False
    imgChk(intIndex).ZOrder 1
    
    Exit Sub
errH:
    If 1 = 2 Then
        Resume
    End If
    err.Clear
End Sub
   
Private Sub UnLoadPaper()
    Dim i As Integer
    
    On Error Resume Next
    
    '先卸载容器内的制件再卸载容器
    For i = 1 To Val(picPaper(0).Tag)
        Unload imgChk(i)
        Unload imgIco(i)
        Unload lblNum(i)
        Unload picPaperB(i)
        Unload picPaper(i)
    Next
    
    For i = 1 To Val(picPaperB(0).Tag)
        Unload imgChk(i + 1000)
        Unload imgIco(i + 1000)
        Unload lblNum(i + 1000)
        Unload picPaperB(i + 1000)
        Unload picPaper(i + 1000)
    Next
    
    For i = 1 To 4
        Unload picContainer(i)
    Next
    
    err.Clear
End Sub

Private Sub imgIco_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub lblNum_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub imgChk_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub picPaper_Click(Index As Integer)
    Dim blnTmp As Boolean
    Dim i As Integer
    
    blnTmp = imgChk(Index).Visible
    imgChk(Index).Visible = Not blnTmp
    
    If Not (Index > 1000 Or mintMid = 0 Or mintMid > Index) Then
        If blnTmp Then
            For i = Index + 1 To mintPageCount
                imgChk(i).Visible = imgChk(Index).Visible
            Next
        Else
            For i = mintMid To Index - 1
                imgChk(i).Visible = imgChk(Index).Visible
            Next
        End If
    End If
    
    If mbln续打页 And Index < 1000 Then
        If mintMid = 1 Then
            chkSeqPage.Visible = imgChk(mintMid).Visible
        Else
            chkSeqPage.Visible = imgChk(mintMid).Visible And Not imgChk(mintMid - 1).Visible
        End If
    End If
End Sub

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub vsc_Scroll()
    If tbsMain.SelectedItem.Key = "常规打印" Then
        picContainer(pic连打_纸面).Top = (-1) * vsc.value * Screen.TwipsPerPixelY
    Else
        picContainer(pic套打_纸面).Top = (-1) * vsc.value * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub cmdClsLastPrint_Click()
'功能：清除打印记录
    Call ClearPrintRs(True)
End Sub

Private Sub txtClearPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtClearPage_Validate(False)
        cmdClear.SetFocus
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtClearPage_Validate(Cancel As Boolean)
    Dim lngTmp As Long
    lngTmp = Val(txtClearPage.Text)
    If lngTmp = 0 Then
        txtClearPage.Text = 1
    ElseIf lngTmp > mlngPrintedMaxPage Then
        txtClearPage.Text = mlngPrintedMaxPage
    Else
        txtClearPage.Text = lngTmp
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearPrintRs(False)
End Sub

Private Sub ClearPrintRs(ByVal bln上次打印 As Boolean)
'功能：从某页开始清除打印
'参数：bln上次打印 －true 清除上次打印，false 清除指定页
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As ADODB.Recordset
    Dim str位置 As String, str打印人 As String
    Dim lngTmp As Long, strTmp As String, str打印时间 As String
    Dim arrSQL As Variant
    Dim lngRows As Long
    Dim i As Long
    Dim lng页号 As Long
    
    If MsgBox("确实要清除已打印的医嘱记录吗，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    arrSQL = Array()
    
    On Error GoTo errH
    
    lngRows = IIF(optReport(opt医嘱_长嘱).value, mlngRows长嘱, mlngRows临嘱)
    
    If bln上次打印 Then
        If optReport(opt医嘱_长嘱).value Then
            '进行重整时间判断
            strSql = "select max(打印时间) as 时间 from 病人医嘱打印 Where 病人id=[1] And 主页id=[2] And Nvl(婴儿,0)=[3] And 期效=0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex)
            If Not IsNull(rsTmp!时间) Then
                If mdat重整时间 > rsTmp!时间 Then
                    MsgBox "上次打印在重整之前，要先回退重整才能清除打印。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        strSql = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & ",null,null,1)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSql
         
        '生成记录
        strSql = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & lngRows & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSql
    Else
        If mdat重整时间 <> CDate("1900-01-01") And optReport(opt医嘱_长嘱).value Then
            strSql = "select max(页号) as 页号 from 病人医嘱打印 a " & _
                " where a.病人id=[1] and a.主页id=[2] and nvl(a.婴儿,0)=[3] and a.期效=[4] and 打印时间<[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, 0, mdat重整时间)
            lng页号 = Val(rsTmp!页号 & "")
        End If
    
        If Val(txtClearPage.Text) <= lng页号 Then
            If MsgBox("清除打印的医嘱单中包含了重整前打过的内容，若要清除请先回退重整操作，选 是 则最清最近一次重整之后打印的内容，选 否 则不清除，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        strSql = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & Val(txtClearPage.Text) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSql
        
        '生成记录
        strSql = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & lngRows & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSql
        
        If Val(txtClearPage.Text) > 1 Then
            strSql = "Select a.打印时间, a.打印人" & vbNewLine & _
                " From 病人医嘱打印 A,病人医嘱记录 b Where a.医嘱id=b.id and b.诊疗类别 in ('5','6') and a.病人id =[1] And a.主页id =[2]" & _
                " and a.婴儿=[3] And a.期效 =[4] And a.页号 =[5] and a.行号=[6]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(opt医嘱_长嘱).value, 0, 1), _
                Val(txtClearPage.Text) - 1, lngRows)
            If Not rsTmp.EOF Then
                str打印人 = rsTmp!打印人 & ""
                strTmp = Format(rsTmp!打印时间, "yyyy-MM-dd HH:mm:ss")
                strTmp = "To_Date('" & strTmp & "','YYYY-MM-DD HH24:MI:SS')"
                str打印时间 = strTmp
            End If
        End If
    End If
    
    '提交数据
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
    
    If str打印人 <> "" Then
        strSql = "Select a.医嘱id,a.页号,a.行号 From 病人医嘱打印 A Where a.打印时间 is null and a.病人id=[1] And a.主页id=[2] and a.婴儿=[3] And a.期效=[4] And a.页号=[5]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(opt医嘱_长嘱).value, 0, 1), Val(txtClearPage.Text) - 1)
        If Not rsTmp.EOF Then
            arrSQL = Array()
            For i = 1 To rsTmp.RecordCount
                strSql = "zl_病人医嘱打印_Update(" & ZVal(rsTmp!医嘱ID) & "," & rsTmp!页号 & "," & rsTmp!行号 & "," & _
                    mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & _
                    str打印时间 & ",'" & str打印人 & "'," & mlngPrintType & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSql
                rsTmp.MoveNext
            Next
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    End If
    
    Call RefreshFace
    
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshPrintPage(Optional ByVal blnAll As Boolean)
'功能：重新生成示打印的记录，刷新界面
'参数：blnAll 是否刷新本病人的所有医嘱单，true 刷新所有，false 只刷当前页
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim lngTmp As Long
    Dim i As Long, j As Long
    Dim lng期效 As Long
    Dim lng婴儿 As Long
    Dim arrSQL As Variant
    Dim lngRows As Long
    
    If Not Me.Visible Then Exit Sub
    
    On Error GoTo errH
 
    Set mrsPrint = Nothing
    mbln续打页 = False
    mintPageCount = 0
    mintStopPageCount = 0
 
 
    lng婴儿 = cboBaby.ListCount - 1
    lng期效 = IIF(optReport(opt医嘱_长嘱).value, 0, 1)
    arrSQL = Array()
    
    If blnAll Then
        For j = 0 To 1
            lngRows = IIF(j = 0, mlngRows长嘱, mlngRows临嘱)
            For i = 0 To cboBaby.ListCount - 1
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & i & "," & j & ",null,null,3)"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & i & "," & j & "," & lngRows & ")"
            Next
        Next
    Else
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & lng婴儿 & "," & lng期效 & ",null,null,3)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & lng婴儿 & "," & lng期效 & "," & IIF(lng期效 = 0, mlngRows长嘱, mlngRows临嘱) & ")"
    End If
    
    '提交数据
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
    
    '判断是否还要添加纸张
    strSql = "select max(a.页号) as 页数 from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID)
    lngTmp = Val(rsTmp!页数 & "") * 2
    If Val(picPaper(0).Tag) < lngTmp Then
        For i = Val(picPaper(0).Tag) + 1 To lngTmp
            Call LoadPaper(0, i)
            Call LoadPaper(1, i)
        Next
    End If
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkCQSHHY_Click()
    Call zlDatabase.SetPara("长嘱单术后换页", chkPrintTurnPage.value, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
End Sub

Private Sub chkCZCZHY_Click()
    Call zlDatabase.SetPara("长嘱单重整换页", chkPrintTurnPage.value, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
End Sub
