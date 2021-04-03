VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSublimeStationSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9585
   Icon            =   "frmSublimeStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt入院天数 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   6555
      MaxLength       =   2
      TabIndex        =   26
      Text            =   "3"
      Top             =   2850
      Width           =   300
   End
   Begin VB.Frame fraFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   3
      Left            =   6570
      TabIndex        =   48
      Top             =   3030
      Width           =   300
   End
   Begin VB.Frame fraSplit 
      Height          =   135
      Left            =   -30
      TabIndex        =   47
      Top             =   4560
      Width           =   9840
   End
   Begin VB.PictureBox picControl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   1770
      ScaleHeight     =   1785
      ScaleWidth      =   2295
      TabIndex        =   38
      Top             =   2670
      Visible         =   0   'False
      Width           =   2295
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1497
         Width           =   200
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   60
         Picture         =   "frmSublimeStationSetup.frx":000C
         ScaleHeight     =   1350
         ScaleWidth      =   2160
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   90
         Width           =   2160
         Begin VB.Shape shpBorder 
            BorderColor     =   &H00C56A31&
            FillColor       =   &H00FF8080&
            Height          =   270
            Left            =   1890
            Top             =   1080
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape shpValue 
            BorderColor     =   &H00C56A31&
            FillColor       =   &H00FF8080&
            Height          =   270
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   270
         End
      End
      Begin VB.Label lblColor 
         Caption         =   "&HFFFFFF"
         Height          =   195
         Left            =   405
         TabIndex        =   42
         Top             =   1500
         UseMnemonic     =   0   'False
         Width           =   1365
      End
   End
   Begin VB.Frame fraMedRec 
      Caption         =   "病案审查反馈设置"
      Height          =   600
      Left            =   4995
      TabIndex        =   23
      Top             =   2115
      Width           =   4485
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1025
         TabIndex        =   44
         Top             =   420
         Width           =   300
      End
      Begin VB.TextBox txtMedRec 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   1040
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "1"
         Top             =   240
         Width           =   300
      End
      Begin VB.Label lblMedRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "显示    天内的病案审查反馈数"
         Height          =   180
         Left            =   645
         TabIndex        =   45
         Top             =   255
         Width           =   2520
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 护理等级颜色"
      Height          =   1530
      Left            =   120
      TabIndex        =   14
      Top             =   2820
      Width           =   4755
      Begin VB.Image img护理等级 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   3
         Left            =   3840
         Picture         =   "frmSublimeStationSetup.frx":0782
         Stretch         =   -1  'True
         Top             =   900
         Width           =   345
      End
      Begin VB.Image img护理等级 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   2
         Left            =   1770
         Picture         =   "frmSublimeStationSetup.frx":0E84
         Stretch         =   -1  'True
         Top             =   900
         Width           =   345
      End
      Begin VB.Image img护理等级 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   3840
         Picture         =   "frmSublimeStationSetup.frx":1586
         Stretch         =   -1  'True
         Top             =   420
         Width           =   345
      End
      Begin VB.Image img护理等级 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1770
         Picture         =   "frmSublimeStationSetup.frx":1C88
         Stretch         =   -1  'True
         Top             =   420
         Width           =   345
      End
      Begin VB.Label lbl护理等级 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "三级护理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2610
         TabIndex        =   30
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl护理等级 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "二级护理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   540
         TabIndex        =   29
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl护理等级 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一级护理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2610
         TabIndex        =   28
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lbl护理等级 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "特级护理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   540
         TabIndex        =   27
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7185
      TabIndex        =   36
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8400
      TabIndex        =   37
      Top             =   4785
      Width           =   1100
   End
   Begin VB.Frame fraAdvice 
      Caption         =   " 医嘱提醒设置 "
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4755
      Begin VB.CheckBox chkWarn 
         Caption         =   "标本拒收暂未启用"
         Enabled         =   0   'False
         Height          =   555
         Index           =   10
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "备血完成"
         Height          =   195
         Index           =   11
         Left            =   1395
         TabIndex        =   51
         Top             =   1635
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "血袋回收"
         Height          =   195
         Index           =   12
         Left            =   2520
         TabIndex        =   50
         Top             =   1635
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "取血通知"
         Height          =   195
         Index           =   9
         Left            =   3670
         TabIndex        =   49
         Top             =   1380
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RIS预约准备"
         Height          =   195
         Index           =   8
         Left            =   2355
         TabIndex        =   11
         Top             =   1380
         Width           =   1320
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RIS预约"
         Height          =   195
         Index           =   7
         Left            =   1395
         TabIndex        =   10
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "启用语音提示"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   2160
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "语音设置(&S)"
         Height          =   350
         Left            =   1860
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "危急值"
         Height          =   195
         Index           =   4
         Left            =   1395
         TabIndex        =   7
         Top             =   1110
         Width           =   870
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "输液拒绝"
         Height          =   195
         Index           =   5
         Left            =   2355
         TabIndex        =   8
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "销帐申请"
         Height          =   195
         Index           =   6
         Left            =   3480
         TabIndex        =   9
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox ChkCollate 
         Caption         =   "医嘱处理后自动定位到病人医嘱页面"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   1890
         Width           =   3900
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "安排"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   6
         Top             =   855
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "新开"
         Height          =   195
         Index           =   0
         Left            =   1395
         TabIndex        =   3
         Top             =   855
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "新停"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   4
         Top             =   855
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "新废"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   5
         Top             =   855
         Width           =   660
      End
      Begin VB.TextBox txtNotifyAdvice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "10"
         Top             =   315
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdvice 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   32
         Top             =   495
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdviceDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   31
         Top             =   765
         Width           =   300
      End
      Begin VB.TextBox txtNotifyAdviceDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "1"
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chkNotifyAdvice 
         Caption         =   "每    分钟自动刷新医嘱提醒区域中的内容"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label lbl提醒内容 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒内容:"
         Height          =   180
         Left            =   570
         TabIndex        =   35
         Top             =   855
         Width           =   810
      End
      Begin VB.Label lblNotifyAdviceDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "将    天内处理的医嘱病人显示在提醒区域"
         Height          =   180
         Left            =   570
         TabIndex        =   34
         Top             =   600
         Width           =   3420
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Left            =   360
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   41
      Top             =   3165
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   " 个性化过滤条件 "
      Height          =   690
      Left            =   4995
      TabIndex        =   15
      Top             =   75
      Width           =   4485
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1320
         TabIndex        =   43
         Top             =   495
         Width           =   300
      End
      Begin VB.TextBox txt入科天数 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   1305
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "3"
         Top             =   315
         Width           =   300
      End
      Begin VB.CheckBox chkPatientFilter 
         Caption         =   "提取入科    天内的住院病人"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   315
         Width           =   3900
      End
   End
   Begin VB.Frame FraCard 
      Caption         =   " 卡片标签设置 "
      Height          =   1095
      Left            =   4995
      TabIndex        =   18
      Top             =   900
      Width           =   4485
      Begin VB.OptionButton optNewCard 
         Caption         =   "床位号"
         Height          =   180
         Index           =   0
         Left            =   1290
         TabIndex        =   21
         Top             =   675
         Width           =   945
      End
      Begin VB.OptionButton optNewCard 
         Caption         =   "床位编制编号+床位号"
         Height          =   180
         Index           =   1
         Left            =   2400
         TabIndex        =   22
         Top             =   675
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CheckBox chkAmount 
         Caption         =   "卡片余额将担保金额计算在内"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label lblNewCard 
         AutoSize        =   -1  'True
         Caption         =   "卡片排序按"
         Height          =   180
         Left            =   300
         TabIndex        =   20
         Top             =   660
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList img24 
      Left            =   225
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CheckBox chkNewPati 
      Caption         =   "待入科列表显示    天内登记的住院病人"
      Height          =   195
      Left            =   5010
      TabIndex        =   25
      Top             =   2850
      Width           =   3900
   End
End
Attribute VB_Name = "frmSublimeStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarColor As OLE_COLOR
Public mstrPrivs As String
Private mlngModual As Long

Private Const ALTERNATE = 1
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" _
    (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" _
    (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Polyline Lib "gdi32" _
    (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'设定一个窗体捕获鼠标，即所有鼠标输入消息都发往该窗体
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'取消鼠标捕获
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private mlngColor As Long
Private mintIndex As Long
Private mobjFileSys As New FileSystemObject

Public Sub ShowMe()
    '由新版住院护士工作站调用，显示标注按钮
    mintIndex = 0
    Me.Show vbModal
End Sub

Private Sub chkNewPati_Click()
    On Error Resume Next
    If chkNewPati.Value = 1 Then
        txt入院天数.Enabled = True
        txt入院天数.SetFocus
    Else
        txt入院天数.Enabled = False
        txt入院天数.Text = ""
    End If
End Sub

Private Sub chkNotifyAdvice_Click()
    txtNotifyAdvice.Enabled = chkNotifyAdvice.Value = 1
    If Visible And txtNotifyAdvice.Enabled Then txtNotifyAdvice.SetFocus
End Sub

Private Sub chkPatientFilter_Click()
    On Error Resume Next
    If chkPatientFilter.Value = 1 Then
        txt入科天数.Enabled = True
        txt入科天数.SetFocus
    Else
        txt入科天数.Enabled = False
        txt入科天数.Text = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    Dim strTmp As String
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If chkNotifyAdvice.Value = 1 And Val(txtNotifyAdvice.Text) = 0 Then
        If txtNotifyAdvice.Text = "" Then
            MsgBox "请设置医嘱提醒的自动刷新间隔。", vbInformation, gstrSysName
        Else
            MsgBox "医嘱提醒的自动刷新间隔至少应为1分钟。", vbInformation, gstrSysName
        End If
        txtNotifyAdvice.SetFocus: Exit Sub
    End If
    If Val(txtNotifyAdviceDay.Text) = 0 Then
        If txtNotifyAdviceDay.Text = "" Then
            MsgBox "请设置要提醒的医嘱天数。", vbInformation, gstrSysName
        Else
            MsgBox "要提醒的医嘱天数至少应为1天。", vbInformation, gstrSysName
        End If
        txtNotifyAdviceDay.SetFocus: Exit Sub
    End If
    If chkPatientFilter.Value = 1 Then
        If Trim(txt入科天数.Text) = "" Then
            MsgBox "请输入入科天数条件！", vbInformation, gstrSysName
            txt入科天数.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt入科天数.Text) Then
            MsgBox "入科天数中含有非法字符！（只能输入数字）", vbInformation, gstrSysName
            txt入科天数.SetFocus
            Exit Sub
        End If
        If Val(txt入科天数.Text) <= 0 Then
            MsgBox "入科天数必须大于零！", vbInformation, gstrSysName
            txt入科天数.SetFocus
            Exit Sub
        End If
    End If
    
    '73646
    If txtMedRec.Text = "" Then
        MsgBox "请设置病理审查反馈提醒的天数。", vbInformation, gstrSysName
        txtMedRec.SetFocus: Exit Sub
    End If
    
    If chkNewPati.Value = 1 Then
        If Trim(txt入院天数.Text) = "" Then
            MsgBox "请输入待入科显示的入院登记天数！", vbInformation, gstrSysName
            txt入院天数.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt入院天数.Text) Then
            MsgBox "待入科显示的入院登记天数中含有非法字符！（只能输入数字）", vbInformation, gstrSysName
            txt入院天数.SetFocus
            Exit Sub
        End If
        If Val(txt入院天数.Text) <= 0 Then
            MsgBox "待入科显示的入院登记天数必须大于零！", vbInformation, gstrSysName
            txt入院天数.SetFocus
            Exit Sub
        End If
    End If
    
    '自动刷新医嘱提醒
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("自动刷新医嘱间隔", IIf(chkNotifyAdvice.Value = 1, Val(txtNotifyAdvice.Text), ""), glngSys, p住院护士站, blnSetup)
    Call zlDatabase.SetPara("自动刷新医嘱天数", Val(txtNotifyAdviceDay.Text), glngSys, p住院护士站, blnSetup)
    strTmp = ""
    For i = 0 To chkWarn.UBound
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zlDatabase.SetPara("自动刷新医嘱类型", strTmp, glngSys, p住院护士站, blnSetup)
    
    '保存入科天数
    If chkPatientFilter.Value = 1 Then
        Call zlDatabase.SetPara("入科天数", txt入科天数.Text, glngSys, 1265, blnSetup)
    Else
        Call zlDatabase.SetPara("入科天数", "0", glngSys, 1265, blnSetup)
    End If
    '保存入院天数 111016
    If chkNewPati.Value = 1 Then
        Call zlDatabase.SetPara("入院天数", txt入院天数.Text, glngSys, 1265, blnSetup)
    Else
        Call zlDatabase.SetPara("入院天数", "0", glngSys, 1265, blnSetup)
    End If
    '保存护理等级的颜色
    Call zlDatabase.SetPara("特级护理颜色", img护理等级(0).Tag, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("一级护理颜色", img护理等级(1).Tag, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("二级护理颜色", img护理等级(2).Tag, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("三级护理颜色", img护理等级(3).Tag, glngSys, 1265, blnSetup)
    '--56960:刘鹏飞,2013-01-17,添加参数"卡片余额含担保金额"
    Call zlDatabase.SetPara("卡片余额含担保金额", chkAmount.Value, glngSys, 1265, blnSetup)
    '54370:刘鹏飞,2013-05-02,添加参数"医嘱处理后自动定位到医嘱页面"
    Call zlDatabase.SetPara("医嘱处理后自动定位到医嘱页面", ChkCollate.Value, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("病案审查反馈天数", txtMedRec.Text, glngSys, p住院护士站, blnSetup)
    '92852:刘鹏飞,2016-01-21,床位卡排序处理
    Call zlDatabase.SetPara("床位卡片排序方式", IIf(optNewCard(0).Value = True, 0, 1), glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("启用语音提示", chkSound.Value, glngSys, p住院护士站, blnSetup)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyCode = vbKeyEscape Then
        ReleaseCapture
        picControl.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strPar As String
    Dim intType As Integer
    
    gblnOK = False
    mlngModual = p住院护士站
    
    chkWarn(9).Visible = gbln血库系统
    chkWarn(11).Visible = gbln血库系统
    '自动刷新医嘱提醒
    strPar = zlDatabase.GetPara("自动刷新医嘱间隔", glngSys, mlngModual, , Array(chkNotifyAdvice), InStr(mstrPrivs, "参数设置") > 0, intType)
    If Val(strPar) > 0 Then
        chkNotifyAdvice.Value = 1: txtNotifyAdvice.Text = Val(strPar)
    End If
    '前面事件中会自动可用，因此后面强制设置
    If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0 Then
        txtNotifyAdvice.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("自动刷新医嘱天数", glngSys, mlngModual, 1, Array(lblNotifyAdviceDay, txtNotifyAdviceDay), InStr(mstrPrivs, "参数设置") > 0)
    txtNotifyAdviceDay.Text = Val(strPar)
    
    strPar = zlDatabase.GetPara("自动刷新医嘱类型", glngSys, mlngModual, "0000000000000", Array(lbl提醒内容, chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), chkWarn(6), chkWarn(7), chkWarn(8), chkWarn(9), chkWarn(10), chkWarn(11), chkWarn(12)), InStr(mstrPrivs, "参数设置") > 0)
    For i = 1 To Len(strPar)
        If i - 1 <= chkWarn.UBound Then
            chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
        End If
    Next
    txt入科天数.Text = zlDatabase.GetPara("入科天数", glngSys, 1265, "3", Array(chkPatientFilter, txt入科天数))
    chkPatientFilter.Value = IIf(Val(txt入科天数.Text) = 0, 0, 1)
    txt入科天数.Enabled = (chkPatientFilter.Value = 1)
    '111016
    txt入院天数.Text = zlDatabase.GetPara("入院天数", glngSys, 1265, "0", Array(chkNewPati, txt入院天数))
    chkNewPati.Value = IIf(Val(txt入院天数.Text) = 0, 0, 1)
    txt入院天数.Enabled = (chkNewPati.Value = 1)
    '--56960:刘鹏飞,2013-01-17,添加参数"卡片余额含担保金额"
    strPar = zlDatabase.GetPara("卡片余额含担保金额", glngSys, 1265, 0, Array(chkAmount), InStr(mstrPrivs, "参数设置") > 0)
    chkAmount.Value = IIf(Val(strPar) = 1, 1, 0)
    '54370:刘鹏飞,2013-05-02,添加参数"医嘱处理后自动定位到医嘱页面"
    strPar = zlDatabase.GetPara("医嘱处理后自动定位到医嘱页面", glngSys, 1265, 0, Array(ChkCollate), InStr(mstrPrivs, "参数设置") > 0)
    ChkCollate.Value = IIf(Val(strPar) = 1, 1, 0)
    strPar = zlDatabase.GetPara("启用语音提示", glngSys, mlngModual, 0, Array(chkSound, cmdSoundSet), InStr(mstrPrivs, "参数设置") > 0)
    chkSound.Value = IIf(Val(strPar) = 1, 1, 0)
    txtMedRec.Text = zlDatabase.GetPara("病案审查反馈天数", glngSys, mlngModual, "3", Array(lblMedRec, txtMedRec), InStr(mstrPrivs, "参数设置") > 0)
    strPar = zlDatabase.GetPara("床位卡片排序方式", glngSys, 1265, 0, Array(lblNewCard, optNewCard(0), optNewCard(1)), InStr(mstrPrivs, "参数设置") > 0)
    If Val(strPar) = 0 Then
        optNewCard(0).Value = True
    Else
        optNewCard(1).Value = True
    End If
    Call InitColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DeleteFile
End Sub

Private Sub InitColor()
    Dim strValue As String
    Dim lng特级 As Long, lng一级 As Long, lng二级 As Long, lng三级 As Long
    Const c紫色 As Long = 8388736
    Const c红色 As Long = 255
    Const c兰色 As Long = 16711680
    Const c白色 As Long = 16777215
    
    Call DeleteFile
    '读取护理等级现有设置(无则取缺省数据)
    strValue = zlDatabase.GetPara("特级护理颜色", glngSys, 1265, "", Array(lbl护理等级(0)))
    lng特级 = IIf(strValue = "", c紫色, Val(strValue))
    strValue = zlDatabase.GetPara("一级护理颜色", glngSys, 1265, "", Array(lbl护理等级(1)))
    lng一级 = IIf(strValue = "", c红色, Val(strValue))
    strValue = zlDatabase.GetPara("二级护理颜色", glngSys, 1265, "", Array(lbl护理等级(2)))
    lng二级 = IIf(strValue = "", c兰色, Val(strValue))
    strValue = zlDatabase.GetPara("三级护理颜色", glngSys, 1265, "", Array(lbl护理等级(3)))
    lng三级 = IIf(strValue = "", c白色, Val(strValue))
    
    '绘图
    mlngColor = lng特级
    Call DrawPoly
    img护理等级(0).Tag = mlngColor
    img护理等级(0).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lng一级
    Call DrawPoly
    img护理等级(1).Tag = mlngColor
    img护理等级(1).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lng二级
    Call DrawPoly
    img护理等级(2).Tag = mlngColor
    img护理等级(2).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lng三级
    Call DrawPoly
    img护理等级(3).Tag = mlngColor
    img护理等级(3).Picture = img24.ListImages("K_" & mintIndex).Picture
End Sub

Private Sub img护理等级_Click(Index As Integer)
    picControl.Tag = Index
    picControl.Visible = True
    Call SetCOLOR(Val(img护理等级(Index).Tag))
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > 0 And x < Picture1.ScaleWidth And y > 0 And y < Picture1.ScaleHeight Then
        SetCapture Picture1.hwnd
        shpBorder.Visible = True
    Else
        ReleaseCapture
        shpBorder.Visible = False
    End If

    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = y \ (18 * Screen.TwipsPerPixelY)
    lCol = x \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    shpBorder.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If Picture1.Point(lX, lY) = -1 Then Exit Sub
    picColor.BackColor = Picture1.Point(lX, lY)
    Select Case CStr(Hex(picColor.BackColor))
    Case "0"
        lblColor = "黑色"
    Case "3399"
        lblColor = "褐色"
    Case "3333"
        lblColor = "橄榄色"
    Case "3300"
        lblColor = "深绿"
    Case "663300"
        lblColor = "深青"
    Case "800000"
        lblColor = "深蓝"
    Case "993333"
        lblColor = "靛蓝"
    Case "333333"
        lblColor = "灰色-80%"
    Case "80"
        lblColor = "深红"
    Case "66FF"
        lblColor = "橙色"
    Case "8080"
        lblColor = "深黄"
    Case "8000"
        lblColor = "绿色"
    Case "808000"
        lblColor = "青色"
    Case "FF0000"
        lblColor = "蓝色"
    Case "996666"
        lblColor = "蓝-灰"
    Case "808080"
        lblColor = "灰色-50%"
    Case "FF"
        lblColor = "红色"
    Case "99FF"
        lblColor = "浅橙色"
    Case "CC99"
        lblColor = "酸橙色"
    Case "669933"
        lblColor = "海绿"
    Case "CCCC33"
        lblColor = "水绿色"
    Case "FF6633"
        lblColor = "浅蓝"
    Case "800080"
        lblColor = "紫罗兰"
    Case "999999"
        lblColor = "灰色-40%"
    Case "FF00FF"
        lblColor = "粉红"
    Case "CCFF"
        lblColor = "金色"
    Case "FFFF"
        lblColor = "黄色"
    Case "FF00"
        lblColor = "鲜绿"
    Case "FFFF00"
        lblColor = "青绿"
    Case "FFCC00"
        lblColor = "天蓝"
    Case "663399"
        lblColor = "梅红"
    Case "C0C0C0"
        lblColor = "灰色-25%"
    Case "CC99FF"
        lblColor = "玫瑰红"
    Case "99CCFF"
        lblColor = "茶色"
    Case "99FFFF"
        lblColor = "浅黄"
    Case "CCFFCC"
        lblColor = "浅绿"
    Case "FFFFCC"
        lblColor = "浅青绿"
    Case "FFCC99"
        lblColor = "淡蓝"
    Case "FF99CC"
        lblColor = "淡紫"
    Case "FFFFFF"
        lblColor = "白色"
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = y \ (18 * Screen.TwipsPerPixelY)
    lCol = x \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    picControl.Visible = False
    
    '按指定颜色作图
    mlngColor = picColor.BackColor
    img护理等级(Val(picControl.Tag)).Tag = mlngColor
    Call DrawPoly
    img护理等级(Val(picControl.Tag)).Picture = img24.ListImages("K_" & mintIndex).Picture
End Sub

Private Sub txtMedRec_GotFocus()
    Call zlControl.TxtSelAll(txtMedRec)
End Sub

Private Sub txtMedRec_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdvice_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdvice)
End Sub

Private Sub txtNotifyAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdviceDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdviceDay)
End Sub

Private Sub txtNotifyAdviceDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub SetCOLOR(vData As OLE_COLOR)
    mvarColor = vData
    Dim lRow As Long, lCol As Long
    shpValue.Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor = "黑色"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor = "褐色"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor = "橄榄色"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor = "深绿"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor = "深青"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor = "深蓝"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor = "靛蓝"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor = "灰色-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor = "深红"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor = "橙色"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor = "深黄"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor = "绿色"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor = "青色"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor = "蓝色"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor = "蓝-灰"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor = "灰色-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor = "红色"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor = "浅橙色"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor = "酸橙色"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor = "海绿"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor = "水绿色"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor = "浅蓝"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor = "紫罗兰"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor = "灰色-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor = "粉红"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor = "金色"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor = "黄色"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor = "鲜绿"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor = "青绿"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor = "天蓝"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor = "梅红"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor = "灰色-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor = "玫瑰红"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor = "茶色"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor = "浅黄"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor = "浅绿"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor = "浅青绿"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor = "淡蓝"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor = "淡紫"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor = "白色"
        lRow = 4
        lCol = 7
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
    shpBorder.Visible = False
    shpValue.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    shpValue.Visible = True
    If vData = tomAutoColor Or vData = -1 Then
    
    Else
        picColor.BackColor = vData
    End If
End Sub

Private Sub AddColor()
    Dim strFile As String
    mintIndex = mintIndex + 1
    '不保存为文件,当创建多个图片时,加入到imagelist里的始终只有最后一个,应该是由于image中保存的是图片ID造成
    
    strFile = App.Path & "\HLDJTMP" & mintIndex & ".BMP"
    SavePicture picDraw.Image, strFile
    picDraw.Picture = LoadPicture(strFile)
    img24.ListImages.Add , "K_" & mintIndex, picDraw.Picture
End Sub

Private Sub DrawPoly()
    Dim lngRgn As Long, lngBrush As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim PtInPoly() As POINTAPI

    '填充区域并划边线
    ReDim PtInPoly(4) As POINTAPI
    PtInPoly(1).x = 0
    PtInPoly(1).y = 0
    PtInPoly(2).x = picDraw.ScaleWidth
    PtInPoly(2).y = 0
    PtInPoly(3).x = picDraw.ScaleWidth
    PtInPoly(3).y = picDraw.ScaleHeight
    PtInPoly(4).x = PtInPoly(1).x
    PtInPoly(4).y = PtInPoly(1).y
    
    '创建系统刷子
    picDraw.Cls
    lngBrush = CreateSolidBrush(mlngColor)

    '如果创建刷子成功,才选入
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn picDraw.hdc, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    End If
    picDraw.Refresh
    
    Call AddColor
End Sub

Private Sub DeleteFile()
    Dim objFile As File
    For Each objFile In mobjFileSys.GetFolder(App.Path).Files
        If Left(objFile.Name, 7) = "HLDJTMP" Then
            mobjFileSys.DeleteFile objFile.Path, True
        End If
    Next
End Sub

Private Sub txt入科天数_GotFocus()
    txt入科天数.SelStart = 0
    txt入科天数.SelLength = txt入科天数.MaxLength
End Sub

Private Sub txt入科天数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt入院天数_GotFocus()
    txt入院天数.SelStart = 0
    txt入院天数.SelLength = txt入院天数.MaxLength
End Sub

Private Sub txt入院天数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
