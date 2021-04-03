VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSetCourse 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ClipControls    =   0   'False
   Icon            =   "frmSetCourse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin VB.CheckBox chkMedicalTeam 
      Caption         =   "入住时必须指定医疗小组"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   4095
      Width           =   4200
   End
   Begin VB.CheckBox chkDeath 
      Caption         =   "出院时，下达了死亡医嘱才允许死亡出院"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   4200
   End
   Begin VB.Frame fraWristlet 
      Caption         =   "病人腕带"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   5220
      Width           =   4935
      Begin VB.OptionButton optWristletPrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   2205
         TabIndex        =   23
         Top             =   285
         Width           =   1500
      End
      Begin VB.OptionButton optWristletPrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   1065
         TabIndex        =   22
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton optWristletPrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   21
         Top             =   285
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "打印设置"
         Height          =   345
         Index           =   0
         Left            =   3690
         TabIndex        =   24
         Top             =   160
         Width           =   990
      End
   End
   Begin VB.CheckBox chkInTime 
      Caption         =   "入院入住时允许修改入院时间"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   3825
      Width           =   4200
   End
   Begin VB.Frame fraBabyWristlet 
      Caption         =   "婴儿腕带"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   5940
      Width           =   4935
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "打印设置"
         Height          =   345
         Index           =   1
         Left            =   3690
         TabIndex        =   29
         Top             =   160
         Width           =   990
      End
      Begin VB.OptionButton optBabyWristletPrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   285
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optBabyWristletPrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   1065
         TabIndex        =   27
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton optBabyWristletPrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   2205
         TabIndex        =   28
         Top             =   285
         Width           =   1500
      End
   End
   Begin VB.CheckBox chkChangeIn 
      Caption         =   "转科入住时护理等级默认为空"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   3555
      Width           =   4200
   End
   Begin VB.CheckBox chkIn 
      Caption         =   "入院入住，允许调整入院科室"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3270
      Width           =   4200
   End
   Begin VB.CheckBox chkAllowOut 
      Caption         =   "出院时，提取入院诊断为默认的出院诊断"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2720
      Width           =   4200
   End
   Begin VB.Frame Fra入科时间 
      Caption         =   "入住时间"
      Height          =   705
      Left            =   120
      TabIndex        =   16
      Top             =   4395
      Width           =   4935
      Begin VB.OptionButton Opt入科时间 
         Caption         =   "系统时间"
         Height          =   180
         Index           =   1
         Left            =   3210
         TabIndex        =   19
         Top             =   330
         Width           =   1215
      End
      Begin VB.OptionButton Opt入科时间 
         Caption         =   "入院时间"
         Height          =   180
         Index           =   0
         Left            =   1965
         TabIndex        =   18
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Lbl缺省入科时间 
         Caption         =   "缺省入科时间"
         Height          =   210
         Left            =   795
         TabIndex        =   17
         Top             =   315
         Width           =   1110
      End
   End
   Begin VB.TextBox txtOutTime 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   690
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "30"
      Top             =   2310
      Width           =   525
   End
   Begin VB.TextBox txtInTime 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   690
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "3"
      Top             =   1965
      Width           =   525
   End
   Begin VB.Frame fra待入科过滤 
      Caption         =   "显示以下科室的待入住病人"
      Height          =   1875
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   4935
      Begin VB.ListBox lstDepartments 
         Height          =   1530
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         ToolTipText     =   "Ctrl+A全选,Ctrl+C全清"
         Top             =   240
         Width           =   4665
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3900
      TabIndex        =   1
      Top             =   6705
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2730
      TabIndex        =   0
      Top             =   6705
      Width           =   1100
   End
   Begin MSComCtl2.UpDown UDInTime 
      Height          =   300
      Left            =   1215
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1965
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   30
      BuddyControl    =   "txtInTime"
      BuddyDispid     =   196623
      OrigLeft        =   2340
      OrigTop         =   210
      OrigRight       =   2580
      OrigBottom      =   450
      Max             =   365
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UDOutTime 
      Height          =   300
      Left            =   1215
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   30
      BuddyControl    =   "txtOutTime"
      BuddyDispid     =   196622
      OrigLeft        =   1215
      OrigTop         =   2310
      OrigRight       =   1455
      OrigBottom      =   2625
      Max             =   365
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "显示在          天以内的出院病人"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   2370
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "显示在          天以内登记入院的病人"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   2025
      Width           =   3240
   End
End
Attribute VB_Name = "frmSetCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String    '权限串
Public mlngModul As Long      '模块号

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strpar As String
    Dim blnSelAll As Boolean


    If txtInTime.Enabled Then
        If Trim(txtInTime.Text) = "" Then
            MsgBox "必须输入要显示的入院时间范围！", vbInformation, gstrSysName
            txtOutTime.SetFocus: Exit Sub
        Else
            zlDatabase.SetPara "入院天数", Val(txtInTime.Text), glngSys, mlngModul, IIf(txtInTime.Enabled = True, True, False)
        End If
    End If
    
    If txtOutTime.Enabled Then
        If Trim(txtOutTime.Text) = "" Then
            MsgBox "必须输入要显示的出院时间范围！", vbInformation, gstrSysName
            txtOutTime.SetFocus: Exit Sub
        Else
            zlDatabase.SetPara "出院天数", Val(txtOutTime.Text), glngSys, mlngModul, IIf(txtOutTime.Enabled = True, True, False)
        End If
    End If

    If Fra入科时间.Enabled Then
        If Opt入科时间(0).Value Then
            Call zlDatabase.SetPara("缺省入科时间", 0, glngSys, mlngModul, IIf(Fra入科时间.Enabled = True, True, False))
        Else
            Call zlDatabase.SetPara("缺省入科时间", 1, glngSys, mlngModul, IIf(Fra入科时间.Enabled = True, True, False))
        End If
    End If
    
    If fra待入科过滤.Enabled Then
        For i = lstDepartments.ListCount - 1 To 0 Step -1
            If lstDepartments.Selected(i) Then
                strpar = strpar & lstDepartments.ItemData(i) & ","
            End If
        Next
        If strpar <> "" Then
            strpar = Left(strpar, Len(strpar) - 1)
            If lstDepartments.ListCount = UBound(Split(strpar, ",")) + 1 Then strpar = "" '全选等于无科室条件
        End If
        zlDatabase.SetPara "待入科病人科室", strpar, glngSys, mlngModul, IIf(fra待入科过滤.Enabled = True, True, False)
    End If
    '问题28138 by lesfeng 2010-03-01
    zlDatabase.SetPara "默认诊断", chkAllowOut.Value, glngSys, mlngModul, IIf(chkAllowOut.Enabled = True, True, False)
    '问题28432 by lesfeng 2010-03-10
    zlDatabase.SetPara "允许调整科室", chkIn.Value, glngSys, mlngModul, IIf(chkIn.Enabled = True, True, False)
    
    zlDatabase.SetPara "护理等级默认为空", chkChangeIn.Value, glngSys, mlngModul, IIf(chkChangeIn.Enabled = True, True, False)
    '问题42701 by ljf
    zlDatabase.SetPara "允许修改入院时间", chkInTime, glngSys, mlngModul, IIf(chkInTime.Enabled = True, True, False)
    
    
    
    '婴儿腕带打印方式
    For i = 0 To optBabyWristletPrint.UBound
        If optBabyWristletPrint(i).Value Then
            zlDatabase.SetPara "婴儿腕带打印", i, glngSys, mlngModul, IIf(optBabyWristletPrint(i).Enabled = True, True, False)
        End If
    Next
    
    '49854:刘鹏飞,2013-10-31,添加病人腕带
    '病人腕带打印方式
    For i = 0 To optWristletPrint.UBound
        If optWristletPrint(i).Value Then
            zlDatabase.SetPara "病人腕带打印", i, glngSys, mlngModul, IIf(optWristletPrint(i).Enabled = True, True, False)
        End If
    Next
    
    '63706:刘鹏飞,2014-08-11,出院死亡
    zlDatabase.SetPara "出院死亡", chkDeath.Value, glngSys, mlngModul, IIf(chkDeath.Enabled = True, True, False)
    
    '72443:刘鹏飞,2014-08-11,入住时必须指定医疗小组
    zlDatabase.SetPara "入住指定医疗小组", chkMedicalTeam.Value, glngSys, mlngModul, IIf(chkMedicalTeam.Enabled = True, True, False)
    
    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(Index As Integer)
    '49854:刘鹏飞,2013-10-31,添加病人腕带
    Select Case Index
        Case 0 '病人腕带
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me)
        Case 1 '婴儿腕带
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = 13 Then cmdOK_Click
    If Shift = vbCtrlMask And fra待入科过滤.Enabled Then
        If KeyCode = vbKeyA Then
            For i = 0 To lstDepartments.ListCount - 1
                lstDepartments.Selected(i) = True
            Next
        ElseIf KeyCode = vbKeyC Then
            For i = 0 To lstDepartments.ListCount - 1
                lstDepartments.Selected(i) = False
            Next
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Long, strpar As String
    Dim rsTmp As ADODB.Recordset
    
    gblnOK = False
    
    '待入科病人科室
    Set rsTmp = GetDepts("临床", "1,2,3")
    Do While Not rsTmp.EOF
        lstDepartments.AddItem rsTmp!编码 & "-" & rsTmp!名称
        lstDepartments.ItemData(lstDepartments.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    strpar = zlDatabase.GetPara("待入科病人科室", glngSys, mlngModul, "", Array(fra待入科过滤), InStr(mstrPrivs, "参数设置") > 0)
    If strpar = "" Then
        For i = 0 To lstDepartments.ListCount - 1
            lstDepartments.Selected(i) = True
        Next
    Else
        For i = 0 To lstDepartments.ListCount - 1
            If InStr("," & strpar & ",", "," & lstDepartments.ItemData(i) & ",") > 0 Then lstDepartments.Selected(i) = True
        Next
    End If
    If lstDepartments.ListCount > 0 Then lstDepartments.TopIndex = 0: lstDepartments.ListIndex = 0
    
    txtInTime.Text = Val(zlDatabase.GetPara("入院天数", glngSys, mlngModul, "3", Array(txtInTime), InStr(mstrPrivs, "参数设置") > 0))
    txtOutTime.Text = Val(zlDatabase.GetPara("出院天数", glngSys, mlngModul, "30", Array(txtOutTime), InStr(mstrPrivs, "参数设置") > 0))
    
    Opt入科时间(zlDatabase.GetPara("缺省入科时间", glngSys, mlngModul, "0", Array(Fra入科时间), InStr(mstrPrivs, "参数设置") > 0)).Value = True
    '问题28138 by lesfeng 2010-03-01
    chkAllowOut.Value = IIf(zlDatabase.GetPara("默认诊断", glngSys, mlngModul, , Array(chkAllowOut), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    '63706:刘鹏飞,2014-08-11,出院死亡
    chkDeath.Value = IIf(zlDatabase.GetPara("出院死亡", glngSys, mlngModul, , Array(chkDeath), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    '问题28432 by lesfeng 2010-03-10
    chkIn.Value = IIf(zlDatabase.GetPara("允许调整科室", glngSys, mlngModul, , Array(chkIn), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    chkChangeIn.Value = IIf(zlDatabase.GetPara("护理等级默认为空", glngSys, mlngModul, , Array(chkChangeIn), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    chkInTime.Value = IIf(zlDatabase.GetPara("允许修改入院时间", glngSys, mlngModul, , Array(chkInTime), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    '72443:刘鹏飞,2014-08-02,入住指定医疗小组
    chkMedicalTeam.Value = IIf(zlDatabase.GetPara("入住指定医疗小组", glngSys, mlngModul, , Array(chkMedicalTeam), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    
    i = Val(zlDatabase.GetPara("婴儿腕带打印", glngSys, mlngModul, , Array(fraBabyWristlet), InStr(mstrPrivs, "参数设置") > 0))
    If i <= optBabyWristletPrint.UBound Then optBabyWristletPrint(i).Value = True
    
    '49854:刘鹏飞,2013-10-31,添加病人腕带
    i = Val(zlDatabase.GetPara("病人腕带打印", glngSys, mlngModul, , Array(fraWristlet), InStr(mstrPrivs, "参数设置") > 0))
    If i <= optWristletPrint.UBound Then optWristletPrint(i).Value = True
End Sub

Private Sub txtInTime_GotFocus()
    SelAll txtInTime
End Sub

Private Sub txtInTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOutTime_GotFocus()
    SelAll txtOutTime
End Sub

Private Sub txtOutTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
