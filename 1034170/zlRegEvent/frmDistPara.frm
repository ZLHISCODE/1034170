VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDistPara 
   Caption         =   "分诊科室设置"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   ControlBox      =   0   'False
   Icon            =   "frmDistPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra条码 
      Caption         =   "条码打印"
      Height          =   735
      Left            =   135
      TabIndex        =   30
      Top             =   7260
      Width           =   6525
      Begin VB.CommandButton cmdBarcodeSet 
         Caption         =   "条码打印设置"
         Height          =   375
         Left            =   4680
         TabIndex        =   34
         Top             =   240
         Width           =   1620
      End
      Begin VB.OptionButton optBarcode 
         Caption         =   "提示选择打印"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   33
         Top             =   360
         Width           =   1770
      End
      Begin VB.OptionButton optBarcode 
         Caption         =   "自动打印"
         Height          =   195
         Index           =   1
         Left            =   1530
         TabIndex        =   32
         Top             =   360
         Width           =   1170
      End
      Begin VB.OptionButton optBarcode 
         Caption         =   "不打印"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   990
      End
   End
   Begin VB.CheckBox chkBusy 
      Caption         =   "医生诊室忙时允许分诊"
      Height          =   300
      Left            =   165
      TabIndex        =   29
      Top             =   9460
      Width           =   4620
   End
   Begin VB.TextBox txt提前分诊 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   585
      TabIndex        =   27
      Text            =   "0"
      Top             =   9115
      Width           =   375
   End
   Begin VB.Frame fra排序 
      Caption         =   "候诊病人排序方式"
      Height          =   1020
      Left            =   135
      TabIndex        =   23
      Top             =   8070
      Width           =   6540
      Begin VB.OptionButton optSort 
         Caption         =   "科室编码,号码,发生时间,登记时间"
         Height          =   210
         Index           =   2
         Left            =   390
         TabIndex        =   26
         Top             =   675
         Width           =   3555
      End
      Begin VB.OptionButton optSort 
         Caption         =   "科室编码,号码,挂号时间"
         Height          =   210
         Index           =   1
         Left            =   2775
         TabIndex        =   25
         Top             =   360
         Width           =   2280
      End
      Begin VB.OptionButton optSort 
         Caption         =   "科室编码,号码,单据号"
         Height          =   210
         Index           =   0
         Left            =   390
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   2280
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7995
      Left            =   7050
      TabIndex        =   21
      Top             =   -120
      Width           =   45
   End
   Begin VB.Frame fra排队单 
      Caption         =   "排队单打印"
      Height          =   735
      Left            =   135
      TabIndex        =   17
      Top             =   6420
      Width           =   6525
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "排队单打印设置"
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   270
         Width           =   1620
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "提示选择打印"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   20
         Top             =   375
         Width           =   1455
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "自动打印"
         Height          =   195
         Index           =   1
         Left            =   1530
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "不打印"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   990
      End
   End
   Begin VB.Frame fra排队叫号 
      Caption         =   "排队叫号设置"
      Height          =   1875
      Left            =   150
      TabIndex        =   7
      Top             =   4440
      Width           =   6540
      Begin VB.CheckBox chk预约排队 
         Caption         =   "预约挂号进入队列"
         Height          =   270
         Left            =   4200
         TabIndex        =   16
         Top             =   1410
         Width           =   1905
      End
      Begin VB.CheckBox chk签到排队 
         Caption         =   "分诊台签到开始排队"
         Height          =   330
         Left            =   2130
         TabIndex        =   15
         Top             =   1380
         Width           =   1935
      End
      Begin VB.CheckBox chk分诊呼叫 
         Caption         =   "分诊后立即呼叫"
         Height          =   300
         Left            =   180
         TabIndex        =   14
         Top             =   1395
         Width           =   2340
      End
      Begin VB.OptionButton opt排队模式 
         Caption         =   "先分诊呼叫,再医生呼叫就诊"
         Height          =   240
         Index           =   2
         Left            =   3870
         TabIndex        =   10
         Top             =   405
         Width           =   2625
      End
      Begin VB.OptionButton opt排队模式 
         Caption         =   "分诊台分诊呼叫或医生主动呼叫"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   720
         Width           =   3045
      End
      Begin VB.OptionButton opt排队模式 
         Caption         =   "禁止全院排队叫号"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   390
         Width           =   1770
      End
      Begin VB.Frame fra呼叫对象 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         TabIndex        =   11
         Top             =   720
         Width           =   6330
         Begin VB.OptionButton opt呼叫对象 
            Caption         =   "医生主动呼叫"
            Height          =   240
            Index           =   1
            Left            =   2070
            TabIndex        =   13
            Top             =   315
            Width           =   1725
         End
         Begin VB.OptionButton opt呼叫对象 
            Caption         =   "分诊台分诊呼叫"
            Height          =   240
            Index           =   0
            Left            =   330
            TabIndex        =   12
            Top             =   315
            Width           =   1725
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7320
      TabIndex        =   2
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   1
      Top             =   270
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   5790
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistPara.frx":058A
            Key             =   "bm"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3360
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Ctrl+A全选,Ctrl+C全消"
      Top             =   540
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5927
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imglst"
      SmallIcons      =   "imglst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   7937
      EndProperty
   End
   Begin MSComCtl2.UpDown upd天数 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4005
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtUD(1)"
      BuddyDispid     =   196632
      BuddyIndex      =   1
      OrigLeft        =   2625
      OrigTop         =   3990
      OrigRight       =   2865
      OrigBottom      =   4290
      Max             =   7
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtUD 
      Alignment       =   1  'Right Justify
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1005
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "0"
      Top             =   4005
      Width           =   675
   End
   Begin VB.Label lbl提前分诊 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提前     小时分诊,设置为零,表示在当前系统时间内挂号病人进行分诊"
      Height          =   180
      Left            =   165
      TabIndex        =   28
      Top             =   9145
      Width           =   5670
   End
   Begin VB.Label lbl有效天数 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "自动刷新            天内的挂号病人,设置为零,表示只刷新当前的挂号病人"
      Height          =   180
      Left            =   225
      TabIndex        =   6
      Top             =   4065
      Width           =   6120
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   285
      Picture         =   "frmDistPara.frx":0B24
      Top             =   30
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    一个分诊台可以同时管理多个门诊临床科室挂号病人，进行分诊相关处理；请选择由本分诊台进行分诊的临床科室(Ctrl+A全选,Ctrl+C全消)"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   930
      TabIndex        =   3
      Top             =   90
      Width           =   5805
   End
End
Attribute VB_Name = "frmDistPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public mlngModul As Long
Private mblnNotClick As Boolean
 

Private Sub cmdBarcodeSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & Int(glngSys \ 100) & "_BILL_1113_1", Me)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ObjItem As ListItem, strTmp As String
    
    For Each ObjItem In Me.lvwMain.ListItems
        If ObjItem.Checked Then
            strTmp = strTmp & "," & Mid(ObjItem.Key, 2)
        End If
    Next
    If strTmp = "" Then
        If MsgBox("你没有设置对任何科室分诊，该分诊台将不能进行分诊操作。" & vbCrLf & "真的暂时不设置吗？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
        strTmp = "0"
    Else
        strTmp = Mid(strTmp, 2)
        If UBound(Split(strTmp, ",")) + 1 = lvwMain.ListItems.Count Then strTmp = ""
    End If
    zlDatabase.SetPara "分诊科室", strTmp, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0 '空表示全部科室
    zlDatabase.SetPara "分诊有效天数", Val(txtUD(1).Text), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0  '空表示全部科室
    
    '1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.3-不排队叫号
    zlDatabase.SetPara "排队叫号模式", IIf(opt排队模式(0).Value, 0, IIf(opt排队模式(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0 '空表示全部科室
    If opt排队模式(1).Value Then
        zlDatabase.SetPara "排队呼叫站点", IIf(opt呼叫对象(0).Value, 0, 1), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0 '空表示全部科室
    End If
    zlDatabase.SetPara "分诊后立即呼叫", IIf(chk分诊呼叫.Enabled = False, 0, chk分诊呼叫.Value), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0 '空表示全部科室
    zlDatabase.SetPara "分诊台签到排队", chk签到排队.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    '问题:44621
    zlDatabase.SetPara "预约生成队列", chk预约排队.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0

    '38165
    zlDatabase.SetPara "排队单打印", IIf(optPrint(0).Value, 0, IIf(optPrint(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "候诊排序方式", IIf(optSort(0).Value, 0, IIf(optSort(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    '77412:李南春，2014/9/3,门诊病人条码打印
    zlDatabase.SetPara "条码打印方式", IIf(optBarcode(0).Value, 0, IIf(optBarcode(1).Value, 1, 2)), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    '51223
    zlDatabase.SetPara "提前N小时分诊", Val(txt提前分诊.Text), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    zlDatabase.SetPara "诊室忙时允许分诊", chkBusy.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0

    Unload Me
End Sub

 
Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & Int(glngSys \ 100) & "_BILL_1113", Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Dim i As Integer
        If UCase(Chr(KeyCode)) = "A" Then
            For i = 1 To lvwMain.ListItems.Count
                lvwMain.ListItems(i).Checked = True
            Next
        ElseIf UCase(Chr(KeyCode)) = "C" Then
            For i = 1 To lvwMain.ListItems.Count
                lvwMain.ListItems(i).Checked = False
            Next
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim ObjItem As ListItem
    Dim blnEnabled As Boolean
    
    Call RestoreWinState(Me, App.ProductName)
    mblnNotClick = True
    Select Case Val(zlDatabase.GetPara("排队叫号模式", glngSys, mlngModul, , Array(opt排队模式(0), opt排队模式(1), opt排队模式(2)), InStr(1, mstrPrivs, ";参数设置;") > 0))
    Case 1 '1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
        opt排队模式(1).Value = True
        blnEnabled = True
    Case 2
        opt排队模式(2).Value = True
        blnEnabled = False
    Case Else
        opt排队模式(0).Value = True
        blnEnabled = False
    End Select
    
    Select Case Val(zlDatabase.GetPara("排队呼叫站点", glngSys, mlngModul, , Array(opt呼叫对象(0), opt呼叫对象(1)), InStr(1, mstrPrivs, ";参数设置;") > 0))
    Case 0  '0-代表分诊台分诊呼叫;1-代表医生主动呼叫
        opt呼叫对象(0).Value = True
    Case Else
        opt呼叫对象(1).Value = True
    End Select
    opt排队模式(1).Tag = IIf(opt呼叫对象(1).Enabled, 1, 0)
    opt呼叫对象(1).Enabled = opt呼叫对象(1).Enabled And blnEnabled
    opt呼叫对象(0).Enabled = opt呼叫对象(0).Enabled And blnEnabled
    
    chk分诊呼叫.Value = IIf(Val(zlDatabase.GetPara("分诊后立即呼叫", glngSys, mlngModul, , Array(chk分诊呼叫), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1, 1, 0)
    chk签到排队.Value = IIf(Val(zlDatabase.GetPara("分诊台签到排队", glngSys, mlngModul, , Array(chk签到排队), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1, 1, 0)
    '问题:44621
    chk预约排队.Value = IIf(Val(zlDatabase.GetPara("预约生成队列", glngSys, mlngModul, , Array(chk预约排队), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1, 1, 0)
    
    chk分诊呼叫.Tag = IIf(chk分诊呼叫.Enabled, 1, 0)
    'chk分诊呼叫.Enabled = Not opt呼叫对象(1).Value And chk分诊呼叫.Enabled
    '问题:43012
    Select Case Val(zlDatabase.GetPara("候诊排序方式", glngSys, mlngModul, , Array(fra排序, optSort(0), optSort(1), optSort(2)), InStr(1, mstrPrivs, ";参数设置;") > 0))
    Case 0
        optSort(0).Value = True
        optSort(1).Value = False
        optSort(2).Value = False
    Case 1
        optSort(1).Value = True
        optSort(0).Value = False
        optSort(2).Value = False
    Case 2
        optSort(2).Value = True
        optSort(0).Value = False
        optSort(1).Value = False
    End Select
    
    '38165
    Select Case Val(zlDatabase.GetPara("排队单打印", glngSys, mlngModul, , Array(optPrint(0), optPrint(1), optPrint(2), fra排队单), InStr(1, mstrPrivs, ";参数设置;") > 0))
    Case 0
        optPrint(0).Value = True
    Case 1
        optPrint(1).Value = True
    Case Else
        optPrint(2).Value = True
    End Select
    '77412:李南春，2014/9/3,门诊病人条码打印
    Select Case Val(zlDatabase.GetPara("条码打印方式", glngSys, mlngModul, , Array(optBarcode(0), optBarcode(1), optBarcode(2), fra条码), InStr(1, mstrPrivs, ";参数设置;") > 0))
    Case 0
        optBarcode(0).Value = True
    Case 1
        optBarcode(1).Value = True
    Case Else
        optBarcode(2).Value = True
    End Select
    strTmp = zlDatabase.GetPara("分诊有效天数", glngSys, mlngModul, , Array(txtUD(1), lbl有效天数), InStr(1, mstrPrivs, ";参数设置;") > 0)
    upd天数.Value = Val(strTmp): txtUD(1).Text = Val(strTmp)
    upd天数.Enabled = txtUD(1).Enabled
    mblnNotClick = False
    
    '先得到以前设置的分诊科室ID,空表示所有诊室
    strTmp = zlDatabase.GetPara("分诊科室", glngSys, mlngModul, , Array(lvwMain), InStr(1, mstrPrivs, ";参数设置;") > 0)
    Me.lvwMain.ListItems.Clear
    On Error GoTo errH
    
    If InStr(mstrPrivs, "所有科室") > 0 Then
        Set rsTmp = GetDepartments("'临床'", "1,3")
    Else
        strSQL = _
            " Select A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And B.工作性质='临床' And B.服务对象 IN(1,3)" & _
            " And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    End If
    
    With rsTmp
        Do While Not .EOF
            Set ObjItem = Me.lvwMain.ListItems.Add(, "K" & !ID, !编码, "bm", "bm")
            ObjItem.SubItems(1) = Nvl(!名称)
            If InStr("," & strTmp & ",", "," & !ID & ",") > 0 Or strTmp = "" Then ObjItem.Checked = True
            .MoveNext
        Loop
    End With
    
    '问题号:51223
     strTmp = zlDatabase.GetPara("提前N小时分诊", glngSys, mlngModul, , Array(txt提前分诊, lbl提前分诊), InStr(1, mstrPrivs, ";参数设置;") > 0)
     If strTmp = "" Then
        txt提前分诊.Text = "0"
     Else
        txt提前分诊.Text = strTmp
     End If
     
     chkBusy.Value = Val(zlDatabase.GetPara("诊室忙时允许分诊", glngSys, mlngModul, , Array(chkBusy), InStr(1, mstrPrivs, ";参数设置;") > 0))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Long
    
    'Me.cmdCancel.Top = ScaleTop + ScaleHeight - cmdCancel.Height - 100   '   lvwMain.Top + lvwMain.Height + 120
    'Me.cmdOK.Top = Me.cmdCancel.Top
    Me.cmdCancel.Left = ScaleWidth - (Me.cmdCancel.Width + 90)
    Me.cmdOK.Left = Me.cmdCancel.Left 'Me.cmdCancel.Left - (Me.cmdOK.Width + 20)
    Me.Frame1.Left = Me.cmdOK.Left - Frame1.Width - 50
    Me.Frame1.Height = ScaleHeight + 100
    '问题号:51223
    txt提前分诊.Top = Me.ScaleHeight - txt提前分诊.Height - chkBusy.Height - 100
    lbl提前分诊.Top = txt提前分诊.Top + (txt提前分诊.Height - lbl提前分诊.Height) / 2
    chkBusy.Top = txt提前分诊.Top + txt提前分诊.Height + 50
    fra排序.Top = txt提前分诊.Top - fra排序.Height - 50
    '77412:李南春，2014/9/3,门诊病人条码打印
    Me.fra条码.Top = fra排序.Top - fra条码.Height - 50
    Me.fra排队单.Top = fra条码.Top - fra排队单.Height - 50
    
    txtUD(1).Top = fra排队单.Top - txtUD(1).Height - 50: upd天数.Top = txtUD(1).Top
    lbl有效天数.Top = txtUD(1).Top + (txtUD(1).Height - lbl有效天数.Height) \ 2
    fra排队叫号.Top = txtUD(1).Top - fra排队叫号.Height - 50
    fra排队叫号.Width = Frame1.Left - fra排队叫号.Left * 2
    fra排队单.Width = Frame1.Left - fra排队单.Left * 2
    fra条码.Width = Frame1.Left - fra条码.Left * 2
    fra排序.Width = Frame1.Left - fra排序.Left * 2
    i = Frame1.Left - (lvwMain.Left * 2)
    lvwMain.Width = IIf(i > Screen.TwipsPerPixelX, i, Screen.TwipsPerPixelX)
    lvwMain.Height = fra排队叫号.Top - 50 - lvwMain.Top 'IIf(i > Screen.TwipsPerPixelY, i, Screen.TwipsPerPixelY)
    
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain.Sorted = True
    If lvwMain.SortKey = ColumnHeader.Index - 1 Then
        If lvwMain.SortOrder = lvwAscending Then
            lvwMain.SortOrder = lvwDescending
        Else
            lvwMain.SortOrder = lvwAscending
        End If
    Else
        lvwMain.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub opt呼叫对象_Click(Index As Integer)
        If mblnNotClick Then Exit Sub
        
'        If opt呼叫对象(1).Value = False Then
'            chk分诊呼叫.Enabled = IIf(Val(chk分诊呼叫.Tag) = 1, True, False)
'        Else
'            chk分诊呼叫.Enabled = False
'        End If
'        If chk分诊呼叫.Enabled = False Then chk分诊呼叫.Value = 0
End Sub

Private Sub opt排队模式_Click(Index As Integer)
        If mblnNotClick Then Exit Sub
        If opt排队模式(1).Value Then
                opt呼叫对象(0).Enabled = IIf(Val(opt排队模式(1).Tag) = 1, True, False)
                opt呼叫对象(1).Enabled = opt呼叫对象(0).Enabled
        Else
                opt呼叫对象(0).Enabled = False
                opt呼叫对象(1).Enabled = opt呼叫对象(0).Enabled
        End If
        chk签到排队.Enabled = opt排队模式(0).Value = False
End Sub
Private Sub txt提前分诊_KeyPress(KeyAscii As Integer)
    '问题号:51223
     zlControl.TxtCheckKeyPress txt提前分诊, KeyAscii, m数字式
End Sub
