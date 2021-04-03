VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSetReplenishTheBalance 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "退费票据打印设置(&2)"
      Height          =   350
      Index           =   2
      Left            =   5790
      TabIndex        =   37
      Top             =   5160
      Width           =   1950
   End
   Begin VB.PictureBox picDelBillFormat 
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   2130
      ScaleHeight     =   1380
      ScaleWidth      =   6015
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2040
      Width           =   6015
      Begin VSFlex8Ctl.VSFlexGrid vsDelBillFormat 
         Height          =   1350
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   5865
         _cx             =   10345
         _cy             =   2381
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSetReplenishTheBalance.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picBillFormat 
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   2910
      ScaleHeight     =   1605
      ScaleWidth      =   5925
      TabIndex        =   32
      Top             =   1710
      Width           =   5925
      Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
         Height          =   1365
         Left            =   30
         TabIndex        =   33
         Top             =   30
         Width           =   5715
         _cx             =   10081
         _cy             =   2408
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSetReplenishTheBalance.frx":0096
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.TextBox txtVaildDays 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "3"
      Top             =   4650
      Width           =   450
   End
   Begin VB.Frame fra结算方式 
      Caption         =   "补结算类别设置"
      Height          =   3165
      Left            =   6180
      TabIndex        =   2
      Top             =   120
      Width           =   1605
      Begin VB.ListBox lst结算方式 
         Height          =   2790
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.OptionButton optDrug 
      Caption         =   "提醒"
      Height          =   180
      Index           =   2
      Left            =   3525
      TabIndex        =   27
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optDrug 
      Caption         =   "禁止"
      Height          =   180
      Index           =   1
      Left            =   2715
      TabIndex        =   26
      Top             =   5820
      Width           =   735
   End
   Begin VB.OptionButton optDrug 
      Caption         =   "不检查"
      Height          =   180
      Index           =   0
      Left            =   1815
      TabIndex        =   25
      Top             =   5820
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox chkPayKey 
      Caption         =   "使用小键盘的加减(+-)来切换支付方式"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5310
      Width           =   3375
   End
   Begin VB.TextBox txt票据张数 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "10"
      Top             =   4950
      Width           =   465
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   5790
      TabIndex        =   16
      Top             =   4260
      Width           =   1950
   End
   Begin VB.Frame fraTitle 
      Caption         =   "本地共用收费票据"
      Height          =   1410
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5925
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   5670
         _cx             =   10001
         _cy             =   1931
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSetReplenishTheBalance.frx":012C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "结算清单打印设置(&3)"
      Height          =   350
      Index           =   1
      Left            =   5790
      TabIndex        =   12
      Top             =   5610
      Width           =   1950
   End
   Begin VB.Frame fraFeeList 
      Caption         =   "结算清单打印方式"
      Height          =   675
      Left            =   1980
      TabIndex        =   7
      Top             =   3390
      Width           =   5760
      Begin VB.OptionButton optPrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   585
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3690
         TabIndex        =   10
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6675
      Left            =   7860
      TabIndex        =   31
      Top             =   -600
      Width           =   45
   End
   Begin VB.Frame fra单位 
      Caption         =   " 药品显示单位 "
      Height          =   1155
      Left            =   135
      TabIndex        =   4
      Top             =   3390
      Width           =   1635
      Begin VB.OptionButton opt单位 
         Caption         =   "售价单位"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton opt单位 
         Caption         =   "门诊单位"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   780
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "收费票据打印设置(&1)"
      Height          =   350
      Index           =   0
      Left            =   5790
      TabIndex        =   11
      Top             =   4710
      Width           =   1950
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8100
      TabIndex        =   28
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8100
      TabIndex        =   29
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   8100
      TabIndex        =   30
      Top             =   5310
      Width           =   1100
   End
   Begin MSComCtl2.UpDown upd票据张数 
      Height          =   300
      Left            =   1635
      TabIndex        =   19
      Top             =   4950
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "txt票据张数"
      BuddyDispid     =   196617
      OrigLeft        =   1605
      OrigTop         =   4860
      OrigRight       =   1860
      OrigBottom      =   5160
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox chk票据张数 
      Caption         =   "票据剩余         张时开始提醒收费员"
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   4965
      Width           =   3450
   End
   Begin VB.TextBox txtNameDays 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   22
      Text            =   "0"
      ToolTipText     =   "0表示查找时不限制时间"
      Top             =   5580
      Width           =   285
   End
   Begin VB.Frame fraShortLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   2730
      TabIndex        =   23
      Top             =   5760
      Width           =   285
   End
   Begin VB.CheckBox chkSeekName 
      Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   5565
      Width           =   4260
   End
   Begin MSComCtl2.UpDown updVaildDays 
      Height          =   270
      Left            =   3480
      TabIndex        =   15
      Top             =   4650
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "txtVaildDays"
      BuddyDispid     =   196612
      OrigLeft        =   1605
      OrigTop         =   4860
      OrigRight       =   1860
      OrigBottom      =   5160
      Max             =   100
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl tbBillSet 
      Height          =   1800
      Left            =   150
      TabIndex        =   36
      Top             =   1560
      Width           =   5925
      _Version        =   589884
      _ExtentX        =   10451
      _ExtentY        =   3175
      _StockProps     =   64
   End
   Begin VB.Label lblVaildDays 
      Caption         =   "可进行保险补充结算的费用有效天数"
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label lblDrugNotFee 
      AutoSize        =   -1  'True
      Caption         =   "药品摆药后退费方式"
      Height          =   180
      Left            =   150
      TabIndex        =   24
      Top             =   5820
      Width           =   1620
   End
End
Attribute VB_Name = "frmSetReplenishTheBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String
Private mlngModule As Long
Private mblnOK As Boolean

Public Function zlSetPara(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '返回:参数设置成功,返回true,否则的返回False
    '编制:李南春
    '日期:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOK = False
    
    Me.Show 1, frmMain
    zlSetPara = mblnOK
End Function

Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1 And txtNameDays.Tag = "1"
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, mlngModule)
End Sub

Private Sub cmdOK_Click()
    Dim blnHavePrivs As Boolean, i As Integer
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    If isValied = False Then Exit Sub
    zlDatabase.SetPara "药品单位显示", IIf(opt单位(0).Value, 0, 1), glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "姓名模糊查找方式", IIf(chkSeekName.Value = 1, "1", "0") & "|" & Val(txtNameDays.Text), glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "票据剩余X张时开始提醒收费员", IIf(chk票据张数.Value = 1, "1", "0") & "|" & Val(txt票据张数.Text), glngSys, mlngModule, blnHavePrivs
    For i = 0 To optPrint.UBound
        If optPrint(i).Value Then
            zlDatabase.SetPara "结算清单打印方式", i, glngSys, mlngModule, blnHavePrivs
        End If
    Next
    Call SaveInvoice
    '47457,82343
    zlDatabase.SetPara "使用加减切换支付方式", IIf(chkPayKey.Value = 1, "1", "0"), glngSys, mlngModule, blnHavePrivs
    '47400,82343
    zlDatabase.SetPara "药品摆药退费方式", IIf(optDrug(0).Value, 0, IIf(optDrug(1).Value, "1", "2")), glngSys, mlngModule, blnHavePrivs
    '84929
    zlDatabase.SetPara "补结算有效天数", Val(txtVaildDays.Text), glngSys, mlngModule, blnHavePrivs
    mblnOK = True
    Unload Me
End Sub

Private Sub chk票据张数_Click()
    txt票据张数.Enabled = chk票据张数.Enabled And chk票据张数.Value = 1
    upd票据张数.Enabled = txt票据张数.Enabled
End Sub

Private Sub cmdPrintSetup_Click(Index As Integer)
    Select Case Index
        Case 0 '门诊医疗费收费
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1124", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124", Me)
            End If
        Case 1 '门诊收费清单
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1124_1", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me)
            End If
        Case 2 '退费发票(红票)
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1124_3", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_3", Me)
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim strTmp As String, blnParSet As Boolean, i As Integer
    
    On Error GoTo errH
    blnParSet = InStr(1, mstrPrivs, "参数设置") > 0
    Call InitTabControl
    
    i = IIf(zlDatabase.GetPara("药品单位显示", glngSys, mlngModule, , Array(opt单位(0), opt单位(1)), blnParSet) = "0", 0, 1)
    opt单位(i).Value = True
    txtNameDays.Enabled = True
    strTmp = zlDatabase.GetPara("姓名模糊查找方式", glngSys, mlngModule, "0|10", Array(txtNameDays, chkSeekName), blnParSet)
    txtNameDays.Text = Val(Split(strTmp & "|", "|")(1))
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
    
    strTmp = zlDatabase.GetPara("票据剩余X张时开始提醒收费员", glngSys, mlngModule, "0|10", Array(txt票据张数, upd票据张数, chk票据张数), blnParSet)
    
    upd票据张数.Value = Val(Split(strTmp & "|", "|")(1))
    txt票据张数.Text = upd票据张数.Value
    chk票据张数.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
    txt票据张数.Enabled = chk票据张数.Enabled And chk票据张数.Value = 1
    upd票据张数.Enabled = txt票据张数.Enabled
    
    i = Val(zlDatabase.GetPara("结算清单打印方式", glngSys, mlngModule, , Array(optPrint(0), optPrint(1), optPrint(2)), blnParSet))
    If i <= optPrint.UBound Then optPrint(i).Value = True
    Call InitShareInvoice
    '47457,82343
    chkPayKey.Value = IIf(Val(zlDatabase.GetPara("使用加减切换支付方式", glngSys, mlngModule, "1", Array(chkPayKey), blnParSet)) = 1, 1, 0)
    '47400,82343
    strTmp = zlDatabase.GetPara("药品摆药退费方式", glngSys, mlngModule, , Array(optDrug(0), optDrug(1), optDrug(2)), blnParSet)
    For i = 0 To 2
        If Val(strTmp) = i Then
            optDrug(i).Value = True: Exit For
        End If
    Next
    '84929
    txtVaildDays.Text = Val(zlDatabase.GetPara("补结算有效天数", glngSys, mlngModule, "3", Array(lblVaildDays, txtVaildDays, updVaildDays), blnParSet))
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '日期:2011-04-28 15:09:10
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModule, vsBill, Me.Name, "共享票据批次列", False, False
    zl_vsGrid_Para_Restore mlngModule, vsBillFormat, Me.Name, "收费发票打印方式", False, False
    zl_vsGrid_Para_Restore mlngModule, vsDelBillFormat, Me.Name, "退费发票打印方式", False, False
    strShareInvoice = zlDatabase.GetPara("共用收费票据批次", glngSys, mlngModule, , , True, intType)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
     '格式:领用ID1,使用类别1|领用IDn,使用类别n|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(1)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!使用类别, " ")
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    '票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号='ZL" & glngSys \ 100 & "_BILL_1124'  " & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("收费票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("收费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    
    '读取参数值
    strShareInvoice = zlDatabase.GetPara("收费发票格式", glngSys, mlngModule, , , True, intType)
    strPrintMode = zlDatabase.GetPara("收费发票打印方式", glngSys, mlngModule, , , True, intType1)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat
         .ColData(.ColIndex("收费票据格式")) = "0"
         .ColData(.ColIndex("收费打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("收费票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("收费打印方式")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("收费票据格式"))) = 1 Or _
            Val(.ColData(.ColIndex("收费打印方式"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    
    '80943,冉俊明,2014-12-18,票据未使用“收费类别”时，加入设置收费类别为空的打印方式和票据格式
    Dim objInvoice As New zlPublicExpense.clsInvoice
    If objInvoice.zlStartFactUseType(EM_收费收据) Then
        strSQL = "" & _
            "   Select 编码, 名称 From 票据使用类别" & _
            "   Order By 编码"
    Else
        strSQL = "" & _
            "   Select 编码, 名称 From 票据使用类别" & _
            "   Union All" & _
            "   Select '', '' From Dual " & _
            "   Order By 编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("收费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("收费票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("收费票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("收费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("收费打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("收费打印方式"), .Rows - 1, .ColIndex("收费打印方式")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("收费票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("收费票据格式"), .Rows - 1, .ColIndex("收费票据格式")) = vbBlue
        End If
    End With
    
    '====================================================================
    '退费票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号='ZL" & glngSys \ 100 & "_BILL_1124_3'  " & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .ColComboList(.ColIndex("退费票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("退费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    
    '读取参数值
    strShareInvoice = zlDatabase.GetPara("退费发票格式", glngSys, mlngModule, , , True, intType)
    strPrintMode = zlDatabase.GetPara("退费发票打印方式", glngSys, mlngModule, , , True, intType1)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsDelBillFormat
         .ColData(.ColIndex("退费票据格式")) = "0"
         .ColData(.ColIndex("退费打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("退费票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("退费打印方式")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("退费票据格式"))) = 1 Or _
            Val(.ColData(.ColIndex("退费打印方式"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    
    '80943,冉俊明,2014-12-18,票据未使用“收费类别”时，加入设置收费类别为空的打印方式和票据格式
    If objInvoice.zlStartFactUseType(EM_收费收据) Then
        strSQL = "" & _
            "   Select 编码, 名称 From 票据使用类别" & _
            "   Order By 编码"
    Else
        strSQL = "" & _
            "   Select 编码, 名称 From 票据使用类别" & _
            "   Union All" & _
            "   Select '', '' From Dual " & _
            "   Order By 编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("退费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("退费票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("退费票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("退费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("退费打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("退费打印方式"), .Rows - 1, .ColIndex("退费打印方式")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("退费票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("退费票据格式"), .Rows - 1, .ColIndex("退费打印方式")) = vbBlue
        End If
    End With
    
    '82990:李南春,2015/3/9,补结算类别设置
    Call Load结算方式(blnHavePrivs)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共享票据批次列", False, False
    zl_vsGrid_Para_Save mlngModule, vsBillFormat, Me.Name, "收费发票打印方式", False, False
    zl_vsGrid_Para_Save mlngModule, vsDelBillFormat, Me.Name, "退费发票打印方式", False, False
End Sub

Private Sub fra票据格式_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共享票据批次列", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共享票据批次列", False, False
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("选择")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存发票相关票据
    '编制:刘兴洪
    '日期:2011-04-28 18:16:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    '保存共享票据
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.TextMatrix(i, .ColIndex("使用类别")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用收费票据批次", strValue, glngSys, mlngModule, blnHavePrivs
    '保存收费格式
    
    Dim strPrintMode As String
    '保存收费格式
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("收费票据格式")))
            strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("收费打印方式")), 1))
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "收费发票格式", strValue, glngSys, mlngModule, blnHavePrivs
        zlDatabase.SetPara "收费发票打印方式", strPrintMode, glngSys, mlngModule, blnHavePrivs
    End With
    
    '====================================================
    '保存退费格式
    strValue = "": strPrintMode = ""
    With vsDelBillFormat
        For i = 1 To .Rows - 1
            strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("退费票据格式")))
            strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("退费打印方式")), 1))
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "退费发票格式", strValue, glngSys, mlngModule, blnHavePrivs
        zlDatabase.SetPara "退费发票打印方式", strPrintMode, glngSys, mlngModule, blnHavePrivs
    End With
    
    '82990:李南春,2015/3/9,补结算类别设置
    strValue = ""
    For i = 0 To lst结算方式.ListCount - 1
        If lst结算方式.Selected(i) = True Then
            strValue = strValue & "|" & lst结算方式.List(i)
        End If
    Next
    strValue = Mid(strValue, 2)
    zlDatabase.SetPara "补结算类别设置", strValue, glngSys, mlngModule, blnHavePrivs
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-28 18:24:16
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
     
    isValied = False
    On Error GoTo errHandle
    '检查每种使用种式只能一个选择
    With vsBill
        str类别 = "-"
        For i = 1 To vsBill.Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("使用类别"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("使用类别")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("使用类别"))) = Trim(.TextMatrix(j, .ColIndex("使用类别"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    使用类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    isValied = True
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtNameDays_GotFocus()
    Call SelAll(txtNameDays)
End Sub

Private Sub txtNameDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNameDays_Validate(Cancel As Boolean)
    If Val(txtNameDays.Text) <= 0 Then
        txtNameDays.Text = 0
    ElseIf Val(txtNameDays.Text) > 999 Then
        txtNameDays.Text = 999
    End If
End Sub

Private Sub Load结算方式(ByVal blnHavePrivs As Boolean)
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objItem As Object
    Dim str支付方式 As String
    
    On Error GoTo errHandle
    '85182:李南春,2015/5/27,参数设置权限控制
    str支付方式 = zlDatabase.GetPara("补结算类别设置", glngSys, mlngModule, , Array(fra结算方式, lst结算方式), blnHavePrivs)
    strSQL = "Select distinct B.编码,B.名称 From 结算方式应用 A,结算方式 B" & vbNewLine & _
            "Where A.应用场合 in ('挂号','收费') And B.名称=A.结算方式" & vbNewLine & _
            "And   (B.性质<>3 And B.性质<>4)" & vbNewLine & _
            "Order by lpad(编码,3,' ')"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "保险补充结算")
    
    lst结算方式.Clear
    Do Until rsTemp.EOF
        lst结算方式.AddItem Nvl(rsTemp!名称)
        If InStr("|" & str支付方式 & "|", "|" & Nvl(rsTemp!名称) & "|") > 0 Then lst结算方式.Selected(lst结算方式.NewIndex) = True
        rsTemp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitTabControl()
    With tbBillSet
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionTop
'        .PaintManager.StaticFrame = True
'        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .InsertItem 0, "收费票据格式", picBillFormat.hWnd, 0
        .InsertItem 1, "退费票据格式", picDelBillFormat.hWnd, 0
        .Item(0).Selected = True
    End With
End Sub

