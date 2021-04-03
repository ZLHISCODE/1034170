VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBalanceBat 
   Caption         =   "批量中途结帐"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBalanceBat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11820
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   45
      ScaleHeight     =   1125
      ScaleWidth      =   11730
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5640
      Width           =   11730
      Begin VB.CommandButton cmdOK 
         Caption         =   "结帐(&O)"
         Default         =   -1  'True
         Height          =   400
         Left            =   8640
         TabIndex        =   15
         Top             =   705
         Width           =   1400
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "退出(&C)"
         Height          =   400
         Left            =   10200
         TabIndex        =   16
         Top             =   705
         Width           =   1400
      End
      Begin VB.ComboBox cbo结算方式 
         Height          =   360
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9675
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1905
      End
      Begin MSMask.MaskEdBox txtDateEnd 
         Height          =   360
         Left            =   280
         TabIndex        =   7
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDateEnd 
         Caption         =   "对                     之前的费用结帐"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   60
         Width           =   4440
      End
      Begin VB.Label lblDeposit 
         Caption         =   "冲预交合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblPay 
         Caption         =   "XX结算合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lbl结算方式 
         Caption         =   "结算方式"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   240
         Left            =   8880
         TabIndex        =   10
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Caption         =   "共完成n个病人结帐"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   480
         Width           =   8295
      End
   End
   Begin VB.Frame fra 
      Height          =   645
      Left            =   90
      TabIndex        =   17
      Top             =   0
      Width           =   11685
      Begin VB.ComboBox cbo使用类别 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   2490
      End
      Begin VB.Label lblRpt 
         AutoSize        =   -1  'True
         Caption         =   "sss"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3810
         TabIndex        =   2
         Top             =   255
         Width           =   405
      End
      Begin VB.Label lbl使用类别 
         Caption         =   "使用类别"
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   4860
      Left            =   2160
      TabIndex        =   4
      Top             =   675
      Width           =   2460
      _cx             =   4339
      _cy             =   8572
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
      BackColorSel    =   13627390
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":617A
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   4860
      Left            =   4680
      TabIndex        =   5
      Top             =   690
      Width           =   7065
      _cx             =   12462
      _cy             =   8572
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
      BackColorSel    =   12640511
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":61C2
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsFeeType 
      Height          =   4875
      Left            =   120
      TabIndex        =   3
      Top             =   675
      Width           =   1980
      _cx             =   3492
      _cy             =   8599
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
      BackColorSel    =   15790320
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":62A7
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
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
Attribute VB_Name = "frmBalanceBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPatis As String '用于记录选择的科室下标记为不结帐的病人ID
Private mlng领用ID As Long
Private mrsRptFormat As ADODB.Recordset
Private mlngShareUseID As Long     '共享批次
Private mstrUseType As String          '使用类别
Private mintPrintMode As Integer    '打印方式:0-不打印;1-提示打印;2-自动打印
Private mintPrintFormat As Integer '打印格式
  
Private Sub cbo使用类别_Click()
    lblRpt.Caption = ""
    mstrUseType = cbo使用类别.Text
    If mrsRptFormat Is Nothing Then Exit Sub
    mrsRptFormat.Filter = "序号=" & cbo使用类别.ItemData(cbo使用类别.ListIndex)
    If Not mrsRptFormat.EOF Then
        lblRpt.Caption = Nvl(mrsRptFormat!说明)
    End If
    mlng领用ID = 0
    mlngShareUseID = zl_GetInvoiceShareID(1137, mstrUseType)    '共享批次
    mintPrintMode = zl_GetInvoicePrintMode(1137, mstrUseType)  '打印方式:0-不打印;1-提示打印;2-自动打印
    mintPrintFormat = zl_GetInvoicePrintFormat(1137, mstrUseType)     '打印格式
    Call RefreshFact
    
    Call vsDept_AfterRowColChange(0, 0, vsDept.Row, vsDept.Col)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, m As Long, blnPrint As Boolean
    Dim rsPati As ADODB.Recordset
    
    For i = 1 To vsDept.Rows - 1
        If vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked Then
            m = m + 1
        End If
    Next
    If m = vsDept.Rows - 1 Then
        MsgBox "请至少选择一个科室.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set rsPati = GetPatiSet
    If rsPati.RecordCount = 0 Then
        MsgBox "请至少选择一个病人.", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not IsDate(txtDateEnd.Text) Then
        MsgBox "费用截止时间格式不正确.", vbInformation, gstrSysName
        txtDateEnd.SetFocus
        Exit Sub
    End If
    
    blnPrint = mintPrintMode <> 0
    If mintPrintMode = 2 Then
        If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
    End If
    
    If blnPrint Then
        If gblnStrictCtrl Then   '严格票据管理
            If Trim(txtInvoice.Text) = "" Then
                MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
            mlng领用ID = GetInvoiceGroupID(IIf(gbytInvoiceKind = 0, 3, 1), rsPati.RecordCount, mlng领用ID, mlngShareUseID, txtInvoice.Text, mstrUseType)
            If mlng领用ID <= 0 Then
                Select Case mlng领用ID
                    Case 0 '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                    Case -3
                        MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入", vbInformation, gstrSysName
                        txtInvoice.SetFocus
                End Select
                Exit Sub
            End If
        Else
            If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If MsgBox("共选择了" & rsPati.RecordCount & "位病人,即将依次进行中途结帐!" & _
        vbCrLf & "请准备好后按确定.", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
        Exit Sub
    End If
    
        
    cmdOK.Enabled = False
    Screen.MousePointer = 11
    
    Call SaveBalance(blnPrint, rsPati)
    
    Call LoadPati(Val(vsDept.RowData(vsDept.Row)))
    Screen.MousePointer = 0
    cmdOK.Enabled = True
    
    gblnOK = True
End Sub

Private Sub GetMaxMinDate(ByVal lngPatiID As Long, ByVal strDateMode As String, ByVal DatEnd As Date, ByRef DatMax As Date, ByRef DatMin As Date)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTable As String
    
    '要和过程Zl_结帐费用记录_Patient中的待结费用游标一致,避免产生没有结帐费用的结帐单.
    '大表拆分:经与张永康落实,从SQL分析上来看,可能针对门诊的费用进行结算,但实质上应该只针对住院病人,因此,本次拆分只替换成住院费用记录
    
    strSql = "" & _
    " Select Max(Max时间) DatMax, Min(Min时间) DatMin" & vbNewLine & _
    " From ( Select Max(" & strDateMode & ") Max时间, Min(" & strDateMode & ") Min时间" & vbNewLine & _
    "        From 住院费用记录 A" & vbNewLine & _
    "        Where A.病人id = [1] And A.结帐id Is Null And A.记录状态 <> 0 And Mod(记录性质, 10) In (2, 3) And" & vbNewLine & _
    "             " & strDateMode & " < [2] " & vbCrLf & _
    "             And Not Exists ( Select 1" & vbNewLine & _
    "                              From 住院费用记录 B" & vbNewLine & _
    "                              Where B.NO = A.NO And B.记录性质 = A.记录性质 And B.序号 = A.序号" & vbNewLine & _
    "                              Group By B.NO, B.记录性质, B.序号" & vbNewLine & _
    "                              Having Nvl(Sum(B.实收金额), 0) = Decode(" & IIf(gblnZero, 1, 0) & ", 1, 1 + Nvl(Sum(B.实收金额), 0), 0))" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select Max(" & strDateMode & ") Max时间, Min(" & strDateMode & ") Min时间" & vbNewLine & _
    "       From " & zlGetFullFieldsTable("住院费用记录") & vbNewLine & _
    "       Where A.病人id = [1] And A.结帐id Is Not Null And Mod(记录性质, 10) In (2, 3) And Nvl(A.实收金额, 0) <> Nvl(A.结帐金额, 0) And" & vbNewLine & _
    "             " & strDateMode & " < [2]" & vbNewLine & _
    "       Group By A.NO, A.序号, Mod(A.记录性质, 10), A.记录状态, A.执行状态" & vbNewLine & _
    "       Having Nvl(Sum(A.实收金额), 0) - Nvl(Sum(A.结帐金额), 0) <> 0)"


    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatiID, DatEnd)
    DatMax = Nvl(rsTmp!DatMax, CDate(0))
    DatMin = Nvl(rsTmp!DatMin, CDate(0))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDateStr(DatTmp As Date) As String
    GetDateStr = "To_Date('" & Format(DatTmp, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Function GetBalanceSum(ByVal Dat收款时间 As Date) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select B.结算方式, Sum(B.冲预交) 结算金额" & vbNewLine & _
            "From 病人结帐记录 A, 病人预交记录 B" & vbNewLine & _
            "Where A.收费时间 = [1] And A.操作员姓名 = [2] And A.ID = B.结帐id" & vbNewLine & _
            "Group By B.结算方式"

    On Error GoTo errH
    Set GetBalanceSum = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Dat收款时间, UserInfo.姓名)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SaveBalance(ByRef blnPrint As Boolean, ByRef rsPati As ADODB.Recordset)
    Dim strNO As String, lng结帐ID As Long, datBalance As Date, lngPatientID As Long, i As Long, j As Long
    Dim arrSQL As Variant, DatMax As Date, DatMin As Date, lngNum As Long, blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    
    lblPay.Visible = False
    lblDeposit.Visible = False
    Err = 0: On Error GoTo Errhand:
    datBalance = zlDatabase.Currentdate '记录为统一的结帐时间

    For i = 1 To rsPati.RecordCount
        arrSQL = Array()
        lngPatientID = rsPati!病人ID
        Call GetMaxMinDate(lngPatientID, IIf(gint费用时间 = 1, "发生时间", "登记时间"), CDate(txtDateEnd.Text), DatMax, DatMin)
        
        If Not (DatMax = DatMin And DatMax = CDate(0)) Then '没有待结费用不结帐
            lblInfo.Caption = "当前进度:共" & rsPati.RecordCount & "位,正在进行第" & i & "位," & rsPati!科室 & ":" & rsPati!姓名
            Me.Refresh
            
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            strNO = zlDatabase.GetNextNo(15)
    
            '1.病人结帐记录
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            '58758
            arrSQL(UBound(arrSQL)) = "zl_病人结帐记录_Insert(" & lng结帐ID & "," & "'" & strNO & "'," & lngPatientID & "," & _
                GetDateStr(datBalance) & "," & GetDateStr(DatMin) & "," & GetDateStr(DatMax) & ",1,0,0,2,NULL,2)"
            
            '2.结帐缴款记录:zl_结帐预交记录_Insert,zl_结帐缴款记录_Insert在Zl_结帐费用记录_Patient中调用,因为结算金额现在未知
            '3.住院费用记录
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_结帐费用记录_Patient('" & strNO & "','" & lngPatientID & "'," & lng结帐ID & "," & _
                GetDateStr(CDate(txtDateEnd.Text)) & "," & gint费用时间 & "," & IIf(gblnZero, 1, 0) & _
                ",'" & cbo结算方式.Text & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & GetDateStr(datBalance) & ")"
                    
            '4.开始票据号
            If blnPrint And Trim(txtInvoice.Text) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_票据起始号_Update('" & strNO & "','" & Trim(txtInvoice.Text) & "',3)"
            End If
        
            
            On Error GoTo errH
            gcnOracle.BeginTrans: blnTrans = True
                For j = 0 To UBound(arrSQL)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSQL(j)), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            lngNum = lngNum + 1 '记录实际结帐人数
            
            '票据打印
            If blnPrint Then
                Call frmPrint.ReportPrint(1, strNO, lng结帐ID, mlng领用ID, mlngShareUseID, mstrUseType, txtInvoice.Text, datBalance, "", "", lngPatientID, mintPrintFormat)
                Call RefreshFact
            End If
        End If
        
        rsPati.MoveNext
    Next
        
    If lngNum = 0 Then
        lblInfo.Caption = "选择了" & rsPati.RecordCount & "位病人,但在指定的截止时间前都不存在未结费用!"
    Else
        lblInfo.Caption = "对" & rsPati.RecordCount & "位病人中,存在未结费用的" & lngNum & "位完成了中途结帐."
        
        Set rsTmp = GetBalanceSum(datBalance)
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "结算方式<>'" & cbo结算方式.Text & "'"
            If rsTmp.RecordCount > 0 Then
                lblDeposit.Caption = "冲预交合计：" & Format(rsTmp!结算金额, "0.00")
                lblDeposit.Visible = True
            End If
            rsTmp.Filter = "结算方式='" & cbo结算方式.Text & "'"
            If rsTmp.RecordCount > 0 Then
                If lblDeposit.Visible = False Then
                    lblPay.Left = lblDeposit.Left
                Else
                    lblPay.Left = lblDeposit.Left + lblDeposit.Width + 200
                End If
                lblPay.Caption = cbo结算方式.Text & "结算合计：" & Format(rsTmp!结算金额, "0.00")
                lblPay.Visible = True
            End If
        End If
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    If lngNum > 0 Then
        lblInfo.Caption = "选择了" & rsPati.RecordCount & "位病人,实际对" & lngNum & "位病人完成了中途结帐."
    End If
    Exit Sub
Errhand:
     If ErrCenter = 1 Then Resume
End Sub
Private Sub RefreshFact()
'功能：刷新收费票据号
    If mintPrintMode = 0 Then Exit Sub
    If gblnStrictCtrl Then
        mlng领用ID = CheckUsedBill(IIf(gbytInvoiceKind = 0, 3, 1), IIf(mlng领用ID > 0, mlng领用ID, mlngShareUseID), , mstrUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            txtInvoice.Text = ""
        Else
            '严格：取下一个号码
            txtInvoice.Text = GetNextBill(mlng领用ID)
        End If
    Else
        '松散：取下一个号码
        txtInvoice.Text = IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
    End If
End Sub
Private Function GetPatiSet() As ADODB.Recordset
    Dim strSql As String, str费别 As String, strDeptIDs As String, i As Long
    
    str费别 = Get费别选择
    If str费别 <> "" Then
        If UBound(Split(str费别, ",")) + 1 < vsFeeType.Rows - 1 Then
            str费别 = "," & str费别 & ","
            strSql = " And Instr([2],','||A.费别||',') > 0"
        End If
    End If
    
    For i = 1 To vsDept.Rows - 1
        If Not (vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked) Then
            strDeptIDs = strDeptIDs & "," & vsDept.RowData(i)
        End If
    Next
    strDeptIDs = Mid(strDeptIDs, 2)
    If UBound(Split(strDeptIDs, ",")) + 1 = vsDept.Rows - 1 Then strDeptIDs = ""
    
    If strDeptIDs <> "" Then
        strSql = strSql & " And B.科室ID In(" & strDeptIDs & ")"
    End If
    
    If mstrPatis <> "" Then
        mstrPatis = "," & mstrPatis & ","
        strSql = strSql & " And Instr([1],','||B.病人ID||',') = 0"
    End If
    
    strSql = "Select Distinct C.名称 as 科室,A.姓名,A.病人ID,A.住院次数,A.住院号" & vbNewLine & _
            "From 病人信息 A, 床位状况记录 B, 部门表 C,病案主页 M " & vbNewLine & _
            "Where A.病人id = B.病人ID  " & _
            "       And B.科室ID = C.ID And A.险类 is Null" & _
            "       And Zl_Billclass(A.病人ID,A.主页ID,0)=[3]  " & strSql & vbNewLine & _
            "Order by 科室,住院号"

    On Error GoTo errH
    Set GetPatiSet = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrPatis, str费别, Trim(cbo使用类别.Text))

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Load()
    Set mrsRptFormat = Nothing
    lblInfo.Caption = ""
    Call LoadUseType
    If Not InitData Then
        Unload Me
    End If
    If vsDept.Rows > 1 Then
        vsDept.Row = 0
        vsDept.Row = 1
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset, i As Long

    Set rsTmp = Get费别
    If rsTmp.RecordCount = 0 Then
        MsgBox "费别未设置,不能使用此功能!", vbInformation, gstrSysName
        Exit Function
    Else
        vsFeeType.Rows = rsTmp.RecordCount + 1
        vsFeeType.ColDataType(0) = flexDTBoolean
        vsFeeType.Cell(flexcpChecked, 1, 0, vsFeeType.Rows - 1, 0) = flexChecked
        vsFeeType.Row = 1: vsFeeType.Col = 1: vsFeeType.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsFeeType.TextMatrix(i, 1) = rsTmp!名称
        rsTmp.MoveNext
    Next
    Call LoadDept
    
    txtDateEnd.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    Set rsTmp = Get结算方式("结帐", 2)
    If rsTmp.RecordCount = 0 Then
        MsgBox "没有设置用于结帐场合的非现金结算方式,不能使用此功能!", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 1 To rsTmp.RecordCount
        cbo结算方式.AddItem rsTmp!名称
        rsTmp.MoveNext
    Next
    cbo结算方式.ListIndex = 0
    
    Call RefreshFact
    
    InitData = True
End Function

Private Function Get费别() As ADODB.Recordset
    Dim strSql As String
 
    strSql = "Select 名称,编码 From 费别 Where 服务对象 In (2, 3) And 属性 = 1 Order by 编码"
    On Error GoTo errH
    Set Get费别 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadDept()
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "Select A.ID, A.名称" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where A.ID = B.部门id And B.服务对象 In (2, 3) And B.工作性质 = '临床'" & vbNewLine & _
            " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & vbNewLine & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " And Exists" & vbNewLine & _
            " (Select 1 From 床位状况记录 C Where C.病人id Is Not Null And C.科室id = A.ID) Order by 名称"
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    vsDept.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
        vsDept.Row = 1: vsDept.Col = 1: vsDept.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsDept.TextMatrix(i, 1) = rsTmp!名称
        vsDept.RowData(i) = Val(rsTmp!ID)
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get费别选择() As String
    Dim i As Long, strTmp As String
    
    For i = 1 To vsFeeType.Rows - 1
        If vsFeeType.Cell(flexcpChecked, i, 0) = flexChecked Then strTmp = strTmp & "," & vsFeeType.TextMatrix(i, 1)
    Next
    Get费别选择 = Mid(strTmp, 2)
End Function

Private Sub LoadPati(ByVal lngDeptID As Long)
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long, str费别 As String
    
    str费别 = Get费别选择
    If str费别 <> "" Then
        If UBound(Split(str费别, ",")) + 1 < vsFeeType.Rows - 1 Then
            str费别 = "," & str费别 & ","
            strSql = " And Instr([2],','||A.费别||',')>0"
        End If
    End If
    
    strSql = "" & _
            "   Select Distinct A.病人ID,A.住院号, Nvl(D.姓名,A.姓名) as 姓名, Nvl(D.性别,A.性别) as 性别, " & _
            "               Nvl(D.年龄,A.年龄) as 年龄, B.费用余额 未结费用, 预交余额 可用预交, A.费别" & vbNewLine & _
            "   From 病人信息 A, 病人余额 B,床位状况记录 C,病案主页 D " & vbNewLine & _
            "   Where C.科室id = [1] And C.病人ID = A.病人id And A.病人id = B.病人id(+) " & _
            "               And A.病人id=D.病人ID(+) And A.主页id = D.主页id(+) " & _
            "               And B.性质(+) = 1  And B.类型(+)=2 And A.险类 is Null " & _
            "               And Zl_Billclass(A.病人ID,A.主页ID,0)=[3] " & strSql & vbNewLine & _
            "   Order by A.住院号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngDeptID, str费别, Trim(cbo使用类别.Text))
    vsPati.Rows = 1 '清除数据,但不清除列标头
    vsPati.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
        Else
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
        End If
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    End If
    
    With vsPati
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 1) = "" & rsTmp!住院号
            .TextMatrix(i, 2) = "" & rsTmp!姓名
            .TextMatrix(i, 3) = "" & rsTmp!性别
            .TextMatrix(i, 4) = "" & rsTmp!年龄
            .TextMatrix(i, 5) = Nvl(rsTmp!未结费用, ""): If Val(.TextMatrix(i, 5)) = 0 Then .TextMatrix(i, 5) = ""
            .TextMatrix(i, 6) = Nvl(rsTmp!可用预交, ""): If Val(.TextMatrix(i, 6)) = 0 Then .TextMatrix(i, 6) = ""
            .TextMatrix(i, 7) = "" & rsTmp!费别
            .RowData(i) = Val(rsTmp!病人ID)
            If Len(mstrPatis) > 0 Then
                If InStr("," & mstrPatis & ",", "," & rsTmp!病人ID & ",") > 0 Then
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            End If
            rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then .Row = 1: .Col = 1: .Col = 0
    End With
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.Width < 12060 Then Me.Width = 12060
    If Me.Height < 7635 Then Me.Height = 7635
    With fra
        .Width = ScaleWidth - .Left * 2
    End With
    With picDown
        .Width = ScaleWidth
        .Top = ScaleHeight - .Height - 100
    End With
     With vsFeeType
        .Height = picDown.Top - .Top - 50
        vsDept.Height = .Height
        vsPati.Height = .Height
        vsPati.Width = ScaleWidth - vsPati.Left - 50
     End With
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRptFormat = Nothing
    mstrPatis = ""
    mlng领用ID = 0
End Sub

Private Sub picDown_Resize()
  Err = 0: On Error Resume Next
    With cmdCancel
        .Left = picDown.ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = .Left - cmdOK.Width - 50
    End With
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed Then vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked    '手动点击时点为灰色，改为选择
    
    If Row <> vsDept.Row Then vsDept.Row = Row
    If vsPati.Rows < 2 Then Exit Sub
    
    If vsDept.Cell(flexcpChecked, Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, Row, 0) = flexTSUnchecked Then
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
    Else
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
    End If
    Call SetPatiLists
End Sub
Private Sub vsdept_DblClick()
    If vsDept.MouseCol = 0 And vsDept.MouseRow = 0 Then
        Call SetVSAll(vsDept)
        Call vsDept_AfterEdit(vsDept.Row, vsDept.Col)
        mstrPatis = ""
    End If
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow And NewRow <> 0 Then Call LoadPati(Val(vsDept.RowData(NewRow)))
End Sub



Private Sub vsFeeType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    i = vsDept.Row
    vsDept.Row = 0
    vsDept.Row = i
    
End Sub

Private Sub vsPati_DblClick()
    If vsPati.MouseCol = 0 And vsPati.MouseRow = 0 Then
        If vsPati.Rows < 2 Then Exit Sub
        
        Call SetVSAll(vsPati)
        Call SetDeptState
        Call SetPatiLists
    End If
End Sub

Private Sub vsFeeType_DblClick()
    Dim i As Long
    If vsFeeType.MouseCol = 0 And vsFeeType.MouseRow = 0 Then
        Call SetVSAll(vsFeeType)
        i = vsDept.Row
        vsDept.Row = 0
        vsDept.Row = i
    End If
End Sub

Private Sub SetVSAll(ByRef vsf As VSFlexGrid)
    If vsf.Rows < 2 Then Exit Sub
    vsf.Cell(flexcpChecked, 1, 0, vsf.Rows - 1, 0) = IIf(Val(vsf.Tag) = 1, flexChecked, flexUnchecked)
    vsf.Tag = IIf(Val(vsf.Tag) = 0, 1, 0)
End Sub


Private Sub vsPati_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
        SetPatiLists
    Else
        Call SetPatistr(Row)
    End If
    Call SetDeptState
End Sub

Private Sub SetPatistr(ByVal lngRow As Long)
'功能：记录没有选择的病人ＩＤ
    If vsPati.Cell(flexcpChecked, lngRow, 0) = flexUnchecked Then
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") = 0 Then
            If mstrPatis = "" Then
                mstrPatis = vsPati.RowData(lngRow)
            Else
                mstrPatis = mstrPatis & "," & vsPati.RowData(lngRow)
            End If
        End If
    Else
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") > 0 Then
            mstrPatis = Replace("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",", ",")
            mstrPatis = Mid(mstrPatis, 2)   '去掉前后的
            If mstrPatis <> "" Then mstrPatis = Mid(mstrPatis, 1, Len(mstrPatis) - 1)
        End If
    End If
    If mstrPatis = "," Then mstrPatis = ""
End Sub

Private Sub SetPatiLists()
'功能:检查当前病人列表，把没有选择的加入到变量中，已选择的，从变量中删除
    Dim i As Long
    
    If vsPati.Rows < 2 Then Exit Sub
    
    For i = 1 To vsPati.Rows - 1
        Call SetPatistr(i)
    Next
End Sub

Private Function SetDeptState() As Boolean
'功能：设置科室选择状态
    Dim i As Long, m As Long
    
    For i = 1 To vsPati.Rows - 1
        If vsPati.Cell(flexcpChecked, i, 0) = flexChecked Then m = m + 1
    Next
    If m = vsPati.Rows - 1 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked
    ElseIf m = 0 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed
    End If
End Function

Private Sub vspati_EnterCell()
    If vsPati.Col = 0 Then
        vsPati.Editable = flexEDKbdMouse
    Else
        vsPati.Editable = flexEDNone
    End If
End Sub
Private Sub vsfeetype_EnterCell()
    If vsFeeType.Col = 0 Then
        vsFeeType.Editable = flexEDKbdMouse
    Else
        vsFeeType.Editable = flexEDNone
    End If
End Sub
Private Sub vsDept_EnterCell()
    If vsDept.Col = 0 Then
        vsDept.Editable = flexEDKbdMouse
    Else
        vsDept.Editable = flexEDNone
    End If
End Sub
Private Sub LoadUseType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载使用类别
    '编制:刘兴洪
    '日期:2011-04-28 15:09:10
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, strSql As String
    Dim varData As Variant, varTemp As Variant
    Dim strRptName As String
    Dim strShareInvoice As String
    
    On Error GoTo errHandle
    
    strShareInvoice = zlDatabase.GetPara("结帐发票格式", glngSys, 1137)
    varData = Split(strShareInvoice, "|")
    
    strRptName = IIf(gbytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    
    '票据格式处理
    strSql = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set mrsRptFormat = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    
    strSql = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cbo使用类别
        .Clear
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!名称)
            .ItemData(.NewIndex) = 0
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .ItemData(.NewIndex) = Val(varTemp(1))
                    Exit For
                End If
            Next
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        mstrUseType = cbo使用类别.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

