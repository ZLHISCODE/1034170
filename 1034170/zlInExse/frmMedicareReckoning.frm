VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMedicareReckoning 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "医保病人结帐校对"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdReturnCash 
      Caption         =   "退现"
      Height          =   330
      Left            =   6330
      TabIndex        =   19
      Top             =   255
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消 (&C)"
      Height          =   420
      Left            =   8160
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2745
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txt找补 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   5040
      Width           =   1755
   End
   Begin VB.TextBox txtMargin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
      Width           =   1755
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   9885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&O)"
      Height          =   420
      Left            =   6465
      TabIndex        =   6
      Top             =   5055
      Width           =   1395
   End
   Begin VB.TextBox txt缴款 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   5025
      Width           =   1755
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
      Height          =   3420
      Left            =   5280
      TabIndex        =   17
      Top             =   855
      Width           =   4230
      _cx             =   7461
      _cy             =   6032
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   0
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
   Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
      Height          =   3420
      Left            =   30
      TabIndex        =   2
      Top             =   870
      Width           =   5205
      _cx             =   9181
      _cy             =   6032
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   0
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
   Begin VB.Label lblDelMoney 
      Caption         =   "退支付宝:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7005
      TabIndex        =   18
      Top             =   255
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Label lbl应缴 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应缴:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7680
      TabIndex        =   15
      Tag             =   "应缴:"
      Top             =   4410
      Width           =   600
   End
   Begin VB.Label lbl医保支付 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医保支付:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5280
      TabIndex        =   14
      Tag             =   "医保支付:"
      Top             =   4410
      Width           =   1080
   End
   Begin VB.Label lbl冲预交 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "冲预交:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2280
      TabIndex        =   13
      Tag             =   "冲预交:"
      Top             =   4410
      Width           =   840
   End
   Begin VB.Label lbl预交余额 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "预交余额:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Tag             =   "预交余额:"
      Top             =   4410
      Width           =   1080
   End
   Begin VB.Label lbl缴款 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金缴款"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label lbl找补 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金找补"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   10
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应补金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   960
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结帐金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frmMedicareReckoning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mbytInFun As Byte '0-费用模块调用,1-医保模块调用
Private mlngModul As Long

Private mlng结帐ID As Long
Private mlng病人ID As Long, mlng主页ID As Long
Private mbln中途结帐 As Boolean     '出院结帐,未冲完的预交金额要退为现金
Private mstr保险结算 As String
Private mstr保险信息 As String      '保险类别,保险密码,保险帐号
Private mbln门诊结帐 As Boolean
Private mcur结帐金额 As Currency
Private mcur预交余额 As Currency
Private mintInsure As Integer       '用来判断是否支持分币处理
Private mcur缴款 As Currency
Private mcur应缴金额 As Currency
Private mstr医保号 As String
Private mcur个帐透支 As Currency
Private mintError As Integer
Private mstrStyle As String
Private mblnThreeDepositAfter As Boolean
Private mcur收费误差 As Currency
Private mblnOk  As Boolean
Private mstrPrivs As String
Private mintDefault As Integer      '缺省结算方式行(为0表示没有)
Private mcurMediCare   As Currency  '医保结算合计,根据[mstr保险结算]计算
Private mblnClickOK As Boolean      '窗体只允许点确定退出
Private mblnCent As Boolean         '医保是否支持分币处理
Private mcur个帐余额 As Currency
Private mstrForceNote As String, mstrCardPrivs As String
Private mcur冲预交合计 As Currency
Private mcur预交合计 As Currency
Private mstr住院次数 As String  '住院次数:多个用逗号分离
Private mint预交类别 As Integer
Private mobjCard As Card
Private mbytInvoiceKind As Byte
Private Type TY_BrushCard    '刷卡类型
    str卡号 As String
    str密码 As String
    str交易流水号 As String    '交易流水号
    str交易说明  As String     '交易信息
    str扩展信息 As String    '交易的扩展信息
    dbl帐户余额 As Double
    dblMoney As Double     '当前退款或刷卡金额
End Type
Private mCurBrushCard As TY_BrushCard   '当前的刷卡信息


'模块参数的私有化
Private Const support分币处理 = 25  '医保病人是否处理分币   ,主要是为了便于医保与医院对帐
Private mstrDec As String
Private mBytMoney As Byte '收费分币处理方法
Private mbytMCMode As Byte '医保病人身份证验模式,包括1-门诊,2-住院两种模式,0-表示非医保
Private mbytMzDeposit As Byte '门诊预交缺省使用方式:0-缺省不使用交;1-按结帐金额使用预交;2-使用所有预交
Private mblnFirst As Boolean
Private mrsCardType As ADODB.Recordset '医疗卡类别
Private mobjPayCards As Cards
Private mblnExternal As Boolean, mstrNO As String
Private mFactProperty As Ty_FactProperty, mstrInvoice As String
Private mlng领用ID As Long, mstrUseType As String, mlngShareUseID As Long
Private mintInvoiceFormat As Integer, mintInvoiceMode As Integer, mint结帐类型 As Integer


Private Sub InitBalanceGrid(ByRef vsGrid As VSFlexGrid)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算列头信息
    '编制:刘兴洪
    '日期:2015-05-04 17:33:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsGrid
        .Redraw = flexRDNone
        .Clear
        .Rows = 2: .Cols = 6: i = 0
        .TextMatrix(0, i) = "结算方式": i = i + 1
        .TextMatrix(0, i) = "金额": i = i + 1
        .TextMatrix(0, i) = "结算号码": i = i + 1
        .TextMatrix(0, i) = "性质": i = i + 1
        .TextMatrix(0, i) = "缺省": i = i + 1
        .TextMatrix(0, i) = "卡类别ID": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "金额" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            If .ColKey(i) = "性质" Or .ColKey(i) = "缺省" Or .ColKey(i) = "卡类别ID" Then
                .ColHidden(i) = True: .ColWidth(0) = 0
            End If
        Next
        .ColWidth(.ColIndex("结算方式")) = 1200
        .ColWidth(.ColIndex("金额")) = 1100
        .ColWidth(.ColIndex("结算号码")) = 1450
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadBalance() As Boolean
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim i As Long, str结算方式 As String, blnFind As Boolean
    
    strSql = "" & _
    "   Select B.性质,A.结算方式,A.冲预交,A.结算号码,卡类别ID,A.校对标志" & _
    "   From 病人预交记录 A, 结算方式 B " & _
    "   Where A.结算方式=B.名称(+) And A.结帐ID=[1] And Mod(A.记录性质,10)<>1  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng结帐ID)
    If rsTemp.EOF Then LoadBalance = True: Exit Function
    If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
        rsTemp.Filter = "性质 <> 3 And 性质 <> 4"
        If rsTemp.RecordCount <> 0 Then
            MsgBox "仅用预交结帐时,不能处理预交和医保之外的结算方式的结帐单据!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With vsfMoney
        Do While Not rsTemp.EOF
            str结算方式 = Nvl(rsTemp!结算方式): blnFind = False
            For i = 1 To .Rows - 1
                If str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式"))) Then
                    blnFind = True
                    If i = mintDefault Then Exit For    '通过计算得到
                    '已经按正式结算结果赋值
                    If InStr(",3,4,", "," & .TextMatrix(i, .ColIndex("性质")) & ",") > 0 Then Exit For
                    .TextMatrix(i, .ColIndex("金额")) = Format(Val(Nvl(rsTemp!冲预交)), "0.00")
                    .TextMatrix(i, .ColIndex("结算号码")) = Nvl(rsTemp!结算号码)
                    .TextMatrix(i, .ColIndex("卡类别ID")) = Nvl(rsTemp!卡类别ID)
                    Exit For
                End If
            Next
            rsTemp.MoveNext
        Loop
    End With
    Call ShowMoney(False)
    LoadBalance = True
End Function

Public Function ShowMeFromOut(ByRef frmParent As Object, ByVal strPrivs As String, _
    ByVal lng结帐ID As Long, Optional ByRef blnThreeDeposit As Boolean, Optional ByRef lng病人ID As Long, _
    Optional ByRef int预交类别 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID来校对医保数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-07 15:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, strValue As String

    On Error GoTo errH
    Call initCardSquareData
    mbytMzDeposit = Val(zlDatabase.GetPara("门诊预交缺省使用方式", glngSys, 1137, 2))
    
    mblnExternal = True
    mlng结帐ID = lng结帐ID
    mstrPrivs = strPrivs
    
    strSql = "" & _
    "   Select a.病人ID,a.记录性质,a.结算方式,a.结算号码,b.性质 结算性质,a.冲预交,a.缴款单位, " & _
    "           a.单位开户行,a.单位帐号,C.中途结帐,C.结帐类型,C.NO" & _
    "   From 病人预交记录 a,结算方式 b,病人结帐记录 C" & _
    "   Where a.记录状态 = 1 And a.结算方式 = B.名称 and A.结帐ID=C.ID " & _
    "          And 结帐id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)
    
    mlng病人ID = Val("" & rsTmp!病人ID)
    
    mbln中途结帐 = Val(Nvl(rsTmp!中途结帐))
    
    mint预交类别 = 2
    If Val(Nvl(rsTmp!结帐类型)) = 1 Then mint预交类别 = 1
    
    mbln门诊结帐 = Val(Nvl(rsTmp!结帐类型)) = 1
    
    mstrNO = Nvl(rsTmp!NO)
   
    rsTmp.Filter = "(记录性质=2 And 结算性质=3) or (记录性质=2 And 结算性质=4)"
    If rsTmp.RecordCount > 0 Then mstr保险信息 = zlCommFun.Nvl(rsTmp!缴款单位, " ") & "," & zlCommFun.Nvl(rsTmp!单位开户行, " ") & "," & zlCommFun.Nvl(rsTmp!单位帐号, " ")


    rsTmp.Filter = 0    '不能取实收金额,因为结帐作废再结帐时,费用明细没有实收金额
    strSql = "" & _
    "   Select Sum(nvl(结帐金额,0)) As 结帐金额" & _
    "   From (  Select nvl(结帐金额,0) as 结帐金额    From 门诊费用记录  Where Nvl(附加标志,0) <> 9 And 结帐id = [1]  UNION ALL  " & _
    "           Select nvl(结帐金额,0) as 结帐金额    From 住院费用记录  Where Nvl(附加标志,0) <> 9 And 结帐id = [1] ) "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)

    mcur结帐金额 = Val(Nvl(rsTmp!结帐金额))


    '保险信息
    rsTmp.Filter = 0
    strSql = "" & _
    "   Select 结算方式,金额 From 保险结算明细 " & _
    "   Where 结帐id = [1] And 结算方式<>'现金' and 标志=1"  '医保管控的过程固定写入了一条"现金"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)
    
    mstr保险结算 = ""   '结算方式|结算金额||
    For i = 1 To rsTmp.RecordCount
        mstr保险结算 = mstr保险结算 & "||" & rsTmp!结算方式 & "|" & rsTmp!金额
        rsTmp.MoveNext
    Next
    If mstr保险结算 <> "" Then mstr保险结算 = Mid(mstr保险结算, 3)


    mintInsure = 0
    If mlng病人ID <> 0 Then
        If mint预交类别 = 1 Then
            '门诊
            strSql = "Select 险类,0 As 主页id From 病人信息 Where 病人id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", mlng病人ID)
            If Not rsTmp.EOF Then mintInsure = zlCommFun.Nvl(rsTmp!险类, 0): mlng主页ID = zlCommFun.Nvl(rsTmp!主页ID, 0)
        Else
            strSql = "Select 险类,主页id From 病案主页 Where 病人id = [1]" & _
                     " And 主页id = (Select Max(主页id) From 病案主页 Where 病人id = [1])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", mlng病人ID)
            If Not rsTmp.EOF Then mintInsure = zlCommFun.Nvl(rsTmp!险类, 0): mlng主页ID = zlCommFun.Nvl(rsTmp!主页ID, 0)
        End If
    End If

    mstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 3, 1)))

    mbytInFun = 1
    Me.Show 1, frmParent
    blnThreeDeposit = mblnThreeDepositAfter
    lng病人ID = mlng病人ID
    int预交类别 = mint预交类别
    ShowMeFromOut = mblnOk

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(ByRef frmParent As Object, ByVal lng结帐ID As Long, _
    ByVal lng病人ID As Long, ByVal bln中途结帐 As Boolean, _
        ByVal cur结帐金额 As Currency, ByVal str保险结算 As String, ByVal str保险信息 As String, _
        ByVal intInsure As Integer, ByVal str缺省金额位数 As String, ByVal byt缺省分币方式 As Byte, _
        ByVal cur缴款 As Currency, ByVal str医保号 As String, _
        ByVal bytMCMode As Byte, ByVal str住院次数 As String, _
        ByVal int预交类别 As Integer, ByRef blnThreeDepositAfter As Boolean, ByVal strStyle As String, ByRef rsCardType As ADODB.Recordset, _
        ByRef objPayCards As Cards, ByRef objCard As Card, ByVal strPrivs As String, bln门诊结帐 As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保预结算与结算不一致的较对窗体
    '入参:bytMCMode=医保病人身份证验模式,包括1-门诊,2-住院两种模式,0-表示非医保
    '     int预交类别-预交类别:0-门诊和住院;1-门诊;2-住院
    '     objCard-上次返回三方卡对象
    '     bln门诊结帐- 是否门诊结帐
    '返回:校对成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-23 10:34:39
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng结帐ID = lng结帐ID: mstr住院次数 = str住院次数: mint预交类别 = int预交类别
    mlng病人ID = lng病人ID: mbln中途结帐 = bln中途结帐
    mstr保险结算 = str保险结算: mstr保险信息 = str保险信息     '用于医保存储:保险类别,保险密码,保险帐号
    mcur结帐金额 = cur结帐金额: mintInsure = intInsure: mstr医保号 = str医保号
    mcur缴款 = cur缴款: mstrDec = str缺省金额位数: mbytMzDeposit = Val(zlDatabase.GetPara("门诊预交缺省使用方式", glngSys, 1137, 2))
    mBytMoney = byt缺省分币方式: mbytMCMode = bytMCMode
    mbln门诊结帐 = bln门诊结帐
    mblnThreeDepositAfter = blnThreeDepositAfter
    mstrStyle = strStyle
    mstrPrivs = strPrivs
    Set mobjCard = objCard
    Set mobjPayCards = objPayCards
    Set mrsCardType = rsCardType
    
    If gblnLED Then 'Led才显示余额
        mcur个帐余额 = gclsInsure.SelfBalance(mlng病人ID, mstr医保号, IIf(mbytMCMode = 1, 10, 40), mcur个帐透支, mintInsure)
    End If
    mbytInFun = 0
    Me.Show 1, frmParent
    ShowMe = mblnOk
    blnThreeDepositAfter = mblnThreeDepositAfter
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    mblnClickOK = True: Unload Me
End Sub

Private Sub cmdOK_Click()
    '检查数据
    Dim strNotValiedNos As String, blnPrint As Boolean
    Dim i As Long, cllPro As Collection, blnPrintBillEmpty As Boolean
    Dim str结帐结算 As String, str误差NO As String, str冲预交 As String
    Dim str三方结帐结算 As String, strCash As String
    Dim objCard As Card, strSql As String
    If Val(txtMargin.Text) <> 0 Then
        If InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then
            If Val(txtMargin.Text) > 0 Then
                MsgBox "病人支付金额不足,请按所显示的差额补款。", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            Else
                MsgBox "病人支付金额过多,请按所显示的差额退款。", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            End If
        Else
            If Val(txtMargin.Text) > 0 Then
                MsgBox "病人冲预交金额不足,请按所显示的差额调整冲预交金额！", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            Else
                MsgBox "病人冲预交金额过多,请按所显示的差额调整冲预交金额！", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
        For i = 1 To vsfMoney.Rows - 1
            If vsfMoney.RowData(i) = 999 Then
                If Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额"))) < 0 Then
                    MsgBox "仅用预交结帐情况下，结帐不支持退款！", vbExclamation, gstrSysName
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    If CheckThreePayDepositValied(objCard) = False Then Exit Sub
    
    '更新数据
    str结帐结算 = ""
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 And .RowData(i) <> 999 Then
                str结帐结算 = str结帐结算 & "||" & .TextMatrix(i, .ColIndex("结算方式")) & "|" & Val(.TextMatrix(i, .ColIndex("金额"))) & "|"
                If InStr(",3,4,", "," & Val(.TextMatrix(i, .ColIndex("性质"))) & ",") = 0 Then
                     'Oracle过程根据结算号码字段判断是否医保,所以缴费的结算号码不能含有,号
                     '结算方式|结算金额|结算号码||.....
                    str结帐结算 = str结帐结算 & IIf(.TextMatrix(i, .ColIndex("结算号码")) = "", " ", .TextMatrix(i, .ColIndex("结算号码")))
                Else
                    str结帐结算 = str结帐结算 & IIf(mstr保险信息 = "", " ", mstr保险信息)
                    '结算方式|结算金额|保险类别,保险密码,保险帐号||.....
                End If
            End If
        Next
    End With
    If str结帐结算 <> "" Then str结帐结算 = Mid(str结帐结算, 3)
   
    
    For i = 1 To vsDeposit.Rows - 1
        If Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("冲预交"))) <> 0 Then     'ID|单据号|金额|记录状态||  Id为零表示冲预交余款(非第一次)
            str冲预交 = str冲预交 & "||" & vsDeposit.TextMatrix(i, vsDeposit.ColIndex("ID"))
            str冲预交 = str冲预交 & "|" & vsDeposit.TextMatrix(i, vsDeposit.ColIndex("单据号"))
            str冲预交 = str冲预交 & "|" & Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("冲预交")))
            str冲预交 = str冲预交 & "|" & Val(vsDeposit.RowData(i))
        End If
    Next
    If str冲预交 <> "" Then str冲预交 = Mid(str冲预交, 3)
    Set cllPro = New Collection
    
    If Not objCard Is Nothing Then
        '结算方式|结算金额|卡类别ID|卡号|交易流水号|交易说明||...
        str三方结帐结算 = objCard.结算方式
        str三方结帐结算 = str三方结帐结算 & "|" & -1 * mCurBrushCard.dblMoney
        str三方结帐结算 = str三方结帐结算 & "|" & objCard.接口序号
        str三方结帐结算 = str三方结帐结算 & "|" & IIf(mCurBrushCard.str卡号 = "", " ", mCurBrushCard.str卡号)
        str三方结帐结算 = str三方结帐结算 & "|" & IIf(mCurBrushCard.str交易流水号 = "", " ", mCurBrushCard.str交易流水号)
        str三方结帐结算 = str三方结帐结算 & "|" & IIf(mCurBrushCard.str交易说明 = "", " ", mCurBrushCard.str交易说明)
    End If
    
    If mstrForceNote <> "" Then
        With vsDeposit
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("卡类别ID"))) <> 0 And Val(.TextMatrix(i, .ColIndex("是否退现"))) = 0 Then
                    strCash = strCash & "," & .TextMatrix(i, .ColIndex("结算方式")) & Format(.TextMatrix(i, .ColIndex("冲预交")), "0.00") & "元"
                End If
            Next i
            If strCash <> "" Then strCash = Mid(strCash, 2)
            mstrForceNote = mstrForceNote & strCash
        End With
    End If
    
    'Zl_住院收费结算_Update
    strSql = "Zl_住院收费结算_Update("
    '  结帐id_In       住院费用记录.结帐id%Type,
    strSql = strSql & "" & mlng结帐ID & ","
    '  结帐结算_In     Varchar2, --结帐结算_IN-非医保时:结算方式|结算金额|结算号码||.....医保时:结算方式|结算金额|保险类别,保险密码,保险帐号||.....
    strSql = strSql & "" & IIf(str结帐结算 = "", "NULL", "'" & str结帐结算 & "'") & ","
    '  冲预交_In       Varchar2, --冲预交_IN= ID|单据号|金额|记录状态||.....
    strSql = strSql & "" & IIf(str冲预交 = "", "Null", "'" & str冲预交 & "'") & ","
    '  缴款_In         病人预交记录.缴款%Type := Null,
    strSql = strSql & "" & IIf(Val(txt缴款.Text) <> 0, "NULL", Val(txt缴款.Text)) & ","
    '  找补_In         病人预交记录.找补%Type := Null,
    strSql = strSql & "" & IIf(Val(txt找补.Text) <> 0, "NULL", Val(txt找补.Text)) & ","
    '  三方帐户结算_In Varchar2 := Null --:结算方式|结算金额|卡类别ID|卡号|交易流水号|交易说明||...
    strSql = strSql & "" & IIf(str三方结帐结算 = "", "NULL", "'" & str三方结帐结算 & "'") & ","
    '  交易说明_In     病人预交记录.交易说明%Type := Null
    strSql = strSql & "'" & mstrForceNote & "')"
    
    zlAddArray cllPro, strSql
    
    On Error GoTo errH
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If ExecuteThreeSwapPayInterface(objCard, mlng结帐ID, mCurBrushCard.dblMoney) = False Then Exit Sub
    
    If mblnExternal Then
        Call ReInitPatiInvoice
        blnPrint = True
        Select Case mintInvoiceMode
        Case 0: blnPrint = False '不打印
        Case 2  '自动打印
            If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) <> vbYes Then
                blnPrint = False
            End If
        End Select
        
        If blnPrint Then
            If gblnStrictCtrl Then   '严格票据管理
                If Trim(mstrInvoice) = "" Then
                    Call RefreshFact
                End If
                mlng领用ID = GetInvoiceGroupID(IIf(mbytInvoiceKind = 0, 3, 1), 1, mlng领用ID, mlngShareUseID, mstrInvoice, mstrUseType)
                If mlng领用ID <= 0 Then
                    Select Case mlng领用ID
                        Case 0 '操作失败
                        Case -1
                            MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                        Case -2
                            MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                        Case -3
                            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内", vbInformation, gstrSysName
                    End Select
                    Exit Sub
                End If
            Else
                If Len(mstrInvoice) <> gbytFactLength And mstrInvoice <> "" Then
                    MsgBox "票据号码长度不正确！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        If blnPrint Then
RePrint:
            Call frmPrint.ReportPrint(1, mstrNO, mlng结帐ID, mlng领用ID, mlngShareUseID, mstrUseType, mstrInvoice, zlDatabase.Currentdate, txt缴款.Text, txtMargin.Text, , mintInvoiceFormat, blnPrintBillEmpty, mbytInvoiceKind + 1)
            strSql = "Zl_票据起始号_Update('" & mstrNO & "','" & Trim(mstrInvoice) & "',3)"
            If gblnStrictCtrl And blnPrintBillEmpty = False And _
                ((mbytInvoiceKind = 0 And InStr(1, mstrPrivs, ";收据打印;") > 0) _
                   Or (mbytInvoiceKind <> 0 And InStr(1, mstrPrivs, ";打印门诊收费票据;") > 0)) Then    'blnPrintBillEmpty:55052
                If zlIsNotSucceedPrintBill(3, mstrNO, strNotValiedNos) = True Then
                    If MsgBox("结帐单据为[" & strNotValiedNos & "]的结帐票据打印未成功,是否重新打印结帐票据?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                    strSql = "Zl_票据起始号_Update('" & mstrNO & "','" & "',3)"
                End If
            End If
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        End If
    End If
    
    mblnOk = True: mblnClickOK = True: Unload Me
    Exit Sub
    
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If Not objCard Then mblnClickOK = True: Unload Me
End Sub

Private Sub ReInitPatiInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim intInsure As Integer
    intInsure = mintInsure
    lng病人ID = mlng病人ID
    lng主页ID = mlng主页ID
    mlng领用ID = 0
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng病人ID, lng主页ID, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType, IIf(mint预交类别 = 1, "1", "2"))
    mintInvoiceMode = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    Call RefreshFact
    mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, IIf(mint预交类别 = 1, "1", "2"))
End Sub

Private Sub RefreshFact()
    Dim bytInvoiceKind As Byte
    '功能：刷新收费票据号
    If mintInvoiceMode = 0 Then Exit Sub
    
    If mint预交类别 = 1 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
    End If
    
    mbytInvoiceKind = bytInvoiceKind
    
    If gblnStrictCtrl Then
        mlng领用ID = CheckUsedBill(IIf(bytInvoiceKind = 0, 3, 1), IIf(mlng领用ID > 0, mlng领用ID, mlngShareUseID), , mstrUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            mstrInvoice = ""
        Else
            '严格：取下一个号码
            mstrInvoice = GetNextBill(mlng领用ID)
        End If
    Else
        '松散：取下一个号码
        mstrInvoice = IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
    End If
End Sub

Private Sub cmdReturnCash_Click()
    Dim dblMoney As Double, lngRow As Long
    Dim str操作员姓名 As String, strDBUser As String
    Dim strPrivs As String
    Dim intCount As Integer, intNotCashCount As Integer
    If mstrForceNote <> "" Then Exit Sub
    
    Call GetDelThreeCardDepositInfor(intCount, intNotCashCount, mblnThreeDepositAfter, mstrStyle)
    If mstrStyle = "" Then Exit Sub
    
    If InStr(";" & mstrCardPrivs & ";", ";三方退款强制退现;") = 0 And intNotCashCount > 0 Then
        str操作员姓名 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
        If str操作员姓名 = "" Then
            MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        mstrForceNote = str操作员姓名 & "强制退现:"
    Else
        If intNotCashCount <> 0 Then
            If MsgBox("选择的结算卡不支持退现,是否强制将其退现？", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        End If
        
        mstrForceNote = UserInfo.姓名 & "强制退现:"
    End If
    
    Call ShowMoney(True)

End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
    Call LedDisplayBank
    
End Sub
Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim rs应用场合 As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim arrMediCare As Variant
    Dim bln允许个帐 As Boolean, blnExist As Boolean
    Dim str可用的医保结算方式 As String
    Dim intCount As Integer
    
    mlngModul = 1137
    
    '变量初始
    If mintInsure <> 0 Then
        mblnCent = gclsInsure.GetCapability(support分币处理, , mintInsure)
    Else
        mblnCent = Not gBytMoney = 0
    End If
    
    mcur收费误差 = 0
    mblnOk = False
    mblnClickOK = False
    mintDefault = 0
    mcurMediCare = 0
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    
    '确定和取消按钮
    If mbytInFun = 0 Then
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False
    Else
        cmdCancel.Visible = True
    End If
    
    '显示预交明细
    Call AdjustDepost
    Set rsTmp = GetDepositBefor(mlng病人ID, mstr住院次数, mint预交类别)
    intCount = 0
    If Not rsTmp Is Nothing Then
        'mbytMzDeposit As Byte '门诊预交缺省使用方式:0-缺省不使用交;1-按结帐金额使用预交;2-使用所有预交
        With vsDeposit
            .Redraw = flexRDNone
            .Rows = IIf(rsTmp.RecordCount <> 0, rsTmp.RecordCount, 1) + 1
            .Cell(flexcpBackColor, 1, .ColIndex("冲预交"), .Rows - 1, .ColIndex("冲预交")) = txtMoney.BackColor
            .Cell(flexcpBackColor, 1, .ColIndex("余额"), .Rows - 1, .ColIndex("余额")) = 12900351
            
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(Nvl(rsTmp!记录状态))
                .Cell(flexcpData, i, .ColIndex("ID")) = Nvl(rsTmp!卡类别ID) & "||" & Nvl(rsTmp!转帐及代扣) & "||" & Nvl(rsTmp!退现) & "||" & Nvl(rsTmp!缺省退现)
                
               If Val(Nvl(rsTmp!卡类别ID)) <> 0 And Nvl(rsTmp!缺省退现) = 0 Then
                    If mblnExternal Then
                        If InStr("," & mstrStyle & ",", rsTmp!结算方式) = 0 Then
                            mstrStyle = mstrStyle & "," & rsTmp!结算方式
                        End If
                    End If
                    intCount = intCount + 1
                End If
                
                .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(rsTmp!ID))
                .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsTmp!NO)
                .TextMatrix(i, .ColIndex("日期")) = Format(rsTmp!日期, "yyyy-MM-dd")
                .TextMatrix(i, .ColIndex("结算方式")) = IIf(IsNull(rsTmp!结算方式), "", rsTmp!结算方式)
                .TextMatrix(i, .ColIndex("余额")) = Format(rsTmp!金额, "0.00")
                
                .TextMatrix(i, .ColIndex("冲预交")) = Format(rsTmp!金额, "0.00")
                .TextMatrix(i, .ColIndex("预交ID")) = Val(Nvl(rsTmp!预交ID))
                .TextMatrix(i, .ColIndex("卡类别ID")) = Val(Nvl(rsTmp!卡类别ID))
                .TextMatrix(i, .ColIndex("是否退现")) = Val(Nvl(rsTmp!退现))
                rsTmp.MoveNext
            Next
            If intCount > 1 And InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then
                mblnThreeDepositAfter = True
            End If
            
            .Row = 1: .Col = .ColIndex("冲预交")
            .Redraw = flexRDBuffered
            
            If mblnExternal And mstrStyle <> "" Then
                mstrStyle = Mid(mstrStyle, 2)
            End If
        End With
    End If
    
    '显示保险结算及现付结算方式,即使不支持使用个帐,也更出来,反正医保的不允许改
    arrMediCare = Array()                   '结算方式|结算金额||
    If mstr保险结算 <> "" Then arrMediCare = Split(mstr保险结算, "||")
    
    On Error GoTo errH
    
    If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
        strSql = _
        " Select Distinct B.编码,B.名称,B.性质,A.缺省标志,1 As 位置" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where (B.性质=3 OR B.性质=4)  " & _
        "       And B.名称=A.结算方式(+) and instr(',7,8,',','||性质||',')=0 " & _
        " Union " & _
        " Select Null As 编码, '冲预交' As 名称, 999 As 性质,0 As 缺省标志,0 As 位置" & _
        " From Dual " & _
        " Order By 位置,性质,编码"
    Else
        strSql = _
        " Select Distinct B.编码,B.名称,B.性质,A.缺省标志,1 As 位置" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where ((A.应用场合='结帐' And B.性质<>3 And B.性质<>4) OR (B.性质=3 OR B.性质=4))  " & _
        "       And B.名称=A.结算方式(+) and instr(',7,8,',','||性质||',')=0 " & _
        " Union " & _
        " Select 编码,名称,性质,缺省标志,0 As 位置" & _
        " From 结算方式 Where 性质=9 " & _
        " Order By 位置,性质,编码"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    strSql = "Select 应用场合,结算方式 From 结算方式应用 Where 应用场合='结帐'"
    Set rs应用场合 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Call InitBalanceGrid(vsfMoney)
    With vsfMoney
        .Redraw = flexRDNone
        .Rows = IIf(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        i = 1
        Do While Not rsTmp.EOF
            .RowData(i) = Nvl(rsTmp!性质, 1)                '用来判断是否可以修改金额,以及是否是现金
            .TextMatrix(i, .ColIndex("结算方式")) = rsTmp!名称
            .TextMatrix(i, .ColIndex("金额")) = "0.00"
            .TextMatrix(i, .ColIndex("性质")) = Nvl(rsTmp!性质, 1)
            .TextMatrix(i, .ColIndex("缺省")) = Nvl(rsTmp!缺省标志, 0)
            '缺省结算方式(没有则用现金) 不可能是医保
            If InStr(",3,4,", "," & Nvl(rsTmp!性质, 1) & ",") = 0 Then
                If Nvl(rsTmp!缺省标志, 0) = 1 Then mintDefault = i
                If Nvl(rsTmp!性质, 1) = 1 And mintDefault = 0 Then mintDefault = i
                If Nvl(rsTmp!性质, 1) = 9 And mintError = 0 Then
                    mintError = i: .Row = i: .Col = 0
                    .CellForeColor = vbRed
                End If
                i = i + 1
            Else
                '保险结算
                blnExist = False
                For j = 0 To UBound(arrMediCare)
                    If Split(arrMediCare(j), "|")(0) = rsTmp!名称 Then
                        blnExist = True
                        rs应用场合.Filter = "结算方式='" & rsTmp!名称 & "'"
                        If rs应用场合.EOF And Nvl(rsTmp!性质, 1) <> 9 Then
                            MsgBox "注意:结算方式[" & rsTmp!名称 & "]未设置应用于[结帐]场合,请到[结算方式管理]中设置!", vbInformation, gstrSysName
                        End If
                        
                        .TextMatrix(i, .ColIndex("金额")) = Split(arrMediCare(j), "|")(1)
                        .TextMatrix(i, .ColIndex("结算号码")) = ""    '无结算号码
                        mcurMediCare = mcurMediCare + Val(.TextMatrix(i, .ColIndex("金额")))
                        Exit For
                    End If
                Next
                If blnExist Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE7CFBA
                     i = i + 1
                End If
                str可用的医保结算方式 = str可用的医保结算方式 & "," & rsTmp!名称
            End If
            rsTmp.MoveNext
        Loop
        .Rows = i: .Redraw = True
    End With
    
    '先检查每一种医保结算方式是否都存在
    If mstr保险结算 <> "" Then
        str可用的医保结算方式 = str可用的医保结算方式 & ","
        For j = 0 To UBound(arrMediCare)
            If InStr(str可用的医保结算方式, "," & Split(arrMediCare(j), "|")(0) & ",") <= 0 Then
                MsgBox "医保结算方式[" & Split(arrMediCare(j), "|")(0) & "]未设置,请先到[结算方式管理]中设置!", vbInformation, gstrSysName
                cmdCancel.Visible = True
                cmdOK.Visible = False
            End If
        Next
    End If
    
    
    '结帐金额
    txtTotal.Text = Format(mcur结帐金额, mstrDec)
    
    '冲预交,结帐金额减去医保正式结帐后的余额
    Call ShowMoney(True)
    If LoadBalance = False Then
        cmdCancel.Visible = True
        cmdOK.Visible = False
    End If
    
    If mintDefault > 0 Then
        vsfMoney.Row = mintDefault: vsfMoney.Col = 0
        vsfMoney.CellFontBold = True
        vsfMoney.Col = 1
    Else        '结算方式没有缺省值,并且无现金方式的情况
        vsfMoney.Row = 1: vsfMoney.Col = 1
    End If
    txt缴款.Text = Format(mcur缴款, "0.00")
    Call LedDisplayBank
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function RecalDeposit(ByRef cur结帐合计 As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算冲预交
    '入参:cur结帐合计-当前的结帐金额
    '出参:cur结帐合计-返回未冲销完成的结帐金额
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-20 14:04:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln负数 As Boolean, i As Long, varData As Variant
    
    On Error GoTo errHandle
    bln负数 = cur结帐合计 < 0
        
    If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
       With vsDeposit
           For i = 1 To .Rows - 1
               If cur结帐合计 = 0 Then
                   .TextMatrix(i, .ColIndex("冲预交")) = "0.00"
               Else
                   If Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Then
                       .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                   Else
                       .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                   End If
                   cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
               End If
           Next i
       End With
       RecalDeposit = True
       Exit Function
    End If
    
    With vsDeposit
        For i = 1 To .Rows - 1
            '卡类别ID||代扣||是否退现||缺省退现
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||", "||")
            
            'mbytMzDeposit-针对门诊结帐有效,0-表示全清;1-代表根据结帐金额来分摊预交;2-预交款全冲
            If mbln门诊结帐 And mbytMzDeposit = 0 Then
                '门诊结帐不使用冲预交
                 .TextMatrix(i, .ColIndex("冲预交")) = Format(0, "0.00")
            ElseIf mblnThreeDepositAfter Then
                '1.三方预交
                If Val(varData(0)) <> 0 Or Val(varData(3)) = 1 Then
                    If bln负数 And Val(varData(3)) <> 1 Then
                        .TextMatrix(i, .ColIndex("冲预交")) = Format(0, "0.00")
                    ElseIf Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Then
                        .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                        cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
                    Else
                        .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                        cur结帐合计 = 0
                    End If
                Else
                   If mbln门诊结帐 Then
                        If mbytMzDeposit = 2 Then
                            '预交款全冲
                            .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                            cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
                        Else
                            '按结帐金额冲
                            If Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Then
                                 .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                                cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
                            Else
                                .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                                cur结帐合计 = 0
                            End If
                            
                        End If
                   Else
                        If Not mbln中途结帐 Or Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Then
                            '出院结帐全部都冲完(冲多了就退现付)
                             .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                            cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
                        Else
                            .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                            cur结帐合计 = 0
                        End If
                    End If
                End If
            ElseIf Not mbln中途结帐 Or (mbln门诊结帐 And mbytMzDeposit = 2) Then
               '2.出院结帐全部都冲完(冲多了就退现付)或门诊结帐缺省全部冲销预交
                .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
            Else
                '3.中途结帐只冲足够的
                If cur结帐合计 = 0 Then
                     .TextMatrix(i, .ColIndex("冲预交")) = "0.00"
                Else
                    If Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Then
                         .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                         cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
                    Else
                         .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                         cur结帐合计 = 0
                    End If
                End If
            End If
        Next
    End With
    RecalDeposit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ShowMoney(Optional ByVal blnAutoSet As Boolean) As String
    '功能：设置和显示界面的各种金额

    Dim i As Long, j As Long
    Dim cur结帐合计 As Currency, curMoney As Currency, curOwn As Currency
    Dim cur预交合计 As Currency, cur冲预交合计 As Currency, cur应缴金额 As Currency
    Dim bln存在补款 As Boolean  '只有当没有缺省结算方式,或者修改缺省结算方式的金额时,才有
    Dim curTmp As Currency
    Dim varData As Variant
    Dim bln负数 As Boolean
    
    
    '设置自动冲预交额及余款的结算金额
    '---------------------------------------------------------------------------------------------
    Call ShowDelThreeSwap

    If blnAutoSet Then
        cur结帐合计 = mcur结帐金额 - mcurMediCare
        If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
            With vsfMoney
                For i = 1 To .Rows - 1
                    If .RowData(i) = 999 Then
                        .TextMatrix(i, 1) = Format(cur结帐合计, "0.00")
                    End If
                Next i
            End With
         End If
        '重新计算预交款
        Call RecalDeposit(cur结帐合计)
    Else
        '修改冲预交或结算金额后
        cur结帐合计 = mcur结帐金额 - GetSumMoney
        If mintDefault <> 0 And (Not Me.ActiveControl Is vsfMoney Or _
                                Me.ActiveControl Is vsfMoney And mintDefault <> vsfMoney.Row) Then
            With vsfMoney
                If Val(.TextMatrix(mintDefault, .ColIndex("性质"))) And mblnCent Then   '现金时要进行分币处理
                    .TextMatrix(mintDefault, .ColIndex("金额")) = Format(CentMoney(Val(.TextMatrix(mintDefault, .ColIndex("金额"))) + cur结帐合计), "0.00")
                Else
                    .TextMatrix(mintDefault, .ColIndex("金额")) = Format(Val(.TextMatrix(mintDefault, .ColIndex("金额"))) + cur结帐合计, "0.00")
                End If
            End With
        Else
            bln存在补款 = True
        End If
    End If
    
    '显示当前冲预交额及差额
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetSumMoney
    
    '这里是差额,不一定用现金,所以不处理分币,lblDelMoney.Tag本次退到三方帐户的金额
    curOwn = Val(txtTotal.Text) - curMoney
    txtMargin.Text = Format(curOwn, "0.00")
    
    '根据差额自动补平并计算'剩余部份尝试设置到缺省结算方式上
    '-----------------------------------------------------------------------------------------------------
    If Val(txtMargin.Text) <> 0 And mintDefault <> 0 And (vsfMoney.Row <> mintDefault Or blnAutoSet) Then
        curTmp = Val(vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("金额"))) + curOwn
        If Abs(curTmp) >= 0.01 Then
            If mintError <> 0 And mblnCent Then
                vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("金额")) = Format(CentMoney(curTmp), "0.00")
            Else
                vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("金额")) = Format(curTmp, "0.00")
            End If
        Else
            vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("金额")) = "0.00"
        End If
        txtMargin.Text = "0.00"
    End If
    
    '计算误差金额(结算金额-结帐金额)
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetSumMoney(cur预交合计, cur冲预交合计, cur应缴金额)

    '有可能应补差额正好是处理分币的误差部份,就不显示了
    If Val(txtMargin.Text) <> 0 And mintDefault <> 0 Then
        If Abs(Val(txtMargin.Text)) < 0.1 Or gBytMoney = 5 And Abs(Val(txtMargin.Text)) < 0.3 Then
            If CentMoney(Val(vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("金额"))) + Val(txtMargin.Text)) = Val(vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("金额"))) Then
                txtMargin.Text = "0.00"
            End If
        End If
    End If
    
    '可能应补部份是小数点的正常误差部份,如果四舍五入小于1分,就不显示了
    If Val(txtMargin.Text) <> 0 And mcur收费误差 + curOwn = 0 And Abs(curOwn) <= 0.005 Then
        txtMargin.Text = "0.00"
    End If
        
    If mintError <> 0 And Val(txtMargin.Text) = 0 Then
        vsfMoney.TextMatrix(mintError, vsfMoney.ColIndex("金额")) = Format(Val(txtTotal.Text) - curMoney, mstrDec)
        If Val(txtTotal.Text) - curMoney <> 0 Then
            vsfMoney.RowHidden(mintError) = False
        Else
            vsfMoney.RowHidden(mintError) = True
        End If
    Else
        mcur收费误差 = Format(curMoney - Val(txtTotal.Text), mstrDec)
        vsfMoney.TextMatrix(mintError, vsfMoney.ColIndex("金额")) = Format(vsfMoney.TextMatrix(mintError, vsfMoney.ColIndex("金额")), mstrDec)
    End If
    
    lbl预交余额.Caption = lbl预交余额.Tag & Format(cur预交合计, "0.00")
    lbl预交余额.ToolTipText = "本次未冲预交之前的预交余额"
    mcur预交合计 = cur预交合计
    lbl冲预交.Caption = lbl冲预交.Tag & Format(cur冲预交合计, "0.00")
    mcur冲预交合计 = cur冲预交合计
    lbl医保支付.Caption = lbl医保支付.Tag & Format(mcurMediCare, "0.00")
    lbl应缴.Caption = lbl应缴.Tag & Format(cur应缴金额, "0.00")
    mcur应缴金额 = cur应缴金额
    
    lbl预交余额.Left = vsDeposit.Left
    lbl冲预交.Left = lbl预交余额.Left + lbl预交余额.Width + 600
    lbl医保支付.Left = vsfMoney.Left
    lbl应缴.Left = lbl医保支付.Left + lbl医保支付.Width + 600
    
    
    Call Calc找补
    Call LedDisplayBank
End Function

Private Sub vsfMoney_GotFocus()
        Call LedDisplayBank
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then        '输入限制
        If InStr(txtMoney.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Beep: Exit Sub
        
        If txtMoney.Left > vsfMoney.Left Then   '结算输入
            If vsfMoney.Col = vsfMoney.Cols - 1 Then    '结算号码,逗号用来在过程中判断是否是医保结算方式
                If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        Else    '预交输入
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
         '结算输入确认
        If txtMoney.Left > vsfMoney.Left Then
            If vsfMoney.Col = vsfMoney.Cols - 1 Then    '输入结算号
                If InStr(txtMoney.Text, "'") > 0 Or InStr(txtMoney.Text, "|") > 0 Or InStr(txtMoney.Text, ",") > 0 Then
                    Exit Sub
                End If
                
                vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Trim(txtMoney.Text)
                txtMoney.Visible = False
            Else
                If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                    zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
                End If
                If Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.ColIndex("性质"))) = 1 And mblnCent Then
                    txtMoney.Text = Format(CentMoney(Val(txtMoney.Text)), "0.00")
                End If
                                
                If Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col)) <> Format(Val(txtMoney.Text), "0.00") Then
                    vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Format(Val(txtMoney.Text), "0.00")
                    txtMoney.Visible = False
                    vsfMoney.SetFocus   '必须在先,ShowMoney中以此判断
                    
                    Call ShowMoney
                Else
                    txtMoney.Visible = False
                    vsfMoney.SetFocus
                End If
            End If
            
            If vsfMoney.Col < vsfMoney.Cols - 2 Then
                vsfMoney.Col = vsfMoney.Col + 1
            Else
                If vsfMoney.Row = vsfMoney.Rows - 1 Then
                    '下一控件处理
                    If Get应缴 > 0 And txt缴款.Visible Then
                        txt缴款.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                Else
                    '下一行处理
                    If Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.ColIndex("性质"))) = 2 Then
                       If vsfMoney.Col = vsfMoney.Cols - 2 Then
                            vsfMoney.Col = vsfMoney.Cols - 1
                       Else
                            vsfMoney.Row = vsfMoney.Row + 1
                            vsfMoney.Col = vsfMoney.Cols - 2
                       End If
                    Else
                        vsfMoney.Row = vsfMoney.Row + 1
                        vsfMoney.Col = vsfMoney.Cols - 2
                    End If
                    
                    If vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(0) - 2) > 1 Then
                        vsfMoney.TopRow = vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(1) - 2)
                    End If
                End If
            End If
        
        '预交输入确认
        Else
            If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
            End If
            
            '修改不能超过上限
            If Val(txtMoney.Text) > Val(vsDeposit.TextMatrix(vsDeposit.Row, 4)) Then
                txtMoney.Text = Val(vsDeposit.TextMatrix(vsDeposit.Row, 4))
            End If
            
            If Val(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.Col)) <> Format(Val(txtMoney.Text), "0.00") Then
                vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.Col) = Format(Val(txtMoney.Text), "0.00")
                txtMoney.Visible = False
                vsDeposit.SetFocus '必须在先
                
                Call ShowMoney
            Else
                txtMoney.Visible = False
                vsDeposit.SetFocus
            End If
            
            If vsDeposit.Row = vsDeposit.Rows - 1 Then
                '下一控件处理
                vsfMoney.SetFocus
            Else
                '下一行处理
                vsDeposit.Row = vsDeposit.Row + 1
                If vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(0) - 2) > 1 Then
                    vsDeposit.TopRow = vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(1) - 2)
                End If
                vsDeposit.Col = vsDeposit.Cols - 1
            End If
        End If
        
        If Val(txt缴款.Text) > 0 Then Call txt缴款_Change
    End If
End Sub

Private Sub txtMoney_LostFocus()
    txtMoney.Visible = False
End Sub

Private Sub txtMoney_Validate(Cancel As Boolean)
    If txtMoney.Visible Then Call txtMoney_KeyPress(13)
End Sub
Private Sub AdjustDepost()
    Dim bln As Boolean, i As Long
    With vsDeposit
        .Redraw = flexRDNone
        .Clear
        .Rows = 2: .Cols = 9: i = 0
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "单据号": .ColWidth(i) = 1100: i = i + 1
        .TextMatrix(0, i) = "日期": .ColWidth(i) = 1050: i = i + 1
        .TextMatrix(0, i) = "结算方式": .ColWidth(i) = 620: i = i + 1
        .TextMatrix(0, i) = "余额": .ColWidth(i) = 1100: i = i + 1
        .TextMatrix(0, i) = "冲预交": .ColWidth(i) = 1100: i = i + 1
        .TextMatrix(0, i) = "预交ID": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "卡类别ID": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否退现": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            .FixedAlignment(i) = IIf(i = 3, flexAlignLeftCenter, flexAlignCenterCenter)
            Select Case .ColKey(i)
            Case "冲预交", "余额"
                .ColAlignment(i) = flexAlignRightCenter
            Case Else
                .ColAlignment(i) = flexAlignLeftCenter
            End Select
        Next
        .Row = 1: .Col = .Cols - 1
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDBuffered
    End With
End Sub
Private Function GetDepositBefor(ByVal lng病人ID As Long, _
    ByVal str住院次数 As String, ByVal int预交类别 As Integer) As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人本次医保结算之前的剩余预交款明细,包含本次冲销的预交
    '入参:lng病人ID-病人ID
    '       str住院次数-住院次数,如:1,2,3
    '       int预交类别-预交类别:0-门诊和住院;1-门诊;2-住院
    '出参:
    '返回:本次医保结算之前的剩余预交款明细,
    '编制:刘兴洪
    '日期:2013-11-14 11:36:01
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strSub1 As String
    Dim strWherePage As String, strPages As String
    On Error GoTo errH
    strPages = "," & str住院次数 & ","
    
    strWherePage = IIf(str住院次数 = "", "", " And instr([3],','||Nvl(A.主页ID,0)||',')>0")
    If int预交类别 <> 0 Then
        strWherePage = strWherePage & " And A.预交类别 =[4]"
    End If
    
    '该子查询用于消除预交款收费及退费时的一正一负,注意系统允许结过帐的预交款进行预交退费,需要加上记录状态判断
    strSub1 = _
    "   Select NO,Sum(Nvl(A.金额,0)) as 金额 " & _
    "   From 病人预交记录 A" & _
    "   Where (A.结帐ID Is Null " & strWherePage & " Or A.结帐ID=[1]) And Nvl(A.金额, 0)<>0 And A.病人ID=[2]" & _
    "   Group by NO  " & _
    "   Having Sum(Nvl(A.金额,0))<>0"
    strSql = _
        "   Select Max(a.Id) As ID, Max(记录状态) As 记录状态, NO, Max(日期) As 日期, a.结算方式, Sum(a.金额) As 金额, 卡类别id, 转帐及代扣, Min(预交id) As 预交id," & vbNewLine & _
        "       Nvl(b.是否退现, 0) As 退现, Nvl(b.是否全退, 0) As 全退, Nvl(b.是否缺省退现, 0) As 缺省退现 " & _
        "   From( " & _
        "       Select A.ID,A.记录状态,A.NO,A.收款时间 as 日期,A.结算方式,Nvl(A.金额,0) as 金额,A.卡类别ID,C.是否转帐及代扣 as 转帐及代扣,A.ID As 预交ID" & _
        "       From 病人预交记录 A,(" & strSub1 & ") B,医疗卡类别 C" & _
        "       Where (A.结帐ID Is Null " & strWherePage & " Or A.结帐ID=[1]) And Nvl(A.金额,0)<>0" & _
        "               And A.结算方式 Not IN (Select 名称 From 结算方式 Where 性质=5)" & _
        "               And A.卡类别ID=C.ID(+) And A.NO=B.NO And A.病人ID=[2]" & _
        "       Union All" & _
        "       Select 0 as ID,A.记录状态,A.NO,Min(A.收款时间) as 日期,A.结算方式,Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0)) as 金额," & _
        "           Max(A.卡类别ID) as 卡类别ID,max(C.是否转帐及代扣) as 转帐及代扣,Min(A.ID) As 预交ID " & _
        "       From 病人预交记录 A,医疗卡类别 C" & _
        "       Where A.记录性质 IN(1,11) And A.结算方式 Not IN (Select 名称 From 结算方式 Where 性质=5) And A.结帐ID is Not NULL And A.卡类别ID=C.ID(+) " & _
        "           And A.结帐ID<>[1] And Nvl(A.金额,0)<>Nvl(A.冲预交,0) And A.病人ID=[2] " & strWherePage & _
        "       Having Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0))<>0" & _
        "       Group by A.记录状态,A.NO,A.结算方式 " & _
        "       ) A, 医疗卡类别 B " & _
        "   Where a.卡类别id = b.Id(+) " & _
        "   Group By a.No, a.结算方式, a.卡类别id, a.转帐及代扣, Nvl(b.是否退现, 0), Nvl(b.是否全退, 0), Nvl(b.是否缺省退现, 0)" & vbNewLine & _
        "   Having Sum(金额) <> 0" & _
        "   Order By Decode(sign(Sum(a.金额)),-1,0,1), Decode(Nvl(a.卡类别id, 0), 0, 0, Decode(Nvl(b.是否退现, 0), 0, 2, 1)) Desc," & vbNewLine & _
        "         Decode(Nvl(a.卡类别id, 0), 0, 0, Decode(Nvl(b.是否全退, 0), 0, 1, 2)) Desc, Nvl(a.卡类别id, 0) Desc, a.No, a.结算方式"
        
    Set GetDepositBefor = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng结帐ID, mlng病人ID, strPages, int预交类别)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetSumMoney(Optional ByRef cur预交合计 As Currency, Optional ByRef cur冲预交合计 As Currency, Optional ByRef cur应缴金额 As Currency) As Currency
    Dim i As Long
    Dim curMoney As Currency
    cur预交合计 = 0: cur冲预交合计 = 0: cur应缴金额 = 0
    With vsDeposit
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            For i = 1 To .Rows - 1
                curMoney = curMoney + Val(.TextMatrix(i, .ColIndex("冲预交")))
                cur预交合计 = cur预交合计 + Val(.TextMatrix(i, .ColIndex("余额")))
                cur冲预交合计 = cur冲预交合计 + Val(.TextMatrix(i, .ColIndex("冲预交")))
            Next
        End If
    End With
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 And .RowData(i) <> 999 Then
                If Val(.TextMatrix(i, .ColIndex("性质"))) <> 9 Then
                    curMoney = curMoney + Val(.TextMatrix(i, .ColIndex("金额")))
                End If
                If InStr(",3,4,9,", "," & Val(.TextMatrix(i, .ColIndex("性质"))) & ",") = 0 Then
                    cur应缴金额 = cur应缴金额 + Val(.TextMatrix(i, .ColIndex("金额")))
                End If
            End If
        Next
    End With
    curMoney = curMoney - Val(lblDelMoney.Tag)
    GetSumMoney = curMoney
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnClickOK Then Cancel = 1: Exit Sub
    mblnExternal = False
    mstrStyle = ""
    mstrForceNote = ""
End Sub

Private Sub vsDeposit_DblClick()
   '刘兴洪:增加权限控制，如果仅用预交结帐，结帐数据不正确，之前刘尔旋更改时，为什么能更改，暂时不知道原因，先限制，避免结帐数据出错
    If InStr(mstrPrivs, ";仅用预交结帐;") > 0 Then Exit Sub
    If Not txtMoney.Visible And vsDeposit.Row >= 1 And vsDeposit.Col = 5 Then
        If mblnThreeDepositAfter And Val(Split(vsDeposit.Cell(flexcpData, vsDeposit.Row, vsDeposit.ColIndex("ID")) & "||", "||")(0)) <> 0 Then
            Exit Sub
        End If
        With txtMoney
            .Left = vsDeposit.Left + vsDeposit.CellLeft + 30
            .Top = vsDeposit.Top + vsDeposit.CellTop + (vsDeposit.CellHeight - txtMoney.Height) / 2 + 15
            .Width = vsDeposit.CellWidth - 30
            .ForeColor = vsDeposit.CellForeColor
            .BackColor = vsDeposit.CellBackColor
            .Alignment = 1
            .Text = vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub vsDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vsDeposit.Col = 0 Then
            vsDeposit.Col = vsDeposit.Col + 1
        ElseIf vsDeposit.Row < vsDeposit.Rows - 1 Then
            vsDeposit.Row = vsDeposit.Row + 1
            vsDeposit.Col = vsDeposit.Cols - 1
            If vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(0) - 2) > 1 Then
                vsDeposit.TopRow = vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(1) - 2)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub vsDeposit_KeyPress(KeyAscii As Integer)
    '刘兴洪:增加权限控制，如果仅用预交结帐，结帐数据不正确，之前刘尔旋更改时，为什么能更改，暂时不知道原因，先限制，避免结帐数据出错
    If InStr(mstrPrivs, ";仅用预交结帐;") > 0 Then Exit Sub
    If Not txtMoney.Visible And KeyAscii <> 13 And KeyAscii <> vbKeyEscape Then
        If mblnThreeDepositAfter And Val(Split(vsDeposit.Cell(flexcpData, vsDeposit.Row, vsDeposit.ColIndex("ID")) & "||", "||")(0)) <> 0 Then
            Exit Sub
        End If
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        With txtMoney
            .Left = vsDeposit.Left + vsDeposit.CellLeft + 30
            .Top = vsDeposit.Top + vsDeposit.CellTop + (vsDeposit.CellHeight - txtMoney.Height) / 2 + 15
            .Width = IIf(vsDeposit.CellWidth - 30 < 0, 50, vsDeposit.CellWidth - 30)
            .ForeColor = vsDeposit.CellForeColor
            .BackColor = vsDeposit.CellBackColor
            .Alignment = 1
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub vsfMoney_DblClick()
     
    If Not txtMoney.Visible And _
        1 = 1 Then
        With vsfMoney
            If InStr(",3,4,9,", "," & .TextMatrix(.Row, .ColIndex("性质")) & ",") > 0 Then Exit Sub
            If .RowData(.Row) = 999 Then Exit Sub
            If .Row <= 0 Or .Col <= .ColIndex("结算方式") Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("卡类别ID"))) > 0 Then Exit Sub
        End With
        
        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = vsfMoney.Left + vsfMoney.CellLeft + 30
            .Top = vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 + 15
            .Width = vsfMoney.CellWidth - 30
            .ForeColor = vsfMoney.CellForeColor
            .BackColor = IIf(vsfMoney.CellBackColor = 0, vbWhite, vsfMoney.CellBackColor)
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub
Private Sub vsfMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsfMoney
        If .Row < 1 Then Exit Sub
        If .Col = 0 Then
            .Col = .Col + 1: Exit Sub
        End If
        
        If .Row < .Rows - 1 Then
            If Val(.TextMatrix(.Row, .ColIndex("性质"))) = 2 Then
               If .Col = .ColIndex("金额") Then
                    .Col = .ColIndex("结算号码")
               Else
                    .Row = .Row + 1
                    .Col = .ColIndex("金额")
               End If
            Else
                .Row = .Row + 1
                .Col = .ColIndex("金额")
            End If
            
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                 .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            Exit Sub
        End If
    End With
    If Get应缴 > 0 Then txt缴款.SetFocus: Exit Sub
    cmdOK.SetFocus
    
End Sub

Private Sub vsfMoney_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And KeyAscii <> 13 And KeyAscii <> vbKeyEscape And _
          1 = 1 Then
        With vsfMoney
            If InStr(",3,4,9,", "," & .TextMatrix(.Row, .ColIndex("性质")) & ",") > 0 Then Exit Sub
            If .RowData(.Row) = 999 Then Exit Sub
            If .Row <= 0 Or .Col <= .ColIndex("结算方式") Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("卡类别ID"))) > 0 Then Exit Sub
        End With
        
        If vsfMoney.Col = 1 Then
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        Else '结算号码特殊字符限制,逗号用来在过程中判断是否是医保结算方式
            If InStr("'||,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        
        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = vsfMoney.Left + vsfMoney.CellLeft + 30
            .Top = vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 + 15
            .Width = vsfMoney.CellWidth - 30
            .ForeColor = vsfMoney.CellForeColor
            .BackColor = IIf(vsfMoney.CellBackColor = 0, vbWhite, vsfMoney.CellBackColor)
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = UCase(Chr(KeyAscii))
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Function Get应缴() As Currency
    Dim i As Long
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("性质"))) = 1 Then
                Get应缴 = Val(.TextMatrix(i, .ColIndex("金额")))
                Exit Function
            End If
        Next
    End With
End Function

 
Private Sub txt缴款_Change()
    Call Calc找补
End Sub
 
Private Sub txt缴款_GotFocus()
    Dim curTotal As Currency
    Call zlControl.TxtSelAll(txt缴款)
    If Not gblnLED Then Exit Sub
    
    curTotal = Get应缴
    '#21 1234.56   --请您付款一千二百三十四点五六元  J
    '#22 1234.56   --预收一千二百三十四点五六元 Y
    '#23 1234.56   --找零一千二百三十四点五六元 Z
    zl9LedVoice.DisplayBank ("")
    If curTotal >= 0 Then
        zl9LedVoice.Speak "#21 " & curTotal
    Else
        zl9LedVoice.Speak "#23 " & Abs(curTotal)
    End If
End Sub
Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
        If txt缴款.Text <> "0.00" Then
            If Val(txt找补.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                txt缴款.SetFocus
                zlControl.TxtSelAll txt缴款
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '病人累加缴款
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc(".") And InStr(txt缴款.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt缴款_LostFocus()
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt缴款_Validate(Cancel As Boolean)
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
'    If Val(txt缴款.Text) = 0 Then Exit Sub
    
'    If CSng(txt找补.Tag) < 0 Then
'        MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
'        Call SelAll(txt缴款): txt缴款.SetFocus
'        Cancel = True: Exit Sub
'    End If
    If Not gblnLED Then Exit Sub
    zl9LedVoice.DispCharge Format(Get应缴, "0.00"), Val(txt缴款.Text), Val(txt找补.Tag)
    zl9LedVoice.Speak "#22 " & txt缴款.Text
    zl9LedVoice.Speak "#23 " & CSng(txt找补.Tag)
    zl9LedVoice.Speak "#3"                  '#3  --请当面点清, 谢谢!
End Sub

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'功能：对指定金额按分币处理规则进行处理,返回处理后的金额
'参数：curMoney=要进行分币处理的金额(为应缴金额,2位小数)
'      mBytMoney=
'         0.不处理
'         1.采取四舍五入法,eg:0.51=0.50;0.56=0.60
'         2.补整收法,eg:0.51=0.60,0.56=0.60
'         3.舍分收法,eg:0.51=0.50,0.56=0.50
'         4.四舍六入五成双,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           四舍六入五成双,详见我国科学技术委员会正式颁布的《数字修约规则》,但根据vb的Round函数,若被舍弃的数字包括几位数字时，不对该数字进行连续修约
'           即银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一
'         5.三七作五、二舍八入,对角进行处理，不需要先对分币进行舍入,即0.29(含)以下都舍掉角，0.80(含)以上都进角，0.3-0.79处理为0.5。
'         6.五舍六入:eg:0.15=0.10:0.16=0.2:   刘兴洪 问题:34519  日期:2010-12-06 09:58:02
'91385,调整“5.三七作五、二舍八入”规则：先对分币进行四舍五入，即0.24(含)以下都舍掉角，0.75(含)以上都进角，0.25-0.74都处理为0.5
'       分币先四舍五入，那么0.00～0.24=0，0.25～0.5=0.50, 0.50～0.74=0.50，0.75～1.00=1，这样舍和入各占50%的比例

    Dim intSign As Integer, curTmp As Currency

    If mBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf mBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '先取两位金额,再处理分币,如:0.248 得0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf mBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf mBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf mBytMoney = 4 Then
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
    ElseIf mBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf mBytMoney = 6 Then
         '刘兴洪 问题:34519 五舍六入:eg:0.15=0.10:0.16=0.2:    日期:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function
 
Private Sub LedDisplayBank()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示医保结算信息
    '编制:刘兴洪
    '日期:2013-10-23 14:50:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double, i As Long
    Dim str个人帐户 As String, str医保其他结算 As String, str老一卡通 As String, str普通结算 As String
    Dim varPara  As Variant, str结算方式 As String
    Dim cur个帐余额 As Currency, dbl现金 As Double, dblMoney As Double
    Dim str医保结算 As String
    
    If Not gblnLED Then Exit Sub
    zl9LedVoice.DisplayBank ""
    If mblnFirst = True Then Exit Sub
    
    str医保结算 = "||帐户余额:" & Format(mcur个帐余额, "0.00")
    With vsfMoney
        For i = 1 To .Rows - 1
            '医保交易
            str结算方式 = Trim(.TextMatrix(i, 0))
            If str结算方式 <> "" Then
                dblMoney = Val(.TextMatrix(i, 1))
                Select Case Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("性质")))
                Case 3
                    str个人帐户 = str个人帐户 & "||" & str结算方式 & ":" & Format(dblMoney, "0.00")
                Case 4
                    str医保其他结算 = str医保其他结算 & "||" & str结算方式 & ":" & Format(dblMoney, "0.00")
                Case 1
                    dbl现金 = dblMoney
                Case Else
                    str普通结算 = str普通结算 & "||" & str结算方式 & ":" & Format(dblMoney, "0.00")
                End Select
            End If
        Next
    End With
    str结算方式 = ""
    If str个人帐户 <> "" Then str医保结算 = str医保结算 & str个人帐户
    If str医保其他结算 <> "" Then str医保结算 = str医保结算 & str医保其他结算
    
    If str医保结算 <> "" Then str结算方式 = str结算方式 & "||医保结算:" & str医保结算
    If str普通结算 <> "" Then str结算方式 = str结算方式 & "" & str普通结算
    If mcur冲预交合计 <> 0 Then str结算方式 = str结算方式 & "||" & "冲预交:" & Format(mcur冲预交合计, "0.00")
    
    If str结算方式 = "" Then Exit Sub
    str结算方式 = Mid(str结算方式, 3)
    varPara = Split(str结算方式, "||")
    
    dblMoney = Val(txt缴款.Text) - dbl现金
    zl9LedVoice.DisplayBank "总费用" & Format(txtTotal.Text, "0.00"), "预交款" & Format(mcur预交合计, "0.00"), _
            "冲预交" & Format(mcur冲预交合计, "0.00"), IIf(dblMoney > 0, "找补", "应缴") & Format(Abs(dblMoney), "0.00")
    '目前最多只能显示10个参数值
    Select Case UBound(varPara)
    Case 0
          zl9LedVoice.DisplayBank varPara(0)
    Case 1
          zl9LedVoice.DisplayBank varPara(0), varPara(1)
    Case 2
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2)
    Case 3
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3)
    Case 4
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4)
    Case 5
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5)
    Case 6
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6)
    Case 7
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7)
    Case 8
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8)
    Case 9
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9)
    Case Else
        str结算方式 = ""
         For i = 10 To UBound(varPara)
            str结算方式 = str结算方式 & ";" & varPara(i)
        Next
        If str结算方式 > "" Then str结算方式 = Mid(str结算方式, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str结算方式
    End Select
    'zl9LedVoice.Speak "#21 " & Format(mcur应缴金额, "0.00")
End Sub

 
Private Sub txt找补_Change()
    txt找补.Tag = ""
End Sub
Private Sub Calc找补()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算找补
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-01-12 17:41:47
    '问题:27360
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl找补 As Double
    Dim cur现金 As Currency, i As Long

    If Val(txt缴款.Text) = 0 Then txt找补.Text = "0.00"
    dbl找补 = FormatEx(Val(txt缴款.Text) - Get应缴, 2)
    txt找补.Text = Format(Abs(dbl找补), "0.00")
    txt找补.Tag = dbl找补
    If dbl找补 <= 0 Then
        lbl找补.Caption = "收款"
        lbl找补.ForeColor = &H0&
    Else
        lbl找补.Caption = "找补"
        lbl找补.ForeColor = vbRed   '35830
    End If
    txt找补.ForeColor = lbl找补.ForeColor
End Sub
Private Function GetThreePayDepositData(ByRef rsTemp As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取三方交易信息
    '出参:rsTemp-返回交易信息(卡类别ID,卡类别名称,结算方式,是否启用,余额,冲预交,剩余款,退预交)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-04-27 09:44:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl冲预交 As Double, dblMoney As Double, dbl余额 As Double
    Dim dblTotal As Double, dblTemp As Double, lngCardTypeID As Long
    Dim varData As Variant
    
    On Error GoTo errHandle
    Set rsTemp = New ADODB.Recordset
    rsTemp.Fields.Append "卡类别ID", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "卡类别名称", adVarChar, 200, adFldIsNullable
    rsTemp.Fields.Append "结算方式", adVarChar, 100, adFldIsNullable
    rsTemp.Fields.Append "是否启用", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "是否退现", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "缺省退现", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "余额", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "冲预交", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "剩余款", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "退预交", adDouble, , adFldIsNullable
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    If mrsCardType Is Nothing Then
        Call initCardSquareData
    ElseIf mrsCardType.State <> 1 Then
        Call initCardSquareData
    End If
    
    dblTotal = Val(txtTotal.Text) - mcurMediCare
    With vsDeposit
        dblMoney = 0: dbl冲预交 = 0: dbl余额 = 0: lngCardTypeID = 0
        For i = 1 To .Rows - 1
            ' 卡类别ID|| 转帐及代扣||退现||缺省退现
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||||", "||")
            If Val(varData(0)) <> 0 And Val(varData(3)) = 0 Then
                
                lngCardTypeID = Val(varData(0))
                dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                rsTemp.Find "卡类别ID=" & lngCardTypeID
                mrsCardType.Filter = "ID=" & lngCardTypeID
                If rsTemp.EOF Then
                    rsTemp.AddNew
                    rsTemp!卡类别ID = lngCardTypeID
                    If Not mrsCardType.EOF Then
                       rsTemp!卡类别名称 = mrsCardType!名称
                       rsTemp!是否启用 = Val(Nvl(mrsCardType!是否启用))
                    Else
                       rsTemp!卡类别名称 = .TextMatrix(i, .ColIndex("结算方式"))
                       rsTemp!是否启用 = 0
                    End If
                    rsTemp!结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                    rsTemp!退预交 = 0
                End If
                rsTemp!余额 = FormatEx(Val(Nvl(rsTemp!余额)) + Val(.TextMatrix(i, .ColIndex("余额"))), 5)
                rsTemp!冲预交 = FormatEx(Val(Nvl(rsTemp!冲预交)) + dbl冲预交, 5)
                rsTemp!剩余款 = FormatEx(Val(Nvl(rsTemp!余额)) - Val(Nvl(rsTemp!冲预交)), 5)
                If FormatEx(dblTotal - dbl冲预交, 6) < 0 Then
                    If dblTotal >= 0 Then
                        dblTemp = dbl冲预交 - dblTotal
                        rsTemp!退预交 = FormatEx(Val(Nvl(rsTemp!退预交)) + dblTemp, 5)
                    Else
                        rsTemp!退预交 = FormatEx(Val(Nvl(rsTemp!退预交)) + dbl冲预交, 5)
                    End If
                    dblTotal = 0
                Else
                    dblTotal = FormatEx(dblTotal - dbl冲预交, 6)
                End If
                rsTemp.Update
            End If
        Next
    End With
    GetThreePayDepositData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowDelThreeSwap()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示退支付宝信息
    '编制:刘兴洪
    '日期:2015-04-27 11:09:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strTittle As String
    Dim intCount  As Integer, intNotCashCount  As Integer
    
    On Error GoTo errHandle
    Call GetDelThreeCardDepositInfor(intCount, intNotCashCount, mblnThreeDepositAfter, mstrStyle)


    
    lblDelMoney.Visible = False
    cmdReturnCash.Visible = False
    lblDelMoney.Tag = "0"
    
    If mstrForceNote <> "" Then
        mblnThreeDepositAfter = False
        GoTo BrushWin
    End If
    If mblnThreeDepositAfter Then
    
    
    
        lblDelMoney.Caption = IIf(mstrStyle <> "", "退：" & mstrStyle, "")
        lblDelMoney.Visible = True
        cmdReturnCash.Visible = lblDelMoney.Visible And mstrStyle <> ""
        Exit Sub
    End If
    If GetThreePayDepositData(rsTemp) = False Then GoTo BrushWin
    
    '无记录时,表示不存在三方交易,直接返回true
    If rsTemp.RecordCount = 0 Then GoTo BrushWin
    rsTemp.Filter = "退预交<>0"
    If rsTemp.RecordCount = 0 Then GoTo BrushWin
    strTittle = ""
    Do While Not rsTemp.EOF
         strTittle = strTittle & IIf(strTittle = "", "", vbCrLf) & "退" & Nvl(rsTemp!卡类别名称) & ":" & Format(Val(Nvl(rsTemp!退预交)), "0.00")
         lblDelMoney.Tag = FormatEx(Val(lblDelMoney.Tag) + Val(Nvl(rsTemp!退预交)), 6)
         rsTemp.MoveNext
    Loop
    lblDelMoney.Caption = strTittle
    lblDelMoney.Visible = True
    lblDelMoney.Top = cmdReturnCash.Top + (cmdReturnCash.Height - lblDelMoney.Height) \ 2
    cmdReturnCash.Visible = lblDelMoney.Visible
BrushWin:
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CheckThreePayDepositValied(ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查三方卡缴预交的合法性
    '出参:返回支付对象(三方卡)
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-04-23 17:32:37
    '规则:
    '     1)目前只支持三方帐户中存在(转帐交易接口的)
    '     2)不能同时存在2种以上的三方帐户交易的,存在的话返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strMsg As String
    Dim dblTotal As Double, dblMoney As Double
    Set objCard = Nothing
    If mblnThreeDepositAfter Or mstrForceNote <> "" Then
        CheckThreePayDepositValied = True
        Exit Function
    End If
    mCurBrushCard.dblMoney = 0
    If GetThreePayDepositData(rsTemp) = False Then Exit Function
    '无记录时,表示不存在三方交易,直接返回true
    If rsTemp.RecordCount = 0 Then CheckThreePayDepositValied = True: Exit Function
    rsTemp.Filter = "退预交<>0"
    If rsTemp.RecordCount = 0 Then CheckThreePayDepositValied = True: Exit Function
    
    If rsTemp.RecordCount >= 2 Then
       Do While Not rsTemp.EOF
            strMsg = strMsg & vbCrLf & Nvl(rsTemp!卡类别名称) & ":" & Format(Nvl(rsTemp!退预交), "0.00")
            rsTemp.MoveNext
       Loop
       MsgBox "当前存在" & rsTemp.RecordCount & "种三方交易需要退款,目前,系统只支持一种三方交易退款(且为代扣交易的)," & _
              "" & "以下为当前需要退款的三方交易:" & _
              strMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Val(Nvl(rsTemp!是否启用)) = 0 Then
       MsgBox Nvl(rsTemp!卡类别名称) & "未启用，不允许退款!" & _
              "", vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If
    If Not GetCurCard(Val(Nvl(rsTemp!卡类别ID)), objCard) Then
       MsgBox Nvl(rsTemp!卡类别名称) & "未启用或读取失败，不允许退款!", vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If

    
    dblMoney = FormatEx(Val(Nvl(rsTemp!退预交)), 6)
    mCurBrushCard.dblMoney = dblMoney
    
    If dblMoney <> FormatEx(Val(lblDelMoney.Tag), 6) Then
       If MsgBox(Nvl(rsTemp!卡类别名称) & "中界面显未金额(" & lblDelMoney.Tag & ")与当前退款金额(" & dblMoney & ") 不一致!" & vbCrLf & "，是否重新刷新界面的退款金额!", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
           Call ShowDelThreeSwap
       End If
       Exit Function
    End If
    
    If CheckThreeSwapIsValied(objCard, dblMoney) = False Then Exit Function
    CheckThreePayDepositValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCurCard(ByVal lngCardTypeID As Long, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前卡对象
    '入参:lngCardTypeID-当前卡类别ID
    '出参:objCard-返回当前退款或缴款的卡对象
    '返回:成功,返回卡对象
    '编制:刘兴洪
    '日期:2015-04-27 10:32:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card
    On Error GoTo errHandle
    Set objCard = Nothing
    For Each objTemp In mobjPayCards
        If objTemp.接口序号 = lngCardTypeID And Not objTemp.消费卡 Then
            Set objCard = objTemp
            GetCurCard = True: Exit Function
        End If
    Next
    GetCurCard = False
    Exit Function
errHandle:
End Function

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证
    '入参:objCard-当前卡
    '     dblMoney-退卡金额
    '返回:刷卡成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 15:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, dbl帐户余额 As Double
    Dim cllSquareBalance As New Collection
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim strXmlIn As String
    
    On Error GoTo errHandle
    
    If objCard.接口序号 <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If mlng病人ID <> 0 Then
        strSql = "Select 姓名,性别,年龄 From 病人信息 Where 病人ID=[1]"
    Else
        strSql = "Select 姓名,性别,年龄 From 病人信息 A,病人结帐记录 B Where  A.病人ID=B.病人ID and B.ID=[2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng结帐ID)
    If rsTemp.EOF Then
        MsgBox "未找到指定的病人,不能调用三方接口交易", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '   zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByRef dbl金额 As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant) As Boolean
    Dim strCardNo As String, strPassWord As String
    strXmlIn = "<IN><CZLX>0</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
        objCard.接口序号, False, _
    rsTemp!姓名, Nvl(rsTemp!性别), Nvl(rsTemp!年龄), dblMoney, strCardNo, strPassWord, _
    False, True, False, False, cllSquareBalance, False, strXmlIn) = False Then Exit Function
    mCurBrushCard.str卡号 = strCardNo
    mCurBrushCard.str密码 = strPassWord
    
    '调用转帐接口
    '    7.1.    zltransferAccountsCheck(转帐检查接口)
    'zlTransferAccountsCheck 转帐检查接口
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  HIS调用模块号
    'lngCardTypeID   Long    In  卡类别ID
    'strCardNo   String  In  卡号
    'dblMoney    Double  In  转帐金额(代扣时为负数)
    'strBalanceIDs   String  In  结帐IDs，多个用逗号分离，表示本次对哪此收费项目进行重新医保补结算
    'strXMLExpend String In   XML串:
    '                            <IN>
    '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-退费业务；2-结帐业务;3-结帐退费业务
    '                            </IN>
    '                    Out  XML串:
    '                            <OUT>
    '                               <ERRMSG>错误信息</ERRMSG >
    '                            </OUT>
    '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
    '说明:
    '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '构造XML串
    If objCard.是否转帐及代扣 Then
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "2"
        zlXML.AppendNode "IN", True
        strXMLExpend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, objCard.接口序号, _
            mCurBrushCard.str卡号, dblMoney, "", strXMLExpend) = False Then
            Call ShowErrMsg(0, strXMLExpend)
            Exit Function
        End If
    End If
                    
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    Dim strExpand As String
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
          mCurBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡)
    mCurBrushCard.dbl帐户余额 = FormatEx(dbl帐户余额, 2)
    
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ExecuteThreeSwapPayInterface(objCard As Card, ByVal lng结帐ID As Long, _
      ByVal dblMoney As Double, Optional ByVal blnMustCommit As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:objCard-当前三方对象
    '     lng结帐ID-结帐ID
    '     dblMoney-本次支付金额
    '     blnMustCommit-必须提交(主要是医保接口)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-04-27 10:45:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String, strXMLExpend As String
    Dim cllPro As Collection, blnTrans As Boolean, rsTmp As ADODB.Recordset, strCardNo As String
    Dim i As Long, strSql As String, lngID As Long, varData As Variant, strExpend As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection, strInXML As String, strOutXML As String
    Dim objXml As New clsXML, dblCheck As Double, dbl冲预交 As Double, lngRow As Long, strValue As String
    dblCheck = dblMoney
    
    blnTrans = True
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    If objCard Is Nothing Then
        gcnOracle.CommitTrans
        ExecuteThreeSwapPayInterface = True: Exit Function
    End If
    
    If mblnThreeDepositAfter Or mstrForceNote <> "" Then
        gcnOracle.CommitTrans
        ExecuteThreeSwapPayInterface = True: Exit Function
    End If
        
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Then gcnOracle.CommitTrans: ExecuteThreeSwapPayInterface = True: Exit Function
    If objCard.是否转帐及代扣 Then
        'zlTransferAccountsMoney
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'strBalanceID    String  In  结算ID
        'dblMoney    Double  In  转帐金额
        'strSwapGlideNO  String  Out 交易流水号
        'strSwapMemo String  Out 交易说明
        'strSwapExtendInfor  String  Out 交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-退费业务；2-结帐业务;3-结帐退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    True:调用成功,False:调用失败
        '说明:
        '１. 在医保补充结算时进行的三方转帐时调用。
        '２. 一般来说，成功转帐后，都应该打印相关的结算票据，可以放在此接口进行处理.
        '３. 在转帐成功后，返回交易流水号和相关交易说明；如果存在其他交易信息，可以放在扩展信息中返回.
        '构造XML串
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "2"
        zlXML.AppendNode "IN", True
        strXMLExpend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, objCard.接口序号, mCurBrushCard.str卡号, _
            lng结帐ID, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
            If Not blnMustCommit Then   '医保必须提交，必须根据病人预交记录中的校对标志来确定
                gcnOracle.RollbackTrans:
            Else
                gcnOracle.CommitTrans
                ExecuteThreeSwapPayInterface = True
            End If
            Call ShowErrMsg(1, strXMLExpend)
            blnTrans = False
            Exit Function
        End If
        
        mCurBrushCard.str交易流水号 = strSwapGlideNO
        mCurBrushCard.str交易说明 = strSwapMemo
        Call zlAddUpdateSwapSQL(False, lng结帐ID, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
        Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    Else
        objXml.ClearXmlText
        
        With vsDeposit
            Call objXml.AppendNode("JSLIST")
            For i = .Rows - 1 To 1 Step -1
                '卡类别ID||代扣标志
                varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||", "||")
                If Val(varData(0)) <> 0 And Val(varData(2)) = 0 And dblCheck > 0 Then
                    dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                    If dblCheck >= dbl冲预交 Then
                        lngID = .TextMatrix(i, .ColIndex("预交ID"))
                        strSql = "Select ID,卡号,交易流水号,交易说明 From 病人预交记录 Where ID = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
                        If Not rsTmp.EOF Then
                            Call objXml.AppendNode("JS")
                                Call objXml.appendData("KH", Nvl(rsTmp!卡号))
                                Call objXml.appendData("JYLSH", Nvl(rsTmp!交易流水号))
                                Call objXml.appendData("JYSM", Nvl(rsTmp!交易说明))
                                Call objXml.appendData("ZFJE", dbl冲预交)
                                Call objXml.appendData("JSLX", 1)
                                Call objXml.appendData("ID", Nvl(rsTmp!ID))
                            Call objXml.AppendNode("JS", True)
                        End If
                        strSql = "Zl_三方退款信息_Insert("
                        strSql = strSql & lng结帐ID & ","
                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                        strSql = strSql & dbl冲预交 & ",'"
                        strSql = strSql & Nvl(rsTmp!卡号) & "','"
                        strSql = strSql & Nvl(rsTmp!交易流水号) & "','"
                        strSql = strSql & Nvl(rsTmp!交易说明) & "')"
                        zlAddArray cllThreeSwap, strSql
                        dblCheck = dblCheck - dbl冲预交
                    Else
                        lngID = .TextMatrix(i, .ColIndex("预交ID"))
                        strSql = "Select ID,卡号,交易流水号,交易说明 From 病人预交记录 Where ID = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
                        If Not rsTmp.EOF Then
                            Call objXml.AppendNode("JS")
                                Call objXml.appendData("KH", Nvl(rsTmp!卡号))
                                Call objXml.appendData("JYLSH", Nvl(rsTmp!交易流水号))
                                Call objXml.appendData("JYSM", Nvl(rsTmp!交易说明))
                                Call objXml.appendData("ZFJE", dblCheck)
                                Call objXml.appendData("JSLX", 1)
                                Call objXml.appendData("ID", Nvl(rsTmp!ID))
                            Call objXml.AppendNode("JS", True)
                        End If
                        strSql = "Zl_三方退款信息_Insert("
                        strSql = strSql & lng结帐ID & ","
                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                        strSql = strSql & dblCheck & ",'"
                        strSql = strSql & Nvl(rsTmp!卡号) & "','"
                        strSql = strSql & Nvl(rsTmp!交易流水号) & "','"
                        strSql = strSql & Nvl(rsTmp!交易说明) & "')"
                        zlAddArray cllThreeSwap, strSql
                        dblCheck = 0
                    End If
                End If
            Next i
            Call objXml.AppendNode("JSLIST", True)
        End With
    
        strInXML = objXml.XmlText
        
        If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, objCard.接口序号, objCard.消费卡, strInXML, _
             lng结帐ID, strOutXML, strExpend) = False Then
            If Not blnMustCommit Then   '医保必须提交，必须根据病人预交记录中的校对标志来确定
                gcnOracle.RollbackTrans:
            Else
                gcnOracle.CommitTrans
                ExecuteThreeSwapPayInterface = True
            End If
            Call ShowErrMsg(1, strXMLExpend)
            blnTrans = False
            Exit Function
        End If
             
        If strOutXML <> "" Then
            If zlXML_Init = False Then Exit Function
            If zlXML_LoadXMLToDOMDocument(strOutXML, False) = False Then Exit Function
            Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("ID", i, strValue)
                strSql = "Zl_三方退款信息_Insert("
                strSql = strSql & lng结帐ID & ","
                strSql = strSql & Val(strValue) & ","
                strSql = strSql & 0 & ",'"
                Call zlXML_GetNodeValue("KH", i, strValue)
                strSql = strSql & strValue & "','"
                Call zlXML_GetNodeValue("TKLSH", i, strValue)
                strSql = strSql & strValue & "','"
                Call zlXML_GetNodeValue("TKSM", i, strValue)
                strSql = strSql & strValue & "',"
                strSql = strSql & 1 & ")"
                zlAddArray cllThreeSwap, strSql
            Next i
        End If
        
        If strExpend <> "" Then
            strSwapExtendInfor = ""
            If zlXML_LoadXMLToDOMDocument(strExpend, False) = False Then Exit Function
            Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("XMMC", i, strValue)
                strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
                Call zlXML_GetNodeValue("XMNR", i, strValue)
                strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
            Next i
        End If
        If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
        strSql = "Select 卡号 From 病人预交记录 Where 结帐ID= [1] And 卡类别ID= [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID, objCard.接口序号)
        If Not rsTmp.EOF Then
            strCardNo = Nvl(rsTmp!卡号)
        End If
        Call zlAddUpdateSwapSQL(False, lng结帐ID, objCard.接口序号, objCard.消费卡, strCardNo, "", "", cllUpdate, 0)
        Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, objCard.接口序号, objCard.消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    End If

    Err = 0: On Error GoTo ErrOtherHand:
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    ExecuteThreeSwapPayInterface = True
    Exit Function
ErrOtherHand:
    ExecuteThreeSwapPayInterface = True
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ShowErrMsg(ByVal BytType As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方转账检查与代扣业务出错提示
    '编制:冉俊明
    '时间:2014-12-2
    '参数:
    '   bytType:0-转账检查,1-转账交易
    '   strXMLErrMsg:格式如下
    '            <OUT>
    '               <ERRMSG>错误信息</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    '解析错误信息
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '提示错误信息
    If Trim(strValue) = "" Then
        If BytType = 0 Then
            strValue = vbCrLf & "交易检查失败！"
        Else
            strValue = vbCrLf & "交易失败！"
        End If
    End If
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
     
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then
        Call CreateSquareCardObject(Me, mlngModul)
        If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    End If
    Set mrsCardType = gobjSquare.objSquareCard.zlGetYLCards
    '所有启用的三方帐户
    Set mobjPayCards = gobjSquare.objSquareCard.zlGetCards(3)
End Sub
 

Private Function GetDelThreeCardDepositInfor(ByRef intThreeCount As Integer, ByRef intNotDelCashCount As Integer, _
    ByRef blnThreeDepositAfter As Boolean, ByRef strDelThreeNames As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退三帐户的预交相关信息
    '入参:
    '出参:intNotDelCashCount-返回不允许退现的个数
    '     intThreeCount-三方帐户个数
    '     blnThreeDepositAfter-三方帐户余额存退(true:存在余额退款,False-不存在余额退款)
    '     strDelThreeNames-发生三方帐户余额退款的名称串，比如：招行,建行
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-10-25 11:59:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double, varData As Variant, dbl冲预交 As Double
    Dim strStyle As String, i As Long
    
    On Error GoTo errHandle
    
    blnThreeDepositAfter = False
    
    dblTotal = RoundEx(Val(txtTotal.Text) - mcurMediCare, 2)
    If mrsCardType Is Nothing Then
        Call initCardSquareData
    ElseIf mrsCardType.State <> 1 Then
        Call initCardSquareData
    End If
    
    intNotDelCashCount = 0
    intThreeCount = 0
    With vsDeposit
        For i = 1 To .Rows - 1
            dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
            ' 卡类别ID ||代扣||退现||缺省退现
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||||", "||")
            If Val(varData(0)) <> 0 Then
                If Val(varData(3)) = 0 And ((dblTotal - dbl冲预交) <= 0 Or dbl冲预交 = 0) Then   '非缺省退现
                    mrsCardType.Filter = "ID=" & Val(varData(0))
                    intThreeCount = intThreeCount + 1
                    If Not mrsCardType.EOF Then
                        If InStr(strStyle & ",", "," & Nvl(mrsCardType!名称) & ",") = 0 Then
                            strStyle = strStyle & "," & mrsCardType!名称
                            If Val(varData(2)) = 0 Then
                               intNotDelCashCount = intNotDelCashCount + 1
                            End If
                        End If
                    End If
                End If
            End If
            
            If FormatEx(dblTotal - dbl冲预交, 6) <= 0 Then
                dblTotal = 0
            Else
                dblTotal = FormatEx(dblTotal - dbl冲预交, 6)
            End If
        Next
    End With
    
    
    If intThreeCount >= 1 And InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then blnThreeDepositAfter = True
    
    If strStyle <> "" Then strStyle = Mid(strStyle, 2)
    strDelThreeNames = strStyle

    GetDelThreeCardDepositInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


