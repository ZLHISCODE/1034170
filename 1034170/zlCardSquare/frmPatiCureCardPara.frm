VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPatiCureCardPara 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7545
      TabIndex        =   4
      Top             =   5085
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   5085
      Width           =   1100
   End
   Begin TabDlg.SSTab stbPage 
      Height          =   4665
      Left            =   135
      TabIndex        =   0
      Top             =   255
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   8229
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "常规(&0)"
      TabPicture(0)   =   "frmPatiCureCardPara.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra退卡方式"
      Tab(0).Control(1)=   "chkLedWelcome"
      Tab(0).Control(2)=   "chk记帐"
      Tab(0).Control(3)=   "fraShortLine"
      Tab(0).Control(4)=   "txtNameDays"
      Tab(0).Control(5)=   "cmdDeviceSetup"
      Tab(0).Control(6)=   "chkSeekName"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "医疗卡票据(&1)"
      TabPicture(1)   =   "frmPatiCureCardPara.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrint"
      Tab(1).Control(1)=   "fraTitle"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "预交款票据(&2)"
      TabPicture(2)   =   "frmPatiCureCardPara.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraPrepay"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fra票据格式"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fra票据格式 
         Caption         =   "预交票据格式"
         Height          =   1305
         Left            =   150
         TabIndex        =   25
         Top             =   3045
         Width           =   8190
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1005
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   7995
            _cx             =   14102
            _cy             =   1773
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
            FormatString    =   $"frmPatiCureCardPara.frx":0054
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
      Begin VB.Frame fraPrint 
         Height          =   615
         Left            =   -74910
         TabIndex        =   19
         Top             =   3900
         Width           =   8145
         Begin VB.OptionButton optPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   2670
            TabIndex        =   23
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   3810
            TabIndex        =   22
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   1665
            TabIndex        =   21
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置(&S)"
            Height          =   345
            Left            =   6570
            TabIndex        =   20
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label lblPrint 
            Caption         =   "发卡打印方式"
            Height          =   285
            Left            =   435
            TabIndex        =   24
            Top             =   270
            Width           =   1500
         End
      End
      Begin VB.Frame fra退卡方式 
         Caption         =   "退卡方式设置"
         Height          =   1050
         Left            =   -74280
         TabIndex        =   14
         Top             =   2235
         Width           =   7320
         Begin VB.OptionButton optBrush 
            Caption         =   "输入单据号退卡或刷卡退卡"
            Height          =   180
            Index           =   3
            Left            =   255
            TabIndex        =   18
            Top             =   765
            Width           =   2460
         End
         Begin VB.OptionButton optBrush 
            Caption         =   "输入单据号后才刷卡退卡"
            Height          =   180
            Index           =   2
            Left            =   3600
            TabIndex        =   17
            Top             =   420
            Width           =   2460
         End
         Begin VB.OptionButton optBrush 
            Caption         =   "必须刷卡退卡"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   16
            Top             =   420
            Width           =   1740
         End
         Begin VB.OptionButton optBrush 
            Caption         =   "不进行刷卡验证"
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   15
            Top             =   420
            Value           =   -1  'True
            Width           =   1740
         End
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   -74265
         TabIndex        =   13
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   1785
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.CheckBox chk记帐 
         Caption         =   "就诊卡费用以记账方式收取"
         Height          =   180
         Left            =   -74265
         TabIndex        =   12
         Top             =   930
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Frame fraShortLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   -71655
         TabIndex        =   10
         Top             =   1515
         Width           =   285
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
         Left            =   -71655
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "0表示查找时不限制时间"
         Top             =   1335
         Width           =   285
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "本地共用预交票据"
         Height          =   2310
         Left            =   150
         TabIndex        =   7
         Top             =   615
         Width           =   8190
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   1920
            Left            =   75
            TabIndex        =   8
            Top             =   270
            Width           =   8025
            _cx             =   14155
            _cy             =   3387
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
            FormatString    =   $"frmPatiCureCardPara.frx":00E2
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
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   -68205
         TabIndex        =   6
         Top             =   3510
         Width           =   1500
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用收费票据"
         Height          =   3345
         Left            =   -74910
         TabIndex        =   1
         Top             =   540
         Width           =   8190
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2925
            Left            =   75
            TabIndex        =   2
            Top             =   270
            Width           =   7995
            _cx             =   14102
            _cy             =   5159
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
            FormatString    =   $"frmPatiCureCardPara.frx":01C2
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
      Begin VB.CheckBox chkSeekName 
         Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
         Height          =   195
         Left            =   -74265
         TabIndex        =   11
         Top             =   1350
         Width           =   4260
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6375
      TabIndex        =   5
      Top             =   5085
      Width           =   1100
   End
End
Attribute VB_Name = "frmPatiCureCardPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnOk As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '返回:参数设置成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-07-14 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOk = False
    
    Me.Show 1, frmMain
    zlSetPara = mblnOk
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
    isValied = False
    On Error GoTo errHandle
    '检查每种使用种式只能一个选择
    With vsBill
        str类别 = "-"
        For i = 1 To vsBill.Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("医疗卡类别"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("医疗卡类别")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("医疗卡类别"))) = Trim(.TextMatrix(j, .ColIndex("医疗卡类别"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    医疗卡类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
  '检查每种使用预交只能一个选择
    With vsPrepay
        str类别 = "-"
        For i = 1 To .Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("预交类型"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("预交类型")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("预交类型"))) = Trim(.TextMatrix(j, .ColIndex("预交类型"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    预交类型为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关票据
    '编制:刘兴洪
    '日期:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String, strPrintMode As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '保存共享票据
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("医疗卡类别")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用医疗卡批次", strValue, glngSys, mlngModule, blnHavePrivs
    '保存预交票据
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("预交类型")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用预交票据批次", strValue, glngSys, mlngModule, blnHavePrivs
    
    '保存预交格式
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("预交打印方式")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "预交发票格式", strValue, glngSys, mlngModule, blnHavePrivs
        zlDatabase.SetPara "预交发票打印方式", strPrintMode, glngSys, mlngModule, blnHavePrivs
    End With
End Sub
Private Sub InitShareInvoice()
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String, rs医疗卡类别 As ADODB.Recordset
    Dim strPrintMode As String, strBillFormat As String, blnHavePrivs As Boolean
    Dim str缺省医疗卡 As String, lng缺省医疗卡 As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ID,编码,名称, nvl(是否固定,0) as 是否固定  from 医疗卡类别  "
    Set rs医疗卡类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rs医疗卡类别.Filter = "名称='就诊卡' and 是否固定=1"
    If rs医疗卡类别.EOF = False Then
        str缺省医疗卡 = rs医疗卡类别!名称: lng缺省医疗卡 = Val(rs医疗卡类别!id)
    End If
    
    zl_vsGrid_Para_Restore mlngModule, vsBill, Me.Name, "共用医疗票据列表", False, False
    strShareInvoice = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModule, , , True, intType)
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
    
    '格式:领用ID1,医疗卡类别ID1|领用IDn,医疗卡类别IDn|...
    varData = Split(strShareInvoice, "|")

    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(5)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            '99007:李南春,2016/7/29，共用医疗卡票据获取使用类别ID
            If Val(Nvl(rsTemp!使用类别ID)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = str缺省医疗卡
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = lng缺省医疗卡
            Else
                rs医疗卡类别.Filter = "ID=" & Val(Nvl(rsTemp!使用类别ID))
                If Not rs医疗卡类别.EOF Then
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = Nvl(rs医疗卡类别!名称)
                Else
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = Nvl(rsTemp!使用类别)
                End If
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = Val(Nvl(rsTemp!使用类别ID))
            End If
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("医疗卡类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    '共用预交票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModule, , , True, intType)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,预交类别ID1|领用IDn,预交类别IDn|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            If Val(Nvl(rsTemp!使用类别, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "门诊和住院共用"
            ElseIf Val(Nvl(rsTemp!使用类别, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交门诊票据"
            Else
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交住院票据"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = Val(Nvl(rsTemp!使用类别))
            
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("预交类型"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    '票据格式处理
    Dim strReport As String
    
    zl_vsGrid_Para_Restore mlngModule, vsBillFormat, Me.Name, "预交发票打印方式", False, False
    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlReadBillFormat(strReport)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("预交打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    
    '读取参数值
    strBillFormat = zlDatabase.GetPara("预交发票格式", glngSys, mlngModule, , , True, intType)
    strPrintMode = zlDatabase.GetPara("预交发票打印方式", glngSys, mlngModule, , , True, intType1)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat
        .TextMatrix(1, 0) = "门诊预交"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "住院预交"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("票据格式")) = "0"
        .ColData(.ColIndex("预交打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("预交打印方式")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("票据格式"))) = 1 Or _
            Val(.ColData(.ColIndex("预交打印方式"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strBillFormat, "|")
    VarType = Split(strPrintMode, "|")
    
    With vsBillFormat
        .Clear 1
        .Rows = 3
        For lngRow = 1 To .Cols - 1
            .TextMatrix(lngRow, .ColIndex("预交打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("预交打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
        Next
        If Val(.ColData(.ColIndex("预交打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("预交打印方式"), .Rows - 1, .ColIndex("预交打印方式")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("票据格式"), .Rows - 1, .ColIndex("票据格式")) = vbBlue
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Sub cmdOK_Click()
    Dim blnHavePrivs As Boolean, intData As Integer
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    If isValied = False Then Exit Sub
    zlDatabase.SetPara "姓名模糊查找", chkSeekName.value, glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "姓名查找天数", Val(txtNameDays.Text), glngSys, mlngModule, blnHavePrivs
   
    zlDatabase.SetPara "卡费记帐", chk记帐.value, glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.value, glngSys, mlngModule, blnHavePrivs
    '问题28130、27929
    intData = 0
    If optBrush(3).value = True Then
        intData = 3
    ElseIf optBrush(1).value = True Then
        intData = 1
    ElseIf optBrush(2).value = True Then
        intData = 2
    End If
    Call zlDatabase.SetPara("退卡刷卡", intData, glngSys, mlngModule, blnHavePrivs)
    zlDatabase.SetPara "发卡打印方式", IIf(optPrint(0).value, 0, IIf(optPrint(1).value, 1, 2)), glngSys, mlngModule, blnHavePrivs
    Call SaveInvoice
    mblnOk = True: Unload Me
End Sub
Private Sub InitPara()
    Dim blnHavePrivs As Boolean, i As Long
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    txtNameDays.Enabled = True
    txtNameDays.Text = zlDatabase.GetPara("姓名查找天数", glngSys, mlngModule, , Array(txtNameDays), blnHavePrivs)
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.value = IIf(zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModule, , Array(chkSeekName), blnHavePrivs) = "1", 1, 0)
    chk记帐.value = IIf(zlDatabase.GetPara("卡费记帐", glngSys, glngModul, , Array(chk记帐), blnHavePrivs) = "1", 1, 0)
    'LED设备
    chkLedWelcome.value = zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModule, 1, Array(chkLedWelcome), blnHavePrivs)
    '问题28130
    Select Case Val(zlDatabase.GetPara("退卡刷卡", glngSys, mlngModule, "0", Array(fra退卡方式, optBrush(0), optBrush(1), optBrush(2), optBrush(3)), InStr(mstrPrivs, "参数设置") > 0))
    Case 0
        optBrush(0).value = True
    Case 1
        optBrush(1).value = True
    Case "2"
        optBrush(2).value = True
    Case "3"
        optBrush(3).value = True
    End Select
    
    i = Val(zlDatabase.GetPara("发卡打印方式", glngSys, mlngModule, , Array(optPrint(0), optPrint(1), optPrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optPrint(i).value = True
      
End Sub
Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.value = 1 And txtNameDays.Tag = "1"
End Sub

Private Sub cmdPrepayPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdPrintSet_Click()
    '打印设置
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me)
End Sub

Private Sub Form_Load()
    Call InitShareInvoice
    Call InitPara
    chkSeekName_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用医疗票据列表", False, False
    zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用预交票据列表", False, False
End Sub
 
Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub
Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Val(.Cell(flexcpData, Row, .ColIndex("医疗卡类别"))) = Val(.Cell(flexcpData, i, .ColIndex("医疗卡类别"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
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
Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("预交类型"))) = Trim(.Cell(flexcpData, i, .ColIndex("预交类型"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
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



