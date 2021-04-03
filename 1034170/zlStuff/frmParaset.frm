VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParaset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmParaset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabMain 
      Height          =   7335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frmParaset.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra移库流程控制"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra排序"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra库房选择"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraBidMess"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "资质校验(&1)"
      TabPicture(1)   =   "frmParaset.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkProduceDate"
      Tab(1).Control(1)=   "fraCheck"
      Tab(1).Control(2)=   "vsfCheck"
      Tab(1).Control(3)=   "lblComment"
      Tab(1).ControlCount=   4
      Begin VB.CheckBox chkProduceDate 
         Caption         =   "生产日期大于注册证效期检查"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Frame fraBidMess 
         Caption         =   "入库单价超中标成本价"
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   4860
         Width           =   3675
         Begin VB.OptionButton optBidMess 
            Caption         =   "禁止"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optBidMess 
            Caption         =   "提示"
            Height          =   180
            Index           =   1
            Left            =   1200
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optBidMess 
            Caption         =   "不限制"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame fraCheck 
         Caption         =   "选择校验方式"
         Height          =   615
         Left            =   -74760
         TabIndex        =   30
         Top             =   5520
         Width           =   7095
         Begin VB.OptionButton optCheck 
            Caption         =   "校验未通过时提醒"
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   32
            Top             =   280
            Width           =   2175
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "校验未通过时禁止保存"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   31
            Top             =   280
            Width           =   2175
         End
      End
      Begin VB.Frame fra库房选择 
         Caption         =   "库房选择"
         Height          =   1665
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   3675
         Begin VB.CheckBox chkStock 
            Caption         =   "允许选择库房"
            Height          =   375
            Left            =   210
            TabIndex        =   27
            Top             =   240
            Width           =   2805
         End
         Begin VB.Label lbl库房选择说明 
            Caption         =   "    如果选择库房，则在单据中有'所有库房'权限人就可以选择不同库房；否则，不能选择库房。"
            Height          =   615
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   3285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "卫材单位"
         Height          =   1665
         Left            =   3840
         TabIndex        =   20
         Top             =   480
         Width           =   3675
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "    请选择一种卫生材料单位，在单据输入中，所有卫生材料将用这种单位。"
            Height          =   405
            Left            =   240
            TabIndex        =   25
            Top             =   1170
            Width           =   3315
         End
         Begin VB.Label lbl盘点表 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "盘点表"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   24
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl盘点单 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "盘点单"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   23
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "控制"
         Height          =   3345
         Left            =   3840
         TabIndex        =   11
         Top             =   2250
         Width           =   3675
         Begin VB.CheckBox chk加成入库 
            Caption         =   "时价卫材以加成率入库"
            Height          =   255
            Left            =   390
            TabIndex        =   46
            Top             =   2040
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk取上次售价 
            Caption         =   "时价卫材入库时取上次售价"
            Height          =   255
            Left            =   390
            TabIndex        =   45
            Top             =   2280
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk分段加成入库 
            Caption         =   "时价卫材按分段加成入库"
            Height          =   255
            Left            =   390
            TabIndex        =   44
            Top             =   2520
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk时价调价 
            Caption         =   "时价卫材按批次调价"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   390
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.CheckBox chkSet库房 
            Caption         =   "允许盘点没有设置存储库房的卫材"
            Height          =   255
            Left            =   390
            TabIndex        =   39
            Top             =   1230
            Visible         =   0   'False
            Width           =   3105
         End
         Begin VB.CheckBox chk高值卫材录入 
            Caption         =   "高值卫材必须填写详细信息"
            Height          =   255
            Left            =   390
            TabIndex        =   13
            Top             =   1755
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk领药财务审核 
            Caption         =   "在领用审核前需要进行财务核查"
            Height          =   255
            Left            =   390
            TabIndex        =   33
            Top             =   1235
            Visible         =   0   'False
            Width           =   3180
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "单据存盘后自动打印"
            Height          =   255
            Left            =   390
            TabIndex        =   19
            Top             =   455
            Width           =   1935
         End
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "单据审核后自动打印"
            Height          =   255
            Left            =   390
            TabIndex        =   18
            Top             =   715
            Width           =   1935
         End
         Begin VB.CheckBox chk修改单据号 
            Caption         =   "允许修改单据号"
            Height          =   255
            Left            =   390
            TabIndex        =   17
            Top             =   972
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.CheckBox chkFixPrice 
            Caption         =   "外购入库定价采购"
            Height          =   255
            Left            =   390
            TabIndex        =   16
            Top             =   195
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.CheckBox chk允许修改批发价 
            Caption         =   "允许修改批发价"
            Height          =   255
            Left            =   390
            TabIndex        =   15
            Top             =   1235
            Visible         =   0   'False
            Width           =   2700
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "单据打印设置(&S)"
            Height          =   350
            Left            =   360
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2880
            Width           =   2925
         End
         Begin VB.CheckBox chk招标卫材 
            Caption         =   "招标卫材可选择非中标单位入库"
            Height          =   255
            Left            =   390
            TabIndex        =   12
            Top             =   1495
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk跟踪在用 
            Caption         =   "不允许具有""跟踪在用""属性的卫材进行领用"
            Height          =   360
            Left            =   390
            TabIndex        =   38
            Top             =   1550
            Width           =   2850
         End
      End
      Begin VB.Frame fra排序 
         Caption         =   "排序方式"
         Height          =   2505
         Left            =   120
         TabIndex        =   7
         Top             =   2250
         Width           =   3675
         Begin VB.ComboBox cbo列名 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox cbo方向 
            Height          =   300
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   390
            Width           =   885
         End
         Begin VB.Label lbl排序说明 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "    本参数的设置，将影响所有编辑窗体中单据的显示内容的排序方式。缺省：按用户输入的顺序显示各单据的内容"
            ForeColor       =   &H80000008&
            Height          =   825
            Left            =   180
            TabIndex        =   10
            Top             =   930
            Width           =   3345
         End
      End
      Begin VB.Frame fra移库流程控制 
         Caption         =   "移库流程控制"
         Height          =   1275
         Left            =   120
         TabIndex        =   4
         Top             =   5880
         Width           =   7365
         Begin VB.CheckBox chkRequestStrike 
            Caption         =   "移库冲销时，移入库房需要先申请冲销"
            Height          =   180
            Left            =   180
            TabIndex        =   40
            Top             =   960
            Value           =   1  'Checked
            Width           =   3705
         End
         Begin VB.CheckBox chk移库流程控制 
            Caption         =   "移库时需要备料、发送、接收这一过程。"
            Height          =   180
            Left            =   180
            TabIndex        =   5
            Top             =   270
            Value           =   1  'Checked
            Width           =   6945
         End
         Begin VB.Label lbl移库说明 
            Caption         =   "注意：如果不打勾，那么在填写移库单后，增加一个审核操作，审核后自动完成备料、发送、接收这一过程。审核前可以修改单据。"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   6945
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
         Height          =   4125
         Left            =   -74760
         TabIndex        =   42
         Top             =   960
         Width           =   7095
         _cx             =   12515
         _cy             =   7276
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
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   25
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParaset.frx":0044
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
         VirtualData     =   0   'False
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
      Begin VB.Label lblComment 
         Caption         =   $"frmParaset.frx":02EF
         Height          =   540
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   7140
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   360
      TabIndex        =   2
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5160
      TabIndex        =   0
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   1
      Top             =   7680
      Width           =   1100
   End
End
Attribute VB_Name = "frmParaset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mlngModule As Long '
Private mstrPrivs As String '
Private mblnHavePriv As Boolean
Private mblnFirstLoad As Boolean    '记录是否第一次加载

Private Sub chk分段加成入库_Click()
    If chk分段加成入库.Value = 1 Then
        chk加成入库.Value = 0
        chk取上次售价.Value = 0
    End If
End Sub

Private Sub chk加成入库_Click()
    If chk加成入库.Value = 1 Then
        chk取上次售价.Value = 0
        chk分段加成入库.Value = 0
    End If
End Sub

Private Sub chk取上次售价_Click()
    If chk取上次售价.Value = 1 Then
        chk加成入库.Value = 0
        chk分段加成入库.Value = 0
    End If
End Sub
Private Function ISValid() As Boolean
    Dim i As Integer
    Dim blnAllUnCheck As Boolean
    
    '资质校验
    If tabMain.TabVisible(1) = True Then
        blnAllUnCheck = True
        With vsfCheck
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("校验")) <> "" Then
                    blnAllUnCheck = False
                    Exit For
                End If
            Next
        End With
        
        '如果选择了校验项目，则必须选择校验方式
        If blnAllUnCheck = False And optCheck(0).Value = 0 And optCheck(1).Value = 0 Then
            MsgBox "请选择资质校验方式！", vbExclamation, gstrSysName
            tabMain.Tab = 1
            If vsfCheck.Enabled Then vsfCheck.SetFocus
            Exit Function
        End If
    End If
    
    ISValid = True
End Function

Private Sub Save资质校验()
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean
    
    If mstrFunction = "卫材外购入库管理" Or mstrFunction = "卫材计划管理" Then
        blnAllUnCheck = True
        
        '保存资质校验项目和方式，格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
        With vsfCheck
            For i = 1 To .Rows - 1
                strCheck = IIf(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("类别")) & "," & .TextMatrix(i, .ColIndex("校验项目")) & "," & _
                    IIf(.TextMatrix(i, .ColIndex("校验")) = "", 0, 1)
                    
                If .TextMatrix(i, .ColIndex("校验")) <> "" Then blnAllUnCheck = False
            Next
        End With
        
        If blnAllUnCheck = True Then
            strCheck = "0|" & strCheck
        ElseIf optCheck(0).Value = True Then
            strCheck = "2|" & strCheck
        Else
            strCheck = "1|" & strCheck
        End If
            
        Call zlDatabase.SetPara("资质校验", strCheck, glngSys, mlngModule)
    End If
    
End Sub

Private Sub Load资质校验()
    Dim i As Integer
    Dim n As Integer
    Dim strCheck As String
    Dim intCheckType As Integer
    Dim arrColumn
    
    On Error Resume Next
    
    If mstrFunction = "卫材外购入库管理" Or mstrFunction = "卫材计划管理" Then
        '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
        strCheck = zlDatabase.GetPara("资质校验", glngSys, mlngModule, "", Array(vsfCheck, fraCheck), mblnHavePriv)
        
        If strCheck <> "" Then
            If mstrFunction = "卫材外购入库管理" Then
                chkProduceDate.Value = IIf(Val(zlDatabase.GetPara("生产日期效期检查", glngSys, mlngModule, "0", Array(chkProduceDate), mblnHavePriv)) = 1, 1, 0)
            End If
            
            If InStr(1, strCheck, "|") > 0 Then
                '校验方式：0-不检查；1－提醒；2－禁止
                intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
                If intCheckType = 2 Then
                    optCheck(0).Value = True
                ElseIf intCheckType = 1 Then
                    optCheck(1).Value = True
                End If
                
                strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)
                 
                If strCheck <> "" Then
                    strCheck = strCheck & ";"
                    arrColumn = Split(strCheck, ";")
                    For n = 0 To UBound(arrColumn)
                        If arrColumn(n) <> "" Then
                            With vsfCheck
                                For i = 1 To .Rows - 1
                                    If Split(arrColumn(n), ",")(0) = .TextMatrix(i, .ColIndex("类别")) And Split(arrColumn(n), ",")(1) = .TextMatrix(i, .ColIndex("校验项目")) Then
                                        If Val(Split(arrColumn(n), ",")(2)) = 1 Then
                                            .TextMatrix(i, .ColIndex("校验")) = "√"
                                        End If
                                    End If
                                Next
                            End With
                        End If
                    Next
                End If
            End If
        End If
    End If
End Sub
Private Sub Cbo列名_Click()
    If cbo方向.ListCount < 1 Then Exit Sub
    cbo方向.Enabled = Not (cbo列名.ListIndex = 0)
    If Not cbo方向.Enabled Then cbo方向.ListIndex = 0
End Sub
Private Sub chk单据累加_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkRequestStrike_Click()
    '当变为不需要申请时，要检查是否有未审核的冲销申请单，如果有则不能改变
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If mblnFirstLoad = True Then
        If chkRequestStrike.Value = 0 Then
            If MsgBox("即将检查是否存在未审核的冲销申请单，可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '该功能是10.34版本新增，增加一个条件填制日期范围，避免全表扫描，因此考虑从34版本修改日期开始
                gstrSQL = "Select 1 From 药品收发记录 Where 单据 = 19 And Mod(记录状态, 3) = 2 And 审核日期 Is Null " & _
                    " And 填制日期 Between To_Date('2014/2/20 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Rownum = 1"
                
                DoEvents
                zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的冲销申请单")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "存在未审核的冲销申请单，不能改变此参数！", vbInformation, gstrSysName
                    chkRequestStrike.Value = 1
                End If
            Else
                chkRequestStrike.Value = 1
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Sub SetCtlEnabled()
    '-----------------------------------------------------------------------------
    '功能:权限设置
    '-----------------------------------------------------------------------------
    Dim blnPara As Boolean
    blnPara = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    chkFixPrice.Enabled = blnPara
    chk修改单据号.Enabled = blnPara
    chk允许修改批发价.Enabled = blnPara
    chk招标卫材.Enabled = blnPara
    chk移库流程控制.Enabled = blnPara
    chk高值卫材录入.Enabled = blnPara
    chk跟踪在用.Enabled = blnPara
    
    If tabMain.TabVisible(1) = True Then
        vsfCheck.Enabled = blnPara
        fraCheck.Enabled = blnPara
    End If
    If mlngModule = 1726 Then
        fra库房选择.Enabled = False
    End If
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/24
    '------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
    
    Call zlDatabase.SetPara("入库单价超中标单价", IIf(optBidMess(0).Value, 0, IIf(optBidMess(1).Value, 1, 2)), glngSys, mlngModule, IIf(fraBidMess.Enabled, True, False))
    Call zlDatabase.SetPara("单据排序", CStr(cbo列名.ListIndex) & CStr(cbo方向.ListIndex), glngSys, mlngModule)   '
    Call zlDatabase.SetPara("存盘打印", IIf(chkSavePrint.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("审核打印", IIf(chkVerifyPrint.Value = 1, 1, 0), glngSys, mlngModule)
    
    If chk高值卫材录入.Visible Then
        Call zlDatabase.SetPara("高值卫材必须填写详细信息", IIf(chk高值卫材录入.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    If mlngModule = 1725 Then
        Call zlDatabase.SetPara("是否选择部门", IIf(chkStock.Value = 1, 1, 0), glngSys, mlngModule)
    Else
        Call zlDatabase.SetPara("是否选择库房", IIf(chkStock.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    Call zlDatabase.SetPara(IIf(mlngModule = 1719, "盘点表单位", "卫材单位"), cboUnit.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("修改单据号", IIf(chk修改单据号.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("跟踪在用", IIf(chk跟踪在用.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("存储库房", IIf(chkSet库房.Value = 1, 1, 0), glngSys, mlngModule)
    
    If chkFixPrice.Visible = True Then
        Select Case mstrFunction
            Case "卫材外购入库管理"
                Call zlDatabase.SetPara("定价采购", IIf(chkFixPrice.Value = 1, 1, 0), glngSys, mlngModule)
                Call zlDatabase.SetPara("修改采购限价", IIf(chk允许修改批发价.Value = 1, 1, 0), glngSys, mlngModule)
                Call zlDatabase.SetPara("招标卫材可选择非中标单位入库", IIf(chk招标卫材.Value = 1, 1, 0), glngSys, mlngModule)
'                Call zlDatabase.SetPara("校验供应商资质", IIf(chk供应商校验.Value = 1, 1, 0), glngSys, mlngModule)
            Case ""
        End Select
    End If
    If CboUnit1.Visible Then
        Call zlDatabase.SetPara("记录单单位", CboUnit1.ListIndex, glngSys, mlngModule)
    End If
    If fra移库流程控制.Visible = True Then
        Call zlDatabase.SetPara("移库流程", IIf(chk移库流程控制.Value = 1, 1, 0), glngSys, mlngModule)
        Call zlDatabase.SetPara("冲销申请", IIf(chkRequestStrike.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    If mlngModule = 1717 Then
        Call zlDatabase.SetPara("审核流程", IIf(chk领药财务审核.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    '卫材调价管理
    If mlngModule = 1726 Then
        zlDatabase.SetPara "时价卫材按批次调价", chk时价调价.Value, glngSys, mlngModule
    End If
    
    Save资质校验
    If mlngModule = 1712 Then
        '卫材外购入库才有生产日期效期检查
        zlDatabase.SetPara "生产日期效期检查", chkProduceDate.Value, glngSys, mlngModule
        
        zlDatabase.SetPara "时价卫材以加价率入库", chk加成入库.Value, glngSys, mlngModule
        zlDatabase.SetPara "时价卫材入库时取上次售价", chk取上次售价.Value, glngSys, mlngModule
        zlDatabase.SetPara "卫材分段加成率", chk分段加成入库.Value, glngSys, mlngModule
    End If
    
    '其他入库
    If mlngModule = 1714 Then
        zlDatabase.SetPara "时价卫材以加价率入库", chk加成入库.Value, glngSys, mlngModule
        zlDatabase.SetPara "时价卫材入库时取上次售价", chk取上次售价.Value, glngSys, mlngModule
        zlDatabase.SetPara "卫材分段加成率", chk分段加成入库.Value, glngSys, mlngModule
    End If
    
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdOk_Click()
    If ISValid = False Then Exit Sub
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Private Sub initPara()
    '-------------------------------------------------------------------------------------------------------------
    '功能:初始化参数设置
    '返回:
    '编制:刘兴宏
    '修改:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim strBidMess As String
    Dim int加成率入库 As Integer    '其他出库也用这个变量
    Dim int取上次售价 As Integer    '其他出库也用这个变量
    Dim int分段加成入库 As Integer  '其他出库也用这个变量
    
    '装入缺省数据
    With cbo列名
        .Clear
        .AddItem "输入顺序"
        .ItemData(.NewIndex) = 0
        .AddItem "编码"
        .ItemData(.NewIndex) = 1
        .AddItem "卫材名称"
        .ItemData(.NewIndex) = 2
        If mstrFunction = "卫材盘点管理" Then
            .AddItem "库房货位"
            .ItemData(.NewIndex) = 3
        End If
        .ListIndex = 0
    End With
    
    With cbo方向
        .Clear
        .AddItem "升序"
        .ItemData(.NewIndex) = 0
        .AddItem "降序"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    fraBidMess.Visible = False
    
    strValue = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00", Array(cbo列名, cbo方向, fra排序, lbl排序说明), mblnHavePriv)
    strValue = IIf(strValue = "", "00", strValue)
    cbo列名.ListIndex = Val(Mid(strValue, 1, 1))
    cbo方向.ListIndex = Val(Right(strValue, 1))
    cbo方向.Enabled = Not (cbo列名.ListIndex = 0)
    
    chkSavePrint.Value = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0", Array(chkSavePrint), mblnHavePriv)) = 1, 1, 0)
    chkVerifyPrint.Value = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0", Array(chkVerifyPrint), mblnHavePriv)) = 1, 1, 0)
    If mlngModule = 1725 Then
        chkStock.Value = IIf(Val(zlDatabase.GetPara("是否选择部门", glngSys, mlngModule, "0", Array(chkStock, fra库房选择, lbl库房选择说明), mblnHavePriv)) = 1, 1, 0)
    Else
        chkStock.Value = IIf(Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModule, "0", Array(chkStock, fra库房选择, lbl库房选择说明), mblnHavePriv)) = 1, 1, 0)
    End If
    
    With CboUnit1
        .Clear
        .AddItem "散装单位"
        .AddItem "包装单位"
    End With

    With cboUnit
        .Clear
        .AddItem "散装单位"
        .AddItem "包装单位"
    End With
    cboUnit.ListIndex = IIf(Val(zlDatabase.GetPara(IIf(mlngModule = 1719, "盘点表单位", "卫材单位"), glngSys, mlngModule, "0", Array(cboUnit, lbl盘点表), mblnHavePriv)) = 1, 1, 0)
    If mstrFunction <> "卫材盘点管理" Then
        CboUnit1.Visible = False
        lbl盘点表.Visible = False
        lbl盘点单.Visible = False
        cboUnit.Left = lbl盘点表.Left
        cboUnit.Width = Frame2.Width - cboUnit.Left - 250
        Label2.Top = lbl盘点单.Top
    Else
        CboUnit1.ListIndex = IIf(Val(zlDatabase.GetPara("记录单单位", glngSys, mlngModule, "0", Array(CboUnit1, lbl盘点单), mblnHavePriv)) = 1, 1, 0)
    End If
     
     
    chk修改单据号.Visible = True
    chk修改单据号.Value = IIf(Val(zlDatabase.GetPara("修改单据号", glngSys, mlngModule, "0", Array(chk修改单据号), mblnHavePriv)) = 1, 1, 0)
    chk跟踪在用.Visible = False
    chkSet库房.Visible = False
    
    Select Case mstrFunction
        Case "卫材盘点管理"
            chkSet库房.Visible = True
            chkSet库房.Value = IIf(Val(zlDatabase.GetPara("存储库房", glngSys, mlngModule, "0", Array(chkSet库房), mblnHavePriv)) = 1, 1, 0)
        Case "卫材外购入库管理"
            chkFixPrice.Visible = True
            chk允许修改批发价.Visible = True
            chk招标卫材.Visible = True
            chk高值卫材录入.Visible = True
            fraBidMess.Visible = True

            chkFixPrice.Value = IIf(Val(zlDatabase.GetPara("定价采购", glngSys, mlngModule, "0", Array(chkFixPrice), mblnHavePriv)) = 1, 1, 0)
            chk允许修改批发价.Value = IIf(Val(zlDatabase.GetPara("修改采购限价", glngSys, mlngModule, "0", Array(chk允许修改批发价), mblnHavePriv)) = 1, 1, 0)
            chk招标卫材.Value = IIf(Val(zlDatabase.GetPara("招标卫材可选择非中标单位入库", glngSys, mlngModule, "0", Array(chk招标卫材), mblnHavePriv)) = 1, 1, 0)
            chk高值卫材录入.Value = IIf(Val(zlDatabase.GetPara("高值卫材必须填写详细信息", glngSys, mlngModule, "0", Array(chk高值卫材录入), mblnHavePriv)) = 1, 1, 0)
            
            strBidMess = zlDatabase.GetPara("入库单价超中标单价", glngSys, mlngModule, , Array(optBidMess(0), optBidMess(1), optBidMess(2), fraBidMess), mblnHavePriv)
            optBidMess(Val(strBidMess)).Value = True
            
            int加成率入库 = Val(zlDatabase.GetPara("时价卫材以加价率入库", glngSys, mlngModule, 1, Array(chk加成入库), mblnHavePriv))
            int取上次售价 = Val(zlDatabase.GetPara("时价卫材入库时取上次售价", glngSys, mlngModule, 0, Array(chk取上次售价), mblnHavePriv))
            int分段加成入库 = Val(zlDatabase.GetPara("卫材分段加成率", glngSys, mlngModule, 0, Array(chk分段加成入库), mblnHavePriv))
            
            '参数规则检查
            If int加成率入库 = 1 Then
                int取上次售价 = 0
                int分段加成入库 = 0
            ElseIf int取上次售价 = 1 Then
                int加成率入库 = 0
                int分段加成入库 = 0
            ElseIf int分段加成入库 = 1 Then
                int加成率入库 = 0
                int取上次售价 = 0
            End If
            
            chk加成入库.Visible = True
            chk取上次售价.Visible = True
            chk分段加成入库.Visible = True
            
            chk加成入库.Value = int加成率入库
            chk取上次售价.Value = int取上次售价
            chk分段加成入库.Value = int分段加成入库
        Case "卫材其他入库管理"
            int加成率入库 = Val(zlDatabase.GetPara("时价卫材以加价率入库", glngSys, mlngModule, 1, Array(chk加成入库), mblnHavePriv))
            int取上次售价 = Val(zlDatabase.GetPara("时价卫材入库时取上次售价", glngSys, mlngModule, 0, Array(chk取上次售价), mblnHavePriv))
            int分段加成入库 = Val(zlDatabase.GetPara("卫材分段加成率", glngSys, mlngModule, 0, Array(chk分段加成入库), mblnHavePriv))
            
            '参数规则检查
            If int加成率入库 = 1 Then
                int取上次售价 = 0
                int分段加成入库 = 0
            ElseIf int取上次售价 = 1 Then
                int加成率入库 = 0
                int分段加成入库 = 0
            ElseIf int分段加成入库 = 1 Then
                int加成率入库 = 0
                int取上次售价 = 0
            End If
            
            chk加成入库.Visible = True
            chk取上次售价.Visible = True
            chk分段加成入库.Visible = True
            
            chk加成入库.Value = int加成率入库
            chk取上次售价.Value = int取上次售价
            chk分段加成入库.Value = int分段加成入库
        Case "卫材计划管理", "卫材申购管理"
            chk修改单据号.Visible = False
        Case "卫材领用管理"
            chk领药财务审核.Visible = True
            chk领药财务审核.Value = IIf(Val(zlDatabase.GetPara("审核流程", glngSys, mlngModule, "0", Array(chk领药财务审核), mblnHavePriv)) = 1, 1, 0)
            chk跟踪在用.Visible = True
            chk跟踪在用.Value = IIf(Val(zlDatabase.GetPara("跟踪在用", glngSys, mlngModule, "0", Array(chk跟踪在用), mblnHavePriv)) = 1, 1, 0)
        Case Else
    End Select
    
    If mstrFunction <> "卫材外购入库管理" Then
        fra排序.Height = Frame3.Height
    End If
    
    fra排序.Enabled = (InStr(1, "卫材付款事务", mstrFunction) = 0)
    If fra排序.Enabled = False Then
        cbo列名.Enabled = False
        cbo方向.Enabled = False
    End If
    If Frame2.Enabled = False Then
        cboUnit.Enabled = False
    End If
    
    fra库房选择.Enabled = (InStr(1, "卫材付款事务", mstrFunction) = 0)
    Me.cmdPrintSet.Enabled = InStr(1, gstrPrivs, ";单据打印;") <> 0
    
    If fra库房选择.Enabled = False Then
        chkStock.Enabled = False
    End If
    
    If mstrFunction = "卫材移库管理" Then
        mblnFirstLoad = False
        chk移库流程控制.Value = IIf(Val(zlDatabase.GetPara("移库流程", glngSys, mlngModule, "0", Array(chk移库流程控制, lbl移库说明, fra移库流程控制), mblnHavePriv)) = 1, 1, 0)
        chkRequestStrike.Value = IIf(Val(zlDatabase.GetPara("冲销申请", glngSys, mlngModule, "0", Array(chkRequestStrike, fra移库流程控制), mblnHavePriv)) = 1, 1, 0)
        mblnFirstLoad = True
    Else
        fra移库流程控制.Visible = False
        
        tabMain.Height = tabMain.Height - fra移库流程控制.Height
        
        cmdHelp.Top = cmdHelp.Top - fra移库流程控制.Height
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = Me.Height - fra移库流程控制.Height
    End If
    
    '资质校验页面
    tabMain.TabVisible(1) = mstrFunction = "卫材外购入库管理" Or mstrFunction = "卫材计划管理"
    If tabMain.TabVisible(1) = True Then
        With vsfCheck
            .MergeCol(0) = True
            .MergeCells = flexMergeRestrictColumns
        End With
        If mstrFunction = "卫材外购入库管理" Then
            fraCheck.Top = tabMain.Height - fraCheck.Height - 100
            chkProduceDate.Top = fraCheck.Top - chkProduceDate.Height - 100
            vsfCheck.Height = chkProduceDate.Top - vsfCheck.Top - 100
        Else
            '卫材计划由于不能输入生产日期 所以不需要此参数
            fraCheck.Top = tabMain.Height - fraCheck.Height - 100
            vsfCheck.Height = fraCheck.Top - vsfCheck.Top - 100
        End If
        
        If mstrFunction = "卫材外购入库管理" Then
            lblComment.Caption = "    说明：卫材外购入库编辑单据时是否校验卫材、生成商、供应商的信息是否完整，及资质是否过期。请选择需要进行校验的项目，并双击“校验”列打勾。"
        ElseIf mstrFunction = "卫材计划管理" Then
            lblComment.Caption = "    说明：卫材计划管理审核单据时是否校验卫材、生成商、供应商的信息是否完整，及资质是否过期。请选择需要进行校验的项目，并双击“校验”列打勾。"
        End If
        
        Load资质校验
    End If
    
    chk时价调价.Visible = False
    If mstrFunction = "卫材调价管理" Then
        chk时价调价.Value = IIf(Val(zlDatabase.GetPara("时价卫材按批次调价", glngSys, mlngModule, "0", Array(chk时价调价), mblnHavePriv)) = 1, 1, 0)
        fra库房选择.Visible = False
        fra排序.Visible = False
        fraBidMess.Visible = False
        Frame3.Visible = True
        Frame3.Top = Frame2.Top
        Frame3.Height = Frame2.Height
        Frame3.Enabled = True
        chk时价调价.Visible = True
        fra移库流程控制.Visible = False
        Frame2.Left = fra库房选择.Left
        chkSavePrint.Visible = False
        chkVerifyPrint.Visible = False
        chk修改单据号.Visible = False
        tabMain.Height = Frame2.Top + Frame2.Height + 100
        tabMain.Width = Frame2.Width + Frame3.Width + 300
        cmdOK.Top = tabMain.Height + tabMain.Top + 150
        cmdHelp.Top = cmdOK.Top
        cmdHelp.Left = 100
        cmdCancel.Top = cmdOK.Top
        cmdCancel.Left = tabMain.Width - cmdCancel.Width
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        Me.Height = tabMain.Top + tabMain.Height + 1100
        Me.Width = tabMain.Width + 250
    End If
End Sub

Public Sub 设置参数(ByVal lngModule As Long, ByVal strPrivs As String, ByVal frmMain As Form, Optional ByVal strFunction As String = "")
    '-------------------------------------------------------------------------------------------------------------
    '功能:设置相关单据操作的控制参数
    '参数:lngModule-模块号
    '     str权限串-权限串
    '     frmMain-调用的主窗体
    '     strFunction-功能说明
    '返回:
    '编制:刘兴宏
    '修改:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs: mlngModule = lngModule: mstrFunction = strFunction
    mblnHavePriv = IsHavePrivs(mstrPrivs, "参数设置")
    
    Call initPara
    Call SetCtlEnabled
    frmParaset.Show vbModal, frmMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_" & glngModul
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFirstLoad = False
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If vsfCheck.Enabled = True Then vsfCheck.SetFocus
        If vsfCheck.TextMatrix(4, 1) = "" Then
            chkProduceDate.Enabled = False
            chkProduceDate.Value = 0
        Else
            chkProduceDate.Enabled = True
        End If
    End If
End Sub

Private Sub vsfCheck_DblClick()
    With vsfCheck
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("校验") Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "√" Then
            .TextMatrix(.Row, .Col) = ""
            If .Row = 4 Then
                '注册证有效期检查
                chkProduceDate.Enabled = False
                chkProduceDate.Value = 0
            End If
        Else
            .TextMatrix(.Row, .Col) = "√"
            If .Row = 4 Then
                '注册证有效期检查
                chkProduceDate.Enabled = True
            End If
        End If
    End With
End Sub


