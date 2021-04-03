VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm保险项目 
   BackColor       =   &H8000000A&
   Caption         =   "医保项目管理"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   10110
   Icon            =   "frm保险项目.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ZL9BillEdit.BillEdit mshSum_S 
      Height          =   2745
      Left            =   3090
      TabIndex        =   4
      Top             =   1020
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4842
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2670
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":0E42
            Key             =   "R"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":115C
            Key             =   "C"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":12B6
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   90
      TabIndex        =   7
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   3450
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":1708
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":1924
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":1B40
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":1D5A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":1F76
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   2760
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":2192
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":23AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":25CA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":27E4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目.frx":2A00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1270
      BandCount       =   2
      _CBWidth        =   10110
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "保险类别"
      Child2          =   "cmb险类"
      MinHeight2      =   300
      Width2          =   2325
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin VB.ComboBox cmb险类 
         Height          =   300
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   3675
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   635
      SimpleText      =   $"frm保险项目.frx":2C1C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm保险项目.frx":2C63
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12779
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "还原(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6750
      TabIndex        =   6
      Top             =   4050
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5340
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Visible         =   0   'False
      Begin VB.Menu mnuEditGet 
         Caption         =   "重新提取项目审核信息(&G)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolSplit 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "不编辑项目编码(&I)"
      End
      Begin VB.Menu mnuViewClass 
         Caption         =   "不编辑医保大类(&C)"
      End
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R) "
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)…"
      End
   End
End
Attribute VB_Name = "frm保险项目"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private int审核标志 As Integer
Private classInsure As New clsInsure

Private Enum ColumnEnum
    cOL编码 = 0
    cOL名称 = 1
    col产地 = 2
    COL规格 = 3
    COL剂型 = 4
    COL单位 = 5
    col价格 = 6
    col改变方式 = 7
    col大类ID = 8
    COL医保编码 = 9
    col医保名称 = 10
    col医保剂型 = 11
    COL医保附注 = 12
    col原编码 = 13
    col大类名称 = 14
    col非医保 = 15
    'Modified By 朱玉宝 地区：长沙 原因：没法，只有加了
    col匹配序列号 = 16
    col审核标志 = 17
End Enum
Private Const mlng编码长度 As Long = 20

Dim mlngListIndex As Long   '保存上次下拉框的选择索引
Dim mblnLoad As Boolean
Dim msngStartX As Single    '移动前鼠标的位置
Dim mstr权限 As String

Dim mstrKey As String       '前一个树节点的关键值
Dim mint险类 As Integer     '当前显示的险类
Dim mint适用地区 As Integer '沈阳专用；0表示其他地区，1表示长春（允许删除已审核的项目）

Dim mlngCol As Long, mblnDesc As Boolean
Private mcnYB As New ADODB.Connection   '医保前置服务器连接

Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub cmdRestore_Click()
    'Modified By 朱玉宝 地区：长沙
    If mint险类 = TYPE_沈阳市 Then
        MsgBox "本医保不支持取消功能，请点击保存！", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("你确认要放弃修改吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    Call FillSum(True)
    mshSum_S.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim lngRow As Long
    
    
    gcnOracle.BeginTrans
    If mint险类 = TYPE_泸州市 Then gcn泸州.BeginTrans
    On Error GoTo errHandle
    
    With mshSum_S
        '处理数据
        For lngRow = 1 To .Rows - 1
            Select Case .TextMatrix(lngRow, col改变方式)
                Case "新增", "修改"
                    '将新增与修改放在一个过程中处理
'                    收费细目ID,险类,大类ID,项目编码,项目名称,附注
                    'Modified by ZYB 2004-08-17
                    If gintInsure = TYPE_乐山 Then
                        gstrSQL = "ZL_保险支付项目_Modify(" & .RowData(lngRow) & "," & mint险类 & "," & _
                                   IIf(Val(.TextMatrix(lngRow, col大类ID)) = 0, "null", .TextMatrix(lngRow, col大类ID)) & ",'" & _
                                   .TextMatrix(lngRow, COL医保编码) & "','" & Split(.TextMatrix(lngRow, col医保名称), "-")(1) & "','" & .TextMatrix(lngRow, COL医保附注) & _
                                   IIf(mint险类 = TYPE_沈阳市, "^^" & .TextMatrix(lngRow, col匹配序列号) & "||" & _
                                   IIf(Trim(.TextMatrix(lngRow, col审核标志)) = "√", 1, IIf(Trim(.TextMatrix(lngRow, col审核标志)) = "×", 2, 0)), "") & _
                                   "'," & IIf(Trim(.TextMatrix(lngRow, col非医保)) = "√", 0, 1) & ")"
                    Else
                        gstrSQL = "ZL_保险支付项目_Modify(" & .RowData(lngRow) & "," & mint险类 & "," & _
                                   IIf(Val(.TextMatrix(lngRow, col大类ID)) = 0, "null", .TextMatrix(lngRow, col大类ID)) & ",'" & _
                                   .TextMatrix(lngRow, COL医保编码) & "','" & .TextMatrix(lngRow, col医保名称) & "','" & .TextMatrix(lngRow, COL医保附注) & _
                                   IIf(mint险类 = TYPE_沈阳市, "^^" & .TextMatrix(lngRow, col匹配序列号) & "||" & _
                                   IIf(Trim(.TextMatrix(lngRow, col审核标志)) = "√", 1, IIf(Trim(.TextMatrix(lngRow, col审核标志)) = "×", 2, 0)), "") & _
                                   "'," & IIf(Trim(.TextMatrix(lngRow, col非医保)) = "√", 0, 1) & ")"
                    End If
                    Call DebugTool("准备保存本次修改")
                    Call ExecuteProcedure(Me.Caption)
                    Call DebugTool("修改成功")
                    
                    gstrSQL = ""
                    If .TextMatrix(lngRow, COL医保编码) <> .TextMatrix(lngRow, col原编码) Then
                        '保存修改记录
                        gstrSQL = "Insert Into 项目对应日志(中心药典序号,中心药典名称,医院药典名称,修改人,发生日期) " & _
                        "values('" & .TextMatrix(lngRow, COL医保编码) & "','" & .TextMatrix(lngRow, col医保名称) & "','" & .TextMatrix(lngRow, cOL名称) & "','" & gstrUserName & "',sysdate)"
                    End If
                    
                    Call DebugTool("保存项目修改日志:" & gstrSQL)
                    If gstrSQL <> "" Then
                        Select Case mint险类
                        Case TYPE_泸州市
                            gcn泸州.Execute gstrSQL
                        Case TYPE_重庆市
                            gcnOracle.Execute gstrSQL
                        Case TYPE_重庆银海版
                            mcnYB.Execute gstrSQL
                        End Select
                    End If
                    Call DebugTool("修改日志保存成功！")
                    
                    .TextMatrix(lngRow, col原编码) = .TextMatrix(lngRow, COL医保编码)
                Case "删除"
                    '删除的项目
                    If .TextMatrix(lngRow, col原编码) <> "" Then
                        gstrSQL = "Insert Into 项目对应日志(中心药典序号,中心药典名称,医院药典名称,修改人,发生日期) " & _
                        "values('000000','无医保项目','" & .TextMatrix(lngRow, cOL名称) & "','" & gstrUserName & "',sysdate)"
                    End If
                    Select Case mint险类
                    Case TYPE_泸州市
                        gcn泸州.Execute gstrSQL
                    Case TYPE_重庆市
                        gcnOracle.Execute gstrSQL
                    Case TYPE_重庆银海版
                        mcnYB.Execute gstrSQL
                    End Select
                    
                    gstrSQL = "ZL_保险支付项目_Delete(" & .RowData(lngRow) & "," & mint险类 & ")"
                    Call ExecuteProcedure(Me.Caption)
                    .TextMatrix(lngRow, col原编码) = .TextMatrix(lngRow, COL医保编码)
            End Select
        Next
        
        '待数据处理完成无误后，再设置数据状态
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, col改变方式) = ""
        Next
    End With
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    gcnOracle.CommitTrans
    If mint险类 = TYPE_泸州市 Then gcn泸州.CommitTrans
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
    If mint险类 = TYPE_泸州市 Then gcn泸州.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call FillTree
    End If
    
    Call mshSum_S_EnterCell(1, cOL编码)
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mstrKey = ""
    mlngCol = 0
    mblnDesc = False
    mblnLoad = True
    
    gstrSQL = "select 序号,名称 from 保险类别 where nvl(是否禁止,0)<>1  order by 序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With cmb险类
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("名称")
            .ItemData(.NewIndex) = rsTemp("序号")
            If rsTemp("序号") = gintInsure Then
                '当前医保。
                '使用API，可以不激活Click事件
                zlControl.CboSetIndex .hwnd, .NewIndex
                Call Fill大类
            End If
            
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            '使用API，可以不激活Click事件
            zlControl.CboSetIndex .hwnd, 0
            Call Fill大类
        End If
    End With
    mint险类 = cmb险类.ItemData(cmb险类.ListIndex)
    
    Call InitSum
    RestoreWinState Me, App.ProductName
    
    mnuViewItem.Checked = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", "False") <> "False"
    If mnuViewItem.Checked = False Then
        '不用判断大类了
        mnuViewClass.Checked = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", "False") <> "False"
    End If
    Call SetSkip
    
    zlControl.CboSetHeight cmb险类, 3600
    '得到查询的时间范围
    If mint险类 = TYPE_沈阳市 Then
        mint适用地区 = 0
        gstrSQL = "Select 参数值 From 保险参数 Where 参数名='适用地区'"
        Call OpenRecordset(rsTemp, "读取适用地区")
        If Not rsTemp.EOF Then
            mint适用地区 = Nvl(rsTemp!参数值, 0)
        End If
        mnuEdit.Visible = True
    End If
End Sub

Private Sub InitSum()
'初始化汇总表的样式
    Dim lngCol As Long
    
    With mshSum_S
        ClearGrid mshSum_S
        
        'Modified By 朱玉宝 地区：长沙 原因：增加列——匹配序列号
        .Cols = 18
        
        .TextMatrix(0, cOL编码) = "编码"
        .TextMatrix(0, cOL名称) = "收费细目"
        .TextMatrix(0, COL规格) = "规格"
        .TextMatrix(0, col产地) = "产地"
        .TextMatrix(0, COL单位) = "单位"
        If mint险类 = TYPE_新都 Then
            .TextMatrix(0, col价格) = "自付比例"
        Else
            .TextMatrix(0, col价格) = "价格"
        End If
        .TextMatrix(0, col改变方式) = "是否修改"
        If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
            .TextMatrix(0, COL医保编码) = "类别"
        Else
            .TextMatrix(0, COL医保编码) = "医保项目编码"
        End If
        If mint险类 = type_大连市 Then
            .TextMatrix(0, col医保名称) = "事业公费比例"
        Else
            .TextMatrix(0, col医保名称) = "医保项目名称"
        End If
        .TextMatrix(0, COL剂型) = "剂型"
        .TextMatrix(0, col医保剂型) = "剂型"
        .TextMatrix(0, col审核标志) = "审核"
        .TextMatrix(0, COL医保附注) = "医保项目附注"
        .TextMatrix(0, col原编码) = "原医保项目编码"
        .TextMatrix(0, col大类ID) = "大类ID"
        .TextMatrix(0, col大类名称) = "医保大类名称"
        
        If mint险类 = TYPE_黔南 Then
            .TextMatrix(0, col非医保) = "检查审批"
        Else
            .TextMatrix(0, col非医保) = "是否医保"
        End If
        
        .TextMatrix(0, col匹配序列号) = "匹配序列号"
        
        .ColWidth(cOL编码) = 1000
        .ColWidth(cOL名称) = 2000
        .ColWidth(COL规格) = 1000
        .ColWidth(col产地) = 600
        .ColWidth(COL单位) = 600
        .ColWidth(col价格) = 800
        .ColWidth(col改变方式) = 0
        .ColWidth(COL医保编码) = 1200
        .ColWidth(col医保名称) = 1200
        .ColWidth(COL医保附注) = 0
        .ColWidth(col原编码) = 0
        .ColWidth(col大类ID) = 0
        .ColWidth(col大类名称) = 1200
        .ColWidth(col非医保) = 800
        .ColWidth(col匹配序列号) = 0
        
        If mint险类 = TYPE_沈阳市 Then
            .ColWidth(COL剂型) = 700
            .ColWidth(col医保剂型) = 700
            .ColWidth(col审核标志) = 400
        Else
            .ColWidth(COL剂型) = 0
            .ColWidth(col医保剂型) = 0
            .ColWidth(col审核标志) = 0
        End If
        
        For lngCol = 0 To .Cols - 1
            .ColAlignment(lngCol) = 1
        Next
        .ColAlignment(col价格) = 7
        .ColAlignment(col非医保) = 4
        
        '设置各列的编辑特性
        .ColData(COL剂型) = 5
        .ColData(col医保剂型) = 5
        .ColData(col审核标志) = 5
        .ColData(cOL编码) = 5 '不能选择
        .ColData(cOL名称) = 5
        .ColData(COL规格) = 5
        .ColData(col产地) = 5
        .ColData(COL单位) = 5
        .ColData(col价格) = 5
        .ColData(col改变方式) = 5
        If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
            .ColData(COL医保编码) = 3
        Else
            .ColData(COL医保编码) = 1
        End If
        If mint险类 = type_大连市 Then
            .ColData(col医保名称) = 4
        Else
            .ColData(col医保名称) = 5
        End If
        .ColData(COL医保附注) = 5
        .ColData(col原编码) = 5
        .ColData(col大类ID) = 5
        .ColData(col大类名称) = 3 '选择器
        .ColData(col非医保) = -1 '选择器
        .ColData(col匹配序列号) = 5
        
        .PrimaryCol = cOL编码
        
        If gintInsure = TYPE_成都南充 Then
            .TxtCheck = True
            .TextMask = "`"
        End If
                
        Call SetSkip
        .AllowAddRow = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSave.Enabled = True Then
        MsgBox "医保项目列表正处于编辑状态，不能退出程序。", vbInformation, gstrSysName
        Cancel = 1
        Exit Sub
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", mnuViewItem.Checked
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", mnuViewClass.Checked
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbrThis.Height)
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrHeight, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    '右边
    'tvwMain_S的位置
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    'picV的位置
    picV.Top = sngTop
    picV.Height = tvwMain_S.Height
    picV.Left = tvwMain_S.Left + tvwMain_S.Width
    
    cmdRestore.Top = sngBottom - cmdRestore.Height - 100
    cmdRestore.Left = ScaleWidth - cmdRestore.Width - 300
    cmdSave.Top = cmdRestore.Top
    cmdSave.Left = cmdRestore.Left - cmdSave.Width - 300
    
    If InStr(mstr权限, "增删改") > 0 Then
        '可以编辑
        sngBottom = cmdRestore.Top - 100
    End If
    
    mshSum_S.Left = picV.Left + picV.Width
    If ScaleWidth - mshSum_S.Left > 0 Then mshSum_S.Width = ScaleWidth - mshSum_S.Left
    mshSum_S.Top = sngTop
    mshSum_S.Height = IIf(sngBottom - mshSum_S.Top > 0, sngBottom - mshSum_S.Top, 0)
    
    Refresh
End Sub

Private Function GetMatch(ByVal rsMatch As ADODB.Recordset, ByVal intType As Integer) As Boolean
    Dim str编码 As String, str匹配序列号 As String, strTmp As String, str匹配类型 As String
    Dim int审核标志 As Integer
    '重新提取医保中心的匹配信息，并更新本地数据库
    'intType=0：诊疗项目;1：药品项目
    
    '取药品类匹配信息
    If Not 调用接口_准备_沈阳市(Function_沈阳市.项目匹配_取匹配项目信息) Then Exit Function
    gstrField_沈阳市 = "hospital_id||audit_status||item_type"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||1||" & intType
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市() Then Exit Function
    If Not 调用接口_指定记录集_沈阳市("ItemMatch") Then Exit Function
'    序号    字段    字段说明    最大长度    备注
'    1   hosp_code   医院目录编码    20
'    2   hosp_name   医院目录名称    60
'    3   hosp_model  医院目录剂型    20
'    4   item_name   中心目录名称    60
'    5   model_name  中心目录剂型    20
'    6   serial_match    匹配序列号  12
'    7   valid_flag  有效标志    1   "0"：无效    "1"：有效
'    8   audit_flag  审核标志    1   "0"：未审核    "1"：审核通过    "2"：审核未通过
'    9   match_type  匹配类型    1   "0"：诊疗项目匹配    "1"：西药匹配    "2"：中成药匹配    "3"：中草药匹配
    If 调用接口_记录数_沈阳市 Then
        Do While True
            Call 调用接口_读取数据_沈阳市("hosp_code", str编码)
            Call 调用接口_读取数据_沈阳市("serial_match", str匹配序列号)
            Call 调用接口_读取数据_沈阳市("match_type", str匹配类型)
            Call 调用接口_读取数据_沈阳市("audit_flag", strTmp)
            int审核标志 = Val(strTmp)
            
            '定位该记录，找出收费细目ID
            rsMatch.Filter = "编码='" & str编码 & "'"
            
            If Not rsMatch.EOF Then
                '更新保险支付项目
                gstrSQL = "ZL_保险支付项目_Modify(" & rsMatch!收费细目ID & "," & TYPE_沈阳市 & "," & rsMatch!大类ID & ",'" & _
                           rsMatch!项目编码 & "','" & rsMatch!项目名称 & "','" & Split(rsMatch!附注, "^^")(0) & "^^" & str匹配序列号 & "||" & int审核标志 & _
                           "'," & rsMatch!是否医保 & ")"
                Call ExecuteProcedure("更新保险支付项目")
            Else
                MsgBox "接口返回的医院编码或标识码[" & str编码 & "]，但在本地保险支付项目中，未找到该收费细目", vbInformation, gstrSysName
            End If
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    MsgBox "已从中心成功获取所有项目的匹配信息！", vbInformation, gstrSysName
    GetMatch = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub mnuEditGet_Click()
    Dim rsMatch As New ADODB.Recordset
    On Error GoTo ErrHand
    If MsgBox("这个操作可能会花很长时间，你确定要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = " (Select ID 收费细目ID,Decode(TRIM(标识主码),NULL,编码,'',编码,标识主码) 编码 From 收费细目 Where 类别 Not In ('5','6','7')" & _
              " Union " & _
              " Select 药品ID 收费细目ID,Decode(Trim(标识码),NULL,编码,'',编码,标识码) 编码 From 药品目录)"
    gstrSQL = " Select B.编码,A.收费细目ID,A.大类ID,A.项目编码,A.项目名称,A.附注,A.是否医保 " & _
              " From 保险支付项目 A," & gstrSQL & " B" & _
              " Where A.收费细目ID=B.收费细目ID And A.险类=" & TYPE_沈阳市
    Call OpenRecordset(rsMatch, "取保险支付项目")
    
    If Not classInsure.InitInsure(gcnOracle, TYPE_沈阳市) Then Exit Sub
    gcnOracle.BeginTrans
    
    rsMatch.Filter = 0
    If Not GetMatch(rsMatch, 0) Then
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    rsMatch.Filter = 0
    If Not GetMatch(rsMatch, 1) Then
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    gcnOracle.CommitTrans
    
    '重新显示本页面信息
    Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuViewFind_Click()
    If cmdSave.Enabled = True Then
        MsgBox "医保项目列表正处于编辑状态，不能使用查找功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    frm保险项目查找.Show vbModal, Me
End Sub

Private Sub cmb险类_Click()
    Call Fill大类
    Call FillSum(False)
    
    'Modified By 朱玉宝 地区：长沙 原因：初始化医保接口
    If cmb险类.ItemData(cmb险类.ListIndex) <> TYPE_沈阳市 Then Exit Sub
    Call classInsure.InitInsure(gcnOracle, TYPE_沈阳市)
End Sub

Private Sub mnuViewClass_Click()
    mnuViewItem.Checked = False
    mnuViewClass.Checked = Not mnuViewClass.Checked
    
    Call SetSkip
End Sub

Private Sub mnuViewItem_Click()
    mnuViewClass.Checked = False
    mnuViewItem.Checked = Not mnuViewItem.Checked
    
    Call SetSkip
End Sub

Private Sub SetSkip()
'设置表格的跳跃属性
    With mshSum_S
        If mnuViewItem.Checked = False Then
        
            If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
            Else
                .ColData(COL医保编码) = 1
            End If
            .LocateCol = COL医保编码
            
            .ColData(col大类名称) = IIf(mnuViewClass.Checked = True, 5, 3)
        Else
            .ColData(col大类名称) = 3 '选择器
            .LocateCol = col大类名称
            If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
            Else
                .ColData(COL医保编码) = 5
            End If
        End If
        If .ColData(.Col) = 5 Then
            '当前列已经不能选择，需重新定位
            .Col = .LocateCol
        End If
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    '只刷新列表内容
    Call FillSum
End Sub

Private Sub mshSum_S_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    '始终是不允许删除的
    Cancel = True
    
    With mshSum_S
        'Modified By 朱玉宝 地区：长沙 原因：已通过医保中心审核的项目不允许删除
        If mint险类 = TYPE_沈阳市 Then
            Call GetItemMatchInfo
            If int审核标志 = 1 And mint适用地区 = 0 Then
                MsgBox "该项目已经通过医保中心审核，不允许删除。请与医保中心联系！", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
        End If
        
        If .TextMatrix(Row, col改变方式) = "新增" Then
            .TextMatrix(Row, col改变方式) = "" '相当于什么都没有做
            'Modified By 朱玉宝 地区：长沙 原因：根据当前动作设置项目匹配信息
            Call SetItemMatch
        Else
            .TextMatrix(Row, col改变方式) = "删除" '标记
            'Modified By 朱玉宝 地区：长沙 原因：根据当前动作设置项目匹配信息
            Call SetItemMatch
        End If
        
        .TextMatrix(Row, COL医保编码) = ""
        .TextMatrix(Row, col医保名称) = ""
        .TextMatrix(Row, col医保剂型) = ""
        .TextMatrix(Row, COL医保附注) = ""
        .TextMatrix(Row, col大类ID) = ""
        .TextMatrix(Row, col大类名称) = ""
        .TextMatrix(Row, col非医保) = ""
        .TextMatrix(Row, col审核标志) = ""
    End With
    cmdSave.Enabled = True
    cmdRestore.Enabled = True
End Sub

Private Sub mshSum_S_cboClick(ListIndex As Long)
    With mshSum_S
        If mint险类 = type_大连市 Or type_大连开发区 = mint险类 Then
            If .Col = COL医保编码 Then
                mlngListIndex = ListIndex
                If .TextMatrix(.Row, COL医保编码) <> .CboText Then
                    .TextMatrix(.Row, COL医保编码) = .CboText
                    Call 标记改变
                End If
                Exit Sub
            End If
        End If
        If .Col <> col大类名称 Then Exit Sub
        
        If .TextMatrix(.Row, col大类名称) <> .CboText Then
            '禁止修改保险大类,只允许通过选择明细以确定大类
            If gintInsure = TYPE_泸州市 Then
                .ListIndex = mlngListIndex
                Exit Sub
            End If
            mlngListIndex = ListIndex
            .TextMatrix(.Row, col大类名称) = .CboText
            Call 标记改变
        Else
            mlngListIndex = ListIndex
        End If
        
        If .CboText = "" Then
            '保存为空
            .TextMatrix(.Row, col大类ID) = ""
            .TextMatrix(.Row, col大类名称) = ""
        Else
            .TextMatrix(.Row, col大类ID) = .ItemData(.ListIndex)
            .TextMatrix(.Row, col大类名称) = .CboText
        End If
        
    End With
End Sub

Private Sub mshSum_S_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshSum_S
        If KeyCode = vbKeyReturn Then
              '刘兴宏(200311)
            If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
                If .Col = COL医保编码 Then
                    If .CboText = "" Then
                            .TextMatrix(.Row, COL医保编码) = " "
                            If mint险类 = type_大连市 Then
                                .Col = col医保名称
                            Else
                                .Col = col大类名称
                            End If
                    Else
                            .TextMatrix(.Row, COL医保编码) = .CboText
                    End If
                    Call 标记改变
                    Exit Sub
                End If
            End If
            If .TextMatrix(.Row, col大类名称) <> .CboText Then
                .TextMatrix(.Row, col大类名称) = .CboText
                Call 标记改变
            End If
            
            If .CboText = "" Then
                '保存为空
                .TextMatrix(.Row, col大类ID) = ""
                .TextMatrix(.Row, col大类名称) = ""
                .Col = col非医保
            Else
                .TextMatrix(.Row, col大类ID) = .ItemData(.ListIndex)
                .TextMatrix(.Row, col大类名称) = .CboText
            End If
        End If
    End With
End Sub

Private Sub mshSum_S_CommandClick()
'功能：提取医保项目供选择
'参数：无
'返回：医保项目编码
    Dim strCode As String
    Dim strSelected As String
    Dim strName As String
    Dim strlastCode As String
    Dim strMemo As String
    
    With mshSum_S
        strCode = .TextMatrix(.Row, COL医保编码)
        Select Case mint险类
            Case TYPE_重庆市, TYPE_重庆银海版
                On Error Resume Next
                If frm保险项目选择重庆.GetCode(strCode, strName, Val(.TextMatrix(.Row, col价格)), mint险类) = True Then
                    strSelected = strCode
                End If
            Case TYPE_宁海
                If frm保险项目选择宁海.GetCode(strCode, strName, mint险类) = True Then
                    strSelected = strCode
                End If
            Case TYPE_渝北农医
                If frm保险项目选择_渝北农医.GetCode(strCode, strName, mint险类) = True Then
                    strSelected = strCode
                End If
            Case TYPE_浙江
                On Error Resume Next
                If frm保险项目选择浙江.GetCode(strCode, strName, strMemo, Val(.TextMatrix(.Row, col价格)), mint险类) = True Then
                    strSelected = strCode
                End If
            Case TYPE_余姚
                On Error Resume Next
                If frm保险项目选择余姚.GetCode(strCode, strName, Val(.TextMatrix(.Row, col价格)), mint险类) = True Then
                    strSelected = strCode
                End If
            Case TYPE_新都
                On Error Resume Next
                If frm保险项目选择新都.GetCode(strCode, strName, Val(.TextMatrix(.Row, col价格)), mint险类) = True Then
                    strSelected = strCode
                End If
                
            Case TYPE_重庆渝北
                '刘兴宏:20040706
                On Error Resume Next
                If frm保险项目选择重庆渝北.GetCode(Me, strCode, strName) = True Then
                    strSelected = Mid(strCode, 2)
                    .TextMatrix(.Row, COL医保附注) = Mid(strCode, 1, 1)
                End If
            Case TYPE_黔南
                On Error Resume Next
                If frm保险项目选择黔南.GetCode(Me, strCode, strName) = True Then
                    strSelected = strCode
                End If
            Case TYPE_成都莲合
                '没有提供获取编码途径
            Case TYPE_成都南充
                If frm保险项目选择南充.GetCode(strCode, strName) Then
                    strSelected = strCode
                End If
            Case TYPE_北京
                strName = .TextMatrix(.Row, col医保名称)
                If frm保险项目选择北京.GetCode(strCode, strName, TYPE_北京) = False Then Exit Sub
                strSelected = strCode
                '如果是药品项目，检查商品名和别名是否在医保中心下发的药品别名中，如果是才允许设置对照
                If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                    If Not CheckTradeName(.RowData(.Row), strCode) Then
                        Exit Sub
                    End If
                End If
            'Modified by ZYB 毕节
            Case TYPE_毕节
                If frm保险项目选择毕节.GetCode(strCode, strName, mint险类) = True Then
                    strSelected = strCode
                End If
            Case Else
                If mint险类 = TYPE_沈阳市 Then
                    Call GetItemMatchInfo
                    If int审核标志 = 1 And mint适用地区 = 0 Then
                        MsgBox "该项目已经通过审核，不允许修改！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                If frm保险项目选择.GetCode(strCode, mint险类) = True Then
                    strSelected = strCode
                    If mint险类 = TYPE_沈阳市 Then
                        Call CheckValid(strCode)
                    End If
                End If
        End Select
        
        If strSelected <> "" Then
            If mint险类 = TYPE_黔南 Then
                .TextMatrix(.Row, COL医保编码) = Mid(strSelected, 2)
                .TextMatrix(.Row, COL医保附注) = Mid(strSelected, 1, 1)
            Else
                .TextMatrix(.Row, COL医保编码) = strSelected
            End If
            If strName = "" Or mint险类 = TYPE_重庆银海版 Or mint险类 = TYPE_重庆渝北 Or mint险类 = TYPE_毕节 Or mint险类 = TYPE_黔南 Then
                Call Get保险名称
            Else
                '已经传入名称，就不用再调用
                .TextMatrix(.Row, col医保名称) = strName
                If mint险类 = TYPE_浙江 Then
                    .TextMatrix(.Row, COL医保附注) = strMemo
                Else
                    .TextMatrix(.Row, COL医保附注) = ""
                End If
                .TextMatrix(.Row, col非医保) = ""
            End If
            Call 标记改变
            'Modified By 朱玉宝 地区：长沙 原因：根据当前动作设置项目匹配信息
            If mint险类 = TYPE_沈阳市 Then
                .TextMatrix(.Row, col医保剂型) = Split(.TextMatrix(.Row, COL医保附注), "||")(3)
            End If
            Call SetItemMatch(False)
        End If
    End With
End Sub

Private Sub mshSum_S_DblClick(Cancel As Boolean)
    With mshSum_S
        If .Active = False Then Exit Sub
        Call 标记改变
    End With
End Sub

Private Sub mshSum_S_EnterCell(Row As Long, Col As Long)
    Static lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    
    If Col = col大类名称 And Trim(mshSum_S.TextMatrix(Row, Col)) = "" Then
        mshSum_S.ListIndex = -1
    End If
    
    '刘兴宏(200311)
    If type_大连开发区 = mint险类 Or type_大连市 = mint险类 Then
        Select Case Col
        Case COL医保编码
            mshSum_S.Clear
            mshSum_S.AddItem ""
            mshSum_S.AddItem "大检"
            mshSum_S.AddItem "特治"
        Case col医保名称
            If type_大连市 = mint险类 Then
                mshSum_S.TxtCheck = True
                mshSum_S.MaxLength = 11
                mshSum_S.TextMask = ".1234567890"
            End If
        Case col大类名称
            Dim strSql As String
              gstrSQL = "select ID,编码,名称 from 保险支付大类 " & _
              "where 险类=" & cmb险类.ItemData(cmb险类.ListIndex) & " order by 编码"
            Call OpenRecordset(rsTemp, Me.Caption)
            mshSum_S.Clear
            Do Until rsTemp.EOF
                mshSum_S.AddItem rsTemp("编码") & "." & rsTemp("名称")
                mshSum_S.ItemData(mshSum_S.NewIndex) = rsTemp("ID")
                rsTemp.MoveNext
            Loop
            
        End Select
    End If
    'Modified By 朱玉宝 地区：长沙 原因：获取项目匹配信息
    If mint险类 <> TYPE_沈阳市 Then Exit Sub
    If lngRow = Row Then Exit Sub
    lngRow = Row
    Call GetItemMatchInfo
End Sub

Private Sub mshSum_S_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '保险项目编码
    Dim str前 As String, strText As String, str类型 As String
    Dim rsTemp As New ADODB.Recordset, blnReturn As Boolean
    Dim strLeft As String
    Dim strTemp As String

    str前 = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", "0") = "0", "%", "") '双向匹配
    
    On Error GoTo errHandle
    
    With mshSum_S
        If mint险类 = type_大连市 And .Col = col医保名称 And KeyCode = vbKeyReturn Then
            strText = Replace(Trim(.Text), "`", "")
            If Not IsNumeric(strText) And strText <> "" Then
                ShowMsgbox "事业公费比例必须为数字型,请重输！"
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Val(strText) > 100 Then
                ShowMsgbox "事业公费比例必须小于100,请重输！"
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If strText = "" Then
                strText = " "
                .Text = " "
                If Trim(.TextMatrix(.Row, .Col)) = "" Then
                    .TextMatrix(.Row, .Col) = " "
                End If
            End If
            .Text = strText
            Call 标记改变
        End If
        If .Col <> COL医保编码 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If .TxtVisible = True Then
                strText = Replace(Trim(.Text), "`", "")
                .Text = strText
                If zlCommFun.StrIsValid(strText, mlng编码长度) = False Then
                    Cancel = True
                    Exit Sub
                End If
                If mint险类 = TYPE_成都南充 Then Exit Sub
                If Trim(strText) = "" Then
                    '不需要再去检查是否有匹配的编码，相当于删除该编码
                    .TextMatrix(.Row, COL医保编码) = Trim(strText)
                Else
                    '产生SQL语句
                    Select Case mint险类
                        Case TYPE_宁海
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '从药品目录中判断
                                str类型 = "药品"
                                gstrSQL = "" & _
                                    " SELECT YPDM AS 医保编码,ZWM AS 项目名称,PYJM AS 简码,YLFL AS 药理分类," & _
                                    "     DECODE(trim(ZFFL),'01','甲类（完全报销）','02','乙类（部分报销）','03','丙类（完全自费）','1','甲类（完全报销）','2','乙类（部分报销）','3','丙类（完全自费）','11','普通诊疗','12','诊疗自负10%','13','诊疗自负15%','14','诊疗自负20%','15','诊疗自负40%','16','监护病房1－5天自负30%','17','监护病房6－10天自负50%','19','自费诊疗','未知') AS 自负分类," & _
                                    "     ZDYYDJ AS 最低医院等级,YPGG AS 规格,YPBZDW AS 包装单位,YPJX AS 剂型,BZYYTS AS 标准用药贴数," & _
                                    "     ltrim(to_Char(BZJG,'9999990.00')) As 标准价格, ltrim(to_Char(ZYXE,'9999990.00')) As 住院限额, ltrim(to_Char(MZXE,'9999990.00')) As 门诊限额, YPCD As 产地,DECODE(SYFW,'0','门诊','1','住院','门诊住院均可使用') As 使用范围, BZSM As 备注" & _
                                    " From SIM_YPML " & _
                                    "Where (upper(YPDM) Like '" & UCase(strText) & "' Or Upper(ZWM) Like '" & UCase(strText) & "' Or Upper(PYJM) Like '" & UCase(strText) & "')"
                            Else
                                '从诊疗目录中判断
                                str类型 = "诊疗"
                                gstrSQL = "" & _
                                " SELECT ZLDM AS 医保编码,ZLMC AS 项目名称,PYJM AS 简码,ZLFL AS 诊疗分类," & _
                                "     DECODE(trim(ZFFL),'01','甲类（完全报销）','02','乙类（部分报销）','03','丙类（完全自费）','1','甲类（完全报销）','2','乙类（部分报销）','3','丙类（完全自费）','11','普通诊疗','12','诊疗自负10%','13','诊疗自负15%','14','诊疗自负20%','15','诊疗自负40%','16','监护病房1－5天自负30%','17','监护病房6－10天自负50%','19','自费诊疗','未知') AS 自负分类," & _
                                "     ltrim(to_Char(BZJG,'9999990.00')) As 标准价格, ltrim(to_Char(ZYXE,'9999990.00')) As 住院限额, ltrim(to_Char(MZXE,'9999990.00')) As 门诊限额, JLDW As 计量单位, ZDYYDJ As 最低医院等级,DECODE(SYFW,'0','门诊','1','住院','门诊住院均可使用') As 使用范围, BZSM As 备注" & _
                                " From SIM_ZLML " & _
                                "Where (upper(ZLDM) Like '" & UCase(strText) & "' Or Upper(ZLMC) Like '" & UCase(strText) & "' Or Upper(PYJM) Like '" & UCase(strText) & "')"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                '强制使记录集为打开状态
                                gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        Case TYPE_渝北农医
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '从药品目录中判断
                                str类型 = "药品"
                                gstrSQL = "" & _
                                    " SELECT YPLSH AS 项目编码,YPMC AS 项目名称,PY AS 简码,GG AS 规格,JX AS 剂型,SCCJ AS 生产厂家," & _
                                    "     CASE WHEN XJFS='0' THEN '甲类' WHEN xjfs='1' THEN '乙类' WHEN xjfs='2' THEN '自费' END AS 药品类别," & _
                                    "     CASE WHEN BXFW ='0' THEN '村级' WHEN BXFW ='1' THEN '乡镇' WHEN BXFW ='2' THEN '区县' END AS 用药范围,ZGXJ AS 最高限价," & _
                                    "     CASE WHEN LB='0' THEN '西药' WHEN LB='1' THEN '中成药' WHEN LB='2' THEN '中草药' WHEN LB='3' THEN '卫生材料' END AS 费用类别" & _
                                    " From YPML" & _
                                    " Where (upper(YPLSH) Like '" & UCase(strText) & "%' Or Upper(YPMC) Like '" & UCase(strText) & "%' Or Upper(PY) Like '" & UCase(strText) & "%')"
                            Else
                                '从诊疗目录中判断
                                str类型 = "诊疗"
                                gstrSQL = "" & _
                                    " SELECT XMBM AS 项目编码,XMMC AS 项目名称,PY AS 简码,CASE WHEN XJFS='0' THEN '甲类' WHEN XJFS='1' THEN '乙类' WHEN XJFS='2' THEN '自费' END AS 项目类别," & _
                                    "     CASE WHEN BXFW='0' THEN '村级' WHEN BXFW='1' THEN '乡镇' WHEN BXFW='2' THEN '区县' END AS 用药范围,ZGXJ AS 最高限价," & _
                                    "     CASE WHEN XMFL='0' THEN '挂号费' WHEN XMFL='1' THEN '检查费' WHEN XMFL='2' THEN '诊疗费' WHEN XMFL='3' THEN '治疗费' WHEN XMFL='4' THEN '耗材费' WHEN XMFL='5' THEN '手术费' WHEN XMFL='6' THEN '护理费' WHEN XMFL='7' THEN '床位费' WHEN XMFL='8' THEN '陪住费'" & _
                                    "          WHEN XMFL='9' THEN '放射费' WHEN XMFL='10' THEN '会诊费' WHEN XMFL='11' THEN '监护费' WHEN XMFL='12' THEN '抢救费' WHEN XMFL='13' THEN 'B超费' WHEN XMFL='14' THEN '彩超费' WHEN XMFL='15' THEN '病历费' WHEN XMFL='16' THEN '检验费' WHEN XMFL='17' THEN '碎石费'" & _
                                    "          WHEN XMFL='18' THEN 'CT费' WHEN XMFL='19' THEN '输氧费' WHEN XMFL='20' THEN '心电图费' WHEN XMFL='21' THEN '调温费' WHEN XMFL='22' THEN '理疗费' WHEN XMFL='23' THEN '接生费' WHEN XMFL='24' THEN '麻醉费' WHEN XMFL='25' THEN '婚检费' WHEN XMFL='26' THEN '体检费' WHEN XMFL='27' THEN '骨疗费' WHEN XMFL='28' THEN '其他费' END AS 费用类别" & _
                                    " From ZLXM" & _
                                " Where (upper(XMBM) Like '" & UCase(strText) & "%' Or Upper(XMMC) Like '" & UCase(strText) & "%' Or Upper(PY) Like '" & UCase(strText) & "%')"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                '强制使记录集为打开状态
                                gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        Case TYPE_重庆市
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '从药品目录中判断
                                str类型 = "药品"
                                gstrSQL = "select YPLSH  医保编码,YPBM 药品编码,TYM 通用名称,SPM 商品名,SPMZJM 商品名简码,YCMC 药厂名称,decode(FYDJ,1,'甲类',2,'乙类','自费') 费用等级 " & _
                                          "      ,PFJ 批发价,nvl(BZDJ,0) 标准单价,ZFBL 自付比例,JX 剂型,BZSL 包装数量,BZDW 包装单位,HL 含量,HLDW 含量单位,RL 容量,RLDW 容量单位 " & _
                                          "      ,DECODE(CFYBZ,1,'是') 处方药标志,decode(GMP,1,'是') GMP标志,decode(YPXJFS,1,'限价') 限价,TQFYDJ 特群项目等级,TQZFBL 特群自付比例,TQBZDJ 特群标准单价 " & _
                                         "   FROM YPML WHERE YPLSH like '" & strText & "%' or Upper(TYM) like '" & str前 & UCase(strText) & "%' Or Upper(SPM) like '" & str前 & UCase(strText) & "%' Or Upper(SPMZJM) like '" & str前 & UCase(strText) & "%'"
                            Else
                                '从诊疗目录中判断
                                str类型 = "诊疗"
                                gstrSQL = "Select XMLSH 医保编码,XMBM 诊疗编码,XMMC 项目名称,ZJM 简码,decode(FYDJ,1,'甲类',2,'乙类','自费') 费用等级,DW 单位 " & _
                                         "       ,nvl(TPJ,0) 特批价,nvl(BZJ,0) 标准单价,ZZBL 在职自付比例,TXBL 退休自付比例,decode(XJFS,1,'统一限价',2,'按医院等级定价',3,'按二级医院标准浮动比例') 限价 " & _
                                         "       ,decode(TPXMBZ,1,'是') 特批项目标志,TQFYDJ 特群项目等级,TQZFBL 特群自付比例,TQBZDJ 特群标准单价,BZ 备注 " & _
                                         "   FROM ZLXM WHERE XMLSH like '" & strText & "%' or Upper(XMMC) like '" & str前 & UCase(strText) & "%' Or Upper(ZJM) like '" & str前 & UCase(strText) & "%'"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                '强制使记录集为打开状态
                                gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        Case TYPE_重庆银海版
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '从药品目录中判断
                                str类型 = "药品"
                                gstrSQL = "select 流水号 医保编码,编码 药品编码,通用名 通用名称,商品名,商品名助记码 商品名简码,药厂名称,decode(项目等级,1,'甲类',2,'乙类','自费') 费用等级 " & _
                                          "      ,批发价,nvl(标准单价,0) 标准单价,自付比例,剂型,包装数量,包装单位,含量,含量单位,容量,容量单位 " & _
                                          "      ,DECODE(处方药标志,1,'是') 处方药标志,decode(GMP标志,1,'是') GMP标志,decode(限价方式,1,'限价') 限价 " & _
                                         "   FROM 中间库_药品目录 WHERE 流水号 like '" & strText & "%' or Upper(通用名) like '" & str前 & UCase(strText) & "%' Or Upper(商品名) like '" & str前 & UCase(strText) & "%' Or Upper(商品名助记码) like '" & str前 & UCase(strText) & "%'"
                            Else
                                '从诊疗目录中判断
                                str类型 = "诊疗"
                                gstrSQL = "Select 流水号 医保编码,项目编码 诊疗编码,项目名称,助记码 简码,decode(项目等级,1,'甲类',2,'乙类','自费') 费用等级,单位 " & _
                                         "       ,nvl(特批价,0) 特批价,nvl(标准单价,0) 标准单价,在职比例 在职自付比例,退休比例 退休自付比例,decode(限价方式,1,'统一限价',2,'按医院等级定价',3,'按二级医院标准浮动比例') 限价 " & _
                                         "       ,decode(特批项目标志,1,'是') 特批项目标志,备注 " & _
                                         "   FROM 中间库_诊疗项目 WHERE 流水号 like '" & strText & "%' or Upper(项目编码) like '" & str前 & UCase(strText) & "%' Or Upper(助记码) like '" & str前 & UCase(strText) & "%'"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                '强制使记录集为打开状态
                                gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        'Modified by 朱玉宝 20031218 地区：福州
                        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                            '20031229:周韬,防止重复
                            gstrSQL = "   Select Distinct A.编码 as 医保编码,A.名称,A.简码,B.名称 as 大类,A.附注 " & _
                                      "   FROM 保险项目 A,保险支付大类 B" & _
                                      "   WHERE A.大类编码=B.编码 And A.险类=" & mint险类 & " And B.险类=A.险类" & _
                                      " And (A.编码 like '" & strText & "%' or Upper(A.名称) like '" & str前 & UCase(strText) & "%' Or Upper(A.简码) like '" & str前 & UCase(strText) & "%')"
                            Call OpenRecordset(rsTemp, Me.Caption)
                        Case TYPE_泸州市
                            gstrSQL = "SELECT A.编码 医保编码,A.名称,A.简码,A.单位,B.名称 AS 大类,C.名称 AS 剂型 " & _
                                      "     ,A.是否是药,A.是否医保,A.最大价格限制,A.首先自付比例,A.价格,A.项目内涵,A.除外内容,A.说明 " & _
                                      "  FROM 保险项目 A,保险支付大类 B,剂型 C " & _
                                      "  WHERE A.险类=" & TYPE_泸州市 & " AND A.大类编码=B.编码(+) AND A.剂型编码=c.编码(+) And (" & _
                                      zlCommFun.GetLike("A", "编码", strText) & " Or " & zlCommFun.GetLike("A", "名称", strText) & " Or " & zlCommFun.GetLike("A", "简码", strText) & ")"
                            rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
                        Case type_大连市, type_大连开发区
                            '200311
                            gstrSQL = "   Select A.编码  医保编码,A.名称,A.简码,B.名称 大类,A.附注 " & _
                                      "   FROM 保险项目 A,保险支付大类 B" & _
                                      "   WHERE A.大类编码=B.编码 and b.险类=" & mint险类 & " And A.险类=" & mint险类 & " and (A.编码 like '" & strText & "%' or Upper(A.名称) like '" & str前 & UCase(strText) & "%' Or Upper(A.简码) like '" & str前 & UCase(strText) & "%')"
                                      
                            Call OpenRecordset(rsTemp, Me.Caption)
                        Case TYPE_北京
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "" & _
                                    " Select A.编码 AS 医保编码,A.类目,A.名称,A.助记码,A.剂量单位 AS 单位,B.名称 As 特殊病,H.名称 AS 项目等级,A.标准单价,A.自付比例,0 限价," & _
                                    " C.名称 AS 处方药,F.名称 AS 剂型,A.用法,A.日常规用量,D.名称 AS 药品分类,G.名称 AS 产地,E.名称 AS 使用限制等级,A.备注,A.生效日期" & _
                                    " From 药品目录 A," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='特殊用药标识' and A.类别=B.类别) B," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='处方药标志' and A.类别=B.类别) C," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='药品分类' and A.类别=B.类别) D," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='使用限制等级' and A.类别=B.类别) E,"
                                gstrSQL = gstrSQL & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='剂型' and A.类别=B.类别) F," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='产地' and A.类别=B.类别) G," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='收费项目等级' and A.类别=B.类别) H" & _
                                    " Where A.特殊病 =B.编码(+) And A.处方药=C.编码(+) And A.药品分类 =D.编码(+)" & _
                                    " And A.使用限制等级=E.编码(+) And A.剂型=F.编码(+) And A.产地=G.编码(+) AND A.药品等级=H.编码(+)" & _
                                    " And (" & zlCommFun.GetLike("A", "编码", strText) & " Or " & zlCommFun.GetLike("A", "名称", strText) & " Or " & zlCommFun.GetLike("A", "助记码", strText) & ")"
                            Else
                                '当前选择是的一个诊疗类别
                                gstrSQL = "" & _
                                    " Select A.编码 AS 医保编码,A.名称,A.助记码,A.单位,B.名称 AS 特殊病,C.名称 AS 项目等级,A.标准单价,A.自付比例,A.限价,A.备注,A.生效日期" & _
                                    "      From 诊疗目录 A," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='特殊用药标识' and A.类别=B.类别) B," & _
                                    "      (Select B.编码,B.名称" & _
                                    "      FROM 指标主表 A,指标体系对照表 B" & _
                                    "      Where A.名称='收费项目等级' and A.类别=B.类别) C" & _
                                    " Where A.特殊病 =B.编码(+) And A.项目等级=C.编码(+)" & _
                                    " AND (" & zlCommFun.GetLike("A", "编码", strText) & " Or " & zlCommFun.GetLike("A", "名称", strText) & " Or " & zlCommFun.GetLike("A", "助记码", strText) & ")"
                            End If
                            If rsTemp.State = 1 Then rsTemp.Close
                            rsTemp.Open gstrSQL, mcnYB
                        Case TYPE_重庆渝北
                                
                                strLeft = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
                                strTemp = "'" & strLeft & strText & "%'"
                                
                                gstrSQL = " select  商品代码 as 医保编码,  医院大类编码, 药品通用中文名, 药品通用英文名,商品名, 商品曾用名, 服务项目结算方式, 报销标识, 医保标识, 是否处方用药, 药品适应症, 限制医生, 审批权限, 别名, 包装规格, " & _
                                         "         最小包装单位, 最小计量单位, 每日最大用量, 指导价格, 招标价格, 基金支付限价1, 基金支付限价2, 基金支付限价3, 实际执行价格, 自付比例1, 自付比例2, 自付比例3, 自付比例4, 自付比例5, 自付比例6, 自付比例7, 自付比例8,  " & _
                                         "         自付比例9, 自付比例10, 自付比例11, 自付比例12, 医院使用状态, 中心使用状态, 标准编号,  " & _
                                         "         五笔助记码1, 五笔助记码2, 五笔助记码3, 拼音助记码1, 拼音助记码2, 拼音助记码3, 备注, 医保经办机构,机构标准编号, 医疗机构编号, " & _
                                         "          修改时间, 目录分类  " & _
                                         "  from 医保服务项目目录" & _
                                         "  where 商品代码 like " & strTemp & " Or 商品名 like " & strTemp & " Or " & _
                                         "        别名 like " & strTemp & " Or 五笔助记码1 like " & UCase(strTemp) & " Or " & _
                                         "        拼音助记码1 like " & UCase(strTemp)
                             Debug.Print Time
                            If gcnOracle_CQYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
                            Else
                                '强制使记录集为打开状态
                                gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                             Debug.Print Time
                             gstrSQL = ""
                        Case TYPE_黔南
                                
                                strLeft = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
                                strTemp = "'" & strLeft & strText & "%'"
                                
                                gstrSQL = "" & _
                                    "  Select  类别,decode(类别,'1','药品','2','诊疗','服务') as 项目类别,类别||编码 as 医保编码,名称, 英文名称,收费类别, 收费等级, 助记码, 单位, 标准价格, 支付标准, 剂型, 规格, 备注, 变更时间, 维护标志  " & _
                                    "  From 医保收费目录" & _
                                    "  Where 编码 like " & strTemp & " Or 名称 like " & strTemp & " Or " & _
                                    "        收费类别 like " & strTemp & " Or 助记码 like " & UCase(strTemp) & _
                                    "   order by 类别,编码"
                                    
                            Debug.Print Time
                            If Not gcnOracle_黔南 Is Nothing Then
                                If gcnOracle_黔南.State = adStateOpen Then
                                    rsTemp.Open gstrSQL, gcnOracle_黔南, adOpenStatic, adLockReadOnly
                                Else
                                    '强制使记录集为打开状态
                                    gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
                                    Call OpenRecordset(rsTemp, Me.Caption)
                                End If
                            Else
                                '强制使记录集为打开状态
                                gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                             Debug.Print Time
                             gstrSQL = ""
                        Case TYPE_浙江
                            strLeft = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
                            strTemp = "'" & strLeft & UCase(strText) & "%'"
                            If gcn浙江.State = 0 Then
                                Call openConn浙江
                            End If
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select AKA060 As 医保编码, AKA061 As 项目名称, AKA069 As 自付比例, AKA066 As 拼音码, AKA070 As 剂型, Decode(AKA065,'1','甲类','2','乙类','丙类') As 类别 From KA02 " & _
                                    "Where AKA060 Like " & strTemp & " Or AKA061 Like " & strTemp & " Or AKA066 Like " & strTemp
                            Else
                                gstrSQL = "Select AKA090 As 医保编码, AKA091 As 项目名称, AKA069 As 自付比例, AKA066 As 拼音码, Decode(AKA065,'1','甲类','2','乙类','丙类') As 类别 From KA03 " & _
                                    "Where AKA090 Like " & strTemp & " Or AKA091 Like " & strTemp & " Or AKA066 Like " & strTemp
                            End If
                            If gcn浙江.State = 1 Then Set rsTemp = gcn浙江.Execute(gstrSQL)
                        Case TYPE_新都
                            Dim cn新都 As New ADODB.Connection
                            
                            strLeft = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
                            strTemp = "'" & strLeft & strText & "%'"
                            
                            cn新都.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\YWCS.MDB;Persist Security Info=True;Jet OLEDB:Database Password=yhybv1.1cdb"
                            cn新都.CursorLocation = adUseClient
                            cn新都.Open

                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select ybxmbm As 医保编码,ybxmmc As 项目名称,zfbl1 As 自付比例 From KYH904 " & _
                                          "Where ybxmbm Like " & UCase(strTemp) & " Or ybxmmc Like " & UCase(strTemp)
                            Else
                                gstrSQL = "Select ybxmbm As 医保编码,ybxmmc As 项目名称,zgxj As 一级医院价格,zgxj1 As 二级医院价格,zgxj2 As 三级医院价格,zfbl1 As 自付比例 From KYH100 " & _
                                          "Where ybxmbm Like " & UCase(strTemp) & " Or ybxmmc Like " & UCase(strTemp)
                            End If
                            Set rsTemp = cn新都.Execute(gstrSQL)
                        Case TYPE_余姚
                            strLeft = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
                            strTemp = "'" & strLeft & UCase(strText) & "%'"
                            If gcn余姚.State = 0 Then
                                Call openConn余姚
                            End If
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select MedicineID As 医保编码,Name As 项目名称,DoseType As 剂型,ZFBL As 自付比例,NameJP As 拼音首码,NameWB As 五笔码 From hi_Medicine " & _
                                    "Where MedicineID Like " & strTemp & " Or Name Like " & strTemp & " Or NameJP Like " & strTemp & " Or NameWB Like " & strTemp
                            Else
                                gstrSQL = "Select DiagnoseID As 医保编码,Name As 项目名称,'' As 剂型,ZFBL As 自付比例,NameJP As 拼音首码,NameWB As 五笔码 From hi_Diagnose " & _
                                    "Where DiagnoseID Like " & strTemp & " Or Name Like " & strTemp & " Or NameJP Like " & strTemp & " Or NameWB Like " & strTemp
                            End If
                            If gcn余姚.State = 1 Then Set rsTemp = gcn余姚.Execute(gstrSQL)
                            
                        'Modified by ZYB 毕节
                        Case TYPE_毕节
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select 药品代码 AS 医保编码,中文名称 AS 项目名称,英文名称,商品名称,药品大类,所属类型,个人自付比例||'%' AS 个人自付比例,个人起付金额 From 药品目录表 A" & _
                                " Where (" & zlCommFun.GetLike("A", "药品代码", strText) & " Or " & zlCommFun.GetLike("A", "中文名称", strText) & ")"
                            Else
                                gstrSQL = "Select 诊疗项目代码 AS 医保编码,诊疗项目名称 AS 项目名称,费用类别,一级医院定额,二级医院定额,三级医院定额,个人自付比例||'%' AS 个人自付比例,个人起付金额 From 诊疗项目表 A" & _
                                " Where (" & zlCommFun.GetLike("A", "诊疗项目代码", strText) & " Or " & zlCommFun.GetLike("A", "诊疗项目名称", strText) & ")"
                            End If
                            If rsTemp.State = 1 Then rsTemp.Close
                            rsTemp.Open gstrSQL, mcnYB
                        Case Else
                            If mint险类 = TYPE_沈阳市 Then
                                Call GetItemMatchInfo
                                If int审核标志 = 1 And mint适用地区 = 0 Then
                                    MsgBox "该项目已经通过审核，不允许修改！", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                            gstrSQL = "Select 编码  医保编码,名称,简码,附注 " & _
                                     "   FROM 保险项目 WHERE 险类=" & mint险类 & " and (编码 like '" & strText & "%' or Upper(名称) like '" & str前 & UCase(strText) & "%' Or Upper(简码) like '" & str前 & UCase(strText) & "%')"
                            Call OpenRecordset(rsTemp, Me.Caption)
                    End Select
                    
                    If rsTemp.RecordCount > 0 Then
                        '出现选择器
                        If rsTemp.RecordCount >= 1 Or rsTemp.Fields.Count > 3 Then
                            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
                            blnReturn = frmListSel.ShowSelect(rsTemp, "医保编码", "医保项目选择", "请选择对应的医保项目：")
                        End If
                    Else
                        If mint险类 = TYPE_成都内江 Then
                            MsgBox "不存指定医保项目，请重输!"
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    If blnReturn = False Then
                        '记录集中没有可选择的数据
                        If rsTemp.RecordCount > 0 Then
                            '记录集有数据，但取消了选择
                            Cancel = True
                            .TxtVisible = True
                            .TxtSetFocus
                            Exit Sub
                        Else
                            If Not (mint险类 = TYPE_沈阳市 Or mint险类 = TYPE_北京 Or mint险类 = TYPE_泸州市 Or mint险类 = TYPE_毕节 Or mint险类 = TYPE_黔南 Or mint险类 = TYPE_宁海 Or mint险类 = TYPE_渝北农医) Then
                                .Text = strText
                                .TextMatrix(.Row, COL医保编码) = strText
                            Else
                                .Text = ""
                                .TextMatrix(.Row, COL医保编码) = ""
                                Cancel = True
                                Exit Sub
                            End If
                        End If
                    Else
                        '肯定是有记录集的
                        If mint险类 = TYPE_黔南 Then
                            .Text = Mid(rsTemp("医保编码"), 2)
                        Else
                            .Text = rsTemp("医保编码")
                        End If
                        Dim str限价 As String
                        Select Case mint险类
                            Case TYPE_重庆市
                                '如果是重庆医保，那对项目的价格进行判断
                                
                                str限价 = Nvl(rsTemp("限价"), "")
                                If str限价 <> "" And Val(.TextMatrix(.Row, col价格)) > 0 Then
                                    '进行了限价
                                    If str类型 = "药品" Then
                                        '药品没有特批价
                                        blnReturn = 价格判断_重庆(Val(.TextMatrix(.Row, col价格)), rsTemp("标准单价"), str限价, False, 0)
                                    Else
                                        blnReturn = 价格判断_重庆(Val(.TextMatrix(.Row, col价格)), rsTemp("标准单价"), str限价, Nvl(rsTemp("特批项目标志"), "") = "是", rsTemp("特批价"))
                                    End If
                                    If blnReturn = False Then
                                        Cancel = True
                                        .TxtVisible = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                End If
                            Case TYPE_重庆银海版
                                '如果是重庆医保，那对项目的价格进行判断
                                str限价 = Nvl(rsTemp("限价"), "")
                                If str限价 <> "" And Val(.TextMatrix(.Row, col价格)) > 0 Then
                                    '进行了限价
                                    If str类型 = "药品" Then
                                        '药品没有特批价
                                        blnReturn = 价格判断_重庆银海版(Val(.TextMatrix(.Row, col价格)), rsTemp("标准单价"), str限价, False, 0)
                                    Else
                                        blnReturn = 价格判断_重庆银海版(Val(.TextMatrix(.Row, col价格)), rsTemp("标准单价"), str限价, Nvl(rsTemp("特批项目标志"), "") = "是", rsTemp("特批价"))
                                    End If
                                    If blnReturn = False Then
                                        Cancel = True
                                        .TxtVisible = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                End If
                            Case TYPE_泸州市
                                If Nvl(rsTemp("最大价格限制"), 0) <> 0 And Val(.TextMatrix(.Row, col价格)) > 0 Then
                                    If rsTemp("最大价格限制") < Val(.TextMatrix(.Row, col价格)) Then
                                        If MsgBox("医院单价" & Format(Val(.TextMatrix(.Row, col价格)), "0.000") & _
                                            " 高于医保中心核准的价格" & Format(rsTemp("最大价格限制"), "0.000") & "，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                                            Cancel = True
                                            .TxtVisible = True
                                            .TxtSetFocus
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case TYPE_北京
                                '如果是药品项目，检查HIS的项目名称是否在药品别名中
                                Dim rsCheck As New ADODB.Recordset
                                If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                    If Not CheckTradeName(.RowData(.Row), rsTemp("医保编码")) Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                End If
                            Case TYPE_沈阳市
                                Call CheckValid(.Text)
                            Case TYPE_浙江
                                .TextMatrix(.Row, COL医保附注) = rsTemp!类别
                                .TextMatrix(.Row, col医保名称) = rsTemp("项目名称")
                            Case TYPE_余姚
                                .TextMatrix(.Row, col医保名称) = rsTemp("项目名称")
                            Case TYPE_黔南
                                '填附注
                                .TextMatrix(.Row, COL医保附注) = rsTemp("类别")
                                .TextMatrix(.Row, col医保名称) = rsTemp("名称")
                        End Select
                        If mint险类 = TYPE_黔南 Then
                            .TextMatrix(.Row, COL医保编码) = Mid(rsTemp("医保编码"), 2)
                        Else
                            .TextMatrix(.Row, COL医保编码) = rsTemp("医保编码")
                        End If
                    End If
                End If
                Call Get保险名称
                Call 标记改变
                'Modified By 朱玉宝 地区：长沙 原因：根据当前动作设置项目匹配信息
                If mint险类 = TYPE_沈阳市 Then
                    If .TextMatrix(.Row, COL医保附注) <> "" Then
                        .TextMatrix(.Row, col医保剂型) = Split(.TextMatrix(.Row, COL医保附注), "||")(3)
                    End If
                End If
                Call SetItemMatch(False)
            Else
                If .TextMatrix(.Row, COL医保编码) = "" Then
                    .TextMatrix(.Row, COL医保编码) = " "
                End If
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Cancel = True
End Sub

Private Sub 标记改变()
    '当前输入已经有效，继续看能否得到其它内容
    cmdRestore.Enabled = True
    cmdSave.Enabled = True
    
    With mshSum_S
        If Trim(.TextMatrix(.Row, COL医保编码)) = "" And Trim(.TextMatrix(.Row, col大类名称)) = "" Then
            .TextMatrix(.Row, col改变方式) = "删除"
        Else
            If Trim(.TextMatrix(.Row, col改变方式)) <> "修改" Then
                '为空，或已经是“新增”
                .TextMatrix(.Row, col改变方式) = "新增"
            End If
        End If
    End With
End Sub

Private Sub Get保险名称()
'功能：根据当前行的保险项目编码，得到其它信息
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long, lngPos As Long
    Dim str大类编码 As String, strTemp As String, varPart As Variant
    
    On Error GoTo errHandle
    With mshSum_S
        If mint险类 = TYPE_重庆市 Then
            If mcnYB.State = adStateOpen Then
                
                gstrSQL = "Select SPM 名称,'' 大类编码,'' 附注  From YPML WHERE yplsh='" & .TextMatrix(.Row, COL医保编码) & "' " & _
                           " Union All " & _
                           " Select XMMC 名称,'' 大类编码,'' 附注  From ZLXM WHERE XMLSH='" & .TextMatrix(.Row, COL医保编码) & "'"
                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
            Else
                '强制使记录集为打开状态
                gstrSQL = "select 名称,大类编码,附注 from 保险项目 where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint险类 = TYPE_重庆银海版 Then
            '调试重庆医保银海版 204-03-29
            If mcnYB.State = adStateOpen Then
                
                gstrSQL = "Select 商品名 名称,lpad(项目等级,6,'0') 大类编码,'' 附注  From 中间库_药品目录 WHERE 流水号='" & .TextMatrix(.Row, COL医保编码) & "' " & _
                           " Union All " & _
                           " Select 项目名称 名称,lpad(项目等级,6,'0') 大类编码,'' 附注  From 中间库_诊疗项目 WHERE 流水号='" & .TextMatrix(.Row, COL医保编码) & "'"
                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
            Else
                '强制使记录集为打开状态
                gstrSQL = "select 名称,大类编码,附注 from 保险项目 where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint险类 = TYPE_成都南充 Then
            gstrSQL = " Select '' DLBM," & gstrCol_ENG & _
                      " From yljcxxk " & _
                      " Where ID=" & Val(.TextMatrix(.Row, COL医保编码))
            Call OpenRecordset(rsTemp, Me.Caption)
        ElseIf mint险类 = TYPE_泸州市 Then
            gstrSQL = "SELECT A.编码 医保编码,A.名称,A.简码,A.大类编码,A.是否医保 " & _
                      "  FROM 保险项目 A " & _
                      "  WHERE A.险类=" & TYPE_泸州市 & " AND A.编码='" & .TextMatrix(.Row, COL医保编码) & "'"
            rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
        ElseIf mint险类 = type_大连市 Or mint险类 = type_大连开发区 Then
              '刘兴宏(200311)
            
            gstrSQL = "SELECT A.编码 医保编码,A.名称,A.简码,A.大类编码,B.是否医保 " & _
                      "  FROM 保险项目 A,保险支付大类 B " & _
                      "  WHERE A.大类编码=B.编码(+) and b.险类=" & cmb险类.ItemData(cmb险类.ListIndex) & " and A.险类=" & cmb险类.ItemData(cmb险类.ListIndex) & " AND A.编码='" & .TextMatrix(.Row, COL医保编码) & "'"
            OpenRecordset rsTemp, Me.Caption
        ElseIf mint险类 = TYPE_北京 Then
            gstrSQL = " SELECT 编码,名称 From 药品目录 WHERE 编码='" & .TextMatrix(.Row, COL医保编码) & "'" & _
                      " Union " & _
                      " SELECT 编码,名称 From 诊疗目录 WHERE 编码='" & .TextMatrix(.Row, COL医保编码) & "'"
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open gstrSQL, mcnYB
        ElseIf mint险类 = TYPE_重庆渝北 Then
            '20040706
            gstrSQL = " SELECT 商品代码 编码,商品名 名称 From 医保服务项目目录 WHERE 商品代码='" & .TextMatrix(.Row, COL医保编码) & "'"
            rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
        'Modified by ZYB 2004-08-17
        ElseIf mint险类 = TYPE_乐山 Then
            gstrSQL = "select substr(附注,Instr(附注,'|',1,3)+1)||'-'||名称 AS 名称,大类编码,附注 from 保险项目 where 编码='" & .TextMatrix(.Row, COL医保编码) & _
                      "' and 险类=" & cmb险类.ItemData(cmb险类.ListIndex)
            Call OpenRecordset(rsTemp, Me.Caption)
        ElseIf mint险类 = TYPE_毕节 Then
            gstrSQL = " SELECT 诊疗项目代码 编码,诊疗项目名称 名称,费用类别 大类 From 诊疗项目表 WHERE 诊疗项目代码='" & .TextMatrix(.Row, COL医保编码) & "'" & _
                      " Union " & _
                      " Select 药品代码 编码,中文名称 名称,药品大类 大类 From 药品目录表 Where 药品代码='" & .TextMatrix(.Row, COL医保编码) & "'"
            rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
        ElseIf mint险类 = TYPE_浙江 Then
            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                gstrSQL = "Select AKA060 As 医保编码, AKA061 As 项目名称, AKA069 As 自付比例, AKA066 As 拼音码, Decode(AKA065,'1','甲类','2','乙类','丙类') As 类别 From KA02 Where AKA060='" & .TextMatrix(.Row, COL医保编码) & "'"
            Else
                gstrSQL = "Select AKA090 As 医保编码, AKA091 As 项目名称, AKA069 As 自付比例, AKA066 As 拼音码, Decode(AKA065,'1','甲类','2','乙类','丙类') As 类别 From KA03 Where AKA090='" & .TextMatrix(.Row, COL医保编码) & "'"
            End If
            Set rsTemp = gcn浙江.Execute(gstrSQL)
        ElseIf mint险类 = TYPE_余姚 Then
            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                gstrSQL = "Select MedicineID As 医保编码,Name As 项目名称,DoseType As 剂型,ZFBL As 自付比例,NameJP As 拼音首码,NameWB As 五笔码 From hi_Medicine Where MedicineID='" & .TextMatrix(.Row, COL医保编码) & "'"
            Else
                gstrSQL = "Select DiagnoseID As 医保编码,Name As 项目名称,ZFBL As 自付比例,NameJP As 拼音首码,NameWB As 五笔码 From hi_Diagnose Where DiagnoseID='" & .TextMatrix(.Row, COL医保编码) & "'"
            End If
            Set rsTemp = gcn余姚.Execute(gstrSQL)
        ElseIf mint险类 = TYPE_宁海 Then
            If mcnYB.State = 1 Then
                gstrSQL = " Select YPDM 编码,ZWM 名称,PYJM 简码 From SIM_YPML " & _
                          " Where (upper(YPDM) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(ZWM) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(PYJM) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "')" & _
                          " UNION " & _
                          " Select ZLDM 编码,ZLMC 名称,PYJM 简码 From SIM_ZLML " & _
                          " Where (upper(ZLDM) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(ZLMC) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(PYJM) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "')"
                If rsTemp.State = 1 Then rsTemp.Close
                rsTemp.Open gstrSQL, mcnYB
            Else
                '强制使记录集为打开状态
                gstrSQL = "select 名称,大类编码,附注 from 保险项目 where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint险类 = TYPE_渝北农医 Then
            If mcnYB.State = 1 Then
                gstrSQL = " Select YPLSH AS 项目编码,YPMC AS 项目名称,PY AS 简码 From YPML " & _
                          " Where (upper(YPLSH) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(YPMC) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(PY) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "')" & _
                          " UNION " & _
                          " Select XMBM AS 项目编码,XMMC AS 项目名称,PY AS 简码 From ZLXM " & _
                          " Where (upper(XMBM) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(XMMC) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "' Or Upper(PY) Like '" & UCase(.TextMatrix(.Row, COL医保编码)) & "')"
                If rsTemp.State = 1 Then rsTemp.Close
                rsTemp.Open gstrSQL, mcnYB
            Else
                '强制使记录集为打开状态
                gstrSQL = "select 名称,大类编码,附注 from 保险项目 where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint险类 = TYPE_黔南 Then
            '20040706
            gstrSQL = " SELECT 类别,编码,名称 From 医保收费目录 WHERE 类别=" & Val(.TextMatrix(.Row, COL医保附注)) & " and 编码='" & .TextMatrix(.Row, COL医保编码) & "'"
            rsTemp.Open gstrSQL, gcnOracle_黔南, adOpenStatic, adLockReadOnly
        Else
            gstrSQL = "select 名称,大类编码,附注 from 保险项目 where 编码='" & .TextMatrix(.Row, COL医保编码) & _
                      "' and 险类=" & cmb险类.ItemData(cmb险类.ListIndex)
            Call OpenRecordset(rsTemp, Me.Caption)
        End If
        
        If rsTemp.RecordCount = 0 Then
            '没有对应的保险项目，只有利用该编码
            .TextMatrix(.Row, col医保名称) = ""
            .TextMatrix(.Row, COL医保附注) = ""
            .TextMatrix(.Row, col非医保) = ""
        ElseIf mint险类 = TYPE_余姚 Then
            .TextMatrix(.Row, col医保名称) = Nvl(rsTemp!项目名称, "")
            .TextMatrix(.Row, COL医保附注) = ""
            .TextMatrix(.Row, col非医保) = ""
        ElseIf mint险类 = TYPE_浙江 Then
            .TextMatrix(.Row, col医保名称) = Nvl(rsTemp!项目名称, "")
            .TextMatrix(.Row, COL医保附注) = Nvl(rsTemp!类别, "丙类")
            .TextMatrix(.Row, col非医保) = ""
        ElseIf mint险类 = TYPE_宁海 Then
            .TextMatrix(.Row, COL医保编码) = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
        ElseIf mint险类 = TYPE_渝北农医 Then
            .TextMatrix(.Row, COL医保编码) = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
        ElseIf mint险类 = TYPE_泸州市 Then
            .TextMatrix(.Row, col医保名称) = Nvl(rsTemp("名称"))
            .TextMatrix(.Row, col非医保) = IIf(rsTemp("是否医保") = 1, "", "√")
            .TextMatrix(.Row, COL医保附注) = ""
            str大类编码 = Nvl(rsTemp("大类编码"))
        ElseIf mint险类 = TYPE_成都南充 Then
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp(ExchangeColName("名称", False))), "", rsTemp(ExchangeColName("名称", False)))
            .TextMatrix(.Row, COL医保附注) = IIf(IsNull(rsTemp(ExchangeColName("药品项目内涵", False))), "", rsTemp(ExchangeColName("药品项目内涵", False)))
        ElseIf mint险类 = type_大连市 Or mint险类 = type_大连开发区 Then
            .TextMatrix(.Row, col非医保) = IIf(rsTemp("是否医保") = 1, "", "√")
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
            str大类编码 = Nvl(rsTemp("大类编码"))
        ElseIf mint险类 = TYPE_北京 Then
            '.TextMatrix(.Row, col医保编码) = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
        ElseIf mint险类 = TYPE_重庆渝北 Then
            .TextMatrix(.Row, COL医保编码) = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
        ElseIf mint险类 = TYPE_毕节 Then
            .TextMatrix(.Row, COL医保编码) = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
            .TextMatrix(.Row, COL医保附注) = IIf(IsNull(rsTemp("大类")), "", rsTemp("大类"))
        ElseIf mint险类 = TYPE_黔南 Then
            .TextMatrix(.Row, COL医保编码) = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
            .TextMatrix(.Row, COL医保附注) = IIf(IsNull(rsTemp("类别")), "", rsTemp("类别"))
        Else
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
            .TextMatrix(.Row, COL医保附注) = IIf(IsNull(rsTemp("附注")), "", rsTemp("附注"))
            str大类编码 = IIf(IsNull(rsTemp("大类编码")), "", rsTemp("大类编码"))
            '自贡医保可以用到附注中的大类编码
            If mint险类 = TYPE_自贡市 Then
                strTemp = .TextMatrix(.Row, COL医保附注)
                strTemp = Mid(strTemp, InStr(strTemp, "|") + 1)    '去掉第一项（剂型）
                strTemp = Mid(strTemp, 1, InStr(strTemp, "|") - 1) '得到第二项（是否医保）
                .TextMatrix(.Row, col非医保) = IIf(strTemp = 0, "√", "")
            ElseIf mint险类 = TYPE_四川眉山 Then
                strTemp = .TextMatrix(.Row, COL医保附注)
                varPart = Split(strTemp, "|")
                If UBound(varPart) >= 3 Then
                    .TextMatrix(.Row, col非医保) = IIf(varPart(2) = "N", "√", "")
                Else
                    .TextMatrix(.Row, col非医保) = ""
                End If
            'Modified by 朱玉宝 20031218 地区：福州
            ElseIf mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or mint险类 = TYPE_南平市 Then
                strTemp = .TextMatrix(.Row, COL医保附注)
                varPart = Split(strTemp, "|")
                If UBound(varPart) >= 3 Then
                    .TextMatrix(.Row, col非医保) = IIf(varPart(3) = "N", "√", "")
                Else
                    .TextMatrix(.Row, col非医保) = ""
                End If
            End If
        End If
        
        For lngIndex = 0 To .ListCount - 1
            lngPos = InStr(.List(lngIndex), ".")
            If lngPos = 0 Then
                strTemp = .List(lngIndex)
            Else
                strTemp = Mid(.List(lngIndex), 1, lngPos - 1)
            End If
            If strTemp = str大类编码 Then
                '找到相匹配的大类编码
                .TextMatrix(.Row, col大类ID) = .ItemData(lngIndex)
                .TextMatrix(.Row, col大类名称) = .List(lngIndex)
                Exit For
            End If
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mshSum_S_KeyPress(KeyAscii As Integer)
    With mshSum_S
        If Not .Active Then Exit Sub
        If .ColData(.Col) = -1 Then Call 标记改变
    End With
End Sub

Private Sub mshSum_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mshSum_S.ToolTipText = mshSum_S.TextMatrix(mshSum_S.MouseRow, mshSum_S.MouseCol)
End Sub

Private Sub mshSum_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rsTemp As New ADODB.Recordset, lngID As Long
    Dim lngRow As Long, lngPos As Long, blnActive As Boolean
    Dim blnEnable As Boolean
    
    If mshSum_S.Active = False Then Exit Sub
    If mshSum_S.MouseRow = 0 Then
        If mlngCol = mshSum_S.MouseCol Then
            mblnDesc = Not mblnDesc
        Else
            mlngCol = mshSum_S.MouseCol
            mblnDesc = False
        End If
        
        blnEnable = cmdRestore.Enabled
        blnActive = mshSum_S.Active
        mshSum_S.Active = False
        mshSum_S.msfObj.MousePointer = vbHourglass
        
        '构成记录集，然后刷新表格
        rsTemp.CursorLocation = adUseClient
        rsTemp.CursorType = adOpenDynamic
        rsTemp.LockType = adLockOptimistic
        With rsTemp.Fields
            .Append "ID", adDouble, adFldIsNullable
            .Append "编码", adVarChar, 20, adFldIsNullable
            .Append "名称", adVarChar, 50, adFldIsNullable
            .Append "规格", adVarChar, 80, adFldIsNullable
            .Append "剂型", adVarChar, 30, adFldIsNullable
            .Append "产地", adVarChar, 100, adFldIsNullable
            .Append "单位", adVarChar, 20, adFldIsNullable
            .Append "是否变价", adInteger, adFldIsNullable
            .Append "价格", adVarNumeric, 20, adFldIsNullable
            .Append "改变方式", adVarChar, 4, adFldIsNullable
            'Modified By 朱玉宝 2003-12-09 地区：乐山
            .Append "项目编码", adVarChar, 50, adFldIsNullable
            .Append "项目名称", adVarChar, 50, adFldIsNullable
            .Append "附注", adVarChar, 50, adFldIsNullable
            .Append "原编码", adVarChar, 20, adFldIsNullable
            .Append "是否医保", adInteger
            .Append "大类ID", adDouble
            .Append "大类编码", adVarChar, 10, adFldIsNullable
            .Append "大类名称", adVarChar, 50, adFldIsNullable
        End With
        
        rsTemp.Open
        With mshSum_S
            For lngRow = 1 To .Rows - 1
                rsTemp.AddNew
                
                rsTemp("ID") = .RowData(lngRow)
                rsTemp("编码") = .TextMatrix(lngRow, cOL编码)
                rsTemp("名称") = .TextMatrix(lngRow, cOL名称)
                rsTemp("规格") = .TextMatrix(lngRow, COL规格)
                rsTemp("剂型") = .TextMatrix(lngRow, COL剂型)
                rsTemp("产地") = Substr(.TextMatrix(lngRow, col产地), 1, 100)
                rsTemp("单位") = .TextMatrix(lngRow, COL单位)
                If .TextMatrix(lngRow, col价格) = "" Then
                    rsTemp("是否变价") = 1
                    rsTemp("价格") = 0
                Else
                    rsTemp("是否变价") = 0
                    rsTemp("价格") = Val(.TextMatrix(lngRow, col价格))
                End If
                rsTemp("改变方式") = .TextMatrix(lngRow, col改变方式)
                rsTemp("项目编码") = .TextMatrix(lngRow, COL医保编码)
                rsTemp("项目名称") = .TextMatrix(lngRow, col医保名称)
                rsTemp("附注") = .TextMatrix(lngRow, COL医保附注)
                rsTemp("原编码") = .TextMatrix(lngRow, col原编码)
                rsTemp("大类ID") = Val(.TextMatrix(lngRow, col大类ID))
                rsTemp("是否医保") = IIf(.TextMatrix(lngRow, col非医保) = "√", 0, 1)
                
                
                lngPos = InStr(.TextMatrix(lngRow, col大类名称), ".")
                If lngPos = 0 Then
                    rsTemp("大类编码") = Null
                    rsTemp("大类名称") = Null
                Else
                    rsTemp("大类编码") = Mid(.TextMatrix(lngRow, col大类名称), 1, lngPos - 1)
                    rsTemp("大类名称") = Mid(.TextMatrix(lngRow, col大类名称), lngPos + 1)
                End If
                
                rsTemp.Update
            Next
            lngID = .RowData(.Row)
        End With
        Call FillGrid(rsTemp, lngID)
    
        mshSum_S.Active = blnActive '恢复
        mshSum_S.msfObj.MousePointer = vbDefault
        MousePointer = vbDefault
        cmdRestore.Enabled = blnEnable
        cmdSave.Enabled = blnEnable
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    '只刷新列表内容
    FillSum
End Sub

Private Sub mshSum_S_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_S_LostFocus()
    mshSum_S.CmdVisible = False
    mshSum_S.CboVisible = False
    
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + x - msngStartX
        If sngTemp > 1500 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwMain_S.SelectedItem
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    
    Set objPrint.Body = mshSum_S.msfObj
    objPrint.Title.Text = nod.Text & "类收费细目医保项目对应表"
    'objRow.Add "医院名称：" & gstr单位名称
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & gstrUserName
    objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
    
Private Sub Fill大类()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    '只刷新列表内容
    
    '首先获得医保大类
    mshSum_S.Active = True
    If gintInsure = TYPE_成都南充 Then
        If mcnYB.State = 1 Then mcnYB.Close
        mcnYB.Open GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
        Exit Sub
    End If
    
    gstrSQL = "select ID,编码,名称 from 保险支付大类 " & _
              "where 险类=" & cmb险类.ItemData(cmb险类.ListIndex) & " order by 编码"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    mshSum_S.Clear
    Do Until rsTemp.EOF
        mshSum_S.AddItem rsTemp("编码") & "." & rsTemp("名称")
        mshSum_S.ItemData(mshSum_S.NewIndex) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    
    Select Case cmb险类.ItemData(cmb险类.ListIndex)
        Case TYPE_重庆市, TYPE_重庆银海版, TYPE_北京, TYPE_毕节, TYPE_宁海
            If mcnYB.State = adStateClosed Then
                '首先读出参数，打开连接
                gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & cmb险类.ItemData(cmb险类.ListIndex)
                Call OpenRecordset(rsTemp, Me.Caption)
                Do Until rsTemp.EOF
                    strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Select Case rsTemp("参数名")
                        Case "医保服务器"
                            strServer = strTemp
                        Case "医保用户名"
                            strUser = strTemp
                        Case "医保用户密码"
                            strPass = strTemp
                    End Select
                    rsTemp.MoveNext
                Loop
                If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
                    Exit Sub
                End If
            End If
        Case TYPE_渝北农医
            Dim strDatabase As String
            If mcnYB.State = adStateClosed Then
                '首先读出参数，打开连接
                gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & cmb险类.ItemData(cmb险类.ListIndex)
                Call OpenRecordset(rsTemp, Me.Caption)
                Do Until rsTemp.EOF
                    strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Select Case rsTemp("参数名")
                        Case "医保服务器"
                            strServer = strTemp
                        Case "医保用户名"
                            strUser = strTemp
                        Case "医保用户密码"
                            strPass = strTemp
                        Case "医保实例名"
                            strDatabase = strTemp
                    End Select
                    rsTemp.MoveNext
                Loop
            End If
            If Not OpenSQLServer(mcnYB, strServer, strUser, strPass, strDatabase) Then Exit Sub
        Case TYPE_泸州市
            '不能连通医保前置机，就不能修改。因为需要保存修改记录
            If 检查医保服务器_泸州 = False Then mshSum_S.Active = False
        Case TYPE_铜仁
            '不能连通医保前置机，就不能修改。因为需要保存修改记录
            If 检查医保服务器_铜仁 = False Then mshSum_S.Active = False
        Case TYPE_重庆渝北
            If gcnOracle_CQYB Is Nothing Or gcnOracle_CQYB.State <> 1 Then
                Call 医保初始化_重庆渝北
            End If
        Case TYPE_黔南
            If gcnOracle_黔南 Is Nothing Then
                
                '重庆新打开医保
                gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & cmb险类.ItemData(cmb险类.ListIndex)
                Call OpenRecordset(rsTemp, Me.Caption)
                Do Until rsTemp.EOF
                    strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Select Case rsTemp("参数名")
                        Case "医保服务器"
                            strServer = strTemp
                        Case "医保用户名"
                            strUser = strTemp
                        Case "医保用户密码"
                            strPass = strTemp
                    End Select
                    rsTemp.MoveNext
                Loop
                Set gcnOracle_黔南 = New ADODB.Connection
                If OraDataOpen(gcnOracle_黔南, strServer, strUser, strPass) = False Then
                    Exit Sub
                End If
            End If
    End Select
End Sub

Private Function FillTree() As Boolean
'功能:装入收费类别和收费细目的所有分类到tvwMain_S
    '本程序中树节点比其它程序的KEY值多一个字符，即第二位的类别编码

    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim nod As Node
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    
    MousePointer = vbHourglass
    
    mstrKey = ""     '全面刷新时就相当于用户没点过任何节点
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    gstrSQL = "select 编码,类别 from 收费类别 where 编码<>'5' and 编码<>'6' and 编码<>'7' order by 编码"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    LockWindowUpdate tvwMain_S.hwnd
    '删除所有节点
    With tvwMain_S.Nodes
        .Clear
        '增加类别
        Do Until rsTemp.EOF
            .Add , , "R" & rsTemp("编码"), "【" & rsTemp("编码") & "】" & rsTemp("类别"), "R", "R"
            tvwMain_S.Nodes("R" & rsTemp("编码")).Sorted = True
            rsTemp.MoveNext
        Loop
        .Add , , "D5", "【5】西成药", "R", "R"
        tvwMain_S.Nodes("D5").Sorted = True
        .Add , , "E6", "【6】中成药", "R", "R"
        tvwMain_S.Nodes("E6").Sorted = True
        .Add , , "F7", "【7】中草药", "R", "R"
        tvwMain_S.Nodes("F7").Sorted = True
        
        '增加普通收费项目分类节点
        gstrSQL = "select id,上级id,类别,编码,名称 from 收费细目  where 类别<>'5' and 类别<>'6' and 类别<>'7' and 末级 <> 1 " & _
             " start with 上级ID is null  connect by prior id=上级ID "
        Call OpenRecordset(rsTemp, Me.Caption)
    
        Do Until rsTemp.EOF
            '添加节点
            If IsNull(rsTemp("上级id")) Then
                .Add "R" & rsTemp("类别"), tvwChild, "C" & rsTemp("类别") & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "C", "C"
            Else
                .Add "C" & rsTemp("类别") & rsTemp("上级id"), tvwChild, "C" & rsTemp("类别") & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "C", "C"
            End If
            tvwMain_S.Nodes("C" & rsTemp("类别") & rsTemp("ID")).Sorted = True
            rsTemp.MoveNext
        Loop
    
        '再装入药品用途分类的数据
        gstrSQL = "select id,上级id,材质,编码,名称 from 药品用途分类  " & _
             " start with 上级ID is null  connect by prior id=上级ID "
        Call OpenRecordset(rsTemp, Me.Caption)
    
        Do Until rsTemp.EOF
            '添加节点
            Select Case rsTemp("材质")
                Case "中成药"
                    If IsNull(rsTemp("上级id")) Then
                        Set nod = .Add("E6", tvwChild, "E6" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    Else
                        Set nod = .Add("E6" & rsTemp("上级id"), tvwChild, "E6" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    End If
                Case "中草药"
                    If IsNull(rsTemp("上级id")) Then
                        Set nod = .Add("F7", tvwChild, "F7" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    Else
                        Set nod = .Add("F7" & rsTemp("上级id"), tvwChild, "F7" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    End If
                Case Else '西成药
                    If IsNull(rsTemp("上级id")) Then
                        Set nod = .Add("D5", tvwChild, "D5" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    Else
                        Set nod = .Add("D5" & rsTemp("上级id"), tvwChild, "D5" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    End If
                End Select
            nod.Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    
    LockWindowUpdate 0
    MousePointer = 0
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes(1)
        nod.Selected = True
    Else
        Err.Clear
        nod.Selected = True
        nod.EnsureVisible
    End If
    Call FillSum
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
    MousePointer = 0
End Function

Public Sub FillSum(Optional ByVal blnForce As Boolean = False)
'功能:装入各种统计数据
    Dim rsTemp As New ADODB.Recordset
    Dim nod As Node
    Dim str材质分类 As String
    Dim lngID As Long

    If tvwMain_S.SelectedItem Is Nothing Then
        ClearGrid mshSum_S
        Call MenuSet
        Exit Sub
    End If
    
    If blnForce = False Then
        If mstrKey = tvwMain_S.SelectedItem.Key And mint险类 = cmb险类.ItemData(cmb险类.ListIndex) Then
            '完全没有改变，不用再刷新
            Exit Sub
        End If
        
        If cmdSave.Enabled = True Then
            If mint险类 <> TYPE_沈阳市 Then
                '已经修改，提示是否需要保存当前的设置
                If MsgBox("保险项目已经修改，是否需要保存？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    Call cmdSave_Click
                End If
            Else
                Call cmdSave_Click
            End If
        End If
    End If
    
    cmdSave = False
    cmdRestore = False
    '提取保险项目的两个主关键字
    mstrKey = tvwMain_S.SelectedItem.Key
    mint险类 = cmb险类.ItemData(cmb险类.ListIndex)
    If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
        Call InitSum
    End If
    Set nod = tvwMain_S.SelectedItem
    
    '根据不同的节点，做出不同的显示
    '手术类别要多显示一列
    If Mid(nod.Key, 2, 1) = "5" Or Mid(nod.Key, 2, 1) = "6" Or Mid(nod.Key, 2, 1) = "7" Then
        '药品的处理要麻烦一些
        mshSum_S.TextMatrix(0, col产地) = "产地"
        
        Select Case Left(nod.Key, 1)
            Case "D"
                str材质分类 = "西成药"
            Case "E"
                str材质分类 = "中成药"
            Case "F"
                str材质分类 = "中草药"
        End Select
        
        If nod.Image = "R" Then
            gstrSQL = "select A.药品ID as ID,A.编码,B.通用名称||decode(M.名称,null,'',b.通用名称,'',' 〖'||M.名称||'〗') as 名称,A.规格,A.产地,A.售价单位 as 单位,D.是否变价,E.名称 剂型 " & _
                        "from 药品目录 A,药品信息 B,收费细目 D,药品剂型 E,(Select distinct 药品id,名称 from 药品别名 ) M " & _
                        "where A.药名ID=B.药名ID and d.id=M.药品ID(+) and B.剂型=E.编码(+) and B.材质分类='" & str材质分类 & "'" & _
                        "      and A.药品ID=D.ID and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select A.药品ID as ID,A.编码,B.通用名称||decode(M.名称,null,'',b.通用名称,'',' 〖'||M.名称||'〗') as 名称,A.规格,A.产地,A.售价单位 as 单位,D.是否变价,E.名称 剂型 " & _
                      "from 药品目录 A,药品信息 B,收费细目 D,药品剂型 E,(Select distinct 药品id,名称 from 药品别名) M ,(select ID from 药品用途分类 start with ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID) C " & _
                      "where A.药名ID=B.药名ID and B.剂型=E.编码(+) and d.id=M.药品ID(+) and B.材质分类='" & str材质分类 & "' and B.用途分类ID=C.ID" & _
                      "       and A.药品ID=D.ID and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
        End If
        
    Else
        '非药品就容易得多了
        mshSum_S.TextMatrix(0, col产地) = "说明"
        
        If nod.Image = "R" Then
            gstrSQL = "select id,编码,名称,规格,说明 as 产地,计算单位 as 单位,是否变价,'' 剂型 from 收费细目 where 末级=1 and 类别='" & Mid(nod.Key, 2, 1) & "' " & _
                        " and (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select id,编码,名称,规格,说明 as 产地,计算单位 as 单位,是否变价,'' 剂型 from 收费细目 where 末级=1 and (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
                        " start with 上级ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID "
        End If
    End If
    
    'Modified by ZYB 2004-08-17
    If mint险类 = TYPE_乐山 Then
        gstrSQL = "select A.ID,A.编码,A.名称,A.规格,A.剂型,A.产地,A.单位,A.是否变价,D.价格,'' as 改变方式" & _
                   " ,B.项目编码,substr(B.附注,Instr(B.附注,'|',1,3)+1)||'-'||B.项目名称 AS 项目名称,B.附注,B.项目编码 as 原编码,B.是否医保,B.大类ID,C.编码 as 大类编码,C.名称 as 大类名称 " & _
                   " from (" & gstrSQL & ") A,保险支付项目 B,保险支付大类 C," & _
                   "      (select sum(现价) as 价格,收费细目ID from 收费价目 where 执行日期<=sysdate and (终止日期>=sysdate or 终止日期 is null) group by 收费细目ID) D " & _
                   " Where A.ID=B.收费细目ID(+) and B.大类ID=c.id(+)  and B.险类(+)= " & mint险类 & _
                   "       and A.ID=D.收费细目ID(+)  "
    Else
        gstrSQL = "select A.ID,A.编码,A.名称,A.规格,A.剂型,A.产地,A.单位,A.是否变价,D.价格,'' as 改变方式" & _
                   " ,B.项目编码,B.项目名称,B.附注,B.项目编码 as 原编码,B.是否医保,B.大类ID,C.编码 as 大类编码,C.名称 as 大类名称 " & _
                   " from (" & gstrSQL & ") A,保险支付项目 B,保险支付大类 C," & _
                   "      (select sum(现价) as 价格,收费细目ID from 收费价目 where 执行日期<=sysdate and (终止日期>=sysdate or 终止日期 is null) group by 收费细目ID) D " & _
                   " Where A.ID=B.收费细目ID(+) and B.大类ID=c.id(+)  and B.险类(+)= " & mint险类 & _
                   "       and A.ID=D.收费细目ID(+)  "
    End If
    
    MousePointer = 11
    Call OpenRecordset(rsTemp, Me.Caption)
    
    lngID = mshSum_S.RowData(mshSum_S.Row)
    Call FillGrid(rsTemp, lngID)
    
    stbThis.Panels(2).Text = "共有收费项目" & rsTemp.RecordCount & "条"
    
    MousePointer = 0
    Call MenuSet
End Sub

Private Sub FillGrid(rsTemp As ADODB.Recordset, ByVal lngID As Long)
    Dim strSort As String
    Dim strDemo As String
    Dim intMatch As Integer
    Dim lngRow As Long, lngRowSelect As Long
    
    Select Case mlngCol
        Case cOL编码
            strSort = "编码"
        Case cOL名称
            strSort = "名称"
        Case COL规格
            strSort = "规格"
        Case col产地
            strSort = "产地"
        Case COL单位
            strSort = "单位"
        Case col价格
            strSort = "价格"
        Case COL医保编码
            strSort = "项目编码"
        Case col医保名称
            strSort = "项目名称"
        Case col大类名称
            strSort = "大类名称"
        Case col非医保
            strSort = "是否医保"
        Case Else
            rsTemp.Sort = "编码"
    End Select
    rsTemp.Sort = strSort & IIf(mblnDesc, " DESC", "")
    
    mshSum_S.TxtVisible = False
    mshSum_S.CboVisible = False
    mshSum_S.Redraw = False
    ClearGrid mshSum_S
    If rsTemp.RecordCount <> 0 Then
        mshSum_S.Rows = rsTemp.RecordCount + 1
    End If
    lngRow = 1
    With mshSum_S
        Do Until rsTemp.EOF
            If rsTemp("ID") = lngID Then
                lngRowSelect = lngRow
            End If
            
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, cOL编码) = rsTemp("编码")
            .TextMatrix(lngRow, cOL名称) = rsTemp("名称")
            .TextMatrix(lngRow, COL规格) = IIf(IsNull(rsTemp("规格")), "", rsTemp("规格"))
            .TextMatrix(lngRow, col产地) = IIf(IsNull(rsTemp("产地")), "", rsTemp("产地"))
            .TextMatrix(lngRow, COL剂型) = IIf(IsNull(rsTemp("剂型")), "", rsTemp("剂型"))
            .TextMatrix(lngRow, COL单位) = IIf(IsNull(rsTemp("单位")), "", rsTemp("单位"))
            .TextMatrix(lngRow, col价格) = IIf(rsTemp("是否变价") = 0, Format(rsTemp("价格"), "0.000"), "")
            .TextMatrix(lngRow, col改变方式) = IIf(IsNull(rsTemp("改变方式")), "", rsTemp("改变方式"))
            .TextMatrix(lngRow, COL医保编码) = IIf(IsNull(rsTemp("项目编码")), "", rsTemp("项目编码"))
            .TextMatrix(lngRow, col医保名称) = IIf(IsNull(rsTemp("项目名称")), "", rsTemp("项目名称"))
            .TextMatrix(lngRow, col原编码) = IIf(IsNull(rsTemp("原编码")), "", rsTemp("原编码"))
            .TextMatrix(lngRow, col大类ID) = IIf(IsNull(rsTemp("大类ID")), "", rsTemp("大类ID"))
            .TextMatrix(lngRow, col非医保) = IIf(rsTemp("是否医保") = "0", "√", "")
            If mint险类 = TYPE_沈阳市 Then
                intMatch = 0
                strDemo = IIf(IsNull(rsTemp("附注")), "", rsTemp("附注"))
                If InStr(1, strDemo, "||") <> 0 Then
                    If InStr(1, strDemo, "^^") <> 0 Then
                        .TextMatrix(lngRow, col医保剂型) = Split(strDemo, "^^")(0)
                        .TextMatrix(lngRow, col医保剂型) = Split(.TextMatrix(lngRow, col医保剂型), "||")(3)
                        .TextMatrix(lngRow, COL医保附注) = Split(strDemo, "^^")(0)
                    Else
                        .TextMatrix(lngRow, col医保剂型) = strDemo
                        .TextMatrix(lngRow, col医保剂型) = Split(.TextMatrix(lngRow, col医保剂型), "||")(3)
                        .TextMatrix(lngRow, COL医保附注) = strDemo
                    End If
                    If InStr(1, strDemo, "^^") <> 0 Then
                        If InStr(1, Split(strDemo, "^^")(1), "||") <> 0 Then
                            .TextMatrix(lngRow, col匹配序列号) = Split(Split(strDemo, "^^")(1), "||")(0)
                            intMatch = Split(Split(strDemo, "^^")(1), "||")(1)
                        Else
                            .TextMatrix(lngRow, col匹配序列号) = Split(strDemo, "^^")(1)
                        End If
                    End If
                Else
                    .TextMatrix(lngRow, COL医保附注) = strDemo
                End If
                If intMatch = 1 Then
                    .TextMatrix(lngRow, col审核标志) = "√"
                ElseIf intMatch = 2 Then
                    .TextMatrix(lngRow, col审核标志) = "×"
                End If
            Else
                .TextMatrix(lngRow, COL医保附注) = IIf(IsNull(rsTemp("附注")), "", rsTemp("附注"))
                .TextMatrix(lngRow, col匹配序列号) = ""
            End If
            
            If IsNull(rsTemp("大类编码")) Or IsNull(rsTemp("大类名称")) Then
                .TextMatrix(lngRow, col大类名称) = ""
            Else
                .TextMatrix(lngRow, col大类名称) = rsTemp("大类编码") & "." & rsTemp("大类名称")
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    If lngRowSelect > 0 And lngRowSelect < mshSum_S.Rows - 1 Then
        mshSum_S.msfObj.TopRow = lngRowSelect
        mshSum_S.Row = lngRowSelect
    End If
    mshSum_S.Redraw = True
End Sub

Private Sub ClearGrid(objGrid As Object)
'功能：清除表格,并完成部分初始化
    Dim i As Long
    
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    With objGrid.msfObj
        .Rows = 2
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
        Next
    
    End With
End Sub

Private Sub MenuSet()
'功能:显示菜单和工具栏的状态(打印)
    Dim blnPrint As Boolean
    
    blnPrint = Not (mshSum_S.Rows = 2 And mshSum_S.TextMatrix(1, 0) = "")

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
    
    If InStr(mstr权限, "增删改") > 0 Then
        mshSum_S.Active = blnPrint
        If mint险类 = TYPE_泸州市 Then
            '强制不能使用
            If gcn泸州.State = adStateClosed Then mshSum_S.Active = False
        End If
    Else
        mshSum_S.Active = False
    End If
End Sub

Public Sub ShowForm(frmParent As Form)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select 序号,名称 from 保险类别 where nvl(是否禁止,0)<>1  order by 序号"
    Call OpenRecordset(rsTemp, "取保险类别")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm保险项目.Visible = True Then
        frm保险项目.Show
        Exit Sub
    End If
    
    mstr权限 = gstrPrivs
    frm保险项目.Show , frmParent
End Sub

'Modified By 朱玉宝 地区：长沙 原因：增加用于设置项目与医保项目的匹配
Private Sub SetItemMatch(Optional ByVal bln删除 As Boolean = True)
    '医保附注中仅保存剂型信息
    'intEdit——1新增;2修改;3删除
    'col改变方式——空或删除，执行删除匹配操作；修改执行先删除，后新增操作；执行新增操作
    Dim str匹配类型 As String, str剂型 As String, str规格 As String, str医院编码 As String
    Dim rsTemp As New ADODB.Recordset
     
    Select Case mint险类
    Case TYPE_沈阳市
        '如果已审核通过，不允许修改或删除
        If int审核标志 = 1 And mint适用地区 = 0 Then
            MsgBox "该项目已经通过医保中心审核，不允许修改或删除。请与医保中心联系！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not classInsure.InitInsure(gcnOracle, TYPE_沈阳市) Then Exit Sub
        str匹配类型 = TranClass
        str剂型 = "无"

        If Trim(mshSum_S.TextMatrix(mshSum_S.Row, col匹配序列号)) <> "" Then
            '删除匹配信息，如果不是修改，直接退出
'            1   serial_match    匹配序列号  12  否
'            2   audit_flag  审核标志    1   否  "0"：未审核；"2"：审核未通过
'            3   edit_staff  操作员工号  5   否
'            4   edit_man    操作员姓名  10  否
            If Not 调用接口_准备_沈阳市(Function_沈阳市.项目匹配_删除匹配信息) Then Exit Sub
            gstrField_沈阳市 = "serial_match||audit_flag||edit_staff||edit_man"
            gstrValue_沈阳市 = mshSum_S.TextMatrix(mshSum_S.Row, col匹配序列号) & "||" & int审核标志 & "||" & gCominfo_沈阳市.操作员工号 & "||" & gstrUserName
            If Not 调用接口_写入口参数_沈阳市(1) Then Exit Sub
            If Not 调用接口_执行_沈阳市 Then Exit Sub
            mshSum_S.TextMatrix(mshSum_S.Row, col匹配序列号) = ""
        End If
        
        If Not bln删除 Then
            '执行设置匹配动作（修改已在上面先删除了）
'            1   hospital_id医疗机构编码    20  否
'            2   match_type 匹配类型        1   否  "0"：诊疗项目匹配；"1"：西药匹配；"2"：中成药匹配；"3"：中草药匹配
'            3   hosp_code  医院目录编码    20  否
'            4   hosp_name  医院目录名称    60  否
'            5   hosp_model 医院目录剂型    20  否
'            6   price      单价            8   是
'            7   item_code  中心目录编码    20  否
'            8   item_name  中心目录名称    60  否
'            9   model_name 中心目录剂型    20  否
'            10  effect_date生效日期            否  格式:YYYY-MM-DD
'            11  expire_date失效日期            否  格式:YYYY-MM-DD
'            12  edit_staff 操作员工号      5   否
'            13  edit_man   操作员姓名      10  否
            If str匹配类型 <> "0" Then
                gstrSQL = "select C.名称 剂型  " & _
                         " from 药品信息 A,药品目录  B,药品剂型 C " & _
                         " where A.药名ID=B.药名ID And A.剂型=C.编码 And B.药品ID = " & mshSum_S.RowData(mshSum_S.Row)
                Call OpenRecordset(rsTemp, "读取药品的剂型名称")
                str剂型 = ToVarchar(rsTemp!剂型, 20)
            End If
            '取收费细目的标识码做为医院编码上传
            If Not (Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "5" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "6" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "7") Then
                gstrSQL = "Select Decode(TRIM(标识主码),NULL,编码,'',编码,标识主码) 编码,规格 From 收费细目 Where ID=" & mshSum_S.RowData(mshSum_S.Row)
            Else
                gstrSQL = "Select Decode(Trim(标识码),NULL,编码,'',编码,标识码) 编码,规格 From 药品目录 Where 药品ID=" & mshSum_S.RowData(mshSum_S.Row)
            End If
            Call OpenRecordset(rsTemp, "取医院编码")
            str医院编码 = Nvl(rsTemp!编码)
            str规格 = Nvl(rsTemp!规格)
            
            If Not 调用接口_准备_沈阳市(Function_沈阳市.项目匹配_项目匹配) Then Exit Sub
            If Not 调用接口_指定记录集_沈阳市("MatchInfo") Then Exit Sub
            
            gstrField_沈阳市 = "hospital_id||match_type||hosp_code||hosp_name||hosp_model||spec||price||" & _
            "item_code||item_name||model_name||effect_date||expire_date||edit_staff||edit_man"
            gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & str匹配类型 & "||" & _
                    str医院编码 & "||" & mshSum_S.TextMatrix(mshSum_S.Row, cOL名称) & "||" & _
                    str剂型 & "||" & str规格 & "||" & mshSum_S.TextMatrix(mshSum_S.Row, col价格) & "||" & _
                    mshSum_S.TextMatrix(mshSum_S.Row, COL医保编码) & "||" & mshSum_S.TextMatrix(mshSum_S.Row, col医保名称) & "||" & _
                    mshSum_S.TextMatrix(mshSum_S.Row, col医保剂型) & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & "||" & _
                    "3000-01-01||" & gCominfo_沈阳市.操作员工号 & "||" & gstrUserName
            If Not 调用接口_写入口参数_沈阳市(1) Then Exit Sub
            If Not 调用接口_执行_沈阳市() Then Exit Sub
            
            '获取匹配序列号，保存
            If Not 调用接口_指定记录集_沈阳市("MatchItem") Then Exit Sub
            Call 调用接口_读取数据_沈阳市("serial_match", str剂型)
            mshSum_S.TextMatrix(mshSum_S.Row, col匹配序列号) = str剂型
            
            '更新费用类型（药品才更新）
            If Not (Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "5" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "6" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "7") Then Exit Sub
            Call 调用接口_读取数据_沈阳市("Staple_flag", str剂型)
            If Val(str剂型) = 1 Then
                str剂型 = "甲类药品"
            ElseIf Val(str剂型) = 2 Then
                str剂型 = "乙类药品"
            Else
                str剂型 = "非基本药品"
            End If
            gstrSQL = "ZL_更新费用类型('" & mshSum_S.TextMatrix(mshSum_S.Row, COL医保编码) & "','" & str剂型 & "')"
            Call ExecuteProcedure("更新费用类型")
        End If
    End Select
End Sub

'Modified By 朱玉宝 地区：长沙 原因：如果是沈阳市铁路局医保，且已设置医保编码，则显示医保中心是否已设置其匹配信息
Private Sub GetItemMatchInfo()
    Dim str匹配类型 As String, str项目编码 As String, strMatch As String
    Dim int原审核标志 As Integer
    Dim rsTemp As New ADODB.Recordset
    
    int原审核标志 = IIf(mshSum_S.TextMatrix(mshSum_S.Row, col审核标志) = "√", 1, IIf(mshSum_S.TextMatrix(mshSum_S.Row, col审核标志) = "×", 2, 0))
    int审核标志 = 0
    stbThis.Panels(2).Text = ""
    If Trim(mshSum_S.TextMatrix(mshSum_S.Row, COL医保编码)) = "" Then Exit Sub
    
    If mint险类 = TYPE_沈阳市 Then

        '取收费细目的标识码做为医院编码上传
        If Not (Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "5" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "6" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "7") Then
            gstrSQL = "Select Decode(TRIM(标识主码),NULL,编码,'',编码,标识主码) 编码 From 收费细目 Where ID=" & mshSum_S.RowData(mshSum_S.Row)
        Else
            gstrSQL = "Select Decode(Trim(标识码),NULL,编码,'',编码,标识码) 编码 From 药品目录 Where 药品ID=" & mshSum_S.RowData(mshSum_S.Row)
        End If
        Call OpenRecordset(rsTemp, "取医院编码")
        str项目编码 = Nvl(rsTemp!编码)

'        1   hospital_id    医疗机构编码    20  否
'        2   his_item_code  医院目录编码    20  否
'        3   medi_item_type 匹配类型        1   否  "0"：诊疗项目匹配；"1"：西药匹配；"2"：中成药匹配；"3"：中草药匹配
'        4   fee_date       费用发生时间        否  格式：YYYY-MM-DD
        stbThis.Panels(2).Text = "获取该项目的匹配信息失败！"
        If Not classInsure.InitInsure(gcnOracle, TYPE_沈阳市) Then Exit Sub
        If Not 调用接口_准备_沈阳市(Function_沈阳市.项目匹配_取单个项目匹配信息) Then Exit Sub
        str匹配类型 = TranClass
        gstrField_沈阳市 = "hospital_id||his_item_code||medi_item_type||fee_date"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & Trim(str项目编码) & "||" & _
                str匹配类型 & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-DD")
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Sub
        If Not 调用接口_执行_沈阳市() Then Exit Sub
        '指定记录集
        If Not 调用接口_指定记录集_沈阳市("MatchInfo") Then Exit Sub
        Call 调用接口_读取数据_沈阳市("audit_flag", strMatch)
        Call DebugTool("审核标志：" & strMatch)
        If strMatch = "" Then strMatch = "0"
        int审核标志 = Val(strMatch)
        
        If int审核标志 = 1 Then
            mshSum_S.TextMatrix(mshSum_S.Row, col审核标志) = "√"
        ElseIf int审核标志 = 2 Then
            mshSum_S.TextMatrix(mshSum_S.Row, col审核标志) = "×"
        Else
            mshSum_S.TextMatrix(mshSum_S.Row, col审核标志) = ""
        End If
        stbThis.Panels(2).Text = "匹配信息：" & IIf(strMatch = "0", "未审核", IIf(strMatch = "1", "审核通过", "审核未通过"))
        
        '更新保险支付项目
        If int审核标志 <> int原审核标志 Then Call 标记改变
    End If
End Sub

'Modified By 朱玉宝 地区：长沙 原因：转换类别为医保接口需要的匹配类型
Private Function TranClass() As String
    Dim strClass As String
    strClass = Mid(tvwMain_S.SelectedItem.Key, 2, 1)
    Select Case strClass
    Case "5"
        TranClass = "1"
    Case "6"
        TranClass = "2"
    Case "7"
        TranClass = "3"
    Case Else
        TranClass = "0"
    End Select
End Function

Private Function CheckValid(ByVal strCode As String) As Boolean
    Dim str大类 As String
    Dim rsTemp As New ADODB.Recordset
    '检查大类是否匹配
    gstrSQL = "Select 附注 From 保险项目 Where 编码='" & strCode & "'"
    Call OpenRecordset(rsTemp, "获取大类")
    str大类 = Mid(rsTemp!附注, 1, 1)
    
    If str大类 <> TranClass Then
        MsgBox "请注意：该医保项目所属类别与当前选择的医院项目的类别不同！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function
