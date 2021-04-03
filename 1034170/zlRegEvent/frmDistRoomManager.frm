VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmDistRoomManager 
   Caption         =   "门诊分诊管理"
   ClientHeight    =   6105
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10530
   Icon            =   "frmDistRoomManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   3810
      TabIndex        =   2
      Top             =   0
      Width           =   3810
      Begin VB.TextBox txtValue 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   615
         TabIndex        =   3
         ToolTipText     =   "定位F3"
         Top             =   30
         Width           =   3165
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   0
         TabIndex        =   4
         Top             =   15
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483634
      End
   End
   Begin VB.Timer tmrBrush 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   120
      Top             =   435
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5748
      Width           =   10524
      _ExtentX        =   18574
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDistRoomManager.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13494
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   1605
      Left            =   1005
      TabIndex        =   1
      Top             =   2565
      Width           =   4290
      _Version        =   589884
      _ExtentX        =   7567
      _ExtentY        =   2831
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmDistRoomManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModul As Long, mstrFindKey As String
Private mbytViewScrop(0 To 3) As Byte  '0-显示已分诊病人;1-显示已接诊病人;2-显示已完成病人;3-显示不就诊病人
Private mblnCard As Boolean     '是否刷卡
Private mobjFindKey As CommandBarPopup
Private WithEvents mfrmTriageMgr  As frmTriageManager
Attribute mfrmTriageMgr.VB_VarHelpID = -1
Private mobjQueue As zlQueueManage.clsQueueManage
Private mstrQueuePrivs As String '排队叫号虚拟模块权限
Private mlngTimerState As Boolean        '临时存放timer状态的变量
Private mbln缺省读卡 As Boolean
Private Enum pg_Page
    pg_分诊页 = 1
    pg_排队页 = 2
End Enum
Private Type ty_Para
        str分诊科室 As String
        int分诊有效天数 As Integer
        byt排队叫号模式 As Byte '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
        byt呼叫站点 As Byte   '0-代表分诊台分诊呼叫;1-代表医生主动呼叫
        bln分诊呼叫 As Boolean
        blnAutoRefresh  As Boolean
        strcurQueueName As String '当前队列名称
        lngcurQueue业务ID As Long     '当前队列业务ID
        str临床部门 As String
        byt分诊台签到排队 As Byte  ';1-分诊签到生成;0-挂号时生成
        byt候诊排序方式 As Byte  '候诊病人的排序方式,0-科室编码,号码,单据号;1-科室编码,号码,挂号时间;
End Type
Private mcllFilter As Variant
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mTy_Para As ty_Para

'-----------------------------------------------------------------------------------
'消息相关变量
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1
Private mstrRegistIdsed As String '已经刷新的挂号ID,用逗号分离
Private mblnExistNewMsg As Boolean    '是否存在新消息
'-----------------------------------------------------------------------------------
'结算卡相关
Private mcllBrushCard As Collection
Private mstrCaption As String
Private mintFindType As Integer

Private Type ty_Square
    lng缺省卡类别ID As Long
    lng卡类别ID  As Long
    bln卡号密文 As Boolean
    int医疗卡长度 As Long
End Type

Private mty_Square As ty_Square

'-----------------------------------------------------------------------------------
Private Sub ClearMenuItem()
    '删除现在的工具栏及顶级菜单项
    Dim lngCount As Long
    For lngCount = cbsThis.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsThis.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsThis.Count To 2 Step -1
        cbsThis(lngCount).Delete
    Next
End Sub


Public Function zlDefCommandBars() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化菜单及工具栏
    '返回：设置成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-01 11:04:33
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrCustom As CommandBarControlCustom
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar, i As Long, strKey As String
    
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BillPrint, "重打排队单(&R)"): cbrControl.BeginGroup = True
         '77412:李南春，2014/9/3,门诊病人条码打印
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BarcodePrint, "条码打印(&B)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "病人签到(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "取消签到(&X)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Triage, "分诊(&M)"): cbrControl.BeginGroup = True 'Ctrl+T
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_ChangeNum, "换号(&C)") 'CTRL+M
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_Leave, "病人不就诊(&L)")
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_Wait, "病人待诊(&W)")
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_BackHospitalize, "回诊(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_BackHospitalizeCancel, "取消回诊(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "完成就诊(&O)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "取消完成(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPati, "病案信息(&I)"): cbrControl.BeginGroup = True 'Ctrl+I
        '73743:李南春,2014-7-21,病人基本信息调整
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPatiBaseInfo, "病人基本信息调整(&D)")
    End With
 
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "条件过滤(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conmenu_View_TriagePati, "显示已分诊病人(&1)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conmenu_View_AdmissionsPati, "显示已接诊病人(&2)")
        Set cbrControl = .Add(xtpControlButton, conmenu_View_OverPati, "显示已完成病人(&3)")
        Set cbrControl = .Add(xtpControlButton, conmenu_View_Leave, "显示不就诊病人(&4)")
        
        Set cbrControl = .Add(xtpControlButton, conmenu_View_AutoRefresh, "自动刷新(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '主菜单右侧的查找
    Set cbrCustom = cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    cbrCustom.Handle = picSearch.Hwnd
    cbrCustom.flags = xtpFlagRightAlign
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("T"), conMenu_Edit_Triage    '分诊
        .Add FCONTROL, Asc("M"), conmenu_Edit_ChangeNum '换号
        .Add FCONTROL, Asc("I"), conMenu_Edit_ModiyPati     '病人信息
        .Add FCONTROL, Asc("F"), conMenu_View_Filter     '条件过滤
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F3, conMenu_View_Find
    End With
    
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "病人签到"): cbrControl.BeginGroup = True
       'Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "取消签到")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Triage, "分诊"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_BackHospitalize, "病人回诊"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "完成就诊"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPati, "病案")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlGetDept() As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取部门信息
    '返回:部门信息IDs:如:123;234;24
    '编制：刘兴洪
    '日期：2010-06-11 20:40:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo Hd
    If InStr(mstrPrivs, "所有科室") > 0 Then
        Set rsTemp = GetDepartments("'临床'", "1,3")
    Else
        strSQL = _
            " Select A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And B.工作性质='临床' And B.服务对象 IN(1,3)" & _
            " And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    End If
    With rsTemp
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & "," & Nvl(rsTemp!ID)
            .MoveNext
        Loop
    End With
    If strSQL <> "" Then zlGetDept = Mid(strSQL, 2)
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub zlRefreshQueueData()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重新获取队列数据
    '编制：刘兴洪
    '日期：2010-06-02 17:53:32
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, rsTemp As ADODB.Recordset, strSQL As String
    Dim strTemp As String
    Dim strQueue() As String, i As Long
    If mobjQueue Is Nothing Or mTy_Para.byt排队叫号模式 = 0 Then Exit Sub
    If Not (InStr(mstrQueuePrivs, ";基本;") > 0) Then Exit Sub
    
    strTemp = IIf(mTy_Para.str分诊科室 = "", mTy_Para.str临床部门, mTy_Para.str分诊科室)
    varData = Split(strTemp, ",")
    i = UBound(varData) + 1
    ReDim Preserve strQueue(1 To i) As String
    For i = 0 To UBound(varData)
        strQueue(i + 1) = varData(i)
    Next
    '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    'zlRefresh(cnOracle As ADODB.Connection, str队列名称() As String, strCurrent队列名称 As String, lngCurrentWorkID As Long) As Long
    '功能:调用刷新指定医嘱id的报告内容，并根据情况提供编辑功能
    '参数:  lngOrderId-医嘱id;
    '返回:成功返回0,否则返回错误代码
    Call mobjQueue.zlRefresh(strQueue, mTy_Para.strcurQueueName, mTy_Para.lngcurQueue业务ID)
End Sub


Private Sub InitVar(Optional blnPatiSet As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化相关变量和参数
    '入参:
    '编制：刘兴洪
    '日期：2010-06-01 16:25:23
    '------------------------------------------------------------------------------------------------------------------------
    Dim Curdate As Date, byt排队叫号模式 As Boolean
    Dim bytNoDay As Byte
   
    byt排队叫号模式 = mTy_Para.byt排队叫号模式
    mstrQueuePrivs = ";" & GetPrivFunc(glngSys, 1160) & ";"
    
    mTy_Para.str分诊科室 = zlDatabase.GetPara("分诊科室", glngSys, mlngModul)
    mTy_Para.int分诊有效天数 = zlDatabase.GetPara("分诊有效天数", glngSys, mlngModul, "1")  '问题:27600
    mTy_Para.byt排队叫号模式 = Val(zlDatabase.GetPara("排队叫号模式", glngSys, mlngModul))
    mTy_Para.byt呼叫站点 = Val(zlDatabase.GetPara("排队呼叫站点", glngSys, mlngModul))
    mTy_Para.bln分诊呼叫 = Val(zlDatabase.GetPara("分诊后立即呼叫", glngSys, mlngModul)) = 1
    mTy_Para.blnAutoRefresh = Val(zlDatabase.GetPara("自动刷新", glngSys, mlngModul, 0)) = 1
    mTy_Para.byt分诊台签到排队 = Val(zlDatabase.GetPara("分诊台签到排队", glngSys, mlngModul, 0))
    mTy_Para.byt候诊排序方式 = Val(zlDatabase.GetPara("候诊排序方式", glngSys, mlngModul, 0)) '候诊病人的排序方式,0-科室编码,号码,单据号;1-科室编码,号码,挂号时间;
    
    mbytViewScrop(0) = IIf(Val(zlDatabase.GetPara("显示分诊病人", glngSys, mlngModul, 0)) = 1, 1, 0)
    mbytViewScrop(1) = IIf(Val(zlDatabase.GetPara("显示在诊病人", glngSys, mlngModul, 0)) = 1, 1, 0)
    mbytViewScrop(2) = IIf(Val(zlDatabase.GetPara("显示已诊病人", glngSys, mlngModul, 0)) = 1, 1, 0)
    mbytViewScrop(3) = IIf(Val(zlDatabase.GetPara("显示不就诊病人", glngSys, mlngModul, 0)) = 1, 1, 0)

    If mTy_Para.str分诊科室 = "" Then
        mTy_Para.str临床部门 = zlGetDept()
    End If
    
    Curdate = zlDatabase.Currentdate
    Set mcllFilter = New Collection
    bytNoDay = IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    
    mcllFilter.Add Array(Format(DateAdd("D", -1 * bytNoDay, Curdate), "yyyy-mm-dd 00:00:00"), Format(Curdate, "yyyy-mm-dd 23:59:59")), "挂号时间"
    mcllFilter.Add Array("", ""), "挂号NO"
    mcllFilter.Add Array("", ""), "发票号"
    mcllFilter.Add "", "挂号员"
    mcllFilter.Add "", "科室"
    mcllFilter.Add "", "门诊号": mcllFilter.Add "", "就诊卡号"
    mcllFilter.Add "", "医保号": mcllFilter.Add "", "病人姓名"
    mcllFilter.Add 0, "KIND"
    mcllFilter.Add 0, "病人ID"
    mcllFilter.Add "  And A.发生时间 Between [1] And [2]", "条件"
    mfrmTriageMgr.zlSetFilterCons mcllFilter
    Call mfrmTriageMgr.zlSetViewScrop(0, mbytViewScrop(0))
    Call mfrmTriageMgr.zlSetViewScrop(1, mbytViewScrop(1))
    Call mfrmTriageMgr.zlSetViewScrop(2, mbytViewScrop(2))
    Call mfrmTriageMgr.zlSetViewScrop(3, mbytViewScrop(3))
    
    mfrmTriageMgr.zl分诊科室 = mTy_Para.str分诊科室
    mfrmTriageMgr.zl有效天数 = mTy_Para.int分诊有效天数
    tmrBrush.Enabled = mTy_Para.blnAutoRefresh
    Call mfrmTriageMgr.zlInitVar(Me, mTy_Para.byt分诊台签到排队, mTy_Para.byt候诊排序方式)
    If blnPatiSet And byt排队叫号模式 <> mTy_Para.byt排队叫号模式 Then
        Call Check排队叫号
        Call InitPage: cbsThis.RecalcLayout
    End If
End Sub

Private Sub InitPage()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化页面
    '编制：刘兴洪
    '日期：2010-06-01 16:12:58
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    
    Err = 0: On Error GoTo Errhand:
    Call tbPage.RemoveAll
    Set ObjItem = tbPage.InsertItem(pg_Page.pg_分诊页, "分诊管理", mfrmTriageMgr.Hwnd, 0)
    ObjItem.Tag = pg_Page.pg_分诊页
    '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    If Not mobjQueue Is Nothing And InStr(mstrQueuePrivs, ";基本;") > 0 And mTy_Para.byt排队叫号模式 <> 0 Then
        Set ObjItem = tbPage.InsertItem(pg_排队页, "排队叫号", mobjQueue.zlGetForm.Hwnd, 0)
        ObjItem.Tag = pg_排队页
    End If
    
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 Private Sub SubPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Select Case tbPage.Selected.Tag
    Case pg_Page.pg_分诊页
        mfrmTriageMgr.zlSubPrint (bytMode)
    Case pg_Page.pg_排队页
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call SubPrint(2)
    Case conMenu_File_Print: Call SubPrint(1)
    Case conMenu_File_Excel: Call SubPrint(3)
    Case conMenu_Manage_Plan '签到
        Call mfrmTriageMgr.zlExc签道(False)
        Call zlRefreshQueueData
    Case conMenu_File_BillPrint '排队单打印
            Call mfrmTriageMgr.zlRePrintBill
    '77412:李南春，2014/9/3,门诊病人条码打印
    Case conMenu_File_BarcodePrint
        Call mfrmTriageMgr.zlPrintBarcode
    Case conMenu_Manage_Logout '取消签到
        Call mfrmTriageMgr.zlExc签道(True)
        Call zlRefreshQueueData
    Case conmenu_Edit_BackHospitalize  '回诊
        Call mfrmTriageMgr.zlExc回诊(False)
        Call zlRefreshQueueData
    Case conmenu_Edit_BackHospitalizeCancel '取消回诊
        Call mfrmTriageMgr.zlExc回诊(True)
        Call zlRefreshQueueData
    Case conMenu_Edit_Triage   ' 分诊
        Call mfrmTriageMgr.zlExecuteTriage(Me)
    Case conmenu_Edit_ChangeNum    '变号
        Call mfrmTriageMgr.zlExcuteChangeNum(Me)
    Case conMenu_Edit_ModiyPati  '调整病人信息
        Call mfrmTriageMgr.zlExcuteEditPati(Me)
    '73743:李南春,2014-7-3,病人基本信息调整
    Case conMenu_Edit_ModiyPatiBaseInfo  '病人基本信息调整
        Call mfrmTriageMgr.zlModiyPatiBaseInfo(Me)
    Case conmenu_Edit_Leave  '病人不就诊
        Call mfrmTriageMgr.zlExcutePatiLeave(Me)
    Case conmenu_Edit_Wait '病人待诊
        Call mfrmTriageMgr.zlExcutePatiWait(Me)
    Case conMenu_Manage_Finish '完成就诊
        Call zlExcutePatiOver: Call tmrBrush_Timer
    Case conMenu_Manage_Redo  '恢复就诊
         Call mfrmTriageMgr.zlExcutePatiCancelOver(Me): Call tmrBrush_Timer
    Case conmenu_View_TriagePati     '显示分诊病人
        mbytViewScrop(0) = IIf(mbytViewScrop(0) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(0, mbytViewScrop(0), True)
        zlDatabase.SetPara "显示分诊病人", mbytViewScrop(0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    Case conmenu_View_AdmissionsPati    '显示在诊病人
        mbytViewScrop(1) = IIf(mbytViewScrop(1) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(1, mbytViewScrop(1), True)
        zlDatabase.SetPara "显示在诊病人", mbytViewScrop(1), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    Case conmenu_View_OverPati    '显示已就诊病人
        mbytViewScrop(2) = IIf(mbytViewScrop(2) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(2, mbytViewScrop(2), True)
        zlDatabase.SetPara "显示已诊病人", mbytViewScrop(2), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    Case conmenu_View_Leave    '显示不就诊病人
        mbytViewScrop(3) = IIf(mbytViewScrop(3) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(3, mbytViewScrop(3), True)
        zlDatabase.SetPara "显示不就诊病人", mbytViewScrop(3), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    Case conmenu_View_AutoRefresh    '自动刷新
        
        mTy_Para.blnAutoRefresh = Not mTy_Para.blnAutoRefresh
        zlDatabase.SetPara "自动刷新", IIf(mTy_Para.blnAutoRefresh, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        tmrBrush.Enabled = mTy_Para.blnAutoRefresh
        Call zlRefreshData
    Case conMenu_View_Refresh   '刷新
        Call zlRefreshData
    Case conMenu_View_Filter  '过滤
        Call zlSetFilterCons
    Case conMenu_View_StatusBar
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
        cbsThis.RecalcLayout
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Parameter: Call zlParaSet
    Case conMenu_View_Find
           If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
    Case conMenu_File_Exit: Unload Me
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                mfrmTriageMgr.zlExcuteReport Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1))
        Else
             If Check排队叫号 Then mobjQueue.zlExecuteCommandBars Control
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    cbsThis.GetClientRect Left, Top, Right, Bottom
   With tbPage
        .Left = Left
        .Top = Top
        .Width = Right - Left
        .Height = Bottom - Top
   End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnTriagePage As Boolean
    
    Err = 0: On Error Resume Next
    blnTriagePage = pg_Page.pg_分诊页 = Val(tbPage.Selected.Tag)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.index
        Case conMenu_EditPopup
          Control.Visible = blnTriagePage
        End Select
    End If
    Select Case Control.ID
    Case conMenu_View_Refresh
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Visible = blnTriagePage
            Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsHaveData
    Case conMenu_Manage_Plan
            If Check排队叫号 Then
                 Control.Visible = mTy_Para.byt分诊台签到排队 = 1 And blnTriagePage
                 Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs允许签道
            Else
                Control.Visible = False
            End If
    Case conMenu_File_BillPrint '重打排队单
            Control.Visible = InStr(1, mstrPrivs, ";分诊排队单;") > 0 And blnTriagePage
    '77412:李南春，2014/9/3,门诊病人条码打印
    Case conMenu_File_BarcodePrint '条码打印
            Control.Visible = InStr(1, mstrPrivs, ";条码打印;") > 0 And blnTriagePage
    Case conMenu_Manage_Logout  '取消签到
            If Check排队叫号 Then
                 Control.Visible = mTy_Para.byt分诊台签到排队 = 1 And blnTriagePage
                 Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs允许取消签道
            Else
                Control.Visible = False
            End If
    Case conmenu_Edit_BackHospitalize   '回诊
            If Check排队叫号 Then
                Control.Visible = blnTriagePage
                Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs允许回诊
            Else
                Control.Visible = False: Control.Enabled = False
            End If
    Case conmenu_Edit_BackHospitalizeCancel  '取消回诊
            If Check排队叫号 Then
                Control.Visible = blnTriagePage
                Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs允许取消回诊
            Else
                Control.Visible = False: Control.Enabled = False
            End If
    Case conMenu_Edit_Triage, conmenu_Edit_ChangeNum  ' 分诊,换号
            Control.Visible = blnTriagePage
            Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsTriage
    Case conMenu_Edit_ModiyPati  '调整病人信息
            Control.Visible = blnTriagePage And InStr(mstrPrivs, ";病案修改;") > 0
    '73743:李南春,2014-7-21,病人基本信息调整
    Case conMenu_Edit_ModiyPatiBaseInfo  '调整病人基本信息
        Control.Visible = blnTriagePage And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;") > 0
    Case conmenu_Edit_Leave  '病人不就诊
        Control.Visible = blnTriagePage
         Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiLeave
    Case conmenu_Edit_Wait '病人待诊
        Control.Visible = blnTriagePage
         Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiWait
    Case conMenu_Manage_Finish '完成就诊
        Control.Visible = blnTriagePage And InStr(mstrPrivs, "完成就诊") > 0 '只有"完成就诊"的才可以进行标注就诊完成功能
        Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiFinish
    Case conMenu_Manage_Redo  '恢复就诊
        Control.Visible = blnTriagePage And InStr(mstrPrivs, "完成就诊") > 0   '只有"完成就诊"的才可以进行标注就诊完成功能
        Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiReDo
    Case conMenu_EditPopup  '编辑
        Control.Visible = blnTriagePage
    Case conmenu_View_TriagePati    '显示已分诊病人
        Control.Checked = (mbytViewScrop(0) = 1)
    Case conmenu_View_AdmissionsPati    '显示已接诊病人
        Control.Checked = (mbytViewScrop(1) = 1)
    Case conmenu_View_OverPati    '显示已完成病人
        Control.Checked = (mbytViewScrop(2) = 1)
    Case conmenu_View_Leave    '显示不就诊病人
        Control.Checked = (mbytViewScrop(3) = 1)
    Case conmenu_View_AutoRefresh   '自动刷新
        If Not IsStartMsgModule Then    '直接调用,不存在性能问题(已经咨询过程福荣)
            Control.Checked = mTy_Para.blnAutoRefresh
        Else
            '启用了消息平台,不允许设置自动刷新
            Control.Visible = False
        End If
    Case conMenu_View_LocationItem, conMenu_View_Find '只有分诊页面才存在.
        Control.Visible = blnTriagePage
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case conMenu_View_FindType       '指定数据
        If Control.Parent Is cbsThis.ActiveMenuBar Then
            Control.Caption = "" & mstrCaption & "↓"
        End If
        Control.Visible = blnTriagePage '42532
    Case Else
             If Check排队叫号 Then mobjQueue.zlUpdateCommandBars Control
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        DoEvents
        If txtValue.Visible = True And txtValue.Enabled Then
            Call txtValue.SetFocus
        End If
    Else
        IDKind.ActiveFastKey
    End If
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Val(tbPage.Selected.Tag) = pg_Page.pg_排队页 Then Exit Sub
    
    If KeyAscii = vbKeyReturn And Not Me.ActiveControl Is txtValue Then
        Call mfrmTriageMgr.zlExcuteFunction
    End If
End Sub
Private Sub Form_Load()
    Err = 0: On Error Resume Next
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.Hwnd)
    Call mobjICCard.SetParent(Me.Hwnd)
    Set mfrmTriageMgr = New frmTriageManager
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitVar
    Call Check排队叫号
    Call zlDefCommandBars
    Call InitPage
    Call InitIDKind
    Call zlRefreshQueueData
'    问题108110,多次调用刷新分诊列表
'    Call zlRefreshData
    
    '初始化消息发送对送
    Call InitMsgModule
    Call mfrmTriageMgr.zlSetobjMsgModule(mobjMsgModule)
End Sub
 
Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    If txtValue.Visible Then txtValue.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtValue.Text = objPatiInfor.卡号
    Call zlRefreshData(True, Trim(txtValue.Text), 1)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    tmrBrush.Enabled = False
    Call SaveWinState(Me, App.ProductName)
    Err = 0: On Error Resume Next
    If Not mobjIDCard Is Nothing Then
         Call mobjIDCard.SetEnabled(False)
         Set mobjIDCard = Nothing
     End If
     If Not mobjICCard Is Nothing Then
         Call mobjICCard.SetEnabled(False)
         Set mobjICCard = Nothing
     End If
    zlDatabase.SetPara "自动刷新", IIf(mTy_Para.blnAutoRefresh, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "显示分诊病人", mbytViewScrop(0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "显示在诊病人", mbytViewScrop(1), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "显示已诊病人", mbytViewScrop(2), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "显示不就诊病人", mbytViewScrop(3), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    If Not mfrmTriageMgr Is Nothing Then Unload mfrmTriageMgr
    Set mfrmTriageMgr = Nothing
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    '拆卸消息发送对象
    Call UnloadMsgModule
End Sub

 
Private Sub mfrmTriageMgr_zlPopuMenu(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub
 

Private Sub mfrmTriageMgr_zlQueueAsk(intType As Integer, strNO As String, lng病人ID As Long, Cancel As Boolean)
  '------------------------------------------------------------------------------------------------------------------------
    '功能：功能操作后,呼
    '入参：intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-病人取消就诊,7-回诊
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-06-03 14:15:46
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte
    Err = 0: On Error GoTo Errhand: '48792
    If Check排队叫号 = False Then Exit Sub
    
    strSQL = "SELECT ID,执行部门ID,诊室,执行人,nvl(病人ID,0) as 病人ID  From 病人挂号记录 where NO=[1] and 记录性质=1 and 记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    strQueueName = Nvl(rsTemp!执行部门id)
'    If Nvl(rsTemp!执行人) <> "" Then
'        strQueueName = strQueueName & ":" & Nvl(rsTemp!执行人)
'    ElseIf Nvl(rsTemp!诊室) <> "" Then
'        strQueueName = strQueueName & ":" & Nvl(rsTemp!诊室)
'    End If
    lngID = Val(Nvl(rsTemp!ID))
    Select Case intType
    Case 1  '-分诊;
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
        If mTy_Para.bln分诊呼叫 = False Then Exit Sub
        mobjQueue.zlQueueExec strQueueName, 0, lngID, IIf(mTy_Para.byt排队叫号模式 = 2, 5, 1)
    Case 2  '换号
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
        If mTy_Para.bln分诊呼叫 = False Then Exit Sub
        mobjQueue.zlQueueExec strQueueName, 0, lngID, IIf(mTy_Para.byt排队叫号模式 = 2, 5, 1)
    Case 3   ' 病人不就诊;
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 3
    Case 4, 6   '病人待诊,'病人取消就诊
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 0
    Case 5  '病人完成就诊
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    Case 7  '回诊
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 6
    End Select
    Call zlRefreshQueueData
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub mfrmTriageMgr_zlShowInfor(strShowInfor As String)
    Me.stbThis.Panels(2).Text = strShowInfor
End Sub

 Private Sub zlParaSet()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：参数设置
    '编制：刘兴洪
    '日期：2010-06-01 15:47:06
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    frmDistPara.mstrPrivs = mstrPrivs
    frmDistPara.mlngModul = mlngModul
    mlngTimerState = IIf(tmrBrush.Enabled, 1, 0): tmrBrush.Enabled = False
    
    frmDistPara.Show 1, Me
    Call InitVar(True)
    Call zlRefreshData
    tmrBrush.Enabled = mlngTimerState = 1
End Sub

Private Sub zlSetFilterCons()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置过滤条件
    '编制：刘兴洪
    '日期：2010-06-01 16:00:34
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Variant
    If mTy_Para.blnAutoRefresh Then
        If MsgBox("自动刷新状态不允许条件过滤。" & vbCrLf & "现在禁止自动刷新吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        mTy_Para.blnAutoRefresh = False:  tmrBrush.Enabled = False
    End If
    Set cllFilter = mcllFilter
    If frmDistFilter.zlShowMe(Me, mlngModul, cllFilter) = False Then
        Exit Sub
    End If
    Set mcllFilter = cllFilter
    txtValue.Text = ""
    Call mfrmTriageMgr.zlSetFilterCons(cllFilter)
    
    mfrmTriageMgr.zlintFindKeys = mintFindType
    Call zlRefreshData(True)
End Sub
 
Private Sub mobjMsgModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '存在消息接收数据
    Dim strRegistIds As String, strRegisterID As String, strRegisterDeptdId  As String
    
    If mblnExistNewMsg Then Exit Sub '有新消息,就不用再确定,直接退出
    If UCase(strMsgItemIdentity) <> "ZLHIS_REGIST_001" Then Exit Sub
    If strMsgContent = "" Then Exit Sub
    If mfrmTriageMgr Is Nothing Then Exit Sub
    
    If Val(tbPage.Selected.Tag) = pg_Page.pg_排队页 Then
        strRegistIds = "," & mobjQueue.GetQueueBusinessDataIDs() & ","
    Else
        strRegistIds = "," & mfrmTriageMgr.zlGetRegistIDsed & ","
    End If
    
    If zlXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
    If zlXML.GetSingleNodeValue("register_info/register_id", strRegisterID) = False Then Exit Sub
    If zlXML.GetSingleNodeValue("register_info/register_dept_id", strRegisterDeptdId) = False Then Exit Sub

    If InStr(1, strRegistIds, "," & Val(strRegisterDeptdId) & ",") = 0 _
        And (InStr(1, "," & mTy_Para.str分诊科室 & ",", "," & strRegisterDeptdId & ",") = True _
              Or mTy_Para.str分诊科室 = "") Then
            mblnExistNewMsg = True
    End If
End Sub

Private Sub picSearch_Resize()
    Err = 0: On Error Resume Next
    With picSearch
        txtValue.Width = .ScaleWidth - IDKind.Width
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
      Dim cbrControl As CommandBarControl, cbrCustom As CommandBarControlCustom
      Dim i As Long
    
        Call LockWindowUpdate(Me.Hwnd)
        Call ClearMenuItem
        Call zlDefCommandBars
'        Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
        If Check排队叫号 Then GoTo GoEnd:
        
        If Val(tbPage.Selected.Tag) = pg_Page.pg_排队页 Then
            '加载排队队列信息
            Call mobjQueue.zlDefCommandBars(cbsThis)
            For i = 1 To cbsThis.Count
                If i <> 1 Then
                    For Each cbrControl In cbsThis(i).Controls
                        cbrControl.Style = xtpButtonIconAndCaption
                    Next
                End If
            Next
        End If
GoEnd:
    Call LockWindowUpdate(0)
    Call zlRefreshData
End Sub

 
Private Sub txtValue_Change()
    If Me.ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_GotFocus()
    Call zlControl.TxtSelAll(txtValue)
    Call zlCommFun.OpenIme(True)
    If txtValue.Text = "" And ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '0-门诊号;1-姓名;2-挂号单;3-就诊卡号;4-医保号
        If IDKind.GetCurCard.名称 = "挂号单" And txtValue.Text <> "" Then txtValue.Text = GetFullNO(txtValue.Text, 12)
        Call zlRefreshData(True, Trim(txtValue.Text))
        zlControl.TxtSelAll txtValue
    End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    '0-门诊号,1-姓名,2-挂号单,3-就诊卡号,4-医保号
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    strKind = IDKind.GetCurCard.名称
    txtValue.PasswordChar = IIf(IDKind.GetCurCard.卡号密文规则 <> "", "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtValue.IMEMode = 0
    
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, gobjSquare.bln缺省卡号密文)
        intLen = gobjSquare.int缺省卡号长度
    Case "门诊号"
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "挂号单"
    Case "医保号"
    Case Else
            If IDKind.GetCurCard.接口序号 <> 0 Then
                blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, IDKind.GetCurCard.卡号密文规则 <> "")
                intLen = IDKind.GetCurCard.卡号长度
            End If
    End Select
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtValue.Text) = intLen - 1 And KeyAscii <> 8 Then
        If KeyAscii <> 13 Then
            txtValue.Text = txtValue.Text & Chr(KeyAscii)
            txtValue.SelStart = Len(txtValue.Text)
        End If
        KeyAscii = 0: mblnCard = True
         Call zlRefreshData(True, Trim(txtValue.Text), 1)
        mblnCard = False
        zlControl.TxtSelAll txtValue
   End If
End Sub
Private Sub txtvalue_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)
    txtValue.Text = Trim(txtValue.Text)
End Sub
Private Sub tmrBrush_Timer()
    Static intNum As Integer
    If IsStartMsgModule Then
        '1.连接成功的,需要1分钟才能刷新一次
        '2.并且需要存在新消息时,才能刷新
        intNum = intNum + 1
        If intNum >= 2 Then '每在30秒执行一次,二次为1分钟
           intNum = 0
           If mblnExistNewMsg Then
                mblnExistNewMsg = False
                Call zlRefreshData
           End If
        End If
    Else
        intNum = 0
        Call zlRefreshData
    End If
End Sub

Private Sub zlRefreshData(Optional blnFilter As Boolean = False, _
    Optional strFindValue As String = "", Optional bytReadType As Byte = 0)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重新刷新数据
    '入参：blnFilter-是否过滤
    '          bytReadType-读取类型(0-不区分;1-刷卡;2-读取身份证;3-读取IC卡)
    '编制：刘兴洪
    '日期：2010-06-02 09:43:08
    '------------------------------------------------------------------------------------------------------------------------
    mlngTimerState = Me.tmrBrush.Enabled: Me.tmrBrush.Enabled = False
    If Val(tbPage.Selected.Tag) = pg_Page.pg_排队页 Then
        Call zlRefreshQueueData
    Else
        mfrmTriageMgr.zlintFindKeys = mintFindType
        Call mfrmTriageMgr.zlRefreshData(blnFilter, strFindValue, bytReadType, IDKind.GetCurCard)
    End If
    Me.tmrBrush.Enabled = mlngTimerState
End Sub
Public Sub zlExcutePatiOver()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：完成就诊
    '编制：刘兴洪
    '日期：2010-05-31 15:52:52
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strMsgbox As String, lng病人ID As Long, lng执行状态 As Long
    Dim strNO As String, str缺省诊室 As String, str缺省医生 As String
    Dim rsTmp As ADODB.Recordset, lngID As Long
    Dim i As Long, strSQL As String
    
    If InStr(mstrPrivs, "完成就诊") = 0 Then Exit Sub
    lng病人ID = mfrmTriageMgr.zlGet病人ID
    If lng病人ID = 0 Then
        MsgBox "不存在的病人！", vbInformation, gstrSysName: Exit Sub
    End If
    lng执行状态 = mfrmTriageMgr.zlGet挂号执行状态
    If lng执行状态 = 1 Then Exit Sub
    If lng执行状态 = 2 Then
        strMsgbox = "医生已经对该病人接诊，正常情况应由医生确定完成！" & vbCrLf & _
                    "除非情况特殊(如医生因故在完成前外出无法继续接诊)" & vbCrLf & _
                    "否则，建议不要进行该操作！" & vbCrLf & vbCrLf & _
                    "真的要标记完成吗？"
        If MsgBox(strMsgbox, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    strNO = mfrmTriageMgr.zlGet挂号NO: str缺省医生 = mfrmTriageMgr.zlGet挂号医生
    str缺省诊室 = mfrmTriageMgr.zlGet挂号诊室
    lngID = mfrmTriageMgr.zlGet挂号ID
    
    On Error GoTo errHandle
    If frmDistOver.zlShowEdit(Me, mstrPrivs, mstrQueuePrivs, mobjQueue, mlngModul, strNO, lng病人ID, str缺省诊室, str缺省医生, mTy_Para.byt排队叫号模式, lngID) = False Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check排队叫号() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查和创建排队叫号功能
    '返回：排队叫号功能所有的都合法,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-06 10:19:43
    '说明：需检查: 权限合法检查;启用了排队叫号的;创建排队叫号成功!
    '------------------------------------------------------------------------------------------------------------------------
    '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    If mTy_Para.byt排队叫号模式 = 0 Then GoTo GoEnd:
    If Not (InStr(mstrQueuePrivs, ";基本;") > 0) Then GoTo GoEnd:
    Err = 0: On Error GoTo GoEnd:
    If mobjQueue Is Nothing Then
        Set mobjQueue = CreateObject("zlQueueManage.clsQueueManage")
        mobjQueue.zlInitVar gcnOracle, glngSys, 0, mTy_Para.int分诊有效天数, mstrQueuePrivs, ""
    End If
    Check排队叫号 = True
    Exit Function
GoEnd:
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    Set mobjQueue = Nothing
End Function

'Private Sub InitMenus()
'    Dim varData As Variant, varTemp As Variant, strKind As String
'    Dim i As Long
'
'    Set mcllBrushCard = New Collection
'    strKind = "姓|姓名|0|0|" & zlGetPatiInforMaxLen.intPatiName & "|0|0||"
'    strKind = strKind & ";" & "门|门诊号|0|0|18|0|0||"
'    strKind = strKind & ";" & "挂|挂号单|0|0|18|0|0||"
'    strKind = strKind & ";" & "就|就诊卡|0|0|18|0|0||"
'    strKind = strKind & ";" & "医|医保号|0|0|64|0|0||"
'    strKind = strKind & ";" & "身|身份证号|0|0|18|0|0||"
'    strKind = strKind & ";" & "IC|IC卡号|0|0|50|0|0||"
'    If Not gobjSquare.objSquareCard Is Nothing Then
'        strKind = gobjSquare.objSquareCard.zlGetIDKindStr(strKind)
'    End If
'    varData = Split(strKind, ";")
'    For i = 0 To UBound(varData)
'        varTemp = Split(varData(i), "|")
'        '取缺省的刷卡方式
'        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
'        '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
'        '第7位后,就只能用索引,不然取不到数
'        mcllBrushCard.Add varTemp, varTemp(1)
'        If Val(varTemp(5)) = 1 Then
'            gobjSquare.bln缺省卡号密文 = Trim(varTemp(7)) <> ""
'            mty_Square.lng缺省卡类别ID = Val(varTemp(3))
'            gobjSquare.int缺省卡号长度 = Val(varTemp(4))
'            mbln缺省读卡 = Val(varTemp(2)) = 1
'        End If
'    Next
'    Call InitCardType
'End Sub

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;挂|挂号单|0", txtValue)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
End Function


Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtValue.Text = "" And Me.ActiveControl Is txtValue Then
        txtValue.Text = strID:
        If txtValue.Text = "" Then
            Call mobjIDCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        '读取类型(0-不区分;1-刷卡;2-读取身份证;3-读取IC卡)
        Call zlRefreshData(True, Trim(txtValue.Text), 2)
    End If
End Sub
Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    If txtValue.Text = "" And Me.ActiveControl Is txtValue Then
        txtValue.Text = strNO
        If txtValue.Text = "" Then
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        '读取类型(0-不区分;1-刷卡;2-读取身份证;3-读取IC卡)
        Call zlRefreshData(True, Trim(txtValue.Text), 3)
    End If
End Sub
 
Private Sub InitMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化消息模块
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    Call IsStartMsgModule   '设置自动刷新
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub UnloadMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拆卸消息模块
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    If mobjMsgModule Is Nothing Then Exit Sub
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Function IsStartMsgModule() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否启用了消息模块对象的(包含连接成功)
    '返回:存在消息模块对象且连接成功的返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-11 14:42:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjMsgModule Is Nothing Then Exit Function
    If mobjMsgModule.IsConnect = False Then Exit Function
    If tmrBrush.Enabled = False Then tmrBrush.Enabled = True
    IsStartMsgModule = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
