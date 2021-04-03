VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProcEdit 
   Caption         =   "编辑过程"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   Icon            =   "frmProcEdit.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   13470
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Index           =   0
      Left            =   105
      ScaleHeight     =   795
      ScaleWidth      =   13185
      TabIndex        =   9
      Top             =   480
      Width           =   13185
      Begin VB.Frame fra 
         Height          =   930
         Left            =   210
         TabIndex        =   10
         Top             =   -75
         Width           =   12930
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   4
            Left            =   10290
            Picture         =   "frmProcEdit.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   150
            Width           =   300
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   150
            Width           =   1410
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   11370
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   150
            Width           =   1380
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1080
            TabIndex        =   13
            Top             =   510
            Width           =   11625
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   2
            Left            =   3480
            TabIndex        =   12
            Text            =   "cbo"
            Top             =   150
            Width           =   7065
         End
         Begin VB.TextBox txt 
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   3465
            TabIndex        =   16
            Top             =   150
            Width           =   6810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "过程类型："
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   20
            Top             =   210
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "过程名称："
            Height          =   180
            Index           =   1
            Left            =   2580
            TabIndex        =   19
            Top             =   210
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "所有者："
            Height          =   180
            Index           =   2
            Left            =   10680
            TabIndex        =   18
            Top             =   210
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "过程说明："
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   17
            Top             =   570
            Width           =   900
         End
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2730
      Index           =   1
      Left            =   360
      ScaleHeight     =   2730
      ScaleWidth      =   3090
      TabIndex        =   7
      Top             =   7230
      Width           =   3090
      Begin RichTextLib.RichTextBox txtProgramEdit 
         Height          =   2175
         Left            =   270
         TabIndex        =   8
         Top             =   210
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   3836
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmProcEdit.frx":D0A4
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2955
      Index           =   2
      Left            =   195
      ScaleHeight     =   2955
      ScaleWidth      =   4410
      TabIndex        =   4
      Top             =   1500
      Width           =   4410
      Begin SHDocVwCtl.WebBrowser wbr 
         Height          =   1935
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
         ExtentX         =   3413
         ExtentY         =   3413
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin RichTextLib.RichTextBox txtLastProgram 
         Height          =   1875
         Left            =   135
         TabIndex        =   6
         Top             =   105
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   3307
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmProcEdit.frx":D141
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Index           =   4
      Left            =   7335
      ScaleHeight     =   2010
      ScaleWidth      =   3030
      TabIndex        =   2
      Top             =   4995
      Visible         =   0   'False
      Width           =   3030
      Begin XtremeSuiteControls.TabControl tbcBase 
         Height          =   1425
         Index           =   1
         Left            =   315
         TabIndex        =   3
         Top             =   255
         Width           =   1995
         _Version        =   589884
         _ExtentX        =   3519
         _ExtentY        =   2514
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Index           =   3
      Left            =   7860
      ScaleHeight     =   2010
      ScaleWidth      =   3030
      TabIndex        =   0
      Top             =   1815
      Visible         =   0   'False
      Width           =   3030
      Begin XtremeSuiteControls.TabControl tbcBase 
         Height          =   1425
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   465
         Width           =   1995
         _Version        =   589884
         _ExtentX        =   3519
         _ExtentY        =   2514
         _StockProps     =   64
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmProcEdit.frx":D1DE
      Left            =   780
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmProcEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjMain As Object
Private mblnReading As Boolean

Private mlngKey As Long
Private mblnDataChanged As Boolean
Private mlngSelStart As Long
Private mblnStartUp As Boolean
Private mrsProcedure As ADODB.Recordset
Private mfrmProcedureRelating As frmProcRelating
Private mstrReportsFile As String
Private mcnOracle As ADODB.Connection
Private mfrmProcedureOwnerCon As frmProcOwnerConn
Private mblnOK As Boolean
Private mintState As Integer
Private mblnLast As Boolean
Private mblnThis As Boolean

Private Sub InitCommandBar()
    '******************************************************************************************************************
    '功能：初始菜单工具栏
    '参数：无
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SaveExit, "完成")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Save, "暂存")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    
End Sub

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane1 As Pane
    Dim objPane2 As Pane
    Dim objPane3 As Pane
    Dim objPane4 As Pane

    Set objPane1 = dkpMain.CreatePane(1, 300, 115, DockLeftOf, objPane1)
    objPane1.Title = "基本信息"
    objPane1.Options = PaneNoCaption

    Set objPane2 = dkpMain.CreatePane(2, 300, 300, DockBottomOf, objPane1)
    objPane2.Title = "过程上次"
    objPane2.Options = PaneNoCaption
    
    Set objPane2 = dkpMain.CreatePane(3, 300, 300, DockBottomOf, objPane2)
    objPane2.Title = "本次过程"
    objPane2.Options = PaneNoCaption
    
'    Set objPane4 = dkpMain.CreatePane(3, 100, 300, DockRightOf, objPane2)
'    objPane4.Title = "说明"
'    objPane4.Options = PaneNoCaption
    
    
    
        
    dkpMain.SetCommandBars cbsMain
    Call gclsBase.DockPannelInit(dkpMain)

End Sub


Public Function ShowMe(ByVal objMain As Object, ByVal lngKey As Long) As Boolean
    Set mobjMain = objMain
    mlngKey = lngKey
    mblnOK = False
    
    Me.Show 1, objMain
    
    ShowMe = mblnOK
    
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    Dim objItem As Object
    Dim lngLoop As Long
    Dim lngKey As Long
    Dim rsSQL As ADODB.Recordset
    Dim strArr() As String
    Dim strFlag As String
    Dim strTemp As String
    Dim blnValidate As Boolean
    Dim strOld As String
    Dim strNew As String
    
    Dim strNewProcedure As String
    Dim strTmpReports As String
    Dim strStandard As String
    Dim strThisProcedure As String
    Dim blnFlag As Boolean
    Dim objFSO As TextStream
    Dim strProcedureName As String
    Dim lngTemp As Long
    Dim lngProcess  As Long
    Dim strUserPwd As String
    Dim strError As String
    
    On Error GoTo errHand
    
    mblnReading = True
    
    Call gclsBase.SQLRecord(rsSQL)
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        Call InitCommandBar
        Call InitDockPannel
        Call gclsBase.TabControlInit(tbcBase(0))
        With tbcBase(0)
            
            With .PaintManager
                .Appearance = xtpTabAppearanceVisio
                .Color = xtpTabColorOffice2003
                .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
                .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            End With
            .InsertItem(0, "上次过程", picPane(2).hwnd, 1).Tag = "上次过程"
        End With
        
        Call gclsBase.TabControlInit(tbcBase(1))
        With tbcBase(1)
            With .PaintManager
                .Appearance = xtpTabAppearanceVisio
                .Color = xtpTabColorOffice2003
                .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
                .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            End With
            .InsertItem(0, "本次过程", picPane(1).hwnd, 1).Tag = "本次过程"
        End With
        txtLastProgram.BackColor = &HE0E0E0
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        txt(0).MaxLength = gclsBase.GetMaxLength("zlProcedure", "名称")
        txt(1).MaxLength = gclsBase.GetMaxLength("zlProcedure", "说明")
        With cbo(0)
            .Clear
            .AddItem "1-变动过程"
            .ItemData(.NewIndex) = 1
            .AddItem "2-空白过程"
            .ItemData(.NewIndex) = 2
            .AddItem "3-用户过程"
            .ItemData(.NewIndex) = 3
            .ListIndex = -1
        End With
        strSQL = "Select Distinct 所有者 from zlSystems a"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
        With cbo(1)
            If rs.BOF = False Then
                .AddItem "--所有者--"
                .ItemData(.NewIndex) = 0
                For lngLoop = 1 To rs.RecordCount
                    .AddItem Nvl(rs("所有者").value)
                    .ItemData(.NewIndex) = lngLoop
                    rs.MoveNext
                Next
                .ListIndex = -1
            End If
        End With
        If mlngKey = 0 Then
            '加载标准过程
        End If
        cbo(2).Text = ""
        cbo(2).Visible = False
        txt(0).Text = "ZLUSER_"
        txt(0).Visible = True
        txt(0).SetFocus
        txt(0).SelStart = Len(txt(0).Text)
        
        '---------------------------------------------------------------------------
        If mlngKey = 0 Then
            Call gclsBase.SetPaneRange(dkpMain, 2, picPane(2).ScaleWidth, 0, picPane(2).ScaleWidth, 0)
            Call Form_Resize
            GoTo errNext
        End If
        strNewProcedure = App.Path & "\NewProcedure"
        strTmpReports = App.Path & "\Reports"
        strStandard = App.Path & "\Standard"
        strThisProcedure = App.Path & "\ThisProcedure"
        mstrReportsFile = strTmpReports
        
        '将本次自定过程对应的上次标准过程与本次标准过程进行比较
        If gobjFile.FolderExists(strStandard) Then
            Call gobjFile.DeleteFolder(strStandard)
        End If
        If gobjFile.FolderExists(strNewProcedure) Then
            Call gobjFile.DeleteFolder(strNewProcedure)
        End If
        If gobjFile.FolderExists(strTmpReports) Then
            Call gobjFile.DeleteFolder(strTmpReports)
        End If
        If gobjFile.FolderExists(strThisProcedure) Then
            Call gobjFile.DeleteFolder(strThisProcedure)
        End If
        DoEvents
        
        Call gobjFile.CreateFolder(strStandard)
        Call gobjFile.CreateFolder(strNewProcedure)
        Call gobjFile.CreateFolder(strTmpReports)
        Call gobjFile.CreateFolder(strThisProcedure)
        
        DoEvents
        
        '
        mblnLast = False
        mblnThis = False
        blnFlag = False
        
        strSQL = "Select A.ID,A.类型,A.名称,B.内容 From zlProcedure A,zlProceduretext B Where A.ID = B.过程ID And A.ID=[1] And B.性质 = " & ProcTextType.上次标准过程 & " Order By B.序号"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
        If rs.BOF = False Then
            strTemp = ""
            Do While Not rs.EOF
                strTemp = strTemp & Nvl(rs("内容").value)
                rs.MoveNext
            Loop
            rs.MoveFirst
            strOld = strStandard & "\" & Nvl(rs("名称").value) & ".sql"
            Set objFSO = gobjFile.CreateTextFile(strOld)
            objFSO.Write strTemp
            objFSO.Close
            mblnLast = True
            strProcedureName = Nvl(rs("名称").value)
        Else
            blnFlag = False
        End If
        
        
        strSQL = "Select A.ID,A.类型,A.名称,B.内容 From zlProcedure A,zlProceduretext B Where A.ID = B.过程ID And A.ID=[1] And B.性质 = " & ProcTextType.本次标准过程 & " Order By B.序号"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
        If rs.BOF = False Then
            strTemp = ""
            Do While Not rs.EOF
                strTemp = strTemp & Nvl(rs("内容").value)
                rs.MoveNext
            Loop
            rs.MoveFirst
            strNew = strNewProcedure & "\" & Nvl(rs("名称").value) & ".sql"
            Set objFSO = gobjFile.CreateTextFile(strNew)
            objFSO.Write strTemp
            objFSO.Close
            blnFlag = True
            mblnThis = True
            strProcedureName = Nvl(rs("名称").value)
        Else
            blnFlag = False
        End If
        
        strSQL = "Select A.ID,A.类型,A.名称,B.内容 From zlProcedure A,zlProceduretext B Where A.ID = B.过程ID And A.ID=[1] And B.性质 = " & ProcTextType.本次自定过程 & " Order By B.序号"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
        If rs.BOF = False Then
            strTemp = ""
            Do While Not rs.EOF
                strTemp = strTemp & Nvl(rs("内容").value)
                rs.MoveNext
            Loop
            rs.MoveFirst
            strNew = strThisProcedure & "\" & Nvl(rs("名称").value) & ".sql"
            Set objFSO = gobjFile.CreateTextFile(strNew)
            objFSO.Write strTemp
            objFSO.Close
            blnFlag = True
            mblnThis = True
            strProcedureName = Nvl(rs("名称").value)
        Else
            blnFlag = False
        End If
        
        If blnFlag = False Then
            Call gclsBase.SetPaneRange(dkpMain, 3, picPane(2).ScaleWidth, 0, picPane(2).ScaleWidth, 0)
            Call Form_Resize
            GoTo errNext
        End If
'        strProcedureName = NVL(rs("名称").value)
        
        If mblnLast = True And mblnThis = True Then
            
            '调用第三方工具对比两个文件夹
            'wincmp3 d:\zlsoft\source\10.29.90\Zl_入院病案主页_Insert.txt /=old d:\zlsoft\source\10.29.50 \ Zl_入院病案主页_Insert.txt /=new /N /R1 /1
'            strCommand = App.Path & "\CompareIt\wincmp3.exe " & strOld & " /=old " & strNew & " /=new" & " /N /R1 /1 /G:HE " & strTmpReports
            
'            strCommand = App.Path & "\CompareIt\wincmp3.exe " & strStandard & "\" & " " & strNewProcedure & "\" & " /G:HE " & strTmpReports
'            lngTemp = Shell(strCommand, vbHide)
'            DoEvents
'            If Err <> 0 Then
'                Err.Clear
'                MsgBox "文件比较失败，请检查工具及文件是否正确", vbExclamation, "中联软件"
'                Exit Function
'            End If
'            lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
'            Do
'                Sleep 100
'                GetExitCodeProcess lngProcess, lngTemp
'            Loop While lngTemp = Still_Active
'            Err.Clear
'            DoEvents
            Call CheckFile(strStandard, strNewProcedure, strTmpReports)
        End If
        
        '--------------------------------------------------------------------------------------------------------------
        If gobjFile.FileExists(strTmpReports & "\" & strProcedureName & ".sql.htm") Then
            wbr.Visible = True
            txtLastProgram.Visible = False
            tbcBase(0).Item(0).Caption = "上次标准过程(左) 与 本次标准过程(右)差异对比"
            Call wbr.Navigate(strTmpReports & "\" & strProcedureName & ".sql.htm")
        Else
            Call gobjFile.DeleteFolder(strTmpReports)
            Call gobjFile.CreateFolder(strTmpReports)
            Call CheckFile(strNewProcedure, strThisProcedure, strTmpReports)
            If gobjFile.FileExists(strTmpReports & "\" & strProcedureName & ".sql.htm") Then
                wbr.Visible = True
                txtLastProgram.Visible = False
                tbcBase(0).Item(0).Caption = "本次标准过程(左) 与 本次自定过程(右)差异对比"
                Call wbr.Navigate(strTmpReports & "\" & strProcedureName & ".sql.htm")
                GoTo 0
            End If
            strSQL = "Select A.ID,B.内容 From zlProcedure A,zlProcedureText B Where A.ID = B.过程ID And A.ID=[1] And B.性质=" & ProcTextType.上次标准过程 & " Order By B.序号"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
            If rs.BOF = False Then
                wbr.Visible = False
                txtLastProgram.Visible = True
                strTemp = ""
                Do While Not rs.EOF
                    strTemp = strTemp & Nvl(rs("内容").value)
                    rs.MoveNext
                Loop
                txtLastProgram.Text = strTemp
            Else
                strSQL = "Select A.ID,B.内容 From zlProcedure A,zlProcedureText B Where A.ID = B.过程ID And A.ID=[1] And B.性质=" & ProcTextType.本次标准过程 & " Order By B.序号"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
                If rs.BOF = False Then
                    wbr.Visible = False
                    txtLastProgram.Visible = True
                    strTemp = ""
                    Do While Not rs.EOF
                        strTemp = strTemp & Nvl(rs("内容").value)
                        rs.MoveNext
                    Loop
                    txtLastProgram.Text = strTemp
                Else
                    Call gclsBase.SetPaneRange(dkpMain, 3, picPane(2).ScaleWidth, 0, picPane(2).ScaleWidth, 0)
                    Call Form_Resize
                End If
            End If
        End If
0:
        On Error Resume Next
        objFSO.Close
        On Error GoTo errHand
        If gobjFile.FolderExists(strNewProcedure) Then
            Call gobjFile.DeleteFolder(strNewProcedure)
        End If
        If gobjFile.FolderExists(strStandard) Then
            Call gobjFile.DeleteFolder(strStandard)
        End If
        If gobjFile.FolderExists(strThisProcedure) Then
            Call gobjFile.DeleteFolder(strThisProcedure)
        End If
        
errNext:
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        If mlngKey = 0 Then GoTo EndHand
        
        strSQL = "Select A.ID,A.名称,A.所有者,A.类型,A.说明,B.内容,A.状态 From zlProcedure A,zlProcedureText B Where A.ID = B.过程ID(+) And A.ID=[1] And B.性质=3 Order By B.序号"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
        If rs.BOF = False Then
            cbo(2).Visible = False
            txt(0).Visible = True
'            cbo(1).Text = NVL(rs("所有者").value)
            Call gclsBase.CboLocate(cbo(1), Nvl(rs("所有者").value, 0), False)
            txt(0).Text = Nvl(rs("名称").value)
            txt(1).Text = Nvl(rs("说明").value)
            strTemp = ""
            Do While Not rs.EOF
                strTemp = strTemp & Nvl(rs("内容").value)
                rs.MoveNext
            Loop
            txtProgramEdit.Text = strTemp
            rs.MoveFirst
            Call gclsBase.CboLocate(cbo(0), Nvl(rs("类型").value, 1), True)
            mintState = Nvl(rs("状态").value, 1)
        End If
        If cbo(0).ItemData(cbo(0).ListIndex) = 1 Or cbo(0).ItemData(cbo(0).ListIndex) = 2 Then
            cbo(1).Enabled = False
        End If
        cbo(0).Locked = True
        txt(0).Locked = True
    '------------------------------------------------------------------------------------------------------------------
    Case "加载过程列表"
        With cbo(2)
            If mrsProcedure Is Nothing Then
                strSQL = "Select distinct Name From all_source Where type in ('PROCEDURE','FUNCTION') and Owner In (Select distinct 所有者 From Zlsystems a) Order By Name"
                Set mrsProcedure = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
            End If
            
            If mrsProcedure.BOF = False Then
                mrsProcedure.MoveFirst
                Do While Not mrsProcedure.EOF
                    .AddItem mrsProcedure("Name").value
                    mrsProcedure.MoveNext
                Loop
                .ListIndex = 0
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "校验数据"
        
        If ValidData = False Then
            Exit Function
        End If
        strFlag = txtProgramEdit.Text
        '变动过程和空白过程
        If (cbo(0).ItemData(cbo(0).ListIndex) = 1 Or cbo(0).ItemData(cbo(0).ListIndex) = 2) Then
            If mlngKey > 0 Then
                strTemp = gclsBase.GetSource(Trim(txt(0).Text))
            Else
                strTemp = gclsBase.GetSource(Trim(cbo(2).Text))
            End If
            If strTemp = "" Then
                MsgBox "该过程不是变动过程！", vbInformation + vbOKOnly, "中联软件"
                Exit Function
            End If
            '检查是否修改了参数
            If gclsBase.CheckRule(strFlag, strTemp) = False Then
                MsgBox "变动过程或空白过程不允许改过程参数！", vbInformation + vbOKOnly, "中联软件"
                Exit Function
            End If
        End If
        '检查过程名称是否匹配
        If gclsBase.CheckProgramName(IIf(Trim(txt(0).Text = ""), Trim(cbo(2).Text), Trim(txt(0).Text)), strFlag) = False Then
            MsgBox "编辑区域过程名称不匹配！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
        '检查过程有效性
        '判断当前登录用户是否与过程所有者匹配
        strSQL = "Select User From Dual"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
        If rs.BOF = False Then
            If Trim(cbo(1).Text) <> Nvl(rs("User").value) Then
                '重新连接
                Set mcnOracle = New ADODB.Connection
                If Contains(gcolOwnerConn, "K" & Trim(cbo(1).Text)) Then
                    Set mcnOracle = gcolOwnerConn.Item("K" & Trim(cbo(1).Text))
                    If mcnOracle.State = adStateClosed Then mcnOracle.Open
                Else
                    If mfrmProcedureOwnerCon Is Nothing Then
                        Set mfrmProcedureOwnerCon = New frmProcOwnerConn
                    End If
                    If mfrmProcedureOwnerCon.ShowDialog(Me, Trim(cbo(1).Text), strUserPwd, blnValidate) Then
                        If blnValidate = False Then GoTo EndHand
                        If mcnOracle.State = adStateOpen Then mcnOracle.Close
                        mcnOracle.Provider = "MSDataShape"
                        mcnOracle.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServer, Trim(cbo(1).Text), TranPasswd(strUserPwd)
                        If mcnOracle.State = adStateClosed Then mcnOracle.Open
                        
                        If err <> 0 Then
                            '保存错误信息
                            strError = err.Description
                            If InStr(strError, "自动化错误") > 0 Then
                                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
                            ElseIf InStr(strError, "ORA-12154") > 0 Then
                                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
                            ElseIf InStr(strError, "ORA-12541") > 0 Then
                                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
                            ElseIf InStr(strError, "ORA-01033") > 0 Then
                                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
                            ElseIf InStr(strError, "ORA-01034") > 0 Then
                                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
                            ElseIf InStr(strError, "ORA-02391") > 0 Then
                                MsgBox "用户" & UCase(cbo(1).Text) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
                            ElseIf InStr(strError, "ORA-01017") > 0 Then
                                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
                            ElseIf InStr(strError, "ORA-28000") > 0 Then
                                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
                            Else
                                MsgBox strError, vbInformation, gstrSysName
                            End If
                            err = 0
                            Exit Function
                        End If
                        gcolOwnerConn.Add mcnOracle, "K" & Trim(cbo(1).Text)
                    Else
                            GoTo EndHand
                    End If
                End If
            End If
        End If
        
        
'        Select Case gclsBase.CheckProgram(IIf(Trim(txt(0).Text) = "", Trim(cbo(2).Text), Trim(txt(0).Text)), txtProgramEdit.Text, mcnOracle)
'        Case 0
'        Case 1
'            MsgBox "过程/函数不能被创建，请检查！", vbInformation + vbOKOnly, "中联软件"
'            Exit Function
'        Case 2
'            MsgBox "过程/函数能被创建，但过程/函数无效，请检查！", vbInformation + vbOKOnly, "中联软件"
'            'Exit Function
'        End Select
        
'        '用户过程
'        If cbo(0).ItemData(cbo(0).ListIndex) = 3 Then
'            strTemp = gclsbase.GetSource(Trim(txt(0).Text))
'            If strTemp <> "" Then
'                If gclsbase.CheckRule(strFlag, strTemp) = False Then
'                    If mfrmProcedureRelating Is Nothing Then
'                        Set mfrmProcedureRelating = New frmProcedureRelating
'                    End If
'                    Call mfrmProcedureRelating.ShowDialog(Me, mlngKey)
'                End If
'            End If
'        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"
        '保存数据
        If mlngKey = 0 Then
        '新增
            lngKey = gclsBase.GetNextId("zlProcedure")
            If cbo(0).ItemData(cbo(0).ListIndex) = 1 Or cbo(0).ItemData(cbo(0).ListIndex) = 2 Then
                strTemp = gclsBase.GetSource(Trim(cbo(2).Text))
            Else
                strTemp = gclsBase.GetSource(Trim(txt(0).Text))
            End If
        Else
            '修改
            lngKey = mlngKey
        End If
        '备份一分标准过程
        
        strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & cbo(0).ItemData(cbo(0).ListIndex) & ",'" & IIf((Trim(txt(0).Text) = ""), Trim(cbo(2).Text), Trim(txt(0).Text)) & "',3,'" & txt(1).Text & "','" & Trim(cbo(1).Text) & "')"
        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        
        strFlag = gclsBase.GetTooLong(Trim(txtProgramEdit.Text))
        strFlag = Replace(strFlag, "'", "''")
        strArr = Split(strFlag, gstrSplite)
        For lngLoop = 0 To UBound(strArr)
            strSQL = "Zl_Zlproceduretext_Update(" & lngKey & ",3," & (lngLoop + 1) & ",'" & strArr(lngLoop) & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        Next
        
        strFlag = ""
        If strTemp <> "" Then
            strFlag = gclsBase.GetTooLong(Trim(strTemp))
            strFlag = Replace(strFlag, "'", "''")
            strArr = Split(strFlag, gstrSplite)
            For lngLoop = 0 To UBound(strArr)
                strSQL = "Zl_Zlproceduretext_Update(" & lngKey & ",4," & (lngLoop + 1) & ",'" & strArr(lngLoop) & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            Next
        End If
        
        If SQLRecordExecute(rsSQL, "") Then
            mblnOK = True
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "暂存数据"
        '验证数据有效性
        
        strTemp = ""
        '保存数据
        If mlngKey = 0 Then
        '新增
            lngKey = gclsBase.GetNextId("zlProcedure")
        Else
            '修改
            lngKey = mlngKey
        End If
        
        strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & cbo(0).ItemData(cbo(0).ListIndex) & ",'" & IIf((Trim(txt(0).Text) = ""), Trim(cbo(2).Text), Trim(txt(0).Text)) & "',2,'" & txt(1).Text & "','" & Trim(cbo(1).Text) & "')"
        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        
        strFlag = gclsBase.GetTooLong(Trim(txtProgramEdit.Text))
        strFlag = Replace(strFlag, "'", "''")
        strArr = Split(strFlag, gstrSplite)
        For lngLoop = 0 To UBound(strArr)
            strSQL = "Zl_Zlproceduretext_Update(" & lngKey & ",3," & (lngLoop + 1) & ",'" & strArr(lngLoop) & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        Next
        
        If SQLRecordExecute(rsSQL, Me.Caption) Then
            mblnOK = True
            mblnDataChanged = False
        End If
    End Select
    
    ExecuteCommand = True
    
    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
EndHand:
    mblnReading = False
    
'    Resume
End Function

Private Function CheckFile(ByVal strPath1 As String, ByVal strPath2 As String, ByVal strReports As String) As Boolean
    Dim strCommand As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    
    strCommand = GetWinSystemPath & "\wincmp3.exe " & strPath1 & "\" & " " & strPath2 & "\" & " /G:HE " & strReports
    lngTemp = Shell(strCommand, vbHide)
    DoEvents
    If err <> 0 Then
        err.Clear
         MsgBox "文件比较失败，请检查" & GetWinSystemPath & "\wincmp3.exe文件是否存在", vbExclamation, "中联软件"
        Exit Function
    End If
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    CheckFile = True
    err.Clear
    DoEvents
End Function

Private Function Contains(ByVal Coll As Collection, ByVal strKey As String) As Boolean
    On Error GoTo errHand
    
    Dim Item As Variant
'    Set item = New ADODB.Connection
    Set Item = Coll.Item(strKey)
    Contains = True
    Set Item = Nothing
    Exit Function
errHand:
    '不存在返回False
    If err.Number = 5 Then Contains = False
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As ADODB.Recordset
    
    If gclsBase.StrIsValid(txt(0).Text, txt(0).MaxLength) = False Then
        gclsBase.LocationObj txt(0)
        Exit Function
    End If
        
    If gclsBase.StrIsValid(txt(1).Text, txt(1).MaxLength) = False Then
        gclsBase.LocationObj txt(1)
        Exit Function
    End If
    If (cbo(0).ItemData(cbo(0).ListIndex) = 1 Or cbo(0).ItemData(cbo(0).ListIndex) = 2) And mlngKey = 0 Then
        If Trim(cbo(2).Text) = "" Then
            MsgBox "过程名称不能为空值，必须输入！", vbInformation + vbOKOnly, "中联软件"
            gclsBase.LocationObj txt(1)
            Exit Function
        End If
    End If
    If cbo(1).ItemData(cbo(1).ListIndex) = 0 Then
        MsgBox "请指定过程所有者！", vbInformation + vbOKOnly, "中联软件"
        Exit Function
    End If
    If cbo(0).ItemData(cbo(0).ListIndex) = 3 Then
        If Trim(txt(0).Text) = "" Then
            MsgBox "过程名称不能为空值，必须输入！", vbInformation + vbOKOnly, "中联软件"
            gclsBase.LocationObj txt(1)
            Exit Function
        End If
        '验证是否选择了所有者
        If Trim(cbo(1).Text) = "" Then
            MsgBox "请选择当前用户过程的所有者！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
        '验证过程名称是否匹配
        If Trim(txt(1).Text) = "" Then
            MsgBox "用户过程的过程说明不能为空！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
    End If
    
    ValidData = True
    
End Function

Private Sub cbo_Change(Index As Integer)
    If mblnReading Then Exit Sub
    Select Case Index
    Case 2
        
        Call SendMessage(cbo(2).hwnd, CB_SHOWDROPDOWN, 1, 0)

    End Select
End Sub

Private Sub cbo_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim strFlag As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If mblnReading Then Exit Sub
    mblnDataChanged = True
    Select Case Index
    Case 0
        Select Case cbo(0).ItemData(cbo(0).ListIndex)
        Case 1, 2
            cbo(2).Visible = True
            txt(0).Visible = False
            cmd(4).Visible = False
            lbl(2).Visible = False
            cbo(1).Visible = False
            txt(0).Text = ""
        Case 3
            cbo(2).Visible = False
            txt(0).Visible = True
            cmd(4).Visible = True
            lbl(2).Visible = True
            cbo(1).Visible = True
            txt(0).Text = "ZLUSER_"
            txt(0).SetFocus
        End Select
        cbo(2).Clear
        txtProgramEdit.Text = ""
        If mlngKey = 0 Then
            '加载标准过程
            If cbo(0).ItemData(cbo(0).ListIndex) <> 3 Then
                Call ExecuteCommand("加载过程列表")
                Call cbo_Click(2)
            End If
        End If
    Case 2
        If Trim(cbo(2).Text) <> "" Then
            txtProgramEdit.Text = gclsBase.GetSource(Trim(cbo(2).Text))
            strSQL = "Select owner From all_source Where type in ('PROCEDURE','FUNCTION') and Name='" & Trim(cbo(2).Text) & "'"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            If rs.BOF = False Then
                cbo(1).Text = rs("owner").value
            End If
        End If
    End Select
    
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case conMenu_Edit_SaveExit
        If ExecuteCommand("校验数据") Then
            If ExecuteCommand("保存数据") Then
                Unload Me
            End If
        End If
    Case conMenu_Edit_Save
        Call ExecuteCommand("暂存数据")
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case conMenu_Edit_Save
        Control.Enabled = mblnDataChanged
    Case conMenu_Edit_SaveExit
    
        Control.Enabled = mblnDataChanged Or mintState = 2
        
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case 1
        Item.Handle = picPane(0).hwnd
    Case 2
        Item.Handle = picPane(3).hwnd
    Case 3
        Item.Handle = picPane(4).hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents
    If ExecuteCommand("初始数据") = False Then GoTo errHand
    
    Call ExecuteCommand("刷新数据")
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    mblnStartUp = True
    mblnDataChanged = False
    Call ExecuteCommand("初始控件")
    
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call gclsBase.SetPaneRange(dkpMain, 1, picPane(1).ScaleWidth, 60, picPane(1).ScaleWidth, 60)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Not (mfrmProgramEditMark Is Nothing) Then
'        Unload mfrmProgramEditMark
'    End If
    If Not (mrsProcedure Is Nothing) Then
        Set mrsProcedure = Nothing
    End If
    If gobjFile.FolderExists(mstrReportsFile) Then
        Call gobjFile.DeleteFolder(mstrReportsFile)
    End If
    If Not (mcnOracle Is Nothing) Then
        Set mcnOracle = Nothing
    End If
    If Not (mfrmProcedureOwnerCon Is Nothing) Then
        Unload mfrmProcedureOwnerCon
    End If
    
    Set mrsProcedure = Nothing
End Sub

Private Sub mfrmProgramEditMark_AfterAdd(ByVal strSign As String)
    Dim LineIndex As Long
    Dim SelRange As CHARRANGE
    Dim TempStr As String
    Dim TempArray() As Byte
    Dim curRow As Long
    Dim lngStart As Long
    Dim lngFlag As Long
    Dim strLeft As String
    Dim strRight As String
    Dim strField As String
    Dim lngPos As Long
    Dim nBar As Long
    Dim lngRet As Long
    
    mlngSelStart = txtProgramEdit.SelStart
    TempArray = StrConv(txtProgramEdit.Text, vbFromUnicode)
    
    '取得当前被选中文本的位置 适用于 RichTextBox
    'TextControl 用 EM_GETSEL 消息
    Call SendMessage(txtProgramEdit.hwnd, EM_EXGETSEL, 0, SelRange)
    
    '根据参数wParam指定的字符位置返回该字符所在的行号
    curRow = SendMessage(txtProgramEdit.hwnd, EM_LINEFROMCHAR, SelRange.cpMin, 0)
    
    '取得指定行第一个字符的位置
    LineIndex = SendMessage(txtProgramEdit.hwnd, EM_LINEINDEX, curRow, 0)
    
    If SelRange.cpMin = LineIndex Then
    
    Else
        TempStr = String(SelRange.cpMin - LineIndex, 13)
        '复制当前行开始到选择文本开始的文本
        CopyMemory ByVal StrPtr(TempStr), ByVal StrPtr(TempArray) + LineIndex, SelRange.cpMin - LineIndex
        TempArray = TempStr
        '删除无用的信息
        ReDim Preserve TempArray(SelRange.cpMin - LineIndex - 1)
        '转换为 Unicode
        TempStr = StrConv(TempArray, vbUnicode)
    End If
    strLeft = Mid(txtProgramEdit.Text, 1, txtProgramEdit.SelStart - Len(TempStr))
    strRight = TempStr & Mid(txtProgramEdit.Text, txtProgramEdit.SelStart + 1)
    strField = "  --" & strSign & vbCrLf
    
'    lngPos = GetScrollPos(txtProgramEdit.hwnd, SB_CTL) '得到当前滚动条位置
    
    txtProgramEdit.Text = strLeft & strField & strRight
    txtProgramEdit.SelStart = mlngSelStart + Len(strField)

'    lngRet = SetScrollPos(txtProgramEdit.hwnd, SB_CTL, lngPos, 1) '将滚动条设置为原来的位置
'    lngRet = SendMessage(txtProgramEdit.hwnd, WM_VSCROLL, lngPos * 65536 + 5, ByVal &O0) '显示滚动条位置对应的内容
End Sub

Private Sub mfrmProgramEditMark_AfterChanged()
    mblnDataChanged = True
End Sub

Private Sub mfrmProgramEditMark_AfterDelete(ByVal strSign As String)
    '寻找文本中的标识并将其删除
    txtProgramEdit.Text = Replace(txtProgramEdit.Text, "  --" & strSign & vbCrLf, "")
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        cbo(1).Move cbo(1).Left, cbo(1).Top, picPane(Index).ScaleWidth - cbo(1).Left - 75
        txt(1).Move txt(1).Left, txt(1).Top, picPane(Index).ScaleWidth - txt(1).Left - 75
        fra.Move 15, -75, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight + 75
    Case 1
        txtProgramEdit.Move 15, 15, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight - 30
    Case 2
        wbr.Move 15, 0, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight
        txtLastProgram.Move 15, 15, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight - 30
    Case 3
        tbcBase(0).Move 0, 0, picPane(Index).ScaleWidth, picPane(Index).ScaleHeight
    Case 4
        tbcBase(1).Move 0, 0, picPane(Index).ScaleWidth, picPane(Index).ScaleHeight
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    If mblnReading Then Exit Sub
    mblnDataChanged = True
End Sub

Private Sub txtProgramEdit_Change()
    If mblnReading Then Exit Sub
    mblnDataChanged = True
End Sub

Private Sub txtProgramEdit_KeyPress(KeyAscii As Integer)
    mblnDataChanged = True
End Sub


