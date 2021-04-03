VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "中华人民共和国传染病报告卡"
   ClientHeight    =   9855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   15105
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar prgSaveData 
      Height          =   330
      Left            =   1365
      TabIndex        =   1
      Top             =   9540
      Visible         =   0   'False
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Max             =   44
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9480
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23733
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   930
      Top             =   375
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":115C
      Left            =   2730
      Top             =   570
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmReport As frmReport     '编辑窗体
Attribute mfrmReport.VB_VarHelpID = -1
Private blnFirstActive As Boolean

Public Sub ShowMe(ByVal frmParent As Object, ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long)
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    Set mfrmReport = New frmReport
    mfrmReport.blnHaveStatus = True
    
    Call mfrmReport.InitReport(bytType, lngPatiID, lngPageID, bytFrom, bytBabyNo, lngDeptID, lngFileId)
    
    If lngPatiID <> 0 Then
        If bytType = 1 Then
            strSql = "select t.最后版本 from 电子病历记录 t where t.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", lngFileId)
            If rsTemp.RecordCount <> 0 Then
                If Nvl(rsTemp!最后版本) = 1 Then
                    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Enabled = False
                    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Cancel).Visible = True
                    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Enabled = False
                    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Visible = False
                Else
                    Call mfrmReport.CanWrite
                End If
            End If
        Else
            Call mfrmReport.CanWrite
        End If
        Call mfrmReport.LoadData(bytType)
    End If
    
    Me.Show 1, frmParent

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Public Sub InitSc()
'初始界面布局
    Dim Pane1 As Pane
    
    On Error GoTo errHand
    
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    dkpMain.DestroyAll
    
    Set Pane1 = dkpMain.CreatePane(1, 250, 250, DockLeftOf, Nothing)
    Pane1.Options = PaneNoCloseable + PaneNoCaption + PaneNoHideable + PaneNoFloatable
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Err GoTo errHand
    
    Select Case Control.ID
        Case conMenu_Manage_Exit
            Call Menu_Exit
            
        Case conMenu_Manage_Finish
            Call Menu_Finish
            
        Case conMenu_Manage_Cancel
            Call Menu_Cancel
            
        Case conMenu_Manage_Save
            Call Menu_Save
    End Select

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Private Sub Menu_Save()
'暂存
    On Error GoTo errHand
    
    prgSaveData.Visible = True
    prgSaveData.Value = 0
    Call mfrmReport.ClearEnterInfo
    Call mfrmReport.SaveData(False)
    prgSaveData.Visible = False
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Private Sub Menu_Exit()
'退出
    Dim result As VbMsgBoxResult
    
    On Error GoTo errHand
    If mfrmReport.HaveChanged = True Then
        result = MsgBox("是否保存修改内容？", vbYesNoCancel + vbQuestion, gstrSysName)
        If result = vbYes Then
            Call Menu_Save
            Unload Me
        ElseIf result = vbNo Then
            Unload Me
        Else
            Exit Sub
        End If
    Else
        Unload Me
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0

End Sub
Private Sub Menu_Finish()
'完成
    On Error GoTo errHand
    If CheckValidity = True Then
        Call mfrmReport.SetEnterInfo
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Enabled = False
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Cancel).Visible = True
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Enabled = False
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Visible = False
        prgSaveData.Visible = True
        prgSaveData.Value = 0
        Call mfrmReport.SaveData(True)
        prgSaveData.Visible = False
        Unload Me
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Private Sub Menu_Cancel()
'取消保存
    Dim strSql As String
    
    On Error GoTo errHand
    Call mfrmReport.CanWrite
    Call mfrmReport.ClearEnterInfo
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Enabled = True
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Cancel).Visible = False
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Enabled = True
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Visible = True

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = mfrmReport.hWnd
    End If
End Sub

Private Sub Form_Activate()
    If blnFirstActive = True Then
        Call mfrmReport.SetMyFocus
        blnFirstActive = False
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    blnFirstActive = True
    Me.WindowState = 2
    Call InitCommandBars

    Call InitSc

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Function CheckValidity() As Boolean
'检查编辑界面的合法性
    CheckValidity = mfrmReport.CheckValidity
End Function

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Save, "暂存(&S)", "暂时保存", 3503, True)
    cbrControl.Style = xtpButtonIconAndCaption
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Finish, "完成(&F)", "完成编辑", 804, False)
    cbrControl.Style = xtpButtonIconAndCaption
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Cancel, "取消完成(&C)", "取消完成", 3504, False)
    cbrControl.Style = xtpButtonIconAndCaption
    cbrControl.Visible = False
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Exit, "退出(&E)", "退出编辑", 191, True)
    cbrControl.Style = xtpButtonIconAndCaption
    
    With cbrMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_Manage_Save
        .Add FCONTROL, vbKeyF, conMenu_Manage_Finish
        .Add FCONTROL, vbKeyU, conMenu_Manage_Cancel
        .Add FCONTROL, vbKeyE, conMenu_Manage_Exit
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'创建该模块内的菜单
    
    On Error GoTo errHand
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If

    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Private Sub Form_Resize()
    On Error Resume Next
    prgSaveData.Width = Me.ScaleWidth - 1400
    prgSaveData.Left = 1400
    prgSaveData.Top = Me.ScaleTop + Me.ScaleHeight - prgSaveData.Height
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmReport
End Sub


Private Sub mfrmReport_HaveSavedSQL()
    On Error Resume Next
    prgSaveData.Value = prgSaveData.Value + 1
    Err.Clear
End Sub
