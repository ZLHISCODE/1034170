VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSendLog 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '????ȱʡ
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7425
      TabIndex        =   10
      Top             =   585
      Width           =   1575
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   0
      Width           =   1185
      Begin VB.ComboBox cboPeiord 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   -30
         Width           =   1215
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   1380
      ScaleHeight     =   240
      ScaleWidth      =   1245
      TabIndex        =   6
      Top             =   0
      Width           =   1275
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   -30
         TabIndex        =   7
         Top             =   -30
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   114753539
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   3390
      ScaleHeight     =   255
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   75
      Width           =   1305
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   -30
         TabIndex        =   5
         Top             =   -30
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   114753539
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   0
      Left            =   3495
      ScaleHeight     =   1980
      ScaleWidth      =   2700
      TabIndex        =   2
      Top             =   3765
      Width           =   2700
      Begin VB.TextBox txtXML 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1305
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   300
         Width           =   1410
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   2
      Left            =   3450
      ScaleHeight     =   1980
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   975
      Width           =   2700
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   210
      Top             =   750
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmLog.frx":0000
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSendLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'????????

Private Enum Command
    ??ʼ?ؼ?
    ??ע????
    ????վ??
    ????վ??
    ˢ??վ??
    ˢ??ָ??վ??
    ?Ƴ?ָ??վ??
End Enum

Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

'######################################################################################################################
'?ӿڷ???
Public Function ShowForm()
    '******************************************************************************************************************
    '???ܣ?
    '??????
    '???أ?
    '******************************************************************************************************************
    Call Form_Activate
End Function

'######################################################################################################################
Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '???ܣ?
    '??????
    '???أ?
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim rsPara As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant
    Dim blnMuliSelect As Boolean
    
    On Error GoTo errHand
    
    
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.??ʼ?ؼ?
        
        Call InitGrid
        Call InitCommandBar
        Call InitDockPannel
    '------------------------------------------------------------------------------------------------------------------
    Case Command.????վ??
        
        Set rsPara = zlCommFun.CreateParameter
        
        With vsf(0)
            blnMuliSelect = False
            For intRow = 1 To .Rows - 1
                If Val(Abs(.TextMatrix(intRow, .ColIndex("ѡ??")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            
            If blnMuliSelect = True Then
                If MsgBox("??ȷ??Ҫ?????Ѿ???ѡ??վ??????Ϣ????", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ????) = vbYes Then
                    For intRow = 1 To .Rows - 1
                        If Val(Abs(.TextMatrix(intRow, .ColIndex("ѡ??")))) = 1 And .TextMatrix(intRow, .ColIndex("????վ")) <> "" Then
                            Call zlCommFun.SetParameter(rsPara, "????վ", .TextMatrix(intRow, .ColIndex("????վ")))
                            Call zlCommFun.SetParameter(rsPara, "??????Ϣ", "0")
                            Call gclsBusiness.ClientsEdit("UPDATE", rsPara)
                        End If
                    Next
                End If
            ElseIf .TextMatrix(.Row, .ColIndex("????վ")) <> "" Then
                Call zlCommFun.SetParameter(rsPara, "????վ", .TextMatrix(.Row, .ColIndex("????վ")))
                Call zlCommFun.SetParameter(rsPara, "??????Ϣ", "0")
                Call gclsBusiness.ClientsEdit("UPDATE", rsPara)
            End If
        End With
        Call ExecuteCommand(Command.ˢ??վ??)
    '------------------------------------------------------------------------------------------------------------------
    Case Command.????վ??
        
        Set rsPara = zlCommFun.CreateParameter
        
        With vsf(0)
            blnMuliSelect = False
            For intRow = 1 To .Rows - 1
                If Val(Abs(.TextMatrix(intRow, .ColIndex("ѡ??")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            
            If blnMuliSelect = True Then
                If MsgBox("??ȷ??Ҫ?????Ѿ???ѡ??վ??????Ϣ????", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ????) = vbYes Then
                    For intRow = 1 To .Rows - 1
                        If Val(Abs(.TextMatrix(intRow, .ColIndex("ѡ??")))) = 1 And .TextMatrix(intRow, .ColIndex("????վ")) <> "" Then
                            Call zlCommFun.SetParameter(rsPara, "????վ", .TextMatrix(intRow, .ColIndex("????վ")))
                            Call zlCommFun.SetParameter(rsPara, "??????Ϣ", "1")
                            Call gclsBusiness.ClientsEdit("UPDATE", rsPara)
                        End If
                    Next
                End If
            ElseIf .TextMatrix(.Row, .ColIndex("????վ")) <> "" Then
                Call zlCommFun.SetParameter(rsPara, "????վ", .TextMatrix(.Row, .ColIndex("????վ")))
                Call zlCommFun.SetParameter(rsPara, "??????Ϣ", "1")
                Call gclsBusiness.ClientsEdit("UPDATE", rsPara)
            End If
        End With
        Call ExecuteCommand(Command.ˢ??վ??)
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ??վ??
        

        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
            
            If Trim(txtLocation.Text) = "" Then
                ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ClientRead())
            Else
                Set rsCondition = zlCommFun.CreateCondition
                Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
                Call zlCommFun.SetCondition(rsCondition, "FilterText", Trim(txtLocation.Text))
                ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ClientRead("FilterData", rsCondition))
            End If
      
            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
        End With

            
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ??ָ??վ??
        
        ExecuteCommand = LoadCustomData(Trim(varParam(0)))
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.?Ƴ?ָ??վ??
        
        With vsf(0)
            
            intRow = mclsVsf(0).FindRow(Trim(varParam(0)), .ColIndex("id"))
            
            If intRow > 0 Then
                If .Rows > 2 Then
                    .RemoveItem .Row
                    mclsVsf(0).AppendRows = True
                Else
                    Call mclsVsf(0).ClearGrid
                End If
            End If
        End With
    
    End Select
    
    
    GoTo EndHand

    '????????
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
EndHand:
End Function

Private Function LoadCustomData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '???ܣ?
    '??????
    '???أ?
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim rsData As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    Set rsData = gclsBusiness.ItemRead("id", rsCondition)
    If rsData.BOF = True Then Exit Function
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(strDataKey, .ColIndex("id"))
        
        If intRow > 0 Then
            '?Ѽ???
            .Row = intRow
        Else
            'δ????
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("id"))) <> "" Then .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
        
        Call mclsVsf(0).LoadGridRow(.Row, rsData)
    End With
    
    mclsVsf(0).AppendRows = True
    
    LoadCustomData = True
    
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '???ܣ?
    '??????
    '???أ?
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[????]", False)
        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTBoolean, "", "[ѡ??]", False)
'        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTString, "", "[ͼ??]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("????վ", 2100, flexAlignLeftCenter, flexDTString, , "????վ", True)
        Call .AppendColumn("IP", 1500, flexAlignLeftCenter, flexDTString, , "IP", True)
        Call .AppendColumn("????ϵͳ", 3000, flexAlignLeftCenter, flexDTString, , "????ϵͳ", True)
        Call .AppendColumn("????", 1500, flexAlignLeftCenter, flexDTString, , "????", True)
        Call .AppendColumn("״̬", 600, flexAlignLeftCenter, flexDTString, , "״̬", True)
        Call .AppendColumn("˵??", 1500, flexAlignLeftCenter, flexDTString, , "˵??", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("????")
        .ConstCol = .ColIndex("????")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("ѡ??"), True, vbVsfEditCheck)
        
    End With
            
    InitGrid = True
    
    Exit Function

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '???ܣ?
    '??????
    '???أ?
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '??ʼ????
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '?˵?????:???????????ݣ?????xtpControlPopup???͵?????ID???¸?ֵ

    cbsMain.ActiveMenuBar.Title = "?˵?"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '??????????:????????????

    Set objBar = cbsMain.Add("??????", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Parameter, "????", True)
    
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "ȫѡ", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "ȫ??")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Check, "????", True)
    objControl.IconId = conMenu_File_Parameter
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Check, "????")
    objControl.IconId = conMenu_File_Parameter
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "ʱ?䣺", , , xtpButtonCaption)
    
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(2).hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "??", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(3).hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "??", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(4).hWnd
    
        
    mstrFindKey = zlDataBase.GetPara("??λ????", ParamInfo.ϵͳ??, mlngModualCode, "????վ")
    If mstrFindKey = "" Then mstrFindKey = "????վ"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, False, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.????վ"): objControl.Parameter = "????վ"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.IP"): objControl.Parameter = "IP"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.????"): objControl.Parameter = "????"
    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "????")
    
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Close, "?ر?")
    objControl.Flags = xtpFlagRightAlign
    
    '------------------------------------------------------------------------------------------------------------------
    '?????Ŀ???????:???????????????Ѵ???

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh           'ˢ??
        .Add 0, vbKeyDelete, conMenu_Edit_Delete
        .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem     '????
    End With
        
    Exit Function
    
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '????:
    '????:
    '????:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 300, 100, DockLeftOf, Nothing)
    objPane.Title = "??¼"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, objPane)
    objPane.Title = "????"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter
        Call frmStationParameter.ShowConfigDialog(Me)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("ѡ??"), .Rows - 1, .ColIndex("ѡ??")) = 1
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("ѡ??"), .Rows - 1, .ColIndex("ѡ??")) = 0
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Reuse
    
        Call ExecuteCommand(Command.????վ??)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Pause
        
        Call ExecuteCommand(Command.????վ??)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter               'ˢ??
                
        Call ExecuteCommand(Command.ˢ??վ??)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    Case conMenu_File_Close
    '--------------------------------------------------------------------------------------------------------------
        Unload Me
        RaiseEvent AfterClose(mlngModualCode)
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    Dim blnMuliSelect As Boolean
    
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Reuse
            
            For intRow = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ??")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            If blnMuliSelect = False Then
                Control.Enabled = (.TextMatrix(.Row, .ColIndex("״̬")) = "????")
            Else
                Control.Enabled = True
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Pause
            
            For intRow = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ??")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            If blnMuliSelect = False Then
                Control.Enabled = (.TextMatrix(.Row, .ColIndex("״̬")) = "????")
            Else
                Control.Enabled = True
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_DeleteParent
        
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem               '????

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify               '?޸?
            
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
                    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
                    
        End Select
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(2).hWnd
    Case 2
        Item.Handle = picPane(0).hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
    
    Call ExecuteCommand(Command.ˢ??վ??)
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1003
    
    Call ExecuteCommand(Command.??ʼ?ؼ?)
    Call ExecuteCommand(Command.??ע????)

    Call zlComLib.RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDataBase.ShowReportMenu(Me, ParamInfo.ϵͳ??, ParamInfo.ģ????, UserInfo.ģ??Ȩ??)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
            
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        txtXML.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 2
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

'        If txtLocation.Text <> "" Then
            txtLocation.Tag = ""
            
            Dim obj As CommandBarControl
            
            Set obj = cbsMain.FindControl(, conMenu_View_Filter, True)
            If obj.Enabled = True Then
                Call cbsMain_Execute(obj)
            End If

'        End If
'        txtLocation.Tag = ""
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf(Index).AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey, .ColIndex("id"))
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        mclsVsf(Index).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
    End With
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Call mclsVsf(Index).BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Dim objMenu As CommandBarControl
    
    Set objMenu = cbsMain.FindControl(, conMenu_Edit_Modify, False)
    If Not (objMenu Is Nothing) Then
        If objMenu.Enabled = True Then
            Call cbsMain_Execute(objMenu)
        End If
    End If
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf(Index).KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mclsVsf(Index).MoveColumn = (vsf(Index).MouseRow = 0)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '?????˵?????
        Call zlCommFun.SendLMouseButton(vsf(Index).hWnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                Call ShowConetneMenu(1).ShowPopup
            End If
        End Select
        
    End Select
End Sub

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '???ܣ?
    '??????
    '???أ?
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '?????˵?????
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("?????˵?", xtpBarPopup)
    
    Select Case bytPlace
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '

        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SelAll, "ȫ????ѡ(&A)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ????ѡ(&U)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Reuse, "????վ??(&N)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Pause, "????վ??(&M)")
        
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(0).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).ValidateEdit(Col, Cancel)
End Sub

