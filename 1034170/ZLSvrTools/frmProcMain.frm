VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmProcMain 
   Caption         =   "�Զ�����̹���"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15990
   Icon            =   "frmProcMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   15990
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3555
      Index           =   0
      Left            =   90
      ScaleHeight     =   3555
      ScaleWidth      =   12405
      TabIndex        =   0
      Top             =   1470
      Width           =   12405
      Begin VB.PictureBox picPane 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEBD7&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   15
         ScaleHeight     =   375
         ScaleWidth      =   10980
         TabIndex        =   1
         Top             =   15
         Width           =   10980
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   9150
            TabIndex        =   5
            ToolTipText     =   "��ֱ�Ӱ��س������й���"
            Top             =   45
            Width           =   1695
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "�û�����"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   2835
            TabIndex        =   4
            Top             =   90
            Width           =   1305
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "�հ׹���"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1455
            TabIndex        =   3
            Top             =   90
            Width           =   1305
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "�䶯����"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Top             =   90
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ��"
            Height          =   180
            Left            =   8625
            TabIndex        =   6
            Top             =   90
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1755
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   930
         Width           =   1935
         _cx             =   3413
         _cy             =   3096
         Appearance      =   1
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         GridColor       =   -2147483632
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
         RowHeightMin    =   330
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   9975
      Width           =   15990
      _ExtentX        =   28205
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmProcMain.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25329
            MinWidth        =   8819
            Text            =   "��ǰ���д�����0����������0��"
            TextSave        =   "��ǰ���д�����0����������0��"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Bindings        =   "frmProcMain.frx":70E6
      Left            =   1080
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmProcMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmProgramEdit As frmProcEdit
Private mfrmDifferenceCheck As frmProcDiffrentCheck
Private mfrmBuildScript As frmProcBuildScript
Private mfrmProcedureRelating As frmProcRelating
Private mfrmCollectUpdate As frmProcCollectUpdate
Private mintProcType As Integer
Private mclsVsf As clsVsf
Private mclsVsfMark As clsVsf
Private mblnReading As Boolean
Private mobjMain As Object

Private mblnStartUp As Boolean

Private Sub InitCommandBar()
    '******************************************************************************************************************
    '���ܣ���ʼ�˵�������
    '��������
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '�༭
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.Id = conMenu_EditPopup
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Word, "�Ѽ��Ǽ�(&S)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "�½��Ǽ�(&N)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Disuse, "�������(&C)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "������(&J)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸Ĺ���(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "�ָ�����(&R)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Change_PaitNote, "���ɽű�(&G)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    

    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrWebSustainer)
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrWebSustainer & "��ҳ(&H)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrWebSustainer & "��̳(&F)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)
    
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
'    objBar.SetIconSize 16, 16
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Word, "�Ѽ�")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "�½�")
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "���", True)
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Untread, "�ָ�")
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Manage_Change_PaitNote, "����", True)
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�", True)
    
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
                
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��
                
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
End Sub

Public Function ShowMe(ByVal objParent As Object)
    Me.Show 1, objParent
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 600, 100, DockLeftOf, objPane)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call gclsBase.DockPannelInit(dkpMain)

End Sub

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim objNode As Node
    Dim intRow As Integer
    Dim rsColum As ADODB.Recordset
    Dim lngKey As Long
    Dim intType1 As Integer
    Dim intType2 As Integer
    Dim intType3 As Integer
    
    On Error GoTo errHand

    Call gclsBase.SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "��ʼ�ؼ�"
            
            Call InitCommandBar
            Call InitDockPannel
        '--------------------------------------------------------------------------------------------------------------
        Case "��ʼ����"
            Set mclsVsf = New clsVsf
            With mclsVsf
                Call .Initialize(Me.Controls, vsf(0), True, False)
                Call .ClearColumn
                Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[���]", False, False, False)
'                Call .AppendColumn("���", 800, flexAlignLeftCenter, flexDTString, , "���", False)
                Call .AppendColumn("����", 3000, flexAlignLeftCenter, flexDTString, , "����", True)
                Call .AppendColumn("״̬", 1000, flexAlignLeftCenter, flexDTString, , "״̬", True)
                Call .AppendColumn("˵��", 3000, flexAlignLeftCenter, flexDTString, , "˵��", True)
                
                .IndicatorMode = 2
                .IndicatorCol = .ColIndex("���")
                .ConstCol = .ColIndex("���")
            
                .AppendRows = True
            End With
            mintProcType = IIf((opt(0).value = True), 1, IIf((opt(1).value = True), 2, 3))
        '--------------------------------------------------------------------------------------------------------------
        Case "ˢ������"
            
            '���ԭ������
            Call mclsVsf.ClearGrid
            strSQL = "Select ID,Decode(����,1,'��׼����',2,'�հ׹���',3,'�û�����') As ����,���� As ����,Decode(״̬,1,'������',2,'������',3,'�ѵ���') As ״̬,˵��,�޸���Ա,�޸�ʱ��,�ϴ��޸���Ա,�ϴ��޸�ʱ�� From zlprocedure Where ����=[1]"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mintProcType)
            If rs.BOF = False Then
                Call mclsVsf.LoadGrid(rs)
            End If
            For intRow = 1 To vsf(0).Rows - 1
                If vsf(0).TextMatrix(intRow, vsf(0).ColIndex("״̬")) = "������" Then
                    vsf(0).Cell(flexcpForeColor, intRow, vsf(0).ColIndex("״̬")) = vbRed
                ElseIf vsf(0).TextMatrix(intRow, vsf(0).ColIndex("״̬")) = "������" Then
                    vsf(0).Cell(flexcpForeColor, intRow, vsf(0).ColIndex("״̬")) = vbBlue
                End If
            Next
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        '--------------------------------------------------------------------------------------------------------------
        Case "ˢ��״̬"
            intType1 = 0
            intType2 = 0
            For intRow = 1 To vsf(0).Rows - 1
                If vsf(0).TextMatrix(intRow, vsf(0).ColIndex("״̬")) = "������" Then
                    intType1 = intType1 + 1
                ElseIf vsf(0).TextMatrix(intRow, vsf(0).ColIndex("״̬")) = "������" Then
                    intType2 = intType2 + 1
                End If
            Next
            stbThis.Panels(2).Text = "��ǰ���д����� " & CStr(intType1) & " ��,������ " & CStr(intType2) & " ����"
        '--------------------------------------------------------------------------------------------------------------
        Case "�ָ�����"
            With vsf(0)
                lngKey = .RowData(.Row)
                If lngKey = 0 Then Exit Function
                '�õ����α䶯���̶�Ӧ�ı�׼����
                strSQL = "Select A.ID,A.����,A.����,B.���� From zlProcedure A,zlProceduretext B Where A.ID = B.����ID And A.ID=[1] And B.���� = 4"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", lngKey)
                If rs.BOF = False Then
                    '�õ��������ݲ�ִ�д˹���
                    strTmp = NVL(rs("����").value)
                    strTmp = "CREATE OR REPLACE " & Trim(strTmp) & vbCrLf & "/"
                End If
                If strTmp <> "" Then
                    Call gcnOracle.Execute(strTmp)
                    strSQL = "Zl_Zlprocedure_Delete(" & lngKey & ")"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                    Call SQLRecordExecute(rsSQL)
                End If
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case "ɾ������"
            
            '��ѯ�Ƿ��й����ڵ��õ�ǰ����
            
            strSQL = "Select ID,���� From zlProcedure Where ID=[1]"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", vsf(0).RowData(vsf(0).Row))
            If rs.BOF = False Then
                strTmp = NVL(rs("����").value)
                strSQL = "Select Distinct Name As �������� from (Select Distinct Name,Type,Text From User_Source Where type in ('PROCEDURE','FUNCTION') and upper(Text) Like [1] And Name <> [2])"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", "%" & UCase(strTmp) & "%", UCase(strTmp))
                If rs.BOF = False Then
                    If mfrmProcedureRelating Is Nothing Then
                        Set mfrmProcedureRelating = New frmProcRelating
                    End If
                    Call mfrmProcedureRelating.ShowDialog(Me, vsf(0).RowData(vsf(0).Row), rs)
                Else
                    If MsgBox("ȷ��ɾ�����̡�" & strTmp & "����?", vbInformation + vbOKCancel, "�������") = vbOK Then
                        strSQL = "Zl_Zlprocedure_Delete(" & vsf(0).RowData(vsf(0).Row) & ")"
                        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                        Call SQLRecordExecute(rsSQL)
                    Else
                        Exit Function
                    End If
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case "�Ƴ�����"
            If vsf(0).Rows > 2 Then
                vsf(0).RemoveItem vsf(0).Row
                mclsVsf.AppendRows = True
            Else
                Call mclsVsf(0).ClearGrid
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case "�������ݿ�"
'            Call gclsPackage.OraDataOpen("192.168.4.19", "ZLHIS", "ZLHIS", "HIS")
        End Select
    Next
    ExecuteCommand = True
    Exit Function
errHand:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long
    Dim objControl As CommandBarControl
    Dim strSQL As String
    
    
    On Error GoTo errHand
    
    Dim rs As ADODB.Recordset
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem
        If mfrmProgramEdit Is Nothing Then
            Set mfrmProgramEdit = New frmProcEdit
        End If
        Call mfrmProgramEdit.ShowMe(Me, 0)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        If mfrmProgramEdit Is Nothing Then Set mfrmProgramEdit = New frmProcEdit
        
        If vsf(0).RowData(vsf(0).Row) > 0 Then
            If mfrmProgramEdit.ShowMe(Me, vsf(0).RowData(vsf(0).Row)) Then
                Call ExecuteCommand("ˢ������")
                Call ExecuteCommand("ˢ��״̬")
            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Disuse
    
        If MsgBox("��ȷ�������������" & vbCrLf & "�˲����Ὣ��������ǰ�Ĺ��̼�¼��Ϊ�ϴι��̼�¼��", vbOKCancel + vbInformation + vbDefaultButton2, "�������") = vbOK Then
            
            gcnOracle.Execute "Zl_Zlproceduretext_Move()"
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        If mfrmDifferenceCheck Is Nothing Then
            Set mfrmDifferenceCheck = New frmProcDiffrentCheck
        End If
        If mfrmDifferenceCheck.ShowMe(Me) Then
            Call ExecuteCommand("ˢ������")
            Call ExecuteCommand("ˢ��״̬")
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Change_PaitNote
        Set rs = gclsBase.GetProcByState(1, 2)
        If rs.BOF = False Then
            MsgBox "��⵽�й��̻�δ������ɣ����Ƚ��е����������ɡ�", vbInformation + vbOKOnly, "�������"
            Exit Sub
        End If
        If mfrmBuildScript Is Nothing Then
            Set mfrmBuildScript = New frmProcBuildScript
        End If
        Call mfrmBuildScript.ShowMe(Me)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        If ExecuteCommand("ɾ������") Then
            Call ExecuteCommand("�Ƴ�����")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread
        If MsgBox("ȷ��Ҫ���˹��ָ̻�Ϊ֮ǰ���ݵı�׼������", vbOKCancel + vbInformation, "�������") = vbOK Then
            If ExecuteCommand("�ָ�����") Then
                Call ExecuteCommand("�Ƴ�����")
            End If
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Word
        If mfrmCollectUpdate Is Nothing Then
            Set mfrmCollectUpdate = New frmProcCollectUpdate
        End If
        If mfrmCollectUpdate.ShowMe(Me) Then
            Call ExecuteCommand("ˢ������")
            Call ExecuteCommand("ˢ��״̬")
        End If
    Case conMenu_File_Exit
        Unload Me
        
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '״̬��
    
        stbThis.Visible = Not stbThis.Visible
        cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '��������
    
'        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((ParamInfo.ϵͳ��) / 100))
        
    Case conMenu_Help_Web_Home 'Web�ϵ�����
    
        Call zlHomePage(Me.hwnd)
        
    Case conMenu_Help_Web_Forum '������̳
    
        Call zlWebForum(Me.hwnd)
        
    Case conMenu_Help_Web_Mail '���ͷ���
    
        Call zlMailTo(Me.hwnd)
        
    Case conMenu_Help_About '����
        
        Call ShowAbout(Me)
        
    End Select
    Exit Sub
errHand:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    With vsf(0)
    Select Case Control.Id
    Case conMenu_Edit_Delete
        Control.Visible = (opt(2).value = True)
    Case conMenu_Edit_Untread
        Control.Visible = (opt(0).value = True Or opt(1).value = True)

    Case conMenu_View_ToolBar_Button            '������
        If cbsMain.Count >= 2 Then
            Control.Checked = cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = stbThis.Visible
    End Select

    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case 1
        Item.Handle = picPane(0).hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    Call ExecuteCommand("��ʼ����")
    Call ExecuteCommand("ˢ������")
    Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    gstrSplite = "FRCHENUSERPROCEDUREFRCHEN"
    Call ExecuteCommand("��ʼ�ؼ�")
'    Call ExecuteCommand("�������ݿ�")
End Sub

Private Sub Form_Resize()
'    On Error Resume Next
'    picPane(1).Move 0, -15, Me.ScaleWidth, picPane(1).ScaleHeight + 30
'    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
    If Not (mclsVsfMark Is Nothing) Then
        Set mclsVsfMark = Nothing
    End If
    If Not (mfrmBuildScript Is Nothing) Then
        Unload mfrmBuildScript
    End If
    
    If Not (mfrmCollectUpdate Is Nothing) Then
        Unload mfrmCollectUpdate
    End If
    
    If Not (mfrmDifferenceCheck Is Nothing) Then
        Unload mfrmDifferenceCheck
    End If
    
    If Not (mfrmProcedureRelating Is Nothing) Then
        Unload mfrmProcedureRelating
    End If
    
    If Not (mfrmProgramEdit Is Nothing) Then
        Unload mfrmProgramEdit
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    mintProcType = (Index + 1)
    Call ExecuteCommand("ˢ������")
    Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        
        picPane(2).Move 15, 15, picPane(Index).ScaleWidth - 30
        vsf(0).Move 15, picPane(2).Top + picPane(2).Height + 15, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight - (picPane(2).Top + picPane(2).Height + 15) - 15
        mclsVsf.AppendRows = True
    Case 2
        txtLocation.Move picPane(Index).ScaleWidth - txtLocation.Width - 75
        lbl1.Move txtLocation.Left - lbl1.Width - 30
    End Select
    
End Sub

Private Sub txtLocation_GotFocus()
    Call gclsBase.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    
    If KeyAscii = vbKeyReturn Then
        intCol = vsf(0).ColIndex("����")
        lngRow = mclsVsf.FindRow(UCase(txtLocation.Text), intCol, 2, vsf(0).Row + 1)
        If lngRow = -1 Then
            lngRow = mclsVsf.FindRow(UCase(txtLocation.Text), intCol, 2)
        End If
        If lngRow > 0 And vsf(0).Row <> lngRow Then
            vsf(0).Row = lngRow
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        End If
        
        Call gclsBase.LocationObj(txtLocation)
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Select Case Index
    Case 1
        If OldRow = NewRow Then Exit Sub
    End Select
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call gclsBase.SendLMouseButton(vsf(Index).hwnd, X, Y)
        Select Case Index
        Case 0
            Set cbrPopupBar = gclsBase.CopyMenu(cbsMain, 2)
            If cbrPopupBar Is Nothing Then Exit Sub
            cbrPopupBar.ShowPopup
        End Select
        
    End Select
End Sub

