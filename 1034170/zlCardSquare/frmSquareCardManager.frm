VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareCardManager 
   Caption         =   "���ѿ�����"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   Icon            =   "frmSquareCardManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12105
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCardList 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   -360
      ScaleHeight     =   2565
      ScaleWidth      =   12405
      TabIndex        =   3
      Top             =   645
      Width           =   12405
      Begin VB.PictureBox picModify 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   4845
         ScaleHeight     =   465
         ScaleWidth      =   5850
         TabIndex        =   11
         Top             =   -90
         Visible         =   0   'False
         Width           =   5850
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   300
            Left            =   4080
            TabIndex        =   17
            Top             =   135
            Width           =   315
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "����޸�(&O)"
            Height          =   350
            Left            =   4770
            TabIndex        =   14
            Top             =   105
            Width           =   1230
         End
         Begin VB.CheckBox chk�޸� 
            Caption         =   "�޸Ŀ�����(&X)"
            Height          =   350
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Value           =   2  'Grayed
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Height          =   330
            Left            =   2115
            TabIndex        =   12
            Top             =   120
            Width           =   2280
         End
         Begin MSComCtl2.DTPicker dtp����Ч���� 
            Height          =   300
            Left            =   2115
            TabIndex        =   16
            Top             =   120
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   113115139
            CurrentDate     =   40156.0854282407
         End
         Begin VB.ComboBox cbo������ 
            Height          =   300
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   130
            Width           =   2310
         End
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "����"
         Height          =   405
         Index           =   3
         Left            =   4050
         TabIndex        =   10
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "�˿�"
         Height          =   405
         Index           =   2
         Left            =   3210
         TabIndex        =   9
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "ʧЧ��"
         Height          =   405
         Index           =   1
         Left            =   2235
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "��Ч��"
         Height          =   405
         Index           =   0
         Left            =   1305
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   2055
         Left            =   105
         TabIndex        =   4
         Top             =   435
         Width           =   6825
         _cx             =   12039
         _cy             =   3625
         Appearance      =   1
         BorderStyle     =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSquareCardManager.frx":6852
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmSquareCardManager.frx":6B7E
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����Ϣ"
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   90
         Width           =   900
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   4590
      ScaleHeight     =   1875
      ScaleWidth      =   5070
      TabIndex        =   0
      Top             =   5175
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8025
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareCardManager.frx":70CC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12356
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   645
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareCardManager.frx":7960
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareCardManager.frx":7CB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   495
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSquareCardManager.frx":8008
      Left            =   1005
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSquareCardManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mblnFirst  As Boolean, mstrPrivs As String, mstrTitle As String    '���ܱ���
Private mlngModule As Long, mstrKey As String
Private Type Ty_CurrCardStartus '��ǰ��״̬
    blnHaveData As Boolean
    bln���� As Boolean
    bln�˿� As Boolean
    bln���� As Boolean
    bln�û�ʹ����Ч�� As Boolean
    bln�������� As Boolean
    bln�����ֵ���� As Boolean
    blnͣ�ÿ� As Boolean
End Type
Private mTy_CardStartus As Ty_CurrCardStartus
Private Enum mPgIndex
    Pg_��ֵ��¼ = 250101
    Pg_���ռ�¼ = 250102
    Pg_���Ѽ�¼ = 250103
End Enum
Private Enum mPaneID
    Pane_Search = 1     '��������
    Pane_CardLists = 2  '���б�
    Pane_CardDetails = 3    '��ϸ�б�
End Enum
Private mlng�ӿڱ�� As Long
Private mrs���ѿ��ӿ� As ADODB.Recordset
Private mfrmSquareCardCallBack As frmSquareCardCallBack
Private WithEvents mfrmSquareCardConsume As frmSquareCardConsume
Attribute mfrmSquareCardConsume.VB_VarHelpID = -1
Private WithEvents mfrmSquareCardInFull As frmSquareCardInFul
Attribute mfrmSquareCardInFull.VB_VarHelpID = -1
Private WithEvents mfrmFilter As frmSquareCardFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mPanSearch As Pane
Private mobjSubFrm As Collection
Private mfrmActive As Form
Private mArrFilter As Variant
Private Const mconMenu_Lable = 3999
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-11-20 16:02:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsCardList
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("����")) = "1|0"
        .ColData(.ColIndex("��־")) = "-1|1"
        .ColData(.ColIndex("��ǰ���")) = "1|0"
        If .ColIndex("ID") >= 0 Then
            .ColData(.ColIndex("ID")) = "-1|1"
            .ColHidden(.ColIndex("ID")) = True
        End If
    End With
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-11-19 15:15:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo Errhand:
    Set mobjSubFrm = New Collection
    Set mfrmSquareCardInFull = New frmSquareCardInFul
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_��ֵ��¼, "��ֵ��Ϣ", mfrmSquareCardInFull.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_��ֵ��¼
    mobjSubFrm.Add mfrmSquareCardInFull, CStr(objItem.Tag)
    '�г�ֵ�����Ȩ��ʱ������ʾ����
    '106681:���ϴ���2017/3/10������Ȩ�޵�ȫ����"���˳�ֵ"
    If zlCheckPrivs(mstrPrivs, "��ֵ") Or zlCheckPrivs(mstrPrivs, "���˳�ֵ") Then
        objItem.Visible = True: i = 0
    Else
        objItem.Visible = False: i = 1:
    End If
    
    Set mfrmSquareCardCallBack = New frmSquareCardCallBack
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_���ռ�¼, "������Ϣ", mfrmSquareCardCallBack.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_���ռ�¼
    mobjSubFrm.Add mfrmSquareCardCallBack, CStr(objItem.Tag)


    Set mfrmSquareCardConsume = New frmSquareCardConsume
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_���Ѽ�¼, "������Ϣ", mfrmSquareCardConsume.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_���Ѽ�¼
    mobjSubFrm.Add mfrmSquareCardConsume, CStr(objItem.Tag)

     With tbPage
        tbPage.Item(i).Selected = True
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
End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-11-18 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    If mfrmFilter Is Nothing Then Set mfrmFilter = New frmSquareCardFilter
    Call mfrmFilter.Init����

    With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(mPaneID.Pane_Search, 400, 400, DockLeftOf, Nothing)
        mPanSearch.Title = "��������": mPanSearch.Options = PaneNoCloseable
        mPanSearch.MinTrackSize.Width = 220: mPanSearch.MaxTrackSize.Width = 300
        Set objPane = .CreatePane(mPaneID.Pane_CardLists, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        Set objPane = .CreatePane(mPaneID.Pane_CardDetails, 400, 400, DockBottomOf, objPane)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    
    zlRestoreDockPanceToReg Me, dkpMan, "����"

End Function
Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�������
    '����:��ǰ�ؼ�������,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-19 14:24:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String
     
    zlIsHaveData = False
    If Me.ActiveControl Is vsCardList Then
        zlIsHaveData = mTy_CardStartus.blnHaveData
    Else
        'dd
    End If
End Function

Private Sub chkStatus_Click(Index As Integer)
    Call SetCardRowColHide
End Sub

Private Sub chk�޸�_Click()
    Call SetModifyEnabled
End Sub
Private Sub SetModifyEnabled()
    Dim blnEnabled As Boolean
    
    blnEnabled = chk�޸�.value = 1
    cmdSel.Visible = False
    With vsCardList
        cbo������.Visible = False
        dtp����Ч����.Visible = False
        txtEdit.Visible = False
        chk�޸�.Visible = True
        cmdModify.Visible = blnEnabled
        Select Case .Col
        Case .ColIndex("��Ч��")
            dtp����Ч����.Visible = blnEnabled
            picModify.Visible = True
        Case .ColIndex("������")
            cbo������.Visible = blnEnabled
            picModify.Visible = True
        Case .ColIndex("�������")
            txtEdit.Visible = blnEnabled
            picModify.Visible = True
            cmdSel.Visible = blnEnabled
        Case Else
            picModify.Visible = False
        End Select
        picModify.Visible = zlCheckPrivs(mstrPrivs, "�޸Ŀ���Ϣ") And picModify.Visible
    End With
End Sub
Private Sub SetModiyCaption()
    With vsCardList
        Select Case .Col
        Case .ColIndex("��Ч��")
            chk�޸�.Caption = "�޸ġ���Ч�ڡ�"
        Case .ColIndex("������")
            chk�޸�.Caption = "�޸ġ������͡�"
        Case .ColIndex("�������")
            chk�޸�.Caption = "�޸ġ��������"
        Case Else
            chk�޸�.Visible = False
        End Select
    End With
End Sub
Private Sub SetModifyDefaultValue()
    Dim i As Long
    
    With vsCardList
        Select Case .Col
        Case .ColIndex("��Ч��")
            If .Row > 0 Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then
                     dtp����Ч����.value = Null
                Else
                    If CDate(.TextMatrix(.Row, .Col)) < dtp����Ч����.MinDate Then
                        dtp����Ч����.value = dtp����Ч����.MinDate
                    Else
                        dtp����Ч����.value = CDate(.TextMatrix(.Row, .Col))
                    End If
                End If
            End If
            cmdSel.Visible = False
        Case .ColIndex("������")
            If .Row > 0 Then
                For i = 0 To cbo������.ListCount - 1
                    If InStr(1, cbo������.List(i), Trim(.TextMatrix(.Row, .Col))) > 0 Then
                        cbo������.ListIndex = i: Exit For
                    End If
                Next
            End If
            cmdSel.Visible = False
        Case .ColIndex("�������")
            If .Row > 0 Then
                txtEdit.Text = Trim(.TextMatrix(.Row, .Col))
                txtEdit.Tag = txtEdit.Text
            End If
            cmdSel = True
        Case Else
            chk�޸�.Visible = False
            cmdSel.Visible = False
        End Select
    End With

End Sub

Private Sub cmdModify_Click()
   Call SaveBatchUpdateCardInfor
End Sub

Private Sub cmdSel_Click()
    If cmdSel.Visible = False Then Exit Sub
    If Select�շ����ѡ����(txtEdit, "") = False Then Exit Sub
    zlCtlSetFocus txtEdit
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Search    '������������
        Item.Handle = mfrmFilter.hWnd
    Case mPaneID.Pane_CardDetails   '��ϸ����Ϣ
        Item.Handle = picList.hWnd
    Case mPaneID.Pane_CardLists '���б�
        Item.Handle = picCardList.hWnd
    End Select
End Sub
Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode������
    '����:���˺�
    '����:2009-11-19 14:15:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng���ѿ�ID As Long, str���� As String, str������ As String
    Dim str������ As String, str�������� As String
    With vsCardList
        If .Row < 0 Then Exit Sub
        lng���ѿ�ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
        str���� = Trim(.TextMatrix(.Row, .ColIndex("����")))
        str������ = Trim(.TextMatrix(.Row, .ColIndex("������")))
        str������ = Trim(.TextMatrix(.Row, .ColIndex("������")))
        str�������� = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��")))
    End With
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "���ѿ�ID=" & lng���ѿ�ID, "����=" & str����, "������=" & str������, "������=" & str������, "��������=" & str��������)
End Sub

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '���:
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-18 16:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
        
      
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "�շ�����(&M)")
        mcbrControl.IconId = 227
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "�ش�ɿ(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With


    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "����(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardModify, "�޸�(&M)"):
         'Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBatchModify, "�����޸�(&L)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "�˿�(&B)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelBack, "ȡ���˿�(&K)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "����(&H)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelCallBack, "ȡ������(&S)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "��Ƭ����(&F)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "��Ƭͣ��(&P)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "��ֵ(&C)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "��ֵ����(&T)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord, "�޸�����(&G)"): mcbrControl.BeginGroup = True
        mcbrControl.Enabled = zlCheckPrivs(mstrPrivs, "�޸�����")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�޸�����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord_Force, "ǿ���޸�����(&O)")
        mcbrControl.Enabled = zlCheckPrivs(mstrPrivs, "ǿ���޸�����")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ǿ���޸�����")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_CardPay
        .Add FCONTROL, Asc("L"), conMenu_Edit_CardBathPay
        
        .Add FCONTROL, Asc("M"), conMenu_Edit_CardModify
        .Add FCONTROL, Asc("C"), conMenu_Edit_CardInFullBack
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F11, conMenu_Edit_RollingCurtain
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardModify, "�޸�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "�˿�"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelBack, "ȡ���˿�")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelCallBack, "ȡ������")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "��Ƭ����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "��Ƭͣ��"): mcbrControl.BeginGroup = True
                
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "��ֵ"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "��ֵ����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "�շ�����(&M)")
        mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    
    Set mcbrComboxToolBar = cbsThis.Add("���ѿ��ӿ�", xtpBarTop)
    mcbrComboxToolBar.ShowTextBelowIcons = False
    mcbrComboxToolBar.ContextMenuPresent = False
    mcbrComboxToolBar.EnableDocking xtpFlagStretched
    
    With mcbrComboxToolBar.Controls
        Set mcbrControl = .Add(xtpControlLabel, mconMenu_Lable, "���ѿ��ӿ�")
        'objControl.Flags = xtpFlagRightAlign
        Set objComBar = .Add(xtpControlComboBox, conMenu_COMBOX_INTERFACE, "���ѿ��ӿ�")
        'objComBar.Flags = xtpFlagRightAlign
        objComBar.Flags = xtpFlagControlStretched
        Dim intIndex As Integer
        intIndex = 1
        With mrs���ѿ��ӿ�
            Do While Not .EOF
                objComBar.AddItem Nvl(!���) & "-" & Nvl(!����)
                objComBar.ItemData(intIndex) = Val(Nvl(!���))
                If mlng�ӿڱ�� = Val(Nvl(!���)) Then
                   objComBar.ListIndex = intIndex
                End If
                intIndex = intIndex + 1
                .MoveNext
            Loop
        End With
        If intIndex > 1 And objComBar.ListIndex <= 0 Then
            objComBar.ListIndex = 1:
        End If
        If objComBar.ListIndex > 0 Then
             mlng�ӿڱ�� = objComBar.ItemData(objComBar.ListIndex)
        End If
        
        objComBar.Width = 120:
        
    End With
 
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub PrintReBill()
    '����:�ش�Ʊ��
    Dim lngID As Long, lng������� As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    strTemp = zlCommFun.ShowMsgbox("�ɿ��ӡ", "��ѡ����Ҫ��ӡ�Ľɿ", "����(&F),��ֵ(&I),ȡ��(&C)", Me, vbDefaultButton2)
    If strTemp = "ȡ��" Or strTemp = "" Then Exit Sub
    
    If strTemp = "����" Then
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
            If lngID <= 0 Then
                ShowMsgbox "ûѡ����ص����ѿ�"
                Exit Sub
            End If
            
        End With
        gstrSQL = "Select ������� From ���ѿ�Ŀ¼ where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
        lng������� = Val(Nvl(rsTemp!�������))
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "�������=" & lng�������, "�ɿ�=" & 0, "�Ҳ�=" & 0, "��ֵID=0", "ReportFormat=1", 2)
    Else
        lngID = mfrmSquareCardInFull.zlGet��ֵID
        If lngID <= 0 Then
            ShowMsgbox "δѡ����صĳ�ֵ��¼"
            Exit Sub
        End If
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "��ֵID=" & lngID, "�ɿ�=" & 0, "�Ҳ�=" & 0, "�������=0", "ReportFormat=2", 2)
    End If
    

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    Dim ctrCombox As CommandBarComboBox
    '------------------------------------
        
    Select Case Control.id
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_PrintSingleBill       '"�ش�ɿ(&R)")
        Call PrintReBill
    Case conMenu_Edit_RollingCurtain   '�շ�����
          Call zlExecuteChargeRollingCurtain(Me)
    Case conMenu_Edit_CardPay    '����(&S)")
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_����, mlng�ӿڱ��) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
        
    Case conMenu_Edit_CardBathPay    '��������(&P)")
    Case conMenu_Edit_CardModify    '�޸�(&M)"):
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
            If lngID < 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("��ǰ״̬")) <> "��Ч" And .TextMatrix(.Row, .ColIndex("��ǰ״̬")) <> "ʧЧ" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_�޸�, mlng�ӿڱ��, lngID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardBack    '�˿�(&B)"): mcbrControl.BeginGroup = True
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
            If lngID <= 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("��ǰ״̬")) <> "��Ч" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_�˿�, mlng�ӿڱ��, lngID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelBack   'ȡ���˿�
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
            If lngID <= 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("��ǰ״̬")) <> "�˿�" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_ȡ���˿�, mlng�ӿڱ��, lngID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCallBack    '����(&H)")
        With vsCardList
'            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
'            If lngID <= 0 Then Exit Sub
            'If .TextMatrix(.Row, .ColIndex("��ǰ״̬")) <> "��Ч" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_����, mlng�ӿڱ��, 0) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelCallBack  'ȡ������
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
            If lngID <= 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("��ǰ״̬")) <> "����" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_ȡ������, mlng�ӿڱ��, lngID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardResume        '��Ƭ����
        If SaveCardResumeAndStop(False) = False Then Exit Sub
    Case conMenu_Edit_CardStop        '��Ƭͣ��
        If SaveCardResumeAndStop(True) = False Then Exit Sub
    Case conMenu_Edit_CardInFull    '��ֵ(&C)")
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
            If lngID <= 0 Then lngID = 0
            If .TextMatrix(.Row, .ColIndex("��ǰ״̬")) <> "��Ч" Then lngID = 0
            If mTy_CardStartus.bln�û�ʹ����Ч�� = False Then lngID = 0
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_��ֵ, mlng�ӿڱ��, lngID) = False Then Exit Sub
        Call mfrmSquareCardInFull.zlReLoadData(mlng�ӿڱ��, lngID)
    Case conMenu_Edit_CardInFullBack    '��ֵ����(&T)")
         
        If mfrmSquareCardInFull.zl��ֵȡ�� = False Then Exit Sub
    Case conMenu_Edit_ChangePassWord    '�޸�����
        If frmModiCardPass.zlModifyPass(Me, mlngModule, mlng�ӿڱ��, True) Then
            Exit Sub
        End If
    Case conMenu_Edit_ChangePassWord_Force  'ǿ���޸�����
        If frmModiCardPass.zlModifyPass(Me, mlngModule, mlng�ӿڱ��, False) Then
            Exit Sub
        End If
    Case conMenu_COMBOX_INTERFACE   '���ѡ��
        Set ctrCombox = Control
        mlng�ӿڱ�� = ctrCombox.ItemData(ctrCombox.ListIndex)
        Call LoadDataToRpt
    Case conMenu_View_Refresh   'ˢ��
        '����ˢ������
        Call LoadDataToRpt
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = zlIsHaveData
    Case conMenu_Edit_RollingCurtain    '�շ�����
        Control.Visible = zlCheckPrivs(mstrPrivs_RollingCurtain, "����")
        Control.Enabled = Control.Visible
    Case conMenu_File_PrintSingleBill           '"�ش�ɿ(&R)"
        Control.Visible = zlCheckPrivs(mstrPrivs, "���ѿ��շ��վ�")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardPay, conMenu_Edit_CardBathPay   '����(&S),��������(&P)
        Control.Visible = zlCheckPrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardModify   ', conMenu_Edit_CardBatchModify   '�޸�(&M),�����޸�(&L)
        Control.Visible = zlCheckPrivs(mstrPrivs, "�޸Ŀ���Ϣ")
        Control.Enabled = Control.Visible And mTy_CardStartus.blnHaveData
    Case conMenu_Edit_CardBack    '�˿�(&B)
        Control.Visible = zlCheckPrivs(mstrPrivs, "�˿�")
        Control.Enabled = Control.Visible And (Not mTy_CardStartus.bln��������) And mTy_CardStartus.bln�û�ʹ����Ч��
    Case conMenu_Edit_CardCancelBack  'ȡ���˿�
        Control.Visible = zlCheckPrivs(mstrPrivs, "�˿�")
        Control.Enabled = Control.Visible And (Not mTy_CardStartus.bln��������) And mTy_CardStartus.bln�˿�
    Case conMenu_Edit_CardCallBack   '����(&H),��������(&J)
        Control.Visible = zlCheckPrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible
    
    Case conMenu_Edit_CardCancelCallBack  'ȡ������
        Control.Visible = zlCheckPrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible And mTy_CardStartus.bln����
    Case conMenu_Edit_CardResume        '��Ƭ����
        Control.Visible = zlCheckPrivs(mstrPrivs, "��Ƭ����")
        Control.Enabled = Control.Visible And mTy_CardStartus.blnͣ�ÿ� And mTy_CardStartus.blnHaveData
    Case conMenu_Edit_CardStop        '��Ƭͣ��
        Control.Visible = zlCheckPrivs(mstrPrivs, "��Ƭͣ��")
        Control.Enabled = Control.Visible And Not mTy_CardStartus.blnͣ�ÿ� And mTy_CardStartus.blnHaveData
    Case conMenu_Edit_CardInFull    '��ֵ(&C)")
        Control.Visible = zlCheckPrivs(mstrPrivs, "��ֵ")
        Control.Enabled = Control.Visible       ' And mTy_CardStartus.bln�û�ʹ����Ч��
        
    Case conMenu_Edit_CardInFullBack    '��ֵ����(&T)")
        Control.Visible = zlCheckPrivs(mstrPrivs, "���˳�ֵ")
        Control.Enabled = Control.Visible And mTy_CardStartus.bln�����ֵ����
    Case conMenu_View_Refresh   'ˢ��
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Parameter     '��������
            frmSquareCardParaSet.ShowParaSet Me, mlngModule, mstrPrivs
        Case Else   '�����������ܵ���
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub

    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    zl_CtlSetFocus vsCardList
    Call vsCardList_GotFocus
    mblnFirst = False
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strShow As String
    Dim i As Long
    mblnFirst = True
    
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mlng�ӿڱ�� = Val(zlDatabase.GetPara("�ϴνӿں�", glngSys, mlngModule, 0, , InStr(1, mstrPrivs, ";��������;") > 0))
    strShow = Trim(zlDatabase.GetPara("����ʾ��ʽ", glngSys, mlngModule, "1011", Array(chkStatus(0), chkStatus(1), chkStatus(2), chkStatus(3)), InStr(1, mstrPrivs, ";��������;") > 0))
    If Len(strShow) < 4 Then strShow = strShow & "11111"
    For i = 0 To 3
        chkStatus(i).value = IIf(Val(Mid(strShow, i + 1, 1)) = 1, 1, 0)
    Next
    dtp����Ч����.MinDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    dtp����Ч����.value = DateAdd("m", 1, dtp����Ч����.MinDate)
    dtp����Ч����.value = Null
    chk�޸�.value = 0
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call InitPanel
    Call InitPage
    Call zlDefCommandBars '��ʼ�˵���������
    Call InitVsGrid
    Set mArrFilter = mfrmFilter.GetFilterCon
    Call LoadDataToRpt
    '����״̬����ص���ɫ
    zlSetStatusPanelCololor Me, stbThis, 3, "ͣ��", vbRed, False
    zlSetStatusPanelCololor Me, stbThis, 4, "����", vbBlue, False
    zlSetStatusPanelCololor Me, stbThis, 5, "ʧЧ", &HFF00FF, False
    zlSetStatusPanelCololor Me, stbThis, 6, "��Ч", Me.ForeColor, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim i As Long, strTemp As String
   SaveWinState Me, App.ProductName, mstrTitle
   
    strTemp = ""
    For i = 0 To 3
        strTemp = strTemp & IIf(chkStatus(i).value = 1, 1, 0)
    Next
   
   zlDatabase.SetPara "����ʾ��ʽ", strTemp, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
   zlDatabase.SetPara "�ϴνӿں�", mlng�ӿڱ��, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
   
   zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlCheckPrivs(mstrPrivs, "��������")
   
   zlSaveDockPanceToReg Me, dkpMan, "����"
   
    '�ر��Ӵ���
    For i = 1 To mobjSubFrm.count
        If Not mobjSubFrm(i) Is Nothing Then Unload mobjSubFrm(i)
    Next
    If Not frmModiCardPass Is Nothing Then Unload frmModiCardPass
End Sub
Private Sub SetCardRowColHide(Optional lngLocalRow As Long = -1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����е���ʾ������
    '���:lngLocalRow -ָ����(-1����ȫ����������)
    '����:
    '����:
    '����:���˺�
    '����:2009-12-22 21:05:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngRows As Long, i As Long
    Dim lngCurRow As Long
    
    Err = 0: On Error GoTo Errhand:
    
    With vsCardList
        i = 1: lngRows = .Rows - 1
        If lngLocalRow < 0 Then
            .Redraw = flexRDNone
        Else
            i = lngLocalRow: lngRows = lngLocalRow
        End If
        
        For lngRow = i To lngRows
            '1-��Ч, 2-����,3-�˿�,4-ʧЧ,8-ͣ��
            .RowHidden(lngRow) = False
            Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("��ǰ״̬")))
            Case 1
                If chkStatus(0).value = 0 Then .RowHidden(lngRow) = True
            Case 2
                If chkStatus(3).value = 0 Then .RowHidden(lngRow) = True
            Case 3
                If chkStatus(2).value = 0 Then .RowHidden(lngRow) = True
            Case 4
                If chkStatus(1).value = 0 Then .RowHidden(lngRow) = True
            End Select
            If .RowHidden(lngRow) = False Then
                If lngCurRow < .Row Then lngCurRow = lngRow
            End If
        Next
        If lngLocalRow < 0 Then
            If lngCurRow > 0 And .RowHidden(.Row) = True Then .Row = lngCurRow
            .Redraw = flexRDBuffered
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume

End Sub


Private Function zlPopuMenus(ByVal blnListView As Boolean) As Boolean
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Err = 0: On Error Resume Next
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Function
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
        cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
    Next

    If Me.cbsThis.ActiveMenuBar.Controls(3).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(3)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls

            Select Case mcbrControl.id
            Case conMenu_View_ShowStoped, conMenu_View_ShowAll, conMenu_View_Refresh
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
                cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
                cbrPopupItem.Checked = mcbrControl.Checked
            End Select
        Next
    End If
    cbrPopupBar.ShowPopup
End Function
Private Function zlCheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:���ݺϷ�,����true�����򷵻�False
    '����:���˺�
    '����:2009-11-19 15:37:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    zlCheckDepend = False
 
    On Error GoTo errHandle
    
    gstrSQL = "Select ����   From ���㷽ʽ Where ���� = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ֽ���㷽ʽ", UserInfo.id)
    If rsTemp.EOF Then
        ShowMsgbox "���㷽ʽ�в�����һ�������ֽ����ʵĽ��㷽ʽ,���ڽ��㷽ʽ����������!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    Set mrs���ѿ��ӿ� = zlGet���ѿ��ӿ�
    mrs���ѿ��ӿ�.Filter = "���ƿ�=1"
    If mrs���ѿ��ӿ�.RecordCount = 0 Then
        ShowMsgbox "���ѿ��ӿ��в�������Ӧ�����ѿ��ӿ�,���ܽ���ά��!"
        Exit Function
    End If
    
    Set rsTemp = zlGet�շ����
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "   û����ص��շ���Ŀ���,����ϵͳ����Ա��ϵ!"
        Exit Function
    End If
    gstrSQL = "Select rownum as ID, ����,����, ȱʡ���, ȱʡ�ۿ�, ȱʡ��־ From ���ѿ�����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ѿ�����")
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "   û��������ص����ѿ�����,����[�ֵ����]������!"
        Exit Function
    End If
    zlComboxLoadFromRecodeset Me.Caption, rsTemp, cbo������, True
    zlCheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub ShowList(ByVal lngModule As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,��ʾ��ص���Ŀ��������Ϣ
    '����:���˺�
    '����:2009-11-19 15:38:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrTitle = strTitle: mstrPrivs = gstrPrivs
    If Not zlCheckDepend Then Exit Sub            '���������Բ���
    Me.Caption = strTitle
    RestoreWinState Me, App.ProductName, mstrTitle
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hWnd, frmMain
    End If
    Me.ZOrder 0
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    
    vRect = GetControlRect(picImgList.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCardList, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlCheckPrivs(mstrPrivs, "��������")
    
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    '���¼�������
    Call LoadDataToRpt
End Sub
Private Function LoadDataToRpt() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-19 15:43:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSubWhere As String, lngRow As Long, lngPre����ID As Long
    Dim rsTemp As ADODB.Recordset, strCurDate As String
    
    Err = 0: On Error GoTo Errhand:
    strSubWhere = ""
    
    If mArrFilter("����ʱ��")(0) <> "1901-01-01" And mArrFilter("����ʱ��")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And (����ʱ�� Between [1] And [2] Or ����ʱ�� Between [3] And [4])"
    ElseIf mArrFilter("����ʱ��")(0) = "1901-01-01" And mArrFilter("����ʱ��")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And (����ʱ�� Between [3] And [4])"
    ElseIf mArrFilter("����ʱ��")(0) <> "1901-01-01" And mArrFilter("����ʱ��")(0) = "1901-01-01" Then
        strSubWhere = strSubWhere & " And (����ʱ�� Between [1] And [2])"
    End If
    If mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        strSubWhere = strSubWhere & " And (���� Between [5] And [6])"
    ElseIf mArrFilter("���ŷ�Χ")(0) = "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        strWhere = strWhere & " And A.����=[6]"
    ElseIf mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) = "" Then
        strWhere = strWhere & " And A.����=[5]"
    End If
    If strSubWhere = "" Then
        '���û�нᶨʱ�䷶Χ,��ֻ�ܲ��ҵ�ǰ���쿨�˺ͷ�����
        If mArrFilter("�쿨��") <> "" Then strWhere = strWhere & " and  A.�쿨�� like [7]"
        If mArrFilter("������") <> "" Then strWhere = strWhere & " and  A.������ like [8]"
    Else
        If mArrFilter("�쿨��") <> "" Then strSubWhere = strSubWhere & " and  �쿨�� like [7]"
        If mArrFilter("������") <> "" Then strSubWhere = strSubWhere & " and  ������ like [8]"
    End If
    
    If Trim(mArrFilter("������")) <> "����" Then strWhere = strWhere & " and  A.������ = [9]"
    
    If Val(mArrFilter("����ͣ�ÿ�")) = 1 Then
        strWhere = strWhere & " And  A.��ǰ״̬ <= 9"   '��Ҫ�õ�����
    Else
        strWhere = strWhere & " And  A.��ǰ״̬+0 <= 9 And A.ͣ������ >= To_Date('3000-01-01', 'yyyy-mm-dd')"   '��Ҫ�õ�����
    End If
    
    If strSubWhere <> "" Then
        strSubWhere = Mid(Trim(strSubWhere), 4)
        gstrSQL = "" & _
        "   Select A.ID, A.������, A.����, A.���, A.����, A.�������, A.�ɷ��ֵ, A.��Ч��, A.����ԭ��, A.������, A.�쿨��, " & _
        "          A.����ʱ��, A.������, A.����ʱ��, A.��ǰ״̬, A.��ע, A.������, A.���۽��, A.��ֵ�ۿ���, A.���, A.ͣ����, " & _
        "          A.ͣ������,decode(A1.����,NULL,'',A1.����||'-'||A1.����) as �쿨����,decode(C.���ѿ�id,NULL,0,1) as ����" & _
        "   From ���ѿ�Ŀ¼ A,���ű� A1, " & _
        "        (Select �ӿڱ��,���ѿ�id From ���˿������¼  where �ӿڱ��=[10] Group By �ӿڱ��,���ѿ�id Having Count(*)>0) C, " & _
        "        (Select ���� ,max(���) as ���  From ���ѿ�Ŀ¼  Where " & strSubWhere & "  Group by  ����) B" & _
        "   Where  a.id=c.���ѿ�ID(+) and a.�쿨����id=A1.ID(+) And c.�ӿڱ��(+)=[10] And  A.���� = B.���� and a.���=b.��� and A.�ӿڱ��=[10]  " & strWhere
    Else
        gstrSQL = "" & _
        "   Select A.ID, A.������, A.����, A.���, A.����, A.�������, A.�ɷ��ֵ, A.��Ч��, A.����ԭ��, A.������, A.�쿨��, " & _
        "          A.����ʱ��, A.������, A.����ʱ��, A.��ǰ״̬, A.��ע, A.������, A.���۽��, A.��ֵ�ۿ���, A.���, A.ͣ����, " & _
        "          A.ͣ������,decode(A1.����,NULL,'',A1.����||'-'||A1.����) as �쿨����,decode(C.���ѿ�id,NULL,0,1) as ����" & _
        "   From ���ѿ�Ŀ¼ A,���ű� A1, " & _
        "        (Select �ӿڱ��,���ѿ�id From ���˿������¼  where �ӿڱ��=[10] Group By �ӿڱ��,���ѿ�id Having Count(*)>0) C, " & _
        "   Where  a.id=c.���ѿ�ID(+) and a.�쿨����id=A1.ID(+) and A.�ӿڱ��=[10]   " & strWhere
    End If
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
        CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
        CStr(mArrFilter("���ŷ�Χ")(0)), CStr(mArrFilter("���ŷ�Χ")(1)), _
        CStr(mArrFilter("�쿨��")), CStr(mArrFilter("������")), _
        CStr(mArrFilter("������")), mlng�ӿڱ��)
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    With vsCardList
        If .Row > 0 Then
            lngPre����ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
        End If
        .Redraw = flexRDNone
        .Clear 1
        .Rows = 2
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .Cell(flexcpData, lngRow, .ColIndex("����")) = Nvl(rsTemp!id)
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .Cell(flexcpData, .Row, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = "********"
            
            .TextMatrix(lngRow, .ColIndex("��ֵ��")) = IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, "��", "")
            
            
            .TextMatrix(lngRow, .ColIndex("��Ч��")) = Format(rsTemp!��Ч��, "yyyy-mm-dd HH:MM:SS")
            If Trim(.TextMatrix(lngRow, .ColIndex("��Ч��"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("��Ч��")) = ""
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ԭ��")) = Nvl(rsTemp!����ԭ��)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("�쿨��")) = Nvl(rsTemp!�쿨��)
            .TextMatrix(lngRow, .ColIndex("�쿨����")) = Nvl(rsTemp!�쿨����)
            
            
            .TextMatrix(lngRow, .ColIndex("ͣ����")) = Nvl(rsTemp!ͣ����)
            .TextMatrix(lngRow, .ColIndex("ͣ������")) = Format(rsTemp!ͣ������, "yyyy-mm-dd HH:MM:SS")
            If Trim(.TextMatrix(lngRow, .ColIndex("ͣ������"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("ͣ������")) = ""
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM:SS")
            If Trim(.TextMatrix(lngRow, .ColIndex("����ʱ��"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("����ʱ��")) = ""
            
            .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("��ֵ�ۿ���")) = Format(rsTemp!��ֵ�ۿ���, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("���")) = Format(rsTemp!������, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("���۽��")) = Format(rsTemp!���۽��, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("��ǰ���")) = Format(rsTemp!���, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("������")) = IIf(Val(Nvl(rsTemp!����)) = 1, "��", "")
            .TextMatrix(lngRow, .ColIndex("��ע")) = Nvl(rsTemp!��ע)
            .TextMatrix(lngRow, .ColIndex("��ǰ״̬")) = ""
            '1-��Ч, 2-����,3-�˿�
            .Cell(flexcpData, .Row, .ColIndex("��Ч��")) = ""
            '1-��Ч, 2-����,3-�˿�,4-ʧЧ,8-ͣ��
            If Format(rsTemp!��Ч��, "yyyy-mm-dd HH:MM:SS") <= strCurDate Then
                .TextMatrix(lngRow, .ColIndex("��ǰ״̬")) = "ʧЧ"
                .Cell(flexcpData, lngRow, .ColIndex("��ǰ״̬")) = 4
                .Cell(flexcpData, lngRow, .ColIndex("��Ч��")) = "4"   'ʧЧ��
            Else
                .TextMatrix(lngRow, .ColIndex("��ǰ״̬")) = Decode(Val(Nvl(rsTemp!��ǰ״̬)), 1, "��Ч", 2, "����", 3, "�˿�", 4, "ʧЧ", ""):
                .Cell(flexcpData, lngRow, .ColIndex("��ǰ״̬")) = Val(Nvl(rsTemp!��ǰ״̬)) Mod 10
            End If
            If lngPre����ID = Val(Nvl(rsTemp!id)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            '������ɫ��
            Call SetGridRowForeColor(lngRow)
            SetCardRowColHide lngRow
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, True
    
    Call vsCardList_AfterRowColChange(-1, 0, vsCardList.Row, 0)
    LoadDataToRpt = True
    Exit Function
Errhand:
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetGridRowForeColor(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����״̬����������ɫ
    '����:���˺�
    '����:2009-11-20 15:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long, int״̬ As Integer
    With vsCardList
        If .TextMatrix(lngRow, .ColIndex("ͣ������")) <> "" Then
            lngColor = vbRed
        ElseIf Val(.Cell(flexcpData, lngRow, .ColIndex("��Ч��"))) = 4 Then
            lngColor = &HFF00FF
        Else
            Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("��ǰ״̬")))
            Case 2, 3
                  lngColor = vbBlue
            Case Else
                '1-��Ч, 2-����,3-�˿�,4-ʧЧ,8-ͣ��
                lngColor = &H80000008
            End Select
        End If
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
    End With
End Sub

Private Sub mfrmSquareCardConsume_zlDblClick(ByVal lng����ID As Long, ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    If InStr(1, mstrPrivs, ";������������ϸ��;") = 0 Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_INSIDE_1503_1", Me, "������ID=" & lng����ID, 1)
End Sub
 

Private Sub mfrmSquareCardInFull_AfterRowChange(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    mTy_CardStartus.bln�����ֵ���� = mfrmSquareCardInFull.zl�������
    
End Sub

Private Sub mfrmSquareCardInFull_zlPopupMenus(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    '�����˵�:��ֵ���
    Dim cbrPopupBar As CommandBar, cbrPopupItem As CommandBarControl
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    With cbrPopupBar.Controls
        If zlCheckPrivs(mstrPrivs, "��ֵ") Then Set cbrPopupItem = .Add(xtpControlButton, conMenu_Edit_CardInFull, "��ֵ(&C)")
        If zlCheckPrivs(mstrPrivs, "����") Then Set cbrPopupItem = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "��ֵ����(&T)"): cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    cbrPopupBar.ShowPopup
End Sub

 
Private Sub picCardList_Resize()
    Err = 0: On Error Resume Next
    With picCardList
        vsCardList.Left = .ScaleLeft
        vsCardList.Width = .ScaleWidth
        vsCardList.Height = .ScaleHeight - vsCardList.Top
        picModify.Width = .ScaleWidth - picModify.Left - 50
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub zlSetInitCardCustomType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ�ĵ�ǰ״̬
    '����:���˺�
    '����:2009-11-19 14:46:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    With mTy_CardStartus
        .blnHaveData = False: .bln���� = False: .bln���� = False
        .bln�˿� = False: .bln�û�ʹ����Ч�� = False: .bln�������� = False
        .blnͣ�ÿ� = False
    End With
    
    With vsCardList
        If .Rows < 2 Then Exit Sub
        If .Row < 1 Then Exit Sub
        mTy_CardStartus.blnHaveData = Trim(.TextMatrix(1, .ColIndex("����"))) <> ""
        '1-��Ч, 2-����,3-�˿�,4-ʧЧ,8-ͣ��
        mTy_CardStartus.bln���� = Val(.Cell(flexcpData, .Row, .ColIndex("��ǰ״̬"))) = 2
        mTy_CardStartus.bln�˿� = Val(.Cell(flexcpData, .Row, .ColIndex("��ǰ״̬"))) = 3
        mTy_CardStartus.bln���� = Val(.Cell(flexcpData, .Row, .ColIndex("��Ч��"))) = 1 'ʧЧ�˵Ŀ�
        mTy_CardStartus.bln�û�ʹ����Ч�� = Val(.Cell(flexcpData, .Row, .ColIndex("��ǰ״̬"))) = 1   '������Чʹ�õĿ�(���û���ʹ�õĿ�)
        mTy_CardStartus.bln�������� = Val(.Cell(flexcpData, .Row, .ColIndex("������"))) = 1
        mTy_CardStartus.blnͣ�ÿ� = Trim(.TextMatrix(.Row, .ColIndex("ͣ������"))) <> ""
    End With
    '�������ֵ����:
    
End Sub

Private Sub picModify_Click()
    Err = 0: On Error Resume Next
    With picModify
        cmdModify.Left = .ScaleWidth - cmdModify.Width - 50
        txtEdit.Width = cmdModify.Left - txtEdit.Left
        cmdSel.Left = txtEdit.Left + txtEdit.Width - cmdSel.Width
        cbo������.Width = txtEdit.Width
        dtp����Ч����.Width = txtEdit.Width
    End With
End Sub

Private Sub txtEdit_Change()
    txtEdit.Tag = ""
End Sub

Private Sub txtEdit_GotFocus()
    zlControl.TxtSelAll txtEdit
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit.Text <> "" And txtEdit.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select�շ����ѡ����(txtEdit, Trim(txtEdit.Text)) = False Then Exit Sub
    zlCtlSetFocus txtEdit
End Sub

Private Sub vsCardList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng���ѿ�ID As Long
    zl_VsGridRowChange vsCardList, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldCol <> NewCol Then
        Call SetModiyCaption: Call SetModifyEnabled
    End If
    
    If OldRow = NewRow Then Exit Sub
    zlCommFun.ShowFlash "����װ������,���Ժ�..."
    With vsCardList
        lng���ѿ�ID = Val(.Cell(flexcpData, NewRow, .ColIndex("����")))
        Call mfrmSquareCardCallBack.zlReLoadData(mlng�ӿڱ��, lng���ѿ�ID)  '���ռ�¼
        Call mfrmSquareCardInFull.zlReLoadData(mlng�ӿڱ��, lng���ѿ�ID) '��ֵ��¼
        Call mfrmSquareCardConsume.zlReLoadData(mlng�ӿڱ��, lng���ѿ�ID)   '���Ѽ�¼
    End With
    '�����е���Ϣ
    Call zlSetInitCardCustomType
    Call SetModifyDefaultValue
    zlCommFun.StopFlash
End Sub

Private Sub vsCardList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlCheckPrivs(mstrPrivs, "��������")
End Sub
Private Sub vsCardList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlCheckPrivs(mstrPrivs, "��������")
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:
    '����:
    '����:���˺�
    '����:2009-11-20 15:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim blnCardList As Boolean
    blnCardList = Me.ActiveControl Is vsCardList
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "���ѿ����"
    
    If CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "����ʱ�䣺" & CStr(mArrFilter("����ʱ��")(0)) & "��" & CStr(mArrFilter("����ʱ��")(1))
    End If
    If CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "����ʱ�䣺" & CStr(mArrFilter("����ʱ��")(0)) & "��" & CStr(mArrFilter("����ʱ��")(1))
    End If
    If objRow.count > 1 Then
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
    End If
    If mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        objRow.Add "���ŷ�Χ��" & CStr(mArrFilter("���ŷ�Χ")(0)) & "��" & CStr(mArrFilter("���ŷ�Χ")(1))
    ElseIf mArrFilter("���ŷ�Χ")(0) = "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        objRow.Add "���ţ�" & CStr(mArrFilter("���ŷ�Χ")(1))
    ElseIf mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) = "" Then
        objRow.Add "���ţ�" & CStr(mArrFilter("���ŷ�Χ")(0))
    End If
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    
    If mArrFilter("�쿨��") <> "" Then objRow.Add "�쿨�ˣ�" & mArrFilter("�쿨��")
    If mArrFilter("������") <> "" Then objRow.Add "�����ˣ�" & mArrFilter("������")
    If mArrFilter("������") <> "" Then objRow.Add "�����ͣ�" & mArrFilter("������")
    If Val(mArrFilter("����ͣ�ÿ�")) = 1 Then objRow.Add "����ͣ�ÿ�"
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsCardList
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
            
        Next
    End With
    
    Err = 0: On Error GoTo Errhand:
    Set objPrint.Body = vsCardList
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    '�ָ�
    With vsCardList
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Function SaveCardResumeAndStop(ByVal blnStop As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ƭͣ�û�����
    '���:blnStop-ͣ�ÿ�Ƭ
    '����:���˺�
    '����:2009-12-14 09:45:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String, lngID As Long, lngRow As Long, i As Long
    With vsCardList
        lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
        If lngID <= 0 Then Exit Function
        If blnStop Then 'ͣ��
        
            If .TextMatrix(.Row, .ColIndex("ͣ������")) <> "" Then Exit Function
            If MsgBox("�����Ҫ�Կ���Ϊ:��" & .TextMatrix(.Row, .ColIndex("����")) & "���ļ�¼����ͣ�ò�����" & vbCrLf & _
                        "   ���ǡ�: ����ͣ�ò���,ͣ�ú�Ŀ�Ƭ�����ܽ���ˢ�����ѻ����ٷ�����" & vbCrLf & _
                        "   ����:��������ͣ�ò���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            If .TextMatrix(.Row, .ColIndex("ͣ������")) = "" Then Exit Function
            If MsgBox("�����Ҫ�Կ���Ϊ:��" & .TextMatrix(.Row, .ColIndex("����")) & "���ļ�¼�������ò�����" & vbCrLf & _
                        "   ���ǡ�: �������ò���,���ú�Ŀ�Ƭ���ܽ���ˢ�����ѻ���ջ����Ŀ�Ƭ���ٷ���������" & vbCrLf & _
                        "   ����:�������������ò���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End With
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
   Err = 0: On Error GoTo Errhand:
    ' Zl_���ѿ�Ŀ¼_Stopandresume
    gstrSQL = "Zl_���ѿ�Ŀ¼_Stopandresume("
    '  Id_In       In ���ѿ�Ŀ¼.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  ͣ����_In   In ���ѿ�Ŀ¼.ͣ����%Type,
    gstrSQL = gstrSQL & IIf(blnStop = False, "NULL", "'" & UserInfo.���� & "'") & ","
    '  ͣ������_In In ���ѿ�Ŀ¼.ͣ������%Type
    gstrSQL = gstrSQL & IIf(blnStop = False, "NULL", "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    With vsCardList
         If blnStop Then 'ͣ��
            If Val(mArrFilter("����ͣ�ÿ�")) = 1 Then
                .TextMatrix(.Row, .ColIndex("ͣ������")) = strDate
                .TextMatrix(.Row, .ColIndex("ͣ����")) = UserInfo.����
            Else
                lngRow = .Row
                If .Rows - 1 <= 2 Then
                      For i = 0 To .Cols - 1
                        .TextMatrix(.Rows - 1, i) = ""
                        .Cell(flexcpData, .Rows - 1, i) = ""
                      Next
                Else 'ɾ����
                     .RemoveItem lngRow
                     If lngRow < .Rows - 1 Then
                        .Row = lngRow
                     Else
                        .Row = .Rows - 1
                     End If
                End If
            End If
        Else
            .TextMatrix(.Row, .ColIndex("ͣ������")) = ""
            .TextMatrix(.Row, .ColIndex("ͣ����")) = ""
        End If
        Call SetGridRowForeColor(.Row)
    End With
    Call zlSetInitCardCustomType
    SaveCardResumeAndStop = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
 
Private Function SaveBatchUpdateCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������¿�Ƭ��Ϣ
    '����:���³ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-05 12:02:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, blnIsdate As Boolean, cllPro As New Collection, strFieldValue As String
    Dim strSQL As String, strIDIn As String, lngRow As Long
    
    With vsCardList
        Select Case .Col
        Case .ColIndex("��Ч��")
           strFields = "��Ч��": blnIsdate = True:
           If IsNull(dtp����Ч����.value) Then
                strFieldValue = "3000-01-01 00:00:00"
           Else
                strFieldValue = Format(dtp����Ч����.value, "yyyy-mm-dd HH:MM:SS")
           End If
        Case .ColIndex("������")
           strFields = "������": blnIsdate = False:
           If cbo������.ListIndex < 0 Then
                ShowMsgbox "������δѡ��,��ѡ������"
                Exit Function
           End If
           strFieldValue = Mid(cbo������.Text, InStr(1, cbo������.Text, ".") + 1)
        Case .ColIndex("�������")
           strFields = "�������": blnIsdate = False:
           If txtEdit.Tag = "" And txtEdit.Text <> "" Then
                ShowMsgbox "�������ѡ�����,����!"
                Exit Function
           End If
           strFieldValue = Trim(txtEdit.Text)
        Case Else
           Exit Function
        End Select
        If MsgBox("���Ƿ����Ҫ�����޸ġ�" & strFields & "����ֵ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End With
    Err = 0: On Error GoTo Errhand:
    strIDIn = ""
    With vsCardList
        For lngRow = 1 To .Rows - 1
            If zlCommFun.ActualLen(strIDIn) >= 3980 Then
                'Zl_���ѿ�Ŀ¼_Batch_Update
                gstrSQL = "Zl_���ѿ�Ŀ¼_Batch_Update("
                '  Ids_In    Varchar2,
                gstrSQL = gstrSQL & "'" & Mid(strIDIn, 2) & "',"
                '  �ֶ�_In   Varchar2,
                gstrSQL = gstrSQL & "'" & strFields & "',"
                '  �ֶ�ֵ_In Varchar2,
                gstrSQL = gstrSQL & "'" & strFieldValue & "',"
                '  IsDate Number:=0
                gstrSQL = gstrSQL & " " & IIf(blnIsdate, 1, 0) & ")"
                AddArray cllPro, gstrSQL
                strIDIn = ""
            End If
            If Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) <> 0 And .RowHidden(lngRow) = False Then
                strIDIn = strIDIn & "," & Val(.Cell(flexcpData, lngRow, .ColIndex("����")))
            End If
        Next
    End With
    If strIDIn <> "" Then
        'Zl_���ѿ�Ŀ¼_Batch_Update
        gstrSQL = "Zl_���ѿ�Ŀ¼_Batch_Update("
        '  Ids_In    Varchar2,
        gstrSQL = gstrSQL & "'" & Mid(strIDIn, 2) & "',"
        '  �ֶ�_In   Varchar2,
        gstrSQL = gstrSQL & "'" & strFields & "',"
        '  �ֶ�ֵ_In Varchar2,
        gstrSQL = gstrSQL & "'" & strFieldValue & "',"
        '  IsDate Number:=0
        gstrSQL = gstrSQL & " " & IIf(blnIsdate, 1, 0) & ")"
        AddArray cllPro, gstrSQL
        strIDIn = ""
    End If
    If cllPro.count = 0 Then
        Exit Function
    End If
    ExecuteProcedureArrAy cllPro, Me.Caption
    
    '��������
    With vsCardList
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) <> 0 And .RowHidden(lngRow) = False Then
                Select Case .Col
                  Case .ColIndex("��Ч��")
                     If IsNull(dtp����Ч����.value) Then
                          .TextMatrix(lngRow, .Col) = ""
                     Else
                          .TextMatrix(lngRow, .Col) = Format(dtp����Ч����.value, "yyyy-mm-dd HH:MM:SS")
                     End If
                  Case .ColIndex("������")
                       .TextMatrix(lngRow, .Col) = Mid(cbo������.Text, InStr(1, cbo������.Text, "-") + 1)
                  Case .ColIndex("�������")
                      .TextMatrix(lngRow, .Col) = Trim(txtEdit.Text)
                  Case Else
                     Exit Function
                  End Select
            End If
        Next
    End With
    SaveBatchUpdateCardInfor = True
    MsgBox "�޸ĳɹ�!"
    
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Function Select�շ����ѡ����(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����: �շ����ѡ����
    '���::objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-01 14:18:58
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ�
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strTittle = "�շ����ѡ����"
    vRect = GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSearch, False)
    
    gstrSQL = "" & _
        " Select rownum as ID,����,����,���� From �շ���Ŀ���"
    If strSearch <> "" Then
        gstrSQL = gstrSQL & _
        "           Where ( ���� like upper([1]) or ���� like [1] or ���� like upper([1]) )"
    End If
    gstrSQL = gstrSQL & vbCrLf & " Order by ����"
  
    Set rsTemp = frmItemSelectMulit.ShowSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, True, strKey)
 
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgbox "û�������������շ����,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If objCtl.Enabled Then objCtl.SetFocus
    With rsTemp
        objCtl.Tag = ""
        Do While Not .EOF
            strTemp = strTemp & "," & Nvl(rsTemp!����)
            objCtl.Tag = objCtl.Tag & "," & Nvl(rsTemp!����)
            .MoveNext
        Loop
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strKey = objCtl.Tag
    objCtl.Text = strTemp
    objCtl.Tag = strKey
    zlCommFun.PressKey vbKeyTab
    Select�շ����ѡ���� = True
End Function

Private Sub vsCardList_DblClick()
    Dim lngID As Long
    With vsCardList
        lngID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
        If lngID <= 0 Then Exit Sub
    End With
    If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_��ѯ, mlng�ӿڱ��, lngID) = False Then Exit Sub
End Sub

Private Sub vsCardList_GotFocus()
    zl_VsGridGotFocus vsCardList, gSysColor.lngGridColorSel
End Sub

Private Sub vsCardList_LostFocus()
    zl_VsGridLOSTFOCUS vsCardList, gSysColor.lngGridColorLost
End Sub
