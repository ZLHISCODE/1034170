VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCISAduitPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmCISAduitPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   0
      Left            =   255
      ScaleHeight     =   5415
      ScaleWidth      =   5880
      TabIndex        =   22
      Top             =   315
      Width           =   5880
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2055
         Width           =   1920
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   7
         Left            =   3360
         TabIndex        =   9
         Top             =   2445
         Width           =   750
      End
      Begin VB.CheckBox ChkEnter 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3060
         TabIndex        =   29
         Top             =   3420
         Width           =   285
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1665
         Width           =   1920
      End
      Begin VB.CheckBox chkAudit 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3060
         TabIndex        =   13
         Top             =   3075
         Width           =   660
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   930
         TabIndex        =   24
         Top             =   165
         Width           =   4815
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   870
         Width           =   1920
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1260
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1605
         TabIndex        =   11
         Top             =   2775
         Width           =   435
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1260
         Width           =   1920
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   0
         Left            =   1260
         TabIndex        =   23
         Top             =   3735
         Width           =   4815
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "����(&D)"
         Height          =   350
         Left            =   4425
         TabIndex        =   16
         Top             =   4440
         Width           =   1100
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "����(&U)"
         Height          =   350
         Left            =   4425
         TabIndex        =   15
         Top             =   3975
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfNo 
         Height          =   1395
         Left            =   1170
         TabIndex        =   14
         Top             =   3975
         Width           =   3135
         _cx             =   5530
         _cy             =   2461
         Appearance      =   0
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&4.סԺҽ����ӡ"
         Height          =   180
         Index           =   9
         Left            =   1020
         TabIndex        =   32
         Top             =   2100
         Width           =   1260
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&8.��������¼��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   1020
         TabIndex        =   30
         Top             =   3450
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&3.����ʱ��"
         Height          =   180
         Index           =   7
         Left            =   1020
         TabIndex        =   6
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&7.�������������ܹ鵵"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   1020
         TabIndex        =   12
         Top             =   3120
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&1.�ύʱ��"
         Height          =   180
         Index           =   1
         Left            =   1020
         TabIndex        =   0
         Top             =   930
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmCISAduitPara.frx":000C
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ���Ӳ�������ȱʡʱ�䷶Χ��"
         Height          =   405
         Left            =   1035
         TabIndex        =   27
         Top             =   555
         Width           =   4065
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   195
         TabIndex        =   26
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&5.�����������ȱʡ����Ϊ         ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1020
         TabIndex        =   8
         Top             =   2460
         Width           =   3330
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&2.�鵵����"
         Height          =   180
         Index           =   0
         Left            =   1020
         TabIndex        =   2
         Top             =   1305
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&6.ÿ��     �����Զ�ˢ�µȴ���������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   1020
         TabIndex        =   10
         Top             =   2790
         Width           =   3150
      End
      Begin VB.Line ln 
         Index           =   0
         X1              =   1605
         X2              =   2055
         Y1              =   2985
         Y2              =   2985
      End
      Begin VB.Line ln 
         Index           =   1
         X1              =   3375
         X2              =   4125
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&2.��Ժ����"
         Height          =   180
         Index           =   4
         Left            =   1020
         TabIndex        =   4
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������˳��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   195
         TabIndex        =   25
         Top             =   3735
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4935
      TabIndex        =   19
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   18
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   17
      Top             =   6000
      Width           =   1100
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Index           =   1
      Left            =   7005
      ScaleHeight     =   4980
      ScaleWidth      =   5835
      TabIndex        =   20
      Top             =   1260
      Width           =   5835
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1680
         Index           =   0
         Left            =   660
         TabIndex        =   21
         Top             =   285
         Width           =   2820
         _cx             =   4974
         _cy             =   2963
         Appearance      =   0
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
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
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   5805
      Left            =   180
      TabIndex        =   28
      Top             =   75
      Width           =   6120
      _Version        =   589884
      _ExtentX        =   10795
      _ExtentY        =   10239
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmCISAduitPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mblnOK As Boolean
Private mfrmMain As Object
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mclsVsfNo As clsVsf
Private mstrPrivs As String

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    If ExecuteCommand("��ȡ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim varTmp As Variant
    Dim varAry As Variant
    Dim blnAllowModify As Boolean

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "��ʼ����"
            
            Set mclsVsf = New clsVsf
            With mclsVsf
                Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)

                Call .AppendColumn("����", 1590, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("����", 3690, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
                
                If IsPrivs(mstrPrivs, "��������") Then
                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                    Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                    .IndicatorCol = 0
                    Set .IndicatorIcon = GetImageList(16).ListImages("��ǰ").Picture
                End If
                
                .AppendRows = True
            End With
            
            '----------------------------------------------------------------------------------------------------------
            Set mclsVsfNo = New clsVsf
            With mclsVsfNo
                Call .Initialize(Me.Controls, vsfNo, True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("����", 1590, flexAlignLeftCenter, flexDTString, "", , True)
                With vsfNo
                    .Rows = 10
                    .RowHidden(0) = True
                End With
                .AppendRows = True
            End With

            
            '----------------------------------------------------------------------------------------------------------
            With tbc
                With .PaintManager
                    .Appearance = xtpTabAppearancePropertyPage2003
                    .BoldSelected = True
                    .COLOR = xtpTabColorDefault
                    .ColorSet.ButtonSelected = COLOR.��ɫ
                    .ShowIcons = True
                End With
                
                .InsertItem 0, "���� ", picPane(0).hWnd, 0
                .InsertItem 1, "�����ҷ�Χ ", picPane(1).hWnd, 0
                .Item(0).Selected = True
            End With
            
            For intCount = 0 To 3
                With cbo(intCount)
                    .Clear
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "������"
                    .AddItem "��  ��"
                    .AddItem "ǰ����"
                    .AddItem "ǰһ��"
                    .AddItem "ǰ����"
                    .AddItem "ǰһ��"
                    .AddItem "ǰ����"
                    .AddItem "ǰ����"
                    .AddItem "ǰ����"
                    .AddItem "ǰһ��"
                    .AddItem "ǰ����"
                End With
            Next
            
            With cbo(4)
                .Clear
                .AddItem "����ҽ����"
                .AddItem "����ҽ����"
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case "�ؼ�״̬"
            
        '--------------------------------------------------------------------------------------------------------------
        Case "��ȡ����"
            
            On Error Resume Next
            chkAudit.Value = zlDatabase.GetPara("���ղ��ܹ鵵", ParamInfo.ϵͳ��, mfrmMain.ģ���, "0", Array(chkAudit), IsPrivs(mstrPrivs, "��������"))
            ChkEnter.Value = zlDatabase.GetPara("��������¼��������", ParamInfo.ϵͳ��, mfrmMain.ģ���, "0", Array(ChkEnter), IsPrivs(mstrPrivs, "��������"))
            cbo(0).Text = zlDatabase.GetPara("���ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(0)), IsPrivs(mstrPrivs, "��������"))
            cbo(1).Text = zlDatabase.GetPara("�鵵ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(1)), IsPrivs(mstrPrivs, "��������"))
            cbo(2).Text = zlDatabase.GetPara("��Ժȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(2)), IsPrivs(mstrPrivs, "��������"))
            cbo(3).Text = zlDatabase.GetPara("ҽ��ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(3)), IsPrivs(mstrPrivs, "��������"))
            cbo(4).Text = zlDatabase.GetPara("סԺҽ����ӡ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "����ҽ����", Array(cbo(4)), IsPrivs(mstrPrivs, "��������"))
            On Error GoTo errHand
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
            If cbo(2).ListCount > 0 And cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
            If cbo(3).ListCount > 0 And cbo(3).ListIndex = -1 Then cbo(3).ListIndex = 0
            If cbo(4).ListCount > 0 And cbo(4).ListIndex = -1 Then cbo(4).ListIndex = 0
            
            txt(7).Text = Val(zlDatabase.GetPara("������������", ParamInfo.ϵͳ��, mfrmMain.ģ���, "7", Array(txt(7)), IsPrivs(mstrPrivs, "��������")))
            txt(0).Text = Val(zlDatabase.GetPara("δ����ˢ��Ƶ��", ParamInfo.ϵͳ��, mfrmMain.ģ���, "5", Array(txt(0)), IsPrivs(mstrPrivs, "��������")))
            
            strTmp = Trim(zlDatabase.GetPara("��������˳��", ParamInfo.ϵͳ��, mfrmMain.ģ���, "5;1;6;2;3;4;8;7;9", Array(vsfNo, cmdUp, cmdDown), IsPrivs(mstrPrivs, "��������")))
            If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
            varTmp = Split(strTmp, ";")
            With vsfNo
                '1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-ҽ������;7-����֤��;8-֪���ļ�
                For intCount = 0 To UBound(varTmp)
                    Select Case varTmp(intCount)
                    Case "1"
                        .TextMatrix(intCount + 1, 0) = "סԺҽ��"
                        .RowData(intCount + 1) = 1
                    Case "2"
                        .TextMatrix(intCount + 1, 0) = "סԺ����"
                        .RowData(intCount + 1) = 2
                    Case "3"
                        .TextMatrix(intCount + 1, 0) = "������"
                        .RowData(intCount + 1) = 3
                    Case "4"
                        .TextMatrix(intCount + 1, 0) = "�����¼"
                        .RowData(intCount + 1) = 4
                    Case "5"
                        .TextMatrix(intCount + 1, 0) = "��ҳ��¼"
                        .RowData(intCount + 1) = 5
                    Case "6"
                        .TextMatrix(intCount + 1, 0) = "ҽ������"
                        .RowData(intCount + 1) = 6
                    Case "7"
                        .TextMatrix(intCount + 1, 0) = "����֤��"
                        .RowData(intCount + 1) = 7
                    Case "8"
                        .TextMatrix(intCount + 1, 0) = "֪���ļ�"
                        .RowData(intCount + 1) = 8
                    Case "9"
                        .TextMatrix(intCount + 1, 0) = "�ٴ�·��"
                        .RowData(intCount + 1) = 9
                    End Select
                Next
            End With
            
            strTmp = Trim(zlDatabase.GetPara("�����ҷ�Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "", Array(vsf(0)), IsPrivs(mstrPrivs, "��������")))
            gstrSQL = "Select ID,���,���� From ��Ա��"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            gstrSQL = "Select a.ID,a.����,a.���� From ���ű� a,��������˵�� b Where a.ID=b.����id And b.��������='�ٴ�' And ( TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null)"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            With vsf(0)
                .Rows = 2
                If strTmp <> "" Then
                    
                    varTmp = Split(strTmp, ";")
                    For intCount = 0 To UBound(varTmp)
                        varAry = Split(varTmp(intCount), ",")
                        rs.Filter = ""
                        rs.Filter = "ID=" & Val(varAry(0))
                        If rs.RecordCount > 0 Then
                            
                            If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, .ColIndex("����")) = AppendCode(rs("����").Value, rs("���").Value)
                            .RowData(.Rows - 1) = rs("ID").Value
                            
                            For intCol = 1 To UBound(varAry)
                                rsTmp.Filter = ""
                                rsTmp.Filter = "ID=" & Val(varAry(intCol))
                                If rsTmp.RecordCount > 0 Then
                                    If .TextMatrix(.Rows - 1, .ColIndex("����")) = "" Then
                                        .TextMatrix(.Rows - 1, .ColIndex("����")) = AppendCode(rsTmp("����").Value, rsTmp("����").Value)
                                        .TextMatrix(.Rows - 1, .ColIndex("����id")) = rsTmp("ID").Value
                                    Else
                                        .TextMatrix(.Rows - 1, .ColIndex("����")) = .TextMatrix(.Rows - 1, .ColIndex("����")) & vbCrLf & AppendCode(rsTmp("����").Value, rsTmp("����").Value)
                                        .TextMatrix(.Rows - 1, .ColIndex("����id")) = .TextMatrix(.Rows - 1, .ColIndex("����id")) & "," & rsTmp("ID").Value
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
                
                .AutoSize .ColIndex("����"), .ColIndex("����")
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case "У������"
            
        '--------------------------------------------------------------------------------------------------------------
        Case "��������"
            
            Call SetPara("���ȱʡ��Χ", cbo(0).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("�鵵ȱʡ��Χ", cbo(1).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("��Ժȱʡ��Χ", cbo(2).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("ҽ��ȱʡ��Χ", cbo(3).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("סԺҽ����ӡ", cbo(4).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("������������", Val(txt(7).Text), mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("δ����ˢ��Ƶ��", Val(txt(0).Text), mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("���ղ��ܹ鵵", chkAudit.Value, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("��������¼��������", ChkEnter.Value, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            strTmp = ""
            With vsfNo
                For intCount = 1 To .Rows - 1
                    If Val(.RowData(intCount)) > 0 Then
                        strTmp = strTmp & ";" & Val(.RowData(intCount))
                    End If
                Next
            End With
            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
            Call SetPara("��������˳��", strTmp, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            
            strTmp = ""
            With vsf(0)
                For intCount = 1 To .Rows - 1
                    If Val(.RowData(intCount)) > 0 Then
                        strTmp = strTmp & ";" & Val(.RowData(intCount)) & "," & Trim(.TextMatrix(intCount, .ColIndex("����id")))
                    End If
                Next
            End With
            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
            If Len(strTmp) > 2000 Then
                ShowSimpleMsg "������Ȩ��̫�࣬�����˲���ֵ�����洢��Χ��"
                Exit Function
            End If
            Call SetPara("�����ҷ�Χ", strTmp, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            
        End Select
    Next

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    cmdOk.Tag = IIf(blnData, "Changed", "")
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = (cmdOk.Tag = "Changed")
End Property

'######################################################################################################################

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkAudit_Click()
    DataChanged = True
End Sub

Private Sub ChkEnter_Click()
    DataChanged = True
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If vsfNo.Row < vsfNo.Rows - 1 Then
        Call mclsVsfNo.MoveRow(vsfNo.Row, 1)
        vsfNo.Row = vsfNo.Row + 1
        DataChanged = True
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    
    If DataChanged Then
        If ExecuteCommand("У������") = False Then Exit Sub
        
        If ExecuteCommand("��������") Then
            
            DataChanged = False
            
            mblnOK = True
        Else
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub


Private Sub cmdUp_Click()
    If vsfNo.Row > 1 Then
        Call mclsVsfNo.MoveRow(vsfNo.Row, -1)
        vsfNo.Row = vsfNo.Row - 1
        DataChanged = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("�������޸ĵĲ������뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
    End If
    
    Set mclsVsf = Nothing
    Set mclsVsfNo = Nothing
    
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    
    With vsf(0)
        Cancel = Not (Val(.RowData(Row)) > 0 And Trim(.TextMatrix(Row, .ColIndex("����id"))) <> "")
        If Cancel = False Then DataChanged = True
    End With
    DataChanged = True
End Sub

Private Sub picPane_Resize(Index As Integer)
    Select Case Index
    Case 1
        
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf.AppendRows = True
        
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.AfterEdit(Row, Col)
    End Select
    
    DataChanged = True
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Select Case Index
    Case 0
        mclsVsf.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
    Case 0
        mclsVsf.AppendRows = True
    End Select
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim varTmp As Variant
    Dim bytRet As Byte
    Dim strTmp As String
    Dim strTmpID As String
    Dim intCount As Integer
    
    With vsf(Index)
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
            Select Case Col
            '----------------------------------------------------------------------------------------------------------
            Case .ColIndex("����")
                
                Set rsData = gclsPackage.GetOperationPerson
                bytRet = ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\�����Աѡ��", "����±���ѡ��һ�������Ա", rsData, rs, 8790, 4500, False, Val(.RowData(Row)))
                            
                If bytRet = 1 Then
                    
                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value), False) = False Then
                        
                        .EditText = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .TextMatrix(Row, .ColIndex("����")) = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        DataChanged = True
                    End If

                    
                    mclsVsf.AppendRows = True
        
                End If
            '----------------------------------------------------------------------------------------------------------
            Case .ColIndex("����")
            
                Set rs = gclsPackage.GetDeptSelect
                Set rsData = CopyRecordStruct(rs)
                Call CopyRecordData(rs, rsData)
                
                If .TextMatrix(Row, .ColIndex("����id")) <> "" Then
                    varTmp = Split(.TextMatrix(Row, .ColIndex("����id")), ",")
                    For intCount = 0 To UBound(varTmp)
                        rsData.Filter = ""
                        rsData.Filter = "ID=" & Val(varTmp(intCount))
                        If rsData.RecordCount > 0 Then
                            rsData("ѡ��").Value = 1
                        End If
                    Next
                End If
                rsData.Filter = ""
                If rsData.RecordCount > 0 Then rsData.MoveFirst
                
                bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,1200,0,;����,900,0,", Me.Name & "\���˿���ѡ��", "����±���ѡ��һ���������˿���", rsData, rs, 8790, 4500, True)
                            
                If bytRet = 1 Then
                    
                    If rs.RecordCount > 0 Then rs.MoveFirst
                    strTmp = ""
                    strTmpID = ""
                    Do While Not rs.EOF
                        strTmp = strTmp & vbCrLf & AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("����").Value))
                        strTmpID = strTmpID & "," & zlCommFun.NVL(rs("ID").Value, 0)
                        rs.MoveNext
                    Loop
                    If strTmp <> "" Then strTmp = Mid(strTmp, 3)
                    If strTmpID <> "" Then strTmpID = Mid(strTmpID, 2)
                    
                    .EditText = strTmp
                    .TextMatrix(Row, .ColIndex("����")) = strTmp
                    .TextMatrix(Row, .ColIndex("����id")) = strTmpID
                    
                    DataChanged = True

                    .AutoSize .ColIndex("����"), .ColIndex("����")
                    mclsVsf.AppendRows = True
        
                End If
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case 1
            
        End Select
    End With
    DataChanged = True
    
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '�༭����
    Select Case Index
        Case 0
            Call mclsVsf.KeyDown(KeyCode, Shift)
    End Select
    If KeyCode = vbKeyDelete Then
        DataChanged = True
    End If
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim StrText As String
    Dim bytRet As Byte
    
    With vsf(Index)
        
        If InStr(.EditText, "'") > 0 Then
            KeyCode = 0
            .EditText = ""
            Exit Sub
        End If
                            
        StrText = .EditText
        
        Select Case Index
        '----------------------------------------------------------------------------------------------------------
        Case 0
            If KeyCode = vbKeyReturn Then
                If Col = .ColIndex("����") Then

                    Set rsData = gclsPackage.GetOperationPerson(UCase(StrText))
                    
                    If ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\�����Ա����", "����±���ѡ��һ�������Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row)), , True) = 1 Then
    
                        If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                               
                        .EditText = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .Cell(flexcpData, Row, Col) = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .TextMatrix(Row, .ColIndex("����")) = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        DataChanged = True
                    Else
                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    End If
    
                End If
            Else
                DataChanged = True
            End If

        End Select
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    
    '�༭����,������
    Select Case Index
    Case 0
        Call mclsVsf.KeyPress(KeyAscii)
    End Select
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.KeyPressEdit(KeyAscii)
    End Select
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case Index
        Case 0
            Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        End Select
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.EditSelAll
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.BeforeEdit(Row, Col, Cancel)
    End Select
End Sub


