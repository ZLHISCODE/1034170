VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl ClinicPlanUnit 
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ScaleHeight     =   9225
   ScaleWidth      =   12000
   Begin VB.PictureBox picPageSub 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   9420
      ScaleHeight     =   975
      ScaleWidth      =   1275
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1275
      Begin XtremeSuiteControls.TabControl tbPageSub 
         Height          =   930
         Index           =   1
         Left            =   60
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   30
         Width           =   1080
         _Version        =   589884
         _ExtentX        =   1905
         _ExtentY        =   1640
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPageSub 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   0
      Left            =   8100
      ScaleHeight     =   975
      ScaleWidth      =   1275
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1275
      Begin XtremeSuiteControls.TabControl tbPageSub 
         Height          =   930
         Index           =   0
         Left            =   60
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   30
         Width           =   1080
         _Version        =   589884
         _ExtentX        =   1905
         _ExtentY        =   1640
         _StockProps     =   64
      End
   End
   Begin VB.CheckBox chkOnlyOneUse 
      Caption         =   "��ռ��ʽ"
      Height          =   300
      Left            =   5310
      TabIndex        =   4
      Top             =   50
      Width           =   1035
   End
   Begin VB.PictureBox picFun 
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   7320
      ScaleHeight     =   4065
      ScaleWidth      =   765
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   765
      Begin VB.CommandButton cmdFun 
         Caption         =   "<<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   105
         TabIndex        =   11
         Top             =   1935
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   1465
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   ">>"
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   995
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   525
         Width           =   555
      End
   End
   Begin VB.PictureBox picUnit 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   0
      Left            =   8130
      ScaleHeight     =   4050
      ScaleWidth      =   2760
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   2760
      Begin VB.CheckBox chkForbidBespeak 
         Caption         =   "��ֹԤԼ"
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   1110
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSelNum 
         Height          =   3285
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   2175
         _cx             =   3836
         _cy             =   5794
         Appearance      =   2
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ClinicPlanUnit.ctx":0000
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
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   930
      Left            =   8100
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   1080
      _Version        =   589884
      _ExtentX        =   1905
      _ExtentY        =   1640
      _StockProps     =   64
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "������ԤԼ"
      Height          =   300
      Index           =   1
      Left            =   2460
      TabIndex        =   2
      Top             =   50
      Width           =   1200
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "������ԤԼ"
      Height          =   300
      Index           =   0
      Left            =   1215
      TabIndex        =   1
      Top             =   50
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "����ſ���ԤԼ"
      Height          =   300
      Index           =   2
      Left            =   3690
      TabIndex        =   3
      Top             =   50
      Width           =   1560
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfNotSelNum 
      Height          =   4065
      Left            =   4890
      TabIndex        =   6
      Top             =   2400
      Width           =   2385
      _cx             =   4207
      _cy             =   7170
      Appearance      =   2
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanUnit.ctx":0070
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
   Begin VSFlex8Ctl.VSFlexGrid vsUnit 
      Height          =   2865
      Left            =   90
      TabIndex        =   5
      Top             =   405
      Width           =   3120
      _cx             =   5503
      _cy             =   5054
      Appearance      =   2
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanUnit.ctx":00E0
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
      Editable        =   2
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
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ԤԼ���Ʒ�ʽ"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   110
      Width           =   1080
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H8000000A&
      Height          =   3615
      Left            =   1620
      Top             =   3810
      Width           =   2655
   End
End
Attribute VB_Name = "ClinicPlanUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobj���к�����λ As ������λ���Ƽ�
Private mobj������λ�� As ������λ���Ƽ�
Private mobj���к��� As ������Ϣ��
Private mblnNotClick As Boolean
Private mblnEdit As Boolean
Private mblnValiedCanSave As Boolean

Private Enum PageSub_Index
    Pg_������λ = 0
    Pg_ԤԼ��ʽ = 1
End Enum

'���Ա���:
Dim m_EditMode As gRegistPlanEditMode
Dim m_IsDataChanged As Boolean

'ȱʡ����ֵ:
Const m_def_EditMode = 0
Const m_def_IsDataChanged = False
'�¼�����:
Event DataIsChanged()


Public Function LoadData(ByVal obj������λ�� As ������λ���Ƽ�, ByVal obj���к��� As ������Ϣ��, _
    ByVal obj���к�����λ As ������λ���Ƽ�, Optional ByVal blnChanged As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���س��ﰲ��
    '���:
    '     obj������λ��-������λ������Ϣ
    '     obj���к�����λ - ���к�����λ���Ƽ� ,������ʾ�鿴
    '     obj���к��� - ���б�ѡ����
    '����:���سɹ�������true,���򷵻�false
    '����:���˺�
    '����:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj������λ�� = obj������λ��
    Set mobj���к��� = obj���к���
    Set mobj���к�����λ = obj���к�����λ

    If mobj������λ�� Is Nothing Then Set mobj������λ�� = New ������λ���Ƽ�
    If mobj���к��� Is Nothing Then Set mobj���к��� = New ������Ϣ��
    m_IsDataChanged = blnChanged
    LoadData = InitData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitFace()
    Dim i As Integer
    
    Err = 0: On Error GoTo Errhand
    With tbPage.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .BoldSelected = True
        .Layout = xtpTabLayoutAutoSize
        .StaticFrame = False
        .ClientFrame = xtpTabFrameBorder
        .Position = xtpTabPositionBottom
    End With
    
    For i = tbPageSub.LBound To tbPageSub.UBound
        With tbPageSub(i).PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .Layout = xtpTabLayoutAutoSize
            .StaticFrame = False
            .ClientFrame = xtpTabFrameBorder
        End With
    Next
    
    tbPage.Enabled = False '���ÿؼ�������������ؼ��İ�ť�ؼ����л����ÿؼ�ʱ���㲢û������ʧȥ
    tbPage.InsertItem Pg_������λ, "�Һź�����λ", picPageSub(0).Hwnd, 0
    tbPage.InsertItem Pg_ԤԼ��ʽ, "ԤԼ��ʽ", picPageSub(1).Hwnd, 0
    tbPage.Item(Pg_������λ).Selected = True
    tbPage.Enabled = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub UnitPageVisible(ByVal blnVisible As Boolean)
    '��������������λ
    Dim i As Integer
    Dim blnDo As Boolean
    
    Err = 0: On Error GoTo Errhand
    'List
    For i = 1 To vsUnit.Rows - 1
        vsUnit.RowHidden(i) = False
        If vsUnit.RowData(i) = 1 Then vsUnit.RowHidden(i) = blnVisible = False
    Next
    'TabPage
    tbPageSub(Pg_������λ).Visible = blnVisible
    If blnVisible = False Then
        tbPage.Enabled = False
        tbPage(Pg_ԤԼ��ʽ).Selected = True
        tbPage.Enabled = True
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetGridColVisible(ByVal bln��ʱ�� As Boolean, ByVal bln��ſ��� As Boolean)
    '���������пɼ�״̬
    Dim i As Integer, j As Integer
    
    Err = 0: On Error GoTo Errhand:
    vsfNotSelNum.ColHidden(-1) = False
    vsfNotSelNum.AllowSelection = False
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        vsfSelNum(i).ColHidden(-1) = False
        vsfSelNum(i).Editable = flexEDNone '�����༭
        vsfSelNum(i).FocusRect = flexFocusNone
        vsfSelNum(i).AllowSelection = False
    Next
    If bln��ʱ�� Then
        If bln��ſ��� Then
            '��ʱ����ſ���"����"�в��ɼ�
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("����")) = True
            vsfNotSelNum.AllowSelection = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("����")) = True
                vsfSelNum(i).AllowSelection = True
            Next
        Else
            '��ʱ�β���ſ���"���"�в��ɼ�
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("���")) = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).Editable = flexEDKbdMouse  '�����༭
                vsfSelNum(i).FocusRect = flexFocusLight
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("���")) = True
            Next
        End If
    Else
        If bln��ſ��� Then
            '����ʱ����ſ���ֻ��"���"�пɼ�
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("ʱ���")) = True
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("����")) = True
            vsfNotSelNum.AllowSelection = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("ʱ���")) = True
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("����")) = True
                vsfSelNum(i).AllowSelection = True
            Next
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearUnitGridData()
    '�����������
    Dim i As Integer
    
    With vsUnit
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                .TextMatrix(i, .ColIndex("��ֹԤԼ")) = 0
                .TextMatrix(i, .ColIndex("����")) = ""
                .Cell(flexcpBackColor, i, .ColIndex("����")) = vsUnit.BackColor
            End If
        Next
    End With
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2016-01-12 12:48:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj���� As ������Ϣ, obj���� As ������Ϣ��
    Dim objVsfGrid As VSFlexGrid, obj������λ As ������λ����
    Dim bln��ʱ�� As Boolean, bln��ſ��� As Boolean, bytԤԼ���� As Byte
    Dim blnFind As Boolean, i As Long, lngRow As Long, j As Long
    
    Err = 0: On Error GoTo Errhand:
    
    '============================================
    '�ȼ������к�����λ����ʼ������
    picFun.Tag = ""
    If mobj���к�����λ Is Nothing Then
        vsUnit.Clear 1: vsUnit.Rows = 1
    Else
        With vsUnit
            .Clear 1
            .Rows = 1
            .Rows = mobj���к�����λ.Count + 1
            lngRow = 1
            For Each obj������λ In mobj���к�����λ
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(obj������λ.���� = 1, "�Һź�����λ", "ԤԼ��ʽ")
                .TextMatrix(lngRow, .ColIndex("������λ")) = obj������λ.������λ����
                .RowData(lngRow) = obj������λ.���� '1-��������;2-ԤԼ��ʽ
                lngRow = lngRow + 1
            Next
            .ColAlignment(.ColIndex("��ֹԤԼ")) = flexAlignCenterCenter
            .ColDataType(.ColIndex("��ֹԤԼ")) = flexDTBoolean
            
            '���ݰ����ͷ���
            .OutlineBar = flexOutlineBarComplete
            .Subtotal flexSTClear
            .Subtotal flexSTNone, .ColIndex("����")
            .SubtotalPosition = flexSTAbove

            .Outline .ColIndex("������λ")
            .OutlineCol = .ColIndex("������λ")

            .MergeCompare = flexMCIncludeNulls
            .MergeCells = flexMergeRestrictRows
            For i = 0 To .Cols - 1
                .MergeCol(i) = True
            Next
            
            For i = 1 To .Rows - 1
                .MergeRow(i) = False
                If .IsSubtotal(i) Then
                    .IsCollapsed(i) = flexOutlineExpanded
                    .MergeRow(i) = True
                    .TextMatrix(i, .ColIndex("������λ")) = .TextMatrix(i, .ColIndex("����"))
                    .RowData(i) = IIf(.TextMatrix(i, .ColIndex("����")) = "�Һź�����λ", 1, 2)
                End If
            Next
        End With
    End If
    '����ҳ��
    Call InitUnitPage
    '�����������
    Call ClearUnitGridData
    
    vsfNotSelNum.Clear 1: vsfNotSelNum.Rows = 1
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        vsfSelNum(i).Clear 1: vsfSelNum(i).Rows = 1
    Next
    '============================================
    
    bln��ʱ�� = mobj���к���.�Ƿ��ʱ��
    bln��ſ��� = mobj���к���.�Ƿ���ſ���
    '0-��ֹԤԼ(��Һ�);1-����������ԤԼ(��Һ�);2-����������ԤԼ(��Һ�);3-����ſ���ԤԼ(��Һ�);4-��������
    bytԤԼ���� = mobj������λ��.ԤԼ���Ʒ�ʽ
    
    '0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
    Call UnitPageVisible(mobj���к���.ԤԼ���� <> 2)
    Call SetGridColVisible(bln��ʱ��, bln��ſ���)
    mblnEdit = bln��ʱ�� And Not bln��ſ���
    
    If bln��ʱ�� = False And bln��ſ��� = False And bytԤԼ���� = 3 Then bytԤԼ���� = 0
    mblnNotClick = True
    optBespeakMode(IIf(bytԤԼ���� = 0 Or bytԤԼ���� = 4, 0, bytԤԼ���� - 1)).Value = True
    chkOnlyOneUse.Value = IIf(mobj������λ��.�Ƿ��ռ, vbChecked, vbUnchecked)
    mblnNotClick = False
    
    '��ǰ���ſ���ԤԼ(��Һ�)�Ƿ�ɼ�
    optBespeakMode(2).Tag = IIf(bln��ʱ�� Or bln��ſ���, "", "1")
    picFun.Tag = IIf(bln��ſ���, "", "1")
    
    If bytԤԼ���� <> 3 Then
        For Each obj������λ In mobj������λ��
            With vsUnit
                .Redraw = flexRDNone
                For i = 1 To .Rows - 1
                    If .RowData(i) = obj������λ.���� And .TextMatrix(i, .ColIndex("������λ")) = obj������λ.������λ���� Then
                        Select Case obj������λ.ԤԼ���Ʒ�ʽ
                        Case 0 '��ֹԤԼ
                            .TextMatrix(i, .ColIndex("��ֹԤԼ")) = 1
                            .Cell(flexcpBackColor, i, .ColIndex("����")) = vbButtonFace
                        Case 1, 2
                            If Not obj������λ.������Ϣ�� Is Nothing Then
                                For Each obj���� In obj������λ.������Ϣ��
                                    '���:���Ʒ�ʽ=0,1,2,4ʱ����Ϊ0;����洢������Ż��ʱ�ε����
                                    '����:���Ʒ�ʽ=0,4ʱ����Ϊ0;���Ʒ�ʽ=1ʱ����ű���,��20,����20%,���Ʒ�ʽ=2ʱ���洢������Լ���������磺10��ʾֻ��ԤԼ10����;���Ʒ�ʽ=3ʱ���洢��Լ������������ŵģ�һ��Ϊ1,����������ҷ�ʱ�εģ��洢��Լ����
                                    .TextMatrix(i, .ColIndex("����")) = FormatEx(obj����.����, 2, False)
                                    Exit For
                                Next
                            End If
                        End Select
                    End If
                Next
                .Redraw = flexRDBuffered
            End With
        Next
    End If

    '�������������Ϣ
    If bln��ʱ�� Or bln��ſ��� Then
        With vsfNotSelNum
            .Redraw = flexRDNone
            For Each obj���� In mobj���к���
                If obj����.�Ƿ�ԤԼ And obj����.���� > 0 Then
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, .ColIndex("���")) = obj����.���
                    .TextMatrix(lngRow, .ColIndex("ʱ���")) = Format(obj����.��ʼʱ��, "hh:mm") & "-" & Format(obj����.��ֹʱ��, "hh:mm")
                    .Cell(flexcpData, lngRow, .ColIndex("ʱ���")) = obj����.��ʼʱ�� & "-" & obj����.��ֹʱ��
                    .TextMatrix(lngRow, .ColIndex("����")) = obj����.����
                    .Cell(flexcpData, lngRow, .ColIndex("����")) = obj����.����
                End If
            Next
            .Redraw = flexRDBuffered
        End With
        
        If bln��ʱ�� And bln��ſ��� = False Then
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                With vsfSelNum(i)
                    .Redraw = flexRDNone
                    For Each obj���� In mobj���к���
                        If obj����.�Ƿ�ԤԼ And obj����.���� > 0 Then
                            .Rows = .Rows + 1
                            lngRow = .Rows - 1
                            .TextMatrix(lngRow, .ColIndex("���")) = obj����.���
                            .TextMatrix(lngRow, .ColIndex("ʱ���")) = Format(obj����.��ʼʱ��, "hh:mm") & "-" & Format(obj����.��ֹʱ��, "hh:mm")
                            .Cell(flexcpData, lngRow, .ColIndex("ʱ���")) = obj����.��ʼʱ�� & "-" & obj����.��ֹʱ��
                            .TextMatrix(lngRow, .ColIndex("����")) = 0
                        End If
                    Next
                    .Redraw = flexRDBuffered
                End With
            Next
        End If
        If vsfNotSelNum.Rows > 1 And vsfNotSelNum.Row < 1 Then vsfNotSelNum.Row = 1

        '���غ�����λ��ѡ�������Ϣ
        For Each obj������λ In mobj������λ��
            Set objVsfGrid = GetUnitVsfGrid(obj������λ.����, obj������λ.������λ����)
            If Not objVsfGrid Is Nothing Then
                Select Case obj������λ.ԤԼ���Ʒ�ʽ
                Case 0 '��ֹԤԼ
                    mblnNotClick = True
                    chkForbidBespeak(objVsfGrid.index).Value = vbChecked
                    mblnNotClick = False
                    objVsfGrid.Editable = flexEDNone
                Case 3
                    If Not obj������λ.������Ϣ�� Is Nothing Then
                        vsfNotSelNum.Redraw = flexRDNone
                        objVsfGrid.Redraw = flexRDNone
                        For Each obj���� In obj������λ.������Ϣ��
                            '���:���Ʒ�ʽ=0,1,2,4ʱ����Ϊ0;����洢������Ż��ʱ�ε����
                            '����:���Ʒ�ʽ=0,4ʱ����Ϊ0;���Ʒ�ʽ=1ʱ����ű���,��20,����20%,���Ʒ�ʽ=2ʱ���洢������Լ���������磺10��ʾֻ��ԤԼ10����;���Ʒ�ʽ=3ʱ���洢��Լ������������ŵģ�һ��Ϊ1,����������ҷ�ʱ�εģ��洢��Լ����
                            If bln��ʱ�� And bln��ſ��� = False Then
                                RemoveItem vsfNotSelNum, objVsfGrid, obj����.���, True, obj����.����
                            Else
                                RemoveItem vsfNotSelNum, objVsfGrid, obj����.���
                            End If
                        Next
                        vsfNotSelNum.Redraw = flexRDBuffered
                        objVsfGrid.Redraw = flexRDBuffered
                    End If
                    objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And mblnEdit, flexEDKbdMouse, flexEDNone)
                End Select
            End If
        Next
    End If
    
Handler:
    Call SetUnitVisible
    Call SetButtonEnable(SelectedVsfGridIndex)
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetUnitVsfGrid(ByVal byt���� As Byte, ByVal str���� As String) As VSFlexGrid
    '�������ͺ����ƻ�ȡ��Ӧ��VSFlexGrid�ؼ�
    '��Σ�
    '   byt���� 1-��������;2-ԤԼ��ʽ
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    If str���� = "" Then Exit Function
     
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        If Val(vsfSelNum(i).Tag) = byt���� And picUnit(i).Tag = str���� Then
            Set GetUnitVsfGrid = vsfSelNum(i)
            Exit Function
        End If
    Next
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chkForbidBespeak_Click(index As Integer)
    Dim objVsfGrid As VSFlexGrid, i As Long
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    Set objVsfGrid = vsfSelNum(index)
    
    objVsfGrid.Redraw = flexRDNone
    vsfNotSelNum.Redraw = flexRDNone
    If Not mobj���к��� Is Nothing Then
        If mobj���к���.�Ƿ��ʱ�� And mobj���к���.�Ƿ���ſ��� = False Then
            For i = 1 To objVsfGrid.Rows - 1
                RemoveItem vsfNotSelNum, objVsfGrid, Val(objVsfGrid.TextMatrix(i, objVsfGrid.ColIndex("���"))), True, 0
            Next
            objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And mblnEdit, flexEDKbdMouse, flexEDNone)
            Exit Sub
        End If
    End If
    
    For i = 1 To objVsfGrid.Rows - 1
        If i > objVsfGrid.Rows - 1 Then Exit For
        RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(i, objVsfGrid.ColIndex("���")))
        i = i - 1
    Next
    objVsfGrid.Redraw = flexRDBuffered
    vsfNotSelNum.Redraw = flexRDBuffered
    
    Call SetButtonEnable(objVsfGrid.index)
    objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And mblnEdit, flexEDKbdMouse, flexEDNone)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetButtonEnable(ByVal index As Integer)
    If m_EditMode <> ED_RegistPlan_Edit _
        Or index < vsfSelNum.LBound Or index > vsfSelNum.UBound Then
        cmdFun(0).Enabled = False
        cmdFun(1).Enabled = False
        cmdFun(2).Enabled = False
        cmdFun(3).Enabled = False
        Exit Sub
    End If
    
    cmdFun(0).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfNotSelNum.Row > 0
    cmdFun(1).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfNotSelNum.Rows > 1
    cmdFun(2).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfSelNum(index).Row > 0
    cmdFun(3).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfSelNum(index).Rows > 1
End Sub

Private Sub RemoveItem(ByVal objVsfGridFrom As VSFlexGrid, ByVal objVsfGridTo As VSFlexGrid, ByVal lngSN As Long, _
    Optional ByVal blnChangeNum As Boolean, Optional lngNum As Long)
    '�ƶ���Ŀ���������
    '������
    '   lngSN ���
    '   blnChangeNum ���ı�����,��ʱ�Σ�����ſ���ʱ
    '   lngNum �ı������
    Dim blnFind As Boolean, i As Integer, j As Integer
    Dim lngRow As Long
    Dim intLow As Integer, intHigh As Integer, intMid As Integer
    
    On Error GoTo Errhand
    If objVsfGridFrom.Rows > 1 Then
        If Val(objVsfGridFrom.TextMatrix(1, objVsfGridFrom.ColIndex("���"))) = lngSN Then
            lngRow = 1
        ElseIf Val(objVsfGridFrom.TextMatrix(objVsfGridFrom.Rows - 1, objVsfGridFrom.ColIndex("���"))) = lngSN Then
            lngRow = objVsfGridFrom.Rows - 1
        End If
    End If
    '���ַ�����
    If lngRow = 0 Then
        intLow = 1
        intHigh = objVsfGridFrom.Rows - 1
        Do While intLow <= intHigh
            intMid = (intLow + intHigh) \ 2
            If Val(objVsfGridFrom.TextMatrix(intMid, objVsfGridFrom.ColIndex("���"))) < lngSN Then '�ں���
                intLow = intMid + 1
            ElseIf Val(objVsfGridFrom.TextMatrix(intMid, objVsfGridFrom.ColIndex("���"))) > lngSN Then '��ǰ��
                intHigh = intMid - 1
            Else
                lngRow = intMid: Exit Do
            End If
        Loop
    End If
    If lngRow = 0 Then Exit Sub
    
    If blnChangeNum Then
        For i = 1 To objVsfGridTo.Rows - 1
            If Val(objVsfGridTo.TextMatrix(i, objVsfGridTo.ColIndex("���"))) = lngSN Then
                objVsfGridTo.TextMatrix(lngRow, objVsfGridTo.ColIndex("����")) = lngNum
                Exit For
            End If
        Next
        '����ʣ������
        lngNum = Val(objVsfGridFrom.Cell(flexcpData, lngRow, objVsfGridFrom.ColIndex("����")))
        For i = vsfSelNum.LBound To vsfSelNum.UBound
            For j = 1 To vsfSelNum(i).Rows - 1
                If Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("���"))) = lngSN Then
                    lngNum = lngNum - Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("����")))
                    Exit For
                End If
            Next
        Next
        objVsfGridFrom.TextMatrix(lngRow, objVsfGridFrom.ColIndex("����")) = lngNum
    Else
        '��˳�����
        blnFind = False
        If objVsfGridTo.Rows <= 1 Then
            With objVsfGridFrom
                objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("���")) & vbTab & .TextMatrix(lngRow, .ColIndex("ʱ���")) & _
                    vbTab & .TextMatrix(lngRow, .ColIndex("����"))
                objVsfGridTo.Cell(flexcpData, objVsfGridTo.Rows - 1, objVsfGridTo.ColIndex("ʱ���")) = .Cell(flexcpData, lngRow, .ColIndex("ʱ���"))
            End With
            blnFind = True
        Else
            If Val(objVsfGridTo.TextMatrix(1, objVsfGridTo.ColIndex("���"))) >= lngSN Then
                With objVsfGridFrom
                    objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("���")) & vbTab & .TextMatrix(lngRow, .ColIndex("ʱ���")) & _
                        vbTab & .TextMatrix(lngRow, .ColIndex("����")), 1
                    objVsfGridTo.Cell(flexcpData, 1, objVsfGridTo.ColIndex("ʱ���")) = .Cell(flexcpData, lngRow, .ColIndex("ʱ���"))
                End With
                blnFind = True
            ElseIf Val(objVsfGridTo.TextMatrix(objVsfGridTo.Rows - 1, objVsfGridTo.ColIndex("���"))) <= lngSN Then
                With objVsfGridFrom
                    objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("���")) & vbTab & .TextMatrix(lngRow, .ColIndex("ʱ���")) & _
                        vbTab & .TextMatrix(lngRow, .ColIndex("����"))
                    objVsfGridTo.Cell(flexcpData, objVsfGridTo.Rows - 1, objVsfGridTo.ColIndex("ʱ���")) = .Cell(flexcpData, lngRow, .ColIndex("ʱ���"))
                End With
                blnFind = True
            End If
        End If
        
        '���ַ�����
        If blnFind = False Then
            intLow = 1
            intHigh = objVsfGridTo.Rows - 1
            Do While intLow <= intHigh
                intMid = (intLow + intHigh) \ 2
                If Val(objVsfGridTo.TextMatrix(intMid - 1, objVsfGridTo.ColIndex("���"))) < lngSN _
                    And Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("���"))) > lngSN Then   '�ҵ�λ���ˣ��ҿ϶����ҵ�
                    With objVsfGridFrom
                        objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("���")) & vbTab & .TextMatrix(lngRow, .ColIndex("ʱ���")) & _
                            vbTab & .TextMatrix(lngRow, .ColIndex("����")), intMid
                        objVsfGridTo.Cell(flexcpData, intMid, objVsfGridTo.ColIndex("ʱ���")) = .Cell(flexcpData, lngRow, .ColIndex("ʱ���"))
                    End With
                    Exit Do
                ElseIf Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("���"))) < lngSN _
                    And Val(objVsfGridTo.TextMatrix(intMid + 1, objVsfGridTo.ColIndex("���"))) > lngSN Then '�ҵ�λ���ˣ��ҿ϶����ҵ�
                    With objVsfGridFrom
                        objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("���")) & vbTab & .TextMatrix(lngRow, .ColIndex("ʱ���")) & _
                            vbTab & .TextMatrix(lngRow, .ColIndex("����")), intMid + 1
                        objVsfGridTo.Cell(flexcpData, intMid + 1, objVsfGridTo.ColIndex("ʱ���")) = .Cell(flexcpData, lngRow, .ColIndex("ʱ���"))
                    End With
                    Exit Do
                End If
                
                If Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("���"))) < lngSN Then '�ں���
                    intLow = intMid + 1
                ElseIf Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("���"))) > lngSN Then '��ǰ��
                    intHigh = intMid - 1
                End If
            Loop
        End If
        objVsfGridFrom.RemoveItem lngRow
        
        If objVsfGridFrom.Rows > 1 And objVsfGridFrom.Row < 1 Then objVsfGridFrom.Row = 1
        If objVsfGridTo.Rows > 1 And objVsfGridTo.Row < 1 Then objVsfGridTo.Row = 1
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkForbidBespeak_GotFocus(index As Integer)
    chkForbidBespeak(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub chkForbidBespeak_LostFocus(index As Integer)
     chkForbidBespeak(index).BackColor = Me.BackColor
End Sub


Private Sub chkForbidBespeak_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkOnlyOneUse_Click()
    Dim i As Integer
    
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    '�����������
    Call ClearUnitGridData
End Sub

Private Sub chkOnlyOneUse_GotFocus()
    chkOnlyOneUse.BackColor = GCTRL_SELBACK_COLOR
End Sub
Private Sub chkOnlyOneUse_LostFocus()
     chkOnlyOneUse.BackColor = Me.BackColor
End Sub
Private Sub chkOnlyOneUse_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Property Get SelectedVsfGridIndex() As Integer
    '��ȡ��ǰѡ�����������
    Dim byt���� As Byte, str���� As String
    Dim objVsfGrid As VSFlexGrid
    
    On Error GoTo ErrHandler
    SelectedVsfGridIndex = -1
    
    With tbPageSub(tbPage.Selected.index)
        If .Selected Is Nothing Then Exit Sub
        str���� = .Selected.Caption
        byt���� = Val(.Selected.Tag)
    End With
     
    Set objVsfGrid = GetUnitVsfGrid(byt����, str����)
    If Not objVsfGrid Is Nothing Then
        SelectedVsfGridIndex = objVsfGrid.index
    End If
    Exit Property
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Property

Private Sub cmdFun_Click(index As Integer)
    Dim objVsfGrid As VSFlexGrid, intSelectedGridIndex As Integer
    Dim blnFind As Boolean, i As Integer
    Dim intStartRow As Integer, intEndRow As Integer
    Dim byt���� As Byte, str���� As String
    
    On Error GoTo Errhand
    intSelectedGridIndex = SelectedVsfGridIndex
    If intSelectedGridIndex < vsfSelNum.LBound Or intSelectedGridIndex > vsfSelNum.UBound Then Exit Sub
    Set objVsfGrid = vsfSelNum(intSelectedGridIndex)
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    vsfNotSelNum.Redraw = flexRDNone
    objVsfGrid.Redraw = flexRDNone
    Select Case index
    Case 0 'ѡ��
        '��������
        intStartRow = vsfNotSelNum.RowSel: intEndRow = vsfNotSelNum.Row
        If vsfNotSelNum.Row < vsfNotSelNum.RowSel Then
            intStartRow = vsfNotSelNum.Row: intEndRow = vsfNotSelNum.RowSel
        End If
        Do While True
            If intStartRow > intEndRow Then Exit Do
            RemoveItem vsfNotSelNum, objVsfGrid, Val(vsfNotSelNum.TextMatrix(intStartRow, vsfNotSelNum.ColIndex("���")))
            intEndRow = intEndRow - 1
        Loop
        If intStartRow > 0 And intStartRow < vsfNotSelNum.Rows Then vsfNotSelNum.Select intStartRow, 0
    Case 1 'ȫѡ��
        For i = 1 To vsfNotSelNum.Rows - 1
            If i > vsfNotSelNum.Rows - 1 Then Exit For
            RemoveItem vsfNotSelNum, objVsfGrid, Val(vsfNotSelNum.TextMatrix(i, vsfNotSelNum.ColIndex("���")))
            i = i - 1
        Next
    Case 2 '�Ƴ�
        '��������
        intStartRow = objVsfGrid.RowSel: intEndRow = objVsfGrid.Row
        If objVsfGrid.Row < objVsfGrid.RowSel Then
            intStartRow = objVsfGrid.Row: intEndRow = objVsfGrid.RowSel
        End If
        Do While True
            If intStartRow > intEndRow Then Exit Do
            RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(intStartRow, objVsfGrid.ColIndex("���")))
            intEndRow = intEndRow - 1
        Loop
        If intStartRow > 0 And intStartRow < objVsfGrid.Rows Then objVsfGrid.Select intStartRow, 0
    Case 3 'ȫ�Ƴ�
        For i = 1 To objVsfGrid.Rows - 1
            If i > objVsfGrid.Rows - 1 Then Exit For
            RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(i, objVsfGrid.ColIndex("���")))
            i = i - 1
        Next
    End Select
    vsfNotSelNum.Redraw = flexRDBuffered
    objVsfGrid.Redraw = flexRDBuffered
    
    Call SetButtonEnable(intSelectedGridIndex)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdFun_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optBespeakMode_Click(index As Integer)
    Dim i As Integer, j As Integer
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    '�����������
    Call ClearUnitGridData
    If Not mobj���к��� Is Nothing Then
        For i = vsfSelNum.LBound To vsfSelNum.UBound
            chkForbidBespeak(i).Value = vbUnchecked
            For j = 1 To vsfSelNum(i).Rows - 1
                If mobj���к���.�Ƿ���ſ��� Then
                    If j > vsfSelNum(i).Rows - 1 Then Exit For
                    RemoveItem vsfSelNum(i), vsfNotSelNum, Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("���")))
                    j = j - 1
                Else
                    RemoveItem vsfNotSelNum, vsfSelNum(i), Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("���"))), True, 0
                End If
            Next
        Next
    End If
    Call SetUnitVisible
    Call SetButtonEnable(SelectedVsfGridIndex)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optBespeakMode_GotFocus(index As Integer)
    optBespeakMode(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub optBespeakMode_LostFocus(index As Integer)
     optBespeakMode(index).BackColor = Me.BackColor
End Sub

Private Sub optBespeakMode_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub picFun_Resize()
    Err = 0: On Error Resume Next
    cmdFun(0).Top = (picFun.ScaleHeight - (cmdFun(0).Height + 100) * 4) / 2
    cmdFun(1).Top = cmdFun(0).Top + cmdFun(0).Height + 100
    cmdFun(2).Top = cmdFun(1).Top + cmdFun(1).Height + 100
    cmdFun(3).Top = cmdFun(2).Top + cmdFun(2).Height + 100
End Sub

Private Sub picPageSub_Resize(index As Integer)
    On Error Resume Next
    With tbPageSub(index)
        .Left = 0
        .Top = 0
        .Height = picPageSub(index).ScaleHeight
        .Width = picPageSub(index).ScaleWidth
    End With
End Sub

Private Sub PicUnit_Resize(index As Integer)
    Err = 0: On Error Resume Next
    With picUnit(index)
        chkForbidBespeak(index).Left = .ScaleLeft + 30
        chkForbidBespeak(index).Top = .ScaleTop + 30
        
        vsfSelNum(index).Left = .ScaleLeft
        vsfSelNum(index).Width = .ScaleWidth
        vsfSelNum(index).Top = chkForbidBespeak(index).Top + chkForbidBespeak(index).Height
        vsfSelNum(index).Height = .ScaleHeight - vsfSelNum(index).Top
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick Then Exit Sub
    Call SelectedPageChanged(tbPage, Item.index)
End Sub

Private Sub SelectedPageChanged(tbPage As TabControl, ByVal index As Integer)
    '�л�ҳ�棬�����ڱ༭�����ݽ��м��
    '���:
    '   Index ��ǰѡ��Page.Index
    Dim intSelectedGridIndex As Integer
    
    On Error GoTo ErrHandler
    If Val(tbPage.Tag) < tbPage.ItemCount Then
        mblnNotClick = True
        tbPage.Enabled = False
        tbPage.Item(Val(tbPage.Tag)).Selected = True
        tbPage.Enabled = True
        
        intSelectedGridIndex = SelectedVsfGridIndex
        If intSelectedGridIndex >= vsfSelNum.LBound And intSelectedGridIndex <= vsfSelNum.UBound Then
            mblnValiedCanSave = True
            vsfSelNum(intSelectedGridIndex).FinishEditing False    '�����ڱ༭�����ݽ��м��
            If mblnValiedCanSave = False Then mblnNotClick = False: Exit Sub
        End If
        
        tbPage.Enabled = False
        tbPage.Item(index).Selected = True
        tbPage.Enabled = True
        mblnNotClick = False
    End If
    
    tbPage.Tag = index
    Call SetButtonEnable(SelectedVsfGridIndex)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnNotClick = False
End Sub

Private Sub tbPageSub_SelectedChanged(index As Integer, ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick Then Exit Sub
    Call SelectedPageChanged(tbPageSub(index), Item.index)
End Sub

Private Sub UserControl_Initialize()
    Call InitFace
    Call SetUnitVisible
End Sub

Private Sub UserControl_Resize()
    Dim sngSplit As Single
    
    Err = 0: On Error Resume Next
    With vsUnit
        .Left = 0
        .Top = optBespeakMode(0).Top + optBespeakMode(0).Height + 50
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - .Top
    End With
    If m_EditMode = ED_RegistPlan_Edit Then
        With vsfNotSelNum
            .Left = vsUnit.Left
            .Top = vsUnit.Top
            .Height = vsUnit.Height
        End With
        With picFun
            .Left = vsfNotSelNum.Left + vsfNotSelNum.Width + 20
            .Top = vsUnit.Top
            .Height = vsUnit.Height
        End With
        With tbPage
            .Left = IIf(picFun.Tag = "", picFun.Left + picFun.Width, picFun.Left) + 20
            .Top = vsUnit.Top
            .Height = vsUnit.Height
            .Width = ScaleWidth - .Left
        End With
    
        '��һ���߿�
        sngSplit = 30
        With tbPage
            .Left = .Left + sngSplit
            .Top = .Top + sngSplit
            .Height = .Height - 2 * sngSplit
            .Width = .Width - 2 * sngSplit
        End With
        With shpBack
            .Left = tbPage.Left - sngSplit
            .Top = tbPage.Top - sngSplit
            .Height = tbPage.Height + 2 * sngSplit
            .Width = tbPage.Width + 2 * sngSplit
        End With
    Else
        With tbPage
            .Left = 0
            .Top = vsUnit.Top
            .Height = vsUnit.Height
            .Width = ScaleWidth - .Left
        End With
    End If
End Sub

Private Sub InitUnitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2016-01-11 14:23:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objParentPage As TabControl
    Dim objUnit As ������λ����, lngRow As Long
    Dim intPageCount As Integer, p As Integer
    Dim intPage As Integer, intParentPage As Integer
    
    Err = 0: On Error GoTo Errhand:
    intParentPage = tbPage.Selected.index
    intPage = -1
    If Not tbPageSub(intParentPage).Selected Is Nothing Then
        intPage = tbPageSub(intParentPage).Selected.index
    End If
    
    For p = tbPageSub.LBound To tbPageSub.UBound
        tbPageSub(p).RemoveAll
    Next
    intPageCount = picUnit.Count
    
    If Not mobj���к�����λ Is Nothing Then
        For Each objUnit In mobj���к�����λ
            If lngRow >= intPageCount Then
                Load chkForbidBespeak(lngRow): chkForbidBespeak(lngRow).Visible = True
                Load vsfSelNum(lngRow): vsfSelNum(lngRow).Visible = True
                Load picUnit(lngRow): picUnit(lngRow).Visible = True
                Set chkForbidBespeak(lngRow).Container = picUnit(lngRow)
                Set vsfSelNum(lngRow).Container = picUnit(lngRow)
                picUnit(lngRow).TabStop = False
            End If
            
            picUnit(lngRow).Visible = True
            Set objParentPage = IIf(objUnit.���� = 1, tbPageSub(Pg_������λ), tbPageSub(Pg_ԤԼ��ʽ))
            Set objItem = objParentPage.InsertItem(objParentPage.ItemCount, objUnit.������λ����, picUnit(lngRow).Hwnd, 0)
            objItem.Tag = objUnit.���� '1-��������;2-ԤԼ��ʽ
            vsfSelNum(lngRow).Tag = objUnit.���� '1-��������;2-ԤԼ��ʽ
            picUnit(lngRow).Tag = objUnit.������λ����
            lngRow = lngRow + 1
        Next
    End If
    
    If lngRow <= intPageCount Then
        For i = IIf(lngRow = 0, 1, lngRow) To picUnit.UBound
            Unload chkForbidBespeak(i)
            Unload vsfSelNum(i)
            Unload picUnit(i)
        Next
    End If
    
    '��ʾ/����"������λ"ҳǩ
    If tbPageSub(Pg_������λ).ItemCount = 0 And intParentPage = Pg_������λ Then
        tbPage.Enabled = False: mblnNotClick = True
        tbPage.Item(Pg_ԤԼ��ʽ).Selected = True
        tbPage.Enabled = True: mblnNotClick = False
        intParentPage = Pg_ԤԼ��ʽ
        
        tbPage(Pg_������λ).Visible = False
    Else
        tbPage(Pg_������λ).Visible = True
    End If
    
    '�ָ�ѡ��ҳǩ
    If intPage > tbPageSub(intParentPage).ItemCount - 1 Then
        intPage = tbPageSub(intParentPage).ItemCount - 1
    End If
    If intPage = -1 Then intPage = 0
    
    If tbPageSub(intParentPage).ItemCount > 0 Then
        '�ֶ�����SelectedChanged�¼�
        Call tbPageSub_SelectedChanged(intParentPage, tbPageSub(intParentPage).Item(intPage))
        tbPageSub(intParentPage).Enabled = False
        tbPageSub(intParentPage).Item(intPage).Selected = True
        tbPageSub(intParentPage).Enabled = True
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnNotClick = False
End Sub

Private Sub InitUnitGrid()
    '��ʼ��������λ�������
    Dim i As Integer
    
    Err = 0: On Error GoTo Errhand:
    With vsfNotSelNum
        .Clear 1
        .Rows = 1
        .HighLight = flexHighlightAlways
        .ColHidden(-1) = False
    End With
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        With vsfSelNum(i)
            .Clear 1
            .Rows = 1
            .Editable = flexEDNone
            .ColHidden(-1) = False
            .HighLight = flexHighlightAlways
            .FocusRect = flexFocusNone
        End With
        mblnNotClick = True
        chkForbidBespeak(i).Value = vbUnchecked
        mblnNotClick = False
    Next
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetUnitVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݺ�����λ��ԤԼ���Ʒ�ʽ�����ö�Ӧ�Ŀؼ���ʾ
    '����:���˺�
    '����:2016-01-12 11:23:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    vsfNotSelNum.Visible = m_EditMode = ED_RegistPlan_Edit And optBespeakMode(2).Value
    picFun.Visible = m_EditMode = ED_RegistPlan_Edit And optBespeakMode(2).Value And Val(picFun.Tag) = 0
    tbPage.Visible = optBespeakMode(2).Value
    vsUnit.Visible = Not optBespeakMode(2).Value
    chkOnlyOneUse.Visible = Not optBespeakMode(2).Value
    optBespeakMode(2).Visible = Val(optBespeakMode(2).Tag) = 0
    If Val(optBespeakMode(2).Tag) = 0 Then
        chkOnlyOneUse.Left = optBespeakMode(2).Left + optBespeakMode(2).Width + 50
    Else
        chkOnlyOneUse.Left = optBespeakMode(2).Left
    End If
    vsUnit.TextMatrix(0, vsUnit.ColIndex("����")) = IIf(optBespeakMode(0).Value, "����(%)", "��Լ��")
    Call UserControl_Resize
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Get������λ���Ƽ�() As ������λ���Ƽ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������Ϣ��Ϣ����
    '����:������Ϣ��
    '����:���˺�
    '����:2016-01-13 12:34:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, intCol As Integer, index As Integer
    Dim objUnits As New ������λ���Ƽ�, objUnit As ������λ����
    Dim lngSum As Double, varTemp As Variant
    Dim strUnitName As String
    Dim objNums As ������Ϣ��, objNum As ������Ϣ
    Dim bln��ֹԤԼ As Boolean
    
    Err = 0: On Error GoTo Errhand:
    '����δ�ı䣬ֱ�ӷ���ԭ���ϵĸ���
    If m_IsDataChanged = False Then
        If mobj������λ��.Count = 0 And mobj���к�����λ.Count > 0 Then
            '��һ�μ��أ�û�иı䣬Ӧ����ȫ��������
            
        Else
            Set Get������λ���Ƽ� = mobj������λ��.Clone
            Exit Function
        End If
    End If
    
    '�����Ѹı䣬���¹��켯�϶���
    With objUnits
        .ԤԼ���Ʒ�ʽ = GetSelectedIndex(optBespeakMode) + 1
        .�Ƿ��ռ = chkOnlyOneUse.Value = vbChecked
        .�Ƿ��޸� = True
    End With
    
    If optBespeakMode(0).Value Or optBespeakMode(1).Value Then
        '���������ƻ���������
        With vsUnit
            For lngRow = 1 To .Rows - 1
                If .IsSubtotal(lngRow) = False Then
                    Set objUnit = New ������λ����
                    objUnit.������λ���� = .TextMatrix(lngRow, .ColIndex("������λ"))
                    objUnit.���� = .RowData(lngRow)
                    
                    If .RowHidden(lngRow) Then '���صľ��ǽ�ֹԤԼ
                        bln��ֹԤԼ = True
                        lngSum = 0
                    Else
                        bln��ֹԤԼ = Abs(Val(.TextMatrix(lngRow, .ColIndex("��ֹԤԼ")))) = 1
                        lngSum = Val(.TextMatrix(lngRow, .ColIndex("����")))
                    End If
                    '0-��ֹԤԼ;1-����������ԤԼ;2-����������ԤԼ;3-����ſ���ԤԼ;4-��������
                    objUnit.ԤԼ���Ʒ�ʽ = IIf(bln��ֹԤԼ, 0, _
                                            IIf(lngSum = 0, 4, IIf(optBespeakMode(0).Value, 1, 2)))
                    Set objNums = New ������Ϣ��
                    If lngSum > 0 Or bln��ֹԤԼ Then
                        Set objNum = New ������Ϣ
                        objNum.��� = 0
                        objNum.���� = lngSum
                        objNums.AddItem objNum
                    End If
                    Set objUnit.������Ϣ�� = objNums
                    objUnits.AddItem objUnit, "K" & objUnit.���� & "_" & objUnit.������λ����
                End If
            Next
        End With
    Else
        '0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
        For index = vsfSelNum.LBound To vsfSelNum.UBound
            If GetLocaleUnit(index, objUnit) Then
                objUnits.AddItem objUnit, "K" & objUnit.���� & "_" & objUnit.������λ����
            End If
        Next
    End If
    Set Get������λ���Ƽ� = objUnits
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetLocaleUnit(ByVal index As Integer, ByRef objUnit As ������λ����) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ĺ�������Ϣ
    '���:index-ָ����ҳ
    '����:objUnit-������λ��Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2016-01-13 18:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objNums As ������Ϣ��, objNum As ������Ϣ
    Dim varTemp As Variant, lngCount As Long
    
    Set objUnit = New ������λ����
    Err = 0: On Error GoTo Errhand:
    objUnit.������λ���� = picUnit(index).Tag
    objUnit.���� = Val(vsfSelNum(index).Tag)
    '0-��ֹԤԼ;1-����������ԤԼ;2-����������ԤԼ;3-����ſ���ԤԼ;4-��������
    If chkForbidBespeak(index).Value = vbChecked _
        Or mobj���к���.ԤԼ���� = 2 And objUnit.���� = 1 Then
        '����ֹ����������λ
        objUnit.ԤԼ���Ʒ�ʽ = 0
    Else
        objUnit.ԤԼ���Ʒ�ʽ = 3
    End If

    Set objNums = New ������Ϣ��
    If objUnit.ԤԼ���Ʒ�ʽ = 3 Then
        With vsfSelNum(index)
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("����"))) <> 0 Then
                    lngCount = lngCount + Val(.TextMatrix(i, .ColIndex("����")))
                    
                    Set objNum = New ������Ϣ
                    objNum.��� = Val(.TextMatrix(i, .ColIndex("���")))
                    If .TextMatrix(i, .ColIndex("ʱ���")) <> "" Then
                        varTemp = Split(.Cell(flexcpData, i, .ColIndex("ʱ���")), "-")
                        objNum.��ʼʱ�� = varTemp(0)
                        objNum.��ֹʱ�� = varTemp(1)
                    End If
                    objNum.���� = Val(.TextMatrix(i, .ColIndex("����")))
                    objNums.AddItem objNum
                End If
            Next
        End With
        'һ����Ŷ�û����������,���ʾ������
        If lngCount = 0 Then objUnit.ԤԼ���Ʒ�ʽ = 4
    End If
    If objUnit.ԤԼ���Ʒ�ʽ = 0 Or objUnit.ԤԼ���Ʒ�ʽ = 4 Then
        '��ֹԤԼ������ʱ����һ����¼���Ա㱣��
        Set objNum = New ������Ϣ
        objNum.��� = 0
        objNum.���� = 0
        objNums.AddItem objNum
    End If
    Set objUnit.������Ϣ�� = objNums
    GetLocaleUnit = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Property Get Get������λ������Ϣ��() As ������λ���Ƽ�
   Set Get������λ������Ϣ�� = Get������λ���Ƽ�
End Property

Public Function IsValied(Optional ByVal blnChanged As Boolean) As Boolean
    '�������
    '����һ���Ƿ�ı䣬���ı��򱾲�ҲҪ���
    Dim lngSum As Double, lng��Լ�� As Long, lngSN As Long
    Dim i As Long, j As Integer, k As Long
    Dim intSelectedGridIndex As Integer
    
    Err = 0: On Error GoTo ErrHandler
    '����δ�ı䲻���
    If m_IsDataChanged = False And blnChanged = False Then IsValied = True: Exit Function
    
    mblnValiedCanSave = True
    vsUnit.FinishEditing False
    If mblnValiedCanSave = False Then Exit Function
    
    mblnValiedCanSave = True
    intSelectedGridIndex = SelectedVsfGridIndex
    If intSelectedGridIndex >= vsfSelNum.LBound And intSelectedGridIndex <= vsfSelNum.UBound Then
        vsfSelNum(intSelectedGridIndex).FinishEditing False
    End If
    If mblnValiedCanSave = False Then Exit Function

    If optBespeakMode(0).Value Then '������
        If chkOnlyOneUse.Value = vbChecked Then
            For i = 1 To vsUnit.Rows - 1
                If vsUnit.IsSubtotal(i) = False Then
                    lngSum = lngSum + Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("����")))
                End If
            Next
            If lngSum > 100 Then
                MsgBox "��ռ��ʽʱ����Լ����֮�Ͳ��ܳ���100��", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        Else
            For i = 1 To vsUnit.Rows - 1
                If vsUnit.IsSubtotal(i) = False Then
                    lngSum = Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("����")))
                    If lngSum > 100 Then
                        MsgBox vsUnit.TextMatrix(i, vsUnit.ColIndex("������λ")) & "��ԤԼ�������ܳ���100��", vbInformation + vbOKOnly, gstrSysName
                        vsUnit.Row = i: vsUnit.Col = vsUnit.ColIndex("����")
                        Exit Function
                    End If
                End If
            Next
        End If
    ElseIf optBespeakMode(1).Value Then '������
        If Not mobj���к��� Is Nothing Then lng��Լ�� = mobj���к���.��Լ��
        If lng��Լ�� > 0 Then '����Լʱ���ü��
            If chkOnlyOneUse.Value = vbChecked Then
                For i = 1 To vsUnit.Rows - 1
                    If vsUnit.IsSubtotal(i) = False Then
                        lngSum = lngSum + Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("����")))
                    End If
                Next
                If lngSum > lng��Լ�� Then
                    MsgBox "��ռ��ʽʱ����Լ��֮�Ͳ��ܳ�����Լ��(" & lng��Լ�� & ")��", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            Else
                For i = 1 To vsUnit.Rows - 1
                    If vsUnit.IsSubtotal(i) = False Then
                        lngSum = Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("����")))
                        If lngSum > lng��Լ�� Then
                            MsgBox vsUnit.TextMatrix(i, vsUnit.ColIndex("������λ")) & "����Լ�����ܳ�����Լ��(" & lng��Լ�� & ")��", vbInformation + vbOKOnly, gstrSysName
                            vsUnit.Row = i: vsUnit.Col = vsUnit.ColIndex("����")
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
    Else '�����
        If Not mobj���к��� Is Nothing Then
            If mobj���к���.�Ƿ��ʱ�� And mobj���к���.�Ƿ���ſ��� = False Then
                For k = 1 To vsfNotSelNum.Rows - 1
                    lngSum = Val(vsfNotSelNum.Cell(flexcpData, k, vsfNotSelNum.ColIndex("����")))
                    lngSN = Val(vsfNotSelNum.TextMatrix(k, vsfNotSelNum.ColIndex("���")))
                    For i = vsfSelNum.LBound To vsfSelNum.UBound
                        For j = 1 To vsfSelNum(i).Rows - 1
                            If Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("���"))) = lngSN Then
                                lngSum = lngSum - Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("����")))
                            End If
                        Next
                    Next
                    If lngSum < 0 Then
                        MsgBox vsfNotSelNum.Cell(flexcpData, k, vsfNotSelNum.ColIndex("ʱ���")) & _
                            " �����ԤԼ�������˸�ʱ�εĿ�ԤԼ����(" & Val(vsfNotSelNum.Cell(flexcpData, k, vsfNotSelNum.ColIndex("����"))) & ")��", _
                            vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                Next
            End If
        End If
    End If
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    On Error Resume Next
    lblEdit.BackColor = UserControl.BackColor
    optBespeakMode(0).BackColor = UserControl.BackColor
    optBespeakMode(1).BackColor = UserControl.BackColor
    optBespeakMode(2).BackColor = UserControl.BackColor
    chkOnlyOneUse.BackColor = UserControl.BackColor
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_IsDataChanged = m_def_IsDataChanged
    m_EditMode = m_def_EditMode
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
End Sub

Private Sub UserControl_Terminate()
    Set mobj������λ�� = Nothing
    Set mobj���к��� = Nothing
    Set mobj���к�����λ = Nothing
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
End Sub

Private Sub vsfNotSelNum_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetButtonEnable(SelectedVsfGridIndex)
End Sub

Private Sub vsfNotSelNum_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
End Sub

Private Sub vsfNotSelNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfNotSelNum_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
End Sub

Private Sub vsfSelNum_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
    RemoveItem vsfNotSelNum, vsfSelNum(index), _
        Val(vsfSelNum(index).TextMatrix(Row, vsfSelNum(index).ColIndex("���"))), _
        True, Val(vsfSelNum(index).TextMatrix(Row, vsfSelNum(index).ColIndex("����")))
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub vsfSelNum_AfterRowColChange(index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    Call SetButtonEnable(index)
    If vsfSelNum(index).Editable = flexEDKbdMouse Then
        vsfNotSelNum.Row = NewRow
        vsfNotSelNum.TopRow = vsfSelNum(index).TopRow
    End If
End Sub

Private Sub vsfSelNum_BeforeEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
    If vsfSelNum(index).ColIndex("����") <> Col Then Cancel = True: Exit Sub
End Sub

Private Sub vsfSelNum_EnterCell(index As Integer)
    If vsfSelNum(index).Col = vsfSelNum(index).ColIndex("����") Then vsfSelNum(index).EditCell
End Sub

Private Sub vsfSelNum_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And vsfSelNum(index).Editable = flexEDKbdMouse Then
        If vsfSelNum(index).Row = vsfSelNum(index).Rows - 1 And vsfSelNum(index).Col = vsfSelNum(index).Cols - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlVsMoveGridCell(vsfSelNum(index), 2)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub vsfSelNum_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfSelNum_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '����λ�����ƣ�����λ���Ȳ��ܴ���9
    If InStr(vsfSelNum(index).EditText, ".") > 0 Then
        If InStr(vsfSelNum(index).EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsfSelNum(index).EditText) >= 9 Then KeyAscii = 0
    End If
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsfSelNum_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Long, lngSN As Long
    Dim i As Integer, j As Integer, lngRow As Long
    
    On Error GoTo Errhand
    '����λ����9λ��ֱ�ӽص�,��ֹ���
    If InStr(vsfSelNum(index).EditText, ".") > 0 Then
        If InStr(vsfSelNum(index).EditText, ".") > 9 Then
            vsfSelNum(index).EditText = Left(vsfSelNum(index).EditText, 9)
        End If
    Else
        vsfSelNum(index).EditText = Left(vsfSelNum(index).EditText, 9)
    End If
    
    lngSN = Val(vsfSelNum(index).TextMatrix(Row, vsfSelNum(index).ColIndex("���")))
    For i = 1 To vsfNotSelNum.Rows - 1
        If Val(vsfNotSelNum.TextMatrix(i, vsfNotSelNum.ColIndex("���"))) = lngSN Then
            lngRow = i: Exit For
        End If
    Next
    '����ʣ������
    lngSum = Val(vsfNotSelNum.Cell(flexcpData, lngRow, vsfNotSelNum.ColIndex("����")))
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        If i <> index Then
            For j = 1 To vsfSelNum(i).Rows - 1
                If Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("���"))) = lngSN Then
                    lngSum = lngSum - Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("����")))
                    Exit For
                End If
            Next
        End If
    Next
    
    If Val(vsfSelNum(index).EditText) > lngSum Then
        MsgBox picUnit(index).Tag & " ԤԼ��(" & Val(vsfSelNum(index).EditText) & ")���ܳ���ʣ��ԤԼ����(" & lngSum & ")��", vbInformation + vbOKOnly, gstrSysName
        Cancel = True: mblnValiedCanSave = False: Exit Sub
    End If
    vsfSelNum(index).EditText = FormatEx(Val(vsfSelNum(index).EditText), 0)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsUnit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = vsUnit.ColIndex("��ֹԤԼ") Then
        If vsUnit.TextMatrix(Row, vsUnit.ColIndex("��ֹԤԼ")) = True Then
            vsUnit.TextMatrix(Row, vsUnit.ColIndex("����")) = ""
            vsUnit.Cell(flexcpBackColor, Row, vsUnit.ColIndex("����")) = vbButtonFace
        Else
            vsUnit.Cell(flexcpBackColor, Row, vsUnit.ColIndex("����")) = vsUnit.BackColor
        End If
    End If
End Sub

Private Sub vsUnit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
    If vsUnit.IsSubtotal(Row) Then Cancel = True: Exit Sub
    If Col = vsUnit.ColIndex("������λ") Then Cancel = True: Exit Sub
    If Col = vsUnit.ColIndex("����") Then
        If Abs(Val(vsUnit.TextMatrix(Row, vsUnit.ColIndex("��ֹԤԼ")))) = 1 Then Cancel = True: Exit Sub
    End If
    '���¼�AfterEdit���������Ϊ�����ڱ༭ʱֱ�Ӱ����棬��鲻��
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub vsUnit_EnterCell()
    If vsUnit.Col = vsUnit.ColIndex("����") Then vsUnit.EditCell
End Sub

Private Sub vsUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vsUnit.Row = vsUnit.Rows - 1 And vsUnit.Col = vsUnit.Cols - 1 Then
            'Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlVsMoveGridCell(vsUnit, 1)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub vsUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsUnit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '����λ�����ƣ�����λ���Ȳ��ܴ���9
    If InStr(vsUnit.EditText, ".") > 0 Then
        If InStr(vsUnit.EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsUnit.EditText) >= 9 Then KeyAscii = 0
    End If
    If optBespeakMode(0).Value Then
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Public Property Let ԤԼ����(ByVal vNewValue As Byte)
    Dim i As Integer, j As Integer
    
    On Error GoTo Errhand
    If mobj���к��� Is Nothing Then Set mobj���к��� = New ������Ϣ��
    mobj���к���.ԤԼ���� = vNewValue
    '0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
    Call UnitPageVisible(mobj���к���.ԤԼ���� <> 2)
    
    '�������
    If mobj���к���.ԤԼ���� = 2 Then
        For i = 1 To vsUnit.Rows - 1
            If vsUnit.RowData(i) = 1 And vsUnit.IsSubtotal(i) = False Then
                vsUnit.TextMatrix(i, vsUnit.ColIndex("��ֹԤԼ")) = 1
                vsUnit.TextMatrix(i, vsUnit.ColIndex("����")) = ""
                vsUnit.Cell(flexcpBackColor, i, vsUnit.ColIndex("����")) = vbButtonFace
            End If
        Next
        For i = vsfSelNum.LBound To vsfSelNum.UBound
            If Val(vsfSelNum(i).Tag) = Val("1-�����ṹ") Then
                chkForbidBespeak(i).Value = vbChecked
                For j = 1 To vsfSelNum(i).Rows - 1
                    If mobj���к���.�Ƿ���ſ��� Then
                        If j > vsfSelNum(i).Rows - 1 Then Exit For
                        RemoveItem vsfSelNum(i), vsfNotSelNum, Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("���")))
                        j = j - 1
                    Else
                        RemoveItem vsfNotSelNum, vsfSelNum(i), Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("���"))), True, 0
                    End If
                Next
            End If
        Next
        Call SetButtonEnable(SelectedVsfGridIndex)
    End If
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Property Let ���к�����Ϣ��(ByVal vNewValue As ������Ϣ��)
    Err = 0: On Error GoTo Errhand
    Set mobj���к��� = vNewValue
    If mobj���к��� Is Nothing Then Set mobj���к��� = New ������Ϣ��
    Set mobj������λ�� = Get������λ���Ƽ�
    Call InitData
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Private Sub vsUnit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Double, lng��Լ�� As Long
    Dim i As Long
    
    On Error GoTo Errhand
    '�༭��ֹԤԼ��ʱ�����
    If Col = vsUnit.ColIndex("��ֹԤԼ") Then Exit Sub
    '����λ����9λ��ֱ�ӽص�,��ֹ���
    If InStr(vsUnit.EditText, ".") > 0 Then
        If InStr(vsUnit.EditText, ".") > 9 Then
            vsUnit.EditText = Left(vsUnit.EditText, 9)
        End If
    Else
        vsUnit.EditText = Left(vsUnit.EditText, 9)
    End If
    
    If chkOnlyOneUse.Value = vbChecked Then
        For i = 1 To vsUnit.Rows - 1
            If i <> vsUnit.Row And vsUnit.IsSubtotal(i) = False Then
                lngSum = lngSum + Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("����")))
            End If
        Next
        lngSum = lngSum + Val(vsUnit.EditText)
        If optBespeakMode(0).Value Then '������
            If lngSum > 100 Then
                MsgBox "��ռ��ʽʱ��������λ������Լ����֮�Ͳ��ܳ���100��", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        ElseIf optBespeakMode(1).Value Then '������
            If Not mobj���к��� Is Nothing Then lng��Լ�� = mobj���к���.��Լ��
            If lng��Լ�� > 0 Then '����Լʱ���ü��
                If lngSum > lng��Լ�� Then
                    MsgBox "��ռ��ʽʱ��������λ������Լ��֮�Ͳ��ܳ�����Լ��(" & lng��Լ�� & ")��", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True: mblnValiedCanSave = False: Exit Sub
                End If
            End If
        End If
    Else
        lngSum = Val(vsUnit.EditText)
        If optBespeakMode(0).Value Then '������
            If lngSum > 100 Then
                MsgBox vsUnit.TextMatrix(vsUnit.Row, vsUnit.ColIndex("������λ")) & " ԤԼ�������ܳ���100��", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        ElseIf optBespeakMode(1).Value Then  '������
            If Not mobj���к��� Is Nothing Then lng��Լ�� = mobj���к���.��Լ��
            If lng��Լ�� > 0 Then '����Լʱ���ü��
                If lngSum > lng��Լ�� Then
                    MsgBox vsUnit.TextMatrix(vsUnit.Row, vsUnit.ColIndex("������λ")) & " ��Լ�����ܳ�����Լ��(" & lng��Լ�� & ")��", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True: mblnValiedCanSave = False: Exit Sub
                End If
            End If
        End If
    End If
    vsUnit.EditText = FormatEx(Val(vsUnit.EditText), 2)
    vsUnit.EditText = IIf(vsUnit.EditText = "0", "", vsUnit.EditText)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property

Public Property Let IsDataChanged(ByVal New_IsDataChanged As Boolean)
    m_IsDataChanged = New_IsDataChanged
    PropertyChanged "IsDataChanged"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=14,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    Dim i As Integer
    
    m_EditMode = IIf(New_EditMode = ED_RegistPlan_UpdateUnit, ED_RegistPlan_Edit, New_EditMode)
    If mobj���к��� Is Nothing Then
        m_EditMode = ED_RegistPlan_View
    ElseIf m_EditMode = ED_RegistPlan_Edit And mobj���к���.ԤԼ���� = Val("1-��ֹԤԼ") Then
        m_EditMode = ED_RegistPlan_View
    End If
    PropertyChanged "EditMode"
    
    For i = optBespeakMode.LBound To optBespeakMode.UBound
        optBespeakMode(i).Enabled = m_EditMode = ED_RegistPlan_Edit
    Next
    chkOnlyOneUse.Enabled = m_EditMode = ED_RegistPlan_Edit
    vsUnit.Editable = IIf(m_EditMode = ED_RegistPlan_Edit, flexEDKbdMouse, flexEDNone)
    picFun.Enabled = m_EditMode = ED_RegistPlan_Edit
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        chkForbidBespeak(i).Enabled = m_EditMode = ED_RegistPlan_Edit
        vsfSelNum(i).Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(i).Value = Unchecked And mblnEdit, flexEDKbdMouse, flexEDNone)
    Next
    
    '����
    SetUnitVisible
    UserControl_Resize
End Property
