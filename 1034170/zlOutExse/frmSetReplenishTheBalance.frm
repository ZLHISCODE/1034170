VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSetReplenishTheBalance 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "�˷�Ʊ�ݴ�ӡ����(&2)"
      Height          =   350
      Index           =   2
      Left            =   5790
      TabIndex        =   37
      Top             =   5160
      Width           =   1950
   End
   Begin VB.PictureBox picDelBillFormat 
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   2130
      ScaleHeight     =   1380
      ScaleWidth      =   6015
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2040
      Width           =   6015
      Begin VSFlex8Ctl.VSFlexGrid vsDelBillFormat 
         Height          =   1350
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   5865
         _cx             =   10345
         _cy             =   2381
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
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSetReplenishTheBalance.frx":0000
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
         ExplorerBar     =   2
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
   Begin VB.PictureBox picBillFormat 
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   2910
      ScaleHeight     =   1605
      ScaleWidth      =   5925
      TabIndex        =   32
      Top             =   1710
      Width           =   5925
      Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
         Height          =   1365
         Left            =   30
         TabIndex        =   33
         Top             =   30
         Width           =   5715
         _cx             =   10081
         _cy             =   2408
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
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSetReplenishTheBalance.frx":0096
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
         ExplorerBar     =   2
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
   Begin VB.TextBox txtVaildDays 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "3"
      Top             =   4650
      Width           =   450
   End
   Begin VB.Frame fra���㷽ʽ 
      Caption         =   "�������������"
      Height          =   3165
      Left            =   6180
      TabIndex        =   2
      Top             =   120
      Width           =   1605
      Begin VB.ListBox lst���㷽ʽ 
         Height          =   2790
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.OptionButton optDrug 
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   3525
      TabIndex        =   27
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optDrug 
      Caption         =   "��ֹ"
      Height          =   180
      Index           =   1
      Left            =   2715
      TabIndex        =   26
      Top             =   5820
      Width           =   735
   End
   Begin VB.OptionButton optDrug 
      Caption         =   "�����"
      Height          =   180
      Index           =   0
      Left            =   1815
      TabIndex        =   25
      Top             =   5820
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox chkPayKey 
      Caption         =   "ʹ��С���̵ļӼ�(+-)���л�֧����ʽ"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5310
      Width           =   3375
   End
   Begin VB.TextBox txtƱ������ 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "10"
      Top             =   4950
      Width           =   465
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   5790
      TabIndex        =   16
      Top             =   4260
      Width           =   1950
   End
   Begin VB.Frame fraTitle 
      Caption         =   "���ع����շ�Ʊ��"
      Height          =   1410
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5925
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   5670
         _cx             =   10001
         _cy             =   1931
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
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSetReplenishTheBalance.frx":012C
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
         ExplorerBar     =   2
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
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "�����嵥��ӡ����(&3)"
      Height          =   350
      Index           =   1
      Left            =   5790
      TabIndex        =   12
      Top             =   5610
      Width           =   1950
   End
   Begin VB.Frame fraFeeList 
      Caption         =   "�����嵥��ӡ��ʽ"
      Height          =   675
      Left            =   1980
      TabIndex        =   7
      Top             =   3390
      Width           =   5760
      Begin VB.OptionButton optPrint 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   585
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   2
         Left            =   3690
         TabIndex        =   10
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6675
      Left            =   7860
      TabIndex        =   31
      Top             =   -600
      Width           =   45
   End
   Begin VB.Frame fra��λ 
      Caption         =   " ҩƷ��ʾ��λ "
      Height          =   1155
      Left            =   135
      TabIndex        =   4
      Top             =   3390
      Width           =   1635
      Begin VB.OptionButton opt��λ 
         Caption         =   "�ۼ۵�λ"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton opt��λ 
         Caption         =   "���ﵥλ"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   780
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "�շ�Ʊ�ݴ�ӡ����(&1)"
      Height          =   350
      Index           =   0
      Left            =   5790
      TabIndex        =   11
      Top             =   4710
      Width           =   1950
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8100
      TabIndex        =   28
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8100
      TabIndex        =   29
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   8100
      TabIndex        =   30
      Top             =   5310
      Width           =   1100
   End
   Begin MSComCtl2.UpDown updƱ������ 
      Height          =   300
      Left            =   1635
      TabIndex        =   19
      Top             =   4950
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "txtƱ������"
      BuddyDispid     =   196617
      OrigLeft        =   1605
      OrigTop         =   4860
      OrigRight       =   1860
      OrigBottom      =   5160
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox chkƱ������ 
      Caption         =   "Ʊ��ʣ��         ��ʱ��ʼ�����շ�Ա"
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   4965
      Width           =   3450
   End
   Begin VB.TextBox txtNameDays 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   22
      Text            =   "0"
      ToolTipText     =   "0��ʾ����ʱ������ʱ��"
      Top             =   5580
      Width           =   285
   End
   Begin VB.Frame fraShortLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   2730
      TabIndex        =   23
      Top             =   5760
      Width           =   285
   End
   Begin VB.CheckBox chkSeekName 
      Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   5565
      Width           =   4260
   End
   Begin MSComCtl2.UpDown updVaildDays 
      Height          =   270
      Left            =   3480
      TabIndex        =   15
      Top             =   4650
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "txtVaildDays"
      BuddyDispid     =   196612
      OrigLeft        =   1605
      OrigTop         =   4860
      OrigRight       =   1860
      OrigBottom      =   5160
      Max             =   100
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl tbBillSet 
      Height          =   1800
      Left            =   150
      TabIndex        =   36
      Top             =   1560
      Width           =   5925
      _Version        =   589884
      _ExtentX        =   10451
      _ExtentY        =   3175
      _StockProps     =   64
   End
   Begin VB.Label lblVaildDays 
      Caption         =   "�ɽ��б��ղ������ķ�����Ч����"
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label lblDrugNotFee 
      AutoSize        =   -1  'True
      Caption         =   "ҩƷ��ҩ���˷ѷ�ʽ"
      Height          =   180
      Left            =   150
      TabIndex        =   24
      Top             =   5820
      Width           =   1620
   End
End
Attribute VB_Name = "frmSetReplenishTheBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String
Private mlngModule As Long
Private mblnOK As Boolean

Public Function zlSetPara(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:�������óɹ�,����true,����ķ���False
    '����:���ϴ�
    '����:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOK = False
    
    Me.Show 1, frmMain
    zlSetPara = mblnOK
End Function

Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1 And txtNameDays.Tag = "1"
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, mlngModule)
End Sub

Private Sub cmdOK_Click()
    Dim blnHavePrivs As Boolean, i As Integer
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    If isValied = False Then Exit Sub
    zlDatabase.SetPara "ҩƷ��λ��ʾ", IIf(opt��λ(0).Value, 0, 1), glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "����ģ�����ҷ�ʽ", IIf(chkSeekName.Value = 1, "1", "0") & "|" & Val(txtNameDays.Text), glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", IIf(chkƱ������.Value = 1, "1", "0") & "|" & Val(txtƱ������.Text), glngSys, mlngModule, blnHavePrivs
    For i = 0 To optPrint.UBound
        If optPrint(i).Value Then
            zlDatabase.SetPara "�����嵥��ӡ��ʽ", i, glngSys, mlngModule, blnHavePrivs
        End If
    Next
    Call SaveInvoice
    '47457,82343
    zlDatabase.SetPara "ʹ�üӼ��л�֧����ʽ", IIf(chkPayKey.Value = 1, "1", "0"), glngSys, mlngModule, blnHavePrivs
    '47400,82343
    zlDatabase.SetPara "ҩƷ��ҩ�˷ѷ�ʽ", IIf(optDrug(0).Value, 0, IIf(optDrug(1).Value, "1", "2")), glngSys, mlngModule, blnHavePrivs
    '84929
    zlDatabase.SetPara "��������Ч����", Val(txtVaildDays.Text), glngSys, mlngModule, blnHavePrivs
    mblnOK = True
    Unload Me
End Sub

Private Sub chkƱ������_Click()
    txtƱ������.Enabled = chkƱ������.Enabled And chkƱ������.Value = 1
    updƱ������.Enabled = txtƱ������.Enabled
End Sub

Private Sub cmdPrintSetup_Click(Index As Integer)
    Select Case Index
        Case 0 '����ҽ�Ʒ��շ�
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1124", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124", Me)
            End If
        Case 1 '�����շ��嵥
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1124_1", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me)
            End If
        Case 2 '�˷ѷ�Ʊ(��Ʊ)
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1124_3", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_3", Me)
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim strTmp As String, blnParSet As Boolean, i As Integer
    
    On Error GoTo errH
    blnParSet = InStr(1, mstrPrivs, "��������") > 0
    Call InitTabControl
    
    i = IIf(zlDatabase.GetPara("ҩƷ��λ��ʾ", glngSys, mlngModule, , Array(opt��λ(0), opt��λ(1)), blnParSet) = "0", 0, 1)
    opt��λ(i).Value = True
    txtNameDays.Enabled = True
    strTmp = zlDatabase.GetPara("����ģ�����ҷ�ʽ", glngSys, mlngModule, "0|10", Array(txtNameDays, chkSeekName), blnParSet)
    txtNameDays.Text = Val(Split(strTmp & "|", "|")(1))
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
    
    strTmp = zlDatabase.GetPara("Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", glngSys, mlngModule, "0|10", Array(txtƱ������, updƱ������, chkƱ������), blnParSet)
    
    updƱ������.Value = Val(Split(strTmp & "|", "|")(1))
    txtƱ������.Text = updƱ������.Value
    chkƱ������.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
    txtƱ������.Enabled = chkƱ������.Enabled And chkƱ������.Value = 1
    updƱ������.Enabled = txtƱ������.Enabled
    
    i = Val(zlDatabase.GetPara("�����嵥��ӡ��ʽ", glngSys, mlngModule, , Array(optPrint(0), optPrint(1), optPrint(2)), blnParSet))
    If i <= optPrint.UBound Then optPrint(i).Value = True
    Call InitShareInvoice
    '47457,82343
    chkPayKey.Value = IIf(Val(zlDatabase.GetPara("ʹ�üӼ��л�֧����ʽ", glngSys, mlngModule, "1", Array(chkPayKey), blnParSet)) = 1, 1, 0)
    '47400,82343
    strTmp = zlDatabase.GetPara("ҩƷ��ҩ�˷ѷ�ʽ", glngSys, mlngModule, , Array(optDrug(0), optDrug(1), optDrug(2)), blnParSet)
    For i = 0 To 2
        If Val(strTmp) = i Then
            optDrug(i).Value = True: Exit For
        End If
    Next
    '84929
    txtVaildDays.Text = Val(zlDatabase.GetPara("��������Ч����", glngSys, mlngModule, "3", Array(lblVaildDays, txtVaildDays, updVaildDays), blnParSet))
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-04-28 15:09:10
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModule, vsBill, Me.Name, "����Ʊ��������", False, False
    zl_vsGrid_Para_Restore mlngModule, vsBillFormat, Me.Name, "�շѷ�Ʊ��ӡ��ʽ", False, False
    zl_vsGrid_Para_Restore mlngModule, vsDelBillFormat, Me.Name, "�˷ѷ�Ʊ��ӡ��ʽ", False, False
    strShareInvoice = zlDatabase.GetPara("�����շ�Ʊ������", glngSys, mlngModule, , , True, intType)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
     '��ʽ:����ID1,ʹ�����1|����IDn,ʹ�����n|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(1)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!ʹ�����, " ")
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    'Ʊ�ݸ�ʽ����
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���='ZL" & glngSys \ 100 & "_BILL_1124'  " & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    
    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("�շѷ�Ʊ��ʽ", glngSys, mlngModule, , , True, intType)
    strPrintMode = zlDatabase.GetPara("�շѷ�Ʊ��ӡ��ʽ", glngSys, mlngModule, , , True, intType1)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat
         .ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = "0"
         .ColData(.ColIndex("�շѴ�ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�շѴ�ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ"))) = 1 Or _
            Val(.ColData(.ColIndex("�շѴ�ӡ��ʽ"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    
    '80943,Ƚ����,2014-12-18,Ʊ��δʹ�á��շ����ʱ�����������շ����Ϊ�յĴ�ӡ��ʽ��Ʊ�ݸ�ʽ
    Dim objInvoice As New zlPublicExpense.clsInvoice
    If objInvoice.zlStartFactUseType(EM_�շ��վ�) Then
        strSQL = "" & _
            "   Select ����, ���� From Ʊ��ʹ�����" & _
            "   Order By ����"
    Else
        strSQL = "" & _
            "   Select ����, ���� From Ʊ��ʹ�����" & _
            "   Union All" & _
            "   Select '', '' From Dual " & _
            "   Order By ����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("�շѴ�ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�շѴ�ӡ��ʽ"), .Rows - 1, .ColIndex("�շѴ�ӡ��ʽ")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�շ�Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    
    '====================================================================
    '�˷�Ʊ�ݸ�ʽ����
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���='ZL" & glngSys \ 100 & "_BILL_1124_3'  " & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .ColComboList(.ColIndex("�˷�Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�˷Ѵ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    
    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("�˷ѷ�Ʊ��ʽ", glngSys, mlngModule, , , True, intType)
    strPrintMode = zlDatabase.GetPara("�˷ѷ�Ʊ��ӡ��ʽ", glngSys, mlngModule, , , True, intType1)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsDelBillFormat
         .ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ")) = "0"
         .ColData(.ColIndex("�˷Ѵ�ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�˷Ѵ�ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ"))) = 1 Or _
            Val(.ColData(.ColIndex("�˷Ѵ�ӡ��ʽ"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    
    '80943,Ƚ����,2014-12-18,Ʊ��δʹ�á��շ����ʱ�����������շ����Ϊ�յĴ�ӡ��ʽ��Ʊ�ݸ�ʽ
    If objInvoice.zlStartFactUseType(EM_�շ��վ�) Then
        strSQL = "" & _
            "   Select ����, ���� From Ʊ��ʹ�����" & _
            "   Order By ����"
    Else
        strSQL = "" & _
            "   Select ����, ���� From Ʊ��ʹ�����" & _
            "   Union All" & _
            "   Select '', '' From Dual " & _
            "   Order By ����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�˷Ѵ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("�˷�Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�˷�Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�˷Ѵ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("�˷Ѵ�ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�˷Ѵ�ӡ��ʽ"), .Rows - 1, .ColIndex("�˷Ѵ�ӡ��ʽ")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�˷�Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("�˷Ѵ�ӡ��ʽ")) = vbBlue
        End If
    End With
    
    '82990:���ϴ�,2015/3/9,�������������
    Call Load���㷽ʽ(blnHavePrivs)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "����Ʊ��������", False, False
    zl_vsGrid_Para_Save mlngModule, vsBillFormat, Me.Name, "�շѷ�Ʊ��ӡ��ʽ", False, False
    zl_vsGrid_Para_Save mlngModule, vsDelBillFormat, Me.Name, "�˷ѷ�Ʊ��ӡ��ʽ", False, False
End Sub

Private Sub fraƱ�ݸ�ʽ_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���淢Ʊ���Ʊ��
    '����:���˺�
    '����:2011-04-28 18:16:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    '���湲��Ʊ��
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "�����շ�Ʊ������", strValue, glngSys, mlngModule, blnHavePrivs
    '�����շѸ�ʽ
    
    Dim strPrintMode As String
    '�����շѸ�ʽ
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�շ�Ʊ�ݸ�ʽ")))
            strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�շѴ�ӡ��ʽ")), 1))
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "�շѷ�Ʊ��ʽ", strValue, glngSys, mlngModule, blnHavePrivs
        zlDatabase.SetPara "�շѷ�Ʊ��ӡ��ʽ", strPrintMode, glngSys, mlngModule, blnHavePrivs
    End With
    
    '====================================================
    '�����˷Ѹ�ʽ
    strValue = "": strPrintMode = ""
    With vsDelBillFormat
        For i = 1 To .Rows - 1
            strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�˷�Ʊ�ݸ�ʽ")))
            strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�˷Ѵ�ӡ��ʽ")), 1))
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "�˷ѷ�Ʊ��ʽ", strValue, glngSys, mlngModule, blnHavePrivs
        zlDatabase.SetPara "�˷ѷ�Ʊ��ӡ��ʽ", strPrintMode, glngSys, mlngModule, blnHavePrivs
    End With
    
    '82990:���ϴ�,2015/3/9,�������������
    strValue = ""
    For i = 0 To lst���㷽ʽ.ListCount - 1
        If lst���㷽ʽ.Selected(i) = True Then
            strValue = strValue & "|" & lst���㷽ʽ.List(i)
        End If
    Next
    strValue = Mid(strValue, 2)
    zlDatabase.SetPara "�������������", strValue, glngSys, mlngModule, blnHavePrivs
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-28 18:24:16
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
     
    isValied = False
    On Error GoTo errHandle
    '���ÿ��ʹ����ʽֻ��һ��ѡ��
    With vsBill
        str��� = "-"
        For i = 1 To vsBill.Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) = Trim(.TextMatrix(j, .ColIndex("ʹ�����"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    ʹ�����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    isValied = True
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtNameDays_GotFocus()
    Call SelAll(txtNameDays)
End Sub

Private Sub txtNameDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNameDays_Validate(Cancel As Boolean)
    If Val(txtNameDays.Text) <= 0 Then
        txtNameDays.Text = 0
    ElseIf Val(txtNameDays.Text) > 999 Then
        txtNameDays.Text = 999
    End If
End Sub

Private Sub Load���㷽ʽ(ByVal blnHavePrivs As Boolean)
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objItem As Object
    Dim str֧����ʽ As String
    
    On Error GoTo errHandle
    '85182:���ϴ�,2015/5/27,��������Ȩ�޿���
    str֧����ʽ = zlDatabase.GetPara("�������������", glngSys, mlngModule, , Array(fra���㷽ʽ, lst���㷽ʽ), blnHavePrivs)
    strSQL = "Select distinct B.����,B.���� From ���㷽ʽӦ�� A,���㷽ʽ B" & vbNewLine & _
            "Where A.Ӧ�ó��� in ('�Һ�','�շ�') And B.����=A.���㷽ʽ" & vbNewLine & _
            "And   (B.����<>3 And B.����<>4)" & vbNewLine & _
            "Order by lpad(����,3,' ')"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ղ������")
    
    lst���㷽ʽ.Clear
    Do Until rsTemp.EOF
        lst���㷽ʽ.AddItem Nvl(rsTemp!����)
        If InStr("|" & str֧����ʽ & "|", "|" & Nvl(rsTemp!����) & "|") > 0 Then lst���㷽ʽ.Selected(lst���㷽ʽ.NewIndex) = True
        rsTemp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitTabControl()
    With tbBillSet
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionTop
'        .PaintManager.StaticFrame = True
'        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .InsertItem 0, "�շ�Ʊ�ݸ�ʽ", picBillFormat.hWnd, 0
        .InsertItem 1, "�˷�Ʊ�ݸ�ʽ", picDelBillFormat.hWnd, 0
        .Item(0).Selected = True
    End With
End Sub

