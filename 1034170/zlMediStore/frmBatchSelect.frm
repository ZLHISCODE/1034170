VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ����ѡ��"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11865
   Icon            =   "frmBatchSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12.753
   ScaleMode       =   0  'User
   ScaleWidth      =   20.929
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDrug 
      Height          =   5535
      Left            =   1560
      ScaleHeight     =   5475
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VSFlex8Ctl.VSFlexGrid vsfDrug 
         Height          =   5470
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3930
         _cx             =   6932
         _cy             =   9648
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBatchSelect.frx":000C
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
   Begin VB.CommandButton cmdQuit 
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   10320
      TabIndex        =   6
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "���(&A)"
      Height          =   300
      Left            =   9240
      TabIndex        =   5
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "���(&O)"
      Height          =   300
      Left            =   8160
      TabIndex        =   4
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   960
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtSelect 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   180
      Width           =   3975
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   10920
      Top             =   120
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
            Picture         =   "frmBatchSelect.frx":0081
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":061B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":6E7D
            Key             =   "���U"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSelectDrug 
      Height          =   5925
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   11775
      _cx             =   20770
      _cy             =   10451
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBatchSelect.frx":7417
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   6780
      Width           =   360
   End
   Begin VB.Label lblCalss 
      AutoSize        =   -1  'True
      Caption         =   "������Ʒ�ּ���"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmBatchSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintUnit As Integer '��ģ�������õ���ʾ��λ 0-ҩ�ⵥλ;1-���ﵥλ;2-סԺ��λ;3-�ۼ۵�λ
Private Const mlngRowHeight As Long = 300 '����и����и�
Private mrsReturn As ADODB.Recordset        '����ѡ��ҩƷ����
Private mblnOk As Boolean   '��¼�Ƿ��ǵ����ȷ����ť
Private mrsFindName As ADODB.Recordset '��¼��ѯ���ݼ�
Private mstrMatch  As String '0-˫��ƥ�� 1-������ƥ��

'����λ
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
Private Const MStrCaption As String = "ҩƷ����ѡ��"

Private Enum vsfSelectDrugCol
    ҩƷID = 0
    ҩƷ��Ϣ = 1
    ҩƷ����
    ��Ʒ��
    ͨ����
    ���
    ����
    ��λ
    �ۼ۵�λ
    ���ﵥλ
    ����ϵ��
    סԺ��λ
    סԺϵ��
    ҩ�ⵥλ
    ҩ��ϵ��
    ����
    �ۼ�
    �ɱ���
    ָ������
    ָ���ۼ�
    ������
End Enum

Public Sub showMe(ByVal frmParent As Form, ByRef rsTemp As ADODB.Recordset, ByRef blnOk As Boolean)
    Me.Show vbModal, frmParent
    blnOk = mblnOk
    Set rsTemp = mrsReturn
End Sub

Private Sub initVsflexgrid()
    With vsfSelectDrug
        .Editable = flexEDNone
        .Cols = vsfSelectDrugCol.������
        .rows = 1
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMove '�ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��

        '�����п�
        .ColWidth(vsfSelectDrugCol.ҩƷID) = 0
        .ColWidth(vsfSelectDrugCol.ҩƷ��Ϣ) = 3000
        .ColWidth(vsfSelectDrugCol.ҩƷ����) = 0
        .ColWidth(vsfSelectDrugCol.��Ʒ��) = 0
        .ColWidth(vsfSelectDrugCol.ͨ����) = 0
        .ColWidth(vsfSelectDrugCol.����) = 1500
        .ColWidth(vsfSelectDrugCol.��λ) = 800
        
        .ColWidth(vsfSelectDrugCol.�ۼ۵�λ) = 0
        .ColWidth(vsfSelectDrugCol.���ﵥλ) = 0
        .ColWidth(vsfSelectDrugCol.����ϵ��) = 0
        .ColWidth(vsfSelectDrugCol.סԺ��λ) = 0
        .ColWidth(vsfSelectDrugCol.סԺϵ��) = 0
        .ColWidth(vsfSelectDrugCol.ҩ�ⵥλ) = 0
        .ColWidth(vsfSelectDrugCol.ҩ��ϵ��) = 0
        
        .ColWidth(vsfSelectDrugCol.����) = 1000
        .ColWidth(vsfSelectDrugCol.�ۼ�) = 1500
        .ColWidth(vsfSelectDrugCol.�ɱ���) = 1500
        .ColWidth(vsfSelectDrugCol.ָ������) = 1500
        .ColWidth(vsfSelectDrugCol.ָ���ۼ�) = 1500
        '������ͷ
        .TextMatrix(0, vsfSelectDrugCol.ҩƷID) = "ҩƷid"
        .TextMatrix(0, vsfSelectDrugCol.ҩƷ��Ϣ) = "ҩƷ��Ϣ"
        .TextMatrix(0, vsfSelectDrugCol.ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, vsfSelectDrugCol.��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, vsfSelectDrugCol.ͨ����) = "ͨ����"
        .TextMatrix(0, vsfSelectDrugCol.���) = "���"
        .TextMatrix(0, vsfSelectDrugCol.����) = "����"
        .TextMatrix(0, vsfSelectDrugCol.��λ) = "��λ"
        
        .TextMatrix(0, vsfSelectDrugCol.�ۼ۵�λ) = "�ۼ۵�λ"
        .TextMatrix(0, vsfSelectDrugCol.���ﵥλ) = "���ﵥλ"
        .TextMatrix(0, vsfSelectDrugCol.����ϵ��) = "����ϵ��"
        .TextMatrix(0, vsfSelectDrugCol.סԺ��λ) = "סԺ��λ"
        .TextMatrix(0, vsfSelectDrugCol.סԺϵ��) = "סԺϵ��"
        .TextMatrix(0, vsfSelectDrugCol.ҩ�ⵥλ) = "ҩ�ⵥλ"
        .TextMatrix(0, vsfSelectDrugCol.ҩ��ϵ��) = "ҩ��ϵ��"
        
        .TextMatrix(0, vsfSelectDrugCol.����) = "����"
        .TextMatrix(0, vsfSelectDrugCol.�ۼ�) = "�ۼ�"
        .TextMatrix(0, vsfSelectDrugCol.�ɱ���) = "�ɱ���"
        .TextMatrix(0, vsfSelectDrugCol.ָ������) = "ָ������"
        .TextMatrix(0, vsfSelectDrugCol.ָ���ۼ�) = "ָ���ۼ�"

        .ColAlignment(vsfSelectDrugCol.ҩƷID) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.ҩƷ��Ϣ) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.ҩƷ����) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.���) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.����) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.��λ) = flexAlignCenterCenter
        .ColAlignment(vsfSelectDrugCol.����) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.�ۼ�) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.�ɱ���) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.ָ������) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.ָ���ۼ�) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

'
'Private Sub setTvwInfo()
'    'Ϊ�����������
'    Dim objNode As Node
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'
'    gstrSQL = " Select ����,���� From ������Ŀ��� " & _
'              " Where Instr([1],����,1) > 0 " & _
'              " Order by ����"
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, mstrCaption, "567")
'
'    If rsTemp Is Nothing Then
'        Exit Sub
'    End If
'
'    With tvwDrug
'        .Nodes.Clear
'        Do While Not rsTemp.EOF
'            .Nodes.Add , , "Root" & rsTemp!����, rsTemp!����, 1, 1
'            .Nodes("Root" & rsTemp!����).Tag = rsTemp!����
'            rsTemp.MoveNext
'        Loop
'    End With
'
'
'    gstrSQL = "Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, '����' As ���" & _
'                " From ���Ʒ���Ŀ¼" & _
'                " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                " Start With �ϼ�id Is Null" & _
'                " Connect By Prior ID = �ϼ�id"
'
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�����ѯ")
'    With rsTemp
'        Do While Not .EOF
'           If IsNull(!�ϼ�id) Then
'                Set objNode = tvwDrug.Nodes.Add("Root" & !����, 4, "K_" & !Id, !���� & "-����", 1, 1)
'            Else
'                Set objNode = tvwDrug.Nodes.Add("K_" & !�ϼ�id, 4, "K_" & !Id, !���� & "-����", 1, 1)
'            End If
'            objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'            .MoveNext
'        Loop
'    End With
'
'    If optVariety.Value = True Then
'        gstrSQL = "Select ID, ����id, ����, ����, Decode(���, 5, '����ҩ', 6, '�г�ҩ', 7, '�в�ҩ') ����, 'Ʒ��' As ���" & _
'                  "  From ������ĿĿ¼" & _
'                  "  Where ����id In (Select ID" & _
'                                   " From ���Ʒ���Ŀ¼" & _
'                                   " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                                   " Start With �ϼ�id Is Null" & _
'                                   " Connect By Prior ID = �ϼ�id)"
'        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "Ʒ��")
'
'        With rsTemp
'            Do While Not .EOF
'                Set objNode = tvwDrug.Nodes.Add("K_" & !����id, 4, !��� & "K_" & !Id, !���� & "-Ʒ��", 1, 1)
'                objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'                .MoveNext
'            Loop
'        End With
'    End If
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

'Private Sub cmdSelect_Click()
'    picDrug.Visible = True
'    tvwDrug.Visible = True
'    Call setTvwInfo
'End Sub

Private Sub cmdCal_Click()
    With vsfSelectDrug
        If MsgBox("ȷ��Ҫ��������Ѿ�ѡ���ҩƷ��", vbYesNo, gstrSysName) = vbYes Then
            .rows = 1
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    Dim intRow As Integer
    Set mrsReturn = New ADODB.Recordset

    With mrsReturn
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ͨ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ʱ��", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        
        .Fields.Append "�ۼ۵�λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "���ﵥλ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "�����װ", adDouble, 11, adFldIsNullable
        .Fields.Append "סԺ��λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "סԺ��װ", adDouble, 11, adFldIsNullable
        .Fields.Append "ҩ�ⵥλ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ҩ���װ", adDouble, 11, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    With vsfSelectDrug
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, vsfSelectDrugCol.ҩƷID) = "" Then Exit For
            mrsReturn.AddNew
            mrsReturn!ҩƷID = .TextMatrix(intRow, vsfSelectDrugCol.ҩƷID)
            mrsReturn!ҩƷ���� = .TextMatrix(intRow, vsfSelectDrugCol.ҩƷ����)
            mrsReturn!��Ʒ�� = .TextMatrix(intRow, vsfSelectDrugCol.��Ʒ��)
            mrsReturn!ͨ���� = .TextMatrix(intRow, vsfSelectDrugCol.ͨ����)
            mrsReturn!��� = .TextMatrix(intRow, vsfSelectDrugCol.���)
            mrsReturn!ʱ�� = IIf(.TextMatrix(intRow, vsfSelectDrugCol.����) = "ʱ��", 1, 0)
            mrsReturn!���� = .TextMatrix(intRow, vsfSelectDrugCol.����)
            mrsReturn!�ۼ۵�λ = .TextMatrix(intRow, vsfSelectDrugCol.�ۼ۵�λ)
            mrsReturn!���ﵥλ = .TextMatrix(intRow, vsfSelectDrugCol.���ﵥλ)
            mrsReturn!�����װ = .TextMatrix(intRow, vsfSelectDrugCol.����ϵ��)
            mrsReturn!סԺ��λ = .TextMatrix(intRow, vsfSelectDrugCol.סԺ��λ)
            mrsReturn!סԺ��װ = .TextMatrix(intRow, vsfSelectDrugCol.סԺϵ��)
            mrsReturn!ҩ�ⵥλ = .TextMatrix(intRow, vsfSelectDrugCol.ҩ�ⵥλ)
            mrsReturn!ҩ���װ = .TextMatrix(intRow, vsfSelectDrugCol.ҩ��ϵ��)
            
            mrsReturn.Update
        Next
    End With
    mblnOk = True
    
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdSelect_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picDrug.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    '��ȡ���õĵ�λ
    mintUnit = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, 1333, 1))
    Select Case mintUnit
        Case 0 'ҩ��
            intUnitTemp = 4
        Case 1 'סԺ
            intUnitTemp = 3
        Case 2 '����
            intUnitTemp = 2
        Case 3 '�ۼ�
            intUnitTemp = 1
    End Select
    '��ȡ������λ����
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)
    
    mstrMatch = IIf(zlDatabase.GetPara("����ƥ��", , , 0) = "0", "%", "")
    mblnOk = False
    Call initVsflexgrid
    
    Call RestoreWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub optClass_Click()
    picDrug.Visible = False
    lblCalss.Caption = "����"
End Sub

Private Sub optClassSub_Click()
    picDrug.Visible = False
    lblCalss.Caption = "����(������)"
End Sub

Private Sub optVariety_Click()
    picDrug.Visible = False
    lblCalss.Caption = "Ʒ��"
End Sub

'Private Sub tvwDrug_NodeClick(ByVal Node As MSComctlLib.Node)
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'    If Node.Key Like "Root" Then Exit Sub
'
'    gstrSQL = "select id,����,����,���㵥λ from ������ĿĿ¼ where  Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' and ����id=[1]"
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", Mid(Node.Key, InStr(1, Node.Key, "_") + 1))
'
'    Set vsfDetails.DataSource = rsTemp
'
'    Exit Sub
'errHandle:
'    If errcenter() = 1 Then Resume
'    Call saveerrlog
'End Sub

'Private Sub tvwDrug_DblClick()
'    '����������д���ֵ
'    Dim lngId As Long
'    Dim rsTemp As ADODB.Recordset
'    Dim intRow As Integer
'    Dim i As Integer
'    Dim blnDou As Boolean '�ظ�����
'    Dim dbl����ϵ�� As Double
'    Dim strUnit As String   '��λ
'
'    On Error GoTo errHandle
'    With tvwDrug
'        If optVariety.Value = True Then
'            If InStr(1, .SelectedItem.Text, "-Ʒ��") <= 0 Then
'                Exit Sub
'            End If
'            gstrSQL = "Select Distinct a.ҩƷid, c.���� As ҩƷ����, c.���� As ͨ����, d.��Ʒ��, c.���, c.�Ƿ��� As ʱ��, c.����, c.���㵥λ As �ۼ۵�λ, a.���ﵥλ, a.�����װ," & _
'                                        " a.סԺ��λ , a.סԺ��װ, a.ҩ�ⵥλ, a.ҩ���װ, a.�ɱ���, e.�ּ�, a.ָ��������, a.ָ�����ۼ�" & _
'                        " From ҩƷ��� A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) D,�շѼ�Ŀ E" & _
'                        " Where a.ҩ��id = b.Id And a.ҩƷid = c.Id And c.Id = d.�շ�ϸĿid(+) and a.ҩƷid=e.�շ�ϸĿid and sysdate between e.ִ������ and e.��ֹ���� and b.id=[1] order by c.����"
'        Else
'            If InStr(1, .SelectedItem.Text, "-����") <= 0 Then
'                Exit Sub
'            End If
'                If optClassSub.Value = True Then '�������ӽڵ�
'                    gstrSQL = "(Select ID From ���Ʒ���Ŀ¼ Where ���� In (1, 2, 3) Start With ID = [1] Connect By Prior ID = �ϼ�id) A,"
'                Else '������
'                    gstrSQL = "(select id from ���Ʒ���Ŀ¼ where ���� in (1,2,3) and id=[1]) A,"
'                End If
'
'                gstrSQL = "Select Distinct c.ҩƷid, d.���� As ҩƷ����, d.���� As ͨ����, f.��Ʒ��, d.���, d.�Ƿ��� As ʱ��, d.����, d.���㵥λ As �ۼ۵�λ, c.���ﵥλ, c.�����װ," & _
'                                        " c.סԺ��λ , c.סԺ��װ, c.ҩ�ⵥλ, c.ҩ���װ, c.�ɱ���, e.�ּ�, c.ָ��������, c.ָ�����ۼ� " & _
'                        " From " & gstrSQL & " ������ĿĿ¼ B, ҩƷ��� C," & _
'                             " �շ���ĿĿ¼ D, �շѼ�Ŀ E, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) F" & _
'                        " Where a.Id = b.����id And b.Id = c.ҩ��id And c.ҩƷid = d.Id And d.Id = e.�շ�ϸĿid And e.�շ�ϸĿid = f.�շ�ϸĿid(+) And" & _
'                              " Sysdate Between e.ִ������ And e.��ֹ���� order by d.����"
'        End If
'        lngId = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "K_") + 2)
'
'        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯҩƷ", lngId)
'        If rsTemp.RecordCount = 0 Then
'            Exit Sub
'        End If
'    End With
'
'    With vsfSelectDrug
'        For intRow = 0 To rsTemp.RecordCount - 1
'            blnDou = False
'            For i = 1 To .rows - 1
'                If .TextMatrix(i, vsfSelectDrugCol.ҩƷid) = rsTemp!ҩƷid Then
'                    blnDou = True
'                End If
'            Next
'            If blnDou = False Then
'                .rows = .rows + 1
'                .RowHeight(.rows - 1) = mlngRowHeight
'
'                Select Case mintUnit
'                    Case 0
'                        dbl����ϵ�� = rsTemp!ҩ���װ
'                        strUnit = rsTemp!ҩ�ⵥλ
'                    Case 1
'                        dbl����ϵ�� = rsTemp!סԺ��װ
'                        strUnit = rsTemp!סԺ��λ
'                    Case 2
'                        dbl����ϵ�� = rsTemp!�����װ
'                        strUnit = rsTemp!���ﵥλ
'                    Case 3
'                        dbl����ϵ�� = 1
'                        strUnit = rsTemp!�ۼ۵�λ
'                End Select
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷid) = rsTemp!ҩƷid
'                If gintҩƷ������ʾ = 1 Then
'                    .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ��Ϣ) = "[" & rsTemp!ҩƷ���� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
'                Else
'                    .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ��Ϣ) = "[" & rsTemp!ҩƷ���� & "]" & rsTemp!ͨ����
'                End If
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ����) = rsTemp!ҩƷ����
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.��Ʒ��) = IIf(IsNull(rsTemp!��Ʒ��), "", rsTemp!��Ʒ��)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.ͨ����) = IIf(IsNull(rsTemp!ͨ����), "", rsTemp!ͨ����)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.��λ) = strUnit
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ۼ۵�λ) = rsTemp!�ۼ۵�λ
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.���ﵥλ) = rsTemp!���ﵥλ
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.����ϵ��) = rsTemp!�����װ
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.סԺ��λ) = rsTemp!סԺ��λ
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.סԺϵ��) = rsTemp!סԺ��װ
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩ�ⵥλ) = rsTemp!ҩ�ⵥλ
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩ��ϵ��) = rsTemp!ҩ���װ
'
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.����) = IIf(rsTemp!ʱ�� = 1, "ʱ��", "����")
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ۼ�) = GetFormat(dbl����ϵ�� * rsTemp!�ּ�, mintPriceDigit)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ɱ���) = GetFormat(dbl����ϵ�� * rsTemp!�ɱ���, mintCostDigit)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.ָ������) = GetFormat(dbl����ϵ�� * rsTemp!ָ��������, mintCostDigit)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.ָ���ۼ�) = GetFormat(dbl����ϵ�� * rsTemp!ָ�����ۼ�, mintPriceDigit)
'
'            End If
'            rsTemp.MoveNext
'        Next
'        picDrug.Visible = False
'    End With
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long
    
    '����ҩƷ
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        txtFind.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����
        Else
            strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        End If
        
        For lngRow = 1 To vsfSelectDrug.rows - 1
            lngFindRow = vsfSelectDrug.FindRow(strҩ��, lngRow, CLng(vsfSelectDrugCol.ҩƷ��Ϣ), True, True)
            If lngFindRow > 0 Then
                vsfSelectDrug.Select lngFindRow, 1, lngFindRow, vsfSelectDrug.Cols - 1
                vsfSelectDrug.TopRow = lngFindRow
                Exit For
            End If
        Next
        
        If lngFindRow > 0 Then  '��ѯ�����ݺ���ƶ�����һ�����˳����β�ѯ
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtSelect_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub

Private Sub txtSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim rsPinzhong As ADODB.Recordset
    Dim objNode As Node
    Dim lng����id As Long
    Dim i As Integer
    
    If KeyCode = vbKeyReturn Then
    
        On Error GoTo errHandle
        
        If Trim(txtSelect.Text) = "" Then Exit Sub
                
        gstrSQL = "Select Distinct a.id,a.����,a.����" & _
                  "  From ������ĿĿ¼ A, ������Ŀ���� B" & _
                    " Where a.Id = b.������Ŀid(+) And a.��� In ('5', '6', '7') And Sysdate Between ����ʱ�� And ����ʱ�� And" & _
                         " (a.���� Like [1] Or a.���� Like [1] Or b.���� Like [1] Or b.���� Like [1])"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & UCase(txtSelect.Text) & mstrMatch)
        If rsTemp.RecordCount = 0 Then
            MsgBox "δ��ѯ��Ʒ�֣�", vbInformation, gstrSysName
            txtSelect.SetFocus
            txtSelect.SelStart = 1
            txtSelect.SelLength = Len(txtSelect.Text)
        Else
            picDrug.Visible = True
            vsfDrug.Visible = True
            Set vsfDrug.DataSource = rsTemp
            vsfDrug.SetFocus
            vsfDrug.Row = 1
        End If
        With vsfDrug
            For i = 0 To .rows - 1
                .RowHeight(i) = mlngRowHeight
            Next
        End With
'        gstrSQL = " Select ����,���� From ������Ŀ��� " & _
'                  " Where Instr([1],����,1) > 0 " & _
'                  " Order by ����"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, "567")
'
'        If rsTemp Is Nothing Then
'            Exit Sub
'        End If
'
'        With tvwDrug
'            .Nodes.Clear
'            Do While Not rsTemp.EOF
'                .Nodes.Add , , "Root" & rsTemp!����, rsTemp!����, 1, 1
'                .Nodes("Root" & rsTemp!����).Tag = rsTemp!����
'                rsTemp.MoveNext
'            Loop
'        End With
        
'        If optVariety.Value = True Then 'Ʒ�ֱ�ѡ��
'            gstrSQL = "Select a.Id, a.�ϼ�id, a.����, a.����, Decode(a.����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, '����' As ���" & _
'                        " From ���Ʒ���Ŀ¼ A," & _
'                             " (Select Distinct a.����id" & _
'                               " From ������ĿĿ¼ A, ������Ŀ���� B" & _
'                               " Where a.Id = b.������Ŀid(+) And a.��� In ('5', '6', '7') And Sysdate Between ����ʱ�� And ����ʱ�� And" & _
'                                     " (a.���� Like [1] Or a.���� Like [1] Or b.���� Like [1] Or b.���� Like [1])) B" & _
'                        " Where a.Id = b.����id And Nvl(To_Char(a.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                        " Start With a.�ϼ�id Is Null" & _
'                        " Connect By Prior a.Id = a.�ϼ�id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!�ϼ�id) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !����, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !�ϼ�id, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    End If
'                    objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'                    .MoveNext
'                Loop
'
'                rsTemp.MoveFirst
'                Do While Not rsTemp.EOF
'                    lng����id = rsTemp!Id
'                    gstrSQL = "Select Distinct a.Id, a.����id, a.����, a.����, Decode(a.���, '5', '����ҩ', '6', '�г�ҩ', '7', '�в�ҩ') ����, 'Ʒ��' As ���" & _
'                                " From ������ĿĿ¼ A" & _
'                                " Where a.��� In ('5', '6', '7') And a.����id=[1] and Sysdate Between a.����ʱ�� And a.����ʱ��"
'                    Set rsPinzhong = zlDatabase.OpenSQLRecord(gstrSQL, "Ʒ��", lng����id)
'
'                    Do While Not rsPinzhong.EOF
'                        Set objNode = tvwDrug.Nodes.Add("K_" & rsPinzhong!����id, 4, rsPinzhong!��� & "K_" & rsPinzhong!Id, rsPinzhong!���� & "-Ʒ��", 1, 1)
'                        objNode.Tag = rsPinzhong!���� & "-" & rsPinzhong!���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'                        rsPinzhong.MoveNext
'                    Loop
'                    rsTemp.MoveNext
'                Loop
'            End With
'        Else
'            gstrSQL = "Select ID, �ϼ�id, ����, ����, ����, ��� from ( " & _
'                    "Select Distinct ID, �ϼ�id, ����, ����, ����, ���" & _
'                        " From (Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, '����' As ���" & _
'                               " From ���Ʒ���Ŀ¼" & _
'                               " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
'                                     " (���� Like [1] Or ���� Like [1] Or ���� Like [1])" & _
'                               " Start With �ϼ�id Is Null" & _
'                               " Connect By Prior ID = �ϼ�id" & _
'                               " Union All" & _
'                               " Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, '����' As ���" & _
'                               " From ���Ʒ���Ŀ¼" & _
'                               " Where ID In (Select �ϼ�id" & _
'                                            " From ���Ʒ���Ŀ¼" & _
'                                            " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
'                                                  " (���� Like [1] Or ���� Like [1] Or ���� Like [1]))))" & _
'                        " Start With �ϼ�id Is Null" & _
'                        " Connect By Prior ID = �ϼ�id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!�ϼ�id) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !����, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !�ϼ�id, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    End If
'                    objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'                    .MoveNext
'                Loop
'            End With
'        End If
'        tvwDrug.SetFocus
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDrug_DblClick()
    Dim lngId As Long
    Dim rsTemp As ADODB.Recordset
    Dim intRow As Integer
    Dim i As Integer
    Dim blnDou As Boolean '�ظ�����
    Dim dbl����ϵ�� As Double
    Dim strUnit As String   '��λ

    On Error GoTo errHandle
    With vsfDrug
        If Val(.TextMatrix(.Row, 0)) = 0 Then
            Exit Sub
        End If
        gstrSQL = "Select Distinct a.ҩƷid, c.���� As ҩƷ����, c.���� As ͨ����, d.��Ʒ��, c.���, c.�Ƿ��� As ʱ��, c.����, c.���㵥λ As �ۼ۵�λ, a.���ﵥλ, a.�����װ," & _
                                    " a.סԺ��λ , a.סԺ��װ, a.ҩ�ⵥλ, a.ҩ���װ, a.�ɱ���, e.�ּ�, a.ָ��������, a.ָ�����ۼ�" & _
                    " From ҩƷ��� A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) D,�շѼ�Ŀ E" & _
                    " Where a.ҩ��id = b.Id And a.ҩƷid = c.Id And c.Id = d.�շ�ϸĿid(+) and a.ҩƷid=e.�շ�ϸĿid and sysdate between e.ִ������ and e.��ֹ���� and b.id=[1] " & _
                    " And (c.����ʱ�� = to_date('3000-01-01','yyyy-mm-dd') or c.����ʱ�� is null ) order by c.����"
        
        lngId = Val(.TextMatrix(.Row, 0))

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯҩƷ", lngId)
        If rsTemp.RecordCount = 0 Then
            Exit Sub
        End If
    End With

    With vsfSelectDrug
        For intRow = 0 To rsTemp.RecordCount - 1
            blnDou = False
            For i = 1 To .rows - 1
                If .TextMatrix(i, vsfSelectDrugCol.ҩƷID) = rsTemp!ҩƷID Then
                    blnDou = True
                End If
            Next
            If blnDou = False Then
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight
            
                Select Case mintUnit
                    Case 0
                        dbl����ϵ�� = rsTemp!ҩ���װ
                        strUnit = rsTemp!ҩ�ⵥλ
                    Case 1
                        dbl����ϵ�� = rsTemp!סԺ��װ
                        strUnit = rsTemp!סԺ��λ
                    Case 2
                        dbl����ϵ�� = rsTemp!�����װ
                        strUnit = rsTemp!���ﵥλ
                    Case 3
                        dbl����ϵ�� = 1
                        strUnit = rsTemp!�ۼ۵�λ
                End Select
                                
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷID) = rsTemp!ҩƷID
                If gintҩƷ������ʾ = 1 Then
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ��Ϣ) = "[" & rsTemp!ҩƷ���� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
                Else
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ��Ϣ) = "[" & rsTemp!ҩƷ���� & "]" & rsTemp!ͨ����
                End If

                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ����) = rsTemp!ҩƷ����
                .TextMatrix(.rows - 1, vsfSelectDrugCol.��Ʒ��) = IIf(IsNull(rsTemp!��Ʒ��), "", rsTemp!��Ʒ��)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ͨ����) = IIf(IsNull(rsTemp!ͨ����), "", rsTemp!ͨ����)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.��λ) = strUnit
                
                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ۼ۵�λ) = rsTemp!�ۼ۵�λ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.���ﵥλ) = rsTemp!���ﵥλ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.����ϵ��) = rsTemp!�����װ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.סԺ��λ) = rsTemp!סԺ��λ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.סԺϵ��) = rsTemp!סԺ��װ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩ�ⵥλ) = rsTemp!ҩ�ⵥλ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩ��ϵ��) = rsTemp!ҩ���װ
                
                
                .TextMatrix(.rows - 1, vsfSelectDrugCol.����) = IIf(rsTemp!ʱ�� = 1, "ʱ��", "����")
                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ۼ�) = GetFormat(dbl����ϵ�� * rsTemp!�ּ�, mintPriceDigit)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ɱ���) = GetFormat(dbl����ϵ�� * rsTemp!�ɱ���, mintCostDigit)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ָ������) = GetFormat(dbl����ϵ�� * rsTemp!ָ��������, mintCostDigit)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ָ���ۼ�) = GetFormat(dbl����ϵ�� * rsTemp!ָ�����ۼ�, mintPriceDigit)
                
            End If
            rsTemp.MoveNext
        Next
        picDrug.Visible = False
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfDrug_DblClick
    End If
End Sub

Private Sub vsfSelectDrug_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub
