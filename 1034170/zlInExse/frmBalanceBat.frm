VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBalanceBat 
   Caption         =   "������;����"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBalanceBat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11820
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   45
      ScaleHeight     =   1125
      ScaleWidth      =   11730
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5640
      Width           =   11730
      Begin VB.CommandButton cmdOK 
         Caption         =   "����(&O)"
         Default         =   -1  'True
         Height          =   400
         Left            =   8640
         TabIndex        =   15
         Top             =   705
         Width           =   1400
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&C)"
         Height          =   400
         Left            =   10200
         TabIndex        =   16
         Top             =   705
         Width           =   1400
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   360
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9675
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1905
      End
      Begin MSMask.MaskEdBox txtDateEnd 
         Height          =   360
         Left            =   280
         TabIndex        =   7
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDateEnd 
         Caption         =   "��                     ֮ǰ�ķ��ý���"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   60
         Width           =   4440
      End
      Begin VB.Label lblDeposit 
         Caption         =   "��Ԥ���ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblPay 
         Caption         =   "XX����ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lbl���㷽ʽ 
         Caption         =   "���㷽ʽ"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
         Height          =   240
         Left            =   8880
         TabIndex        =   10
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Caption         =   "�����n�����˽���"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   480
         Width           =   8295
      End
   End
   Begin VB.Frame fra 
      Height          =   645
      Left            =   90
      TabIndex        =   17
      Top             =   0
      Width           =   11685
      Begin VB.ComboBox cboʹ����� 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   2490
      End
      Begin VB.Label lblRpt 
         AutoSize        =   -1  'True
         Caption         =   "sss"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3810
         TabIndex        =   2
         Top             =   255
         Width           =   405
      End
      Begin VB.Label lblʹ����� 
         Caption         =   "ʹ�����"
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   4860
      Left            =   2160
      TabIndex        =   4
      Top             =   675
      Width           =   2460
      _cx             =   4339
      _cy             =   8572
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
      BackColorSel    =   13627390
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":617A
      ScrollTrack     =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   4860
      Left            =   4680
      TabIndex        =   5
      Top             =   690
      Width           =   7065
      _cx             =   12462
      _cy             =   8572
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
      BackColorSel    =   12640511
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":61C2
      ScrollTrack     =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsFeeType 
      Height          =   4875
      Left            =   120
      TabIndex        =   3
      Top             =   675
      Width           =   1980
      _cx             =   3492
      _cy             =   8599
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
      BackColorSel    =   15790320
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":62A7
      ScrollTrack     =   -1  'True
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
Attribute VB_Name = "frmBalanceBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPatis As String '���ڼ�¼ѡ��Ŀ����±��Ϊ�����ʵĲ���ID
Private mlng����ID As Long
Private mrsRptFormat As ADODB.Recordset
Private mlngShareUseID As Long     '��������
Private mstrUseType As String          'ʹ�����
Private mintPrintMode As Integer    '��ӡ��ʽ:0-����ӡ;1-��ʾ��ӡ;2-�Զ���ӡ
Private mintPrintFormat As Integer '��ӡ��ʽ
  
Private Sub cboʹ�����_Click()
    lblRpt.Caption = ""
    mstrUseType = cboʹ�����.Text
    If mrsRptFormat Is Nothing Then Exit Sub
    mrsRptFormat.Filter = "���=" & cboʹ�����.ItemData(cboʹ�����.ListIndex)
    If Not mrsRptFormat.EOF Then
        lblRpt.Caption = Nvl(mrsRptFormat!˵��)
    End If
    mlng����ID = 0
    mlngShareUseID = zl_GetInvoiceShareID(1137, mstrUseType)    '��������
    mintPrintMode = zl_GetInvoicePrintMode(1137, mstrUseType)  '��ӡ��ʽ:0-����ӡ;1-��ʾ��ӡ;2-�Զ���ӡ
    mintPrintFormat = zl_GetInvoicePrintFormat(1137, mstrUseType)     '��ӡ��ʽ
    Call RefreshFact
    
    Call vsDept_AfterRowColChange(0, 0, vsDept.Row, vsDept.Col)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, m As Long, blnPrint As Boolean
    Dim rsPati As ADODB.Recordset
    
    For i = 1 To vsDept.Rows - 1
        If vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked Then
            m = m + 1
        End If
    Next
    If m = vsDept.Rows - 1 Then
        MsgBox "������ѡ��һ������.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set rsPati = GetPatiSet
    If rsPati.RecordCount = 0 Then
        MsgBox "������ѡ��һ������.", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not IsDate(txtDateEnd.Text) Then
        MsgBox "���ý�ֹʱ���ʽ����ȷ.", vbInformation, gstrSysName
        txtDateEnd.SetFocus
        Exit Sub
    End If
    
    blnPrint = mintPrintMode <> 0
    If mintPrintMode = 2 Then
        If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
    End If
    
    If blnPrint Then
        If gblnStrictCtrl Then   '�ϸ�Ʊ�ݹ���
            If Trim(txtInvoice.Text) = "" Then
                MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
            mlng����ID = GetInvoiceGroupID(IIf(gbytInvoiceKind = 0, 3, 1), rsPati.RecordCount, mlng����ID, mlngShareUseID, txtInvoice.Text, mstrUseType)
            If mlng����ID <= 0 Then
                Select Case mlng����ID
                    Case 0 '����ʧ��
                    Case -1
                        MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -2
                        MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -3
                        MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,����������", vbInformation, gstrSysName
                        txtInvoice.SetFocus
                End Select
                Exit Sub
            End If
        Else
            If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If MsgBox("��ѡ����" & rsPati.RecordCount & "λ����,�������ν�����;����!" & _
        vbCrLf & "��׼���ú�ȷ��.", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
        Exit Sub
    End If
    
        
    cmdOK.Enabled = False
    Screen.MousePointer = 11
    
    Call SaveBalance(blnPrint, rsPati)
    
    Call LoadPati(Val(vsDept.RowData(vsDept.Row)))
    Screen.MousePointer = 0
    cmdOK.Enabled = True
    
    gblnOK = True
End Sub

Private Sub GetMaxMinDate(ByVal lngPatiID As Long, ByVal strDateMode As String, ByVal DatEnd As Date, ByRef DatMax As Date, ByRef DatMin As Date)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTable As String
    
    'Ҫ�͹���Zl_���ʷ��ü�¼_Patient�еĴ�������α�һ��,�������û�н��ʷ��õĽ��ʵ�.
    '�����:������������ʵ,��SQL����������,�����������ķ��ý��н���,��ʵ����Ӧ��ֻ���סԺ����,���,���β��ֻ�滻��סԺ���ü�¼
    
    strSql = "" & _
    " Select Max(Maxʱ��) DatMax, Min(Minʱ��) DatMin" & vbNewLine & _
    " From ( Select Max(" & strDateMode & ") Maxʱ��, Min(" & strDateMode & ") Minʱ��" & vbNewLine & _
    "        From סԺ���ü�¼ A" & vbNewLine & _
    "        Where A.����id = [1] And A.����id Is Null And A.��¼״̬ <> 0 And Mod(��¼����, 10) In (2, 3) And" & vbNewLine & _
    "             " & strDateMode & " < [2] " & vbCrLf & _
    "             And Not Exists ( Select 1" & vbNewLine & _
    "                              From סԺ���ü�¼ B" & vbNewLine & _
    "                              Where B.NO = A.NO And B.��¼���� = A.��¼���� And B.��� = A.���" & vbNewLine & _
    "                              Group By B.NO, B.��¼����, B.���" & vbNewLine & _
    "                              Having Nvl(Sum(B.ʵ�ս��), 0) = Decode(" & IIf(gblnZero, 1, 0) & ", 1, 1 + Nvl(Sum(B.ʵ�ս��), 0), 0))" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select Max(" & strDateMode & ") Maxʱ��, Min(" & strDateMode & ") Minʱ��" & vbNewLine & _
    "       From " & zlGetFullFieldsTable("סԺ���ü�¼") & vbNewLine & _
    "       Where A.����id = [1] And A.����id Is Not Null And Mod(��¼����, 10) In (2, 3) And Nvl(A.ʵ�ս��, 0) <> Nvl(A.���ʽ��, 0) And" & vbNewLine & _
    "             " & strDateMode & " < [2]" & vbNewLine & _
    "       Group By A.NO, A.���, Mod(A.��¼����, 10), A.��¼״̬, A.ִ��״̬" & vbNewLine & _
    "       Having Nvl(Sum(A.ʵ�ս��), 0) - Nvl(Sum(A.���ʽ��), 0) <> 0)"


    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatiID, DatEnd)
    DatMax = Nvl(rsTmp!DatMax, CDate(0))
    DatMin = Nvl(rsTmp!DatMin, CDate(0))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDateStr(DatTmp As Date) As String
    GetDateStr = "To_Date('" & Format(DatTmp, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Function GetBalanceSum(ByVal Dat�տ�ʱ�� As Date) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select B.���㷽ʽ, Sum(B.��Ԥ��) ������" & vbNewLine & _
            "From ���˽��ʼ�¼ A, ����Ԥ����¼ B" & vbNewLine & _
            "Where A.�շ�ʱ�� = [1] And A.����Ա���� = [2] And A.ID = B.����id" & vbNewLine & _
            "Group By B.���㷽ʽ"

    On Error GoTo errH
    Set GetBalanceSum = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Dat�տ�ʱ��, UserInfo.����)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SaveBalance(ByRef blnPrint As Boolean, ByRef rsPati As ADODB.Recordset)
    Dim strNO As String, lng����ID As Long, datBalance As Date, lngPatientID As Long, i As Long, j As Long
    Dim arrSQL As Variant, DatMax As Date, DatMin As Date, lngNum As Long, blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    
    lblPay.Visible = False
    lblDeposit.Visible = False
    Err = 0: On Error GoTo Errhand:
    datBalance = zlDatabase.Currentdate '��¼Ϊͳһ�Ľ���ʱ��

    For i = 1 To rsPati.RecordCount
        arrSQL = Array()
        lngPatientID = rsPati!����ID
        Call GetMaxMinDate(lngPatientID, IIf(gint����ʱ�� = 1, "����ʱ��", "�Ǽ�ʱ��"), CDate(txtDateEnd.Text), DatMax, DatMin)
        
        If Not (DatMax = DatMin And DatMax = CDate(0)) Then 'û�д�����ò�����
            lblInfo.Caption = "��ǰ����:��" & rsPati.RecordCount & "λ,���ڽ��е�" & i & "λ," & rsPati!���� & ":" & rsPati!����
            Me.Refresh
            
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            strNO = zlDatabase.GetNextNo(15)
    
            '1.���˽��ʼ�¼
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            '58758
            arrSQL(UBound(arrSQL)) = "zl_���˽��ʼ�¼_Insert(" & lng����ID & "," & "'" & strNO & "'," & lngPatientID & "," & _
                GetDateStr(datBalance) & "," & GetDateStr(DatMin) & "," & GetDateStr(DatMax) & ",1,0,0,2,NULL,2)"
            
            '2.���ʽɿ��¼:zl_����Ԥ����¼_Insert,zl_���ʽɿ��¼_Insert��Zl_���ʷ��ü�¼_Patient�е���,��Ϊ����������δ֪
            '3.סԺ���ü�¼
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_���ʷ��ü�¼_Patient('" & strNO & "','" & lngPatientID & "'," & lng����ID & "," & _
                GetDateStr(CDate(txtDateEnd.Text)) & "," & gint����ʱ�� & "," & IIf(gblnZero, 1, 0) & _
                ",'" & cbo���㷽ʽ.Text & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & GetDateStr(datBalance) & ")"
                    
            '4.��ʼƱ�ݺ�
            If blnPrint And Trim(txtInvoice.Text) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_Ʊ����ʼ��_Update('" & strNO & "','" & Trim(txtInvoice.Text) & "',3)"
            End If
        
            
            On Error GoTo errH
            gcnOracle.BeginTrans: blnTrans = True
                For j = 0 To UBound(arrSQL)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSQL(j)), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            lngNum = lngNum + 1 '��¼ʵ�ʽ�������
            
            'Ʊ�ݴ�ӡ
            If blnPrint Then
                Call frmPrint.ReportPrint(1, strNO, lng����ID, mlng����ID, mlngShareUseID, mstrUseType, txtInvoice.Text, datBalance, "", "", lngPatientID, mintPrintFormat)
                Call RefreshFact
            End If
        End If
        
        rsPati.MoveNext
    Next
        
    If lngNum = 0 Then
        lblInfo.Caption = "ѡ����" & rsPati.RecordCount & "λ����,����ָ���Ľ�ֹʱ��ǰ��������δ�����!"
    Else
        lblInfo.Caption = "��" & rsPati.RecordCount & "λ������,����δ����õ�" & lngNum & "λ�������;����."
        
        Set rsTmp = GetBalanceSum(datBalance)
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "���㷽ʽ<>'" & cbo���㷽ʽ.Text & "'"
            If rsTmp.RecordCount > 0 Then
                lblDeposit.Caption = "��Ԥ���ϼƣ�" & Format(rsTmp!������, "0.00")
                lblDeposit.Visible = True
            End If
            rsTmp.Filter = "���㷽ʽ='" & cbo���㷽ʽ.Text & "'"
            If rsTmp.RecordCount > 0 Then
                If lblDeposit.Visible = False Then
                    lblPay.Left = lblDeposit.Left
                Else
                    lblPay.Left = lblDeposit.Left + lblDeposit.Width + 200
                End If
                lblPay.Caption = cbo���㷽ʽ.Text & "����ϼƣ�" & Format(rsTmp!������, "0.00")
                lblPay.Visible = True
            End If
        End If
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    If lngNum > 0 Then
        lblInfo.Caption = "ѡ����" & rsPati.RecordCount & "λ����,ʵ�ʶ�" & lngNum & "λ�����������;����."
    End If
    Exit Sub
Errhand:
     If ErrCenter = 1 Then Resume
End Sub
Private Sub RefreshFact()
'���ܣ�ˢ���շ�Ʊ�ݺ�
    If mintPrintMode = 0 Then Exit Sub
    If gblnStrictCtrl Then
        mlng����ID = CheckUsedBill(IIf(gbytInvoiceKind = 0, 3, 1), IIf(mlng����ID > 0, mlng����ID, mlngShareUseID), , mstrUseType)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            txtInvoice.Text = ""
        Else
            '�ϸ�ȡ��һ������
            txtInvoice.Text = GetNextBill(mlng����ID)
        End If
    Else
        '��ɢ��ȡ��һ������
        txtInvoice.Text = IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
    End If
End Sub
Private Function GetPatiSet() As ADODB.Recordset
    Dim strSql As String, str�ѱ� As String, strDeptIDs As String, i As Long
    
    str�ѱ� = Get�ѱ�ѡ��
    If str�ѱ� <> "" Then
        If UBound(Split(str�ѱ�, ",")) + 1 < vsFeeType.Rows - 1 Then
            str�ѱ� = "," & str�ѱ� & ","
            strSql = " And Instr([2],','||A.�ѱ�||',') > 0"
        End If
    End If
    
    For i = 1 To vsDept.Rows - 1
        If Not (vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked) Then
            strDeptIDs = strDeptIDs & "," & vsDept.RowData(i)
        End If
    Next
    strDeptIDs = Mid(strDeptIDs, 2)
    If UBound(Split(strDeptIDs, ",")) + 1 = vsDept.Rows - 1 Then strDeptIDs = ""
    
    If strDeptIDs <> "" Then
        strSql = strSql & " And B.����ID In(" & strDeptIDs & ")"
    End If
    
    If mstrPatis <> "" Then
        mstrPatis = "," & mstrPatis & ","
        strSql = strSql & " And Instr([1],','||B.����ID||',') = 0"
    End If
    
    strSql = "Select Distinct C.���� as ����,A.����,A.����ID,A.סԺ����,A.סԺ��" & vbNewLine & _
            "From ������Ϣ A, ��λ״����¼ B, ���ű� C,������ҳ M " & vbNewLine & _
            "Where A.����id = B.����ID  " & _
            "       And B.����ID = C.ID And A.���� is Null" & _
            "       And Zl_Billclass(A.����ID,A.��ҳID,0)=[3]  " & strSql & vbNewLine & _
            "Order by ����,סԺ��"

    On Error GoTo errH
    Set GetPatiSet = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrPatis, str�ѱ�, Trim(cboʹ�����.Text))

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Load()
    Set mrsRptFormat = Nothing
    lblInfo.Caption = ""
    Call LoadUseType
    If Not InitData Then
        Unload Me
    End If
    If vsDept.Rows > 1 Then
        vsDept.Row = 0
        vsDept.Row = 1
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset, i As Long

    Set rsTmp = Get�ѱ�
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ѱ�δ����,����ʹ�ô˹���!", vbInformation, gstrSysName
        Exit Function
    Else
        vsFeeType.Rows = rsTmp.RecordCount + 1
        vsFeeType.ColDataType(0) = flexDTBoolean
        vsFeeType.Cell(flexcpChecked, 1, 0, vsFeeType.Rows - 1, 0) = flexChecked
        vsFeeType.Row = 1: vsFeeType.Col = 1: vsFeeType.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsFeeType.TextMatrix(i, 1) = rsTmp!����
        rsTmp.MoveNext
    Next
    Call LoadDept
    
    txtDateEnd.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    Set rsTmp = Get���㷽ʽ("����", 2)
    If rsTmp.RecordCount = 0 Then
        MsgBox "û���������ڽ��ʳ��ϵķ��ֽ���㷽ʽ,����ʹ�ô˹���!", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 1 To rsTmp.RecordCount
        cbo���㷽ʽ.AddItem rsTmp!����
        rsTmp.MoveNext
    Next
    cbo���㷽ʽ.ListIndex = 0
    
    Call RefreshFact
    
    InitData = True
End Function

Private Function Get�ѱ�() As ADODB.Recordset
    Dim strSql As String
 
    strSql = "Select ����,���� From �ѱ� Where ������� In (2, 3) And ���� = 1 Order by ����"
    On Error GoTo errH
    Set Get�ѱ� = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadDept()
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "Select A.ID, A.����" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where A.ID = B.����id And B.������� In (2, 3) And B.�������� = '�ٴ�'" & vbNewLine & _
            " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & vbNewLine & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " And Exists" & vbNewLine & _
            " (Select 1 From ��λ״����¼ C Where C.����id Is Not Null And C.����id = A.ID) Order by ����"
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    vsDept.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
        vsDept.Row = 1: vsDept.Col = 1: vsDept.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsDept.TextMatrix(i, 1) = rsTmp!����
        vsDept.RowData(i) = Val(rsTmp!ID)
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get�ѱ�ѡ��() As String
    Dim i As Long, strTmp As String
    
    For i = 1 To vsFeeType.Rows - 1
        If vsFeeType.Cell(flexcpChecked, i, 0) = flexChecked Then strTmp = strTmp & "," & vsFeeType.TextMatrix(i, 1)
    Next
    Get�ѱ�ѡ�� = Mid(strTmp, 2)
End Function

Private Sub LoadPati(ByVal lngDeptID As Long)
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long, str�ѱ� As String
    
    str�ѱ� = Get�ѱ�ѡ��
    If str�ѱ� <> "" Then
        If UBound(Split(str�ѱ�, ",")) + 1 < vsFeeType.Rows - 1 Then
            str�ѱ� = "," & str�ѱ� & ","
            strSql = " And Instr([2],','||A.�ѱ�||',')>0"
        End If
    End If
    
    strSql = "" & _
            "   Select Distinct A.����ID,A.סԺ��, Nvl(D.����,A.����) as ����, Nvl(D.�Ա�,A.�Ա�) as �Ա�, " & _
            "               Nvl(D.����,A.����) as ����, B.������� δ�����, Ԥ����� ����Ԥ��, A.�ѱ�" & vbNewLine & _
            "   From ������Ϣ A, ������� B,��λ״����¼ C,������ҳ D " & vbNewLine & _
            "   Where C.����id = [1] And C.����ID = A.����id And A.����id = B.����id(+) " & _
            "               And A.����id=D.����ID(+) And A.��ҳid = D.��ҳid(+) " & _
            "               And B.����(+) = 1  And B.����(+)=2 And A.���� is Null " & _
            "               And Zl_Billclass(A.����ID,A.��ҳID,0)=[3] " & strSql & vbNewLine & _
            "   Order by A.סԺ��"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngDeptID, str�ѱ�, Trim(cboʹ�����.Text))
    vsPati.Rows = 1 '�������,��������б�ͷ
    vsPati.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
        Else
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
        End If
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    End If
    
    With vsPati
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 1) = "" & rsTmp!סԺ��
            .TextMatrix(i, 2) = "" & rsTmp!����
            .TextMatrix(i, 3) = "" & rsTmp!�Ա�
            .TextMatrix(i, 4) = "" & rsTmp!����
            .TextMatrix(i, 5) = Nvl(rsTmp!δ�����, ""): If Val(.TextMatrix(i, 5)) = 0 Then .TextMatrix(i, 5) = ""
            .TextMatrix(i, 6) = Nvl(rsTmp!����Ԥ��, ""): If Val(.TextMatrix(i, 6)) = 0 Then .TextMatrix(i, 6) = ""
            .TextMatrix(i, 7) = "" & rsTmp!�ѱ�
            .RowData(i) = Val(rsTmp!����ID)
            If Len(mstrPatis) > 0 Then
                If InStr("," & mstrPatis & ",", "," & rsTmp!����ID & ",") > 0 Then
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            End If
            rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then .Row = 1: .Col = 1: .Col = 0
    End With
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.Width < 12060 Then Me.Width = 12060
    If Me.Height < 7635 Then Me.Height = 7635
    With fra
        .Width = ScaleWidth - .Left * 2
    End With
    With picDown
        .Width = ScaleWidth
        .Top = ScaleHeight - .Height - 100
    End With
     With vsFeeType
        .Height = picDown.Top - .Top - 50
        vsDept.Height = .Height
        vsPati.Height = .Height
        vsPati.Width = ScaleWidth - vsPati.Left - 50
     End With
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRptFormat = Nothing
    mstrPatis = ""
    mlng����ID = 0
End Sub

Private Sub picDown_Resize()
  Err = 0: On Error Resume Next
    With cmdCancel
        .Left = picDown.ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = .Left - cmdOK.Width - 50
    End With
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed Then vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked    '�ֶ����ʱ��Ϊ��ɫ����Ϊѡ��
    
    If Row <> vsDept.Row Then vsDept.Row = Row
    If vsPati.Rows < 2 Then Exit Sub
    
    If vsDept.Cell(flexcpChecked, Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, Row, 0) = flexTSUnchecked Then
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
    Else
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
    End If
    Call SetPatiLists
End Sub
Private Sub vsdept_DblClick()
    If vsDept.MouseCol = 0 And vsDept.MouseRow = 0 Then
        Call SetVSAll(vsDept)
        Call vsDept_AfterEdit(vsDept.Row, vsDept.Col)
        mstrPatis = ""
    End If
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow And NewRow <> 0 Then Call LoadPati(Val(vsDept.RowData(NewRow)))
End Sub



Private Sub vsFeeType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    i = vsDept.Row
    vsDept.Row = 0
    vsDept.Row = i
    
End Sub

Private Sub vsPati_DblClick()
    If vsPati.MouseCol = 0 And vsPati.MouseRow = 0 Then
        If vsPati.Rows < 2 Then Exit Sub
        
        Call SetVSAll(vsPati)
        Call SetDeptState
        Call SetPatiLists
    End If
End Sub

Private Sub vsFeeType_DblClick()
    Dim i As Long
    If vsFeeType.MouseCol = 0 And vsFeeType.MouseRow = 0 Then
        Call SetVSAll(vsFeeType)
        i = vsDept.Row
        vsDept.Row = 0
        vsDept.Row = i
    End If
End Sub

Private Sub SetVSAll(ByRef vsf As VSFlexGrid)
    If vsf.Rows < 2 Then Exit Sub
    vsf.Cell(flexcpChecked, 1, 0, vsf.Rows - 1, 0) = IIf(Val(vsf.Tag) = 1, flexChecked, flexUnchecked)
    vsf.Tag = IIf(Val(vsf.Tag) = 0, 1, 0)
End Sub


Private Sub vsPati_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
        SetPatiLists
    Else
        Call SetPatistr(Row)
    End If
    Call SetDeptState
End Sub

Private Sub SetPatistr(ByVal lngRow As Long)
'���ܣ���¼û��ѡ��Ĳ��ˣɣ�
    If vsPati.Cell(flexcpChecked, lngRow, 0) = flexUnchecked Then
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") = 0 Then
            If mstrPatis = "" Then
                mstrPatis = vsPati.RowData(lngRow)
            Else
                mstrPatis = mstrPatis & "," & vsPati.RowData(lngRow)
            End If
        End If
    Else
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") > 0 Then
            mstrPatis = Replace("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",", ",")
            mstrPatis = Mid(mstrPatis, 2)   'ȥ��ǰ���
            If mstrPatis <> "" Then mstrPatis = Mid(mstrPatis, 1, Len(mstrPatis) - 1)
        End If
    End If
    If mstrPatis = "," Then mstrPatis = ""
End Sub

Private Sub SetPatiLists()
'����:��鵱ǰ�����б���û��ѡ��ļ��뵽�����У���ѡ��ģ��ӱ�����ɾ��
    Dim i As Long
    
    If vsPati.Rows < 2 Then Exit Sub
    
    For i = 1 To vsPati.Rows - 1
        Call SetPatistr(i)
    Next
End Sub

Private Function SetDeptState() As Boolean
'���ܣ����ÿ���ѡ��״̬
    Dim i As Long, m As Long
    
    For i = 1 To vsPati.Rows - 1
        If vsPati.Cell(flexcpChecked, i, 0) = flexChecked Then m = m + 1
    Next
    If m = vsPati.Rows - 1 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked
    ElseIf m = 0 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed
    End If
End Function

Private Sub vspati_EnterCell()
    If vsPati.Col = 0 Then
        vsPati.Editable = flexEDKbdMouse
    Else
        vsPati.Editable = flexEDNone
    End If
End Sub
Private Sub vsfeetype_EnterCell()
    If vsFeeType.Col = 0 Then
        vsFeeType.Editable = flexEDKbdMouse
    Else
        vsFeeType.Editable = flexEDNone
    End If
End Sub
Private Sub vsDept_EnterCell()
    If vsDept.Col = 0 Then
        vsDept.Editable = flexEDKbdMouse
    Else
        vsDept.Editable = flexEDNone
    End If
End Sub
Private Sub LoadUseType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʹ�����
    '����:���˺�
    '����:2011-04-28 15:09:10
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, strSql As String
    Dim varData As Variant, varTemp As Variant
    Dim strRptName As String
    Dim strShareInvoice As String
    
    On Error GoTo errHandle
    
    strShareInvoice = zlDatabase.GetPara("���ʷ�Ʊ��ʽ", glngSys, 1137)
    varData = Split(strShareInvoice, "|")
    
    strRptName = IIf(gbytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    
    'Ʊ�ݸ�ʽ����
    strSql = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set mrsRptFormat = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    
    strSql = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cboʹ�����
        .Clear
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!����)
            .ItemData(.NewIndex) = 0
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .ItemData(.NewIndex) = Val(varTemp(1))
                    Exit For
                End If
            Next
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        mstrUseType = cboʹ�����.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

