VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMedicareReckoning 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ҽ�����˽���У��"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdReturnCash 
      Caption         =   "����"
      Height          =   330
      Left            =   6330
      TabIndex        =   19
      Top             =   255
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ �� (&C)"
      Height          =   420
      Left            =   8160
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2745
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txt�Ҳ� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   5040
      Width           =   1755
   End
   Begin VB.TextBox txtMargin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
      Width           =   1755
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   9885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ ��(&O)"
      Height          =   420
      Left            =   6465
      TabIndex        =   6
      Top             =   5055
      Width           =   1395
   End
   Begin VB.TextBox txt�ɿ� 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   5025
      Width           =   1755
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
      Height          =   3420
      Left            =   5280
      TabIndex        =   17
      Top             =   855
      Width           =   4230
      _cx             =   7461
      _cy             =   6032
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
      Height          =   3420
      Left            =   30
      TabIndex        =   2
      Top             =   870
      Width           =   5205
      _cx             =   9181
      _cy             =   6032
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VB.Label lblDelMoney 
      Caption         =   "��֧����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7005
      TabIndex        =   18
      Top             =   255
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Label lblӦ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7680
      TabIndex        =   15
      Tag             =   "Ӧ��:"
      Top             =   4410
      Width           =   600
   End
   Begin VB.Label lblҽ��֧�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��֧��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5280
      TabIndex        =   14
      Tag             =   "ҽ��֧��:"
      Top             =   4410
      Width           =   1080
   End
   Begin VB.Label lbl��Ԥ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ԥ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2280
      TabIndex        =   13
      Tag             =   "��Ԥ��:"
      Top             =   4410
      Width           =   840
   End
   Begin VB.Label lblԤ����� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ԥ�����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Tag             =   "Ԥ�����:"
      Top             =   4410
      Width           =   1080
   End
   Begin VB.Label lbl�ɿ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ�ɿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label lbl�Ҳ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ��Ҳ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   10
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   960
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ʽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frmMedicareReckoning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mbytInFun As Byte '0-����ģ�����,1-ҽ��ģ�����
Private mlngModul As Long

Private mlng����ID As Long
Private mlng����ID As Long, mlng��ҳID As Long
Private mbln��;���� As Boolean     '��Ժ����,δ�����Ԥ�����Ҫ��Ϊ�ֽ�
Private mstr���ս��� As String
Private mstr������Ϣ As String      '�������,��������,�����ʺ�
Private mbln������� As Boolean
Private mcur���ʽ�� As Currency
Private mcurԤ����� As Currency
Private mintInsure As Integer       '�����ж��Ƿ�֧�ֱַҴ���
Private mcur�ɿ� As Currency
Private mcurӦ�ɽ�� As Currency
Private mstrҽ���� As String
Private mcur����͸֧ As Currency
Private mintError As Integer
Private mstrStyle As String
Private mblnThreeDepositAfter As Boolean
Private mcur�շ���� As Currency
Private mblnOk  As Boolean
Private mstrPrivs As String
Private mintDefault As Integer      'ȱʡ���㷽ʽ��(Ϊ0��ʾû��)
Private mcurMediCare   As Currency  'ҽ������ϼ�,����[mstr���ս���]����
Private mblnClickOK As Boolean      '����ֻ�����ȷ���˳�
Private mblnCent As Boolean         'ҽ���Ƿ�֧�ֱַҴ���
Private mcur������� As Currency
Private mstrForceNote As String, mstrCardPrivs As String
Private mcur��Ԥ���ϼ� As Currency
Private mcurԤ���ϼ� As Currency
Private mstrסԺ���� As String  'סԺ����:����ö��ŷ���
Private mintԤ����� As Integer
Private mobjCard As Card
Private mbytInvoiceKind As Byte
Private Type TY_BrushCard    'ˢ������
    str���� As String
    str���� As String
    str������ˮ�� As String    '������ˮ��
    str����˵��  As String     '������Ϣ
    str��չ��Ϣ As String    '���׵���չ��Ϣ
    dbl�ʻ���� As Double
    dblMoney As Double     '��ǰ�˿��ˢ�����
End Type
Private mCurBrushCard As TY_BrushCard   '��ǰ��ˢ����Ϣ


'ģ�������˽�л�
Private Const support�ֱҴ��� = 25  'ҽ�������Ƿ���ֱ�   ,��Ҫ��Ϊ�˱���ҽ����ҽԺ����
Private mstrDec As String
Private mBytMoney As Byte '�շѷֱҴ�����
Private mbytMCMode As Byte 'ҽ���������֤��ģʽ,����1-����,2-סԺ����ģʽ,0-��ʾ��ҽ��
Private mbytMzDeposit As Byte '����Ԥ��ȱʡʹ�÷�ʽ:0-ȱʡ��ʹ�ý�;1-�����ʽ��ʹ��Ԥ��;2-ʹ������Ԥ��
Private mblnFirst As Boolean
Private mrsCardType As ADODB.Recordset 'ҽ�ƿ����
Private mobjPayCards As Cards
Private mblnExternal As Boolean, mstrNO As String
Private mFactProperty As Ty_FactProperty, mstrInvoice As String
Private mlng����ID As Long, mstrUseType As String, mlngShareUseID As Long
Private mintInvoiceFormat As Integer, mintInvoiceMode As Integer, mint�������� As Integer


Private Sub InitBalanceGrid(ByRef vsGrid As VSFlexGrid)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������ͷ��Ϣ
    '����:���˺�
    '����:2015-05-04 17:33:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsGrid
        .Redraw = flexRDNone
        .Clear
        .Rows = 2: .Cols = 6: i = 0
        .TextMatrix(0, i) = "���㷽ʽ": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        .TextMatrix(0, i) = "�������": i = i + 1
        .TextMatrix(0, i) = "����": i = i + 1
        .TextMatrix(0, i) = "ȱʡ": i = i + 1
        .TextMatrix(0, i) = "�����ID": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "���" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            If .ColKey(i) = "����" Or .ColKey(i) = "ȱʡ" Or .ColKey(i) = "�����ID" Then
                .ColHidden(i) = True: .ColWidth(0) = 0
            End If
        Next
        .ColWidth(.ColIndex("���㷽ʽ")) = 1200
        .ColWidth(.ColIndex("���")) = 1100
        .ColWidth(.ColIndex("�������")) = 1450
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadBalance() As Boolean
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim i As Long, str���㷽ʽ As String, blnFind As Boolean
    
    strSql = "" & _
    "   Select B.����,A.���㷽ʽ,A.��Ԥ��,A.�������,�����ID,A.У�Ա�־" & _
    "   From ����Ԥ����¼ A, ���㷽ʽ B " & _
    "   Where A.���㷽ʽ=B.����(+) And A.����ID=[1] And Mod(A.��¼����,10)<>1  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
    If rsTemp.EOF Then LoadBalance = True: Exit Function
    If InStr(1, mstrPrivs, ";����Ԥ������;") > 0 Then
        rsTemp.Filter = "���� <> 3 And ���� <> 4"
        If rsTemp.RecordCount <> 0 Then
            MsgBox "����Ԥ������ʱ,���ܴ���Ԥ����ҽ��֮��Ľ��㷽ʽ�Ľ��ʵ���!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With vsfMoney
        Do While Not rsTemp.EOF
            str���㷽ʽ = Nvl(rsTemp!���㷽ʽ): blnFind = False
            For i = 1 To .Rows - 1
                If str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))) Then
                    blnFind = True
                    If i = mintDefault Then Exit For    'ͨ������õ�
                    '�Ѿ�����ʽ��������ֵ
                    If InStr(",3,4,", "," & .TextMatrix(i, .ColIndex("����")) & ",") > 0 Then Exit For
                    .TextMatrix(i, .ColIndex("���")) = Format(Val(Nvl(rsTemp!��Ԥ��)), "0.00")
                    .TextMatrix(i, .ColIndex("�������")) = Nvl(rsTemp!�������)
                    .TextMatrix(i, .ColIndex("�����ID")) = Nvl(rsTemp!�����ID)
                    Exit For
                End If
            Next
            rsTemp.MoveNext
        Loop
    End With
    Call ShowMoney(False)
    LoadBalance = True
End Function

Public Function ShowMeFromOut(ByRef frmParent As Object, ByVal strPrivs As String, _
    ByVal lng����ID As Long, Optional ByRef blnThreeDeposit As Boolean, Optional ByRef lng����ID As Long, _
    Optional ByRef intԤ����� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��У��ҽ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-05-07 15:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, strValue As String

    On Error GoTo errH
    Call initCardSquareData
    mbytMzDeposit = Val(zlDatabase.GetPara("����Ԥ��ȱʡʹ�÷�ʽ", glngSys, 1137, 2))
    
    mblnExternal = True
    mlng����ID = lng����ID
    mstrPrivs = strPrivs
    
    strSql = "" & _
    "   Select a.����ID,a.��¼����,a.���㷽ʽ,a.�������,b.���� ��������,a.��Ԥ��,a.�ɿλ, " & _
    "           a.��λ������,a.��λ�ʺ�,C.��;����,C.��������,C.NO" & _
    "   From ����Ԥ����¼ a,���㷽ʽ b,���˽��ʼ�¼ C" & _
    "   Where a.��¼״̬ = 1 And a.���㷽ʽ = B.���� and A.����ID=C.ID " & _
    "          And ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
    
    mlng����ID = Val("" & rsTmp!����ID)
    
    mbln��;���� = Val(Nvl(rsTmp!��;����))
    
    mintԤ����� = 2
    If Val(Nvl(rsTmp!��������)) = 1 Then mintԤ����� = 1
    
    mbln������� = Val(Nvl(rsTmp!��������)) = 1
    
    mstrNO = Nvl(rsTmp!NO)
   
    rsTmp.Filter = "(��¼����=2 And ��������=3) or (��¼����=2 And ��������=4)"
    If rsTmp.RecordCount > 0 Then mstr������Ϣ = zlCommFun.Nvl(rsTmp!�ɿλ, " ") & "," & zlCommFun.Nvl(rsTmp!��λ������, " ") & "," & zlCommFun.Nvl(rsTmp!��λ�ʺ�, " ")


    rsTmp.Filter = 0    '����ȡʵ�ս��,��Ϊ���������ٽ���ʱ,������ϸû��ʵ�ս��
    strSql = "" & _
    "   Select Sum(nvl(���ʽ��,0)) As ���ʽ��" & _
    "   From (  Select nvl(���ʽ��,0) as ���ʽ��    From ������ü�¼  Where Nvl(���ӱ�־,0) <> 9 And ����id = [1]  UNION ALL  " & _
    "           Select nvl(���ʽ��,0) as ���ʽ��    From סԺ���ü�¼  Where Nvl(���ӱ�־,0) <> 9 And ����id = [1] ) "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)

    mcur���ʽ�� = Val(Nvl(rsTmp!���ʽ��))


    '������Ϣ
    rsTmp.Filter = 0
    strSql = "" & _
    "   Select ���㷽ʽ,��� From ���ս�����ϸ " & _
    "   Where ����id = [1] And ���㷽ʽ<>'�ֽ�' and ��־=1"  'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
    
    mstr���ս��� = ""   '���㷽ʽ|������||
    For i = 1 To rsTmp.RecordCount
        mstr���ս��� = mstr���ս��� & "||" & rsTmp!���㷽ʽ & "|" & rsTmp!���
        rsTmp.MoveNext
    Next
    If mstr���ս��� <> "" Then mstr���ս��� = Mid(mstr���ս���, 3)


    mintInsure = 0
    If mlng����ID <> 0 Then
        If mintԤ����� = 1 Then
            '����
            strSql = "Select ����,0 As ��ҳid From ������Ϣ Where ����id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", mlng����ID)
            If Not rsTmp.EOF Then mintInsure = zlCommFun.Nvl(rsTmp!����, 0): mlng��ҳID = zlCommFun.Nvl(rsTmp!��ҳID, 0)
        Else
            strSql = "Select ����,��ҳid From ������ҳ Where ����id = [1]" & _
                     " And ��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = [1])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", mlng����ID)
            If Not rsTmp.EOF Then mintInsure = zlCommFun.Nvl(rsTmp!����, 0): mlng��ҳID = zlCommFun.Nvl(rsTmp!��ҳID, 0)
        End If
    End If

    mstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 3, 1)))

    mbytInFun = 1
    Me.Show 1, frmParent
    blnThreeDeposit = mblnThreeDepositAfter
    lng����ID = mlng����ID
    intԤ����� = mintԤ�����
    ShowMeFromOut = mblnOk

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(ByRef frmParent As Object, ByVal lng����ID As Long, _
    ByVal lng����ID As Long, ByVal bln��;���� As Boolean, _
        ByVal cur���ʽ�� As Currency, ByVal str���ս��� As String, ByVal str������Ϣ As String, _
        ByVal intInsure As Integer, ByVal strȱʡ���λ�� As String, ByVal bytȱʡ�ֱҷ�ʽ As Byte, _
        ByVal cur�ɿ� As Currency, ByVal strҽ���� As String, _
        ByVal bytMCMode As Byte, ByVal strסԺ���� As String, _
        ByVal intԤ����� As Integer, ByRef blnThreeDepositAfter As Boolean, ByVal strStyle As String, ByRef rsCardType As ADODB.Recordset, _
        ByRef objPayCards As Cards, ByRef objCard As Card, ByVal strPrivs As String, bln������� As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��Ԥ��������㲻һ�µĽ϶Դ���
    '���:bytMCMode=ҽ���������֤��ģʽ,����1-����,2-סԺ����ģʽ,0-��ʾ��ҽ��
    '     intԤ�����-Ԥ�����:0-�����סԺ;1-����;2-סԺ
    '     objCard-�ϴη�������������
    '     bln�������- �Ƿ��������
    '����:У�Գɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-23 10:34:39
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng����ID = lng����ID: mstrסԺ���� = strסԺ����: mintԤ����� = intԤ�����
    mlng����ID = lng����ID: mbln��;���� = bln��;����
    mstr���ս��� = str���ս���: mstr������Ϣ = str������Ϣ     '����ҽ���洢:�������,��������,�����ʺ�
    mcur���ʽ�� = cur���ʽ��: mintInsure = intInsure: mstrҽ���� = strҽ����
    mcur�ɿ� = cur�ɿ�: mstrDec = strȱʡ���λ��: mbytMzDeposit = Val(zlDatabase.GetPara("����Ԥ��ȱʡʹ�÷�ʽ", glngSys, 1137, 2))
    mBytMoney = bytȱʡ�ֱҷ�ʽ: mbytMCMode = bytMCMode
    mbln������� = bln�������
    mblnThreeDepositAfter = blnThreeDepositAfter
    mstrStyle = strStyle
    mstrPrivs = strPrivs
    Set mobjCard = objCard
    Set mobjPayCards = objPayCards
    Set mrsCardType = rsCardType
    
    If gblnLED Then 'Led����ʾ���
        mcur������� = gclsInsure.SelfBalance(mlng����ID, mstrҽ����, IIf(mbytMCMode = 1, 10, 40), mcur����͸֧, mintInsure)
    End If
    mbytInFun = 0
    Me.Show 1, frmParent
    ShowMe = mblnOk
    blnThreeDepositAfter = mblnThreeDepositAfter
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    mblnClickOK = True: Unload Me
End Sub

Private Sub cmdOK_Click()
    '�������
    Dim strNotValiedNos As String, blnPrint As Boolean
    Dim i As Long, cllPro As Collection, blnPrintBillEmpty As Boolean
    Dim str���ʽ��� As String, str���NO As String, str��Ԥ�� As String
    Dim str�������ʽ��� As String, strCash As String
    Dim objCard As Card, strSql As String
    If Val(txtMargin.Text) <> 0 Then
        If InStr(1, mstrPrivs, ";����Ԥ������;") = 0 Then
            If Val(txtMargin.Text) > 0 Then
                MsgBox "����֧������,�밴����ʾ�Ĳ��", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            Else
                MsgBox "����֧��������,�밴����ʾ�Ĳ���˿", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            End If
        Else
            If Val(txtMargin.Text) > 0 Then
                MsgBox "���˳�Ԥ������,�밴����ʾ�Ĳ�������Ԥ����", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            Else
                MsgBox "���˳�Ԥ��������,�밴����ʾ�Ĳ�������Ԥ����", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If InStr(1, mstrPrivs, ";����Ԥ������;") > 0 Then
        For i = 1 To vsfMoney.Rows - 1
            If vsfMoney.RowData(i) = 999 Then
                If Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("���"))) < 0 Then
                    MsgBox "����Ԥ����������£����ʲ�֧���˿", vbExclamation, gstrSysName
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    If CheckThreePayDepositValied(objCard) = False Then Exit Sub
    
    '��������
    str���ʽ��� = ""
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("���"))) <> 0 And .RowData(i) <> 999 Then
                str���ʽ��� = str���ʽ��� & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & "|" & Val(.TextMatrix(i, .ColIndex("���"))) & "|"
                If InStr(",3,4,", "," & Val(.TextMatrix(i, .ColIndex("����"))) & ",") = 0 Then
                     'Oracle���̸��ݽ�������ֶ��ж��Ƿ�ҽ��,���ԽɷѵĽ�����벻�ܺ���,��
                     '���㷽ʽ|������|�������||.....
                    str���ʽ��� = str���ʽ��� & IIf(.TextMatrix(i, .ColIndex("�������")) = "", " ", .TextMatrix(i, .ColIndex("�������")))
                Else
                    str���ʽ��� = str���ʽ��� & IIf(mstr������Ϣ = "", " ", mstr������Ϣ)
                    '���㷽ʽ|������|�������,��������,�����ʺ�||.....
                End If
            End If
        Next
    End With
    If str���ʽ��� <> "" Then str���ʽ��� = Mid(str���ʽ���, 3)
   
    
    For i = 1 To vsDeposit.Rows - 1
        If Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("��Ԥ��"))) <> 0 Then     'ID|���ݺ�|���|��¼״̬||  IdΪ���ʾ��Ԥ�����(�ǵ�һ��)
            str��Ԥ�� = str��Ԥ�� & "||" & vsDeposit.TextMatrix(i, vsDeposit.ColIndex("ID"))
            str��Ԥ�� = str��Ԥ�� & "|" & vsDeposit.TextMatrix(i, vsDeposit.ColIndex("���ݺ�"))
            str��Ԥ�� = str��Ԥ�� & "|" & Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("��Ԥ��")))
            str��Ԥ�� = str��Ԥ�� & "|" & Val(vsDeposit.RowData(i))
        End If
    Next
    If str��Ԥ�� <> "" Then str��Ԥ�� = Mid(str��Ԥ��, 3)
    Set cllPro = New Collection
    
    If Not objCard Is Nothing Then
        '���㷽ʽ|������|�����ID|����|������ˮ��|����˵��||...
        str�������ʽ��� = objCard.���㷽ʽ
        str�������ʽ��� = str�������ʽ��� & "|" & -1 * mCurBrushCard.dblMoney
        str�������ʽ��� = str�������ʽ��� & "|" & objCard.�ӿ����
        str�������ʽ��� = str�������ʽ��� & "|" & IIf(mCurBrushCard.str���� = "", " ", mCurBrushCard.str����)
        str�������ʽ��� = str�������ʽ��� & "|" & IIf(mCurBrushCard.str������ˮ�� = "", " ", mCurBrushCard.str������ˮ��)
        str�������ʽ��� = str�������ʽ��� & "|" & IIf(mCurBrushCard.str����˵�� = "", " ", mCurBrushCard.str����˵��)
    End If
    
    If mstrForceNote <> "" Then
        With vsDeposit
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("�����ID"))) <> 0 And Val(.TextMatrix(i, .ColIndex("�Ƿ�����"))) = 0 Then
                    strCash = strCash & "," & .TextMatrix(i, .ColIndex("���㷽ʽ")) & Format(.TextMatrix(i, .ColIndex("��Ԥ��")), "0.00") & "Ԫ"
                End If
            Next i
            If strCash <> "" Then strCash = Mid(strCash, 2)
            mstrForceNote = mstrForceNote & strCash
        End With
    End If
    
    'Zl_סԺ�շѽ���_Update
    strSql = "Zl_סԺ�շѽ���_Update("
    '  ����id_In       סԺ���ü�¼.����id%Type,
    strSql = strSql & "" & mlng����ID & ","
    '  ���ʽ���_In     Varchar2, --���ʽ���_IN-��ҽ��ʱ:���㷽ʽ|������|�������||.....ҽ��ʱ:���㷽ʽ|������|�������,��������,�����ʺ�||.....
    strSql = strSql & "" & IIf(str���ʽ��� = "", "NULL", "'" & str���ʽ��� & "'") & ","
    '  ��Ԥ��_In       Varchar2, --��Ԥ��_IN= ID|���ݺ�|���|��¼״̬||.....
    strSql = strSql & "" & IIf(str��Ԥ�� = "", "Null", "'" & str��Ԥ�� & "'") & ","
    '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
    strSql = strSql & "" & IIf(Val(txt�ɿ�.Text) <> 0, "NULL", Val(txt�ɿ�.Text)) & ","
    '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
    strSql = strSql & "" & IIf(Val(txt�Ҳ�.Text) <> 0, "NULL", Val(txt�Ҳ�.Text)) & ","
    '  �����ʻ�����_In Varchar2 := Null --:���㷽ʽ|������|�����ID|����|������ˮ��|����˵��||...
    strSql = strSql & "" & IIf(str�������ʽ��� = "", "NULL", "'" & str�������ʽ��� & "'") & ","
    '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null
    strSql = strSql & "'" & mstrForceNote & "')"
    
    zlAddArray cllPro, strSql
    
    On Error GoTo errH
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If ExecuteThreeSwapPayInterface(objCard, mlng����ID, mCurBrushCard.dblMoney) = False Then Exit Sub
    
    If mblnExternal Then
        Call ReInitPatiInvoice
        blnPrint = True
        Select Case mintInvoiceMode
        Case 0: blnPrint = False '����ӡ
        Case 2  '�Զ���ӡ
            If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) <> vbYes Then
                blnPrint = False
            End If
        End Select
        
        If blnPrint Then
            If gblnStrictCtrl Then   '�ϸ�Ʊ�ݹ���
                If Trim(mstrInvoice) = "" Then
                    Call RefreshFact
                End If
                mlng����ID = GetInvoiceGroupID(IIf(mbytInvoiceKind = 0, 3, 1), 1, mlng����ID, mlngShareUseID, mstrInvoice, mstrUseType)
                If mlng����ID <= 0 Then
                    Select Case mlng����ID
                        Case 0 '����ʧ��
                        Case -1
                            MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                        Case -2
                            MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                        Case -3
                            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��", vbInformation, gstrSysName
                    End Select
                    Exit Sub
                End If
            Else
                If Len(mstrInvoice) <> gbytFactLength And mstrInvoice <> "" Then
                    MsgBox "Ʊ�ݺ��볤�Ȳ���ȷ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        If blnPrint Then
RePrint:
            Call frmPrint.ReportPrint(1, mstrNO, mlng����ID, mlng����ID, mlngShareUseID, mstrUseType, mstrInvoice, zlDatabase.Currentdate, txt�ɿ�.Text, txtMargin.Text, , mintInvoiceFormat, blnPrintBillEmpty, mbytInvoiceKind + 1)
            strSql = "Zl_Ʊ����ʼ��_Update('" & mstrNO & "','" & Trim(mstrInvoice) & "',3)"
            If gblnStrictCtrl And blnPrintBillEmpty = False And _
                ((mbytInvoiceKind = 0 And InStr(1, mstrPrivs, ";�վݴ�ӡ;") > 0) _
                   Or (mbytInvoiceKind <> 0 And InStr(1, mstrPrivs, ";��ӡ�����շ�Ʊ��;") > 0)) Then    'blnPrintBillEmpty:55052
                If zlIsNotSucceedPrintBill(3, mstrNO, strNotValiedNos) = True Then
                    If MsgBox("���ʵ���Ϊ[" & strNotValiedNos & "]�Ľ���Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����´�ӡ����Ʊ��?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                    strSql = "Zl_Ʊ����ʼ��_Update('" & mstrNO & "','" & "',3)"
                End If
            End If
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        End If
    End If
    
    mblnOk = True: mblnClickOK = True: Unload Me
    Exit Sub
    
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If Not objCard Then mblnClickOK = True: Unload Me
End Sub

Private Sub ReInitPatiInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim intInsure As Integer
    intInsure = mintInsure
    lng����ID = mlng����ID
    lng��ҳID = mlng��ҳID
    mlng����ID = 0
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng����ID, lng��ҳID, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType, IIf(mintԤ����� = 1, "1", "2"))
    mintInvoiceMode = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    Call RefreshFact
    mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, IIf(mintԤ����� = 1, "1", "2"))
End Sub

Private Sub RefreshFact()
    Dim bytInvoiceKind As Byte
    '���ܣ�ˢ���շ�Ʊ�ݺ�
    If mintInvoiceMode = 0 Then Exit Sub
    
    If mintԤ����� = 1 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("�������Ʊ������", glngSys, mlngModul, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, mlngModul, "0"))
    End If
    
    mbytInvoiceKind = bytInvoiceKind
    
    If gblnStrictCtrl Then
        mlng����ID = CheckUsedBill(IIf(bytInvoiceKind = 0, 3, 1), IIf(mlng����ID > 0, mlng����ID, mlngShareUseID), , mstrUseType)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            mstrInvoice = ""
        Else
            '�ϸ�ȡ��һ������
            mstrInvoice = GetNextBill(mlng����ID)
        End If
    Else
        '��ɢ��ȡ��һ������
        mstrInvoice = IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
    End If
End Sub

Private Sub cmdReturnCash_Click()
    Dim dblMoney As Double, lngRow As Long
    Dim str����Ա���� As String, strDBUser As String
    Dim strPrivs As String
    Dim intCount As Integer, intNotCashCount As Integer
    If mstrForceNote <> "" Then Exit Sub
    
    Call GetDelThreeCardDepositInfor(intCount, intNotCashCount, mblnThreeDepositAfter, mstrStyle)
    If mstrStyle = "" Then Exit Sub
    
    If InStr(";" & mstrCardPrivs & ";", ";�����˿�ǿ������;") = 0 And intNotCashCount > 0 Then
        str����Ա���� = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
        If str����Ա���� = "" Then
            MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        mstrForceNote = str����Ա���� & "ǿ������:"
    Else
        If intNotCashCount <> 0 Then
            If MsgBox("ѡ��Ľ��㿨��֧������,�Ƿ�ǿ�ƽ������֣�", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        End If
        
        mstrForceNote = UserInfo.���� & "ǿ������:"
    End If
    
    Call ShowMoney(True)

End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
    Call LedDisplayBank
    
End Sub
Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim rsӦ�ó��� As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim arrMediCare As Variant
    Dim bln������� As Boolean, blnExist As Boolean
    Dim str���õ�ҽ�����㷽ʽ As String
    Dim intCount As Integer
    
    mlngModul = 1137
    
    '������ʼ
    If mintInsure <> 0 Then
        mblnCent = gclsInsure.GetCapability(support�ֱҴ���, , mintInsure)
    Else
        mblnCent = Not gBytMoney = 0
    End If
    
    mcur�շ���� = 0
    mblnOk = False
    mblnClickOK = False
    mintDefault = 0
    mcurMediCare = 0
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    
    'ȷ����ȡ����ť
    If mbytInFun = 0 Then
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False
    Else
        cmdCancel.Visible = True
    End If
    
    '��ʾԤ����ϸ
    Call AdjustDepost
    Set rsTmp = GetDepositBefor(mlng����ID, mstrסԺ����, mintԤ�����)
    intCount = 0
    If Not rsTmp Is Nothing Then
        'mbytMzDeposit As Byte '����Ԥ��ȱʡʹ�÷�ʽ:0-ȱʡ��ʹ�ý�;1-�����ʽ��ʹ��Ԥ��;2-ʹ������Ԥ��
        With vsDeposit
            .Redraw = flexRDNone
            .Rows = IIf(rsTmp.RecordCount <> 0, rsTmp.RecordCount, 1) + 1
            .Cell(flexcpBackColor, 1, .ColIndex("��Ԥ��"), .Rows - 1, .ColIndex("��Ԥ��")) = txtMoney.BackColor
            .Cell(flexcpBackColor, 1, .ColIndex("���"), .Rows - 1, .ColIndex("���")) = 12900351
            
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(Nvl(rsTmp!��¼״̬))
                .Cell(flexcpData, i, .ColIndex("ID")) = Nvl(rsTmp!�����ID) & "||" & Nvl(rsTmp!ת�ʼ�����) & "||" & Nvl(rsTmp!����) & "||" & Nvl(rsTmp!ȱʡ����)
                
               If Val(Nvl(rsTmp!�����ID)) <> 0 And Nvl(rsTmp!ȱʡ����) = 0 Then
                    If mblnExternal Then
                        If InStr("," & mstrStyle & ",", rsTmp!���㷽ʽ) = 0 Then
                            mstrStyle = mstrStyle & "," & rsTmp!���㷽ʽ
                        End If
                    End If
                    intCount = intCount + 1
                End If
                
                .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(rsTmp!ID))
                .TextMatrix(i, .ColIndex("���ݺ�")) = Nvl(rsTmp!NO)
                .TextMatrix(i, .ColIndex("����")) = Format(rsTmp!����, "yyyy-MM-dd")
                .TextMatrix(i, .ColIndex("���㷽ʽ")) = IIf(IsNull(rsTmp!���㷽ʽ), "", rsTmp!���㷽ʽ)
                .TextMatrix(i, .ColIndex("���")) = Format(rsTmp!���, "0.00")
                
                .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(rsTmp!���, "0.00")
                .TextMatrix(i, .ColIndex("Ԥ��ID")) = Val(Nvl(rsTmp!Ԥ��ID))
                .TextMatrix(i, .ColIndex("�����ID")) = Val(Nvl(rsTmp!�����ID))
                .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(Nvl(rsTmp!����))
                rsTmp.MoveNext
            Next
            If intCount > 1 And InStr(1, mstrPrivs, ";����Ԥ������;") = 0 Then
                mblnThreeDepositAfter = True
            End If
            
            .Row = 1: .Col = .ColIndex("��Ԥ��")
            .Redraw = flexRDBuffered
            
            If mblnExternal And mstrStyle <> "" Then
                mstrStyle = Mid(mstrStyle, 2)
            End If
        End With
    End If
    
    '��ʾ���ս��㼰�ָ����㷽ʽ,��ʹ��֧��ʹ�ø���,Ҳ������,����ҽ���Ĳ������
    arrMediCare = Array()                   '���㷽ʽ|������||
    If mstr���ս��� <> "" Then arrMediCare = Split(mstr���ս���, "||")
    
    On Error GoTo errH
    
    If InStr(1, mstrPrivs, ";����Ԥ������;") > 0 Then
        strSql = _
        " Select Distinct B.����,B.����,B.����,A.ȱʡ��־,1 As λ��" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where (B.����=3 OR B.����=4)  " & _
        "       And B.����=A.���㷽ʽ(+) and instr(',7,8,',','||����||',')=0 " & _
        " Union " & _
        " Select Null As ����, '��Ԥ��' As ����, 999 As ����,0 As ȱʡ��־,0 As λ��" & _
        " From Dual " & _
        " Order By λ��,����,����"
    Else
        strSql = _
        " Select Distinct B.����,B.����,B.����,A.ȱʡ��־,1 As λ��" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where ((A.Ӧ�ó���='����' And B.����<>3 And B.����<>4) OR (B.����=3 OR B.����=4))  " & _
        "       And B.����=A.���㷽ʽ(+) and instr(',7,8,',','||����||',')=0 " & _
        " Union " & _
        " Select ����,����,����,ȱʡ��־,0 As λ��" & _
        " From ���㷽ʽ Where ����=9 " & _
        " Order By λ��,����,����"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    strSql = "Select Ӧ�ó���,���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó���='����'"
    Set rsӦ�ó��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Call InitBalanceGrid(vsfMoney)
    With vsfMoney
        .Redraw = flexRDNone
        .Rows = IIf(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        i = 1
        Do While Not rsTmp.EOF
            .RowData(i) = Nvl(rsTmp!����, 1)                '�����ж��Ƿ�����޸Ľ��,�Լ��Ƿ����ֽ�
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = rsTmp!����
            .TextMatrix(i, .ColIndex("���")) = "0.00"
            .TextMatrix(i, .ColIndex("����")) = Nvl(rsTmp!����, 1)
            .TextMatrix(i, .ColIndex("ȱʡ")) = Nvl(rsTmp!ȱʡ��־, 0)
            'ȱʡ���㷽ʽ(û�������ֽ�) ��������ҽ��
            If InStr(",3,4,", "," & Nvl(rsTmp!����, 1) & ",") = 0 Then
                If Nvl(rsTmp!ȱʡ��־, 0) = 1 Then mintDefault = i
                If Nvl(rsTmp!����, 1) = 1 And mintDefault = 0 Then mintDefault = i
                If Nvl(rsTmp!����, 1) = 9 And mintError = 0 Then
                    mintError = i: .Row = i: .Col = 0
                    .CellForeColor = vbRed
                End If
                i = i + 1
            Else
                '���ս���
                blnExist = False
                For j = 0 To UBound(arrMediCare)
                    If Split(arrMediCare(j), "|")(0) = rsTmp!���� Then
                        blnExist = True
                        rsӦ�ó���.Filter = "���㷽ʽ='" & rsTmp!���� & "'"
                        If rsӦ�ó���.EOF And Nvl(rsTmp!����, 1) <> 9 Then
                            MsgBox "ע��:���㷽ʽ[" & rsTmp!���� & "]δ����Ӧ����[����]����,�뵽[���㷽ʽ����]������!", vbInformation, gstrSysName
                        End If
                        
                        .TextMatrix(i, .ColIndex("���")) = Split(arrMediCare(j), "|")(1)
                        .TextMatrix(i, .ColIndex("�������")) = ""    '�޽������
                        mcurMediCare = mcurMediCare + Val(.TextMatrix(i, .ColIndex("���")))
                        Exit For
                    End If
                Next
                If blnExist Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE7CFBA
                     i = i + 1
                End If
                str���õ�ҽ�����㷽ʽ = str���õ�ҽ�����㷽ʽ & "," & rsTmp!����
            End If
            rsTmp.MoveNext
        Loop
        .Rows = i: .Redraw = True
    End With
    
    '�ȼ��ÿһ��ҽ�����㷽ʽ�Ƿ񶼴���
    If mstr���ս��� <> "" Then
        str���õ�ҽ�����㷽ʽ = str���õ�ҽ�����㷽ʽ & ","
        For j = 0 To UBound(arrMediCare)
            If InStr(str���õ�ҽ�����㷽ʽ, "," & Split(arrMediCare(j), "|")(0) & ",") <= 0 Then
                MsgBox "ҽ�����㷽ʽ[" & Split(arrMediCare(j), "|")(0) & "]δ����,���ȵ�[���㷽ʽ����]������!", vbInformation, gstrSysName
                cmdCancel.Visible = True
                cmdOK.Visible = False
            End If
        Next
    End If
    
    
    '���ʽ��
    txtTotal.Text = Format(mcur���ʽ��, mstrDec)
    
    '��Ԥ��,���ʽ���ȥҽ����ʽ���ʺ�����
    Call ShowMoney(True)
    If LoadBalance = False Then
        cmdCancel.Visible = True
        cmdOK.Visible = False
    End If
    
    If mintDefault > 0 Then
        vsfMoney.Row = mintDefault: vsfMoney.Col = 0
        vsfMoney.CellFontBold = True
        vsfMoney.Col = 1
    Else        '���㷽ʽû��ȱʡֵ,�������ֽ�ʽ�����
        vsfMoney.Row = 1: vsfMoney.Col = 1
    End If
    txt�ɿ�.Text = Format(mcur�ɿ�, "0.00")
    Call LedDisplayBank
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function RecalDeposit(ByRef cur���ʺϼ� As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����Ԥ��
    '���:cur���ʺϼ�-��ǰ�Ľ��ʽ��
    '����:cur���ʺϼ�-����δ������ɵĽ��ʽ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-20 14:04:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln���� As Boolean, i As Long, varData As Variant
    
    On Error GoTo errHandle
    bln���� = cur���ʺϼ� < 0
        
    If InStr(1, mstrPrivs, ";����Ԥ������;") > 0 Then
       With vsDeposit
           For i = 1 To .Rows - 1
               If cur���ʺϼ� = 0 Then
                   .TextMatrix(i, .ColIndex("��Ԥ��")) = "0.00"
               Else
                   If Val(.TextMatrix(i, .ColIndex("���"))) <= Format(cur���ʺϼ�, "0.00") Then
                       .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                   Else
                       .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(cur���ʺϼ�, "0.00")
                   End If
                   cur���ʺϼ� = cur���ʺϼ� - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
               End If
           Next i
       End With
       RecalDeposit = True
       Exit Function
    End If
    
    With vsDeposit
        For i = 1 To .Rows - 1
            '�����ID||����||�Ƿ�����||ȱʡ����
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||", "||")
            
            'mbytMzDeposit-������������Ч,0-��ʾȫ��;1-������ݽ��ʽ������̯Ԥ��;2-Ԥ����ȫ��
            If mbln������� And mbytMzDeposit = 0 Then
                '������ʲ�ʹ�ó�Ԥ��
                 .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(0, "0.00")
            ElseIf mblnThreeDepositAfter Then
                '1.����Ԥ��
                If Val(varData(0)) <> 0 Or Val(varData(3)) = 1 Then
                    If bln���� And Val(varData(3)) <> 1 Then
                        .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(0, "0.00")
                    ElseIf Val(.TextMatrix(i, .ColIndex("���"))) <= Format(cur���ʺϼ�, "0.00") Then
                        .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                        cur���ʺϼ� = cur���ʺϼ� - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                    Else
                        .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(cur���ʺϼ�, "0.00")
                        cur���ʺϼ� = 0
                    End If
                Else
                   If mbln������� Then
                        If mbytMzDeposit = 2 Then
                            'Ԥ����ȫ��
                            .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                            cur���ʺϼ� = cur���ʺϼ� - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                        Else
                            '�����ʽ���
                            If Val(.TextMatrix(i, .ColIndex("���"))) <= Format(cur���ʺϼ�, "0.00") Then
                                 .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                                cur���ʺϼ� = cur���ʺϼ� - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                            Else
                                .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(cur���ʺϼ�, "0.00")
                                cur���ʺϼ� = 0
                            End If
                            
                        End If
                   Else
                        If Not mbln��;���� Or Val(.TextMatrix(i, .ColIndex("���"))) <= Format(cur���ʺϼ�, "0.00") Then
                            '��Ժ����ȫ��������(����˾����ָ�)
                             .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                            cur���ʺϼ� = cur���ʺϼ� - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                        Else
                            .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(cur���ʺϼ�, "0.00")
                            cur���ʺϼ� = 0
                        End If
                    End If
                End If
            ElseIf Not mbln��;���� Or (mbln������� And mbytMzDeposit = 2) Then
               '2.��Ժ����ȫ��������(����˾����ָ�)���������ȱʡȫ������Ԥ��
                .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                cur���ʺϼ� = cur���ʺϼ� - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
            Else
                '3.��;����ֻ���㹻��
                If cur���ʺϼ� = 0 Then
                     .TextMatrix(i, .ColIndex("��Ԥ��")) = "0.00"
                Else
                    If Val(.TextMatrix(i, .ColIndex("���"))) <= Format(cur���ʺϼ�, "0.00") Then
                         .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                         cur���ʺϼ� = cur���ʺϼ� - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                    Else
                         .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(cur���ʺϼ�, "0.00")
                         cur���ʺϼ� = 0
                    End If
                End If
            End If
        Next
    End With
    RecalDeposit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ShowMoney(Optional ByVal blnAutoSet As Boolean) As String
    '���ܣ����ú���ʾ����ĸ��ֽ��

    Dim i As Long, j As Long
    Dim cur���ʺϼ� As Currency, curMoney As Currency, curOwn As Currency
    Dim curԤ���ϼ� As Currency, cur��Ԥ���ϼ� As Currency, curӦ�ɽ�� As Currency
    Dim bln���ڲ��� As Boolean  'ֻ�е�û��ȱʡ���㷽ʽ,�����޸�ȱʡ���㷽ʽ�Ľ��ʱ,����
    Dim curTmp As Currency
    Dim varData As Variant
    Dim bln���� As Boolean
    
    
    '�����Զ���Ԥ������Ľ�����
    '---------------------------------------------------------------------------------------------
    Call ShowDelThreeSwap

    If blnAutoSet Then
        cur���ʺϼ� = mcur���ʽ�� - mcurMediCare
        If InStr(1, mstrPrivs, ";����Ԥ������;") > 0 Then
            With vsfMoney
                For i = 1 To .Rows - 1
                    If .RowData(i) = 999 Then
                        .TextMatrix(i, 1) = Format(cur���ʺϼ�, "0.00")
                    End If
                Next i
            End With
         End If
        '���¼���Ԥ����
        Call RecalDeposit(cur���ʺϼ�)
    Else
        '�޸ĳ�Ԥ����������
        cur���ʺϼ� = mcur���ʽ�� - GetSumMoney
        If mintDefault <> 0 And (Not Me.ActiveControl Is vsfMoney Or _
                                Me.ActiveControl Is vsfMoney And mintDefault <> vsfMoney.Row) Then
            With vsfMoney
                If Val(.TextMatrix(mintDefault, .ColIndex("����"))) And mblnCent Then   '�ֽ�ʱҪ���зֱҴ���
                    .TextMatrix(mintDefault, .ColIndex("���")) = Format(CentMoney(Val(.TextMatrix(mintDefault, .ColIndex("���"))) + cur���ʺϼ�), "0.00")
                Else
                    .TextMatrix(mintDefault, .ColIndex("���")) = Format(Val(.TextMatrix(mintDefault, .ColIndex("���"))) + cur���ʺϼ�, "0.00")
                End If
            End With
        Else
            bln���ڲ��� = True
        End If
    End If
    
    '��ʾ��ǰ��Ԥ������
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetSumMoney
    
    '�����ǲ��,��һ�����ֽ�,���Բ�����ֱ�,lblDelMoney.Tag�����˵������ʻ��Ľ��
    curOwn = Val(txtTotal.Text) - curMoney
    txtMargin.Text = Format(curOwn, "0.00")
    
    '���ݲ���Զ���ƽ������'ʣ�ಿ�ݳ������õ�ȱʡ���㷽ʽ��
    '-----------------------------------------------------------------------------------------------------
    If Val(txtMargin.Text) <> 0 And mintDefault <> 0 And (vsfMoney.Row <> mintDefault Or blnAutoSet) Then
        curTmp = Val(vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("���"))) + curOwn
        If Abs(curTmp) >= 0.01 Then
            If mintError <> 0 And mblnCent Then
                vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("���")) = Format(CentMoney(curTmp), "0.00")
            Else
                vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("���")) = Format(curTmp, "0.00")
            End If
        Else
            vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("���")) = "0.00"
        End If
        txtMargin.Text = "0.00"
    End If
    
    '���������(������-���ʽ��)
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetSumMoney(curԤ���ϼ�, cur��Ԥ���ϼ�, curӦ�ɽ��)

    '�п���Ӧ����������Ǵ���ֱҵ�����,�Ͳ���ʾ��
    If Val(txtMargin.Text) <> 0 And mintDefault <> 0 Then
        If Abs(Val(txtMargin.Text)) < 0.1 Or gBytMoney = 5 And Abs(Val(txtMargin.Text)) < 0.3 Then
            If CentMoney(Val(vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("���"))) + Val(txtMargin.Text)) = Val(vsfMoney.TextMatrix(mintDefault, vsfMoney.ColIndex("���"))) Then
                txtMargin.Text = "0.00"
            End If
        End If
    End If
    
    '����Ӧ��������С�������������,�����������С��1��,�Ͳ���ʾ��
    If Val(txtMargin.Text) <> 0 And mcur�շ���� + curOwn = 0 And Abs(curOwn) <= 0.005 Then
        txtMargin.Text = "0.00"
    End If
        
    If mintError <> 0 And Val(txtMargin.Text) = 0 Then
        vsfMoney.TextMatrix(mintError, vsfMoney.ColIndex("���")) = Format(Val(txtTotal.Text) - curMoney, mstrDec)
        If Val(txtTotal.Text) - curMoney <> 0 Then
            vsfMoney.RowHidden(mintError) = False
        Else
            vsfMoney.RowHidden(mintError) = True
        End If
    Else
        mcur�շ���� = Format(curMoney - Val(txtTotal.Text), mstrDec)
        vsfMoney.TextMatrix(mintError, vsfMoney.ColIndex("���")) = Format(vsfMoney.TextMatrix(mintError, vsfMoney.ColIndex("���")), mstrDec)
    End If
    
    lblԤ�����.Caption = lblԤ�����.Tag & Format(curԤ���ϼ�, "0.00")
    lblԤ�����.ToolTipText = "����δ��Ԥ��֮ǰ��Ԥ�����"
    mcurԤ���ϼ� = curԤ���ϼ�
    lbl��Ԥ��.Caption = lbl��Ԥ��.Tag & Format(cur��Ԥ���ϼ�, "0.00")
    mcur��Ԥ���ϼ� = cur��Ԥ���ϼ�
    lblҽ��֧��.Caption = lblҽ��֧��.Tag & Format(mcurMediCare, "0.00")
    lblӦ��.Caption = lblӦ��.Tag & Format(curӦ�ɽ��, "0.00")
    mcurӦ�ɽ�� = curӦ�ɽ��
    
    lblԤ�����.Left = vsDeposit.Left
    lbl��Ԥ��.Left = lblԤ�����.Left + lblԤ�����.Width + 600
    lblҽ��֧��.Left = vsfMoney.Left
    lblӦ��.Left = lblҽ��֧��.Left + lblҽ��֧��.Width + 600
    
    
    Call Calc�Ҳ�
    Call LedDisplayBank
End Function

Private Sub vsfMoney_GotFocus()
        Call LedDisplayBank
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then        '��������
        If InStr(txtMoney.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Beep: Exit Sub
        
        If txtMoney.Left > vsfMoney.Left Then   '��������
            If vsfMoney.Col = vsfMoney.Cols - 1 Then    '�������,���������ڹ������ж��Ƿ���ҽ�����㷽ʽ
                If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        Else    'Ԥ������
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
         '��������ȷ��
        If txtMoney.Left > vsfMoney.Left Then
            If vsfMoney.Col = vsfMoney.Cols - 1 Then    '��������
                If InStr(txtMoney.Text, "'") > 0 Or InStr(txtMoney.Text, "|") > 0 Or InStr(txtMoney.Text, ",") > 0 Then
                    Exit Sub
                End If
                
                vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Trim(txtMoney.Text)
                txtMoney.Visible = False
            Else
                If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                    zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
                End If
                If Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.ColIndex("����"))) = 1 And mblnCent Then
                    txtMoney.Text = Format(CentMoney(Val(txtMoney.Text)), "0.00")
                End If
                                
                If Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col)) <> Format(Val(txtMoney.Text), "0.00") Then
                    vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Format(Val(txtMoney.Text), "0.00")
                    txtMoney.Visible = False
                    vsfMoney.SetFocus   '��������,ShowMoney���Դ��ж�
                    
                    Call ShowMoney
                Else
                    txtMoney.Visible = False
                    vsfMoney.SetFocus
                End If
            End If
            
            If vsfMoney.Col < vsfMoney.Cols - 2 Then
                vsfMoney.Col = vsfMoney.Col + 1
            Else
                If vsfMoney.Row = vsfMoney.Rows - 1 Then
                    '��һ�ؼ�����
                    If GetӦ�� > 0 And txt�ɿ�.Visible Then
                        txt�ɿ�.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                Else
                    '��һ�д���
                    If Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.ColIndex("����"))) = 2 Then
                       If vsfMoney.Col = vsfMoney.Cols - 2 Then
                            vsfMoney.Col = vsfMoney.Cols - 1
                       Else
                            vsfMoney.Row = vsfMoney.Row + 1
                            vsfMoney.Col = vsfMoney.Cols - 2
                       End If
                    Else
                        vsfMoney.Row = vsfMoney.Row + 1
                        vsfMoney.Col = vsfMoney.Cols - 2
                    End If
                    
                    If vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(0) - 2) > 1 Then
                        vsfMoney.TopRow = vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(1) - 2)
                    End If
                End If
            End If
        
        'Ԥ������ȷ��
        Else
            If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
            End If
            
            '�޸Ĳ��ܳ�������
            If Val(txtMoney.Text) > Val(vsDeposit.TextMatrix(vsDeposit.Row, 4)) Then
                txtMoney.Text = Val(vsDeposit.TextMatrix(vsDeposit.Row, 4))
            End If
            
            If Val(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.Col)) <> Format(Val(txtMoney.Text), "0.00") Then
                vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.Col) = Format(Val(txtMoney.Text), "0.00")
                txtMoney.Visible = False
                vsDeposit.SetFocus '��������
                
                Call ShowMoney
            Else
                txtMoney.Visible = False
                vsDeposit.SetFocus
            End If
            
            If vsDeposit.Row = vsDeposit.Rows - 1 Then
                '��һ�ؼ�����
                vsfMoney.SetFocus
            Else
                '��һ�д���
                vsDeposit.Row = vsDeposit.Row + 1
                If vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(0) - 2) > 1 Then
                    vsDeposit.TopRow = vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(1) - 2)
                End If
                vsDeposit.Col = vsDeposit.Cols - 1
            End If
        End If
        
        If Val(txt�ɿ�.Text) > 0 Then Call txt�ɿ�_Change
    End If
End Sub

Private Sub txtMoney_LostFocus()
    txtMoney.Visible = False
End Sub

Private Sub txtMoney_Validate(Cancel As Boolean)
    If txtMoney.Visible Then Call txtMoney_KeyPress(13)
End Sub
Private Sub AdjustDepost()
    Dim bln As Boolean, i As Long
    With vsDeposit
        .Redraw = flexRDNone
        .Clear
        .Rows = 2: .Cols = 9: i = 0
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "���ݺ�": .ColWidth(i) = 1100: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1050: i = i + 1
        .TextMatrix(0, i) = "���㷽ʽ": .ColWidth(i) = 620: i = i + 1
        .TextMatrix(0, i) = "���": .ColWidth(i) = 1100: i = i + 1
        .TextMatrix(0, i) = "��Ԥ��": .ColWidth(i) = 1100: i = i + 1
        .TextMatrix(0, i) = "Ԥ��ID": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�����ID": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColHidden(i) = True: .ColWidth(i) = 0: i = i + 1
        
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            .FixedAlignment(i) = IIf(i = 3, flexAlignLeftCenter, flexAlignCenterCenter)
            Select Case .ColKey(i)
            Case "��Ԥ��", "���"
                .ColAlignment(i) = flexAlignRightCenter
            Case Else
                .ColAlignment(i) = flexAlignLeftCenter
            End Select
        Next
        .Row = 1: .Col = .Cols - 1
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDBuffered
    End With
End Sub
Private Function GetDepositBefor(ByVal lng����ID As Long, _
    ByVal strסԺ���� As String, ByVal intԤ����� As Integer) As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˱���ҽ������֮ǰ��ʣ��Ԥ������ϸ,�������γ�����Ԥ��
    '���:lng����ID-����ID
    '       strסԺ����-סԺ����,��:1,2,3
    '       intԤ�����-Ԥ�����:0-�����סԺ;1-����;2-סԺ
    '����:
    '����:����ҽ������֮ǰ��ʣ��Ԥ������ϸ,
    '����:���˺�
    '����:2013-11-14 11:36:01
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strSub1 As String
    Dim strWherePage As String, strPages As String
    On Error GoTo errH
    strPages = "," & strסԺ���� & ","
    
    strWherePage = IIf(strסԺ���� = "", "", " And instr([3],','||Nvl(A.��ҳID,0)||',')>0")
    If intԤ����� <> 0 Then
        strWherePage = strWherePage & " And A.Ԥ����� =[4]"
    End If
    
    '���Ӳ�ѯ��������Ԥ�����շѼ��˷�ʱ��һ��һ��,ע��ϵͳ�������ʵ�Ԥ�������Ԥ���˷�,��Ҫ���ϼ�¼״̬�ж�
    strSub1 = _
    "   Select NO,Sum(Nvl(A.���,0)) as ��� " & _
    "   From ����Ԥ����¼ A" & _
    "   Where (A.����ID Is Null " & strWherePage & " Or A.����ID=[1]) And Nvl(A.���, 0)<>0 And A.����ID=[2]" & _
    "   Group by NO  " & _
    "   Having Sum(Nvl(A.���,0))<>0"
    strSql = _
        "   Select Max(a.Id) As ID, Max(��¼״̬) As ��¼״̬, NO, Max(����) As ����, a.���㷽ʽ, Sum(a.���) As ���, �����id, ת�ʼ�����, Min(Ԥ��id) As Ԥ��id," & vbNewLine & _
        "       Nvl(b.�Ƿ�����, 0) As ����, Nvl(b.�Ƿ�ȫ��, 0) As ȫ��, Nvl(b.�Ƿ�ȱʡ����, 0) As ȱʡ���� " & _
        "   From( " & _
        "       Select A.ID,A.��¼״̬,A.NO,A.�տ�ʱ�� as ����,A.���㷽ʽ,Nvl(A.���,0) as ���,A.�����ID,C.�Ƿ�ת�ʼ����� as ת�ʼ�����,A.ID As Ԥ��ID" & _
        "       From ����Ԥ����¼ A,(" & strSub1 & ") B,ҽ�ƿ���� C" & _
        "       Where (A.����ID Is Null " & strWherePage & " Or A.����ID=[1]) And Nvl(A.���,0)<>0" & _
        "               And A.���㷽ʽ Not IN (Select ���� From ���㷽ʽ Where ����=5)" & _
        "               And A.�����ID=C.ID(+) And A.NO=B.NO And A.����ID=[2]" & _
        "       Union All" & _
        "       Select 0 as ID,A.��¼״̬,A.NO,Min(A.�տ�ʱ��) as ����,A.���㷽ʽ,Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0)) as ���," & _
        "           Max(A.�����ID) as �����ID,max(C.�Ƿ�ת�ʼ�����) as ת�ʼ�����,Min(A.ID) As Ԥ��ID " & _
        "       From ����Ԥ����¼ A,ҽ�ƿ���� C" & _
        "       Where A.��¼���� IN(1,11) And A.���㷽ʽ Not IN (Select ���� From ���㷽ʽ Where ����=5) And A.����ID is Not NULL And A.�����ID=C.ID(+) " & _
        "           And A.����ID<>[1] And Nvl(A.���,0)<>Nvl(A.��Ԥ��,0) And A.����ID=[2] " & strWherePage & _
        "       Having Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0))<>0" & _
        "       Group by A.��¼״̬,A.NO,A.���㷽ʽ " & _
        "       ) A, ҽ�ƿ���� B " & _
        "   Where a.�����id = b.Id(+) " & _
        "   Group By a.No, a.���㷽ʽ, a.�����id, a.ת�ʼ�����, Nvl(b.�Ƿ�����, 0), Nvl(b.�Ƿ�ȫ��, 0), Nvl(b.�Ƿ�ȱʡ����, 0)" & vbNewLine & _
        "   Having Sum(���) <> 0" & _
        "   Order By Decode(sign(Sum(a.���)),-1,0,1), Decode(Nvl(a.�����id, 0), 0, 0, Decode(Nvl(b.�Ƿ�����, 0), 0, 2, 1)) Desc," & vbNewLine & _
        "         Decode(Nvl(a.�����id, 0), 0, 0, Decode(Nvl(b.�Ƿ�ȫ��, 0), 0, 1, 2)) Desc, Nvl(a.�����id, 0) Desc, a.No, a.���㷽ʽ"
        
    Set GetDepositBefor = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng����ID, strPages, intԤ�����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetSumMoney(Optional ByRef curԤ���ϼ� As Currency, Optional ByRef cur��Ԥ���ϼ� As Currency, Optional ByRef curӦ�ɽ�� As Currency) As Currency
    Dim i As Long
    Dim curMoney As Currency
    curԤ���ϼ� = 0: cur��Ԥ���ϼ� = 0: curӦ�ɽ�� = 0
    With vsDeposit
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            For i = 1 To .Rows - 1
                curMoney = curMoney + Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                curԤ���ϼ� = curԤ���ϼ� + Val(.TextMatrix(i, .ColIndex("���")))
                cur��Ԥ���ϼ� = cur��Ԥ���ϼ� + Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
            Next
        End If
    End With
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("���"))) <> 0 And .RowData(i) <> 999 Then
                If Val(.TextMatrix(i, .ColIndex("����"))) <> 9 Then
                    curMoney = curMoney + Val(.TextMatrix(i, .ColIndex("���")))
                End If
                If InStr(",3,4,9,", "," & Val(.TextMatrix(i, .ColIndex("����"))) & ",") = 0 Then
                    curӦ�ɽ�� = curӦ�ɽ�� + Val(.TextMatrix(i, .ColIndex("���")))
                End If
            End If
        Next
    End With
    curMoney = curMoney - Val(lblDelMoney.Tag)
    GetSumMoney = curMoney
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnClickOK Then Cancel = 1: Exit Sub
    mblnExternal = False
    mstrStyle = ""
    mstrForceNote = ""
End Sub

Private Sub vsDeposit_DblClick()
   '���˺�:����Ȩ�޿��ƣ��������Ԥ�����ʣ��������ݲ���ȷ��֮ǰ����������ʱ��Ϊʲô�ܸ��ģ���ʱ��֪��ԭ�������ƣ�����������ݳ���
    If InStr(mstrPrivs, ";����Ԥ������;") > 0 Then Exit Sub
    If Not txtMoney.Visible And vsDeposit.Row >= 1 And vsDeposit.Col = 5 Then
        If mblnThreeDepositAfter And Val(Split(vsDeposit.Cell(flexcpData, vsDeposit.Row, vsDeposit.ColIndex("ID")) & "||", "||")(0)) <> 0 Then
            Exit Sub
        End If
        With txtMoney
            .Left = vsDeposit.Left + vsDeposit.CellLeft + 30
            .Top = vsDeposit.Top + vsDeposit.CellTop + (vsDeposit.CellHeight - txtMoney.Height) / 2 + 15
            .Width = vsDeposit.CellWidth - 30
            .ForeColor = vsDeposit.CellForeColor
            .BackColor = vsDeposit.CellBackColor
            .Alignment = 1
            .Text = vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub vsDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vsDeposit.Col = 0 Then
            vsDeposit.Col = vsDeposit.Col + 1
        ElseIf vsDeposit.Row < vsDeposit.Rows - 1 Then
            vsDeposit.Row = vsDeposit.Row + 1
            vsDeposit.Col = vsDeposit.Cols - 1
            If vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(0) - 2) > 1 Then
                vsDeposit.TopRow = vsDeposit.Row - (vsDeposit.Height \ vsDeposit.RowHeight(1) - 2)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub vsDeposit_KeyPress(KeyAscii As Integer)
    '���˺�:����Ȩ�޿��ƣ��������Ԥ�����ʣ��������ݲ���ȷ��֮ǰ����������ʱ��Ϊʲô�ܸ��ģ���ʱ��֪��ԭ�������ƣ�����������ݳ���
    If InStr(mstrPrivs, ";����Ԥ������;") > 0 Then Exit Sub
    If Not txtMoney.Visible And KeyAscii <> 13 And KeyAscii <> vbKeyEscape Then
        If mblnThreeDepositAfter And Val(Split(vsDeposit.Cell(flexcpData, vsDeposit.Row, vsDeposit.ColIndex("ID")) & "||", "||")(0)) <> 0 Then
            Exit Sub
        End If
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        With txtMoney
            .Left = vsDeposit.Left + vsDeposit.CellLeft + 30
            .Top = vsDeposit.Top + vsDeposit.CellTop + (vsDeposit.CellHeight - txtMoney.Height) / 2 + 15
            .Width = IIf(vsDeposit.CellWidth - 30 < 0, 50, vsDeposit.CellWidth - 30)
            .ForeColor = vsDeposit.CellForeColor
            .BackColor = vsDeposit.CellBackColor
            .Alignment = 1
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub vsfMoney_DblClick()
     
    If Not txtMoney.Visible And _
        1 = 1 Then
        With vsfMoney
            If InStr(",3,4,9,", "," & .TextMatrix(.Row, .ColIndex("����")) & ",") > 0 Then Exit Sub
            If .RowData(.Row) = 999 Then Exit Sub
            If .Row <= 0 Or .Col <= .ColIndex("���㷽ʽ") Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("�����ID"))) > 0 Then Exit Sub
        End With
        
        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = vsfMoney.Left + vsfMoney.CellLeft + 30
            .Top = vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 + 15
            .Width = vsfMoney.CellWidth - 30
            .ForeColor = vsfMoney.CellForeColor
            .BackColor = IIf(vsfMoney.CellBackColor = 0, vbWhite, vsfMoney.CellBackColor)
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub
Private Sub vsfMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsfMoney
        If .Row < 1 Then Exit Sub
        If .Col = 0 Then
            .Col = .Col + 1: Exit Sub
        End If
        
        If .Row < .Rows - 1 Then
            If Val(.TextMatrix(.Row, .ColIndex("����"))) = 2 Then
               If .Col = .ColIndex("���") Then
                    .Col = .ColIndex("�������")
               Else
                    .Row = .Row + 1
                    .Col = .ColIndex("���")
               End If
            Else
                .Row = .Row + 1
                .Col = .ColIndex("���")
            End If
            
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                 .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            Exit Sub
        End If
    End With
    If GetӦ�� > 0 Then txt�ɿ�.SetFocus: Exit Sub
    cmdOK.SetFocus
    
End Sub

Private Sub vsfMoney_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And KeyAscii <> 13 And KeyAscii <> vbKeyEscape And _
          1 = 1 Then
        With vsfMoney
            If InStr(",3,4,9,", "," & .TextMatrix(.Row, .ColIndex("����")) & ",") > 0 Then Exit Sub
            If .RowData(.Row) = 999 Then Exit Sub
            If .Row <= 0 Or .Col <= .ColIndex("���㷽ʽ") Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("�����ID"))) > 0 Then Exit Sub
        End With
        
        If vsfMoney.Col = 1 Then
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        Else '������������ַ�����,���������ڹ������ж��Ƿ���ҽ�����㷽ʽ
            If InStr("'||,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        
        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = vsfMoney.Left + vsfMoney.CellLeft + 30
            .Top = vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 + 15
            .Width = vsfMoney.CellWidth - 30
            .ForeColor = vsfMoney.CellForeColor
            .BackColor = IIf(vsfMoney.CellBackColor = 0, vbWhite, vsfMoney.CellBackColor)
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = UCase(Chr(KeyAscii))
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Function GetӦ��() As Currency
    Dim i As Long
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("����"))) = 1 Then
                GetӦ�� = Val(.TextMatrix(i, .ColIndex("���")))
                Exit Function
            End If
        Next
    End With
End Function

 
Private Sub txt�ɿ�_Change()
    Call Calc�Ҳ�
End Sub
 
Private Sub txt�ɿ�_GotFocus()
    Dim curTotal As Currency
    Call zlControl.TxtSelAll(txt�ɿ�)
    If Not gblnLED Then Exit Sub
    
    curTotal = GetӦ��
    '#21 1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
    '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
    '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
    zl9LedVoice.DisplayBank ("")
    If curTotal >= 0 Then
        zl9LedVoice.Speak "#21 " & curTotal
    Else
        zl9LedVoice.Speak "#23 " & Abs(curTotal)
    End If
End Sub
Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
        If txt�ɿ�.Text <> "0.00" Then
            If Val(txt�Ҳ�.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                txt�ɿ�.SetFocus
                zlControl.TxtSelAll txt�ɿ�
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '�����ۼӽɿ�
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt�ɿ�_LostFocus()
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
'    If Val(txt�ɿ�.Text) = 0 Then Exit Sub
    
'    If CSng(txt�Ҳ�.Tag) < 0 Then
'        MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
'        Call SelAll(txt�ɿ�): txt�ɿ�.SetFocus
'        Cancel = True: Exit Sub
'    End If
    If Not gblnLED Then Exit Sub
    zl9LedVoice.DispCharge Format(GetӦ��, "0.00"), Val(txt�ɿ�.Text), Val(txt�Ҳ�.Tag)
    zl9LedVoice.Speak "#22 " & txt�ɿ�.Text
    zl9LedVoice.Speak "#23 " & CSng(txt�Ҳ�.Tag)
    zl9LedVoice.Speak "#3"                  '#3  --�뵱�����, лл!
End Sub

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'���ܣ���ָ�����ֱҴ��������д���,���ش����Ľ��
'������curMoney=Ҫ���зֱҴ���Ľ��(ΪӦ�ɽ��,2λС��)
'      mBytMoney=
'         0.������
'         1.��ȡ�������뷨,eg:0.51=0.50;0.56=0.60
'         2.�����շ�,eg:0.51=0.60,0.56=0.60
'         3.����շ�,eg:0.51=0.50,0.56=0.50
'         4.�����������˫,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           �����������˫,����ҹ���ѧ����ίԱ����ʽ�䲼�ġ�������Լ����,������vb��Round����,�������������ְ�����λ����ʱ�����Ը����ֽ���������Լ
'           �����м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ
'         5.�������塢�������,�Խǽ��д�������Ҫ�ȶԷֱҽ�������,��0.29(��)���¶�����ǣ�0.80(��)���϶����ǣ�0.3-0.79����Ϊ0.5��
'         6.��������:eg:0.15=0.10:0.16=0.2:   ���˺� ����:34519  ����:2010-12-06 09:58:02
'91385,������5.�������塢������롱�����ȶԷֱҽ����������룬��0.24(��)���¶�����ǣ�0.75(��)���϶����ǣ�0.25-0.74������Ϊ0.5
'       �ֱ����������룬��ô0.00��0.24=0��0.25��0.5=0.50, 0.50��0.74=0.50��0.75��1.00=1������������ռ50%�ı���

    Dim intSign As Integer, curTmp As Currency

    If mBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf mBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '��ȡ��λ���,�ٴ���ֱ�,��:0.248 ��0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf mBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf mBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf mBytMoney = 4 Then
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
    ElseIf mBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf mBytMoney = 6 Then
         '���˺� ����:34519 ��������:eg:0.15=0.10:0.16=0.2:    ����:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function
 
Private Sub LedDisplayBank()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҽ��������Ϣ
    '����:���˺�
    '����:2013-10-23 14:50:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double, i As Long
    Dim str�����ʻ� As String, strҽ���������� As String, str��һ��ͨ As String, str��ͨ���� As String
    Dim varPara  As Variant, str���㷽ʽ As String
    Dim cur������� As Currency, dbl�ֽ� As Double, dblMoney As Double
    Dim strҽ������ As String
    
    If Not gblnLED Then Exit Sub
    zl9LedVoice.DisplayBank ""
    If mblnFirst = True Then Exit Sub
    
    strҽ������ = "||�ʻ����:" & Format(mcur�������, "0.00")
    With vsfMoney
        For i = 1 To .Rows - 1
            'ҽ������
            str���㷽ʽ = Trim(.TextMatrix(i, 0))
            If str���㷽ʽ <> "" Then
                dblMoney = Val(.TextMatrix(i, 1))
                Select Case Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("����")))
                Case 3
                    str�����ʻ� = str�����ʻ� & "||" & str���㷽ʽ & ":" & Format(dblMoney, "0.00")
                Case 4
                    strҽ���������� = strҽ���������� & "||" & str���㷽ʽ & ":" & Format(dblMoney, "0.00")
                Case 1
                    dbl�ֽ� = dblMoney
                Case Else
                    str��ͨ���� = str��ͨ���� & "||" & str���㷽ʽ & ":" & Format(dblMoney, "0.00")
                End Select
            End If
        Next
    End With
    str���㷽ʽ = ""
    If str�����ʻ� <> "" Then strҽ������ = strҽ������ & str�����ʻ�
    If strҽ���������� <> "" Then strҽ������ = strҽ������ & strҽ����������
    
    If strҽ������ <> "" Then str���㷽ʽ = str���㷽ʽ & "||ҽ������:" & strҽ������
    If str��ͨ���� <> "" Then str���㷽ʽ = str���㷽ʽ & "" & str��ͨ����
    If mcur��Ԥ���ϼ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||" & "��Ԥ��:" & Format(mcur��Ԥ���ϼ�, "0.00")
    
    If str���㷽ʽ = "" Then Exit Sub
    str���㷽ʽ = Mid(str���㷽ʽ, 3)
    varPara = Split(str���㷽ʽ, "||")
    
    dblMoney = Val(txt�ɿ�.Text) - dbl�ֽ�
    zl9LedVoice.DisplayBank "�ܷ���" & Format(txtTotal.Text, "0.00"), "Ԥ����" & Format(mcurԤ���ϼ�, "0.00"), _
            "��Ԥ��" & Format(mcur��Ԥ���ϼ�, "0.00"), IIf(dblMoney > 0, "�Ҳ�", "Ӧ��") & Format(Abs(dblMoney), "0.00")
    'Ŀǰ���ֻ����ʾ10������ֵ
    Select Case UBound(varPara)
    Case 0
          zl9LedVoice.DisplayBank varPara(0)
    Case 1
          zl9LedVoice.DisplayBank varPara(0), varPara(1)
    Case 2
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2)
    Case 3
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3)
    Case 4
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4)
    Case 5
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5)
    Case 6
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6)
    Case 7
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7)
    Case 8
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8)
    Case 9
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9)
    Case Else
        str���㷽ʽ = ""
         For i = 10 To UBound(varPara)
            str���㷽ʽ = str���㷽ʽ & ";" & varPara(i)
        Next
        If str���㷽ʽ > "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str���㷽ʽ
    End Select
    'zl9LedVoice.Speak "#21 " & Format(mcurӦ�ɽ��, "0.00")
End Sub

 
Private Sub txt�Ҳ�_Change()
    txt�Ҳ�.Tag = ""
End Sub
Private Sub Calc�Ҳ�()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����Ҳ�
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-01-12 17:41:47
    '����:27360
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�Ҳ� As Double
    Dim cur�ֽ� As Currency, i As Long

    If Val(txt�ɿ�.Text) = 0 Then txt�Ҳ�.Text = "0.00"
    dbl�Ҳ� = FormatEx(Val(txt�ɿ�.Text) - GetӦ��, 2)
    txt�Ҳ�.Text = Format(Abs(dbl�Ҳ�), "0.00")
    txt�Ҳ�.Tag = dbl�Ҳ�
    If dbl�Ҳ� <= 0 Then
        lbl�Ҳ�.Caption = "�տ�"
        lbl�Ҳ�.ForeColor = &H0&
    Else
        lbl�Ҳ�.Caption = "�Ҳ�"
        lbl�Ҳ�.ForeColor = vbRed   '35830
    End If
    txt�Ҳ�.ForeColor = lbl�Ҳ�.ForeColor
End Sub
Private Function GetThreePayDepositData(ByRef rsTemp As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������Ϣ
    '����:rsTemp-���ؽ�����Ϣ(�����ID,���������,���㷽ʽ,�Ƿ�����,���,��Ԥ��,ʣ���,��Ԥ��)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-04-27 09:44:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl��Ԥ�� As Double, dblMoney As Double, dbl��� As Double
    Dim dblTotal As Double, dblTemp As Double, lngCardTypeID As Long
    Dim varData As Variant
    
    On Error GoTo errHandle
    Set rsTemp = New ADODB.Recordset
    rsTemp.Fields.Append "�����ID", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "���������", adVarChar, 200, adFldIsNullable
    rsTemp.Fields.Append "���㷽ʽ", adVarChar, 100, adFldIsNullable
    rsTemp.Fields.Append "�Ƿ�����", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "�Ƿ�����", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "ȱʡ����", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "���", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "��Ԥ��", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "ʣ���", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "��Ԥ��", adDouble, , adFldIsNullable
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    If mrsCardType Is Nothing Then
        Call initCardSquareData
    ElseIf mrsCardType.State <> 1 Then
        Call initCardSquareData
    End If
    
    dblTotal = Val(txtTotal.Text) - mcurMediCare
    With vsDeposit
        dblMoney = 0: dbl��Ԥ�� = 0: dbl��� = 0: lngCardTypeID = 0
        For i = 1 To .Rows - 1
            ' �����ID|| ת�ʼ�����||����||ȱʡ����
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||||", "||")
            If Val(varData(0)) <> 0 And Val(varData(3)) = 0 Then
                
                lngCardTypeID = Val(varData(0))
                dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                rsTemp.Find "�����ID=" & lngCardTypeID
                mrsCardType.Filter = "ID=" & lngCardTypeID
                If rsTemp.EOF Then
                    rsTemp.AddNew
                    rsTemp!�����ID = lngCardTypeID
                    If Not mrsCardType.EOF Then
                       rsTemp!��������� = mrsCardType!����
                       rsTemp!�Ƿ����� = Val(Nvl(mrsCardType!�Ƿ�����))
                    Else
                       rsTemp!��������� = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                       rsTemp!�Ƿ����� = 0
                    End If
                    rsTemp!���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                    rsTemp!��Ԥ�� = 0
                End If
                rsTemp!��� = FormatEx(Val(Nvl(rsTemp!���)) + Val(.TextMatrix(i, .ColIndex("���"))), 5)
                rsTemp!��Ԥ�� = FormatEx(Val(Nvl(rsTemp!��Ԥ��)) + dbl��Ԥ��, 5)
                rsTemp!ʣ��� = FormatEx(Val(Nvl(rsTemp!���)) - Val(Nvl(rsTemp!��Ԥ��)), 5)
                If FormatEx(dblTotal - dbl��Ԥ��, 6) < 0 Then
                    If dblTotal >= 0 Then
                        dblTemp = dbl��Ԥ�� - dblTotal
                        rsTemp!��Ԥ�� = FormatEx(Val(Nvl(rsTemp!��Ԥ��)) + dblTemp, 5)
                    Else
                        rsTemp!��Ԥ�� = FormatEx(Val(Nvl(rsTemp!��Ԥ��)) + dbl��Ԥ��, 5)
                    End If
                    dblTotal = 0
                Else
                    dblTotal = FormatEx(dblTotal - dbl��Ԥ��, 6)
                End If
                rsTemp.Update
            End If
        Next
    End With
    GetThreePayDepositData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowDelThreeSwap()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��֧������Ϣ
    '����:���˺�
    '����:2015-04-27 11:09:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strTittle As String
    Dim intCount  As Integer, intNotCashCount  As Integer
    
    On Error GoTo errHandle
    Call GetDelThreeCardDepositInfor(intCount, intNotCashCount, mblnThreeDepositAfter, mstrStyle)


    
    lblDelMoney.Visible = False
    cmdReturnCash.Visible = False
    lblDelMoney.Tag = "0"
    
    If mstrForceNote <> "" Then
        mblnThreeDepositAfter = False
        GoTo BrushWin
    End If
    If mblnThreeDepositAfter Then
    
    
    
        lblDelMoney.Caption = IIf(mstrStyle <> "", "�ˣ�" & mstrStyle, "")
        lblDelMoney.Visible = True
        cmdReturnCash.Visible = lblDelMoney.Visible And mstrStyle <> ""
        Exit Sub
    End If
    If GetThreePayDepositData(rsTemp) = False Then GoTo BrushWin
    
    '�޼�¼ʱ,��ʾ��������������,ֱ�ӷ���true
    If rsTemp.RecordCount = 0 Then GoTo BrushWin
    rsTemp.Filter = "��Ԥ��<>0"
    If rsTemp.RecordCount = 0 Then GoTo BrushWin
    strTittle = ""
    Do While Not rsTemp.EOF
         strTittle = strTittle & IIf(strTittle = "", "", vbCrLf) & "��" & Nvl(rsTemp!���������) & ":" & Format(Val(Nvl(rsTemp!��Ԥ��)), "0.00")
         lblDelMoney.Tag = FormatEx(Val(lblDelMoney.Tag) + Val(Nvl(rsTemp!��Ԥ��)), 6)
         rsTemp.MoveNext
    Loop
    lblDelMoney.Caption = strTittle
    lblDelMoney.Visible = True
    lblDelMoney.Top = cmdReturnCash.Top + (cmdReturnCash.Height - lblDelMoney.Height) \ 2
    cmdReturnCash.Visible = lblDelMoney.Visible
BrushWin:
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CheckThreePayDepositValied(ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������Ԥ���ĺϷ���
    '����:����֧������(������)
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-04-23 17:32:37
    '����:
    '     1)Ŀǰֻ֧�������ʻ��д���(ת�ʽ��׽ӿڵ�)
    '     2)����ͬʱ����2�����ϵ������ʻ����׵�,���ڵĻ�����False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strMsg As String
    Dim dblTotal As Double, dblMoney As Double
    Set objCard = Nothing
    If mblnThreeDepositAfter Or mstrForceNote <> "" Then
        CheckThreePayDepositValied = True
        Exit Function
    End If
    mCurBrushCard.dblMoney = 0
    If GetThreePayDepositData(rsTemp) = False Then Exit Function
    '�޼�¼ʱ,��ʾ��������������,ֱ�ӷ���true
    If rsTemp.RecordCount = 0 Then CheckThreePayDepositValied = True: Exit Function
    rsTemp.Filter = "��Ԥ��<>0"
    If rsTemp.RecordCount = 0 Then CheckThreePayDepositValied = True: Exit Function
    
    If rsTemp.RecordCount >= 2 Then
       Do While Not rsTemp.EOF
            strMsg = strMsg & vbCrLf & Nvl(rsTemp!���������) & ":" & Format(Nvl(rsTemp!��Ԥ��), "0.00")
            rsTemp.MoveNext
       Loop
       MsgBox "��ǰ����" & rsTemp.RecordCount & "������������Ҫ�˿�,Ŀǰ,ϵͳֻ֧��һ�����������˿�(��Ϊ���۽��׵�)," & _
              "" & "����Ϊ��ǰ��Ҫ�˿����������:" & _
              strMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Val(Nvl(rsTemp!�Ƿ�����)) = 0 Then
       MsgBox Nvl(rsTemp!���������) & "δ���ã��������˿�!" & _
              "", vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If
    If Not GetCurCard(Val(Nvl(rsTemp!�����ID)), objCard) Then
       MsgBox Nvl(rsTemp!���������) & "δ���û��ȡʧ�ܣ��������˿�!", vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If

    
    dblMoney = FormatEx(Val(Nvl(rsTemp!��Ԥ��)), 6)
    mCurBrushCard.dblMoney = dblMoney
    
    If dblMoney <> FormatEx(Val(lblDelMoney.Tag), 6) Then
       If MsgBox(Nvl(rsTemp!���������) & "�н�����δ���(" & lblDelMoney.Tag & ")�뵱ǰ�˿���(" & dblMoney & ") ��һ��!" & vbCrLf & "���Ƿ�����ˢ�½�����˿���!", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
           Call ShowDelThreeSwap
       End If
       Exit Function
    End If
    
    If CheckThreeSwapIsValied(objCard, dblMoney) = False Then Exit Function
    CheckThreePayDepositValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCurCard(ByVal lngCardTypeID As Long, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ������
    '���:lngCardTypeID-��ǰ�����ID
    '����:objCard-���ص�ǰ�˿��ɿ�Ŀ�����
    '����:�ɹ�,���ؿ�����
    '����:���˺�
    '����:2015-04-27 10:32:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card
    On Error GoTo errHandle
    Set objCard = Nothing
    For Each objTemp In mobjPayCards
        If objTemp.�ӿ���� = lngCardTypeID And Not objTemp.���ѿ� Then
            Set objCard = objTemp
            GetCurCard = True: Exit Function
        End If
    Next
    GetCurCard = False
    Exit Function
errHandle:
End Function

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����֤
    '���:objCard-��ǰ��
    '     dblMoney-�˿����
    '����:ˢ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 15:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, dbl�ʻ���� As Double
    Dim cllSquareBalance As New Collection
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim strXmlIn As String
    
    On Error GoTo errHandle
    
    If objCard.�ӿ���� <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If mlng����ID <> 0 Then
        strSql = "Select ����,�Ա�,���� From ������Ϣ Where ����ID=[1]"
    Else
        strSql = "Select ����,�Ա�,���� From ������Ϣ A,���˽��ʼ�¼ B Where  A.����ID=B.����ID and B.ID=[2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng����ID)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ�ָ���Ĳ���,���ܵ��������ӿڽ���", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '   zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByRef dbl��� As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String, _
    Optional ByRef bln�˷� As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln���� As Boolean = False, _
    Optional ByVal bln�����ֹ As Boolean = True, _
    Optional ByRef varSquareBalance As Variant) As Boolean
    Dim strCardNo As String, strPassWord As String
    strXmlIn = "<IN><CZLX>0</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
        objCard.�ӿ����, False, _
    rsTemp!����, Nvl(rsTemp!�Ա�), Nvl(rsTemp!����), dblMoney, strCardNo, strPassWord, _
    False, True, False, False, cllSquareBalance, False, strXmlIn) = False Then Exit Function
    mCurBrushCard.str���� = strCardNo
    mCurBrushCard.str���� = strPassWord
    
    '����ת�ʽӿ�
    '    7.1.    zltransferAccountsCheck(ת�ʼ��ӿ�)
    'zlTransferAccountsCheck ת�ʼ��ӿ�
    '������  ��������    ��/��   ��ע
    'frmMain Object  In  ���õ�������
    'lngModule   Long    In  HIS����ģ���
    'lngCardTypeID   Long    In  �����ID
    'strCardNo   String  In  ����
    'dblMoney    Double  In  ת�ʽ��(����ʱΪ����)
    'strBalanceIDs   String  In  ����IDs������ö��ŷ��룬��ʾ���ζ��Ĵ��շ���Ŀ��������ҽ��������
    'strXMLExpend String In   XML��:
    '                            <IN>
    '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�˷�ҵ��2-����ҵ��;3-�����˷�ҵ��
    '                            </IN>
    '                    Out  XML��:
    '                            <OUT>
    '                               <ERRMSG>������Ϣ</ERRMSG >
    '                            </OUT>
    '    Boolean ��������    �������ݺϷ�,����True:���򷵻�False
    '˵��:
    '��. ��ҽ���������ʱ���е�����ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
    '��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
    '����XML��
    If objCard.�Ƿ�ת�ʼ����� Then
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "2"
        zlXML.AppendNode "IN", True
        strXMLExpend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, objCard.�ӿ����, _
            mCurBrushCard.str����, dblMoney, "", strXMLExpend) = False Then
            Call ShowErrMsg(0, strXMLExpend)
            Exit Function
        End If
    End If
                    
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        strCardNo-����
    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
    '����:dblMoney-�����ʻ����
    Dim strExpand As String
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.�ӿ����, _
          mCurBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�)
    mCurBrushCard.dbl�ʻ���� = FormatEx(dbl�ʻ����, 2)
    
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ExecuteThreeSwapPayInterface(objCard As Card, ByVal lng����ID As Long, _
      ByVal dblMoney As Double, Optional ByVal blnMustCommit As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�����ӿ�)
    '���:objCard-��ǰ��������
    '     lng����ID-����ID
    '     dblMoney-����֧�����
    '     blnMustCommit-�����ύ(��Ҫ��ҽ���ӿ�)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-04-27 10:45:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String, strXMLExpend As String
    Dim cllPro As Collection, blnTrans As Boolean, rsTmp As ADODB.Recordset, strCardNo As String
    Dim i As Long, strSql As String, lngID As Long, varData As Variant, strExpend As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection, strInXML As String, strOutXML As String
    Dim objXml As New clsXML, dblCheck As Double, dbl��Ԥ�� As Double, lngRow As Long, strValue As String
    dblCheck = dblMoney
    
    blnTrans = True
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    If objCard Is Nothing Then
        gcnOracle.CommitTrans
        ExecuteThreeSwapPayInterface = True: Exit Function
    End If
    
    If mblnThreeDepositAfter Or mstrForceNote <> "" Then
        gcnOracle.CommitTrans
        ExecuteThreeSwapPayInterface = True: Exit Function
    End If
        
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Then gcnOracle.CommitTrans: ExecuteThreeSwapPayInterface = True: Exit Function
    If objCard.�Ƿ�ת�ʼ����� Then
        'zlTransferAccountsMoney
        '������  ��������    ��/��   ��ע
        'frmMain Object  In  ���õ�������
        'lngModule   Long    In  HIS����ģ���
        'lngCardTypeID   Long    In  �����ID
        'strCardNo   String  In  ����
        'strBalanceID    String  In  ����ID
        'dblMoney    Double  In  ת�ʽ��
        'strSwapGlideNO  String  Out ������ˮ��
        'strSwapMemo String  Out ����˵��
        'strSwapExtendInfor  String  Out ������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
        'strXMLExpend String In   XML��:
        '                            <IN>
        '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�˷�ҵ��2-����ҵ��;3-�����˷�ҵ��
        '                            </IN>
        '                    Out  XML��:
        '                            <OUT>
        '                               <ERRMSG>������Ϣ</ERRMSG >
        '                            </OUT>
        '    Boolean ��������    True:���óɹ�,False:����ʧ��
        '˵��:
        '��. ��ҽ���������ʱ���е�����ת��ʱ���á�
        '��. һ����˵���ɹ�ת�ʺ󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
        '��. ��ת�ʳɹ��󣬷��ؽ�����ˮ�ź���ؽ���˵���������������������Ϣ�����Է�����չ��Ϣ�з���.
        '����XML��
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "2"
        zlXML.AppendNode "IN", True
        strXMLExpend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, objCard.�ӿ����, mCurBrushCard.str����, _
            lng����ID, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
            If Not blnMustCommit Then   'ҽ�������ύ��������ݲ���Ԥ����¼�е�У�Ա�־��ȷ��
                gcnOracle.RollbackTrans:
            Else
                gcnOracle.CommitTrans
                ExecuteThreeSwapPayInterface = True
            End If
            Call ShowErrMsg(1, strXMLExpend)
            blnTrans = False
            Exit Function
        End If
        
        mCurBrushCard.str������ˮ�� = strSwapGlideNO
        mCurBrushCard.str����˵�� = strSwapMemo
        Call zlAddUpdateSwapSQL(False, lng����ID, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
        Call zlAddThreeSwapSQLToCollection(False, lng����ID, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    Else
        objXml.ClearXmlText
        
        With vsDeposit
            Call objXml.AppendNode("JSLIST")
            For i = .Rows - 1 To 1 Step -1
                '�����ID||���۱�־
                varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||", "||")
                If Val(varData(0)) <> 0 And Val(varData(2)) = 0 And dblCheck > 0 Then
                    dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                    If dblCheck >= dbl��Ԥ�� Then
                        lngID = .TextMatrix(i, .ColIndex("Ԥ��ID"))
                        strSql = "Select ID,����,������ˮ��,����˵�� From ����Ԥ����¼ Where ID = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
                        If Not rsTmp.EOF Then
                            Call objXml.AppendNode("JS")
                                Call objXml.appendData("KH", Nvl(rsTmp!����))
                                Call objXml.appendData("JYLSH", Nvl(rsTmp!������ˮ��))
                                Call objXml.appendData("JYSM", Nvl(rsTmp!����˵��))
                                Call objXml.appendData("ZFJE", dbl��Ԥ��)
                                Call objXml.appendData("JSLX", 1)
                                Call objXml.appendData("ID", Nvl(rsTmp!ID))
                            Call objXml.AppendNode("JS", True)
                        End If
                        strSql = "Zl_�����˿���Ϣ_Insert("
                        strSql = strSql & lng����ID & ","
                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                        strSql = strSql & dbl��Ԥ�� & ",'"
                        strSql = strSql & Nvl(rsTmp!����) & "','"
                        strSql = strSql & Nvl(rsTmp!������ˮ��) & "','"
                        strSql = strSql & Nvl(rsTmp!����˵��) & "')"
                        zlAddArray cllThreeSwap, strSql
                        dblCheck = dblCheck - dbl��Ԥ��
                    Else
                        lngID = .TextMatrix(i, .ColIndex("Ԥ��ID"))
                        strSql = "Select ID,����,������ˮ��,����˵�� From ����Ԥ����¼ Where ID = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
                        If Not rsTmp.EOF Then
                            Call objXml.AppendNode("JS")
                                Call objXml.appendData("KH", Nvl(rsTmp!����))
                                Call objXml.appendData("JYLSH", Nvl(rsTmp!������ˮ��))
                                Call objXml.appendData("JYSM", Nvl(rsTmp!����˵��))
                                Call objXml.appendData("ZFJE", dblCheck)
                                Call objXml.appendData("JSLX", 1)
                                Call objXml.appendData("ID", Nvl(rsTmp!ID))
                            Call objXml.AppendNode("JS", True)
                        End If
                        strSql = "Zl_�����˿���Ϣ_Insert("
                        strSql = strSql & lng����ID & ","
                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                        strSql = strSql & dblCheck & ",'"
                        strSql = strSql & Nvl(rsTmp!����) & "','"
                        strSql = strSql & Nvl(rsTmp!������ˮ��) & "','"
                        strSql = strSql & Nvl(rsTmp!����˵��) & "')"
                        zlAddArray cllThreeSwap, strSql
                        dblCheck = 0
                    End If
                End If
            Next i
            Call objXml.AppendNode("JSLIST", True)
        End With
    
        strInXML = objXml.XmlText
        
        If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, strInXML, _
             lng����ID, strOutXML, strExpend) = False Then
            If Not blnMustCommit Then   'ҽ�������ύ��������ݲ���Ԥ����¼�е�У�Ա�־��ȷ��
                gcnOracle.RollbackTrans:
            Else
                gcnOracle.CommitTrans
                ExecuteThreeSwapPayInterface = True
            End If
            Call ShowErrMsg(1, strXMLExpend)
            blnTrans = False
            Exit Function
        End If
             
        If strOutXML <> "" Then
            If zlXML_Init = False Then Exit Function
            If zlXML_LoadXMLToDOMDocument(strOutXML, False) = False Then Exit Function
            Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("ID", i, strValue)
                strSql = "Zl_�����˿���Ϣ_Insert("
                strSql = strSql & lng����ID & ","
                strSql = strSql & Val(strValue) & ","
                strSql = strSql & 0 & ",'"
                Call zlXML_GetNodeValue("KH", i, strValue)
                strSql = strSql & strValue & "','"
                Call zlXML_GetNodeValue("TKLSH", i, strValue)
                strSql = strSql & strValue & "','"
                Call zlXML_GetNodeValue("TKSM", i, strValue)
                strSql = strSql & strValue & "',"
                strSql = strSql & 1 & ")"
                zlAddArray cllThreeSwap, strSql
            Next i
        End If
        
        If strExpend <> "" Then
            strSwapExtendInfor = ""
            If zlXML_LoadXMLToDOMDocument(strExpend, False) = False Then Exit Function
            Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("XMMC", i, strValue)
                strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
                Call zlXML_GetNodeValue("XMNR", i, strValue)
                strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
            Next i
        End If
        If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
        strSql = "Select ���� From ����Ԥ����¼ Where ����ID= [1] And �����ID= [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, objCard.�ӿ����)
        If Not rsTmp.EOF Then
            strCardNo = Nvl(rsTmp!����)
        End If
        Call zlAddUpdateSwapSQL(False, lng����ID, objCard.�ӿ����, objCard.���ѿ�, strCardNo, "", "", cllUpdate, 0)
        Call zlAddThreeSwapSQLToCollection(False, lng����ID, objCard.�ӿ����, objCard.���ѿ�, strCardNo, strSwapExtendInfor, cllThreeSwap)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    End If

    Err = 0: On Error GoTo ErrOtherHand:
    '��������������Ϣ
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    ExecuteThreeSwapPayInterface = True
    Exit Function
ErrOtherHand:
    ExecuteThreeSwapPayInterface = True
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ShowErrMsg(ByVal BytType As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ת�˼�������ҵ�������ʾ
    '����:Ƚ����
    'ʱ��:2014-12-2
    '����:
    '   bytType:0-ת�˼��,1-ת�˽���
    '   strXMLErrMsg:��ʽ����
    '            <OUT>
    '               <ERRMSG>������Ϣ</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    '����������Ϣ
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '��ʾ������Ϣ
    If Trim(strValue) = "" Then
        If BytType = 0 Then
            strValue = vbCrLf & "���׼��ʧ�ܣ�"
        Else
            strValue = vbCrLf & "����ʧ�ܣ�"
        End If
    End If
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
     
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then
        Call CreateSquareCardObject(Me, mlngModul)
        If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    End If
    Set mrsCardType = gobjSquare.objSquareCard.zlGetYLCards
    '�������õ������ʻ�
    Set mobjPayCards = gobjSquare.objSquareCard.zlGetCards(3)
End Sub
 

Private Function GetDelThreeCardDepositInfor(ByRef intThreeCount As Integer, ByRef intNotDelCashCount As Integer, _
    ByRef blnThreeDepositAfter As Boolean, ByRef strDelThreeNames As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ʻ���Ԥ�������Ϣ
    '���:
    '����:intNotDelCashCount-���ز��������ֵĸ���
    '     intThreeCount-�����ʻ�����
    '     blnThreeDepositAfter-�����ʻ�������(true:��������˿�,False-����������˿�)
    '     strDelThreeNames-���������ʻ�����˿�����ƴ������磺����,����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-10-25 11:59:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double, varData As Variant, dbl��Ԥ�� As Double
    Dim strStyle As String, i As Long
    
    On Error GoTo errHandle
    
    blnThreeDepositAfter = False
    
    dblTotal = RoundEx(Val(txtTotal.Text) - mcurMediCare, 2)
    If mrsCardType Is Nothing Then
        Call initCardSquareData
    ElseIf mrsCardType.State <> 1 Then
        Call initCardSquareData
    End If
    
    intNotDelCashCount = 0
    intThreeCount = 0
    With vsDeposit
        For i = 1 To .Rows - 1
            dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
            ' �����ID ||����||����||ȱʡ����
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||||", "||")
            If Val(varData(0)) <> 0 Then
                If Val(varData(3)) = 0 And ((dblTotal - dbl��Ԥ��) <= 0 Or dbl��Ԥ�� = 0) Then   '��ȱʡ����
                    mrsCardType.Filter = "ID=" & Val(varData(0))
                    intThreeCount = intThreeCount + 1
                    If Not mrsCardType.EOF Then
                        If InStr(strStyle & ",", "," & Nvl(mrsCardType!����) & ",") = 0 Then
                            strStyle = strStyle & "," & mrsCardType!����
                            If Val(varData(2)) = 0 Then
                               intNotDelCashCount = intNotDelCashCount + 1
                            End If
                        End If
                    End If
                End If
            End If
            
            If FormatEx(dblTotal - dbl��Ԥ��, 6) <= 0 Then
                dblTotal = 0
            Else
                dblTotal = FormatEx(dblTotal - dbl��Ԥ��, 6)
            End If
        Next
    End With
    
    
    If intThreeCount >= 1 And InStr(1, mstrPrivs, ";����Ԥ������;") = 0 Then blnThreeDepositAfter = True
    
    If strStyle <> "" Then strStyle = Mid(strStyle, 2)
    strDelThreeNames = strStyle

    GetDelThreeCardDepositInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


