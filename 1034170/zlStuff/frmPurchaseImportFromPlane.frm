VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchaseImportFromPlane 
   Caption         =   "����ƻ���"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12210
   Icon            =   "frmPurchaseImportFromPlane.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12210
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   3855
      TabIndex        =   18
      Top             =   7080
      Width           =   3855
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1680
         TabIndex        =   22
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "��ͣ��"
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.CheckBox chk������ͣ������ 
      Caption         =   "������ͣ������"
      Height          =   180
      Left            =   7440
      TabIndex        =   15
      Top             =   7045
      Width           =   1815
   End
   Begin VB.Frame frmCondition 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   12012
      Begin VB.ComboBox cboStock 
         Height          =   276
         Left            =   8700
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   205
         Visible         =   0   'False
         Width           =   1872
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   195
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   127467523
         CurrentDate     =   36263
      End
      Begin VB.CheckBox chkNoTime 
         Caption         =   "����"
         Height          =   180
         Left            =   1440
         TabIndex        =   14
         Tag             =   "1|0"
         Top             =   262
         Width           =   735
      End
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   6120
         MaxLength       =   8
         TabIndex        =   6
         Top             =   193
         Width           =   1605
      End
      Begin VB.CommandButton cmd��ȡ 
         Caption         =   "��ȡ(&G)"
         Height          =   350
         Left            =   10800
         TabIndex        =   5
         Top             =   168
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   4080
         TabIndex        =   8
         Top             =   195
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   127467523
         CurrentDate     =   36263
      End
      Begin VB.Label lbl����ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ⷿ"
         Height          =   180
         Left            =   7920
         TabIndex        =   17
         Top             =   253
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ƻ����������"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   260
         Width           =   1260
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   3840
         TabIndex        =   10
         Top             =   255
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No"
         Height          =   180
         Left            =   5880
         TabIndex        =   9
         Top             =   252
         Width           =   180
      End
   End
   Begin VB.PictureBox picLine 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "������ⵥ(&O)"
      Height          =   350
      Left            =   9480
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   10920
      TabIndex        =   0
      Top             =   6960
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7428
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseImportFromPlane.frx":030A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16457
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2208
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "˫�����ݣ�ѡ��Ҫ����ļƻ�����"
      Top             =   840
      Width           =   12012
      _cx             =   21188
      _cy             =   3895
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
      BackColorSel    =   16764622
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchaseImportFromPlane.frx":0B9E
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
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   2772
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   12012
      _cx             =   21188
      _cy             =   4890
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
      BackColorSel    =   16764622
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchaseImportFromPlane.frx":0D0B
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
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "ע�⣺δ���ù�Ӧ�̵����Ľ����ᵼ�룡"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   3240
   End
End
Attribute VB_Name = "frmPurchaseImportFromPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSum As Long '��¼�������ļƻ�����δ����ͣ�����ĸ���
Private mstrMsg As String '�������ļƻ�����ͣ������δ����ʱ����ʾ��Ϣ

'�����洫�����
Dim mfrmMain As Form
Dim mStr�ⷿ As String
Dim mlng�ⷿid As Long
Dim mintUnit As Integer                 '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Dim mbln���пⷿ As Boolean
Dim mblnSuccess As Boolean
Private mint��ѯ��ʽ As Integer     '���������ǲ�ѯ�ƻ��������깺��:0-�ƻ���;1-�깺��
Private mlngMode As Long
Private mint����� As Integer             '��ʾ���ĳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mint��ȷ���� As Integer             '�����Ƿ����γ���
Private mint������ʾ As Integer             '��Զ��۲��������ĸ�����������

'��������
Dim mOraFMT As g_FmtString

Private Sub ��ȡ���ĳɱ��ۺ��ۼ�(ByRef rsData As ADODB.Recordset, ByVal lng����ID As Long, ByVal lng���� As Long, _
                                ByVal int�Ƿ���� As Integer, ByVal bln�Ƿ��� As Boolean, ByVal dbl����ϵ�� As Double)
    Dim rsprice As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If mint��ȷ���� = 1 Then
        gstrSQL = "Select a.ʵ������, a.ʵ�ʽ��, a.ʵ�ʲ��, a.���ۼ�, a.ƽ���ɱ���, b.�ּ�, c.�ɱ���" & vbNewLine & _
                "From ҩƷ��� A, �շѼ�Ŀ B, �������� C" & vbNewLine & _
                "Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.����id And a.ҩƷid = [1] And Nvl(a.����, 0) = [2] And b.ִ������ <= Sysdate And" & vbNewLine & _
                "      b.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd') And Rownum < 2"
    Else
        gstrSQL = "Select a.ʵ������, a.ʵ�ʽ��, a.ʵ�ʲ��, a.���ۼ�, a.ƽ���ɱ���, b.�ּ�, c.�ɱ���" & vbNewLine & _
                "From ҩƷ��� A, �շѼ�Ŀ B, �������� C" & vbNewLine & _
                "Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.����id And a.ҩƷid = [1] And b.ִ������ <= Sysdate And" & vbNewLine & _
                "      b.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd') And Rownum < 2"
    End If
         
    Set rsprice = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ĳɱ��ۺ��ۼ�", lng����ID, lng����)
    
    If mint��ȷ���� = 1 Then
        If int�Ƿ���� = 1 Then
            If bln�Ƿ��� = False Then     '���۷���
                rsData!�ɱ��� = IIf(IsNull(rsprice!ƽ���ɱ���), 0, rsprice!ƽ���ɱ���) * dbl����ϵ��
                rsData!�ۼ� = IIf(IsNull(rsprice!�ּ�), 0, rsprice!�ּ�) * dbl����ϵ��
            Else                            'ʱ�۷���
                rsData!�ɱ��� = IIf(IsNull(rsprice!ƽ���ɱ���), 0, rsprice!ƽ���ɱ���) * dbl����ϵ��
                rsData!�ۼ� = IIf(IsNull(rsprice!���ۼ�), 0, rsprice!���ۼ�) * dbl����ϵ��
            End If
            
            rsData!�ɱ���� = rsData!ʵ������ * rsData!�ɱ���
            rsData!�ۼ۽�� = rsData!ʵ������ * rsData!�ۼ�
        Else
            If bln�Ƿ��� = False Then     '���۲�����
                rsData!�ɱ��� = IIf(IsNull(rsprice!ƽ���ɱ���), 0, rsprice!ƽ���ɱ���) * dbl����ϵ��
                rsData!�ۼ� = IIf(IsNull(rsprice!�ּ�), 0, rsprice!�ּ�) * dbl����ϵ��
            Else                            'ʱ�۲�����
                rsData!�ɱ��� = IIf(IsNull(rsprice!ƽ���ɱ���), 0, rsprice!ƽ���ɱ���) * dbl����ϵ��
                rsData!�ۼ� = rsprice!ʵ�ʽ�� / rsprice!ʵ������ * dbl����ϵ��
            End If
            
            rsData!�ɱ���� = rsData!ʵ������ * rsData!�ɱ���
            rsData!�ۼ۽�� = rsData!ʵ������ * rsData!�ۼ�
        End If
    Else
        If int�Ƿ���� = 1 Then
            If bln�Ƿ��� = False Then     '���۷���
                rsData!�ɱ��� = IIf(IsNull(rsprice!�ɱ���), 0, rsprice!�ɱ���) * dbl����ϵ��
                rsData!�ۼ� = IIf(IsNull(rsprice!�ּ�), 0, rsprice!�ּ�) * dbl����ϵ��
            Else                            'ʱ�۷���
                rsData!�ɱ��� = IIf(IsNull(rsprice!�ɱ���), 0, rsprice!�ɱ���) * dbl����ϵ��
                rsData!�ۼ� = IIf(IsNull(rsprice!�ּ�), 0, rsprice!�ּ�) * dbl����ϵ��
            End If
            
            rsData!�ɱ���� = rsData!ʵ������ * rsData!�ɱ���
            rsData!�ۼ۽�� = rsData!ʵ������ * rsData!�ۼ�
        Else
            If bln�Ƿ��� = False Then     '���۲�����
                rsData!�ɱ��� = IIf(IsNull(rsprice!ƽ���ɱ���), 0, rsprice!ƽ���ɱ���) * dbl����ϵ��
                rsData!�ۼ� = IIf(IsNull(rsprice!�ּ�), 0, rsprice!�ּ�) * dbl����ϵ��
            Else                            'ʱ�۲�����
                rsData!�ɱ��� = IIf(IsNull(rsprice!ƽ���ɱ���), 0, rsprice!ƽ���ɱ���) * dbl����ϵ��
                rsData!�ۼ� = rsprice!ʵ�ʽ�� / rsprice!ʵ������ * dbl����ϵ��
            End If
            
            rsData!�ɱ���� = rsData!ʵ������ * rsData!�ɱ���
            rsData!�ۼ۽�� = rsData!ʵ������ * rsData!�ۼ�
        End If
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function �����(ByVal lng�ⷿID As Long, ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln������� As Boolean
    Dim bln�ⷿ���� As Boolean
    Dim bln���÷��� As Boolean
    Dim bln�Ƿ��� As Boolean
    
    ����� = False
    On Error GoTo ErrHandle
    
    '����ǲ��������ģ��Ҳ�����棬��ֱ���˳��˹���
    '---------------------------------------
    '��ȡ��ǰ���ķ������
    gstrSQL = "Select Nvl(a.�ⷿ����, 0) �ⷿ����, Nvl(a.���÷���, 0) ���÷���, b.�Ƿ���" & vbNewLine & _
            "From �������� A, �շ���ĿĿ¼ B" & vbNewLine & _
            "Where a.����id = b.Id And a.����id = [1]"

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ���ķ������", lng����ID)
    bln�ⷿ���� = rsTemp!�ⷿ����
    bln���÷��� = rsTemp!���÷���
    bln�Ƿ��� = rsTemp!�Ƿ���
    
    '��ȡ����ⷿ��������
    gstrSQL = "Select 1 From ��������˵�� Where �������� In '���ϲ���' And ����id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ���ķ������", lng�ⷿID)
    If rsTemp.EOF Then
        bln������� = bln�ⷿ����
    Else
        bln������� = bln���÷���
    End If

    If bln������� = False And mint����� <> 2 And bln�Ƿ��� = False Then
        ����� = True
        If mint����� = 1 Then
            mint������ʾ = 1
        End If
        Exit Function
    End If
    '---------------------------------------
    
    '���û�п���¼����ֱ���˳�
    gstrSQL = "" & _
        "   Select Count(*) ��¼�� From ҩƷ��� " & _
        "   Where �ⷿID=[1] And ����=1 And ҩƷID=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����������Ƿ����", lng�ⷿID, lng����ID)
    If rsTemp!��¼�� <> 0 Then
        ����� = True
        Exit Function
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ���ķֽ�(ByRef rsData As ADODB.Recordset, ByVal lng�ⷿID As Long, ByVal lng����ID As Long, _
                    ByVal dbl��д���� As Double, ByVal dbl����ϵ�� As Double)

    Dim rsTemp As New ADODB.Recordset
    Dim dbl�������� As Double
    Dim dblʣ������ As Double
    Dim bln������� As Boolean
    Dim bln�ⷿ���� As Boolean
    Dim bln���÷��� As Boolean
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl�ۼ� As Double
    Dim dbl�ۼ۽�� As Double
    Dim str���� As String
    Dim lng����ID As Long
    Dim bln�Ƿ��� As Boolean
          
    On Error GoTo ErrHandle
    
    '��������
    str���� = rsData!����
    lng�ⷿID = rsData!�ⷿid
    lng����ID = rsData!����ID
    
    '��ȡ��ǰ���ķ������
    gstrSQL = "Select Nvl(a.�ⷿ����, 0) �ⷿ����, Nvl(a.���÷���, 0) ���÷���, b.�Ƿ���" & vbNewLine & _
            "From �������� A, �շ���ĿĿ¼ B" & vbNewLine & _
            "Where a.����id = b.Id And a.����id = [1]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ���ķ������", lng����ID)
    bln�ⷿ���� = rsTemp!�ⷿ����
    bln���÷��� = rsTemp!���÷���
    bln�Ƿ��� = rsTemp!�Ƿ���
    
    '��ȡ����ⷿ��������
    gstrSQL = "Select 1 From ��������˵�� Where �������� In '���ϲ���' And ����id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ���ķ������", lng�ⷿID)
    If rsTemp.EOF Then
        bln������� = bln�ⷿ����
    Else
        bln������� = bln���÷���
    End If
    
    '�������γ�������ĵ�������
    If mint��ȷ���� = 0 Then
        Call ��ȡ���ĳɱ��ۺ��ۼ�(rsData, rsData!����ID, 0, IIf(bln�������, 1, 0), bln�Ƿ���, dbl����ϵ��)
        Exit Sub
    End If
    
    '�������������ֽ�;���ⲻ����,������ֽ�
    gstrSQL = " Select Nvl(��������,0)/" & dbl����ϵ�� & " ��������,Nvl(����,0) ����, " & _
        " �ϴβ��� as ����,�ϴ����� as ����,Ч��,���Ч��,�ϴ��������� as �������� " & _
        " From ҩƷ��� Where ����=1 and �ⷿid = [1] And ҩƷid = [2] And nvl(��������,0)<>0 " & _
        " Order by Nvl(����,0) "
        
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ÿ��", lng�ⷿID, lng����ID)
        
    If bln������� Then
        dblʣ������ = dbl��д����
        If dblʣ������ > rsTemp!�������� Then
            rsData.Delete
            Do While Not rsTemp.EOF
                If rsTemp!�������� > 0 Then
                    If dblʣ������ > rsTemp!�������� Then
                        rsData.AddNew
                            
                        rsData!ʵ������ = rsTemp!��������
                        rsData!���� = rsTemp!����
                        rsData!���� = rsTemp!����
                        rsData!���� = rsTemp!����
                        rsData!Ч�� = rsTemp!Ч��
                        rsData!���Ч�� = rsTemp!���Ч��
                        rsData!����ID = lng����ID
                        rsData!����ϵ�� = dbl����ϵ��
                        rsData!�ⷿid = lng�ⷿID
                        rsData!����ID = lng����ID
                        
                        Call ��ȡ���ĳɱ��ۺ��ۼ�(rsData, rsData!����ID, rsData!����, 1, bln�Ƿ���, dbl����ϵ��)
                        
                        dblʣ������ = dblʣ������ - rsTemp!��������
                    Else
                        rsData.AddNew
                        
                        rsData!ʵ������ = dblʣ������
                        rsData!���� = rsTemp!����
                        rsData!���� = rsTemp!����
                        rsData!���� = rsTemp!����
                        rsData!Ч�� = rsTemp!Ч��
                        rsData!���Ч�� = rsTemp!���Ч��
                        rsData!����ID = lng����ID
                        rsData!����ϵ�� = dbl����ϵ��
                        rsData!�ⷿid = lng�ⷿID
                        rsData!����ID = lng����ID
                        
                        Call ��ȡ���ĳɱ��ۺ��ۼ�(rsData, rsData!����ID, rsData!����, 1, bln�Ƿ���, dbl����ϵ��)
                        
                        Exit Do
                    End If
                End If
                rsTemp.MoveNext
            Loop
        Else
            rsData!ʵ������ = dbl��д����
            rsData!���� = rsTemp!����
            rsData!���� = rsTemp!����
            rsData!���� = rsTemp!����
            rsData!Ч�� = rsTemp!Ч��
            rsData!���Ч�� = rsTemp!���Ч��
            
            Call ��ȡ���ĳɱ��ۺ��ۼ�(rsData, rsData!����ID, rsData!����, 1, bln�Ƿ���, dbl����ϵ��)
        End If
    Else
        '���ݿ�����ж���д�����Ƿ���ڿ���������
        '1)�� �������� < ��д�������� ��д���� = ��������
        '2)�� �������� >= ��д�������� ��д���� = ��д����
'        gstrSQL = " Select sum(Nvl(��������,0)/" & dbl����ϵ�� & ") �������� From ҩƷ��� Where �ⷿid = [1] And ҩƷid = [2]"
'
'        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ÿ��", lng�ⷿID, lng����ID)
    
        If mint����� = 2 Then
            If rsTemp!�������� < dbl��д���� Then
                rsData!ʵ������ = rsTemp!��������
                If rsTemp!�������� = 0 Then
                    rsData.Delete
                    Exit Sub
                End If
            Else
                rsData!ʵ������ = dbl��д����
            End If
        Else
            rsData!ʵ������ = dbl��д����
        End If
        
        rsData!���� = 0
        
        Call ��ȡ���ĳɱ��ۺ��ۼ�(rsData, rsData!����ID, rsData!����, 0, bln�Ƿ���, dbl����ϵ��)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ȡ��ǰ�ⷿ����ͨ����
Private Sub getDept()
    
    Dim rsTemp As New ADODB.Recordset
    
    '��鲢װ������ⷿ
    err = 0: On Error Resume Next
    Set rsTemp = ReturnSQL(mlng�ⷿid, Me.Caption, True, , 1716)
    With rsTemp
        cboStock.Clear
        Do While Not .EOF
            cboStock.AddItem !����
            cboStock.ItemData(cboStock.NewIndex) = !Id
            .MoveNext
        Loop
        If cboStock.ListIndex < 0 Then cboStock.ListIndex = 0
    End With
End Sub

'�������������
Private Function GetDepend() As Boolean
    Dim strMsg As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    GetDepend = False
    With rsTemp
        '��������������Ƿ�����
        strMsg = "û�����������ƿ����⼰�����������������������ã�"
        
        gstrSQL = "" & _
            "   SELECT B.Id,B.ϵ�� " & _
            "   FROM ҩƷ�������� A, ҩƷ������ B " & _
            "   Where A.���id = B.ID  AND A.���� = 34"
            
        zldatabase.OpenRecordset rsTemp, gstrSQL, "�����ƿ����"
        
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "ϵ��=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û�����������ƿ������������������������ã�"
            GoTo ErrHand
        End If
        .Filter = "ϵ��=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û�����������ƿ�ĳ����������������������ã�"
            GoTo ErrHand
        End If
        .Filter = 0
        .Close
    End With
    
    If mlngMode = 1716 Then
        Set rsTemp = ReturnSQL(mlng�ⷿid, "�����ƿ����", True, , 1716)
        strMsg = "û���κο�����ⷿ������[���Ĳ�������]���������������ã�"
    ElseIf mlngMode = 1722 Then
        Set rsTemp = ReturnSQL(mlng�ⷿid, "�����������", True, , 1722)
        strMsg = "û���κοⷿ�������죬����[���Ĳ�������]���������������ã�"
    End If
    rsTemp.Filter = "ID<>" & mlng�ⷿid
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsTemp.Close
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetDetail()
    Dim rsTemp As New Recordset
    Dim IntBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim strUnit As String                       '��λ����:�����ﵥλ��סԺ��λ��
    Dim str��װϵ�� As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        Select Case mintUnit
            Case 0
                str��װϵ�� = "1"
            Case Else
                str��װϵ�� = "D.����ϵ��"
        End Select
        
        
        gstrSQL = "" & _
            "   SELECT b.���,'['||M.����||']'||M.���� as ����, M.���," & IIf(mintUnit = 0, "M.���㵥λ", "D.��װ��λ") & " as  ��λ," & _
            "           trim(to_char(b.ǰ������ /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ǰ������," & _
            "           trim(to_char(b.�������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ��������," & _
            "           trim(to_char(b.������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) �������," & _
            "           trim(to_char(b.�ƻ����� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) �ƻ�����," & _
            "           trim(to_char(b.���� *" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ")) ����," & _
            "           trim(to_char(b.���," & mOraFMT.FM_��� & ")) ���, " & _
            " Trim(To_Char(Decode(M.�Ƿ���, 0, P.�ּ� * " & str��װϵ�� & ", B.���� * " & str��װϵ�� & " * (1+(1 / (1 - D.ָ������� / 100) - 1))), " & mOraFMT.FM_���ۼ� & ")) �ۼ�, " & _
            " Trim(To_Char(Decode(M.�Ƿ���, 0, P.�ּ� , B.���� * (1+(1 / (1 - D.ָ������� / 100) - 1))) * B.�ƻ�����," & mOraFMT.FM_��� & ")) �ۼ۽��, " & _
            " b.�ϴι�Ӧ�� as ��Ӧ��,b.�ϴ������� as ������,b.����ID " & _
            "   FROM ���ϲɹ��ƻ� a, ���ϼƻ����� b,���ű� c,�������� D,�շ���ĿĿ¼ M, �շѼ�Ŀ P " & _
            "   Where a.id = b.�ƻ�id " & _
            "           and nvl(a.�ⷿid,0)=c.id(+) " & _
            "           and b.����id=d.����id and b.����id=M.id  And M.ID = P.�շ�ϸĿid " & _
            "   And (P.��ֹ���� Is Null Or Sysdate Between P.ִ������ And Nvl(P.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
            "           AND b.�ƻ�ID =[1] " & _
            "   Order by ���"
        
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�ƻ�����", Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ID"))))
        
        With vsfDetail
            .Redraw = flexRDNone
            Set .DataSource = rsTemp.DataSource
            .Redraw = flexRDDirect
            If mint��ѯ��ʽ = 1 Then
                .TextMatrix(0, 7) = "�깺����"
                .ColHidden(.ColIndex("ǰ������")) = True
                .ColHidden(.ColIndex("��������")) = True
                .ColHidden(.ColIndex("�������")) = True
            End If
        End With
        
        With vsfDetail
        
            '���ز���id��
            .ColHidden(.ColIndex("����ID")) = True
            For i = 1 To .Rows - 1
                '�ж��Ƿ�ͣ�ã�ͣ����ʾδ
                If �Ƿ�ͣ��(Val(.TextMatrix(i, .ColIndex("����ID")))) Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF00FF
                End If
            Next
            
        End With
        
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetImportData(ByVal strNo As String) As ADODB.Recordset
    On Error GoTo ErrHandle
    If mint��ѯ��ʽ = 0 Then
        gstrSQL = "Select ����id, ����, Sum(ʵ������) As ʵ������, �ɱ���, Sum(�ɱ����) As �ɱ����, �ۼ�, Sum(�ۼ۽��) As �ۼ۽��, ��Ӧ��id, ���� " & _
            " From (Select B.����id, M.����, B.�ƻ����� As ʵ������, B.���� As �ɱ���, B.��� As �ɱ����," & _
            " Decode(M.�Ƿ���, 0, P.�ּ�, B.���� * (1+(1 / (1 - D.ָ������� / 100) - 1))) As �ۼ�, " & _
            " Decode(M.�Ƿ���, 0, P.�ּ�, B.���� * (1+(1 / (1 - D.ָ������� / 100) - 1))) * B.�ƻ����� As �ۼ۽��, G.ID As ��Ӧ��id, B.�ϴ������� As ���� " & _
            " From ���ϲɹ��ƻ� A, ���ϼƻ����� B, ���ű� C, �������� D, �շ���ĿĿ¼ M, �շѼ�Ŀ P, ��Ӧ�� G, " & _
            " Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) L " & _
            " Where A.ID = B.�ƻ�id And Nvl(A.�ⷿid, 0) = C.ID(+) And B.����id = D.����id And B.����id = M.ID And M.ID = P.�շ�ϸĿid And " & _
            " (P.��ֹ���� Is Null Or Sysdate Between P.ִ������ And Nvl(P.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) And a.���� = 0 And " & _
            " (G.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or G.����ʱ�� Is Null) And Substr(G.����, 5, 1) = 1 And " & _
            " Nvl(G.ĩ��, 0) = 1 And B.�ϴι�Ӧ�� = G.���� And A.NO = L.Column_Value) " & _
            " Group By ����id, ����, �ɱ���, �ۼ�, ��Ӧ��id, ���� " & _
            " Order By ��Ӧ��id, ���� "
    ElseIf mint��ѯ��ʽ = 1 And (mlngMode = 1716 Or mlngMode = 1722) Then
        gstrSQL = "Select ����id, ����, Sum(ʵ������) As ʵ������, ����, �ⷿid, ����id, ����ϵ��" & vbNewLine & _
                "From (Select b.����id, m.����, b.�ƻ����� / d.����ϵ�� As ʵ������, d.����ϵ��, b.�ϴ������� As ����, �ⷿid, ����id" & vbNewLine & _
                "       From ���ϲɹ��ƻ� A, ���ϼƻ����� B, ���ű� C, �������� D, �շ���ĿĿ¼ M, �շѼ�Ŀ P," & vbNewLine & _
                "            Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) L" & vbNewLine & _
                "       Where a.Id = b.�ƻ�id And Nvl(a.�ⷿid, 0) = c.Id(+) And b.����id = d.����id And b.����id = m.Id And m.Id = p.�շ�ϸĿid And" & vbNewLine & _
                "             (p.��ֹ���� Is Null Or Sysdate Between p.ִ������ And Nvl(p.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) And a.���� = 1 And" & vbNewLine & _
                "             a.No = l.Column_Value)" & vbNewLine & _
                "Group By ����id, ����, ����, �ⷿid, ����id, ����ϵ��" & vbNewLine & _
                "Order By ����"
    End If
            
    Set GetImportData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�ƻ���ϸ", strNo)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList()
    Dim rsTemp As New Recordset
    Dim lng����ID As Long
    
    On Error GoTo ErrHandle
    If mint��ѯ��ʽ = 1 And (mlngMode = 1716 Or mlngMode = 1722) Then
        lng����ID = cboStock.ItemData(cboStock.ListIndex)
    Else
        lng����ID = 0
    End If
    
    If mint��ѯ��ʽ = 0 Then
        gstrSQL = "" & _
            "   SELECT id,'' As ѡ��,�ڼ�,no, decode(�ƻ�����,1,'�¶ȼƻ�',2,'���ȼƻ�',3,'��ȼƻ�','�ܶȼƻ�') as �ƻ����� ," & _
            "           decode(���Ʒ���,1,'����ͬ�����β��շ�',2,'�ٽ��ڼ�ƽ�����շ�',3,'���ϴ���������շ�',4, '���������������շ�', '�����깺���շ�') as ���Ʒ��� ," & _
            "           ������,to_char(��������,'yyyy-mm-dd HH24:MI:SS') as ��������, �����," & _
            "           to_char(�������,'yyyy-mm-dd HH24:MI:SS') as �������,����˵�� " & _
            "   From ���ϲɹ��ƻ� a " & _
            "  Where ����=0 And ������� Is Not Null "
    Else
        gstrSQL = "" & _
                "   SELECT id,'' As ѡ��,�ڼ�,no, decode(�ƻ�����,1,'�¶ȼƻ�',2,'���ȼƻ�',3,'��ȼƻ�','�ܶȼƻ�') as �ƻ����� ," & _
                "           decode(���Ʒ���,1,'����ͬ�����β��շ�',2,'�ٽ��ڼ�ƽ�����շ�','���ϴ���������շ�') as ���Ʒ��� ," & _
                "           ������,to_char(��������,'yyyy-mm-dd HH24:MI:SS') as ��������, �����," & _
                "           to_char(�������,'yyyy-mm-dd HH24:MI:SS') as �������,����˵�� " & _
                "   From ���ϲɹ��ƻ� a " & _
                "  Where ����=1 And ������� Is Not Null "
    End If
    
    
    If mint��ѯ��ʽ = 0 Then
        If mbln���пⷿ = True Then
            gstrSQL = gstrSQL & " And (nvl(�ⷿid,0) =[1] Or �ⷿid Is Null) "
        Else
            gstrSQL = gstrSQL & " And nvl(�ⷿid,0) =[1]"
        End If
    ElseIf mint��ѯ��ʽ = 1 Then
        If mlngMode = 1716 Then
            gstrSQL = gstrSQL & " And nvl(�ⷿid,0) =[1] and  ����id = [5] "
        ElseIf mlngMode = 1722 Then
            gstrSQL = gstrSQL & " And nvl(����id,0) =[1] and  �ⷿid = [5] "
        End If
    End If
    
    If chkNoTime.Value = 0 Then
        gstrSQL = gstrSQL & " and ������� Between [2] And [3] "
    End If
    
    If Trim(txtNo.Text) <> "" Then
        gstrSQL = gstrSQL & " And No=[4] "
    End If
         
    gstrSQL = gstrSQL & " ORDER BY �ڼ�,no "

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�ɹ��ƻ�", _
        mlng�ⷿid, _
        CDate(Format(dtp��ʼʱ��.Value, "yyyy-mm-dd") & " 00:00:00"), _
        CDate(Format(dtp����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59"), _
        txtNo.Text, _
        lng����ID)
    
    With vsfList
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        .Redraw = flexRDDirect
        If mint��ѯ��ʽ = 1 Then
            .ColHidden(.ColIndex("���Ʒ���")) = True
        End If
        If rsTemp.EOF = False Then .Row = 1
        vsfDetail.Rows = 1
    End With
    
    staThis.Panels(2).Text = "��ǰ����" & rsTemp.RecordCount & "�ŵ��ݣ�û��ѡ�񵥾�"
    
    Call vsfList_EnterCell
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveCard() As Boolean
    Dim intRow As Integer
    Dim strNo�� As String
    Dim rsData As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngCur��Ӧ��ID As Long
    Dim int��� As Integer
    Dim strNo As String
    Dim strDate As String
    Dim blnBeginTrans As Boolean
    
    On Error GoTo ErrHandle
    
    With vsfList
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("ѡ��")) = "��" Then
                strNo�� = IIf(strNo�� = "", "", strNo�� & ",") & .TextMatrix(intRow, .ColIndex("NO"))
            End If
        Next
    End With
    
    If mint��ѯ��ʽ = 0 And strNo�� = "" Then
        MsgBox "��ѡ��Ҫ����Ĳɹ��ƻ����ݣ�", vbOKOnly, gstrSysName
        Exit Function
    ElseIf mint��ѯ��ʽ = 1 And strNo�� = "" Then
        MsgBox "��ѡ��Ҫ������깺���ݣ�", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set rsData = GetImportData(strNo��)
    
    If rsData Is Nothing Then Exit Function
    If rsData.EOF Then
        If mint��ѯ��ʽ = 1 Then
            MsgBox "�޷��������ݣ�������ѡ�е������Ƿ��п�档"
        End If
        Exit Function
    End If
    
'    If mint��ѯ��ʽ = 1 Then
'        If MsgBox("�����Զ����Ƴ��ⷿ���зֽ⣬�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'            Exit Function
'        End If
'    End If
    
    If mint��ѯ��ʽ = 1 Then
        '�������ؼ�¼��
        With rsTmp
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable
            .Fields.Append "�ɱ���", adDouble, , adFldIsNullable
            .Fields.Append "�ɱ����", adDouble, , adFldIsNullable
            .Fields.Append "�ۼ�", adDouble, , adFldIsNullable
            .Fields.Append "�ۼ۽��", adDouble, , adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "Ч��", adDate, , adFldIsNullable
            .Fields.Append "���Ч��", adDate, , adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ϵ��", adDouble, 18, adFldIsNullable
            .Fields.Append "�ⷿID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Open
            
            rsData.MoveFirst
            Do While Not rsData.EOF
                If �Ƿ���(rsData!����ID) Then
                    .AddNew
                    !ʵ������ = IIf(IsNull(rsData!ʵ������), 0, rsData!ʵ������)
'                    !�ɱ��� = IIf(IsNull(rsData!�ɱ���), 0, rsData!�ɱ���)
'                    !�ɱ���� = IIf(IsNull(rsData!�ɱ����), 0, rsData!�ɱ����)
'                    !�ۼ� = IIf(IsNull(rsData!�ۼ�), 0, rsData!�ۼ�)
'                    !�ۼ۽�� = IIf(IsNull(rsData!�ۼ۽��), 0, rsData!�ۼ۽��)
                    !���� = IIf(IsNull(rsData!����), "", rsData!����)
                    !����ID = IIf(IsNull(rsData!����ID), 0, rsData!����ID)
                    !����ϵ�� = IIf(IsNull(rsData!����ϵ��), 1, rsData!����ϵ��)
                    !�ⷿid = IIf(IsNull(rsData!�ⷿid), 0, rsData!�ⷿid)
                    !����ID = IIf(IsNull(rsData!����ID), 0, rsData!����ID)
                    .Update
                End If
                rsData.MoveNext
            Loop
            
            rsTmp.Sort = "����ID"
        End With
        
        If mlngSum > 0 Then
            If mlngMode = 1716 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "��������ͣ�ã��ⲿ�����Ľ��������ƿⵥ�У�", "��" & mlngSum & "��������ͣ�ã��ⲿ�����Ľ��������ƿⵥ�У�"), vbInformation, gstrSysName
            ElseIf mlngMode = 1722 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "��������ͣ�ã��ⲿ�����Ľ����������쵥�У�", "��" & mlngSum & "��������ͣ�ã��ⲿ�����Ľ����������쵥�У�"), vbInformation, gstrSysName
            End If
            mlngSum = 0
            mstrMsg = ""
        End If
        
        '�����
        If mlngMode = 1716 Then
            mint����� = Get������(mlng�ⷿid)
        ElseIf mlngMode = 1722 Then
            mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
        End If
        
        '[�����γ���]��[�����Ϊ"�����ֹ"]����û�п������Ĳ��ܳ��⡣
        If mint����� = 2 Or mint��ȷ���� = 1 Then
            If Not rsTmp.EOF Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                '������������Ƿ��п��
                If mlngMode = 1716 Then
                    If �����(mlng�ⷿid, rsTmp!����ID) = False Then
                        rsTmp.Delete
                    End If
                ElseIf mlngMode = 1722 Then
                    If �����(rsTmp!�ⷿid, rsTmp!����ID) = False Then
                        rsTmp.Delete
                    End If
                End If
                
                rsTmp.MoveNext
            Loop
            
            If mint������ʾ = 1 Then
                If MsgBox("�������Ŀ�治�㣬�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                   Exit Function
               End If
            End If
            
            rsTmp.UpdateBatch
            
            If rsTmp.EOF And rsTmp.RecordCount = 0 Then
                MsgBox "�޷��������ݣ�������ѡ�е������Ƿ��п��ÿ�档"
                Exit Function
            End If
            
            rsTmp.MoveFirst
        End If
        
        '�����γ��⡣��Ӧ�÷ֽ⵽��Ӧ�������ϡ�
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If mlngMode = 1716 Then
                '�����Ľ��зֽ�
                Call ���ķֽ�(rsTmp, mlng�ⷿid, rsTmp!����ID, rsTmp!ʵ������, rsTmp!����ϵ��)
            End If
            If mlngMode = 1722 Then
                '�����Ľ��зֽ�
                Call ���ķֽ�(rsTmp, cboStock.ItemData(cboStock.ListIndex), rsTmp!����ID, rsTmp!ʵ������, rsTmp!����ϵ��)
            End If
            rsTmp.MoveNext
        Loop
        
        rsTmp.UpdateBatch
        
        If rsTmp.EOF And rsTmp.RecordCount = 0 Then
            MsgBox "�޷��������ݣ�������ѡ�е������Ƿ��п��ÿ�档"
            Exit Function
        End If
            
        rsTmp.MoveFirst
    End If
    
    strDate = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    If mint��ѯ��ʽ = 0 Then
        With rsData
            Do While Not .EOF
                If �Ƿ���(Val(!����ID)) Then
                    int��� = int��� + 1
                    If lngCur��Ӧ��ID <> !��Ӧ��ID Then
                        lngCur��Ӧ��ID = !��Ӧ��ID
                        int��� = 0
                        strNo = zldatabase.GetNextNo(68, mlng�ⷿid)
                    End If
                    
                    gstrSQL = "zl_�����⹺_INSERT("
                    '  No_In         In ҩƷ�շ���¼.NO%Type,
                    gstrSQL = gstrSQL & "'" & strNo & "',"
                    '  ���_In       In ҩƷ�շ���¼.���%Type,
                    gstrSQL = gstrSQL & "" & int��� & ","
                    '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                    gstrSQL = gstrSQL & "" & mlng�ⷿid & ","
                    '  ��ҩ��λid_In In ҩƷ�շ���¼.��ҩ��λid%Type,
                    gstrSQL = gstrSQL & "" & !��Ӧ��ID & ","
                    '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                    gstrSQL = gstrSQL & "" & !����ID & ","
                    '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                    gstrSQL = gstrSQL & "'" & !���� & "',"
                    '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
                    gstrSQL = gstrSQL & "" & !ʵ������ & ","
                    '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!�ɱ���, g_С��λ��.obj_���С��.�ɱ���С��) & ","
                    '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!�ɱ����, g_С��λ��.obj_ɢװС��.���С��) & ","
                    '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                    gstrSQL = gstrSQL & "100,"
                    '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!�ۼ�, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                    '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!�ۼ۽��, g_С��λ��.obj_ɢװС��.���С��) & ","
                    '  ���_In       In ҩƷ�շ���¼.���%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!�ۼ۽��, g_С��λ��.obj_ɢװС��.���С��) - Round(!�ɱ����, g_С��λ��.obj_ɢװС��.���С��) & ","
                    '  ���۲��_In   In ҩƷ�շ���¼.���%Type := Null,Ŀǰ������÷��ֶ�
                    gstrSQL = gstrSQL & "Null,"
                    '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '   ע��֤��_In   In ҩƷ�շ���¼.ע��֤��%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ������_In     In ҩƷ�շ���¼.������%Type := Null,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '  �������_In   In Ӧ����¼.�������%Type := Null
                    gstrSQL = gstrSQL & "Null,"
                    '  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                    gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS'),"
                    '  �˲���_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  �˲�����_In   In ҩƷ�շ���¼.��ҩ����%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  ����_In       In ҩƷ�շ���¼.����%Type := 0,
                    gstrSQL = gstrSQL & "0,"
                    '  �˻�_In       In Number := 1
                    gstrSQL = gstrSQL & "1)"
                        
                    If blnBeginTrans = False Then gcnOracle.BeginTrans
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    blnBeginTrans = True
                
                End If
                        
                .MoveNext
            Loop
        End With
    Else
        With rsTmp
            Do While Not .EOF
'                If �Ƿ���(Val(!����ID)) Then
                    int��� = int��� + 1
                    If mlngMode = 1716 Then    '�����ƿ�
                        If int��� = 1 Then
                            strNo = sys.GetNextNo(72, mlng�ⷿid)
                        Else
                            '��Ϊ�ƿ���2���ⷿ�����������"2"����
                            int��� = int��� + 1
                        End If
                            
                        gstrSQL = "Zl_�����ƿ�_Insert("
                        '  No_In         In ҩƷ�շ���¼.No%Type,
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '  ���_In       In ҩƷ�շ���¼.���%Type,
                        gstrSQL = gstrSQL & "" & int��� & ","
                        '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                        gstrSQL = gstrSQL & "" & mlng�ⷿid & ","
                        '  �Է�����id_In In ҩƷ�շ���¼.�Է�����id%Type,
                        gstrSQL = gstrSQL & "" & !����ID & ","
                        '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                        gstrSQL = gstrSQL & "" & !����ID & ","
                        '  ����_In       In ҩƷ�շ���¼.����%Type,
                        gstrSQL = gstrSQL & IIf(mint��ȷ���� = 1, "" & !���� & ",", "0,")
                        '  ��д����_In   In ҩƷ�շ���¼.��д����%Type,
                        gstrSQL = gstrSQL & "" & !ʵ������ * !����ϵ�� & ","
                        '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type,
                        gstrSQL = gstrSQL & "" & !ʵ������ * !����ϵ�� & ","
                        '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ɱ��� / !����ϵ��, g_С��λ��.obj_���С��.�ɱ���С��) & ","
                        '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ɱ����, g_С��λ��.obj_ɢװС��.���С��) & ","
                        '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ۼ� / !����ϵ��, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                        '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ۼ۽��, g_С��λ��.obj_ɢװС��.���С��) & ","
                        '  ���_In       In ҩƷ�շ���¼.���%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ۼ۽��, g_С��λ��.obj_ɢװС��.���С��) - Round(!�ɱ����, g_С��λ��.obj_ɢװС��.���С��) & ","
                        '  ������_In     In ҩƷ�շ���¼.������%Type,
                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                        '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "'" & !���� & "',"
                        '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "'" & !���� & "',"
                        '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!Ч��) = "", "Null", "to_date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!���Ч��) = "", "Null", "to_date('" & Format(!���Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                        gstrSQL = gstrSQL & "Null,"
                        '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                        gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS'))"
                            
                    ElseIf mint��ѯ��ʽ = 1 And mlngMode = 1722 Then    '��������
                        If int��� = 1 Then
                            strNo = sys.GetNextNo(72, mlng�ⷿid)
                        Else
                            '��Ϊ�ƿ���2���ⷿ�����������"2"����
                            int��� = int��� + 1
                        End If
                            
                        gstrSQL = "Zl_��������_Insert("
                        '  No_In         In ҩƷ�շ���¼.No%Type,
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '  ���_In       In ҩƷ�շ���¼.���%Type,
                        gstrSQL = gstrSQL & "" & int��� & ","
                        '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                        gstrSQL = gstrSQL & "" & !�ⷿid & ","
                        '  �Է�����id_In In ҩƷ�շ���¼.�Է�����id%Type,
                        gstrSQL = gstrSQL & "" & mlng�ⷿid & ","
                        '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                        gstrSQL = gstrSQL & "" & !����ID & ","
                        '  ����_In       In ҩƷ�շ���¼.����%Type,
                        gstrSQL = gstrSQL & IIf(mint��ȷ���� = 1, "" & !���� & ",", "0,")
                        '  ��д����_In   In ҩƷ�շ���¼.��д����%Type,
                        gstrSQL = gstrSQL & "" & !ʵ������ * !����ϵ�� & ","
                        '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type,
                        gstrSQL = gstrSQL & "" & !ʵ������ * !����ϵ�� & ","
                        '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ɱ��� / !����ϵ��, g_С��λ��.obj_���С��.�ɱ���С��) & ","
                        '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ɱ����, g_С��λ��.obj_ɢװС��.���С��) & ","
                        '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ۼ� / !����ϵ��, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                        '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ۼ۽��, g_С��λ��.obj_ɢװС��.���С��) & ","
                        '  ���_In       In ҩƷ�շ���¼.���%Type,
                        gstrSQL = gstrSQL & "" & Round(!�ۼ۽��, g_С��λ��.obj_ɢװС��.���С��) - Round(!�ɱ����, g_С��λ��.obj_ɢװС��.���С��) & ","
                        '  ������_In     In ҩƷ�շ���¼.������%Type,
                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                        '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "'" & !���� & "',"
                        '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "'" & !���� & "',"
                        '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!Ч��) = "", "Null", "to_date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!���Ч��) = "", "Null", "to_date('" & Format(!���Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                        gstrSQL = gstrSQL & "Null,"
                        '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                        gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS'))"
                    
                    End If
                    
                    If blnBeginTrans = False Then gcnOracle.BeginTrans
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    blnBeginTrans = True
'                End If
            
                .MoveNext
            Loop
        End With
    End If
        
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '��ʾ��Ϣ
    If mlngSum > 0 Then
        If mint��ѯ��ʽ = 0 Then
            MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "��������ͣ�ã��ⲿ�����Ľ��������⹺��ⵥ�У�", "��" & mlngSum & "��������ͣ�ã��ⲿ�����Ľ��������⹺��ⵥ�У�"), vbInformation, gstrSysName
        Else
            If mlngMode = 1716 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "��������ͣ�ã��ⲿ�����Ľ��������ƿⵥ�У�", "��" & mlngSum & "��������ͣ�ã��ⲿ�����Ľ��������ƿⵥ�У�"), vbInformation, gstrSysName
            ElseIf mlngMode = 1722 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "��������ͣ�ã��ⲿ�����Ľ����������쵥�У�", "��" & mlngSum & "��������ͣ�ã��ⲿ�����Ľ����������쵥�У�"), vbInformation, gstrSysName
            End If
        End If
        mlngSum = 0
        mstrMsg = ""
    End If
    
    SaveCard = True
    Exit Function
ErrHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'���ܣ��ж������Ƿ�ͣ�ã��ٸ��ݸ�ѡ��������ͣ�����ġ�����ֵ
'����ѡʱ��������ͣ�����ģ��������ж������Ƿ�ͣ��ֱ�ӷ���TRUE
'������ѡʱ����������ͣ�����ģ����ж������Ƿ�ͣ�ã�ͣ�÷���false
Private Function �Ƿ���(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If lng����ID = 0 Then Exit Function
    
    If chk������ͣ������.Value = 1 Then '������ͣ������
        �Ƿ��� = True
        Exit Function
    Else '��������ͣ������
    
        '�ж������Ƿ�ͣ��
        gstrSQL = "select ����,��� from �շ���ĿĿ¼ where ID = [1] and nvl(����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD')"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��������Ƿ�ͣ��", lng����ID)
        
        If rsTemp.RecordCount = 0 Then 'rsTemp.RecordCount = 0˵��������δͣ��
            �Ƿ��� = True
        Else
            �Ƿ��� = False
            
            mlngSum = mlngSum + 1
            If mlngSum <= 3 Then 'ƴ��ʾ��Ϣ��
                mstrMsg = mstrMsg & "��" & rsTemp!���� & "(" & rsTemp!��� & ")��" & Chr(10)
            End If
            
        End If
    End If

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(frmMain As Form, ByVal str�ⷿ As String, ByVal lng�ⷿID As Long, ByVal intUnit As Integer, _
                    ByVal bln���пⷿ As Boolean, Optional blnSuccess As Boolean = False, _
                    Optional int��ѯ��ʽ As Integer, Optional lngMode As Integer, Optional int��ȷ���� As Integer)
    
    Set mfrmMain = frmMain
    
    mStr�ⷿ = str�ⷿ
    mlng�ⷿid = lng�ⷿID
    mintUnit = intUnit
    mbln���пⷿ = bln���пⷿ
    mint��ѯ��ʽ = int��ѯ��ʽ
    mlngMode = lngMode
    mint��ȷ���� = int��ȷ����
    
    If int��ѯ��ʽ = 1 Then
        If Not GetDepend Then Exit Sub
    End If
    
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
End Sub





Private Sub chkNoTime_Click()
    If chkNoTime.Value = 0 Then
        dtp��ʼʱ��.Enabled = True
        dtp����ʱ��.Enabled = True
    Else
        dtp��ʼʱ��.Enabled = False
        dtp����ʱ��.Enabled = False
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    mblnSuccess = SaveCard
    If mblnSuccess = True Then
        Unload Me
    End If
End Sub

Private Sub cmd��ȡ_Click()
    GetList
End Sub


Private Sub Form_Activate()
    Me.Caption = Me.Caption & "(" & mStr�ⷿ & ")"
End Sub

Private Sub Form_Load()

    chk������ͣ������.Value = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9Stuff", "������ͣ������", 0)
    
    staThis.Panels(2).Picture = picColor
    
    dtp����ʱ��.Value = zldatabase.Currentdate
    dtp��ʼʱ��.Value = DateAdd("m", -1, Me.dtp����ʱ��.Value)
    
    If mint��ѯ��ʽ = 1 Then
        chk������ͣ������.Value = 0
        chk������ͣ������.Visible = False
        lblʱ��.Caption = "�깺���������"
        vsfList.ColHidden(5) = True     '����[���Ʒ���]
        vsfDetail.ColHidden(4) = True   '����[ǰ������]
        vsfDetail.ColHidden(5) = True   '����[��������]
        vsfDetail.ColHidden(6) = True   '����[�������]
        Me.Caption = "�����깺��"
        If mlngMode = 1716 Then
            CmdSave.Caption = "�����ƿⵥ(&O)"
        ElseIf mlngMode = 1722 Then
            CmdSave.Caption = "�������쵥(&O)"
        End If
        vsfDetail.TextMatrix(0, 7) = "�깺����"
        
        If mlngMode = 1716 Or mlngMode = 1722 Then
            If mlngMode = 1722 Then
                lbl����ⷿ.Caption = "���Ͽⷿ"
            End If
            lbl����ⷿ.Visible = True
            cboStock.Visible = True
            Call getDept
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    Dim dblStateHeight As Double
    
    On Error Resume Next
    
    If Me.Height < 8325 Then Me.Height = 8325
    If Me.Width < 12420 Then Me.Width = 12420
    
    dblStateHeight = IIf(staThis.Visible, staThis.Height, 0)
    
    With CmdCancel
        .Top = Me.ScaleHeight - dblStateHeight - .Height - 200
        .Left = Me.ScaleWidth - .Width - 200
    End With
    
    With CmdSave
        .Top = CmdCancel.Top
        .Left = CmdCancel.Left - .Width - 200
    End With
    
    With chk������ͣ������
        .Top = CmdSave.Top + (CmdSave.Height - .Height) / 2
        .Left = CmdSave.Left - .Width - 200
    End With
    
    With lblMsg
        .Top = chk������ͣ������.Top
    End With
    
    With frmCondition
        .Width = Me.ScaleWidth - 200
    End With
    
    With vsfList
        .Width = frmCondition.Width
    End With
    
    With picLine
        .Top = vsfList.Top + vsfList.Height
        .Width = frmCondition.Width
    End With
    
    With vsfDetail
        .Top = picLine.Top + picLine.Height
        .Width = frmCondition.Width
        .Height = CmdCancel.Top - .Top - 200
    End With
        
    With cmd��ȡ
        .Left = frmCondition.Width - .Width - 200
    End With
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '����ע�����Ϣ(�Ƿ���ʾͣ������)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\zl9Stuff", "������ͣ������", chk������ͣ������.Value
End Sub

Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfList.Height + y <= 500 Or vsfDetail.Height - y <= 500 Then Exit Sub
        
        picLine.Top = picLine.Top + y
        vsfList.Height = vsfList.Height + y
        vsfDetail.Height = vsfDetail.Height - y
        vsfDetail.Top = vsfDetail.Top + y
        
        Me.Refresh
    End If
End Sub


Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If mint��ѯ��ʽ = 1 Then
        If KeyCode = vbKeyReturn Then
            If Len(txtNo) < 8 And Len(txtNo) > 0 Then
                txtNo.Text = GetFullNO(txtNo.Text, 72, mlng�ⷿid)
            End If
            zlCommFun.PressKey (vbKeyTab)
        End If
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        If Len(txtNo) < 8 And Len(txtNo) > 0 Then
            txtNo.Text = GetFullNO(txtNo.Text, 77, mlng�ⷿid)
            GetList
        End If
    End If
End Sub


Private Sub vsfList_DblClick()
    Dim intRow As Integer
    Dim intSelectCount As Integer
    
    With vsfList
        If .Row = 0 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        .Redraw = flexRDNone
        
        If .TextMatrix(.Row, .ColIndex("ѡ��")) = "��" Then
            .TextMatrix(.Row, .ColIndex("ѡ��")) = ""
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000008
        Else
            .TextMatrix(.Row, .ColIndex("ѡ��")) = "��"
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlue
        End If
        
        .Redraw = flexRDDirect
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("ѡ��")) = "��" Then
                intSelectCount = intSelectCount + 1
            End If
        Next
        
        If intSelectCount = 0 Then
            staThis.Panels(2).Text = "��ǰ����" & .Rows - 1 & "�ŵ��ݣ�û��ѡ�񵥾�"
        Else
            staThis.Panels(2).Text = "��ǰ����" & .Rows - 1 & "�ŵ��ݣ�ѡ����" & intSelectCount & "�ŵ���"
        End If
    End With
End Sub
Private Sub vsfList_EnterCell()
    GetDetail
End Sub


'���ܣ��ж��Ƿ�ͣ��,true - ͣ��
Private Function �Ƿ�ͣ��(ByVal lngҩƷID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If lngҩƷID = 0 Then Exit Function

    
    '�ж�ҩƷ�Ƿ�ͣ��
    gstrSQL = "select ����,��� from �շ���ĿĿ¼ where ID = [1] and nvl(����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD') "
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "���ҩƷ�Ƿ�ͣ��", lngҩƷID)
    
    �Ƿ�ͣ�� = rsTemp.RecordCount <> 0  '˵����ҩƷδͣ��

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

