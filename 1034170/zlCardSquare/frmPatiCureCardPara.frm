VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPatiCureCardPara 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7545
      TabIndex        =   4
      Top             =   5085
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   5085
      Width           =   1100
   End
   Begin TabDlg.SSTab stbPage 
      Height          =   4665
      Left            =   135
      TabIndex        =   0
      Top             =   255
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   8229
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frmPatiCureCardPara.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra�˿���ʽ"
      Tab(0).Control(1)=   "chkLedWelcome"
      Tab(0).Control(2)=   "chk����"
      Tab(0).Control(3)=   "fraShortLine"
      Tab(0).Control(4)=   "txtNameDays"
      Tab(0).Control(5)=   "cmdDeviceSetup"
      Tab(0).Control(6)=   "chkSeekName"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "ҽ�ƿ�Ʊ��(&1)"
      TabPicture(1)   =   "frmPatiCureCardPara.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrint"
      Tab(1).Control(1)=   "fraTitle"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ԥ����Ʊ��(&2)"
      TabPicture(2)   =   "frmPatiCureCardPara.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraPrepay"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraƱ�ݸ�ʽ"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraƱ�ݸ�ʽ 
         Caption         =   "Ԥ��Ʊ�ݸ�ʽ"
         Height          =   1305
         Left            =   150
         TabIndex        =   25
         Top             =   3045
         Width           =   8190
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1005
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   7995
            _cx             =   14102
            _cy             =   1773
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
            FormatString    =   $"frmPatiCureCardPara.frx":0054
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
      Begin VB.Frame fraPrint 
         Height          =   615
         Left            =   -74910
         TabIndex        =   19
         Top             =   3900
         Width           =   8145
         Begin VB.OptionButton optPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   2670
            TabIndex        =   23
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   3810
            TabIndex        =   22
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   1665
            TabIndex        =   21
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����(&S)"
            Height          =   345
            Left            =   6570
            TabIndex        =   20
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label lblPrint 
            Caption         =   "������ӡ��ʽ"
            Height          =   285
            Left            =   435
            TabIndex        =   24
            Top             =   270
            Width           =   1500
         End
      End
      Begin VB.Frame fra�˿���ʽ 
         Caption         =   "�˿���ʽ����"
         Height          =   1050
         Left            =   -74280
         TabIndex        =   14
         Top             =   2235
         Width           =   7320
         Begin VB.OptionButton optBrush 
            Caption         =   "���뵥�ݺ��˿���ˢ���˿�"
            Height          =   180
            Index           =   3
            Left            =   255
            TabIndex        =   18
            Top             =   765
            Width           =   2460
         End
         Begin VB.OptionButton optBrush 
            Caption         =   "���뵥�ݺź��ˢ���˿�"
            Height          =   180
            Index           =   2
            Left            =   3600
            TabIndex        =   17
            Top             =   420
            Width           =   2460
         End
         Begin VB.OptionButton optBrush 
            Caption         =   "����ˢ���˿�"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   16
            Top             =   420
            Width           =   1740
         End
         Begin VB.OptionButton optBrush 
            Caption         =   "������ˢ����֤"
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   15
            Top             =   420
            Value           =   -1  'True
            Width           =   1740
         End
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Left            =   -74265
         TabIndex        =   13
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   1785
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���￨�����Լ��˷�ʽ��ȡ"
         Height          =   180
         Left            =   -74265
         TabIndex        =   12
         Top             =   930
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Frame fraShortLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   -71655
         TabIndex        =   10
         Top             =   1515
         Width           =   285
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
         Left            =   -71655
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "0��ʾ����ʱ������ʱ��"
         Top             =   1335
         Width           =   285
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "���ع���Ԥ��Ʊ��"
         Height          =   2310
         Left            =   150
         TabIndex        =   7
         Top             =   615
         Width           =   8190
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   1920
            Left            =   75
            TabIndex        =   8
            Top             =   270
            Width           =   8025
            _cx             =   14155
            _cy             =   3387
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
            FormatString    =   $"frmPatiCureCardPara.frx":00E2
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
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   -68205
         TabIndex        =   6
         Top             =   3510
         Width           =   1500
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع����շ�Ʊ��"
         Height          =   3345
         Left            =   -74910
         TabIndex        =   1
         Top             =   540
         Width           =   8190
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2925
            Left            =   75
            TabIndex        =   2
            Top             =   270
            Width           =   7995
            _cx             =   14102
            _cy             =   5159
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
            FormatString    =   $"frmPatiCureCardPara.frx":01C2
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
      Begin VB.CheckBox chkSeekName 
         Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
         Height          =   195
         Left            =   -74265
         TabIndex        =   11
         Top             =   1350
         Width           =   4260
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6375
      TabIndex        =   5
      Top             =   5085
      Width           =   1100
   End
End
Attribute VB_Name = "frmPatiCureCardPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnOk As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:�������óɹ�,����true,����ķ���False
    '����:���˺�
    '����:2011-07-14 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOk = False
    
    Me.Show 1, frmMain
    zlSetPara = mblnOk
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
    isValied = False
    On Error GoTo errHandle
    '���ÿ��ʹ����ʽֻ��һ��ѡ��
    With vsBill
        str��� = "-"
        For i = 1 To vsBill.Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����"))) = Trim(.TextMatrix(j, .ColIndex("ҽ�ƿ����"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    ҽ�ƿ����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
  '���ÿ��ʹ��Ԥ��ֻ��һ��ѡ��
    With vsPrepay
        str��� = "-"
        For i = 1 To .Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("Ԥ������")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) = Trim(.TextMatrix(j, .ColIndex("Ԥ������"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    Ԥ������Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:���˺�
    '����:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String, strPrintMode As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '���湲��Ʊ��
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "����ҽ�ƿ�����", strValue, glngSys, mlngModule, blnHavePrivs
    '����Ԥ��Ʊ��
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("Ԥ������")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "����Ԥ��Ʊ������", strValue, glngSys, mlngModule, blnHavePrivs
    
    '����Ԥ����ʽ
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("Ԥ����ӡ��ʽ")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "Ԥ����Ʊ��ʽ", strValue, glngSys, mlngModule, blnHavePrivs
        zlDatabase.SetPara "Ԥ����Ʊ��ӡ��ʽ", strPrintMode, glngSys, mlngModule, blnHavePrivs
    End With
End Sub
Private Sub InitShareInvoice()
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String, rsҽ�ƿ���� As ADODB.Recordset
    Dim strPrintMode As String, strBillFormat As String, blnHavePrivs As Boolean
    Dim strȱʡҽ�ƿ� As String, lngȱʡҽ�ƿ� As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '�ָ��п��
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ID,����,����, nvl(�Ƿ�̶�,0) as �Ƿ�̶�  from ҽ�ƿ����  "
    Set rsҽ�ƿ���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rsҽ�ƿ����.Filter = "����='���￨' and �Ƿ�̶�=1"
    If rsҽ�ƿ����.EOF = False Then
        strȱʡҽ�ƿ� = rsҽ�ƿ����!����: lngȱʡҽ�ƿ� = Val(rsҽ�ƿ����!id)
    End If
    
    zl_vsGrid_Para_Restore mlngModule, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
    strShareInvoice = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModule, , , True, intType)
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
    
    '��ʽ:����ID1,ҽ�ƿ����ID1|����IDn,ҽ�ƿ����IDn|...
    varData = Split(strShareInvoice, "|")

    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(5)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            '99007:���ϴ�,2016/7/29������ҽ�ƿ�Ʊ�ݻ�ȡʹ�����ID
            If Val(Nvl(rsTemp!ʹ�����ID)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = strȱʡҽ�ƿ�
                .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = lngȱʡҽ�ƿ�
            Else
                rsҽ�ƿ����.Filter = "ID=" & Val(Nvl(rsTemp!ʹ�����ID))
                If Not rsҽ�ƿ����.EOF Then
                    .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsҽ�ƿ����!����)
                Else
                    .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsTemp!ʹ�����)
                End If
                .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = Val(Nvl(rsTemp!ʹ�����ID))
            End If
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    '����Ԥ��Ʊ������
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModule, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
    
    strShareInvoice = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, mlngModule, , , True, intType)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,Ԥ�����ID1|����IDn,Ԥ�����IDn|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            If Val(Nvl(rsTemp!ʹ�����, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "�����סԺ����"
            ElseIf Val(Nvl(rsTemp!ʹ�����, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ������Ʊ��"
            Else
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ��סԺƱ��"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = Val(Nvl(rsTemp!ʹ�����))
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ԥ������"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    'Ʊ�ݸ�ʽ����
    Dim strReport As String
    
    zl_vsGrid_Para_Restore mlngModule, vsBillFormat, Me.Name, "Ԥ����Ʊ��ӡ��ʽ", False, False
    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlReadBillFormat(strReport)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    
    '��ȡ����ֵ
    strBillFormat = zlDatabase.GetPara("Ԥ����Ʊ��ʽ", glngSys, mlngModule, , , True, intType)
    strPrintMode = zlDatabase.GetPara("Ԥ����Ʊ��ӡ��ʽ", glngSys, mlngModule, , , True, intType1)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat
        .TextMatrix(1, 0) = "����Ԥ��"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "סԺԤ��"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
        .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Or _
            Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strBillFormat, "|")
    VarType = Split(strPrintMode, "|")
    
    With vsBillFormat
        .Clear 1
        .Rows = 3
        For lngRow = 1 To .Cols - 1
            .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
        Next
        If Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("Ԥ����ӡ��ʽ"), .Rows - 1, .ColIndex("Ԥ����ӡ��ʽ")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Sub cmdOK_Click()
    Dim blnHavePrivs As Boolean, intData As Integer
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    If isValied = False Then Exit Sub
    zlDatabase.SetPara "����ģ������", chkSeekName.value, glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "������������", Val(txtNameDays.Text), glngSys, mlngModule, blnHavePrivs
   
    zlDatabase.SetPara "���Ѽ���", chk����.value, glngSys, mlngModule, blnHavePrivs
    zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chkLedWelcome.value, glngSys, mlngModule, blnHavePrivs
    '����28130��27929
    intData = 0
    If optBrush(3).value = True Then
        intData = 3
    ElseIf optBrush(1).value = True Then
        intData = 1
    ElseIf optBrush(2).value = True Then
        intData = 2
    End If
    Call zlDatabase.SetPara("�˿�ˢ��", intData, glngSys, mlngModule, blnHavePrivs)
    zlDatabase.SetPara "������ӡ��ʽ", IIf(optPrint(0).value, 0, IIf(optPrint(1).value, 1, 2)), glngSys, mlngModule, blnHavePrivs
    Call SaveInvoice
    mblnOk = True: Unload Me
End Sub
Private Sub InitPara()
    Dim blnHavePrivs As Boolean, i As Long
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    txtNameDays.Enabled = True
    txtNameDays.Text = zlDatabase.GetPara("������������", glngSys, mlngModule, , Array(txtNameDays), blnHavePrivs)
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.value = IIf(zlDatabase.GetPara("����ģ������", glngSys, mlngModule, , Array(chkSeekName), blnHavePrivs) = "1", 1, 0)
    chk����.value = IIf(zlDatabase.GetPara("���Ѽ���", glngSys, glngModul, , Array(chk����), blnHavePrivs) = "1", 1, 0)
    'LED�豸
    chkLedWelcome.value = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModule, 1, Array(chkLedWelcome), blnHavePrivs)
    '����28130
    Select Case Val(zlDatabase.GetPara("�˿�ˢ��", glngSys, mlngModule, "0", Array(fra�˿���ʽ, optBrush(0), optBrush(1), optBrush(2), optBrush(3)), InStr(mstrPrivs, "��������") > 0))
    Case 0
        optBrush(0).value = True
    Case 1
        optBrush(1).value = True
    Case "2"
        optBrush(2).value = True
    Case "3"
        optBrush(3).value = True
    End Select
    
    i = Val(zlDatabase.GetPara("������ӡ��ʽ", glngSys, mlngModule, , Array(optPrint(0), optPrint(1), optPrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optPrint(i).value = True
      
End Sub
Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.value = 1 And txtNameDays.Tag = "1"
End Sub

Private Sub cmdPrepayPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdPrintSet_Click()
    '��ӡ����
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me)
End Sub

Private Sub Form_Load()
    Call InitShareInvoice
    Call InitPara
    chkSeekName_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
    zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub
 
Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub
Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Val(.Cell(flexcpData, Row, .ColIndex("ҽ�ƿ����"))) = Val(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
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
Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("Ԥ������"))) = Trim(.Cell(flexcpData, i, .ColIndex("Ԥ������"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
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



