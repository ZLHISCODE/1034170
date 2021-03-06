VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFilesUpgradeLogView 
   Caption         =   "升级日志信息"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8775
   Icon            =   "frmFilesUpgradeLogView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8775
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   7440
      TabIndex        =   0
      Top             =   4920
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfShift 
      Height          =   4728
      Left            =   60
      TabIndex        =   1
      Top             =   12
      Width           =   8556
      _cx             =   15092
      _cy             =   8340
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      GridColor       =   -2147483626
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
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
End
Attribute VB_Name = "frmFilesUpgradeLogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mclsVsfShift As clsVsf
Attribute mclsVsfShift.VB_VarHelpID = -1
Private mstrName    As String
Private mintType    As Integer
Private mrsLog      As ADODB.Recordset

Public Sub ShowMe(ByVal strName As String, ByVal intType As Integer)
    On Error GoTo ErrH
    mstrName = strName
    mintType = intType
    gstrSQL = "Select 处理日期,内容 From zltools.zlClientUpdatelog  WHERE 工作站='" & mstrName & "' And NVL(类型,0)=" & mintType & " order by 处理日期"
    Set mrsLog = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "客户端升级日志")
    If mrsLog.RecordCount > 0 Then
        Me.Show 1, frmMDIMain
    End If
    Exit Sub
ErrH:
    err.Clear
End Sub
'关闭
Private Sub cmdCancel_Click()
    mstrName = ""
    Unload Me
End Sub

Private Sub Form_Load()
    If mstrName <> "" Then
        Me.Caption = mstrName & IIf(mintType, "--升级检查结果", "--升级日志信息")
    End If
    Call InitVSF
    Call LoadVSF
End Sub

Private Sub Form_Resize()
    With vsfShift
        .Top = 15
        .Left = 15
        .Width = ScaleWidth - 30
        .Height = ScaleHeight - 450
    End With
    cmdCancel.Top = vsfShift.Top + vsfShift.Height + 30
    cmdCancel.Left = vsfShift.Left + vsfShift.Width - cmdCancel.Width
    mclsVsfShift.AppendRows = True
End Sub

Private Sub InitVSF()
     
    Set mclsVsfShift = New clsVsf
     
    With mclsVsfShift
        Call .Initialize(Me.Controls, vsfShift, True, False)
        Call .ClearColumn
        Call .AppendColumn("日期", 2400, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("内容", 2200, flexAlignLeftCenter, flexDTString, "", , True)
        
        .AppendRows = True
        .IndicatorMode = 2
        Call .InitializeEdit(False, False, False)
    End With
    
    
    vsfShift.FixedCols = 0
End Sub

Private Sub LoadVSF()
    Dim strTemp() As String
    Dim i As Integer
    
    With vsfShift
'        .Redraw = flexRDNone
       
        If mrsLog.RecordCount > 0 Then
            .Rows = mrsLog.RecordCount + 1
            mrsLog.MoveFirst
            Do Until mrsLog.EOF
                i = i + 1
                .TextMatrix(i, .ColIndex("日期")) = Trim(mrsLog!处理日期)
                .TextMatrix(i, .ColIndex("内容")) = Trim(mrsLog!内容)
                mrsLog.MoveNext
            Loop
        
        End If
        .ShowCell .Rows - 1, .ColIndex("日期")
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsfShift = Nothing
End Sub

