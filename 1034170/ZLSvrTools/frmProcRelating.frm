VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcRelating 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关联过程"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   Icon            =   "frmProcRelating.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1860
      Index           =   0
      Left            =   330
      ScaleHeight     =   1860
      ScaleWidth      =   5955
      TabIndex        =   2
      Top             =   930
      Width           =   5955
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1185
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _cx             =   3413
         _cy             =   2090
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
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "继续(&O)"
      Height          =   350
      Left            =   3915
      TabIndex        =   1
      Top             =   3990
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&O)"
      Height          =   350
      Left            =   5100
      TabIndex        =   0
      Top             =   3990
      Width           =   1100
   End
   Begin VB.Label Label1 
      Caption         =   $"frmProcRelating.frx":6852
      Height          =   600
      Left            =   810
      TabIndex        =   9
      Top             =   240
      Width           =   5250
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   75
      Picture         =   "frmProcRelating.frx":68EE
      Top             =   105
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "当前过程名称："
      Height          =   180
      Left            =   285
      TabIndex        =   8
      Top             =   2925
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "zl_导诊资源目录_Update"
      Height          =   180
      Left            =   1530
      TabIndex        =   7
      Top             =   2910
      Width           =   1980
   End
   Begin VB.Label Label4 
      Caption         =   "您可以直接删除当前存储过程，正在使用它的其他存储过程将会以文本文件的方式存储在本地，请及时对其他存储过程进行处理。"
      Height          =   390
      Left            =   285
      TabIndex        =   6
      Top             =   3195
      Width           =   6015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "调用查看路径："
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   3675
      Width           =   1260
   End
   Begin VB.Label lbl路径 
      Caption         =   "C:\AppSoft\RelProc.ini"
      Height          =   195
      Left            =   1575
      TabIndex        =   4
      Top             =   3675
      Width           =   2490
   End
End
Attribute VB_Name = "frmProcRelating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjMain As Object
Private mclsVsf As clsVsf
Private mlngKey As Long
Private mrsData As ADODB.Recordset
Private mblnStartUp As Boolean

Public Function ShowDialog(ByVal objMain As Object, ByVal lngKey As Long, ByVal rsData As ADODB.Recordset) As Boolean
    Set mobjMain = objMain
    Set mrsData = rsData
    mlngKey = lngKey
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    If ExecuteCommand("刷新数据") = False Then Exit Function
    
    Me.Show 1, mobjMain
    ShowDialog = True
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String
    Dim objItem As Object
    Dim i As Integer
    Dim strTmp As String
    Dim objFSO As TextStream
    
    On Error GoTo ErrHand
    
    Call gclsBase.SQLRecord(rsSQL)
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
    
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True)
            Call .ClearColumn
            Call .AppendColumn("过程名称", 2000, flexAlignLeftCenter, flexDTString, , "Name", True)
            .AppendRows = True
        End With
    Case "刷新数据"
        If mrsData.BOF = False Then
            mrsData.MoveFirst
            With vsf(0)
                For i = 1 To .Rows - 1
                    .TextMatrix(i, .ColIndex("过程名称")) = Nvl(mrsData("过程名称").value)
                Next
            End With
        End If
    Case "删除数据"
        strSQL = "Zl_Zlprocedure_Delete(" & mlngKey & ")"
        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        If SQLRecordExecute(rsSQL) Then
            '生成脚本
            If gobjFile.FileExists(lbl路径.Caption) = False Then
                Set objFSO = gobjFile.CreateTextFile(lbl路径.Caption, True)
            Else
                Set objFSO = gobjFile.OpenTextFile(lbl路径.Caption, ForAppending)
            End If
            For i = 1 To vsf(0).Rows - 1
                Call objFSO.WriteLine(vsf(0).TextMatrix(i, vsf(0).ColIndex("过程名称")))
            Next
        End If
    End Select
    ExecuteCommand = True
    Exit Function
ErrHand:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Sub cmdOK_Click()
    On Error GoTo ErrHand
    
    If ExecuteCommand("删除数据") Then
        Unload Me
    End If
    
    Exit Sub
ErrHand:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsf(0).Move 15, 15, picPane(0).ScaleWidth - 30, picPane(0).ScaleHeight - 30
    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
End Sub

