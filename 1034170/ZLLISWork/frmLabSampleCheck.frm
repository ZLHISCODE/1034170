VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabSampleCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "送检标本核对"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14355
   Icon            =   "frmLabSampleCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14355
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   360
      Left            =   5610
      TabIndex        =   24
      Top             =   7590
      Width           =   1035
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   14280
      TabIndex        =   14
      Top             =   0
      Width           =   14310
      Begin VB.TextBox txtSampleCode 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   915
         TabIndex        =   0
         Top             =   248
         Width           =   3150
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "扫描条码"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "送检科室"
         Height          =   180
         Index           =   1
         Left            =   5220
         TabIndex        =   20
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "送检时间"
         Height          =   180
         Index           =   3
         Left            =   10365
         TabIndex        =   19
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lblInto 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   0
         Left            =   6210
         TabIndex        =   18
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblInto 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   8610
         TabIndex        =   17
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblInto 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   11265
         TabIndex        =   16
         Top             =   330
         Width           =   90
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   8550
         X2              =   9780
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   6135
         X2              =   7365
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   11265
         X2              =   13530
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "送检人"
         Height          =   180
         Index           =   2
         Left            =   7860
         TabIndex        =   15
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdQuet 
      Caption         =   "退出(&Q)"
      Height          =   360
      Left            =   13245
      TabIndex        =   4
      Top             =   7590
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "核对(&D)"
      Height          =   360
      Left            =   11865
      TabIndex        =   3
      Top             =   7590
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   8115
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   6630
      Left            =   -45
      ScaleHeight     =   6600
      ScaleWidth      =   14355
      TabIndex        =   1
      Top             =   840
      Width           =   14385
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6285
         Index           =   2
         Left            =   120
         ScaleHeight     =   6285
         ScaleWidth      =   7050
         TabIndex        =   12
         Top             =   45
         Width           =   7050
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   6165
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   270
            Width           =   7065
            _cx             =   12462
            _cy             =   10874
            Appearance      =   3
            BorderStyle     =   1
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "待登记标本(0)"
            Height          =   180
            Index           =   4
            Left            =   45
            TabIndex        =   22
            Top             =   45
            Width           =   1170
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4230
         Index           =   1
         Left            =   7275
         ScaleHeight     =   4230
         ScaleWidth      =   7005
         TabIndex        =   9
         Top             =   45
         Width           =   7005
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   4140
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Top             =   285
            Width           =   7065
            _cx             =   12462
            _cy             =   7302
            Appearance      =   2
            BorderStyle     =   1
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "已扫描标本(0)"
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1170
         End
      End
      Begin VB.Frame fraNS 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   7185
         MousePointer    =   7  'Size N S
         TabIndex        =   6
         Top             =   4335
         Width           =   7095
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   7260
         ScaleHeight     =   1815
         ScaleWidth      =   7065
         TabIndex        =   5
         Top             =   4545
         Width           =   7065
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   1545
            Index           =   2
            Left            =   30
            TabIndex        =   7
            Top             =   225
            Width           =   6960
            _cx             =   12277
            _cy             =   2725
            Appearance      =   2
            BorderStyle     =   1
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "本批已登记或已核收标本(0)"
            Height          =   180
            Index           =   6
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   2250
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "双击""待登记""表格中的数据行可以将数据添加到""已扫描""表格中,双击""已扫描""表格中的数据行可以将数据退回到""待登记""表格中"
      ForeColor       =   &H00004000&
      Height          =   465
      Left            =   90
      TabIndex        =   23
      Top             =   7560
      Width           =   5310
   End
End
Attribute VB_Name = "frmLabSampleCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnUse As Boolean                       '当前批次是否使用
Private mstrPrivs As String
Private mlngBatch As Long
Private mlngSampleCount As Long                  '本批标本总数
Private mObjSelectVSF As VSFlexGrid              '单击的VSF控件
Private mstrFind As String
Private WithEvents mfrmFind As frmLabSampleCheckFind
Attribute mfrmFind.VB_VarHelpID = -1

Public Sub ShowME(Objfrm As Object)
    Me.Show vbModal, Objfrm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 6 Then
        Call cmdFind_Click
    End If
End Sub

Private Sub mfrmFind_Finded(ByVal blnFind As Boolean, ByVal strVale As String)
    '定位:
    Dim varTmp As Variant, strSampleCode As String
    If blnFind Then
        varTmp = Split(strVale, ",")
        strSampleCode = varTmp(0)
        Call findSample(strSampleCode)
'        Call RptItem_SelectionChanged
    End If
End Sub

Private Sub saveSample()
    '获取标本号
    Dim i As Integer
    Dim strSampleIDs As String
    
    With Me.vsfList(1)
        If .Rows <= 1 Then Exit Sub
        
        For i = 1 To .Rows - 1
            strSampleIDs = strSampleIDs & .TextMatrix(i, .ColIndex("医嘱ID")) & ","
        Next
    End With
    Call SaveRegister(strSampleIDs, Me.vsfList(1))
End Sub

Private Function SaveRegister(ByVal strSampleIDs As String, objVsf As VSFlexGrid) As Boolean
    '签收标本       strSampleCodes-传入的医嘱id,以","分隔
    Dim var_Tmp As Variant
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim intTimeLimit As Integer         '送检时限单位分钟
    Dim blnTimeLimit As Boolean         '是否超过送检时限 true = 超过
    Dim strAdvice As String
    Dim blnShowMsg As Boolean
    Dim blnSave As Boolean              '是否强制通过
    Dim i As Integer
    Dim strArr() As String
    Dim blnTran As Boolean
    Dim strErr As String
    Dim arrSql() As String
    
    On Error GoTo ErrHand
    
    var_Tmp = Split(strSampleIDs, ",")
    blnShowMsg = True
    blnSave = False
    For i = 0 To UBound(var_Tmp) - 1
        If Chk划价费用(Me, CStr(var_Tmp(i)), 0) = False Then
            MsgBox var_Tmp(i) & "没有划价", vbInformation, "提示"
            Exit Function
        End If
    Next
    
    With objVsf
        '对已扫描的医嘱进行排序再登记
        If .Rows > 1 Then
            .Cell(flexcpSort, 1, .ColIndex("医嘱ID"), .Rows - 1, .ColIndex("医嘱ID")) = flexSortNumericAscending
            .Cell(flexcpSort, 1, .ColIndex("相关ID"), .Rows - 1, .ColIndex("相关ID")) = flexSortNumericAscending
        End If
        
        For i = 1 To .Rows - 1
            '处理是否超过采集时限
            strSQL = "select 送检时限 from 检验项目选项 where 诊疗项目id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(.TextMatrix(i, .ColIndex("诊疗项目ID"))))
            If rsTmp.EOF = True Then
                intTimeLimit = 0
            Else
                intTimeLimit = Val(Nvl(rsTmp("送检时限")))
            End If
            
            If IsDate(.TextMatrix(i, .ColIndex("采样时间"))) = False And intTimeLimit > 0 Then
                blnTimeLimit = True
            Else
                If IsDate(.TextMatrix(i, .ColIndex("采样时间"))) = True Then
                    If DateDiff("n", .TextMatrix(i, .ColIndex("采样时间")), zlDatabase.Currentdate) > intTimeLimit _
                        And intTimeLimit > 0 Then
                        '超过送检时限
                        blnTimeLimit = True
                    End If
                Else
                    If intTimeLimit > 0 Then
                        blnTimeLimit = True
                    End If
                End If
            End If
            
            If blnTimeLimit = True Then
                '超时处理，查看是否有权限，有权限时只提示
                If InStr(mstrPrivs, "强制通过送检时限") > 0 Then
                    If blnShowMsg = True Then
                        '提示
                        If MsgBox("本批采样时间为《" & .TextMatrix(i, .ColIndex("采样时间")) & "》" & vbCrLf & _
                                "已超过采样时限" & intTimeLimit & "分钟,送检延迟！" & vbCrLf & _
                                "您有强制通过送检时限权限" & vbCrLf & _
                                "是否强制通过?", vbQuestion + vbYesNo) = vbYes Then
                            blnSave = True
                        End If
                        blnShowMsg = False
                    End If
                    If blnSave = True Then
                        strAdvice = strAdvice & "|" & .TextMatrix(i, .ColIndex("医嘱ID")) & "," & .TextMatrix(i, .ColIndex("相关ID"))
                        Call vsfDataToVsfData(.TextMatrix(i, .ColIndex("条码")), objVsf, Me.vsfList(2))
                        .RowHidden(i) = True
                    End If
                Else
                    '拒绝登记
                    If blnShowMsg = True Then
                        MsgBox ("本批采样时间为《" & .TextMatrix(i, .ColIndex("采样时间")) & "》" & vbCrLf & _
                                "已超过采样时限" & intTimeLimit & "分钟,不允许登记！")
                        blnShowMsg = False
                    End If
                End If

            ElseIf .TextMatrix(i, .ColIndex("采样时间")) = "" Then
                '处理强制登记未采样标本
                If InStr(mstrPrivs, "强制登记未采样标本") > 0 Then
                    '提示
                    If blnShowMsg = True Then
'                        If MsgBox("当前《" & .TextMatrix(i, .ColIndex("申请项目")) & "》未采样!", vbInformation + vbQuestion) = vbYes Then
'                            blnSave = True
'                        End If
'                        If blnSave = True Then
                            strAdvice = strAdvice & "|" & .TextMatrix(i, .ColIndex("医嘱ID")) & "," & .TextMatrix(i, .ColIndex("相关ID"))
'                        End If
                        blnShowMsg = False
                    End If
                Else
                    '拒绝登记
                    If blnShowMsg = True Then
                        MsgBox "当前《" & .TextMatrix(i, .ColIndex("申请项目")) & "》未采样,不允许登记！", vbInformation
                        blnShowMsg = False
                    End If
                End If
            Else
                strAdvice = strAdvice & "|" & .TextMatrix(i, .ColIndex("医嘱ID")) & "," & .TextMatrix(i, .ColIndex("相关ID"))
                Call vsfDataToVsfData(.TextMatrix(i, .ColIndex("条码")), objVsf, Me.vsfList(2))
                .RowHidden(i) = True
            End If
        Next
    End With
    Call RemoveHiddenItem(objVsf)
    Call showNum
    
    '登记
    If strAdvice <> "" Then
        If mblnUse = True Or mlngBatch = 0 Then
            '得到一个新的批次
            strSQL = "select 病人医嘱发送_接收批次.nextval from dual "
            zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
            mlngBatch = rsTmp(0)
            mblnUse = False
        End If
        
        strArr = Str2Array(Mid(strAdvice, 2), "|", 4000)
        ReDim arrSql(UBound(strArr))
        
        For i = 0 To UBound(strArr)
            strSQL = "Zl_病人医嘱发送_SampleInput('" & strArr(i) & "','" & UserInfo.姓名 & "','" & mlngBatch & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            arrSql(i) = strSQL
        Next
        
        gcnOracle.BeginTrans
        blnTran = True
        For i = 0 To UBound(arrSql)
            If arrSql(i) <> "" Then
                zlDatabase.ExecuteProcedure arrSql(i), gstrSysName
            End If
        Next
        gcnOracle.CommitTrans
        blnTran = False
        
        '登记信息写入LIS
        Call WriterCheckSampleToLIS(strAdvice, UserInfo.姓名, mlngBatch, strErr)
        
        '写入LIS报错，取消之前的登记
        If strErr <> "" Then
            For i = 0 To UBound(strArr)
                strSQL = "Zl_病人医嘱发送_SampleInput('" & strArr(i) & "',NULL," & mlngBatch & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                arrSql(i) = strSQL
            Next
            
            gcnOracle.BeginTrans
            blnTran = True
            For i = 0 To UBound(arrSql)
                If arrSql(i) <> "" Then
                    zlDatabase.ExecuteProcedure arrSql(i), gstrSysName
                End If
            Next
            gcnOracle.CommitTrans
            blnTran = False
            
            MsgBox strErr, vbInformation, "标本登记"
            Exit Function
        End If
        
        mblnUse = True
    End If
    
    SaveRegister = True
    
    Exit Function
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub WriterCheckSampleToLIS(strAdvices As String, strName As String, strBatchNO As Long, Optional strError As String)
    '功能   把签收信息写入LIS
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.SampleCheckinInfoWrite(strAdvices, strName, strBatchNO, strErr) = False Then
            strError = "写入签收信息到LIS申请单出错!" & vbCrLf & strErr
        End If
    End If
End Sub

Private Sub cmdFind_Click()
    Dim strFindSQL As String, strFindFiled As String
    If mfrmFind Is Nothing Then Set mfrmFind = New frmLabSampleCheckFind
    strFindSQL = "select 条码 from (Select Distinct a.医嘱id, a.样本条码 条码, b.姓名, b.标本部位 As 标本, " & _
                 " b.医嘱内容 申请项目, a.送检人, c.名称 送检科室, b.诊疗项目id, a.采样时间, a.标本送出时间 送检时间," & _
                 " a.标本发送批号 , a.接收人, a.接收时间 From 病人医嘱发送 A, 病人医嘱记录 B, 部门表 C" & _
                 " Where a.医嘱id = b.Id And a.执行部门id = c.Id And a.执行状态 In (0) And" & _
                 " a.标本发送批号 In (Select 标本发送批号 From 病人医嘱发送 Where 样本条码 =100008332023)) where" & _
                 " 条码 like [1] or 姓名 like [1] or 标本 LIKE [1] or 申请项目 like [1]"
    Call mfrmFind.ShowFind(strFindSQL)
End Sub

Private Sub cmdQuet_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call saveSample
End Sub

Private Sub Form_Load()
    Call vsfSeting(Me.vsfList(0), 0)
    Call vsfSeting(Me.vsfList(1), 1)
    Call vsfSeting(Me.vsfList(2), 2)
    mstrPrivs = gstrPrivs       '初使化权限
    
    If mobjLisInsideComm Is Nothing Then
        Dim strErr As String
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        '初始化LIS接口部件
        If Not mobjLisInsideComm Is Nothing Then
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "初始化LIS接口失败！" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If
    
End Sub

Private Sub vsfSeting(ByVal objVsf As VSFlexGrid, Optional Index As Integer)
    Dim intFontSize As Integer
    Dim lbl As Label, lblInto As Label
    
    intFontSize = 11
    With objVsf
        .Clear
        .FixedCols = 0
        .Cols = 16
        .Rows = 1
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShow
        .ColWidth(0) = 1
        .TextMatrix(0, 1) = "条码": .ColKey(1) = "条码": .ColWidth(.ColIndex("条码")) = 2000: .Cell(flexcpAlignment, 0, .ColIndex("条码")) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "姓名": .ColKey(2) = "姓名": .ColWidth(.ColIndex("姓名")) = 1200: .Cell(flexcpAlignment, 0, .ColIndex("姓名")) = flexAlignCenterCenter
        .TextMatrix(0, 3) = "性别": .ColKey(3) = "性别": .ColWidth(.ColIndex("性别")) = 1200: .Cell(flexcpAlignment, 0, .ColIndex("性别")) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "标本": .ColKey(4) = "标本": .ColWidth(.ColIndex("标本")) = 1000: .Cell(flexcpAlignment, 0, .ColIndex("标本")) = flexAlignCenterCenter
        .TextMatrix(0, 5) = "申请项目": .ColKey(5) = "申请项目": .Cell(flexcpAlignment, 0, .ColIndex("申请项目")) = flexAlignCenterCenter
        .TextMatrix(0, 6) = "送检人": .ColKey(6) = "送检人": .Cell(flexcpAlignment, 0, .ColIndex("送检人")) = flexAlignCenterCenter: .ColHidden(.ColIndex("送检人")) = True
        .TextMatrix(0, 7) = "送检科室": .ColKey(7) = "送检科室": .Cell(flexcpAlignment, 0, .ColIndex("送检科室")) = flexAlignCenterCenter: .ColHidden(.ColIndex("送检科室")) = True
        .TextMatrix(0, 8) = "送检时间": .ColKey(8) = "送检时间": .Cell(flexcpAlignment, 0, .ColIndex("送检时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("送检时间")) = True
        .TextMatrix(0, 9) = "接收人": .ColKey(9) = "接收人": .Cell(flexcpAlignment, 0, .ColIndex("接收人")) = flexAlignCenterCenter: .ColHidden(.ColIndex("接收人")) = True
        .TextMatrix(0, 10) = "接收时间": .ColKey(10) = "接收时间": .Cell(flexcpAlignment, 0, .ColIndex("接收时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("接收时间")) = True
        .TextMatrix(0, 11) = "医嘱ID": .ColKey(11) = "医嘱ID": .Cell(flexcpAlignment, 0, .ColIndex("医嘱ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("医嘱ID")) = True
        .TextMatrix(0, 12) = "诊疗项目ID": .ColKey(12) = "诊疗项目ID": .Cell(flexcpAlignment, 0, .ColIndex("诊疗项目ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("诊疗项目ID")) = True
        .TextMatrix(0, 13) = "采样时间": .ColKey(13) = "采样时间": .Cell(flexcpAlignment, 0, .ColIndex("采样时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("采样时间")) = True
        .TextMatrix(0, 14) = "紧急": .ColKey(14) = "紧急": .ColWidth(.ColIndex("紧急")) = 1200: .Cell(flexcpAlignment, 0, .ColIndex("紧急")) = flexAlignCenterCenter
        .TextMatrix(0, 15) = "相关ID": .ColKey(15) = "相关ID": .Cell(flexcpAlignment, 0, .ColIndex("相关ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("相关ID")) = True
        .Cell(flexcpAlignment, 0, 1) = 3 '标题居中对齐
        .BackColorBkg = vbWhite
        .FontSize = intFontSize
    End With
    For Each lbl In Me.lbl
        lbl.FontSize = intFontSize
    Next
    For Each lblInto In Me.lblInto
        lblInto.FontSize = intFontSize
    Next
    Me.txtSampleCode.FontSize = intFontSize
    Me.txtSampleCode.Height = Me.lbl(0).Height
    
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrHand
    Me.txtSampleCode.Left = Me.lbl(0).Left + Me.lbl(0).Width + 100
    Me.lblInto(0).Move Me.lbl(1).Left + Me.lbl(1).Width + 100, Me.lbl(1).Top
'    With Me.Line1(0)
'        .X1 = Me.lblInto(0).Left
'        .X2 = Me.lblInto(0).Left + Me.lblInto(0).Width
'        .Y1 = Me.lbl(1).Top + Me.lbl(1).Height + 50
'        .Y2 = .Y1
'    End With
    Me.lblInto(1).Move Me.lbl(2).Left + Me.lbl(2).Width + 100, Me.lblInto(0).Top
'    With Me.Line1(1)
'        .X1 = Me.lblInto(1).Left
'        .X2 = Me.lblInto(1).Left + Me.lblInto(1).Width
'        .Y1 = Me.lbl(2).Top + Me.lbl(2).Height + 50
'        .Y2 = .Y1
'    End With
    Me.lblInto(2).Move Me.lbl(3).Left + Me.lbl(3).Width + 100, Me.lblInto(0).Top
'    With Me.Line1(2)
'        .X1 = Me.lblInto(2).Left
'        .X2 = Me.lblInto(2).Left + Me.lblInto(2).Width
'        .Y1 = Me.lbl(3).Top + Me.lbl(3).Height + 50
'        .Y2 = .Y1
'    End With
    Me.lblInto(1).Left = Me.lbl(2).Left + Me.lbl(2).Width + 100
    Me.lblInto(2).Left = Me.lbl(3).Left + Me.lbl(3).Width + 100
    Me.StatusBar.Panels(1).Width = Me.Width
    
    Me.picTop.Move 0, 0, Me.Width
    Me.PicMain.Move 0, Me.picTop.Height, Me.picTop.Width
    Exit Sub
ErrHand:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnUse = False
    mlngBatch = 0
    mlngSampleCount = 0
End Sub

Private Sub fraNS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.fraNS.Top = Me.fraNS.Top + Y
        Call picMain_Resize
    End If
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error GoTo ErrHand
    Select Case Index
        Case 0
            Me.lbl(6).Move 50, 50
            Me.vsfList(2).Move 50, Me.lbl(6).Top + Me.lbl(6).Height + 50, Me.pic(Index).Width - 150, Me.pic(Index).Height - Me.lbl(6).Height - 100
        Case 1
            Me.lbl(5).Move 50, 50
            Me.vsfList(1).Move 50, Me.lbl(5).Top + Me.lbl(5).Height + 50, Me.pic(Index).Width - 150, Me.pic(Index).Height - Me.lbl(5).Height - 100
        Case 2
            Me.lbl(4).Move 50, 50
            Me.vsfList(0).Move 50, Me.lbl(4).Top + Me.lbl(4).Height + 50, Me.pic(Index).Width, Me.pic(Index).Height - Me.lbl(4).Height - 100
    End Select
    Exit Sub
ErrHand:
    
End Sub

Private Sub picMain_Resize()
    On Error GoTo ErrHand
    Me.pic(2).Move 0, 0, Me.PicMain.Width / 2 - 10, Me.PicMain.Height
    Me.pic(1).Move Me.PicMain.Width / 2 + 10, 0, Me.PicMain.Width / 2 - 60, Me.fraNS.Top
    Me.pic(0).Move Me.PicMain.Width / 2 + 10, Me.fraNS.Top + Me.fraNS.Height, Me.PicMain.Width / 2 - 60, Me.PicMain.Height - Me.fraNS.Top - Me.fraNS.Height
    Exit Sub
ErrHand:
    
End Sub


Private Sub txtSampleCode_GotFocus()
   Call selectAll(Me.txtSampleCode)
End Sub

Private Sub selectAll(ByVal objTxt As TextBox)
    objTxt.SelStart = 0
    objTxt.SelLength = Len(objTxt.Text)
End Sub

Private Sub txtSampleCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call setVSFData(Trim(Me.txtSampleCode.Text))
            Call selectAll(Me.txtSampleCode)
    End Select
End Sub


Private Sub findSample(ByVal strSampleCode As String)
    Dim i As Long
    
    With Me.vsfList(0)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("条码")) = strSampleCode Then
                Me.vsfList(1).Select 0, 1
                Me.vsfList(2).Select 0, 1
                .Select i, 1, i, 4
                .ShowCell i, 1
            End If
        Next
    End With
    With Me.vsfList(1)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("条码")) = strSampleCode Then
                Me.vsfList(0).Select 0, 1
                Me.vsfList(2).Select 0, 1
                .Select i, 1, i, 4
                .ShowCell i, 1
            End If
        Next
    End With
    With Me.vsfList(2)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("条码")) = strSampleCode Then
                Me.vsfList(0).Select 0, 1
                Me.vsfList(1).Select 0, 1
                .Select i, 1, i, 4
                .ShowCell i, 1
            End If
        Next
    End With
End Sub

Private Function setVSFData(ByVal strSampleCode As String) As Boolean
    '功能               绑定数据到VSF
    'strSampleCode      扫描的条码
    Dim strSampleCodesLeft As String '左边VSF的所有条码
    Dim strSampleCodesRight As String '右边VSF的所有条码
    Dim strSampleCodesYDJ As String     '已登记或已核收条码
    Dim var_Tmp As Variant
    Dim rsData As Recordset
    Dim i As Integer, j As Integer
    
    Set rsData = ReadData(strSampleCode)
    mlngSampleCount = rsData.RecordCount
    
    If rsData.EOF = True Then
        MsgBox "条码不正确或者本批已全部登记,请检查    ", vbInformation, "提示"
        Exit Function
    End If
    
    With Me.vsfList(0)
        For i = 1 To .Rows - 1
            strSampleCodesLeft = strSampleCodesLeft & .TextMatrix(i, .ColIndex("条码")) & ","
        Next
    End With
    With Me.vsfList(1)
        For i = 1 To .Rows - 1
            strSampleCodesRight = strSampleCodesRight & .TextMatrix(i, .ColIndex("条码")) & ","
        Next
    End With
    With Me.vsfList(2)
        For i = 1 To .Rows - 1
            strSampleCodesYDJ = strSampleCodesYDJ & .TextMatrix(i, .ColIndex("条码")) & ","
        Next
    End With

    If InStr(strSampleCodesLeft, strSampleCode & ",") > 0 And InStr(strSampleCodesRight, strSampleCode & ",") = 0 Then
        '将左边VSF已扫描的条码加入到右边VSF
               
        With Me.vsfList(0)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("条码")) = strSampleCode Then
                    
                    
                    With Me.vsfList(1)
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, .ColIndex("条码")) = vsfList(0).TextMatrix(i, .ColIndex("条码")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("条码")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("姓名")) = vsfList(0).TextMatrix(i, .ColIndex("姓名")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("姓名")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("性别")) = vsfList(0).TextMatrix(i, .ColIndex("性别")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("性别")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("标本")) = vsfList(0).TextMatrix(i, .ColIndex("标本")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("标本")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = vsfList(0).TextMatrix(i, .ColIndex("申请项目")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("申请项目")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("送检人")) = vsfList(0).TextMatrix(i, .ColIndex("送检人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检人")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("送检科室")) = vsfList(0).TextMatrix(i, .ColIndex("送检科室")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检科室")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("送检时间")) = vsfList(0).TextMatrix(i, .ColIndex("送检时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检时间")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("接收人")) = vsfList(0).TextMatrix(i, .ColIndex("接收人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收人")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("接收时间")) = vsfList(0).TextMatrix(i, .ColIndex("接收时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收时间")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("医嘱ID")) = vsfList(0).TextMatrix(i, .ColIndex("医嘱ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("医嘱ID")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("相关ID")) = vsfList(0).TextMatrix(i, .ColIndex("相关ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("相关ID")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("诊疗项目ID")) = vsfList(0).TextMatrix(i, .ColIndex("诊疗项目ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("诊疗项目ID")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("采样时间")) = vsfList(0).TextMatrix(i, .ColIndex("采样时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("采样时间")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("紧急")) = vsfList(0).TextMatrix(i, .ColIndex("紧急")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("紧急")) = flexAlignLeftCenter
                        Me.lblInto(0).Caption = vsfList(1).TextMatrix(.Rows - 1, .ColIndex("送检科室"))
                        Me.lblInto(1).Caption = vsfList(1).TextMatrix(.Rows - 1, .ColIndex("送检人"))
                        Me.lblInto(2).Caption = vsfList(1).TextMatrix(.Rows - 1, .ColIndex("送检时间"))
                    End With
                    .RowHidden(i) = True
                    Call Form_Resize
                End If
            Next
            Call RemoveHiddenItem(Me.vsfList(0))
'            Me.lbl(4).Caption = "待登记标本(" & Me.vsfList(0).Rows - 1 & ")"
'            Me.lbl(5).Caption = "已扫描标本(" & Me.vsfList(1).Rows - 1 & ")"
            Call showNum
            Me.StatusBar.Panels(1).Text = "本批标本个数:" & mlngSampleCount & "个"
        End With
    ElseIf InStr(strSampleCodesRight, strSampleCode & ",") = 0 And InStr(strSampleCodesYDJ, strSampleCode & ",") = 0 Then
        '绑定数据
'        If Me.vsfList(0).Rows > 1 Then
'            If MsgBox("该条码不属于本批次,是否放弃现有批次扫面新批次?    ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                Exit Function
'            End If
'        End If
'        '初始化表格
'        Call vsfSeting(vsfList(0), 0)
'        Call vsfSeting(vsfList(1), 1)
'        Call vsfSeting(vsfList(2), 2)
        '绑定数据
        For i = 1 To rsData.RecordCount
            If IIf(IsNull(rsData("接收人")), "", rsData("接收人")) = "" Then    '未登记
                With Me.vsfList(0)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("条码")) = rsData("条码"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("条码")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsData("姓名"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("姓名")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsData("性别"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("性别")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("标本")) = rsData("标本"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("标本")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsData("申请项目"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("申请项目")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检人")) = IIf(IsNull(rsData("送检人")), "", rsData("送检人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检人")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检科室")) = rsData("送检科室"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检科室")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检时间")) = rsData("送检时间"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("接收人")) = IIf(IsNull(rsData("接收人")), "", rsData("接收人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收人")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("接收时间")) = IIf(IsNull(rsData("接收时间")), "", rsData("接收时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("医嘱ID")) = IIf(IsNull(rsData("医嘱ID")), "", rsData("医嘱ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("医嘱ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("相关ID")) = IIf(IsNull(rsData("相关ID")), "", rsData("相关ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("相关ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("诊疗项目ID")) = IIf(IsNull(rsData("诊疗项目ID")), "", rsData("诊疗项目ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("诊疗项目ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("采样时间")) = IIf(IsNull(rsData("采样时间")), "", rsData("采样时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("采样时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("紧急")) = IIf(IsNull(rsData("紧急")), "", rsData("紧急")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("紧急")) = flexAlignLeftCenter
                End With
            ElseIf IIf(IsNull(rsData("接收人")), "", rsData("接收人")) <> "" Then   '已登记
                With Me.vsfList(2)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("条码")) = rsData("条码"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("条码")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsData("姓名"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("姓名")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsData("性别"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("性别")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("标本")) = rsData("标本"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("标本")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsData("申请项目"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("申请项目")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检人")) = IIf(IsNull(rsData("送检人")), "", rsData("送检人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检人")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检科室")) = rsData("送检科室"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检科室")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检时间")) = rsData("送检时间"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("接收人")) = IIf(IsNull(rsData("接收人")), "", rsData("接收人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收人")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("接收时间")) = IIf(IsNull(rsData("接收时间")), "", rsData("接收时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("医嘱ID")) = IIf(IsNull(rsData("医嘱ID")), "", rsData("医嘱ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("医嘱ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("相关ID")) = IIf(IsNull(rsData("相关ID")), "", rsData("相关ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("相关ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("诊疗项目ID")) = IIf(IsNull(rsData("诊疗项目ID")), "", rsData("诊疗项目ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("诊疗项目ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("采样时间")) = IIf(IsNull(rsData("采样时间")), "", rsData("采样时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("采样时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("紧急")) = IIf(IsNull(rsData("紧急")), "", rsData("紧急")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("紧急")) = flexAlignLeftCenter
                End With
            End If
            rsData.MoveNext
        Next
        Call showNum
        Me.StatusBar.Panels(1).Text = "本批标本个数:" & mlngSampleCount & "个"
        Call setVSFData(strSampleCode)
    ElseIf InStr(strSampleCodesRight, strSampleCode & ",") > 0 Then
        MsgBox "该条码已经在已扫描条码区   ", vbInformation
        With Me.vsfList(1)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("条码")) = strSampleCode Then
                    .Select i, 1
                    .ShowCell i, 1
                End If
            Next
        End With
        Exit Function
    ElseIf InStr(strSampleCodesYDJ, strSampleCode & ",") > 0 Then
        MsgBox "该条码已登记或已核收   ", vbInformation
        With Me.vsfList(2)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("条码")) = strSampleCode Then
                    .Select i, 1
                    .ShowCell i, 1
                End If
            Next
        End With
        Exit Function
    End If
    
End Function

Private Sub showNum()
    Me.lbl(4).Caption = "待登记标本(" & Me.vsfList(0).Rows - 1 & ")"
    Me.lbl(5).Caption = "已扫描标本(" & Me.vsfList(1).Rows - 1 & ")"
    Me.lbl(6).Caption = "本批已登记或已核收标本(" & Me.vsfList(2).Rows - 1 & ")"
End Sub

Private Function ReadData(ByVal strSampleCode As String) As Recordset
    '功能               返回扫描条码对应发送批号下的所有条码
    'strSampleCode      扫描的条码
    Dim strSQL As String
    Dim rsSampleCodes As Recordset
    
    strSQL = "Select Distinct b.相关id, b.Id 医嘱id, a.样本条码 条码, b.姓名, b.性别, b.标本部位 As 标本, b.医嘱内容 申请项目, a.送检人, c.名称 送检科室, b.诊疗项目id, a.采样时间," & vbNewLine & _
            "                a.标本送出时间 送检时间, a.标本发送批号, a.接收人, a.接收时间, Decode(b.紧急标志, 1, '紧急', '') As 紧急" & vbNewLine & _
            "From 病人医嘱发送 A, 病人医嘱记录 B, 部门表 C, 诊疗项目目录 D" & vbNewLine & _
            "Where a.医嘱id = b.Id And a.执行部门id = c.Id And b.诊疗项目id = d.Id And d.类别 = 'C' And a.执行状态 In (0) And" & vbNewLine & _
            "      a.标本发送批号 In (Select 标本发送批号 From 病人医嘱发送 Where 样本条码 = [1])" & vbNewLine & _
            "Order By Nvl(b.相关id, b.Id), 医嘱id"
    
    Set rsSampleCodes = zlDatabase.OpenSQLRecord(strSQL, "批量检查条码", strSampleCode)
        
    Set ReadData = rsSampleCodes
End Function

Private Sub vsfList_Click(Index As Integer)
    Select Case Index
        Case 0
            Set mObjSelectVSF = Me.vsfList(0)
        Case 1
            Set mObjSelectVSF = Me.vsfList(1)
        Case 2
            Set mObjSelectVSF = Me.vsfList(2)
    End Select
End Sub

Private Sub vsfList_DblClick(Index As Integer)
    Dim strSampleCode As String
    
    With vsfList(Index)
        If .MouseRow > 0 Then
            strSampleCode = .TextMatrix(.MouseRow, .ColIndex("条码"))
            
            Select Case Index
                Case 0
                    Call vsfDataToVsfData(strSampleCode, Me.vsfList(0), Me.vsfList(1))
                    Call RemoveHiddenItem(Me.vsfList(0))
                Case 1
                    Call vsfDataToVsfData(strSampleCode, Me.vsfList(1), Me.vsfList(0))
                    Call RemoveHiddenItem(Me.vsfList(1))
            End Select
        
            Call showNum
            Me.StatusBar.Panels(1).Text = "本批标本个数:" & mlngSampleCount & "个"
        End If
    End With
End Sub

Private Sub vsfDataToVsfData(ByVal strSampleCode As String, objVSFFrom As VSFlexGrid, objVSFTo As VSFlexGrid)
    '将数据从一个VSF转移到另一个VSF
    'strSampleCode-用于匹配的条码
    'indexFrom-数据来源的VSF索引
    'indexTo-要添加数据的VSF索引
    
    Dim i As Long
    With objVSFFrom
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("条码")) = strSampleCode And .RowHidden(i) = False Then
                With objVSFTo
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("条码")) = objVSFFrom.TextMatrix(i, .ColIndex("条码")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("条码")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("姓名")) = objVSFFrom.TextMatrix(i, .ColIndex("姓名")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("姓名")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("性别")) = objVSFFrom.TextMatrix(i, .ColIndex("性别")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("性别")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("标本")) = objVSFFrom.TextMatrix(i, .ColIndex("标本")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("标本")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = objVSFFrom.TextMatrix(i, .ColIndex("申请项目")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("申请项目")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检人")) = objVSFFrom.TextMatrix(i, .ColIndex("送检人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检人")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检科室")) = objVSFFrom.TextMatrix(i, .ColIndex("送检科室")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检科室")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("送检时间")) = objVSFFrom.TextMatrix(i, .ColIndex("送检时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("送检时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("接收人")) = objVSFFrom.TextMatrix(i, .ColIndex("接收人")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收人")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("接收时间")) = objVSFFrom.TextMatrix(i, .ColIndex("接收时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("接收时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("医嘱ID")) = objVSFFrom.TextMatrix(i, .ColIndex("医嘱ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("医嘱ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("相关ID")) = objVSFFrom.TextMatrix(i, .ColIndex("相关ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("相关ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("诊疗项目ID")) = objVSFFrom.TextMatrix(i, .ColIndex("诊疗项目ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("诊疗项目ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("采样时间")) = objVSFFrom.TextMatrix(i, .ColIndex("采样时间")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("采样时间")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("紧急")) = objVSFFrom.TextMatrix(i, .ColIndex("紧急")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("紧急")) = flexAlignLeftCenter
                End With
                .RowHidden(i) = True
                Exit For
            End If
        Next
    End With
'    Call RemoveHiddenItem(objVSFFrom)
End Sub

Private Sub RemoveHiddenItem(objVsf As VSFlexGrid)
    Dim i As Long
begin:
    With objVsf
        For i = 1 To .Rows - 1
            If .RowHidden(i) = True Then
                .RemoveItem i
                GoTo begin
            End If
        Next
    End With
End Sub

Private Sub VSFList_RowColChange(Index As Integer)
    With Me.vsfList(Index)
        If .Rows > 1 Then
            If .TextMatrix(1, 1) <> "" Then
                Me.lblInto(0).Caption = .TextMatrix(.RowSel, .ColIndex("送检科室"))
                Me.lblInto(1).Caption = .TextMatrix(.RowSel, .ColIndex("送检人"))
                Me.lblInto(2).Caption = .TextMatrix(.RowSel, .ColIndex("送检时间"))
                Call Form_Resize
            End If
        End If
    End With
End Sub

