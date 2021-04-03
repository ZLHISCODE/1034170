VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBalanceDeposit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结帐三方预交退款"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   Icon            =   "frmBalanceDeposit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   825
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   4335
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4965
      TabIndex        =   1
      ToolTipText     =   "热键：F2"
      Top             =   4305
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6405
      TabIndex        =   2
      ToolTipText     =   "热键:Esc"
      Top             =   4305
      Width           =   1410
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDeposit 
      Height          =   3600
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   7980
      _cx             =   14076
      _cy             =   6350
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceDeposit.frx":06EA
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
   Begin VB.Label lbl误差金额 
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3120
      TabIndex        =   7
      Top             =   4350
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl误差 
      AutoSize        =   -1  'True
      Caption         =   "误差费:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2175
      TabIndex        =   6
      Top             =   4350
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   330
      Picture         =   "frmBalanceDeposit.frx":0829
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "以下是本次结帐病人的三方卡预交情况,  请根据需要处理退款    "
      Height          =   360
      Left            =   885
      TabIndex        =   5
      Top             =   135
      Width           =   3420
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "退现"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   3
      Top             =   4395
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmBalanceDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsDeposit As ADODB.Recordset, mrsInfo As ADODB.Recordset
Private mblnUnload As Boolean
Private mlng结帐ID As Long, mlng病人ID As Long
Private mlngModul As Long, mblnAll As Boolean
Private mblnDateMoved As Boolean
Private mstr住院次数 As String
Private mstrDepositDate    As String
Private mint预交类别    As Integer
Private mstrCardPrivs As String, mstrForceNote As String
Private mstr强制退现操作员 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveData
End Sub

Private Sub SaveData()
    Dim i As Integer, cllSQL As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSql As String, strFailNo As String, strXMLExpend As String, dblMoney As Double
    Dim strCardNo As String, strPassWord As String, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllSquareBalance As Collection, strIDs As String, strNos As String, dbl误差 As Double
    Dim rsTmp As ADODB.Recordset, lngRow As Long, j As Integer, strValue As String
    Dim strInXML As String, strOutXML As String, strExpend As String, strBalanceIDs As String
    If lbl误差金额.Visible Then
        dbl误差 = Val(lbl误差金额.Caption)
    End If
    For i = 1 To vsfDeposit.Rows - 1
        Set cllSQL = New Collection
        Set cllSquareBalance = New Collection
        Set cllThreeSwap = New Collection
        With vsfDeposit
            If Val(.TextMatrix(i, .ColIndex("退款金额"))) <> 0 Then
                If Val(.TextMatrix(i, .ColIndex("退现"))) = 0 Then
                    If .TextMatrix(i, .ColIndex("转账")) = 1 Then
                        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
                                Val(.RowData(i)), False, _
                            mrsInfo!姓名, mrsInfo!性别, mrsInfo!年龄, Val(.TextMatrix(i, .ColIndex("退款金额"))), strCardNo, strPassWord, _
                            False, True, False, False, cllSquareBalance) = False Then
                            strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("结算卡名称"))
                        Else
                            zlXML.ClearXmlText
                            zlXML.AppendNode "IN"
                            zlXML.appendData "CZLX", "2"
                            zlXML.AppendNode "IN", True
                            strXMLExpend = zlXML.XmlText
                            zlXML.ClearXmlText
                            If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, Val(.RowData(i)), _
                                strCardNo, Val(.TextMatrix(i, .ColIndex("退款金额"))), "", strXMLExpend) = False Then
                                strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("结算卡名称"))
                            Else
                                mrsDeposit.Filter = "三方卡名称='" & .TextMatrix(i, .ColIndex("结算卡名称")) & "'"
                                dblMoney = Val(.TextMatrix(i, .ColIndex("退款金额")))
                                Do While Not mrsDeposit.EOF
                                    If dblMoney > 0 Then
                                        If dblMoney > Val(mrsDeposit!金额) Then
                                            strSql = "Zl_结帐预交记录_三方退款(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    Val(mrsDeposit!金额) & "," & mlng结帐ID & "," & mlng病人ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!结算方式) & "')"
                                            dblMoney = dblMoney - Val(mrsDeposit!金额)
                                        Else
                                            strSql = "Zl_结帐预交记录_三方退款(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    dblMoney & "," & mlng结帐ID & "," & mlng病人ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!结算方式) & "')"
                                            dblMoney = 0
                                        End If
                                        zlAddArray cllSQL, strSql
                                    End If
                                    mrsDeposit.MoveNext
                                Loop
                                zlExecuteProcedureArrAy cllSQL, Me.Caption, True
                                zlXML.ClearXmlText
                                zlXML.AppendNode "IN"
                                zlXML.appendData "CZLX", "2"
                                zlXML.AppendNode "IN", True
                                strXMLExpend = zlXML.XmlText
                                zlXML.ClearXmlText
                                If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, Val(.RowData(i)), _
                                    strCardNo, Val(.TextMatrix(i, .ColIndex("退款金额"))), "", strXMLExpend) = False Then
                                    gcnOracle.RollbackTrans
                                    strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("结算卡名称"))
                                Else
                                    zlXML.ClearXmlText
                                    zlXML.AppendNode "IN"
                                        zlXML.appendData "CZLX", "2"
                                    zlXML.AppendNode "IN", True
                                    strXMLExpend = zlXML.XmlText
                                    zlXML.ClearXmlText
                                    If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, Val(.RowData(i)), strCardNo, _
                                        mlng结帐ID, Val(.TextMatrix(i, .ColIndex("退款金额"))), strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
                                        gcnOracle.RollbackTrans
                                        strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("结算卡名称"))
                                    Else
                                        Set cllUpdate = New Collection
                                        Set cllThreeSwap = New Collection
    '                                    Call zlAddUpdateSwapSQL(False, mlng结帐ID, Val(.RowData(i)), False, strCardNo, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
                                        Call zlAddThreeSwapSQLToCollection(False, mlng结帐ID, Val(.RowData(i)), False, strCardNo, strSwapExtendInfor, cllThreeSwap)
                                        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
                                        zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, False, True
                                    End If
                                End If
                            End If
                        End If
                    Else
                        '多笔退款
                        strBalanceIDs = ""
                        zlXML.ClearXmlText
                        mrsDeposit.Filter = "三方卡名称='" & .TextMatrix(i, .ColIndex("结算卡名称")) & "'"
                        dblMoney = Val(.TextMatrix(i, .ColIndex("退款金额")))
                        Call zlXML.AppendNode("JSLIST")
                        Do While Not mrsDeposit.EOF
                            If dblMoney > 0 Then
                                If dblMoney > Val(mrsDeposit!金额) Then
                                    strSql = "Zl_结帐预交记录_三方退款(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    Val(mrsDeposit!金额) & "," & mlng结帐ID & "," & mlng病人ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!结算方式) & "')"
                                    zlAddArray cllSQL, strSql
                                    strSql = "Select ID,卡号,交易流水号,交易说明 From 病人预交记录 Where ID = [1]"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsDeposit!预交ID))
                                    If Not rsTmp.EOF Then
                                        Call zlXML.AppendNode("JS")
                                            Call zlXML.appendData("KH", Nvl(rsTmp!卡号))
                                            Call zlXML.appendData("JYLSH", Nvl(rsTmp!交易流水号))
                                            Call zlXML.appendData("JYSM", Nvl(rsTmp!交易说明))
                                            Call zlXML.appendData("ZFJE", Val(mrsDeposit!金额))
                                            Call zlXML.appendData("JSLX", 1)
                                            Call zlXML.appendData("ID", Nvl(rsTmp!ID))
                                        Call zlXML.AppendNode("JS", True)
                                        strSql = "Zl_三方退款信息_Insert("
                                        strSql = strSql & mlng结帐ID & ","
                                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                                        strSql = strSql & Val(mrsDeposit!金额) & ",'"
                                        strSql = strSql & Nvl(rsTmp!卡号) & "','"
                                        strSql = strSql & Nvl(rsTmp!交易流水号) & "','"
                                        strSql = strSql & Nvl(rsTmp!交易说明) & "')"
                                        zlAddArray cllThreeSwap, strSql
                                        strBalanceIDs = strBalanceIDs & "," & Val(Nvl(rsTmp!ID))
                                    End If
                                    dblMoney = dblMoney - Val(mrsDeposit!金额)
                                Else
                                    strSql = "Zl_结帐预交记录_三方退款(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    dblMoney & "," & mlng结帐ID & "," & mlng病人ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!结算方式) & "')"
                                    zlAddArray cllSQL, strSql
                                    strSql = "Select ID,卡号,交易流水号,交易说明 From 病人预交记录 Where ID = [1]"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsDeposit!预交ID))
                                    If Not rsTmp.EOF Then
                                        Call zlXML.AppendNode("JS")
                                            Call zlXML.appendData("KH", Nvl(rsTmp!卡号))
                                            Call zlXML.appendData("JYLSH", Nvl(rsTmp!交易流水号))
                                            Call zlXML.appendData("JYSM", Nvl(rsTmp!交易说明))
                                            Call zlXML.appendData("ZFJE", dblMoney)
                                            Call zlXML.appendData("JSLX", 1)
                                            Call zlXML.appendData("ID", Nvl(rsTmp!ID))
                                        Call zlXML.AppendNode("JS", True)
                                        strSql = "Zl_三方退款信息_Insert("
                                        strSql = strSql & mlng结帐ID & ","
                                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                                        strSql = strSql & dblMoney & ",'"
                                        strSql = strSql & Nvl(rsTmp!卡号) & "','"
                                        strSql = strSql & Nvl(rsTmp!交易流水号) & "','"
                                        strSql = strSql & Nvl(rsTmp!交易说明) & "')"
                                        zlAddArray cllThreeSwap, strSql
                                        strBalanceIDs = strBalanceIDs & "," & Val(Nvl(rsTmp!ID))
                                    End If
                                    dblMoney = 0
                                End If
                            End If
                            mrsDeposit.MoveNext
                        Loop
                        Call zlXML.AppendNode("JSLIST", True)
                        strXMLExpend = zlXML.XmlText
                        strInXML = zlXML.XmlText
                        If strBalanceIDs <> "" Then strBalanceIDs = "1|" & Mid(strBalanceIDs, 2)
                        
                        If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, Val(.RowData(i)), False, strCardNo, _
                            strBalanceIDs, Val(.TextMatrix(i, .ColIndex("退款金额"))), strSwapGlideNO, strSwapMemo, strXMLExpend) = False Then
                            strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("结算卡名称"))
                        Else
                            zlExecuteProcedureArrAy cllSQL, Me.Caption, True
                            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, True, True
                            If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, Val(.RowData(i)), False, strInXML, _
                                 mlng结帐ID, strOutXML, strExpend) = False Then
                                gcnOracle.RollbackTrans:
                                strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("结算卡名称"))
                            Else
                                '提交
                                Set cllThreeSwap = New Collection
                                If zlXML_Init = True Then
                                    If strOutXML <> "" Then
                                        If zlXML_LoadXMLToDOMDocument(strOutXML, False) Then
                                            Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
                                            For j = 0 To lngRow - 1
                                                Call zlXML_GetNodeValue("ID", i, strValue)
                                                strSql = "Zl_三方退款信息_Insert("
                                                strSql = strSql & mlng结帐ID & ","
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
                                            Next j
                                        End If
                                    End If
                                    
                                    If strExpend <> "" Then
                                        strSwapExtendInfor = ""
                                        If zlXML_LoadXMLToDOMDocument(strExpend, False) Then
                                            Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
                                            For j = 0 To lngRow - 1
                                                Call zlXML_GetNodeValue("XMMC", j, strValue)
                                                strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
                                                Call zlXML_GetNodeValue("XMNR", j, strValue)
                                                strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
                                            Next j
                                        End If
                                    End If
                                    If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
                                End If
                                strSql = "Select 卡号 From 病人预交记录 Where 结帐ID= [1] And 卡类别ID= [2]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng结帐ID, Val(.RowData(i)))
                                If Not rsTmp.EOF Then
                                    strCardNo = Nvl(rsTmp!卡号)
                                End If
    '                            Call zlAddUpdateSwapSQL(False, mlng结帐ID, Val(.RowData(i)), False, strCardNo, "", "", cllUpdate, 0)
                                Call zlAddThreeSwapSQLToCollection(False, mlng结帐ID, Val(.RowData(i)), False, strCardNo, strSwapExtendInfor, cllThreeSwap)
    '                            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
                                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, True, True
                                gcnOracle.CommitTrans
                            End If
                        End If
                    End If
                Else
                    '退现
                    mrsDeposit.Filter = "三方卡名称='" & .TextMatrix(i, .ColIndex("结算卡名称")) & "'"
                    dblMoney = Val(.TextMatrix(i, .ColIndex("退款金额")))
                    
                    Do While Not mrsDeposit.EOF
                        If dblMoney > 0 Then
                            If dblMoney > Val(mrsDeposit!金额) Then
                                strSql = "Zl_结帐预交记录_三方退款(" & Val(mrsDeposit!ID) & "," & _
                                        "'" & mrsDeposit!NO & "'" & ",1," & _
                                        Val(mrsDeposit!金额) & "," & mlng结帐ID & "," & mlng病人ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!结算方式) & "'" & IIf(lbl误差金额.Visible And dbl误差 <> 0, "," & dbl误差 & ",'", ",Null,'") & mstrForceNote & "')"
                                dblMoney = dblMoney - Val(mrsDeposit!金额)
                                If dbl误差 <> 0 And lbl误差.Visible Then
                                    dbl误差 = 0
                                End If
                            Else
                                strSql = "Zl_结帐预交记录_三方退款(" & Val(mrsDeposit!ID) & "," & _
                                        "'" & mrsDeposit!NO & "'" & ",1," & _
                                        dblMoney & "," & mlng结帐ID & "," & mlng病人ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!结算方式) & "'" & IIf(lbl误差金额.Visible And dbl误差 <> 0, "," & dbl误差 & ",'", ",Null,'") & mstrForceNote & "')"
                                dblMoney = 0
                                If dbl误差 <> 0 And lbl误差.Visible Then
                                    dbl误差 = 0
                                End If
                            End If
                            zlAddArray cllSQL, strSql
                        End If
                        mrsDeposit.MoveNext
                    Loop
                    zlExecuteProcedureArrAy cllSQL, Me.Caption
                End If
            End If
        End With
    Next i
    If strFailNo <> "" Then
        MsgBox "以下三方卡的预交款在退款过程中出现错误,请使用余额退款功能对该类预交款进行退款!" & vbCrLf & Mid(strFailNo, 2)
    End If
    mblnUnload = True
    Unload Me
End Sub

Public Sub ShowMe(frmMain As Object, lngModule As Long, lng结帐ID As Long, lng病人ID As Long, blnAll As Boolean, _
                  Optional ByVal blnDateMoved As Boolean = False, Optional ByVal str住院次数 As String = "", Optional ByVal strDepositDate As String = "", Optional ByVal int预交类别 As Integer)
    mlngModul = lngModule
    mlng结帐ID = lng结帐ID
    mlng病人ID = lng病人ID
    mblnAll = blnAll
    mblnDateMoved = blnDateMoved
    mstr住院次数 = str住院次数
    mstrDepositDate = strDepositDate
    mint预交类别 = int预交类别
    On Error Resume Next
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim i As Integer
    Dim lngRow As Long
    
    mblnUnload = False
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    
    With vsfDeposit
        .Clear 1: .Rows = 2
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
        Next
    End With
    Set mrsDeposit = GetThreeDeposit(mlng病人ID, mblnDateMoved, mstr住院次数, mstrDepositDate, mint预交类别)
    With vsfDeposit
        Do While Not mrsDeposit.EOF
            lngRow = 0
            For i = 1 To .Rows - 1
                If .RowData(i) = Nvl(mrsDeposit!卡类别ID) Then
                    lngRow = i
                    Exit For
                End If
            Next i
            If lngRow = 0 Then lngRow = .Rows - 1: .Rows = .Rows + 1
            
            .TextMatrix(lngRow, .ColIndex("结算卡名称")) = Nvl(mrsDeposit!三方卡名称)
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(mrsDeposit!结算方式)
            .TextMatrix(lngRow, .ColIndex("余额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("余额"))) + Val(Nvl(mrsDeposit!金额)), "0.00")
            
            If mblnAll Then
                .TextMatrix(lngRow, .ColIndex("退款金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("退款金额"))) + Val(Nvl(mrsDeposit!金额)), "0.00")
            Else
                .TextMatrix(lngRow, .ColIndex("退款金额")) = Format(0, "0.00")
            End If
            
            .TextMatrix(lngRow, 4) = 0
            If Val(mrsDeposit!退现) = 1 Then
                '允许退现,可以修改
                .Cell(flexcpData, lngRow, .ColIndex("退现")) = 1
                .Cell(flexcpBackColor, lngRow, .ColIndex("退现")) = vbWhite
            Else
                .Cell(flexcpData, lngRow, .ColIndex("退现")) = 0
                .Cell(flexcpBackColor, lngRow, .ColIndex("退现")) = &H8000000F
            End If
            
            .TextMatrix(lngRow, .ColIndex("预交ID")) = Nvl(mrsDeposit!预交ID)
            .TextMatrix(lngRow, .ColIndex("转账")) = Nvl(mrsDeposit!代扣)
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(mrsDeposit!ID)
            .TextMatrix(lngRow, .ColIndex("记录状态")) = Nvl(mrsDeposit!记录状态)
            
            .RowData(lngRow) = Nvl(mrsDeposit!卡类别ID)
            mrsDeposit.MoveNext
        Loop
    End With
    
    If mrsDeposit.RecordCount = 0 Then
        mblnUnload = True
        Unload Me: Exit Sub
    End If
    
    vsfDeposit.Rows = vsfDeposit.Rows - 1
    strSql = "Select 姓名,年龄,性别 From 病人信息 Where 病人ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnUnload = False Then
        If MsgBox("是否确定取消预交款退款?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    mstrForceNote = ""
    mstr强制退现操作员 = ""
    mblnUnload = False
End Sub

Private Sub vsfDeposit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    With vsfDeposit
        If Col = .ColIndex("退款金额") Then
            If IsNumeric(.TextMatrix(Row, .ColIndex("退款金额"))) = False Then
                MsgBox "请输入正确的退款金额!", vbInformation, gstrSysName
                 .TextMatrix(Row, 3) = "0.00"
            End If
            If Val(vsfDeposit.TextMatrix(Row, .ColIndex("退款金额"))) > Val(.TextMatrix(Row, .ColIndex("余额"))) Then
                MsgBox "输入的退款金额过大,请检查", vbInformation, gstrSysName
                 .TextMatrix(Row, .ColIndex("退款金额")) = "0.00"
            End If
             .TextMatrix(Row, .ColIndex("退款金额")) = Format(Val(.TextMatrix(Row, .ColIndex("退款金额"))), "0.00")
        End If
        Call RecalCash
        
        If Col = .ColIndex("退现") And Val(.Cell(flexcpData, Row, Col)) = 0 Then
            mstrForceNote = ""
            For i = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(i, .ColIndex("退现")))) = 1 Then
                    mstrForceNote = mstrForceNote & IIf(mstrForceNote = "", mstr强制退现操作员 & "强制退现:", ";") & .TextMatrix(i, 0) & "," & Format(.TextMatrix(i, 3), "0.00") & "元"
                End If
            Next i
        End If
    End With
End Sub

Private Sub vsfDeposit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblMoney As Double, lngRow As Long
    Dim str操作员姓名 As String, strDBUser As String
    Dim strPrivs As String
    With vsfDeposit
        If Col <> .ColIndex("退现") Then Exit Sub
        
        If Val(.Cell(flexcpData, Row, Col)) = 0 Then
            If InStr(";" & mstrCardPrivs & ";", ";三方退款强制退现;") = 0 Then
                If mstr强制退现操作员 = "" Then
                    mstr强制退现操作员 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
                    If mstr强制退现操作员 = "" Then
                        MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                If mstr强制退现操作员 = "" Then
                    If MsgBox("选择的结算卡不支持退现,是否强制将其退现？", _
                                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Cancel = True: Exit Sub
                    mstr强制退现操作员 = UserInfo.姓名
                End If
            End If
        End If
    End With
End Sub

Private Sub RecalCash()
    '重算现金金额
    Dim i As Integer, dblSum As Double
    Dim dbl误差 As Double, dbl实际 As Double
    Dim cur误差 As Currency
    dblSum = 0
    With vsfDeposit
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, .ColIndex("退现")))) = 1 Then
                dblSum = dblSum + Val(.TextMatrix(i, .ColIndex("退款金额")))
            End If
        Next i
    End With
    
    If dblSum = 0 Then
        txtMoney.Visible = False
        lblMoney.Visible = False
        lbl误差.Visible = False
        lbl误差金额.Visible = False
    Else
        txtMoney.Visible = True
        lblMoney.Visible = True
        dbl实际 = CentMoney(dblSum)
        cur误差 = Val(dblSum) - Val(dbl实际)
        txtMoney.Text = Format(dbl实际, "0.00")
        lbl误差.Visible = cur误差 <> 0
        lbl误差金额.Visible = cur误差 <> 0
        lbl误差金额.Caption = Format(cur误差, "0.######")
    End If
End Sub

Private Sub vsfDeposit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDeposit
        Select Case Col
        Case .ColIndex("退款金额")
        Case .ColIndex("退现")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Function GetThreeDeposit(lng病人ID As Long, _
    Optional blnDateMoved As Boolean, Optional strTime As String, _
    Optional ByVal strPepositDate As String = "", _
    Optional int预交类别 As Integer = 0) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人剩余预交款明细(三方卡)
    '入参:strTime-住院次数,如:1,2,3
    '        bln门诊转住院-是否门诊费用转住院(只能充指定的预交)
    '        strPepositDate-指定的预交日期
    '       int预交类别-0-门诊和住院;1-门诊;2- 住院
    '出参:
    '返回:预交明细数据
    '编制:刘尔旋
    '日期:2016-2-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strSub1 As String
    Dim strWherePage As String, strTable As String
    Dim strWhere As String, strDate As String
    On Error GoTo errH
    
    If int预交类别 = 1 Then strTime = ""    '69500
    
    strWherePage = IIf(strTime = "", "", " And instr(','||[2]||',',','||Nvl(A.主页ID,0)||',')>0")
    strTable = IIf(blnDateMoved, zlGetFullFieldsTable("病人预交记录"), "病人预交记录 A")
    strWhere = "": strDate = "1974-04-28 00:00:00"
    If strPepositDate <> "" Then
        If IsDate(strPepositDate) Then
            strDate = strPepositDate
            strWhere = " And A.收款时间=[3]"
        End If
    End If
    
    If int预交类别 <> 0 Then
        strWhere = strWhere & " And A.预交类别 =[4]"
    End If
    
    strSql = "" & _
    "    Select a.No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, " & vbCrLf & _
    "           Min(Decode(a.结帐id, Null, a.Id, 0) * Decode(a.记录状态, 1, 1, 0))*1 As ID, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0*Null, a.Id), 0)) As 预交id, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, '', a.实际票号), '')) As 票据号, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, '', To_Char(a.收款时间, 'yyyy-mm-dd hh24:mi:ss')), '')) As 日期, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, '', a.结算方式), '')) As 结算方式, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0*Null, a.卡类别id), 0)) As 卡类别id, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0*Null, a.结算卡序号), 0)) As 结算卡序号, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, '', a.卡号), '')) As 卡号, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, '', a.交易流水号), '')) As 交易流水号, " & vbCrLf & _
    "           Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, '', a.交易说明), '')) As 交易说明 " & vbCrLf & _
    "     From  病人预交记录 A" & vbCrLf & _
    "     Where a.记录性质 In (1, 11) And a.病人id = [1]   AND 卡类别ID IS NOT NULL " & strWherePage & strWhere & vbCrLf & _
    "           And ( nvl(金额,0)>=0 or a.记录状态<>1 or Not exists(select 1 from 病人预交记录 where 病人ID=[1] and a.卡类别id=卡类别ID and a.交易流水号=交易流水号 and 金额>=0 and 记录性质=1)) " & vbCrLf & _
    "     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0 " & vbCrLf & _
    "     Group By a.No "
    
    strSql = strSql & "" & vbCrLf & _
    "     UNION ALL  " & vbCrLf & _
    "    Select b.No, Sum(Nvl(a.金额, 0) - Nvl(a.冲预交, 0)) As 金额, 1 As 记录状态, 0 As ID, " & vbCrLf & _
    "             0 As 预交id, Min(b.实际票号) As 票据号,To_Char(Min(b.收款时间), 'yyyy-mm-dd hh24:mi:ss') As 日期, Min(b.结算方式) As 结算方式, " & vbCrLf & _
    "             Min(b.卡类别id) As 卡类别id, 0 * Null As 结算卡序号, Min(b.卡号) As 卡号, Min(b.交易流水号) As 交易流水号, Min(b.交易说明) As 交易说明" & vbCrLf & _
    "    From 病人预交记录 A, " & vbCrLf & _
    "          ( Select b.No, b.交易流水号, b.Id, b.收款时间, b.卡号, b.交易说明, b.卡类别id, b.结帐id, b.结算方式, b.实际票号 " & vbCrLf & _
    "            From 病人预交记录 B 　" & vbCrLf & _
    "            Where 病人id = [1] And 记录性质 = 1 And 记录状态 = 1 And 卡类别id Is Not Null and 金额>0 " & _
    "           ) B,医疗卡类别 C " & vbCrLf & _
    "     Where a.记录性质 In (1, 11) And a.病人id = [1] And a.卡类别id Is Not Null And a.金额 < 0 And a.卡类别id = b.卡类别id And  a.交易流水号 = b.交易流水号 " & strWherePage & strWhere & vbCrLf & _
    "            and a.卡类别ID=c.ID and nvl(c.是否转帐及代扣,0)=0  and a.NO not in (select NO From 病人预交记录 where 病人ID=[1] and mod(记录性质,10)=1 and 预交类别=1 having sum(nvl(金额,0))-sum(nvl(冲预交,0))=0 Group by NO) " & vbCrLf & _
    "     Group By b.No " & vbCrLf & _
    "     Having Sum(Nvl(a.金额, 0) - Nvl(a.冲预交, 0)) <> 0"
    
    strSql = "" & _
    " Select " & vbCrLf & _
    "       max( Decode(Nvl(a.卡类别id, 0), 0, 0, Decode(Nvl(c.是否退现, 0), 0, 2, 1))) as 排序1, " & vbCrLf & _
    "       max( Decode(Nvl(a.卡类别id, 0), 0, 0, Decode(Nvl(c.是否全退, 0), 0, 1, 2))) as 排序2, " & vbCrLf & _
    "       a.No, max(a.票据号) as 票据号,max(a.Id) as ID, sum(a.金额) as 金额, max(a.记录状态) as 记录状态, max(a.预交id) as 预交id, max(a.日期) as 日期, " & vbCrLf & _
    "       max(a.结算方式) as 结算方式, max(a.卡类别id) as 卡类别id,max( a.结算卡序号) as 结算卡序号, max(a.卡号) as 卡号, max(a.交易流水号) as 交易流水号, max(a.交易说明) as 交易说明, max(Nvl(c.是否转帐及代扣,0)) As 代扣, " & vbCrLf & _
    "       max(b.性质) As 结算性质, max(Nvl(c.是否退现,0)) As 退现, max(Nvl(c.是否全退,0)) As 全退, max(c.名称) As 三方卡名称,  max(c.是否缺省退现) As 缺省退现 " & vbCrLf & _
    " From ( " & strSql & "  ) A, 结算方式 B, 医疗卡类别 C " & vbCrLf & _
    " Where a.结算方式 = b.名称(+) And a.卡类别id = c.Id(+) And b.性质 <> 5 " & vbCrLf & _
    " group by A.NO " & vbCrLf & _
    " having sum(金额) <>0 " & vbCrLf & _
    " Order By 排序1 desc ,排序2 desc, 卡类别id Desc, No, 结算方式 " & vbCrLf & _
    " "
    
    '主要是适用支付宝更改,代扣标志为1,其他的在10.35版本中支持(先使用支付宝缴的预交部分)
    Set GetThreeDeposit = zlDatabase.OpenSQLRecord(strSql, "mdlInExse", lng病人ID, strTime, CDate(strDate), int预交类别)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
