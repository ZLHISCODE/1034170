VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBalanceDeposit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������Ԥ���˿�"
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
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
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
      ToolTipText     =   "�ȼ���F2"
      Top             =   4305
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
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
      ToolTipText     =   "�ȼ�:Esc"
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
   Begin VB.Label lbl����� 
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����Ǳ��ν��ʲ��˵�������Ԥ�����,  �������Ҫ�����˿�    "
      Height          =   360
      Left            =   885
      TabIndex        =   5
      Top             =   135
      Width           =   3420
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "����"
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
Private mlng����ID As Long, mlng����ID As Long
Private mlngModul As Long, mblnAll As Boolean
Private mblnDateMoved As Boolean
Private mstrסԺ���� As String
Private mstrDepositDate    As String
Private mintԤ�����    As Integer
Private mstrCardPrivs As String, mstrForceNote As String
Private mstrǿ�����ֲ���Ա As String

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
    Dim cllSquareBalance As Collection, strIDs As String, strNos As String, dbl��� As Double
    Dim rsTmp As ADODB.Recordset, lngRow As Long, j As Integer, strValue As String
    Dim strInXML As String, strOutXML As String, strExpend As String, strBalanceIDs As String
    If lbl�����.Visible Then
        dbl��� = Val(lbl�����.Caption)
    End If
    For i = 1 To vsfDeposit.Rows - 1
        Set cllSQL = New Collection
        Set cllSquareBalance = New Collection
        Set cllThreeSwap = New Collection
        With vsfDeposit
            If Val(.TextMatrix(i, .ColIndex("�˿���"))) <> 0 Then
                If Val(.TextMatrix(i, .ColIndex("����"))) = 0 Then
                    If .TextMatrix(i, .ColIndex("ת��")) = 1 Then
                        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
                                Val(.RowData(i)), False, _
                            mrsInfo!����, mrsInfo!�Ա�, mrsInfo!����, Val(.TextMatrix(i, .ColIndex("�˿���"))), strCardNo, strPassWord, _
                            False, True, False, False, cllSquareBalance) = False Then
                            strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                        Else
                            zlXML.ClearXmlText
                            zlXML.AppendNode "IN"
                            zlXML.appendData "CZLX", "2"
                            zlXML.AppendNode "IN", True
                            strXMLExpend = zlXML.XmlText
                            zlXML.ClearXmlText
                            If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, Val(.RowData(i)), _
                                strCardNo, Val(.TextMatrix(i, .ColIndex("�˿���"))), "", strXMLExpend) = False Then
                                strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                            Else
                                mrsDeposit.Filter = "����������='" & .TextMatrix(i, .ColIndex("���㿨����")) & "'"
                                dblMoney = Val(.TextMatrix(i, .ColIndex("�˿���")))
                                Do While Not mrsDeposit.EOF
                                    If dblMoney > 0 Then
                                        If dblMoney > Val(mrsDeposit!���) Then
                                            strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    Val(mrsDeposit!���) & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
                                            dblMoney = dblMoney - Val(mrsDeposit!���)
                                        Else
                                            strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    dblMoney & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
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
                                    strCardNo, Val(.TextMatrix(i, .ColIndex("�˿���"))), "", strXMLExpend) = False Then
                                    gcnOracle.RollbackTrans
                                    strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                                Else
                                    zlXML.ClearXmlText
                                    zlXML.AppendNode "IN"
                                        zlXML.appendData "CZLX", "2"
                                    zlXML.AppendNode "IN", True
                                    strXMLExpend = zlXML.XmlText
                                    zlXML.ClearXmlText
                                    If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, Val(.RowData(i)), strCardNo, _
                                        mlng����ID, Val(.TextMatrix(i, .ColIndex("�˿���"))), strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
                                        gcnOracle.RollbackTrans
                                        strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                                    Else
                                        Set cllUpdate = New Collection
                                        Set cllThreeSwap = New Collection
    '                                    Call zlAddUpdateSwapSQL(False, mlng����ID, Val(.RowData(i)), False, strCardNo, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
                                        Call zlAddThreeSwapSQLToCollection(False, mlng����ID, Val(.RowData(i)), False, strCardNo, strSwapExtendInfor, cllThreeSwap)
                                        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
                                        zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, False, True
                                    End If
                                End If
                            End If
                        End If
                    Else
                        '����˿�
                        strBalanceIDs = ""
                        zlXML.ClearXmlText
                        mrsDeposit.Filter = "����������='" & .TextMatrix(i, .ColIndex("���㿨����")) & "'"
                        dblMoney = Val(.TextMatrix(i, .ColIndex("�˿���")))
                        Call zlXML.AppendNode("JSLIST")
                        Do While Not mrsDeposit.EOF
                            If dblMoney > 0 Then
                                If dblMoney > Val(mrsDeposit!���) Then
                                    strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    Val(mrsDeposit!���) & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
                                    zlAddArray cllSQL, strSql
                                    strSql = "Select ID,����,������ˮ��,����˵�� From ����Ԥ����¼ Where ID = [1]"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsDeposit!Ԥ��ID))
                                    If Not rsTmp.EOF Then
                                        Call zlXML.AppendNode("JS")
                                            Call zlXML.appendData("KH", Nvl(rsTmp!����))
                                            Call zlXML.appendData("JYLSH", Nvl(rsTmp!������ˮ��))
                                            Call zlXML.appendData("JYSM", Nvl(rsTmp!����˵��))
                                            Call zlXML.appendData("ZFJE", Val(mrsDeposit!���))
                                            Call zlXML.appendData("JSLX", 1)
                                            Call zlXML.appendData("ID", Nvl(rsTmp!ID))
                                        Call zlXML.AppendNode("JS", True)
                                        strSql = "Zl_�����˿���Ϣ_Insert("
                                        strSql = strSql & mlng����ID & ","
                                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                                        strSql = strSql & Val(mrsDeposit!���) & ",'"
                                        strSql = strSql & Nvl(rsTmp!����) & "','"
                                        strSql = strSql & Nvl(rsTmp!������ˮ��) & "','"
                                        strSql = strSql & Nvl(rsTmp!����˵��) & "')"
                                        zlAddArray cllThreeSwap, strSql
                                        strBalanceIDs = strBalanceIDs & "," & Val(Nvl(rsTmp!ID))
                                    End If
                                    dblMoney = dblMoney - Val(mrsDeposit!���)
                                Else
                                    strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    dblMoney & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
                                    zlAddArray cllSQL, strSql
                                    strSql = "Select ID,����,������ˮ��,����˵�� From ����Ԥ����¼ Where ID = [1]"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsDeposit!Ԥ��ID))
                                    If Not rsTmp.EOF Then
                                        Call zlXML.AppendNode("JS")
                                            Call zlXML.appendData("KH", Nvl(rsTmp!����))
                                            Call zlXML.appendData("JYLSH", Nvl(rsTmp!������ˮ��))
                                            Call zlXML.appendData("JYSM", Nvl(rsTmp!����˵��))
                                            Call zlXML.appendData("ZFJE", dblMoney)
                                            Call zlXML.appendData("JSLX", 1)
                                            Call zlXML.appendData("ID", Nvl(rsTmp!ID))
                                        Call zlXML.AppendNode("JS", True)
                                        strSql = "Zl_�����˿���Ϣ_Insert("
                                        strSql = strSql & mlng����ID & ","
                                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                                        strSql = strSql & dblMoney & ",'"
                                        strSql = strSql & Nvl(rsTmp!����) & "','"
                                        strSql = strSql & Nvl(rsTmp!������ˮ��) & "','"
                                        strSql = strSql & Nvl(rsTmp!����˵��) & "')"
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
                            strBalanceIDs, Val(.TextMatrix(i, .ColIndex("�˿���"))), strSwapGlideNO, strSwapMemo, strXMLExpend) = False Then
                            strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                        Else
                            zlExecuteProcedureArrAy cllSQL, Me.Caption, True
                            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, True, True
                            If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, Val(.RowData(i)), False, strInXML, _
                                 mlng����ID, strOutXML, strExpend) = False Then
                                gcnOracle.RollbackTrans:
                                strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                            Else
                                '�ύ
                                Set cllThreeSwap = New Collection
                                If zlXML_Init = True Then
                                    If strOutXML <> "" Then
                                        If zlXML_LoadXMLToDOMDocument(strOutXML, False) Then
                                            Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
                                            For j = 0 To lngRow - 1
                                                Call zlXML_GetNodeValue("ID", i, strValue)
                                                strSql = "Zl_�����˿���Ϣ_Insert("
                                                strSql = strSql & mlng����ID & ","
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
                                strSql = "Select ���� From ����Ԥ����¼ Where ����ID= [1] And �����ID= [2]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, Val(.RowData(i)))
                                If Not rsTmp.EOF Then
                                    strCardNo = Nvl(rsTmp!����)
                                End If
    '                            Call zlAddUpdateSwapSQL(False, mlng����ID, Val(.RowData(i)), False, strCardNo, "", "", cllUpdate, 0)
                                Call zlAddThreeSwapSQLToCollection(False, mlng����ID, Val(.RowData(i)), False, strCardNo, strSwapExtendInfor, cllThreeSwap)
    '                            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
                                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, True, True
                                gcnOracle.CommitTrans
                            End If
                        End If
                    End If
                Else
                    '����
                    mrsDeposit.Filter = "����������='" & .TextMatrix(i, .ColIndex("���㿨����")) & "'"
                    dblMoney = Val(.TextMatrix(i, .ColIndex("�˿���")))
                    
                    Do While Not mrsDeposit.EOF
                        If dblMoney > 0 Then
                            If dblMoney > Val(mrsDeposit!���) Then
                                strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                        "'" & mrsDeposit!NO & "'" & ",1," & _
                                        Val(mrsDeposit!���) & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "'" & IIf(lbl�����.Visible And dbl��� <> 0, "," & dbl��� & ",'", ",Null,'") & mstrForceNote & "')"
                                dblMoney = dblMoney - Val(mrsDeposit!���)
                                If dbl��� <> 0 And lbl���.Visible Then
                                    dbl��� = 0
                                End If
                            Else
                                strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                        "'" & mrsDeposit!NO & "'" & ",1," & _
                                        dblMoney & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "'" & IIf(lbl�����.Visible And dbl��� <> 0, "," & dbl��� & ",'", ",Null,'") & mstrForceNote & "')"
                                dblMoney = 0
                                If dbl��� <> 0 And lbl���.Visible Then
                                    dbl��� = 0
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
        MsgBox "������������Ԥ�������˿�����г��ִ���,��ʹ������˿�ܶԸ���Ԥ��������˿�!" & vbCrLf & Mid(strFailNo, 2)
    End If
    mblnUnload = True
    Unload Me
End Sub

Public Sub ShowMe(frmMain As Object, lngModule As Long, lng����ID As Long, lng����ID As Long, blnAll As Boolean, _
                  Optional ByVal blnDateMoved As Boolean = False, Optional ByVal strסԺ���� As String = "", Optional ByVal strDepositDate As String = "", Optional ByVal intԤ����� As Integer)
    mlngModul = lngModule
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mblnAll = blnAll
    mblnDateMoved = blnDateMoved
    mstrסԺ���� = strסԺ����
    mstrDepositDate = strDepositDate
    mintԤ����� = intԤ�����
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
    Set mrsDeposit = GetThreeDeposit(mlng����ID, mblnDateMoved, mstrסԺ����, mstrDepositDate, mintԤ�����)
    With vsfDeposit
        Do While Not mrsDeposit.EOF
            lngRow = 0
            For i = 1 To .Rows - 1
                If .RowData(i) = Nvl(mrsDeposit!�����ID) Then
                    lngRow = i
                    Exit For
                End If
            Next i
            If lngRow = 0 Then lngRow = .Rows - 1: .Rows = .Rows + 1
            
            .TextMatrix(lngRow, .ColIndex("���㿨����")) = Nvl(mrsDeposit!����������)
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Nvl(mrsDeposit!���㷽ʽ)
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���"))) + Val(Nvl(mrsDeposit!���)), "0.00")
            
            If mblnAll Then
                .TextMatrix(lngRow, .ColIndex("�˿���")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�˿���"))) + Val(Nvl(mrsDeposit!���)), "0.00")
            Else
                .TextMatrix(lngRow, .ColIndex("�˿���")) = Format(0, "0.00")
            End If
            
            .TextMatrix(lngRow, 4) = 0
            If Val(mrsDeposit!����) = 1 Then
                '��������,�����޸�
                .Cell(flexcpData, lngRow, .ColIndex("����")) = 1
                .Cell(flexcpBackColor, lngRow, .ColIndex("����")) = vbWhite
            Else
                .Cell(flexcpData, lngRow, .ColIndex("����")) = 0
                .Cell(flexcpBackColor, lngRow, .ColIndex("����")) = &H8000000F
            End If
            
            .TextMatrix(lngRow, .ColIndex("Ԥ��ID")) = Nvl(mrsDeposit!Ԥ��ID)
            .TextMatrix(lngRow, .ColIndex("ת��")) = Nvl(mrsDeposit!����)
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(mrsDeposit!ID)
            .TextMatrix(lngRow, .ColIndex("��¼״̬")) = Nvl(mrsDeposit!��¼״̬)
            
            .RowData(lngRow) = Nvl(mrsDeposit!�����ID)
            mrsDeposit.MoveNext
        Loop
    End With
    
    If mrsDeposit.RecordCount = 0 Then
        mblnUnload = True
        Unload Me: Exit Sub
    End If
    
    vsfDeposit.Rows = vsfDeposit.Rows - 1
    strSql = "Select ����,����,�Ա� From ������Ϣ Where ����ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnUnload = False Then
        If MsgBox("�Ƿ�ȷ��ȡ��Ԥ�����˿�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    mstrForceNote = ""
    mstrǿ�����ֲ���Ա = ""
    mblnUnload = False
End Sub

Private Sub vsfDeposit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    With vsfDeposit
        If Col = .ColIndex("�˿���") Then
            If IsNumeric(.TextMatrix(Row, .ColIndex("�˿���"))) = False Then
                MsgBox "��������ȷ���˿���!", vbInformation, gstrSysName
                 .TextMatrix(Row, 3) = "0.00"
            End If
            If Val(vsfDeposit.TextMatrix(Row, .ColIndex("�˿���"))) > Val(.TextMatrix(Row, .ColIndex("���"))) Then
                MsgBox "������˿������,����", vbInformation, gstrSysName
                 .TextMatrix(Row, .ColIndex("�˿���")) = "0.00"
            End If
             .TextMatrix(Row, .ColIndex("�˿���")) = Format(Val(.TextMatrix(Row, .ColIndex("�˿���"))), "0.00")
        End If
        Call RecalCash
        
        If Col = .ColIndex("����") And Val(.Cell(flexcpData, Row, Col)) = 0 Then
            mstrForceNote = ""
            For i = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(i, .ColIndex("����")))) = 1 Then
                    mstrForceNote = mstrForceNote & IIf(mstrForceNote = "", mstrǿ�����ֲ���Ա & "ǿ������:", ";") & .TextMatrix(i, 0) & "," & Format(.TextMatrix(i, 3), "0.00") & "Ԫ"
                End If
            Next i
        End If
    End With
End Sub

Private Sub vsfDeposit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblMoney As Double, lngRow As Long
    Dim str����Ա���� As String, strDBUser As String
    Dim strPrivs As String
    With vsfDeposit
        If Col <> .ColIndex("����") Then Exit Sub
        
        If Val(.Cell(flexcpData, Row, Col)) = 0 Then
            If InStr(";" & mstrCardPrivs & ";", ";�����˿�ǿ������;") = 0 Then
                If mstrǿ�����ֲ���Ա = "" Then
                    mstrǿ�����ֲ���Ա = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
                    If mstrǿ�����ֲ���Ա = "" Then
                        MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                If mstrǿ�����ֲ���Ա = "" Then
                    If MsgBox("ѡ��Ľ��㿨��֧������,�Ƿ�ǿ�ƽ������֣�", _
                                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Cancel = True: Exit Sub
                    mstrǿ�����ֲ���Ա = UserInfo.����
                End If
            End If
        End If
    End With
End Sub

Private Sub RecalCash()
    '�����ֽ���
    Dim i As Integer, dblSum As Double
    Dim dbl��� As Double, dblʵ�� As Double
    Dim cur��� As Currency
    dblSum = 0
    With vsfDeposit
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, .ColIndex("����")))) = 1 Then
                dblSum = dblSum + Val(.TextMatrix(i, .ColIndex("�˿���")))
            End If
        Next i
    End With
    
    If dblSum = 0 Then
        txtMoney.Visible = False
        lblMoney.Visible = False
        lbl���.Visible = False
        lbl�����.Visible = False
    Else
        txtMoney.Visible = True
        lblMoney.Visible = True
        dblʵ�� = CentMoney(dblSum)
        cur��� = Val(dblSum) - Val(dblʵ��)
        txtMoney.Text = Format(dblʵ��, "0.00")
        lbl���.Visible = cur��� <> 0
        lbl�����.Visible = cur��� <> 0
        lbl�����.Caption = Format(cur���, "0.######")
    End If
End Sub

Private Sub vsfDeposit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDeposit
        Select Case Col
        Case .ColIndex("�˿���")
        Case .ColIndex("����")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Function GetThreeDeposit(lng����ID As Long, _
    Optional blnDateMoved As Boolean, Optional strTime As String, _
    Optional ByVal strPepositDate As String = "", _
    Optional intԤ����� As Integer = 0) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ʣ��Ԥ������ϸ(������)
    '���:strTime-סԺ����,��:1,2,3
    '        bln����תסԺ-�Ƿ��������תסԺ(ֻ�ܳ�ָ����Ԥ��)
    '        strPepositDate-ָ����Ԥ������
    '       intԤ�����-0-�����סԺ;1-����;2- סԺ
    '����:
    '����:Ԥ����ϸ����
    '����:������
    '����:2016-2-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strSub1 As String
    Dim strWherePage As String, strTable As String
    Dim strWhere As String, strDate As String
    On Error GoTo errH
    
    If intԤ����� = 1 Then strTime = ""    '69500
    
    strWherePage = IIf(strTime = "", "", " And instr(','||[2]||',',','||Nvl(A.��ҳID,0)||',')>0")
    strTable = IIf(blnDateMoved, zlGetFullFieldsTable("����Ԥ����¼"), "����Ԥ����¼ A")
    strWhere = "": strDate = "1974-04-28 00:00:00"
    If strPepositDate <> "" Then
        If IsDate(strPepositDate) Then
            strDate = strPepositDate
            strWhere = " And A.�տ�ʱ��=[3]"
        End If
    End If
    
    If intԤ����� <> 0 Then
        strWhere = strWhere & " And A.Ԥ����� =[4]"
    End If
    
    strSql = "" & _
    "    Select a.No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, " & vbCrLf & _
    "           Min(Decode(a.����id, Null, a.Id, 0) * Decode(a.��¼״̬, 1, 1, 0))*1 As ID, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0*Null, a.Id), 0)) As Ԥ��id, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, '', a.ʵ��Ʊ��), '')) As Ʊ�ݺ�, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, '', To_Char(a.�տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss')), '')) As ����, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, '', a.���㷽ʽ), '')) As ���㷽ʽ, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0*Null, a.�����id), 0)) As �����id, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0*Null, a.���㿨���), 0)) As ���㿨���, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, '', a.����), '')) As ����, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, '', a.������ˮ��), '')) As ������ˮ��, " & vbCrLf & _
    "           Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, '', a.����˵��), '')) As ����˵�� " & vbCrLf & _
    "     From  ����Ԥ����¼ A" & vbCrLf & _
    "     Where a.��¼���� In (1, 11) And a.����id = [1]   AND �����ID IS NOT NULL " & strWherePage & strWhere & vbCrLf & _
    "           And ( nvl(���,0)>=0 or a.��¼״̬<>1 or Not exists(select 1 from ����Ԥ����¼ where ����ID=[1] and a.�����id=�����ID and a.������ˮ��=������ˮ�� and ���>=0 and ��¼����=1)) " & vbCrLf & _
    "     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0 " & vbCrLf & _
    "     Group By a.No "
    
    strSql = strSql & "" & vbCrLf & _
    "     UNION ALL  " & vbCrLf & _
    "    Select b.No, Sum(Nvl(a.���, 0) - Nvl(a.��Ԥ��, 0)) As ���, 1 As ��¼״̬, 0 As ID, " & vbCrLf & _
    "             0 As Ԥ��id, Min(b.ʵ��Ʊ��) As Ʊ�ݺ�,To_Char(Min(b.�տ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As ����, Min(b.���㷽ʽ) As ���㷽ʽ, " & vbCrLf & _
    "             Min(b.�����id) As �����id, 0 * Null As ���㿨���, Min(b.����) As ����, Min(b.������ˮ��) As ������ˮ��, Min(b.����˵��) As ����˵��" & vbCrLf & _
    "    From ����Ԥ����¼ A, " & vbCrLf & _
    "          ( Select b.No, b.������ˮ��, b.Id, b.�տ�ʱ��, b.����, b.����˵��, b.�����id, b.����id, b.���㷽ʽ, b.ʵ��Ʊ�� " & vbCrLf & _
    "            From ����Ԥ����¼ B ��" & vbCrLf & _
    "            Where ����id = [1] And ��¼���� = 1 And ��¼״̬ = 1 And �����id Is Not Null and ���>0 " & _
    "           ) B,ҽ�ƿ���� C " & vbCrLf & _
    "     Where a.��¼���� In (1, 11) And a.����id = [1] And a.�����id Is Not Null And a.��� < 0 And a.�����id = b.�����id And  a.������ˮ�� = b.������ˮ�� " & strWherePage & strWhere & vbCrLf & _
    "            and a.�����ID=c.ID and nvl(c.�Ƿ�ת�ʼ�����,0)=0  and a.NO not in (select NO From ����Ԥ����¼ where ����ID=[1] and mod(��¼����,10)=1 and Ԥ�����=1 having sum(nvl(���,0))-sum(nvl(��Ԥ��,0))=0 Group by NO) " & vbCrLf & _
    "     Group By b.No " & vbCrLf & _
    "     Having Sum(Nvl(a.���, 0) - Nvl(a.��Ԥ��, 0)) <> 0"
    
    strSql = "" & _
    " Select " & vbCrLf & _
    "       max( Decode(Nvl(a.�����id, 0), 0, 0, Decode(Nvl(c.�Ƿ�����, 0), 0, 2, 1))) as ����1, " & vbCrLf & _
    "       max( Decode(Nvl(a.�����id, 0), 0, 0, Decode(Nvl(c.�Ƿ�ȫ��, 0), 0, 1, 2))) as ����2, " & vbCrLf & _
    "       a.No, max(a.Ʊ�ݺ�) as Ʊ�ݺ�,max(a.Id) as ID, sum(a.���) as ���, max(a.��¼״̬) as ��¼״̬, max(a.Ԥ��id) as Ԥ��id, max(a.����) as ����, " & vbCrLf & _
    "       max(a.���㷽ʽ) as ���㷽ʽ, max(a.�����id) as �����id,max( a.���㿨���) as ���㿨���, max(a.����) as ����, max(a.������ˮ��) as ������ˮ��, max(a.����˵��) as ����˵��, max(Nvl(c.�Ƿ�ת�ʼ�����,0)) As ����, " & vbCrLf & _
    "       max(b.����) As ��������, max(Nvl(c.�Ƿ�����,0)) As ����, max(Nvl(c.�Ƿ�ȫ��,0)) As ȫ��, max(c.����) As ����������,  max(c.�Ƿ�ȱʡ����) As ȱʡ���� " & vbCrLf & _
    " From ( " & strSql & "  ) A, ���㷽ʽ B, ҽ�ƿ���� C " & vbCrLf & _
    " Where a.���㷽ʽ = b.����(+) And a.�����id = c.Id(+) And b.���� <> 5 " & vbCrLf & _
    " group by A.NO " & vbCrLf & _
    " having sum(���) <>0 " & vbCrLf & _
    " Order By ����1 desc ,����2 desc, �����id Desc, No, ���㷽ʽ " & vbCrLf & _
    " "
    
    '��Ҫ������֧��������,���۱�־Ϊ1,��������10.35�汾��֧��(��ʹ��֧�����ɵ�Ԥ������)
    Set GetThreeDeposit = zlDatabase.OpenSQLRecord(strSql, "mdlInExse", lng����ID, strTime, CDate(strDate), intԤ�����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
