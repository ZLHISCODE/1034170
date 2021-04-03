VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "票据打印"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1

Private mbytInFun As Byte               '1-新单打印,2-重打
Private mlng领用ID As Long              '上次领用ID
Private mstrPrintNO As String           '结帐单据号
Private mlngBalanceID As Long           '结帐ID
Private mstrInvoice As String           '开始票据号
Private mdateBalance As Date            '结帐或重打的时间
Private mblnPrinted As Boolean          '打印票据数据生成是否成功
Private mlngShareUseID As Long '共享批次
Private mstrUseType As String
Private mbytInvoiceKind As Byte     '1-住院,2-门诊

Private Sub Form_Load()
    Set mobjReport = New clsReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    mbytInFun = 0
    mlng领用ID = 0
    mstrPrintNO = ""
    mlngBalanceID = 0
    mstrInvoice = ""
    mbytInvoiceKind = 0
    mdateBalance = CDate(0)
    mblnPrinted = False
End Sub


Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrInvoice As Variant)
    Dim strSql As String, i As Integer, strInvoices As String
    
    '没有票据号,严格控制票据时不打印,不严格控制票据时只打印不处理票据数据
    If mstrInvoice = "" Then
        Cancel = gblnStrictCtrl
        mblnPrinted = Not gblnStrictCtrl
        Exit Sub
    End If
    
    mblnPrinted = False
    '1.严格控制票据时，根据实际的票据张数,重新检查领用ID和票据号
    If gblnStrictCtrl Then
        mlng领用ID = GetInvoiceGroupID(IIf(mbytInvoiceKind = 0, IIf(gbytInvoiceKind = 0, 3, 1), IIf(mbytInvoiceKind = 1, 3, 1)), TotalPages, mlng领用ID, mlngShareUseID, mstrInvoice, mstrUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case -1
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "你没有足够的自用和共用的票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "你没有足够的的共用票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                        "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
                Case -4
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]所在的领用批次没有足够的票据！" & _
                        "请先打印其它票据,用完当前领用批次后,重打该单据！", vbInformation, gstrSysName
                Case Else
                    MsgBox "票据领用信息访问失败！将来，你可以重打单据[" & mstrPrintNO & "]", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    
    '2.产生票据使用数据
    On Error GoTo errH
    Select Case mbytInFun
        Case 1
            strSql = "zl_病人结帐票据_Insert('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & _
                ",'" & UserInfo.姓名 & "',To_Date('" & Format(mdateBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & TotalPages & "," & gbytInvoiceKind & ")"
        
        Case 2
            strSql = "zl_病人结帐记录_RePrint('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & _
                ",'" & UserInfo.姓名 & "'," & TotalPages & "," & gbytInvoiceKind & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSql, "票据数据生成")
    mblnPrinted = True
    
    '3.传递所用的票据号信息
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
        
    '不严格控制票据时保存到注册表
    If Not gblnStrictCtrl Then
        zlDatabase.SetPara "当前结帐票据号", mstrInvoice, glngSys, 1137
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub
Public Sub ReportPrint(ByVal bytInfun As Byte, ByVal strNo As String, ByVal lngBalanceID As Long, _
                        ByRef lngLastUseID As Long, ByVal lngShareUseID As Long, ByVal strUseType As String, ByVal strInvoice As String, Optional ByVal dateBalance As Date, _
                        Optional str缴款 As String, Optional str找补 As String, Optional lngPatientID As Long, _
                        Optional intLocalFormat As Integer, Optional blnPrintBillEmpty As Boolean = False, Optional bytInvoiceKind As Byte = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐票据打印
    '入参:bytInfun:1-新单打印,2-重打
    '       strNO:结帐单据号,不带引号
    '       lngBalanceID:结帐ID
    '       lngLastUseID:最近使用的领用批次ID,初次时为0
    '       lngShareUseID:共享批次
    '       strUseType:使用类别
    '       strInvoice:开始票据号，不带引号,不严格控制票据时允许传入空,严格控制时调用前已前检查不能为空
    '       dateBalance :结帐时间,仅新单打印才传入
    '       lngPatientID:合约单位结帐按病人分别打印,每次打印传入当前病人ID
    '       intLocalFormat:按指定的格式打印
    '出参:
    '       blnPrintBillEmpty-是否打印空票据(55052)
    '返回:
    '编制:刘兴洪
    '日期:2011-05-03 17:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strReportNO As String, strSql As String, strFormat As String
    Dim arrInvoice As Variant
    blnPrintBillEmpty = False
    '1.变量传递
    mbytInFun = bytInfun: mstrPrintNO = strNo
    mlngBalanceID = lngBalanceID: mlng领用ID = lngLastUseID
    mstrInvoice = strInvoice: mdateBalance = dateBalance
    mlngShareUseID = lngShareUseID
    mstrUseType = strUseType
    mbytInvoiceKind = bytInvoiceKind
    If bytInvoiceKind = 0 Then
        If gbytInvoiceKind = 0 Then
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137"
        Else
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137_2"
        End If
    Else
        If bytInvoiceKind = 1 Then
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137"
        Else
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137_2"
        End If
    End If
    '选择的打印格式
    strFormat = IIf(intLocalFormat <= 0, "", "ReportFormat=" & intLocalFormat)
    mblnPrinted = False
    '2.打印调用
    Select Case mbytInFun
        Case 1  '新单打印
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)   '调用打印方法但不打印，只生成了票据使用数据
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then GoTo ClearInvoice
                
                If Not gobjTax Is Nothing And gblnTax Then
                    gstrTax = gobjTax.zlTaxInPrint(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                If gblnBillPrint Then
                    If gobjBillPrint.zlPrintBill("", mlngBalanceID) = False Then GoTo ClearInvoice
                End If
                
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结帐ID=" & mlngBalanceID, "病人ID=" & lngPatientID, "PrintEmpty=0", str缴款, str找补, strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then GoTo ClearInvoice
            End If
        Case 2  '重打
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then Exit Sub
                
                If Not gobjTax Is Nothing And gblnTax Then
                    MsgBox "请在准备好之后按确定开始打印。", vbInformation, gstrSysName
                    gstrTax = gobjTax.zlTaxInReput(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                If gblnBillPrint Then
                    If gobjBillPrint.zlRePrintBill("", mlngBalanceID, strInvoice) = False Then Exit Sub
                End If
                
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结帐ID=" & mlngBalanceID, "病人ID=" & lngPatientID, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
    End Select
    '3.传回最近使用的领用ID
    lngLastUseID = mlng领用ID
    Exit Sub
    
ClearInvoice:
    On Error GoTo errH
    strSql = "Zl_票据起始号_Update('" & strNo & "','',3)"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


