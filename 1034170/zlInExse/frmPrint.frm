VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Ʊ�ݴ�ӡ"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
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

Private mbytInFun As Byte               '1-�µ���ӡ,2-�ش�
Private mlng����ID As Long              '�ϴ�����ID
Private mstrPrintNO As String           '���ʵ��ݺ�
Private mlngBalanceID As Long           '����ID
Private mstrInvoice As String           '��ʼƱ�ݺ�
Private mdateBalance As Date            '���ʻ��ش��ʱ��
Private mblnPrinted As Boolean          '��ӡƱ�����������Ƿ�ɹ�
Private mlngShareUseID As Long '��������
Private mstrUseType As String
Private mbytInvoiceKind As Byte     '1-סԺ,2-����

Private Sub Form_Load()
    Set mobjReport = New clsReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    mbytInFun = 0
    mlng����ID = 0
    mstrPrintNO = ""
    mlngBalanceID = 0
    mstrInvoice = ""
    mbytInvoiceKind = 0
    mdateBalance = CDate(0)
    mblnPrinted = False
End Sub


Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrInvoice As Variant)
    Dim strSql As String, i As Integer, strInvoices As String
    
    'û��Ʊ�ݺ�,�ϸ����Ʊ��ʱ����ӡ,���ϸ����Ʊ��ʱֻ��ӡ������Ʊ������
    If mstrInvoice = "" Then
        Cancel = gblnStrictCtrl
        mblnPrinted = Not gblnStrictCtrl
        Exit Sub
    End If
    
    mblnPrinted = False
    '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
    If gblnStrictCtrl Then
        mlng����ID = GetInvoiceGroupID(IIf(mbytInvoiceKind = 0, IIf(gbytInvoiceKind = 0, 3, 1), IIf(mbytInvoiceKind = 1, 3, 1)), TotalPages, mlng����ID, mlngShareUseID, mstrInvoice, mstrUseType)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case -1
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ�,�ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ�������������ش򵥾�[" & mstrPrintNO & "]", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    
    '2.����Ʊ��ʹ������
    On Error GoTo errH
    Select Case mbytInFun
        Case 1
            strSql = "zl_���˽���Ʊ��_Insert('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & _
                ",'" & UserInfo.���� & "',To_Date('" & Format(mdateBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & TotalPages & "," & gbytInvoiceKind & ")"
        
        Case 2
            strSql = "zl_���˽��ʼ�¼_RePrint('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & _
                ",'" & UserInfo.���� & "'," & TotalPages & "," & gbytInvoiceKind & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSql, "Ʊ����������")
    mblnPrinted = True
    
    '3.�������õ�Ʊ�ݺ���Ϣ
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
        
    '���ϸ����Ʊ��ʱ���浽ע���
    If Not gblnStrictCtrl Then
        zlDatabase.SetPara "��ǰ����Ʊ�ݺ�", mstrInvoice, glngSys, 1137
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub
Public Sub ReportPrint(ByVal bytInfun As Byte, ByVal strNo As String, ByVal lngBalanceID As Long, _
                        ByRef lngLastUseID As Long, ByVal lngShareUseID As Long, ByVal strUseType As String, ByVal strInvoice As String, Optional ByVal dateBalance As Date, _
                        Optional str�ɿ� As String, Optional str�Ҳ� As String, Optional lngPatientID As Long, _
                        Optional intLocalFormat As Integer, Optional blnPrintBillEmpty As Boolean = False, Optional bytInvoiceKind As Byte = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݴ�ӡ
    '���:bytInfun:1-�µ���ӡ,2-�ش�
    '       strNO:���ʵ��ݺ�,��������
    '       lngBalanceID:����ID
    '       lngLastUseID:���ʹ�õ���������ID,����ʱΪ0
    '       lngShareUseID:��������
    '       strUseType:ʹ�����
    '       strInvoice:��ʼƱ�ݺţ���������,���ϸ����Ʊ��ʱ�������,�ϸ����ʱ����ǰ��ǰ��鲻��Ϊ��
    '       dateBalance :����ʱ��,���µ���ӡ�Ŵ���
    '       lngPatientID:��Լ��λ���ʰ����˷ֱ��ӡ,ÿ�δ�ӡ���뵱ǰ����ID
    '       intLocalFormat:��ָ���ĸ�ʽ��ӡ
    '����:
    '       blnPrintBillEmpty-�Ƿ��ӡ��Ʊ��(55052)
    '����:
    '����:���˺�
    '����:2011-05-03 17:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strReportNO As String, strSql As String, strFormat As String
    Dim arrInvoice As Variant
    blnPrintBillEmpty = False
    '1.��������
    mbytInFun = bytInfun: mstrPrintNO = strNo
    mlngBalanceID = lngBalanceID: mlng����ID = lngLastUseID
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
    'ѡ��Ĵ�ӡ��ʽ
    strFormat = IIf(intLocalFormat <= 0, "", "ReportFormat=" & intLocalFormat)
    mblnPrinted = False
    '2.��ӡ����
    Select Case mbytInFun
        Case 1  '�µ���ӡ
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)   '���ô�ӡ����������ӡ��ֻ������Ʊ��ʹ������
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
                
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "����ID=" & mlngBalanceID, "����ID=" & lngPatientID, "PrintEmpty=0", str�ɿ�, str�Ҳ�, strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then GoTo ClearInvoice
            End If
        Case 2  '�ش�
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then Exit Sub
                
                If Not gobjTax Is Nothing And gblnTax Then
                    MsgBox "����׼����֮��ȷ����ʼ��ӡ��", vbInformation, gstrSysName
                    gstrTax = gobjTax.zlTaxInReput(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                If gblnBillPrint Then
                    If gobjBillPrint.zlRePrintBill("", mlngBalanceID, strInvoice) = False Then Exit Sub
                End If
                
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "����ID=" & mlngBalanceID, "����ID=" & lngPatientID, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
    End Select
    '3.�������ʹ�õ�����ID
    lngLastUseID = mlng����ID
    Exit Sub
    
ClearInvoice:
    On Error GoTo errH
    strSql = "Zl_Ʊ����ʼ��_Update('" & strNo & "','',3)"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


