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
Private mbytInFun As Byte                 '1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ; 4-����Ʊ��;6-�˷�Ʊ��(��Ʊ)��ӡ
Private mlng����ID As Long              '�ϴ�����ID
Private mstrPrintNO As String           'Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
Private mstrInvoice As String           '��ʼƱ�ݺ�
Private mdatFeeDate As Date             '���õ������ݵĵǼ�ʱ��
Private mblnPrinted As Boolean          'Ʊ�����������Ƿ�ɹ�(�Ƿ��Ѵ�ӡ)
Private mstrReclaimInvoice As String    'Ҫ����յķ�Ʊ��,��1-����ϵͳԤ���������Ʊ�ź�2-�����û������������Ʊ����Ч
Private mstrPrivs As String
Private mlngShareUseID As Long '��ӡ�Ĺ�������ID
Private mstrUseType As String
Private mbln����Ʊ�� As Boolean
Private mblnOnePatiPrint As Boolean, mlng��ӡID As Long '�����˲���Ʊ��ʱʹ��

Private Sub Form_Load()
    mstrPrivs = ";" & GetPrivFunc(glngSys, 1121)
    Set mobjReport = New clsReport
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrInvoice As Variant)
    Dim strSQL As String, i As Integer, strInvoices As String
    
    If mbytInFun <> 6 Then
        '56963
        If gTy_Module_Para.bytƱ�ݷ������ <> 0 And mbln����Ʊ�� Then Exit Sub
        
        If gTy_Module_Para.blnһ��Ʊ�� Then TotalPages = 1  '�շ�ÿ�δ�ӡֻ��һ��Ʊ��
    End If
    'û��Ʊ�ݺ�,�ϸ����Ʊ��ʱ����ӡ,���ϸ����Ʊ��ʱֻ��ӡ������Ʊ������
    If mstrInvoice = "" Then
        Cancel = gblnStrictCtrl
        mblnPrinted = Not gblnStrictCtrl
        Exit Sub
    End If
    If CheckInvoiceValied(TotalPages, mbytInFun = 6) = False Then Cancel = True: Exit Sub
    
    On Error GoTo errH
    '2.����Ʊ������
    Select Case mbytInFun
        Case 1
         'Create Or Replace Procedure Zl_�����շ�Ʊ��_Insert
            strSQL = "Zl_�����շ�Ʊ��_Insert("
            '  No_In           Varchar2,
            strSQL = strSQL & "" & IIf(mblnOnePatiPrint, "NULL", "'" & Replace(mstrPrintNO, "'", "") & "'") & ","
            '  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
            strSQL = strSQL & "" & ZVal(mlng����ID) & ","
            '  ʹ����_In       Ʊ��ʹ����ϸ.ʹ����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ʹ��ʱ��_In     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
            strSQL = strSQL & "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  ��ӡid_In       Ʊ�ݴ�ӡ����.Id%Type := 0,
            strSQL = strSQL & "" & IIf(mblnOnePatiPrint, mlng��ӡID, 0) & ","
            '  Ʊ������_In     Number := 1,
            strSQL = strSQL & "" & TotalPages & ","
            '  ҽ���ӿڴ�ӡ_In Number := 0,
            strSQL = strSQL & "0,"
            '  �����˴�ӡ_In Number:=0
            strSQL = strSQL & "" & IIf(mblnOnePatiPrint, 1, 0) & ")"
        Case 2, 3
            '����Ƕ��ţ�ֻ��Ҫ��һ�ŵ��ݺž�����(�޸Ķ����е�һ��ʱ,���һ�����µ�)
            strSQL = "zl_�����շѼ�¼_RePrint('" & Replace(Split(mstrPrintNO, ",")(0), "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & ",'" & UserInfo.���� & "'," & _
                    "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(mbytInFun = 2, "0", "1") & "," & TotalPages & ")"
        Case 6 '�˷ѷ�Ʊ(��Ʊ)��ӡ
            'Zl_�����˷�Ʊ��_Insert
            strSQL = "Zl_�����˷�Ʊ��_Insert("
            '  �������_In   ����Ԥ����¼.�������%Type,
            strSQL = strSQL & "" & Val(mstrPrintNO) & ","
            '  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
            strSQL = strSQL & "" & ZVal(mlng����ID) & ","
            '  ʹ����_In       Ʊ��ʹ����ϸ.ʹ����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ʹ��ʱ��_In     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
            strSQL = strSQL & "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  Ʊ������_In Number:=1
            strSQL = strSQL & "" & TotalPages & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ����������")
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
        zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mstrInvoice, glngSys, 1121, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub
Private Function CheckInvoiceValied(Optional int���� As Integer = 1, _
    Optional ByVal blnDelFeePrint As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ�Ϸ�(�ϸ����Ʊ��ʱ)
    '���:int���� -��Ҫ�ķ�Ʊ����
    '   blnDelFeePrint-�˷ѷ�Ʊ(��Ʊ)��ӡ
    '����:
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2013-03-27 13:01:41
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gblnStrictCtrl Then CheckInvoiceValied = True: Exit Function
    '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
    mlng����ID = GetInvoiceGroupID(1, int����, mlng����ID, mlngShareUseID, mstrInvoice, mstrUseType)
    '���ݺϷ�
    If mlng����ID > 0 Then CheckInvoiceValied = True: Exit Function
    Select Case mlng����ID
        Case -1
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "��û���㹻�����ú͹��õ�Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
        Case -2
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "��û���㹻�ĵĹ���Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
        Case -3
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "Ʊ�ݺ�[" & mstrInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
        Case -4
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "Ʊ�ݺ�[" & mstrInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ��ش�õ��ݣ�", vbInformation, gstrSysName
        Case Else
            MsgBox "Ʊ��������Ϣ����ʧ�ܣ������������" & IIf(blnDelFeePrint, "�ش�õ��ݣ�", "�ش򵥾�[" & mstrPrintNO & "]��"), vbInformation, gstrSysName
    End Select
End Function
Private Sub TaxInterface(ByVal byt���� As Byte, ByVal strPrintNO As String, ByVal strModiNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˰�ش�ӡ�ӿ�
    '���:byt����-1-������ӡ(���޸�);2-�ش�;3-�˷�
    '        strPrintNO-Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
    '        strModiNos-�޸Ķ൥���е�һ��ʱ,ָ�ö��ŵ��ݵ�����NO���ö��ŷָ�:'F0000001','F0000002',...
    '����:���˺�
    '����:2013-03-27 14:24:03
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    'δ����˰��,ֱ�ӷ���
    If Not gblnTax Then Exit Sub
    If byt���� = 3 Then
        '�˷�
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If byt���� = 2 Then
        '�ش�
        MsgBox "����׼����֮��ȷ����ʼ��ӡ��", vbInformation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If strModiNos <> "" Then
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strModiNos)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    End If
    gstrTax = gobjTax.zlTaxOutPrint(gcnOracle, strPrintNO)
    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
Private Function BillPrint(ByVal byt���� As Byte, ByVal strPrintNO As String, _
    ByVal strModiNos As String, ByRef strInvoice As String, ByRef strClearNOs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݴ�ӡ�ӿ�
    '���:byt����-1-������ӡ(���޸Ĵ�ӡ);2-�ش��ӡ;3-�˷�
    '        strPrintNO-Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
    '        strModiNos-�޸Ķ൥���е�һ��ʱ,ָ�ö��ŵ��ݵ�����NO���ö��ŷָ�:'F0000001','F0000002',...
    '         strInvoice-��Ʊ��(�ش�ʱ��Ч)
    '����:strClearNOs-��Ҫ����ĵ��ݺ�
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-03-27 14:36:28
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gblnBillPrint Then BillPrint = True: Exit Function
    If byt���� = 3 Then
        '�˷�
        '�˷�����֮ǰ�ȵ���Ʊ���ջأ�zlEraseBill
        BillPrint = gobjBillPrint.zlRePrintBill(strPrintNO, 0, strInvoice)
        Exit Function
    End If
    If byt���� = 2 Then
        '�ش�
       BillPrint = gobjBillPrint.zlRePrintBill(strPrintNO, 0, strInvoice)
       Exit Function
    End If
    If strModiNos <> "" Then
        If gobjBillPrint.zlEraseBill(strModiNos, 0) = False Then strClearNOs = Replace(strModiNos, "'", ""): Exit Function
    End If
    If gobjBillPrint.zlPrintBill(strPrintNO, 0) = False Then strClearNOs = Replace(strPrintNO, "'", ""): Exit Function
    BillPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function InsureReprint(ByVal blnVirtualPrint As Boolean, ByVal strNos As String, _
    ByVal lng����ID As Long, ByVal bln�˷� As Boolean, ByRef strInvoice As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���µ���ҽ����ӡ�ӿ�
    '���:blnVirtualPrint-�Ƿ����ҽ���ӿڴ�ӡ
    '       strNos-���ݺ�
    '       bln�˷�-�Ƿ��˷�
    '       strInvoice-��Ʊ��
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-03-27 17:01:02
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer
    On Error GoTo errHandle
    If Not blnVirtualPrint Then InsureReprint = True: Exit Function
    intInsure = ChargeExistInsure(strNos, 0, lng����ID, , bln�˷�)
    Call gclsInsure.RePrintBill(intInsure, lng����ID, strInvoice)
    InsureReprint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ReportPrint(ByVal bytInFun As Byte, ByVal strNos As String, ByVal strAllNOs As String, ByVal strReclaimInvoice As String, _
                        ByRef lngLastUseID As Long, ByVal lngShareUseID As Long, ByVal strInvoice As String, _
                        ByVal datFeeDate As Date, _
                        Optional str�ɿ� As String, Optional str�Ҳ� As String, Optional bln�ֱ��ӡ As Boolean, _
                        Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                        Optional ByVal blnDelRecord As Boolean, Optional strUseType As String = "", _
                        Optional blnPrintBillEmpty As Boolean, _
                        Optional blnOnePatiPrint As Boolean, Optional lng��ӡID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�ݴ�ӡ,�������
    '���:bytInfun :1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ,4-����Ʊ��(ֻ��:2-��ϵͳԤ�������3-�û��Զ�����ʱ��ת��),6-�˷�Ʊ��(��Ʊ)��ӡ
    '       strNOs - �µ�ʱҪ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...,
    '                   - �޸�ʱ,�����µ��ݺ�,ֻ��һ��,���ڴ�ӡȡ���������ʼƱ�ݺ�
    '                   - �˷�Ʊ��(��Ʊ)��ӡʱ������������
    '       strAllNOs-�޸Ķ൥���е�һ��ʱ,ָ�ö��ŵ��ݵ�����NO���ö��ŷָ�:'F0000001','F0000002',...
    '       strReclaimInvoice-Ҫ����յķ�Ʊ��,����ö��ŷ���'F0000001','F0000002',...
    '       lngLastUseID-���ʹ�õ���������ID,����ʱΪ0
    '       strInvoice-��ʼƱ�ݺţ���������,���ϸ����Ʊ��ʱ�������,�ϸ����ʱ����ǰ��ǰ��鲻��Ϊ��
    '       datFeeDate-���õ������ݵĵǼ�ʱ��
    '       intPrintFormat-��ӡ��ʽ(��ӡ��ʽ���)
    '       blnVirtualPrint-ҽ���ӿ��ڵ��ô�ӡ��HISֻ��Ʊ�Ų�ʵ�ʴ�ӡ
    '       blnDelRecord-�ش�ʱ���Ƿ��Ƕ��˷Ѽ�¼�����ش�(Ŀǰֻ�б���ҽ��(ҽ���ӿڴ�ӡƱ��)������)
    '       lngShareUseID-��������
    '       strUseType-ʹ�����
    '       blnOnePatiPrint-�����˲���Ʊ��(���ֽ������)
    '       lng��ӡID-����Ĵ�ӡID(blnOnePatiPrint=trueʱ����),������Ը��ݴ�ӡID�ӡ���ʱƱ�ݴ�ӡ���ݡ�����ʱ��������ȡ��Ӧ���շѵ���
    '                 ֮����Ҫ��ʱ����Ҫԭ������Ϊ�����˴�ӡʱ�����ݺſ��ܻ���ɣ����������屨����������
    ' ����:
    '   blnPrintBillEmpty-�Ƿ��ӡ�Ŀձ�����()
    '����:���˺�
    '����:2011-04-29 12:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    Dim i As Integer, j As Integer, strPrintNO As String, blnPrint As Boolean, blnTrans As Boolean
    Dim strReportNO As String, strSQL As String, strClearNOs As String, strFormat As String, lngBalanceID As Long
    Dim blnNotPrint As Boolean, varTmp As Variant '�ձ�������Ҫ��Ϊ�˴�����ú����ķ���ֵ
    Dim str��Ʊ�� As String, intƱ������ As Integer
    
    blnPrintBillEmpty = False
    mbln����Ʊ�� = False
    '1.��������
    mlngShareUseID = lngShareUseID
    mbytInFun = bytInFun
    mlng����ID = lngLastUseID: mstrUseType = strUseType
    mstrInvoice = strInvoice
    mdatFeeDate = datFeeDate
    mstrReclaimInvoice = strReclaimInvoice
    mblnOnePatiPrint = blnOnePatiPrint: mlng��ӡID = lng��ӡID
    
    If mbytInFun = 6 Then '�˷�Ʊ��(��Ʊ)��ӡ
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1121_7"
    Else
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1121_1"
    End If
    strFormat = IIf(intPrintFormat = 0, "", "ReportFormat=" & intPrintFormat)
    
    mstrPrintNO = ""
    mblnPrinted = False
    blnNotPrint = (Not gobjTax Is Nothing And gblnTax) Or blnVirtualPrint
    '2.��ӡ����
    Select Case mbytInFun
        Case 1 '�µ���ӡ,�޸��ش���ش�Ʊ��
            If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
               '1.����ϵͳԤ�������ӡƱ��;2-�����û��Զ������Ʊ��
               '�ȷ���Ʊ��:
               If CheckInvoiceValied(1) = False Then
                    strClearNOs = Replace(strNos, "'", "")
                    GoTo ClearInvoice: Exit Sub
               End If

               If zlExeCuteBillNoSplit(False, 1, mlng����ID, strNos, 0, mstrInvoice, mdatFeeDate, 1, str��Ʊ��, intƱ������, mlng��ӡID) = False Then
                    strClearNOs = Replace(strNos, "'", "")
                    GoTo ClearInvoice:
                    Exit Sub
                End If
               If intƱ������ = 0 Then
                    strClearNOs = Replace(strNos, "'", "")
                    GoTo ClearInvoice:
                    Exit Sub     'û������Ʊ��,ֱ�ӷ���
                End If
               mstrReclaimInvoice = str��Ʊ��
               mbln����Ʊ�� = True
               If blnNotPrint Then
                  Call TaxInterface(1, mstrPrintNO, strAllNOs)      '��ӡ˰��Ʊ��
               Else
                   'Ʊ�ݽӿ�
                    If BillPrint(1, mstrPrintNO, strAllNOs, "", strClearNOs) = False Then Exit Sub
                    '���ñ���
                   Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                   "��Ʊ��=" & str��Ʊ��, "NO='NO'", "��ӡID=" & mlng��ӡID, "PrintEmpty=0", str�ɿ�, str�Ҳ�, strFormat, 2)
                   If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
               End If
               Exit Sub
            End If
            If bln�ֱ��ӡ And UBound(Split(strNos, ",")) > 0 And strAllNOs = "" Then
            '������޸ĵĶ����е�һ�ţ���ʹ���ڲ����Ƿֱ��ӡ����ȻҪһ�����Ϊԭʼ��һ���ģ���
                For i = 0 To UBound(Split(strNos, ","))
                    mblnPrinted = False '���������ʼ����Ϊ���ܴ�ӡ�����ڵ�BeforePrint֮ǰ�ͳ�������
                    mstrPrintNO = Split(strNos, ",")(i)
                    blnPrint = True
                    If gTy_Module_Para.bln������ Then
                        'һ�ŵ���ֻ�й����Ѳ���ӡ
                        If BillOnlyFactMoney(Replace(mstrPrintNO, "'", "")) Then blnPrint = False
                    End If
                    If blnPrint Then
                        If blnNotPrint Then
                            Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)    '���ô�ӡ����������ӡ��ֻ������Ʊ��ʹ������
                            If Not mblnPrinted Then Exit For
                            Call TaxInterface(1, mstrPrintNO, "")        '��ӡ˰��Ʊ��
                            'Ʊ�ݽӿ�
                             If BillPrint(1, mstrPrintNO, "", "", "") = False Then Exit For
                        Else
                            'Ʊ�ݽӿ�
                             If BillPrint(1, mstrPrintNO, "", "", "") = False Then Exit For
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNo", "NO=" & mstrPrintNO, "��ӡID=" & mlng��ӡID, "PrintEmpty=0", str�ɿ�, str�Ҳ�, strFormat, 2)
                            If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                            If Not mblnPrinted And mobjReport.DataIsEmpty = False Then Exit For '109708
                        End If
                        If mobjReport.DataIsEmpty Then   '109708
                            strClearNOs = strClearNOs & "," & mstrPrintNO
                        Else
                            If i < UBound(Split(strNos, ",")) Then 'ȡ��һƱ�ݺ�,��BeforePrint��ʹ��
                                If gblnStrictCtrl Then
                                    mstrInvoice = GetNextBill(mlng����ID)   '����ʱ,���ؿ�
                                Else
                                    mstrInvoice = IncStr(mstrInvoice)
                                End If
                            End If
                        End If
                    End If
                Next
                '��;ʧ�ܴ���
                If i < UBound(Split(strNos, ",")) + 1 Then 'ע�⣺mobjReport_BeforePrint�е���ʾ����֮ǰ
                    If i = 0 Then
                        MsgBox "����[" & strNos & "]һ��Ҳû�д�ӡ!" & vbCrLf & _
                            "����ָ���µ�Ʊ�ݺź�ʹ���ش��ܴ�ӡ��", vbInformation, gstrSysName
                    Else
                        MsgBox "����[" & strNos & "]ֻ��ӡ��ǰ" & i & "��!" & vbCrLf & _
                            "ʣ�µ�����ָ���µ�Ʊ�ݺź�ʹ���ش��ܴ�ӡ��", vbInformation, gstrSysName
                    End If
                    For j = i To UBound(Split(strNos, ","))
                        strClearNOs = strClearNOs & "," & Split(strNos, ",")(j)
                    Next
                    strClearNOs = Replace(Mid(strClearNOs, 2), "'", "")
                    GoTo ClearInvoice
                End If
            Else
                '�޸Ķ����е�һ��ʱ,�޸��²��������ŷ������,��ΪmobjReport_BeforePrint��Ҫȡ��һ�����ջ�ԭ����
                mstrPrintNO = IIf(strAllNOs <> "", strAllNOs & ",", "") & strNos
                If blnNotPrint Then
                    Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp) '���ô�ӡ����������ӡ��ֻ������Ʊ��ʹ������
                    If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice '�޸�ʱ,ֻ����µ��ݵĿ�ʼƱ�ݺ�
                    Call TaxInterface(1, mstrPrintNO, strAllNOs)
                Else
                   'Ʊ�ݽӿ�
                    If BillPrint(1, mstrPrintNO, strAllNOs, "", strClearNOs) = False Then: GoTo ClearInvoice
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "��ӡID=" & mlng��ӡID, "PrintEmpty=0", str�ɿ�, str�Ҳ�, strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
                End If
            End If
        Case 2, 4 '�ش�
            mstrPrintNO = strNos
            If gTy_Module_Para.bytƱ�ݷ������ <> 0 And (mstrReclaimInvoice <> "" Or mbytInFun = 4) Then
               '1.����ϵͳԤ�������ӡƱ��;2-�����û��Զ������Ʊ��
               '���շ�Ʊ����Ϊ��ʱ,���ܰ��·�ʽ�ش�Ʊ��
               '�ȷ���Ʊ��:
               If CheckInvoiceValied(1) = False Then Exit Sub
               str��Ʊ�� = mstrReclaimInvoice
               '1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
               If zlExeCuteBillNoSplit(False, IIf(mbytInFun = 4, 2, 3), mlng����ID, mstrPrintNO, 0, mstrInvoice, mdatFeeDate, 1, str��Ʊ��, intƱ������) = False Then Exit Sub
               mstrReclaimInvoice = str��Ʊ��
               mbln����Ʊ�� = True
               If intƱ������ = 0 Then Exit Sub
               
               If blnNotPrint Then
                    Call TaxInterface(2, mstrPrintNO, strAllNOs)      '��ӡ˰��Ʊ��
                    ''����ҽ���ش�ӿ�
                    If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
               Else
                   'Ʊ�ݽӿ�
                    If BillPrint(2, mstrPrintNO, strAllNOs, strInvoice, strClearNOs) = False Then Exit Sub
                    '���ñ���
                   Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                   "��Ʊ��=" & str��Ʊ��, "NO='NO'", "��ӡID=" & mlng��ӡID, "PrintEmpty=0", "", "", strFormat, 2)
                   If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
               End If
               Exit Sub
            End If
            
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                Call TaxInterface(2, mstrPrintNO, strAllNOs)      '��ӡ˰��Ʊ��
                    ''����ҽ���ش�ӿ�
                    If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
            Else
                If bln�ֱ��ӡ And UBound(Split(strNos, ",")) > 0 And strAllNOs = "" Then
                    For i = 0 To UBound(Split(strNos, ","))
                        mblnPrinted = False '���������ʼ����Ϊ���ܴ�ӡ�����ڵ�BeforePrint֮ǰ�ͳ�������
                        mstrPrintNO = Split(strNos, ",")(i)
                        blnPrint = True
                        If gTy_Module_Para.bln������ Then
                            'һ�ŵ���ֻ�й����Ѳ���ӡ
                            If BillOnlyFactMoney(Replace(mstrPrintNO, "'", "")) Then blnPrint = False
                        End If
                        If blnPrint Then
                            If blnNotPrint Then
                                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                                If Not mblnPrinted Then Exit For
                                Call TaxInterface(3, mstrPrintNO, "")
                            Else
                                If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit For
                                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "��ӡID=" & mlng��ӡID, "PrintEmpty=0", "", "", strFormat, 2)
                                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                                If Not mblnPrinted And mobjReport.DataIsEmpty = False Then Exit For '109708
                            End If
                            If mobjReport.DataIsEmpty = False Then '109708
                                If i < UBound(Split(strNos, ",")) Then 'ȡ��һƱ�ݺ�,��BeforePrint��ʹ��
                                    If gblnStrictCtrl Then
                                        mstrInvoice = GetNextBill(mlng����ID)   '����ʱ,���ؿ�
                                    Else
                                        mstrInvoice = IncStr(mstrInvoice)
                                    End If
                                End If
                            End If
                        End If
                    Next i
                Else
                    'Ʊ�ݽӿ�
                     If BillPrint(2, mstrPrintNO, "", strInvoice, strClearNOs) = False Then Exit Sub
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "��ӡID=" & mlng��ӡID, "PrintEmpty=0", "", "", strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then Exit Sub
                End If
            End If
        Case 3  '�˷�
            If gTy_Module_Para.bytƱ�ݷ������ <> 0 And mstrReclaimInvoice <> "" Then
                '1.����ϵͳԤ�������ӡƱ��;2-�����û��Զ������Ʊ��
                '���շ�Ʊ����Ϊ��ʱ,���ܰ��·�ʽ�ش�Ʊ��
                '�ȷ���Ʊ��:
                If CheckInvoiceValied(1) = False Then Exit Sub
                str��Ʊ�� = mstrReclaimInvoice
                '1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
                If zlExeCuteBillNoSplit(False, 4, mlng����ID, strNos, 0, mstrInvoice, mdatFeeDate, 1, str��Ʊ��, intƱ������) = False Then Exit Sub
                mstrReclaimInvoice = str��Ʊ��
                mbln����Ʊ�� = True
                If intƱ������ = 0 Then Exit Sub
                
                If blnNotPrint Then
                     Call TaxInterface(3, strNos, "")       '��ӡ˰��Ʊ��
                Else
                    'Ʊ�ݽӿ�
                  If BillPrint(3, strNos, "", strInvoice, "") = False Then Exit Sub
                     '���ñ���
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                    "��Ʊ��=" & str��Ʊ��, "NO='NO'", "��ӡID=" & mlng��ӡID, "PrintEmpty=0", "", "", strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                End If
                Exit Sub
            End If
            If bln�ֱ��ӡ And UBound(Split(strNos, ",")) > 0 And strAllNOs = "" Then
                For i = 0 To UBound(Split(strNos, ","))
                    mblnPrinted = False '���������ʼ����Ϊ���ܴ�ӡ�����ڵ�BeforePrint֮ǰ�ͳ�������
                    mstrPrintNO = Split(strNos, ",")(i)
                    blnPrint = True
                    If gTy_Module_Para.bln������ Then
                        'һ�ŵ���ֻ�й����Ѳ���ӡ
                        If BillOnlyFactMoney(Replace(mstrPrintNO, "'", "")) Then blnPrint = False
                    End If
                    If blnPrint Then
                        If blnNotPrint Then
                            Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                            If Not mblnPrinted Then Exit For
                            Call TaxInterface(3, mstrPrintNO, "")
                        Else
                            If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit For
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "��ӡID=" & mlng��ӡID, "PrintEmpty=0", "", "", strFormat, 2)
                                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                            If Not mblnPrinted And mobjReport.DataIsEmpty = False Then Exit For '109708
                        End If
                        If mobjReport.DataIsEmpty = False Then '109708
                            If i < UBound(Split(strNos, ",")) Then 'ȡ��һƱ�ݺ�,��BeforePrint��ʹ��
                                If gblnStrictCtrl Then
                                    mstrInvoice = GetNextBill(mlng����ID)   '����ʱ,���ؿ�
                                Else
                                    mstrInvoice = IncStr(mstrInvoice)
                                End If
                            End If
                        End If
                    End If
                Next i
            Else
                mstrPrintNO = strNos
                If blnNotPrint Then
                    Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                    If Not mblnPrinted Then Exit Sub
                    Call TaxInterface(3, mstrPrintNO, "")
                Else
                    If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit Sub
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "��ӡID=" & mlng��ӡID, "PrintEmpty=0", "", "", strFormat, 2)
                        If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then Exit Sub
                End If
            End If
        Case 6 '��Ʊ��ӡ
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
'                Call TaxInterface(3, mstrPrintNO, "")
            Else
'                If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit Sub
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "�������=" & Val(strNos), "PrintEmpty=0", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
    End Select
    '3.�������ʹ�õ�����ID
    lngLastUseID = mlng����ID
'    Exit Sub
ClearInvoice:
    On Error GoTo errH
    
    If strClearNOs = "" Then Exit Sub
    strClearNOs = Replace(strClearNOs, "'", "")
    If Left(strClearNOs, 1) = "," Then strClearNOs = Mid(strClearNOs, 2)
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(Split(strClearNOs, ","))
            strPrintNO = Split(strClearNOs, ",")(i)
            strSQL = "Zl_Ʊ����ʼ��_Update('" & strPrintNO & "','',1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




