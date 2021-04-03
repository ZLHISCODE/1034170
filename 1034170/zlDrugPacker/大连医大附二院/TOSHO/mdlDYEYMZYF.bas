Attribute VB_Name = "mdlDYEYMZYF"
Option Explicit

Public gobjSOAP As Object  '�ӿڶ���
Public gstrIP As String    '����ip
Public gblnShowMsg As Boolean   '�Ƿ񵯳��Ի�����ʾ�������շ���Ҫ��
Public Const gstrUnit_DYEY = "����ҽ�ƴ�ѧ�����ڶ�ҽԺ"
Public Const gstrUnit_YZSZYY = "��������ҽԺ"
Public Const gstrUnit_JLSZXYY = "����������ҽԺ"

Public Const GINT_SEND_TYPE = 1           '0-����ʼ��ҩ���̣�1-�п�ʼ��ҩ��������ҩ����
Public Const GINT_STARTSEND_TYPE = 1      '0-��ť��ʽ��ʼ��ҩ��1-ˢ����ʽ��ʼ��ҩ

Private Type IPINFO
    dwAddr As Long   ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(5) As IPINFO  'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Enum gType
    IntDrug = 101 '�ϴ�ҩƷ��������
    IntStore = 102 '�ϴ�ҩƷ�������
    IntDetail = 201 '�ϴ�������ϸ
    IntStartList = 202  '�ϴ�������������ʼ��ҩ
    IntEndList = 203    '�ϴ�����������������ҩ
End Enum

Private mStrSql As String

Public Function DYEY_MZ_TransData(ByVal intType As Integer, ByVal intOprId As Integer, ByVal strUserCode As String, ByVal strUserName As String, ByVal arrXML As Variant, ByRef strReturn As String, Optional ByVal strNO As String, Optional ByVal LngStockID As Long) As Boolean
'1.��WebService��������
'2.���ӿں�������
    Dim i As Integer
    Dim intRetval As Integer
    Dim strRetmsg As String
    Dim blnShow As Boolean
    Dim lngDrugStockID As Long
    
    On Error GoTo errHandle
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "�����ϴ���", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "�����ϴ���"
        End If
    End If
    If gstrIP = "" Then
        gstrIP = GetLocalIP
    End If
    
    For i = 0 To UBound(arrXML)
        If gobjSOAP.TransConsisData(intOprId, intType, CStr(arrXML(i)), gstrIP, strUserCode, strUserName, intRetval, strRetmsg) <> 1 Then
            If gblnShowMsg Then
                MsgBox strRetmsg, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strRetmsg
            End If
            If blnShow Then frmDYEY_MZ_TransDrug.UnloadMe
            Exit Function
        End If
        
        If intType = gType.IntDrug Or intType = gType.IntStore Then
            If i = 0 Then
                frmDYEY_MZ_TransDrug.Show
                blnShow = True
            End If
            
            Call frmDYEY_MZ_TransDrug.ChangePrg(i + 1, UBound(arrXML) + 1, intType)
        ElseIf intType = gType.IntDetail Then
            lngDrugStockID = GetStockID(arrXML(i))
            If lngDrugStockID = 0 Then lngDrugStockID = 176
            'If Not SetSendWin(LngStockID, strNO, intRetval) Then
            'If Not SetSendWin(176, strNO, intRetval) Then   '��ʱ�ⷿidΪ176''''''''''''''''''''''''''''''''''''''''''''
            If Not SetSendWin(lngDrugStockID, strNO, intRetval) Then
                If gblnShowMsg Then
                    MsgBox "���������ķ�ҩ����ʧ�ܣ�", vbCritical, GSTR_MESSAGE
                Else
                    strReturn = "���������ķ�ҩ����ʧ�ܣ�"
                End If
            End If
        End If
    Next
    
    DYEY_MZ_TransData = True
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "�ϴ���ɣ�", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "�ϴ���ɣ�"
        End If
    End If
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
End Function

Public Function GetXML_Drug() As Variant
'��ҩƷ������Ϣ��֯��ָ����XML��ʽ
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strErrMsg As String
    
    On Error GoTo errHandle
'    MsgBox "��ȡ����"
    strErrMsg = "��ȡ����"
    mStrSql = "Select Distinct a.id ҩƷ���, a.���� ҩƷ����, e.���� ҩƷ��Ʒ��, a.��� ҩƷ���, a.��� ҩƷ��װ���, b.���ﵥλ ҩƷ��λ," & vbNewLine & _
              "    round(b.ҩ���װ/b.�����װ, 2) ��װ��,b.����ɷ����,a.���� ҩƷ����, c.�ּ� * b.�����װ ҩƷ�۸�, d.ҩƷ����, " & vbNewLine & _
              "    b.�����װ, a.����ʱ�� ������ʱ��, f.���� ҩƷƴ��, d.������� " & vbNewLine & _
              "From �շ���ĿĿ¼ a, ҩƷ��� b, �շѼ�Ŀ c, ҩƷ���� d, �շ���Ŀ���� e, �շ���Ŀ���� f " & vbNewLine & _
              "Where a.Id = b.ҩƷid And a.Id = c.�շ�ϸĿid And b.ҩ��id = d.ҩ��id And a.Id = e.�շ�ϸĿid(+) And a.Id = f.�շ�ϸĿid(+) And " & vbNewLine & _
              "    (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And Sysdate Between c.ִ������ And " & vbNewLine & _
              "    Nvl(c.��ֹ����, Sysdate) And e.����(+) = 3 And f.����(+) = 1 And f.����(+) = 1"
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_Drug")
    strErrMsg = "���ݻ�ȡ���"
    strXML = ""
    arrXML = Array()
    
    strErrMsg = "XML��ʼ"
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_DRUGSVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!ҩƷ���) & """"
                strDrug = strDrug & vbCrLf & "DRUG_NAME = """ & SpecialChar(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "TRADE_NAME = """ & SpecialChar(!ҩƷ��Ʒ��) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SPEC = """ & SpecialChar(!ҩƷ���) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PACKAGE = """ & NVL(!�����װ) & """"  ' & SpecialChar(!ҩƷ��װ���) & """"
                strDrug = strDrug & vbCrLf & "DRUG_UNIT = """ & SpecialChar(!ҩƷ��λ) & """"
                strDrug = strDrug & vbCrLf & "FIRM_ID = """ & SpecialChar(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PRICE = """ & NVL(!ҩƷ�۸�) & """"
                strDrug = strDrug & vbCrLf & "DRUG_FORM = """ & SpecialChar(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SORT = """ & SpecialChar(!�������) & """"
                strDrug = strDrug & vbCrLf & "BARCODE = """""
                strDrug = strDrug & vbCrLf & "LAST_DATE = """ & Format(!������ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PINYIN = """ & SpecialChar(!ҩƷƴ��) & """"
                strDrug = strDrug & vbCrLf & "DRUG_CONVERTATION = """ & NVL(!��װ��) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_DRUGSVW>"
                
                If Len(strXML & strDrug) > 3900 Then
                    '����ǰ����ӵ�����
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "װ������1"
                    '����ƴ���µ�XML
                    strXML = strTitle & vbCrLf & strDrug
                Else
                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                End If
                
                rsTemp.MoveNext
                If .EOF And strXML <> "" Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "װ������2"
                End If
            Loop
        End If
    End With
    
    strErrMsg = "��ȡ����"
    GetXML_Drug = arrXML
    strErrMsg = "��������"
    Exit Function

errHandle:
    Debug.Print strErrMsg
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetXML_RecipeDetail(ByVal LngStockID As Long, ByVal strNO As String) As Variant
'��������ϸ��֯��ָ����XML��ʽ
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strSQL As String
    Dim i As Integer
    Dim rsDetails As Recordset
    Dim strDetail As String
    
    '��ʱ�ⷿΪ176''''''''''''''''''''''''''''''''''''''''''''''''
'    LngStockID = 176
    
    On Error GoTo errHandle
    '��ȡ��������Ϣ
    strSQL = "Select a.�������� ����ʱ��, a.����, a.No �������, a.�ⷿid ��ҩҩ��, c.����id ���￨��, a.���� ��������, Decode(a.���ȼ�, 1, '01', '00') ��������, " & vbNewLine & _
             "    c.�������� ���߳�������, c.�Ա� �����Ա�, c.��� �������, c.ҽ�Ƹ��ʽ ҽ������, Sum(d.Ӧ�ս��) ����, Sum(d.ʵ�ս��) ʵ������," & vbNewLine & _
             "    f.id ��������, d.������ ����ҽ��, d.������ ¼����, Decode(a.���ȼ�, 1, '1', '2') ��ҩ���ȼ� " & vbNewLine & _
             "From δ��ҩƷ��¼ a, ������Ϣ c, ������ü�¼ d, ҩƷ�շ���¼ e, ���ű� f " & vbNewLine & _
             "Where a.���� = e.���� And a.No = e.No And a.�ⷿid = e.�ⷿid And a.����id = c.����id And e.����id = d.Id And " & vbNewLine & _
             "    d.��������id = f.Id " & IIf(LngStockID = 0, "", " And a.�ⷿid=[1] ")

    If InStr(1, strNO, "|") < 1 Then
        strSQL = strSQL & " And a.����=[2] And a.NO=[3] "
    Else
        strSQL = strSQL & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSQL = strSQL & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSQL = strSQL & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSQL = strSQL & ") "
    End If
            
    strSQL = strSQL & _
             "Group By a.��������, a.����, a.No, a.�ⷿid, c.����id, a.����, Decode(a.���ȼ�, 1, '01', '00'), c.��������, c.�Ա�, " & vbNewLine & _
             "    c.���, c.ҽ�Ƹ��ʽ, f.id, d.������,d.������, Decode(a.���ȼ�, 1, '1', '2') "
             

    mStrSql = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
    mStrSql = "select * from (" & mStrSql & ") Order By ��ҩҩ��, ���￨�� "
    
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeDetail", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    
    
    '��ȡ������ϸ��Ϣ
    strSQL = "Select Distinct a.��������, a.����, a.No, a.���, b.id ҩƷ����, b.���� ҩƷ����, c.���� ҩƷ��Ʒ��, b.��� ҩƷ���, b.��� ҩƷ��װ���, " & vbNewLine & _
             "    d.���ﵥλ ҩƷ��λ, a.���� ҩƷ����, a.���ۼ� * d.�����װ ҩƷ�۸�, a.ʵ������ / d.�����װ ����, e.Ӧ�ս�� ����,e.����id," & vbNewLine & _
             "    e.ʵ�ս�� ʵ������, a.���� ҩƷ����, a.�ⷿid, a.�÷�, f.ִ��Ƶ��, g.���㵥λ ������λ " & vbNewLine & _
             "From ҩƷ�շ���¼ a, �շ���ĿĿ¼ b, �շ���Ŀ���� c, ҩƷ��� d, ������ü�¼ e, ����ҽ����¼ f, ������ĿĿ¼ g " & vbNewLine & _
             "Where a.ҩƷid = b.Id And a.ҩƷid = c.�շ�ϸĿid(+) And a.ҩƷid = d.ҩƷid And a.����id = e.Id and d.ҩ��id=g.id " & vbNewLine & _
             "    And e.ҽ����� = f.Id(+) And c.����(+) = 3 " & IIf(LngStockID = 0, "", " And a.�ⷿid=[1] ")

    If InStr(1, strNO, "|") < 1 Then
        strSQL = strSQL & " And a.����=[2] And a.NO=[3] "
    Else
        strSQL = strSQL & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSQL = strSQL & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSQL = strSQL & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSQL = strSQL & ") "
    End If
    
    mStrSql = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
    Set rsDetails = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeDetail", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    strXML = ""
    arrXML = Array()
    
    '�ⷿIDΪ0�����������������
    If LngStockID = 0 Then
        If GetXML_RecipeDetailEx(rsTemp, rsDetails, arrXML) Then
            GetXML_RecipeDetail = arrXML
        End If
        Exit Function
    End If
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & SpecialChar(!�������) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!��ҩҩ��) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_ID = """ & NVL(!���￨��) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!��������) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_TYPE = """ & NVL(!��������) & """"
                strDrug = strDrug & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!���߳�������), "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "SEX = """ & SpecialChar(!�����Ա�) & """"
                strDrug = strDrug & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!�������) & """"
                strDrug = strDrug & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!ҽ������) & """"
                strDrug = strDrug & vbCrLf & "PRESC_ATTR = """""
                strDrug = strDrug & vbCrLf & "PRESC_INFO = """""
                strDrug = strDrug & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!�������))
                strDrug = strDrug & vbCrLf & "RCPT_REMARK = """""
                strDrug = strDrug & vbCrLf & "REPETITION = ""1"""
                strDrug = strDrug & vbCrLf & "COSTS = """ & NVL(!����) & """"
                strDrug = strDrug & vbCrLf & "PAYMENTS = """ & NVL(!ʵ������) & """"
                strDrug = strDrug & vbCrLf & "ORDERED_BY = """ & NVL(!��������) & """"
                strDrug = strDrug & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!����ҽ��) & """"
                strDrug = strDrug & vbCrLf & "ENTERED_BY = """ & SpecialChar(!¼����) & """"
                strDrug = strDrug & vbCrLf & "DISPENSE_PRI = """ & NVL(!��ҩ���ȼ�) & """"
                strDrug = strDrug & vbCrLf & ">"
                
                '������ϸ��¼��ȷ���뵥�ݶ�Ӧ
                rsDetails.Filter = "no='" & !������� & "' and ����=" & NVL(!����) & " and �ⷿid=" & NVL(!��ҩҩ��)
                rsDetails.Sort = "���"
'                rsDetails.Filter = "no='" & !������� & "' and ��������='" & CDate(!����ʱ��) & "'"
                
                strDetail = ""
                Do While Not rsDetails.EOF
                    strDetail = strDetail & vbCrLf & "<CONSIS_PRESC_DTLVW"
                    strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetails!��������, "yyyy-MM-DDThh:mm:ss") & """"
                    strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetails!no) & """"
                    strDetail = strDetail & vbCrLf & "ITEM_NO = """ & NVL(rsDetails!���) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetails!ҩƷ��Ʒ��) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetails!ҩƷ���) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetails!ҩƷ��װ���) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetails!ҩƷ��λ) & """"
                    strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetails!ҩƷ�۸�) & """"
                    strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetails!����) & """"
                    strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetails!����) & """"
                    strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetails!ʵ������) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetails!������λ) & """"
                    strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetails!�÷�) & """"
                    strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetails!ִ��Ƶ��) & """"
                    strDetail = strDetail & vbCrLf & ">"
                    strDetail = strDetail & vbCrLf & "</CONSIS_PRESC_DTLVW>"
                    rsDetails.MoveNext
                Loop
                strDrug = strDrug & strDetail
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeDetail = arrXML
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function GetXML_RecipeDetailEx(ByVal rsBill As ADODB.Recordset, ByVal rsDetail As ADODB.Recordset, ByRef varXML As Variant) As Boolean
'���ܣ�����ⷿIDΪ0��������ִ���ⷿID�벡��ID����XML�ַ���
'������
'  rsBill���������ݼ���
'  rsDetail����ϸ���ݼ���
'  varXML�����ɵ�XML�ַ������飨ʵ�Σ���
'���أ�True�ɹ�   Falseʧ��
    Const STR_ROOT_BEGIN = "<ROOT>"
    Const STR_ROOT_END = "</ROOT>"
    Const STR_BILL = "CONSIS_PRESC_MSTVW"
    Const STR_DETAIL = "CONSIS_PRESC_DTLVW"
    Dim strXML As String, strBill As String, strDetail As String
    Dim lng�ⷿID As Long, lng����ID As Long
    Dim varReturn As Variant
    
    On Error GoTo errHandle
    varReturn = Array()
    With rsBill
        If .RecordCount <= 0 Then Exit Function
        .MoveFirst
        lng�ⷿID = NVL(!��ҩҩ��, 0)
        lng����ID = NVL(!���￨��, 0)
        Do
            If .EOF Then Exit Do
            '����
            strBill = "<" & STR_BILL & " "
            strBill = strBill & vbCrLf & "PRESC_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "PRESC_NO = """ & SpecialChar(!�������) & """"
            strBill = strBill & vbCrLf & "DISPENSARY = """ & NVL(!��ҩҩ��) & """"
            strBill = strBill & vbCrLf & "PATIENT_ID = """ & NVL(!���￨��) & """"
            strBill = strBill & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!��������) & """"
            strBill = strBill & vbCrLf & "PATIENT_TYPE = """ & NVL(!��������) & """"
            strBill = strBill & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!���߳�������), "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "SEX = """ & SpecialChar(!�����Ա�) & """"
            strBill = strBill & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!�������) & """"
            strBill = strBill & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!ҽ������) & """"
            strBill = strBill & vbCrLf & "PRESC_ATTR = """""
            strBill = strBill & vbCrLf & "PRESC_INFO = """""
            strBill = strBill & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!�������))
            strBill = strBill & vbCrLf & "RCPT_REMARK = """""
            strBill = strBill & vbCrLf & "REPETITION = ""1"""
            strBill = strBill & vbCrLf & "COSTS = """ & NVL(!����) & """"
            strBill = strBill & vbCrLf & "PAYMENTS = """ & NVL(!ʵ������) & """"
            strBill = strBill & vbCrLf & "ORDERED_BY = """ & NVL(!��������) & """"
            strBill = strBill & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!����ҽ��) & """"
            strBill = strBill & vbCrLf & "ENTERED_BY = """ & SpecialChar(!¼����) & """"
            strBill = strBill & vbCrLf & "DISPENSE_PRI = """ & NVL(!��ҩ���ȼ�) & """"
            strBill = strBill & vbCrLf & ">"
            
            '������ϸ��¼��ȷ���뵥�ݶ�Ӧ
            strDetail = ""
            rsDetail.Filter = "no='" & !������� & "' and ����=" & NVL(!����) & " and �ⷿid=" & NVL(!��ҩҩ��) & " and ����id=" & NVL(!���￨��)
            rsDetail.Sort = "���"
            Do
                If rsDetail.EOF Then Exit Do
                '��ϸ
                strDetail = strDetail & vbCrLf & "<" & STR_DETAIL & " "
                strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetail!��������, "yyyy-MM-DDThh:mm:ss") & """"
                strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetail!no) & """"
                strDetail = strDetail & vbCrLf & "ITEM_NO = """ & rsDetail!��� & """"
                strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetail!ҩƷ��Ʒ��) & """"
                strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetail!ҩƷ���) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetail!ҩƷ��װ���) & """"
                strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetail!ҩƷ��λ) & """"
                strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetail!ҩƷ�۸�) & """"
                strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetail!����) & """"
                strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetail!����) & """"
                strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetail!ʵ������) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetail!������λ) & """"
                strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetail!�÷�) & """"
                strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetail!ִ��Ƶ��) & """"
                strDetail = strDetail & vbCrLf & ">"
                strDetail = strDetail & vbCrLf & "</" & STR_DETAIL & ">"
                rsDetail.MoveNext
            Loop While Not rsDetail.EOF
            
            strBill = strBill & strDetail
            strBill = strBill & "</" & STR_BILL & ">"
            
            '��ֲ�ͬ�ⷿID�Ͳ���ID�ĵ�����ϸ
            If lng�ⷿID = NVL(!��ҩҩ��, 0) And lng����ID = NVL(!���￨��, 0) Then
                strXML = strXML & strBill & vbCrLf
            Else
                strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
                ReDim Preserve varReturn(UBound(varReturn) + 1)
                varReturn(UBound(varReturn)) = strXML
                strXML = strBill & vbCrLf
            End If
            
            lng�ⷿID = NVL(!��ҩҩ��, 0)
            lng����ID = NVL(!���￨��, 0)
            
            .MoveNext
        Loop While Not .EOF
        
        strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
        ReDim Preserve varReturn(UBound(varReturn) + 1)
        varReturn(UBound(varReturn)) = strXML
        varXML = varReturn
        GetXML_RecipeDetailEx = True
    
    End With
    
    Exit Function
    
errHandle:
    Set varXML = Nothing
End Function

Public Function GetXML_RecipeList(ByVal LngStockID As Long, ByVal strNO As String) As Variant
'����������֯��ָ����XML��ʽ
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant, arrTmp As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    mStrSql = "Select ��������,No From ҩƷ�շ���¼ Where �ⷿid=[1]"
    
    If InStr(1, strNO, "|") < 1 Then
        mStrSql = mStrSql & " And ����=[2] And NO=[3]"
    Else
        mStrSql = mStrSql & " And ("
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(arrTmp)
            If i = UBound(arrTmp) Then
                mStrSql = mStrSql & "(����=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "')"
            Else
                mStrSql = mStrSql & "(����=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "') or "
            End If
        Next
        mStrSql = mStrSql & ")"
    End If
    mStrSql = mStrSql & " and (��¼״̬=1 or mod(��¼״̬,3)=1) "
    
    If InStr(1, strNO, "|") < 1 Then
        Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeList", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeList", LngStockID)
    End If
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!��������, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & NVL(!no) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeList = arrXML
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetXML_Stock(ByVal LngStockID As Long) As Variant
'��ҩƷ�����Ϣ��֯��ָ����XML��ʽ
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    
    On Error GoTo errHandle
    mStrSql = "Select a.id ҩƷ���,c.�ⷿid ��ҩҩ��,sum(c.ʵ������/e.�����װ) ҩƷ����,d.�ⷿ��λ ҩƷ��λ " & vbNewLine & _
              "From �շ���ĿĿ¼ a, ҩƷ��� c, ҩƷ�����޶� d,ҩƷ��� e " & vbNewLine & _
              "Where a.Id = c.ҩƷid And e.ҩƷid=c.ҩƷid And d.�ⷿid(+) = c.�ⷿid And d.ҩƷid(+) = c.ҩƷid And c.�ⷿid=[1] " & vbNewLine & _
              "Group By a.id, c.�ⷿid, d.�ⷿ��λ " & vbNewLine & _
              "Having Sum(c.ʵ������/e.�����װ)<>0 "

    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_Stock", LngStockID)
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PHC_STORAGEVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!ҩƷ���) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!��ҩҩ��) & """"
                strDrug = strDrug & vbCrLf & "DRUG_QUANTITY = """ & NVL(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "LOCATIONINFO = """ & SpecialChar(!ҩƷ��λ) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PHC_STORAGEVW>"

'��ҵ���ܿ��Բ���4K����
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                
'                If Len(strXML & strDrug) > 3900 Then
'                    '����ǰ����ӵ�����
'                    strXML = strXML & vbCrLf & "</ROOT>"
'                    ReDim Preserve arrXML(UBound(arrXML) + 1)
'                    arrXML(UBound(arrXML)) = strXML
'
'                    '����ƴ���µ�XML
'                    strXML = strTitle & vbCrLf & strDrug
'                Else
'                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
'                End If
                
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_Stock = arrXML
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function SetSendWin(ByVal LngStockID As Long, ByVal strNO As String, ByVal intOpr As Integer) As Boolean
'����HIS��ָ�������ķ�ҩ����
    Dim i As Integer
    Dim arrTmp As Variant
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    mStrSql = "Select ���� From ��ҩ���� Where ҩ��id=[1] And ����=[2]"
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "SetSendWin", LngStockID, CStr(intOpr))
    
    If Not rsTemp.EOF Then
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(Split(strNO, "|"))
            mStrSql = "Zl_δ��ҩƷ��¼_���䷢ҩ����("
'            mStrSql = mStrSql & "'" & Split(Split(strNO, "|"), ",")(1) & "',"
'            mStrSql = mStrSql & Split(Split(strNO, "|"), ",")(0) & ","
            mStrSql = mStrSql & "'" & Split(arrTmp(i), ",")(1) & "',"
            mStrSql = mStrSql & Split(arrTmp(i), ",")(0) & ","
            mStrSql = mStrSql & LngStockID & ","
            mStrSql = mStrSql & "'" & rsTemp!���� & "')"
            Call gobjComLib.zldatabase.ExecuteProcedure(mStrSql, "SetSendWin")
        Next
        SetSendWin = True
    Else
        If gblnShowMsg Then
            MsgBox "û���ҵ�����Ϊ��" & intOpr & "���Ĵ��ڣ����飡", vbCritical, GSTR_MESSAGE
        End If
    End If
    
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gobjComLib.ErrCenter() = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
End Function


Public Function GetLocalIP() As String
'ȡ����IP
    Dim Ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    
    
    On Error GoTo EndRow
        GetIpAddrTable ByVal 0&, Ret, True
    
    
        If Ret <= 0 Then Exit Function
        ReDim bBytes(0 To Ret - 1) As Byte
        ReDim TempList(0 To Ret - 1) As String
        
        'retrieve the data
        GetIpAddrTable bBytes(0), Ret, False
          
        'Get the first 4 bytes to get the entry's.. ip installed
        CopyMemory Listing.dEntrys, bBytes(0), 4
        
        For Tel = 0 To Listing.dEntrys - 1
            'Copy whole structure to Listing..
            CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
            TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        Next Tel
        'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        GetLocalIP = TempIP 'Return The TempIP
    Exit Function
EndRow:
    GetLocalIP = ""
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Private Function GetRCPT_INFO(ByVal strNO As String) As String
'���ܣ���ȡ�����Ϣ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select MAX(DECODE(Id,1,�������,''))||';'||MAX(DECODE(Id,2,�������,'')) as ��� " & vbNewLine & _
             "From ( " & vbNewLine & _
             "      Select Rownum As Id,������� " & vbNewLine & _
             "      From (Select �������||decode(�Ƿ�����,1,'?','') ������� " & vbNewLine & _
             "            From ������ϼ�¼ " & vbNewLine & _
             "            Where ����id=(Select distinct ����id " & vbNewLine & _
             "                          From ( Select a.����id From ������ü�¼ a Left Join ����ҽ����¼ b On a.ҽ�����=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And ��¼����=1 ) ) " & vbNewLine & _
             "              And ��ҳid=(Select distinct Case When ��ҳid Is Null Then (Select Id From ���˹Һż�¼ Where No=c.�Һŵ�) Else ��ҳId End As ��ҳid " & vbNewLine & _
             "                          From ( Select null ��ҳid, b.�Һŵ� From ������ü�¼ a Left Join ����ҽ����¼ b On a.ҽ�����=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And ��¼����=1 ) c ) " & vbNewLine & _
             "union all " & vbNewLine & _
             "Select a.ժҪ As ������� From ���˹Һż�¼ a " & vbNewLine & _
             "Where No= (Select distinct Case When b.�Һŵ� Is Null Then ' ' Else b.�Һŵ� End As No " & vbNewLine & _
             "           From ������ü�¼ a Left Join ����ҽ����¼ b On a.ҽ����� = b.Id " & vbNewLine & _
             "           Where a.No = [1] And ��¼���� = 1 ) ) ) "
    On Error GoTo errHandle
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(strSQL, "��ȡ�����Ϣ", strNO)
    If Not rsTemp.EOF Then
        GetRCPT_INFO = IIf(Trim(NVL(rsTemp!���)) = ";", """""", """" & Trim(NVL(rsTemp!���)) & """")
    Else
        GetRCPT_INFO = """"""
    End If
    rsTemp.Close
    Exit Function
    
errHandle:
    GetRCPT_INFO = """"""
End Function

Private Function SpecialChar(ByVal strVal As Variant) As String
'���ܣ������ַ�ת��
'˵����
' < ת &lt;
' > ת &gt;
' & ת &amp;
' ' ת &apos;
' " ת &quot;
    Dim strReturn As String
    
    If IsNull(strVal) Then
        strVal = ""
        GoTo errHandle
    End If
    If strVal = "" Then
        GoTo errHandle
    End If
    On Error GoTo errHandle
    strReturn = strVal
    strReturn = Replace(strReturn, "<", "&lt;")
    strReturn = Replace(strReturn, ">", "&gt;")
    strReturn = Replace(strReturn, "&", "&amp;")
    strReturn = Replace(strReturn, "'", "&apos;")
    strReturn = Replace(strReturn, """", "&quot;")
    SpecialChar = strReturn
    Exit Function
    
errHandle:
    SpecialChar = strVal
End Function

Private Function GetStockID(ByVal strText As String) As Long
'���ܣ���ȡXML�ı��е�ҩ��ID
    Const STR_KEY = "DISPENSARY = "
    Dim LngStockID As Long
    Dim intStart As Integer
    
    If strText = "" Then Exit Function
    
    intStart = InStr(strText, STR_KEY)
    If intStart > 0 Then
        LngStockID = Val(Mid(strText, intStart + Len(STR_KEY) + 1))
    End If
    GetStockID = LngStockID
    
End Function

