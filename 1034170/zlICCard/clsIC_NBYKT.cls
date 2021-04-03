VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIC_NBYKT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

Private Const strFile As String = "C:\APPSOFT\NBYKT.INI"
Private gstr��ַ As String, gstr�˿ں� As String, gstrĿ¼ As String, gstr�û��� As String, gstr���� As String, gstr������ַ As String, gbln��Ϣת�� As Boolean
Private gstrҽԺ���� As String, gstrIP As String, gstrMAC As String
Private mdomOutput As New MSXML2.DOMDocument

'------------------------------------------------------------
'������ȡ����MAC��ַ�������
Private Const MAX_ADAPTER_NAME_LENGTH As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH As Long = 8
Private Const ERROR_SUCCESS As Long = 0

Private Type MAC_ADDRESS_STRING
IpAddr(0 To 15) As Byte
End Type

Private Type MAC_MASK_STRING
IpMask(0 To 15) As Byte
End Type

Private Type MAC_ADDR_STRING
dwNext As Long
IpAddress As MAC_ADDRESS_STRING
IpMask As MAC_MASK_STRING
dwContext As Long
End Type

Private Type MAC_ADAPTER_INFO
dwNext As Long
ComboIndex As Long '����
sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3)) As Byte
sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
dwAddressLength As Long
sMACAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1)) As Byte
dwIndex As Long
uType As Long
uDhcpEnabled As Long
CurrentIpAddress As Long
IpAddressList As MAC_ADDR_STRING
GatewayList As MAC_ADDR_STRING
DhcpServer As MAC_ADDR_STRING
bHaveWins As Long
PrimaryWinsServer As MAC_ADDR_STRING
SecondaryWinsServer As MAC_ADDR_STRING
LeaseObtained As Long
LeaseExpires As Long
End Type

Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (pTcpTable As Any, pdwSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
'------------------------------------------------------------


'------------------------------------------------------------
'������ȡ�����������
Private Declare Function GetPWD Lib "Authen.dll" (ByVal strUser As String) As String
'------------------------------------------------------------

Private Function LocalMACAddress() As String
    Dim cbRequired As Long
    Dim buff() As Byte
    Dim Adapter As MAC_ADAPTER_INFO
    Dim AdapterStr As MAC_ADDR_STRING
    Dim ptr1 As Long
    Dim sIPAddr As String
    Dim found As Boolean
    Dim iFound As Integer
    Dim sReturn As String
    
    iFound = 0
    sIPAddr = ""
    sReturn = ""
    Call GetAdaptersInfo(ByVal 0&, cbRequired)
    If cbRequired = 0 Then Exit Function
    ReDim buff(0 To cbRequired - 1) As Byte
    If GetAdaptersInfo(buff(0), cbRequired) <> ERROR_SUCCESS Then Exit Function
    
    '��ȡ�����buff()�е����ݵ�ָ��
    ptr1 = VarPtr(buff(0))
    Do While (ptr1 <> 0)
        '����һ������������ת����MAC_ADAPTER_INFO�ṹ��
        CopyMemory Adapter, ByVal ptr1, LenB(Adapter)
        With Adapter
            'IpAddress.IpAddr��Ա������DHCP��IP��ַ
            Dim k As Long
            For k = 1 To .dwAddressLength
                sReturn = sReturn & Right("0" & Hex(AscB(MidB(.sMACAddress, k, 1))), 2) & "-"
            Next k
            sReturn = Left(sReturn, Len(sReturn) - 1) & vbCrLf
            ptr1 = .dwNext
        End With 'With Adapter
        '����������ʱ��ptr1��ֵΪ0
    Loop 'Do While (ptr1 <> 0)
    '���ؽ���ַ���
    LocalMACAddress = sReturn
End Function

Private Function zl_Ip_Address_FromOrc(Optional strDefaultIp_Address As String = "") As String
    '-----------------------------------------------------------------------------------------------------------
    '����:ͨ��oracle��ȡ�ļ������IP��ַ
    '���:strDefaultIp_Address-ȱʡIP��ַ
    '����:
    '����:����IP��ַ
    '����:���˺�
    '����:2009-01-21 11:08:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strIp_Address As String, strSQL As String
    Err = 0: On Error GoTo errHand:
     strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡIP��ַ")
    If rsTemp.EOF = False Then
        strIp_Address = gobjCommFun.Nvl(rsTemp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = strDefaultIp_Address
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    zl_Ip_Address_FromOrc = strIp_Address
    Exit Function
errHand:
    
End Function

Private Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp >= lngLen Then
        '��Ҫ�пո������
        strTmp = ToVarchar(strCode, lngLen)
        lngTmp = LenB(StrConv(strTmp, vbFromUnicode))
    End If
    
    strTmp = strTmp & String(lngLen - lngTmp, strChar)
    'ȡ��������ַ�
    'Rpad = Replace(strTmp, Chr(0), strChar)
    Rpad = strTmp
End Function

Private Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'���ܣ����ı���Varchar2�ĳ��ȼ��㷽�����нض�
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Dim varReturn As Variant
    varReturn = IIf(IsNull(varValue), DefaultValue, varValue)
    Nvl = Replace(varReturn, "'", "")
End Function

Private Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    
    rsTemp.Open strSQL, gcnConnect, adOpenStatic, adLockReadOnly
    Set rsTemp.ActiveConnection = Nothing
End Sub

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_CloseCard = True
End Function

Private Function clsCardDev_GetSpare() As Currency
    '
End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim strLine As String
    Dim strName As String
    Dim strValue As String
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
  
     Dim rsTemp  As New ADODB.Recordset
     Dim agcnconnect As New ADODB.Connection
     Dim conn As String, str�û� As String, str���� As String
     
    On Error GoTo errHand
    
   
      '�������ļ��ж�ȡ��ַ\Ŀ¼���˿ں�
       str�û� = "system"
       str���� = "abc123"
    gbln��Ϣת�� = False
    If objFileSys.FileExists(strFile) Then
        Set objStream = objFileSys.OpenTextFile(strFile, ForReading)
        Do While Not objStream.AtEndOfStream
            strLine = objStream.ReadLine
            strName = Trim(Split(strLine, "=")(0))
            strValue = Trim(Split(strLine, "=")(1))
            Select Case strName
            Case "��ַ"
                gstr��ַ = strValue
            Case "�˿ں�"
                gstr�˿ں� = strValue
            Case "Ŀ¼"
                gstrĿ¼ = strValue
            Case "��Ϣת��"
                gbln��Ϣת�� = (Val(strValue) = 1)
            Case "������"
                  conn = Nvl(strValue, "orcl")
            Case "�û�"
                  str�û� = Nvl(strValue, "system")
            Case "����"
                  str���� = Nvl(strValue, "abc123")
            End Select
        Loop
        objStream.Close
    Else
        gLastErr = "δ��������һ��ͨ����!"
        Exit Function
    End If
    
    
    
    
    If gcnConnect Is Nothing Then '���û�д���gcnconnect���Լ�ȡ��his����
       
        If OraDataOpen(agcnconnect, conn, str�û�, str����, False) = False Then
                MsgBox "�޷����ӵ��м�⣬�������ò������м�������Ƿ�������ȷ��"
                Exit Function
            End If
        Set gcnConnect = agcnconnect
    
    End If
    
    Dim strSQL As String
    rsTemp.CursorLocation = adUseClient
    If rsTemp.State = 1 Then rsTemp.Close
    strSQL = " Select ҽԺ���� From һ��ͨĿ¼ Where ����=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����Ƿ���ڸò��˵���Ϣ", "����һ��ͨ")
    If rsTemp.RecordCount = 0 Then Exit Function
    If IsNull(rsTemp!ҽԺ����) Then Exit Function
    gstrҽԺ���� = rsTemp!ҽԺ����
    gstrIP = zl_Ip_Address_FromOrc()
    gstrMAC = Replace(Split(LocalMACAddress, vbCrLf)(0), "-", "")
      'ȡ�û���
    gstr�û��� = gstrҽԺ���� & gstrMAC
    'ȡ����
    gstr���� = GetPWD(gstr�û���)
'    gstr�û��� = "100100002170A67A35"
'    gstr���� = "CCBA1FDD5E7C1CB3E342D50AB2EA28FD"
    
  
    
    gstr������ַ = "http://" & gstr��ַ & IIf(gstr�˿ں� <> "", ":" & gstr�˿ں�, "") & "/" & gstrĿ¼ & "/PersonInfoService?wsdl"
    clsCardDev_initCard = True
    Exit Function
errHand:
    gLastErr = Err.Description
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    '
End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    '����
    clsCardDev_ReadCard = frmIdentifyNBYKT.ReadCard(gstr������ַ, gstr�û���, gstr����, gbln��Ϣת��)
    If clsCardDev_ReadCard = "" Then blnCancel = True
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    '
End Function

'��������ֵ�ľֲ�����
Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    frmSet_NBYKT.Show 1
End Sub

Private Function Get��Ϣ�ӱ�(ByVal lng����ID As Long) As String
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "Select ��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ�ӱ�", lng����ID, "ʡ")
    If rsTemp.RecordCount <> 0 Then
        Get��Ϣ�ӱ� = Nvl(rsTemp!��Ϣֵ)
    Else
        Get��Ϣ�ӱ� = ""
    End If
    
    strSQL = "Select ��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ�ӱ�", lng����ID, "��")
    If rsTemp.RecordCount <> 0 Then
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|" & Nvl(rsTemp!��Ϣֵ)
    Else
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|"
    End If
    
    strSQL = "Select ��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ�ӱ�", lng����ID, "�ֵ�")
    If rsTemp.RecordCount <> 0 Then
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|" & Nvl(rsTemp!��Ϣֵ)
    Else
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|"
    End If
    
    strSQL = "Select ��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ�ӱ�", lng����ID, "��λ��ַ")
    If rsTemp.RecordCount <> 0 Then
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|" & Nvl(rsTemp!��Ϣֵ)
    Else
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|"
    End If
    
    strSQL = "Select ��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ�ӱ�", lng����ID, "�ֻ���")
    If rsTemp.RecordCount <> 0 Then
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|" & Nvl(rsTemp!��Ϣֵ)
    Else
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|"
    End If
    
    strSQL = "Select ��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ�ӱ�", lng����ID, "EMAIL")
    If rsTemp.RecordCount <> 0 Then
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|" & Nvl(rsTemp!��Ϣֵ)
    Else
        Get��Ϣ�ӱ� = Get��Ϣ�ӱ� & "|"
    End If
End Function

Private Function clsCardDev_UploadSwap(ByVal lng����ID As Long, strOut As String) As Boolean
    Dim lngID As Long
    Dim strSQL As String
    Dim str������Ϣ�ӱ� As String
    Dim strPatient As String
    Dim rsCard As New ADODB.Recordset
    Dim rsPatient As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '����ģ��:���￨����,����ҽ����ɽ���,�ҺŽ���,��Ժ�Ǽ�,��Ժ�Ǽ�
    '��ɽ�������Ϣ�Լ����ŵĸ����ϴ�
   If clsCardDev_initCard() = False Then
      Exit Function
      End If
      
   
 '   MsgBox "׼���ϴ���"
    str������Ϣ�ӱ� = Get��Ϣ�ӱ�(lng����ID)
    
    If rsPatient.State = 1 Then rsPatient.Close
    strSQL = " Select ����ID,סԺ���� AS ��ҳID,IC����,���￨��,��������,����,�Ա�,��������,����֤��,����֤��,����״��,��ͥ��ַ,��ͥ��ַ�ʱ�," & _
                   "        ��ͥ�绰,������λ,��λ�ʱ�,ְҵ,��λ�绰,��ϵ������,��ϵ�˵绰,һ��ͨ����ʱ�� " & _
                   " From ������Ϣ" & _
                   " Where ����ID=[1]"
    Set rsPatient = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID)
    If rsPatient.RecordCount = 0 Then Exit Function
    '������¼���ֻ����һ�����������ļ�¼
    If rsCard.State = 1 Then rsCard.Close
    strSQL = " Select A.�ɿ���,A.�ɿ�����,A.�ɿ�����ҽԺ,A.�¿���,A.����ʱ��,A.�ɿ�����ʱ��,A.�ɿ����� " & _
                " From ���˷�����¼ A" & _
                " Where Nvl(�ϴ���־,0)=0 And ����ID=[1] And ����ʱ��=(Select MAX(����ʱ��) From ���˷�����¼ Where ����ID=[1])" & _
                " Order by ����ʱ��"
    Set rsCard = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���˷�����¼", lng����ID)
    If rsCard.RecordCount = 0 Then Exit Function       '���ܸ����˲�����Ϣ��,���Բ����˳�
    
    '-----------------------------------------------
    '�������֤�����뷽ʽ(�ɿ�����=9),��ɾ���ɿ�����
 '   MsgBox "׼���ϴ��ɿ�����"
    If Val(rsCard!�ɿ�����) = 9 Then
    '    MsgBox "�����ϴ��ɿ�����"
        '׼���ϴ���������
        lngID = gobjDatabase.GetNextId("LOGID")
        strPatient = Rpad(gstrҽԺ���� & Format(gobjDatabase.CurrentDate, "yyyyMMdd") & Right("00000000" & CStr(lngID), 8), 30)       'LOGID
        strPatient = strPatient & Rpad(gstrIP, 15)                                                                  'IP
        strPatient = strPatient & Rpad(gstrMAC, 48)                                                                 'MAC
        strPatient = strPatient & gstrҽԺ����                                                                      'ҽԺ����
        If IsNull(rsPatient!IC����) Or IsNull(rsPatient!��������) Then
            strPatient = strPatient & Rpad("", 22)             '������,�½����Ĳ��˴���
        Else
            strPatient = strPatient & Rpad(IIf(Val(rsPatient!IC����) = lng����ID, "", rsPatient!IC����), 22)            '������,�½����Ĳ��˴���
        End If
        strPatient = strPatient & Rpad(rsPatient!����, 20)                                                          '����
        strPatient = strPatient & IIf(rsPatient!�Ա� = "��", 0, 1)                                                  '�Ա�
        strPatient = strPatient & Format(rsPatient!��������, "yyyyMMdd")                                            '��������
        strPatient = strPatient & IIf(IsNull(rsPatient!����֤��), 9, 0)                                             '֤������
        strPatient = strPatient & IIf(IsNull(rsPatient!����֤��), Rpad(Nvl(rsPatient!����֤��), 60), Rpad(Nvl(rsPatient!����֤��), 60))   '֤������
        
        'ȡ����Ѫ��
        strSQL = "Select ��Ϣֵ From ������ҳ�ӱ� Where ����ID=[1] And ��ҳID=[2] And ��Ϣ��=[3]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ȡ����Ѫ��", rsPatient!����ID, rsPatient!��ҳID, "Ѫ��")
        If rsTemp.RecordCount <> 0 Then
            strPatient = strPatient & Left(rsTemp!��Ϣֵ, 1)                                                        'Ѫ��
        Else
            strPatient = strPatient & " "                                                                           'Ѫ��
        End If
        
        strPatient = strPatient & IIf(rsPatient!����״�� = "�ѻ�", 0, 1)                                            '����״��
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(5), 60)                                           'EMAIL
        strPatient = strPatient & "2"                                                                   '������
        strPatient = strPatient & Rpad(Trim(rsCard!�ɿ���), 20)                                                     '����
        If Trim(rsCard!�ɿ���) <> "" Then
            strPatient = strPatient & IIf(rsPatient!�������� = "����", Rpad(Trim(rsCard!�ɿ�����), 20), String(20, " "))                             'ҽ����/���￨����
            strPatient = strPatient & rsCard!�ɿ�����ʱ��                                                               '����ʱ��
            strPatient = strPatient & Rpad(rsCard!�ɿ�����ҽԺ, 10)                                                              '�ɿ�����ҽԺ
        Else
            strPatient = strPatient & String(20, " ")                             'ҽ����/���￨����
            strPatient = strPatient & String(14, " ")                                                              '����ʱ��
            strPatient = strPatient & String(10, " ")                                                              '�ɿ�����ҽԺ
        End If
        strPatient = strPatient & IIf(Trim(rsCard!�ɿ���) <> "", "4", "5")                                                                             '��״̬
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '����ʱ��
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ͥ��ַ), 100)                                                '��ͥ��ַ
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(0), 6)                                            'ʡ/��
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(1), 6)                                            '��
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(2), 6)                                            '�ֵ�
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ͥ��ַ�ʱ�), 6)                                                  '��ͥ�ʱ�
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ͥ�绰), 20)                                                 '��ͥ�绰
        strPatient = strPatient & Rpad(Nvl(rsPatient!������λ), 100)                                                '������λ
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(3), 100)                                          '��λ��ַ
        strPatient = strPatient & Rpad(Nvl(rsPatient!��λ�ʱ�), 6)                                                  '��λ�ʱ�
        strPatient = strPatient & Rpad(Nvl(rsPatient!ְҵ), 100)                                                    'ְҵ
        strPatient = strPatient & Rpad(Nvl(rsPatient!��λ�绰), 20)                                                 '��λ�绰
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(4), 20)                                           '�ֻ���
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ϵ������), 20)                                               '��������
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ϵ�˵绰), 60)                                               '�����绰
        strPatient = strPatient & "1"                                                                               '��������
        'strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                               '����ʱ��
'        If Trim(rsCard!�ɿ���) <> "" Then
'            strPatient = strPatient & Rpad(Trim(rsCard!�ɿ���), 20)                                                           '�ɿ���
'            strPatient = strPatient & rsCard!�ɿ�����ҽԺ                                                               '�ɿ�����ҽԺ
'        Else
            strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
            strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
'        End If
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        strPatient = strPatient & rsPatient!һ��ͨ����ʱ��                                                          '����ʱ��
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '����ʱ��
        Call ���ýӿ�("putPersonInfo", strPatient)
    End If
    '-----------------------------------------------
    
    
    '-----------------------------------------------
    
    '׼���ϴ���������
  '  MsgBox "׼���ϴ��¿�����"
   ' If Not IsNull(rsCard!�¿���) Or (IsNull(rsCard!�¿���) And (rsCard!�ɿ����� = 2 Or rsPatient!�������� Like "���*")) Then
     If Not IsNull(rsCard!�¿���) Or (IsNull(rsCard!�¿���) And rsCard!�ɿ����� = 2) Then
     
      ' MsgBox "�����ϴ��¿�����"
        lngID = gobjDatabase.GetNextId("LOGID")
        strPatient = Rpad(gstrҽԺ���� & Format(gobjDatabase.CurrentDate, "yyyyMMdd") & Right("00000000" & CStr(lngID), 8), 30)       'LOGID
        strPatient = strPatient & Rpad(gstrIP, 15)                                                                  'IP
        strPatient = strPatient & Rpad(gstrMAC, 48)                                                                 'MAC
        strPatient = strPatient & gstrҽԺ����                                                                      'ҽԺ����
        If IsNull(rsPatient!IC����) Or (rsPatient!�������� Like "���*") Or IsNull(rsPatient!��������) Then
            strPatient = strPatient & Rpad("", 22)             '������,�½����Ĳ��˴���
        Else
            strPatient = strPatient & Rpad(IIf(Val(rsPatient!IC����) = lng����ID, "", rsPatient!IC����), 22)            '������,�½����Ĳ��˴���
        End If
        strPatient = strPatient & Rpad(rsPatient!����, 20)                                                          '����
        strPatient = strPatient & IIf(rsPatient!�Ա� = "��", 0, 1)                                                  '�Ա�
        strPatient = strPatient & Format(rsPatient!��������, "yyyyMMdd")                                            '��������
        strPatient = strPatient & IIf(IsNull(rsPatient!����֤��), 9, 0)                                             '֤������
        strPatient = strPatient & IIf(IsNull(rsPatient!����֤��), Rpad(Nvl(rsPatient!����֤��), 60), Rpad(Nvl(rsPatient!����֤��), 60))   '֤������
        
        'ȡ����Ѫ��
        strSQL = "Select ��Ϣֵ From ������ҳ�ӱ� Where ����ID=[1] And ��ҳID=[2] And ��Ϣ��=[3]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ȡ����Ѫ��", rsPatient!����ID, rsPatient!��ҳID, "Ѫ��")
        If rsTemp.RecordCount <> 0 Then
            strPatient = strPatient & Left(rsTemp!��Ϣֵ, 1)                                                        'Ѫ��
        Else
            strPatient = strPatient & " "                                                                           'Ѫ��
        End If
        
        strPatient = strPatient & IIf(rsPatient!����״�� = "�ѻ�", 0, 1)                                            '����״��
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(5), 60)                                           'EMAIL
        '�¿���Ϊ��˵�����˱�������ͨ�þ��￨����Ϊ��˵�������µ�ͨ�þ��￨�����Զ��Ӳ�����Ϣ��ȡ���￨�ż���
        strPatient = strPatient & IIf(IsNull(rsCard!�¿���), rsCard!�ɿ�����, IIf(rsCard!�ɿ����� = 0 Or rsCard!�ɿ����� = 1, rsCard!�ɿ�����, 2))                                '������
        strPatient = strPatient & Rpad(Nvl(rsPatient!���￨��, Nvl(rsCard!�¿���, rsCard!�ɿ���)), 20)                                                    '����
        If InStr(1, "0,1", IIf(IsNull(rsCard!�¿���), rsCard!�ɿ�����, 2)) <> 0 Then    'ҽ����
            strPatient = strPatient & Rpad(Nvl(rsPatient!���￨��, Nvl(rsCard!�¿���, rsCard!�ɿ���)), 20)
            strPatient = strPatient & Format(rsCard!����ʱ��, "              ")                                         '����ʱ��
            If rsPatient!�������� Like "���*" Then
                strPatient = strPatient & Rpad("", 10)                                                            '����ҽԺ
            Else
                strPatient = strPatient & Rpad(gstrҽԺ����, 10)                                                            '����ҽԺ
            End If
        Else
            strPatient = strPatient & Rpad(Mid(rsPatient!���￨��, 1, 4) & Mid(rsPatient!���￨��, 11), 20)         'ҽ����/���￨����
            If Not IsNull(rsCard!�¿���) Then
                strPatient = strPatient & Format(rsCard!����ʱ��, "yyyyMMddHHmmss")                                         '����ʱ��
                strPatient = strPatient & Rpad(gstrҽԺ����, 10)                                                            '����ҽԺ
            Else
                '��Ժ����
                strPatient = strPatient & rsCard!�ɿ�����ʱ��                                         '����ʱ��
                strPatient = strPatient & Rpad(rsCard!�ɿ�����ҽԺ, 10)                                                            '����ҽԺ
            End If
        End If
        strPatient = strPatient & "0"                                                                               '��״̬
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '����ʱ��
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ͥ��ַ), 100)                                                '��ͥ��ַ
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(0), 6)                                            'ʡ/��
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(1), 6)                                            '��
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(2), 6)                                            '�ֵ�
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ͥ��ַ�ʱ�), 6)                                                  '��ͥ�ʱ�
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ͥ�绰), 20)                                                 '��ͥ�绰
        strPatient = strPatient & Rpad(Nvl(rsPatient!������λ), 100)                                                '������λ
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(3), 100)                                          '��λ��ַ
        strPatient = strPatient & Rpad(Nvl(rsPatient!��λ�ʱ�), 6)                                                  '��λ�ʱ�
        strPatient = strPatient & Rpad(Nvl(rsPatient!ְҵ), 100)                                                    'ְҵ
        strPatient = strPatient & Rpad(Nvl(rsPatient!��λ�绰), 20)                                                 '��λ�绰
        strPatient = strPatient & Rpad(Split(str������Ϣ�ӱ�, "|")(4), 20)                                           '�ֻ���
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ϵ������), 20)                                               '��������
        strPatient = strPatient & Rpad(Nvl(rsPatient!��ϵ�˵绰), 60)                                               '�����绰
        'strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                               '����ʱ��
'        If Trim(rsCard!�ɿ���) <> "" And Not IsNull(rsCard!�¿���) Then
'            strPatient = strPatient & "1"                                                                           '��������
'            strPatient = strPatient & Rpad(Trim(rsCard!�ɿ���), 20)                                                       '�ɿ���
'            strPatient = strPatient & rsCard!�ɿ�����ҽԺ                                                           '�ɿ�����ҽԺ
'        Else
'            strPatient = strPatient & "0"                                                                           '��������
'            strPatient = strPatient & String(20, " ")                                                               '�ɿ���
'            strPatient = strPatient & String(6, " ")                                                                '�ɿ�����ҽԺ
'        End If
        If rsCard!�ɿ����� = "9" Then
            strPatient = strPatient & "1"                                                                           '��������
            strPatient = strPatient & String(20, " ")                                                               '�ɿ���
            strPatient = strPatient & String(6, " ")                                                                '�ɿ�����ҽԺ
        Else
            If Not IsNull(rsCard!�¿���) Or rsPatient!�������� Like "���*" Then
                strPatient = strPatient & "0"                                                                           '��������
            Else
                strPatient = strPatient & "1"                                                                           '��������
            End If
            '����Ǿɾ��￨,����Ҫ���Ͼɿ���,�����������д
            If rsCard!�ɿ����� = 2 And rsPatient!�������� = "���￨" Then
                strPatient = strPatient & Rpad(Trim(rsCard!�ɿ���), 20)                                                       '�ɿ���
                strPatient = strPatient & rsCard!�ɿ�����ҽԺ                                                           '�ɿ�����ҽԺ
            Else
                strPatient = strPatient & String(20, " ")                                                               '�ɿ���
                strPatient = strPatient & String(6, " ")                                                                '�ɿ�����ҽԺ
            End If
        End If
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        strPatient = strPatient & String(20, " ")                                                                   '�ɿ���
        strPatient = strPatient & String(6, " ")                                                                    '�ɿ�����ҽԺ
        
        If rsPatient!�������� Like "���*" Or _
            (rsPatient!�������� = "ͨ�þ��￨" And Nvl(rsCard!�ɿ�����ҽԺ) = gstrҽԺ����) Or _
            rsPatient!�������� = "���￨" Then '�½�����������ɿ����¿�
            strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '����ʱ��
        Else    '��������������˵Ľ���ʱ��
            strPatient = strPatient & Nvl(rsPatient!һ��ͨ����ʱ��, Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss"))                                                          '����ʱ��
        End If
        
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '����ʱ��
        clsCardDev_UploadSwap = ���ýӿ�("putPersonInfo", strPatient)
    End If
    '-----------------------------------------------
    
    '���²��˷�����¼���ϴ���־
    If clsCardDev_UploadSwap Then
        gcnConnect.Execute "zl_���˷�����¼_�ϴ�(" & lng����ID & ")", , adCmdStoredProc
    End If
    
    clsCardDev_UploadSwap = True
    Exit Function
errHand:
    gLastErr = Err.Description
End Function

Private Function ���ýӿ�(ByVal strFunction As String, ByVal strCreate As String) As Boolean
'    ----------------------------------------------------------------
    '��������   �����ýӿں���
    '��д��     ������
'    ��д����   ��2009-07-31
'    ----------------------------------------------------------------
    Dim strURL As String, strSoapRequest As String
    Dim str���� As String, lng���к� As Long, str������Ϣ As String
    Dim objHttp As MSXML2.XMLHTTP
    On Error GoTo errHand
    
    Set objHttp = New MSXML2.XMLHTTP
    strURL = gstr������ַ & "?op=" & strFunction
    
    strSoapRequest = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>" & _
                "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">" & _
                "<soapenv:Header>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:user>" & gstr�û��� & "</ns:user>" & _
                        "<ns:pwd>" & gstr���� & "</ns:pwd>" & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Header>" & _
                "<soapenv:Body>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:FieldStr>" & strCreate & "</ns:FieldStr>" & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Body>" & _
                "</soapenv:Envelope>"
       
    If gbln��Ϣת�� = False Then
        objHttp.Open "post", strURL, False
        objHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        objHttp.setRequestHeader "Content-Length", Len(strSoapRequest)
        objHttp.setRequestHeader "SOAPAction", strURL
        
        '���ݷ��ص�״̬��Ϣ���ж��Ƿ�ɹ�
        objHttp.send (strSoapRequest)
        If objHttp.status <> 200 Then
            MsgBox "������Ϣ��[" & objHttp.status & "]" & objHttp.responseText
            Exit Function
        End If
    Else
        'д������
        If Not SendRequest(str����, lng���к�, strFunction, strURL, strSoapRequest) Then Exit Function
        
        '��ʾ�ȴ�����
        If frmWait.SendRequest(str����, lng���к�, str������Ϣ) = False Then
            If str������Ϣ <> "" Then MsgBox str������Ϣ
            Exit Function
        End If
    End If
    
    '�ϵ����ô�
    Set mdomOutput = New MSXML2.DOMDocument
    If gbln��Ϣת�� = False Then
        If mdomOutput.loadXML(objHttp.responseText) = False Then
            MsgBox "���׺�����" & strFunction & "���������ݸ�ʽ����ȷ��"
            Exit Function
        End If
    Else
        If mdomOutput.loadXML(str������Ϣ) = False Then
            MsgBox "���׺�����" & strFunction & "���������ݸ�ʽ����ȷ��"
            Exit Function
        End If
    End If
    
    ���ýӿ� = True
    Exit Function
errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Function SendRequest(str���� As String, lng���к� As Long, _
    ByVal strFuncName As String, ByVal strURL As String, ByVal strSoapRequest As String) As Boolean
    Dim blnTrans As Boolean
    Dim strRow As String
    Dim intRow As Integer, intCount As Integer
    On Error GoTo errHand
    '������������д�����ݱ�
    
    str���� = Format(gobjDatabase.CurrentDate, "yyyyMMdd")
    lng���к� = gobjDatabase.GetNextId("��Ϣת��")
    
    gcnConnect.BeginTrans
    blnTrans = True
    
    '��������
    gcnConnect.Execute "zl_��Ϣ����_Insert('" & str���� & "'," & lng���к� & ",'" & strFuncName & "','" & strURL & "')", , adCmdStoredProc
    
    '�������������
    intCount = Len(strSoapRequest) \ 1000
    If Len(strSoapRequest) Mod 1000 <> 0 Then intCount = intCount + 1
    For intRow = 0 To intCount
        strRow = Mid(strSoapRequest, intRow * 1000 + 1, 1000)
        gcnConnect.Execute "zl_��Ϣת��_Insert('" & str���� & "'," & lng���к� & "," & intRow + 1 & ",'" & strRow & "')", , adCmdStoredProc
    Next
    
    gcnConnect.CommitTrans
    blnTrans = False
    SendRequest = True
    Exit Function
errHand:
    If blnTrans Then gcnConnect.RollbackTrans
    MsgBox Err.Description
End Function
