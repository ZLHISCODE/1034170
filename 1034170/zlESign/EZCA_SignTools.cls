VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EZCA_SignTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private EZCAClientTools As Object
Private CapicomCert As Object
Private SoapClient As Object

Private SoapAction As String
'wsdl��ַ
Private END_POINT_URL As String
Private Const BASE64CHR          As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)     As String

'�����û���ʶ
Public Function WF_SelectCert() As String
    Dim ls_ret, user_id As String
    Dim lstr() As String
    Dim users() As String
    Dim Cert_User As String
    
     '1���ؼ��Ƿ�֧��
    ls_ret = EZCAClientTools.SOF_CheckSupport()
    If ls_ret <> 0 Then Exit Function

    '2ȡ�û��б���ѡ���û�
    ls_ret = EZCAClientTools.SOF_GetUserList()
    If ls_ret = "" Then
        MsgBox "�����գӣ£ˣţ٣�"
        Exit Function
    Else
        If InStr(ls_ret, "&&&") > 0 Then    '����û���
            users = Split(ls_ret, "&&&")
            'ѡ���û�
            Cert_User = frmSelectUser.ShowMe(ls_ret)
        Else
            Cert_User = ls_ret
        End If
    End If
    
    '3�ټ���û��Ƿ���ȷѡ�񣬲�ȡ�û�Ψһ��ʾ
    If Cert_User = "" Then Exit Function

    lstr = Split(Cert_User, "||")
    If UBound(lstr) < 1 Then Exit Function
    If lstr(1) = "" Then Exit Function
    
    user_id = lstr(1)
    
    WF_SelectCert = user_id
End Function

'ǩ��������
Public Function WF_SignData(bussinessid As String, sdata As String, CertID As String) As String
'������bussinessId ҵ��Ψһ��ʶ��sdata ��ǩ��ԭ�ģ�timeStampState �Ƿ���Ҫʱ���(1��Ҫ0����Ҫ����signTypeǩ����������(0:�ַ�����1���ļ�,2:XML)��personid����ID��
'���أ����ͣ�0��ʾ�ɹ�������ʧ��
sdata = EncodeBase64String(sdata)  '���ȶ�ԭ�Ľ���base64����
Const timestampstate = 0
Const signtype = 0
Const personid = ""
 
    Dim li_result As Integer
    Dim ls_ret, signed_data As String
    Dim user_id, cert_id As String
    Dim lstr() As String
    Dim users() As String
    Dim Cert_User As String
    Dim strEdate As String
    
     '1���ؼ��Ƿ�֧��
    ls_ret = EZCAClientTools.SOF_CheckSupport()
    If ls_ret = 0 Then
        'MsgBox "test::" & "֧��"
    Else
        'MsgBox "test::" & "��֧��" ��֧��ǩ���������˳�
        WF_SignData = 1000
        Exit Function
    End If
    '2ȡ�û��б���ѡ���û�
     ls_ret = EZCAClientTools.SOF_GetUserList()
    'MsgBox "test::" & ls_ret
    
    If ls_ret = "" Then
        MsgBox "�����գӣ£ˣţ٣�"
        WF_SignData = 1001
        Exit Function
    Else
        If InStr(ls_ret, "&&&") > 0 Then    '����û���
            users = Split(ls_ret, "&&&")
            'ѡ���û�
            Cert_User = frmSelectUser.ShowMe(ls_ret)
        Else
            Cert_User = ls_ret
        End If
    End If
     '3�ټ���û��Ƿ���ȷѡ�񣬲�ȡ�û�Ψһ��ʾ
    If Cert_User = "" Then
        WF_SignData = 1002
        Exit Function
    Else
        lstr = Split(Cert_User, "||")
        If UBound(lstr) < 1 Then
            WF_SignData = 1003
            Exit Function
        End If
        If lstr(1) <> "" Then
            user_id = lstr(1)
        Else
            WF_SignData = 1004
            Exit Function
        End If
    End If
    
    '�û���Keyƥ���Լ��
    If user_id <> CertID Then
        MsgBox "�գӣ£ˣţ�����Ա������"
        WF_SignData = 1001
        Exit Function
    End If
    
    '2018-4-13 �����޸Ĺ���ʱ������,һ�ε�¼ֻ��Ҫ����һ��
    If Not gblnShow Then
        strEdate = EZCAClientTools.SOF_GetCertInfo(EZCAClientTools.SOF_ExportUserCert(user_id), 18)
        If DateDiff("d", Date, Format(strEdate, "yyyy-mm-dd")) <= 30 Then
            If MsgBox("���֤�黹��" & DateDiff("d", Date, Format(strEdate, "yyyy-mm-dd")) & "�����,�뾡�쵽��Ϣ�Ƹ���!���ε�¼�Ƿ������ѣ�", vbYesNo) = vbYes Then
                gblnShow = True
            End If
        End If
    End If
    'ȡ֤�����к�
    cert_id = get_Cert_Id(user_id)
    'MsgBox cert_id
    '4����ǩ������
     signed_data = EZCAClientTools.SOF_SignData(user_id, sdata)
     If signed_data = "" Then
        MsgBox "ǩ��ʧ��"
        WF_SignData = 1005
        Exit Function
    End If
    'MsgBox signed_data
     '5�������ݱ��淽��
    'WF_DataSave(ҵ��ΨһID��ǩ��ԭ�ģ�ǩ��ֵ���Ƿ���Ҫʱ���(1��Ҫ0����Ҫ����ǩ����������(0:�ַ�����1���ļ�,2:XML)��ǩ����������(û�и���Ĭ��Ϊ��)��ǩ��֤�����кţ��û���ΨһID�ţ�ǩ��ֵ���� ��0:P7,1:P1��������id��
    ls_ret = WF_DataSave(bussinessid, sdata, signed_data, "0", "0", "", cert_id, user_id, "1", Val(personid))
    
    If ls_ret = 10000 Then
        '����ɹ�
        WF_SignData = signed_data
    Else
        Select Case ls_ret
            Case "10001"
                MsgBox "ǩ���洢ʧ��,������:" & ls_ret
            Case "10002"
                MsgBox "ǩ��������֤ʧ��,������:" & ls_ret
            Case "10003"
                MsgBox "֤����Ч,������:" & ls_ret
            Case "10004"
                MsgBox "֤�����,������:" & ls_ret
        End Select
        WF_SignData = ls_ret
    End If
 End Function
 
 'ȡ֤�����к�
 Private Function get_Cert_Id(ByVal userid As String) As String
 '�������û�ID  ���أ��û�֤������к�
   
    Dim pubcert As String
    pubcert = EZCAClientTools.SOF_ExportUserCert(userid)
    get_Cert_Id = EZCAClientTools.SOF_GetCertInfo(pubcert, 2)

 
 End Function
'�������ݱ����WEBSERVICE����
Public Function WF_DataSave(ByVal bussinessid As String, ByVal InData As String, ByVal signed_data As String, ByVal timestampstate As String, ByVal signtype As String, ByVal signMethodType As String, ByVal SignCertSerialNo As String, ByVal appid As String, ByVal signdatatype As String, ByVal personid As String) As String
    WF_DataSave = SoapClient.VD_DataSave(bussinessid, InData, signed_data, timestampstate, signtype, signMethodType, SignCertSerialNo, appid, signdatatype, personid)
End Function
'��֤�洢���ݿ�ĵ���ǩ��
Public Function WF_VerifyServerSigneData(businessId As String, SignCertSerialNo As String, signdatatype As String) As String
'BusinessId ҵ��ϵͳ��ΨһID��,SignCertSerialNoǩ��֤�����к�,signDataTypeǩ��ֵ���� ��0:P7,1:P1��
    WF_VerifyServerSigneData = SoapClient.VD_VerifyServerSigneData(businessId, SignCertSerialNo, signdatatype)

End Function

'���÷���������֤����ǩ��
Public Function WF_VerifySigneDataByServer(ByVal InData As String, ByVal SignData As String, ByVal signMethodType As String, ByVal signdatatype As String, ByVal signCert As String, ByVal signtype As String) As String
'BusinessId ҵ��ϵͳ��ΨһID��,SignCertSerialNoǩ��֤�����к�,signDataTypeǩ��ֵ���� ��0:P7,1:P1��

    WF_VerifySigneDataByServer = SoapClient.VD_VerifySigneData(InData, SignData, signMethodType, signdatatype, signCert, signtype)

End Function

'���ÿͻ��˿ؼ�����ǩ����֤
Public Function WF_VerifySigneData(ByVal cert As String, ByVal InData As String, ByVal sSignValue As String)
    InData = EncodeBase64String(InData)  '���ȶ�ԭ�Ľ���base64����
    WF_VerifySigneData = EZCAClientTools.SOF_VerifySignedData(cert, InData, sSignValue)
End Function

'ͨ��֤�����кŻ��û�Ψһ��ţ���ȡ�û���Ӧ�Ĺ�Կ֤�顣��������һ��ֻ��ʹ������һ��
Public Function WF_GetUserCert(UID As String, IDType As String) As String
'UID ��ű���,IDType�������ͣ�01��֤�����кţ�02���û�ID�ţ�
   WF_GetUserCert = SoapClient.VD_GetUserCert(UID, IDType)
End Function

'������Ҫǩ����ԭ�����ݻ�HASHֵ������Base64�����ǩ��ֵ
Public Function WF_SignService(InData As String, signtype As String) As String
'inDate[IN]  ԭ�����ݻ�HASHֵ�����������תΪBase64���룩
'SignType    ԭ�������ͣ�1��ԭ�ģ�2��HASHֵ
    WF_SignService = SoapClient.VD_SignService(InData, signtype)
End Function

'�������˵�ǩ����֤
Public Function WF_VerifyService(InData As String, SignData As String, certbase64 As String, signtype As String) As String
'inDate ԭ�����ݻ�HASHֵ�����������תΪBase64���룩,SignData ǩ��ֵ,certbase64 ǩ��֤��BASE64����,signtype ǩ��ֵ����P7Ϊ0��P1Ϊ1
    WF_VerifyService = SoapClient.VD_VerifyService(InData, SignData, certbase64, signtype)

End Function

'���Է���
Public Function Of_test(bussinessid As String, sdata As String, timestampstate As String, signtype As String, personid) As Integer

''������bussinessId ҵ��Ψһ��ʶ��sdata ��ǩ��ԭ�ģ�timeStampState �Ƿ���Ҫʱ���(1��Ҫ0����Ҫ����signTypeǩ����������(0:�ַ�����1���ļ�,2:XML)��personid����ID��
''���أ����ͣ�0��ʾ�ɹ�������ʧ��

 
Dim li_result As Integer
Dim sd_p7 As String
Dim GS_CertID As String
Dim cert_id As String
Dim signed_data As String
Dim ls_ret As String

''ֱ�Ӹ�ֵ
GS_CertID = "EZCA@5014990setup"
cert_id = "10d755cfb9ba15023852aabd88887ee9"
'signed_data = "mIyRj3yyQaNZxv/3qhVp8FQQdoTNXAEnmO/IbgwbqSZbGQvsX/IT0cmU2uhH+rMNMhffdHZTRyYcLDMAlcYqSv1B+6ENsqnNgehaSgF9qfI8HC6BpyriPDjUrtxdKtpl3iXKLVnwjAMEbNpb3kVZKIxYbSNNn27hyqX7DfIppg0="
'base64������P1ǩ��ֵ
signed_data = "1mqsQoe/WuZ+ImXU4q5F7UeCAdbgZuAI+HwNm8M9opJ3eXUlQVY8YR0PBZjLv7rUB7/zF/FlZmXMWTazztPViQNyG55SU1KaIiNV8Ggua+bRpYfHoAM7ZnXDzD1H9LyAxvTQ3O+5unmk68dnetf9bagAZrxMbR8ocN7ckRAtu0o="
'sd_p7 = "MIIE1wYJKoZIhvcNAQcCoIIEyDCCBMQCAQExDzANBgkqhkiG9w0BAQUFADAmBgkqhkiG9w0BBwGgGQQXMzQzNDM01tC5+sjLw/GjrNKq09DLqgCgggNQMIIDTDCCArWgAwIBAgIQENdVz7m6FQI4Uqq9iIh+6TANBgkqhkiG9w0BAQUFADBxMQswCQYDVQQGEwJDTjESMBAGA1UECBMJQ0hPTkdRSU5HMTYwNAYDVQQKEy1DaG9uZ3FpbmcgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IENlbnRlciBDTy5MVEQxFjAUBgNVBAMTDUNob25ncWluZyAgQ0EwHhcNMTEwNTE4MDc1MDI1WhcNMTIwNTE3MDc1MDI1WjA+MQswCQYDVQQGEwJDTjEPMA0GA1UECh4GW8xY65AaMQ0wCwYDVQQLHgRfAFPRMQ8wDQYDVQQDHgZbWVwPWh8wgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBAMUDRoCEGMdDVEQDnGsJBZyc/hhoNWNDuvcG8Naq8UZgXNJJ3HUfkWi8Q5FyLWwfh/aiiHnHsq/Ymr3c03oi1ewKxiIwIRzf5CeF0H0CG88CU6yj55p9JD/JCzOjMXjT7yzbIOTFzhoMy" _
& "L6cooDEE7bYYsdXh2i+HUO8bazpm+3pAgMBAAGjggEWMIIBEjBpBggrBgEFBQcBAQRdMFswMAYIKwYBBQUHMAKGJGh0dHA6Ly93d3cuZGZ6eGNhLmNvbS9DQS9jYWlzc3VlLmh0bTAnBggrBgEFBQcwAYYbaHR0cDovL3d3dy5kZnp4Y2EuY29tOjIwNDQzMB8GA1UdIwQYMBaAFHrTXfuu1AQNTkXzgSsiItPP/FIWMAkGA1UdEwQCMAAwLwYDVR0fBCgwJjAkoCKgIIYeaHR0cDovL3d3dy5kZnp4Y2EuY29tL2NybDguY3JsMAsGA1UdDwQEAwIGwDAdBgNVHQ4EFgQUNwzOUFtyh3NIEX5UJQT/2nImNTIwHAYFKlYVAQE" _
& "EEwwRRVpDQUA1MDE0OTkwc2V0dXAwDQYJKoZIhvcNAQEFBQADgYEACh94eDDIJ0fgaf03/sBhI9OFcqnmQmxMuOhQD+/OAH+e9W0LcHKYi6+/ZtOcclWH8IIQezB3dxMcyMHBKY5vwf483mgpgid7sNAk/JZYjlva18xnUyrehiQWXpy49SPhwov41uFfCZEesKt4P7lAa3GGbEgB1p1t8Iulho2xO1ExggEwMIIBLAIBATCBhTBxMQswCQYDVQQGEwJDTjESMBAGA1UECBMJQ0hPTkdRSU5HMTYwNAYDVQQKEy1DaG9uZ3FpbmcgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IENlbnRlciBDTy5MVEQxFjAUBgNVBAMTDUNob25ncWluZyAgQ0ECEBDXVc+5uhUCOFKqvYiIfukwDQYJKoZIhvcNAQEFBQAwDQYJKoZIhvcNAQEBBQAEgYANpinyDfulyuFun00jbViMKFlF3lvabAQDjPBZLcol3mXaKl3crtQ4POIqp4EuHDzyqX0BSlrogc2psg2h+0H9SirGlQAzLBwmR1N2dN8XMg2z+kfo2pTJ0RPyX+wLGVsmqRsMbsjvmCcBXM2EdhBU8GkVqvf/xlmjQbJ8j5GMmA=="
''
'base64������P7ǩ��ֵ
sd_p7 = "MIIE4QYJKoZIhvcNAQcCoIIE0jCCBM4CAQExDzANBgkqhkiG9w0BAQUFADAwBgkqhkiG9w0BBwGgIwQhTXpRek5ETTAxdEM1K3NqTHcvR2pyTktxMDlETHFnPT0AoIIDUDCCA0wwggK1oAMCAQICEBDXVc+5uhUCOFKqvYiIfukwDQYJKoZIhvcNAQEFBQAwcTELMAkGA1UEBhMCQ04xEjAQBgNVBAgTCUNIT05HUUlORzE2MDQGA1UEChMtQ2hvbmdxaW5nIENlcnRpZmljYXRlIEF1dGhvcml0eSBDZW50ZXIgQ08uTFREMRYwFAYDVQQDEw1DaG9uZ3FpbmcgIENBMB4XDTExMDUxODA3NTAyNVoXDTEyMDUxNzA3NTAyNVowPjELMAkGA1UEBhMCQ04xDzANBgNVBAoeBlvMWOuQGjENMAsGA1UECx4EXwBT0TEPMA0GA1UEAx4GW1lcD1ofMIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDFA0aAhBjHQ1REA5xrCQWcnP4YaDVjQ7r3BvDWqvFGYFzSSdx1H5FovEORci1sH4f2ooh5x7Kv2Jq93NN6ItXsCsYiMCEc3+QnhdB9AhvPAlOso+eafSQ/yQszozF40+8s2yDkxc4a" _
& "DMi+nKKAxBO22GLHV4dovh1DvG2s6Zvt6QIDAQABo4IBFjCCARIwaQYIKwYBBQUHAQEEXTBbMDAGCCsGAQUFBzAChiRodHRwOi8vd3d3LmRmenhjYS5jb20vQ0EvY2Fpc3N1ZS5odG0wJwYIKwYBBQUHMAGGG2h0dHA6Ly93d3cuZGZ6eGNhLmNvbToyMDQ0MzAfBgNVHSMEGDAWgBR60137rtQEDU5F84ErIiLTz/xSFjAJBgNVHRMEAjAAMC8GA1UdHwQoMCYwJKAioCCGHmh0dHA6Ly93d3cuZGZ6eGNhLmNvbS9jcmw4LmNybDALBgNVHQ8EBAMCBsAwHQYDVR0OBBYEFDcMzlBbcodzSBF+VCUE/9pyJjU" _
& "yMBwGBSpWFQEBBBMMEUVaQ0FANTAxNDk5MHNldHVwMA0GCSqGSIb3DQEBBQUAA4GBAAofeHgwyCdH4Gn9N/7AYSPThXKp5kJsTLjoUA/vzgB/nvVtC3BymIuvv2bTnHJVh/CCEHswd3cTHMjBwSmOb8H+PN5oKYIne7DQJPyWWI5b2tfMZ1Mq3oYkFl6cuPUj4cKL+NbhXwmRHrCreD+5QGtxhmxIAdadbfCLpYaNsTtRMYIBMDCCASwCAQEwgYUwcTELMAkGA1UEBhMCQ04xEjAQBgNVBAgTCUNIT05HUUlORzE2MDQGA1UEChMtQ2hvbmdxaW5nIENlcnRpZmljYXRlIEF1dGhvcml0eSBDZW50ZXIgQ08uTFREMRYwFAYDVQQDEw1DaG9uZ3FpbmcgIENBAhAQ11XPuboVAjhSqr2IiH7pMA0GCSqGSIb3DQEBBQUAMA0GCSqGSIb3DQEBAQUABIGASrstEJHc3nAoH21MvGYAqG3913pnx+ukebq579zQ9MaAvPRHPczDdWY7A6DHh6XR5msuaPBVIyKaUlNSnhtyA4nV086zNlnMZWZl8RfzvwfUur/LmAUPHWE8VkEldXl3kqI9w5sNfPgI4Gbg1gGCR+1FruLUZSJ+5lq/h0KsatY="

If signed_data = "" Then
    '1���ؼ��Ƿ�֧��
     li_result = EZCAClientTools.SOF_CheckSupport()
     If li_result = 0 Then
          '֧��
     Else
          MsgBox "��֧��ǩ���������˳�"
          Of_test = 1000
     End If
    '2ȡ�û��б���ѡ���û�
    If GS_CertID = "" Then
         ls_ret = EZCAClientTools.SOF_GetUserList()
         If ls_ret = "" Then
              MsgBox "�����գӣ£ˣţ٣�"
              Of_test = 1001
         End If
    End If
    
    '3ȡ֤�����к�
     cert_id = EZCAClientTools.SOF_GetCertInfo(EZCAClientTools.SOF_ExportUserCert(GS_CertID), 2)
    
    
    ' '4����ǩ������
      signed_data = EZCAClientTools.SOF_SignData(GS_CertID, sdata)
      If signed_data = "" Then
          MsgBox "ǩ��ʧ��"
          Of_test = 1005
     End If
End If
'  '5�������ݱ��淽��
' 'WF_DataSave(ҵ��ΨһID��ǩ��ԭ�ģ�ǩ��ֵ���Ƿ���Ҫʱ���(1��Ҫ0����Ҫ����ǩ����������(0:�ַ�����1���ļ�,2:XML)��ǩ����������(û�и���Ĭ��Ϊ��)��ǩ��֤�����кţ��û���ΨһID�ţ�ǩ��ֵ���� ��0:P7,1:P1��������id��

 ls_ret = WF_DataSave(bussinessid, sdata, signed_data, "0", "0", "", cert_id, GS_CertID, "1", personid)
         'WF_DataSave(bussinessid , inData , signed_data , timestampstate , signtype , signMethodType , SignCertSerialNo , ByVal appid , signdatatype , personid )
 If ls_ret = "10000" Then
'    '����ɹ�
      Of_test = 0
 Else
      Select Case ls_ret
            Case "10001"
                MsgBox "ǩ���洢ʧ��,������:" + ls_ret
            Case "10002"
                MsgBox "ǩ��������֤ʧ��,������:" + ls_ret
            Case "10003"
                 MsgBox "֤����Ч,������:" + ls_ret
            Case "10004"
                MsgBox "֤�����,������:" & ls_ret
        End Select
      Of_test = 1006
       MsgBox "������:" + ls_ret
 End If
 
'P7ǩ������
If sd_p7 = "" Then
    sd_p7 = EZCAClientTools.SOF_SignDataByP7(GS_CertID, sdata)
    'Form1.Text1.Text = sd_p7
End If
If signed_data = "" Then
    MsgBox "ǩ��ʧ��"
    Of_test = 1007
End If
'  '5�������ݱ��淽��
' 'WF_DataSave(ҵ��ΨһID��ǩ��ԭ�ģ�ǩ��ֵ���Ƿ���Ҫʱ���(1��Ҫ0����Ҫ����ǩ����������(0:�ַ�����1���ļ�,2:XML)��ǩ����������(û�и���Ĭ��Ϊ��)��ǩ��֤�����кţ��û���ΨһID�ţ�ǩ��ֵ���� ��0:P7,1:P1��������id��

 ls_ret = WF_DataSave(bussinessid, sdata, sd_p7, "0", "0", "", cert_id, GS_CertID, "0", personid)
         'WF_DataSave(bussinessid , inData , signed_data , timestampstate , signtype , signMethodType , SignCertSerialNo , ByVal appid , signdatatype , personid )
 If ls_ret = "10000" Then
'    '����ɹ�
      Of_test = 0
 Else
      Select Case ls_ret
            Case "10001"
                MsgBox "ǩ���洢ʧ��,������:" + ls_ret
            Case "10002"
                MsgBox "ǩ��������֤ʧ��,������:" + ls_ret
            Case "10003"
                 MsgBox "֤����Ч,������:" + ls_ret
            Case "10004"
                MsgBox "֤�����,������:" & ls_ret
        End Select
      Of_test = 1008
       MsgBox "������:" + ls_ret
 End If

'��������ǩ������
signed_data = WF_SignService(sdata, "1")
If signed_data = "" Then
    MsgBox "ǩ��ʧ��"
    Of_test = 1009
End If
'  '5�������ݱ��淽��
' 'WF_DataSave(ҵ��ΨһID��ǩ��ԭ�ģ�ǩ��ֵ���Ƿ���Ҫʱ���(1��Ҫ0����Ҫ����ǩ����������(0:�ַ�����1���ļ�,2:XML)��ǩ����������(û�и���Ĭ��Ϊ��)��ǩ��֤�����кţ��û���ΨһID�ţ�ǩ��ֵ���� ��0:P7,1:P1��������id��

 ls_ret = WF_DataSave(bussinessid, sdata, signed_data, "0", "3", "", "", "", "1", personid)
         'WF_DataSave(bussinessid , inData , signed_data , timestampstate , signtype , signMethodType , SignCertSerialNo , ByVal appid , signdatatype , personid )
 If ls_ret = "10000" Then
'    '����ɹ�
      Of_test = 0
 Else
      Select Case ls_ret
            Case "10001"
                MsgBox "ǩ���洢ʧ��,������:" + ls_ret
            Case "10002"
                MsgBox "ǩ��������֤ʧ��,������:" + ls_ret
            Case "10003"
                 MsgBox "֤����Ч,������:" + ls_ret
            Case "10004"
                MsgBox "֤�����,������:" & ls_ret
        End Select
       Of_test = 1010
       MsgBox "������:" + ls_ret
 End If


 Of_test = 0
 
 
 
End Function
'����webservice��URL��ַ
Public Function Of_SetUrl(ad As String)
    '��ʼ��,��������ַ�ɷ��ڷ��������ݿ��л������ļ��У��Ա����Ժ������������ַ

    'namespace
    SoapAction = ad '"http://127.0.0.1:8090/ezca_signserver/services/ezcawebservice"
    'wsdl��ַ
    END_POINT_URL = ad '+"?WSDL" ' "http://127.0.0.1:8090/ezca_signserver/services/ezcawebservice?WSDL"


End Function
'webservice���Ӷ���ĳ�ʼ��
Public Function Of_init()
    SoapClient.MSSoapInit (SoapAction)
End Function

'ȡ֤��ʹ����
Public Function WF_GetCertOwner(CertID As String) As String
 '������CertId �û�Ψһ��ʶ��
 'itype ȡֵ����
' CERT_SUBJECT    1   �û���
'CERT_UNIQUEID   2   ֤��ʵ��Ψһ��ʶ
'CERT_DEPT   3   ����
'CERT_ISSUE  4   �䷢��DN
'CERT_DEVICETYPE 8   ֤���������
 '���أ��ַ�����֤��ʹ����
Dim itype As Integer
itype = 1
WF_GetCertOwner = EZCAClientTools.SOF_GetUserInfo(CertID, itype)
 
End Function

'ȡ֤��ʹ����
Public Function WF_GetCertDN(CertID As String) As String
 '������CertId �û�Ψһ��ʶ��
 'itype ȡֵ����
'0 x00000001 ֤��汾
'0 x00000002 ֤�����к�
'0 x00000003 ֤��ǩ���㷨��ʶ
'0 x00000004 ֤��䷢�߹���(C)
'0 x00000005 ֤��䷢����֯��(O)
'0 x00000006 ֤��䷢�߲�����(OU)
'0 x00000007 ֤��䷢�����ڵ�ʡ?������?ֱϽ��(S)
'0 x00000008 ֤��䷢��ͨ������(CN)
'0 x00000009 ֤��䷢�����ڵĳ���?����(L)
'0 x00000010 ֤��䷢��Email
'0 x00000011 ֤����Ч��: ��ʼ����
'0 x00000012 ֤����Ч��: ��ֹ����
'0 x00000013 ֤��ӵ���߹���(C)
'0 x00000014 ֤��ӵ������֯��(O)
'0 x00000015 ֤��ӵ���߲�����(OU)
'0 x00000016 ֤��ӵ�������ڵ�ʡ?������?ֱϽ��(S)
'0 x00000017 ֤��ӵ����ͨ������(CN)
'0 x00000018 ֤��ӵ�������ڵĳ���?����(L)
'0 x00000019 ֤��ӵ����Email
'0 x00000020 ֤��䷢��DN
'0 x00000021 ֤��ӵ����DN
'0 x00000022 ֤�鹫Կ��Ϣ
'0 x00000023 ֤����չ����Ϣ
'0 x00000024 �䷢����Կ��ʶ��
'0 x00000025 ֤���������Կ��ʶ��
'0 x00000026 ��Կ��;
'0 x00000027 ˽Կ��Ч��
'0 x00000028 ֤�����
'0 x00000029 ����ӳ��
'0 x00000030 ��������
'0 x00000031 ��������
'0 x00000032 ��չ��Կ��;
'0 x00000033 CRL������
'0 x00000034 netscape����
'0 x00000035 ֤��ʵ��Ψһ��ʶ
'0 x00000036 ��������֤����
'0 x00000037 ������ᱣ�պ�
'0 x00000038 ��ҵ����ע���
'0 x00000039 ��ҵ��֯��������
'0 x00000040 ˰��Ǽ�֤��
 '���أ��ַ�����֤������
Dim itype As Integer
itype = 33
Dim cert_str As String
cert_str = EZCAClientTools.SOF_ExportUserCert(CertID)
WF_GetCertDN = EZCAClientTools.SOF_GetCertInfo(cert_str, itype)
End Function

'ȡ֤�����к�
Public Function WF_GetCertSN(CertID As String) As String
 '������CertId �û�Ψһ��ʶ��
 'itype ȡֵ����
'0 x00000001 ֤��汾
'0 x00000002 ֤�����к�
'0 x00000003 ֤��ǩ���㷨��ʶ
'0 x00000004 ֤��䷢�߹���(C)
'0 x00000005 ֤��䷢����֯��(O)
'0 x00000006 ֤��䷢�߲�����(OU)
'0 x00000007 ֤��䷢�����ڵ�ʡ?������?ֱϽ��(S)
'0 x00000008 ֤��䷢��ͨ������(CN)
'0 x00000009 ֤��䷢�����ڵĳ���?����(L)
'0 x00000010 ֤��䷢��Email
'0 x00000011 ֤����Ч��: ��ʼ����
'0 x00000012 ֤����Ч��: ��ֹ����
'0 x00000013 ֤��ӵ���߹���(C)
'0 x00000014 ֤��ӵ������֯��(O)
'0 x00000015 ֤��ӵ���߲�����(OU)
'0 x00000016 ֤��ӵ�������ڵ�ʡ?������?ֱϽ��(S)
'0 x00000017 ֤��ӵ����ͨ������(CN)
'0 x00000018 ֤��ӵ�������ڵĳ���?����(L)
'0 x00000019 ֤��ӵ����Email
'0 x00000020 ֤��䷢��DN
'0 x00000021 ֤��ӵ����DN
'0 x00000022 ֤�鹫Կ��Ϣ
'0 x00000023 ֤����չ����Ϣ
'0 x00000024 �䷢����Կ��ʶ��
'0 x00000025 ֤���������Կ��ʶ��
'0 x00000026 ��Կ��;
'0 x00000027 ˽Կ��Ч��
'0 x00000028 ֤�����
'0 x00000029 ����ӳ��
'0 x00000030 ��������
'0 x00000031 ��������
'0 x00000032 ��չ��Կ��;
'0 x00000033 CRL������
'0 x00000034 netscape����
'0 x00000035 ֤��ʵ��Ψһ��ʶ
'0 x00000036 ��������֤����
'0 x00000037 ������ᱣ�պ�
'0 x00000038 ��ҵ����ע���
'0 x00000039 ��ҵ��֯��������
'0 x00000040 ˰��Ǽ�֤��
 '���أ��ַ�����֤������
Dim itype As Integer
itype = 2
Dim cert_str As String
cert_str = EZCAClientTools.SOF_ExportUserCert(CertID)
WF_GetCertSN = EZCAClientTools.SOF_GetCertInfo(cert_str, itype)
End Function

'ȡ֤�鹫Կ
Public Function WF_GetPublicCert(CertID As String) As String
'������CertId �û�Ψһ��ʶ��
'���أ��ַ�����֤�鹫Կ
WF_GetPublicCert = EZCAClientTools.SOF_ExportUserCert(CertID)
End Function

'��ʾ֤����Ϣ
Public Function WF_ShowCert(CertID As String) As String
'������CertId �û�Ψһ��ʶ��
Dim CertContent As String
CertContent = WF_GetPublicCert(CertID)
If CertContent <> "" Then
    CapicomCert.Import (CertContent)
    CapicomCert.Display
End If

End Function
 
Private Sub Class_Initialize()
Set SoapClient = CreateObject("MSSOAP.SoapClient30")  'SOAP���Ӷ��󣬵�WEBSERVICE��
Set CapicomCert = CreateObject("CAPICOM.Certificate")   '΢��CAPICOM֤�������ʾ֤����Ϣ��
Set EZCAClientTools = CreateObject("SANITATIONSYSTEMCLIENT.EZCASanitationSystemClient")  '������Ѷ�����ͻ��˿ؼ����ͻ���ǩ����ǩ��

'webservice�����ʼ��
Dim strURL As String
'strUrl = GetSetting("ZLSOFT", "ע����Ϣ", "CQCA_URL", "")
strURL = ReadIni("CA", "URL", App.Path & "\TSA.ini")
'strUrl = "http://192.168.4.40:8090/ezca_signserver/services/ezcawebservice?wsdl"
If strURL = "" Then
    Err.Raise -1, , "û������ǩ����������ַ���������á�"
    Exit Sub
End If
SoapAction = strURL

SoapClient.MSSoapInit (SoapAction)
End Sub