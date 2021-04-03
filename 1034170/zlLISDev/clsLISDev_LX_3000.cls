VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LX_3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:LX3000
'����:
'����:����
'�û�:��������ҽԺ

Private mstrCode As String    '��������
Private mlngStep_In As Long   '�����ַ�����������
Private mstrACK As String     '�Ƿ�Ӧ��ACK,����Ӧ��ACK(06),��-Ӧ�� ETX(03)

Private mlngStep_Out As Long  '��˫����Ϣ����
Private mstr������ As String

Private mstrWaitSend As String '�����ͱ걾


Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long, i_Count As Integer
    Dim str_in As String, str_Line As String, str_���� As String
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
    Dim str_WaitSendLine As String
    Dim strNow As String
    Dim strData As String
    Dim strImage As String
    Dim intLoop As Integer
    Dim aLine() As String
    Dim aItem() As String
    Dim str�걾�� As String
    Dim strImagePath As String
    Dim strLine As String
    Dim strFilePaht As String
    Dim strValue As String
    Dim strItem As String
    Dim str���� As String

    mstrCode = mstrCode & strSource
    If InStr(mstrCode, Chr(&H38) & Chr(&H0) & Chr(&HCCCC) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0)) > 0 Then
        strCmd = "OK"
    End If
    
    '��ʼ��
    If InStr(mstrCode, Chr(&H9) & Chr(&H0) & Chr(&H0) & Chr(&H0)) <= 0 Then
        Exit Sub
    End If
    
    '������
    If InStr(mstrCode, Chr(&H2E) & Chr(&H6A) & Chr(&H70) & Chr(&H67)) <= 0 Then
        Exit Sub
    End If
    
    strImagePath = ReadIni("LX_3000", "ͼ��·��", App.Path & "\LX_3000.ini")
    If strImagePath = "" Then strImagePath = "c:\LX3000\":   Call WriteIni("LX_3000", "ͼ��·��", strImagePath, App.Path & "\LX_3000.ini")
    
    
    '�����������
    '����ת��Ϊ�ַ���ʽ
'    strData = Mid(strSource, 1, Len(strSource) - 1)
    lng_E = InStr(mstrCode, Chr(&H2E) & Chr(&H6A) & Chr(&H70) & Chr(&H67)) + 3
    lng_S = InStr(mstrCode, Chr(&H9) & Chr(&H0) & Chr(&H0) & Chr(&H0))
    
    strData = Mid(mstrCode, lng_S, lng_E - lng_S)
    mstrCode = Mid$(mstrCode, lng_E + 1)
    
    aLine = Split(strData, vbCrLf)
    For intLoop = 0 To UBound(aLine)
        If InStr(aLine(intLoop), "EXAMINEDATE:") > 0 And UBound(Split(aLine(intLoop), ":")) > 5 Then
            aItem = Split(aLine(intLoop), ":")
            
            
            strItem = Trim(Replace(aItem(5), "ITEMVALUE", ""))
            If strItem <> "" Then
                strResult = strResult & "||" & Trim(Mid(Trim(aItem(1)), 1, 10)) & "|" & Trim(Mid(Trim(aItem(3)), 1, 5)) & "| |��|0|"
                strResult = strResult & Trim(Replace(aItem(5), "ITEMVALUE", "")) & "|"
                
                If Trim(Replace(aItem(7), "ITEMDESC", "")) = "" Then
                    strValue = Trim(Replace(aItem(6), "ITEMCHAR", " "))
                Else
                    strValue = Trim(Replace(aItem(7), "ITEMDESC", " "))
                End If
                If strValue = "" Then strValue = " "
                strResult = strResult & strValue
                
'                strImage = Trim(Replace(aItem(9), "DELETEMARK", "")) & ".jpg"
                str�걾�� = Trim(Mid(Trim(aItem(3)), 1, 5))
                str���� = Trim(Mid(Trim(aItem(1)), 1, 10))
            End If
        End If
    Next
    strResult = Mid(strResult, 3)

    If str�걾�� <> "" Then
        '����ͼ��
        strImage = Mid(strData, InStr(strData, ".jpg") - 20, 25)
        strImage = Mid(strImage, InStr(strImage, "2"))
        strImage = Mid(strImage, 1, InStr(strImage, "g"))
        strImage = strImagePath & Format(str����, "yyyymm") & "\" & Format(str����, "yyyymmdd") & "\" & Trim(strImage)
    End If

    If strImage <> "" Then
        strResult = strResult & vbCrLf & "A;110;1;" & strImage
    End If
    
    strReserved = ""
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep����ǰ���Ͳ����Ĳ��衣�����������ͽ������򷵻�0�������Զ���1���ء�
    'strResponse����ǰ��������Ӧ����
    'strSampleInfo�����͵ı걾��Ϣ��
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5�����أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��

    Dim var_�걾 As Variant, i_Count As Integer, i_Loop As Integer, str_��¼ As String, str_�걾�� As String
    Dim i_���� As Integer, str_���� As String, str_������Ŀ As String, str_�걾���� As String
    Dim str_SendItem As String
    
    
    
End Function


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function

Private Function SaveBase64Bmp(ByVal str�걾�� As String, strCode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
    
    Dim ByteData() As Byte, lngCount As Long, strfilename As String, lngFileNum
    Dim aTmp() As String
    
    aTmp = Split(strCode, ",")
    
    ReDim ByteData(UBound(aTmp))
    
    For lngCount = 0 To UBound(aTmp) - 1
        ByteData(lngCount) = "&H" & aTmp(lngCount)
    Next
        
    
    
    strfilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_1" & ".jpg"
    lngFileNum = FreeFile
    Open strfilename For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBase64Bmp = strfilename

End Function