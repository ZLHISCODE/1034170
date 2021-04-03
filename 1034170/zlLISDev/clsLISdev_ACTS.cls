VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_ACTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ACTS���߱����챨�浥
'����:????
'����:����
'�û�:����һ����ҽԺ
'ʱ��:20100508
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    Dim str_Code As String, lng_S As Long, lng_len As Long
    Dim str_Day As String, str_���� As String, str�걾�� As String, strValues As String
    Dim int_i As Integer, strRecord As String
    Dim str_WBC As String, str_RBC As String, str_PLT As String
    Dim str_Data As String
    Dim strNow As String
    Dim strLine() As String
    Dim intloop As Integer
    Dim strTmp As String
    Dim strItem() As String
    Dim strData As String
    Dim intItem As Integer
    Dim strImage As String
    Dim strPath As String
    
    
    '����
'    If InStr(strSource, "gooyonetSystemlogin") > 0 Then
'        '����ȷ��
'        strCmd = "gooyonetSystemlogin�����              1469                "
'        Exit Sub
'    End If
    
    If InStr(strSource, Chr(&H2)) <= 0 Or InStr(strSource, Chr(&H3)) <= 0 Then
        strReserved = strSource
        Exit Sub
    End If
    strNow = Mid$(strSource, InStr(strSource, Chr(&H2)), InStr(strSource, Chr(&H3)))
    strReserved = Mid$(strSource, InStr(strSource, Chr(&H3)) + 1)
    
    strPath = ReadIni("SETUP", "PATH", App.Path & "\ACTS.ini")
    If strPath = "" Then strPath = "D:\ACTS\DataFile\":   Call WriteIni("SETUP", "PATH", "D:\ACTS\DataFile\", App.Path & "\ACTS.ini")
    
    '��ʼ����
    strLine = Split(strNow, vbCrLf)
    For intloop = 0 To UBound(strLine)
        'ȥ���س�����
        strTmp = Replace$(strLine(intloop), vbLf, "")
        strTmp = Replace$(strTmp, vbCr, "")
        strTmp = Replace$(strTmp, Chr(&H2), "")
        strTmp = Replace$(strTmp, Chr(&H3), "")
        
        If InStr(strTmp, "|") > 0 Then
            strItem = Split(strTmp, "|")
            '�ж��Ƿ�����ֵ�������ֵ���Ǳ걾��
            If IsNumeric(strItem(2)) = True Then
                strData = strItem(3) & " " & Format(Now, "hh:mm:ss") & "|" & strItem(2) & "| |��|0"
            End If
            
            '������
            If Mid(strItem(2), 1, 2) = "YR" Then
                For intItem = 3 To UBound(strItem) Step 5
                    strData = strData & "|" & strItem(intItem + 1) & "|" & strItem(intItem + 3)
                Next
            End If
            
            '������
            If Mid(strItem(2), 1, 2) = "ZR" Then
                strPath = strPath & Mid(strItem(3), 1, 4) & "\" & Val(Mid(strItem(3), 5, 2)) & "\" & Val(Mid(strItem(3), 7, 2)) & "\"
                strImage = " ;110;1;" & strPath & Mid(strItem(3), 1, InStr(strItem(3), ".")) & "jpg"
                strImage = strImage & "^" & "  ;110;1;" & strPath & Mid(strItem(4), 1, InStr(strItem(4), ".")) & "jpg"
            End If
        End If
    Next
    strResult = strData & vbCrLf & strImage
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function



Private Function Graph_Text(ByVal str_bin As String) As String
    '����ͼ���ı�����
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String
    lng_len = Len_Bin(str_bin)

    If lng_len < 1 Then
        Graph_Text = ""
        Exit Function
    End If

    For lng_i = 1 To lng_len
        str_bit = Replace(Mid_bin(str_bin, lng_i, 1, False), ",", "")
    
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format(CByte("&H" & str_bit), "000")
        End If
    Next
    Graph_Text = str_Return
End Function

