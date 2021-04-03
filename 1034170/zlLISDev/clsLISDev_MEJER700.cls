VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MEJER700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:MEJER700��Һ����
'����:����Mejer��˾
'����:�µ���
'�û�:��Ԫ����ҽԺ

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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    '�������������д�������������ע�����벻Ҫɾ����
    Dim str_ʱ�� As String, str_������ As String, str_��Ŀ As String, str_��� As String, str_ͨ���� As String
    Dim lon_str As Long, lon_end As Long
    Dim str_Data As String, strLog As String
    Dim varLine As Variant
    Dim i As Integer, str_Line As String, ITYP As Integer
    Dim str_���� As String
    Dim str_sj As String
    Dim strIniFile As String
    Dim objStream As Scripting.TextStream
    Dim objFileSystem As New Scripting.FileSystemObject
    
    strResult = ""
    
    strIniFile = App.Path & "\MEJER700.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "ITYP=0 ;0=�����������Խ�� 1=ֻ�ж��Խ��"
       objStream.Close
       Set objStream = Nothing
    End If
    ITYP = Val(Replace(ReadIni("Options", "ITYP", strIniFile, 0), ";0=�����������Խ�� 1=ֻ�ж��Խ��", ""))
    
    lon_str = InStr(strSource, Chr(&H2))
    lon_end = InStr(strSource, Chr(&H3))
    If lon_str > 0 And lon_end > 0 Then
        str_Data = Mid(strSource, lon_str, lon_end)
        strReserved = Mid(strSource, lon_end + 1)
        str_��Ŀ = "": str_ʱ�� = "": str_ͨ���� = "": str_��� = ""
        str_Data = Replace(str_Data, Chr(&H2), "")
        str_Data = Replace(str_Data, Chr(&H3), "")
        varLine = Split(str_Data, Chr(&HD) & Chr(&HA))
        For i = 1 To UBound(varLine)
            str_Line = varLine(i)
            If InStr(str_Data, "#") > 0 Then
              If i = 1 Then
                 str_������ = Val(Mid(str_Line, InStr(str_Line, "#") + 1, 4))
                 str_���� = Trim(Mid(str_Line, InStr(str_Line, "#") + 8))
              End If
              If i = 2 Then
                 str_sj = Trim(Mid(str_Line, InStr(str_Line, "#") + 8))
                 str_ʱ�� = str_sj & " " & str_����
              End If
            End If
            If i >= 3 And i <= 13 Then
                  '�滻��λΪ��ֵ
                  str_Line = Replace(str_Line, "mmol/L", "")
                  str_Line = Replace(str_Line, "Cell/uL", "")
                  str_Line = Replace(str_Line, "g/L", "")
                  str_Line = Replace(str_Line, "leu/uL", "")
                  str_Line = Replace(str_Line, "ery/uL", "")
                  str_Line = Replace(str_Line, "umol/L", "")
                  str_Line = Trim(str_Line)
                  
                  str_ͨ���� = Trim(Mid(str_Line, 1, 4))
                  str_ͨ���� = Trim(Replace(str_ͨ����, "*", ""))
                  If InStr(str_Data, "#") > 0 Then
                     If ITYP = 0 Then
                       str_��� = Trim(Mid(str_Line, 5))
                     Else
                       If InStr(Trim(Mid(str_Line, 5)), "+") > 0 Or InStr(Trim(Mid(str_Line, 5)), "-") > 0 Then
                         str_��� = Trim(Mid(str_Line, 5, 5))
                       Else
                         str_��� = Trim(Mid(str_Line, 5))
                       End If
                     End If
                  Else
                    If (UCase(str_ͨ����) = "URO") Or (UCase(str_ͨ����) = "VC") Then
                      str_��� = Trim(Mid(str_Line, 5, 6))
                    Else
                      str_��� = Trim(Mid(str_Line, 5, 9))
                    End If
                  End If
                  
                  str_��� = Replace(str_���, "Brown Yellow", "�ػ�ɫ")
                  str_��� = Replace(str_���, "Thin Yellow", "����ɫ")
                  str_��� = Replace(str_���, "Yellow", "��ɫ")
                  str_��� = Replace(str_���, "Green", "��ɫ")
                  str_��� = Replace(str_���, "Primrose", "����ɫ")
                  str_��� = Replace(str_���, "Red", "��ɫ")
                  str_��� = Replace(str_���, "Orange", "��ɫ")
                  str_��� = Replace(str_���, "Blue", "��ɫ")
                  str_��� = Replace(str_���, "Brown", "��ɫ")
                  str_��� = Replace(str_���, "Unselected", "δѡ")
                  str_��� = Replace(str_���, "Other", "����")
                  str_��� = Replace(str_���, "Filemot", "�ƺ�ɫ")
                  str_��� = Replace(str_���, "Little Dark", "��ģ��")
                  str_��� = Replace(str_���, "Dark", "ģ��")
                  str_��� = Replace(str_���, "Clear", "�峺")
                  str_��� = Replace(str_���, "Feculent", "����")
                  str_��� = Replace(str_���, "Negative", "����")
'                  str_��� = Replace(str_���, "Normal", "����")
                  str_��� = Replace(str_���, "Positive", "����")
                  str_��� = Replace(str_���, "Small", "1+")
                  str_��� = Replace(str_���, "Moderate", "2+")
                  str_��� = Replace(str_���, "Large", "3+")
                  str_��� = Replace(str_���, "Trace", "΢��")
                  str_��� = Replace(str_���, "Thin", "ǳ��ɫ")
                  str_��� = Trim(str_���)
                  str_��Ŀ = str_��Ŀ & "|" & str_ͨ���� & "|" & str_���
            End If
        Next
        If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
            strResult = str_ʱ�� & "|" & str_������ & "| |��Һ|0" & str_��Ŀ
        End If
    Else
        strReserved = strSource
    End If
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