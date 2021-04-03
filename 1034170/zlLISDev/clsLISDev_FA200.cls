VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FA200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:FA-200 ��Һ������,FA-300
'����:ɽ����ɭ
'����:�¶�
'�û�:�»�

'�޸���:���
'�޸�����:2009-08-25

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
    Dim str_ʱ�� As String, str_������ As String, str_��Ŀ As String, str_��� As String
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String
    Dim varLine As Variant
    Dim i As Integer, str_Line As String
    
    lng_S = InStr(strSource, Chr(&H2))
    lng_E = InStr(strSource, Chr(&H3))
    strResult = ""
    If lng_S > 0 And lng_E > 0 Then
        str_Data = Mid(strSource, lng_S, lng_E - lng_S + 1)
        strReserved = Mid(strSource, lng_E + 1)
        str_��Ŀ = ""
        str_ʱ�� = ""
        varLine = Split(str_Data, Chr(&HD) & Chr(&HA))
        For i = LBound(varLine) To UBound(varLine)
            str_Line = varLine(i)
            If Left(str_Line, 1) = Chr(&H2) Then str_Line = Mid(str_Line, 2)
            If Left(str_Line, 1) = Chr(&HA) Then str_Line = Mid(str_Line, 2)
            If Left(str_Line, 1) = Chr(&HD) Then str_Line = Mid(str_Line, 2)
            If UCase(str_Line) Like "*NO.*" Then
                str_Line = UCase(str_Line)
                str_������ = Val(Mid(str_Line, InStr(str_Line, "NO.") + 7))
                
            ElseIf str_Line Like "####/##/##*" Then
                str_ʱ�� = Format(CDate(str_Line), "yyyy-MM-dd HH:MM")
                
            ElseIf Not (UCase(str_Line) Like "*:*") Then
                    str_Tmp = Trim(Mid(str_Line, 6))
                    str_Tmp = Replace(UCase(str_Tmp), "MMOL/L", "")
                    str_Tmp = Replace(UCase(str_Tmp), "CELLS/UL", "")
                    str_Tmp = Replace(UCase(str_Tmp), "CELL/UL", "")
                    str_Tmp = Replace(UCase(str_Tmp), "UMOL/L", "")
                    str_Tmp = Replace(UCase(str_Tmp), "MG/L", "")
                    str_Tmp = Replace(UCase(str_Tmp), "LEU/UL", "")
                    str_Tmp = Replace(UCase(str_Tmp), "G/L", "")
'                    If str_Line Like "*+*" Then
'                        str_Tmp = Trim(Mid(str_Line, 6, InStr(1, Mid(str_Line, 6), "+")))
'                    End If
                    If str_Tmp <> "" And Trim(Mid(str_Line, 2, 4)) <> "" Then
                        str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 2, 4)) & "|" & Trim(str_Tmp)
                    End If
            End If
            
        Next
        If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
            strResult = str_ʱ�� & "|" & str_������ & "|" & " |" & "��Һ" & "|0" & str_��Ŀ
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

