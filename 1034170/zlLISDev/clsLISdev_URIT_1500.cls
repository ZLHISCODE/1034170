VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_URIT_1500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:URIT_1500 ��Һ����
'����:�����ص��Ӽ������޹�˾
'����:�����
'�û�:����������ҽԺ

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
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
    Dim lon_str As Long, lon_end As Long
    Dim str_Data As String
    Dim varLine As Variant
    Dim i As Integer, str_Line As String
    
    On Error GoTo errHand
    lon_str = InStr(strSource, Chr(&H2))
    lon_end = InStr(strSource, Chr(&H3))
    If lon_str > 0 And lon_end > 0 Then
        str_Data = Mid(strSource, lon_str, lon_end)
        strReserved = Mid(strSource, lon_end + 1)
        str_��Ŀ = ""
        str_ʱ�� = ""
        varLine = Split(str_Data, Chr(&HD) & Chr(&HA))
        str_������ = Val(Mid(varLine(1), InStr(varLine(1), "NO.") + 3, 4))
        str_ʱ�� = Format(Trim(Mid(varLine(1), 12, 10)) & " " & Trim(Mid(varLine(2), 13, 9)), "yyyy-mm-dd hh:mm:ss")
        For i = 4 To UBound(varLine)
            If varLine(i) <> "" Then
                str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(varLine(i), 2, 4)) & "|" & IIf(Trim(Mid(varLine(i), 6, 7)) = "", Trim(Mid(varLine(i), 13, 9)), Trim(Mid(varLine(i), 6, 6)))
            End If
        Next
        If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
            str_��Ŀ = Replace(str_��Ŀ, "Cell/uL", " ")
            str_��Ŀ = Replace(str_��Ŀ, "mmol/L", " ")
            str_��Ŀ = Replace(str_��Ŀ, "umol/L", " ")
            str_��Ŀ = Replace(str_��Ŀ, "g/L", " ")
            strResult = str_ʱ�� & "|" & str_������ & "|" & " |��Һ" & "|0" & str_��Ŀ
        End If
    Else
        strReserved = strSource
    End If
    
    Exit Sub
    
errHand:
    WriteLog "clsLISdev_URIT_1500", "�����" & Err.Number & "��" & "�������� ��" & Err.Description, ""
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function





