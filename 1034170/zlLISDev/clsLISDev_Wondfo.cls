VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Wondfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�ɲ�II
'����:���ڣ�wondfo��
'����:�����
'�û�:��֦����Ժ
'����ʱ�䣺2014��2��14��

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
    '   ͼ�����ݷ��ڼ��������ûس����з�(vbcrlf)�ָ�
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim strLine As String, str��Ŀ As String, str�����˱걾�ʿ� As String
    Dim lngLength As Long
    Dim blnEnd As Boolean '�Ƿ��ѽ�����

    On Error GoTo errHandle
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "*FF*EE*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(strReserved, "EE") - InStr(strReserved, "FF") + 2
            strLine = Mid(strReserved, InStr(strReserved, "FF"), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, "EE") + 2)
            
            str���� = "": str�걾�� = "":  strItem = "": str�����˱걾�ʿ� = ""
            '����
            strItem = ""
            If strLine Like "FF*EE" Then
                varLine = Split(strLine, "&")
                If UBound(varLine) = 6 Then
                    str�걾�� = Val(varLine(1))
                    str���� = Format(varLine(4), "yyyy-MM-dd hh:mm:ss")
                    str�����˱걾�ʿ� = " |ѪҺ|0"
                    strItem = strItem & "|CRP|" & Trim(varLine(2)) & "|hsCRP|" & Trim(varLine(3))
                Else
                    str�걾�� = Val(varLine(1))
                    str���� = Format(varLine(3), "yyyy-MM-dd hh:mm:ss")
                    str�����˱걾�ʿ� = " |ѪҺ|0"
                    strItem = strItem & "|PCT|" & Trim(varLine(2))
                End If
            End If
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            
            If Not strReserved Like "*FF*EE*" Then
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
errHandle:
'    WriteLog "clsLISDev_Wondfo.Analyse",  , Err.Number & " " & Err.Description
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



