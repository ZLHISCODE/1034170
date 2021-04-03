VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ReLIA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�������������
'�豸���̣�δ֪
'�豸�ͺţ�RELIA
'ʹ��ҽԺ:�ɶ����Բ�ҽԺ

Private mstrCode As String

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
    Dim str_S As Long, str_E As Long, str_SS As Long
    Dim str_Code As String, str_Line As String
    Dim VAR_DAT As Variant
    Dim i As Integer, j As Integer
    Dim str���� As String, str�걾�� As String, str��Ŀ As String, str��� As String, str������Ŀ As String, str������ As String
    
    str���� = "": str_S = 0: str_E = 0
    str�걾�� = "": str��Ŀ = "": str��� = ""
    mstrCode = strSource
    '�ж��ַ�����������
    str_S = InStr(mstrCode, "BEGIN REPORT")

    If str_S > 0 Then
       str_E = InStr(mstrCode, "END REPORT")
    Else
       str_E = 0
    End If

    If str_S > 0 And str_E > 0 Then
        Do While str_S > 0 And str_E > 0  '�ж�һ���������ж����ʼ�ͽ����Ŀ�ʼ
             str_Code = Mid(mstrCode, str_S, str_E - str_S + 12)
             mstrCode = Mid(mstrCode, str_E + 15)
              '��������к��������ı걾�Ϳ�ʼ�������ݵĽ���
             VAR_DAT = Split(str_Code, Chr(&HD) + Chr(&HA))
             For i = LBound(VAR_DAT) To UBound(VAR_DAT)
                 str_Line = VAR_DAT(i)
                 If str_Line Like "*Specimen ID:*" Then
                    str�걾�� = Val(Mid(str_Line, InStr(str_Line, ":") + 1))
                ElseIf str_Line Like "*Test finish time:*" Then
                    str���� = Trim(Mid(str_Line, InStr(str_Line, ":") + 1))
                    str���� = Format(CDate(str����), "yyyy-MM-dd HH:MM") & ":" & Format(Now, "SS")
                    If Not IsDate(str����) Then str���� = Format(Now, "yyyy-MM-dd HH:MM:SS")
                 ElseIf str_Line Like "*Test name:*" Then
                    str������Ŀ = Trim(Mid(str_Line, InStr(str_Line, ":") + 1))
                 ElseIf InStr(str_Line, str�걾��) > 0 And InStr(str_Line, str������Ŀ) And InStr(str_Line, ":") > 0 And str������Ŀ <> "" And str�걾�� <> "" Then
                    str������ = Trim(Replace(Replace(Replace(Replace(Mid(str_Line, InStr(str_Line, ":") + 1), "ng/ml", ""), "pg/ml", ""), "��", ""), "��", ""))
                 End If
             Next
             If str�걾�� <> "" And str���� <> "" And str������Ŀ <> "" And str������ <> "" Then
                str��� = IIf(Trim(str���) = "", "", str��� & "||") & IIf(Trim(str������Ŀ & "|" & str������) = "", "", str���� & "|" & str�걾�� & "| |Ѫ��|0|" & str������Ŀ & "|" & str������)
                str���� = ""
                str�걾�� = ""
                str������Ŀ = ""
                str������ = ""
             End If
        'ѭ�������ж�
             str_S = InStr(mstrCode, "BEGIN REPORT")
             If str_S > 0 Then
                str_E = InStr(mstrCode, "END REPORT")
             Else
                str_E = 0
             End If
            
        Loop
    End If
    strResult = str���
    strReserved = mstrCode
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""

End Function

