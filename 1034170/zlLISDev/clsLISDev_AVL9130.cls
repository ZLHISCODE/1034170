VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AVL9130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AVL 9130
'����:����
'����:�¶�
'�û�:��ú����ҽԺ

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
    Dim str_Code As String, lng_S As Long, lng_E As Long
    Dim var_line As Variant, str_Line As String, int_Loop As Integer
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String
    
    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    
    Do While lng_S > 0 And lng_E > 0
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        
        var_line = Split(str_Code, Chr(&HD) & Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_line) To UBound(var_line)
            str_Line = var_line(int_Loop)
            If str_Line Like "*AVL 9130*" Then
                str_�걾�� = "": str_���� = "": str_��Ŀ = ""
                str_���� = var_line(int_Loop + 2)
                
                str_���� = "20" & Mid(str_����, 6, 2) & "-" & TrunMonth(Mid(str_����, 3, 3)) & "-" & Left(str_����, 2)
                If Not IsDate(str_����) Then str_���� = Format(Now, "yyyy-MM-dd")
                
                str_�걾�� = var_line(int_Loop + 8)
                If str_�걾�� Like "Sample No.*" Then
                    str_�걾�� = Val(Replace(str_�걾��, "Sample No.", ""))
                    If Val(str_�걾��) <> 0 Then
                        
                        strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0"
                        str_��Ŀ = var_line(int_Loop + 11)
                        strResult = strResult & "|" & Trim(Split(str_��Ŀ, "=")(0)) & "|" & Val(Trim(Replace(Replace(Split(str_��Ŀ, "=")(1), "mmol/L", ""), "", "")))
                        str_��Ŀ = var_line(int_Loop + 12)
                        strResult = strResult & "|" & Trim(Split(str_��Ŀ, "=")(0)) & "|" & Val(Trim(Replace(Replace(Split(str_��Ŀ, "=")(1), "mmol/L", ""), "", "")))
                        str_��Ŀ = var_line(int_Loop + 13)
                        strResult = strResult & "|" & Trim(Split(str_��Ŀ, "=")(0)) & "|" & Val(Trim(Replace(Replace(Split(str_��Ŀ, "=")(1), "mmol/L", ""), "", "")))
                    End If
                End If
            End If
        Next
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
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
    
    SendSample = ""
    
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function
