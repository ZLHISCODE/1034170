VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_FUS200_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��������: FUS200ȫ�Զ������������
'�����ͺ�: FUS200
'����:��������ҽ����е���޹�˾
'����:����
'����:2012-10-17



Private mstrCode As String    '������ܵ�������


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


    Dim lng_S As Long                                   '��ʼ��λ��
    Dim lng_E As Long                                   '������λ��
    Dim strDate As String                               '��������ʱ��
    Dim strNo As String                                 '�������ı걾��
    Dim str_Items As String                             '��������ͨ���뼰���
    Dim str_Sample As String
    Dim var_Tmp As Variant
    Dim intloop As Integer
    Dim strLine As String
    
    '������ܵ�������
    mstrCode = mstrCode & strSource

    '�ж��Ƿ��յ���һ������������
    lng_S = InStr(mstrCode, Chr(&HB))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H1C))
    Else
        lng_E = InStr(mstrCode, Chr(&H1C))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        '�յ����������ݣ�ȡ��
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        'ȡ�����ڣ��걾�ţ�������
        var_Tmp = Split(str_Sample, Chr(&HD))
        For intloop = LBound(var_Tmp) To UBound(var_Tmp)
            strLine = var_Tmp(intloop)
            If InStr(strLine, Chr(&HB)) > 0 Then
                strDate = Split(strLine, "|")(6)
                strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                strDate = Format(strDate, "YYYY-MM-DD hh:mm:ss")
            ElseIf InStr(strLine, "PID") > 0 Then
                strNo = Split(strLine, "|")(3)
            ElseIf InStr(strLine, "OBX") > 0 And InStr(strLine, "NM") > 0 And InStr(strLine, "Chemistry") > 0 Then
                If Split(strLine, "|")(3) = "WBC" Then
                    str_Items = str_Items & "|" & Split(strLine, "|")(3) & "S|" & Split(Split(strLine, "|")(5), "^")(1)
                Else
                    If Split(Split(strLine, "|")(5), "^")(1) <> "" Then
                        str_Items = str_Items & "|" & Split(strLine, "|")(3) & "|" & Split(Split(strLine, "|")(5), "^")(1)
                    Else
                        str_Items = str_Items & "|" & Split(strLine, "|")(3) & "|" & Split(Split(strLine, "|")(5), "^")(2)
                    End If
                End If
            ElseIf InStr(strLine, "OBX") > 0 And InStr(strLine, "NM") > 0 And InStr(strLine, "Sediment") > 0 Then
                str_Items = str_Items & "|" & Split(strLine, "|")(3) & "|" & Split(strLine, "|")(5)
            End If
        Next
        '����ȡ���Ľ��
        If strNo <> 0 And str_Items <> "" Then
            If strResult <> "" Then
                strResult = strResult & "||" & strDate & "|" & strNo & "| |��Һ|0" & str_Items
            Else
                strResult = strDate & "|" & strNo & "| |��Һ|0" & str_Items
            End If
        End If
        '�жϻ����е������Ƿ�ȡ��
        lng_S = InStr(mstrCode, Chr(&HB))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H1C))
        Else
            lng_E = InStr(mstrCode, Chr(&H1C))
        End If
    Loop
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