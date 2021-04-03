VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Clinitek500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Clinitek 500 ��Һ������
'����:�¹��ݶ�(Bayer)
'����:�¶�
'�û�:��ͷ����ҽԺ
Dim mstrCode As String

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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant
    Dim str_Line As String, int_Loop As Integer
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String
    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    
    
    If lng_S > 0 And lng_E > 0 Then
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        var_Code = Split(str_Code, vbCrLf)
        For int_Loop = 0 To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If Left(str_Line, 1) = "#" Then
                str_���� = "20" & Right(str_Line, 8)
                str_�걾�� = Trim(Mid(str_Line, InStr(str_Line, "-") + 1, Len(str_Line) - InStr(str_Line, "-") - 8))
            ElseIf Len(str_Line) = 20 Then
                str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & Trim(Mid(str_Line, 4))
            End If
        Next
        
        If Val(str_�걾��) <> 0 And IsDate(str_����) And str_��Ŀ <> "" Then
            str_��Ŀ = Replace(str_��Ŀ, "umol/L", "")
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |��Һ|0" & str_��Ŀ
        End If
        
        If strResult <> "" Then strResult = Mid(strResult, 3)
    End If
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

