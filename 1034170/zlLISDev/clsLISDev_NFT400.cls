VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_NFT400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: ����������
'����: ����
'����: �¶�
'�û�: ɽ������ɽҽԺ
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
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim var_Code As Variant, str_Line As String, i As Integer
    Dim str_NO As String, str_Date As String, str_Item As String, str_Value As String
    
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    If InStr(strSource, ">") > 0 Then strCmd = "<Saved|1F3>"
    
    lng_S = InStr(mstrCode, "<")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ">")
    Else
        lng_E = InStr(mstrCode, ">")
    End If
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------
        str_Item = Trim$(Split(str_Code, "|")(3))
        str_Value = Trim$(Split(str_Code, "|")(4))
        str_NO = Trim$(Split(str_Code, "|")(2))
        If Val(str_NO) > 0 And str_Value <> "" And str_Item <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_NO & "| |Ѫ��|0|" & str_Item & "|" & str_Value
        ElseIf Trim("" & str_NO) <> "" And str_Value <> "" And str_Item <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_NO & "| |Ѫ��|1|" & str_Item & "|" & str_Value
        End If
        '--------------------------------
        lng_S = InStr(mstrCode, "<")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, ">")
        Else
            lng_E = InStr(mstrCode, ">")
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
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


