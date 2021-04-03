VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XC_A30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:XC-A30Ѫ����
'����:����Խ��
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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim str_Code As String, mstrCode As String
    Dim var_Code As Variant
    Dim lng_count As Integer
    
    Dim str_Segment As String 'һ����Ϣ��(��chr(&HA)�ָ�)
    
    Dim lng_S As Integer
    Dim lng_E As Integer
    
    Dim str_�걾�� As String, str_���� As String
    Dim str_��Ŀ As String, str_��� As String
    Dim str_Return As String
    
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lng_S = InStr(mstrCode, Chr(&H18))
    lng_E = InStr(mstrCode, Chr(&H24))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HD))
        str_Return = "": str_�걾�� = "": str_���� = ""
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            str_Segment = Replace(str_Segment, Chr(&HD), "")
            str_Segment = Replace(str_Segment, Chr(&HA), "")
            
            If IsNumeric(Trim(Mid(str_Segment, 1, 2))) Then
              str_�걾�� = Trim(Mid(str_Segment, 1, 2))
              str_���� = Format(Now, "yyyy-MM-dd HH:MM:SS")
              str_��Ŀ = "ESR"
              str_��� = Trim(Mid(str_Segment, 25))
              str_��� = Trim(Replace(str_���, Chr(&H20), ""))
              str_Return = str_��Ŀ & "|" & str_���
              If Val(str_�걾��) > 0 And str_Return <> "" Then
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0|" & str_Return
              End If
            End If
        Next
        lng_S = InStr(mstrCode, Chr(&H18))
        lng_E = InStr(mstrCode, Chr(&H24))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
 
End Sub


