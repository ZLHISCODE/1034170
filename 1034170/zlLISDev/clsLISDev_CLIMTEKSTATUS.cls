VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CLIMTEKSTATUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:CLIMTEKSTATUS ��Һ������
'����:
'����:�¶�
'�û�:�»�,���������
'2011-05-11
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
    Dim str_Line As String, i As Integer
    Dim str_Date As String, str_SampleNo As String, str_Item As String
    mstrCode = mstrCode & strSource
    var_Code = Split(mstrCode, ",")
    If UBound(var_Code) >= 41 Then
        str_SampleNo = var_Code(0)
        str_Date = var_Code(1)
        str_Date = Mid(str_Date, 7, 4) & "-" & Mid(str_Date, 4, 2) & "-" & Mid(str_Date, 1, 2) & " " & var_Code(2)
        str_Date = Format(CDate(str_Date), "yyyy-MM-dd HH:mm")
        For i = 9 To 41 Step 3
            If Trim(var_Code(i)) <> "" And Trim(var_Code(i + 1)) <> "" Then
                str_Item = str_Item & "|" & Trim(var_Code(i)) & "|" & Trim(var_Code(i + 1))
            End If
        Next
        If Val(str_SampleNo) > 0 And IsDate(str_Date) And Trim(str_Item) <> "" Then
            strResult = strResult & "||" & str_Date & "|" & str_SampleNo & "| |��Һ|0" & Replace(str_Item, "Trace", "΢��")
        End If
        str_Code = ""
        For i = 42 To UBound(var_Code)
            str_Code = str_Code & "," & var_Code(i)
            
        Next
        mstrCode = str_Code
    End If
        
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
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function
