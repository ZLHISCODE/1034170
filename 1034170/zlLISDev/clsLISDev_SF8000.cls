VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SF8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:SF-8000ȫ�Զ���Ѫ������
'����:����ϣ��
'����:�¶�
'�û�:�����зο�ҽԺ
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
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String
    mstrCode = mstrCode & strSource
    lng_E = InStr(mstrCode, Chr(&HA))
    
    Do While lng_E > 0
        str_Code = Mid(mstrCode, 1, lng_E)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------------
        If Len(str_Code) > 12 Then
            str_�걾�� = Trim(Mid(str_Code, 1, 4))
            str_��Ŀ = Trim(Mid(str_Code, 6, 6))
            str_��� = Trim(Mid(str_Code, 12))
            str_��Ŀ = Replace(str_��Ŀ, "g/L", "")
            
            If Val(str_�걾��) > 0 And str_��Ŀ <> "" And str_��� <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_�걾�� & "| |Ѫ��|0|" & str_��Ŀ & "|" & str_���
            End If
        End If
        
        '----------------------------------------
        lng_E = InStr(mstrCode, Chr(&HA))
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
    
End Function

