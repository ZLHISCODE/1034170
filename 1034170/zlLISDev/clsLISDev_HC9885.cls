VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HC9885"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:HC-9885 ����ʷ�����
'����:���ں���
'����:�¶�
'�û�:�����ο�ҽԺ

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
    Dim str_�걾�� As String, str_��Ŀ As String, str_Tmp As String
    Dim aItem() As String
    
    mstrCode = mstrCode & Replace(strSource, "!", "")
    mstrCode = Replace(mstrCode, Chr(&H0), "")
    
    lng_E = InStr(mstrCode, Chr(&HA))
    strResult = ""
    Do While lng_E > 0
        
        str_Code = Trim(Mid(mstrCode, 1, lng_E))
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "��ʼ����", str_Code, ""
        '----------------------------------------
        str_�걾�� = "": str_��Ŀ = "": str_Tmp = ""
        str_Code = DelSapce(str_Code)
        aItem = Split(str_Code, " ")
        If UBound(aItem) >= 6 Then
            If Trim(aItem(0)) <> "" Then str_�걾�� = aItem(0)
            If Trim(aItem(1)) <> "" Then str_��Ŀ = str_��Ŀ & "|01|" & aItem(1)
            If Trim(aItem(2)) <> "" Then str_��Ŀ = str_��Ŀ & "|02|" & aItem(2)
            If Trim(aItem(3)) <> "" Then str_��Ŀ = str_��Ŀ & "|03|" & aItem(3)
            If Trim(aItem(4)) <> "" Then str_��Ŀ = str_��Ŀ & "|04|" & aItem(4)
            If Trim(aItem(5)) <> "" Then str_��Ŀ = str_��Ŀ & "|05|" & aItem(5)
            If Trim(aItem(6)) <> "" Then str_��Ŀ = str_��Ŀ & "|06|" & aItem(6)
            
            If UBound(aItem) >= 7 Then _
                If Trim(aItem(7)) <> "" Then str_��Ŀ = str_��Ŀ & "|07|" & aItem(7)
            If UBound(aItem) >= 8 Then _
                If Trim(aItem(8)) <> "" Then str_��Ŀ = str_��Ŀ & "|08|" & aItem(8)
        End If
        
        
'        str_�걾�� = Mid(str_Code, 1, 3)
'        str_Tmp = Trim(Mid(str_Code, 7, 4))
'        If str_Tmp <> "" Then str_��Ŀ = str_��Ŀ & "|01|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 14, 5))
'        If str_Tmp <> "" Then str_��Ŀ = str_��Ŀ & "|02|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 22, 4))
'        If str_Tmp <> "" Then str_��Ŀ = str_��Ŀ & "|03|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 29, 4))
'        If str_Tmp <> "" Then str_��Ŀ = str_��Ŀ & "|04|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 36, 4))
'        If str_Tmp <> "" Then str_��Ŀ = str_��Ŀ & "|05|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 43, 4))
'        If str_Tmp <> "" Then str_��Ŀ = str_��Ŀ & "|06|" & str_Tmp
        
        If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_�걾�� & "| |Ѫ��|0" & str_��Ŀ
            WriteLog "��ɽ���", "", strResult
        End If
        str_��Ŀ = ""
        str_�걾�� = ""
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
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function