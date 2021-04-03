VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_AFT_TCO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AFT-TCO2������̼������
'����:���ڿ���
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
    Dim str_Return As String, strLog As String
    
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    If Len(strSource) < 45 Then strReserved = strSource: Exit Sub
    
    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " ����ԭʼ���ݣ�" & strSource
    Call writelislog(strLog)
    
    mstrCode = Replace(mstrCode, "Ѫ��������", Chr(&HB))
    mstrCode = Replace(mstrCode, "MMOL/L", Chr(&H1C))
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HA))
        str_Return = "": str_�걾�� = "": str_���� = ""
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            Select Case lng_count
                Case 0
                   str_�걾�� = Trim(str_Segment)
                   str_�걾�� = Trim(Replace(str_�걾��, Chr(&HB), ""))
                Case 1
                   str_Segment = Replace(str_Segment, Chr(&HA), "")
                   str_Segment = Replace(str_Segment, Chr(&HD), "")
                   str_���� = Trim(str_Segment) & ":" & Format(Now, "SS")
                Case 3
                   str_Segment = Replace(str_Segment, Chr(&H1C), "")
                   str_Segment = Replace(str_Segment, Chr(&HA), "")
                   str_��Ŀ = Trim(Mid(str_Segment, 1, InStr(str_Segment, "=") - 1))
                   str_��� = Trim(Mid(str_Segment, InStr(str_Segment, "=") + 1))
                   str_��� = Trim(Replace(str_���, Chr(&H20), ""))
                   str_Return = str_Return & "|" & str_��Ŀ & "|" & str_���
            End Select
        Next
        If Val(str_�걾��) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_Return
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
'    writelislog (strResult)
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_CO2_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function

