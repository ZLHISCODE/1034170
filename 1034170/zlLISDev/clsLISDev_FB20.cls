VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FB20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:FB20
'����:ɽ����ɭʵҵ���޹�˾
'����:�ƽ���
'�û�:�ɶ�����ҽԺ
Private mstrCode As String  '����������
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
    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim lng_count As Long
    Dim str_Segment As String 'һ����Ϣ��(��chr(&HD)�ָ�)
    Dim var_Filed As Variant   '�ֶ�
    Dim strTmp As String
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_�걾�� As String, str_���� As String, str_������� As String
    Dim str_�ʿ� As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    Dim str_Return As String, strChanle As String, strChanle2 As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 1)
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If str_Segment Like "20*/*/*" Then
              str_���� = Format(str_Segment, "yyyy-MM-dd hh:mm") & ":" & Format(Now, "ss")
            ElseIf str_Segment Like "NO*CH*" Then
                '�걾�ţ���Ŀ
                strTmp = Trim(Mid(str_Segment, InStr(str_Segment, "NO:") + 3, InStr(str_Segment, ":CH") - 4))
                str_�걾�� = strTmp
                strChanle = Trim(Mid(str_Segment, InStr(str_Segment, "CH") + 4))
            ElseIf InStr(str_Segment, "[") > 0 And InStr(str_Segment, "]") > 0 Then
                strChanle2 = Trim(Mid(str_Segment, 1, InStr(str_Segment, ":") - 1))
                str_��� = Trim(Mid(str_Segment, InStr(str_Segment, ":") + 1, InStr(str_Segment, "[") - InStr(str_Segment, ":") - 2))
                str_��Ŀ = str_��Ŀ & "|" & strChanle & "_" & strChanle2 & "|" & str_���
            End If
        Next
        
        
        If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "2", 1, 0) & str_��Ŀ
        End If
        str_��Ŀ = "": strTmp = "": strChanle = "": strChanle2 = ""
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(mstrCode, Chr(&H3))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "FB20_Analyse", strSource, Err.Number & " " & Err.Description
End Sub