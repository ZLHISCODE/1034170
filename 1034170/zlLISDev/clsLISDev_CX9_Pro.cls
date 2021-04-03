VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CX9_Pro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:CX9 Pro
'����:���������������(Beckman coulter)
'�û�:�ɶ���ͯר��ҽԺ

Private mstr_���� As String
Private mstr_�걾�� As String
Private mstr_Items As String
Private curTP As Currency '07C
Private curALB As Currency '08A

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    
    Dim lng_S As Long, lng_E As Long
    Dim str_Line As String, varTmp As Variant
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String, str_��� As String
    
    strReserved = strSource
    If strReserved Like "*[[]*[]]*" Then
        
        lng_S = InStr(strReserved, "[")
        lng_E = InStr(strReserved, "]")
        str_Line = Mid(strReserved, lng_S, lng_E - lng_S + 1)
        strReserved = Mid(strReserved, lng_E + 1)
        If InStr(strReserved, "[") <= 0 Then strReserved = ""
        
        varTmp = Split(str_Line, ",")
        If UBound(varTmp) > 3 Then
            If varTmp(1) = "702" And varTmp(2) = "01" Then
                mstr_���� = varTmp(3)
                mstr_���� = Format(CDate(Mid(mstr_����, 5, 2) & "-" & Mid(mstr_����, 3, 2) & "-" & Mid(mstr_����, 1, 2)), "yyyy-MM-dd")
                mstr_�걾�� = Val(varTmp(12))
                
                mstr_Items = ""
                
            ElseIf varTmp(1) = "702" And varTmp(2) = "03" Then
                str_���� = varTmp(3)
                str_���� = Format(CDate(Mid(str_����, 5, 2) & "-" & Mid(str_����, 3, 2) & "-" & Mid(str_����, 1, 2)), "yyyy-MM-dd")
                
                str_�걾�� = Val(varTmp(9))
                str_��Ŀ = Trim(varTmp(10))
                str_��� = Val(varTmp(15))
                
                If Val(str_�걾��) = Val(mstr_�걾��) And str_���� = mstr_���� Then
                    mstr_Items = mstr_Items & "|" & str_��Ŀ & "|" & str_���
                    
                    If str_��Ŀ = "07C" Then curTP = Val(str_���)
                    If str_��Ŀ = "08A" Then curALB = Val(str_���)
                    
                    
                End If
            ElseIf varTmp(1) = "702" And varTmp(2) = "05" Then
                str_�걾�� = Val(varTmp(6))
                If mstr_���� <> "" And Val(mstr_�걾��) = Val(str_�걾��) And mstr_Items <> "" Then
                    strResult = mstr_���� & "|" & Val(mstr_�걾��) & "| |Ѫ|" & IIf(Val(mstr_�걾��) >= 1000, "1", "0") & mstr_Items
                    
                    If curTP <> 0 And curALB <> 0 Then
                        strResult = strResult & "|GLB|" & curTP - curALB & "|A/G|" & Format(curALB / (curTP - curALB), "0.000")
                    End If
                End If
                mstr_���� = "": mstr_�걾�� = "": mstr_Items = ""
            End If
        End If
        
    End If
End Sub

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