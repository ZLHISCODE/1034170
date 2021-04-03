VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AVL_OPTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:OPTI CCA Ѫ��������
'����:Osmetech
'����:������
'�û�:̫ԭ¯��
'ʱ��:2012-12-17
Private mstrCode As String '��������
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

    Dim strDate As String
    Dim strNo As String
    Dim strItems As String
    Dim varItem As Variant
    Dim strItem As String
    Dim strLine As String
    Dim str_Sample As String
    Dim lng_S As Long
    Dim lng_E As Long
    Dim intloop As Integer
    Dim strAllItem As String
    
    On Err GoTo errHandle
    mstrCode = mstrCode & strSource    '���滺�������
    'Ӧ����
    strAllItem = ",ACID/BASE,PH,PCO2,PO2,BE,tCO2,HCO3,BB,BEact,BEecf,stHCO3,st.pH,cH+,Na+,K+,Ca++,nCa++,Cl-,AnGap,tHb,SO2,Hct(c),AaDO2,O2Ct,P50(c),"
    lng_E = 0: lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    
    Do While lng_S > 0 And lng_E > 0
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        varItem = Split(str_Sample, Chr(&HA) & Chr(&HD))
        For intloop = LBound(varItem) To UBound(varItem)
            strLine = Trim(varItem(intloop))
            strLine = Replace(strLine, "", "")
            strLine = Replace(strLine, "", "")
            
            If strLine <> "" Then
                If strLine Like "*##-???-##*" Then
                    strLine = Replace(strLine, " ", "")
                    strLine = Mid(strLine, 8, 2) & "-" & Mid(strLine, 4, 3) & "-" & Mid(strLine, 1, 2) & " " & Mid(strLine, 10, 2) & ":" & Mid(strLine, 13, 2)
                    strDate = Format(strLine, "yyyy-mm-dd hh:mm:ss")
                ElseIf InStr(strLine, "Sample") > 0 Then
                    strNo = Mid(strLine, InStr(strLine, ":") + 1)
                ElseIf InStr(strLine, " -") > 0 Then
                '----------------------------------������-----------------------------------
                ElseIf InStr(strLine, "Ref.Lim") > 0 Then
                '----------------------------------������-----------------------------------
                ElseIf InStr(strAllItem, Split(Trim(strLine), " ")(0)) > 0 Then
                    
                    If Split(strLine, " ")(0) Like "ACID/BASE" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 10, 4)
                    ElseIf Split(strLine, " ")(0) Like "pH" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 3, 4)
                    ElseIf Split(strLine, " ")(0) Like "st.pH" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 6, 5)
                    ElseIf Split(strLine, " ")(0) Like "tHb" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 4, 3)
                    ElseIf Split(strLine, " ")(0) Like "SO2" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 4, 2)
                    ElseIf Split(strLine, " ")(0) Like "Hct(c)" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 7, 2)
                    ElseIf Split(strLine, " ")(0) Like "O2Ct" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 5, 4)
                    ElseIf Split(strLine, " ")(0) Like "Hb" Then
                        strItem = ""
                    ElseIf Split(strLine, " ")(0) Like "P50" Then
                        strItem = ""
                    Else
                        strItem = Trim(Mid(strLine, Len(Split(strLine, " ")(0)) + 1, InStr(strLine, "mm") - Len(Split(strLine, " ")(0)) - 1))
                        strItem = Split(strLine, " ")(0) & "|" & strItem
                    End If
                    If strItem <> "" Then
                        strItems = strItems & "|" & strItem
                    End If
                End If
            End If
        Next
        
        If strDate <> "" And strNo <> "" And strItems <> "" Then
           strResult = strResult & "||" & strDate & "|" & Val(strNo) & "| |ѪҺ|0" & strItems
        End If
            
        lng_E = 0: lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Loop
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
    End If
errHandle:
    WriteLog "Test_analyse", strSource, Err.Number & " " & Err.Description
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





