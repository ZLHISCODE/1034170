VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_EasyBloodGas_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:MEDICA EasyBloodGas
'����:MEDICA
'����:�¶�
'�û�:����ҽԺ

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String
    Dim str�� As String, str�� As String, str�� As String, strʱ�� As String, strָ�� As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    
    Dim lng_S As Long '��ʼ
    Dim lng_E As Long '����
    Dim str_Line As String
    Dim lng_�걾�� As String
    Dim var_data As Variant
    Dim i As Long
    Dim str_Code As String, str_��� As String
    
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    strָ�� = ",pH,PCO2,PO2,pH(T),PCO2(T),PO2(T),HCO2,BEb,BEecf,SBC,%SO2c,O2ct,A-aDO2,RI"
    
    mstrCode = mstrCode & strSource
    '---ȡ����������
    lng_S = InStr(mstrCode, Chr(&HD) & Chr(&H3D))     '�� ��ͷ��������������Ŀ�ʼλ��
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ";")  '�� ������������������Ľ���λ��
        If lng_E > 0 Then
            lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD))
        Else
            lng_E = 0
        End If
    Else
        lng_E = 0
    End If
    
    Do While lng_S > 0 And lng_E > 0 '�����ʼ����������Ϊ��
        '----- ��ʼ�������
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 5) '���شӿ�ʼλ�õ�����λ��֮����ַ�
        mstrCode = Mid(mstrCode, lng_E + 5)
        strItem = ""
        str���� = ""
        str�걾�� = ""
        
        var_data = Split(str_Code, Chr(&HD)) 'ȡÿ������
        For iRow = LBound(var_data) To UBound(var_data)
            str_Line = var_data(iRow)
            
            If InStr(str_Line, "ANALYSIS RESULTS") > 0 Then
                str�걾�� = Val(Split((str_Line), Chr(&H20))(2))
                strItem = ""
                str�����˱걾�ʿ� = " |ѪҺ|0"
            ElseIf InStr(str_Line, ";") > 0 Then
                str���� = Replace(Replace(str_Line, Chr(&H3F), ""), ";", "")
                str�� = Mid(str����, 1, InStr(str����, "-") - 1): strʱ�� = Split(str����, " ")(1)
                str�� = Split(str����, " ")(0)
                str�� = Split(str��, "-")(1)
                str�� = Split(str����, " ")(0)
                str�� = Split(str��, "-")(2)
                str���� = Format(CDate(str�� & "-" & TrunMonth(str��) & "-" & str��) & " " & strʱ��, "yyyy-MM-dd hh:mm")
            Else
                If InStr(str_Line, "pH(T)") > 0 Then
                    strItem = strItem & "|pH(T)|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PCO[(T)") > 0 Then
                    strItem = strItem & "|PCO2(T)|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PO[(T)") > 0 Then
                    strItem = strItem & "|PO2(T)|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PCO[") > 0 Then
                    strItem = strItem & "|PCO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PO[") > 0 Then
                    strItem = strItem & "|PO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "pH") > 0 Then
                    strItem = strItem & "|pH|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "TCO[") > 0 Then
                    strItem = strItem & "|TCO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "HCO]") > 0 Then
                    strItem = strItem & "|HCO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "BEb") > 0 Then
                    If Split(trimString(str_Line), " ")(1) = "-" Then
                        strItem = strItem & "|BEb|" & Split(trimString(str_Line), " ")(1) & Split(trimString(str_Line), " ")(2)
                    Else
                        strItem = strItem & "|BEb|" & Split(trimString(str_Line), " ")(1)
                    End If
                    
                ElseIf InStr(str_Line, "BEecf") > 0 Then
                    If Split(trimString(str_Line), " ")(1) = "-" Then
                        strItem = strItem & "|BEecf|" & Split(trimString(str_Line), " ")(1) & Split(trimString(str_Line), " ")(2)
                    Else
                        strItem = strItem & "|BEecf|" & Split(trimString(str_Line), " ")(1)
                    End If
                    
                ElseIf InStr(str_Line, "SBC") > 0 Then
                    strItem = strItem & "|SBC|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "%SO[c") > 0 Then
                    strItem = strItem & "|%SO2c|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "O2ct") > 0 Then
                    strItem = strItem & "|O2ct|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "A-aDO[") > 0 Then
                    strItem = strItem & "|A-aDO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "RI") > 0 Then
                    strItem = strItem & "|RI|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "Hb") > 0 Then
                    strItem = strItem & "|Hb|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "FIO[") > 0 Then
                    strItem = strItem & "|FIO|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "B.P.") > 0 Then
                    strItem = strItem & "|BP|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PAT.TEMP") > 0 Then
                    strItem = strItem & "|PATTEMP|" & Split(trimString(str_Line), " ")(1)
                End If
            End If
            
        Next
        
        If strItem <> "" And str���� <> "" And Val(str�걾��) > 0 Then
            strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
        End If
            
        '---ȡ����������
        lng_S = InStr(mstrCode, Chr(&HD) & Chr(&H3D))     '�� ��ͷ��������������Ŀ�ʼλ��
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, ";")  '�� ������������������Ľ���λ��
            If lng_E > 0 Then
                lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD))
            Else
                lng_E = 0
            End If
        Else
            lng_E = 0
        End If
    Loop
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    '���豸���͡������յ�������
    'strCmd = Chr(&H6&)
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
    
    SendSample = ""
    
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function

Private Function getNumber_Right(ByVal strSource As String) As String
    '���ҿ�ʼȡ���֣������ո����
    Dim strTmp As String, lngLength As Long
    Dim strIn As String
    strIn = strSource
    
    If Len(strIn) > 0 Then
        For lngLength = 1 To Len(strSource)
            strTmp = Right(strIn, 1)
            If IsNumeric(strTmp) Then
                getNumber_Right = strTmp & getNumber_Right
                strIn = Mid(strIn, 1, Len(strIn) - 1)
            ElseIf strTmp = "." Then
                If InStr(getNumber_Right, ".") <= 0 Then
                    getNumber_Right = strTmp & getNumber_Right
                    strIn = Mid(strIn, 1, Len(strIn) - 1)
                Else
                    Exit For
                End If
            ElseIf strTmp = " " Then
                If getNumber_Right <> "" Then
                    Exit For
                Else
                    strIn = Mid(strIn, 1, Len(strIn) - 1)
                End If
            End If
        Next
    End If
    If getNumber_Right <> "" Then
        If InStr(strSource, "- ") > 0 Then getNumber_Right = "-" & getNumber_Right
    End If
End Function

Private Function trimString(ByVal Str As String) As String
        Dim strTemp       As String
        
        strTemp = Trim(Str)
        Do While InStr(1, strTemp, Space(2)) > 0
                strTemp = Replace(strTemp, Space(2), Space(1))
        Loop
        trimString = strTemp
End Function



