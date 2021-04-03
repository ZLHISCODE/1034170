VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_EasyBloodGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:MEDICA EasyBloodGas
'����:MEDICA
'����:�¶�
'�û�:����ҽԺ

Private strDeviceString '������Ϣ���豸��
Private strLastString As String '�ϴη��͵��豸��
Private mStrλ�� As String  '˫��ͨ���á�

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String
    Dim str�� As String, str�� As String, str�� As String, strָ�� As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    strָ�� = ",pH,PCO2,PO2,pH(T),PCO2(T),PO2(T),HCO2,BEb,BEecf,SBC,%SO2c,O2ct,A-aDO2,RI"
    If strSource Like "*[#]*-##-##;*" Then
        '----- ��ʼ�������
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ����������Ϣ
            strLine = GetStr_Section(strReserved, "#", ";")
            strReserved = Mid(strReserved, InStr(strReserved, ";") + 1)

            '--�������
            varLine = Split(strLine, Chr(&HD&))
            For iRow = LBound(varLine) To UBound(varLine)
                If varLine(iRow) Like "[#]*" Then
                    str�걾�� = Val(Replace(varLine(iRow), "#", ""))
                    strItem = ""
                    str�����˱걾�ʿ� = " |ѪҺ|0"
                ElseIf varLine(iRow) Like "*-##-##;*" Then
                    str���� = Replace(Replace(varLine(iRow), Chr(&H3F), ""), ";", "")
                    str�� = Mid(str����, 1, InStr(str����, "-") - 1): str���� = Mid(str����, InStr(str����, "-") + 1)
                    str�� = Mid(str����, 1, InStr(str����, "-") - 1)
                    str�� = Mid(str����, InStr(str����, "-") + 1, 2)
                    If str�� Like "##" And str�� Like "##" And str�� Like "##" Then
                        str���� = Format(CDate(str�� & "-" & str�� & "-" & str��), "yyyy-MM-dd")
                    Else
                        str���� = Format(Now, "yyyy-MM-dd")
                    End If
                Else
                    strTmp = varLine(iRow)
                    If InStr(strTmp, "*****") <= 0 Then
                        If InStr(strTmp, "pH(T)") > 0 Then
                            strItem = strItem & "|pH(T)|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "PCO[(T)") > 0 Then
                            strItem = strItem & "|PCO2(T)|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "PO[(T)") > 0 Then
                            strItem = strItem & "|PO2(T)|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "PCO[") > 0 Then
                            strItem = strItem & "|PCO2|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "PO[") > 0 Then
                            strItem = strItem & "|PO2|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "pH") > 0 Then
                            strItem = strItem & "|pH|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "TCO[") > 0 Then
                            strItem = strItem & "|TCO2|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "HCO]") > 0 Then
                            strItem = strItem & "|HCO2|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "BEb") > 0 Then
                            strItem = strItem & "|BEb|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "BEecf") > 0 Then
                            strItem = strItem & "|BEecf|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "SBC") > 0 Then
                            strItem = strItem & "|SBC|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "%SO[c") > 0 Then
                            strItem = strItem & "|%SO2c|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "O2ct") > 0 Then
                            strItem = strItem & "|O2ct|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "A-aDO[") > 0 Then
                            strItem = strItem & "|A-aDO2|" & getNumber_Right(Mid(strTmp, 1, 14))
                        ElseIf InStr(strTmp, "RI") > 0 Then
                            strItem = strItem & "|RI|" & getNumber_Right(Mid(strTmp, 1, 14))
                        End If
                    End If
                End If
            Next
            If strItem <> "" And str���� <> "" And Val(str�걾��) > 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            If Not strReserved Like "**[#]*-##-##;*" Then
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
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

