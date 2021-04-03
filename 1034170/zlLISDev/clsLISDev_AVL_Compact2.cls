VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AVL_Compact2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AVL Compact 2
'����:����
'����:�¶�
'�û�:���������ڶ�ҽԺ

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String
    Dim str�� As String, str�� As String, str�� As String, strָ�� As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    strָ�� = ",pH,PCO2,PO2,BE,BEecf,BB,HCO3,cH+,st.HCO3,st.pH,TCO2,O2sat,O2cont,AaDO2,"
    If strSource Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
        '----- ��ʼ�������
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ����������Ϣ
            strLine = GetStr_Section(strReserved, Chr(&H2), Chr(&H3))
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H3)) + 1)

            '--�������
            varline = Split(strLine, vbCrLf)
            For iRow = LBound(varline) To UBound(varline)
                If varline(iRow) Like "No.:*" Then
                    str�걾�� = Val(Split(varline(iRow), ":")(1))
                    strItem = ""
                    str�����˱걾�ʿ� = " |ѪҺ|0"
                ElseIf varline(iRow) Like "*,##-*-## ##:##*" Then
                    str���� = Mid(varline(iRow), InStr(varline(iRow), ",") + 1)
                    str�� = Mid(str����, 1, 2): str���� = Mid(str����, 4)
                    str�� = TrunMonth(Mid(str����, 1, InStr(str����, "-") - 1))
                    str�� = str�� & " " & Mid(str����, InStr(str����, ":") - 2, 5)
                    str�� = Mid(str����, InStr(str����, "-") + 1, 2)
                    str���� = Format(CDate(str�� & "-" & str�� & "-" & str��), "yyyy-MM-dd HH:mm")
                Else
                    strTmp = varline(iRow)
                    If InStr(strTmp, "Baro") > 0 Then
                        strItem = strItem & "|Baro|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "#tHb A") > 0 Then
                        strItem = strItem & "|#tHb A|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "#Temp") > 0 Then
                        If Trim(Mid(strTmp, 13, 5)) <> "" Then
                            strItem = strItem & "|#Temp|" & Trim(Mid(strTmp, 13, 5))
                        ElseIf Trim(Mid(strTmp, 7, 5)) <> "" Then
                            strItem = strItem & "|#Temp|" & Trim(Mid(strTmp, 7, 5))
                        End If
                    ElseIf Mid(strTmp, 1, 2) = "pH" Then
                        If Trim(Mid(strTmp, 13, 5)) <> "" Then
                            strItem = strItem & "|pH|" & Trim(Mid(strTmp, 13, 5))
                        ElseIf Trim(Mid(strTmp, 7, 5)) <> "" Then
                            strItem = strItem & "|pH|" & Trim(Mid(strTmp, 7, 5))
                        End If
                    ElseIf InStr(strTmp, "PCO2") > 0 Then
                        If Trim(Mid(strTmp, 13, 5)) <> "" Then
                            strItem = strItem & "|PCO2|" & Trim(Mid(strTmp, 13, 5))
                        ElseIf Trim(Mid(strTmp, 7, 5)) <> "" Then
                            strItem = strItem & "|PCO2|" & Trim(Mid(strTmp, 7, 5))
                        End If
                    ElseIf InStr(strTmp, "PO2") > 0 Then
                        If Trim(Mid(strTmp, 13, 5)) <> "" Then
                            strItem = strItem & "|PO2|" & Trim(Mid(strTmp, 13, 5))
                        ElseIf Trim(Mid(strTmp, 7, 5)) <> "" Then
                            strItem = strItem & "|PO2|" & Trim(Mid(strTmp, 7, 5))
                        End If
                    ElseIf Mid(strTmp, 1, 4) = " BE " Then
                        strItem = strItem & "|BE|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "BEecf") > 0 Then
                        strItem = strItem & "|BEecf|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "BB") > 0 Then
                        strItem = strItem & "|BB|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, " HCO3") > 0 Then
                        strItem = strItem & "|HCO3|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "cH+") > 0 Then
                        strItem = strItem & "|cH+|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "st.HCO3") > 0 Then
                        strItem = strItem & "|st.HCO3|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "st.pH") > 0 Then
                        strItem = strItem & "|st.pH|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "TCO2") > 0 Then
                        strItem = strItem & "|TCO2|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "O2sat") > 0 Then
                        strItem = strItem & "|O2sat|" & getNumber_Right(Mid(strTmp, 1, 17))
                    ElseIf InStr(strTmp, "O2cont") > 0 Then
                        strItem = strItem & "|O2cont|" & getNumber_Right(Mid(strTmp, 1, 17))

                    ElseIf InStr(strTmp, "AaDO2") > 0 Then
                        strItem = strItem & "|AaDO2|" & getNumber_Right(Mid(strTmp, 1, 17))

                    ElseIf InStr(strTmp, "#RQ") > 0 Then
                        strItem = strItem & "|#RQ|" & getNumber_Right(Mid(strTmp, 1, 12))
                        If InStr(strTmp, "#FIO2") > 0 Then
                            strItem = strItem & "|#FIO2|" & getNumber_Right(Mid(strTmp, 12))
                        End If
                    End If

                End If
            Next
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            If Not strReserved Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
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