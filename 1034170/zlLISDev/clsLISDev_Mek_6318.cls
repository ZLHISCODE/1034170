VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mek_6318"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Mek 6318
'����:�ձ����
'����:�¶�
'�û�:��������ҽԺ
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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varline As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Dim strImage As String  'ͼ������
    Dim varImage As Variant 'ͼ������
    Dim lng_S As Long, lng_E As Long
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, "")
    Else
        lng_E = InStr(mstrCode, "")
    End If
    
    Do While lng_S > 0 And lng_E > 0

        'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
        strLine = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        strImage = Mid(mstrCode, lng_E + 1)
        
        If InStr(strImage, "") > 0 Then
            strImage = Mid(strImage, 1, InStr(strImage, ""))
            strImage = Replace(strImage, "", "")
        End If
        
        varline = Split(strLine, Chr(13))
        If strImage <> "" Then
            varImage = Split(strImage, Chr(13))
            If UBound(varImage) < 768 Then
                strLine = ""
                lng_S = 0: lng_E = 0
                Exit Do
            End If
        End If
        WriteLog "�յ���������", strLine & strImage, Replace(mstrCode, strImage, "")
        
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
        '����
        
        '�������
        For iRow = 0 To UBound(varline)
            Select Case iRow
            Case 0 '����
                If varline(iRow) Like "*##-*-##*" Then
                    str���� = Mid(varline(iRow), InStr(varline(iRow), "-") - 2, 2)
                    strTmp = Mid(varline(iRow), InStr(varline(iRow), "-") + 1)
                    strTmp = Mid(strTmp, InStr(strTmp, "-") + 1, 2) & "-" & TrunMonth(Mid(strTmp, 1, InStr(strTmp, "-") - 1))
                    str���� = Format(CDate(strTmp & "-" & str����), "yyyy-mm-dd")
                Else
                    str���� = Format(Now, "yyyy-MM-dd")
                End If
            Case 1 '�걾��
                str�걾�� = Val(varline(iRow))
                str�����˱걾�ʿ� = " |ѪҺ|0"
            Case 2 'WBC  ��ϸ������
                strItem = "|WBC|" & Format_6318(varline(iRow))
            Case 3 'LYM% �ܰ�ϸ���ٷֱ�
                strItem = strItem & "|LYMP|" & Format_6318(varline(iRow))
            Case 4 'MON% ��ֵϸ���ٷֱ�
                strItem = strItem & "|MONP|" & Format_6318(varline(iRow))
            Case 5 'GRA% ����ϸ���ٷֱ�
                strItem = strItem & "|GRAP|" & Format_6318(varline(iRow))
            Case 6 'LYM# �ܰ�ϸ������ֵ
                strItem = strItem & "|LYMA|" & Format_6318(varline(iRow))
            Case 7 'MON# ��ֵϸ������ֵ
                strItem = strItem & "|MONA|" & Format_6318(varline(iRow))
            Case 8 'GRA# ����ϸ������ֵ
                strItem = strItem & "|GRAA|" & Format_6318(varline(iRow))
            Case 9 'EOS# ��������ϸ������ֵ
                strItem = strItem & "|EOS|" & Format_6318(varline(iRow))
            Case 10 'RBC  ��ϸ������
                strItem = strItem & "|RBC|" & Format_6318(varline(iRow))
            Case 11 'HGB Ѫ�쵰��
                strItem = strItem & "|HGB|" & Format_6318(varline(iRow))
            Case 12 'HCT  ��ϸ��ѹ��
                strItem = strItem & "|HCT|" & Format_6318(varline(iRow))
            Case 13 'MCV  ƽ����ϸ�����
                strItem = strItem & "|MCV|" & Format_6318(varline(iRow))
            Case 14 'MCH  ƽ��Ѫ�쵰�׺���
                strItem = strItem & "|MCH|" & Format_6318(varline(iRow))
            Case 15 'MCHC ƽ��Ѫ�쵰��Ũ��
                strItem = strItem & "|MCHC|" & Format_6318(varline(iRow))
            Case 16 'RDW  ��ϸ��ƽ������
                strItem = strItem & "|RDW|" & Format_6318(varline(iRow))
            Case 17 'PLT ѪС������
                strItem = strItem & "|PLT|" & Format_6318(varline(iRow))
            Case 18 'PCT  ѪС��ѹ��
                strItem = strItem & "|PCT|" & Format_6318(varline(iRow))
            Case 19 'MPV  ƽ��ѪС�����
                strItem = strItem & "|MPV|" & Format_6318(varline(iRow))
            Case 20 'PDW  ѪС��ƽ������
                strItem = strItem & "|PDW|" & Format_6318(varline(iRow))
            End Select
        Next
        
        '����ͼ��
        If strImage <> "" Then
            For iRow = 0 To UBound(varImage)
                If iRow >= 0 And iRow <= 256 Then
                    If iRow = 0 Then strImage = "^WBC;0;0,271,15-50,51-100,87-,123-200,159-,195-300,235-,271-400"
                    strImage = strImage & ";" & Val(varImage(iRow))
                End If
                If iRow > 256 And iRow <= 512 Then
                    If iRow = 257 Then strImage = strImage & "^RBC;0;0,256,18-,56-100,94-,132-200,170-,208-300,246-"
                    strImage = strImage & ";" & Val(varImage(iRow))
                End If
                If iRow > 512 And iRow <= 768 Then
                    If iRow = 513 Then strImage = strImage & "^PLT;0;0,220,9-2,42-,75-20,108-,141-40,174-,207-60"
                    strImage = strImage & ";" & Val(varImage(iRow))
                End If
            Next
        End If
        If strItem <> "" And Val(str�걾��) <> 0 Then
            strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
        End If

        lng_S = InStr(mstrCode, "")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, "")
        Else
            lng_E = InStr(mstrCode, "")
        End If
    Loop


    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
        If strImage <> "" Then
            strResult = strResult & vbCrLf & Mid(strImage, 2)
        End If
    End If

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

Private Function Format_6318(ByVal strVal As String) As String
    '��ʽת��
    '��>,<��ԭ�����أ�����תΪ���ַ���
    If InStr(strVal, ">") > 0 Or InStr(strVal, "<") > 0 Then
        Format_6318 = Trim(strVal)
    Else
        Format_6318 = Val(Trim(strVal))
    End If
End Function

