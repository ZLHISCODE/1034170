VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisdev_ADVIACentaurXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��������: ��ѧ�������߷�����
'�����ͺ�: ADVIA centaur XP
'����:�¹�������
'����:
'�û�:��������ҽԺ(XP)

Private mstrCode As String  '����������
Private mstrWaitSend As String  '����������

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
    mstrWaitSend = ""
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

    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim var_Code As Variant, int_Loop As Integer
    Dim str_Code As String 'һ��������¼
    Dim str_Line As String 'һ������
    Dim strSampleNo As String, strTestItem As String, strTestValue As String, strDate As String
    Dim str_SendLine As String, str_Item As String
    Dim str_Tmp As String
    On Error GoTo hErr
    strResult = "": strCmd = ""
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
'    If Right(strSource, 1) = Chr(&H4) Then
'        '�յ�һ�����������ݣ���ʱ�������δ�������ݣ�����˫��ģʽ
'        If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
'            SaveLog "XP-��������δ�������ݣ��л���˫��ģʽ ", strSource, mstrWaitSend
'            mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
'            strCmd = "1|"
'            strResult = Format(Now, "YYYY-MM-DD") & "|^^0"
'            Exit Sub
'        ElseIf mstrWaitSend <> "" Then
'            'û�з���������ݣ�����
'            SaveLog "XP-��������δ���͵Ĳ��������ݣ�����", strSource, mstrWaitSend
'            mstrWaitSend = ""
'        End If
'    End If
    '-------------------------------------------------------------------------------------------------------------
    If InStr(mstrCode, Chr(&H5)) > 0 Then                                         '�յ������� 05 ,Ӧ�� 06
        strCmd = Chr(&H6)
    ElseIf InStr(mstrCode, Chr(&HA)) > 0 Then                               '�յ��н����� 03 ,Ӧ�� 06
         strCmd = Chr(&H6)
    End If
    '-------------------------------------------------------------------------------------------------------------
 
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = 0
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
 
    Do While lng_S > 0 And lng_E > 0    '�յ�һ�������ļ�¼,��ʼ����
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        Do While str_Code Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
            lng_S = InStr(str_Code, Chr(&H17))
            lng_E = InStr(lng_S, str_Code, Chr(&H2)) + 1
            str_Code = Mid(str_Code, 1, lng_S - 1) & Mid(str_Code, lng_E + 1)
        Loop
        '--------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)

            If str_Line Like "O|*" Then
               str_Tmp = Split(str_Line, "|")(2)
               str_Tmp = Split(str_Tmp, "^")(0)
               str_Tmp = Val(str_Tmp)
               If strSampleNo <> str_Tmp Then
                    If str_Item <> "" And Val(strSampleNo) > 0 And strDate <> "" Then
                        If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
                        strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Ѫ��|0" & str_Item
                    End If
                    strSampleNo = str_Tmp
                    str_Item = ""
               End If
            ElseIf str_Line Like "R|*|*^DOSE|*" Then
               strDate = Split(str_Line, "|")(12)
               If strDate <> "" Then
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                    If Not IsDate(strDate) Then strDate = ""
               End If
               strTestItem = Split(str_Line, "|")(2)
               strTestItem = Trim(Split(strTestItem, "^")(3))
               strTestValue = Trim(Split(str_Line, "|")(3))
               If strTestValue <> "" Then
                    str_Item = str_Item & "|" & Replace(strTestItem, "^", "") & "|" & strTestValue
               End If
            ElseIf str_Line Like "Q|*|ALL|*" Then
                '˫��
                strSampleNo = Trim(Split(str_Line, "|")(3))
                If strSampleNo <> "" Then
                    If Left$(strSampleNo, 1) = "^" Then strSampleNo = Mid$(strSampleNo, 2)
                    If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then
                        strCmd = "1|"
                        strResult = "||" & Format(Now, "yyyy-MM-dd") & "|" & "^^" & strSampleNo
                    End If
                End If
                
            End If
        Next
        
        If strCmd <> "1|" Then
            If str_Item <> "" And Val(strSampleNo) > 0 And strDate <> "" Then
                If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
                strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Ѫ��|0" & str_Item
            End If
        Else
            Exit Do
        End If
        '--------------------------------------------------------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        lng_E = 0
        If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    SaveLog "Centaur-�յ�" & strSource & ",��" & CStr(Erl()) & "��," & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '   strSample:������Ҫ���������͵ı걾��Ϣ
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
    Dim varInData As Variant, iLoop As Integer, strData As String, i As Integer
    Dim strNo As String, strEmergency As String, strTmp As String
    Dim strChannel As String, strSend As String, strItems As String, iRow As Integer
    Dim strSegment As String

    On Error GoTo hErr

100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '�ս���˫��ģʽ����û�����ɷ�������

106     SaveLog "Centaur-����˫��" & strSample & ",����" & mstrWaitSend
108     varInData = Split(strSample, "||")
110     For iLoop = LBound(varInData) To UBound(varInData)
112         strData = varInData(iLoop)
            '--����
114         strNo = Trim(Split(strData, "|")(5))
            '---����
116         If Val(Split(strData, "|")(4)) = 1 Then
118             strEmergency = "S"
            Else
120             strEmergency = "R"
            End If
            '----ָ��
         
122         strItems = ""
124         For i = 10 To UBound(Split(strData, "|"))
126             strChannel = Trim$("" & Split(strData, "|")(i))
128             strItems = strItems & "^^^" & strChannel & "\"
            Next
        
130         strSend = ""
132         If strItems <> "" And strNo <> "" Then
134             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
            
136             strSend = "H|\^&" & Chr(&HD) & "P|1" & Chr(&HD)
138             strSend = strSend & "O|1|" & strNo & "||" & strItems & "|" & strEmergency & "||||||||||||||||||||O\Q" & Chr(&HD)
140             strSend = strSend & "L|1" & Chr(&HD) & Chr(&H3)
            
142             iRow = 0: strSegment = ""
144             Do While strSend <> ""
146                 If Len(strSend) < 230 Then
148                     iRow = iRow + 1
150                     strSegment = Chr(&H5) & strSegment & Chr(&H2) & iRow & strSend & GetCheckCode(iRow & strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
152                     strSend = ""
                    Else
154                     iRow = iRow + 1
156                     strTmp = Mid(strSend, 1, 230) & Chr(&H17)
158                     strSegment = strSegment & Chr(&H2) & iRow & strTmp & GetCheckCode(iRow & strTmp) & Chr(&HD) & Chr(&HA)
160                     strSend = Mid(strSend, 231)
                    End If
                Loop
                
162             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment
            End If

164         SendSample = GetSendData
166         If mstrWaitSend <> "" Then iSendStep = 1    '����δ�������ݣ����˳�˫��ģʽ

        Next
168     SaveLog "Centaur-����˫������������,����:" & SendSample & ",״̬" & iSendStep & ",����:" & mstrWaitSend
    Else
        '�յ�����Ӧ�𣬴ӻ�����ȡ��ָ���

170     If InStr(strResponse, Chr(&H6)) > 0 Then
172         strTmp = "�յ�06"
174         SendSample = GetSendData
176         If mstrWaitSend = "" Then iSendStep = 0
178     ElseIf InStr(strResponse, Chr(&H5)) > 0 Then
            '�յ���������,����ͨѶʧ�ܣ�����δ���͵����ݣ��л�״̬
180         strTmp = "�յ�05,���ȴ����������������ݣ��˳�˫��"
182         mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
183         mstrCode = mstrCode & Chr(&H5)
184         SendSample = Chr(&H6)
186         iSendStep = 0
188     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
190         strTmp = "�յ�04,�˳�˫��,��ջ���"
192         SendSample = ""
194         mstrWaitSend = ""
196         iSendStep = 0
        Else
198         If mstrWaitSend = "" Then iSendStep = 0
200         strTmp = "�յ�����:" & strResponse
        End If
        
        
202     If iSendStep = 0 Then
204         strTmp = strTmp & ",�˳�˫��"
        Else
206         strTmp = strTmp & ",����˫��"
        End If
        
208     If SendSample <> "" Then
210         strTmp = strTmp & ",����:" & SendSample
        Else
212         strTmp = strTmp & ",�����ݷ���"
        End If
214     If mstrWaitSend <> "" Then
216         strTmp = strTmp & ",����" & mstrWaitSend
        Else
218         strTmp = strTmp & ",�����ݻ���"
        End If
220     SaveLog "Centaur-˫��" & strTmp
    End If
    Exit Function
hErr:
222     SaveLog "Centaur-��" & CStr(Erl()) & "���յ�" & strResponse & ",���ִ���:" & Err.Description & "����ǰ����" & mstrWaitSend
End Function

Private Function GetSendData() As String
    '�ӻ����л�ȡһ�����ݣ����ء����������û�����ݣ����ؿմ�

    
    If mstrWaitSend = "" Then Exit Function
    
    If Left$(mstrWaitSend, 1) = Chr(&H5) Then
        GetSendData = Chr(&H5)
        mstrWaitSend = Mid(mstrWaitSend, 2)
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H2) Then
        If InStr(1, mstrWaitSend, Chr(&HA)) > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(1, mstrWaitSend, Chr(&HA)))
            mstrWaitSend = Mid$(mstrWaitSend, InStr(1, mstrWaitSend, Chr(&HA)) + 1)
        End If
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H4) Then
        GetSendData = Chr(&H4)
        mstrWaitSend = Mid(mstrWaitSend, 2)
    End If

End Function
Private Sub SaveLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------
    
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    If Dir(App.Path & "\����.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    strOutput = Replace(strOutput, Chr(&HD), "<CR>")
    strOutput = Replace(strOutput, Chr(&HA), "<LF>")
    objStream.WriteLine (strDate & ":   " & strOutput)
    objStream.Close
    Set objStream = Nothing
End Sub
Private Function GetCheckCode(ByVal str_in As String) As String
    '����У����
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function
