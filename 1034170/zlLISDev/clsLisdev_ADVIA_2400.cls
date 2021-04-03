VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisdev_ADVIA_2400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ADVIA 2400 ȫ�Զ���ѧ������
'����:�ݶ� ������
'����:
'�û�:��������ҽԺ

Private mstrCode As String      '����������
Private mstrWaitSend As String  '����������
'Private mstrRackNo As String    '��¼�����Ӧ���̺ţ����ŵ���Ϣ
Private mstrEndBarcode As String '��¼ÿ��ָ�����һ������
Private mintRow As Integer       '˳���

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
 
'    mstrRackNo = ""
    mstrWaitSend = ""
    mstrCode = ""
    mstrEndBarcode = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function
 
Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
        
        '���ܣ���������
        '������
        'strSource��Ҫ������ԭʼ��
        'strResult�����صļ�����(����������������밴���±�׼��֯���)
        '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
        '   ��0��Ԫ�أ�����ʱ��
        '   ��1��Ԫ�أ��������^�Ƿ���^�����  ����ֻ���������
        '   ��2��Ԫ�أ�������
        '   ��3��Ԫ�أ��걾
        '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
        '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
        '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
        'strReserved���������ɽ�����ԭʼ������Ҫ��������������
        'strCmd�������Ҫ���ɷ������豸���͵�����
        Dim strCode As String, strQC As String, strSampleNo As String, strChannel As String, strValue As String
        Dim varCode As Variant, iRow As Integer, strLine As String, intStart As String, strItems As String
        Dim lngS As Long, lngE As Long, i As Integer, strDate As String
        Dim iCount As Integer, strSendBlock As String
        
        On Error GoTo hErr
100     strCmd = "": strResult = ""
102     mstrCode = mstrCode & strSource
104     If InStr(strSource, Chr(&H5)) > 0 Then
106         strCmd = Chr(&H6)
108     ElseIf InStr(strSource, Chr(&HA)) > 0 Then
110         strCmd = Chr(&H6)
        End If
    
112     lngE = 0
114     lngS = InStr(mstrCode, Chr(&H5))
116     If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
        
118     If Right(strSource, 1) = Chr(&H4) Then
            '�յ�һ�����������ݣ���ʱ�������δ�������ݣ�����˫��ģʽ
120         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
122             Call SaveLog(" ��������δ�������ݣ��л���˫��ģʽ,����" & mstrWaitSend & ",���" & mintRow)
                
124             If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
126                 mstrWaitSend = Mid$(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
                Else
128                 mstrWaitSend = Chr(&H5) & "|" & mstrWaitSend
130                 mintRow = 0
                End If
132             strCmd = "1|"
134             strResult = Format(Now, "YYYY-MM-DD") & "|^^0"
                Exit Sub
            End If
        End If
        
136     Do While lngS > 0 And lngE > 0
138         strCode = Mid(mstrCode, lngS, lngE - lngS + 1)
140         mstrCode = Mid$(mstrCode, lngE + 1)
            '-----------------------------------------------------
142         strQC = "0"
144         strSampleNo = ""
146         strDate = ""
148         strItems = ""
150         intStart = -1
152         iCount = 0
154         strValue = ""
156         varCode = Split(strCode, Chr(&H2))

158         For iRow = LBound(varCode) To UBound(varCode)
160             strLine = varCode(iRow)
162             If Mid$(strLine, 2, 1) = "R" Then
164                 If Not (Mid$(strLine, 44, 1) = "M" Or Mid$(strLine, 44, 1) = "D" Or Mid$(strLine, 44, 1) = "U") Then
166                     If strSampleNo <> "" And strDate <> "" And strItems <> "" Then
168                         If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
170                         strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|" & strQC & strItems
172                         strQC = "0"
174                         strSampleNo = ""
176                         strDate = ""
178                         strItems = ""
180                         intStart = -1
                        
                        End If
182                     If Mid$(strLine, 19, 1) = "N" Then
                            '--��ͨ�걾
184                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2)
186                         strSampleNo = Trim$(Mid$(strLine, 21, 13))
188                         intStart = 91
190                     ElseIf Mid$(strLine, 25, 1) = "N" Then
                            '--��ͨ�걾
192                         strSampleNo = Trim$(Mid$(strLine, 27, 13))
194                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2) & " " & Mid$(strLine, 19, 2) & ":" & Mid$(strLine, 21, 2) & ":" & Mid$(strLine, 23, 2)
196                         intStart = 97
198                     ElseIf Mid$(strLine, 19, 1) = "I" Then
                            '---����
200                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2)
202                         strSampleNo = Trim$(Mid$(strLine, 21, 13)) & "^1^"
204                         intStart = 91
206                     ElseIf Mid$(strLine, 25, 1) = "I" Then
                            '---����
208                         strSampleNo = Trim$(Mid$(strLine, 27, 13)) & "^1^"
210                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2) & " " & Mid$(strLine, 19, 2) & ":" & Mid$(strLine, 21, 2) & ":" & Mid$(strLine, 23, 2)
212                         intStart = 97
214                     ElseIf Mid$(strLine, 19, 1) = "C" Then
                            '---�ʿ�
216                         strQC = "1"
218                         strSampleNo = Trim$(Mid$(strLine, 21, 13))
220                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2)
222                         intStart = 91
224                     ElseIf Mid$(strLine, 25, 1) = "C" Then
                            '---�ʿ�
226                         strQC = "1"
228                         strSampleNo = Trim$(Mid$(strLine, 27, 13)) & "^1^"
230                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2) & " " & Mid$(strLine, 19, 2) & ":" & Mid$(strLine, 21, 2) & ":" & Mid$(strLine, 23, 2)
232                         intStart = 97
                        End If
                        '3+1+8+3=15
234                     For i = intStart To Len(strLine) Step 15
236                         If i + 15 > Len(strLine) Then Exit For
238                         strChannel = Trim$(Mid$(strLine, i, 3))
240                         strValue = Trim$(Mid$(strLine, i + 4, 8))
                        
242                         If strChannel <> "" And IsNumeric(strValue) Then strItems = strItems & "|" & strChannel & "|" & strValue
                        Next
                    Else
244                     If intStart = 91 Then
246                         intStart = 41
248                     ElseIf intStart = 97 Then
250                         intStart = 47
                        End If
                    
252                     For i = intStart To Len(strLine) Step 15
254                         If i + 15 > Len(strLine) Then Exit For
256                         strChannel = Trim$(Mid$(strLine, i, 3))
258                         strValue = Trim$(Mid$(strLine, i + 4, 8))
                        
260                         If strChannel <> "" And IsNumeric(strValue) Then strItems = strItems & "|" & strChannel & "|" & strValue
                        Next
                    
                    End If
262             ElseIf Mid$(strLine, 2, 1) = "Q" And Mid$(strLine, 10, 1) = "0" Then
                    '˫��,
                    '------
264                 SaveLog "2400-˫���ѯ����:" & strLine
266                 strValue = ""
268                 If Mid$(strLine, 1, 1) = "1" Then iCount = 0
270                 For i = 11 To Len(strLine) Step 13
272                     If i + 13 > Len(strLine) Then Exit For
274                     strValue = Trim(Mid$(strLine, i, 13))
276                     If Len(strValue) >= 12 Then
278                         strCmd = "1|"
280                         iCount = iCount + 1
282                         If iCount = 8 Then iCount = 0

    '274                         mstrRackNo = mstrRackNo & "," & Mid$(strLine, 4, 6) & "|" & iCount
284                         strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & strValue
286                         If Mid$(strLine, 1, 1) = "1" And i = 11 Then
288                             SaveLog "------------2400˫��ʼ������05,��ʼ����," & strValue
290                             mstrWaitSend = mstrWaitSend & "|" & Chr(&H5)
292                             mintRow = 0
                            End If
                        Else
294                         SaveLog "2400��" & strValue & " С��12λ����������"
                            
                        End If
                    Next
                    
    '286                 If mstrRackNo <> "" Then mstrRackNo = Mid$(mstrRackNo, 2)
296             ElseIf Mid$(strLine, 2, 1) = "Q" And Mid$(strLine, 10, 1) <> "0" Then
                    
298                 SaveLog "2400-˫���ѯ����(������),Ӧ��:04" & strLine
300                 'strCmd = Chr(&H4)   '������ֹͨѶ����
                    '--
                End If
            Next
            
302         If strValue <> "" And strCmd = "1|" Then
304             If Len(strValue) >= 12 Then
306                 If InStr(mstrEndBarcode, "|" & strValue) <= 0 Then mstrEndBarcode = mstrEndBarcode & "|" & strValue
308                 SaveLog "2400˫�򻺴��δλ����Ϊ" & mstrEndBarcode '& ",����ļܺ�Ϊ" & mstrRackNo
                End If
            End If
            
310         If strSampleNo <> "" And strDate <> "" And strItems <> "" Then
312             If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
314             strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|" & strQC & strItems
            End If
            '-----------------------------------------------------
316         lngE = 0
318         lngS = InStr(mstrCode, Chr(&H5))
320         If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
        Loop
    
322     If strResult <> "" Then strResult = Mid$(strResult, 3)
        
        Exit Sub
hErr:
    SaveLog "2400-��" & CStr(Erl()) & "��,���ִ�����ʾ," & Err.Description & ",�յ�������Ϊ" & strSource

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

    
    Dim strTmp As String
    On Error GoTo hErr
    
100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '�ս���˫��ģʽ����û�����ɷ�������
106     If strSample = "" Then
108         SendSample = GetSendData
110         If mstrWaitSend <> "" Then iSendStep = 1    '����δ�������ݣ����˳�˫��ģʽ
112         SaveLog "2400-�ӵ���ֱ���л���" & IIf(iSendStep = 0, "�˳�˫��", "����˫��") & "״̬,���η���" & SendSample & ",δ��������:" & mstrWaitSend
        Else
114         If GetSendCmd(strSample) Then
116             SendSample = GetSendData
118             If mstrWaitSend <> "" Then iSendStep = 1    '����δ�������ݣ����˳�˫��ģʽ
            End If
120         SaveLog "2400-�������ݽ�����" & IIf(iSendStep = 0, "�˳�˫��", "����˫��") & "״̬,���η���" & SendSample & ",δ��������:" & mstrWaitSend
        End If

    Else
        '�յ�����Ӧ�𣬴ӻ�����ȡ��ָ���
122     If InStr(strResponse, Chr(&H6)) > 0 Then

124         SendSample = GetSendData
126         If mstrWaitSend = "" Then iSendStep = 0
128         strTmp = "06 "
130     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then

132         SendSample = GetSendData
134         If mstrWaitSend = "" Then iSendStep = 0
136         strTmp = "15 Err "
138     ElseIf InStr(strResponse, Chr(&H11)) > 0 Then

140         SendSample = GetSendData
142         If mstrWaitSend = "" Then iSendStep = 0
144         strTmp = "11 Next "
146     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
            '��ʱ-����δ���͵ı걾��
148         mstrWaitSend = ""
150         mstrEndBarcode = ""
152         iSendStep = 0
            
154         strTmp = "04 ��ֹ"
156     ElseIf InStr(strResponse, Chr(&H5)) > 0 Then
            '�յ�05����ͣ�����˫���䣬���뵥��ģʽ
            '���δ���͵�����
    '152         mstrWaitSend = ""
    '156         mstrEndBarcode = ""
158         If mstrWaitSend <> "" Then
160             mstrWaitSend = Chr(&H5) & "|" & mstrWaitSend
162             mintRow = 0
            End If
164         iSendStep = 0
166         strTmp = "05 ����"
168         mstrCode = mstrCode & Chr(&H5)
170         SendSample = Chr(&H6)
        Else
172         strTmp = strResponse
        End If
174     strTmp = "2400-����˫���յ�" & strTmp & "," & IIf(iSendStep = 0, "-----------�˳�˫��", "����˫��") & "״̬"
176     If SendSample <> "" Then
178         strTmp = strTmp & ",���η���:" & SendSample
        Else
180         strTmp = strTmp & ",û�����ݷ���"
        End If
182     If mstrWaitSend <> "" Then
184         strTmp = strTmp & ",����������:" & mstrWaitSend
        Else
186         strTmp = strTmp & ",�����ѷ���"
        End If
188     SaveLog strTmp
    End If
    Exit Function
hErr:
190     SaveLog "2400-����˫���" & CStr(Erl()) & "��,���ִ�����ʾ" & Err.Description & ",�յ�:" & strResponse
End Function

Private Function GetSendCmd(ByVal strSample As String) As Boolean
    '���ɴ��������ݣ���������һ�����򷵻�True
    Dim varInData As Variant, iLoop As Integer, strData As String
    Dim strNo As String, strBlood As String, strEmergency As String, strRackNo As String
    Dim strChannel As String, strSend As String, iCount As Integer, iFlag As Integer
    Dim iRow As Integer, i As Integer
    
106     SaveLog "2400-׼�����ɴ���������," & strSample
108     varInData = Split(strSample, "||")
110     For iLoop = LBound(varInData) To UBound(varInData)
112         strData = varInData(iLoop)
            '--����
114         strNo = Trim(Split(strData, "|")(5))
            '--�걾����
116         strBlood = Split(strData, "|")(3)
118         If InStr(strBlood, "��") > 0 Then
120             strBlood = "2"
            Else
122             strBlood = "1"
            End If
            '---����
124         If Val(Split(strData, "|")(4)) = 1 Then
126             strEmergency = "I"
            Else
128             strEmergency = "N"
            End If
            '---�̺�
            strRackNo = "0101" & Format(UBound(Split(strData, "|")) - 9, "000")
            '----ָ��
154         iCount = 0: strSend = "": iFlag = 0

156         For i = 10 To UBound(Split(strData, "|"))
158             If iCount = 0 Then
                        '�״�����-ͷ
160                 strSend = "O " & strRackNo & strEmergency & "0" & strNo & String$(13 - Len(strNo), " ") & Space$(7) & Space$(16) & Space$(16)
162                 strSend = strSend & "M000" & Format(CDate(Split(strData, "|")(0)), "yyyyMMdd") & " 1.0" & strBlood & "1"
                End If
                    '--- ͨ����Ϊ3λ
164             strChannel = GetChannel(Trim$(Split(strData, "|")(i)))
166             If strChannel <> "" Then strSend = strSend & strChannel & "M"

168             iCount = iCount + 1
            Next
            
170         If strSend <> "" Then mstrWaitSend = mstrWaitSend & "|" & strSend

180         If mstrEndBarcode = "" Then SaveLog "2400-������뻺��Ϊ��,��ǰ������������" & strNo
            
            '---���Ρ���ָ���е����һ�����룬�����
192         If InStr(mstrEndBarcode & "|", "|" & strNo & "|") > 0 Then
                If strSend <> "" Then mstrWaitSend = mstrWaitSend & "|" & Chr(&H4)
194             SaveLog "2400-���һ�����룬����04,��:" & mstrEndBarcode & "�����" & strNo
196             mstrEndBarcode = Replace(mstrEndBarcode, "|" & strNo, "")
                GetSendCmd = True
            End If
        Next
        
End Function

Private Function GetChannel(ByVal strIn As String) As String
    '����ͨ����
    
    If Len(Trim(strIn)) <= "3" Then
        GetChannel = String(3 - Len(Trim(strIn)), "0") & Trim(strIn)
        If GetChannel = Space(3) Then
            SaveLog "2400-˫��ͨ����Ϊ��," & strIn
            GetChannel = ""
        End If
    Else
        SaveLog "2400-˫��ͨ���볬��," & strIn
    End If

End Function
Private Function GetSendData() As String
    '�ӻ����л�ȡһ�����ݣ����ء����������û�����ݣ����ؿմ�
    
    Dim strSend As String
    
    If Left$(mstrWaitSend, 1) = "|" Then mstrWaitSend = Mid$(mstrWaitSend, 2)
    If mstrWaitSend = "" Then Exit Function
    If InStr(mstrWaitSend, "|") <= 0 Then
        strSend = mstrWaitSend
        mstrWaitSend = ""
    Else
        strSend = Mid$(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
        mstrWaitSend = Mid$(mstrWaitSend, InStr(mstrWaitSend, "|") + 1)
    End If
    
    If Len(strSend) > 1 Then
        mintRow = mintRow + 1
        If mintRow = 8 Then mintRow = 0
        strSend = mintRow & strSend & " " & Chr(&H3)   '���һ�����ݹ̶�Ϊ 03
        strSend = Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
    
    End If
    GetSendData = strSend
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
    strFileName = App.Path & "\LisDev_" & Format(Date, "yyyyMMdd") & ".LOG"
    
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