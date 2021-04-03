VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Immage_800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: Immage800
'����:������
'����:����
'�û�:��������Ժ,��������ҽԺ
'2012-10-17 �¶� ����˫��
Private mstrCode As String
Private mstrWaitSend As String

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    'GetStartCmd = Chr(6)
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
    
        Dim varLine As Variant
        Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
        Dim strValues As String
        Dim blnEnd As Boolean '�Ƿ��ѽ�����

        Dim strNowData As String
        Dim strData As String
        Dim intloop As Integer
        Dim varItem As Variant
        Dim strDate As String, strSampleNo As String
        Dim lng_S As Long, lng_E As Long
        Dim strItem As String
        Dim varTmp As Variant
        On Error GoTo hErr
        '��ʼ������ֵ
100     strResult = "": strReserved = "": strCmd = ""
    
102     If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
104     If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)
    
106     mstrCode = mstrCode & strSource
    
        'û���ҵ���ʼ���ͽ�����ʱ�˳�
108     lng_S = InStr(mstrCode, Chr(&H5))  '-"H|\^&"
    '110     lng_S = lng_S - 2
110     lng_E = 0
112     If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    '116         lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '"L|1|N"
    '118         If lng_E > 0 Then
    '120             lng_E = InStr(lng_E, mstrCode, Chr(&HA))
    '            Else
    '122             lng_E = 0
    '            End If
    '        End If
    
114     If Not (lng_S > 0 And lng_E > 0) Then Exit Sub
116     SaveLog "800", "������������ʼ����", strCmd
        'ȡ�õ�ǰһ�Σ�ʣ�µķ��ء��´��ٽ�����
118     strNowData = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
120     mstrCode = Mid(mstrCode, lng_E + 1)
        
        '��ʼ����
122     varLine = Split(strNowData, vbCrLf)
124     strItem = ""
126     For intloop = LBound(varLine) To UBound(varLine)
128         If InStr(varLine(intloop), "|") > 0 Then
130             varItem = Split(varLine(intloop), "|")
132             strTmp = UCase(Right(varItem(0), 1))
134             Select Case strTmp
                    Case "H"
136                     strTmp = Trim(varItem(13))
138                     strDate = ""
140                     If Len(strTmp) >= 14 Then
142                         strDate = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & _
                                      Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                        End If
144                 Case "Q"
146                     strTmp = varItem(2)
148                     varTmp = Split(strTmp, "\")
150                     For iRow = LBound(varTmp) To UBound(varTmp)
152                         strTmp = varTmp(iRow)
154                         If IsNumeric(Mid(strTmp, 2)) And Mid(strTmp, 2) >= 12 Then
156                             strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Mid$(strTmp, 2)
                            End If
                        Next
158                     If strResult <> "" Then strCmd = "1|"
160                     SaveLog "800", "˫��" & strResult, strCmd
162                 Case "O"                '�걾��
164                     If Not IsDate(strDate) Then
166                         strTmp = Trim(varItem(7))
168                         strDate = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & _
                                      Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                        End If
                        
170                     strSampleNo = Trim(Split(varItem(2), "^")(0))
172                     If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
174                     strData = strDate & "|" & strSampleNo & "| |Blood|0"
176                 Case "R"
178                     strItem = strItem & "|" & Split(varItem(2), "^")(3) & "|" & varItem(3)
                End Select
            End If
        Next
180     If strItem <> "" And strData <> "" Then
182         strResult = strResult & "||" & strData & strItem
        End If
        
184     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
186     SaveLog "immage800", CStr(Erl()) & "," & Err.Description, ""
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    Dim varInData As Variant, iLoop As Integer, strData As String, i As Integer
    Dim strBarcode As String, strEmergency As String, strTmp As String
    Dim strChannel As String, strSend As String, strItems As String, iRow As Integer
    Dim strSegment As String, strDate As String, strSampleType As String
    Dim strSampeNo As String
    On Error GoTo hErr

100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '�ս���˫��ģʽ����û�����ɷ�������

106     SaveLog "Immage800-����˫�����ɴ���������", strSample, ""
108     varInData = Split(strSample, "||")
110     For iLoop = LBound(varInData) To UBound(varInData)
112         strData = varInData(iLoop)
114         strDate = Format(CDate(Split(strData, "|")(0)), "yyyyMMddHHmmss")
116         strSampeNo = Trim(Split(strData, "|")(1))
            '--�걾 S'=Ѫ��,'U' =�����,'F' =�Լ���Һ,��P�� =Ѫ��,'T' =��ʱ��,ע����Ĭ�ϲ������ʹ�ϵͳ����
118         strSampleType = Trim(Split(strData, "|")(3))
120         If InStr(strSampleType, "Ѫ��") > 0 Then
122             strSampleType = "P"
124         ElseIf InStr(strSampleType, "��ʱ��") > 0 Then
126             strSampleType = "T"
128         ElseIf InStr(strSampleType, "��") > 0 Then
130             strSampleType = "U"
132         ElseIf InStr(strSampleType, "Ѫ��") > 0 Then
134             strSampleType = "S"
            Else
136             strSampleType = ""
            End If
            '--����
138         strBarcode = Trim(Split(strData, "|")(5))
            '---����
140         If Val(Split(strData, "|")(4)) = 1 Then
142             strEmergency = "S"
            Else
144             strEmergency = "R"
            End If
            '----ָ��
         
146         strItems = ""
148         For i = 10 To UBound(Split(strData, "|"))
150             strChannel = Trim$("" & Split(strData, "|")(i))
152             strItems = strItems & "^^^" & strChannel & "\"
            Next
        
154         strSend = ""
156         If strItems <> "" And strBarcode <> "" Then
158             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
160             iRow = 0
162             strSegment = Chr(&H5)
164             iRow = iRow + 1
166             strSend = iRow & "H|\^&" & Chr(&HD) & Chr(&H3)
168             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                
                'P|1||A43749||Smith^Tom^J||19950315^120^H|�Ա�M/F/U|||||ҽ��||||||||||||λ��<CR>
                'P|1||�걾��|<CR>
170             iRow = iRow + 1
172             strSend = iRow & "P|1||" & strSampeNo & "|" & Chr(&HD) & Chr(&H3)
174             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                
                'O|1|SID_00002^3^1^^1||^^^IGG^^N^2\^^^IGA|S|||19950316135300|||N||||S|||1^1.00|||||||O<CR>
176             iRow = iRow + 1
178             strSend = iRow & "O|1|" & strBarcode & "||" & strItems & "|" & strEmergency & "||" & _
                          IIf(InStr("T,U", strSampleType) > 0, strDate, "") & _
                          "||||N||||" & strSampleType & "|||1^1.00" & Chr(&HD) & Chr(&H3)
180             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                
182             iRow = iRow + 1
184             strSend = iRow & "L|1|N" & Chr(&HD) & Chr(&H3)
186             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
                
188             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment
            End If

190         SendSample = GetSendData
192         If mstrWaitSend <> "" Then iSendStep = 1    '����δ�������ݣ����˳�˫��ģʽ

        Next
194     SaveLog "image800-����˫�����ɴ���������", "����:" & SendSample, iSendStep & "����:" & mstrWaitSend
    Else
        '�յ�����Ӧ�𣬴ӻ�����ȡ��ָ���
196     SaveLog "image800-˫�򣬷������� ", strResponse, iSendStep
198     If InStr(strResponse, Chr(&H6)) > 0 Then

200         SendSample = GetSendData
202         If mstrWaitSend = "" Then iSendStep = 0
204         SaveLog "image800-˫�򣬷�������", "����:" & SendSample, iSendStep & "����:" & mstrWaitSend
206     ElseIf strResponse = Chr(&H5) Then
208         SaveLog "image800-˫�򣬷�������", "�յ�05:" & SendSample, iSendStep & "����:" & mstrWaitSend
210         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
212             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
214             mstrWaitSend = ""
            End If
216         iSendStep = 0
218         SendSample = Chr(&H6)
        End If

    End If
    Exit Function
hErr:
220     SaveLog "image800-˫��", strResponse, CStr(Erl()) & "��," & Err.Description
End Function


Private Sub SaveLog(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
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
    objStream.WriteLine (String(50, "��"))
    objStream.WriteLine ("ִ��ʱ��:" & strDate & "�汾:" & App.major & "." & App.minor & "." & App.Revision)
    objStream.WriteLine ("Func:" & strFunc)
    objStream.WriteLine ("In:" & StrInput)
    objStream.WriteLine ("Out:" & strOutput)
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
