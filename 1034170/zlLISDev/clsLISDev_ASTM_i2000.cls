VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ASTM_i2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: Axsym i2000
'����: ����
'����: �¶�
'�û�: ������ҽԺ
'����: 2012-10-25
'�޸ļ�¼:
'
Private mstrData As String
Private mstrWaitSend As String

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

    Dim lngPos As Long
    Dim lngSeqi As Long
    Dim lngCount As Long
    Dim lngCount1 As Long
    Dim lngSeqj As Long
    Dim lngSeqjcnt As Long
    Dim lngfn As Long

    
    Dim varReply As Variant
    Dim strReplyMsg As String
    Dim strInstrReply As String
    Dim strChecksum As String
  
    
    Dim varRecord As Variant
    Dim varField As Variant
    Dim lngS As Long, lngE As Long  '��ʼ������λ��
    Dim strData As String
    Dim strLine As String
    Dim strBarcode As String    '����
    Dim strSampleNo As String   '�걾��
    Dim strQC As String         '�Ƿ��ʿ�
    Dim strDate As String       '����
    Dim strValue As String, strChannel As String '�����ͨ����
    Dim strItem As String
    
    On Error GoTo hErr
    strData = strSource
    strResult = ""
    
    If strData <> "" Then
        '---   Ӧ�𲿷� ----
        Call SaveLog("�յ��ַ�: " & strData)
        mstrData = mstrData & strData
        If InStr(strData, Chr(&H5)) > 0 Then
            strCmd = Chr(&H6)
            Call SaveLog("���յ��������͹�����ͨѶ�����ENQ��")
        ElseIf InStr(strData, Chr(&H3)) > 0 Then
            strCmd = Chr(&H6)
            Call SaveLog("�յ�ETX��Ӧ��ACK�󣬵ȴ�ͨѶ������EOT��")
        ElseIf InStr(strData, Chr(&H17)) > 0 Then
            strCmd = Chr(&H6)
            Call SaveLog("�յ�ETB��Ӧ��ACK�󣬼����ȴ���һ֡��")
        End If
    
    
        lngS = InStr(mstrData, Chr(&H5))
        lngE = 0
        If lngS > 0 Then lngE = InStr(lngS, mstrData, Chr(&H4))
        
        Do While lngS > 0 And lngE > 0
        
            Call SaveLog("���յ�һ����������")
            strData = Mid(mstrData, lngS, lngE - lngS + 1)
            mstrData = Mid(mstrData, lngE + 1)
            
            Do While strData Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
                lngS = InStr(strData, Chr(&H17))
                lngE = InStr(lngS, strData, Chr(&H2)) + 1
                strData = Mid(strData, 1, lngS - 1) & Mid(strData, lngE + 1)
            Loop
            
            '------------------------------------------------------------------

            Call SaveLog("��ʼ����")
            varRecord = Split(strData, Chr(&HD))
            strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
            
            For lngSeqi = LBound(varRecord) To UBound(varRecord)
                strLine = varRecord(lngSeqi)
    
                If strLine Like "*" & Chr(&H2) & "?Q|*" Then
                    varField = Split(strLine, "|")
                    strBarcode = Trim("" & varField(2))
                    Call SaveLog("�������롰" & strBarcode & "��������ָ���ʼ��ȡ������Ŀ��Ϣ��")
                    If Left(strBarcode, 1) = "^" Then strBarcode = Mid(strBarcode, 2)
                    
                    strCmd = "1|"
                    strResult = Format(Now, "yyyy-MM-dd") & "|^^" & strBarcode
                    
                    Exit Sub
                ElseIf strLine Like "*" & Chr(&H2) & "?O|*" Then
                    varField = Split(strLine, "|")
                    strSampleNo = Trim("" & varField(2))
                    If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
                    strQC = Trim("" & varField(11))
                    If strQC = "Q" Then
                        strQC = "1"
                    Else
                        strQC = "0"
                    End If
                ElseIf strLine Like "*" & Chr(&H2) & "?R|*" Then
                    varField = Split(strLine, "|")
                    strChannel = Trim("" & varField(2))
                    strValue = Trim("" & varField(3))
                    If strDate = "" Then
                        strDate = Trim("" & varField(12))
                        strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                    End If
                    If Right(strChannel, 2) = "^F" Then
                        strChannel = Split(strChannel, "^")(3)
                        If strChannel <> "" And strValue <> "" Then
                            strItem = strItem & "|" & strChannel & "|" & strValue
                        End If
                    End If
                End If
            
            Next
            
            If strItem <> "" And strDate <> "" And strSampleNo <> "" Then
                strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Serum|" & strQC & strItem
                Call SaveLog("���֡�" & strDate & "���걾��Ϊ��" & strSampleNo & "���ļ�������" & strItem & "��")
            End If
            '------------------------------------------------------------------
            '����Ƿ�������������
            lngS = InStr(mstrData, Chr(&H5))
            lngE = 0
            If lngS > 0 Then lngE = InStr(mstrData, Chr(&H4))
        Loop
    End If
    If strResult <> "" Then strResult = Mid$(strResult, 3)
    
    Exit Sub
hErr:
    Call SaveLog(CStr(Erl()) & "��," & Err.Description)
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
106     varInData = Split(strSample, "||")
108     For iLoop = LBound(varInData) To UBound(varInData)
110         strData = varInData(iLoop)
112         strDate = Format(CDate(Split(strData, "|")(0)), "yyyyMMddHHmmss")
114         strSampeNo = Trim(Split(strData, "|")(1))
            '--�걾
116         strSampleType = "" 'Trim(Split(strData, "|")(3))
            '--����
118         strBarcode = Trim(Split(strData, "|")(5))
            '---����
120         If Val(Split(strData, "|")(4)) = 1 Then
122             strEmergency = "S"
            Else
124             strEmergency = ""
            End If
            '----ָ��
126         strItems = ""
128         For i = 10 To UBound(Split(strData, "|"))
130             strChannel = Trim$("" & Split(strData, "|")(i))
132             strItems = strItems & "^^^" & strChannel & "\"
            Next
134         strSend = ""
136         If strItems <> "" And strBarcode <> "" Then
138             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
140             iRow = 0
142             strSegment = Chr(&H5)
144             iRow = iRow + 1
146             strSend = iRow & "H|\^&||||||||||P|1" & Chr(&HD) & Chr(&H3)
148             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                'P|1||A43749||Smith^Tom^J||19950315^120^H|�Ա�M/F/U|||||ҽ��||||||||||||λ��<CR>
                'P|1||�걾��|<CR>
150             iRow = iRow + 1
152             strSend = iRow & "P|1||" & strSampeNo & "||" & strSampeNo & "||||||||||||||||||||" & Chr(&HD) & Chr(&H3)
154             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                'O|1|SID_00002^3^1^^1||^^^IGG^^N^2\^^^IGA|S|||19950316135300|||N||||S|||1^1.00|||||||O<CR>
156             iRow = iRow + 1
158             strSend = iRow & "O|1|" & strBarcode & "||" & strItems & "|" & strEmergency & "||" & strDate & _
                          "||||A||||||||||||||Q" & Chr(&HD) & Chr(&H3)
160             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
162             iRow = iRow + 1
164             strSend = iRow & "L|1" & Chr(&HD) & Chr(&H3)
166             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
168             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment
            End If
170         SendSample = GetSendData
172         If mstrWaitSend <> "" Then iSendStep = 1    '����δ�������ݣ����˳�˫��ģʽ

        Next
174     If SendSample <> "" Then
176         SaveLog "˫�����ɴ���������,����:" & SendSample & "����:" & mstrWaitSend
        Else
178         SaveLog "˫�����ɴ���������,����:" & mstrWaitSend
        End If
        
    Else
        '�յ�����Ӧ�𣬴ӻ�����ȡ��ָ���

180     If InStr(strResponse, Chr(&H6)) > 0 Then
182         SendSample = GetSendData
184         If mstrWaitSend = "" Then iSendStep = 0
186         SaveLog "����Ӧ��06��ͨѶ�ɹ�������:" & SendSample & "����:" & mstrWaitSend
188     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
190         SendSample = GetSendData
192         If mstrWaitSend = "" Then iSendStep = 0
194         SaveLog "����Ӧ��04����ʾϣ����ͣ������Э���Ե���06,ͨѶ�ɹ�������:" & SendSample & "����:" & mstrWaitSend
196     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then
198         SendSample = GetSendData
200         If mstrWaitSend = "" Then iSendStep = 0
202         SaveLog "����Ӧ��15��ͨѶʧ�ܡ�����:" & SendSample & "����:" & mstrWaitSend
204     ElseIf strResponse = Chr(&H5) Then
206         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
208             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
210             mstrWaitSend = ""
            End If
212         Call SaveLog("����Ӧ��05��������������Ȩ�����η���ȡ��")
214         iSendStep = 0
216         SendSample = Chr(&H6)
        Else
218         Call SaveLog("����Ӧ��" & strResponse)
        End If
220     If mstrWaitSend = "" Then iSendStep = 0    '�ѷ����꣬�˳�˫��ģʽ
    End If
    Exit Function
hErr:
222     SaveLog "��" & CStr(Erl()) & "�д���" & Err.Description

End Function

'------- �������ڲ�����
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
        Call SaveLog("����ͨѶ������05������ͨѶ��ʼ��")
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H2) Then
        If InStr(1, mstrWaitSend, Chr(&HA)) > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(1, mstrWaitSend, Chr(&HA)))
            mstrWaitSend = Mid$(mstrWaitSend, InStr(1, mstrWaitSend, Chr(&HA)) + 1)
        End If
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H4) Then
        GetSendData = Chr(&H4)
        mstrWaitSend = Mid(mstrWaitSend, 2)
        Call SaveLog("����ͨѶ������04������ͨѶ������")
    End If
End Function

Private Sub SaveLog(ByVal StrInput As String)
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
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub
