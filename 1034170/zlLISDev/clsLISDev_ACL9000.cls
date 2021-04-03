VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACL9000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ACL9000
'����:�������������
'����:�¶�
'�û�:��������ҽԺ(����)������ҽԺ(˫��)
'�޸���:������
'�޸�����:2012-12-5

Private mstrCode As String
Private mstrWaitSend As String

Public Function GetStartCmd() As String
'�豸��׼�ӿ�
'���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
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

    Dim varLine As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, lngLength As Long
    Dim str�걾���� As String
    Dim blnEnd As Boolean    '�Ƿ��ѽ�����
    Dim i As Integer, str_Line As String
    On Error GoTo hErr

    '��ʼ������ֵ
100 strResult = "": strReserved = "": strCmd = ""
102 If InStr(strSource, Chr(&H5)) > 0 Then
104     strCmd = Chr(&H6)
106 ElseIf InStr(strSource, Chr(&HA)) > 0 Then
108     strCmd = Chr(&H6)
    End If

110 str�걾���� = Trim(ReadIni("ACL9000", "�걾����", App.Path & "\ACL9000.ini"))
112 If str�걾���� = "" Then str�걾���� = "ѪҺ": Call WriteIni("ACL9000", "�걾����", str�걾����, App.Path & "\ACL9000.ini")
    
114 mstrCode = mstrCode & strSource

116 If mstrCode Like "***" Then
118     blnEnd = False
        '        strReserved = strSource
120     Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
122         lngLength = InStr(mstrCode, "") - InStr(mstrCode, "")
124         strLine = Mid(mstrCode, InStr(mstrCode, ""), lngLength)
126         mstrCode = Mid(mstrCode, InStr(mstrCode, "") + 1)

128         If strLine Like "*" & Chr(&H2) & "?Q|*" Then                 '˫��ͨѶ
130             varLine = Split(strLine, vbCrLf)
132             For i = LBound(varLine) To UBound(varLine)
134                 str_Line = varLine(i)
136                 If str_Line Like Chr(&H5) & Chr(&H2) & "?H|*" Then
138                     str���� = Split(str_Line, "|")(13)
140                     If IsNumeric(str����) And Len(str����) >= 12 Then
142                         str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2) & " " & Mid(str����, 9, 2) & ":" & Mid(str����, 11, 2)
                        Else
144                         str���� = Format(Now, "yyyy-MM-dd HH:mm")
                        End If
146                 ElseIf str_Line Like Chr(&H2) & "?Q|*" Then
148                     If InStr(str_Line, "^") > 0 Then
150                         str�걾�� = Split(Split(str_Line, "|")(2), "^")(1)
                        End If
                    End If
                Next

152             If Val(str�걾��) <> 0 Then
154                 strResult = strResult & "||" & str���� & "|^^" & Trim(str�걾��)
                End If

156             If strResult <> "" Then
158                 strCmd = "1|" & strCmd                                      '˫��ָ��
                End If

160             If Not mstrCode Like "" Then
162                 blnEnd = True
                End If

164         ElseIf strLine Like "*" & Chr(&H2) & "?R|*" Then             '�������

166             str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
                '����
168             varLine = Split(strLine, vbCrLf)
170             strItem = ""
172             For iRow = LBound(varLine) To UBound(varLine)
174                 If InStr(varLine(iRow), "|") > 0 Then
176                     varItem = Split(varLine(iRow), "|")
178                     If UBound(varItem) >= 12 Then
180                         If Trim("" & varItem(0)) Like "*O" And Left(varItem(4), 3) = "^^^" Then    '�걾����
182                             str��Ŀ = Replace(varItem(4), "^^^", "")
                                'һ���µı걾��ʼ,�ȱ����ϴεĽ��
184                             If strItem <> "" And str�걾�� <> "" Then
                                    If Len(str�걾��) >= 12 Then
                                        strResult = strResult & "||" & str���� & "|^^" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                                    Else
                                        strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                                    End If
                                End If

188                             str�걾�� = Trim("" & varItem(2)): str�����˱걾�ʿ� = " |" & str�걾���� & "|" & IIf(varItem(11) = "Q", "1", "0")
190                             strItem = ""

                            End If
192                         If InStr(varItem(2), "^^^") > 0 And Trim(Replace(varItem(2), "^^^", "")) = Trim(str��Ŀ) Then
194                             If Trim(varItem(3)) <> "" Then
                                    '�����
196                                 strItem = strItem & "|" & varItem(1) & Replace(varItem(2), "^^^", "") & "|" & varItem(3)
198                                 str���� = Mid(varItem(12), 1, 4) & "-" & Mid(varItem(12), 5, 2) & "-" & Mid(varItem(12), 7, 2) & " " & Mid(varItem(12), 9, 2) & ":" & Mid(varItem(12), 11, 2) & ":" & Mid(varItem(12), 13, 2)
                                End If
                            End If
                        End If
                    End If
                Next

200             If strItem <> "" And str�걾�� <> "" Then
202                 If Len(str�걾��) >= 12 Then
204                     strResult = strResult & "||" & str���� & "|^^" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                    Else
206                     strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                    End If
208                 If Not mstrCode Like "" Then
210                     blnEnd = True
                    End If
                End If
            End If
        Loop

    End If
212 If strResult <> "" Then
214     If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
hErr:
216 SaveLog "Analyse��" & CStr(Erl()) & "��," & Err.Description
End Sub


Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        Dim varSampleInfo As Variant
        Dim strNo As String
        Dim strItems As String
        Dim strLine As String
        Dim intloop As Integer
        Dim str_Line As String
        Dim strSendCmd As String
        Dim i As Integer
        Dim intNO As Integer
        Dim j As Integer

100     intNO = 2
102     varSampleInfo = Split(strSampleInfo, "||")
104     SendSample = ""
106     blnSuccess = True

108     If iSendStep = 0 Then
110         For intloop = LBound(varSampleInfo) To UBound(varSampleInfo)
112             strLine = varSampleInfo(intloop)
114             If Val(Split(strLine, "|")(5)) > 0 Then
116                 strNo = Split(strLine, "|")(5)
                Else
118                 strNo = Split(strLine, "|")(1)
                End If

120             If strNo <> "" And UBound(Split(strLine, "|")) >= 10 Then
                    'ָ�ʼ
122                 str_Line = "1H|\^&||||||||ACL9000||P|1|" & Format(Now, "yyyyMMddHHmmss") & Chr(&HD) & Chr(&H3)
124                 strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                    '������Ϣ
126                 str_Line = "2P|1|||||||U||||||||||||||||||||||||||" & Chr(&HD) & Chr(&H3)
128                 strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
130                 strItems = ""
132                 For i = 10 To UBound(Split(strLine, "|"))
134                     If InStr(strItems & ",", "," & Mid(Split(strLine, "|")(i), 2) & ",") <= 0 Then
136                         strItems = strItems & "," & Mid(Split(strLine, "|")(i), 2)
138                         intNO = intNO + 1
140                         j = j + 1
                            '����ָ��
142                         str_Line = intNO & "O|" & j & "|" & strNo & "||^^^" & Mid(Split(strLine, "|")(i), 2) & "|||||||||||^|||||||||||O||||||" & Chr(&HD) & Chr(&H3)
144                         strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                        End If
                    Next
                    '����
146                 intNO = intNO + 1
148                 str_Line = intNO & "L|1|N" & Chr(&HD) & Chr(&H3)
150                 strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","

152                 iSendStep = 1    '��ʼͨ��
154                 mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & strSendCmd & Chr(&H4)

                End If
            Next
156         SendSample = GetSendData
        Else
158         If InStr(strResponse, Chr(&H6)) > 0 Then SendSample = GetSendData            '�յ�06
        End If
160     If mstrWaitSend = "" Then iSendStep = 0 '�˳�˫��ͨ
        Exit Function
hErr:
162     SaveLog "SendSample��" & CStr(Erl()) & "��," & Err.Description
End Function

Private Function GetSendData() As String
    If mstrWaitSend <> "" Then
        If Left(mstrWaitSend, 1) = "," Then
            mstrWaitSend = Mid(mstrWaitSend, 2)
        End If
        If InStr(mstrWaitSend, ",") > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
            mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
        Else
            GetSendData = mstrWaitSend
            mstrWaitSend = ""
        End If
    End If

End Function
Private Function GetCheckCode(ByVal str_in As String) As String
'��ȡУ����
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
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
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"

    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub
