VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Access2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: Access2 ȫ�Զ�΢���ӻ�ѧ�������߷�����
'����: BECKMAN COULTER
'����: �¶�
'�޸ģ�������
'�û�: �Ϻ�����ҽԺ(����0),����������ҽԺ��˫��
Private mstrCode As String
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
        Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, i As Integer, str_Line As String
        Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_Items As String
        Dim str_�������� As String
    
        On Error GoTo hErr
        'Ӧ����
100     If InStr(strSource, Chr(&H5)) > 0 Then
102         strCmd = Chr(&H6)
104     ElseIf InStr(strSource, Chr(&HA)) > 0 Then
106         strCmd = Chr(&H6)
        End If
108     If strSource <> "" Then mstrCode = mstrCode & strSource
    
110     lng_S = InStr(mstrCode, Chr(&H5))
112     If lng_S > 0 Then
114         lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
116         lng_E = InStr(mstrCode, Chr(&H4))
        End If
    
118     Do While lng_S > 0 And lng_E > 0
120         str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
122         mstrCode = Mid(mstrCode, lng_E + 1)
            '-------------------------------------------------------
124         str_Items = "": str_���� = "": str_�걾�� = ""
        
126         var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
128         For i = LBound(var_Code) To UBound(var_Code)
130             str_Line = var_Code(i)
132             If str_Line Like Chr(&H5) & Chr(&H2) & "*H|*" Then
                
134                 Call ReadConf(App.Path & "\Access2.ini", "ȡ����ʱ��", str_��������)
136                 If str_�������� = "1" Then
138                     str_���� = Format(Now, "yyyy-MM-dd HH:mm")
                    Else
140                     str_���� = Split(str_Line, "|")(13)
142                     If IsNumeric(str_����) And Len(str_����) >= 12 Then
144                         str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2)
                        Else
146                         str_���� = Format(Now, "yyyy-MM-dd HH:mm")
                        End If
                    End If
148             ElseIf str_Line Like Chr(&H2) & "*Q|*" Then '˫��ͨѶ
150                 str_�걾�� = Split(Split(str_Line, "|")(2), "^")(1)
152                 strCmd = "1|"
154             ElseIf str_Line Like Chr(&H2) & "*O|*" Then
156                 str_�걾�� = Trim(Split(str_Line, "|")(2))
158             ElseIf str_Line Like Chr(&H2) & "*R|*^*" Then
160                 str_��� = Trim(Split(str_Line, "|")(3))
162                 str_��Ŀ = Trim(Split(Split(str_Line, "|")(2), "^")(3))
164                 If str_��� <> "" And str_��Ŀ <> "" Then str_Items = str_Items & "|" & str_��Ŀ & "|" & str_���
                End If
            Next
        
166         If Len(Trim(str_�걾��)) >= 12 And Val(str_�걾��) > 0 Then
168             str_�걾�� = "^^" & str_�걾��
170         ElseIf Val(str_�걾��) <= 0 Then
172             str_�걾�� = ""
            End If
        
174         If str_�걾�� <> "" And IsDate(str_����) Then
176             strResult = strResult & "||" & str_���� & "|" & str_�걾��
178             If str_Items <> "" Then strResult = strResult & "| |Ѫ��|0" & str_Items
            End If
            '-------------------------------------------------------
180         lng_S = InStr(mstrCode, Chr(&H5))
182         If lng_S > 0 Then
184             lng_E = InStr(lng_S, mstrCode, Chr(&H4))
            Else
186             lng_E = InStr(mstrCode, Chr(&H4))
            End If
        Loop
188     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
190     SaveLog "Analyse ��" & CStr(Erl()) & "�д���" & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
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
    On Error GoTo hErr
100 varSampleInfo = Split(strSample, "||")
102 SendSample = ""
104 blnSuccess = True

106 If iSendStep = 0 Then
108     For intloop = LBound(varSampleInfo) To UBound(varSampleInfo)
110         strLine = varSampleInfo(intloop)
112         If Val(Split(strLine, "|")(5)) > 0 Then
114             strNo = Split(strLine, "|")(5)
            Else
116             strNo = Split(strLine, "|")(1)
            End If

118         If strNo <> "" And UBound(Split(strLine, "|")) >= 10 Then
                'ָ�ʼ
120             str_Line = "1H|\^&|||Host LIS|||||||P|1|" & Format(Now, "yyyyMMddHHmmss") & Chr(&HD) & Chr(&H3)
122             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                '������Ϣ
124             str_Line = "2P|1|" & Chr(&HD) & Chr(&H3)
126             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","

128             For i = 10 To UBound(Split(strLine, "|"))
130                 'strItems = Split(strLine, "|")(i)
                    '����ָ��
132                 strItems = strItems & "^^^" & Split(strLine, "|")(i) & "^1\"
                    
                Next
                If strItems <> "" Then strItems = Left$(strItems, Len(strItems) - 1)
134             str_Line = "3O|1|" & strNo & "||" & strItems & "|R||||||A||||Serum" & Chr(&HD) & Chr(&H3)
136             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                '����
138             str_Line = "4L|1|F" & Chr(&HD) & Chr(&H3)
140             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","

142             mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & strSendCmd & Chr(&H4)

            End If
        Next
144     SendSample = GetSendData
146     If mstrWaitSend <> "" Then iSendStep = 1               '��ʼͨ��
    Else
        
148     If InStr(strResponse, Chr(&H6)) > 0 Then
150         SendSample = GetSendData
152         If mstrWaitSend = "" Then iSendStep = 0
154         SaveLog "����Ӧ��06��ͨѶ�ɹ�������:" & SendSample & "����:" & mstrWaitSend
156     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
158         SendSample = GetSendData
160         If mstrWaitSend = "" Then iSendStep = 0
162         SaveLog "����Ӧ��04����ʾϣ����ͣ������Э���Ե���06,ͨѶ�ɹ�������:" & SendSample & "����:" & mstrWaitSend
164     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then
166         SendSample = GetSendData
168         If mstrWaitSend = "" Then iSendStep = 0
170         SaveLog "����Ӧ��15��ͨѶʧ�ܡ�����:" & SendSample & "����:" & mstrWaitSend
172     ElseIf strResponse = Chr(&H5) Then
174         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
176             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
178             mstrWaitSend = ""
            End If
180         Call SaveLog("����Ӧ��05��������������Ȩ�����η���ȡ��")
182         iSendStep = 0
184         SendSample = Chr(&H6)
        Else
186         Call SaveLog("����Ӧ��" & strResponse)
        End If
188     If mstrWaitSend = "" Then iSendStep = 0    '�ѷ����꣬�˳�˫��ģʽ

    End If
    Exit Function
hErr:
190 SaveLog "SendSample��" & CStr(Erl()) & "�д���" & Err.Description
End Function




Private Function GetSendData() As String
    '�ӻ����л�ȡһ�����ݣ����ء����������û�����ݣ����ؿմ�
    
    Dim strSend As String
    
    If Left$(mstrWaitSend, 1) = "," Then mstrWaitSend = Mid$(mstrWaitSend, 2)
    If mstrWaitSend = "" Then Exit Function
    If InStr(mstrWaitSend, ",") <= 0 Then
        strSend = mstrWaitSend
        mstrWaitSend = ""
    Else
        strSend = Mid$(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
        mstrWaitSend = Mid$(mstrWaitSend, InStr(mstrWaitSend, ",") + 1)
    End If
    
    GetSendData = strSend
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '�������ļ�
    'strIniFile :ini�����ļ���
    'strItems   :Ҫ��ȡ���ֶΣ��磺DATABASE NAME=th_xlb�е� DATABASE NAME,����ֶ���,�ָ�
    'strReturn  :��ȡ�ķ���ֵ����:DATABASE NAME=th_xlb�е� th_xlb ���������ֵ��,�ָ�
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFileSys As New FileSystemObject
    Dim objStream As TextStream

    Dim intloop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '�������ļ��ж�ȡ�������ļ��е�����cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intloop = 0 To intCount
        arrItems(intloop, 0) = Split(strItems, ",")(intloop)
    Next
    
    If objFileSys.FileExists(strIniFile) Then
        Set objStream = objFileSys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If arrLine(0) = arrItems(intloop, 0) Then
                        arrItems(intloop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intloop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intloop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub



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

Private Function GetCheckCode(ByVal str_in As String) As String
    '��ȡУ����
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function



