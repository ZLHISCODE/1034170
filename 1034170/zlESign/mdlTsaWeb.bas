Attribute VB_Name = "mdlTsaWeb"
Option Explicit
'��������ʱ���

Private mobjTsa As Object       '����׼���ҽԺ��ʱ����ӿ�

Public Function TSA_initObj() As Boolean
    On Error Resume Next
    Set mobjTsa = Nothing
    Set mobjTsa = CreateObject("tsaMiddleware.UtilUdp")
    If Err.Number <> 0 Then
        'MsgBox "ʱ����ؼ�û�а�װ��", vbExclamation, gstrSysName
        Exit Function
    End If
    TSA_initObj = True
End Function

Public Function TSA_UnloadObj()
    'ጷŌ���
    If Not mobjTsa Is Nothing Then Set mobjTsa = Nothing
End Function

Private Function GetReturnInfo(ByVal strSign As String) As String
    'ʱ���������Ϣת������
    If strSign = "0001" Then
        GetReturnInfo = "����ͨ���쳣"
    ElseIf strSign = "0002" Then
        GetReturnInfo = "ϵͳ�쳣"
    ElseIf strSign = "0003" Then
        GetReturnInfo = "ϵͳ��æ"
    ElseIf strSign = "0004" Then
        GetReturnInfo = "���ݲ������Ϸ�"
    ElseIf strSign = "0005" Then
        GetReturnInfo = "�û������������"
    ElseIf strSign = "0006" Then
        GetReturnInfo = "���ݿ��쳣"
    ElseIf strSign = "0007" Then
        GetReturnInfo = "DLL�����ļ���ȡ����"
    ElseIf strSign = "1001" Then
        GetReturnInfo = "������Ӧʧ��"
    ElseIf strSign = "1002" Then
        GetReturnInfo = "���������ѼӸǹ�ʱ���"
    ElseIf strSign = "1003" Then
        GetReturnInfo = "�������ݵȴ��Ӹ�ʱ���"
    ElseIf strSign = "2001" Then
        GetReturnInfo = "δ����ʱ���"
    ElseIf strSign = "2002" Then
        GetReturnInfo = "У��ʧ��"
    ElseIf strSign = "2010" Then
        GetReturnInfo = "��֤�ɹ�"
    Else
        GetReturnInfo = strSign
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "ʱ����ӿڷ�����ʾ��" & GetReturnInfo
    End If
End Function

Public Function Times_Tamp(ByVal strSource As String, ByRef strTimeStamp As String) As Boolean
        'ȡʱ���
        Dim intCount As Integer, strSign As String
        On Error GoTo hErr
        
        If mobjTsa Is Nothing Then Exit Function
        
100     strSign = mobjTsa.sendTimestamp(strSource, "sha1")
102     If strSign <> "1000" And strSign <> "1002" And strSign <> "1003" Then
104         strSign = GetReturnInfo(strSign)
106         MsgBox "����ʱ���ʧ�ܣ�" & strSign, vbExclamation, gstrSysName
            Times_Tamp = False
            Exit Function
        Else
108         intCount = 0
110         Do While intCount <= 100
112             strSign = mobjTsa.gettimestampinfo(strSource, "sha1")
                'ǩ���л���ʱ��
114             If InStr(strSign, "#") > 0 Then
116                 strTimeStamp = Split(strSign, "#")(0)
118                 If IsDate(strTimeStamp) Then
120                     strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
                        Times_Tamp = True
                        Exit Function
                    Else
122                     MsgBox "��ȡ��ʱ�������һ�����ڣ�" & strTimeStamp, vbExclamation, gstrSysName
                    End If
124             ElseIf strSign <> "1003" And strSign <> "2001" Then
126                 strSign = GetReturnInfo(strSign)
128                 MsgBox "��ȡʱ���ʧ�ܣ�" & strSign, vbExclamation, gstrSysName
                    Exit Function
                End If
130             intCount = intCount + 1
            Loop
        End If
132     Times_Tamp = True
        Exit Function
hErr:
134    MsgBox "ȡʱ���-��" & CStr(Erl()) & "��," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function verify_Timestamp(ByVal strSource As String) As Boolean
    '��֤ʱ���
    Dim strData As String
    If mobjTsa Is Nothing Then Exit Function
    strData = mobjTsa.verifyTimestamp(strSource, "sha1")
    If strData <> "2010" Then
        MsgBox "��֤ʱ���ʧ�ܣ�" & GetReturnInfo(strData), vbExclamation, gstrSysName
        Exit Function
    End If
    verify_Timestamp = True
End Function

Private Function verify_getTimestamp(ByVal strSource As String) As String
    '��ȡʱ���  ������Ҽӵġ�
    Dim strData As String
    Dim strTimeStamp As String
    If mobjTsa Is Nothing Then Exit Function
    
    strData = mobjTsa.gettimestampinfo(strSource, "sha1")
    If strData = "2001" Then
        MsgBox "��ȡ��֤ʱ���ʧ�ܣ�" & GetReturnInfo(strData), vbExclamation, gstrSysName
        verify_getTimestamp = "��"
        Exit Function
    End If
    
    If InStr(strData, "#") > 0 Then
        strTimeStamp = Split(strData, "#")(0)
        If IsDate(strTimeStamp) Then
            strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
        Else
            MsgBox "��ȡ��ʱ�������һ�����ڣ�" & strTimeStamp, vbExclamation, gstrSysName
            verify_getTimestamp = "��"
            Exit Function
        End If
    End If
    verify_getTimestamp = strTimeStamp
    
End Function



