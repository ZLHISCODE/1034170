VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AC600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AC600 �ǻ�Ѫ�쵰����
'����:
'����:����
'�û�:��������УҽԺ

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
    Dim strCode As String
    Dim i As Integer, lngE As Long, lngS As Long
    Dim strDate As String, strSample As String, StrQ As String, Items As String, strE As String
    

    If strSource <> "" Then
        strReserved = strReserved & strSource
    End If

    On Error GoTo errHand
    
    Call writelislog("�յ������ݣ�" & strReserved)
    lngE = InStr(strReserved, Chr(&H2))
    lngS = InStr(strReserved, Chr(&H3))
    If lngE > 0 And lngS > 0 Then
        strCode = strReserved
        strReserved = Mid(strReserved, lngS + 1)
        strDate = Format(Mid(strCode, 2, 19), "yyyy-mm-dd hh:mm:ss")
        StrQ = Mid(strCode, 21, 15)

        strSample = Val(Right(Split(StrQ, "-")(1), 4))
        StrQ = Trim(Left(Split(StrQ, "-")(1), 3))
        If StrQ = "ST" Then
            strE = "1"
        ElseIf StrQ = "CAL" Then
            StrQ = "1"
            strE = "0"
        Else
            StrQ = "0"
            strE = "0"
        End If
        Items = Items & "|HbA1c|" & Val(Mid(Mid(strCode, 36, 26), 9, 6))
        Items = Items & "|IFCC|" & Val(Mid(Mid(strCode, 62, 23), 16, 4))
        Items = Items & "|NGSP|" & Val(Mid(Mid(strCode, 85), 26, 5))
        If Items <> "" Then
            strResult = strDate & "|" & Val(strSample) & IIf(strE = "1", "^1^", "") & "| |Ѫ��|" & IIf(StrQ = "1", "1", "0") & Items
        End If
    End If
    
    Exit Sub
errHand:
    WriteLog "clsLISDev_AC600", Err.Description, "�����������" & strCode
End Sub
Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_AC600_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function