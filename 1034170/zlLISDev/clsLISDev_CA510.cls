VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:CA510
'����:ϣɭ����
'����:�¶�
'�û�:�ɶ�XXҽԺ

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
        
    Dim strCode As String
    Dim lngS As Long, lngE As Long, i As Integer
    Dim str�걾�� As String, str��Ŀ As String, str��� As String, str���� As String, str�ʿ� As String
    Dim strDateFmt As String, lngStart As Long
    Dim strItems As String
    
    strReserved = strSource
    
    If strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        lngS = InStr(strReserved, Chr(2))
        lngE = InStr(strReserved, Chr(3))
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
        
        strDateFmt = ReadIni("CA510", "���ڸ�ʽ", App.Path & "\CA510.ini", "")
        If strDateFmt = "" Then
            strDateFmt = "YYMMDD"
            Call WriteIni("CA510", "���ڸ�ʽ", "YYMMDD", App.Path & "\CA510.ini")
        End If
        lngStart = Val(ReadIni("CA510", "����λ��", App.Path & "\CA510.ini", ""))
        If lngStart = 0 Then
            lngStart = 52
            Call WriteIni("CA510", "����λ��", "52", App.Path & "\CA510.ini")
        End If
        
        str�ʿ� = Mid(strCode, 10, 1)
        If str�ʿ� = "C" Then
            str�ʿ� = "1"
        Else
            str�ʿ� = "0"
        End If
        If strDateFmt = "YYMMDD" Then
            str���� = "20" & Mid(strCode, 11, 2) & "-" & Mid(strCode, 13, 2) & "-" & Mid(strCode, 15, 2)
        ElseIf strDateFmt = "MMDDYY" Then
            str���� = "20" & Mid(strCode, 15, 2) & "-" & Mid(strCode, 11, 2) & "-" & Mid(strCode, 13, 2)
        ElseIf strDateFmt = "DDMMYY" Then
            str���� = "20" & Mid(strCode, 15, 2) & "-" & Mid(strCode, 13, 2) & "-" & Mid(strCode, 11, 2)
        End If
        
        If Not IsDate(str����) Then
            str���� = Format(Now, "yyyy-MM-dd")
        Else
            str���� = Format(CDate(str����), "yyyy-MM-dd")
        End If
        
        lngS = lngStart - 12
        If lngS <= 10 Then Exit Sub
        
        str�걾�� = Val(Mid(strCode, lngS - 10, 10))
        strCode = Mid(strCode, lngStart)
        
        strItems = ""
        Do While Len(strCode) > 0
            str��Ŀ = Mid(strCode, 1, 9)
            str��� = Trim(Mid(str��Ŀ, 4))
            str��Ŀ = Trim(Left(str��Ŀ, 3))

            If str��Ŀ <> "" And str��� <> "" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Trim(str���)
            End If
            strCode = Mid(strCode, 10)
            
        Loop
        If Val(str�걾��) > 0 And strItems <> "" Then
            strResult = str���� & "|" & Val(str�걾��) & "| |ѪҺ|" & str�ʿ� & strItems
        End If
    End If
    
End Sub


Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

