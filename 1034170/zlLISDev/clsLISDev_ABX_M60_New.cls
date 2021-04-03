VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_M60_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ABX Micros 60
'����:ABX
'����:�¶�
'�û�:�����·ҽԺ

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
    
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varLine As Variant, intRow As Integer, strLine As String
    Dim str�걾�� As String, str���� As String, strItems As String, lngItem As Long
    
    lngS = InStr(strSource, Chr(2))
    lngE = InStr(strSource, Chr(3))
    strReserved = strSource
    If lngS > 0 And lngE > 0 And lngE > lngS Then
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
        varLine = Split(strCode, Chr(&HD))
        For intRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(intRow)
            If InStr("!,2,3,4,5,6,7,8,@,A,B,C,#,%,',"""",$,&", Mid(strLine, 1, 1)) > 0 Then
                '��Ŀ
                lngItem = lngItem + 1
                strItems = strItems & "|" & lngItem & "|" & Val(Mid(strLine, 3))
            ElseIf Mid(strLine, 1, 1) = "q" Then
                '����
                str���� = Format(CDate(Mid(strLine, 9, 2) & "-" & Mid(strLine, 6, 2) & "-" & Mid(strLine, 3, 2)), "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            ElseIf Mid(strLine, 1, 1) = "u" Then
                '�걾��
                str�걾�� = Val(Mid(strLine, 3))
            End If
        Next
        
        If Val(str�걾��) > 0 And strItems <> "" Then
            strResult = str���� & "|" & str�걾�� & "| |ȫѪ|0" & strItems
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