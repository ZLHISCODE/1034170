VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ZS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����֮��ZS-2ø����
'����:�Ϻ���ʢ
'����:Ī��ά
'�û�:�ձ�������ҽԺ
'ʱ��:2013/12/19
Private mintStp As Integer


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

    Dim intLoop As Integer, i As Integer
    Dim varline As Variant
    Dim strLine As String, j As Integer
    Dim strData(8, 12) As String, str�޽�� As String
    Dim lng_S As Long, lng_E As Long
    Dim str���2 As String, str���3 As String, str���4 As String, str���5 As String
    Dim str���6 As String, str���7 As String, str���8 As String, str���1 As String

    On Error GoTo errH
    If mintStp = 1 Then
        strReserved = strSource
        lng_S = InStr(strReserved, "")
        If lng_S > 0 Then
            lng_E = InStr(strReserved, "")
        End If
        If lng_S > 0 And lng_E > 0 Then
            strReserved = Replace(strReserved, "-", " ")
            varline = Split(strReserved, Chr(&H20))(0)
            strLine = Split(strReserved, Chr(&H20))(1)
            If strLine > 0 Then
                For i = 1 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str���1 = str���1 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 2 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str���2 = str���2 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 3 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str���3 = str���3 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 4 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str���4 = str���4 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 5 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str���5 = str���5 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 6 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str���6 = str���6 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 7 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str���7 = str���7 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 8 To 96
                    If i = 96 Then
                        strLine = Format(Mid(Replace(Split(strReserved, Chr(&H20))(i), "-", " "), 1, 4) * 0.001, "0.000")
                        str���8 = str���8 & ";" & Format(Mid(Replace(Split(strReserved, Chr(&H20))(i), "-", " "), 1, 4) * 0.001, "0.000")
                    Else
                        strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                        str���8 = str���8 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                        i = i + 8 - 1
                    End If
                Next
            Else
                '�����Ǵ���ø���Ǵ���ȫ��O�Ľ��
                For i = 1 To 96
                    strLine = Format(0, "0.000")
                    str���1 = str���1 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 2 To 96
                    strLine = Format(0, "0.000")
                    str���2 = str���2 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 3 To 96
                    strLine = Format(0, "0.000")
                    str���3 = str���3 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 4 To 96
                    strLine = Format(0, "0.000")
                    str���4 = str���4 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 5 To 96
                    strLine = Format(0, "0.000")
                    str���5 = str���5 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 6 To 96
                    strLine = Format(0, "0.000")
                    str���6 = str���6 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 7 To 96
                    strLine = Format(0, "0.000")
                    str���7 = str���7 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 8 To 96
                    strLine = Format(0, "0.000")
                    str���8 = str���8 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
            End If
            strResult = Mid(str���1, 2) & "|" & Mid(str���2, 2) & "|" & Mid(str���3, 2) & "|" & Mid(str���4, 2) & "|" & Mid(str���5, 2) & "|" & Mid(str���6, 2) & "|" & Mid(str���7, 2) & "|" & Mid(str���8, 2) & "|"
        End If
    End If
    Exit Sub
errH:
    WriteLog "ZS2", "zs", Err.Number & ":" & Err.Description
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

Public Function CmdAnalyse(strCmdType As String) As String
'����:  �ֽ�����Ϊ����ִ�е�����

'MsgBox strCmdType
    Select Case strCmdType
    Case "Begin"                    '��ʼ ����
        CmdAnalyse = ""
    Case "End"                      '���� ����
        CmdAnalyse = ""
    Case "Revert"                   '�ظ�����
        CmdAnalyse = ""
    Case "Play"                     '�������
        If mintStp = 0 Then
            CmdAnalyse = ""
        Else
            CmdAnalyse = ""
        End If
    Case "ReadData_TimeOut"
        CmdAnalyse = 50
    Case "ReadData"                 '����
        If mintStp = 0 Then
            CmdAnalyse = Chr(&H1B) & Chr(&H56) & Chr(&H55) & Chr(&H2) & Chr(&H4) & Chr(&HD)
            mintStp = 1
        Else
            Sleep (3000)
            CmdAnalyse = ""
            mintStp = 1
        End If
    Case "Out"                      '����
        CmdAnalyse = ""
    Case "In"                       '����
        CmdAnalyse = ""
        '����ָ��Ϊ�̶���,���뱣��,������û�ж�Ӧָ��,��Ϊ�մ�����
        '----------------------------------------------------------------------------------
        '                               '����,�������������������еĲ����趨��Ӧ��ָ��
    Case "450"                      '����һ
        CmdAnalyse = ""
    Case "620"                      '������
        CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '���巽ʽ,�������������������еĽ��巽ʽ�趨��Ӧ��ָ��
    Case "����"                     '����
        CmdAnalyse = ""
    Case "����"                     '��
        CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '���Ƶ��,�������������������е����Ƶ���趨��Ӧ��ָ��
    Case "��"                       '
        CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '�հ���ʽ,�������������������еĿհ���ʽ�趨��Ӧ��ָ��
    Case "����"
        CmdAnalyse = ""
    Case "�Լ�"
        CmdAnalyse = ""
    Case Else

    End Select

End Function











