VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ST360"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:KHB ST-360ø����
'����:�ƻ�
'����:�¶�
'�û�:�����зο�ҽԺ
Private mintStp As Integer
Private mstrReturn As String
Private mstrCode As String

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
    Dim strLine As String
    Dim strData(8, 12) As String, str�޽�� As String
    Dim strCode As String, strTwo As String, strTmp As String
    Dim lngS As Long, lngE As Long, strIniFile As String
    Dim strWave As String   '���ص���һ�廹�Ƕ���
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream
    
    
    strIniFile = App.Path & "\St_360.ini"
    
    strWave = "NO"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "ָ���ʽ = һ��"
       objStream.WriteLine "���ݸ�ʽ = NO"
       objStream.Close
       Set objStream = Nothing
    Else
        Set objStream = objFileSystem.OpenTextFile(strIniFile)
        Do While Not objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If strLine Like "*���ݸ�ʽ*=*" Then
                strWave = Trim(Split(strLine, "=")(1))
            End If
        Loop
        objStream.Close
        Set objStream = Nothing
    End If
    
    
    For intLoop = 1 To 8
        For i = 1 To 12
            strData(intLoop - 1, i - 1) = "0.000"
        Next
    Next
    
    For intLoop = 0 To 7
        strLine = ""
        For i = 0 To 11
            strLine = strLine & ";" & strData(intLoop, i)
        Next
        If Trim(strLine) <> "" Then
            str�޽�� = str�޽�� & "|" & Mid(strLine, 2)
        End If
    Next

    If mstrReturn <> "" Then
        strResult = mstrReturn

        WriteLog "�ڶ��η�����������", strCode, strResult
        mstrReturn = ""
        Exit Sub
    End If
    
    mstrCode = mstrCode & strSource
    lngS = InStr(mstrCode, "A")
    If lngS > 5 Then lngS = lngS - 5
    lngE = 0
    If lngS > 0 Then
        lngE = InStr(lngS, mstrCode, "H")
        If strWave = "NO" Then
            If lngE > 0 Then
                lngE = InStr(lngE, mstrCode, "A")
                If lngE > 0 Then
                    lngE = InStr(lngE, mstrCode, "H")
                    If lngE > 0 Then
                        lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
                    Else
                        lngE = 0
                    End If
                Else
                    lngE = 0
                End If
            Else
                lngE = 0
            End If
        Else
            If lngE > 0 Then
                lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
            Else
                lngE = 0
            End If
        End If
    End If
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid$(mstrCode, lngS, lngE - lngS + 2)
        mstrCode = Mid$(mstrCode, lngE + 2)
        '--------------------------------------------------------------------
        
            '������һ�ξͽ����μ������ݷ���,����ֳ�����,�����ݷ���
             
            
            If UBound(Split(strCode, Chr(&HD) & Chr(&HA))) < 9 Then Exit Sub
            
            '��ʼ����
            strResult = "": strTmp = ""
            For intLoop = LBound(Split(strCode, Chr(&HD) & Chr(&HA))) To UBound(Split(strCode, Chr(&HD) & Chr(&HA)))
                strLine = Trim(Split(strCode, Chr(&HD) & Chr(&HA))(intLoop))
                If InStr(",A,B,C,D,E,F,G,H,", "," & Trim(Mid(strLine, 1, 1)) & ",") > 0 Then
                    
                    If strTwo <> "ABCDEFGH" Then
                        '��һ�η��صĽ��
                        If InStr(strTwo, Trim(Mid(strLine, 1, 1))) <= 0 Then strTwo = strTwo & Trim(Mid(strLine, 1, 1))
                        varline = Split(strLine, " ")
                        For i = LBound(varline) + 1 To UBound(varline)
                            If i <= 12 Then
                               strData(intLoop - 1, i - 1) = Format(Val(varline(i)), "0.000")
                            End If
                        Next
                    Else
                        '�ڶ��η��صĽ��
                        strTmp = ""
                        varline = Split(strLine, " ")
                        For i = LBound(varline) + 1 To UBound(varline)
                            If i <= 12 Then
                                strTmp = strTmp & ";" & Format(Val(varline(i)), "0.000")
                            End If
                        Next
                        If Trim(strTmp) <> "" Then
                            mstrReturn = mstrReturn & Mid(strTmp, 2) & "|"
                        End If
                    End If
                End If
            Next

            
            For intLoop = 0 To 7
                strLine = ""
                For i = 0 To 11
                    strLine = strLine & ";" & strData(intLoop, i)
                Next
                If Trim(strLine) <> "" Then
                    strResult = strResult & "|" & Mid(strLine, 2)
                End If
            Next
            
            
            If strResult <> "" Then
                If str�޽�� = strResult Then
                    strResult = ""
                Else
                    strResult = Mid(strResult, 2)
                End If
            End If
            WriteLog "��һ�η�����������", strCode, strResult
        
        '--------------------------------------------------------------------
        lngS = InStr(mstrCode, "A")
        lngE = 0
        If lngS > 0 Then
            lngE = InStr(lngS, mstrCode, "H")
            If strWave = "NO" Then
                If lngE > 0 Then
                    lngE = InStr(lngE, mstrCode, "A")
                    If lngE > 0 Then
                        lngE = InStr(lngE, mstrCode, "H")
                        If lngE > 0 Then
                            lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
                        Else
                            lngE = 0
                        End If
                    Else
                        lngE = 0
                    End If
                Else
                    lngE = 0
                End If
            Else
                If lngE > 0 Then
                    lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
                Else
                    lngE = 0
                End If
            End If
        End If
    Loop
    
        
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
    Dim lngTime As Long
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String
    Dim strType As String
    
    strIniFile = App.Path & "\St_360.ini"
    strType = "һ��"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "ָ���ʽ = һ��"
       objStream.WriteLine "���ݸ�ʽ = NO"
       objStream.Close
       Set objStream = Nothing
    Else
        Set objStream = objFileSystem.OpenTextFile(strIniFile)
        Do While Not objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If strLine Like "*ָ���ʽ*=*" Then
                strType = Trim(Split(strLine, "=")(1))
            End If
        Loop
        objStream.Close
        Set objStream = Nothing
    End If

    Select Case strCmdType
        Case "Begin"                    '��ʼ ����
            CmdAnalyse = ""
        Case "End"                      '���� ����
            CmdAnalyse = ""
        Case "Revert"                   '�ظ�����
            CmdAnalyse = "e_end"
        Case "Play"                     '�������
            If mintStp = 0 Then
                CmdAnalyse = "" 'CmdAnalyse = "e(0)|e_end"
            Else
                CmdAnalyse = ""
            End If
        Case "ReadData_TimeOut"
            CmdAnalyse = 90
        Case "ReadData"                 '����
            If mintStp = 1 Then
                CmdAnalyse = "|"         '��ָ��
                mintStp = 0
            Else
                If strType = "һ��" Then
                    CmdAnalyse = "p(1,3)"
                Else
                    CmdAnalyse = "p(2,3,7)"
                End If
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
        Case "e(0)"
            CmdAnalyse = "e(0)"
        Case "e(1)"
            CmdAnalyse = "e(1)"
        Case "e(2)"
            CmdAnalyse = "e(2)"
        Case "e(3)"
            CmdAnalyse = "e(3)"
        Case "e(4)"
            CmdAnalyse = "e(4)"
        Case "e(5)"
            CmdAnalyse = "e(5)"
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
        Case "Z5_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z10_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z15_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z20_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z25_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z30_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case Else
            If Left(strCmdType, 1) = "Z" Then
                '���ʱ��
                lngTime = Val(Mid(strCmdType, 2))
                If lngTime >= 0 And lngTime <= 60 Then
                    CmdAnalyse = "Z(" & Format(lngTime, "00") & ")"
                End If
            End If
    End Select
    
End Function



