VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BR680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BIO-RAD 680ø����
'����:����
'����:�¶�
'�û�:����������ְ��ҽԺ

Private mstrCode As String
Private mintStp As Integer   '������ô���
Private mstrReturn As String '����ο�����

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
        Dim varline As Variant, strTmp As String
        Dim strLine As String
        Dim strData(8, 12) As String, str�޽�� As String
        Dim lng_S As Long, lng_E As Long, intRow As Integer, intCol As Integer
        On Error GoTo errH
100     mstrCode = mstrCode & strSource
        strReserved = ""
102     strResult = ""
104     If mstrCode Like "*begin*end*begin*end" & Chr(&HD) & "*" Then
            WriteLog "analyse", mstrCode, strResult
            '---���������
106         lng_S = InStr(mstrCode, "begin" & Chr(&HD))
108         lng_E = InStr(lng_S, mstrCode, "end" & Chr(&HD)) + 3
110         strTmp = Mid$(mstrCode, lng_S, lng_E - lng_S + 1)
112         mstrCode = Mid$(mstrCode, lng_E + 1)
            
114         varline = Split(strTmp, Chr("&H0D"))
116         intRow = 0: intCol = 0
118         For i = LBound(varline) To UBound(varline)
120             strTmp = varline(i)
122             strLine = ""
124             If Len(strTmp) = 6 * 12 Then
126                 Do While strTmp <> ""
128                     strData(intRow, intCol) = Val(Mid$(strTmp, 1, 6))
130                     intCol = intCol + 1
132                     If intCol > 11 Then
134                         intRow = intRow + 1
136                         intCol = 0
                        End If
                        'strLine = strLine & ";" & Val(Mid$(strTmp, 1, 6))
138                     strTmp = Mid$(strTmp, 7)
                    Loop
                End If

            Next
        
            '�ο��������
142         lng_S = InStr(mstrCode, "begin" & Chr(&HD))
144         lng_E = InStr(lng_S, mstrCode, "end" & Chr(&HD)) + 3
146         strTmp = Mid$(mstrCode, lng_S, lng_E - lng_S + 1)
148         mstrCode = Mid$(mstrCode, lng_E + 1)
        
150         varline = Split(strTmp, Chr("&H0D"))
152         intRow = 0: intCol = 0
154         For i = LBound(varline) To UBound(varline)
156             strTmp = varline(i)
158             strLine = ""
160             If Len(strTmp) = 6 * 12 Then
162                 Do While strTmp <> ""
164                     strData(intRow, intCol) = Val(strData(intRow, intCol)) - Val(Mid$(strTmp, 1, 6))
166                     intCol = intCol + 1
168                     If intCol > 11 Then
170                         intRow = intRow + 1
172                         intCol = 0
                        End If
174                     strTmp = Mid$(strTmp, 7)
                    Loop
                End If

            Next

176         For intRow = 0 To 7
178             strLine = ""
180             For intCol = 0 To 11
182                 strLine = strLine & ";" & strData(intRow, intCol)
                Next
184             strResult = strResult & "|" & Mid(strLine, 2)
            Next
186         strResult = Mid$(strResult, 2)
        End If
        Exit Sub
errH:
192     WriteLog "BR680.Analyse", CStr(Erl()) & "��," & Err.Description & vbCrLf & mstrCode, strResult
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

    Select Case strCmdType
        Case "Begin"                    '��ʼ ����
            CmdAnalyse = "EIA.READER AQ" & Chr(&HD) & "|" & "ERE 8073" & Chr(&HD) & "|" & "ERE 0000" & Chr(&HD)
        Case "End"                      '���� ����
            CmdAnalyse = "EIA.READER RL" & Chr(&HD) & "|" & "ERE 0000" & Chr(&HD)
        Case "Revert"                   '�ظ�����
            CmdAnalyse = ""
        Case "Play"                     '�������
            If mintStp = 1 Then
                CmdAnalyse = "|"         '��ָ��
            Else
                CmdAnalyse = "EIA.READER ISTATUS" & Chr(&HD) & "|" & "ERE 0000 0 0 **.* **.*" & Chr(&HD)
            End If
        Case "ReadData_TimeOut"
            CmdAnalyse = 80
        Case "ReadData"                 '����
            If mintStp = 1 Then
                CmdAnalyse = "|"         '��ָ��
                mintStp = 0
            Else
                CmdAnalyse = "EIA.READER RPLATE 1 5 H 2 4" & Chr(&HD)
                mintStp = 1
            End If
        Case "Out"                      '����
            CmdAnalyse = ""
        Case "In"                       '����
            CmdAnalyse = ""
        '����ָ��Ϊ�̶���,���뱣��,������û�ж�Ӧָ��,��Ϊ�մ�����
        '----------------------------------------------------------------------------------
        Case "450"                      '����һ
            CmdAnalyse = ""
        Case "630"                      '������
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
        Case "��"                       '
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '�հ���ʽ,�������������������еĿհ���ʽ�趨��Ӧ��ָ��
        Case "����"
            CmdAnalyse = ""
        Case "�Լ�"
            CmdAnalyse = ""
    End Select
    
End Function

