VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_RT6100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:RT6000
'����:�׶�
'����:
'�û�:�⴨������ҽԺ
Private mintStp As Integer


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public str��־ As String




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
    Dim strData(8, 12) As String, str��� As String, lngMod As Long, lng_E As Long, lng_S As Long, j As Long
    Dim str���1 As String, strDate As String
    
    
    
    If mintStp = 1 Then
'        ������һ�ξͽ����μ������ݷ��� , ����ֳ�����, �����ݷ���
        strReserved = strSource
        '��ʼ����
        strResult = ""
        If InStr(strSource, "TRANSFER FINISH") > 0 Then
            varline = Split(strSource, "B,1,")  'B,1,��1��Ϊ��ø���Ǵ���ʱ�����1
            lng_S = LBound(varline)
            lng_E = UBound(varline)
            For i = lng_S + 1 To lng_E
                
                strLine = varline(i)
                str��� = Split(strLine, ",")(7)
                
                lngMod = i Mod 12
                If i > 10 And lngMod = 0 Then
                
                    str���1 = str���1 & ";" & str��� & "|"
                    str��� = ""
                Else
                    str���1 = str���1 & ";" & str���
                    str��� = ""
                End If
                
            Next
            If lng_E < 96 Then
                For j = lng_E + 1 To 96
                    lngMod = j Mod 12
                    If lngMod = 0 Then
                        str���1 = str���1 & ";" & "0.000" & "|"
                        str��� = ""
                    Else
                        str���1 = str���1 & ";" & "0.000"
                        str��� = ""
                    End If
                Next
            End If
       End If
       
       If CheckTheChar("|", str���1) = 8 Then
            strResult = Replace(Mid(str���1, 1, Len(str���1) - 1), "*", "")
            varline = Split(strResult, "|")
            For intLoop = LBound(varline) To UBound(varline)
                strDate = strDate & "|" & Mid(varline(intLoop), 2)
            Next
        strResult = Mid(strDate, 2)
        strReserved = ""
        strSource = ""
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
            CmdAnalyse = 90
        Case "ReadData"                 '����
            If mintStp = 0 Then
                CmdAnalyse = " |"
                mintStp = 1
            Else
                Sleep (500)
                CmdAnalyse = " |"
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





Function CheckTheChar(TheChar, TheString) As Long
    '\'TheChar=\"Ҫ�����ַ���\"
    '\'TheString=\"�������ַ���\"
    Dim n As Long
    If InStr(TheString, TheChar) Then
        For n = 1 To Len(TheString)
            If Mid(TheString, n, Len(TheChar)) = TheChar Then
                CheckTheChar = CheckTheChar + 1
            End If
        Next
    'CheckTheChar=\"����ַ�\"&CheckTheChar&\"��\"
    'Else
    'CheckTheChar=\"0��\"
    End If
End Function


