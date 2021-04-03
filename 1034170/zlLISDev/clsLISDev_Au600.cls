VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AU600
'����:���ְ�˹
'����:�¶�
'�û�:��������Ժ

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, lngLength As Long, iRows As Integer
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    ''
    Dim dblTP As Double, dblALB As Double, dblTBIL As Double, dblDBIL As Double

    
    Dim intE As Integer, intP As Integer, intQ As Integer
    Dim intEmergency As Integer
    
    intE = Val(ReadIni("AU600", "E", App.Path & "\AU600.ini"))
    If intE = 0 Then intE = 8000:   Call WriteIni("AU600", "E", 8000, App.Path & "\AU600.ini")
    
    intQ = Val(ReadIni("AU600", "Q", App.Path & "\AU600.ini"))
    If intQ = 0 Then intE = 6000:   Call WriteIni("AU600", "Q", 6000, App.Path & "\AU600.ini")
    
    intP = Val(ReadIni("AU600", "P", App.Path & "\AU600.ini"))
    If intP = 0 Then intE = 3000:   Call WriteIni("AU600", "P", 3000, App.Path & "\AU600.ini")
    
    intEmergency = Val(ReadIni("AU600", "�Ƿ�ʹ�ü���", App.Path & "\AU600.ini"))
    If intEmergency = 0 Then Call WriteIni("AU600", "�Ƿ�ʹ�ü���", CStr(intEmergency), App.Path & "\AU600.ini")
    
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "***" Or strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        strReserved = Replace(strReserved, "", "")
        strReserved = Replace(strReserved, "", "")
        
        strReserved = Replace(strReserved, "", "")
        strReserved = Replace(strReserved, "", "")
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            
            
            lngLength = InStr(strReserved, "") - InStr(strReserved, "") + 1
            
            strLine = Mid(strReserved, InStr(strReserved, ""), lngLength)
            If InStr(2, strLine, "") > 0 Then
                strLine = Mid(strLine, InStr(2, strLine, ""))
            End If
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            
            If strLine Like "RB*" Or strLine Like "RE*" Then  '����ʼ�����������������
                ' #
                strCmd = ""
            ElseIf strLine Like "R*" Then '����ɨ������
                '# ��ȡ���ڣ����뷵��
                'strCmd="1|" & ����^^����
            ElseIf strLine Like "D *" Or strLine Like "DQ*" Then
                strCmd = Chr(&H6)
                str���� = Format(Now, "yyyy-MM-dd HH:mm:ss")
                strTmp = Mid(strLine, 7, 1)
                If strTmp = "E" Then '�ʿ�
                    str�걾�� = intE + Val(Replace(Mid(strLine, 8, 3), " ", "")) & IIf(intEmergency = 0, "^1", "^0")
                    str�����˱걾�ʿ� = " |Ѫ��|0"
                ElseIf strTmp = "Q" Then  '����
                    str�걾�� = intQ + Val(Mid(strLine, 8, 3)) * 10 + Val(Mid(strLine, 11, 5)) & IIf(intEmergency = 0, "^1", "^0")
                    str�����˱걾�ʿ� = " |Ѫ��|1"
                ElseIf strTmp = "P" Then
                    str�걾�� = intP + Val(Replace(Mid(strLine, 8, 3), " ", "")) & IIf(intEmergency = 0, "^1", "^0")
                    str�����˱걾�ʿ� = " |Ѫ��|0"
                Else
                    str�걾�� = Val(Replace(Mid(strLine, 7, 4), " ", "")) & "^0"
                    str�����˱걾�ʿ� = " |Ѫ��|0"
                End If
                str��Ŀ = Mid(strLine, 16, InStr(strLine, ""))
                If Len(str��Ŀ) >= 13 Then
                    iRows = Len(str��Ŀ) / 13
                    For iRow = 1 To iRows
                        strItem = strItem & "|" & Mid(str��Ŀ, 1, 2) & "|" & Val(Trim(Mid(str��Ŀ, 3, 9)))
                        str��Ŀ = Mid(str��Ŀ, 14)
                        
                        If Mid(str��Ŀ, 1, 2) = "04" Then
                            dblTP = Val(Trim(Mid(str��Ŀ, 3, 9)))
                        ElseIf Mid(str��Ŀ, 1, 2) = "05" Then
                            dblALB = Val(Trim(Mid(str��Ŀ, 3, 9)))
                        ElseIf Mid(str��Ŀ, 1, 2) = "06" Then
                            dblTBIL = Val(Trim(Mid(str��Ŀ, 3, 9)))
                        ElseIf Mid(str��Ŀ, 1, 2) = "07" Then
                            dblDBIL = Val(Trim(Mid(str��Ŀ, 3, 9)))
                        End If
                        
                        
                    Next
                    ' �򵰰�GLO=�ܵ���TP-�׵���ALB
                    ' �����A/G=�׵���ALB/�򵰰�GLO
                    If dblTP <> 0 And dblALB <> 0 Then
                        strItem = strItem & "|GLO|" & Format(dblTP - dblALB, "0.0")
                        If dblTP - dblALB <> 0 Then
                            strItem = strItem & "|AG|" & Format(dblALB / (dblTP - dblALB), "0.0")
                        End If
                    End If
                    '��ӵ�����IDBIL=�ܵ�����TBIL-ֱ�ӵ�����DBIL
                    If dblDBIL <> 0 And dblTBIL <> 0 Then
                        strItem = strItem & "|IDBIL|" & Format(dblTBIL - dblDBIL, "0.0")
                    End If
                End If
                '����
            End If
            
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            
            If Not strReserved Like "***" Then
                blnEnd = True
            ElseIf strReserved Like "*R*" Then '��������Ϣ��ֹͣ�������,���ء�
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    
    '������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        SendSample = ""

End Function

