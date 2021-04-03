VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7170A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  ����7180
'ҽԺ:  �ɶ�����
'����:  ����
'ʱ��:  2008/1/16
Const str_MOR = ">"
Const str_REP = "?"
Private str_REC As String
Private mstrSend As String
Private mLastSend As String

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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    '����ͼ��ķ�ʽ��
    '                   1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '                   2.�ж��ͼ������ʱʹ��"^"���ָ�
    '                   3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ
    '                     a) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);X1;X2;X3;X4;X5...
    '                     b) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                        ��:00000100001000010000100010;00000100001000010000100010;
    '                        ˵��:1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                             2.�ж��ٸ��ֺž��ж�����
    '                             3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                             3.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    
    Dim strData As String
    Dim aData() As String
    Dim intloop As Integer

    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    Dim str_Date  As String
    Dim str_Ji As String                '�����־
    Dim str_QC As String              '�ʿ�
    Dim str_Tmp As String
    Dim curALB As Currency, curGLO As Currency, curHBA1C As Currency, curHB As Currency
    
    Dim curTP As Currency, curTBIL As Currency, curDBIL As Currency, curALT As Currency, curAST As Currency
    Dim curIBIL As Currency, curASTALT As Currency
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": strSampleNO = ""
    'û���ҵ���ʼλ��ʱ�˳�
'    If InStr(strSource, Chr(&H2)) <= 0 Then
'        strReserved = strSource
'        'strCmd = Chr(&H6)
'        Exit Sub
'    End If
'    'û���ҵ�����λ��ʱ�˳�
'    If InStr(strSource, Chr(&H3)) <= 0 Then
'        strReserved = strSource
'        'strCmd = Chr(&H6)
'        Exit Sub
'    End If
    str_Date = "": str_Ji = "": str_QC = ""
    If Not strSource Like "***" Then
        strReserved = strSource
        Exit Sub
    End If
    
    lngBegin = InStr(strSource, Chr(&H2))
    lngEnd = InStr(strSource, Chr(&H3))
    
    'ȡ��ǰһ�εĽ���������ķ���
    strLine = Mid(strSource, lngBegin, lngEnd - lngBegin + 1)
    strReserved = Mid(strSource, lngEnd + 1)
    If InStr(strReserved, Chr(&H2)) <= 0 Then
        strReserved = ""    'û���ҵ���ʼλ��ʱ�˳�
    End If
    
    If str_REC = "" Then
        
        Call ReadConf(App.Path & "\HITA7170A.ini", "Ӧ��", str_REC)
        If str_REC = "" Then
            str_REC = "A"
            Call WriteIni("7170A", "Ӧ��", "A", App.Path & "\HITA7170A.ini")
        End If
    End If
    
    If strLine = str_MOR Or strLine = str_REP Then
        '˫��ͨ�ŵ���������ԭ��Ӧ�𼴿�
        If mstrSend = "" Then
            strCmd = str_MOR
        End If
    ElseIf strLine Like ";N1*" And Len(strLine) = 42 Then
        
        strCmd = "1|" & str_REC  '�յ�ɨ��ı걾������
    Else
        '��ʼ����(û���ҵ�ʱ�䣬��ʱʹ�ñ���ʱ��)
        strCmd = str_REC
        
        If InStr(",:,1,2,", Mid(strLine, 2, 1)) > 0 Then
        
            If UCase(Mid(strLine, 3, 1)) = "Q" Then
                str_Ji = "1"
            ElseIf UCase(Mid(strLine, 3, 1)) = "G" Then
                Exit Sub
            End If
            
            str_QC = "0"
            If Mid(strLine, 2, 1) = "1" Then
                str_QC = "1"
            ElseIf Mid(strLine, 3, 1) = "F" Then
                str_QC = "F1"
            End If
            If Val(Mid(strLine, 5, 5)) = 0 Then 'ȡ����
                
                If Val(Mid(strLine, 15, 2)) <> 0 Then
                    '����ǰ��λ��Ϊ0 ����������
                    strData = Now & "|^" & IIf(str_Ji = "1", "1", "") & "^" & Val(Mid(strLine, 15, 13)) & "| |Ѫ��|" & str_QC
                Else
                    '����ǰ��λΪ0,�Ͱ�������Ϊ�걾�Ŵ���
                    strData = Now & "|" & Val(Mid(strLine, 15, 13)) & IIf(str_Ji = "1", "^1^", "") & "| |Ѫ��|" & str_QC
                End If
            Else
                str_Date = Trim(Mid(strLine, 32, 10))
                If IsNumeric(str_Date) And Len(str_Date) = 10 Then
                    str_Date = "20" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 1, 2) & "-" & Mid(str_Date, 3, 2) & " " & Mid(str_Date, 7, 2) & ":" & Mid(str_Date, 9, 2)
                Else
                    str_Date = Format(Now, "yyyy-MM-dd hh:MM:ss")
                End If
                If str_QC = "F1" Then
                    If Trim(Mid(strLine, 14, 13)) = "CONT.01" Then '��ú��
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|1014" & IIf(str_Ji = "1", "^1^", "") & "| |Ѫ��|1"
                    ElseIf Trim(Mid(strLine, 14, 13)) = "CONT.02" Then '��ú��
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|2014" & IIf(str_Ji = "1", "^1^", "") & "| |Ѫ��|1"
                    ElseIf IsNumeric(Trim(Mid(strLine, 14, 13))) Then   '������
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|" & Trim(Mid(strLine, 14, 13)) & "^1^" & "| |Ѫ��|0"
                    Else
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|" & Trim(Mid(strLine, 5, 5)) & IIf(str_Ji = "1", "^1^", "") & "| |Ѫ��|1"
                    End If
                Else
                    strData = str_Date & "|" & Val(Mid(strLine, 5, 5)) & IIf(str_Ji = "1", "^1^", "") & "| |Ѫ��|" & str_QC
                End If
            End If
            '��ʼ�������
            strLine = Replace(Mid(strLine, 45), "", "")
            
            For intloop = 0 To Len(strLine) / 10
                If Len(strLine) < 10 Then Exit For
                str_Tmp = Mid(strLine, 1, 10)
                If InStr(str_Tmp, "%") > 0 Then str_Tmp = Replace(str_Tmp, "%", "")
                strData = strData & "|" & Trim(Mid(str_Tmp, 1, 3)) & "|" & Val(Trim(Mid(str_Tmp, 4)))
                
                strLine = Mid(strLine, 11)
            Next
            
            strResult = strData
            If strData <> "" Then
                aData = Split(strData, "|")
                For intloop = 5 To UBound(aData)
                    If intloop Mod 2 = 1 Then
                        If Val(Trim(aData(intloop))) = 12 Then curTP = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 13 Then curALB = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 28 Then curHBA1C = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 29 Then curHB = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 15 Then curTBIL = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 16 Then curDBIL = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 1 Then curALT = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 30 Then curAST = Val(Trim(aData(intloop + 1)))
                    End If
                Next
                If curALB <> 0 And curTP <> 0 Then
                    If curTP >= curALB Then
                        curGLO = curTP - curALB
                        strResult = strResult & "|X002|" & Format(curGLO, "0.000")
                    End If
                    If curGLO <> 0 Then
                        strResult = strResult & "|110|" & Format(curALB / curGLO, "0.000")
                    End If
                End If
                If curHBA1C <> 0 And curHB <> 0 Then
                    strResult = strResult & "|1111|" & Format(curHBA1C * 87.6 / curHB + 2.27, "0.000")
                End If
                
                If curTBIL <> 0 And curDBIL <> 0 Then
                    curIBIL = curTBIL - curDBIL
                    strResult = strResult & "|94|" & Format(curIBIL, "0.000")
                End If
                
                If curAST <> 0 And curALT <> 0 Then
                    strResult = strResult & "|X006|" & Format(curAST / curALT, "0.000")
                End If
            
            End If
        End If
    
    End If
    'strCmd = Chr(&H6)
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
    Dim varItem As Variant, strItem As String, i As Integer, strSampleNO As String
    Dim str�̺� As String, str����  As String, lng_S As Long, lng_E As Long
    '2008-12-29 11:39:00|110|�ƺ���|ѪҺ|0| |0,10| | | |29|44
    On Error GoTo errHandle
    
    blnSuccess = True
    If iSendStep = 0 Then
        
        If strSample <> "" Then
            strItem = String(88, "0")
            
            varItem = Split(strSample, "|")
            
            strSampleNO = Val(varItem(1))
            strSampleNO = Space(5 - Len(strSampleNO)) & strSampleNO
            
            If InStr(varItem(6), ",") > 0 Then
                blnSuccess = True
                
                str�̺� = Split(varItem(6), ",")(0)
                str���� = Split(varItem(6), ",")(1)
                str���� = Space(3 - Len(str����)) & str����
                
                SendSample = ";N1" & strSampleNO & str�̺� & str���� & Space(29)
        
                        
                For i = 10 To UBound(varItem)
                    If varItem(i) <= 88 And varItem(i) > 0 Then
                        strItem = Mid(strItem, 1, varItem(i) - 1) & "1" & Mid(strItem, varItem(i) + 1)
                    End If
                Next
                If strItem <> String(88, "0") Then
                    SendSample = SendSample & "88" & strItem & "00000"
                End If
            Else
                WriteLog "sendSample", "�̺�,���Ŵ���,ӦΪ0,1�ĸ�ʽ", varItem(6)
            End If
            If SendSample <> "" Then
                If mstrSend <> "" Then
                    lng_S = InStr(mstrSend, Chr(&H2))
                    lng_E = InStr(mstrSend, Chr(&H3))
                    SendSample = Mid(mstrSend, lng_S, lng_E - lng_S + 1)
                    mstrSend = Mid(mstrSend, lng_E + 1)
                End If
                mstrSend = mstrSend & SendSample
            
                If mLastSend = SendSample Then
                    SendSample = ""
                Else
                    mLastSend = SendSample
                End If
                iSendStep = 1
            End If
        End If
    Else
        If strResponse = str_MOR Or strResponse = str_REP Then
            If mstrSend <> "" Then
                lng_S = InStr(mstrSend, Chr(&H2))
                lng_E = InStr(mstrSend, Chr(&H3))
                SendSample = Mid(mstrSend, lng_S, lng_E - lng_S + 1)
                mstrSend = Mid(mstrSend, lng_E + 1)
                If mLastSend = SendSample Then
                    SendSample = ""
                    If mstrSend <> "" Then
                        lng_S = InStr(mstrSend, Chr(&H2))
                        lng_E = InStr(mstrSend, Chr(&H3))
                        SendSample = Mid(mstrSend, lng_S, lng_E - lng_S + 1)
                        mLastSend = SendSample
                    End If
                Else
                    mLastSend = SendSample
                End If
            End If
        End If
        
        If mstrSend = "" Then iSendStep = 0
    End If
    WriteLog "SendSample", strSample, "SendSample=" & SendSample & vbNewLine & "iStep=" & iSendStep & vbNewLine & "mStrsend=" & mstrSend & vbNewLine & "mLastSend=" & mLastSend
    Exit Function
errHandle:
    WriteLog "sendSample", strSample, Err.Description
End Function


Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '�������ļ�
    'strIniFile :ini�����ļ���
    'strItems   :Ҫ��ȡ���ֶΣ��磺DATABASE NAME=th_xlb�е� DATABASE NAME,����ֶ���,�ָ�
    'strReturn  :��ȡ�ķ���ֵ����:DATABASE NAME=th_xlb�е� th_xlb ���������ֵ��,�ָ�
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
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
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If Trim(arrLine(0)) = arrItems(intloop, 0) Then
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