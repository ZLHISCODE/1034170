VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_VIT_250"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'2009-04-09 �¶��޸�
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
    
    Dim iStrPoint As Long, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aItemData() As String
    Dim i As Long
    Dim strCode As String, lng_S As Long, lng_E As Long
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    '�յ�һ��Ӧ��
    'If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, "!000a")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, "h")
        If lng_E > 0 Then
            If Mid(mstrCode, lng_E - 4, 1) = "!" Then
                lng_E = InStr(lng_E, mstrCode, Chr(&HA))
            Else
                lng_E = 0
            End If
        End If
    Else
        lng_E = 0
    End If
    Do While lng_S > 0 And lng_E > 0
        strCode = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "V250 ����:", strCode, ""
    '----------------------------------------------------------------------
        iStrPoint = 1: blnReadItem = False
        Do While iStrPoint <= Len(strCode)
            If Not blnReadItem Then '��ǰ�Ƿ����ڶ�ȡ������Ŀ
                '������Ϣ�Ŀ�ʼλ��
                iPos = iStrPoint
                strBuffer = LTrim(ReadLine(strCode, iPos))
                If strBuffer Like "*!000*" Then
                    '��¼һ���������
                    iCurrGrpEntry = iStrPoint
                    '�����Ƿ�����
                    If ifEnd(strCode, iStrPoint) Then
                        '��Ϣ���������ٱ���
                        iCurrGrpEntry = 0
                        
                        iStrPoint = iPos
                        aItemData = Split(strBuffer, " ")
                        '����ʱ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(Mid(aItemData(2), 7, 2) & "-" & Mid(aItemData(2), 9, 2) & "-" & _
                            Mid(aItemData(2), 11, 2), "yyyy-mm-dd") & " " & Mid(aItemData(2), 1, 2) & ":" & _
                            Mid(aItemData(2), 3, 2) & ":00"
                                
                        '������ţ�����Ϊ����
                        strData = strData & "|" & Val(Mid(aItemData(2), 13))
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ"
                        '�ʿ�Ʒ
                        strData = strData & "|0"
                    
                        '׼������
                        blnReadItem = True
                    Else
                        Exit Do
                    End If
                Else
                    iStrPoint = iPos
                End If
            Else
                strBuffer = LTrim(ReadLine(strCode, iStrPoint))
                If Len(strBuffer) > 0 Then
                    If InStr(strBuffer, " ") > 0 Then '������
                        aItemData = Split(strBuffer, " ")
                        iPos = InStr(aItemData(0), "!")
                        strData = strData & "|" & UCase(Mid(aItemData(0), iPos + 5, 4))
                        If InStr("+-", Right(strData, 1)) > 0 Then strData = Mid(strData, 1, Len(strData) - 1)
                        If Len(aItemData(0)) > 10 Then
                            strData = strData & "|" & Val(Mid(aItemData(0), 11))
                        Else
                            strData = strData & "|" & Val(Mid(strBuffer, Len(aItemData(0)) + 1))
                        End If
                    Else
                        'һ���걾����
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        blnReadItem = False
                    End If
                End If
            End If
        Loop
    '----------------------------------------------------------------------
        lng_S = InStr(mstrCode, "!000a")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, "h")
            If lng_E > 0 Then
                If Mid(mstrCode, lng_E - 4, 1) = "!" Then
                    lng_E = InStr(lng_E, mstrCode, Chr(&HA))
                Else
                    lng_E = 0
                End If
            End If
        Else
            lng_E = 0
        End If
    Loop
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "") As String
    '������Ҫ���������͵ı걾��Ϣ
    '   �걾֮����||�ָ�
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

Private Function ReadSubString(ByVal strSource As String, iPoint As Long, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function

Private Function ReadLine(ByVal strSource, iPoint As Long) As String
    '��ָ��λ�ÿ�ʼ��ȡһ�У�ͬʱָ�붨λ����һ������֮��
    Dim iPos As Long
    
    iPos = iPoint
    Do While iPos <= Len(strSource)
        If Mid(strSource, iPos, 1) = Chr(13) Then
            Exit Do
        End If
        iPos = iPos + 1
    Loop
    If iPos > Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function ifEnd(ByVal strSource, ByVal lngBeginPos As Long) As Boolean
    '���ҽ�����Ƿ����꣺�����һ���Ƿ��пո�
    Dim iPos As Long, strTmp As String
    
    ifEnd = False
    For iPos = Len(strSource) - 1 To lngBeginPos Step -1
        If Mid(strSource, iPos, 1) = Chr(13) Then Exit For
    Next
    If iPos < lngBeginPos Then
        Exit Function
    Else
        strTmp = Mid(strSource, iPos + 1)
        If Len(strTmp) > 0 And InStr(strTmp, " ") = 0 Then ifEnd = True
    End If
End Function