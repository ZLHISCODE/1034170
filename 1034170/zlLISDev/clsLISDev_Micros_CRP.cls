VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Micros_CRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ABX Micros CRP
'����:(δ֪)
'����:�¶�
'�û�:����������Ժ

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
    
    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    str��Ŀ = ",!,@,#,$,%,&,',"",A,B,C,2,3,4,5,6,7,8,"
    If strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            strLine = Mid(strReserved, InStr(strReserved, ""), InStr(strReserved, ""))
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            '����
            varLine = Split(strLine, Chr(&HD&))
            
            For iRow = 0 To UBound(varLine)
                strTmp = varLine(iRow)
                Select Case Left(strTmp, 1)
                    Case "q"  '����
                        str���� = Format(CDate(Mid(strTmp, 9, 2) & "-" & Mid(strTmp, 6, 2) & "-" & Mid(strTmp, 3, 2)), "yyyy-MM-dd")
                    Case "u"  '�걾��
                        str�걾�� = Val(Mid(strTmp, 3))
                        str�����˱걾�ʿ� = " |ѪҺ|0"
                    Case Else
                        If InStr(str��Ŀ, "," & Left(strTmp, 1) & ",") > 0 Then
                            strValues = Mid(strTmp, 3, 5)
                            If strValues <> "--.--" Then
                                strItem = strItem & "|" & SwapWord(Left(strTmp, 1)) & "|" & Val(strValues)
                            End If
                        End If
                End Select
            
            Next
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            If Not strReserved Like "***" Then
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
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function

Private Function SwapWord(ByVal strWord) As String
    '���������ص������ַ�ת��Ϊ��ͨ�ַ�,���ⷵ�ص�HIS�г���.
    Select Case strWord
    Case "!"
        SwapWord = "1"
    Case "@"
        SwapWord = "9"
    Case "#"
        SwapWord = "D"
    Case "%"
        SwapWord = "E"
    Case "'"
        SwapWord = "F"
    Case """"
        SwapWord = "G"
    Case "$"
        SwapWord = "H"
    Case "&"
        SwapWord = "I"
    Case Else
        SwapWord = strWord
    End Select
End Function
