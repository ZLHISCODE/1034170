VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_U1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:������1200
'����:
'����:����ǰ
'�û�:��Ϫ����ͷ����Ժ

Private mstrSource As String
Private mstr�걾�� As String
Private mstr���� As String

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
    Dim strLine As String, str���� As String, str�걾�� As String, str��Ŀ As String, str��� As String, str��Ŀ��� As String
    Dim str�ʿ� As String
    Dim i As Long
    Dim brr()  As String
    Dim arr()  As String
    
    mstrSource = mstrSource & strSource
    
    On Error GoTo errHand
    Do While InStr(mstrSource, "MSH") > 0
        If InStr(mstrSource, "MSH") > 0 Then
            strLine = Mid(mstrSource, 1, InStr(mstrSource, "MSH") - 1) '��ȡ�걾��ʼǰ���� ��ֹ����һ���걾δ��ȫ��������
            If Len(strLine) > 0 Then
                If mstr���� <> "" Then
                    arr = Split(strLine, Chr(13)) '������chr(13)�ָ�
                    For i = LBound(arr) To UBound(arr)
                        brr = Split(arr(i), "|")
                        If UBound(brr) > 2 Then
                            Select Case brr(0)
                                Case "MSH" '�϶��޴˶�����
                                    str���� = brr(6)
                                    str���� = Mid(brr(6), 1, 4) & "-" & Mid(brr(6), 5, 2) & "-" & Mid(brr(6), 7, 2) & " " & Mid(brr(6), 9, 2) & ":" & Mid(brr(6), 11, 2) & ":" & Mid(brr(6), 13, 2)
                                    mstr���� = str����
                                Case "OBR"
                                    str�걾�� = brr(3)
                                    mstr�걾�� = str�걾��
                                Case "OBX"
                                    str��Ŀ = Mid(brr(3), InStr(brr(3), "^") + 1)
                                    str��� = Trim(brr(5))
                                   If InStr("D_Vc,D_WBC,D_KET,D_BIL,D_PRO,D_GLU,D_BLD", str��Ŀ) > 0 Then
                                       str��� = Trim(Mid(str���, 1, 2))
                                       str��� = Replace(Replace(Replace(Replace(Replace(str���, "+-", "��"), "+2", "++"), "+3", "+++"), "+4", "++++"), "+5", "+++++")
                                   End If
                                   str��Ŀ��� = str��Ŀ��� & "|" & str��Ŀ & "|" & str���
                                Case Else
                                '������
                            End Select
                        End If
                    Next i
                    If str��Ŀ��� <> "" Then
                        str�ʿ� = "0"
                        If Val(str�걾��) = "999" Then str�ʿ� = "1"  '�ʿر걾��ָ���걾�����ʿ�
                        
                        strResult = strResult & "||" & mstr���� & "|" & mstr�걾�� & "| |��|" & str�ʿ� & str��Ŀ���
                    End If
                End If
    
            End If
            mstrSource = Mid(mstrSource, InStr(mstrSource, "MSH")) '��ȡ�µı걾��ʼ
            mstr���� = "": mstr�걾�� = "" '����ϸ��걾����Ϣ
            If InStr(3, mstrSource, "MSH") > 0 Then '�ж�ʣ�������Ƿ���
                strLine = Mid(mstrSource, 1, InStr(3, mstrSource, "MSH") - 1)  '���ʱֻȡһ�������ı걾
                mstrSource = Mid(mstrSource, InStr(3, mstrSource, "MSH"))  'ʣ��������¸�ѭ������
            Else
                strLine = mstrSource 'ȡ��ǰ��Ϣ
                mstrSource = ""
            End If
                arr = Split(strLine, Chr(13))
                For i = LBound(arr) To UBound(arr)
                    brr = Split(arr(i), "|")
                    If UBound(brr) > 2 Then
                        Select Case brr(0)
                            Case "MSH"
                                str���� = brr(6)
                                str���� = Mid(brr(6), 1, 4) & "-" & Mid(brr(6), 5, 2) & "-" & Mid(brr(6), 7, 2) & " " & Mid(brr(6), 9, 2) & ":" & Mid(brr(6), 11, 2) & ":" & Mid(brr(6), 13, 2)
                                mstr���� = str����
                            Case "OBR"
                                str�걾�� = brr(3)
                                mstr�걾�� = str�걾��
                            Case "OBX"
                                str��Ŀ = Mid(brr(3), InStr(brr(3), "^") + 1)
                                str��� = Trim(brr(5))
                               If InStr("D_Vc,D_WBC,D_KET,D_BIL,D_PRO,D_GLU,D_BLD", str��Ŀ) > 0 Then
                                   str��� = Trim(Mid(str���, 1, 2))
                                   str��� = Replace(Replace(Replace(Replace(Replace(str���, "+-", "��"), "+2", "++"), "+3", "+++"), "+4", "++++"), "+5", "+++++")
                               End If
                               str��Ŀ��� = str��Ŀ��� & "|" & str��Ŀ & "|" & str���
                            Case Else
                            '������
                        End Select
                    End If
                Next i
                If str��Ŀ��� <> "" Then
                    str�ʿ� = "0"
                    If Val(str�걾��) = "999" Then str�ʿ� = "1"  '�ʿر걾��ָ���걾�����ʿ�
                    strResult = strResult & "||" & mstr���� & "|" & mstr�걾�� & "| |��|" & str�ʿ� & str��Ŀ���
                End If
        End If
    Loop
    If strResult <> "" Then
        strResult = Mid(strResult, 3) 'ȥ����ǰ��� ||
        strResult = Replace(Replace(strResult, "|-|", "|- |"), "|+|", "|+ |")
    End If
    Exit Sub
errHand:
    WriteLog "clsLISDev_U1200", Err.Number & ":" & Err.Description, ""
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











