VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_Spife3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:spife 3000
'����:����Helena
'����:�¶�
'�û�:3201ҽԺ

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, strImg As String
    Dim lngS As Long, lngE As Long
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    Do While mstrCode Like "*H|\^&*L|1*"
        '----- ��ʼ�������

        'ȡһ����������Ϣ
        lngS = InStr(mstrCode, "H|\")
        lngE = InStr(lngS, mstrCode, Chr(&HD) & "L|1")
        strLine = Mid(mstrCode, lngS, lngE - lngS + 4)
        mstrCode = Mid(mstrCode, lngE + 4)
        
        str���� = ""
        '--�������
        varline = Split(strLine, Chr(&HD))
        For iRow = LBound(varline) To UBound(varline)
            If varline(iRow) Like "P|*" Then
                varItem = Split(varline(iRow), "|")
'                    str�걾�� = Val(Split(varItem(3), "^")(2))
                str�걾�� = Val(varItem(4))
                str�����˱걾�ʿ� = " |ѪҺ|0"
                strItem = ""
                
            ElseIf varline(iRow) Like "R|*^^^*^*" Then
                varItem = Split(varline(iRow), "|")
                
                strTmp = Replace(varItem(2), "^^^", "")
                If InStr(strTmp, "TraceValues") > 0 Then
                    strImg = varItem(3)
                Else
                    If UBound(Split(strTmp, "^")) > 1 Then
                        If Split(strTmp, "^")(2) = "%" Then
                            strItem = strItem & "|" & Split(strTmp, "^")(1) 'ͨ����
                            strTmp = varItem(3)
                            strItem = strItem & "|" & strTmp
                            If str���� = "" Then
                                str���� = Trim(varItem(11))
                                If IsNumeric(str����) And Len(str����) > 12 Then
                                    str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2) & " " & Mid(str����, 9, 2) & ":" & Mid(str����, 11, 2)
                                Else
                                    str���� = Format(Now, "yyyy-MM-dd")
                                End If
                            End If
                        
                        End If
                    Else
                        strItem = strItem & "|" & Split(strTmp, "^")(1) 'ͨ����
                        strTmp = varItem(3)
                        strItem = strItem & "|" & strTmp
                        If str���� = "" Then
                            str���� = Trim(varItem(11))
                            If IsNumeric(str����) And Len(str����) > 12 Then
                                str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2) & " " & Mid(str����, 9, 2) & ":" & Mid(str����, 11, 2)
                            Else
                                str���� = Format(Now, "yyyy-MM-dd")
                            End If
                        End If
                        
                    End If

                End If
            End If
        Next
        If strItem <> "" Then
            strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
        End If
        If strImg <> "" Then
            strImg = Replace(strImg, "^", ";")
            strImg = Replace(strImg, "0.", "")
            strResult = strResult & vbCrLf & "��;0;100,200,0-0,25-,50-5,75-,100-10,125-,150-15,175-;" & strImg
        End If

    Loop
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
    
End Function

