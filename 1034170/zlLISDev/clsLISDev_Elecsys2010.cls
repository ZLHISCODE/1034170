VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Elecsys2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Elecsys 2010
'����:����
'����:�¶�
'�û�:����ҽԺ ˫��

Private strDeviceString '������Ϣ���豸��
Private strLastString As String '�ϴη��͵��豸��
Private mStrλ�� As String  '˫��ͨ���á�
Private strTmp As String

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str���� As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Dim lngS As Long, lngE As Long
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
        strReserved = Mid(strSource, InStr(strSource, Chr(&H5)) + 1)
    Else
        strReserved = strSource
        If Replace(strReserved, strTmp, "") Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
            strCmd = Chr(&H6)
            strTmp = strReserved
        End If
    End If
    
    If strSource Like "*" & Chr(2) & "1H|\^&**" Then
        '----- ��ʼ�������
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ����������Ϣ
            lngS = InStr(strReserved, Chr(2) & "1H")
            lngE = InStr(strReserved, "")
            strLine = Mid(strReserved, lngS, lngE - lngS)
            strReserved = Mid(strReserved, lngE + 1)
            If strReserved = Chr(4) Then strReserved = ""
            If strLine Like "*?Q|*L|1*" Then
                If strResult <> "" Then
                    '���ѽ����Ľ���ȷ���,��������˫��Ľ��
                    strReserved = strLine & strReserved
                    blnEnd = True
                    Exit Sub
                End If
                '˫��ͨ��: �յ����������ı걾�ţ�����š�
                varLine = Split(strLine, vbCrLf)
                For iRow = LBound(varLine) To UBound(varLine)
                    If varLine(iRow) Like "?Q|*" Then
                        varItem = Split(varLine(iRow), "|")
                        str���� = Split(varItem(2), "^")(1)
                        str�걾�� = Val(Split(varItem(2), "^")(2))
                        str���� = Format(Now, "yyyy-MM-dd")
                        
                        mStrλ�� = Replace(varItem(2), "^" & str���� & "^", "") 'SendSample������Ҫ�á�
                        mStrλ�� = Mid(mStrλ��, 1, InStr(mStrλ��, "^^") - 1)
                        strResult = str���� & "|" & str�걾�� & "^0^" & str����
                        strCmd = "1|"
                    End If
                Next
                blnEnd = True
                Exit Sub
            End If
            
            '--�������
            varLine = Split(strLine, vbCrLf)
            For iRow = LBound(varLine) To UBound(varLine)
                If varLine(iRow) Like "?O|*" Then
                    varItem = Split(varLine(iRow), "|")
'                    str�걾�� = Val(Split(varItem(3), "^")(2))
                    str�걾�� = Val(varItem(2))
                    str���� = Mid(varItem(6), 1, 8)
                    str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2)
                    str�����˱걾�ʿ� = " |ѪҺ|0"
                    strItem = ""
                    
                ElseIf varLine(iRow) Like "?R|*^^^*" Then
                    varItem = Split(varLine(iRow), "|")
                    strTmp = Split(varItem(2), "^")(3)
                    strItem = strItem & "|" & strTmp  'ͨ����
                    strTmp = varItem(3)
                    If InStr(strTmp, "^") > 0 Then
                        '���ԵĽ��
                        strItem = strItem & "|" & Mid(strTmp, InStr(strTmp, "^") + 1)
                        
                    Else
                        '�����Ľ��
                        strItem = strItem & "|" & IIf(Trim(strTmp) = "", " ", strTmp)
                    End If
                End If
            Next
            If strItem <> "" Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            If Not strReserved Like "*1H|\^&**" Then
                blnEnd = True
            End If
        Loop
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
    
    Dim varItem As Variant, varRecord As Variant, str�����Ŀ As String
    Dim i As Integer
    If iSendStep = 0 Then
        SendSample = Chr(5)
        iSendStep = iSendStep + 1
        Exit Function
    End If
    
    If strResponse Like "*" & Chr(6) & "*" Then
        If iSendStep = 1 Then
            '������Ϣͷ: [STX]1H|\^&|||ASTM-Host[CR][ETX]59[CR][LF]
            SendSample = Chr(2) & "1H|\^&|||ASTM-Host" & Chr(13) & Chr(3) & "59" & Chr(13) & Chr(10)
        ElseIf iSendStep = 2 Then
            '���ز�����Ϣ��  [STX]2P|1||000004[CR][ETX]5B[CR][LF]
            varItem = Split(strSample, "|")
            SendSample = Chr(2) & "2P|1||" & varItem(1) & Chr(13) & Chr(3) & "5B" & Chr(13) & Chr(10)
        ElseIf iSendStep = 3 Then
            '����Ҫ������Ŀ : [STX]3O|1|000004|278^0^19|^^^10^0\^^^20^0|R||||||N||||||||||||||O[CR][ETX]58[CR][LF]
            varItem = Split(strSample, "|")
            SendSample = Chr(2) & "3O|1|" & varItem(1) & "|" & mStrλ�� & "|"
            str�����Ŀ = ""
            For i = 10 To UBound(varItem)
                str�����Ŀ = str�����Ŀ & "\^^^" & varItem(i) & "^0"
            Next
            str�����Ŀ = Mid(str�����Ŀ, 2)
            SendSample = SendSample & str�����Ŀ & "|R||||||N||||||||||||||O" & Chr(13) & Chr(3) & "58" & Chr(13) & Chr(10)
        ElseIf iSendStep = 4 Then
            '������Ϣβ: [STX]4L|1[CR][ETX]3D[CR][LF]
            SendSample = Chr(2) & "4L|1" & Chr(13) & Chr(3) & "3D" & Chr(13) & Chr(10)
        ElseIf iSendStep = 5 Then
            '���ؽ����� : 
            SendSample = Chr(4)
        End If
        If iSendStep >= 5 Then
            iSendStep = 0
        Else
            iSendStep = iSendStep + 1
        End If
    End If
    
End Function







