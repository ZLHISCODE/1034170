Attribute VB_Name = "mdl����"
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--�������ӿ�
    '����˵��:
    '   msgType-ҵ����������,�����µĲ�����
    '   packageType-���ݽ�����ʽ���ͣ�ϵͳ��������ʱʹ��,�����µĲ�����
    '   packageLength-���ݴ��ĳ���,�����µĲ�����
    '   str-���ݴ�,����ʱ��ͨ�����ݴ������������������ʱ�����ݴ��а������ص�����
    '   strCom:�������󴮿ڣ����ݶ��������λ�ã�����������ȡֵ��'com1','com2')
    '����:
    '   I.  ����������ֵ����0ʱ����ʾ�ɹ����ַ����а�����ҵ����󷵻ص�����
    '   II. ����������ֵ������0ʱ���μ��������һ����Ӧ����Ҫ�����������Ȼ������ʵ��Ĵ���

'��������������
Private Declare Function IC_Read_Base Lib "ICCNII32.DLL" (ByVal szData As String) As Long
Private Declare Function IC_Read_Plus Lib "ICCNII32.DLL" (nSequence As Long, ByVal szData As String) As Long
    
Private Declare Function KfqTransData Lib "OltpTransKfq03.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    
'--��ͨ�ӿ�
Private Declare Function OltpTransData Lib "OltpTransIc03.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
     ByVal str As String, ByVal strCom As String) As Long
'����Ϊ�������Ĳ�����
'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 95                     ʵʱ�鿨���������鿨��
'1002            12                  420                    ʵʱ����
'1003            7                   297                    ʵʱҽ����ϸ�����ύ
'1004            9                   136                    ʵʱסԺ�Ǽ������ύ
'1006            12                  420                    ʵʱ����Ԥ��
'1008            101                 95                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�

'����Ϊ�����еĲ�����
'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 94                     ʵʱ�鿨���������鿨��
'1002            12                  424                    ʵʱ����
'1003            7                   230                    ʵʱҽ����ϸ�����ύ
'1004            9                   206                    ʵʱסԺ�Ǽ������ύ
'1006            12                  424                    ʵʱ����Ԥ��
'1008            101                 94                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�
'1005            8                   274                    ʵʱҽ������
'1007            2                   55                     �����ʻ���ѯ
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public gblnKFQCom_����  As Boolean   'true-�������ӿ�,False-��ͨ�ӿ�

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Public g�������_���� As �������
Private Type �������
    ���˱��            As String
    ����                As String
    �Ա�                As String
    ��������            As String
    ����                As Integer
    ���֤��            As String
    IC����              As Long
    �������            As Long
    ְ����ҽ���        As String
    ���������ʻ����    As Double
    ���������ʻ����    As Double
    ͳ���ۼ�            As Double
    �½ɷѻ���          As Double
    �ʻ�״̬            As String
    �α����1           As String
    �α����2           As String
    �α����3           As String
    �α����4           As String
    �α����5           As String
    
    ת�ﵥ��            As String           '�����֤ʱ����
    ҽ������            As Long             '�����֤ʱѡ��,��������
    �������            As Long             '�����֤ʱѡ��,������ǽ��㷽ʽ����
    ֧�����            As Double           '
    ��ϱ���            As String           '��ϱ���ʱ����,������Ч
    �������            As String           '�������ʱ����,������Ч
    
    �����ʻ�ԭʼֵ      As Double          '������ѯ��ȡ
    �����ʻ���ǰֵ      As Double          '������ѯ��ȡ
    �����ʻ�״̬        As Double          '������ѯ��ȡ
    ����              As Double
End Type

Public Const gblnģ��ӿ� = False     'ģ��ӿ�����

Public gstrҽԺ����_���� As String        'ҽԺ����,ֻ��Ϊ4λ
Public gintComPort_���� As Integer
Public gbln������ϸʱʵ�ϴ� As Boolean
Public gblnסԺ��ϸʱʵ�ϴ� As Boolean

Private Function Readģ������(ByVal lng���Ĵ��� As Long, _
        msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
        str As String)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,��������
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strArr
    Dim strArr1
    Dim strText As String
    Dim strTemp As String
    
    If Dir(App.Path & "\����ҽ��\����ҽ��ģ������" & lng���Ĵ��� & ".txt") <> "" Then
            Set objText = objFile.OpenTextFile(App.Path & "\����ҽ��\����ҽ��ģ������" & lng���Ĵ��� & ".txt")
            Do While Not objText.AtEndOfStream
                strTemp = Trim(objText.ReadLine)
                strArr = Split(strTemp, "||")
                strArr1 = Split(strArr(0), "|")
                If Val(strArr1(0)) = msgType Then
                     str = strArr(1)
                     Exit Do
                End If
            Loop
            objText.Close
    End If
    
End Function
Public Function ��ȡ�������_����(ByVal lng���Ĵ��� As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵�������,������Ϣ����g�������_����
    '--�����:lng���Ĵ���(2��������)
    '--������:
    '--��  ��:��ȡ�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    
    Dim strInfor As String
    Dim lngReturn As Long
    Dim int�Ա� As Integer
    ��ȡ�������_���� = False
    Err = 0
    On Error GoTo ErrHand:
    '�ܺ�ȫ���� 2003-12-17
    '����˴�������ո�ֵʱ�������������˴���ֱ���˳�
    strInfor = Space(100)
    If gblnģ��ӿ� Then
        Readģ������ lng���Ĵ���, 1001, 101, 94, strInfor
        If strInfor = "" Then Exit Function
    Else
        If lng���Ĵ��� = 2 Then
            '1001    101 95  ʵʱ�鿨���������鿨��
            lngReturn = KfqTransData(1001, 101, 95, strInfor, "com" & gintComPort_����)
        Else
            '1001    101 94  ʵʱ�鿨���������鿨��
            lngReturn = OltpTransData(1001, 101, 94, strInfor, "com" & gintComPort_����)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn))
            Exit Function
        End If
    End If
    'ȡ���ظ�
    strInfor = Mid(strInfor, 2)
    With g�������_����
        .ҽ������ = lng���Ĵ���
        If lng���Ĵ��� = 2 Then
            .���˱�� = Substr(strInfor, 1, 10) '���˱���    1   10      ���ķ���
            .���� = Substr(strInfor, 11, 8)     '����    11  8       ���ķ���
            .���֤�� = Substr(strInfor, 19, 18)    '���֤��    19  18      ���ķ���
            .IC���� = Substr(strInfor, 37, 7)       'IC����  37  7       ���ķ���
            .������� = Val(Substr(strInfor, 44, 4))    '�������    44  4       ���ķ���
            .ְ����ҽ��� = Substr(strInfor, 48, 1)     'ְ����ҽ���    48  1   A��ְ��B����    ���ķ���
            .���������ʻ���� = Val(Substr(strInfor, 49, 10)) '���������ʻ����    49  10      ���ķ���
            .���������ʻ���� = Val(Substr(strInfor, 59, 10)) '���������ʻ����    59  10      ���ķ���
            .ͳ���ۼ� = Val(Substr(strInfor, 69, 10)) 'ͳ���ۼ�    69  10      ���ķ���
            .�½ɷѻ��� = Val(Substr(strInfor, 79, 10)) '�½ɷѻ���  79  10  �½ɷѹ���  ���ķ���
            .�ʻ�״̬ = Substr(strInfor, 89, 1) '�ʻ�״̬    89  1   A������B��ֹ����Cȫֹ����D����  ���ķ���
            .�α����1 = Substr(strInfor, 90, 1) '�α����1   90  1   �Ƿ����ܸ߶� 1 ���� 0 ������    ���ķ���
            .�α����2 = Substr(strInfor, 91, 1) '�α����2   91  1   �Ƿ����ܲ�������ҵ����������Ա������'0 ������ 1 ��ҵ 2 ����Ա    ���ķ���
            .�α����3 = Substr(strInfor, 92, 1) '�α����3   92  1   0 �󱣡�1 �±�  ���ķ���
            .�α����4 = Substr(strInfor, 93, 1) '�α����4   93  1   ����    ���ķ���
            .�α����5 = Substr(strInfor, 94, 1) '�α����5   94  1   ����    ���ķ���
        Else
            .���˱�� = Substr(strInfor, 1, 8)  '���˱��    CHAR    1   8   ҽ�����    ����
            .���� = Substr(strInfor, 9, 8)      '����    CHAR    9   8       ����
            .���֤�� = Substr(strInfor, 17, 18)    '���֤��    CHAR    17  18  18λ��15λ  ����
            .IC���� = Substr(strInfor, 35, 7)       'IC����  NUM 35  7       ����
            .������� = Val(Substr(strInfor, 42, 4))    '�������    NUM 42  4       ����
            
            '�ܺ�ȫ���� 2003-12-17
            '���룺Q��ҵ����
            .ְ����ҽ��� = Substr(strInfor, 46, 1)     'ְ����ҽ���    CHAR    46  1   A��ְ��B���ݡ�L���ݡ�T���Q��ҵ����  ����
            .���������ʻ���� = Val(Substr(strInfor, 47, 10))   '���������ʻ����    NUM 47  10      ����
            .���������ʻ���� = Val(Substr(strInfor, 57, 10))   '���������ʻ����    NUM 57  10  �����ڹ���Ա��������    ����
            .ͳ���ۼ� = Val(Substr(strInfor, 67, 10))   'ͳ���ۼ�    NUM 67  10      ����
            .�½ɷѻ��� = Val(Substr(strInfor, 77, 10)) '�½ɷѻ���  NUM 77  10  �½ɷѹ���  ����
            .�ʻ�״̬ = Substr(strInfor, 87, 1)         '�ʻ�״̬    CHAR    87  1   A������B��ֹ����Cȫֹ����D����  ����
            .�α����1 = Substr(strInfor, 88, 1)        '�α����1   CHAR    88  1   �Ƿ����ܸ߶�: 0 �����ܸ߶1 ���ܸ߶2 ҽ�Ʊ��ղ�����    ����
            .�α����2 = Substr(strInfor, 89, 1)        '�α����2   CHAR    89  1   �Ƿ����ܲ�������ҵ����������Ա������0 ������ 1 ��ҵ 2 ����Ա    ����
            .�α����3 = Substr(strInfor, 90, 1)        '�α����3   CHAR    90  1   0 �󱣡�1 �±�  ����
            .�α����4 = Substr(strInfor, 91, 1)        '�α����4   CHAR    91  1   0���������á�1��������  ����
            .�α����5 = Substr(strInfor, 92, 1)        '�α����5   CHAR    92  1   0���˲����á�1���˿���  ����
        End If
        int�Ա� = Val(IIf(Len(.���֤��) = 18, Mid(.���֤��, 17, 1), Right(.���֤��, 1))) Mod 2
        '�������֤ȡ����Ӧ���Ա�
        .�Ա� = IIf(int�Ա� = 0, "Ů", "��")
        .�������� = zlCommFun.GetIDCardDate(Trim(.���֤��))
        '��������
        If IsDate(.��������) And .�������� <> "" Then
            .���� = Abs(Int((zlDatabase.Currentdate - CDate(.��������)) / 365))
        Else
            .���� = 0
        End If
        
    End With
    ��ȡ�������_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    ��ȡ�������_���� = False
End Function

Public Function ҵ������_����( _
            ByVal lng���Ĵ��� As Long, _
            ByVal lngMsgType As Long, _
            strTans As String _
    ) As Boolean
    
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����ص�ҵ������,��������Ӧ�Ľ��
    '--�����:lng���Ĵ���(2��������)
    '   lngMsgType-ҵ����������
    '   lngPackageType-���ݽ�����ʽ����
    '   lngPackageLength-���ݴ��ĳ���
    '   strTans-���ݴ�,����ʱ��ͨ�����ݴ������������������ʱ�����ݴ��а������ص�����
    '����:
    '   �ɹ�-true,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngPackageType As Long
    Dim lngPackageLength As Long
    Dim i As Long
    Dim strTmp As String
    
    i = lngMsgType
    
    '����Ϊ�������Ĳ�����
    'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 95                     ʵʱ�鿨���������鿨��
    '1002            12                  420                    ʵʱ����
    '1003            7                   297                    ʵʱҽ����ϸ�����ύ
    '1004            9                   136                    ʵʱסԺ�Ǽ������ύ
    '1006            12                  420                    ʵʱ����Ԥ��
    '1008            101                 95                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�
    
    '����Ϊ�����еĲ�����
    'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 94                     ʵʱ�鿨���������鿨��
    '1002            12                  424                    ʵʱ����
    '1003            7                   230                    ʵʱҽ����ϸ�����ύ
    '1004            9                   206                    ʵʱסԺ�Ǽ������ύ
    '1006            12                  424                    ʵʱ����Ԥ��
    '1008            101                 94                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�
    '1005            8                   274                    ʵʱҽ������
    '1007            2                   55                     �����ʻ���ѯ
    
    Dim strInfor As String
    Dim lngReturn As Long
    ҵ������_���� = False
    Err = 0
    On Error Resume Next
    If lng���Ĵ��� = 2 Then
        strTmp = Switch(i = 1001, "101|95", i = 1002, "12|420", i = 1003, "7|297", i = 1004, "9|136", i = 1006, "12|420", _
            i = 1008, "101|95")
        If Err <> 0 Then
            strTmp = "|"
        End If
    Else
            strTmp = Switch(i = 1001, "101|94", i = 1002, "12|424", i = 1003, "7|230", i = 1004, "9|206", i = 1006, "12|424", _
                i = 1008, "101|94", i = 1005, "8|274", i = 1007, "2|55")
        If Err <> 0 Then
            strTmp = "|"
        End If
    End If
    lngPackageType = Val(Split(strTmp, "|")(0))
    lngPackageLength = Val(Split(strTmp, "|")(1))
    
    Err = 0
    On Error GoTo ErrHand:
    strInfor = strTans
    If gblnģ��ӿ� Then
        Readģ������ lng���Ĵ���, lngMsgType, lngPackageType, lngPackageLength, strInfor
        If strInfor = "" Then
            strTans = strInfor
            Exit Function
        End If
    Else
        '������˵,������ҵ�����͵������ж���ǰ�ӿո�.�����ؼ���:" " &
        strInfor = " " & strInfor
        If lng���Ĵ��� = 2 Then
            lngReturn = KfqTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_����)
        Else
            lngReturn = OltpTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_����)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn))
            strTans = ""
            Exit Function
        End If
    End If
    'ȡ���ظ�
    strInfor = Mid(strInfor, 2)
    
    strTans = strInfor
    ҵ������_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    strTans = ""
    ҵ������_���� = False
End Function


Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '--�����:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '--������:
    '--��  ��:�Ӵ�
    '-----------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    
    Substr = Trim(StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode))
    Exit Function
ErrHand:
    Substr = ""
End Function

Public Function ҽ����ʼ��_����() As Boolean

    Dim rsTemp  As New ADODB.Recordset
    Dim strReg As String
    
    '���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
    '���أ���ʼ���ɹ�������true�����򣬷���false
    
    On Error Resume Next
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & gintInsure
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡҽԺ����")
    gstrҽԺ����_���� = NVL(rsTemp!ҽԺ����, "")
    
    '���ö˿ں�
    Call GetRegInFor(g����ģ��, "����", "�˿ں�", strReg)

    If Val(strReg) = 0 Then
        gintComPort_���� = 1
    Else
        gintComPort_���� = IIf(Val(strReg) > 99, 1, Val(strReg))
    End If
    
    Call GetRegInFor(g����ģ��, "����", "������", strReg)
    
    If gintInsure = TYPE_���������� Then
        gblnKFQCom_���� = True
    Else
        gblnKFQCom_���� = False
    End If
    '�����ϴ���ϸ����
    gstrSQL = "Select * From ���ղ��� where ������ in ('������ϸʱʵ�ϴ�','סԺ��ϸʱʵ�ϴ�') and ����=" & gintInsure
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ղ���"
    gbln������ϸʱʵ�ϴ� = True
    gblnסԺ��ϸʱʵ�ϴ� = True
    Do While Not rsTemp.EOF
        Select Case NVL(rsTemp!������)
        Case "������ϸʱʵ�ϴ�"
            gbln������ϸʱʵ�ϴ� = IIf(Val(NVL(rsTemp!����ֵ)) = 1, True, False)
        Case "סԺ��ϸʱʵ�ϴ�"
            gblnסԺ��ϸʱʵ�ϴ� = IIf(Val(NVL(rsTemp!����ֵ)) = 1, True, False)
        End Select
        rsTemp.MoveNext
    Loop
    ҽ����ʼ��_���� = True
End Function

Public Function �������_����(ByVal lng����id As Long) As Currency
    '����: ���ݲ���idȡ�����
    '����: ����id
    '����: ���ظ����ʻ����
    Dim rsAcc As New ADODB.Recordset
    
    
    '����ʧ�����˳�
    gstrSQL = "Select Nvl(�ʻ����,0) �ʻ����,����֤�� From �����ʻ� Where ����=" & gintInsure
    gstrSQL = gstrSQL & " And ����id=" & lng����id
    
    Call OpenRecordset(rsAcc, "��ȡ�ʻ����")
    
    With g�������_����
        .���������ʻ���� = NVL(rsAcc!�ʻ����, 0)
        .���������ʻ���� = Val(NVL(rsAcc!����֤��))
        �������_���� = .���������ʻ����
    End With
End Function

Public Function ҽ������_����(ByVal lng���� As Long, ByVal lngҽ������ As Integer) As Boolean
    ҽ������_���� = frmSet����.ShowME(lng����, lngҽ������)
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����id As Long) As String
    Dim str��ע As String, rsPatient As New ADODB.Recordset
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    
    ��ݱ�ʶ_���� = frmIdentify����.GetPatient(bytType, lng����id)
End Function
Public Function ��ݱ�ʶ_����2(ByVal strCard As String, ByVal strPass As String, Optional lng����id As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    ��ݱ�ʶ_����2 = frmIdentify����.GetPatient(3, lng����id)
End Function

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = strTmp
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = Len(strCode)
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    End If
    Rpad = strTmp
End Function
Private Function Get�������(ByVal bytҵ�� As Byte, ByVal int���� As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��������ʶ
    '--�����:bytҵ��-(0-����,1-����)
    '         int���� ����:(1-��ͨ����,2-��������,3-�����,4-������������)
    '                 סԺ:(5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ)
    '--������:
    '--��  ��:ҽ�����ĵķ����ʶ
    '-----------------------------------------------------------------------------------------------------------
    'ҽ�����ĵľ������Ķ�Ӧϵͳ
    
    '1 �������
    'A ����������
    '3 �������
    '7 ����������
    '5 ����󲡽���
    'B ����󲡳���
    'S ������������
    'T ������������
    
    '2 סԺ����
    'D סԺ�����
    '9 סԺ�岹��  �˹����ݲ���
    '4 ��ͥ��������
    'C ��ͥ���������
    '8 ��ͥ��������     '�˹����ݲ���
    'O ��������סԺ����
    'P ��������סԺ����
    'Q ���˱��ս���
    'R ���˱��ճ���


    Dim i As Integer
    Dim strTmp As String
    i = int����
    
    '���˺��ע:200404
    '     ����:1-1,2-3,3-5,4-"S"
    '     סԺ:5-2,6-4,7-"O",8-"Q"
            
            
    Select Case int����
        Case 1  '1-��ͨ����
            strTmp = Decode(bytҵ��, 0, "1", "A")
        Case 2  '2-��������
            strTmp = Decode(bytҵ��, 0, "3", "7")
        Case 3  '3-�����
            strTmp = Decode(bytҵ��, 0, "5", "B")
        Case 4  '4-������������
            strTmp = Decode(bytҵ��, 0, "S", "T")
        Case 5  '5-��ͨסԺ,
            strTmp = Decode(bytҵ��, 0, "2", "D")
        Case 6  '6-��ͥ����סԺ,
            strTmp = Decode(bytҵ��, 0, "4", "C")
        Case 7  '7-��������סԺ
            strTmp = Decode(bytҵ��, 0, "O", "P")
        Case 8  '8-���˱���סԺ
            strTmp = Decode(bytҵ��, 0, "Q", "R")
        Case Else
            strTmp = ""
    End Select
    Get������� = strTmp
End Function
Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    Dim curTotal As Currency, cur�����ʻ� As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl���� As Double
    Dim dbl���Ʒ� As Double
    Dim dbl���� As Double
    Dim dbl����Է� As Double
    Dim dbl�������Ʒ� As Double
    Dim dbl���������Է� As Double
    Dim dbl�������Էѷ��� As Double
    Dim dbl�Ǳ��շ��� As Double
    Dim dblͳ����� As Double
    Dim dbl������ As Double     '��Դ�����������
    Dim dbl�𸶱�׼ As Double
    
    Dim lng����id As Long
    
    Dim str��ϱ��� As String  '��������
    Dim strҽʦ���� As String
    Dim str����Ա���� As String
    Dim str������� As String
    Dim str���������ʶ As String
    Dim strTmp As String
    Dim strҽ�� As String
    Dim str��ϸ As String       '��ϸ��
    Dim str���ұ��� As String
    Dim dbl���� As Double
    Dim str��Ŀͳ�Ʒ��� As String
    Dim str��Ŀ���� As String
    Dim dbl��Ŀ���� As Double
    
    
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '��ϸ�ֶ�
    '   ����ID,�շ����,�վݷ�Ŀ,���㵥λ,������,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��,ժҪ,�Ƿ���
    
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ�
    
    '������֧��������ڱ���,�Ա���㱣�Ѽ��Է�
    gstrSQL = "Select * From ����֧������"
    zlDatabase.OpenRecordset rs����, gstrSQL, "����֧������"
    
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lng����ID As Long
    With rs��ϸ
        'ȷ������
        If Not .EOF Then
            lng����id = NVL(!����ID, 0)
            gstrSQL = "  select ����id from �����ʻ� where ����id=" & lng����id & "  and ����=" & gintInsure & "  and ҽ����='" & g�������_����.���˱�� & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
            If Not rsTemp.EOF Then
                lng����ID = NVL(rsTemp!����ID, 0)
            Else
                lng����ID = 0
            End If
          '����׼��Ŀ
            gstrSQL = "Select * from ������׼��Ŀ  where ����ID=  " & lng����ID
            zlDatabase.OpenRecordset rs��׼��Ŀ, gstrSQL, "��ȡ������Ŀ����"
            
        End If
        
        'ȡ�����η������õĽ��ϼ�
        Do While Not .EOF
            '---��˳��,�Խ���Ƿ�Ϊ���������ж�,���Ϊ������׼ִ��ҽ���շ�
            If !ʵ�ս�� < 0 Then
                ShowMsgbox "�õ����а����н��Ϊ��������Ŀ,����ִ��ҽ���շ�!����������շ�"
                �����������_���� = False
                Exit Function
            End If
            
            If lng����ID <> 0 Then
                    '��һ��,ȷ��������շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=1 and �շ�ϸĿid=" & NVL(!�շ�ϸĿID, 0)
                    If rs��׼��Ŀ.EOF Then
                        gstrSQL = "Select ����,���� from �շ�ϸĿ where id=" & NVL(!�շ�ϸĿID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ�ϸĿ"
                        ShowMsgbox "�շ�ϸĿΪ��" & NVL(rsTemp!����) & "������Ŀ���ǲ��������趨����Ŀ."
                        Exit Function
                    End If
                    
                    '�ڶ���,ȷ������ı��մ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=1 and  �շ�ϸĿid=" & NVL(!����֧������ID, 0)
                    If rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�ڽ����д����˽�������ı���֧������,���ܼ�����"
                        Exit Function
                    End If
                    '������,'ȷ����ֹ���շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=2 and �շ�ϸĿid=" & NVL(!�շ�ϸĿID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        gstrSQL = "Select ����,���� from �շ�ϸĿ where id=" & NVL(!�շ�ϸĿID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ�ϸĿ"
                        ShowMsgbox "�շ�ϸĿΪ��" & NVL(rsTemp!����) & "������Ŀ�Ǳ���ֹʹ�õ���Ŀ." & vbCrLf & "���ܼ���!"
                        Exit Function
                    End If
                    '���Ĳ�,'ȷ����ֹ�Ĵ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=2 and �շ�ϸĿid=" & NVL(!����֧������ID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�ڽ����д����˽�ֹʹ�õı���֧������,���ܼ�����"
                    End If
            End If
        
            '���ж��Ƿ�������ҽ����Ӧ��Ŀ����
            gstrSQL = " Select ��Ŀ����,��Ŀ���� From ����֧����Ŀ" & _
                      " Where ����=" & gintInsure & " And �շ�ϸĿID=" & !�շ�ϸĿID
                      
            Call OpenRecordset(rsTemp, "�ж��Ƿ������˶�Ӧ��ҽ����Ŀ")
            If rsTemp.EOF = True Then
                MsgBox "����Ŀδ����ҽ����Ŀ�����ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            If strҽ�� = "" Then
                strҽ�� = NVL(!������)
            End If
            
            str��Ŀ���� = NVL(rsTemp!��Ŀ����)
            dbl��Ŀ���� = Val(NVL(rsTemp!��Ŀ����))
            lng����id = NVL(!����ID, 0)
            gstrSQL = "" & _
                " Select b.������,b.����ֵ from �շ���� a,���ղ��� b " & _
                " Where a.���=b.������ and b.����=" & gintInsure & _
                "        and a.����='" & NVL(!�շ����) & "'"
            
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "���Ѽ���"
            
            If rsTemp.EOF Then
                strTmp = ""
            Else
                strTmp = NVL(rsTemp!����ֵ)
            End If
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '���㱣��
                rs����.Find "id=" & NVL(!����֧������ID, 0), , adSearchForward, 1
                If Not rs����.EOF Then
                    dblͳ����� = NVL(rs����!ͳ��ȶ�, 0) / 100
                Else
                    dblͳ����� = 1
                End If
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,Q��ҵ����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                If gintInsure <> TYPE_���������� And g�������_����.ְ����ҽ��� = "L" _
                    And g�������_����.�α����3 = "0" And NVL(!�Ƿ�ҽ��, 0) = 1 Then  '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '  ������  ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�����ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dblͳ����� = 1
                End If
                
                If gintInsure = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dblͳ����� = dbl��Ŀ����
                End If
                
                If gintInsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                    '�����Q��ҵ����,�������Ϊ100�Է�,�������Ǳ��շ�����
                    If dblͳ����� = 0 Then
                        '�Է�100
                        strTmp = ""
                    Else
                        '�ԷѲ��ַ��� �������Էѷ�����
                    End If
                End If
                                
                '�ܺ�ȫ���� 2003-12-17
                '����������Ŀ��ֻҪ�Ǳ�ʶΪ�����Ρ��ģ���Ӧ�����������
                'If NVL(!�շ����) = "����" And str��Ŀ���� = "����" Then
                If str��Ŀ���� = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If str��Ŀ���� = "���" Then
                    strTmp = "����"
                End If
                '����۳��ԷѲ��ֵķ���
                If Not rsTemp.EOF Then
                    Select Case strTmp
                        Case "����"
                            dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                        Case "��ҩ��"
                            dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                        Case "��ҩ��"
                            dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                        Case "��ҩ��"
                            dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                        Case "����"
                            dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                        Case "���Ʒ�"
                            dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                        '�ܺ�ȫ���� 2003-12-17
                        '������ҽ�������������޷���Ӧ��Ŀ�����������ȡ�õģ�
                        Case "����"
                                 If gintInsure = TYPE_������ Then
                                       '---��˳��
                                       '�����кͿ������Դ����ô���ͬ,
                                       '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                       dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                                       
                                       If g�������_����.ְ����ҽ��� = "Q" Then
                                           '�ԷѲ��ַ��뱣�����Էѷ�����
                                       Else
                                           dbl����Է� = dbl����Է� + Round(NVL(!ʵ�ս��, 0) * (1 - dblͳ�����), 2)
                                       End If
                                 Else
                                       dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0), 2)
                                       dbl����Է� = dbl����Է� + Round(NVL(!ʵ�ս��, 0) * (1 - dblͳ�����), 2)
                                End If
                        Case "�������Ʒ�"
                            '�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                            If gintInsure = TYPE_������ Then
                                dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!ʵ�ս��, 0), 2)
                            Else
                                dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!ʵ�ս��, 0) * dblͳ�����, 2)
                            End If
                        
                            If gintInsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                                '�ԷѲ��ַ��� �������Էѷ�����
                            Else
                                dbl���������Է� = dbl���������Է� + Round(NVL(!ʵ�ս��, 0) * (1 - dblͳ�����), 2)
                            End If
                    End Select
                    If gintInsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                        '�ԷѲ��ַ��� �������Էѷ�����
                        If dblͳ����� <> 0 Then
                            If !�Ƿ�ҽ�� = 1 Then
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!ʵ�ս��, 0) * (1 - dblͳ�����), 2)
                            End If
                        Else
                            '100�ԷѲ��ַ���Ǳ��շ�����
                            dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!ʵ�ս��, 0), 2)
                        End If
                    Else
'                            If InStr(1, "567", NVL(!�շ����, 0)) <> 0 And Len(NVL(!�շ����)) = 1 Then
                                If gintInsure = TYPE_���������� Then
                                    If !�Ƿ�ҽ�� = 1 And dblͳ����� <> 0 Then
                                        '����ҩƷ�Է�  NUM 155 10  ҽ����ҩ�ԷѲ���    Ժ����д
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!ʵ�ս��, 0) * (1 - dblͳ�����), 2)
                                    Else
                                        '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
                                        dbl������ = dbl������ + Round(NVL(!ʵ�ս��, 0) * (1 - dblͳ�����), 2)
                                    End If
                                Else
                                    If strTmp <> "�������Ʒ�" And strTmp <> "����" And !�Ƿ�ҽ�� = 1 And dblͳ����� <> 0 Then
                                        'ҽ����ҩ�Լ����˴�졢��������������Ŀ���ԷѲ���
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!ʵ�ս��, 0) * (1 - dblͳ�����), 2)
                                    End If
                                    
                                    If !�Ƿ�ҽ�� <> 1 Or dblͳ����� = 0 Then
                                        '��ҽ����ҩ�Լ�������Ŀ
                                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!ʵ�ս��, 0), 2)
                                    End If
                                End If
 '                           End If
                        End If
                    End If
            End If
            curTotal = curTotal + Round(NVL(!ʵ�ս��, 0), 2)
            .MoveNext
        Loop
    End With
    
    '��������
'    gstrSQL = "" & _
'        "   Select �������� " & _
'        "   From �ʻ������Ϣ " & _
'        "   where ����=" & gintInsure & " and ����ID=" & lng����id & " and  ���=to_char(sysdate,'yyyy')"
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��������"
    
'    If rsTemp.EOF Then
'        dbl�𸶱�׼ = 0
'    Else
'        dbl�𸶱�׼ = NVL(rsTemp!��������, 0)
'    End If
    If strҽ�� <> "" Then
        gstrSQL = "Select ��� From ��Ա��  where ����='" & strҽ�� & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ�����"
        If Not rsTemp.EOF Then
            strҽ�� = NVL(rsTemp!���)
            If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                strҽ�� = Substr(strҽ��, 1, 6)
            End If
        Else
            strҽ�� = ""
        End If
    End If
    '�ҳ���������
    str��ϱ��� = g�������_����.��ϱ���
    str������� = g�������_����.�������
    With g�������_����
        dbl�𸶱�׼ = .����
        If .ҽ������ = 2 Then   '������
            strInfor = Lpad(gstrҽԺ����_����, 6)       'ҽԺ����
        Else
            strInfor = Lpad(gstrҽԺ����_����, 4)       'ҽԺ����
        End If
        strInfor = strInfor & " "      '�������ʶ
        If gintInsure = TYPE_���������� Then     '������
            strInfor = strInfor & Lpad(.���˱��, 10)       '���˱��
        Else
            strInfor = strInfor & Lpad(.���˱��, 8)      '���˱��
        End If
        strInfor = strInfor & Lpad(.IC����, 7)       'IC����
        strInfor = strInfor & Lpad(.������� + 1, 4)      '�������
        strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '����ʱ��
        strInfor = strInfor & String(10, " ") '��־��
        
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����, 2))), 10) '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10) '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10) '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10)  '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����, 2))), 10)  '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl���Ʒ�, 2))), 10)   '���Ʒ�
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����, 2))), 10)   '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl�������Ʒ�, 2))), 10)   '�������Ʒ�
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����Է�, 2))), 10)   '����Է�
        If gintInsure = TYPE_���������� Then
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl���������Է�, 2))), 10)    '�����Է�    NUM 145 10      Ժ����д
        End If
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl�������Էѷ���, 2))), 10)    '�������Էѷ���
        
        If gintInsure = TYPE_���������� Then       '������
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl������, 2))), 10)    '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
        Else
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl�Ǳ��շ���, 2))), 10)    '�Ǳ��շ���
        End If
        
        strInfor = strInfor & String(10, " ")    '���ķ���:���������ʻ����;������:���������ʻ����  NUM 175 10  ���������ʻ������������ʻ�  ���ķ���
        strInfor = strInfor & String(10, " ")    '���ķ���:�����ͳ��֧���ۼ�  NUM 185 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ���ķ���
            
        If gintInsure = TYPE_���������� Then
                strInfor = strInfor & Lpad(.���������ʻ����, 10)  '����ǰ�����ʻ����  NUM 195 10  �����鿨���ؽ��    Ժ����д
                strInfor = strInfor & Lpad(Trim(CStr(.���������ʻ����)), 10)   '����ǰ�����˻����  NUM 205 10  �����鿨���ؽ��    Ժ����д
                strInfor = strInfor & Lpad(Trim(CStr(.ͳ���ۼ�)), 10)    '����ǰͳ��֧���ۼ�  NUM 215 10  ����ͳ���ۼƣ�����ͳ���ۼƸ����鿨���ؽ�� Ժ����д
        Else
            '����ǰ�����ʻ��������鿨���ؽ�����������������Ӧ����������ѯ�������ʻ������д��
            If Get�������(0, .�������) = "S" Then
                strInfor = strInfor & Lpad(.�����ʻ���ǰֵ, 10)   '����ǰ�����ʻ����
                strInfor = strInfor & Lpad("0", 10)   '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInfor = strInfor & Lpad("0", 10)   '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
            Else
                strInfor = strInfor & Lpad(.���������ʻ����, 10)  '����ǰ�����ʻ����
                strInfor = strInfor & Lpad(Trim(CStr(.���������ʻ����)), 10)   '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInfor = strInfor & Lpad(Trim(CStr(.ͳ���ۼ�)), 10)    '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
            End If
        End If
        
        strInfor = strInfor & String(10, " ")    '���ķ���:���λ��������ʻ�֧��(������������㣬��ʾ�����ʻ�֧��)
        strInfor = strInfor & String(10, " ")    '���ķ���:���β��������ʻ�֧��(������������㷵��0)
        strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ��֧��
        strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ���Ը�
        strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ��֧��
        strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ���Ը�
        strInfor = strInfor & String(10, " ")    '���ķ���:���λ�����������֧�� ��������:����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
        strInfor = strInfor & String(10, " ")    '���ķ���:���ηǻ�����������֧����������:����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
        strInfor = strInfor & String(10, " ")    '���ķ���:���α��շ�Χ���Ը���������:�޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
        
        If gintInsure <> TYPE_���������� Then
            strInfor = strInfor & Lpad(Trim(CStr(dbl���������Է�)), 10)    '�������������Ը�
        End If
        
        strInfor = strInfor & Lpad(Trim(CStr(dbl�𸶱�׼)), 10)    '�𸶱�׼��������:����סԺ�ż���  NUM 315 10      Ժ����д
        strInfor = strInfor & Lpad(.ת�ﵥ��, 6)     'ת�ﵥ��
        strInfor = strInfor & Lpad(Get�������(0, .�������), 1)     '�������
        If gintInsure <> TYPE_���������� Then
            
            strInfor = strInfor & Lpad(.�α����3, 1)    '�α����3:0 �󱣡�1 �±��������鿨���
        End If
        strInfor = strInfor & Lpad(.ְ����ҽ���, 1)       'ְ����ҽ���
        
        strInfor = strInfor & Lpad(.��ϱ���, 16)    '��ϱ���
        
        strInfor = strInfor & Lpad(strҽ��, 6)    'ҽʦ����
        strInfor = strInfor & Lpad(UserInfo.���, 6)    '����Ա����
        strInfor = strInfor & Lpad(.�������, 30)    '�������
        'A-������B-��ת��C-δ����D-������E-����
        strInfor = strInfor & "A"    '���������ʶ
        strInfor = strInfor & String(8, " ")      '��Ժ����
        
        If gintInsure = TYPE_���������� Then       '������
        Else
            strInfor = strInfor & String(16, " ")      '����ʱ��
        End If
        strInfor = strInfor & String(10, " ")      '�������
    End With
    
    '��������ӿ�(1006    12  423   ʵʱ����Ԥ��
    �����������_���� = ҵ������_����(IIf(gintInsure = TYPE_����������, 2, 1), 1006, strInfor)
    If �����������_���� = False Then
        Exit Function
    End If
    
    '������:
    '    ���λ��������ʻ�֧��    NUM 225 10      ���ķ���
    '    ���β��������ʻ�֧��    NUM 235 10      ���ķ���
    '    ���λ���ͳ��֧��    NUM 245 10      ���ķ���
    '    ���λ���ͳ���Ը�    NUM 255 10      ���ķ���
    '    ���β���ͳ��֧��    NUM 265 10      ���ķ���
    '    ���β���ͳ���Ը�    NUM 275 10      ���ķ���
    '    ���λ�����������֧��    NUM 285 10  ����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
    '    ���ηǻ�����������֧��  NUM 295 10  ����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
    '    ���α��շ�Χ���Ը�  NUM 305 10  �޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
    '������:
    '    ���λ��������ʻ�֧��    NUM 211 10  ������������㣬��ʾ�����ʻ�֧��    ����
    '    ���β��������ʻ�֧��    NUM 221 10  ������������㷵��0 ����
    '    ���λ���ͳ��֧��    NUM 231 10      ����
    '    ���λ���ͳ���Ը�    NUM 241 10      ����
    '    ���β���ͳ��֧��    NUM 251 10  ������������㣬���ֶ����ڴ����������֧��  ����
    '    ���β���ͳ���Ը�    NUM 261 10      ����
    '    ���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧�� 2�� ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
    '    ���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��   2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����    ����
    '    ���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    
    Dim i As Long
    If gintInsure = TYPE_���������� Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    'ȷ�����ν��㷽ʽ
    str���㷽ʽ = "�����ʻ�;" & Format(Val(Substr(strInfor, i + 10, 10)), "###0.00;-###0.00;0;0") & ";0" '���λ��������ʻ�֧��,�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "�����ʻ�;" & Format(Val(Substr(strInfor, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strInfor, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strInfor, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "��������;" & Format(Val(Substr(strInfor, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "�ǲ�������;" & Format(Val(Substr(strInfor, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    
    �����������_���� = True
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    Dim lng����id As Long
    �������_���� = Set�����������(False, lng����ID, cur�����ʻ�, lng����id, strSelfNo)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    �������_���� = False
End Function
Private Function Set�����������(ByVal bln���� As Boolean, lng����ID As Long, cur�����ʻ� As Currency, lng����id As Long, strSelfNo As String) As Boolean
  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID��
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim curTotal As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl���� As Double
    Dim dbl���Ʒ� As Double
    Dim dbl���� As Double
    Dim dbl����Է� As Double
    Dim dbl�������Ʒ� As Double
    Dim dbl���������Է� As Double
    Dim dbl�������Էѷ��� As Double
    Dim dbl�Ǳ��շ��� As Double
    Dim dblͳ����� As Double
    Dim dbl������ As Double     '��Դ�����������
    Dim dbl�𸶱�׼ As Double
    Dim dbl���� As Double
    Dim strҽ�� As String
    Dim str��ϸ As String       '��ϸ��
    Dim str���ұ��� As String
    Dim str��Ŀ���� As String
    Dim str��Ŀͳ�Ʒ��� As String
    Dim strTmp As String
    Dim intҵ�� As Integer
    Dim lng����ID As Long
    Dim strNO As String
    Dim lng��¼���� As Long
    
    Dim dbl�����ʻ���� As Double
    Dim dblͳ��֧���ۼ� As Double
    Dim dbl�����ʻ�֧�� As Double
    Dim dbl�����ʻ�֧�� As Double
    Dim dbl����ͳ��֧�� As Double
    Dim dbl����ͳ���Ը� As Double
    Dim dbl����ͳ��֧�� As Double
    Dim dbl����ͳ���Ը� As Double
    Dim dbl��������֧�� As Double
    Dim dbl�ǲ�������֧�� As Double
    Dim dbl���շ�Χ���Ը� As Double
    
    Dim dbl����ǰ�����ʻ����  As Double
    Dim dbl����ǰ�����˻����  As Double
    Dim dbl����ǰͳ���ۼ�  As Double
    Dim lngTmp As Long
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim lng����ID As Long
    
    intҵ�� = IIf(bln����, 1, 0)
     Set����������� = False
   
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ���������ڽ��㣬����ʹ��Ȧ��ӿڣ����������Ǯ���������ڣ������ӿ��ڽ��
    '���������Ҫͨ��������������ȡ����Ȧ�����ǽӿڷ��أ���Ҫ�޸�
    
    On Error GoTo ErrHand
    '���¶���
    If ��ȡ�������_����(IIf(gintInsure = TYPE_����������, 2, 1)) = False Then
        Exit Function
    End If
    If bln���� Then
        '��֤�Ƿ�Ϊ�ò��˵�IC��
        gstrSQL = "Select * From  �����ʻ� where ����id=" & lng����id
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵�ҽ����"
        If rsTemp.EOF Then
            ShowMsgbox "�ò����ڱ����ʻ����޼�¼!"
            Exit Function
        End If
        
        If g�������_����.IC���� <> NVL(rsTemp!����) Then
            ShowMsgbox "�ò��˵�IC���������,�����ǲ����������˵�IC��!"
            Exit Function
        End If
        'ȷ���������,ת�ﵥ��,��ϱ���,�������
        ' ֧��˳���_IN(�������;ת�ﵥ��;��ϱ���),��ע(�������_IN)
        gstrSQL = "Select ֧��˳���,��ע from ���ս����¼  where ��¼ID=" & lng����ID
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�������"
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "�ڽ����¼���޽����¼!"
            Exit Function
        End If
        Dim strArr
        strArr = Split(NVL(rsTemp!֧��˳���), ";")
        
        '�������;ת�ﵥ��;��ϱ���
        '1-��ͨ����("1", "A"),2-��������("3", "7")
        '3-�����("5", "B"),4-������������("S", "T")
        If UBound(strArr) >= 2 Then
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g�������_����.ת�ﵥ�� = strArr(1)
            g�������_����.��ϱ��� = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g�������_����.ת�ﵥ�� = strArr(1)
        Else
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
        End If
        g�������_����.������� = NVL(rsTemp!��ע)
        
        
        'ȷ���˷Ѽ�¼
        '�˷�
          gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B " & _
                    " where A.NO=B.NO and A.��¼����=B.��¼����  and A.��¼״̬=2 and B.����ID=" & lng����ID
          Call OpenRecordset(rsTemp, "�����˷�")
          If rsTemp.EOF Then
            ShowMsgbox "�����ڲ��˷��ó�����¼!"
            Exit Function
          Else
            lng����ID = rsTemp("����ID")
          End If
          
    End If
    '�򿪱��ν�����ϸ��¼
    gstrSQL = " " & _
        "  Select Rownum ��ʶ��,A.ID,A.����ID,A.�շ�ϸĿid,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������ as ҽ��,H.��� as ҽ�����, " & _
        "      A.����*A.���� as ����,A.���㵥λ,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ�� as ʵ�ս��,F.����ֵ,G.id as ����id,G.ͳ��ȶ�, " & _
        "      A.�շ����,B.���� as ��Ŀ����,B.���� as ��Ŀ����,B.��ʶ����||B.��ʶ���� as ���ұ���, " & _
        "      D.��Ŀ���� ҽ������,D.��Ŀ���� as ҽ������,J.���� as ����,D.�Ƿ�ҽ��,C.���� ��������,E.���� �ܵ�����, " & _
        "      L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.��λ����,L.˳���,L.����֤��,L.�ʻ����,L.��ǰ״̬,L.����ID,L.��ְ,L.�����,L.�Ҷȼ�,L.����ʱ�� " & _
        "  From (Select * From ���˷��ü�¼ Where ����ID=" & IIf(bln����, lng����ID, lng����ID) & " and  Nvl(���ӱ�־,0)<>9 ) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E,  " & _
        "       (Select U.*,K.����ֵ From �շ���� U,���ղ��� K where U.���=K.������ and K.����=" & gintInsure & "  ) F, " & _
        "       (Select distinct Q.ҩƷid,T.���� From ҩƷĿ¼ Q,ҩƷ��Ϣ R,ҩƷ���� T  Where  Q.ҩ��id=R.ҩ��id and R.����=T.���� ) J, " & _
        "       ����֧������ G,��Ա�� H,�����ʻ� L" & _
        "  Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+) and A.����id=L.����id and L.����=" & gintInsure & " and A.�շ����=F.����(+)  and d.����id=G.id and a.�շ�ϸĿid=J.ҩƷid(+) " & _
        "        And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����= " & gintInsure & " and a.������=H.����(+) " & _
        "  Order by A.ID"
        
    '�ϴ�������ϸ��¼
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ���ν��ʷ�����ϸ"
    
    With rs��ϸ
        If Not .EOF Then
            lng����id = NVL(!����ID, 0)
            strҽ�� = NVL(!ҽ�����)
            If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                strҽ�� = Substr(strҽ��, 1, 6)
            End If
            lng����ID = NVL(!����ID, 0)
            '����׼��Ŀ
            gstrSQL = "Select * from ������׼��Ŀ  where ����ID=  " & lng����ID
            zlDatabase.OpenRecordset rs��׼��Ŀ, gstrSQL, "��ȡ������Ŀ����"
        End If
        Do While Not .EOF
            If lng����ID <> 0 Then
                    '��һ��,ȷ��������շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=1 and �շ�ϸĿid=" & NVL(!�շ�ϸĿID, 0)
                    If rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�շ�ϸĿΪ��" & NVL(!��Ŀ����) & "������Ŀ���ǲ��������趨����Ŀ."
                        Exit Function
                    End If
                    '�ڶ���,ȷ������ı��մ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=1 and  �շ�ϸĿid=" & NVL(!����ID, 0)
                    If rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�ڽ����д����˽�������ı���֧������,���ܼ�����"
                        Exit Function
                    End If
                    '������,'ȷ����ֹ���շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=2 and �շ�ϸĿid=" & NVL(!�շ�ϸĿID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�շ�ϸĿΪ��" & NVL(!��Ŀ����) & "������Ŀ�Ǳ���ֹʹ�õ���Ŀ." & vbCrLf & "���ܼ���!"
                        Exit Function
                    End If
                    '���Ĳ�,'ȷ����ֹ�Ĵ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=2 and �շ�ϸĿid=" & NVL(!����ID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�ڽ����д����˽�ֹʹ�õı���֧������,���ܼ�����"
                    End If
            End If
            strTmp = NVL(!����ֵ)
            lng����id = NVL(!����ID, 0)
            'ȷ���������
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str��Ŀͳ�Ʒ��� = ""
                Else
                    str��Ŀͳ�Ʒ��� = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '����
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                    
                If NVL(!����, 0) <> TYPE_���������� And Val(NVL(!��λ����, "99")) = 0 And NVL(!��ְ, 0) = 3 And NVL(!�Ƿ�ҽ��, 0) = 1 Then   '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '������    ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dbl���� = 1
                Else
                    dbl���� = NVL(!ͳ��ȶ�, 0) / 100
                End If
                
                If NVL(!����, 0) = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dbl���� = Val(NVL(!ҽ������))
                End If
                
                If NVL(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                    '�����Q��ҵ����,�������Ϊ100�Է�,�������Ǳ��շ�����
                    If dblͳ����� = 0 Then
                        '�Է�100
                        strTmp = ""
                    Else
                        '�ԷѲ��ַ��� �������Էѷ�����
                    End If
                End If
                
                If NVL(!ҽ������) = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If NVL(!ҽ������) = "���" Then
                    strTmp = "����"
                End If
                
                Select Case strTmp
                    Case "����"
                            dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                    Case "��ҩ��"
                           dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                    Case "��ҩ��"
                            dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                    Case "��ҩ��"
                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                    Case "����"
                        dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                    Case "���Ʒ�"
                        dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                    Case "����"
                          If gintInsure = TYPE_������ Then
                                '---��˳��
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                                
                                If g�������_����.ְ����ҽ��� = "Q" Then
                                    '�ԷѲ��ַ��뱣�����Էѷ�����
                                Else
                                    dbl����Է� = dbl����Է� + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
                                End If
                          Else
                                dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0), 2)
                                dbl����Է� = dbl����Է� + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
                         End If
'
'                        If gintInsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
'                            '�ԷѲ��ַ��� �������Էѷ�����
'                        Else
'                            dbl����Է� = dbl����Է� + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
'                        End If
                    Case "�������Ʒ�"
                        '�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                        If gintInsure = TYPE_������ Then
                            dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!ʵ�ս��, 0), 2)
                        Else
                            dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                        End If
                        If gintInsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                            '�ԷѲ��ַ��� �������Էѷ�����
                        Else
                            dbl���������Է� = dbl���������Է� + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
                        End If
                End Select
                If gintInsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                    '�ԷѲ��ַ��� �������Էѷ�����
                    If dblͳ����� <> 0 Then
                        If !�Ƿ�ҽ�� = 1 Then
                            dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
                        End If
                    Else
                        '100�ԷѲ��ַ���Ǳ��շ�����
                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!ʵ�ս��, 0), 2)
                    End If
                Else
'                        If InStr(1, "567", NVL(!�շ����, 0)) <> 0 And Len(NVL(!�շ����)) = 1 Then
                        If gintInsure = TYPE_���������� Then
                            If !�Ƿ�ҽ�� = 1 And dbl���� <> 0 Then
                                '����ҩƷ�Է�  NUM 155 10  ҽ����ҩ�ԷѲ���    Ժ����д
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
                            Else
                                '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
                                dbl������ = dbl������ + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
                            End If
                        Else
                            If strTmp <> "�������Ʒ�" And strTmp <> "����" And !�Ƿ�ҽ�� = 1 And dbl���� <> 0 Then
                                'ҽ����ҩ�Լ����˴�졢��������������Ŀ���ԷѲ���
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!ʵ�ս��, 0) * (1 - dbl����), 2)
                            End If
                                            
                            'Ҫ����100%���ԷѲ�������Ǳ��շ�����
                            If !�Ƿ�ҽ�� <> 1 Or dbl���� = 0 Then
                                '��ҽ����ҩ�Լ�������Ŀ
                                dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!ʵ�ս��, 0), 2)
                            End If
                          End If
    '                    End If
               End If
            Else
                dbl���� = 1
                str��Ŀͳ�Ʒ��� = ""
            End If

            '�ϴ���ϸ��¼,ʵʱҽ����ϸ����
            '����������ϸ�ϴ�
            If gbln������ϸʱʵ�ϴ� Then
                
                    If NVL(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = Lpad(gstrҽԺ����_����, 6)     'ҽԺ����    CHAR    1   6       Ժ����д
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ����), 10)  '���ձ��    CHAR    7   10      Ժ����д
                    Else
                        str��ϸ = Lpad(gstrҽԺ����_����, 4)     'ҽԺ����    CHAR    1   4       Ժ��
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ����), 8)   '���˱��    CHAR    5   8       Ժ��
                    End If
                
                    str��ϸ = str��ϸ & Space(10)   '��־��  CHAR    13  10  ������ϸ�Կո�λ,סԺ��סԺ��  Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!˳���, 0), 4)   '�������    NUM 23  4   סԺ��ϸ�����������Ժ�Ǽ�ʱ�������������ϸ:                         ������ڱ��ν���������� Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!NO, 0), 10)       '������  NUM 27  10      Ժ��
                    
                    If NVL(!����, 0) = TYPE_���������� Then '������
                    Else
                        str��ϸ = str��ϸ & Lpad(CStr(.AbsolutePosition), 10)       '������Ŀ���    NUM 37  10  ��Ӧ�����ŵļǼ���Ŀ���    Ժ��
                    End If
                    '������Ϊ���ݺ�  CHAR    41  10  ҽ���ţ�    Ժ����д
                    str��ϸ = str��ϸ & Space(10)       'ҽ����  CHAR    47  10  ������Ӧҽ����ҽ����¼�ţ�������ϸ��û��ҽ����ҽԺ�Կո�λ    Ժ��
                    
                    str��ϸ = str��ϸ & Get�������(intҵ��, NVL(!�Ҷȼ�, 0))         '�������    CHAR    57  1   ȡֵ���"�������"˵��  Ժ��
                    
                    If NVL(!����, 0) = TYPE_���������� Then '������
                        '������Ϊ����ʱ��    DATETIME    52  16  ��ȷ���루������ʱ�䣩��ʽΪ��yyyymmddhhmiss�����Կո�λ  Ժ����д
                        str��ϸ = str��ϸ & Rpad(Format(!����ʱ��, "yyyymmddHHmmss"), 16)
                    Else
                        str��ϸ = str��ϸ & Rpad(Format(!�Ǽ�ʱ��, "yyyymmddHHmmss"), 16)      '��������ʱ�䣨Ͷҩʱ�䣩    DATETIME    58  16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ    Ժ��
                    End If
                    
                    str��ϸ = str��ϸ & Lpad(NVL(!���ұ���), 20)      '��Ŀ����    CHAR    74  20  �Ƽ���Ŀ����    Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!��Ŀ����), 20)      '��Ŀ����    CHAR    94  20      Ժ��
        
                    If NVL(!����, 0) = TYPE_���������� Then '������
                    Else
        
                        If !�Ƿ�ҽ�� = 1 Then
                            str��ϸ = str��ϸ & Lpad(1 - dbl����, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                        Else
                            str��ϸ = str��ϸ & Lpad(1, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                        End If
                        str��ϸ = str��ϸ & Lpad(str��Ŀͳ�Ʒ���, 1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(NVL(!����), 6)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!ʵ�ʼ۸�), 8) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!���㵥λ), 4) '��λ    CHAR    135 4       Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!����), 20)      '����    CHAR    139 20  �����Ƭ����    Ժ��
                    
                    If NVL(!����, 0) = TYPE_���������� Then '������
                        '��ȡ���˵�����.
                        gstrSQL = "Select ����,Ƶ��,�÷� From ҩƷ�շ���¼ where ����id=" & NVL(!ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵���Ƶ��"
                        If rsTemp.EOF Then
                            str��ϸ = str��ϸ & Space(5)       'ÿ������    NUM 146 5       Ժ����д
                            str��ϸ = str��ϸ & Space(20)      'ʹ��Ƶ��    CHAR    151 20  �磺1��2��  Ժ����д
                            str��ϸ = str��ϸ & Space(50)      '�÷�    CHAR    171 50  �磺�ڷ�    Ժ����д
                        Else
                            str��ϸ = str��ϸ & Lpad(NVL(rsTemp!����), 5)      'ÿ������    NUM 146 5       Ժ����д
                            str��ϸ = str��ϸ & Lpad(NVL(rsTemp!Ƶ��), 20)      'ʹ��Ƶ��    CHAR    151 20  �磺1��2��  Ժ����д
                            str��ϸ = str��ϸ & Lpad(NVL(rsTemp!�÷�), 50)      '�÷�    CHAR    171 50  �磺�ڷ�    Ժ����д
                        End If
                        str��ϸ = str��ϸ & Space(4)      'ִ������    NUM 221 4       Ժ����д
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ�����), 6)      'ҽʦ����    CHAR    225 6       Ժ����д
                    Else
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ��), 8)      'ҽʦ����    CHAR    159 8       Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(g�������_����.��ϱ���, 16)      '��ϱ���    CHAR    167 16      Ժ��
                    str��ϸ = str��ϸ & Lpad(g�������_����.�������, 30)     '�������    CHAR    183 30      Ժ��
                    If NVL(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = str��ϸ & Lpad(NVL(!��������), 20)    '�Ʊ�����    CHAR    277 20      Ժ����д
                    Else
                        str��ϸ = str��ϸ & Space(16)     '����ʱ��    DATETIME    213 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ��Ժ�˿ո�λ  ����
                    End If
                    
                
                '�ϴ���ϸ
                '1003    7   230 ʵʱҽ����ϸ�����ύ
                Set����������� = ҵ������_����(IIf(NVL(!����, 0) = TYPE_����������, 2, 1), 1003, str��ϸ)
                If Set����������� = False Then
                    ShowMsgbox "�������ʱҽ����ϸ�����ύʧ��,���ܼ���!"
                    Exit Function
                End If
                'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & NVL(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            End If
            '�����ܶ�,����
            curTotal = curTotal + Round(NVL(!ʵ�ս��, 0), 2)
            .MoveNext
        Loop
    End With
    Set����������� = False
    '����ʱ,���»�ȡ���˵������Ϣ.
'    If bln���� Then
'        Call ��ȡ������Ϣ_����(lng����id)
'    End If
    '��������
        dbl�𸶱�׼ = g�������_����.����

    If bln���� Then
       '��ȷ���ϴ����ķ��ص�����

        gstrSQL = "" & _
            "   Select *  " & _
            "   From ���ս����¼ " & _
            "   Where ��¼id=" & lng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�����շ�ʱ���ص�����"
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "�������ϴ��շѵĽ����¼!"
            Exit Function
        End If
        '/???
        'ԭ���̲���:
        '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
        "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
        '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
        '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
        '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
        '������ֵ����Ϊ:
        '       ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN, _
        '       dbl�����ʻ����,dblͳ��֧���ۼ�,dbl��������֧��,dbl�����ʻ�֧��,סԺ����_IN,����_IN,dbl���շ�Χ���Ը�,ʵ������_IN
        '       �������ý��_IN,dbl����ͳ��֧��,dbl����ͳ���Ը�,
        '       dbl����ͳ��֧��,dbl����ͳ���Ը�,dbl�ǲ�������֧��,�����Ը����_IN,dbl�����ʻ�֧��
        '       ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
           dbl�����ʻ���� = Round(NVL(rsTemp!�ʻ��ۼ�����, 0), 2)
           dblͳ��֧���ۼ� = Round(NVL(rsTemp!�ʻ��ۼ�֧��, 0), 2)
           dbl��������֧�� = Round(NVL(rsTemp!�ۼƽ���ͳ��, 0), 2)
           dbl�����ʻ�֧�� = Round(NVL(rsTemp!�ۼ�ͳ�ﱨ��, 0), 2)
           dbl�𸶱�׼ = Round(NVL(rsTemp!����, 0), 2)
           dbl���շ�Χ���Ը� = Round(NVL(rsTemp!�ⶥ��, 0), 2)
           dbl����ͳ��֧�� = Round(NVL(rsTemp!ȫ�Ը����, 0), 2)
           dbl����ͳ���Ը� = Round(NVL(rsTemp!�����Ը����, 0), 2)
           dbl����ͳ��֧�� = Round(NVL(rsTemp!����ͳ����, 0), 2)
           dbl����ͳ���Ը� = Round(NVL(rsTemp!ͳ�ﱨ�����, 0), 2)
           dbl�ǲ�������֧�� = Round(NVL(rsTemp!���Ը����, 0), 2)
           dbl�����ʻ�֧�� = Round(NVL(rsTemp!�����ʻ�֧��, 0), 2)
           dbl����ǰ�����ʻ���� = Round(NVL(rsTemp!����ǰ�����ʻ����, 0), 2)
           dbl����ǰ�����˻���� = Round(NVL(rsTemp!����ǰ�����˻����, 0), 2)
           dbl����ǰͳ���ۼ� = Round(NVL(rsTemp!����ǰͳ���ۼ�, 0), 2)
    End If
    '���ҽ������
    
    '�ҳ���������
    With g�������_����
        If gintInsure = TYPE_���������� Then    '������
            strInfor = Lpad(gstrҽԺ����_����, 6)       'ҽԺ����
        Else
            strInfor = Lpad(gstrҽԺ����_����, 4)       'ҽԺ����
        End If
        strInfor = strInfor & " "      '�������ʶ
        If gintInsure = TYPE_���������� Then   '������
            strInfor = strInfor & Lpad(.���˱��, 10)         '���˱��
        Else
            strInfor = strInfor & Lpad(.���˱��, 8)      '���˱��
        End If
        strInfor = strInfor & Lpad(.IC����, 7)       'IC����
        .������� = .������� + 1
        strInfor = strInfor & Lpad(.�������, 4)       '�������
        strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '����ʱ��
        strInfor = strInfor & String(10, " ") '��־��
        
        '�ܺ�ȫ���� 2003-12-17
        '���ڲ����ǽ��㻹�ǳ�����������Ϊ���������Դ˴�ֻ��ȡ����ֵ
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����), 2))), 10) '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl��ҩ��), 2))), 10) '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl��ҩ��), 2))), 10) '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl��ҩ��), 2))), 10)  '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����), 2))), 10)  '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl���Ʒ�), 2))), 10)   '���Ʒ�
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����), 2))), 10)    '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl�������Ʒ�), 2))), 10)   '�������Ʒ�
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����Է�), 2))), 10)   '����Է�
        If gintInsure = TYPE_���������� Then
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl���������Է�), 2))), 10)    '�����Է�    NUM 145 10      Ժ����д
        End If
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl�������Էѷ���), 2))), 10)    '�������Էѷ���
        
        If gintInsure = TYPE_���������� Then       '������
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl������), 2))), 10)    '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
        Else
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl�Ǳ��շ���), 2))), 10)    '�Ǳ��շ���
        End If
      '�ܺ�ȫ���� 2003-12-22
        '�˴�����ǳ���Ӧ��ͬʱ��ȡ�ϴν��������д
        
'        strInfor = strInfor & String(10, " ")    '���ķ���:���������ʻ����;������:���������ʻ����  NUM 175 10  ���������ʻ������������ʻ�  ���ķ���
'        strInfor = strInfor & String(10, " ")    '���ķ���:�����ͳ��֧���ۼ�  NUM 185 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ���ķ���
        strInfor = strInfor & Lpad(dbl�����ʻ����, 10)
        strInfor = strInfor & Lpad(dblͳ��֧���ۼ�, 10)
        
        '����ǰ�����ʻ��������鿨���ؽ�����������������Ӧ����������ѯ�������ʻ������д��
        Dim dbl����ǰ���(1 To 3) As Double '1-����ǰ�����ʻ����,2-����ǰ�����˻����,3-����ǰͳ��֧���ۼ�
        
        dbl����ǰ���(1) = .���������ʻ����
        dbl����ǰ���(2) = .���������ʻ����
        dbl����ǰ���(3) = .ͳ���ۼ�
        
        '����ǰ�����ʻ��������鿨���ؽ�����������������Ӧ����������ѯ�������ʻ������д��
        If bln���� Then
                strInfor = strInfor & Lpad(dbl����ǰ�����ʻ����, 10)   '����ǰ�����ʻ����
                strInfor = strInfor & Lpad(dbl����ǰ�����˻����, 10)    '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInfor = strInfor & Lpad(dbl����ǰͳ���ۼ�, 10)     '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
        Else
            If gintInsure <> TYPE_���������� And Get�������(0, .�������) = "S" Then
                strInfor = strInfor & Lpad(.�����ʻ���ǰֵ, 10)   '����ǰ�����ʻ����
                strInfor = strInfor & Lpad("0", 10)   '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInfor = strInfor & Lpad("0", 10)   '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
                dbl����ǰ���(1) = .�����ʻ���ǰֵ
                dbl����ǰ���(2) = 0
                dbl����ǰ���(3) = 0
            Else
                strInfor = strInfor & Lpad(.���������ʻ����, 10)  '����ǰ�����ʻ����
                strInfor = strInfor & Lpad(Trim(CStr(.���������ʻ����)), 10)   '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInfor = strInfor & Lpad(Trim(CStr(.ͳ���ۼ�)), 10)    '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
            End If
        End If
        
        If bln���� Then
            'dbl�����ʻ���� = Round(NVL(rsTemp!�ʻ��ۼ�����, 0), 2)
            'dblͳ��֧���ۼ� = Round(NVL(rsTemp!�ʻ��ۼ�֧��, 0), 2)
            'dbl�𸶱�׼ = Round(NVL(rsTemp!����, 0), 2)
            
            strInfor = strInfor & Lpad(dbl�����ʻ�֧��, 10) ' = Round(NVL(rsTemp!�����ʻ�֧��, 0), 2)
            strInfor = strInfor & Lpad(dbl�����ʻ�֧��, 10) ' = Round(NVL(rsTemp!�ۼ�ͳ�ﱨ��, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ��֧��, 10) ' = Round(NVL(rsTemp!ȫ�Ը����, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ���Ը�, 10) ' = Round(NVL(rsTemp!�����Ը����, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ��֧��, 10) ' = Round(NVL(rsTemp!����ͳ����, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ���Ը�, 10) ' = Round(NVL(rsTemp!ͳ�ﱨ�����, 0), 2)
            strInfor = strInfor & Lpad(dbl��������֧��, 10) ' = Round(NVL(rsTemp!�ۼƽ���ͳ��, 0), 2)
            strInfor = strInfor & Lpad(dbl�ǲ�������֧��, 10) ' = Round(NVL(rsTemp!���Ը����, 0), 2)
            strInfor = strInfor & Lpad(dbl���շ�Χ���Ը�, 10) ' = Round(NVL(rsTemp!�ⶥ��, 0), 2)
        Else
        
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ��������ʻ�֧��(������������㣬��ʾ�����ʻ�֧��)
            strInfor = strInfor & String(10, " ")    '���ķ���:���β��������ʻ�֧��(������������㷵��0)
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ��֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ���Ը�
            strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ��֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ���Ը�
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ�����������֧�� ��������:����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
            strInfor = strInfor & String(10, " ")    '���ķ���:���ηǻ�����������֧����������:����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
            strInfor = strInfor & String(10, " ")    '���ķ���:���α��շ�Χ���Ը���������:�޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
        End If
        If gintInsure <> TYPE_���������� Then
            strInfor = strInfor & Lpad(Trim(CStr(dbl���������Է�)), 10)    '�������������Ը�
        End If
        
       
        '�ܺ�ȫ���� 2003-12-22
        '��������Ҫ�𸶱�׼��Ӧ�ô�0
        '        strInfor = strInfor & Lpad(Trim(CStr(dbl�𸶱�׼)), 10)    '�𸶱�׼
        strInfor = strInfor & Lpad(0, 10)
        strInfor = strInfor & Lpad(.ת�ﵥ��, 6)     'ת�ﵥ��
        strInfor = strInfor & Lpad(Get�������(intҵ��, .�������), 1)     '�������
        
        If gintInsure <> TYPE_���������� Then
            strInfor = strInfor & Lpad(.�α����3, 1)    '�α����3:0 �󱣡�1 �±��������鿨���
        End If
        strInfor = strInfor & Lpad(.ְ����ҽ���, 1)       'ְ����ҽ���
        
        strInfor = strInfor & Lpad(.��ϱ���, 16)    '��ϱ���
        strInfor = strInfor & Lpad(strҽ��, 6)    'ҽʦ����
        strInfor = strInfor & Lpad(UserInfo.���, 6)    '����Ա����
        strInfor = strInfor & Lpad(.�������, 30)    '�������
        'A-������B-��ת��C-δ����D-������E-����
        strInfor = strInfor & "A"    '���������ʶ
        strInfor = strInfor & String(8, " ")      '��Ժ����
        
        If gintInsure = TYPE_���������� Then       '������
        Else
            strInfor = strInfor & String(16, " ")      '����ʱ��
        End If
        strInfor = strInfor & String(10, " ")      '�������
    End With
    
    '����1002    12  423 ʵʱ����
    Set����������� = ҵ������_����(IIf(gintInsure = TYPE_����������, 2, 1), 1002, strInfor)
    If Set����������� = False Then
        Exit Function
    End If
    
    '
   
    
   
    
    '������:
    '   ���������ʻ����  NUM 175 10  ���������ʻ������������ʻ�  ���ķ���
    '   �����ͳ��֧���ۼ�  NUM 185 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ���ķ���
    
    '    ���λ��������ʻ�֧��    NUM 225 10      ���ķ���
    '    ���β��������ʻ�֧��    NUM 235 10      ���ķ���
    '    ���λ���ͳ��֧��    NUM 245 10      ���ķ���
    '    ���λ���ͳ���Ը�    NUM 255 10      ���ķ���
    '    ���β���ͳ��֧��    NUM 265 10      ���ķ���
    '    ���β���ͳ���Ը�    NUM 275 10      ���ķ���
    '    ���λ�����������֧��    NUM 285 10  ����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
    '    ���ηǻ�����������֧��  NUM 295 10  ����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
    '    ���α��շ�Χ���Ը�  NUM 305 10  �޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
    '������:
    '   ���������ʻ����  NUM 161 10  ��  ����ǻ���ҽ�ƽ����ʾ�����������ʻ������������ʻ��� ��������������ʾ: �����ʻ��������� ����
    '   �����ͳ��֧���ۼ�  NUM 171 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ����
    
    '    ���λ��������ʻ�֧��    NUM 211 10  ������������㣬��ʾ�����ʻ�֧��    ����
    '    ���β��������ʻ�֧��    NUM 221 10  ������������㷵��0 ����
    '    ���λ���ͳ��֧��    NUM 231 10      ����
    '    ���λ���ͳ���Ը�    NUM 241 10      ����
    '    ���β���ͳ��֧��    NUM 251 10  ������������㣬���ֶ����ڴ����������֧��  ����
    '    ���β���ͳ���Ը�    NUM 261 10      ����
    '    ���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧�� 2�� ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
    '    ���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��   2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����    ����
    '    ���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    
    Dim i As Long
    If gintInsure = TYPE_���������� Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    
    dbl�����ʻ���� = Val(Substr(strInfor, i - 40, 10))
    dblͳ��֧���ۼ� = Val(Substr(strInfor, i - 30, 10))  '�����ͳ��֧���ۼ�=����ͳ���ۼƣ�����ͳ���ۼ�
    
    dbl�����ʻ�֧�� = Val(Substr(strInfor, i + 10, 10)) '���λ��������ʻ�֧��=������������㣬��ʾ�����ʻ�֧��
    dbl�����ʻ�֧�� = Val(Substr(strInfor, i + 20, 10))    '���β��������ʻ�֧��    NUM 221 10  ������������㷵��0
    dbl����ͳ��֧�� = Val(Substr(strInfor, i + 30, 10))   '���λ���ͳ��֧��    NUM 231 10      ����
    dbl����ͳ���Ը� = Val(Substr(strInfor, i + 40, 10))     '���λ���ͳ���Ը�    NUM 241 10      ����
    dbl����ͳ��֧�� = Val(Substr(strInfor, i + 50, 10))     '���β���ͳ��֧��    NUM 251 10      ����
    dbl����ͳ���Ը� = Val(Substr(strInfor, i + 60, 10))     '���β���ͳ���Ը�    NUM 261 10      ����
    dbl��������֧�� = Val(Substr(strInfor, i + 70, 10))     '���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧��2��   ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
    dbl�ǲ�������֧�� = Val(Substr(strInfor, i + 80, 10))     '���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����
    dbl���շ�Χ���Ը� = Val(Substr(strInfor, i + 90, 10))     '���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    
    '/???
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    '    ����_IN,��ҩ��_IN,��ҩ��_IN,��ҩ��_IN,����_IN,���Ʒ�_IN,����_IN,����Է�_IN,�������Ʒ�_IN,���������Է�_IN,�������Էѷ���_IN,�Ǳ��շ���_IN,ͳ�����_IN,������
    '   ����ǰ�����ʻ����,����ǰ�����˻����,����ǰͳ���ۼ�
    '������ֵ����Ϊ:
    '       ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN, _
    '       dbl�����ʻ����,dblͳ��֧���ۼ�,dbl��������֧��,dbl�����ʻ�֧��,סԺ����_IN,����_IN,dbl���շ�Χ���Ը�,ʵ������_IN
    '       �������ý��_IN,dbl����ͳ��֧��,dbl����ͳ���Ը�,
    '       dbl����ͳ��֧��,dbl����ͳ���Ը�,dbl�ǲ�������֧��,����ǰ�����ʻ����(�μ�:˵��) ,dbl�����ʻ�֧��
    '       ֧��˳���_IN(�������;ת�ﵥ��;��ϱ���),��ҳID_IN,��;����_IN,�������_IN
    '    ����_IN,��ҩ��_IN,��ҩ��_IN,��ҩ��_IN,����_IN,���Ʒ�_IN,����_IN,����Է�_IN,�������Ʒ�_IN,���������Է�_IN,�������Էѷ���_IN,�Ǳ��շ���_IN,ͳ�����_IN,������,
    '   ����ǰ�����ʻ����,����ǰ�����˻����,����ǰͳ���ۼ�
    '˵��:
    '
    gstrSQL = "zl_���ս����¼_insert(1," & IIf(bln����, lng����ID, lng����ID) & "," & gintInsure & "," & lng����id & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
       dbl�����ʻ���� & "," & dblͳ��֧���ۼ� & "," & dbl��������֧�� & "," & dbl�����ʻ�֧�� & "," & "Null" & "," & dbl�𸶱�׼ & "," & dbl���շ�Χ���Ը� & "," & dbl�𸶱�׼ & "," & _
       curTotal & "," & dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & _
       dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & dbl�ǲ�������֧�� & ",Null," & dbl�����ʻ�֧�� & ",'" & _
       Get�������(intҵ��, g�������_����.�������) & ";" & g�������_����.ת�ﵥ�� & ";" & g�������_����.��ϱ��� & "',null,null,'" & g�������_����.������� & "'," & _
        dbl���� & "," & dbl��ҩ�� & "," & dbl��ҩ�� & "," & dbl��ҩ�� & "," & dbl���� & "," & dbl���Ʒ� & "," & dbl���� & "," & dbl����Է� & "," & dbl�������Ʒ� & "," & dbl���������Է� & "," & dbl�������Էѷ��� & "," & dbl�Ǳ��շ��� & "," & dblͳ����� & "," & dbl������ & "," & _
         dbl����ǰ���(1) & "," & dbl����ǰ���(2) & "," & dbl����ǰ���(3) & _
         " )"
             
    zlDatabase.ExecuteProcedure gstrSQL, "���������շ�����"
    
    
    Set����������� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����id As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Err = 0
    On Error GoTo ErrHand:
    ����������_���� = Set�����������(True, lng����ID, cur�����ʻ�, lng����id, "")
    Exit Function
ErrHand:
    ����������_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����id As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo ErrHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select * From �����ʻ� where  ����=" & gintInsure & "  and ����id=" & lng����id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    strת�ﵥ�� = NVL(rsTemp!��Ա���)
    lng���� = IIf(gintInsure = 83, 2, 1)
    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6) 'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 10)     '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4) 'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 8)     '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    
    strInfor = strInfor & Lpad(NVL(rsTemp!˳���, 1), 4)      '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    
    str������� = Decode(NVL(rsTemp!�Ҷȼ�, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=" & lng����id & _
            "       and A.����ID=" & lng����id & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID"
            
    Call OpenRecordset(rsTemp, "��ȡ��Ժ��Ϣ")
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = NVL(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!סԺ��, 0), 10)      '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(NVL(rsTemp!��Ժ����), 8)      '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(NVL(rsTemp!��Ժ����ʱ��), 16)     '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    If lng���� = 2 Then
        '������Ϊ:סԺ 2���Ҵ� 4ȡ��סԺ�Ǽ� C
        strInfor = strInfor & IIf(str������� = "4", "4", "2")
    Else
        strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    End If

    gstrSQL = "Select * From ��λ״����¼ D where ����ID=" & NVL(rsTemp!��ǰ����ID, 0) & " And ����=" & NVL(rsTemp!��ǰ����, 0)
    Call OpenRecordset(rsTemp, "��ȡ��λ��Ϣ")
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(NVL(rsTemp!�����)) & "��" & Trim(NVL(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = NVL(rsTemp!��Ժ���)
        strȷ��������� = NVL(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
    If lng���� = 2 Then
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
    Else
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
        strInfor = strInfor & Lpad(strȷ����ϱ���, 16)  'ȷ����ϱ���    CHAR    98  16      N   Ժ��
        strInfor = strInfor & Lpad(strȷ���������, 30)  'ȷ���������    CHAR    114 30      N   Ժ��
    End If
    strInfor = strInfor & Lpad(str��Ժ����, 20)  '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    If lng���� = 2 Then
    Else
        strInfor = strInfor & str��λ��              '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    End If
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)   'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)   '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    If lng���� = 2 Then
    Else
        strInfor = strInfor & "A"   '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
        strInfor = strInfor & Space(16)   '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    End If
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    ��Ժ�Ǽ�_���� = ҵ������_����(lng����, 1004, strInfor)
    If ��Ժ�Ǽ�_���� = False Then
        ShowMsgbox "ʵʱסԺ�Ǽ������ύʧ��!"
        Exit Function
    End If
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & gintInsure & ")"
    Call ExecuteProcedure("������Ժ�Ǽ�")
    ��Ժ�Ǽ�_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function ��Ժ�Ǽǳ���_����(lng����id As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
                
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    
    gstrSQL = " Select Count(*) Records From ���˷��ü�¼ " & _
              " Where ����ID=" & lng����id & " And ��ҳID=" & lng��ҳID
    Call OpenRecordset(rsTemp, "������Ժ���")
    
    If rsTemp!Records <> 0 Then
        MsgBox "�Ѿ����ڷ��ü�¼���������������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If

    
    On Error GoTo ErrHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select * From �����ʻ� where  ����=" & gintInsure & "  and ����id=" & lng����id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "������Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    strת�ﵥ�� = NVL(rsTemp!��Ա���)
    lng���� = IIf(gintInsure = 83, 2, 1)

    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6) 'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 10)     '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4) 'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 8)     '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    
    strInfor = strInfor & Lpad(NVL(rsTemp!˳���, 1), 4)      '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    
    str������� = Decode(NVL(rsTemp!�Ҷȼ�, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=" & lng����id & _
            "       and A.����ID=" & lng����id & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID"
            
    Call OpenRecordset(rsTemp, "��ȡ��Ժ��Ϣ")
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = NVL(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!סԺ��, 0), 10)      '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(NVL(rsTemp!��Ժ����), 8)      '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(NVL(rsTemp!��Ժ����ʱ��), 16)      '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    
    If lng���� = 2 Then
        '������Ϊ:סԺ 2���Ҵ� 4ȡ��סԺ�Ǽ� C
        strInfor = strInfor & "C"                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    Else
        strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    End If
    gstrSQL = "Select * From ��λ״����¼ D where ����ID=" & NVL(rsTemp!��ǰ����ID, 0) & " And ����=" & NVL(rsTemp!��ǰ����, 0)
    Call OpenRecordset(rsTemp, "��ȡ��λ��Ϣ")
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(NVL(rsTemp!�����)) & "��" & Trim(NVL(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1 and a.����id=" & lng����id & " and a.��ҳid=" & lng��ҳID
         
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = NVL(rsTemp!��Ժ���)
        strȷ��������� = NVL(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
    If lng���� = 2 Then
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
    Else
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
        strInfor = strInfor & Lpad(strȷ����ϱ���, 16)  'ȷ����ϱ���    CHAR    98  16      N   Ժ��
        strInfor = strInfor & Lpad(strȷ���������, 30)  'ȷ���������    CHAR    114 30      N   Ժ��
    End If
    strInfor = strInfor & Lpad(str��Ժ����, 20)  '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    If lng���� = 2 Then
    Else
        strInfor = strInfor & str��λ��              '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    End If
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)   'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)   '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    If lng���� = 2 Then
    Else
        strInfor = strInfor & "C"   '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
        strInfor = strInfor & Space(16)   '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    End If
    
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    ��Ժ�Ǽǳ���_���� = ҵ������_����(lng����, 1004, strInfor)
    If ��Ժ�Ǽǳ���_���� = False Then
        ShowMsgbox "ʵʱסԺ�Ǽǳ��������ύʧ��!"
        Exit Function
    End If
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & gintInsure & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����id As Long, lng��ҳID As Long) As Boolean
    '
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & gintInsure & ")"
    Call ExecuteProcedure("��Ժ�Ǽ�")
    ��Ժ�Ǽ�_���� = True
End Function
Public Function ��Ժ�Ǽǳ���_����(lng����id As Long, lng��ҳID As Long) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    
    '�������鿨
     lng���� = IIf(gintInsure = 83, 2, 1)
    
     If ��ȡ�������_����(lng����) = False Then Exit Function
    
    '����δ����õĲ��˲�������HIS��Ժ��������Ϊ�Ѱ���ҽ����Ժ���������ٰ���HIS��Ժ
    If Not ����δ�����(lng����id, lng��ҳID) Then
        MsgBox "ҽ���ѳ�Ժ�Ĳ��˲���������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
               
    On Error GoTo ErrHand
    
    '��ȡ���˵���ر�����Ϣ
    gstrSQL = "select * From �����ʻ� where  ����=" & gintInsure & "  and ����id=" & lng����id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "������Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    
    strת�ﵥ�� = NVL(rsTemp!��Ա���)
    

    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6) 'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 10)     '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4) 'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 8)     '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    
    strInfor = strInfor & Lpad(g�������_����.�������, 4)       '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    
    str������� = Decode(NVL(rsTemp!�Ҷȼ�, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=" & lng����id & _
            "       and A.����ID=" & lng����id & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID"
            
    Call OpenRecordset(rsTemp, "��ȡ��Ժ��Ϣ")
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = NVL(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!סԺ��, 0), 10)      '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(NVL(rsTemp!��Ժ����), 8)      '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(NVL(rsTemp!��Ժ����ʱ��), 16)      '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    
    If lng���� = 2 Then
        '������Ϊ:סԺ 2���Ҵ� 4ȡ��סԺ�Ǽ� C
        strInfor = strInfor & "C"                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    Else
        strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    End If
    gstrSQL = "Select * From ��λ״����¼ D where ����ID=" & NVL(rsTemp!��ǰ����ID, 0) & " And ����=" & NVL(rsTemp!��ǰ����, 0)
    Call OpenRecordset(rsTemp, "��ȡ��λ��Ϣ")
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(NVL(rsTemp!�����)) & "��" & Trim(NVL(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1 and a.����id=" & lng����id & " and a.��ҳid=" & lng��ҳID
         
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = NVL(rsTemp!��Ժ���)
        strȷ��������� = NVL(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
    If lng���� = 2 Then
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
    Else
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
        strInfor = strInfor & Lpad(strȷ����ϱ���, 16)  'ȷ����ϱ���    CHAR    98  16      N   Ժ��
        strInfor = strInfor & Lpad(strȷ���������, 30)  'ȷ���������    CHAR    114 30      N   Ժ��
    End If
    
    strInfor = strInfor & Lpad(str��Ժ����, 20)  '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    If lng���� = 2 Then
    Else
        strInfor = strInfor & str��λ��              '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    End If
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)   'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)   '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    
    If lng���� = 2 Then
    Else
        strInfor = strInfor & "A"   '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
        strInfor = strInfor & Space(16)   '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    End If
    
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    ��Ժ�Ǽǳ���_���� = ҵ������_����(lng����, 1004, strInfor)
    If ��Ժ�Ǽǳ���_���� = False Then
        ShowMsgbox "ʵʱסԺ�Ǽǳ��������ύʧ��!"
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & gintInsure & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub ��ȡ������Ϣ_����(ByVal lng����id As Long)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵������Ϣ,����ֵ����G�������
    '--�����:lng����id
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    '��ȡҽ�����������Ϣ�������¹��ýṹ��
        
    gstrSQL = "" & _
        "   Select *" & _
        "   From �����ʻ�" & _
        "   Where ����=" & gintInsure & " And ����ID=" & lng����id
    Call OpenRecordset(rsTemp, "��ȡҽ�����˵������Ϣ")
    
    If Not rsTemp.EOF Then
        With g�������_����
            .IC���� = NVL(rsTemp!����, 0)
            .���˱�� = NVL(rsTemp!ҽ����)
            .ҽ������ = IIf(gintInsure = 83, 2, 1) ' NVL(rsTemp!����, 1)
            .������� = NVL(rsTemp!˳���, 0)
            .ת�ﵥ�� = NVL(rsTemp!��Ա���)
            .���������ʻ���� = NVL(rsTemp!�ʻ����, 0)
            .���������ʻ���� = Val(NVL(rsTemp!����֤��))
            
            .ְ����ҽ��� = Decode(NVL(rsTemp!��ְ, 1), 1, "A", 2, "B", 3, "L", 4, "T", 5, "Q", "")
            .������� = NVL(rsTemp!�Ҷȼ�, 0)
            .�α����3 = NVL(rsTemp!��λ����, 0)
            '.���� = NVL(rsTemp!ͳ�ﱨ���ۼ�, 0)
        End With
    End If
End Sub
Private Function Get�������_����(lng����id As Long, lng��ҳID As Long) As String
    '����:��ȡ���������ʶ
    '     A-������B-��ת��C-δ����D-������E-����
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.��Ժ���" & _
             " From ������ A,��������Ŀ¼ B " & _
             " Where A.����ID=" & lng����id & " And A.����ID=B.ID(+) And A.��ҳID=" & lng��ҳID & _
             "       And A.������� in (2,3)" & _
             " Order by A.������� Desc"
    
    rsInNote.CursorLocation = adUseClient
    Call OpenRecordset(rsInNote, "ҽ���ӿ�", strTmp)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = NVL(rsInNote!��Ժ���)
    End If
    strTmp = Decode(strTmp, "����", "A", "��ת", "B", "δ��", "C", "����", "D", "����", "E")
    
End Function
Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����id As Long) As String
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '      �ֶ�:��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID, _
    '           �շ����,�շ�ϸĿID,�շ�����,��������,���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��, _
    '           �Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    Dim rsTemp As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
'    Dim rs���� As New ADODB.Recordset
    Dim curTotal As Currency
    
    Dim lng��ҳID As Long
    Dim cur�����Ը� As Currency, cur�����ʻ� As Currency
    Dim str��Ժ��� As String, str������� As String
    Dim str����ʱ�� As String, str����ʱ�� As String
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl���� As Double
    Dim dbl���Ʒ� As Double
    Dim dbl���� As Double
    Dim dbl����Է� As Double
    Dim dbl�������Ʒ� As Double
    Dim dbl���������Է� As Double
    Dim dbl�������Էѷ��� As Double
    Dim dbl�Ǳ��շ��� As Double
    Dim dbl���� As Double
    Dim dbl������ As Double     '��Դ�����������
    Dim dbl�𸶱�׼ As Double
    
    Dim str��ϱ��� As String  '��������
    Dim strҽʦ���� As String
    Dim str����Ա���� As String
    Dim str������� As String
    Dim str���������ʶ As String
    Dim strTmp As String
    Dim strҽ�� As String
    Dim str��ϸ As String       '��ϸ��
    Dim str���ұ��� As String
    Dim str��Ŀͳ�Ʒ��� As String
    Dim str��Ժ���� As String
    Dim dbl��Ŀ���� As Double
    Dim strסԺ�� As String
    
    Dim intMouse As Integer
    On Error GoTo ErrHand
    intMouse = Screen.MousePointer
    
    '���������ǰ����֤���
    Screen.MousePointer = 1
    If ��ݱ�ʶ_����(4, lng����id) = "" Then
        Screen.MousePointer = intMouse
        סԺ�������_���� = ""
        Exit Function
    End If
    Screen.MousePointer = intMouse
    
'    '��ȡ������Ϣ
'    Call ��ȡ������Ϣ_����(lng����id)
    
    cur�����ʻ� = g�������_����.���������ʻ����

    gstrSQL = " Select B.סԺ���� ��ҳID,to_char(A.��Ժ����,'yyyy') ��Ժ���,A.��Ժ����,B.סԺ��" & _
              " From ������ҳ A,������Ϣ B" & _
              " Where B.����ID=" & lng����id & " And A.��ҳID=B.סԺ���� And A.����ID=B.����ID"

    Call OpenRecordset(rsTemp, "��ȡ������Ժʱ��")
    str��Ժ��� = rsTemp!��Ժ���
    lng��ҳID = rsTemp!��ҳID
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyymmdd")
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str����ʱ�� = str����ʱ��
    str������� = Mid(str����ʱ��, 1, 4)
    strסԺ�� = NVL(rsTemp!סԺ��)
    
    '���»�ȡ��¼
    Set rs���� = GetסԺ�����¼(lng����id)
    If rs����.RecordCount <= 0 Then
        ShowMsgbox "����Ŀδ����ҽ����Ŀ�����ܽ���!"
        Exit Function
    End If
    dbl�𸶱�׼ = g�������_����.����

    With rs����
        '�ϴ�������ϸ
        curTotal = 0
        Do While Not .EOF
        
            If !��� < 0 Or !�۸� < 0 Or !��¼״̬ <> 1 Then
                MsgBox "�ò��˵Ĵ������������Ŀ�Ľ���۸�Ϊ����,����״̬����ȷ,��������½���!", vbOKOnly
                Exit Function
            End If
        
            If strҽ�� = "" Then
                strҽ�� = NVL(!ҽ�����)
                If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                    strҽ�� = Substr(strҽ��, 1, 6)
                End If
            End If
            curTotal = curTotal + NVL(!���, 0)
            
            lng����id = NVL(!����ID, 0)
            strTmp = NVL(!����ֵ)
            
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '���㱣��
                dbl���� = NVL(!סԺ�ȶ�, 0) / 100
                
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                If g�������_����.ҽ������ <> 2 And g�������_����.ְ����ҽ��� = "L" And g�������_����.�α����3 = "0" And NVL(!������Ŀ��, 0) = 1 Then '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '  ������  ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dbl���� = 1
                End If
                If NVL(!ҽ����Ŀ����) = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If NVL(!ҽ����Ŀ����) = "���" Then
                    strTmp = "����"
                End If
                
                If g�������_����.ҽ������ <> 2 And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dbl���� = Val(NVL(rsTemp!ҽ����Ŀ����))
                End If
                
                If g�������_����.ҽ������ <> 2 And g�������_����.ְ����ҽ��� = "Q" Then
                    '�����Q��ҵ����,�������Ϊ100�Է�,�������Ǳ��շ�����
                    If dbl���� = 0 Then
                        '�Է�100
                        strTmp = ""
                    Else
                        '�ԷѲ��ַ��� �������Էѷ�����
                    End If
                End If
                
                '����Ǵ�λ,���谴���·�ʽ����,��������ͳ���������,������Ϊ100���Է�,���ֿ������ʹ�����
                If NVL(!�շ����) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select ���Ӵ�λ From ���˱䶯��¼ " & _
                        "   Where ����=" & NVL(!����, 0) & _
                        "         And ( (to_date('" & Format(!�Ǽ�ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss') between ��ʼʱ�� and ��ֹʱ��) or" & _
                        "               ( ��ֹʱ�� is null  and ��ʼʱ��<=to_date('" & Format(!�Ǽ�ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))) " & _
                        "         And ���� is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ���Ƿ�Ϊ����!"
                    If rsTemp.RecordCount >= 1 Then
                       If rsTemp!���Ӵ�λ = 1 Then
                            '��ʾ������λ,Ϊȫ�Է�
                            dbl���� = 0
                       End If
                    End If
                End If
                If dbl���� <> 0 Then
                    Select Case strTmp
                        Case "����"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!���, 0), 2)
                                    End If
                                Else
                                    dbl���� = dbl���� + Round(NVL(!���, 0) * dbl����, 2)
                                End If
                        Case "��ҩ��"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���, 0), 2)
                                    End If
                                Else
                                    dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���, 0) * dbl����, 2)
                                End If
                        Case "��ҩ��"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���, 0), 2)
                                    End If
                                Else
                                    dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���, 0) * dbl����, 2)
                                End If
                        Case "��ҩ��"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���, 0), 2)
                                    End If
                                Else
                                    dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���, 0) * dbl����, 2)
                                End If
                        Case "����"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴽��
                                        dbl���� = dbl���� + Round(NVL(!���, 0), 2)
                                    End If
                                Else
                                    dbl���� = dbl���� + Round(NVL(!���, 0) * dbl����, 2)
                                End If
                        Case "���Ʒ�"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴽��
                                        dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!���, 0), 2)
                                    End If
                                Else
                                    dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!���, 0) * dbl����, 2)
                                End If
                        Case "����"
                            If NVL(!�㷨, 0) = 2 Then
                                        '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!��׼����, 0), 2)
                                        If g�������_����.ҽ������ = 1 And g�������_����.ְ����ҽ��� = "Q" Then
                                            '�ԷѲ��ַ��� �������Էѷ�����
                                        Else
                                             dbl����Է� = dbl����Է� + NVL(!���, 0) - NVL(!��׼����, 0)
                                        End If
                                    Else
                                        '�򰴽��
                                        dbl���� = dbl���� + Round(NVL(!���, 0), 2)
                                    End If
                            Else
                                    
                                If g�������_����.ҽ������ = 1 Then
                                    '---��˳��
                                    '�����кͿ������Դ����ô���ͬ,
                                    '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                    dbl���� = dbl���� + Round(NVL(!���, 0) * dbl����, 2)
                                    
                                    If g�������_����.ְ����ҽ��� = "Q" Then
                                        '�ԷѲ��ַ��뱣�����Էѷ�����
                                    Else
                                        dbl����Է� = dbl����Է� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                    End If
                                
                                Else
                                    
                                    dbl���� = dbl���� + Round(NVL(!���, 0), 2)
                                    
                                    dbl����Է� = dbl����Է� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                    
                                End If
                                    
                                    
                                    
    '                                If g�������_����.ҽ������ = 1 Then
    '                                    dbl���� = dbl���� + Round(NVL(!���, 0), 2)
    '                                Else
    '                                    dbl���� = dbl���� + Round(NVL(!���, 0) * dbl����, 2)
    '                                End If
    '
    '                                If g�������_����.ҽ������ = 1 And g�������_����.ְ����ҽ��� = "Q" Then
    '                                    '�ԷѲ��ַ��� �������Էѷ�����
    '                                Else
    '                                    dbl����Է� = dbl����Է� + Round(NVL(!���, 0) * (1 - dbl����), 2)
    '                                End If
                            End If
                        Case "�������Ʒ�"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        '�򰴶������
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴽��
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!���, 0), 2)
                                        If g�������_����.ҽ������ = 1 And g�������_����.ְ����ҽ��� = "Q" Then
                                            '�ԷѲ��ַ��� �������Էѷ�����
                                        Else
                                            dbl���������Է� = dbl���������Է� + NVL(!���, 0) - NVL(!��׼����, 0)
                                        End If
                                    End If
                                Else
                                    '�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                    If g�������_����.ҽ������ = 1 Then
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!���, 0), 2)
                                    Else
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!���, 0) * dbl����, 2)
                                    End If
                                    If g�������_����.ҽ������ = 1 And g�������_����.ְ����ҽ��� = "Q" Then
                                        '�ԷѲ��ַ��� �������Էѷ�����
                                    Else
                                        dbl���������Է� = dbl���������Է� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                    End If
                                End If
                    End Select
                End If
                If g�������_����.ҽ������ <> 2 And g�������_����.ְ����ҽ��� = "Q" Then
                        '�ԷѲ��ַ��� �������Էѷ�����
                         If NVL(!�㷨, 0) = 2 Then
                                '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                If NVL(!��׼����, 0) < NVL(!���, 0) Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���, 0) - NVL(!��׼����, 0), 2)
                                End If
                          Else
                                If dbl���� <> 0 Then
                                    If !������Ŀ�� = 1 Then
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                    End If
                                Else
                                    '100�ԷѲ��ַ���Ǳ��շ�����
                                    dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!���, 0), 2)
                                End If
                          End If
                Else
                         If gintInsure = TYPE_���������� Then
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���, 0) And !������Ŀ�� = 1 Then
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���, 0) - NVL(!��׼����, 0), 2)
                                    End If
                                    If NVL(!��׼����, 0) < NVL(!���, 0) And !������Ŀ�� <> 1 Then
                                        dbl������ = dbl������ + Round(NVL(!���, 0) - NVL(!��׼����, 0), 2)
                                    End If
                                Else
                         
                                    If !������Ŀ�� = 1 And dbl���� <> 0 Then
                                        '����ҩƷ�Է�  NUM 155 10  ҽ����ҩ�ԷѲ���    Ժ����д
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                    Else
                                        '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
                                        dbl������ = dbl������ + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                    End If
                                End If
                         Else
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                     If strTmp <> "�������Ʒ�" And strTmp <> "����" And !������Ŀ�� = 1 And NVL(!��׼����, 0) < NVL(!���, 0) Then
                                        ''ҽ����ҩ�Լ����˴�졢��������������Ŀ���ԷѲ���
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���, 0) - NVL(!��׼����, 0), 2)
                                     End If
                                    If !������Ŀ�� <> 1 Or dbl���� = 0 Then
                                        '��ҽ����ҩ�Լ�������Ŀ
                                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!���, 0), 2)
                                    End If
                                Else
                                    If strTmp <> "�������Ʒ�" And strTmp <> "����" And !������Ŀ�� = 1 And dbl���� <> 0 Then
                                        'ҽ����ҩ�Լ����˴�졢��������������Ŀ���ԷѲ���
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                    End If
                                    If !������Ŀ�� <> 1 Or dbl���� = 0 Then
                                        '��ҽ����ҩ�Լ�������Ŀ
                                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!���, 0), 2)
                                    End If
                                End If
                         End If
                 End If
            End If
            .MoveNext
        Loop
        
        With g�������_����
            If .ҽ������ = 2 Then   '������
                strInfor = Lpad(gstrҽԺ����_����, 6)       'ҽԺ����
            Else
                strInfor = Lpad(gstrҽԺ����_����, 4)       'ҽԺ����
            End If
            strInfor = strInfor & " "      '�������ʶ
            If .ҽ������ = 2 Then   '������
                strInfor = strInfor & Lpad(.���˱��, 10)       '���˱��
            Else
                strInfor = strInfor & Lpad(.���˱��, 8)      '���˱��
            End If
            strInfor = strInfor & Lpad(.IC����, 7)       'IC����
            strInfor = strInfor & Lpad(.������� + 1, 4)      '�������, �鿨���ؽ��ֵ��1
            strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '����ʱ��
            strInfor = strInfor & Lpad(strסԺ��, 10) '��־��:סԺ��ΪסԺ��
            
            
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����, 2))), 10) '����
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10) '��ҩ��
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10) '��ҩ��
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10) '��ҩ��
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����, 2))), 10) '����
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl���Ʒ�, 2))), 10)  '���Ʒ�
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����, 2))), 10)  '����
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl�������Ʒ�, 2))), 10)  '�������Ʒ�
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl����Է�, 2))), 10)  '����Է�
            If .ҽ������ = 2 Then
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl���������Է�, 2))), 10)   '�����Է�    NUM 145 10      Ժ����д
            End If
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl�������Էѷ���, 2))), 10)   '�������Էѷ���
            
            If .ҽ������ = 2 Then        '������
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl������, 2))), 10)   '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
            Else
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl�Ǳ��շ���, 2))), 10)    '�Ǳ��շ���
            End If
            
            strInfor = strInfor & String(10, " ")    '���ķ���:���������ʻ����;������:���������ʻ����  NUM 175 10  ���������ʻ������������ʻ�  ���ķ���
            strInfor = strInfor & String(10, " ")    '���ķ���:�����ͳ��֧���ۼ�  NUM 185 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ���ķ���
            
            '����ǰ�����ʻ��������鿨���ؽ�����������������Ӧ����������ѯ�������ʻ������д��
            strInfor = strInfor & Lpad(.���������ʻ����, 10)  '����ǰ�����ʻ����
            strInfor = strInfor & Lpad(Trim(CStr(.���������ʻ����)), 10)   '����ǰ�����˻����(�����鿨���ؽ�������������������0)
            strInfor = strInfor & Lpad(Trim(CStr(.ͳ���ۼ�)), 10)    '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ��������ʻ�֧��(������������㣬��ʾ�����ʻ�֧��)
            strInfor = strInfor & String(10, " ")    '���ķ���:���β��������ʻ�֧��(������������㷵��0)
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ��֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ���Ը�
            strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ��֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ���Ը�
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ�����������֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���ηǻ�����������֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���α��շ�Χ���Ը�
              
              
            If .ҽ������ <> 2 Then
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl���������Է�, 2))), 10)   '�������������Ը�
            End If
            
            strInfor = strInfor & Lpad(Trim(CStr(dbl�𸶱�׼)), 10)    '�𸶱�׼
              
            strInfor = strInfor & Lpad(.ת�ﵥ��, 6)     'ת�ﵥ��
            strInfor = strInfor & Lpad(Get�������(0, .�������), 1)     '�������
            If .ҽ������ <> 2 Then
                strInfor = strInfor & Lpad(.�α����3, 1)    '�α����3:0 �󱣡�1 �±��������鿨���
            End If
            strInfor = strInfor & Lpad(.ְ����ҽ���, 1)       'ְ����ҽ���
              
            strInfor = strInfor & Lpad(.��ϱ���, 16)    '��ϱ���
            strInfor = strInfor & Lpad(strҽ��, 6)    'ҽʦ����
            strInfor = strInfor & Lpad(UserInfo.���, 6)    '����Ա����
            strInfor = strInfor & Lpad(.�������, 30)    '�������
            
            'A-������B-��ת��C-δ����D-������E-����
            strInfor = strInfor & Lpad(Get�������_����(lng����id, lng��ҳID), 1)   '���������ʶ
            strInfor = strInfor & Lpad(IIf(str��Ժ���� = "", Format(zlDatabase.Currentdate, "yyyyMMDD"), str��Ժ����), 8) '��Ժ����
            
            If .ҽ������ = 2 Then       '������
            Else
                strInfor = strInfor & String(16, " ")      '����ʱ��
            End If
            strInfor = strInfor & String(10, " ")      '�������
          End With
    
        '��������ӿ�(1006    12  423   ʵʱ����Ԥ��
        If ҵ������_����(g�������_����.ҽ������, 1006, strInfor) = False Then
            ShowMsgbox "סԺ�������ʧ��!"
            Exit Function
        End If
        
        
        g�������_����.֧����� = curTotal
    End With
    
    Dim str���㷽ʽ  As String
    
  
    '������:
    '    ���λ��������ʻ�֧��    NUM 225 10      ���ķ���
    '    ���β��������ʻ�֧��    NUM 235 10      ���ķ���
    '    ���λ���ͳ��֧��    NUM 245 10      ���ķ���
    '    ���λ���ͳ���Ը�    NUM 255 10      ���ķ���
    '    ���β���ͳ��֧��    NUM 265 10      ���ķ���
    '    ���β���ͳ���Ը�    NUM 275 10      ���ķ���
    '    ���λ�����������֧��    NUM 285 10  ����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
    '    ���ηǻ�����������֧��  NUM 295 10  ����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
    '    ���α��շ�Χ���Ը�  NUM 305 10  �޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
    '������:
    '    ���λ��������ʻ�֧��    NUM 211 10  ������������㣬��ʾ�����ʻ�֧��    ����
    '    ���β��������ʻ�֧��    NUM 221 10  ������������㷵��0 ����
    '    ���λ���ͳ��֧��    NUM 231 10      ����
    '    ���λ���ͳ���Ը�    NUM 241 10      ����
    '    ���β���ͳ��֧��    NUM 251 10  ������������㣬���ֶ����ڴ����������֧��  ����
    '    ���β���ͳ���Ը�    NUM 261 10      ����
    '    ���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧�� 2�� ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
    '    ���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��   2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����    ����
    '    ���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    
    Dim i As Long
    
    If g�������_����.ҽ������ = 2 Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    'ȷ�����ν��㷽ʽ
    str���㷽ʽ = "�����ʻ�;" & Format(Val(Substr(strInfor, i + 10, 10)), "###0.00;-###0.00;0;0") & ";0" '���λ��������ʻ�֧��,�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "�����ʻ�;" & Format(Val(Substr(strInfor, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strInfor, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strInfor, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "��������;" & Format(Val(Substr(strInfor, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "�ǲ�������;" & Format(Val(Substr(strInfor, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    
    סԺ�������_���� = str���㷽ʽ
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Get�����ʻ����_����() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�����ʻ����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    'ҽԺ����    CHAR    1   4       Ժ��
    '���˱��    CHAR    5   8       Ժ��
    '��������    CHAR    13  16  ĿǰΪ: WZMB    Ժ��
    '�������    NUM 29  4       ����
    '�����ʻ�ԭʼֵ  NUM 33  10  ÿ�β������ۼ�ֵ    ����
    '�����ʻ���ǰֵ  NUM 43  10      ����
    '�ʻ�״̬    CHAR    53  1   A������Cֹ��    ����

    
    Dim strTmp As String
    Err = 0
    On Error GoTo ErrHand:
    With g�������_����
        strTmp = Lpad(gstrҽԺ����_����, 4)      'ҽԺ����    CHAR    1   4       Ժ��
        strTmp = strTmp & Lpad(.���˱��, 8) '���˱��    CHAR    5   8       Ժ��
        strTmp = strTmp & Lpad("WZMB", 16)  '��������    CHAR    13  16  ĿǰΪ: WZMB    Ժ��
        strTmp = strTmp & Space(4)  '�������    NUM 29  4       ����
        strTmp = strTmp & Space(10)  '�����ʻ�ԭʼֵ  NUM 33  10  ÿ�β������ۼ�ֵ    ����
        strTmp = strTmp & Space(10)  '�����ʻ���ǰֵ  NUM 43  10      ����
        strTmp = strTmp & Space(1)   '�ʻ�״̬    CHAR    53  1   A������Cֹ��    ����
        '��ҽ�����Ĳ�ѯ����
        '   1007    2   55  �����ʻ���ѯ
        Get�����ʻ����_���� = ҵ������_����(.ҽ������, 1007, strTmp)
        If Get�����ʻ����_���� = False Then
            .�����ʻ�ԭʼֵ = 0
            .�����ʻ���ǰֵ = 0
            Exit Function
        End If
        .�����ʻ�ԭʼֵ = Val(Substr(strTmp, 33, 10))
        .�����ʻ���ǰֵ = Val(Substr(strTmp, 43, 10))
    End With
    Exit Function
ErrHand:
End Function
Private Function סԺ���㼰����_����(ByVal bln���� As Boolean, ByVal lng����id As Long, ByVal lng����ID As Long, ByVal ԭ����id As Long, ByVal lng��ҳID As Long) As Boolean

    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim strסԺ�� As String
    Dim str��Ŀͳ�Ʒ���  As String
    Dim strInfor As String  '�������ķ��ش�
    Dim curTotal As Double
    Dim dbl���� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl��ҩ�� As Double
    Dim dbl���� As Double
    Dim dbl���Ʒ� As Double
    Dim dbl���� As Double
    Dim dbl����Է� As Double
    Dim dbl�������Ʒ� As Double
    Dim dbl���������Է� As Double
    Dim dbl�������Էѷ��� As Double
    Dim dbl�Ǳ��շ��� As Double
    Dim dbl���� As Double
    Dim dbl������ As Double     '��Դ�����������
    Dim dbl�𸶱�׼ As Double
   
    Dim dbl�����ʻ���� As Double
    Dim dblͳ��֧���ۼ� As Double
    Dim dbl�����ʻ�֧�� As Double
    Dim dbl�����ʻ�֧�� As Double
    Dim dbl����ͳ��֧�� As Double
    Dim dbl����ͳ���Ը� As Double
    Dim dbl����ͳ��֧�� As Double
    Dim dbl����ͳ���Ը� As Double
    Dim dbl��������֧�� As Double
    Dim dbl�ǲ�������֧�� As Double
    Dim dbl���շ�Χ���Ը� As Double
    
    Dim dbl����ǰ�����ʻ����  As Double
    Dim dbl����ǰ�����˻����  As Double
    Dim dbl����ǰͳ���ۼ�  As Double
    
    Dim strҽ�� As String
    Dim str��ϸ As String       '��ϸ��
    Dim str���ұ��� As String
    Dim intҵ�� As Integer
    Dim str��Ժ���� As String
    
    intҵ�� = IIf(bln����, 1, 0)
    
    Err = 0
    On Error GoTo ErrHand:
    
    'סԺӦ�ñ���֧�������е�סԺ�ȶ�
    gstrSQL = " " & _
        "        select a.id,a.��¼����,a.��ҳid,a.��¼״̬,a.�Ǽ�ʱ��,a.no,a.���˲���id,a.����,a.���,a.��ʶ�� as סԺ��,a.���˿���id,a.����id,a.�շ����,b.���,a.���㵥λ, " & _
        "               A.���㵥λ,A.����*Nvl(A.����,1) ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ�� ,a.������ as ҽ��,c.��� as ҽ�����, " & _
        "               a.ҽ�����, A.ʵ�ս��,A.�Ƿ��ϴ�, " & _
        "               F.����ֵ,D.���� as ��Ŀ����,D.���� as ��Ŀ����,D.��ʶ����||D.��ʶ���� as ���ұ���, " & _
        "               E.��Ŀ���� as ҽ������,E.��Ŀ���� as ҽ������,e.�Ƿ�ҽ��,e.����id,G.סԺ�ȶ� as ͳ��ȶ�,G.��׼����,G.�㷨,H.���� as ��������,J.���� as ����, " & _
        "               L.����,l.���� , l.����, l.ҽ����, l.��Ա���, l.��λ����, l.˳���, l.����֤��, l.�ʻ����, l.��ǰ״̬, l.����ID, l.��ְ, l.�����, l.�Ҷȼ�, l.����ʱ�� " & _
        "        from ���˷��ü�¼ a,�շ���� b,��Ա�� c,�շ�ϸĿ D,����֧����Ŀ E,����֧������ G,�����ʻ� L,���ű� H, " & _
        "             (Select U.*,K.����ֵ From �շ���� U,���ղ��� K where U.���=K.������ and K.����=" & gintInsure & "  ) F ," & _
        "             (Select distinct Q.ҩƷid,T.���� From ҩƷĿ¼ Q,ҩƷ��Ϣ R,ҩƷ���� T  Where  Q.ҩ��id=R.ҩ��id and R.����=T.���� ) J " & _
        "        where a.�շ����=b.���� and a.�շ�ϸĿid=J.ҩƷid(+)   and  Nvl(a.���ӱ�־,0)<>9 and a.�շ�ϸĿid=D.id and a.������=c.����(+) and a.�շ����=F.����(+) and " & _
        "              a.�շ�ϸĿid=E.�շ�ϸĿID  and E.����id=G.id and a.����id=L.����ID and a.��������id=h.id  and " & _
        "              a.����ID = " & lng����id & " And a.����ID = " & lng����ID & " And E.���� = " & gintInsure
        
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡסԺ������ϸ"
    
    'ȷ���ò����Ƿ��Ѿ���Ժ
    gstrSQL = "Select * From ������ҳ where ����id=" & lng����id & " and ��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�����Ƿ��Ժ"
    str��Ժ���� = ""
    If rsTemp.EOF Then
        strTmp = Get��Ժ���(lng����id, lng��ҳID, , True)
    Else
        If IsNull(rsTemp!��Ժ����) Then
            strTmp = Get��Ժ���(lng����id, lng��ҳID, , True)
        Else
            strTmp = ��ȡ���Ժ���(lng����id, lng��ҳID, False, , True)
            str��Ժ���� = Format(rsTemp!��Ժ����, "yyyymmdd")
        End If
    End If
    If InStr(1, strTmp, "|") <> 0 Then
        g�������_����.��ϱ��� = Split(strTmp, "|")(1)
        g�������_����.������� = Split(strTmp, "|")(0)
    End If
    
    With rs��ϸ
        If Not .EOF Then
            strסԺ�� = NVL(!סԺ��)
        End If
        Do While Not .EOF
            strTmp = NVL(!����ֵ)
            lng����id = NVL(!����ID, 0)
            If strҽ�� = "" Then
                strҽ�� = NVL(!ҽ�����)
                If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                    strҽ�� = Substr(strҽ��, 1, 6)
                End If
            End If
            'ȷ���������
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str��Ŀͳ�Ʒ��� = ""
                Else
                    str��Ŀͳ�Ʒ��� = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '����
                dbl���� = NVL(!ͳ��ȶ�, 0) / 100
                If NVL(!����, 0) <> TYPE_���������� And Val(NVL(!��λ����, "99")) = 0 And NVL(!��ְ, 0) = 3 And NVL(!�Ƿ�ҽ��, 0) = 1 Then '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '    ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dbl���� = 1
                End If
                If NVL(!ҽ������) = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If NVL(!ҽ������) = "���" Then
                    strTmp = "����"
                End If
                If NVL(!����, 0) = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dbl���� = Val(NVL(!ҽ������))
                End If
                
                If NVL(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                    '�����Q��ҵ����,�������Ϊ100�Է�,�������Ǳ��շ�����
                    If dbl���� = 0 Then
                        '�Է�100
                        strTmp = ""
                    Else
                        '�ԷѲ��ַ��� �������Էѷ�����
                    End If
                End If
                '����Ǵ�λ,���谴���·�ʽ����,��������ͳ���������,������Ϊ100���Է�,���ֿ������ʹ�����
                If NVL(!�շ����) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select ���Ӵ�λ From ���˱䶯��¼ " & _
                        "   Where ����=" & NVL(!����, 0) & _
                        "         And ( (to_date('" & Format(!�Ǽ�ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss') between ��ʼʱ�� and ��ֹʱ��) or" & _
                        "               ( ��ֹʱ�� is null  and ��ʼʱ��<=to_date('" & Format(!�Ǽ�ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))) " & _
                        "         And ���� is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ���Ƿ�Ϊ����!"
                    If rsTemp.RecordCount >= 1 Then
                       If rsTemp!���Ӵ�λ = 1 Then
                            '��ʾ������λ,Ϊȫ�Է�
                            dbl���� = 0
                       End If
                    End If
                End If
                If dbl���� <> 0 Then
                    '---��˳��
                    '---��Ϊ�Ա�������Ϊ0����Ŀ����,�Ƕ��κ��˶�����Ϊ�������Է�
                    '---��Ϊ������������,Ҳ����Ϊȫ�Էѷ��ô���,��˷���dbl����=0��ֱ���������������Էѷ��ô���
                    '---ԭ���Ĵ��������dbl����=0�İ������������ƷѺͱ�������ö������˴������÷������ظ�
                               
                    Select Case strTmp
                        Case "����"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                Else
                                    dbl���� = dbl���� + Round(NVL(!���ʽ��, 0) * dbl����, 2)
                                End If
                        Case "��ҩ��"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                Else
                                    dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���ʽ��, 0) * dbl����, 2)
                                End If
                        Case "��ҩ��"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                Else
                                    dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���ʽ��, 0) * dbl����, 2)
                                End If
                        Case "��ҩ��"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴶������
                                        dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                Else
                                    dbl��ҩ�� = dbl��ҩ�� + Round(NVL(!���ʽ��, 0) * dbl����, 2)
                                End If
                        Case "����"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴽��ʽ��
                                        dbl���� = dbl���� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                Else
                                    dbl���� = dbl���� + Round(NVL(!���ʽ��, 0) * dbl����, 2)
                                End If
                        Case "���Ʒ�"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴽��ʽ��
                                        dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                Else
                                    dbl���Ʒ� = dbl���Ʒ� + Round(NVL(!���ʽ��, 0) * dbl����, 2)
                                End If
                        Case "����"
                            If NVL(!�㷨, 0) = 2 Then
                                        '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl���� = dbl���� + Round(NVL(!��׼����, 0), 2)
                                        If NVL(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                                            '�ԷѲ��ַ��� �������Էѷ�����
                                        Else
                                             dbl����Է� = dbl����Է� + NVL(!���ʽ��, 0) - NVL(!��׼����, 0)
                                        End If
                                    Else
                                        '�򰴽��ʽ��
                                        dbl���� = dbl���� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                            Else
                                    If NVL(!����, 0) = TYPE_������ Then
                                        '---��˳��
                                        '�����кͿ������Դ����ô���ͬ,
                                        '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                        
                                          dbl���� = dbl���� + Round(NVL(!���, 0) * dbl����, 2)
                                          
                                          If g�������_����.ְ����ҽ��� = "Q" Then
                                              '�ԷѲ��ַ��뱣�����Էѷ�����
                                          Else
                                              dbl����Է� = dbl����Է� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                          End If
                                      
                                      Else
                                          
                                          dbl���� = dbl���� + Round(NVL(!���, 0), 2)
                                          
                                          dbl����Է� = dbl����Է� + Round(NVL(!���, 0) * (1 - dbl����), 2)
                                          
                                      End If
                                    
                                    'If NVL(!����, 0) = TYPE_������ Then
                                    '    dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0), 2)
                                    'Else
                                    '    dbl���� = dbl���� + Round(NVL(!ʵ�ս��, 0) * dbl����, 2)
                                    'End If
                                    '
                                    'If NVL(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                                    '    '�ԷѲ��ַ��� �������Էѷ�����
                                    'Else
                                    '    dbl����Է� = dbl����Է� + Round(NVL(!���ʽ��, 0) * (1 - dbl����), 2)
                                    'End If
                            End If
                        Case "�������Ʒ�"
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        '�򰴶������
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!��׼����, 0), 2)
                                    Else
                                        '�򰴽��ʽ��
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!���ʽ��, 0), 2)
                                        If NVL(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                                            '�ԷѲ��ַ��� �������Էѷ�����
                                        Else
                                            dbl���������Է� = dbl���������Է� + NVL(!���ʽ��, 0) - NVL(!��׼����, 0)
                                        End If
                                    End If
                                Else
                                    '�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                    If NVL(!����, 0) = TYPE_������ Then
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!���ʽ��, 0), 2)
                                    Else
                                        dbl�������Ʒ� = dbl�������Ʒ� + Round(NVL(!���ʽ��, 0) * dbl����, 2)
                                    End If
                                    If NVL(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                                        '�ԷѲ��ַ��� �������Էѷ�����
                                    Else
                                        dbl���������Է� = dbl���������Է� + Round(NVL(!���ʽ��, 0) * (1 - dbl����), 2)
                                    End If
                                End If
                    End Select
                End If
                If NVL(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                        '�ԷѲ��ַ��� �������Էѷ�����
                         If NVL(!�㷨, 0) = 2 Then
                                '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���ʽ��, 0) - NVL(!��׼����, 0), 2)
                                End If
                          Else
                                If dbl���� <> 0 Then
                                    If !�Ƿ�ҽ�� = 1 Then
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���ʽ��, 0) * (1 - dbl����), 2)
                                    End If
                                Else
                                    '100�ԷѲ��ַ���Ǳ��շ�����
                                    dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!���ʽ��, 0), 2)
                                End If
                          End If
                Else
                         If gintInsure = TYPE_���������� Then
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) And !�Ƿ�ҽ�� = 1 Then
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���ʽ��, 0) - NVL(!��׼����, 0), 2)
                                    End If
                                    If NVL(!��׼����, 0) < NVL(!���ʽ��, 0) And !�Ƿ�ҽ�� <> 1 Then
                                        dbl������ = dbl������ + Round(NVL(!���ʽ��, 0) - NVL(!��׼����, 0), 2)
                                    End If
                                Else
                         
                                    If !�Ƿ�ҽ�� = 1 And dbl���� <> 0 Then
                                        '����ҩƷ�Է�  NUM 155 10  ҽ����ҩ�ԷѲ���    Ժ����д
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���ʽ��, 0) * (1 - dbl����), 2)
                                    Else
                                        '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
                                        dbl������ = dbl������ + Round(NVL(!���ʽ��, 0) * (1 - dbl����), 2)
                                    End If
                                End If
                         Else
                                If NVL(!�㷨, 0) = 2 Then
                                    '���˺�:200404,���õ��㷨2(���ö����),������б�������.
                                     If strTmp <> "�������Ʒ�" And strTmp <> "����" And !�Ƿ�ҽ�� = 1 And NVL(!��׼����, 0) < NVL(!���ʽ��, 0) Then
                                        ''ҽ����ҩ�Լ����˴�졢��������������Ŀ���ԷѲ���
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���ʽ��, 0) - NVL(!��׼����, 0), 2)
                                     End If
                                    If !�Ƿ�ҽ�� <> 1 Or dbl���� = 0 Then
                                        '��ҽ����ҩ�Լ�������Ŀ
                                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                Else
                                    If strTmp <> "�������Ʒ�" And strTmp <> "����" And !�Ƿ�ҽ�� = 1 And dbl���� <> 0 Then
                                        'ҽ����ҩ�Լ����˴�졢��������������Ŀ���ԷѲ���
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(NVL(!���ʽ��, 0) * (1 - dbl����), 2)
                                    End If
                                    If !�Ƿ�ҽ�� <> 1 Or dbl���� = 0 Then
                                        '��ҽ����ҩ�Լ�������Ŀ
                                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(NVL(!���ʽ��, 0), 2)
                                    End If
                                End If
                         End If
                 End If
            Else
                dbl���� = 1
                str��Ŀͳ�Ʒ��� = ""
            End If

 
            '�ϴ���ϸ��¼,ʵʱҽ����ϸ����
            If gblnסԺ��ϸʱʵ�ϴ� And bln���� = False And NVL(!�Ƿ��ϴ�, 0) = 0 Then
                    If NVL(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = Lpad(gstrҽԺ����_����, 6)     'ҽԺ����    CHAR    1   6       Ժ����д
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ����), 10)  '���ձ��    CHAR    7   10      Ժ����д
                    Else
                        str��ϸ = Lpad(gstrҽԺ����_����, 4)     'ҽԺ����    CHAR    1   4       Ժ��
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ����), 8)   '���˱��    CHAR    5   8       Ժ��
                    End If
                    
                    str��ϸ = str��ϸ & Lpad(NVL(!סԺ��, 0), 10) '��־��  CHAR    13  10  ������ϸ�Կո�λ,סԺ��סԺ��  Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!˳���, 0), 4)   '�������    NUM 23  4   סԺ��ϸ�����������Ժ�Ǽ�ʱ�������������ϸ:                         ������ڱ��ν���������� Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!NO, 0), 10)       '������  NUM 27  10      Ժ��
                    
                    If NVL(!����, 0) = TYPE_���������� Then '������
                    Else
                        str��ϸ = str��ϸ & Lpad(NVL(!���, 0), 10)      '������Ŀ���    NUM 37  10  ��Ӧ�����ŵļǼ���Ŀ���    Ժ��
                    End If
                    
                    '������Ϊ���ݺ�  CHAR    41  10  ҽ���ţ�    Ժ����д
                    str��ϸ = str��ϸ & Lpad(NVL(!ҽ�����, 0), 10)     'ҽ����  CHAR    47  10  ������Ӧҽ����ҽ����¼�ţ�������ϸ��û��ҽ����ҽԺ�Կո�λ    Ժ��
                    g�������_����.������� = NVL(!�Ҷȼ�, 0)
                    
                    str��ϸ = str��ϸ & Get�������(intҵ��, NVL(!�Ҷȼ�, 0))         '�������    CHAR    57  1   ȡֵ���"�������"˵��  Ժ��
                    
                    If NVL(!����, 0) = TYPE_���������� Then  '������
                        '������Ϊ����ʱ��    DATETIME    52  16  ��ȷ���루������ʱ�䣩��ʽΪ��yyyymmddhhmiss�����Կո�λ  Ժ����д
                        str��ϸ = str��ϸ & Rpad(Format(!����ʱ��, "yyyymmddHHmmss"), 16)
                    Else
                        str��ϸ = str��ϸ & Rpad(Format(!�Ǽ�ʱ��, "yyyymmddHHmmss"), 16)      '��������ʱ�䣨Ͷҩʱ�䣩    DATETIME    58  16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ    Ժ��
                    End If
                    
                    str��ϸ = str��ϸ & Lpad(NVL(!���ұ���), 20)      '��Ŀ����    CHAR    74  20  �Ƽ���Ŀ����    Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!��Ŀ����), 20)      '��Ŀ����    CHAR    94  20      Ժ��
        
                    If NVL(!����, 0) = TYPE_���������� Then  '������
                    Else
        
                        If !�Ƿ�ҽ�� = 1 Then
                            str��ϸ = str��ϸ & Lpad(1 - dbl����, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                        Else
                            str��ϸ = str��ϸ & Lpad(1, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                        End If
                        str��ϸ = str��ϸ & Lpad(str��Ŀͳ�Ʒ���, 1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(NVL(!����), 6)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!ʵ�ʼ۸�), 8) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!���㵥λ), 4) '��λ    CHAR    135 4       Ժ��
                    str��ϸ = str��ϸ & Lpad(NVL(!����), 20)      '����    CHAR    139 20  �����Ƭ����    Ժ��
                    
                    If NVL(!����, 0) = TYPE_���������� Then  '������
                        '��ȡ���˵�����.
                        gstrSQL = "Select ����,Ƶ��,�÷� From ҩƷ�շ���¼ where ����id=" & NVL(!ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵���Ƶ��"
                        If rsTemp.EOF Then
                            str��ϸ = str��ϸ & Space(5)       'ÿ������    NUM 146 5       Ժ����д
                            str��ϸ = str��ϸ & Space(20)      'ʹ��Ƶ��    CHAR    151 20  �磺1��2��  Ժ����д
                            str��ϸ = str��ϸ & Space(50)      '�÷�    CHAR    171 50  �磺�ڷ�    Ժ����д
                        Else
                            str��ϸ = str��ϸ & Lpad(NVL(rsTemp!����), 5)      'ÿ������    NUM 146 5       Ժ����д
                            str��ϸ = str��ϸ & Lpad(NVL(rsTemp!Ƶ��), 20)      'ʹ��Ƶ��    CHAR    151 20  �磺1��2��  Ժ����д
                            str��ϸ = str��ϸ & Lpad(NVL(rsTemp!�÷�), 50)      '�÷�    CHAR    171 50  �磺�ڷ�    Ժ����д
                        End If
                        str��ϸ = str��ϸ & Space(4)      'ִ������    NUM 221 4       Ժ����д
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ�����), 6)      'ҽʦ����    CHAR    225 6       Ժ����д
                    Else
                        str��ϸ = str��ϸ & Lpad(NVL(!ҽ��), 8)      'ҽʦ����    CHAR    159 8       Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(g�������_����.��ϱ���, 16)      '��ϱ���    CHAR    167 16      Ժ��
                    str��ϸ = str��ϸ & Lpad(g�������_����.�������, 30)     '�������    CHAR    183 30      Ժ��
                    If NVL(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = str��ϸ & Lpad(NVL(!��������), 20)    '�Ʊ�����    CHAR    277 20      Ժ����д
                    Else
                        str��ϸ = str��ϸ & Space(16)     '����ʱ��    DATETIME    213 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ��Ժ�˿ո�λ  ����
                    End If
                    
                    '�ϴ���ϸ
                    '1003    7   230 ʵʱҽ����ϸ�����ύ
                    סԺ���㼰����_���� = ҵ������_����(IIf(NVL(!����, 0) = TYPE_����������, 2, 1), 1003, str��ϸ)
                    If סԺ���㼰����_���� = False Then
                        ShowMsgbox "סԺ����������ϸ�����ύʧ��,���ܼ���!"
                        Exit Function
                    End If
                    '�ϴ�ҽ����ϸ
                    If NVL(!ҽ�����, 0) <> 0 Then
                    
                        If ҽ����ϸ�����ύ(!ҽ�����, NVL(!סԺ��), str��Ŀͳ�Ʒ���) = False Then
                            ShowMsgbox "ҽ����ϸ�����ύʧ��,���ܼ���!"
                            Exit Function
                        End If
                    End If
                    'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & NVL(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            End If
            '�����ܶ�,����
            curTotal = curTotal + Round(NVL(!���ʽ��, 0), 2)
            .MoveNext
        Loop
    End With
    
 
    '��д�����¼
    '��������
    dbl�𸶱�׼ = g�������_����.����
    
    If bln���� Then
          '��ȷ���ϴ����ķ��ص�����
    
           gstrSQL = "" & _
               "   Select *  " & _
               "   From ���ս����¼ " & _
               "   Where ��¼id=" & ԭ����id
           zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�����շ�ʱ���ص�����"
           If rsTemp.RecordCount = 0 Then
               ShowMsgbox "�������ϴ��շѵĽ����¼!"
               Exit Function
           End If
           '/???
           'ԭ���̲���:
           '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
           "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
           '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
           '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
           '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
           '������ֵ����Ϊ:
           '       ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN, _
           '       dbl�����ʻ����,dblͳ��֧���ۼ�,dbl��������֧��,dbl�����ʻ�֧��,סԺ����_IN,����_IN,dbl���շ�Χ���Ը�,ʵ������_IN
           '       �������ý��_IN,dbl����ͳ��֧��,dbl����ͳ���Ը�,
           '       dbl����ͳ��֧��,dbl����ͳ���Ը�,dbl�ǲ�������֧��,�����Ը����_IN,dbl�����ʻ�֧��
           '       ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
            dbl�����ʻ���� = Round(NVL(rsTemp!�ʻ��ۼ�����, 0), 2)
            dblͳ��֧���ۼ� = Round(NVL(rsTemp!�ʻ��ۼ�֧��, 0), 2)
            dbl��������֧�� = Round(NVL(rsTemp!�ۼƽ���ͳ��, 0), 2)
            dbl�����ʻ�֧�� = Round(NVL(rsTemp!�ۼ�ͳ�ﱨ��, 0), 2)
            dbl�𸶱�׼ = Round(NVL(rsTemp!����, 0), 2)
            dbl���շ�Χ���Ը� = Round(NVL(rsTemp!�ⶥ��, 0), 2)
            dbl����ͳ��֧�� = Round(NVL(rsTemp!ȫ�Ը����, 0), 2)
            dbl����ͳ���Ը� = Round(NVL(rsTemp!�����Ը����, 0), 2)
            dbl����ͳ��֧�� = Round(NVL(rsTemp!����ͳ����, 0), 2)
            dbl����ͳ���Ը� = Round(NVL(rsTemp!ͳ�ﱨ�����, 0), 2)
            dbl�ǲ�������֧�� = Round(NVL(rsTemp!���Ը����, 0), 2)
            dbl�����ʻ�֧�� = Round(NVL(rsTemp!�����ʻ�֧��, 0), 2)
            
            dbl����ǰ�����ʻ���� = Round(NVL(rsTemp!����ǰ�����ʻ����, 0), 2)
            dbl����ǰ�����˻���� = Round(NVL(rsTemp!����ǰ�����˻����, 0), 2)
            dbl����ǰͳ���ۼ� = Round(NVL(rsTemp!����ǰͳ���ۼ�, 0), 2)
              
       End If
    '�ҳ���������
    With g�������_����
        If gintInsure = TYPE_���������� Then    '������
            strInfor = Lpad(gstrҽԺ����_����, 6)       'ҽԺ����
        Else
            strInfor = Lpad(gstrҽԺ����_����, 4)       'ҽԺ����
        End If
        strInfor = strInfor & " "      '�������ʶ
        If gintInsure = TYPE_���������� Then    '������
            strInfor = strInfor & Lpad(.���˱��, 10)       '���˱��
        Else
            strInfor = strInfor & Lpad(.���˱��, 8)      '���˱��
        End If
        strInfor = strInfor & Lpad(.IC����, 7)       'IC����
        strInfor = strInfor & Lpad(.������� + 1, 4)      '�������
        strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '����ʱ��
        strInfor = strInfor & Lpad(strסԺ��, 10) '��־��
        
        
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����), 2))), 10) '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl��ҩ��), 2))), 10) '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl��ҩ��), 2))), 10) '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl��ҩ��), 2))), 10) '��ҩ��
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����), 2))), 10) '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl���Ʒ�), 2))), 10)  '���Ʒ�
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����), 2))), 10)  '����
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl�������Ʒ�), 2))), 10)  '�������Ʒ�
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl����Է�), 2))), 10)  '����Է�
        If gintInsure = TYPE_���������� Then        '������
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl���������Է�), 2))), 10)   '�����Է�    NUM 145 10      Ժ����д
        End If
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl�������Էѷ���), 2))), 10)   '�������Էѷ���
        
        If gintInsure = TYPE_���������� Then        '������
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl������), 2))), 10)   '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д
        Else
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl�Ǳ��շ���), 2))), 10)    '�Ǳ��շ���
        End If
    
        Dim dbl����ǰ���(1 To 3) As Double '1-����ǰ�����ʻ����,2-����ǰ�����˻����,3-����ǰͳ��֧���ۼ�
        dbl����ǰ���(1) = .���������ʻ����
        dbl����ǰ���(2) = .���������ʻ����
        dbl����ǰ���(3) = .ͳ���ۼ�
        
        If bln���� Then
            strInfor = strInfor & Lpad(dbl�����ʻ����, 10)
            strInfor = strInfor & Lpad(dblͳ��֧���ۼ�, 10)
            strInfor = strInfor & Lpad(dbl����ǰ�����ʻ����, 10)   '����ǰ�����ʻ����
            strInfor = strInfor & Lpad(dbl����ǰ�����˻����, 10)    '����ǰ�����˻����(�����鿨���ؽ�������������������0)
            strInfor = strInfor & Lpad(dbl����ǰͳ���ۼ�, 10)     '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
            strInfor = strInfor & Lpad(dbl�����ʻ�֧��, 10) ' = Round(NVL(rsTemp!�����ʻ�֧��, 0), 2)
            strInfor = strInfor & Lpad(dbl�����ʻ�֧��, 10) ' = Round(NVL(rsTemp!�ۼ�ͳ�ﱨ��, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ��֧��, 10) ' = Round(NVL(rsTemp!ȫ�Ը����, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ���Ը�, 10) ' = Round(NVL(rsTemp!�����Ը����, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ��֧��, 10) ' = Round(NVL(rsTemp!����ͳ����, 0), 2)
            strInfor = strInfor & Lpad(dbl����ͳ���Ը�, 10) ' = Round(NVL(rsTemp!ͳ�ﱨ�����, 0), 2)
            strInfor = strInfor & Lpad(dbl��������֧��, 10) ' = Round(NVL(rsTemp!�ۼƽ���ͳ��, 0), 2)
            strInfor = strInfor & Lpad(dbl�ǲ�������֧��, 10) ' = Round(NVL(rsTemp!���Ը����, 0), 2)
            strInfor = strInfor & Lpad(dbl���շ�Χ���Ը�, 10) ' = Round(NVL(rsTemp!�ⶥ��, 0), 2)
        Else
            strInfor = strInfor & String(10, " ")    '���ķ���:���������ʻ����;������:���������ʻ����  NUM 175 10  ���������ʻ������������ʻ�  ���ķ���
            strInfor = strInfor & String(10, " ")    '���ķ���:�����ͳ��֧���ۼ�  NUM 185 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ���ķ���
            '����ǰ�����ʻ��������鿨���ؽ�����������������Ӧ����������ѯ�������ʻ������д��
            strInfor = strInfor & Lpad(.���������ʻ����, 10)  '����ǰ�����ʻ����
            strInfor = strInfor & Lpad(Trim(CStr(.���������ʻ����)), 10)   '����ǰ�����˻����(�����鿨���ؽ�������������������0)
            strInfor = strInfor & Lpad(Trim(CStr(.ͳ���ۼ�)), 10)    '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ��������ʻ�֧��(������������㣬��ʾ�����ʻ�֧��)
            strInfor = strInfor & String(10, " ")    '���ķ���:���β��������ʻ�֧��(������������㷵��0)
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ��֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ���ͳ���Ը�
            strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ��֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���β���ͳ���Ը�
            strInfor = strInfor & String(10, " ")    '���ķ���:���λ�����������֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���ηǻ�����������֧��
            strInfor = strInfor & String(10, " ")    '���ķ���:���α��շ�Χ���Ը�
        End If
        
        If gintInsure <> TYPE_���������� Then        '������
            strInfor = strInfor & Lpad(Trim(CStr(dbl���������Է�)), 10)    '�������������Ը�
        End If
        
        strInfor = strInfor & Lpad(Trim(CStr(dbl�𸶱�׼)), 10)    '�𸶱�׼
        
        strInfor = strInfor & Lpad(.ת�ﵥ��, 6)     'ת�ﵥ��
        strInfor = strInfor & Lpad(Get�������(intҵ��, .�������), 1)     '�������
        If gintInsure <> TYPE_���������� Then
            strInfor = strInfor & Lpad(.�α����3, 1)    '�α����3:0 �󱣡�1 �±��������鿨���
        End If
        
        strInfor = strInfor & Lpad(.ְ����ҽ���, 1)       'ְ����ҽ���
        
        strInfor = strInfor & Lpad(.��ϱ���, 16)    '��ϱ���
        strInfor = strInfor & Lpad(strҽ��, 6)    'ҽʦ����
        strInfor = strInfor & Lpad(UserInfo.���, 6)    '����Ա����
        strInfor = strInfor & Lpad(.�������, 30)    '�������
        'A-������B-��ת��C-δ����D-������E-����
        strInfor = strInfor & Lpad(Get�������_����(lng����id, lng��ҳID), 1)    '���������ʶ
        strInfor = strInfor & Lpad(str��Ժ����, 8)      '��Ժ����
        
        If gintInsure = TYPE_���������� Then        '������
        Else
            strInfor = strInfor & String(16, " ")      '����ʱ��
        End If
        strInfor = strInfor & String(10, " ")      '�������
    End With
    '����1002    12  423 ʵʱ����
    סԺ���㼰����_���� = ҵ������_����(IIf(gintInsure = TYPE_����������, 2, 1), 1002, strInfor)
    
    
    '��������¼
 
   
    '������:
    '   ���������ʻ����  NUM 175 10  ���������ʻ������������ʻ�  ���ķ���
    '   �����ͳ��֧���ۼ�  NUM 185 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ���ķ���
    
    '    ���λ��������ʻ�֧��    NUM 225 10      ���ķ���
    '    ���β��������ʻ�֧��    NUM 235 10      ���ķ���
    '    ���λ���ͳ��֧��    NUM 245 10      ���ķ���
    '    ���λ���ͳ���Ը�    NUM 255 10      ���ķ���
    '    ���β���ͳ��֧��    NUM 265 10      ���ķ���
    '    ���β���ͳ���Ը�    NUM 275 10      ���ķ���
    '    ���λ�����������֧��    NUM 285 10  ����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
    '    ���ηǻ�����������֧��  NUM 295 10  ����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
    '    ���α��շ�Χ���Ը�  NUM 305 10  �޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
    '������:
    '   ���������ʻ����  NUM 161 10  ��  ����ǻ���ҽ�ƽ����ʾ�����������ʻ������������ʻ��� ��������������ʾ: �����ʻ��������� ����
    '   �����ͳ��֧���ۼ�  NUM 171 10  ����ͳ���ۼƣ�����ͳ���ۼ�  ����
    
    '    ���λ��������ʻ�֧��    NUM 211 10  ������������㣬��ʾ�����ʻ�֧��    ����
    '    ���β��������ʻ�֧��    NUM 221 10  ������������㷵��0 ����
    '    ���λ���ͳ��֧��    NUM 231 10      ����
    '    ���λ���ͳ���Ը�    NUM 241 10      ����
    '    ���β���ͳ��֧��    NUM 251 10  ������������㣬���ֶ����ڴ����������֧��  ����
    '    ���β���ͳ���Ը�    NUM 261 10      ����
    '    ���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧�� 2�� ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
    '    ���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��   2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����    ����
    '    ���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    
    Dim i As Long
    If gintInsure = TYPE_���������� Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    
    dbl�����ʻ���� = Val(Substr(strInfor, i - 40, 10))
    dblͳ��֧���ۼ� = Val(Substr(strInfor, i - 30, 10))  '�����ͳ��֧���ۼ�=����ͳ���ۼƣ�����ͳ���ۼ�
    
    dbl�����ʻ�֧�� = Val(Substr(strInfor, i + 10, 10)) '���λ��������ʻ�֧��=������������㣬��ʾ�����ʻ�֧��
    dbl�����ʻ�֧�� = Val(Substr(strInfor, i + 20, 10))    '���β��������ʻ�֧��    NUM 221 10  ������������㷵��0
    dbl����ͳ��֧�� = Val(Substr(strInfor, i + 30, 10))   '���λ���ͳ��֧��    NUM 231 10      ����
    dbl����ͳ���Ը� = Val(Substr(strInfor, i + 40, 10))     '���λ���ͳ���Ը�    NUM 241 10      ����
    dbl����ͳ��֧�� = Val(Substr(strInfor, i + 50, 10))     '���β���ͳ��֧��    NUM 251 10      ����
    dbl����ͳ���Ը� = Val(Substr(strInfor, i + 60, 10))     '���β���ͳ���Ը�    NUM 261 10      ����
    dbl��������֧�� = Val(Substr(strInfor, i + 70, 10))     '���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧��2��   ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
    dbl�ǲ�������֧�� = Val(Substr(strInfor, i + 80, 10))     '���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����
    dbl���շ�Χ���Ը� = Val(Substr(strInfor, i + 90, 10))     '���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    
    '/???
       'ԭ���̲���:
       '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
       "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
       '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
       '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
       '    ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN,
       '    ����_IN,��ҩ��_IN,��ҩ��_IN,��ҩ��_IN,����_IN,���Ʒ�_IN,����_IN,����Է�_IN,�������Ʒ�_IN,���������Է�_IN,�������Էѷ���_IN,�Ǳ��շ���_IN,ͳ�����_IN,������
        '   ����ǰ�����ʻ����,����ǰ�����˻����,����ǰͳ���ۼ�
       '������ֵ����Ϊ:
       '       ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN, _
       '       dbl�����ʻ����,dblͳ��֧���ۼ�,dbl��������֧��,dbl�����ʻ�֧��,סԺ����_IN,����_IN,dbl���շ�Χ���Ը�,ʵ������_IN
       '       �������ý��_IN,dbl����ͳ��֧��,dbl����ͳ���Ը�,
       '       dbl����ͳ��֧��,dbl����ͳ���Ը�,dbl�ǲ�������֧��,dbl�����ʻ�֧��
       '       ֧��˳���_IN(�������;ת�ﵥ��;��ϱ���),��ҳID_IN,��;����_IN,�������_IN
       '    ����_IN,��ҩ��_IN,��ҩ��_IN,��ҩ��_IN,����_IN,���Ʒ�_IN,����_IN,����Է�_IN,�������Ʒ�_IN,���������Է�_IN,�������Էѷ���_IN,�Ǳ��շ���_IN,ͳ�����_IN,������
        '   ����ǰ�����ʻ����,����ǰ�����˻����,����ǰͳ���ۼ�
              
       gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����id & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          dbl�����ʻ���� & "," & dblͳ��֧���ۼ� & "," & dbl��������֧�� & "," & dbl�����ʻ�֧�� & "," & "Null" & "," & dbl�𸶱�׼ & "," & dbl���շ�Χ���Ը� & "," & dbl�𸶱�׼ & "," & _
          curTotal & "," & dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & _
          dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & dbl�ǲ�������֧�� & ",Null," & dbl�����ʻ�֧�� & ",'" & _
          Get�������(intҵ��, g�������_����.�������) & ";" & g�������_����.ת�ﵥ�� & ";" & g�������_����.��ϱ��� & "'," & lng��ҳID & ",null,'" & g�������_����.������� & "'," & _
           dbl���� & "," & dbl��ҩ�� & "," & dbl��ҩ�� & "," & dbl��ҩ�� & "," & dbl���� & "," & dbl���Ʒ� & "," & dbl���� & "," & dbl����Է� & "," & dbl�������Ʒ� & "," & dbl���������Է� & "," & dbl�������Էѷ��� & "," & dbl�Ǳ��շ��� & "," & dbl���� & "," & dbl������ & "," & _
            dbl����ǰ���(1) & "," & dbl����ǰ���(2) & "," & dbl����ǰ���(3) & _
            " )"
            
        zlDatabase.ExecuteProcedure gstrSQL, "����סԺ�����շ�����"
        Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function סԺ����_����(lng����ID As Long, ByVal lng����id As Long) As Boolean

    Dim cur�����ʻ� As Currency
    Dim lng��ҳID As Long
    Dim blnError As Boolean
    Dim str��Ժ��� As String, str������� As String
    Dim str����ʱ�� As String, str����ʱ�� As String
    Dim str������ As String
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    On Error GoTo ErrHand
    
    Call �������_����(lng����id)

    cur�����ʻ� = g�������_����.���������ʻ����

    gstrSQL = " Select B.סԺ���� ��ҳID,to_char(A.��Ժ����,'yyyy') ��Ժ��� " & _
              " From ������ҳ A,������Ϣ B" & _
              " Where B.����ID=" & lng����id & " And A.��ҳID=B.סԺ���� And A.����ID=B.����ID"
    Call OpenRecordset(rsTemp, "��ȡ������Ժʱ��")
    str��Ժ��� = rsTemp!��Ժ���
    lng��ҳID = rsTemp!��ҳID

    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str����ʱ�� = str����ʱ��
    str������� = Mid(str����ʱ��, 1, 4)

    סԺ����_���� = סԺ���㼰����_����(False, lng����id, lng����ID, lng����ID, lng��ҳID)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function Get����(ByVal strְ����ҽ��� As String, ByVal lng���� As Long) As Double
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------

       Dim strCaption As String
       Dim rsTmp As New ADODB.Recordset
       strCaption = Decode(strְ����ҽ���, "A", "��ְ", "B", "����", "L", "����", "T", "����", "Q", "��ҵ����", "��ְ")
    
        gstrSQL = "" & _
            "   Select d.���*a.����/100 as ����" & _
            "   From ����֧������ a,������Ⱥ b, " & _
            "      (Select * From ���������  " & _
            "       where ((" & lng���� & ">=���� and " & lng���� & "<=����) or (" & lng���� & ">���� and ����=0) ) and ����=" & gintInsure & _
            "       ) c,����֧���޶� d " & _
            " where a.����=" & gintInsure & " and b.���� =a.���� and a.��ְ=b.��� and b.����='" & strCaption & "' and  " & _
            "       a.�����=c.����� and a.��ְ=c.��ְ and a.����=d.���� and d.���='" & Format(zlDatabase.Currentdate, "yyyy") & "' and d.����='1'"
    
       Err = 0
       On Error GoTo ErrHand:
       zlDatabase.OpenRecordset rsTmp, gstrSQL, "��������"
       If Not rsTmp.EOF Then
            Get���� = NVL(rsTmp!����, 0)
       Else
            Get���� = 0
       End If
       Exit Function
ErrHand:
        If ErrCenter = 1 Then
            Resume
        End If
       Get���� = 0
   
End Function
Public Function סԺ�������_����(lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim str�˵���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng����id As Long
    Dim lng��ҳID As Long
    
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    On Error GoTo ErrHand
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    lng����ID = rsTemp("ID") '�������ݵ�ID

    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID='" & lng����ID & "'"
    Call OpenRecordset(rsTemp, "����ҽ��")
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ��ս����¼���޸ý����¼!"
        Exit Function
    End If
    lng����id = NVL(rsTemp!����ID, 0)
    lng��ҳID = NVL(rsTemp!��ҳID, 0)
        
        
    '���¶���
    If ��ȡ�������_����(IIf(gintInsure = TYPE_����������, 2, 1)) = False Then
        Exit Function
    End If
    
    Dim strArr
    strArr = Split(NVL(rsTemp!֧��˳���), ";")
    
    '�������;ת�ﵥ��;��ϱ���
    '5-��ͨסԺ("2", "D"),6-��ͥ����סԺ("4", "C")
    '7-��������סԺ("O", "P"),8-���˱���סԺ("Q", "R")
    With rsTemp
        If UBound(strArr) >= 2 Then
            g�������_����.������� = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g�������_����.ת�ﵥ�� = strArr(1)
            g�������_����.��ϱ��� = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g�������_����.������� = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g�������_����.ת�ﵥ�� = strArr(1)
        Else
            g�������_����.������� = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
        End If
        g�������_����.������� = NVL(rsTemp!��ע)
    End With
    
    '��֤�Ƿ�Ϊ�ò��˵�IC��
    gstrSQL = "Select * From  �����ʻ� where ����id=" & lng����id
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵�ҽ����"
    If rsTemp.EOF Then
        ShowMsgbox "�ò����ڱ����ʻ����޼�¼!"
        Exit Function
    End If
    
    If g�������_����.IC���� <> NVL(rsTemp!����) Then
        ShowMsgbox "�ò��˵�IC���������,�����ǲ����������˵�IC��!"
        Exit Function
    End If
    '���ó�������ӿ�
    סԺ�������_���� = סԺ���㼰����_����(True, lng����id, lng����ID, lng����ID, lng��ҳID)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ҽ����ֹ_����() As Boolean
    ҽ����ֹ_���� = True
End Function

Public Function �����Ǽ�_����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��д�뵥��ͷ����д�뵥����
    '��¼״̬��1-����;����Ϊɾ���������ô�����ֻ�����ŵ���ɾ�����ٲ����µ���
    On Error GoTo ErrHand
    �����Ǽ�_���� = False
    If gblnסԺ��ϸʱʵ�ϴ� = False Then
        �����Ǽ�_���� = True
        Exit Function
    End If
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient

        gstrSQL = " " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ���� ,C.�Ƿ�ҽ��,C.ͳ��ȶ�,F.סԺ���� AS ��ҳid, " & _
            "        G.��ʶ����||G.��ʶ���� AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From ���˷��ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.��Ŀ����,M.�Ƿ�ҽ��,M.�շ�ϸĿid,Q.ͳ��ȶ�  From ����֧����Ŀ M,����֧������ Q Where M.����=" & TYPE_������ & " and M.����ID=Q.id) C " & _
            " Where     a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=82 AND F.סԺ����= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_������ & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=" & lng��¼���� & " and  A.��¼״̬=" & lng��¼״̬ & " And A.NO='" & str���ݺ� & "'" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 "
        
        gstrSQL = gstrSQL & " Union all " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ���� ,C.�Ƿ�ҽ��,C.ͳ��ȶ�,F.סԺ���� AS ��ҳid, " & _
            "        G.��ʶ����||G.��ʶ���� AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From ���˷��ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.��Ŀ����,M.�Ƿ�ҽ��,M.�շ�ϸĿid,Q.ͳ��ȶ�  From ����֧����Ŀ M,����֧������ Q Where M.����=" & TYPE_���������� & " and M.����ID=Q.id) C " & _
            " Where     a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=83 AND F.סԺ����= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_���������� & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=" & lng��¼���� & " and  A.��¼״̬=" & lng��¼״̬ & " And A.NO='" & str���ݺ� & "'" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
            " Order by ����ID"
            
        Call OpenRecordset(rsTemp, "�����Ǽ�")
        
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�������¼����ҽ����������������ʧ�ܣ�[�����Ǽ�]", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    �����Ǽ�_���� = �ϴ�����_����(rsTemp)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function �ϴ�����_����(ByVal rsExse As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:rsExse-��ϸ����
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------


    Dim lng����id As Long
    Dim curTotal As Currency
    Dim blnUpload As Boolean
    Dim rsPara As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str��ϸ As String
    Dim str��Ŀͳ�Ʒ��� As String
    Dim strTmp As String
    Err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "select * from ���ղ��� where ���� in (82,83)"
    zlDatabase.OpenRecordset rsPara, gstrSQL, "�ϴ�������ȡ����"
    With rsExse
        Do While Not .EOF
            lng����id = NVL(!����ID, 0)
            'ȷ���������
            '�ϴ���ϸ��¼,ʵʱҽ����ϸ����
                
            If NVL(!����, 0) = TYPE_���������� Then '������
                str��ϸ = Lpad(gstrҽԺ����_����, 6)     'ҽԺ����    CHAR    1   6       Ժ����д
                str��ϸ = str��ϸ & Lpad(NVL(!ҽ����), 10)  '���ձ��    CHAR    7   10      Ժ����д
            Else
                str��ϸ = Lpad(gstrҽԺ����_����, 4)     'ҽԺ����    CHAR    1   4       Ժ��
                str��ϸ = str��ϸ & Lpad(NVL(!ҽ����), 8)   '���˱��    CHAR    5   8       Ժ��
            End If
            
            str��ϸ = str��ϸ & Lpad(NVL(!סԺ��, 0), 10) '��־��  CHAR    13  10  ������ϸ�Կո�λ,סԺ��סԺ��  Ժ��
            str��ϸ = str��ϸ & Lpad(NVL(!˳���, 0), 4)   '�������    NUM 23  4   סԺ��ϸ�����������Ժ�Ǽ�ʱ�������������ϸ:                         ������ڱ��ν���������� Ժ��
            str��ϸ = str��ϸ & Lpad(NVL(!NO, 0), 10)       '������  NUM 27  10      Ժ��
            
            If NVL(!����, 0) = TYPE_���������� Then '������
            Else
                str��ϸ = str��ϸ & Lpad(CStr(NVL(!���, 0)), 10)      '������Ŀ���    NUM 37  10  ��Ӧ�����ŵļǼ���Ŀ���    Ժ��
            End If
            
            '������Ϊ���ݺ�  CHAR    41  10  ҽ���ţ�    Ժ����д
            str��ϸ = str��ϸ & Lpad(NVL(!ҽ�����, 0), 10)     'ҽ����  CHAR    47  10  ������Ӧҽ����ҽ����¼�ţ�������ϸ��û��ҽ����ҽԺ�Կո�λ    Ժ��
            
            str��ϸ = str��ϸ & Get�������(0, NVL(!�Ҷȼ�))      '�������    CHAR    57  1   ȡֵ���"�������"˵��  Ժ��
            
            If NVL(!����, 0) = TYPE_���������� Then '������
                '������Ϊ����ʱ��    DATETIME    52  16  ��ȷ���루������ʱ�䣩��ʽΪ��yyyymmddhhmiss�����Կո�λ  Ժ����д
                str��ϸ = str��ϸ & Rpad(NVL(!����ʱ��), 16)
            Else
                str��ϸ = str��ϸ & Rpad(NVL(!�Ǽ�ʱ��), 16)      '��������ʱ�䣨Ͷҩʱ�䣩    DATETIME    58  16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ    Ժ��
            End If
            
            str��ϸ = str��ϸ & Lpad(NVL(!���ұ���), 20)      '��Ŀ����    CHAR    74  20  �Ƽ���Ŀ����    Ժ��
            str��ϸ = str��ϸ & Lpad(NVL(!��Ŀ����), 20)      '��Ŀ����    CHAR    94  20      Ժ��

            If NVL(!����, 0) = TYPE_���������� Then '������
            Else

                If !�Ƿ�ҽ�� = 1 Then
                    str��ϸ = str��ϸ & Lpad(1 - NVL(!ͳ��ȶ�, 0), 6)   '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                Else
                    str��ϸ = str��ϸ & Lpad(1, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                End If
                rsPara.Filter = 0
                rsPara.Filter = " ������='" & NVL(!���) & "' and ����=" & NVL(!����, 0)
                str��Ŀͳ�Ʒ��� = ""
                If Not rsPara.EOF Then
                    strTmp = NVL(rsPara!����ֵ)
                    If InStr(1, strTmp, ";") <> 0 And strTmp <> ";" Then
                        strTmp = Split(strTmp, ";")(1)
                        If strTmp <> "" Then
                            str��Ŀͳ�Ʒ��� = Substr(strTmp, 1, 1)
                            str��ϸ = str��ϸ & Substr(strTmp, 1, 1)   '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                        Else
                            str��ϸ = str��ϸ & Space(1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                        End If
                    Else
                        str��ϸ = str��ϸ & Space(1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    End If
                Else
                        str��ϸ = str��ϸ & Space(1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                End If
            End If
            
            str��ϸ = str��ϸ & Lpad(NVL(!����), 6)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
            str��ϸ = str��ϸ & Lpad(NVL(!ʵ�ʼ۸�), 8) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
            str��ϸ = str��ϸ & Lpad(NVL(!���㵥λ), 4) '��λ    CHAR    135 4       Ժ��
            str��ϸ = str��ϸ & Lpad(NVL(!����), 20)      '����    CHAR    139 20  �����Ƭ����    Ժ��
            
            If NVL(!����, 0) = TYPE_���������� Then  '������
                '��ȡ���˵�����.
                gstrSQL = "Select ����,Ƶ��,�÷� From ҩƷ�շ���¼ where ����id=" & NVL(!ID, 0)
                zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵���Ƶ��"
                If rsTemp.EOF Then
                    str��ϸ = str��ϸ & Space(5)       'ÿ������    NUM 146 5       Ժ����д
                    str��ϸ = str��ϸ & Space(20)      'ʹ��Ƶ��    CHAR    151 20  �磺1��2��  Ժ����д
                    str��ϸ = str��ϸ & Space(50)      '�÷�    CHAR    171 50  �磺�ڷ�    Ժ����д
                Else
                    str��ϸ = str��ϸ & Lpad(NVL(rsTemp!����), 5)      'ÿ������    NUM 146 5       Ժ����д
                    str��ϸ = str��ϸ & Lpad(NVL(rsTemp!Ƶ��), 20)      'ʹ��Ƶ��    CHAR    151 20  �磺1��2��  Ժ����д
                    str��ϸ = str��ϸ & Lpad(NVL(rsTemp!�÷�), 50)      '�÷�    CHAR    171 50  �磺�ڷ�    Ժ����д
                End If
                str��ϸ = str��ϸ & Space(4)      'ִ������    NUM 221 4       Ժ����д
                str��ϸ = str��ϸ & Lpad(NVL(!ҽ�����), 6)      'ҽʦ����    CHAR    225 6       Ժ����д
            Else
                str��ϸ = str��ϸ & Lpad(NVL(!ҽ��), 8)      'ҽʦ����    CHAR    159 8       Ժ��
            End If
            'ȷ��������
            
            strTmp = Get��Ժ���(NVL(!����ID), NVL(!��ҳID, 0), False, True)
            If InStr(1, strTmp, "|") <> 0 Then
                
                str��ϸ = str��ϸ & Lpad(Split(strTmp, "|")(1), 16)     '��ϱ���    CHAR    167 16      Ժ��
                strTmp = Split(strTmp, "|")(0)
                strTmp = Lpad(strTmp, 30)
                strTmp = Substr(strTmp, 1, 30)
                str��ϸ = str��ϸ & strTmp     '�������    CHAR    183 30      Ժ��
            Else
                str��ϸ = str��ϸ & Space(16)      '��ϱ���    CHAR    167 16      Ժ��
                str��ϸ = str��ϸ & Space(30)     '�������    CHAR    183 30      Ժ��
            End If
            
            If NVL(!����, 0) = TYPE_���������� Then  '������
                str��ϸ = str��ϸ & Lpad(NVL(!��������), 20)    '�Ʊ�����    CHAR    277 20      Ժ����д
            Else
                str��ϸ = str��ϸ & Space(16)     '����ʱ��    DATETIME    213 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ��Ժ�˿ո�λ  ����
            End If
            
            '�ϴ���ϸ
            '1003    7   230 ʵʱҽ����ϸ�����ύ
            �ϴ�����_���� = ҵ������_����(IIf(NVL(!����, 0) = TYPE_����������, 2, 1), 1003, str��ϸ)
            If �ϴ�����_���� = False Then
                ShowMsgbox "�������ʱҽ����ϸ�����ύʧ��,���ܼ���!"
                Exit Function
            End If
            '�ϴ�ҽ����ϸ
            If NVL(!ҽ�����, 0) <> 0 Then
                �ϴ�����_���� = False
                If ҽ����ϸ�����ύ(NVL(!ҽ�����, 0), NVL(!סԺ��), str��Ŀͳ�Ʒ���) = False Then
                    ShowMsgbox "ҽ����ϸ�����ύʧ��,���ܼ���!"
                    Exit Function
                End If
            End If

            'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & NVL(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
            zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            .MoveNext
        Loop
    End With
    �ϴ�����_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function Decode(ParamArray arrPar() As Variant) As Variant

'���ܣ�ģ��Oracle��Decode����

    Dim varValue As Variant, i As Integer

    

    i = 1

    varValue = arrPar(0)

    Do While i <= UBound(arrPar)

        If i = UBound(arrPar) Then

            Decode = arrPar(i): Exit Function

        ElseIf varValue = arrPar(i) Then

            Decode = arrPar(i + 1): Exit Function

        Else

            i = i + 2

        End If

    Loop

End Function

Private Function ҽ����ϸ�����ύ(ByVal lngҽ��ID As Long, ByVal strסԺ�� As String, ByVal str��Ŀͳ�Ʒ��� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡҽ����ϸ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    '5.  ʵʱҽ�������ύ�ӿ�

    '��������ҽ���ӿ�
    If gintInsure = TYPE_���������� Then
        ҽ����ϸ�����ύ = True
        Exit Function
    End If
    gstrSQL = " " & _
         " select ID,����� as �����,decode(��Ч,1,1,0) as  ҽ������,ҩƷ����,������λ,ִ��Ƶ��,Ƶ�ʴ���,ҽ������, " & _
         "        ��ҽ��ҽ��,��ҽ��ʱ��,to_char(��ʼִ��ʱ��,'yyyymmddhh24miss') as ��ʼִ��ʱ��,У�Ի�ʿ as ִ��ҽ����ʿ����," & _
         "        ͣҽ��ҽ��,to_char(ͣҽ��ʱ��,'yyyymmddhh24miss') as ͣҽ��ʱ��,����˵��" & _
         " from ҽ����¼  " & _
         " Where id=" & lngҽ��ID
    Err = 0
    On Error GoTo ErrHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ����ϸ��¼"
    If rsTemp.EOF Then
        ShowMsgbox "�޶�Ӧ��ҽ����¼!"
        Exit Function
    End If
    With g�������_����
        strInfor = Lpad(gstrҽԺ����_����, 4)             'ҽԺ����    CHAR    1   4       Ժ��
        strInfor = strInfor & Lpad(.���˱��, 8)    '���˱��    CHAR    5   8       Ժ��
        strInfor = strInfor & Lpad(.�������, 4)    '�������    NUM 13  4   ���������Ժʱ�������  Ժ��
        strInfor = strInfor & Lpad(strסԺ��, 10)    '��־��  CHAR    17  10      Ժ��
        strInfor = strInfor & Lpad(lngҽ��ID, 10)     'ҽ����  CHAR    27  10      Ժ��
        
        'ע�ڣ�ҽ�������������ʾҽ����ͬʱʹ�õ���Ŀ�����磺������ҽ����¼�ֱ�Ϊ��ù�غ��Ȼ���ע��Һ��ҽ��Ҫ������ҩ��ͬʱ������ʹ�ã���ʱ�Ϳ��Խ�������¼�ķ������Ϊ��ͬ��ֵ��ֵ�����ݲ�ͬҽԺ���Ը�������ҽ��ϵͳ�ľ�������Զ���ֻҪ����2λ�ַ���ʶ���Ըû���ͬʱʹ�õļǼ���Ŀ���ɡ�
        strInfor = strInfor & Lpad(lngҽ��ID, 10)     'ҽ�������  CHAR    37  3   ���ע��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ������, 0), 1)   'ҽ������    CHAR    40  1   1 ������0 ��ʱҽ��
        strInfor = strInfor & Space(20)   '��Ŀ����    CHAR    41  20  �Ƽ���Ŀ���룬����������ҽ�����磺���ճ�Ժ�� ��Ŀ����ͳһ��'000000' Ժ��
        strInfor = strInfor & Space(20)   '��Ŀ����    CHAR    61  20      Ժ��
        strInfor = strInfor & Lpad(str��Ŀͳ�Ʒ���, 1)  '��Ŀͳ�Ʒ���    CHAR    81  1   ���ע��    Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҩƷ����, 0), 15) 'ÿ������    CHAR    82  15  ���磺10    Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!������λ), 4) '������λ    CHAR    97  4   ���磺ml
        strInfor = strInfor & Lpad(NVL(rsTemp!ִ��Ƶ��), 20) 'ʹ��Ƶ��    CHAR    101 20  �磺1��2��  Ժ��
        strInfor = strInfor & Substr(Lpad(NVL(rsTemp!ҽ������), 50), 1, 50) '�÷�    CHAR    121 50  �磺�ڷ���������ע��20��/���ӣ��������� Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!��ҽ��ҽ��), 8) '��ҽ��ҽʦ����  CHAR    171 8       Ժ��
        strInfor = strInfor & Rpad(NVL(rsTemp!��ʼִ��ʱ��), 16) '��ʼִ��ʱ��    DATATIME    179 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ���������  Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ִ��ҽ����ʿ����), 8) 'ִ��ҽ����ʿ����    CHAR    195 8       Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ͣҽ��ҽ��), 8) '��ֹҽ��ҽʦ����    CHAR    203 8       Ժ��
        strInfor = strInfor & Rpad(NVL(rsTemp!ͣҽ��ʱ��), 16) '��ֹҽ��ʱ��    DATATIME    211 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڳ���ҽ��������������ʱҽ��  Ժ��
        strInfor = strInfor & Substr(Lpad(NVL(rsTemp!����˵��), 30), 1, 30) '��ע    CHAR    227 30  ������ʱҽ��ִ�з���������������    Ժ��
        strInfor = strInfor & Space(16)                  '����ʱ��    DATATIME    257 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  ����
    End With
    '1005    8   274 ʵʱҽ������
    ҽ����ϸ�����ύ = ҵ������_����(g�������_����.ҽ������, 1005, strInfor)
    Exit Function
ErrHand:
    '���ûװҽ���Ͳ�ִ��
    ҽ����ϸ�����ύ = True
End Function

Private Function Get���˱䶯��¼(ByVal lng����id As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵ı䶯���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
        "   Select  ����,���Ӵ�λ,��ʼʱ��,��ֹʱ��,��λ�ȼ�id " & _
        "   From ���˱䶯��¼  " & _
        "   Where  ����id=" & lng����id & " and ��ҳid=" & lng��ҳID & " and ���� is not null"
    Err = 0
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˱䶯���"
    Set Get���˱䶯��¼ = rsTemp
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Set Get���˱䶯��¼ = Nothing
    Exit Function
End Function
Private Function GetסԺ�����¼(ByVal lng����id As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��������δ���¼
    '--�����:
    '--������:
    '--��  ��:δ�����
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset

    '
    strSql = _
        "   Select  A.��¼����,A.��¼״̬,A.NO,A.���,A.����," & _
        "           A.����ID,A.��ҳID,A.Ӥ����," & _
        "           A.���մ���ID,A.�շ����,A.�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������," & _
        "           Decode(Sign(Instr(B.���,'��')),0,B.���,Substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
        "           Decode(Sign(Instr(B.���,'��')),0,B.���,Substr(B.���,Instr(B.���,'��')+1)) as ����," & _
        "           A.����,Decode(A.����,0,0,Round(A.���/A.����,4)) as �۸�,A.���,A.ҽ��,w.��� as ҽ�����,A.�Ǽ�ʱ��," & _
        "           A.�Ƿ��ϴ�,A.�Ƿ���,A.������Ŀ��,A.ժҪ,C.��Ŀ���� as ҽ����Ŀ����," & _
        "           C.��Ŀ���� as ҽ����Ŀ����,Q.����ֵ,Q.������,J.ͳ��ȶ�,J.סԺ�ȶ�,J.��׼����,J.�㷨" & _
        "   From (" & _
        "           Select  Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.����,A.NO,Nvl(A.�۸񸸺�,���) as ���,A.����ID,A.��ҳID,Nvl(A.Ӥ����,0) as Ӥ����," & _
        "                   A.������ as ҽ��,A.��������ID,A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0) as ���մ���ID,Avg(Nvl(A.����,1)*A.����) as ����," & _
        "                   Sum(A.��׼����) as ��׼����,Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as ���,A.�Ǽ�ʱ��,Nvl(A.�Ƿ��ϴ�,0) as �Ƿ��ϴ�,Nvl(A.�Ƿ���,0) as �Ƿ���,Nvl(A.������Ŀ��,0) as ������Ŀ��,A.ժҪ" & _
        "           From ���˷��ü�¼ A,������Ŀ B" & _
        "           Where A.���ʷ���=1 And A.������ĿID=B.ID And A.����ID=" & lng����id & _
        "           Group by    Mod(A.��¼����,10),A.��¼״̬,A.NO,Nvl(A.�۸񸸺�,���),A.����ID,A.��ҳID,A.����,Nvl(A.Ӥ����,0),A.������," & _
        "                       A.��������ID,A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0),A.�Ǽ�ʱ��,Nvl(A.�Ƿ��ϴ�,0),Nvl(A.�Ƿ���,0),Nvl(A.������Ŀ��,0),A.ժҪ" & _
        "           Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0) A,�շ�ϸĿ B,���ű� X," & _
        "           (Select * From ����֧����Ŀ Where ����=" & gintInsure & ") C," & _
        "           (Select M.����, L.������,L.����ֵ from �շ���� M,���ղ��� L  Where M.���=L.������ and L.����=" & gintInsure & ")  Q," & _
        "           (Select * from ����֧������  Where ����=" & gintInsure & ")  J,��Ա�� W" & _
        "   Where     A.�շ�ϸĿID=B.ID and a.ҽ��=w.����(+) and C.����id=J.ID and a.�շ����=Q.����(+) And A.�շ�ϸĿID=C.�շ�ϸĿID And A.��������ID=X.ID"
    Err = 0
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsTmp, strSql, "��ȡ����ҽ��δ�����"
    Set GetסԺ�����¼ = rsTmp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Set GetסԺ�����¼ = Nothing
    Exit Function
End Function


Private Function Set����ҺŽ�������(ByVal bln���� As Boolean, lng����ID As Long, cur�����ʻ� As Currency, lng����id As Long, strSelfNo As String) As Boolean
  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID��
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Set����ҺŽ������� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �ҺŽ���_����(ByVal lng����ID As Long) As Boolean
     �ҺŽ���_���� = Set����ҺŽ�������(False, lng����ID, 0, 0, 0)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function �Һų���_����(ByVal lng����ID As Long) As Boolean
    �Һų���_���� = Set����ҺŽ�������(False, lng����ID, 0, 0, 0)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function ��Ժ������Ϣ_����(lng����id As Long, lng��ҳID As Long) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo ErrHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select * From �����ʻ� where  ����=" & gintInsure & "  and ����id=" & lng����id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    strת�ﵥ�� = NVL(rsTemp!��Ա���)
    lng���� = IIf(gintInsure = 83, 2, 1)
    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6) 'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 10)     '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4) 'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(NVL(rsTemp!ҽ����), 8)     '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    
    strInfor = strInfor & Lpad(NVL(rsTemp!˳���, 1), 4)      '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    
    str������� = Decode(NVL(rsTemp!�Ҷȼ�, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=" & lng����id & _
            "       and A.����ID=" & lng����id & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID"
            
    Call OpenRecordset(rsTemp, "��ȡ��Ժ��Ϣ")
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = NVL(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!סԺ��, 0), 10)      '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(NVL(rsTemp!��Ժ����), 8)      '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(NVL(rsTemp!��Ժ����ʱ��), 16)     '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    If lng���� = 2 Then
        '������Ϊ:סԺ 2���Ҵ� 4ȡ��סԺ�Ǽ� C
        strInfor = strInfor & IIf(str������� = "4", "4", "2")
    Else
        strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    End If

    gstrSQL = "Select * From ��λ״����¼ D where ����ID=" & NVL(rsTemp!��ǰ����ID, 0) & " And ����=" & NVL(rsTemp!��ǰ����, 0)
    Call OpenRecordset(rsTemp, "��ȡ��λ��Ϣ")
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(NVL(rsTemp!�����)) & "��" & Trim(NVL(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = NVL(rsTemp!��Ժ���)
        strȷ��������� = NVL(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
    If lng���� = 2 Then
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
    Else
        strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
        strInfor = strInfor & Lpad(str��Ժ�������, 30)  '��Ժ�������    CHAR    68  30      y Ժ��
        strInfor = strInfor & Lpad(strȷ����ϱ���, 16)  'ȷ����ϱ���    CHAR    98  16      N   Ժ��
        strInfor = strInfor & Lpad(strȷ���������, 30)  'ȷ���������    CHAR    114 30      N   Ժ��
    End If
    strInfor = strInfor & Lpad(str��Ժ����, 20)  '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    If lng���� = 2 Then
    Else
        strInfor = strInfor & str��λ��              '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    End If
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)   'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)   '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    If lng���� = 2 Then
    Else
        strInfor = strInfor & "M"   '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
        strInfor = strInfor & Space(16)   '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    End If
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    ��Ժ������Ϣ_���� = ҵ������_����(lng����, 1004, strInfor)
    If ��Ժ������Ϣ_���� = False Then
        ShowMsgbox "ʵʱסԺ�Ǽ������ύʧ��!"
        Exit Function
    End If
    ��Ժ������Ϣ_���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function



Public Function GetItemInfo_����(ByVal lngPatiID As Long, ByVal lngItemID As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�������˵������ʾ��Ϣ
    '--�����:
    '--������:
    '--��  ��:��ʾ��
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strҽ�Ƹ��ʽ As String
    Dim int���� As Integer
    Dim bln��Ժ As Boolean
    Dim dblͳ����� As Double
    Dim strMsgInfor As String
    
    '��һ��:ȷ���Ƿ�ҽ������
    gstrSQL = "Select ����id,����,nvl(��ǰ״̬,0) as ״̬ from �����ʻ�  where ����id=" & lngPatiID & " and ����=" & gintInsure
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�ж��Ƿ�Ϊҽ������!"
    If rsTemp.EOF Then
        rsTemp.Close
        GetItemInfo_���� = ""
        Exit Function
    End If
    
    int���� = NVL(rsTemp!����, 0)
    bln��Ժ = NVL(rsTemp!״̬, 0) > 0
    '�ڶ���:ȷ��ҽ�Ƹ��ʽ
    gstrSQL = "Select ҽ�Ƹ��ʽ from ������Ϣ where ����id=" & lngPatiID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ�Ƹ���ʽ"
    strҽ�Ƹ��ʽ = NVL(rsTemp!ҽ�Ƹ��ʽ)
        
    '��������ȷ���շ�ϸĿ���������
    gstrSQL = "" & _
        "   Select b.����,b.����,b.����,b.�㷨,a.��Ŀ����,b.ͳ��ȶ�,b.��׼����,b.סԺ�ȶ�,a.�Ƿ�ҽ�� " & _
        "   From ����֧����Ŀ a,����֧������ b " & _
        "   where a.����id=b.id and a.����=b.���� and a.�շ�ϸĿid=" & lngItemID & " and a.����=" & int����
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����֧������"
    strMsgInfor = ""
    If InStr(1, "������ҽ�Ʊ���;��ҵ����;���˱���;��������;��ҵ����;", IIf(strҽ�Ƹ��ʽ = "", "D", strҽ�Ƹ��ʽ) & ";") <> 0 Then
        '   ҽ�Ƹ��ʽΪ������ҽ�������ҽ������ҵ���ݡ����˱��ա��������ա���ҵ���յģ������������մ�����ҽ���ӿڵ�ҽ����Ŀ�����е�ҽ�����ඨ���еı�������������ʾ
        '   ҽ�Ƹ��ʽΪ������ҽ���ģ������������տ�����ҽ���ӿڵ�ҽ�����ඨ���еı�������������ʾ
        If bln��Ժ Then
            If NVL(rsTemp!�㷨, 0) = 2 Then
                 '���˺�:200404,���õ��㷨2(���ö����),������б�������
                 strMsgInfor = "����Ŀ�̶�����:" & Format(rsTemp!��׼����, "#####0.00;-####0.00; ;") & "Ԫ"
            Else
                 strMsgInfor = "����Ŀ��������:" & Format(rsTemp!סԺ�ȶ�, "#####0.00;-####0.00; ;") & "%"
            End If
        Else
                 strMsgInfor = "����Ŀ��������:" & Format(rsTemp!ͳ��ȶ�, "#####0.00;-####0.00; ;") & "%"
        End If
    ElseIf InStr(1, "����ҽ��;��ͬ��λ", "") <> 0 Then
        '   ҽ�Ƹ��ʽΪ����ҽ�ơ���ͬ��λ�ģ������������մ�����ҽ���ӿڵ�ҽ����Ŀ�����е���ҵ���ѱ������������ʾ��
        strMsgInfor = "����Ŀ���ѱ���:" & Format(Val(NVL(rsTemp!��Ŀ����)), "#####0.00;-####0.00; ;") & "%"
    End If
    If strMsgInfor <> "" Then
        ShowMsgbox strMsgInfor
    End If
    GetItemInfo_���� = strMsgInfor
End Function
