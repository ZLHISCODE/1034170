Attribute VB_Name = "mdl�山ũҽ"
Option Explicit
Public gstrBusiness_�山ũҽ As String
Public gstrInput_�山ũҽ As String
Public gstrOutput_�山ũҽ As String

Private Const mstrAmountFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrPriceFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrDateFormat As String = "yyyy-MM-dd HH:mm:ss"
Private Const gstrSplit_�山ũҽ As String = "|"
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public Enum Business_�山ũҽ
    ��ȡ������Ϣ_�山ũҽ = 300
    ����Ǽ�_�山ũҽ = 101
    ������Ϣ�޸�_�山ũҽ = 102
    ����Ǽ�ȡ��_�山ũҽ = 103
    ������ϸ�ϴ�_�山ũҽ = 104
    ������ϸ����_�山ũҽ = 105
    Ԥ����_�山ũҽ = 106
    ��ʽ����_�山ũҽ = 107
    ��������_�山ũҽ = 108
End Enum

Private Type ComInfo_�山ũҽ
    ҽԺ���� As String
    ҽԺ���� As String
    ҵ������ As String
    ҽ��֤�� As String
    ���˱�� As String
    ������ˮ�� As String
    ������ˮ�� As String
    �������� As String                      '���������֤�󷵻صļ�������
    ����֢ As String
    �ܷ��� As Currency                      'HIS
    �ܷ���_���� As Currency                 '���ĵķ����ܶ�
    ���������� As String
End Type
Public gComInfo_�山ũҽ As ComInfo_�山ũҽ

Private gobjYB As Object   '���������ö���ı�����
Private mblnInit As Boolean
Private strFields As String, strValues As String
Private mrsOutExse As New ADODB.Recordset

Public Function ��ݱ�ʶ_�山ũҽ(Optional bytType As Byte, Optional lng����ID As Long) As String
    Dim strInput As String
    Dim strIdentify As String
    Dim strRegistCode As String             '�Һŵ���
    Dim strRegisterOffice As String         '�������
    Dim strRegisterDoctor As String         'ҽ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    strIdentify = frmIdentify�山ũҽ.GetPatient(bytType, lng����ID)
    If strIdentify = "" Then Exit Function
    If Not (bytType = 1 Or bytType = 0 Or bytType = 3) Then Exit Function
    
    '��������Ǽ�
    If bytType = 0 Then
        '��Σ�����ҽ�ƺ��멦����ҽ�Ʋ�����ҽԺ����ĹҺź��멦�����ҽ�����ҽԺ����Ŀ��ҩ������ҽ����" & _
        ҽԺ����ϩ�ҽԺ����Ǽǵ����ک�����֢����������Ļ������멦��������Ļ������Ʃ����쵥λ��������
        'ȡ����ҺŵĿ�����ҽ��
        gstrSQL = " Select B.���� AS �Һſ���,ִ���� AS ҽ�� " & _
                  " From ���˷��ü�¼ A,���ű� B " & _
                  " Where A.��¼����=4 And ��¼״̬=1 And ����ID=" & lng����ID & _
                  " And �Ǽ�ʱ�� Between to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 00:00:00','yyyy-MM-dd hh24:mi:ss')" & _
                  " And to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss') And Rownum<2"
        Call OpenRecordset(rsTemp, "ȡ����ҺŵĿ�����ҽ��")
        If rsTemp.RecordCount = 0 Then
            MsgBox "����û����Ч�ĹҺż�¼,�޷������������Ǽǣ�", vbInformation, gstrSysName
            Exit Function
        End If
        strRegisterOffice = rsTemp!�Һſ���
        strRegisterDoctor = rsTemp!ҽ��
        
        '��ȡ�Һŵ��ţ�ʮλ��Ψһ��ʶ
        strRegistCode = Right(CStr(zlDatabase.GetNextId("���ű�")), 10)
        strInput = gComInfo_�山ũҽ.ҽ��֤�� & gstrSplit_�山ũҽ & strRegistCode & gstrSplit_�山ũҽ & _
            gComInfo_�山ũҽ.ҵ������ & gstrSplit_�山ũҽ & strRegisterOffice & gstrSplit_�山ũҽ & _
            strRegisterDoctor & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.�������� & gstrSplit_�山ũҽ & _
            Format(zlDatabase.Currentdate(), mstrDateFormat) & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.����֢ & gstrSplit_�山ũҽ & _
            gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & _
            gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & UserInfo.����
        Call ���ýӿ�_׼��_�山ũҽ(����Ǽ�_�山ũҽ, strInput)
        If Not ���ýӿ�_�山ũҽ() Then Exit Function
        
        '������ˮ��
        gComInfo_�山ũҽ.������ˮ�� = Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(1)
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�山ũҽ & ",'˳���','''" & gComInfo_�山ũҽ.������ˮ�� & "''')"
        Call ExecuteProcedure("���������ˮ��")
        
        '��ʼ����¼��
        strFields = "������ˮ��," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & "��ϸ��ˮ��," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & _
            "��������," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & "ҽ������," & adLongVarChar & ",50" & gstrSplit_�山ũҽ & _
            "��Ŀ����," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & "���," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & _
            "����," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & "����," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & _
            "����," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & "�ϴ���ˮ��," & adLongVarChar & ",20"
        Call Record_Init(mrsOutExse, strFields)
    End If
    
    '���±����ʻ������Ϣ��ͳ�����š�ҵ�����ͣ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�山ũҽ & ",'ҵ������','''" & gComInfo_�山ũҽ.ҵ������ & "''')"
    Call ExecuteProcedure("����ҵ������")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�山ũҽ & ",'����֢','''" & gComInfo_�山ũҽ.����֢ & "''')"
    Call ExecuteProcedure("���沢��֢")
    
    '���ز�����Ϣ��
    ��ݱ�ʶ_�山ũҽ = strIdentify
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ȡ������Ǽ�_�山ũҽ(Optional bytType As Byte, Optional lng����ID As Long)
    'ȡ�����ξ���Ǽǣ����Ԥ����ʱ���ϴ�������ϸ������ȡ����ϸ����ȡ������Ǽ�
    If bytType <> 0 Then Exit Sub       'ֻ�����������Һ�
    On Error GoTo ErrHand
    
    '�������ϴ��ϴ������д�����ϸ
    With mrsOutExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Call ���ýӿ�_׼��_�山ũҽ(������ϸ����_�山ũҽ, !�ϴ���ˮ��)
            Call ���ýӿ�_�山ũҽ
            .MoveNext
        Loop
    End With
    
    'ȡ������Ǽ�
    Call ���ýӿ�_׼��_�山ũҽ(����Ǽ�ȡ��_�山ũҽ, gComInfo_�山ũҽ.������ˮ��)
    Call ���ýӿ�_�山ũҽ
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ҽ����ʼ��_�山ũҽ(Optional ByVal blnTest As Boolean = False) As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim strServer As String, strUser As String, strPass As String, strDatabase As String
    Dim rsTemp As New ADODB.Recordset
    Dim cnTest As New ADODB.Connection

    On Error Resume Next
    
    If mblnInit = False Then
        If Not blnTest Then '����ǲ��ԣ���˵���Ǳ��ղ������ô�����
            '��������ҽ��������������
            gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=" & TYPE_�山ũҽ
            Call OpenRecordset(rsTemp, "��ȡ���ղ���")
            
            Do Until rsTemp.EOF
                Select Case rsTemp("������")
                    Case "ҽ���û���"
                        strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Case "ҽ��������"
                        strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Case "ҽ���û�����"
                        strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Case "ҽ��ʵ����"
                        strDatabase = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                End Select
                rsTemp.MoveNext
            Loop
            
            If OpenSQLServer(cnTest, strServer, strUser, strPass, strDatabase) = False Then
                MsgBox "�޷����ӵ�ǰ�û������鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Set gobjYB = CreateObject("HisSel.Handld")
        '��������Ƿ���
        If gobjYB Is Nothing Then
            MsgBox "ҽ����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            '��������ҽ�������� 204-04-07
            Exit Function
        End If
        'ȡҽԺ����
        gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�山ũҽ
        Call OpenRecordset(rsTemp, "��ȡҽԺ����")
        gComInfo_�山ũҽ.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
        'ȡҽԺ����
        gstrSQL = "Select JGMC ҽԺ���� From JGDJ Where JGBM='" & gComInfo_�山ũҽ.ҽԺ���� & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, cnTest
        gComInfo_�山ũҽ.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
        
        cnTest.Close
        If Not blnTest Then mblnInit = True
    End If
    
    ҽ����ʼ��_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ������_�山ũҽ() As Boolean
    '��
End Function

Public Function ҽ����ֹ_�山ũҽ() As Boolean
    On Error Resume Next
    
    Set gobjYB = Nothing
    ҽ����ֹ_�山ũҽ = True
End Function

Public Sub ���ýӿ�_׼��_�山ũҽ(ByVal strBusiness As String, Optional ByVal strInput As String = "")
    gstrBusiness_�山ũҽ = strBusiness
    gstrInput_�山ũҽ = strInput
End Sub

Public Function ���ýӿ�_�山ũҽ() As Boolean
    Dim arrOutput
    Dim lngResult As Long
    On Error GoTo ErrHand
    
    Call DebugTool("�ӿ������Ϣ:" & gstrInput_�山ũҽ)
    Call gobjYB.Business(gstrBusiness_�山ũҽ, gstrInput_�山ũҽ, gstrOutput_�山ũҽ)
    arrOutput = Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)
    lngResult = Val(arrOutput(0))
    If lngResult < 0 Then               '������Ϣ
        MsgBox "��������[" & gstrBusiness_�山ũҽ & "]�������[" & lngResult & "]" & arrOutput(1), vbInformation, gstrSysName
        Exit Function
    End If
    
    ���ýӿ�_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �����������_�山ũҽ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strInput As String
    Dim lng����ID As Long
    Dim str�������� As String, str������ As String, strҽ������ As String, str��Ŀ���� As String, str��� As String, str���� As String
    Dim dbl�ʻ�֧�� As Double, dbl�ֽ� As Double, dbl�Żݽ�� As Double
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo ErrHand
    
    lng����ID = rs��ϸ!����ID
    str�������� = Format(zlDatabase.Currentdate, mstrDateFormat)
    '�������ϴ��ϴ������д�����ϸ
    With mrsOutExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strInput = !�ϴ���ˮ��
            Call ���ýӿ�_׼��_�山ũҽ(������ϸ����_�山ũҽ, strInput)
            If Not ���ýӿ�_�山ũҽ() Then Exit Function
            .MoveNext
        Loop
    End With
    
    '��ȡ�ò��˵ľ���ʱ��
    gstrSQL = "Select to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') As ����ʱ�� From �����ʻ�" & _
        " Where ����=" & TYPE_�山ũҽ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ�ò��˵ľ���ʱ��")
    str������ = GetSequence(Format(rsTemp!����ʱ��, "yyMMddHHmmss")) & Left(CStr(Rnd() * 100), 2)
    
    '��ʼ����¼��
    strFields = "������ˮ��," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & "��ϸ��ˮ��," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & _
        "��������," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & "ҽ������," & adLongVarChar & ",50" & gstrSplit_�山ũҽ & _
        "��Ŀ����," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & "���," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & _
        "����," & adLongVarChar & ",100" & gstrSplit_�山ũҽ & "����," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & _
        "����," & adLongVarChar & ",20" & gstrSplit_�山ũҽ & "�ϴ���ˮ��," & adLongVarChar & ",20"
    Call Record_Init(mrsOutExse, strFields)
        
    '�õ����ν�����ܷ���
    strFields = "������ˮ��" & gstrSplit_�山ũҽ & "��ϸ��ˮ��" & gstrSplit_�山ũҽ & _
            "��������" & gstrSplit_�山ũҽ & "ҽ������" & gstrSplit_�山ũҽ & _
            "��Ŀ����" & gstrSplit_�山ũҽ & "���" & gstrSplit_�山ũҽ & _
            "����" & gstrSplit_�山ũҽ & "����" & gstrSplit_�山ũҽ & "����" & gstrSplit_�山ũҽ & "�ϴ���ˮ��"
    With rs��ϸ
        If .RecordCount > 99 Then
            MsgBox "���ﴦ����ϸ���ܳ���99����¼��", vbInformation, gstrSysName
            Exit Function
        End If
        '������ܶ�
        gComInfo_�山ũҽ.�ܷ��� = 0
        Do While Not .EOF
            gComInfo_�山ũҽ.�ܷ��� = gComInfo_�山ũҽ.�ܷ��� + !ʵ�ս��
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            '��ȡ�շ�ϸĿ�������Ϣ
            gstrSQL = " Select A.��� AS �շ����,A.����,A.���,B.��Ŀ���� From �շ�ϸĿ A,����֧����Ŀ B" & _
                      " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=" & TYPE_�山ũҽ & " And A.ID=" & !�շ�ϸĿID
            Call OpenRecordset(rsItem, "��ȡ��Ŀ��Ϣ")
            str��Ŀ���� = Nvl(rsItem!����)
            strҽ������ = Nvl(rsItem!��Ŀ����)
            str��� = Nvl(rsItem!���)
            If InStr(1, str���, "|") <> 0 Then str��� = Mid(str���, 1, InStr(1, str���, "|") - 1)
            
            '�����ҩƷ��ȡ����
            str���� = ""
            If InStr(1, "5,6,7", rsItem!�շ����) <> 0 Then
                gstrSQL = "SELECT ���� FROM ҩƷ���� WHERE ����=(SELECT ���� FROM ҩƷ��Ϣ WHERE ҩ��ID=(SELECT ҩ��ID FROM ҩƷĿ¼ WHERE ҩƷID=" & !�շ�ϸĿID & "))"
                Call OpenRecordset(rsItem, "��ȡҩƷ�ļ���")
                str���� = Nvl(rsItem!����)
            End If
            
            '������ϸ�ϴ���Σ�����ҽ�Ʋ��˵ľ�����ˮ��|ҽԺ�����Ĵ����ı���|ҽԺ������������|ҩƷ��������Ŀ����ҽ�ƶ��յı���| & _
            ҩƷ����������Ŀ������|ҩƷ�Ĺ��|ҩƷ�ļ���|ҩƷ����������Ŀ�ĵ���|ҩƷ�����������ƴ���
            strInput = gComInfo_�山ũҽ.������ˮ�� & gstrSplit_�山ũҽ & str������ & String(2 - Len(CStr(.AbsolutePosition)), "0") & .AbsolutePosition & gstrSplit_�山ũҽ & _
                str�������� & gstrSplit_�山ũҽ & strҽ������ & gstrSplit_�山ũҽ & Left(str��Ŀ����, 15) & gstrSplit_�山ũҽ & _
                Left(str���, 10) & gstrSplit_�山ũҽ & Left(str����, 10) & gstrSplit_�山ũҽ & Format(!����, mstrPriceFormat) & gstrSplit_�山ũҽ & Format(!����, mstrAmountFormat)
            
            Call ���ýӿ�_׼��_�山ũҽ(������ϸ�ϴ�_�山ũҽ, strInput)
            If Not ���ýӿ�_�山ũҽ() Then Exit Function
            
            '�����ϴ��Ĵ�����ϸд���¼��
            strValues = gComInfo_�山ũҽ.������ˮ�� & gstrSplit_�山ũҽ & str������ & String(2 - Len(CStr(.AbsolutePosition)), "0") & .AbsolutePosition & gstrSplit_�山ũҽ & _
                str�������� & gstrSplit_�山ũҽ & strҽ������ & gstrSplit_�山ũҽ & str��Ŀ���� & gstrSplit_�山ũҽ & _
                str��� & gstrSplit_�山ũҽ & str���� & gstrSplit_�山ũҽ & Format(!����, mstrPriceFormat) & gstrSplit_�山ũҽ & Format(!����, mstrAmountFormat) & gstrSplit_�山ũҽ & Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(1)
            Call Record_Add(mrsOutExse, strFields, strValues)
            
            .MoveNext
        Loop
    End With
    
    'Ԥ�������Σ����˾���Ǽǵ���ˮ��|��������|סԺ�Ĵ���|���쵥λ|������|��������
    strInput = gComInfo_�山ũҽ.������ˮ�� & gstrSplit_�山ũҽ & "01" & gstrSplit_�山ũҽ & "0" & gstrSplit_�山ũҽ & _
        gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & UserInfo.���� & gstrSplit_�山ũҽ & str��������
    gComInfo_�山ũҽ.���������� = strInput
    Call ���ýӿ�_׼��_�山ũҽ(Ԥ����_�山ũҽ, strInput)
    If Not ���ýӿ�_�山ũҽ() Then Exit Function
    
    '���Σ�ִ�д��멦������ˮ�ũ���ν���ҽԺ�ܵĽ�����ҽԺ�¸�����ܽ�����ҽ�ư칫�ҳ��Ͽ��Բμӱ����Ľ�ʵ�ʱ����Ľ������Ը��Ľ��
    dbl�Żݽ�� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(2)) - Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(3))
    dbl�ʻ�֧�� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(5))
    dbl�ֽ� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(6))
    
    str���㷽ʽ = "��ͥ�ʻ�;" & dbl�ʻ�֧�� & ";0|�Żݽ��;" & dbl�Żݽ�� & ";0"
    �����������_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_�山ũҽ(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    Dim lng����ID As Long
    Dim strInput As String
    Dim str�������� As String, str������ˮ�� As String, str����˳��� As String
    Dim dbl����ͳ�� As Double, dblͳ�ﱨ�� As Double, dbl�ֽ� As Double, dbl�Żݽ�� As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    'Ԥ�������Σ����˾���Ǽǵ���ˮ��|��������|סԺ�Ĵ���|���쵥λ|������|��������,��������
    str�������� = Format(zlDatabase.Currentdate, mstrDateFormat)
    strInput = gComInfo_�山ũҽ.���������� & gstrSplit_�山ũҽ & str��������
    Call ���ýӿ�_׼��_�山ũҽ(��ʽ����_�山ũҽ, strInput)
    If Not ���ýӿ�_�山ũҽ() Then Exit Function
    
    '���Σ�ִ�д��멦������ˮ�ũ���ν���ҽԺ�ܵĽ�����ҽԺ�¸�����ܽ�����ҽ�ư칫�ҳ��Ͽ��Բμӱ����Ľ�ʵ�ʱ����Ľ������Ը��Ľ��
    str������ˮ�� = Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(1)
    dbl�Żݽ�� = Format(Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(2)) - Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(3)), "#.00")
    dbl����ͳ�� = Format(Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(4)), "#.00")
    dblͳ�ﱨ�� = Format(Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(5)), "#0.00")   'ͳ�ﱨ�����Ǽ�ͥ�ʻ�����֧����
    dbl�ֽ� = Format(Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(6)), "#0.00")
    
    'ȡ����ID
    gstrSQL = "Select ����ID From ���˷��ü�¼ Where ����ID=" & lng����ID & " And Rownum<2"
    Call OpenRecordset(rsTemp, "ȡ�ò��˵�ID")
    lng����ID = rsTemp!����ID
    
    'ȡ����˳���
    gstrSQL = "Select ˳��� From �����ʻ� Where ����=" & TYPE_�山ũҽ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ����˳���")
    str����˳��� = Nvl(rsTemp!˳���)
    
    '���汾�ν������
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�山ũҽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_�山ũҽ.�ܷ��� & "," & dbl�ֽ� & "," & dbl�Żݽ�� & "," & dbl����ͳ�� & "," & dblͳ�ﱨ�� & ",0,0," & _
        dblͳ�ﱨ�� & ",'" & str����˳��� & "|" & str������ˮ�� & "',null,null,null)"
    Call ExecuteProcedure("���������շ�����")
    
    �������_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ����������_�山ũҽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim str������ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    'ȡ������¼�Ľ���ID�����ݺ�
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���²����Ľ���ID")
    lng����ID = rsTemp!����id
    
    'ȡ������ˮ��
    gstrSQL = "Select * From ���ս����¼ Where ����=1 And ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ������ˮ��")
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ�ԭʼ�����¼���޷�����������������", vbInformation, gstrSysName
        Exit Function
    End If
    gComInfo_�山ũҽ.������ˮ�� = Split(rsTemp!֧��˳���, gstrSplit_�山ũҽ)(0)
    str������ˮ�� = Split(rsTemp!֧��˳���, gstrSplit_�山ũҽ)(1)
    
    '���ý������
    Call ���ýӿ�_׼��_�山ũҽ(��������_�山ũҽ, str������ˮ��)
    If Not ���ýӿ�_�山ũҽ() Then Exit Function
    
    'ȡ������Ǽ�
    Call ���ýӿ�_׼��_�山ũҽ(����Ǽ�ȡ��_�山ũҽ, gComInfo_�山ũҽ.������ˮ��)
    Call ���ýӿ�_�山ũҽ
    
    '���汾�ν������
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�山ũҽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & ",0,0," & _
        -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",null,null,null,null)"
    Call ExecuteProcedure("����������")
    
    ����������_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�山ũҽ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim strInput As String
    Dim strRegistCode As String             '�Һŵ���
    Dim strInHospitalDate As String         '��Ժ����
    Dim strRegisterOffice As String         '�������
    Dim strRegisterDoctor As String         'ҽ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    'ȡ������ҽ��
    gstrSQL = " Select A.��Ժ����,B.���� ����,A.סԺҽʦ ҽ�� From ������ҳ A,���ű� B " & _
              " Where A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID"
    Call OpenRecordset(rsTemp, "ȡ������ҽ��")
    strInHospitalDate = Format(rsTemp!��Ժ����, mstrDateFormat)
    strRegisterDoctor = Nvl(rsTemp!ҽ��)
    strRegisterOffice = Nvl(rsTemp!����)
    
    '��Σ�����ҽ�ƺ��멦����ҽ�Ʋ�����ҽԺ����ĹҺź��멦�����ҽ�����ҽԺ����Ŀ��ҩ������ҽ����" & _
    ҽԺ����ϩ�ҽԺ����Ǽǵ����ک�����֢����������Ļ������멦��������Ļ������Ʃ����쵥λ��������
    '��ȡ�Һŵ��ţ�ʮλ��Ψһ��ʶ
    strRegistCode = Right(CStr(zlDatabase.GetNextId("���ű�")), 10)
    strInput = gComInfo_�山ũҽ.ҽ��֤�� & gstrSplit_�山ũҽ & strRegistCode & gstrSplit_�山ũҽ & _
        gComInfo_�山ũҽ.ҵ������ & gstrSplit_�山ũҽ & strRegisterOffice & gstrSplit_�山ũҽ & _
        strRegisterDoctor & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.�������� & gstrSplit_�山ũҽ & _
        strInHospitalDate & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.����֢ & gstrSplit_�山ũҽ & _
        gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & _
        gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & UserInfo.����
    Call ���ýӿ�_׼��_�山ũҽ(����Ǽ�_�山ũҽ, strInput)
    If Not ���ýӿ�_�山ũҽ() Then Exit Function
    
    '������ˮ��
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�山ũҽ & ",'˳���','''" & Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(1) & "''')"
    Call ExecuteProcedure("���������ˮ��")
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�山ũҽ & ")"
    Call ExecuteProcedure("������Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�山ũҽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "��ҽ�����˴���δ����ã��������������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȡԭ������ˮ��
    gstrSQL = "Select ˳��� From �����ʻ� Where ����=" & TYPE_�山ũҽ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ������ˮ��")
    gComInfo_�山ũҽ.������ˮ�� = rsTemp!˳���
    
    '���þ���Ǽ����Ͻӿ�
    Call ���ýӿ�_׼��_�山ũҽ(����Ǽ�ȡ��_�山ũҽ, gComInfo_�山ũҽ.������ˮ��)
    If Not ���ýӿ�_�山ũҽ Then Exit Function
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�山ũҽ & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    
    ��Ժ�Ǽǳ���_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_�山ũҽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo ErrHand
    
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�山ũҽ & ")"
    Call ExecuteProcedure("��Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�山ũҽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo ErrHand
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�山ũҽ & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_�山ũҽ(strSelfNo As String) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '����: ��ȡ�α����˸����ʻ����
    '����: strSelfNO-���˸��˱��
    '����: ���ظ����ʻ����Ľ��
    '�����������ؼ�ͥ�ʻ���סԺ���ظ����ʻ����
    gstrSQL = "Select Nvl(�ʻ����,0) AS �����ʻ�,Nvl(��ͥ�ʻ����,0) AS ��ͥ�ʻ�,����ID From �����ʻ� Where ҽ����='" & strSelfNo & "'"
    Call OpenRecordset(rsTemp, "��ȡ����ID")
    �������_�山ũҽ = rsTemp!�����ʻ� + rsTemp!��ͥ�ʻ�
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����ϴ�_�山ũҽ(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal strNO As String) As Boolean
    Dim lng����ID As Long, lng����ID As Long
    Dim strInput As String
    Dim blnInsure As Boolean
    Dim str������ˮ�� As String, str������ As String
    Dim str��Ŀ���� As String, strҽ������ As String, str��� As String, str���� As String
    Dim rsDetail As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo ErrHand
    '�ϴ�������ϸ
    '�򿪱��δ��ϴ��Ĵ�����ϸ
    gstrSQL = " Select ID,��¼����,��¼״̬,NO,���,�շ����,����ID,�շ�ϸĿID,�Ǽ�ʱ��,Nvl(����,1)*���� AS ����,ʵ�ս��/(Nvl(����,1)*����) AS �۸�" & _
              " From ���˷��ü�¼" & _
              " Where ��¼����=" & int���� & " ANd ��¼״̬=" & int״̬ & " And NO='" & strNO & "' And Nvl(�Ƿ��ϴ�,0)=0" & _
              " Order by ����ID"
    Call OpenRecordset(rsDetail, "��ȡ���δ��ϴ��Ĵ�����ϸ")
    
    '�ȼ����ϸ�������������ʣ�ֻ���ҽ�����˵��������ʵĴ�����ϸ��
    With rsDetail
        lng����ID = 0
        If int״̬ = 1 Then
            Do While Not .EOF
                If lng����ID <> !����ID Then
                    lng����ID = !����ID
                    blnInsure = IsYBPatient(lng����ID, str������ˮ��)
                End If
                If blnInsure Then
                    If !���� < 0 Then
                        MsgBox "�山ũ�����ҽ�ƽӿڲ�֧��Ϊҽ�����˽��и������ʣ���ֱ�ӳ���ԭʼ������ϸ��", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                .MoveNext
            Loop
        End If
        If .RecordCount <> 0 Then .MoveFirst
        
        '�ϴ��������ϴ������㼸����û�ϴ��ɹ�Ҳ�����浥�ݣ�
        lng����ID = 0
        Do While Not .EOF
            If lng����ID <> !����ID Then
                lng����ID = !����ID
                blnInsure = IsYBPatient(lng����ID, str������ˮ��)
            End If
            
            If blnInsure Then
                '�Է���ID��ʮλ����Ϊ���δ�����ϸ��ˮ��
                lng����ID = !ID
                str������ = Right(CStr(lng����ID), 10)
                
                '��ȡ�շ�ϸĿ�������Ϣ
                gstrSQL = " Select A.��� AS �շ����,A.����,A.���,B.��Ŀ���� From �շ�ϸĿ A,����֧����Ŀ B" & _
                          " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=" & TYPE_�山ũҽ & " And A.ID=" & !�շ�ϸĿID
                Call OpenRecordset(rsItem, "��ȡ��Ŀ��Ϣ")
                str��Ŀ���� = Nvl(rsItem!����)
                strҽ������ = Nvl(rsItem!��Ŀ����)
                str��� = Nvl(rsItem!���)
                If InStr(1, str���, "|") <> 0 Then str��� = Mid(str���, 1, InStr(1, str���, "|") - 1)
                
                '�����ҩƷ��ȡ����
                str���� = ""
                If InStr(1, "5,6,7", rsItem!�շ����) <> 0 Then
                    gstrSQL = "SELECT ���� FROM ҩƷ���� WHERE ����=(SELECT ���� FROM ҩƷ��Ϣ WHERE ҩ��ID=(SELECT ҩ��ID FROM ҩƷĿ¼ WHERE ҩƷID=" & !�շ�ϸĿID & "))"
                    Call OpenRecordset(rsItem, "��ȡҩƷ�ļ���")
                    str���� = Nvl(rsItem!����)
                End If
                
                '������ϸ�ϴ���Σ�����ҽ�Ʋ��˵ľ�����ˮ��|ҽԺ�����Ĵ����ı���|ҽԺ������������|ҩƷ��������Ŀ����ҽ�ƶ��յı���| & _
                ҩƷ����������Ŀ������|ҩƷ�Ĺ��|ҩƷ�ļ���|ҩƷ����������Ŀ�ĵ���|ҩƷ�����������ƴ���
                If int״̬ <> 2 Then
                    strInput = str������ˮ�� & gstrSplit_�山ũҽ & str������ & gstrSplit_�山ũҽ & _
                        Format(!�Ǽ�ʱ��, mstrDateFormat) & gstrSplit_�山ũҽ & strҽ������ & gstrSplit_�山ũҽ & Left(str��Ŀ����, 15) & gstrSplit_�山ũҽ & _
                        Left(str���, 10) & gstrSplit_�山ũҽ & Left(str����, 10) & gstrSplit_�山ũҽ & Format(!�۸�, mstrPriceFormat) & gstrSplit_�山ũҽ & Format(!����, mstrAmountFormat)
                Else
                    'ȡԭʼ���ü�¼ID
                    gstrSQL = "Select ժҪ From ���˷��ü�¼ Where ��¼����=" & !��¼���� & " And ��¼״̬=3 And NO='" & !no & "' And ���=" & !���
                    Call OpenRecordset(rsItem, "ȡ���ü�¼ID")
                    strInput = Nvl(rsItem!ժҪ)
                End If
                
                Call ���ýӿ�_׼��_�山ũҽ(IIf(int״̬ <> 2, ������ϸ�ϴ�_�山ũҽ, ������ϸ����_�山ũҽ), strInput)
                If Not ���ýӿ�_�山ũҽ() Then
                    �����ϴ�_�山ũҽ = True
                    Exit Function
                End If
                
                '���ϴ���־
                If int״̬ <> 2 Then
                    gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & lng����ID & ",0,'" & Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(1) & "')"
                Else
                    gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !no & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                End If
                Call ExecuteProcedure("���ϴ���־")
            End If
            .MoveNext
        Loop
    End With
    
    �����ϴ�_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�山ũҽ(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim intDo As Integer
    Dim lng����ID As Long, lng��ҳID As Long, lngסԺ���� As Long
    Dim dbl�ʻ�֧�� As Double, dbl�ֽ� As Double, dbl�Żݽ�� As Double
    Dim bln�������� As Boolean                  '��������(01)��תԺ����(02)
    Dim strInput As String
    Dim str������ As String, str������ˮ�� As String
    Dim str��Ŀ���� As String, strҽ������ As String, str��� As String, str���� As String
    Dim rsItem As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '�����ȳ�Ժ�����ܽ��н���
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "�����ȳ�Ժ�����ܽ��н��㣡", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ��Ժ��ʽ
    gstrSQL = "Select ��Ժ��ʽ,סԺ����,��ҳID From ������ҳ Where (����ID,��ҳID) in (Select ����ID,סԺ���� From ������Ϣ Where ����ID=" & lng����ID & ")"
    Call OpenRecordset(rsItem, "ȡ��Ժ��ʽ")
    lng��ҳID = rsItem!��ҳID
    lngסԺ���� = Nvl(rsItem!סԺ����, 0)
    bln�������� = IIf(rsItem!��Ժ��ʽ = "תԺ", False, True)
    
    '��ȡ���˵ľ�����ˮ��
    gstrSQL = "Select ˳��� From �����ʻ� Where ����=" & TYPE_�山ũҽ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsItem, "��ȡ���˵ľ�����ˮ��")
    str������ˮ�� = Nvl(rsItem!˳���)
    
    '��ȡ���η�����ϸ
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,C.��Ŀ���� AS ҽ����Ŀ����,B.����,B.����,A.ʵ�ս�� AS ���" & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸�,A.������ AS ҽ��,A.�Ǽ�ʱ�� " & _
              "  From ���˷��ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              "  where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & " and A.���ʷ���=1 And A.����Ա���� is not null AND A.ʵ�ս�� IS NOT NULL " & _
              "        and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= " & TYPE_���� & _
              "  Order by A.����ID,A.����ʱ��"
    Call OpenRecordset(rs��ϸ, "��ȡ���η�����ϸ")
    
    With rsExse
        '������ܶ�
        gComInfo_�山ũҽ.�ܷ��� = 0
        Do While Not .EOF
            gComInfo_�山ũҽ.�ܷ��� = gComInfo_�山ũҽ.�ܷ��� + !���
            .MoveNext
        Loop
    End With
        
    With rs��ϸ
        For intDo = 1 To 2
            .Filter = IIf(intDo = 1, "��¼״̬<>2", "��¼״̬=2")
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                If Nvl(!�Ƿ��ϴ�, 0) = 0 Then
                    '�Է���ID��ʮλ����Ϊ���δ�����ϸ��ˮ��
                    gstrSQL = "Select ID,ʵ�ս�� From ���˷��ü�¼ Where ��¼����=" & !��¼���� & " And ��¼״̬=" & !��¼״̬ & " And NO='" & !no & "' And ���=" & !���
                    Call OpenRecordset(rsItem, "ȡ���ü�¼ID")
                    If Not IsNull(rsItem!ʵ�ս��) Then
                        lng����ID = rsItem!ID
                        str������ = Right(rsItem!ID, 10)
                        
                        '��ȡ�շ�ϸĿ�������Ϣ
                        gstrSQL = " Select A.��� AS �շ����,A.����,A.���,B.��Ŀ���� From �շ�ϸĿ A,����֧����Ŀ B" & _
                                  " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=" & TYPE_�山ũҽ & " And A.ID=" & !�շ�ϸĿID
                        Call OpenRecordset(rsItem, "��ȡ��Ŀ��Ϣ")
                        str��Ŀ���� = Nvl(rsItem!����)
                        strҽ������ = Nvl(rsItem!��Ŀ����)
                        str��� = Nvl(rsItem!���)
                        If InStr(1, str���, "|") <> 0 Then str��� = Mid(str���, 1, InStr(1, str���, "|") - 1)
                        
                        '�����ҩƷ��ȡ����
                        str���� = ""
                        If InStr(1, "5,6,7", rsItem!�շ����) <> 0 Then
                            gstrSQL = "SELECT ���� FROM ҩƷ���� WHERE ����=(SELECT ���� FROM ҩƷ��Ϣ WHERE ҩ��ID=(SELECT ҩ��ID FROM ҩƷĿ¼ WHERE ҩƷID=" & !�շ�ϸĿID & "))"
                            Call OpenRecordset(rsItem, "��ȡҩƷ�ļ���")
                            str���� = Nvl(rsItem!����)
                        End If
                        
                        '������ϸ�ϴ���Σ�����ҽ�Ʋ��˵ľ�����ˮ��|ҽԺ�����Ĵ����ı���|ҽԺ������������|ҩƷ��������Ŀ����ҽ�ƶ��յı���| & _
                        ҩƷ����������Ŀ������|ҩƷ�Ĺ��|ҩƷ�ļ���|ҩƷ����������Ŀ�ĵ���|ҩƷ�����������ƴ���
                        If intDo = 1 Then
                            strInput = str������ˮ�� & gstrSplit_�山ũҽ & str������ & gstrSplit_�山ũҽ & _
                                Format(!�Ǽ�ʱ��, mstrDateFormat) & gstrSplit_�山ũҽ & strҽ������ & gstrSplit_�山ũҽ & Left(str��Ŀ����, 15) & gstrSplit_�山ũҽ & _
                                Left(str���, 10) & gstrSplit_�山ũҽ & Left(str����, 10) & gstrSplit_�山ũҽ & Format(!�۸�, mstrPriceFormat) & gstrSplit_�山ũҽ & Format(!����, mstrAmountFormat)
                        Else
                            'ȡԭʼ���ü�¼ID
                            gstrSQL = "Select ժҪ From ���˷��ü�¼ Where ��¼����=" & !��¼���� & " And ��¼״̬=3 And NO='" & !no & "' And ���=" & !���
                            Call OpenRecordset(rsItem, "ȡ���ü�¼ID")
                            strInput = Nvl(rsItem!ժҪ)
                        End If
                        
                        Call ���ýӿ�_׼��_�山ũҽ(IIf(intDo = 1, ������ϸ�ϴ�_�山ũҽ, ������ϸ����_�山ũҽ), strInput)
                        If Not ���ýӿ�_�山ũҽ() Then Exit Function
                        
                        '���ϴ���־
                        If intDo = 1 Then
                            gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & lng����ID & ",0,'" & Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(1) & "')"
                        Else
                            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !no & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                        End If
                        Call ExecuteProcedure("���ϴ���־")
                    End If
                End If
                .MoveNext
            Loop
        Next
    End With
    
    '����סԺԤ���㣬��Σ����˾���Ǽǵ���ˮ��|�������|סԺ�Ĵ���|���쵥λ|������|��������
    strInput = str������ˮ�� & gstrSplit_�山ũҽ & IIf(bln��������, "01", "02") & gstrSplit_�山ũҽ & _
        lngסԺ���� & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & UserInfo.���� & gstrSplit_�山ũҽ & Format(zlDatabase.Currentdate, mstrDateFormat)
    Call ���ýӿ�_׼��_�山ũҽ(Ԥ����_�山ũҽ, strInput)
    If Not ���ýӿ�_�山ũҽ() Then Exit Function
    
    '���Σ�ִ�д��멦������ˮ�ũ���ν���ҽԺ�ܵĽ�����ҽԺ�¸�����ܽ�����ҽ�ư칫�ҳ��Ͽ��Բμӱ����Ľ�ʵ�ʱ����Ľ������Ը��Ľ��
    dbl�Żݽ�� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(2)) - Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(3))
    dbl�ʻ�֧�� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(5))
    dbl�ֽ� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(6))
    
    סԺ�������_�山ũҽ = "�����ʻ�;" & dbl�ʻ�֧�� & ";0|�Żݽ��;" & dbl�Żݽ�� & ";0"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�山ũҽ(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim strInput As String
    Dim str������ˮ�� As String, str������ˮ�� As String
    Dim lngסԺ���� As Long, lng��ҳID As Long
    Dim bln�������� As Boolean
    Dim dbl����ͳ�� As Double, dblͳ�ﱨ�� As Double, dbl�ֽ� As Double, dbl�Żݽ�� As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    'ȡ��Ժ��ʽ
    gstrSQL = "Select ��ҳID,��Ժ��ʽ,סԺ���� From ������ҳ Where (����ID,��ҳID) in (Select ����ID,סԺ���� From ������Ϣ Where ����ID=" & lng����ID & ")"
    Call OpenRecordset(rsTemp, "ȡ��Ժ��ʽ")
    lng��ҳID = Nvl(rsTemp!��ҳID, 1)
    lngסԺ���� = Nvl(rsTemp!סԺ����, 0)
    bln�������� = IIf(rsTemp!��Ժ��ʽ = "תԺ", False, True)
    
    '��ȡ���˵ľ�����ˮ��
    gstrSQL = "Select ˳��� From �����ʻ� Where ����=" & TYPE_�山ũҽ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���˵ľ�����ˮ��")
    str������ˮ�� = Nvl(rsTemp!˳���)
    
    '����סԺ���㣬��Σ����˾���Ǽǵ���ˮ��|�������|סԺ�Ĵ���|���쵥λ|������|��������
    strInput = str������ˮ�� & gstrSplit_�山ũҽ & IIf(bln��������, "01", "02") & gstrSplit_�山ũҽ & _
        lngסԺ���� & gstrSplit_�山ũҽ & gComInfo_�山ũҽ.ҽԺ���� & gstrSplit_�山ũҽ & _
        UserInfo.���� & gstrSplit_�山ũҽ & Format(zlDatabase.Currentdate, mstrDateFormat) & gstrSplit_�山ũҽ & _
        gstrSplit_�山ũҽ & Format(zlDatabase.Currentdate, mstrDateFormat)
    Call ���ýӿ�_׼��_�山ũҽ(��ʽ����_�山ũҽ, strInput)
    If Not ���ýӿ�_�山ũҽ() Then Exit Function
    
    '���Σ�ִ�д��멦������ˮ�ũ���ν���ҽԺ�ܵĽ�����ҽԺ�¸�����ܽ�����ҽ�ư칫�ҳ��Ͽ��Բμӱ����Ľ�ʵ�ʱ����Ľ������Ը��Ľ��
    str������ˮ�� = Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(1)
    dbl�Żݽ�� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(2)) - Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(3))
    dbl����ͳ�� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(4))
    dblͳ�ﱨ�� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(5))
    dbl�ֽ� = Val(Split(gstrOutput_�山ũҽ, gstrSplit_�山ũҽ)(6))
    
    '���汾�ν������
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�山ũҽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng��ҳID & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_�山ũҽ.�ܷ��� & "," & dbl�ֽ� & "," & dbl�Żݽ�� & "," & dbl����ͳ�� & "," & dblͳ�ﱨ�� & ",0,0," & _
        dblͳ�ﱨ�� & ",'" & str������ˮ�� & "|" & str������ˮ�� & "',null,null,null)"
    Call ExecuteProcedure("����סԺ��������")

    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call ExecuteProcedure("�����ʼ�¼�����ϴ���־")
    
    סԺ����_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�山ũҽ(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    Dim lng����ID As Long
    Dim str������ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    'ȡ����ID
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B where A.NO=B.NO and A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���²����Ľ���ID")
    lng����ID = rsTemp!ID
    
    'ȡ������ˮ��
    gstrSQL = "Select * From ���ս����¼ Where ����=2 And ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ������ˮ��")
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ�ԭʼ�����¼���޷�����סԺ���������", vbInformation, gstrSysName
        Exit Function
    End If
    str������ˮ�� = Split(rsTemp!֧��˳���, gstrSplit_�山ũҽ)(1)
    
    '���ý������
    Call ���ýӿ�_׼��_�山ũҽ(��������_�山ũҽ, str������ˮ��)
    If Not ���ýӿ�_�山ũҽ() Then Exit Function
    
    '���汾�ν������
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�山ũҽ & "," & rsTemp!����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & Nvl(rsTemp!��ҳID, 1) & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & ",0,0," & _
        -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",null,null,null,null)"
    Call ExecuteProcedure("����������")
    
    סԺ�������_�山ũҽ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsYBPatient(ByVal lng����ID As Long, str������ˮ�� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '�ж�ָ�����˱����Ƿ���ҽ����ݾ���
    gstrSQL = " Select 1 From ������ҳ Where ����=" & TYPE_�山ũҽ & " And (����ID,��ҳID) IN " & _
              "     (Select ����ID,סԺ���� From ������Ϣ Where ����ID=" & lng����ID & ")"
    Call OpenRecordset(rsTemp, "�ж�ָ�����˱����Ƿ���ҽ����ݾ���")
    IsYBPatient = (rsTemp.RecordCount <> 0)
    
    If IsYBPatient Then
        'ȡ���˵ľ�����ˮ��
        gstrSQL = "Select ˳��� From �����ʻ� Where ����=" & TYPE_�山ũҽ & " And ����ID=" & lng����ID
        Call OpenRecordset(rsTemp, "ȡ���˵ľ�����ˮ��")
        str������ˮ�� = Nvl(rsTemp!˳���)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer
    Dim intRecords As Integer
    '������:����
    '��������:2000-11-02
    'Ҳʹ���ڱ���
    Set RecTarget = New ADODB.Recordset
    
    With RecTarget
        If .State = 1 Then .Close
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, adLongVarChar, 100, adFldIsNullable     '0:��ʾ����
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        Do While Not SourceRec.EOF
            If Nvl(SourceRec!�Ƿ��ϴ�, 0) = 0 Then
                .AddNew
                For intFields = 0 To SourceRec.Fields.Count - 1
                    .Fields(intFields) = SourceRec.Fields(intFields).Value
                Next
                .Update
            End If
            If Nvl(SourceRec!�Ƿ��ϴ�, 0) = 0 Then
                intRecords = intRecords + 1
                If intRecords = 20 Then
                    SourceRec.MoveNext
                    Exit Do
                End If
            End If
            SourceRec.MoveNext
        Loop
    End With
    
    Set CopyNewRec = RecTarget
End Function

Private Function GetSequence(ByVal strInput As String) As String
    Dim intDo As Integer, intPos As Integer
    Dim strText As String, strSequence As String
    
    intPos = 1
    For intDo = 1 To 6
        strText = Mid(strInput, intPos, 2)
        intPos = intPos + 2
        strSequence = strSequence & Chr(asc("0") + Val(strText))
    Next
    GetSequence = strSequence
End Function

Public Function OpenSQLServer(cnYB As ADODB.Connection, ByVal strServer As String, ByVal strUser As String, ByVal strPass As String, Optional ByVal strDatabase As String = "") As Boolean
    On Error GoTo ErrHand
    With cnYB
        If .State = 1 Then .Close
        .Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & strUser & ";Password=" & strPass & ";Initial Catalog=" & strDatabase & ";Data Source=" & strServer
    End With
    
    OpenSQLServer = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
