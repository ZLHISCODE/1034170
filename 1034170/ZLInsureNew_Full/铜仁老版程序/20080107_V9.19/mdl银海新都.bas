Attribute VB_Name = "mdl�����¶�"
Option Explicit
'API��������

'1�������ϴ�
Private Declare Function DataUnloading Lib "yhybReckoning.dll" Alias "_DataUnloading@12" _
        (ByVal str_UploadData As String, ByVal str_UploadLsh As String, ByVal str_Fzxbm As String) As String

'2���ʻ�֧��
Private Declare Function reckoning Lib "yhybReckoning.dll" Alias "_reckoning@64" (ByVal str���� As String, _
        ByVal strҽ���� As String, ByVal str������ As String, ByVal str���� As String, _
        ByVal str����˳��� As String, ByVal str֧����� As String, ByVal strҽԺ���� As String, _
        ByVal str��Ժ���� As String, ByVal dbl�ʻ�֧�� As String, ByVal dat֧��ʱ�� As String, _
        ByVal dbl�ܶ� As String, ByVal dblȫ�Է� As String, ByVal dbl�ҹ��Ը� As String, _
        ByVal dbl������ As String, ByVal str������ As String, ByVal str������ As String) As String

'3����ȡ��ǰҽԺ������Ϣ
Private Declare Function GetHospitalInfo Lib "yhybReckoning.dll" Alias "_GetHospitalInfo@0" () As String

'4��������ϸ�ָ�
'Private Declare Function DivideUp Lib "yhybDivideUp.dll" Alias "_DivideUp@24" _
        (ByVal str�����ı�� As String, ByVal strҽ����Ŀ���� As String, ByVal str֧����� As String, _
        ByVal strҽ����Ա��� As String, ByVal dbl�ָ��� As Double) As String
Private Declare Function DivideUp Lib "yhybReckoning.dll" Alias "_DivideUp@24" _
        (ByVal str�����ı�� As String, ByVal strҽ����Ŀ���� As String, ByVal str֧����� As String, _
        ByVal strҽ����Ա��� As String, ByVal dbl�ָ��� As Double) As String

'5�������֧�����
Private Declare Function GetPayCount Lib "yhybReckoning.dll" Alias "_GetPayCount@48" _
        (ByVal str�����ı�� As String, ByVal str֧����� As String, _
        ByVal dbl�����Ը� As Double, ByVal dblȫ�Է� As Double, ByVal dbl�ҹ��Է� As Double, _
        ByVal dbl���� As Double, ByVal dbl�ʻ���� As Double) As String

'6�����ý���
Private Declare Function CalculateFeeCD Lib "yhybBill.dll" Alias "_CalculateFeeCD@84" _
        (ByVal dbl�����ܶ� As Double, ByVal dbl���� As Double, ByVal dblͳ���޶� As Double, _
        ByVal dblͳ��֧���ۼ� As Double, ByVal intʵ������ As Integer, ByVal dbl�ѽ������� As Double, _
        ByVal dbl�ѽ���ҹ��Ը� As Double, ByVal dbl���������� As Double, ByVal dblȫ�Է� As Double, _
        ByVal dbl�ҹ��Է� As Double, ByVal dblͳ�ﱨ������ As Double) As String
'7��ҽ������Ŀ¼�ļ�
Private Declare Function MakeTxt Lib "yhybReckoning.dll" Alias "_MakeTxt@8" (ByVal str����Ŀ¼�ļ� As String, _
        ByVal str����Ŀ¼�ļ� As String) As String

'8������������
Private Declare Function GetKard Lib "yhybReckoning.dll" Alias "_GetKard@4" (ByVal str_UploadData As String) As String

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public mint���õ���_�¶� As Integer
Private mstrҽ���� As String
Private mstr���� As String
Private mlng����ID As Long
Private mstr����� As String
Private mstrInfo As String                      '������Ϣ�����ڲ�����־�ļ�
Private mstr������ˮ�� As String                '����סԺ�����������ҵ���������˳���δ���µ������ʻ��У��������סԺ��˳���
Private mcol����ϸ As New Collection

Public Function ҽ����ʼ��_�¶�() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    '��ȡ��ǰ�ӿ����õ���
    mint���õ���_�¶� = 0
    '�������´���,���ܻ��ж������ʹ�ñ��ӿ�
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=" & TYPE_�¶� & " And ���=1"
    Call OpenRecordset(rsTemp, "��ȡ��ǰ�ӿ����õ���")
    If Not rsTemp.EOF Then mint���õ���_�¶� = Nvl(rsTemp!����ֵ, 0)
    
    ҽ����ʼ��_�¶� = True
End Function

Public Function ҽ������_�¶�() As Boolean
'���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
'���أ��ӿ����óɹ�������true�����򣬷���false
    Dim strConn As String
    
    ҽ������_�¶� = frmSet�¶�����.ShowSet
End Function

Public Function ��ݱ�ʶ_�¶�(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim str���� As String, strҽ���� As String, str���� As String
    Dim str���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��Ա��� As String, str��λ���� As String, str��λ���� As String
    Dim strIdentify As String, str���� As String, str����� As String
    Dim datCurr As Date, strҽԺ���� As String
    Dim strReturn As String, str��ˮ�� As String, strסԺ˳��� As String, str���ı�� As String, strInput As String, arrOutput As Variant
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur�����ʻ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, cur�������� As Currency, cur�����ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, bln��ȡ�ʻ������Ϣ As Boolean, curͳ���޶� As Currency
    bln��ȡ�ʻ������Ϣ = False
    
    '��ʼ��һЩ����
    mlng����ID = 0
    mstr����� = ""
    mstrҽ���� = ""
    mstr���� = ""
    
    '��ò���ҽ���š������ı�ŵ���Ϣ
    If frmIdentify����.GetIdentify(TYPE_�¶�, str����, strҽ����, str���ı��, str����) = False Then Exit Function
    
    '���ò����Ƿ���ҽ���������סԺ
    Dim rsTemp As New ADODB.Recordset
    '���ò����Ƿ���Ժ
    gstrSQL = "select nvl(��ǰ״̬,0) as ��ǰ״̬,˳��� from �����ʻ� where ҽ����='" & strҽ���� & "' and ����=" & TYPE_�¶�
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = False Then
        If rsTemp("��ǰ״̬") = 1 Then
            '˫������������Ժ�ڼ䷢������ҵ��
            strסԺ˳��� = Nvl(rsTemp!˳���)
'            If mint���õ���_�¶� = 1 Then
'                MsgBox "�ò�������ҽ�������Ժ�������ٽ��������֤��", vbInformation, gstrSysName
'                Exit Function
'            End If
        End If
    End If
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '���������֤
    If Get��ˮ��("A", strҽԺ����, str��ˮ��) = False Then Exit Function
    '����|���˱���|�����ı��|����|��ȡ������ˮ��#
    strInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & str���� & "|" & IIf(bytType = 1, "31", "11") & "#"
    Call WriteLog("DataUnloading(" & strInput & "," & str��ˮ�� & "," & str���ı�� & ")")
    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '˫���ӵ��жϣ��������Ϊ111111��˵���ǳ�ʼ���룬����Ҫ���û��޸ģ����˳����ν���
'    If mint���õ���_�¶� = 1 Then
'        If str���� = "111111" Then
'            MsgBox "������Ϊ�籣�ֳ�ʼ���룬�������������룡", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    'ȡ�÷���ֵ
    str���� = arrOutput(1)
    strҽ���� = arrOutput(3)
    
    str���� = arrOutput(4)
    str�Ա� = IIf(arrOutput(5) = "2", "Ů", "��")
    str���֤���� = arrOutput(6)
    str�������� = arrOutput(7)
    If IsDate(str��������) = False Then
        str�������� = Get��������(str���֤����, 0)
    End If
    If IsDate(str��������) Then
        lng���� = DateDiff("yyyy", CDate(str��������), zlDataBase.Currentdate)
        str�������� = Format(CDate(str��������), "yyyy-MM-dd")
    Else
        str�������� = Format(zlDataBase.Currentdate, "yyyy-MM-dd")
    End If
    
    str��Ա��� = arrOutput(8)
    str��λ���� = arrOutput(9)
    str��λ���� = arrOutput(10)
    '˫������������Ժ�ڼ䷢������ҵ����ˣ��ڽ�������ҵ��ʱ�����סԺ˳��Ų�Ϊ�գ�˵����Ժ��������˳���
    str��ˮ�� = arrOutput(12)
    mstr������ˮ�� = arrOutput(12)
    If strסԺ˳��� <> "" Then str��ˮ�� = strסԺ˳���
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    strIdentify = str���� & ";" & strҽ���� & ";;" & str���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";" & str��λ���� & "(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    cur�����ʻ� = arrOutput(11)
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";" & str��ˮ��                 '9.˳���
    str���� = str���� & ";" & str��Ա���               '10��Ա���
    str���� = str���� & ";" & arrOutput(11)             '11�ʻ����
    str���� = str���� & ";" & IIf(strסԺ˳��� <> "", "1", "0")                       '12��ǰ״̬
    str���� = str���� & ";"                             '13����ID
    str���� = str���� & ";" & IIf(Left(str��Ա���, 1) = "��", 2, 1)     '14��ְ(1,2)
    str���� = str���� & ";" & str���ı��               '15����֤�� ����ҽ�����ڱ���ҽ�������ı��루���⽨��ҽ�����ģ�
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";" & cur�����ʻ�             '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";"                             '23�������� (1����������)
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
    
    gstrSQL = "Select * From �����ʻ� Where ҽ����='" & strҽ���� & "' And ����=" & gintInsure
    Call OpenRecordset(rsTemp, "ȡ����ID")
    If Not rsTemp.EOF Then
        lng����ID = rsTemp!����ID
    End If
    datCurr = zlDataBase.Currentdate
    If lng����ID <> 0 Then          '��������Ѵ��ڣ����ȡ�ʻ������Ϣ
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, curͳ���޶�)
        bln��ȡ�ʻ������Ϣ = True
    End If
    
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID)
    
    If bln��ȡ�ʻ������Ϣ = True Then          '�����ȡ���ʻ������Ϣ��������д��
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
            cur�����ʻ� & ",0," & _
            cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & curͳ���޶� & ")"
        Call ExecuteProcedure(gstrSysName)
    End If
    
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_�¶� = strIdentify & ";" & lng����ID & str����
        
        mstrҽ���� = strҽ����
        mstr���� = str����
    Else
        mstr������ˮ�� = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�¶�(strSelfNo As String, ByVal bytPlace As Byte) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset, str���� As String, strҽ���� As String, str���� As String
    Dim strReturn As String, str��ˮ�� As String, str���ı�� As String, strInput As String, arrOutput  As Variant
    Dim strҽԺ���� As String
    
    On Error GoTo errHandle
    
    
    If bytPlace = balanԤ�� Then
        '�ڲ�����Ժ���Ԥ��֮��ɱ仯�����Ե��²�������Ѿ���׼ȷ��
        '��ò���ҽ���š������ı�ŵ���Ϣ
        If frmIdentify����.GetIdentify(TYPE_�¶�, str����, strҽ����, str���ı��, str����) = False Then Exit Function
        
        If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
        
        '���������֤
        If Get��ˮ��("A", strҽԺ����, str��ˮ��) = False Then Exit Function
        strInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & str���� & "|11#"
        Call WriteLog("DataUnloading(" & strInput & "," & str��ˮ�� & "," & str���ı�� & ")")
        strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        mstrҽ���� = strҽ����
        mstr���� = str����
        �������_�¶� = Val(arrOutput(11))
    Else
        '�����ݿ��ж�ȡ����Ϊ�ղŲű����˵ģ�Ӧ����׼ȷ�ģ�
        gstrSQL = "Select �ʻ���� From �����ʻ� where ����=" & gintInsure & " and ����=0 and ҽ����='" & strSelfNo & "'"
        Call OpenRecordset(rsTemp, gstrSysName)
        
        If rsTemp.EOF = False Then
            �������_�¶� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_�¶�(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strҽ���� As String, strInput As String, arrOutput  As Variant, strReturn As String
    Dim dbl�����ʻ� As Double
    Dim lng����ID As Long, datCurr As Date, lng��� As Long
    Dim str���ı�� As String, str����˳��� As String, str��Ա��� As String, strҽԺ���� As String
    Dim dbl�ܽ�� As Double, dblȫ�Է� As Double, dbl�ҹ��Ը� As Double, dbl���� As Double, dbl��� As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_�¶� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    datCurr = zlDataBase.Currentdate
    
    '�ӱ����ʻ���õǼ���Ϣ
    gstrSQL = "select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "from �����ʻ� where ����ID=" & lng����ID & " and ����=" & gintInsure
    Call OpenRecordset(rsTemp, "����Ԥ��")
    'str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str����˳��� = mstr������ˮ��
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    str��Ա��� = IIf(IsNull(rsTemp("��Ա���")), "", rsTemp("��Ա���"))
    strҽ���� = rsTemp("ҽ����")
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '��������Ѿ�����ķ�����ϸ
    Set mcol����ϸ = Nothing
    
    'Ȼ����봦����ϸ
    Do Until rs��ϸ.EOF
        '�õ�������ϸ
        gstrSQL = "select A.����,A.����,A.���,A.���,A.���㵥λ,B.��Ŀ����,B.��ע,C.��� as ���� " & _
                 " from �շ�ϸĿ A,����֧����Ŀ B,�շ���� C " & _
                 " where A.���=C.���� and  A.ID=" & rs��ϸ("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID and B.����=" & gintInsure
        Call OpenRecordset(rsTemp, "����Ԥ��")
        
        '���з��÷ָ�
        strReturn = DivideUp(str���ı��, ToVarchar(rsTemp("��Ŀ����"), 12), "11", str��Ա���, Val(rs��ϸ("����")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        '�ڶ�������˳�������Ϊ���㵥��
        strInput = str����˳��� & "|" & str����˳���
        strInput = strInput & "|" & str����˳��� & "_" & lng���      '���
        strInput = strInput & "|" & strҽ���� & "|" & str���ı�� & "|" & strҽԺ���� & "|000"
        strInput = strInput & "|" & ToVarchar(rsTemp("��Ŀ����"), 12)  'ҽ����ˮ��
        strInput = strInput & "|" & ToVarchar(rsTemp("����"), 10)      '�շѴ�������
        strInput = strInput & "|" & Format(rs��ϸ("����"), "0.00")
        strInput = strInput & "|" & Format(rs��ϸ("����"), "0.00")
        strInput = strInput & "|" & Format(rs��ϸ("ʵ�ս��"), "0.00")
        strInput = strInput & "|" & arrOutput(4)                       '�Ը�����
        strInput = strInput & "|" & Format(Val(arrOutput(1)) * rs��ϸ("����"), "#0.00") 'ȫ�ԷѲ���
        strInput = strInput & "|" & Format(Val(arrOutput(2)) * rs��ϸ("����"), "#0.00") '�ҹ��ԷѲ���
        strInput = strInput & "|" & Format(Val(arrOutput(3)) * rs��ϸ("����"), "#0.00") '����������
        strInput = strInput & "||11"                                   '�����־��֧�����
        strInput = strInput & "|" & ToVarchar(UserInfo.����, 56)       '������������
        strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)       '��������ҽ��
        strInput = strInput & "|" & ToVarchar(UserInfo.����, 56)       '�ܵ���������
        strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)       '�ܵ�����ҽ��
        strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)        '������
        strInput = strInput & "|" & Format(datCurr + lng��� / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '����ʱ��
        strInput = strInput & "|" & ToVarchar(rsTemp("����"), 200)       '�շ���Ŀ
        strInput = strInput & "|" & ToVarchar(rsTemp("���"), 200)       '���
        strInput = strInput & "|"                                        '����
        strInput = strInput & "|" & ToVarchar(rsTemp("���㵥λ"), 30)    '��λ
        strInput = strInput & "|||"                                      'Ӣ��������ѧ��
        strInput = strInput & lng��� & "#"                             '���
        Call WriteLog(strInput)
        mcol����ϸ.Add strInput  '���Ƚ���ϸ���棬������ʱ���ϴ�
        
        lng��� = lng��� + 1
        dbl�ܽ�� = dbl�ܽ�� + Val(rs��ϸ("ʵ�ս��"))
        dblȫ�Է� = dblȫ�Է� + Val(arrOutput(1)) * rs��ϸ("����")
        dbl�ҹ��Ը� = dbl�ҹ��Ը� + Val(arrOutput(2)) * rs��ϸ("����")
        dbl���� = dbl���� + Val(arrOutput(3)) * rs��ϸ("����")    'Ŀǰʹ�����������֡�����
        
        rs��ϸ.MoveNext
    Loop
    
    '�õ��������
    dbl��� = �������_�¶�(strҽ����, balan����)
    With g��������
        .�������ý�� = dbl�ܽ��
        .ȫ�Էѽ�� = dblȫ�Է�
        .�����Ը���� = dbl�ҹ��Ը�
        .����ͳ���� = dbl����
        .֧��˳��� = str����˳���
    End With
    '����Ԥ����
    strReturn = GetPayCount(str���ı��, "11", dbl�ܽ��, dblȫ�Է�, dbl�ҹ��Ը�, dbl����, dbl���)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    dbl�����ʻ� = Val(arrOutput(1))                 'ȡ�ӿ������ʻ�֧���Ľ��
    If mint���õ���_�¶� = 0 Then
        '˫������ȫ�����ʻ�֧�����ӿڷ��ص��ʻ�֧��������Ч
        dbl�����ʻ� = IIf(dbl��� < dbl�ܽ��, dbl���, dbl�ܽ��)
    End If
    str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";1"   '�����޸ĸ����ʻ�
    �����������_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�¶�(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim strҽ���� As String, strInput As String, arrOutput  As Variant, strReturn As String
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset
    Dim datCurr As Date, var��ϸ As Variant, rsTemp As New ADODB.Recordset
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, str���� As String, str��ˮ�� As String
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From ���˷��ü�¼ Where ����ID=" & lng����ID
    Call OpenRecordset(rs��ϸ, "����Ԥ��")
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    lng����ID = rs��ϸ("����ID")
    datCurr = rs��ϸ("�Ǽ�ʱ��")
    
    If mstrҽ���� <> strSelfNo Then
        MsgBox "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    '����ʻ������Ϣ
    gstrSQL = "select ����,ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "from �����ʻ� where ����ID=" & lng����ID & " and ����=" & gintInsure
    Call OpenRecordset(rs��ϸ, "����Ԥ��")
    str����˳��� = mstr������ˮ��
    str���ı�� = IIf(IsNull(rs��ϸ("���ı��")), "", rs��ϸ("���ı��"))
    str���� = IIf(IsNull(rs��ϸ("����")), "", rs��ϸ("����")) '���뿨��û�п���
    strҽ���� = rs��ϸ("ҽ����")
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '�ϴ�������ϸ��ͳһ��һ����ˮ�ţ�������
    If Get��ˮ��("G", strҽԺ����, str��ˮ��) = False Then Exit Function
    For Each var��ϸ In mcol����ϸ
        Call WriteLog("�ϴ�:" & var��ϸ)
        strReturn = DataUnloading(var��ϸ, str��ˮ��, str���ı��)
        Call WriteLog("����:" & strReturn)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    Next
    
    '���ý���
    With g��������
    Call WriteLog("����(" & str���� & "," & strҽ���� & "," & str���ı�� & "," & mstr���� & "," & str����˳��� & "," & "11" & "," & strҽԺ���� & "," & "000" & "," & CStr(cur�����ʻ�) & "," & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "," & _
               CStr(.�������ý��) & "," & CStr(.ȫ�Էѽ��) & "," & CStr(.�����Ը����) & "," & CStr(.����ͳ����) & "," & ToVarchar(UserInfo.����, 20) & "," & ToVarchar(.֧��˳���, 20) & ")")
    strReturn = reckoning(str����, strҽ����, str���ı��, mstr����, str����˳���, "11", strҽԺ����, "000", Format(cur�����ʻ�, "0.##"), Format(datCurr, "yyyy-MM-dd HH:mm:ss"), _
               Format(.�������ý��, "0.##"), Format(.ȫ�Էѽ��, "0.##"), Format(.�����Ը����, "0.##"), Format(.����ͳ����, "0.##"), ToVarchar(UserInfo.����, 20), ToVarchar(.֧��˳���, 20))
    Call WriteLog("����:" & strReturn)
    End With
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '��������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim cur�����ۼ� As Currency, cur�������� As Currency, curͳ���޶� As Currency
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, curͳ���޶�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & curͳ���޶� & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & g��������.�������ý�� & ",0,0," & _
        0 & "," & 0 & ",0,0," & cur�����ʻ� & ",'')"
    Call ExecuteProcedure(gstrSysName)
    '---------------------------------------------------------------------------------------------

    �������_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_�¶�(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��

    ����������_�¶� = True
End Function

Public Function �����ʻ�תԤ��_�¶�(lngԤ��ID As Long, cur�����ʻ� As Currency, strSelfNo As String, str˳��� As String, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim strҽ���� As String, strInput As String, arrOutput  As Variant, strReturn As String
    Dim datCurr As Date, var��ϸ As Variant, rs��ϸ As New ADODB.Recordset
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, str���� As String, str��ˮ�� As String
    
    On Error GoTo errHandle
    
    datCurr = zlDataBase.Currentdate
    
    If mstrҽ���� <> strSelfNo Then
        MsgBox "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    '����ʻ������Ϣ
    gstrSQL = "select ����,ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "from �����ʻ� where ����ID=" & lng����ID & " and ����=" & gintInsure
    Call OpenRecordset(rs��ϸ, "��Ԥ����")
    str����˳��� = IIf(IsNull(rs��ϸ("�������")), "", rs��ϸ("�������"))
    str���ı�� = IIf(IsNull(rs��ϸ("���ı��")), "", rs��ϸ("���ı��"))
    str���� = IIf(IsNull(rs��ϸ("����")), "", rs��ϸ("����")) '���뿨û�п���
    strҽ���� = rs��ϸ("ҽ����")
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '�����жϽ���Ƿ����ʹ��
    strReturn = GetPayCount(str���ı��, "31", cur�����ʻ�, 0, 0, cur�����ʻ�, cur�����ʻ�)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    If Val(arrOutput(1)) < cur�����ʻ� Then
        MsgBox "�����ʻ���������֧��Ԥ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ý���
    Call WriteLog("reckoning(" & str���� & "," & strҽ���� & "," & str���ı�� & "," & mstr���� & "," & str����˳��� & "," & "31" & "," & strҽԺ���� & "," & "000" & "," & CStr(cur�����ʻ�) & "," & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "," & _
               CStr(cur�����ʻ�) & "," & CStr(0) & "," & CStr(0) & "," & CStr(cur�����ʻ�) & "," & ToVarchar(UserInfo.����, 20) & "," & ToVarchar(str����˳���, 20) & ")")
    strReturn = reckoning(str����, strҽ����, str���ı��, mstr����, str����˳���, "31", strҽԺ����, "000", CDbl(cur�����ʻ�), Format(datCurr, "yyyy-MM-dd HH:mm:ss"), _
               CDbl(cur�����ʻ�), CDbl(0), CDbl(0), CDbl(cur�����ʻ�), ToVarchar(UserInfo.����, 20), ToVarchar(str����˳���, 20))
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '��������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
            
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & cur�����ʻ� & ",0,0," & _
        0 & "," & 0 & ",0,0," & cur�����ʻ� & ",'')"
    Call ExecuteProcedure(gstrSysName)
    '---------------------------------------------------------------------------------------------

    �����ʻ�תԤ��_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�¶�(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String, Optional ByVal blnFirst As Boolean = True) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strInput As String, arrOutput  As Variant, arrTmp As Variant
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str��ˮ�� As String, strReturn As String
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, str���� As String
    Dim str��Ժ��� As String, str��Ժ��� As String, strҽԺ��� As String, str��Ժ��� As String
    Dim intValue As Integer
    Dim dblͳ���޶� As Double, dblͳ���ۼ� As Double

    On Error GoTo errHandle
    
    '��ȡ���ղ���ֵ���Ծ���ҽ��������Ժʱ���Ƿ�ͬʱ�ϴ���Ժ��Ϣ
    intValue = 1
'    gstrSQL = "Select Nvl(����ֵ,0) Value From ���ղ��� Where ����=" & TYPE_�¶� & " And ������='�ϴ���Ժ��Ϣ'"
'    Call OpenRecordset(rsTemp, "��ȡ�ϴ���Ժ��Ϣ����ֵ")
'
'    If Not rsTemp.EOF Then
'        intValue = rsTemp!Value
'    End If
    
    '���ҽ����
    gstrSQL = "select ҽ����,����,˳��� as �������,����֤�� as ���ı�� from �����ʻ� where ����=" & TYPE_�¶� & " and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    
    str���� = IIf(IsNull(rsTemp("����")), "", rsTemp("����")) '��������뿨,���ž�Ϊ��
    strҽ���� = rsTemp("ҽ����")
    str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.�������,A.������Ϣ from ������ A where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & _
              " and A.������� in (1,3) and A.��ϴ���=1"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    Do Until rsTemp.EOF
        If rsTemp("�������") = 1 Then
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        Else
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        End If
        rsTemp.MoveNext
    Loop
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
    
    '���������Ժ��Ϣ
    datCurr = zlDataBase.Currentdate
    gstrSQL = " select A.��Ժ����,A.�Ǽ�ʱ��,B.���� ��Ժ���� " & _
              " from ������ҳ A,���ű� B " & _
              " Where A.��Ժ����ID=B.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    str��Ժ��� = Year(rsTemp!��Ժ����)
    
    '���ҽԺ���
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Function
    
    '������Ժ�ӿ�
    If blnFirst Then
        '�������
        If Get��ˮ��("C", strҽԺ����, str��ˮ��) = False Then Exit Function
        strInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & mstr���� & _
                    "|" & str����˳��� & "|" & strҽԺ���� & _
                    "|000|0|000|31|0" & _
                    "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss") & _
                    "|" & ToVarchar(UserInfo.����, 20) & _
                    "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "#"
        Call WriteLog("DataUnloadint(" & strInput & "," & str��ˮ�� & "," & str���ı�� & ")")
        strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        '�����˫����ҽ������Ҫ�Ի���ͳ���޶�����жϣ����Ϊ�㣬���ֹ������Ժ����ʾ����ͨ���˰���ͬʱ��
        '��Ҫ������ͳ���޶���ͳ��֧���ۼ���ʾ����������Ա
        If mint���õ���_�¶� = 0 Then
            If Val(arrOutput(6)) = 0 Then
                MsgBox "����ͳ���޶�Ϊ�㣬����������ҽ�������Ժ���밴��ͨ���˰�����Ժ��", vbInformation, gstrSysName
                Exit Function
            End If
            dblͳ���޶� = Val(arrOutput(6))
            dblͳ���ۼ� = Val(arrOutput(8))
        End If
        
        '�ϴ���Ժ�Ǽ�
        If intValue = 1 Then
            If Get��ˮ��("E", strҽԺ����, str��ˮ��) = False Then Exit Function
            strInput = str����˳��� & "|" & strҽ���� & "|" & strҽԺ���� & "|000|" & strҽԺ��� & "|31|0"
            strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)    '��Ժ������
            strInput = strInput & "|" & ToVarchar(rsTemp("��Ժ����"), 20)  '��Ժ����
            strInput = strInput & "|" & str��Ժ���
            strInput = strInput & "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")
            strInput = strInput & "|" & Format(rsTemp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss") & "|||��Ժ�Ǽ�|" & Format("2000-01-01", "yyyy-MM-dd HH:mm:ss") & "|" & Format("2000-01-01", "yyyy-MM-dd HH:mm:ss") & "|9#"
            Call WriteLog("�ϴ���Ժ�Ǽ�(" & strInput & "," & str��ˮ�� & "," & str���ı�� & ")")
            strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
            Call WriteLog("����:" & strReturn)
            If JudgeReturn(strReturn, arrTmp) = False Then Exit Function
        End If
        
        Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
        Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
        Dim intסԺ�����ۼ� As Integer
                
        '�ʻ������Ϣ   ע����ֶ���ʵ���ô�֮��Ķ�Ӧ��ϵ
        '��������    ----   סԺ����
        '�����ۼ�    ----   ����ͳ��֧���ۼ�
        '����ͳ���޶�  ----   סԺͳ���޶�
        '���ͳ���޶�  ----   ʵ������
        '���ͳ���ۼ�  ----   ͳ�ﱨ������
        Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & str��Ժ��� & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & _
            arrOutput(5) & "," & arrOutput(8) & "," & arrOutput(6) & "," & arrOutput(3) & "," & arrOutput(11) & ")"
        Call ExecuteProcedure(gstrSysName)
        
        '����״̬���޸�
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
        Call ExecuteProcedure(gstrSysName)
        
        If mint���õ���_�¶� = 0 Then
            MsgBox "�òα����˵�סԺ�����Ϣ��" & vbCrLf & _
                   "    ����ͳ���޶��" & Format(dblͳ���޶�, "#0.00") & _
                   "    ͳ��֧���ۼƣ���" & Format(dblͳ���ۼ�, "#0.00"), vbInformation, gstrSysName
        End If
    End If
    
    ��Ժ�Ǽ�_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�¶�(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str��Ժ��� As String, str��Ժ��� As String
    Dim strInput As String, arrOutput  As Variant, str��ˮ�� As String, strReturn As String
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, strҽ���� As String
    Dim strҽԺ��� As String
    
    On Error GoTo errHandle
    
    '���ҽ����
    gstrSQL = "select ҽ����,����,˳��� as �������,����֤�� as ���ı�� from �����ʻ� where ����=" & TYPE_�¶� & " and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    strҽ���� = rsTemp("ҽ����")
    str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.�������,A.������Ϣ from ������ A where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & _
              " and A.������� in (1,3) and A.��ϴ���=1"
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    Do Until rsTemp.EOF
        If rsTemp("�������") = 1 Then
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        Else
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        End If
        rsTemp.MoveNext
    Loop
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
        
    '���������Ժ��Ϣ
    datCurr = zlDataBase.Currentdate
    gstrSQL = "select A.����ҽʦ,A.סԺҽʦ,A.�Ǽ�ʱ��,A.��Ժ����,A.��Ժ����,A.��Ժ��ʽ,B.���� as ��Ժ����,C.���� as ��Ժ���� " & _
             " from ������ҳ A,���ű� B,���ű� C " & _
             " Where A.��Ժ����ID = B.ID And A.��Ժ����ID = C.ID And A.����ID =" & lng����ID & " And A.��ҳID = " & lng��ҳID
    Call OpenRecordset(rsTemp, "��Ժ�Ǽ�")
    
    '���ҽԺ���
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Function

    '���ó�Ժ�ӿ�
    If Get��ˮ��("E", strҽԺ����, str��ˮ��) = False Then Exit Function
    strInput = str����˳��� & "|" & strҽ���� & "|" & strҽԺ���� & "|000|" & strҽԺ��� & "|31|" & _
                IIf(Format(rsTemp("��Ժ����"), "yyyy") = Format(rsTemp("��Ժ����"), "yyyy"), "0", "1")
    strInput = strInput & "|" & ToVarchar(rsTemp("����ҽʦ"), 20)  '��Ժ������
    strInput = strInput & "|" & ToVarchar(rsTemp("��Ժ����"), 20)  '��Ժ����
    strInput = strInput & "|" & str��Ժ���
    strInput = strInput & "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")
    strInput = strInput & "|" & Format(rsTemp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
    strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)       '��Ժ������
    strInput = strInput & "|" & ToVarchar(rsTemp("��Ժ����"), 20)  '��Ժ����
    strInput = strInput & "|" & str��Ժ���
    strInput = strInput & "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")
    strInput = strInput & "|" & Format(zlDataBase.Currentdate, "yyyy-MM-dd HH:mm:ss") '��Ժ����ʱ��
    strInput = strInput & "|" & Switch(rsTemp("��Ժ��ʽ") = "����", 0, rsTemp("��Ժ��ʽ") = "����", 1, rsTemp("��Ժ��ʽ") = "תԺ", 2, True, 9) & "#"
    
    Call WriteLog("DataUnloadint(" & strInput & "," & str��ˮ�� & "," & str���ı�� & ")")
    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    
    ��Ժ�Ǽ�_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ���ʴ���_�¶�(strNO As String, int���� As Integer, int״̬ As Integer, Optional lng����ID As Long) As Boolean
'���ܣ���סԺ���˵ļ��ʵ����ϴ���ҽ��ǰ�÷�����
'������lng����ID=�Ƿ�ֻ�ϴ�������ָ�����˵ķ���
    Dim strInput As String, arrOutput   As Variant, strReturn As String
    Dim rsBill As New ADODB.Recordset, rsTemp As New ADODB.Recordset, rs�շ���� As New ADODB.Recordset
    Dim lng��ǰ���� As Long
    '���ô���ʹ�õı���
    Dim str���ı�� As String, str����˳��� As String, str��Ա��� As String, strҽԺ���� As String
    Dim str��ˮ�� As String, str�շ���� As String, strҽ���� As String
    
    ���ʴ���_�¶� = True '���ȱ�֤�����ܵõ����档��ʹ�����ϴ��ܣ�Ҳ�������Ժ�����ϴ���
    On Error GoTo errHandle
    
    '�г������շ����
    gstrSQL = "Select ����,��� as ���� From �շ����"
    Call OpenRecordset(rs�շ����, gstrSysName)
    
    '��ȡ������ϸ(ҽ����,˳���,�Ǽ�ʱ��,��Ŀ����,��Ŀ����,����,���,����,����,���,ҽ��,��������)
    '�����зǸ�ҽ���ķ��ò���,δ����ҽ������Ĳ���,��˳��ŵĲ���,Ӥ���Ѳ��ϴ�������������
    gstrSQL = _
        "Select Nvl(A.�۸񸸺�,���) as ���," & _
        " A.����ID,A.��ҳID,F.ҽ����,F.˳���,A.�Ǽ�ʱ��,D.��Ŀ����,B.���� as ��Ŀ����,A.�շ����, " & _
        " Decode(Instr(B.���,'��'),0,B.���,Substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
        " Decode(Instr(B.���,'��'),0,'',Substr(B.���,Instr(B.���,'��')+1)) as ����," & _
        " Avg(Nvl(A.����,1)*A.����) as ����,Sum(A.��׼����) as ����,Sum(A.ʵ�ս��) as ���," & _
        " A.������ as ҽ��,C.���� as ��������" & _
        " From ���˷��ü�¼ A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,������ҳ E,�����ʻ� F" & _
        " Where A.��¼״̬<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID" & _
        " And A.����ID=E.����ID And A.��ҳID=E.��ҳID And A.����ID=F.����ID" & _
        " And F.˳��� is Not NULL And Nvl(A.Ӥ����,0)=0" & _
        " And D.����=" & gintInsure & " And E.����=" & gintInsure & " And F.����=" & gintInsure & _
        " And A.NO='" & strNO & "' And A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & _
        IIf(lng����ID = 0, "", " And A.����ID=" & lng����ID) & _
        " Group by Nvl(A.�۸񸸺�,���),A.����ID,A.��ҳID,F.ҽ����,F.˳���," & _
        " A.�Ǽ�ʱ��,D.��Ŀ����,B.����,A.�շ����,B.���,A.������,C.����" & _
        " Order by ����ID,���"
    rsBill.CursorLocation = adUseClient
    Call OpenRecordset(rsBill, "���ʴ���")
    
    Do Until rsBill.EOF
        '���ʵ����ж������,Ҫ�ֱ���
        If lng��ǰ���� <> rsBill("����ID") Then
            '�Ըò�������Ӧ�ĳ�ʼ������-------------------------------------------------
            lng��ǰ���� = rsBill("����ID")
            
            '�õ���Ժ������Ϣ���Ѿ�������֤�ģ�
            gstrSQL = "Select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
                      "       ,NVL(A.��������,0) as סԺ����,NVL(A.�����ۼ�,0) as ����ͳ��֧���ۼ�" & _
                      "       ,NVL(A.����ͳ���޶�,0) as סԺͳ���޶�,NVL(A.���ͳ���޶�,0) as ʵ������,NVL(A.���ͳ���ۼ�,0) as ͳ�ﱨ������" & _
                      "  From �ʻ������Ϣ A,������ҳ B,�����ʻ� C " & _
                      "  where B.����ID=" & lng��ǰ���� & " and B.��ҳID=" & rsBill("��ҳID") & " and A.����ID=B.����ID and A.����=" & gintInsure & " and A.���=to_char(B.��Ժ����,'yyyy')" & _
                      "     and C.����ID=A.����ID and C.����=A.����"
            Call OpenRecordset(rsTemp, "���ʴ���")
            str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
            str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
            strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
            str��Ա��� = IIf(IsNull(rsTemp("��Ա���")), "", rsTemp("��Ա���"))
            
            If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
            If Get��ˮ��("G", strҽԺ����, str��ˮ��) = False Then Exit Function
        End If
            
        '���з��÷ָ�
        Call WriteLog("DivideUp(" & str���ı�� & "," & ToVarchar(rsBill!��Ŀ����, 12) & "," & "31" & "," & str��Ա��� & "," & Val(rsBill!����) & ")")
        strReturn = DivideUp(str���ı��, ToVarchar(rsBill("��Ŀ����"), 12), "31", str��Ա���, Val(rsBill("����")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        rs�շ����.Filter = "���� = '" & rsBill("�շ����") & "'"
        If rs�շ����.EOF = False Then str�շ���� = rs�շ����("����")
        
        '�ڶ�������˳�������Ϊ���㵥��
        strInput = str����˳��� & "|" & str����˳���
        strInput = strInput & "|" & strNO & "_" & rsBill("���") & "_" & int���� & "_" & int״̬  '���
        strInput = strInput & "|" & strҽ���� & "|" & str���ı�� & "|" & strҽԺ���� & "|000"
        strInput = strInput & "|" & ToVarchar(rsBill("��Ŀ����"), 12)  'ҽ����ˮ��
        strInput = strInput & "|" & ToVarchar(str�շ����, 10)      '�շѴ�������
        strInput = strInput & "|" & Format(rsBill("����"), "0.00")
        strInput = strInput & "|" & Format(rsBill("����"), "0.00")
        strInput = strInput & "|" & Format(rsBill("���"), "0.00")
        strInput = strInput & "|" & arrOutput(4)                       '�Ը�����
        strInput = strInput & "|" & Format(Val(arrOutput(1)) * rsBill("����"), "#0.00") 'ȫ�ԷѲ���
        strInput = strInput & "|" & Format(Val(arrOutput(2)) * rsBill("����"), "#0.00") '�ҹ��ԷѲ���
        strInput = strInput & "|" & Format(Val(arrOutput(3)) * rsBill("����"), "#0.00") '����������
        strInput = strInput & "||31"                                   '�����־��֧�����
        strInput = strInput & "|" & ToVarchar(rsBill("��������"), 56)  '������������
        strInput = strInput & "|" & ToVarchar(rsBill("ҽ��"), 20)      '��������ҽ��
        strInput = strInput & "|" & ToVarchar(rsBill("��������"), 56)  '�ܵ���������
        strInput = strInput & "|" & ToVarchar(rsBill("ҽ��"), 20)      '�ܵ�����ҽ��
        strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)        '������
        strInput = strInput & "|" & Format(rsBill("�Ǽ�ʱ��") + rsBill("���") / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '����ʱ��
        strInput = strInput & "|" & ToVarchar(rsBill("��Ŀ����"), 200)       '�շ���Ŀ
        strInput = strInput & "|" & ToVarchar(rsBill("���"), 200)       '���
        strInput = strInput & "|" & ToVarchar(rsBill("����"), 200)       '����
        strInput = strInput & "|"                                        '��λ
        strInput = strInput & "||"                                      'Ӣ��������ѧ��
        'modify by ccy ,Ψһ
        strInput = strInput & Format(rsBill("�Ǽ�ʱ��"), "yyyyMMddHHmmss") & rsBill("���") & "#"
        
        Call WriteLog("DataUnloading(" & strInput & "," & str��ˮ�� & "," & str���ı�� & ")")
        strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        
        gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & strNO & "'," & rsBill("���") & "," & int���� & "," & int״̬ & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        rsBill.MoveNext
    Loop
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�¶�(rsExse As Recordset, ByVal lng����ID As Long, ByVal strҽ���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs�շ���� As New ADODB.Recordset

    Dim strInput As String, arrOutput   As Variant, strReturn As String
    Dim str���ı�� As String, str����˳��� As String, str��Ա��� As String, strҽԺ���� As String
    Dim cur�����ʻ� As Double, curͳ��֧�� As Double
    Dim dbl�ܽ�� As Double, dblȫ�Է� As Double, dbl�ҹ��Ը� As Double, dbl������ As Double
    Dim dblסԺ���� As Double, dbl����ͳ��֧���ۼ� As Double, dblסԺͳ���޶� As Double, lngʵ������ As Long, dblͳ�ﱨ������ As Double
    Dim strҽ�� As String, datCurr As Date, str��ˮ�� As String, str�շ���� As String
    
    On Error GoTo errHandle
    mlng����ID = 0         '��ʼ����ֻҪһѡ���ˣ��ͻ���ñ����̣�Ҳ�ͻ����0
    
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    rsExse.MoveFirst
    
    datCurr = zlDataBase.Currentdate
    With g��������
        .����ID = rsExse("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & rsExse("����ID")
        Call OpenRecordset(rsTemp, "�������")
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
    End With
    
    '���½��д������
    Dim str����_New As String, strҽ����_New As String, str���ı��_New As String, str����_New As String
    If frmIdentify����.GetIdentify(TYPE_�¶�, str����_New, strҽ����_New, str���ı��_New, str����_New) = False Then
        '�����֤δͨ��
        Exit Function
    End If
'    If strҽ���� <> strҽ����_New Then
'        MsgBox "�ÿ����ǵ�ǰ���˵ģ�����һ�¡�", vbInformation, gstrSysName
'        Exit Function
'    End If
    If ��Ժ�Ǽ�_�¶�(g��������.����ID, g��������.��ҳID, strҽ����, False) = False Then
        Exit Function
    End If
    
    '�õ���Ժ������Ϣ���Ѿ�������֤�ģ�
    gstrSQL = "Select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "       ,NVL(A.��������,0) as סԺ����,NVL(A.�����ۼ�,0) as ����ͳ��֧���ۼ�" & _
              "       ,NVL(A.����ͳ���޶�,0) as סԺͳ���޶�,NVL(A.���ͳ���޶�,0) as ʵ������,NVL(A.���ͳ���ۼ�,0) as ͳ�ﱨ������" & _
              "  From �ʻ������Ϣ A,������ҳ B,�����ʻ� C " & _
              "  where B.����ID=" & lng����ID & " and B.��ҳID=" & g��������.��ҳID & " and A.����ID=B.����ID and A.����=" & gintInsure & " and A.���=to_char(B.��Ժ����,'yyyy')" & _
              "     and C.����ID=A.����ID and C.����=A.����"
    Call OpenRecordset(rsTemp, "סԺԤ��")
    str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    str��Ա��� = IIf(IsNull(rsTemp("��Ա���")), "", rsTemp("��Ա���"))
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    dblסԺ���� = rsTemp("סԺ����")
    dbl����ͳ��֧���ۼ� = rsTemp("����ͳ��֧���ۼ�")
    dblסԺͳ���޶� = rsTemp("סԺͳ���޶�")
    lngʵ������ = rsTemp("ʵ������")
    dblͳ�ﱨ������ = rsTemp("ͳ�ﱨ������")
    
    '�г������շ����
    gstrSQL = "Select ����,��� as ���� From �շ����"
    Call OpenRecordset(rs�շ����, gstrSysName)
    '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
    cn�ϴ�.ConnectionString = gcnOracle.ConnectionString
    cn�ϴ�.Open
    
    Screen.MousePointer = vbHourglass
    
    
    If Get��ˮ��("G", strҽԺ����, str��ˮ��) = False Then Exit Function
    Do Until rsExse.EOF
        '���з��÷ָ�
        Call WriteLog("���÷ָ�(" & str���ı�� & "," & ToVarchar(rsExse!ҽ����Ŀ����, 12) & ",31," & str��Ա��� & "," & Val(rsExse!�۸�) & ")")
        strReturn = DivideUp(str���ı��, ToVarchar(rsExse("ҽ����Ŀ����"), 12), "31", str��Ա���, Val(rsExse("�۸�")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        dbl�ܽ�� = dbl�ܽ�� + rsExse("���")
        dblȫ�Է� = dblȫ�Է� + Val(arrOutput(1)) * rsExse("����")
        dbl�ҹ��Ը� = dbl�ҹ��Ը� + Val(arrOutput(2)) * rsExse("����")
        dbl������ = dbl������ + Val(arrOutput(3)) * rsExse("����")
        
'        If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" Then
'            'ֻ�ϴ�ֻ���ݹ�������
'            rs�շ����.Filter = "���� = '" & rsExse("�շ����") & "'"
'            If rs�շ����.EOF = False Then str�շ���� = rs�շ����("����")
'
'            '�ڶ�������˳�������Ϊ���㵥��
'            strInput = str����˳��� & "|" & str����˳���
'            strInput = strInput & "|" & rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") & "_" & rsExse("��¼״̬")  '���
'            strInput = strInput & "|" & strҽ���� & "|" & str���ı�� & "|" & strҽԺ���� & "|000"
'            strInput = strInput & "|" & ToVarchar(rsExse("ҽ����Ŀ����"), 12)  'ҽ����ˮ��
'            strInput = strInput & "|" & ToVarchar(str�շ����, 10)      '�շѴ�������
'            strInput = strInput & "|" & Format(rsExse("����"), "0.00")
'            strInput = strInput & "|" & Format(rsExse("�۸�"), "0.00")
'            strInput = strInput & "|" & Format(rsExse("���"), "0.00")
'            strInput = strInput & "|" & arrOutput(4)                       '�Ը�����
'            strInput = strInput & "|" & Val(arrOutput(1)) * rsExse("����") 'ȫ�ԷѲ���
'            strInput = strInput & "|" & Val(arrOutput(2)) * rsExse("����") '�ҹ��ԷѲ���
'            strInput = strInput & "|" & Val(arrOutput(3)) * rsExse("����") '����������
'            strInput = strInput & "||31"                                   '�����־��֧�����
'            strInput = strInput & "|" & ToVarchar(rsExse("��������"), 56)  '������������
'            strInput = strInput & "|" & ToVarchar(rsExse("ҽ��"), 20)      '��������ҽ��
'            strInput = strInput & "|" & ToVarchar(rsExse("��������"), 56)  '�ܵ���������
'            strInput = strInput & "|" & ToVarchar(rsExse("ҽ��"), 20)      '�ܵ�����ҽ��
'            strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)        '������
'            strInput = strInput & "|" & Format(rsExse("�Ǽ�ʱ��") + rsExse("���") / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '����ʱ��
'            strInput = strInput & "|" & ToVarchar(rsExse("�շ�����"), 200)       '�շ���Ŀ
'            strInput = strInput & "|" & ToVarchar(rsExse("���"), 200)       '���
'            strInput = strInput & "|" & ToVarchar(rsExse("����"), 200)       '����
'            strInput = strInput & "|"                                        '��λ
'            strInput = strInput & "||"                                      'Ӣ��������ѧ��
'            'modify by ccy ,Ψһ
'            strInput = strInput & Format(rsExse("�Ǽ�ʱ��"), "yyyyMMddHHmmss") & rsExse("���") & "#"
'
'            Call WriteLog("DataUnloading(" & strInput & "," & str��ˮ�� & "," & str���ı�� & ")")
'            strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
'            If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
'
'
'            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
'            cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
'        End If
        
        rsExse.MoveNext
    Loop
    
    '����Ԥ����
    '2107,404.2,44020,0,37,0,0,1604,103,400,.824
    Call WriteLog("Ԥ����:" & dbl�ܽ�� & "," & dblסԺ���� & "," & dblסԺͳ���޶� & "," & dbl����ͳ��֧���ۼ� & "," & lngʵ������ & "," & 0 & "," & 0 & "," & _
                dbl������ & "," & dblȫ�Է� & "," & dbl�ҹ��Ը� & "," & dblͳ�ﱨ������)
    strReturn = CalculateFeeCD(dbl�ܽ��, dblסԺ����, dblסԺͳ���޶�, dbl����ͳ��֧���ۼ�, lngʵ������, 0, 0, _
                dbl������, dblȫ�Է�, dbl�ҹ��Ը�, dblͳ�ﱨ������)
    Call WriteLog("����:" & strReturn)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    curͳ��֧�� = Val(arrOutput(2))
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = dbl�ܽ��
        .ʵ������ = Val(arrOutput(1))
        .ͳ�ﱨ����� = curͳ��֧��
        .�����Ը���� = Val(arrOutput(4))
    
        .����ͳ���� = dbl������
        .ȫ�Էѽ�� = dblȫ�Է�
        .�����Ը���� = dbl�ҹ��Ը�
        .�����ʻ�֧�� = Val(arrOutput(3)) '����ͳ���Ը�����
    End With
    
    סԺ�������_�¶� = "ҽ������;" & curͳ��֧�� & ";0"
    
    mlng����ID = lng����ID  '��ʾ�ò����Ѿ��������������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�¶�(lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant, str��ˮ�� As String, strReturn As String
    Dim strҽԺ��� As String, strҽԺ���� As String, strҽ���� As String
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, datCurr As Date, cur���� As Currency
    
    If mlng����ID <> lng����ID Then
        MsgBox "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    datCurr = zlDataBase.Currentdate
    
    '�õ���Ժ������Ϣ
    gstrSQL = "Select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���,C.��λ����  " & _
              " ,D.����,D.�Ա�,D.�������� " & _
              " ,nvl(A.��������,0) as סԺ����,nvl(A.�����ۼ�,0) as ����ͳ��֧���ۼ�,nvl(A.����ͳ���޶�,0) as סԺͳ���޶�,nvl(A.���ͳ���޶�,0) as ʵ������,nvl(A.���ͳ���ۼ�,0) as ͳ�ﱨ������" & _
              "  From �ʻ������Ϣ A,������ҳ B,�����ʻ� C,������Ϣ D " & _
              "  where B.����ID=" & lng����ID & " and B.��ҳID=" & g��������.��ҳID & " and A.����ID=B.����ID and A.����=" & gintInsure & " and A.���=to_char(B.��Ժ����,'yyyy')" & _
              "     and C.����ID=A.����ID and C.����=A.����   and B.����ID=D.����ID"
    Call OpenRecordset(rsTemp, "סԺԤ��")
    If GetҽԺ����(strҽԺ����, rsTemp("���ı��")) = False Then Exit Function
    If GetҽԺ����(strҽԺ���, rsTemp("���ı��"), True) = False Then Exit Function
    
    cur���� = rsTemp("סԺ����")
    '���ý���
    If Get��ˮ��("F", strҽԺ����, str��ˮ��) = False Then Exit Function
    strInput = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))  '����˳���
    strInput = strInput & "|" & ToVarchar(rsTemp("���ı��"), 4)        '�����ı��
    strInput = strInput & "|" & ToVarchar(rsTemp("ҽ����"), 20)         '���˱���
    strInput = strInput & "|" & ToVarchar(rsTemp("��λ����"), 12)        '��λ����
    strInput = strInput & "|" & ToVarchar(rsTemp("����"), 20)            '����
    strInput = strInput & "|" & ToVarchar(IIf(rsTemp("�Ա�") = "Ů", "2", "1"), 4)         '�Ա�
    strInput = strInput & "|" & Format(rsTemp("��������"), "yyyy-MM-dd") '��������
    strInput = strInput & "|" & Format(rsTemp("ʵ������"), "0")         'ʵ������
    strInput = strInput & "|"                                           '�ɷ�����
    strInput = strInput & "|" & strҽԺ����
    strInput = strInput & "|000"                                        '��Ժ����
    strInput = strInput & "|" & strҽԺ���                             'ҽԺ���
    strInput = strInput & "|31"                                         '֧�����
    strInput = strInput & "|0"                                          '���ֲ���־
    strInput = strInput & "|000"                                        '���ֲ�����
    strInput = strInput & "|" & ToVarchar(rsTemp("�������"), 20)       '������
    strInput = strInput & "|"                                           '�˵����
    strInput = strInput & "|" & ToVarchar(rsTemp("��Ա���"), 20)       'ҽ����Ա���
    With g��������
        strInput = strInput & "|" & Format(cur����, "0.00")        '����
        strInput = strInput & "|" & Format(.�������ý��, "0.00")    '�����ܶ�
        strInput = strInput & "|" & Format(.ȫ�Էѽ��, "0.00")      'ȫ�ԷѲ���
        strInput = strInput & "|" & Format(.�����Ը����, "0.00")    '�ҹ��Ը�����
        strInput = strInput & "|" & Format(.����ͳ����, "0.00")    '����������
        strInput = strInput & "|" & Format(.ʵ������, "0.00")      '�������߲���
        strInput = strInput & "|" & Format(.ͳ�ﱨ�����, "0.00")    '����ҽ��ͳ��֧������
        strInput = strInput & "|" & Format(.�����ʻ�֧��, "0.00")    '����ҽ��ͳ���Ը�����
        strInput = strInput & "|" & Format(0, "0.00")                '����ͳ��֧������
        strInput = strInput & "|" & Format(0, "0.00")                '����ͳ���Ը�����
        strInput = strInput & "|" & Format(.�����Ը����, "0.00")    '�����Ը����
        strInput = strInput & "|" & Format(0, "0.00")                '�����˻�֧�����
    End With
    strInput = strInput & "|"                                              '��Ʊ��
    strInput = strInput & "|" & ToVarchar(UserInfo.����, 20)               '������
    strInput = strInput & "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "#"    '����ʱ��

    Call WriteLog("DataUnloading(" & strInput & "," & str��ˮ�� & "," & rsTemp!���ı�� & ")")
    strReturn = DataUnloading(strInput, str��ˮ��, rsTemp("���ı��"))
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '��д�����
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    With g��������
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(datCurr) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
            cur����ͳ���ۼ� + .����ͳ���� & "," & _
            curͳ�ﱨ���ۼ� + .ͳ�ﱨ����� & "," & intסԺ�����ۼ� & ")"
        Call ExecuteProcedure(gstrSysName)
        
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
            Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",NULL," & .ʵ������ & "," & g��������.�������ý�� & _
            "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & .�����Ը���� & ",0,''," & .��ҳID & ")"
        Call ExecuteProcedure(gstrSysName)
        
        '���ս������
        gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",NULL)"
        Call ExecuteProcedure(gstrSysName)
    End With
        
    סԺ����_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�¶�(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    סԺ�������_�¶� = False
End Function

Private Function GetҽԺ����(ByRef strҽԺ���� As String, ByVal str�����ı��� As String, Optional ByVal blnҽԺ��� As Boolean) As Boolean
'���ܣ��õ�ҽԺ��ҽ������
    Dim strReturn As String, arrOutput As Variant
    Dim strTemp As String, varList As Variant, lngIndex As Long, strHospital As String
    
    On Error GoTo errHandle
    
    strReturn = GetHospitalInfo()
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '���Ƚ��ִ���ԭ
    strTemp = ""
    For lngIndex = 1 To UBound(arrOutput)
        strTemp = strTemp & "|" & arrOutput(lngIndex)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '֧����һ�����ӵ�|
    If Right(strTemp, 1) = "#" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) '֧������#
    
    varList = Split(strTemp, "$")
    
    For lngIndex = 0 To UBound(varList)
        arrOutput = Split(varList(lngIndex), "|")
        
        If UBound(arrOutput) > 3 Then
            If arrOutput(3) = str�����ı��� Then
                If blnҽԺ��� = True Then
                    strHospital = arrOutput(2) 'ҽԺ���
                Else
                    strHospital = arrOutput(0) 'ҽԺ����
                End If
            End If
        End If
    Next
    
    If strHospital <> "" Then
        strҽԺ���� = strHospital
        GetҽԺ���� = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get���ı���() As String
'���ܣ��õ�ҽԺ��ҽ������
    Dim strReturn As String, arrOutput As Variant
    Dim strTemp As String, varList As Variant, lngIndex As Long, strHospital As String
    Dim strҽԺ���� As String, rsTmp As New ADODB.Recordset
        
    On Error GoTo errHandle
    '��ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�¶�
    Call OpenRecordset(rsTmp, gstrSysName)
    
    If IsNull(rsTmp("ҽԺ����")) = True Then
        MsgBox "����δ����ҽԺ��ţ��޷�ִ��ҽ�����ף�", vbExclamation, gstrSysName
        Exit Function
    End If
    strҽԺ���� = rsTmp!ҽԺ����
    
    strReturn = GetHospitalInfo()
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '���Ƚ��ִ���ԭ
    strTemp = ""
    For lngIndex = 1 To UBound(arrOutput)
        strTemp = strTemp & "|" & arrOutput(lngIndex)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '֧����һ�����ӵ�|
    If Right(strTemp, 1) = "#" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) '֧������#
    
    varList = Split(strTemp, "$")
    
    For lngIndex = 0 To UBound(varList)
        arrOutput = Split(varList(lngIndex), "|")
        
        If UBound(arrOutput) > 3 Then
            If arrOutput(0) = strҽԺ���� Then
                Get���ı��� = arrOutput(3) '���ı���
                Exit For
            End If
        End If
    Next
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function JudgeReturn(ByRef strReturn As String, ByRef varOut As Variant) As Boolean
'���ܣ��жϷ���ֵ�Ƿ�Ϸ���
    Dim varArray As Variant, lngReturn As Long, lngPos As Long
    Dim strSuggest
    
    strReturn = TruncZero(strReturn)
    lngPos = InStr(strReturn, "#")
    If lngPos > 0 Then
        strReturn = Mid(strReturn, 1, lngPos - 1)
    End If
    
    varArray = Split(strReturn, "|")
    
    lngReturn = Val(varArray(0))
    If lngReturn < 0 Then
        'ҵ�����ʧ��
        If UBound(varArray) > 0 Then
            strReturn = "ҽ��ҵ����ʧ�ܡ�" & vbCrLf & varArray(1)
        Else
            strReturn = "ҽ��ҵ����ʧ�ܡ�"
        End If
        
        Select Case lngReturn
            Case -1101
                strSuggest = "�����������ʶ�𲢻�ȡ�µ���ˮ�š�"
            Case -1102, -1210, -1216, -1404, -1405, -1502
                strSuggest = "��Ҫ������˾��顣"
            Case -1201, -1203, -1204, -1205, -1207, -1213, -1215, -1217, -1220
                strSuggest = "��Ҫ���籣��ȷ�ϡ�"
            Case -1208
                strSuggest = "�����������뿨�����ɲ��˵Ĵſ�����ˢ��"
        End Select
        
        If strSuggest <> "" Then
            strReturn = strSuggest & vbCrLf & vbCrLf & "���鴦������" & strSuggest
        End If
        
        Screen.MousePointer = vbDefault
        MsgBox strReturn, vbExclamation, gstrSysName
        Exit Function
    End If
    
    varOut = varArray
    JudgeReturn = True
End Function

Private Function Get��ˮ��(ByVal str��־ As String, ByVal strҽԺ���� As String, ByRef str��ˮ�� As String) As Boolean
    Dim datCurr As Date
    
    datCurr = zlDataBase.Currentdate
    '[��Ϣ��־+ҽԺ����+YYMMDD+6λ��ˮ��]
    str��ˮ�� = str��־ & strҽԺ���� & Format(datCurr, "yyMMddHHmmss")
    Get��ˮ�� = True
End Function

Public Function ҽ����Ŀ_�¶�(rsTemp As ADODB.Recordset) As Boolean
'���ܣ�ҽ������ҩƷĿ¼��ѯ
    Dim str���� As String, str���� As String, str���� As String
    Dim strPath As String, strFile As String, strReturn As String, arrOutput As Variant
    Dim lngFile  As Long, str���ı�� As String
    
    
    str���ı�� = Get���ı���
    If str���ı�� = "" Then Exit Function
    
    '���ýӿڣ������ļ�
    strFile = Space(255)
    GetTempPath 255, strFile
    strPath = TrimStr(strFile)
    strFile = strPath & "MakeTxt.txt"
    
    strReturn = MakeTxt(strFile, strPath & "Temp.txt") '����Ŀ¼��Ȼ��Ҫ,��Ҳ���봫
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    
    On Error GoTo errHandle
    Do Until EOF(lngFile)
        Line Input #lngFile, strReturn
        
        arrOutput = Split(strReturn, vbTab)
        If UBound(arrOutput) >= 11 Then
            str���� = arrOutput(0)
            str���� = ToVarchar(arrOutput(1), 40)
            str���� = ToVarchar(zlCommFun.SpellCode(arrOutput(1)), 10)
        End If
        If str���� <> "" And arrOutput(11) = str���ı�� Then
            'ֻȡ��ǰ���ĵ�ҽ������,�������ĵı�����ܲ�ͬ
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY"), Array("1", str����, str����, str����)
            rsTemp.Update
        End If
    Loop
    Close #lngFile
    Kill strFile
    Kill strPath & "Temp.txt"
    
    ҽ����Ŀ_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Close #lngFile
    
End Function

Public Function ������_�¶�(ByVal str������ As String, strҽ���� As String, str���� As String, str���ı�� As String) As Boolean
'���ܣ����ſ����ݽ��н���
    Dim strReturn  As String, arrOutput As Variant
    
    On Error GoTo errHandle
    
    If str������ = "" Then
        MsgBox "���Ƚ���ˢ��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    strReturn = GetKard(str������)  '����Ϊҽ���š����š�ҽԺ���롢�����ı��
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    strҽ���� = arrOutput(1)
    str���� = arrOutput(2)
    str���ı�� = arrOutput(3)
    ������_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��������_�¶�(ByVal str���� As String, ByVal strҽ���� As String, ByVal str���ı�� As String, _
            ByVal strԭ���� As String, ByVal str������ As String) As Boolean

'���ܣ��޸��û�����
    Dim strInput As String, arrOutput   As Variant, strReturn As String
    Dim strҽԺ���� As String, str��ˮ�� As String
    
    On Error GoTo errHandle
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    If Get��ˮ��("B", strҽԺ����, str��ˮ��) = False Then Exit Function
    
    strInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & strԭ���� & "|" & str������ & "#"
    
    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    MsgBox "�����뱣��ɹ���", vbInformation, gstrSysName
    ��������_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub �˶��ʻ�֧��_�¶�(ByVal lng����ID As Long)
    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
    Dim int��¼��_Client As Integer, cur���_Client As Currency
    Dim lng��ҳID As Long
    Dim strInput As String, strReturn As String, arrOutput
    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���Գ�Ժ���˽��м��
    On Error GoTo ErrHand
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ȡ�ϴ�סԺ����ҳID����Ϊ�ù�����Ҫ���ڳ�Ժ��ʹ�ã���˼ٶ��ò���δ�ٴ���Ժ
    gstrSQL = "Select nvl(סԺ����,1) ��ҳID From ������Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���ϴ�סԺʱ����ҳID")
    lng��ҳID = rsTemp!��ҳID
    
    'ȡ�ʻ�֧����¼����֧�����
    gstrSQL = "Select Sum(A.��Ԥ��) �ʻ�֧��,Count(*) ��¼��  " & _
             " From ����Ԥ����¼ A, " & _
             "      (Select ����ID,����ID  " & _
             "      From ���˷��ü�¼ " & _
             "      Where ����ID=1 And ��ҳID=1) B " & _
             " Where A.����ID=B.����ID And A.���㷽ʽ='�����ʻ�'"
    Call OpenRecordset(rsTemp, "ȡ�ʻ�֧�����¼��")
    int��¼��_Client = Nvl(rsTemp!��¼��, 0)
    cur���_Client = Nvl(rsTemp!�ʻ�֧��, 0)
    
    '��ȡ������Ϣ
    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
            " Where ����ID=" & lng����ID & " And ����=" & TYPE_�¶�
    Call OpenRecordset(rsTemp, "��ȡ������Ϣ")
    str������� = rsTemp!�������
    str���ı�� = rsTemp!���ı��
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
    
    '���ú˶Խӿ�
    If Get��ˮ��("H", strҽԺ����, str��ˮ��) = False Then Exit Sub
    strInput = ToVarchar(str���ı��, 4)
    strInput = strInput & "|" & ToVarchar(strҽԺ����, 8)
    strInput = strInput & "|" & str�������
    strInput = strInput & "|" & str������� & "|%#"
    
'    MsgBox "�˶��ʻ�֧����DataUnloading" & strInput
    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼��;2-֧���
    int��¼��_OUT = arrOutput(1)
    cur���_OUT = arrOutput(2)
    
    If Format(cur���_OUT, "#####0.00;-#####0.00;0;") <> Format(cur���_Client, "#####0.00;-#####0.00;0;") Then
        MsgBox "�����ʻ�֧������ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
               "����ʵ���ʻ�֧����" & cur���_Client & String(4, " ") & "ҽ������ͳ�Ƴ����ʻ�֧����" & cur���_OUT & vbCrLf & _
               "�����ʻ�֧��������" & int��¼��_Client & String(4, " ") & "ҽ������ͳ�Ƴ���֧��������" & int��¼��_OUT
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶����Ժ_�¶�(ByVal lng����ID As Long)
    '���Գ�Ժ���˽��м��
    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
    Dim int��¼��_Client As Integer, cur���_Client As Currency
    Dim lng��ҳID As Long
    Dim strInput As String, strReturn As String, arrOutput
    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���Գ�Ժ���˽��м��
    On Error GoTo ErrHand
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    int��¼��_Client = 1
    
    '��ȡ������Ϣ
    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
            " Where ����ID=" & lng����ID & " And ����=" & TYPE_�¶�
    Call OpenRecordset(rsTemp, "��ȡ������Ϣ")
    str������� = rsTemp!�������
    str���ı�� = rsTemp!���ı��
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
    
    '���ú˶Խӿ�
    If Get��ˮ��("I", strҽԺ����, str��ˮ��) = False Then Exit Sub
    strInput = ToVarchar(str���ı��, 4)
    strInput = strInput & "|" & ToVarchar(strҽԺ����, 8)
    strInput = strInput & "|" & str�������
    strInput = strInput & "|" & str������� & "|#"
    
'    MsgBox "�˶����Ժ��¼��DataUnloading" & strInput
    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼����
    int��¼��_OUT = arrOutput(1)
    
    If int��¼��_OUT <> int��¼��_Client Then
        MsgBox "�������Ժ��¼��ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
               "�������Ժ��¼����" & int��¼��_Client & String(4, " ") & "ҽ������ͳ�Ƴ������Ժ��¼����" & int��¼��_OUT
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶Է��ý���_�¶�(ByVal lng����ID As Long)
    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
    Dim cur����_OUT As Currency, curȫ�Է�_OUT As Currency
    Dim cur�����Ը�_OUT As Currency, curʵ������_OUT As Currency
    Dim curͳ��֧��_OUT As Currency, curͳ���Ը�_OUT As Currency
    Dim cur�����Ը�_OUT As Currency, cur�ʻ�֧��_OUT As Currency
    Dim int��¼��_Client As Integer, cur���_Client As Currency
    Dim cur����_Client As Currency, curȫ�Է�_Client As Currency
    Dim cur�����Ը�_Client As Currency, curʵ������_Client As Currency
    Dim curͳ��֧��_Client As Currency, curͳ���Ը�_Client As Currency
    Dim cur�����Ը�_Client As Currency, cur�ʻ�֧��_Client As Currency
    Dim lng��ҳID As Long
    Dim strInput As String, strReturn As String, arrOutput
    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���Գ�Ժ���˽��м��
    On Error GoTo ErrHand
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ȡ�ϴ�סԺ����ҳID����Ϊ�ù�����Ҫ���ڳ�Ժ��ʹ�ã���˼ٶ��ò���δ�ٴ���Ժ
    gstrSQL = "Select nvl(סԺ����,1) ��ҳID From ������Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���ϴ�סԺʱ����ҳID")
    lng��ҳID = rsTemp!��ҳID
    
    'ȡ�ʻ�֧����¼����֧�����
    gstrSQL = "SELECT SUM(�������ý��) ��������,SUM(����ͳ����) ����ͳ��,SUM(ͳ�ﱨ�����) ͳ�ﱨ��, " & _
             " SUM(�����Ը����) �����Ը�,SUM(����) ����,SUM(ʵ������) ʵ������," & _
             " SUM(�����Ը����) �����Ը�,SUM(�����ʻ�֧��) �����ʻ�֧��,Count(*) ��¼�� " & _
             " FROM  " & _
             "      (SELECT ����ID,����ID FROM ���˷��ü�¼ " & _
             "      WHERE ����ID=" & lng����ID & " AND ��ҳID= " & lng��ҳID & _
             "      ) A,���ս����¼ B " & _
             " WHERE A.����ID=B.����ID AND B.��¼ID=A.����ID AND B.����=" & TYPE_�¶� & " AND B.����=2 " & _
             " GROUP BY A.����ID"
    Call OpenRecordset(rsTemp, "ȡ�ʻ�֧�����¼��")
    int��¼��_Client = Nvl(rsTemp!��¼��, 0)
    cur���_Client = Nvl(rsTemp!��������, 0)
    curͳ��֧��_Client = Nvl(rsTemp!ͳ�ﱨ��, 0)
    cur�ʻ�֧��_Client = Nvl(rsTemp!�����ʻ�֧��, 0)
    
    '��ȡ������Ϣ
    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
            " Where ����ID=" & lng����ID & " And ����=" & TYPE_�¶�
    Call OpenRecordset(rsTemp, "��ȡ������Ϣ")
    str������� = rsTemp!�������
    str���ı�� = rsTemp!���ı��
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
    
    '���ú˶Խӿ�
    If Get��ˮ��("J", strҽԺ����, str��ˮ��) = False Then Exit Sub
    strInput = ToVarchar(str���ı��, 4)
    strInput = strInput & "|" & ToVarchar(strҽԺ����, 8)
    strInput = strInput & "|" & str�������
    strInput = strInput & "|" & str������� & "|%|%|%#"
    
'    MsgBox "�˶Է��ý��㣺DataUnloading" & strInput
    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼��;2-֧���
    int��¼��_OUT = arrOutput(1)
    cur����_OUT = arrOutput(2)
    cur���_OUT = arrOutput(3)
    curȫ�Է�_OUT = arrOutput(4)
    cur�����Ը�_OUT = arrOutput(5)
    'cur����ͳ��_OUT = arrOutput(6)
    curʵ������_OUT = arrOutput(7)
    curͳ��֧��_OUT = arrOutput(8)
    curͳ���Ը�_OUT = arrOutput(9)
    cur�����Ը�_OUT = arrOutput(10)
    cur�ʻ�֧��_OUT = arrOutput(11)
    
    'ֻҪͳ��֧�����ʻ�֧���������ܶ�һ�¼���
    If Not (Format(cur���_OUT, "#####0.00;-#####0.00;0;") = Format(cur���_Client, "#####0.00;-#####0.00;0;") _
    And Format(curͳ��֧��_OUT, "#####0.00;-#####0.00;0;") = Format(curͳ��֧��_Client, "#####0.00;-#####0.00;0;") _
    And Format(cur�ʻ�֧��_OUT, "#####0.00;-#####0.00;0;") = Format(cur�ʻ�֧��_Client, "#####0.00;-#####0.00;0;")) Then
        MsgBox "���ؽ���������ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
               "��ҽ���������ܶ" & cur���_OUT & String(4, " ") & "ͳ��֧����" & curͳ��֧��_OUT & String(4, " ") & "�ʻ�֧����" & cur�ʻ�֧��_OUT & vbCrLf & _
               "�����أ������ܶ" & cur���_Client & String(4, " ") & "ͳ��֧����" & curͳ��֧��_Client & String(4, " ") & "�ʻ�֧����" & cur�ʻ�֧��_Client
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶Է�����ϸ_�¶�(ByVal lng����ID As Long)
'    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
'    Dim int��¼��_Client As Integer, cur���_Client As Currency
'    Dim lng��ҳID As Long
'    Dim strInput As String, strReturn As String, arrOutput
'    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
'    Dim rsTemp As New ADODB.Recordset
'    '���Գ�Ժ���˽��м��
'    On Error GoTo ErrHand
'
'    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
'        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    'ȡ�ϴ�סԺ����ҳID����Ϊ�ù�����Ҫ���ڳ�Ժ��ʹ�ã���˼ٶ��ò���δ�ٴ���Ժ
'    gstrSQL = "Select nvl(סԺ����,1) ��ҳID From ������Ϣ Where ����ID=" & lng����ID
'    Call OpenRecordset(rsTemp, "���ϴ�סԺʱ����ҳID")
'    lng��ҳID = rsTemp!��ҳID
'
'    'ȡ�ʻ�֧����¼����֧�����
'    gstrSQL = "Select Sum(A.��Ԥ��) �ʻ�֧��,Count(*) ��¼��  " & _
'             " From ����Ԥ����¼ A, " & _
'             "      (Select ����ID,����ID  " & _
'             "      From ���˷��ü�¼ " & _
'             "      Where ����ID=1 And ��ҳID=1) B " & _
'             " Where A.����ID=B.����ID And A.���㷽ʽ='�����ʻ�'"
'    Call OpenRecordset(rsTemp, "ȡ�ʻ�֧�����¼��")
'    int��¼��_Client = NVL(rsTemp!��¼��, 0)
'    cur���_Client = NVL(rsTemp!�ʻ�֧��, 0)
'
'    '��ȡ������Ϣ
'    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
'            " Where ����ID=" & lng����ID & " And ����=" & TYPE_�¶�
'    Call OpenRecordset(rsTemp, "��ȡ������Ϣ")
'    str������� = rsTemp!�������
'    str���ı�� = rsTemp!���ı��
'    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
'    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
'
'    '���ú˶Խӿ�
'    If Get��ˮ��("H", strҽԺ����, str��ˮ��) = False Then Exit Sub
'    strInput = ToVarchar(str���ı��, 4)
'    strInput = strInput & "|" & ToVarchar(strҽԺ����, 8)
'    strInput = strInput & "|" & str�������
'    strInput = strInput & "|" & str������� & "|%#"
'
'    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
'    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
'
'    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼��;2-֧���
'    int��¼��_OUT = arrOutput(1)
'    cur���_OUT = arrOutput(2)
'
'    If Format(cur���_OUT, "#####0.00;-#####0.00;0;") <> Format(cur���_Client, "#####0.00;-#####0.00;0;") Then
'        MsgBox "�����ʻ�֧������ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
'               "����ʵ���ʻ�֧����" & cur���_Client & String(4, " ") & "ҽ������ͳ�Ƴ����ʻ�֧����" & cur���_OUT & vbCrLf & _
'               "�����ʻ�֧��������" & int��¼��_Client & String(4, " ") & "ҽ������ͳ�Ƴ���֧��������" & int��¼��_OUT
'    Else
'        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
'    End If
'    Exit Sub
'ErrHand:
'    If ErrCenter = 1 Then Resume
End Sub

Private Sub WriteLog(ByVal strInfo As String)
    Dim strFileName As String
    Dim objSystem As FileSystemObject
    Dim objStream As TextStream
    
    If Val(GetSetting("ZLSOFT", "ҽ��", "����", 0)) = 0 Then Exit Sub
    strFileName = "C:\" & Format(Date, "yyyyMMdd") & ".txt"
    Set objSystem = New FileSystemObject
    If Not objSystem.FileExists(strFileName) Then Call objSystem.CreateTextFile(strFileName, False)
    Set objStream = objSystem.OpenTextFile(strFileName, ForAppending, False, TristateMixed)
    objStream.WriteLine (strInfo)
    objStream.Close
End Sub
