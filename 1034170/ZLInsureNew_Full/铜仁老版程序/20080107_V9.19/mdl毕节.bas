Attribute VB_Name = "mdl�Ͻ�"
Option Explicit
'IC����ز�����Ҫ�ṩ��������
'�ʻ�����[��������]������ҽ�ơ��󲡡�����Ա��סԺͳ�סԺ��ǰ�����ų⣬�Խ�����Ӱ�죬ֱ�Ӱ����߱���м���
'�������: ҩƷ�����Ƹ�����Ϊһ�ʴ�������
'ҽ�Ʒ���֧����ϸ��: ������ܱ�
'����סԺ����: ����ֵ������ͳ���ۼƽ��
'֧�ֽ��㷽ʽ����: �����ʻ���ͳ�����
'ÿ�����г���ǰ������Ƿ����δ�ϴ���δ���أ���������������
'��������ʱ������Ϊ����������=ҽ�Ʊ���
'�����㷨
'   ���ȫ���ʻ�֧����֧��
'   סԺ���Ȱ���Ŀ��������ͳ����ٰ����ּ���ʵ�ʽ���ͳ�����󰴷ֵ�����ó�ʵ��ͳ�����֧�����
'���ڸ�ʽ: yyyy.MM.dd��ʱ���ʽ:yyyy.MM.dd HH:mm:ss
'����ʻ�ͣ�ã���ֻ�ܵ�����ͨ���˴���
'δ������Ŀ���������ҩƷ�����Ը�������Ϊȫ�Ը�������ȫ����ͳ������ֵ�����
'��֧����;����,����ʱ�����Ժ
'ʵʱ�����ѻ�ʹ�������Ķ˿��ƣ���ʵʱ�ϴ�������ϸ���������������Ķ˵�δ������ϸ�������ν�����ϸ�ϴ������ģ��ٽ�������Ϣ�ϴ�������
'��Ҫ��ɽ��㲿�������סԺ���㣩���ϴ����أ����úͽ������ݵ��ϴ����������ݵ����أ���IC����д���������ĵĲ���
'�������ʹ���˸����ʻ��ģ���Ҫ�������ĵĸ����ʻ��������صĸ����ʻ��������¿�
'Ҫ��ÿ���������˳�������Ҫ����ʹ�ã�������������һ��
'���������д���м�⣬�������ϸ���Ƿ��ϴ���Ϊ1������ϴ������ģ����м���е��Ƿ��ϴ���Ϊ1
Private Enum ic
    shbzh = 0
    xm
    dwdm
    xb
    csrq
    cjqzrq
    jyqkdm
    yxkh
    grjbdm
    ye
    zhjzrq
    yydm
    pass
End Enum

Private Type IC_Struct
    ��ᱣ�Ϻ� As String
    ���� As String
    ��λ���� As String
    �Ա� As String
    �������� As String
    �μӹ������� As String
    ��ҵ������� As String
    ��Ч���� As Integer
    ���˼������ As String
    �����ʻ���� As Double
    ���������� As String
    ������ҽԺ���� As String
    ����IC������    As String
End Type
Public IC_Data_�Ͻ� As IC_Struct

Private Type gCominfo
    strHospitalCode As String       'ҽԺ����
    strHospitalName As String       'ҽԺ����
    strConnectPass As String        '��������
    blnOnLine As Boolean            'ʵʱ���������ѻ�
    blnICPassVerify As Boolean      '�Ƿ�ʹ��IC������
    blnDiseaseCash As Boolean       '�Ƿ����ò����Ը�
    blnPhysicCash As Boolean        '�Ƿ�����ҩƷ�����Ը�
    blnYearBase As Boolean          '�Ƿ������
    '����סԺʹ��
    str������ˮ�� As String         '������ˮ��
    dbl�����ܶ� As Double
    dbl���ͳ�� As Double           '���ͳ���ۼƣ�ָ����ֵ�����ʱ��ͳ�����ۼƣ�
    dblͳ���� As Double           'ͬ�ϣ�ֻ�Ǳ��ν���ֵ�����ʱ��ͳ����
    dbl��ȱ��� As Double           '����ȱ�������ۼƣ���Ҫ�ٴμ���
    dblͳ�ﱨ�� As Double           '����ͳ�ﱨ�������������𸶣����ȥdbl��ȱ����õ�����ʵ��ͳ��֧�����
    str��ᱣ�Ϻ� As String         'ҽ����
    str��Ч���� As String           '����
End Type
Public gCominfo_�Ͻ� As gCominfo

Public gcnGYBJYB As New ADODB.Connection
Public gobjCenter As Object

Private mblnInit As Boolean
Private mstrFirstStart As String        '��¼��һ�ε�¼���ڣ������ͬ���ֹʹ��

Private Const gstrҩƷ���� As String = "000"
Private Const gstrҩƷ���� As String = "��ҩ��"
Private Const gstr���ƴ��� As String = "01"
Private Const gstr���ƴ��� As String = "����"

Public Function GetAge(ByVal strServer As String, ByVal strTest As String) As Long
    Dim strServerYear As String, strTestYear As String
    Dim strServerMonth As String, strTestMonth As String
    Dim strServerDay As String, strTestDay As String
    Dim lngAge As Long
    Dim intDef As Integer
    
    '�����������,��:δ��31��,��30����
    If Not IsDate(strServer) Then
        MsgBox "����ĵ�һ���������������ͣ�[GetAge]", vbInformation, gstrSysName
        Exit Function
    End If
    If Not IsDate(strTest) Then
        MsgBox "����ĵڶ����������������ͣ�[GetAge]", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ֽ��ꡢ�¡���
    strServerYear = Mid(strServer, 1, 4)
    strServerMonth = Mid(strServer, 6, 2)
    strServerDay = Mid(strServer, 9, 2)
    strTestYear = Mid(strTest, 1, 4)
    strTestMonth = Mid(strTest, 6, 2)
    strTestDay = Mid(strTest, 9, 2)
    
    '�Ȱ����㣬�ó���ŵ�����
    lngAge = Val(strServerYear) - Val(strTestYear)
    '����������·ݴ��ڳ����·ݣ�������ֱ�ӷ��أ����С�ڣ��������1�������ͬ��������ж�
    intDef = Val(strServerMonth) - Val(strTestMonth)
    If intDef > 0 Then
        GetAge = lngAge
        Exit Function
    ElseIf intDef < 0 Then
        GetAge = (lngAge - 1)
        Exit Function
    Else
        intDef = Val(strServerDay) - Val(strTestDay)
        If intDef >= 0 Then
            GetAge = lngAge
            Exit Function
        Else
            GetAge = (lngAge - 1)
            Exit Function
        End If
    End If
End Function

Public Function ��ݱ�ʶ_�Ͻ�(Optional bytType As Byte, Optional lng����ID As Long) As String
    ��ݱ�ʶ_�Ͻ� = frmIdentify�Ͻ�.GetIdentify(bytType, lng����ID)
End Function

Public Function ҽ����ʼ��_�Ͻ�(Optional ByVal blnTest As Boolean = False) As Boolean
    '���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
    '���أ���ʼ���ɹ�������true�����򣬷���false
    Dim bln��ֹ��¼ As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strServer As String, strUser As String, strPass As String
    On Error GoTo ErrHand
    
    If mblnInit Then
        ҽ����ʼ��_�Ͻ� = True
        Exit Function
    End If
    
    On Error Resume Next
    If gobjCenter Is Nothing Then
        Err = 0
        Set gobjCenter = CreateObject("Interface.clsInterface")
        If Err <> 0 Then
            MsgBox "�޷������������Ĳ���������ҽ�����Ļ򿪷�����ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    On Error GoTo ErrHand
    'ȡ��������е�ҽԺ����
    gCominfo_�Ͻ�.strHospitalCode = ""
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�Ͻ�
    Call OpenRecordset(rsTemp, "ȡҽԺ����")
    If rsTemp.RecordCount <> 0 Then
        gCominfo_�Ͻ�.strHospitalCode = Nvl(rsTemp!ҽԺ����)
    End If
    If gCominfo_�Ͻ�.strHospitalCode = "" Then
        MsgBox "��δ��ʼ�������ڱ��������������ñ�ҽ�ƻ����ĵ�λ���룡", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������ҽ��������������
    gCominfo_�Ͻ�.strConnectPass = ""
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=" & TYPE_�Ͻ�
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "��������"
                gCominfo_�Ͻ�.strConnectPass = Nvl(rsTemp!����ֵ)
        End Select
        rsTemp.MoveNext
    Loop

    If Not OraDataOpen(gcnGYBJYB, strServer, strUser, strPass) Then Exit Function
    
    'ȡ���ӷ�ʽ
    If Not ��ȡ���ӷ�ʽ() Then Exit Function
    
    '����Ƿ����δ�ϴ��ķ�����ϸ��������ݣ��������������ʹ�ã���ʾ�û�ʹ���ϴ�����
    If Not ����Ƿ��ϴ���ϸ Then Exit Function
    
    '�������Ƿ���й����أ����û�У�Ҳ��ֹʹ�ã�ͬʱ��ȡ����λ���ƣ�
    If Not ����Ƿ����� Then Exit Function
    
    '����Ƿ�����ͨ����
    If Not ��ȡ�������� Then Exit Function
    ҽ����ʼ��_�Ͻ� = �����������
    Call �ر���������
    
    If mstrFirstStart = "" Then mstrFirstStart = Format(zlDataBase.Currentdate, "yyyy-MM-dd")
    
    mblnInit = True
    ҽ����ʼ��_�Ͻ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ������_�Ͻ�() As Boolean
    ҽ������_�Ͻ� = frmSet�Ͻ�.��������()
End Function

Public Function ҽ����ֹ_�Ͻ�() As Boolean
    On Error Resume Next
    
    If gCominfo_�Ͻ�.blnOnLine Then
        If Not gobjCenter Is Nothing Then
            Call gobjCenter.CloseConnector
            Set gobjCenter = Nothing
        End If
    End If
End Function

Public Function �����������_�Ͻ�(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim dbl��� As Double
    Dim cur�ʻ�֧�� As Double
    With rs��ϸ
        Do While Not .EOF
            dbl��� = dbl��� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    
    '��������ʻ������ڱ��ν���������ʻ�֧������ڽ����������ڸ����ʻ����
    cur�ʻ�֧�� = IIf(IC_Data_�Ͻ�.�����ʻ���� >= dbl���, dbl���, IC_Data_�Ͻ�.�����ʻ����)
    
    str���㷽ʽ = "�����ʻ�;" & Format(cur�ʻ�֧��, "#####0.00") & ";1"
    �����������_�Ͻ� = True
End Function

Public Function �������_�Ͻ�(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    Dim strCard As String
    Dim blnTrans As Boolean
    Dim lng����ID As Long
    Dim blnҩƷ As Boolean
    Dim str����ǼǺ� As String
    Dim cur�ʻ�֧�� As Currency
    Dim cur�ʻ�֧���� As Currency           '��¼ҩƷ�������У��ʻ�ʵ��֧����
    Dim dbl�����ܶ� As Double
    Dim rsDetail As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo ErrHand
    
    cur�ʻ�֧�� = cur�����ʻ�
    '��ȡ���ν��������ϸ
    gstrSQL = " Select  A.����ID,A.�շ����,A.�շ�ϸĿID,round(A.ʵ�ս��,2) ʵ�ս��,B.��Ŀ����,Nvl(B.��Ŀ����,C.����) AS ��Ŀ����,B.��ע" & _
              " From ���˷��ü�¼ A,����֧����Ŀ B,�շ�ϸĿ C" & _
              " Where A.����ID=" & lng����ID & " And A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=" & TYPE_�Ͻ� & _
              " And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0" & _
              " And A.�շ�ϸĿID=C.ID"
    Call OpenRecordset(rsDetail, "��ȡ���ν��������ϸ")
    lng����ID = rsDetail!����ID
    
    '�жϿ��ǲ��ǵ�ǰ���˵�
    gstrSQL = "Select ����,ҽ���� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_�Ͻ�
    Call OpenRecordset(rsCheck, "�жϿ��ǲ��ǵ�ǰ���˵�")
    
    '����
    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
    Call ����ת��_�Ͻ�(strCard, True)
    If Not (IC_Data_�Ͻ�.��Ч���� = Nvl(rsCheck!����, 0) And IC_Data_�Ͻ�.��ᱣ�Ϻ� = rsCheck!ҽ����) Then
        MsgBox "��ǰIC�����Ǹò��˵Ŀ���ÿ���ʧЧ������ҽ��������ϵ��", vbInformation, gstrSysName
        Call IC_End(True)
        Exit Function
    End If
        
    str����ǼǺ� = Get��ˮ��_�Ͻ�
    blnTrans = True
    If Not ����_��ʼ Then
        Call IC_End(True)
        Exit Function
    End If
    
    cur�ʻ�֧���� = 0
    With rsDetail
        Do While Not .EOF
            '��д�м������(��дҩƷ������ϸ�����Ʒ�����ϸ��ҽ�Ʒ���֧����ϸ��)
            blnҩƷ = (InStr(1, "5,6,7", !�շ����) <> 0)
            If blnҩƷ Then
                gstrSQL = "" & _
                    " INSERT INTO ҩƷ������ϸ��" & _
                    " (ID,��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
                    " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                    " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                    " VALUES" & _
                    " (ҩƷ������ϸ��_ID.Nextval,'" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
                    "'" & Nvl(!��Ŀ����, gstrҩƷ����) & "','" & !��Ŀ���� & "','" & Nvl(!��ע, gstrҩƷ����) & "'," & _
                    "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','����'," & !ʵ�ս�� & "," & _
                    "0," & cur�ʻ�֧���� & "," & !ʵ�ս�� - cur�ʻ�֧���� & "," & _
                    "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','��','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
                gcnGYBJYB.Execute gstrSQL
            Else
                gstrSQL = "" & _
                    " INSERT INTO ���Ʒ�����ϸ��" & _
                    " (ID,��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
                    " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                    " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                    " VALUES" & _
                    " (���Ʒ�����ϸ��_ID.Nextval,'" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
                    "'" & Nvl(!��Ŀ����, gstr���ƴ���) & "','" & !��Ŀ���� & "','" & Nvl(!��ע, gstr���ƴ���) & "'," & _
                    "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','����'," & !ʵ�ս�� & "," & _
                    "0," & cur�ʻ�֧���� & "," & !ʵ�ս�� - cur�ʻ�֧���� & "," & _
                    "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','��','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
                gcnGYBJYB.Execute gstrSQL
            End If
            
            '��д���Ŀ����ݣ�ͬ��
            If gCominfo_�Ͻ�.blnOnLine Then
                If blnҩƷ Then
                    gstrSQL = "" & _
                        " INSERT INTO ҩƷ������ϸ��" & _
                        " (��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
                        " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                        " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա)" & _
                        " VALUES" & _
                        " ('" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
                        "'" & Nvl(!��Ŀ����, gstrҩƷ����) & "','" & !��Ŀ���� & "','" & Nvl(!��ע, gstrҩƷ����) & "'," & _
                        "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','����'," & !ʵ�ս�� & "," & _
                        "0," & cur�ʻ�֧���� & "," & !ʵ�ս�� - cur�ʻ�֧���� & "," & _
                        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "')"
                    If Not ExecuteSQL(gstrSQL) Then
                        Call IC_End(True)
                        Exit Function
                    End If
                Else
                    gstrSQL = "" & _
                        " INSERT INTO ���Ʒ�����ϸ��" & _
                        " (��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
                        " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                        " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա)" & _
                        " VALUES" & _
                        " ('" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
                        "'" & Nvl(!��Ŀ����, gstr���ƴ���) & "','" & !��Ŀ���� & "','" & Nvl(!��ע, gstr���ƴ���) & "'," & _
                        "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','����'," & !ʵ�ս�� & "," & _
                        "0," & cur�ʻ�֧���� & "," & !ʵ�ս�� - cur�ʻ�֧���� & "," & _
                        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "')"
                    If Not ExecuteSQL(gstrSQL) Then
                        Call IC_End(True)
                        Exit Function
                    End If
                End If
            End If
            
            cur�ʻ�֧�� = cur�ʻ�֧�� - cur�ʻ�֧����
            dbl�����ܶ� = dbl�����ܶ� + !ʵ�ս��
            .MoveNext
        Loop
    End With
    
    'Ϊ���з�����ϸ�����ϴ����
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '��д���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ͻ� & "," & lng����ID & "," & _
        Format(zlDataBase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        dbl�����ܶ� & "," & dbl�����ܶ� - cur�����ʻ� & ",0,0,0,0," & _
        0 & "," & cur�����ʻ� & ",'" & str����ǼǺ� & "')"
    Call ExecuteProcedure("���汣�ս����¼")
    
    gstrSQL = "" & _
        " INSERT INTO ҽ�Ʒ���֧����ϸ��" & _
        " (ID,��ᱣ�Ϻ�,����,����סԺ��,����ʱ��,�������,�ܷ���,ͳ�����֧��," & _
        " �����ʻ�֧��,�����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ��ϴ�)" & _
        " VALUES" & _
        " (ҽ�Ʒ���֧����ϸ��_ID.Nextval,'" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
        "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','����'," & dbl�����ܶ� & "," & _
        "0," & cur�����ʻ� & "," & dbl�����ܶ� - cur�����ʻ� & "," & _
        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = "" & _
            " INSERT INTO ҽ�Ʒ���֧����ϸ��" & _
            " (��ᱣ�Ϻ�,����,����סԺ��,����ʱ��,�������,�ܷ���,ͳ�����֧��," & _
            " �����ʻ�֧��,�����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա)" & _
            " VALUES" & _
            " ('" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
            "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','����'," & dbl�����ܶ� & "," & _
            "0," & cur�����ʻ� & "," & dbl�����ܶ� - cur�����ʻ� & "," & _
            "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "')"
        If Not ExecuteSQL(gstrSQL) Then
            Call IC_End(True)
            Exit Function
        End If
    End If
    
    '���ʻ������Ŀ�ĸ����ʻ�����Ҫ���£��м��ĸ����ʻ�����Ҫ���£�
    'סԺ����Ҫ���µ�ǰסԺ�����ֶΣ�ͳ��֧����סԺ������
    IC_Data_�Ͻ�.�����ʻ���� = IC_Data_�Ͻ�.�����ʻ���� - cur�����ʻ�
    IC_Data_�Ͻ�.���������� = Format(zlDataBase.Currentdate, "yyyy.MM.dd")
    IC_Data_�Ͻ�.������ҽԺ���� = gCominfo_�Ͻ�.strHospitalCode
    Call ����ת��_�Ͻ�(strCard, False)
    
    gstrSQL = " Update �����ʻ����� " & _
              " Set ���=Nvl(���,0)-" & Val(cur�����ʻ�) & _
              " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then 'SqlServer�ǿպ�����IsNULL();��Oracle��Nvl()
        gstrSQL = " Update �����ʻ����� " & _
                  " Set ����֧��=IsNull(����֧��,0)+" & Val(cur�����ʻ�) & "," & _
                  "     �ۼ�֧��=IsNull(�ۼ�֧��,0)+" & Val(cur�����ʻ�) & "," & _
                  "     ���=IsNull(���,0)-" & Val(cur�����ʻ�) & _
                  " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
        If Not ExecuteSQL(gstrSQL) Then
            Call IC_End(True)
            Exit Function
        End If
    End If
    
    If Not gobjCenter.IC_WriteCard(strCard) Then
        Call ����_�ع�
        Call IC_End(True)
        Exit Function
    End If
    
    If ����_�ύ Then
        �������_�Ͻ� = True
    Else
        Call ����_�ع�
    End If
    
    Call IC_End
    
    '���ʹ�ø����ʻ�֧������ʾ������ʾ��
    If cur�����ʻ� <> 0 Then Call Frm������ʾ��.ShowME(IC_Data_�Ͻ�.����, IC_Data_�Ͻ�.�����ʻ���� + cur�����ʻ�, _
        IC_Data_�Ͻ�.�����ʻ����, dbl�����ܶ�, cur�����ʻ�, dbl�����ܶ� - cur�����ʻ�)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_�ع�
    Call IC_End(True)
End Function

Public Function ����������_�Ͻ�(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    'ֻ����������������ϣ���ֻ�������һ�ʣ�������ϵ�������ҽ�ƻ���������ΪH000��ɾ���м�������ĵķ�����ϸ��֧����ϸ
    Dim lng����ID As Long
    Dim strCard As String
    Dim str����ǼǺ� As String, str�˵���ˮ�� As String
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '�жϿ��ǲ��ǵ�ǰ���˵�
    gstrSQL = "Select ����,ҽ���� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_�Ͻ�
    Call OpenRecordset(rsCheck, "�жϿ��ǲ��ǵ�ǰ���˵�")
    
    '����
    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
    Call ����ת��_�Ͻ�(strCard, True)
    If Not (IC_Data_�Ͻ�.��Ч���� = Nvl(rsCheck!����, 0) And IC_Data_�Ͻ�.��ᱣ�Ϻ� = rsCheck!ҽ����) Then
        Call IC_End(True)
        MsgBox "��ǰIC�����Ǹò��˵Ŀ���ÿ���ʧЧ������ҽ��������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���������ҽ�ƻ������붼����ͬ����ֱ���˳�
    If IC_Data_�Ͻ�.������ҽԺ���� <> gCominfo_�Ͻ�.strHospitalCode Then
        Call IC_End(True)
        MsgBox "����������ҽ�ƻ�������������Ϊ�������˵���", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ���ν���ID
    gstrSQL = "select distinct A.����ID,A.NO from ���˷��ü�¼ A,���˷��ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���²����Ľ���ID")
    lng����ID = rsTemp!����ID
    
    '��ȡԭʼ�ı��ս����¼
    gstrSQL = "Select * From ���ս����¼ Where ����=1 AND ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡԭʼ�ı��ս����¼")
    str����ǼǺ� = Nvl(rsTemp!֧��˳���)
    
    str�˵���ˮ�� = InputBox("������ԭʼ���ݵľ���ǼǺţ�", "����ǼǺ�")
    '�жϿ��ڼ�¼��������ҽԺ���������Ƿ������ĵ�һ�£����ڴ����ѻ���ֻ���м��ȡ��
    gstrSQL = " Select Max(����סԺ��) ����ǼǺ� From ҽ�Ʒ���֧����ϸ��" & _
              " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "' And ҽ�ƻ�������='" & gCominfo_�Ͻ�.strHospitalCode & "'"
    With rsCheck
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnGYBJYB
        If IsNull(!����ǼǺ�) Then
            Call IC_End(True)
            MsgBox "û���ҵ����һ�εľ���ǼǺţ��޷��˵���", vbInformation, gstrSysName
            Exit Function
        End If
        If Nvl(!����ǼǺ�) <> str����ǼǺ� Then
            Call IC_End(True)
            MsgBox "ֻ���˸ò����ڱ�Ժ��������һ�����ﵥ�ݣ�", vbInformation, gstrSysName
            Exit Function
        End If
        If str����ǼǺ� <> str�˵���ˮ�� Then
            Call IC_End(True)
            MsgBox "����ľ���ǼǺ���ԭ���ݵľ���ǼǺŲ������޷��˵���", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    'Ϊ���з�����ϸ�����ϴ����
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '���汣�ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ͻ� & "," & lng����ID & "," & _
        Format(zlDataBase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & ",0,0,0,0,0," & -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",NULL)"
    Call ExecuteProcedure("���汣�ս����¼")
    
    blnTrans = True
    If Not ����_��ʼ Then
        Call IC_End(True)
        Exit Function
    End If
    
    'ɾ���м�⡢���Ŀ��еķ�����ϸ��֧����ϸ��¼
    gstrSQL = " Delete ҩƷ������ϸ�� " & _
              " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "' And סԺ��='" & str����ǼǺ� & "' And ҽ�ƻ�������='" & gCominfo_�Ͻ�.strHospitalCode & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    gstrSQL = " Delete ���Ʒ�����ϸ�� " & _
              " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "' And סԺ��='" & str����ǼǺ� & "' And ҽ�ƻ�������='" & gCominfo_�Ͻ�.strHospitalCode & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    gstrSQL = " Delete ҽ�Ʒ���֧����ϸ�� " & _
              " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "' And ����סԺ��='" & str����ǼǺ� & "' And ҽ�ƻ�������='" & gCominfo_�Ͻ�.strHospitalCode & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    'д�����޸�������ʻ�֧���ۼ�
    '���ʻ������Ŀ�ĸ����ʻ�����Ҫ���£��м��ĸ����ʻ�����Ҫ���£�
    '���ܱ��ν����Ƿ�ʹ�ø����ʻ�����Ҫд����Ŀ���Ǹ���������ҽ�ƻ�������
    IC_Data_�Ͻ�.�����ʻ���� = IC_Data_�Ͻ�.�����ʻ���� + cur�����ʻ�
    IC_Data_�Ͻ�.������ҽԺ���� = "H000"
    Call ����ת��_�Ͻ�(strCard, False)
    
    gstrSQL = " Update �����ʻ����� " & _
              " Set ���=Nvl(���,0)+" & Val(cur�����ʻ�) & _
              " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = " Update �����ʻ����� " & _
                  " Set ����֧��=IsNull(����֧��,0)-" & Val(cur�����ʻ�) & "," & _
                  "     �ۼ�֧��=IsNull(�ۼ�֧��,0)-" & Val(cur�����ʻ�) & "," & _
                  "     ���=IsNull(���,0)+" & Val(cur�����ʻ�) & _
                  " Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
        If Not ExecuteSQL(gstrSQL) Then
            Call IC_End(True)
            Exit Function
        End If
    End If
    
    If Not gobjCenter.IC_WriteCard(strCard) Then
        Call IC_End(True)
        Call ����_�ع�
        Exit Function
    End If
    
    If ����_�ύ Then
        ����������_�Ͻ� = True
    Else
        Call ����_�ع�
    End If
    
    Call IC_End
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_�ع�
    Call IC_End(True)
End Function

Public Function ��Ժ�Ǽ�_�Ͻ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    '����Ժ�Ǽ�ǰ����Ҫ���������֤,���,���������Ϣ��ֱ�Ӵӿ������л�ȡ
    Dim blnTrans As Boolean
    Dim str����ǼǺ� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    str����ǼǺ� = Get��ˮ��_�Ͻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ͻ� & ")"
    Call ExecuteProcedure("������Ժ�Ǽ�")
    
    '��ȡ������Ժ�����Ϣ
    gstrSQL = " Select B.��Ժ����,D.���� As ��Ժ����,B.��Ժ����,B.�Ǽ��� As ����Ա,E.��������,Sum(Nvl(F.���,0)) As Ԥ����" & _
              " From �����ʻ� A,������ҳ B,������Ϣ C,���ű� D,zlbj.����Ŀ¼�� E,����Ԥ����¼ F" & _
              " Where A.����=" & TYPE_�Ͻ� & " And A.����ID=" & lng����ID & " And B.��ҳID=" & lng��ҳID & " ANd A.����ID=B.����ID And B.����ID=C.����ID And B.��ҳID=C.סԺ����" & _
              " And A.����ID=E.ID(+) And B.��Ժ����ID=D.ID(+) And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+) And F.��¼����(+)=1 " & _
              " Group by B.��Ժ����,D.����,B.��Ժ����,B.�Ǽ���,E.��������"
    Call OpenRecordset(rsTemp, "��ȡ������Ժ�����Ϣ")
    
    If Not ����_��ʼ Then Exit Function
    
    '������Ժ�ǼǼ�¼(�м�������Ŀ�)
    gstrSQL = " Insert Into ��Ժ�ǼǱ�" & _
              " (ID,��ᱣ�Ϻ�,����,סԺ��,��Ժʱ��,��Ժʱ��,��������," & _
              " ����,��λ��,Ԥ����,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
              " Values" & _
              " (��Ժ�ǼǱ�_ID.Nextval,'" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
              "'" & Format(rsTemp!��Ժ����, "yyyy.MM.dd") & "',NULL,'" & rsTemp!�������� & "','" & Nvl(rsTemp!��Ժ����) & "'," & _
              "'" & Nvl(rsTemp!��Ժ����) & "'," & Val(Nvl(rsTemp!Ԥ����, 0)) & ",'" & Nvl(rsTemp!����Ա, "ZLHIS") & "','��'," & IIf(gCominfo_�Ͻ�.blnOnLine, "'10'", "'00'") & ")"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = " Insert Into סԺ�ǼǱ�" & _
                  " (��ᱣ�Ϻ�,����,סԺ��,��Ժʱ��,��Ժʱ��,��������," & _
                  " ����,��λ��,Ԥ����,����Ա,�Ƿ����,ҽ�ƻ�������,ҽ�ƻ�������)" & _
                  " Values" & _
                  " ('" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "','" & IC_Data_�Ͻ�.���� & "','" & str����ǼǺ� & "'," & _
                  "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd") & "','','" & rsTemp!�������� & "','" & Nvl(rsTemp!��Ժ����) & "'," & _
                  "'" & Nvl(rsTemp!��Ժ����) & "'," & Val(Nvl(rsTemp!Ԥ����, 0)) & ",'" & Nvl(rsTemp!����Ա, "ZLHIS") & "','��'," & _
                  "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "')"
        If Not ExecuteSQL(gstrSQL) Then Exit Function
    End If
    
    If ����_�ύ Then
        ��Ժ�Ǽ�_�Ͻ� = True
    Else
        Call ����_�ع�
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_�ع�
End Function

Public Function ��Ժ�Ǽǳ���_�Ͻ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    'ɾ����Ժ�ǼǱ��ɣ�������εǼ��˷��ã�����������Ժ
    Dim strסԺ�� As String
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select סԺ�� From zlbj.��Ժ�ǼǱ� Where ��ᱣ�Ϻ�=" & _
            "   (Select ҽ���� From �����ʻ� Where ����ID=" & lng����ID & ")"
    Call OpenRecordset(rsTemp, "ȡ������סԺ��")
    strסԺ�� = rsTemp!סԺ��
    
    '��ȡ����סԺ�Ƿ�������
    gstrSQL = " Select Count(*) Records From ҩƷ������ϸ�� Where סԺ��='" & strסԺ�� & "'" & _
              " Union ALL" & _
              " Select Count(*) Records From ���Ʒ�����ϸ�� Where סԺ��='" & strסԺ�� & "'"
    gstrSQL = "Select SUM(Records) AS Records From (" & gstrSQL & ")"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If !Records > 0 Then
            MsgBox "�Ѿ��������ã����ܳ�����Ժ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If Not ����_��ʼ Then Exit Function
    
    gstrSQL = "Delete ��Ժ�ǼǱ� Where סԺ��='" & strסԺ�� & "'"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = "Delete סԺ�ǼǱ� Where סԺ��='" & strסԺ�� & "'"
        If Not ExecuteSQL(gstrSQL) Then Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ͻ� & ")"
    Call ExecuteProcedure("�����Ժ�Ǽ�")
    
    If ����_�ύ Then
        ��Ժ�Ǽǳ���_�Ͻ� = True
    Else
        Call ����_�ع�
    End If
End Function

Public Function ��Ժ�Ǽ�_�Ͻ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim blnTrans As Boolean
    Dim str��ᱣ�Ϻ� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '����ȫ��������ܳ�Ժ
    If ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "ֻ�з��ý����������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȡ�α��˵���ᱣ�Ϻ�
    gstrSQL = "Select ҽ���� From �����ʻ� Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ�α��˵���ᱣ�Ϻ�")
    str��ᱣ�Ϻ� = Nvl(rsTemp!ҽ����)
    
    If Not ����_��ʼ Then Exit Function
    
    If Not ���˱䶯��¼�ϴ�_�Ͻ�(lng����ID, lng��ҳID, False) Then
        Call ����_�ع�
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ͻ� & ")"
    Call ExecuteProcedure("�����Ժ�Ǽ�")
    
    '����סԺ����
    gstrSQL = " Update �����ʻ����� " & _
              " Set סԺ����=Nvl(סԺ����,0)+1" & _
              " Where ��ᱣ�Ϻ�='" & str��ᱣ�Ϻ� & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then 'SqlServer�ǿպ�����IsNULL();��Oracle��Nvl()
        gstrSQL = " Update �����ʻ����� " & _
                  " Set סԺ����=IsNull(סԺ����,0)+1" & _
                  " Where ��ᱣ�Ϻ�='" & str��ᱣ�Ϻ� & "'"
        If Not ExecuteSQL(gstrSQL) Then Exit Function
    End If
    
    If ����_�ύ Then
        ��Ժ�Ǽ�_�Ͻ� = True
    Else
        Call ����_�ع�
    End If
    blnTrans = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_�ع�
End Function

Public Function ��Ժ�Ǽǳ���_�Ͻ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    '��Ƴ�˵����������Ժ
    MsgBox "ҽ�����߲���������Ժ������ҽ��������ϵ��", vbInformation, gstrSysName
    ��Ժ�Ǽǳ���_�Ͻ� = False
End Function

Public Function �������_�Ͻ�(ByVal lng����ID As Long) As Currency
    '����: ��ȡ�α����˸����ʻ����
    '��Ϊ����ÿ�ξ���ʱ������£���סԺֻ�ܽ���һ�Σ����Կ���ֱ���Կ��е����Ϊ׼
    Dim rsTemp As New ADODB.Recordset
    
    With rsTemp
        If .State = 1 Then .Close
        .Open "Select Nvl(�ʻ����,0) ��� From �����ʻ� Where ����=" & TYPE_�Ͻ� & " And ����ID=" & lng����ID, gcnOracle
        �������_�Ͻ� = !���
    End With
End Function

Public Function סԺ�������_�Ͻ�(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim lng��ҳID As Long
    Dim dbl����ͳ�� As Double       '������ϸ�Ľ���ͳ����
    Dim dbl�����Ը� As Double
    Dim dbl�����ʻ� As Double
    Dim blnҩƷ As Boolean
    Dim blnTrans As Boolean
    Dim bln�ԷѲ��� As Boolean
    Dim str��Ŀ���� As String, str��Ŀ���� As String, str���� As String, str���� As String
    Dim dbl��� As Double, dbl���� As Double, dbl���� As Double
    
    Dim rsTemp As New ADODB.Recordset
    Dim cnOracle As New ADODB.Connection
    On Error GoTo ErrHand
    
    With gCominfo_�Ͻ�
        .dbl�����ܶ� = 0
        .dbl��ȱ��� = 0
        .dbl���ͳ�� = 0
        .dblͳ�ﱨ�� = 0
        .dblͳ���� = 0
    End With
    
    With cnOracle
        If .State = 1 Then .Close
        .Open gcnOracle.ConnectionString
    End With
    
    'ȡ�ò��˵���ᱣ�Ϻ�
    gstrSQL = " Select B.����,A.����,A.ҽ���� From �����ʻ� A,������Ϣ B" & _
              " Where A.����=" & TYPE_�Ͻ� & " And A.����ID=B.����ID And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ�ò��˵���ᱣ�Ϻ�")
    gCominfo_�Ͻ�.str��ᱣ�Ϻ� = rsTemp!ҽ����
    gCominfo_�Ͻ�.str��Ч���� = Nvl(rsTemp!����, 0)
    str���� = rsTemp!����
    
    'ȡ�ò��˵�סԺ��ˮ��
    gstrSQL = "Select סԺ�� From ��Ժ�ǼǱ� Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "' And ��Ժʱ�� Is Null"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "û���ҵ��ò�����Ч����Ժ��¼���޷����н��㣡", vbInformation, gstrSysName
            Exit Function
        End If
        gCominfo_�Ͻ�.str������ˮ�� = Nvl(!סԺ��)
    End With
    
    '����
    Dim strCard As String
    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
    Call ����ת��_�Ͻ�(strCard, True)
    If Not (IC_Data_�Ͻ�.��Ч���� = gCominfo_�Ͻ�.str��Ч���� And IC_Data_�Ͻ�.��ᱣ�Ϻ� = gCominfo_�Ͻ�.str��ᱣ�Ϻ�) Then
        Call IC_End(True)
        MsgBox "��ǰIC�����Ǹò��˵Ŀ���ÿ���ʧЧ������ҽ��������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ò��˵Ŀ�״̬
    '����������Ҫ���м���л�ȡ
    gstrSQL = "Select סԺ����,�ʻ�����,����ԭ��,��Ч����,����סԺ����,����ʱ��,����˵�� " & _
        " From �����ʻ����� Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
    If gCominfo_�Ͻ�.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then
            Call IC_End(True)
            Call gobjCenter.CloseConnector
            Exit Function
        End If
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsTemp
        If .RecordCount = 0 Then
            Call IC_End(True)
            MsgBox "û���ָò��˵���Ч��¼������������ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
        If Nvl(!�ʻ�����, "��") = "��" Then
            Call IC_End(True)
            MsgBox "�ò��˵��ʻ��Ѿ������ᣬֻ�����ֽ���㣡" & vbCrLf & "����ԭ��" & Nvl(!����ԭ��) & vbCrLf & "����˵����" & Nvl(!����˵��) & vbCrLf & "����ʱ�䣺" & Nvl(!����ʱ��), vbInformation, gstrSysName
            bln�ԷѲ��� = True
            Exit Function
        End If
        If Nvl(IC_Data_�Ͻ�.��Ч����, 0) <> Nvl(!��Ч����, 0) Then
            Call IC_End(True)
            MsgBox "��ǰ��IC��Ƭ��һ����Ч�Ŀ���", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    If gCominfo_�Ͻ�.blnOnLine Then gobjCenter.CloseConnector
    
    lng��ҳID = rsExse!��ҳID
    '��������Ժ���״̬�ϴ�������
    If Not ���˱䶯��¼�ϴ�_�Ͻ�(lng����ID, lng��ҳID) Then
        Call IC_End(True)
        Exit Function
    End If
    
    If Not ����_��ʼ() Then
        Call IC_End(True)
        Exit Function
    End If
    cnOracle.BeginTrans         '�ϴ�����ϸ�ʹ��ϱ�ǣ������ύ
    blnTrans = True
    
    '���ݴ����¼���������ͳ���ֻ����δ���㲿����ϸ���ǵ�Ȼ��δ�ϴ��Ŀ϶���û�м��㣩
    With rsExse
        Do While Not .EOF
            dbl��� = Nvl(!���, 0)
            If Nvl(!�Ƿ��ϴ�, 0) = 0 Then
                '����ͳ����
                str��Ŀ���� = "": str��Ŀ���� = "": str���� = ""
                blnҩƷ = (InStr(1, "5,6,7", !�շ����) <> 0)
                dbl���� = !����     '����������Ϊ����
                dbl��� = Nvl(!���, 0)
                dbl���� = Nvl(!���, 0) / dbl����
                
                'ȡ��ҽ����Ŀ�����Ϣ
                gstrSQL = " Select A.��Ŀ����,A.��Ŀ����,B.����,A.��ע As ���� From ����֧����Ŀ A,�շ�ϸĿ B " & _
                          " Where B.ID=" & !�շ�ϸĿID & " And B.ID=A.�շ�ϸĿID(+) And A.����(+)=" & TYPE_�Ͻ�
                Call OpenRecordset(rsTemp, "ȡ��ҽ����Ŀ�����Ϣ")
                If rsTemp.RecordCount <> 0 Then
                    str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
                    str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
                    str���� = Nvl(rsTemp!����)
                End If
                
                dbl����ͳ�� = Calcͳ����_��ϸ(blnҩƷ, str��Ŀ����, str��Ŀ����, dbl����, dbl����)
                
                'д���м��
                If blnҩƷ Then
                    gstrSQL = "" & _
                        " INSERT INTO ҩƷ������ϸ��" & _
                        " (ID,��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
                        " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                        " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                        " VALUES" & _
                        " (ҩƷ������ϸ��_ID.Nextval,'" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "','" & str���� & "','" & gCominfo_�Ͻ�.str������ˮ�� & "'," & _
                        "'" & IIf(str��Ŀ���� = "", gstrҩƷ����, str��Ŀ����) & "','" & Nvl(rsTemp!��Ŀ����, rsTemp!����) & "','" & IIf(str���� = "", gstrҩƷ����, str����) & "'," & _
                        "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!���, 0) & "," & _
                        "" & dbl����ͳ�� & ",0," & Nvl(!���, 0) - dbl����ͳ�� & "," & _
                        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & !ҽ�� & "','��','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
                    gcnGYBJYB.Execute gstrSQL
                Else
                    gstrSQL = "" & _
                        " INSERT INTO ���Ʒ�����ϸ��" & _
                        " (ID,��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
                        " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                        " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                        " VALUES" & _
                        " (���Ʒ�����ϸ��_ID.Nextval,'" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "','" & str���� & "','" & gCominfo_�Ͻ�.str������ˮ�� & "'," & _
                        "'" & IIf(str��Ŀ���� = "", gstr���ƴ���, str��Ŀ����) & "','" & Nvl(rsTemp!��Ŀ����, rsTemp!����) & "','" & IIf(str���� = "", gstr���ƴ���, str����) & "'," & _
                        "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!���, 0) & "," & _
                        "" & dbl����ͳ�� & ",0," & Nvl(!���, 0) - dbl����ͳ�� & "," & _
                        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & !ҽ�� & "','��','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
                    gcnGYBJYB.Execute gstrSQL
                End If
                
                'дҽ�����Ŀ⣨����סԺ�ǼǱ������в������ƣ����ң���λ�ţ���Ժʱ�����Ժʱ�䣬��ˣ������ű�����ͬ���ݲ�����д��
                If gCominfo_�Ͻ�.blnOnLine Then
                    If blnҩƷ Then
                        gstrSQL = "" & _
                            " INSERT INTO סԺδ����ҩƷ�����ռ��ʱ�" & _
                            " (��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
                            " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                            " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                            " VALUES" & _
                            " ('" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "','" & str���� & "','" & gCominfo_�Ͻ�.str������ˮ�� & "'," & _
                            "'" & IIf(str��Ŀ���� = "", gstrҩƷ����, str��Ŀ����) & "','" & Nvl(rsTemp!��Ŀ����, rsTemp!����) & "','" & IIf(str���� = "", gstrҩƷ����, str����) & "'," & _
                            "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!���, 0) & "," & _
                            "" & dbl����ͳ�� & ",0," & Nvl(!���, 0) - dbl����ͳ�� & "," & _
                            "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & Nvl(!ҽ��, "ZLHIS") & "','��','��')"
                        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
                    Else
                        gstrSQL = "" & _
                            " INSERT INTO סԺδ�������Ʒ����ռ��ʱ�" & _
                            " (��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
                            " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                            " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                            " VALUES" & _
                            " ('" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "','" & str���� & "','" & gCominfo_�Ͻ�.str������ˮ�� & "'," & _
                            "'" & IIf(str��Ŀ���� = "", gstr���ƴ���, str��Ŀ����) & "','" & Nvl(rsTemp!��Ŀ����, rsTemp!����) & "','" & IIf(str���� = "", gstr���ƴ���, str����) & "'," & _
                            "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!���, 0) & "," & _
                            "" & dbl����ͳ�� & ",0," & Nvl(!���, 0) - dbl����ͳ�� & "," & _
                            "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & Nvl(!ҽ��, "ZLHIS") & "','��','��')"
                        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
                    End If
                End If
                    
                '���ϴ����
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                cnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
            gCominfo_�Ͻ�.dbl�����ܶ� = gCominfo_�Ͻ�.dbl�����ܶ� + dbl���
            .MoveNext
        Loop
    End With
    
    If ����_�ύ Then
        cnOracle.CommitTrans
    Else
        Call ����_�ع�
        cnOracle.RollbackTrans
        Call IC_End(True)
        Exit Function
    End If
    blnTrans = False
    
    '���м����ȡ����δ����ķ�����ϸ��ͳ����
    gstrSQL = " Select Sum(Nvl(ͳ�������,0)) ͳ����" & _
              " From ҩƷ������ϸ�� A,��Ժ�ǼǱ� B" & _
              " Where A.��ᱣ�Ϻ�=B.��ᱣ�Ϻ� And Nvl(A.�Ƿ����,'��')='��' And A.�������='סԺ' And A.סԺ��=B.סԺ��" & _
              " And A.��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "' And A.סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
    gstrSQL = gstrSQL & " Union All" & _
              " Select Sum(Nvl(ͳ�������,0)) ͳ����" & _
              " From ���Ʒ�����ϸ�� A,��Ժ�ǼǱ� B" & _
              " Where A.��ᱣ�Ϻ�=B.��ᱣ�Ϻ� And Nvl(A.�Ƿ����,'��')='��' And A.�������='סԺ' And A.סԺ��=B.סԺ��" & _
              " And A.��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "' And A.סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
    gstrSQL = " Select Sum(ͳ����) ͳ���� From (" & gstrSQL & ")"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        gCominfo_�Ͻ�.dblͳ���� = !ͳ����
    End With
    
    '����ѡ��Ĳ����ٴμ������ͳ����
    gCominfo_�Ͻ�.dblͳ���� = Calcͳ����_����(gCominfo_�Ͻ�.dblͳ����, lng����ID)
    
    'ȡ����ۼ�
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = " Select IsNull(���,0) ���,IsNull(����סԺ����,0) ���ͳ���ۼ� " & _
                  " From �����ʻ����� " & _
                  " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'"
    Else
        gstrSQL = " Select Nvl(���,0) ���,nvl(����סԺ����,0) ���ͳ���ۼ� " & _
                  " From �����ʻ����� " & _
                  " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'"
    End If
    
    If gCominfo_�Ͻ�.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then
            Call gobjCenter.CloseConnector
            Call IC_End(True)
            Exit Function
        End If
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    With rsTemp
        If .RecordCount <> 0 Then
            gCominfo_�Ͻ�.dbl���ͳ�� = !���ͳ���ۼ�
            dbl�����ʻ� = !���
        End If
    End With
    
    If gCominfo_�Ͻ�.blnOnLine Then
        Call gobjCenter.CloseConnector
    End If
    
    '�ٸ��ݲ��˵Ĳα����֣���ҵ����ȣ��ֵ�����ó����յ�ͳ�ﱨ�����
    '����ǰ�����𸶣������ֻ��һ�����ߣ�����ÿ�ν��㶼Ҫ�����ߣ����ۼ���ʼ��Ҫ�õ�
    'If gCominfo_�Ͻ�.blnYearBase Then gCominfo_�Ͻ�.dbl��ȱ��� = Calcͳ����_�ֵ�(gCominfo_�Ͻ�.dbl���ͳ��, gCominfo_�Ͻ�.str��ᱣ�Ϻ�)
    'gCominfo_�Ͻ�.dblͳ�ﱨ�� = Calcͳ����_�ֵ�(gCominfo_�Ͻ�.dblͳ���� + IIf(gCominfo_�Ͻ�.blnYearBase, gCominfo_�Ͻ�.dbl���ͳ��, 0), gCominfo_�Ͻ�.str��ᱣ�Ϻ�)
    gCominfo_�Ͻ�.dblͳ�ﱨ�� = Calcͳ����_�ֵ�(gCominfo_�Ͻ�.dblͳ����, gCominfo_�Ͻ�.str��ᱣ�Ϻ�)
    
    'ʵ�ʱ��ν����ͳ�ﱨ�����
    'gCominfo_�Ͻ�.dblͳ�ﱨ�� = gCominfo_�Ͻ�.dblͳ�ﱨ�� - gCominfo_�Ͻ�.dbl��ȱ���
    gCominfo_�Ͻ�.dblͳ�ﱨ�� = gCominfo_�Ͻ�.dblͳ�ﱨ��
    
    '��������ʻ�֧����
    dbl�����Ը� = gCominfo_�Ͻ�.dbl�����ܶ� - gCominfo_�Ͻ�.dblͳ�ﱨ��
    dbl�����ʻ� = IIf(dbl�����ʻ� >= dbl�����Ը�, dbl�����Ը�, dbl�����ʻ�)
    
    If bln�ԷѲ��� Then
        dbl�����Ը� = gCominfo_�Ͻ�.dbl�����ܶ�
        dbl�����ʻ� = 0
        gCominfo_�Ͻ�.dblͳ�ﱨ�� = 0
        gCominfo_�Ͻ�.dblͳ���� = 0
    End If
    
    סԺ�������_�Ͻ� = "�����ʻ�;" & dbl�����ʻ� & ";1"
    סԺ�������_�Ͻ� = סԺ�������_�Ͻ� & "|ҽ������;" & gCominfo_�Ͻ�.dblͳ�ﱨ�� & ";0"
    
    Call IC_End(True)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then
        Call ����_�ع�
        cnOracle.RollbackTrans
    End If
    Call IC_End(True)
End Function

Public Function סԺ����_�Ͻ�(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    On Error GoTo ErrHand
    Dim str���� As String
    Dim strCard As String
    Dim blnTrans As Boolean
    Dim dbl�����ʻ� As Double
    Dim lng��ҳID As Long
    Dim intסԺ���� As Integer
    Dim rsTemp As New ADODB.Recordset
    'ҽ��Ҫ���Ժ���������Զ���Ժ����HIS���в������ƣ���Ҫʵʩ��Աע��
    
    '����
    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
    Call ����ת��_�Ͻ�(strCard, True)
    If Not (IC_Data_�Ͻ�.��Ч���� = gCominfo_�Ͻ�.str��Ч���� And IC_Data_�Ͻ�.��ᱣ�Ϻ� = gCominfo_�Ͻ�.str��ᱣ�Ϻ�) Then
        Call IC_End(True)
        MsgBox "��ǰIC�����Ǹò��˵Ŀ���ÿ���ʧЧ������ҽ��������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    str���� = IC_Data_�Ͻ�.����
    
    'ȡ���˵���ҳID
    gstrSQL = "Select nvl(סԺ����,0) AS ��ҳID From ������Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ���˵���ҳID")
    lng��ҳID = rsTemp!��ҳID
    
    'ȡ���˵�סԺ����
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = "Select IsNull(סԺ����,0) סԺ���� From �����ʻ����� Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
    Else
        gstrSQL = "Select Nvl(סԺ����,0) סԺ���� From �����ʻ����� Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
    End If
    
    If gCominfo_�Ͻ�.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then
            Call gobjCenter.CloseConnector
            Call IC_End(True)
            Exit Function
        End If
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    
    intסԺ���� = rsTemp!סԺ����
    
    If gCominfo_�Ͻ�.blnOnLine Then
        Call gobjCenter.CloseConnector
    End If
    
    'ȡ���ν���ʵ�ʸ����ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & TYPE_�Ͻ� & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���θ����ʻ�֧����")
    dbl�����ʻ� = 0
    If Not rsTemp.EOF Then
        dbl�����ʻ� = rsTemp!�����ʻ�
    End If
    
    If Not ����_��ʼ Then
        Call IC_End(True)
        Exit Function
    End If
    blnTrans = True
    
    '����д���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ͻ� & "," & lng����ID & "," & _
        Format(zlDataBase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & intסԺ���� & "," & 0 & "," & 0 & "," & 0 & "," & _
        gCominfo_�Ͻ�.dbl�����ܶ� & "," & gCominfo_�Ͻ�.dbl�����ܶ� - gCominfo_�Ͻ�.dblͳ�ﱨ�� - dbl�����ʻ� & ",0," & _
        gCominfo_�Ͻ�.dblͳ���� & "," & gCominfo_�Ͻ�.dblͳ�ﱨ�� & ",0,0," & dbl�����ʻ� & ",'" & gCominfo_�Ͻ�.str������ˮ�� & "')"
    Call ExecuteProcedure("���汣�ս����¼")
    
    '���м���иò������з�����ϸ�Ľ����־����
    gstrSQL = " Update ҩƷ������ϸ�� " & _
              " Set �Ƿ����='��' " & _
              " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'" & _
              " And סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
    gcnGYBJYB.Execute gstrSQL
    gstrSQL = " Update ���Ʒ�����ϸ�� " & _
              " Set �Ƿ����='��' " & _
              " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'" & _
              " And סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
    gcnGYBJYB.Execute gstrSQL
    
    gstrSQL = "" & _
        " INSERT INTO ҽ�Ʒ���֧����ϸ��" & _
        " (ID,��ᱣ�Ϻ�,����,����סԺ��,����ʱ��,�������,�ܷ���,ͳ�����֧��," & _
        " �����ʻ�֧��,�����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ��ϴ�)" & _
        " VALUES" & _
        " (ҽ�Ʒ���֧����ϸ��_ID.Nextval,'" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "','" & str���� & "','" & gCominfo_�Ͻ�.str������ˮ�� & "'," & _
        "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & gCominfo_�Ͻ�.dbl�����ܶ� & "," & _
        gCominfo_�Ͻ�.dblͳ�ﱨ�� & "," & dbl�����ʻ� & "," & gCominfo_�Ͻ�.dbl�����ܶ� - gCominfo_�Ͻ�.dblͳ�ﱨ�� - dbl�����ʻ� & "," & _
        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_�Ͻ�.blnOnLine Then
        '�����ĵ�δ������ϸת�������ϸ����
        gstrSQL = "" & _
            " INSERT INTO ҩƷ������ϸ��" & _
            "     (��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
            "     ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
            "     �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա)" & _
            " Select ��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
            "     ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
            "     �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա" & _
            " From סԺδ����ҩƷ�����ռ��ʱ�" & _
            " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'" & _
            " And סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        
        gstrSQL = "" & _
            " INSERT INTO ���Ʒ�����ϸ��" & _
            "     (��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
            "     ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
            "     �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա)" & _
            " Select ��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
            "     ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
            "     �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա" & _
            " From סԺδ�������Ʒ����ռ��ʱ�" & _
            " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'" & _
            " And סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        
        'ɾ������δ������ϸ
        gstrSQL = " Delete סԺδ����ҩƷ�����ռ��ʱ�" & _
                  " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'" & _
                  " And סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        gstrSQL = " Delete סԺδ�������Ʒ����ռ��ʱ�" & _
                  " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'" & _
                  " And סԺ��='" & gCominfo_�Ͻ�.str������ˮ�� & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        
        '����ʱ�����ν�����ϸд�����Ŀ�
        gstrSQL = "" & _
            " INSERT INTO ҽ�Ʒ���֧����ϸ��" & _
            " (��ᱣ�Ϻ�,����,����סԺ��,����ʱ��,�������,�ܷ���,ͳ�����֧��," & _
            " �����ʻ�֧��,�����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա)" & _
            " VALUES" & _
            " ('" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "','" & str���� & "','" & gCominfo_�Ͻ�.str������ˮ�� & "'," & _
            "'" & Format(zlDataBase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & gCominfo_�Ͻ�.dbl�����ܶ� & "," & _
            gCominfo_�Ͻ�.dblͳ�ﱨ�� & "," & dbl�����ʻ� & "," & gCominfo_�Ͻ�.dbl�����ܶ� - gCominfo_�Ͻ�.dblͳ�ﱨ�� - dbl�����ʻ� & "," & _
            "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "')"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    '������Ժ�ǼǱ�
    gstrSQL = " Update ��Ժ�ǼǱ�" & _
              " Set �Ƿ����='��'" & _
              " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "' And ��Ժʱ�� Is NULL"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = " Update סԺ�ǼǱ�" & _
                  " Set �Ƿ����='��'" & _
                  " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "' And IsNull(��Ժʱ��,'')=''"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    '���ʻ������Ŀ�ĸ����ʻ�����Ҫ���£��м��ĸ����ʻ�����Ҫ���£�
    'סԺ����Ҫ���µ�ǰסԺ�����ֶΣ�ͳ��֧����סԺ������
    IC_Data_�Ͻ�.�����ʻ���� = IC_Data_�Ͻ�.�����ʻ���� - dbl�����ʻ�
    IC_Data_�Ͻ�.���������� = Format(zlDataBase.Currentdate, "yyyy.MM.dd")
    IC_Data_�Ͻ�.������ҽԺ���� = gCominfo_�Ͻ�.strHospitalCode
    Call ����ת��_�Ͻ�(strCard, False)
    
    '�ڳ�Ժʱ���Զ�����סԺ���� Set סԺ����=Nvl(סԺ����,0)+1
    gstrSQL = " Update �����ʻ����� " & _
              " Set ���=Nvl(���,0)-" & Val(dbl�����ʻ�) & "," & _
              "     ����סԺ����=Nvl(����סԺ����,0)+" & gCominfo_�Ͻ�.dblͳ���� & _
              " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_�Ͻ�.blnOnLine Then 'SqlServer�ǿպ�����IsNULL();��Oracle��Nvl()
        gstrSQL = " Update �����ʻ����� " & _
                  " Set ����֧��=IsNull(����֧��,0)+" & Val(dbl�����ʻ�) & "," & _
                  "     �ۼ�֧��=IsNull(�ۼ�֧��,0)+" & Val(dbl�����ʻ�) & "," & _
                  "     ���=IsNull(���,0)-" & Val(dbl�����ʻ�) & "," & _
                  "     ����סԺ����=IsNull(����סԺ����,0)+" & gCominfo_�Ͻ�.dblͳ���� & _
                  " Where ��ᱣ�Ϻ�='" & gCominfo_�Ͻ�.str��ᱣ�Ϻ� & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    If Not gobjCenter.IC_WriteCard(strCard) Then
        Call ����_�ع�
        Call IC_End(True)
        Exit Function
    End If
    
    If ����_�ύ Then
        סԺ����_�Ͻ� = True
    Else
        Call ����_�ع�
    End If
    blnTrans = False
    
    Call IC_End
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_�ع�
    Call IC_End(True)
End Function

Public Function סԺ�������_�Ͻ�(lng����ID As Long) As Boolean
    'ֻ����������������ϣ�סԺ������
    MsgBox "ҽ���ӿڲ�֧������סԺ���㵥��", vbInformation, gstrSysName
    סԺ�������_�Ͻ� = False
End Function

Public Function �����ϴ�_�Ͻ�(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal str���ݺ� As String) As Boolean
    Dim blnTrans As Boolean                 '��ǰ�Ƿ���������
    Dim blnInsure As Boolean                '�����Ƿ���Ϊҽ�����˵���ݽ��о���
    Dim blnҩƷ As Boolean
    Dim int��� As Integer
    Dim lng����ID As Long
    Dim dblͳ���� As Double
    Dim str����ǼǺ� As String, str��ᱣ�Ϻ� As String, str���� As String
    Dim rsDetail As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '�����β����Ĵ���ȫ�����������Ŀ��δ���ռ��ʱ��м�����ϸ����
    If Not ����_��ʼ Then Exit Function
    blnTrans = True
    
    gstrSQL = " Select A.����ID,A.�շ����,A.�շ�ϸĿID,A.���,A.����ʱ��," & _
              " round(A.ʵ�ս��,2) ʵ�ս��,A.ʵ�ս��/(A.����*A.����) As ����,(A.����*A.����) AS ����," & _
              " B.��Ŀ����,B.��Ŀ����,B.��ע As ����,C.����" & _
              " From ���˷��ü�¼ A,����֧����Ŀ B,�շ�ϸĿ C" & _
              " Where A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=" & TYPE_�Ͻ� & _
              " And A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & " And A.NO='" & str���ݺ� & "'" & _
              " And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.�Ƿ��ϴ�,0)=0"
    Call OpenRecordset(rsDetail, "��ȡ���Ŵ�����ϸ")
    
    With rsDetail
        Do While Not .EOF
            If lng����ID <> !����ID Then
                '��鱾���Ƿ���ҽ�������Ժ
                gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=" & !����ID & " And A.��ҳID=B.סԺ���� And A.����=" & TYPE_�Ͻ�
                Call OpenRecordset(rsInsure, "�ж��Ƿ�ҽ������")
                blnInsure = (rsInsure!Records = 1)
                If blnInsure Then
                    lng����ID = !����ID
                    '��ȡ���˵���ᱣ�Ϻ�
                    gstrSQL = "Select ҽ���� As ��ᱣ�Ϻ� From �����ʻ� Where ����ID=" & lng����ID
                    Call OpenRecordset(rsTemp, "��ȡ���˵���ᱣ�Ϻ�")
                    str��ᱣ�Ϻ� = Nvl(rsTemp!��ᱣ�Ϻ�)
                    'ȡ������Ժ�ľ���ǼǺ�(���м�����Ժ�ǼǱ���ȡ)
                    gstrSQL = " Select B.ҽ���� As ��ᱣ�Ϻ�,C.����,A.סԺ��" & _
                              " From zlbj.��Ժ�ǼǱ� A,�����ʻ� B,������Ϣ C" & _
                              " Where A.��ᱣ�Ϻ�=B.ҽ���� And A.��Ժʱ�� Is NULL And B.����ID=C.����ID" & _
                              " And B.����=" & TYPE_�Ͻ� & " And B.����ID=" & lng����ID
                    With rsTemp
                        If .State = 1 Then .Close
                        .Open gstrSQL, gcnOracle
                        If .RecordCount = 0 Then
                            MsgBox "û���ҵ��ò���[��ᱣ�Ϻţ�" & str��ᱣ�Ϻ� & "]����Ч��Ժ��¼,���ܸò����Ѿ���Ժ��", vbInformation, gstrSysName
                            Call ����_�ع�
                            Exit Function
                        End If
                        str����ǼǺ� = Nvl(!סԺ��)
                        str��ᱣ�Ϻ� = Nvl(!��ᱣ�Ϻ�)
                        str���� = Nvl(!����)
                    End With
                    
                    '�жϿ�����Ч�ԣ������Ч���ж�Ԥ������Ƿ�������з����ܶ���С����ʾ����׼���浥��
                    If Not CheckCard(str��ᱣ�Ϻ�) Then
                        If Not BalanceLack(lng����ID) Then
                            MsgBox "�ò���[��ᱣ�Ϻţ�" & str��ᱣ�Ϻ� & "]�Ŀ��ѱ�����������ȫ�����ֽ𣬶�Ԥ����㣬��ɿ", vbInformation, gstrSysName
                            Call ����_�ع�
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            If blnInsure Then
                int��� = !���
                blnҩƷ = (InStr(1, "5,6,7", !�շ����) <> 0)
                '���㵱����ϸ�Ľ���ͳ����
                dblͳ���� = Calcͳ����_��ϸ(blnҩƷ, Nvl(!��Ŀ����), Nvl(!��Ŀ����), Nvl(!����, 0), Nvl(!����, 0))
                
                'д�м��
                If blnҩƷ Then
                    gstrSQL = "" & _
                        " INSERT INTO ҩƷ������ϸ��" & _
                        " (ID,��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
                        " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                        " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                        " VALUES" & _
                        " (ҩƷ������ϸ��_ID.Nextval,'" & str��ᱣ�Ϻ� & "','" & str���� & "','" & str����ǼǺ� & "'," & _
                        "'" & Nvl(!��Ŀ����, gstrҩƷ����) & "','" & Nvl(!��Ŀ����, !����) & "','" & Nvl(!����, gstrҩƷ����) & "'," & _
                        "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!ʵ�ս��, 0) & "," & _
                        "" & dblͳ���� & ",0," & Nvl(!ʵ�ս��, 0) - dblͳ���� & "," & _
                        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','��','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
                    gcnGYBJYB.Execute gstrSQL
                Else
                    gstrSQL = "" & _
                        " INSERT INTO ���Ʒ�����ϸ��" & _
                        " (ID,��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
                        " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                        " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                        " VALUES" & _
                        " (���Ʒ�����ϸ��_ID.Nextval,'" & str��ᱣ�Ϻ� & "','" & str���� & "','" & str����ǼǺ� & "'," & _
                        "'" & Nvl(!��Ŀ����, gstr���ƴ���) & "','" & Nvl(!��Ŀ����, !����) & "','" & Nvl(!����, gstr���ƴ���) & "'," & _
                        "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!ʵ�ս��, 0) & "," & _
                        "" & dblͳ���� & ",0," & Nvl(!ʵ�ս��, 0) - dblͳ���� & "," & _
                        "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','��','" & IIf(gCominfo_�Ͻ�.blnOnLine, "��", "��") & "')"
                    gcnGYBJYB.Execute gstrSQL
                End If
                
                'дҽ�����Ŀ⣨����סԺ�ǼǱ������в������ƣ����ң���λ�ţ���Ժʱ�����Ժʱ�䣬��ˣ������ű�����ͬ���ݲ�����д��
                If gCominfo_�Ͻ�.blnOnLine Then
                    If blnҩƷ Then
                        gstrSQL = "" & _
                            " INSERT INTO סԺδ����ҩƷ�����ռ��ʱ�" & _
                            " (��ᱣ�Ϻ�,����,סԺ��,ҩƷ����,ҩƷ����,ҩƷ����," & _
                            " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                            " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                            " VALUES" & _
                            " ('" & str��ᱣ�Ϻ� & "','" & str���� & "','" & str����ǼǺ� & "'," & _
                            "'" & Nvl(!��Ŀ����, gstrҩƷ����) & "','" & Nvl(!��Ŀ����, !����) & "','" & Nvl(!����, gstrҩƷ����) & "'," & _
                            "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!ʵ�ս��, 0) & "," & _
                            "" & dblͳ���� & ",0," & Nvl(!ʵ�ս��, 0) - dblͳ���� & "," & _
                            "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','��','��')"
                        If Not ExecuteSQL(gstrSQL) Then
                            Call ����_�ع�
                            Exit Function
                        End If
                    Else
                        gstrSQL = "" & _
                            " INSERT INTO סԺδ�������Ʒ����ռ��ʱ�" & _
                            " (��ᱣ�Ϻ�,����,סԺ��,������Ŀ����,������Ŀ����,�������," & _
                            " ����ʱ��,�������,�ܷ���,ͳ�������,�����ʻ����," & _
                            " �����Ը����,ҽ�ƻ�������,ҽ�ƻ�������,����Ա,�Ƿ����,�Ƿ��ϴ�)" & _
                            " VALUES" & _
                            " ('" & str��ᱣ�Ϻ� & "','" & str���� & "','" & str����ǼǺ� & "'," & _
                            "'" & Nvl(!��Ŀ����, gstr���ƴ���) & "','" & Nvl(!��Ŀ����, !����) & "','" & Nvl(!����, gstr���ƴ���) & "'," & _
                            "'" & Format(!����ʱ��, "yyyy.MM.dd HH:mm:ss") & "','סԺ'," & Nvl(!ʵ�ս��, 0) & "," & _
                            "" & dblͳ���� & ",0," & Nvl(!ʵ�ս��, 0) - dblͳ���� & "," & _
                            "'" & gCominfo_�Ͻ�.strHospitalCode & "','" & gCominfo_�Ͻ�.strHospitalName & "','" & UserInfo.���� & "','��','��')"
                        If Not ExecuteSQL(gstrSQL) Then
                            Call ����_�ع�
                            Exit Function
                        End If
                    End If
                End If
                
                '���ϴ����
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & str���ݺ� & "'," & int��� & "," & int���� & "," & int״̬ & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
            .MoveNext
        Loop
    End With
    
    If ����_�ύ Then
        �����ϴ�_�Ͻ� = True
    Else
        Call ����_�ع�
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_�ع�
End Function

Public Function ���²���_�Ͻ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim lng����ID As Long
    Dim blnTrans As Boolean
    Dim str��ᱣ�Ϻ� As String
    Dim str�������� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '���²���
    If Not frm����ѡ��_�Ͻ�.����ѡ��(lng����ID, lng����ID, str��������) Then Exit Function
    
    If Not ����_��ʼ Then Exit Function
    gcnOracle.BeginTrans
    blnTrans = True
    
    '���±����ʻ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�Ͻ� & ",'����ID','" & lng����ID & "')"
    Call ExecuteProcedure("���²�����Ϣ")
    
    '�����м�������Ŀ�
    gstrSQL = " Update ��Ժ�ǼǱ� " & _
              " Set ��������='" & str�������� & "'" & _
              " Where ��ᱣ�Ϻ�='" & str��ᱣ�Ϻ� & "'"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_�Ͻ�.blnOnLine Then
        gstrSQL = " Update סԺ�ǼǱ� " & _
                  " Set ��������='" & str�������� & "'" & _
                  " Where ��ᱣ�Ϻ�='" & str��ᱣ�Ϻ� & "'"
        If Not ExecuteSQL(gstrSQL) Then Exit Function
    End If
    
    If ����_�ύ Then
        gcnOracle.CommitTrans
        ���²���_�Ͻ� = True
    Else
        gcnOracle.RollbackTrans
        Call ����_�ع�
    End If
    
    blnTrans = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ����_�ع�
    End If
End Function

Public Function ���˱䶯��¼�ϴ�_�Ͻ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln��ʼ���� As Boolean = True) As Boolean
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    '�����˿��ң���λ������ҽ��������ȼ������仯ʱ�������¼�����ҽ��ֻ���Ŀ��ҡ���λ��Ԥ�����벡�����ƣ�
    gstrSQL = " Select A.ҽ���� As ��ᱣ�Ϻ�,to_Char(B.��Ժ����,'yyyy.MM.dd') AS ��Ժ����," & _
              " D.���� As ��ǰ����,C.��ǰ����,E.��������,Sum(Nvl(F.���,0)) As Ԥ����" & _
              " From �����ʻ� A,������ҳ B,������Ϣ C,���ű� D,zlbj.����Ŀ¼�� E,����Ԥ����¼ F" & _
              " Where A.����=" & TYPE_�Ͻ� & " And A.����ID=" & lng����ID & " And B.��ҳID=" & lng��ҳID & " ANd A.����ID=B.����ID And B.����ID=C.����ID And B.��ҳID=C.סԺ����" & _
              " And A.����ID=E.ID(+) And C.��ǰ����ID=D.ID(+) And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+) And F.��¼����(+)=1 " & _
              " Group by A.ҽ����,B.��Ժ����,D.����,C.��ǰ����,B.�Ǽ���,E.��������"
    Call OpenRecordset(rsTemp, "��ȡ������Ժ�����Ϣ")
    
    If bln��ʼ���� Then
        blnTrans = True
        If Not ����_��ʼ Then Exit Function
    End If
    
    '�޸���Ժ�ǼǼ�¼(�м�������Ŀ�)
    '���ڳ�Ժʱ���޵�ǰ�����봲λ�ţ���ˣ�ֻҪ��Ժʱ�䲻Ϊ�գ����ٸ��¿����봲λ��
    If Nvl(rsTemp!��Ժ����) = "" Then
        gstrSQL = " Update ��Ժ�ǼǱ�" & _
                  " Set ��������='" & Nvl(rsTemp!��������) & "'," & _
                  "     ����='" & Nvl(rsTemp!��ǰ����) & "'," & _
                  "     ��λ��='" & Nvl(rsTemp!��ǰ����) & "'," & _
                  "     Ԥ����=" & Nvl(rsTemp!Ԥ����, 0) & _
                  " Where ��ᱣ�Ϻ�='" & rsTemp!��ᱣ�Ϻ� & "' And ��Ժʱ�� Is NULL"
    Else
        gstrSQL = " Update ��Ժ�ǼǱ�" & _
                  " Set ��������='" & Nvl(rsTemp!��������) & "'," & _
                  "     ��Ժʱ��='" & Nvl(rsTemp!��Ժ����) & "'" & _
                  " Where ��ᱣ�Ϻ�='" & rsTemp!��ᱣ�Ϻ� & "' And ��Ժʱ�� Is NULL"
    End If
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_�Ͻ�.blnOnLine Then
        If Nvl(rsTemp!��Ժ����) = "" Then
            gstrSQL = " Update סԺ�ǼǱ�" & _
                      " Set ��������='" & Nvl(rsTemp!��������) & "'," & _
                      "     ����='" & Nvl(rsTemp!��ǰ����) & "'," & _
                      "     ��λ��='" & Nvl(rsTemp!��ǰ����) & "'," & _
                      "     Ԥ����=" & Nvl(rsTemp!Ԥ����, 0) & _
                      " Where ��ᱣ�Ϻ�='" & rsTemp!��ᱣ�Ϻ� & "' And IsNull(��Ժʱ��,'')=''"
        Else
            gstrSQL = " Update סԺ�ǼǱ�" & _
                      " Set ��������='" & Nvl(rsTemp!��������) & "'," & _
                      "     ��Ժʱ��='" & Nvl(rsTemp!��Ժ����) & "'" & _
                      " Where ��ᱣ�Ϻ�='" & rsTemp!��ᱣ�Ϻ� & "' And IsNull(��Ժʱ��,'')=''"
        End If
        If Not ExecuteSQL(gstrSQL, bln��ʼ����) Then Exit Function
    End If
    
    If bln��ʼ���� Then
        If ����_�ύ Then
            ���˱䶯��¼�ϴ�_�Ͻ� = True
        Else
            Call ����_�ع�
        End If
    Else
        ���˱䶯��¼�ϴ�_�Ͻ� = True
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_�ع�
End Function

Private Function Calcͳ����_��ϸ(ByVal blnҩƷ As Boolean, ByVal str��Ŀ���� As String, ByVal str��Ŀ���� As String, _
    ByVal dbl���� As Double, ByVal dbl���� As Double) As Double
    Dim dblͳ���� As Double, dbl�Ը����� As Double, dbl�𸶽�� As Double, dbl�޼� As Double
    Dim dbl��� As Double
    Dim bln���� As Boolean
    Dim strҽԺ���� As String
    Dim rsCalc As New ADODB.Recordset
    '���㵥����ϸ�Ľ���ͳ����
    
    '��ȡҽԺ����
    If Not blnҩƷ Then
        gstrSQL = "Select ���� From ҽ�ƻ������������ Where ��λ����='" & gCominfo_�Ͻ�.strHospitalCode & "'"
        With rsCalc
            If .State = 1 Then .Close
            .Open gstrSQL, gcnGYBJYB
            strҽԺ���� = !����
        End With
    End If
    
    '��ȡ��Ŀ�Ļ�����Ϣ
    dbl��� = dbl���� * dbl����
    bln���� = (dbl��� < 0)
    gstrSQL = " Select nvl(�����Ը�����,0) As �Ը�����,Nvl(�����𸶽��,0) As �𸶽��" & _
              "" & IIf(blnҩƷ, "", ",nvl(һ��ҽԺ����,0) һ��,Nvl(����ҽԺ����,0) ����,Nvl(����ҽԺ����,0) ����") & _
              " From " & IIf(blnҩƷ, "ҩƷĿ¼��", "������Ŀ��") & _
              " Where " & IIf(blnҩƷ, "ҩƷ����", "������Ŀ����") & "='" & str��Ŀ���� & "'" & _
              " And " & IIf(blnҩƷ, "��������", "������Ŀ����") & "='" & str��Ŀ���� & "'"
    With rsCalc
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            '�����δ�������Ŀ�����ϵͳ������סԺ����ҩƷ�����Ը���Ϊ�棬�����ϸΪȫ�Ը�������ȫ������ͳ��
            dbl�Ը����� = IIf(gCominfo_�Ͻ�.blnPhysicCash, 100, 0)
            dbl�𸶽�� = 0
        Else
            dbl�Ը����� = IIf(gCominfo_�Ͻ�.blnPhysicCash, !�Ը�����, 0)
            dbl�𸶽�� = IIf(gCominfo_�Ͻ�.blnPhysicCash, !�𸶽��, 0)
        End If
        
        '���ж�ʵ�ʵ��ۣ���������޼ۣ����޼�Ϊ׼
        If Not blnҩƷ And .RecordCount > 0 Then
            dbl�޼� = IIf(strҽԺ���� = "һ��ҽԺ", !һ��, IIf(strҽԺ���� = "����ҽԺ", !����, !����))
        End If
        
        '���¼��㱾����ϸ��ʵ�ʽ��
        dbl��� = Abs(IIf(dbl���� >= dbl�޼� And dbl�޼� <> 0, dbl�޼�, dbl����) * dbl����) * IIf(bln����, -1, 1)
        
        '�ȿ��𸶽�������С�ڵ����㣬��ֱ���˳�
        dbl��� = (Abs(dbl���) - dbl�𸶽��) * IIf(bln����, -1, 1)
        If dbl��� <= 0 And Not bln���� Then Exit Function
        Calcͳ����_��ϸ = Round(dbl��� * (100 - dbl�Ը�����) / 100, 2)
    End With
End Function

Private Function Calcͳ����_����(ByVal dbl����ͳ�� As Double, ByVal lng����ID As Long) As Double
    Dim dbl�Ը����� As Double, dbl�𸶽�� As Double
    Dim lng����ID As Long
    Dim rsCalc As New ADODB.Recordset
    '����ѡ��Ĳ����ٴμ������ͳ����
    
    '��ȡ�ò��˵�ǰ�Ĳ���ID
    gstrSQL = "Select Nvl(����ID,0) AS ����ID From �����ʻ� Where ����=" & TYPE_�Ͻ� & " And ����ID=" & lng����ID
    Call OpenRecordset(rsCalc, "��ȡ�ò��˵�ǰ�Ĳ���ID")
    lng����ID = rsCalc!����ID
    
    '��ȡ�ò��ֵ�����
    gstrSQL = "Select Nvl(�����Ը�����,0) AS �Ը�����,Nvl(�����𸶽��,0) AS �𸶽�� From ����Ŀ¼�� Where ID=" & lng����ID
    With rsCalc
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "û���ҵ����ּ�¼���޷����סԺ���㣬����������ϵ��", vbInformation, gstrSysName
            Exit Function
'            dbl�Ը����� = IIf(gCominfo_�Ͻ�.blnDiseaseCash, 100, 0)
'            dbl�𸶽�� = 0
        Else
            dbl�Ը����� = IIf(gCominfo_�Ͻ�.blnDiseaseCash, !�Ը�����, 0)
            dbl�𸶽�� = IIf(gCominfo_�Ͻ�.blnDiseaseCash, !�𸶽��, 0)
        End If
    End With
    
    '�ȿ�ȥ�𸶽��ٰ���������ͳ����
    dbl����ͳ�� = dbl����ͳ�� - dbl�𸶽��
    If dbl����ͳ�� <= 0 Then Exit Function
    dbl����ͳ�� = dbl����ͳ�� * (100 - dbl�Ը�����) / 100
    Calcͳ����_���� = dbl����ͳ��
End Function

Private Function Calcͳ����_�ֵ�(ByVal dblͳ���� As Double, ByVal str��ᱣ�Ϻ� As String) As Double
    Dim intDo As Integer, intLoops As Integer   '����ѭ������ģ������סԺͳ���ֻѭ��һ��
    Dim blnMatch As Boolean
    
    Dim bln��һ�� As Boolean
    Dim dblʵ������ As Double
    
    Dim str�μ����� As String
    Dim str�������� As String
    Dim str��ҵ��� As String
    Dim strҽԺ���� As String
    Dim str�Ա� As String
    Dim lng���� As Long
    Dim lng���� As Long
    
    Dim dbl���ͳ���ۼ� As Double
    Dim dbl�����ۼ� As Double
    Dim dbl������ As Double
    Dim dbl���뱨�� As Double
    Dim dbl���� As Double
    Dim dbl���� As Double
    Dim dbl�𸶽�� As Double
    Dim dbl�������� As Double
    Dim rsBase As New ADODB.Recordset       '�α��˻�����Ϣ
    Dim rsDisease As New ADODB.Recordset    '���ֻ�����Ϣ
    Dim rsRule As New ADODB.Recordset       '��������
    '�������ͳ�������ͳ�ﱨ�����
    
    dbl���ͳ���ۼ� = gCominfo_�Ͻ�.dbl���ͳ��
    '��ȡ���ò��˵Ļ�����Ϣ
    gstrSQL = "" & _
        " Select A.�Ա�,A.��������,A.�μӹ���ʱ��,A.��ҵ���,A.�α�����,B.���� AS ҽԺ����" & _
        " From �����ʻ����� A,ҽ�ƻ������������ B" & _
        " Where A.��ᱣ�Ϻ�='" & str��ᱣ�Ϻ� & "'"
    With rsBase
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "û���ҵ��ò��˵Ļ�����Ϣ���޷����н��㣡[��ȡ���˻�����Ϣ]", vbInformation, gstrSysName
            Exit Function
        Else
            str�μ����� = Nvl(!�α�����, String(8, "0"))
            str�μ����� = Mid(str�μ�����, 1, 4)        'ֻ��ǰ��λ��������
            str��ҵ��� = !��ҵ���
            strҽԺ���� = !ҽԺ����
            str�Ա� = !�Ա�
            lng���� = GetAge(Format(zlDataBase.Currentdate, "yyyy-MM-dd"), Replace(!��������, ".", "-"))
            lng���� = GetAge(Format(zlDataBase.Currentdate, "yyyy-MM-dd"), Replace(!�μӹ���ʱ��, ".", "-"))
        End If
    End With
    
    '��α����ֹ���λ��ֻ��ǰ��λ���ã���ǰ��λ�����λ���ų�
    '��ȡ�������֣���������ԭʼ˳��λ��Ӧstr�μ����֣�
    gstrSQL = "Select Rownum ���,�������� From �α����ֱ� Where ��������='ҽ�Ʊ���'" '����ʱҲֻ���˻������ֵ���ҽ�Ʊ��յ�
    With rsDisease
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "�α����ֲ�ȫ���޷����н��㣬����������ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    '׼�����зֵ�����
    intLoops = IIf(Right(str�μ�����, 1) = "1", 4, 3)
    intDo = IIf(Right(str�μ�����, 1) = "1", 4, 1)
    bln��һ�� = True
    For intDo = intDo To intLoops
        If Mid(str�μ�����, intDo, 1) = "1" Then
            rsDisease.Filter = "���=" & intDo
            If rsDisease.RecordCount = 0 Then
                rsDisease.Filter = 0
                MsgBox "�α����ֲ�ȫ���޷����н��㣬����������ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
            str�������� = rsDisease!��������
            
            '��ȡҽ�Ʊ��ս������ߣ�ȡ�����Ŀ��ܶ���������������������ֻ������һ����¼��
            gstrSQL = " Select �����,�����,��ҵ���,���ö�,�𸶽��,�������� " & _
                      " From ҽ�Ʒ�֧�����߱�" & _
                      " Where ��������='" & str�������� & "' And Nvl(�Ա�,'" & str�Ա� & "')='" & str�Ա� & "' And Nvl(ҽԺ����,'" & strҽԺ���� & "')='" & strҽԺ���� & "'" & _
                      " Order By ���ö�"
            With rsRule
                If .State = 1 Then .Close
                .Open gstrSQL, gcnGYBJYB
                If .RecordCount > 0 Then
                    blnMatch = False
                    Do While Not .EOF
                        blnMatch = CheckMatch(Nvl(!�����, "00-99"), lng����, "-")
                        If blnMatch Then blnMatch = CheckMatch(Nvl(!�����, "00-99"), lng����, "-")
                        If blnMatch Then blnMatch = (Nvl(!��ҵ���, str��ҵ���) = str��ҵ���)
                        If blnMatch Then
                            '���öεıȽ���Ҫ��������
                            dbl���� = Split(!���ö�, "-")(0)
                            dbl���� = Split(!���ö�, "-")(1)
                            dbl�������� = Nvl(!��������, 0)
                            dbl�𸶽�� = Nvl(!�𸶽��, 0)
                            blnMatch = ((dblͳ���� + dbl���ͳ���ۼ�) >= dbl���� And dbl���� > dbl���ͳ���ۼ�)
                            '���ܵ�һ��ƥ��񣬶�������ȡ����
                            If bln��һ�� Then
                                '���������𸶣�����֧�����ߣ�������Ϊ�㣬�����������
                                If gCominfo_�Ͻ�.blnYearBase Then
                                    If dbl���ͳ���ۼ� >= dbl�𸶽�� Then
                                        dblʵ������ = 0
                                    Else
                                        dblʵ������ = dbl�𸶽�� - dbl���ͳ���ۼ�
                                        If dblʵ������ < 0 Then dblʵ������ = 0
                                    End If
                                Else
                                    dblʵ������ = dbl�𸶽��
                                End If
                                bln��һ�� = False
                            End If
                        End If
                        
                        '�ҵ���Ӧ�ģ��Ͱ�������м��㣨ֻ�е�һ�ε��������ã�
                        If blnMatch Then
                            '�ó�����˶εĽ��
                            If (dblͳ���� + dbl���ͳ���ۼ�) <= dbl���� Then
                                dbl���� = (dblͳ���� + dbl���ͳ���ۼ�)
                            End If
                            If dbl���ͳ���ۼ� >= dbl���� Then
                                dbl������ = dbl���� - dbl���ͳ���ۼ�
                            Else
                                dbl������ = dbl���� - dbl����
                            End If
                            If dbl������ >= dblʵ������ Then
                                dbl������ = dbl������ - dblʵ������
                                dblʵ������ = 0
                            Else
                                dblʵ������ = dblʵ������ - dbl������
                                dbl������ = 0
                            End If
                            dbl���뱨�� = dbl������ * dbl��������
                            dbl�����ۼ� = dbl�����ۼ� + dbl���뱨��
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
    Next
    rsDisease.Filter = 0
    
    Calcͳ����_�ֵ� = dbl�����ۼ�
End Function

Private Function CheckMatch(ByVal str��Χ As String, ByVal strValue As String, ByVal str�ָ� As String) As Boolean
    'strȱʡ��str��ΧΪ��ʱ��ȱʡֵ
    Dim arrData
    arrData = Split(str��Χ, str�ָ�)
    CheckMatch = (strValue >= arrData(0) And strValue <= arrData(1))
End Function

Private Function ��ȡ���ӷ�ʽ() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���м������ȡ���ӷ�ʽ
    gstrSQL = "Select NVL(���ӷ�ʽ,'�ѻ�') ���ӷ�ʽ From ҽ�ƻ������������ Where ��λ����='" & gCominfo_�Ͻ�.strHospitalCode & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "���컹δ���������أ�����ִ�����س���[��ȡ���ӷ�ʽ]", vbInformation, gstrSysName
            Exit Function
        Else
            gCominfo_�Ͻ�.blnOnLine = (!���ӷ�ʽ <> "�ѻ�")
        End If
    End With
    ��ȡ���ӷ�ʽ = True
End Function

Private Function �����������() As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '������������Ƿ�������һ��,��ͬ���ֹ��¼����Ϊÿ�춼Ҫ���أ�����ֻ���м���н��бȽϣ�
    gstrSQL = " Select ��������,��λ����,�Ƿ�ʹ��IC������,סԺ���ò����Ը�,סԺ����ҩƷ�����Ը�,סԺ�������" & _
              " From ҽ�ƻ������������ Where ��λ����='" & gCominfo_�Ͻ�.strHospitalCode & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "û���ҵ���ҽ�ƻ����Ļ�����Ϣ������������ϵ��[�����������]", vbInformation, gstrSysName
            Exit Function
        Else
            If gCominfo_�Ͻ�.strConnectPass <> Nvl(!��������) Then
                MsgBox "����������󣬿��������ѽ�ֹ��ҽ�ƻ���ʹ�ã�����������ϵ��[�����������]", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '��ϵͳ��������ȫ�ֱ���
        With gCominfo_�Ͻ�
            .strHospitalName = Nvl(rsTemp!��λ����)
            .blnICPassVerify = (InStr(1, "��,NO", UCase(Nvl(rsTemp!�Ƿ�ʹ��IC������, "��"))) = 0)
            .blnDiseaseCash = (InStr(1, "��,NO", UCase(Nvl(rsTemp!סԺ���ò����Ը�, "��"))) = 0)
            .blnPhysicCash = (InStr(1, "��,NO", UCase(Nvl(rsTemp!סԺ����ҩƷ�����Ը�, "��"))) = 0)
            .blnYearBase = (InStr(1, "��,NO", UCase(Nvl(rsTemp!סԺ�������, "��"))) = 0)
        End With
    End With
    ����������� = True
End Function

Private Function ��ȡ��������() As Boolean
    Dim strCurDate As String
    
    If gCominfo_�Ͻ�.blnOnLine = False Then
        ��ȡ�������� = True
        Exit Function
    End If
    
    '�ȼ�鵱ǰ�����Ƿ����һ��ʹ�õ�������ͬ����ͬ������ʹ��
    If mstrFirstStart <> "" Then
        strCurDate = Format(zlDataBase.Currentdate, "yyyy-MM-dd")
        If mstrFirstStart <> strCurDate Then
            MsgBox "�������������򣬲�����������ҽ�����ף�[��ȡ��������]", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����Ƿ�����ͨ����
    If Not gobjCenter.InitConnect("") Then Exit Function
    ��ȡ�������� = True
End Function

Private Sub �ر���������()
    If gCominfo_�Ͻ�.blnOnLine = False Then Exit Sub
    Call gobjCenter.CloseConnector
End Sub

Private Function ����Ƿ��ϴ���ϸ() As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If gCominfo_�Ͻ�.blnOnLine Then
        ����Ƿ��ϴ���ϸ = True
        Exit Function
    End If
    
    '����Ƿ����δ�ϴ�����ϸ
    gstrSQL = " Select 1 " & _
        " From �����ʻ� A,���˷��ü�¼ B,������Ϣ C" & _
        " Where A.����=" & TYPE_�Ͻ� & " And A.����ID=B.����ID And C.����ID=B.����ID And B.��ҳID=C.סԺ����" & _
        " And B.��¼����=3 And Nvl(B.�Ƿ��ϴ�,0)=0 And B.����ʱ�� Between Sysdate-3 and Sysdate-1 And Rownum<2"
    Call OpenRecordset(rsTemp, "����Ƿ����δ�ϴ��ķ�����ϸ")
    ����Ƿ��ϴ���ϸ = (rsTemp.RecordCount = 0)
    If ����Ƿ��ϴ���ϸ = False Then MsgBox "����δ�ϴ��ķ�����ϸ�����������ϴ����س�����ϸ�ϴ���ҽ�����ģ�", vbInformation, gstrSysName
End Function

Private Function ����Ƿ�����() As Boolean
    Dim strCurDate As String, strDownDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If gCominfo_�Ͻ�.blnOnLine Then
        ����Ƿ����� = True
        Exit Function
    End If
    
    '��鵱���Ƿ�����
    strCurDate = Format(zlDataBase.Currentdate, "yyyy-MM-dd")
    gstrSQL = " Select �������� From ҽ�ƻ������������ Where ��λ����='" & gCominfo_�Ͻ�.strHospitalCode & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 1 Then
            strDownDate = Nvl(!��������)
        End If
    End With
    If strCurDate <> strDownDate Then
        MsgBox "���컹δ���������أ�����ִ�����س���[����Ƿ�����]", vbInformation, gstrSysName
        Exit Function
    End If
    
    ����Ƿ����� = True
End Function

Private Function ����_��ʼ() As Boolean
    On Error GoTo ErrHand
    If gCominfo_�Ͻ�.blnOnLine Then
        If Not ��ȡ�������� Then Exit Function
        Call gobjCenter.BeginTrans
    End If
    gcnGYBJYB.BeginTrans
    
    ����_��ʼ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ����_�ύ() As Boolean
    On Error GoTo ErrHand
    gcnGYBJYB.CommitTrans
    If gCominfo_�Ͻ�.blnOnLine Then Call gobjCenter.CommitTrans
    
    ����_�ύ = True
    If gCominfo_�Ͻ�.blnOnLine Then Call �ر���������
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ����_�ع�() As Boolean
    On Error GoTo ErrHand
    gcnGYBJYB.RollbackTrans
    If gCominfo_�Ͻ�.blnOnLine Then Call gobjCenter.RollbackTrans
    
    ����_�ع� = True
    If gCominfo_�Ͻ�.blnOnLine Then �ر���������
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ExecuteSQL(ByVal strSQL As String, Optional ByVal bln�ع� As Boolean = True) As Boolean
    '���bln�ع�=TRUE����˵��������������У��ù����Զ��ع����񣬼�����㺯���ķ��Ӵ���
    If Not gobjCenter.ExecuteSQL(strSQL) Then
        If bln�ع� Then Call ����_�ع�
        Exit Function
    End If
    ExecuteSQL = True
End Function

Public Sub ����ת��_�Ͻ�(CardData As String, Optional blnRead As Boolean = True)
    'IC�������ݹ���
    'shbzh           ��ᱣ�Ϻ�         32      18      string
    'xm              ����               50      10      string
    'dwdm            ��λ����           60      15      string
    'xb              �Ա�               75      2       string
    'csrq            ��������           77      10      string
    'cjgzrq          �μӹ�������       87      10      string
    'jyqkdm          ��ҵ�������       97      1       string
    'yxkh            ��Ч����           99      2       int
    'grjbdm          ���˼������       101     2       string
    'ye              �����ʻ����       120     10      decimal
    'zhjzrq          ����������       151     10      string
    'yydm            ������ҽԺ����   161     4       string
    'pass            ����ic������       168     8       string
    Dim arrData
    
    If Not blnRead Then
        CardData = IC_Data_�Ͻ�.��ᱣ�Ϻ� & "||" & IC_Data_�Ͻ�.���� & "||" & IC_Data_�Ͻ�.��λ���� & "||" & _
            IC_Data_�Ͻ�.�Ա� & "||" & IC_Data_�Ͻ�.�������� & "||" & IC_Data_�Ͻ�.�μӹ������� & "||" & _
            IC_Data_�Ͻ�.��ҵ������� & "||" & IC_Data_�Ͻ�.��Ч���� & "||" & IC_Data_�Ͻ�.���˼������ & "||" & _
            IC_Data_�Ͻ�.�����ʻ���� & "||" & IC_Data_�Ͻ�.���������� & "||" & IC_Data_�Ͻ�.������ҽԺ���� & "||" & _
            IC_Data_�Ͻ�.����IC������
    Else
        arrData = Split(CardData, "||")
        IC_Data_�Ͻ�.��ᱣ�Ϻ� = arrData(ic.shbzh)
        IC_Data_�Ͻ�.���� = arrData(ic.xm)
        IC_Data_�Ͻ�.��λ���� = arrData(ic.dwdm)
        IC_Data_�Ͻ�.�Ա� = arrData(ic.xb)
        IC_Data_�Ͻ�.�������� = arrData(ic.csrq)
        IC_Data_�Ͻ�.�μӹ������� = arrData(ic.cjqzrq)
        IC_Data_�Ͻ�.��ҵ������� = arrData(ic.jyqkdm)
        IC_Data_�Ͻ�.��Ч���� = arrData(ic.yxkh)
        IC_Data_�Ͻ�.���˼������ = arrData(ic.grjbdm)
        IC_Data_�Ͻ�.�����ʻ���� = arrData(ic.ye)
        IC_Data_�Ͻ�.���������� = arrData(ic.zhjzrq)
        IC_Data_�Ͻ�.������ҽԺ���� = arrData(ic.yydm)
        IC_Data_�Ͻ�.����IC������ = arrData(ic.pass)
    End If
End Sub

Private Function Get��ˮ��_�Ͻ�() As String
    Dim lng����ǼǺ� As Long
    Dim rsCheck As New ADODB.Recordset

    '��ȡ��ǰ����ǼǺ�
    gstrSQL = "Select ����ǼǺ�_ID.Nextval from dual"
    With rsCheck
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        lng����ǼǺ� = .Fields(0).Value
    End With
    Get��ˮ��_�Ͻ� = gCominfo_�Ͻ�.strHospitalCode & Format(zlDataBase.Currentdate, "yyyyMMdd") & String(6 - Len(CStr(lng����ǼǺ�)), "0") & lng����ǼǺ�
End Function

Private Function CheckCard(ByVal str��ᱣ�Ϻ� As String) As Boolean
    '���ò��˵Ŀ�״̬����������ᣬ�򷵻ؼ�
    Dim rsTemp As New ADODB.Recordset
    
    '����������Ҫ���м���л�ȡ
    gstrSQL = "Select סԺ����,�ʻ�����,����ԭ��,��Ч����,����סԺ����,����ʱ��,����˵�� " & _
        " From �����ʻ����� Where ��ᱣ�Ϻ�='" & str��ᱣ�Ϻ� & "'"
    If gCominfo_�Ͻ�.blnOnLine Then
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then Exit Function
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        If Nvl(!�ʻ�����, "��") = "��" Then Exit Function
    End With
    CheckCard = True
End Function

Private Function BalanceLack(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��鲡��Ԥ������Ƿ��㹻����������ʾ�ɿ���ݲ�����
    gstrSQL = "Select Ԥ�����,������� From ������� Where ����=1 And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡԤ�����������")
    If rsTemp.RecordCount = 0 Then Exit Function
    If Nvl(rsTemp!Ԥ�����, 0) < Nvl(rsTemp!�������, 0) Then Exit Function
    BalanceLack = True
End Function

Private Sub IC_End(Optional ByVal blnPull As Boolean = False)
    '�ڴ�IC�豸����������Ƿ����������������ڵ�����رն˿�
    Call gobjCenter.IC_PullCard
    If blnPull Then Exit Sub
    
    Call gobjCenter.IC_CloseDevice
End Sub
