Attribute VB_Name = "mdlPublicExpense"
Option Explicit
Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String    '��Ʒ����
Public gstrSQL As String
Public glngSys As Long
Public glngMainModule As Long '�����ߵ�ģ���
Public gstrMainPrivs As String '�����ߵ����Ȩ��
Public gblnOK As Boolean
Public gclsInsure As New clsInsure 'ҽ������
Public gstrDBUser As String '������
Public gcnOracle As ADODB.Connection
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gobjSquare As Object '�����㲿��
Public gobjPlugIn As Object '��ҹ���

'�Һ��ò���
Public gstrRooms As String
Public glngModul As Long
Public gbytState As Byte
Public gstrDocs As String
Public gstrDeptIDs As String
Public gstrPrivs As String
Public gblnBill�Һ� As Boolean
Public gbytRegistMode As Byte
Public gdatRegistTime As Date

Public Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum

Public grsҽ�Ƹ��ʽ As ADODB.Recordset

Private Type TY_Decimal_Precision 'С������
    byt_Bit As Byte 'С��λ��:��ʾ���㵽С�����ڶ���λ��
    strFormt_VB As String   'VB��ʽ��:0.0000;...
    strFormt_ORA As String  'Oracle��ʽ��:999990.00000...
End Type

Private Type ty_SysPara
    bln�����������۷���  As Boolean
    bytƱ�ݷ������ As Byte   'Ʊ�ݷ������:0-����ʵ�ʴ�ӡ����Ʊ��;1-����ϵͳԤ���������;2-�����û��Զ���������
    Money_Decimal As TY_Decimal_Precision  '���ý��С����ʽ
    Price_Decimal As TY_Decimal_Precision  '���õ���С����ʽ
    bln��������ۿ�  As Boolean
    bytҩƷ������ʾ As Byte '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
    byt����ҩƷ��ʾ As Byte '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��

    byt������˷�ʽ As Byte '������˷�ʽ:0-δ��˲��������ʣ�ȱʡΪ0;1-���ʱ�����������ú�ҽ��������ҽ�������ͷ��õ�����
    blnδ��ƽ�ֹ����  As Boolean
    bln����ִ�з��� As Boolean 'ִ��֮�������Զ�����
    blnִ�к���� As Boolean
    blnִ��ǰ�Ƚ��� As Boolean 'һ��ִͨ��ǰ���շѻ�������
    bln�������������� As Boolean '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
    intҽ������ As Integer '�Ƿ��סԺҽ�����˵���Ŀ����������м��:0-�����,1-��鲢����,2-��鲢��ֹ
    dblMaxMoney As Double   '�������
    bytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    bln������֤ As Boolean '����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
    bln����ƥ�䷽ʽ�л� As Boolean '�����ڴ��ڽ���Ĺ������л�����ƥ�䷽ʽ�л�

    blnסԺ�Զ����� As Boolean 'סԺ������ɺ��Ƿ��Զ�����
    bln�����Զ����� As Boolean '���������ɺ��Ƿ��Զ�����
    bln�շѺ��Զ���ҩ As Boolean '
    bln���뷢ҩ As Boolean
    strҽ���������� As String 'ҽ�����������ķ�������
    str���ѷ������� As String '���Ѳ��������ķ�������
    strLike As String
    bytCode As Byte
    bln�շ���� As Boolean '�Ƿ������������
    blnFeeKindCode As Boolean '�������ʱ,��λ�����շ�������
    strMatchMode As String '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
    blnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
    bln�������ۼ���  As Boolean
    blnסԺ���ۼ��� As Boolean
End Type

Public gSysPara As ty_SysPara
Public Enum gEm_BulidIng_SQLType
    EM_Bulid_�ַ� = 0
    EM_Bulid_���� = 1
End Enum
Public Const gstrCompentsName = ""
Public Enum Enum_Inside_Program
    pסԺ���ʲ��� = 1150
    pҽ�����ѹ��� = 1257
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    
End Enum
Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    רҵ����ְ�� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO

'----------------------------------------------------
'����������
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gstrNodeNo As String 'վ����

Public Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�Ȩ�ֱ���
    '����:���˺�
    '����:2014-03-20 16:07:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, varTmp As Variant
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = gobjComlib.GetUnitName
    
    
    With gSysPara
        .bln�����������۷��� = gobjDatabase.GetPara(98, glngSys) = "1"
        With .Money_Decimal '���ý��С��λ��
            .byt_Bit = Val(gobjDatabase.GetPara(9, glngSys, , 2))
            .strFormt_VB = "0." & String(.byt_Bit, "0")
            .strFormt_ORA = "FM" & String(14, "9") & "0." & String(.byt_Bit, "9")
        End With
        With .Price_Decimal  '���õ���С��λ��
            .byt_Bit = Val(gobjDatabase.GetPara(157, glngSys, , 5))
            .strFormt_VB = "0." & String(.byt_Bit, "0")
            .strFormt_ORA = "FM" & String(14, "9") & "0." & String(.byt_Bit, "9")
        End With
        '���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
        strTmp = Trim(gobjDatabase.GetPara("Ʊ�ݷ������", glngSys, 1121, "0||0;0;0;0;0"))
        varTmp = Split(strTmp & "||", "||")
        .bytƱ�ݷ������ = Val(varTmp(0))
        .bln��������ۿ� = Val(gobjDatabase.GetPara(93, glngSys)) <> 0
        .bytҩƷ������ʾ = Val(gobjDatabase.GetPara("ҩƷ������ʾ", , , "2"))
        .byt����ҩƷ��ʾ = gobjDatabase.GetPara("����ҩƷ��ʾ", , , 0)
        .byt������˷�ʽ = Val(gobjDatabase.GetPara(185, glngSys))    '49501
        .blnδ��ƽ�ֹ���� = Val(gobjDatabase.GetPara(215, glngSys)) = 1 '51612
        .bln�������ۼ��� = gobjDatabase.GetPara("�������۲��˼���", glngSys, 1150) = "1"
        .blnסԺ���ۼ��� = gobjDatabase.GetPara("סԺ���۲��˼���", glngSys, 1150) = "1"

        .bln����ִ�з��� = Val(gobjDatabase.GetPara(33, glngSys)) <> 0
        .blnִ�к���� = Val(gobjDatabase.GetPara(81, glngSys)) <> 0
        '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
        .blnִ��ǰ�Ƚ��� = Val(gobjDatabase.GetPara(163, glngSys)) <> 0
        '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
        .bln�������������� = Val(gobjDatabase.GetPara(232, glngSys)) <> 0
        'ҽ��������
        .intҽ������ = Val(gobjDatabase.GetPara(59, glngSys, , 1))
        '���ʷ���������ѽ��
        .dblMaxMoney = Val(gobjDatabase.GetPara(60, glngSys))
    
        '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
        .bytBillOpt = Val(gobjDatabase.GetPara(23, glngSys))
        'һ��ͨ������֤
        .bln������֤ = Val(gobjDatabase.GetPara(28, glngSys)) <> 0
        .bln����ƥ�䷽ʽ�л� = Val(gobjDatabase.GetPara("����ƥ�䷽ʽ�л�", , , "1")) = 1
        '�����Զ�����
        .bln�����Զ����� = Val(gobjDatabase.GetPara(92, glngSys)) <> 0
        'סԺ�Զ�����
        .blnסԺ�Զ����� = Val(gobjDatabase.GetPara(63, glngSys)) <> 0
        '�Զ���ҩ��ҩ
        .bln�շѺ��Զ���ҩ = gobjDatabase.GetPara(45, glngSys) = "1"
        '�����շ��뷢ҩ����
        .bln���뷢ҩ = gobjDatabase.GetPara(15, glngSys) = "1"
        'ҽ����������
        .strҽ���������� = "'" & Replace(gobjDatabase.GetPara(41, glngSys), "|", "','") & "'"
    
        '���ѷ�������
        .str���ѷ������� = "'" & Replace(gobjDatabase.GetPara(42, glngSys), "|", "','") & "'"
            
        '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
        .strMatchMode = gobjDatabase.GetPara(44, glngSys, , "00")
        
        .strLike = IIf(gobjDatabase.GetPara("����ƥ��") = "0", "%", "")
        .bytCode = Val(gobjDatabase.GetPara("���뷽ʽ"))
        '�Ƿ�Ҫ�������������
        .bln�շ���� = Val(gobjDatabase.GetPara(72, glngSys, , 1)) <> 0
        '���������ʱ,���������Ŀʱ,��λ����������
        .blnFeeKindCode = Val(gobjDatabase.GetPara(144, glngSys)) <> 0 And Not .bln�շ����
        'ָ��ҩ��ʱ���ƿ��
        .blnStock = Val(gobjDatabase.GetPara(18, glngSys)) <> 0
    End With
End Sub
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetRoom(str�ű� As String) As String
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
            
    strSQL = "Select ID,Nvl(���﷽ʽ,0) as ���� From �ҺŰ��� Where ����=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublicExpense", str�ű�)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!���� = 0 Then Exit Function '������
    
    '��������
    If rsTmp!���� = 1 Then
        'ָ������
        strSQL = "Select �������� From �ҺŰ������� Where �ű�ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublicExpense", CLng(rsTmp!ID))
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        strSQL = _
            " Select ��������,Sum(NUM) as NUM From (" & _
                " Select ��������,0 as NUM From �ҺŰ������� Where �ű�ID=[1]" & _
                " Union ALL" & _
                " Select ����,Count(����) as NUM From ���˹Һż�¼" & _
                " Where Nvl(ִ��״̬,0)=0 And ��¼����=1 and ��¼״̬=1 and  ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű�=[2]" & _
                " And ���� IN(Select �������� From �ҺŰ������� Where �ű�ID=[1])" & _
                " Group by ����)" & _
            " Group by �������� Order by Num"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublicExpense", CLng(rsTmp!ID), str�ű�)
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        strSQL = "Select �ű�ID,��������,��ǰ���� From �ҺŰ������� Where �ű�ID=" & rsTmp!ID
        Set rsTmp = New ADODB.Recordset
        Call gobjDatabase.OpenRecordset(rsTmp, strSQL, "mdlPublicExpense", adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!��ǰ����), 0, rsTmp!��ǰ����) = 1 Then
                    GetRoom = rsTmp!��������
                    rsTmp!��ǰ���� = 0
                    
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!��ǰ���� = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '������һ��ƽ������
            If GetRoom = "" Then
                rsTmp.MoveFirst
                GetRoom = rsTmp!��������
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!��ǰ���� = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPatiMoney(ByVal bytType As Byte, ByVal lng����ID As Long, ByRef objPatiFee As clsPatiFeeinfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˵���ط�����Ϣ
    '���:bytType-0-����;1-סԺ
    '     lng����ID-����ID
     '����:
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 16:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Set objPatiFee = New clsPatiFeeinfor
    On Error GoTo errHandle
    If bytType = 0 Then
        strSQL = "" & _
        "   Select Nvl(Ԥ�����,0) Ԥ�����,Nvl(�������,0) �������,0 as Ԥ�����,0 as ������ " & _
        "   From ������� " & _
        "   Where ����=1 And ����=1 And ����ID=[1]" & _
        "   "
    Else
        strSQL = "" & _
        "   Select Nvl(Ԥ�����,0) Ԥ�����,Nvl(�������,0) �������,0 as Ԥ����� ,0 as ������" & _
        "   From ������� " & _
        "   Where ����=1 And ����=2 And ����ID=[1]" & _
        "   Union ALL " & _
        "   Select 0 as Ԥ�����,0 as �������,Sum(B.���) as Ԥ����� ,0 as ������" & _
        "   From ������Ϣ A,����ģ����� B" & _
        "   Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1]"
    End If
    strSQL = strSQL & "" & _
    "   Union ALL " & _
    "   Select 0 as Ԥ�����,0 as �������,0 as Ԥ�����,������" & _
    "   From ������Ϣ B " & _
    "   Where ����ID=[1]"
    
    strSQL = "" & _
    "   Select Nvl(Sum(Ԥ�����),0) as Ԥ�����,Nvl(Sum(�������),0) as �������,Nvl(Sum(Ԥ�����),0) as Ԥ����� " & _
    "   From (" & strSQL & ")"
    
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���˵���ط��ý��", lng����ID)
    If rsTemp.EOF Then GetPatiMoney = True: Exit Function
    With objPatiFee
        .Ԥ����� = FormatEx(Val(Nvl(rsTemp!Ԥ�����)), 6)
        .δ����� = FormatEx(Val(Nvl(rsTemp!�������)), 6)
        .Ԥ����� = FormatEx(Val(Nvl(rsTemp!Ԥ�����)), 6)
        .������ = FormatEx(Val(Nvl(rsTemp!������)), 6)
        .ʣ��� = FormatEx(.Ԥ����� + .Ԥ����� - .δ�����, 6)
    End With
    GetPatiMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function FromIDsBulidIngSQL(ByVal bytBulidType As gEm_BulidIng_SQLType, _
    ByVal strValues As String, _
    ByRef varPara As Variant, ByRef strBulitSQL As String, _
    ByVal strAliaName As String, Optional intStartPara As Integer = 1 _
    ) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����IDs����ȡ��ص�SQL,��:select ... From str2List Union ALL Selelct ..
    '���:strValues-ֵ,����ö��ŷ���
    '     strAliaName-����
    '     bytType-0-�ַ���;1-������;
    '     intStartPara-�����Ĳ���
    '����:varPara-���صĲ���ֵ������
    '     strBulitSQL-���صĹ�����SQL��
    '����:�����ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-25 17:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, strTemp As String
    Dim i As Long, j As Long, strSQL As String
    Dim strPara() As Variant, strTable As String, strColumnName As String
    
    On Error GoTo errHandle
    
    strColumnName = " Column_Value "
    If strAliaName <> "" Then strColumnName = strColumnName & " As " & strAliaName
    
    If bytBulidType = EM_Bulid_�ַ� Then
        strTable = "Table(f_str2list([0]))"
    Else
        strTable = "Table(f_Num2list([0]))"
    End If
    
    j = intStartPara
    ReDim Preserve strPara(0 To j - 1) As Variant
    
    
    varData = Split(strValues, ",")
    strTemp = ""
    For i = 0 To UBound(varData)
        If gobjCommFun.ActualLen(strTemp & "," & varData(i)) > 4000 Then
            strSQL = strSQL & " Union ALL  Select " & strColumnName & " From " & Replace(strTable, "[0]", "[" & j & "]")
            ReDim Preserve strPara(0 To j - 1) As Variant
            strPara(j - 1) = Mid(strTemp, 2)
            j = j + 1
            strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSQL = strSQL & " Union ALL  Select " & strColumnName & " From " & Replace(strTable, "[0]", "[" & j & "]")
        ReDim Preserve strPara(0 To j - 1) As Variant
        strPara(j - 1) = Mid(strTemp, 2)
    End If
    
    varPara = strPara
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    strBulitSQL = strSQL
    FromIDsBulidIngSQL = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function


Public Function GetFeeMoneyFromAdviceIDs(ByVal strҽ��IDs As String, _
    ByRef dblOutӦ�ս�� As Double, ByRef dblOutʵ�ս�� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ��IDs����ȡӦ�պ�ʵ�ս��
    '���:strҽ��IDs-ҽ��ID,����ö��ŷ���
    '����:dblOutӦ�ս��-Ӧ�ս��
    '     dblOutʵ�ս��-ʵ�ս��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-25 16:11:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varPara As Variant
    dblOutӦ�ս�� = 0: dblOutʵ�ս�� = 0
    If strҽ��IDs = "" Then Exit Function
    
    '���ܴ���4000
    If gobjCommFun.ActualLen(strҽ��IDs) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_����, strҽ��IDs, varPara, strSQL, "ҽ��ID") = False Then Exit Function
        strSQL = "" & _
        " Select /*+ RULE */ " & _
        "   Nvl(Sum(Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(ʵ�ս��), 0) As ʵ�ս�� " & _
        " From (With ҽ������ As (" & strSQL & ") " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From ������ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id " & _
        "        Union All " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From סԺ���ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id)"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ��ID��ȡ��صķ��ý��", varPara)
    Else
        strSQL = "" & _
        " Select /*+ RULE */ " & _
        "   Nvl(Sum(Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(ʵ�ս��), 0) As ʵ�ս�� " & _
        " From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Num2list([1]))) " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From ������ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id " & _
        "        Union All " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From סԺ���ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id)"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ��ID��ȡ��صķ��ý��", strҽ��IDs)
    End If
    
    On Error GoTo errHandle
    dblOutӦ�ս�� = FormatEx(Val(Nvl(rsTemp!Ӧ�ս��)), 6)
    dblOutʵ�ս�� = FormatEx(Val(Nvl(rsTemp!ʵ�ս��)), 6)
    
    GetFeeMoneyFromAdviceIDs = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function AdviceIsCharged(ByVal strҽ��IDs As String, _
    ByVal strNos As String, ByRef bytOutChargeStatus As Byte, Optional ByRef strOutδ��ҽ��IDs As String, _
    Optional ByRef bytOutBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ҽ���Ƿ��Ѿ��շ�
    '���:strҽ��IDs-ҽ��ID,����ö��ŷ���
    '����:bytOutChargeStatus-�շ�״̬(0-δ�շ�,1-��ȫ�շ�;2-�����շ�)
    '     strOutδ��ҽ��IDs-����δ�շѻ�δ����˵�ҽ��ID
    '     bytOutBillType:���ص�ǰ�ĵ�������(0-�������κε���;1-�շѵ�;2-���ʵ�;3-�շѺͼ��ʶ���)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 09:48:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varPara As Variant
    Dim bytStatus As Byte
    strOutδ��ҽ��IDs = "": bytOutBillType = 0: bytOutChargeStatus = 0
    If strNos = "" And strҽ��IDs = "" Then Exit Function
    
    If strҽ��IDs <> "" Then
        '���ܴ���4000
        If gobjCommFun.ActualLen(strҽ��IDs) > 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_����, strҽ��IDs, varPara, strSQL, "ҽ��ID") = False Then Exit Function
            strSQL = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (" & strSQL & ") " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ��ID��ȡ��صķ��ý��", varPara)
        Else
            strSQL = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Num2list([1]))) " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ��ID��ȡ��صķ��ý��", strҽ��IDs)
        End If
    Else
        '�����ݺŴ���
        '���ܴ���4000
        If gobjCommFun.ActualLen(strNos) > 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSQL, "NO") = False Then Exit Function
            strSQL = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (" & strSQL & ") " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ��ID��ȡ��صķ��ý��", varPara)
        Else
            strSQL = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Str2list([1]))) " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ��ID��ȡ��صķ��ý��", strҽ��IDs)
        End If
        
    End If
    On Error GoTo errHandle
    With rsTemp
        bytStatus = -1
        Do While Not .EOF
             If Val(Nvl(!��¼״̬)) = 0 Then  'δ�շ�
                If Val(Nvl(!ҽ�����)) <> 0 Then
                    strOutδ��ҽ��IDs = strOutδ��ҽ��IDs & "," & Nvl(rsTemp!ҽ�����)
                End If
             End If
             If bytStatus = -1 Then
                If Val(Nvl(!��¼״̬)) = 0 Then
                    bytStatus = IIf(Val(Nvl(!��¼״̬)) = 0, 0, 1)
                End If
             ElseIf bytStatus = 0 And (Val(Nvl(!��¼״̬)) = 1 Or Val(Nvl(!��¼״̬)) = 3) Then
                bytStatus = 2   '�����շ�
             ElseIf bytStatus = 1 And Val(Nvl(!��¼״̬)) = 0 Then
                bytStatus = 2 '�����շ�
             End If
             
             If bytOutBillType = 0 Then
                bytOutBillType = Val(Nvl(!��¼����))
             ElseIf bytOutBillType <> Val(Nvl(!��¼����)) Then
                '��������
                bytOutBillType = 3
             End If
            .MoveNext
        Loop
    End With
    bytOutChargeStatus = bytStatus
    AdviceIsCharged = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function BillExistNotBalance(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��շѵ����Ƿ����δ�շѵ�
    '���:strNOs:ָ���ĵ��ݺ�,�������,�ö��ŷ���
    '����:
    '����:�����д���δ�շѵ�,����true,���򷵻�False
    '����:Ƚ����
    '����:2016-08-25 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    '���ܴ���4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSQL, "NO") = False Then Exit Function
        strSQL = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
                " From ������ü�¼ A,(" & strSQL & ") B" & vbNewLine & _
                " Where Mod(a.��¼����, 10) = 1 And a.NO = b.NO And a.��¼״̬ = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺ����ж��Ƿ����δ�շѵ�", varPara)
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSQL = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
                " From ������ü�¼ A,(Select Column_Value As NO From Table(f_str2list([1]))) B" & vbNewLine & _
                " Where Mod(a.��¼����, 10) = 1 And a.NO = b.NO And a.��¼״̬ = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺ����ж��Ƿ����δ�շѵ�", strNos)
    Else
        strSQL = "Select 1" & vbNewLine & _
                " From ������ü�¼" & vbNewLine & _
                " Where Mod(��¼����, 10) = 1 And NO = [1] And ��¼״̬ = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺ����ж��Ƿ����δ�շѵ�", strNos)
    End If
    
    If rsTemp.EOF Then
        BillExistNotBalance = False '��ȫ���շ�
    Else
        BillExistNotBalance = True '����δ�շ�
    End If
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBillChargeStatus(ByVal strNos As String, ByRef bytOutStatus As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѵ��ݵļƷ�״̬
    '���:strNOs:ָ���ĵ��ݺ�,�������,�ö��ŷ���
    '����:bytOutStatus:0-δ�շ�;1-�����շ�/�˷�;2-ȫ���շ�;3-ȫ���˷�
    '����:��ȡ�ɹ�,����true,���򷵻�False(��δ�ҵ����ݲ���)
    '����:���˺�
    '����:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '���ܴ���4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSQL, "NO") = False Then Exit Function
        strSQL = "Select /*+cardinality(b,10)*/ Sum(a.���� * Nvl(a.����, 1)) As ʣ������," & vbNewLine & _
                "        Sum(Decode(a.��¼����, 1, 1, 0) * Decode(a.��¼״̬, 2, 0, 1) * a.���� * Nvl(a.����, 1)) As ԭʼ����," & vbNewLine & _
                "        Sum(Decode(a.��¼״̬, 0, 1, 0) * a.���� * Nvl(a.����, 1)) As δ������" & vbNewLine & _
                " From ������ü�¼ A,(" & strSQL & ") B " & _
                " Where Mod(a.��¼����, 10) = 1 And a.�۸񸸺� Is Null And a.NO = b.NO"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", varPara)
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSQL = "Select /*+cardinality(b,10)*/ Sum(a.���� * Nvl(a.����, 1)) As ʣ������," & vbNewLine & _
                "        Sum(Decode(a.��¼����, 1, 1, 0) * Decode(a.��¼״̬, 2, 0, 1) * a.���� * Nvl(a.����, 1)) As ԭʼ����," & vbNewLine & _
                "        Sum(Decode(a.��¼״̬, 0, 1, 0) * a.���� * Nvl(a.����, 1)) As δ������" & vbNewLine & _
                " From ������ü�¼ A,(Select Column_Value As NO From Table(f_str2list([1]))) B " & _
                " Where Mod(a.��¼����, 10) = 1 And a.�۸񸸺� Is Null And a.NO = b.NO"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", strNos)
    Else
        strSQL = "Select Sum(���� * Nvl(����, 1)) As ʣ������," & vbNewLine & _
                "        Sum(Decode(��¼����, 1, 1, 0) * Decode(��¼״̬, 2, 0, 1) * ���� * Nvl(����, 1)) As ԭʼ����," & vbNewLine & _
                "        Sum(Decode(��¼״̬, 0, 1, 0) * ���� * Nvl(����, 1)) As δ������" & vbNewLine & _
                " From ������ü�¼" & vbNewLine & _
                " Where Mod(��¼����, 10) = 1 And �۸񸸺� Is Null And NO = [1]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", strNos)
    End If
    
    If Val(Nvl(rsTemp!ԭʼ����)) = 0 Then Exit Function
    If Val(Nvl(rsTemp!ԭʼ����)) = Val(Nvl(rsTemp!δ������)) Then
        bytOutStatus = 0 'δ�շ�
    ElseIf Val(Nvl(rsTemp!ԭʼ����)) = Val(Nvl(rsTemp!ʣ������)) And Val(Nvl(rsTemp!δ������)) = 0 Then
        bytOutStatus = 2 'ȫ���շ�
    ElseIf Val(Nvl(rsTemp!ʣ������)) = 0 Then
        bytOutStatus = 3 'ȫ���˷�
    Else
        bytOutStatus = 1 '�����շ�/�˷�
    End If
    GetBillChargeStatus = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceStatus(ByVal strNos As String, ByRef bytOutStatus As Byte, _
    Optional bln���� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϼ��ʵ��Ƿ��Ѿ�����(ֻ����ʵ�)
    '���:strNOs:ָ���ĵ��ݺ�,�������,�ö��ŷ���
    '     bln����-������ʵ�
    '����:bytOutStatus:0-δ����;1-���ֽ���;2-ȫ������
    '����:��ȡ�ɹ�,����true,���򷵻�False(��δ�ҵ����ݲ���)
    '����:���˺�
    '����:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSQL As String, rsTemp As ADODB.Recordset
    Dim strTable As String
    
    bytOutStatus = 0
    On Error GoTo errHandle
    strTable = IIf(bln����, "������ü�¼", "סԺ���ü�¼")
    '���ܴ���4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSQL, "NO") = False Then Exit Function
        
        strSQL = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When δ���� <> 0 And ���ʽ�� <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As ���ʱ�־" & vbNewLine & _
        "From (Select /*+Cardinality(B,10)*/" & vbNewLine & _
        "        a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.����Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From ������ü�¼ A, (" & strSQL & ") B" & vbNewLine & _
        "       Where a.No = b.No And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.�۸񸸺�, a.���))"

        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", varPara)
        
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSQL = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When δ���� <> 0 And ���ʽ�� <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As ���ʱ�־" & vbNewLine & _
        "From (Select /*+Cardinality(B,10)*/" & vbNewLine & _
        "        a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.����Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From ������ü�¼ A, Table(f_Str2list([1])) B" & vbNewLine & _
        "       Where a.No = b.Column_Value And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.�۸񸸺�, a.���))"
        
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ��ID��ȡ��صķ��ý��", strNos)
    Else
        strSQL = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When δ���� <> 0 And ���ʽ�� <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As ���ʱ�־" & vbNewLine & _
        "From (Select " & vbNewLine & _
        "        a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.����Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From ������ü�¼ A " & vbNewLine & _
        "       Where a.No = [1] And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.�۸񸸺�, a.���))"

        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݵ��ݺŻ�ȡ���ʵ��Ƿ��Ѿ�����", strNos)
    End If
    bytOutStatus = Val(Nvl(rsTemp!���ʱ�־))
    GetBalanceStatus = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceExpenseDetails(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByRef rsOutDetails As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵķ�����ϸ����
    '���:frmMain -����������
    '    lngModule -ģ���
    '    lng����id -����ID
    '����:rsOutDetails-��������(���õ��ţ��շ�����շ����ơ��շ����������ʽ��շѵ��ۡ����㵥λ��ִ�п��ң�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    Dim blnNOMoved As Boolean
    
    Set rsOutDetails = Nothing
    blnNOMoved = gobjDatabase.NOMoved("���˽��ʼ�¼", "", "ID", lng����ID, gstrCompentsName & ":�������Ƿ�ת������ʷ���ռ�")
    
   strSQL = "" & _
    "   Select A.����ʱ��, A.NO,nvl(�۸񸸺�,���) as ���,A.�շ����,A.�շ�ϸĿID," & _
    "           Avg(Nvl(����,1)) *Avg(����) as ����,A.���㵥λ,sum(A.���ʽ��) as ���ʽ��,sum(a.��׼���� ) as �շѵ���, " & _
    "           a.ִ�в���ID" & _
    "   From " & IIf(blnNOMoved, "H", "") & "������ü�¼ A" & _
    "   Where A.����ID=[1]" & _
    "   Group by A.����ʱ��, A.NO,nvl(�۸񸸺�,���),A.�շ����,A.�շ�ϸĿID,A.���㵥λ,a.ִ�в���ID" & _
    "   Union ALL " & _
    "   Select A.����ʱ��, A.NO,nvl(�۸񸸺�,���) as ���,A.�շ����,A.�շ�ϸĿID," & _
    "           Avg(Nvl(����,1)) *Avg(����) as ����,A.���㵥λ,sum(A.���ʽ��) as ���ʽ��,sum(a.��׼���� ) as �շѵ���, " & _
    "           a.ִ�в���ID" & _
    "   From " & IIf(blnNOMoved, "H", "") & "סԺ���ü�¼ A" & _
    "   Where A.����ID=[1] " & _
    "   Group by A.����ʱ��, A.NO,nvl(�۸񸸺�,���),A.�շ����,A.�շ�ϸĿID,A.���㵥λ,a.ִ�в���ID" & _
    "   "
    strSQL = _
    "  Select    A.NO as ���õ���,A.���,A.�շ����,Nvl(E.����,D.����) as �շ�����,A.���� as �շ�����, " & _
    "             a.���ʽ��,a.�շѵ��� ,A.���㵥λ,Nvl(B.����,'δ֪') as ִ�п��� " & _
    " From (" & strSQL & ") A,���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
    " Where A.ִ�в���ID=B.ID(+) And A.�շ�ϸĿID=D.ID" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=3" & _
    " Order by ����ʱ�� Desc,���õ��� Desc,���"
    Set rsOutDetails = gobjDatabase.OpenSQLRecord(strSQL, gstrCompentsName & ":���ݽ���ID��ȡ��������", lng����ID)
    GetBalanceExpenseDetails = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function GetBalanceInfor(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByRef rsOutBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����������
    '���:frmMain -����������
    '    lngModule -ģ���
    '    lng����id -����ID
    '����:rsOutDetails-��������( ���㷽ʽ��������������,ҽ�ƿ����ID,���ѿ�,������ˮ��,����˵��,ˢ�����ţ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    Dim blnNOMoved As Boolean
    
    Set rsOutBalance = Nothing
    blnNOMoved = gobjDatabase.NOMoved("���˽��ʼ�¼", "", "ID", lng����ID, gstrCompentsName & ":�������Ƿ�ת������ʷ���ռ�")
    
   strSQL = "" & _
    "   Select decode(mod(A.��¼����,10),1,'[��Ԥ��]', A.���㷽ʽ) as ���㷽ʽ,  " & _
    "       ��Ԥ�� as ������,A.�������, " & _
    "       A.�����ID,A.���㿨���,decode(nvl(A.���㿨���,0),0,0,1) as ���ѿ�, " & _
    "       A.������ˮ��,A.����˵��,A.���� as ˢ������ " & _
    "   From " & IIf(blnNOMoved, "H", "") & "����Ԥ����¼ A" & _
    "   Where A.����ID=[1]"
    Set rsOutBalance = gobjDatabase.OpenSQLRecord(strSQL, gstrCompentsName & ":���ݽ���ID��ȡ��������", lng����ID)
    GetBalanceInfor = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function
Public Function GetInsidePrivs(ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ڲ�ģ���������е�Ȩ��
    '���:lngProg-�����
    '   blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    '����:
    '����:����Ȩ�޴�
    '����:���˺�
    '����:2014-04-09 11:58:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = gobjComlib.GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function
Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ҳ�����������
    '���:objErr ������� strFunName �ӿڷ�������
    '����:
    '����:���˺�
    '����:2014-04-09 13:27:19
    '˵��:�����������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ������" & vbCrLf & _
            objErr.Number & vbCrLf & _
            objErr.Description, vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Public Function CreatePlugIn(ByVal lngModule As Long, _
    Optional ByVal int���� As Integer) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugIn = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    If gobjPlugIn Is Nothing Then
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    End If
    If gobjPlugIn Is Nothing Then Exit Function
    
    Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngModule, int����)
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "Initialize")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    CreatePlugIn = True
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.����ID = Nvl(rsTmp!����ID, 0)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = Getרҵ����ְ��(UserInfo.ID)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.�û���
End Function

Public Function Getרҵ����ְ��(ByVal lng��ԱID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��רҵ����ְ��
    '����:����ָд��Ա��רҵ����ְ��
    '����:���˺�
    '����:2014-04-09 13:45:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
 
    strSQL = "Select רҵ����ְ�� From ��Ա�� Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Աרҵְ��", lng��ԱID)
    
    Getרҵ����ְ�� = "" & rsTmp!רҵ����ְ��
  
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    '����:������Ա����,����ö��ŷ���
    '����:���˺�
    '����:2014-04-09 13:46:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    If str���� <> "" Then
        strSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա����", str����)
    Else
        strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա����", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ActualMoney(str�ѱ� As String, ByVal lng������ĿID As Long, ByVal curӦ�ս�� As Currency, _
    Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal dbl���� As Double, Optional ByVal dbl�Ӱ�Ӽ��� As Double) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�ϸĿID��������ĿID(ǰ������),Ӧ�ս��,���ѱ����õķֶα������۹������ʵ�ս�
    '     ���ҩƷ���ɱ����ձ����������ʵ�ս��
    '���:str�ѱ�=���˷ѱ�����ǰ���̬�ѱ�,�����ʽΪ"���˷ѱ�,��̬�ѱ�1,��̬�ѱ�2,..."
    '      lng�ⷿID,dbl����,��ҩƷ����Ŀ���ɱ��ۼ��մ���ʱ����Ҫ����
    '      dbl����=�����������ڵ��ۼ�����
    '      dbl�Ӱ�Ӽ���=С������,�����Ӧ�ս���Ѱ��Ӱ�Ӽۼ���ʱ��Ҫ�����ڻ�ԭ������
    '����:
    '����:���أ������۹���ͱ��������ʵ�ս��,����Ƕ�̬�ѱ�,��"str�ѱ�"�������Żݷѱ�(ע�����δ���ۼ���,����ԭ������,Ҳ���ܷ��ص�һ��)
    '����:���˺�
    '����:2014-04-09 13:54:17
    '˵��:
    '   ���ɱ��ۼ��ձ������۵����ּ��㷽��(ʵ����һ��)��
    '       1.���۽�� = �ɱ���� * (1 + ���ձ���)
    '       2.���۽�� = �ɱ��� * (1 + ���ձ���) * ��������
    '   ��صļ��㹫ʽ��
    '      �ɱ��� = ҩƷ�ۼ� * (1 - �����)
    '      �ɱ���� = �ۼ۽�� * (1 - �����) = �ɱ��� * ��������
    '      �п����ʱ:����� = ����� / �����,����:����� = ָ�������
    '      ���ڷ���ҩƷ��Ӧÿ���������ηֱ����ɱ��ۺͳɱ����
    '      ����ʱ�۷�����"ҩƷ�ۼ�=Nvl(���ۼ�,ʵ�ʽ��/ʵ������)"��������ʱ��ҩƷ��治��ʱ��������ۼ��㡣
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, str�ѱ�, lng�շ�ϸĿID, lng������ĿID, curӦ�ս�� / (1 + dbl�Ӱ�Ӽ���), dbl����, lng�ⷿID)
        
    str�ѱ� = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl�Ӱ�Ӽ���), gSysPara.Money_Decimal.strFormt_VB)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
    '���:vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    '����:
    '����:���ظ�ʽ���Ĵ�
    '����:���˺�
    '����:2014-04-09 14:05:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = IIf(blnShowZero, 0, "")
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ģ��Oracle��Decode����
    '����:��������������ֵ
    '����:���˺�
    '����:2014-04-09 14:04:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
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

Public Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
    '���:strText-�����ı�
    '     blnTime=�Ƿ���ʱ�䲿��
    '����:
    '����:�������������ڴ�(yyyy-MM-dd[ HH:mm])
    '����:���˺�
    '����:2014-04-09 14:03:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = gobjDatabase.CurrentDate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function
Public Function NeedName(strList As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ж��Իس����ָ�
    '���:strList:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
    '     2-�ָ��������ȼ����س���(Chr(13)��> - > [] > ()
    '����:
    '����: ��ȡ����
    '����:���˺�
    '����:2014-04-09 14:03:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
        Exit Function
    End If
    '��[]�ָ�
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
        If gobjCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
            Exit Function
        End If
    End If
    '��()�ָ�
    If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
        If gobjCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
            Exit Function
        End If
    End If
    '��-�ָ�
    NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    
End Function
Public Function BillExistBalance(ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�����շѻ��۵��Ƿ�����Ѿ��շѵ�����
    '���:strNO-���ݺ�
    '����:
    '����:���շѷ���true,���򷵻�False
    '����:���˺�
    '����:2014-04-09 14:12:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select ID From ������ü�¼ Where ��¼����=1 And ��¼״̬ IN(1,3) And NO=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "BillExistBalance", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function ExistIOClass(bytBill As Byte) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ����ָ�������������͵�������
    '����:����������ID
    '����:���˺�
    '����:2014-04-09 14:17:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!���ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetBillMax���(ByVal strNO As String, ByVal int��¼���� As Integer, str�Ǽ�ʱ�� As String, int������Դ As Integer) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ݵ�ǰ��������+1
    '���:str�Ǽ�ʱ��=���ҽ��ֻ�����˲���������ʱ����Ҫ�����ɵ��շѻ��۵�(NO��ͬ)��ʱ���������ɵ�һ�¡�
    '     int������Դ:1-���2-סԺ
    '����:
    '����:���ص�ǰ������+1
    '����:���˺�
    '����:2014-04-09 14:18:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIf(int��¼���� = 1 Or (int��¼���� = 2 And int������Դ = 1), "������ü�¼", "סԺ���ü�¼")
    On Error GoTo errHandle
    
    str�Ǽ�ʱ�� = ""
    strSQL = "Select Max(���) as ���,Max(�Ǽ�ʱ��) as ʱ�� From " & strTab & " Where NO=[1] And ��¼����=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strNO, int��¼����)
    If Not rsTmp.EOF Then
        GetBillMax��� = Nvl(rsTmp!���, 0) + 1
        If Not IsNull(rsTmp!ʱ��) Then
            str�Ǽ�ʱ�� = Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        GetBillMax��� = 1
    End If
    Exit Function
    
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��0��ת��Ϊ"NULL"��,������SQL���ʱ��
    '���:blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
    '����:
    '����:����������SQL���
    '����:���˺�
    '����:2014-04-09 14:23:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If IsNull(varValue) Then
        ZVal = "NULL"
    Else
        ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
    End If
End Function


Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function GetPatiDayMoney(lng����ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵��췢���ķ����ܶ�
    '����:���ز��˵ĵ��շ����ܶ�
    '����:���˺�
    '����:2014-04-09 14:59:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
    ByVal cur���ս�� As Currency, ByVal cur���ʽ�� As Currency, ByVal cur������� As Currency, _
    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
    intWarn As Integer, Optional ByVal bln���� As Boolean, _
    Optional blnNotCheck��� As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Բ��˼��ʽ��б�����ʾ
    '���:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
    '     str�շ����=��ǰҪ�������,���ڷ��౨��
    '     str�������=�������,������ʾ
    '     bln����=���ɻ��۷���ʱ�ı��������ƾ���Ƿ��ǿ�Ƽ���Ȩ��ʱ�Ĵ���
    '     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
    '     blnNotCheck���:���������м��(��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
    '����:
    '����:intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
    '     0;û�б���,����
    '     1:������ʾ���û�ѡ�����
    '     2:������ʾ���û�ѡ���ж�
    '     3:������ʾ�����ж�
    '     4:ǿ�Ƽ��ʱ���,����
    '����:���˺�
    '����:2014-04-09 15:00:33
    '˵��:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim bln�ѱ��� As Boolean, byt��־ As Byte
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str���� As String, i As Long
    
    BillingWarn = 0
    
    '�����������:NULL��û������,0�������˵�
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־1 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־2 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־3 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ�����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    'ʾ����"-" �� ",ABC,567,DEF"
    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
    
    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
        If byt��־ = 2 Then
            If str�ѱ���� Like "-*" Then
                byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
            Else
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str�շ����) > 0 Then
                        byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                        'Exit For 'ȡ��˵����סԺ����ģ��
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str������� <> "" Then str������� = """" & str������� & """����"
    str���� = IIf(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
    curʣ���� = curʣ���� + cur������� - cur���ʽ��
    cur���ս�� = cur���ս�� + cur���ʽ��
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If curʣ���� < 0 Then
                        byt��ʽ = 2
                        If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�,�����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    ElseIf curʣ���� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ������ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If curʣ���� < 0 Then
                            byt��ʽ = 2
                            If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�,�����ò��˼�����", frmParent)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If vMsg = vbCancel Then intWarn = 0
                                        BillingWarn = 2
                                    ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                        If vMsg = vbIgnore Then intWarn = 1
                                        BillingWarn = 4
                                    End If
                                Else
                                    If intWarn = 0 Then
                                        BillingWarn = 2
                                    ElseIf intWarn = 1 Then
                                        BillingWarn = 4
                                    End If
                                End If
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
            End If
        End If
    End If
End Function


Public Function zlIsCheckMedicinePayMode(ByVal strҽ�Ƹ������� As String, _
    Optional ByRef blnҽ�� As Boolean, Optional ByRef bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ�Ƹ��ʽ�Ƿ񹫷ѻ�ҽ��
    '���:strҽ�Ƹ�������-ҽ�Ƹ�������
    '����:blnҽ��-true,��ʾҽ��
    '        bln����-true,��ʾ�ǹ���
    '����:��ҽ���򹫷�ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2012-01-17 16:25:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "": blnҽ�� = False: bln���� = False
    If grsҽ�Ƹ��ʽ Is Nothing Then
        strSQL = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    ElseIf grsҽ�Ƹ��ʽ.State <> 1 Then
        strSQL = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    End If
    If strSQL <> "" Then
        Set grsҽ�Ƹ��ʽ = gobjDatabase.OpenSQLRecord(strSQL, "��ȡҽ�Ƹ��ʽ")
    End If
    grsҽ�Ƹ��ʽ.Find "����='" & strҽ�Ƹ������� & "'", , adSearchForward, 1
    If grsҽ�Ƹ��ʽ.EOF Then Exit Function
    blnҽ�� = Val(Nvl(grsҽ�Ƹ��ʽ!�Ƿ�ҽ��)) = 1
    bln���� = Val(Nvl(grsҽ�Ƹ��ʽ!�Ƿ񹫷�)) = 1
    zlIsCheckMedicinePayMode = blnҽ�� Or bln����
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function ShowHelp(ByVal ChmName As String, SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��������
    '����:ChmName:CHM��ʽ�ļ�(Ŀǰ�������:App.ProductName)
    '     SHwnd:���봰�ھ��(��Ϊ��������)
    '     htmName:��ӳ��CHM�е�htm�ļ�����
    '����:���˺�
    '����:2014-05-15 15:49:52
    '-----------------------------------------------------------------------------------------------------------------------------
    ShowHelp = gobjComlib.ShowHelp(ChmName, SHwnd, htmName, Sys)
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ������״̬�����󶥱߽糬��ʱ�����Զ�����Ϊ0
    '���:objForm:Ҫ�ָ��Ĵ���
    '     strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
    '     strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-15 15:53:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
   RestoreWinState = gobjComlib.RestoreWinState(objForm, strProjectName, strUserDef)
End Function

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���洰�弰���и��ֿؼ���״̬
    '���: objForm:Ҫ����Ĵ���
    '      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
    '      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    '����:���˺�
    '����:2014-05-15 15:55:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
   SaveWinState = gobjComlib.SaveWinState(objForm, strProjectName, strUserDef)
End Function
Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo: zlGetComLib = True
    Err = 0: On Error GoTo 0
End Function
 



Public Function zlGetDefaultWindow(ByVal str��� As String, ByVal lngҩ��ID As Long, _
    ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ��ҩ����������
    '���:str���-�շ����
    '     lngҩ��ID-ҩ��ID
    '     lngModule-ģ���
    '����:
    '����:����ȱʡ�ķ�ҩ����
    '����:���˺�
    '����:2014-07-23 18:38:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    Dim str���� As String, lng��ҩ�� As Long
    Dim str�ɴ� As String, lng��ҩ�� As Long
    Dim str�д� As String, lng��ҩ�� As Long
    Select Case str���
        Case "5"
            str���� = gobjDatabase.GetPara(50, glngSys, lngModule)
            lng��ҩ�� = Val(gobjDatabase.GetPara(18, glngSys, lngModule))
            If InStr(str����, ":") > 0 Then '������û�д�ҩ��ID
                 strTmp = str����
            ElseIf lng��ҩ�� > 0 And str���� <> "" Then
                strTmp = lng��ҩ�� & ":" & str����
            End If
        Case "6"
            str�ɴ� = gobjDatabase.GetPara(51, glngSys, lngModule)
            lng��ҩ�� = Val(gobjDatabase.GetPara(19, glngSys, lngModule))
            If InStr(str�ɴ�, ":") > 0 Then
                 strTmp = str�ɴ�
            ElseIf lng��ҩ�� > 0 And str�ɴ� <> "" Then
                 strTmp = lng��ҩ�� & ":" & str�ɴ�
            End If
        Case "7"
            str�д� = gobjDatabase.GetPara(49, glngSys, lngModule)
            lng��ҩ�� = Val(gobjDatabase.GetPara(20, glngSys, lngModule))
            If InStr(str�д�, ":") > 0 Then
                 strTmp = str�д�
            ElseIf lng��ҩ�� > 0 And str�д� <> "" Then
                 strTmp = lng��ҩ�� & ":" & str�д�
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str���
                Case "5"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    zlGetDefaultWindow = strTmp
End Function

Public Function zlGet��ҩ����(ByVal lngModule As Long, ByVal curDate As Date, ByVal lngҩ��ID As Long, ByVal str��� As String, _
    str���� As String, str�ɴ� As String, str�д� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ��Ӧ�ķ�ҩ����
    '���:lngҩ��ID=ִ�в���ID
    '     curDate=��ǰʱ��
    '����:����ҩƷ��Ӧ�ķ�ҩ����
    '����:���˺�
    '����:2014-07-23 18:40:35
    '˵��:��ͬһ������ҩ���ķ�ҩ������ƽ������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lng��ҩ�� As Long, lng��ҩ�� As Long, lng��ҩ�� As Long
    
    On Error GoTo errH
    
    'ָ��ʱ�̶�����(ָ����ָû�ж�Ӧҩ���ϰ�ʱָ��)
    Select Case str���
        Case "5"
            lng��ҩ�� = Val(gobjDatabase.GetPara(18, glngSys, lngModule))

            If str���� <> "" Then
                zlGet��ҩ���� = str����
            ElseIf lng��ҩ�� > 0 Then
                zlGet��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
                str���� = zlGet��ҩ����
            End If
        Case "6"
            lng��ҩ�� = Val(gobjDatabase.GetPara(19, glngSys, lngModule))
            If str�ɴ� <> "" Then
                zlGet��ҩ���� = str�ɴ�
            ElseIf lng��ҩ�� > 0 Then
                zlGet��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
                str�ɴ� = zlGet��ҩ����
            End If
        Case "7"
            lng��ҩ�� = Val(gobjDatabase.GetPara(20, glngSys, lngModule))
            If str�д� <> "" Then
                zlGet��ҩ���� = str�д�
            ElseIf lng��ҩ�� > 0 Then
                zlGet��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
                str�д� = zlGet��ҩ����
            End If
    End Select
    
    
    If zlGet��ҩ���� <> "" Then
        strSQL = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngҩ��ID, zlGet��ҩ����)
        If rsTmp.EOF Then zlGet��ҩ���� = ""
        Exit Function
    End If
    
    '��̬�����ϰ�ķ�ר�Ҵ���
    
    If Val(gobjDatabase.GetPara(19, glngSys, , 0)) = 0 Then
        'æ���Ż���ʽ
        '57332:and Not exists(Select 1 from ������ü�¼ where A.NO=NO  And  ��¼����=decode(A.����,8,1,24,1,0) and ����״̬=1)
        
        strSQL = _
        "   Select ��ҩ����,Sum(Num) as Num  " & _
        "   From (  Select ���� as ��ҩ����,0 as NUM  " & _
        "                From ��ҩ����" & _
        "                Where �ϰ��=1 And Nvl(ר��,0)=0 And ҩ��ID=[2]" & _
        "               Union" & _
        "               Select ��ҩ����,Count(��ҩ����) as Num " & _
        "               From δ��ҩƷ��¼ A" & _
        "               Where �������� Between Trunc(To_Date([1])) And Trunc(To_Date([1])+1)-1/24/60/60 " & _
        "                           And ��ҩ���� IN (Select ���� From ��ҩ���� Where �ϰ��=1 And Nvl(ר��,0)=0 And ҩ��ID=[2])" & _
        "                           And Not exists(Select 1 from ������ü�¼ B where B.NO=A.NO  And  B.��¼����=decode(A.����,8,1,24,1,0) and nvl(B.����״̬,0)=1) " & _
        "               Group by ��ҩ���� " & _
        "           ) " & _
        "   Group by ��ҩ���� " & _
        "   Order by Num"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlOutExse", curDate, lngҩ��ID)
        If Not rsTmp.EOF Then
            zlGet��ҩ���� = Nvl(rsTmp!��ҩ����)
        End If
    Else
        'ƽ�����䷽ʽ
        strSQL = "Select Zl_Get_��ҩ����_Average([1]) as ��ҩ���� From dual"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublicExpense", lngҩ��ID)
        If Not rsTmp.EOF Then
            zlGet��ҩ���� = Nvl(rsTmp!��ҩ����)
        End If
    End If
    
    If zlGet��ҩ���� <> "" Then
        Select Case str���
            Case "5"
                str���� = zlGet��ҩ����
            Case "6"
                str�ɴ� = zlGet��ҩ����
            Case "7"
                str�д� = zlGet��ҩ����
        End Select
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function Getδ��ҩƷ��ҩ����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As String
    '-------------------------------------------------------------------------
    '���ܣ��жϵ�ǰ�����Ƿ������ִͬ�в��ŵ�δ��ҩƷ���������򷵻�δ��ҩƷ�ķ�ҩ����
    '���أ���������ִͬ�в��ŵ�δ��ҩƷ���򷵻�δ��ҩƷ�ķ�ҩ���ڣ����򷵻ؿ�
    '���ƣ�Ƚ����
    '���ڣ�2014-04-09
    '���⣺71902
    '˵����
    '   ͬһ���˲��˲�ͬʱ��ζ��ŵ����շѣ�����ͬһ����ҩ���ڣ����㲡��ȡҩ
    '-------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select ��ҩ����" & vbNewLine & _
            "From δ��ҩƷ��¼" & vbNewLine & _
            "Where ���� = 8 And ��ҩ���� Is Not Null And ����id = [1] And �ⷿid = [2]" & vbNewLine & _
            "Order By ���շ� Desc, �������� Desc"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ����δ��ҩƷ��ҩ����", lng����ID, lngִ�в���ID)
    
    If Not rsTemp.EOF Then
        Getδ��ҩƷ��ҩ���� = Nvl(rsTemp!��ҩ����)
    End If
    rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetDrugWindow(ByVal lngModule As Long, ByVal lngҩ��ID As Long, ByVal str��� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ�ķ�ҩ����,�������ָ����ȱʡ,����ָ��Ϊ׼,����,����ǻ��۵�,���Ե�һҩƷ�еĴ���Ϊ׼,��������������ͬҩƷ�Ĵ���Ϊ׼
    '����:���ط�ҩ����
    '����:���˺�
    '����:2014-07-23 18:49:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��ҩ���� As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim p As Integer, i As Integer, varData As Variant, varTemp As Variant
    Err = 0: On Error GoTo errH:
    str��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
    If str��ҩ���� = "" Then Exit Function
    strSQL = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡȱʡ��ҩ����", lngҩ��ID, str��ҩ����)
    If rsTmp.EOF Then Exit Function
    zlGetDrugWindow = str��ҩ����
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlAddUpdateSwapSQL(ByVal blnԤ�� As Boolean, ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    str���� As String, str������ˮ�� As String, str����˵�� As String, _
    ByRef cllPro As Collection, Optional intУ�Ա�־ As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������ˮ�ź���ˮ˵��
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    '����:cllPro-����SQL��
    '����:���˺�
    '����:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "Zl_�����ӿڸ���_Update("
    '  �����id_In   ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '  ���ѿ�_In     Number,
    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ����ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type
    strSQL = strSQL & "'" & str����˵�� & "',"
    'Ԥ����ɿ�_In Number := 0
    strSQL = strSQL & "" & IIf(blnԤ��, 1, 0) & ","
    '�˷ѱ�־ :1-�˷�;0-����
    strSQL = strSQL & "0,"
    'У�Ա�־
    strSQL = strSQL & "" & IIf(intУ�Ա�־ = 0, "NULL", intУ�Ա�־) & ")"
    zlAddArray cllPro, strSQL
End Function

Public Function zlAddThreeSwapSQLToCollection(ByVal blnԤ���� As Boolean, _
    ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���� As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    ' ����:cllPro-����SQL��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
    varData = Split(strExpend, "||")
    Dim str������Ϣ As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                    str������Ϣ = Mid(str������Ϣ, 3)
                    'Zl_�������㽻��_Insert
                    strSQL = "Zl_�������㽻��_Insert("
                    '�����id_In ����Ԥ����¼.�����id%Type,
                    strSQL = strSQL & "" & lng�����ID & ","
                    '���ѿ�_In   Number,
                    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                    '����_In     ����Ԥ����¼.����%Type,
                    strSQL = strSQL & "'" & str���� & "',"
                    '����ids_In  Varchar2,
                    strSQL = strSQL & "'" & strIDs & "',"
                    '������Ϣ_In Varchar2:������Ŀ|��������||...
                    strSQL = strSQL & "'" & str������Ϣ & "',"
                    'Ԥ����ɿ�_In Number := 0
                    strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
                    zlAddArray cllPro, strSQL
                    str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & "||" & strTemp
            End If
        End If
    Next
    If str������Ϣ <> "" Then
        str������Ϣ = Mid(str������Ϣ, 3)
        'Zl_�������㽻��_Insert
        strSQL = "Zl_�������㽻��_Insert("
        '�����id_In ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '���ѿ�_In   Number,
        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
        '����_In     ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '����ids_In  Varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '������Ϣ_In Varchar2:������Ŀ|��������||...
        strSQL = strSQL & "'" & str������Ϣ & "',"
        'Ԥ����ɿ�_In Number := 0
        strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
        zlAddArray cllPro, strSQL
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckUsedBill(bytKind As Byte, ByVal lng����ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '���ܣ���鵱ǰ����Ա�Ƿ��п���Ʊ������(���û���),�����ؿ��õ�����ID
    '������bytKind=Ʊ��
    '      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
    '      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
    '˵����
    '    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
    '    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
    '    3.���ж�������ʱ,ȱʡ���ٵ�����,��������,"���ʹ�õ�����"ԭ��
    '���أ�
    '      ������Ʊ������ID>0
    '      0=ʧ��
    '      -1:û������(�����δ����)��Ҳû�й���(δ����)
    '      -2:���õĹ���������
    '      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    
    '����Ա��ʣ�������Ʊ�ݼ�
    strSQL = _
        "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
        "From Ʊ�����ü�¼" & vbNewLine & _
        "Where Ʊ�� = [1] And ʹ�÷�ʽ = 1 And ʣ������ > 0 And ������ = [2] And (Nvl(ʹ�����,'LXH')=[3] or  ʹ����� is NULL)" & vbNewLine & _
        "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� Desc, ��ʼ����"
    Set rsSelf = gobjDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, UserInfo.����, IIf(strUseType = "", "LXH", strUseType))
    If lng����ID = 0 Then
        '�����е�һ�μ��,��û�����ñ��ع���
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function 'Ҳû������Ʊ��
        '������Ʊ��,������ԭ�򷵻�
        lngReturn = rsSelf!ID
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        strSQL = "Select ID,ʹ�÷�ʽ,ʣ������,ǰ׺�ı�,��ʼ����,��ֹ���� From Ʊ�����ü�¼ Where Ʊ��=[1]  And (Nvl(ʹ�����,'LXH')=[3] or  ʹ����� is NULL) And ID=[2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, lng����ID, IIf(strUseType = "", "LXH", strUseType))
        '����26352 by ���ջ� 2009-11-20
        If rsTmp.EOF Then CheckUsedBill = -2: Exit Function
        
        If rsTmp!ʹ�÷�ʽ = 2 Then '����,Ҫ�ȿ���û������
            If Not rsSelf.EOF Then
                '�����õģ�����
                lngReturn = rsSelf!ID
            Else
                'û������ȡ����
                If rsTmp!ʣ������ = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                lngReturn = rsTmp!ID
                blnTmp = True
            End If
        Else
            '����Ʊ��
            If rsTmp!ʣ������ > 0 Then
                '��ʣ��
                lngReturn = rsTmp!ID
            Else
                '������ʣ�������
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '��������Ҳû��ʣ��
                lngReturn = rsSelf!ID
            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!��ʼ����) And UCase(strBill) <= UCase(rsTmp!��ֹ����) And Len(strBill) = Len(rsTmp!��ʼ����)) Then
                lngReturn = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '����������,�������������м��
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng����ID As Long) As String
'���ܣ�������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
'˵����1.��ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
'      2.�ſ��ѱ���ĺ���
    Dim rsMain As ADODB.Recordset
    Dim rsDelete As ADODB.Recordset
    Dim strSQL As String, strBill As String
    
    On Error GoTo errH
    
    strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����" & _
        " From Ʊ�����ü�¼ Where ʣ������>0 And ID=[1]"
    Set rsMain = gobjDatabase.OpenSQLRecord(strSQL, "ȡһ��Ʊ�ݺ�", lng����ID)
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!��ǰ����) Then
        strBill = UCase(rsMain!��ʼ����)
    Else
        strBill = UCase(gobjCommFun.IncStr(rsMain!��ǰ����))
    End If
    
     '�����:25448
     '���˺�:ȡ����;����=1 And ԭ��=5 And ���:ԭ���ǿ��ܴ����Ѿ�ʹ���˵�Ʊ��,ʹ���˵�,���ų�
     'Ʊ��: 1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
     '����:1-����(ԭ����1��3��5��������)��2-�ջ�(ԭ����2��4��������)
     'ԭ��:1-��������Ʊ�ݣ�2-�����ջط�Ʊ��3-�ش򷢳�Ʊ�ݣ�4-�ش��ջ�Ʊ�ݣ�5-��������Ʊ��
     
    strSQL = "Select Upper(����) as ���� From Ʊ��ʹ����ϸ" & _
        " Where ����||''>=[1] And ����ID=[2]" & _
        " Order by ����"
        
    Set rsDelete = gobjDatabase.OpenSQLRecord(strSQL, "ȡһ��Ʊ�ݺ�", strBill, lng����ID)
    Do While True
        '��鷶Χ
        If Left(strBill, Len("" & rsMain!ǰ׺�ı�)) <> UCase("" & rsMain!ǰ׺�ı�) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!��ʼ����) And strBill <= UCase(rsMain!��ֹ����)) Then
            Exit Function
        End If
                
        '�ſ������
        rsDelete.Filter = "����='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = gobjCommFun.IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         Set gobjSquare.objSquareCard = Nothing
     End If
     Set gobjSquare = Nothing
     If Err <> 0 Then Err.Clear: Err = 0
End Sub

Public Function zlGetFeeFields(Optional strTableName As String = "������ü�¼", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ������ֵ
    '��Σ�strTableName:��:������ü�¼;סԺ���ü�¼;....
    '      blnReadDatabase-�����ݿ��ж�ȡ
    '���Σ�
    '���أ��ֶμ�
    '���ƣ����˺�
    '���ڣ�2010-03-10 10:41:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo Errhand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "������ü�¼"
        zlGetFeeFields = "" & _
        "Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, " & _
        "����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, " & _
        "�Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, " & _
        "����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, " & _
        "���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
        Exit Function
    Case "סԺ���ü�¼"
        zlGetFeeFields = "" & _
         " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, " & _
         " �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, " & _
         " ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, " & _
         " ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, " & _
         " ����id , ���ʽ��, ���մ���ID, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
         Exit Function
    Case "���˽��ʼ�¼"
        zlGetFeeFields = "Id, No, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������, ��ע"
        Exit Function
    Case "����Ԥ����¼"
        zlGetFeeFields = "" & _
        " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, " & _
        " ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�,Ԥ�����,�����ID,���㿨���,����,������ˮ��,����˵��,������λ,�������,У�Ա�־"
        Exit Function
    Case "��Ա��"
        zlGetFeeFields = "" & _
        "Id, ���, ����, ����, ����֤��, ��������, �Ա�, ����, ��������, �칫�ҵ绰, �����ʼ�, ִҵ���, ִҵ��Χ, " & _
        "����ְ��, רҵ����ְ��, Ƹ�μ���ְ��, ѧ��, ��ѧרҵ, ��ѧʱ��, ��ѧ����, ������ѵ, ���п���, ���˼��, ����ʱ��, " & _
        "����ʱ��, ����ԭ��, ����, վ��"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ����Ϣ", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!Column_Name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
Errhand:
    zlGetFeeFields = "*"
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function ReadRegistPrice(ByVal lng��ĿID As Long, ByVal bln���� As Boolean, ByVal bln���￨ As Boolean, _
    Optional str�ѱ� As String, Optional rsItems As ADODB.Recordset, Optional rsIncomes As ADODB.Recordset) As Long
'���ܣ���ȡָ���Һ���Ŀ��Ӧ�ķ�����Ϣ����¼����
'������lng��ĿID=��ʾ�Ƿ��ȡ�Һŷ���(Ҫ���ĹҺ���ĿID)
'      bln����=��ʾ�Ƿ��ȡ����������(���ܽ���ȡ������)
'      bln���￨=��ʾ�Ƿ��ȡ���￨����(��Һŷѻ�����һ����ȡ)
'      str�ѱ�=�Һŷѱ�
'      rsItems(Out)=�����Һ���Ŀ��������Ŀ,������New��ʽ����
'      rsInComes(Out)=����������Ŀ���������,������New��ʽ����
'���أ���ȡ����Ŀ����,ͬʱrsItems,rsInCome=Nothing
'˵������������Ϊ1,����趨���δ���,��Ϊ�̶�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngԭ��ID As Long
    
    Set rsItems = Nothing
    Set rsIncomes = Nothing
    
    '��ȡ�Һ���Ŀ��������Ŀ�ķ���
    If lng��ĿID <> 0 Then
        strSQL = _
            "Select 1 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=[1]" & _
            " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))"
        strSQL = strSQL & " Union ALL " & _
            "Select 2 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " D.�������� as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=[1]" & _
            " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))"
    End If
    
    '��ȡ���������Ѷ�Ӧ�ķ���
    If bln���� Then
        strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
            "Select 3 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,A.ִ�п��� as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.�շ�ϸĿID And D.�ض���Ŀ='������'" & _
            " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))"
    End If
    
    If strSQL = "" Then Exit Function
    
    '������,����,����˳������
    strSQL = "Select * From (" & strSQL & ") Order by ����,��Ŀ����,�������"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng��ĿID)
    If Not rsTmp.EOF Then
        '�ȴ�����¼��
        Set rsItems = New ADODB.Recordset
        rsItems.Fields.Append "����", adSmallInt '1-����,2-����,3-������,4-���￨��
        rsItems.Fields.Append "ִ�п���ID", adBigInt
        rsItems.Fields.Append "���", adVarChar, 1
        rsItems.Fields.Append "��ĿID", adBigInt
        rsItems.Fields.Append "��Ŀ����", adVarChar, 80
        rsItems.Fields.Append "���㵥λ", adVarChar, 20, adFldIsNullable
        rsItems.Fields.Append "����", adSingle
        rsItems.Fields.Append "������Ŀ��", adSmallInt, , adFldIsNullable
        rsItems.Fields.Append "���մ���ID", adBigInt, , adFldIsNullable
        rsItems.Fields.Append "���ձ���", adVarChar, 80
        
        rsItems.CursorLocation = adUseClient
        rsItems.LockType = adLockOptimistic
        rsItems.CursorType = adOpenStatic
        rsItems.Open
        
        Set rsIncomes = New ADODB.Recordset
        rsIncomes.Fields.Append "��ĿID", adBigInt
        rsIncomes.Fields.Append "������ĿID", adBigInt
        rsIncomes.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
        rsIncomes.Fields.Append "����", adSingle
        rsIncomes.Fields.Append "Ӧ��", adCurrency
        rsIncomes.Fields.Append "ʵ��", adCurrency
        rsIncomes.Fields.Append "ͳ����", adCurrency, , adFldIsNullable
        rsIncomes.CursorLocation = adUseClient
        rsIncomes.LockType = adLockOptimistic
        rsIncomes.CursorType = adOpenStatic
        rsIncomes.Open
        
        For i = 1 To rsTmp.RecordCount
            '�Һ���Ŀ����
            If lngԭ��ID <> rsTmp!��ĿID Then
                rsItems.AddNew
                rsItems!���� = rsTmp!����
                 '0-����ȷ����,1-�������ڿ���,2-�������ڲ���,3-���������ڿ���,4-ָ������
                If rsTmp!ִ�п������� = -1 Then
                    rsItems!ִ�п���ID = 0      '0-��ʾ�Һſ���
                Else
                    rsItems!ִ�п���ID = Get�Һ�ִ�п���ID(rsTmp!��ĿID, rsTmp!ִ�п�������)
                End If
                
                rsItems!��� = rsTmp!���
                rsItems!��ĿID = rsTmp!��ĿID
                rsItems!��Ŀ���� = rsTmp!��Ŀ����
                rsItems!���㵥λ = rsTmp!���㵥λ
                rsItems!���� = Format(Nvl(rsTmp!����, 0), "0.000")
                rsItems.Update
            End If
            lngԭ��ID = rsTmp!��ĿID
            
            '������Ŀ����
            rsIncomes.AddNew
            rsIncomes!��ĿID = rsTmp!��ĿID
            rsIncomes!������ĿID = rsTmp!������ĿID
            rsIncomes!�վݷ�Ŀ = rsTmp!�վݷ�Ŀ
            rsIncomes!���� = Format(Nvl(rsTmp!����, 0), "0.00")
            rsIncomes!Ӧ�� = Format(rsItems!���� * rsIncomes!����, "0.00")
            If Nvl(rsTmp!���ηѱ�, 0) = 1 Then
                rsIncomes!ʵ�� = rsIncomes!Ӧ��
            Else
                rsIncomes!ʵ�� = Format(GetActualMoney(str�ѱ�, rsTmp!������ĿID, rsIncomes!Ӧ��, rsTmp!��ĿID), "0.00")
            End If
            rsIncomes.Update
            rsTmp.MoveNext
        Next
        ReadRegistPrice = rsItems.RecordCount
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
    Set rsItems = Nothing
    Set rsIncomes = Nothing
End Function

Public Function ReadExRegistPrice(ByRef rsExpenses As ADODB.Recordset, ByRef blnAppointPrice As Boolean, _
            Optional ByVal lng����ID As Long, Optional ByVal int���� As Integer, Optional ByVal str�ű� As String, _
            Optional ByVal str���� As String, Optional ByVal str�Ա� As String, Optional ByVal str���� As String, _
            Optional ByVal str����֤�� As String, Optional ByVal str�ѱ� As String, Optional ByVal strҽ�Ƹ��ʽ As String) As Boolean
    '==================================================================================================================================
    '���ܣ����ݹҺż�������Ϣ��ȡ���ӷ�
    '��Σ��Һż����˵Ļ�����Ϣ
    '���Σ�rsExpenses(Out)=���ӷѵ��շ���Ŀ
    '���أ�
    '˵����
    '==================================================================================================================================
    Dim strFee As String, str������ As String, strTmp As String, strSQL As String
    Dim varFees As Variant, strTmp1() As String, strTmp2() As String
    Dim strDateCondition As String, strWherePriceGrade As String
    Dim i As Long, j As Long
    Dim rsFee As ADODB.Recordset
    
    On Error GoTo Errhand
    Set rsExpenses = Nothing
    '����: �շ�ϸĿID|����|����|Ӧ��|ʵ��,....����ö��ŷָ�,����NULLʱ��������,ֻ�����շ�ϸĿIDʱ���շѼ�ĿΪ׼��
    '��֧�ַ�����ͬ���շ�ϸĿID�����ܴ������κϲ�
    strFee = "Select zl_Fun_CustomRegExpenses([1],[2],[3],[4],[5],[6],[7],[8],[9]) As ���ӷ� From Dual"
    Set rsFee = gobjDatabase.OpenSQLRecord(strFee, "zl_Fun_CustomRegExpenses", _
                lng����ID, int����, str�ű�, str����, str�Ա�, str����, str����֤��, str�ѱ�, strҽ�Ƹ��ʽ)
    If rsFee.EOF Then ReadExRegistPrice = True: Exit Function
    str������ = Nvl(rsFee!���ӷ�)
    If str������ = "" Then ReadExRegistPrice = True: Exit Function
    blnAppointPrice = InStr(1, str������, "|") > 0
        
    If blnAppointPrice Then
        strTmp1() = Split(str������, ",")
        str������ = ""
        ReDim varFees(UBound(strTmp1))
        For i = 0 To UBound(strTmp1)
            strTmp2() = Split(strTmp1(i) & "||||", "|")
            varFees(i) = strTmp2()
            str������ = str������ & "," & strTmp2(0)
        Next
        If str������ <> "" Then str������ = Mid(str������, 2)
'        str������ = Replace(str������, "|", ":")
        
        strSQL = "" & _
            "Select /* +cardinality(D,10) */ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_str2list([1])) D " & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.Column_Value " & _
            " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))"
            
        'ָ���۸�Ӧ�ô��ڴ�����
'        strSQL = strSQL & " Union ALL " & _
'            "Select /* +cardinality(E,10) */ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
'            " D.C2 * D.������� as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,D.C3 as ����,-1 as ִ�п�������, " & _
'            " D.C4 as Ӧ��, D.C5 as ʵ��" & _
'            " From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D, Table(f_str2list([1])) E" & _
'            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=E.C1 " & _
'            " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
'            strWherePriceGrade
    Else
        '��ָ�����������ǰ�Ĵ�����ʽ����
        strSQL = "" & _
            "Select /* +cardinality(D,10) */ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_str2list([1])) D " & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.Column_Value " & _
            " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))"

        strSQL = strSQL & " Union ALL " & _
            "Select /* +cardinality(E,10) */ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " D.�������� as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D, Table(f_str2list([1])) E" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=E.Column_Value " & _
            " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))"
    End If

    Set rsFee = gobjDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str������)
    
    If rsFee.EOF Then Exit Function
    
    '�ȴ�����¼��
    Set rsExpenses = New ADODB.Recordset
    With rsExpenses
        .Fields.Append "����", adSmallInt '1-����,2-����,3-������,4-���￨��,5-���ӷ�
        .Fields.Append "ִ�п���ID", adBigInt
        .Fields.Append "���", adVarChar, 1
        .Fields.Append "��ĿID", adBigInt
        .Fields.Append "��Ŀ����", adVarChar, 80
        .Fields.Append "���㵥λ", adVarChar, 20, adFldIsNullable
        .Fields.Append "������ĿID", adBigInt
        .Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
        .Fields.Append "����", adSingle
        .Fields.Append "����", adSingle
        .Fields.Append "Ӧ��", adCurrency
        .Fields.Append "ʵ��", adCurrency
        .Fields.Append "������Ŀ��", adSmallInt, , adFldIsNullable
        .Fields.Append "���մ���ID", adBigInt, , adFldIsNullable
        .Fields.Append "���ձ���", adVarChar, 80
        .Fields.Append "ͳ����", adCurrency, , adFldIsNullable
    
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    
        Do While Not rsFee.EOF
            .AddNew
            
            !���� = rsFee!����
            !ִ�п���ID = 0
            !��� = rsFee!���
            !��ĿID = rsFee!��ĿID
            !��Ŀ���� = rsFee!��Ŀ����
            !���㵥λ = rsFee!���㵥λ
            !������ĿID = rsFee!������ĿID
            !�վݷ�Ŀ = rsFee!�վݷ�Ŀ
            If blnAppointPrice Then
                For i = 0 To UBound(varFees)
                    If Val(varFees(i)(0)) = !��ĿID Then
                        !���� = Format(varFees(i)(1), "0.000")
                        !���� = Format(varFees(i)(2), "0.00")
                        !Ӧ�� = Format(Val(varFees(i)(3)), "0.00")
                        !ʵ�� = Format(Val(varFees(i)(4)), "0.00")
                        Exit For
                    End If
                Next
            Else
                !���� = Format(Nvl(rsFee!����, 0), "0.000")
                !���� = Format(Nvl(rsFee!����, 0), "0.00")
                !Ӧ�� = Format(rsFee!���� * rsFee!����, "0.00")
                If Nvl(rsFee!���ηѱ�, 0) = 1 Then
                    !ʵ�� = !Ӧ��
                Else
                    !ʵ�� = Format(GetActualMoney(str�ѱ�, !������ĿID, !Ӧ��, !��ĿID), "0.00")
                End If
            End If
            
            .Update
            rsFee.MoveNext
        Loop
    End With
    ReadExRegistPrice = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set rsExpenses = Nothing
End Function

Public Function GetActualMoney(ByVal str�ѱ� As String, ByVal lng����ID As Long, ByVal curӦ�� As Currency, ByVal lng�շ�ϸĿID As Long) As Currency
'���ܣ�����ָ���ķѱ��������Ŀ���շ���Ŀ,����ָ������ʵ���տ���
'������
'   str�ѱ�   ���ѱ�
'   lng����ID  ��������ĿID
'   curӦ�գ�Ӧ�ս��ֵ
'���أ�ʵ��Ӧ�յĽ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select ʵ�ձ���" & vbNewLine & _
            "From �ѱ���ϸ" & vbNewLine & _
            "Where �ѱ� = [1] And �շ�ϸĿid = [3] And Abs([4]) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select ʵ�ձ���" & vbNewLine & _
            "From �ѱ���ϸ A" & vbNewLine & _
            "Where �ѱ� = [1] And ������Ŀid = [2] And Abs([4]) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And Not Exists" & vbNewLine & _
            " (Select 1 From �ѱ���ϸ C Where C.�ѱ� = A.�ѱ� And C.�շ�ϸĿid = [3])"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, str�ѱ�, lng����ID, lng�շ�ϸĿID, curӦ��)
    If rsTmp.EOF Then
        GetActualMoney = curӦ��
    Else
        GetActualMoney = curӦ�� * rsTmp!ʵ�ձ��� / 100
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get�Һ�ִ�п���ID(ByVal lng��ĿID As Long, ByVal intִ�п������� As Integer) As Long
'���ܣ���ȡ�ҺŸ�����Ŀ(������,���￨��)���շ���Ŀ��ִ�п���
'������
'���أ����������,��ʾ�Һſ���(ҽ�����ڿ���)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Get�Һ�ִ�п���ID = UserInfo.����ID
    
    Select Case intִ�п�������
        Case 0 '0-����ȷ����
        Case 1 '1-�������ڿ���
            Get�Һ�ִ�п���ID = 0
        Case 2 '2-�������ڲ���
            Get�Һ�ִ�п���ID = 0
        Case 3 '3-����Ա����
        Case 4 '4-ָ������
            strSQL = "Select ִ�п���ID From �շ�ִ�п��� Where �շ�ϸĿID=[1] And Nvl(������Դ,1)=1 "
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng��ĿID)
            
            If Not rsTmp.EOF Then Get�Һ�ִ�п���ID = rsTmp!ִ�п���ID
        Case 5 'Ժ��ִ��(Ԥ��,������δ��)
        Case 6 '�����˿���
    End Select
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str�������� As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�������,���ò�ͬ�������͵���ʾ��ɫ
    '���:objPatiControl-���˿ؼ�(�ı���,��ǩ)
    '    str��������-��������
    '    lngDefaultColor-ȱʡ���˵���ʾ��ɫ
    '����:True-������ɫ�ɹ���False-ʧ��
    '����:���ϴ�
    '����:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str�������� <> "" Then
        lngColor = gobjDatabase.GetPatiColor(str��������)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function GetMoneyInfoRegist(lng����ID As Long, Optional dblModiMoney As Double, _
    Optional blnInsure As Boolean, _
    Optional int���� As Integer = -1, _
    Optional bln������ͳ�� As Boolean = False, _
    Optional bytModiMoneyType As Byte = 0) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵�ʣ���
    '���:blnInsure=�Ƿ��ſ�ҽ�����˵�Ԥ�����
    '       curModiMoney=�޸�ʱ,ԭ���ݵĵ�ǰ���˵ķ��úϼ�
    '       int����:����(0-�����סԺ����;1-����;2-סԺ),-1��ʾ����
    '       bytModiMoneyType-�޸ķ��õ����(�ڰ����ͳ��ʱ��Ч)
    '       blnFamilyMoney-�Ƿ��ȡ�������
    '����:
    '����:����ʣ���
    '����:���˺�
    '����:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, blnҽ�� As Boolean, lng��ҳId As Long
    Dim strSQL As String
    On Error GoTo errH
    If blnInsure Then
        strSQL = "Select A.����,A.��ҳID From ������ҳ A,������Ϣ B" & _
                " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
                " And B.����ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
        If Not rsTmp.EOF Then
            blnҽ�� = Not IsNull(rsTmp!����)
            lng��ҳId = rsTmp!��ҳID
        End If
    End If
    strSQL = "Select " & IIf(bln������ͳ��, "����,", "") & _
            "       Nvl(�������,0) As �������,Nvl(Ԥ�����,0) As Ԥ�����" & _
            " From �������" & _
            " Where ����=1 And ����ID=[1] " & IIf(int���� = -1, "", " And ����=[4]")
  
    If dblModiMoney <> 0 Then   '����Ҫ��Union��ʽ,���ֱ��ȥ��,�ڲ�������޼�¼ʱ,���᷵�ؼ�¼
        strSQL = strSQL & " Union All " & _
                " Select " & IIf(bln������ͳ��, "[4] as ����,", "") & _
                "       -1*[3] as �������,0 as Ԥ����� From Dual"
    End If
    
    '���Ϊҽ��סԺ���ˣ����ڷ���������ſ�Ԥ���еķ���(���ڱ���)
    If blnInsure And blnҽ�� Then
        strSQL = strSQL & " Union All " & _
        " Select  " & IIf(bln������ͳ��, "Decode(��ҳID,NULL,1,0,1,2) as ����,", "") & _
        "       -1*Nvl(���,0) as �������,0 as Ԥ�����" & _
        " From ����ģ�����" & _
        " Where ����ID=[1] And ��ҳID=[2] "
    End If
    strSQL = "Select " & IIf(bln������ͳ��, "����,", "") & _
            "       nvl(Sum(�������),0) as �������,nvl(Sum(Ԥ�����),0) as Ԥ����� " & _
            " From (" & strSQL & ")" & vbCrLf & _
            IIf(bln������ͳ��, " Group by ���� ", _
                IIf(bln������ͳ��, " Group by ����", ""))
    
    Set GetMoneyInfoRegist = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳId, dblModiMoney, int����)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub

Public Function zlFormatNum(ByVal strMoney As String) As String
    strMoney = Replace(strMoney, Chr(44), "")
    zlFormatNum = strMoney
End Function

Public Function CheckChargeItemByPlugIn(objPlugIn As Object, _
    lngSys As Long, ByVal lngModule As Long, _
    ByVal intType As Integer, ByVal intMode As Integer, _
    ByRef rsDetail As ADODB.Recordset, Optional strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ҳ������շ���Ŀ��Ч�Խ��м��
    '���:lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '     intType:0-����;1-סԺ
    '     intMode:0-¼����ϸʱ�ĳ�����;1-���浥��ǰ�Ļ��ܼ��
    '     rsDetail-����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������
    '     strExpend-���Ժ���չ��������
    '����:strExpend-���Ժ���չ��������
    '����:���ݺϷ�����true,���򷵻�False
    '����:Ƚ����
    '����:2017-04-19 10:09:26
    '�����:105189
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '1.û����Ҳ���ʱ����Ϊ���ͨ��
    '2.��Ҳ�������CheckChargeItem�ӿڣ�Ҳ��Ϊ���ͨ��
    If objPlugIn Is Nothing Then CheckChargeItemByPlugIn = True: Exit Function
    
    On Error Resume Next
    If objPlugIn.CheckChargeItem(lngSys, lngModule, intType, intMode, rsDetail, strExpend) = False Then
        'ע�⣬�ӿڲ�����ʱҲ�����
        If Err <> 0 Then
            If Err.Number = 438 Then '�ӿڲ����ڣ���Ϊ���ͨ��
                CheckChargeItemByPlugIn = True
                Exit Function
            End If
            Call zlPlugInErrH(Err, "CheckChargeItem")
        End If
        Exit Function
    End If
    CheckChargeItemByPlugIn = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPatiID(ByVal lngModel As Long, ByVal frmParent As Object, _
                                        ByVal strIDnumber As String, ByVal objControl As Object, _
                                        Optional ByVal strPatiName As String = "", _
                                        Optional ByVal strPatiSex As String = "", _
                                        Optional ByRef blnCancel As Boolean = False) As Long
    '����:���ݲ�������֤��(����,�Ա�)��ȡ����id,����id�п����Ƕ��
    '���:lngModel-ģ���
    '       frmParent-��ʾ�ĸ�����
    '       vRect-�ؼ�����Ļ�е�λ��
    '       objControl-��������֤��ˢ����֤�Ŀؼ�
    '       strIDnumber-����֤��
    '       strPatiName-��������
    '       strPatiSex-�����Ա�
    Dim strSQL As String, strPatiIDs As String
    Dim rsTmp  As ADODB.Recordset
    Dim vRect As RECT
    On Error GoTo Errhand
    strSQL = "Select zl_Custom_PatiIDs_Get([1],[2],[3],[4]) As ����IDs From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lngModel, strIDnumber, strPatiName, strPatiSex)
    If rsTmp.EOF Then
        GetPatiID = 0: Exit Function
    End If
    strPatiIDs = Nvl(rsTmp!����IDs)
    If InStr(strPatiIDs, ",") > 0 Then
        strSQL = _
                    " Select /*+cardinality(B,10)*/ distinct A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.����֤��,A.��ͥ��ַ,A.������λ " & _
                    " From ������Ϣ A, Table(f_Str2List([1])) B " & _
                    " Where a.����ID=b.Column_Value" & _
                    " Order by ����,�Ա�,����"
        strSQL = "Select  *  From (" & strSQL & ") Where Rownum < 101"
        
        vRect = zlControl.GetControlRect(objControl.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, objControl.Height, blnCancel, False, True, strPatiIDs)
        If Not rsTmp Is Nothing Then
            If Val(rsTmp!ID) <> 0 Then GetPatiID = Val(rsTmp!ID)
        End If
    Else
        GetPatiID = strPatiIDs
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
