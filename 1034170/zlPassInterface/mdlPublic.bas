Attribute VB_Name = "mdlPublic"
Option Explicit

'����---------------------
Public Const G_STR_PASS As String = "������ҩ����"
Public Const G_STR_MATCH As String = "abcdefghigklmnopkrstuvwxyzABCDEFGHIGKLMNOPKRSTUVWXYZ0123456789"" </>_="
Public Const G_INT_MODEL_0 As Integer = 0
Public Const G_INT_MODEL_1 As Integer = 1
Public Const G_STR_SPLIT As String = "&&"
Public Const SW_SHOWNORMAL = 1

Public Const G_STR_PARA_MK4 As String = "PASS_1_MK4"
Public Const G_STR_PARA_DTBS As String = "PASS_2_DTBS"
Public Const G_STR_PARA_HZYY As String = "PASS_5_HZYY"
Public Const G_STR_PARA_ZL As String = "PASS_6_ZL"

'ȫ�ֱ���-------------------------------
Public gfrmMain As Object                   '������
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��

Public gstrSysName As String                'ϵͳ����
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngSys As Long
Public gbytUseType As Byte                  '0-ҽ���´�
                                            '1-�ٴ�·����Ŀ��ҽ������
                                            '2-�ٴ�·�����·������Ŀ��ҽ����������ѡ����
                                            '3-ҽ��˳�����(������ʾ��ֹͣ��ҽ������Ϊ�ƶ�ʱ������Щҽ�������Ҫһ�����)
Public glngObject As Long                   '��Ǵ�������
Public gobjPlugIn   As Object


'�ַ�����UTF-8����
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

'------------------------------------------------------------------
'������ҩ������ò���
'------------------------------------------------------------------

Public gbytPass As Byte             'ZLHIS��ʹ��PASS�ӿ�����,0-δʹ��,1-����,2-��ͨ,3-̫Ԫͨ,4-ҩ��ʿ
Public gbytBlackLamp As Byte        '�Ƿ��������ҩƷ
Public gbytSuperVolume As Byte      '�Ƿ��ֹ������ҩƷ
Public gbytOutBlackLamp As Byte     '�Ƿ�����Ժ��ִ�еĽ���ҩƷҽ��
Public gbytReason As Byte           '����ҩƷҪ����дԭ��
Public gobjPass As Object           '3-̫Ԫͨ�ӿڶ���
Public gbytOpenLog As Byte          '������ͨ�ӿڵ�����־ 0-�����ã�1-����
Public gbytSysSet As Byte           '��������ʹ��ϵͳ���� 1-��ʾ��0-����
Public gstrVersion As String        '��ʶ�ӿڰ汾��
Public gstrIP As String             '������IP
Public gstrPort As String           '�������˿ں�
Public gstrPortPlus As String           '�������˿ں�
Public gstrHOSCODE As String        'ҽԺ����
Public gbytType As Byte          '��������:0-����;1-�ǹ���
'---------------

Public gint�����Ǽ���Ч���� As Integer
Public gblnInitOK As Boolean         '��ʼ���ɹ�

'��¼�û��ṹ
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
    רҵ��������  As String
    ��ҩ���� As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Enum DataEnum
    responseText = 1
    responseBody = 2
End Enum
Public UserInfo As TYPE_USER_INFO

Public gobjCOL As clsVSCOL           '��ǰҽ����ӳ��
Public gobjAdvice As Object         '��ǰҽ���б���� vsAdvice
Public gobjCmdAlley As Object           '��ǰPASS����ʷ��ť

Public glngModel As Long                '��ǰ����gbytModel 0-����༭,1-סԺ�༭��2-סԺҽ���嵥,3-��ʿУ��,4-����ҽ���嵥
Public gobjDiags As clsDiags              '����
Public gint���� As Integer              ' ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
Public gcolPASSExe As Collection        '�˵�����ӳ��
Public gcolPASSState As Collection      '�����˵�״ֵ̬ӳ��


Public gobjMap As clsPassMap  'ӳ�����
Public gobjPati As clsPatient
Public gbytOpen As Byte   '�����������
Public glngDrugID As Long    '��¼����һ�δ��˵�ҩƷID

'�������ܺ�
Public Enum G_PASS_MK
    MK_���PASS�˵�״̬ = 0
    MK_סԺ������� = 1
    MK_סԺ�ύ��� = 2
    MK_�ֹ�������� = 3
    MK_��ҩ���� = 6
    MK_ϵͳ���� = 11
    MK_��ҩ�о� = 12
    MK_ҩƷ�����Ϣ = 13
    MK_��ҩ;�������Ϣ = 14
    MK_����״̬����ʷ�鿴 = 21
    MK_����״̬����ʷ = 22
    MK_���ﱣ����� = 33
    MK_ҩ���ٴ���Ϣ�ο� = 101
    MK_ҩƷ˵���� = 102
    MK_������ҩ���� = 103
    MK_����ֵ = 104
    MK_ҽԺҩƷ��Ϣ = 105
    MK_ҽҩ��Ϣ���� = 106
    MK_�й�ҩ�� = 107
    MK_ҩ��_ҩ���໥���� = 201
    MK_ҩ��_ʳ���໥ʹ�� = 202
    MK_����ע������� = 203
    MK_����ע������� = 204
    MK_����֢ = 205
    MK_������ = 206
    MK_��������ҩ = 207
    MK_��ͯ��ҩ = 208
    MK_��������ҩ = 209
    MK_��������ҩ = 210
    MK_�رո������� = 402 '�رյ�ǰ���и�������
    MK_��ʾ��ʾ���� = 403  '��ʾ��ʾ��ʾ����
End Enum

Public Enum G_PASS_MK4
    MK4_���PASS�˵�״̬ = 0
    MK4_���
    MK4_�Զ����
    MK4_ҩƷ˵���� = 11
    MK4_ҩ��ר�� = 21
    MK4_������ҩ���� = 31
    MK4_�й�ҩ�� = 41
    MK4_ҩƷ��Ҫ��Ϣ = 51
    MK4_ҩ���໥���� = 61
    MK4_ҩʳ�໥���� = 62
    MK4_�������� = 63
    MK4_����Ũ�� = 64
    MK4_ҩ�����֢ = 65
    MK4_ҩ����Ӧ֢ = 66
    MK4_������Ӧ = 67
    MK4_���𺦼��� = 68
    MK4_���𺦼��� = 69
    MK4_��ͯ��ҩ = 70
    MK4_������ҩ = 71
    MK4_������ҩ = 72
    MK4_������ҩ = 73
    MK4_������ҩ = 74
    MK4_�Ա���ҩ = 75
    MK4_ϸ����ҩ�� = 76
End Enum

'������3.0�˵�����ֵ
Public Enum G_MK_INDEX
    MK_IX_ҩ���ٴ���Ϣ�ο� = 0
    MK_IX_ҩƷ˵���� = 1
    MK_IX_�й�ҩ��
    MK_IX_������ҩ����
    MK_IX_����ֵ
    MK_IX_ר����Ϣ
    MK_IX_ҩ���໥����
    MK_IX_ҩʳ�໥����
    MK_IX_����ע�������
    MK_IX_����ע�������
    MK_IX_����֢
    MK_IX_������
    MK_IX_��������ҩ
    MK_IX_��ͯ��ҩ
    MK_IX_��������ҩ
    MK_IX_��������ҩ
    MK_IX_ҽҩ��Ϣ����
    MK_IX_ҩƷ�����Ϣ
    MK_IX_��ҩ;�������Ϣ
    MK_IX_ҽԺҩƷ��Ϣ
    MK_IX_ϵͳ����
    MK_IX_��ҩ�о�
    MK_IX_����
    MK_IX_���
End Enum
'������4.0�˵�����ֵ
Public Enum G_MK4_INDEX
    MK4_IX_ҩƷ˵���� = 0
    MK4_IX_ҩ��ר�� = 1
    MK4_IX_�й�ҩ��
    MK4_IX_������ҩ����
    MK4_IX_ҩƷ��Ҫ��Ϣ
    MK4_IX_ר����Ϣ
    MK4_IX_ҩ���໥����
    MK4_IX_ҩʳ�໥����
    MK4_IX_��������
    MK4_IX_����Ũ��
    MK4_IX_ҩ�����֢
    MK4_IX_ҩ����Ӧ֢
    MK4_IX_������Ӧ
    MK4_IX_���𺦼���
    MK4_IX_���𺦼���
    MK4_IX_��ͯ��ҩ
    MK4_IX_������ҩ
    MK4_IX_������ҩ
    MK4_IX_������ҩ
    MK4_IX_������ҩ
    MK4_IX_�Ա���ҩ
    MK4_IX_ϸ����ҩ��
    MK4_IX_ϵͳ����
    MK4_IX_��ҩ�о�
    MK4_IX_���
End Enum

'̫Ԫͨ ���ܺ�
Public Enum G_PASS_TYT
    TYT_��ҩ�淶 = 0
    TYT_ҩ����� = 1
    TYT_ҩƷ��ʾ = 2
    TYT_ҽҩ֪ʶ�� = 3
    TYT_ϵͳ���� = 4
    TYT_������� = 5
End Enum

Public Enum G_PASS_HZYY
    HZYY_ҩƷ˵���� = 0
    HZYY_ҩ����� = 1
End Enum

Public Enum G_PASS_UseStation
    US_InDoctor = 0     'סԺҽ��վ
    US_InNurse = 1      'סԺ��ʿվ
    US_Intech = 2       'סԺҽ��վ
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    PҩƷ������ҩ = 1341        '1341    ҩƷ������ҩ
    PҩƷ���ŷ�ҩ = 1342        '1342    ҩƷ���ŷ�ҩ
    PPIVA���� = 1345        '1345    PIVA����
End Enum

Public Enum G_TYPE_FUN
    FUN_ҽ����Ϣ = 1
    FUN_�����Ϣ = 2
    FUN_������� = 3
    FUN_ҽ����Ϣ_DTBS = 4
    FUN_����� = 5
    FUN_ҽ����Ϣ_HZYY = 6
    FUN_�����_HZYY = 7
End Enum

Public Enum G_TYPE_FLOATWIN
    FLOATWIN_CLOSE = 0   '�ر�
    FLOATWIN_DRUG = 1    'ҩƷ��Ϣ��ʾ��
    FLOATWIN_WARN = 2    '��ʾ����
End Enum

'������ָ������Ļ�����ϵ�λ��
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'��ô�������Ļ�����е�λ��
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'�ж�ָ���ĵ��Ƿ���ָ���ľ����ڲ�
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'׼������ʹ����ʼ������ǰ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'�����ƶ�����
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'��ȡ����״̬
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Const GWL_EXSTYLE  As Long = (-20)
Public Const WS_EX_TOPMOST As Long = &H8
Public Const HWND_TOPMOST As Long = -1

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�

    gbytPass = Val(zlDatabase.GetPara(30, glngSys))  '�ӿ�����
    If gbytPass = UNPASS Then Exit Function
    gstrVersion = zlDatabase.GetPara(228, glngSys) '��ʶ�ӿڰ汾��
    
    '��ʼ�ɹ��������ظ���ȡ����ֵ��gbytPass����ģ����û�Ȩ�޽��õ�ԭ�����Ϊ:0-UNPASS����ÿ����Ҫ���¶�ȡ��
    gbytOpenLog = Val(zlDatabase.GetPara(225, glngSys))
    If gbytPass = MK Then
        gbytSysSet = Val(zlDatabase.GetPara(226, glngSys))
    ElseIf gbytPass = DT Or gbytPass = YWS Then '��ͨ
        If gstrVersion = "4.0" Then gstrHOSCODE = zlDatabase.GetPara(90001, glngSys, , "1513")
    ElseIf gbytPass = HZYY Then '��������
        Call HZYY_GetPara
    End If
    gbytBlackLamp = Val(zlDatabase.GetPara(161, glngSys))  '�Ƿ��������ҩƷ
    gbytSuperVolume = Val(zlDatabase.GetPara(182, glngSys)) '�Ƿ��ֹ������ҩƷ
    gbytOutBlackLamp = Val(zlDatabase.GetPara(189, glngSys)) '�Ƿ�����Ժ��ִ�еĽ���ҩƷҽ��
    
    'Ƥ�Խ����Чʱ��
    gint�����Ǽ���Ч���� = Val(zlDatabase.GetPara(70, glngSys))

    
    InitSysPar = True
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
'���ܣ���ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    If str���� <> "" Then
        strSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str����)
    Else
        strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getרҵ����ְ��(ByVal lng��ԱID As Long) As String
'���ܣ���ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
   
    strSQL = "Select רҵ����ְ�� From ��Ա�� Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��ԱID)
    
    Getרҵ����ְ�� = "" & rsTmp!רҵ����ְ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Get���˹�����¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bytFunc As Byte) As ADODB.Recordset
'���ܣ���ȡ���˹�����¼
'bytFunc=1 ����ҩƷID
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng��ҳID = 0 Then
        If bytFunc = 0 Then
            strSQL = "Select Distinct ҩ��ID,ҩ����,����Դ����,������Ӧ From ���˹�����¼ Where ����ID=[1] And ���=1 And Nvl(����ʱ��,��¼ʱ��)>Trunc(Sysdate-[3])"
        Else
            strSQL = "Select Distinct Decode(b.ҩƷid, Null, a.ҩ��id, b.ҩƷid) As ҩ��id, a.ҩ����, a.����Դ����, a.������Ӧ From ���˹�����¼ A, ҩƷ��� B Where a.ҩ��id = b.ҩ��id(+) And ����ID=[1] And ���=1 And Nvl(����ʱ��,��¼ʱ��)>Trunc(Sysdate-[3])"
        End If
    Else
        If bytFunc = 0 Then
            strSQL = "Select Distinct ҩ��ID,ҩ����,����Դ����,������Ӧ From ���˹�����¼ Where ����ID=[1] And ��ҳID=[2] And ���=1"
        Else
            strSQL = "Select Distinct Decode(b.ҩƷid, Null, a.ҩ��id, b.ҩƷid) As ҩ��id, a.ҩ����, a.����Դ����, a.������Ӧ From ���˹�����¼ A, ҩƷ��� B Where a.ҩ��id = b.ҩ��id(+) And ����ID=[1] And ��ҳID=[2] And ���=1"
        End If
    End If
    Set Get���˹�����¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID, gint�����Ǽ���Ч����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������ϼ�¼(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str���� As String) As ADODB.Recordset
'���ܣ���ȡ������ϼ�¼
'������lng����ID�����ﲡ�˴��Һ�ID��סԺ���˴���ҳID
'       �������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'       ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select a.ID,a.����id, a.���id, a.�������, a.��ϴ���, Nvl(b.����, c.����) As ����, NVL(Nvl(b.����, c.����),a.�������) ����,A.��¼���� " & vbNewLine & _
             "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & vbNewLine & _
             "Where a.����id = [1] And a.��ҳid = [2] And ȡ��ʱ�� Is Null And ��¼��Դ IN (1, 3) And Instr(',' ||[3]|| ',', ',' || ������� || ',') > 0 And a.����id = b.Id(+) And" & vbNewLine & _
             "      a.���id = c.Id(+)" & vbNewLine & _
             "Order By ��¼��Դ, �������, ��ϴ���"
    Set Get������ϼ�¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng����ID, str����)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˲��������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ����ݲ���ID����ҳID��ȡ���˲��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH

    If lng��ҳID = 0 Then
        lng��ҳID = Val(zlDatabase.GetPara(21, glngSys))
        strSQL = "Select ���������" & vbNewLine & _
                 "From ���˹Һż�¼" & vbNewLine & _
                 "Where ����id = [1] And �Ǽ�ʱ�� > Trunc(Sysdate-[2]) And ��������� Is Not Null And Rownum = 1"
    Else
        strSQL = "Select ��Ϣֵ As ���������" & vbNewLine & _
                 "From ������ҳ�ӱ� Where ����id = [1] And ��ҳid = [2] And ��Ϣ�� = '���������'"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            Get���˲�������� = Get���˲�������� & "," & rsTmp!���������
            rsTmp.MoveNext
        Wend
        Get���˲�������� = Mid(Get���˲��������, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���������¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˹�����¼
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��������ID,��������,������ʼʱ��,��������ʱ�� From ���������¼ Where ����ID=[1] And ��ҳID=[2] "

    Set Get���������¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiOperation(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String) As ADODB.Recordset
'���ܣ���ȡ���˹�����¼
    Dim strSQL As String
    
    On Error GoTo errH
    If str�Һŵ� = "" Then
        strSQL = " And a.����id = [1] And a.��ҳid = [2] "
    Else
        strSQL = "  And a.�Һŵ� = [3] "
    End If
    strSQL = "Select a.Id, a.����ʱ��, c.����, c.����" & vbNewLine & _
               "From ����ҽ����¼ A, ������϶��� B, ��������Ŀ¼ C" & vbNewLine & _
               "Where a.������Ŀid = b.����id And b.����id = c.Id And a.������� = 'F' And a.ҽ��״̬  In (1,2,3,5,8) " & strSQL
    Set GetPatiOperation = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID, str�Һŵ�)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiSymptom(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ����ݲ���ID����ҳID��ȡ����֢״��̫Ԫͨ�ӿ�ʹ�ã�
'lng��ҳId :���ﴫ�Һ�ID
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.����,a.���� From ����֢״��¼ a " & vbNewLine & _
            "Where a.����ID=[1] And a.��ҳID=[2] "
    Set GetPatiSymptom = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ����ݲ���ID����ҳID��ȡ���˻�����Ϣ
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select A.סԺ��, A.��ǰ����, A.��������, Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�, Nvl(B.����, A.����) ����, A.�����, A.������,A.���֤��,B.���,B.����" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B" & vbNewLine & _
            "Where A.����id = B.����id And A.����id = [1] And B.��ҳid = [2]"

    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetƵ����Ϣ_����(ByVal strƵ�� As String, intƵ�ʴ��� As Integer, _
    intƵ�ʼ�� As Integer, str�����λ As String, str��Χ As String, Optional strƵ�ʱ��� As String) As Boolean
'���ܣ�����Ƶ�ʵ������Ϣ
'������strƵ��=Ƶ������
'      str��Χ=1-��ҽ,2-��ҽ,-1-һ����,-2-������
'���أ���������ȡ��ʱ������True�����򷵻�False
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
    
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    strSQL = "Select Ƶ�ʴ���,Ƶ�ʼ��,�����λ,���� From ����Ƶ����Ŀ Where ����=[1] And Instr([2],','||���÷�Χ||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strƵ��, "," & str��Χ & ",")
    If Not rsTmp.EOF Then
        intƵ�ʴ��� = Nvl(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = Nvl(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = Nvl(rsTmp!�����λ)
        strƵ�ʱ��� = "" & rsTmp!����
        GetƵ����Ϣ_���� = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function get��������(ByVal lng����ID As Long) As String
'���ܣ�������Ա������ȡ����
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select ���� From ���ű� Where id = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID)
    If rsTmp.RecordCount > 0 Then get�������� = rsTmp!����
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get��Ա���(ByVal str���� As String) As String
'���ܣ�������Ա������ȡ����
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select ��� From ��Ա�� Where ���� = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", str����)
    If rsTmp.RecordCount > 0 Then Get��Ա��� = rsTmp!���
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDoctorTitleType(ByVal strDoctTitle As String) As String
'���ܣ�����ҽ��ְ�Ʒ���ְ�����
'����ֵ��
'C --�����ڣ����ڣ�������ҽʦ������ҽʦ��ר��
'B������ҽʦ����ʦ
'A�������ϵ�����ְ��

    If InStr(";������;����;������ҽʦ;����ҽʦ;ר��;", ";" & strDoctTitle & ";") > 0 Then
        GetDoctorTitleType = "C"
    ElseIf InStr(";����ҽʦ;��ʦ;", ";" & strDoctTitle & ";") > 0 Then
        GetDoctorTitleType = "B"
    Else
        GetDoctorTitleType = "A"
    End If

End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
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
            UserInfo.רҵ�������� = RowValue("רҵ����ְ��", UserInfo.רҵ����ְ��, "����", "����")
            GetUserInfo = True
        End If
    End If
    gstrDBUser = UserInfo.�û���
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
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
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function PassCheckPrivs(ByVal lngModel As Long, Optional ByVal blnInit As Byte = False) As Boolean
'����:����ģ��Ż�ȡģ����е�Ȩ��
'����:blnInit -�Ƿ��ʼ��(סԺҽ��վ��ʼ��ʱ��Ҫ�ж�סԺҽ���´��סԺҽ�����͵ĺ�����ҩ���Ȩ��)
    Dim blnDo As Boolean
    
    Select Case lngModel
    
    Case PM_����༭, PM_����ҽ���嵥
        If InStr(GetInsidePrivs(p����ҽ���´�), "������ҩ���") > 0 Then blnDo = True
    Case PM_סԺҽ���嵥
        If blnInit Then
            If InStr(GetInsidePrivs(pסԺҽ���´�) & GetInsidePrivs(pסԺҽ������), "������ҩ���") > 0 Then blnDo = True
        Else
            If InStr(GetInsidePrivs(pסԺҽ���´�), "������ҩ���") > 0 Then blnDo = True
        End If
    Case PM_סԺ�༭
        If InStr(GetInsidePrivs(pסԺҽ���´�), "������ҩ���") > 0 Then blnDo = True
    Case PM_��ʿУ��
        If InStr(GetInsidePrivs(pסԺҽ������), "������ҩ���") > 0 Then blnDo = True
    Case PM_סԺ��ҳ
        blnDo = True
    Case PM_������ҩ, PM_���ŷ�ҩ, PM_PIVA����
        If InStr(GetInsidePrivs(lngModel), "������ҩ���") > 0 Then blnDo = True
    End Select
    
    PassCheckPrivs = blnDo
End Function

Public Function Getҩ����ҩ;��ID(ByVal lngҽ��ID As Long) As Long
'���ܣ�����ҩ���ĸ�ҩ;���е�ҽ��ID��ȡ��������ĿID
'˵��: DockInAdviceCheckWarn����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    strSQL = "Select a.������ĿID From ����ҽ����¼ a " & vbNewLine & _
            "Where a.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
    If rsTmp.RecordCount > 0 Then
        Getҩ����ҩ;��ID = rsTmp!������ĿID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��Ŀ����(lng��Ŀid As Long) As String
'���ܣ�����������Ŀ����
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ���� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlAdvice", lng��Ŀid)
    If Not rsTmp.EOF Then Get��Ŀ���� = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵��:��frmDockInAdviceһ����ҩ����һ��
    Dim i As Long, blnTmp As Boolean
    With gobjAdvice
        If .TextMatrix(lngRow, gobjCOL.intCOL�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Public Function InitAdviceRS(Optional ByVal bytFunc As Byte = 1) As ADODB.Recordset
'����:����ҽ����¼
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '�ֶ�����|�ֶ�����|�ֶγ��� ȱʡ�ֶ����� ΪadVarChar
    Select Case bytFunc
    
    Case FUN_ҽ����Ϣ
        strFields = "ҽ��ID||18,���ID||18,ҽ����Ч||1,ҽ�����||5,ҽ��״̬||3,��������||100,��������ID||18,����ҽ������||10,����ҽ��||100," & _
        "ҩƷID||18,ҩƷ����||100,��������||16,������λ||20,Ƶ��||50,�÷�||100,�÷�ID||18,����ʱ��||20,��ʼʱ��||20,����ʱ��||20,����||16,������λ||20," & _
        "��ҩĿ��||1,ҽ������||100,����||100,�������||18"
    Case FUN_�����Ϣ
        strFields = "���||2,��ϱ���||20,�������||100"
    Case FUN_�������
        strFields = "ҽ��ID|adBigInt|18,ҩƷ����||1000,�Ƿ����|adInteger|1,����ҩƷ˵��||100,״̬|adInteger|1"
    Case FUN_ҽ����Ϣ_DTBS
        strFields = "ҽ��ID||18,���ID||18,ҽ����Ч||1,ҽ�����||5,ҽ��״̬||3,�������||3,��������||100,��������ID||18,����ҽ������||10,����ҽ��||100," & _
        "������ĿID||18,ҩƷID||18,ҩƷ����||100,��������||16,������λ||20,Ƶ��||50,�÷�||100,�÷�ID||18,����ʱ��||20,��ʼʱ��||20,����ʱ��||20,����||16,������λ||20," & _
        "��ҩĿ��||1,ҽ������||100,��ʾ|adInteger|1,����||3,���||100,Ƶ�ʱ���||5,��ҩ����||1000,��־||1,��Ժ��ҩ|adInteger|1"
    Case FUN_ҽ����Ϣ_HZYY
        strFields = "ҽ��ID||18,���ID||18,ҽ����Ч||1,ҽ�����||5,ҽ��״̬||3,�������||3,��������||100,��������ID||18,����ҽ��ID||10,����ҽ��||100," & _
        "������ĿID||18,ҩƷID||18,ҩƷ����||100,��������||16,������λ||20,Ƶ��||50,�÷�||100,�÷�ID||18,����ʱ��||20,��ʼʱ��||20,����ʱ��||20,����||16,������λ||20," & _
        "��ҩĿ��||50,ҽ������||100,��ʾ|adInteger|1,����||3,���||100,Ƶ�ʱ���||5,��ҩ����||1000,��־||1,��Ժ��ҩ|adInteger|1,����ID|adBigInt|18,����||100," & _
        "רҵ����ְ��||50,��Һ||3"
    Case FUN_�����
        strFields = "��ʾֵ||3,ҽ��ID||18"
    Case FUN_�����_HZYY
        strFields = "DrugName||100,DrugID||18,advice||1000,source||100,GroupNo||18,Type||200,Message||1000,Severity|adInteger|2,recipeId||18"
    End Select
    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            For j = LBound(arrSubFeld) To UBound(arrSubFeld)
                
            Next
            strFieldName = arrSubFeld(0)
            If UCase(arrSubFeld(1) & "") = UCase("adVarChar") Then
                FieldType = adVarChar
            Else
                FieldType = adVarChar
            End If
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitAdviceRS = rs
End Function

Public Function RowValue(ByVal strTable As String, Optional ByVal arrValues As Variant, Optional ByVal strGetFields As String, Optional ByVal strWhereField As String = "ID") As Variant
'���ܣ���ȡָ����ָ���ֶ���Ϣ
'������strTable=��ȡ���ݵı�
'          arrValues=����ֵ�����Դ����飬Ҳ���Դ�����ֵ��Ҳ���Բ�����������ȡȫ��
'          strGetField=��ȡ���ֶ�,����ֶ��Զ��ŷָͬSQL��д��ȡ�ֶ�һ��
'          strWhereField=���˵��ֶΣ����ֶ�Ϊ�򵥵���ֵ���ַ����ͻ��������ͣ����������޷�֧��
'���أ�
'ֻ������һ����������ض���һ��ֵ��δ����NULLֵ����
'      strGetField=�����ֶ�
'      arrValues=Ϊ����ֵ���򲻸���һ��Ԫ�ص�����
'������������ؼ�¼��

    Dim rsTmp As New ADODB.Recordset, blnReturnRec As Boolean
    Dim strSQL As String
    Dim strWhere As String
    Dim arrPars As Variant
    Dim i As Long, strPars As String
    
    On Error GoTo errH
    blnReturnRec = True
    If TypeName(arrValues) = "Variant()" Then
        arrPars = arrValues
        For i = LBound(arrValues) To UBound(arrValues)
            strPars = strPars & ",[" & i + 1 & "]"
        Next
        If strGetFields <> "" Then '�������Ԫ�ز�����һ��,�һ�ȡ����Ԫ�أ��򲻷��ؼ�¼��
            If UBound(arrValues) - LBound(arrValues) + 1 <= 1 And Not strGetFields Like "*,*" Then blnReturnRec = False
        End If
        If strPars <> "" Then
            strWhere = " Where " & strWhereField & " In (" & strPars & ")"
        End If
    ElseIf TypeName(arrValues) <> "Error" Then
        '����ֵʱ������ȡ�����ֶΣ��򲻷��ڼ�¼��
         If strGetFields <> "" And Not strGetFields Like "*,*" Then blnReturnRec = False
        arrPars = Array(arrValues)
        strWhere = " Where " & strWhereField & "=[1]"
    Else
        strWhere = ""
    End If
    
    If strGetFields = "" Then strGetFields = "*"
    strSQL = "Select " & strGetFields & " From " & strTable & strWhere
    If strWhere <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "RowValue", arrPars)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "RowValue")
    End If
    If blnReturnRec Then
        Set RowValue = rsTmp
    Else
        If Not rsTmp.EOF Then
            RowValue = rsTmp.Fields(strGetFields).Value
        Else '��ȡ��ֵʱ��û�л�ȡ����ֵ���򷵻�Ĭ��ֵ
            If IsType(rsTmp.Fields(strGetFields).Type, adVarChar) Then
                RowValue = ""
            ElseIf IsType(rsTmp.Fields(strGetFields).Type, adInteger) Then
                RowValue = 0
            Else
                RowValue = Null
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, Optional blnShowZero As Boolean = True, Optional ByVal blnAddZero As Boolean) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,
'          intBit=���С��λ��
'         blnShowZero=vNumberΪ0ʱ�Ƿ���ʾ0ֵ
'         blnAddZero=С��λ�����Ƿ���
'���أ���ʽ������ַ���
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
    
    If Not blnAddZero Then 'С��λ�����ʾ�㡣��1.0100 ��Ϊ1.01
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
    Else 'С��λ�����㲹��.��3λС����1.1��Ϊ1.100
        strNumber = Format(vNumber, "#0." & String(intBit, "0"))
    End If
    FormatEx = strNumber
End Function

Public Function FullDate(ByVal strText As String, Optional blnTime As Boolean = True, Optional ByVal strMintime As String, Optional strMaxtTime As String) As String
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
'������blnTime=�Ƿ���ʱ�䲿��
'������strMintime=����ʱ�������
'          strOutTime=����ʱ�������
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = zlDatabase.Currentdate
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
    
    If IsDate(strTmp) Then
        If strMintime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(strMintime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMintime
            End If
        End If
        If strMaxtTime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") > Format(strMaxtTime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMaxtTime
            End If
        End If
        If Not blnTime Then
            strTmp = Format(strTmp, "yyyy-MM-dd")
        End If
    End If
    FullDate = strTmp
End Function

Public Function GetDrugID(ByVal str������ĿID As String) As Variant
'����:����ҩƷID���¼
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH
    arrTmp = Split(str������ĿID, ",")
    If UBound(arrTmp) = 0 Then
        strSQL = "Select ҩ��ID,ҩƷID from ҩƷ��� where ҩ��id=[1] and rownum <2"
    ElseIf UBound(arrTmp) > 0 Then
        strSQL = "Select a.ҩ��id, Max(a.ҩƷid) As ҩƷid" & vbNewLine & _
        "From ҩƷ��� A" & vbNewLine & _
        "Where a.ҩ��id In (Select * From Table(f_Num2list([1])))" & vbNewLine & _
        "Group By a.ҩ��id"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", str������ĿID)
    
    If UBound(arrTmp) = 0 Then
        If Not rs.EOF Then
            GetDrugID = rs!ҩƷID & ""
        Else
            GetDrugID = ""
        End If
    Else
        Set GetDrugID = rs
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get��Ա��Ϣ(ByVal str���� As String) As ADODB.Recordset
'���ܣ�������Ա����,���

    Dim strSQL As String
 
    strSQL = "Select ����,��� From ��Ա�� Where ���� In (Select * From Table(f_Str2list([1])))"
    
    On Error GoTo errH
    Set Get��Ա��Ϣ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", str����)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRS(ByVal strTableName As String, ByVal strFileds As String, ByVal strInput As String, _
        Optional ByVal strWhere As String = "ID", Optional ByVal bytModel As Byte = 0, Optional ByVal bytType As Byte = 0) As Variant
'����:����ָ����ָ���ֶεļ�¼��
'������strTableName-����
'     strFileds
'     strInput ��ʽ1(1����������)��ID1,ID2,...
'              ��ʽ2(2����������)������1,��Χ1;����2,��Χ2;...
'             strSQL = "Select ����, ����, ���÷�Χ" & vbNewLine & _
'                "From ����Ƶ����Ŀ" & vbNewLine & _
'                "Where (����, ���÷�Χ) In (Select /*+cardinality(B,10)*/" & vbNewLine & _
'                "                      C1, C2" & vbNewLine & _
'                "                     From Table(f_Str2list2('ÿ�����,1|ÿ������,1', ';', ',')) B)"
'    bytModel=1 ��������Ϊ����
'    ��bytModel=1ʱ�� bytType=0-����� C1,C2 ͬΪ�ַ��� =1-C1(Number),C2(Number);=2-C1(char),C2(Number);=3-C1(Number),C2(Char)
'    ��bytModel=0ʱ�� bytType=0-f_num2list; bytType=1 f_Str2list


    Dim strSQL As String
    Dim strSub As String
    Dim strFun As String
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    If bytModel = 1 Then
        If bytType = 0 Then
            strSub = " C1,C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 1 Then
            strSub = " C1,C2 "
            strFun = "f_num2list2"
        ElseIf bytType = 2 Then
            strSub = "C1,To_Number(C2) As C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 3 Then
            strSub = " To_Number(C1) As C1,C2 "
            strFun = "f_Str2list2"
        End If
        strSQL = " Select  " & strFileds & vbNewLine & _
                " From  " & strTableName & vbNewLine & _
                " Where (" & strWhere & ") In (Select /*+cardinality(B,10)*/" & vbNewLine & _
                "                    " & strSub & vbNewLine & _
                "                     From Table(" & strFun & "([1], ';', ',')) B)"
    Else
        If bytType = 0 Then
            strFun = "f_num2list"
        ElseIf bytType = 1 Then
            strFun = "f_Str2list"
        End If
        arrTmp = Split(strInput, ",")
        If UBound(arrTmp) = 0 Or strInput = "" Then
            strSQL = "Select " & strFileds & "  From " & strTableName & " Where " & strWhere & " = [1]"
        ElseIf UBound(arrTmp) > 0 Then
            strSQL = "Select " & strFileds & vbNewLine & _
            "From " & strTableName & vbNewLine & _
            "Where " & strWhere & " In (Select /*+cardinality(A,10)*/ * From Table(" & strFun & "([1]))A )"
        End If
    End If
    Set GetRS = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", strInput)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AddDrugReason(ByRef objMap As Object, ByRef rsOut As ADODB.Recordset) As Boolean
'------------------------------------------------------------------------
'����:����ҩƷ��ӽ���˵��
'����:
'objMap-���������
'rsOut-�������
'����:True-����ҽ�����棨�����ڽ���ҩƷ,������д����˵��;���ڽ���ҩƷ��������д����˵����,False-��ֹҽ�����棨���ڽ���ҩƷ�ҽ���ҩƷ˵��δ������д��
'˵��:��ҩ�䷽����˵����������ҩ
'-----------------------------------------------------------------------
    Dim i As Long
    Dim strReason As String
    
    If rsOut Is Nothing Then AddDrugReason = True: Exit Function
    
    rsOut.Filter = "�Ƿ����=1"
    
    For i = 1 To rsOut.RecordCount
        strReason = rsOut!����ҩƷ˵�� & ""
        Call zlCommFun.ShowMsgBox("����˵��", "^��鷢�ֽ�����ҩ:��" & rsOut!ҩƷ���� & "��" & _
            vbCrLf & vbCrLf & "����¼�������ҩ˵����������ҽ����^", "!ȷ��(&O),?ȡ��(&C)", objMap.frmMain, vbInformation, , , , , , "����˵����", 99, strReason)
        If strReason = "" Then
            Exit Function
        Else
            rsOut!����ҩƷ˵�� = strReason
        End If
        rsOut.MoveNext
    Next
    AddDrugReason = True
End Function

Public Function ReadXML(ByVal strXML As String) As ADODB.Recordset
'����:���ص���ҩƷ���ʾֵ
'xmlģ��
'    <his_results_xml fun_id="1006">
'    <result>
'       <type>ZDXGYWSY</type>
'       <level>2</level>
'       <prescA_hiscode>669</prescA_hiscode>
'       <mediA_hiscode>14686</mediA_hiscode>
'       <mediA_name>�����Ƭ</mediA_name>
'       <groupA>669</groupA>
'       <prescB_hiscode /><mediB_hiscode />
'       <mediB_name />
'        <groupB />
'    </result>
'    <result>
'    <type>XHZYWT</type>
'    <level>2</level>
'    <prescA_hiscode>669</prescA_hiscode>
'    <mediA_hiscode>14686</mediA_hiscode>
'    <mediA_name>�����Ƭ</mediA_name>
'    <groupA>669</groupA>
'    <prescB_hiscode>671</prescB_hiscode><mediB_hiscode>14250</mediB_hiscode>
'    <mediB_name>ά����CƬ</mediB_name>
'    <groupB>671</groupB>
'   </result>
'   <types>;ZDXGYWSY;XHZYWT;YHGXHCGYFYLWT_PC;YHGXHCGYFYLWT_DR;</types>
'</his_results_xml>


    Dim xmlDoc As DOMDocument
    Dim xmlRoot As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodes As IXMLDOMNodeList
    Dim rsRet As ADODB.Recordset
    
    Dim str��ʾֵ As String
    Dim strҽ��ID As String
    
    On Error GoTo errH
    
    Set xmlDoc = New DOMDocument
    xmlDoc.loadXML (strXML)
    '����������κ�Ԫ�أ����˳�
    If xmlDoc.documentElement Is Nothing Then
        Set xmlDoc = Nothing
        Exit Function
    End If
    
    Set rsRet = InitAdviceRS(FUN_�����)
    '��ȡXML����
    Set xmlRoot = xmlDoc.selectSingleNode("his_results_xml")
    Set xmlNodes = xmlRoot.selectNodes("result")

    If Not xmlNodes Is Nothing Then
        For Each xmlNode In xmlNodes
            str��ʾֵ = xmlNode.selectSingleNode("level").Text
            If Val(str��ʾֵ) > 0 Then
                strҽ��ID = xmlNode.selectSingleNode("prescA_hiscode").Text
                If Val(strҽ��ID) <> 0 Then
                    rsRet.Filter = "ҽ��ID ='" & strҽ��ID & "'"
                    If Not rsRet.EOF Then
                        If Val(rsRet!��ʾֵ & "") < Val(str��ʾֵ) Then
                            rsRet!��ʾֵ = str��ʾֵ
                        End If
                    Else
                        rsRet.AddNew
                        rsRet!��ʾֵ = str��ʾֵ
                        rsRet!ҽ��ID = strҽ��ID
                        rsRet.Update
                    End If
                End If
                strҽ��ID = xmlNode.selectSingleNode("prescB_hiscode").Text
                If Val(strҽ��ID) > 0 Then
                    rsRet.Filter = "ҽ��ID ='" & strҽ��ID & "'"
                    
                    If Not rsRet.EOF Then
                        If Val(rsRet!��ʾֵ & "") < Val(str��ʾֵ) Then
                            rsRet!��ʾֵ = str��ʾֵ
                        End If
                    Else
                        rsRet.AddNew
                        rsRet!��ʾֵ = str��ʾֵ
                        rsRet!ҽ��ID = strҽ��ID
                        rsRet.Update
                    End If
                End If
            End If
        Next
    End If
    
    If rsRet.RecordCount > 0 Then rsRet.Filter = ""
    
    Set ReadXML = rsRet
    Exit Function
errH:
    MsgBox "ReadXML �����:" & Err.Number & "��������:" & Err.Description, vbOKOnly, gstrSysName
End Function

Public Function Get��ҩ�䷽(ByVal str��IDs As String) As ADODB.Recordset
'����:����ҩƷID���¼
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select a.Id, a.���id, a.ҽ����Ч, a.ҽ��״̬,a.�������,a.�շ�ϸĿID as ҩƷID,a.ҽ������ As ҩƷ����,a.���,a.��������, d.���㵥λ As ������λ,a.ִ��Ƶ�� as Ƶ��, a.�����λ, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.��ʼִ��ʱ�� As ��ʼʱ��," & vbNewLine & _
            "       a.ִ����ֹʱ�� As ��ֹʱ��, a.����ʱ��, a.ͣ��ʱ��, c.���� As �÷�, c.Id As �÷�id, a.ִ������, b.ִ������ As ��ִ������,a.����ҽ��,a.��ҩĿ��,a.����, " & vbNewLine & _
            "       a.�ܸ�����,f.סԺ��λ As ������λ,f.���ﵥλ, a.ҽ������, a.��������id" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ����¼ B, ������ĿĿ¼ C, ������ĿĿ¼ D, ҩƷ��� F" & vbNewLine & _
            "Where a.���id = b.Id And b.������Ŀid = c.Id And a.������Ŀid = d.Id And a.�շ�ϸĿid = f.ҩƷid(+) And" & vbNewLine & _
            "      a.���id in (Select * From Table(f_Num2list([1]))) And a.������� = '7'"


    Set rs = zlDatabase.OpenSQLRecord(strSQL, "��ҩ��ID", str��IDs)

    Set Get��ҩ�䷽ = rs

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����(ByVal str��IDs As String) As ADODB.Recordset
'����:����ҩƷID���¼
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select a.Id, a.ҽ������" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B " & vbNewLine & _
            "Where A.������ĿID = B.ID And A.������� ='E' And B.�������� = '2' And b.ִ�з��� = 1 And NVL(a.ҽ������,'��') <> '��' And " & vbNewLine & _
            "      a.ID in (Select /*+cardinality(A,10)*/ * From Table(f_Num2list([1])) A) "


    Set rs = zlDatabase.OpenSQLRecord(strSQL, "ҽ������", str��IDs)

    Set Get���� = rs

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncGetDripInfo(ByVal lngIndex As Long, ByVal strDrip As String) As String
'����:����ָ����JSON��
'�ַ�������:
'{ "type":"druginfo","index":"drug001","driprate":"60","driptime":"120"}
'driprate��   60   ��ʾ  ÿ����60��
'driptime����ʾ��������Ҫ��ʱ�䣬���û�оʹ��մ�
'��������Ǹ�����ֵ���ʹ����ġ�
'��λΪ���� ���㵥λ1����=20��

    Dim strRet As String
    Dim arrTmp As Variant
    
    strRet = "{""type"":""druginfo"",""index"":""" & lngIndex & """,""driprate"":""60"",""driptime"":""""}"
    If InStr(strDrip, "��/����") > 0 Then
        strDrip = Replace(strDrip, "��/����", "")
        arrTmp = Split(strDrip, "-")
        If UBound(arrTmp) = 1 Then
            strDrip = arrTmp(1)
        Else
            strDrip = arrTmp(0)
        End If
        strRet = Replace(strRet, "60", strDrip)
    ElseIf InStr(strDrip, "����/Сʱ") > 0 Then
        strDrip = Replace(strDrip, "����/Сʱ", "")
        arrTmp = Split(strDrip, "-")
        If UBound(arrTmp) = 1 Then
            strDrip = arrTmp(1)
        Else
            strDrip = arrTmp(0)
        End If
        strDrip = (Val(strDrip) \ 60) * 20
        strRet = Replace(strRet, "60", strDrip)
    Else
        strRet = ""
    End If
    
    FuncGetDripInfo = strRet
End Function

Public Function StrConvToNormal(ByVal strIn As String) As String
'���ܣ���StrConv(str,vbFromUnicode)ת��ʱ,��ʱ����Ϊ�������뵼��ת����xml������һ��������Ч��XML����
    Dim strChar As String
    Dim strRet As String
    Dim i As Long
    
    For i = 1 To Len(strIn)
        strChar = Mid(strIn, i, 1)
        If InStr(G_STR_MATCH & "=", strChar) > 0 Then
            strRet = strRet & strChar
        End If
    Next
    StrConvToNormal = strRet
End Function

'* ************************************** *
'* ģ�����ƣ�modCharset.bas
'* ģ�鹦�ܣ�GB2312��UTF8�໥ת������
'* ���ߣ�lyserver
'* ************************************** *

'- ------------------------------------------- -
'  ����˵����GB2312ת��ΪUTF8
'- ------------------------------------------- -

Public Function GB2312ToUTF8(strIn As String, Optional ByVal ReturnValueType As VbVarType = vbString) As Variant
    Dim adoStream As Object

    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 2 'adTypeText
    adoStream.Open
    adoStream.WriteText strIn
    adoStream.Position = 0
    adoStream.Type = 1 'adTypeBinary
    GB2312ToUTF8 = adoStream.Read()
    adoStream.Close

    If ReturnValueType = vbString Then GB2312ToUTF8 = Mid(GB2312ToUTF8, 1)
End Function

'- ------------------------------------------- -
'  ����˵����UTF8ת��ΪGB2312
'- ------------------------------------------- -
Public Function UTF8ToGB2312(ByVal varIn As Variant) As String
    Dim bytesData() As Byte
    Dim adoStream As Object

    bytesData = varIn
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 1 'adTypeBinary
    adoStream.Open
    adoStream.Write bytesData
    adoStream.Position = 0
    adoStream.Type = 2 'adTypeText
    UTF8ToGB2312 = adoStream.ReadText()
    adoStream.Close
End Function


Public Function WinHttpPost(ByVal strUrl As String, ByVal strData As String, ByVal DataStic As DataEnum, Optional ByVal strHeader As String, Optional ByVal strMethod As String = "POST") As Variant
'֧��HTTPS����
'����:strHeader ��ֵ��ʽ��HeaderName:HeaderValue ����:CONTENT-TYPE:application/json
    Dim XMLHTTP As WinHttp.WinHttpRequest
    Dim DataS As String
    Dim DataB() As Byte
    Dim varHeader As Variant
    Dim varHeaderItem As Variant
    Dim i As Long

    On Error GoTo errH:
       
8      Set XMLHTTP = New WinHttpRequest
9      XMLHTTP.Open strMethod, strUrl
10      If strHeader <> "" Then
            varHeader = Split(strHeader, ",")
            For i = LBound(varHeader) To UBound(varHeader)
                varHeaderItem = Split(varHeader(i), ":")
                XMLHTTP.setRequestHeader varHeaderItem(0), varHeaderItem(1)
            Next
        End If

13     XMLHTTP.send strData

110     Do Until XMLHTTP.Status = 200
112         DoEvents
        Loop

    '-----------------------------��������
114 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
116     DataS = XMLHTTP.responseText
118     WinHttpPost = DataS
120 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
122     DataB = XMLHTTP.responseBody
124     WinHttpPost = DataS
126 Case responseBody + responseText
        '---------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
128     DataS = BytesToStr(XMLHTTP.responseBody)
130     WinHttpPost = DataS
132 Case Else
        '--------------------------------��Ч�ķ���
134     WinHttpPost = ""
    End Select

    '------------------------------------�ͷſռ�
136     Set XMLHTTP = Nothing

    Exit Function

errH:
138     WinHttpPost = ""
140     MsgBox "WinHttpPostʧ�ܣ�" & vbNewLine & "�����:" & Err.Number & vbCrLf & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
End Function

'==========================================================
'| ģ �� �� | XMLHTTP
'| ˵    �� | ���Inet�ؼ���ʵ������ͨѶ
'---------------------------------------------------------------------------����Begin����---------------------------------------------------------------------------------------
'==========================================================
Public Function HttpGet(ByVal Url As String, ByVal DataStic As DataEnum) As Variant
    Dim XMLHTTP As Object
    Dim DataS As String
    Dim DataB() As Byte

    On Error GoTo errH:

100 Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
102 XMLHTTP.Open "get", Url, True
104 XMLHTTP.send

106 Do While XMLHTTP.readyState <> 4
108     DoEvents
    Loop

    '--------------------------------------��������
110 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
112     DataS = XMLHTTP.responseText
114     HttpGet = DataS
116 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
118     DataB = XMLHTTP.responseBody
120     HttpGet = DataB
122 Case responseBody + responseText
        '------------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
124     DataS = BytesToStr(XMLHTTP.responseBody)
126     HttpGet = DataS
128 Case Else
        '--------------------------------��Ч�ķ���
130     HttpGet = ""
    End Select

    '--------------------------------------�ͷſռ�
132 Set XMLHTTP = Nothing

    Exit Function

errH:
134 HttpGet = ""
136 MsgBox "HttpGetʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
End Function

Public Function HttpPost(ByVal strUrl As String, ByVal strData As String, ByVal DataStic As DataEnum, Optional ByVal strCONTENTTYPE As String) As Variant
    Dim XMLHTTP As Object
    Dim DataS As String
    Dim DataB() As Byte

    On Error GoTo errH:

100 Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
102 XMLHTTP.Open "POST", strUrl, True
104 XMLHTTP.setRequestHeader "Content-Length", Len(HttpPost)
    If strCONTENTTYPE = "" Then
106     XMLHTTP.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
    Else
        XMLHTTP.setRequestHeader "CONTENT-TYPE", strCONTENTTYPE  '"application/x-www-form-urlencoded; charset=utf-8"
    End If
108 XMLHTTP.send (strData)

110 Do Until XMLHTTP.readyState = 4
112     DoEvents
    Loop

    '-----------------------------��������
114 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
116     DataS = XMLHTTP.responseText
118     HttpPost = DataS
120 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
122     DataB = XMLHTTP.responseBody
124     HttpPost = DataS
126 Case responseBody + responseText
        '---------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
128     DataS = BytesToStr(XMLHTTP.responseBody)
130     HttpPost = DataS
    Case 6
        HttpPost = XMLHTTP.responseXML
132 Case Else
        '--------------------------------��Ч�ķ���
134     HttpPost = ""
    End Select

    '------------------------------------�ͷſռ�
136     Set XMLHTTP = Nothing

    Exit Function

errH:
138     HttpPost = ""
140     MsgBox "HttpPostʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
End Function

Private Function BytesToStr(ByVal vInput As Variant) As String
    
    Dim strReturn       As String
    Dim i               As Long
    Dim intPrevCharCode As Integer
    Dim intNextCharCode As Integer

    For i = 1 To LenB(vInput)
        intPrevCharCode = AscB(MidB(vInput, i, 1))
        If intPrevCharCode < &H80 Then
            strReturn = strReturn & Chr(intPrevCharCode)
        Else
            intNextCharCode = AscB(MidB(vInput, i + 1, 1))
            strReturn = strReturn & Chr(CLng(intPrevCharCode) * &H100 + CInt(intNextCharCode))
            i = i + 1
        End If
    Next

    BytesToStr = strReturn
End Function

Public Function CreatePlugInOK() As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModel)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub


Public Sub UpdateThirdPara(ByVal strCAName As String, ByVal intParaNum As Integer, ByVal strParaName As String, ByVal strParaValue As String, ByVal strParaTip As String)
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Zl_�����ӿ�����_Update('" & strCAName & "'," & intParaNum & ",'" & strParaName & "','" & strParaValue & "','" & strParaTip & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, gstrSysName)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetThirdPara(ByVal strCAName As String, ByVal varPara As Variant) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    If TypeName(varPara) = "String" Then
        strSQL = " And ������ = [3]"
    Else
        strSQL = " And ������ = [2]"
    End If
    On Error GoTo errH
    strSQL = "Select ����ֵ From �����ӿ����� Where �ӿ��� = [1] " & strSQL
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strCAName, Val(varPara), CStr(varPara))
    If Not rsTmp.EOF Then GetThirdPara = rsTmp!����ֵ & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function URLEncode(ByVal strParameter As String, Optional strEncodeType As String = "utf8") As String
          Dim strTemp As String
          Dim strRet As String
          Dim strInput As String
          
          Dim i As Long
          Dim lngValue As Long
          Dim lngLen As Long
          Dim lngMax As Long
          
          Dim bytData() As Byte
10        On Error GoTo errH
20        lngLen = 32767
30        Do While Len(strParameter) > 0
40            lngMax = Len(strParameter)
50            If lngMax > lngLen Then
60                strInput = Mid(strParameter, 1, lngLen)
70                strParameter = Mid(strParameter, lngLen + 1, lngMax - lngLen)
80            Else
90                strInput = strParameter
100               strParameter = ""
110           End If
120           strTemp = ""
130           If "UTF8" = UCase(strEncodeType) Then
140               bytData = StringToUTF8Bytes(strInput)
150           Else
160               bytData = StrConv(strInput, vbFromUnicode)
170           End If
              
180           For i = 0 To UBound(bytData)
190               lngValue = bytData(i)
200               If (lngValue >= 48 And lngValue <= 57) Or _
                      (lngValue >= 65 And lngValue <= 90) Or _
                      (lngValue >= 97 And lngValue <= 122) Or _
                       InStr("$-_.+*'()", Chr(lngValue)) > 0 Then
                       '�����ַ���ת"$-_.+*'()"
210                   strTemp = strTemp & Chr(lngValue)
220               ElseIf lngValue = 32 Then
                      '�ո�
230                   strTemp = strTemp & "+"
240               Else
250                   If lngValue <= 15 Then
260                       strTemp = strTemp & "%0" & UCase(Hex(lngValue))
270                   Else
280                       strTemp = strTemp & "%" & UCase(Hex(lngValue))
290                   End If
300               End If
310           Next
320           strRet = strRet & strTemp
330       Loop
340       URLEncode = strRet
350       Exit Function
errH:
360       MsgBox "��������:" & Err.Description & vbCrLf & _
                  "������:" & Erl() & vbCrLf & _
                  "�����:" & Err.Number, vbExclamation, G_STR_PASS
End Function

'======================================================================================================================
'����           StringToUTF8Bytes       ���ַ���ת��ΪUTF-8������ֽ�����
'����ֵ         Byte()                  16�����ַ���ת�����ֽ���
'����б�:
'������         ����                    ˵��
'strInput      String                  16�����ַ���
'======================================================================================================================
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '�ȼ��������ֽ���
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    'Ȼ��ת��
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function


