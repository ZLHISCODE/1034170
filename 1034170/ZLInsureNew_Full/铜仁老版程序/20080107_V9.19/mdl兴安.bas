Attribute VB_Name = "mdl�˰�"
Option Explicit
Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
End Type
Public InitInfor_�˰� As InitbaseInfor
Public g�������_�˰� As �������
'��ʾ��ǰ���еĴ����API����
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000
Const OFS_MAXPATHNAME = 128
Const OF_EXIST = &H4000

 
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
'�رյ�ǰ���еĴ����API����
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10
Public Declare Function apiOpenFile Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long


Private Type �������
    ���˱��            As String
    ����                As String
    ����                As String
    �Ա�                As String
    ����                As Integer
    ��λ����            As String
    ��Ա���            As String
    ��������            As String
    �ʻ�״̬            As String
    �ʻ����            As Double
    ��������ҩƷ���    As Double
    ���ִ���            As String
    ��������            As String
    ���������Ը���      As Double
    ������������ͳ��    As Double
    �����ܶ�            As Double       '��ʾ��ǰ�����ܶ�
    �������            As Variant      '������.
    byt����             As Byte         ''0-�����շѣ�1-סԺ
    סԺ�ǼǺ�          As String       'סԺ�ǼǺ�
    �������û���ͳ��    As Double       'סԺ��
    �������ô�ͳ��    As Double       'סԺ��
    ����ڼ���סԺ      As Integer      'סԺ��
    ����סԺ�𸶱�׼    As Double       'סԺ��
    ҽԺ����            As String       'סԺ��
    ҽԺ����            As String       'סԺ��
    ������ˮ��          As String
End Type
Public Enum ҵ������_�˰�
        �������ݿ����� = 0
        �ر����ݿ�����
        ����Աע��
        ��ȡ������Ϣ
        ��ȡҽ����Ŀ��Ϣ
        ����Ԥ����
        ������ϸд��
        ��������ύ
        ����������
        סԺ�Ǽ�
        ȡ��סԺ�Ǽ�
        סԺ��ϸд��
        סԺ��ϸȡ��
        סԺ����
        סԺ����ȡ��
        סԺ����ʼ
        סԺ�����ύ
        סԺ����ع�
End Enum
Private gobj�˰� As Object             '�����˰�����Dll
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'���ú�����������
Public Function ҽ����ʼ��_�˰�() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼҽ������ر���
    '--�����:
    '--������:
    '--��  ��:��ʼ���ɹ�������true�����򣬷���false
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim strUser As String
    Dim strServer As String
    Dim strPass As String
    Dim rsTemp As New ADODB.Recordset
    
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�˰�.ģ������ = True
    Else
        InitInfor_�˰�.ģ������ = False
    End If
    
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�˰�
    
    Call OpenRecordset(rsTemp, "��ȡҽԺ����")
    InitInfor_�˰�.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
        
        
    '������תվ����
    If ExcuteExeFile() = False Then Exit Function
    
    Err = 0
    On Error GoTo ErrHand:
    '�����ݿ�����.
    Dim intReturn As Integer
    
    '��ҽ�����ݿ�
    If ҵ������_�˰�(�������ݿ�����, "", "") = False Then
        Exit Function
    End If
    ҽ����ʼ��_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ExcuteExeFile() As Boolean
    'ִ�м�����
    Dim mError As String
    Dim strFile As String
    
    ExcuteExeFile = False
    
    '�ȹص�����
    Err = 0
    On Error Resume Next
    Call �رռ���
    
    strFile = App.Path & "\����������.exe"
    If FindFile(strFile) = False Then
        ShowMsgbox "�ļ�(" & App.Path & "\����������.exe)������!����������˾��ϵ"
        Exit Function
    End If
    Err = 0
    On Error Resume Next
    mError = Shell(strFile, vbNormalFocus)
    ExcuteExeFile = True
End Function

Public Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--����:����ָ�����ļ��Ƿ����
    '--����: ������ڴ��ļ�ΪTrue,����ΪFlase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Public Function ҽ����ֹ_�˰�() As Boolean
    
    Err = 0
    On Error Resume Next
    Set gobj�˰� = Nothing
    '��ҽ�����ݿ�
    If ҵ������_�˰�(�ر����ݿ�����, "", "") = False Then
        Exit Function
    End If
    
    Call �رռ���
    ҽ����ֹ_�˰� = True
End Function


Public Sub �رռ���()
    Dim app_hwnd As Long
    app_hwnd = FindWindow(vbNullString, "����������")
    SendMessage app_hwnd, WM_CLOSE, 0, 0
End Sub

Public Function ��ݱ�ʶ_�˰�(Optional bytType As Byte, Optional lng����ID As Long) As String
    Dim str��ע As String, rsPatient As New ADODB.Recordset
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    
    ��ݱ�ʶ_�˰� = frmIdentify�˰�.GetPatient(bytType, lng����ID)
    
End Function
Public Function ��ݱ�ʶ_�˰�2(ByVal strCard As String, ByVal strPass As String, Optional lng����ID As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    ��ݱ�ʶ_�˰�2 = frmIdentify�˰�.GetPatient(3, lng����ID)
End Function

Private Function Get������Ϣ(ByVal lng����ID As Long)
    Dim rsTemp As New ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '--�����ʻ�����Ա�˵��:
    '--����id, ����, ����, ���ţ�ҽ������), ҽ����(���˱��), ����, ��Ա���(��Ա���), ��λ����(��λ����), ˳���(��),
    '--����֤��(���ֱ���-��������), �ʻ����(�ʻ����), ��ǰ״̬, ����id����), ��ְ(1), �����(����), �Ҷȼ�(��),
    '--����ʱ��(��)
    
    Dim strTemp As String
    Dim strArr
    
    Err = 0
    On Error GoTo ErrHand:
    
    DebugTool "����Get������Ϣ����"
    
   '�����ʻ�:�����ֶ�:��������ҩƷ,���������Ը���,������������,�������û���ͳ��,�������ô�ͳ��,����סԺ����,����סԺ�𸶱�׼
    
    gstrSQL = "select a.����,a.ҽ����,a.����,a.��Ա���,a.��λ����,b.������λ,a.˳���,a.����֤��,a.�ʻ����,a.��ǰ״̬,a.����id,a.��ְ,a.�����,a.�Ҷȼ�,a.����ʱ��," & _
             "        b.����,b.�Ա�, b.����, b.��������, b.���֤��,A.��������ҩƷ,A.���������Ը���,A.������������,A.�������û���ͳ��,A.�������ô�ͳ��,A.����סԺ����,A.����סԺ�𸶱�׼" & _
             " from �����ʻ� a,������Ϣ b " & _
             " WHERE a.����id=" & lng����ID & " AND a.����id=b.����id and a.����=" & TYPE_�˰�
 
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
    
    With g�������_�˰�
        .���� = Nvl(rsTemp!����)
        .���˱�� = Nvl(rsTemp!ҽ����)
        .���� = Nvl(rsTemp!����)
        .�Ա� = Nvl(rsTemp!�Ա�)
        .��λ���� = Nvl(rsTemp!��λ����)
        .���� = Nvl(rsTemp!�����, 0)
        .��Ա��� = Nvl(rsTemp!��Ա���)
        .סԺ�ǼǺ� = Nvl(rsTemp!˳���)
        strTemp = Nvl(rsTemp!����֤��, "")
        If strTemp <> "" And InStr(1, strTemp, "-") <> 0 Then
            .���ִ��� = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
            .�������� = Mid(strTemp, InStr(1, strTemp, "-") + 1)
        Else
            .���ִ��� = ""
            .�������� = ""
        End If
        .�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
        
        .����ڼ���סԺ = Nvl(rsTemp!����סԺ����, 1)
        .�������ô�ͳ�� = Nvl(rsTemp!�������ô�ͳ��, 0)
        .�������û���ͳ�� = Nvl(rsTemp!�������û���ͳ��, 0)
        .����סԺ�𸶱�׼ = Nvl(rsTemp!����סԺ�𸶱�׼, 0)
        .������������ͳ�� = Nvl(rsTemp!������������, 0)
        .���������Ը��� = Nvl(rsTemp!���������Ը���, 0)
        .��������ҩƷ��� = Nvl(rsTemp!��������ҩƷ, 0)
    End With
    DebugTool "�˳�Get������Ϣ����"
Exit Function
ErrHand:
    DebugTool "��ȡ������Ϣʧ��" & vbCrLf & " �����:" & Err.Number & vbCrLf & " ������Ϣ:" & Err.Description
End Function

Public Function ��ݼ���_�˰�() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:Զ����ݼ���
    '--�����:
    '--������:
    '--��  ��:�ɹ�true,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutput As String
    Dim strInput As String
    Dim blnReturn As Boolean
    Dim strArr
    
        
    Err = 0
    On Error GoTo ErrHand:
    ��ݼ���_�˰� = False
    
    If g�������_�˰�.byt���� = 0 Then
        strInput = g�������_�˰�.����
    Else
        strInput = g�������_�˰�.���˱�� & vbTab
        strInput = strInput & InitInfor_�˰�.ҽԺ����
    End If
    
    DebugTool "������ݼ�����"
    
    'ҵ������
    blnReturn = ҵ������_�˰�(��ȡ������Ϣ, strInput, strOutput)
    
    If blnReturn = False Then
        Exit Function
    End If
    If strOutput = "" Then
        '���˺� /*200408*/
        DebugTool "��ȡ������Ϣʱ�����˴�����Ϊ����!"
        Exit Function
    End If
    strArr = Split(strOutput, vbTab)
    
    '�����ñ�����ֵ
    With g�������_�˰�
        'byt���� 0-����,1-סԺ
        If .byt���� = 0 Then
            .���˱�� = strArr(0)
            .���� = strArr(1)
            .�Ա� = strArr(2)
            .���� = Val(strArr(3))
            .��λ���� = strArr(4)
            .��Ա��� = strArr(5)
            .�ʻ�״̬ = strArr(6)
            .�ʻ���� = Val(strArr(7))
            .��������ҩƷ��� = Val(strArr(8))
            .���ִ��� = strArr(9)
            .�������� = strArr(10)
            .���������Ը��� = Val(strArr(11))
            .������������ͳ�� = Val(strArr(12))
            ��ݼ���_�˰� = True
            DebugTool "��ݼ���ɹ�"
            Exit Function
        End If
        .סԺ�ǼǺ� = strArr(0)
        .���� = strArr(1)
        .�Ա� = strArr(2)
        .���� = Val(strArr(3))
        .��λ���� = strArr(4)
        .��Ա��� = strArr(5)
        .�������û���ͳ�� = Val(strArr(6))
        .�������ô�ͳ�� = Val(strArr(7))
        .����ڼ���סԺ = Val(strArr(8))
        .����סԺ�𸶱�׼ = Val(strArr(9))
        .ҽԺ���� = strArr(10)
        .ҽԺ���� = strArr(11)
        .���ִ��� = ""
        .�������� = ""
    End With
    ��ݼ���_�˰� = True
    DebugTool "��ݼ���ɹ�"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    ��ݼ���_�˰� = False
End Function

Public Function �����������_�˰�(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strArr
    Dim strInput  As String
    Dim strOutput  As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    DebugTool "���������������ӿ�"
    
    With rs��ϸ
        Do While Not .EOF
            gstrSQL = "Select ����,���� From �շ�ϸĿ where id=" & Nvl(!�շ�ϸĿID, 0)
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ�ϸĿ����"
            If ҵ������_�˰�(��ȡҽ����Ŀ��Ϣ, Nvl(rsTemp!����), strOutput) = False Then
                Exit Function
            End If
            If strOutput = "" Then
                DebugTool "��ȡҽ����Ŀ��Ϣʱ,�����Ϊ����"
                Exit Function
            End If
            strArr = Split(strOutput, vbTab)
            
            '���:���,���,�������,ҽ����־,ҽ������,��������
            strInput = .AbsolutePosition & vbTab
            strInput = strInput & Nvl(!ʵ�ս��, 0) & vbTab
            strInput = strInput & strArr(0) & vbTab
            strInput = strInput & strArr(1) & vbTab
            strInput = strInput & strArr(2) & vbTab
            If g�������_�˰�.�������� = "��ͨҽ������" Then
                strInput = strInput & "1" & vbTab
            Else
                'ժҪ����:�մ���;������;�ε�λ;������;��������;��������
                strTemp = Nvl(!ժҪ)
                strTemp = strTemp & vbTab & ":" & vbTab & ":" & vbTab & ":" & vbTab & ":" & vbTab & ":" & vbTab & ":"
                strTemp = Split(strTemp, vbTab)(4)
                strTemp = Split(strTemp, ":")(1)
                strInput = strInput & IIf(Val(strTemp) = 0, 1, Val(strTemp)) & vbTab
            End If
            
            If ҵ������_�˰�(����Ԥ����, strInput, strOutput) = False Then
                Exit Function
            End If
            If strOutput = "" Then
                DebugTool "����Ԥ����ʱ,�����Ϊ����"
                Exit Function
            End If
            '����:���θ����ʻ����,���θ����ʻ����,�����Ը��ν��,����ͳ���ʽ��,�����Ը����,����ֵ
            strArr = Split(strOutput, vbTab)
            .MoveNext
        Loop
    End With
    
    g�������_�˰�.������� = strArr
    
    str���㷽ʽ = "�����ʻ�;" & Format(Val(strArr(0)), "###0.00;-###0.00;0;0") & ";0" '���λ��������ʻ�֧��,�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "ҽ������;" & Format(Val(strArr(3)), "###0.00;-###0.00;0;0") & ";0"
    
    DebugTool "�����������ɹ�"
    �����������_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get��ϸ��¼(ByRef lng����ID As Long, Optional strNO As String, Optional lng��¼���� As Long, Optional lng��¼״̬ As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���ν��ʵ���ϸ��¼
    '--�����:lng����ID-���ν��ʵ�ID��¼
    '         strno-���δ����ĵ��ݺ�,lng��¼����=��¼����,lng��¼״̬
    '--������:
    '--��  ��:SQL���
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim strFields As String
    
    If lng����ID = 0 Then
            strSql = " " & _
                "  Select Rownum ��ʶ��,A.ID,A.����ID,a.��ҳid,A.�շ�ϸĿid,������Ŀid,A.NO,A.��� ,A.��¼����,A.��¼״̬,A.����ʱ�� as ����ʱ�� ,c.���� as ��������,a.������ as ����ҽ��,nvl(a.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
                "      A.����*A.���� as ����,A.���㵥λ,A.ʵ�ս�� as ʵ�ʽ��,Round(A.ʵ�ս��/(A.����*A.����),4) as ʵ�ʼ۸�,A.ʵ�ս�� as ʵ�ս��, " & _
                "      A.�շ����,A.ժҪ,A.����Ա���� as ������," & _
                "      L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.˳���,L.����ID,L.����ʱ�� ,J.����,J.���� as ��Ʒ��,J.���" & _
                "  From (Select * From ���˷��ü�¼ Where ��¼״̬<>0 and NO='" & strNO & "' and ��¼����=" & lng��¼���� & " and ��¼״̬=" & lng��¼״̬ & " and  Nvl(���ӱ�־,0)<>9 ) A,���ű� C," & _
                "       �����ʻ� L,�շ�ϸĿ J " & _
                "  Where A.��������id=C.id(+)  and  A.����id=L.����id  and a.�շ�ϸĿid=J.id and L.����=" & gintInsure & "  " & _
                "  Order by a.NO,A.��¼����,A.��¼״̬,a.���"
                
    Else
        strSql = " " & _
            "  Select Rownum ��ʶ��,A.ID,A.����ID,a.��ҳid,A.�շ�ϸĿid,������Ŀid,A.NO,A.��� ,A.��¼����,A.��¼״̬,A.����ʱ�� as ����ʱ�� ,c.���� as ��������,a.������ as ����ҽ��,nvl(a.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "      A.����*A.���� as ����,A.���㵥λ,A.ʵ�ս�� as ʵ�ʽ��,Round(A.���ʽ��/(A.����*A.����),4) as ʵ�ʼ۸�,A.���ʽ�� as ʵ�ս��, " & _
            "      A.�շ����,A.ժҪ,A.����Ա���� as ������," & _
            "      L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.˳���,L.����ID,L.����ʱ��,J.���� ,J.���� as ��Ʒ��,J.���" & _
            "  From (Select * From ���˷��ü�¼ Where ��¼״̬<>0 and ����ID=" & lng����ID & " and  Nvl(���ӱ�־,0)<>9 ) A,���ű� C," & _
            "       �����ʻ� L,�շ�ϸĿ J " & _
            "  Where A.��������id=C.id(+) and  A.����id=L.����id and a.�շ�ϸĿid=J.id and L.����=" & gintInsure & _
            "   Order by a.NO,A.��¼����,A.��¼״̬,a.���"
    End If
    Get��ϸ��¼ = strSql
End Function
Public Function �������_�˰�(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    
    Dim lng����ID As Long, strOutput As String, strInput As String
    Dim strArr
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strҽ������ As String
    �������_�˰� = False
    
    DebugTool "�����������"
        
    Err = 0
    On Error GoTo ErrHand:
    
    '��ȡ��ϸ��¼
    gstrSQL = Get��ϸ��¼(lng����ID)
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ��ϸ��¼"
    If rs��ϸ.RecordCount = 0 Then
        ShowMsgbox "û��һ����ϸ��¼,���ܽ��н���!"
        Exit Function
    End If
    DebugTool "��ʼ����ϸ"
    g�������_�˰�.�����ܶ� = 0
    With rs��ϸ
            lng����ID = Nvl(!����ID, 0)
            Do While Not .EOF
                '��ȡҽ���������Ϣ
                If ҵ������_�˰�(��ȡҽ����Ŀ��Ϣ, Nvl(!����), strOutput) = False Then Exit Function
                If strOutput = "" Then
                    DebugTool "�ڶ�ȡ������Ŀ��Ϣʱ,û�д�����!"
                    Exit Function
                End If
                strArr = Split(strOutput, vbTab)
                
                '��ͨ����
                '���:kbname(ҽ������),ysname(��������),xh(���),fycode(���ô���),fyname(��������),gg(���),dw(��λ),dj(����),sl(����),je(���),fylb(�������),ypbz(ҽ�����),ybdm(ҽ������)
                '��������:
                '���:kbname(ҽ������),ysname(��������),xh(���),fycode(���ô���),fyname(��������),gg(���),dw(��λ),dj(����),sl(����),je(���),fylb(�������),ypbz(ҽ�����),ybdm(ҽ������),yf(�մ���),yl(������),yfdw(�ε�λ),mryl(������),cfts(��������),cfzl(��������)
                
                'д��ϸ��¼
                strInput = Nvl(!��������) & vbTab
                strInput = strInput & Nvl(!����ҽ��) & vbTab
                strInput = strInput & Nvl(!���, 0) & vbTab
                strInput = strInput & Nvl(!����) & vbTab
                strInput = strInput & Nvl(!��Ʒ��) & vbTab
                strInput = strInput & Nvl(!���) & vbTab
                strInput = strInput & Nvl(!���㵥λ) & vbTab
                strInput = strInput & Nvl(!ʵ�ʼ۸�, 0) & vbTab
                strInput = strInput & Nvl(!����, 0) & vbTab
                strInput = strInput & Nvl(!ʵ�ս��, 0) & vbTab
                strInput = strInput & strArr(0) & vbTab
                strInput = strInput & strArr(1) & vbTab
                strҽ������ = strArr(2)
                
                If g�������_�˰�.�������� = "��ͨҽ������" Then
                    strInput = strInput & strҽ������
                Else
                    'ժҪ����:�մ���;������;�ε�λ;������;��������;��������
                    '�մ���:2    ������:2    �ε�λ:Ƭ   ������:4    ��������:5  ��������:20 �ʻ����:0
                    strInput = strInput & strҽ������ & vbTab
                    strTemp = Nvl(!ժҪ, ":0" & vbTab & ":0" & vbTab & ":" & vbTab & ":0" & vbTab & ":1" & vbTab & ":0")
                    strTemp = strTemp & ":0" & vbTab & ":0" & vbTab & ":" & vbTab & ":0" & vbTab & ":1" & vbTab & ":0"
                    
                    strArr = Split(strTemp, vbTab)
                    
                    strInput = strInput & Val(Split(strArr(0), ":")(1)) & vbTab
                    strInput = strInput & Val(Split(strArr(1), ":")(1)) & vbTab
                    strInput = strInput & Split(strArr(2), ":")(1) & vbTab
                    strInput = strInput & Val(Split(strArr(3), ":")(1)) & vbTab
                    strInput = strInput & Val(Split(strArr(4), ":")(1)) & vbTab
                    strInput = strInput & Val(Split(strArr(5), ":")(1)) & vbTab
                End If
                strOutput = ""
                If ҵ������_�˰�(������ϸд��, strInput, strOutput) = False Then Exit Function
                If strOutput = "" Then
                    DebugTool "��������ϸд��ʱ,û�д�����!"
                    Exit Function
                End If
                strArr = Split(strOutput, vbTab)
                'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                'ժҪֵ:�մ���;������;�ε�λ;������;��������;��������;�������
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(!ժҪ) & vbTab & "�ʻ����:" & Val(strArr(0)) & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
                g�������_�˰�.�����ܶ� = g�������_�˰�.�����ܶ� + Nvl(!ʵ�ս��, 0)
                .MoveNext
            Loop
    End With
    
    DebugTool "��ϸ�ϴ��ɹ�������ʼ���㽻���ύ"

    '���Խ���
    strInput = ""
    If ҵ������_�˰�(��������ύ, strInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, vbTab)
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN(���θ����ʻ����),�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(�����Ը��ν��),�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN(����ͳ���ʽ��),���Ը����_IN,�����Ը����_IN(�����Ը����),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    
    With g�������_�˰�
        gstrSQL = "zl_���ս����¼_insert( 1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          "NUll," & Val(.�������(4)) & ",Null,NULL,NULL,null,Null,NULL," & _
         .�����ܶ� & "," & Val(.�������(0)) & ",Null," & _
         "Null," & Val(.�������(1)) & ",Null," & Val(.�������(2)) & "," & Val(.�������(3)) & ",'" & _
         strArr(0) & "',Null,Null,NULl)"
    End With
    Call ExecuteProcedure("���汣�ս����¼")
    DebugTool "�������ɹ�"
    �������_�˰� = True
    Exit Function
ErrHand:
    DebugTool "�������(�������_�˰�)" & vbCrLf & " �����:" & Err.Number & vbCrLf & "������Ϣ:" & Err.Description
End Function

Private Function Get����ID(ByVal lng����ID As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��ǰ������¼��IDֵ
    '--�����:
    '--������:
    '--��  ��:����ID
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    'ȡ������¼�Ľ���ID
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���²����Ľ���ID")
    If rsTemp.EOF Then
        Get����ID = 0
    Else
        Get����ID = Nvl(rsTemp!����ID, 0)
    End If
End Function

Public Function ����������_�˰�(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    

    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset
    Dim str������ˮ�� As String
    Dim lng����ID As Long
    Dim strOutput As String
    Dim strArr
    
    ����������_�˰� = False
    
    Err = 0
    On Error GoTo ErrHand
    DebugTool "��������������"
    
    gstrSQL = "Select * From ���ս����¼  where ��¼id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������ˮ��"
    
    lng����ID = Get����ID(lng����ID)
    str������ˮ�� = Nvl(rsTemp!֧��˳���)
    
    '����ȡ���������
    If ҵ������_�˰�(����������, str������ˮ��, strOutput) = False Then Exit Function
    If strOutput = "" Then
        strOutput = "0"
    End If
    strArr = Split(strOutput, vbTab)
    
    DebugTool "���뱣�汣�ս����¼"
   
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN(���θ����ʻ����),�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(�����Ը��ν��),�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN(����ͳ���ʽ��),���Ը����_IN,�����Ը����_IN(�����Ը����),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert( 1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      "NUll," & -1 * Nvl(rsTemp!�ʻ��ۼ�֧��, 0) & ",Null,NULL,NULL,null,Null,NULL," & _
     -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & ",Null," & _
     "Null," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & ",Null," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & _
     strArr(0) & "',Null,Null,NULl)"
    Call ExecuteProcedure("���汣�ս����¼")
    DebugTool "�����������ɹ�"
    ����������_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ҽ������_�˰�() As Boolean
    ҽ������_�˰� = True
End Function

Public Function ��Ժ�Ǽ�_�˰�(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    
    Err = 0: On Error GoTo ErrHand
    
    DebugTool "������Ժ�Ǽǽӿ�"
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ�����,���Ƚ��н���!"
        Exit Function
    End If
    
   ' Call Get������Ϣ(����ID)
    
    If סԺ��Ϣ�ύ(lng����ID, lng��ҳID) = False Then Exit Function


    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˰� & ")"
    Call ExecuteProcedure("������Ժ�Ǽ�")
    
    DebugTool "������Ժ�ɹ�"
    ��Ժ�Ǽ�_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�˰�(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    
    Err = 0
    On Error GoTo ErrHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_�˰� = False
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    
    Get������Ϣ lng����ID
    '����סԺ��Ϣ
    strInput = g�������_�˰�.סԺ�ǼǺ� & vbTab
    strInput = strInput & InitInfor_�˰�.ҽԺ����
    
    
    If ҵ������_�˰�(ȡ��סԺ�Ǽ�, strInput, strOutput) = False Then Exit Function
        
    DebugTool "����ҽ����ȡ��ҵ��ɹ�,����ʼ���±����ʻ������״̬��"
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˰� & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function ��Ժ�Ǽ�_�˰�(lng����ID As Long, lng��ҳID As Long) As Boolean
    
    Err = 0
    On Error GoTo ErrHand:
    DebugTool "�����Ժ�Ǽ�"
    
    ��Ժ�Ǽ�_�˰� = False
    Get������Ϣ lng����ID
    
    If ��Ժ������Ϣ_�˰�(lng����ID, lng��ҳID) = False Then Exit Function

    
    
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˰� & ")"
    Call ExecuteProcedure("��Ժ�Ǽ�")
    
    DebugTool "��Ժ�Ǽǳɹ�"
    ��Ժ�Ǽ�_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function ��Ժ�Ǽǳ���_�˰�(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand
    DebugTool "�����Ժ�Ǽǳ���!"
    ��Ժ�Ǽǳ���_�˰� = False
    
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "�ò����Ѿ�����,�����ٽ��г�Ժ����."
        Exit Function
    End If
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˰� & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    
    DebugTool "��Ժ�Ǽǳ����ɹ�!"
    ��Ժ�Ǽǳ���_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_�˰�(ByVal lng����ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    
    '����ʧ�����˳�
    Err = 0: On Error GoTo ErrHand:
    DebugTool "�����ȡ�����ʻ����(�������_�˰�)"
    gstrSQL = "Select Nvl(�ʻ����,0) �ʻ���� From �����ʻ� Where ����=" & gintInsure
    gstrSQL = gstrSQL & " And ����id=" & lng����ID
    
    Call OpenRecordset(rsTemp, "��ȡ�ʻ����")
    �������_�˰� = Nvl(rsTemp!�ʻ����, 0)
    
    DebugTool "��ȡ�ɹ�,���Ϊ:" & Nvl(rsTemp!�ʻ����, 0)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�˰�(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand:
    סԺ�������_�˰� = ""
    If bln���ʴ� = False Then Exit Function
    
    gstrSQL = "Select ��ǰ״̬ From �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�ж��Ƿ��Ժû��"
    If Nvl(rsTemp!��ǰ״̬, 0) = 1 Then
        ShowMsgbox "�ò��˻������Ժ,���Բ��ܽ���!"
        Exit Function
    End If
    With rsExse
        g�������_�˰�.�����ܶ� = 0
        Do While Not .EOF
            g�������_�˰�.�����ܶ� = g�������_�˰�.�����ܶ� + Nvl(!���, 0)
            .MoveNext
        Loop
    End With
'    If bln���ʴ� Then
'        '���ʴ�����֤���
'         If frmIdentify�˰�.GetPatient(4, lng����id) = "" Then Exit Function
'    End If
    Call Get������Ϣ(lng����ID)
    סԺ�������_�˰� = "ҽ������;" & g�������_�˰�.�����ܶ� & ";0"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function �ϴ�������ϸ(ByVal rs��ϸ As ADODB.Recordset) As Boolean

    '�ϴ�������ϸ
    Dim strInput As String, strOutput As String, strTemp As String
    Dim strArr
    Dim rsTemp As New ADODB.Recordset
     
    �ϴ�������ϸ = False
    DebugTool "�����ϴ�������ϸ����    "
    Err = 0
    On Error GoTo ErrHand:
    g�������_�˰�.�����ܶ� = 0
    With rs��ϸ
        If .RecordCount = 0 Then
            ShowMsgbox "���ν���û����ص���ϸ��¼,���ܽ����½�"
            Exit Function
        End If
        
        'дδ�ϴ�����ϸ��¼
        Do While Not .EOF
            
            If Nvl(!�Ƿ��ϴ�, 0) <> 1 And Nvl(!ʵ�ʽ��, 0) <> 0 Then
                    
                    If !��¼״̬ <> 1 Then
                        '��ʾ�������ļ�¼
                        '��ȷ��ԭʼ��¼�е���Ŀ���
                        gstrSQL = "Select ժҪ From ���˷��ü�¼ where mod(��¼״̬,3)=0 and NO='" & Nvl(!NO) & "' and ��¼����=" & Nvl(!��¼����, 0) & " and ���=" & Nvl(!���)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����������Ŀ���"
                        If rsTemp.RecordCount = 0 Then
                            ShowMsgbox "������ԭʼ����δ�ҵ�!" & Nvl(!NO)
                            Exit Function
                        End If
                        strTemp = Nvl(rsTemp!ժҪ) & vbTab & vbTab
                        strArr = Split(strTemp, vbTab)
                        If Trim(strArr(0)) = "" Then
                            ShowMsgbox "ԭʼ���ݼ�¼δ�ҵ���Ӧ����Ŀ���!" & vbCrLf & "���ݺ�:" & Nvl(!NO) & vbCrLf & " ����:" & Nvl(!����, 0) & vbCrLf & " �к�Ϊ:" & Nvl(!���)
                            Exit Function
                        End If
                        strTemp = strArr(0)
                        strInput = Nvl(!˳���) & vbTab
                        strInput = strInput & InitInfor_�˰�.ҽԺ���� & vbTab
                        strInput = strInput & Val(strArr(0)) & vbTab
                        strInput = strInput & Nvl(!������)
                        
                        If ҵ������_�˰�(סԺ��ϸȡ��, strInput, strOutput) = False Then Exit Function
                        strArr = Split(strOutput, vbTab)
                        
                        'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                        'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                        gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & ";;;;;" & vbTab & Val(strTemp) & vbTab & Val(strArr(0)) & "')"
                        zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
                    Else
                    
                            '��ȡҽ����Ŀ��Ϣ
                            If ҵ������_�˰�(��ȡҽ����Ŀ��Ϣ, Nvl(!����), strOutput) = False Then Exit Function
                            
                            
                            If strOutput = "" Then
                                DebugTool "�ڽ��д�����ϸ�ϴ������еĶ�ȡҽ����Ϣʱ�����ش�Ϊ����"
                                Exit Function
                            End If
                            strArr = Split(strOutput, vbTab)
                            '���:lsh(סԺ�ǼǺ�),yycode(ҽԺ����),rq(��������),kbname(ҽ������),ysname(��������),fycode(���ô���),fyname(��������),gg(���),dw(��λ),
                            '       dj(����),sl(����),je(���),fylb(�������),ypbz(ҽ�����),ybdm(ҽ������),czyname(������)
                            '����:bl(�����Ը�����),xh(��Ŀ���)
                            
                            strInput = Nvl(!˳���) & vbTab
                            strInput = strInput & InitInfor_�˰�.ҽԺ���� & vbTab
                            strInput = strInput & Format(!����ʱ��, "yyyyMMDD") & vbTab
                            strInput = strInput & Nvl(!��������) & vbTab
                            strInput = strInput & Nvl(!����ҽ��) & vbTab
                            strInput = strInput & Nvl(!����) & vbTab
                            strInput = strInput & Nvl(!��Ʒ��) & vbTab
                            strInput = strInput & Nvl(!���) & vbTab
                            strInput = strInput & Nvl(!���㵥λ) & vbTab
                            strInput = strInput & Nvl(!ʵ�ʼ۸�, 0) & vbTab
                            strInput = strInput & Nvl(!����, 0) & vbTab
                            strInput = strInput & Nvl(!ʵ�ս��, 0) & vbTab
                            strInput = strInput & strArr(0) & vbTab
                            strInput = strInput & strArr(1) & vbTab
                            strInput = strInput & strArr(2) & vbTab
                            strInput = strInput & Nvl(!������)
                                                        
                            If ҵ������_�˰�(סԺ��ϸд��, strInput, strOutput) = False Then Exit Function
                            strArr = Split(strOutput, vbTab)
                            
                            'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Val(strArr(0)) & vbTab & Val(strArr(1)) & "')"
                            zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
                        End If
                    End If
                g�������_�˰�.�����ܶ� = g�������_�˰�.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    �ϴ�������ϸ = True
    Exit Function
ErrHand:
    
    DebugTool "�ϴ�������ϸʧ��!" & vbCrLf & "�����:" & Err.Number & vbCrLf & "��������:" & Err.Description
 Resume
End Function
Private Function Get���SQL(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim strSql As String
    
    strSql = "Select C.סԺ��,C.��ǰ����id,A.��Ժ���� ,c.סԺ��,to_char(A.ȷ������,'yyyyMMdd') as ȷ������,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����ʱ��," & _
        " to_char(A.��Ժ����,'yyyyMMdd') ��Ժ����, A.��Ժ��ʽ,a.��Ժ���� ,a.��Ժ����,H.���� as ��Ժ����,G.��Ժ��� " & _
        " From ������ҳ A,���ű� B,������Ϣ C,���ű� H, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����||'-'||b.����,'')) AS ��Ժ��� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1  and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by ����id,��ҳid)   D," & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����||'-'||b.����,'')) AS ��Ժ��� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� = 3 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by ����id,��ҳid)   G" & _
        " Where A.����id=C.����id and C.����id=" & lng����ID & _
        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID and A.��Ժ����ID=H.id(+) " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        "       and A.��ҳid=G.��ҳid(+) and a.����id=G.����id(+) " & _
        ""
    Get���SQL = strSql
End Function
Public Function סԺ����_�˰�(lng����ID As Long, ByVal lng����ID As Long) As Boolean

    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim lng��ҳID As Long
    Dim strInput As String
    Dim strOutput As String
    Dim strArr
    Dim str���㷽ʽ As String
    
    סԺ����_�˰� = False
        
    DebugTool "����סԺ����ӿ�"
    
    
    gstrSQL = Get��ϸ��¼(lng����ID)
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ������ϸ��¼"
    
    '��ȡ��Ժ���������Ϣ
    lng��ҳID = Nvl(rs��ϸ!��ҳID, 0)
    gstrSQL = Get���SQL(lng����ID, Nvl(rs��ϸ!��ҳID, 0))
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�����Ϣ"
    
    Err = 0
    On Error GoTo ErrHand
    If ҵ������_�˰�(סԺ����ʼ, InitInfor_�˰�.ҽԺ����, "") = False Then Exit Function
    If �ϴ�������ϸ(rs��ϸ) = False Then
        If ҵ������_�˰�(סԺ����ع�, "", "") = False Then Exit Function
        Exit Function
    End If
    
    If ҵ������_�˰�(סԺ�����ύ, "", "") = False Then
        Exit Function
    End If
     
    
    'סԺ����
    strInput = g�������_�˰�.סԺ�ǼǺ� & vbTab
    strInput = strInput & InitInfor_�˰�.ҽԺ���� & vbTab
    strInput = strInput & Format(rsTemp!��Ժ����, "yyyymmdd") & vbTab
    strInput = strInput & Nvl(rsTemp!��Ժ���) & vbTab
    strInput = strInput & Nvl(rsTemp!��Ժ����) & vbTab
    
    gstrSQL = "Select ����Ա���� From ���˽��ʼ�¼ where  ID=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������"
    strInput = strInput & Nvl(rsTemp!����Ա����) & vbTab
        
    If ҵ������_�˰�(סԺ����, strInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, vbTab)
    
    str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & Val(strArr(1))
    str���㷽ʽ = str���㷽ʽ & "||��ͳ��|" & Val(strArr(2))
    'str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & Val(strArr(0))
    
    
    str���㷽ʽ = Mid(str���㷽ʽ, 3)
    
    '������صĽ�����Ϣ
    gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
    Call ExecuteProcedure("����Ԥ����¼")
    Dim intMouse As Integer
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    '��ʾ�������
    If frm������Ϣ.ShowMe(lng����ID) = False Then
        Exit Function
    End If
    Screen.MousePointer = intMouse
  
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN(���θ����ʻ����),�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(�����Ը��ν��),�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN(����ͳ���ʽ��),���Ը����_IN(�󲡼��ʽ��),�����Ը����_IN(�����Ը����),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(������ˮ��,סԺ��סԺ�ǼǺ�),��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert( 2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      "NUll," & g�������_�˰�.�ʻ���� & ",Null,NULL," & g�������_�˰�.����ڼ���סԺ & ",null,Null,NULL," & _
     g�������_�˰�.�����ܶ� & "," & 0 & ",Null," & _
     "Null," & Val(strArr(1)) & "," & Val(strArr(2)) & "," & 0 & "," & Val(strArr(0)) & ",'" & _
     g�������_�˰�.סԺ�ǼǺ� & "'," & lng��ҳID & ",Null,NULl)"
    Call ExecuteProcedure("���汣�ս����¼")
    
    DebugTool "סԺ����ɹ�"
    סԺ����_�˰� = True
    Exit Function
ErrHand:
    
    DebugTool "סԺ����(סԺ����_�˰�)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
End Function

Public Function סԺ�������_�˰�(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    
   

    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    
    Dim strOutput As String
    Dim strInput As String
    
    סԺ�������_�˰� = False
    
    Err = 0
    On Error GoTo ErrHand
    DebugTool "����סԺ�������"
    
    gstrSQL = "Select * From ���ս����¼  where ��¼id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ�ǼǺ�"
    
    lng����ID = Nvl(rsTemp!����ID, 0)
    lng��ҳID = Nvl(rsTemp!��ҳID, 0)
    
    lng����ID = Get����ID(lng����ID)
       
    strInput = Nvl(rsTemp!֧��˳���) & vbTab
    strInput = strInput & InitInfor_�˰�.ҽԺ���� & vbTab
    
    DebugTool "����סԺ����ȡ������"
    '����ȡ���������
    If ҵ������_�˰�(סԺ����ȡ��, strInput, strOutput) = False Then Exit Function
    
    DebugTool "���뱣�汣�ս����¼"
   
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN(���θ����ʻ����),�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(�����Ը��ν��),�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN(����ͳ���ʽ��),���Ը����_IN(�󲡼��ʽ��),�����Ը����_IN(�����Ը����),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert( 1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      "NUll," & -1 * Nvl(rsTemp!�ʻ��ۼ�֧��, 0) & ",Null,NULL," & Nvl(rsTemp!סԺ����, 1) & ",null,Null,NULL," & _
     -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & ",Null," & _
     "Null," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & "," & -1 * Nvl(rsTemp!���Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & _
      Nvl(rsTemp!֧��˳���) & "'," & lng��ҳID & ",Null,NULl)"
    Call ExecuteProcedure("���汣�ս����¼")
    DebugTool "סԺ��������ɹ�"
    סԺ�������_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    סԺ�������_�˰� = False
End Function

Public Function �����Ǽ�_�˰�(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng����ID As Long
    Dim rs��ϸ As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand:
    
    �����Ǽ�_�˰� = False
    
    
    '��һ��: ��ȡ������ϸ��¼
    gstrSQL = Get��ϸ��¼(0, str���ݺ�, lng��¼����, lng��¼״̬)
    Call OpenRecordset(rs��ϸ, "��ȡ������ϸ")
    
    If rs��ϸ.RecordCount = 0 Then
        ShowMsgbox "û����ϸ��¼!"
        Exit Function
    End If
    '���:ҽԺ����
    If ҵ������_�˰�(סԺ����ʼ, InitInfor_�˰�.ҽԺ����, "") = False Then Exit Function
    
    If �ϴ�������ϸ(rs��ϸ) = False Then
        If ҵ������_�˰�(סԺ����ع�, "", "") = False Then Exit Function
        Exit Function
    End If
    
    If ҵ������_�˰�(סԺ�����ύ, "", "") = False Then
        Exit Function
    End If
    �����Ǽ�_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If ҵ������_�˰�(סԺ����ع�, "", "") = False Then Exit Function
End Function
Public Sub WriteParaInfor_�˰�(ByVal strInfor As String)
        '��������Ϣд���ļ���
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        Dim strFile As String
        
        strFile = "C:\�ӿڽ�����.log"
        If Not Dir(strFile) <> "" Then
            objFile.CreateTextFile strFile
        End If
        Set objText = objFile.OpenTextFile(strFile, ForWriting)
        objText.WriteLine strInfor
        objText.Close
End Sub

Public Function ҵ������_�˰�(ByVal intҵ������ As ҵ������_�˰�, ByVal strInputString As String, ByRef strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInput As String, intReturn As Integer, strReturn As String
    Dim strOutput(0 To 10) As String, dblOutPut(0 To 10) As Double, intOutPut(0 To 5) As Integer
    Dim strArr1
    Dim strArr(0 To 20) As String
    Dim strReg As String
    
    Dim i As Integer
    
    DebugTool "����ҵ��������(ҵ������Ϊ:" & intҵ������ & "),�������Ϊ" & vbCrLf & strInputString
    
    
    ҵ������_�˰� = False
    
    strInput = strInputString
    
    If InitInfor_�˰�.ģ������ Then
        '��ȡģ������
        Readģ������ intҵ������, strInputString, strOutPutstring
         ҵ������_�˰� = True
        Exit Function
    End If
    
    
    Err = 0
    On Error GoTo ErrHand:
    
    strReg = "Z|" & intҵ������
    WriteParaInfor_�˰� strInputString
     SaveSetting "ZLSOFT", "ҽ��", "���ݽ���", strReg
    If ����ȴ�() = False Then Exit Function
    strReg = GetSetting("ZLSOFT", "ҽ��", "���ݽ���", "|")
    If strReg = "" Then strReg = "|"
    strArr1 = Split(strReg, "|")
    
    intReturn = Val(strArr1(1))
    Select Case intҵ������
        Case �������ݿ�����
            If intReturn < -99999 Then
                ShowMsgbox "ҽ����������ʧ��!����ҽ���ӿ�����ϵ!"
                Exit Function
            ElseIf intReturn < 0 Then:
                ShowMsgbox "��ҽ�����ĵ����ݿ�����ʧ��!": Exit Function
            End If
        Case �ر����ݿ�����
            If intReturn < 0 Then: ShowMsgbox "����ҽ�����ĵ����ݿ�����ʧ��!": Exit Function
        Case ����Աע��
            '���:yycode(ҽԺ����),czycode(����Ա����),czyname(����Ա����),jylx(�������� 1-��ͨҽ�������ְ�����ݡ����Ϲأ�,2-����ҽ��������ݡ����ҡ���������
            '����:lsh(������ˮ��)
            'intReturn = gobj�˰�.Operator_Login(strArr(0), strArr(1), strArr(2), strArr(3), strOutput(0))
            '��������
            If intReturn = -4 Then
                ShowMsgbox "��Ч�Ĳ���Ա!": Exit Function
            ElseIf intReturn = -5 Then
                ShowMsgbox "���Ǳ�ҽԺ�Ĳ���Ա!": Exit Function
            ElseIf intReturn = -99 Then
                ShowMsgbox "ҽ�������ڲ�����,����ҽ���ṩ����ϵ!": Exit Function
            End If
            Read���� strReturn
        Case ��ȡ������Ϣ
            Select Case intReturn
            Case 0  '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݼ����ӣ�": Exit Function
            Case -3: ShowMsgbox "��Ч�Ŀ���,���鿨���Ƿ���ȷ��": Exit Function
            Case -4: ShowMsgbox "δ��������,����ʹ�ã�": Exit Function
            Case -5: ShowMsgbox "�ÿ��Ѿ�ֹͣʹ�ã�": Exit Function
            Case -41: ShowMsgbox "δ����סԺ�Ǽ�": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������
            Read���� strReturn
        Case ��ȡҽ����Ŀ��Ϣ
            '���:yyfycode(ҽԺ���ô���)
            '����:yplb(ҽ����Ŀ���),ybbz(ҽ����Ŀ��־),ybdm(ҽ����Ŀ����)
            Select Case intReturn
            Case Is >= 0 '�����ɹ�
            Case -6: ShowMsgbox "��Ч����Ŀ���룡": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������
            Read���� strReturn
        Case ����Ԥ����
            '���:xh(���),je(���),fylb(�������),ypbz(ҽ�����),ypcode(ҩƷ����),cfts(��������)
            '����:bczfdje(�����Ը��ν��),bctcje(����ͳ���ʽ��),bczfje(�����Ը����),bczhje(���θ����ʻ����),grzhye(���θ����ʻ�����),bz(��ע)
            Select Case intReturn
            Case 0  'ʹ�ø����ʻ�
            Case 1  'ʹ���Ը���
            Case 2  'ʹ��ͳ��
            Case 3  '��ͳ������
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -97: ShowMsgbox "�ϴι�ҩδ���꣡": Exit Function
            Case -98: ShowMsgbox "�����ʻ������ʹ�����Է�ҩƷ��"
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            Read���� strReturn
            strArr1 = Split(strReturn, vbTab)
            If Trim(strArr1(6)) <> "" Then
                ShowMsgbox strArr1(6)
            End If
            
        Case ������ϸд��
            Select Case intReturn
            Case Is >= 0 '�����ɹ�,�����ʻ����
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -21: ShowMsgbox "����ͨ��ϸ��¼ʧ�ܣ�": Exit Function
            Case -31: ShowMsgbox "������ҽ����ϸ��¼ʧ�ܣ�": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            Read���� strReturn
        Case ��������ύ
            
            '���:��
            '����:mzcode(������ˮ��),grzhye(�ʻ����),bczhje(���ν��׽��),xjzf(�����ֽ��Ը���)
            
            Select Case intReturn
            Case Is >= 0  '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -22: ShowMsgbox "δд��������ϸ��¼��": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            Read���� strReturn
            
        Case ����������
            '���:mzcode(������ˮ��)
            '����:��
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -23: ShowMsgbox "û�д˱ʽ��ף�": Exit Function
            Case -24: ShowMsgbox "�˱ʽ�����ȡ����": Exit Function
            Case -25: ShowMsgbox "�˱ʽ����ѽ��㣡": Exit Function
            Case -26: ShowMsgbox "��������": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            Read���� strReturn
        Case סԺ�Ǽ�
            '���:lsh(סԺ�ǼǺ�),yycode(ҽԺ����),ryrq(��Ժ����),zyh(סԺ��),kbname(��������),ysname(ҽ������),cwcode(��λ��),ryzd(��Ժ���),zt(����״̬(�Ǽ� �޸�)
            '����:��
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -41: ShowMsgbox "δ����סԺ�Ǽ�������": Exit Function
            Case -44: ShowMsgbox "δ����סԺ�Ǽǣ�": Exit Function
            Case -42: ShowMsgbox "��Ч�Ĳ���״̬��": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            strReturn = ""
        Case ȡ��סԺ�Ǽ�
            '���:    lsh(סԺ�ǼǺ�),yycode(ҽԺ����)
            '����:
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -41: ShowMsgbox "δ����סԺ�Ǽ�������": Exit Function
            Case -43: ShowMsgbox "�м��ʷ��ã�����ȡ����": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            strReturn = ""
        Case סԺ��ϸд��
            '���:lsh(סԺ�ǼǺ�),yycode(ҽԺ����),rq(��������),kbname(ҽ������),ysname(��������),fycode(���ô���),fyname(��������),gg(���),dw(��λ),dj(����),sl(����),je(���),fylb(�������),ypbz(ҽ�����),ybdm(ҽ������),czyname(������)
            '����:bl(�����Ը�����),xh(��Ŀ���)
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -44: ShowMsgbox "δ����סԺ�Ǽǣ�": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            Read���� strReturn
        Case סԺ����
            '���:lsh(סԺ�ǼǺ�),yycode(ҽԺ����),cyrq(��Ժ����),Cyzd(��Ժ���),kbname(��Ժ����),czyname(������)
            '����:xjzf(����Ӧ���ֽ�),tcjzje(ͳ����ʽ��),dbjzje(�󲡼��ʽ��)
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -45: ShowMsgbox "δ����סԺ���㣡": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            Read���� strReturn
        Case סԺ����ȡ��
            '���:lsh(סԺ�ǼǺ�),yycode(ҽԺ����)
            '����:��
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -45: ShowMsgbox "δ����סԺ���㣡": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            strReturn = ""
        Case סԺ��ϸȡ��
            '���:lsh(סԺ�ǼǺ�),yycode(ҽԺ����),czyname(������),xh(��Ŀ���)

            '����:
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -2: ShowMsgbox "û�����ݿ�����,����ҽ�����ݿ��Ƿ��Ѿ����Ӻã�": Exit Function
            Case -44: ShowMsgbox "δ����סԺ�Ǽǣ�": Exit Function
            Case -46: ShowMsgbox "û����Ӧ�ķ��ã�": Exit Function
            Case Else: ShowMsgbox "�ڵ�ȡ�˰�ҽ���ӿ�ʱ�����ڲ�����,����ӿڹ�Ӧ����ϵ��": Exit Function '-99
            End Select
            '��������,�ڴ���������������˸�����ֵ
            Read���� strReturn
        Case סԺ����ʼ
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -1: ShowMsgbox "סԺ����ʼʧ�ܣ�": Exit Function
            End Select
            '��������,�ڴ���������������˸�����ֵ
        Case סԺ�����ύ
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -1: ShowMsgbox "סԺ�����ύʧ�ܣ�": Exit Function
            End Select
        Case סԺ����ع�
            Select Case intReturn
            Case 0   '�����ɹ�
            Case -1: ShowMsgbox "סԺ����ع�ʧ�ܣ�": Exit Function
            End Select
    End Select
    strOutPutstring = strReturn
    ҵ������_�˰� = True
    DebugTool "ҵ������ɹ�(ҵ������Ϊ:" & intҵ������ & ")." & vbCrLf & "�������Ϊ" & vbCrLf & strInputString & vbCrLf & "�������Ϊ:" & vbCrLf & strReturn
    Exit Function
ErrHand:
    DebugTool "ҵ������ʧ��(ҵ������Ϊ:" & intҵ������ & ")." & vbCrLf & "�������Ϊ" & vbCrLf & strInputString & vbCrLf & "�������Ϊ:" & vbCrLf & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Readģ������(ByVal intҵ������ As ҵ������_�˰�, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,�Ա����
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim strName As String
    
    strFile = App.Path & "\ģ���ύ��.txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Select Case intҵ������
    Case �������ݿ�����
        Exit Function
    Case �ر����ݿ�����
        Exit Function
    Case ����Աע��
        strName = "����Աע��"
    Case ��ȡ������Ϣ
        strName = "��ȡ������Ϣ"
    Case ��ȡҽ����Ŀ��Ϣ
        strName = "��ȡҽ����Ŀ��Ϣ"
    Case ����Ԥ����
        strName = "����Ԥ����"
    Case ������ϸд��
        strName = "������ϸд��"
    Case ��������ύ
        strName = "��������ύ"
    Case ����������
        strName = "����������"
    Case סԺ�Ǽ�
        strName = "סԺ�Ǽ�"
    Case ȡ��סԺ�Ǽ�
        Exit Function
    Case סԺ��ϸд��
        Exit Function
    Case סԺ��ϸȡ��
        Exit Function
    Case סԺ����
        strName = "סԺ����"
    Case סԺ����ȡ��
        strName = "סԺ����ȡ��"
    End Select
   
    Dim blnStart As Boolean
    Dim strArr
    Err = 0
    On Error GoTo ErrHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                    
                If blnStart Then
                    If strText = "" Then
                        strText = "" & vbTab
                    End If
                    strArr = Split(strText, "|")
                    
                    If Val(strArr(0)) = 1 Then
                        str = strArr(1)
                        Exit Do
                    End If
                Else
                     If "<" & strName & ">" = strText Then
                         blnStart = True
                     End If
                End If
                If "</" & strName & ">" = strText Then
                    Exit Do
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
    Exit Function
ErrHand:
    DebugTool Err.Description
    Exit Function
End Function
Public Function �ҺŽ���_�˰�(ByVal lng����ID As Long) As Boolean
  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    
    �ҺŽ���_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function �Һų���_�˰�(ByVal lng����ID As Long) As Boolean

    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Err = 0
    On Error GoTo ErrHand
    
    �Һų���_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Sub WriteDebugInfor_�˰�(ByVal strInfor As String)
        '��������Ϣд���ļ���
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        Dim intDebug As Integer
        
        intDebug = GetSetting("ZLSOFT", "ҽ��", "����д���ı��ļ�", 0)
        If intDebug <> 1 Then Exit Sub

        Dim strFile As String
        Dim rsTemp As New ADODB.Recordset
        strFile = App.Path & "\Test.log"
        
        If Not Dir(strFile) <> "" Then
            objFile.CreateTextFile strFile
        End If
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfor
        objText.Close
End Sub
Public Function ��Ժ������Ϣ_�˰�(lng����ID As Long, lng��ҳID As Long) As Boolean
 
    
    ��Ժ������Ϣ_�˰� = False
    On Error GoTo ErrHand
    DebugTool "������Ժ������Ϣ�ӿ�"
    
    If סԺ��Ϣ�ύ(lng����ID, lng��ҳID, True) = False Then Exit Function
    
    DebugTool "��Ժ������Ϣ�޸ĳɹ�"
    ��Ժ������Ϣ_�˰� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function סԺ��Ϣ�ύ(lng����ID As Long, lng��ҳID As Long, Optional bln�޸� As Boolean = False) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    
    'дסԺ��Ϣ
    DebugTool "��ȡ��Ժ�������Ϣ"
    Err = 0
    On Error GoTo ErrHand:
    סԺ��Ϣ�ύ = False
    
    '��ȡ��ز�����Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����,to_char(A.ȷ������,'yyyyMMdd') as ȷ������,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����ʱ��," & _
        " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����  ,to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժʱ��,D.��Ժ��� " & _
        " From ������ҳ A,���ű� B,������Ϣ C, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����||'-'||b.����,'')) AS ��Ժ��� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by  ����id,��ҳid)   D" & _
        " Where A.����id=C.����id and C.����id=" & lng����ID & _
        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        ""
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ��Ϣ"
    
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    '���:סԺ�ǼǺ�,ҽԺ����,��Ժ����,סԺ��,��������,ҽ������,��λ��,��Ժ���,����״̬
    strInput = g�������_�˰�.סԺ�ǼǺ� & vbTab
    strInput = strInput & InitInfor_�˰�.ҽԺ���� & vbTab
    strInput = strInput & Nvl(rsTemp!��Ժ����) & vbTab
    strInput = strInput & Nvl(rsTemp!סԺ��) & vbTab
    strInput = strInput & Nvl(rsTemp!��Ժ����) & vbTab
    strInput = strInput & Nvl(rsTemp!סԺҽʦ) & vbTab
    strInput = strInput & Nvl(rsTemp!��ǰ����) & vbTab
    strInput = strInput & Nvl(rsTemp!��Ժ���) & vbTab
    strInput = strInput & IIf(bln�޸�, "�޸�", "�Ǽ�")
    
    DebugTool "����סԺ�޸�����"
    
    If ҵ������_�˰�(סԺ�Ǽ�, strInput, strOutput) = False Then
        Exit Function
    End If
    DebugTool "��Ժ������Ϣд��ɹ�"
    
    סԺ��Ϣ�ύ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetItemInfo_�˰�(ByVal lngPatiID As Long, ByVal lngItemID As Long, Optional ByVal strժҪ As String, Optional intType As Integer = 0) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�������˵������ʾ��Ϣ
    '--�����:
    '--������:
    '--��  ��:��ʾ��
    '-----------------------------------------------------------------------------------------------------------
    Dim strMsgInfor As String
    Dim strԭժҪ As String
    strԭժҪ = strժҪ
    If g�������_�˰�.�������� = "��ͨҽ������" Then
        GetItemInfo_�˰� = strԭժҪ
        Exit Function
    End If
    strMsgInfor = strժҪ
    If frm������Ϣ����_�˰�.EditCard(strMsgInfor) = False Then
        GetItemInfo_�˰� = strԭժҪ
        Exit Function
    End If
    GetItemInfo_�˰� = strMsgInfor
End Function
Private Function Read����(ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,�Ա����
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim strName As String
    
    strFile = "C:\�ӿڽ�����.log"
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Err = 0
    On Error GoTo ErrHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                Exit Do
            Loop
            objText.Close
            strOutPutstring = strText
    End If
    Exit Function
ErrHand:
    Exit Function
End Function

Private Function ����ȴ�() As Boolean
    '�ȴ����ݴ���,true����ɹ�,fale����ʧ��
    Dim strReg As String
    Dim strArr1
    ����ȴ� = False
    
    Dim strDate As String
    strDate = Format(DateAdd("s", 400, Now), "yyyymmdd HH:MM:SS")
    Do While True
        '�ȴ��������
        strReg = GetSetting("ZLSOFT", "ҽ��", "���ݽ���", "|")
        strArr1 = Split(strReg, "|")
        If strArr1(0) = "H" Then
            Exit Do
        End If
        
        If strArr1(0) = "" Then
            Exit Function
        End If
        If Format(Now, "yyyymmdd HH:MM:SS") >= strDate Then
            '���׵ȴ���������ȡ�����ν���,
            ShowMsgbox "���׵ȴ���������ȡ�����ν���"
            Exit Function
        End If
    Loop
    ����ȴ� = True
End Function


