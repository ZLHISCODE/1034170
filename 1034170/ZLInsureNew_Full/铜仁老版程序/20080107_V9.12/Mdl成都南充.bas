Attribute VB_Name = "Mdl�ɶ��ϳ�"
Option Explicit
Public Const gstrSplit���� As String = "��"
Public Const gstrSplitС�� As String = "��"
Public Const gstr������Ŀ As String = "��λ�ѡ���ҩ�ѡ��г�ҩ�ѡ��в�ҩ�ѡ������ѡữ��ѡ���ѡ����Ʒѡ���Ϸѡ�Ѫ�ѡ������ѡ����ٷѡỤ��ѡ����ѡ�CT�ѡ�˴Ź����������"
Public Const gstrCol_ENG As String = "BH,ID,ZWMC,JLDW,DJ,YPXM,YPLX,YPZLX,YPXLX,YPSHQF,XZSYFW,YPXMLH"
Public Const gstrCol_CHI As String = "���,ҽ����ĿID,��������,������λ,����,ҩƷ��Ŀ,ҩƷ����,ҩƷ������,ҩƷС����,ҩƷʹ������,ҩƷʹ�÷�Χ,ҩƷ��Ŀ�ں�"
Public gcnInterbase As New ADODB.Connection

Private rsTemp As New ADODB.Recordset
Private Const giniPath As String = "c:\his_yb"
Private Const giniFile As String = "his_yb.ini"
Private strSql As String
Private strProcedure As String
Private intReturn As Integer

Type Bill_Head
    סԺ�� As String
    ������ˮ�� As String
    ����ʱ�� As Date
    ҽ�� As String
    ���� As String
    ���� As String
End Type
Type Bill_Body
    ������ϸ��ˮ�� As Long
    ҽ���շ�ϸĿ As Long
    ���� As Currency
    ���� As Single
    ������Ŀ As String
    ������λ As String
    '--���²�������ҩƷ��Ч������Ϊ��
    ��Ʒ�� As String
    ��� As String
    ���� As String  '�գ����࣬����
    ���� As Currency
End Type
Private ����ͷ As Bill_Head
Private ������ As Bill_Body

Public Function ҽ������_�ɶ��ϳ�() As Boolean
'���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
'���أ��ӿ����óɹ�������true�����򣬷���false
    Dim strConn As String
    
    On Error GoTo ErrHand
    If frmSet�ɶ�.ShowSet(TYPE_�ɶ��ϳ�) = False Then
        Exit Function
    End If
    
    strConn = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
    '���½�����ҽ���������Ĺ�������
    If gcnInterbase.State = adStateClosed Then
        On Error Resume Next
        gcnInterbase.Open strConn
        If Err = 0 Then
            ҽ������_�ɶ��ϳ� = True
        Else
            Err.Clear
        End If
    Else
        ҽ������_�ɶ��ϳ� = True
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ҽ����ʼ��_�ɶ��ϳ�() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    On Error GoTo ErrHand
    '������ҽ���������Ĺ�������
    Dim strConn As String
    strConn = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
    Err = 0
    On Error Resume Next
    With gcnInterbase
        If .State = adStateOpen Then .Close
        .ConnectionString = strConn
        .Open
        If Err <> 0 Then
            MsgBox "���ܽ�����ҽ�������������ӣ��޷�ִ��ҽ������", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    
    ҽ����ʼ��_�ɶ��ϳ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��ݱ�ʶ_�ɶ��ϳ�(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    On Error GoTo ErrHand
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify�ɶ��ϳ�.ShowCard(bytType, lng����ID)
    ��ݱ�ʶ_�ɶ��ϳ� = strTmpIden
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����Ǽ�_�ɶ��ϳ�_ָ������(ByVal lngPatient As Long) As Boolean
    '�ϴ�ʣ�µķ��ã���Ҫ�Ǵ�λ�ѡ�����ѵȣ�
    '��д�뵥��ͷ����д�뵥����
    '��¼״̬��1-����;����Ϊɾ���������ô�����ֻ�����ŵ���ɾ�����ٲ����µ���
    On Error GoTo ErrHand
    �����Ǽ�_�ɶ��ϳ�_ָ������ = False
    ����ͷ.סԺ�� = ""
    
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        gstrSQL = " Select A.ID ������ϸ��ˮ��,A.��ʶ�� as סԺ��,A.����ID,A.��¼����,A.��¼״̬,A.NO,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ����ʱ��," & _
                  " A.������ ҽ��,'' ����,B.���� ����,C.��Ŀ���� ҽ���շ�ϸĿ,A.��׼���� ����,A.����*Nvl(A.����,1) ����,D.��� ������Ŀ,'['||E.����||']'||E.���� �շ�ϸĿ," & _
                  " E.��� �շ�ϸĿ���,substrb(E.���,1,60) ���,E.�������� ����,A.��׼���� ����" & _
                  " From ���˷��ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=" & TYPE_�ɶ��ϳ� & ") C,�շ���� D,�շ�ϸĿ E " & _
                  " Where A.ִ�в���ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And A.�շ����=D.���� And A.�շ�ϸĿID=E.ID And Nvl(A.�Ƿ��ϴ�,0)=0" & _
                  " And A.����ID=" & lngPatient & _
                  " Order by A.��ʶ��"
        Call OpenRecordset(rsTemp, "�����Ǽ�")
    End With
    
    If Not �ϴ�_�����Ǽ�(rsTemp) Then Exit Function
    
    �����Ǽ�_�ɶ��ϳ�_ָ������ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����Ǽ�_�ɶ��ϳ�(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '��д�뵥��ͷ����д�뵥����
    '��¼״̬��1-����;����Ϊɾ���������ô�����ֻ�����ŵ���ɾ�����ٲ����µ���
    On Error GoTo ErrHand
    �����Ǽ�_�ɶ��ϳ� = False
    ����ͷ.סԺ�� = ""
    
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        gstrSQL = " Select A.ID ������ϸ��ˮ��,A.��ʶ�� as סԺ��,A.����ID,A.��¼����,A.��¼״̬,A.NO,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ����ʱ��," & _
                  " A.������ ҽ��,'' ����,B.���� ����,C.��Ŀ���� ҽ���շ�ϸĿ,A.��׼���� ����,A.����*Nvl(A.����,1) ����,D.��� ������Ŀ,'['||E.����||']'||E.���� �շ�ϸĿ," & _
                  " E.��� �շ�ϸĿ���,substrb(E.���,1,60) ���,E.�������� ����,A.��׼���� ����" & _
                  " From ���˷��ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=" & TYPE_�ɶ��ϳ� & ") C,�շ���� D,�շ�ϸĿ E " & _
                  " Where A.��¼����=" & lng��¼���� & " And A.��¼״̬=" & lng��¼״̬ & " And A.NO='" & str���ݺ� & "'" & _
                  " And A.ִ�в���ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And A.�շ����=D.���� And A.�շ�ϸĿID=E.ID And Nvl(A.�Ƿ��ϴ�,0)=0" & _
                  " Order by A.��ʶ��"
        Call OpenRecordset(rsTemp, "�����Ǽ�")
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�������¼����ҽ����������������ʧ�ܣ�[�����Ǽ�]", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If Not �ϴ�_�����Ǽ�(rsTemp) Then Exit Function
    
    �����Ǽ�_�ɶ��ϳ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function Get��ˮ��(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, strNO)
    Get��ˮ�� = lng��¼���� & lng��¼״̬ & (Asc(Mid(strNO, 1, 1)) - 55) & Mid(strNO, 2)
End Function

Public Function ����ɾ��_�ɶ��ϳ�(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    Dim blnNew As Boolean
    Dim blnExec As Boolean 'Modified.By.ZYB 2003-01-23 ��ҽ�����˲�ִ���ϴ�
    '��д�뵥��ͷ����д�뵥����
    '��¼״̬��1-����;����Ϊɾ���������ô�����ֻ�����ŵ���ɾ�����ٲ����µ���
    On Error GoTo ErrHand
    ����ɾ��_�ɶ��ϳ� = False
    ����ͷ.סԺ�� = ""
    
    gcnInterbase.BeginTrans
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        gstrSQL = " Select A.ID ������ϸ��ˮ��,A.��ʶ�� as סԺ��,A.����ID,A.��¼����,A.��¼״̬,A.NO,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ����ʱ��," & _
                  " A.������ ҽ��,'' ����,B.���� ����,C.��Ŀ���� ҽ���շ�ϸĿ,A.��׼���� ����,A.����*Nvl(A.����,1) ����,D.��� ������Ŀ,'['||E.����||']'||E.���� �շ�ϸĿ" & _
                  " From ���˷��ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=" & TYPE_�ɶ��ϳ� & ") C,�շ���� D,�շ�ϸĿ E " & _
                  " Where A.��¼����=" & lng��¼���� & " And A.��¼״̬=" & lng��¼״̬ & " And A.NO='" & str���ݺ� & "'" & _
                  " And A.ִ�в���ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And A.�շ����=D.���� And A.�շ�ϸĿID=E.ID And Nvl(A.�Ƿ��ϴ�,0)=0" & _
                  " Order by A.��ʶ��"
        Call OpenRecordset(rsTemp, "����ɾ��")
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�������¼����ҽ����������������ʧ�ܣ�[����ɾ��]", vbInformation, gstrSysName
            gcnInterbase.RollbackTrans
            Exit Function
        End If
        
        Do While Not .EOF
            'д�봦��ͷ
            blnNew = (����ͷ.סԺ�� <> GetסԺ��(rsTemp!����ID))
            blnExec = IsYBPatient(rsTemp!����ID)
            If blnNew And blnExec Then
                With ����ͷ
                    .סԺ�� = GetסԺ��(rsTemp!����ID)
                    .������ˮ�� = Get��ˮ��(lng��¼����, 1, str���ݺ�)
                    .����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                    .ҽ�� = IIf(IsNull(rsTemp!ҽ��), "", rsTemp!ҽ��)
                    .���� = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .���� = rsTemp!����
                End With
                
                strProcedure = "DELETE_CFJLK"
                strSql = "Execute Procedure DELETE_CFJLK('" & ����ͷ.סԺ�� & "'," & ����ͷ.������ˮ�� & ")"
                If Not ExecProc(strSql) Then gcnInterbase.RollbackTrans: Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    '���²��˷��ü�¼�����ϴ���־Ϊ��
    If Not �����ϴ���־(rsTemp) Then
        gcnInterbase.RollbackTrans
        Exit Function
    End If
    
    gcnInterbase.CommitTrans
    ����ɾ��_�ɶ��ϳ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnInterbase.RollbackTrans
End Function

Public Function �ϴ�_�����Ǽ�(ByVal rsTemp As ADODB.Recordset) As Boolean
    Dim blnNew As Boolean
    Dim blnExec As Boolean 'Modified.By.ZYB 2003-01-23 ��ҽ�����˲�ִ���ϴ�
    On Error GoTo ErrHand
    
    �ϴ�_�����Ǽ� = False
    gcnInterbase.BeginTrans
    
    With rsTemp
        Do While Not .EOF
            'д�봦��ͷ
            blnNew = (����ͷ.סԺ�� <> GetסԺ��(rsTemp!����ID))
            blnExec = IsYBPatient(rsTemp!����ID)
            If blnExec Then
                If blnNew Then
                    With ����ͷ
                        .סԺ�� = GetסԺ��(rsTemp!����ID)
                        .������ˮ�� = Get��ˮ��(rsTemp!��¼����, rsTemp!��¼״̬, rsTemp!no)
                        .����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                        .ҽ�� = IIf(IsNull(rsTemp!ҽ��), "", rsTemp!ҽ��)
                        .���� = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                        .���� = rsTemp!����
                    End With
                    
                    strProcedure = "ADD_CFJLK"
                    strSql = "Execute Procedure ADD_CFJLK('" & ����ͷ.סԺ�� & "'," & ����ͷ.������ˮ�� & _
                             ",'" & ����ͷ.����ʱ�� & "','" & ����ͷ.ҽ�� & "',NULL,'" & ����ͷ.���� & "')"
                    If Not ExecProc(strSql) Then gcnInterbase.RollbackTrans: Exit Function
                End If
                
                'д�봦����ϸ
                With ������
                    .������ϸ��ˮ�� = rsTemp!������ϸ��ˮ��
                    .ҽ���շ�ϸĿ = IIf(IsNull(rsTemp!ҽ���շ�ϸĿ), 0, rsTemp!ҽ���շ�ϸĿ)
                    .���� = rsTemp!����
                    .���� = rsTemp!����
                    .������λ = ""
                    .������Ŀ = Get������Ŀ(rsTemp!������Ŀ)
                    If InStr(1, ",5,6,7,", "," & rsTemp!�շ�ϸĿ��� & ",") <> 0 Then
                        .��Ʒ�� = rsTemp!�շ�ϸĿ
                        .��� = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                        .���� = IIf(rsTemp!���� = "����ҩ", "����", IIf(rsTemp!���� = "����ҩ", "����", ""))
                    Else
                        .��Ʒ�� = ""
                        .��� = ""
                        .���� = ""
                    End If
                    .���� = 0
                    
                    If .ҽ���շ�ϸĿ = 0 Then
                        MsgBox rsTemp!�շ�ϸĿ & "δ���ö�Ӧ��ҽ����Ŀ��[�ϴ�����]", vbInformation, gstrSysName
                        gcnInterbase.RollbackTrans
                        Exit Function
                    End If
                    If .������Ŀ = "" Then
                        MsgBox "������Hisϵͳ�е��շ������ҽ��ϵͳ�з�����Ŀ�Ķ��չ�ϵ��[�������]", vbInformation, gstrSysName
                        gcnInterbase.RollbackTrans
                        Exit Function
                    End If
                End With
                
                strProcedure = "ADD_CFMXK"
                strSql = "Execute Procedure ADD_CFMXK('" & ����ͷ.סԺ�� & "'," & ����ͷ.������ˮ�� & _
                        "," & ������.������ϸ��ˮ�� & "," & ������.ҽ���շ�ϸĿ & ",'" & ������.������λ & _
                        "'," & ������.���� & "," & ������.���� & ",'" & ������.������Ŀ & "','" & ������.��Ʒ�� & _
                        "','" & ������.��� & "','" & ������.���� & "'," & ������.���� & ")"
                If Not ExecProc(strSql) Then gcnInterbase.RollbackTrans: Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    '���²��˷��ü�¼�����ϴ���־Ϊ��
    If Not �����ϴ���־(rsTemp) Then
        gcnInterbase.RollbackTrans
        Exit Function
    End If
    
    gcnInterbase.CommitTrans
    �ϴ�_�����Ǽ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnInterbase.RollbackTrans
End Function

Public Function ��Ժ�Ǽ�_�ɶ��ϳ�(ByVal lngPatient As Long) As Boolean
    Dim strObj As String, strסԺ�� As String, blnExist As Boolean
    Dim strLine As TextStream, FileSys As New FileSystemObject
    '����Ժ���˵�סԺ��д�뱾����(c:\his_yb\his_yb.ini)
    '��ʽΪ��zyh=11111111
    'ͬʱ���±����ʻ��Ͳ�����ҳ
    
    On Error GoTo ErrHand
    ��Ժ�Ǽ�_�ɶ��ϳ� = False
    
    '��Ժ�Ǽ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatient & "," & gintInsure & ")"
    Call ExecuteProcedure("�ϳ�ҽ��")
    
    '���·�������ڣ�������
    If Not FileSys.FolderExists(giniPath) Then FileSys.CreateFolder (giniPath)
    '����ļ����ڣ���ɾ�������²���
    blnExist = FileSys.FileExists(giniPath & "\" & giniFile)
    If blnExist Then Call FileSys.DeleteFile(giniPath & "\" & giniFile, True)
    strסԺ�� = GetסԺ��(lngPatient)
    '�����Ƿ���ڸö���
    Set strLine = FileSys.OpenTextFile(giniPath & "\" & giniFile, ForWriting, True)
    
    Call strLine.WriteLine("[String]")  'Modified.By.ZYB 2003-01-23
    Call strLine.WriteLine("ZYH=" & strסԺ��)
    strLine.Close
    
    ��Ժ�Ǽ�_�ɶ��ϳ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_�ɶ��ϳ�(ByVal lngPatient As Long) As Boolean
    On Error GoTo ErrHand
    ��Ժ�Ǽ�_�ɶ��ϳ� = False
    
    '�ϴ�ʣ�µķ��ã���Ҫ�Ǵ�λ�ѡ�����ѵȣ�
    If Not �����Ǽ�_�ɶ��ϳ�_ָ������(lngPatient) Then Exit Function
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatient & "," & gintInsure & ")"
    Call ExecuteProcedure("�ϳ�ҽ��")
    
    ��Ժ�Ǽ�_�ɶ��ϳ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽǳ���_�ɶ��ϳ�(ByVal lngPatient As Long) As Boolean
    On Error GoTo ErrHand
    ��Ժ�Ǽǳ���_�ɶ��ϳ� = False

    '�ָ���Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatient & "," & gintInsure & ")"
    Call ExecuteProcedure("�ϳ�ҽ��")
    
    ��Ժ�Ǽǳ���_�ɶ��ϳ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetסԺ��(ByVal lngPatient As Long) As String
    On Error GoTo ErrHand
    Dim rsTemp As New ADODB.Recordset
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        ''Modified.By.ZYB 2003-01-23 ����ÿ��סԺ����סԺ�Ŷ�����Ψһ�����Լ���סԺ����
        gstrSQL = "Select סԺ��||'_'||סԺ���� סԺ�� From ������Ϣ Where ����ID=" & lngPatient
        Call OpenRecordset(rsTemp, "�ϳ�ҽ��")
        
        GetסԺ�� = !סԺ��
    End With
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function IsYBPatient(ByVal lngPatient As Long) As Boolean
    Dim rsYbPatient As New ADODB.Recordset
    On Error GoTo ErrHand
    '�ж��Ƿ���ҽ������
    IsYBPatient = False
    
    With rsYbPatient
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        gstrSQL = "Select Count(*) Records From �����ʻ� Where ����=" & gintInsure & " And ����ID=" & lngPatient
        Call OpenRecordset(rsYbPatient, "�ϳ�ҽ��")
        
        If .EOF Then Exit Function
        If IsNull(!Records) Then Exit Function
        If !Records = 0 Then Exit Function
    End With
    
    IsYBPatient = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ�������_�ɶ��ϳ�(ByVal lngPatient As Long) As String
    On Error GoTo ErrHand
    Dim rsPay As New ADODB.Recordset, CurPay As Currency, strסԺ�� As String
    סԺ�������_�ɶ��ϳ� = ""
    
    strסԺ�� = GetסԺ��(lngPatient)
    
    strProcedure = "GET_SBBXJE"
    With rsPay
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Execute Procedure GET_SBBXJE('" & strסԺ�� & "')", gcnInterbase
        CurPay = IIf(IsNull(!BXXJ), 0, !BXXJ)
        intReturn = !SUCC
    End With
    If intReturn <> 0 Then
        Call IsError
        סԺ�������_�ɶ��ϳ� = ""
        Exit Function
    End If
    סԺ�������_�ɶ��ϳ� = "ҽ������;" & CurPay & ";0"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ����_�ɶ��ϳ�(ByVal lng����ID As Long, ByVal rstmp As ADODB.Recordset) As Boolean
    Dim CurPay As Currency
    '���벡����ҽ�����ݿ��г�Ժ�����ɵ��ý������GET_SBBXJE
    '��֧�ֽ�����˷Ѳ�����������ҽ�����ĽӴ����
    
    CurPay = Split(סԺ�������_�ɶ��ϳ�(rstmp!����ID), ";")(1)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & rstmp!����ID & "," & _
        Int(Format(zlDatabase.Currentdate, "yyyy")) & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & 0 & "," & 0 & "," & 0 & "," & CurPay & ",0," & _
        0 & "," & 0 & ",NULL," & 0 & "," & 0 & ")"
    Call ExecuteProcedure("�ϳ�ҽ��")
    
    סԺ����_�ɶ��ϳ� = True
End Function

Public Function Get������Ŀ(ByVal str�շ���� As String) As String
    On Error GoTo ErrHand
    Dim str������Ŀ As String, arrItem, intItem As Integer
    '��ȡ���շ�����Ӧ��ҽ��������Ŀ
    Get������Ŀ = ""
    str������Ŀ = Trim(GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LCItem"), ""))
    If str������Ŀ = "" Then Exit Function
    
    arrItem = Split(str������Ŀ, gstrSplit����)
    For intItem = 0 To UBound(arrItem)
        If Split(arrItem(intItem), gstrSplitС��)(0) = str�շ���� Then
            Get������Ŀ = Split(arrItem(intItem), gstrSplitС��)(1)
            Exit For
        End If
    Next
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ExchangeColName(ByVal strCol As String, Optional ByVal blnExchange As Boolean = True) As String
    Dim arrEng, arrChi, arrTemp
    Dim intExchange As Integer, intFind As Integer
    'Ӣ��������������������ת��
    On Error GoTo ErrHand
    
    arrEng = Split(gstrCol_ENG, ",")
    arrChi = Split(gstrCol_CHI, ",")
    If blnExchange Then
        arrTemp = arrEng
    Else
        arrTemp = arrChi
    End If
    
    For intExchange = 0 To UBound(arrTemp)
        If arrTemp(intExchange) = strCol Then
            intFind = intExchange
            Exit For
        End If
    Next
    
    If blnExchange Then
        ExchangeColName = arrChi(intFind)
    Else
        ExchangeColName = arrEng(intFind)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����ϴ���־(ByVal rsTemp As ADODB.Recordset) As Boolean
    On Error GoTo ErrHand
    �����ϴ���־ = False
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & !������ϸ��ˮ�� & ")"
            Call ExecuteProcedure("�ϳ�ҽ��")
            .MoveNext
        Loop
    End With
    �����ϴ���־ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ExecProc(ByVal strExec As String, Optional ByVal bln��ʾ As Boolean = True) As Boolean
    Dim rsExecute As New ADODB.Recordset
    On Error GoTo ErrHand
    With rsExecute
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open strExec, gcnInterbase
        If .RecordCount = 0 Then
            MsgBox "��ҽ���������������ݹ����У�����δ֪����", vbInformation, gstrSysName
            Exit Function
        End If
        intReturn = .Fields(0).Value
    End With
    
    ExecProc = Not IsError(bln��ʾ)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsError(Optional ByVal bln��ʾ As Boolean = True) As Boolean
    On Error GoTo ErrHand
    Dim strMsg As String
    IsError = False
    If intReturn = 0 Then Exit Function
    strProcedure = UCase(strProcedure)
    
    Select Case strProcedure
    Case "ADD_CFJLK"
        Select Case intReturn
        Case 1
            strMsg = "����������סԺ,���ܵǼǴ�����"
        Case 2
            strMsg = "�ô�����¼�룡"
        Case 3
            strMsg = "����ʱ��С����Ժʱ�䣡"
        Case 4
            strMsg = "��������ˣ��������ӣ�"
        End Select
    Case "DELETE_CFJLK"
        Select Case intReturn
        Case 1
            strMsg = "������¼�봦��,����ɾ����"
        Case 2
            strMsg = "��������ˣ�����ɾ����"
        Case 3
            strMsg = "��������ˣ�����ɾ����"
        End Select
    Case "UPDATE_CFJLK"
        Select Case intReturn
        Case 1
            strMsg = "������¼�봦��,�����޸ģ�"
        Case 2
            strMsg = "��������ˣ�����ɾ����"
        Case 3
            strMsg = "��������ˣ�����ɾ����"
        Case 4
            strMsg = "����ʱ��С����Ժʱ�䣡"
        End Select
    Case "ADD_CFMXK"
        Select Case intReturn
        Case 1
            strMsg = "����������סԺ,����¼�봦����"
        Case 2
            strMsg = "�����ȵǼǴ���,���ܵǼǴ�����ϸ��"
        Case 3
            strMsg = "��������ˣ��������ӣ�"
        Case 4
            strMsg = "����ʱ��С����Ժʱ�䣡"
        Case 5
            strMsg = "ҩƷû�ҵ��������ҩƷ��Ϣ�⣡"
        End Select
    Case "DELETE_CFMXK"
        Select Case intReturn
        Case 1
            strMsg = "������¼�봦����ϸ,����ɾ����"
        Case 2
            strMsg = "��������ˣ�����ɾ����"
        Case 3
            strMsg = "��������ˣ�����ɾ����"
        End Select
    Case "UPDATE_CFMXK"
        Select Case intReturn
        Case 1
            strMsg = "������¼�봦����ϸ,�����޸ģ�"
        Case 2
            strMsg = "��������ˣ������޸ģ�"
        Case 3
            strMsg = "��������ˣ������޸ģ�"
        End Select
    Case "GET_SBBXJE"
        Select Case intReturn
        Case 1
            strMsg = "û������סԺ��"
        Case 2
            strMsg = "������ҽ�����ݿ��г�Ժ����ܽ��н��㣡"
        End Select
    End Select
    IsError = True
    If bln��ʾ Then MsgBox strMsg, vbInformation, gstrSysName
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function



