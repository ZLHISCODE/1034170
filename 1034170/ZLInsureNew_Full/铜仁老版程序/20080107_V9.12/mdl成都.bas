Attribute VB_Name = "mdl�ɶ�"
Option Explicit
Public gcnSybase As New ADODB.Connection

Public Function ҽ������_�ɶ�() As Boolean
'���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
'���أ��ӿ����óɹ�������true�����򣬷���false
    Dim strConn As String
    
    If frmSet�ɶ�.ShowSet(TYPE_�ɶ���) = False Then
        Exit Function
    End If
    
    strConn = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ConnectionString"), "dsn=cnnSyb;uID=face;pwd=facepass")
    '���½�����ҽ���������Ĺ�������
    If gcnSybase.State = adStateClosed Then
        On Error Resume Next
        gcnSybase.Open strConn
        If Err = 0 Then
            ҽ������_�ɶ� = True
        Else
            Err.Clear
        End If
    Else
        ҽ������_�ɶ� = True
    End If
End Function

Public Function ҽ����ʼ��_�ɶ�() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false

    '������ҽ���������Ĺ�������
    Dim strConn As String
    strConn = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ConnectionString"), "")
    Err = 0
    On Error Resume Next
    With gcnSybase
        If .State = 1 Then .Close
        '.ConnectionString = "dsn=cnnSyb;uID=face;pwd=facepass"
        .ConnectionString = strConn
        .Open
        If Err <> 0 Then
            MsgBox "���ܽ�����ҽ�������������ӣ��޷�ִ��ҽ������", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    
    ҽ����ʼ��_�ɶ� = True
End Function

Public Function ��ݱ�ʶ_�ɶ�2(ByVal strCard As String, ByVal strPass As String, Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strCard-ˢ���õ���strPass-�������룻bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
'Ȩ�ޣ����ű�_ID,������Ϣ,�����ʻ�,zl_������Ϣ_Insert,zl_������Ϣ_Update,zl_�����ʻ�_insert,zl_�ʻ������Ϣ_Insert
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strInfo As String
    Dim strҽ���� As String, str���� As String
    Dim strSerial As String, strSwapNo As String '����˳���
    Dim cur��� As Currency
    Dim curסԺ���� As Currency, cur�������� As Currency, curסԺ�޶� As Currency
    
    If strCard = "" Then Exit Function
    
    '������ҽ���źͿ���
    Call ExecuteZ015(strCard, strҽ����, str����)
    If strҽ���� = "" And str���� = "" Then
        MsgBox "ˢ������ʧ�ܣ������ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��֤���
    With rsTmp
        If .State = 1 Then .Close
        strSql = "select ���ű�_id.nextval||'1' from dual"
        .CursorLocation = adUseClient
        Call OpenRecordset(rsTmp, "�ɶ�ҽ��", strSql)
        strSwapNo = .Fields(0).Value
        strSerial = getSerial(strҽ����)
        
        'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
        strSql = "z001('z001','" & UserInfo.վ�� & "','" & strSwapNo & "','" & strPass & "','" & UserInfo.��� & "'," & _
            "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & strSwapNo & "','" & IIf(bytType = 0, "11", "31") & "','" & str���� & "')"
        gcnSybase.Execute strSql, , adCmdStoredProc
        
        strSql = "select code from zjycl  where jysxh='" & strSwapNo & "' and jybh='z001'"
        If .State = 1 Then .Close
        .Open strSql, gcnSybase, adOpenStatic, adLockReadOnly
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "����""z001""���ִ���""" & !code & """:" & vbCrLf & String(2, "��") & GetErrInfo(!code) & String(2, vbTab), vbInformation, gstrSysName
            Exit Function
        Else
            strSql = "select * from grjbxx where grbm='" & strҽ���� & "'"
            If .State = 1 Then .Close
            .Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
            If Not .EOF Then
                'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
                strInfo = str���� & ";" & strҽ���� & ";" & strPass & ";" & _
                        TrimStr(.Fields("xm").Value) & ";" & _
                        IIf(TrimStr(.Fields("xb").Value) = "1", "��", "Ů") & ";" & _
                        TrimStr(.Fields("csrq").Value) & ";" & _
                        TrimStr(.Fields("sfz").Value) & ";" & _
                        TrimStr(.Fields("dwmc").Value) & "(" & Trim(.Fields("dwbm").Value) & ")"
                
                cur��� = IIf(IsNull(!grzhlnye), 0, !grzhlnye) + IIf(IsNull(!grzhbnye), 0, !grzhbnye)
                '200308z012
                If bytType <> 0 Then
                    curסԺ���� = IIf(IsNull(!zyjs), 0, !zyjs)
                    cur�������� = IIf(IsNull(!tcbxbl), 0, !tcbxbl)
                    curסԺ�޶� = IIf(IsNull(!zyxe), 0, !zyxe)
                End If
                
                lng����ID = BuildPatiInfo(bytType, strInfo & ";;;;" & cur��� & ";;;;;;;" & _
                    cur��� & ";;;;;;" & curסԺ���� & ";" & cur�������� & ";" & curסԺ�޶�, lng����ID)
                
                '���ظ�ʽ:�м���벡��ID
                ��ݱ�ʶ_�ɶ�2 = strInfo & ";" & lng����ID & ";;;;" & cur��� & ";;;;;;;" & cur��� & ";;;;;"
            End If
        End If
    End With
End Function

Public Function ��ݱ�ʶ_�ɶ�(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify�ɶ�
    Dim strPatiInfo As String, cur��� As Currency
    Dim curסԺ���� As Currency, curסԺ�޶� As Currency, cur�������� As Currency
    
    frmIDentified.Tag = bytType
    frmIDentified.Show 1
    'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    strPatiInfo = frmIDentified.mstrPatiInfo
    cur��� = frmIDentified.mcur���
    curסԺ���� = frmIDentified.mcurסԺ����
    cur�������� = frmIDentified.mcur��������
    curסԺ�޶� = frmIDentified.mcurסԺ�޶�
    Unload frmIDentified
    
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������� (1����������);
        '24��������;25�����ۼ�;26����ͳ���޶�
        
        '200308z012
        lng����ID = BuildPatiInfo(bytType, strPatiInfo & ";;;;" & cur��� & ";;;;;;;" & _
            cur��� & ";;;;;;" & curסԺ���� & ";" & cur�������� & ";" & curסԺ�޶�, lng����ID)
        If lng����ID = 0 Then Exit Function
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = strPatiInfo & ";" & lng����ID & ";;;;" & cur��� & ";;;;;;;" & cur��� & ";;;;;"
    End If
    ��ݱ�ʶ_�ɶ� = strPatiInfo
End Function

Public Function �������_�ɶ�(strSelfNo As String, Optional bytYear As Byte) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��,bytYear-�������,0-�������,1-�������,2-�������
'����: ���ظ����ʻ����Ľ��
    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next
    With rsTmp
        gstrSQL = "Select grzhlnye,grzhbnye From grjbxx Where grbm='" & strSelfNo & "'"
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If .RecordCount > 0 Then
            Select Case bytYear
            Case 1
                �������_�ɶ� = .Fields(1).Value
            Case 2
                �������_�ɶ� = .Fields(0).Value
            Case Else
                �������_�ɶ� = .Fields(0).Value + .Fields(1).Value
            End Select
        Else
            �������_�ɶ� = 0
        End If
    End With
End Function

Public Function �������_�ɶ�(lng����ID As Long, lng����ID As Long, strҽ���� As String, str���� As String, str���� As String, curȫ�Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'      strҽ����     ҽ����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£�
'        ��Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    Dim strSerial As String, lngCount As Long, cur��� As Currency
    Dim rsList As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim cur����֧�� As Currency, cur�������� As Currency, cur�����Ը� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date
    Dim cur�������� As Currency, cur�����ۼ� As Currency, cur����ͳ���޶� As Currency
    
    strSerial = getSerial(strҽ����)
    
    With rsList
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
        gstrSQL = "Select * From ���˷��ü�¼ Where Nvl(���ӱ�־,0)<>9 And ����ID=" & lng����ID
        gstrSQL = "Select A.NO,A.�Ǽ�ʱ��,A.������ as ҽ��," & _
                "   A.����*A.���� as ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ��," & _
                "   D.��Ŀ���� as �շ���Ŀ,B.���� as ��Ŀ����," & _
                "   decode(Instr(B.���,'��'),0,B.���,substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
                "   decode(Instr(B.���,'��'),0,'',substr(B.���,Instr(B.���,'��')+1)) as ����," & _
                "   C.���� as ��������" & _
                " From (" & gstrSQL & ") A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D " & _
                " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=" & gintInsure & _
                " Order by A.ID"
        Call OpenRecordset(rsList, "�ɶ�ҽ��")
        
        If .RecordCount = 0 Then
            MsgBox "û����д�շѼ�¼��", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '���������ϸ(Z003)
        Dim strFeeKind As String
        lngCount = 0
        Do While Not .EOF
            lngCount = lngCount + 1
            gstrSQL = "Select sfdlmc From sfxmdl Where sfdlbm='" & Left(!�շ���Ŀ, 3) & "'"
            With rsTmp
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open gstrSQL, gcnSybase, adOpenKeyset
                strFeeKind = .Fields(0).Value
            End With
            gstrSQL = "insert into zfymx(jysxh,sfsj,pcno,grbm," & _
                    "   sfdlbm,sfxmbm,sl,sjjg," & _
                    "   cd,gg,yfyl,fyze,zfbl," & _
                    "   txbz,bpbz,qzfbf,ggzfbf,yxbxbf,fyshbz," & _
                    "   sfy,jbr,bz,sfdlmc,sfxmmc," & _
                    "   sjph,xh,yybm,ksbm,fylx," & _
                    "   tjdm,ysxm,ksmc,blh,zyh) " & _
                    " values ('" & lng����ID & "3',getdate(),'" & UserInfo.վ�� & "','" & strҽ���� & "'," & _
                    "   '" & Left(!�շ���Ŀ, 3) & "','" & !�շ���Ŀ & "'," & !���� & "," & !ʵ�ʼ۸� & "," & _
                    "   '" & !���� & "','" & !��� & "',''," & !���ʽ�� & ",0," & _
                    "   '','',0,0,0,''," & _
                    "   '" & UserInfo.���� & "','" & UserInfo.���� & "','','" & strFeeKind & "','" & !��Ŀ���� & "'," & _
                    "   '" & lng����ID & "3','" & lngCount & "','" & Trim(gstrҽԺ����) & "','',''," & _
                    "   '','" & !ҽ�� & "','" & !�������� & "','" & !no & "','')"
            gcnSybase.Execute gstrSQL
            
            cur�������� = cur�������� + !���ʽ��
            .MoveNext
        Loop
    End With
    
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    gstrSQL = "z003('z003','" & UserInfo.վ�� & "','" & lng����ID & "3','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & lng����ID & "3','11','" & str���� & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc

    '����Ƿ���ȷ(zjycl)
    With rsTmp
        gstrSQL = "select code from zjycl where jysxh='" & lng����ID & "3' And jybh='z003'"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "����""z003""���ִ���""" & !code & """:" & vbCrLf & String(2, "��") & GetErrInfo(!code) & String(2, vbTab), vbInformation, gstrSysName
            �������_�ɶ� = False
            Exit Function
        End If
    End With
    
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    gstrSQL = "z008('z008','" & UserInfo.վ�� & "','" & lng����ID & "8','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','','11','" & str���� & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    With rsTmp
        '����Ƿ���ȷ(zjycl)
        gstrSQL = "select code from zjycl where jysxh='" & lng����ID & "8' And jybh='z008'"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "����""z008""���ִ���""" & !code & """:" & vbCrLf & String(2, "��") & GetErrInfo(!code) & String(2, vbTab), vbInformation, gstrSysName
            �������_�ɶ� = False: Exit Function
        End If
        '---------------------------------------------------------------------------------------------
        '��д�����
        curDate = zlDatabase.Currentdate
                
        cur��� = �������_�ɶ�(strҽ����)
        
        '������ʻ�֧�����
        gstrSQL = "Select ��Ԥ�� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ����ID=" & lng����ID
        Call OpenRecordset(rsTmp, "�ɶ�ҽ��")
        
        If Not .EOF Then cur����֧�� = IIf(IsNull(!��Ԥ��), 0, !��Ԥ��)
                
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, cur����ͳ���޶�)
                        
        '200308z012:"��������=סԺ����","����ͳ���޶�=��������"
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & cur����ͳ���޶� & ")"
        Call ExecuteProcedure("�ɶ�ҽ��")
        
        '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
            Year(curDate) & "," & cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & cur�������� & "," & _
            curȫ�Ը� & "," & cur�����Ը� & ",NULL,NULL,NULL,NULL," & cur����֧�� & ",NULL)"
        Call ExecuteProcedure("�ɶ�ҽ��")
        '---------------------------------------------------------------------------------------------
    End With
    �������_�ɶ� = True
End Function

Public Function �����ʻ�תԤ��_�ɶ�(lngԤ��ID As Long, curMoney As Currency, rsԤ����¼ As ADODB.Recordset) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim strҽ���� As String, str���� As String, strSerial As String, str���� As String
    Dim lng����ID As Long, lng��ҳID As Long, cur��� As Currency, cur��� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, cur�������� As Currency
    Dim cur�����ۼ� As Currency, cur����ͳ���޶� As Currency
    Dim rsTmp As New ADODB.Recordset, curDate As Date
    Dim strDJZT As String
    
    With rsԤ����¼
        lng����ID = rsԤ����¼!����ID
        lng��ҳID = IIf(IsNull(rsԤ����¼!��ҳID), 0, rsԤ����¼!��ҳID)
        str���� = TrimStr(IIf(IsNull(!����), "", !����))
        strҽ���� = TrimStr(IIf(IsNull(!ҽ����), str����, !ҽ����))
        str���� = TrimStr(IIf(IsNull(!����), "", !����))
        strSerial = getSerial(strҽ����)
        
        cur��� = !���
        cur��� = �������_�ɶ�(strҽ����, 1) 'ȡ�������,�������϶��������ʽ��
    End With
    
    strDJZT = Trim(GetGrjbxx(strҽ����, "djzt"))
    If strDJZT <> "120" Then
        MsgBox "��ҽ��������δ��Ժ,����ִ�и����ʻ�תԤ�����ף�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�������ݵ������ʻ�֧����
    gstrSQL = "insert into zgrzhzf(jysxh,pcno,grbm," & _
            "   yybm,zfsj,bnzhzf,lnzhzf,jbr,zfyy,bz)" & _
            " values ('" & lngԤ��ID & "A','" & UserInfo.վ�� & "','" & strҽ���� & "'," & _
            "   '" & Trim(gstrҽԺ����) & "',getdate()," & _
            IIf(cur��� >= cur���, cur���, cur���) & "," & _
            IIf(cur��� >= cur���, 0, cur��� - cur���) & "," & _
            "   '" & UserInfo.���� & "','','')"
    gcnSybase.Execute gstrSQL
    
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    gstrSQL = "z010('z010','" & UserInfo.վ�� & "','" & lngԤ��ID & "A','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & lngԤ��ID & "A'," & _
        IIf(lng��ҳID = 0, "'11'", "'31'") & ",'" & str���� & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    With rsTmp
        '����Ƿ���ȷ(zjycl)
        gstrSQL = "Select code From zjycl Where jysxh='" & lngԤ��ID & "A' And jybh='z010'"
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "����""z010""���ִ���""" & !code & """:" & vbCrLf & String(2, "��") & GetErrInfo(!code) & String(2, vbTab), vbInformation, gstrSysName
            �����ʻ�תԤ��_�ɶ� = False: Exit Function
        End If
        '---------------------------------------------------------------------------------------------
        '��д�����
        curDate = zlDatabase.Currentdate
        
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, cur����ͳ���޶�)
        If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
        
        cur��� = �������_�ɶ�(strҽ����) 'ȡ�������
        
        '200308z012:"��������=סԺ����","����ͳ���޶�=��������"
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & cur����ͳ���޶� & ")"
        Call ExecuteProcedure("�ɶ�ҽ��")
        
        '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա�����Ԥ��ID�϶�Ϊ����)
        gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & gintInsure & "," & lng����ID & "," & _
            Year(curDate) & "," & cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & ",NULL," & cur�������� & "," & _
            cur��� & ",NULL,NULL,NULL,NULL,NULL,NULL," & cur��� & ",NULL)"
        Call ExecuteProcedure("�ɶ�ҽ��")
        '---------------------------------------------------------------------------------------------
    End With
    �����ʻ�תԤ��_�ɶ� = True
End Function

Public Function ��Ժ�Ǽ�_�ɶ�(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim jysxh As String, INDate As String, strInNote As String
    Dim strSelfNo As String, strSelfPwd As String, strSerial As String, strKH As String
    Dim rsTmp As New ADODB.Recordset, curDate As Date

    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curסԺ���� As Currency, cur�������� As Currency, curסԺ�޶� As Currency

    jysxh = zlDatabase.GetNextId("���ű�") & "2"
    'New
    gstrSQL = "Select A.��Ժ����,A.��Ժ����,B.����,D.סԺ��,SysDate as ����ʱ��,C.����,C.ҽ����,C.���� " & _
            " From ������ҳ A,���ű� B,�����ʻ� C,������Ϣ D " & _
            " Where A.����ID=D.����ID And A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
            " And A.��Ժ����ID=B.ID And A.����ID=C.����ID And C.����=" & gintInsure
    Call OpenRecordset(rsTmp, "�ɶ�ҽ��")
    
    strKH = TrimStr(IIf(IsNull(rsTmp!����), "", rsTmp!����))
    strSelfNo = TrimStr(IIf(IsNull(rsTmp!ҽ����), strKH, rsTmp!ҽ����))
    strSelfPwd = TrimStr(IIf(IsNull(rsTmp!����), "", rsTmp!����))
    
    If strSelfNo = "" Then
        MsgBox "û�д˲��˻�˲��˲���ҽ�����ˣ�", vbExclamation, gstrSysName
        Exit Function
    End If
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("�ɶ�ҽ��")
            
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
    strSerial = getSerial(strSelfNo)
    
    Dim mSqlTemp As String
    mSqlTemp = ""
    '�ύסԺ�ǼǱ�
    mSqlTemp = "insert into zzydj(jysxh,pcno,yybm,grbm,ryzd,rysj,ryks,rycw,ryjbr,blh,zyh,sftzb,tzbbxbl,bpbz,jbsj)" & _
            " values('" & jysxh & "','" & UserInfo.վ�� & "','" & Trim(gstrҽԺ����) & "','" & strSelfNo & "'," & _
            "'" & strInNote & "','" & Format(rsTmp!��Ժ����, "yyyy-MM-dd hh:mm:ss") & "','" & rsTmp("����") & "','" & rsTmp("��Ժ����") & "','" & _
            UserInfo.��� & "','','" & rsTmp("סԺ��") & "','',0,'','" & Format(rsTmp!����ʱ��, "yyyy-MM-dd hh:mm:ss") & "')"
    gcnSybase.Execute mSqlTemp
    rsTmp.Close
    
    '�ύ���׵ǼǱ�
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    gstrSQL = "z002('z002','" & UserInfo.վ�� & "','" & jysxh & "','" & strSelfPwd & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strSelfNo & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & jysxh & "','31','" & strKH & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    '����Ƿ���ȷ(zjycl)
    gstrSQL = "Select code From zjycl Where jysxh='" & jysxh & "' And jybh='z002'"
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.Open gstrSQL, gcnSybase, adOpenStatic, adLockReadOnly
    If Trim(rsTmp("code").Value) <> "0000" Then
        MsgBox "����""z002""���ִ���""" & rsTmp!code & """:" & vbCrLf & String(2, "��") & GetErrInfo(rsTmp!code) & String(2, vbTab), vbInformation, gstrSysName
        ��Ժ�Ǽ�_�ɶ� = False
        Exit Function
    End If
    
    '200308z012:ɾ��ȡ˳���,���˲���ʹ�ù̶�˳���
    
    '��д�ʻ������Ϣ
    curDate = zlDatabase.Currentdate
    Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
        
    '200308z012:����סԺ�����ͱ�������
    curסԺ���� = Val(GetGrjbxx(strSelfNo, "zyjs")) '���浽"��������"
    cur�������� = Val(GetGrjbxx(strSelfNo, "tcbxbl")) '���浽"�����ۼ�"
    curסԺ�޶� = Val(GetGrjbxx(strSelfNo, "zyxe")) '���浽"����ͳ���޶�"
    
    '200308z012:"��������=סԺ����","����ͳ���޶�=��������"
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & curסԺ���� & "," & cur�������� & "," & curסԺ�޶� & ")"
    Call ExecuteProcedure("�ɶ�ҽ��")
    
    ��Ժ�Ǽ�_�ɶ� = True
End Function

Public Function ��Ժ�Ǽ�_�ɶ�(lng����ID As Long, lng��ҳID As Long, rs���� As ADODB.Recordset) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTmp As New ADODB.Recordset
    Dim jysxh As String, OutDate As String
    Dim strSelfNo As String, strSelfPwd As String, strSerial As String, strKH As String
    
    'New
    strKH = TrimStr(IIf(IsNull(rs����!����), "", rs����!����))
    strSelfNo = TrimStr(IIf(IsNull(rs����!ҽ����), strKH, rs����!ҽ����))
    strSelfPwd = TrimStr(IIf(IsNull(rs����!����), "", rs����!����))
    
    strSerial = getSerial(strSelfNo)
    jysxh = zlDatabase.GetNextId("���ű�") & "B"
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & gintInsure & ")"
    Call ExecuteProcedure("�ɶ�ҽ��")
    
    '�ύ��Ժ�ǼǱ�
    gstrSQL = "Select A.��Ժ����,A.��Ժ����,SysDate as ����ʱ��,B.סԺ��,A.��Ժ��ʽ,C.����" & _
        " From ������ҳ A,������Ϣ B,���ű� C" & _
        " Where A.��Ժ����ID=C.ID And A.����ID=B.����ID And A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID
    rsTmp.CursorLocation = adUseClient
    Call OpenRecordset(rsTmp, "�ɶ�ҽ��")
    
    gstrSQL = "insert into zcydj(jysxh,pcno,grbm,yybm,cysj,cyzd,cycw,cyjbr,blh,zyh,jbsj,cyyy,cyks,zyzt) " & _
            "values('" & jysxh & "','" & UserInfo.վ�� & "','" & strSelfNo & "','" & Trim(gstrҽԺ����) & "','" & _
            Format(rsTmp!��Ժ����, "yyyy-MM-dd hh:mm:ss") & "','','" & IIf(IsNull(rsTmp!��Ժ����), "", rsTmp!��Ժ����) & "','" & UserInfo.��� & "','','" & _
            rsTmp("סԺ��") & "','" & Format(rsTmp!����ʱ��, "yyyy-MM-dd hh:mm:ss") & "','" & IIf(IsNull(rsTmp!��Ժ��ʽ), "", rsTmp!��Ժ��ʽ) & "','" & rsTmp("����") & "','')"
    gcnSybase.Execute gstrSQL
    
    '�ύ���׵ǼǱ�
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    gstrSQL = "z011('z011','" & UserInfo.վ�� & "','" & jysxh & "','" & strSelfPwd & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strSelfNo & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & jysxh & "','31','" & strKH & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    '����Ƿ���ȷ(zjycl)
    gstrSQL = "Select code From zjycl Where jysxh='" & jysxh & "' And jybh='z011'"
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.Open gstrSQL, gcnSybase, adOpenStatic, adLockReadOnly
    If Trim(rsTmp("code").Value) <> "0000" Then
        MsgBox "����""z011""���ִ���""" & rsTmp!code & """:" & vbCrLf & String(2, "��") & GetErrInfo(rsTmp!code) & String(2, vbTab), vbInformation, gstrSysName
        ��Ժ�Ǽ�_�ɶ� = False
        Exit Function
    End If
    ��Ժ�Ǽ�_�ɶ� = True
End Function

Public Function סԺ�������_�ɶ�(rsList As ADODB.Recordset, strҽ���� As String, str���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsList-��Ҫ����ķ�����ϸ��¼���ϣ�strҽ����-ҽ���ţ�str����-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim str˳��� As String, str���� As String
    Dim lng��� As Integer, lng����ID As Long
    Dim strSerial As String, str���� As String
    Dim strSql As String, str��ע As String
    Dim blnTran As Boolean, i As Long
    Dim rsTmp As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    
    Dim cur�ܶ� As Currency, cur�޶� As Currency, cur�۸� As Currency
    Dim cur���� As Currency, curȫ�Է� As Currency
    Dim cur�����Ը�20 As Currency, cur�����Ը�80 As Currency
    Dim cur��λ�����Ը� As Currency, cur��λ�޶�� As Currency
    
    On Error GoTo errH
    
    rsList.Filter = "Ӥ����=0"
    If rsList.RecordCount = 0 Then Exit Function
    
    g��������.����ID = rsList!����ID
    g��������.��ҳID = rsList!��ҳID
    lng����ID = rsList!����ID
    
    '��ȡ���˵�һЩ�ʻ���Ϣ
    strSql = "Select * From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & gintInsure
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then Exit Function
    
    str���� = TrimStr(IIf(IsNull(rsTmp!����), "", rsTmp!����))
    strҽ���� = TrimStr(IIf(IsNull(rsTmp!ҽ����), str����, rsTmp!ҽ����))
    str���� = TrimStr(IIf(IsNull(rsTmp!����), "", rsTmp!����))
    strSerial = getSerial(strҽ����)
    
    '����Z003���׵�˳��źͿ�ʼ���
    lng��� = 1
    str˳��� = zlDatabase.GetNextId("���˽��ʼ�¼")
    
    '��SybaseFace���ȡ�շ�ϸĿ�����嵥
    strSql = "select * from sfxmdl"
    Set rs���� = New ADODB.Recordset
    rs����.CursorLocation = adUseClient
    rs����.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
    
    '���������ϸzfymx
    gcnOracle.BeginTrans: blnTran = True
    For i = 1 To rsList.RecordCount
        If rsList!��ҳID > g��������.��ҳID Then g��������.��ҳID = rsList!��ҳID
        
        '��λ����Ϊ�ǵ����޶�,�������,����ҲҪ����
        If Left(rsList!ҽ����Ŀ����, 3) = "002" And Mid(rsList!ҽ����Ŀ����, 8, 1) = "2" Then
            cur�۸� = rsList!��� / IIf(rsList!���� = 0, 1, rsList!����)
        Else
            cur�۸� = rsList!�۸�
        End If
        
        'ֻ�ϴ�δ�ϴ�����
        '-----------------------------------------------------------------------------
        If rsList!�Ƿ��ϴ� = 0 Then
            '��ȡ�շѴ�������
            str���� = ""
            rs����.Filter = "sfdlbm='" & Left(rsList!ҽ����Ŀ����, 3) & "'"
            If Not rs����.EOF Then str���� = NVL(rs����!sfdlmc)
        
            '����zfymx,����ϸ�����ڽ���(z003)
            'sfsjҪ�õ�ǰʱ��,��Ȼ�����ٴ�ʱ��Υ��ΨһԼ��
            With rsList
                str��ע = "Ԥ���ϴ�:" & !no & ",���:" & !���
                strSql = _
                    "insert into zfymx(" & _
                    "jysxh,sfsj,pcno,grbm,sfdlbm,sfxmbm,sl,sjjg,cd,gg,yfyl,fyze,zfbl,txbz,bpbz,qzfbf,ggzfbf,yxbxbf," & _
                    "fyshbz,sfy,jbr,bz,sfdlmc,sfxmmc,sjph,xh,yybm,ksbm,fylx,tjdm,ysxm,ksmc,blh,zyh) values (" & _
                    "'" & str˳��� & "3',getdate()," & _
                    "'" & UserInfo.վ�� & "','" & strҽ���� & "','" & Left(!ҽ����Ŀ����, 3) & "','" & !ҽ����Ŀ���� & "'," & _
                    Format(!����, "0.00") & "," & Format(cur�۸�, "0.00") & ",'" & IIf(IsNull(!����), "", !����) & "'," & _
                    "'" & IIf(IsNull(!���), "", !���) & "',''," & Format(!���, "0.00") & ",0,'','',0,0,0,''," & _
                    "'" & UserInfo.���� & "','" & UserInfo.���� & "','" & str��ע & "','" & str���� & "','" & !�շ����� & "'," & _
                    "'" & str˳��� & "3','" & lng��� & "','" & Trim(gstrҽԺ����) & "','','',''," & _
                    "'" & IIf(IsNull(!ҽ��), "", !ҽ��) & "','" & !�������� & "','" & lng����ID & "','" & lng����ID & "')"
            End With
            gcnSybase.Execute strSql
        
            '��Ǹ÷������ϴ�(��δ�ύ)
            strSql = "ZL_���˷��ü�¼_�ϴ�('" & rsList!no & "'," & rsList!��� & "," & rsList!��¼���� & "," & rsList!��¼״̬ & ")"
            gcnOracle.Execute strSql, , adCmdStoredProc
            
            lng��� = lng��� + 1
        End If
        '-----------------------------------------------------------------------------
                
        '����z012��ҪһЩ���
        '-----------------------------------------------------------------------------
        cur�ܶ� = cur�ܶ� + Format(rsList!���, "0.00")
        If Left(rsList!ҽ����Ŀ����, 3) <> "002" Then
            '�Ǵ�λ��
            Select Case Mid(rsList!ҽ����Ŀ����, 8, 1)
                Case "1" '����
                    cur���� = cur���� + Format(rsList!���, "0.00")
                Case "2" '�����Ը�
                    cur�����Ը�20 = cur�����Ը�20 + Format(Format(rsList!���, "0.00") * 0.2, "0.00")
                    cur�����Ը�80 = cur�����Ը�80 + Format(Format(rsList!���, "0.00") * 0.8, "0.00")
                Case "3" 'ȫ�Է�
                    curȫ�Է� = curȫ�Է� + Format(rsList!���, "0.00")
            End Select
        Else
            '��λ�Ѳ���
            If Mid(rsList!ҽ����Ŀ����, 8, 1) = "1" Then 'ȫ���Ѵ�λ��
                'cur��λ�����Ը� = cur��λ�����Ը� + 0
                cur��λ�޶�� = cur��λ�޶�� + Format(rsList!���, "0.00")
            ElseIf Mid(rsList!ҽ����Ŀ����, 8, 1) = "3" Then 'ȫ�ԷѴ�λ��
                'cur��λ�޶�� = cur��λ�޶�� + 0
                cur��λ�����Ը� = cur��λ�����Ը� + Format(rsList!���, "0.00")
            ElseIf Mid(rsList!ҽ����Ŀ����, 8, 1) = "2" Then '�����޶λ��
                Select Case Mid(rsList!ҽ����Ŀ����, 11, 2)
                    Case "01": cur�޶� = 5
                    Case "02": cur�޶� = 6.5
                    Case "03": cur�޶� = 7
                    Case "04": cur�޶� = 7.5
                    Case "05": cur�޶� = 10
                    Case "06": cur�޶� = 12
                    Case "07": cur�޶� = 5.5
                    Case "08": cur�޶� = 8
                    Case "09": cur�޶� = 9.75
                    Case "10": cur�޶� = 10.5
                    Case "11": cur�޶� = 11.25
                    Case "12": cur�޶� = 15
                    Case "13": cur�޶� = 2
                    Case "14": cur�޶� = 40
                    Case "15": cur�޶� = 25
                    Case Else: cur�޶� = 0
                End Select
                
                If cur�۸� <= cur�޶� Then
                    cur��λ�޶�� = cur��λ�޶�� + Format(cur�۸�, "0.00") * Format(rsList!����, "0.00")
                Else
                    cur��λ�޶�� = cur��λ�޶�� + cur�޶� * Format(rsList!����, "0.00")
                    cur��λ�����Ը� = cur��λ�����Ը� + (Format(cur�۸�, "0.00") - cur�޶�) * Format(rsList!����, "0.00")
                End If
            End If
        End If
        '-----------------------------------------------------------------------------
        
        rsList.MoveNext
    Next
   
    '�ύ������ϸ
    '-----------------------------------------------------------------------------
    If lng��� > 1 Then
        'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
        strSql = "z003('z003','" & UserInfo.վ�� & "','" & str˳��� & "3','" & str���� & "','" & UserInfo.��� & "'," & _
            "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & str˳��� & "3','31','" & str���� & "')"
        gcnSybase.Execute strSql, , adCmdStoredProc
    
        '����Ƿ���ȷ(zjycl)
        strSql = "Select code From zjycl where grbm='" & strҽ���� & "' and jysxh='" & str˳��� & "3' and jybh='z003'"
        Set rsTmp = New ADODB.Recordset
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
        If rsTmp.EOF Then
            gcnOracle.RollbackTrans
            MsgBox "δ���ֽ��״�������", vbInformation, gstrSysName
            Exit Function
        ElseIf Trim(rsTmp!code) <> "0000" Then
            gcnOracle.RollbackTrans
            MsgBox "����""z003""���ִ���""" & rsTmp!code & """:" & vbCrLf & String(2, "��") & GetErrInfo(rsTmp!code) & String(2, vbTab), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    '-----------------------------------------------------------------------------
    
    '���㱨�����
    '-----------------------------------------------------------------------------
    '����˳���,�ͻ������,���˱���,ҽԺ����,�����ܶ�,ȫ�Ը�����,�ҹ��Ը�����,����������
    '����ͳ�ﲿ��,ͳ��֧������,ͳ��֧������,�����ʻ�֧��,�籣�ֿ۳�,��ע,֧�����
    '֧�����:?
    strSql = "insert into zyxnjs(" & _
        "jysxh,pcno,grbm,yybm,fyze,qzfbf,ggzfbf,yxbxbf,jrjsbf,tczifbl,tczhifbf,grzhzf,sbjkc,bz) values(" & _
        "'" & str˳��� & "C','" & UserInfo.վ�� & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "'," & _
        Format(cur�ܶ�, "0.00") & "," & Format(curȫ�Է� + cur��λ�����Ը�, "0.00") & "," & Format(cur�����Ը�20, "0.00") & "," & _
        Format(cur�����Ը�80 + cur��λ�޶�� + cur����, "0.00") & ",0,0,0,0,0,'')"
    gcnSybase.Execute strSql
    
    '���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    strSql = "z012('z012','" & UserInfo.վ�� & "','" & str˳��� & "C','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & str˳��� & "C','31','" & str���� & "')"
    gcnSybase.Execute strSql, , adCmdStoredProc
    
    '����Ƿ���ȷ(zjycl)
    strSql = "Select code From zjycl where grbm='" & strҽ���� & "' and jysxh='" & str˳��� & "C' and jybh='z012'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        MsgBox "δ���ֽ��״�������", vbInformation, gstrSysName
        Exit Function
    ElseIf Trim(rsTmp!code) <> "0000" Then
        MsgBox "����""z012""���ִ���""" & rsTmp!code & """:" & vbCrLf & String(2, "��") & GetErrInfo(rsTmp!code) & String(2, vbTab), vbInformation, gstrSysName
        Exit Function
    End If

    '����:����ͳ�ﲿ��,ͳ��֧������,�����ʻ�֧��
    strSql = "Select jrjsbf,tczhifbf,grzhzf From zyxnjs where grbm='" & strҽ���� & "' and jysxh='" & str˳��� & "C'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
    If Not rsTmp.EOF Then סԺ�������_�ɶ� = "ҽ������;" & rsTmp!tczhifbf & ";0"
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Function

Public Function סԺ����_�ɶ�(lng����ID As Long, rs�ʻ� As ADODB.Recordset) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ�
'        ��˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״���
'        ����������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס�
'        �������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�
'        (���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    Dim strҽ���� As String, str���� As String, str���� As String
    Dim strSerial As String, lng����ID As Long
    Dim str���� As String, strSql As String, i As Long
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, cur�������� As Currency, curDate As Date
    Dim cur�����ۼ� As Currency, cur����ͳ���޶� As Currency
    
    Dim curסԺ���� As Currency, cur�������� As Currency, cur֧������ As Double
    Dim cur����ͳ�� As Currency, curͳ��֧�� As Currency
    Dim cur�����Ը� As Currency, curȫ�Ը� As Currency
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '��ȡ��һЩ�ʻ���Ϣ
    lng����ID = rs�ʻ�!����ID
    str���� = TrimStr(IIf(IsNull(rs�ʻ�!����), "", rs�ʻ�!����))
    strҽ���� = TrimStr(IIf(IsNull(rs�ʻ�!ҽ����), str����, rs�ʻ�!ҽ����))
    str���� = TrimStr(IIf(IsNull(rs�ʻ�!����), "", rs�ʻ�!����))
    strSerial = getSerial(strҽ����)
    
    '��������
    '---------------------------------------------------------------------------------------------------------------
    'ɾ����Ӧ˳��ŵ�zjycl,zfzjs,�Ա����ظ�
    strSql = "delete from zjycl where grbm='" & strҽ���� & "' and jysxh='" & lng����ID & "8' and jybh='z008'"
    gcnSybase.Execute strSql
    strSql = "delete from zfzjs where grbm='" & strҽ���� & "' and jysxh='" & lng����ID & "8'"
    gcnSybase.Execute strSql
    
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    'jysxhҪʹ�õ�ǰ�Ľ���ID,�Ա�ִ�н�������z013֮ǰ��ȡ��Ӧ��Ϣ
    strSql = "z008('z008','" & UserInfo.վ�� & "','" & lng����ID & "8','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','','31','" & str���� & "')"
    gcnSybase.Execute strSql, , adCmdStoredProc
         
     '����Ƿ���ȷ(zjycl)
    strSql = "Select code From zjycl Where grbm='" & strҽ���� & "' and jysxh='" & lng����ID & "8' And jybh='z008'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        MsgBox "δ���ֽ��״�������", vbInformation, gstrSysName
        Exit Function
    ElseIf Trim(rsTmp!code) <> "0000" Then
        MsgBox "����""z008""���ִ���""" & rsTmp!code & """:" & vbCrLf & String(2, "��") & GetErrInfo(rsTmp!code) & String(2, vbTab), vbInformation, gstrSysName
        Exit Function
    End If
    
    '��д�����
    '---------------------------------------------------------------------------------------------------------------
    curDate = zlDatabase.Currentdate

    'סԺ����,�����ܶ�,����ͳ�ﲿ��,ͳ��֧������,ȫ�Ը�����
    strSql = "select zyjs,fyze,yxbxbf,tczhifbf,qzfbf,tczifbl from zfzjs where jysxh='" & lng����ID & "8' and grbm='" & strҽ���� & "'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        MsgBox "δ���ظ��������¼��", vbInformation, gstrSysName
        Exit Function
    End If
    
    cur֧������ = IIf(IsNull(rsTmp!tczifbl), 0, rsTmp!tczifbl) * 100 'Ϊ�˱������㹻��С��λ������ԭ�б����ϳ���100
    curסԺ���� = rsTmp!zyjs
    cur�������� = rsTmp!fyze
    cur����ͳ�� = rsTmp!yxbxbf
    curͳ��֧�� = rsTmp!tczhifbf
    curȫ�Ը� = rsTmp!qzfbf
    cur�����Ը� = cur�������� - curȫ�Ը� - cur����ͳ��
    
    '�ȽϽ�������Ԥ�����Ƿ�һ��
    strSql = "Select ��Ԥ�� From ����Ԥ����¼ Where ��¼����=2 And ���㷽ʽ='ҽ������' And ����ID=" & lng����ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset
    If rsTmp.EOF Then
        If curͳ��֧�� <> 0 Then
            MsgBox "δ����Ԥ���¼��", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf curͳ��֧�� <> NVL(rsTmp!��Ԥ��, 0) Then
        MsgBox "ͳ��֧�����Ϊ:" & Format(curͳ��֧��, "0.00") & " ,��Ԥ����Ľ����һ�£�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����ν��ʼ�¼���Ϊ���ϴ�
    strSql = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    gcnOracle.Execute strSql, , adCmdStoredProc
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, cur����ͳ���޶�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '200308z012:"��������=סԺ����","����ͳ���޶�=��������"
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� + cur����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & cur����ͳ���޶� & ")"
    Call ExecuteProcedure("�ɶ�ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & curסԺ���� & "," & cur֧������ & "," & curסԺ���� & "," & _
        cur�������� & "," & curȫ�Ը� & "," & cur�����Ը� & "," & cur����ͳ�� & "," & curͳ��֧�� & "," & _
        "NULL,NULL,NULL,NULL)"
    Call ExecuteProcedure("�ɶ�ҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & cur����ͳ�� & "," & curͳ��֧�� & ",NULL)"
    Call ExecuteProcedure("�ɶ�ҽ��")
    
    סԺ����_�ɶ� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function סԺ�������_�ɶ�(lng����ID As Long, rs�ʻ� As ADODB.Recordset) As Boolean
'----------------------------------------------------------------
'���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
'������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
'      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ�
'        �ڲ��˷��ü�¼�и��ݽ���ID���ң�
'      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ�
'        ��������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
'----------------------------------------------------------------
    Dim strҽ���� As String, str���� As String, str���� As String
    Dim cur�����ܶ� As Currency, strSerial As String, lng����ID As Long
    Dim str������ As String, str˳��� As String
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, cur�������� As Currency
    Dim cur�����ۼ� As Currency, cur����ͳ���޶� As Currency
    Dim curDate As Date, lng��ID As Long
    
    Dim curסԺ���� As Currency, cur�������� As Currency, cur֧������ As Double
    Dim cur����ͳ�� As Currency, curͳ��֧�� As Currency
    Dim cur�����Ը� As Currency, curȫ�Ը� As Currency
    
    Dim rsTmp As ADODB.Recordset, strSql As String
        
    On Error GoTo errH
        
    '������Ϣ
    lng����ID = rs�ʻ�!����ID
    str���� = TrimStr(IIf(IsNull(rs�ʻ�!����), "", rs�ʻ�!����))
    strҽ���� = TrimStr(IIf(IsNull(rs�ʻ�!ҽ����), str����, rs�ʻ�!ҽ����))
    str���� = TrimStr(IIf(IsNull(rs�ʻ�!����), "", rs�ʻ�!����))
    strSerial = getSerial(strҽ����)
    
    'ԭ"�����ܶ�,������"
    strSql = "select fyze,jsbh from zfzjs where jysxh='" & lng����ID & "8' and grbm='" & strҽ���� & "'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        MsgBox "���ʼ�¼δ�ҵ���", vbInformation, gstrSysName
        Exit Function
    End If
    str������ = rsTmp!jsbh
    cur�����ܶ� = IIf(IsNull(rsTmp!fyze), 0, rsTmp!fyze)
    
    '������ý����
    strSql = _
        "insert into zfyjs(jysxh,pcno,grbm,yybm,zyjs," & _
        " nspgz,fyze,qzfbf,ggzfbf,yxbxbf,jrjsbf,tczifbl," & _
        " tczhifbf,grzhzf,zfsm,sbjkc,jbr,sfy,jbsj,bz,jsbh)" & _
        " values('" & lng����ID & "D','" & UserInfo.վ�� & "','" & _
        strҽ���� & "','" & Trim(gstrҽԺ����) & "',0,0," & _
        cur�����ܶ� & ",0,0,0,0,0,0,0,'',0,'" & UserInfo.��� & "'," & _
        "'" & UserInfo.��� & "',getdate() ,'','" & str������ & "')"
    gcnSybase.Execute strSql
    
    '�ύ���׵ǼǱ�
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    str˳��� = zlDatabase.GetNextId("���˽��ʼ�¼") & "D"
    strSql = "z013('z013','" & UserInfo.վ�� & "','" & str˳��� & "','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & str˳��� & "','31','" & str���� & "')"
    gcnSybase.Execute strSql, , adCmdStoredProc
    
    '����Ƿ���ȷ(zjycl)
    strSql = "Select code From zjycl Where jysxh='" & str˳��� & "' And jybh='z013'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnSybase, adOpenKeyset
    If rsTmp.EOF Then
        MsgBox "δ���ֽ��״�������", vbInformation, gstrSysName
        Exit Function
    ElseIf Trim(rsTmp!code) <> "0000" Then
        MsgBox "����""z013""���ִ���""" & rsTmp!code & """:" & vbCrLf & String(2, "��") & GetErrInfo(rsTmp!code) & String(2, vbTab), vbInformation, gstrSysName
        Exit Function
    End If
    
    '----------------------------------------------------------------------------------
    '��д�����
    curDate = zlDatabase.Currentdate
    '��ȡ���Ϻ�Ľ���ID
    strSql = "Select A.ID From ���˽��ʼ�¼ A,���˽��ʼ�¼ B" & _
        " Where A.NO=B.NO And A.��¼״̬=2 And B.��¼״̬=3" & _
        " And B.ID=" & lng����ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset
    If rsTmp.EOF Then
        MsgBox "δ�������ϵĽ������ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ID = rsTmp!ID
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, cur����ͳ���޶�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
    
    strSql = "Select * From ���ս������ Where Nvl(����,0)=0 And ����ID=" & lng����ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rsTmp.EOF Then
        cur����ͳ�� = IIf(IsNull(rsTmp!����ͳ����), 0, rsTmp!����ͳ����)
        curͳ��֧�� = IIf(IsNull(rsTmp!ͳ�ﱨ�����), 0, rsTmp!ͳ�ﱨ�����)
    End If
    
    strSql = "Select * From ���ս����¼ Where ����=2 And ��¼ID=" & lng����ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rsTmp.EOF Then
        cur֧������ = IIf(IsNull(rsTmp!�ⶥ��), 0, rsTmp!�ⶥ��)
        curסԺ���� = IIf(IsNull(rsTmp!ʵ������), 0, rsTmp!ʵ������)
        cur�������� = IIf(IsNull(rsTmp!�������ý��), 0, rsTmp!�������ý��)
        If cur����ͳ�� = 0 Then cur����ͳ�� = IIf(IsNull(rsTmp!����ͳ����), 0, rsTmp!����ͳ����)
        If curͳ��֧�� = 0 Then curͳ��֧�� = IIf(IsNull(rsTmp!ͳ�ﱨ�����), 0, rsTmp!ͳ�ﱨ�����)
        cur�����Ը� = IIf(IsNull(rsTmp!�����Ը����), 0, rsTmp!�����Ը����)
        curȫ�Ը� = IIf(IsNull(rsTmp!ȫ�Ը����), 0, rsTmp!ȫ�Ը����)
    End If
    
    '�����µ����ϼ�¼
    '200308z012:"��������=סԺ����","����ͳ���޶�=��������"
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� - cur����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� - curͳ��֧�� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & cur����ͳ���޶� & ")"
    Call ExecuteProcedure("�ɶ�ҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng��ID & ",0," & -1 * cur����ͳ�� & "," & -1 * curͳ��֧�� & ",NULL)"
    Call ExecuteProcedure("�ɶ�ҽ��")
    
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng��ID & "," & gintInsure & "," & lng����ID & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & _
        intסԺ�����ۼ� & "," & curסԺ���� & "," & cur֧������ & "," & curסԺ���� & "," & -1 * cur�������� & "," & _
        -1 * curȫ�Ը� & "," & -1 * cur�����Ը� & "," & -1 * cur����ͳ�� & "," & -1 * curͳ��֧�� & "," & _
        "NULL,NULL,NULL,NULL)"
    Call ExecuteProcedure("�ɶ�ҽ��")

    סԺ�������_�ɶ� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function getSerial(strSelfNo As String) As String
'----------------------------------------------------------
'���ܣ���ȡ����˳���
'----------------------------------------------------------
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset

    strSql = "select sxh from grjbxx where grbm='" & strSelfNo & "'"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnSybase, adOpenKeyset
    If Not rsTmp.EOF Then getSerial = rsTmp.Fields(0).Value
End Function

Public Function GetGrjbxx(strSelfNo As String, strField As String) As Variant
'���ܣ���ȡgrjbxx��ָ���ֶε�ֵ
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset

    strSql = "select " & strField & " from grjbxx where grbm='" & strSelfNo & "'"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnSybase, adOpenKeyset
    If Not rsTmp.EOF Then
        GetGrjbxx = IIf(IsNull(rsTmp.Fields(strField).Value), "", rsTmp.Fields(strField).Value)
    End If
End Function

Public Sub ExecuteZ015(ByVal strCard As String, ByRef strҽ���� As String, ByRef str���� As String)
'���ܣ�ִ��Z015����
'������
'   �룺strCard=ˢ��������
'   ����strҽ����=���ݿ����ݽ�����ҽ����
'   ����str����=���ݿ����ݽ����Ŀ���
'˵���������ڳɶ��½ӿ�
    Dim cmdSybase As New ADODB.Command
    
    On Error GoTo errH
    
    With cmdSybase
        Set .ActiveConnection = gcnSybase
        .Parameters.Append .CreateParameter("vid", adVarChar, adParamInput, 30, strCard)
        .Parameters.Append .CreateParameter("vgrbm", adVarChar, adParamOutput, 20)
        .Parameters.Append .CreateParameter("vkh", adVarChar, adParamOutput, 20)
        .CommandType = adCmdStoredProc
        .CommandText = "z015"
        .Execute
        strҽ���� = TrimStr(IIf(IsNull(.Parameters("vgrbm").Value), "", .Parameters("vgrbm").Value))
        str���� = TrimStr(IIf(IsNull(.Parameters("vkh").Value), "", .Parameters("vkh").Value))
    End With
    Exit Sub
errH:
    MsgBox Err.Number & vbCrLf & vbTab & Err.Description, vbInformation, gstrSysName
End Sub

Public Function �ҺŽ���_�ɶ�(lng����ID As Long, lng����ID As Long, strҽ���� As String, str���� As String, str���� As String) As Boolean
'���ܣ����Һ��շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID=�Һż�¼�Ľ���ID��
'Ȩ�ޣ������ʻ�,���˷��ü�¼,�շ�ϸĿ,���ű�,����֧����Ŀ,����Ԥ����¼,�ʻ������Ϣ,zl_�ʻ������Ϣ_insert,zl_���ս����¼_insert
    Dim strSerial As String, lngCount As Long, cur��� As Currency
    Dim rsList As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim cur����֧�� As Currency, cur�������� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date
    Dim cur�������� As Currency, cur�����ۼ� As Currency, cur����ͳ���޶� As Currency
    
    Dim strFeeKind As String
    
    strSerial = getSerial(strҽ����)
    
    With rsList
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
        gstrSQL = "Select A.NO,A.�Ǽ�ʱ��,A.������ as ҽ��," & _
                "   A.����*A.���� as ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ��," & _
                "   D.��Ŀ���� as �շ���Ŀ,B.���� as ��Ŀ����," & _
                "   decode(Instr(B.���,'��'),0,B.���,substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
                "   decode(Instr(B.���,'��'),0,'',substr(B.���,Instr(B.���,'��')+1)) as ����," & _
                "   C.���� as ��������" & _
                " From (Select * From ���˷��ü�¼ Where ����ID=" & lng����ID & ") A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D " & _
                " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=" & gintInsure & _
                " Order by A.ID"
        Call OpenRecordset(rsList, "�ɶ�ҽ��")
        
        If .EOF Then
            MsgBox "û����д�Һż�¼��", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '���������ϸ(Z003)
        lngCount = 0
        Do While Not .EOF
            lngCount = lngCount + 1
            gstrSQL = "Select sfdlmc From sfxmdl Where sfdlbm='" & Left(!�շ���Ŀ, 3) & "'"
            With rsTmp
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open gstrSQL, gcnSybase, adOpenKeyset
                strFeeKind = .Fields(0).Value
            End With
            gstrSQL = "insert into zfymx(jysxh,sfsj,pcno,grbm," & _
                    "   sfdlbm,sfxmbm,sl,sjjg," & _
                    "   cd,gg,yfyl,fyze,zfbl," & _
                    "   txbz,bpbz,qzfbf,ggzfbf,yxbxbf,fyshbz," & _
                    "   sfy,jbr,bz,sfdlmc,sfxmmc," & _
                    "   sjph,xh,yybm,ksbm,fylx," & _
                    "   tjdm,ysxm,ksmc,blh,zyh) " & _
                    " values ('" & lng����ID & "3',getdate(),'" & UserInfo.վ�� & "','" & strҽ���� & "'," & _
                    "   '" & Left(!�շ���Ŀ, 3) & "','" & !�շ���Ŀ & "'," & !���� & "," & !ʵ�ʼ۸� & "," & _
                    "   '" & !���� & "','" & !��� & "',''," & !���ʽ�� & ",0," & _
                    "   '','',0,0,0,''," & _
                    "   '" & UserInfo.���� & "','" & UserInfo.���� & "','','" & strFeeKind & "','" & !��Ŀ���� & "'," & _
                    "   '" & lng����ID & "3','" & lngCount & "','" & Trim(gstrҽԺ����) & "','',''," & _
                    "   '','" & !ҽ�� & "','" & !�������� & "','" & !no & "','')"
            gcnSybase.Execute gstrSQL
            
            cur�������� = cur�������� + !���ʽ��
            .MoveNext
        Loop
    End With
    
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    gstrSQL = "z003('z003','" & UserInfo.վ�� & "','" & lng����ID & "3','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & lng����ID & "3','11','" & str���� & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc

    '����Ƿ���ȷ(zjycl)
    With rsTmp
        gstrSQL = "select code from zjycl where jysxh='" & lng����ID & "3' And jybh='z003'"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "����""z003""���ִ���""" & !code & """:" & vbCrLf & String(2, "��") & GetErrInfo(!code) & String(2, vbTab), vbInformation, gstrSysName
            �ҺŽ���_�ɶ� = False
            Exit Function
        End If
    End With
    
    'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
    gstrSQL = "z008('z008','" & UserInfo.վ�� & "','" & lng����ID & "8','" & str���� & "','" & UserInfo.��� & "'," & _
        "'" & strSerial & "','" & strҽ���� & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','','11','" & str���� & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    With rsTmp
        '����Ƿ���ȷ(zjycl)
        gstrSQL = "select code from zjycl where jysxh='" & lng����ID & "8' And jybh='z008'"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "����""z008""���ִ���""" & !code & """:" & vbCrLf & String(2, "��") & GetErrInfo(!code) & String(2, vbTab), vbInformation, gstrSysName
            �ҺŽ���_�ɶ� = False: Exit Function
        End If
        '---------------------------------------------------------------------------------------------
        '��д�����
        curDate = zlDatabase.Currentdate
                
        cur��� = �������_�ɶ�(strҽ����)
        
        '������ʻ�֧�����
        gstrSQL = "Select ��Ԥ�� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ����ID=" & lng����ID
        Call OpenRecordset(rsTmp, "�ɶ�ҽ��")
        
        If Not .EOF Then cur����֧�� = IIf(IsNull(!��Ԥ��), 0, !��Ԥ��)
                
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, cur����ͳ���޶�)
                        
        '200308z012:"��������=סԺ����","����ͳ���޶�=��������"
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & _
            intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & cur����ͳ���޶� & ")"
        Call ExecuteProcedure("�ɶ�ҽ��")
        
        '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & gintInsure & "," & lng����ID & "," & _
            Year(curDate) & "," & cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & cur�������� & "," & _
            0 & "," & 0 & ",NULL,NULL,NULL,NULL," & cur����֧�� & ",NULL)"
        Call ExecuteProcedure("�ɶ�ҽ��")
        '---------------------------------------------------------------------------------------------
    End With
    �ҺŽ���_�ɶ� = True
End Function

Public Function ���ʴ���_�ɶ�(strNO As String, int���� As Integer, int״̬ As Integer, Optional lng����ID As Long) As Boolean
'���ܣ���סԺ���˵ļ��ʵ����ϴ���ҽ��ǰ�÷�����
'������lng����ID=�Ƿ�ֻ�ϴ�������ָ�����˵ķ���
    Dim rsBill As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngPatiID As Long
    Dim lng��� As Long, str���� As String
    Dim str��ע As String
    Dim i As Long
    
    On Error GoTo errH
    
    '��ȡ������ϸ(ҽ����,˳���,�Ǽ�ʱ��,��Ŀ����,��Ŀ����,����,���,����,����,���,ҽ��,��������)
    '�����зǸ�ҽ���ķ��ò���,δ����ҽ������Ĳ���,��˳��ŵĲ���,Ӥ���Ѳ��ϴ�������������
    strSql = _
        "Select Nvl(A.�۸񸸺�,���) as ���," & _
        " A.����ID,F.ҽ����,F.˳���,A.�Ǽ�ʱ��,D.��Ŀ����,B.���� as ��Ŀ����, " & _
        " Decode(Instr(B.���,'��'),0,B.���,Substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
        " Decode(Instr(B.���,'��'),0,'',Substr(B.���,Instr(B.���,'��')+1)) as ����," & _
        " Avg(Nvl(A.����,1)*A.����) as ����,Sum(A.��׼����) as ����,Sum(A.ʵ�ս��) as ���," & _
        " A.������ as ҽ��,C.���� as ��������" & _
        " From ���˷��ü�¼ A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,������ҳ E,�����ʻ� F" & _
        " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID" & _
        " And A.����ID=E.����ID And A.��ҳID=E.��ҳID And A.����ID=F.����ID" & _
        " And F.˳��� is Not NULL And Nvl(A.Ӥ����,0)=0" & _
        " And D.����=" & gintInsure & " And E.����=" & gintInsure & " And F.����=" & gintInsure & _
        " And A.NO='" & strNO & "' And A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & _
        IIf(lng����ID = 0, "", " And A.����ID=" & lng����ID) & _
        " Group by Nvl(A.�۸񸸺�,���),A.����ID,F.ҽ����,F.˳���," & _
        " A.�Ǽ�ʱ��,D.��Ŀ����,B.����,B.���,A.������,C.����" & _
        " Order by ����ID,���"
    rsBill.CursorLocation = adUseClient
    rsBill.Open strSql, gcnOracle, adOpenKeyset, adLockReadOnly
    
    For i = 1 To rsBill.RecordCount
        '���ʵ����ж������,Ҫ�ֱ���
        If rsBill!����ID <> lngPatiID Then
            lngPatiID = rsBill!����ID
            
            '��ȡ�ò������ϴ���������
            strSql = "select max(convert(integer,xh)) as xh from zfymx where jysxh='" & rsBill!˳��� & "7' and grbm='" & rsBill!ҽ���� & "'"
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
            lng��� = 1
            If Not rsTmp.EOF Then lng��� = IIf(IsNull(rsTmp!xh), 0, rsTmp!xh) + 1
        End If
        
        '��ȡ�շѴ�������
        strSql = "select sfdlmc from sfxmdl where sfdlbm='" & Left(rsBill!��Ŀ����, 3) & "'"
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.Open strSql, gcnSybase, adOpenKeyset, adLockReadOnly
        str���� = ""
        If Not rsTmp.EOF Then str���� = rsTmp!sfdlmc
        
        '����zfymx,����ϸ�������������(z007)
        With rsBill
            If int״̬ = 1 Then
                str��ע = "����:" & strNO & ",���:" & !���
            Else
                str��ע = "����:" & strNO & ",���:" & !���
            End If
            strSql = _
                "insert into zfymx(" & _
                "jysxh,sfsj,pcno,grbm,sfdlbm,sfxmbm,sl,sjjg,cd,gg,yfyl,fyze,zfbl,txbz,bpbz,qzfbf,ggzfbf,yxbxbf," & _
                "fyshbz,sfy,jbr,bz,sfdlmc,sfxmmc,sjph,xh,yybm,ksbm,fylx,tjdm,ysxm,ksmc,blh,zyh) values (" & _
                "'" & !˳��� & "7','" & Format(!�Ǽ�ʱ��, "yyyy-MM-dd hh:mm:ss") & "'," & _
                "'" & UserInfo.վ�� & "','" & !ҽ���� & "','" & Left(!��Ŀ����, 3) & "','" & !��Ŀ���� & "'," & _
                Format(!����, "0.00") & "," & Format(!����, "0.00") & ",'" & IIf(IsNull(!����), "", !����) & "'," & _
                "'" & IIf(IsNull(!���), "", !���) & "',''," & Format(!���, "0.00") & ",0,'','',0,0,0,''," & _
                "'" & UserInfo.���� & "','" & UserInfo.���� & "','" & str��ע & "','" & str���� & "','" & !��Ŀ���� & "'," & _
                "'" & !˳��� & "7','" & lng��� & "','" & Trim(gstrҽԺ����) & "','','',''," & _
                "'" & IIf(IsNull(!ҽ��), "", !ҽ��) & "','" & !�������� & "','" & !����ID & "','" & !����ID & "')"
        End With
        gcnSybase.Execute strSql
        
        '������ϴ�
        strSql = "ZL_���˷��ü�¼_�ϴ�('" & strNO & "'," & rsBill!��� & "," & int���� & "," & int״̬ & ")"
        gcnOracle.Execute strSql, , adCmdStoredProc
        
        lng��� = lng��� + 1
        
        rsBill.MoveNext
    Next
    
    ���ʴ���_�ɶ� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function


