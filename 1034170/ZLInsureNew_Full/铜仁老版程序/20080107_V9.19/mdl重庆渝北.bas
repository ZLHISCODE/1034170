Attribute VB_Name = "mdl�����山"
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------------
'API��ҽ���ӿ�����
    Private Type Struct
        lngAppCode  As Long   '��־����ִ��״̬���롣����1ʱ��ʾ����ִ������������С��0ʱ��ʾ����ִ���쳣�����
        strErrMsg  As String  '������ִ��״̬����AppCodС��0ʱ����������ִ�е��쳣�������Ϣ��
    End Type
    
    '����API����
    '����:����Զ�����ݷ��񣬷���Զ�����ݷ���ķ������
    'Private Declare Function DataUpload Lib "YHMdcrDataUpldSvr.dll" Alias "_DataUpload@12" ( _
         strInputString As String, strOutPutstring As String, AppInfo As Struct) As Boolean
    '�½ӿ�
   ' Private Declare Function DataUpload Lib "YHMdcrDataUpldSvr.dll" Alias "_DataUpload@4" (ByVal strInputString As String) As as Boolean
 
    Private tmpStruct As Struct

    '�����������ݺ���
    '   strPerNo-���˱��
    '   strCardNO-����
    '   strExInfor-Ӧ��ִ����Ϣ
    Private Declare Sub srd_4428_info Lib "Mwic_32.dll" ( _
         ByVal strPerNO As String, ByVal strCardNO As String, ByVal strExInfor As String)


    
    '���ض�����Ϣ
    Private Declare Function ExportKB01 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKB01@8" (ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean
    
    '��ȡ������
    Private Declare Function GetAKC190 Lib "YHMdcrAsistntSvr.dll" Alias "_GetAKC190@12 (ByVal strYab003 As String, ByRef strAkc190 As String, ByRef tmpStrut As Struct) As Boolean" ()
    
    '��ȡ������
    Private Declare Function GetYKA105 Lib "YHMdcrAsistntSvr.dll" Alias "_GetYKA105@12" (ByVal strYab003 As String, ByRef strYka105 As String, ByRef tmpStrut As Struct) As Boolean
    
'-------------------------------------------------------------------------------------------------------------------------------------
'���õı�������
Public gcnOracle_CQYB       As New Connection        '���ӵ�oracle���ݿ�(�м��)


Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    ������� As String                      'ҽ�ƻ������
    ����������� As String
    ҽ���������� As String
    ҽ�ƻ������� As String
    ����״̬��ʶ As String
    �����޼�    As Double
    ������Ŀid As Long
    
    ������������ As Boolean
End Type
Public InitInfor_�����山 As InitbaseInfor

Public Enum ҵ������_�����山
    ��ȡϵͳʱ�� = 0
    ��ݼ���
    �޸�����
    IC���ʻ�֧��
    �ʸ����������˶�
    ������Ϣд��
    ������ϸд��
    ���������Ϣд��
    ������д��
    �˶��ʻ�֧����Ϣ
    �˶Ծ�����Ϣ
    �˶Դ�����ϸ��Ϣ
    �˶Է��ý�����
    �˶Է��ý��������Ϣ
    ����������ĿĿ¼
    ����ICD_10��Ϣ
    ��������Ŀ¼
    �������־��������Ϣ
    ����ҽ��������Ϣ
    ��ȡ�ͻ�����ʶ��
    ������������
    ��ȡ������
    ��ȡ������ˮ��
    ��ȡ������
    ���ý���
    ����������¼
End Enum
Public g�������_�����山 As �������
Private Type �������
    ���˱��            As String
    ����                As String
    ����                As String
    ����                As String
    �Ա�                As String
    ���֤��            As String
    ��������            As String
    ҽ����Ա���        As String
    ҽ���չ����        As String
    ҽ�Ʋ������        As String
    �籣���칹������    As String
    ��λ����            As Long
    ��λ����            As String
    ����                As Integer
    �ۼƽɷ�����        As Integer
    
    �ʻ����            As Double
    
    ����ID              As Long
    ���ֱ���            As String
    ֧�����            As String
    ������            As String
    ������㷽ʽ        As String       '��ʼʱ��ֵ,��Ҫ�ǵ�ǰֻ��һ�־�����㷽ʽ��:0-����Ŀ����
    ������            As String
    
    �����־            As Integer      '��ʾ��ǰΪ��ȡ�Ľ��㷽ʽ 0-����,1-סԺ,2-�Һ�,3-סԺ�����Ǽ�
    ����ID              As Long         '��ʾ��ǰ�Ľ���ID
    ����                As Boolean      '��ʾ��ǰ�Ƿ�Ϊ����
    �����ܶ�            As Double       '��ʾ��ǰ�����ܶ�
    ��;����            As Boolean      'ʾʾ��ǰ����Ϊ��;����
    ����ID              As Long
    �������            As Boolean      '��ǰ�Ƿ�Ϊ�������
    
    ��Ʊ��              As String       '��ǰ����ķ�Ʊ����
    ����ʱ��            As String       '��yyyy-mm-dd��ʽ
    
    lng����id           As Long
    lng��ҳID           As Long
    
    ������Ϣ            As String       '��ǰ�Ľ�����Ϣ,��Ҫ�������������
    ��������          As Double
 End Type
 
 
 Public Enum CodeType
    ҽ����Ա��� = 0
    ҽ���չ����
    ҽ�Ʋ������
 End Enum

'���Xml������
Private gobjXMLInPut As MSXML2.DOMDocument
Private gobjXMLOutput As MSXML2.DOMDocument
Private Const gstrXMLRootPart  As String = "XMLBODY"       '���ڵ�
Private gstrAppPath      As String
Private gobj���ý���   As Object
Public gobjYingHaiDll As Object                     '

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'���ú�����������

Public Function ҽ����ʼ��_�����山() As Boolean
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
        InitInfor_�����山.ģ������ = True
    Else
        InitInfor_�����山.ģ������ = False
    End If
    
    Call GetRegInFor(g����ģ��, "����", "������������", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�����山.������������ = True
    Else
        InitInfor_�����山.������������ = False
    End If
    InitInfor_�����山.������������ = InitInfor_�����山.������������ Or InitInfor_�����山.ģ������
    
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�����山
    Call OpenRecordset(rsTemp, "��ȡҽԺ����")
    InitInfor_�����山.ҽԺ���� = Nvl(rsTemp!ҽԺ����)

    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=" & TYPE_�����山
    Call OpenRecordset(rsTemp, "�山ҽ��")
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    
    '�޼�ȷ��
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������='������������' and ����=" & TYPE_�����山
    Call OpenRecordset(rsTemp, "�山ҽ��")
    If Not rsTemp.EOF Then
        InitInfor_�����山.�����޼� = Val(Nvl(rsTemp!����ֵ))
    Else
        InitInfor_�����山.�����޼� = 200
    End If
    '�޼�ȷ��
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������='�����ʻ�' and ����=" & TYPE_�����山
    Call OpenRecordset(rsTemp, "�山ҽ��")
    If Not rsTemp.EOF Then
        InitInfor_�����山.������Ŀid = Val(Nvl(rsTemp!����ֵ))
    Else
        InitInfor_�����山.������Ŀid = 0
    End If
    
    
    If OraDataOpen(gcnOracle_CQYB, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������㷽ʽ��ֵ��Ŀǰ��Ҫֻ��һ�ַ�ʽ:��:0-����Ŀ����
    g�������_�����山.������㷽ʽ = "0"
    
    '��ʼ����ҽ������ı��ļ�Ŀ¼
    gstrAppPath = App.Path & "\ҽ��"
    
    '������̬�ķ��ý������
    Err = 0
    On Error Resume Next
    Set gobj���ý��� = Nothing
    
    Set gobj���ý��� = CreateObject("PB80.n_yhmedicarereckon")
    
    If gobj���ý��� Is Nothing Or Err <> 0 Then
        If InitInfor_�����山.ģ������ Then
        Else
            ShowMsgbox "���ý��㲿���д���,����ҽ���ӿ�����ϵ."
            Exit Function
        End If
    End If
    
    Set gobjYingHaiDll = Nothing
    Set gobjYingHaiDll = CreateObject("PB80.n_dll_in")
    If gobjYingHaiDll Is Nothing Then
        ShowMsgbox "����ҽ���ӿڳ�������ҽ���ӿ��ṩ����ϵ!"
        Exit Function
    End If
    
    
    '������Ϣ
     Call ���ض���ҽ�ƻ���
    
    ҽ����ʼ��_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��ݱ�ʶ_�����山(Optional bytType As Byte, Optional lng����id As Long) As String
    Dim str��ע As String, rsPatient As New ADODB.Recordset
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    
    ��ݱ�ʶ_�����山 = frmIdentify�����山.GetPatient(bytType, lng����id)
    
End Function
Public Function ��ݱ�ʶ_�����山2(ByVal strCard As String, ByVal strPass As String, Optional lng����id As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    ��ݱ�ʶ_�����山2 = frmIdentify�����山.GetPatient(3, lng����id)
    
End Function

Private Function Get������Ϣ(ByVal lng����id As Long)
    Dim rsTemp As New ADODB.Recordset
    '�����ʻ�Ŀǰ���ֵ
    '--����id, ����, ����, ���ţ�ҽ������), ҽ����(���˱��), ����(֧����� ), ��Ա���(�α���Ա���ڵ��籣�����������), ��λ����(��λ����(��λ����)), ˳���(��),
    '--����֤��(ҽ����Ա���|ҽ���չ����|ҽ�Ʋ������|�ۼƽɷ�����), �ʻ����(�ʻ����), ��ǰ״̬, ����id������ID), ��ְ(1), �����(����), �Ҷȼ�, ����ʱ��
    Dim strTemp As String
    Dim strArr
    
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "select a.����,a.ҽ����,a.����,a.��Ա���,a.��λ����,b.������λ,a.˳���,a.����֤��,a.�ʻ����,a.��ǰ״̬,a.����id,a.��ְ,a.�����,a.�Ҷȼ�,a.����ʱ��," & _
             "        b.����,decode( b.�Ա�,'��','1','Ů','2','3') as �Ա�, b.����, b.��������, b.���֤��,A.������,A.������,A.֧����� " & _
             " from �����ʻ� a,������Ϣ b " & _
             " WHERE a.����id=" & lng����id & " AND a.����id=b.����id and a.����=" & TYPE_�����山
 
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
    
    With g�������_�����山
        .���� = Nvl(rsTemp!����)
        .���˱�� = Nvl(rsTemp!ҽ����)
        .���� = Nvl(rsTemp!����)
        .�Ա� = Nvl(rsTemp!�Ա�)
        .���� = Nvl(rsTemp!�����, 0)
        .�������� = Format(rsTemp!��������, "yyyy-mm-dd")
        .��λ���� = Val(Nvl(rsTemp!��λ����))
        
        strTemp = Nvl(rsTemp!������λ)
        If InStr(1, strTemp, "(") <> 0 Then
            .��λ���� = Mid(strTemp, 1, InStr(1, strTemp, "(") - 1)
        Else
            .��λ���� = strTemp
        End If
        
        .���� = Nvl(rsTemp!����)
        .֧����� = Nvl(rsTemp!֧�����)
        .�籣���칹������ = Nvl(rsTemp!��Ա���)
        strTemp = Nvl(rsTemp!����֤��, "|||")
        strTemp = IIf(strTemp = "", "|||", strTemp)
        strArr = Split(strTemp, "|")
        .ҽ����Ա��� = strArr(0)
        .ҽ���չ���� = strArr(1)
        .ҽ�Ʋ������ = strArr(2)
        .�ۼƽɷ����� = Val(strArr(3))
        .�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
        
        .���֤�� = Nvl(rsTemp!���֤��)
        .����ID = Nvl(rsTemp!����ID, 0)
        .������ = Nvl(rsTemp!������)
        .������ = Nvl(rsTemp!������)
        
        If .����ID <> 0 Then
           gstrSQL = "Select ���� From ҽ������Ŀ¼ where id=" & .����ID
           OpenRecordset_ZLYB rsTemp, "��ȡ����"
           If rsTemp.EOF Then
                .���ֱ��� = "00000"
           Else
                .���ֱ��� = Nvl(rsTemp!����, "000000")
           End If
        Else
            .���ֱ��� = "000000"
        End If
    End With
Exit Function
ErrHand:
        DebugTool "��ȡ������Ϣʧ��" & vbCrLf & " �����:" & Err.Number & vbCrLf & " ������Ϣ:" & Err.Description
End Function
Private Sub OpenRecordset_ZLYB(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSQL, strSql))
    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Private Function ���ض���ҽ�ƻ���() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ض���ҽ�ƻ������
    '--�����:
    '--������:
    '--��  ��:���سɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strXMLText As String
    Dim objStruct As Struct
    Dim blnTrue  As Boolean
    
    strFile = gstrAppPath & "\������Ϣ.txt"
    
    ���ض���ҽ�ƻ��� = False
    
    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(gstrAppPath) Then
        '�������ļ��У��贴��
        objFile.CreateFolder gstrAppPath
    End If
    
    objFile.CreateTextFile strFile, True
    
    DebugTool "����(" & "���ض���ҽ�ƻ���" & ")"
    objStruct.strErrMsg = Space(5000)
    Err = 0
    On Error GoTo ErrHand:
     '���ض���ҽ�ƹ�����Ϣ
     ExportKB01 strFile, objStruct
     If objStruct.lngAppCode < 0 Then
        ShowMsgbox "���ض���ҽ����Ϣ����"
     End If
    Set objText = objFile.OpenTextFile(strFile)
    '�洢���̲���:
    '����id, ��ҳid, ������, ������, �˵������, ������¼���, �����������, ҽ����Ա���, ҽ���չ����, ҽ�Ʋ������,
    '���, ���޽��, �Ը����, ֧�����, ����Ա����, �����Ը����, �ۼƽɷ�����, ʵ������, ҽ���������, �ʻ�֧��, �ֶα�׼,
    'ȫ�Էѽ��, �ҹ��Է�, �����Ը�, ����֧�����, ����Աͳ��֧��, �����Ը��ۼ�
    
    Call intXML
    blnTrue = False
    strXMLText = ""
    Do While Not objText.AtEndOfStream
        strXMLText = objText.ReadLine
        blnTrue = True
        Exit Do
    Loop
    If strXMLText = "" Then
        DebugTool "�ļ�������(���ض���ҽ�ƻ���)���ļ�:" & strFile
        Exit Function
    End If
    If GetXML��(strXMLText, False) = False Then
        DebugTool "XML��ʽ��Ч����ʽ:" & strXMLText
        Exit Function
    End If
    With InitInfor_�����山
        .����������� = GetXMLOutput("YAB003")
        .ҽ���������� = GetXMLOutput("AAB300")
        .ҽԺ���� = GetXMLOutput("AKB020")
        .ҽ�ƻ������� = GetXMLOutput("AKB021")
        .������� = GetXMLOutput("AKB023")
        .����״̬��ʶ = GetXMLOutput("YKB002")
    End With
    Exit Function
ErrHand:
  DebugTool "���ض���ҽ�ƻ�������(���ض���ҽ�ƻ���)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description

End Function
Public Function ҽ����ֹ_�����山() As Boolean
    If gcnOracle_CQYB.State = 1 Then
        gcnOracle_CQYB.Close
    End If
    Set gobjYingHaiDll = Nothing
    Set gobj���ý��� = Nothing
    
    ҽ����ֹ_�����山 = True
End Function
Public Function ��ݼ���_�����山() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:Զ����ݼ���
    '--�����:
    '--������:
    '--��  ��:�ɹ�true,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutPut As String
    Dim strXMLText As String
    Dim blnReturn As Boolean
    Err = 0
    On Error GoTo ErrHand:
    
    ��ݼ���_�����山 = False
    If intXML = False Then Exit Function
        
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "03"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g�������_�����山.����, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g�������_�����山.����, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    
    strXMLText = gobjXMLInPut.documentElement.xml
    'ȡ��ǰ��XML��
    strXMLText = ȡ��XML��ǰ����ʶ(strXMLText)
        
    'ҵ������
    
    blnReturn = ҵ������_�����山(��ݼ���, strXMLText, strOutPut)
    If blnReturn = False Then
        Exit Function
    End If
    
    '�����
    strXMLText = strOutPut
    
    '��ȡ����
    If GetXML��(strXMLText) = False Then
        ShowMsgbox "��ݼ��𷵻ش��Ǵ����XML��,���ܼ���!"
        Exit Function
    End If
    '�����ñ�����ֵ
    With g�������_�����山
        .���˱�� = GetXMLOutput("aac001")
        .���֤�� = GetXMLOutput("aac002")
        .���� = GetXMLOutput("aac003")
        .�Ա� = GetXMLOutput("aac004")
        .�������� = GetXMLOutput("aac006")
        .ҽ����Ա��� = GetXMLOutput("akc021")
        .ҽ���չ���� = GetXMLOutput("ykc120")
        .ҽ�Ʋ������ = GetXMLOutput("ykc121")
        .�籣���칹������ = GetXMLOutput("yab003")
        .��λ���� = Val(GetXMLOutput("aab001"))
        .��λ���� = GetXMLOutput("aab004")
        .�ۼƽɷ����� = Val(GetXMLOutput("ykc021"))
        .���� = Val(GetXMLOutput("akc023"))
        .�ʻ���� = Val(GetXMLOutput("ThisBaseICUsable")) + Val(GetXMLOutput("ThisOfficialICUsable"))
    End With
    
    ��ݼ���_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    ��ݼ���_�����山 = False
End Function
Private Function ȡ��XML��ǰ����ʶ(ByVal strXMLText As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:ȡ��XML��ǰ����ʶ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strXML As String
    
     strXML = Substr(strXMLText, Len(gstrXMLRootPart) + 3, LenBString(strXMLText) - Len(gstrXMLRootPart) * 2 - 5)
     If Right(strXML, 2) = "</" Then
        strXML = Mid(strXML, 1, Len(strXML) - 2)
     End If
    ȡ��XML��ǰ����ʶ = strXML
End Function
Private Function LenBString(ByVal strTxt As String) As Long
     LenBString = LenB(StrConv(strTxt, vbFromUnicode))
End Function

Private Function �ʸ���˴����˶�(ByVal lng����id As Long, ByVal str��ʼ����ʱ�� As String, ByVal str��������ʱ�� As String, Optional bln���������Ϣ As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���д����˶�
    '--�����:
    '--������:
    '--��  ��:��¼��
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText  As String
    Dim strOutPut As String
    
    If intXML = False Then Exit Function
          
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g�������_�����山.�籣���칹������, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "06"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g�������_�����山.����, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "ChkCardSymbol", 2
    
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g�������_�����山.����, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", Substr(g�������_�����山.������, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    AppendXMLNode gobjXMLInPut.documentElement, "aka123", IIf(g�������_�����山.����ID = 0, 0, 1)
    
    AppendXMLNode gobjXMLInPut.documentElement, "yka026", Substr(g�������_�����山.���ֱ���, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", Substr(g�������_�����山.֧�����, 1, 6)
    
    '-Ŀǰֻ��һ�־�����㷽ʽ,����
    AppendXMLNode gobjXMLInPut.documentElement, "yka222", Substr(g�������_�����山.������㷽ʽ, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", str��ʼ����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", str��������ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "SaveSymbol", IIf(bln���������Ϣ, 2, 1)
    
    strXMLText = gobjXMLInPut.documentElement.xml
    
    'ȡ��ǰ��XML��
    strXMLText = ȡ��XML��ǰ����ʶ(strXMLText)
        
       
    'ҵ������
    �ʸ���˴����˶� = ҵ������_�����山(�ʸ����������˶�, strXMLText, strOutPut)
    If �ʸ���˴����˶� = False Then
        Exit Function
    End If
    
    '�����
    strXMLText = strOutPut
    
    �ʸ���˴����˶� = False
    '��֤XML�Ƿ���ȷ
    If GetXML��(strXMLText) = False Then
        ShowMsgbox "�ʸ����������˶����ش��Ǵ����XML��,���ܼ���!"
        Exit Function
    End If
    
    �ʸ���˴����˶� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function Save������Ϣ(ByVal lng����id As Long, Optional bln������� As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����������Ϣ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim strXMLText As String
    Dim strFile As String   '���������ļ�
    Dim strText As String
    Dim strTemp As String
    Dim strҽ�ƻ������ As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    
    
    'aac001  Number  15  0   ���˱��
    'aae073  Number  15  0   �������
    'akb021  String  50      ����ҽ�Ʒ����������
    'akc190  String  20      ������
    'akb020  String  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    'ykb012  String  8       ת��ǰ����������
    'akb023  String  6       ҽ�ƻ�����𣬼������
    'aac002  String  18      ��ݺ���
    'aac003  String  20      ����
    'aac004  String  1       �Ա𣬼������
    'aac006  Date        ��  ��������
    'yab003  String  4       �α���Ա���ڵ��籣����������룬����λ��
    'aab001  Number  15  0   ��λ���
    'aab004  String  50      ��λ����
    'PastBaseICUsable    Number  14  2   ����ҽ������IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    'LastBaseICUsable    Number  14  2   ����ҽ������IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    'ThisBaseICUsable    Number  14  2   ����ҽ�Ʊ���IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    'NotPastBaseICUsable Number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻�����IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    'NotLastBaseICUsable Number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻�����IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    'NotThisBaseICUsable Number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻�����IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    'PastOfficialICUsable    Number  14  2   ����Ա����IC����Ȧ���ſ�ģʽ�µ����˻���ʵ���)
    'LastOfficialICUsable    Number  14  2   ����Ա����IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    'ThisOfficialICUsable    Number  14  2   ����Ա����IC����Ȧ���(�ſ�ģʽ�µ����˻���ʵ���)
    
    
    'ykc114  Number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
    'ykc007  String  1       �Ƿ�����µĽ����׼��'0' ����Ҫ��'1' ��Ҫ
    'akc021  String  6       ҽ����Ա��𣬼������
    'ykc021  Number  3       �ۼƽɷ�����
    'akc023  Number  3       ʵ������
    'yka114  Number  14  2   �𸶱�׼
    'yka115  Number  14  2   ��������
    'yka116  Number  14  2   ����֧���ۼ�
    'yka117  Number  14  2   ��������ɲ����޶�
    'yka118  Number  14  2   ����֧���ۼ�
    'yka203  Number  14  2   ���λ���ҽ��֧���޶��׼
    'yka119  Number  14  2   ���λ���ҽ��֧���޶�
    'yka120  Number  14  2   ���λ���ҽ�ƽ���ͳ���ۼ�
    'yka204  Number  14  2   ���δ��ҽ���޶��׼
    'yka121  Number  14  2   ���δ��ҽ��֧���޶�
    'yka122  Number  14  2   ���δ��ҽ�ƽ���ͳ���ۼ�
    'yka123  Number  14  2   ���ι���Ա֧���޶�
    'yka124  Number  14  2   ���ι���Ա����ͳ���ۼ�
    'ykc008  String  4000        ��������Ϣ
    'ykc022  Number  3   0   ���������ۼ�סԺ����
    'ykc006  Number  3   0   ����ͳ���ۼ�סԺ����
    'ykc141  Number  14  2   ���ν���˻��޶�
    'ykc142  Number  14  2   �������˻�֧���ۼ�
    'yka125  Number  14  2   ͳ���֧�����������Ը������ۼ�
    'yka126  number  14  2   ͳ�ﲻ��֧�����������Ը������ۼ�
    'ykc120  string  6       ҽ���չ���𣬼������
    'ykc121  string  6       ����ҽ�Ʋ�����𣬼������
    'yka273  number  14  2   ������������֧���޶��׼
    'yka274  number  14  2   ������������֧���޶�
    'yka275  number  14  2   ���������ܶ�
    'akc315  string  6       ҽ�ƴ���������𣬼������
    'ykc054  number  14  2   ���������������ҽ��֧���ۼ�
    
    
    '���̲���:
    '   �������, ����id, ���˱��, �����������, ������, ҽԺ���, ����������, ҽ�ƻ������, �����������, ��λ���, ��λ����,
    '   �ʻ����, ������¼��, �½����׼, ҽ����Ա���, �ۼƽɷ�����, ʵ������, �𸶱�׼, ��������, ����֧���ۼ�, ���ﲹ���޶�, ����֧���ۼ�, ����֧����׼,
    '   ����֧���޶�, ���������ۼ�, ����޶��׼, ���֧���޶�, �������ۼ�, ����Ա֧���޶�, ����Ա�����ۼ�, ��������Ϣ, �����ۼƴ���, ͳ���ۼƴ���,
    '   ����ʻ��޶�, ����ʻ��ۼ�, ��֧�Ը��ۼ�, ����֧�Ը��ۼ�, �չ����, �������, ���������׼, ���������޶�, ���������ܶ�, �����������,
    '   ��������֧���ۼ�
    
    '������Ϣ����
    strFile = gstrAppPath & "\����������Ϣ.txt"
    
    If Not objFile.FolderExists(gstrAppPath) Then
        '�������ļ��У��贴��
        objFile.CreateFolder gstrAppPath
    End If
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile, True
    End If
    Set objText = objFile.OpenTextFile(strFile, ForWriting)
    
    Err = 0
    On Error GoTo ErrHand:
    
    g�������_�����山.�ʻ���� = Val(GetXMLOutput("ThisBaseICUsable")) + Val(GetXMLOutput("ThisOfficialICUsable"))
    g�������_�����山.������ = GetXMLOutput("akc190")
    
    
    
    strHead = "ZL_�ʸ����������˶�_INSERT("
    'aae073  Number  15  0   �������
    strHead = strHead & Val(GetXMLOutput("aae073")) & ","
    strHead = strHead & lng����id & ",'"
    'aac001  Number  15  0   ���˱��
    strHead = strHead & GetXMLOutput("aac001") & "','"
    'akb021  String  50      ����ҽ�Ʒ����������
    strHead = strHead & GetXMLOutput("akb021") & "','"
    'akc190  String  20      ������
    strHead = strHead & GetXMLOutput("akc190") & "','"
    'akb020  String  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    strHead = strHead & GetXMLOutput("akb020") & "','"
    'ykb012  String  8       ת��ǰ����������
    strHead = strHead & GetXMLOutput("ykb012") & "','"
    'akb023  String  6       ҽ�ƻ�����𣬼������
    strTemp = GetXMLOutput("akb023")
    strHead = strHead & strTemp & "','"
    strҽ�ƻ������ = strTemp
    'yab003  String  4       �α���Ա���ڵ��籣����������룬����λ��
    strHead = strHead & GetXMLOutput("yab003") & "','"
    'aab001  Number  15  0   ��λ���
    strHead = strHead & GetXMLOutput("aab001") & "','"
    'aab004  String  50      ��λ����
    strHead = strHead & GetXMLOutput("aab004") & "',"
    
    strHead = strHead & g�������_�����山.�ʻ���� & ","
    
    '�����¼����
    Dim lngCount As Long
    Dim lngRow As Long
    lngCount = GetOutXMLRows("ykc114")
    For lngRow = 0 To lngCount - 1
        gstrSQL = ""
        strText = ""
        'ykc114  Number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
        strTemp = GetXMLOutput("ykc114", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        strText = strText & g�������_�����山.������ & vbTab
        strText = strText & g�������_�����山.֧����� & vbTab
        strText = strText & GetXMLOutput("akc021", , lngRow) & vbTab
        strText = strText & g�������_�����山.ҽ���չ���� & vbTab
        strText = strText & g�������_�����山.ҽ�Ʋ������ & vbTab
        strText = strText & strҽ�ƻ������ & vbTab
        '--����ȱ�־
        strText = strText & "0" & vbTab
        strText = strText & g�������_�����山.���ֱ��� & vbTab
                
        
        'ykc007  String  1       �Ƿ�����µĽ����׼��'0' ����Ҫ��'1' ��Ҫ
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc007", , lngRow)) & ",'"
        'akc021  String  6       ҽ����Ա��𣬼������
        gstrSQL = gstrSQL & GetXMLOutput("akc021", , lngRow) & "',"
        'ykc021  Number  3       �ۼƽɷ�����
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc021", , lngRow)) & ","
        'akc023  Number  3       ʵ������
        strTemp = GetXMLOutput("akc023", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        strText = strText & Val(GetXMLOutput("ykc021", , lngRow)) & vbTab   '�ۼƽɷ�����
        strText = strText & Val(GetXMLOutput("ykc006", , lngRow)) & vbTab    '����ͳ��֧���ۼ�סԺ����
        
        'yka114  Number  14  2   �𸶱�׼
        strTemp = GetXMLOutput("yka114", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        'yka115  Number  14  2   ��������

        strTemp = GetXMLOutput("yka115", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        g�������_�����山.�������� = Val(strTemp)
        
        'yka116  Number  14  2   ����֧���ۼ�
        strTemp = GetXMLOutput("yka116", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka117  Number  14  2   ��������ɲ����޶�
        strTemp = GetXMLOutput("yka117", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka118  Number  14  2   ����֧���ۼ�
        strTemp = GetXMLOutput("yka118", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        strTemp = GetXMLOutput("ykc141", , lngRow)  'ykc141  Number  14  2   ���ν���˻��޶�
        strText = strText & Val(strTemp) & vbTab
        strTemp = GetXMLOutput("ykc142", , lngRow) 'ykc142  Number  14  2   �������˻�֧���ۼ�
        strText = strText & Val(strTemp) & vbTab
        
        'yka203  Number  14  2   ���λ���ҽ��֧���޶��׼
        strTemp = GetXMLOutput("yka203", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka119  Number  14  2   ���λ���ҽ��֧���޶�
        strTemp = GetXMLOutput("yka119", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka120  Number  14  2   ���λ���ҽ�ƽ���ͳ���ۼ�
        strTemp = GetXMLOutput("yka120", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka204  Number  14  2   ���δ��ҽ���޶��׼
        strTemp = GetXMLOutput("yka204", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka121  Number  14  2   ���δ��ҽ��֧���޶�
        strTemp = GetXMLOutput("yka121", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka122  Number  14  2   ���δ��ҽ�ƽ���ͳ���ۼ�
        strTemp = GetXMLOutput("yka122", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka123  Number  14  2   ���ι���Ա֧���޶�
        strTemp = GetXMLOutput("yka123", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka124  Number  14  2   ���ι���Ա����ͳ���ۼ�
        strTemp = GetXMLOutput("yka124", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ",'"
        strText = strText & Val(strTemp) & vbTab
        
        'yka125  Number  14  2   ͳ���֧�����������Ը������ۼ�
        strTemp = GetXMLOutput("yka125", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        
        'yka126  number  14  2   ͳ�ﲻ��֧�����������Ը������ۼ�
        strTemp = GetXMLOutput("yka126", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        
        
        '1       number  14  2   ����ҽ�Ʊ����˻���֧����
         strTemp = GetXMLOutput("ThisBaseICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        '2       number  14  2   ����ҽ�������˻���֧����
        '3       number  14  2   ����ҽ�������˻���֧����
        strText = strText & 0 & vbTab
        strText = strText & 0 & vbTab
        
        '4       number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻������֧�����
        '5       number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻������֧�����
        '6       number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻������֧�����
        strText = strText & 0 & vbTab
        strText = strText & 0 & vbTab
        strText = strText & 0 & vbTab
        '7       number  14  2   ����Ա���������˻���֧����
        strTemp = GetXMLOutput("ThisOfficialICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        '8       number  14  2   ����Ա���������˻���֧����
        '9       number  14  2   ����Ա���������˻���֧����
        strText = strText & 0 & vbTab
        strText = strText & 0 & vbTab
 

        strText = strText & g�������_�����山.�籣���칹������ & vbTab

        
        'ykc008  String  4000        ��������Ϣ
        strTemp = GetXMLOutput("ykc008", , lngRow)
        strText = strText & strTemp & vbTab
        gstrSQL = gstrSQL & strTemp & "',"
        
        
        
        'ykc022  Number  3   0   ���������ۼ�סԺ����
        
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc022", , lngRow)) & ","
        'ykc006  Number  3   0   ����ͳ���ۼ�סԺ����
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc006", , lngRow)) & ","
        
        'ykc141  Number  14  2   ���ν���˻��޶�
        strTemp = GetXMLOutput("ykc141", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        'ykc142  Number  14  2   �������˻�֧���ۼ�
        strTemp = GetXMLOutput("ykc142", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        
        'yka125  Number  14  2   ͳ���֧�����������Ը������ۼ�
        strTemp = GetXMLOutput("yka125", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        
        'yka126  number  14  2   ͳ�ﲻ��֧�����������Ը������ۼ�
        strTemp = GetXMLOutput("yka126", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ",'"
        
        
        'ykc120  string  6       ҽ���չ���𣬼������
        gstrSQL = gstrSQL & GetXMLOutput("ykc120", , lngRow) & "','"
        
        'ykc121  string  6       ����ҽ�Ʋ�����𣬼������
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc121", , lngRow)) & "',"
        'yka273  number  14  2   ������������֧���޶��׼
        strTemp = GetXMLOutput("yka273", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka274  number  14  2   ������������֧���޶�
        strTemp = GetXMLOutput("yka274", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka275  number  14  2   ���������ܶ�
        strTemp = GetXMLOutput("yka275", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ",'"
        strText = strText & Val(strTemp) & vbTab
        
        'akc315  string  6       ҽ�ƴ���������𣬼������
        strTemp = GetXMLOutput("akc315", , lngRow)
        gstrSQL = gstrSQL & strTemp & "',"
        strText = strText & strTemp & vbTab
        
        'ykc054  number  14  2   ���������������ҽ��֧���ۼ�
        strTemp = GetXMLOutput("ykc054", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ")"
        strText = strText & Val(strTemp) & vbTab
        
        '���������
        gstrSQL = strHead & gstrSQL
        If bln������� = False Then
            '�ڹ���ǰ�õ�������
            gcnOracle_CQYB.Execute gstrSQL, , adCmdStoredProc
        End If
        '���������ļ�
        objText.WriteLine strText
    Next
    Save������Ϣ = True
    objText.Close
    Exit Function
ErrHand:
   DebugTool "������Ϣ�������(Save������Ϣ)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
    objText.Close
End Function
Private Function GetOutXMLRows(ByVal strName As String) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡXML����
    '--�����:
    '--������:
    '--��  ��:����
    '-----------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    Err = 0
    On Error Resume Next
    GetOutXMLRows = gobjXMLOutput.getElementsByTagName(strName).Length
    If Err <> 0 Then
        strErrMsg = "�������:" & vbCrLf & "   " & Err.Description
    End If
    DebugTool "��ȡXML�ļ�¼����(GetOutXMLRows)�� " & strName & "��" & vbCrLf & strErrMsg
End Function
Private Function IsertIntoҽ����ϸ(ByVal lng����ID As Long, ByVal strNO As String, ByVal lng��� As Long, ByVal lng��¼���� As Long, ByVal strCode As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�����¼
    '--�����:
    '--������:strCode-������Ŀ����(��Ҫ���ڹҺ�)
    '--��  ��:�����¼
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    If g�������_�����山.���� Then
        gstrSQL = "Select ID From ���˷��ü�¼ where no='" & strNO & "' and ��¼����=" & lng��¼���� & " and ��¼״̬=3 and ���=" & lng���
        OpenRecordset rsTemp, "�˵���ˮ��"
        If rsTemp.EOF Then
            IsertIntoҽ����ϸ = False
            Exit Function
        End If
        gstrSQL = "Select ������ˮ��,������,������ְ��,������־,��Ŀ���� From ҽ����ϸ���� where ����id= " & Nvl(rsTemp!ID)
        OpenRecordset_ZLYB rsTemp, "�˵����ݲ���"
        '--ZL_ҽ����ϸ����_INSERT(
            '����ID_IN IN ҽ����ϸ����.����ID%TYPE,
            '������_IN IN ҽ����ϸ����.������%TYPE,
            '������ְ��_IN IN ҽ����ϸ����.������ְ��%TYPE,
            '������־_IN IN ҽ����ϸ����.������־%TYPE,
            '������_IN IN ҽ����ϸ����.������%TYPE,
            '������_IN IN ҽ����ϸ����.������%TYPE,
            '�˵���ˮ��_IN   IN ҽ����ϸ����.�˵���ˮ��%type
            ')
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "ԭʼ���õ��ݲ�����,��˲�!"
            Exit Function
        End If
        gstrSQL = "ZL_ҽ����ϸ����_INSERT(" & _
         lng����ID & ",'" & _
         Nvl(rsTemp!������) & "','" & _
         Nvl(rsTemp!������ְ��) & "'," & _
         Nvl(rsTemp!������־, 1) & ",'" & _
         g�������_�����山.������ & "','" & _
         g�������_�����山.������ & "'," & _
         Nvl(rsTemp!������ˮ��, 0) & ",'" & _
         Nvl(rsTemp!��Ŀ����) & "')"
         
         
    Else
        gstrSQL = "ZL_ҽ����ϸ����_INSERT(" & _
         lng����ID & ",'" & _
         "" & "','" & _
         "" & "'," & _
         1 & ",'" & _
         g�������_�����山.������ & "','" & _
         g�������_�����山.������ & "'," & _
           "NULL" & ",'" & _
           strCode & "')"
    End If
    Err = 0
    On Error GoTo ErrHand:
    Call SQLTest(App.ProductName, "����ҽ����ϸ����", gstrSQL)
    gcnOracle_CQYB.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
    IsertIntoҽ����ϸ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    IsertIntoҽ����ϸ = False
End Function
Private Function Saveҽ����ϸ����(ByVal rs��ϸ As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp  As String
    Dim strCode As String
    Err = 0
    On Error GoTo ErrHand:
    g�������_�����山.�����ܶ� = 0
    With rs��ϸ
         .MoveFirst
         Do While Not .EOF
                 If Nvl(!��Ŀ����) = "" Then
                     ShowMsgbox "����δ����ҽ����Ŀ,���ڱ�����Ŀ��������Ӧ�Ķ�Ӧ��ϵ!"
                     Exit Function
                 End If
                 If g�������_�����山.�����־ = 3 Then
                    'ȡ��ǰ�Ľ�����
                    g�������_�����山.������ = Nvl(!�����)
                    g�������_�����山.������ = Nvl(!������)
                 End If
                 If g�������_�����山.���� Then
                        If IsertIntoҽ����ϸ(!ID, Nvl(!NO), Nvl(!���, 0), Nvl(!��¼����, 0), "") = False Then Exit Function
                 Else
                    If g�������_�����山.�����־ = 2 And InitInfor_�����山.������Ŀid <> 0 Then
                       If Nvl(!������Ŀid, 0) = InitInfor_�����山.������Ŀid Then
                          If frm������Ŀ����ѡ��.ShowCard(strCode) = False Then Exit Function
                          
                       End If
                    End If
                     '����ҩƷ��ȷ����۸��������
                     If Nvl(!ʵ�ʼ۸�, 0) > InitInfor_�����山.�����޼� Then
                         strTemp = frm�޼�����_�山.Get������Ϣ(!ID, strCode)
                     Else
                         IsertIntoҽ����ϸ !ID, Nvl(!NO), Nvl(!���, 0), Nvl(!��¼����, 0), strCode
                     End If
                     
                 End If
             g�������_�����山.�����ܶ� = g�������_�����山.�����ܶ� + Nvl(!ʵ�ս��, 0)
             .MoveNext
         Loop
     End With
     Saveҽ����ϸ���� = True

    Exit Function
ErrHand:
  DebugTool "����ҽ����ϸ����д��(Saveҽ����ϸ����)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
End Function
Private Function Get������־(ByVal lng����ID As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������־
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ������ˮ��,������,������ְ��,������־ From ҽ����ϸ���� where ����id= " & lng����ID
    OpenRecordset_ZLYB rsTemp, "��ȡ������־"
    If rsTemp.EOF Then
        Get������־ = 0
    Else
        Get������־ = Nvl(rsTemp!������־, 0)
    End If
End Function
Private Function Get������Ŀ����(ByVal lng����ID As Long)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ��Ŀ���� From ҽ����ϸ���� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ŀ����"
    If rsTemp.EOF Then
        Get������Ŀ���� = ""
    Else
        Get������Ŀ���� = Nvl(rsTemp!��Ŀ����)
    End If
End Function
Private Function Save������ϸ�ı��ļ�(ByVal rs��ϸ As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ϸ���ı��ļ�
    '--�����:
    '--������:
    '--��  ��:�����ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp��ϸ As New ADODB.Recordset
    Dim strFile As String
    
    strFile = gstrAppPath & "\������ϸ��Ϣ.txt"
    
    Save������ϸ�ı��ļ� = False
    
    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(gstrAppPath) Then
        '�������ļ��У��贴��
        objFile.CreateFolder gstrAppPath
    End If
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile, True
    End If
    Set objText = objFile.OpenTextFile(strFile, ForWriting)
    
    
    If rs��ϸ Is Nothing Then Exit Function
    
    Dim byt���� As Byte
    Dim lng��ˮ�� As Long
    
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
                
                Set rsTemp = Get������Ŀ(Nvl(!��Ŀ����))
                Set rsTmp��ϸ = Getҽ����ϸ����(!ID)
                If rsTemp.RecordCount = 0 Then
                    ShowMsgbox "����ҽ����Ŀ,��˲�!"
                    Exit Function
                End If
                '����ҩƷ��ȷ����۸��������
                strText = Nvl(!������) & vbTab
                strText = strText & Nvl(rsTmp��ϸ!������ˮ��, 0) & vbTab
                     
                If Nvl(rsTemp!������־, 0) = 1 Then
                    strText = strText & "59000000000010000" & vbTab  'ΪҽԺ�Լ������ı���,�̶�����
                Else
                    strText = strText & Nvl(rsTemp!ҽ������) & vbTab
                End If
                
                If g�������_�����山.�����־ = 2 And Nvl(!������Ŀid) = InitInfor_�����山.������Ŀid Then
                    
                    strText = strText & Nvl(rsTmp��ϸ!��Ŀ����) & vbTab
                    strText = strText & Nvl(rsTmp��ϸ!��Ŀ����) & vbTab
                Else
                    strText = strText & Nvl(!��Ŀ����) & vbTab
                    If Nvl(rsTemp!������־, 0) = 1 Then
                        
                        strText = strText & Nvl(rsTemp!��׼���) & vbTab
                    Else
                        strText = strText & Nvl(!��Ŀ����) & vbTab
                    End If
                End If
                strText = strText & 1 & vbTab       'Ŀǰ�ñ���ֻתֵΪ1
                strText = strText & Nvl(rsTmp��ϸ!������) & vbTab
                strText = strText & Nvl(rsTmp��ϸ!�˵���ˮ��) & vbTab
                strText = strText & Nvl(!ʵ�ʼ۸�, 0) & vbTab
                strText = strText & Nvl(!����, 0) & vbTab
                strText = strText & Nvl(!ʵ�ս��, 0) & vbTab
                strText = strText & Nvl(rsTmp��ϸ!������־, 0) & vbTab
                strText = strText & Nvl(!�����������) & vbTab
                strText = strText & Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS") & vbTab
                If Not rsTemp.EOF Then
                    strText = strText & Nvl(rsTemp!Ŀ¼����)
                Else
                    strText = strText & ""
                End If
                objText.WriteLine strText
            .MoveNext
        Loop
    End With
    objText.Close
    Save������ϸ�ı��ļ� = True
    Exit Function
ErrHand:
    Resume
     DebugTool "��ϸ��Ϣ�������(Save������ϸ�ı��ļ�)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
    objText.Close
End Function

Private Function Get������Ŀ(ByVal str��Ŀ��� As String) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select Ŀ¼����,������־,ҽ������,����,��׼���,��Ʒ����,��Ʒ��,�Ը�����1 as �Ը�����  From ҽ��������ĿĿ¼ where ��Ʒ����='" & str��Ŀ��� & "'"
    With rsTemp
        .Open gstrSQL, gcnOracle_CQYB
    End With
    Set Get������Ŀ = rsTemp
End Function
Private Function Get������ˮ��(ByVal lng����ID As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������ˮ��
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ������ˮ�� From ҽ����ϸ���� where ����ID=" & lng����ID
    
    Call SQLTest(App.ProductName, "��ȡ������ˮ��", gstrSQL)
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle_CQYB
    Call SQLTest
    If rsTemp.EOF Then
        Get������ˮ�� = 0
    Else
        Get������ˮ�� = Nvl(rsTemp!������ˮ��, 0)
    End If
End Function
Private Function Getҽ����ϸ����(ByVal lng����ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������ˮ��
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    DebugTool "����(" & "Getҽ����ϸ����" & ")"
    
    Err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "Select * From ҽ����ϸ���� where ����ID=" & lng����ID
    
    Call SQLTest(App.ProductName, "��ȡҽ����ϸ����", gstrSQL)
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle_CQYB
    Call SQLTest
    Set Getҽ����ϸ���� = rsTemp
    
    Exit Function
ErrHand:
  DebugTool "��ȡҽ����ϸ���ó���(Getҽ����ϸ����)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
End Function

Private Function Save��ʷ���ý������ı�(ByVal lng����id As Long, ByVal lng��ҳID As Long, Optional bln���� As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ϸ���ı��ļ�
    '--�����:
    '--������:
    '--��  ��:�����ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim rsTemp As New ADODB.Recordset
    Dim strFile As String
    
    strFile = gstrAppPath & "\���ν�����Ϣ.txt"
    
    Save��ʷ���ý������ı� = False
    
    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(gstrAppPath) Then
        '�������ļ��У��贴��
        objFile.CreateFolder gstrAppPath
    End If
    objFile.CreateTextFile strFile, True
    Set objText = objFile.OpenTextFile(strFile, ForWriting)
     
    If bln���� Then
        '����ֻ��һ�����ļ�
        Save��ʷ���ý������ı� = True
        Exit Function
    End If
    
    
    gstrSQL = "Select * From ���ý����� where ������='" & g�������_�����山.������ & "' and ����id=" & lng����id & " order by ������ "
    
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnOracle_CQYB, adOpenStatic
        Do While Not .EOF
                    strText = Nvl(!������) & vbTab
                    strText = strText & Nvl(!������) & vbTab
                    strText = strText & Nvl(!�˵������) & vbTab
                    strText = strText & Nvl(!������¼���) & vbTab
                    strText = strText & Nvl(!�����������) & vbTab
                    strText = strText & Nvl(!���) & vbTab
                    strText = strText & Nvl(!�ֶα�׼) & vbTab
                    strText = strText & Nvl(!ȫ�Էѽ��, 0) & vbTab
                    strText = strText & Nvl(!�ҹ��Է�, 0) & vbTab
                    strText = strText & Nvl(!���Ͻ��, 0) & vbTab
                    strText = strText & Nvl(!�����Ը�, 0) & vbTab
                    strText = strText & Nvl(!����֧�����, 0) & vbTab
                    strText = strText & Nvl(!����Աͳ��֧��, 0) & vbTab
                    strText = strText & Nvl(!�����Ը��ۼ�, 0) & vbTab
                    objText.WriteLine strText
                .MoveNext
            Loop
    End With
    objText.Close
    Save��ʷ���ý������ı� = True
    Exit Function
ErrHand:
    DebugTool "��ʷ���ý�������ѯ����(Save��ʷ���ý�����)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
    objText.Close
End Function
Private Function ���ý���ֽ�(ByVal strFile As String, ByVal lng����id As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ý���ֽ�
    '--�����:
    '--������:
    '--��  ��:�����ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strText As String
    Dim strXMLText As String
    Dim strOutPut As String
    Dim blnFirst As Boolean
    Dim objXMLItem As MSXML2.IXMLDOMElement
    Dim strXMLtext1 As String
    Dim dblTmp(0 To 10) As Double
    Dim dblSumMony(0 To 10) As Double
    Dim dblSumSubMony(0 To 10) As Double
    Dim dblSumSubmony1(0 To 10) As Double
    Dim rsTemp As New ADODB.Recordset
    Dim str����ʱ�� As String
    Dim i As Long
    
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    '0-ȫ�Է�:decode(�ֶα�׼,'02',yka056+yka057,0)
    '1-���Ը�����:decode(�ֶα�׼,'05',yka057,'06',yka057,0)
    '2-���Ͻ��:decode(�ֶα�׼,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
    '3-�����ܶ�:decode(�ֶα�׼,'02','yka056,'02','yka053+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
    '4-��������:decode(�ֶα�׼,'03',yka106+yka057,0)
    '5-����ҽ���Ը�:decode(�ֶα�׼,'05',yka106,0)
    '6-����ҽ��ͳ��֧��:decode(�ֶα�׼,'05',yka107,0)
    '7-����Ը�:decode(�ֶα�׼,'06','yka106,0)
    '8-���֧��:decode(�ֶα�׼,'06','yka107
    '����
    '   9-����Ա����:decode(�ֶα�׼,'05',yka063,'07',yka063,0)
    
    'סԺ
    '   9-����Ա����:decode(�ֶα�׼,'07',yka063,0)
    '10-�����Ը�:decode(�ֶα�׼,'08',yka106+yka057,0)
    
    Dim strTemp As String
    Dim strValues As String
    Dim strvalues1 As String
    Dim strArr
    Dim str������Ϣ As String
    Dim str������¼��  As String
    
    Dim bytType As Byte         '1-������Ϣд��ʧ��,2-���ý��д��ʧ��,3-�۸����ʻ�ʧ��
    
    DebugTool "���ý���ֽ�!"
    
    Err = 0
    On Error GoTo ErrHand:
    
    Set objText = objFile.OpenTextFile(strFile)
    
    '�洢���̲���:
    'ID,����id, ��ҳid, ������, ������, �˵������, ������¼���, �����������, ҽ����Ա���, ҽ���չ����, ҽ�Ʋ������,
    '���, ���޽��, �Ը����, ֧�����, ����Ա����, �����Ը����, �ۼƽɷ�����, ʵ������, ҽ���������, �ʻ�֧��, �ֶα�׼,
    'ȫ�Էѽ��, �ҹ��Է�, �����Ը�, ����֧�����, ����Աͳ��֧��, �����Ը��ۼ�
    
    Call intXML
    blnFirst = True
    For i = 0 To 10
        dblSumMony(i) = 0
        dblSumSubMony(i) = 0
        dblSumSubmony1(i) = 0
        dblTmp(i) = 0
    Next
    Dim lngID As Long
    
    Do While Not objText.AtEndOfStream
            
        gstrSQL = "Select ���ý�����_ID.nextval as ID from dual"
        OpenRecordset_ZLYB rsTemp, "��ȡ������"
        lngID = Nvl(rsTemp!ID, 0)
        
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        
        strSql = "ZL_���ý�����_INSERT("
        
        strSql = strSql & lngID & ","
        strSql = strSql & lng����id & ","
        strSql = strSql & IIf(lng��ҳID = 0, "Null", lng��ҳID) & ","
        strSql = strSql & "'" & strArr(0) & "',"
        strSql = strSql & "'" & strArr(1) & "',"
        strSql = strSql & "'" & strArr(2) & "',"
        strSql = strSql & "" & Val(strArr(3)) & ","
        strSql = strSql & "'" & strArr(4) & "',"
        strSql = strSql & "'" & strArr(5) & "',"
        strSql = strSql & "'" & strArr(6) & "',"
        strSql = strSql & "'" & strArr(7) & "',"  'ҽ�Ʋ������
        strSql = strSql & "'" & strArr(8) & "',"
        
        '10��ΪXMLֵ,��ֽ�ֵ
        strXMLText = strArr(9)
        str������Ϣ = strXMLText
        'GKC010  string  800     ����ҽ���Ӷ���Ϣ���������ŵ�Ԫ����������Ԫ��
        'SubRkn �������ŵ�Ԫ����������Ԫ��
        '    AKA160  number  14  2   �Ӷ����޽��
        '    YKA106  number  14  2   �Ը����
        '    YKA 107 number  14  2   ֧�����
        '    YKA 063 number  14  2   ����Ա�������
        '    YKA057  number  14  2   �����Ը�����
        
        strSql = strSql & "'" & strArr(9) & "',"
        strSql = strSql & 0 & ","
        strSql = strSql & 0 & ","
        strSql = strSql & 0 & ","
        strSql = strSql & 0 & ","
        strSql = strSql & 0 & ","
        
        '�ۼƽɷ�����
        strSql = strSql & "" & Val(strArr(10)) & ","
        strSql = strSql & "" & Val(strArr(11)) & ","
        strSql = strSql & "'" & strArr(12) & "',"
        
        strSql = strSql & "" & Val(strArr(13)) & ","
        '�ֶα�׼
        strSql = strSql & "'" & strArr(22) & "',"
        
        strSql = strSql & "" & Val(strArr(23)) & ","
        
        strSql = strSql & "" & Val(strArr(24)) & ","
        strSql = strSql & "" & Val(strArr(25)) & ","
        strSql = strSql & "" & Val(strArr(26)) & ","
        strSql = strSql & "" & Val(strArr(27)) & ","
        strSql = strSql & "" & Val(strArr(28)) & ","
        strSql = strSql & "" & Val(strArr(29)) & ")"
        
        If g�������_�����山.������� Then
            '��������ò��ű����������
        Else
            '�������ݿ���
            gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
            If insertInto����(lngID, str������Ϣ) = False Then
                DebugTool "��������Ŀ����!"
            End If
        End If
        
        'XML���ý��д��
        If blnFirst Then
            AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g�������_�����山.�籣���칹������, 1, 4)
            AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "12"
            AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    
            'BaseInfo                ����������ܶι��е���ͬ�Ļ�����Ϣ���֣��������ŵ�Ԫ����������Ԫ��
            Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "BaseInfo", "")
            '    akc190  string  20      ������
            AppendXMLNode objXMLItem, "akc190", strArr(0)
            '    yka103  string  20      ������
            AppendXMLNode objXMLItem, "yka103", strArr(1)
            '    yka198  string  20      �˵���Ӧ������
            AppendXMLNode objXMLItem, "yka198", strArr(2)
            '    ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
            AppendXMLNode objXMLItem, "ykc114", strArr(3)
            '    yab003  string  4       �籣�����������
            AppendXMLNode objXMLItem, "yab003", strArr(4)
            strValues = ""
            
            strValues = strValues & strArr(0) & vbTab
            strValues = strValues & strArr(1) & vbTab
            strValues = strValues & strArr(2) & vbTab
            strValues = strValues & Val(strArr(3)) & vbTab
            str������¼�� = strArr(3)
            strValues = strValues & strArr(4) & vbTab
            strValues = strValues & strArr(12) & vbTab
            blnFirst = False
        End If
        
        '��ȷ����ص��ִ�
        'ReckonInfo              ����������ܶεĽ���ֶ���Ϣ���������ŵ�Ԫ����������Ԫ��
        Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "ReckonInfo", "")
        
        'akc190  string  20      ������
         AppendXMLNode objXMLItem, "akc190", strArr(0)
        'yka103  string  20      ������
         AppendXMLNode objXMLItem, "yka103", strArr(1)
        'yka198  string  20      �˵���Ӧ������
         AppendXMLNode objXMLItem, "yka198", strArr(2)
        'ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
         AppendXMLNode objXMLItem, "ykc114", strArr(3)
        'yab003  string  4       �籣�����������
         AppendXMLNode objXMLItem, "yab003", strArr(4)
        'aka213  string  2       �ֶα�׼��03 ���ߣ� 05 ����ҽ�� ��06 ���ҽ�ƣ�07 ����
         AppendXMLNode objXMLItem, "aka213", strArr(22)
        'yka056  number  14  2   ȫ�Էѽ��
         AppendXMLNode objXMLItem, "yka056", strArr(23)
        'yka057  number  14  2   �ҹ��Էѽ��
         AppendXMLNode objXMLItem, "yka057", strArr(24)
        'yka111  number  14  2   ���Ϸ�Χ���
         AppendXMLNode objXMLItem, "yka111", strArr(25)
        'yka106  number  14  2   �Ը����
         AppendXMLNode objXMLItem, "yka106", strArr(26)
        'yka107  number  14  2   ֧�����
         AppendXMLNode objXMLItem, "yka107", strArr(27)
        'yka063  number  14  2   ����Աͳ��֧�����
         AppendXMLNode objXMLItem, "yka063", strArr(28)
        'yka221  number  14  2   ����ҽ�Ʋ��������Ը��ۼƽ��
         AppendXMLNode objXMLItem, "yka221", strArr(29)
        'Akc315  String  3       ҽ������ְ��
         AppendXMLNode objXMLItem, "Akc315", strArr(12)
         
        
        '���ݷֶα�׼,�������ֵ
        '0-ȫ�Է�:decode(�ֶα�׼,'02',yka056+yka057,0)
        '1-���Ը�����:decode(�ֶα�׼,'05',yka057,'06',yka057,0)
        '2-���Ͻ��:decode(�ֶα�׼,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
        '����
        '3-�����ܶ�:decode(�ֶα�׼,'02','yka056,'02','yka057+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka063+yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
        'סԺ
        '3-�����ܶ�:decode(�ֶα�׼,'02','yka056,'02','yka057+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
        
        '4-��������:decode(�ֶα�׼,'03',yka106+yka057,0)
        '5-����ҽ���Ը�:decode(�ֶα�׼,'05',yka106,0)
        
        '6-����ҽ��ͳ��֧��:decode(�ֶα�׼,'05',yka107,0)
        '7-����Ը�:decode(�ֶα�׼,'06','yka106,0)
        '8-���֧��:decode(�ֶα�׼,'06','yka107
        '����
        '   9-����Ա����:decode(�ֶα�׼,'05',yka063,'03',yka063,'07',yka063,0)
        
        'סԺ
        '   9-����Ա����:decode(�ֶα�׼,'07',yka063,0)
        '10-�����Ը�:decode(�ֶα�׼,'08',yka106+yka057,0)
        '���
        
        dblTmp(0) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)), 0)
        dblTmp(1) = Decode(strArr(22), "05", Val(strArr(24)), "06", Val(strArr(24)), 0)
        dblTmp(2) = Decode(strArr(22), "03", Val(strArr(25)), "03", Val(strArr(25)), "04", Val(strArr(25)), "05", Val(strArr(25)), "06", Val(strArr(25)), "07", Val(strArr(25)), "10", Val(strArr(25)), 0)
        
        If g�������_�����山.�����־ = 1 Then
            dblTmp(3) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)) + Val(strArr(28)), "03", Val(strArr(26)) + Val(strArr(24)), "07", Val(strArr(28)), "04", Val(strArr(25)), "05", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "06", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "08", Val(strArr(25)) + Val(strArr(24)), "10", Val(strArr(25)), 0)
        Else
            'dblTmp(3) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)) + Val(strArr(28)), "03", Val(strArr(26)) + Val(strArr(24)), "07", Val(strArr(28)), "04", Val(strArr(25)), "05", Val(strArr(28)) + Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "06", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "08", Val(strArr(25)) + Val(strArr(24)), "10", Val(strArr(25)), 0)
            '��ʱ���롰yka063(03��)
            dblTmp(3) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)) + Val(strArr(28)), "03", Val(strArr(26)) + Val(strArr(24)) + Val(strArr(28)), "07", Val(strArr(28)), "04", Val(strArr(25)), "05", Val(strArr(28)) + Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "06", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "08", Val(strArr(25)) + Val(strArr(24)), "10", Val(strArr(25)), 0)
        End If
        
        dblTmp(4) = Decode(strArr(22), "03", Val(strArr(24)) + Val(strArr(26)), 0)
        dblTmp(5) = Decode(strArr(22), "05", Val(strArr(26)), 0)
        
        dblTmp(6) = Decode(strArr(22), "05", Val(strArr(27)), 0)
        
        dblTmp(7) = Decode(strArr(22), "06", Val(strArr(26)), 0)
        
        dblTmp(8) = Decode(strArr(22), "06", Val(strArr(27)), 0)
        
        If g�������_�����山.�����־ = 1 Then
            dblTmp(9) = Decode(strArr(22), "07", Val(strArr(28)), 0)
        Else
            'dblTmp(9) = Decode(strArr(22), "05", Val(strArr(28)), "07", Val(strArr(28)), 0)
            '������03�ε�yka063
            dblTmp(9) = Decode(strArr(22), "05", Val(strArr(28)), "03", Val(strArr(28)), "07", Val(strArr(28)), 0)
        End If
        dblTmp(10) = Decode(strArr(22), "08", Val(strArr(26)) + Val(strArr(24)), 0)
        
        '�ֱ����
        If strArr(1) = strArr(2) Then
            For i = 0 To 10
                dblSumSubMony(i) = dblSumSubMony(i) + dblTmp(i)
            Next
        Else
            For i = 0 To 10
                dblSumSubmony1(i) = dblSumSubmony1(i) + dblTmp(i)
            Next
            
            strvalues1 = strvalues1 & strArr(0) & vbTab
            strvalues1 = strvalues1 & strArr(1) & vbTab
            strvalues1 = strvalues1 & strArr(2) & vbTab
            strvalues1 = strvalues1 & Val(str������¼��) & vbTab
            strvalues1 = strvalues1 & strArr(4) & vbTab
            strvalues1 = strvalues1 & strArr(12) & vbTab
        End If
        
        '���ܺ�
        For i = 0 To 10
            dblSumMony(i) = dblSumMony(i) + dblTmp(i)
        Next
    Loop
    
    objText.Close
    
    '�������������,�����轫ֵд��
    If Get���㷽ʽ(dblSumMony) = False Then
        Exit Function
    End If
    
'    If Format(g�������_�����山.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(dblSumMony(3), "###0.00;-###0.00;0;0") Then
'        Dim blnYes As Boolean
'        '�����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
'        ShowMsgbox "���ν����ܶ�(" & g�������_�����山.�����ܶ� & ") ��" & vbCrLf & _
'                    "   ���ķ��ص��ܶ�(" & dblSumMony(3) & ")���²��ܽ���?"
'        Exit Function
'    End If
    
    'д����ý�����
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    strXMLtext1 = strXMLText
        
        
    '0-ȫ�Է�:decode(�ֶα�׼,'02',yka056+yka057,0)
    '1-���Ը�����:decode(�ֶα�׼,'05',yka057,'06',yka057,0)
    '2-���Ͻ��:decode(�ֶα�׼,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
    '3-�����ܶ�:decode(�ֶα�׼,'02','yka056,'02','yka053+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
    '4-��������:decode(�ֶα�׼,'03',yka106+yka057,0)
    '5-����ҽ���Ը�:decode(�ֶα�׼,'05',yka106,0)
    '6-����ҽ��ͳ��֧��:decode(�ֶα�׼,'05',yka107,0)
    '7-����Ը�:decode(�ֶα�׼,'06','yka106,0)
    '8-���֧��:decode(�ֶα�׼,'06','yka107
    '����
        '   9-����Ա����:decode(�ֶα�׼,'05',yka063,'03',yka063,'07',yka063,0)
    
    'סԺ
    '   9-����Ա����:decode(�ֶα�׼,'07',yka063,0)
    '10-�����Ը�:decode(�ֶα�׼,'08',yka106+yka057,0)
    Dim dbl�����ʻ� As Double
    Dim dbl�ֽ�   As Double
    
    
    
    i = 0
ECal2:
    
    'д����ý��������Ϣ
    strArr = Split(strValues, vbTab)
    
    If g�������_�����山.�����־ = 1 Then
        dbl�ֽ� = IIf(i = 99, dblSumSubmony1(0), dblSumSubMony(0))
        If ҽ�������Ѿ���Ժ(lng����id) Then
            dbl�����ʻ� = IIf(i = 99, dblSumSubmony1(1) + dblSumSubmony1(5) + dblSumSubmony1(7) + dblSumSubmony1(10) + dblSumSubmony1(4), dblSumSubMony(1) + dblSumSubMony(5) + dblSumSubMony(7) + dblSumSubMony(10) + dblSumSubMony(4))
        Else
            dbl�ֽ� = dbl�ֽ� + IIf(i = 99, dblSumSubmony1(1) + dblSumSubmony1(5) + dblSumSubmony1(7) + dblSumSubmony1(10) + dblSumSubmony1(4), dblSumSubMony(1) + dblSumSubMony(5) + dblSumSubMony(7) + dblSumSubMony(10) + dblSumSubMony(4))
            dbl�����ʻ� = 0
        End If
        
        If g�������_�����山.�ʻ���� <= dbl�����ʻ� Then
            dbl�ֽ� = dbl�ֽ� + dbl�����ʻ� - g�������_�����山.�ʻ����
            dbl�����ʻ� = g�������_�����山.�ʻ����
        End If
    Else
        dbl�����ʻ� = dblSumMony(1) + dblSumMony(5) + dblSumMony(7) + dblSumMony(10) + dblSumMony(4)
        
        dbl�ֽ� = dblSumMony(0)
        
        If g�������_�����山.�ʻ���� <= dbl�����ʻ� Then
            dbl�ֽ� = dbl�ֽ� + dbl�����ʻ� - g�������_�����山.�ʻ����
            dbl�����ʻ� = g�������_�����山.�ʻ����
        End If
    End If
    
    '���̲���:
    '    ����id, ��ҳid, ������, ������, �˵������, ������¼���, ���˱��, ��λ���, ����, �Ա�, ��������, ʵ������,
    '    �ۼƽɷ�����, ҽ����Ա���, ҽ�ƻ�������, ��֧��������, ҽ�ƻ������, ���ֲ���־, ֧�����, ���ֱ���, ��������,
    '    ҽ�Ʒ��ܶ�, ȫ�Է��ܶ�, �ҹ��Է��ܶ�, ���Ϸ�Χ�ܶ�, �����ʻ�֧���ܶ�, �����ֽ�֧���ܶ�, ����ʱ��, �����������,
    '    ҽ���չ���� , ҽ�Ʋ������, ������㷽ʽ, ��Ʊ��, ��ע, �ֶμ������, ҽ���������
    
    strSql = "ZL_���û�����Ϣ_INSERT(" & lng����id & ","
    strSql = strSql & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ","
    strSql = strSql & "'" & strArr(0) & "',"
    strSql = strSql & "'" & strArr(1) & "',"
    strSql = strSql & "'" & strArr(2) & "',"
    strSql = strSql & "" & Val(strArr(3)) & ","
    strSql = strSql & "" & g�������_�����山.���˱�� & ","
    strSql = strSql & "" & g�������_�����山.��λ���� & ","
    strSql = strSql & "'" & g�������_�����山.���� & "',"
    strSql = strSql & "'" & g�������_�����山.�Ա� & "',"
    
    If g�������_�����山.�������� = "" Then
        strSql = strSql & "NULL,"
    Else
        strSql = strSql & "to_date('" & g�������_�����山.�������� & "','yyyy-mm-dd'),"
    End If
    
    strSql = strSql & "" & g�������_�����山.���� & ","
    strSql = strSql & "" & g�������_�����山.�ۼƽɷ����� & ","
    strSql = strSql & "'" & g�������_�����山.ҽ����Ա��� & "',"
    strSql = strSql & "'" & InitInfor_�����山.ҽԺ���� & "',"
    strSql = strSql & "'" & "01" & "',"
    strSql = strSql & "'" & "" & "',"       'ҽ�ƻ������
    strSql = strSql & "'" & IIf(g�������_�����山.����ID <> 0, "1", "0") & "',"
        
    strSql = strSql & "'" & g�������_�����山.֧����� & "',"
    strSql = strSql & "'" & g�������_�����山.���ֱ��� & "',"
    
    strSql = strSql & "" & 0 & ","      '��������
    If g�������_�����山.�����־ = 1 Then
        strSql = strSql & "" & IIf(i = 99, dblSumSubmony1(3), dblSumSubMony(3)) & ","
        strSql = strSql & "" & IIf(i = 99, dblSumSubmony1(0), dblSumSubMony(0)) & ","
        strSql = strSql & "" & IIf(i = 99, dblSumSubmony1(1), dblSumSubMony(1)) & ","
        strSql = strSql & "" & IIf(i = 99, dblSumSubmony1(2), dblSumSubMony(2)) & ","
    Else
        strSql = strSql & "" & dblSumMony(3) & ","
        strSql = strSql & "" & dblSumMony(0) & ","
        strSql = strSql & "" & dblSumMony(1) & ","
        strSql = strSql & "" & dblSumMony(2) & ","
    End If
    strSql = strSql & "" & dbl�����ʻ� & ","
    strSql = strSql & "" & dbl�ֽ� & ","
    strSql = strSql & "to_date('" & str����ʱ�� & "','yyyy-mm-dd HH24:mi:ss'),"
    strSql = strSql & "'" & strArr(4) & "',"
    strSql = strSql & "'" & g�������_�����山.ҽ���չ���� & "',"
    strSql = strSql & "'" & g�������_�����山.ҽ�Ʋ������ & "',"
    strSql = strSql & "'" & g�������_�����山.������㷽ʽ & "',"
    strSql = strSql & "'" & g�������_�����山.��Ʊ�� & "',"
    strSql = strSql & "'" & "" & "',"
    strSql = strSql & "'" & str������Ϣ & "',"
    strSql = strSql & "'" & strArr(5) & "')"
            
    
    If g�������_�����山.������� Then
        '������㲻��������
    Else
        '��������
        gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
    End If
    
    Call intXML
    
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g�������_�����山.�籣���칹������, 1, 4)
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ10, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "10"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    
    'akc190  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", strArr(0)
    'yka103  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "yka103", strArr(1)
    'yka198  string  20      �˵���Ӧ������
    AppendXMLNode gobjXMLInPut.documentElement, "yka198", strArr(2)
    
    'ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc114", strArr(3)
    'aac001  number  15  0   ���˱��
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g�������_�����山.���˱��
    'aab001  number  15  0   ��λ���
    AppendXMLNode gobjXMLInPut.documentElement, "aab001", g�������_�����山.��λ����
    'aac003  string  20      ����
    AppendXMLNode gobjXMLInPut.documentElement, "aac003", g�������_�����山.����
    'aac004  string  1       �Ա𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "aac004", g�������_�����山.�Ա�
    
    'aac006  date    ��      ��������
    AppendXMLNode gobjXMLInPut.documentElement, "aac006", g�������_�����山.��������
    'akc023  number  3       ʵ������
    AppendXMLNode gobjXMLInPut.documentElement, "akc023", g�������_�����山.����
    'ykc021  number  3       �ۼƽɷ�����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc021", g�������_�����山.�ۼƽɷ�����
    'akc021  string  6       ҽ����Ա��𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g�������_�����山.ҽ����Ա���
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_�����山.ҽԺ����
    'ykb006  string  3       ����ҽ�ƻ�����֧�������
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"            '��֧��������
    'akb023  string  6       ҽ�ƻ�����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "akb023", InitInfor_�����山.�������
    
    'aka123  string  1       ���ֲ���־���������
    AppendXMLNode gobjXMLInPut.documentElement, "aka123", IIf(g�������_�����山.����ID <> 0, "1", "0")         '���ֲ���־
    'aka130  string  6       ֧����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g�������_�����山.֧�����
    'yka026  string  20      ���ֱ���
    AppendXMLNode gobjXMLInPut.documentElement, "yka026", g�������_�����山.���ֱ���
    'yka115  number  14  2   ��������
    AppendXMLNode gobjXMLInPut.documentElement, "yka115", g�������_�����山.��������            '��������
    
    If g�������_�����山.�����־ = 1 Then
        'yka055  number  14  2   ҽ�Ʒ��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", IIf(i = 99, dblSumSubmony1(3), dblSumSubMony(3))            '
        'yka056  number  14  2   ȫ�Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", IIf(i = 99, dblSumSubmony1(0), dblSumSubMony(0))             '
        'yka057  number  14  2   �ҹ��Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", IIf(i = 99, dblSumSubmony1(1), dblSumSubMony(1))              '
        'yka111  number  14  2   ���Ϸ�Χ�ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", IIf(i = 99, dblSumSubmony1(2), dblSumSubMony(2))                 '
    Else
        'yka055  number  14  2   ҽ�Ʒ��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", dblSumMony(3)                '
        'yka056  number  14  2   ȫ�Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", dblSumMony(0)              '
        'yka057  number  14  2   �ҹ��Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", dblSumMony(1)               '
        'yka111  number  14  2   ���Ϸ�Χ�ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", dblSumMony(2)                '
    End If
    'yka112  number  14  2   �����˻�֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "yka112", dbl�����ʻ�                 '
    'yka113  number  14  2   �����ֽ�֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "yka113", dbl�ֽ�                  '
    'aae036  date        ��  ����ʱ��
    '����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "aae036", str����ʱ��                  '
    'yab003  string  4       �籣�����������
    AppendXMLNode gobjXMLInPut.documentElement, "yab003", strArr(4)                  '
    'ykc120  string  6       ҽ���չ���𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc120", g�������_�����山.ҽ���չ����                   '
    'ykc121  string  6       ����ҽ�Ʋ�����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc121", g�������_�����山.ҽ�Ʋ������                    '
    'yka222  string  6       ������㷽ʽ
    AppendXMLNode gobjXMLInPut.documentElement, "yka222", g�������_�����山.������㷽ʽ                    '
    'yka110  string  20      ��Ʊ��
    AppendXMLNode gobjXMLInPut.documentElement, "yka110", g�������_�����山.��Ʊ��                                '
    'aae013  string  100     ��ע
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""                              '
    
    'gkc010  string  800     �ֶμ������(סԺ��)
    AppendXMLNode gobjXMLInPut.documentElement, "gkc010", "||GKC010_LXH||"                              '
    'akc315  string  3       ҽ�ƴ���������𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "akc315", strArr(5)                              '
        
    'д�������Ϣ
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    strXMLText = Replace(strXMLText, "||GKC010_LXH||", str������Ϣ)
    If g�������_�����山.������� Then
    Else
        If ҵ������_�����山(���������Ϣд��, strXMLText, strOutPut) = False Then
            
            If g�������_�����山.�����־ = 1 And i = 99 Then
                '��Ϊ���ڱ������Ļ�����Ϣ,����������Ѿ��ϴ�����ϸ��¼
                
            Else
                '����ǻ�����Ϣд��ʧ��,��ֱ���˳�����
            End If
            Exit Function
        End If
        
        If g�������_�����山.�����־ = 1 And i <> 99 And Trim(strvalues1) <> "" Then
            '����еڶ���д�������Ϣ
            i = 99
            strValues = strvalues1
            GoTo ECal2:
        End If
        strOutPut = ""
        'д���
        If ҵ������_�����山(������д��, strXMLtext1, strOutPut) = False Then
            '�϶�������ϸ������Ϣ,�����贫���෴��,����������Ϣ�ͽ��
            Call ���û�����Ϣ����(g�������_�����山.������)
            Exit Function
        End If
    
    End If
    
    
    '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
  '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(����Ա����),�ʻ��ۼ�֧��_IN(���֧��),�ۼƽ���ͳ��_IN(����ҽ���Ը�),�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����(��������),�ⶥ��_IN(֧�����+10000),ʵ������_IN,
    '   �������ý��_IN(��������),ȫ�Ը����_IN(ȫ�Ը�),�����Ը����_IN(�����Ը�),
    '   ����ͳ����_IN(���Ͻ��),ͳ�ﱨ�����_IN(����ҽ��ͳ��֧��),���Ը����_IN(����Ը�),�����Ը����_IN(�����Ը�),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(������),��ҳID_IN,��;����_IN,��ע_IN(������)
     
     '0-ȫ�Է�:decode(�ֶα�׼,'02',yka056+yka057,0)
    '1-���Ը�����:decode(�ֶα�׼,'05',yka057,'06',yka057,0)
    '2-���Ͻ��:decode(�ֶα�׼,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
    '3-�����ܶ�:decode(�ֶα�׼,'02','yka056,'02','yka053+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
    '4-��������:decode(�ֶα�׼,'03',yka106+yka057,0)
    '5-����ҽ���Ը�:decode(�ֶα�׼,'05',yka106,0)
    '6-����ҽ��ͳ��֧��:decode(�ֶα�׼,'05',yka107,0)
    '7-����Ը�:decode(�ֶα�׼,'06','yka106,0)
    '8-���֧��:decode(�ֶα�׼,'06','yka107
    '9-����Ա����:decode(�ֶα�׼,'07',yka063,0)
    '10-�����Ը�:decode(�ֶα�׼,'08',yka106+yka057,0)
    If g�������_�����山.������� Then
        '������㲻��������
    Else
        Err = 0
        On Error Resume Next
        
        gstrSQL = "zl_���ս����¼_insert(" & IIf(g�������_�����山.�����־ = 1, 2, 1) & "," & g�������_�����山.����ID & "," & gintInsure & "," & lng����id & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          dblSumMony(9) & "," & dblSumMony(8) & "," & dblSumMony(5) & ",NULL,NULL," & dblSumMony(4) & "," & "1" & g�������_�����山.֧����� & ",NULL," & _
           dblSumMony(3) & "," & dblSumMony(0) & "," & dblSumMony(1) & "," & _
            "" & dblSumMony(2) & "," & dblSumMony(6) & "," & dblSumMony(7) & "," & dblSumMony(10) & "," & dbl�����ʻ� & ",'" & _
           g�������_�����山.������ & "'," & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & "," & IIf(g�������_�����山.��;���� = 1, "1", "NULL") & ",'" & _
           g�������_�����山.������ & "')"
        Call ExecuteProcedure("���汣�ս����¼")
        If Err <> 0 Then
            DebugTool "���±��ս����¼ʱ����!" & vbCrLf & " �����:" & Err.Number & " ��������:" & Err.Description
            Err.Clear
            '�϶�������ϸ������Ϣ�ͷ��ý����,�����贫���෴��,����������Ϣ�ͽ��
            Call ���û�����Ϣ����(g�������_�����山.������)
            Call ���ý���������(g�������_�����山.������)
            Exit Function
        End If
        '������ʻ�
        If g�������_�����山.�����־ = 1 Then
            If ҽ�������Ѿ���Ժ(g�������_�����山.lng����id) Then
                '�ۼ������ʻ�
                If IC���ʻ�֧��_�����山(dbl�����ʻ�, str����ʱ��, g�������_�����山.������) = False Then
                
                    '�϶�������ϸ������Ϣ�ͷ��ý����,�����贫���෴��,����������Ϣ�ͽ��
                    Call ���û�����Ϣ����(g�������_�����山.������)
                    Call ���ý���������(g�������_�����山.������)
                    Exit Function
                End If
            End If
        Else
            '�ۼ������ʻ�
            If IC���ʻ�֧��_�����山(dbl�����ʻ�, str����ʱ��, g�������_�����山.������) = False Then
                '�϶�������ϸ������Ϣ�ͷ��ý����,�����贫���෴��,����������Ϣ�ͽ��
                Call ���û�����Ϣ����(g�������_�����山.������)
                Call ���ý���������(g�������_�����山.������)
                Exit Function
            End If
        End If
    
    End If
    ���ý���ֽ� = True
    
    
    Exit Function
ErrHand:
   DebugTool "���ý���������(���ý���ֽ�)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
   objText.Close
End Function
Private Function Get���㷽ʽ(ByVal strDblCur As Variant) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ݴ���ֵȷ����Ӧ�Ľ��㷽ʽ
    '--�����:
    '--������:str���㷽ʽ
    '--��  ��:�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    Dim dbl�����ʻ� As Double
    Dim dbl�ֽ� As Double
    
    'strDblCur���ݷֶα�׼,�������ֵ
        '0-ȫ�Է�:decode(�ֶα�׼,'02',yka056+yka057,0)
        '1-���Ը�����:decode(�ֶα�׼,'05',yka057,'06',yka057,0)
        '2-���Ͻ��:decode(�ֶα�׼,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
        '3-�����ܶ�:decode(�ֶα�׼,'02','yka056,'02','yka057+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
        
        '4-��������:decode(�ֶα�׼,'03',yka106+yka057,0)
        '5-����ҽ���Ը�:decode(�ֶα�׼,'05',yka106,0)
        
        '6-����ҽ��ͳ��֧��:decode(�ֶα�׼,'05',yka107,0)
        '7-����Ը�:decode(�ֶα�׼,'06','yka106,0)
        '8-���֧��:decode(�ֶα�׼,'06','yka107
        '9-����Ա����:decode(�ֶα�׼,'07',yka063,0)
        '10-�����Ը�:decode(�ֶα�׼,'08',yka106+yka057,0)
        
    'ҽ������=����ҽ��ͳ��֧��+
    '�����ʻ�֧��=���Ը�����+����ҽ���Ը�+�����Ը�+���Ը�����
    '������=����Ա����
    
    Err = 0
    On Error GoTo ErrHand:
    DebugTool "����(" & "Get���㷽ʽ" & ")"
    
    dbl�ֽ� = strDblCur(0)
    
    If g�������_�����山.�����־ = 1 Then
        If ҽ�������Ѿ���Ժ(g�������_�����山.lng����id) Then
            '��Ժ������������ʻ�
            dbl�����ʻ� = strDblCur(1) + strDblCur(5) + strDblCur(7) + strDblCur(10) + strDblCur(4)
        Else
            '��;�����޸����ʻ�֧��
            dbl�����ʻ� = 0
        End If
    Else
            dbl�����ʻ� = strDblCur(1) + strDblCur(5) + strDblCur(7) + strDblCur(10) + strDblCur(4)
    End If
    
    
    If g�������_�����山.�ʻ���� <= dbl�����ʻ� Then
        dbl�ֽ� = dbl�ֽ� + dbl�����ʻ� - g�������_�����山.�ʻ����
        dbl�����ʻ� = g�������_�����山.�ʻ����
    End If
    str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & strDblCur(6)
    str���㷽ʽ = str���㷽ʽ & "||����|" & strDblCur(8)
    str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & strDblCur(9)
    str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & dbl�����ʻ�
    
    If Format(g�������_�����山.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(strDblCur(3), "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '�����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
        ShowMsgbox "���ν����ܶ�(" & g�������_�����山.�����ܶ� & ") ��" & vbCrLf & _
                    "   ���ķ��ص��ܶ�(" & strDblCur(3) & ")���²��ܽ���?"
        Exit Function
    End If
    
   
    '�������
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        g�������_�����山.������Ϣ = str���㷽ʽ
        If g�������_�����山.������� Then
            '������㲻�����������
        Else
            If g�������_�����山.�����־ = 0 Or g�������_�����山.�����־ = 2 Then
                gstrSQL = "zl_���˽����¼_Update(" & g�������_�����山.����ID & ",'" & str���㷽ʽ & "',0)"
                Call ExecuteProcedure("����Ԥ����¼")
            Else
                If g�������_�����山.������� Then
                    gstrSQL = "zl_���˽����¼_Update(" & g�������_�����山.����ID & ",'" & str���㷽ʽ & "',1)"
                    Call ExecuteProcedure("����Ԥ����¼")
                End If
            End If
        End If
        g�������_�����山.������Ϣ = Replace(g�������_�����山.������Ϣ, "||", "[")
        g�������_�����山.������Ϣ = Replace(g�������_�����山.������Ϣ, "|", ";")
        g�������_�����山.������Ϣ = Replace(g�������_�����山.������Ϣ, "[", ";0|")
        g�������_�����山.������Ϣ = g�������_�����山.������Ϣ & ";0"
    End If
    
    '��ʾ������Ϣ
    If g�������_�����山.������� Or g�������_�����山.�����־ = 1 Then
    Else
        If frm������Ϣ.ShowME(g�������_�����山.����ID, True) = False Then
            Get���㷽ʽ = False
            Exit Function
        End If
    End If
    Get���㷽ʽ = True
    Exit Function
ErrHand:
  DebugTool "���没�����¼����(Get���㷽ʽ)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
End Function
'20041012:���˺�:��Ϊ�����������ı��ļ��ĸ�ʽ����(�����ý��������ȷ��.����ɶ�ȡ��¼Ϊ���еģ���������.����ֻ�в�ȡ���س����������������ļ�
'Private Function Save������ϸ����ָ�(ByVal strFile As String) As Boolean
'    '-----------------------------------------------------------------------------------------------------------
'    '--��  ��:������ý�����������ϸ
'    '--�����:
'    '--������:
'    '--��  ��:
'    '-----------------------------------------------------------------------------------------------------------
'    DebugTool "����:Save������ϸ����ָ�"
'    Dim strSql As String
'    Dim objFile As New FileSystemObject
'    Dim objText As TextStream
'    Dim strText As String
'    Dim strTemp  As String
'    Dim strArr
'
'    Dim strXMLText As String
'
'    If g�������_�����山.�����־ <> 1 Then
'        '���ﲿ��,����û������ı�,�����ޱ��������ϸ��Ϣ
'        Save������ϸ����ָ� = True
'        Exit Function
'    End If
'
'    Err = 0
'    On Error GoTo ErrHand:
'
'    Set objText = objFile.OpenTextFile(strFile)
'    '��ϸ���̲���(��������):
'    '   ������ˮ��,������, ��Ŀ����, ��Ʒ������, ������Ʒ����, ������, �����������, ��Ŀ���㷽ʽ, �����ܶ�, �ʻ�֧����, �ֶα�׼,
'    '   ȫ�Էѽ��, �ҹ��Էѽ��, ���Ϸ�Χ���, �Ը����, ֧�����, ����Աͳ��֧��, �����Ը��ۼ�, �Ը�����
'
'
'    Do While Not objText.AtEndOfStream
'        strTemp = Trim(objText.ReadLine)
'        strArr = Split(strTemp, vbTab)
'            '�ı���ʽ
'                '            AKC190  string  20      ������
'                '            YKA104  number  15  0   �˵���Ӧ������ˮ��
'                '            YKA002  string  20      ҽ����Ŀ����
'                '            YKA231  string  20      ҽ����Ŀ��Ʒ������
'                '            YKA247  string  20      ��������ҽ����Ŀ��Ʒ������
'                '            YKA096  number  20      �Ը�����
'                '            YKA272  string  4       Ŀ¼����
'                '            AKC225  string  6       ʵ�ʼ۸�
'                '            AKC226  number  14  2   ����
'                '            YKA055  number  14  2   �����ܶ�
'                '            YKA056  number  14  2   �Ը����
'                '            YKA057  number  14  2   �ҹ��Ը����
'                '            YKA111  number  14  2   ���Ϸ�Χ���ֽ��
'                '            YKA103  number  14  2   �˵���Ӧ������
'            '���̲���:
'            '        ������_IN ҽ�����÷�����Ϣ.������%type,
'            '        �˵���ˮ��_IN   ҽ�����÷�����Ϣ.�˵���ˮ��%type,
'            '        ��Ŀ����_IN ҽ�����÷�����Ϣ.��Ŀ����%type,
'            '        ��Ʒ������_IN   ҽ�����÷�����Ϣ.��Ʒ������%type,
'            '        ������Ʒ������_IN   ҽ�����÷�����Ϣ.������Ʒ������%type,
'            '        �Ը�����_IN ҽ�����÷�����Ϣ.�Ը�����%type,
'            '        Ŀ¼����_IN ҽ�����÷�����Ϣ.Ŀ¼����%type,
'            '        ʵ�ʼ۸�_IN ҽ�����÷�����Ϣ.Ŀ¼����%type,
'            '        ����_IN     ҽ�����÷�����Ϣ.����%type,
'            '        �����ܶ�_IN ҽ�����÷�����Ϣ.�����ܶ�%type,
'            '        �Ը����_IN ҽ�����÷�����Ϣ.�Ը����%type,
'            '        �ҹ��Ը����_IN ҽ�����÷�����Ϣ.�ҹ��Ը����%type,
'            '        ���Ϸ�Χ���_IN ҽ�����÷�����Ϣ.���Ϸ�Χ���%type,
'            '        �˵�������_IN ҽ�����÷�����Ϣ.�˵�������%type
'
'            strSql = "ZL_ҽ�����÷�����Ϣ_INSERT("
'            strSql = strSql & "'" & strArr(0) & "',"
'            strSql = strSql & "" & Val(strArr(1)) & ","
'            strSql = strSql & "'" & strArr(2) & "',"
'            strSql = strSql & "'" & strArr(3) & "',"
'            strSql = strSql & "'" & strArr(4) & "',"
'            strSql = strSql & "" & Val(strArr(5)) & ","
'            strSql = strSql & "'" & strArr(6) & "',"
'            strSql = strSql & "" & Val(strArr(7)) & ","
'            strSql = strSql & "" & Val(strArr(8)) & ","
'            strSql = strSql & "" & Val(strArr(9)) & ","
'            strSql = strSql & "" & Val(strArr(10)) & ","
'            strSql = strSql & "" & Val(strArr(11)) & ","
'            strSql = strSql & "" & Val(strArr(12)) & ","
'            strSql = strSql & "" & Val(strArr(13)) & ")"
'
'
'            'ֻ��סԺ����
'            '20040720ȡ��
'
'
'            '       StrSQL = "ZL_ҽ����ϸ����_UPDATE("
'
'            '            '������ˮ��
'            '            StrSQL = StrSQL & Val(strArr(1)) & ","
'            '            StrSQL = StrSQL & "'" & strArr(0) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(2) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(3) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(4) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(5) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(6) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(7) & "',"
'            '
'            '            StrSQL = StrSQL & "" & Val(strArr(8)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(9)) & ","
'            '
'            '            '�ֶα�׼
'            '            StrSQL = StrSQL & "" & Val(strArr(18)) & ","
'            '
'            '            StrSQL = StrSQL & "" & Val(strArr(19)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(20)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(21)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(22)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(23)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(24)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(25)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(26)) & ")"
'            '�޸ķ��÷�����Ϣ
'            gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
'    Loop
'    objText.Close
'    Save������ϸ����ָ� = True
'    Exit Function
'ErrHand:
'
'    DebugTool "��ϸ�ָ��(Save������ϸ����ָ�)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
'   objText.Close
'End Function
Private Function Save������ϸ����ָ�(ByVal strFile As String, ByVal lng����id As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ý�����������ϸ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    '20041012:���˺�:��Ϊ�����������ı��ļ��ĸ�ʽ����(�����ý��������ȷ��.����ɶ�ȡ��¼Ϊ���еģ���������.����ֻ�в�ȡ���س����������������ļ�
    DebugTool "����:Save������ϸ����ָ�"
    Dim strSql As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strText As String
    Dim strTemp  As String
    Dim strArr
    Dim strArr1
    Dim i As Long
    
    Dim strXMLText As String
    
    If g�������_�����山.�����־ <> 1 Then
        '���ﲿ��,����û������ı�,�����ޱ��������ϸ��Ϣ
        Save������ϸ����ָ� = True
        Exit Function
    End If
    If g�������_�����山.������� Then
            '��������ò��ű����������
            Save������ϸ����ָ� = True
            Exit Function
    End If
    Err = 0
    On Error GoTo ErrHand:
    
    Set objText = objFile.OpenTextFile(strFile)
    '��ϸ���̲���(��������):
    '   ������ˮ��,������,����ID, ��Ŀ����, ��Ʒ������, ������Ʒ����, ������, �����������, ��Ŀ���㷽ʽ, �����ܶ�, �ʻ�֧����, �ֶα�׼,
    '   ȫ�Էѽ��, �ҹ��Էѽ��, ���Ϸ�Χ���, �Ը����, ֧�����, ����Աͳ��֧��, �����Ը��ۼ�, �Ը�����
    strText = Trim(objText.ReadAll)
    strArr1 = Split(strText, vbCr)
    
    For i = 0 To UBound(strArr1)
        If Trim(strArr1(i)) <> "" Then
            strArr = Split(strArr1(i), vbTab)
                '�ı���ʽ
                    '            AKC190  string  20      ������
                    '            YKA104  number  15  0   �˵���Ӧ������ˮ��
                    '            YKA002  string  20      ҽ����Ŀ����
                    '            YKA231  string  20      ҽ����Ŀ��Ʒ������
                    '            YKA247  string  20      ��������ҽ����Ŀ��Ʒ������
                    '            YKA096  number  20      �Ը�����
                    '            YKA272  string  4       Ŀ¼����
                    '            AKC225  string  6       ʵ�ʼ۸�
                    '            AKC226  number  14  2   ����
                    '            YKA055  number  14  2   �����ܶ�
                    '            YKA056  number  14  2   �Ը����
                    '            YKA057  number  14  2   �ҹ��Ը����
                    '            YKA111  number  14  2   ���Ϸ�Χ���ֽ��
                    '            YKA103  number  14  2   �˵���Ӧ������
                '���̲���:
                '        ����ID,
                '        ��ҳID,
                '        ������_IN ҽ�����÷�����Ϣ.������%type,
                '        �˵���ˮ��_IN   ҽ�����÷�����Ϣ.�˵���ˮ��%type,
                '        ��Ŀ����_IN ҽ�����÷�����Ϣ.��Ŀ����%type,
                '        ��Ʒ������_IN   ҽ�����÷�����Ϣ.��Ʒ������%type,
                '        ������Ʒ������_IN   ҽ�����÷�����Ϣ.������Ʒ������%type,
                '        �Ը�����_IN ҽ�����÷�����Ϣ.�Ը�����%type,
                '        Ŀ¼����_IN ҽ�����÷�����Ϣ.Ŀ¼����%type,
                '        ʵ�ʼ۸�_IN ҽ�����÷�����Ϣ.Ŀ¼����%type,
                '        ����_IN     ҽ�����÷�����Ϣ.����%type,
                '        �����ܶ�_IN ҽ�����÷�����Ϣ.�����ܶ�%type,
                '        �Ը����_IN ҽ�����÷�����Ϣ.�Ը����%type,
                '        �ҹ��Ը����_IN ҽ�����÷�����Ϣ.�ҹ��Ը����%type,
                '        ���Ϸ�Χ���_IN ҽ�����÷�����Ϣ.���Ϸ�Χ���%type,
                '        �˵�������_IN ҽ�����÷�����Ϣ.�˵�������%type
                        
                strSql = "ZL_ҽ�����÷�����Ϣ_INSERT("
                strSql = strSql & "" & lng����id & ","
                strSql = strSql & "" & lng��ҳID & ","
                strSql = strSql & "'" & strArr(0) & "',"
                strSql = strSql & "" & Val(strArr(1)) & ","
                strSql = strSql & "'" & strArr(2) & "',"
                strSql = strSql & "'" & strArr(3) & "',"
                strSql = strSql & "'" & strArr(4) & "',"
                strSql = strSql & "" & Val(strArr(5)) & ","
                strSql = strSql & "'" & strArr(6) & "',"
                strSql = strSql & "" & Val(strArr(7)) & ","
                strSql = strSql & "" & Val(strArr(8)) & ","
                strSql = strSql & "" & Val(strArr(9)) & ","
                strSql = strSql & "" & Val(strArr(10)) & ","
                strSql = strSql & "" & Val(strArr(11)) & ","
                strSql = strSql & "" & Val(strArr(12)) & ","
                strSql = strSql & "" & Val(strArr(13)) & ")"
                
                
                'ֻ��סԺ����
                '20040720ȡ��
                
                
                '       StrSQL = "ZL_ҽ����ϸ����_UPDATE("
                             
                '            '������ˮ��
                '            StrSQL = StrSQL & Val(strArr(1)) & ","
                '            StrSQL = StrSQL & "'" & strArr(0) & "',"
                '            StrSQL = StrSQL & "'" & strArr(2) & "',"
                '            StrSQL = StrSQL & "'" & strArr(3) & "',"
                '            StrSQL = StrSQL & "'" & strArr(4) & "',"
                '            StrSQL = StrSQL & "'" & strArr(5) & "',"
                '            StrSQL = StrSQL & "'" & strArr(6) & "',"
                '            StrSQL = StrSQL & "'" & strArr(7) & "',"
                '
                '            StrSQL = StrSQL & "" & Val(strArr(8)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(9)) & ","
                '
                '            '�ֶα�׼
                '            StrSQL = StrSQL & "" & Val(strArr(18)) & ","
                '
                '            StrSQL = StrSQL & "" & Val(strArr(19)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(20)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(21)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(22)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(23)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(24)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(25)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(26)) & ")"
                '�޸ķ��÷�����Ϣ
                gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
            End If
    Next
    objText.Close
    Save������ϸ����ָ� = True
    Exit Function
ErrHand:

    DebugTool "��ϸ�ָ��(Save������ϸ����ָ�)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
   objText.Close
End Function

Private Function ���˷��ý���(ByVal lng����id As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ý���
    '--�����:
    '--������:
    '--��  ��:����ɹ�����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInFile1 As String          '���ν������������Ϣ�ļ���ŵ�ַ���ļ���
    Dim strInFile2 As String          '���ν�����ϸ��Ϣ��ŵ�ַ���ļ���
    Dim strInFile3 As String          '���ν�����Ϣ��ŵ�ַ���ļ���������Ϊ�գ�
    Dim strOutFile1 As String         '������ϸ�ָ���ı��ļ���ŵ�ַ���ļ��������Ʊ���ָ�
    Dim strOutXMLFile1 As String      '������ϸ�ָ���ı��ļ���ŵ�ַ���ļ�������XML��ʽ
    Dim stroutFile2 As String         '������������Ϣ���ı��ļ���ŵ�ַ���ļ������Ʊ���ָ�
    Dim stroutXMLFile2 As String      '������������Ϣ���ı��ļ���ŵ�ַ���ļ�������XML��ʽ
    Dim strErrMsg As String
    Dim lngAppCode As Long
    Dim gobjFile As New FileSystemObject
    
    Dim blnReturn As Boolean
    
    strInFile1 = gstrAppPath & "\����������Ϣ.txt"
    strInFile2 = gstrAppPath & "\������ϸ��Ϣ.txt"
    strInFile3 = gstrAppPath & "\���ν�����Ϣ.txt"
    
    strOutFile1 = gstrAppPath & "\������ϸ�ָ�.txt"
    strOutXMLFile1 = gstrAppPath & "\������ϸ�ָ�XML.txt"
    
    stroutFile2 = gstrAppPath & "\������������Ϣ.txt"
    stroutXMLFile2 = gstrAppPath & "\������������ϢXML.txt"
    ���˷��ý��� = False
    
    DebugTool "���˷��ý������"
    Err = 0
    On Error GoTo ErrHand:
    ���˷��ý��� = False
    If InitInfor_�����山.ģ������ Then
        Readģ������ ���ý���, "", ""
    Else
        Debug.Print Time
        blnReturn = gobj���ý���.chargereckoning(strInFile1, strInFile2, strInFile3, g�������_�����山.�籣���칹������, g�������_�����山.������, g�������_�����山.������, strOutFile1, strOutXMLFile1, stroutFile2, stroutXMLFile2, lngAppCode, strErrMsg)
        Debug.Print Time
        If blnReturn = False Then
            ShowMsgbox "�����:" & lngAppCode & vbCrLf & "������Ϣ:" & strErrMsg
            GoTo DelFile:
            Exit Function
        End If
    End If
    '��ϸ�ֽ�
    If Save������ϸ����ָ�(strOutFile1, lng����id, lng��ҳID) = False Then
        GoTo DelFile:
        Exit Function
    End If
    
    '���ý���ֽ�
    If ���ý���ֽ�(stroutFile2, lng����id, lng��ҳID) = False Then
        GoTo DelFile:
        Exit Function
    End If
    
    ���˷��ý��� = True
    GoTo DelFile:
    
    Exit Function
ErrHand:
    DebugTool "�������:" & Err.Number & "   ��Ϣ:" & Err.Description
    
DelFile:
    '�����ʱ�ļ�.
    Err = 0
    On Error Resume Next
    If gobjFile.FileExists(strOutFile1) = True Then
        gobjFile.DeleteFile strOutFile1, True
    End If
    If gobjFile.FileExists(stroutFile2) = True Then
        gobjFile.DeleteFile stroutFile2, True
    End If
    If gobjFile.FileExists(strOutXMLFile1) = True Then
        gobjFile.DeleteFile strOutXMLFile1, True
    End If
    If gobjFile.FileExists(stroutXMLFile2) = True Then
        gobjFile.DeleteFile stroutXMLFile2, True
    End If
End Function

Private Function Get������() As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������
    '--�����:
    '--������:
    '--��  ��:�µĽ������
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String, strErrMsg As String, intAppCode As Integer
    Dim blnReturn As Boolean
    
    If Not g�������_�����山.���� Then
        Get������ = ""
        Err = 0
        On Error GoTo ErrHand:
        If InitInfor_�����山.ģ������ Then
            gstrSQL = "Select ҽ������Ŀ¼_ID.nextval as ��� from dual"
            OpenRecordset_ZLYB rsTemp, "��ȡ������"
            Get������ = Nvl(rsTemp!���)
            Exit Function
        End If
        Call intXML
         AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g�������_�����山.�籣���칹������
        'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ102, ��ʶ��Сд���У�����λ��
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "102"
        
        Get������ = ""
        strInPut = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
        
        Err = 0
        On Error GoTo ErrHand:
        blnReturn = gobjYingHaiDll.dll_main_in(strInPut, strOutPut, intAppCode, strErrMsg)
        
        If blnReturn = False Then
          '��������,��ʾ����Ϣ
            ShowMsgbox strErrMsg
            Get������ = ""
            Exit Function
        End If
        Get������ = strOutPut
        Exit Function
    End If
    
    gstrSQL = "Select ֧��˳���,��ע ժҪ From ���ս����¼ where ��¼ID= " & g�������_�����山.����ID & " And ����=" & TYPE_�����山
    
    OpenRecordset rsTemp, "��ȡ������"
    
    If rsTemp.EOF Then
        Get������ = ""
    Else
        Get������ = Nvl(rsTemp!֧��˳���)
        g�������_�����山.������ = Substr(Nvl(rsTemp!ժҪ), 1, 20)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function Get������_�����山() As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������
    '--�����:
    '--������:
    '--��  ��:�µĽ������
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String
    Dim strOutPut As String, strErrMsg As String, intAppCode As Integer
    Dim blnReturn As Boolean
    
    
    If InitInfor_�����山.ģ������ Then
        gstrSQL = "Select ҽ������Ŀ¼_ID.nextval as ��� from dual"
        OpenRecordset_ZLYB rsTemp, "��ȡ������"
        Get������_�����山 = Nvl(rsTemp!���)
        Exit Function
    End If
    
     Call intXML
     AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g�������_�����山.�籣���칹������
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ08, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "101"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����

    Get������_�����山 = ""
    strInPut = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    Err = 0
    On Error GoTo ErrHand:
    blnReturn = gobjYingHaiDll.dll_main_in(strInPut, strOutPut, intAppCode, strErrMsg)
    
    If blnReturn = False Then
      '��������,��ʾ����Ϣ
        ShowMsgbox strErrMsg
        Get������_�����山 = ""
        Exit Function
    End If
    Get������_�����山 = strOutPut
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function �����������_�����山(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    
    'Ŀǰ��֧�������������
    
    str���㷽ʽ = ""
    �����������_�����山 = True
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
    
    If lng����ID = 0 And g�������_�����山.�����־ <> 1 Then
            '--��ȷ��������־
            strSql = " " & _
                "  Select Rownum ��ʶ��,A.ID,A.����ID,a.��ҳid,A.�շ�ϸĿid,������Ŀid,A.NO,A.��� ,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ�� as ����ʱ�� ,c.���� as ��������,a.������ as ����ҽ��,nvl(a.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
                "      A.����*A.���� as ����,A.���㵥λ,Round(A.ʵ�ս��/(A.����*A.����),4) as ʵ�ʼ۸�,A.ʵ�ս�� as ʵ�ս��, " & _
                "      A.�շ����,D.��Ŀ����,D.��Ŀ����,l.��Ա��� as �����������,l.������ �����,l.֧�����," & _
                "      l.������ ,'' as ������־,L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.����ID,L.����ʱ�� ,J.���� as ��Ʒ��" & _
                "  From (Select * From ���˷��ü�¼ Where ��¼״̬<>0 and NO='" & strNO & "' and ��¼����=" & lng��¼���� & " and ��¼״̬=" & lng��¼״̬ & " and  Nvl(���ӱ�־,0)<>9 ) A,���ű� C," & _
                "       ����֧����Ŀ D,�����ʻ� L,�շ�ϸĿ J " & _
                "  Where A.��������id=C.id(+) and  A.����id=L.����id  and a.�շ�ϸĿid=J.id and L.����=" & gintInsure & "  And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����= " & gintInsure & _
                "  Order by a.NO,A.��¼����,A.��¼״̬,a.���"
    Else
        If g�������_�����山.�����־ = 1 And lng����ID = 0 Then
            'סԺ�贫������ϸ��¼,����ݽ����ż��������ȷ��.

            strSql = "Select Rownum ��ʶ��, " & _
                     "          A.ID,A.����ID,a.��ҳid,A.�շ�ϸĿid ,������Ŀid,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ�� as ����ʱ��,c.���� as ��������,a.������ as ����ҽ��, " & _
                     "          nvl(a.�Ƿ��ϴ�, 0) �Ƿ��ϴ�,A.���� * A.���� as ����,A.���㵥λ,Round(A.ʵ�ս�� / (A.���� * A.����), 4) as ʵ�ʼ۸�, " & _
                     "          A.ʵ�ս�� as ʵ�ս��,A.�շ���� ,b.������ˮ��,b.�˵���ˮ��,D.��Ŀ����,D.��Ŀ����, " & _
                     "          L.��Ա��� as �����������,l.������  as �����,L.֧����� ,l.������,b.������־, " & _
                     "          L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.����ID,L.����ʱ�� ,J.���� as ��Ʒ��" & _
                     "   From  ���˷��ü�¼ a , " & _
                     "          ҽ����ϸ���� b,���ű� C,����֧����Ŀ D,�����ʻ� L,�շ�ϸĿ J  " & _
                     "   Where a.��¼״̬<>0 and nvl(a.���ӱ�־,0)<>9 and A.��������id = C.id(+) and a.id=b.����id and b.������='" & g�������_�����山.������ & "' " & IIf(g�������_�����山.lng����id = 0, "", " And A.����id =" & g�������_�����山.lng����id) & "   and A.����id = L.����id and L.���� = " & gintInsure & "  And " & _
                     "          A.�շ�ϸĿID = D.�շ�ϸĿID and a.�շ�ϸĿid=J.id And D.���� =  " & gintInsure & _
                     "    Order by a.NO,A.��¼����,A.��¼״̬,a.���"
        Else
            '--��ȷ��������־
            strSql = " " & _
                "  Select Rownum ��ʶ��,A.ID,A.����ID,a.��ҳid,A.�շ�ϸĿid,������Ŀid,A.NO,A.��� ,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ�� as ����ʱ�� ,c.���� as ��������,a.������ as ����ҽ��,nvl(a.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
                "      A.����*A.���� as ����,A.���㵥λ,Round(A.���ʽ��/(A.����*A.����),4) as ʵ�ʼ۸�,A.���ʽ�� as ʵ�ս��, " & _
                "      A.�շ����,D.��Ŀ����,D.��Ŀ����,��Ա��� as �����������,'" & g�������_�����山.������ & "' as �����,L.֧�����," & _
                "      '" & g�������_�����山.������ & "' as ������ ,'' as ������־,L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.����ID,L.����ʱ�� ,J.���� as ��Ʒ��" & _
                "  From (Select * From ���˷��ü�¼ Where ��¼״̬<>0 and ����ID=" & lng����ID & " and  Nvl(���ӱ�־,0)<>9 ) A,���ű� C," & _
                "       ����֧����Ŀ D,�����ʻ� L,�շ�ϸĿ J " & _
                "  Where A.��������id=C.id(+) and  A.����id=L.����id and a.�շ�ϸĿid=J.id and L.����=" & gintInsure & "  And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����= " & gintInsure & _
                "   Order by a.NO,A.��¼����,A.��¼״̬,a.���"
        End If
    End If

    Get��ϸ��¼ = strSql
End Function

Private Function ������ϸ�ϴ�(ByVal rs��ϸ As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ϸ�ϴ�
    '--�����:��ϸ��¼���ֶ�:'ID,����ID,�շ�ϸĿID,NO,������ˮ��,�˵���ˮ��,��¼����,��¼״̬,����ʱ��,��������,����ҽ��,,����,���㵥λ,ʵ�ʼ۸�,ʵ�ս��,�շ����,��Ŀ����,��Ŀ����,�����������,�����,������,������־,����,����,����,ҽ����,��Ա���,����ID,����ʱ��
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
   Dim rsTemp As New ADODB.Recordset
   Dim rs��Ŀ As New ADODB.Recordset
   Dim strXMLText As String
   Dim blnTrue As Boolean
   Dim strOutPut As String
   
    Err = 0
    On Error GoTo ErrHand:
    DebugTool "����(" & "������ϸ�ϴ�" & ")"
    If rs��ϸ Is Nothing Then Exit Function
    If rs��ϸ.RecordCount = 0 Then Exit Function
    
     ''ID,����ID,�շ�ϸĿID,NO,������ˮ��,�˵���ˮ��,��¼����,��¼״̬,����ʱ��,����,���㵥λ,ʵ�ʼ۸�,ʵ�ս��,�շ����,��Ŀ����,��Ŀ����,�����������,�����,������,������־,����,����,����,ҽ����,��Ա���,����ID,����ʱ��
    With rs��ϸ
        .Filter = 0
        .Filter = "�Ƿ��ϴ�=0"
        If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
        blnTrue = True
        Do While Not .EOF
                 If Nvl(!��Ŀ����) = "" Then
                     ShowMsgbox "����δ����ҽ����Ŀ,���ڱ�����Ŀ��������Ӧ�Ķ�Ӧ��ϵ!"
                     Exit Function
                 End If
                Call intXML
                Set rsTemp = Getҽ����ϸ����(!ID)
                If g�������_�����山.�����־ = 2 And Nvl(!������Ŀid, 0) = InitInfor_�����山.������Ŀid Then
                    Set rs��Ŀ = Get������Ŀ(Nvl(rsTemp!��Ŀ����))
                Else
                    Set rs��Ŀ = Get������Ŀ(Nvl(!��Ŀ����))
                    
                End If
                If rs��Ŀ.RecordCount = 0 Then Exit Function
                
                If g�������_�����山.�����־ = 3 Then
                    g�������_�����山.֧����� = Nvl(!֧�����)
                    g�������_�����山.�籣���칹������ = Nvl(!��Ա���)
                End If
                'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
                AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g�������_�����山.�籣���칹������, 1, 4)
                'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ09, ��ʶ��Сд���У�����λ��
                AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "09"
                'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
                AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
                'akc190  string  20      ������
                AppendXMLNode gobjXMLInPut.documentElement, "akc190", Nvl(!������)
                'yka105  number  15  0   ������ˮ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka105", Nvl(rsTemp!������ˮ��, 0)
                'yka002  string  20      ҽ����Ŀ����
                AppendXMLNode gobjXMLInPut.documentElement, "yka002", Nvl(!��Ŀ����)
                
                'yka103  string  20      ������
                AppendXMLNode gobjXMLInPut.documentElement, "yka103", Nvl(rsTemp!������)
                'yka104  number  15      �˵���Ӧ������ˮ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka104", Nvl(rsTemp!�˵���ˮ��)
                'aka130  string  6       ֧����𣬼������
                AppendXMLNode gobjXMLInPut.documentElement, "aka130", g�������_�����山.֧�����
                'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
                AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_�����山.ҽԺ����
                'ykb006  string  3       ����ҽ�ƻ�����֧�������
                AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"                '/***/��ȷ��������
                'aac001  number  15  0   ���˱��
                AppendXMLNode gobjXMLInPut.documentElement, "aac001", Nvl(!ҽ����)
                'akc226  number  14  4   ����
                AppendXMLNode gobjXMLInPut.documentElement, "akc226", Nvl(!����, 0)
                'akc225  number  14  4   ʵ�ʼ۸�
                AppendXMLNode gobjXMLInPut.documentElement, "akc225", Nvl(!ʵ�ʼ۸�, 0)
                'yka055  number  14  2   ҽ�Ʒ��ܶ�
                AppendXMLNode gobjXMLInPut.documentElement, "yka055", Nvl(!ʵ�ս��, 0)
                'yka096  number  14  4   �Ը�����
                AppendXMLNode gobjXMLInPut.documentElement, "yka096", Nvl(rsTemp!�Ը�����, 0)
                'yka056  number  14  2   ȫ�Էѽ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka056", Nvl(rsTemp!ȫ�Էѽ��, 0)
                'yka057  number  14  2   �ҹ��Էѽ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka057", Nvl(rsTemp!�ҹ��Էѽ��, 0)
                'yka111  number  14  2   ���Ϸ�Χ���
                AppendXMLNode gobjXMLInPut.documentElement, "yka111", Nvl(rsTemp!���Ϸ�Χ���, 0)
                'yka012  string  6       ������Ŀ���㷽ʽ���������
                AppendXMLNode gobjXMLInPut.documentElement, "yka012", "0"
                'yka098  string  50      ������������
                AppendXMLNode gobjXMLInPut.documentElement, "yka098", Nvl(!��������)
                'yka099  string  20      ����ҽ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka099", Nvl(!����ҽ��)
                'yka101  string  50      �ܵ���������
                AppendXMLNode gobjXMLInPut.documentElement, "yka101", Nvl(!��������)
                'yka102  string  20      �ܵ�ҽ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka102", Nvl(!����ҽ��)
                'aae036  date        ��  ����ʱ��
                AppendXMLNode gobjXMLInPut.documentElement, "aae036", Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS")
                'ykc166  date        ��  ��ϸ����ʱ��
                AppendXMLNode gobjXMLInPut.documentElement, "ykc166", Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS")
                'yab003  string  4       �籣�����������
                AppendXMLNode gobjXMLInPut.documentElement, "yab003", g�������_�����山.�籣���칹������
                'yka231  string  20      ��Ʒ������
                AppendXMLNode gobjXMLInPut.documentElement, "yka231", Nvl(rs��Ŀ!��Ʒ����)
                'yka247  String  20      �Է�ҩƷ��Ӧ��Ʒ������
                AppendXMLNode gobjXMLInPut.documentElement, "yka247", IIf(rs��Ŀ!������־ = 1, Nvl(rs��Ŀ!��׼���), Nvl(rs��Ŀ!��Ʒ����))
                'yka232  string  100     ��Ʒ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka232", Nvl(rs��Ŀ!��Ʒ��)
                'ykc130  string  6       ��ҩ���
                AppendXMLNode gobjXMLInPut.documentElement, "ykc130", "0" '/**/��ȷ��������
                'yka249  string  20      ����ҽ������
                AppendXMLNode gobjXMLInPut.documentElement, "yka249", Nvl(rsTemp!������)
                'yka250  string  20      ����ҽ��ְ��
                AppendXMLNode gobjXMLInPut.documentElement, "yka250", Nvl(rsTemp!������ְ��)
                'aae013  string  100     ��ע
                AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""        '/**
                'gkc013  string  6       ��Ŀ������־
                AppendXMLNode gobjXMLInPut.documentElement, "yka250", Nvl(rsTemp!������־, 0)
                'gkc014  string  50      ����
                AppendXMLNode gobjXMLInPut.documentElement, "gkc014", Nvl(rs��Ŀ!����, 0)
                'yka272  String  6       Ŀ¼����
                AppendXMLNode gobjXMLInPut.documentElement, "yka272", Nvl(rs��Ŀ!Ŀ¼����, 0)
                
                strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
                
                WriteDebugInfor_�����山 strXMLText
                
                'ҵ��������ϸ�ύ
                If ҵ������_�����山(������ϸд��, strXMLText, strOutPut) = False Then
                    blnTrue = False
                Else
                    '�����ϴ���־
                    'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
                End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    ������ϸ�ϴ� = blnTrue
    Exit Function
ErrHand:
  DebugTool "������ϸ�ϴ�����(������ϸ�ϴ�)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
End Function
Private Function Get��Ʊ����(ByVal lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ����ID,ʵ��Ʊ�� From ���˽��ʼ�¼ Where ID=" & lng����ID & " And Rownum<2"
    Call OpenRecordset(rsTemp, "��ȡ��Ʊ��")
    If rsTemp.EOF Then
    Get��Ʊ���� = ""
    Else
        Get��Ʊ���� = Nvl(rsTemp!ʵ��Ʊ��)
    End If

End Function


Public Function �������_�����山(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    
    Dim lng����id As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    
    �������_�����山 = False
    
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    WriteDebugDate_�����山 "===��    ��:�������"
    WriteDebugDate_�����山 "===��ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    
    g�������_�����山.�����־ = 0
    g�������_�����山.���� = False
    
    g�������_�����山.������ = Get������
    g�������_�����山.����ID = lng����ID
    g�������_�����山.��Ʊ�� = Get��Ʊ����(lng����ID)
    g�������_�����山.������� = False

  
    gstrSQL = "Select ����id, �Ǽ�ʱ�� From ���˷��ü�¼ where rownum<=1 and ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�Ǽ�ʱ��"
    
    If g�������_�����山.����ʱ�� > Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") Then
        g�������_�����山.����ʱ�� = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '���浱ǰ״̬�Ľ�����
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & Nvl(rsTemp!����ID, 0) & "," & TYPE_�����山 & ",'������','''" & g�������_�����山.������ & "''')"
    Call ExecuteProcedure("���������")
    
    gcnOracle_CQYB.BeginTrans
    
    
    If ���˽���(lng����ID) = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    
    gcnOracle_CQYB.CommitTrans
    �������_�����山 = True
    Exit Function
ErrHand:
    DebugTool "�������(�������_�����山)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
    gcnOracle_CQYB.RollbackTrans
End Function
Private Function Get����ID() As Long
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��ǰ������¼��IDֵ
    '--�����:
    '--������:
    '--��  ��:����ID
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    'ȡ������¼�Ľ���ID
    gstrSQL = "select distinct A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & g�������_�����山.����ID
    Call OpenRecordset(rsTemp, "���²����Ľ���ID")
    If rsTemp.EOF Then
        Get����ID = 0
    Else
        Get����ID = Nvl(rsTemp!����ID, 0)
    End If

End Function

Public Function ����������_�����山(lng����ID As Long, cur�����ʻ� As Currency, lng����id As Long) As Boolean
    

    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Err = 0
    On Error GoTo ErrHand
    
    ����������_�����山 = False
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    WriteDebugDate_�����山 "===��    ��:����������`"
    WriteDebugDate_�����山 "===��ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    
    '��ȡ������Ϣ
    Call Get������Ϣ(lng����id)
    
    g�������_�����山.����ID = lng����ID
    g�������_�����山.�����־ = 0
    g�������_�����山.���� = False
    g�������_�����山.������ = Get������
    g�������_�����山.���� = True
    g�������_�����山.����ID = Get����ID
    g�������_�����山.������� = False
    g�������_�����山.lng����id = lng����id
    
    '���浱ǰ״̬�Ľ�����
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����id & "," & TYPE_�����山 & ",'������','''" & lng����id & "''')"
    Call ExecuteProcedure("���������")
    
    gcnOracle_CQYB.BeginTrans
    ����������_�����山 = ���˽������(lng����ID)
    If ����������_�����山 = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    gcnOracle_CQYB.CommitTrans
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle_CQYB.RollbackTrans
    
End Function
Public Function ҽ������_�����山() As Boolean
    ҽ������_�����山 = frmSet�����山.��������
End Function

Public Function ��Ժ�Ǽ�_�����山(lng����id As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    Dim blnYes As Boolean
    On Error GoTo ErrHand
    
    
    If ����δ�����(lng����id, lng��ҳID) Then
        ShowMsgbox "���˴���δ�������,�Ƿ����?", True, blnYes
        If blnYes = False Then
            Exit Function
        End If
    End If
    g�������_�����山.���� = False
    g�������_�����山.������� = False
    g�������_�����山.�����־ = 2
    ''��ȡ�����
    'g�������_�����山.������ = Get������_�����山
    g�������_�����山.������ = Get������
    
    
    '���¾�����
'    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�����山 & ",'������','''" & g�������_�����山.������ & "''')"
'    Call ExecuteProcedure("���������")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����id & "," & TYPE_�����山 & ",'������','''" & g�������_�����山.������ & "''')"
    Call ExecuteProcedure("���������")
    
    '��ȡ��ز�����Ϣ
    '(SELECT A.����id,A.��ҳid,max(decode(A.���,1,B.����,'')) ��������,max(decode(A.���,1,A.������Ϣ,'')) AS ��������,max(decode(A.���,1,a.�п�,'')) �п�,max(decode(A.���,1,A.����,'')) ����    from ������� A ,��������Ŀ¼ B     where A.����ID=B.ID and A.���=1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & "    Group BY  A.����id,A.��ҳid ) E,
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,c.���￨�� as ������,c.סԺ��,to_char(A.ȷ������,'yyyy-MM-dd hh24:mi:ss') as ȷ������,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժ����ʱ��," & _
        " to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd') ��Ժ����  ,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd') ��Ժʱ��,D.��Ժ���,D.��Ժ���1,D.��Ժ���2,D.��Ժ���3,'' ��������,'' ��������,'' �п�,'' ����,'' ����ҽʦ,'' ����ҽʦ" & _
        " From ������ҳ A,���ű� B,������Ϣ C, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����||'-'||b.����,'')) AS ��Ժ���, max(DECODE(a.��ϴ���,2,b.����||'-'||b.����,'')) AS ��Ժ���1,max(DECODE(a.��ϴ���,3,b.����||'-'||b.����,'')) AS ��Ժ���2, max(DECODE(a.��ϴ���,4,b.����||'-'||b.����,'')) AS ��Ժ���3 From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����id & " Group by  ����id,��ҳid)   D" & _
        "        " & _
        "        " & _
        " Where A.����id=C.����id and C.����id=" & lng����id & _
        "       and A.����ID=" & lng����id & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        "" & _
        "       " & _
        ""
        'and A.��ҳid=F.��ҳid(+) and a.����id=F.����id(+)
        '(SELECT ����id,��ҳid,max(decode(��Ϣ��,'����ҽʦ',��Ϣֵ,'')) ����ҽʦ,max(decode(��Ϣ��,'����ҽʦ',��Ϣֵ,'')) ����ҽʦ from ������ҳ�ӱ� where ��ҳid=" & lng��ҳID & " and ����id=" & lng����ID & "    Group BY  ����id,��ҳid ) F
        'and A.��ҳid=E.��ҳid(+) and a.����id=E.����id(+)
    Call OpenRecordset(rsTemp, "��ȡ��Ժ��Ϣ")
    
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    '���������.
 
    If �ʸ���˴����˶�(lng����id, Format(rsTemp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss"), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) = False Then
        Exit Function
    End If
    
    '�ڶ���:д���ʸ���������,���������������ļ�
    If Save������Ϣ(lng����id, False) = False Then
        Exit Function
    End If
    
    Call intXML
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g�������_�����山.�籣���칹������
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ08, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "08"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   ���˱��
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g�������_�����山.���˱��
    'akc021  string  6       ҽ����Ա���
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g�������_�����山.ҽ����Ա���
    'akc190  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_�����山.ҽԺ����
    'ykb006  string  3       ����ҽ�ƻ�����֧�������
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    'aka130  string  6       ֧����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g�������_�����山.֧�����
    'akc192  date    ��      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", Nvl(rsTemp!��Ժ����)
    'akc193  string  100     ��Ժ���
    AppendXMLNode gobjXMLInPut.documentElement, "akc193", Nvl(rsTemp!��Ժ���)
    'ykc011  string  50      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc011", Nvl(rsTemp!��Ժ����)
    'ykc013  string  20      ��Ժ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc013", Nvl(rsTemp!������)
    'ykc014  date        ��  ��Ժ����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc014", Nvl(rsTemp!��Ժ����ʱ��)
    'akc195  string  6       ��Ժԭ�򣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "akc195", ""
    'akc194  date    ��      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", ""
    'akc196  string  100     ��Ժ���
    AppendXMLNode gobjXMLInPut.documentElement, "akc196", ""
    'ykc015  string  50      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc015", ""
    'ykc016  string  20      ��Ժ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc017  date        ��  ��Ժ����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc017", ""
    'ykc023  string  6       סԺ״̬
    '0-��Ժ,1-��Ժ 2-תԺ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc023", "0"
    'ykc009  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc009", Nvl(rsTemp!������)
    'ykc010  string  20      סԺ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc010", Nvl(rsTemp!סԺ��)
    'ykc149  string  100     ��Ժ���1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc149", Nvl(rsTemp!��Ժ���1)
    'ykc150  string  100     ��Ժ���2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc150", Nvl(rsTemp!��Ժ���2)
    'ykc151  string  100     ��Ժ���3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc151", Nvl(rsTemp!��Ժ���3)
    'ykc012  string  12      ��Ժ��λ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc012", Nvl(rsTemp!��ǰ����)
    'ykc152  string  100     ��Ժ���1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc152", ""
    'ykc153  string  100     ��Ժ���2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc153", ""
    'ykc154  string  100     ��Ժ���3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc154", ""
    'ykc016  string  12      ��Ժ��λ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc155  string  20      ��������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc155", Nvl(rsTemp!��������)
    
    'ykc156  string  100     ��������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc156", Nvl(rsTemp!��������)
    'ykc157  date        ��  ȷ��ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc157", Nvl(rsTemp!ȷ������)
    'ykc158  string  4       �����пڷ���
    AppendXMLNode gobjXMLInPut.documentElement, "ykc158", Nvl(rsTemp!�п�)
    'ykc159  string  4       �����п����ϼ���
    AppendXMLNode gobjXMLInPut.documentElement, "ykc159", Nvl(rsTemp!����)
    'ykc160  string  20      סԺҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc160", Nvl(rsTemp!סԺҽʦ)
    'ykc161  string  20      ����ҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc161", Nvl(rsTemp!����ҽʦ)
    'ykc162  string  20      ����ҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc162", Nvl(rsTemp!����ҽʦ)
    'aae013  string  100     ��ע
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""
    
    Dim strXMLText As String
    Dim strOutPut As String
    strXMLText = gobjXMLInPut.xml
    strXMLText = ȡ��XML��ǰ����ʶ(strXMLText)
        
        
        
    If ҵ������_�����山(������Ϣд��, strXMLText, strOutPut, "") = False Then
        ��Ժ�Ǽ�_�����山 = False
        Exit Function
    End If

    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & TYPE_�����山 & ")"
    Call ExecuteProcedure("������Ժ�Ǽ�")
    ��Ժ�Ǽ�_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�����山(lng����id As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    ShowMsgbox "��ҽ����֧����Ժ����!"
    Exit Function
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & TYPE_�����山 & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_�����山(lng����id As Long, lng��ҳID As Long) As Boolean
    
    Dim str��Ժԭ�� As String
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '1       ��Ժԭ�� ����
    '2       ��Ժԭ�� ��ת
    '3       ��Ժԭ�� δ��
    '4       ��Ժԭ�� ����
    '5       ��Ժԭ�� תԺ
    '9       ��Ժԭ�� ����
   '��ȡ������Ϣ
    Call Get������Ϣ(lng����id)

    
    '��ȡԭ������ˮ��
    
    gstrSQL = "Select ������,������ From �����ʻ� Where ����=" & TYPE_�����山 & " And ����ID=" & lng����id
    Call OpenRecordset(rsTemp, "��ȡ�����źͽ�����")
    
    g�������_�����山.������ = Nvl(rsTemp!������)
    g�������_�����山.������ = Nvl(rsTemp!������)
    
    '��ȡ��ز�����Ϣ
    
    gstrSQL = Get���SQL(lng����id, lng��ҳID)
    Call OpenRecordset(rsTemp, "��ȡ��Ժ��Ϣ")
    
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If

'    If �ʸ���˴����˶�(lng����ID, Format(rsTemp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss"), Format(rsTemp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")) = False Then
'        Exit Function
'    End If
'
'    '�ڶ���:д���ʸ���������,���������������ļ�
'    If Save������Ϣ(lng����ID, False) = False Then
'        Exit Function
'    End If

    Call intXML
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g�������_�����山.�籣���칹������
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ08, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "08"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   ���˱��
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g�������_�����山.���˱��
    'akc021  string  6       ҽ����Ա���
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g�������_�����山.ҽ����Ա���
    'akc190  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_�����山.ҽԺ����
    'ykb006  string  3       ����ҽ�ƻ�����֧�������
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    'aka130  string  6       ֧����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g�������_�����山.֧�����
    'akc192  date    ��      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", Nvl(rsTemp!��Ժ����)
    'akc193  string  100     ��Ժ���
    AppendXMLNode gobjXMLInPut.documentElement, "akc193", Nvl(rsTemp!��Ժ���)
    'ykc011  string  50      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc011", Nvl(rsTemp!��Ժ����)
    'ykc013  string  20      ��Ժ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc013", Nvl(rsTemp!������)
    'ykc014  date        ��  ��Ժ����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc014", Nvl(rsTemp!��Ժ����ʱ��)
    'akc195  string  6       ��Ժԭ�򣬼������
    str��Ժԭ�� = IIf(IsNull(rsTemp!��Ժ��ʽ), "", rsTemp!��Ժ��ʽ)
    '1       ��Ժԭ�� ����
    '2       ��Ժԭ�� ��ת
    '3       ��Ժԭ�� δ��
    '4       ��Ժԭ�� ����
    '5       ��Ժԭ�� תԺ
    '9       ��Ժԭ�� ����
    '��������ת��δ��������������
      Select Case str��Ժԭ��
      Case "����"
          str��Ժԭ�� = 1
      Case "��ת"
          str��Ժԭ�� = 2
      Case "δ��"
          str��Ժԭ�� = 3
      Case "����"
          str��Ժԭ�� = 4
      Case "תԺ"
          str��Ժԭ�� = 5
      Case Else
          str��Ժԭ�� = 9
      End Select
      
    AppendXMLNode gobjXMLInPut.documentElement, "akc195", str��Ժԭ��
    'akc194  date    ��      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", Nvl(rsTemp!��Ժ����)
    'akc196  string  100     ��Ժ���
    AppendXMLNode gobjXMLInPut.documentElement, "akc196", Nvl(rsTemp!��Ժ���)
    'ykc015  string  50      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc015", Nvl(rsTemp!��Ժ����)
    'ykc016  string  20      ��Ժ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", Nvl(rsTemp!����Ա)
    'ykc017  date        ��  ��Ժ����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc017", Format(rsTemp!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
    'ykc023  string  6       סԺ״̬
    '0-��Ժ,1-��Ժ 2-תԺ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc023", IIf(str��Ժԭ�� = "5", "2", "1")
    'ykc009  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc009", Nvl(rsTemp!������)
    'ykc010  string  20      סԺ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc010", Nvl(rsTemp!סԺ��)
    'ykc149  string  100     ��Ժ���1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc149", Nvl(rsTemp!��Ժ���1)
    'ykc150  string  100     ��Ժ���2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc150", Nvl(rsTemp!��Ժ���2)
    'ykc151  string  100     ��Ժ���3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc151", Nvl(rsTemp!��Ժ���3)
    'ykc012  string  12      ��Ժ��λ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc012", Nvl(rsTemp!��Ժ����)
    'ykc152  string  100     ��Ժ���1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc152", Nvl(rsTemp!��Ժ���1)
    'ykc153  string  100     ��Ժ���2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc153", Nvl(rsTemp!��Ժ���2)
    'ykc154  string  100     ��Ժ���3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc154", Nvl(rsTemp!��Ժ���3)
    'ykc016  string  12      ��Ժ��λ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", Nvl(rsTemp!��Ժ����)
    'ykc155  string  20      ��������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc155", Nvl(rsTemp!��������)
    
    'ykc156  string  100     ��������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc156", Nvl(rsTemp!��������)
    'ykc157  date        ��  ȷ��ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc157", IIf(Nvl(rsTemp!ȷ������) = "", Format(rsTemp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss"), Nvl(rsTemp!ȷ������))
    'ykc158  string  4       �����пڷ���
    AppendXMLNode gobjXMLInPut.documentElement, "ykc158", Nvl(rsTemp!�п�)
    'ykc159  string  4       �����п����ϼ���
    AppendXMLNode gobjXMLInPut.documentElement, "ykc159", Nvl(rsTemp!����)
    'ykc160  string  20      סԺҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc160", Nvl(rsTemp!סԺҽʦ)
    'ykc161  string  20      ����ҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc161", Nvl(rsTemp!����ҽʦ)
    'ykc162  string  20      ����ҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc162", Nvl(rsTemp!����ҽʦ)
    'aae013  string  100     ��ע
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""
    
    Dim strXMLText As String
    Dim strOutPut As String
    strXMLText = gobjXMLInPut.xml
    strXMLText = ȡ��XML��ǰ����ʶ(strXMLText)
    
    If ҵ������_�����山(������Ϣд��, strXMLText, strOutPut, "") = False Then
        ��Ժ�Ǽ�_�����山 = False
          Exit Function
    End If
     
    If Not ����δ�����(lng����id, lng��ҳID) Then
        '���������δ�����,��ɾ��������Ϣ
        '��������¼����
        Call ������¼����_�����山
    End If
     
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & TYPE_�����山 & ")"
    Call ExecuteProcedure("��Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get���SQL(ByVal lng����id As Long, ByVal lng��ҳID As Long) As String
    Dim strSql As String
    
    strSql = "Select C.סԺ��,C.��ǰ����id,A.��Ժ���� ,c.���￨�� as ������,c.סԺ��,to_char(A.ȷ������,'yyyy-MM-dd hh24:mi:ss') as ȷ������,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժ����ʱ��," & _
        " to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժʱ��,J.��ֹʱ��,J.����Ա,D.��Ժ���,D.��Ժ���1,D.��Ժ���2,D.��Ժ���3,A.��Ժ��ʽ,to_Char(a.��Ժ����,'yyyy-MM-DD') as ��Ժ����,a.��Ժ���� as ��Ժʱ��,a.��Ժ����,H.���� as ��Ժ����,'' ��������,'' ��������,'' �п�,'' ����,'' ����ҽʦ,'' ����ҽʦ,G.��Ժ���,G.��Ժ���1,g.��Ժ���2,g.��Ժ���3" & _
        " From ������ҳ A,���ű� B,������Ϣ C,���ű� H, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����||'-'||b.����,'')) AS ��Ժ���, max(DECODE(a.��ϴ���,2,b.����||'-'||b.����,'')) AS ��Ժ���1,max(DECODE(a.��ϴ���,3,b.����||'-'||b.����,'')) AS ��Ժ���2, max(DECODE(a.��ϴ���,4,b.����||'-'||b.����,'')) AS ��Ժ���3 From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1  and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����id & " Group by ����id,��ҳid)   D," & _
        "        " & _
        "       (Select ����id,��ҳid,Max(��ֹʱ��) as ��ֹʱ��,max(����Ա����) ����Ա From ���˱䶯��¼ where  ��ֹԭ��=1 and ����id=" & lng����id & " and ��ҳid=" & lng��ҳID & " Group by ����id,��ҳid) J," & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����||'-'||b.����,'')) AS ��Ժ���, max(DECODE(a.��ϴ���,2,b.����||'-'||b.����,'')) AS ��Ժ���1,max(DECODE(a.��ϴ���,3,b.����||'-'||b.����,'')) AS ��Ժ���2, max(DECODE(a.��ϴ���,4,b.����||'-'||b.����,'')) AS ��Ժ���3 From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� = 3 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����id & " Group by ����id,��ҳid)   G" & _
        " Where A.����id=C.����id and C.����id=" & lng����id & _
        "       and A.����ID=" & lng����id & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID and A.��Ժ����ID=H.id(+) " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        "       and A.��ҳid=J.��ҳid(+) and a.����id=J.����id(+)" & _
        "       " & _
        "       and A.��ҳid=G.��ҳid(+) and a.����id=G.����id(+) " & _
        ""
    Get���SQL = strSql
End Function

Public Function ��Ժ�Ǽǳ���_�����山(lng����id As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
     
     '��ȡ��ز�����Ϣ
      '��ȡ������Ϣ
    
    '�����ڽ��ʺ�Ҫ����IC��,���Բ��ܶ��Ѿ������˵Ĳ��˽���ȡ������
    
    If Not ����δ�����(lng����id, lng��ҳID) Then
        ShowMsgbox "���ܶԲ�����δ����õĲ��˽��г�����Ժ,�����°�����Ժ."
        Exit Function
    End If
    
    Call Get������Ϣ(lng����id)
 
    gstrSQL = Get���SQL(lng����id, lng��ҳID)
    Call OpenRecordset(rsTemp, "��ȡ��Ժ��Ϣ")
    
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    Call intXML
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g�������_�����山.�籣���칹������
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ08, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "08"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   ���˱��
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g�������_�����山.���˱��
    'akc021  string  6       ҽ����Ա���
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g�������_�����山.ҽ����Ա���
    'akc190  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_�����山.ҽԺ����
    'ykb006  string  3       ����ҽ�ƻ�����֧�������
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    'aka130  string  6       ֧����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g�������_�����山.֧�����
    'akc192  date    ��      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", Nvl(rsTemp!��Ժ����)
    'akc193  string  100     ��Ժ���
    AppendXMLNode gobjXMLInPut.documentElement, "akc193", Nvl(rsTemp!��Ժ���)
    'ykc011  string  50      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc011", Nvl(rsTemp!��Ժ����)
    'ykc013  string  20      ��Ժ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc013", Nvl(rsTemp!������)
    'ykc014  date        ��  ��Ժ����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc014", Nvl(rsTemp!��Ժ����ʱ��)
    'akc195  string  6       ��Ժԭ�򣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "akc195", ""
    'akc194  date    ��      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", ""
    'akc196  string  100     ��Ժ���
    AppendXMLNode gobjXMLInPut.documentElement, "akc196", ""
    'ykc015  string  50      ��Ժ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc015", ""
    'ykc016  string  20      ��Ժ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc017  date        ��  ��Ժ����ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc017", ""
    'ykc023  string  6       סԺ״̬
    
    '0-��Ժ,1-��Ժ 2-תԺ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc023", "0"
    'ykc009  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc009", Nvl(rsTemp!������)
    'ykc010  string  20      סԺ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc010", Nvl(rsTemp!סԺ��)
    'ykc149  string  100     ��Ժ���1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc149", Nvl(rsTemp!��Ժ���1)
    'ykc150  string  100     ��Ժ���2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc150", Nvl(rsTemp!��Ժ���2)
    'ykc151  string  100     ��Ժ���3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc151", Nvl(rsTemp!��Ժ���3)
    'ykc012  string  12      ��Ժ��λ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc012", Nvl(rsTemp!��Ժ����)
    'ykc152  string  100     ��Ժ���1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc152", ""
    'ykc153  string  100     ��Ժ���2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc153", ""
    'ykc154  string  100     ��Ժ���3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc154", ""
    'ykc016  string  12      ��Ժ��λ
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc155  string  20      ��������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc155", Nvl(rsTemp!��������)
    
    'ykc156  string  100     ��������
    AppendXMLNode gobjXMLInPut.documentElement, "ykc156", Nvl(rsTemp!��������)
    'ykc157  date        ��  ȷ��ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "ykc157", Nvl(rsTemp!ȷ������)
    'ykc158  string  4       �����пڷ���
    AppendXMLNode gobjXMLInPut.documentElement, "ykc158", Nvl(rsTemp!�п�)
    'ykc159  string  4       �����п����ϼ���
    AppendXMLNode gobjXMLInPut.documentElement, "ykc159", Nvl(rsTemp!����)
    'ykc160  string  20      סԺҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc160", Nvl(rsTemp!סԺҽʦ)
    'ykc161  string  20      ����ҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc161", Nvl(rsTemp!����ҽʦ)
    'ykc162  string  20      ����ҽʦ����
    AppendXMLNode gobjXMLInPut.documentElement, "ykc162", Nvl(rsTemp!����ҽʦ)
    'aae013  string  100     ��ע
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""
    
    Dim strXMLText As String
    Dim strOutPut As String
    strXMLText = gobjXMLInPut.xml
    strXMLText = ȡ��XML��ǰ����ʶ(strXMLText)
    
    If ҵ������_�����山(������Ϣд��, strXMLText, strOutPut, "") = False Then
        ��Ժ�Ǽǳ���_�����山 = False
        Exit Function
    End If
    
'    If Not ����δ�����(lng����ID, lng��ҳID) Then
'
'        If �ʸ���˴����˶�(lng����ID, Format(rsTemp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss"), Format(zlDataBase.Currentdate, "yyyy-MM-dd HH:mm:ss")) = False Then
'            Exit Function
'        End If
'        '�ڶ���:д���ʸ���������,���������������ļ�
'        If Save������Ϣ(lng����ID, False) = False Then
'            Exit Function
'        End If
'    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����id & "," & TYPE_�����山 & ")"
    Call ExecuteProcedure("��������Ժ�Ǽ�")
    
    ��Ժ�Ǽǳ���_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_�����山(ByVal lng����id As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    
    '����ʧ�����˳�
    gstrSQL = "Select Nvl(�ʻ����,0) �ʻ����,����֤�� From �����ʻ� Where ����=" & gintInsure
    gstrSQL = gstrSQL & " And ����id=" & lng����id
    
    Call OpenRecordset(rsTemp, "��ȡ�ʻ����")
    
    With g�������_����
        �������_�����山 = Nvl(rsTemp!�ʻ����, 0)
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function סԺ�������_�����山(rsExse As Recordset, ByVal lng����id As Long, Optional bln���ʴ� As Boolean = True) As String

    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim intMouse As Integer
    
    סԺ�������_�����山 = ""
    
    DebugTool "����סԺ�������"
    
    Call Get������Ϣ(lng����id)
    DebugTool "�Ѿ���ȡ������Ϣ,�����������֤."
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    WriteDebugDate_�����山 "===��    ��:סԺ�������"
    WriteDebugDate_�����山 "===��ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    
    If bln���ʴ� Then
        '�����½����鿨
        intMouse = Screen.MousePointer
        Screen.MousePointer = 1
        If Trim(frmIdentify�����山.GetPatient(4, 0)) = "" Then
            Exit Function
        End If
        Screen.MousePointer = intMouse
    Else
        '�����ط�
    End If
    
    
    g�������_�����山.�����־ = 1
    g�������_�����山.���� = False
    
    g�������_�����山.�����ܶ� = 0
    
    '�󱾴��ܶ�
    DebugTool "ȡ���η����ܶ�"
    With rsExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            g�������_�����山.�����ܶ� = g�������_�����山.�����ܶ� + Nvl(rsExse!���, 0)
            .MoveNext
        Loop
    End With
    
    'ȡ����ʱ��
    DebugTool "ȡ������Ϣ!"
    gstrSQL = "Select ����ʱ��,������,������ From �����ʻ� where ����=" & TYPE_�����山 & " and ����id=" & lng����id
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����ʱ��"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "�ڱ����ʻ��в����ڴ�ҽ������!"
        Exit Function
    End If
    
    g�������_�����山.����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    g�������_�����山.������ = Nvl(rsTemp!������)
    g�������_�����山.������ = Nvl(rsTemp!������)
    g�������_�����山.lng����id = lng����id
    
    '������ڷ��ü�¼�е��Զ����ʲ���,�����ڱ���
    gstrSQL = "" & _
        "   Select ID,NO,��¼״̬,��¼����,��� From ���˷��ü�¼ " & _
        "   Where ����id=" & lng����id & " and id not in(Select a.id From ���˷��ü�¼ a,ҽ����ϸ���� b Where a.id=b.����id And a.����id=" & lng����id & ") " & _
        "   Order by ��¼����,NO,��¼״̬,���"
    
    DebugTool "�����Զ����ʲ��ֵ���ϸ��¼(���м�����)!"
    
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ�Զ�������ϸ��¼"
    With rs��ϸ
        .Filter = "��¼״̬<>2"
        g�������_�����山.���� = False
        Do While Not .EOF
             IsertIntoҽ����ϸ !ID, Nvl(!NO), Nvl(!���, 0), Nvl(!��¼����, 0), ""
            .MoveNext
        Loop
        .Filter = "��¼״̬=2"
        If .RecordCount <> 0 Then .MoveFirst
        g�������_�����山.���� = True
        
        '������˵�����ˮ��
        Do While Not .EOF
            IsertIntoҽ����ϸ !ID, Nvl(!NO), Nvl(!���, 0), Nvl(!��¼����, 0), ""
            .MoveNext
        Loop
    End With
        
    '��������ҽ����ϸ���ü�¼�ľ����ż�������
    Err = 0
    On Error Resume Next
    
    DebugTool "���±��ν���Ľ�����(���м��ҽ����ϸ����)!"
    gcnOracle_CQYB.Execute "UPdate ҽ����ϸ���� set ������='" & g�������_�����山.������ & "' where ������ is null and ������='" & g�������_�����山.������ & "'"
    
    If Err <> 0 Then
        ShowMsgbox "�ڸ���ҽ������ʱ����!"
        Exit Function
    End If
    
    g�������_�����山.���� = False
    g�������_�����山.������� = True
    
    gcnOracle_CQYB.BeginTrans
    
    DebugTool "���벡�˽���!"
    If ���˽���(0) = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    DebugTool "�������!"
    סԺ�������_�����山 = g�������_�����山.������Ϣ
    gcnOracle_CQYB.CommitTrans
    Exit Function
ErrHand:
    DebugTool "�������ʱ��������" & vbCrLf & " �����:" & Err.Number & vbCrLf & "������Ϣ:" & Err.Description
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�����山(lng����ID As Long, ByVal lng����id As Long) As Boolean

    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    WriteDebugDate_�����山 "===��    ��:סԺ����"
    WriteDebugDate_�����山 "===��ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_�����山 "================================================================================================================================================================================================================================================"
    
    סԺ����_�����山 = False
    
    Call Get������Ϣ(lng����id)
    
    g�������_�����山.����ID = lng����ID
    g�������_�����山.�����־ = 1
    g�������_�����山.���� = False
    g�������_�����山.��Ʊ�� = Get��Ʊ����(lng����ID)
    g�������_�����山.lng����id = lng����id
    
    '�󱾴ν���ķ����ܶ�
    gstrSQL = "Select Sum(nvl(���ʽ��,0)) as �ܷ��� From ���˷��ü�¼ where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�ܷ���"
    g�������_�����山.�����ܶ� = Nvl(rsTemp!�ܷ���, 0)
    
    'ȡ����ʱ��
    gstrSQL = "Select ����ʱ��,������,������ From �����ʻ� where ����=" & TYPE_�����山 & " and ����id=" & lng����id
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����ʱ��"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "�ڱ����ʻ��в����ڴ�ҽ������!"
        Exit Function
    End If
    
    g�������_�����山.����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    g�������_�����山.������ = Nvl(rsTemp!������)
    g�������_�����山.������ = Nvl(rsTemp!������)
    g�������_�����山.������� = False
    
       
    Err = 0
    On Error GoTo ErrHand
    
    gcnOracle_CQYB.BeginTrans
    
    If ���˽���(lng����ID) = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    '������ɺ󣬽���ǰ�Ľ�������Ϊ��.
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����id & "," & TYPE_�����山 & ",'������','''" & "" & "''')"
    Call ExecuteProcedure("���������")
    gcnOracle_CQYB.CommitTrans
    
    סԺ����_�����山 = True
    Exit Function
ErrHand:
    DebugTool "סԺ����(סԺ����_�����山)" & vbCrLf & " �����:" & Err & vbCrLf & "������Ϣ:" & Err.Description
    gcnOracle_CQYB.RollbackTrans
End Function

Public Function סԺ�������_�����山(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    MsgBox "ҽ����֧�ֽ������ϣ���ֱ�����ϼ��ʵ��ݺ��ٽ��ʣ�", vbInformation, gstrSysName
    סԺ�������_�����山 = False
End Function
Public Function �����Ǽ�_�����山(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng����id As Long
    Dim blnUpload As Boolean
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand:
    
    
    g�������_�����山.�����־ = 3
    g�������_�����山.���� = lng��¼״̬ <> 1
    
    
    
    '���·�������
    gstrSQL = " Select distinct a.����id,b.������,b.������  From ���˷��ü�¼ a,�����ʻ� b  where a.no='" & str���ݺ� & "'   and a.��¼״̬=" & lng��¼״̬ & " and a.��¼����=" & lng��¼���� & " and a.����id=b.����id  and b.����=" & gintInsure
            
    OpenRecordset rsTemp, "��ȡ���������Ϣ"
    '��Ҫ���Ǽ��ʱ�����ֱ��
    
    Do While Not rsTemp.EOF
        If IsNull(rsTemp!������) Then
            g�������_�����山.������ = Get������
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & Nvl(rsTemp!����ID, 0) & "," & TYPE_�����山 & ",'������','''" & g�������_�����山.������ & "''')"
            Call ExecuteProcedure("���������")
        End If
        rsTemp.MoveNext
    Loop
    
    
    '��һ��: ��ȡ������ϸ��¼
    gstrSQL = Get��ϸ��¼(0, str���ݺ�, lng��¼����, lng��¼״̬)
    
    Call OpenRecordset(rs��ϸ, "��ȡ������ϸ")
    If rs��ϸ.RecordCount = 0 Then
        ShowMsgbox "û����ϸ��¼�����������Ŀδ������Ӧ�Ķ���"
        Exit Function
    End If
    gcnOracle_CQYB.BeginTrans
    
    If Saveҽ����ϸ����(rs��ϸ) = False Then Exit Function
    
    '�ڶ���:������ϸ�ϴ�
    If ������ϸ�ϴ�(rs��ϸ) = False Then
        ShowMsgbox "�ڽ��д�����ϸ�ϴ�ʱ����һ�����ϵ���ϸ�ϴ�ʧ��,���Ժ�ע�ⲹ��!"
    End If
    gcnOracle_CQYB.CommitTrans
    �����Ǽ�_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnOracle_CQYB.RollbackTrans
End Function

'---------------------------------------------------------------------------------------------------------------------------
'����:
'
'---------------------------------------------------------------------------------------------------------------------------
Public Function ҵ������_�����山(ByVal intҵ������ As ҵ������_�����山, ByVal strInputString As String, ByRef strOutPutstring As String, Optional ByRef strErrMsg As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮
    '         strOutPutString-�����
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInPut As String
    Dim strOutPut As String
    Dim AppStruct As Struct
    Dim blnReturn As Boolean    '���ش���
    Dim intAppCode As Integer
    
    
    
    strInPut = strInputString
    
    strOutPut = ""
    If InitInfor_�����山.ģ������ Then
        '��ȡģ������
        Readģ������ intҵ������, strInputString, strOutPutstring
         ҵ������_�����山 = True
        Exit Function
    End If
  
    AppStruct.strErrMsg = Space(4500)
    strOutPut = ""
    
    'ҵ������
    'blnReturn = DataUpload(strInPut, strOutput, AppStruct)
    '���½ӿڶ���
    
     
    Err = 0
    On Error GoTo ErrHand:
    
    blnReturn = gobjYingHaiDll.dll_main_in(strInPut, strOutPut, intAppCode, strErrMsg)
    
    If blnReturn = False Then
      '��������,��ʾ����Ϣ
        ShowMsgbox strErrMsg
        ҵ������_�����山 = False
        Exit Function
    End If
    strOutPutstring = strOutPut
    ҵ������_�����山 = True
    
    Exit Function
    
    
    strErrMsg = ""
    If AppStruct.lngAppCode = 1 Then
        ҵ������_�����山 = True
    ElseIf AppStruct.lngAppCode < 0 Then
        '��������,��ʾ����Ϣ
        ShowMsgbox AppStruct.strErrMsg
        strErrMsg = AppStruct.strErrMsg
        ҵ������_�����山 = False
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Readģ������(ByVal intҵ������ As ҵ������_�����山, ByVal strInputString As String, ByRef strOutPutstring As String)
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
    
    strFile = App.Path & "\ģ���ύ��.txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
        
    Select Case intҵ������
        Case ��ȡϵͳʱ��
            strTemp = "��ȡϵͳʱ��"
            '�Ա���ʱ��Ϊ׼
        Case ��ݼ���
            strTemp = "��ݼ���"
        Case �޸�����
            strTemp = "�޸�����"
        Case IC���ʻ�֧��
            strTemp = "IC���ʻ�֧��"
        Case �ʸ����������˶�
            strTemp = "�ʸ����������˶�"
        Case ������Ϣд��
            strTemp = "������Ϣд��"
        Case ������ϸд��
            strTemp = "������ϸд��"
        Case ���������Ϣд��
            strTemp = "���������Ϣд��"
        Case ������д��
            strTemp = "������д��"
        Case �˶��ʻ�֧����Ϣ
            strTemp = "�˶��ʻ�֧����Ϣ"
        Case �˶Ծ�����Ϣ
            strTemp = "�˶Ծ�����Ϣ"
        Case �˶Դ�����ϸ��Ϣ
            strTemp = "�˶Դ�����ϸ��Ϣ"
        Case �˶Է��ý�����
            strTemp = "�˶Է��ý�����"
        Case �˶Է��ý��������Ϣ
            strTemp = "�˶Է��ý��������Ϣ"
        Case ����������ĿĿ¼
            strTemp = "����������ĿĿ¼"
        Case ����ICD_10��Ϣ
            strTemp = "����ICD_10��Ϣ"
        Case ��������Ŀ¼
            strTemp = "��������Ŀ¼"
        Case �������־��������Ϣ
            strTemp = "�������־��������Ϣ"
        Case ����ҽ��������Ϣ
            strTemp = "����ҽ��������Ϣ"
        Case ��ȡ�ͻ�����ʶ��
            strTemp = "��ȡ�ͻ�����ʶ��"
        Case ������������
            strTemp = "������������"
        Case ��ȡ������
            strTemp = "��ȡ������"
        Case ��ȡ������ˮ��
            strTemp = "��ȡ������ˮ��"
        Case ��ȡ������
            strTemp = "��ȡ������"
        Case ���ý���
            strTemp = "���ý���"
        Case ����������¼
            strTemp = "����������¼"
    End Select
    
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine "[" & strTemp & "]"
    objText.WriteLine strInputString
    objText.Close
    If intҵ������ = ������������ Then
        strFile = App.Path & "\������.txt"
    Else
        strFile = App.Path & "\ҽ��ģ������.txt"
    End If
    
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
                If intҵ������ = ������������ Then
                    strArr = Split(strText, vbTab)
                    If Val(strArr(0)) = 1 Then
                        With g�������_�����山
                             .���˱�� = strArr(1)
                             .���� = strArr(2)
                         End With
                        str = strArr(1) & vbTab & strArr(2)
                        Exit Do
                    End If
                Else
                    If blnStart Then
                        If strText = "" Then
                            strText = "" & vbTab
                        End If
                        strArr = Split(strText, vbTab)
                        If strArr(0) = strInputString Then
                            str = strArr(1)
                            Exit Do
                        End If
                   Else
                        If "<" & strTemp & ">" = strText Then
                            blnStart = True
                        End If
                   End If
                    If "</" & strTemp & ">" = strText Then
                        Exit Do
                    End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    
    End If
    
    Exit Function
ErrHand:
    DebugTool Err.Description
   
End Function
Private Function GetXML��(ByVal strInputXMLString As String, Optional blnLoadRoot As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��XML������Խ�gobjOutput
    '--�����:blnLoadRoot-�Ƿ��Զ�����Root�ӵ�
    '--������:
    '--��  ��:���سɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    
    If blnLoadRoot Then
        strXMLText = "<" & gstrXMLRootPart & ">" & strInputXMLString & "</" & gstrXMLRootPart & ">"
    Else
        strXMLText = strInputXMLString
    End If
    
    GetXML�� = gobjXMLOutput.loadXML(strXMLText)
End Function
Private Function intXML() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼXML
    '--�����:
    '--������:
    '--��  ��:��ʼ�ɹ�����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim nodData As MSXML2.IXMLDOMElement
    
       
    On Error Resume Next
    Set gobjXMLInPut = New MSXML2.DOMDocument
    Set gobjXMLOutput = New MSXML2.DOMDocument
    If Err <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set nodData = gobjXMLInPut.createElement(gstrXMLRootPart)
    Set gobjXMLInPut.documentElement = nodData
    intXML = True
End Function
Private Function AppendXMLNode(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String) As MSXML2.IXMLDOMElement
    '���ܣ���ָ��XMLԪ����������Ԫ��
    Set AppendXMLNode = gobjXMLInPut.createElement(Name)
    AppendXMLNode.Text = Value
    nodParent.appendChild AppendXMLNode
End Function
Public Function GetXMLOutput(ByVal Name As String, Optional blnName As Boolean = True, Optional lngRow As Long = 0) As String

    '���ܣ��õ�ָ��Ԫ�ص�ֵ
    '����:blnName-����������ȡֵ
    Dim xmlElement As MSXML2.IXMLDOMElement
    If blnName Then
        Set xmlElement = gobjXMLOutput.getElementsByTagName(Name).Item(lngRow)
    Else
        Set xmlElement = gobjXMLOutput.documentElement.selectSingleNode(Name)
    End If
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        GetXMLOutput = xmlElement.Text
    End If
End Function
Public Function �޸�����_�����山(ByVal strOldPassWord As String, ByVal strNewPassWord As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:Զ���޸�ҽ����Ա���ʻ�����
    '--�����:
    '--������:
    '--��  ��:�ɹ�true,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutPut As String
    Dim strXMLText As String
    Dim blnReturn As Boolean
    Err = 0
    On Error GoTo ErrHand:
    
    �޸�����_�����山 = False
    If intXML = False Then Exit Function
        
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g�������_�����山.�籣���칹������, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "04"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g�������_�����山.����, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(strOldPassWord, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "New_ykc005", Substr(strNewPassWord, 1, 6)
    
    strXMLText = gobjXMLInPut.documentElement.xml
    'ȡ��ǰ��XML��
    strXMLText = Mid(strXMLText, Len(gstrXMLRootPart) + 3, Len(strXMLText) - 3)
        
    'ҵ������
    
    blnReturn = ҵ������_�����山(�޸�����, strXMLText, strOutPut)
    If blnReturn = False Then
        Exit Function
    End If
    
    '�����
    �޸�����_�����山 = True
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
    �޸�����_�����山 = False
End Function
Public Function ������_�����山() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������������
    '--�����:strCardData-��������
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO As String
    Dim strCardNO As String
    Dim strExInfor As String
    Err = 0
    On Error GoTo ErrHand:
    If InitInfor_�����山.������������ Then
        Readģ������ ������������, "", strExInfor
        
    Else
        strNO = Space(10)
        strCardNO = Space(12)
        strExInfor = Space(4)
        Call srd_4428_info(strNO, strCardNO, strExInfor)
        With g�������_�����山
            .���˱�� = strNO
            .���� = strCardNO
        End With
    End If
    ������_�����山 = True
    Exit Function
ErrHand:
    ������_�����山 = False
    ShowMsgbox "IC������,����ʶ��!"
End Function
Public Function �˶Բ��˾�����Ϣ_�����山(ByVal lng����id As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�˶Բ��˾�����Ϣ
    '--�����:
    '--������:
    '--��  ��:���غ˶Է�Χ�ڵļ�¼��
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    
    Call Get������Ϣ(lng����id)
    
    Call intXML
    
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "14"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    If ҵ������_�����山(�˶Ծ�����Ϣ, strXMLText, strOutPut, "") = False Then
        ShowMsgbox "��ȡ�˶Ծ�����Ϣʱ��Ժ"
        Exit Function
    End If
    
    If GetXML��(strOutPut) = False Then
        ShowMsgbox "�˶Ծ�����Ϣ�з��ش�����һ����Ч��XML����"
        Exit Function
    End If
    lngCount = Val(GetXMLOutput("RecordCount"))
    
    
    '���������¼��
    gstrSQL = "Select count(distinct a.����id||' '||a.��ҳid)  as ���� From ������ҳ a,�����ʻ� b where a.����id=b.����id "
    
    OpenRecordset rsTemp, "��ȡҽԺҽ�����˵������Ϣ"
    
    strTemp = "���ļ�¼��Ϊ:" & lngCount & "|" & Nvl(rsTemp!����, 0)
    frmShowMsg.ShowInFor strTemp
End Function

Public Function �˶Բ����ʻ�֧����Ϣ_�����山(ByVal lng����id As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�˶Բ��˾�����Ϣ
    '--�����:
    '--������:
    '--��  ��:���غ˶Է�Χ�ڵļ�¼��
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    Dim dbl�ܶ� As Double
    Dim strTemp As String
    
    Get������Ϣ lng����id
    
    Call intXML
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ13, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "13"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       �籣�����������
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    'akc190  string  20      ��Ҫ�˶Ե��˻�֧����Ϣ�ľ�����
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    
    
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    If ҵ������_�����山(�˶��ʻ�֧����Ϣ, strXMLText, strOutPut, "") = False Then
        ShowMsgbox "�˶��ʻ�֧����Ϣʱ,ҵ������ʧ�ܣ�"
        Exit Function
    End If
    
    If GetXML��(strOutPut) = False Then
        ShowMsgbox "�˶��ʻ�֧����Ϣ�з��ش�����һ����Ч��XML����"
        Exit Function
    End If
    'RecordCount number  15      �ں˶Է�Χ�ڵ�������Ϣ�ļ�¼����
    lngCount = Val(GetXMLOutput("RecordCount"))
    'DefrayAmount    string  14  2   �ں˶Է�Χ�ڵ����м�¼���ʻ�������֧���ܶ��ۼ�ֵ
    dbl�ܶ� = Val(GetXMLOutput("DefrayAmount"))
    
    '���������¼��
    gstrSQL = "Select count(��¼ID) as ��¼��,sum(�����ʻ�֧��) as �ܶ�  From ���ս����¼ where nvl(�����ʻ�֧��,0)<>0 and ��ע='" & g�������_�����山.������ & "'"
    OpenRecordset rsTemp, "��ȡҽԺ�˶��ʻ�֧����Ϣ"
    
    strTemp = "��¼��Ϊ:" & lngCount & "|" & Nvl(rsTemp!��¼��, 0) & "||֧���ܶ�Ϊ:" & Format(dbl�ܶ�, "####0.00;-####0.00; ;") & "Ԫ|" & Format(Nvl(rsTemp!�ܶ�, 0), "####0.00;-####0.00; ;") & "Ԫ"
    frmShowMsg.ShowInFor strTemp
End Function


Public Function �˶Բ��˷��ý��������Ϣ_�����山(ByVal lng����id As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�˶Է��ý��������Ϣ
    '--�����:
    '--������:
    '--��  ��:���غ˶Է�Χ�ڵļ�¼��
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    Dim dbl�ܶ� As Double
    Get������Ϣ lng����id
    Call intXML
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ15, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "13"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       �籣�����������
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    'akc190  string  20      ��Ҫ�˶Ե��˻�֧����Ϣ�ľ�����
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    
    
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    If ҵ������_�����山(�˶Է��ý��������Ϣ, strXMLText, strOutPut, "") = False Then
        ShowMsgbox "�˶Է��ý��������Ϣʱ,ҵ������ʧ�ܣ�"
        Exit Function
    End If
    
    If GetXML��(strOutPut) = False Then
        ShowMsgbox "�˶Է��ý��������Ϣ�з��ش�����һ����Ч��XML����"
        Exit Function
    End If
    gstrSQL = "Select count(distinct ������ ) as ��¼��,0 as �����ܶ�,sum(ȫ�Էѽ��) as ȫ�Է�,sum(�ҹ��Է�) as �ҹ��Է�,sum(���Ͻ��) as ���Ͻ��,sum(�����Ը�) as �����ʻ�,0 as �����ֽ�    From ���ý����� where ������='" & g�������_�����山.������ & "'"
    OpenRecordset_ZLYB rsTemp, "��ȡ���ý��������Ϣ"
    
    'RecordCount number  15      �ں˶Է�Χ�ڵ�������Ϣ�ļ�¼����
    strTemp = "��¼��:" & Val(GetXMLOutput("RecordCount")) & "|" & Nvl(rsTemp!��¼��, 0)
    'yka055  number  14  2   �ں˶Է�Χ�ڵ����м�¼��ҽ�Ʒ��ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "ҽ�Ʒ����ܶ�:" & Format(Val(GetXMLOutput("yka055")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!�����ܶ�, 0)
    'yka056  number  14  2   �ں˶Է�Χ�ڵ����м�¼��ȫ�Է��ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "ȫ�Է�  �ܶ�:" & Format(Val(GetXMLOutput("yka056")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!ȫ�Է�, 0)
    'yka057  number  14  2   �ں˶Է�Χ�ڵ����м�¼�Ĺҹ��Է��ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "�ҹ��Է��ܶ�:" & Format(Val(GetXMLOutput("yka057")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!�ҹ��Է�, 0)
    'yka111  number  14  2   �ں˶Է�Χ�ڵ����м�¼�ķ��Ϸ�Χ�ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "���Ͻ���ܶ�:" & Format(Val(GetXMLOutput("yka111")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!���Ͻ��, 0)
    'yka112  number  14  2   �ں˶Է�Χ�ڵ����м�¼�ĸ����˻�֧���ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "�����ʻ��ܶ�:" & Format(Val(GetXMLOutput("yka112")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!�����ʻ�, 0)
    'yka113  number  14  2   �ں˶Է�Χ�ڵ����м�¼�ĸ����ֽ�֧���ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "�����ֽ��ܶ�:" & Format(Val(GetXMLOutput("yka113")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!�����ֽ�, 0)
    '���������¼��
    frmShowMsg.ShowInFor strTemp
End Function


Public Function �˶Դ�����ϸ��Ϣ_�����山(ByVal lng����id As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�˶Դ�����ϸ��Ϣ
    '--�����:
    '--������:
    '--��  ��:���غ˶Է�Χ�ڵļ�¼��
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    Dim dbl�ܶ� As Double
    
    Call Get������Ϣ(lng����id)
    Call intXML
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ15, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "16"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       �籣�����������
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    'akc190  string  20      ��Ҫ�˶Ե��˻�֧����Ϣ�ľ�����
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    
    
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    If ҵ������_�����山(�˶Դ�����ϸ��Ϣ, strXMLText, strOutPut, "") = False Then
        ShowMsgbox "�˶Դ�����ϸ��Ϣʱ,ҵ������ʧ�ܣ�"
        Exit Function
    End If
    
    If GetXML��(strOutPut) = False Then
        ShowMsgbox "�˶Դ�����ϸ��Ϣ�з��ش�����һ����Ч��XML����"
        Exit Function
    End If
    gstrSQL = "Select count(ID) as ��¼��,Sum(����*����) as ����,sum(Round(A.ʵ�ս��/(A.����*A.����),2)) as �۸�,sum(ʵ�ս��) as �����ܶ�   From ���˷��ü�¼ a  where ID in (Select ����ID From ҽ����ϸ���� where  ������='" & g�������_�����山.������ & "')"
    OpenRecordset rsTemp, "��ȡ������ϸ��Ϣ"
    
    'RecordCount number  15      �ں˶Է�Χ�ڵ�������Ϣ�ļ�¼����
    strTemp = "��ϸ��¼��:" & Val(GetXMLOutput("RecordCount")) & "|" & Nvl(rsTemp!��¼��, 0)
    'akc226  number  14  2   �ں˶Է�Χ�ڵ����м�¼�������ۼ�ֵ
    strTemp = strTemp & "||" & "��ϸ������:" & Format(Val(GetXMLOutput("akc226")), "####0.0000;-####0.00; 0;") & "|" & Nvl(rsTemp!����, 0)
    'akc225  number  14  2   �ں˶Է�Χ�ڵ����м�¼��ʵ�ʼ۸��ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "ʵ�ʼ۸��ܶ�:" & Format(Val(GetXMLOutput("akc225")), "####0.00;-####0.000;0 ;") & "|" & Nvl(rsTemp!�۸�, 0)
    'yka055  number  14  2   �ں˶Է�Χ�ڵ����м�¼��ҽ�Ʒ��ܶ��ۼ�ֵ
    strTemp = strTemp & "||" & "ҽ�Ʒ��ܶ�:" & Format(Val(GetXMLOutput("yka055")), "####0.00;-####0.000;0 ;") & "|" & Nvl(rsTemp!�����ܶ�, 0)
    
    '���������¼��
    frmShowMsg.ShowInFor strTemp
End Function


Public Function �˶Է��ý�����_�����山(ByVal lng����id As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�˶Դ�����ϸ��Ϣ
    '--�����:
    '--������:
    '--��  ��:���غ˶Է�Χ�ڵļ�¼��
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    Dim dbl�ܶ� As Double
    
    Call Get������Ϣ(lng����id)
    
    Call intXML
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ15, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "18"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       �籣�����������
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    'akc190  string  20      ��Ҫ�˶Ե��˻�֧����Ϣ�ľ�����
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g�������_�����山.������
    
    
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    If ҵ������_�����山(�˶Է��ý�����, strXMLText, strOutPut, "") = False Then
        ShowMsgbox "�˶Է��ý�����ʱ,ҵ������ʧ�ܣ�"
        Exit Function
    End If
    
    If GetXML��(strOutPut) = False Then
        ShowMsgbox "�˶Է��ý������з��ش�����һ����Ч��XML����"
        Exit Function
    End If
    
    gstrSQL = "Select count(distinct ������ ) as ��¼��,0 as �����ܶ�,sum(ȫ�Էѽ��) as ȫ�Է�,sum(�ҹ��Է�) as �ҹ��Է�,sum(���Ͻ��) as ���Ͻ��,sum(�����Ը�) as �����ʻ�," & _
    "       sum(����֧�����) as ֧�����,sum(����Աͳ��֧��) as ����Աͳ��֧��,sum(�����Ը��ۼ�) as �����Ը��ۼ�  " & _
    "   From ���ý����� where ������='" & g�������_�����山.������ & "'"
    
    OpenRecordset_ZLYB rsTemp, "��ȡ���ý��������Ϣ"
    
    'RecordCount number  15      �ں˶Է�Χ�ڵ�������Ϣ�ļ�¼����
    strTemp = "��¼��:" & Val(GetXMLOutput("RecordCount")) & "|" & Nvl(rsTemp!��¼��, 0)
    'aka213  string  6       �ֶα�׼
    strTemp = strTemp & "||" & "�ֶα�׼:" & Val(GetXMLOutput("aka213")) & "|" & "0"
    'yka056  number  14  2   �ں˶Է�Χ�ڵ����м�¼��ȫ�Էѽ���ۼ�ֵ
    strTemp = strTemp & "||" & "ȫ�Է�  �ܶ�:" & Format(Val(GetXMLOutput("yka056")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!ȫ�Է�, 0)
    'yka057  number  14  2   �ں˶Է�Χ�ڵ����м�¼�Ĺҹ��Էѽ���ۼ�ֵ
    strTemp = strTemp & "||" & "�ҹ��Է��ܶ�:" & Format(Val(GetXMLOutput("yka057")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!�ҹ��Է�, 0)
    'yka111  number  14  2   �ں˶Է�Χ�ڵ����м�¼�ķ��Ϸ�Χ����ۼ�ֵ
    strTemp = strTemp & "||" & "���Ͻ���ܶ�:" & Format(Val(GetXMLOutput("yka111")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!���Ͻ��, 0)
    
    'yka106  number  14  2   �ں˶Է�Χ�ڵ����м�¼���Ը�����ۼ�ֵ
    strTemp = strTemp & "||" & "    �Ը��ܶ�:" & Format(Val(GetXMLOutput("yka106")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!�����ʻ�, 0)
    'yka107  number  14  2   �ں˶Է�Χ�ڵ����м�¼��֧������ۼ�ֵ
    strTemp = strTemp & "||" & "    ֧�����:" & Format(Val(GetXMLOutput("yka107")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!֧�����, 0)
    'yka063  number  14  2   �ں˶Է�Χ�ڵ����м�¼�Ĺ���Աͳ��֧������ۼ�ֵ
    strTemp = strTemp & "||" & "����Աͳ��֧��:" & Format(Val(GetXMLOutput("yka063")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!����Աͳ��֧��, 0)
    'yka221  number  14  2   �ں˶Է�Χ�ڵ����м�¼������ҽ�Ʋ��������Ը��ۼƽ��
    strTemp = strTemp & "||" & "�����Ը��ۼ�:" & Format(Val(GetXMLOutput("yka063")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!�����Ը��ۼ�, 0)
    
    frmShowMsg.ShowInFor strTemp
End Function
Private Function IC���ʻ�֧��_�����山(ByVal dbl����֧�� As Double, ByVal str����ʱ�� As String, ByVal str�˵������ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����Զ��IC��֧��
    '--�����:
    '--������:
    '--��  ��:֧���ɹ�,����true,����False
    '-----------------------------------------------------------------------------------------------------------
    IC���ʻ�֧��_�����山 = False
    
    Err = 0
    On Error GoTo ErrHand:
    
    Call intXML
    
    'YAB003  String  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    If g�������_�����山.���� = True Then
        'SvrcID  String  2       Զ�����ݷ����ʶ����ֵ05, ��ʶ��Сд���У�����λ��
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "19"
    Else
        'SvrcID  String  2       Զ�����ݷ����ʶ����ֵ05, ��ʶ��Сд���У�����λ��
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "05"
    End If
    'CtrInf  String  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'code    String  20      �α���Ա��ҽ������
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g�������_�����山.����, 1, 20)
    'ykc005  String  6       ����α���Ա�����ҽ����֤���룬����λ���������ַ�
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g�������_�����山.����, 1, 6)
    'akc190  String  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", Substr(g�������_�����山.������, 1, 20)
    'aka130  String  6       ֧����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", Substr(g�������_�����山.֧�����, 1, 6)
    'akb020  String  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    'ykb006  String  3       ����ҽ�ƻ�����֧�������
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    


    If g�������_�����山.���� = True Then
        AppendXMLNode gobjXMLInPut.documentElement, "DefrayAmount", dbl����֧��
    End If
    
    'PastBaseDefray  Number  14  2   ����ҽ������֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "PastBaseDefray", 0
    'LastBaseDefray  Number  14  2   ����ҽ������֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "LastBaseDefray", 0
    'ThisBaseDefray  Number  14  2   ����ҽ�Ʊ���֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "ThisBaseDefray", dbl����֧��
    'NotPastBaseDefray   Number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻�����֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "NotPastBaseDefray", 0
    'NotLastBaseDefray   Number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻�����֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "NotLastBaseDefray", 0
    'NotThisBaseDefray   Number  14  2   ����ҽ�Ʊ��껮��Ǳ����˻�����֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "NotThisBaseDefray", 0
    'PastOfficialDefray  Number  14  2   ����Ա����֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "PastOfficialDefray", 0
    'LastOfficialDefray  Number  14  2   ����Ա����֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "LastOfficialDefray", 0
    'ThisOfficialDefray  Number  14  2   ����Ա����֧���ܶ�
    AppendXMLNode gobjXMLInPut.documentElement, "ThisOfficialDefray", 0
    'aae036  Date        ��  �˻�֧���ľ���ʱ��
    AppendXMLNode gobjXMLInPut.documentElement, "aae036", str����ʱ��
    'Yka198  String  20      �˵���Ӧ�����ţ��˴�Ϊ�����ţ�
    
    AppendXMLNode gobjXMLInPut.documentElement, "Yka198", str�˵������
    
    Dim strXMLText As String, strOutPut As String
    
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    WriteDebugInfor_�����山 strXMLText
    If ҵ������_�����山(IC���ʻ�֧��, strXMLText, strOutPut, "") = False Then
        ShowMsgbox "IC���ʻ�֧��ʱ,ҵ������ʧ�ܣ�"
        Exit Function
    End If
    IC���ʻ�֧��_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function ������¼����_�����山() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������¼����
    '--�����:
    '--������:
    '--��  ��:���ϳɹ�,����true,����False
    '-----------------------------------------------------------------------------------------------------------
    ������¼����_�����山 = False
    
    Err = 0
    On Error GoTo ErrHand:
    
    If g�������_�����山.������� Then
        ������¼����_�����山 = True
        Exit Function
    End If
        
    Call intXML
    
    'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_�����山.�����������, 1, 4)
    'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ26, ��ʶ��Сд���У�����λ��
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "26"
    'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   ���˱��
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g�������_�����山.���˱��
    
    'ykc005  string  6       ����α���Ա�����ҽ����֤���룬����λ���������ַ�
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g�������_�����山.����, 1, 6)
    'akc190  string  20      ������
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", Substr(g�������_�����山.������, 1, 20)
    'aka130  string  6       ֧����𣬼������
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", Substr(g�������_�����山.֧�����, 1, 6)
    'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_�����山.ҽԺ����, 1, 8)
    
    Dim strXMLText As String, strOutPut As String
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    
    If ҵ������_�����山(����������¼, strXMLText, strOutPut, "") = False Then
        ShowMsgbox "������¼����ʱ,ҵ������ʧ�ܣ�"
        Exit Function
    End If
    ������¼����_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Get��������_�����山(ByVal intCode As CodeType, ByVal strCode As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������,��ȡ�̶�ֵ
    '--�����:
    '--������:
    '--��  ��:����
    '-----------------------------------------------------------------------------------------------------------
    Dim strName As String
    Select Case intCode
    Case ҽ����Ա���
        strName = Switch(strCode = "41", "�¸�ְ��", strCode = "21", "����", strCode = "22", "������ذ���", strCode = "34", "������ذ���", strCode = "12", "��ְ����פ��", strCode = "11", "��ְ", strCode = "31", "����", strCode = "33", "�����Ҽ��˲о���", strCode = "32", "�Ϻ��", strCode = "51", "����ǰ�Ϲ���", True, "������Ա")
    Case ҽ�Ʋ������
        strName = Switch(strCode = "1", "����ҽ�Ʋ���", True, "������ҽ�Ʋ���")
    Case ҽ���չ����
        strName = Switch(strCode = "0", "��ҽ���չ���Ա", strCode = "1", "ҽ���չ���Ա����", True, "ҽ���չ���Ա����")
    End Select
    Get��������_�����山 = strName
End Function
Private Function Get����ʱ��(ByVal lng����id As Long) As String
    '���ܣ���ȡ����ʱ��
    '������
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ����ʱ�� From �����ʻ� where ����=" & gintInsure & " and ����id=" & lng����id
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����ʱ��"
    
    If rsTemp.RecordCount = 0 Then
        Get����ʱ�� = ""
        Exit Function
    End If
    Get����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-mm-dd")
End Function


Public Function �ҺŽ���_�����山(ByVal lng����ID As Long) As Boolean
  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    
    Dim strʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    
    �ҺŽ���_�����山 = False
    
    g�������_�����山.�����־ = 2
    g�������_�����山.���� = False
    g�������_�����山.������ = Get������
    g�������_�����山.����ID = lng����ID
    g�������_�����山.��Ʊ�� = Get��Ʊ����(lng����ID)
    g�������_�����山.������� = False

    
    
    gstrSQL = "Select ����id,�Ǽ�ʱ�� From ���˷��ü�¼ where rownum<=1 and ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�Ǽ�ʱ��"
    
    If g�������_�����山.����ʱ�� > Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") Then
        g�������_�����山.����ʱ�� = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '���浱ǰ״̬�Ľ�����
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & Nvl(rsTemp!����ID, 0) & "," & TYPE_�����山 & ",'������','''" & g�������_�����山.������ & "''')"
    Call ExecuteProcedure("���������")
        
    Err = 0
    On Error GoTo ErrHand
    gcnOracle_CQYB.BeginTrans
    
    �ҺŽ���_�����山 = ���˽���(lng����ID)
    
    If �ҺŽ���_�����山 = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
        gcnOracle_CQYB.CommitTrans
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
  gcnOracle_CQYB.RollbackTrans
End Function


Public Function �Һų���_�����山(ByVal lng����ID As Long) As Boolean

    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand
    
    �Һų���_�����山 = False
    gstrSQL = "Select ����ID From ���˷��ü�¼  where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����id"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "û�б������ĹҺ���Ŀ,���ܳ���"
        
        Exit Function
    End If
    '��ȷ����ʱ�����Ƿ��з��ý���,�����ܽ��г���
    
    
    g�������_�����山.lng����id = Nvl(rsTemp!����ID, 0)
    '��ȡ������Ϣ
    Call Get������Ϣ(g�������_�����山.lng����id)
    
    g�������_�����山.����ID = lng����ID
    g�������_�����山.�����־ = 2
    g�������_�����山.���� = False
    g�������_�����山.������ = Get������
    g�������_�����山.���� = True
    g�������_�����山.����ID = Get����ID
    g�������_�����山.������� = False

    '���浱ǰ״̬�Ľ�����
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & g�������_�����山.lng����id & "," & TYPE_�����山 & ",'������','''" & g�������_�����山.������ & "''')"
    Call ExecuteProcedure("���������")
    
    gcnOracle_CQYB.BeginTrans
    
    �Һų���_�����山 = ���˽������(lng����ID)
    If �Һų���_�����山 = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    gcnOracle_CQYB.CommitTrans
    �Һų���_�����山 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle_CQYB.RollbackTrans
End Function

Private Function ���˽������(ByVal lng����ID As Long) As Boolean
    '�������˽���
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim str����� As String
    On Error GoTo ErrHand
    
    ���˽������ = False
    
    Err = 0
    On Error GoTo ErrHand
        
    
    '��һ��:�����ǳ���,���Բ�����д����˶�,�����ȡԭ���ݵĽ�����
    gstrSQL = "select  ֧��˳���,��ע from ���ս����¼ where ����=" & gintInsure & " and ��¼id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
    
    str����� = Nvl(rsTemp!֧��˳���)
    g�������_�����山.������ = Nvl(rsTemp!��ע)
    
    '�ڶ���:д�뱾������ϸ������ı�
    
    '   ��ȡ������ϸ��¼
    gstrSQL = Get��ϸ��¼(g�������_�����山.����ID)
    
    Call OpenRecordset(rs��ϸ, "��ȡ������ϸ")
    
    If rs��ϸ.RecordCount = 0 Then
        Exit Function
    End If
    
    If Saveҽ����ϸ����(rs��ϸ) = False Then Exit Function
    
    
    '������:�������ķ��ý�������������Ϣ�ϴ�(�Ը���ʽ�ϴ�)
     If ���ý����������ϴ�(str�����) = False Then Exit Function
     
    '���Ĳ�:������ϸ�ϴ�
    If ������ϸ�ϴ�(rs��ϸ) = False Then
        ShowMsgbox "�ڽ��д�����ϸ�ϴ�ʱ����һ�����ϵ���ϸ�ϴ�ʧ��,���Ժ�ע�ⲹ��!"
    End If
    ���˽������ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ���ý���������(ByVal str������ As String) As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    Dim strOutPut As String
    Dim strXMLText As String
    Dim blnFirst As Boolean
    Dim strSql  As String
    Dim objXMLItem As MSXML2.IXMLDOMElement

    ���ý��������� = False
    
    '�������ý�����,д�෴��
    gstrSQL = "" & _
                "   Select id, ����id, ��ҳid, ������, ������, �˵������, ������¼���, ҽ����Ա���, ҽ���չ����, ҽ�Ʋ������," & _
                "       ���, ����ҽ����Ϣ, -1*���޽�� ���޽�� , -1*�Ը���� �Ը����, �����������,֧����� ֧�����,-1* ����Ա���� ����Ա����," & _
                "       -1*�����Ը���� �����Ը����, �ۼƽɷ�����, ʵ������," & _
                "       ҽ���������, -1*�ʻ�֧�� �ʻ�֧��, �ֶα�׼, -1*ȫ�Էѽ�� ȫ�Էѽ��, -1*�ҹ��Է� �ҹ��Է�, -1*���Ͻ�� ���Ͻ��, -1*�����Ը� �����Ը�," & _
                "       -1*����֧����� ����֧�����, -1*����Աͳ��֧�� ����Աͳ��֧��,-1*�����Ը��ۼ�  �����Ը��ۼ� " & _
                "   From ���ý����� " & _
                "   where ������='" & str������ & "' and ������='" & g�������_�����山.������ & "'"
                    
    Err = 0
    On Error GoTo ErrHand:
    
    OpenRecordset_ZLYB rsTemp, "����������"
    
    If rsTemp.EOF Then
        ShowMsgbox "������Ϊ:" & str������ & " �Ľ���Ų�����,����ʧ��!"
        Exit Function
    End If
 
    '�洢���̲���:
    'ID,����id, ��ҳid, ������, ������, �˵������, ������¼���, �����������, ҽ����Ա���, ҽ���չ����, ҽ�Ʋ������,
    '���,����ҽ����Ϣ, ���޽��, �Ը����, ֧�����, ����Ա����, �����Ը����, �ۼƽɷ�����, ʵ������, ҽ���������, �ʻ�֧��, �ֶα�׼,
    'ȫ�Էѽ��, �ҹ��Է�,���Ͻ��_IN, �����Ը�, ����֧�����, ����Աͳ��֧��, �����Ը��ۼ�
    
    Call intXML
    blnFirst = True
    Dim lngID As Long
    With rsTemp
        Do While Not .EOF
            gstrSQL = "Select ���ý�����_ID.nextval as ID from dual"
            OpenRecordset_ZLYB rsTemp, "��ȡ������"
            lngID = Nvl(rsTemp!ID, 0)
            
            strSql = "ZL_���ý�����_INSERT("
            
            
            strSql = strSql & lngID & ","
            strSql = strSql & Nvl(!����ID, 0) & ","
            
            strSql = strSql & Nvl(!��ҳID, 0) & ","
            strSql = strSql & "'" & Nvl(!������) & "',"
            strSql = strSql & "'" & g�������_�����山.������ & "',"
            strSql = strSql & "'" & Nvl(!������) & "',"
            strSql = strSql & "" & Nvl(!������¼���, 0) & ","
            strSql = strSql & "'" & Nvl(!�����������) & "',"
            strSql = strSql & "'" & Nvl(!ҽ����Ա���) & "',"
            strSql = strSql & "'" & Nvl(!ҽ���չ����) & "',"
            strSql = strSql & "'" & Nvl(!ҽ�Ʋ������) & "',"  'ҽ�Ʋ������
            strSql = strSql & "'" & Nvl(!���) & "',"
            strSql = strSql & "'" & Nvl(!����ҽ����Ϣ) & "',"
            strSql = strSql & "" & Nvl(!���޽��, 0) & ","
            
            
            strSql = strSql & "" & Nvl(!�Ը����, 0) & ","
            strSql = strSql & "" & Nvl(!֧�����, 0) & ","
            strSql = strSql & "" & Nvl(!����Ա����, 0) & ","
            strSql = strSql & "" & Nvl(!�����Ը����, 0) & ","
            strSql = strSql & "" & Nvl(!�ۼƽɷ�����, 0) & ","
            strSql = strSql & "" & Nvl(!ʵ������, 0) & ","
            
            strSql = strSql & "'" & Nvl(!ҽ���������) & "',"
            strSql = strSql & "" & Nvl(!�ʻ�֧��, 0) & ","
            strSql = strSql & "'" & Nvl(!�ֶα�׼) & "',"
            
            
            
            strSql = strSql & "" & Nvl(!ȫ�Էѽ��, 0) & ","
            strSql = strSql & "" & Nvl(!�ҹ��Է�, 0) & ","
            strSql = strSql & "" & Nvl(!���Ͻ��, 0) & ","
            strSql = strSql & "" & Nvl(!�����Ը�, 0) & ","
            strSql = strSql & "" & Nvl(!����֧�����, 0) & ","
            
            strSql = strSql & "" & Nvl(!����Աͳ��֧��, 0) & ","
            strSql = strSql & "" & Nvl(!�����Ը��ۼ�, 0) & ")"
            
                        
            '�������ݿ���
            gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
            
            If insertInto����(lngID, Nvl(!����ҽ����Ϣ)) = False Then
                DebugTool "�����������!"
            End If
            'XML���ý��д��
            If blnFirst Then
                AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!�����������)
                AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "12"
                AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
                'BaseInfo                ����������ܶι��е���ͬ�Ļ�����Ϣ���֣��������ŵ�Ԫ����������Ԫ��
                Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "BaseInfo", "")
                '    akc190  string  20      ������
                AppendXMLNode objXMLItem, "akc190", Nvl(!������)
                '    yka103  string  20      ������
                AppendXMLNode objXMLItem, "yka103", g�������_�����山.������
                '    yka198  string  20      �˵���Ӧ������
                AppendXMLNode objXMLItem, "yka198", Nvl(!������)
                '    ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
                AppendXMLNode objXMLItem, "ykc114", Nvl(!������¼���, 0)
                '    yab003  string  4       �籣�����������
                AppendXMLNode objXMLItem, "yab003", Nvl(!�����������)
                blnFirst = False
            End If
            
            
            '��ȷ����ص��ִ�
            'ReckonInfo              ����������ܶεĽ���ֶ���Ϣ���������ŵ�Ԫ����������Ԫ��
            Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "ReckonInfo", "")
            
            'akc190  string  20      ������
             AppendXMLNode objXMLItem, "akc190", Nvl(!������)
            'yka103  string  20      ������
             AppendXMLNode objXMLItem, "yka103", g�������_�����山.������
             
            'yka198  string  20      �˵���Ӧ������
             AppendXMLNode objXMLItem, "yka198", Nvl(!������)
            'ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
             AppendXMLNode objXMLItem, "ykc114", Nvl(!������¼���)
            'yab003  string  4       �籣�����������
             AppendXMLNode objXMLItem, "yab003", Nvl(!�����������)
            'aka213  string  2       �ֶα�׼��03 ���ߣ� 05 ����ҽ�� ��06 ���ҽ�ƣ�07 ����
             AppendXMLNode objXMLItem, "aka213", Nvl(!�ֶα�׼)
            'yka056  number  14  2   ȫ�Էѽ��
             AppendXMLNode objXMLItem, "yka056", Nvl(!ȫ�Էѽ��, 0)
            'yka057  number  14  2   �ҹ��Էѽ��
             AppendXMLNode objXMLItem, "yka057", Nvl(!�ҹ��Է�, 0)
            'yka111  number  14  2   ���Ϸ�Χ���
             AppendXMLNode objXMLItem, "yka111", Nvl(!���Ͻ��, 0)
            'yka106  number  14  2   �Ը����
             AppendXMLNode objXMLItem, "yka106", Nvl(!�����Ը�, 0)
            'yka107  number  14  2   ֧�����
             AppendXMLNode objXMLItem, "yka107", Nvl(!֧�����, 0)
            'yka063  number  14  2   ����Աͳ��֧�����
             AppendXMLNode objXMLItem, "yka063", Nvl(!����Աͳ��֧��, 0)
            'yka221  number  14  2   ����ҽ�Ʋ��������Ը��ۼƽ��
             AppendXMLNode objXMLItem, "yka221", Nvl(!�����Ը��ۼ�, 0)
            'Akc315  String  3       ҽ������ְ��
             AppendXMLNode objXMLItem, "Akc315", Nvl(!ҽ���������)
            .MoveNext
        Loop
    End With
      
    'д����ý�����
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    WriteDebugInfor_�����山 strXMLText
    
    If ҵ������_�����山(������д��, strXMLText, strOutPut) = False Then
        Exit Function
    End If
    WriteDebugInfor_�����山 strOutPut
    ���ý��������� = True
    Exit Function
    
ErrHand:
    DebugTool "���ý���������ʧ��!" & vbCrLf & " �����:" & Err.Number & vbCrLf & " ��������: " & Err.Description
 End Function
 Private Function ���û�����Ϣ����(ByVal str������ As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strXMLText As String
    Dim strSql As String
    Dim str����ʱ�� As String, strOutPut As String
    
    
    ���û�����Ϣ���� = False
    'д����ý��������Ϣ
    '���û�����Ϣ����
    gstrSQL = " " & _
    "           Select ����id, ��ҳid, ������, ������, �˵������, ������¼���, ���˱��, ��λ���, ����, �Ա�, ��������, ʵ������, " & _
    "                   �ۼƽɷ�����, ҽ����Ա���, ҽ�ƻ�������, ��֧��������, ҽ�ƻ������, ���ֲ���־, ֧�����, ���ֱ���, ��������, " & _
    "                   -1*ҽ�Ʒ��ܶ� ҽ�Ʒ��ܶ�, -1*ȫ�Է��ܶ� ȫ�Է��ܶ�, -1*�ҹ��Է��ܶ� �ҹ��Է��ܶ�, -1*���Ϸ�Χ�ܶ� ���Ϸ�Χ�ܶ�, -1*�����ʻ�֧���ܶ� �����ʻ�֧���ܶ�," & _
    "                  -1*�����ֽ�֧���ܶ� �����ֽ�֧���ܶ�, ����ʱ��, �����������, " & _
    "                   ҽ���չ���� , ҽ�Ʋ������, ������㷽ʽ, ��Ʊ��, ��ע, �ֶμ������, ҽ��������� " & _
    "           From ���û�����Ϣ " & _
    "           where ������='" & g�������_�����山.������ & "' and ������='" & str������ & "'"

    '���̲���:
    '    ����id, ��ҳid, ������, ������, �˵������, ������¼���, ���˱��, ��λ���, ����, �Ա�, ��������, ʵ������,
    '    �ۼƽɷ�����, ҽ����Ա���, ҽ�ƻ�������, ��֧��������, ҽ�ƻ������, ���ֲ���־, ֧�����, ���ֱ���, ��������,
    '    ҽ�Ʒ��ܶ�, ȫ�Է��ܶ�, �ҹ��Է��ܶ�, ���Ϸ�Χ�ܶ�, �����ʻ�֧���ܶ�, �����ֽ�֧���ܶ�, ����ʱ��, �����������,
    '    ҽ���չ���� , ҽ�Ʋ������, ������㷽ʽ, ��Ʊ��, ��ע, �ֶμ������, ҽ���������
    
    OpenRecordset_ZLYB rsTemp, "��ȡ���ý��������Ϣ"
    
    With rsTemp
    
        strSql = "ZL_���û�����Ϣ_INSERT(" & Nvl(!����ID, 0) & ","
        strSql = strSql & Nvl(!��ҳID, 0) & ","
        strSql = strSql & "'" & Nvl(!������) & "',"
        strSql = strSql & "'" & g�������_�����山.������ & "',"
        strSql = strSql & "'" & Nvl(!������) & "',"
        strSql = strSql & "" & Nvl(!������¼���, 0) & ","
        strSql = strSql & "" & Nvl(!���˱��, 0) & ","
        strSql = strSql & "" & Nvl(!��λ���, 0) & ","
        strSql = strSql & "'" & Nvl(!����) & "',"
        strSql = strSql & "'" & Nvl(!�Ա�) & "',"
        If IsNull(!��������) Then
            strSql = strSql & "NULL,"
        Else
            strSql = strSql & "to_date('" & Format(!��������, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
        End If
        
        strSql = strSql & "" & Nvl(!ʵ������, 0) & ","
        strSql = strSql & "" & Nvl(!�ۼƽɷ�����, 0) & ","
        strSql = strSql & "'" & Nvl(!ҽ����Ա���) & "',"  'ҽ����Ա���
        strSql = strSql & "'" & Nvl(!ҽ�ƻ�������) & "',"
        strSql = strSql & "'" & Nvl(!��֧��������) & "',"
        strSql = strSql & "'" & Nvl(!ҽ�ƻ������) & "',"
        strSql = strSql & "'" & Nvl(!���ֲ���־) & "',"
        strSql = strSql & "'" & Nvl(!֧�����) & "',"
        strSql = strSql & "'" & Nvl(!���ֱ���) & "',"
        strSql = strSql & "" & Nvl(!��������, 0) & ","
        strSql = strSql & "" & Nvl(!ҽ�Ʒ��ܶ�, 0) & ","

        
        strSql = strSql & "" & Nvl(!ȫ�Է��ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!�ҹ��Է��ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!���Ϸ�Χ�ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!�����ʻ�֧���ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!�����ֽ�֧���ܶ�, 0) & ","
        If IsNull(!����ʱ��) Then
            strSql = strSql & "NULL,"
        Else
            strSql = strSql & "to_date('" & Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        End If
        
        strSql = strSql & "'" & Nvl(!�����������) & "',"
        strSql = strSql & "'" & Nvl(!ҽ���չ����) & "',"
        strSql = strSql & "'" & Nvl(!ҽ�Ʋ������) & "',"
        strSql = strSql & "'" & Nvl(!������㷽ʽ) & "',"
        strSql = strSql & "'" & Nvl(!��Ʊ��) & "',"
        strSql = strSql & "'" & Nvl(!��ע) & "',"
        strSql = strSql & "'" & Nvl(!�ֶμ������) & "',"
        strSql = strSql & "'" & Nvl(!ҽ���������) & "')"
            
        '��������
        gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
        
        Call intXML
    
        'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
        AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!�����������)
        'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ10, ��ʶ��Сд���У�����λ��
        
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "10"
        'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
        AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
        'akc190  string  20      ������
        AppendXMLNode gobjXMLInPut.documentElement, "akc190", Nvl(!������)
        'yka103  string  20      ������
        AppendXMLNode gobjXMLInPut.documentElement, "yka103", g�������_�����山.������
        'yka198  string  20      �˵���Ӧ������
        AppendXMLNode gobjXMLInPut.documentElement, "yka198", Nvl(!������)
        
        'ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
        AppendXMLNode gobjXMLInPut.documentElement, "ykc114", Nvl(!������¼���, 0)
        'aac001  number  15  0   ���˱��
        AppendXMLNode gobjXMLInPut.documentElement, "aac001", Nvl(!���˱��, 0)
        'aab001  number  15  0   ��λ���
        AppendXMLNode gobjXMLInPut.documentElement, "aab001", Nvl(!��λ���, 0)
        'aac003  string  20      ����
        AppendXMLNode gobjXMLInPut.documentElement, "aac003", Nvl(!����)
        'aac004  string  1       �Ա𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "aac004", Nvl(!�Ա�)
        
        'aac006  date    ��      ��������
        AppendXMLNode gobjXMLInPut.documentElement, "aac006", Format(!��������, "yyyy-mm-dd")
        'akc023  number  3       ʵ������
        AppendXMLNode gobjXMLInPut.documentElement, "akc023", Nvl(!ʵ������, 0)
        'ykc021  number  3       �ۼƽɷ�����
        AppendXMLNode gobjXMLInPut.documentElement, "ykc021", Nvl(!�ۼƽɷ�����, 0)
        'akc021  string  6       ҽ����Ա��𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "akc021", Nvl(!ҽ����Ա���)
        'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
        AppendXMLNode gobjXMLInPut.documentElement, "akb020", Nvl(!ҽ�ƻ�������)
        'ykb006  string  3       ����ҽ�ƻ�����֧�������
        AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"          '��֧��������
        'akb023  string  6       ҽ�ƻ�����𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "akb023", InitInfor_�����山.�������
        
        'aka123  string  1       ���ֲ���־���������
        AppendXMLNode gobjXMLInPut.documentElement, "aka123", Nvl(!���ֲ���־, 0)      '���ֲ���־
        'aka130  string  6       ֧����𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "aka130", Nvl(!֧�����)
        'yka026  string  20      ���ֱ���
        AppendXMLNode gobjXMLInPut.documentElement, "yka026", Nvl(!���ֱ���)
        
        '    '    ����id, ��ҳid, ������, ������, �˵������, ������¼���, ���˱��, ��λ���, ����, �Ա�, ��������, ʵ������,
        '    �ۼƽɷ�����, ҽ����Ա���, ҽ�ƻ�������, ��֧��������, ҽ�ƻ������, ���ֲ���־, ֧�����, ���ֱ���, ��������,
        '    ҽ�Ʒ��ܶ�, ȫ�Է��ܶ�, �ҹ��Է��ܶ�, ���Ϸ�Χ�ܶ�, �����ʻ�֧���ܶ�, �����ֽ�֧���ܶ�, ����ʱ��, �����������,
        '    ҽ���չ���� , ҽ�Ʋ������, ������㷽ʽ, ��Ʊ��, ��ע, �ֶμ������, ҽ���������
        
        'yka115  number  14  2   ��������
        AppendXMLNode gobjXMLInPut.documentElement, "yka115", Nvl(!��������, 0)           '��������
        'yka055  number  14  2   ҽ�Ʒ��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", Nvl(!ҽ�Ʒ��ܶ�, 0)
        'yka056  number  14  2   ȫ�Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", Nvl(!ȫ�Է��ܶ�, 0)              '
        'yka057  number  14  2   �ҹ��Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", Nvl(!�ҹ��Է��ܶ�, 0)               '
        'yka111  number  14  2   ���Ϸ�Χ�ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", Nvl(!���Ϸ�Χ�ܶ�, 0)                '
        'yka112  number  14  2   �����˻�֧���ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka112", Nvl(!�����ʻ�֧���ܶ�, 0)                 '
        'yka113  number  14  2   �����ֽ�֧���ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka113", Nvl(!�����ֽ�֧���ܶ�, 0)                  '
        'aae036  date        ��  ����ʱ��
        str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        '����ʱ��
        AppendXMLNode gobjXMLInPut.documentElement, "aae036", str����ʱ��                 '
        'yab003  string  4       �籣�����������
        AppendXMLNode gobjXMLInPut.documentElement, "yab003", Nvl(!�����������)               '
        'ykc120  string  6       ҽ���չ���𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "ykc120", Nvl(!ҽ���չ����)                  '
        'ykc121  string  6       ����ҽ�Ʋ�����𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "ykc121", Nvl(!ҽ�Ʋ������)
        'yka222  string  6       ������㷽ʽ
        AppendXMLNode gobjXMLInPut.documentElement, "yka222", Nvl(!������㷽ʽ) '
        'yka110  string  20      ��Ʊ��
        AppendXMLNode gobjXMLInPut.documentElement, "yka110", Nvl(!��Ʊ��)                                '
        'aae013  string  100     ��ע
        AppendXMLNode gobjXMLInPut.documentElement, "aae013", Nvl(!��ע)                              '
        'gkc010  string  800     �ֶμ������(סԺ��)
        AppendXMLNode gobjXMLInPut.documentElement, "gkc010", Nvl(!�ֶμ������)                              '
        'akc315  string  3       ҽ�ƴ���������𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "akc315", Nvl(!ҽ���������)                              '
    End With
    
    'д�������Ϣ
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    WriteDebugInfor_�����山 strXMLText
    
    If ҵ������_�����山(���������Ϣд��, strXMLText, strOutPut) = False Then
        Exit Function
    End If
    ���û�����Ϣ���� = True
    Exit Function
ErrHand:
    DebugTool "���û�����Ϣ����ʧ��!" & vbCrLf & " �����:" & Err.Number & vbCrLf & " ��������: " & Err.Description
End Function
Private Function ���ý����������ϴ�(ByVal str������ As String) As Boolean
    '���ݽ�����,�������εĽ�����Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim blnFirst As Boolean
    Dim objXMLItem As MSXML2.IXMLDOMElement
    Dim strXMLText As String
    Dim strXMLtext1 As String
    Dim strOutPut As String
    Dim str����ʱ�� As String
    Dim rs���� As New ADODB.Recordset
    If g�������_�����山.�����־ = 0 Or g�������_�����山.�����־ = 2 Then
        gstrSQL = "Select * From ���ս����¼ where ����=1 and  ��¼id=" & g�������_�����山.����ID
    Else
        gstrSQL = "Select * From ���ս����¼ where ����=2 and  ��¼id=" & g�������_�����山.����ID
    End If
     
     '��ȡԭ�����¼�е�����,�Ա����
     zlDatabase.OpenRecordset rs����, gstrSQL, "��ȡԭ�����¼�е�����"
     If rs����.RecordCount = 0 Then
        ShowMsgbox "��ȡԭ�����¼ʱ����������ʷ�����¼,���ܱ�����!"
        Exit Function
     End If
    
    gstrSQL = "" & _
            "   Select id,����id, ��ҳid, ������, ������, �˵������, ������¼���, ҽ����Ա���, ҽ���չ����, ҽ�Ʋ������," & _
            "       ���, ����ҽ����Ϣ, -1*���޽�� ���޽�� , -1*�Ը���� �Ը����, �����������,֧����� ֧�����,-1* ����Ա���� ����Ա����," & _
            "       -1*�����Ը���� �����Ը����, �ۼƽɷ�����, ʵ������," & _
            "       ҽ���������, -1*�ʻ�֧�� �ʻ�֧��, �ֶα�׼, -1*ȫ�Էѽ�� ȫ�Էѽ��, -1*�ҹ��Է� �ҹ��Է�, -1*���Ͻ�� ���Ͻ��, -1*�����Ը� �����Ը�," & _
            "       -1*����֧����� ����֧�����, -1*����Աͳ��֧�� ����Աͳ��֧��,-1*�����Ը��ۼ�  �����Ը��ۼ� " & _
            "   From ���ý����� " & _
            "   where ������='" & str������ & "' and ������='" & g�������_�����山.������ & "'"
                
    Err = 0
    On Error GoTo ErrHand:
    OpenRecordset_ZLYB rsTemp, "����������"
    
    ���ý����������ϴ� = False
    If rsTemp.EOF Then
        ShowMsgbox "������Ϊ:" & str������ & " �Ľ���Ų�����,����ʧ��!"
        Exit Function
    End If
 
    '�洢���̲���:
    '����id, ��ҳid, ������, ������, �˵������, ������¼���, �����������, ҽ����Ա���, ҽ���չ����, ҽ�Ʋ������,
    '���,����ҽ����Ϣ, ���޽��, �Ը����, ֧�����, ����Ա����, �����Ը����, �ۼƽɷ�����, ʵ������, ҽ���������, �ʻ�֧��, �ֶα�׼,
    'ȫ�Էѽ��, �ҹ��Է�,���Ͻ��_IN, �����Ը�, ����֧�����, ����Աͳ��֧��, �����Ը��ۼ�
    
    Call intXML
    Dim lngID As Long
    blnFirst = True
    With rsTemp
        Do While Not .EOF
            strSql = "ZL_���ý�����_INSERT("
            
            gstrSQL = "Select ���ý�����_ID.nextval as ID from dual"
            OpenRecordset_ZLYB rsTmp, "��ȡ������"
            lngID = Nvl(rsTmp!ID, 0)
                
            strSql = strSql & lngID & ","
            strSql = strSql & Nvl(!����ID, 0) & ","
            strSql = strSql & Nvl(!��ҳID, 0) & ","
            strSql = strSql & "'" & Nvl(!������) & "',"
            strSql = strSql & "'" & g�������_�����山.������ & "',"
            strSql = strSql & "'" & Nvl(!������) & "',"
            strSql = strSql & "" & Nvl(!������¼���, 0) & ","
            strSql = strSql & "'" & Nvl(!�����������) & "',"
            strSql = strSql & "'" & Nvl(!ҽ����Ա���) & "',"
            strSql = strSql & "'" & Nvl(!ҽ���չ����) & "',"
            strSql = strSql & "'" & Nvl(!ҽ�Ʋ������) & "',"  'ҽ�Ʋ������
            strSql = strSql & "'" & Nvl(!���) & "',"
            strSql = strSql & "'" & Nvl(!����ҽ����Ϣ) & "',"
            strSql = strSql & "" & Nvl(!���޽��, 0) & ","
            
            
            strSql = strSql & "" & Nvl(!�Ը����, 0) & ","
            strSql = strSql & "" & Nvl(!֧�����, 0) & ","
            strSql = strSql & "" & Nvl(!����Ա����, 0) & ","
            strSql = strSql & "" & Nvl(!�����Ը����, 0) & ","
            strSql = strSql & "" & Nvl(!�ۼƽɷ�����, 0) & ","
            strSql = strSql & "" & Nvl(!ʵ������, 0) & ","
            
            strSql = strSql & "'" & Nvl(!ҽ���������) & "',"
            strSql = strSql & "" & Nvl(!�ʻ�֧��, 0) & ","
            strSql = strSql & "'" & Nvl(!�ֶα�׼) & "',"
            
            
            
            strSql = strSql & "" & Nvl(!ȫ�Էѽ��, 0) & ","
            strSql = strSql & "" & Nvl(!�ҹ��Է�, 0) & ","
            strSql = strSql & "" & Nvl(!���Ͻ��, 0) & ","
            strSql = strSql & "" & Nvl(!�����Ը�, 0) & ","
            strSql = strSql & "" & Nvl(!����֧�����, 0) & ","
            
            strSql = strSql & "" & Nvl(!����Աͳ��֧��, 0) & ","
            strSql = strSql & "" & Nvl(!�����Ը��ۼ�, 0) & ")"
            
                    
            '�������ݿ���
            gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
            
            If insertInto����(lngID, Nvl(!����ҽ����Ϣ)) = False Then
                DebugTool "�����������!"
            End If
            
            'XML���ý��д��
            If blnFirst Then
                AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!�����������)
                AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "12"
                AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
                'BaseInfo                ����������ܶι��е���ͬ�Ļ�����Ϣ���֣��������ŵ�Ԫ����������Ԫ��
                Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "BaseInfo", "")
                '    akc190  string  20      ������
                AppendXMLNode objXMLItem, "akc190", Nvl(!������)
                '    yka103  string  20      ������
                AppendXMLNode objXMLItem, "yka103", g�������_�����山.������
                '    yka198  string  20      �˵���Ӧ������
                AppendXMLNode objXMLItem, "yka198", Nvl(!������)
                '    ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
                AppendXMLNode objXMLItem, "ykc114", Nvl(!������¼���, 0)
                '    yab003  string  4       �籣�����������
                AppendXMLNode objXMLItem, "yab003", Nvl(!�����������)
                blnFirst = False
            End If
            
            '��ȷ����ص��ִ�
            'ReckonInfo              ����������ܶεĽ���ֶ���Ϣ���������ŵ�Ԫ����������Ԫ��
            Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "ReckonInfo", "")
            
            'akc190  string  20      ������
             AppendXMLNode objXMLItem, "akc190", Nvl(!������)
            'yka103  string  20      ������
             AppendXMLNode objXMLItem, "yka103", g�������_�����山.������
             
            'yka198  string  20      �˵���Ӧ������
             AppendXMLNode objXMLItem, "yka198", Nvl(!������)
            'ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
             AppendXMLNode objXMLItem, "ykc114", Nvl(!������¼���)
            'yab003  string  4       �籣�����������
             AppendXMLNode objXMLItem, "yab003", Nvl(!�����������)
            'aka213  string  2       �ֶα�׼��03 ���ߣ� 05 ����ҽ�� ��06 ���ҽ�ƣ�07 ����
             AppendXMLNode objXMLItem, "aka213", Nvl(!�ֶα�׼)
            'yka056  number  14  2   ȫ�Էѽ��
             AppendXMLNode objXMLItem, "yka056", Nvl(!ȫ�Էѽ��, 0)
            'yka057  number  14  2   �ҹ��Էѽ��
             AppendXMLNode objXMLItem, "yka057", Nvl(!�ҹ��Է�, 0)
            'yka111  number  14  2   ���Ϸ�Χ���
             AppendXMLNode objXMLItem, "yka111", Nvl(!���Ͻ��, 0)
            'yka106  number  14  2   �Ը����
             AppendXMLNode objXMLItem, "yka106", Nvl(!�����Ը�, 0)
            'yka107  number  14  2   ֧�����
             AppendXMLNode objXMLItem, "yka107", Nvl(!����֧�����, 0)
            'yka063  number  14  2   ����Աͳ��֧�����
             AppendXMLNode objXMLItem, "yka063", Nvl(!����Աͳ��֧��, 0)
            'yka221  number  14  2   ����ҽ�Ʋ��������Ը��ۼƽ��
             AppendXMLNode objXMLItem, "yka221", Nvl(!�����Ը��ۼ�, 0)
            'Akc315  String  3       ҽ������ְ��
             AppendXMLNode objXMLItem, "Akc315", Nvl(!ҽ���������)
            .MoveNext
        Loop
    End With
      
    

    'д����ý�����
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    strXMLtext1 = strXMLText
      
    'д����ý��������Ϣ
    '���û�����Ϣ����
    gstrSQL = " " & _
    "           Select ����id, ��ҳid, ������, ������, �˵������, ������¼���, ���˱��, ��λ���, ����, �Ա�, ��������, ʵ������, " & _
    "                   �ۼƽɷ�����, ҽ����Ա���, ҽ�ƻ�������, ��֧��������, ҽ�ƻ������, ���ֲ���־, ֧�����, ���ֱ���, ��������, " & _
    "                   -1*ҽ�Ʒ��ܶ� ҽ�Ʒ��ܶ�, -1*ȫ�Է��ܶ� ȫ�Է��ܶ�, -1*�ҹ��Է��ܶ� �ҹ��Է��ܶ�, -1*���Ϸ�Χ�ܶ� ���Ϸ�Χ�ܶ�, -1*�����ʻ�֧���ܶ� �����ʻ�֧���ܶ�," & _
    "                  -1*�����ֽ�֧���ܶ� �����ֽ�֧���ܶ�, ����ʱ��, �����������, " & _
    "                   ҽ���չ���� , ҽ�Ʋ������, ������㷽ʽ, ��Ʊ��, ��ע, �ֶμ������, ҽ��������� " & _
    "           From ���û�����Ϣ " & _
    "           where ������='" & g�������_�����山.������ & "' and ������='" & str������ & "'"

    '���̲���:
    '    ����id, ��ҳid, ������, ������, �˵������, ������¼���, ���˱��, ��λ���, ����, �Ա�, ��������, ʵ������,
    '    �ۼƽɷ�����, ҽ����Ա���, ҽ�ƻ�������, ��֧��������, ҽ�ƻ������, ���ֲ���־, ֧�����, ���ֱ���, ��������,
    '    ҽ�Ʒ��ܶ�, ȫ�Է��ܶ�, �ҹ��Է��ܶ�, ���Ϸ�Χ�ܶ�, �����ʻ�֧���ܶ�, �����ֽ�֧���ܶ�, ����ʱ��, �����������,
    '    ҽ���չ���� , ҽ�Ʋ������, ������㷽ʽ, ��Ʊ��, ��ע, �ֶμ������, ҽ���������
    OpenRecordset_ZLYB rsTemp, "��ȡ���ý��������Ϣ"
    With rsTemp
    
        strSql = "ZL_���û�����Ϣ_INSERT(" & Nvl(!����ID, 0) & ","
        strSql = strSql & Nvl(!��ҳID, 0) & ","
        strSql = strSql & "'" & Nvl(!������) & "',"
        strSql = strSql & "'" & g�������_�����山.������ & "',"
        strSql = strSql & "'" & Nvl(!������) & "',"
        strSql = strSql & "" & Nvl(!������¼���, 0) & ","
        strSql = strSql & "" & Nvl(!���˱��, 0) & ","
        strSql = strSql & "" & Nvl(!��λ���, 0) & ","
        strSql = strSql & "'" & Nvl(!����) & "',"
        strSql = strSql & "'" & Nvl(!�Ա�) & "',"
        If IsNull(!��������) Then
            strSql = strSql & "NULL,"
        Else
            strSql = strSql & "to_date('" & Format(!��������, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
        End If
        
        strSql = strSql & "" & Nvl(!ʵ������, 0) & ","
        strSql = strSql & "" & Nvl(!�ۼƽɷ�����, 0) & ","
        strSql = strSql & "'" & Nvl(!ҽ����Ա���) & "',"  'ҽ����Ա���
        strSql = strSql & "'" & Nvl(!ҽ�ƻ�������) & "',"
        strSql = strSql & "'" & Nvl(!��֧��������) & "',"
        strSql = strSql & "'" & Nvl(!ҽ�ƻ������) & "',"
        strSql = strSql & "'" & Nvl(!���ֲ���־) & "',"
        strSql = strSql & "'" & Nvl(!֧�����) & "',"
        strSql = strSql & "'" & Nvl(!���ֱ���) & "',"
        strSql = strSql & "" & Nvl(!��������, 0) & ","
        strSql = strSql & "" & Nvl(!ҽ�Ʒ��ܶ�, 0) & ","

        
        strSql = strSql & "" & Nvl(!ȫ�Է��ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!�ҹ��Է��ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!���Ϸ�Χ�ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!�����ʻ�֧���ܶ�, 0) & ","
        strSql = strSql & "" & Nvl(!�����ֽ�֧���ܶ�, 0) & ","
        If IsNull(!����ʱ��) Then
            strSql = strSql & "NULL,"
        Else
            strSql = strSql & "to_date('" & Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        End If
        
        strSql = strSql & "'" & Nvl(!�����������) & "',"
        strSql = strSql & "'" & Nvl(!ҽ���չ����) & "',"
        strSql = strSql & "'" & Nvl(!ҽ�Ʋ������) & "',"
        strSql = strSql & "'" & Nvl(!������㷽ʽ) & "',"
        strSql = strSql & "'" & Nvl(!��Ʊ��) & "',"
        strSql = strSql & "'" & Nvl(!��ע) & "',"
        strSql = strSql & "'" & Nvl(!�ֶμ������) & "',"
        strSql = strSql & "'" & Nvl(!ҽ���������) & "')"
            
        '��������
        gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
        
        Call intXML
        
    
        'YAB003  string  4       �ڶ���ҽ�ƻ�������Ĳα���Ա���ڵ��籣����������룬����λ��
        AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!�����������)
        'SvrcID  string  2       Զ�����ݷ����ʶ����ֵ10, ��ʶ��Сд���У�����λ��
        
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "10"
        'CtrInf  string  20      ������Ϣ��Ԥ��, ��ʶ��Сд����
        AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
        'akc190  string  20      ������
        AppendXMLNode gobjXMLInPut.documentElement, "akc190", Nvl(!������)
        'yka103  string  20      ������
        AppendXMLNode gobjXMLInPut.documentElement, "yka103", g�������_�����山.������
        'yka198  string  20      �˵���Ӧ������
        AppendXMLNode gobjXMLInPut.documentElement, "yka198", Nvl(!������)
        
        'ykc114  number  15  0   ������¼��ţ���ʾ��ͬһ��������µĶ���������Ϣ
        AppendXMLNode gobjXMLInPut.documentElement, "ykc114", Nvl(!������¼���, 0)
        'aac001  number  15  0   ���˱��
        AppendXMLNode gobjXMLInPut.documentElement, "aac001", Nvl(!���˱��, 0)
        'aab001  number  15  0   ��λ���
        AppendXMLNode gobjXMLInPut.documentElement, "aab001", Nvl(!��λ���, 0)
        'aac003  string  20      ����
        AppendXMLNode gobjXMLInPut.documentElement, "aac003", Nvl(!����)
        'aac004  string  1       �Ա𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "aac004", Nvl(!�Ա�)
        
        'aac006  date    ��      ��������
        AppendXMLNode gobjXMLInPut.documentElement, "aac006", Format(!��������, "yyyy-mm-dd")
        'akc023  number  3       ʵ������
        AppendXMLNode gobjXMLInPut.documentElement, "akc023", Nvl(!ʵ������, 0)
        'ykc021  number  3       �ۼƽɷ�����
        AppendXMLNode gobjXMLInPut.documentElement, "ykc021", Nvl(!�ۼƽɷ�����, 0)
        'akc021  string  6       ҽ����Ա��𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "akc021", Nvl(!ҽ����Ա���)
        'akb020  string  8       ����ҽ�ƻ����ھ���α���Ա���ڵ�ҽ�������еı��
        AppendXMLNode gobjXMLInPut.documentElement, "akb020", Nvl(!ҽ�ƻ�������)
        'ykb006  string  3       ����ҽ�ƻ�����֧�������
        AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"          '��֧��������
        'akb023  string  6       ҽ�ƻ�����𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "akb023", InitInfor_�����山.�������
        
        'aka123  string  1       ���ֲ���־���������
        AppendXMLNode gobjXMLInPut.documentElement, "aka123", Nvl(!���ֲ���־, 0)      '���ֲ���־
        'aka130  string  6       ֧����𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "aka130", Nvl(!֧�����)
        'yka026  string  20      ���ֱ���
        AppendXMLNode gobjXMLInPut.documentElement, "yka026", Nvl(!���ֱ���)
        
        '    '    ����id, ��ҳid, ������, ������, �˵������, ������¼���, ���˱��, ��λ���, ����, �Ա�, ��������, ʵ������,
        '    �ۼƽɷ�����, ҽ����Ա���, ҽ�ƻ�������, ��֧��������, ҽ�ƻ������, ���ֲ���־, ֧�����, ���ֱ���, ��������,
        '    ҽ�Ʒ��ܶ�, ȫ�Է��ܶ�, �ҹ��Է��ܶ�, ���Ϸ�Χ�ܶ�, �����ʻ�֧���ܶ�, �����ֽ�֧���ܶ�, ����ʱ��, �����������,
        '    ҽ���չ���� , ҽ�Ʋ������, ������㷽ʽ, ��Ʊ��, ��ע, �ֶμ������, ҽ���������
        
        'yka115  number  14  2   ��������
        AppendXMLNode gobjXMLInPut.documentElement, "yka115", Nvl(!��������, 0)           '��������
        'yka055  number  14  2   ҽ�Ʒ��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", Nvl(!ҽ�Ʒ��ܶ�, 0)
        'yka056  number  14  2   ȫ�Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", Nvl(!ȫ�Է��ܶ�, 0)              '
        'yka057  number  14  2   �ҹ��Է��ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", Nvl(!�ҹ��Է��ܶ�, 0)               '
        'yka111  number  14  2   ���Ϸ�Χ�ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", Nvl(!���Ϸ�Χ�ܶ�, 0)                '
        'yka112  number  14  2   �����˻�֧���ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka112", Nvl(!�����ʻ�֧���ܶ�, 0)                 '
        'yka113  number  14  2   �����ֽ�֧���ܶ�
        AppendXMLNode gobjXMLInPut.documentElement, "yka113", Nvl(!�����ֽ�֧���ܶ�, 0)                  '
        'aae036  date        ��  ����ʱ��
        str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        '����ʱ��
        AppendXMLNode gobjXMLInPut.documentElement, "aae036", str����ʱ��                 '
        'yab003  string  4       �籣�����������
        AppendXMLNode gobjXMLInPut.documentElement, "yab003", Nvl(!�����������)               '
        'ykc120  string  6       ҽ���չ���𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "ykc120", Nvl(!ҽ���չ����)                  '
        'ykc121  string  6       ����ҽ�Ʋ�����𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "ykc121", Nvl(!ҽ�Ʋ������)
        'yka222  string  6       ������㷽ʽ
        AppendXMLNode gobjXMLInPut.documentElement, "yka222", Nvl(!������㷽ʽ) '
        'yka110  string  20      ��Ʊ��
        AppendXMLNode gobjXMLInPut.documentElement, "yka110", Nvl(!��Ʊ��)                                '
        'aae013  string  100     ��ע
        AppendXMLNode gobjXMLInPut.documentElement, "aae013", Nvl(!��ע)                              '
        'gkc010  string  800     �ֶμ������(סԺ��)
        AppendXMLNode gobjXMLInPut.documentElement, "gkc010", Nvl(!�ֶμ������)                              '
        'akc315  string  3       ҽ�ƴ���������𣬼������
        AppendXMLNode gobjXMLInPut.documentElement, "akc315", Nvl(!ҽ���������)                              '
            
    End With
    'д�������Ϣ
    strXMLText = ȡ��XML��ǰ����ʶ(gobjXMLInPut.xml)
    WriteDebugInfor_�����山 strXMLText
    
    If ҵ������_�����山(���������Ϣд��, strXMLText, strOutPut) = False Then
        Exit Function
    End If
    
    WriteDebugInfor_�����山 strXMLtext1
    
    '������ý�����
    If ҵ������_�����山(������д��, strXMLtext1, strOutPut) = False Then
        Exit Function
    End If
    
    '���˸����ʻ���
    If IC���ʻ�֧��_�����山(rsTemp!�����ʻ�֧���ܶ�, str����ʱ��, Nvl(rsTemp!������)) = False Then
            Exit Function
    End If
   
   'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
  '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(����Ա����),�ʻ��ۼ�֧��_IN(���֧��),�ۼƽ���ͳ��_IN(����ҽ���Ը�),�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����(��������),�ⶥ��_IN(֧�����+10000),ʵ������_IN,
    '   �������ý��_IN(��������),ȫ�Ը����_IN(ȫ�Ը�),�����Ը����_IN(�����Ը�),
    '   ����ͳ����_IN(���Ͻ��),ͳ�ﱨ�����_IN(����ҽ��ͳ��֧��),���Ը����_IN(����Ը�),�����Ը����_IN(�����Ը�),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(������),��ҳID_IN,��;����_IN,��ע_IN(������)
    
    gstrSQL = "zl_���ս����¼_insert(" & IIf(g�������_�����山.�����־ = 1, 2, 1) & "," & g�������_�����山.����ID & "," & gintInsure & "," & g�������_�����山.lng����id & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      -1 * Nvl(rs����!�ʻ��ۼ�����, 0) & "," & -1 * Nvl(rs����!�ʻ��ۼ�֧��, 0) & "," & -1 * Nvl(rs����!�ۼƽ���ͳ��, 0) & ",NULL,NULL," & -1 * Nvl(rs����!����, 0) & "," & Nvl(rs����!�ⶥ��, 0) & ",NULL," & _
       -1 * Nvl(rs����!�������ý��, 0) & "," & -1 * Nvl(rs����!ȫ�Ը����, 0) & "," & -1 * Nvl(rs����!�����Ը����, 0) & "," & _
        "" & -1 * Nvl(rs����!����ͳ����, 0) & "," & -1 * Nvl(rs����!ͳ�ﱨ�����, 0) & "," & -1 * Nvl(rs����!���Ը����, 0) & "," & -1 * Nvl(rs����!�����Ը����, 0) & "," & -1 * Nvl(rs����!�����ʻ�֧��, 0) & ",'" & _
       g�������_�����山.������ & "'," & IIf(Nvl(rs����!��ҳID, 0) = 0, "NULL", Nvl(rs����!��ҳID, 0)) & "," & IIf(g�������_�����山.��;���� = 1, "1", "NULL") & ",'" & _
       g�������_�����山.������ & "')"
       
    Call ExecuteProcedure("���汣�ս����¼")
    ���ý����������ϴ� = True
    Exit Function
ErrHand:
    
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function ���˽���(ByVal lng����ID As Long) As Boolean
    '���˷��ý���
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim strCurrDate As String
    Dim str��ʼʱ�� As String
    Dim lng����id  As Long
    Err = 0
    
    On Error GoTo ErrHand:
    
    '��һ��:��ȷ���ʸ����������˶�
    
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:SS")
    
    If InitInfor_�����山.ģ������ Then
        str��ʼʱ�� = "2004-07-10 21:40:29"
        strCurrDate = "2004-07-10 21:40:29"
    Else
        str��ʼʱ�� = g�������_�����山.����ʱ��
        If g�������_�����山.�����־ = 1 Then
            'סԺ�Ļ�,�俪ʼʱ��Ӧ�ô�00:00:00�뿪ʼ��.
            str��ʼʱ�� = Format(str��ʼʱ��, "yyyy-mm-dd" & " 00:00:00")
        End If
    End If
    lng����id = g�������_�����山.lng����id
    
    If g�������_�����山.������� Then
          '��������¼����
        Call ������¼����_�����山
    End If
    
    WriteDebugDate_�����山 "��ȡ�ʸ�������Ϣ��ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g�������_�����山.�����־ = 1 And g�������_�����山.������� = False Then
        'סԺ���㲻�ܽ����ʸ���˺�д,��������㼴�����.
    Else
        If �ʸ���˴����˶�(lng����id, str��ʼʱ��, strCurrDate) = False Then
            '�п��ܸ�������¼�Ѿ�����,��������һ��,�ٽ��к˶�.
            Call ������¼����_�����山
            If �ʸ���˴����˶�(lng����id, str��ʼʱ��, strCurrDate) = False Then
                Exit Function
            End If
        End If
        
        '�ڶ���:д���ʸ���������,���������������ļ�
        If Save������Ϣ(lng����id) = False Then
            Exit Function
        End If
   End If
   WriteDebugDate_�����山 "��ȡ�ʸ�������Ϣ����ʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    '������:д�뱾������ϸ������ı�
    '   ��ȡ������ϸ��¼
    WriteDebugDate_�����山 "��ȡ��ϸ��Ϣ��ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g�������_�����山.�����־ = 1 Then
        '������סԺ,��ȷ�����ξ����������ϸ����
        gstrSQL = Get��ϸ��¼(0)
    Else
        'ֻ�Ǳ��ν������ϸ����
        gstrSQL = Get��ϸ��¼(lng����ID)
    End If
    
    Call OpenRecordset(rs��ϸ, "��ȡ������ϸ")
    
    If rs��ϸ.RecordCount = 0 Then
        ShowMsgbox "û����ϸ��¼�����������Ŀδ������Ӧ�Ķ���"
        '��������¼����
        GoTo CancelRecordVerify:
        Exit Function
    End If
    WriteDebugDate_�����山 "��ȡ��ϸ��Ϣ����ʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    WriteDebugDate_�����山 "������ϸ���ݿ�ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g�������_�����山.�����־ = 1 Then
        '�����������ϸ��¼
    Else
        If Saveҽ����ϸ����(rs��ϸ) = False Then

            '��������¼����
            GoTo CancelRecordVerify:
            Exit Function
        End If
    End If
    WriteDebugDate_�����山 "������ϸ���ݽ���ʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    WriteDebugDate_�����山 "������ϸ�����ı��ļ���ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If Save������ϸ�ı��ļ�(rs��ϸ) = False Then
        '��������¼����
        GoTo CancelRecordVerify:
        Exit Function
    End If
    WriteDebugDate_�����山 "������ϸ�����ı��ļ�����ʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")

    WriteDebugDate_�����山 "������ʷ���ý�������ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    '���Ĳ�:������ʷ�ķ��ý�����
    If g�������_�����山.�����־ = 1 Then
        rs��ϸ.MoveFirst
        g�������_�����山.lng��ҳID = Nvl(rs��ϸ!��ҳID, 0)
        If g�������_�����山.������� Then
            If Save��ʷ���ý������ı�(g�������_�����山.lng����id, Nvl(rs��ϸ!��ҳID, 0), False) = False Then
                    '��������¼����
                    GoTo CancelRecordVerify:
                    Exit Function
            End If
        End If
    Else
        If Save��ʷ���ý������ı�(0, 0) = False Then
                '��������¼����
                GoTo CancelRecordVerify:
                Exit Function
        End If
    End If
    WriteDebugDate_�����山 "������ʷ���ý���������ʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
        
    '���岽:����б��ؼ���,�����ý���
    WriteDebugDate_�����山 "���ط��ý��㿪ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If ���˷��ý���(lng����id, 0) = False Then

        '��������¼����
        GoTo CancelRecordVerify:
        Exit Function
    End If
    WriteDebugDate_�����山 "���ط��ý������ʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    '������:������ϸ�ϴ�
    WriteDebugDate_�����山 "�����ϴ���ʼʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g�������_�����山.������� Then
        '��������ò����ϴ���ϸ
    Else
        If ������ϸ�ϴ�(rs��ϸ) = False Then
            ShowMsgbox "�ڽ��д�����ϸ�ϴ�ʱ����һ�����ϵ���ϸ�ϴ�ʧ��,���Ժ�ע�ⲹ��!"
        End If
    End If
    ���˽��� = True
    WriteDebugDate_�����山 "�����ϴ�����ʱ��:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    GoTo CancelRecordVerify:
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
CancelRecordVerify:
    If g�������_�����山.�����־ <> 1 Then
    
        '��������¼����
        Call ������¼����_�����山
    Else
        If g�������_�����山.������� = False And ҽ�������Ѿ���Ժ(lng����id) = True And ���˽��� = True Then
            '��������¼����
            Call ������¼����_�����山
        End If
    End If
End Function
Public Sub WriteDebugInfor_�����山(ByVal strInfor As String)
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

Public Sub WriteDebugDate_�����山(ByVal strInfor As String)
        '��������Ϣд���ļ���
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        Dim intDebug As Integer
        
        intDebug = GetSetting("ZLSOFT", "ҽ��", "����ʱ��", 0)
        If intDebug <> 1 Then Exit Sub

        Dim strFile As String
        Dim rsTemp As New ADODB.Recordset
        strFile = App.Path & "\Test.log"
        
        If Not Dir(strFile) <> "" Then
            objFile.CreateTextFile strFile
        End If
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        If InStr(1, strInfor, "==") <> 0 Then
            objText.WriteLine strInfor
        Else
            objText.WriteLine "����:" & g�������_�����山.���� & vbTab & strInfor
        End If
        objText.Close
        
End Sub

Private Function insertInto����(ByVal lng���id As Long, ByVal XMLTEXT As String) As Boolean
    '����:
    '���̲���:
    '   ZL_���ý�������_INSERT
    '���ID_IN       IN ���ý�������.���ID%TYPE,
    '���_IN
    '���޽��_IN     IN ���ý�������.���޽��%TYPE,
    '�Ը����_IN     IN ���ý�������.�Ը����%TYPE,
    '֧�����_IN     IN ���ý�������.֧�����%TYPE,
    '����Ա����_IN   IN ���ý�������.����Ա����%TYPE,
    '���Ը�_IN       IN ���ý�������.���Ը�%TYPE
    DebugTool "������ý����������"
    If Trim(XMLTEXT) = "" Then insertInto���� = True: Exit Function
    insertInto���� = False
    Set gobjXMLOutput = New MSXML2.DOMDocument

    If GetXML��(XMLTEXT) = False Then Exit Function
    
    Dim lngCount As Long
    Dim lngRow As Long
    
    lngCount = GetOutXMLRows("SubRkn")
    
    
    Err = 0
    On Error GoTo ErrHand:
    For lngRow = 0 To lngCount - 1
        gstrSQL = "ZL_���ý�������_INSERT("
        gstrSQL = gstrSQL & "" & lng���id & ","
        gstrSQL = gstrSQL & "" & lngRow & ","
        
        'AKA160  number  14  2   �Ӷ����޽��
        gstrSQL = gstrSQL & Val(GetXMLOutput("aka160", , lngRow)) & ","
        'YKA106  number  14  2   �Ը����
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka106", , lngRow)) & ","
        'YKA 107 number  14  2   ֧�����
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka107", , lngRow)) & ","
        'YKA 063 number  14  2   ����Ա�������
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka063", , lngRow)) & ","
        'YKA057  number  14  2   �����Ը�����
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka057", , lngRow)) & ")"
        gcnOracle_CQYB.Execute gstrSQL
    Next

    insertInto���� = True
    Exit Function
ErrHand:
    DebugTool "ִ���������" & vbCrLf & " �����:" & Err.Number & vbCrLf & "������Ϣ:" & Err.Description
End Function







