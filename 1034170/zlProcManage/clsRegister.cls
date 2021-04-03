VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'����������ZL9Rigister����
Private mcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Private mcolPrivs As Collection              '��ǰ�û��߱������г���Ĺ���Ȩ��
Private mlngCount As Long
Private mstrUserName As String
Private mstrPassword As String
Private mstrServer As String

Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum
'zlRegCheck,zlRegInfoʹ�þɵ�
Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    Static srsInfo As New ADODB.Recordset
    Static sblnTemp As Boolean
    Dim strInfo As String, aryInfo() As String
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHand
    If blnTemp Or sblnTemp <> blnTemp Or (srsInfo.State <> adStateOpen) Then
        sblnTemp = blnTemp
        strSQL = "Select Item, Text From Table(Cast(zltools.f_Reg_Info(" & IIf(blnTemp, 1, 0) & ") As zlTools.t_Reg_Rowset))"
        If srsInfo.State = adStateOpen Then srsInfo.Close
        srsInfo.Open strSQL, mcnOracle, adOpenKeyset, adLockReadOnly
    End If
    
    srsInfo.Filter = "Item='" & strItem & "'"
    If srsInfo.RecordCount <> 1 Then zlRegInfo = "": Exit Function
    strInfo = "" & srsInfo!Text
    If (strItem = "��λ����" Or strItem = "��Ʒ������" Or strItem = "����֧����") And intBits <> -1 Then
        aryInfo = Split(strInfo, ";")
        If intBits > UBound(aryInfo) Then
            strInfo = ""
        Else
            strInfo = aryInfo(intBits)
        End If
    End If
    zlRegInfo = strInfo
    Exit Function

errHand:
    zlRegInfo = ""
End Function

'--------------------------------------------------
'���ܣ���֤ϵͳע����Ȩ����ȷ��
'������blnTemp-�Ƿ��δ�������ʱע����Ϣ��֤
'���أ���ȷ����"";���󷵻ش�����Ϣ
'--------------------------------------------------
Public Function zlRegCheck(Optional blnTemp As Boolean) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHand
    
    If blnTemp Then
        strSQL = "Select zltools.f_Reg_Audit(1) As Stamp From zlRegFile r Where ��Ŀ = '��Ȩ֤��'"
    Else
        strSQL = "Select zltools.f_Reg_Audit(0) As Stamp From zlRegInfo r Where ��Ŀ = '��Ȩ֤��'"
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open strSQL, mcnOracle, adOpenKeyset, adLockReadOnly
    If rsTemp.RecordCount <= 0 Then
        zlRegCheck = "����Ȩ��Ϣ!"
        Exit Function
    End If
    
    If Left(rsTemp.Fields(0).value, 6) <> "ERROR-" Then
        zlRegCheck = ""
    Else
        zlRegCheck = rsTemp.Fields(0).value
    End If
    Exit Function
    
errHand:
    zlRegCheck = Err.Description
End Function

Public Function zlRegTool(Optional ByVal blnTemp As Boolean) As Long
'���ܣ������Ȩ������Ϣ
'������blnTemp  :�Ƿ��δ�������ʱע����Ϣ��֤
'���أ���2�Ĺ���ĩλ�η����ع�������
    Dim rsTool As ADODB.Recordset
    Dim strSQL As String, lngRetu As Long
    
    On Error GoTo errHand
    strSQL = "Select Prog From Table(Cast(zltools.f_Reg_Tool([1]) As zlTools.t_Reg_Rowset))"
    Set rsTool = OpenSQLRecord(strSQL, "zlRegTool", IIf(blnTemp, 1, 0))
    lngRetu = 0
    Do While Not rsTool.EOF
        lngRetu = lngRetu + 2 ^ ((Val("" & rsTool.Fields(0).value) Mod 10) - 1)
        rsTool.MoveNext
    Loop
    zlRegTool = lngRetu
    Exit Function
errHand:
    MsgBox "��ȡע�Ṥ�߳���:" & vbNewLine & Err.Description, vbExclamation, "zlRegister"
    zlRegTool = 0
End Function

Public Function zlRegFunc(lngSys As Long, lngProgId As Long) As String
'���ܣ����ص�ǰ�û����е�ָ������Ĺ��ܴ�
'������lngSys     :����ǹ̶�ģ�飬��Ϊ0
'      lngProgId  :�������(ģ���)
'���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��
    Dim rsTmp As ADODB.Recordset, blnNew As Boolean
    Dim strSQL As String, strPrivs As String
    Dim blnRegCheck As Boolean
        
    If mcolPrivs Is Nothing Then
        Set mcolPrivs = New Collection
        blnNew = True
    Else
        On Error Resume Next
        strPrivs = mcolPrivs("_" & lngSys & "_" & lngProgId)
        If Err.Number > 0 Then
            blnNew = True
            Err.Clear
            On Error GoTo 0
        End If
    End If
    
    If blnNew Then
        strSQL = "Select Text as ���� From Table(Cast(zltools.f_Reg_Func([1],[2]) as zlTools.t_Reg_Rowset))"
        On Error GoTo errHand

        Set rsTmp = OpenSQLRecord(strSQL, "GetPrivFunc", lngSys, lngProgId)
                
        Do While Not rsTmp.EOF
            strPrivs = strPrivs & ";" & rsTmp!����
            rsTmp.MoveNext
        Loop
        strPrivs = Mid(strPrivs, 2)
        mcolPrivs.Add strPrivs, "_" & lngSys & "_" & lngProgId
    End If
        
    zlRegFunc = strPrivs
    Exit Function
errHand:
    MsgBox "��ȡע�Ṧ�ܳ���:" & vbNewLine & Err.Description, vbExclamation, "zlRegister"
    zlRegFunc = ""
End Function

Public Function zlRegBuild(ByVal strFile As String, ByRef pgbState As Object) As Boolean
'���ܣ�����Ȩע���ļ���Ϣд����ʱ��
'������strFile  :ע�����ļ�·�������ļ�����
'      pgbState :�������ؼ�(ProgressBar)
    Dim rsFile As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errHand
    
    With rsFile
        .Open strFile, , adOpenKeyset, adLockReadOnly
        
        pgbState.Visible = True
        pgbState.Max = .RecordCount
                
        mcnOracle.Execute "Delete zlRegFile"
                
        strSQL = ""
        i = 1
        Do While Not .EOF
            If i = 1 Then
                strSQL = "Select '" & !��Ŀ & "'," & !ϵͳ & "," & !��� & ", '" & !���� & "' From Dual"
            Else
                strSQL = strSQL & " Union All Select '" & !��Ŀ & "'," & !ϵͳ & "," & !��� & ", '" & !���� & "' From Dual"
            End If
            
            If i = 100 Then
                strSQL = "Insert Into zlRegFile(��Ŀ, ϵͳ, ���, ����) " & strSQL
                mcnOracle.Execute strSQL
                i = 1
                strSQL = ""
            Else
                i = i + 1
            End If
            
            pgbState.value = .Bookmark
            .MoveNext
        Loop
        If strSQL <> "" Then
            strSQL = "Insert Into zlRegFile(��Ŀ, ϵͳ, ���, ����) " & strSQL
            mcnOracle.Execute strSQL
        End If
    End With
    pgbState.Visible = False
    
    zlRegBuild = True
    Exit Function
errHand:
    MsgBox "ע����Ȩ�ļ�д�����ݿ�ʱ���ִ������飡" & vbNewLine & Err.Description, vbExclamation, "zlRegister"
    pgbState.Visible = False
End Function

Private Sub Class_Initialize()
    mlngCount = mlngCount + 1
End Sub

Private Sub Class_Terminate()
    'ͬһ�������У��ڲ�ͬ������,�Բ�ͬ�Ķ����������set xx=nothing��ÿ�ζ��ᴥ�����¼�
    mlngCount = mlngCount - 1
    If mlngCount = 0 Then Set mcolPrivs = Nothing
End Sub

Public Sub zlRegInit(cnMain As ADODB.Connection)
'���ܣ���ʼ��ע�Ჿ���ڵ����Ӷ���(һ�����ڵ�¼����ã�������ȡһ�����Ӷ���ʱ���س�ʼ��)
    Set mcnOracle = cnMain
    
    If mstrUserName = "" Then
        '���֮ǰû�е��ù�GetConnection���������û���Ϊ�ա������Ӷ����ConnectionString�н��������Ӵ�
        '�������һ������ҽ�����ߵȲ��ӵ���̨��¼�ĳ����ڵ���zlComlib����ʱ��ʼ������
        Call GetConnectionInfo(cnMain.ConnectionString, mstrServer, mstrUserName, mstrPassword)
    End If
End Sub

Public Function GetConnection(ByVal strServer As String, ByVal strUserName As String, ByVal strPassword As String, ByVal blnTransPassword As Boolean, _
     Optional ByVal bytProvider As enuProvider = MSODBC, Optional ByRef strError As String = "���뷵�ش�����Ϣ", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
'���ܣ� ��ָ�������ݿ⣬��������ʵ������ADO���Ӷ���(�����10.35.10��ǰ�����룬���µ�ת�������������)
'������ strServer       :�������������߿���ֱ��ָ��IP:Port/SID
'       strUserName     :�û���
'       strPassword     :����
'       blnTransPassword:�Ƿ��������ת��
'       bytProvider     :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ
'       strError        :����ʧ�ܺ����ָ���˴˲������򷵻ش�����Ϣ��δָ��ʱֱ�ӵ�����ʾ��Ϣ��
'       blnSaveAccount  :�����û��������롢����������ȫ�ֱ�����һ�㣬���ڵ�¼����ʱ���棬���ӿ�ReGetConnection��GetUserName��GetServerName��GetPassword��LoginValidateʹ�ã�
'���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)
    Dim cnOracle As ADODB.Connection
    Dim strPWD_New As String, strPWD_Old As String
    Dim blnPrompt As Boolean
    
    blnPrompt = (strError = "���뷵�ش�����Ϣ")
    strError = ""
    
    If blnTransPassword Then
        strPWD_Old = TransPassword(strPassword)
    Else
        strPWD_Old = strPassword
    End If
    Set cnOracle = OraDBOpen(strServer, strUserName, strPWD_Old, bytProvider, strError, blnSaveAccount)
    
    If blnPrompt And strError <> "" Then
        MsgBox strError, vbExclamation, "��ʾ"
    End If
    
    mstrPassword = strPWD_Old
    Set GetConnection = cnOracle
End Function


Public Function ReGetConnection(ByVal bytProvider As enuProvider, ByRef strError As String) As ADODB.Connection
'���ܣ����ص�¼����̨ʱ�����Ӷ��󣬻��߸���֮ǰ�򿪵����ݿ����Ӷ������»�ȡһ��OLEDB��MSODBC��ʽ�򿪵����Ӷ���
'������bytProvider  :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ,9-��¼����̨ʱ�����Ӷ���
'      strError     :���ش�����ʧ�ܺ�Ĵ�����Ϣ
'���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)
    
    If bytProvider = OriginalConnection Then
        Set ReGetConnection = mcnOracle
    Else
        Set ReGetConnection = OraDBOpen(mstrServer, mstrUserName, mstrPassword, bytProvider, strError, False)
    End If
End Function

Public Function LoginValidate(ByVal strServer As String, ByVal strUserName As String, ByRef strPassword As String, ByRef strError As String, _
    Optional lngInstance As Long) As Boolean
'���ܣ����ݷ����������û�����������֤�û���¼�������10.35.10��ǰ�����룬���Զ����µ�ת������������룩
'������strServer    :�������������߿���ֱ��ָ��IP:Port/SID,��������ֵ����ȡ��¼ϵͳ(����GetConnection����ʱ)ʹ�õķ�������
'      strUserName  :�û���
'      strPassword  :����ת���������(ָ���ĳ���ʹ���ŷ���ת����ģ�δָ�����򷵻ش�����ʾ��Ϣ)
'      strError     :��֤ʧ��ʱ���ش�����Ϣ
'      lngInstance  :��ǰӦ�ó���ʵ���ľ���������Ҫ����ת��������룬��ǰû�д����������������̶�ʱ����Ҫ���룩
'���أ���֤��¼�Ƿ�ɹ�
    Dim cnOracle As ADODB.Connection
    Dim strPWD_New As String, strPWD_Old As String
        
    If strServer = "" Then strServer = mstrServer
    strPWD_New = TransPassword(strPassword)
        
    Set cnOracle = OraDBOpen(strServer, strUserName, strPWD_New, MSODBC, strError, False)
        
    LoginValidate = (cnOracle.State = adStateOpen)
    If LoginValidate Then
        strPassword = strPWD_New
        cnOracle.Close
    End If
End Function


Public Function CreateUser(ByVal cnOracle As ADODB.Connection, ByVal strUserName As String, ByVal strPassword As String, ByRef strError As String) As Boolean
'���ܣ������û�����������
'������ cnOracle        :ADO���Ӷ���
'       strUserName     :�û���
'       strPassword     :����
'       strError        :����ʧ�ܺ�Ĵ�����Ϣ
'���أ��Ƿ񴴽��ɹ�
    Dim strPWD_New As String, strSQL As String
    
    strPWD_New = TransPassword(strPassword)
    strSQL = "Create User " & strUserName & " identified by """ & strPWD_New & """"
    On Error Resume Next
    cnOracle.Execute strSQL
    
    '�û��Ѵ��ڡ����벻���ϸ��Ӷ�Ҫ��ȴ���
    If Err.Number <> 0 Then
        strError = Err.Description
    End If
    
    CreateUser = Err.Number = 0
End Function

Public Function UpdateUserPassword(ByVal cnOracle As ADODB.Connection, ByVal strUserName As String, ByVal strPassword As String, Optional ByVal blnTransPassword As Boolean = True, Optional ByRef strError As String) As Boolean
'���ܣ��޸��û�����(������޸��Լ������룬�����֮ǰ���������)
'������ cnOracle        :ADO���Ӷ���
'       strUserName     :�û���
'       strPassword     :�µ�����
'       blnTransPassword:�Ƿ�ת������
'       strError        :�޸�ʧ�ܺ�Ĵ�����Ϣ
'���أ��޸��Ƿ�ɹ�
    strError = UpdatePassword(cnOracle, strUserName, strPassword, blnTransPassword)
        
    If strError = "" And mstrUserName = strUserName Then mstrPassword = strPassword
    UpdateUserPassword = strError = ""
End Function

Public Function GetUserName(Optional ByVal blnUcase As Boolean = True) As String
'���ܣ���ȡ��ǰ��¼���û���
'������blnUcase     :�Ƿ�ת��Ϊ��д��ĸ
    GetUserName = IIf(blnUcase, UCase(mstrUserName), mstrUserName)
End Function

Public Function GetServerName(Optional ByVal blnUcase As Boolean = True) As String
'���ܣ���ȡ��ǰ��¼�ķ�������
'������blnUcase     :�Ƿ�ת��Ϊ��д��ĸ
    GetServerName = IIf(blnUcase, UCase(mstrServer), mstrServer)
End Function


Public Function GetPassword(Optional lngInstance As Long) As String
'���ܣ���ȡ��ǰ��¼(����GetConnection֮��)�û����ܺ������
'������lngInstance-��ǰӦ�ó���ʵ���ľ����û�д����������������̶�ʱ����Ҫ���룩
'˵����������ָ���ĳ���ʹ��ڷ��ʱ��ӿ�(δ����������£�������ʾ���֣����磺δ��Ȩ�ĳ���δ��Ȩ�Ĵ���)
    GetPassword = mstrPassword
End Function

'======================================================================
'˽�з���
'======================================================================
Private Function TransPassword(ByVal strPassword) As String
'���ܣ�ת���û�����,������û�����дת��
'������
    Dim intBit As Integer, intLen As Integer
    Dim strPWDNew As String, strBit As String
    
    strPassword = UCase(Trim(strPassword))

     If Len(strPassword) = 0 Then TransPassword = "": Exit Function
     
     strPWDNew = ""
     For intBit = 1 To Len(strPassword)
         strBit = Mid(strPassword, intBit, 1)
         Select Case (intBit Mod 3)
         Case 1
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                    strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                    strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                    strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
         Case 2
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                    strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                    strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                    strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
         Case 0
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                    strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                    strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                    strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
         End Select
     Next
    
    TransPassword = strPWDNew
End Function

Private Function OraDBOpen(ByVal strServer As String, ByVal strUserName As String, ByVal strPassword As String, _
    ByVal bytProvider As enuProvider, ByRef strError As String, ByVal blnSaveAccount As Boolean) As ADODB.Connection
'���ܣ� ��ָ�������ݿ⣬������ADO���Ӷ���
'������ strServer���������������߿���ֱ��ָ��IP:Port/SID
'       strUserName���û���
'       strUserPwd������
'       bytProvider�������ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ
'       blnSaveAccount�������û��������롢����������ȫ�ֱ�����һ�㣬��¼ʱ�����򱣴棬��֤�û�����ʱ�����棩
'���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    Dim strPersist_Security_Info As String
    Dim arrTmp As Variant, strIp As String, strPort As String, strSID As String
    
    On Error Resume Next

    Set OraDBOpen = New ADODB.Connection
        
    With OraDBOpen
        If InStr(strServer, "/") > 0 Then
            arrTmp = Split(strServer, "/")
            strSID = arrTmp(1)
            If InStr(arrTmp(0), ":") > 0 Then
                arrTmp = Split(arrTmp(0), ":")
                strIp = arrTmp(0)
                strPort = arrTmp(1)
            Else
                strIp = arrTmp(0)
                strPort = "1521"
            End If
            strServer = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIp & ")(PORT=" & strPort & "))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
            
            '�������ּ���ADDRESS_LIST��д������ODBC�£�ֻ֧��SID����֧��SERVICE_NAME;OLEDB�����ֶ�֧��
            'If bytProvider = enuProvider.MSODBC Then
            'strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIP & ")(PORT=" & strPort & ")))(CONNECT_DATA=(SID=" & strSID & ")))"
        End If
        
        '��Persist Security InfoΪfalseʱ�����Ӷ����ConnectionString�����У����������룬MSDataShape��ʽ��������������������,���ԣ���ģ������洢���Ա�����һ�����ӷ�ʽʱʹ��
        If blnSaveAccount Then
            mstrServer = strServer
            mstrUserName = strUserName
            mstrPassword = strPassword
        End If
        
        strPersist_Security_Info = ";Persist Security Info=False" '��������ߴӷ��ص����Ӷ����л���û����룬��ָ�������ԵĻ���ȱʡ��false
        
        'ȱʡΪadUseServer�������ָ�����䣬������OLEDB�򿪵����ӣ�����Command����Execute�������ص�Recordset�����ActiveConnection = Nothing�ᱨ��:�����ʱ����������(MSODBC��ʽ�򿪵����Ӳ��ᱨ��)
        .CursorLocation = adUseClient
        
        If bytProvider = enuProvider.MSODBC Then
            .Provider = "MSDataShape"
            .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer & strPersist_Security_Info, strUserName, strPassword
        Else
            .Provider = "OraOLEDB.Oracle"
            .Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassword
            'DistribTX=1,�����ֲ�����(ȱʡ);DistribTx=0:���ηֲ�����oracle8.1.7�汾��BUG������10.35.10֮ǰ�Ĺ������ߵ�¼ʱ�ǽ��õġ�
            'PLSQLRSet=1 ���ڲ��������α�����Ĵ洢���̣�Ҳ��д��Extended Properties=PLSQLRSet=1
        End If
    End With
    
    If Err = 0 Then
        strError = ""
    Else
        strError = Err.Description
        On Error GoTo 0
        
        If InStr(strError, "�Զ�������") > 0 Then
            If bytProvider = enuProvider.MSODBC Then
                strError = "msoracl32.dll"
            Else
                strError = "OraOLEDB.dll"
            End If
            strError = "�޷��������Ӷ����������ݷ��ʲ���(" & strError & ")�Ƿ�������װ��ע�ᡣ"
        ElseIf InStr(strError, "ORA-12505") > 0 Then
            strError = "ORA-12505,��������ǰ�޷�ʶ���������������������� SID,��������������õ�ʵ�����ơ�"
            
        ElseIf InStr(strError, "ORA-12170") > 0 Then
            strError = "ORA-12170,���ӳ�ʱ��������������Ƿ���ȷ�������Ƿ�ɷ��ʣ��Լ��Ƿ񱻷���������ǽ��ֹ��"
            
        ElseIf InStr(strError, "ORA-12154") > 0 Then
            strError = "ORA-12154,�޷���������������" & vbCrLf & "���鱾����Oracle�����ļ�(tnsnames.ora)���Ƿ���ڵ�ǰʹ�õķ�������"
            
        ElseIf InStr(strError, "ORA-12541") > 0 Then
            strError = "ORA-12541,�޷����ӷ�����������������ϵ�Oracle�����������Ƿ�������"
            
        ElseIf InStr(strError, "ORA-01033") > 0 Then
            strError = "ORA-01033,ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�"
            
        ElseIf InStr(strError, "ORA-01034") > 0 Then
            strError = "ORA-01034,ORACLE�����ã��������ݿ�ʵ���Ƿ�������"
            
        ElseIf InStr(strError, "ORA-02391") > 0 Then
            strError = "ORA-02391,�û�" & strUserName & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ������������¼��)��"
            
        ElseIf InStr(strError, "ORA-01017") > 0 Then
            strError = "ORA-01017,��Ч���û��������룬��¼���ܾ���"
        
        ElseIf InStr(strError, "ORA-28000") > 0 Then
            strError = "ORA-28000,���û��Ѿ������ã���������¼��"
        End If
    End If
End Function

Private Function UpdatePassword(ByVal cnOracle As ADODB.Connection, ByVal strUserName As String, ByVal strPassword As String, ByVal blnTransPassword As Boolean) As String
'���ܣ��޸��û�����
'���أ��޸ĳ�����Ĵ�������
    Dim strPWD_New As String
    
    If blnTransPassword Then
        strPWD_New = TransPassword(strPassword)
    Else
        strPWD_New = strPassword
    End If
    
    '���������븴�Ӷ�Ҫ��ʱ�������ж��ŵ������ַ���������Ҫ��˫����
    'Υ�����밲ȫ���Ե���������
    On Error Resume Next
    cnOracle.Execute "Alter user " & strUserName & " identified by """ & strPWD_New & """"
    
    If Err.Number <> 0 Then
       UpdatePassword = Err.Description
       Err.Clear
    End If
End Function

Private Sub GetConnectionInfo(ByVal strConect As String, ByRef strServerName As String, ByRef strUserName As String, ByRef strPassword As String)
'���ܣ� ����MSODBC���Ӷ����е�ORACLE���е� ���������û���������
'���أ� �ɹ�ʧ�ܣ�����True��ʧ�ܣ�����False

    Dim i As Integer
    Dim strTemp As String
    If strConect = "" Then Exit Sub
            
    strServerName = ""
    strUserName = ""
    strPassword = ""
    strConect = Replace(strConect, """", "")
    
    If InStr(strConect, "ODBC") > 0 Then
        'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=DYYY";Persist Security Info=True;User ID=zlhis;Password=his;Data Provider=MSDASQL"
        'Provider=MSDataShape.1;Persist Security Info=False;User ID=ZLHIS;Data Provider=MSDASQL;
        '��ȡ strServerName(SecurityΪfalseʱ���޷����)
        i = InStrRev(strConect, "Server=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 6)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            End If
        End If
    Else
        'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
        'Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=ZLHIS;Data Source="DYYY"
        i = InStrRev(strConect, "Data Source=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 11)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            Else    'SecurityΪfalseʱ��û��;��
                strServerName = strTemp
            End If
        End If
    End If
    
    '��ȡ strUserName
    i = InStrRev(strConect, "User ID=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 7)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strUserName = Left(strTemp, i - 1)
        End If
    End If
    
    '��ȡ strPassword
    i = InStrRev(strConect, "Password=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 8)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strPassword = Left(strTemp, i - 1)
        End If
    End If
End Sub
