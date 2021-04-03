VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutExse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ҫ���������
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���
    Set gfrmMain = frmMain: glngMain = 0
    If Not CheckValid Then Exit Sub

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    If glngSys Like "8??" Then
        App.Title = "����ҩ�����"
        gstrCustomerAppellation = "�ͻ�"
    Else
        gstrCustomerAppellation = "����"
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ز�����ģ��Ĺ������弤��ʱ���ٴζ�ȡ(�����ȶ�����ΪForm_Loadʱ���ܻ��õ�),��Ϊģ��֮�䴰���л���ı�ȫ�ֱ�����ֵ
    Call InitLocPar(lngModul)
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, lngModul)
    Call CreatePlugIn(lngModul)
    Call InitSysPar
    
    Select Case glngModul
        Case 1120 'ҩ�����۹���
            frmManagePrice.Show , frmMain
        Case 1121 '�����շѹ���
            If gfrmMain Is Nothing Then
                frmManageCharge.Show
            Else
                frmManageCharge.Show , frmMain
            End If
        Case 1122 '������ʹ���
            frmManageBilling.Show , frmMain
        Case 1124 '���ղ������
            frmReplenishTheBalanceManage.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���
    Set gfrmMain = Nothing: glngMain = lngMain
    If glngMain = 0 Then glngMain = -1
    
    If Not CheckValid Then Exit Sub

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    If glngSys Like "8??" Then
        App.Title = "����ҩ�����"
        gstrCustomerAppellation = "�ͻ�"
    Else
        gstrCustomerAppellation = "����"
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ز�����ģ��Ĺ������弤��ʱ���ٴζ�ȡ(�����ȶ�����ΪForm_Loadʱ���ܻ��õ�),��Ϊģ��֮�䴰���л���ı�ȫ�ֱ�����ֵ
    Call InitLocPar(lngModul)
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(gfrmMain, lngModul)
    Call CreatePlugIn(lngModul)
    Call InitSysPar
    
    Select Case glngModul
        Case 1120 'ҩ�����۹���
            zlCommFun.ShowChildWindow frmManagePrice.hWnd, lngMain
        Case 1121 '�����շѹ���
            zlCommFun.ShowChildWindow frmManageCharge.hWnd, lngMain
        Case 1122 '������ʹ���
            zlCommFun.ShowChildWindow frmManageBilling.hWnd, lngMain
        Case 1124 '���ղ������
            zlCommFun.ShowChildWindow frmReplenishTheBalanceManage.hWnd, lngMain
    End Select
End Sub

Public Function Charge(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strDBUser As String, ByVal bytInFun As Byte, ByVal bytInState As Byte) As Boolean
'���ܣ�ֱ�ӵ����շ�/���۹���
'���أ��Ƿ��ѻ���
    '��������
    Static blnNoFirst As Boolean  '�Ƿ��һ�ε���,��ʼ��ȫ�ֱ���
    
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    
    If Not blnNoFirst Then
        blnNoFirst = True
        
        Set gfrmMain = frmMain: glngMain = 0
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        glngSys = lngSys
        gstrDBUser = strDBUser
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrUnitName = GetUnitName
        gstrPrivs = ";" & GetPrivFunc(lngSys, Choose(bytInFun + 1, 1121, 1120, 1122)) & ";"
        
        If InStr(gstrPrivs, ";����;") = 0 Then
            MsgBox "��û�б�����ʹ�û��۹��ܵ�Ȩ��,����ʹ��,����ϵͳ����Ա��ϵ!", vbInformation, gstrSysName
            Exit Function
        End If
                    
        '�û���Ϣ
        Call GetUserInfo
            
    End If
    '��ʼ������
    Call InitLocPar(Choose(bytInFun + 1, 1121, 1120, 1122))
    Call InitSysPar
    Call CreateSquareCardObject(gfrmMain, Choose(bytInFun + 1, 1121, 1120, 1122))
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(Choose(bytInFun + 1, 1121, 1120, 1122))
    
    '���û���
    Err.Clear
    On Error Resume Next
    frmCharge.mlngModul = Choose(bytInFun + 1, 1121, 1120, 1122)
    frmCharge.mstrPrivs = gstrPrivs
    frmCharge.mbytInFun = bytInFun
    frmCharge.mbytInState = bytInState
    If gfrmMain Is Nothing And glngMain = 0 Then
        frmCharge.Show
    Else
        frmCharge.Show 1, frmMain
    End If
    
    Charge = gblnOK
End Function
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Load frmPrint
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
    CloseSquareCardObject
    Set grsҽ�Ƹ��ʽ = Nothing
    Unload frmPrint
End Sub
Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function zlCheckExcuteItemValied(frmMain As Object, cnOracle As ADODB.Connection, ByVal strDBUser As String, _
                                         ByVal lngSys As Long, ByVal lngModule As Long, ByVal lng����ID As Long, _
                                         ByVal str�շ���� As String, ByVal strNos As String, ByVal str�շ�ϸĿIDs As String) _
                                         As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�����ƺ���㷽ʽ�����ִ����Ŀ�ĺϷ���
    '���:frmMain-���õ�������
    '     cnOracle-����Oracle����
    '     strDBUser-���ݿ��û���
    '     lngSys-ϵͳ��
    '     lngModule-���õ�ģ���(��Ҫ�ǿ��Ǻ������ܴ��ڸ���ģ���������ж�)
    '     lng����ID-��ǰ���Ĳ���ID
    '     str�շ����-����ִ�е��շ����,���Դ�����,���ʱ�ö��ŷ���:����:4,5,6,7��,Ϊ��ʱ���������շ������м��
    '     strNOs-����ִ�еĵ��ݺ�(���ʵ���),���Դ����������ʱ�ö��ŷ���,����:4,5�ȣ�Ϊ��ʱ�������ݵ��ݺŽ��м��
    '     str�շ�ϸĿIDs-��ʾ����ִ�е��շ���Ŀ,����Ϊ������ö��ŷ���,���磺123,456�ȣ�Ϊ��ʱ�������շ�ϸĿIDs���м��
    '����:True��ʾ���ݺϷ�,����ִ�У�False-��ʾ���ݲ��Ϸ���������ִ�и���Ŀ
    '����:������
    '����:2013-12-24
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim rsCheck As New ADODB.Recordset, strSQL As String
    Dim strReturn As String, strArray() As String
    Set gcnOracle = cnOracle
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrDBUser = strDBUser: glngSys = lngSys: glngModul = lngModule
    strSQL = "Select zl_Get_ExcuteItem_Infor([1],[2],[3],[4],[5]) As ������Ϣ From Dual"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "���ִ����Ŀ", lngModule, lng����ID, str�շ����, strNos, str�շ�ϸĿIDs)
    strReturn = Nvl(rsCheck!������Ϣ)
    If strReturn <> "" And InStr(strReturn, "|") > 0 Then
        strArray = Split(strReturn, "|")
        If Val(strArray(0)) = 2 Then
            MsgBox "���ݲ��Ϸ�����ִֹ�У�" & vbCrLf & "��ʾ��Ϣ:" & strArray(1), vbInformation, gstrSysName
            zlCheckExcuteItemValied = False
            Exit Function
        End If
        If Val(strArray(0)) = 1 Then
            If MsgBox("ȷ��Ҫִ����Ŀ��" & vbCrLf & "��ʾ��Ϣ:" & strArray(1), vbYesNo + vbQuestion, gstrSysName) = vbYes Then
                zlCheckExcuteItemValied = True
                Exit Function
            Else
                zlCheckExcuteItemValied = False
                Exit Function
            End If
        End If
    End If
    If strReturn = "" Then
        zlCheckExcuteItemValied = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function