VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################

Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ��¼���ݿ��û���
    '******************************************************************************************************************
    
    Set gcnOracle = cnMain
    
    ParamInfo.ϵͳ���� = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    gstrDBUser = strDBUser
    UserInfo.���ݿ��û� = strDBUser
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, ParamInfo.ϵͳ����
        Exit Sub
    End If

    Set gfrmMain = frmMain
    Set gobjEmr = frmMain.mobjEmr
    
    UserInfo.ģ��Ȩ�� = GetPrivFunc(lngSys, lngModul)
    UserInfo.��λ���� = zlRegInfo("��λ����")
    gstrPrivs = UserInfo.ģ��Ȩ��
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    glngModul = lngModul
    ParamInfo.ģ��� = lngModul
    ParamInfo.ϵͳ�� = lngSys
    ParamInfo.����RIS = Val(zlDatabase.GetPara(255, lngSys)) = 1
    glngSys = lngSys
    
    gstrUserName = UserInfo.����
    SetDbUser UserInfo.�û���
    gstrDeptName = UserInfo.��������
    glngDeptId = UserInfo.����ID
    gblnInsure = True
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, ParamInfo.ϵͳ��, UserInfo.ģ��Ȩ��)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, ParamInfo.ϵͳ��, blnShowInTaskBar)
    Call gobjPath.InitCISPath(gcnOracle, frmMain, ParamInfo.ϵͳ��, UserInfo.ģ��Ȩ��)
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    
    Call InitSysPara
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand

    Select Case lngModul
    Case 1550                  '�������ֱ�׼
        frm���ֱ�׼ά��.Show , frmMain
    Case 1551                   '��������׼
        frmAuditItem.Show , frmMain
    Case 1560                   '���Ӳ������
        If IsPrivs(gstrPrivs, "������鲡��") = False And IsPrivs(gstrPrivs, "���Ĺ鵵����") = False And IsPrivs(gstrPrivs, "���ĳ�鲡��") = False Then
            ShowSimpleMsg "�Բ�����û�в鿴���Ӳ������ģ����κ�Ȩ�ޣ����ȵ��û���Ȩ���������ã�"
            Exit Sub
        End If
        frmCISAduit.Show , frmMain
    Case 1561                   '���Ӳ�������
        frmCISBorrow.Show , frmMain
    Case 1562                   '���Ӳ�������
        frm��������.Show , frmMain
    Case 1563                   '������������
        frmEPRAuditMan.Show , frmMain
    Case 1564                   '���Ӳ�������
        frm�������չ���.Show , frmMain
    Case 1566                   '���Ӳ�����ӡ  80880
        frmMedRecPrint.Show , frmMain
    End Select
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
    '******************************************************************************************************************
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '   strDBUser:��ǰ��¼���ݿ��û���
    '******************************************************************************************************************
    
    Set gcnOracle = cnMain
    
    ParamInfo.ϵͳ���� = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    UserInfo.���ݿ��û� = strDBUser
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, ParamInfo.ϵͳ����
        Exit Sub
    End If

    Set gfrmMain = Nothing
    If Not objLogin Is Nothing Then
        Set gobjEmr = objLogin.mobjEmr
    End If

    UserInfo.ģ��Ȩ�� = GetPrivFunc(lngSys, lngModul)
    UserInfo.��λ���� = zlRegInfo("��λ����")
    gstrPrivs = UserInfo.ģ��Ȩ��
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    glngModul = lngModul
    ParamInfo.ģ��� = lngModul
    ParamInfo.ϵͳ�� = lngSys
    ParamInfo.����RIS = Val(zlDatabase.GetPara(255, lngSys)) = 1
    glngSys = lngSys
    
    gstrUserName = UserInfo.����
    SetDbUser UserInfo.�û���
    gstrDeptName = UserInfo.��������
    glngDeptId = UserInfo.����ID
    gblnInsure = True
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, Nothing, ParamInfo.ϵͳ��, UserInfo.ģ��Ȩ��)
    Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, ParamInfo.ϵͳ��, True)
    Call gobjPath.InitCISPath(gcnOracle, Nothing, ParamInfo.ϵͳ��, UserInfo.ģ��Ȩ��)
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    Call InitSysPara
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand

    Select Case lngModul
    Case 1550                  '�������ֱ�׼
        zlCommFun.ShowChildWindow frm���ֱ�׼ά��.hWnd, lngMain
    Case 1551                   '�������Ŀ¼
        zlCommFun.ShowChildWindow frmAuditItem.hWnd, lngMain
    Case 1560                   '���Ӳ������
        If IsPrivs(gstrPrivs, "������鲡��") = False And IsPrivs(gstrPrivs, "���Ĺ鵵����") = False And IsPrivs(gstrPrivs, "���ĳ�鲡��") = False Then
            ShowSimpleMsg "�Բ�����û�в鿴���Ӳ������ģ����κ�Ȩ�ޣ����ȵ��û���Ȩ���������ã�"
            Exit Sub
        End If
        
        zlCommFun.ShowChildWindow frmCISAduit.hWnd, lngMain, True
    Case 1561                   '���Ӳ�������
        zlCommFun.ShowChildWindow frmCISBorrow.hWnd, lngMain, True
    Case 1562                   '���Ӳ�������
        zlCommFun.ShowChildWindow frm��������.hWnd, lngMain, True
    Case 1563                   '������������
        zlCommFun.ShowChildWindow frmEPRAuditMan.hWnd, lngMain, True
    Case 1564                   '���Ӳ�������
        zlCommFun.ShowChildWindow frm�������չ���.hWnd, lngMain, True
    Case 1566                   '���Ӳ�����ӡ  80880
        zlCommFun.ShowChildWindow frmMedRecPrint.hWnd, lngMain, True
    End Select
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CloseWindows() As Boolean
    '******************************************************************************************************************
    '����:�ر������Ӵ���
    '******************************************************************************************************************
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    
    blnChildren = gobjRichEPR.CloseWindows()
    blnChildren = blnChildren And gobjKernel.CloseWindows()
    blnChildren = blnChildren And gobjPath.CloseWindows()
    CloseWindows = blnChildren And (Forms.count = 0)
    
End Function

Private Sub Class_Initialize()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
        
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub Class_Terminate()

    On Error Resume Next

'    Unload frmPubIcons
    Unload frmPubResource
End Sub

