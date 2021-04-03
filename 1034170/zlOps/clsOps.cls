VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################

Public Version As String

'######################################################################################################################

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '******************************************************************************************************************
    Set gcnOracle = cnMain
    
    '��ȡ��ʼֵ
    '------------------------------------------------------------------------------------------------------------------
    ParamInfo.ϵͳ���� = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrSysName = ParamInfo.ϵͳ����
    gblnShowInTaskBar = blnShowInTaskBar
    gblnInsure = True
    
    gstrDBUser = strDBUser
    UserInfo.���ݿ��û� = strDBUser
    
    Call GetUserInfo
    
    UserInfo.ģ��Ȩ�� = GetPrivFunc(lngSys, lngModul)
    UserInfo.��λ���� = zlRegInfo("��λ����")
    ParamInfo.ģ��� = lngModul
    ParamInfo.ϵͳ�� = lngSys
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, ParamInfo.ϵͳ��, UserInfo.ģ��Ȩ��)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, ParamInfo.ϵͳ��, blnShowInTaskBar)
    
    
    '��ȡϵͳ����
    '------------------------------------------------------------------------------------------------------------------
    Call InitSysPara
    
    'ִ��ģ��
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    Select Case lngModul
    Case 1801               '������������
        frmOpsScheme.Show , frmMain
    Case 1804               '�����ҹ���վ
        Call gclsInsure.InitOracle(gcnOracle)
        frmOpsStation.Show , frmMain
    End Select
    
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '******************************************************************************************************************
    Set gcnOracle = cnMain
    
    '��ȡ��ʼֵ
    '------------------------------------------------------------------------------------------------------------------
    ParamInfo.ϵͳ���� = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrSysName = ParamInfo.ϵͳ����
    gblnShowInTaskBar = blnShowInTaskBar
    gblnInsure = True
    
    gstrDBUser = strDBUser
    UserInfo.���ݿ��û� = strDBUser
    
    Call GetUserInfo
    
    UserInfo.ģ��Ȩ�� = GetPrivFunc(lngSys, lngModul)
    UserInfo.��λ���� = zlRegInfo("��λ����")
    ParamInfo.ģ��� = lngModul
    ParamInfo.ϵͳ�� = lngSys
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, Nothing, ParamInfo.ϵͳ��, UserInfo.ģ��Ȩ��)
    Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, ParamInfo.ϵͳ��, blnShowInTaskBar)
    
    '��ȡϵͳ����
    '------------------------------------------------------------------------------------------------------------------
    Call InitSysPara
    
    'ִ��ģ��
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    Select Case lngModul
    Case 1801               '������������
        zlCommFun.ShowChildWindow frmOpsScheme.hWnd, lngMain
    Case 1804               '�����ҹ���վ
        Call gclsInsure.InitOracle(gcnOracle)
        zlCommFun.ShowChildWindow frmOpsStation.hWnd, lngMain, True
    End Select
    
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
    CloseWindows = blnChildren And (Forms.Count = 0)

End Function

Private Sub Class_Initialize()
    Version = App.major & "." & App.minor & "." & App.Revision
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
End Sub
