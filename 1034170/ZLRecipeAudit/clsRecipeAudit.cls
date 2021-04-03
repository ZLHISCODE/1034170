VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecipeAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
'˵��: ΪZLHIS����̨���ŵĽӿ���
'--------------------------------------------------------------------------------
Option Explicit

Private mstrVersion As String

Public Property Get Version() As String
    mstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    Version = mstrVersion
End Property

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, ByRef cnMain As ADODB.Connection, ByRef frmMain As Object, ByRef strDbUser As String)
'����:ZLHIS����̨����ģ��ķ���
'����:
'  lngSys��ϵͳ��
'  lngModul��ģ���
'  cnMain��Oracle���ݿ����Ӷ���
'  frmMain������̨�������
'  strDBUser�����ݿ��¼�û���

    glngSys = lngSys
    glngModule = lngModul
    UserInfo.�û��� = strDbUser
    Set gcnOracle = cnMain
    If gfrmNavigation Is Nothing Then
        Set gfrmNavigation = frmMain
    End If
        
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")       '��ȡ����ϵͳ����
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")   '��ȡ��Ʒ����
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrUnitName = zl9ComLib.GetUnitName    '��ȡע��ĵ�λ������Ϣ
    
    Select Case glngModule
        Case enuModule.���ﴦ�����_1351
            frmRAOutPatient.Show , frmMain
            
        Case enuModule.סԺҩ�����_1352
            frmRAInPatient.Show , frmMain
            
        Case enuModule.���������Ŀ_1353
            frmRAItems.Show , frmMain
        
        Case enuModule.�����������_1354
            frmRASetup.Show , frmMain
        
        Case enuModule.�������ͳ��_1355
            frmRAStatistics.Show , frmMain
    End Select
    
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, ByRef cnMain As ADODB.Connection, ByRef lngMain As Long, ByRef strDbUser As String)
'����:BH����̨����ģ��ķ���
'����:
'  lngSys��ϵͳ��
'  lngModul��ģ���
'  cnMain��Oracle���ݿ����Ӷ���
'  lngMain������̨�������ľ����
'  strDBUser�����ݿ��¼�û���

    glngSys = lngSys
    glngModule = lngModul
    UserInfo.�û��� = strDbUser
    Set gcnOracle = cnMain
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrUnitName = zl9ComLib.GetUnitName
    
    Select Case glngModule
        Case enuModule.���ﴦ�����_1351
            zlCommFun.ShowChildWindow frmRAOutPatient.hwnd, lngMain
        
        Case enuModule.סԺҩ�����_1352
            zlCommFun.ShowChildWindow frmRAInPatient.hwnd, lngMain
        
        Case enuModule.���������Ŀ_1353
            zlCommFun.ShowChildWindow frmRAItems.hwnd, lngMain
        
        Case enuModule.�����������_1354
            zlCommFun.ShowChildWindow frmRASetup.hwnd, lngMain
            
        Case enuModule.�������ͳ��_1355
            zlCommFun.ShowChildWindow frmRAStatistics.hwnd, lngMain
    End Select
End Sub

Public Function CloseWindows() As Boolean
'����:�رղ��������еĴ���
'����:True�ɹ���Falseʧ��

    Dim frmThis As Form
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    gintHoursRecipe = 24
End Sub