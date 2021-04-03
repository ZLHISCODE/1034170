VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPACSWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VERSION As String

'��ҩ�䷽��������Ҫ�Ĳ�����Ϣ
Public Type TYPE_PatiInfoEx
        lng����ID As Long
        lng��ҳID As Long '��ҳID
        str�Һŵ� As String '�Һŵ���
        str�Ա� As String  '�����Ա�
        intӤ�� As Integer
        int���� As Integer 'ҽ�����˵�����
        blnҽ�� As Boolean '�Ƿ�ҽ������
        lng���˿���ID As Long '����ȷ����ҩ�䷽��ȱʡҩ��
End Type

Private mfrmPacsMain As frmPacsMain
Private mfrmVideoMain As frmPacsMain
Private mfrmPatholMain As frmPacsMain
Private mobjPacsInterface As Object

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
On Error GoTo errHandle
    
    Set gcnOracle = cnMain
    Set gfrmMain = frmMain
    gstrDBUser = strDBUser
    glngModul = lngModul
    glngSys = lngSys

    '��ȡע������ж��Ƿ����õ�����־
    gblnUseDebugLog = IIf(Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPacsMain", "UseDebugLog", 0)) = 0, False, True)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPacsMain", "UseDebugLog", IIf(gblnUseDebugLog, 1, 0))
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub
    
    Call WriteLog("CodeMan -> Step 1")
    
    If Not InitSysPar Then Exit Sub '��ʼ��ϵͳ����
    
    Select Case lngModul
        Case 1049, 2280, 2281, 2282, 2283, 2284, 2286
            If mobjPacsInterface Is Nothing Then Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS���ܱ���༭���ӿڲ���zlPacsInterfaceC")
            If Not mobjPacsInterface Is Nothing Then Call mobjPacsInterface.CodeMain(glngSys, lngModul, "", frmMain.hWnd, UCase(gcnOracle.Properties(23)), gcnOracle.Properties(24), UCase(Split(gcnOracle.Properties(8), "=")(2)))
            
        Case 1056 'Ӱ������Ŀ
            frmRadLists.Show , gfrmMain
        Case 1057 'Ӱ���豸Ŀ¼
            frmPACSDevice.Show , gfrmMain
        Case 1079
            frmPatholSpecimenCfg.Show , gfrmMain
        Case 1290 'Ӱ��ʦ����վ
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
            Set gobjEmr = frmMain.mobjEMR
            
            Call InitPacsControlWithInPro(gcnOracle, frmMain.hWnd, lngSys, lngModul, frmMain, gblnUseDebugLog)
            
            gstrOracleOwner = zlDatabase.GetPara("XWOracleӵ����", glngSys, G_LNG_XWPACSVIEW_MODULE, "zlhis")
            
            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = New frmPacsMain
            End If
            
            mfrmPacsMain.ShowStation lngModul, gfrmMain
            mfrmPacsMain.Caption = "Ӱ��ҽ������վ"
        
            Set gfrmPacsMain = mfrmPacsMain
        Case 1291   'Ӱ��ɼ�����վ
            Call WriteLog("CodeMan -> Step 1.1������Ӱ��ɼ�ģ�����̡�")

            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
            Set gobjEmr = frmMain.mobjEMR
            
            Call InitPacsControlWithInPro(gcnOracle, frmMain.hWnd, lngSys, lngModul, frmMain, gblnUseDebugLog)
            
            Call WriteLog("CodeMan -> Step 1.2������CISKernel��RICHEPR��ʼ����")
            
            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = New frmPacsMain
            End If
                        
            mfrmVideoMain.ShowStation lngModul, gfrmMain
            mfrmVideoMain.Caption = "Ӱ��ɼ�����վ"
            
            Call WriteLog("CodeMan -> Step 1.3������Ӱ��ɼ���������ʾ���á�")
            
        Case 1292   'Ӱ�����̿���
            frmWorkFlow.Show , gfrmMain
        Case 1294   'Ӱ��������վ
        
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
            Set gobjEmr = frmMain.mobjEMR
            
            Call InitPacsControlWithInPro(gcnOracle, frmMain.hWnd, lngSys, lngModul, frmMain, gblnUseDebugLog)
            
            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = New frmPacsMain
            End If
                        
            mfrmPatholMain.ShowStation lngModul, gfrmMain
            mfrmPatholMain.Caption = "������鹤��վ"
            
        Case 1295   '�����鵵����վ
            frmPatholArchivesStation.Show , gfrmMain
        Case 1296   '�����軹����վ
            frmPatholReborrowStation.Show , gfrmMain
        Case 1297   '����������ʧ��������վ
            frmPatholLoseStation.Show , gfrmMain
    End Select
    
    Call WriteLog("CodeMan -> Step 2")
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
On Error GoTo errHandle
    Set gcnOracle = cnMain
    Set gfrmMain = Nothing
    gstrDBUser = strDBUser
    glngModul = lngModul
    glngSys = lngSys

    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub

    If Not InitSysPar Then Exit Sub '��ʼ��ϵͳ����
    
    '��ȡע������ж��Ƿ����õ�����־
    gblnUseDebugLog = IIf(Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPacsMain", "UseDebugLog", 0)) = 0, False, True)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPacsMain", "UseDebugLog", IIf(gblnUseDebugLog, 1, 0))
        

    Select Case lngModul
        Case 1049, 2280, 2281, 2282, 2283, 2284, 2286
            If mobjPacsInterface Is Nothing Then Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS���ܱ���༭���ӿڲ���zlPacsInterfaceC")
            If Not mobjPacsInterface Is Nothing Then Call mobjPacsInterface.CodeMain(glngSys, lngModul, "", lngMain, UCase(gcnOracle.Properties(23)), gcnOracle.Properties(24), UCase(Split(gcnOracle.Properties(8), "=")(2)))
            
        Case 1056 'Ӱ������Ŀ
            zlCommFun.ShowChildWindow frmRadLists.hWnd, lngMain
        Case 1057 'Ӱ���豸Ŀ¼
            zlCommFun.ShowChildWindow frmPACSDevice.hWnd, lngMain
        Case 1079 'Ӱ���豸Ŀ¼
            zlCommFun.ShowChildWindow frmPatholSpecimenCfg.hWnd, lngMain
        Case 1290 'Ӱ��ʦ����վ
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, False)
            
            Set gobjEmr = objLogin.mobjEMR
            
            Call InitPacsControlWithInPro(gcnOracle, lngMain, lngSys, lngModul, Nothing, gblnUseDebugLog)
            
            gstrOracleOwner = zlDatabase.GetPara("XWOracleӵ����", glngSys, G_LNG_XWPACSVIEW_MODULE, "zlhis")
                        
            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = New frmPacsMain
            End If
            
            zlCommFun.ShowChildWindow mfrmPacsMain.hWnd, lngMain, True
            
            mfrmPacsMain.ShowStation lngModul, gfrmMain     '��ʱgfrmMainΪnothing
            mfrmPacsMain.Caption = "Ӱ��ҽ������վ"
            
            Set gfrmPacsMain = mfrmPacsMain
        Case 1291   'Ӱ��ɼ�����վ
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, False)
            
            Set gobjEmr = objLogin.mobjEMR
            
            Call InitPacsControlWithInPro(gcnOracle, lngMain, lngSys, lngModul, Nothing, gblnUseDebugLog)
            
            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = New frmPacsMain
            End If
            
            zlCommFun.ShowChildWindow mfrmVideoMain.hWnd, lngMain, True
            
            mfrmVideoMain.ShowStation lngModul, gfrmMain    '��ʱgfrmMainΪnothing
            mfrmVideoMain.Caption = "Ӱ��ɼ�����վ"
            
        Case 1292   'Ӱ�����̿���
            zlCommFun.ShowChildWindow frmWorkFlow.hWnd, lngMain
        Case 1294   'Ӱ��������վ
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, False)
            
            Set gobjEmr = objLogin.mobjEMR
            
            Call InitPacsControlWithInPro(gcnOracle, lngMain, lngSys, lngModul, Nothing, gblnUseDebugLog)
            
            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = New frmPacsMain
            End If
            
            zlCommFun.ShowChildWindow mfrmPatholMain.hWnd, lngMain, True
            
            mfrmPatholMain.ShowStation lngModul, gfrmMain   '��ʱgfrmMainΪnothing
            mfrmPatholMain.Caption = "������鹤��վ"
            
        Case 1295   '�����鵵����վ
            zlCommFun.ShowChildWindow frmPatholArchivesStation.hWnd, lngMain, True
        Case 1296   '�����軹����վ
            zlCommFun.ShowChildWindow frmPatholReborrowStation.hWnd, lngMain, True
        Case 1297   '����������ʧ��������վ
            zlCommFun.ShowChildWindow frmPatholLoseStation.hWnd, lngMain, True
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
On Error GoTo errHandle
    
    Dim frmThis As Form
    
    For Each frmThis In Forms
        UnLoadForm frmThis
    Next
    
    Call gobjEmr.CloseForms
    CloseWindows = (Forms.Count = 0)

errHandle:
End Function

Private Sub UnLoadForm(objForm As Object)
On Error GoTo errHandle
    Unload objForm
Exit Sub
errHandle:
    Debug.Print err.Description
End Sub

Private Sub Class_Initialize()
    VERSION = App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
On Error GoTo errHandle
    Unload frmPacsMain
    Set frmPacsMain = Nothing
    Set mobjPacsInterface = Nothing
errHandle:
End Sub