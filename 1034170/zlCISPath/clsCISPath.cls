VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InitCISPath(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal MainPrivs As String, Optional ByRef objEmr As Object)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    Set gobjEmr = objEmr

    glngSys = lngSys
    gMainPrivs = MainPrivs
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    
    '��ʼ��
    Call GetUserInfo
    Call InitSysPar
    
    'clsCISJob���ѵ���gobjKernel.InitCISKernel
End Sub

Public Sub zlExecuteCode(ByVal lngSys As Long, ByVal lngModul As Long, ByVal frmParent As Variant, ByVal strPrivs As String)
    glngSys = lngSys: glngModul = lngModul
    
    If strPrivs = "" Then
        gstrPrivs = GetPrivFunc(glngSys, glngModul)
    Else
        gstrPrivs = strPrivs
    End If
    
    Select Case lngModul
    Case 1078 '�ٴ�·������
        If IsObject(frmParent) Then
            frmPathManage.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmPathManage.Hwnd, frmParent
        End If
    Case 1275 '�ٴ�·������
        If IsObject(frmParent) Then
            frmPathTrack.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmPathTrack.Hwnd, frmParent
        End If
    End Select
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = gobjKernel.CloseWindows And (Forms.count = 0)
End Function
