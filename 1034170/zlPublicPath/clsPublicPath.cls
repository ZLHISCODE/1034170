VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public objDockPath As zlPublicPath.clsDockPath

Public Function InitCommon(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, Optional ByVal strվ�� As String, Optional ByRef frmMain As Object, _
        Optional ByVal lngModule As Long, Optional ByVal MainPrivs As String, Optional ByRef objEmr As Object) As Boolean
'���ܣ���ʼ��ҽ����������
'������MainPrivs=������Ȩ�޴�
'     frmMain =�����洰�����
'     mlngModule=������ģ���
'     objMip=��Ϣƽ̨����
    Set gcnOracle = cnOracle
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    glngSys = lngSys
    glngModule = lngModule
    gMainPrivs = MainPrivs
    gstrNodeNo = strվ��

    On Error Resume Next
    If gobjPath Is Nothing Then
        Set gobjPath = New zlCISPath.clsCISPath
        Call gobjPath.InitCISPath(cnOracle, frmMain, lngSys, MainPrivs, objEmr)
    End If
    
    InitCommon = True
End Function

Private Sub Class_Initialize()
    If objDockPath Is Nothing Then Set objDockPath = New zlPublicPath.clsDockPath
End Sub

Public Sub zlExecuteCode(ByVal lngSys As Long, ByVal lngModul As Long, ByVal frmParent As Variant, ByVal strPrivs As String)
    Call gobjPath.zlExecuteCode(lngSys, lngModul, frmParent, strPrivs)
End Sub

Public Function CloseWindows() As Boolean
    If gobjPath Is Nothing Then CloseWindows = True: Exit Function
    CloseWindows = gobjPath.CloseWindows
End Function