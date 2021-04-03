VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPartogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
  '
End Sub

Private Sub Class_Terminate()
    Set mclsUnzip = Nothing
    Set mclsZip = Nothing
End Sub


Public Sub InitPartogram(cnMain As ADODB.Connection, ByVal lngSys As Long)
    Set gcnOracle = cnMain
    Call InitCommon(gcnOracle)
        
    glngSys = lngSys
    gstrDbOwner = GetDbOwner(glngSys)
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("��Ʒ����"), Default:="����")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    
    Call GetUserInfo
End Sub

Public Function GetPartogram() As Object
    Set GetPartogram = frmPartogram
End Function

Public Function GetNewPartogram() As Object
    Set GetNewPartogram = New frmPartogram
End Function

Public Function zlPartogramPara(ByVal frmParent As Object, Optional ByVal strPrivs As String) As Boolean
    zlPartogramPara = frmPartogramPara.ShowPara(frmParent, strPrivs)
End Function

Public Sub zlPrintSet(ByVal frmParent As Object, Optional ByVal bytMode As Byte = 1)
'-----------------------------------------------------------------------------------------------------------------------------
'����:���̴�ӡ����
'-----------------------------------------------------------------------------------------------------------------------------
    Call frmPrintSet.ShowMe(frmParent, bytMode)
End Sub

Public Function zlPartogramEditor(ByVal frmParent As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, ByVal strPrivs As String, _
    Optional ByVal blnEdit As Boolean, Optional ByVal bytSize As Byte = 0) As Boolean
    Dim frmEditor As New frmPartogramEditor
    zlPartogramEditor = frmEditor.ShowMe(frmParent, lngFileID, lngPatiID, lngPageId, lngDeptID, 0, strPrivs, blnEdit, bytSize)
End Function

Public Sub zlPrintData(ByVal frmParent As Object, ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "")
    '----------------------------------------------------------------------------------------------------------------------------
    '����:��ɲ���ͼ��ӡԤ��
    '----------------------------------------------------------------------------------------------------------------------------
    If TypeName(frmParent) <> "frmPartogram" Then
        MsgBox "�봫�봴������ͼʱ��ȡ�Ĵ������", vbInformation, gstrSysName
        Exit Sub
    End If
    Call frmParent.PrintData(bytMode, strPrintDevice)
End Sub

Public Sub PrintPartogram(ByVal objParent As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDtpID As Long, ByVal lngFileIndex As Long, ByVal lngFilePage As Long, Optional ByVal blnPrint As Boolean = True, Optional ByVal strPrintDevice As String = "")
    '�ⲿ�ӿڵ��ã�����ĳ���˲���ͼ��ӡ
    Call ShowPrintPartogram(objParent, lngFileID, lngPatiID, lngPageId, lngDtpID, lngFileIndex, lngFilePage, blnPrint, strPrintDevice)
End Sub