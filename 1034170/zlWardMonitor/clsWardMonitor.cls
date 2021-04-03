VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWardMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function CEC_Initialize Lib "CecDeviceToHis.dll" (ByVal ServerIP As String, ByVal ServerPort As Long, ByVal DispWindowhwnd As Long, ByVal recallFunAddr As Long, ByVal lngobj As Long) As Boolean
Private Declare Function CEC_Uninitialize Lib "CecDeviceToHis.dll" () As Boolean

Private Declare Function CEC_DevNo2His Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngType As Long, ByVal strInPatient As String) As Boolean
Private Declare Function CEC_His2DevNo Lib "CecDeviceToHis.dll" (ByVal lngType As Long, ByVal strInPatient As String) As Long
'lngType:1�໤�Ǵ���, 2HIS����, 3�������

Private Declare Function CEC_ShowWindows Lib "CecDeviceToHis.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
Private Declare Function CEC_SelectBedNo Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long) As Boolean
Private Declare Function CEC_GetListBedNo Lib "CecDeviceToHis.dll" (ByVal strListBedNo As String) As Boolean

Private mblnUsable As Boolean

Public Function Init(ByVal cnOracle As ADODB.Connection, ByRef lngWindowHWND As Long, Optional ByVal lngWindowMode As Long = 3) As Boolean
'���ܣ���ʼ���໤�Ǵ���
'������WindowHWND-��ʾ�໤�ǵĴ��ھ��,WindowMode-��ʽģʽ
'        0-���� 1-��ͨ  2-��󻯣�û�б�����  3-���,�б����� 4-��С�� 5-���ش�λ�����б�,�����˳���ť
    Dim ServerIP As String, ServerPort As String
    
    Call GetServerInfo(ServerIP, ServerPort)
    If ServerPort = "" Or ServerIP = "" Then GoTo errH
    
    On Error GoTo errH
    Call CEC_Initialize(ServerIP, ServerPort, lngWindowHWND, AddressOf RequestData, 0)
    '����ص�����ָ�룬���ڽ�����������
    
    Set gcnOracle = cnOracle
    
    Init = True
    mblnUsable = True
    Exit Function
errH:
    Init = False
    mblnUsable = False
End Function

Public Sub ShowInfor(ByVal strInPatient As String)
'���ܣ���ʾָ�����˵ļ໤��Ϣ
'������strInPatient-סԺ��
    Dim lngDevice As Long
        
    lngDevice = CEC_His2DevNo(3, strInPatient)
    If lngDevice = 0 Then
        Call ShowWindow(False)
    Else
        Call ShowWindow(True)
        Call CEC_SelectBedNo(lngDevice)
    End If
End Sub

Public Sub ShowWindow(ByVal blnShow As Boolean)
'���ܣ���ʾ�����ؼ໤����
'������blnShow:true��ʾ��false����
    Dim lngWindowMode As Long   '0-���� 1-��ͨ  2-��󻯣�û�б�����  3-���,�б����� 4-��С�� 5-���ش�λ�����б�,�����˳���ť
    If blnShow Then
        lngWindowMode = 3
    Else
        lngWindowMode = 0
    End If
    Call CEC_ShowWindows(0, lngWindowMode)
End Sub

Public Sub GetListBedNo(ByRef strListBedNo As String)
'���ܣ���ȡ�������ӵ������ϵ����м໤�Ǵ���
    Dim strList As String * 4000, arrtmp As Variant, i As Long, lngDevice As Long, strPatiNO As String * 20
       
    Call CEC_GetListBedNo(strList)  '���ص����豸���б� '���ظ�ʽ��2|1,2����ǰ���2�������豸������'|'���ٵ�Ϊ���б�
    If InStr(strList, "|") > 1 Then
        arrtmp = Split(strList, "|")
        For i = 1 To Val(arrtmp(0))
            lngDevice = Val(arrtmp(i))
            strPatiNO = ""
            Call CEC_DevNo2His(lngDevice, 3, strPatiNO)
            If Trim(strPatiNO) <> "" Then
                strListBedNo = IIf(strListBedNo = "", "", strListBedNo & ",") & Val(strPatiNO)
            End If
        Next
    End If
End Sub


Private Sub Class_Terminate()
    If mblnUsable Then Call CEC_Uninitialize
End Sub

Public Function ShowPatientInfor(ByVal strInPatient As String) As String
    ShowPatientInfor = GetPatientInfor(strInPatient)
End Function

Private Sub GetServerInfo(ByRef ServerIP As String, ByRef ServerPort As String)
'���ܣ���ȡ�����ļ��еķ�����IP�Ͷ˿�
    Dim objFile As Object
    Dim objText As Object
    Dim strFile As String
    Dim strTmp As String
    
    On Error Resume Next
    Set objFile = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\zlWardMonitor.ini"
    Set objText = objFile.OpenTextFile(strFile, 1) '1-ForReading
    If Err.Number <> 0 Then
        Set objFile = Nothing
        Exit Sub
    End If
    
    strTmp = objText.ReadLine
    If InStr(strTmp, ":") > 0 Then ServerIP = Split(strTmp, ":")(1)
    strTmp = objText.ReadLine
    If InStr(strTmp, ":") > 0 Then ServerPort = Split(strTmp, ":")(1)
    
    If ServerIP = "" Or ServerPort = "" Then
        Call WriteLog("��ȡ�����������ļ�zlWardMonitor.iniʧ��.")
    End If
    objText.Close
    Set objFile = Nothing
End Sub


Public Sub WriteLog(ByVal strInfo As String)
    '��������Ϣд���ļ���
    Dim objFile As Object
    Dim objText As Object
    Dim strFile As String
    
    On Error Resume Next
    Set objFile = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\zlWardMonitor.Log"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, 8) '8-ForAppending
    objText.WriteLine Now()
    objText.WriteLine strInfo
    objText.Close
    Set objFile = Nothing
End Sub