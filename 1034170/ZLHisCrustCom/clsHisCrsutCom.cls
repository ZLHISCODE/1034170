VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHisCrsutCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mobjHisCrustCom     As Object
Private mpbSend             As PropertyBag

Public Function InitHisCrustCom(ByRef objHisCrust As Object, ByVal strAppPath As String, ByVal strCommand As String, ByVal strLogFile As String) As Boolean
'���ܣ���ʼ������
'objHisCrust=�Զ�������ƵĴ����������������Զ�����������һЩ���顣
'strAppPath=�Զ���������·��
'strCommand=�Զ�����������
'�Ƿ��ʼ���ɹ�
    Set mobjHisCrustCom = objHisCrust
    gstrAppPath = strAppPath
    gstrCommand = strCommand
    gobjTrace.LogFile = strLogFile
    gobjTrace.WriteSection "��ʼʹ��ZLHISCRUSTCOM.DLL����", SL_LevelThree
    InitHisCrustCom = InitCommon
    gobjTrace.WriteInfo "InitHisCrustCom", "END", True
End Function

Public Function UpdateBaseFile(Optional ByVal blnUpgrade As Boolean = True) As Boolean
    On Error GoTo errH
    If gotCurType <> OT_CheckFile Then
        If Not UpgradeBase() Then
            Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:����ʧ�� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
            Call SetOperateProcess(gotCurType, OS_Failure, SumErrMsg)          '��ʶ��������
            GoTo ReCall                      'ǿ���˳�����
        End If
    End If
    gobjTrace.WriteInfo "UpdateBaseFile", "END", True
    Call gobjTrace.CloseLog
    UpdateBaseFile = True
    Exit Function
errH:
    MsgBox Err.Description, vbInformation, App.Title
    Err.Clear
ReCall:
    Call CallHISEXE
    Call gobjMe.ExitApp
    Exit Function
AutoLogin:
    Call CallHISEXE(True)
    Call gobjMe.ExitApp
    Exit Function
End Function

Public Function UpdateFile() As Boolean
    On Error GoTo errH
    gobjTrace.LogFile = gobjTrace.LogFile
    gobjTrace.WriteInfo "UpdateFile", "BGN", True
    If Not GetUpgradeFileList Then
        Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:����ʧ�� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, IIf(gotCurType = OT_CheckFile, OS_Failure, OS_NotInProcessing), SumErrMsg)          '��ʶ��������
        GoTo ReCall
    End If
    If grsFileUpgrade.RecordCount = 0 Then
        Call RecordErrMsg(MT_InitEnv, "�ļ��嵥��ȡ", "û�п��������ļ���ϵͳ�Զ��˳���")
        Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:������� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, OS_Completed, SumErrMsg, glngFileBatch)          '��ʶ��������
        GoTo ReCall
    End If
    Call GetKILLProcess
    Call frmHisCrust.tmrStart_Timer
    gobjTrace.WriteInfo "UpdateFile", "END", True
    Call gobjTrace.CloseLog
    UpdateFile = True
    Exit Function
errH:
    MsgBox Err.Description, vbInformation, App.Title
    Err.Clear
ReCall:
    Call CallHISEXE
    Call gobjMe.ExitApp
    Exit Function
AutoLogin:
    Call CallHISEXE(True)
    Call gobjMe.ExitApp
    Exit Function
End Function

Public Sub ExitApp()
'���ܣ�����������������ֹ��������
    Call mobjHisCrustCom.ExitApp
End Sub

Public Function SendCommand(ByVal strCommand As String, ByVal strParas As String, Optional ByRef pbSend As PropertyBag, Optional ByRef pbReturn As PropertyBag) As Boolean
'���ܣ���������������ִ������,���ֽӿڼ���
'strCommand:ִ�е�����
'strParas:ִ������Ĳ���
'pbSend�����Ӵ�������
'pbReturn:���ӷ�������
'���أ������Ƿ�ִ�гɹ�
    If pbSend Is Nothing Then
        SendCommand = mobjHisCrustCom.ReceiveCommand(strCommand, strParas, mpbSend, pbReturn)
        Set mpbSend = Nothing
    Else
        SendCommand = mobjHisCrustCom.ReceiveCommand(strCommand, strParas, pbSend, pbReturn)
    End If
    
End Function

Public Function ReceiveCommand(ByVal strCommand As String, Optional ByVal strParas As String, Optional ByRef pbSend As PropertyBag, Optional ByRef pbReturn As PropertyBag) As Boolean
'���ܣ����������DLLִ�����������չ,���ֽӿڼ���
'strCommand:ִ�е�����
'strParas:ִ������Ĳ���
'pbSend�����Ӵ�������
'pbReturn:���ӷ�������
'���أ������Ƿ�ִ�гɹ�
'    Dim pbReturn As New PropertyBag
'
'
'    pbReturn.WriteProperty "���ص�����1", "���ص�ֵ1"
'    pbReturn.WriteProperty "���ص�����2", "���ص�ֵ2"

End Function


Public Sub AddSendValue(ByVal strKeyName As String, varValue As Variant)
'���ܣ����������뵽���ݴ��Ͱ�
'strKeyName:���ݹؼ���
'varValue:��������
    If mpbSend Is Nothing Then
        Set mpbSend = New PropertyBag
    End If
    mpbSend.WriteProperty strKeyName, varValue
End Sub

Private Sub Class_Initialize()
    Set gobjMe = Me
End Sub

Private Sub Class_Terminate()
    Call gobjTrace.CloseLog
    Set gobjTrace = Nothing
    Set gobjMe = Nothing
    Set grsFileUpgrade = Nothing
    Set gcnOracle = Nothing
    Set gobjFSO = Nothing
    Set gcllSetPath = Nothing
    Set gclsRegCom = Nothing
    Set grsErrRec = Nothing
    Set gobj7zZip = Nothing
    Set gclsConnect = Nothing
End Sub