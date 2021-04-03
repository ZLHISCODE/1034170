Attribute VB_Name = "mdlSampleReprot"
Option Explicit

Public gstrSysName As String                        'ϵͳ����
Public gstrProductName As String                    'OEM��Ʒ����
Public gstrUnitName As String                       '�û���λ����
Public gcnOracle As New ADODB.Connection                 '�������ݿ�����

Public UserInfo As TYPE_USER_INFO

'�û���Ϣ
Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String '��Ա����
    ���� As String
    DeptID As Long '����ID
    DeptNo As String '���ű��
    DeptName As String '��������
    DBUser As String '���ݿ��û�
End Type

Public glngSys As Long                              'ϵͳ��
Public glngModule As Long                           'ģ���
Public gobjLISInsideComm As Object
Public gobjComLib As Object

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.DeptID = Nvl(rsTmp!����ID, 0)
            UserInfo.DeptNo = rsTmp!������ & ""
            UserInfo.DeptName = rsTmp!������ & ""
            UserInfo.DBUser = rsTmp!�û��� & ""
            GetUserInfo = True
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub InitObjLis()
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLISInsideComm Is Nothing Then
        On Error Resume Next
        Set gobjLISInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLISInsideComm Is Nothing Then
            If gobjLISInsideComm.InitComponentsHIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLISInsideComm = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

