VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInOutMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' strTag=������Ϣ�����ڴ洢���ﲡ����Ƭ�ļ���·�����Ժ���չʱ����|�ָ�
Public Event Closed(ByVal blnEditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String) '��ҳ�ر��¼�

Private mfrmInMedRecEdit As Object
Private mfrmOutMedRecEdit As frmOutMedRecEdit
Private mclsProperty As clsProperty                   '�����࣬�������沿�ֱ�������������
'������
Public Property Get Pros() As clsProperty
    Set Pros = mclsProperty
End Property
'��ҳ�Ƿ��Ǵ�״̬��ֻ������
Public Property Get IsOpen() As Boolean
    IsOpen = mclsProperty.IsOpen
End Property
'��ҳ�Ƿ�ϲ���,ֻ������
Public Property Get IsMerged() As Boolean
    IsMerged = mclsProperty.IsMerged
End Property
'��Ժ��ҳ����Ƿ����ı�,ֻ������
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsProperty.IsDiagChange
End Property
'IsDiagInput:ָ����ϵ���ҳ����Ƿ�����
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsProperty.IsDiagInput
End Property

Public Sub InitMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule, Optional objCommunity As Object, Optional objInsure As clsInsure)
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    If lngModule = P����������� Then Exit Sub
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(fҽ����ҳ, lngSys, lngModule, objMip, objCommunity, objInsure)
End Sub

Public Function ShowInMedRecEdit(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lngPathState As Long, _
                            Optional ByVal str���� As String, Optional ByVal strPrivs As String, Optional ByVal intOpenMode As Integer, Optional ByVal blnModal As Boolean) As Boolean
'������
'      frmParent=���ô���
'      lngModule=����ģ��
'      lng����ID=��ǰ����ID
'      lng��ҳID=��ǰ���˵���ҳID
'      lng����ID=��ǰ���˵Ŀ���ID
'      lngPathState=��ǰ���˵�·��״̬ -1=δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
'      str����=Ҫʾ¼���������ͣ���"3,13"��ʽ
'      strPrivs=��ǰ�û�Ȩ��
'      intOpenMode=��ҳ��ģʽ
'      blnModal=����򿪷�ʽ��True-ģʽ�򿪴��壨����ֻ�йرղ��ܲ����������壩��False-��ģ�򿪴��壨����δ�ر�Ҳ���Բ����������壩
'���أ�ShowInMedRecEdit=�Ƿ���ȷ����ť�رմ���
    If gcolclsMain Is Nothing Then
        Set gcolclsMain = New Collection
    End If
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
        On Error Resume Next
        If Not gcolclsMain("סԺ��ҳ") Is Nothing Then
            gcolclsMain.Remove ("סԺ��ҳ")
            gcolclsMain.Add Me, "סԺ��ҳ"
        Else
            gcolclsMain.Add Me, "סԺ��ҳ"
        End If
        If Err.Number <> 0 Then On Error GoTo 0
    End If
    If gcolclsPros Is Nothing Then
        Set gcolclsPros = New Collection
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    mclsProperty.PatiType = PF_סԺ
    mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("������ҳ��׼", mclsProperty.SysNo, pסԺҽ��վ, "0")) + 1
    '��ֹ������ʱж�ز���
    If Not mclsProperty.CurrentForm Is Nothing Then
        On Error Resume Next
        Unload mclsProperty.CurrentForm
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
    mclsProperty.Privs = "" '�Զ���ȡ
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.OpenMode = intOpenMode
    mclsProperty.����ID = lng����ID
    mclsProperty.��ҳID = lng��ҳID
    mclsProperty.��Ժ����ID = lng����ID
    mclsProperty.PathState = lngPathState
    mclsProperty.MustDiagType = str����
    mclsProperty.InfosChange = False
    mclsProperty.LoadFinish = False
    
    ShowInMedRecEdit = frmMain.ShowMe(blnModal)
    On Error Resume Next
    If Not gcolclsPros("" & TypeName(gclsPros.CurrentForm)) Is Nothing Then
        gcolclsPros.Remove ("" & TypeName(gclsPros.CurrentForm))
        gcolclsPros.Add mclsProperty, "" & TypeName(gclsPros.CurrentForm)
    Else
        gcolclsPros.Add mclsProperty, "" & TypeName(gclsPros.CurrentForm)
    End If
    If Err.Number <> 0 Then On Error GoTo 0
    '����
    If blnModal Then
        ShowInMedRecEdit = mclsProperty.IsOK
    End If
End Function

Public Function ShowOutMedRecEdit(frmParent As Object, ByVal str�Һŵ� As String, ByVal strPrivs As String, Optional ByVal intOpenMode As Integer, Optional ByRef strPictureFile As String) As Boolean
'������blnDiagnose=�Ƿ����������д���
'���أ�blnDiagnose=�Ƿ���д�˲��˵����
    Dim i As Long
    
    If gcolclsMain Is Nothing Then
        Set gcolclsMain = New Collection
    End If
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
        On Error Resume Next
        If Not gcolclsMain("������ҳ") Is Nothing Then
            gcolclsMain.Remove ("������ҳ")
            gcolclsMain.Add Me, "������ҳ"
        Else
            gcolclsMain.Add Me, "������ҳ"
        End If
        If Err.Number <> 0 Then On Error GoTo 0
    End If
    If gcolclsPros Is Nothing Then
        Set gcolclsPros = New Collection
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    '��ֹ������ʱж�ز���
    If Not mclsProperty.CurrentForm Is Nothing Then
        On Error Resume Next
        Unload mclsProperty.CurrentForm
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
    Set mfrmOutMedRecEdit = New frmOutMedRecEdit
    Set mclsProperty.CurrentForm = mfrmOutMedRecEdit
    mclsProperty.Privs = "" '�Զ���ȡ
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.PatiType = PF_����
    mclsProperty.����ID = 0
    mclsProperty.��ҳID = 0
    mclsProperty.MedPageSandard = ST_������ҳ
    mclsProperty.OpenMode = intOpenMode
    mclsProperty.RegistNo = str�Һŵ�
    gclsPros.PictureFile = strPictureFile
    ShowOutMedRecEdit = mfrmOutMedRecEdit.ShowMe
    On Error Resume Next
    If Not gcolclsPros("" & TypeName(gclsPros.CurrentForm)) Is Nothing Then
        gcolclsPros.Remove ("" & TypeName(gclsPros.CurrentForm))
        gcolclsPros.Add mclsProperty, "" & TypeName(gclsPros.CurrentForm)
    Else
        gcolclsPros.Add mclsProperty, "" & TypeName(gclsPros.CurrentForm)
    End If
    If Err.Number <> 0 Then On Error GoTo 0
    strPictureFile = gclsPros.PictureFile
End Function


Public Function PrintOrPriviewInMedRec(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByRef objReport As Object, ByVal lng����ID As Long, ByRef objForm As Object, Optional intPage As Integer) As Boolean
'���ܣ���ҳ��ӡ��Ԥ��
'������intType=2����ӡ����=1��Ԥ����0=����
'     mobjReport-��ӡ������lng����ID-���˿��ң�mobjForm-������
'     intPage=1-4��ӡ��ҳ������ʽ��=5��ӡ����+��ҳ1��=6��ӡ����+��ҳ2
    Call PrintInMedRec(intType, lng����ID, lng��ҳID, lng����ID, intPage, objReport, objForm)
    PrintOrPriviewInMedRec = True
End Function
'
Public Sub Closed(ByVal blnEditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String)
'���ܣ������¼�
' strTag=������Ϣ�����ڴ洢���ﲡ����Ƭ�ļ���·�����Ժ���չʱ����|�ָ�
    RaiseEvent Closed(blnEditCancel, str����ID, str���ID, strTag)
End Sub

Public Sub EditDelivery(ByRef objFrmMain As Object, ByVal lngPatiID As Long, ByVal lngMainID As Long)
'����:�����������Ǽ�ʱ�ǼǷ�����Ϣ
    Call CmdDeliceryInfoClick(1, objFrmMain, lngPatiID, lngMainID)
End Sub

Public Sub Hide()
    On Error Resume Next
    Call frmMain.Hide
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Public Function FormUnLoad() As Boolean
    On Error Resume Next
    Unload frmMain
    If Not gblnUnload Then
        FormUnLoad = True
    Else
        FormUnLoad = False
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Public Sub PlugWndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    '��Ҹ�ҳ�Զ������Ϣ��������
    Dim wzDelta As Integer
    Select Case msg
        Case WM_MOUSEWHEEL   '����
            wzDelta = HIWORD(wParam)
            With gclsPros.CurrentForm
                If wzDelta > 0 Then        '���Ϲ���
                    Call ChangePage(False, , , False)
                Else                        '���¹���
                    Call ChangePage(True, , , False)
                End If
            End With
    End Select
End Sub

