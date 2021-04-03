VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPageMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'���ò������λ��,���ȷ��ʱ����
Public Event MedRecSaveLocation(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'������ҳ��ˢ�²����б�
Public Event SavePage(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
Private mfrmPageMedRecEdit As Object
Private mclsProperty As clsProperty                   '�����࣬�������沿�ֱ�������������
'��ҳ�Ƿ��Ǵ�״̬��ֻ������
Public Property Get IsOpen() As Boolean
    IsOpen = mclsProperty.IsOpen
End Property
'��ҳ�Ƿ�ϲ���,ֻ������
Public Property Get IsMerged() As Boolean
    IsMerged = mclsProperty.IsMerged
End Property
'��Ժ��Ҫ����Ƿ����ı�,ֻ������
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsProperty.IsDiagChange
End Property
'IsDiagInput:ָ����ϵ���ҳ����Ƿ�����
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsProperty.IsDiagInput
End Property

Public Sub InitPageMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule)
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(f������ҳ, lngSys, lngModule, objMip)
End Sub

Public Function ShowPageMedRecEdit(frmParent As Object, ByVal strסԺ�� As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal intOpenMode As Integer, Optional bln��Ŀ As Boolean, Optional blnChange����Ա As Boolean) As Boolean
'������lngModel=����ģ��
'      frmParent=���ô���
'      lng����ID=��ǰ����ID
'      lng��ҳID=��ǰ���˵���ҳID
'      strסԺ��=��ǰ���˵�סԺ��
'      bln��Ŀ=�����е���ҳ����
'      intOpenMode=��ҳ��ģʽ
'      blnChange����Ա=�Ƿ���Ը�������Ա
'���أ�ShowInMedRecEdit=�Ƿ���ȷ����ť�رմ���
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
    mclsProperty.PatiType = PF_סԺ
    If Not mclsProperty.ShareMedRec Then
        mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("������ҳ��׼", mclsProperty.SysNo, p��������, "0")) + 1
    Else
        mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("������ҳ��׼", 100, pסԺҽ��վ, "0")) + 1
    End If
    If gblnUnload Then
        Unload frmMain
    Else
        If Not mclsProperty.CurrentForm Is Nothing Then
            '82639:�������ԭ��δ֪��������ٷ��֣�ĳ�ιرմ��壬�ú���ִ�н���������Form_Load���ݣ����´���δж�ص����ٴ�showʱ
            'ֱ������Form_load����ʾ����һ�����˵���Ϣ���˴�Ӧ�Բ����ǣ�ǿ��ж�ش���
            On Error Resume Next
            Unload mclsProperty.CurrentForm
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    End If
    
    mclsProperty.Privs = "" '�Զ���ȡ
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.OpenMode = intOpenMode
    If intOpenMode = EM_�༭ Then
        mclsProperty.IsExistPati = True
    End If
    mclsProperty.����ID = lng����ID
    mclsProperty.��ҳID = lng��ҳID
    mclsProperty.InNo = strסԺ��
    mclsProperty.Is��Ŀ = bln��Ŀ
    mclsProperty.Change����Ա = blnChange����Ա
    ShowPageMedRecEdit = frmMain.ShowMe(False)
    On Error Resume Next
    If Not gcolclsPros("" & TypeName(gclsPros.CurrentForm)) Is Nothing Then
        gcolclsPros.Remove ("" & TypeName(gclsPros.CurrentForm))
        gcolclsPros.Add mclsProperty, "" & TypeName(gclsPros.CurrentForm)
    Else
        gcolclsPros.Add mclsProperty, "" & TypeName(gclsPros.CurrentForm)
    End If
    If Err.Number <> 0 Then On Error GoTo 0
End Function

Private Sub Class_Terminate()
    Set mclsProperty = Nothing
    Set mfrmPageMedRecEdit = Nothing
End Sub

Public Sub MedRecSaveLocation(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    RaiseEvent MedRecSaveLocation(lng����ID, lng��ҳID)
End Sub

Public Sub SavePage(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    RaiseEvent SavePage(lng����ID, lng��ҳID)
End Sub
Public Function unLoadPage() As Boolean
    Unload frmMain
    If Not gblnUnload Then
        unLoadPage = True
    Else
        unLoadPage = False
    End If
End Function