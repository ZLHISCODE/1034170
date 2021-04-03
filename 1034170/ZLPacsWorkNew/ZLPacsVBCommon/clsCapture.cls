VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjWork_ImageCap As Object

Property Get HasVideo() As Boolean
'判断该对象是否拥有视频采集
     If Not mobjWork_ImageCap Is Nothing Then
        HasVideo = mobjWork_ImageCap.HasVideo
    End If
End Property

Public Function ContainerHwnd() As Long
    If Not mobjWork_ImageCap Is Nothing Then
        ContainerHwnd = mobjWork_ImageCap.ContainerHwnd
    End If
End Function

Public Function ContainerObj() As Object
    If Not mobjWork_ImageCap Is Nothing Then
        Set ContainerObj = mobjWork_ImageCap.ContainerObj
    End If
End Function

Public Function VideoTitle() As String
    If Not mobjWork_ImageCap Is Nothing Then
        VideoTitle = mobjWork_ImageCap.VideoTitle
    End If
End Function

Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long)
    Dim lngSendNO As Long
    Dim lngStudyState As Long
    Dim blnMoved As Boolean
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If Not mobjWork_ImageCap Is Nothing Then
        strSql = "Select 发送号,执行过程 From 病人医嘱发送 Where 医嘱ID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID)
        
        If rsData.RecordCount > 0 Then
            lngSendNO = Nvl(rsData!发送号, 0)
            lngStudyState = Nvl(rsData!执行过程, 0)
        End If
        
        Call mobjWork_ImageCap.zlUpdateStudyInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
    End If
End Sub

Public Sub zlRefreshData()
    If Not mobjWork_ImageCap Is Nothing Then Call mobjWork_ImageCap.zlRefreshData
End Sub

Public Sub zlRefreshVideoWindow(Optional blnSingleWindow As Boolean = False)
    If Not mobjWork_ImageCap Is Nothing Then
        mobjWork_ImageCap.ParentWindowKey = "采集" & IIf(blnSingleWindow = True, "Dock", "")
        mobjWork_ImageCap.zlRefreshVideoWindow
    End If
End Sub

Public Sub InitActiveVideoModuleObj(objDbAccess As Object, _
                                    ByVal strDbUser As String, _
                                    ByVal lngSysNo As Long, _
                                    ByVal lngCallModule As Long, _
                                    ByVal lngCurDepartId As Long, _
                                    Optional blnForceRefresh As Boolean = False, _
                                    Optional blnOpenDebugMode As Boolean = False)
'初始化ActivexExe视频采集模块对象
    Dim strPrivs As String
    
    If mobjWork_ImageCap Is Nothing Then
        Set mobjWork_ImageCap = CreateObject("zl9PacsImageCap.clsPacsCapture")
        
        Call InitComLib(objDbAccess, strDbUser)
        
        strPrivs = ";" & gobjComlib.GetPrivFunc(lngSysNo, lngCallModule) & ";"
        
        Call mobjWork_ImageCap.zlInitModule(objDbAccess, lngSysNo, lngCallModule, strPrivs, lngCurDepartId, 0, Nothing, blnForceRefresh, blnOpenDebugMode)
    End If
End Sub

Public Sub ClassTerminate()
    If Not mobjWork_ImageCap Is Nothing Then
        Set mobjWork_ImageCap = Nothing
    End If
End Sub
