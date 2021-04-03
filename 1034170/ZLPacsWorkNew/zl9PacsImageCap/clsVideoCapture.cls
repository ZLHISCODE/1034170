VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideoCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************************************************
'*
'*  ���ܣ��ɼ���װ���󣨶�vfw��wdm�Լ�������Ƶ�������ͽ���֧�֣�
'*  ����ʱ�䣺2010-01-19
'*  �����ˣ�TJH
'*
'*************************************************************************************************************





Option Explicit






'��Ƶ��������
Public Enum TVideoDriverType
  vdtWDM = 0
  vdtVFW = 1
  vdtTWAIN = 2
  vdtCustom = 3
  '������Ҫ֧�ֵ���������......
End Enum









Private mVideoDriverType As TVideoDriverType '���浱ǰҪʹ�õ���������

Private mVfwWindow As PictureBox  'vfw����Ƶ��ʾ����
Private mVfwControl As clsVfwCapture 'vfw����Ƶ���ƶ���

Private mWdmControl As ZLDSVideoProcess.DSCapture  'wdm����Ƶ�ɼ����

Private mCustomControl As Object    'ר����Ƶ�ɼ����

Private mblnIsAllowStartupVideo As Boolean '�Ƿ�����������Ƶ����ΪFALSEʱ���ö�������з���������ʹ��







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'�������Զ���
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get VideoDriverType() As TVideoDriverType
  VideoDriverType = mVideoDriverType
End Property

Property Get Visible() As Boolean
  Visible = mWdmControl.Visible Or mVfwWindow.Visible
End Property

Property Let WdmVisible(value As Boolean)
  mWdmControl.Visible = True
End Property

Property Let VfwVisible(value As Boolean)
  mVfwWindow.Visible = value
End Property

Property Let VideoDriverType(value As TVideoDriverType)
  mVideoDriverType = value
  
  '���òɼ��������ʾ״̬
  Select Case value
    Case vdtWDM
      mWdmControl.Visible = True
      mVfwWindow.Visible = False
    Case vdtVFW
      mVfwWindow.Visible = True
      mWdmControl.Visible = False
    'Case ...
  End Select
End Property

'�жϲɼ�����Ƿ�����
Property Get IsStartup() As Boolean
  Select Case mVideoDriverType
    Case vdtWDM
      IsStartup = mWdmControl.PreviewState
    Case vdtVFW
      IsStartup = mVfwControl.hWnd <> 0
    Case Else
      IsStartup = False
  End Select
End Property

Property Get IsAllowStartupVideo() As Boolean
  IsAllowStartupVideo = mblnIsAllowStartupVideo
End Property

Property Let IsAllowStartupVideo(value As Boolean)
  mblnIsAllowStartupVideo = value
End Property


Property Get VideoSize() As TVideoSize

  'ȡ����Ƶ����
  Select Case mVideoDriverType
    Case vdtVFW
      Dim vfwParameter As clsVfwParameterCfg
      
      Set vfwParameter = mVfwControl.GetCaptureParameter()
      
      VideoSize.Width = vfwParameter.CapVideoWidth
      VideoSize.Height = vfwParameter.CapVideoHeight
      
    Case vdtWDM
      Dim wdmParameter As TCaptureParameter
      Dim wdmParameterEnum As DSCapParameterEnum
      Dim wdmVideoSize As TVideoSize
            
      Set wdmParameterEnum = New DSCapParameterEnum
            
      Call mWdmControl.GetCaptureParameter(wdmParameter)
      
      VideoSize = wdmParameterEnum.VideoSizeConvert(wdmParameter.VideoSize)
    Case Else
      VideoSize.Width = 800
      VideoSize.Height = 600
  End Select
  
End Property












''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'���󷽷�����
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'��ʼ��ƵԤ��
Public Sub StartPreview()

  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  Select Case mVideoDriverType
    Case vdtWDM 'wdm����Ƶ�ɼ�
      Call mWdmControl.StartPreview
      Call mWdmControl.RefreshWindow
    Case vdtVFW 'vfw����Ƶ�ɼ�
      Call mVfwControl.StartPreview(mVfwWindow.hWnd)
      Call mVfwControl.UpdateCaptureWindowPos(mVfwWindow.ScaleWidth, mVfwWindow.ScaleHeight)
    'Case ...
  End Select
  
End Sub


'ֹͣ��ƵԤ��
Public Sub StopPreview()

  Select Case mVideoDriverType
    Case vdtWDM 'wdm����Ƶ�ɼ�
      Call mWdmControl.StopPreview
    Case vdtVFW 'vfw����Ƶ�ɼ�
      Call mVfwControl.StopPreview
    'Case   ...
  End Select
  
End Sub


'ȫ����ʾ
Public Sub FullScreen(owner As Object, ByVal ownerHandle As Long)
  
  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  Select Case mVideoDriverType
    Case vdtWDM 'wdm����Ƶ�ɼ�
      mWdmControl.ShowModel = smFit
      
      Call mWdmControl.ShowFullScreen(0, GetMonitorIndex(ownerHandle))
      
      mWdmControl.ShowModel = smStretch
    Case vdtVFW 'vfw����Ƶ�ɼ�
      'Dim vfwParameter As clsVfwCaptureParameterCfg
      'Set vfwParameter = mVfwControl.GetCaptureParameter()
      
      'vfwParameter.VideoShowWay = swFit
      
      'Call mVfwControl.SetCaptureParameter(vfwParameter)
      
      mVfwControl.CaptureParameterInf.VideoShowWay = swFit
      
      Call mVfwControl.ShowFullScreen(owner, mVfwWindow, GetMonitorIndex(ownerHandle))
      
      mVfwControl.CaptureParameterInf.VideoShowWay = swStretch
      'vfwParameter.VideoShowWay = swStretch
      
      'Call mVfwControl.SetCaptureParameter(vfwParameter)
    'Case ...
  End Select
  
End Sub


'ˢ����Ƶ����
Public Sub RefreshVideoWindow()

  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  Select Case mVideoDriverType
    Case vdtWDM 'wdm����Ƶ�ɼ�
      Call mWdmControl.RefreshWindow
    Case vdtVFW 'vfw����Ƶ�ɼ�
      Call mVfwControl.UpdateCaptureWindowPos(mVfwWindow.ScaleWidth, mVfwWindow.ScaleHeight)
    'Case ...
  End Select
  
End Sub


'�ɼ�ͼ���ļ�
Public Sub CaptureImgToFile(ByVal strFileName As String)

  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If
  
  '�ɼ�ͼ��
  Select Case mVideoDriverType
    Case vdtVFW
      Call mVfwControl.CaptureBmpImgToFile(strFileName)
    Case vdtWDM
      Call mWdmControl.CaptureBmpImageToFile(strFileName)
    'Case Other......
  End Select
  
End Sub


'���ɼ�ͼ�񱣴浽������
Public Function CaptureImgToClipBoard() As Boolean
  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    CaptureImgToClipBoard = False
    Exit Function
  End If
  
  On Error GoTo errHandle
    
    '�ɼ�ͼ��
    Select Case mVideoDriverType
        Case vdtVFW
            CaptureImgToClipBoard = mVfwControl.CaptureImageToClipBoard
        Case vdtWDM
            CaptureImgToClipBoard = (mWdmControl.CaptureImgToClipBoard = "")
            'Case Other......
    End Select
  
    
  Exit Function
errHandle:
  CaptureImgToClipBoard = False
End Function


Public Function CaptureImageFromMemory() As IPictureDisp
  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Set CaptureImageFromMemory = Nothing
    Exit Function
  End If
  
  On Error GoTo errHandle
    
    Dim iPic As IPictureDisp
    Set iPic = Nothing
    
    '�ɼ�ͼ��
    Select Case mVideoDriverType
        Case vdtVFW
        
            Clipboard.Clear
            
            Call mVfwControl.CaptureImageToClipBoard
            
            Set iPic = Clipboard.GetData(2)
            
            Clipboard.Clear
        Case vdtWDM
            Set iPic = mWdmControl.CaptureBmpImage
        'Case Other......
    End Select
  
    Set CaptureImageFromMemory = iPic
    
  Exit Function
errHandle:
  Set CaptureImageFromMemory = Nothing
End Function


Public Function CaptureImg() As IPictureDisp
  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Set CaptureImg = Nothing
    Exit Function
  End If
  
  On Error GoTo errHandle
  
    Dim strTempFile As String
  
    '�����ɼ�Ŀ¼
    If Dir(App.Path & "\TmpImage\", vbDirectory) = "" Then
      Call MkDir(App.Path & "\TmpImage\")
    End If
    strTempFile = App.Path & "\TmpImage\" & "DAEAE94F2CBC46018D6F1496C00447E3.dat"
    
  
    '�ɼ�ͼ��
    Select Case mVideoDriverType
        Case vdtVFW
            Call mVfwControl.CaptureBmpImgToFile(strTempFile)
        Case vdtWDM
            Call mWdmControl.CaptureBmpImageToFile(strTempFile)
            'Case Other......
    End Select
  
    If Dir(strTempFile) <> "" Then
        '���ļ�����ͼ��
        Set CaptureImg = LoadPicture(strTempFile)
  
        'ɾ���ɼ���ͼ���ļ�
        Call Kill(strTempFile)
    Else
        Set CaptureImg = Nothing
    End If
  
  Exit Function
errHandle:
  Set CaptureImg = Nothing
End Function


'��ʼ¼��
Public Function StartVideo(ByVal strFileName As String) As String
  StartVideo = ""
  
  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    StartVideo = "������������ǰ��ƵԴ��"
    Exit Function
  End If
  
  '��ʼ��Ƶ¼��
  Select Case mVideoDriverType
    Case vdtVFW
      Call mVfwControl.StartVideo(strFileName)
    Case vdtWDM
      mWdmControl.IsShowState = True
      
      StartVideo = mWdmControl.StartCaptureVideo(strFileName)
      
      If Trim(StartVideo) <> "" Then mWdmControl.IsShowState = False
    'Case Other......
  End Select
  
End Function


'ֹͣ¼��
Public Sub StopVideo()

  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If

  Dim strTempFileName As String
  
  'ֹͣ��Ƶ¼��
  Select Case mVideoDriverType
    Case vdtVFW
      Call mVfwControl.StopVideo
    Case vdtWDM
      Call mWdmControl.StopCaptureVideo(strTempFileName)
      
      mWdmControl.IsShowState = False
    'Case Other......
  End Select
  
End Sub


'��ʾ�ɼ��������ô���
Public Sub ShowCaptureParameterCfgDialog(objOwner As Object, Optional hWndParent As Long = 0)

  '������������Ƶ����ֱ���˳�
  If Not mblnIsAllowStartupVideo Then
    Exit Sub
  End If

  Select Case mVideoDriverType
    Case vdtWDM 'wdm����Ƶ�ɼ�
      mWdmControl.HideCfgItem = hciVideoShowWay + hciVideoState + hciImageCapture
      
      Call mWdmControl.ShowCaptureParameterCfgDialog(hWndParent)
    Case vdtVFW 'vfw����Ƶ�ɼ�
      Call mVfwControl.ShowCaptureParameterCfgDialog(objOwner, vpiShowWay)
    Case vdtCustom
        If Not mCustomControl Is Nothing Then Call mCustomControl.zlConfigDevice
    'Case Other......
  End Select
End Sub


'��ȡ��Ƶ¼��ı���������
Function GetEncoderName() As String
    Dim curCap As TCaptureParameter
    
    GetEncoderName = ""
    If mVideoDriverType <> vdtWDM Then Exit Function
    
    Call mWdmControl.GetCaptureParameter(curCap)
    GetEncoderName = curCap.EncoderName
End Function


'��ȡ��Ƶ¼���¼�Ƴ���
Function GetTimeLen() As Long
    GetTimeLen = 0
    If mVideoDriverType <> vdtWDM Then Exit Function
    
    GetTimeLen = mWdmControl.RecordTimeLen
End Function


'����VFW�豸����
Public Sub ConnectedVfwDeviceObj(ByRef vfwWindow As Object)
   Set mVfwWindow = vfwWindow
End Sub


'����WDM�豸����
Public Sub ConnectedWdmDeviceObj(ByRef wdmControl As Object)
  Set mWdmControl = wdmControl
End Sub

'����ר����Ƶ�ɼ��豸����
Public Sub ConnectedCustomDeviceObj(ByRef customControl As Object)
    Set mCustomControl = customControl
End Sub

'����ָ���ɼ�������
Public Sub ConnectedTrdDeviceObj()
  '�÷���Ŀǰ����Ҫʵ��......
  Call err.Raise(-1, "ConnectedTrdDeviceObj", "��֧�ָ÷�����")
End Sub


'������Ƶ��ʾģʽ
Public Sub SetVideoShowWay(ByVal curShowWay As Long)
    Select Case curShowWay
      Case swNormal
        mWdmControl.ShowModel = smNormal
        mVfwControl.CaptureParameterInf.VideoShowWay = swNormal
      Case swFit
        mWdmControl.ShowModel = smFit
        mVfwControl.CaptureParameterInf.VideoShowWay = swFit
      Case swStretch
        mWdmControl.ShowModel = smStretch
        mVfwControl.CaptureParameterInf.VideoShowWay = swStretch
      Case swAutoFitCut
        mWdmControl.ShowModel = smAutoFitCut
        mVfwControl.CaptureParameterInf.VideoShowWay = swAutoFitCut
      Case swWindowAutoFit
        mWdmControl.ShowModel = smWindAutoFit
        mVfwControl.CaptureParameterInf.VideoShowWay = swWindowAutoFit
      Case Else
        mWdmControl.ShowModel = smStretch
        mVfwControl.CaptureParameterInf.VideoShowWay = swStretch
    End Select
End Sub


'�������ļ���ȡ�ɼ�����
Public Sub ReadCaptureParameterFromFile(ByVal strCfgName As String)

  '��ȡvfw����
  mVfwControl.ParameterCfgFileName = strCfgName
  Call mVfwControl.ReadVfwCaptureParameterFromFile
 

  '��ȡwdm����
  mWdmControl.ParameterCfgFileName = strCfgName
  Call mWdmControl.ReadParameterFromFile
  
End Sub


'����ɼ��������ļ�
Public Sub SaveCaptureParameterToFile(ByVal strCfgName As String)
  
  '����vfw����
  mVfwControl.ParameterCfgFileName = strCfgName
  Call mVfwControl.SaveVfwCaptureParameterToFile
  
  '����wdm����
  mWdmControl.ParameterCfgFileName = strCfgName
  Call mWdmControl.SaveParameterToFile

End Sub


'�����ʼ��
Private Sub Class_Initialize()
  mblnIsAllowStartupVideo = True
  
  Set mVfwControl = New clsVfwCapture
    
  '��mVfwControl����ʾģʽ����Ϊ����ģʽ������Ӧͼ��ü��ĵ���
  Dim vfwParameter As clsVfwParameterCfg
  Set vfwParameter = mVfwControl.GetCaptureParameter()
  
  vfwParameter.VideoShowWay = swStretch
  
  Call mVfwControl.SetCaptureParameter(vfwParameter)
End Sub


'�ͷ���Դ
Private Sub Class_Terminate()
  Set mVfwControl = Nothing
End Sub