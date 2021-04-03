VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.Form frmReportImage 
   BorderStyle     =   0  'None
   Caption         =   "报告图像"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMiniCache 
      Height          =   3855
      Left            =   4080
      ScaleHeight     =   3795
      ScaleWidth      =   4155
      TabIndex        =   15
      Top             =   2760
      Width           =   4215
      Begin MSComCtl2.DTPicker DTPimg 
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   162463745
         CurrentDate     =   42674
      End
      Begin VB.ComboBox cboCache 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin zl9PacsControl.ucImagePreview ucMiniCache 
         Height          =   1215
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   3780
         _ExtentX        =   6694
         _ExtentY        =   2143
         BackColor       =   8421504
         ShowCheckbox    =   -1  'True
      End
   End
   Begin VB.PictureBox picMiniViewer 
      Height          =   1365
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   3615
      TabIndex        =   13
      Top             =   5280
      Width           =   3675
      Begin zl9PacsControl.ucImagePreview ucMiniImageViewer 
         Height          =   975
         Left            =   45
         TabIndex        =   14
         Top             =   120
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1746
         BackColor       =   8421504
      End
   End
   Begin VB.PictureBox picMenu 
      Height          =   540
      Left            =   2100
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   585
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin MSComctlLib.ImageList listCur 
      Bindings        =   "frmReportImage.frx":0000
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportImage.frx":0014
            Key             =   "Pen"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picReportImage 
      Height          =   2055
      Left            =   3600
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin DicomObjects.DicomViewer dcmReportImage 
         Height          =   1695
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _Version        =   262147
         _ExtentX        =   3413
         _ExtentY        =   2990
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picMark 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.PictureBox picNumMark 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1030
         Left            =   300
         ScaleHeight     =   1035
         ScaleWidth      =   2040
         TabIndex        =   4
         Top             =   1300
         Width           =   2040
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":032E
            Height          =   510
            Index           =   1
            Left            =   490
            Picture         =   "frmReportImage.frx":0F70
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":1BB2
            Height          =   510
            Index           =   4
            Left            =   510
            Picture         =   "frmReportImage.frx":27F4
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":3436
            Height          =   510
            Index           =   2
            Left            =   1000
            Picture         =   "frmReportImage.frx":4078
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":4CBA
            Height          =   510
            Index           =   5
            Left            =   1010
            Picture         =   "frmReportImage.frx":58FC
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":653E
            Height          =   510
            Index           =   3
            Left            =   1560
            Picture         =   "frmReportImage.frx":7180
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":7DC2
            Height          =   510
            Index           =   6
            Left            =   1510
            Picture         =   "frmReportImage.frx":8A04
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":9646
            Height          =   1020
            Index           =   0
            Left            =   0
            Picture         =   "frmReportImage.frx":A288
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "自动编号"
            Top             =   0
            Value           =   1  'Checked
            Width           =   510
         End
      End
      Begin DicomObjects.DicomViewer dcmMark 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _Version        =   262147
         _ExtentX        =   2990
         _ExtentY        =   1720
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdate As Date '用于按时间过滤后台图
Private mintTagMaxTag As Integer '标识X中的最大X（用于判定是否更新菜单栏信息）
Private mintTagNow As Integer '当前标识
Private mintTagMax As Integer '最大标识
Private mblDel As Boolean '是否点击过删除，用于同步刷新采集模块界面

Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Private mlngAdviceID As Long    '医嘱ID
Private mintEditType As Integer '病历状态 0 创建，1书写，2 修订
Private mlngReportID As Long    '报告内容ID
Private mlngFileID As Long      '报告单格式ID
Private mlngShowBigImg As Long          '是否显示大图,0-不显示；1-鼠标移动时显示；2-鼠标单击显示独立窗口
Private mintImageDblClick As Integer    '缩略图双击后的操作 0--直接写入报告；1--打开图片编辑窗口
Private mblnEditable As Boolean         '是否可以编辑内容
Private mintMoustType As Integer        '鼠标工作类型
Private mblnUserInvoke As Boolean       '是否用户操作触发
Private mblnMoved As Boolean            '是否已经转储
Private mintCurImgIndex As Integer      '当前选中的图像
Private mintShowPhotoNumber As Integer  '当前界面能够显示出图像的最佳数量
Private mlngModule As Long

Public mSelMiniImg As DicomImage
Private mSelReportImg As DicomImage
Private mSelViewerIndex As Integer  '当前被选中的报告图象框ID，从1开始计数
Private mselReportImgIndex As Integer   '当前被选中的报告图像ID，从1开始计数
Private mdblMarkZoom As Double          '当前标记图中实际像素和标记之间的缩放比例
Private lngColor(10) As Long             '标记图中圆形编号使用的9个颜色
Private mlngCY1 As Long                 '标记图的高度
Private mlngMarkW As Long               '标记图的宽度
Private mlngCY2 As Long                 '报告图的高度
Private mlngRptImgW As Long             '报告图的宽度
Private mlngCY3 As Long                 '缩略图图的高度

Public pMarkModified As Boolean        '标记图的标记有改动
Public pImageModified As Boolean       '记录报告图像是否修改，如果没有修改，则保存报告的时候不再保存图像
Public pobjMarks As cPicMarks          '当前标记图的标注对象
Public pMarkImageID As Double            '当前标记图在数据库表“电子病历内容”表中的ID
Public pTableID As String              '当前图像所在表格的ID串，用“;”分隔。影响能否保存报告图，问题108069对这个变量的清0做了调整。


Private mintShowMarkImage As Integer   '是否显示标记图   0-隐藏标记图  1-显示标记图
Private mblnIsInitFace As Boolean        '是否已经加载窗体
Private mobjImgCTables() As cEPRTable

Private blnLoadImages As Boolean        '记录本次刷新是否加载了图像


Private mdcmGlobal As New DicomGlobal    '定义UIDRoot=1

Private mrsImageCache As New ADODB.Recordset
Private mdcmUID As New DicomGlobal
Private mlngReleationType As Integer    '1--导出，2--导入
Private mlngCurDeptId As Long
Private mlngStudyState As Long
Private mstrTmpQueryPath As String
Private mblnUseAfterCapture As Boolean
Private mblnTmpUseAfterCapture As Boolean
Private mstrAfterImgPath As String
Private mobjFile As New FileSystemObject
Private mstrInstance As String
Private mblnImageShield As Boolean     '是否屏蔽大图

Private WithEvents mobjImageProcess As zl9PacsControl.clsImageProcess
Attribute mobjImageProcess.VB_VarHelpID = -1
Private mobjPacsCapture As Object

Public Event OnImageCountChanged(ByVal intType As Integer, ByVal isNeedRefreshTitle As Boolean) '图像数量改变 intType: 0:发送到检查   1:删除
Public Event AfterReleationImage(ByVal lngReleationType As Long)
Public Event AfterShowBigImage()

Private Enum MarkType
    自动编号 = 0: 编号1: 编号2: 编号3: 编号4: 编号5: 编号6
End Enum

Property Get ImageCount() As Long
    ImageCount = ucMiniImageViewer.CurImageCount
End Property

Property Get ReportImageCount() As Long
    Dim i As Long
    Dim lngResult As Long
    
    lngResult = 0
    For i = 0 To dcmReportImage.Count - 1
        lngResult = lngResult + dcmReportImage(i).Images.Count
    Next i
    
    ReportImageCount = lngResult
End Property

Property Get dcmImages() As Object
    Set dcmImages = ucMiniImageViewer.ImgViewer.Images
End Property

Public Sub RefreshAfterImage()
    LoadMiniCache
End Sub

Public Sub MovePage(ByVal lngPageType As TMoveType)
'移动缩略图页面
    ucMiniImageViewer.MovePage (lngPageType)
End Sub


Public Sub zlRefresh(ByVal lngAdviceID As Long, FileID As Long, _
        ReportID As Long, blnSingleWindow As Boolean, lngShowBigImg As Long, _
        intImageDblClick As Integer, blnEditable As Boolean, _
        ByVal blnMoved As Boolean, ByVal intMinImageCount As Integer, blnFormIsSelected As Boolean, _
        ByVal lngModule As Long, ByVal lngCurDeptId As Long, ByVal lngStudyState As Long, _
        ByVal blnIsSaveRefresh As Boolean)
        
    Dim i As Integer
    Dim intShowMarkImage As Integer
        
    Call GetNowTag(True)
    
    mlngCurDeptId = lngCurDeptId
    mlngStudyState = lngStudyState
    mlngAdviceID = lngAdviceID
    mlngFileID = FileID
    mlngReportID = ReportID
    mlngShowBigImg = lngShowBigImg
    mintImageDblClick = intImageDblClick
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    mintShowPhotoNumber = intMinImageCount
    mlngModule = lngModule
    mblnSingleWindow = blnSingleWindow
    mstrAfterImgPath = IIf(Len(App.Path) > 3, App.Path & "\TmpAfterImage\", App.Path & "\TmpAfterImage\")
    
    Call InitCTables
    
    intShowMarkImage = DecideMarkImagesVisible    '判断标记图是否可见
    
    If mlngModule = 1291 Then
        mblnUseAfterCapture = Val(GetDeptPara(mlngCurDeptId, "启用后台采集", 1, True)) = 1
    Else
        mblnUseAfterCapture = False
    End If
    
    ucMiniImageViewer.BigImageWay = lngShowBigImg
    ucMiniImageViewer.PreViewTime = Val(GetDeptPara(mlngCurDeptId, "移动预览延时", 0))
    ucMiniImageViewer.ShowPopup = False
    ucMiniImageViewer.ImgLoadType = IIf(GetServiceStatus = SERVICE_RUNNING, FileLoadType.Service, FileLoadType.Normal)
    
    Call GetLocalPar
    ucMiniImageViewer.ImageShield = mblnImageShield
    
    '只加载 报告图像 字段中的报告图，若字段内容为空，再加载所有报告图
    ucMiniImageViewer.OnlyLoadReportImage = True
    
    
    '判断如果是 独立窗口 或者 没有加载过窗体 或者 标记图状态已经改变，则重新加载初始化界面
    If mblnSingleWindow Or (Not mblnIsInitFace Or (mintShowMarkImage <> intShowMarkImage)) Or (mblnTmpUseAfterCapture <> mblnUseAfterCapture) Then
        mintShowMarkImage = intShowMarkImage
'        Call InitLoaclParas     '读取本机参数
        Call InitFaceScheme     '初始化窗体界面
    End If
    
     mblnTmpUseAfterCapture = mblnUseAfterCapture
    
    '重新初始化内部参数
    pMarkImageID = 0
    pImageModified = False
    pMarkModified = False
    dcmMark.Images.Clear
    
    If Not (pobjMarks Is Nothing) Then
        For i = 1 To pobjMarks.Count
            pobjMarks.Remove 1
        Next i
    End If
    
    
    '标记本次刷新还没有加载图像
    blnLoadImages = False
    
    '如果窗体是正在被显示的，则加载图像
    If blnFormIsSelected = True Then ' And Me.Visible
        '根据需要加载图像
        If Not blnIsSaveRefresh Then    '如果是签名或者保存报告，则不载入图像
            Call LoadImages
        Else
            Call LoadReportImages
            blnLoadImages = True
        End If
    Else
        Call ClearReportImages
    End If
    
    '设置界面控件是否可以编辑
    picMark.Enabled = mblnEditable
    picReportImage.Enabled = mblnEditable
    picMiniViewer.Enabled = mblnEditable
End Sub

Private Sub ClearReportImages()
    Dim i As Integer
    
    pTableID = ""
    '初始化各个对象
    For i = 1 To dcmReportImage.Count - 1
        Unload dcmReportImage(i)
    Next i
    dcmMark.Images.Clear
End Sub

Public Sub RefPacsPic(Optional ByVal lngEventType As TVideoEventType = vetUpdateImg)
    '读取和显示当前可选报告图像
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        Call LoadMiniCache
    End If
    
    If lngEventType <> vetAfterUpdateImg Then Call LoadMiniImages
End Sub

Private Sub cboCache_Click()
    Dim strQueryPath As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    ucMiniCache.ClearCurrentPageImage
    
    Set rsTmp = mrsImageCache
    rsTmp.Filter = ""
        
    If mrsImageCache.RecordCount <= 0 Then Exit Sub
    
    mrsImageCache.MoveFirst
    Set rsTmp = mrsImageCache

    rsTmp.Filter = "姓名='" & Trim(Mid(cboCache.Text, 1, 5)) & "'"

    If rsTmp.RecordCount < 1 Then Exit Sub
    strQueryPath = Nvl(rsTmp!路径)

    If strQueryPath = "" Then Exit Sub


    Call ucMiniCache.RefreshImage(slLocal, strQueryPath, mblnMoved)
    mstrTmpQueryPath = strQueryPath
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
    err.Clear
End Sub

Private Sub cboCache_DropDown()
On Error GoTo errHandle
    Call SendMessage(cboCache.hWnd, &H160, 500, 0)
errHandle:
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    
    Select Case control.ID
        Case comMenu_Cap_Process '图像处理
            Call OpenImageProcessWind
        Case conMenu_Cap_DevSet
            If mblnUseAfterCapture And mlngModule <> 1290 Then Call ucMiniCache.ShowPageConfig
            Call ucMiniImageViewer.ShowPageConfig
        Case conMenu_PacsReport_DelImage    '删除图像
            If dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex And mselReportImgIndex <> 0 Then
                dcmReportImage(mSelViewerIndex).Images.Remove mselReportImgIndex
                Call picReportImage_Resize
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveUp      '前移图像
            If mselReportImgIndex > 1 And dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex - 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveDown    '后移图像
            If mselReportImgIndex > 0 And dcmReportImage(mSelViewerIndex).Images.Count > mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex + 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_DelMarks    '清除标注
            If dcmMark.Images.Count > 0 Then
                dcmMark.Images(1).Labels.Clear
                dcmMark.Refresh
                For i = 1 To pobjMarks.Count
                    pobjMarks.Remove 1
                Next i
                pMarkModified = True
            End If
        Case conMenu_View_Refresh           '刷新
            '读取和显示当前可选报告图像
            Call LoadMiniImages
        Case conMenu_PacsReport_DelMiniImage    '删除报告图
            
        Case conMenu_PacsReport_SelMiniImage    '提取报告图
            Dim resImages As DicomImages
            
            Set resImages = frmSelectRepImage.ShowMe(Me, mlngAdviceID)
            '把当前图形添加到图象框中
            If resImages.Count > 0 Then
                For i = 1 To resImages.Count
                    dcmReportImage(mSelViewerIndex).Images.Add resImages(i)
                    dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).BorderColour = vbWhite
                Next i
                dcmReportImage(mSelViewerIndex).CurrentIndex = 1
                Call picReportImage_Resize
                pImageModified = True
            End If
            
        Case conMenu_Cap_SendToAdvice        '发送到检查
            mlngReleationType = 2
            Call ReleationImage
        
        Case conMenu_Cap_SendToAfter     '发送到后台
            mlngReleationType = 1
            Call ReleationImage
        
        Case conMenu_Manage_DeleteImage '删除临时图象
            mlngReleationType = 2
            mblDel = True
            Call DelTempImage
        
        Case conMenu_Manage_RefreshImg  '刷新缓存
            Call LoadMiniCache
        
        Case conMenu_Cap_ImageShield    '屏蔽大图
            control.Checked = Not control.Checked
            
            mblnImageShield = control.Checked
            ucMiniImageViewer.ImageShield = mblnImageShield
            Call SaveLocalPar
    End Select
End Sub

Private Sub CheckSendOnImageCountChangedChanged(ByVal intType As Integer)
'intType 0:发送到检查  1:删除图像
'isNeedRefreshTitle:是否需要更新数量

    If (InStr(cboCache.Text, "标识" & Lpad((mintTagMaxTag), 3, "0")) > 0) Then
        RaiseEvent OnImageCountChanged(intType, True)
    Else
        RaiseEvent OnImageCountChanged(intType, False)
    End If
End Sub

Private Sub DelTempImage()
    Dim rsImageDatas As ADODB.Recordset
    Dim i As Long
    
    '在数据库中查询图像数据
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "请选择需要删除的检查图像。", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    '当前检查UID在数据库中不存在，则退出本程序
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "请选择需要删除的检查图像。", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "是否确认删除所选图像？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    Call DelTempImages(rsImageDatas)
End Sub

Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'删除本地缓存中的文件，在界面上删除ucpre控件的选中图像
    Dim blfinished As Boolean
    Dim i As Long
    Dim curTime As Date
    Dim intTMP As Integer
On Error GoTo errHandle
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        If mobjFile.FileExists(rsImageDatas!路径) Then Call mobjFile.DeleteFile(rsImageDatas!路径)
        
        rsImageDatas.MoveNext
    Wend
    
        '删除界面图像
    blfinished = False
    For i = ucMiniCache.CurImageCount To 1 Step -1
        If ucMiniCache.ImgChecked(i) Then
            Call ucMiniCache.DeleteImage(i)
            blfinished = True
        End If
    Next

    If blfinished = False Then
        Call ucMiniCache.DeleteImage(ucMiniCache.Selectindex)
    End If

    '同时需要删除cbo项目
    Call ClearEmptyFolder(False)
    If ucMiniCache.CurImageCount = 0 Then
        curTime = zlDatabase.Currentdate
        '是当天并且选中的是当前标识，就不进行清空操作
        If Not ((Format(DTPimg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCache.Text, "标识" & Lpad((mintTagMaxTag), 3, "0")) > 0)) Then
            intTMP = cboCache.ListIndex
            Call cboCache.RemoveItem(cboCache.ListIndex)
            If cboCache.ListCount > intTMP - 1 Then
                cboCache.ListIndex = intTMP - 1
            Else
                cboCache.ListIndex = 0
            End If
        End If
    End If
        
    DelTempImages = True
    
    Exit Function
errHandle:
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Function ReleationImage() As Boolean
    Dim strHint As String
    Dim rsImageDatas As ADODB.Recordset
    Dim strTmpFile As String
    Dim i As Integer
    
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "请选择需要进行关联的检查图像。", vbInformation, Me.Caption)
        Exit Function
    End If
        
    '当前检查UID在数据库中不存在，则退出本程序
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "请选择需要进行关联的检查图像。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If mlngReleationType = 2 Then
        '关联图像提示
        strHint = GetReleationHintInfo(mlngAdviceID, rsImageDatas)
        
        If strHint = "" Then
            Call MsgBoxD(Me, "不能查询到需要关联的数据信息，结束关联。", vbOKOnly, Me.Caption)
            Exit Function
        End If
        
        If MsgBoxD(Me, strHint, vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        
    Else
        '取消关联提示
        If MsgBoxD(Me, "是否确认将所选图像发送到后台？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
    End If

    If mlngReleationType = 2 Then '等于2表示关联图像
        ReleationImage = StartReleation(mlngAdviceID, rsImageDatas)
        Call ClearEmptyFolder(False)
    Else
        ReleationImage = CancelReleation(mlngAdviceID, rsImageDatas)
    End If
        
    '操作后清除红框，防止出现2个红框的BUG
    For i = 1 To ucMiniImageViewer.CurImageCount
        ucMiniImageViewer.ImgViewer.Images(i).BorderColour = vbWhite
    Next
    
    If ReleationImage Then RefPacsPic
    RaiseEvent AfterReleationImage(mlngReleationType)
End Function

'取得关联提示信息
Private Function GetReleationHintInfo(lngAdviceID As Long, rsReleationImage As ADODB.Recordset) As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strResult As String
    Dim strStudyInf As String
    
    GetReleationHintInfo = ""
    
    strSQL = "select 检查号,姓名,性别,年龄 from 影像检查记录 where 医嘱ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    GetReleationHintInfo = "是否确认将选择的图像发送到[" & Nvl(rsTemp!姓名) & "(" & Nvl(rsTemp!检查号) & ") " & Nvl(rsTemp!性别) & " " & Nvl(rsTemp!年龄) & "]的检查中？"
End Function

Private Function GetReleationImageIds() As ADODB.Recordset
'查询关联或者要取消关联的图像ID
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String
    Dim strTmpFile As String
    Dim rsImageDatas As ADODB.Recordset

    j = 0
    strUninTable = ""
    strFilter = ""
    strValue = ""
    
    '构造查询语句
    If mlngReleationType = 1 Then
        For i = 1 To ucMiniImageViewer.CurImageCount
            If ucMiniImageViewer.ImgChecked(i) Then
                If j > 79 Then
                    strFilter = strFilter & " Or 图像UID ='" & ucMiniImageViewer.ImgViewer.Images(i).InstanceUID & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 3600 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         
                         strValue = ""
                         j = j + 1
                    End If
                    
                    strValue = strValue & "," & ucMiniImageViewer.ImgViewer.Images(i).InstanceUID
                End If
            End If
        Next
                
        '若所有图像都没有被选中的红点，则有红框的图像视为选中
        If Not ucMiniImageViewer.SelectImage Is Nothing And strValue = "" Then strValue = strValue & "," & ucMiniImageViewer.SelectImage.InstanceUID
                
    Else
        Set rsImageDatas = New ADODB.Recordset
        rsImageDatas.Fields.Append "序列UID", adVarChar, 4000
        rsImageDatas.Fields.Append "检查UID", adVarChar, 4000
        rsImageDatas.Fields.Append "路径", adVarChar, 4000
        rsImageDatas.Open
            
        For i = 1 To ucMiniCache.CurImageCount
            If ucMiniCache.ImgChecked(i) Then
                strTmpFile = ucMiniCache.ImgViewer.Images(i).tag.FilePath
                rsImageDatas.AddNew
                rsImageDatas!序列UID = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                rsImageDatas!检查uid = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                rsImageDatas!路径 = strTmpFile
                rsImageDatas.Update
            End If
        Next
        
        '没有图像处于红点状态，就选择有红框的
        If rsImageDatas.RecordCount = 0 Then
            If ucMiniCache.CurImageCount > 0 Then
                If Not ucMiniCache.SelectImage Is Nothing Then
                    strTmpFile = ucMiniCache.SelectImage.tag.FilePath
                    rsImageDatas.AddNew
                    rsImageDatas!序列UID = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                    rsImageDatas!检查uid = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                    rsImageDatas!路径 = strTmpFile
                    rsImageDatas.Update
                End If
            End If
        End If
                
        If rsImageDatas.RecordCount > 0 Then rsImageDatas.MoveFirst
                
        Set GetReleationImageIds = rsImageDatas
        Exit Function
    End If
    
    If strValue <> "" Then
        strValues(j) = Mid(strValue, 2)
        strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
    End If
    
    '如果没有需要查找的图像UID，则返回空数据集
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
        Set GetReleationImageIds = Nothing
        Exit Function
    End If
    
    If strFilter <> "" Then strUninTable = strUninTable & " Union All Select 图像UID from [影像图象] where  ( " & Mid(strFilter, 4) & ")"
    
    strSQL = "Select /*+ RULE*/ D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd, Decode(C.位置一,Null,C.位置二,C.位置一) as 设备号," & _
        "D.IP地址 As Host,B.序列UID,B.检查UID,C.影像类别, " & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL,A.图像UID, c.姓名,c.性别,c.年龄,c.检查号 " & _
        "From 影像检查图象 A, 影像检查序列 B, 影像检查记录 C,影像设备目录 D,(" & Replace(strUninTable, "[影像图象]", "影像检查图象") & ") E " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And A.序列UID=B.序列UID and B.检查UID=C.检查UID and A.图像UID = E.图像UID "
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    
    Set GetReleationImageIds = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValues(0), strValues(1), strValues(2), strValues(3), _
        strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10), _
        strValues(11), strValues(12), strValues(13), strValues(14), strValues(15), strValues(16), strValues(17), _
        strValues(18), strValues(19), strValues(20), strValues(21), strValues(22), strValues(23), strValues(24), strValues(25), strValues(26), _
        strValues(27), strValues(28), strValues(29), strValues(30), strValues(31), strValues(32), strValues(33), strValues(34), strValues(35), strValues(36), _
        strValues(37), strValues(38), strValues(39), strValues(40), strValues(41), strValues(42), strValues(43), strValues(44), strValues(45), strValues(46), _
        strValues(47), strValues(48), strValues(49), strValues(50), strValues(51), strValues(52), strValues(53), strValues(54), strValues(55), strValues(56), _
        strValues(57), strValues(58), strValues(59), strValues(60), strValues(61), strValues(62), strValues(63), strValues(64), strValues(65), strValues(66), _
        strValues(67), strValues(68), strValues(69), strValues(70), strValues(71), strValues(72), strValues(73), strValues(74), strValues(75), strValues(76), _
        strValues(77), strValues(78), strValues(79), strValues(80))
End Function

Private Function StartReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'开始关联
On Error GoTo errHandle
    Dim strSQL As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objFtp As New clsFtp
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    strSQL = "select 检查UID,接收日期 from 影像检查记录 where 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgBoxD(Me, "找不到待关联的检查信息。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If Trim(Nvl(rsTmp!检查uid)) = "" Or Trim(Nvl(rsTmp!接收日期)) = "" Then
        '尚未采集图像，需要生成新的检查UID
        strNewStudyUID = CreateStudyUid(rsImageDatas!检查uid)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
            Exit Function
        End If
        
        '更新存储设备信息
        strSQL = "Zl_影像检查_更新设备(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        strNewStudyUID = Nvl(rsTmp!检查uid)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    
    '连接FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgBoxD(Me, "FTP连接失败，请检查网络设置。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    '移动图像文件
    If Not MoveImageToStudy(objFtp, rsImageDatas, strNewFtpVirtualPath, objMoveList) Then Exit Function
          
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        '创建新的序列UID
        strNewSeriesUid = CreateSeriesUid(rsImageDatas!序列UID, strNewStudyUID)
        
        '更新图像关联数据
        strSQL = "Zl_影像检查_图像导入(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & mobjFile.GetFileName(Nvl(rsImageDatas!路径)) & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        rsImageDatas.MoveNext
    Wend
    
    '提交事务
    Call gcnOracle.CommitTrans
    
    '说明全部上传成功,删除本地临时图像
    Call DelTempImages(rsImageDatas)
    
    StartReleation = True
    
    Exit Function
errHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '继续抛出错误
End Function

Private Function CreateSeriesUid(ByVal strSeriesUID As String, ByVal strStudyUID As String) As String
'创建序列UID
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim strNewSeriesUid As String
    
    strNewSeriesUid = strSeriesUID 'M_STR_SERIES_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)
    
    strSQL = "select 序列UID from 影像检查序列 where 序列UID = [1] And 检查UID <> [2]"
              
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存", strNewSeriesUid, strStudyUID)
    
    If rsData.RecordCount > 0 Then
        '创建一个新的检查UID
        strSQL = "Select 影像检查UID序号_ID.Nextval From Dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存")
        
        If Len(strNewSeriesUid) <= 55 Then
            strNewSeriesUid = strNewSeriesUid & ".A" & rsData(0)
        Else
            strNewSeriesUid = Left(strNewSeriesUid, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateSeriesUid = strNewSeriesUid
End Function

Private Function CancelReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'撤销关联
On Error GoTo errHandle
    Dim strSQL As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim objFtp As New clsFtp
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    CancelReleation = False
    
    curDate = zlDatabase.Currentdate
    
    '撤销图像关联
    strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
    Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
    If Trim(strNewFtpIp) = "" Then
        Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    '连接FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgBoxD(Me, "FTP连接失败，请检查网络设置。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If Not MoveImageToAfter(objFtp, rsImageDatas, objMoveList) Then Exit Function
    
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    '更新数据
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        strSQL = "Zl_影像检查_图像导出(" & mlngAdviceID & ",'" & Nvl(rsImageDatas!图像UID) & "')"
                                        
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        rsImageDatas.MoveNext
    Wend
    
    Call gcnOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
errHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    Call OutputDebug("CancelReleation", err)
    Call RaiseErr(err)
End Function

Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo errHandle
'转移图像成功后，在删除临时图像和原有FTP的图像和目录，清场操作出现错误可以不处理
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strTmpFile = App.Path & "\TmpImage\" & Nvl(rsImageDatas!图像UID)
        
        strImageUID = Nvl(rsImageDatas!图像UID)
        
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
                
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
        
'       为避免重新下载图像，如果本地存在图像文件，则不用进行删除
        
        If FileExists(strTmpFile) Then Call Kill(strTmpFile)
        If FileExists(strTmpFile & ".jpg") Then Call Kill(strTmpFile & ".jpg")
        
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        '删除空的ftp目录
        Call objSrcFtp.FuncFtpDelDir(Replace(strVirtualPath, strImageUID, ""), strImageUID)
                
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
Exit Sub
errHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub

'撤销图像的移动
Private Sub CancelImageMove(ByVal strFTPIP As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo errHandle

    Call objFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
errHandle:
    objFtp.FuncFtpDisConnect
End Sub

Private Function MoveImageToStudy(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, strNewFtpVirtualPath As String, ByRef objMoveList As Collection) As Boolean
    Dim i As Long
    
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        If objFtp.FuncUploadFile(strNewFtpVirtualPath, rsImageDatas!路径, mobjFile.GetFileName(rsImageDatas!路径)) <> 0 Then
            '失败后删除之前上传的文件
            For i = 0 To objMoveList.Count - 1
                Call objFtp.FuncDelFile(strNewFtpVirtualPath, objMoveList(i))
            Next
            
            Call MsgBoxD(Me, "图像移动失败，请检查FTP传输是否正常。", vbInformation, Me.Caption)
            
            Exit Function
        Else
            Call objMoveList.Add(rsImageDatas!路径)
        End If
        
        rsImageDatas.MoveNext
    Wend
    
    MoveImageToStudy = True
End Function

Private Function MoveImageToAfter(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, ByRef objMoveList As Collection) As Boolean
    Dim i As Long
    Dim strDestPath As String
    
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        strDestPath = GetAfterImagePath(rsImageDatas!图像UID, rsImageDatas!序列UID, rsImageDatas!检查uid, rsImageDatas!影像类别)
        If mobjFile.FolderExists(strDestPath) = False Then Call MkLocalDir(strDestPath)
        
        If objFtp.FuncDownloadFile(rsImageDatas!Root & rsImageDatas!Url, strDestPath & rsImageDatas!图像UID, rsImageDatas!图像UID) <> 0 Then
            '失败后删除之前下载的文件
            For i = 0 To objMoveList.Count - 1
                Call mobjFile.DeleteFile(objMoveList(i))
            Next
            
            Call MsgBoxD(Me, "图像移动失败，请检查FTP传输是否正常。", vbInformation, Me.Caption)
            
            Exit Function
        Else
            Call objMoveList.Add(strDestPath & rsImageDatas!图像UID)
        End If
        
        rsImageDatas.MoveNext
    Wend
    
    Call MsgBoxD(Me, "已将选中图像发送到[检查" & mintTagNow & "]中", vbInformation, "提示")
        
    MoveImageToAfter = True
End Function

Public Function GetAfterImagePath(ByVal strImageName As String, ByVal strSeriesUID As String, ByVal strStudyUID As String, ByVal strModality As String) As String
    Dim strTmpPath As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim curDate As Date
    Dim strDate As String
    Dim intTMP As Integer
    
    curDate = zlDatabase.Currentdate
    strDate = Format(curDate, "yyyymmdd")
        
    strTmpPath = ""
    
    If mobjFile.FolderExists(mstrAfterImgPath & "\") Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterImgPath & "\").SubFolders   '时间层
            If objFolder1.Name = strDate Then '时间只判断当天

                For Each objFolder2 In mobjFile.GetFolder(objFolder1.Path).SubFolders   '检查层
                
                    If InStr(objFolder2.Name, "检查" & mintTagNow) > 0 Then '判断是否有这个检查+当前标识的目录，若有，直接使用，
                        
                        For Each objFolder3 In mobjFile.GetFolder(objFolder2.Path).SubFolders   '序列层
                            strTmpPath = objFolder3.Path & "\"
                            GoTo step2
                        Next
                   
                    End If
                Next
                
                Exit For '终止时间层文件夹的搜索
            End If
        Next
    End If
    
    If strTmpPath = "" Then
        strTmpPath = mstrAfterImgPath & "\" & Format(curDate, "yyyymmdd") & "\" & "检查" & mintTagNow & "-" & strStudyUID & "\" & strSeriesUID & "\"
    End If
        
    '找到目录后停止前面的遍历，直接进入step2
step2:
    GetAfterImagePath = strTmpPath
End Function

Private Sub MoveReportImage(ByVal strDicomFile As String, ByVal strImgUid As String, _
    objSrcFtp As clsFtp, objDestFtp As clsFtp, ByVal strSourceVirtualPath As String, ByVal strDestVirtualPath As String, _
    objMoveList As Collection, Optional ByVal lngWay As Long = 0)
On Error GoTo errHandle
'移动报告图
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim lngResult As Long
    
    If lngWay = 0 Then
        Call objSrcFtp.FuncDelFile(strSourceVirtualPath, strImgUid & ".jpg")
        
        '如果本地中存在从源ftp中下载的dicom图像，则将图像转换成jpg，并保存到目的ftp设备中
        If FileExists(strDicomFile) Then
            Call dcmImages.Clear
            Set dcmImg = dcmImages.ReadFile(strDicomFile)
    
            Call dcmImg.FileExport(strDicomFile & ".jpg", "JPG")
            Call objDestFtp.FuncUploadFile(strDestVirtualPath, strDicomFile & ".jpg", strImgUid & ".jpg")
            
            If FileExists(strDicomFile & ".jpg") Then Call Kill(strDicomFile & ".jpg")
        End If
    Else
        '如果源ftp设备中不存在该图像，则不进行移动
        If objDestFtp.FuncFtpFileExists(strSourceVirtualPath, strImgUid & ".jpg") Then
            lngResult = objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
            
            If lngResult <> 0 Then
                '如果文件移动失败，则端开连接重试一次
                Call objDestFtp.FuncFtpDisConnect
'                Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                Call objDestFtp.ResotreFtpConnect
                
                Call objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
                
                '记录已经被移动过的文件，以便在处理数据失败的时候，还可对移动的图像进行恢复
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strSourceVirtualPath & "/" & strImgUid & ".jpg" & ">" & strDestVirtualPath & "/" & strImgUid & ".jpg")
                End If
            End If
        End If
    End If
Exit Sub
errHandle:
    Call OutputDebug("MoveReportImage", err)
End Sub

Private Sub GetStorageDevice(ByVal lngAdviceID As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFTPIP As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'获取新的存储设备信息，如果设备存储信息部存在，则需要进行增加
'如果是取消关联，则使用strNewStudyUID将不能从数据库中查找到对应的数据
'strDeviceNum:设备号
'strFtpIp: ftp地址
'strFtpUrl: ftp目录
'strVirtualPath: ftp虚拟存储路径
'strFtpUser: ftp用户名
'strFtpPwd: ftp密码



    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    
    strFTPIP = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSQL = "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像检查记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1]"
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '如果执行到这里，说明是执行图像关联,需要判断当前检查的存储设备是否有效，如果无效需生成新的存储设备
        If Trim(rsData!接收日期) = "" Then
            blnIsGetNewDevice = True
        Else
            strDeviceNO = Nvl(rsData!位置一)
            strFTPIP = Nvl(rsData!host)
            strFtpUrl = Nvl(rsData!Root)
            strFTPUser = Nvl(rsData!FtpUser)
            strFTPPwd = Nvl(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & Nvl(rsData!Url)
        End If
    End If
    
    
    If blnIsGetNewDevice Then
        '生成新的检查UID和存储设备,如果执行到这里，说明是取消关联
        
        If mlngModule = 1290 Then
            '查询医技工作站中，检查所对应的存储设备
            strSQL = "select d.参数值 " & _
                        " from 医技执行房间 a, 病人医嘱发送 b, 影像DICOM服务对 c, 影像DICOM服务参数 d " & _
                        " Where a.科室ID = b.执行部门id And a.执行间 = b.执行间 And a.检查设备 = c.设备号 " & _
                        " and c.服务功能='图像接收' and c.服务ID=d.服务ID and d.参数名称='存储设备' and b.医嘱id=[1]"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
            
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD Me, "未找到图像存储设备,请确认当前检查所用设备是否在影像设备目录的服务配置中配置了图像存储。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strDeviceNO = Nvl(rsTemp!参数值)
        Else
            '查询非医技工作站中的图像存储设备
            strDeviceNO = GetDeptPara(mlngCurDeptId, "存储设备号")
            
            If Val(strDeviceNO) <= 0 Then
                MsgBoxD Me, "未找到图像存储设备,请确认在影像流程管理中是否对该科室配置了图像采集存储设备。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        strSQL = "Select 设备号,设备名,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL,FTP用户名,FTP密码,IP地址 " & _
                    " From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.tag, strDeviceNO)
        
        '如果存储设备停用，则直接退出
        If rsTemp.RecordCount <= 0 Then
            MsgBoxD Me, "未找到存储设备,请确认设备号为 [" & strDeviceNO & "] 的设备是否启用。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strFtpUrl = Nvl(rsTemp("URL"))
        strFTPIP = Nvl(rsTemp("IP地址"))
        strFTPUser = Nvl(rsTemp("FTP用户名"))
        strFTPPwd = Nvl(rsTemp("FTP密码"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFTPIP, strFTPUser, strFTPPwd
        On Error GoTo errHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '创建FTP目录
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
errHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    
    Select Case control.ID
        Case conMenu_Cap_SendToAdvice
            If mlngAdviceID <= 0 Or mlngStudyState < 2 Then control.Enabled = False
    End Select
    
    Exit Sub
errHandle:
    
End Sub

Private Sub chkMark_Click(Index As Integer)
    Dim i As Integer
    If mblnUserInvoke = False Then
        mblnUserInvoke = True
    Select Case Index
        Case 0
            mintMoustType = MarkType.自动编号
        Case 1
            mintMoustType = MarkType.编号1
        Case 2
            mintMoustType = MarkType.编号2
        Case 3
            mintMoustType = MarkType.编号3
        Case 4
            mintMoustType = MarkType.编号4
        Case 5
            mintMoustType = MarkType.编号5
        Case 6
            mintMoustType = MarkType.编号6
    End Select
    For i = 0 To 6
        chkMark(i).value = 0
    Next i
    chkMark(Index).value = 1
    mblnUserInvoke = False
    End If
End Sub

Private Sub dcmMark_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lTemp As DicomLabel
    Dim strNum As Integer
    
    If Button = 1 And dcmMark.Images.Count > 0 And picMark.MousePointer = 99 Then
        '画标注
        '两种类型的标注，一种是直接自动编号，另一种是手工编号
        pobjMarks.Add pobjMarks.Count + 1
        pobjMarks(pobjMarks.Count).Selected = False
        pobjMarks(pobjMarks.Count).类型 = 6     '圆形编号
        If mintMoustType = MarkType.自动编号 Then
            pobjMarks(pobjMarks.Count).内容 = pobjMarks.Count
        Else
            Select Case mintMoustType
                Case MarkType.编号1
                    pobjMarks(pobjMarks.Count).内容 = 1
                Case MarkType.编号2
                    pobjMarks(pobjMarks.Count).内容 = 2
                Case MarkType.编号3
                    pobjMarks(pobjMarks.Count).内容 = 3
                Case MarkType.编号4
                    pobjMarks(pobjMarks.Count).内容 = 4
                Case MarkType.编号5
                    pobjMarks(pobjMarks.Count).内容 = 5
                Case MarkType.编号6
                    pobjMarks(pobjMarks.Count).内容 = 6
            End Select
        End If
        '点集没有留空
        Set lTemp = New DicomLabel
        lTemp.Left = X
        lTemp.Top = Y
        lTemp.Width = 20
        lTemp.Height = 20
        lTemp.ImageTied = True
        lTemp.Rescale dcmMark.Images(1)
        pobjMarks(pobjMarks.Count).X1 = lTemp.Left / mdblMarkZoom
        pobjMarks(pobjMarks.Count).Y1 = lTemp.Top / mdblMarkZoom
        pobjMarks(pobjMarks.Count).X2 = pobjMarks(pobjMarks.Count).X1
        pobjMarks(pobjMarks.Count).Y2 = pobjMarks(pobjMarks.Count).Y1
        pobjMarks(pobjMarks.Count).填充色 = lngColor(pobjMarks.Count Mod 9 + 1)
        pobjMarks(pobjMarks.Count).填充方式 = -2
        '线条色留空，字体色留空
        pobjMarks(pobjMarks.Count).线型 = 1
        pobjMarks(pobjMarks.Count).线宽 = 1
        Set pobjMarks(pobjMarks.Count).字体 = New StdFont '  "宋体"
        drawPicMarks dcmMark.Images(1), pobjMarks
        dcmMark.Refresh
        
        pMarkModified = True
    End If
End Sub

Private Sub dcmMark_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If dcmMark.Images.Count = 1 Then
        '设置鼠标
        If dcmMark.ImageXPosition(X, Y) > 0 And dcmMark.ImageXPosition(X, Y) < dcmMark.Images(1).SizeX _
           And dcmMark.ImageYPosition(X, Y) > 0 And dcmMark.ImageYPosition(X, Y) < dcmMark.Images(1).SizeY Then
            picMark.MousePointer = 99
            picMark.MouseIcon = listCur.ListImages("Pen").Picture
        Else
            picMark.MousePointer = 0
            picMark.MouseIcon = Nothing
        End If
    End If
End Sub

Private Sub dcmMark_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 2 Then ShowPopupMark
End Sub

Private Sub OpenImageProcessWind()
    Dim i As Long
    
    If Not mSelMiniImg Is Nothing Then
        For i = 1 To mSelMiniImg.Labels.Count
            If mSelMiniImg.Labels(i).tag = "SELECT" Or mSelMiniImg.Labels(i).tag = "BORDER" Then
                mSelMiniImg.Labels(i).Visible = False
            End If
        Next
    End If
        
    If mobjImageProcess Is Nothing Then
        Set mobjImageProcess = New zl9PacsControl.clsImageProcess
    
    End If
    
    mobjImageProcess.ShowImageProcess mlngAdviceID, mSelMiniImg, (ucMiniImageViewer.PageNumber - 1) * ucMiniImageViewer.PageImgCount + mintCurImgIndex, Me, mblnMoved, 0
'    Call frmReportImageEdit.zlShowMe(mSelMiniImg, Me, mintCurImgIndex, mSelViewerIndex, mlngModule)
    
    If Not mSelMiniImg Is Nothing Then
        For i = 1 To mSelMiniImg.Labels.Count
            mSelMiniImg.Labels(i).Visible = True
        Next
    End If
End Sub

Public Sub DcmAddImage(dcmImage As DicomImage, SelViewerIndex As Integer)
'把当前图形添加到图象框中
    Dim i As Integer
    
    If Not dcmImage Is Nothing Then
        For i = 1 To dcmImage.Labels.Count
            If dcmImage.Labels(i).tag = "SELECT" Or dcmImage.Labels(i).tag = "BORDER" Then
                dcmImage.Labels(i).Visible = False
            End If
        Next
        
        dcmReportImage(SelViewerIndex).Images.Add dcmImage
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).BorderColour = vbWhite
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
        dcmReportImage(SelViewerIndex).CurrentIndex = 1
        Call picReportImage_Resize
        pImageModified = True
        
        For i = 1 To dcmImage.Labels.Count
            dcmImage.Labels(i).Visible = True
        Next
    End If
End Sub

Public Sub DcmAddXWImage(dcmImage As DicomImage)
'把当前图形添加到图象框中
    Dim i As Integer
    
    If Not dcmImage Is Nothing Then
        For i = 1 To dcmImage.Labels.Count
            dcmImage.Labels(i).Visible = False
        Next
        
        dcmReportImage(mSelViewerIndex).Images.Add dcmImage
        dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).BorderColour = vbWhite
        dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
        dcmReportImage(mSelViewerIndex).CurrentIndex = 1
        Call picReportImage_Resize
        pImageModified = True
        
        For i = 1 To dcmImage.Labels.Count
            dcmImage.Labels(i).Visible = True
        Next
    End If
End Sub

Private Sub dcmReportImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    
'    If dcmReportImage(Index).Images.Count = 0 Then Exit Sub
    
    mSelViewerIndex = Index
    mselReportImgIndex = dcmReportImage(Index).ImageIndex(X, Y)
    
    For i = 1 To dcmReportImage.Count - 1
        dcmReportImage(i).Labels(1).ForeColour = vbWhite
        dcmReportImage(i).Refresh
    Next i
    dcmReportImage(Index).Labels(1).ForeColour = vbRed
    dcmReportImage(Index).Refresh
    
    If mselReportImgIndex <> 0 Then
        For i = 1 To dcmReportImage(Index).Images.Count
            dcmReportImage(Index).Images(i).BorderColour = vbWhite
        Next i
        dcmReportImage(Index).Images(mselReportImgIndex).BorderColour = vbBlue
    End If
    
    
End Sub

Private Sub dcmReportImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If Button = 2 Then Call ShowPopupImage(0)
End Sub

Private Sub ShowPopupCache()

End Sub

Private Sub ShowPopupImage(ByVal intType As Integer)
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'intType:0--报告图，1--缩略图，2--缓存图
'------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    If intType <> 2 Then
        If ucMiniImageViewer.CurImageCount < 1 Then Exit Sub
    End If
    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If intType = 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelImage, "删除")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveUp, "前移")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveDown, "后移")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_SelMiniImage, "提取报告图")
        ElseIf intType = 1 Then
            Set cbrControl = .Add(xtpControlButton, comMenu_Cap_Process, "图像处理")
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "分页设置")
            cbrControl.BeginGroup = True
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImageShield, "屏蔽大图")
            Call GetLocalPar
            cbrControl.Checked = mblnImageShield
            
            If mlngModule = 1291 And mblnUseAfterCapture Then Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SendToAfter, "发送到后台")
            cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "分页设置")
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SendToAdvice, "发送到检查")
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_DeleteImage, "删除")
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_RefreshImg, "刷新")
        End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub ShowPopupMark()
    '------------------------------------------------
'功能：创建鼠标右键弹出菜单
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelMarks, "清除标注")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub DTPImg_Change()
On Error GoTo errH
    mdate = DTPimg.value
    Call LoadMiniCache
    ucMiniCache.RedrawSelf
    Call dkpMain.RedrawPanes
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
End Sub

Private Sub Form_Activate()
    '根据需要加载图像
    
    '注：在Form的Activate和Paint时间中必须调用LoadImages方法
    '因为如果只在Activate方法中调用LoadImages方法，可能造成报告图不会在第一时间显示，必须用鼠标点击一下报告图才会显示
    '如果只在Paint方法中调用LoadImages方法，由于该方法中使用了UnLoad卸载控件数组，可能造成“不能从该上下文中卸载”的错误
    
    Call LoadImages
    Call GetNowTag(True)
End Sub

Private Sub Form_Load()
    
    DTPimg.value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    mdate = DTPimg.value
    Call LoadMiniCache
    
    '标记本次刷新已经加载图像
    blnLoadImages = True
    
    '标记窗体已经首次加载
    mblnIsInitFace = False
        
    mintMoustType = MarkType.自动编号

    
    '设置默认颜色
    lngColor(1) = RGB(186, 186, 186)
    lngColor(2) = RGB(255, 215, 0)
    lngColor(3) = RGB(255, 0, 255)
    lngColor(4) = RGB(255, 0, 130)
    lngColor(5) = RGB(0, 255, 0)
    lngColor(6) = RGB(130, 255, 255)
    lngColor(7) = RGB(255, 255, 0)
    lngColor(8) = RGB(0, 0, 255)
    lngColor(9) = RGB(0, 160, 0)
    
    '定义UIDRoot=1
    mdcmGlobal.RegString("UIDRoot") = "1"
    
    Call InitLoaclParas     '读取本机参数
'    Call InitFaceScheme     '初始化窗体界面
    
    Call RegXWAddReportImgWindow(Me.hWnd, Me)
End Sub


Private Sub InitLoaclParas()
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage"
    End If
    
    ucMiniImageViewer.PageImgCount = Val(GetSetting("ZLSOFT", strRegPath, "报告缩略图数量", 5))
    If mlngModule = 1291 Then
        mblnUseAfterCapture = Val(GetDeptPara(mlngCurDeptId, "启用后台采集", 1, True)) = 1
    Else
        mblnUseAfterCapture = False
    End If
    
    '读取标记图区域，报告图区域 和缩略图区域的高度
    mlngCY1 = GetSetting("ZLSOFT", strRegPath, "CY1", 180)
    mlngMarkW = GetSetting("ZLSOFT", strRegPath, "MarkW", 300)
    mlngCY2 = GetSetting("ZLSOFT", strRegPath, "CY2", 400)
    mlngRptImgW = GetSetting("ZLSOFT", strRegPath, "RptImgW", 100)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 200)
End Sub

Private Sub Form_Paint()
    '根据需要加载图像
    Call LoadImages
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
        
    Call ClearEmptyFolder(True)
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage"
    End If
    
    Call SaveSetting("ZLSOFT", strRegPath, "报告缩略图数量", ucMiniImageViewer.PageImgCount)
    
    '保存标记图区域，报告图区域和缩略图区域的高度
    '285是Pane的标题高度，使用了标题，就需要加回这个高度
    SaveSetting "ZLSOFT", strRegPath, "CY1", picMark.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "MarkW", picMark.Width
    SaveSetting "ZLSOFT", strRegPath, "CY2", picReportImage.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "RptImgW", picReportImage.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", picMiniCache.Height
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX3", Me.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", Me.Height
    
    Call DisRegXWAddReportImgWindow(Me.hWnd)
End Sub

Private Sub mobjImageProcess_AfterSaveStady()
    Call LoadMiniImages
End Sub

Private Sub mobjImageProcess_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    If mstrInstance = dcmImage.InstanceUID Then Exit Sub
    Select Case lngImageType
        Case 0  '标记图
'            Call DcmAddMarkImage(dcmImage)
        Case 1  '报告图
            Call DcmAddImage(dcmImage, mSelViewerIndex)
        Case 2  '检查图
            If mobjPacsCapture Is Nothing Then
                Set mobjPacsCapture = CreateObject("zl9PacsImageCap.clsPacsCapture")
                
                Call mobjPacsCapture.zlInitModule(gcnOracle, glngSys, mlngModule, gstrPrivs, mlngCurDeptId, Me.hWnd, Me, True, gblnUseDebugLog)
            End If
            
            Call mobjPacsCapture.SaveImageToStady(dcmImage, mlngAdviceID)
            
            Set mobjPacsCapture = Nothing
    End Select
    
    mstrInstance = dcmImage.InstanceUID
End Sub

Private Sub mobjImageProcess_OnUnload()
    Set mobjImageProcess = Nothing
End Sub

Private Sub picMark_Resize()
    If picMark.Height = 0 Or picMark.Width = 0 Then Exit Sub
    
    On Error Resume Next
    
    '判断宽高比
    If picMark.Width / picMark.Height > 2 Then  '数字标记放在右边
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = Abs(picMark.ScaleWidth - picNumMark.ScaleWidth - 50)
        dcmMark.Height = picMark.ScaleHeight
        
        picNumMark.Left = dcmMark.Width
        If picMark.Height > picNumMark.Height Then
            picNumMark.Top = (picMark.ScaleHeight - picNumMark.ScaleHeight) / 2
        Else
            picNumMark.Top = 0
        End If
    Else    '数字标记放在下面
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = picMark.ScaleWidth
        dcmMark.Height = Abs(picMark.ScaleHeight - picNumMark.ScaleHeight - 50)
        
        If picMark.Width > picNumMark.Width Then
            picNumMark.Left = (picMark.ScaleWidth - picNumMark.ScaleWidth) / 2
        Else
            picNumMark.Left = 0
        End If
        picNumMark.Top = dcmMark.Height
    End If
End Sub

Private Sub picMiniCache_Resize()
On Error Resume Next
    DTPimg.Left = 0
    DTPimg.Top = 0
    DTPimg.Width = 1400
    DTPimg.Height = 300
    
    cboCache.Left = DTPimg.Width
    cboCache.Top = 0
    cboCache.Width = picMiniCache.ScaleWidth - DTPimg.Width
    cboCache.Height = 300
    
    ucMiniCache.Left = 0
    ucMiniCache.Top = cboCache.Top + cboCache.Height
    ucMiniCache.Width = picMiniCache.ScaleWidth
    ucMiniCache.Height = picMiniCache.ScaleHeight - ucMiniCache.Top
End Sub

Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane
    
    With dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    If mintShowMarkImage = 1 Then
        picMark.Visible = True
        dcmMark.Visible = True
        picNumMark.Visible = True
        
        Set Pane1 = dkpMain.CreatePane(1, mlngMarkW, mlngCY1, DockTopOf, Nothing)
        Pane1.Title = "标记图"
        Pane1.Handle = picMark.hWnd
        Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '根据宽高比，摆放报告图的位置
        If ((mlngCY1 = mlngCY2) And (mlngMarkW + mlngRptImgW > mlngCY1)) _
            Or (((mlngCY1 <> mlngCY2)) And (mlngMarkW + mlngRptImgW > mlngCY1 + mlngCY2)) Then
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockLeftOf, Pane1)
        Else
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockBottomOf, Pane1)
        End If
    Else
        picMark.Visible = False
        dcmMark.Visible = False
        picNumMark.Visible = False
        
        Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockTopOf, Nothing)
    End If
    

    Pane2.Title = "报告图"
    Pane2.Handle = picReportImage.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set Pane3 = dkpMain.CreatePane(3, 0, mlngCY3, DockBottomOf, Nothing)
    Pane3.Title = "缩略图"
    Pane3.Handle = picMiniViewer.hWnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        Set pane4 = dkpMain.CreatePane(4, 0, mlngCY3, DockBottomOf, Nothing)
        pane4.Title = "后台图"
        pane4.Handle = picMiniCache.hWnd
        pane4.Options = PaneNoCloseable Or PaneNoFloatable
        pane4.AttachTo Pane3
        picMiniCache.Visible = True
    Else
        picMiniCache.Visible = False
    End If
    
    Pane3.Selected = True
    
    mblnIsInitFace = True
End Sub

Private Function GetTag(ByVal FolderName As String, ByRef strType As String) As Integer
'解析文件夹名称中的标识号，FolderName：目标目录名，strType： 返回“标识” 或 “检查”
On Error GoTo errH
    Dim i As Integer
    Dim strTmp As String
    
    strType = Mid(FolderName, 1, 2)
    strTmp = Mid(FolderName, 3, Len(FolderName) - 2)
    i = InStr(strTmp, "-")
    GetTag = Val(Mid(strTmp, 1, i - 1))
    
    Exit Function
errH:
    GetTag = 0
End Function

Private Function GetStudyUIDFromFolderName(ByVal FolderName As String) As String
'解析文件夹名称中的检查UID并返回，若出错返回文件夹名
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
    
    i = InStr(FolderName, "-")
    j = Len(FolderName)
    
    GetStudyUIDFromFolderName = Mid(FolderName, i + 1, j - i)
    Exit Function
errH:
    GetStudyUIDFromFolderName = FolderName
End Function

Function LoadMiniCache() As Boolean
    Dim i As Integer
    Dim strQueryPath As String
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strStudyUID As String, strSeriesUID As String
    Dim lngStudyNo As Long, lngSeriesNo As Long

    Dim strAfterTime As String
    Dim dtChose As Date
    Dim intTMP As Integer
    Dim strTag As String  '三位数的标识
    Dim strType As String
    Dim curDate As Date
    
    If mblnUseAfterCapture = False Then Exit Function
    
    curDate = zlDatabase.Currentdate
    DTPimg = mdate

    Set mrsImageCache = New ADODB.Recordset
    mrsImageCache.Fields.Append "姓名", adVarChar, 100
    mrsImageCache.Fields.Append "检查号", adVarChar, 18
    mrsImageCache.Fields.Append "检查UID", adVarChar, 64
    mrsImageCache.Fields.Append "序列号", adVarChar, 18
    mrsImageCache.Fields.Append "序列UID", adVarChar, 64
    mrsImageCache.Fields.Append "检查日期", adVarChar, 20
    mrsImageCache.Fields.Append "路径", adVarChar, 4000
    mrsImageCache.Open
    
    If mobjFile.FolderExists(mstrAfterImgPath) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterImgPath).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterImgPath).SubFolders
            If InStr(objFolder2.Name, Format(mdate, "yyyymmdd")) > 0 Then ''如果不是选择的时间则跳过

                If objFolder2.SubFolders.Count > 0 Then
                    For Each objFolder3 In objFolder2.SubFolders                            '检查UID层
                            
                        If objFolder3.SubFolders.Count >= 0 Then
                            strAfterTime = Format(objFolder3.DateCreated, "YYYY-MM-DD HH:MM:SS")
                            strStudyUID = GetStudyUIDFromFolderName(objFolder3.Name)
                                                                  
                            lngStudyNo = lngStudyNo + 1
                            lngSeriesNo = 0
                                    
                            For Each objFolder4 In objFolder3.SubFolders                    '序列UID层
                                    
                                strSeriesUID = objFolder4.Name
                                lngSeriesNo = lngSeriesNo + 1
                                       
                                mrsImageCache.AddNew
                                strTag = Lpad(GetTag(objFolder3.Name, strType), 3, "0")
                                mrsImageCache!姓名 = strType & strTag
                                mrsImageCache!检查号 = lngStudyNo
                                mrsImageCache!检查uid = strStudyUID
                                mrsImageCache!序列号 = lngSeriesNo
                                mrsImageCache!序列UID = strSeriesUID
                                mrsImageCache!检查日期 = strAfterTime
                                mrsImageCache!路径 = objFolder4.Path
                                mrsImageCache.Update
                            Next
                        End If
                    Next
                                        Exit For '此时已经遍历完所选时间，跳出时间选择
                End If
            End If '时间选择
        Next
    End If
    
    If mrsImageCache.RecordCount > 0 Then
        mrsImageCache.Sort = "检查日期 desc"
        mrsImageCache.MoveFirst
    End If

    cboCache.Clear
    ucMiniCache.ImgViewer.Images.Clear
    
    For i = 0 To mrsImageCache.RecordCount - 1
        If i = 0 Then strQueryPath = Nvl(mrsImageCache!路径)
        
        cboCache.AddItem Nvl(mrsImageCache!姓名) & "     时间：" & Format(Nvl(mrsImageCache!检查日期), "HH:MM:SS")
        mrsImageCache.MoveNext
    Next
    
    If mrsImageCache.RecordCount > 0 Then
        If cboCache.ListIndex < 0 Then
            cboCache.ListIndex = 0
        End If
    End If
    
End Function

Public Function LoadMiniImages() As Boolean
    ucMiniImageViewer.ShowCheckBox = mlngModule <> 1290
    Call ucMiniImageViewer.RefreshImage(slAdvice, mlngAdviceID, mblnMoved, True)
End Function

Private Sub LoadReportImages()
    
    On Error GoTo errH
    
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim i As Integer
    Dim j As Long
    
    '初始化各个对象
    Call ClearReportImages
        
    iRImageCount = 0
    For i = 1 To UBound(mobjImgCTables)
    
        Set cTable = mobjImgCTables(i)
            
        '记录图像所在表格ID
        If pTableID = "" Then
            pTableID = cTable.ID
        Else
            pTableID = pTableID & ";" & cTable.ID
        End If
        
        '创建viewer
        iRImageCount = iRImageCount + 1
        Load dcmReportImage(iRImageCount)
        dcmReportImage(iRImageCount).BorderStyle = 1
        dcmReportImage(iRImageCount).Labels.AddNew
        dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
        dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
        dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
        dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
        dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
        dcmReportImage(iRImageCount).Visible = True
        
        mSelViewerIndex = iRImageCount

        '记录图像框的宽度和高度，该宽高比例用于后续对图像行列布局
        If cTable.ExtendTag <> "" Then
            If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
            Else
                dcmReportImage(iRImageCount).tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
            End If
        Else
            dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
        End If
        
        
        For j = 1 To cTable.Pictures.Count
            strPicFile = App.Path & "\PACSPic" & j & ".JPG"
            If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

            Set oPicture = cTable.Pictures(j).OrigPic
            SavePicture oPicture, strPicFile
            If objFile.FileExists(strPicFile) Then
                '显示标记图和报告图
                If cTable.Pictures(j).PictureType = EPRMarkedPicture And dcmMark.Images.Count = 0 Then

                    '只处理第一个标记图
                    dcmMark.Images.AddNew
                    
                    dcmMark.Images(1).FileImport strPicFile, "BMP"
                    dcmMark.Images(1).tag = cTable.Pictures(j).ID
                    '保存标记图基础数据
                    Set pobjMarks = cTable.Pictures(j).PicMarks
                    pMarkImageID = cTable.Pictures(j).ID

                    mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(j).Width * Screen.TwipsPerPixelX
                    '显示标注
                    If cTable.Pictures(j).PicMarks.Count > 0 Then
                        drawPicMarks dcmMark.Images(1), cTable.Pictures(j).PicMarks
                    End If
                Else

                    dcmReportImage(iRImageCount).Images.AddNew
                    dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).FileImport strPicFile, "BMP"
                    If cTable.Pictures(j).PicName = "" Then
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
                    Else
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).tag = cTable.Pictures(j).PicName
                    End If
                    
                    dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderWidth = 3
                    dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderColour = vbWhite
                    dcmReportImage(iRImageCount).CurrentIndex = 1
                    mselReportImgIndex = 1
                End If
                '删除临时图像
                Kill strPicFile
            End If
        Next j
    Next i
    
    If dcmReportImage.Count > 1 Then dcmReportImage(1).Labels(1).ForeColour = vbRed
    Call picReportImage_Resize
    

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function InitCTables(Optional ByVal lngFormatId As Long = 0) As Boolean
'初始化病历格式中的图像表格对象
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable
    Dim blnGetTable As Boolean
    Dim lngUbound As Long
    Dim i As Long
    
    ReDim mobjImgCTables(0)
    
    InitCTables = False
    
    If lngFormatId <> 0 Then
      '范文格式，查 病历范文内容
        strSQL = "Select Id As 表格Id From 病历范文内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngFormatId)

        If rsTemp.RecordCount > 0 Then
            If rsTemp.RecordCount < dcmReportImage.Count - 1 Then
                If MsgBoxD(Me, "新格式中图象框数量少于当前格式，当前的部分图象框会被删除，是否更换格式？", vbOKCancel) = vbCancel Then
                    Exit Function
                Else
                    '先删除多余的图象框
                    For i = dcmReportImage.Count - 1 - rsTemp.RecordCount To 1 Step -1
                        Unload dcmReportImage(dcmReportImage.Count - 1)
                    Next i
                End If
            End If
        End If
    Else
        '如果存在报告内容，则从报告内容中读取数据，否则从报告单格式中读取数据
        If mlngReportID <> 0 Then
            strSQL = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
                " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
                " Order By 对象序号"
            If mblnMoved = True Then
                strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
            End If
    
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "报告内容中读取", mlngReportID)
        Else
            strSQL = "Select Id As 表格Id From 病历文件结构" & vbNewLine & _
                " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
                " Order By 对象序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "从报告单格式中读取", mlngFileID)
        End If
    End If
    
    
    Do While Not rsTemp.EOF
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_病历文件定义, mlngFileID, Val("" & rsTemp!表格ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_单病历审核, mlngReportID, Val("" & rsTemp!表格ID))
        End If
        
        If blnGetTable Then
            lngUbound = UBound(mobjImgCTables) + 1
            ReDim Preserve mobjImgCTables(lngUbound)
            
            Set mobjImgCTables(lngUbound) = cTable
        End If
        
        Call rsTemp.MoveNext
    Loop
    
    InitCTables = True

End Function


Private Function DecideMarkImagesVisible() As Integer
'------------------------------------------------
'功能：判断当前选中检查标记图是否可见
'参数：无
'返回：int类型，1-显示标记图  2-隐藏标记图
'-----------------------------------------------
    Dim cTable As cEPRTable
    Dim i As Integer
    Dim j As Long
    
    For i = 1 To UBound(mobjImgCTables)
        Set cTable = mobjImgCTables(i)
        
        If cTable.Pictures.Count > 0 Then
            For j = 1 To cTable.Pictures.Count
                If cTable.Pictures(j).PictureType = EPRMarkedPicture Then
                    DecideMarkImagesVisible = 1
                    Exit Function
                Else
                    DecideMarkImagesVisible = 0
                End If
            Next
        Else
            DecideMarkImagesVisible = 0
        End If
    Next i

End Function


Private Sub drawPicMarks(img As DicomImage, thisMarks As cPicMarks)
'显示标注，只支持数字编号标注
    Dim i As Integer
    Dim iLabelCount As Integer
    
    img.Labels.Clear
    For i = 1 To thisMarks.Count
        If thisMarks(i).类型 = 6 Then   '圆形编号
            With thisMarks(i)
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).BackColour = IIf(.填充色 = 0, vbYellow, .填充色)
                img.Labels(iLabelCount).Transparent = False
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True
                
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True

                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelText
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).FontSize = 11
                img.Labels(iLabelCount).FontName = "Arial Bold"
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).AutoSize = True
                img.Labels(iLabelCount).Text = .内容
                img.Labels(iLabelCount).ImageTied = True
            End With
        End If
    Next i
End Sub
 
Private Sub picMiniViewer_Resize()
On Error Resume Next
    ucMiniImageViewer.Left = 0
    ucMiniImageViewer.Top = 0
    ucMiniImageViewer.Width = picMiniViewer.ScaleWidth
    ucMiniImageViewer.Height = picMiniViewer.ScaleHeight
End Sub

Private Sub picReportImage_Resize()
    Dim i As Integer
    Dim rectH As Long, rectW As Long    '图象框可以使用的区域宽高
    Dim picH As Long, picW As Long      '图像实际宽高，作为比例使用
    Dim iCols As Integer, iRows As Integer
    Dim dImg As DicomImage
    
    If dcmReportImage.Count = 1 Then Exit Sub
    
    On Error Resume Next
    
    '首先计算每个图象框可占用的最大宽高
    
    rectH = picReportImage.Height / (dcmReportImage.Count - 1)
    rectW = picReportImage.Width
    If rectH < 100 Or rectW < 100 Then Exit Sub
    
    For i = 1 To dcmReportImage.Count - 1
        '按照图像比例，计算图象框的真实宽度和高度
        picW = Val(Split(dcmReportImage(i).tag, "|")(0))
        picH = Val(Split(dcmReportImage(i).tag, "|")(1))
        
        dcmReportImage(i).Height = rectH - 100
        dcmReportImage(i).Width = rectW - 100
        
        dcmReportImage(i).Left = 0
        dcmReportImage(i).Top = rectH * (i - 1)
        
        dcmReportImage(i).Labels(1).Width = Abs(dcmReportImage(i).Width / Screen.TwipsPerPixelX - 2)
        dcmReportImage(i).Labels(1).Height = Abs(dcmReportImage(i).Height / Screen.TwipsPerPixelY - 1)

        
        '调整图像显示布局
        ResizeRegion dcmReportImage(i).Images.Count, picW, picH, iRows, iCols
        dcmReportImage(i).MultiColumns = iCols
        dcmReportImage(i).MultiRows = iRows
    Next i
End Sub

Public Sub zlChangeFormat(FormatID As Long)

    On Error GoTo errH
    
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim blnHasMarkImage As Boolean
    Dim objFile As New Scripting.FileSystemObject
    Dim i As Integer
    Dim j As Long
    
    '初始化各个对象
    Call ClearReportImages
    
    If InitCTables(FormatID) = True Then
        '读取图象框中的标记图和报告图
        iRImageCount = 0
        pTableID = ""
        
        For i = 1 To UBound(mobjImgCTables)
            Set cTable = mobjImgCTables(i)
            
            iRImageCount = iRImageCount + 1
            If iRImageCount > dcmReportImage.Count - 1 Then
                '创建Viewer
                Load dcmReportImage(iRImageCount)
                dcmReportImage(iRImageCount).BorderStyle = 1
                dcmReportImage(iRImageCount).Labels.AddNew
                dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
                dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
                dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
                dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
                dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
                dcmReportImage(iRImageCount).Visible = True
            End If
            mSelViewerIndex = iRImageCount
            
            '记录图像框的宽度和高度，该宽高比例用于后续对图像行列布局
            If cTable.ExtendTag <> "" Then
                If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                    dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
                Else
                    dcmReportImage(iRImageCount).tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                End If
            Else
                dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
            End If
            
            '更新标记图
            For j = 1 To cTable.Pictures.Count
                strPicFile = App.Path & "\PACSPic" & j & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                Set oPicture = cTable.Pictures(j).OrigPic
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '显示标记图
                    If cTable.Pictures(j).PictureType = EPRMarkedPicture Then
                        blnHasMarkImage = True
                        '先清除当前标记图，再更新
                        dcmMark.Images.Clear
                        dcmMark.Images.AddNew
                        dcmMark.Images(1).FileImport strPicFile, "BMP"
                        dcmMark.Images(1).tag = cTable.Pictures(j).ID
                        '如果当前没有标记，则读取新格式中标记图的标记
                        If pobjMarks Is Nothing Then
                            Set pobjMarks = cTable.Pictures(j).PicMarks
                        End If
                        pMarkImageID = cTable.Pictures(j).ID

                        mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(j).Width * Screen.TwipsPerPixelX
                        '显示标注
                        If pobjMarks.Count > 0 Then
                            drawPicMarks dcmMark.Images(1), pobjMarks
                        End If
                    End If
                    '删除临时图像
                    Kill strPicFile
                End If
            Next j
        Next i
    End If
    
    If blnHasMarkImage = False Then
        '当前格式没有标记图，删除当前显示的标记图
        pMarkImageID = 0
        
        dcmMark.Images.Clear
        If Not (pobjMarks Is Nothing) Then
            For i = 1 To pobjMarks.Count
                pobjMarks.Remove 1
            Next i
        End If
    End If
    Call picReportImage_Resize

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ucMiniCache_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
'    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
'    If Button = 1 Then ucMiniCache.ImgChecked(ucMiniCache.SelectIndex) = Not ucMiniCache.ImgChecked(ucMiniCache.SelectIndex)
End Sub

Private Sub ucMiniCache_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(2)
End Sub

Private Sub ucMiniImageViewer_AfterSaveStudy()
    Call LoadMiniImages
End Sub

Private Sub ucMiniImageViewer_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
    If ucMiniImageViewer.CurImageCount > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '判断当前双击的操作动作
        If mintImageDblClick = 0 Then   '直接写入报告
            Dim dcmImage As DicomImage
            
            Set dcmImage = mSelMiniImg
            
            '调用将当前图形添加到图象框过程
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            
        Else                            '先打开图片编辑窗口
            Call OpenImageProcessWind
        End If

    End If
    
    blnContinue = False
End Sub

Private Sub ucMiniImageViewer_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub

    If Button = 1 And mlngShowBigImg = 3 Then Call OpenImageProcessWind
End Sub

Private Sub ucMiniImageViewer_OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mlngShowBigImg = 1 Then RaiseEvent AfterShowBigImage
End Sub

Private Sub ucMiniImageViewer_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(1)
End Sub

Private Sub ucMiniImageViewer_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    If mstrInstance = dcmImage.InstanceUID Then Exit Sub
    Select Case lngImageType
        Case 0  '标记图
'            Call DcmAddMarkImage(dcmImage)
        Case 1  '报告图
            Call DcmAddImage(dcmImage, mSelViewerIndex)
        Case 2  '检查图
            If mobjPacsCapture Is Nothing Then
                Set mobjPacsCapture = CreateObject("zl9PacsImageCap.clsPacsCapture")
                
                Call mobjPacsCapture.zlInitModule(gcnOracle, glngSys, mlngModule, gstrPrivs, mlngCurDeptId, Me.hWnd, Me, True, gblnUseDebugLog)
            End If
            
            Call mobjPacsCapture.SaveImageToStady(dcmImage, mlngAdviceID)
            
            Set mobjPacsCapture = Nothing
    End Select
    
    mstrInstance = dcmImage.InstanceUID
End Sub

Private Sub ucMiniImageViewer_OnSelChange(ByVal lngSelectedIndex As Long)
    Set mSelMiniImg = ucMiniImageViewer.SelectImage
End Sub

Private Sub LoadImages()
'------------------------------------------------
'功能：加载报告图和缩略图
'参数：
'返回：无，直接加载图像，并修噶 blnLoadImages状态
'-----------------------------------------------
    '如果本次刷新没有加载图像，则加载图像
    If blnLoadImages = False Then
        '读取后台采集的图像
        If mblnUseAfterCapture And mlngModule <> 1290 Then
            Call LoadMiniCache
        End If
        
        '读取和显示当前可选报告图像
        Call LoadMiniImages
        '根据报告单格式，或者报告内容格式，读取标记图和报告图
        Call LoadReportImages
        '标记本次刷新已经加载图像
        blnLoadImages = True
    End If
End Sub

Private Sub ClearEmptyFolder(ByVal blNoReason As Boolean)
'intType 0:发送到检查  1:删除图像
'blNoReason 是否跳过条件执行本过程？用于关闭程序的时候执行
'清空空目录及对应的下拉框，若当前选中的是当天最新标识，则不执行此操作
'首先判断下拉框当前选中是否是当天最新标识
    Dim curTime As Date
    Dim strTime As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim strType As String
    Dim strTag As String
    Dim i As Long
    Dim blDT As Boolean
    Dim blTag As Boolean
    
    On Error GoTo errH
    blDT = False
    blTag = False
    
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        If Not mblDel Then
            Call CheckSendOnImageCountChangedChanged(0)
        Else
            Call CheckSendOnImageCountChangedChanged(1)
        End If
    End If
    mblDel = False
   
    If blNoReason = False Then
        curTime = zlDatabase.Currentdate
        '是当天并且选中的是当前标识，就不进行清空操作

        If (Format(DTPimg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCache.Text, "标识" & Lpad((mintTagMaxTag), 3, "0")) > 0) Then
            Debug.Print "终止清空"
            Exit Sub
        Else
            Debug.Print "继续清空"
        End If
        
    End If
    
    If mobjFile.FolderExists(mstrAfterImgPath) = False Then Exit Sub

    If mobjFile.GetFolder(mstrAfterImgPath).SubFolders.Count > 0 Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterImgPath).SubFolders '''时间
            If objFolder1.Name = Format(curTime, "yyyymmdd") Then blDT = True
            If objFolder1.SubFolders.Count > 0 Then
            
                For Each objFolder2 In objFolder1.SubFolders '''检查uid
                    If InStr(objFolder2.Name, "标识" & mintTagMaxTag) > 0 Then blTag = True
                    If objFolder2.SubFolders.Count > 0 Then
                    
                        For Each objFolder3 In objFolder2.SubFolders '''序列UID
                            If objFolder3.Files.Count = 0 Then
                                '若是当天最新标识则不清空目录
                                If Not (blDT And blTag) Then Call mobjFile.DeleteFolder(objFolder3.Path)
                                
                            End If
                        Next
                        
                        If objFolder2.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder2.Path)
                    Else
                        Call mobjFile.DeleteFolder(objFolder2.Path)
                    End If
                Next
                
                If objFolder1.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder1.Path)
            Else
                Call mobjFile.DeleteFolder(objFolder1.Path)
            End If
        Next
    End If
    
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
End Sub

Private Function GetTodayTagMax(ByVal curDate As Date) As Integer
    '计算当天最大标识
    Dim strDate As String
    Dim intTMP As Integer
    Dim strType As String
    Dim strStudyUID As String
    Dim objFolder2 As Folder, objFolder3 As Folder
    
    On Error GoTo errH
    
    mintTagMax = 1
    mintTagMaxTag = 1
    
    strDate = Format(curDate, "yyyymmdd")
    
    If mobjFile.FolderExists(mstrAfterImgPath) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterImgPath).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterImgPath).SubFolders
            If InStr(objFolder2.Name, strDate) > 0 Then                                 '时间选择
            
                If objFolder2.SubFolders.Count > 0 Then                                  '时间层是否有子目录
                
                    For Each objFolder3 In objFolder2.SubFolders                            '检查UID层
                    
                        If objFolder3.SubFolders.Count > 0 Then

                            strStudyUID = GetStudyUIDFromFolderName(objFolder3.Name)
                            
                            intTMP = GetTag(objFolder3.Name, strType)
                            If intTMP > mintTagMax Then mintTagMax = intTMP
                            
                            If strType = "标识" Then
                                If intTMP > mintTagMaxTag Then mintTagMaxTag = intTMP
                            End If
                            
                        End If

                    Next
                    
                End If '时间层是否有子目录
                
            End If '时间选择
        Next
    End If
    
    GetTodayTagMax = mintTagMax
    
    Exit Function
errH:
    BUGEX "GetTodayTagMax output= -1"
    GetTodayTagMax = -1
End Function

Private Function GetNowTag(ByVal blIsNeedAddOne As Boolean) As Integer
'获得当前标识,blIsNeedAddOne:是否额外+1 如果加载或者初始化，应该+1，发送到后台这种情况不用
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    '更新当天最大标识
    mintTagNow = GetTodayTagMax(curDate)
    
    If blIsNeedAddOne = True Then mintTagNow = mintTagNow + 1

End Function

Public Sub UseAfterImgChanged(ByVal blUse As Boolean)
    Dim objImage As Pane, objCache As Pane, objTmp As Pane
    Dim blHavePane As Boolean
    Dim i As Integer
    
    mblnUseAfterCapture = blUse
    
    If blUse = True Then
    '是否需要创建的判断
        blHavePane = False
        
        For i = 1 To 5
            Set objTmp = dkpMain.FindPane(i)
            
            If Not objTmp Is Nothing Then
            
                If objTmp.Title = "缩略图" Then Set objImage = dkpMain.FindPane(i)
                If objTmp.Title = "后台图" Then blHavePane = True
                
            End If
        Next
        
        If blHavePane = False And Not objImage Is Nothing Then
            picMiniCache.Visible = True
            
            Set objCache = dkpMain.CreatePane(4, 0, 400, DockLeftOf)
            objCache.Title = "后台图"
            objCache.Handle = picMiniCache.hWnd
            objCache.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
                        
            If objImage.Title = "缩略图" Then
                objCache.AttachTo objImage
                objImage.Selected = True
                LoadMiniCache
            End If
        End If
    Else
    '是否需要销毁的判断
        blHavePane = False

        For i = 1 To 5
            Set objTmp = dkpMain.FindPane(i)
            
            If Not objTmp Is Nothing Then
                If objTmp.Title = "后台图" Then
                    blHavePane = True
                    Exit For
                End If
            End If
        Next
    
        If blHavePane = True Then Call dkpMain.DestroyPane(objTmp)
        picMiniCache.Visible = False
        
    End If
    
    Exit Sub
errH:
    Call err.Raise(0, , "后台图标签处理错误" & err.Description)
End Sub

Public Sub initParaForAfterImage(ByVal lngCurDeptId As Long, ByVal lngModule As Long)
    mlngCurDeptId = lngCurDeptId
    mlngModule = lngModule
End Sub

Private Sub SaveLocalPar()
'保存本地参数
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\frmReportImage", "屏蔽大图", IIf(mblnImageShield, 1, 0)
End Sub

Private Sub GetLocalPar()
'读取本地参数

    mblnImageShield = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmReportImage", "屏蔽大图", 0)) = 1
End Sub
