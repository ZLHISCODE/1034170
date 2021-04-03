VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.UserControl ClinicPlanBar 
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14265
   ScaleHeight     =   8610
   ScaleWidth      =   14265
   Begin MSComctlLib.ImageList img11 
      Left            =   5700
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClinicPlanBar.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClinicPlanBar.ctx":050A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClinicPlanBar.ctx":0A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClinicPlanBar.ctx":0F1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDetailedList 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7500
      Left            =   2970
      ScaleHeight     =   7500
      ScaleWidth      =   10845
      TabIndex        =   3
      Top             =   180
      Width           =   10845
      Begin zl9RegEvent.ClinicPlanDetailPages cpdClinicPlanDetailedPag 
         Height          =   4095
         Left            =   810
         TabIndex        =   25
         Top             =   1710
         Width           =   4365
         _ExtentX        =   5900
         _ExtentY        =   5424
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin zl9RegEvent.CustomButton btnLeft 
         Height          =   315
         Left            =   30
         TabIndex        =   24
         Top             =   75
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Picture         =   "ClinicPlanBar.ctx":1428
         BackColor       =   -2147483643
      End
      Begin VB.PictureBox picApply 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   4740
         ScaleHeight     =   825
         ScaleMode       =   0  'User
         ScaleWidth      =   9540
         TabIndex        =   9
         Top             =   60
         Width           =   9540
         Begin VB.PictureBox picApplyWeek 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   930
            ScaleHeight     =   255
            ScaleWidth      =   6765
            TabIndex        =   28
            Top             =   390
            Width           =   6765
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "周一"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "周二"
               Height          =   180
               Index           =   1
               Left            =   947
               TabIndex        =   34
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "周三"
               Height          =   180
               Index           =   2
               Left            =   1894
               TabIndex        =   33
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "周四"
               Height          =   180
               Index           =   3
               Left            =   2841
               TabIndex        =   32
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "周五"
               Height          =   180
               Index           =   4
               Left            =   3788
               TabIndex        =   31
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "周六"
               Height          =   180
               Index           =   5
               Left            =   4735
               TabIndex        =   30
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "周日"
               Height          =   180
               Index           =   6
               Left            =   5685
               TabIndex        =   29
               Top             =   30
               Width           =   690
            End
         End
         Begin VB.PictureBox picApplyRule 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   900
            ScaleHeight     =   345
            ScaleWidth      =   7905
            TabIndex        =   14
            Top             =   0
            Width           =   7905
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "当前"
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   75
               Value           =   -1  'True
               Width           =   705
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "单日"
               Height          =   240
               Index           =   1
               Left            =   735
               TabIndex        =   7
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "双日"
               Height          =   240
               Index           =   2
               Left            =   1530
               TabIndex        =   8
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "星期"
               Height          =   240
               Index           =   3
               Left            =   2325
               TabIndex        =   11
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "轮循"
               Height          =   240
               Index           =   4
               Left            =   3150
               TabIndex        =   12
               Top             =   75
               Width           =   735
            End
            Begin VB.Frame fraLoopSkip 
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Height          =   330
               Left            =   3975
               TabIndex        =   15
               Top             =   30
               Width           =   3735
               Begin VB.TextBox txtSkip 
                  Height          =   285
                  Left            =   2820
                  Locked          =   -1  'True
                  TabIndex        =   17
                  Text            =   "7"
                  Top             =   15
                  Width           =   330
               End
               Begin VB.ComboBox cboDays 
                  Height          =   300
                  Left            =   1110
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   0
                  Width           =   1260
               End
               Begin MSComCtl2.UpDown updSkip 
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   18
                  Top             =   15
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   503
                  _Version        =   393216
                  Value           =   1
                  BuddyControl    =   "txtSkip"
                  BuddyDispid     =   196618
                  OrigLeft        =   3225
                  OrigTop         =   15
                  OrigRight       =   3480
                  OrigBottom      =   300
                  Max             =   30
                  Min             =   1
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.Label lblLoopSkipDays 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "间隔       天"
                  Height          =   180
                  Left            =   2445
                  TabIndex        =   20
                  Top             =   60
                  Width           =   1170
               End
               Begin VB.Label lblLoopDate 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "开始轮循日期"
                  Height          =   180
                  Left            =   0
                  TabIndex        =   19
                  Top             =   60
                  Width           =   1080
               End
            End
         End
         Begin VB.Label lblApp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "应用于(&Y)"
            Height          =   180
            Left            =   15
            TabIndex        =   10
            Top             =   90
            Width           =   810
         End
      End
      Begin zl9RegEvent.CustomButton btnRight 
         Height          =   315
         Left            =   390
         TabIndex        =   23
         Top             =   75
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Picture         =   "ClinicPlanBar.ctx":1932
         BackColor       =   -2147483643
      End
      Begin VB.Label lblTittle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "星期二"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   13
         Top             =   75
         Width           =   990
      End
   End
   Begin VB.PictureBox picSouceList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   2910
      Left            =   90
      ScaleHeight     =   2910
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   5895
      Width           =   3195
      Begin zl9RegEvent.ShowSourceInfor SourceInfor 
         Height          =   1635
         Left            =   450
         TabIndex        =   21
         Top             =   540
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2884
         BackColor       =   16773091
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpSourceLine 
         BackColor       =   &H00ECEDC2&
         BorderColor     =   &H80000003&
         Height          =   2295
         Left            =   30
         Top             =   30
         Width           =   3150
      End
      Begin VB.Label lblSourceTittle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号源信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   5
         Top             =   90
         Width           =   780
      End
      Begin VB.Image imgSignalSource 
         Height          =   240
         Left            =   45
         Picture         =   "ClinicPlanBar.ctx":1E3C
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox picWorkTimeList 
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   90
      ScaleHeight     =   2400
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   3450
      Width           =   3195
      Begin VB.PictureBox picLvwBack 
         BackColor       =   &H00FFEFE3&
         BorderStyle     =   0  'None
         Height          =   1245
         Left            =   330
         ScaleHeight     =   1245
         ScaleWidth      =   2475
         TabIndex        =   26
         Top             =   510
         Width           =   2475
         Begin MSComctlLib.ListView lvwWorkTime 
            Height          =   1035
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   1826
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16773091
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "时间段"
               Object.Width           =   9596
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "开始时间"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "终止时间"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Shape shpWorkLine 
         BackColor       =   &H00FFEFE3&
         BorderColor     =   &H80000003&
         Height          =   2295
         Left            =   30
         Top             =   30
         Width           =   3150
      End
      Begin VB.Image imgWork 
         Height          =   240
         Left            =   60
         Picture         =   "ClinicPlanBar.ctx":23C6
         Top             =   75
         Width           =   240
      End
      Begin VB.Label lblCalendbarTittle 
         BackStyle       =   0  'Transparent
         Caption         =   "上班时段"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   4
         Top             =   90
         Width           =   810
      End
   End
   Begin VB.PictureBox picDateList 
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   150
      ScaleHeight     =   3060
      ScaleWidth      =   3045
      TabIndex        =   0
      Top             =   420
      Width           =   3045
      Begin zl9RegEvent.CalendarSel cldsCalenbarSel 
         Height          =   2085
         Left            =   180
         TabIndex        =   22
         Top             =   390
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   3678
         BackColor       =   16773091
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowStyle       =   2
      End
      Begin VB.Shape shpItemSel 
         BorderColor     =   &H80000003&
         Height          =   2625
         Left            =   0
         Top             =   0
         Width           =   3000
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "ClinicPlanBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'缺省属性值
Const m_def_Enabled = 0
'属性变量:
Dim m_Enabled As Boolean

Public Enum Pancel_Index
    Pan_日历 = 1001
    Pan_时间段 = 1002
    Pan_号源 = 1003
    Pan_详情 = 1004
End Enum
Private mobj出诊安排 As 出诊安排
Private mstrCurDay As String

Public Function LoadData(ByVal obj出诊安排 As 出诊安排) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载出诊安排
    '入参:obj出诊记录集-出诊记录集
    '出参:
    '返回:加载成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Err = 0: On Error GoTo Errhand:
    Set mobj出诊安排 = obj出诊安排
    mstrCurDay = ""
    blnOK = InitData
    LoadData = blnOK
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 
Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objListItem As ListItem
    Dim objTemp As 出诊记录集, obj上班时段 As 上班时段
    Dim dtCur As Date
    
    Err = 0: On Error GoTo Errhand:
    '出诊时间
    lvwWorkTime.ListItems.Clear
'    lvwWorkTime.ColumnHeaders(1).Width = 2000
    lvwWorkTime.View = lvwReport
    For Each obj上班时段 In mobj出诊安排.所有上班时段
        Set objListItem = lvwWorkTime.ListItems.Add(, , obj上班时段.时间段 & "(" & Format(obj上班时段.开始时间, "hh:mm") & "-" & Format(obj上班时段.结束时间, "hh:mm") & ")")
        objListItem.SubItems(1) = obj上班时段.开始时间
        objListItem.SubItems(2) = obj上班时段.结束时间
        objListItem.Tag = obj上班时段.时间段
    Next
    If mobj出诊安排.更新合作单位 Then picWorkTimeList.Enabled = False
    
    '号源信息
    SourceInfor.LoadData mobj出诊安排.出诊号源
    
    cldsCalenbarSel.LoadData mobj出诊安排
    
    '轮询开始日期
    If cldsCalenbarSel.ShowStyle = Show_Plan_Day And cboDays.Enabled Then
        If mobj出诊安排.排班方式 = 1 Then
            dtCur = mobj出诊安排.开始时间
            cboDays.Clear
            Do While True
                cboDays.AddItem Format(dtCur, "yyyy/mm/dd")
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
            Loop
            If cboDays.ListCount > 0 Then cboDays.ListIndex = 0
        End If
    End If
    
    Call LoadDetailData
    
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadDetailData()
    Dim objTemp As 出诊记录集
    
    On Error GoTo Errhand
    '当前项目
    If mobj出诊安排.Count = 0 Then
        Set objTemp = New 出诊记录集
    Else
        Set objTemp = mobj出诊安排(1).Clone
    End If
    
    mstrCurDay = objTemp.出诊日期
    Call SetTitleText
    
    '时间段
    CheckWorkTime objTemp
    '安排
    cpdClinicPlanDetailedPag.LoadData objTemp
    
    If IsDate(mstrCurDay) Then
        If DateDiff("d", mstrCurDay, Now) > 0 Then
            lvwWorkTime.Enabled = False
            cpdClinicPlanDetailedPag.Enabled = False
        Else
            lvwWorkTime.Enabled = m_Enabled And True
            cpdClinicPlanDetailedPag.Enabled = m_Enabled And True
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckWorkTime(ByVal obj出诊记录集 As 出诊记录集) As Boolean
    '选择时间段
    Dim objListItem As ListItem, i As Integer
    Dim objItem As 出诊记录
    
    On Error GoTo Errhand
    For i = 1 To lvwWorkTime.ListItems.Count
        lvwWorkTime.ListItems(i).Checked = False
    Next
    If obj出诊记录集 Is Nothing Then Exit Function
    For Each objItem In obj出诊记录集
        If objItem.时间段 <> "" Then
            For i = 1 To lvwWorkTime.ListItems.Count
                If objItem.时间段 = lvwWorkTime.ListItems(i).Tag Then
                    lvwWorkTime.ListItems(i).Checked = True
                    Exit For
                End If
            Next
        End If
    Next
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Docking控件
    '编制:刘兴洪
    '日期:2016-01-08 14:34:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    Dim strReg As String
    Dim panThis As Pane, panLeft As Pane
    
    On Error GoTo Errhand
    sngWidth = picDateList.Width / Screen.TwipsPerPixelX
    sngHeight = picDateList.Height / Screen.TwipsPerPixelY
    
    Set panLeft = dkpMain.CreatePane(Pancel_Index.Pan_日历, sngWidth, sngHeight, DockTopOf, Nothing)
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Title = "": panLeft.Tag = Pancel_Index.Pan_日历
    panLeft.handle = picDateList.hWnd
    
    panLeft.MinTrackSize.Height = sngHeight
    panLeft.MaxTrackSize.Height = sngHeight
    panLeft.MaxTrackSize.Width = sngWidth
    panLeft.MinTrackSize.Width = sngWidth * 2 / 3
    
    
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_详情, sngWidth, 300, DockRightOf, panLeft)
    panThis.Title = ""
    panThis.Tag = Pancel_Index.Pan_详情
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.handle = picDetailedList.hWnd
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_时间段, sngWidth, 300, DockBottomOf, panLeft)
    panThis.Title = "上班时间"
    panThis.Tag = Pancel_Index.Pan_时间段
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.handle = picWorkTimeList.hWnd
    
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_号源, sngWidth, 300, DockBottomOf, panThis)
    panThis.Title = "当前号源信息"
    panThis.Tag = Pancel_Index.Pan_号源
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.handle = picSouceList.hWnd
     
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Call picDateList_Resize
    'Set dkpMain.PaintManager.CaptionFont = use.Font
    
    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub btnRight_Click()
    Dim intIndex As Integer
    
    On Error GoTo Errhand
    If Not mobj出诊安排 Is Nothing Then
        If IsDate(mstrCurDay) Then
            If DateDiff("d", mobj出诊安排.开始时间, DateAdd("d", 1, mstrCurDay)) >= 0 _
                And DateDiff("d", DateAdd("d", 1, mstrCurDay), mobj出诊安排.终止时间) >= 0 Then
                mstrCurDay = Format(DateAdd("d", 1, mstrCurDay), "yyyy-mm-dd")
            End If
        Else
            If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj出诊安排.排班规则 <> 1 Then
                If mobj出诊安排.排班规则 = 6 Then
                    If Val(mstrCurDay) + 1 <= 31 Then mstrCurDay = Val(mstrCurDay) + 1 & "日"
                End If
            Else '星期
                intIndex = GetWeekIndex(mstrCurDay)
                If intIndex + 1 >= 0 And intIndex + 1 <= 6 Then
                    mstrCurDay = GetWeekName(intIndex + 1)
                End If
            End If
        End If
    End If
    Call SetButtonEnabled
    
    If mobj出诊安排 Is Nothing Then Exit Sub
    Call CurPlanChanged(mstrCurDay)
    cldsCalenbarSel.LoadData mobj出诊安排
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetButtonEnabled()
    '设置上一个、下一个按钮可用状态
    Dim intIndex As Integer
    
    On Error GoTo Errhand
    If mobj出诊安排 Is Nothing Then
        btnLeft.Enabled = False
        btnRight.Enabled = False
    Else
        If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj出诊安排.排班规则 <> 1 Then
            '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
            If mobj出诊安排.排班规则 = 6 Then
                btnLeft.Enabled = True
                btnRight.Enabled = True
                If Val(mstrCurDay) <= 1 Or mobj出诊安排.已保存出诊安排.Count = 0 Then
                    btnLeft.Enabled = False
                End If
                If Val(mstrCurDay) >= 31 Or mobj出诊安排.已保存出诊安排.Count = 0 Then
                    btnRight.Enabled = False
                End If
            Else
                btnLeft.Enabled = False
                btnRight.Enabled = False
            End If
        Else
            btnLeft.Enabled = True
            btnRight.Enabled = True
            If IsDate(mstrCurDay) Then '日期
                If DateDiff("d", mobj出诊安排.开始时间, mstrCurDay) <= 0 Then
                    btnLeft.Enabled = False
                End If
                If DateDiff("d", mobj出诊安排.终止时间, mstrCurDay) >= 0 Then
                    btnRight.Enabled = False
                End If
            Else '星期
                intIndex = GetWeekIndex(mstrCurDay)
                If intIndex <= 0 Then
                    btnLeft.Enabled = False
                End If
                If intIndex >= 6 Then
                    btnRight.Enabled = False
                End If
            End If
        End If
    End If
    
    Set btnLeft.Picture = img11.ListImages(IIf(btnLeft.Enabled, 1, 3)).Picture
    Set btnRight.Picture = img11.ListImages(IIf(btnRight.Enabled, 2, 4)).Picture
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cldsCalenbarSel_SelectedChanged()
    Call LoadDetailData
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pancel_Index.Pan_日历
        Item.handle = picDateList.hWnd
    Case Pancel_Index.Pan_时间段
        Item.handle = picWorkTimeList.hWnd
    Case Pancel_Index.Pan_号源
        Item.handle = picSouceList.hWnd
    Case Pancel_Index.Pan_详情
        Item.handle = picDetailedList.hWnd
    End Select
End Sub

Private Sub lvwWorkTime_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, obj出诊安排 As 出诊安排
    Dim blnChecked As Boolean, objItem As 出诊安排, objTemp As 出诊安排
    Dim lngFindIndex As Long
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    
    On Error GoTo Errhand
    blnChecked = Item.Checked
    Item.Checked = Not blnChecked
    If mobj出诊安排.Count = 0 Then
        MsgBox IIf(cldsCalenbarSel.ShowStyle = Show_Plan_Rule, "出诊规则未选择！", "出诊日期未选择！"), vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If blnChecked Then
        '限制上班时间段，不能有交叉
        Dim dtCurStart As Date, dtCurEnd As Date
        Dim dtStart As Date, dtEnd As Date
        dtStart = Item.SubItems(1): dtEnd = Item.SubItems(2)
        If DateDiff("n", dtStart, dtEnd) <= 0 Then dtEnd = DateAdd("d", 1, dtEnd)
        For i = 1 To lvwWorkTime.ListItems.Count
            If lvwWorkTime.ListItems(i).Checked Then
                dtCurStart = lvwWorkTime.ListItems(i).SubItems(1): dtCurEnd = lvwWorkTime.ListItems(i).SubItems(2)
                If DateDiff("n", dtCurStart, dtCurEnd) <= 0 Then dtCurEnd = DateAdd("d", 1, dtCurEnd)
                
                If Not (DateDiff("n", dtCurStart, dtEnd) <= 0 Or DateDiff("n", dtCurEnd, dtStart) >= 0) Then
                    MsgBox "当前上班时段的时间范围与已选择上班时段【" & Left(lvwWorkTime.ListItems(i).Text, InStr(lvwWorkTime.ListItems(i).Text, "(") - 1) & "】的时间范围有重叠，不能同时选择！", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        Next
    End If
    Item.Checked = blnChecked
    
    mobj出诊安排.RemoveAll
    Set obj出诊记录集 = cpdClinicPlanDetailedPag.Get出诊记录集
    mobj出诊安排.AddItem obj出诊记录集, "K" & obj出诊记录集.出诊日期
    
    If Item.Checked Then
        Set obj出诊记录集 = mobj出诊安排(1).Clone
        mobj出诊安排.RemoveAll

        With mobj出诊安排.出诊号源
            Set obj出诊记录 = New 出诊记录
            obj出诊记录.时间段 = Item.Tag
            Set obj出诊记录.上班时段 = mobj出诊安排.所有上班时段("K" & obj出诊记录.时间段).Clone
            obj出诊记录.是否分时段 = .是否分时段
            obj出诊记录.是否序号控制 = .是否序号控制
            obj出诊记录.预约控制 = .预约控制
            obj出诊记录.分诊方式 = .分诊方式
            obj出诊记录.更新合作单位 = mobj出诊安排.更新合作单位

            Set obj出诊记录.安排门诊诊室集.所有分诊诊室 = .分诊诊室集.所有分诊诊室.Clone
            obj出诊记录.安排门诊诊室集.分诊方式 = .分诊方式
            Set obj出诊记录.安排门诊诊室集 = .分诊诊室集.Clone

            Set obj出诊记录.号序信息集 = New 号序信息集
            Set obj出诊记录.号序信息集.上班时段 = obj出诊记录.上班时段
            obj出诊记录.号序信息集.时间段 = obj出诊记录.时间段
            obj出诊记录.号序信息集.是否分时段 = .是否分时段
            obj出诊记录.号序信息集.是否序号控制 = .是否序号控制
            obj出诊记录.号序信息集.预约控制 = .预约控制
            obj出诊记录.号序信息集.出诊频次 = .出诊频次

            Set obj出诊记录.合作单位控制集 = New 合作单位控制集
            Set obj出诊记录.合作单位控制集.号序信息集 = obj出诊记录.号序信息集.Clone
            Set obj出诊记录.合作单位控制集.所有合作单位 = mobj出诊安排.所有合作单位.Clone

            obj出诊记录集.AddItem obj出诊记录
        End With
        mobj出诊安排.AddItem obj出诊记录集, "K" & obj出诊记录集.出诊日期
    Else
        Set obj出诊记录集 = mobj出诊安排(1)
        For i = 1 To obj出诊记录集.Count
            If obj出诊记录集(i).时间段 = Item.Tag Then
                obj出诊记录集.Remove i: Exit For
            End If
        Next
    End If
    cpdClinicPlanDetailedPag.LoadData mobj出诊安排(1)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picDateList_Resize()
    Err = 0: On Error Resume Next
    With picDateList
        shpItemSel.Top = .ScaleTop
        shpItemSel.Width = .ScaleWidth
        shpItemSel.Left = .ScaleLeft
        shpItemSel.Height = .ScaleHeight
        
        cldsCalenbarSel.Left = .ScaleLeft + 50
        cldsCalenbarSel.Top = .ScaleTop + 50
        cldsCalenbarSel.Width = .ScaleWidth - 100
        cldsCalenbarSel.Height = .ScaleHeight - 100
    End With
End Sub

Private Sub picDetailedList_Resize()
    Err = 0: On Error Resume Next
    With picDetailedList
        btnLeft.Left = .ScaleLeft
        btnRight.Left = btnLeft.Left + btnLeft.Width + 20
        lblTittle.Left = btnRight.Left + btnRight.Width + 50
        
        picApply.Left = 3500
        picApply.Top = lblTittle.Top
        picApply.Width = .ScaleWidth - picApply.Left
        
        cpdClinicPlanDetailedPag.Left = .ScaleLeft
        cpdClinicPlanDetailedPag.Top = .ScaleTop + IIf(picApply.Visible, picApply.Top + picApply.Height, lblTittle.Top + lblTittle.Height) + 50
        cpdClinicPlanDetailedPag.Width = .ScaleWidth
        cpdClinicPlanDetailedPag.Height = .ScaleHeight - cpdClinicPlanDetailedPag.Top + 30
    End With
End Sub

Private Sub picLvwBack_Resize()
    Err = 0: On Error Resume Next
    lvwWorkTime.Move 0, 0, picLvwBack.ScaleWidth, picLvwBack.ScaleHeight + 240
End Sub

Private Sub picWorkTimeList_Resize()
    Err = 0: On Error Resume Next
    With picWorkTimeList
        shpWorkLine.Top = .ScaleTop
        shpWorkLine.Width = .ScaleWidth
        shpWorkLine.Left = .ScaleLeft
        shpWorkLine.Height = .ScaleHeight
        
        picLvwBack.Left = lblCalendbarTittle.Left
        picLvwBack.Top = lblCalendbarTittle.Top + lblCalendbarTittle.Height + 50
        picLvwBack.Width = .ScaleWidth - picLvwBack.Left - 50
        picLvwBack.Height = .ScaleHeight - picLvwBack.Top - 50
    End With
End Sub

Private Sub picSouceList_Resize()
    Err = 0: On Error Resume Next
    With picSouceList
        shpSourceLine.Top = .ScaleTop
        shpSourceLine.Width = .ScaleWidth
        shpSourceLine.Left = .ScaleLeft
        shpSourceLine.Height = .ScaleHeight
        
        SourceInfor.Left = lblSourceTittle.Left
        SourceInfor.Top = lblSourceTittle.Top + lblSourceTittle.Height + 50
        SourceInfor.Width = .ScaleWidth - SourceInfor.Left - 50
        SourceInfor.Height = .ScaleHeight - SourceInfor.Top - 50
    End With
End Sub

Private Sub UserControl_Initialize()
    Call InitPanel
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get ShowStyle() As gShowStyle
    ShowStyle = cldsCalenbarSel.ShowStyle
End Property

Public Property Let ShowStyle(ByVal New_ShowStyle As gShowStyle)
    cldsCalenbarSel.ShowStyle = New_ShowStyle
    picApply.Visible = New_ShowStyle = Show_Plan_Day And m_Enabled
End Property

Private Sub btnLeft_Click()
    Dim intIndex As Integer
    
    On Error GoTo Errhand
    If Not mobj出诊安排 Is Nothing Then
        If IsDate(mstrCurDay) Then
            If DateDiff("d", mobj出诊安排.开始时间, DateAdd("d", -1, mstrCurDay)) >= 0 _
                And DateDiff("d", DateAdd("d", -1, mstrCurDay), mobj出诊安排.终止时间) >= 0 Then
                mstrCurDay = Format(DateAdd("d", -1, mstrCurDay), "yyyy-mm-dd")
            End If
        Else
            If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj出诊安排.排班规则 <> 1 Then
                If mobj出诊安排.排班规则 = 6 Then
                    If Val(mstrCurDay) - 1 > 0 Then mstrCurDay = Val(mstrCurDay) - 1 & "日"
                End If
            Else '星期
                intIndex = GetWeekIndex(mstrCurDay)
                If intIndex - 1 >= 0 And intIndex - 1 <= 6 Then
                    mstrCurDay = GetWeekName(intIndex - 1)
                End If
            End If
        End If
    End If
    Call SetButtonEnabled
    
    If mobj出诊安排 Is Nothing Then Exit Sub
    Call CurPlanChanged(mstrCurDay)
    cldsCalenbarSel.LoadData mobj出诊安排
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optRule_Click(index As Integer)
    '功能:设置应用于的显示
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler:
    fraLoopSkip.Visible = False
    picApplyWeek.Visible = False
    If index = 3 Then '按星期
        picApplyWeek.Visible = True
        picApplyWeek.Top = picApplyRule.Top + picApplyRule.Height
        picApply.Height = picApplyWeek.Top + picApplyWeek.Height
        For i = chkWeek.LBound To chkWeek.UBound
            chkWeek(i).Value = vbUnchecked
        Next
    Else
        picApply.Height = picApplyRule.Top + picApplyRule.Height
    End If
    fraLoopSkip.Visible = index = 4
    picDetailedList_Resize
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetTitleText()
    Dim strTemp As String, i As Integer
    
    On Error GoTo Errhand
    lblTittle.Caption = "无安排"
    Select Case cldsCalenbarSel.ShowStyle
    Case Show_Plan_Rule
        picApply.Visible = m_Enabled
        picApplyRule.Visible = False
        picApplyWeek.Visible = False
        If mobj出诊安排.排班规则 = 1 Then
            picApplyWeek.Visible = True
            picApplyWeek.Top = picApplyRule.Top + 60
            picApply.Height = picApplyWeek.Top + picApplyWeek.Height
            For i = chkWeek.LBound To chkWeek.UBound
                chkWeek(i).Value = vbUnchecked
            Next
        Else
            picApply.Visible = False
        End If
    Case Show_Plan_Week
        picApply.Visible = m_Enabled
        picApplyRule.Visible = False
        picApplyWeek.Visible = True
        picApplyWeek.Top = picApplyRule.Top + 60
        picApply.Height = picApplyWeek.Top + picApplyWeek.Height
    Case Show_Plan_Day
        picApply.Visible = m_Enabled
        If Not mobj出诊安排 Is Nothing Then
            If mobj出诊安排.排班方式 = 1 Then '按月
                If optRule(0).Value Then optRule(1).Value = True
                optRule(0).Value = True
            ElseIf mobj出诊安排.排班方式 = 2 Then '按周
                picApplyRule.Visible = False
                picApplyWeek.Visible = True
                picApplyWeek.Top = picApplyRule.Top + 60
                picApply.Height = picApplyWeek.Top + picApplyWeek.Height
            End If
        End If
    End Select
    Call SetButtonEnabled
    
    If mstrCurDay <> "" Then
        '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
        Select Case mobj出诊安排.排班规则
        Case 2
            strTemp = "按单日出诊"
        Case 3
            strTemp = "按双日出诊"
        Case 4, 5
            strTemp = "按" & Val(mstrCurDay) & "天轮循"
        Case 6
            strTemp = ""
            If mobj出诊安排.已保存出诊安排.排班规则 = 6 Then
                For i = 1 To mobj出诊安排.已保存出诊安排.Count
                    strTemp = strTemp & "," & mobj出诊安排.已保存出诊安排(i).出诊日期
                Next
            End If
            If mobj出诊安排.更新合作单位 = False Then
                strTemp = strTemp & "," & mstrCurDay
                If mobj出诊安排.应用于 <> "" Then
                    strTemp = strTemp & "," & mobj出诊安排.应用于
                End If
            End If
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            strTemp = "按每月的" & ZlNumStrSort(strTemp, True) & "日固定出诊"
        Case Else
            strTemp = mstrCurDay
        End Select
    End If
        
    If strTemp = "" Then strTemp = "无安排"
    lblTittle.Caption = IIf(IsDate(strTemp), Format(strTemp, "yyyy-mm-dd"), strTemp)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function IsValied() As Boolean
    '检查数据
    Dim blnSelected As Boolean
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    For i = 1 To lvwWorkTime.ListItems.Count
        If lvwWorkTime.ListItems(i).Checked Then
            blnSelected = True: Exit For
        End If
    Next
    If blnSelected = False Then
        MsgBox "未设置上班时段！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    If cpdClinicPlanDetailedPag.IsValied() = False Then Exit Function
    IsValied = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Property Get Get出诊安排() As 出诊安排
    Dim obj出诊安排 As New 出诊安排
    Dim obj出诊记录集 As 出诊记录集
    
    On Error GoTo Errhand
    Set obj出诊安排 = mobj出诊安排.Clone
    obj出诊安排.RemoveAll
    Set obj出诊记录集 = cpdClinicPlanDetailedPag.Get出诊记录集
    obj出诊安排.AddItem obj出诊记录集, "K" & obj出诊记录集.出诊日期
    If obj出诊安排.排班规则 <> 6 Then
        '模板的特定日期的应用于由选择时已确定
        obj出诊安排.应用于 = GetApplyToStr()
    End If
    Set Get出诊安排 = obj出诊安排
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Private Function GetApplyToStr() As String
    '获取应用于字符串
    Dim strApplyTo As String, i As Integer
    Dim dtCur As Date, str出诊日期 As String
    Dim varTemp As Variant
    
    On Error GoTo Errhand
    If mobj出诊安排 Is Nothing Then Exit Function
    If m_Enabled = False Then Exit Function
    If picApply.Visible = False Then Exit Function
    If mobj出诊安排.Count > 0 Then
        str出诊日期 = mobj出诊安排(1).出诊日期
    End If
    
    If picApplyRule.Visible Then
        If optRule(1).Value Then '单日
            If mobj出诊安排.排班方式 = 1 Then
                dtCur = mobj出诊安排.开始时间
                Do While True
                    If Day(dtCur) Mod 2 = 1 And DateDiff("n", str出诊日期, dtCur) <> 0 Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                    End If
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
                Loop
            End If
        ElseIf optRule(2).Value Then '双日
            If mobj出诊安排.排班方式 = 1 Then
                dtCur = mobj出诊安排.开始时间
                Do While True
                    If Day(dtCur) Mod 2 = 0 And DateDiff("n", str出诊日期, dtCur) <> 0 Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                    End If
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
                Loop
            End If
        ElseIf optRule(3).Value Then '星期
            If picApplyWeek.Visible Then
                dtCur = mobj出诊安排.开始时间
                Do While True
                    For i = chkWeek.LBound To chkWeek.UBound
                        If chkWeek(i).Value = vbChecked Then
                            If Weekday(dtCur, vbMonday) = i + 1 And DateDiff("n", str出诊日期, dtCur) <> 0 Then
                                strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                            End If
                        End If
                    Next
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
                Loop
            End If
        ElseIf optRule(4).Value Then '轮循
            If mobj出诊安排.排班方式 = 1 Then
                If Not (cboDays.ListIndex = -1 Or Val(txtSkip.Text) = 0) Then
                    dtCur = CDate(cboDays) '开始时间
                    Do While True
                        If DateDiff("n", str出诊日期, dtCur) <> 0 Then
                            strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                        End If
                        dtCur = DateAdd("d", Val(txtSkip.Text), dtCur)
                        If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
                    Loop
                End If
            End If
        End If
        If strApplyTo <> "" Then strApplyTo = Mid(strApplyTo, 2)
        GetApplyToStr = strApplyTo
        Exit Function
    End If
    If picApplyWeek.Visible Then
        If cldsCalenbarSel.ShowStyle = Show_Plan_Day Then '按周排班
            dtCur = mobj出诊安排.开始时间
            Do While True
                For i = chkWeek.LBound To chkWeek.UBound
                    If chkWeek(i).Value = vbChecked Then
                        If Weekday(dtCur, vbMonday) = i + 1 And DateDiff("n", str出诊日期, dtCur) <> 0 Then
                            strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                        End If
                    End If
                Next
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
            Loop
        Else
            For i = chkWeek.LBound To chkWeek.UBound
                If chkWeek(i).Value = vbChecked Then
                    If GetWeekName(i) <> str出诊日期 Then
                        strApplyTo = strApplyTo & "," & GetWeekName(i)
                    End If
                End If
            Next
        End If
        If strApplyTo <> "" Then strApplyTo = Mid(strApplyTo, 2)
        GetApplyToStr = strApplyTo
        Exit Function
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    On Error GoTo Errhand
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    SetEnabled UserControl.Controls, New_Enabled
    If btnLeft.Enabled Then
        Set btnLeft.Picture = img11.ListImages(1).Picture: Set btnRight.Picture = img11.ListImages(2).Picture
    Else
        Set btnLeft.Picture = img11.ListImages(3).Picture: Set btnRight.Picture = img11.ListImages(4).Picture
    End If
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

Private Sub UserControl_Terminate()
    Set mobj出诊安排 = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

Private Sub CurPlanChanged(ByVal str出诊日期 As String)
    '当前选择项目改变
    Dim objItem As New 出诊记录集
    
    On Error GoTo Errhand
    mobj出诊安排.RemoveAll
    If mobj出诊安排.已保存出诊安排.Exits("K" & str出诊日期) Then
        mobj出诊安排.AddItem mobj出诊安排.已保存出诊安排("K" & str出诊日期)
    Else
        objItem.出诊日期 = str出诊日期
        mobj出诊安排.AddItem objItem
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
