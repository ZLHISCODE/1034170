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
            Name            =   "����"
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
               Caption         =   "��һ"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "�ܶ�"
               Height          =   180
               Index           =   1
               Left            =   947
               TabIndex        =   34
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   2
               Left            =   1894
               TabIndex        =   33
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   3
               Left            =   2841
               TabIndex        =   32
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   4
               Left            =   3788
               TabIndex        =   31
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   5
               Left            =   4735
               TabIndex        =   30
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H80000005&
               Caption         =   "����"
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
               Caption         =   "��ǰ"
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
               Caption         =   "����"
               Height          =   240
               Index           =   1
               Left            =   735
               TabIndex        =   7
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "˫��"
               Height          =   240
               Index           =   2
               Left            =   1530
               TabIndex        =   8
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   240
               Index           =   3
               Left            =   2325
               TabIndex        =   11
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H80000005&
               Caption         =   "��ѭ"
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
                  Caption         =   "���       ��"
                  Height          =   180
                  Left            =   2445
                  TabIndex        =   20
                  Top             =   60
                  Width           =   1170
               End
               Begin VB.Label lblLoopDate 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ʼ��ѭ����"
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
            Caption         =   "Ӧ����(&Y)"
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
         Caption         =   "���ڶ�"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "��Դ��Ϣ"
         BeginProperty Font 
            Name            =   "����"
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
               Text            =   "ʱ���"
               Object.Width           =   9596
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "��ʼʱ��"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "��ֹʱ��"
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
         Caption         =   "�ϰ�ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
'ȱʡ����ֵ
Const m_def_Enabled = 0
'���Ա���:
Dim m_Enabled As Boolean

Public Enum Pancel_Index
    Pan_���� = 1001
    Pan_ʱ��� = 1002
    Pan_��Դ = 1003
    Pan_���� = 1004
End Enum
Private mobj���ﰲ�� As ���ﰲ��
Private mstrCurDay As String

Public Function LoadData(ByVal obj���ﰲ�� As ���ﰲ��) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���س��ﰲ��
    '���:obj�����¼��-�����¼��
    '����:
    '����:���سɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Err = 0: On Error GoTo Errhand:
    Set mobj���ﰲ�� = obj���ﰲ��
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
    '����:��ʼ������
    '����:���˺�
    '����:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objListItem As ListItem
    Dim objTemp As �����¼��, obj�ϰ�ʱ�� As �ϰ�ʱ��
    Dim dtCur As Date
    
    Err = 0: On Error GoTo Errhand:
    '����ʱ��
    lvwWorkTime.ListItems.Clear
'    lvwWorkTime.ColumnHeaders(1).Width = 2000
    lvwWorkTime.View = lvwReport
    For Each obj�ϰ�ʱ�� In mobj���ﰲ��.�����ϰ�ʱ��
        Set objListItem = lvwWorkTime.ListItems.Add(, , obj�ϰ�ʱ��.ʱ��� & "(" & Format(obj�ϰ�ʱ��.��ʼʱ��, "hh:mm") & "-" & Format(obj�ϰ�ʱ��.����ʱ��, "hh:mm") & ")")
        objListItem.SubItems(1) = obj�ϰ�ʱ��.��ʼʱ��
        objListItem.SubItems(2) = obj�ϰ�ʱ��.����ʱ��
        objListItem.Tag = obj�ϰ�ʱ��.ʱ���
    Next
    If mobj���ﰲ��.���º�����λ Then picWorkTimeList.Enabled = False
    
    '��Դ��Ϣ
    SourceInfor.LoadData mobj���ﰲ��.�����Դ
    
    cldsCalenbarSel.LoadData mobj���ﰲ��
    
    '��ѯ��ʼ����
    If cldsCalenbarSel.ShowStyle = Show_Plan_Day And cboDays.Enabled Then
        If mobj���ﰲ��.�Ű෽ʽ = 1 Then
            dtCur = mobj���ﰲ��.��ʼʱ��
            cboDays.Clear
            Do While True
                cboDays.AddItem Format(dtCur, "yyyy/mm/dd")
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
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
    Dim objTemp As �����¼��
    
    On Error GoTo Errhand
    '��ǰ��Ŀ
    If mobj���ﰲ��.Count = 0 Then
        Set objTemp = New �����¼��
    Else
        Set objTemp = mobj���ﰲ��(1).Clone
    End If
    
    mstrCurDay = objTemp.��������
    Call SetTitleText
    
    'ʱ���
    CheckWorkTime objTemp
    '����
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

Private Function CheckWorkTime(ByVal obj�����¼�� As �����¼��) As Boolean
    'ѡ��ʱ���
    Dim objListItem As ListItem, i As Integer
    Dim objItem As �����¼
    
    On Error GoTo Errhand
    For i = 1 To lvwWorkTime.ListItems.Count
        lvwWorkTime.ListItems(i).Checked = False
    Next
    If obj�����¼�� Is Nothing Then Exit Function
    For Each objItem In obj�����¼��
        If objItem.ʱ��� <> "" Then
            For i = 1 To lvwWorkTime.ListItems.Count
                If objItem.ʱ��� = lvwWorkTime.ListItems(i).Tag Then
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
    '����:��ʼ��Docking�ؼ�
    '����:���˺�
    '����:2016-01-08 14:34:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    Dim strReg As String
    Dim panThis As Pane, panLeft As Pane
    
    On Error GoTo Errhand
    sngWidth = picDateList.Width / Screen.TwipsPerPixelX
    sngHeight = picDateList.Height / Screen.TwipsPerPixelY
    
    Set panLeft = dkpMain.CreatePane(Pancel_Index.Pan_����, sngWidth, sngHeight, DockTopOf, Nothing)
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Title = "": panLeft.Tag = Pancel_Index.Pan_����
    panLeft.handle = picDateList.hWnd
    
    panLeft.MinTrackSize.Height = sngHeight
    panLeft.MaxTrackSize.Height = sngHeight
    panLeft.MaxTrackSize.Width = sngWidth
    panLeft.MinTrackSize.Width = sngWidth * 2 / 3
    
    
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_����, sngWidth, 300, DockRightOf, panLeft)
    panThis.Title = ""
    panThis.Tag = Pancel_Index.Pan_����
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.handle = picDetailedList.hWnd
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_ʱ���, sngWidth, 300, DockBottomOf, panLeft)
    panThis.Title = "�ϰ�ʱ��"
    panThis.Tag = Pancel_Index.Pan_ʱ���
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.handle = picWorkTimeList.hWnd
    
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_��Դ, sngWidth, 300, DockBottomOf, panThis)
    panThis.Title = "��ǰ��Դ��Ϣ"
    panThis.Tag = Pancel_Index.Pan_��Դ
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.handle = picSouceList.hWnd
     
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Call picDateList_Resize
    'Set dkpMain.PaintManager.CaptionFont = use.Font
    
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub btnRight_Click()
    Dim intIndex As Integer
    
    On Error GoTo Errhand
    If Not mobj���ﰲ�� Is Nothing Then
        If IsDate(mstrCurDay) Then
            If DateDiff("d", mobj���ﰲ��.��ʼʱ��, DateAdd("d", 1, mstrCurDay)) >= 0 _
                And DateDiff("d", DateAdd("d", 1, mstrCurDay), mobj���ﰲ��.��ֹʱ��) >= 0 Then
                mstrCurDay = Format(DateAdd("d", 1, mstrCurDay), "yyyy-mm-dd")
            End If
        Else
            If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj���ﰲ��.�Ű���� <> 1 Then
                If mobj���ﰲ��.�Ű���� = 6 Then
                    If Val(mstrCurDay) + 1 <= 31 Then mstrCurDay = Val(mstrCurDay) + 1 & "��"
                End If
            Else '����
                intIndex = GetWeekIndex(mstrCurDay)
                If intIndex + 1 >= 0 And intIndex + 1 <= 6 Then
                    mstrCurDay = GetWeekName(intIndex + 1)
                End If
            End If
        End If
    End If
    Call SetButtonEnabled
    
    If mobj���ﰲ�� Is Nothing Then Exit Sub
    Call CurPlanChanged(mstrCurDay)
    cldsCalenbarSel.LoadData mobj���ﰲ��
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetButtonEnabled()
    '������һ������һ����ť����״̬
    Dim intIndex As Integer
    
    On Error GoTo Errhand
    If mobj���ﰲ�� Is Nothing Then
        btnLeft.Enabled = False
        btnRight.Enabled = False
    Else
        If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj���ﰲ��.�Ű���� <> 1 Then
            '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
            If mobj���ﰲ��.�Ű���� = 6 Then
                btnLeft.Enabled = True
                btnRight.Enabled = True
                If Val(mstrCurDay) <= 1 Or mobj���ﰲ��.�ѱ�����ﰲ��.Count = 0 Then
                    btnLeft.Enabled = False
                End If
                If Val(mstrCurDay) >= 31 Or mobj���ﰲ��.�ѱ�����ﰲ��.Count = 0 Then
                    btnRight.Enabled = False
                End If
            Else
                btnLeft.Enabled = False
                btnRight.Enabled = False
            End If
        Else
            btnLeft.Enabled = True
            btnRight.Enabled = True
            If IsDate(mstrCurDay) Then '����
                If DateDiff("d", mobj���ﰲ��.��ʼʱ��, mstrCurDay) <= 0 Then
                    btnLeft.Enabled = False
                End If
                If DateDiff("d", mobj���ﰲ��.��ֹʱ��, mstrCurDay) >= 0 Then
                    btnRight.Enabled = False
                End If
            Else '����
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
    Case Pancel_Index.Pan_����
        Item.handle = picDateList.hWnd
    Case Pancel_Index.Pan_ʱ���
        Item.handle = picWorkTimeList.hWnd
    Case Pancel_Index.Pan_��Դ
        Item.handle = picSouceList.hWnd
    Case Pancel_Index.Pan_����
        Item.handle = picDetailedList.hWnd
    End Select
End Sub

Private Sub lvwWorkTime_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, obj���ﰲ�� As ���ﰲ��
    Dim blnChecked As Boolean, objItem As ���ﰲ��, objTemp As ���ﰲ��
    Dim lngFindIndex As Long
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    
    On Error GoTo Errhand
    blnChecked = Item.Checked
    Item.Checked = Not blnChecked
    If mobj���ﰲ��.Count = 0 Then
        MsgBox IIf(cldsCalenbarSel.ShowStyle = Show_Plan_Rule, "�������δѡ��", "��������δѡ��"), vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If blnChecked Then
        '�����ϰ�ʱ��Σ������н���
        Dim dtCurStart As Date, dtCurEnd As Date
        Dim dtStart As Date, dtEnd As Date
        dtStart = Item.SubItems(1): dtEnd = Item.SubItems(2)
        If DateDiff("n", dtStart, dtEnd) <= 0 Then dtEnd = DateAdd("d", 1, dtEnd)
        For i = 1 To lvwWorkTime.ListItems.Count
            If lvwWorkTime.ListItems(i).Checked Then
                dtCurStart = lvwWorkTime.ListItems(i).SubItems(1): dtCurEnd = lvwWorkTime.ListItems(i).SubItems(2)
                If DateDiff("n", dtCurStart, dtCurEnd) <= 0 Then dtCurEnd = DateAdd("d", 1, dtCurEnd)
                
                If Not (DateDiff("n", dtCurStart, dtEnd) <= 0 Or DateDiff("n", dtCurEnd, dtStart) >= 0) Then
                    MsgBox "��ǰ�ϰ�ʱ�ε�ʱ�䷶Χ����ѡ���ϰ�ʱ�Ρ�" & Left(lvwWorkTime.ListItems(i).Text, InStr(lvwWorkTime.ListItems(i).Text, "(") - 1) & "����ʱ�䷶Χ���ص�������ͬʱѡ��", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        Next
    End If
    Item.Checked = blnChecked
    
    mobj���ﰲ��.RemoveAll
    Set obj�����¼�� = cpdClinicPlanDetailedPag.Get�����¼��
    mobj���ﰲ��.AddItem obj�����¼��, "K" & obj�����¼��.��������
    
    If Item.Checked Then
        Set obj�����¼�� = mobj���ﰲ��(1).Clone
        mobj���ﰲ��.RemoveAll

        With mobj���ﰲ��.�����Դ
            Set obj�����¼ = New �����¼
            obj�����¼.ʱ��� = Item.Tag
            Set obj�����¼.�ϰ�ʱ�� = mobj���ﰲ��.�����ϰ�ʱ��("K" & obj�����¼.ʱ���).Clone
            obj�����¼.�Ƿ��ʱ�� = .�Ƿ��ʱ��
            obj�����¼.�Ƿ���ſ��� = .�Ƿ���ſ���
            obj�����¼.ԤԼ���� = .ԤԼ����
            obj�����¼.���﷽ʽ = .���﷽ʽ
            obj�����¼.���º�����λ = mobj���ﰲ��.���º�����λ

            Set obj�����¼.�����������Ҽ�.���з������� = .�������Ҽ�.���з�������.Clone
            obj�����¼.�����������Ҽ�.���﷽ʽ = .���﷽ʽ
            Set obj�����¼.�����������Ҽ� = .�������Ҽ�.Clone

            Set obj�����¼.������Ϣ�� = New ������Ϣ��
            Set obj�����¼.������Ϣ��.�ϰ�ʱ�� = obj�����¼.�ϰ�ʱ��
            obj�����¼.������Ϣ��.ʱ��� = obj�����¼.ʱ���
            obj�����¼.������Ϣ��.�Ƿ��ʱ�� = .�Ƿ��ʱ��
            obj�����¼.������Ϣ��.�Ƿ���ſ��� = .�Ƿ���ſ���
            obj�����¼.������Ϣ��.ԤԼ���� = .ԤԼ����
            obj�����¼.������Ϣ��.����Ƶ�� = .����Ƶ��

            Set obj�����¼.������λ���Ƽ� = New ������λ���Ƽ�
            Set obj�����¼.������λ���Ƽ�.������Ϣ�� = obj�����¼.������Ϣ��.Clone
            Set obj�����¼.������λ���Ƽ�.���к�����λ = mobj���ﰲ��.���к�����λ.Clone

            obj�����¼��.AddItem obj�����¼
        End With
        mobj���ﰲ��.AddItem obj�����¼��, "K" & obj�����¼��.��������
    Else
        Set obj�����¼�� = mobj���ﰲ��(1)
        For i = 1 To obj�����¼��.Count
            If obj�����¼��(i).ʱ��� = Item.Tag Then
                obj�����¼��.Remove i: Exit For
            End If
        Next
    End If
    cpdClinicPlanDetailedPag.LoadData mobj���ﰲ��(1)
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
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    If Not mobj���ﰲ�� Is Nothing Then
        If IsDate(mstrCurDay) Then
            If DateDiff("d", mobj���ﰲ��.��ʼʱ��, DateAdd("d", -1, mstrCurDay)) >= 0 _
                And DateDiff("d", DateAdd("d", -1, mstrCurDay), mobj���ﰲ��.��ֹʱ��) >= 0 Then
                mstrCurDay = Format(DateAdd("d", -1, mstrCurDay), "yyyy-mm-dd")
            End If
        Else
            If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj���ﰲ��.�Ű���� <> 1 Then
                If mobj���ﰲ��.�Ű���� = 6 Then
                    If Val(mstrCurDay) - 1 > 0 Then mstrCurDay = Val(mstrCurDay) - 1 & "��"
                End If
            Else '����
                intIndex = GetWeekIndex(mstrCurDay)
                If intIndex - 1 >= 0 And intIndex - 1 <= 6 Then
                    mstrCurDay = GetWeekName(intIndex - 1)
                End If
            End If
        End If
    End If
    Call SetButtonEnabled
    
    If mobj���ﰲ�� Is Nothing Then Exit Sub
    Call CurPlanChanged(mstrCurDay)
    cldsCalenbarSel.LoadData mobj���ﰲ��
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optRule_Click(index As Integer)
    '����:����Ӧ���ڵ���ʾ
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler:
    fraLoopSkip.Visible = False
    picApplyWeek.Visible = False
    If index = 3 Then '������
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
    lblTittle.Caption = "�ް���"
    Select Case cldsCalenbarSel.ShowStyle
    Case Show_Plan_Rule
        picApply.Visible = m_Enabled
        picApplyRule.Visible = False
        picApplyWeek.Visible = False
        If mobj���ﰲ��.�Ű���� = 1 Then
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
        If Not mobj���ﰲ�� Is Nothing Then
            If mobj���ﰲ��.�Ű෽ʽ = 1 Then '����
                If optRule(0).Value Then optRule(1).Value = True
                optRule(0).Value = True
            ElseIf mobj���ﰲ��.�Ű෽ʽ = 2 Then '����
                picApplyRule.Visible = False
                picApplyWeek.Visible = True
                picApplyWeek.Top = picApplyRule.Top + 60
                picApply.Height = picApplyWeek.Top + picApplyWeek.Height
            End If
        End If
    End Select
    Call SetButtonEnabled
    
    If mstrCurDay <> "" Then
        '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
        Select Case mobj���ﰲ��.�Ű����
        Case 2
            strTemp = "�����ճ���"
        Case 3
            strTemp = "��˫�ճ���"
        Case 4, 5
            strTemp = "��" & Val(mstrCurDay) & "����ѭ"
        Case 6
            strTemp = ""
            If mobj���ﰲ��.�ѱ�����ﰲ��.�Ű���� = 6 Then
                For i = 1 To mobj���ﰲ��.�ѱ�����ﰲ��.Count
                    strTemp = strTemp & "," & mobj���ﰲ��.�ѱ�����ﰲ��(i).��������
                Next
            End If
            If mobj���ﰲ��.���º�����λ = False Then
                strTemp = strTemp & "," & mstrCurDay
                If mobj���ﰲ��.Ӧ���� <> "" Then
                    strTemp = strTemp & "," & mobj���ﰲ��.Ӧ����
                End If
            End If
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            strTemp = "��ÿ�µ�" & ZlNumStrSort(strTemp, True) & "�չ̶�����"
        Case Else
            strTemp = mstrCurDay
        End Select
    End If
        
    If strTemp = "" Then strTemp = "�ް���"
    lblTittle.Caption = IIf(IsDate(strTemp), Format(strTemp, "yyyy-mm-dd"), strTemp)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function IsValied() As Boolean
    '�������
    Dim blnSelected As Boolean
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    For i = 1 To lvwWorkTime.ListItems.Count
        If lvwWorkTime.ListItems(i).Checked Then
            blnSelected = True: Exit For
        End If
    Next
    If blnSelected = False Then
        MsgBox "δ�����ϰ�ʱ�Σ�", vbInformation + vbOKOnly, gstrSysName
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

Public Property Get Get���ﰲ��() As ���ﰲ��
    Dim obj���ﰲ�� As New ���ﰲ��
    Dim obj�����¼�� As �����¼��
    
    On Error GoTo Errhand
    Set obj���ﰲ�� = mobj���ﰲ��.Clone
    obj���ﰲ��.RemoveAll
    Set obj�����¼�� = cpdClinicPlanDetailedPag.Get�����¼��
    obj���ﰲ��.AddItem obj�����¼��, "K" & obj�����¼��.��������
    If obj���ﰲ��.�Ű���� <> 6 Then
        'ģ����ض����ڵ�Ӧ������ѡ��ʱ��ȷ��
        obj���ﰲ��.Ӧ���� = GetApplyToStr()
    End If
    Set Get���ﰲ�� = obj���ﰲ��
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Private Function GetApplyToStr() As String
    '��ȡӦ�����ַ���
    Dim strApplyTo As String, i As Integer
    Dim dtCur As Date, str�������� As String
    Dim varTemp As Variant
    
    On Error GoTo Errhand
    If mobj���ﰲ�� Is Nothing Then Exit Function
    If m_Enabled = False Then Exit Function
    If picApply.Visible = False Then Exit Function
    If mobj���ﰲ��.Count > 0 Then
        str�������� = mobj���ﰲ��(1).��������
    End If
    
    If picApplyRule.Visible Then
        If optRule(1).Value Then '����
            If mobj���ﰲ��.�Ű෽ʽ = 1 Then
                dtCur = mobj���ﰲ��.��ʼʱ��
                Do While True
                    If Day(dtCur) Mod 2 = 1 And DateDiff("n", str��������, dtCur) <> 0 Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                    End If
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
                Loop
            End If
        ElseIf optRule(2).Value Then '˫��
            If mobj���ﰲ��.�Ű෽ʽ = 1 Then
                dtCur = mobj���ﰲ��.��ʼʱ��
                Do While True
                    If Day(dtCur) Mod 2 = 0 And DateDiff("n", str��������, dtCur) <> 0 Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                    End If
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
                Loop
            End If
        ElseIf optRule(3).Value Then '����
            If picApplyWeek.Visible Then
                dtCur = mobj���ﰲ��.��ʼʱ��
                Do While True
                    For i = chkWeek.LBound To chkWeek.UBound
                        If chkWeek(i).Value = vbChecked Then
                            If Weekday(dtCur, vbMonday) = i + 1 And DateDiff("n", str��������, dtCur) <> 0 Then
                                strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                            End If
                        End If
                    Next
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
                Loop
            End If
        ElseIf optRule(4).Value Then '��ѭ
            If mobj���ﰲ��.�Ű෽ʽ = 1 Then
                If Not (cboDays.ListIndex = -1 Or Val(txtSkip.Text) = 0) Then
                    dtCur = CDate(cboDays) '��ʼʱ��
                    Do While True
                        If DateDiff("n", str��������, dtCur) <> 0 Then
                            strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                        End If
                        dtCur = DateAdd("d", Val(txtSkip.Text), dtCur)
                        If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
                    Loop
                End If
            End If
        End If
        If strApplyTo <> "" Then strApplyTo = Mid(strApplyTo, 2)
        GetApplyToStr = strApplyTo
        Exit Function
    End If
    If picApplyWeek.Visible Then
        If cldsCalenbarSel.ShowStyle = Show_Plan_Day Then '�����Ű�
            dtCur = mobj���ﰲ��.��ʼʱ��
            Do While True
                For i = chkWeek.LBound To chkWeek.UBound
                    If chkWeek(i).Value = vbChecked Then
                        If Weekday(dtCur, vbMonday) = i + 1 And DateDiff("n", str��������, dtCur) <> 0 Then
                            strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy/mm/dd")
                        End If
                    End If
                Next
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
            Loop
        Else
            For i = chkWeek.LBound To chkWeek.UBound
                If chkWeek(i).Value = vbChecked Then
                    If GetWeekName(i) <> str�������� Then
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    Set mobj���ﰲ�� = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

Private Sub CurPlanChanged(ByVal str�������� As String)
    '��ǰѡ����Ŀ�ı�
    Dim objItem As New �����¼��
    
    On Error GoTo Errhand
    mobj���ﰲ��.RemoveAll
    If mobj���ﰲ��.�ѱ�����ﰲ��.Exits("K" & str��������) Then
        mobj���ﰲ��.AddItem mobj���ﰲ��.�ѱ�����ﰲ��("K" & str��������)
    Else
        objItem.�������� = str��������
        mobj���ﰲ��.AddItem objItem
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
