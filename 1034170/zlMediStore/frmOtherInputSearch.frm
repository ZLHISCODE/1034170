VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmOtherInputSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4200
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmOtherInputSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2055
      Left            =   1680
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   3975
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmOtherInputSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmOtherInputSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   2715
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   5505
         Begin VB.CheckBox Chk类别 
            Caption         =   "类别"
            Height          =   300
            Left            =   480
            TabIndex        =   15
            Top             =   1440
            Width           =   960
         End
         Begin VB.CommandButton Cmd药品 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   11
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox Txt药品 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   10
            Top             =   360
            Width           =   3255
         End
         Begin VB.CheckBox Chk药品 
            Caption         =   "药品"
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Width           =   990
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1620
            MaxLength       =   8
            TabIndex        =   17
            Top             =   2100
            Width           =   1365
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   18
            Top             =   2100
            Width           =   1365
         End
         Begin VB.ComboBox Cbo类别 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1440
            Width           =   3495
         End
         Begin VB.CheckBox Chk生产商 
            Caption         =   "生产商"
            Height          =   300
            Left            =   480
            TabIndex        =   12
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox txt生产商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   13
            Top             =   900
            Width           =   3255
         End
         Begin VB.CommandButton Cmd生产商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   14
            Top             =   900
            Width           =   255
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   750
            TabIndex        =   32
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   3120
            TabIndex        =   31
            Top             =   2160
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   5520
         Begin VB.TextBox txt开始No 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt结束NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk填制 
            Caption         =   "未审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "包含冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   8
            Top             =   2280
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   160628739
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   4
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   160628739
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   160628739
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   7
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   160628739
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   29
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   28
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核日期"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   27
            Top             =   1905
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   3
            Left            =   3345
            TabIndex        =   26
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制日期"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   25
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   24
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   20
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   19
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmOtherInputSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '查找字符串
Private BlnAdvance As Boolean '是否展开
Private mdatStart As Date   '开始时间
Private mdatEnd As Date     '结束时间
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '父窗体
Public lng药品ID As Long
Private mstrSelectTag As String     '当前选择的对象

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    str填制人 As String
    str审核人 As String
    str产地 As String
    lng入出类别 As Long
End Type

Private SQLCondition As Type_SQLCondition
Public Function GetSearch(ByVal FrmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strNO开始 As String, _
        ByRef strNO结束 As String, _
        ByRef date填制时间开始 As Date, _
        ByRef date填制时间结束 As Date, _
        ByRef date审核时间开始 As Date, _
        ByRef date审核时间结束 As Date, _
        ByRef lng药品 As Long, _
        ByRef str填制人 As String, _
        ByRef str审核人 As String, _
        ByRef str产地 As String, _
        ByRef lng入出类别 As Long) As String
    mstrFind = ""
    Set mfrmMain = FrmMain
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    
    strNO开始 = SQLCondition.strNO开始
    strNO结束 = SQLCondition.strNO结束
    date填制时间开始 = SQLCondition.date填制时间开始
    date填制时间结束 = SQLCondition.date填制时间结束
    date审核时间开始 = SQLCondition.date审核时间开始
    date审核时间结束 = SQLCondition.date审核时间结束
    lng药品 = SQLCondition.lng药品
    str审核人 = SQLCondition.str审核人
    str填制人 = SQLCondition.str填制人
    str产地 = SQLCondition.str产地
    lng入出类别 = SQLCondition.lng入出类别
    
End Function

Private Sub Cbo类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        cmd确定.SetFocus
    End If
    
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        cmd确定.SetFocus
    End If
End Sub



Private Sub Chk类别_Click()
    Cbo类别.Enabled = IIf(Chk类别.Value = 1, True, False)
End Sub

Private Sub Chk类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Chk类别.Value = 1 Then
        Cbo类别.SetFocus
    Else
        Txt填制人.SetFocus
    End If
End Sub

Private Sub chk审核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If chk审核.Value = 0 Then
            cmd确定.SetFocus
        Else
            SendKeys vbTab
        End If
    End If
    
End Sub

Private Sub Chk生产商_Click()
    Me.txt生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
    Cmd生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
End Sub

Private Sub Chk生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        If Chk生产商.Value = 1 Then
            txt生产商.SetFocus
        ElseIf Chk类别.Visible = True Then
            Chk类别.SetFocus
        Else
            Txt填制人.SetFocus
        End If
    End If
End Sub

Private Sub chk填制_Click()
    dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
    
End Sub

Private Sub chk填制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk药品_Click()
    Txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    Cmd药品.Enabled = IIf(Chk药品.Value = 1, True, False)
End Sub

Private Sub Chk药品_GotFocus()
    sstFilter.Tab = 1
    Chk药品.SetFocus
End Sub

Private Sub Chk药品_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk药品.Value = 1 Then
        Txt药品.SetFocus
    Else
        Chk生产商.SetFocus
    End If
End Sub




Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = 24
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    '检查数据
    If Chk药品.Value = 1 Then
        If Txt药品.Tag = 0 Then
            MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
            Me.Txt药品.SetFocus
            Exit Sub
        End If
    End If
    
    If Chk生产商.Value = 1 Then
        If txt生产商.Tag = 0 Then
            MsgBox "请选择需查询的药品生产商信息！", vbInformation, gstrSysName
            Me.txt生产商.SetFocus
            Exit Sub
        End If
    End If
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "对不起，必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If
    
    mstrFind = ""
    '基本查询条件
    Dim i As Integer
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                    & " or (A.审核日期 Between [5] And [6]))"
        Else
            mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                    & " or (A.审核日期 Between [5] And [6] and a.记录状态 =1))  "
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.审核日期 Between [5] And [6] "
        Else
            mstrFind = " And A.审核日期 Between [5] And [6] and a.记录状态 =1 "
            
        End If
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.填制日期 Between [3] And [4]) and 审核日期 is null "
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        txt开始No.Text = GetFullNO(txt开始No.Text, intNO, lng库房ID)
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        txt结束NO.Text = GetFullNO(txt结束NO.Text, intNO, lng库房ID)
    End If
    
    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [1] And A.No <=[2] "
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [1] "
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= [2] "
    
    SQLCondition.strNO开始 = Me.txt开始No
    SQLCondition.strNO结束 = Me.txt结束NO
    SQLCondition.date填制时间开始 = CDate(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date审核时间开始 = CDate(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date审核时间结束 = CDate(Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59")
    
    '扩展查询条件
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If Chk药品.Value = 1 Then
        lng药品ID = Txt药品.Tag
        mstrFind = mstrFind & " And A.药品ID + 0=[7] "
    End If
    
    If Chk生产商.Value = 1 Then mstrFind = mstrFind & " And A.产地=[12] "
    If Chk类别.Value = 1 Then mstrFind = mstrFind & " And A.入出类别ID=[15] "
    If Me.Txt审核人 <> "" Then mstrFind = mstrFind & " And A.审核人 like [10] "
    If Me.Txt填制人 <> "" Then mstrFind = mstrFind & " And A.填制人 like [11] "
    
    SQLCondition.lng药品 = Val(Txt药品.Tag)
    SQLCondition.str产地 = txt生产商
    SQLCondition.str审核人 = Me.Txt审核人 & "%"
    SQLCondition.str填制人 = Me.Txt填制人 & "%"
    SQLCondition.lng入出类别 = Cbo类别.ItemData(Cbo类别.ListIndex)
    
    Unload Me
End Sub

Private Sub Cmd生产商_Click()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
    Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-药品生产商", gstrNodeNo)
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "所有药品生产商"
        .lngMode = 1
        .Show 1, Me
        If .BlnSuccess = True Then
            txt生产商.Tag = 1
            txt生产商.Text = .CurrentName
        End If
    End With
    Unload FrmSelect
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品其他入库管理", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    End If
    
'    Set RecReturn = Frm药品选择器.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , , False)

    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品ID
    
    Chk生产商.SetFocus
End Sub

Private Sub dtp结束时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp开始时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.dtp结束时间(Index).SetFocus
End Sub

Private Sub Form_Load()
    Me.dtp结束时间(0) = zldatabase.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    Me.Txt药品.Tag = 0
    Me.txt生产商.Tag = 0
    lng药品ID = 0
    
    sstFilter.Tab = 0
    BlnAdvance = False
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    
    gstrSQL = "Select 编码,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null"
    Set rsCompete = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "药品生产商", gstrNodeNo)
    With rsCompete
        If .EOF Then
            MsgBox "药品生产商信息不全,请在字典管理中设置药品生产商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    gstrSQL = "SELECT B.Id,b.名称 " _
            & "FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID AND A.单据 = 4 "
    Set rsCompete = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsCompete
        If .EOF Then
            MsgBox "药品其他入库没有设置相应的入出类别，请检查药品入出分类！", vbInformation, gstrSysName
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Cbo类别.AddItem .Fields(1)
            Cbo类别.ItemData(Cbo类别.NewIndex) = .Fields(0)
            .MoveNext
        Loop
        Cbo类别.ListIndex = 0
        .Close
    End With
    CheckCompete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Maker"
                txt生产商.SetFocus
                txt生产商.SelStart = 0
                txt生产商.SelLength = Len(txt生产商.Text)
            
            Case "Booker"
                Txt填制人.SetFocus
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
            Case "Verify"
                Txt审核人.SetFocus
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
        End Select
        Cancel = True
    End If
    Call ReleaseSelectorRS
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Maker"
                    txt生产商.Text = .TextMatrix(.Row, 1)
                    txt生产商.Tag = 1
                    Chk类别.SetFocus
                Case "Booker"
                    Txt填制人 = .TextMatrix(.Row, 2)
                    Txt审核人.SetFocus
                Case "Verify"
                    Txt审核人 = .TextMatrix(.Row, 2)
                    cmd确定.SetFocus
            End Select
            .Visible = False
            
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
        End If
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk药品.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk药品.SetFocus
        End If
    End If
    
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = 24
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = GetFullNO(txt结束NO.Text, intNO, lng库房ID)
        End If
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = 24
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = GetFullNO(txt开始No.Text, intNO, lng库房ID)
        End If
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then cmd确定.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            cmd确定.SetFocus
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取审核人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", _
                        Me.Txt审核人 & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt审核人.Top + Txt审核人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt审核人.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt审核人.Top - Txt审核人.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt审核人 = IIf(IsNull(!姓名), "", !姓名)
                cmd确定.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.txt生产商 = "" Then Exit Sub
        If Trim(txt生产商) = "" Then Exit Sub
        txt生产商 = UCase(txt生产商)
    
        Dim rsTemp As New ADODB.Recordset
        
        On Error GoTo errHandle
        gstrSQL = "Select 编码,简码,名称 From 药品生产商 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2]) " & _
                  "Order By 编码"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[药品生产商]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.txt生产商 & "%", _
                        Me.txt生产商 & "%", gstrNodeNo)
       
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Maker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + txt生产商.Top + txt生产商.Height
                    .Left = sstFilter.Left + fra附加条件.Left + txt生产商.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt生产商 = IIf(IsNull(!名称), "", !名称)
                txt生产商.Tag = 1
            End If
        End With
        
        If Chk类别.Visible = True Then
            If Chk类别.Value = 1 Then
                Cbo类别.SetFocus
            Else
                Chk类别.SetFocus
            End If
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            Txt审核人.SetFocus
            Exit Sub
        End If
        Txt填制人.Text = UCase(Txt填制人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取填制人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", _
                        Me.Txt填制人 & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt填制人.Top + Txt填制人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt填制人.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt填制人.Top - Txt填制人.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt填制人 = IIf(IsNull(!姓名), "", !姓名)
                Me.Txt审核人.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra附加条件.Left + Txt药品.Left
    sngTop = Me.Top + sstFilter.Top + fra附加条件.Top + Txt药品.Top + Txt药品.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt药品.Height - 3630
    End If
    
    strkey = Trim(Txt药品.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品其他入库管理", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    End If
    
'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , , False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品ID
    
    Chk生产商.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

