VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppRemove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "应用系统拆卸"
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmAppRemove.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6600
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   1605
      TabIndex        =   17
      Top             =   3510
      Width           =   1100
   End
   Begin VB.Frame fraSys 
      Height          =   1365
      Left            =   2085
      TabIndex        =   9
      Top             =   1125
      Width           =   3945
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "所有者"
         Height          =   180
         Left            =   210
         TabIndex        =   16
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lblOwner 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   795
         TabIndex        =   15
         Top             =   930
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "系统名"
         Height          =   180
         Left            =   210
         TabIndex        =   13
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblSysName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   795
         TabIndex        =   12
         Top             =   225
         Width           =   2895
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   795
         TabIndex        =   11
         Top             =   570
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "版本号"
         Height          =   180
         Left            =   210
         TabIndex        =   10
         Top             =   630
         Width           =   540
      End
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4005
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   0
      Width           =   1335
      Begin VB.Image imgRemove 
         Height          =   2550
         Left            =   120
         Picture         =   "frmAppRemove.frx":058A
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdGetIni 
      Caption         =   "选择(&S)…"
      Height          =   350
      Left            =   4935
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4095
      TabIndex        =   3
      Top             =   3510
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   60
      TabIndex        =   5
      Top             =   3405
      Width           =   7140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5185
      TabIndex        =   2
      Top             =   3510
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   4005
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppRemove.frx":5B70
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8070
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "11:02"
            Key             =   "STANUM"
         EndProperty
      EndProperty
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
   Begin VB.Label lblNote 
      Caption         =   "    如果完成了上述工作，可以在正确指定应用安装配置文件后执行拆卸，将自动清除该系统的所有数据、独立的所有者和独立的存储空间。"
      Height          =   525
      Index           =   1
      Left            =   1605
      TabIndex        =   8
      Top             =   540
      Width           =   4680
   End
   Begin VB.Label lbliniFile 
      AutoSize        =   -1  'True
      Caption         =   "应用安装配置文件"
      Height          =   180
      Left            =   2085
      TabIndex        =   6
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2085
      TabIndex        =   1
      Top             =   3000
      Width           =   3945
   End
   Begin VB.Label lblNote 
      Caption         =   "    拆卸操作是对指定系统的彻底清除，建议在拆卸前保留多个完整可靠的数据备份；"
      Height          =   375
      Index           =   0
      Left            =   1605
      TabIndex        =   4
      Top             =   105
      Width           =   4680
   End
End
Attribute VB_Name = "frmAppRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intDefSysCode As Integer                '系统编号
Dim strDefSysName As String                 '系统名称
Dim strDefVersion As String                 '版本号
Dim strDefSpace   As String                 '表空间

Dim mbln帐套 As Boolean    '本次安装是否是属于帐套安装
Dim mlng帐套 As Long

Dim objFile As New FileSystemObject
Dim objText As TextStream

Dim rsTemp As New ADODB.Recordset
Dim strSQL As String, strTemp As String
Dim intCount As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetIni_Click()
    With frmMDIMain.dlgMain
        .Filename = lblFileName.Caption
        .DialogTitle = "选择应用安装配置文件"
        .Filter = "(应用安装配置文件)|zlSetup.ini"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        Else
            lblFileName.Caption = .Filename
        End If
    End With
    
    If CheckIniFile(lblFileName.Caption, True) = False Then
        cmdOk.Enabled = False
        lblFileName.Caption = ""
        cmdGetIni.SetFocus
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.Name
End Sub

Private Sub cmdOK_Click()
    Me.MousePointer = 11
    If DeleteSystem = False Then
        Me.MousePointer = 0
        Exit Sub
    End If
    Me.MousePointer = 0
    Unload Me
End Sub

Private Function DeleteSystem() As Boolean
    Dim msgSystem As VbMsgBoxResult
    Dim strMsg As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSpaces As String
    Dim aryTbs() As String, aryChr() As String
    Dim blnEnjoy As Boolean
    
    gstrSQL = "Select 1 From zltools.zlbakspaces where 当前<>1 and 系统=" & Val(lblSysName.Tag)
    rsTemp.Open gstrSQL, gcnOracle
    If Not rsTemp.EOF Then
        strMsg = "被拆卸系统存在历史数据空间,是否继续拆卸？" & vbCrLf & _
             "选择【是】：将保留历史业务数据(即自动剥离)，必要时可通过“再植”恢复。" & vbCrLf & _
             "选择【否】：将退出拆卸程序，你可以在管理工具->数据管理->数据转移中删" & vbCrLf & Space(12) & "除历史数据空间后，再进行拆卸。"
        msgSystem = MsgBox(strMsg, vbQuestion Or vbYesNo Or vbDefaultButton3, gstrSysName)
        If msgSystem = vbNo Then Exit Function
    End If
    
    If MsgBox("系统拆卸操作，将删除该系统所有数据(包括表和表空间)，无法恢复。建议在做此操作前对数据库进行一次备份。" & vbCrLf & "确定要继续吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
        
    cmdGetIni.Enabled = False
    cmdCancel.Enabled = False
    cmdOk.Enabled = False
    
    On Error GoTo 0
    
    DoEvents
    '判断是否共享其他系统
    With rsTemp
        strSQL = "select 1 from zlSystems where upper(所有者)='" & UCase(lblOwner.Caption) & "'"
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        blnEnjoy = (.RecordCount > 1)
    End With
    '非共享系统
    If blnEnjoy = False Then
        With rsTemp
            If .State = adStateOpen Then .Close
            .Open "select 1 from gv$session where USERNAME='" & UCase(lblOwner.Caption) & "'", gcnOracle
            If .EOF = False Then
                MsgBox "系统所有者正连接到数据库上，无法完成卸载操作。请用SYS用户卸载！", vbExclamation, gstrSysName
                cmdGetIni.Enabled = True
                cmdCancel.Enabled = True
                cmdOk.Enabled = True
                Exit Function
            End If
        End With
    Else
        If UCase(lblOwner.Caption) <> UCase(gstrUserName) And gstrUserName <> "SYS" Then
            MsgBox "当前用户不是系统所有者或SYS用户，无法完成卸载操作！", vbExclamation, gstrSysName
            cmdGetIni.Enabled = True
            cmdCancel.Enabled = True
            cmdOk.Enabled = True
            Exit Function
        End If
    End If
    '搜索表空间及数据文件
    aryTbs = Split(strDefSpace, "||")
    For intCount = 0 To UBound(aryTbs)
        aryChr = Split(aryTbs(intCount), "|")
        strSpaces = IIf(strSpaces = "", Trim(aryChr(1)), strSpaces & "," & Trim(aryChr(1)))
    Next
    
    Call UnInstall(Me, blnEnjoy, stbThis, lblOwner.Caption, strSpaces, mlng帐套, lblSysName.Tag)
    MsgBox strDefSysName & "拆卸完成！", vbInformation, gstrSysName
    DeleteSystem = True
End Function

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    With imgRemove
        .Top = PicSetup.ScaleTop
        .Left = PicSetup.ScaleLeft
        .Height = PicSetup.ScaleHeight
        .Width = PicSetup.ScaleWidth
    End With
    With frmAppStart.lvwSys.SelectedItem
        lblSysName.Tag = Mid(.Key, 2)
        lblSysName.Caption = .Text
        lblVersion.Caption = .SubItems(1)
        lblOwner.Caption = .SubItems(3)
    End With
    
    Call Judge帐套
    
    If mbln帐套 = False Then
        '完全删除
        With rsTemp
            strSQL = "select 文件名 from zlSysFiles where 系统=" & lblSysName.Tag & " and 操作=1"
            If .State = adStateOpen Then .Close
            .Open strSQL, gcnOracle, adOpenKeyset
            If Not .EOF And Not .BOF Then
                lblFileName.Caption = .Fields(0).value
            End If
        End With
        If Not gobjFile.FileExists(lblFileName.Caption) Then
            If gobjFile.FileExists(App.Path & "\zlSetup.ini") Then
                lblFileName.Caption = App.Path & "\zlSetup.ini"
            End If
        End If
        
        If Trim(lblFileName.Caption) <> "" Then
            If CheckIniFile(lblFileName.Caption) = False Then
                lblFileName.Caption = ""
            Else
                cmdOk.Enabled = True
            End If
        End If
    Else
        '帐套删除
        cmdOk.Enabled = True
        cmdGetIni.Enabled = False
        lbliniFile.Enabled = False
        lblFileName.Enabled = False
    End If
End Sub

Private Sub Judge帐套()
    '判断是否应该把本次安装作为帐套安装
    Dim lng系统号 As Long, lngTemp As Long
    Dim lstTemp As ListItem

    
    mbln帐套 = False
    lng系统号 = lblSysName.Tag \ 100
    For Each lstTemp In frmAppStart.lvwSys.ListItems
        lngTemp = Mid(lstTemp.Key, 2)
        If lngTemp \ 100 = lng系统号 Then
            '系统相同
            
            If lngTemp <> lblSysName.Tag Then
                '有另一个帐套存在
                mbln帐套 = True
                mlng帐套 = lngTemp - lblSysName.Tag
                Exit For
            End If
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If cmdCancel.Enabled = False Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Function CheckIniFile(strFile As String, Optional blnMsg As Boolean) As Boolean
    err = 0
    On Error Resume Next
        
    '配置文件正确性检查
    Set objText = objFile.OpenTextFile(strFile)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统号]" Then
        intDefSysCode = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统名]" Then
        strDefSysName = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[版本号]" Then
        strDefVersion = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[表空间]" Then
        strDefSpace = UCase(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    objText.Close
    
    If err <> 0 Then
        CheckIniFile = False
        If blnMsg Then MsgBox "安装配置文件不正确", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '配置文件符合性检查
    If intDefSysCode <> Int(lblSysName.Tag / 100) Then
        err.Raise 10
        If blnMsg Then MsgBox "选择文件不是 " & lblSysName.Caption & " 的安装配置文件", vbExclamation, gstrSysName
    ElseIf Trim(strDefVersion) <> lblVersion.Caption Then
        err.Raise 10
        If blnMsg Then MsgBox "选择文件与 " & lblSysName.Caption & " 版本不符 ", vbExclamation, gstrSysName
    End If
    If err = 0 Then
        CheckIniFile = True
    Else
        CheckIniFile = False
    End If
End Function
