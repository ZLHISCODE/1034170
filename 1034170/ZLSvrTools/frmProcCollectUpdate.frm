VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcCollectUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "搜集更新"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmProcCollectUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   0
      Left            =   135
      ScaleHeight     =   4500
      ScaleWidth      =   10185
      TabIndex        =   6
      Top             =   1485
      Width           =   10185
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1140
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   105
         Width           =   1935
         _cx             =   3413
         _cy             =   2011
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "frmProcCollectUpdate.frx":6852
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   5
      Top             =   75
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "开始搜集(&S)"
      Height          =   350
      Left            =   7800
      TabIndex        =   4
      Top             =   6075
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   9240
      TabIndex        =   3
      Top             =   6090
      Width           =   1100
   End
   Begin VB.OptionButton opt 
      Caption         =   "当前数据库"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   1125
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.OptionButton opt 
      Caption         =   "其他数据库"
      Height          =   255
      Index           =   1
      Left            =   1650
      TabIndex        =   1
      Top             =   1125
      Width           =   1380
   End
   Begin VB.CommandButton cmdConnet 
      Caption         =   "连接配置(&L)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2970
      TabIndex        =   0
      Top             =   1065
      Width           =   1290
   End
   Begin MSComctlLib.ProgressBar pbr 
      Height          =   105
      Left            =   120
      TabIndex        =   8
      Top             =   6525
      Visible         =   0   'False
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3720
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "test"
      Height          =   180
      Left            =   135
      TabIndex        =   11
      Top             =   6300
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "搜集登记过程/函数"
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
      Left            =   1110
      TabIndex        =   10
      Top             =   105
      Width           =   2820
   End
   Begin VB.Label Label2 
      Caption         =   "请在下方选择当前版本的脚本配置文件，以便和当前版本数据库中的过程进行比较，得出有更改的过程。"
      Height          =   210
      Left            =   1140
      TabIndex        =   9
      Top             =   555
      Width           =   9180
   End
End
Attribute VB_Name = "frmProcCollectUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjMain As Object
Private mclsVsf As clsVsf
Private mblnOK As Boolean
Private WithEvents mfrmPageConfigure As frmProcConfigure
Attribute mfrmPageConfigure.VB_VarHelpID = -1

Private mcnOracle As ADODB.Connection

Public Function ShowMe(ByVal objMain As Object) As Boolean
    On Error GoTo errHand
    
    mblnOK = False
    
    Set mobjMain = objMain
    Me.Show 1, mobjMain
    
    ShowMe = mblnOK
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim intRow As Integer
    Dim intFlag As Integer
    Dim strSQL As String
    Dim objItem As Object
    Dim strUpPath As String
    Dim strFlag As String
    
    On Error GoTo errHand
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True)
            Call .ClearColumn
            Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[序号]", False, False, False)
            Call .AppendColumn("版本号", 0, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("系统名称", 1800, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("安装脚本", 2700, flexAlignLeftCenter, flexDTString, , "", True)
            
            Call .InitializeEdit(True, False, True)
            Call .InitializeEditColumn(.ColIndex("安装脚本"), True, vbVsfEditCommand)

            .IndicatorMode = 2
            .IndicatorCol = .ColIndex("序号")
            .ConstCol = .ColIndex("序号")
                
            .AppendRows = True
        End With
'        lblState.ForeColor = &HFF&
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        With vsf(0)
            strSQL = "Select A.编号,A.版本号,A.名称 as 系统名称,B.文件名 From zlSystems A,zlSysFiles B Where A.编号 = B.系统 And B.操作=1"
            Set rs = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "")
            If rs.BOF = False Then
                For intRow = 0 To rs.RecordCount - 1
                    intFlag = intFlag + 1
                    If .Rows < intFlag + 1 Then .Rows = intFlag + 1
                    .TextMatrix(intRow + 1, .ColIndex("系统名称")) = rs("系统名称").value
                    .TextMatrix(intRow + 1, .ColIndex("安装脚本")) = rs("文件名").value
                    
                    strFlag = rs("版本号").value
                    .TextMatrix(intRow + 1, .ColIndex("版本号")) = strFlag
                    strFlag = Split(strFlag, ".")(0) & "." & Split(strFlag, ".")(1) & ".0"
                    
                    '缺省升级脚本
                    strUpPath = Split(rs("文件名").value, "应用脚本")(0) & "升级脚本\" & strFlag & "\zlUpgrade.ini"
                                                            
                    .RowData(intRow + 1) = rs("编号").value
                    rs.MoveNext
                Next
            End If
        End With
        mclsVsf.UpdateSerial
        mclsVsf.AppendRows = True
    End Select
    ExecuteCommand = True
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnet_Click()
    If mfrmPageConfigure Is Nothing Then
        Set mfrmPageConfigure = New frmProcConfigure
    End If
    Call mfrmPageConfigure.ShowConfigure(Me)
End Sub

Private Sub cmdOK_Click()
    '1.建立两个临时文件夹
    Dim strTmp1 As String
    Dim strProcedure As String
    Dim strTmpReports As String
    Dim strFlag As String
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim lngLoop As Long
    Dim i As Integer
    Dim lngLine As Long
    Dim objFileLines As Long
    Dim strLine As String
    Dim strFMT As String
    Dim blnBlock As Boolean
    Dim blnSQL As Boolean
    Dim strPro As String
    Dim strTemp As String
    Dim strFileProName As String
    Dim objFileTemp As TextStream
    Dim objFolder As Folder
    Dim objFolderTemp As Folder
    Dim objCurFolder As Folder
    Dim objFile As File
    Dim objFileFlag As File
    Dim rsInit As ADODB.Recordset
    Dim intSysNumLast As Integer
    Dim lngTemp As Long
    Dim strCommand As String
    Dim lngProcess As Long
    Dim rsSQL As ADODB.Recordset
    Dim blnNew As Boolean
    Dim strOwner As String
    Dim strIniPath As String
    Dim strIni1 As String
    Dim strIniSys As String
    Dim strIniApp As String
    Dim lngSys As Long
    
    Dim objPercent As New clsPercent
    
    On Error GoTo errHand
    
    cmdOK.Enabled = False
    
    Call gclsBase.SQLRecord(rsSQL)
    
    lblTitle.Caption = "正在清除临时目录.."
    lblTitle.Visible = True
    DoEvents
    
    strTmp1 = App.Path & "\Tmp1"
    strProcedure = App.Path & "\Procedure"
    strTmpReports = App.Path & "\Reports"
    
    If mcnOracle Is Nothing Then
        MsgBox "请先进行连接配置，已确认搜集来源！", vbInformation + vbOKOnly, "中联软件"
        Exit Sub
    End If
       
    '------------------------------------------------------------------------------------------------------------------
    With vsf(0)
        
'        strSQL = "Delete From zlproceduretext where 性质 in (1,2))"
'        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
'        strSQL = "Delete from zlproceduretext where 过程id in (select id from zlprocedure where 类型 in (1,3))"
'        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
'        strSQL = "Delete from zlprocedure where 类型 in (1,3)"
'        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        
                
        strSQL = "Select 编号,名称,版本号 From zlSystems a"
        Set rsData = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "")
        If rsData.BOF = True Then
            MsgBox "当前数据库没有安装任何系统。", vbInformation + vbOKOnly, "中联软件"
            GoTo errEnd
        End If
        For i = 1 To .Rows - 1
            rsData.Filter = ""
            rsData.Filter = "编号=" & .RowData(i)
            If .TextMatrix(i, vsf(0).ColIndex("安装脚本")) = "" Then
                MsgBox "请选择" & .TextMatrix(i, .ColIndex("系统名称")) & "安装脚本"
                GoTo errEnd
            End If
            Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")))
            rsInit.Filter = "项目='版本号'"
            strIniApp = rsInit("内容").value '安装脚本版本号

            rsData.Filter = ""
            rsData.Filter = "编号=" & .RowData(i)
            strIniSys = Trim(rsData("版本号").value) '数据库版本号

            If strIniSys <> strIniApp Then
                MsgBox .TextMatrix(i, .ColIndex("系统名称")) & "数据库系统版本与配置文件版本不匹配。", vbInformation + vbOKOnly, "中联软件"
                GoTo errEnd
            End If
        Next
    End With
    If gobjFile.FolderExists(strTmp1) Then Call gobjFile.DeleteFolder(strTmp1, True)
    If gobjFile.FolderExists(strProcedure) Then Call gobjFile.DeleteFolder(strProcedure)
    If gobjFile.FolderExists(strTmpReports) Then gobjFile.DeleteFolder (strTmpReports)
        
    DoEvents
    
    Call gobjFile.CreateFolder(strTmpReports)
    Call gobjFile.CreateFolder(strTmp1)
    Call gobjFile.CreateFolder(strProcedure)
    
    '------------------------------------------------------------------------------------------------------------------
    '将数据库过程生成单个脚本文件
    lblTitle.Caption = "正在准备数据库.."
    strSQL = "Select Name,Type,Text From user_source Where type in ('PROCEDURE','FUNCTION') Order by Name,Line"
    Set rs = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "")
    
    If rs.BOF = False Then
        pbr.Visible = True
        
        Call objPercent.InitPercent(pbr, rs.RecordCount)
                
        lblTitle.Visible = True
        strFlag = ""
        For lngLoop = 0 To rs.RecordCount - 1
            If strFlag <> Nvl(rs("Name").value) And strFlag <> "" Then
            
                '判断文件名是否非法
                If Not (InStr(strFlag, "\") > 0 Or _
                    InStr(strFlag, "/") > 0 Or _
                    InStr(strFlag, ":") > 0 Or _
                    InStr(strFlag, " ") > 0 Or _
                    InStr(strFlag, "*") > 0 Or _
                    InStr(strFlag, "?") > 0 Or _
                    InStr(strFlag, """") > 0 Or _
                    InStr(strFlag, "<") > 0 Or _
                    InStr(strFlag, ">") > 0 Or _
                    InStr(strFlag, "|") > 0) Then
                    
                    '创建单个过程脚本文件
                    Set objFileTemp = gobjFile.CreateTextFile(strTmp1 & "\" & strFlag & ".sql", True)
                    
                    '写入上一个查询到的过程
                    Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                       strTemp = Left(strTemp, Len(strTemp) - 1)
                    Loop
                    
                    DoEvents
                    
                    strTemp = "CREATE OR REPLACE " & strTemp
                    objFileTemp.Write strTemp
                End If
                
                strTemp = ""
                strTemp = strTemp & UCase(Nvl(rs("Text").value))
                    
            ElseIf strFlag = "" Then
                
                strTemp = strTemp & UCase(Nvl(rs("Text").value))
                
            Else
                strTemp = strTemp & Nvl(rs("Text").value)
            End If
            
            strFlag = Nvl(rs("Name").value)
            rs.MoveNext
            
            Call objPercent.LoopPercent
        Next
        
        If strTemp <> "" Then
        
            '判断文件名是否非法
            If Not (InStr(strFlag, "\") > 0 Or _
                InStr(strFlag, "/") > 0 Or _
                InStr(strFlag, ":") > 0 Or _
                InStr(strFlag, " ") > 0 Or _
                InStr(strFlag, "*") > 0 Or _
                InStr(strFlag, "?") > 0 Or _
                InStr(strFlag, """") > 0 Or _
                InStr(strFlag, "<") > 0 Or _
                InStr(strFlag, ">") > 0 Or _
                InStr(strFlag, "|") > 0) Then
                
                '创建单个过程脚本文件
                Set objFileTemp = gobjFile.CreateTextFile(strTmp1 & "\" & strFlag & ".sql", True)
                
                '写入上一个查询到的过程
                Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                   strTemp = Left(strTemp, Len(strTemp) - 1)
                Loop
                
                DoEvents
                strTemp = "CREATE OR REPLACE " & strTemp
                objFileTemp.Write strTemp
            End If
            
            strTemp = ""
        End If
        objFileTemp.Close
        pbr.Visible = False
    End If
    
    
    '------------------------------------------------------------------------------------------------------------------
    For i = 1 To vsf(0).Rows - 1
        
        '提取安装脚本和升级脚本的过程再生成单个脚本文件
        '读取安装脚本
        If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本"))) Then
            MsgBox "无法打开脚本文件" & vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")) & ",执行中断。", vbExclamation, gstrSysName
            GoTo errEnd
        Else
            strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本"))) - 11)
            strIniPath = strIniPath & "zlProgram.sql"
        End If
        lblTitle.Caption = "正在提取“" & vsf(0).TextMatrix(i, vsf(0).ColIndex("系统名称")) & "”安装脚本.."
        
        Call CheckProcedure(strIniPath, strProcedure)
        
        pbr.value = 0
        pbr.Visible = False
        DoEvents
        
        '提取升级脚本
        strIniSys = vsf(0).TextMatrix(i, vsf(0).ColIndex("版本号"))
        If Split(strIniSys, ".")(2) = 0 Then
            GoTo errNext
        ElseIf Not gobjFile.FolderExists(Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")), "应用脚本")(0) & "升级脚本\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0") Then
            MsgBox "无法检测到升级脚本文件夹,执行中断。", vbExclamation, gstrSysName
            GoTo errEnd
        Else
            strIniPath = Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")), "应用脚本")(0) & "升级脚本\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0" & "\"
        End If
'        If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本"))) And vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")) <> "" Then
'            MsgBox "无法打开脚本文件" & vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")) & ",执行中断。", vbExclamation, gstrSysName
'            GoTo errEnd
'        ElseIf Trim(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本"))) = "" Then
'            GoTo errNext
'        Else
'            strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本"))) - 13)
'        End If

'        Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")))
'        If Not CheckINIValid(rsInit, "系统号|目标版本") Then
'            MsgBox "升迁配置文件格式不正确。", vbExclamation, "中联软件"
'            GoTo errEnd
'        End If
        lblTitle.Caption = "正在提取" & vsf(0).TextMatrix(i, vsf(0).ColIndex("系统名称")) & "升级脚本.."
'        rsInit.Filter = "项目='目标版本'"
'        intSysNumLast = Split(rsInit("内容").value, ".")(2) '得到升级配置文件的版本号
        intSysNumLast = Split(strIniSys, ".")(2)
        For lngLoop = 10 To intSysNumLast Step 10
            strFlag = Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & "." & CStr(lngLoop)
            Call CheckProcedure(strIniPath & "ZL" & vsf(0).RowData(i) / 100 & "_" & strFlag & ".sql", strProcedure)
        Next
errNext:

    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '将数据库中的过程与脚本进行比对，生成html报告
    strCommand = GetWinSystemPath & "\wincmp3.exe " & strTmp1 & "\ " & strProcedure & "\ /G:HE " & strTmpReports
    err.Clear
    DoEvents
    lblTitle.Caption = "正在比较.."
    lngTemp = Shell(strCommand, vbHide)
    DoEvents
    If err <> 0 Then
        err.Clear
         MsgBox "文件比较失败，请检查" & GetWinSystemPath & "\wincmp3.exe文件是否存在", vbExclamation, "中联软件"
        GoTo errEnd
    End If
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    CloseHandle lngProcess
    
    DoEvents
    
    '------------------------------------------------------------------------------------------------------------------
    
    Set objFolder = gobjFile.GetFolder(strTmpReports)
    
    '报告中存在的即为需要调整的过程
    For Each objFile In objFolder.Files
        Dim strFileName As String
        Dim lngKey As Long
        Dim strContent As String
        Dim lngMaxLength As Long
        Dim str As String
        Dim lngRow As Long
        Dim strArr() As String
        
        DoEvents
        
        strFileName = Split(objFile.name, ".")(0)
        lblTitle.Caption = "正在产生变动过程：" & strFileName
        
        
        '获取过程所有者
        strOwner = gclsBase.GetOwnerInfo(strFileName)
        
        Set rs = gclsBase.GetProInfo(strFileName)
        
        '过程不存在，自动添加为变动过程或空白过程
        '-------------------------------------------------------------------------------------------------------------------------------------------------
        If rs.BOF = False Then
            lngKey = Nvl(rs("ID").value)
            If rs("类型").value = 2 Then
                strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.空白过程 & ",'" & strFileName & "'," & ProcState.已调整 & ",'','" & strOwner & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            Else
                strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.变动过程 & ",'" & strFileName & "'," & ProcState.已调整 & ",'','" & strOwner & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            End If
        Else
            lngKey = gclsBase.GetNextId("zlProcedure")
            strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.变动过程 & ",'" & strFileName & "'," & ProcState.已调整 & ",'','" & strOwner & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        End If
        
        
        '保存本次自定义过程
        Set objFileTemp = gobjFile.OpenTextFile(strTmp1 & "\" & Split(objFile.name, ".")(0) & "." & Split(objFile.name, ".")(1))
        '读取过程内容
        strContent = ""
        Do While Not objFileTemp.AtEndOfStream
            strLine = objFileTemp.ReadLine
            If strContent = "" Then
                strContent = strContent & Replace(strLine, "'", "''")
            Else
                strContent = strContent & vbCrLf & Replace(strLine, "'", "''")
            End If
        Loop
        objFileTemp.Close
        lngMaxLength = 3900
        If LenB(StrConv(strContent, vbFromUnicode)) > lngMaxLength Then
            strFlag = ""
            str = ""
            For lngRow = 1 To Len(strContent)
                str = str & Mid(strContent, lngRow, 1)
                If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngRow = Len(strContent)) And Mid(strContent, lngRow, 1) <> "'" Then
                    strFlag = strFlag & gstrSplite & str
                    str = ""
                End If
            Next
            strFlag = Mid(strFlag, Len(gstrSplite) + 1)
            strContent = strFlag
        End If
        strArr = Split(strContent, gstrSplite)
        '先删除内容
        strSQL = "Zl_Zlproceduretext_Delete(" & lngKey & "," & ProcTextType.本次自定过程 & ")"
        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        '再插入内容
        For lngRow = 0 To UBound(strArr)
            strSQL = "Zl_Zlproceduretext_Update(" & lngKey & "," & ProcTextType.本次自定过程 & " ," & (lngRow + 1) & ",'" & strArr(lngRow) & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        Next
               
        
        '本次标准过程
        '-------------------------------------------------------------------------------------------------------------------------------------------------
        Set objFileTemp = gobjFile.OpenTextFile(strProcedure & "\" & Split(objFile.name, ".")(0) & "." & Split(objFile.name, ".")(1))
        '读取过程内容
        strContent = ""
        Do While Not objFileTemp.AtEndOfStream
            strLine = objFileTemp.ReadLine
            If strContent = "" Then
                strContent = strContent & Replace(strLine, "'", "''")
            Else
                strContent = strContent & vbCrLf & Replace(strLine, "'", "''")
            End If
        Loop
        objFileTemp.Close
        lngMaxLength = 3900
        If LenB(StrConv(strContent, vbFromUnicode)) > lngMaxLength Then
            strFlag = ""
            str = ""
            For lngRow = 1 To Len(strContent)
                str = str & Mid(strContent, lngRow, 1)
                If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngRow = Len(strContent)) And Mid(strContent, lngRow, 1) <> "'" Then
                    strFlag = strFlag & gstrSplite & str
                    str = ""
                End If
            Next
            strFlag = Mid(strFlag, Len(gstrSplite) + 1)
            strContent = strFlag
        End If
        strArr = Split(strContent, gstrSplite)
        '先删除内容
        strSQL = "Zl_Zlproceduretext_Delete(" & lngKey & "," & ProcTextType.本次标准过程 & ")"
        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        For lngRow = 0 To UBound(strArr)
            strSQL = "Zl_Zlproceduretext_Update(" & lngKey & "," & ProcTextType.本次标准过程 & "," & (lngRow + 1) & ",'" & strArr(lngRow) & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        Next
    Next
    
    '------------------------------------------------------------------------------------------------------------------
    lblTitle.Caption = "正在产生用户过程.."
    pbr.Visible = True
    Set objFolder = gobjFile.GetFolder(strTmp1)
        
    Call objPercent.InitPercent(pbr, objFolder.Files.Count)

    For Each objFile In objFolder.Files
        lblTitle.Caption = "正在产生用户过程：" & objFile.name
        DoEvents
                
        blnNew = False
        If gobjFile.FileExists(strProcedure & "\" & objFile.name) Then
            blnNew = True
        End If
        If blnNew = False Then
            
            '数据库中的过程在脚本中没有，说明是用户过程
            '添加用户过程

            strFileName = Split(objFile.name, ".")(0)
            Set rs = gclsBase.GetProInfo(strFileName)
            If rs.BOF = False Then
                lngKey = Nvl(rs("ID").value)
            Else
                '过程不存在，自动添加为用户过程
                lngKey = gclsBase.GetNextId("zlProcedure")
            End If
            Set objFileTemp = gobjFile.OpenTextFile(strTmp1 & "\" & Split(objFile.name, ".")(0) & "." & Split(objFile.name, ".")(1))
            
            '读取过程内容
            strContent = ""
            Do While Not objFileTemp.AtEndOfStream
                strLine = objFileTemp.ReadLine
                If strContent = "" Then
                    strContent = strContent & Replace(strLine, "'", "''")
                Else
                    strContent = strContent & vbCrLf & Replace(strLine, "'", "''")
                End If
            Loop
            
            
            
            objFileTemp.Close
            lngMaxLength = 3900
            If LenB(StrConv(strContent, vbFromUnicode)) > lngMaxLength Then
                strFlag = ""
                str = ""
                For lngRow = 1 To Len(strContent)
                    str = str & Mid(strContent, lngRow, 1)
                    If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngRow = Len(strContent)) And Mid(strContent, lngRow, 1) <> "'" Then
                        strFlag = strFlag & gstrSplite & str
                        str = ""
                    End If
                Next
                strFlag = Mid(strFlag, Len(gstrSplite) + 1)
                strContent = strFlag
            End If
            
            strArr = Split(strContent, gstrSplite)
            
            strOwner = gclsBase.GetOwnerInfo(strFileName)
            strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.用户过程 & ",'" & strFileName & "'," & ProcState.已调整 & ",'','" & strOwner & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            
            '先删除内容
            strSQL = "Zl_Zlproceduretext_Delete(" & lngKey & "," & ProcTextType.本次自定过程 & ")"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            For lngRow = 0 To UBound(strArr)
                strSQL = "Zl_Zlproceduretext_Update(" & lngKey & "," & ProcTextType.本次自定过程 & "," & (lngRow + 1) & ",'" & strArr(lngRow) & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            Next
            
        End If
        
        Call objPercent.LoopPercent
        
    Next
    
    lblTitle.Caption = "正在提交数据..."
    DoEvents
    
    On Error Resume Next
    objFileTemp.Close
    Set objFileTemp = Nothing
    On Error GoTo errHand
    
    Call SQLRecordExecute(rsSQL, "")
    
    lblTitle.Caption = "正在清除临时数据..."
    DoEvents
    '删除临时文件夹
    '------------------------------------------------------------------------------------------------------------------
    If gobjFile.FolderExists(strTmp1) Then
        Call gobjFile.DeleteFolder(strTmp1, True)
    End If
    If gobjFile.FolderExists(strProcedure) Then
        Call gobjFile.DeleteFolder(strProcedure, True)
    End If
    If gobjFile.FolderExists(strTmpReports) Then
        Call gobjFile.DeleteFolder(strTmpReports, True)
    End If
    
    MsgBox "搜索登记完成！", vbInformation, Me.Caption
    lblTitle.Visible = False
    pbr.Visible = False
    mblnOK = True
    cmdOK.Enabled = True
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errEnd:
    mblnOK = True
    cmdOK.Enabled = True
    Exit Sub
errHand:
    MsgBox "搜索登记失败！" & vbCrLf & err.Description, vbInformation, Me.Caption
    cmdOK.Enabled = True
End Sub

Private Sub Form_Load()
    Call ExecuteCommand("初始控件")
    
    Call opt_Click(0)
    
End Sub

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    Do While InStr(strText, "  ") > 0
        strText = Replace(strText, "  ", " ")
    Loop
    TrimEx = strText
End Function

Public Function TrimComment(ByVal strSQL As String) As String
'功能：去掉写在单行strSQL语句后面的"--"注释
'说明：主要是RunSQLFile的子函数
    Dim blnStr As Boolean
    Dim i As Long, k As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                k = i: Exit For
            End If
        Next
        If k > 0 Then strSQL = RTrim(Left(strSQL, k - 1))
    End If
    TrimComment = strSQL
End Function

Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'功能：将指定INI配置文件的内容读取到记录集中
'返回：Nothing或包含"项目,内容"的记录集,其中同一项目可能有多行内容
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "项目", adVarChar, 100
    rsTmp.Fields.Append "内容", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = Null
                rsTmp.Update
            End If
            
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!项目 = strItem
            rsTmp!内容 = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!项目 = strItem
        rsTmp!内容 = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Private Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'功能：检查对应的配置文件格式是否正确
'参数：rsINI=存放配置文件内容的记录集，包含"项目,内容"字段
'      strItem=配置文件中必须要求有内容的项目串,如"项目1|项目2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "项目='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If IsNull(rsINI!内容) Then Exit Function
    Next
    CheckINIValid = True
End Function

Private Function CheckProcedure(ByVal strFile As String, Optional strFilePath As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim lngLine As Long
    Dim strLine As String
    Dim strTemp As String
    Dim strFMT As String
    Dim blnSQL As Boolean
    Dim blnBlock As Boolean
    Dim strFlag As String
    Dim strFileProName As String
    Dim lngFileLines As Long
    Dim objFileTemp As TextStream
    Dim objFile As TextStream
    Dim blnFlag As Boolean
    Dim objPercent As New clsPercent
    Dim lngMsg As Long
    
    On Error GoTo errHand
    
    pbr.value = 0
    pbr.Visible = True

    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    If objFile.AtEndOfStream Then
        objFile.Close
        Exit Function
    End If
        
    Do While Not objFile.AtEndOfStream
        objFile.ReadLine
    Loop
    lngFileLines = objFile.Line
    
    Call objPercent.InitPercent(pbr, lngFileLines)
    
    objFile.Close
    
    Dim blnSpaceProc As Boolean
    
    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objFile.AtEndOfStream
        lngLine = objFile.Line '当前行号:未读取行之前,行指针未移到下一行
        strLine = objFile.ReadLine
        strFMT = UCase(TrimComment(TrimEx(strLine)))
        If strFMT Like "PROMPT *" Then GoTo NextLine
        
        
        If blnBlock Then
            If strFMT = "/" Then
                blnSQL = True
                blnBlock = False
                Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                   strTemp = Left(strTemp, Len(strTemp) - 1)
                Loop
                
                
                objFileTemp.Write "CREATE OR REPLACE " & strTemp
                DoEvents
                objFileTemp.Close
                strTemp = ""
                
                If blnSpaceProc = True Then
                    blnSpaceProc = False
                    
                    Set objFileTemp = gobjFile.OpenTextFile(strFilePath & "\" & strFileProName & ".sql")
                    strTemp = objFileTemp.ReadAll
                    objFileTemp.Close
                    strTemp = GetBlankProcedure(strTemp)
                    
                    DoEvents
                    Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
                    objFileTemp.Write strTemp
                    objFileTemp.Close
                    strTemp = ""
                End If
                
            Else
                strTemp = strTemp & vbCrLf & strLine
            End If
        ElseIf strFMT Like "CREATE OR REPLACE PROCEDURE *" Or strFMT Like "CREATE PROCEDURE *" _
            Or strFMT Like "CREATE OR REPLACE FUNCTION *" Or strFMT Like "CREATE FUNCTION *" _
            Or strFMT Like "CREATE OR REPLACE TRIGGER *" Or strFMT Like "CREATE TRIGGER *" _
            Or strFMT Like "CREATE OR REPLACE TYPE *" Or strFMT Like "CREATE TYPE *" _
            Or strFMT Like "CREATE OR REPLACE PACKAGE *" Or strFMT Like "CREATE PACKAGE *" Then
            
            blnBlock = True
            
            '创建单个过程脚本文件
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            
            If InStr(strFlag, "(") > 0 Then strFlag = Left(strFlag, InStr(strFlag, "(") - 1)
            If InStr(strFlag, ".") > 0 Then strFlag = Split(strFlag, ".")(1)
            strFileProName = Split(strFlag, " ")(1)
            If gobjFile.FileExists(strFilePath & "\" & strFileProName & ".sql") Then
                Call gobjFile.DeleteFile(strFilePath & "\" & strFileProName & ".sql")
            End If
            
            '检查是否为空白过程
            blnSpaceProc = False
            If IsSpaceProcedure("ZLHIS", strFileProName) = True Then
                blnSpaceProc = True
            End If
            
            Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
             
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            strTemp = strTemp & UCase(strFlag)
        End If
        
        Call objPercent.LoopPercent

NextLine:
    Loop
    objFile.Close
    pbr.Visible = False
    pbr.value = 0
'    MsgBox blnFlag
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Function GetWinPath() As String
    '--------------------------------------------------------------------------------------------------------------
    '--功能:获取系统目录
    '--------------------------------------------------------------------------------------------------------------
    Dim Buffer As String
    Dim StrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    StrWinPath = Left(Buffer, rtn)
    GetWinPath = StrWinPath
End Function

Private Sub Form_Resize()
    On Error Resume Next
    vsf(0).Move 15, 15, picPane(0).ScaleWidth - 30, picPane(0).ScaleHeight - 30
    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
    If Not (mfrmPageConfigure Is Nothing) Then
        Unload mfrmPageConfigure
    End If
'    Call InitCommon(gcnOracle)
End Sub

Private Sub mfrmPageConfigure_AfterConn(ByVal cnOracle As ADODB.Connection)
    Set mcnOracle = cnOracle
    
    Call ExecuteCommand("初始数据")
'    lblState.Caption = "已连接"
'    lblState.ForeColor = &HC000&
End Sub

Private Sub opt_Click(Index As Integer)
    cmdConnet.Enabled = (opt(1).value = True)
    
    Select Case Index
    Case 0
        Set mcnOracle = gcnOracle
        Call ExecuteCommand("初始数据")
'        lblState.Caption = "已连接"
'        lblState.ForeColor = &HC000&
    Case 1
        mclsVsf.ClearGrid
    End Select
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsf(0)
        Select Case Col
        '--------------------------------------------------------------------------------------------------------------
        Case .ColIndex("安装脚本")
            With dlg
                .DialogTitle = "选择应用安装配置文件"
                .Filter = "(应用安装配置文件)|zlSetup.ini"
                .ShowOpen
                If .FileName = "" Then
                    Exit Sub
                Else
                    vsf(0).TextMatrix(vsf(0).Row, vsf(0).Col) = .FileName
                End If
            End With
        End Select
        
        Call mclsVsf.SetFocus(, , True)
    End With
End Sub

Private Sub vsf_DblClick(Index As Integer)
    mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub




