VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoJobset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动作业设置"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "frmAutoJobset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtType 
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "用户自定义"
      Top             =   120
      Width           =   4245
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   480
      Width           =   4245
   End
   Begin VB.TextBox txtSys 
      Height          =   300
      Left            =   1320
      TabIndex        =   33
      Top             =   480
      Width           =   4245
   End
   Begin VB.PictureBox pic背景 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1320
      ScaleHeight     =   2145
      ScaleWidth      =   4050
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   4080
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编号100的系统中增加一个自动调价的作业"
         Height          =   180
         Index           =   2
         Left            =   450
         TabIndex        =   27
         Top             =   1740
         Width           =   3330
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZL100_USERJOB自动调价"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   26
         Top             =   1980
         Width           =   1890
      End
      Begin VB.Label lbl说明 
         BackStyle       =   0  'Transparent
         Caption         =   "命名规则中蓝体部分由用户输入；对服务器管理用户，不需要系统号。"
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   1005
         Width           =   3345
      End
      Begin VB.Label lbl标题 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "举例"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label lbl标题 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   810
         Width           =   390
      End
      Begin VB.Label lbl用户 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[系统号]        功能"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   690
         TabIndex        =   22
         Top             =   450
         Width           =   2100
      End
      Begin VB.Label lbl固定 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZL        _USERJOB"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   21
         Top             =   450
         Width           =   1890
      End
      Begin VB.Label lbl标题 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户自定作业命名规则"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   150
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新参数"
      Height          =   350
      Left            =   5850
      TabIndex        =   28
      Top             =   1560
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CheckBox chk规则 
      Caption         =   "命名规则"
      Height          =   350
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1140
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtJobName 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   4245
   End
   Begin VB.TextBox txtJobComment 
      ForeColor       =   &H00808080&
      Height          =   1230
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1590
      Width           =   4245
   End
   Begin VB.CommandButton cmdWhat 
      Caption         =   "…"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5280
      TabIndex        =   1
      Top             =   1200
      Width           =   285
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5850
      TabIndex        =   15
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5850
      TabIndex        =   14
      Top             =   120
      Width           =   1100
   End
   Begin VB.Frame fraPara 
      Caption         =   "执行参数"
      Height          =   840
      Left            =   1320
      TabIndex        =   12
      Top             =   4410
      Width           =   4245
      Begin VB.TextBox txtPara 
         Height          =   300
         Index           =   0
         Left            =   1035
         TabIndex        =   6
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label lblPara 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "登记时间"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   13
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame fraCycle 
      Caption         =   "执行周期"
      Height          =   1080
      Left            =   1320
      TabIndex        =   9
      Top             =   3255
      Width           =   4245
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   2100
         TabIndex        =   4
         Top             =   645
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   213647363
         UpDown          =   -1  'True
         CurrentDate     =   37031.0416666667
      End
      Begin VB.ComboBox cboMonth 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   645
         Width           =   900
      End
      Begin VB.ComboBox cboDay 
         Height          =   300
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   645
         Width           =   1030
      End
      Begin VB.ComboBox cboWeek 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   645
         Width           =   1030
      End
      Begin VB.ComboBox cboCycle 
         Height          =   300
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   225
         Width           =   720
      End
      Begin VB.TextBox txtCycle 
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label lblCycle 
         AutoSize        =   -1  'True
         Caption         =   "循环时间"
         Height          =   180
         Left            =   285
         TabIndex        =   11
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "执行时间"
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   705
         Width           =   720
      End
   End
   Begin VB.CheckBox chkAutoJob 
      Caption         =   "设置为后台自动作业(&A)"
      Height          =   210
      Left            =   1320
      TabIndex        =   3
      Top             =   2910
      Width           =   2850
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统"
      Height          =   180
      Left            =   900
      TabIndex        =   37
      Top             =   600
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分类"
      Height          =   180
      Left            =   900
      TabIndex        =   36
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "说明"
      Height          =   180
      Left            =   900
      TabIndex        =   17
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label lblJobWhat 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   4005
   End
   Begin VB.Label lblWhat 
      AutoSize        =   -1  'True
      Caption         =   "内容"
      Height          =   180
      Left            =   900
      TabIndex        =   16
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   255
      Picture         =   "frmAutoJobset.frx":000C
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblJob 
      AutoSize        =   -1  'True
      Caption         =   "作业"
      Height          =   180
      Left            =   900
      TabIndex        =   8
      Top             =   960
      Width           =   360
   End
   Begin VB.Menu mnuProcedures 
      Caption         =   "Procedure"
      Visible         =   0   'False
      Begin VB.Menu mnuWhat 
         Caption         =   "mnuWhat"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAutoJobset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum DateUnit
    DU_天 = 0
    DU_周 = 1
    DU_月 = 2
    DU_季度 = 3
End Enum

Private Enum vsfCol
    Col_系统 = 0
    Col_类别 = 1
    Col_序号 = 2
    Col_说明 = 3
    Col_参数 = 4
    Col_名称 = 5
    Col_调用过程 = 6
    Col_作业号 = 7
    Col_自动执行 = 8
    Col_状态 = 9
    Col_开始执行时间 = 10
    Col_间隔时间 = 11
    Col_系统编号 = 12
    Col_所有者
End Enum
Private mDateNow As Date
Private mstrPro As String

Private Sub cboCycle_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngDay As Long
    Dim lngMonth As Long
    Dim lngMaxDay As Long
    
    Select Case cboCycle.ListIndex
    Case DU_天
        cboMonth.Visible = False
        cboWeek.Visible = False
        cboDay.Visible = False
        dtpStart.Width = 2145
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        dtpStart.Left = txtCycle.Left
        
        If cboCycle.Text = cboCycle.Tag Then
            dtpStart.value = dtpStart.Tag
        Else
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_周
        cboMonth.Visible = False
        cboWeek.Visible = True
        cboDay.Visible = False
        dtpStart.Width = 1125
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboWeek.Left = txtCycle.Left
        dtpStart.Left = cboWeek.Left + cboWeek.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            cboWeek.ListIndex = Weekday(CDate(dtpStart.Tag)) - 1
            dtpStart.value = dtpStart.Tag
        Else
            cboWeek.ListIndex = 1
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_月
        cboMonth.Visible = False
        cboWeek.Visible = False
        cboDay.Visible = True
        dtpStart.Width = 1125
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboDay.Left = txtCycle.Left
        dtpStart.Left = cboDay.Left + cboDay.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            '获取指定月最大天数
            lngMaxDay = Right(DateSerial(Year(dtpStart.Tag), Month(dtpStart.Tag) + 1, 0), 2)
            lngDay = Format(dtpStart.Tag, "d")
            If lngDay <= 28 Then
                cboDay.Text = lngDay & "日"
            ElseIf lngDay = lngMaxDay Then
                cboDay.Text = "月末"
            ElseIf lngDay = lngMaxDay - 1 Then
                cboDay.Text = "月末-1"
            ElseIf lngDay = lngMaxDay - 2 Then
                cboDay.Text = "月末-2"
            End If
            dtpStart.value = dtpStart.Tag
        Else
            cboDay.ListIndex = 0
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_季度
        cboWeek.Visible = False
        cboMonth.Visible = True
        cboDay.Visible = True
        dtpStart.Width = 1125
        txtCycle.Width = 2310
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboMonth.Left = txtCycle.Left
        cboDay.Left = cboMonth.Left + cboMonth.Width - 20
        dtpStart.Left = cboDay.Left + cboDay.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            '获得指定月是第几个月
            lngMonth = Format(dtpStart.Tag, "M") Mod 3 - 1
            If lngMonth = 0 Then
                cboMonth.Text = "第一月"
            ElseIf lngMonth = 1 Then
                cboMonth.Text = "第二月"
            Else
                lngMonth = 2
                cboMonth.Text = "第三月"
            End If
            '获取指定月最大天数
            lngMaxDay = Right(DateSerial(Year(CDate(dtpStart.Tag)), Month(CDate(dtpStart.Tag)) + 1, 0), 2)
            lngDay = Format(dtpStart.Tag, "d")
            If lngDay <= 28 Then
                cboDay.Text = lngDay & "日"
            ElseIf lngDay = lngMaxDay Then
                cboDay.Text = "月末"
            ElseIf lngDay = lngMaxDay - 1 Then
                cboDay.Text = "月末-1"
            ElseIf lngDay = lngMaxDay - 2 Then
                cboDay.Text = "月末-2"
            End If
            dtpStart.value = dtpStart.Tag
        Else
            cboMonth.ListIndex = 0
            cboDay.ListIndex = 0
            dtpStart.value = "2001/5/20 1:00:00"
        End If
        
        '存入当前季度中第一月的月份
        cboMonth.Tag = Format(mDateNow, "M") - lngMonth
    End Select
End Sub

Private Sub chk规则_Click()
    pic背景.Visible = chk规则.value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strParas As String
    Dim lngCount As Long
    Dim strExecuteTime As String, strQuarterly As String
    Dim rsTemp As ADODB.Recordset
    Dim cnTools As ADODB.Connection
    Dim lngMaxDay As Long
    Dim lngSys As Long
    Dim strOrder As String
    Dim strSQL As String
    
    If Trim(lblJobWhat.Caption) = "" Then
        MsgBox "未设置作业内容！", vbExclamation, gstrSysName
        Exit Sub
    End If
    If Val(txtCycle.Text) = 0 Then
        MsgBox "未正确设置作业循环时间！", vbExclamation, gstrSysName
        txtCycle.SetFocus: Exit Sub
    End If
    
    strParas = ""
    If fraPara.Visible Then
        For lngCount = 0 To lblPara.UBound
            If lblPara(lngCount).Visible = False Then Exit For
            If Trim(txtPara(lngCount).Text) = "" Then
                MsgBox lblPara(lngCount).Caption & " 参数未指定值！", vbExclamation, gstrSysName
                Exit Sub
            End If
            strParas = strParas & ";" & lblPara(lngCount).Caption & "," & txtPara(lngCount).Text
        Next
    End If
    If strParas <> "" Then strParas = Mid(strParas, 2)
    
    '将获取到的执行日期信息转换为具体的日期
    Select Case cboCycle.ListIndex
    Case DU_天
        strExecuteTime = Format(mDateNow, "yyyy-MM-dd") & " " & Format(dtpStart.value, "HH:mm:ss")
    Case DU_周
        strExecuteTime = Format(DateAdd("d", cboWeek.ListIndex + 1 - Weekday(mDateNow), mDateNow), "yyyy-MM-dd") & " " & Format(dtpStart.value, "HH:mm:ss")
    Case DU_月
        If cboDay.ListIndex <= 27 Then
            strExecuteTime = Format(mDateNow, "yyyy-MM") & "-" & Val(cboDay.Text) & " " & Format(dtpStart.value, "HH:mm:ss")
        Else
            lngMaxDay = Right(DateSerial(Year(mDateNow), Month(mDateNow) + 1, 0), 2)
            strExecuteTime = Format(mDateNow, "yyyy-MM") & "-" & lngMaxDay - (cboDay.ListCount - cboDay.ListIndex - 1) & " " & Format(dtpStart.value, "HH:mm:ss")
        End If
    Case DU_季度
        If cboDay.ListIndex <= 27 Then
            strExecuteTime = Format(mDateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & Val(cboDay.Text) & " " & Format(dtpStart.value, "HH:mm:ss")
        Else
            strQuarterly = Format(mDateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & "01 11:11:11"
            lngMaxDay = Right(DateSerial(Year(CDate(strQuarterly)), Month(CDate(strQuarterly)) + 1, 0), 2)
            strExecuteTime = Format(mDateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & lngMaxDay - (cboDay.ListCount - cboDay.ListIndex - 1) & " " & Format(dtpStart.value, "HH:mm:ss")
        End If
    End Select
    
    If txtType.Text = "系统设定" Then
        lngCount = 1
    ElseIf txtType.Text = "数据转移" Then
        lngCount = 2
    Else
        lngCount = 3
    End If
            
    If Me.Tag = "ADD" Then
        If InStr(mstrPro, UCase(lblJobWhat.Caption) & ",") = 0 Then
            mstrPro = mstrPro & UCase(lblJobWhat.Caption) & ","
        End If
        lngSys = Val(cmbSystem.ItemData(cmbSystem.ListIndex))
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Job_number", lngSys)
        If rsTemp.RecordCount > 0 Then
            strOrder = Nvl(Val(rsTemp.Fields(0)), 1)
        Else
            strOrder = 1
        End If
        strSQL = "insert into zlAutoJobs(系统,类型,序号,名称,说明,内容,参数,执行时间,间隔时间,时间单位)" & _
                " values (" & IIf(lngSys = 0, "Null", lngSys) & ",3," & Val(strOrder) & "," & _
                "'" & txtJobName.Text & "'," & _
                "'" & txtJobComment.Text & "'," & _
                "'" & lblJobWhat.Caption & "'," & _
                " '" & strParas & "'," & _
                "to_date('" & strExecuteTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                "" & Val(txtCycle.Text) & _
                ",'" & cboCycle.Text & "')"
    Else
        strSQL = "update zlAutoJobs" & _
                " set 名称='" & txtJobName.Text & "'," & _
                "说明='" & txtJobComment.Text & "'," & _
                "内容='" & lblJobWhat.Caption & "'," & _
                "参数='" & strParas & "'," & _
                "执行时间=to_date('" & strExecuteTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                "间隔时间=" & Val(txtCycle.Text) & "," & _
                "时间单位='" & cboCycle.Text & "'" & _
                " Where Nvl(系统,0)=" & Val(lblSys.Tag) & _
                " and 类型=" & lngCount & _
                " and 序号=" & txtJobName.Tag
    End If
    err = 0
    On Error Resume Next
    gcnOracle.Execute strSQL
    If err <> 0 Then
        MsgBox "作业设置保存失败，请检查设置情况！" & vbNewLine & err.Description, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    err = 0
    If Me.Tag = "ADD" Then
        lngSys = lngSys
    Else
        lngSys = lblSys.Tag
    End If
    If lngSys = 0 Then
        Set cnTools = GetConnection("ZLTOOLS")
        If cnTools Is Nothing Then Exit Sub
    Else
        Set cnTools = gcnOracle
    End If
    If chkAutoJob.value = 1 Then
        If Tag = "ADD" Then                      '新作业
            strSQL = "zl_JobSubmit(" & IIf(lngSys = 0, "Null", lngSys) & ",3," & Val(strOrder) & ")"
        ElseIf Val(chkAutoJob.Tag) = 0 Then      '首次设置为自动作业
            strSQL = "zl_JobSubmit(" & IIf(lngSys = 0, "Null", lngSys) & "," & lngCount & "," & txtJobName.Tag & ")"
        Else                                        '修改已经启用的作业
            strSQL = "zl_JobChange(" & IIf(lngSys = 0, "Null", lngSys) & "," & lngCount & "," & txtJobName.Tag & ")"
        End If
        cnTools.Execute strSQL, , adCmdStoredProc
    Else
        If Val(chkAutoJob.Tag) <> 0 Then         '取消自动作业
            strSQL = "zl_JobRemove(" & IIf(lngSys = 0, "Null", lngSys) & "," & lngCount & "," & txtJobName.Tag & ")"
            cnTools.Execute strSQL, , adCmdStoredProc
        End If
    End If
    If err <> 0 Then
        MsgBox "虽然作业设置保存，但未能成功设置为自动作业。请检查数据库系统！" & vbNewLine & err.Description, vbExclamation, gstrSysName
    End If
    
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If MsgBox("是否根据数据归档转移处设置的时间更新参数？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_depict", Val(lblSys.Tag), Val(txtJobName.Tag))
    If rsTemp.RecordCount > 0 Then
        txtPara(0).Text = Val(IIf(IsNull(rsTemp.Fields(0)), "150", rsTemp.Fields(0)))
    Else
        txtPara(0).Text = 150
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdWhat_Click()
    Dim cnTools As ADODB.Connection
    Dim rsTemp As New ADODB.Recordset
    Dim lngSys As Long
    Dim strOwner As String
    Dim strProcedure As String
    Dim varTemp As Variant
    Dim strTemp As String
    Dim i As Long
   
    On Error GoTo errHandle
    
    lngSys = Val(cmbSystem.ItemData(cmbSystem.ListIndex))
    strOwner = Split(cmbSystem.Text, "|")(1)
    If Val(cmdWhat.Tag) = 0 Then
        If lngSys = 0 Then
            Set cnTools = GetConnection("ZLTOOLS")
            If cnTools Is Nothing Then Exit Sub
        Else
            Set cnTools = gcnOracle
        End If
        Set rsTemp = cnTools.Execute("SELECT Object_Name  From All_Objects " & vbNewLine & _
                                      "WHERE Object_Type = 'PROCEDURE' AND Object_Name LIKE 'ZL" & CStr(IIf(lngSys = 0, "", lngSys)) & "_USERJOB%' " & vbNewLine & _
                                      " AND Status = 'VALID' AND Owner = '" & strOwner & "'")
        With rsTemp
            Do While Not .EOF
                If .AbsolutePosition - 1 > mnuWhat.UBound Then Load mnuWhat(.AbsolutePosition - 1)
                mnuWhat(.AbsolutePosition - 1).Caption = .Fields(0).value
                mnuWhat(.AbsolutePosition - 1).Visible = True
                .MoveNext
            Loop
            cmdWhat.Tag = .RecordCount
        End With
    End If
    If Val(cmdWhat.Tag) > 0 Then
        PopupMenu mnuProcedures, 2
        varTemp = Split(mstrPro, ",")
        strTemp = " And (t.Name ='" & lblJobWhat.Caption & "'"
        For i = 0 To UBound(varTemp)
            If varTemp(i) <> "" Then
                strTemp = strTemp & " or t.Name='" & varTemp(i) & "'"
            End If
        Next
        strTemp = strTemp & ")"
        gstrSQL = "Select t.Name, upper(t.Text) Text" & vbNewLine & _
                "From User_Source t" & vbNewLine & _
                "Where t.Type = 'PROCEDURE' " & strTemp & " And Substr(Trim(t.Text), 1, 2) <> '--'" & vbNewLine & _
                "Order By t.Line"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption)
        rsTemp.Filter = "Name='" & lblJobWhat.Caption & "'"
        strTemp = ""
        Do While Not rsTemp.EOF
            strTemp = strTemp & rsTemp!Text & vbCrLf
            rsTemp.MoveNext
        Loop
        varTemp = Split(mstrPro, ",")
        For i = 0 To UBound(varTemp)
            If InStr(strTemp, varTemp(i)) > 0 And varTemp(i) <> "" Then
                strProcedure = strProcedure & varTemp(i) & ","
            End If
        Next
        If strProcedure <> "" Then
            MsgBox "[" & lblJobWhat.Caption & "]过程代码中存在已有作业的过程:" & strProcedure & "请检查是否添加！"
        End If
        strProcedure = ","
        rsTemp.Filter = "Name<>'" & lblJobWhat.Caption & "'"
        Do While Not rsTemp.EOF
            If InStr(rsTemp!Text, lblJobWhat.Caption) > 0 Then
                If InStr(strProcedure, "," & rsTemp!Name & ",") = 0 Then strProcedure = strProcedure & rsTemp!Name & ","
            End If
            rsTemp.MoveNext
        Loop
        If strProcedure <> "," Then
            MsgBox "[" & lblJobWhat.Caption & "]过程在已有作业的过程代码中存在:" & Mid(strProcedure, 2) & "请检查是否添加！"
        End If
    Else
        MsgBox "没有可选的存储过程", vbExclamation, gstrSysName
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Activate()
    Dim i As Long
    
    If frmAutoJobset.Tag = "2" Then cmdUpdate.Visible = True
    cboCycle.Clear
    cboCycle.AddItem "天"
    cboCycle.AddItem "周"
    cboCycle.AddItem "月"
    cboCycle.AddItem "季度"
    cboWeek.Clear
    cboWeek.AddItem "星期日"
    cboWeek.AddItem "星期一"
    cboWeek.AddItem "星期二"
    cboWeek.AddItem "星期三"
    cboWeek.AddItem "星期四"
    cboWeek.AddItem "星期五"
    cboWeek.AddItem "星期六"
    cboMonth.Clear
    cboMonth.AddItem "第一月"
    cboMonth.AddItem "第二月"
    cboMonth.AddItem "第三月"
    cboDay.Clear
    For i = 1 To 28
        cboDay.AddItem i & "日"
    Next
    cboDay.AddItem "月末-2"
    cboDay.AddItem "月末-1"
    cboDay.AddItem "月末"
    
    '将当前数据库时间存入变量
    mDateNow = CurrentDate()
    
    cboCycle.Text = IIf(cboCycle.Tag = "", "天", cboCycle.Tag)
End Sub

Public Sub RunSet(ByVal vsfJobs As VSFlexGrid)
'功能：运行设置
    Dim strParas As String
    Dim aryPara() As String
    Dim intCount As Integer

    With vsfJobs
        txtType.Enabled = True
        txtType.Text = .TextMatrix(.Row, Col_类别)
        txtType.Enabled = False
        txtSys.Visible = True
        txtSys.Enabled = True
        cmbSystem.Visible = False
        cmbSystem.Enabled = False
        lblSys.Tag = .TextMatrix(.Row, Col_系统编号)
        txtSys.Text = .TextMatrix(.Row, Col_系统)
        txtSys.Enabled = False
        txtJobName.Tag = .TextMatrix(.Row, Col_序号)                 '序号
        txtJobName.Text = .TextMatrix(.Row, Col_名称)                    '名称
        chkAutoJob.value = IIf(.TextMatrix(.Row, Col_自动执行) = "是", 1, 0) '自动执行
        If .TextMatrix(.Row, Col_状态) = "未知" Then
            chkAutoJob.Tag = 0                                '作业号
        Else
            chkAutoJob.Tag = .TextMatrix(.Row, Col_作业号)             '作业号
        End If
        lblJobWhat.Caption = .TextMatrix(.Row, Col_调用过程)             '内容
        txtJobComment.Text = .TextMatrix(.Row, Col_说明)            '说明
        dtpStart.value = IIf(.TextMatrix(.Row, Col_开始执行时间) = "", CurrentDate(), .TextMatrix(.Row, Col_开始执行时间)) '开始执行时间
        dtpStart.Tag = dtpStart.value
        txtCycle.Text = Val(.TextMatrix(.Row, Col_间隔时间))  '间隔时间
        cboCycle.Tag = Replace(.TextMatrix(.Row, Col_间隔时间), txtCycle.Text, "") '时间单位
        strParas = Trim(.TextMatrix(.Row, Col_参数))
    End With
    If vsfJobs.TextMatrix(vsfJobs.Row, Col_类别) = "用户自定义" Then                                         '类型
        Me.Tag = 3
        cmdWhat.Enabled = True
        txtJobComment.Locked = False
        txtJobComment.ForeColor = Me.ForeColor
    ElseIf vsfJobs.TextMatrix(vsfJobs.Row, Col_类别) = "数据转移" Then
        Me.Tag = 2
        fraPara.Enabled = False
    Else
        Me.Tag = 1
    End If
    
    If strParas = "" Then
        Me.Height = fraCycle.Top + fraCycle.Height + 600
        fraPara.Visible = False
    Else
        fraPara.Visible = True
        aryPara = Split(strParas, ";")
        For intCount = 0 To UBound(aryPara)
            If intCount > lblPara.UBound Then Load lblPara(intCount)
            If intCount > txtPara.UBound Then Load txtPara(intCount)
            lblPara(intCount).Top = intCount * 400 + 375
            txtPara(intCount).Top = intCount * 400 + 315
            lblPara(intCount).Left = txtPara(0).Left - lblPara(intCount).Width - 45
            txtPara(intCount).Left = txtPara(0).Left
            lblPara(intCount).Caption = Left(aryPara(intCount), InStr(1, aryPara(intCount), ",") - 1)
            txtPara(intCount).Text = Mid(aryPara(intCount), InStr(1, aryPara(intCount), ",") + 1)
            lblPara(intCount).Visible = True
            txtPara(intCount).Visible = True
        Next
        fraPara.Height = (UBound(aryPara) + 1) * 400 + 375
        Me.Height = fraPara.Top + fraPara.Height + 600
    End If
    Me.Show 1, frmMDIMain
End Sub

Public Sub Add(ByRef strPro As String)
'功能：新增
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    txtType.Enabled = True
    txtType.Text = "用户自定义"
    txtType.Enabled = False
    txtSys.Enabled = False
    txtSys.Visible = False
    cmbSystem.Visible = True
    cmbSystem.Enabled = True
    Me.Tag = "ADD"
    cmdWhat.Enabled = True
    chk规则.Visible = True
    txtJobComment.Locked = False
    txtJobComment.ForeColor = Me.ForeColor
    Me.Height = fraCycle.Top + fraCycle.Height + 600
    fraPara.Visible = False
    dtpStart.value = CurrentDate()
    If gblnDBA Then
        cmbSystem.AddItem "服务器管理工具" & "|" & "ZLTOOLS"
        cmbSystem.ItemData(cmbSystem.NewIndex) = 0
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", gstrUserName)
    End If

    Do Until rsTemp.EOF
        cmbSystem.AddItem rsTemp!名称 & "|" & rsTemp!所有者
        cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp!编号
        rsTemp.MoveNext
    Loop
    '缺省不为服务器管理工具
    If cmbSystem.ListCount > 0 Then
        For i = 0 To cmbSystem.ListCount - 1
            If cmbSystem.ItemData(i) <> 0 Then
                cmbSystem.ListIndex = i: Exit For
            End If
        Next
        If cmbSystem.ListIndex = -1 Then cmbSystem.ListIndex = 0
    End If
    mstrPro = strPro
    Me.Show 1, frmMDIMain
    strPro = mstrPro
End Sub

Private Sub mnuWhat_Click(Index As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    Dim aryPara As Variant
    
    On Error GoTo errHandle
    lblJobWhat.Caption = mnuWhat(Index).Caption
    With rsTemp
        If gblnDBA Then
            strSQL = "select rtrim(ltrim(upper(text))) from dba_source where name='" & mnuWhat(Index).Caption & "' and OWNER='" & Split(cmbSystem.Text, "|")(1) & "'"
        Else
            strSQL = "select rtrim(ltrim(upper(text))) from user_source where name='" & mnuWhat(Index).Caption & "'"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & " " & Replace(Replace(Replace(Replace(Trim(.Fields(0).value), vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
            If InStr(1, strSQL, " AS ") > 0 Then Exit Do
            If InStr(1, strSQL, " IS ") > 0 Then Exit Do
            If InStr(1, strSQL, ")AS ") > 0 Then Exit Do
            If InStr(1, strSQL, ")IS ") > 0 Then Exit Do
            If Right(strSQL, 3) = " AS" Then Exit Do
            If Right(strSQL, 3) = " IS" Then Exit Do
            If Right(strSQL, 3) = ")AS" Then Exit Do
            If Right(strSQL, 3) = ")IS" Then Exit Do
            .MoveNext
        Loop
        strSQL = Replace(Replace(Replace(Replace(strSQL, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
        If InStr(1, strSQL, "(") > 0 Then
            strSQL = Mid(strSQL, InStr(1, strSQL, "(") + 1)
            strSQL = Left(strSQL, InStr(1, strSQL, ")") - 1)
        Else
            strSQL = ""
        End If
        
        For intCount = 0 To lblPara.UBound
            lblPara(intCount).Visible = False
            txtPara(intCount).Visible = False
        Next
    
        If strSQL = "" Then
            Height = fraCycle.Top + fraCycle.Height + 600
            fraPara.Visible = False
        Else
            fraPara.Visible = True
            aryPara = Split(strSQL, ",")
            For intCount = 0 To UBound(aryPara)
                aryPara(intCount) = Trim(aryPara(intCount))
                If intCount > lblPara.UBound Then Load lblPara(intCount)
                If intCount > txtPara.UBound Then Load txtPara(intCount)
                lblPara(intCount).Top = intCount * 400 + 375
                txtPara(intCount).Top = intCount * 400 + 315
                lblPara(intCount).Left = txtPara(0).Left - lblPara(intCount).Width - 45
                txtPara(intCount).Left = txtPara(0).Left
                lblPara(intCount).Caption = Left(aryPara(intCount), InStr(1, aryPara(intCount), " ") - 1)
                txtPara(intCount).Text = ""
                lblPara(intCount).Visible = True
                txtPara(intCount).Visible = True
            Next
            fraPara.Height = (UBound(aryPara) + 1) * 400 + 375
            Height = fraPara.Top + fraPara.Height + 600
        End If
    
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub


