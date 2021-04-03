VERSION 5.00
Begin VB.Form frmReportSetup 
   BorderStyle     =   0  'None
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraReportSetup 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Frame fra 
         Height          =   1485
         Index           =   5
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   7695
         Begin VB.Frame fra 
            Height          =   1125
            Index           =   9
            Left            =   480
            TabIndex        =   47
            Top             =   270
            Width           =   2055
            Begin VB.CheckBox chkIgnorePosi 
               Caption         =   "忽略结果的阴阳性"
               Height          =   180
               Left            =   120
               TabIndex        =   50
               ToolTipText     =   "不记录和处理阴阳性。"
               Top             =   0
               Width           =   1800
            End
            Begin VB.CheckBox chkReportAfterResult 
               Caption         =   "无诊断内容为阴性"
               Height          =   180
               Left            =   120
               TabIndex        =   49
               ToolTipText     =   "书写报告时，没有录入诊断，则默认记录为阴性。"
               Top             =   720
               Width           =   1740
            End
            Begin VB.CheckBox chkDefaultPosi 
               Caption         =   "诊断结果默认阳性"
               Height          =   300
               Left            =   120
               TabIndex        =   48
               ToolTipText     =   "弹出阴阳性选择窗口，默认选择阳性。"
               Top             =   300
               Width           =   1815
            End
         End
         Begin VB.CheckBox chkConformDetermine 
            Caption         =   "符合情况判断"
            Height          =   180
            Left            =   2640
            TabIndex        =   46
            ToolTipText     =   "激活符合情况功能和菜单"
            Top             =   280
            Width           =   1455
         End
         Begin VB.CheckBox chkReportLevel 
            Caption         =   "报告质量等级"
            Height          =   180
            Left            =   2640
            TabIndex        =   45
            Top             =   657
            Width           =   1410
         End
         Begin VB.CheckBox chkImageLevel 
            Caption         =   "影像质量等级"
            Height          =   180
            Left            =   2640
            TabIndex        =   44
            Top             =   1035
            Width           =   1410
         End
         Begin VB.TextBox txtReportLevel 
            Height          =   270
            Left            =   4050
            TabIndex        =   43
            Text            =   "甲,乙"
            Top             =   600
            Width           =   1035
         End
         Begin VB.TextBox txtImageLevel 
            Height          =   270
            Left            =   4050
            TabIndex        =   42
            Text            =   "甲,乙"
            ToolTipText     =   "用于评定影像质量的登记，最多四个等级"
            Top             =   990
            Width           =   1035
         End
         Begin VB.Frame fra 
            Caption         =   "录入时机"
            Height          =   1150
            Index           =   6
            Left            =   5280
            TabIndex        =   38
            Top             =   240
            Width           =   2055
            Begin VB.OptionButton optResultInput 
               Caption         =   "诊断签名后"
               Height          =   240
               Index           =   0
               Left            =   210
               TabIndex        =   41
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optResultInput 
               Caption         =   "审核签名后"
               Height          =   240
               Index           =   1
               Left            =   210
               TabIndex        =   40
               Top             =   525
               Width           =   1230
            End
            Begin VB.OptionButton optResultInput 
               Caption         =   "报告打印前"
               Height          =   240
               Index           =   2
               Left            =   210
               TabIndex        =   39
               Top             =   810
               Width           =   1290
            End
         End
      End
      Begin VB.Frame fraEditorSetUp 
         Caption         =   "报告文档编辑器设置"
         Height          =   4215
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   7695
         Begin VB.Frame Frame8 
            Caption         =   "查看历史报告"
            Height          =   1215
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   7215
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "PACS报告编辑器"
               Height          =   255
               Index           =   1
               Left            =   4080
               TabIndex        =   32
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "电子病历编辑器"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   31
               Top             =   600
               Value           =   -1  'True
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "报告编辑器"
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   7730
         Begin VB.OptionButton optReportEditor 
            Caption         =   "PACS智能报告编辑器"
            Height          =   255
            Index           =   2
            Left            =   4680
            TabIndex        =   28
            Top             =   240
            Width           =   1932
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "电子病历编辑器"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "PACS报告编辑器"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "报告设置"
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   7695
         Begin VB.Frame Frame7 
            Caption         =   "打印格式选择方式"
            Height          =   1335
            Left            =   4440
            TabIndex        =   33
            Top             =   1800
            Width           =   2895
            Begin VB.CheckBox chkPrintFormat 
               Caption         =   "单选报告格式"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   960
               Width           =   2295
            End
            Begin VB.OptionButton optPrintFormat 
               Caption         =   "始终保持默认格式"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   35
               Top             =   600
               Width           =   2415
            End
            Begin VB.OptionButton optPrintFormat 
               Caption         =   "记录最后一次打印格式"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   34
               Top             =   320
               Value           =   -1  'True
               Width           =   2415
            End
         End
         Begin VB.CheckBox chkUntreadPrinted 
            Caption         =   "审核打印后允许回退"
            Height          =   180
            Left            =   480
            TabIndex        =   27
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkSpecialContent 
            Caption         =   "显示专科报告内容："
            Height          =   180
            Left            =   480
            TabIndex        =   23
            Top             =   1080
            Width           =   2055
         End
         Begin VB.ComboBox cboSpecialContent 
            Height          =   300
            Left            =   480
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   1440
            Width           =   6855
         End
         Begin VB.CheckBox chkExitAfterPrint 
            Caption         =   "打印后退出"
            Height          =   180
            Left            =   2760
            TabIndex        =   21
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "报告文本段名称"
            Height          =   1335
            Left            =   480
            TabIndex        =   14
            Top             =   1800
            Width           =   3255
            Begin VB.TextBox txtAdvice 
               Height          =   270
               Left            =   1560
               TabIndex        =   17
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtResult 
               Height          =   270
               Left            =   1560
               TabIndex        =   16
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtCheckView 
               Height          =   270
               Left            =   1560
               TabIndex        =   15
               Top             =   225
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "建    议："
               Height          =   255
               Left            =   360
               TabIndex        =   20
               Top             =   975
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "诊断意见："
               Height          =   255
               Left            =   360
               TabIndex        =   19
               Top             =   615
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "检查所见："
               Height          =   255
               Left            =   360
               TabIndex        =   18
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CheckBox chkShowVideoCapture 
            Caption         =   "显示视频采集区域"
            Height          =   180
            Left            =   2760
            TabIndex        =   13
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtMinImageCount 
            Height          =   270
            Left            =   6840
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "8"
            Top             =   315
            Width           =   495
         End
         Begin VB.CheckBox chkShowImage 
            Caption         =   "显示报告图像区域                               报告缩略图显示数量："
            CausesValidation=   0   'False
            Height          =   180
            Left            =   480
            TabIndex        =   11
            Top             =   360
            Width           =   6375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "报告词句双击后"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "直接写入报告"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "打开词句编辑窗口"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   1750
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "缩略图双击后"
         Height          =   855
         Left            =   2520
         TabIndex        =   4
         Top             =   5640
         Width           =   2895
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "打开图片编辑窗口"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   1750
         End
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "直接写入报告"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "词句模板显示"
         Height          =   855
         Left            =   5400
         TabIndex        =   1
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optShowWord 
            Caption         =   "双击标题"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optShowWord 
            Caption         =   "直接显示"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmReportSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptId As Long   '科室ID
Private mblnRefreshed As Boolean

Public Sub zlRefresh(lngDeptID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
    Dim lngHintType As Long
    
    mblnRefreshed = True            '数据被刷新过了，可以保存
    
    mlngDeptId = lngDeptID
    optReportEditor(0).value = True '默认使用电子病历编辑器编辑报告
    chkShowImage.value = 0          '默认不显示图像区域
    chkShowVideoCapture.value = 0   '默认不显示视频采集区域
    
    chkSpecialContent.value = 0     '默认不显示专科报告
    cboSpecialContent.Enabled = False
    chkExitAfterPrint.value = 0     '默认打印后不退出
    optWordDblClick(0).value = True '默认双击词句后直接写入报告
    optImageDblClick(0).value = True '默认报告缩略图双击后直接写入报告
    txtCheckView.Text = "检查所见"  '默认为检查所见
    txtResult.Text = "诊断意见"     '默认为诊断意见
    txtAdvice.Text = "建议"         '默认为建议
    optShowWord(0).value = True     '默认为直接显示词句模板
    chkUntreadPrinted.value = 0     '默认为审核打印后不允许回退
    
    chkIgnorePosi.value = 0     '忽略结果阴阳性
    chkReportAfterResult.value = 0 '无影像诊断为阴性
    chkDefaultPosi.value = 0        '诊断结果默认阳性为未勾选
    chkConformDetermine.value = 1       '符合情况判定默认为选中
    txtImageLevel.Text = "甲,乙"     '默认影像质量等级
    txtReportLevel.Text = "甲,乙"    '默认报告质量等级
    
    On Error GoTo err
    strSQL = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId)
    
    While Not rsTemp.EOF
        Select Case rsTemp!参数名
            Case "报告编辑器"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optReportEditor(0).value = True
                ElseIf Nvl(rsTemp!参数值, 0) = 1 Then
                    optReportEditor(1).value = True
                Else
                    optReportEditor(2).value = True
                End If
            Case "查看历史报告"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optHistoryReportEditor(0).value = True
                Else
                    optHistoryReportEditor(1).value = True
                End If
                
            Case "显示报告图像"
                chkShowImage.value = Nvl(rsTemp!参数值, 0)
            Case "报告缩略图数量"
                txtMinImageCount.Text = Nvl(rsTemp!参数值, "8")
            Case "显示视频采集"
                chkShowVideoCapture.value = Nvl(rsTemp!参数值, 0)
            Case "打印后退出"
                chkExitAfterPrint.value = Nvl(rsTemp!参数值, 0)

            Case "显示专科报告"
                chkSpecialContent.value = Nvl(rsTemp!参数值, 0)
                cboSpecialContent.Enabled = IIf(chkSpecialContent.value = 1, True, False)
            Case "专科报告页"
                cboSpecialContent.Text = Nvl(rsTemp!参数值)
            Case "报告词句双击操作"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optWordDblClick(0).value = True
                Else
                    optWordDblClick(1).value = True
                End If
            Case "缩略图双击操作"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optImageDblClick(0).value = True
                Else
                    optImageDblClick(1).value = True
                End If
            Case "检查所见名称"
                txtCheckView.Text = Nvl(rsTemp!参数值, "检查所见")
            Case "诊断意见名称"
                txtResult.Text = Nvl(rsTemp!参数值, "诊断意见")
            Case "建议名称"
                txtAdvice.Text = Nvl(rsTemp!参数值, "建议")
            Case "显示词句示范"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optShowWord(0).value = True
                Else
                    optShowWord(1).value = True
                End If
            Case "审核打印后允许回退"
                chkUntreadPrinted.value = Nvl(rsTemp!参数值, 0)
            Case "打印格式选择方式"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optPrintFormat(0).value = True
                Else
                    optPrintFormat(1).value = True
                End If
            Case "单选报告格式"
                    chkPrintFormat.value = IIf(Nvl(rsTemp!参数值, 0), 1, 0)
            Case "诊断结果提示类型"
                lngHintType = Nvl(rsTemp!参数值, 0)
                optResultInput(lngHintType).value = True
            Case "诊断结果默认阳性"
                chkDefaultPosi.value = Nvl(rsTemp!参数值, 0)
            Case "无影像诊断为阴性"
                chkReportAfterResult.value = Nvl(rsTemp!参数值, 0)
            Case "忽略结果阴阳性"
                chkIgnorePosi.value = Nvl(rsTemp!参数值, 0)
            Case "符合情况判定"
                chkConformDetermine.value = Nvl(rsTemp!参数值, 0)
            Case "影像质量判定"
                chkImageLevel.value = Nvl(rsTemp!参数值, 0)
            Case "影像质量等级"
                txtImageLevel.Text = Nvl(rsTemp!参数值, "甲,乙")
                txtImageLevel.Enabled = chkImageLevel.value = 1
            Case "报告质量判定"
                chkReportLevel.value = Nvl(rsTemp!参数值, 0)
            Case "报告质量等级"
                txtReportLevel.Text = Nvl(rsTemp!参数值, "甲,乙")
                txtReportLevel.Enabled = chkReportLevel.value = 1
        End Select
        rsTemp.MoveNext
    Wend
    
    If optReportEditor(2).value Then
        fraEditorSetUp.Visible = True
        
    Else
        fraEditorSetUp.Visible = False
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Public Sub zlSave()
    Dim intMatch As Integer
    Dim strSQL As String
    Dim intTxtLen As Integer
    
    On Error GoTo errHand
    
    If mblnRefreshed = False Then Exit Sub          '数据没有被刷新，所以不保存
    
    If txtImageLevel.Enabled Then
        '将中文状态下的 逗号替换成英文状态
        txtImageLevel.Text = Replace(txtImageLevel.Text, "，", ",")
        
        intTxtLen = Len(txtImageLevel.Text) - Len(Replace(txtImageLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "影像等级最少为2种，最多为4种，请重新填写。", vbOKOnly, "提示信息"
            txtImageLevel.Text = Nvl(GetDeptPara(mlngDeptId, "影像质量等级", "甲,乙"))
            txtImageLevel.SetFocus
            Exit Sub
        End If
    End If
    
    
    If txtReportLevel.Enabled Then
        '将中文状态下的 逗号替换成英文状态
        txtReportLevel.Text = Replace(txtReportLevel.Text, "，", ",")
        
        intTxtLen = Len(txtReportLevel.Text) - Len(Replace(txtReportLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "报告等级最少为2种，最多为4种，请重新填写。", vbOKOnly, "提示信息"
            txtReportLevel.Text = Nvl(GetDeptPara(mlngDeptId, "报告质量等级", "甲,乙"))
            txtReportLevel.SetFocus
            Exit Sub
        End If
    End If
    
    If optReportEditor(0).value = True Then         '电子病历编辑器
        intMatch = 0
    ElseIf optReportEditor(1).value = True Then     'PACS报告编辑器
        intMatch = 1
    ElseIf optReportEditor(2).value = True Then     '报告文档编辑器
        intMatch = 2
    End If
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '报告编辑器','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '显示报告图像','" & chkShowImage.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '报告缩略图数量','" & txtMinImageCount.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '显示视频采集','" & chkShowVideoCapture.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '打印后退出','" & chkExitAfterPrint.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '显示专科报告','" & chkSpecialContent.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '专科报告页','" & cboSpecialContent.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optWordDblClick(0).value = True Then         '报告词句双击后直接写入报告
        intMatch = 0
    ElseIf optWordDblClick(1).value = True Then     '报告词句双击后打开编辑窗口
        intMatch = 1
    End If
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '报告词句双击操作','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optImageDblClick(0).value = True Then         '缩略图双击后直接写入报告
        intMatch = 0
    ElseIf optImageDblClick(1).value = True Then     '缩略图双击后打开图像编辑窗口
        intMatch = 1
    End If
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '缩略图双击操作','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '检查所见名称','" & txtCheckView.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '诊断意见名称','" & txtResult.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '建议名称','" & txtAdvice.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optShowWord(0).value = True Then         '直接显示词句示范
        intMatch = 0
    ElseIf optShowWord(1).value = True Then     '双击标题后显示词句示范
        intMatch = 1
    End If
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '显示词句示范','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '审核打印后允许回退','" & chkUntreadPrinted.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optReportEditor(2) Then
        strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '查看历史报告','" & IIf(optHistoryReportEditor(0).value, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '打印格式选择方式','" & IIf(optPrintFormat(0).value, 0, 1) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '单选报告格式','" & IIf(chkPrintFormat.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '符合情况判定','" & IIf(chkConformDetermine.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '忽略结果阴阳性','" & IIf(chkIgnorePosi.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '无影像诊断为阴性','" & IIf(chkReportAfterResult.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '诊断结果默认阳性','" & IIf(chkDefaultPosi.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '影像质量判定','" & IIf(chkImageLevel.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '影像质量等级','" & txtImageLevel.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '报告质量判定','" & IIf(chkReportLevel.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '报告质量等级','" & txtReportLevel.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_影像流程参数_UPDATE( " & mlngDeptId & ", '诊断结果提示类型','" & IIf(optResultInput(0).value = True, 0, IIf(optResultInput(1).value = True, 1, 2)) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub chkSpecialContent_Click()
    If chkSpecialContent.value = 1 Then
        cboSpecialContent.Enabled = True
    Else
        cboSpecialContent.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    mblnRefreshed = False
    '装载专科报告名称
    cboSpecialContent.Clear
    cboSpecialContent.AddItem (Report_Form_frmReportES)
    cboSpecialContent.AddItem (Report_Form_frmReportPathology)
    cboSpecialContent.AddItem (Report_Form_frmReportUS)
    cboSpecialContent.AddItem (Report_Form_frmReportCustom)
End Sub

Private Sub Form_Resize()
    fraReportSetup.Left = (Me.ScaleWidth - fraReportSetup.Width) / 2
End Sub


Private Sub optReportEditor_Click(Index As Integer)
    Dim hService As Long
    Dim hSCManager As Long

On Error GoTo errHandle

    fraEditorSetUp.Visible = Index = 2
    
    Exit Sub
errHandle:
    
End Sub

Private Sub chkImageLevel_Click()
    txtImageLevel.Enabled = chkImageLevel.value = 1
End Sub

Private Sub chkReportAfterResult_Click()
    If chkReportAfterResult.value = vbChecked Then
        chkIgnorePosi.Enabled = False
        chkIgnorePosi.value = vbUnchecked
    Else
        chkIgnorePosi.Enabled = True
    End If
End Sub

Private Sub chkReportLevel_Click()
    txtReportLevel.Enabled = chkReportLevel.value = 1
End Sub
