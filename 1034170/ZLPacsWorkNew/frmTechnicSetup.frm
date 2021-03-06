VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTechnicSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmTechnicSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicAction 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6900
      Left            =   120
      ScaleHeight     =   6870
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chkFormat 
         Caption         =   "报到打印前选择格式"
	     Enabled         =   0   'False
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         ToolTipText     =   "病人报到后自动打印申请单。"
         Top             =   3720
         Width           =   2025
      End
      Begin VB.Frame fraValid 
         Caption         =   "图像校对"
         Height          =   800
         Left            =   120
         TabIndex        =   48
         Top             =   5400
         Width           =   6375
         Begin VB.CheckBox chkValid 
            Caption         =   "打开影像工作站时自动校对前一天的图像"
            Height          =   375
            Left            =   240
            TabIndex        =   49
            ToolTipText     =   "打开工作站时自动校对"
            Top             =   240
            Width           =   5130
         End
      End
      Begin VB.ComboBox cbxMoneyExeModle 
         Height          =   300
         ItemData        =   "frmTechnicSetup.frx":000C
         Left            =   1410
         List            =   "frmTechnicSetup.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4935
         Width           =   2205
      End
      Begin VB.Frame frmPatholParameter 
         Height          =   3855
         Left            =   105
         TabIndex        =   21
         Top             =   30
         Width           =   3495
         Begin VB.TextBox txtMoleculeReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   35
            Top             =   2595
            Width           =   1335
         End
         Begin VB.TextBox txtSpecialStainReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   34
            Top             =   2235
            Width           =   1335
         End
         Begin VB.TextBox txtImmuneReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   33
            Top             =   1875
            Width           =   1335
         End
         Begin VB.TextBox txtNormalReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   32
            Top             =   1515
            Width           =   1335
         End
         Begin VB.TextBox txtGrossDescribe 
            Height          =   270
            Left            =   1680
            TabIndex        =   31
            Top             =   1155
            Width           =   1335
         End
         Begin VB.CheckBox chkIsDirectPrint 
            Caption         =   "是否直接打印"
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox chkKuaiPian 
            Caption         =   "快片"
            Height          =   375
            Left            =   2760
            TabIndex        =   29
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkShiJian 
            Caption         =   "尸检"
            Height          =   375
            Left            =   2280
            TabIndex        =   28
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkHuiZhen 
            Caption         =   "会诊"
            Height          =   375
            Left            =   1800
            TabIndex        =   27
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkXiBao 
            Caption         =   "细胞"
            Height          =   375
            Left            =   1320
            TabIndex        =   26
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkBingDong 
            Caption         =   "冰冻"
            Height          =   375
            Left            =   840
            TabIndex        =   25
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkChangGui 
            Caption         =   "常规"
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   3330
            Width           =   495
         End
         Begin VB.TextBox txtDecalinHintTime 
            Height          =   270
            Left            =   1800
            TabIndex        =   23
            Text            =   "30"
            Top             =   500
            Width           =   495
         End
         Begin VB.CheckBox chkDecalin 
            Caption         =   "脱钙任务声音提醒"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "当脱钙时间到了会有声音提示。"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblSpecialStainReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "特染报告模板："
            Height          =   180
            Left            =   360
            TabIndex        =   43
            Top             =   2280
            Width           =   1260
         End
         Begin VB.Label lblMoleculeReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分子报告模板："
            Height          =   180
            Left            =   360
            TabIndex        =   42
            Top             =   2640
            Width           =   1260
         End
         Begin VB.Label lblImmuneReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "免疫报告模板："
            Height          =   180
            Left            =   360
            TabIndex        =   41
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label lblNormalReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "常规报告模板："
            Height          =   180
            Left            =   360
            TabIndex        =   40
            Top             =   1560
            Width           =   1260
         End
         Begin VB.Label lblGrossDescribe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "巨检描述模板："
            Height          =   180
            Left            =   360
            TabIndex        =   39
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "对应词句分类"
            Height          =   180
            Left            =   1800
            TabIndex        =   38
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label labHint 
            Caption         =   "当以下检查完成时自动弹出质量窗口："
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   3060
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "提醒间隔时长：      秒"
            Height          =   255
            Left            =   600
            TabIndex        =   36
            Top             =   550
            Width           =   2055
         End
      End
      Begin VB.CommandButton CmdDevSet 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   1320
         TabIndex        =   19
         Top             =   6390
         Width           =   1170
      End
      Begin VB.CheckBox ChkOpenReport 
         Caption         =   "开始检查后自动打开报告窗口"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         ToolTipText     =   "在报到后自动打开报告窗口。"
         Top             =   2690
         Width           =   2640
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5235
         TabIndex        =   17
         Top             =   6390
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3975
         TabIndex        =   16
         Top             =   6390
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   60
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6390
         Width           =   1100
      End
      Begin VB.CheckBox chkPatTrack 
         Caption         =   "病人状态跟踪"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         ToolTipText     =   "在对病人一系列的检查过程中，始终保持当前选中的是同一个病人。"
         Top             =   2176
         Width           =   2640
      End
      Begin VB.CheckBox chkBatchInput 
         Caption         =   "连续输入申请"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         ToolTipText     =   "允许连续登记。"
         Top             =   634
         Width           =   2640
      End
      Begin VB.CheckBox chkView 
         Caption         =   "填写报告时打开观片站"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         ToolTipText     =   "打开报告书写窗口时自动打开观片站。"
         Top             =   1662
         Width           =   2280
      End
      Begin VB.Frame Frame6 
         Height          =   30
         Index           =   0
         Left            =   -120
         TabIndex        =   11
         Top             =   6240
         Width           =   6615
      End
      Begin VB.CheckBox chkCancelCheck 
         Caption         =   "不显示被取消的登记"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         ToolTipText     =   "在病人检查列表中不显示已经被取消的登记记录。"
         Top             =   1148
         Width           =   2640
      End
      Begin VB.CommandButton cmd3DSetup 
         Caption         =   "3D设置"
         Height          =   350
         Left            =   2640
         TabIndex        =   9
         Top             =   6390
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CheckBox chkAutoPrint 
         Caption         =   "报到后自动打印申请单"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         ToolTipText     =   "病人报到后自动打印申请单。"
         Top             =   120
         Width           =   2100
      End
      Begin VB.CheckBox chkExitAfterSign 
         Caption         =   "签名后退出"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         ToolTipText     =   "报告签名后自动退出报告书写。"
         Top             =   3204
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Caption         =   "登记模式"
         Height          =   880
         Left            =   120
         TabIndex        =   2
         Top             =   3900
         Width           =   3495
         Begin VB.CheckBox chkPasv 
            Caption         =   "启用被动传输"
            Height          =   255
            Left            =   1320
            TabIndex        =   47
            Top             =   240
            Width           =   1380
         End
         Begin VB.CheckBox chkInputOutInfo 
            Caption         =   "录入外院信息"
            Height          =   255
            Left            =   1710
            TabIndex        =   46
            ToolTipText     =   "在登记窗口录入送检单位和送检医生。"
            Top             =   865
            Width           =   1590
         End
         Begin VB.OptionButton optCheckInMode 
            Caption         =   "精简模式"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   6
            ToolTipText     =   "只显示和录入必要项目。"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optCheckInMode 
            Caption         =   "正常模式"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   5
            ToolTipText     =   "显示和录入全部项目。"
            Top             =   570
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CheckBox chkAddons 
            Caption         =   "不显示附加主述"
            Height          =   255
            Left            =   1710
            TabIndex        =   4
            ToolTipText     =   "在登记报到窗口不显示附加主述一项。"
            Top             =   540
            Width           =   1590
         End
         Begin VB.CheckBox chkReagent 
            Caption         =   "不显示造影剂"
            Height          =   255
            Left            =   1710
            TabIndex        =   3
            ToolTipText     =   "在登记报到窗口不显示造影剂一项。"
            Top             =   225
            Width           =   1680
         End
      End
      Begin VB.ComboBox cbxMainPage 
         Height          =   300
         ItemData        =   "frmTechnicSetup.frx":0041
         Left            =   3840
         List            =   "frmTechnicSetup.frx":0043
         TabIndex        =   1
         Top             =   4485
         Width           =   2655
      End
      Begin MSComDlg.CommonDialog dlgFont 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label labMoneyExemodel 
         Caption         =   "费用执行模式："
         Height          =   270
         Left            =   120
         TabIndex        =   44
         Top             =   4980
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "主要工作页面："
         Height          =   180
         Left            =   3840
         TabIndex        =   20
         Top             =   4200
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmTechnicSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mlng科室ID As Long 'IN:当前执行科室ID
Public mblnOK As Boolean
Public mlngModul As Long
Public mstrPrivs As String '模块权限

'病人列表字体设置
Private mTitleFont As StdFont       '病人列表表头字体
Private mTextFont As StdFont        '病人列表内容字体

Private Sub chkAutoPrint_Click()
    chkFormat.Enabled = chkAutoPrint.value = 1
End Sub


Private Sub cmd3DSetup_Click()
    frm3DSetup.ShowMe Me, mstrPrivs
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDevSet_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1101)
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOk_Click()
On Error GoTo errHandle
    Dim strPar As String, i As Long
    '计数变量
    Dim GrossNum As Long, NormalNum As Long, ImmuneNum As Long, SpecialNum As Long, MoleculeNum As Long
    Dim strSQL As String
    Dim rsExpression As ADODB.Recordset
    

    zlDatabase.SetPara "工作首页", cbxMainPage.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0

    zlDatabase.SetPara "报告时观片", chkView.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "连续登记申请", chkBatchInput.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "病人跟踪", chkPatTrack.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "报到打印前选择格式", chkFormat.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    
    zlDatabase.SetPara "开始检查自动打开报告", ChkOpenReport.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "不显示造影剂", chkReagent.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "不显示附加主述", chkAddons.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "录入外院信息", chkInputOutInfo.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "登记模式", IIf(optCheckInMode(1).value = True, 1, 2), glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "不显示被取消的登记", chkCancelCheck.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "报到后自动打印申请单", chkAutoPrint.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    
    SaveSetting "ZLSOFT", "公共模块\ZL9PACSWork\frmTechnicSetup", "启用被动传输", IIf(chkPasv.value = 1, 1, 0)
    Call zlDatabase.SetPara("PACS报告签名后退出", chkExitAfterSign.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    
    If mlngModul = 1291 Then
        Call zlDatabase.SetPara("采集费用执行模式", cbxMoneyExeModle.ListIndex, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    ElseIf mlngModul = 1290 Then
        Call zlDatabase.SetPara("医技费用执行模式", cbxMoneyExeModle.ListIndex, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    Else
        Call zlDatabase.SetPara("病理费用执行模式", cbxMoneyExeModle.ListIndex, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    End If
    
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        '保存病理系统相关参数
        zlDatabase.SetPara "脱钙声音提醒", chkDecalin.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "提醒间隔时长", txtDecalinHintTime.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "常规质量窗口", chkChangGui.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "快速石蜡质量窗口", chkKuaiPian.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "冰冻质量窗口", chkBingDong.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "细胞质量窗口", chkXiBao.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "会诊质量窗口", chkHuiZhen.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "尸检质量窗口", chkShiJian.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        
        '保存是否直接打印设置
        zlDatabase.SetPara "是否直接打印", chkIsDirectPrint.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
        
        strSQL = "select 名称 from 病历词句分类"
        Set rsExpression = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

        For i = 1 To rsExpression.RecordCount
            
            If txtGrossDescribe.Text <> "" Then
                '如果用户输入的分类和数据库匹配 则将参数保存到数据库中
                If rsExpression("名称").value = txtGrossDescribe.Text Then
                    '执行保存参数
                    zlDatabase.SetPara "巨检描述模板", txtGrossDescribe.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
                Else
                    GrossNum = GrossNum + 1
                End If
                
                If GrossNum = rsExpression.RecordCount Then
                    MsgBox "巨检描述模板对应的分类，数据库中不存在！"
                End If
            Else
                zlDatabase.SetPara "巨检描述模板", txtGrossDescribe.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
            End If
            
             If txtNormalReport.Text <> "" Then
                '如果用户输入的分类和数据库匹配 则将参数保存到数据库中
                If rsExpression("名称").value = txtNormalReport.Text Then
                    '执行保存参数
                    zlDatabase.SetPara "常规报告模板", txtNormalReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
                Else
                    NormalNum = NormalNum + 1
                End If
                
                If NormalNum = rsExpression.RecordCount Then
                    MsgBox "常规报告模板对应的分类，数据库中不存在！"
                End If
            Else
                zlDatabase.SetPara "常规报告模板", txtNormalReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
            End If
            
            If txtImmuneReport.Text <> "" Then
                '如果用户输入的分类和数据库匹配 则将参数保存到数据库中
                If rsExpression("名称").value = txtImmuneReport.Text Then
                    '执行保存参数
                    zlDatabase.SetPara "免疫报告模板", txtImmuneReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
                Else
                    ImmuneNum = ImmuneNum + 1
                End If
                
                If ImmuneNum = rsExpression.RecordCount Then
                    MsgBox "免疫报告模板对应的分类，数据库中不存在！"
                End If
            Else
                zlDatabase.SetPara "免疫报告模板", txtImmuneReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
            End If
            
            If txtSpecialStainReport.Text <> "" Then
                '如果用户输入的分类和数据库匹配 则将参数保存到数据库中
                If rsExpression("名称").value = txtSpecialStainReport.Text Then
                    '执行保存参数
                    zlDatabase.SetPara "特染报告模板", txtSpecialStainReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
                Else
                    SpecialNum = SpecialNum + 1
                End If
                
                If SpecialNum = rsExpression.RecordCount Then
                    MsgBox "特染报告模板对应的分类，数据库中不存在！"
                End If
             Else
                zlDatabase.SetPara "特染报告模板", txtSpecialStainReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
            End If
            
            If txtMoleculeReport.Text <> "" Then
                '如果用户输入的分类和数据库匹配 则将参数保存到数据库中
                If rsExpression("名称").value = txtMoleculeReport.Text Then
                    '执行保存参数
                    zlDatabase.SetPara "分子报告模板", txtMoleculeReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
                Else
                    MoleculeNum = MoleculeNum + 1
                End If
                
                If MoleculeNum = rsExpression.RecordCount Then
                    MsgBox "分子报告模板对应的分类，数据库中不存在！"
                End If
            Else
                zlDatabase.SetPara "分子报告模板", txtMoleculeReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
            End If
            
            If Not rsExpression.EOF Then
                rsExpression.MoveNext
            End If

        Next
    End If
    
    zlDatabase.SetPara "图像校对", chkValid.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0
    
    mblnOK = True
    Unload Me
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call CmdHelp_Click
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub InitFaceScheme()
    Dim Item As TabControlItem
    
    If mlngModul = 1290 And InStr(mstrPrivs, "三维重建设置") <> 0 Then
        cmd3DSetup.Visible = True
    Else
        cmd3DSetup.Visible = False
    End If
    
    
    '如果是病理系统，则不进行执行间的设置
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        chkReagent.Visible = False
        chkInputOutInfo.Visible = False     '病理系统不进行登记外院送检单位和送检医生
    Else
        frmPatholParameter.Visible = False
    End If
End Sub


Private Sub Form_Load()
    InitFaceScheme
    mblnOK = False
    Dim intTemp As Integer
    Dim strTemp As String
    Dim i As Integer
    Dim blnChkVisible As Boolean
    
    chkPasv.ForeColor = &HFF0000
'    labMoneyExemodel.Visible = IIf(mlngModul = G_LNG_VIDEOSTATION_MODULE Or mlngModul = G_LNG_PACSSTATION_MODULE, True, False)
'    cbxMoneyExeModle.Visible = labMoneyExemodel.Visible
    
        
    '根据不同的工作站加载不同的 主要工作页面 参数
    Select Case mlngModul

        Case 1290 '医技工作站
            cbxMainPage.Clear
            cbxMainPage.AddItem ("")
            cbxMainPage.AddItem ("影像")
            cbxMainPage.AddItem ("报告")
            cbxMainPage.AddItem ("费用")
            cbxMainPage.AddItem ("医嘱")
            cbxMainPage.AddItem ("病历")
            chkPasv.Left = optCheckInMode(2).Left
            chkPasv.Top = optCheckInMode(2).Top + 300
            
            cbxMoneyExeModle.Clear
            cbxMoneyExeModle.AddItem ("报到时执行")
            cbxMoneyExeModle.AddItem ("报告时执行")
            cbxMoneyExeModle.ListIndex = Val(zlDatabase.GetPara("医技费用执行模式", glngSys, mlngModul, 0, Array(cbxMoneyExeModle), InStr(mstrPrivs, ";参数设置;") > 0))
        Case 1291 '采集工作站
            cbxMainPage.Clear
            cbxMainPage.AddItem ("")
            cbxMainPage.AddItem ("采集")
            cbxMainPage.AddItem ("报告")
            cbxMainPage.AddItem ("费用")
            cbxMainPage.AddItem ("医嘱")
            cbxMainPage.AddItem ("病历")
            chkPasv.Left = optCheckInMode(2).Left
            chkPasv.Top = optCheckInMode(2).Top + 300
            
            cbxMoneyExeModle.ListIndex = Val(zlDatabase.GetPara("采集费用执行模式", glngSys, mlngModul, 0, Array(cbxMoneyExeModle), InStr(mstrPrivs, ";参数设置;") > 0))
        Case 1294 '病理工作站
            cbxMainPage.Clear
            cbxMainPage.AddItem ("")
            cbxMainPage.AddItem ("采集")
            cbxMainPage.AddItem ("核收")
            cbxMainPage.AddItem ("取材")
            cbxMainPage.AddItem ("制片")
            cbxMainPage.AddItem ("特殊")
            cbxMainPage.AddItem ("诊断")
            cbxMainPage.AddItem ("报告")
            cbxMainPage.AddItem ("费用")
            cbxMainPage.AddItem ("医嘱")
            cbxMainPage.AddItem ("病历")
            chkPasv.Left = optCheckInMode(1).Left + 1575
            chkPasv.Top = optCheckInMode(1).Top
            
            cbxMoneyExeModle.ListIndex = Val(zlDatabase.GetPara("病理费用执行模式", glngSys, mlngModul, 0, Array(cbxMoneyExeModle), InStr(mstrPrivs, ";参数设置;") > 0))
    End Select
    
    CmdDevSet.Enabled = InStr(mstrPrivs, ";参数设置;") > 0
    CmdOK.Enabled = InStr(mstrPrivs, ";参数设置;") > 0
    chkView.value = Val(zlDatabase.GetPara("报告时观片", glngSys, mlngModul, 0, Array(chkView), InStr(mstrPrivs, ";参数设置;") > 0))
    chkPasv.value = IIf(Val(GetSetting("ZLSOFT", "公共模块\ZL9PACSWork\frmTechnicSetup", "启用被动传输", 0)) = 1, 1, 0)
    chkBatchInput.value = Val(zlDatabase.GetPara("连续登记申请", glngSys, mlngModul, 0, Array(chkBatchInput), InStr(mstrPrivs, ";参数设置;") > 0))
    chkPatTrack.value = Val(zlDatabase.GetPara("病人跟踪", glngSys, mlngModul, 0, Array(chkPatTrack), InStr(mstrPrivs, ";参数设置;") > 0))
    ChkOpenReport.value = Val(zlDatabase.GetPara("开始检查自动打开报告", glngSys, mlngModul, 0, Array(ChkOpenReport), InStr(mstrPrivs, ";参数设置;") > 0))
    chkReagent.value = Val(zlDatabase.GetPara("不显示造影剂", glngSys, mlngModul, 0, Array(chkReagent), InStr(mstrPrivs, ";参数设置;") > 0))
    chkAddons.value = Val(zlDatabase.GetPara("不显示附加主述", glngSys, mlngModul, 0, Array(chkAddons), InStr(mstrPrivs, ";参数设置;") > 0))
    chkInputOutInfo.value = Val(zlDatabase.GetPara("录入外院信息", glngSys, mlngModul, 0, Array(chkInputOutInfo), InStr(mstrPrivs, ";参数设置;") > 0))
    intTemp = Val(zlDatabase.GetPara("登记模式", glngSys, mlngModul, 0, Array(optCheckInMode(1)), InStr(mstrPrivs, ";参数设置;") > 0))
    intTemp = Val(zlDatabase.GetPara("登记模式", glngSys, mlngModul, 0, Array(optCheckInMode(2)), InStr(mstrPrivs, ";参数设置;") > 0))
    chkFormat.value = Val(zlDatabase.GetPara("报到打印前选择格式", glngSys, mlngModul, 0, Array(chkFormat), InStr(mstrPrivs, ";参数设置;") > 0))
    If intTemp = 1 Then
        optCheckInMode(1).value = True
    Else
        optCheckInMode(2).value = True
    End If

    chkCancelCheck.value = Val(zlDatabase.GetPara("不显示被取消的登记", glngSys, mlngModul, 0, Array(chkCancelCheck), InStr(mstrPrivs, ";参数设置;") > 0))
    chkAutoPrint.value = Val(zlDatabase.GetPara("报到后自动打印申请单", glngSys, mlngModul, 0, Array(chkAutoPrint), InStr(mstrPrivs, ";参数设置;") > 0))
    chkExitAfterSign.value = Val(zlDatabase.GetPara("PACS报告签名后退出", glngSys, mlngModul, "1", Array(chkExitAfterSign), InStr(mstrPrivs, ";参数设置;") > 0))
    cbxMainPage.Text = zlDatabase.GetPara("工作首页", glngSys, mlngModul, "", Array(cbxMainPage), InStr(mstrPrivs, ";参数设置;") > 0)
    
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        chkDecalin.value = Val(zlDatabase.GetPara("脱钙声音提醒", glngSys, mlngModul, 1, Array(chkDecalin), InStr(mstrPrivs, ";参数设置;") > 0))
        txtDecalinHintTime.Text = Val(zlDatabase.GetPara("提醒间隔时长", glngSys, mlngModul, "30", Array(txtDecalinHintTime), InStr(mstrPrivs, ";参数设置;") > 0))
        chkChangGui.value = Val(zlDatabase.GetPara("常规质量窗口", glngSys, mlngModul, 1, Array(chkChangGui), InStr(mstrPrivs, ";参数设置;") > 0))
        chkKuaiPian.value = Val(zlDatabase.GetPara("快速石蜡质量窗口", glngSys, mlngModul, 1, Array(chkKuaiPian), InStr(mstrPrivs, ";参数设置;") > 0))
        chkBingDong.value = Val(zlDatabase.GetPara("冰冻质量窗口", glngSys, mlngModul, 1, Array(chkBingDong), InStr(mstrPrivs, ";参数设置;") > 0))
        chkXiBao.value = Val(zlDatabase.GetPara("细胞质量窗口", glngSys, mlngModul, 1, Array(chkXiBao), InStr(mstrPrivs, ";参数设置;") > 0))
        chkHuiZhen.value = Val(zlDatabase.GetPara("会诊质量窗口", glngSys, mlngModul, 1, Array(chkHuiZhen), InStr(mstrPrivs, ";参数设置;") > 0))
        chkShiJian.value = Val(zlDatabase.GetPara("尸检质量窗口", glngSys, mlngModul, 1, Array(chkShiJian), InStr(mstrPrivs, ";参数设置;") > 0))
        '读取是否直接打印参数信息
        chkIsDirectPrint.value = Val(zlDatabase.GetPara("是否直接打印", glngSys, mlngModul, 0, Array(chkIsDirectPrint), InStr(mstrPrivs, ";参数设置;") > 0))
        '读取模板对应词句分类参数
        txtGrossDescribe.Text = zlDatabase.GetPara("巨检描述模板", glngSys, mlngModul, "", Array(txtGrossDescribe), InStr(mstrPrivs, ";参数设置;") > 0)
        txtNormalReport.Text = zlDatabase.GetPara("常规报告模板", glngSys, mlngModul, "", Array(txtNormalReport), InStr(mstrPrivs, ";参数设置;") > 0)
        txtImmuneReport.Text = zlDatabase.GetPara("免疫报告模板", glngSys, mlngModul, "", Array(txtImmuneReport), InStr(mstrPrivs, ";参数设置;") > 0)
        txtSpecialStainReport.Text = zlDatabase.GetPara("特染报告模板", glngSys, mlngModul, "", Array(txtSpecialStainReport), InStr(mstrPrivs, ";参数设置;") > 0)
        txtMoleculeReport.Text = zlDatabase.GetPara("分子报告模板", glngSys, mlngModul, "", Array(txtMoleculeReport), InStr(mstrPrivs, ";参数设置;") > 0)

    End If

    chkValid.value = Val(zlDatabase.GetPara("图像校对", glngSys, mlngModul, "", , InStr(mstrPrivs, ";参数设置;") > 0))
    
    Call ResizeFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng科室ID = 0
End Sub

Private Sub TxtLike_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtShowPhotoNumber_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub Txt默认天数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub optCheckInMode_Click(Index As Integer)
    If optCheckInMode(1).value = True Then
        chkReagent.Enabled = False
        chkAddons.Enabled = False
    Else
        chkReagent.Enabled = True
        chkAddons.Enabled = True
    End If
End Sub

Private Sub ResizeFace()
    On Error Resume Next
    
    If mlngModul <> G_LNG_PATHOLSYS_NUM Then
        chkAutoPrint.Left = 240
        chkAutoPrint.Top = 120
        
        chkFormat.Top = chkAutoPrint.Top
        
        chkCancelCheck.Left = 240
        chkCancelCheck.Top = chkAutoPrint.Top + chkAutoPrint.Height + 120
        
        chkView.Left = chkView.Left
        chkView.Top = chkCancelCheck.Top
        
        chkPatTrack.Left = 240
        chkPatTrack.Top = chkCancelCheck.Top + chkCancelCheck.Height + 120
        
        ChkOpenReport.Left = ChkOpenReport.Left
        ChkOpenReport.Top = chkPatTrack.Top
        
        chkExitAfterSign.Left = 240
        chkExitAfterSign.Top = chkPatTrack.Top + chkPatTrack.Height + 120
        
        chkBatchInput.Top = chkExitAfterSign.Top
        
        Load Frame6(1)
        With Frame6(1)
            .Left = Frame6(0).Left
            .Top = chkExitAfterSign.Top + chkExitAfterSign.Height + 120
            .Width = Frame6(0).Width
            .Height = 25
            
            .Caption = ""
            .Visible = True
        End With
        
        Frame2.Top = Frame6(1).Top + Frame6(1).Height + 200
        Frame2.Height = 1185

        Label2.Top = Frame6(1).Top + Frame6(1).Height + 150
        
        cbxMainPage.Top = Label2.Top + Label2.Height + 100
        
        labMoneyExemodel.Left = cbxMainPage.Left
        labMoneyExemodel.Top = cbxMainPage.Top + cbxMainPage.Height + 150
        
        cbxMoneyExeModle.Left = labMoneyExemodel.Left
        cbxMoneyExeModle.Top = labMoneyExemodel.Top + labMoneyExemodel.Height
        cbxMoneyExeModle.Width = cbxMainPage.Width
        
        fraValid.Top = Frame2.Top + Frame2.Height + 120
        fraValid.Left = Frame2.Left
        
        Frame6(0).Top = fraValid.Top + fraValid.Height + 150
        
        cmdHelp.Top = Frame6(0).Top + Frame6(0).Height + 150
        CmdDevSet.Top = cmdHelp.Top
        cmd3DSetup.Top = cmdHelp.Top
        CmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        PicAction.Height = cmdCancel.Top + cmdCancel.Height + 150
        
        Me.Height = PicAction.Top + PicAction.Height + 600
    End If
End Sub
