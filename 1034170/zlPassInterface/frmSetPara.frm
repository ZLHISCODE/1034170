VERSION 5.00
Begin VB.Form frmSetPara 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   Icon            =   "frmSetPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9780
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "kj"
      Height          =   5415
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   8
         Left            =   1320
         TabIndex        =   19
         Top             =   4800
         Width           =   8175
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   2
         Left            =   1320
         TabIndex        =   18
         ToolTipText     =   "示例:http://serverAddr/v4/engineIncrementAsync"
         Top             =   2400
         Width           =   8175
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   3
         Left            =   1320
         TabIndex        =   17
         ToolTipText     =   "示例：http://serverAddr/v4/invalidIncrement"
         Top             =   3000
         Width           =   8175
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   10
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "示例:http://127.0.0.1:80/zlcx/data_detail.action"
         Top             =   3600
         Width           =   8175
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   9
         Left            =   1320
         TabIndex        =   15
         ToolTipText     =   "示例:http://serverAddr/valid"
         Top             =   4200
         Width           =   8175
      End
      Begin VB.Frame fraContent 
         BackColor       =   &H80000005&
         Caption         =   "CONTENT-TYPE"
         Height          =   1215
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   5055
         Begin VB.OptionButton optContent 
            BackColor       =   &H80000005&
            Caption         =   "application/xml;charset=utf-8"
            Height          =   375
            Index           =   1
            Left            =   360
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   14
            Top             =   360
            Width           =   3015
         End
         Begin VB.OptionButton optContent 
            BackColor       =   &H80000005&
            Caption         =   "application/x-www-form-urlencoded;charset=utf-8"
            Height          =   375
            Index           =   0
            Left            =   360
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   13
            Top             =   720
            Width           =   4575
         End
      End
      Begin VB.Frame fraOpt 
         BackColor       =   &H80000005&
         Caption         =   "药品说明书"
         Height          =   1215
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Width           =   3975
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "产品非公用"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   11
            Top             =   720
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "产品公用"
            Height          =   255
            Index           =   0
            Left            =   360
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   10
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   1
         Left            =   1320
         MaxLength       =   32
         TabIndex        =   8
         Top             =   1800
         Width           =   8175
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "患教地址"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "用药审查端口号"
         Top             =   4890
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "干预地址"
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   24
         Top             =   2490
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "删除地址"
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "不含参数部分。示例:http://serverAddr/v4/invalidIncrement"
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "说明书地址"
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "药品说明书端口号"
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "有效处方地址"
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "http://serverAddr/valid"
         Top             =   4290
         Width           =   1080
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "医院编码"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1890
         Width           =   720
      End
   End
   Begin VB.TextBox txtIn 
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   4
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9780
      TabIndex        =   0
      Top             =   5640
      Width           =   9780
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chk 
         BackColor       =   &H80000005&
         Caption         =   "启用药师审方干预系统"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "医院编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmSetPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CMD_ENUM
    CMD_OK = 0
    CMD_CANCEL = 1
End Enum

Private Sub cmdPara_Click(Index As Integer)
    Dim blnOK As Boolean
    Dim strPara As String
    Dim strSQL As String
    Dim lngID As Long
    Dim rsTmp As ADODB.Recordset
    
    If Index = CMD_OK Then
        If gbytPass = DT And gstrVersion = "4.0" Then
            strPara = Trim(txtIn.Text)
        ElseIf gbytPass = MK And gstrVersion = "4.0" Then
            strPara = IIf(chk(0).Value = vbChecked, 1, 0)
        ElseIf gbytPass = HZYY Then
            Call HZYY_SetPara
        End If
        On Error GoTo errH
        If strPara <> "" Then
            strSQL = "Select count(1) as RowCount  From zlParameters Where 系统 = [1] And Nvl(模块, 0) = 0 And 参数号 = 90001"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "合理用药监测配置", glngSys)
            If Not rsTmp.EOF Then
                If rsTmp!RowCount = 0 Then
                    lngID = zlDatabase.GetNextId("zlParameters")
                    strSQL = "Insert Into zlParameters(ID, 系统, 模块, 参数号, 参数名, 参数值) Values (" & lngID & ", " & glngSys & ", Null, 90001, '合理用药监测配置','" & strPara & "')"
                    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
                    blnOK = True
                End If
            End If
            If Not blnOK Then
                Call zlDatabase.SetPara(90001, strPara, glngSys)
            End If
        End If
    End If
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strPara As String
    
    Me.Width = 4215
    Me.Height = 2535
    fra(0).Visible = False: fra(1).Visible = False
    txtIn.Visible = False
    lblInfo.Visible = False
    If gbytPass = DT And gstrVersion = "4.0" Then
        txtIn.Visible = True
        lblInfo.Visible = True
        strPara = zlDatabase.GetPara(90001, glngSys, , "1513")
        txtIn.Text = strPara
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        fra(0).Visible = True
        strPara = zlDatabase.GetPara(90001, glngSys, , "0")
        chk(0).Value = IIf(Val(strPara) = 1, vbChecked, vbUnchecked)
    ElseIf gbytPass = HZYY Then
        Me.Height = 6540
        Me.Width = 9990
        fra(1).Visible = True
        Call HZYY_GetPara
        txtPara(1).Text = gstrHOSCODE
        txtPara(2).Text = gstrUrlCheck
        txtPara(3).Text = gstrUrlDel
        txtPara(8).Text = gstrEduURL
        txtPara(9).Text = gstrUrlUpLoad
        txtPara(10).Text = gstrUrlDrug
        optType(0).Value = (gbytType = 0)
        optType(1).Value = (gbytType = 1)
        optContent(0).Value = (gbytContentType = 0)
        optContent(1).Value = (gbytContentType = 1)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If gbytPass = MK And gstrVersion = "4.0" Then
        fra(0).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    ElseIf gbytPass = HZYY Then
        fra(1).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    End If
    cmdPara(CMD_CANCEL).Left = picBottom.Width - 1100 - 120
    cmdPara(CMD_OK).Left = cmdPara(CMD_CANCEL).Left - 1100 - 60
End Sub

Private Sub optContent_Click(Index As Integer)
    If optContent(Index).Value Then
        gbytContentType = Index
    End If
End Sub

Private Sub optType_Click(Index As Integer)
    gbytType = Index
End Sub

Private Sub txtPara_Change(Index As Integer)
    If gbytPass = HZYY Then
        If Index = 1 Then
            gstrHOSCODE = txtPara(Index)
        ElseIf Index = 2 Then
            gstrUrlCheck = txtPara(Index)
        ElseIf Index = 3 Then
            gstrUrlDel = txtPara(Index)
        ElseIf Index = 8 Then
            gstrEduURL = txtPara(Index)
        ElseIf Index = 9 Then
            gstrUrlUpLoad = txtPara(Index)
        ElseIf Index = 10 Then
            gstrUrlDrug = txtPara(Index)
        End If
    End If
End Sub

