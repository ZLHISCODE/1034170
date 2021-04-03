VERSION 5.00
Begin VB.Form frmSet涪陵 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmSet涪陵.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdTrans 
      Caption         =   "上传"
      Height          =   350
      Left            =   90
      TabIndex        =   11
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CheckBox chk床位 
      Caption         =   "上传床位信息"
      Height          =   210
      Left            =   420
      TabIndex        =   10
      Top             =   1860
      Width           =   3375
   End
   Begin VB.CheckBox chk诊疗 
      Caption         =   "上传诊疗项目信息"
      Height          =   210
      Left            =   420
      TabIndex        =   9
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CheckBox chk药品 
      Caption         =   "上传药品编码信息"
      Height          =   210
      Left            =   420
      TabIndex        =   8
      Top             =   1260
      Width           =   3375
   End
   Begin VB.CheckBox chk疾病 
      Caption         =   "上传疾病编码信息"
      Height          =   210
      Left            =   420
      TabIndex        =   7
      Top             =   960
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   750
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3810
      TabIndex        =   5
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2670
      TabIndex        =   4
      Top             =   2400
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2250
      Width           =   5265
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1545
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "号串口"
      Height          =   180
      Index           =   4
      Left            =   1950
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "当前串口(&D)"
      Height          =   180
      Index           =   3
      Left            =   450
      TabIndex        =   1
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmSet涪陵"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng险类 As Long
 
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtEdit) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call ExecuteProcedure(Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'端口号','" & txtEdit.Text & "',1)"
    Call ExecuteProcedure(Me.Caption)
    
    gcnOracle.CommitTrans
    gintComPort = txtEdit.Text
    mblnReturn = True
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTrans_Click()
    Dim rsTemp As New ADODB.Recordset, iLoop As Long, strTemp As String
'    gstr医保机构编码 = "500102"
'    gstr医院编码 = "5001020003"
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Sub
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    If chk疾病.Value = 1 Then
        gstrSQL = "Select id as 编码,名称 From 保险病种"
        Call OpenRecordset(rsTemp, gstrSysName)
        chk疾病.Caption = "上传疾病编码信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstr医保机构编码, gstr医院编码, "0", rsTemp!编码, rsTemp!名称, "", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk疾病.Caption = "上传疾病编码信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk疾病.Value = 0
    End If
    If chk药品.Value = 1 Then
        gstrSQL = "select a.类别 as 类别,a.id as 编码,a.名称 as 名称,b.药品来源 as 药品来源 from 收费细目 a,药品目录 b where a.类别 In ('5','6','7') and a.编码=b.编码"
        Call OpenRecordset(rsTemp, gstrSysName)
        chk药品.Caption = "上传药品编码信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstr医保机构编码, gstr医院编码, "1", rsTemp!类别 & "_" & rsTemp!编码, rsTemp!名称, IIf(rsTemp!药品来源 = "进口", "03", "02"), gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk药品.Caption = "上传药品编码信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk药品.Value = 0
    End If
    If chk诊疗.Value = 1 Then
        gstrSQL = "select * from 收费细目 where 类别 Not In ('J','5','6','7')"
        Call OpenRecordset(rsTemp, gstrSysName)
        chk诊疗.Caption = "上传诊疗项目信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstr医保机构编码, gstr医院编码, "2", rsTemp!类别 & "_" & rsTemp!ID, rsTemp!名称, "", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk诊疗.Caption = "上传诊疗项目信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk诊疗.Value = 0
    End If
    If chk床位.Value = 1 Then
        gstrSQL = "select * from 收费细目 where 类别='J'"
        Call OpenRecordset(rsTemp, gstrSysName)
        chk床位.Caption = "上传床位信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = wyyglxx(gstr医保机构编码, gstr医院编码, "3", rsTemp!类别 & "_" & rsTemp!ID, rsTemp!名称, " ", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk床位.Caption = "上传床位信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk床位.Value = 0
    End If
    MsgBox "基本项目信息上传完成", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnReturn = False
    
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=" & mlng险类
    Call OpenRecordset(rsTemp, "读取参数")
    
    If Not rsTemp.EOF Then txtEdit.Text = rsTemp!参数值
End Sub

Public Function ShowME(ByVal lng险类 As Long) As Boolean
    mlng险类 = lng险类
    Me.Show 1
    ShowME = mblnReturn
End Function
