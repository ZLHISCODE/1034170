VERSION 5.00
Begin VB.Form frmStuffPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmStuffPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "3、卫材分批属性自动设置"
      Height          =   1335
      Left            =   180
      TabIndex        =   19
      Top             =   3120
      Width           =   4620
      Begin VB.OptionButton optsetall 
         Caption         =   "库房和发料部门分批"
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1980
      End
      Begin VB.OptionButton optSet库房 
         Caption         =   "仅库房分批"
         Height          =   210
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   1500
      End
      Begin VB.OptionButton optSetNotall 
         Caption         =   "库房和发料部门都不分批"
         Height          =   210
         Left            =   2160
         TabIndex        =   21
         Top             =   840
         Width           =   2340
      End
      Begin VB.OptionButton optSet手动 
         Caption         =   "手工设置分批属性"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1740
      End
   End
   Begin VB.Frame fra 
      Caption         =   "3、允许“应用于…”的范围"
      Height          =   4185
      Left            =   4890
      TabIndex        =   14
      Top             =   270
      Width           =   3120
      Begin VB.CheckBox chk存储库房 
         Caption         =   "应用于分类下所有卫生材料(&N)"
         Height          =   324
         Index           =   2
         Left            =   144
         TabIndex        =   6
         Top             =   840
         Width           =   2760
      End
      Begin VB.CheckBox chk存储库房 
         Caption         =   "应用于本级所有卫生材料(&B)"
         Height          =   324
         Index           =   1
         Left            =   144
         TabIndex        =   5
         Top             =   540
         Width           =   2712
      End
      Begin VB.CheckBox chk存储库房 
         Caption         =   "应用于所有卫生材料(&A)"
         Height          =   324
         Index           =   0
         Left            =   144
         TabIndex        =   4
         Top             =   285
         Width           =   2364
      End
      Begin VB.Label lblInfor 
         Caption         =   "   如:没有勾上此栏目中的“应用于所有卫生材料”，则在存储库房设置界面中的『应用于所有“卫生材料”(4)』将不能选择！"
         ForeColor       =   &H00000000&
         Height          =   870
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   2820
      End
      Begin VB.Label lblInfor 
         Caption         =   "    本栏目主要是控制卫生材料管理的存储库房设置界面中的“应用于...”功能。"
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   0
         Left            =   165
         TabIndex        =   16
         Top             =   1800
         Width           =   2910
      End
      Begin VB.Label lblInfor 
         Caption         =   "说明:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   15
         Top             =   1440
         Width           =   2550
      End
   End
   Begin VB.Frame fraCodeMode 
      Height          =   1515
      Left            =   180
      TabIndex        =   10
      Top             =   360
      Width           =   4620
      Begin VB.OptionButton opt编码模式 
         Caption         =   "&3) 分类号+顺序编号"
         Height          =   210
         Index           =   2
         Left            =   900
         TabIndex        =   2
         Top             =   1155
         Width           =   3420
      End
      Begin VB.OptionButton opt编码模式 
         Caption         =   "&2) 材料类别+分类号+顺序编号"
         Height          =   210
         Index           =   1
         Left            =   900
         TabIndex        =   1
         Top             =   825
         Width           =   3420
      End
      Begin VB.OptionButton opt编码模式 
         Caption         =   "&1) 同类顺序编号"
         Height          =   210
         Index           =   0
         Left            =   900
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.Label lbl编码模式 
         Caption         =   "1、编码缺省递增模式"
         Height          =   180
         Left            =   600
         TabIndex        =   18
         Top             =   0
         Width           =   1750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   60
         Picture         =   "frmStuffPara.frx":000C
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   9
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5625
      TabIndex        =   7
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6870
      TabIndex        =   8
      Top             =   4725
      Width           =   1100
   End
   Begin VB.Frame fraIncome 
      Height          =   855
      Left            =   180
      TabIndex        =   12
      Top             =   2040
      Width           =   4620
      Begin VB.ComboBox cbo收入项目 
         ForeColor       =   &H80000012&
         Height          =   300
         Index           =   0
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label lblIncome 
         AutoSize        =   -1  'True
         Caption         =   "2、各材质对应缺省收入项目"
         Height          =   180
         Left            =   585
         TabIndex        =   11
         Top             =   0
         Width           =   2250
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   60
         Picture         =   "frmStuffPara.frx":08D6
         Top             =   60
         Width           =   480
      End
      Begin VB.Label LblNote 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   13
         Top             =   390
         Visible         =   0   'False
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmStuffPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnActive As Boolean
Private mintTabIndex As Integer
Private mlng收入项目ID As Long
Private mstrPrivs As String
Private mrs收入项目 As New ADODB.Recordset
Private mblnHavePriv As Boolean
Private Const mlngModule = 1711

Public Sub ShowMe(ByVal strPrivs As String, ByVal frmMain As Object)
    '----------------------------------------------------------------------------------
    '功能:参数设置入口
    '参数:mstrPrivs -权限串
    '     frmMain-调用父窗口
    '返回:
    '编制:刘兴宏
    '日期:2007/12/24
    '----------------------------------------------------------------------------------
    mstrPrivs = strPrivs
    Me.Show 1, frmMain
End Sub

 
Private Sub cbo收入项目_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'Private Sub chk规格连续_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'End Sub

'Private Sub chk品种规格_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'End Sub

 
Private Sub chk存储库房_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
            zlCommFun.PressKey vbKeyTab
        End If
End Sub

'Private Sub chk品种连续_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'End Sub

Private Sub CmdCancel_Click()
    gblnIncomeItem = False
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub CmdOK_Click()
    Dim intSave As Integer
    Dim strReg As String
    
    If SaveSet = False Then Exit Sub
    gblnIncomeItem = True
    Unload Me
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/24
    '------------------------------------------------------------------------------------------
    Dim int编码模式 As Integer, str应用范围 As String
    Dim intSet分批 As Integer
    
    '格式:3位字符构成,1,代表允许，0代表不允许,如111.其中第一位代表所有,第二位代表本级所有,第三位代表分类下所有
    str应用范围 = IIf(chk存储库房(0).Value = 1, "1", "0")
    str应用范围 = str应用范围 & IIf(chk存储库房(1).Value = 1, "1", "0")
    str应用范围 = str应用范围 & IIf(chk存储库房(2).Value = 1, "1", "0")
    If Me.opt编码模式(0).Value = True Then
       int编码模式 = 0
    ElseIf Me.opt编码模式(1).Value = True Then
       int编码模式 = 1
    Else
       int编码模式 = 2
    End If
    
    If optSet手动.Value = True Then
        intSet分批 = 0
    ElseIf optSet库房.Value = True Then
        intSet分批 = 1
    ElseIf optsetall.Value = True Then
        intSet分批 = 2
    ElseIf optSetNotall.Value = True Then
        intSet分批 = 3
    End If
    
    err = 0: On Error GoTo ErrHand:
    
    gcnOracle.BeginTrans
'    Call zlDatabase.SetPara("品种增加模式", Me.chk品种连续.Value, glngSys, mlngModule)
'    Call zlDatabase.SetPara("品种规格模式", Me.chk品种规格.Value, glngSys, mlngModule)
'    Call zlDatabase.SetPara("规格增加模式", Me.chk规格连续.Value, glngSys, mlngModule)
    Call zlDatabase.SetPara("编码递增模式", int编码模式, glngSys, mlngModule)
    Call zlDatabase.SetPara("允许应用于的范围", str应用范围, glngSys, mlngModule)
    Call zlDatabase.SetPara("收入项目对应", cbo收入项目(1).ItemData(cbo收入项目(1).ListIndex), glngSys, mlngModule)
    Call zlDatabase.SetPara("卫材分批属性自动设置", intSet分批, glngSys, mlngModule)
    
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub Form_Activate()
    If Not mblnActive Then Unload Me: Exit Sub
    
End Sub

Private Sub Form_Load()
    '根据用户权限，装入控件
    Dim intValue As Integer
    Dim strReg As String
    Dim intSet分批 As Integer
    
    mblnHavePriv = IsHavePrivs(mstrPrivs, "参数设置")
    
    On Error GoTo errHandle
    mblnActive = False
    
'    chk品种连续.Value = IIf(Val(zlDatabase.GetPara("品种增加模式", glngSys, mlngModule, , Array(chk品种连续), mblnHavePriv)) = 1, 1, 0)
'    chk品种规格.Value = IIf(Val(zlDatabase.GetPara("品种规格模式", glngSys, mlngModule, , Array(chk品种规格), mblnHavePriv)) = 1, 1, 0)
'    chk规格连续.Value = IIf(Val(zlDatabase.GetPara("规格增加模式", glngSys, mlngModule, , Array(chk规格连续), mblnHavePriv)) = 1, 1, 0)
    
    
    intValue = Val(zlDatabase.GetPara("编码递增模式", glngSys, mlngModule, , Array(opt编码模式(0), opt编码模式(1), opt编码模式(2), lbl编码模式, fraCodeMode), mblnHavePriv))
    If intValue = 0 Then
        Me.opt编码模式(0).Value = True: Me.opt编码模式(1).Value = False: Me.opt编码模式(2).Value = False
    ElseIf intValue = 1 Then
        Me.opt编码模式(0).Value = False: Me.opt编码模式(1).Value = True: Me.opt编码模式(2).Value = False
    Else
        Me.opt编码模式(0).Value = False: Me.opt编码模式(1).Value = False: Me.opt编码模式(2).Value = True
    End If
    '格式:3位字符构成,1,代表允许，0代表不允许,如111.其中第一位代表所有,第二位代表本级所有,第三位代表分类下所有
    strReg = zlDatabase.GetPara("允许应用于的范围", glngSys, mlngModule, , Array(fra, chk存储库房(0), chk存储库房(1), chk存储库房(2)), mblnHavePriv)
        
    If Len(strReg) < 3 Then
        '默认全选中
        strReg = "111"
    End If
    chk存储库房(0).Value = IIf(Val(Mid(strReg, 1, 1)) = 1, 1, 0)
    chk存储库房(1).Value = IIf(Val(Mid(strReg, 2, 1)) = 1, 1, 0)
    chk存储库房(2).Value = IIf(Val(Mid(strReg, 3, 1)) = 1, 1, 0)
    
    
    intSet分批 = Val(zlDatabase.GetPara("卫材分批属性自动设置", glngSys, mlngModule, 0))
    Select Case intSet分批
        Case 0
            optSet手动.Value = True
        Case 1
            optSet库房.Value = True
        Case 2
            optsetall.Value = True
        Case 3
            optSetNotall.Value = True
    End Select
    
    gstrSQL = "Select ID,编码||'-'||名称 名称 From 收入项目 Where 末级=1"
    zlDatabase.OpenRecordset mrs收入项目, gstrSQL, Me.Caption
    With mrs收入项目
        If .EOF Then
            MsgBox "请初始化收入项目（收入项目）！", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    mlng收入项目ID = Val(zlDatabase.GetPara("收入项目对应", glngSys, mlngModule, "", Array(lblIncome, fraIncome, cbo收入项目, LblNote(0)), mblnHavePriv))
    mintTabIndex = 10

    Call AddCons("卫生材料")
    mblnActive = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddCons(ByVal strName As String)
    Dim intIdx As Integer
    intIdx = LblNote.UBound + 1
    Load LblNote(intIdx)
    Load cbo收入项目(intIdx)
    
    mintTabIndex = mintTabIndex + 1
    With LblNote(intIdx)
        .Caption = strName
        .TabIndex = mintTabIndex
        .Container = fraIncome
        .Top = IIf(intIdx = 1, LblNote(0).Top, LblNote(intIdx - 1).Top) + IIf(intIdx = 1, 0, LblNote(0).Height + 200)
        .Left = LblNote(0).Left + LblNote(0).Width - .Width
        .Visible = True
    End With
    mintTabIndex = mintTabIndex + 1
    With cbo收入项目(intIdx)
        .Container = fraIncome
        .Left = cbo收入项目(0).Left
        .Top = IIf(intIdx = 1, cbo收入项目(0).Top, cbo收入项目(intIdx - 1).Top) + IIf(intIdx = 1, 0, cbo收入项目(0).Height + 100)
        .TabIndex = mintTabIndex
        .Visible = True
    End With
    Call AddItem(cbo收入项目(intIdx), strName)
End Sub

Private Sub AddItem(ByVal cboObj As ComboBox, ByVal strName As String)
    Dim i As Integer

    With mrs收入项目
        .MoveFirst
        Do While Not .EOF
            cboObj.AddItem !名称
            cboObj.ItemData(cboObj.NewIndex) = !Id
            .MoveNext
        Loop
        For i = 0 To cboObj.ListCount - 1
            If cboObj.ItemData(i) = mlng收入项目ID Then
                cboObj.ListIndex = i
                Exit Sub
            End If
        Next
        For i = 0 To cboObj.ListCount - 1
            If strName = "卫生材料" Then    '如果是卫生材料那就先找材料费的，没有则找含有材料的，两者都没有的这默认选中第一个
                If cboObj.List(i) Like "*材料费" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
        For i = 0 To cboObj.ListCount - 1
            If strName = "卫生材料" Then
                If cboObj.List(i) Like "*材料*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
        cboObj.ListIndex = 0
    End With
End Sub

Private Sub opt编码模式_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
