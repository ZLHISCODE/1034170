VERSION 5.00
Begin VB.Form frmIdentify成都内江 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo类别 
      Height          =   300
      ItemData        =   "frmIdentify成都内江.frx":0000
      Left            =   915
      List            =   "frmIdentify成都内江.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1290
      Width           =   2295
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   6
      Top             =   4710
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4710
      TabIndex        =   5
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   4
      Top             =   705
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -465
      TabIndex        =   3
      Top             =   4500
      Width           =   8340
   End
   Begin VB.TextBox TxtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   915
      MaxLength       =   20
      TabIndex        =   2
      Top             =   915
      Width           =   2295
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4605
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   915
      Width           =   2385
   End
   Begin VB.CommandButton cmd修改密码 
      Caption         =   "修改密码"
      Height          =   350
      Left            =   300
      TabIndex        =   0
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   915
      TabIndex        =   41
      Top             =   4050
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "制卡单位"
      Height          =   180
      Index           =   17
      Left            =   180
      TabIndex        =   40
      Top             =   4095
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   4605
      TabIndex        =   39
      Top             =   4050
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "在职情况"
      Height          =   180
      Index           =   16
      Left            =   3825
      TabIndex        =   38
      Top             =   4095
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   4605
      TabIndex        =   37
      Top             =   3645
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   15
      Left            =   3825
      TabIndex        =   36
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "请在密码栏输入病人的IC卡密码后回车,将读取病人的相关信息。"
      Height          =   180
      Left            =   720
      TabIndex        =   35
      Top             =   465
      Width           =   5130
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmIdentify成都内江.frx":0004
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保卡号"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   34
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      Height          =   180
      Index           =   1
      Left            =   3825
      TabIndex        =   33
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   540
      TabIndex        =   32
      Top             =   1762
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   4185
      TabIndex        =   31
      Top             =   1755
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   30
      Top             =   2145
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出身日期"
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   29
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "工况类别"
      Height          =   180
      Index           =   6
      Left            =   3825
      TabIndex        =   28
      Top             =   2145
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡有效期"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   27
      Top             =   3300
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   540
      TabIndex        =   26
      Top             =   2925
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "统筹编号"
      Height          =   180
      Index           =   9
      Left            =   3825
      TabIndex        =   25
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "补卡次数"
      Height          =   180
      Index           =   10
      Left            =   3825
      TabIndex        =   24
      Top             =   3300
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "制卡日期"
      Height          =   180
      Index           =   11
      Left            =   3825
      TabIndex        =   23
      Top             =   2925
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位号码"
      Height          =   180
      Index           =   12
      Left            =   180
      TabIndex        =   22
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   21
      Top             =   1305
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   915
      TabIndex        =   20
      Top             =   1710
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4605
      TabIndex        =   19
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   915
      TabIndex        =   18
      Top             =   2100
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4605
      TabIndex        =   17
      Top             =   2100
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   915
      TabIndex        =   16
      Top             =   2490
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4605
      TabIndex        =   15
      Top             =   2490
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   915
      TabIndex        =   14
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4605
      TabIndex        =   13
      Top             =   2880
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   915
      TabIndex        =   12
      Top             =   3255
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   4605
      TabIndex        =   11
      Top             =   3255
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   915
      TabIndex        =   10
      Top             =   3645
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   13
      Left            =   4185
      TabIndex        =   9
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "交易类别"
      Height          =   180
      Index           =   14
      Left            =   180
      TabIndex        =   8
      Top             =   1350
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify成都内江"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
Private mlng病人ID As Long
Private mstrReturn As String
Private mblnChange As Boolean
Private mblnFirst As Boolean

Private Sub cbo类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd修改密码_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    Dim strInPut As String, strOutPut As String
    
    If InitInfor_成都内江.读卡器_内江 = 0 Then
        '明华修改密码
    
        strNewPassWord = frm修改密码.ChangePassword(strOldPassWord, strOldPassWord)
        
        If strOldPassWord = strNewPassWord Then Exit Sub
        If strNewPassWord = "" Then Exit Sub
        '    a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
        '    b)  OldPassword：输入参数，为原密码，要求长度为6，字符串中只能包含0到9的数字；
        '    c)  NewPassword：输入参数，为新密码，要求长度为6，字符串中只能包含0到9的数字。
        strInPut = InitInfor_成都内江.串号号_内江
        strInPut = strInPut & vbTab & strOldPassWord
        strInPut = strInPut & vbTab & strNewPassWord
        
        If 业务请求_成都内江(更改密码_内江, strInPut, strOutPut) = False Then Exit Sub
        txtEdit(1).Text = strNewPassWord
    Else
        '其他只能读卡
    End If
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '如果不是明华的读卡，则不需输入密码
    If InitInfor_成都内江.读卡器_内江 = 0 Then Exit Sub
    txtEdit(1).Enabled = False
    txtEdit(1).BackColor = txtEdit(0).BackColor
    
    '读卡
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
    Me.cmd修改密码.Caption = "读卡(&R)"
    
    
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        txtEdit(Index).Tag = ""
        g病人身份_成都内江.个人编号 = ""
        g病人身份_成都内江.卡号 = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    mblnChange = True
    
    If Index = 1 Then
        '密码输入完毕
        '需获取病人信息
         SetOKCtrl False
         If ReadCardInFo = False Then Exit Sub
        '初始值
        Call LoadCtrlData
        SetOKCtrl True
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
Private Function ReadCardInFo() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取读卡信息
    '--入参数:
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInPut As String
     '读取病人信息
        '   a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
        '   b)  UserPassword：输入参数，为用户密码，要求长度为6，字符串中只能包含0到9的数字；
           
    ReadCardInFo = False
    strInPut = InitInfor_成都内江.串号号_内江
    If InitInfor_成都内江.读卡器_内江 = 0 Then
        '明华需输入密码
        If Trim(txtEdit(1)) = "" Then
            ShowMsgbox "请输入IC卡密码!"
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        End If
        strInPut = strInPut & vbTab & txtEdit(1).Text
    End If
    
    Err = 0
    On Error GoTo ErrHand:
    
    If 获取参保人员信息_成都内江(strInPut) = False Then Exit Function
    ReadCardInFo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmd确定.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "还没有进行身份验证！", vbInformation, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    If Trim(g病人身份_成都内江.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    If cbo类别.Text = "" Then
        ShowMsgbox "交易类别未选择"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查录前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=" & TYPE_成都内江 & " and 医保号='" & g病人身份_成都内江.个人编号 & "'"
            Call OpenRecordset(rsTemp, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '设置
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim lng疾病ID As Long
    
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str类别 As String
    Dim int当前状态 As Integer
    
    
    If IsValid = False Then Exit Sub
    
    
    g病人身份_成都内江.交易类别 = Split(cbo类别.Text, "-")(0)
    int当前状态 = 0
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & gintInsure & " and  医保号='" & g病人身份_成都内江.个人编号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证(工况类别);7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号(统筹地区编码|制卡日期|卡有效日期);16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_成都内江
        
        strIdentify = .卡号                               '0卡号
        strIdentify = strIdentify & ";" & .个人编号           '1医保号
        strIdentify = strIdentify & ";"                     '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", .性别)              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & .身份证号           '6身份证
        strIdentify = strIdentify & ";" & IIf(.单位号码 = "", "", "(" & .单位号码 & ")")            '7.单位名称(编码)
        strAddition = ";0"                                          '8.中心代码
        strAddition = strAddition & ";"                             '9.顺序号
        strAddition = strAddition & ";" & .交易类别                 '10人员身份
        strAddition = strAddition & ";" & .帐户余额                 '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态               '12当前状态
        strAddition = strAddition & ";"                             '13病种ID
        strAddition = strAddition & ";1"                            '14在职(1,2,3)
        strAddition = strAddition & ";" & .统筹编号 & "|" & .制卡日期 & "|" & .卡有效期 & "|" & .制卡单位 & "|" & .在职情况    '15退休证号
        strAddition = strAddition & ";" & .补卡次数                     '16年龄段
        strAddition = strAddition & ";" & .工况类别                            '17灰度级
        strAddition = strAddition & ";" & .帐户余额                             '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID)
    
    g病人身份_成都内江.lng病人ID = mlng病人ID
    
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    
    DebugTool "进入身份验证,并开始加入基本信息"
    
    If LoadBaseData = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    DebugTool "加入成功(身份验证)"
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Function LoadBaseData() As Boolean
    '加载基础数据
    Dim rsTemp As New ADODB.Recordset
    LoadBaseData = False
    On Error GoTo ErrHand:
      
    If mbytType = 0 Or mbytType = 3 Then
        cbo类别.AddItem "0-普通门诊"
    Else
        cbo类别.AddItem "1-普通住院"
    End If
    cbo类别.ListIndex = cbo类别.NewIndex
    LoadBaseData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    With g病人身份_成都内江
        txtEdit(0) = .卡号
        lblEdit(1) = .个人编号
        lblEdit(2) = .姓名
        lblEdit(3) = Decode(.性别, "1", "男", "2", "女", .性别)
        lblEdit(4) = .身份证号
        lblEdit(5) = .工况类别
        lblEdit(6) = .出生日期
        lblEdit(7) = .统筹编号
        lblEdit(8) = .年龄
        lblEdit(9) = .制卡日期
        lblEdit(10) = .卡有效期
        lblEdit(11) = .补卡次数
        lblEdit(12) = .单位号码
        lblEdit(13) = Format(.帐户余额, "####0.00;#####0.00; ;")
        lblEdit(14) = .制卡单位
        lblEdit(15) = .在职情况
   End With
End Sub
