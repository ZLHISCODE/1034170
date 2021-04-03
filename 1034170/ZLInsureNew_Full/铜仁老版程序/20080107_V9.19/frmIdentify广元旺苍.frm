VERSION 5.00
Begin VB.Form frmIdentify广元旺苍 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   300
      Left            =   810
      MaxLength       =   25
      TabIndex        =   5
      Tag             =   "社会保障号"
      Top             =   1320
      Width           =   2265
   End
   Begin VB.CommandButton cmd验卡 
      Caption         =   "重新读卡(&R)"
      Height          =   350
      Left            =   300
      TabIndex        =   25
      Top             =   3705
      Width           =   1305
   End
   Begin VB.ComboBox cbo社保 
      Height          =   300
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   2265
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5580
      TabIndex        =   23
      Top             =   3705
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4290
      TabIndex        =   22
      Top             =   3705
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   26
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -525
      TabIndex        =   24
      Top             =   3480
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "记录号"
      Height          =   180
      Index           =   2
      Left            =   3780
      TabIndex        =   6
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      Height          =   180
      Index           =   4
      Left            =   3600
      TabIndex        =   18
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   4425
      TabIndex        =   19
      Top             =   2565
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   4425
      TabIndex        =   3
      Top             =   900
      Width           =   2265
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "医保病人基本信息显示，可以通过[重新读卡]按钮重新进行读取病人基本信息。"
      Height          =   180
      Left            =   630
      TabIndex        =   27
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify广元旺苍.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      Height          =   180
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   8
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label lblInf 
      AutoSize        =   -1  'True
      Caption         =   "医保证号"
      Height          =   180
      Left            =   90
      TabIndex        =   4
      Top             =   1380
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   16
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   6
      Left            =   3600
      TabIndex        =   14
      Top             =   2205
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "社保机构"
      Height          =   180
      Index           =   7
      Left            =   90
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   450
      TabIndex        =   12
      Top             =   2205
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   10
      Left            =   90
      TabIndex        =   20
      Top             =   3045
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   810
      TabIndex        =   9
      Top             =   1740
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   4425
      TabIndex        =   7
      Top             =   1320
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   4425
      TabIndex        =   11
      Top             =   1740
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   810
      TabIndex        =   13
      Top             =   2160
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   810
      TabIndex        =   17
      Top             =   2565
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   810
      TabIndex        =   21
      Top             =   3000
      Width           =   5865
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   4425
      TabIndex        =   15
      Top             =   2145
      Width           =   2265
   End
End
Attribute VB_Name = "frmIdentify广元旺苍"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '第一次起动系统时调用
Private mblnChange As Boolean
Private Sub cbo社保_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd验卡_Click()
   If 获取参保人员信息 = False Then
        cmd确定.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmd确定.Enabled = True
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    cmd确定.Enabled = False
End Sub

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
    Dim strInput As String, strOutPut As String
    Dim lng状态 As Long
    
    IsValid = False
    If Trim(g病人身份_广元旺苍.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        If cmd验卡.Enabled Then cmd验卡.SetFocus
        Exit Function
    End If
    
     If cbo社保.Text = "" Then
        ShowMsgbox "社保机构还未选择"
        Exit Function
    End If
      
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查录前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=" & TYPE_广元旺苍 & " and 医保号='" & g病人身份_广元旺苍.医保证号 & "'"
            Call OpenRecordset(rsTemp, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
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
    Dim strInput  As String, strOutPut As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str社保 As String
    Dim int当前状态 As Integer
    Dim lng状态 As Long
    
    
    g病人身份_广元旺苍.机构编码 = Split(cbo社保.Text, "-")(0)
    g病人身份_广元旺苍.社保中心 = cbo社保.ItemData(cbo社保.ListIndex)
    If IsValid = False Then Exit Sub
    
    int当前状态 = 0
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & gintInsure & " and  医保号='" & g病人身份_广元旺苍.医保证号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人id, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_广元旺苍
        
        strIdentify = .医保卡号                                '0卡号
        strIdentify = strIdentify & ";" & .医保证号             '1医保号
        strIdentify = strIdentify & ";"                    '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", .性别)              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & .身份证号码            '6身份证
        strIdentify = strIdentify & ";" & .单位名称     '7.单位名称(编码)
        strAddition = ";0" & .社保中心                                           '8.中心代码
        strAddition = strAddition & ";" & .记录号                               '9.顺序号
        strAddition = strAddition & ";"                                '10人员身份
        strAddition = strAddition & ";" & .帐户余额                  '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";"             '13病种ID
        strAddition = strAddition & ";1"                        '14在职(1,2,3)
        strAddition = strAddition & ";" & .机构编码            '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";"                         '17灰度级
        strAddition = strAddition & ";"                         '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID)
    If mlng病人ID = 0 Then Exit Sub
    
    If mbytType = 0 Or mbytType = 3 Then
    Else
    End If
    g病人身份_广元旺苍.病人id = mlng病人ID
    
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
    If Load社保机构 = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    DebugTool "加入成功(身份验证)"
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g病人身份_广元旺苍
        lblEdit(0).Caption = .医保卡号
        txtEdit.Text = .医保证号
        lblEdit(1).Caption = .姓名
        lblEdit(2).Caption = .记录号
        lblEdit(3).Caption = Decode(.性别, "1", "男", "2", "女", .性别)
        lblEdit(4).Caption = .年龄
        lblEdit(5).Caption = .身份证号码
        lblEdit(6).Caption = .出生日期
        lblEdit(7).Caption = Format(.帐户余额, "####0.00;-####0.00;;")
        lblEdit(8).Caption = .单位名称
    End With
End Sub
Private Sub Form_Load()
        mblnFirst = True
End Sub

Private Function Load社保机构() As Boolean
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "" & _
        "   Select * From 保险中心目录 " & _
        "   Order by 编码"
        
    Err = 0
    On Error GoTo ErrHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption & "社保机构目录"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "不存社保机构目录，请在参数中下载机构!"
        Exit Function
    End If
    
    With rsTemp
        cbo社保.Clear
        Do While Not .EOF
            cbo社保.AddItem Nvl(!编码) & "--" & Nvl(!名称)
            cbo社保.ItemData(cbo社保.NewIndex) = Nvl(!序号, 0)
            .MoveNext
        Loop
    End With
    cbo社保.ListIndex = 0
    SetDefaultSel
    cbo社保.Enabled = False
    Load社保机构 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo ErrHand:
    Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
    If cbo社保.ListCount = 0 Then Exit Function
    For i = 0 To cbo社保.ListCount
        If Split(cbo社保.List(i), "--")(0) = strReg Then
            cbo社保.ListIndex = i
            Exit For
        End If
    Next
    If cbo社保.ListIndex < 0 Then
        cbo社保.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 获取参保人员信息() As Boolean
    '获取参保人员信息
    Dim strInput As String
    Dim strOutPut As String
    Dim strArr
    
    获取参保人员信息 = False
    
    
    Err = 0
    On Error GoTo ErrHand:
   
    If 业务请求_广元旺苍(获得参保人员资料, "", strOutPut) = False Then
        Call ClearData
        Exit Function
    End If
    
    strArr = Split(strOutPut, "||")
    '返回:医保卡号||医保证号||个人记录号||姓名||身份证号码||单位名称||性别||出生日期
    
    With g病人身份_广元旺苍
        .医保卡号 = strArr(0)
        .医保证号 = strArr(1)
        .记录号 = strArr(2)
        .姓名 = strArr(3)
        .身份证号码 = strArr(4)
        .单位名称 = strArr(5)
        .性别 = strArr(6)
        .出生日期 = strArr(7)
        .年龄 = Get年龄(.出生日期)
        .机构编码 = Split(cbo社保.Text, "--")(0)
    End With
    
    '获取帐户余额
    '    YBJGBH  PCHAR   保险机构编号
    '    CPASSWORD   PCHAR   持卡人卡密码
    '有问题，根据机构编号怎么获取.
    strInput = g病人身份_广元旺苍.机构编码
    strInput = strInput & vbTab & g病人身份_广元旺苍.密码
    If 业务请求_广元旺苍(获取帐户余额_旺苍, strInput, strOutPut) = False Then Exit Function
    g病人身份_广元旺苍.帐户余额 = Val(strOutPut)
    
    获取参保人员信息 = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get年龄(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as 年龄 from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
        Get年龄 = Int(Nvl(rsTemp!年龄, 0))
        Exit Function
    End If
    Exit Function
ErrHand:
End Function
Private Sub ClearData()
    Dim i As Long
    '清除相关信息
    With g病人身份_广元旺苍
        .医保卡号 = ""
        .医保证号 = ""
        .记录号 = ""
        .姓名 = ""
        .身份证号码 = ""
        .单位名称 = ""
        .性别 = ""
        .出生日期 = ""
        .年龄 = 0
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m文本式
End Sub

