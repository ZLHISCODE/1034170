VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentify重庆渝北 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmIdentify重庆渝北.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3525
      Left            =   -165
      TabIndex        =   39
      Top             =   5055
      Visible         =   0   'False
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   6218
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   285
      Left            =   6255
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3990
      Width           =   255
   End
   Begin VB.CommandButton cmd修改密码 
      Caption         =   "修改密码"
      Height          =   350
      Left            =   360
      TabIndex        =   38
      Top             =   4635
      Width           =   1100
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4590
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   908
      Width           =   1980
   End
   Begin VB.TextBox TxtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   930
      MaxLength       =   20
      TabIndex        =   1
      Top             =   908
      Width           =   2385
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -405
      TabIndex        =   25
      Top             =   4425
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   510
      Width           =   8340
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   9
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5490
      TabIndex        =   10
      Top             =   4635
      Width           =   1100
   End
   Begin VB.TextBox txt病种 
      Height          =   300
      Left            =   930
      TabIndex        =   7
      Top             =   3975
      Width           =   5610
   End
   Begin VB.ComboBox cbo类别 
      Height          =   300
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1275
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "支付类别"
      Height          =   180
      Index           =   14
      Left            =   195
      TabIndex        =   4
      Top             =   1335
      Width           =   720
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "病种(&F)"
      Height          =   180
      Left            =   285
      TabIndex        =   6
      Top             =   4035
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   13
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   4590
      TabIndex        =   37
      Top             =   3210
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   4590
      TabIndex        =   36
      Top             =   2805
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   930
      TabIndex        =   35
      Top             =   3615
      Width           =   5625
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4590
      TabIndex        =   34
      Top             =   2430
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   930
      TabIndex        =   33
      Top             =   2805
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   930
      TabIndex        =   32
      Top             =   3210
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   4590
      TabIndex        =   31
      Top             =   2025
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   930
      TabIndex        =   30
      Top             =   2430
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   930
      TabIndex        =   29
      Top             =   2025
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4590
      TabIndex        =   28
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   930
      TabIndex        =   27
      Top             =   1650
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   26
      Top             =   1283
      Width           =   1980
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   12
      Left            =   3840
      TabIndex        =   22
      Top             =   3255
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗补助类别"
      Height          =   180
      Index           =   11
      Left            =   3480
      TabIndex        =   19
      Top             =   2850
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   10
      Left            =   210
      TabIndex        =   21
      Top             =   3667
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗照顾类别"
      Height          =   180
      Index           =   9
      Left            =   3480
      TabIndex        =   17
      Top             =   2475
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   570
      TabIndex        =   18
      Top             =   2850
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位编码"
      Height          =   180
      Index           =   7
      Left            =   210
      TabIndex        =   20
      Top             =   3262
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗人员类别"
      Height          =   180
      Index           =   6
      Left            =   3480
      TabIndex        =   15
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出身日期"
      Height          =   180
      Index           =   5
      Left            =   210
      TabIndex        =   16
      Top             =   2482
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   14
      Top             =   2077
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   4200
      TabIndex        =   13
      Top             =   1695
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   570
      TabIndex        =   12
      Top             =   1702
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      Height          =   180
      Index           =   1
      Left            =   3855
      TabIndex        =   11
      Top             =   1335
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保卡号"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify重庆渝北.frx":000C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "通过IC卡验证人员身份，并将验证结果信息显示出来。"
      Height          =   180
      Left            =   630
      TabIndex        =   24
      Top             =   270
      Width           =   4320
   End
End
Attribute VB_Name = "frmIdentify重庆渝北"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
'API的医保接口声明
Private Type Struct
    lngAppCode  As Long   '标志服务执行状态代码。等于1时表示服务执行正常结束，小于0时表示服务执行异常或错误。
    strErrMsg  As String  '当服务执行状态代码AppCod小于0时，描述服务执行的异常或错误信息。
End Type
'获取就诊编号
Private Declare Function GetAKC190 Lib "YHMdcrAsistntSvr.dll" Alias "_GetAKC190@12" (ByVal strYab003 As String, ByRef strAkc190 As String, ByRef tmpStrut As Struct) As Boolean
     
Dim mblnChange As Boolean
Private Sub cbo类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd病种_Click()
     
        Dim rsTemp As New ADODB.Recordset
        
        With rsTemp
            If .State = 1 Then .Close

            gstrSQL = "" & _
                "   Select id, 编码, 名称, 支付类别, 助记码, 病种结算办法, 经办构构代码 " & _
                "   From 医保病种目录"
                
            Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
            .Open gstrSQL, gcnOracle_CQYB
            Call SQLTest
            If .EOF Then
                MsgBox "不存在任何病种,请下载！", vbInformation, gstrSysName
                Exit Sub
            End If
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt病种.Top - .Height
                    .Left = txt病种.Left + txt病种.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = 1000
                    .ColWidth(5) = 1400
                    .ColWidth(6) = 2000
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt病种 = "[" & Nvl(!编码) & "]" & IIf(IsNull(!名称), "", !名称)
                txt病种.Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
End Sub

Private Sub cmd修改密码_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    
    strNewPassWord = frm修改密码.ChangePassword(strOldPassWord, strOldPassWord)
    If strOldPassWord = strNewPassWord Then Exit Sub
    If strNewPassWord = "" Then Exit Sub
      
    If 修改密码_重庆渝北(strOldPassWord, strNewPassWord) = True Then
        g病人身份_重庆渝北.密码 = strNewPassWord
        cmd确定_Click
        Unload Me
        Exit Sub
    End If
End Sub



Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        txtEdit(Index).Tag = ""
    End If
    If Index = 0 And mblnChange = False Then
        g病人身份_重庆渝北.个人编号 = ""
        g病人身份_重庆渝北.卡号 = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    mblnChange = True
    If Index = 0 Then
        SetOKCtrl False
        mblnChange = True
    ElseIf Index = 1 Then
        '密码输入完毕
        '需获取病人信息
         SetOKCtrl False
        
        '需解析卡内数据
        If 解析卡_重庆渝北 = False Then
            Exit Sub
        End If
         If Trim(txtEdit(Index)) = "" Then
            If mbytType = 0 Then
                '如果是门诊,需检查是否当前就诊的,且已经存在该帐户时,不重新输入密码.
                                
                '取密码
                 gstrSQL = "Select 密码,就诊时间 From 保险帐户  where 险类=" & TYPE_重庆渝北 & " and 医保号='" & g病人身份_重庆渝北.个人编号 & "'"
                 zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                 
                 If rsTemp.RecordCount = 0 Then
                     ShowMsgbox "请输入密码!"
                    txtEdit(Index).SetFocus
                     Exit Sub
                 End If
                 If Format(rsTemp!就诊时间, "yyyy-mm-dd") <> Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                    ShowMsgbox "请输入密码!"
                    txtEdit(Index).SetFocus
                    Exit Sub
                 End If
                 txtEdit(Index) = Trim(Nvl(rsTemp!密码))
                 If txtEdit(Index) = "" Then
                    ShowMsgbox "请输入密码!"
                    txtEdit(Index).SetFocus
                    Exit Sub
                 End If
            Else
                ShowMsgbox "请输入密码!"
                txtEdit(Index).SetFocus
                Exit Sub
            End If
         End If
         
        txtEdit(0).Text = g病人身份_重庆渝北.卡号
        lblEdit(1).Caption = g病人身份_重庆渝北.个人编号
         
         g病人身份_重庆渝北.密码 = Trim(txtEdit(Index))
        If g病人身份_重庆渝北.卡号 = "" Then
            g病人身份_重庆渝北.卡号 = Trim(txtEdit(0).Text)
        End If
        If 身份鉴别_重庆渝北 = False Then
            Exit Sub
        End If
        
        If g病人身份_重庆渝北.姓名 = "" Then
            ShowMsgbox "无效的用户验证,请核查!"
            Exit Sub
        End If
        
        
        '如果是门诊,需先进行挂号处理,否则是不能进行相应的处理的.
        If mbytType = 0 Then
            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            gstrSQL = "Select 1 From 病人费用记录 " & _
                    "   where 记录状态=1 and 记录性质=4  and rownum<=1 and 登记时间 between to_date('" & strCurrDate & " 00:00:00','yyyy-mm-dd hh24:mi:ss') and to_date('" & strCurrDate & " 23:59:59','yyyy-mm-dd hh24:mi:ss') and 病人id in (select 病人id From 保险帐户  where 险类=" & TYPE_重庆渝北 & " and 医保号='" & g病人身份_重庆渝北.个人编号 & "')"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
            If rsTemp.RecordCount = 0 Then
                ShowMsgbox "该医保病人未进行挂号,不能进行门诊结算!"
                Exit Sub
            End If
        End If
        '初始值
        Call LoadCtrlData
        SetOKCtrl True
    End If
    zlCommFun.PressKey vbKeyTab
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
    IsValid = False
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "还没有输入医保卡号！", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    
    If Trim(g病人身份_重庆渝北.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Trim(txt病种) <> "" And Val(txt病种.Tag) = 0 Then
        ShowMsgbox "病种选择错误,请重新选择!"
        txt病种.SetFocus
        Exit Function
    End If
    If cbo类别.Text = "" Then
        ShowMsgbox "支付类别未选择"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查录前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=" & TYPE_重庆渝北 & " and 医保号='" & g病人身份_重庆渝北.个人编号 & "'"
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
    
    lng疾病ID = Val(txt病种.Tag)
    
    If lng疾病ID <> 0 And txt病种.Text <> "" Then
        g病人身份_重庆渝北.病种编码 = Mid(txt病种.Text, 2, InStr(1, txt病种.Text, "]") - 2)
    Else
        g病人身份_重庆渝北.病种编码 = "000000"
    End If
    g病人身份_重庆渝北.病种ID = lng疾病ID
    
    g病人身份_重庆渝北.支付类别 = Mid(cbo类别.Text, 1, InStr(1, cbo类别.Text, "-") - 1)
    int当前状态 = 0
    
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & gintInsure & " and  医保号='" & g病人身份_重庆渝北.个人编号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_重庆渝北
        
        strIdentify = .卡号                               '0卡号
        strIdentify = strIdentify & ";" & .个人编号           '1医保号
        strIdentify = strIdentify & ";" & .密码                 '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", "未知")              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & .身份证号           '6身份证
        strIdentify = strIdentify & ";" & .单位名称 & IIf(.单位编码 = 0, "", "(" & .单位编码 & ")")          '7.单位名称(编码)
        strAddition = ";0"                                          '8.中心代码
        strAddition = strAddition & ";"                             '9.顺序号
        strAddition = strAddition & ";" & .社保经办构构代码          '10人员身份
        strAddition = strAddition & ";" & .帐户余额                 '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";" & IIf(lng疾病ID = 0, "", lng疾病ID)             '13病种ID
        strAddition = strAddition & ";1"                            '14在职(1,2,3)
        strAddition = strAddition & ";" & .医疗人员类别 & "|" & .医疗照顾类别 & "|" & .医疗补助类别 & "|" & .累计缴费月数     '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";"                             '17灰度级
        strAddition = strAddition & ";" & .帐户余额                             '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID)
    
    If mbytType = 3 Or mbytType = 1 Then
        '如果是挂号或入院登记,需确定新的就诊编号
        g病人身份_重庆渝北.就诊编号 = Get就诊编号_重庆渝北
        If g病人身份_重庆渝北.就诊编号 = "" Then
            ShowMsgbox "在获取就诊编号时为空了,请检查"
            Exit Sub
        End If
        
        '更新保险帐户的相关信息
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & gintInsure & ",'就诊编号','''" & g病人身份_重庆渝北.就诊编号 & "''')"
        Call ExecuteProcedure("保存就诊编号")
        
        If mbytType = 1 Then
            '为了保证先按普通入院再进行补充入院的就诊时间需更改.
             gstrSQL = "Select 入院日期 From 病案主页 where 病人id=" & mlng病人ID & " And 出院日期 is null"
             zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
             If Not rsTemp.EOF Then
                    '应该是补充登记
                    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & gintInsure & ",'就诊时间','" & Format(rsTemp!入院日期, "yyyy-mm-dd HH:MM:SS") & "',1)"
                    Call ExecuteProcedure("保存就诊时间")
             End If
        End If
    Else
        '就诊时间还原
        '更新保险帐户的相关信息
        If g病人身份_重庆渝北.就诊时间 <> "" Then
            gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & gintInsure & ",'就诊时间','" & g病人身份_重庆渝北.就诊时间 & "',1)"
            Call ExecuteProcedure("保存就诊时间")
        End If
    End If
    
    '取保险帐户中的就诊编号
     gstrSQL = "Select 就诊编号,就诊时间 From 保险帐户  where 病人id=" & mlng病人ID & " and 险类=" & gintInsure
     zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
     If rsTemp.RecordCount = 0 Then
         ShowMsgbox "在保险帐户中不存在该病人"
         Exit Sub
     End If
    g病人身份_重庆渝北.就诊编号 = Nvl(rsTemp!就诊编号)
    g病人身份_重庆渝北.就诊时间 = Format(rsTemp!就诊时间, "yyyy-MM-dd HH:mm:ss")
    g病人身份_重庆渝北.lng病人ID = mlng病人ID
    
    '更新保险帐户的相关信息
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & gintInsure & ",'支付类别','''" & g病人身份_重庆渝北.支付类别 & "''')"
    Call ExecuteProcedure("保存就诊编号")
    
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
    
    With rsTemp
    
        .Open "Select * From 支付类别 where 标志=2 or 标志=" & IIf(mbytType = 3, 0, IIf(mbytType = 4, 1, mbytType)) & " order by 编码", gcnOracle_CQYB
        Do While Not .EOF
            cbo类别.AddItem Nvl(!编码) & "-" & Nvl(!名称)
            If !缺省 = 1 Then
                cbo类别.ListIndex = cbo类别.NewIndex
            End If
            .MoveNext
        Loop
        If cbo类别.ListIndex < 0 Then
            If cbo类别.ListCount <> 0 Then
                cbo类别.ListIndex = 0
            End If
        End If
    End With
    If cbo类别.ListCount = 0 Then
        ShowMsgbox "支付类别未初始化,请与系统管理员联系!"
        Exit Function
    End If
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
    With g病人身份_重庆渝北
        lblEdit(2).Caption = .姓名
        lblEdit(3).Caption = Decode(.性别, "1", "男", "2", "女", "未知")
        lblEdit(4).Caption = .身份证号
        lblEdit(5).Caption = .出生日期
        lblEdit(6).Caption = Get代码数据_重庆渝北(医疗人员类别, .医疗人员类别)
        lblEdit(7).Caption = .单位编码
        lblEdit(8).Caption = .年龄
        '目前没有类别
        lblEdit(9).Caption = ""          'Get代码数据_重庆渝北(医疗照顾类别, .医疗照顾类别)
        lblEdit(10).Caption = .单位名称
        lblEdit(11).Caption = Get代码数据_重庆渝北(医疗补助类别, .医疗补助类别)
        lblEdit(12).Caption = Format(.帐户余额, "####0.00;#####0.00; ;")
    End With
    
    gstrSQL = "Select 病种ID,支付类别,就诊时间 from 保险帐户 where 医保号='" & g病人身份_重庆渝北.个人编号 & "' and 险类=" & TYPE_重庆渝北
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取相关病种"
    If rsTemp.EOF Then Exit Sub
    g病人身份_重庆渝北.支付类别 = Nvl(rsTemp!支付类别)
    g病人身份_重庆渝北.就诊时间 = Format(rsTemp!就诊时间, "yyyy-MM-dd HH:mm:ss")
    
    gstrSQL = "Select * From 医保病种目录 where ID=" & Nvl(rsTemp!病种ID, 0)
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.ProductName, "获取病种信息", gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle_CQYB
    Call SQLTest
    If rsTemp.EOF Then
        Exit Sub
    End If
    txt病种.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
    txt病种.Tag = Nvl(rsTemp!ID, 0)
    Dim i As Long
    For i = 0 To cbo类别.ListCount - 1
        If InStr(1, cbo类别.List(i), g病人身份_重庆渝北.支付类别 & "-") <> 0 Then
            cbo类别.ListIndex = i
            Exit For
        End If
    Next
    
End Sub
Private Sub mshSelect_Click()
    With mshSelect
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            SetColumnSort mshSelect, mintPreCol, mintsort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshSelect
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub


'对列头进行排序
Private Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer)
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
               .Sort = flexSortStringNoCaseAscending
               intPreSort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               intPreSort = flexSortStringNoCaseDescending
            End If
            intPreCol = intCol
            .Row = FindRow(mshFilter, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub


Private Sub txt病种_Change()
    txt病种.Tag = ""
End Sub

Private Sub txt病种_GotFocus()
    OpenIme GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
    zlControl.TxtSelAll txt病种
End Sub

Private Sub txt病种_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    If KeyCode = vbKeyReturn Then
        If Me.txt病种 = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If Trim(txt病种) = "" Then Exit Sub
        If Trim(txt病种.Tag) <> "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txt病种 = UCase(txt病种)
    
        Dim rsTemp As New ADODB.Recordset
        
        With rsTemp
            If .State = 1 Then .Close

            gstrSQL = "" & _
                "   Select id, 编码, 名称, 支付类别, 助记码, 病种结算办法, 经办构构代码 " & _
                "   From 医保病种目录" & _
                "   Where " & zlCommFun.GetLike("", "编码", Me.txt病种) & " Or " & _
                            zlCommFun.GetLike("", "名称", Me.txt病种) & " Or " & _
                            zlCommFun.GetLike("", "助记码", Me.txt病种)
            
            Call SQLTest(App.ProductName, Me.Caption, strSql)
            .Open gstrSQL, gcnOracle_CQYB
            Call SQLTest
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt病种.Top - .Height
                    .Left = txt病种.Left + txt病种.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = 1000
                    .ColWidth(5) = 1400
                    .ColWidth(6) = 1400
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt病种 = "[" & Nvl(!编码) & "]" & IIf(IsNull(!名称), "", !名称)
                txt病种.Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
    End If
End Sub

Private Sub txt病种_LostFocus()
    OpenIme ""
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            txt病种.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
            txt病种.Tag = .TextMatrix(.Row, 0)
            If cmd确定.Enabled Then cmd确定.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub
'寻找与某一单元值相等的行
Private Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .Rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function
