VERSION 5.00
Begin VB.Form frmIdentify南京市 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病保病人身份验证"
   ClientHeight    =   3480
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4635
   Icon            =   "frmIdentify南京市.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra基本 
      Caption         =   "医保病人基本信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4404
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   10
         Top             =   1860
         Width           =   1692
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   8
         Top             =   855
         Width           =   1692
      End
      Begin VB.CommandButton cmd病种信息 
         Caption         =   "…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3624
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1335
         Width           =   372
      End
      Begin VB.TextBox txt门诊病种 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   4
         Top             =   1335
         Width           =   1692
      End
      Begin VB.TextBox txt姓名 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "正确姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   11
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   9
         Top             =   915
         Width           =   480
      End
      Begin VB.Label lbl门诊病种 
         AutoSize        =   -1  'True
         Caption         =   "门诊病种(&F)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   3
         Top             =   1410
         Width           =   1320
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "单据号(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   1
         Top             =   420
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3432
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2232
      TabIndex        =   6
      Top             =   2880
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify南京市"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstrIdentify As String
Private mlng病人ID As Long, mlng病种ID As Long
Private mstr病人姓名 As String
Private mstr病种编码 As String
Private mstr病种名称 As String

Private Sub cmdCancle_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim strIdentify As String
    Dim strAddition As String
    Dim lngsequence As String
    On Error GoTo errHandle
    
    '判断是否输入医保病人姓名
    If Trim(Text1.Text) = "" Then
        MsgBox "未提取到医保病人姓名", vbInformation, gstrSysName
        txt姓名.SetFocus
        Exit Sub
    End If
    
    mstr病人姓名 = Trim(Text1.Text)
    
    If Trim(txt门诊病种.Text) = "" Or txt门诊病种 <> mstr病种名称 Then
        MsgBox "门诊病种未录入或有误", vbInformation, gstrSysName
        txt门诊病种.SetFocus
        Exit Sub
    End If
        
    '此处无法取得卡号和医保号,所以暂时填入保险病种序列,以后得到卡号后再进行修改
    lngsequence = Right(String(20, "0") & Text1.Tag, 20)
'      strInfo='0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
'      8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(1,2,3);15退休证号;16年龄段;17灰度级
'      18帐户增加累计;19帐户支出累计;20进入统筹累计;21统筹报销累计;22住院次数累计;23就诊类别
'      24本次起付线;25起付线累计;26基本统筹限额
    
    strIdentify = lngsequence & ";"                                       '0卡号
    strIdentify = strIdentify & lngsequence & ";"                  '1医保号（个人编号）
    strIdentify = strIdentify & ";"                                 '2密码
    strIdentify = strIdentify & mstr病人姓名 & ";"                   '3姓名
    strIdentify = strIdentify & ";"                                 '4性别
    strIdentify = strIdentify & ";"                                '5出生日期
    strIdentify = strIdentify & ";"                                 '6身份证
    strIdentify = strIdentify & ";"                               '7.单位名称(编码)
    strAddition = "0;"                                          '8.中心代码
    strAddition = strAddition & ";"                               '9.顺序号
    strAddition = strAddition & ";"                            '10人员身份
    strAddition = strAddition & "10000;"                              '11帐户余额
    strAddition = strAddition & "0;"                            '12当前状态
    strAddition = strAddition & mlng病种ID & ";"                 '13病种ID
    strAddition = strAddition & "1;"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";"                             '18帐户增加累计
    strAddition = strAddition & ";"                            '19帐户支出累计
    strAddition = strAddition & "0;"                            '20进入统筹累计
    strAddition = strAddition & "0;"                            '21统筹报销累计
    strAddition = strAddition & "0;"                             '22住院次数累计
    strAddition = strAddition & ";"                             '23就诊类型
    
    If mlng病人ID = 0 Then
        mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID)
    End If
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrIdentify = strIdentify & mlng病人ID & ";" & strAddition
    End If
    If Trim(Text2.Text) <> "" Then
        mstr病人姓名 = Trim(Text2.Text)
    Else
        mstr病人姓名 = Trim(Text1.Text)
    End If
    
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MsgBox mlng病人ID & "，" & lngsequence & "，" & Text1.Tag & "，" & strIdentify & strAddition, vbInformation, gstrSysName
End Sub


Private Sub cmd病种信息_Click()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select id,编码,名称,decode(类别,1,'慢性病',2,'特殊病','普通病') as 病种 from 保险病种 where 险类=" & TYPE_南京市
    Call OpenRecordset(rsTemp, "选择病种")
    
    If frmListSel.ShowSelect(rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：") Then
        txt门诊病种.Text = rsTemp!名称
        mlng病种ID = rsTemp!ID
        mstr病种编码 = rsTemp!编码
        mstr病种名称 = rsTemp!名称
    Else
        txt门诊病种.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If mbytType = 0 Then
        txt门诊病种.Enabled = True
    Else
        txt门诊病种.Enabled = False
    End If
End Sub



Private Sub txt门诊病种_GotFocus()
    OpenIme ("")
    Call zlControl.TxtSelAll(txt门诊病种)
End Sub

Private Sub txt门诊病种_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strText As String
    Dim blnReturn As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errorhandle
    '读出门诊病种
    
    strText = txt门诊病种.Text
    gstrSQL = "select A.id,A.编码,A.名称 from 保险病种 A where A.险类=" & TYPE_南京市 & " and (" & _
              zlCommFun.GetLike("A", "编码", strText) & " or " & zlCommFun.GetLike("A", "名称", strText) & " or " & zlCommFun.GetLike("A", "简码", strText) & ")"
    Call OpenRecordset(rsTemp, "门诊病种")
    
    If rsTemp.RecordCount = 1 Then
        blnReturn = True
    Else
        blnReturn = frmListSel.ShowSelect(rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
    End If
    
    If blnReturn Then
        txt门诊病种.Text = rsTemp!名称
        mlng病种ID = rsTemp!ID
        mstr病种编码 = rsTemp!编码
        mstr病种名称 = rsTemp!名称
        zlCommFun.PressKey (vbKeyTab)
    Else
        txt门诊病种_GotFocus
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt姓名_GotFocus()
    Call zlControl.TxtSelAll(txt姓名)
End Sub

Public Function Identify(ByVal bytType As Byte) As String
    mbytType = bytType
    Me.Show 1
    Identify = mstrIdentify
    With gPatInfo_南京市
        .病人姓名 = mstr病人姓名
        .病种编码 = mstr病种编码
        .病种名称 = mstr病种名称
    End With
End Function

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    Dim strInput As String, strSql As String, rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt姓名 <> "" Then
        If Left(txt姓名.Text, 1) <> "." Then Exit Sub
        strInput = txt姓名.Text
        If Not IsNumeric(Mid(strInput, 2)) Then Exit Sub
        If Len(Mid(strInput, 2)) <= 4 Then
            strSql = PreFixNO & Format(CDate("1992-" & Format(zlDatabase.Currentdate, "MM-dd")) - CDate("1992-01-01") + 1, "000") & Format(Mid(strInput, 2), "0000") '按天顺序编号
        Else
            strSql = GetFullNO(Mid(strInput, 2))
        End If
        '门诊记帐时必须要挂号建档
        strSql = "Select 病人id,姓名,标识号 From 病人费用记录 Where NO='" & strSql & "' And 记录性质=4 And 记录状态=1"
        Set rsTemp = gcnOracle.Execute(strSql)
        If rsTemp.EOF Then
            MsgBox "错误的挂号单号", vbInformation, gstrSysName
            Exit Sub
        End If
        strSql = "Select * From 病人信息 Where 病人ID=" & rsTemp!病人ID
        Set rsTemp = gcnOracle.Execute(strSql)
        If rsTemp.EOF Then
            MsgBox "读取病人信息出错", vbInformation, gstrSysName
            Exit Sub
        ElseIf IsNull(rsTemp!门诊号) Then
            MsgBox "该病人的医保卡号没有录入", vbInformation, gstrSysName
            Exit Sub
        Else
            Text1.Text = rsTemp!姓名
            Text1.Tag = rsTemp!门诊号
        End If
        zlCommFun.PressKey (vbKeyTab)
    End If
    Exit Sub
errHandle:
    MsgBox "此医保病人没有建立病案，请重新挂号并建立病案", vbInformation, gstrSysName
End Sub


