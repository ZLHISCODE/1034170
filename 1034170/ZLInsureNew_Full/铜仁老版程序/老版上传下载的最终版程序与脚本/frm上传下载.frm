VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm上传下载 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "上传下载程序"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frm上传下载.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8130
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox pic目录 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   300
      ScaleHeight     =   1305
      ScaleWidth      =   7575
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   870
      Width           =   7575
      Begin VB.ComboBox cbo社保局 
         Height          =   300
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   930
         Width           =   5355
      End
      Begin VB.CommandButton cmd目录 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   7200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
      End
      Begin VB.CommandButton cmd目录 
         Caption         =   "…"
         Height          =   240
         Index           =   1
         Left            =   7200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   585
         Width           =   255
      End
      Begin VB.TextBox txtDown 
         Height          =   300
         Left            =   2130
         MaxLength       =   40
         TabIndex        =   6
         Top             =   540
         Width           =   5355
      End
      Begin VB.TextBox txtUp 
         Height          =   300
         Left            =   2130
         MaxLength       =   40
         TabIndex        =   3
         Top             =   150
         Width           =   5355
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "社保局"
         Height          =   180
         Index           =   0
         Left            =   1530
         TabIndex        =   8
         Top             =   990
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frm上传下载.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "本地下载目录"
         Height          =   180
         Index           =   7
         Left            =   990
         TabIndex        =   5
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "本地上传目录"
         Height          =   180
         Index           =   6
         Left            =   990
         TabIndex        =   2
         Top             =   210
         Width           =   1080
      End
   End
   Begin MSComctlLib.TabStrip tabHost 
      Height          =   1965
      Left            =   135
      TabIndex        =   0
      Top             =   390
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   3466
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra进度 
      Caption         =   "进度报告"
      Height          =   705
      Left            =   135
      TabIndex        =   19
      Top             =   4560
      Width           =   7905
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   255
         Left            =   900
         TabIndex        =   21
         Top             =   300
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl项目 
         AutoSize        =   -1  'True
         Caption         =   "项目"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.Frame fra上传 
      Caption         =   "上传"
      Height          =   1965
      Left            =   4125
      TabIndex        =   11
      Top             =   2490
      Width           =   3915
      Begin VB.CommandButton cmd恢复 
         Caption         =   "恢复性上传"
         Height          =   350
         Left            =   2640
         TabIndex        =   18
         Top             =   780
         Width           =   1100
      End
      Begin VB.CommandButton cmd上传 
         Caption         =   "开始上传"
         Height          =   350
         Left            =   2640
         TabIndex        =   17
         Top             =   1440
         Width           =   1100
      End
      Begin VB.Label lbl上传 
         AutoSize        =   -1  'True
         Caption         =   "最近上传日期："
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1110
         Width           =   1260
      End
      Begin VB.Label lbl上传说明 
         Caption         =   "   上传数据每天只能执行一次。如果失败可以重新执行。"
         Height          =   495
         Left            =   1110
         TabIndex        =   13
         Top             =   360
         Width           =   2625
      End
      Begin VB.Image img上传 
         Height          =   480
         Left            =   210
         Picture         =   "frm上传下载.frx":0884
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame fra下载 
      Caption         =   "下载"
      Height          =   1965
      Left            =   135
      TabIndex        =   10
      Top             =   2490
      Width           =   3915
      Begin VB.CommandButton cmd下载 
         Caption         =   "开始下载"
         Height          =   350
         Left            =   2580
         TabIndex        =   15
         Top             =   1440
         Width           =   1100
      End
      Begin VB.Label lbl下载 
         AutoSize        =   -1  'True
         Caption         =   "最近下载日期："
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   1050
         Width           =   1260
      End
      Begin VB.Image img下载 
         Height          =   480
         Left            =   210
         Picture         =   "frm上传下载.frx":114E
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lbl下载说明 
         Caption         =   "   在第一次使用医保接口前必须完成先进行一次下载。下载程序可以随时进行。"
         Height          =   615
         Left            =   1110
         TabIndex        =   12
         Top             =   360
         Width           =   2565
      End
   End
   Begin VB.Label lbl标题 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注：由于不同主机位于不同的网络中，所以上传下载只能针对当前主机进行。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   90
      Width           =   6120
   End
End
Attribute VB_Name = "frm上传下载"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'用于下载文件
Private Enum HostField
    hf主机代码 = 0
    hf主机名称 = 1
    hf医保年 = 2
    hf有效期 = 3
    hf装钱序号 = 4
    hf黑名单下载序号 = 5
    hf病种下载序号 = 6
    hf项目下载序号 = 7
    hf政策下载序号 = 8
    hf离休干部下载序号 = 9
    hf补充人员下载序号 = 10
    hf是否可用 = 11
    hf代理下发 = 12
End Enum

Private Enum HostParamField
    hp主机代码 = 0
    hp起始日期 = 1
    hp终止日期 = 2
    hp上传IP = 3
    hp上传用户 = 4
    hp上传密码 = 5
    hp上传目录 = 6
    hp下载IP = 7
    hp下载用户 = 8
    hp下载密码 = 9
    hp下载目录 = 10
    hp装钱IP地址1 = 13
    hp装钱IP地址2 = 14
    hp代理下发 = 15
End Enum

Private Enum CenterField
    cf主机代码 = 0
    cf中心代码 = 1
    cf中心名称 = 2
    cf个人账户可支付首先自付 = 3
    cf可写卡冲票 = 4
End Enum

Private Enum PolicyField
    pol中心代码 = 0
    pol算法 = 1
    pol数值 = 2
    pol次数限制 = 3
    pol起付线限制 = 4
    pol减起付线限制 = 5
    pol起付线在段中 = 6
    pol统筹封顶线 = 7
    pol统筹段数 = 8
    pol段值类型 = 9
    pol封顶类型 = 10
    pol使用累计报销 = 11
    pol当年慢性病月份 = 12
    pol开展补充保险报销 = 13
    pol补充报销比例 = 14
    pol补充报销限额 = 15
    pol补充报销限额类型 = 16
    pol补充报销减起付金 = 17
    pol开展补助报销 = 18
    pol开展慢病报销 = 19
    pol开展大病报销 = 20
    pol乙类项目价格 = 21
    pol跨年起付金类型 = 22 '0-补原起付金；1补今年差价；2交起付金
    pol跨年增加住院次数 = 23 '1增加一次，0不增加
End Enum

Private Enum ParamField
    par中心代码 = 0
    par单据类型 = 1
    par医院等级 = 2
    par职工身份 = 3
    par起付线 = 4
    par第一段起始值 = 5
    par第一段报销比例 = 6
    par第二段起始值 = 7
    par第二段报销比例 = 8
    par第三段起始值 = 9
    par第三段报销比例 = 10
    par第四段起始值 = 11
    par第四段报销比例 = 12
    par第五段起始值 = 13
    par第五段报销比例 = 14
End Enum

Private Enum ItemField
    if项目序号 = 0
    if数字编码 = 1
    if拼音编码 = 2
    if药典名称 = 3
    if单位 = 4
    if剂型编码 = 5
    if大类编码 = 6
    if是否是药 = 7
    if是否医保 = 8
    if最大价格限制 = 9
    if首先自付比例 = 10
    if价格 = 11
    if项目内涵 = 12
    if除外内容 = 13
    if说明 = 14
    if省级限价 = 15
    if市级限价 = 16
    if县级限价 = 17
    if乡级限价 = 18
    if特检项目 = 19
    if特检自付比例 = 20
End Enum

'作为公共使用的文件系统变量
Dim mobjFileSys As New FileSystemObject

Private mstr主机编码 As String
Private mstr主机名称 As String
Private mstr中心InOracle As String
Private mstr序号InOracle As String
Private mstr中心InStr As String

Private mstr原医保年 As String
Private mlng装钱序号 As Long
Private mlng黑名单下载序号 As Long
Private mlng病种下载序号 As Long
Private mlng项目下载序号 As Long
Private mlng政策下载序号 As Long
Private mlng离休干部序号 As Long
Private mlng补充人员下载序号 As Long

Private mstr上传IP As String
Private mstr上传用户 As String
Private mstr上传密码 As String
Private mstr下载IP As String
Private mstr下载用户 As String
Private mstr下载密码 As String
Private mstr远程上传目录 As String
Private mstr远程下载目录 As String
Private mstr本地上传目录 As String
Private mstr本地下载目录 As String

Private mstr医院编码 As String
Private mstr医院级别 As String

'上传程序专用
Private mstr开始日期 As String
Private mstr结束日期 As String
Private mstr日结日期 As String '也就是费用发生日期
Private mstr缺省开始日期 As String

Private mdat开始日期 As Date
Private mdat结束日期 As Date

Private mblnLoad As Boolean

Private Sub cbo社保局_Click()
    gcn医保.Execute "Update 保险主机 A set 社保局='" & cbo社保局.Text & "' where  A.险类 = " & TYPE_铜仁市 & " And A.编码 = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    If mblnLoad = False Then Exit Sub
    
    On Error GoTo errHandle
    
    gstrSQL = "SELECT A.名称,A.编码 FROM 保险主机 A,保险主机参数 B " & _
              " Where A.险类 = " & TYPE_铜仁市 & " And A.险类 = B.险类 And A.编码 = B.主机 " & _
              "    AND nvl(B.起始日期,to_date('2000-01-01','yyyy-MM-dd'))<=SYSDATE  AND nvl(B.终止日期,to_date('3000-01-01','yyyy-MM-dd'))>=trunc(SYSDATE)" & _
              " Order by A.编码"
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有发现可以进行上传下载的医保主机，请检查初始化数据是否正确。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    tabHost.Tabs.Clear
    Do Until rsTemp.EOF
        tabHost.Tabs.Add , "K" & rsTemp("编码"), rsTemp("名称")
        rsTemp.MoveNext
    Loop
    tabHost.Tabs(1).Selected = True
    
    mblnLoad = False
    
    '执行相应的功能
    If gintType = 1 Then
        '下载
        Call cmd下载_Click
    ElseIf gintType = 2 Then
        '上传
        Call cmd上传_Click
    End If
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    
End Sub

Private Sub tabHost_Click()
    Dim i As Integer, j As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '取该中心下属社保局
    Me.cbo社保局.Clear
    gstrSQL = "Select 编码||'-'||名称 AS 社保局 from zlyb.保险中心目录 where 主机编码='" & Mid(tabHost.SelectedItem.Key, 2) & "' order by 序号"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    With rsTemp
        Do While Not .EOF
            Me.cbo社保局.AddItem !社保局
            .MoveNext
        Loop
    End With
    
    gstrSQL = "SELECT A.本地上传地址,A.本地下载地址,A.社保局 FROM 保险主机 A " & _
              " Where A.险类 = " & TYPE_铜仁市 & " And A.编码 = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    txtDown.Text = NVL(rsTemp("本地下载地址"))
    txtUp.Text = NVL(rsTemp("本地上传地址"))
    j = Me.cbo社保局.ListCount
    For i = 1 To j
        If Me.cbo社保局.List(i - 1) = NVL(rsTemp!社保局) Then
            Me.cbo社保局.ListIndex = i - 1
            Me.cbo社保局.Enabled = False
            Exit For
        End If
    Next
    
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Sub cmd目录_Click(Index As Integer)
    Dim strTitle As String
    Dim strPath As String
    
    If Index = 0 Then
        strTitle = "请选择保存上传文件的目录："
    Else
        strTitle = "请选择保存下载文件的目录："
    End If
    
    strPath = OpenDir(Me, strTitle)
    If StrIsValid(strPath, 50, , "目录名") = False Then
        Exit Sub
    End If
    If strPath <> "" Then
        '保存目录名
        If Index = 0 Then
            gcn医保.Execute "Update 保险主机 A set 本地上传地址='" & strPath & "' where  A.险类 = " & TYPE_铜仁市 & " And A.编码 = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
            txtUp.Text = strPath
        Else
            gcn医保.Execute "Update 保险主机 A set 本地下载地址='" & strPath & "' where  A.险类 = " & TYPE_铜仁市 & " And A.编码 = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
            txtDown.Text = strPath
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    lbl下载.Caption = "最近下载日期：无"
    lbl上传.Caption = "最近上传日期：无"
    lbl下载.Tag = ""
    lbl上传.Tag = ""
    
    gstrSQL = "select 传输模式,max(操作日期) as 日期 from 上传下载 group by 传输模式"
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Do Until rsTemp.EOF
        If rsTemp("传输模式") = 2 Then
            lbl下载.Caption = "最近下载日期：" & Format(rsTemp("日期"), "yyyy-MM-dd hh:mm:ss")
            lbl下载.Tag = Format(rsTemp("日期"), "yyyy-MM-dd hh:mm:ss")
        Else
            lbl上传.Caption = "最近上传日期：" & Format(rsTemp("日期"), "yyyy-MM-dd")
            lbl上传.Tag = Format(rsTemp("日期"), "yyyy-MM-dd")
        End If
        rsTemp.MoveNext
    Loop
    mblnLoad = True
    pgb.Value = 0
End Sub

Private Sub cmd恢复_Click()
    Dim datMax As Date
    
    If IsDate(lbl上传.Tag) = False Then
        MsgBox "尚未进行过数据上传。", vbInformation, gstrSysName
        Exit Sub
    End If
    datMax = CDate(lbl上传.Tag)
    mdat开始日期 = datMax
    mdat结束日期 = datMax
    If frm重复上传日期.GetTimeScope(mdat开始日期, mdat结束日期, datMax) = False Then
        Exit Sub
    End If
    mdat开始日期 = mdat开始日期 - 1  '为了处理开始那天的数据
    
    SetEnable False
   
    DoEvents
    Call 上传数据(True)
    '在上传数据中可能有提交的出错事务，强制回滚
    On Error Resume Next
    gcnOracle.RollbackTrans
    gcn医保.RollbackTrans
    
    Call Form_Load
    
    SetEnable True
    MsgBox "恢复性上传操作处理完成。", vbInformation, gstrSysName
   
End Sub

Private Sub cmd上传_Click()
    SetEnable False
    
    Call 上传数据
    '在上传数据中可能有提交的出错事务，强制回滚
    On Error Resume Next
    gcnOracle.RollbackTrans
    gcn医保.RollbackTrans
    
    Call Form_Load
    
    SetEnable True
    
    If gintType <> 0 Then
        Unload Me
    Else
        MsgBox "上传操作处理完成。", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmd下载_Click()
    SetEnable False
    
    Call 下载数据
    Call Form_Load
    
    SetEnable True
    
    If gintType <> 0 Then
        Unload Me
    Else
        MsgBox "下载操作处理完成。", vbInformation, gstrSysName
    End If
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    cmd下载.Enabled = blnEnable
    cmd上传.Enabled = blnEnable
    cmd恢复.Enabled = blnEnable
    
    If blnEnable = False Then
        MousePointer = vbHourglass
    Else
        MousePointer = vbDefault
    End If
End Sub

Private Sub 下载数据()
    Dim rsHost As New ADODB.Recordset
    Dim varHost As Variant

    On Error GoTo errHandle
    If Get医院参数 = False Then Exit Sub
    
    '由于不同的中心服务器不同，网络要分别拨号，所以分别连接
    rsHost.Open "select * from 保险主机 where 险类=" & TYPE_铜仁市 & " And 编码 = '" & Mid(tabHost.SelectedItem.Key, 2) & "'", gcn医保, adOpenStatic, adLockReadOnly

    '由于可能有多个医保中心，因此做一个循环处理
    Do Until rsHost.EOF
        '获得参数
        If Get主机参数(rsHost) = False Then Exit Sub
        
        '下载核心数据包
        If DownHost(rsHost("编码"), varHost) = False Then
            MsgBox mstr主机名称 & "不能完成数据下载。其对应中心代码在下载文件中没有找到。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Get中心列表(rsHost("编码"))
        
        If Is离线装钱(rsHost("编码")) = True Then
            '仅设置为离线装钱才处理
            If mlng装钱序号 = 0 Or Val(varHost(hf装钱序号)) > mlng装钱序号 Then
                '完成装钱清单的下载
                If Down装钱(varHost(hf装钱序号), varHost(hf医保年)) = False Then
                    Exit Sub
                End If
            End If
        End If
        
        If mlng离休干部序号 = 0 Or Val(varHost(hf离休干部下载序号)) > mlng离休干部序号 Then
            '完成装钱清单的下载
            If Down离休(varHost(hf离休干部下载序号)) = False Then
                Exit Sub
            End If
        End If
        
        '完成单位信息、护理限额数据的下载（每次都下载）
        If Down单位信息 = False Then Exit Sub
        If Down护理限额 = False Then Exit Sub
        
        If mlng补充人员下载序号 = 0 Or Val(varHost(hf补充人员下载序号)) > mlng补充人员下载序号 Then
            '完成装钱清单的下载
            If Down补充(varHost(hf补充人员下载序号)) = False Then
                Exit Sub
            End If
        End If

        If mlng黑名单下载序号 = 0 Or Val(varHost(hf黑名单下载序号)) > mlng黑名单下载序号 Then
            '完成黑名单的下载
            If Down黑名单(varHost(hf黑名单下载序号), varHost(hf医保年)) = False Then
                Exit Sub
            End If
        End If
        
        If mlng病种下载序号 = 0 Or Val(varHost(hf病种下载序号)) > mlng病种下载序号 Then
            '完成疾病清单的下载
            If Down病种(varHost(hf病种下载序号)) = False Then
                Exit Sub
            End If
        End If
        
        If mlng政策下载序号 = 0 Or Val(varHost(hf政策下载序号)) > mlng政策下载序号 Then
            '完成项目的下载
            If Down政策(varHost(hf政策下载序号), varHost(hf医保年)) = False Then
                Exit Sub
            End If
        End If
        
        If mlng项目下载序号 = 0 Or Val(varHost(hf项目下载序号)) > mlng项目下载序号 Then
            '完成项目的下载
            If Down项目(varHost(hf项目下载序号)) = False Then
                Exit Sub
            End If
        End If

        '记录下载日志
        gstrSQL = "insert into 上传下载 (操作日期,用户名,传输模式,中心代码,文件名) " & _
                  "values(sysdate,substr(user,1,20),'2','" & rsHost("编码") & "','Center.pak')"
        gcn医保.Execute gstrSQL

        rsHost.MoveNext
    Loop

    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Function DownHost(主机代码 As String, var序号 As Variant) As Boolean
'功能：下载与主机相关的信息
'参数：主机代码   当前下载的主机代码
'      var序号    其它文件是否需要下载的判断序号
'返回：能成功处理，返回True
    Dim objText As TextStream, strLine As String
    Dim str节 As String '表示当前是处理下载文件的哪一段
    Dim varTemp As Variant, rsTemp As New ADODB.Recordset
    Dim col代理 As New Collection, lng序号 As Long, bln处理 As Boolean
    Dim str开始日期 As String, str终止日期 As String
    
    '首先下载Center.pak包，得到Center文件
    If DownLoadFile("Center.pak") = False Then
        Exit Function
    End If
    
    Set mobjFileSys = New FileSystemObject
    Set objText = mobjFileSys.OpenTextFile(mstr本地下载目录 & "Center")
    gcn医保.BeginTrans '使用事务控制的目的是要将主机参数的新增与删除放在一起，否则可能出法没有主机参数的情况
    On Error GoTo errHandle
    
    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        If strLine <> "" Then
            If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
                '段标志
                str节 = UCase(Mid(strLine, 2, Len(strLine) - 2))
            Else
                varTemp = Split(strLine, "|")
                Select Case str节
                    Case "HOSTS"
                        If varTemp(hf主机代码) = 主机代码 Then
                            '当前主机，直接更新一些没有其它关联的参数
                            gstrSQL = "Update 保险主机 Set 有效期=" & GetDateForOracle(varTemp(hf有效期)) & _
                                    ",是否可用=" & varTemp(hf是否可用) & "  Where 险类=" & TYPE_铜仁市 & " and 编码='" & 主机代码 & "'"
                            gcn医保.Execute gstrSQL
                            '同时，设置返回值
                            bln处理 = True
                            var序号 = Split(strLine, "|")
                            
                            '----------------------------为了保证更新主机参数、中心信息只做一次，所以在此执行
                            '删除上传下载信息，待会在处理[HOSPARAMS]会重建
                            gstrSQL = "Delete 保险主机参数 Where 险类=" & TYPE_铜仁市 & " and 主机='" & 主机代码 & "'"
                            gcn医保.Execute gstrSQL
                            
                            '停止所有中心，然后在处理[CENTERAGENCY]时再启用
                            gstrSQL = "Update 保险中心目录 Set 运行模式=0 Where 险类=" & TYPE_铜仁市
                            gcn医保.Execute gstrSQL
                        Else
                            '属于代理下发，则需要先检查是否有已经代理下发过。如果已经存在，则不处理
                            gstrSQL = "Select 名称 From 保险主机 Where 险类=" & TYPE_铜仁市 & " and 编码='" & varTemp(hf主机代码) & "'"
                            Call OpenRecordset(rsTemp)
                            If rsTemp.RecordCount > 0 Then
                                '已经代理过
                                col代理.Add True, "K" & varTemp(hf主机代码)
                            Else
                                '建立该主机的信息
                                col代理.Add False, "K" & varTemp(hf主机代码)
                                gstrSQL = "Insert Into 保险主机 (险类,编码,名称,是否可用,有效期) VALUES (" & _
                                    TYPE_铜仁市 & ",'" & varTemp(hf主机代码) & "','" & varTemp(hf主机名称) & "'," & varTemp(hf是否可用) & _
                                    "," & GetDateForOracle(varTemp(hf有效期)) & ")"
                                gcn医保.Execute gstrSQL
                            End If
                        End If
                    Case "HOSTPARAMS"
                        str开始日期 = varTemp(hp起始日期)
                        If IsDate(str开始日期) = True Then
                            str开始日期 = "To_date('" & Format(CDate(str开始日期), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                        Else
                            str开始日期 = "To_date('2003-01-01','yyyy-MM-dd')"
                        End If
                        str终止日期 = varTemp(hp终止日期)
                        If IsDate(str开始日期) = True Then
                            str终止日期 = "To_date('" & Format(CDate(str终止日期), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                        Else
                            str终止日期 = "To_date('3000-01-01','yyyy-MM-dd')"
                        End If
                        If varTemp(hp主机代码) = 主机代码 Then
                            '当前主机，直接更新一些没有其它关联的参数
                            gstrSQL = "Insert Into 保险主机参数 (险类,主机,起始日期,终止日期,上传IP,上传用户,上传密码,上传目录,下载IP,下载用户,下载密码,下载目录,装钱IP地址1,装钱IP地址2) " & _
                                      " VALUES (" & TYPE_铜仁市 & ",'" & 主机代码 & "'," & str开始日期 & "," & str终止日期 & _
                                      ",'" & varTemp(hp上传IP) & "','" & varTemp(hp上传用户) & "','" & varTemp(hp上传密码) & "','" & varTemp(hp上传目录) & _
                                          "','" & varTemp(hp下载IP) & "','" & varTemp(hp下载用户) & "','" & varTemp(hp下载密码) & "','" & varTemp(hp下载目录) & _
                                          "','" & varTemp(hp装钱IP地址1) & "','" & varTemp(hp装钱IP地址2) & "')"
                            gcn医保.Execute gstrSQL
                        Else
                            '属于代理下发，则需要先检查是否有已经代理下发过。如果已经存在，则不处理
                            If col代理("K" & varTemp(hp主机代码)) = False Then
                                gstrSQL = "Insert Into 保险主机参数 (险类,主机,起始日期,终止日期,上传IP,上传用户,上传密码,上传目录,下载IP,下载用户,下载密码,下载目录,装钱IP地址1,装钱IP地址2) " & _
                                          " VALUES (" & TYPE_铜仁市 & ",'" & varTemp(hp主机代码) & "'," & str开始日期 & "," & str终止日期 & _
                                          ",'" & varTemp(hp上传IP) & "','" & varTemp(hp上传用户) & "','" & varTemp(hp上传密码) & "','" & varTemp(hp上传目录) & _
                                          "','" & varTemp(hp下载IP) & "','" & varTemp(hp下载用户) & "','" & varTemp(hp下载密码) & "','" & varTemp(hp下载目录) & _
                                          "','" & varTemp(hp装钱IP地址1) & "','" & varTemp(hp装钱IP地址2) & "')"
                                gcn医保.Execute gstrSQL
                            End If
                        End If
                    Case "CENTERS"
                        '首先检查该中心是否存在
                        gstrSQL = "Select Rowid RID From 保险中心目录 Where  险类=" & TYPE_铜仁市 & " and 编码='" & varTemp(cf中心代码) & "'"
                        Call OpenRecordset(rsTemp)
                        If rsTemp.RecordCount > 0 Then
                            '已经存在，只需要更新
                            gstrSQL = "Update 保险中心目录 Set 主机编码='" & varTemp(cf主机代码) & "',个人账户可支付首先自付=" & varTemp(cf个人账户可支付首先自付) & _
                                        ",可写卡冲票=" & varTemp(cf可写卡冲票) & " where RowID='" & rsTemp("RID") & "'"
                            gcn医保.Execute gstrSQL
                        Else
                            '新增该中心（需要取得最大序号）
                            lng序号 = GetMax("保险中心目录", "序号", 1, " Where 险类=" & TYPE_铜仁市)
                            gstrSQL = "Insert Into 保险中心目录 (险类,序号,编码,名称,主机编码,运行模式,个人账户可支付首先自付,可写卡冲票) values(" & _
                                TYPE_铜仁市 & "," & lng序号 & ",'" & varTemp(cf中心代码) & "','" & varTemp(cf中心名称) & "','" & varTemp(cf主机代码) & _
                                "',0," & varTemp(cf个人账户可支付首先自付) & "," & varTemp(cf可写卡冲票) & ")"
                            gcn医保.Execute gstrSQL
                            
                            '同时，向HIS中插入
                            gstrSQL = "Insert Into 保险中心目录 (险类,序号,编码,名称) values(" & _
                                TYPE_铜仁市 & "," & lng序号 & ",'" & varTemp(cf中心代码) & "','" & varTemp(cf中心名称) & "')"
                            gcnOracle.Execute gstrSQL
                        End If
                    Case "HOSTAGENCY"
                        If varTemp(0) = 主机代码 And varTemp(1) = mstr医院编码 Then
                            gstrSQL = "Update 保险主机 Set 装钱模式=" & varTemp(2) & " Where 险类=" & TYPE_铜仁市 & " and 编码='" & 主机代码 & "'"
                            gcn医保.Execute gstrSQL
                        End If
                    Case "CENTERAGENCY"
                        If varTemp(1) = mstr医院编码 Then
                            '除了设置床位费限价外，还允许使用该主机
                            gstrSQL = "Update 保险中心目录 Set 运行模式=1,每天床位费限价=" & varTemp(2) & " Where 险类=" & TYPE_铜仁市 & " and 编码='" & varTemp(0) & "'"
                            gcn医保.Execute gstrSQL
                            gstrSQL = "Update 保险主机 Set 是否可用=1 Where 险类=" & TYPE_铜仁市 & " and 编码='" & 主机代码 & "'"
                            gcn医保.Execute gstrSQL
                        End If
                End Select
            End If
        End If
    Loop
    
    objText.Close  '关闭文件，否则无法得到解压下一个Center文件
    gcn医保.CommitTrans
    DownHost = bln处理
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    objText.Close
    gcn医保.RollbackTrans
End Function

Private Function Down装钱(ByVal lng序号 As Long, ByVal str医保年 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant
    
    '下载表
    If DownLoadFile("inmoneylist.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcn医保.BeginTrans
    On Error GoTo errHandle

    '首先删除当前医保中心的内容
    lbl项目.Caption = "装钱清单"
    gcn医保.Execute "Delete from 装钱清单 where 中心代码 IN (" & mstr中心InOracle & ")"

    '打开数据文件
    Call OpenText(mstr本地下载目录 & "inmoneylist", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr中心InStr, "," & varFields(0)) > 0 Then
                '插入装钱清单,注意此处对金额进行了加密
                gstrSQL = "insert into 装钱清单 (中心代码,卡号,医保年,装钱期次,帐户注入,划拨月份) values ('" & varFields(0) & _
                    "','" & varFields(1) & "','" & str医保年 & "'," & lng序号 & ",'" & EncryptStr(varFields(2), "256", True) & "','" & varFields(3) & "')"
                gcn医保.Execute gstrSQL
            End If
        End If
    Loop

    '更新参数表
    Call Update主机参数("装钱序号", lng序号)

    gcn医保.CommitTrans
    Down装钱 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
End Function

Private Function Down离休(ByVal lng序号 As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant
    
    '下载表
    If DownLoadFile("levwork.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcn医保.BeginTrans
    On Error GoTo errHandle

    '首先删除当前医保中心的内容
    lbl项目.Caption = "离休人员"
    gcn医保.Execute "Delete from 离休人员 where 中心代码 IN (" & mstr中心InOracle & ")"

    '打开数据文件
    Call OpenText(mstr本地下载目录 & "levwork", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr中心InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into 离休人员 (中心代码,姓名,生日,性别,医保号,身份证号,单位医保号,身份代码,单位性质,是否困难企业) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "','" & varFields(3) & _
                     "','" & varFields(4) & "','" & varFields(5) & "','" & varFields(6) & "','" & varFields(7) & "','" & varFields(8) & "','" & varFields(9) & "')"
                gcn医保.Execute gstrSQL
            End If
        End If
    Loop

    '更新参数表
    Call Update主机参数("离休干部下载序号", lng序号)

    gcn医保.CommitTrans
    Down离休 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
End Function

Private Function Down单位信息() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant

    rsTemp.CursorLocation = adUseClient

    gcn医保.BeginTrans
    On Error GoTo errHandle

    '首先删除当前医保中心的内容
    lbl项目.Caption = "单位信息"
    gcn医保.Execute "Delete from 单位信息 where 中心代码 IN (" & mstr中心InOracle & ")"

    '打开数据文件
    Call OpenText(mstr本地下载目录 & "SPECRETIREPAYPARAMS", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr中心InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into 单位信息 (中心代码,性质,单位负担比例,统筹负担比例,个人负担比例,是否困难企业) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "','" & varFields(3) & _
                     "','" & varFields(4) & "','" & varFields(5) & "')"
                gcn医保.Execute gstrSQL
            End If
        End If
    Loop

    gcn医保.CommitTrans
    Down单位信息 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
End Function

Private Function Down护理限额() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant

    rsTemp.CursorLocation = adUseClient

    gcn医保.BeginTrans
    On Error GoTo errHandle

    '首先删除当前医保中心的内容
    lbl项目.Caption = "单位信息"
    gcn医保.Execute "Delete from 护理限额 where 中心单位 IN (" & mstr中心InOracle & ")"

    '打开数据文件
    Call OpenText(mstr本地下载目录 & "TENDLEVY", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr中心InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into 护理限额 (中心单位,级别,费用) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "')"
                gcn医保.Execute gstrSQL
            End If
        End If
    Loop

    gcn医保.CommitTrans
    Down护理限额 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
End Function

Private Function Down补充(ByVal lng序号 As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant
    
    '下载表
    If DownLoadFile("experlist.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcn医保.BeginTrans
    On Error GoTo errHandle

    '首先删除当前医保中心的内容
    lbl项目.Caption = "补充人员"
    gcn医保.Execute "Delete from 补充人员 where 中心代码 IN (" & mstr中心InOracle & ")"

    '打开数据文件
    Call OpenText(mstr本地下载目录 & "experlist", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr中心InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into 补充人员 (中心代码,职工编码) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "')"
                gcn医保.Execute gstrSQL
            End If
        End If
    Loop

    '更新参数表
    Call Update主机参数("补充人员下载序号", lng序号)

    gcn医保.CommitTrans
    Down补充 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
End Function

Private Function Down黑名单(ByVal lng序号 As Long, ByVal str医保年 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long
    
    '下载表
    If DownLoadFile("cardblklist.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcn医保.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("cardblklist", "unitblklist")
    For lngCount = LBound(varTables) To UBound(varTables)
        '首先删除当前医保中心的内容
        If varTables(lngCount) = "cardblklist" Then
            lbl项目.Caption = "黑名单人员"
            '黑名单的处理稍稍有一点不同
            If str医保年 > mstr原医保年 Then
                '跨医保年的处理
                gcn医保.Execute "Delete from 黑名单 where 医保年='" & mstr原医保年 & "' and 灰度<>'1' " & _
                                " and 中心代码 IN (" & mstr中心InOracle & ")"
            End If
            gcn医保.Execute "Delete from 黑名单 where 医保年='" & str医保年 & "' And 中心代码 IN (" & mstr中心InOracle & ")"
        Else
            lbl项目.Caption = "黑名单单位"
            gcn医保.Execute "Delete from 单位黑名单 where 中心代码 IN (" & mstr中心InOracle & ")"
        End If
            
        '打开数据文件
        Call OpenText(mstr本地下载目录 & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                If InStr(mstr中心InStr, "," & varFields(0)) > 0 Then
                    If varTables(lngCount) = "cardblklist" Then
                        gstrSQL = "insert into 黑名单 (中心代码,卡号,灰度,医保年) values ('" & varFields(0) & _
                            "','" & varFields(1) & "','" & varFields(2) & "','" & str医保年 & "')"
                    Else
                        gstrSQL = "insert into 单位黑名单 (中心代码,编码,名称,灰度) values ('" & varFields(0) & _
                            "','" & varFields(1) & "','" & varFields(2) & "','" & varFields(3) & "')"
                    End If
                    gcn医保.Execute gstrSQL
                End If
            End If
        Loop
    Next
    
    '更新参数表
    Call Update主机参数("黑名单下载序号", lng序号)
    Call Update主机参数("医保年", str医保年)

    gcn医保.CommitTrans
    Down黑名单 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
End Function

Private Function Down病种(ByVal lng序号 As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long, lng类别 As Long
    
    '下载表
    If DownLoadFile("sickdefine.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnOracle.BeginTrans
    gcn医保.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("sickdefine", "sickkind")
    For lngCount = LBound(varTables) To UBound(varTables)
        '首先删除当前医保中心的内容
        If varTables(lngCount) = "sickdefine" Then
            lbl项目.Caption = "保险病种"
            gcn医保.Execute "Delete from 保险病种 where 险类=" & TYPE_铜仁市
        Else
            lbl项目.Caption = "病种类型"
            gcn医保.Execute "Delete from 保险病种支付 where 中心代码 IN (" & mstr中心InOracle & ")"
        End If
            
        '打开数据文件
        Call OpenText(mstr本地下载目录 & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                If varTables(lngCount) = "sickdefine" Then
                    '保险病种
                    gstrSQL = "insert into 保险病种 (险类,编码,名称,简码,类别) values (" & _
                                TYPE_铜仁市 & ",'" & varFields(0) & "','" & _
                                Replace(varFields(3), "'", "''") & "','" & varFields(2) & "','" & varFields(4) & "')"
                    gcn医保.Execute gstrSQL
                    
                    '同时更新HIS的病种，注意HIS的病种类型只支持3种，所以将1-5慢病,6-9都归为特殊病
                    gstrSQL = "select rowid as RID FROM 保险病种 where 险类=" & TYPE_铜仁市 & " and 编码='" & varFields(0) & "'"
                    If Val(varFields(4)) >= 6 Then
                        lng类别 = 2
                    ElseIf Val(varFields(4)) >= 1 Then
                        lng类别 = 1
                    Else
                        lng类别 = 0
                    End If
                    If rsTemp.State = adStateOpen Then rsTemp.Close
                    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                    If rsTemp.RecordCount = 0 Then
                        '该病种不存在，新增
                        gstrSQL = "insert into 保险病种 (险类,ID,编码,名称,简码,类别) values (" & _
                                    TYPE_铜仁市 & ",保险病种_ID.nextval,'" & varFields(0) & "','" & _
                                    Replace(varFields(3), "'", "''") & "','" & varFields(2) & "','" & lng类别 & "')"
                    Else
                        '病种已经存在，修改
                        gstrSQL = "update 保险病种  set 名称='" & Replace(varFields(3), "'", "''") & "',简码='" & varFields(2) & _
                            "',类别='" & lng类别 & "' where  rowid='" & rsTemp("RID") & "'"
                    End If
                    gcnOracle.Execute gstrSQL
                Else
                    '病种类型
                    If InStr(mstr中心InStr, "," & varFields(0)) > 0 Then
                        gstrSQL = "INSERT INTO 保险病种支付 (中心代码,病种类型代码,病种类型名称,支付比例,限额,限额类型,累计基本保险支付,累计基本保险费用,门诊报销影响限额,备案时间影响限额,住院天数影响限额,报销月份影响限额,住院影响额度,起付线金额,个人帐户使用方法,起付线算报销,个人帐户算报销,统筹封顶影响报销)  Values('" & _
                            varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "'," & varFields(3) & _
                            "," & varFields(4) & "," & varFields(5) & "," & varFields(6) & "," & varFields(7) & _
                            "," & varFields(8) & "," & varFields(9) & "," & varFields(10) & "," & varFields(11) & _
                            "," & varFields(12) & "," & varFields(13) & "," & varFields(14) & "," & varFields(15) & _
                            "," & varFields(16) & "," & varFields(17) & ") "
                        gcn医保.Execute gstrSQL
                    End If
                End If
            End If
        Loop
    Next
    
    '更新参数表
    Call Update主机参数("病种下载序号", lng序号)

    gcnOracle.CommitTrans
    gcn医保.CommitTrans
    Down病种 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
    gcnOracle.RollbackTrans
End Function

Private Function Down政策(ByVal lng序号 As Long, ByVal str医保年 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long
    Dim lng中心 As Long, lng职工身份 As Long
    Dim cnServer(1 To 2) As ADODB.Connection, lngServer As Long
    
    Set cnServer(1) = gcnOracle
    Set cnServer(2) = gcn医保
    
    Dim rs起付线 As New ADODB.Recordset
    
    rs起付线.Fields.Append "中心", adBigInt, 19, adFldIsNullable
    rs起付线.Fields.Append "在职", adBigInt, 19, adFldIsNullable
    rs起付线.Fields.Append "起付线", adSingle, 15, adFldIsNullable
    rs起付线.CursorLocation = adUseClient
    rs起付线.LockType = adLockOptimistic
    rs起付线.CursorType = adOpenStatic
    rs起付线.Open
    
    '下载表
    If DownLoadFile("policy.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnOracle.BeginTrans
    gcn医保.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("payparams", "paypolicy", "subpayparams", "medikind", "conformation")
    For lngCount = LBound(varTables) To UBound(varTables)
        '首先删除当前医保中心的内容
        Select Case varTables(lngCount)
            Case "paypolicy"
                lbl项目.Caption = "保险政策"
            Case "payparams"
                lbl项目.Caption = "保险参数"
            Case "subpayparams"
                lbl项目.Caption = "补助政策"
                gcn医保.Execute "Delete from 保险补助比例 where 险类=" & TYPE_铜仁市 & " And 年度=" & str医保年 & _
                                " And 中心 IN (" & mstr序号InOracle & ")"
            Case "medikind"
                lbl项目.Caption = "保险大类"
                gcn医保.Execute "Delete from 保险支付大类 where 险类=" & TYPE_铜仁市
            Case "conformation"
                lbl项目.Caption = "剂型"
                gcn医保.Execute "Delete from 剂型"
        End Select
            
        '打开数据文件
        Call OpenText(mstr本地下载目录 & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                Select Case varTables(lngCount)
                    Case "payparams"
                        '基本分段报销
                        'Modified by ZYB 20060617
                        '***只下载本医院等级的政策
                        If varFields(par医院等级) = Mid(mstr医院级别, 1, 1) And varFields(par单据类型) = 3 Then '本处只处理第一位
                            For lngServer = 1 To 2
                                If rsTemp.State = adStateOpen Then rsTemp.Close
                                rsTemp.Open "select 序号 from 保险中心目录 where 险类=" & TYPE_铜仁市 & " and 编码='" & varFields(par中心代码) & "'", cnServer(lngServer), adOpenStatic, adLockReadOnly
                                If rsTemp.RecordCount = 1 Then
                                    '存在该中心
                                    lng中心 = rsTemp("序号")
                                    'Modified by ZYB 20060617
                                    '1*=在职;2*=退休;其它=离休
                                    '下发的文件中,0表示在职,1表示退休,没有下发离休的
                                    lng职工身份 = Switch(Left(varFields(par职工身份), 1) = "0", "1", Left(varFields(par职工身份), 1) = "1", 2, True, 3)
                                    
                                    '保存起付线，以供处理医保政策使用
                                    If lngServer = 2 Then '只处理医保服务器
                                        rs起付线.AddNew
                                        rs起付线("中心") = lng中心
                                        rs起付线("在职") = lng职工身份
                                        rs起付线("起付线") = Val(varFields(par起付线))
                                        rs起付线.Update
                                    End If
                                     
                                     '虽然费用档都是差不多的，但也每次都处理
                                    strLine = "1;第一档;0;" & Format(Val(varFields(par第二段起始值)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "2;第二档;" & Format(Val(varFields(par第二段起始值)), "########0.00;-########0.00; ; ") & ";" & Format(Val(varFields(par第三段起始值)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "3;第三档;" & Format(Val(varFields(par第三段起始值)), "########0.00;-########0.00; ; ") & ";" & Format(Val(varFields(par第四段起始值)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "4;第四档;" & Format(Val(varFields(par第四段起始值)), "########0.00;-########0.00; ; ") & ";" & Format(Val(varFields(par第五段起始值)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "5;第五档;" & Format(Val(varFields(par第五段起始值)), "########0.00;-########0.00; ; ") & ";0;"
                                    gstrSQL = "zl_保险费用档_Update(" & TYPE_铜仁市 & "," & lng中心 & ",'" & strLine & "')"
                                    cnServer(lngServer).Execute gstrSQL, , adCmdStoredProc
                                    
                                    '年龄段
                                    If lng职工身份 = 3 Then
                                        gstrSQL = "zl_保险年龄段_Update(" & TYPE_铜仁市 & "," & lng中心 & ",3,1,1,1,'1;离休;0;0;')"
                                    Else
                                        gstrSQL = "zl_保险年龄段_Update(" & TYPE_铜仁市 & "," & lng中心 & "," & lng职工身份 & ",0,0,0,'1;" & IIf(lng职工身份 = 1, "在职", "退休") & ";0;0;')"
                                    End If
                                    cnServer(lngServer).Execute gstrSQL, , adCmdStoredProc
                                    
                                    '保险支付比例
                                    gstrSQL = "Delete 保险支付比例 WHERE 险类=" & TYPE_铜仁市 & " AND 中心=" & lng中心 & " AND 年度=" & str医保年 & " and 在职=" & lng职工身份
                                    With cnServer(lngServer)
                                        .Execute gstrSQL
                                        .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                            TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & "," & lng职工身份 & ",1,1," & Val(varFields(par第一段报销比例)) * 100 & ")"
                                        .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                            TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & "," & lng职工身份 & ",1,2," & Val(varFields(par第二段报销比例)) * 100 & ")"
                                        .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                            TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & "," & lng职工身份 & ",1,3," & Val(varFields(par第三段报销比例)) * 100 & ")"
                                        .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                            TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & "," & lng职工身份 & ",1,4," & Val(varFields(par第四段报销比例)) * 100 & ")"
                                        .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                            TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & "," & lng职工身份 & ",1,5," & Val(varFields(par第五段报销比例)) * 100 & ")"
                                
                                        If lng职工身份 = 1 Then
                                            '强行处理离休病人（担心政策文件中没有离休病人的下载）
                                             gstrSQL = "zl_保险年龄段_Update(" & TYPE_铜仁市 & "," & lng中心 & ",3,1,1,1,'1;离休;0;0;')"
                                             .Execute gstrSQL
                                             '保险支付比例
                                             gstrSQL = "Delete 保险支付比例 WHERE 险类=" & TYPE_铜仁市 & " AND 中心=" & lng中心 & " AND 年度=" & str医保年 & " and 在职=3"
                                             .Execute gstrSQL
                                             .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                                     TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & ",3,1,1,100)"
                                             .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                                     TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & ",3,1,2,100)"
                                             .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                                     TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & ",3,1,3,100)"
                                             .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                                     TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & ",3,1,4,100)"
                                             .Execute "INSERT INTO 保险支付比例(险类,中心,年度,在职,年龄段,档次,比例) values(" & _
                                                     TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & ",3,1,5,100)"
                                        
                                        End If
                                    End With
                                End If
                            Next
                        End If
                    Case "paypolicy"
                        '报销政策
                        Dim cur起付线 As Double, cur实际起付线 As Double, cur减起付线 As Double
                        Dim lng次数限制 As Long, lng住院次数 As Long
                        
                        '构成保险支付限额处理串
                        strLine = ""
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "select 序号 from 保险中心目录 where 险类=" & TYPE_铜仁市 & " and 编码='" & varFields(pol中心代码) & "'", gcn医保, adOpenStatic, adLockReadOnly
                        If rsTemp.RecordCount = 1 Then
                            '存在该中心
                            lng中心 = rsTemp("序号")
                            
                            rs起付线.Filter = "中心=" & lng中心
                            Do Until rs起付线.EOF
                                cur起付线 = rs起付线("起付线")
                                strLine = strLine & rs起付线("在职") & ";" & "A;" & varFields(pol统筹封顶线) & ";" '统筹封顶线
                                '假设可以住院6次
                                For lng住院次数 = 0 To 5
                                    '首先得到有效的住院次数
                                    If lng住院次数 > (Val(varFields(pol次数限制)) - 1) And Val(varFields(pol次数限制)) > 0 Then
                                        '********采用次数限制
                                        '最多只能算这几次
                                        lng次数限制 = Val(varFields(pol次数限制)) - 1
                                    Else
                                        '次数限制可能为-1
                                        '第一次住院，该值为0
                                        lng次数限制 = lng住院次数
                                    End If
                                    
                                    If varFields(pol算法) = "-" Then
                                        '递减算法
                                        cur减起付线 = Val(varFields(pol数值)) * lng次数限制
                                    Else
                                        '按比例减少
                                        cur减起付线 = cur起付线 * Val(varFields(pol数值)) * lng次数限制
                                    End If
                                    
                                    If cur减起付线 > Val(varFields(pol减起付线限制)) And Val(varFields(pol减起付线限制)) > 0 Then
                                        '********采用减起付线限制
                                        '减起付线限制可能为-1
                                        cur减起付线 = Val(varFields(pol减起付线限制))
                                    End If
                                    
                                    cur实际起付线 = cur起付线 - cur减起付线
                                    
                                    If cur实际起付线 < Val(varFields(pol起付线限制)) And Val(varFields(pol起付线限制)) > 0 Then
                                        '********采用起付线限制
                                        '起付线限制可能为-1
                                        cur实际起付线 = Val(varFields(pol起付线限制))
                                    End If
                                    
                                    strLine = strLine & rs起付线("在职") & ";" & (lng住院次数 + 1) & ";" & cur实际起付线 & ";"
                                Next
                                rs起付线.MoveNext
                            Loop
                            gstrSQL = "zl_保险支付限额_Update(" & TYPE_铜仁市 & "," & lng中心 & "," & str医保年 & ",'" & strLine & "')"
                            gcn医保.Execute gstrSQL, , adCmdStoredProc
                            
                            For lngServer = 1 To 2
                                gstrSQL = "Delete 保险费用档 Where 险类=" & TYPE_铜仁市 & " and 中心=" & lng中心 & " And 档次>" & varFields(pol统筹段数)
                                cnServer(lngServer).Execute gstrSQL
                                
                                gstrSQL = "Update 保险费用档 Set 上限=0 Where 险类=" & TYPE_铜仁市 & " and 中心=" & lng中心 & " And 档次=" & varFields(pol统筹段数)
                                cnServer(lngServer).Execute gstrSQL
                                
'                                gstrSQL = "Delete 保险支付比例 Where 险类=" & TYPE_铜仁市 & " and 中心=" & lng中心 & " And 年度=" & str医保年 & " And 档次>" & varFields(pol统筹段数)
'                                cnServer(lngServer).Execute gstrSQL, , adCmdStoredProc
'
                            Next
                        End If
                        gstrSQL = "Update 保险中心目录 Set " & _
                                  "  起付线在段中=" & varFields(pol起付线在段中) & "," & _
                                  "  段值类型=" & varFields(pol段值类型) & "," & _
                                  "  封顶类型=" & varFields(pol封顶类型) & "," & _
                                  "  使用累计报销=" & varFields(pol使用累计报销) & "," & _
                                  "  当年慢性病月份=" & varFields(pol当年慢性病月份) & "," & _
                                  "  开展补充保险报销=" & varFields(pol开展补充保险报销) & "," & _
                                  "  补充报销比例=" & varFields(pol补充报销比例) & "," & _
                                  "  补充报销限额=" & varFields(pol补充报销限额) & "," & _
                                  "  补充报销限额类型=" & varFields(pol补充报销限额类型) & "," & _
                                  "  补充报销减起付金=" & varFields(pol补充报销减起付金) & "," & _
                                  "  开展补助报销=" & varFields(pol开展补助报销) & "," & _
                                  "  开展慢病报销=" & varFields(pol开展慢病报销) & "," & _
                                  "  开展大病报销=" & varFields(pol开展大病报销) & "," & _
                                  "  跨年起付金类型=" & varFields(pol跨年起付金类型) & "," & _
                                  "  跨年增加住院次数=" & varFields(pol跨年增加住院次数) & "," & _
                                  "  乙类项目价格=" & Val(varFields(pol乙类项目价格)) & "" & _
                                  " Where 险类=" & TYPE_铜仁市 & " And 编码='" & varFields(0) & "'"
                        gcn医保.Execute gstrSQL
                    Case "subpayparams"
                        '补助报销比例
                        gstrSQL = "insert into 保险补助比例(险类,中心,年度,段值,比例)" & _
                                  " SELECT 险类,序号," & str医保年 & " AS 年度," & varFields(1) & "," & varFields(2) & _
                                  " FROM 保险中心目录 WHERE 险类=" & TYPE_铜仁市 & " And 编码='" & varFields(0) & "'"
                        gcn医保.Execute gstrSQL
                    Case "medikind"
                        '医保大类
                        gstrSQL = "insert into 保险支付大类 (险类,编码,名称,非医保编码) values (" & _
                                    TYPE_铜仁市 & ",'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "')"
                        gcn医保.Execute gstrSQL
                        
                        '同时更新HIS的病种，注意HIS的病种类型只支持3种，所以将2-9都归为特殊病
                        gstrSQL = "select rowid as RID FROM 保险支付大类 where 险类=" & TYPE_铜仁市 & " and 编码='" & varFields(0) & "'"
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                        If rsTemp.RecordCount = 0 Then
                            '该保险支付大类不存在，新增
                            gstrSQL = "insert into 保险支付大类 (险类,ID,编码,名称,简码,性质,算法,统筹比额,是否医保,服务对象) values (" & _
                                        TYPE_铜仁市 & ",保险支付大类_ID.nextval,'" & varFields(0) & "','" & _
                                        varFields(1) & "','',1,1,0,1,3)"
                            gcnOracle.Execute gstrSQL
                        End If
                    Case "conformation"
                        '剂型
                        gstrSQL = "INSERT INTO 剂型 (编码,名称) VALUES ('" & _
                                varFields(0) & "','" & Replace(varFields(1), "'", "") & "')"
                        gcn医保.Execute gstrSQL
                End Select
            End If
        Loop
    Next
    
    '更新参数表
    Call Update主机参数("政策下载序号", lng序号)

    gcnOracle.CommitTrans
    gcn医保.CommitTrans
    Down政策 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
    gcnOracle.RollbackTrans
End Function

Private Function Down项目(ByVal lng序号 As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long
    Dim lng是否医保 As Long, str大类编码 As String, rs保险大类 As New ADODB.Recordset
    
    '下载表
    If DownLoadFile("itemcenter.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnOracle.BeginTrans
    gcn医保.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("itemcenter", "agencyspecitem", "sickspecitem")
    For lngCount = LBound(varTables) To UBound(varTables)
        '首先删除当前医保中心的内容
        Select Case varTables(lngCount)
            Case "itemcenter"
                lbl项目.Caption = "保险项目"
                gcn医保.Execute "Delete from 保险项目 where 险类=" & TYPE_铜仁市
'                gcnOracle.Execute "Delete from 保险项目 where 险类=" & TYPE_铜仁市

                '得到非医保大类的编码
                rs保险大类.Open "Select 编码,非医保编码 From 保险支付大类", gcn医保, adOpenStatic, adLockReadOnly
            Case "agencyspecitem"
                lbl项目.Caption = "医院特殊项目"
            Case "sickspecitem"
                lbl项目.Caption = "病种特准项目"
                gcn医保.Execute "Delete from 保险病种项目"
        End Select
            
        '打开数据文件
        Call OpenText(mstr本地下载目录 & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                Select Case varTables(lngCount)
                    Case "itemcenter"
                        '保险项目
                            '根据医保级别判断该项目是否医保。下载的等级大于医院等级，说明该院不满足条件，不能作为医保项目
                            lng是否医保 = IIf(varFields(if是否医保) > Mid(mstr医院级别, 1, 1), 0, 1)
                            str大类编码 = varFields(if大类编码)
                            If lng是否医保 = 0 Then
                                '如果是非医保项目，则换一个大类
                                rs保险大类.Filter = "编码='" & str大类编码 & "'"
                                If rs保险大类.EOF = False Then
                                    str大类编码 = NVL(rs保险大类("非医保编码"), str大类编码)
                                End If
                            End If
                            
                            gstrSQL = "INSERT INTO 保险项目 (险类,编码,名称,简码,单位,剂型编码,大类编码,是否是药,是否医保,最大价格限制," & _
                                      "首先自付比例,价格,项目内涵,除外内容,说明,省级限价,市级限价,县级限价,乡级限价,特检项目,特检自付比例) VALUES ( " & _
                                      TYPE_铜仁市 & ",'" & varFields(if项目序号) & "','" & Replace(varFields(if药典名称), "'", "") & "','" & Replace(varFields(if拼音编码), "'", "") & _
                                      "','" & varFields(if单位) & "','" & varFields(if剂型编码) & "','" & str大类编码 & "','" & varFields(if是否是药) & _
                                      "','" & lng是否医保 & "','" & varFields(if最大价格限制) & "','" & varFields(if首先自付比例) & "','" & varFields(if价格) & _
                                      "','" & varFields(if项目内涵) & "','" & varFields(if除外内容) & "','" & varFields(if说明) & _
                                      "','" & varFields(if省级限价) & "','" & varFields(if市级限价) & "','" & varFields(if县级限价) & "','" & varFields(if乡级限价) & "','" & varFields(if特检项目) & "','" & varFields(if特检自付比例) & "')"
                            gcn医保.Execute gstrSQL
                            
'                            gstrSQL = "INSERT INTO 保险项目 (险类,编码,名称,简码,大类编码 VALUES ( " & _
'                                            TYPE_铜仁市 & ",'" & varFields(if项目序号) & "','" & varFields(if药典名称) & "','" & MidUni(varFields(if拼音编码), 1, 10) & _
'                                            "','" & varFields(if大类编码) & "')"
'                            gcnOracle.Execute gstrSQL
                            
                            '更新现在保险支付项目的是否医保、附注
                            gstrSQL = "update 保险支付项目 A " & _
                                      "  set A.项目名称='" & varFields(if药典名称) & "',A.是否医保=" & lng是否医保 & _
                                      "  where A.项目编码='" & varFields(if项目序号) & "' and A.险类=" & TYPE_铜仁市
                            
                            gcnOracle.Execute gstrSQL
                    Case "agencyspecitem"
                        '医院特殊项目
                        If varFields(0) = mstr医院编码 Then
                            lng是否医保 = IIf(varFields(2) = 1, 0, 1)
                            
                            gstrSQL = "update 保险项目 A " & _
                                      "  set A.首先自付比例=" & varFields(2) & ",A.是否医保=" & lng是否医保 & _
                                      "  where A.项目编码='" & varFields(1) & "' and A.险类=" & TYPE_铜仁市
                            gcn医保.Execute gstrSQL
                            
                            gstrSQL = "update 保险支付项目 A " & _
                                      "  set A.是否医保=" & lng是否医保 & _
                                      "  where A.项目编码='" & varFields(if项目序号) & "' and A.险类=" & TYPE_铜仁市
                            gcnOracle.Execute gstrSQL
                        End If
                    Case "sickspecitem"
                        '病种类型
                        gstrSQL = "INSERT INTO 保险病种项目 (病种序号,项目序号,首先自付比例) VALUES ('" & _
                                varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "')"
                        gcn医保.Execute gstrSQL
                End Select
            End If
        Loop
    Next
    
    '根据医院的等级更新最大价格限价
    Select Case mstr医院级别
    Case "33"
        gstrSQL = "Update 保险项目 Set 最大价格限制=三甲最高限价"
    Case "32"
        gstrSQL = "Update 保险项目 Set 最大价格限制=三乙最高限价"
    Case "23"
        gstrSQL = "Update 保险项目 Set 最大价格限制=二甲最高限价"
    Case "22"
        gstrSQL = "Update 保险项目 Set 最大价格限制=二乙最高限价"
    Case "13", "12"
        gstrSQL = "Update 保险项目 Set 最大价格限制=一级最高限价"
    End Select
    '更新参数表
    Call Update主机参数("项目下载序号", lng序号)

    gcnOracle.CommitTrans
    gcn医保.CommitTrans
    Down项目 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcn医保.RollbackTrans
    gcnOracle.RollbackTrans
End Function

Private Function DownLoadFile(ByVal strFile As String) As Boolean
'功能：下载指定的文件，并且完成解压、解密
    Dim zipfilesIn As ZIPnames
    Dim zipfilesEx As ZIPnames
    Dim lngReturn As Long
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    '下载文件
    lngReturn = FTPDownLoad(mstr下载IP, "21", mstr下载用户, mstr下载密码, mstr远程下载目录, strFile, mstr本地下载目录 & strFile)
    If lngReturn <> 0 Then
        MsgBox "对于“" & mstr主机名称 & "”，文件" & strFile & "下载失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '对文件进行解密
    strTemp = mstr本地下载目录 & Mid(strFile, 1, Len(strFile) - 4) & ".zip"
    DecryptFiles mstr本地下载目录 & strFile, strTemp
    
    '解压文件
    If VBUnzip(strTemp, mstr本地下载目录, 1, 1, 0, 0, 0, 0, zipfilesIn, zipfilesEx) = False Then
        Exit Function
    End If
    
    DownLoadFile = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub 上传数据(Optional ByVal bln恢复 As Boolean = False)
    Dim rsHost As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim str主机编码 As String
    Dim dat结束日期 As Date, dat开始日期 As Date   '查询条件使用   开始日期<=登记日期<结束日期
    Dim datBegin As Date, datEnd As Date           '临时变量，主要用于分段处理
    Dim bln需要上传 As Boolean
    Dim str上传文件  As String
    
    If Get医院参数 = False Then Exit Sub
    rsHost.Open "select * from 保险主机 where 险类=" & TYPE_铜仁市 & " And 编码 = '" & Mid(tabHost.SelectedItem.Key, 2) & "'", gcn医保, adOpenStatic, adLockReadOnly
    
    '缺省是上传今天的数据
    dat结束日期 = CDate(Format(DateAdd("d", 1, Currentdate), "yyyy-MM-dd"))
    If Me.cbo社保局.ListCount <> 0 Then str主机编码 = Mid(Me.cbo社保局.Text, 1, 4)
    
    On Error GoTo errHandle
    '由于可能有多个医保中心，因此做一个循环处理
    Do Until rsHost.EOF
        '获得参数
        Call Get主机参数(rsHost)
        
        '首先完成数据文件的产生
        mstr主机编码 = rsHost("编码")
        mstr主机名称 = rsHost("名称")
        
        Call Get中心列表(rsHost("编码"))
        
        If bln恢复 = False Then
            '1、得到最近一次上传的情况
            gstrSQL = "select max(操作日期) as 上传 from 上传下载 where 传输模式=0 and 中心代码='" & mstr主机编码 & "'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
            
            
            If IsNull(rsTemp("上传")) = True Then
                '未进行过任何处理，用一个相当小的值作为开始
                dat开始日期 = CDate("1900-01-01")
                bln需要上传 = True
            Else
                If rsTemp("上传") < dat结束日期 Then
                    '需要上传今天的数据
                    dat开始日期 = rsTemp("上传")
                    bln需要上传 = True
                Else
                    '今天的工作已经进行，什么也不需要做
                    MsgBox mstr主机名称 & "的数据今天已经上传完成，不用再处理。", vbInformation, gstrSysName
                    bln需要上传 = False
                End If
            End If
        Else
            '进行恢复性上传
            bln需要上传 = True
            dat开始日期 = mdat开始日期 + 1
            dat结束日期 = mdat结束日期 + 1
        End If
        
        If bln需要上传 = True Then
            '首先产生数据
            If dat开始日期 = CDate("1900-01-01") Then
                '从未进行过上传，一次性处理
                datBegin = dat开始日期
                datEnd = dat结束日期
            Else
                datBegin = dat开始日期
                datEnd = dat开始日期 + 1
            End If
            
            Do Until datEnd > dat结束日期
                mstr开始日期 = "to_date('" & Format(datBegin, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                mstr结束日期 = "to_date('" & Format(datEnd, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                mstr日结日期 = "to_date('" & Format(datEnd - 1, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                mstr缺省开始日期 = "to_date('" & Format(DateAdd("d", -15, datEnd), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                
                gcnOracle.BeginTrans
                gcn医保.BeginTrans
                '产生准备上传的数据
                If 日结(bln恢复) = False Then
                    '日结失败
                    gcnOracle.RollbackTrans
                    gcn医保.RollbackTrans
                    Exit Sub
                End If
                
                If bln恢复 = False Then
                    '记录上传日志
                    gstrSQL = "insert into 上传下载 (操作日期,用户名,传输模式,中心代码,文件名) " & _
                              "values(" & mstr结束日期 & ",substr(user,1,20),'0','" & mstr主机编码 & "','" & mstr医院编码 & Format(datEnd - 1, "yyMMdd") & ".pak')"
                    gcn医保.Execute gstrSQL
                End If
                
                '然后合成文件，上传数据
'                If UpLoadFile(mstr主机编码 & mstr医院编码 & Format(datEnd - 1, "yyMMdd") & ".pak") = True Then
                If UpLoadFile(str主机编码 & mstr医院编码 & Format(datEnd - 1, "yyMMdd") & ".pak") = True Then
                    '对当前医保中心的数据进行提交
                    gcnOracle.CommitTrans
                    gcn医保.CommitTrans
                Else
                    gcnOracle.RollbackTrans
                    gcn医保.RollbackTrans
                End If
                '按天分段处理
                datBegin = datBegin + 1
                datEnd = datEnd + 1
            Loop
        End If '可以上传
        rsHost.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Function 日结(Optional ByVal bln恢复 As Boolean = False) As Boolean
'功能：产生门诊的上传数据
    Dim rsTemp As New ADODB.Recordset, str年度 As String
    Dim cur全自费 As Currency, cur首先自付 As Currency, cur统筹 As Currency
    
    On Error GoTo errHandle
    
    '上传之前保存费用的报销部分是正确的
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.收费细目ID,C.项目编码,B.编码,B.名称,A.实收金额 " & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 单价 " & _
              "  From 保险结算记录 D,病人费用记录 A,收费细目 B,保险支付项目 C " & _
              "  where D.性质 = 1 And D.险类 =" & TYPE_铜仁市 & "  And D.记录ID = A.结帐ID And A.登记时间 >=" & mstr开始日期 & " And A.登记时间 <" & mstr结束日期 & _
              "         AND A.实收金额 IS NOT NULL and nvl(A.是否上传,0)=0 And Nvl(A.附加标志,0)<>9 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= " & TYPE_铜仁市 & _
              "  Order by A.病人ID,A.发生时间"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        If Calc费用分割(rsTemp, cur全自费, cur首先自付, cur统筹) = False Then
            Exit Function
        End If
    End If
        
    rsTemp.Close
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.收费细目ID,C.项目编码,B.编码,B.名称,A.实收金额 " & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 单价 " & _
              "  From 病案主页 D,病人费用记录 A,收费细目 B,保险支付项目 C " & _
              "  where D.病人ID =A.病人ID And D.主页ID=A.主页ID And D.险类 =" & TYPE_铜仁市 & " And A.登记时间 >=" & mstr缺省开始日期 & "  And A.登记时间 <" & mstr结束日期 & _
              "        AND A.实收金额 IS NOT NULL and nvl(A.是否上传,0)=0 And Nvl(A.附加标志,0)<>9 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= " & TYPE_铜仁市 & _
              "  Order by A.病人ID,A.发生时间"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        If Calc费用分割(rsTemp, cur全自费, cur首先自付, cur统筹) = False Then
            Exit Function
        End If
    End If
    
    gstrSQL = "SELECT 医保年 FROM 保险主机 WHERE 险类=" & TYPE_铜仁市 & " AND 编码='" & mstr主机编码 & "'"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    str年度 = rsTemp("医保年")
    
    '1、门诊主表
    gstrSQL = "select 中心代码,'" & mstr医院编码 & "' as 医院代码,序号,发票号,姓名,性别,年龄 " & _
              "         ,卡号,医保号,身份证号,单位医保号,身份代码,是否公务员,是否医疗照顾对象,参加补充保险,帐户累计增加,帐户累计支出 " & _
              "         ,统筹已支付金额,统筹已支付费用,慢病已支付金额,慢病已支付费用,慢病起付金已支付,备案日期 " & _
              "         ,门诊个人帐户支付金额,住院个人帐户支付金额,额度已支付金额,部门名称,医生名称,病种代码,病种名称,病种类型 " & _
              "         ,发生费用金额,全自付金额,首先自付金额,个人帐户支付,统筹总支付,统筹总自付,统筹基金支付,统筹基金自付 " & _
              "         ,补充基金支付,补充基金自付,超基本封顶线,超补充封顶线,进入统筹支付,进入统筹费用,进入慢病支付,进入慢病费用 " & _
              "         ,增加住院次数,进入额度支付,进入门诊个人帐户支付,进入慢性病起付金,进入住院个人帐户支付 " & _
              "         ,帐户累计增加,个人帐户支付,支付顺序号,卡灰度级,冲票标志,被冲票据号,票据日期,A.年度," & mstr日结日期 & " as 日结日期 " & _
              "  from 保险结算记录 A " & _
              "  Where A.性质 = 1 And A.险类 =" & TYPE_铜仁市 & " And A.票据日期 <" & mstr结束日期 & _
              IIf(bln恢复, " And A.票据日期 >= " & mstr开始日期, " And Nvl(A.是否上传,0)=0 And A.票据日期 >=" & mstr缺省开始日期) & _
              "       and A.中心代码 in (" & mstr中心InOracle & ")"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UClinicBill", "门诊主表")
    
    '2、门诊明细
    gstrSQL = "select '" & mstr医院编码 & "' as 医院代码,B.No||decode(B.记录状态,2,'2','1') as 序号,substr(E.编码,1,8) as 中心药典编码,trim(substr(E.名称,1,40)) as 中心药典名称 " & _
              "         ,trim(substr(G.名称,1,40)) as 医院药典名称,Round(B.结帐金额/(B.数次*B.付数),4) as 实际价格,B.数次*B.付数*decode(B.记录状态,2,-1,1) as 数量,B.结帐金额*decode(B.记录状态,2,-1,1) " & _
              "         ,B.保险项目否,E.大类编码,E.首先自付比例,E.剂型编码,nvl(substr(decode(Instr(G.规格,'┆'),0,G.规格,substr(G.规格,1,Instr(G.规格,'┆')-1)),1,40),' ') as 规格 " & _
              "  from 保险结算记录 A," & gstrOwner & ".病人费用记录 B," & gstrOwner & ".保险帐户 C,保险项目 E," & gstrOwner & ".收费细目 G " & _
              "  Where A.性质 = 1 And A.险类 =" & TYPE_铜仁市 & " And A.记录ID = B.结帐ID And Nvl(B.附加标志,0)<>9 " & " And A.票据日期 <" & mstr结束日期 & _
              IIf(bln恢复, " And A.票据日期 >= " & mstr开始日期, " And Nvl(A.是否上传,0)=0 And Nvl(B.实收金额,0)<>0 And A.票据日期 >=" & mstr缺省开始日期) & _
              "       and B.保险编码=E.编码 and C.险类=E.险类 and B.收费细目ID=G.ID and A.病人ID=C.病人ID and C.险类=" & TYPE_铜仁市 & " and C.中心 IN (" & mstr序号InOracle & ")"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UClinicBillDetail", "门诊明细")
    
    '3、门诊大类
    gstrSQL = "select '" & mstr医院编码 & "' as 医院代码,B.No||decode(B.记录状态,2,'2','1') as 序号,E.大类编码,sum(B.结帐金额)*decode(B.记录状态,2,-1,1) as 金额 " & _
              "  from 保险结算记录 A," & gstrOwner & ".病人费用记录 B," & gstrOwner & ".保险帐户 C,保险项目 E " & _
              "  Where A.性质 = 1 And A.险类 =" & TYPE_铜仁市 & " And A.记录ID = B.结帐ID And Nvl(B.附加标志,0)<>9 And A.票据日期 <" & mstr结束日期 & _
              IIf(bln恢复, " And A.票据日期 >= " & mstr开始日期, " And Nvl(A.是否上传,0)=0 And A.票据日期 >=" & mstr缺省开始日期) & _
              "       and A.病人ID=C.病人ID and C.险类=" & TYPE_铜仁市 & " and B.保险编码=E.编码 and C.险类=E.险类 " & _
              " and C.中心 IN (" & mstr序号InOracle & ")" & _
              "  group by B.No,B.记录状态,E.大类编码"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UClinicMediKind", "门诊大类")
    
    '4、入院登记
    '如果病人当天入院，又当天出院，接着又去门诊。则可能“保险帐户”的“病种ID”为空。
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    gstrSQL = "select substr(E.编码,1,4) as 中心代码,'" & mstr医院编码 & "' as 医院代码,A.病人ID||'_'||A.主页ID as 序号,D.住院号 " & _
              "         ,substr(D.姓名,1,8) as 姓名,substr(D.性别,1,2) as 性别,floor(MONTHS_BETWEEN(A.入院日期,D.出生日期)/12) as 年龄 " & _
              "         ,substr(C.卡号,1,8) as 卡号,substr(C.医保号,1,8)  as 医保号,D.身份证号,substr(C.单位编码,1,5) as 单位医保号 " & _
              "         ,substr(C.人员身份,1,2) as 人员身份,F.名称,substr(A.登记人,1,8) as 医生,nvl(substr(G.名称,1,50),'无病种')  as 入院病种 " & _
              "         ,trunc(A.入院日期)," & str年度 & " as 年度," & mstr日结日期 & " as 日结日期 " & _
              "  from 病案主页 A,保险帐户 C,病人信息 D,保险中心目录 E,部门表 F,保险病种 G " & _
              "  Where A.险类 =" & TYPE_铜仁市 & " And A.登记时间 <" & mstr结束日期 & " And A.入院科室ID = F.ID And A.入院日期 Is Not Null " & _
              "       and A.病人ID=C.病人ID and C.险类=" & TYPE_铜仁市 & " and A.病人ID=D.病人ID and C.险类=E.险类 and C.中心=E.序号 and C.病种ID=G.ID(+) " & _
              " and E.序号 IN (" & mstr序号InOracle & ")" & IIf(bln恢复, " And A.登记时间 >= " & mstr开始日期, " And A.登记时间 >=" & mstr缺省开始日期 & " and nvl(A.是否上传,0)=0")
    '处理补充登记(相同的记录只取一条)
    gstrSQL = gstrSQL & vbCrLf & " Union " & vbCrLf & _
              "select substr(E.编码,1,4) as 中心代码,'" & mstr医院编码 & "' as 医院代码,A.病人ID||'_'||A.主页ID as 序号,D.住院号 " & _
              "         ,substr(D.姓名,1,8) as 姓名,substr(D.性别,1,2) as 性别,floor(MONTHS_BETWEEN(A.入院日期,D.出生日期)/12) as 年龄 " & _
              "         ,substr(C.卡号,1,8) as 卡号,substr(C.医保号,1,8)  as 医保号,D.身份证号,substr(C.单位编码,1,5) as 单位医保号 " & _
              "         ,substr(C.人员身份,1,2) as 人员身份,F.名称,substr(A.登记人,1,8) as 医生,nvl(substr(G.名称,1,50),'无病种') as 入院病种 " & _
              "         ,trunc(A.入院日期)," & str年度 & " as 年度," & mstr日结日期 & " as 日结日期 " & _
              "  from 病案主页 A,保险帐户 C,病人信息 D,保险中心目录 E,部门表 F,保险病种 G " & _
              "  Where A.险类 =" & TYPE_铜仁市 & " And C.就诊时间 <" & mstr结束日期 & " And A.入院科室ID = F.ID And A.入院日期 Is Not Null " & _
              "       and A.病人ID=C.病人ID and C.险类=" & TYPE_铜仁市 & " and A.病人ID=D.病人ID and C.险类=E.险类 and C.中心=E.序号 and C.病种ID=G.ID(+)  " & _
              "       and E.序号 IN (" & mstr序号InOracle & ")" & IIf(bln恢复, " And trunc(A.登记时间) < " & mstr开始日期 & " And C.就诊时间 >= " & mstr开始日期, " and nvl(A.是否上传,0)=0")
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UInHosRegister", "入院登记")
    gstrSQL = "update 病案主页 A Set A.是否上传 = 1 " & _
              " where A.险类=" & TYPE_铜仁市 & " and A.入院日期 Is Not Null and A.登记时间<" & mstr结束日期 & _
              " and exists (select B.险类 from 保险帐户 B where A.病人ID =B.病人ID And B.险类=A.险类 and B.中心 IN (" & mstr序号InOracle & "))"
    gcnOracle.Execute gstrSQL
    
    '5、记账主表
    '2003-03-03 支掉后面的条件，为了保证合计金额的正确  and B.序号=1 " &
    gstrSQL = "select '" & mstr医院编码 & "' as 医院代码,B.病人ID||'_'||B.主页ID as 入院序号,B.No||decode(B.记录状态,2,'2','1') as 序号 " & _
              "         ,F.名称,substr(B.操作员姓名,1,8) as 医生 " & _
              "         ,sum(B.实收金额)*decode(B.记录状态,2,-1,1) as 金额,decode(b.记录状态,2,-1,1) as 冲票,B.登记时间," & str年度 & " as 年度," & mstr日结日期 & " as 日结日期 " & _
              "  from 病人费用记录 B,保险帐户 C,病案主页 D,保险中心目录 E,部门表 F " & _
              "  where B.记录性质 in (2,3) And Nvl(B.附加标志,0)<>9 and B.登记时间<" & mstr结束日期 & _
              "       and B.病人ID=C.病人ID and C.险类=" & TYPE_铜仁市 & " and B.病人ID=D.病人ID AND B.主页ID=D.主页ID AND D.险类=C.险类 and C.险类=E.险类 and C.中心=E.序号 and B.开单部门ID=F.ID " & _
              "       and E.序号 IN (" & mstr序号InOracle & ")" & IIf(bln恢复, " and B.登记时间>=" & mstr开始日期, " And B.登记时间 >=" & mstr缺省开始日期 & " and Nvl(B.是否上传,0)<>1") & _
              "  group by B.NO,B.病人ID,B.主页ID,F.名称,B.操作员姓名,B.记录状态,B.登记时间"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UInHosBill", "记账主表")
    
    '6、记账明细
    gstrSQL = "select '" & mstr医院编码 & "' as 医院代码,B.病人ID||'_'||B.主页ID as 入院序号,B.No||decode(B.记录状态,2,'2','1') as 序号,substr(E.编码,1,8) as 中心药典编码,trim(substr(E.名称,1,40)) as 中心药典名称 " & _
              "         ,trim(substr(G.名称,1,40)) as 医院药典名称,Round(B.实收金额/(B.数次*B.付数),4) as 实际价格,B.数次*B.付数*decode(B.记录状态,2,-1,1) as 数量 " & _
              "         ,B.实收金额*decode(B.记录状态,2,-1,1) as 金额" & _
              "         ,B.保险项目否,E.大类编码,E.首先自付比例 " & _
              "         ,E.剂型编码,nvl(substr(decode(Instr(G.规格,'┆'),0,G.规格,substr(G.规格,1,Instr(G.规格,'┆')-1)),1,40),' ') as 规格 " & _
              "  from " & gstrOwner & ".病人费用记录 B," & gstrOwner & ".保险帐户 C,保险项目 E," & gstrOwner & ".收费细目 G," & gstrOwner & ".病案主页 H " & _
              "  where B.记录性质 in (2,3) And Nvl(B.附加标志,0)<>9 and B.登记时间<" & mstr结束日期 & _
              "       and B.病人ID=C.病人ID and C.险类=" & TYPE_铜仁市 & " and B.病人ID=H.病人ID AND B.主页ID=H.主页ID AND H.险类=C.险类 " & _
              "       and B.保险编码=E.编码 and E.险类=C.险类 and B.收费细目ID=G.ID And Nvl(B.数次,0)<>0" & _
              "       and C.中心 IN (" & mstr序号InOracle & ")" & IIf(bln恢复, " and B.登记时间>=" & mstr开始日期, " And B.登记时间 >=" & mstr缺省开始日期 & "    AND Nvl(B.是否上传,0)<>1")
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UInHosBillDetail", "记账明细")
        
    '7、住院结算
    gstrSQL = "select 中心代码,'" & mstr医院编码 & "' as 医院代码,序号,发票号,A.病人ID||'_'||A.主页ID as 住院登记号,B.住院号,A.医保号,A.身份证号,单位医保号,A.姓名,A.性别,A.年龄 " & _
              "         ,A.卡号,A.身份代码,A.是否公务员,是否医疗照顾对象,参加补充保险,帐户累计增加,帐户累计支出 " & _
              "         ,统筹已支付金额,统筹已支付费用 " & _
              "         ,住院个人帐户支付金额,A.住院次数,'" & mstr医院级别 & "' as 医院等级,部门名称,医生名称,治愈情况,A.入院日期,A.出院日期,A.住院天数 " & _
              "         ,发生费用金额,全自付金额,首先自付金额,转外首先自付,A.住院次数,起付线,实际起付线,统筹总自付,个人帐户支付,统筹总支付,统筹总自付,统筹基金支付,统筹基金自付 " & _
              "         ,补充基金支付,补充基金自付,补助基金支付,补助基金自付,第一段支付,第一段自付,第二段支付,第二段自付,第三段支付,第三段自付,第四段支付,第四段自付,第五段支付,第五段自付" & _
              "         ,超基本封顶线,超补充封顶线,进入统筹支付,进入统筹费用,进入慢病支付,进入慢病费用 " & _
              "         ,增加住院次数,进入额度支付,进入门诊个人帐户支付,进入慢性病起付金,进入住院个人帐户支付 " & _
              "         ,帐户累计增加,帐户累计支出+个人帐户支付,支付顺序号,卡灰度级,中途结帐,冲票标志,被冲票据号,票据日期,A.年度," & mstr日结日期 & " as 日结日期 " & _
              "  from 保险结算记录 A," & gstrOwner & ".病人信息 B" & _
              "  Where A.性质 = 2 And A.险类 =" & TYPE_铜仁市 & "  And A.票据日期 <" & mstr结束日期 & _
              "       and A.中心代码 in (" & mstr中心InOracle & ") And A.病人ID=B.病人ID" & IIf(bln恢复, "  And A.票据日期 >=" & mstr开始日期, " And A.票据日期 >=" & mstr缺省开始日期 & "  And Nvl(A.是否上传,0)<>1")
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UInHosBalance", "住院结算")
    
    '8、结算医保大类
    gstrSQL = "select '" & mstr医院编码 & "' as 医院代码,A.序号,C.编码,sum(B.结帐金额) as 金额 " & _
              "  from 保险结算记录 A," & gstrOwner & ".病人费用记录 B," & gstrOwner & ".保险支付大类 C " & _
              "  Where A.性质 = 2 And A.险类 = " & TYPE_铜仁市 & " And A.记录ID = B.结帐id+0 And Nvl(B.附加标志,0)<>9 And A.票据日期 <" & mstr结束日期 & _
              "        And B.保险大类ID = C.ID AND A.病人ID=B.病人ID and C.险类=" & TYPE_铜仁市 & IIf(bln恢复, " And A.票据日期 >= " & mstr开始日期, " And A.票据日期 >= " & mstr缺省开始日期 & " And Nvl(A.是否上传,0)<>1 ") & _
              "       and A.中心代码 IN (" & mstr中心InOracle & ")" & _
              "  group by A.序号,C.编码"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UInHosMediKind", "结算医保大类")
    
    gstrSQL = "update 病人费用记录 A Set A.是否上传 = 1 " & _
              " where A.记录性质 in (2,3) and A.登记时间>=" & mstr缺省开始日期 & " and A.登记时间<" & mstr结束日期 & _
              " and exists (select B.险类 from 病案主页 B,保险帐户 C where A.病人ID =B.病人ID and A.主页ID =B.主页ID and B.险类=" & TYPE_铜仁市 & _
              " and B.病人ID=C.病人ID And Nvl(A.是否上传,0)<>1 And B.险类=C.险类 and C.中心 IN (" & mstr序号InOracle & "))"
    gcnOracle.Execute gstrSQL
    
    '9、出院病种
    gstrSQL = "select '" & mstr医院编码 & "' as 医院代码,A.序号 " & _
              "         ,B.编码,B.名称,B.类别,B.名称 " & _
              "  from 保险结算记录 A," & gstrOwner & ".保险帐户 C," & gstrOwner & ".保险病种 B " & _
              "  Where A.险类 =" & TYPE_铜仁市 & " And A.票据日期 <" & mstr结束日期 & IIf(bln恢复, " And A.票据日期 >= " & mstr开始日期, " And A.票据日期 >= " & mstr缺省开始日期 & " and Nvl(A.是否上传,0)<>1") & _
              "        and C.病种ID=B.ID and A.病人ID=C.病人ID And Nvl(A.是否上传,0)<>1 and C.险类=" & TYPE_铜仁市 & " and C.中心 IN (" & mstr序号InOracle & ")"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UInHosSick", "出院病种")
    
    gstrSQL = "update 保险结算记录 A Set A.是否上传 = 1 " & _
              " where Nvl(A.是否上传,0)=0 and 险类=" & TYPE_铜仁市 & " And 票据日期 >=" & mstr缺省开始日期 & " And 票据日期 <" & mstr结束日期
    gcn医保.Execute gstrSQL
    
    '10、项目对应变动记录
    gstrSQL = "select '" & mstr医院编码 & "' as 医院代码,中心药典序号,trim(中心药典名称) 中心药典名称,trim(医院药典名称) 医院药典名称,发生日期 " & _
              "  from 项目对应日志 A " & _
              "  Where A.发生日期 >= " & mstr开始日期 & " And A.发生日期 <" & mstr结束日期
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    Call 产生文件(rsTemp, gstrSQL, "UAssociateItems", "项目关联表")
    
    日结 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub 产生文件(rsData As ADODB.Recordset, ByVal strSource As String, ByVal strFile As String, ByVal str项目 As String)
'根据记录集产生文件
    Dim txtFile As TextStream
    Dim fld As ADODB.Field
    Dim strLine As String, lngLines As Long
    
    
    Set txtFile = mobjFileSys.CreateTextFile(mstr本地上传目录 & strFile)
    
    lbl项目.Caption = str项目
    lngLines = rsData.RecordCount
    DoEvents
    Do Until rsData.EOF
        SetProgress lngLines, rsData.AbsolutePosition
        
        strLine = ""
        For Each fld In rsData.Fields
            If IsNull(fld.Value) Then
                If fld.Type = adNumeric Then
                    strLine = strLine & "0|"
                Else
                    strLine = strLine & "|"
                End If
            Else
                If fld.Type = adDBTimeStamp Then
                    strLine = strLine & Format(fld.Value, "yyyy-MM-dd HH:mm:ss") & "|"
                Else
                    strLine = strLine & fld.Value & "|"
                End If
            End If
        Next
        '结尾仍保留|存在
'        strLine = Mid(strLine, 1, Len(strLine) - 1)
        
        '写入一行记录
        txtFile.WriteLine strLine
        rsData.MoveNext
    Loop
    txtFile.Close
End Sub

Private Function UpLoadFile(ByVal strFile As String) As Boolean
'功能：上传指定的文件，并且完成解压、解密
    Dim zipFile As ZIPnames
    Dim lngCount As Integer, zipname As String
    Dim recurse As Integer, updat As Integer, freshen As Integer, junk As Integer
    
    Dim lngReturn As Long
    Dim strPath As String
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    strPath = mstr本地上传目录
    strPath = strPath & IIf(Right(strPath, 1) <> "\", "\", "")
    
    '首先对文件进行压缩
    junk = 1    ' 1=throw away path names
    recurse = 0 ' 1=recurse -R 2=recurse -r 2=most useful :)
    updat = 0   ' 1=update only if newer
    freshen = 0 ' 1=freshen - overwrite only
    
    zipFile.s(0) = ""
    zipFile.s(1) = strPath & "UClinicBill"
    zipFile.s(2) = strPath & "UClinicBillDetail"
    zipFile.s(3) = strPath & "UClinicMediKind"
    zipFile.s(4) = strPath & "UInHosRegister"
    zipFile.s(5) = strPath & "UInHosBill"
    zipFile.s(6) = strPath & "UInHosBillDetail"
    zipFile.s(7) = strPath & "UInHosBalance"
    zipFile.s(8) = strPath & "UInHosMediKind"
    zipFile.s(9) = strPath & "UInHosSick"
    zipFile.s(10) = strPath & "UAssociateItems"
    lngCount = 11
    zipname = strPath & strFile & ".zip"
    
    If mobjFileSys.FileExists(zipname) = True Then
        mobjFileSys.DeleteFile zipname, True
    End If
    If VBZip(lngCount, zipname, zipFile, junk, recurse, updat, freshen, strPath) = False Then
        Exit Function
    End If
    
    
    '对文件进行加密
    EncryptFiles zipname, strPath & strFile
    
    '上传文件
    lngReturn = FTPUpLoad(mstr上传IP, "21", mstr上传用户, mstr上传密码, strPath & strFile, mstr远程上传目录, strFile)
    If lngReturn <> 0 Then
        MsgBox "对于“" & mstr主机名称 & "”，文件" & strFile & "上传失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    UpLoadFile = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Function Get主机参数(rsHost As ADODB.Recordset) As Boolean
'功能：根据中心序号，得到针对该中心的参数内容
    Dim rsTemp As New ADODB.Recordset
    
    '初始化变量
    mstr主机编码 = NVL(rsHost("编码"))
    mstr主机名称 = NVL(rsHost("名称"))
    
    mstr原医保年 = NVL(rsHost("医保年"))
    mlng装钱序号 = NVL(rsHost("装钱序号"), 0)
    mlng黑名单下载序号 = NVL(rsHost("黑名单下载序号"), 0)
    mlng病种下载序号 = NVL(rsHost("病种下载序号"), 0)
    mlng项目下载序号 = NVL(rsHost("项目下载序号"), 0)
    mlng政策下载序号 = NVL(rsHost("政策下载序号"), 0)
    mlng离休干部序号 = NVL(rsHost("离休干部下载序号"), 0)
    mlng补充人员下载序号 = NVL(rsHost("补充人员下载序号"), 0)
    
    mstr本地下载目录 = NVL(rsHost("本地下载地址"))
    mstr本地上传目录 = NVL(rsHost("本地上传地址"))
    If mstr本地下载目录 = "" Or mstr本地上传目录 = "" Then
        MsgBox "请设置主机“" & rsHost("名称") & "”的本地上传目录和本地下载目录。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr本地下载目录 = mstr本地下载目录 & IIf(Right(mstr本地下载目录, 1) <> "\", "\", "")
    mstr本地上传目录 = mstr本地上传目录 & IIf(Right(mstr本地上传目录, 1) <> "\", "\", "")
    
    '得到远程主机的信息
    gstrSQL = "SELECT B.* FROM 保险主机参数 B " & _
              " Where  B.险类=" & TYPE_铜仁市 & " And B.主机='" & rsHost("编码") & "' " & _
              "    AND nvl(B.起始日期,to_date('2000-01-01','yyyy-MM-dd'))<=SYSDATE  AND nvl(B.终止日期,to_date('3000-01-01','yyyy-MM-dd'))>=trunc(SYSDATE)"
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount <> 1 Then
        MsgBox "主机“" & rsHost("名称") & "”的上传下载参数有错。", vbInformation, gstrSysName
        Exit Function
    End If
    mstr上传IP = NVL(rsTemp("上传IP"))
    mstr上传用户 = NVL(rsTemp("上传用户"))
    mstr上传密码 = NVL(rsTemp("上传密码"))
    mstr下载IP = NVL(rsTemp("下载IP"))
    mstr下载用户 = NVL(rsTemp("下载用户"))
    mstr下载密码 = NVL(rsTemp("下载密码"))
    mstr远程上传目录 = NVL(rsTemp("上传目录"))
    mstr远程下载目录 = NVL(rsTemp("下载目录"))
    
    Get主机参数 = True
End Function

Private Function Get医院参数() As Boolean
'功能：根据医院的相关参数，如医院编码、医院等级
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    '得到医院代码
    gstrSQL = "select 医院编码 from 保险类别 where 序号=" & TYPE_铜仁市
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "请初始化医院的医保编码。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If IsNull(rsTemp("医院编码")) = True Then
        MsgBox "请初始化医院的医保编码。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr医院编码 = Mid(rsTemp("医院编码"), 1, 4)
    
    '得到医院等级
    gstrSQL = "select 参数值 from 保险参数 where 险类=" & TYPE_铜仁市 & " And 参数名='医院级别'"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "请在医保参数中初始化医院的医院级别。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr医院级别 = Mid(rsTemp("参数值"), 1, 2)
    Get医院参数 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub SetProgress(lngSum As Long, lngValue As Long)
'显示进度值
    If lngSum = 0 Then
        pgb.Value = 0
    Else
        pgb.Value = lngValue / lngSum * 100
    End If
End Sub

Private Function Is离线装钱(ByVal 主机代码 As String) As Boolean
'功能：判断当前主机是否是离线装钱
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 装钱模式 From 保险主机 Where 险类=" & TYPE_铜仁市 & " and 编码='" & 主机代码 & "'"
    Call OpenRecordset(rsTemp)
    
    Is离线装钱 = (NVL(rsTemp("装钱模式"), 0) = 2)
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub Get中心列表(ByVal 主机代码 As String)
'功能：判断当前主机是否是离线装钱
    Dim rsTemp As New ADODB.Recordset
    
    mstr中心InOracle = ""
    mstr序号InOracle = ""
    mstr中心InStr = ""
    On Error GoTo errHandle
    
    gstrSQL = "Select 序号,编码 From 保险中心目录 Where 险类=" & TYPE_铜仁市 & " and 主机编码='" & 主机代码 & "'"
    Call OpenRecordset(rsTemp)
    
    Do Until rsTemp.EOF
        mstr序号InOracle = mstr序号InOracle & "," & rsTemp("序号")
        mstr中心InOracle = mstr中心InOracle & ",'" & rsTemp("编码") & "'"
        mstr中心InStr = mstr中心InStr & "," & rsTemp("编码")
        rsTemp.MoveNext
    Loop
    
    If mstr中心InOracle = "" Then
        mstr序号InOracle = "''"
        mstr中心InOracle = "''"
    Else
        mstr序号InOracle = Mid(mstr序号InOracle, 2)
        mstr中心InOracle = Mid(mstr中心InOracle, 2)
    End If
    
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Function GetDateForOracle(ByVal 日期 As String) As String
'功能：根据普通的日期串得到基于Oracle的日期值
    GetDateForOracle = "To_date('" & Format(CDate(AddDate(日期)), "yyyy-MM-dd") & "','yyyy-MM-dd')"
End Function

Private Sub OpenText(ByVal FileName As String, TextFile As TextStream, Lines As Long)
'功能：打开指定文件，并得到其行数
    Set TextFile = mobjFileSys.OpenTextFile(FileName)
    Do While Not TextFile.AtEndOfStream
        TextFile.ReadLine
    Loop
    Lines = TextFile.Line
    Set TextFile = mobjFileSys.OpenTextFile(FileName)

End Sub

Private Sub Update主机参数(ByVal 字段名 As String, ByVal 值 As String)
'功能：更新与保险主机相关的参数
    gstrSQL = "Update 保险主机 Set " & 字段名 & "='" & 值 & "' Where 险类= " & TYPE_铜仁市 & " And 编码='" & mstr主机编码 & "'"
    gcn医保.Execute gstrSQL
End Sub

Private Function Calc费用分割(rs费用明细 As ADODB.Recordset, _
                 cur全自费 As Currency, cur首先自付 As Currency, cur统筹 As Currency) As Boolean
'功能：根据费用明细，重新计算明细中费用的报销金额。计算好的金额可以直接上传
'参数：rs费用明细  费用明细，包含费用的细目ID、单价、数量、金额
'      是否更新     是否需要对数据库中病人费用记录的医保数据进行更新。门诊预算时不能做
'      cur全自费    输出参数，费用中全自费部分的金额
'      cur首先自付  输出参数，费用中首先自付部分的金额
'      cur统筹      输出参数，费用中统筹部分的金额
'      费用分割     输入参数，为否表示限价从病人费用记录中读取，仅计算当前那笔记录
'返回：本函数成功完成所有功能，为True
'调用位置：门诊预算、门诊结算、住院记帐、住院预算、住院结算、费用明细上传

    Dim str中心编码 As String, str病种编码 As String, lng病人ID As Long
    Dim rs保险大类 As New ADODB.Recordset
    Dim rs病种特准 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset, str项目编码 As String, str细目名称 As String
    Dim cur金额 As Currency, cur最大价格 As Currency, cur单价 As Currency, cur自付比例 As Currency, cur床位费 As Currency, cur乙类项目 As Currency
    Dim cur统筹金额 As Currency, cur自付 As Currency, lng保险大类ID As Long, lng保险项目否 As Long
    Dim gcnUpdate As ADODB.Connection
    
    Set gcnUpdate = New ADODB.Connection
    With gcnUpdate
        If .State = 1 Then .Close
        .Open gcnOracle.ConnectionString
        .BeginTrans
    End With
    
    cur全自费 = 0
    cur首先自付 = 0
    cur统筹 = 0
    
    On Error GoTo errHandle
    '得到所有医保大类
    gstrSQL = "SELECT A.ID,A.编码 FROM 保险支付大类 A Where A.险类 =" & TYPE_铜仁市
    rs保险大类.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    'Modified by zyb ##2003-08-31
    If rs费用明细.RecordCount > 0 Then rs费用明细.MoveFirst
    Do Until rs费用明细.EOF
        If lng病人ID <> rs费用明细("病人ID") Then
            lng病人ID = rs费用明细("病人ID")
            '不同的病人，可能属于不同的中心，其床位限价也可能不同，所以要单独处理
            gstrSQL = "SELECT B.编码 中心,C.编码 AS 病种编码 " & _
                "FROM 保险帐户 A,保险中心目录 B,保险病种 C " & _
                "WHERE A.病人ID=" & lng病人ID & " AND A.险类=" & TYPE_铜仁市 & " AND A.险类=B.险类 AND nvl(A.中心,0)=nvl(B.序号,0) AND A.病种ID=C.ID(+)"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
            
            '得到该医保病人的病种特准项目
            gstrSQL = "SELECT A.项目序号,A.首先自付比例 FROM 保险病种项目 A Where A.病种序号 ='" & rsTemp("病种编码") & "'"
            If rs病种特准.State = adStateOpen Then rs病种特准.Close
            rs病种特准.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
            
            '得到该中心规定的床位费限价
            str中心编码 = rsTemp("中心")
            gstrSQL = "Select 每天床位费限价,乙类项目价格 From 保险中心目录 Where 险类=" & TYPE_铜仁市 & " And 编码='" & rsTemp("中心") & "'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
            cur床位费 = rsTemp("每天床位费限价")
            cur乙类项目 = NVL(rsTemp("乙类项目价格"), 0)
        End If
        
        If IsNull(rs费用明细("项目编码")) = True Then
            MsgBox "请为" & rs费用明细("名称") & "设置医保编码。", vbInformation, gstrSysName
            gcnUpdate.RollbackTrans
            Exit Function
        End If
        str项目编码 = rs费用明细("项目编码")
        str细目名称 = rs费用明细("名称")
        
        '获得保险项目的详细信息，方便计算
        gstrSQL = "Select * from 保险项目 Where 险类=" & TYPE_铜仁市 & " And 编码='" & str项目编码 & "'"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
        If rsTemp.EOF Then
            MsgBox str细目名称 & "的保险编码有误，不能完成结算。", vbInformation, gstrSysName
            gcnUpdate.RollbackTrans
            Exit Function
        End If
        
        If rs费用明细("收费类别") = "J" Then
            '床位费
            lng保险项目否 = 1
            If rs费用明细("单价") <= cur床位费 Then
                cur统筹金额 = rs费用明细("实收金额")
            Else
                cur统筹金额 = cur床位费 * rs费用明细("数量")
            End If
            cur统筹 = cur统筹 + cur统筹金额
            cur全自费 = cur全自费 + (rs费用明细("实收金额") - cur统筹金额)
        Else
            '求出该项目的最大可以报销的价格
            cur最大价格 = IIf(NVL(rsTemp("最大价格限制"), 0) = 0, NVL(rsTemp("价格"), 0), rsTemp("最大价格限制"))
            If cur最大价格 > 0 And cur最大价格 < rs费用明细("单价") Then
                '该项目存在最大限价，并且比医院价格要低
                cur单价 = cur最大价格
            Else
                cur单价 = rs费用明细("单价")
            End If
            
            rs病种特准.Filter = "项目序号='" & str项目编码 & "'"
            If rs病种特准.EOF = False Then
                '是否医保项目，按此处作准
                lng保险项目否 = IIf(rs病种特准("首先自付比例") = 1, 0, 1)
                cur自付比例 = rs病种特准("首先自付比例")
            Else
                '以保险项目中的值为准
                lng保险项目否 = rsTemp("是否医保")
                cur自付比例 = rsTemp("首先自付比例")
                
                If lng保险项目否 = 1 And cur乙类项目 > 0 And _
                    (rs费用明细("收费类别") <> "5" And rs费用明细("收费类别") <> "6" And rs费用明细("收费类别") <> "7") Then
                    
                    '对于按价格开区分甲类或乙类项目的中心
                    If rs费用明细("单价") >= cur乙类项目 Then
                        cur自付比例 = 0.2
                    Else
                        cur自付比例 = 0
                    End If
                End If
                
                '虽然定义为保险项目，但由于自付比例，仍改为全自费
                If lng保险项目否 = 1 And rsTemp("首先自付比例") = 1 Then lng保险项目否 = 0
            End If
            
            If lng保险项目否 = 0 Then
                '全自费项目
                cur统筹金额 = 0
                cur全自费 = cur全自费 + rs费用明细("实收金额")
            Else
                If cur最大价格 = 0 Or rs费用明细("单价") <= cur最大价格 Then
                    '没有价格限制，或者限制的价格还没有超过
                    cur统筹金额 = rs费用明细("实收金额") * (1 - cur自付比例)
                Else
                    '有价格限制，就只能取最大价格
                    cur统筹金额 = cur最大价格 * rs费用明细("数量") * (1 - cur自付比例)
                End If
                cur统筹 = cur统筹 + cur统筹金额
                
                'Modified by zyb ##2003-08-31
                '当存在最大价格限制时,其首先自付的计算规则应该是(全自付=超限部分+非医保项目的费用;实收金额=统筹金额+首先自付+全自付)
                If cur最大价格 > 0 And cur最大价格 < rs费用明细("单价") Then
                    cur自付 = (cur最大价格 * rs费用明细("数量") - cur统筹金额)
                Else
                    cur自付 = (rs费用明细("实收金额") - cur统筹金额)
                End If
                cur首先自付 = cur首先自付 + cur自付
                cur全自费 = cur全自费 + (rs费用明细("实收金额") - cur统筹金额 - cur自付)
                'Modified end
            End If
        End If
        
        rs保险大类.Filter = "编码='" & rsTemp("大类编码") & "'"
        If rs保险大类.EOF = False Then
            lng保险大类ID = rs保险大类("ID")
        Else
            lng保险大类ID = 0
        End If
        
        '不做事务控制，这样可以与门诊收费放在一个事务中。然后住院数据都是已经保存好了的，随便怎么计算都无所谓
        'Modified by zyb ##2003-09-01(因为统一改为预结算时全部重算,所以不更新是否上传标志)
        gstrSQL = "ZL_病人费用记录_更新医保(" & rs费用明细("ID") & "," & cur统筹金额 & "," & _
            lng保险大类ID & "," & lng保险项目否 & ",'" & str项目编码 & "',NULL," & cur最大价格 & ")"
        gcnUpdate.Execute gstrSQL, , adCmdStoredProc
        
        rs费用明细.MoveNext
    Loop
    
    gcnUpdate.CommitTrans
    Calc费用分割 = True
    Exit Function
errHandle:
    MsgBox "  费用分割时,发生下列错误:" & vbCrLf & "  " & Err.Description
    gcnUpdate.RollbackTrans
End Function

