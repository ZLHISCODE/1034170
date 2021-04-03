VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillEdit 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "票据领用单"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBillEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraCheck 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4815
      TabIndex        =   31
      Top             =   3435
      Width           =   3030
      Begin VB.OptionButton optResult 
         Caption         =   "不符"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton optResult 
         Caption         =   "相符"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   23
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1335
         Left            =   120
         MaxLength       =   200
         TabIndex        =   26
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "核对备注(&D)"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblResult 
         Caption         =   "核对结果(&C)"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   4530
      TabIndex        =   27
      Top             =   5910
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   5940
      TabIndex        =   28
      Top             =   5910
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -210
      TabIndex        =   30
      Top             =   5790
      Width           =   8295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   120
      TabIndex        =   29
      Top             =   5910
      Width           =   1200
   End
   Begin VB.Frame fraUse 
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   120
      TabIndex        =   32
      Top             =   750
      Width           =   7725
      Begin VB.ComboBox cbo类别 
         Height          =   360
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2565
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   360
         Left            =   4020
         TabIndex        =   35
         Top             =   1155
         Width           =   285
      End
      Begin VB.CommandButton cmd批次 
         Caption         =   "…"
         Height          =   330
         Left            =   2955
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   330
      End
      Begin VB.ComboBox cmb领用人 
         Height          =   360
         Left            =   1380
         TabIndex        =   13
         Text            =   "cmb领用人"
         Top             =   1635
         Width           =   1920
      End
      Begin VB.ComboBox cmb使用方式 
         Height          =   360
         Left            =   5805
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1635
         Width           =   1785
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2115
         Width           =   1920
      End
      Begin VB.ComboBox cmb票种 
         Height          =   360
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   1
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1155
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   2
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1155
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   3
         Left            =   4650
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1155
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   4
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1155
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   5
         Left            =   1395
         MaxLength       =   20
         TabIndex        =   5
         Top             =   705
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   5805
         TabIndex        =   19
         Top             =   2115
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   136183811
         CurrentDate     =   37007
      End
      Begin VB.Label lblUserType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用类别(&K)"
         Height          =   240
         Left            =   3735
         TabIndex        =   2
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "领用人(&G)"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1695
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用方式(&M)"
         Height          =   240
         Index           =   1
         Left            =   4350
         TabIndex        =   14
         Top             =   1695
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人(&R)"
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   2175
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记时间(&D)"
         Height          =   240
         Index           =   3
         Left            =   4350
         TabIndex        =   18
         Top             =   2175
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据种类(&K)"
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   0
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号码范围(&B)"
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   1215
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   240
         Index           =   5
         Left            =   4350
         TabIndex        =   34
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label lbl说明 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   1860
         Left            =   0
         TabIndex        =   20
         Top             =   3015
         Width           =   4605
      End
      Begin VB.Label Label2 
         Caption         =   "详细情况"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   2685
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "批次(&P)"
         Height          =   240
         Index           =   7
         Left            =   480
         TabIndex        =   4
         Top             =   765
         Width           =   840
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票据领用单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2160
      TabIndex        =   21
      Top             =   240
      Width           =   2190
   End
End
Attribute VB_Name = "frmBillEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytInFun As Byte '0-票据领用单,1-核对领用单
Private mblnChange As Boolean     '为真时表示已改变了
Private mdatCurrnet As Date
Private mlng领用ID As Long
Private mstr最小号码 As String
Private mstr最大号码 As String
Private mstr票据长度 As String '表示各种票据的号码长度，各位分别为1-收费,2-预交,3-结帐,4-挂号,5-就诊卡  77777
Private mlng长度 As Long       '当前票据种类的长度
Private mbln药店  As Boolean
Private mstrPrivs As String
Private mrsPerson As ADODB.Recordset
Private mlngPreID As Long
Private mlngModule As Long
Private mbln入库确定领用 As Boolean      '33725
Private mrs报损 As ADODB.Recordset
Private mrs分段 As ADODB.Recordset
Private mstr入库开始号 As String, mstr入库结束号 As String
Private mint上次票据 As Integer
Private mblnNotClick As Boolean
Private mstr类别 As String
Private mintKind As Integer '主界面选择的票种
Private mstrPreType(1 To 5) As String '上次选择的类别
Private mcllCardProperty As Collection  '卡号长度,前缀文本,密文
Private Function Select入库票据(ByVal objCtl As Object, _
    ByVal strKey As String, ByVal int票据 As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的票据
    '入参:objCtl-控件(目前只支持文本框)
    '     strKey-输入的建值
    '     int票据-当前选择的票据
    '出参:
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2010-11-18 11:08:09
    '问题:33725
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strWhere As String
    Dim str开始号码 As String, int前缀 As Integer, lng号码 As Long
    Dim blnCancel As Boolean, sngX As Single, sngY As Single, lngH As Long, i As Long
    Dim vRect As RECT, strSearch1 As String, blnFind As Boolean
    Dim str类别 As String
    
    If Not mbln入库确定领用 Then zlCommFun.PressKey vbKeyTab: Exit Function
    'zlDatabase.ShowSQLSelect
    '功能：多功能选择器
    '参数：
    '     frmMain=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim str使用类别 As String
    strSearch1 = strKey
    Err = 0: On Error GoTo ErrHand:
    If strKey <> "" Then
        If zlCommFun.IsNumOrChar(strKey) Then
            strWhere = " And (A.ID=[3] or A.开始号码 like upper([2]) or A.终止号码 like upper([2]) )"
        Else
            strWhere = " And (A.登记人 like upper([2]) or A.备注 like upper([2]) )"
        End If
        strKey = GetMatchingSting(strKey, False)
    End If
    
    If int票据 = 1 Or int票据 = 3 Then
        strWhere = strWhere & " And nvl(A.使用类别,'LXH')=[4]"
        str类别 = Trim(cbo类别.Text)
        If str类别 = "" Then str类别 = "LXH"
        str使用类别 = " A.使用类别,"
    ElseIf int票据 = 2 Then
       '58071
        strWhere = strWhere & " And nvl(A.使用类别,'0')=[4]"
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        str使用类别 = " decode(nvl(A.使用类别,'0'),'0','','1','门诊','住院') as  使用类别,"
    ElseIf int票据 = 5 Then
        strWhere = strWhere & " And nvl(A.使用类别,'0')=[4]"
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        str使用类别 = " nvl(A.使用类别,'就诊卡') as  使用类别,"
    End If
    
     gstrSQL = "" & _
     "  Select A.Id, A.ID as 入库批次,A.票种,A.使用类别 as 使用类别ID," & str使用类别 & "A.前缀文本,  " & _
     "          A.开始号码, A.终止号码, A.入库数量, A.剩余数量, A.备注, A.登记人, A.登记时间 " & _
     "  From 票据入库记录 A " & IIf(mintKind = 5, ",医疗卡类别 M", "") & _
     "  Where nvl(A.剩余数量,0)>0 And A.票种=[1] And A.有无票据=1  " & strWhere & _
            IIf(mintKind = 5, " And to_number(nvl(A.使用类别,'0'))=M.ID(+) ", "")
    
   '坐标定位
    vRect = GetControlRect(objCtl.hWnd)
    sngX = vRect.Left - 15
    sngY = vRect.Top
    lngH = objCtl.Height
     
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "领用票据选择", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, int票据, strKey, Val(strSearch1), str类别)
    
   If blnCancel = True Then
        Call zl_CtlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "未找到满足条件的票据入库记录,请检查"
        Call zl_CtlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    Call zl_CtlSetFocus(objCtl, True)
    objCtl.Text = rsTemp!ID
    objCtl.Tag = Val(rsTemp!ID)
    
    str开始号码 = Nvl(rsTemp!开始号码)
    txtEdit(1).Text = Trim(Nvl(rsTemp!前缀文本))
    int前缀 = Len(txtEdit(1).Text)
    txtEdit(3).Text = Nvl(rsTemp!前缀文本)
    txtEdit(2).Text = Trim(Mid(str开始号码, int前缀 + 1))
    lng号码 = Len(txtEdit(2).Text)

    txtEdit(2).Tag = txtEdit(2).Text
    txtEdit(4).Text = Mid(Nvl(rsTemp!终止号码), int前缀 + 1)
    txtEdit(4).Tag = txtEdit(4).Text
    blnFind = False
    With cbo类别
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            If int票据 = 2 Or int票据 = 5 Then
              If .ItemData(i) = Val(Nvl(rsTemp!使用类别ID)) Then
                    blnFind = True
                    .ListIndex = i: Exit For
              End If
            Else
                If Trim(.List(i)) = Trim(Nvl(rsTemp!使用类别ID)) Then
                    blnFind = True
                    .ListIndex = i: Exit For
                End If
            End If
        Next
        
        If blnFind = False And Not (int票据 = 2 Or int票据 = 5) Then
            .AddItem Nvl(rsTemp!使用类别ID, " ")
            .ListIndex = .NewIndex
        End If
        .Tag = .Text
        mblnNotClick = False
    End With
    
    mstr入库开始号 = Nvl(rsTemp!开始号码): mstr入库结束号 = Nvl(rsTemp!终止号码):
    Call Load分段票号(Val(objCtl.Text), mstr入库开始号, mstr入库结束号)
    Dim varTemp As Variant
    If mrs分段.RecordCount <> 0 Then
        mrs分段.MoveFirst
        varTemp = Split(Nvl(mrs分段!票据范围) & "-", "-")
        If varTemp(1) = "" Then varTemp(1) = varTemp(0)
        txtEdit(2).Text = Mid(varTemp(0), int前缀 + 1)
        txtEdit(4).Text = Mid(varTemp(1), int前缀 + 1)
    Else
        txtEdit(2).Text = "": txtEdit(4).Text = ""
    End If
    '104238:李南春，2017/2/15，医疗卡要根据领用的票据长度确定号码长度
    If int票据 = 5 Then
        mlng长度 = Len(mstr入库开始号)
        If mlng长度 = 1 Or mlng长度 = 2 Then
            '不能输前缀
            txtEdit(1).Enabled = False
            txtEdit(1).Text = ""
            txtEdit(3).Enabled = False
            txtEdit(3).Text = ""
        End If
        Call txtEdit_Change(1)
    End If
    
    txtEdit(2).Tag = txtEdit(2).Text: txtEdit(4).Tag = txtEdit(4).Text
    zlCommFun.PressKey vbKeyTab
rsTemp.Close
    Select入库票据 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub Load分段票号(ByVal lng批次 As Long, ByVal str入库开始号 As String, ByVal str入库终止号 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载分段票号数据
    '编制:刘兴洪
    '日期:2010-11-18 17:27:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strKey As String, str开始号码 As String, int前缀 As Integer, lng号码 As Long
    
    On Error GoTo errHandle
    Call Init分段票号(mrs分段)
    int前缀 = Len(txtEdit(1).Text): lng号码 = Len(str入库开始号) - int前缀
    '获取当前批次的最大编加和最小编号
    gstrSQL = "" & _
    "   Select 开始号码,nvl(终止号码,开始号码) as 终止号码 From 票据报损记录 Where 入库ID=[1]  " & _
    "   Union ALL " & _
    "   Select 开始号码,nvl(终止号码,开始号码) as 终止号码 From 票据领用记录 Where 批次=[1] and 票种=(Select Max(票种) From 票据入库记录 Where id=[1] ) " & IIf(mlng领用ID <> 0, " And ID<>[2] ", "") & _
    "   Order by 开始号码"
    Set mrs报损 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng批次, mlng领用ID)
    If Not mrs报损.EOF Then
        i = 1
        str开始号码 = str入库开始号
        Do While Not mrs报损.EOF
            If str开始号码 < Nvl(mrs报损!开始号码) Then
                strKey = txtEdit(1).Text & Lpad(zlAddNum(Mid(Nvl(mrs报损!开始号码), int前缀 + 1), False), lng号码, "0")
                If strKey <> str开始号码 Then
                    strKey = str开始号码 & "-" & strKey
                End If
                mrs分段.AddNew
                mrs分段!ID = i
                mrs分段!序号 = i
                mrs分段!票据范围 = strKey
                mrs分段.Update
                i = i + 1
            End If
            str开始号码 = txtEdit(1).Text & Lpad(zlAddNum(Mid(Nvl(mrs报损!终止号码), int前缀 + 1), True), lng号码, "0")
            mrs报损.MoveNext
        Loop
        strKey = str入库终止号
        If str开始号码 <= strKey And str开始号码 <> "" Then
            If str开始号码 <> strKey Then
                strKey = str开始号码 & "-" & strKey
            End If
            mrs分段.AddNew
            mrs分段!ID = i
            mrs分段!序号 = i
            mrs分段!票据范围 = strKey
            mrs分段.Update
        End If
    Else
        mrs分段.AddNew
        mrs分段!ID = 1
        mrs分段!序号 = 1
        mrs分段!票据范围 = str入库开始号 & "-" & str入库终止号
        mrs分段.Update
    End If
 

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Init分段票号(rs分段 As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始分段票号的数据结构
    '编制:刘兴洪
    '日期:2010-11-18 14:33:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rs分段 = New ADODB.Recordset
    With rs分段
        If .State = adStateOpen Then .Close
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "票据范围", adLongVarChar, 200, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub


Private Sub InitContext()
    mbln药店 = (glngSys \ 100 = 8)
    
    mstr最小号码 = ""
    mstr最大号码 = ""
    
    dtpDate.Value = mdatCurrnet
    dtpDate.MaxDate = mdatCurrnet
    
    cmb票种.Clear
    If mbln药店 = False Then
        If InStr(1, mstrPrivs, ";收费收据;") > 0 Then
             cmb票种.AddItem "1-收费收据":        cmb票种.ItemData(cmb票种.NewIndex) = 1
        End If
        
        If InStr(1, mstrPrivs, ";预交收据;") > 0 Or _
          (InStr(1, mstrPrivs, ";预交门诊票据;") > 0 _
          Or InStr(1, mstrPrivs, ";预交住院票据;") > 0) Then
            cmb票种.AddItem "2-预交收据":        cmb票种.ItemData(cmb票种.NewIndex) = 2
        End If
        
        If InStr(1, mstrPrivs, ";结帐收据;") > 0 Then
          cmb票种.AddItem "3-结帐收据":        cmb票种.ItemData(cmb票种.NewIndex) = 3
        End If
        
        If InStr(1, mstrPrivs, ";挂号收据;") > 0 Then
          cmb票种.AddItem "4-挂号收据":        cmb票种.ItemData(cmb票种.NewIndex) = 4
        End If
        
        If InStr(1, mstrPrivs, ";医疗卡;") > 0 Then
           cmb票种.AddItem "5-医疗卡":          cmb票种.ItemData(cmb票种.NewIndex) = 5
        End If
'        cmb票种.AddItem "1-收费收据":        cmb票种.ItemData(cmb票种.NewIndex) = 1
'        cmb票种.AddItem "2-预交收据":        cmb票种.ItemData(cmb票种.NewIndex) = 2
'        cmb票种.AddItem "3-结帐收据":        cmb票种.ItemData(cmb票种.NewIndex) = 3
'        cmb票种.AddItem "4-挂号收据":        cmb票种.ItemData(cmb票种.NewIndex) = 4
'        cmb票种.AddItem "5-医疗卡":          cmb票种.ItemData(cmb票种.NewIndex) = 5
    Else
        cmb票种.AddItem "1-收费收据":        cmb票种.ItemData(cmb票种.NewIndex) = 1
        cmb票种.AddItem "5-会员卡":          cmb票种.ItemData(cmb票种.NewIndex) = 5
    End If
    
    cmb使用方式.Clear
    cmb使用方式.AddItem "1-自用"
    cmb使用方式.AddItem "2-共用"
    cmb使用方式.ListIndex = 0
    
    '初始化票据打印
    'On Error Resume Next
    'BillInit gcnOracle
End Sub
Private Sub cbo类别_Click()
    Dim lng票种 As Long
    Dim blnChange As Boolean
    mblnChange = True
    
    If cmb票种.ListIndex < 0 Then Exit Sub
    lng票种 = cmb票种.ItemData(cmb票种.ListIndex)
    '115348:李南春,2017/10/24,使用类别未改变不刷新界面信息
    If Not (lng票种 = 1 Or lng票种 = 3 Or lng票种 = 2 Or lng票种 = 5) Then Exit Sub
    If lng票种 = 2 Or lng票种 = 5 Then
        blnChange = Val(cbo类别.Tag) <> cbo类别.ItemData(cbo类别.ListIndex)
        cbo类别.Tag = cbo类别.ItemData(cbo类别.ListIndex)
    Else
        blnChange = Trim(cbo类别.Tag) <> Trim(cbo类别.Text)
        cbo类别.Tag = Trim(cbo类别.Text)
    End If
    If blnChange = False Then Exit Sub
    
    If lng票种 = 5 Then
        If cbo类别.ListIndex >= 0 Then
            mlng长度 = mcllCardProperty(cbo类别.ListIndex + 1)(0)
            If mlng长度 = 1 Or mlng长度 = 2 Then
                txtEdit(1).Text = ""
            End If
            Call txtEdit_Change(1)
        End If
    End If
    If mblnNotClick Then GoTo hdYLK
    If mbytInFun = 0 And mlng领用ID = 0 Then
        txtEdit(5).Text = "": txtEdit(1).Text = "": txtEdit(2).Text = "": txtEdit(3).Text = "": txtEdit(4).Text = ""
    End If
hdYLK:
    If lng票种 = 5 Then
        txtEdit(1).Text = UCase(mcllCardProperty(cbo类别.ListIndex + 1)(1))
        txtEdit(1).Enabled = mcllCardProperty(cbo类别.ListIndex + 1)(1) = "" And Not mbln入库确定领用 And mlng长度 > 2: txtEdit(3).Enabled = txtEdit(1).Enabled
        txtEdit(1).BackColor = IIf(txtEdit(1).Enabled, txtEdit(2).BackColor, cmdOK.BackColor)
        txtEdit(3).BackColor = txtEdit(1).BackColor
    End If
    cmdOK.Enabled = True
End Sub

Private Sub cbo类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmb领用人_Validate(Cancel As Boolean)
    If cmb领用人.ListIndex < 0 Then zlControl.CboLocate cmb领用人, mlngPreID, True
    If cmb领用人.ListIndex < 0 And cmb领用人.Text <> "" Then cmb领用人.Text = ""
End Sub
Private Sub cmb票种_Click()
    '选择相应的人员
    Dim rsTmp As New ADODB.Recordset
    Dim int票种 As Integer
    Dim blnChange As Boolean
    
    On Error GoTo errHandle
    '115348:李南春,2017/10/24,票种未改变不刷新界面信息
    If Val(cmb票种.Tag) <> cmb票种.ItemData(cmb票种.ListIndex) Then
        blnChange = True
        cmb票种.Tag = cmb票种.ItemData(cmb票种.ListIndex)
    End If
    If blnChange = False Then Exit Sub
    int票种 = cmb票种.ItemData(cmb票种.ListIndex)
    Call LoadCombox
    If int票种 = 2 Or int票种 = 5 Then
        If cbo类别.ListIndex < 0 Then
            mstrPreType(int票种) = ""
        Else
            mstrPreType(int票种) = cbo类别.ItemData(cbo类别.ListIndex)
        End If
    Else
        mstrPreType(int票种) = cbo类别.Text
    End If
    mblnChange = True
    '得到当前票据种类的长度
'    mlng长度 = Val(Mid(mstr票据长度, cmb票种.ItemData(cmb票种.ListIndex), 1))
'    If mlng长度 = 0 Then
'        mlng长度 = 10
'    End If
    If int票种 = 5 Then
        If cbo类别.ListIndex >= 0 Then
            mlng长度 = mcllCardProperty(cbo类别.ListIndex + 1)(0)
        End If
    Else
        mlng长度 = Val(Split(mstr票据长度, "|")(int票种 - 1))
    End If
    If mlng长度 = 1 Or mlng长度 = 2 Then
        '不能输前缀
        txtEdit(1).Enabled = False
        txtEdit(1).Text = ""
        txtEdit(3).Enabled = False
        txtEdit(3).Text = ""
    Else
        txtEdit(1).Enabled = True
        txtEdit(3).Enabled = True
        If int票种 = 5 Then
            txtEdit(1).Enabled = mcllCardProperty(cbo类别.ListIndex + 1)(1) = ""
            txtEdit(3).Enabled = txtEdit(1).Enabled
        End If
    End If
    Call txtEdit_Change(1)
    
    Select Case int票种
        Case 1      '1-收费收据
            gstrSQL = " And B.人员性质='门诊收费员'"
        Case 2      '2-预交收据
            gstrSQL = " And B.人员性质 in ('预交收款员','入院登记员')"
        Case 3      '3-结帐收据
            gstrSQL = " And B.人员性质='住院结帐员'"
        Case 4      '4-挂号收据
            gstrSQL = " And B.人员性质='门诊挂号员'"
        Case 5      '5-医疗卡 或者称为  会员卡
            gstrSQL = " And B.人员性质 in ('发卡登记人','入院登记员')"
        Case Else
            Exit Sub
    End Select
    gstrSQL = "Select distinct A.ID,A.编号, A.姓名,A.简码 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID " & gstrSQL & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) order by A.姓名"
    Set mrsPerson = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    cmb领用人.Clear
    Do Until mrsPerson.EOF
        cmb领用人.AddItem mrsPerson("姓名")
        cmb领用人.ItemData(cmb领用人.NewIndex) = Val(Nvl(mrsPerson!ID))
        mrsPerson.MoveNext
    Loop
    If cmb领用人.ListCount > 0 Then cmb领用人.ListIndex = 0
    
    With cmb票种
        If mint上次票据 <> .ItemData(.ListIndex) Then
            gstrSQL = "Select 1 From 票据入库记录 where rownum=1 and 票种=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .ItemData(.ListIndex))
            mbln入库确定领用 = Not rsTmp.EOF
            If mbln入库确定领用 Then
                txtEdit(1).Text = "": txtEdit(2).Text = "":
                txtEdit(3).Text = "": txtEdit(4).Text = "":
                txtEdit(5).Text = ""
                If int票种 = 5 Then
                    txtEdit(1).Text = mcllCardProperty(cbo类别.ListIndex + 1)(1)
                    txtEdit(3).Text = mcllCardProperty(cbo类别.ListIndex + 1)(1)
                End If
            End If
            mint上次票据 = .ItemData(.ListIndex)
        End If
    End With
    Call SetCtrlEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmb票种_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
Private Sub cmb领用人_KeyPress(KeyAscii As Integer)
   Dim lngIdx As Long, lng医生ID As Long
     '刘兴洪 问题:27378 日期:2010-01-27 16:20:02
    Dim strAllCaption As String
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cmb领用人.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
       If mrsPerson Is Nothing Then Exit Sub
    If zlPersonSelect(Me, mlngModule, cmb领用人, mrsPerson, cmb领用人.Text, True, strAllCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
 
End Sub

Private Sub cmb领用人_Click()
    If cmb领用人.ListIndex >= 0 Then mlngPreID = cmb领用人.ItemData(cmb领用人.ListIndex)
    mblnChange = True
    cmdOK.Enabled = True
End Sub

Private Sub cmb使用方式_Click()
    mblnChange = True
    cmdOK.Enabled = True
End Sub

Private Sub cmb使用方式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdSel_Click()
    If Select票号 = False Then Exit Sub
    zl_CtlSetFocus cmb领用人
End Sub
Private Sub cmd批次_Click()
    If Select入库票据(txtEdit(5), "", cmb票种.ItemData(cmb票种.ListIndex)) = False Then
        Exit Sub
    End If
End Sub

Private Sub dtpDate_Change()
    mblnChange = True
    cmdOK.Enabled = True
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mint上次票据 = -1
    mbln入库确定领用 = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lng领用ID As Long
    Dim strUserName As String
    
    If ValidateContent() = False Then Exit Sub
    If Val(zlDatabase.GetPara("领用票据签字确认", glngSys, mlngModule, 0)) = 1 Then
        '问题:40775
        strUserName = zlDatabase.UserIdentify(Me, "请输入签字用户名和密码!", glngSys, mlngModule, "")
        If strUserName = "" Then Exit Sub
        If strUserName <> cmb领用人.Text Then
            MsgBox "领用人与签字人不一致,不能继续!", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If Save(lng领用ID, strUserName) = False Then Exit Sub
    
    If mbytInFun = 0 Then
        If mlng领用ID <> 0 Then
            '修改
            mblnChange = False
            Unload Me
            Exit Sub
        Else
            '连续新增
            If Mid(frmBillSupervise.lvwMain.SelectedItem.Key, 2) = cmb票种.ItemData(cmb票种.ListIndex) Then
                '票种相同，更新主界面中
                Dim lst As ListItem
                Set lst = frmBillSupervise.lvw领用_S.ListItems.Add(, "C" & lng领用ID, "开始号码", "Item", "Item")
                frmBillSupervise.ShowItem lst
                If frmBillSupervise.lvw领用_S.ListItems.Count = 1 Then
                    '只有一张领用单
                    lst.Selected = True
                    frmBillSupervise.Fill汇总
                End If
            End If
            txtEdit(2).Text = ""
            txtEdit(4).Text = ""
            
            '问题号:115671,焦博,2017/11/15,领用票据后，确定按钮禁用,使光标停留在票种的下拉框中。
            cmdOK.Enabled = False
            cmb票种.SetFocus
            If mstr入库开始号 <> "" Then
                Call Load分段票号(Val(txtEdit(5).Text), mstr入库开始号, mstr入库结束号)
                If mrs分段.RecordCount <> 0 Then
                    Dim varTemp As Variant
                    mrs分段.MoveFirst
                    varTemp = Split(Nvl(mrs分段!票据范围) & "-", "-")
                    If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                    txtEdit(2).Text = Mid(varTemp(0), Len(txtEdit(1).Text) + 1)
                    txtEdit(4).Text = Mid(varTemp(1), Len(txtEdit(1).Text) + 1)
                Else
                    txtEdit(5).Text = "": zl_CtlSetFocus txtEdit(5)
                End If
            End If
        End If
        mblnChange = False
    Else
        mblnChange = False
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 And txtEdit(1).Text <> txtEdit(3).Text Then txtEdit(3).Text = txtEdit(1).Text
    If Index = 3 And txtEdit(1).Text <> txtEdit(3).Text Then txtEdit(1).Text = txtEdit(3).Text
    If Index = 1 Or Index = 3 Then
         
        txtEdit(2).MaxLength = mlng长度 - LenB(StrConv(txtEdit(1).Text, vbFromUnicode))
        txtEdit(4).MaxLength = txtEdit(2).MaxLength
    End If
    If Index = 5 Then
        txtEdit(Index).Tag = "": Set mrs分段 = Nothing
    End If
    Call ShowSum
    cmdOK.Enabled = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    SelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 1 Or Index = 3 Then
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
    End If
    txtEdit(Index).Text = Trim(txtEdit(Index).Text)
End Sub
Private Function Select票号() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择具体的票据号
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-18 16:24:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '弹出选择器
    Dim rsReturn As ADODB.Recordset, varTemp As Variant
    If Not mbln入库确定领用 Then
        Select票号 = True: Exit Function
    End If
    
    On Error GoTo errHandle
    If mrs分段 Is Nothing Then
        ShowMsgbox "请先确定入库批次,请检查"
        zlCtlSetFocus txtEdit(5)
        Exit Function
    End If
    
    mrs分段.Filter = 0
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtEdit(2), mrs分段, True, "", "ID", rsReturn) Then
        If rsReturn.RecordCount <> 0 Then
            varTemp = Split(rsReturn!票据范围 & "-", "-")
            txtEdit(2).Text = Mid(varTemp(0), Len(txtEdit(1).Text) + 1)
            txtEdit(4).Text = Mid(varTemp(1), Len(txtEdit(3).Text) + 1)
            zlCtlSetFocus cmb领用人
        End If
    End If
    mrs分段.Filter = 0
    
    Select票号 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 5 And mbln入库确定领用 Then
            If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
            If Select入库票据(txtEdit(Index), Trim(txtEdit(Index).Text), cmb票种.ItemData(cmb票种.ListIndex)) = False Then
                Exit Sub
            End If
            Exit Sub
        End If
        If Not (Index = 2 Or Index = 4) Then
            If Trim(txtEdit(Index)) = "" Then
                If Select票号 = False Then Exit Sub
            End If
        End If
        zlCommFun.PressKey vbKeyTab: Exit Sub
    Else
        
    End If
    If Index = 1 Or Index = 3 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
    Else
        If Not (Index = 5 And mbln入库确定领用) Then
            If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Function ValidateContent() As Boolean
'功能:检查输入内容的是否有效
'返回:有效则返回True,否则返回False
    Dim i As Integer, int票种 As Integer, strTemp As String
    Dim str类别 As String
    Dim rsTmp As New ADODB.Recordset
    Dim bln张数过多 As Boolean '问题号:43366
    Dim lng长度 As Long, byt发卡控制 As Byte
    
    On Error GoTo errHandle
    
    ValidateContent = False
    If mbytInFun = 0 Then
        If cmb票种.ListIndex < 0 Then
            MsgBox "请先选择要领用的票种。", vbExclamation, gstrSysName
            If cmb票种.Visible And cmb票种.Enabled Then cmb票种.SetFocus
            Exit Function
        End If
        int票种 = cmb票种.ItemData(cmb票种.ListIndex)
        
        
        '字符串检查
        For i = 1 To 4
            If StrIsValid(txtEdit(i).Text, txtEdit(i).MaxLength) = False Then
                txtEdit(i).SetFocus
                SelAll txtEdit(i)
                Exit Function
            End If
        Next
        
        For i = 1 To Len(txtEdit(2).Text)
            strTemp = Mid(txtEdit(2), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox "开始号码中含有非数字字符，字母只能作为前缀。", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                SelAll txtEdit(2)
                Exit Function
            End If
        Next
        For i = 1 To Len(txtEdit(4).Text)
            strTemp = Mid(txtEdit(4), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox "终止号码中含有非数字字符，字母只能作为前缀。", vbExclamation, gstrSysName
                txtEdit(4).SetFocus
                SelAll txtEdit(4)
                Exit Function
            End If
        Next
        If mbln入库确定领用 Then
            If txtEdit(5).Tag = "" Then
                    MsgBox "入库批次未选择,不能领用。", vbExclamation, gstrSysName
                    zl_CtlSetFocus txtEdit(5)
                    Exit Function
            End If
        End If
        If Len(txtEdit(2).Text) <> txtEdit(2).MaxLength Then
            If Not mbln入库确定领用 And int票种 = 5 Then
                lng长度 = mcllCardProperty(cbo类别.ListIndex + 1)(0)
                byt发卡控制 = mcllCardProperty(cbo类别.ListIndex + 1)(3)
                Select Case byt发卡控制
                    Case 0
                        MsgBox "开始号码的长度不够，应该有" & lng长度 & "位!", vbExclamation, gstrSysName
                        txtEdit(2).SetFocus
                        SelAll txtEdit(2)
                        Exit Function
                    Case 2
                        If MsgBox("开始号码的长度少于" & lng长度 & "位!,是否继续？", vbExclamation + vbYesNo, gstrSysName) = vbNo Then
                            txtEdit(2).SetFocus
                            SelAll txtEdit(2)
                            Exit Function
                        End If
                End Select
            Else
                MsgBox "开始号码的长度不够，应该有" & mlng长度 & "位。", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                SelAll txtEdit(2)
                Exit Function
            End If
        End If
        If Len(txtEdit(2).Text) = 0 Then
            MsgBox "开始号码不能为空。", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            SelAll txtEdit(2)
            Exit Function
        End If
        If Len(txtEdit(2).Text) <> Len(txtEdit(4).Text) Then
            MsgBox "终止号码的长度要和开始号码的相同。", vbExclamation, gstrSysName
            txtEdit(4).SetFocus
            SelAll txtEdit(4)
            Exit Function
        End If
        If txtEdit(2).Text > txtEdit(4).Text Then
            MsgBox "开始号码必须小于终止号码。", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            SelAll txtEdit(2)
            Exit Function
        End If
        If txtEdit(2).Text = "0000000000" And txtEdit(4).Text = "9999999999" Then
            MsgBox "不能使用这个号码范围。", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            SelAll txtEdit(2)
            Exit Function
        End If
        If mstr最小号码 <> "" Then
            If Len(txtEdit(2).Text) <> Len(txtEdit(2).Tag) Then
                MsgBox "这张领用单的票据已经使用，号码长度不能改变。" & vbCrLf & "号码长度应该是" & Len(txtEdit(1).Text & txtEdit(2).Tag) & "位。", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                SelAll txtEdit(2)
                Exit Function
            End If
            If txtEdit(1).Text & txtEdit(2).Text > mstr最小号码 Then
                MsgBox "这张领用单的票据已经使用，" & vbCrLf & "开始号码最大只可以到" & mstr最小号码 & "。", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                SelAll txtEdit(2)
                Exit Function
            End If
            If txtEdit(3).Text & txtEdit(4).Text < mstr最大号码 Then
                MsgBox "这张领用单的票据已经使用，" & vbCrLf & "号码已经用到" & mstr最大号码 & ",终止号码必须大于它。", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                SelAll txtEdit(2)
                Exit Function
            End If
        End If
        If cmb领用人.Text = "" Then
            MsgBox "领用人不能为空。", vbExclamation, gstrSysName
            cmb领用人.SetFocus
            Exit Function
        End If
        
        '问题号:43366,54259
        If Len(CalcTotal) > 11 Then
            bln张数过多 = True
        ElseIf Len(CalcTotal) < 11 Then
            bln张数过多 = False
        ElseIf CalcTotal > "9999999999" Then
            bln张数过多 = True
        ElseIf CalcTotal < "9999999999" Then
            bln张数过多 = False
        End If
        
        '检查号码总张数是否过大
        If bln张数过多 Then
            MsgBox "号码可用的数量异常过大，请检查开始结束号码的正确性。", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            SelAll txtEdit(2)
            Exit Function
        End If
        
        
'        '检查号码范围是否过大
'        If CalcTotal > 999999999# Then
'            MsgBox "号码可用的数量异常过大，请检查开始结束号码的正确性。", vbExclamation, gstrSysName
'            txtEdit(2).SetFocus
'            SelAll txtEdit(2)
'            Exit Function
'        End If
        '检查是否有使用类别
        
        If cbo类别.ListIndex < 0 Then
            MsgBox "注意:" & vbCrLf & IIf(int票种 = 2, "   预交类别", IIf(int票种 = 5, "    医疗卡类别", "    使用类别")) & "没有选择,请选择!", vbInformation + vbOKOnly, gstrSysName
            zl_CtlSetFocus cbo类别: Exit Function
            Exit Function
        End If

        Select Case int票种
            Case 1, 3
                '收费和结帐
                str类别 = Trim(cbo类别.Text)
            Case 2
                '预交
                str类别 = cbo类别.ItemData(cbo类别.ListIndex)
                If Val(str类别) = 0 Then str类别 = ""
            Case 5
                '就诊卡
                str类别 = cbo类别.ItemData(cbo类别.ListIndex)
                If Val(str类别) = 0 Then str类别 = ""
            Case Else
                str类别 = ""
        End Select
        
        '判断领用是否重复
        '问题号:115760,焦博,2017/12/5,相同类别的票据或医疗卡才需要检查重复,不同类别的不需要检查
        gstrSQL = "" & _
        "   Select ID,nvl(剩余数量,0) 剩余数量 " & _
        "   From 票据领用记录 " & _
        "   Where ID<>[3] And 票种=[4] And nvl(使用类别,'LXH') =nvl([5], 'LXH')" & _
        "       And (开始号码<=[1] and 终止号码>=[1] or 开始号码<=[2] and 终止号码>=[2]) And length(开始号码)=length([1]) "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtEdit(1).Text & txtEdit(2).Text, txtEdit(3).Text & txtEdit(4).Text, mlng领用ID, Left(cmb票种.Text, 1), str类别)
        If rsTmp.RecordCount > 0 Then
            If MsgBox("与本次领用票据号有重叠的领用记录存在" & IIf(rsTmp("剩余数量") > 0, "，并且还有未使用完的票据。", "。") & vbCrLf & "你还需要继续吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    Else
        If zlCommFun.ActualLen(txtRemarks.Text) > txtRemarks.MaxLength Then
            MsgBox "备注信息不允许超过" & txtRemarks.MaxLength & "个字符!", vbExclamation, gstrSysName
            If txtRemarks.Enabled Then txtRemarks.SetFocus
            Exit Function
        End If
    End If
    ValidateContent = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Save(ByRef lng领用ID As Long, ByVal strUserName As String) As Boolean
'功能:保存编辑的内容
'参数:lng领用ID-新增时返回新记录的领用ID
'返回值:成功返回True,否则为False
    Dim strTemp As String, strSQL As String, str类别 As String
    Dim int票种 As Integer
    On Error GoTo errHandle
    Save = False
    int票种 = cmb票种.ItemData(cmb票种.ListIndex)
    str类别 = ""
    If int票种 = 1 Or int票种 = 3 Then str类别 = Trim(cbo类别.Text)
    If int票种 = 2 Then
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        If Val(str类别) = 0 Then str类别 = ""
    ElseIf int票种 = 5 Then
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        If Val(str类别) = 0 Then str类别 = ""
    End If
    
    str类别 = IIf(str类别 = "", "NULL", "'" & str类别 & "'")
    
    If mbytInFun = 0 Then
        If mlng领用ID = 0 Then
            '新增
            lng领用ID = GetNextId("票据领用记录")
            strSQL = "zl_票据领用记录_insert(" & lng领用ID & "," & Left(cmb票种.Text, 1) & "," & str类别 & _
                ",'" & cmb领用人.Text & "','" & txtEdit(1).Text & "','" & txtEdit(1).Text & txtEdit(2).Text & _
                "','" & txtEdit(3).Text & txtEdit(4).Text & "'," & Left(cmb使用方式.Text, 1) & _
                ",to_date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),'" & txtEdit(0).Text & _
                "'," & CalcTotal & ",'" & txtEdit(5).Text & "'," & IIf(strUserName = "", "NULL", "'" & strUserName & "'") & ")"
        Else
            '修改
            strSQL = "zl_票据领用记录_update(" & mlng领用ID & "," & str类别 & ",'" & cmb领用人.Text & _
                "','" & txtEdit(1).Text & txtEdit(2).Text & "','" & txtEdit(3).Text & txtEdit(4).Text & _
                "','" & txtEdit(1).Text & "'," & Left(cmb使用方式.Text, 1) & _
                ",to_date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),'" & txtEdit(0).Text & "','" & txtEdit(5).Text & "'," & IIf(strUserName = "", "NULL", "'" & strUserName & "'") & ")"
        End If
    Else
        strSQL = "zl_票据领用记录_check(" & mlng领用ID & "," & IIf(optResult(0).Value, 1, 0) & ",'" & UserInfo.姓名 & "','" & txtRemarks.Text & "',0)"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Save = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    lng领用ID = 0
End Function

Private Sub ShowSum()
'功能:显示汇总信息
    Dim strTemp As String
    strTemp = vbCrLf & "  开始号码：" & Replace(txtEdit(1).Text, "&", "&&") & txtEdit(2).Text & vbCrLf
    strTemp = strTemp & "  终止号码：" & Replace(txtEdit(3).Text, "&", "&&") & txtEdit(4).Text & vbCrLf
    If txtEdit(2).Text = "" Or txtEdit(4).Text = "" Then
        strTemp = strTemp & "  票据总张数：" & vbCrLf
    Else
        strTemp = strTemp & "  票据总张数：" & CalcTotal & vbCrLf
    End If
    If mstr最小号码 <> "" Then
        strTemp = strTemp & "  已经使用的最小号码：" & Replace(mstr最小号码, "&", "&&") & vbCrLf
        strTemp = strTemp & "  已经使用的最大号码：" & Replace(mstr最大号码, "&", "&&") & vbCrLf
    End If
    
    lbl说明.Caption = strTemp
End Sub
Public Function ShowMe(ByVal frmOwner As Form, bytInFun As Byte, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lng领用ID As Long, Optional ByVal str类别 As String = "", Optional intKind As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:用来与调用的财务监控窗口进行通讯的程序,用来增加缴款记录
    '入参:bytInFun:0-领用与修改,1-核对领用单
    '       str类别-缺省的使用类别
    '       intKind-主界面传入的票种
    '出参:
    '返回:编辑成功返回True,否则为False
    '编制:刘兴洪
    '日期:2011-05-05 16:43:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long, blnFind As Boolean, i As Long
    Dim int票种 As Integer
    mlngModule = lngModule: mstrPrivs = strPrivs: mbytInFun = bytInFun
    mstr类别 = str类别: mintKind = intKind
    int票种 = mintKind: mstr入库开始号 = "": mstr入库结束号 = ""
    mstr最大号码 = "": mstr最小号码 = ""
    '42618
    On Error GoTo errHandle
    If mintKind <> 0 Then
        mstrPreType(mintKind) = mstr类别
    End If
    If UserInfo.姓名 = "" Then
        MsgBox "请为你自己指定相应人员，否则不能使用本功能。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mdatCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mlng领用ID = lng领用ID
    Call InitContext
            
    mstr票据长度 = zlDatabase.GetPara(20, glngSys, , , "7|7|7|7|7")
    Set mrs报损 = Nothing
    Set mrs分段 = Nothing
    
    If mlng领用ID = 0 Then
        '新增
        mstr最小号码 = ""
        mstr最大号码 = ""
        With cmb票种
            For i = 0 To .ListCount - 1
                If .ItemData(i) = mintKind Then .ListIndex = i: Exit For
            Next
        End With
        With cbo类别
            For i = 0 To .ListCount - 1
                    If mintKind = 2 Or mintKind = 5 Then
                        If Trim(.ItemData(i)) = Val(mstr类别) Then .ListIndex = i: Exit For
                    Else
                        If Trim(.List(i)) = mstr类别 Then .ListIndex = i: Exit For
                    End If
            Next
        End With
         txtEdit(0).Text = UserInfo.姓名
        
        On Error Resume Next
        cmb领用人.Text = UserInfo.姓名
        If Err <> 0 Then
            If InStr(mstrPrivs, "所有操作员") = 0 Then
                MsgBox "你不具备相应的人员性质，没有权限领用票据。", vbInformation, gstrSysName
                mblnChange = False: Unload Me: Exit Function
            End If
        End If
        If InStr(mstrPrivs, "所有操作员") = 0 Then cmb领用人.Enabled = False
        On Error GoTo errHandle
    Else
        '修改,或核对
        gstrSQL = "" & _
        "   Select A.票种,A.使用类别,A.领用人,A.前缀文本,A.开始号码,A.终止号码,A.使用方式,A.登记时间,A.登记人,A.当前号码,A.剩余数量,A.批次," & _
        "           B.开始号码 as 入库开始号,B.终止号码 as 入库终止号 " & _
        "   From 票据领用记录 A,票据入库记录 B  " & _
        "   Where A.ID=[1] And nvl(A.批次,0)=B.ID(+) and A.票种 =B.票种(+)"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng领用ID)
        
        If rsTmp.RecordCount = 0 Then Exit Function
        
        For lngIndex = 0 To cmb票种.ListCount - 1
            If cmb票种.ItemData(lngIndex) = rsTmp("票种") Then
                '找到该种票据的类型
                cmb票种.ListIndex = lngIndex
                Exit For
            End If
        Next
        int票种 = Val(Nvl(rsTmp("票种")))
        With cbo类别
            mblnNotClick = True
            blnFind = False
            For i = 0 To .ListCount - 1
                If int票种 = 2 Then
                    If .ItemData(i) = Val(Nvl(rsTmp!使用类别)) Then
                       .ListIndex = i: blnFind = True: Exit For
                    End If
                ElseIf int票种 = 5 Then
                    If .ItemData(i) = Val(Nvl(rsTmp!使用类别)) Then
                       .ListIndex = i: blnFind = True: Exit For
                    End If
                Else
                    If .List(i) = Nvl(rsTmp!使用类别) Then
                       .ListIndex = i: blnFind = True: Exit For
                    End If
                End If
            Next
            '58071
            If blnFind = False And int票种 <> 5 Then
                .AddItem Nvl(rsTmp!使用类别, " ")
                .ListIndex = .NewIndex
            End If
            .Enabled = IIf(Nvl(rsTmp!入库开始号) = "", True, False)
            lblUserType.Enabled = .Enabled
            mblnNotClick = False
        End With
        
        mlng长度 = Len(Nvl(rsTmp!开始号码))
        
        cmb票种.Enabled = False
        
        txtEdit(1).Text = IIf(IsNull(rsTmp("前缀文本")), "", rsTmp("前缀文本"))
        txtEdit(2).Text = Mid(rsTmp("开始号码"), Len(txtEdit(1).Text) + 1)
        txtEdit(2).Tag = txtEdit(2).Text
        txtEdit(4).Text = Mid(rsTmp("终止号码"), Len(txtEdit(1).Text) + 1)
        txtEdit(4).Tag = txtEdit(4).Text
        txtEdit(5).Text = "" & rsTmp!批次
        txtEdit(5).Tag = "" & rsTmp!批次
        cmb使用方式.ListIndex = rsTmp("使用方式") - 1
        txtEdit(0).Text = UserInfo.姓名
        dtpDate.Value = rsTmp("登记时间")
        
        On Error Resume Next
        cmb领用人.Text = rsTmp("领用人")
        If Err <> 0 Then
            cmb领用人.AddItem rsTmp("领用人")
            cmb领用人.Text = rsTmp("领用人")
        End If
        If InStr(mstrPrivs, "所有操作员") = 0 Then cmb领用人.Enabled = False
        On Error GoTo errHandle
        If Nvl(rsTmp!入库开始号) <> "" And mbytInFun = 0 Then
            Call Load分段票号(Val(Nvl(rsTmp!批次)), Nvl(rsTmp!入库开始号), Nvl(rsTmp!入库终止号))
        End If
        
        rsTmp.Close
        gstrSQL = "select nvl(min(号码),' ') as 最小号码,nvl(max(号码),' ')  as 最大号码 from 票据使用明细 where 领用ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng领用ID)
        
        mstr最小号码 = Trim(rsTmp("最小号码"))
        mstr最大号码 = Trim(rsTmp("最大号码"))
        If mstr最小号码 <> "" Then
            '号码已经使用，有些内容就不能更改
            txtEdit(1).Enabled = False
            txtEdit(3).Enabled = False
            Call ShowSum
        End If
    End If
    
    mblnChange = False
    Me.Caption = IIf(mbytInFun = 0, "票据领用单", "核对领用单")
    If mbytInFun = 0 Then
        fraCheck.Visible = False
        lbl说明.Width = lbl说明.Width + (cmb使用方式.Left + cmb使用方式.Width - (lbl说明.Left + lbl说明.Width))
    Else
        fraUse.Enabled = False
    End If
    Call SetCtrlEnable
    frmBillEdit.Show vbModal, frmOwner
    ShowMe = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetCtrlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enable属性和Visible属性
    '编制:刘兴洪
    '日期:2010-11-18 10:59:12
    '问题:33725
    '---------------------------------------------------------------------------------------------------------------------------------------------
     If mbytInFun = 0 And mlng领用ID = 0 Then
        cmd批次.Visible = mbln入库确定领用
     Else
        cmd批次.Visible = False
        txtEdit(5).Enabled = txtEdit(5).Enabled And Not mbln入库确定领用    '批次
     End If
    cmdSel.Visible = mbln入库确定领用 And mbytInFun = 0
    txtEdit(1).Enabled = txtEdit(1).Enabled And Not mbln入库确定领用        '开始前缀文本
    txtEdit(3).Enabled = txtEdit(3).Enabled And Not mbln入库确定领用    ''结束前缀文本
    If txtEdit(1).Enabled = False Then
         txtEdit(1).BackColor = cmdOK.BackColor
    Else
         txtEdit(1).BackColor = txtEdit(2).BackColor
    End If
    txtEdit(3).BackColor = txtEdit(1).BackColor
    cmdOK.Enabled = True
End Sub

Private Function CalcTotal() As String
'功能：获取可用号码总数
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    '问题43366
    If InStr(1, txtEdit(4).Text, ".") > 0 Or InStr(1, txtEdit(2).Text, ".") > 0 Then
        ShowMsgbox "号码范围不能输入小数,请重新输入!"
        Exit Function
    End If
    
    '问题:28048:
     CalcTotal = NumberSubtrac(txtEdit(4).Text, txtEdit(2).Text)
    'CalcTotal = CDec(txtEdit(4).Text) - CDec(txtEdit(2).Text) + 1
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function LoadCombox() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载Combox数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-27 10:22:29
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int票种 As Integer, strSQL As String, rsTemp As ADODB.Recordset
    Dim str类别 As String
    If cmb票种.ListIndex < 0 Then Exit Function
    
    On Error GoTo errHandle
    
    int票种 = cmb票种.ItemData(cmb票种.ListIndex)
    str类别 = mstrPreType(int票种)
    Select Case int票种
    Case 1, 3 '1-收费,3-结帐
        strSQL = "Select 编码,名称,简码,缺省标志 From 票据使用类别 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo类别
            .Clear
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = 1
                If Val(Nvl(rsTemp!缺省标志)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            .AddItem " "    '允许设置为空
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            mblnNotClick = False
            .Enabled = True: lblUserType.Enabled = True
        End With
    Case 2 '预交
        mblnNotClick = True
        With cbo类别
            .Clear
            If InStr(1, mstrPrivs, ";预交门诊票据;") > 0 Then
                .AddItem "门诊预交": .ItemData(.NewIndex) = 1
                If Val(str类别) = 1 Then .ListIndex = .NewIndex
            End If
            If InStr(1, mstrPrivs, ";预交住院票据;") > 0 Then
                .AddItem "住院预交": .ItemData(.NewIndex) = 2
                If Val(str类别) = 2 Then .ListIndex = .NewIndex
            End If
            '58071
            If InStr(1, mstrPrivs, ";预交门诊票据;") > 0 And InStr(1, mstrPrivs, ";预交住院票据;") > 0 Then
                .AddItem " "
                .ItemData(.NewIndex) = 0
                If Val(str类别) = 0 Then .ListIndex = .NewIndex
            End If
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            .Enabled = True
        End With
        mblnNotClick = False
    Case 5 '医疗卡
        strSQL = "Select ID,编码,名称,缺省标志,卡号长度,卡号密文,前缀文本,发卡控制 From 医疗卡类别 where nvl(是否启用,0) >=1 Order by 编码 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo类别
            .Clear
            Set mcllCardProperty = New Collection
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
                mcllCardProperty.Add Array(Val(Nvl(rsTemp!卡号长度)), CStr(Nvl(rsTemp!前缀文本)), CStr(Nvl(rsTemp!卡号密文)), Val(Nvl(rsTemp!发卡控制))), "K" & Val(Nvl(rsTemp!ID))
                If Val(Nvl(rsTemp!缺省标志)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If Val(str类别) = Val(Nvl(rsTemp!ID)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            mblnNotClick = False: .Enabled = True
        End With
    Case Else
        cbo类别.Enabled = False: lblUserType.Enabled = False
        cbo类别.ListIndex = -1
    End Select
    LoadCombox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

