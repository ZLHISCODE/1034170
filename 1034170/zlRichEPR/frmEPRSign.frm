VERSION 5.00
Begin VB.Form frmEPRSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "书写签名"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6330
   Icon            =   "frmEPRSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -270
      TabIndex        =   15
      Top             =   1785
      Width           =   6555
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   3630
      TabIndex        =   25
      Top             =   1710
      Width           =   30
   End
   Begin VB.CheckBox chkOrgPic 
      Caption         =   "签名原图"
      Height          =   195
      Left            =   3735
      TabIndex        =   22
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.TextBox txtHeight 
      Height          =   270
      Left            =   4965
      TabIndex        =   20
      Text            =   "50"
      Top             =   2625
      Width           =   390
   End
   Begin VB.CheckBox chkSignPic 
      Caption         =   "签名使用图片"
      Height          =   195
      Left            =   3735
      TabIndex        =   19
      Top             =   1965
      Width           =   1395
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2610
      Width           =   2310
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -270
      TabIndex        =   18
      Top             =   510
      Width           =   6555
   End
   Begin VB.CheckBox chkPreText 
      Caption         =   "将签名级别作为前缀加入(&P)"
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   1950
      Width           =   2565
   End
   Begin VB.CheckBox chkHandSign 
      Caption         =   "显示手签位置(&H)"
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   2257
      Width           =   1695
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "数字签名(&E)"
      Height          =   195
      Left            =   3105
      TabIndex        =   7
      Top             =   1013
      Width           =   1365
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1387
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   1365
   End
   Begin VB.OptionButton optName 
      Caption         =   "指定用户(&U)"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1320
   End
   Begin VB.OptionButton optName 
      Caption         =   "当前用户(&C)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5130
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3885
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   4110
   End
   Begin VB.PictureBox pic签名图片 
      AutoRedraw      =   -1  'True
      Height          =   810
      Left            =   5415
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   21
      Top             =   2325
      Width           =   810
   End
   Begin VB.Label lblH 
      Caption         =   "高"
      Height          =   225
      Left            =   5055
      TabIndex        =   24
      Top             =   2415
      Width           =   180
   End
   Begin VB.Label lblWH 
      Height          =   225
      Left            =   3720
      TabIndex        =   23
      Top             =   3015
      Width           =   1605
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "签名时间(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   2670
      Width           =   990
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   240
      TabIndex        =   17
      Top             =   3255
      Width           =   5970
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "签名效果预览:"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户密码(&P)"
      Height          =   180
      Left            =   510
      TabIndex        =   5
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "张三"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1605
      TabIndex        =   14
      Top             =   660
      Width           =   360
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签名级别(&L)"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmEPRSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '父窗体
Private Sign As cEPRSign                    '签名对象

Private mlngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mblnOk As Boolean
Private mintSign As Integer                 '显示姓名还是显示签名 避免同名情况
Private msSource As String                 '数字签名的源字符串
Private mpicSign  As StdPicture
Private morgSign  As StdPicture             '签名原始图(人员表.签名图片)


'################################################################################################################
'## 功能：  显示本窗体
'##
'## 参数：  edtThis     :IN     编辑器控件
'##         fParent     :IN     父窗体
'##         strSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
'################################################################################################################
Public Function ShowMe(ByRef edtThis As Editor, ByRef fParent As Object, _
                        ByVal sSource As String, ByRef picSign As StdPicture) As cEPRSign
    
    Dim bytFileKind As Byte    '是否护理病历
    bytFileKind = fParent.Document.EPRPatiRecInfo.病历种类
    Set mpicSign = Nothing
    Set morgSign = Nothing
    
    Dim lngStart As Long, strPreText As String
    mintSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    
    Me.cboTime.Clear
    Me.cboTime.AddItem "不显示"
    Me.cboTime.AddItem Format(Now(), "yyyy-MM-dd hh:mm")
    Me.cboTime.AddItem Format(Now(), "yyyy年MM月dd日 hh:mm")
    
    lngStart = edtThis.Selection.StartPos
    strPreText = edtThis.Range(lngStart - 1, lngStart)
    If strPreText = ":" Or strPreText = "：" Then
        Me.chkPreText.Value = vbUnchecked
    Else
        Me.chkPreText.Value = vbChecked
    End If
    
    Set Sign = New cEPRSign
    Set frmParent = fParent
    msSource = sSource
    
    '根据签名级别来初始化“签名级别＂
    Select Case bytFileKind
    Case cpr护理病历
        cmbLevel.AddItem "1 - 护士"
        cmbLevel.AddItem "3 - 护士长"
        cmbLevel.ListIndex = 0
        If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 1
    Case cpr诊疗报告
        cmbLevel.AddItem "1 - 医生"
        cmbLevel.AddItem "2 - 主治"
        cmbLevel.AddItem "3 - 主任"
        cmbLevel.ListIndex = 0
        If frmParent.Document.用户签名级别 >= cprSL_主治 Then cmbLevel.ListIndex = 1
        If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 2
    Case Else
        cmbLevel.AddItem "1 - 经治医师"
        cmbLevel.AddItem "2 - 主治医师"
        cmbLevel.AddItem "3 - 副主任医师"
        cmbLevel.AddItem "4 - 主任医师"
        cmbLevel.ListIndex = 0
        If frmParent.Document.用户签名级别 >= cprSL_主治 Then cmbLevel.ListIndex = 1
        If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 2
        If frmParent.Document.用户签名级别 >= cprSL_正高 Then cmbLevel.ListIndex = 3
    End Select
    
    '读取当前签名方式（系统参数26）
    Dim lS As Long
    Select Case fParent.Document.EPRPatiRecInfo.病历种类
    Case cpr门诊病历
        lS = 1
    Case cpr住院病历
        lS = 2
    Case cpr诊疗报告
        Select Case fParent.Document.EPRFileInfo.lngModule
            Case 1290, 1291, 1294
                lS = 7
            Case Else
                lS = 3
        End Select
        
    Case cpr护理病历
        lS = 4
    Case Else
        Select Case fParent.Document.EPRPatiRecInfo.病人来源
        Case cprPF_门诊
            lS = 1
        Case cprPF_住院
            lS = 2
        Case Else
            lS = 2  '否则，以住院为准
        End Select
    End Select
    
    mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), lS, 1)) '门诊,住院,医技,护理,药品,LIS,PACS (1111111),为空默认采用密码模式
    If mlngPassType = 1 Then
        If gstrESign = "" Or (lS = 3 And gstrESign = "0") Then '医技工作站书写报告没有调用clsDockxx类,如果先刷新"住院病历"页面，再填写报告会在clsDockInEPR中产生gstrESign = "0"
            gstrESign = getPassESign(3, fParent.Document.EPRPatiRecInfo.科室ID)
        End If
        mlngPassType = Val(gstrESign)
    End If
    
    lblUserName.Caption = gstrUserName
    chkEsign.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", vbUnchecked)
    
    txtHeight.Text = zlDatabase.GetPara("签名图片高度", glngSys, 1070, "50", Array(txtHeight), InStr(gstrPrivsEpr, "参数设置") > 0)
    txtHeight.ToolTipText = txtHeight.Text: txtHeight.Tag = txtHeight.Text
    chkOrgPic.Value = zlDatabase.GetPara("签名使用原图", glngSys, 1070, "1", Array(chkOrgPic, lblH), InStr(gstrPrivsEpr, "参数设置") > 0)
    chkOrgPic.Tag = chkOrgPic.Value
    
    chkSignPic.Value = zlDatabase.GetPara("签名使用图片", glngSys, 1070, "0", Array(chkSignPic), InStr(gstrPrivsEpr, "参数设置") > 0)
    chkSignPic.Tag = chkSignPic.Value
    
    chkHandSign.Value = zlDatabase.GetPara("显示手签位置", glngSys, 1070, "0", Array(chkHandSign), InStr(gstrPrivsEpr, "参数设置") > 0)
    chkHandSign.Tag = chkHandSign.Value
    
    chkPreText.Value = zlDatabase.GetPara("将签名级别作为前缀加入", glngSys, 1070, "0", Array(chkPreText), InStr(gstrPrivsEpr, "参数设置") > 0)
    chkPreText.Tag = chkPreText.Value

    cboTime.ListIndex = zlDatabase.GetPara("签名时间", glngSys, 1070, "0", Array(cboTime), InStr(gstrPrivsEpr, "参数设置") > 0)
    cboTime.Tag = cboTime.ListIndex
   
    Call RefControls
    
    Me.Show vbModal, frmParent
    If mblnOk Then
        Set ShowMe = Sign
        If Sign.签名图片 Then
            Set picSign = mpicSign
        Else
            Set picSign = Nothing
        End If
    Else
        Set picSign = Nothing
        Set ShowMe = Nothing
    End If
    Set mpicSign = Nothing
    Set morgSign = Nothing
End Function

'################################################################################################################
'## 功能：  保存签名到内部签名组并刷新显示（验证密码或者数字签名）
'################################################################################################################
Private Function Validation() As Boolean
    Dim blnSpecify As Boolean, strSpecifySign, lngSpecifyId As Long, lngSpecifyLevel As Long
    Dim lngCertID As Long, strSign As String, str时间戳 As String, objSignPic As Object, str时间Base64 As String
    Dim rsTemp As ADODB.Recordset, l As Long
    
    On Error GoTo errHand
    Dim lngPatiId As Long, lngPageId As Long, bFileType As Byte
    lngPatiId = frmParent.Document.EPRPatiRecInfo.病人ID
    lngPageId = frmParent.Document.EPRPatiRecInfo.主页ID
    bFileType = frmParent.Document.EPRPatiRecInfo.病历种类

    If optName(1).Value Then  '指定帐号签名
        blnSpecify = True
        txtName = Trim(txtName)
        txtPass = Trim(txtPass)
        
        If frmParent.Document.EPRPatiRecInfo.病历种类 = cpr住院病历 Or frmParent.Document.EPRPatiRecInfo.病历种类 = cpr门诊病历 Or frmParent.Document.EPRPatiRecInfo.病历种类 = cpr护理病历 Then
            gstrSQL = "Select 1 From 上机人员表 A, 部门人员 B Where a.用户名 = [1] And a.人员id = b.人员id And b.部门id = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查签名用户与当前用户是否同科室", UCase(txtName.Text), frmParent.Document.EPRPatiRecInfo.科室ID)
            If rsTemp.EOF Then
                MsgBox "指定签名用户与当前操作人员不属于同一科室，禁止操作该科室病人病历！", vbExclamation, gstrSysName: Exit Function
            End If
        End If
        
        If chkEsign.Value = vbUnchecked Then '密码签名
            If Trim(txtPass) = "" Then MsgBox "指定帐号密码不能为空，请检查！", vbExclamation: Exit Function
            If Not OraDataOpen(txtName, IIf(UCase(txtName) = "SYS" Or UCase(txtName) = "SYSTEM", txtPass, TranPasswd(txtPass))) Then
                MsgBox "指定帐号/密码错误,请重新输入登录帐号和密码！", vbInformation + vbOKOnly, gstrSysName: Exit Function
            End If
        End If
        
        gstrSQL = "Select ID,姓名,签名 From 人员表 p Where ID=(Select 人员ID From 上机人员表 Where 用户名='" & UCase(txtName) & "')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo")
        If rsTemp.EOF Then MsgBox "指定帐号不存在，请重新输入登录帐号和密码!", vbInformation, gstrSysName: Exit Function
        
        If mintSign = 0 Then
            strSpecifySign = rsTemp!姓名
        Else
            strSpecifySign = NVL(rsTemp!签名, rsTemp!姓名)         '显示签名
        End If
        lngSpecifyId = rsTemp.Fields("ID")   '用户ID
        
        lngSpecifyLevel = GetUserSignLevel(lngSpecifyId, rsTemp!姓名, frmParent.Document.EPRPatiRecInfo.病人ID, frmParent.Document.EPRPatiRecInfo.主页ID) '获取指定用户的签名级别
        If lngSpecifyLevel = cprSL_空白 Then MsgBox "指定帐号尚未设置签名级别，请在人员管理中调整聘任职务！", vbInformation, gstrSysName: Exit Function
        For l = 1 To frmParent.Document.Signs.Count
            If frmParent.Document.Signs(l).签名级别 > lngSpecifyLevel Then
                MsgBox "当前病历已有更高级别的签名,当前签名级别无权审签本病历", vbInformation, gstrSysName: Exit Function
            End If
        Next
    End If
    
    If Not (IIf(blnSpecify, lngSpecifyLevel, frmParent.Document.用户签名级别) >= Val(cmbLevel.Text)) Then '
        MsgBox "用户拥有的签名级别低于选定的签名级别,请重新选定签名级别！", vbInformation, gstrSysName: Exit Function
    End If

    If chkEsign.Value = vbChecked Then '数字签名,在此窗口中对签名对象进行初始化，此窗口关闭后，数据保存，提取数据生成源内容进行签名，若签名对象初始化失败则不保存
        If gobjESign Is Nothing Then
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            If gobjESign.Initialize(gcnOracle, glngSys) = False Then Exit Function
        End If
        
        If gobjESign.CheckCertificate(IIf(blnSpecify, UCase(txtName), gstrDBUser)) = False Then Exit Function
        
        '停用的，只能用密码签名
        If Not gobjESign.CertificateStoped(IIf(blnSpecify, strSpecifySign, gstrUserName)) Then
            strSign = gobjESign.signature(msSource, IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), ""), lngCertID, str时间戳, objSignPic, str时间Base64, False, lngPatiId, IIf(bFileType = cpr门诊病历, 0, lngPageId), IIf(bFileType <> cpr门诊病历, 0, lngPageId)) '返回签名信息,lngCertID返回签名使用的证书记录ID
            If strSign = "" Then MsgBox "数字签名失败！请再次签名！", vbInformation + vbOKOnly, gstrSysName: Exit Function
        Else
            chkEsign.Value = vbUnchecked
        End If
    End If
    
    Sign.姓名 = IIf(blnSpecify, strSpecifySign, IIf(mintSign = 0, gstrUserName, gstrSignName))
    Sign.签名人ID = IIf(blnSpecify, lngSpecifyId, glngUserId)
    Sign.签名级别 = Val(cmbLevel.Text)
    If Sign.签名级别 > cprSL_主任 Then Sign.签名级别 = cprSL_主任
    
    If Me.chkPreText.Value = vbChecked Then
        Sign.前置文字 = Trim(Mid(Me.cmbLevel.Text, 4)) & "："
    Else
        Sign.前置文字 = ""
    End If
    Sign.显示手签 = (chkHandSign.Value = vbChecked)
    Sign.签名方式 = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.签名时间 = zlDatabase.Currentdate()
    Select Case Me.cboTime.ListIndex
    Case 1: Sign.显示时间 = "yyyy-MM-dd hh:mm"
    Case 2: Sign.显示时间 = "yyyy年MM月dd日 hh:mm"
    Case Else: Sign.显示时间 = ""
    End Select
    
    '签名规则=2 使用RTF.Text做为数字签名原文 见cEPRSign注释
    Sign.签名规则 = 2
    Sign.签名信息 = strSign
    Sign.证书ID = lngCertID
    Sign.时间戳 = str时间戳
    Sign.时间戳信息 = str时间Base64
'    '签名规则=3 使用保存数据库后的内容文本（不含签名要素，签名对象,图片、表格及子对象）为数字签名原文
'    '数字签名信息在保存后进行数字签名后返回并单独保存
'    Sign.签名规则 = 3
'    Sign.签名信息 = IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), "") '如果数字签名，先存签名帐号，用于数字签名传入参数,签名完成后更改
'    Sign.证书ID = 0
'    Sign.时间戳 = ""
    
    If chkSignPic.Value = 1 And pic签名图片.Picture.Handle <> 0 And chkSignPic.Visible Then
        Sign.签名图片 = True
        Set mpicSign = pic签名图片.Picture
    ElseIf chkSignPic.Value = 1 And pic签名图片.Picture.Handle = 0 And chkSignPic.Visible Then
        MsgBox IIf(optName(0).Value, "当前", "指定") & "帐号没有可用的签名图，不能使用图片签名功能，请联系管理员！", vbExclamation, gstrSysName
        Exit Function
    Else
        Sign.签名图片 = False
        Set mpicSign = Nothing
    End If
    
    If pic签名图片.Tag <> "" Then '删除临时图片
        Kill pic签名图片.Tag
        pic签名图片.Tag = ""
    End If
    
    Validation = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## 功能：  验证用户名密码是否正确
'################################################################################################################
Private Function OraDataOpen(ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    Dim strSQL As String
    Dim strError As String
    Dim Cn As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    With Cn
        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
        .Open gcnOracle.ConnectionString, strUserName, strUserPwd
        If Err <> 0 Then
            OraDataOpen = False
            Exit Function
        End If
        .Close
    End With
    Set Cn = Nothing
    OraDataOpen = True
    Exit Function
errHand:
    Set Cn = Nothing
    OraDataOpen = False
    Err = 0
End Function

'################################################################################################################
'## 功能：  密码转换函数
'##
'## 参数：  strOld  :原密码
'##
'## 返回：  加密生成的密码
'################################################################################################################
Public Function TranPasswd(strOld As String) As String
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew
End Function

'################################################################################################################
'## 功能：  刷新控件
'################################################################################################################
Private Sub RefControls()
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case mlngPassType
        Case 0
            '密码签名
            chkEsign.Value = vbUnchecked
            chkEsign.Visible = False
        Case 1
            '1－数字
            chkEsign.Value = vbChecked
            chkEsign.Move txtPass.Left, txtPass.Top
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        End Select
    Else
        chkEsign.Enabled = True
        txtPass.Enabled = True
        txtName.Enabled = True
        Select Case mlngPassType
        Case 0
            '密码签名
            chkEsign.Value = vbUnchecked
            txtPass.Enabled = True
        Case 1
            '1－数字
            chkEsign.Value = vbChecked
            chkEsign.Move txtPass.Left, txtPass.Top
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        End Select
    End If
End Sub

Private Sub cboTime_Click()
     Call Preview
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then
        txtPass.SetFocus
    End If
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHandSign_Click()
     Call Preview
End Sub

Private Sub chkHandSign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkOrgPic_Click()
    If chkOrgPic.Value = vbUnchecked Then
        txtHeight.Visible = True
        lblH.Visible = True
    Else
        txtHeight.Visible = False
        lblH.Visible = False
    End If
    DrawSignPicture
End Sub

Private Sub chkPreText_Click()
    Call Preview
End Sub

Private Sub chkPreText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkSignPic_Click()
Dim strFile As String, rsTemp As ADODB.Recordset
    If chkSignPic.Value = 1 Then
        pic签名图片.Tag = ""
        pic签名图片.ToolTipText = ""
        lblWH.Caption = "": lblWH.Visible = True: lblH.Visible = True
        chkOrgPic.Visible = True
        txtHeight.Visible = True
        pic签名图片.Cls
        Set pic签名图片.Picture = Nothing
        If optName(1).Value And Trim(txtName) = "" Then Exit Sub '点击"指定帐号"
        gstrSQL = "Select b.签名图片 From 上机人员表 A, 人员表 B Where a.用户名 = '" & IIf(optName(0).Value, gstrDBUser, UCase(txtName)) & "' And a.人员id = b.id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "签名图片")
        If Not rsTemp.EOF Then
            strFile = zlDatabase.ReadPicture(rsTemp, "签名图片")
            pic签名图片.Tag = strFile
        End If
        strFile = pic签名图片.Tag
        
        If strFile <> "" Then
            Set morgSign = LoadPicture(strFile)
            pic签名图片.ToolTipText = "原始大小:" & CLng(pic签名图片.ScaleX(morgSign.Width, vbHimetric, vbPixels)) & " X " & CLng(pic签名图片.ScaleY(morgSign.Height, vbHimetric, vbPixels))
            DrawSignPicture
        End If
        chkPreText.Value = vbUnchecked: chkPreText.Enabled = False
        chkHandSign.Value = vbUnchecked: chkHandSign.Enabled = False
        cboTime.ListIndex = 0:          cboTime.Enabled = False
    Else
        Set morgSign = Nothing
        Set pic签名图片.Picture = Nothing
        pic签名图片.ToolTipText = ""
        lblWH.Caption = "": lblWH.Visible = False: lblH.Visible = False
        pic签名图片.Cls
        chkOrgPic.Visible = False
        txtHeight.Visible = False: pic签名图片.Move pic签名图片.Left, pic签名图片.Top, 810, 810
        Call DrawSignPicture
        chkPreText.Enabled = True
        chkHandSign.Enabled = True
        cboTime.Enabled = True
    End If
End Sub

Private Sub cmbLevel_Click()
    Call Preview
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    If pic签名图片.Tag <> "" Then '删除临时图片
        Kill pic签名图片.Tag
        pic签名图片.Tag = ""
    End If
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If InStr(gstrPrivsEpr, "参数设置") > 0 Then
        If chkHandSign.Tag <> chkHandSign.Value Then Call zlDatabase.SetPara("显示手签位置", chkHandSign.Value, glngSys, 1070)
        If chkPreText.Tag <> chkPreText.Value Then Call zlDatabase.SetPara("将签名级别作为前缀加入", chkPreText.Value, glngSys, 1070)
        If cboTime.Tag <> cboTime.ListIndex Then Call zlDatabase.SetPara("签名时间", cboTime.ListIndex, glngSys, 1070)
        If chkSignPic.Tag <> chkSignPic.Value Then Call zlDatabase.SetPara("签名使用图片", chkSignPic.Value, glngSys, 1070)
        If chkOrgPic.Tag <> chkOrgPic.Value Then Call zlDatabase.SetPara("签名使用原图", chkOrgPic.Value, glngSys, 1070)
        If txtHeight.Tag <> txtHeight.Text Then Call zlDatabase.SetPara("签名图片高度", txtHeight.Text, glngSys, 1070)
    End If
    If Validation Then
        mblnOk = True
        Unload Me
    End If
End Sub

Private Sub Preview()
    Dim StrText As String, bln手签 As Boolean, str前置文字 As String
    
    If Me.chkPreText.Value = vbChecked Then
        str前置文字 = Trim(Mid(Me.cmbLevel.Text, 4)) & "："
    Else
        str前置文字 = ""
    End If
    bln手签 = (chkHandSign.Value = vbChecked)
    StrText = str前置文字 & IIf(mintSign = 0, gstrUserName, gstrSignName) & IIf(bln手签, "，手签：_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        StrText = StrText & "，" & Me.cboTime.Text
    End If
    lblPreview.Caption = StrText
    
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngPassType = 2 Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", chkEsign.Value
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cmbLevel", cmbLevel.ListIndex
    Set frmParent = Nothing
End Sub

Private Sub Label1_Click()

End Sub

Private Sub optName_Click(Index As Integer)
    Call RefControls
    Call chkSignPic_Click
    If Index = 1 Then
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    End If
End Sub

Private Sub optName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub optPassType_Click(Index As Integer)
    If Index = 1 Then
        txtPass.Enabled = True
        If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus
    Else
        txtPass.Enabled = False
    End If
End Sub

Private Sub optPassType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtHeight_Change()
    On Error Resume Next
    DrawSignPicture
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus: Call Preview: Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
        End If
    End If
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtNames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtName_LostFocus()
    Call chkSignPic_Click
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < 32 Or KeyAscii > 126 Then KeyAscii = 0
    If InStr("""@\ ", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Function getPassType(ByVal lngFileKind As Long, ByVal lngPatientSource As Long, ByVal lngDeptId As Long, ByVal lngModule As Long) As Long
Dim rsTemp As New ADODB.Recordset, lS As Long
    On Error GoTo errHand
    '0-门诊医嘱和病历；1-住院医生医嘱和病历；2-住院护士医嘱；3-医技医嘱和报告；4-护理记录和护理病历；5-药品发药；6-LIS;7-PACS
    Select Case lngFileKind
        Case cpr门诊病历
            lS = 0
        Case cpr住院病历
            lS = 1
        Case cpr诊疗报告
            Select Case lngModule
                Case 1290, 1291, 1294
                    lS = 7
                Case Else
                    lS = 3
            End Select
        Case cpr护理病历
            lS = 4
        Case Else
            Select Case lngModule
            Case cprPF_门诊
                lS = 0
            Case cprPF_住院
                lS = 1
            Case Else
                lS = 1  '否则，以住院为准
            End Select
    End Select
    
    gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) as 启用 From Dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取电子签名控制部门", lS, lngDeptId)
    If rsTemp.EOF Then
        getPassType = 1
    Else
        getPassType = rsTemp!启用
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub DrawSignPicture()
    On Error Resume Next
    If Not morgSign Is Nothing Then
        If chkOrgPic.Value = vbChecked Then
            Set pic签名图片.Picture = morgSign
            pic签名图片.Appearance = 0: pic签名图片.BorderStyle = 0
            If pic签名图片.Width <> 810 Then pic签名图片.Move pic签名图片.Left, pic签名图片.Top, 810, 810
            pic签名图片.PaintPicture pic签名图片.Picture, 0, 0, pic签名图片.ScaleX(pic签名图片.Width, vbTwips, vbPixels), pic签名图片.ScaleY(pic签名图片.Height, vbTwips, vbPixels)
            lblWH.Caption = CLng(pic签名图片.ScaleX(morgSign.Width, vbHimetric, vbPixels)) & " X " & CLng(pic签名图片.ScaleY(morgSign.Height, vbHimetric, vbPixels)) & " Pixels"
        Else
            Dim lngWidth As Long
            lngWidth = CLng(txtHeight.Text * (morgSign.Width / morgSign.Height))
            pic签名图片.Appearance = 0: pic签名图片.BorderStyle = 0
            pic签名图片.Move pic签名图片.Left, pic签名图片.Top, pic签名图片.ScaleX(lngWidth, vbPixels, vbTwips), pic签名图片.ScaleY(txtHeight.Text, vbPixels, vbTwips)
            pic签名图片.PaintPicture morgSign, 0, 0, lngWidth, txtHeight.Text
            Set pic签名图片.Picture = pic签名图片.Image
            lblWH.Caption = lngWidth & " X " & txtHeight.Text & " Pixels"
        End If
    End If
    Err.Clear
End Sub
