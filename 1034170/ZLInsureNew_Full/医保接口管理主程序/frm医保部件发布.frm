VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm医保部件发布 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保部件发布"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frm医保部件发布.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8130
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd模块 
      Caption         =   "模块(&M)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6780
      TabIndex        =   10
      Top             =   4470
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3840
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd方法 
      Caption         =   "方法(&F)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6780
      TabIndex        =   11
      Top             =   4920
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img方法 
      Left            =   3210
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保部件发布.frx":1CFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img菜单 
      Left            =   840
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保部件发布.frx":2F7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd医保接口 
      Caption         =   "…"
      Height          =   285
      Left            =   5940
      TabIndex        =   4
      Top             =   960
      Width           =   285
   End
   Begin MSComctlLib.ListView lvw所支持的方法 
      Height          =   3675
      Left            =   2850
      TabIndex        =   9
      Top             =   1710
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   6482
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img方法"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6780
      TabIndex        =   13
      Top             =   990
      Width           =   1100
   End
   Begin VB.CommandButton cmd发布 
      Caption         =   "发布(&P)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6780
      TabIndex        =   12
      Top             =   510
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5745
      Left            =   6480
      TabIndex        =   14
      Top             =   -300
      Width           =   45
   End
   Begin MSComctlLib.ListView lvw菜单 
      Height          =   3675
      Left            =   0
      TabIndex        =   7
      Top             =   1710
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   6482
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img菜单"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -180
      TabIndex        =   5
      Top             =   1440
      Width           =   6705
   End
   Begin VB.TextBox txt名称 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2115
      MaxLength       =   20
      TabIndex        =   1
      Top             =   570
      Width           =   2775
   End
   Begin VB.TextBox txt险类 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   0
      Top             =   180
      Width           =   405
   End
   Begin VB.TextBox txt医保部件 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2115
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   3825
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   5415
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   635
      SimpleText      =   $"frm医保部件发布.frx":3DCE
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm医保部件发布.frx":3E15
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl险类 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "险类(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1425
      TabIndex        =   17
      Top             =   240
      Width           =   630
   End
   Begin VB.Label lbl名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1410
      TabIndex        =   16
      Top             =   630
      Width           =   630
   End
   Begin VB.Label lbl所支持的方法 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "所支持的方法"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   1530
      Width           =   3495
   End
   Begin VB.Label lbl菜单 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "菜单清单"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   1530
      Width           =   2745
   End
   Begin VB.Label lbl医保部件 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保部件(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1065
      TabIndex        =   2
      Top             =   1020
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   180
      Picture         =   "frm医保部件发布.frx":46A9
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frm医保部件发布"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mrsPrivs As New ADODB.Recordset      '各方法的权限
Private mrsModul As New ADODB.Recordset      '各模块使用到的方法
Private mrsMethod As New ADODB.Recordset     '本接口所使用到的方法

Private Sub cmd发布_Click()
    '以下变量用于Regist.txt
    Dim intInsure As Integer
    Dim strModuls As String
    Dim strFunctions As String
    Dim strPrivs As String
    Dim str权限 As String
    Dim str注册码 As String
    '以下变量用于Regist.sql
    Dim strInsert As String
    
    Dim intItem As Integer, intCount As Integer
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    
    intInsure = Val(txt险类.Text)
    If intInsure = 0 Then
        MsgBox "请先选择医保接口部件！", vbInformation, gstrSysname
        Exit Sub
    End If
    
    '根据所选择的模块、方法及各方法的权限，产生注册脚本及注册码
    If objFileSys.FileExists(txt医保部件.Tag & "\Regist.txt") Then
        If MsgBox("医保部件所在目录中，已经存在接口注册文件，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
    End If
    
    '获取使用模块清单
    intCount = lvw菜单.ListItems.Count
    For intItem = 1 To intCount
        If lvw菜单.ListItems(intItem).Checked Then
            strModuls = strModuls & "," & Mid(lvw菜单.ListItems(intItem).Key, 3)
        End If
    Next
    If strModuls <> "" Then strModuls = Mid(strModuls, 2)
    
    '获取支持的方法清单
    With mrsModul
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strFunctions = strFunctions & vbCrLf & !模块 & "|" & !权限串 & "|" & !方法
            .MoveNext
        Loop
    End With
    If strFunctions <> "" Then strFunctions = Mid(strFunctions, 3)
    
    '获取各方法的权限
    With mrsPrivs
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "方法,对象"
        Do While Not .EOF
            For intItem = 1 To 5
                If Mid(!权限, intItem, 1) = 1 Then
                    Select Case intItem
                    Case 1
                        str权限 = "SELECT"
                    Case 2
                        str权限 = "INSERT"
                    Case 3
                        str权限 = "UPDATE"
                    Case 4
                        str权限 = "DELETE"
                    Case 5
                        str权限 = "EXECUTE"
                    End Select
                    strPrivs = strPrivs & vbCrLf & !方法 & "|" & !对象 & "|" & str权限
                End If
            Next
            .MoveNext
        Loop
    End With
    If strPrivs <> "" Then strPrivs = Mid(strPrivs, 3)
    
    '写Regist.txt
    Set objStream = objFileSys.CreateTextFile(txt医保部件.Tag & "\Regist.txt", True)
    objStream.WriteLine "[MODULS]"
    objStream.WriteLine strModuls
    objStream.WriteBlankLines 1
    objStream.WriteLine "[FUNCTIONS]"
    objStream.WriteLine strFunctions
    objStream.WriteBlankLines 1
    objStream.WriteLine "[PRIVS]"
    objStream.WriteLine strPrivs
    objStream.Close
    
    '转换为实际的权限SQL
'    With mrsModul
'        If .RecordCount <> 0 Then .MoveFirst
'        Do While Not .EOF
'            mrsPrivs.Filter = "方法='" & !方法 & "'"
'            Do While Not mrsPrivs.EOF
'                For intItem = 1 To 5
'                    If Mid(mrsPrivs!权限, intItem, 1) = 1 Then
'                        Select Case intItem
'                        Case 1
'                            str权限 = "SELECT"
'                        Case 2
'                            str权限 = "INSERT"
'                        Case 3
'                            str权限 = "UPDATE"
'                        Case 4
'                            str权限 = "DELETE"
'                        Case 5
'                            str权限 = "EXECUTE"
'                        End Select
'                        strInsert = strInsert & vbCrLf & "Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100," & _
'                        !模块 & ",'" & !权限串 & "','USER','" & mrsPrivs!对象 & "','" & str权限 & "');"
'                    End If
'                Next
'                mrsPrivs.MoveNext
'            Loop
'            mrsPrivs.Filter = 0
'            .MoveNext
'        Loop
'    End With
'    If strInsert <> "" Then strInsert = Mid(strInsert, 3)
    
    '得到管理数据的插入SQL语句
    intCount = lvw菜单.ListItems.Count
    For intItem = 1 To intCount
        If lvw菜单.ListItems(intItem).Checked Then
            strInsert = strInsert & vbCrLf & _
                "Insert into zlInsureModuls(险类,序号) Values (" & intInsure & "," & _
                Mid(lvw菜单.ListItems(intItem).Key, 3) & ");"
        End If
    Next
    With mrsModul
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strInsert = strInsert & vbCrLf & _
                "Insert into zlInsureFuncs(险类,序号,功能,方法) Values (" & intInsure & "," & _
                !模块 & ",'" & !权限串 & "','" & !方法 & "');"
            .MoveNext
        Loop
    End With
    With mrsPrivs
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            For intItem = 1 To 5
                If Mid(mrsPrivs!权限, intItem, 1) = 1 Then
                    Select Case intItem
                    Case 1
                        str权限 = "SELECT"
                    Case 2
                        str权限 = "INSERT"
                    Case 3
                        str权限 = "UPDATE"
                    Case 4
                        str权限 = "DELETE"
                    Case 5
                        str权限 = "EXECUTE"
                    End Select
                    strInsert = strInsert & vbCrLf & _
                        "Insert Into zlInsurePrivs(险类,方法,对象,权限) Values(" & intInsure & "," & _
                        "'" & !方法 & "','" & !对象 & "','" & str权限 & "');"
                End If
            Next
            .MoveNext
        Loop
    End With
    If strInsert <> "" Then strInsert = Mid(strInsert, 3)
    
    '写Regist.sql，实际的权限脚本
    Set objStream = objFileSys.CreateTextFile(txt医保部件.Tag & "\Regist.sql", True)
    objStream.WriteLine strInsert
    objStream.Close
    
    MsgBox "注册文件已经产生！", vbInformation, gstrSysname
End Sub

Private Sub cmd方法_Click()
    If Val(txt险类.Text) = 0 Then
        MsgBox "请先选择医保接口部件！", vbInformation, gstrSysname
        Exit Sub
    End If
    Call MakeMethodRecord
    Call frm权限设置.ShowEditor(mrsPrivs, mrsMethod)
End Sub

Private Sub cmd模块_Click()
    If Val(txt险类.Text) = 0 Then
        MsgBox "请先选择医保接口部件！", vbInformation, gstrSysname
        Exit Sub
    End If
    Call MakeMethodRecord
    Call frm方法设置.ShowEditor(mrsModul, mrsMethod)
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd医保接口_Click()
    Dim strMessage As String
    Dim strFile As String, strPath As String
    Dim arrMessage
    Dim objTest As Object
    
    cmd发布.Enabled = False
    cmd模块.Enabled = False
    cmd方法.Enabled = False
    
    With dlg
        .Filter = "医保部件(*.dll)|*.dll"
        .ShowOpen
        Call GetFileOrPath(.FileName, strFile, strPath)
        strFile = Mid(strFile, 1, Len(strFile) - 4)
        txt医保部件.Tag = strPath
    End With

    '1、
    If Mid(strFile, 1, 5) <> "ZL9I_" Then
        MsgBox "请选择合法的医保接口部件！错误代码1", vbInformation, gstrSysname
        Exit Sub
    End If
    '2、
    On Error Resume Next
    Err = 0
    Set objTest = CreateObject(strFile & ".CLS" & Mid(strFile, 4))
    If Err <> 0 Then
        MsgBox "请选择合法的医保接口部件！错误代码2", vbInformation, gstrSysname
        Exit Sub
    End If
    '3、
    Err = 0
    strMessage = objTest.I_RegInfo
    If Err <> 0 Then
        MsgBox "请选择合法的医保接口部件！错误代码3", vbInformation, gstrSysname
        Set objTest = Nothing
        Exit Sub
    End If
    
    arrMessage = Split(strMessage, "|")
    If Not (UBound(arrMessage) >= 1) Then
        MsgBox "请选择合法的医保接口部件！错误代码3.1", vbInformation, gstrSysname
        Exit Sub
    End If
    
    If Val(arrMessage(0)) = 0 Then
        MsgBox "医保接口的险类不能为空！", vbInformation, gstrSysname
        Exit Sub
    End If
    If Trim(UCase(arrMessage(1))) = "" Then
        MsgBox "医保接口的名称不能为空！", vbInformation, gstrSysname
        Exit Sub
    End If
    
    txt险类.Text = Val(arrMessage(0))
    txt名称.Text = UCase(arrMessage(1))
    txt医保部件.Text = UCase(strFile) & ".DLL"
    
    cmd发布.Enabled = True
    cmd模块.Enabled = True
    cmd方法.Enabled = True
    
    Call ShowRegist
    Exit Sub
ErrHand:
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim intItem As Integer, intCount As Integer
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    '加入菜单体系
    mstrSQL = "Select 序号,标题 From zlPrograms Where Upper(部件)='ZL9INSURE' Order By 序号"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "读取菜单体系")
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvw菜单.ListItems.Add(, "K_" & rsTemp!序号, rsTemp!标题, , 1)
            lvwItem.Checked = True
            .MoveNext
        Loop
    End With
    
    '加入方法列表（固定）
    With lvw所支持的方法
        Call .ListItems.Add(, "K_" & 方法.身份验证, "Identify()", , 1)
        Call .ListItems.Add(, "K_" & 方法.身份验证_自助挂号, "Identify2()", , 1)
        Call .ListItems.Add(, "K_" & 方法.帐户余额, "SelfBalance()", , 1)
        Call .ListItems.Add(, "K_" & 方法.门诊挂号, "RegistSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.门诊挂号作废, "RegistDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.门诊虚拟结算, "ClinicPreSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.门诊结算, "ClinicSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.门诊结算作废, "ClinicDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.个人帐户转预交, "TransferSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.预交退个人帐户, "TransferDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.住院虚拟结算, "WipeoffMoney()", , 1)
        Call .ListItems.Add(, "K_" & 方法.住院结算, "SettleSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.住院结算作废, "SettleDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.入院登记, "ComeInSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.入院登记撤销, "ComeInDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.出院登记, "LeaveSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.出院登记撤销, "LeaveDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.费用明细上传, "TranChargeDetail()", , 1)
        Call .ListItems.Add(, "K_" & 方法.住院信息变动, "ModiPatiSwap()", , 1)
        Call .ListItems.Add(, "K_" & 方法.获取医保项目信息, "GetItemInfo()", , 1)
        Call .ListItems.Add(, "K_" & 方法.病种选择, "ChooseDisease()", , 1)
        
        For intItem = 1 To 总数
            Set lvwItem = .ListItems(intItem)
            lvwItem.Checked = True
        Next
    End With
    
End Sub

Private Sub lvw菜单_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '保险类别必选
    If Item.Key = "K_1600" Then Item.Checked = True
End Sub

Private Sub txt名称_GotFocus()
    Call zlControl.TxtSelAll(txt名称)
End Sub

Private Sub txt险类_GotFocus()
    Call zlControl.TxtSelAll(txt险类)
End Sub

Private Sub txt医保部件_GotFocus()
    Call zlControl.TxtSelAll(txt医保部件)
End Sub

Private Sub GetFileOrPath(ByVal strInput As String, strFile As String, strPath As String)
    Dim intPos As Integer
    '根据完整的文件路径，返回文件路径及文件名：C:\Appsoft\Apply\zl9Insure.dll，返回的文件名是zl9Insure.dll，而路径名是C:\Appsoft\Apply
    intPos = 1
    Do While True
        If InStr(intPos, strInput, "\") = 0 Then Exit Do
        intPos = InStr(intPos, strInput, "\") + 1
    Loop
    If intPos = 1 Then Exit Sub
    
    strPath = UCase(Mid(strInput, 1, intPos - 2))
    strFile = UCase(Mid(strInput, intPos))
End Sub

Private Sub MakeMethodRecord()
    Dim intItem As Integer, intCount As Integer
    Dim strMethod As String
    Dim strField As String, strValue As String
    
    strField = "序号," & adDouble & ",18|方法," & adLongVarChar & ",50"
    Call Record_Init(mrsMethod, strField)
    
    strField = "序号|方法"
    intCount = lvw所支持的方法.ListItems.Count
    For intItem = 1 To intCount
        If lvw所支持的方法.ListItems(intItem).Checked Then
            strMethod = lvw所支持的方法.ListItems(intItem).Text
            strMethod = Mid(strMethod, 1, Len(strMethod) - 2)
            strValue = Mid(lvw所支持的方法.ListItems(intItem).Key, 3) & "|" & strMethod
            Call Record_Add(mrsMethod, strField, strValue)
        End If
    Next
End Sub

Private Sub ShowRegist()
    Dim strTemp As String
    Dim strBase As String
    Dim strFunctions As String
    Dim strPrivs As String
    Dim arrData
    Dim intItem As Integer, intCount As Integer
    Dim strFields As String
    Dim str方法 As String, str对象 As String, str权限 As String
    On Error GoTo ErrHand
    
    '初始化权限记录集
    strFields = "方法," & adLongVarChar & "," & 50 & "|对象," & adLongVarChar & "," & 50 & "|权限," & adLongVarChar & "," & 5
    Call Record_Init(mrsPrivs, strFields)
    strFields = "模块," & adDouble & "," & 18 & "|权限串," & adLongVarChar & "," & 50 & "|方法," & adLongVarChar & "," & 50
    Call Record_Init(mrsModul, strFields)
    
    If Not ReadFile(strBase, strFunctions, strPrivs) Then Exit Sub
    
    '清除所有选择
    intCount = lvw菜单.ListItems.Count
    For intItem = 1 To intCount
        lvw菜单.ListItems(intItem).Checked = False
    Next
    intCount = lvw所支持的方法.ListItems.Count
    For intItem = 1 To intCount
        lvw所支持的方法.ListItems(intItem).Checked = False
    Next
    
    '根据注册文件显示
    '菜单
    arrData = Split(strBase, ",")
    intCount = UBound(arrData)
    For intItem = 0 To intCount
        lvw菜单.ListItems("K_" & arrData(intItem)).Checked = True
    Next
    
    '方法
    arrData = Split(strFunctions, vbCrLf)
    intCount = UBound(arrData)
    For intItem = 0 To intCount
        If InStr(1, strTemp & ",", "," & UCase(Split(arrData(intItem), "|")(2)) & ",") = 0 Then
            strTemp = strTemp & "," & UCase(Split(arrData(intItem), "|")(2))
            Select Case UCase(Split(arrData(intItem), "|")(2))
            Case "IDENTIFY"
                lvw所支持的方法.ListItems("K_" & 方法.身份验证).Checked = True
            Case "IDENTIFY2"
                lvw所支持的方法.ListItems("K_" & 方法.身份验证_自助挂号).Checked = True
            Case "SELFBALANCE"
                lvw所支持的方法.ListItems("K_" & 方法.帐户余额).Checked = True
            Case "REGISTSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.门诊挂号).Checked = True
            Case "REGISTDELSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.门诊挂号作废).Checked = True
            Case "CLINICPRESWAP"
                lvw所支持的方法.ListItems("K_" & 方法.门诊虚拟结算).Checked = True
            Case "CLINICSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.门诊结算).Checked = True
            Case "CLINICDELSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.门诊结算作废).Checked = True
            Case "TRANSFERSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.个人帐户转预交).Checked = True
            Case "TRANSFERDELSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.预交退个人帐户).Checked = True
            Case "WIPEOFFMONEY"
                lvw所支持的方法.ListItems("K_" & 方法.住院虚拟结算).Checked = True
            Case "SETTLESWAP"
                lvw所支持的方法.ListItems("K_" & 方法.住院结算).Checked = True
            Case "SETTLEDELSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.住院结算作废).Checked = True
            Case "COMEINSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.入院登记).Checked = True
            Case "COMEINDELSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.入院登记撤销).Checked = True
            Case "LEAVESWAP"
                lvw所支持的方法.ListItems("K_" & 方法.出院登记).Checked = True
            Case "LEAVEDELSWAP"
                lvw所支持的方法.ListItems("K_" & 方法.出院登记撤销).Checked = True
            Case "TRANCHARGEDETAIL"
                lvw所支持的方法.ListItems("K_" & 方法.费用明细上传).Checked = True
            Case "MODIPATIDWAP"
                lvw所支持的方法.ListItems("K_" & 方法.住院信息变动).Checked = True
            Case "GETITEMINFO"
                lvw所支持的方法.ListItems("K_" & 方法.获取医保项目信息).Checked = True
            Case "CHOOSEDISEASE"
                lvw所支持的方法.ListItems("K_" & 方法.病种选择).Checked = True
            End Select
        End If
    Next
    
    '权限
    For intItem = 0 To intCount
        Call Record_Add(mrsModul, "模块|权限串|方法", arrData(intItem))
    Next
    arrData = Split(strPrivs, vbCrLf)
    intCount = UBound(arrData)
    str权限 = "00000"
    For intItem = 0 To intCount
        If (str方法 <> Split(arrData(intItem), "|")(0) Or str对象 <> Split(arrData(intItem), "|")(1)) Then
            If str方法 <> "" Then Call Record_Add(mrsPrivs, "方法|对象|权限", str方法 & "|" & str对象 & "|" & str权限)
            str方法 = Split(arrData(intItem), "|")(0)
            str对象 = Split(arrData(intItem), "|")(1)
            str权限 = "00000"
        End If
        
        Select Case UCase(Split(arrData(intItem), "|")(2))
        Case "SELECT"
            str权限 = "1" & Mid(str权限, 2)
        Case "INSERT"
            str权限 = Mid(str权限, 1, 1) & "1" & Mid(str权限, 3)
        Case "UPDATE"
            str权限 = Mid(str权限, 1, 2) & "1" & Mid(str权限, 4)
        Case "DELETE"
            str权限 = Mid(str权限, 1, 3) & "1" & Mid(str权限, 5)
        Case Else
            str权限 = Mid(str权限, 1, 4) & "1"
        End Select
    Next
    Call Record_Add(mrsPrivs, "方法|对象|权限", str方法 & "|" & str对象 & "|" & str权限)
    
    Exit Sub
    
ErrHand:
    MsgBox "装入注册文件时发生未知错误！", vbInformation, gstrSysname
End Sub

Private Function ReadFile(strBase As String, strFunctions As String, strPrivs As String) As Boolean
    Dim intState As Integer
    Dim strLine As String
    Dim strPath As String
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    Const strRegist As String = "Regist.txt"
    '解析文件
    strPath = txt医保部件.Tag & "\" & strRegist
    If Not objFileSys.FileExists(strPath) Then Exit Function
    Set objStream = objFileSys.OpenTextFile(strPath, ForReading)
    Do While Not objStream.AtEndOfStream
        strLine = UCase(objStream.ReadLine)
        Select Case strLine
        Case "[MODULS]"
            intState = 1
        Case "[FUNCTIONS]"
            intState = 2
        Case "[PRIVS]"
            intState = 3
        Case Else
            If Trim(strLine) <> "" Then
                Select Case intState
                Case 1  'MODULS
                    strBase = strLine
                Case 2  'FUNCTIONS
                    strFunctions = strFunctions & IIf(strFunctions = "", "", vbCrLf) & strLine
                Case 3  'PRIVS
                    strPrivs = strPrivs & IIf(strPrivs = "", "", vbCrLf) & strLine
                End Select
            End If
        End Select
    Loop
    
    strBase = Trim(strBase)
    strFunctions = Trim(strFunctions)
    strPrivs = Trim(strPrivs)
    
    objStream.Close
    ReadFile = Not (strBase = "" Or strFunctions = "" Or strPrivs = "")
End Function
