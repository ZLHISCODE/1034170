VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm医保接口注册 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "注册医保接口"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frm医保接口注册.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt表空间 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      MaxLength       =   100
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txt中间库用户名 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   4
      Top             =   570
      Width           =   2775
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   17
      Top             =   3030
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1860
      TabIndex        =   16
      Top             =   3030
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "注册信息"
      Height          =   1530
      Left            =   150
      TabIndex        =   7
      Top             =   1380
      Width           =   4065
      Begin VB.TextBox txt说明 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1080
         Width           =   2745
      End
      Begin VB.TextBox txt密钥 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1530
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox txt名称 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         MaxLength       =   40
         TabIndex        =   11
         Top             =   690
         Width           =   2745
      End
      Begin VB.TextBox txt序号 
         Height          =   300
         Left            =   810
         MaxLength       =   3
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lbl说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   12
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl密钥 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "密钥"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   14
         Top             =   1590
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   10
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbl序号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "序号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   8
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.CommandButton cmd医保接口 
      Caption         =   "…"
      Height          =   285
      Left            =   3900
      TabIndex        =   2
      Top             =   180
      Width           =   285
   End
   Begin VB.TextBox txt医保接口部件 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   2475
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl表空间 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户表空间"
      Height          =   180
      Left            =   450
      TabIndex        =   5
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label lbl中间库用户名 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "中间库用户名"
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   630
      Width           =   1080
   End
   Begin VB.Label lbl医保接口部件 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医保接口部件"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frm医保接口注册"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mintRegist As Integer       '0-失败或取消;1-正常;2-升级;3-重新注册（指共用相同的医保部件以及中间库用户等，但险类不同）
Private mintInsure As Integer       '保险类别序号
Private mstrInsureUser As String    '中间库用户名
Private mstrInsureName As String    '医保接口名称
Private mstrInsureTablespace As String
Private mstrPath As String          '注册文件路径
Private mstrComponent As String     '部件名称
Private mstrDemo As String          '说明
Private mbln重新注册 As Boolean     '重新注册标志

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim blnRegist As Boolean        '是否允许重复注册
    Dim strFile As String, strMessage As String
    Dim objTest As Object
    Dim rsTest As New ADODB.Recordset
    '检查保险类别中是否已存在该险类，如果存在，表明是重新安装，这时只运行spNew.sql
    
    '步骤：
    '1、检查文件的合法性（zl9I_文件名）
    '2、创建指定的医保接口部件，失败则退出
    '3、部件合法性验证，失败则退出
    '4、检查已注册清单中，是否存在不同险类，相同部件名，如果存在则提示
    '----如果该医保接口已安装，再次安装说明需要运行修正脚本，5-9跳过----
    '5、检查注册文件的合法性，失败则退出
    '6、运行安装脚本，失败则退出
    
    If Trim(txt医保接口部件.Text) = "" Then
        MsgBox "请选择需要注册的医保部件！", vbInformation, gstrSysName
        cmd医保接口.SetFocus
        Exit Sub
    End If
    strFile = Mid(txt医保接口部件.Text, 1, Len(txt医保接口部件.Text) - 4)
    
    '检查是否允许重复注册
    Set objTest = CreateObject(strFile & ".CLS" & Mid(strFile, 4))
    strMessage = objTest.I_Support(I_Support重复注册)
    blnRegist = (Val(strMessage) = 1)
    
    '4、
    mstrSQL = " Select A.序号,A.名称,B.部件 As 医保部件,B.用户名,B.表空间" & _
              " From 保险类别 A,zlInsureComponents B" & _
              " Where A.序号=B.险类 And Upper(B.部件)='" & txt医保接口部件.Text & "'"
    Call zlDatabase.OpenRecordset(rsTest, mstrSQL, "装入已注册的医保接口")
    
    mintRegist = 0
    If rsTest.RecordCount <> 0 Then
        rsTest.Filter = "序号<>" & txt序号.Text
        If rsTest.RecordCount <> 0 Then
            If blnRegist Then
                '允许继续注册的原因是，可能一个医保接口部件应用于多个地区的医保，福州地区就是例子
                If MsgBox("发现其他已注册的医保接口，其部件名称与本次注册医保接口的部件名称一致，是否继续注册？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then rsTest.Filter = 0: Exit Sub
                mintRegist = 3
                '用户名与表空间以现有的为准
                txt中间库用户名.Text = Nvl(rsTest!用户名)
                txt表空间.Text = Nvl(rsTest!表空间)
            Else
                MsgBox "该医保接口不允许重复注册使用！", vbInformation, gstrSysName
                rsTest.Filter = 0: Exit Sub
            End If
        Else
            rsTest.Filter = "序号=" & txt序号.Text
            '说明是再次注册
            If rsTest.RecordCount <> 0 Then mintRegist = 2
        End If
        rsTest.Filter = 0
    End If
    
    '如果存在相同险类但部件名称不同的，也认为是重新注册
    If mintRegist = 0 Then
        mstrSQL = " Select A.序号,A.名称,upper(B.部件) As 医保部件" & _
                  " From 保险类别 A,zlInsureComponents B" & _
                  " Where A.序号=B.险类 And A.序号=" & txt序号.Text
        Call zlDatabase.OpenRecordset(rsTest, mstrSQL, "装入已注册的医保接口")
        If rsTest.RecordCount <> 0 Then
            MsgBox "发现其他已注册的医保接口，其保险序号与当前接口一致，但医保接口部件不同，不允许继续注册！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If mintRegist = 0 Then mintRegist = 1
    
    '检查数据合法性
    If Trim(txt序号.Text) = "" Then
        MsgBox "保险类别序号不能为空！", vbInformation, gstrSysName
        txt序号.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txt序号.Text) Then
        MsgBox "保险类别序号中含有非法字符！", vbInformation, gstrSysName
        txt序号.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Me.txt说明.Text, vbFromUnicode)) > 100 Then
        MsgBox "说明不能超过50个汉字或100个字符！", vbInformation, gstrSysName
        txt说明.SetFocus
        Exit Sub
    End If
    
    mstrComponent = Me.txt医保接口部件.Text
    mstrPath = Me.txt医保接口部件.Tag
    mintInsure = Val(txt序号.Text)
    mstrInsureUser = Trim(txt中间库用户名.Text)
    mstrInsureTablespace = Trim(txt表空间.Text)
    mstrInsureName = txt名称.Text
    mstrDemo = Trim(txt说明.Text)
    
    Unload Me
    Exit Sub
End Sub

Private Sub cmd医保接口_Click()
    Dim strFile As String, strPath As String
    Dim strMessage As String
    Dim arrMessage
    Dim str序号 As String, str名称 As String, str说明 As String
    Dim objTest As Object
    On Error GoTo ErrHand
    
    With CommonDialog1
        .Filter = "医保部件(*.dll)|*.dll"
        .ShowOpen
        Call GetFileOrPath(.FileName, strFile, strPath)
        strFile = Mid(strFile, 1, Len(strFile) - 4)
    End With

    '1、
    If Mid(strFile, 1, 5) <> "ZL9I_" Then
        MsgBox "请选择合法的医保接口部件！错误代码1", vbInformation, gstrSysName
        Exit Sub
    End If
    '2、
    On Error Resume Next
    Err = 0
    Set objTest = CreateObject(strFile & ".CLS" & Mid(strFile, 4))
    If Err <> 0 Then
        MsgBox "请选择合法的医保接口部件！错误代码2", vbInformation, gstrSysName
        Exit Sub
    End If
    '3、
    Err = 0
    strMessage = objTest.I_RegInfo
    If Err <> 0 Then
        MsgBox "请选择合法的医保接口部件！错误代码3", vbInformation, gstrSysName
        Set objTest = Nothing
        Exit Sub
    End If
    
    '3.1
    arrMessage = Split(strMessage, "|")
    If Not (UBound(arrMessage) >= 1) Then
        MsgBox "请选择合法的医保接口部件！错误代码3.1", vbInformation, gstrSysName
        Exit Sub
    End If
    str序号 = Val(arrMessage(0))
    str名称 = UCase(arrMessage(1))
    str说明 = UCase(arrMessage(2))
    If str序号 = 0 Or Trim(str名称) = "" Then
        MsgBox "医保接口注册信息不完整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '4、判断是否支持中间库
    Err = 0
    strMessage = objTest.I_Support(I_Support中间库)
    If Err <> 0 Then
        MsgBox "请选择合法的医保接口部件！错误代码4", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Me.txt序号.Text = str序号
    Me.txt名称.Text = str名称
    Me.txt说明.Text = str说明

    Me.txt医保接口部件.Text = strFile & ".DLL"
    Me.txt医保接口部件.Tag = strPath
    
    '支持中间库，则用户名与表空间缺省为医保部件的文件名
    If Val(strMessage) = 1 Then
        txt中间库用户名.Text = strFile
        txt表空间.Text = strFile
    End If
ErrHand:
    Exit Sub
End Sub

Public Function ShowRegist(intInsure As Integer, strInsureUser As String, strInsureTablespace As String, _
    strInsureName As String, strDemo As String, strComponent As String, strPath As String) As Integer
    'intInsure:保险类别序号
    'strInsureName:医保接口名称,eg“重庆市医保中心”
    'strComponent:医保部件名称
    'strPath:医保部件文件路径
    mintRegist = 0
    mintInsure = 0
    mstrInsureUser = ""
    mstrInsureTablespace = ""
    mstrInsureName = ""
    mstrDemo = ""
    mstrComponent = ""
    mstrPath = ""
    
    Me.Show 1
    
    ShowRegist = mintRegist
    If mintRegist > 0 Then
        strPath = mstrPath
        strComponent = mstrComponent
        intInsure = mintInsure
        strInsureUser = mstrInsureUser
        strInsureTablespace = mstrInsureTablespace
        strInsureName = mstrInsureName
        strDemo = mstrDemo
    End If
End Function

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt序号_GotFocus()
    Call zlControl.TxtSelAll(txt序号)
End Sub
