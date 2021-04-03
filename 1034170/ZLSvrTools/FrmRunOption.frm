VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRunOption 
   BackColor       =   &H80000005&
   Caption         =   "系统运行选项"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "FrmRunOption.frx":0000
   ScaleHeight     =   7335
   ScaleWidth      =   10230
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   255
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
      Begin VB.Image imgMain 
         Height          =   480
         Left            =   0
         Picture         =   "FrmRunOption.frx":04F9
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   6075
      Left            =   930
      TabIndex        =   13
      Top             =   570
      Width           =   5730
      Begin VB.CheckBox chkShutDown 
         BackColor       =   &H80000005&
         Caption         =   "允许关闭锁定的导航台"
         Height          =   255
         Left            =   255
         TabIndex        =   30
         Tag             =   "24"
         Top             =   5070
         Width           =   4455
      End
      Begin VB.CheckBox chkSpecial 
         BackColor       =   &H80000005&
         Caption         =   "复杂度控制（至少包含一个数字、字母、特殊符号）"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Tag             =   "23"
         Top             =   4200
         Width           =   4455
      End
      Begin VB.CheckBox chkLenCtrl 
         BackColor       =   &H80000005&
         Caption         =   "启用密码长度控制"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Tag             =   "20"
         Top             =   3855
         Width           =   1750
      End
      Begin VB.TextBox txtLen 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   23
         Tag             =   "21"
         Text            =   "3"
         Top             =   3840
         Width           =   300
      End
      Begin VB.TextBox txtLen 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   2910
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "22"
         Text            =   "12"
         Top             =   3840
         Width           =   300
      End
      Begin VB.TextBox Txt报表路径 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2070
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "6"
         Top             =   3090
         Width           =   3195
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "…"
         Height          =   300
         Left            =   5280
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3090
         Width           =   285
      End
      Begin VB.CheckBox Chk使用日志 
         BackColor       =   &H80000005&
         Caption         =   "运行日志记录(&S)"
         Height          =   210
         Left            =   240
         TabIndex        =   0
         Tag             =   "1"
         Top             =   315
         Width           =   1695
      End
      Begin VB.TextBox Txt运行日志最大条目数 
         Height          =   300
         Left            =   1965
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "2"
         Top             =   547
         Width           =   1755
      End
      Begin VB.CheckBox Chk是否记录运行错误 
         BackColor       =   &H80000005&
         Caption         =   "记录运行错误(&A)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1230
         Width           =   1695
      End
      Begin VB.TextBox Txt错误日志最大条目数 
         Height          =   300
         Left            =   1965
         MaxLength       =   18
         TabIndex        =   5
         Tag             =   "4"
         Top             =   1440
         Width           =   1755
      End
      Begin VB.TextBox Txt消息最大条目数 
         Height          =   300
         Left            =   1965
         MaxLength       =   18
         TabIndex        =   7
         Tag             =   "5"
         Top             =   2145
         Width           =   1755
      End
      Begin MSComCtl2.UpDown udLen 
         Height          =   270
         Index           =   1
         Left            =   3210
         TabIndex        =   25
         Top             =   3840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtLen(1)"
         BuddyDispid     =   196615
         BuddyIndex      =   1
         OrigLeft        =   3240
         OrigTop         =   3855
         OrigRight       =   3495
         OrigBottom      =   4110
         Max             =   16
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown udLen 
         Height          =   270
         Index           =   0
         Left            =   2340
         TabIndex        =   29
         Top             =   3840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtLen(0)"
         BuddyDispid     =   196615
         BuddyIndex      =   0
         OrigLeft        =   2370
         OrigTop         =   3855
         OrigRight       =   2625
         OrigBottom      =   4110
         Max             =   16
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.Label lblShutDown 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRunOption.frx":227B
         ForeColor       =   &H8000000D&
         Height          =   510
         Left            =   525
         TabIndex        =   31
         Top             =   5400
         Width           =   4350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRunOption.frx":22BD
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   480
         TabIndex        =   28
         Top             =   4545
         Width           =   3960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000005&
         Caption         =   "-->"
         Height          =   135
         Left            =   2640
         TabIndex        =   26
         Top             =   3915
         Width           =   375
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(请选择服务器上的Apply目录做为报表路径)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   21
         Top             =   3480
         Width           =   3510
      End
      Begin VB.Label Lbl报表路径 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL报表保存路径(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   3150
         Width           =   1800
      End
      Begin VB.Label LblOption1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(是否自动记录用户的使用系统的情况)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1950
         TabIndex        =   18
         Top             =   315
         Width           =   3060
      End
      Begin VB.Label Lbl运行日志条目数 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "日志最多保存天数(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1710
      End
      Begin VB.Label LblOption2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(使用日志最多保存的天数，超过时系统将自动删除超时的记录)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   870
         Width           =   5040
      End
      Begin VB.Label LblOption3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(是否记录使用过程中发生的各种错误)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1950
         TabIndex        =   16
         Top             =   1230
         Width           =   3060
      End
      Begin VB.Label Lbl错误日志最大条目数 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "错误最多保存天数(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   1490
         Width           =   1710
      End
      Begin VB.Label LblOption4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(错误记录最多保存的天数，超过时系统将自动删除超时的记录)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   1770
         Width           =   5040
      End
      Begin VB.Label Lbl消息最大条目数 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "消息保存最大天数(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   2205
         Width           =   1710
      End
      Begin VB.Label lblOption5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(消息最多能保存的天数，超过时系统将其自动删除。       天数为0时表示永久保存)"
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   480
         TabIndex        =   14
         Top             =   2520
         Width           =   4680
      End
   End
   Begin VB.CommandButton Cmd还原 
      Cancel          =   -1  'True
      Caption         =   "还原(&R)"
      Height          =   350
      Left            =   2190
      TabIndex        =   12
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton Cmd保存 
      Caption         =   "保存(&O)"
      Height          =   350
      Left            =   900
      TabIndex        =   11
      Top             =   6750
      Width           =   1100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统运行选项"
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
      Left            =   195
      TabIndex        =   20
      Top             =   150
      Width           =   1440
   End
End
Attribute VB_Name = "FrmRunOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RecOption As New ADODB.Recordset

Private Sub chkLenCtrl_Click()
    Dim blnEnabled  As Boolean
    blnEnabled = (chkLenCtrl.value = 1)
    txtLen(0).Enabled = blnEnabled
    txtLen(1).Enabled = blnEnabled
    udLen(0).Enabled = blnEnabled
    udLen(1).Enabled = blnEnabled
    Cmd保存.Enabled = True
End Sub

Private Sub chkShutDown_Click()
    Cmd保存.Enabled = True
End Sub

Private Sub chkSpecial_Click()
    Cmd保存.Enabled = True
End Sub

Private Sub cmdSelect_Click()
    Dim strPath As String
    strPath = OpenFolder(Me, "Excel报表保存路径：")
    If strPath = "" Then Exit Sub
    Txt报表路径 = strPath
    Cmd保存.Enabled = True
End Sub

Private Sub Cmd保存_Click()
    If Txt错误日志最大条目数.Enabled = True And Val(Txt错误日志最大条目数.Text) > 10 ^ 8 Then
        MsgBox "错误日志最大条目数太大。", vbInformation, gstrSysName
        Txt错误日志最大条目数.SetFocus
        Exit Sub
    End If
    If Txt运行日志最大条目数.Enabled = True And Val(Txt运行日志最大条目数.Text) > 10 ^ 8 Then
        MsgBox "运行日志最大条目数太大。", vbInformation, gstrSysName
        Txt运行日志最大条目数.SetFocus
        Exit Sub
    End If
    If Txt消息最大条目数.Enabled = True And Val(Txt消息最大条目数.Text) > 10 ^ 8 Then
        MsgBox "消息最大条目数太大。", vbInformation, gstrSysName
        Txt消息最大条目数.SetFocus
        Exit Sub
    End If
    If StrIsValid(Txt报表路径.Text, 50) = False Then
        Txt报表路径.SetFocus
        Exit Sub
    End If
    If SaveCons = False Then Exit Sub
End Sub

Private Sub Chk使用日志_Click()
    Cmd保存.Enabled = True
    Txt运行日志最大条目数.Enabled = Chk使用日志.value = 1
End Sub

Private Sub Chk是否记录运行错误_Click()
    Cmd保存.Enabled = True
    Txt错误日志最大条目数.Enabled = Chk是否记录运行错误.value = 1
End Sub

Private Sub Cmd还原_Click()
    Call InitCons
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ActiveControl Is Txt错误日志最大条目数 Or ActiveControl Is Txt消息最大条目数 Or ActiveControl Is Txt运行日志最大条目数 Then
        If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call InitCons
End Sub

Private Sub InitCons()
    Dim ConThis As Control
    '--初始化各控件的值--
    
    For Each ConThis In Controls
        If Val(ConThis.Tag) <> 0 Then
            Set RecOption = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Zloption", Val(ConThis.Tag))
            With RecOption
                If Val(ConThis.Tag) = 6 Then
                    ConThis.Enabled = Not (.EOF)
                    CmdSelect.Enabled = Not (.EOF)
                End If
                
                Select Case TypeName(ConThis)
                Case "CheckBox"
                    If .EOF Then
                        ConThis.value = 0
                    Else
                        ConThis.value = IIf(IsNull(!Option_Value), 0, !Option_Value)
                    End If
                Case "TextBox"
                    If .EOF Then
                        ConThis.Text = ""
                    Else
                        ConThis.Text = IIf(IsNull(!Option_Value), "", !Option_Value)
                    End If
                End Select
            End With
        End If
    Next
    Txt运行日志最大条目数.Enabled = Chk使用日志.value = 1
    Txt错误日志最大条目数.Enabled = Chk是否记录运行错误.value = 1
    
    Cmd保存.Enabled = False
End Sub

Private Function SaveCons() As Boolean
    Dim ConThis As Control, StrValue As String
    '--保存各控件的值--
    
    SaveCons = False
    On Error Resume Next
    err = 0
    
    gcnOracle.BeginTrans
    For Each ConThis In Controls
        If Val(ConThis.Tag) <> 0 Then
            Select Case TypeName(ConThis)
            Case "CheckBox"
                StrValue = ConThis.value
            Case "TextBox"
                StrValue = IIf(ConThis.Enabled = True, ConThis.Text, "")
            End Select
            gcnOracle.Execute "Update ZlOptions Set 参数值='" & StrValue & "' Where 参数号=" & Val(ConThis.Tag)
        End If
    Next
    
    If err <> 0 Then
        MsgBox "更新运行参数时，发生错误！", vbInformation, gstrSysName
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    gcnOracle.CommitTrans
    MsgBox "运行参数修改成功！", vbInformation, gstrSysName
    Cmd保存.Enabled = False
    SaveCons = True
End Function

Private Sub SelLen(ByVal ConObj As TextBox)
    With ConObj
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLen_Change(Index As Integer)
    Cmd保存.Enabled = True
    If Val(txtLen(0).Text) > Val(txtLen(1).Text) Then
        If Index = 0 Then
            txtLen(1).Text = txtLen(0).Text
        Else
            txtLen(0).Text = txtLen(1).Text
        End If
    End If
End Sub

Private Sub txtLen_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLen_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtLen(Index).Text) < udLen(Index).Min Then
        txtLen(Index).Text = udLen(Index).Min
    ElseIf Val(txtLen(Index).Text) > udLen(Index).Max Then
        txtLen(Index).Text = udLen(Index).Max
    End If
    If Val(txtLen(0).Text) > Val(txtLen(1).Text) Then
        If Index = 0 Then
            txtLen(1).Text = txtLen(0).Text
        Else
            txtLen(0).Text = txtLen(1).Text
        End If
    End If
    If Val(txtLen(1 - Index).Text) < udLen(1 - Index).Min Then
        txtLen(1 - Index).Text = udLen(1 - Index).Min
    ElseIf Val(txtLen(1 - Index).Text) > udLen(1 - Index).Max Then
        txtLen(1 - Index).Text = udLen(1 - Index).Max
    End If
End Sub

Private Sub Txt报表路径_Change()
    Cmd保存.Enabled = True
End Sub

Private Sub Txt报表路径_GotFocus()
    SelAll Txt报表路径
End Sub

Private Sub Txt错误日志最大条目数_Change()
    Cmd保存.Enabled = True
End Sub

Private Sub Txt错误日志最大条目数_GotFocus()
    SelLen Txt错误日志最大条目数
End Sub

Private Sub Txt消息最大条目数_Change()
    Cmd保存.Enabled = True
End Sub

Private Sub Txt消息最大条目数_GotFocus()
    SelLen Txt消息最大条目数
End Sub

Private Sub Txt运行日志最大条目数_Change()
    Cmd保存.Enabled = True
End Sub

Private Sub Txt运行日志最大条目数_GotFocus()
    SelLen Txt运行日志最大条目数
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

