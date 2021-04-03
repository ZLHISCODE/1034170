VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMediPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmMediPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt数字码 
      Height          =   300
      Left            =   5400
      TabIndex        =   25
      Top             =   3240
      Width           =   720
   End
   Begin MSComCtl2.UpDown udg数字码 
      Height          =   300
      Left            =   6120
      TabIndex        =   24
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txt数字码"
      BuddyDispid     =   196609
      OrigLeft        =   6720
      OrigTop         =   3720
      OrigRight       =   6975
      OrigBottom      =   4020
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Frame fra分批 
      Caption         =   "3、药品分批属性自动设置"
      ForeColor       =   &H00800000&
      Height          =   1080
      Left            =   120
      TabIndex        =   18
      Top             =   2300
      Width           =   4035
      Begin VB.OptionButton optAllNotSet 
         Caption         =   "药库和药房都不分批"
         Height          =   200
         Left            =   1920
         TabIndex        =   22
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optAllSet 
         Caption         =   "药库和药房分批"
         Height          =   200
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton opt手动 
         Caption         =   "手工设置分批属性"
         Height          =   200
         Left            =   120
         TabIndex        =   20
         Top             =   390
         Width           =   1735
      End
      Begin VB.OptionButton optOnly药库 
         Caption         =   "仅药库分批"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fra售价方式 
      Caption         =   "2、新增规格售价计算方式"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   1300
      Width           =   4035
      Begin VB.OptionButton opt分段加成 
         Caption         =   "按分段加成计算售价"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   3735
      End
      Begin VB.OptionButton opt一般加成 
         Caption         =   "按一般加成率计算售价"
         Height          =   200
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraIncome 
      Height          =   1005
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4035
      Begin VB.ComboBox cbo收入项目 
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   315
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label LblNote 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   14
         Top             =   390
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   60
         Picture         =   "frmMediPara.frx":000C
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lblIncome 
         AutoSize        =   -1  'True
         Caption         =   "1、各材质对应缺省收入项目"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   585
         TabIndex        =   13
         Top             =   0
         Width           =   2250
      End
   End
   Begin VB.Frame frmStockRange 
      Caption         =   "4、设置存储库房时允许应用于的范围"
      ForeColor       =   &H00800000&
      Height          =   3030
      Left            =   4215
      TabIndex        =   3
      Top             =   105
      Width           =   3585
      Begin VB.CheckBox chk应用范围 
         Caption         =   "仅应用于当前选择的药品(&1)"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   285
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chk应用范围 
         Caption         =   "应用于所有当前选择的同品种药品(&2)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   9
         Top             =   660
         Value           =   1  'Checked
         Width           =   3270
      End
      Begin VB.CheckBox chk应用范围 
         Caption         =   "应用于所有当前选择的同材质药品(&3)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   1035
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin VB.CheckBox chk应用范围 
         Caption         =   "应用于所有当前选择的同剂型药品(&4)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   7
         Top             =   1410
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin VB.CheckBox chk应用范围 
         Caption         =   "应用于所有同级的药品(&5)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   4
         Left            =   210
         TabIndex        =   6
         Top             =   1785
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox chk应用范围 
         Caption         =   "应用于所有当前分类下的药品(&6)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   5
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2985
      End
      Begin VB.Label lblComment 
         Caption         =   "提示：没有选择到的应用范围在设置存储库房时将不能选择。"
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   2880
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5325
      TabIndex        =   0
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6675
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin VB.Label lbl数字码 
      AutoSize        =   -1  'True
      Caption         =   "数字码长度"
      Height          =   180
      Left            =   4320
      TabIndex        =   23
      Top             =   3300
      Width           =   900
   End
End
Attribute VB_Name = "frmMediPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnActive As Boolean
Private intTabIndex As Integer
Private lng西成药 As Long, lng中草药 As Long, lng中成药 As Long
Private strPrivs As String
Private rs收入项目 As New ADODB.Recordset
Dim mblnSetPara As Boolean      '是否具有参数设置权限
Private Sub SetFramSize()
    Dim dlbTopTmp As Double
    Dim n As Integer
    
    frmStockRange.Top = fraIncome.Top
    lblComment.Top = chk应用范围(5).Top + chk应用范围(5).Height + 200 'frmStockRange.Height - lblComment.Height - 100
    fra售价方式.Top = fraIncome.Top + fraIncome.Height + 100
    fra分批.Top = fra售价方式.Height + fra售价方式.Top + 100
    frmStockRange.Height = udg数字码.Top + udg数字码.Height - 400
    lbl数字码.Top = frmStockRange.Top + frmStockRange.Height + 200
    txt数字码.Top = lbl数字码.Top - 60
    udg数字码.Top = txt数字码.Top
    
    
    With cmdOK
        .Top = fra分批.Top + fra分批.Height + 150
        .TabIndex = intTabIndex
    End With
    With cmdCancel
        .Top = cmdOK.Top
        .TabIndex = intTabIndex + 1
    End With
    With cmdHelp
        .Top = cmdOK.Top
        .TabIndex = intTabIndex + 2
    End With
    Me.Height = cmdOK.Top + cmdOK.Height + 550
'    dlbTopTmp = lblComment.Top - chk应用范围(0).Top
'
'    dlbTopTmp = Int(dlbTopTmp / 6)
    
'    For n = 1 To 5
'        chk应用范围(n).Top = chk应用范围(n - 1).Top + dlbTopTmp
'    Next
End Sub

Public Sub ShowMe(ByVal strPrivss As String, ByVal frmParent As Object)
    strPrivs = strPrivss
    Me.Show 1, frmParent
End Sub

'Private Sub chk品种连续_Click()
'    If Me.chk品种连续.Value = 1 Then
'        Me.chk品种规格.Value = 0: Me.chk品种规格.Enabled = False
'    Else
'        Me.chk品种规格.Enabled = True
'    End If
'End Sub

Private Sub cmdCancel_Click()
    gblnIncomeItem = False
    Unload Me
End Sub
Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim intSave As Integer
    Dim strRange As String
    Dim intSetMethod As Integer '库房设置方式,0-手工设置分批属性（默认值）；1-仅药库分批；2-药库和药房分批；3-药库和药房都不分批
    Dim n As Integer
        
'    If Me.chk品种连续.Value = 0 Then
'        If Me.chk品种规格.Value = 0 Then
'            zldatabase.SetPara "品种增加模式", 0, glngSys, 1023
'        Else
'            zldatabase.SetPara "品种增加模式", 2, glngSys, 1023
'        End If
'    Else
'        zldatabase.SetPara "品种增加模式", 1, glngSys, 1023
'    End If
'    If Me.chk规格连续.Value = 0 Then
'        zldatabase.SetPara "规格增加模式", 0, glngSys, 1023
'    Else
'        zldatabase.SetPara "规格增加模式", 1, glngSys, 1023
'    End If
    
    If opt手动.Value = True Then
        intSetMethod = 0
    ElseIf optOnly药库.Value = True Then
        intSetMethod = 1
    ElseIf optAllSet.Value = True Then
        intSetMethod = 2
    ElseIf optAllNotSet.Value = True Then
        intSetMethod = 3
    End If
    
    For intSave = 1 To LblNote.UBound
        zlDatabase.SetPara intSave, cbo收入项目(intSave).ItemData(cbo收入项目(intSave).ListIndex), glngSys, 1023
    Next
    
    For n = 0 To chk应用范围.Count - 1
        strRange = strRange & chk应用范围(n).Value
    Next
    
    If opt一般加成.Value = True Then
        zlDatabase.SetPara "售价按加成计算", 0, glngSys, 1023
    Else
        zlDatabase.SetPara "售价按加成计算", 1, glngSys, 1023
    End If
    
    zlDatabase.SetPara "应用范围", strRange, glngSys, 1023
    zlDatabase.SetPara "药品分批属性自动设置", intSetMethod, glngSys, 1023
    zlDatabase.SetPara "数字码", txt数字码.Text, glngSys, 1023
    
    gblnIncomeItem = True
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not blnActive Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strRange As String
    Dim n As Integer
    Dim int品种增加 As Integer
    Dim int规格增加 As Integer
    Dim strTmp As String
    Dim intTmp As Integer
    Dim int售价 As Integer
    Dim intSet分批 As Integer   '分批设置方式 0-手工设置分批属性（默认值）；1-仅药库分批；2-药库和药房分批；3-药库和药房都不分批
    Dim rsTemp As ADODB.Recordset
    
    mblnSetPara = InStr(strPrivs, "参数设置") > 0

    '根据用户权限，装入控件
    On Error GoTo errHandle
    intTabIndex = 2
    blnActive = False
    
    gstrSql = "select nvl(max(length(简码)),0) 长度 from 收费项目别名 where 码类=3"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "数字码长度")
    
    If rsTemp!长度 = 0 Then
        udg数字码.Min = 7
    Else
        udg数字码.Min = rsTemp!长度
    End If
    udg数字码.Max = 40
'    int品种增加 = Val(zldatabase.GetPara("品种增加模式", glngSys, 1023, 0, Array(chk品种连续, chk品种规格), mblnSetPara))
'    Select Case int品种增加
'    Case 1
'        Me.chk品种连续.Value = 1
'        Me.chk品种规格.Value = 0: Me.chk品种规格.Enabled = False
'    Case 2
'        Me.chk品种连续.Value = 0
'        Me.chk品种规格.Value = 1: Me.chk品种规格.Enabled = True And mblnSetPara = True
'    Case Else
'        Me.chk品种连续.Value = 0
'        Me.chk品种规格.Enabled = True And mblnSetPara = True
'    End Select
    
'    int规格增加 = Val(zldatabase.GetPara("规格增加模式", glngSys, 1023, 0, Array(chk规格连续), mblnSetPara))
    
'    If int规格增加 = 0 Then
'        Me.chk规格连续.Value = 0
'    Else
'        Me.chk规格连续.Value = 1
'    End If

    gstrSql = "Select ID,编码||'-'||名称 名称 From 收入项目 Where 末级=1"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rs收入项目 = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rs收入项目
        If .EOF Then
            MsgBox "请初始化收入项目（收入项目）！", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    lng西成药 = Val(zlDatabase.GetPara("西成药收入项目", glngSys, 1023, 0))
    lng中成药 = Val(zlDatabase.GetPara("中成药收入项目", glngSys, 1023, 0))
    lng中草药 = Val(zlDatabase.GetPara("中草药收入项目", glngSys, 1023, 0))
    
    If strPrivs Like "*西成药*" Then Call AddCons("西成药")
    If strPrivs Like "*中成药*" Then Call AddCons("中成药")
    If strPrivs Like "*中草药*" Then Call AddCons("中草药")
    
    For n = 0 To cbo收入项目.UBound
        If n = 0 Then strTmp = "西成药收入项目"
        If n = 1 Then strTmp = "中成药收入项目"
        If n = 2 Then strTmp = "中草药收入项目"
        
        intTmp = Val(zlDatabase.GetPara(strTmp, glngSys, 1023, 0, Array(cbo收入项目(n)), mblnSetPara))
    Next
    
    strRange = zlDatabase.GetPara("应用范围", glngSys, 1023, "111111", Array(frmStockRange, chk应用范围(1), chk应用范围(2), chk应用范围(3), chk应用范围(4), chk应用范围(5)), mblnSetPara)
    For n = 1 To chk应用范围.Count - 1
        chk应用范围(n).Value = Mid(strRange, n + 1, 1)
    Next

    
    int售价 = Val(zlDatabase.GetPara("售价按加成计算", glngSys, 1023, 0))
    
    If int售价 = 0 Then
        opt一般加成.Value = True
        opt分段加成.Value = False
    Else
        opt一般加成.Value = False
        opt分段加成.Value = True
    End If
    
    intSet分批 = Val(zlDatabase.GetPara("药品分批属性自动设置", glngSys, 1023, 0))
    Select Case intSet分批
        Case 0
            opt手动.Value = True
        Case 1
            optOnly药库.Value = True
        Case 2
            optAllSet.Value = True
        Case 3
            optAllNotSet.Value = True
    End Select
    txt数字码.Text = Val(zlDatabase.GetPara("数字码", glngSys, 1023, 7))
    
    '设置窗体大小及各控件位置
    fraIncome.Height = cbo收入项目(cbo收入项目.UBound).Top + cbo收入项目(0).Height + 200
    
    Call SetFramSize
    
    blnActive = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddCons(ByVal strName As String)
    Dim intIdx As Integer
    intIdx = LblNote.UBound + 1
    Load LblNote(intIdx)
    Load cbo收入项目(intIdx)
    
    LblNote(intIdx).ForeColor = LblNote(0).ForeColor
    cbo收入项目(intIdx).ForeColor = cbo收入项目(0).ForeColor
    
    intTabIndex = intTabIndex + 1
    With LblNote(intIdx)
        .Caption = strName
        .TabIndex = intTabIndex
        .Container = fraIncome
        .Top = IIf(intIdx = 1, LblNote(0).Top, LblNote(intIdx - 1).Top) + IIf(intIdx = 1, 0, LblNote(0).Height + 200)
        .Left = LblNote(0).Left + LblNote(0).Width - .Width
        .Visible = True
    End With
    intTabIndex = intTabIndex + 1
    With cbo收入项目(intIdx)
        .Container = fraIncome
        .Left = cbo收入项目(0).Left
        .Top = IIf(intIdx = 1, cbo收入项目(0).Top, cbo收入项目(intIdx - 1).Top) + IIf(intIdx = 1, 0, cbo收入项目(0).Height + 100)
        .TabIndex = intTabIndex
        .Visible = True
    End With
    Call AddItem(cbo收入项目(intIdx), strName)
End Sub

Private Sub AddItem(ByVal cboObj As ComboBox, ByVal strName As String)
'    Dim lngIdx As Integer
    Dim i As Integer
    
'    Select Case strName
'    Case "西成药"
'        lngIdx = lng西成药
'    Case "中成药"
'        lngIdx = lng中成药
'    Case "中草药"
'        lngIdx = lng中草药
'    End Select
    
    With rs收入项目
        .MoveFirst
        Do While Not .EOF
            cboObj.AddItem !名称
            cboObj.ItemData(cboObj.NewIndex) = !ID
            .MoveNext
        Loop

        For i = 0 To cboObj.ListCount - 1
            If strName = "西成药" Then
                If cboObj.List(i) Like "*西药*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
            If strName = "中成药" Then
                If cboObj.List(i) Like "*中成药*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
            If strName = "中草药" Then
                If cboObj.List(i) Like "*中药*" Or cboObj.List(i) Like "*草药*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
    End With
End Sub


Private Sub txt数字码_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub udg数字码_Change()
    txt数字码.Text = udg数字码.Value
End Sub
