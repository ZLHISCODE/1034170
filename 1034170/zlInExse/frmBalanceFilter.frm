VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBalanceFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   5145
      TabIndex        =   12
      Top             =   1740
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   105
      TabIndex        =   14
      Top             =   0
      Width           =   4920
      Begin VB.TextBox txtClinic 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1920
         Width           =   1470
      End
      Begin VB.CheckBox chkFeeOrigin 
         Caption         =   "其它"
         Height          =   255
         Index           =   3
         Left            =   3735
         TabIndex        =   26
         Top             =   2777
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkFeeOrigin 
         Caption         =   "体检"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   25
         Top             =   2777
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkFeeOrigin 
         Caption         =   "住院"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   24
         Top             =   2777
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkFeeOrigin 
         Caption         =   "门诊"
         Height          =   255
         Index           =   0
         Left            =   975
         TabIndex        =   23
         Top             =   2777
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkType 
         Caption         =   "作废记录"
         Height          =   210
         Index           =   1
         Left            =   3360
         TabIndex        =   21
         Top             =   720
         Width           =   1020
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1920
         Width           =   1470
      End
      Begin VB.CheckBox chkType 
         Caption         =   "结帐记录"
         Height          =   210
         Index           =   0
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2340
         Width           =   1470
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1098
         Width           =   1470
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1098
         Width           =   1470
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   960
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2340
         Width           =   1470
      End
      Begin VB.TextBox txtFactBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         TabIndex        =   4
         Top             =   1512
         Width           =   1470
      End
      Begin VB.TextBox txtFactEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3240
         TabIndex        =   5
         Top             =   1512
         Width           =   1470
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   975
         TabIndex        =   1
         Top             =   684
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   199819267
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   975
         TabIndex        =   0
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   199819267
         CurrentDate     =   36588
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2760
         TabIndex        =   30
         Top             =   1572
         Width           =   180
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员"
         Height          =   180
         Left            =   2580
         TabIndex        =   29
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2760
         TabIndex        =   28
         Top             =   1158
         Width           =   180
      End
      Begin VB.Label lblClinicNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   2580
         TabIndex        =   27
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费用来源"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   2814
         Width           =   720
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   180
         TabIndex        =   20
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   744
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   360
         TabIndex        =   18
         Top             =   1158
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   540
         TabIndex        =   17
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   1572
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5145
      TabIndex        =   13
      Top             =   810
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5145
      TabIndex        =   11
      Top             =   390
      Width           =   1100
   End
End
Attribute VB_Name = "frmBalanceFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrFilter As String
Public mblnDateMoved As Boolean '当前所选条件的数据是否在后备数据表中
Public mstr来源 As String

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlcontrol.CboMatchIndex(cbo操作员.hWnd, KeyAscii)
        If lngIdx = -1 And cbo操作员.ListCount > 0 Then lngIdx = 0
        cbo操作员.ListIndex = lngIdx
    End If
End Sub

Private Sub chkFeeOrigin_Click(Index As Integer)
    If chkFeeOrigin(0).Value = 0 And chkFeeOrigin(1).Value = 0 And chkFeeOrigin(2).Value = 0 And chkFeeOrigin(3).Value = 0 Then
        chkFeeOrigin(Index).Value = 1
    End If
End Sub

Private Sub chkType_Click(Index As Integer)
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then chkType(Index).Value = 1
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub


Private Sub cmdOK_Click()
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "结束票据号不能小于开始票据号！", vbInformation, gstrSysName
            txtFactEnd.SetFocus: Exit Sub
        End If
    End If
    
    If DateDiff("d", dtpBegin.Value, dtpEnd.Value) > 30 Then
        If txt姓名.Text = "" And txt住院号.Text = "" And txtClinic.Text = "" Then
            If MsgBox("过滤的时间范围超过了三十天,读取数据可能会比较耗时,是否继续?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
                Exit Sub
            End If
        End If
    End If
    
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim curDate As Date, i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txt住院号.Text = ""
    txt姓名.Text = ""
    
    curDate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(curDate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    
    '操作员
    cbo操作员.Clear
    cbo操作员.AddItem "所有结帐人"
    cbo操作员.ListIndex = 0
    Set rsTmp = GetPersonnel("住院结帐员")
    For i = 1 To rsTmp.RecordCount
        cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
        rsTmp.MoveNext
    Next
    cbo.SetListWidthAuto cbo操作员, zlcontrol.OneCharWidth(cbo操作员.Font) * 70 / cbo操作员.Width
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFactBegin_GotFocus()
    SelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    SelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    SelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlcontrol.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 15)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 15)
End Sub

Private Sub txtNoEnd_GotFocus()
    SelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlcontrol.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Sub MakeFilter()
    Dim strSQL As String, strSQLtmp As String, i As Integer
    
    mstrFilter = " And A.收费时间 Between [1] And [2]"
    
    mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[3]"
    End If
    
    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
       '无需根据票据号判断,直接根据单据的登记时间判断
       strSQLtmp = IIf(txtFactEnd.Text = "", " =[5]", " Between [5] And [6]")
       
       If mblnDateMoved Then
           strSQL = "" & _
            "(  Select A.NO" & _
            "   From 票据打印内容 A,票据使用明细 B" & _
            "   Where A.数据性质=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.打印ID And B.票种=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.性质=1" & _
            "         And B.号码 " & strSQLtmp & ")  Union All" & _
            " (Select A.NO " & _
            " From H票据打印内容 A,H票据使用明细 B" & _
            " Where A.数据性质=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.打印ID And B.票种=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.性质=1" & _
            " And B.号码 " & strSQLtmp & ")"
       Else
           strSQL = "Select A.NO" & _
           " From 票据打印内容 A,票据使用明细 B" & _
           " Where A.数据性质=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.打印ID And B.票种=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.性质=1" & _
           " And B.号码 " & strSQLtmp
       End If
    End If
    If strSQL <> "" Then mstrFilter = mstrFilter & " And A.NO IN(" & strSQL & ")"
    
    
    If txt住院号.Text <> "" Then
        mstrFilter = mstrFilter & " And C.病人ID in (Select 病人ID From 病案主页 where 住院号=[7])"
    End If
    
    '问题65105,刘尔旋:过滤增加门诊号条件
    If txtClinic.Text <> "" Then
        mstrFilter = mstrFilter & " And C.病人ID in (Select 病人ID From 病人信息 where 门诊号=[10]) And (Nvl(A.结帐类型,0)=1 Or Nvl(A.结帐类型,0)=0)"
    End If
    
    If txt姓名.Text <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt姓名.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(C.姓名) Like [8]"
        Else
            mstrFilter = mstrFilter & " And C.姓名 Like [8]"
        End If
    End If
    
    If cbo操作员.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And A.操作员姓名||''=[9]"
    End If
    If Not (chkType(0).Value = 1 And chkType(1).Value = 1) Then
        If chkType(0).Value = 1 Then
            mstrFilter = mstrFilter & " And A.记录状态 IN(1,3)"
        Else
            mstrFilter = mstrFilter & " And A.记录状态=2"
        End If
    End If
        
    mstr来源 = ""
    For i = 0 To chkFeeOrigin.Count - 1
        mstr来源 = mstr来源 & IIf(chkFeeOrigin(i).Value = 1, 1, 0) '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    Next
 
    
End Sub

Private Sub txt姓名_GotFocus()
    SelAll txt姓名
End Sub

Private Sub txt住院号_GotFocus()
    SelAll txt住院号
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
