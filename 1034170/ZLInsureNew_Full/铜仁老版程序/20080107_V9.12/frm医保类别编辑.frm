VERSION 5.00
Begin VB.Form frm医保类别编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保类别编辑"
   ClientHeight    =   3315
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   5640
   Icon            =   "frm医保类别编辑.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdGet 
      Caption         =   "…"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2160
      Width           =   315
   End
   Begin VB.CheckBox chk禁止 
      Caption         =   "本系统中禁止使用(&S)"
      Height          =   225
      Left            =   1215
      TabIndex        =   10
      Top             =   2910
      Width           =   2025
   End
   Begin VB.CheckBox chk中心 
      Caption         =   "具有多个医保中心(&R)"
      Height          =   225
      Left            =   1215
      TabIndex        =   9
      Top             =   2571
      Width           =   2025
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4320
      TabIndex        =   13
      Top             =   2745
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1215
      MaxLength       =   12
      TabIndex        =   7
      Top             =   2157
      Width           =   2430
   End
   Begin VB.TextBox txtEdit 
      Height          =   1080
      Index           =   2
      Left            =   1215
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   963
      Width           =   2790
   End
   Begin VB.Frame Frame1 
      Height          =   3570
      Left            =   4155
      TabIndex        =   14
      Top             =   -195
      Width           =   30
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1215
      MaxLength       =   3
      TabIndex        =   1
      Top             =   135
      Width           =   555
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1215
      MaxLength       =   20
      TabIndex        =   3
      Top             =   549
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4320
      TabIndex        =   12
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4320
      TabIndex        =   11
      Top             =   150
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "医院编号(&B)"
      Height          =   180
      Index           =   3
      Left            =   210
      TabIndex        =   6
      Top             =   2217
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "应用说明(&E)"
      Height          =   180
      Index           =   2
      Left            =   195
      TabIndex        =   4
      Top             =   1020
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "医保序号(&S)"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "医保名称(&N)"
      Height          =   180
      Index           =   1
      Left            =   195
      TabIndex        =   2
      Top             =   609
      Width           =   990
   End
End
Attribute VB_Name = "frm医保类别编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum编辑
    Text序号 = 0
    Text名称 = 1
    Text说明 = 2
    Text医院编码 = 3
End Enum

Dim mstr序号 As String         '当前编辑的医保类别序号
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub cmdGet_Click()
    Dim strReturn As String
    
    If mstr序号 = "10" Then
        strReturn = 医院编码_重庆
        If strReturn <> "" Then
            txtEdit(Text医院编码) = strReturn
            mblnChange = True
        End If
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
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
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    
    MousePointer = vbHourglass
    If Save医保类别() = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    MousePointer = vbDefault
    
    mblnOK = True
    mblnChange = False
    
    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能:分析输入有关医保类别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim lngIndex As Integer
    Dim strTemp As String
    For lngIndex = Text序号 To Text医院编码
        If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
            txtEdit(lngIndex).SetFocus
            zlControl.TxtSelAll txtEdit(lngIndex)
            Exit Function
        End If
        
        If lngIndex = Text序号 Or lngIndex = Text名称 Then
            If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                txtEdit(lngIndex).Text = ""
                MsgBox "序号或名称都不能为空。", vbExclamation, gstrSysName
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    
    If txtEdit(Text序号).Enabled = True Then
        If IsNumeric(txtEdit(Text序号)) = False Or Val(txtEdit(Text序号).Text) <= 900 Then
            MsgBox "序号只能是大于900的整数。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text序号)
            txtEdit(Text序号).SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Function Save医保类别() As Boolean
'功能:保存编辑的内容到医保类别表中
'参数:
'返回值:成功返回True,否则为False
    Dim lng序号 As Long
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    If mstr序号 = "" Then       '新增一条记录
        lng序号 = Val(txtEdit(Text序号).Text)
        gstrSQL = "zl_保险类别_Insert(" & lng序号 & _
            ",'" & txtEdit(Text名称).Text & "','" & txtEdit(Text说明).Text & _
            "','" & txtEdit(Text医院编码).Text & "'," & chk中心.Value & "," & chk禁止.Value & ")"
        Call ExecuteProcedure(Me.Caption)
        
        If chk中心.Value = 0 Then
            '单中心医保预先就增加中心
            gstrSQL = "zl_保险中心目录_Insert(" & lng序号 & ",0,'1','" & txtEdit(Text名称).Text & "')"
            Call ExecuteProcedure(Me.Caption)
        End If
    Else    '修改
        gstrSQL = "zl_保险类别_Update(" & mstr序号 & _
            ",'" & txtEdit(Text名称).Text & "','" & txtEdit(Text说明).Text & _
            "','" & txtEdit(Text医院编码).Text & "'," & chk禁止.Value & ")"
        Call ExecuteProcedure(Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    
    '在主界面上做相应的调整
    If mstr序号 = "" Then
        '新增
        Set lst = frm医保类别.lvwKind_S.ListItems.Add(, "K" & lng序号, " ", "Common", "Common")
        lst.Selected = True
        lst.EnsureVisible
    Else
        '修改
        Set lst = frm医保类别.lvwKind_S.SelectedItem
    End If
    lst.Text = txtEdit(Text名称).Text
    lst.SubItems(1) = txtEdit(Text序号).Text
    lst.SubItems(2) = txtEdit(Text医院编码).Text
    lst.SubItems(3) = txtEdit(Text说明).Text
    lst.Tag = chk中心.Value
    lst.Ghosted = (chk禁止.Value = 1)
    
    Save医保类别 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Public Function 编辑医保类别(ByVal str序号 As String) As Boolean
'功能:用来与调用的医保类别管理窗口进行通讯的程序
'参数:str序号           当前编辑的医保类别的的序号
'返回值:编辑成功返回True,否则为False
    Dim rs医保类别 As New ADODB.Recordset
    Dim i As Integer
    
    mstr序号 = str序号
    If str序号 = "10" Then
        If 医保初始化_重庆 = True Then
            cmdGet.Enabled = True
        End If
    End If
    
    mblnOK = False
    
    rs医保类别.CursorLocation = adUseClient
    
    If str序号 <> "" Then
        gstrSQL = "Select 序号,名称,说明,医院编码,具有中心,是否禁止" & _
            " From 保险类别  Where 序号=" & str序号
        Call OpenRecordset(rs医保类别, Me.Caption)
        
        txtEdit(Text序号).Text = rs医保类别("序号")
        txtEdit(Text名称).Text = rs医保类别("名称")
        txtEdit(Text说明).Text = IIf(IsNull(rs医保类别("说明")), "", rs医保类别("说明"))
        txtEdit(Text医院编码).Text = IIf(IsNull(rs医保类别("医院编码")), "", rs医保类别("医院编码"))
        chk中心.Value = IIf(rs医保类别("具有中心") = 1, 1, 0)
        chk禁止.Value = IIf(rs医保类别("是否禁止") = 1, 1, 0)
        
        lblEdit(Text序号).Enabled = False
        txtEdit(Text序号).Enabled = False
        chk中心.Enabled = False
    Else
        txtEdit(Text序号).Text = zlDatabase.GetMax("保险类别", "序号", 3)
        If Val(txtEdit(Text序号).Text) < 901 Then txtEdit(Text序号).Text = 901
    End If
    
    mblnChange = False
    frm医保类别编辑.Show vbModal
    编辑医保类别 = mblnOK
End Function

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text名称, Text说明
          zlCommFun.OpenIme True
        Case Text序号, Text医院编码
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 '使之不响
        SendKeys "{Tab}", 1
    Else
        If Index = Text序号 Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub chk中心_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub chk禁止_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub


