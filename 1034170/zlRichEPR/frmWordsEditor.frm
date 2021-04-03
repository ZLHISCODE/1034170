VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWordsEditor 
   AutoRedraw      =   -1  'True
   Caption         =   "词句选择"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9360
   Icon            =   "frmWordsEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   3465
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   3465
      MousePointer    =   7  'Size N S
      TabIndex        =   10
      Top             =   3150
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   3375
      TabIndex        =   9
      Top             =   2865
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   4125
      MousePointer    =   9  'Size W E
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   9360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6030
      Width           =   9360
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5865
         TabIndex        =   7
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   4770
         TabIndex        =   6
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.Frame fraUD 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3465
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   3765
      Width           =   5475
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "详细内容"
         Height          =   180
         Left            =   105
         TabIndex        =   12
         Top             =   30
         Width           =   720
      End
   End
   Begin RichTextLib.RichTextBox rtfSentence 
      Height          =   1245
      Left            =   3555
      TabIndex        =   2
      Top             =   4680
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   2196
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmWordsEditor.frx":058A
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   3285
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   2400
      Left            =   3390
      TabIndex        =   1
      Top             =   225
      Width           =   5760
      _cx             =   10160
      _cy             =   4233
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmWordsEditor.frx":0627
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgList 
         Left            =   420
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWordsEditor.frx":069C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWordsEditor.frx":0C36
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWordsEditor.frx":11D0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWordsEditor.frx":176A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWordsEditor.frx":1D04
            Key             =   "Expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5865
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   10345
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line lin 
      Index           =   1
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line lin 
      Index           =   2
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line lin 
      Index           =   3
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3045
      Y2              =   3045
   End
   Begin VB.Line lin 
      Index           =   4
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Line lin 
      Index           =   5
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line lin 
      Index           =   6
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line lin 
      Index           =   7
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3165
      Y2              =   3165
   End
End
Attribute VB_Name = "frmWordsEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================================================
Public mblnShow As Boolean '该窗体是否正在显示
Private mstrInput As String
Private mstrSentence As String
Private mstrLike As String
Private mintType As Integer
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mblnOK As Boolean

Private mlngPreY As Long

Private mrsPati As New ADODB.Recordset

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strInput As String, Optional ByVal intType As Integer = 3) As String
    mstrSentence = ""
    mstrInput = strInput
    mintType = intType
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    
    On Error Resume Next
    Me.Show 1, frmParent
    Err.Clear: On Error GoTo 0
    
    If mblnOK Then
        ShowMe = mstrSentence
    Else
        ShowMe = mstrInput
    End If
End Function

Private Function ShowTree() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strMatch As String
    
    On Error GoTo errH
        
    Screen.MousePointer = 11
    
    strMatch = "f_Sentence_Matched(A.ID,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10])=1"
    strSQL = _
        " Select Max(Level) As 级数, ID, 上级id, 编码, 名称, 说明" & _
        " From 病历词句分类" & _
        " Start With ID In (" & _
        "   Select A.分类id From 病历词句分类 B, 病历词句示范 A" & _
        "   Where A.分类id = B.ID And Nvl(Substr(B.范围, [1], 1), '0') = '1' And " & strMatch & _
        "   And ((Nvl(A.通用级, 0) = 0" & _
        "       Or A.通用级 = 1 And A.科室id In(Select A.部门id From 部门人员 A, 上机人员表 B Where A.人员id = B.人员id And B.用户名 = User)" & _
        "       Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User)))" & _
        "   Group By A.分类id)" & _
        " Connect By Prior 上级id = ID" & _
        " Group By ID, 上级id, 编码, 名称, 说明" & _
        " Order By 级数 Desc, 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
        CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "")
    
    
    tvw_s.Nodes.Clear
    Set objNode = tvw_s.Nodes.Add(, , "_", "所有词句", "Close")
    objNode.ExpandedImage = "Expend"
    objNode.Expanded = True
    
    Do While Not rsTmp.EOF
        Set objNode = tvw_s.Nodes.Add("_" & NVL(rsTmp!上级ID), tvwChild, "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Close")
        objNode.ExpandedImage = "Expend"
        'objNode.Expanded = True
        
        rsTmp.MoveNext
    Loop

    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Selected = True
    End If
    If Not tvw_s.SelectedItem Is Nothing Then
        tvw_s.SelectedItem.EnsureVisible
    End If
    
    Screen.MousePointer = 0
    ShowTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowList(Optional ByVal lng分类id As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strMatch As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    strMatch = "f_Sentence_Matched(A.ID,[2],[3],[4],[5],[6],[7],[8],[9],[10],[11])=1"
    If lng分类id <> 0 Then
        '按树形读取数据
        strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Trim(B.内容文本) as 内容文本" & _
            " From 病历词句组成 B,病历词句示范 A" & _
            " Where A.ID=B.词句ID(+) And B.排列次序(+)=1 And A.分类ID=[1] And " & strMatch & _
            "   And ((Nvl(A.通用级, 0) = 0" & _
            "       Or A.通用级 = 1 And A.科室id In(Select A.部门id From 部门人员 A, 上机人员表 B Where A.人员id = B.人员id And B.用户名 = User)" & _
            "       Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User))) Order by A.编号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng分类id, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
            CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "")
    Else
        '按输入读取数据
        strSQL = "Select A.ID,A.编号,A.名称,A.通用级,LPad(B.排列次序,3,'0')||Trim(B.内容文本) as 内容文本" & _
            " From 病历词句分类 C,病历词句组成 B,病历词句示范 A" & _
            " Where A.ID=B.词句ID And Nvl(B.内容性质,0)=0 And A.分类ID=C.ID And Nvl(Substr(C.范围, [1], 1), '0') = '1'" & _
            "   And (A.编号 Like [1]||'%'" & _
            "       Or A.名称 Like " & IIf(mstrLike <> "", "'%'||", "") & "[1]||'%'" & _
            "       Or B.内容文本 Like " & IIf(mstrLike <> "", "'%'||", "") & "[1]||'%')" & _
            "   And ((Nvl(A.通用级, 0) = 0" & _
            "       Or A.通用级 = 1 And A.科室id In(Select A.部门id From 部门人员 A, 上机人员表 B Where A.人员id = B.人员id And B.用户名 = User)" & _
            "       Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User)))"
        
        strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Substr(Min(A.内容文本),4) as 内容文本" & _
            " From (" & strSQL & ") A Where " & strMatch & " Group by A.ID,A.编号,A.名称,A.通用级 Order by A.编号"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstrInput, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
            CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "")
    End If
        
    vsList.Redraw = flexRDNone
    vsList.Rows = vsList.FixedRows
    
    If Not rsTmp.EOF Then
        vsList.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            vsList.RowData(i) = Val(rsTmp!ID)
            vsList.TextMatrix(i, 1) = rsTmp!编号
            vsList.TextMatrix(i, 2) = rsTmp!名称
            vsList.TextMatrix(i, 3) = NVL(rsTmp!内容文本)
            vsList.Cell(flexcpPicture, i, 0) = imgList.ListImages(NVL(rsTmp!通用级, 0) + 1).Picture
            
            rsTmp.MoveNext
        Next
        vsList.Cell(flexcpPictureAlignment, 1, 0, vsList.Rows - 1, 0) = 4
        vsList.Row = 1: vsList.Col = 2
    End If
    vsList.Redraw = flexRDDirect
    
    Screen.MousePointer = 0
    ShowList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If rtfSentence.Text = "" Then
        MsgBox "没有可用的词句内容。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrSentence = rtfSentence.Text
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String, i As Long
    Dim vRect As RECT, lngMaxH As Long
    
    mblnShow = True
    mblnOK = False
    mstrSentence = ""
    Me.rtfSentence.Text = mstrInput
    
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gstrSQL = "Select B.主页ID as 就诊ID,A.性别,Nvl(B.婚姻状况,A.婚姻状况) as 婚姻状况," & _
        " B.住院目的,B.当前病况 as 病人病情,B.入院方式" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlng病人ID, mlng主页ID)
    '读取词句数据
    Call ShowTree
    
    '界面显示处理
    Call RestoreWinState(Me, App.ProductName, IIf(mstrInput <> "", 1, 0))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    tvw_s.Left = 0
    tvw_s.Top = 0
    tvw_s.Height = Me.ScaleHeight - picBottom.Height
    
    fraLR.Left = tvw_s.Left + tvw_s.Width
    fraLR.Top = 0
    fraLR.Height = tvw_s.Height
    
    vsList.Top = 0
    vsList.Left = fraLR.Left + fraLR.Width
    vsList.Height = Me.ScaleHeight - rtfSentence.Height - fraUD.Height - picBottom.Height
    vsList.Width = Me.ScaleWidth - fraLR.Width - tvw_s.Width
    
    fraUD.Top = vsList.Top + vsList.Height
    fraUD.Left = vsList.Left
    fraUD.Width = vsList.Width
    
    rtfSentence.Top = fraUD.Top + fraUD.Height
    rtfSentence.Left = vsList.Left
    rtfSentence.Width = vsList.Width
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    
    Call SaveWinState(Me, App.ProductName, IIf(mstrInput <> "", 1, 0))
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 1 Then
            If Me.Width + x < 4000 Or Me.Width + x > 9600 Then Exit Sub
            Me.Width = Me.Width + x
        ElseIf Index = 2 Then
            If Me.Height + y < rtfSentence.Height * 2 Or Me.Height + y > 7200 Then Exit Sub
            Me.Height = Me.Height + y
        End If
        Call Form_Resize
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or vsList.Width - x < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + x
        tvw_s.Width = tvw_s.Width + x
        
        vsList.Left = vsList.Left + x
        vsList.Width = vsList.Width - x
        
        fraUD.Left = fraUD.Left + x
        fraUD.Width = fraUD.Width - x
        
        rtfSentence.Left = rtfSentence.Left + x
        rtfSentence.Width = rtfSentence.Width - x
        
        Me.Refresh
    End If
End Sub

Private Sub fraUD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlngPreY = y
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If vsList.Height + (y - mlngPreY) < 1000 Or rtfSentence.Height - (y - mlngPreY) < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + (y - mlngPreY)
        vsList.Height = vsList.Height + (y - mlngPreY)
        rtfSentence.Top = rtfSentence.Top + (y - mlngPreY)
        rtfSentence.Height = rtfSentence.Height - (y - mlngPreY)
        
        Me.Refresh
    End If
End Sub

Private Sub picBottom_GotFocus()
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    
    If picBottom.ScaleWidth - cmdCancel.Width * 2 < 3500 Then Exit Sub
    cmdCancel.Left = picBottom.ScaleWidth - cmdCancel.Width * 2
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
End Sub

Private Sub rtfSentence_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Val(Mid(Node.Key, 2)) <> 0 Then
        Call ShowList(Val(Mid(Node.Key, 2)))
    Else
        vsList.Rows = vsList.FixedRows
    End If
End Sub

Private Sub vsList_DblClick()
    With vsList
        If .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call LoadWords
        End If
    End With
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call vsList_DblClick
    End If
End Sub

Private Sub LoadWords()
    Dim lngStart As Long, lngStart_LAST As Long
    Dim strText As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsValue As New ADODB.Recordset
    On Error GoTo errHand
    
    lngStart_LAST = rtfSentence.SelStart
    If lngStart_LAST = 0 Then lngStart_LAST = Len(rtfSentence.Text)
    rtfSentence.Tag = rtfSentence.Text
    
    gstrSQL = "Select 内容性质,内容文本,要素名称,要素单位 From 病历词句组成 Where 词句ID=[1] Order by 排列次序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, Val(vsList.RowData(vsList.Row)))
    
    rtfSentence.Text = ""
    Do While Not rsTemp.EOF
        lngStart = Len(rtfSentence.Text)
        rtfSentence.SelStart = lngStart
        rtfSentence.SelLength = 0
        Select Case rsTemp!内容性质
        Case 0 '自由文字
            strText = NVL(rsTemp!内容文本)
            With rtfSentence
                .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                .SelUnderline = False
            End With
        Case 1, 2 '1-临时诊治要素,2-固定诊治要素
            If Not IsNull(rsTemp!内容文本) Then
                strText = rsTemp!内容文本
            Else
                strText = ""
                gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4]) as 内容 From Dual"
                Set rsValue = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, CStr(rsTemp!要素名称), mlng病人ID, mlng主页ID, 2)
                If Not rsTemp.EOF Then strText = IIf(Not IsNull(rsValue!内容), rsValue!内容 & NVL(rsTemp!要素单位), "")
                If strText = "" Then strText = "{" & rsTemp!要素名称 & "}" & NVL(rsTemp!要素单位)
            End If
            With rtfSentence
                .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                .SelUnderline = True
            End With
        End Select
        rsTemp.MoveNext
    Loop
    
    rtfSentence.Text = Mid(rtfSentence.Tag, 1, lngStart_LAST) & rtfSentence.Text & Mid(rtfSentence.Tag, lngStart_LAST + 1)
    rtfSentence.SelStart = lngStart_LAST
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
