VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCISBorrowPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmCISBorrowPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   2
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   1
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4935
      TabIndex        =   0
      Top             =   6000
      Width           =   1100
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   0
      Left            =   255
      ScaleHeight     =   5415
      ScaleWidth      =   5880
      TabIndex        =   3
      Top             =   330
      Width           =   5880
      Begin VB.CheckBox chkBorrowAccount 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2685
         TabIndex        =   18
         Top             =   3495
         Width           =   660
      End
      Begin VB.CheckBox chkBorrowReason 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2685
         TabIndex        =   16
         Top             =   3105
         Width           =   660
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   2460
         TabIndex        =   14
         Top             =   2670
         Width           =   1920
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   2460
         TabIndex        =   12
         Top             =   2265
         Width           =   1920
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   3
         Left            =   930
         TabIndex        =   9
         Top             =   1920
         Width           =   4815
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   870
         Width           =   1920
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   930
         TabIndex        =   5
         Top             =   165
         Width           =   4815
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "允许自由录入借阅原因"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   825
         TabIndex        =   19
         Top             =   3540
         Width           =   1800
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "必须录入借阅申请理由"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   825
         TabIndex        =   17
         Top             =   3150
         Width           =   1800
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "借阅时的最长期限为                      天"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   825
         TabIndex        =   15
         Top             =   2730
         Width           =   3780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "借阅时的期限缺省为                      天"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   825
         TabIndex        =   13
         Top             =   2325
         Width           =   3780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "其他设置"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   16
         Left            =   180
         TabIndex        =   11
         Top             =   1905
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "缺省时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   195
         TabIndex        =   10
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "查询电子病案借阅申请的缺省时间范围。"
         Height          =   405
         Left            =   1035
         TabIndex        =   8
         Top             =   555
         Width           =   4065
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmCISBorrowPara.frx":000C
         Top             =   390
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省范围(&1)"
         Height          =   180
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Top             =   930
         Width           =   990
      End
   End
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   5790
      Left            =   105
      TabIndex        =   4
      Top             =   30
      Width           =   5970
      _Version        =   589884
      _ExtentX        =   10530
      _ExtentY        =   10213
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmCISBorrowPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mblnOK As Boolean
Private mfrmMain As Object
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mstrPrivs As String
Private mblnBorrowAccount As Boolean '允许自由录入借阅原因

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始数据") = False Then Exit Function
    If ExecuteCommand("读取参数") = False Then Exit Function
    
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim blnAllowModify As Boolean

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "初始数据"
            With tbc
                With .PaintManager
                    .Appearance = xtpTabAppearancePropertyPage2003
                    .BoldSelected = True
                    .COLOR = xtpTabColorDefault
                    .ColorSet.ButtonSelected = COLOR.白色
                    .ShowIcons = True
                End With
                
                .InsertItem 0, "基本 ", picPane(0).Hwnd, 0
                .Item(0).Selected = True
            End With
            

            cbo(0).Clear
            cbo(0).AddItem "今  天"
            cbo(0).AddItem "昨  天"
            cbo(0).AddItem "本  周"
            cbo(0).AddItem "本  月"
            cbo(0).AddItem "本  季"
            cbo(0).AddItem "本半年"
            cbo(0).AddItem "本  年"
            cbo(0).AddItem "前三天"
            cbo(0).AddItem "前一周"
            cbo(0).AddItem "前半月"
            cbo(0).AddItem "前一月"
            cbo(0).AddItem "前二月"
            cbo(0).AddItem "前三月"
            cbo(0).AddItem "前半年"
            cbo(0).AddItem "前一年"
            cbo(0).AddItem "前二年"
            
'            lbl(3).ForeColor = COLOR.公共模块色
'            txt(7).ForeColor = COLOR.公共模块色
'
'            lbl(0).ForeColor = COLOR.公共模块色
'            txt(0).ForeColor = COLOR.公共模块色
            
        '--------------------------------------------------------------------------------------------------------------
        Case "控件状态"
            
'            blnAllowModify = IsPrivs(mstrPrivs, "参数设置")
'            lbl(3).Enabled = blnAllowModify
'            txt(7).Enabled = blnAllowModify
'            lbl(0).Enabled = blnAllowModify
'            txt(0).Enabled = blnAllowModify
        '--------------------------------------------------------------------------------------------------------------
        Case "读取参数"
            
            On Error Resume Next
            cbo(0).Text = zlDatabase.GetPara("登记缺省范围", ParamInfo.系统号, mfrmMain.模块号, "今  天", Array(cbo(0)), IsPrivs(mstrPrivs, "参数设置"))
            On Error GoTo errHand
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            txt(7).Text = Val(zlDatabase.GetPara("病案借阅期限", ParamInfo.系统号, mfrmMain.模块号, "7", Array(txt(7)), IsPrivs(mstrPrivs, "参数设置")))
            txt(0).Text = Val(zlDatabase.GetPara("借阅最长期限", ParamInfo.系统号, mfrmMain.模块号, "30", Array(txt(0)), IsPrivs(mstrPrivs, "参数设置")))
            chkBorrowReason.Value = zlDatabase.GetPara("必须录入借阅原因", ParamInfo.系统号, mfrmMain.模块号, "0", chkBorrowReason, IsPrivs(mstrPrivs, "参数设置"))
            chkBorrowAccount.Value = zlDatabase.GetPara("允许自由录入借阅原因", ParamInfo.系统号, mfrmMain.模块号, "0", chkBorrowAccount, IsPrivs(mstrPrivs, "参数设置"))
            
        '--------------------------------------------------------------------------------------------------------------
        Case "校验数据"
            
            If Val(txt(0).Text) < Val(txt(7).Text) Then
                ShowSimpleMsg "借阅的最长期限不能小于病案借阅的缺省期限!"
                Exit Function
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case "保存数据"
            
            Call SetPara("登记缺省范围", cbo(0).Text, mfrmMain.模块号, IsPrivs(mstrPrivs, "参数设置"))
            Call SetPara("病案借阅期限", Val(txt(7).Text), mfrmMain.模块号, IsPrivs(mstrPrivs, "参数设置"))
            Call SetPara("借阅最长期限", Val(txt(0).Text), mfrmMain.模块号, IsPrivs(mstrPrivs, "参数设置"))
            Call SetPara("必须录入借阅原因", chkBorrowReason.Value, mfrmMain.模块号, IsPrivs(mstrPrivs, "参数设置"))
            Call SetPara("允许自由录入借阅原因", chkBorrowAccount.Value, mfrmMain.模块号, IsPrivs(mstrPrivs, "参数设置"))
        End Select
    Next

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    cmdOK.Tag = IIf(blnData, "Changed", "")
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = (cmdOK.Tag = "Changed")
End Property

'######################################################################################################################

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkBorrowAccount_Click()
    DataChanged = True
End Sub

Private Sub chkBorrowReason_Click()
    DataChanged = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    
    If DataChanged Then
        If ExecuteCommand("校验数据") = False Then Exit Sub
        
        If ExecuteCommand("保存数据") Then
            
            DataChanged = False
            
            mblnOK = True
        Else
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("新增或修改的参数必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
    End If
    
    Set mclsVsf = Nothing
End Sub

Private Sub txt_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 5
        zlCommFun.OpenIme True
    End Select
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 5
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


