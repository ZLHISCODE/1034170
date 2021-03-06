VERSION 5.00
Begin VB.Form frmSet成都郊县 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmSet成都郊县.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   90
      TabIndex        =   5
      Top             =   1230
      Width           =   4275
   End
   Begin VB.TextBox txtIC端口号 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3810
      TabIndex        =   4
      Text            =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.ComboBox cbo卡类型 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   750
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   7
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1770
      TabIndex        =   6
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CheckBox Chk入院信息 
      Caption         =   "入院登记的同时，上传医保病人入院信息(&1)"
      Height          =   345
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblIC端口号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "端口号"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3180
      TabIndex        =   3
      Top             =   810
      Width           =   540
   End
   Begin VB.Label lbl卡类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡类型"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   1
      Top             =   810
      Width           =   540
   End
End
Attribute VB_Name = "frmSet成都郊县"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnOK As Boolean

Public Function ShowSet() As Boolean
    blnOK = False
    
    Me.Show 1
    ShowSet = blnOK
End Function

Private Sub cbo卡类型_Click()
    Me.lblIC端口号.Enabled = (cbo卡类型.ListIndex <> 0)
    Me.txtIC端口号.Enabled = (cbo卡类型.ListIndex <> 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "zl_保险参数_Delete(" & gintInsure & ",NULL)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & gintInsure & ",NULL,'上传入院信息'," & Chk入院信息.Value & ",1)", , adCmdStoredProc
    gcnOracle.CommitTrans
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "卡类型", Me.cbo卡类型.ListIndex)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", txtIC端口号.Text)
    blnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    
    '增加初始化数据
    Me.cbo卡类型.Clear
    Me.cbo卡类型.AddItem "磁卡"
    Me.cbo卡类型.AddItem "IC卡-JKP428"
    Me.cbo卡类型.AddItem "IC卡-ICIOX"
    Me.cbo卡类型.ListIndex = 0
    
    '将以前的参数取出来显示在界面中
    gstrSQL = "Select 参数名,Nvl(参数值,0) Value From 保险参数 Where 险类=22 "
    Call OpenRecordset(rsTmp, "获取上传入院信息参数值")
    With rsTmp
        Do While Not rsTmp.EOF
            Select Case !参数名
            Case "上传入院信息"
                Chk入院信息.Value = rsTmp!Value
            End Select
            .MoveNext
        Loop
    End With
    
    cbo卡类型.ListIndex = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "卡类型", 0)
    txtIC端口号.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", 1)
End Sub
