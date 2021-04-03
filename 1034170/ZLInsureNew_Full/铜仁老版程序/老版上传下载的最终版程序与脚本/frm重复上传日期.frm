VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm重复上传日期 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "重新上传时间设置"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frm重复上传日期.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   5
      Top             =   810
      Width           =   1100
   End
   Begin VB.Frame fraScope 
      Caption         =   "时间范围"
      Height          =   1455
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3405
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1830
         TabIndex        =   1
         Top             =   870
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   23789571
         CurrentDate     =   36279
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1830
         TabIndex        =   2
         Top             =   390
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   23789571
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.Label lblTimeStop 
         AutoSize        =   -1  'True
         Caption         =   "结束时间(&E)"
         Height          =   180
         Left            =   780
         TabIndex        =   4
         Top             =   930
         Width           =   990
      End
      Begin VB.Label lblTimeStart 
         AutoSize        =   -1  'True
         Caption         =   "开始时间(&B)"
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   450
         Width           =   990
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   150
         Picture         =   "frm重复上传日期.frx":000C
         Top             =   420
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   6
      Top             =   270
      Width           =   1100
   End
   Begin VB.Label lbl提示 
      Caption         =   "注意：本功能只在医保中心数据受到破坏，作数据恢复时使用。"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1860
      Width           =   3585
   End
End
Attribute VB_Name = "frm重复上传日期"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mdatBegin As Date, mdatEnd As Date

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "开始时间大于结束时间了。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("如果使用不当，可能破坏医保中心的数据。" & vbCrLf & "是否确定？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function GetTimeScope(datBegin As Date, datEnd As Date, ByVal datMax As Date) As Boolean
                
    dtpBegin.Value = datBegin
    dtpEnd.Value = datEnd
    
    '严格限制日期
    dtpBegin.MaxDate = datMax
    dtpEnd.MaxDate = dtpBegin.MaxDate
    
    frm重复上传日期.Show vbModal, frm上传下载
    
    GetTimeScope = mblnOK
    If mblnOK = True Then
        datBegin = mdatBegin
        datEnd = mdatEnd
    End If
End Function


