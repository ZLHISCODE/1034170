VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceStopTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "停止医嘱"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4170
   Icon            =   "frmAdviceStopTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -255
      TabIndex        =   3
      Top             =   1305
      Width           =   4845
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2625
      TabIndex        =   2
      Top             =   1470
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1485
      TabIndex        =   1
      Top             =   1470
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   645
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   96862211
      UpDown          =   -1  'True
      CurrentDate     =   39668.3388888889
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行终止时间"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   375
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frmAdviceStopTime.frx":058A
      Top             =   255
      Width           =   720
   End
End
Attribute VB_Name = "frmAdviceStopTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsAdvice As ADODB.Recordset
Private mlng医嘱ID As String
Private mblnOK As Boolean

Private mstrTime As String

Public Function ShowMe(frmParent As Object, ByVal lng医嘱ID As Long) As String
    mlng医嘱ID = lng医嘱ID
    Me.Show 1, frmParent
    If mblnOK Then
        ShowMe = mstrTime
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '检查合法性
    '必须大于开始执行时间
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    If Format(dtpTime.value, "yyyy-MM-dd HH:mm") <= Format(mrsAdvice!开始执行时间, "yyyy-MM-dd HH:mm") Then
        MsgBox "输入的执行终止时间必须大于医嘱的开始执行时间 " & Format(mrsAdvice!开始执行时间, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
        dtpTime.SetFocus: Exit Sub
    End If
    '登记执行时间>上次执行时间
    mstrTime = GetAdviceStopTime(mlng医嘱ID)
    If mstrTime <> "" Then
        If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mstrTime, "yyyy-MM-dd HH:mm") Then
            MsgBox "不能停止到执行时间 " & mstrTime & " 之前，请调整停止时间，如果确实要停止到执行时间之前，请先取消执行登记。", vbInformation, gstrSysName
            dtpTime.SetFocus: Exit Sub
        End If
    End If
    '不应小于上次执行时间
    If Not IsNull(mrsAdvice!上次执行时间) Then
        If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm") Then
            If MsgBox("输入的执行终止时间小于医嘱的上次执行时间 " & Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm") & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                dtpTime.SetFocus: Exit Sub
            End If
        End If
    End If
    
    mstrTime = Format(dtpTime.value, "yyyy-MM-dd HH:mm")
    mblnOK = True
    Unload Me
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call cmdOK_Click
End Sub

Private Sub Form_Activate()
    dtpTime.SetFocus
    Me.Refresh
    zlCommFun.PressKey vbKeyRight
    zlCommFun.PressKey vbKeyRight
    zlCommFun.PressKey vbKeyRight
End Sub

Private Sub Form_Load()
    Dim datCurr As Date
    Dim strSQL As String
    
    mblnOK = False
    datCurr = zlDatabase.Currentdate
    
    On Error GoTo errH
    strSQL = "Select 开始执行时间,执行终止时间,上次执行时间,开嘱时间 From 病人医嘱记录 Where ID=[1]"
    Set mrsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
    
    'If IsNull(mrsAdvice!执行终止时间) Then
    If gbln长期医嘱次日生效 Then
        dtpTime.value = CDate(Format(datCurr + 1, "yyyy-MM-dd 00:00"))
    Else
        dtpTime.value = CDate(Format(datCurr, "yyyy-MM-dd HH:mm"))
    End If
    If Not IsNull(mrsAdvice!上次执行时间) Then
        If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm") Then
            dtpTime.value = Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm")
        End If
    End If
'    Else
'        dtpTime.Value = Format(mrsAdvice!执行终止时间, "yyyy-MM-dd HH:mm")
'    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
        Set mrsAdvice = Nothing
    End If
End Sub
