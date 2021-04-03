VERSION 5.00
Begin VB.Form frmBJCAJS 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3930
   Icon            =   "frmBJCAJS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3930
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkPara 
      BackColor       =   &H80000005&
      Caption         =   "不启用时间戳"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   3930
      TabIndex        =   0
      Top             =   1425
      Width           =   3930
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   2760
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmBJCAJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CMD_ENUM
    CMD_OK = 0
    CMD_CANCEL = 1
End Enum

Private Sub chkPara_Click()
    gudtPara.blnISTS = chkPara.Value = vbChecked
End Sub

Private Sub cmdPara_Click(Index As Integer)
    Dim lngID As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnOk As Boolean
    
    If Index = CMD_OK Then
        gstrPara = BJCAJS_SetParaStr
        On Error GoTo errH
        strSQL = "Select count(1) as RowCount  From zlParameters Where 系统 = [1] And Nvl(模块, 0) = 0 And 参数号 = 90000"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "电子签名参数", glngSys)
        If Not rsTmp.EOF Then
            If rsTmp!RowCount = 0 Then
                lngID = gobjComLib.zlDatabase.GetNextId("zlParameters")
                strSQL = "Insert Into zlParameters(ID, 系统, 模块, 参数号, 参数名, 参数值) Values (" & lngID & ", " & glngSys & ", Null, 90000, '电子签名参数','" & gstrPara & "')"
                Call gobjComLib.zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
                blnOk = True
            End If
        End If
        If Not blnOk Then
            Call gobjComLib.zlDatabase.SetPara(90000, gstrPara, glngSys)
        End If
    End If
    
    Unload Me
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Load()
    Call BJCAJS_GetPara
    chkPara.Value = gudtPara.blnISTS
End Sub
