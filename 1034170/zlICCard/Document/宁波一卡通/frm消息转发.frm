VERSION 5.00
Begin VB.Form frm消息转发 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "消息转发"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   Icon            =   "frm消息转发.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3810
      Top             =   2640
   End
   Begin VB.CommandButton cmd退出 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   405
      Left            =   6870
      TabIndex        =   5
      Top             =   5160
      Width           =   1275
   End
   Begin VB.CommandButton cmd开启服务 
      Caption         =   "开启服务(&S)"
      Height          =   405
      Left            =   5430
      TabIndex        =   4
      Top             =   5160
      Width           =   1275
   End
   Begin VB.TextBox txt请求内容 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   4275
      Left            =   1350
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   660
      Width           =   7065
   End
   Begin VB.Label lbl请求内容 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请求内容"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   2
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl操作说明 
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1350
      TabIndex        =   1
      Top             =   300
      Width           =   7665
   End
   Begin VB.Label lbl当前操作 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "当前操作"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frm消息转发"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnExit As Boolean
Dim blnProcess As Boolean

Private Sub cmd开启服务_Click()
    Timer1.Enabled = True
    cmd开启服务.Enabled = False
End Sub

Private Sub cmd退出_Click()
    '必须把当前的请求处理完后（保存到数据库中，并更新相应状态），才询问是否退出
    If blnProcess Then
        If MsgBox("正在与中心交互数据,是否退出？", vbQuestion + vbYesNo + vbDefaultButton2, "消息转发") = vbNo Then Exit Sub
    End If
    
    blnExit = True
    Unload Me
End Sub

Public Function 调用接口(ByVal str日期 As String, ByVal lng序列号 As Long, ByVal strFuncName As String, ByVal strURL As String) As Boolean
    Dim strSql As String
    Dim strSoapRequest As String
    Dim objHttp As MSXML2.XMLHTTP
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '标记为正在处理中
    Me.lbl操作说明.Caption = "正在向中心发送序列号=[" & lng序列号 & "]的请求，请稍候......"
    gstrSQL = "zl_消息主表_Insert('" & str日期 & "'," & lng序列号 & ",'" & strFuncName & "','" & strURL & "',NULL,1)"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '提取待发送数据
    strSql = " Select 发送数据 From 消息转发 Where 日期='" & str日期 & "' And 序列号=" & lng序列号 & " Order by 行号"
    Call OpenRecordset(rsTemp, "提取待发送数据", strSql)
    With rsTemp
        Do While Not .EOF
            strSoapRequest = strSoapRequest & IIf(IsNull(!发送数据), "", !发送数据)
            .MoveNext
        Loop
    End With
    Me.txt请求内容.Text = strSoapRequest
    
    '准备发送
    Set objHttp = New MSXML2.XMLHTTP
    objHttp.Open "post", strURL, False
    objHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objHttp.setRequestHeader "Content-Length", Len(strSoapRequest)
    objHttp.setRequestHeader "SOAPAction", strURL
    
    '根据返回的状态信息来判断是否成功
    objHttp.send (strSoapRequest)
    If objHttp.Status <> 200 Then GoTo errHand
    
    '将返回数据写入数据库
    Call WriteData(str日期, lng序列号, strFuncName, strURL, objHttp.responseText)
    Me.lbl操作说明.Caption = "序列号=[" & lng序列号 & "]的请求已完成"
    
    调用接口 = True
    Exit Function
errHand:
    If Err <> 0 Then
        gstrSQL = "zl_消息主表_Insert('" & str日期 & "'," & lng序列号 & ",'" & strFuncName & "','" & strURL & "','" & Err.Description & "',3)"
    Else
        gstrSQL = "zl_消息主表_Insert('" & str日期 & "'," & lng序列号 & ",'" & strFuncName & "','" & strURL & "','返回信息：[" & objHttp.Status & "]" & objHttp.responseText & "',3)"
    End If
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
End Function

Private Sub WriteData(ByVal str日期 As String, ByVal lng序列号 As Long, ByVal strFuncName As String, ByVal strURL As String, ByVal strData As String)
    Dim blnTrans As Boolean
    Dim strRow As String
    Dim intRow As Integer, intCount As Integer
    On Error GoTo errHand
    '将返回数据写入数据库
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    '更新标志
    gcnOracle.Execute "zl_消息主表_Insert('" & str日期 & "'," & lng序列号 & ",'" & strFuncName & "','" & strURL & "',NULL,9)", , adCmdStoredProc
    
    '写入接收数据
    intCount = Len(strData) \ 1000
    If Len(strData) Mod 1000 <> 0 Then intCount = intCount + 1
    For intRow = 0 To intCount
        strRow = Mid(strData, intRow * 1000 + 1, 1000)
        gcnOracle.Execute "zl_消息接收_Insert('" & str日期 & "'," & lng序列号 & "," & intRow + 1 & ",'" & Replace(strRow, "'", "''") & "')", , adCmdStoredProc
    Next
    
    gcnOracle.CommitTrans
    blnTrans = False
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    blnExit = False
End Sub

Private Sub Timer1_Timer()
    Dim rsTemp As New ADODB.Recordset
    '每处理完一个，就释放一次控制权
    
    Timer1.Enabled = False
    gstrSQL = " Select 日期,序列号,FUNCNAME,URL From 消息主表 Where nvl(标志,0)=0 Order by 日期,序列号"
    Call OpenRecordset(rsTemp, "提取待处理清单")
    With rsTemp
        Do While Not .EOF
            DoEvents
            
            If blnExit Then Exit Sub
            
            blnProcess = True
            Call 调用接口(!日期, !序列号, !FuncName, !url)
            blnProcess = False
            
            DoEvents
            .MoveNext
        Loop
    End With
    
    Timer1.Enabled = True
End Sub
