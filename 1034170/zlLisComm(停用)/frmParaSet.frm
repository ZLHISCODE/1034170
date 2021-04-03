VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "frmParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk核收 
      Caption         =   "允许发送已核收标本(提取病人信息时，可以提取已核收的标本)"
      Height          =   195
      Left            =   2955
      TabIndex        =   41
      Top             =   2940
      Width           =   5595
   End
   Begin VB.TextBox txt间隔 
      Height          =   270
      Left            =   3195
      TabIndex        =   39
      Top             =   2595
      Width           =   510
   End
   Begin VB.Frame fraSaveAs 
      Height          =   1110
      Left            =   2835
      TabIndex        =   35
      Top             =   3270
      Width           =   5880
      Begin VB.ComboBox cboSaveAs 
         Height          =   300
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   180
         Width           =   3780
      End
      Begin VB.Label Label9 
         Caption         =   "数据保存到指定仪器"
         Height          =   210
         Left            =   105
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "    请不要随意更改，这个设置仅用于将“检验仪器”接收到的数据保存到“指定仪器”。"
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   540
         TabIndex        =   37
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "移除(&M)"
      Height          =   350
      Left            =   1260
      TabIndex        =   32
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
      Height          =   350
      Left            =   135
      TabIndex        =   31
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "清空接收日志"
      Height          =   225
      Left            =   2910
      TabIndex        =   29
      Top             =   4575
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7515
      TabIndex        =   28
      Top             =   4545
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6225
      TabIndex        =   27
      Top             =   4545
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   4545
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1100
   End
   Begin TabDlg.SSTab sstbSet 
      Height          =   2040
      Left            =   2790
      TabIndex        =   0
      Top             =   450
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   3598
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "COM通信设置(&M)"
      TabPicture(0)   =   "frmParaSet.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "chkCom"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "TCP/IP通信设置(&T)"
      TabPicture(1)   =   "frmParaSet.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ChkIP"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraIP"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraIP 
         Caption         =   "特性"
         Height          =   1035
         Left            =   210
         TabIndex        =   14
         Top             =   855
         Width           =   5505
         Begin VB.ComboBox cboInMode 
            Height          =   300
            Left            =   4305
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   615
            Width           =   1080
         End
         Begin VB.OptionButton OptHost 
            Caption         =   "作为主机"
            Height          =   255
            Index           =   0
            Left            =   2805
            TabIndex        =   20
            Top             =   225
            Width           =   1095
         End
         Begin VB.OptionButton OptHost 
            Caption         =   "作为终端"
            Height          =   225
            Index           =   1
            Left            =   1230
            TabIndex        =   19
            Top             =   225
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtPort 
            Height          =   300
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   16
            Text            =   "66666"
            Top             =   615
            Width           =   630
         End
         Begin VB.TextBox txtIP 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   780
            MaxLength       =   15
            TabIndex        =   15
            Text            =   "0.0.0.0"
            Top             =   615
            Width           =   1500
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "接收模式"
            Height          =   255
            Left            =   3495
            TabIndex        =   24
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblPort 
            Alignment       =   1  'Right Justify
            Caption         =   "端口"
            Height          =   180
            Left            =   2025
            TabIndex        =   18
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblIP 
            Alignment       =   1  'Right Justify
            Caption         =   "仪器IP"
            Height          =   180
            Left            =   30
            TabIndex        =   17
            Top             =   660
            Width           =   690
         End
      End
      Begin VB.CheckBox ChkIP 
         Caption         =   "启用TCP/IP通信"
         Height          =   240
         Left            =   3900
         TabIndex        =   13
         Top             =   585
         Width           =   1680
      End
      Begin VB.CheckBox chkCom 
         Caption         =   "启用COM通信"
         Height          =   240
         Left            =   -70740
         TabIndex        =   12
         Top             =   450
         Width           =   1440
      End
      Begin VB.Frame Frame1 
         Caption         =   "端口特性"
         Height          =   1335
         Left            =   -74895
         TabIndex        =   1
         Top             =   615
         Width           =   5640
         Begin VB.TextBox txtCom 
            Height          =   270
            Left            =   480
            TabIndex        =   33
            Top             =   630
            Width           =   510
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   9
            ItemData        =   "frmParaSet.frx":0044
            Left            =   4155
            List            =   "frmParaSet.frx":0046
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   990
            Width           =   1200
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   1
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   255
            Width           =   1230
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   4
            ItemData        =   "frmParaSet.frx":0048
            Left            =   4155
            List            =   "frmParaSet.frx":004A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   630
            Width           =   1215
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   3
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   630
            Width           =   1230
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   2
            Left            =   4155
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   255
            Width           =   1215
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   5
            ItemData        =   "frmParaSet.frx":004C
            Left            =   2100
            List            =   "frmParaSet.frx":005C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   990
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "COM"
            Height          =   180
            Left            =   135
            TabIndex        =   34
            Top             =   675
            Width           =   315
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "接收模式"
            Height          =   255
            Left            =   3390
            TabIndex        =   22
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "传输速度"
            Height          =   255
            Left            =   1260
            TabIndex        =   11
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "停止位"
            Height          =   285
            Left            =   3390
            TabIndex        =   10
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "奇偶位"
            Height          =   285
            Left            =   1425
            TabIndex        =   9
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "数据位"
            Height          =   285
            Left            =   3390
            TabIndex        =   8
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "握手协议"
            Height          =   255
            Left            =   1260
            TabIndex        =   7
            Top             =   1035
            Width           =   735
         End
      End
   End
   Begin VB.ListBox Lst仪器 
      Height          =   4020
      Left            =   75
      TabIndex        =   25
      Top             =   360
      Width           =   2565
   End
   Begin VB.Label Label3 
      Caption         =   "每        秒自动应答（取值为0-3600,设为0，表示不使用此功能)"
      Height          =   195
      Left            =   2925
      TabIndex        =   40
      ToolTipText     =   "需要接口程序支持才会发送命令。"
      Top             =   2625
      Width           =   5715
   End
   Begin VB.Label lbl 
      Caption         =   "本机连接仪器"
      Height          =   195
      Left            =   135
      TabIndex        =   30
      Top             =   75
      Width           =   1260
   End
End
Attribute VB_Name = "frmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ifOK As Boolean
Private mblnEdit As Boolean '是否有权限进行修改

Private iLastDev As Long

Public Function ShowMe(objParent As Object) As Boolean
    Me.chkClear.Value = IIf(gblnClearData, 1, 0)
    Me.Show vbModal, objParent
    ShowMe = ifOK
End Function

Private Sub cboAttr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkCom_Click()
    If chkCom.Value = 0 Then
        ChkIP.Value = 1
        sstbSet.Tab = 1
    Else
        ChkIP.Value = 0
    End If
End Sub

Private Sub ChkIP_Click()
    If ChkIP.Value = 0 Then
        chkCom.Value = 1
        sstbSet.Tab = 0
    Else
        chkCom.Value = 0
    End If
End Sub

Private Sub cmdAdd_Click()
    If frmSelect.Select仪器 Then
        iLastDev = -1
        LoadPropertySettings
        If Lst仪器.ListCount > 0 Then Lst仪器.ListIndex = 0
    End If
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim lngID As Long, i As Integer
    Dim lastIndex As Long
    If Lst仪器.ListCount <= 0 Then Exit Sub
    lngID = Lst仪器.ItemData(Lst仪器.ListIndex)
    If lngID > 0 Then

        For i = LBound(g仪器) To UBound(g仪器)
            If lngID = g仪器(i).ID Then
                g仪器(i).ID = 0
                Exit For
            End If
        Next
        lastIndex = Lst仪器.ListIndex
        Lst仪器.RemoveItem lastIndex
                
        iLastDev = -1
        If lastIndex - 1 >= 0 Then
            Lst仪器.ListIndex = lastIndex - 1
        Else
            If Lst仪器.ListCount > 0 Then Lst仪器.ListIndex = 0
        End If
    End If
    
End Sub

Private Sub cmdHelp_Click()
    gobjComLib.ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, blnNoDev As Boolean, strMsg As String, lng仪器ID As Long, str已设值 As String
    Dim strIDs As String
    '将当前设置保存到内存中


    If Not mblnEdit Then
        ifOK = True
        Unload Me
        Exit Sub
    End If
    
    If Lst仪器.ListCount > 0 Then
        iLastDev = Lst仪器.ListIndex: Lst仪器_Click
    End If
    blnNoDev = True
    str已设值 = ""
    strMsg = ""
    
    For i = LBound(g仪器) To UBound(g仪器)
        If g仪器(i).ID > 0 Then

            blnNoDev = False
            '检查参数
            
            If g仪器(i).类型 = 1 Then
                'TCP/IP
                
                If ValidateIP(g仪器(i).IP) Then strMsg = strMsg & vbNewLine & g仪器(i).编码名称 & " IP错误"
                
                If ValidatePort(g仪器(i).IP端口) Then strMsg = strMsg & vbNewLine & g仪器(i).编码名称 & " IP端口错误"
                
                If Not ValidateIP(g仪器(i).IP) And Not ValidatePort(g仪器(i).IP端口) Then
                    If InStr(str已设值, "," & g仪器(i).IP & ":" & g仪器(i).IP端口) > 0 Then
                        strMsg = strMsg & vbNewLine & g仪器(i).编码名称 & " IP地址和端口重复设置"
                    Else
                        str已设值 = str已设值 & "," & g仪器(i).IP & ":" & g仪器(i).IP端口
                    End If
                End If
            Else
                'COM
                If g仪器(i).COM口 = 0 Then
                    strMsg = strMsg & vbNewLine & g仪器(i).编码名称 & " COM口设置错误"
                Else
                    If InStr(str已设值, ",COM" & g仪器(i).COM口) > 0 Then
                        strMsg = strMsg & vbNewLine & g仪器(i).编码名称 & " COM口重复设置"
                    Else
                        str已设值 = str已设值 & ",COM" & g仪器(i).COM口
                    End If
                End If
            End If
            
            If Val(g仪器(i).自动应答) < 0 Or Val(g仪器(i).自动应答) > 3600 Then
                strMsg = strMsg & vbNewLine & g仪器(i).编码名称 & " 自动应答时间在0 - 3600秒之间"
            End If

        End If
    Next
    
    
    If blnNoDev Then
        If MsgBox("没有连接任何仪器，系统将不能接收检验数据！是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Lst仪器.SetFocus: Exit Sub
        End If
    Else
        If MsgBox("系统将重新连接检验仪器，数据接收过程将暂停！是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Lst仪器.SetFocus: Exit Sub
        End If
    End If
    
    If strMsg <> "" Then
        MsgBox "参数设置有误，请检查：" & strMsg, vbQuestion
        Exit Sub
    End If

    SavePortsSetting
    If gblnFromDB Then
        Call gobjDatabase.SetPara("清空接收日志", Me.chkClear.Value, glngSys, 1208)
    Else
        
        Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv", "清空接收日志", CStr(Me.chkClear.Value))
    End If

    ifOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim objControl As Object
    mblnEdit = InStr(";" & gstrPrivs & ";", ";通讯参数设置;") > 0

    If Not mblnEdit Then
        For Each objControl In Me.Controls
            If InStr("chkClear,cmdHelp,cmdOK,cmdCancel,lvwComm,sstbSet", objControl.Name) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        Next
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    ifOK = False
    mblnEdit = False

    iLastDev = -1
    LoadPropertySettings
    If Lst仪器.ListCount > 0 Then Lst仪器.ListIndex = 0
    
End Sub

Private Sub LoadPropertySettings()
    Dim rsDev As ADODB.Recordset
    '载入串口速率设定---波特率
    Dim i As Integer
    With cboAttr(1)
        .AddItem "110"
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "28800"
        .AddItem "38400"
        .AddItem "56000"
        .AddItem "57600"
        .AddItem "115200"
        .AddItem "128000"
        .AddItem "256000"
    End With
    
    ' 载入数据位设置
    With cboAttr(2)
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
    End With
    
    ' 载入奇偶检验设置
    With cboAttr(3)
        .AddItem "None"
        .AddItem "Odd"
        .AddItem "Even"
        .AddItem "Mark"
        .AddItem "Space"
    End With
    
    ' 载入停止位设置
    With cboAttr(4)
        .AddItem "1"
        .AddItem "1.5"
        .AddItem "2"
    End With
    '
    
    With cboAttr(9) '接收模式
        .Clear
        .AddItem "字符"
        .AddItem "流模式"
    End With
    
    With cboInMode
        .Clear
        .AddItem "字符"
        .AddItem "流模式"
    End With
    
    '检验仪器
    Set rsDev = GetDevices
'    With cboAttr(0)
'        .Clear
'        .AddItem "未指定设备"
'        .ItemData(0) = 0

        cboSaveAs.Clear
        cboSaveAs.AddItem "缺省"
        cboSaveAs.ItemData(0) = 0

    If Not rsDev Is Nothing Then
        Do While Not rsDev.EOF
'                .AddItem "(" & rsDev("编码") & ")" & rsDev("名称")
'                .ItemData(.ListCount - 1) = rsDev("ID")

            cboSaveAs.AddItem "(" & rsDev("编码") & ")" & rsDev("名称")
            cboSaveAs.ItemData(cboSaveAs.ListCount - 1) = rsDev("ID")
    
            rsDev.MoveNext
        Loop
    End If
    Lst仪器.Clear
    For i = LBound(g仪器) To UBound(g仪器)
       If g仪器(i).ID > 0 Then
           rsDev.Filter = "ID=" & g仪器(i).ID
           If Not rsDev.EOF Then
               Lst仪器.AddItem "(" & rsDev("编码") & ")" & rsDev("名称")
               Lst仪器.ItemData(Lst仪器.ListCount - 1) = rsDev("ID")
           End If
       End If
    Next
     
'    End With
End Sub


Private Sub Lst仪器_Click()
    Dim lng仪器ID As Long
    Dim i As Integer
    On Error GoTo errH
    
    If iLastDev > -1 Then
        lng仪器ID = Val(Lst仪器.ItemData(iLastDev))
        
         For i = LBound(g仪器) To UBound(g仪器)
            If Val(g仪器(i).ID) = lng仪器ID Then
                '保存修改
                g仪器(i).IP = txtIP
                g仪器(i).IP端口 = CLng(Val(txtPort))
                g仪器(i).SaveAsID = Val(cboSaveAs.ItemData(cboSaveAs.ListIndex))
                g仪器(i).波特率 = CLng(Val(cboAttr(1).Text))
                g仪器(i).数据位 = cboAttr(2).Text
                g仪器(i).类型 = ChkIP.Value
                g仪器(i).COM口 = CInt(Val(txtCom))
                g仪器(i).校验位 = Left(cboAttr(3).Text, 1)
                g仪器(i).停止位 = cboAttr(4).Text
                g仪器(i).握手 = cboAttr(5).ListIndex
                g仪器(i).主机 = IIf(OptHost(0).Value, 1, 0)
                g仪器(i).字符模式 = IIf(chkCom.Value = 1, cboAttr(9).ListIndex, cboInMode.ListIndex)
                If IsNumeric(Trim(Me.txt间隔.Text)) Then
                    g仪器(i).自动应答 = Trim(txt间隔.Text)
                End If
                g仪器(i).可发已核标本 = Val(chk核收.Value)
                Exit For
            End If
        Next
    End If
    lng仪器ID = Val(Lst仪器.ItemData(Lst仪器.ListIndex))
    
    If lng仪器ID > 0 Then
        For i = LBound(g仪器) To UBound(g仪器)
            
            If Val(g仪器(i).ID) = lng仪器ID Then
                
                If g仪器(i).类型 = 0 Then
                    txtCom = g仪器(i).COM口
                    ChkIP.Value = 0
                    chkCom.Value = 1
                    sstbSet.Tab = 0
                    Me.cboAttr(1).Text = g仪器(i).波特率
                    Me.cboAttr(2).Text = g仪器(i).数据位
                    Me.cboAttr(3).Text = Switch(UCase(g仪器(i).校验位) = "N", "None", _
                        UCase(g仪器(i).校验位) = "E", "Even", _
                        UCase(g仪器(i).校验位) = "O", "Odd", _
                        UCase(g仪器(i).校验位) = "M", "Mark", _
                        UCase(g仪器(i).校验位) = "S", "Space")
                    Me.cboAttr(4).Text = g仪器(i).停止位
                    Me.cboAttr(5).ListIndex = Val(g仪器(i).握手)

                Else
                    txtCom = g仪器(i).COM口
                    ChkIP.Value = 1
                    chkCom.Value = 0
                    sstbSet.Tab = 1
                                    
                    txtPort = g仪器(i).IP端口
                    txtIP = g仪器(i).IP
                    OptHost(0).Value = g仪器(i).主机 = 1
                    
                    If OptHost(0).Value Then
                        Call OptHost_Click(1)
                    Else
                        Call OptHost_Click(0)
                    End If
                End If
                Me.cboAttr(9).ListIndex = Val(g仪器(i).字符模式)
                cboInMode.ListIndex = Val(g仪器(i).字符模式)
                Me.txt间隔.Text = CStr(g仪器(i).自动应答)
                If Left(Me.txt间隔, 1) = "." Then Me.txt间隔.Text = "0" & Me.txt间隔.Text
                
                Me.cboSaveAs.ListIndex = GetComboxIndex(cboSaveAs, g仪器(i).SaveAsID)
                Me.chk核收.Value = g仪器(i).可发已核标本
            End If
        Next
        
    End If
    iLastDev = Lst仪器.ListIndex
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub OptHost_Click(Index As Integer)
    If Index = 0 Then
        lblIP.Caption = "本地IP"
        lblPort.Caption = "端口"
    Else
        lblIP.Caption = "仪器IP"
        lblPort.Caption = "端口"
    End If
End Sub

