VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
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
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk���� 
      Caption         =   "�������Ѻ��ձ걾(��ȡ������Ϣʱ��������ȡ�Ѻ��յı걾)"
      Height          =   195
      Left            =   2955
      TabIndex        =   41
      Top             =   2940
      Width           =   5595
   End
   Begin VB.TextBox txt��� 
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
         Caption         =   "���ݱ��浽ָ������"
         Height          =   210
         Left            =   105
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "    �벻Ҫ������ģ�������ý����ڽ����������������յ������ݱ��浽��ָ����������"
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   540
         TabIndex        =   37
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "�Ƴ�(&M)"
      Height          =   350
      Left            =   1260
      TabIndex        =   32
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   135
      TabIndex        =   31
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "��ս�����־"
      Height          =   225
      Left            =   2910
      TabIndex        =   29
      Top             =   4575
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7515
      TabIndex        =   28
      Top             =   4545
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6225
      TabIndex        =   27
      Top             =   4545
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
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
      TabCaption(0)   =   "COMͨ������(&M)"
      TabPicture(0)   =   "frmParaSet.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "chkCom"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "TCP/IPͨ������(&T)"
      TabPicture(1)   =   "frmParaSet.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ChkIP"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraIP"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraIP 
         Caption         =   "����"
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
            Caption         =   "��Ϊ����"
            Height          =   255
            Index           =   0
            Left            =   2805
            TabIndex        =   20
            Top             =   225
            Width           =   1095
         End
         Begin VB.OptionButton OptHost 
            Caption         =   "��Ϊ�ն�"
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
            Caption         =   "����ģʽ"
            Height          =   255
            Left            =   3495
            TabIndex        =   24
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblPort 
            Alignment       =   1  'Right Justify
            Caption         =   "�˿�"
            Height          =   180
            Left            =   2025
            TabIndex        =   18
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblIP 
            Alignment       =   1  'Right Justify
            Caption         =   "����IP"
            Height          =   180
            Left            =   30
            TabIndex        =   17
            Top             =   660
            Width           =   690
         End
      End
      Begin VB.CheckBox ChkIP 
         Caption         =   "����TCP/IPͨ��"
         Height          =   240
         Left            =   3900
         TabIndex        =   13
         Top             =   585
         Width           =   1680
      End
      Begin VB.CheckBox chkCom 
         Caption         =   "����COMͨ��"
         Height          =   240
         Left            =   -70740
         TabIndex        =   12
         Top             =   450
         Width           =   1440
      End
      Begin VB.Frame Frame1 
         Caption         =   "�˿�����"
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
            Caption         =   "����ģʽ"
            Height          =   255
            Left            =   3390
            TabIndex        =   22
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "�����ٶ�"
            Height          =   255
            Left            =   1260
            TabIndex        =   11
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ֹͣλ"
            Height          =   285
            Left            =   3390
            TabIndex        =   10
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "��żλ"
            Height          =   285
            Left            =   1425
            TabIndex        =   9
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "����λ"
            Height          =   285
            Left            =   3390
            TabIndex        =   8
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "����Э��"
            Height          =   255
            Left            =   1260
            TabIndex        =   7
            Top             =   1035
            Width           =   735
         End
      End
   End
   Begin VB.ListBox Lst���� 
      Height          =   4020
      Left            =   75
      TabIndex        =   25
      Top             =   360
      Width           =   2565
   End
   Begin VB.Label Label3 
      Caption         =   "ÿ        ���Զ�Ӧ��ȡֵΪ0-3600,��Ϊ0����ʾ��ʹ�ô˹���)"
      Height          =   195
      Left            =   2925
      TabIndex        =   40
      ToolTipText     =   "��Ҫ�ӿڳ���֧�ֲŻᷢ�����"
      Top             =   2625
      Width           =   5715
   End
   Begin VB.Label lbl 
      Caption         =   "������������"
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
Private mblnEdit As Boolean '�Ƿ���Ȩ�޽����޸�

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
    If frmSelect.Select���� Then
        iLastDev = -1
        LoadPropertySettings
        If Lst����.ListCount > 0 Then Lst����.ListIndex = 0
    End If
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim lngID As Long, i As Integer
    Dim lastIndex As Long
    If Lst����.ListCount <= 0 Then Exit Sub
    lngID = Lst����.ItemData(Lst����.ListIndex)
    If lngID > 0 Then

        For i = LBound(g����) To UBound(g����)
            If lngID = g����(i).ID Then
                g����(i).ID = 0
                Exit For
            End If
        Next
        lastIndex = Lst����.ListIndex
        Lst����.RemoveItem lastIndex
                
        iLastDev = -1
        If lastIndex - 1 >= 0 Then
            Lst����.ListIndex = lastIndex - 1
        Else
            If Lst����.ListCount > 0 Then Lst����.ListIndex = 0
        End If
    End If
    
End Sub

Private Sub cmdHelp_Click()
    gobjComLib.ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, blnNoDev As Boolean, strMsg As String, lng����ID As Long, str����ֵ As String
    Dim strIDs As String
    '����ǰ���ñ��浽�ڴ���


    If Not mblnEdit Then
        ifOK = True
        Unload Me
        Exit Sub
    End If
    
    If Lst����.ListCount > 0 Then
        iLastDev = Lst����.ListIndex: Lst����_Click
    End If
    blnNoDev = True
    str����ֵ = ""
    strMsg = ""
    
    For i = LBound(g����) To UBound(g����)
        If g����(i).ID > 0 Then

            blnNoDev = False
            '������
            
            If g����(i).���� = 1 Then
                'TCP/IP
                
                If ValidateIP(g����(i).IP) Then strMsg = strMsg & vbNewLine & g����(i).�������� & " IP����"
                
                If ValidatePort(g����(i).IP�˿�) Then strMsg = strMsg & vbNewLine & g����(i).�������� & " IP�˿ڴ���"
                
                If Not ValidateIP(g����(i).IP) And Not ValidatePort(g����(i).IP�˿�) Then
                    If InStr(str����ֵ, "," & g����(i).IP & ":" & g����(i).IP�˿�) > 0 Then
                        strMsg = strMsg & vbNewLine & g����(i).�������� & " IP��ַ�Ͷ˿��ظ�����"
                    Else
                        str����ֵ = str����ֵ & "," & g����(i).IP & ":" & g����(i).IP�˿�
                    End If
                End If
            Else
                'COM
                If g����(i).COM�� = 0 Then
                    strMsg = strMsg & vbNewLine & g����(i).�������� & " COM�����ô���"
                Else
                    If InStr(str����ֵ, ",COM" & g����(i).COM��) > 0 Then
                        strMsg = strMsg & vbNewLine & g����(i).�������� & " COM���ظ�����"
                    Else
                        str����ֵ = str����ֵ & ",COM" & g����(i).COM��
                    End If
                End If
            End If
            
            If Val(g����(i).�Զ�Ӧ��) < 0 Or Val(g����(i).�Զ�Ӧ��) > 3600 Then
                strMsg = strMsg & vbNewLine & g����(i).�������� & " �Զ�Ӧ��ʱ����0 - 3600��֮��"
            End If

        End If
    Next
    
    
    If blnNoDev Then
        If MsgBox("û�������κ�������ϵͳ�����ܽ��ռ������ݣ��Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Lst����.SetFocus: Exit Sub
        End If
    Else
        If MsgBox("ϵͳ���������Ӽ������������ݽ��չ��̽���ͣ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Lst����.SetFocus: Exit Sub
        End If
    End If
    
    If strMsg <> "" Then
        MsgBox "���������������飺" & strMsg, vbQuestion
        Exit Sub
    End If

    SavePortsSetting
    If gblnFromDB Then
        Call gobjDatabase.SetPara("��ս�����־", Me.chkClear.Value, glngSys, 1208)
    Else
        
        Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv", "��ս�����־", CStr(Me.chkClear.Value))
    End If

    ifOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim objControl As Object
    mblnEdit = InStr(";" & gstrPrivs & ";", ";ͨѶ��������;") > 0

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
    If Lst����.ListCount > 0 Then Lst����.ListIndex = 0
    
End Sub

Private Sub LoadPropertySettings()
    Dim rsDev As ADODB.Recordset
    '���봮�������趨---������
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
    
    ' ��������λ����
    With cboAttr(2)
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
    End With
    
    ' ������ż��������
    With cboAttr(3)
        .AddItem "None"
        .AddItem "Odd"
        .AddItem "Even"
        .AddItem "Mark"
        .AddItem "Space"
    End With
    
    ' ����ֹͣλ����
    With cboAttr(4)
        .AddItem "1"
        .AddItem "1.5"
        .AddItem "2"
    End With
    '
    
    With cboAttr(9) '����ģʽ
        .Clear
        .AddItem "�ַ�"
        .AddItem "��ģʽ"
    End With
    
    With cboInMode
        .Clear
        .AddItem "�ַ�"
        .AddItem "��ģʽ"
    End With
    
    '��������
    Set rsDev = GetDevices
'    With cboAttr(0)
'        .Clear
'        .AddItem "δָ���豸"
'        .ItemData(0) = 0

        cboSaveAs.Clear
        cboSaveAs.AddItem "ȱʡ"
        cboSaveAs.ItemData(0) = 0

    If Not rsDev Is Nothing Then
        Do While Not rsDev.EOF
'                .AddItem "(" & rsDev("����") & ")" & rsDev("����")
'                .ItemData(.ListCount - 1) = rsDev("ID")

            cboSaveAs.AddItem "(" & rsDev("����") & ")" & rsDev("����")
            cboSaveAs.ItemData(cboSaveAs.ListCount - 1) = rsDev("ID")
    
            rsDev.MoveNext
        Loop
    End If
    Lst����.Clear
    For i = LBound(g����) To UBound(g����)
       If g����(i).ID > 0 Then
           rsDev.Filter = "ID=" & g����(i).ID
           If Not rsDev.EOF Then
               Lst����.AddItem "(" & rsDev("����") & ")" & rsDev("����")
               Lst����.ItemData(Lst����.ListCount - 1) = rsDev("ID")
           End If
       End If
    Next
     
'    End With
End Sub


Private Sub Lst����_Click()
    Dim lng����ID As Long
    Dim i As Integer
    On Error GoTo errH
    
    If iLastDev > -1 Then
        lng����ID = Val(Lst����.ItemData(iLastDev))
        
         For i = LBound(g����) To UBound(g����)
            If Val(g����(i).ID) = lng����ID Then
                '�����޸�
                g����(i).IP = txtIP
                g����(i).IP�˿� = CLng(Val(txtPort))
                g����(i).SaveAsID = Val(cboSaveAs.ItemData(cboSaveAs.ListIndex))
                g����(i).������ = CLng(Val(cboAttr(1).Text))
                g����(i).����λ = cboAttr(2).Text
                g����(i).���� = ChkIP.Value
                g����(i).COM�� = CInt(Val(txtCom))
                g����(i).У��λ = Left(cboAttr(3).Text, 1)
                g����(i).ֹͣλ = cboAttr(4).Text
                g����(i).���� = cboAttr(5).ListIndex
                g����(i).���� = IIf(OptHost(0).Value, 1, 0)
                g����(i).�ַ�ģʽ = IIf(chkCom.Value = 1, cboAttr(9).ListIndex, cboInMode.ListIndex)
                If IsNumeric(Trim(Me.txt���.Text)) Then
                    g����(i).�Զ�Ӧ�� = Trim(txt���.Text)
                End If
                g����(i).�ɷ��Ѻ˱걾 = Val(chk����.Value)
                Exit For
            End If
        Next
    End If
    lng����ID = Val(Lst����.ItemData(Lst����.ListIndex))
    
    If lng����ID > 0 Then
        For i = LBound(g����) To UBound(g����)
            
            If Val(g����(i).ID) = lng����ID Then
                
                If g����(i).���� = 0 Then
                    txtCom = g����(i).COM��
                    ChkIP.Value = 0
                    chkCom.Value = 1
                    sstbSet.Tab = 0
                    Me.cboAttr(1).Text = g����(i).������
                    Me.cboAttr(2).Text = g����(i).����λ
                    Me.cboAttr(3).Text = Switch(UCase(g����(i).У��λ) = "N", "None", _
                        UCase(g����(i).У��λ) = "E", "Even", _
                        UCase(g����(i).У��λ) = "O", "Odd", _
                        UCase(g����(i).У��λ) = "M", "Mark", _
                        UCase(g����(i).У��λ) = "S", "Space")
                    Me.cboAttr(4).Text = g����(i).ֹͣλ
                    Me.cboAttr(5).ListIndex = Val(g����(i).����)

                Else
                    txtCom = g����(i).COM��
                    ChkIP.Value = 1
                    chkCom.Value = 0
                    sstbSet.Tab = 1
                                    
                    txtPort = g����(i).IP�˿�
                    txtIP = g����(i).IP
                    OptHost(0).Value = g����(i).���� = 1
                    
                    If OptHost(0).Value Then
                        Call OptHost_Click(1)
                    Else
                        Call OptHost_Click(0)
                    End If
                End If
                Me.cboAttr(9).ListIndex = Val(g����(i).�ַ�ģʽ)
                cboInMode.ListIndex = Val(g����(i).�ַ�ģʽ)
                Me.txt���.Text = CStr(g����(i).�Զ�Ӧ��)
                If Left(Me.txt���, 1) = "." Then Me.txt���.Text = "0" & Me.txt���.Text
                
                Me.cboSaveAs.ListIndex = GetComboxIndex(cboSaveAs, g����(i).SaveAsID)
                Me.chk����.Value = g����(i).�ɷ��Ѻ˱걾
            End If
        Next
        
    End If
    iLastDev = Lst����.ListIndex
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub OptHost_Click(Index As Integer)
    If Index = 0 Then
        lblIP.Caption = "����IP"
        lblPort.Caption = "�˿�"
    Else
        lblIP.Caption = "����IP"
        lblPort.Caption = "�˿�"
    End If
End Sub

