VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMgrUserGrant 
   Caption         =   "��������Ȩ"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   Icon            =   "frmMgrUserGrant.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   7575
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   1
      Left            =   3390
      Picture         =   "frmMgrUserGrant.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   0
      Left            =   3960
      Picture         =   "frmMgrUserGrant.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   375
   End
   Begin MSComctlLib.TreeView tvwGranted 
      Height          =   5895
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10398
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwNoGrant 
      Height          =   5895
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10398
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "�����û�(&F)"
      Height          =   350
      Left            =   6120
      TabIndex        =   2
      Top             =   65
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   4200
      TabIndex        =   1
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4980
      TabIndex        =   3
      Top             =   6975
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6255
      TabIndex        =   4
      Top             =   6975
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -360
      TabIndex        =   0
      Top             =   525
      Width           =   10110
   End
   Begin MSComctlLib.ImageList Img16 
      Left            =   3600
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1A5E
            Key             =   "�Զ�����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":82C0
            Key             =   "ϵͳװж����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":EB22
            Key             =   "����ת��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":15384
            Key             =   "�û�ע�����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1BBE6
            Key             =   "ϵͳ��Ǩ����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":22448
            Key             =   "ϵͳ��������"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":28CAA
            Key             =   "������־����"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":2F50C
            Key             =   "������־����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":35D6E
            Key             =   "ϵͳ����ѡ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":3C5D0
            Key             =   "�������޸�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":42E32
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":49694
            Key             =   "վ���ļ��ռ�"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":4FEF6
            Key             =   "������Ч����"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":56758
            Key             =   "��̨��ҵ����"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":5CFBA
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6381C
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6A07E
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":708E0
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":77142
            Key             =   "����״̬���"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":7D9A4
            Key             =   "�û���װ�ű�"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":84206
            Key             =   "վ�㲿������"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":8AA68
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":912CA
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":97B2C
            Key             =   "�û���Ȩ����"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":9E38E
            Key             =   "��ɫ��Ȩ����"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":A4BF0
            Key             =   "�˵�����滮"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":AB452
            Key             =   "վ�����п���"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B1CB4
            Key             =   "Ȩ�޹���"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B258E
            Key             =   "װж����"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B2E68
            Key             =   "���ݹ���"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B3742
            Key             =   "���й���"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B3CDC
            Key             =   "ר���"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B45B6
            Key             =   "DBA����"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":BAE18
            Key             =   "�ռ����"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C167A
            Key             =   "SQL����"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7EDC
            Key             =   "�Ự����"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":CE73E
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":D4FA0
            Key             =   "SQL����"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":DB802
            Key             =   "���ݿ�����"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "�ԡ����Ʊ򡱽�����Ȩ����"
      Height          =   180
      Left            =   960
      TabIndex        =   7
      Top             =   150
      UseMnemonic     =   0   'False
      Width           =   3090
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgOne 
      Height          =   480
      Left            =   300
      Picture         =   "frmMgrUserGrant.frx":E2064
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblModul 
      AutoSize        =   -1  'True
      Caption         =   "����Ȩ����(&A)"
      Height          =   180
      Left            =   210
      TabIndex        =   6
      Top             =   660
      Width           =   1170
   End
   Begin VB.Label lblGranted 
      AutoSize        =   -1  'True
      Caption         =   "����Ȩ����(&G)"
      Height          =   180
      Left            =   4335
      TabIndex        =   5
      Top             =   660
      Width           =   1170
   End
End
Attribute VB_Name = "frmMgrUserGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUser As String
Private mstrProg As String
Private mstrAccount As String 'Ϊ�ձ�ʾ���û���Ȩ
Private mblnOK As Boolean

Public Function GrantToProg(ByVal strAccount As String, ByVal strUser As String, ByVal strProg As String) As Boolean
    mstrUser = strUser
    mstrAccount = strAccount
    mstrProg = strProg
    mblnOK = False
    Me.Show 1
    GrantToProg = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Call FindPersonnel
End Sub

Private Sub MoveProg(objMoveIn As TreeView, objMoveOut As TreeView)
    Dim i As Long, y As Long
    Dim strDel As String, Node As Node
    
    For i = objMoveOut.Nodes.Count To 1 Step -1
        err = 0
        On Error Resume Next
        If objMoveOut.Nodes(i).Checked And Not objMoveOut.Nodes(i).Parent Is Nothing Then
            If err = 0 Then
                err = 0
                If objMoveIn.Nodes(objMoveOut.Nodes(i).Parent.Key).Key <> "" Then
                    If err <> 0 Then
                        '��������
                        Set Node = objMoveIn.Nodes.Add(, , objMoveOut.Nodes(i).Parent.Key, objMoveOut.Nodes(i).Parent.Text, objMoveOut.Nodes(i).Parent.Image, objMoveOut.Nodes(i).Parent.SelectedImage)
                        Node.Expanded = objMoveOut.Nodes(i).Parent.Expanded
                        Node.Checked = objMoveOut.Nodes(i).Parent.Checked
                        Node.ForeColor = objMoveOut.Nodes(i).Parent.ForeColor
                    End If
                     '��������
                    Set Node = objMoveIn.Nodes.Add(objMoveOut.Nodes(i).Parent.Key, tvwChild, objMoveOut.Nodes(i).Key, objMoveOut.Nodes(i).Text, objMoveOut.Nodes(i).Image, objMoveOut.Nodes(i).SelectedImage)
                    Node.Expanded = objMoveOut.Nodes(i).Expanded
                    Node.Checked = objMoveOut.Nodes(i).Checked
                    Node.ForeColor = objMoveOut.Nodes(i).ForeColor
                    'ɾ������
                    If objMoveOut.Nodes(i).Parent.Children = 1 Then
                        objMoveOut.Nodes.Remove objMoveOut.Nodes(i).Parent.Index
                    Else
                        objMoveOut.Nodes.Remove i
                    End If
                    
                End If
                On Error GoTo 0
            End If
        End If
    Next
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        Call MoveProg(tvwGranted, tvwNoGrant)
    ElseIf Index = 1 Then
        Call MoveProg(tvwNoGrant, tvwGranted)
    End If
End Sub

Private Sub cmdOK_Click()
'���ܣ���Ȩ
    Dim i As Integer, strProg As String
    Dim StrJiami() As Byte
    Dim strPwText As String
    Dim rsTemp As New ADODB.Recordset
    
    If mstrAccount = "" Then
        MsgBox "���Ȳ�����Ҫ��Ȩ���û���", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    For i = 1 To tvwGranted.Nodes.Count
        If Not tvwGranted.Nodes(i).Parent Is Nothing Then
            strProg = strProg & "," & Trim(Mid(tvwGranted.Nodes(i).Key, 2))
        End If
    Next
    strProg = Mid(strProg, 2)
    '���ܼ���
    If strProg <> "" Then
        Call DES_Encode(StrConv(strProg, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("��λ����", False, 0))
        strPwText = FuncByteTo16Code(StrJiami)
    End If
    On Error GoTo errHandle
    gstrSQL = "Select 1 From zlMgrGrant Where �û���='" & mstrAccount & "'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount > 0 Then
        If strPwText = "" Then
            gstrSQL = "Delete zlMgrGrant Where �û���='" & mstrAccount & "'"
        Else
            gstrSQL = "Update zlMgrGrant Set ����='" & strPwText & "' Where �û���='" & mstrAccount & "'"
        End If
    Else
        gstrSQL = "Insert into zlMgrGrant(�û���,����) values('" & mstrAccount & "','" & strPwText & "')"
    End If
    gcnOracle.Execute gstrSQL
    '���¹���Ա�˻���Ϣ
    rsTemp.Close
    gstrSQL = "Select 1 From zlRegInfo where ��Ŀ='����Ա'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
     If rsTemp.RecordCount > 0 Then
        gstrSQL = "Update zlRegInfo Set ����='" & gstrUserName & "' Where ��Ŀ='����Ա'"
    Else
        gstrSQL = "Insert into zlRegInfo(��Ŀ,����) values('����Ա','" & gstrUserName & "')"
    End If
    gcnOracle.Execute gstrSQL
    '��֤��
    strPwText = ""
    ReDim Preserve StrJiami(0)
    If gstrPassword <> "" Then
        Call DES_Encode(StrConv(gstrPassword, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("��λ����", False, 0))
        strPwText = FuncByteTo16Code(StrJiami)
    End If
    rsTemp.Close
    gstrSQL = "Select 1 From zlRegInfo where ��Ŀ='��֤��'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
     If rsTemp.RecordCount > 0 Then
        gstrSQL = "Update zlRegInfo Set ����='" & strPwText & "' Where ��Ŀ='��֤��'"
    Else
        gstrSQL = "Insert into zlRegInfo(��Ŀ,����) values('��֤��','" & strPwText & "')"




    End If
    gcnOracle.Execute gstrSQL
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub
Private Sub Form_Load()
    If mstrAccount = "" Then
        lblNote.Caption = "���������û�������Ա��������롣"
        txtFind.Visible = True
        cmdFind.Visible = True
    Else
        lblNote.Caption = "���ڶ�""" & mstrUser & """���й�������Ȩ��"
        txtFind.Visible = False
        cmdFind.Visible = False
    End If

    Call FillProg
End Sub

Private Sub FillProg()
'���ܣ���书��
    Dim rsTemp As New ADODB.Recordset
    Dim strProg As String, Node As Node
    Dim i As Long
    
    On Error GoTo errHandle
    '��ʾ���û����еĽ�ɫ
    gstrSQL = "Select /*+Rule */ a.���,a.����,A.�ϼ�,Column_Value as Ȩ��" & vbNewLine & _
            "From zlSvrTools A, (Select Column_Value From Table(Cast(f_Str2list('" & mstrProg & "') As Zltools.t_Strlist))) C" & vbNewLine & _
            "Where  a.��� = c.Column_Value(+) And A.����<>'��������Ȩ'" & vbNewLine & _
            "Order By a.���"

    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Do Until rsTemp.EOF
        With IIf(rsTemp!Ȩ�� & "" = "", tvwNoGrant, tvwGranted)
            '�ϼ����߶���
            If IsNull(rsTemp("�ϼ�")) Then
                Set Node = tvwNoGrant.Nodes.Add(, , "D" & rsTemp("���"), "��" & rsTemp("���") & "��" & rsTemp("����"))
                tvwNoGrant.Nodes("D" & rsTemp("���")).Sorted = True
                tvwNoGrant.Nodes("D" & rsTemp("���")).Expanded = True
                tvwNoGrant.Nodes("D" & rsTemp("���")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(rsTemp!���� & "").Index
                err.Clear: On Error GoTo errHandle
                Set Node = tvwGranted.Nodes.Add(, , "D" & rsTemp("���"), "��" & rsTemp("���") & "��" & rsTemp("����"))
                tvwGranted.Nodes("D" & rsTemp("���")).Sorted = True
                tvwGranted.Nodes("D" & rsTemp("���")).Expanded = True
                tvwGranted.Nodes("D" & rsTemp("���")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(rsTemp!���� & "").Index
                err.Clear: On Error GoTo errHandle
            Else
                Set Node = .Nodes.Add("D" & rsTemp("�ϼ�"), tvwChild, "C" & rsTemp("���"), rsTemp("����"))
                .Nodes("C" & rsTemp("���")).Sorted = True
                Node.Checked = False
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(rsTemp!���� & "").Index
                err.Clear: On Error GoTo errHandle
            End If
        End With
        rsTemp.MoveNext
    Loop
    'ɾ��û������ķ���
    For i = tvwNoGrant.Nodes.Count To 1 Step -1
        If tvwNoGrant.Nodes(i).Children = 0 And tvwNoGrant.Nodes(i).Parent Is Nothing Then
            tvwNoGrant.Nodes.Remove i
        End If
    Next
    For i = tvwGranted.Nodes.Count To 1 Step -1
        If tvwGranted.Nodes(i).Children = 0 And tvwGranted.Nodes(i).Parent Is Nothing Then
            tvwGranted.Nodes.Remove i
        End If
    Next
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraLine.Width = Me.Width
    cmdOK.Move Me.Width - cmdCancel.Width - cmdOK.Width - 400, Me.Height - cmdOK.Height - 650
    cmdCancel.Move cmdOK.Left + cmdOK.Width + 100, cmdOK.Top
    tvwNoGrant.Width = Me.Width \ 2 - 885
    tvwGranted.Width = Me.Width \ 2 - 885
    tvwNoGrant.Height = cmdOK.Top - tvwNoGrant.Top - 100
    tvwGranted.Height = tvwNoGrant.Height
    tvwGranted.Left = tvwNoGrant.Left + tvwNoGrant.Width + 1185
    cmdMove(1).Left = tvwNoGrant.Left + tvwNoGrant.Width + 150
    cmdMove(0).Left = cmdMove(1).Left + cmdMove(1).Width + 150
    lblGranted.Left = tvwGranted.Left
End Sub

Private Sub tvwGranted_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call NodeCheckMode(Node, tvwGranted)
End Sub

Private Sub tvwNoGrant_NodeCheck(ByVal Node As MSComctlLib.Node)
     Call NodeCheckMode(Node, tvwNoGrant)
End Sub

Private Sub NodeCheckMode(ByRef Node As MSComctlLib.Node, ByRef objtvwThis As TreeView)
'���ܣ�������ѡ�и��ڵ㣬�Զ�ѡ�������ӽڵ㣬ѡ�������ӽڵ㣬���ڵ�Ҳѡ��
    Dim i As Long
    Dim blnIsNothing As Boolean
    
    LockWindowUpdate objtvwThis.hwnd
    If Node.Parent Is Nothing Then
        For i = Node.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Key Then
                    objtvwThis.Nodes(i).Checked = Node.Checked
                End If
            End If
        Next
    Else
        For i = Node.Parent.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Parent.Key Then
                    If Not objtvwThis.Nodes(i).Checked = Node.Checked Then blnIsNothing = True
                End If
            End If
        Next
        If blnIsNothing Then
            Node.Parent.Checked = False
        Else
            Node.Parent.Checked = Node.Checked
        End If
    End If
    LockWindowUpdate 0
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindPersonnel
    End If
End Sub

Private Sub FindPersonnel()
'���ܣ�������Ա
    Dim rsTemp As New Recordset
    Dim objPoint As POINTAPI
    
    If txtFind.Text = "" Then Exit Sub
    gstrSQL = "Select b.�û���, c.����, c.����, d.���� As ��������" & vbNewLine & _
            "From  Zlmgrgrant A,�ϻ���Ա�� B, ��Ա�� C, ���ű� D, ������Ա E" & vbNewLine & _
            "Where a.�û���(+) = b.�û��� And b.��Աid = c.Id And c.Id = e.��Աid And d.Id = e.����id And A.�û��� is null And e.ȱʡ = 1 And B.�û��� <> '" & gstrUserName & "'" & _
            " And(b.�û��� like '" & UCase(Trim(txtFind.Text)) & "%' Or c.���� Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.���� Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.���=' & UCase(Trim(txtFind.Text)) & ')" & _
            " Order By c.����"
    Set rsTemp = New ADODB.Recordset
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ҵ��û������ڣ������Ѿ�ӵ����Ȩ�ޣ����顣", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        Exit Sub
    End If
    Call ClientToScreen(txtFind.hwnd, objPoint)
    
    If frmSelectList.ShowSelect(Nothing, rsTemp, "�û���,900,0,1;����,900,0,1;����,650,0,0;��������,1500,0,1", objPoint.X * 15 - 30, objPoint.y * 15 + cmdFind.Height - 30, txtFind.Width + cmdFind.Width + 1300, 3000, "", "������Ա", , , True) = False Then
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        rsTemp.Filter = 0
        Exit Sub
    Else
        txtFind.Text = rsTemp!���� & ""
        mstrAccount = rsTemp!�û��� & ""
        mstrUser = rsTemp!���� & ""
    End If
End Sub
