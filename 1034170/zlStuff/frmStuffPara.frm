VERSION 5.00
Begin VB.Form frmStuffPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmStuffPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "3�����ķ��������Զ�����"
      Height          =   1335
      Left            =   180
      TabIndex        =   19
      Top             =   3120
      Width           =   4620
      Begin VB.OptionButton optsetall 
         Caption         =   "�ⷿ�ͷ��ϲ��ŷ���"
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1980
      End
      Begin VB.OptionButton optSet�ⷿ 
         Caption         =   "���ⷿ����"
         Height          =   210
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   1500
      End
      Begin VB.OptionButton optSetNotall 
         Caption         =   "�ⷿ�ͷ��ϲ��Ŷ�������"
         Height          =   210
         Left            =   2160
         TabIndex        =   21
         Top             =   840
         Width           =   2340
      End
      Begin VB.OptionButton optSet�ֶ� 
         Caption         =   "�ֹ����÷�������"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1740
      End
   End
   Begin VB.Frame fra 
      Caption         =   "3������Ӧ���ڡ����ķ�Χ"
      Height          =   4185
      Left            =   4890
      TabIndex        =   14
      Top             =   270
      Width           =   3120
      Begin VB.CheckBox chk�洢�ⷿ 
         Caption         =   "Ӧ���ڷ�����������������(&N)"
         Height          =   324
         Index           =   2
         Left            =   144
         TabIndex        =   6
         Top             =   840
         Width           =   2760
      End
      Begin VB.CheckBox chk�洢�ⷿ 
         Caption         =   "Ӧ���ڱ���������������(&B)"
         Height          =   324
         Index           =   1
         Left            =   144
         TabIndex        =   5
         Top             =   540
         Width           =   2712
      End
      Begin VB.CheckBox chk�洢�ⷿ 
         Caption         =   "Ӧ����������������(&A)"
         Height          =   324
         Index           =   0
         Left            =   144
         TabIndex        =   4
         Top             =   285
         Width           =   2364
      End
      Begin VB.Label lblInfor 
         Caption         =   "   ��:û�й��ϴ���Ŀ�еġ�Ӧ���������������ϡ������ڴ洢�ⷿ���ý����еġ�Ӧ�������С��������ϡ�(4)��������ѡ��"
         ForeColor       =   &H00000000&
         Height          =   870
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   2820
      End
      Begin VB.Label lblInfor 
         Caption         =   "    ����Ŀ��Ҫ�ǿ����������Ϲ���Ĵ洢�ⷿ���ý����еġ�Ӧ����...�����ܡ�"
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   0
         Left            =   165
         TabIndex        =   16
         Top             =   1800
         Width           =   2910
      End
      Begin VB.Label lblInfor 
         Caption         =   "˵��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   15
         Top             =   1440
         Width           =   2550
      End
   End
   Begin VB.Frame fraCodeMode 
      Height          =   1515
      Left            =   180
      TabIndex        =   10
      Top             =   360
      Width           =   4620
      Begin VB.OptionButton opt����ģʽ 
         Caption         =   "&3) �����+˳����"
         Height          =   210
         Index           =   2
         Left            =   900
         TabIndex        =   2
         Top             =   1155
         Width           =   3420
      End
      Begin VB.OptionButton opt����ģʽ 
         Caption         =   "&2) �������+�����+˳����"
         Height          =   210
         Index           =   1
         Left            =   900
         TabIndex        =   1
         Top             =   825
         Width           =   3420
      End
      Begin VB.OptionButton opt����ģʽ 
         Caption         =   "&1) ͬ��˳����"
         Height          =   210
         Index           =   0
         Left            =   900
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.Label lbl����ģʽ 
         Caption         =   "1������ȱʡ����ģʽ"
         Height          =   180
         Left            =   600
         TabIndex        =   18
         Top             =   0
         Width           =   1750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   60
         Picture         =   "frmStuffPara.frx":000C
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   9
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5625
      TabIndex        =   7
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6870
      TabIndex        =   8
      Top             =   4725
      Width           =   1100
   End
   Begin VB.Frame fraIncome 
      Height          =   855
      Left            =   180
      TabIndex        =   12
      Top             =   2040
      Width           =   4620
      Begin VB.ComboBox cbo������Ŀ 
         ForeColor       =   &H80000012&
         Height          =   300
         Index           =   0
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label lblIncome 
         AutoSize        =   -1  'True
         Caption         =   "2�������ʶ�Ӧȱʡ������Ŀ"
         Height          =   180
         Left            =   585
         TabIndex        =   11
         Top             =   0
         Width           =   2250
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   60
         Picture         =   "frmStuffPara.frx":08D6
         Top             =   60
         Width           =   480
      End
      Begin VB.Label LblNote 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   13
         Top             =   390
         Visible         =   0   'False
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmStuffPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnActive As Boolean
Private mintTabIndex As Integer
Private mlng������ĿID As Long
Private mstrPrivs As String
Private mrs������Ŀ As New ADODB.Recordset
Private mblnHavePriv As Boolean
Private Const mlngModule = 1711

Public Sub ShowMe(ByVal strPrivs As String, ByVal frmMain As Object)
    '----------------------------------------------------------------------------------
    '����:�����������
    '����:mstrPrivs -Ȩ�޴�
    '     frmMain-���ø�����
    '����:
    '����:���˺�
    '����:2007/12/24
    '----------------------------------------------------------------------------------
    mstrPrivs = strPrivs
    Me.Show 1, frmMain
End Sub

 
Private Sub cbo������Ŀ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'Private Sub chk�������_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'End Sub

'Private Sub chkƷ�ֹ��_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'End Sub

 
Private Sub chk�洢�ⷿ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
            zlCommFun.PressKey vbKeyTab
        End If
End Sub

'Private Sub chkƷ������_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'End Sub

Private Sub CmdCancel_Click()
    gblnIncomeItem = False
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub CmdOK_Click()
    Dim intSave As Integer
    Dim strReg As String
    
    If SaveSet = False Then Exit Sub
    gblnIncomeItem = True
    Unload Me
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2007/12/24
    '------------------------------------------------------------------------------------------
    Dim int����ģʽ As Integer, strӦ�÷�Χ As String
    Dim intSet���� As Integer
    
    '��ʽ:3λ�ַ�����,1,��������0��������,��111.���е�һλ��������,�ڶ�λ����������,����λ�������������
    strӦ�÷�Χ = IIf(chk�洢�ⷿ(0).Value = 1, "1", "0")
    strӦ�÷�Χ = strӦ�÷�Χ & IIf(chk�洢�ⷿ(1).Value = 1, "1", "0")
    strӦ�÷�Χ = strӦ�÷�Χ & IIf(chk�洢�ⷿ(2).Value = 1, "1", "0")
    If Me.opt����ģʽ(0).Value = True Then
       int����ģʽ = 0
    ElseIf Me.opt����ģʽ(1).Value = True Then
       int����ģʽ = 1
    Else
       int����ģʽ = 2
    End If
    
    If optSet�ֶ�.Value = True Then
        intSet���� = 0
    ElseIf optSet�ⷿ.Value = True Then
        intSet���� = 1
    ElseIf optsetall.Value = True Then
        intSet���� = 2
    ElseIf optSetNotall.Value = True Then
        intSet���� = 3
    End If
    
    err = 0: On Error GoTo ErrHand:
    
    gcnOracle.BeginTrans
'    Call zlDatabase.SetPara("Ʒ������ģʽ", Me.chkƷ������.Value, glngSys, mlngModule)
'    Call zlDatabase.SetPara("Ʒ�ֹ��ģʽ", Me.chkƷ�ֹ��.Value, glngSys, mlngModule)
'    Call zlDatabase.SetPara("�������ģʽ", Me.chk�������.Value, glngSys, mlngModule)
    Call zlDatabase.SetPara("�������ģʽ", int����ģʽ, glngSys, mlngModule)
    Call zlDatabase.SetPara("����Ӧ���ڵķ�Χ", strӦ�÷�Χ, glngSys, mlngModule)
    Call zlDatabase.SetPara("������Ŀ��Ӧ", cbo������Ŀ(1).ItemData(cbo������Ŀ(1).ListIndex), glngSys, mlngModule)
    Call zlDatabase.SetPara("���ķ��������Զ�����", intSet����, glngSys, mlngModule)
    
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub Form_Activate()
    If Not mblnActive Then Unload Me: Exit Sub
    
End Sub

Private Sub Form_Load()
    '�����û�Ȩ�ޣ�װ��ؼ�
    Dim intValue As Integer
    Dim strReg As String
    Dim intSet���� As Integer
    
    mblnHavePriv = IsHavePrivs(mstrPrivs, "��������")
    
    On Error GoTo errHandle
    mblnActive = False
    
'    chkƷ������.Value = IIf(Val(zlDatabase.GetPara("Ʒ������ģʽ", glngSys, mlngModule, , Array(chkƷ������), mblnHavePriv)) = 1, 1, 0)
'    chkƷ�ֹ��.Value = IIf(Val(zlDatabase.GetPara("Ʒ�ֹ��ģʽ", glngSys, mlngModule, , Array(chkƷ�ֹ��), mblnHavePriv)) = 1, 1, 0)
'    chk�������.Value = IIf(Val(zlDatabase.GetPara("�������ģʽ", glngSys, mlngModule, , Array(chk�������), mblnHavePriv)) = 1, 1, 0)
    
    
    intValue = Val(zlDatabase.GetPara("�������ģʽ", glngSys, mlngModule, , Array(opt����ģʽ(0), opt����ģʽ(1), opt����ģʽ(2), lbl����ģʽ, fraCodeMode), mblnHavePriv))
    If intValue = 0 Then
        Me.opt����ģʽ(0).Value = True: Me.opt����ģʽ(1).Value = False: Me.opt����ģʽ(2).Value = False
    ElseIf intValue = 1 Then
        Me.opt����ģʽ(0).Value = False: Me.opt����ģʽ(1).Value = True: Me.opt����ģʽ(2).Value = False
    Else
        Me.opt����ģʽ(0).Value = False: Me.opt����ģʽ(1).Value = False: Me.opt����ģʽ(2).Value = True
    End If
    '��ʽ:3λ�ַ�����,1,��������0��������,��111.���е�һλ��������,�ڶ�λ����������,����λ�������������
    strReg = zlDatabase.GetPara("����Ӧ���ڵķ�Χ", glngSys, mlngModule, , Array(fra, chk�洢�ⷿ(0), chk�洢�ⷿ(1), chk�洢�ⷿ(2)), mblnHavePriv)
        
    If Len(strReg) < 3 Then
        'Ĭ��ȫѡ��
        strReg = "111"
    End If
    chk�洢�ⷿ(0).Value = IIf(Val(Mid(strReg, 1, 1)) = 1, 1, 0)
    chk�洢�ⷿ(1).Value = IIf(Val(Mid(strReg, 2, 1)) = 1, 1, 0)
    chk�洢�ⷿ(2).Value = IIf(Val(Mid(strReg, 3, 1)) = 1, 1, 0)
    
    
    intSet���� = Val(zlDatabase.GetPara("���ķ��������Զ�����", glngSys, mlngModule, 0))
    Select Case intSet����
        Case 0
            optSet�ֶ�.Value = True
        Case 1
            optSet�ⷿ.Value = True
        Case 2
            optsetall.Value = True
        Case 3
            optSetNotall.Value = True
    End Select
    
    gstrSQL = "Select ID,����||'-'||���� ���� From ������Ŀ Where ĩ��=1"
    zlDatabase.OpenRecordset mrs������Ŀ, gstrSQL, Me.Caption
    With mrs������Ŀ
        If .EOF Then
            MsgBox "���ʼ��������Ŀ��������Ŀ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    mlng������ĿID = Val(zlDatabase.GetPara("������Ŀ��Ӧ", glngSys, mlngModule, "", Array(lblIncome, fraIncome, cbo������Ŀ, LblNote(0)), mblnHavePriv))
    mintTabIndex = 10

    Call AddCons("��������")
    mblnActive = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddCons(ByVal strName As String)
    Dim intIdx As Integer
    intIdx = LblNote.UBound + 1
    Load LblNote(intIdx)
    Load cbo������Ŀ(intIdx)
    
    mintTabIndex = mintTabIndex + 1
    With LblNote(intIdx)
        .Caption = strName
        .TabIndex = mintTabIndex
        .Container = fraIncome
        .Top = IIf(intIdx = 1, LblNote(0).Top, LblNote(intIdx - 1).Top) + IIf(intIdx = 1, 0, LblNote(0).Height + 200)
        .Left = LblNote(0).Left + LblNote(0).Width - .Width
        .Visible = True
    End With
    mintTabIndex = mintTabIndex + 1
    With cbo������Ŀ(intIdx)
        .Container = fraIncome
        .Left = cbo������Ŀ(0).Left
        .Top = IIf(intIdx = 1, cbo������Ŀ(0).Top, cbo������Ŀ(intIdx - 1).Top) + IIf(intIdx = 1, 0, cbo������Ŀ(0).Height + 100)
        .TabIndex = mintTabIndex
        .Visible = True
    End With
    Call AddItem(cbo������Ŀ(intIdx), strName)
End Sub

Private Sub AddItem(ByVal cboObj As ComboBox, ByVal strName As String)
    Dim i As Integer

    With mrs������Ŀ
        .MoveFirst
        Do While Not .EOF
            cboObj.AddItem !����
            cboObj.ItemData(cboObj.NewIndex) = !Id
            .MoveNext
        Loop
        For i = 0 To cboObj.ListCount - 1
            If cboObj.ItemData(i) = mlng������ĿID Then
                cboObj.ListIndex = i
                Exit Sub
            End If
        Next
        For i = 0 To cboObj.ListCount - 1
            If strName = "��������" Then    '��������������Ǿ����Ҳ��Ϸѵģ�û�����Һ��в��ϵģ����߶�û�е���Ĭ��ѡ�е�һ��
                If cboObj.List(i) Like "*���Ϸ�" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
        For i = 0 To cboObj.ListCount - 1
            If strName = "��������" Then
                If cboObj.List(i) Like "*����*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
        cboObj.ListIndex = 0
    End With
End Sub

Private Sub opt����ģʽ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
