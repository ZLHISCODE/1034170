VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMediPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmMediPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt������ 
      Height          =   300
      Left            =   5400
      TabIndex        =   25
      Top             =   3240
      Width           =   720
   End
   Begin MSComCtl2.UpDown udg������ 
      Height          =   300
      Left            =   6120
      TabIndex        =   24
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txt������"
      BuddyDispid     =   196609
      OrigLeft        =   6720
      OrigTop         =   3720
      OrigRight       =   6975
      OrigBottom      =   4020
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Frame fra���� 
      Caption         =   "3��ҩƷ���������Զ�����"
      ForeColor       =   &H00800000&
      Height          =   1080
      Left            =   120
      TabIndex        =   18
      Top             =   2300
      Width           =   4035
      Begin VB.OptionButton optAllNotSet 
         Caption         =   "ҩ���ҩ����������"
         Height          =   200
         Left            =   1920
         TabIndex        =   22
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optAllSet 
         Caption         =   "ҩ���ҩ������"
         Height          =   200
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton opt�ֶ� 
         Caption         =   "�ֹ����÷�������"
         Height          =   200
         Left            =   120
         TabIndex        =   20
         Top             =   390
         Width           =   1735
      End
      Begin VB.OptionButton optOnlyҩ�� 
         Caption         =   "��ҩ�����"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fra�ۼ۷�ʽ 
      Caption         =   "2����������ۼۼ��㷽ʽ"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   1300
      Width           =   4035
      Begin VB.OptionButton opt�ֶμӳ� 
         Caption         =   "���ֶμӳɼ����ۼ�"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   3735
      End
      Begin VB.OptionButton optһ��ӳ� 
         Caption         =   "��һ��ӳ��ʼ����ۼ�"
         Height          =   200
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraIncome 
      Height          =   1005
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4035
      Begin VB.ComboBox cbo������Ŀ 
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   315
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label LblNote 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   14
         Top             =   390
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   60
         Picture         =   "frmMediPara.frx":000C
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lblIncome 
         AutoSize        =   -1  'True
         Caption         =   "1�������ʶ�Ӧȱʡ������Ŀ"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   585
         TabIndex        =   13
         Top             =   0
         Width           =   2250
      End
   End
   Begin VB.Frame frmStockRange 
      Caption         =   "4�����ô洢�ⷿʱ����Ӧ���ڵķ�Χ"
      ForeColor       =   &H00800000&
      Height          =   3030
      Left            =   4215
      TabIndex        =   3
      Top             =   105
      Width           =   3585
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "��Ӧ���ڵ�ǰѡ���ҩƷ(&1)"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   285
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "Ӧ�������е�ǰѡ���ͬƷ��ҩƷ(&2)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   9
         Top             =   660
         Value           =   1  'Checked
         Width           =   3270
      End
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "Ӧ�������е�ǰѡ���ͬ����ҩƷ(&3)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   1035
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "Ӧ�������е�ǰѡ���ͬ����ҩƷ(&4)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   7
         Top             =   1410
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "Ӧ��������ͬ����ҩƷ(&5)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   4
         Left            =   210
         TabIndex        =   6
         Top             =   1785
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "Ӧ�������е�ǰ�����µ�ҩƷ(&6)"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   5
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2985
      End
      Begin VB.Label lblComment 
         Caption         =   "��ʾ��û��ѡ�񵽵�Ӧ�÷�Χ�����ô洢�ⷿʱ������ѡ��"
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   2880
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5325
      TabIndex        =   0
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6675
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      Caption         =   "�����볤��"
      Height          =   180
      Left            =   4320
      TabIndex        =   23
      Top             =   3300
      Width           =   900
   End
End
Attribute VB_Name = "frmMediPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnActive As Boolean
Private intTabIndex As Integer
Private lng����ҩ As Long, lng�в�ҩ As Long, lng�г�ҩ As Long
Private strPrivs As String
Private rs������Ŀ As New ADODB.Recordset
Dim mblnSetPara As Boolean      '�Ƿ���в�������Ȩ��
Private Sub SetFramSize()
    Dim dlbTopTmp As Double
    Dim n As Integer
    
    frmStockRange.Top = fraIncome.Top
    lblComment.Top = chkӦ�÷�Χ(5).Top + chkӦ�÷�Χ(5).Height + 200 'frmStockRange.Height - lblComment.Height - 100
    fra�ۼ۷�ʽ.Top = fraIncome.Top + fraIncome.Height + 100
    fra����.Top = fra�ۼ۷�ʽ.Height + fra�ۼ۷�ʽ.Top + 100
    frmStockRange.Height = udg������.Top + udg������.Height - 400
    lbl������.Top = frmStockRange.Top + frmStockRange.Height + 200
    txt������.Top = lbl������.Top - 60
    udg������.Top = txt������.Top
    
    
    With cmdOK
        .Top = fra����.Top + fra����.Height + 150
        .TabIndex = intTabIndex
    End With
    With cmdCancel
        .Top = cmdOK.Top
        .TabIndex = intTabIndex + 1
    End With
    With cmdHelp
        .Top = cmdOK.Top
        .TabIndex = intTabIndex + 2
    End With
    Me.Height = cmdOK.Top + cmdOK.Height + 550
'    dlbTopTmp = lblComment.Top - chkӦ�÷�Χ(0).Top
'
'    dlbTopTmp = Int(dlbTopTmp / 6)
    
'    For n = 1 To 5
'        chkӦ�÷�Χ(n).Top = chkӦ�÷�Χ(n - 1).Top + dlbTopTmp
'    Next
End Sub

Public Sub ShowMe(ByVal strPrivss As String, ByVal frmParent As Object)
    strPrivs = strPrivss
    Me.Show 1, frmParent
End Sub

'Private Sub chkƷ������_Click()
'    If Me.chkƷ������.Value = 1 Then
'        Me.chkƷ�ֹ��.Value = 0: Me.chkƷ�ֹ��.Enabled = False
'    Else
'        Me.chkƷ�ֹ��.Enabled = True
'    End If
'End Sub

Private Sub cmdCancel_Click()
    gblnIncomeItem = False
    Unload Me
End Sub
Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim intSave As Integer
    Dim strRange As String
    Dim intSetMethod As Integer '�ⷿ���÷�ʽ,0-�ֹ����÷������ԣ�Ĭ��ֵ����1-��ҩ�������2-ҩ���ҩ��������3-ҩ���ҩ����������
    Dim n As Integer
        
'    If Me.chkƷ������.Value = 0 Then
'        If Me.chkƷ�ֹ��.Value = 0 Then
'            zldatabase.SetPara "Ʒ������ģʽ", 0, glngSys, 1023
'        Else
'            zldatabase.SetPara "Ʒ������ģʽ", 2, glngSys, 1023
'        End If
'    Else
'        zldatabase.SetPara "Ʒ������ģʽ", 1, glngSys, 1023
'    End If
'    If Me.chk�������.Value = 0 Then
'        zldatabase.SetPara "�������ģʽ", 0, glngSys, 1023
'    Else
'        zldatabase.SetPara "�������ģʽ", 1, glngSys, 1023
'    End If
    
    If opt�ֶ�.Value = True Then
        intSetMethod = 0
    ElseIf optOnlyҩ��.Value = True Then
        intSetMethod = 1
    ElseIf optAllSet.Value = True Then
        intSetMethod = 2
    ElseIf optAllNotSet.Value = True Then
        intSetMethod = 3
    End If
    
    For intSave = 1 To LblNote.UBound
        zlDatabase.SetPara intSave, cbo������Ŀ(intSave).ItemData(cbo������Ŀ(intSave).ListIndex), glngSys, 1023
    Next
    
    For n = 0 To chkӦ�÷�Χ.Count - 1
        strRange = strRange & chkӦ�÷�Χ(n).Value
    Next
    
    If optһ��ӳ�.Value = True Then
        zlDatabase.SetPara "�ۼ۰��ӳɼ���", 0, glngSys, 1023
    Else
        zlDatabase.SetPara "�ۼ۰��ӳɼ���", 1, glngSys, 1023
    End If
    
    zlDatabase.SetPara "Ӧ�÷�Χ", strRange, glngSys, 1023
    zlDatabase.SetPara "ҩƷ���������Զ�����", intSetMethod, glngSys, 1023
    zlDatabase.SetPara "������", txt������.Text, glngSys, 1023
    
    gblnIncomeItem = True
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not blnActive Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strRange As String
    Dim n As Integer
    Dim intƷ������ As Integer
    Dim int������� As Integer
    Dim strTmp As String
    Dim intTmp As Integer
    Dim int�ۼ� As Integer
    Dim intSet���� As Integer   '�������÷�ʽ 0-�ֹ����÷������ԣ�Ĭ��ֵ����1-��ҩ�������2-ҩ���ҩ��������3-ҩ���ҩ����������
    Dim rsTemp As ADODB.Recordset
    
    mblnSetPara = InStr(strPrivs, "��������") > 0

    '�����û�Ȩ�ޣ�װ��ؼ�
    On Error GoTo errHandle
    intTabIndex = 2
    blnActive = False
    
    gstrSql = "select nvl(max(length(����)),0) ���� from �շ���Ŀ���� where ����=3"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "�����볤��")
    
    If rsTemp!���� = 0 Then
        udg������.Min = 7
    Else
        udg������.Min = rsTemp!����
    End If
    udg������.Max = 40
'    intƷ������ = Val(zldatabase.GetPara("Ʒ������ģʽ", glngSys, 1023, 0, Array(chkƷ������, chkƷ�ֹ��), mblnSetPara))
'    Select Case intƷ������
'    Case 1
'        Me.chkƷ������.Value = 1
'        Me.chkƷ�ֹ��.Value = 0: Me.chkƷ�ֹ��.Enabled = False
'    Case 2
'        Me.chkƷ������.Value = 0
'        Me.chkƷ�ֹ��.Value = 1: Me.chkƷ�ֹ��.Enabled = True And mblnSetPara = True
'    Case Else
'        Me.chkƷ������.Value = 0
'        Me.chkƷ�ֹ��.Enabled = True And mblnSetPara = True
'    End Select
    
'    int������� = Val(zldatabase.GetPara("�������ģʽ", glngSys, 1023, 0, Array(chk�������), mblnSetPara))
    
'    If int������� = 0 Then
'        Me.chk�������.Value = 0
'    Else
'        Me.chk�������.Value = 1
'    End If

    gstrSql = "Select ID,����||'-'||���� ���� From ������Ŀ Where ĩ��=1"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rs������Ŀ = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rs������Ŀ
        If .EOF Then
            MsgBox "���ʼ��������Ŀ��������Ŀ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    lng����ҩ = Val(zlDatabase.GetPara("����ҩ������Ŀ", glngSys, 1023, 0))
    lng�г�ҩ = Val(zlDatabase.GetPara("�г�ҩ������Ŀ", glngSys, 1023, 0))
    lng�в�ҩ = Val(zlDatabase.GetPara("�в�ҩ������Ŀ", glngSys, 1023, 0))
    
    If strPrivs Like "*����ҩ*" Then Call AddCons("����ҩ")
    If strPrivs Like "*�г�ҩ*" Then Call AddCons("�г�ҩ")
    If strPrivs Like "*�в�ҩ*" Then Call AddCons("�в�ҩ")
    
    For n = 0 To cbo������Ŀ.UBound
        If n = 0 Then strTmp = "����ҩ������Ŀ"
        If n = 1 Then strTmp = "�г�ҩ������Ŀ"
        If n = 2 Then strTmp = "�в�ҩ������Ŀ"
        
        intTmp = Val(zlDatabase.GetPara(strTmp, glngSys, 1023, 0, Array(cbo������Ŀ(n)), mblnSetPara))
    Next
    
    strRange = zlDatabase.GetPara("Ӧ�÷�Χ", glngSys, 1023, "111111", Array(frmStockRange, chkӦ�÷�Χ(1), chkӦ�÷�Χ(2), chkӦ�÷�Χ(3), chkӦ�÷�Χ(4), chkӦ�÷�Χ(5)), mblnSetPara)
    For n = 1 To chkӦ�÷�Χ.Count - 1
        chkӦ�÷�Χ(n).Value = Mid(strRange, n + 1, 1)
    Next

    
    int�ۼ� = Val(zlDatabase.GetPara("�ۼ۰��ӳɼ���", glngSys, 1023, 0))
    
    If int�ۼ� = 0 Then
        optһ��ӳ�.Value = True
        opt�ֶμӳ�.Value = False
    Else
        optһ��ӳ�.Value = False
        opt�ֶμӳ�.Value = True
    End If
    
    intSet���� = Val(zlDatabase.GetPara("ҩƷ���������Զ�����", glngSys, 1023, 0))
    Select Case intSet����
        Case 0
            opt�ֶ�.Value = True
        Case 1
            optOnlyҩ��.Value = True
        Case 2
            optAllSet.Value = True
        Case 3
            optAllNotSet.Value = True
    End Select
    txt������.Text = Val(zlDatabase.GetPara("������", glngSys, 1023, 7))
    
    '���ô����С�����ؼ�λ��
    fraIncome.Height = cbo������Ŀ(cbo������Ŀ.UBound).Top + cbo������Ŀ(0).Height + 200
    
    Call SetFramSize
    
    blnActive = True
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
    
    LblNote(intIdx).ForeColor = LblNote(0).ForeColor
    cbo������Ŀ(intIdx).ForeColor = cbo������Ŀ(0).ForeColor
    
    intTabIndex = intTabIndex + 1
    With LblNote(intIdx)
        .Caption = strName
        .TabIndex = intTabIndex
        .Container = fraIncome
        .Top = IIf(intIdx = 1, LblNote(0).Top, LblNote(intIdx - 1).Top) + IIf(intIdx = 1, 0, LblNote(0).Height + 200)
        .Left = LblNote(0).Left + LblNote(0).Width - .Width
        .Visible = True
    End With
    intTabIndex = intTabIndex + 1
    With cbo������Ŀ(intIdx)
        .Container = fraIncome
        .Left = cbo������Ŀ(0).Left
        .Top = IIf(intIdx = 1, cbo������Ŀ(0).Top, cbo������Ŀ(intIdx - 1).Top) + IIf(intIdx = 1, 0, cbo������Ŀ(0).Height + 100)
        .TabIndex = intTabIndex
        .Visible = True
    End With
    Call AddItem(cbo������Ŀ(intIdx), strName)
End Sub

Private Sub AddItem(ByVal cboObj As ComboBox, ByVal strName As String)
'    Dim lngIdx As Integer
    Dim i As Integer
    
'    Select Case strName
'    Case "����ҩ"
'        lngIdx = lng����ҩ
'    Case "�г�ҩ"
'        lngIdx = lng�г�ҩ
'    Case "�в�ҩ"
'        lngIdx = lng�в�ҩ
'    End Select
    
    With rs������Ŀ
        .MoveFirst
        Do While Not .EOF
            cboObj.AddItem !����
            cboObj.ItemData(cboObj.NewIndex) = !ID
            .MoveNext
        Loop

        For i = 0 To cboObj.ListCount - 1
            If strName = "����ҩ" Then
                If cboObj.List(i) Like "*��ҩ*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
            If strName = "�г�ҩ" Then
                If cboObj.List(i) Like "*�г�ҩ*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
            If strName = "�в�ҩ" Then
                If cboObj.List(i) Like "*��ҩ*" Or cboObj.List(i) Like "*��ҩ*" Then
                    cboObj.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
    End With
End Sub


Private Sub txt������_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub udg������_Change()
    txt������.Text = udg������.Value
End Sub
