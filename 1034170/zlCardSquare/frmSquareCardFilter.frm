VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSquareCardFilter 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   4590
      Left            =   60
      ScaleHeight     =   4590
      ScaleWidth      =   3885
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   3885
      Begin VB.CheckBox chk��ͣ�� 
         Caption         =   "����ͣ�õĿ�"
         Height          =   450
         Left            =   645
         TabIndex        =   20
         Top             =   3720
         Width           =   2685
      End
      Begin VB.CommandButton cmdˢ�� 
         Caption         =   "����(&F)"
         Height          =   390
         Left            =   2700
         TabIndex        =   19
         Top             =   4185
         Width           =   1050
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   2
         Left            =   630
         TabIndex        =   17
         Top             =   3285
         Width           =   3105
      End
      Begin VB.TextBox txtEdit 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   15
         Top             =   615
         Width           =   1290
      End
      Begin VB.TextBox txtEdit 
         Height          =   315
         Index           =   0
         Left            =   630
         TabIndex        =   14
         Top             =   630
         Width           =   1290
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   135
         Width           =   3090
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "���������ڹ���"
         Height          =   375
         Index           =   1
         Left            =   45
         TabIndex        =   2
         Top             =   1950
         Width           =   1665
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "���������ڹ���"
         Height          =   375
         Index           =   0
         Left            =   45
         TabIndex        =   1
         Top             =   1125
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   0
         Left            =   630
         TabIndex        =   3
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   111542275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   0
         Left            =   2430
         TabIndex        =   5
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   111542275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   1
         Left            =   630
         TabIndex        =   6
         Top             =   2310
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   111542275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   7
         Top             =   2310
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   111542275
         CurrentDate     =   37007
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   3
         Left            =   630
         TabIndex        =   4
         Top             =   2865
         Width           =   3105
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "�쿨��"
         Height          =   180
         Index           =   3
         Left            =   15
         TabIndex        =   18
         Top             =   3360
         Width           =   540
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   2025
         TabIndex        =   16
         Top             =   690
         Width           =   180
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   13
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   15
         TabIndex        =   11
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   2040
         TabIndex        =   10
         Top             =   1545
         Width           =   180
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   4
         Left            =   2040
         TabIndex        =   9
         Top             =   2370
         Width           =   180
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   1
         Left            =   15
         TabIndex        =   8
         Top             =   2940
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSquareCardFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Private mArrFilter As Variant
Private mstrPrivs As String, mlngModule As Long
Private Enum mtxtIdx
    idx_�������� = 0
    idx_�������� = 1
    idx_�쿨�� = 2
    idx_������ = 3
End Enum
'--------------------------------------------------------------------------------------------------------
Public Event zlRefreshCon(ByVal arrFilter As Variant)

Private Function GetFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-18 14:27:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, strReg As String
    
    '������ѯ����
    Set cllFilter = New Collection
    If cbo����.ListIndex = 0 Then
        cllFilter.Add "����", "������"
    Else
        cllFilter.Add Split(cbo����.Text, "-")(1), "������"
    End If
    cllFilter.Add Array(Trim(txtEdit(mtxtIdx.idx_��������).Text), Trim(txtEdit(mtxtIdx.idx_��������).Text)), "���ŷ�Χ"
    
    cllFilter.Add Trim(txtEdit(mtxtIdx.idx_�쿨��).Text), "�쿨��"
    cllFilter.Add Trim(txtEdit(mtxtIdx.idx_������).Text), "������"
    
    If chkDate(0).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(0).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(0).Value, "yyyy-mm-dd") & " 23:59:59"), "����ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "����ʱ��"
    End If
    If chkDate(1).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(1).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(1).Value, "yyyy-mm-dd") & " 23:59:59"), "����ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "����ʱ��"
    End If
    cllFilter.Add IIf(chk��ͣ��.Value = 1, 1, 0), "����ͣ�ÿ�"
    Set mArrFilter = cllFilter
    
End Function

Private Sub cmdˢ��_Click()
    If chkDate(0).Value = 0 And chkDate(1).Value = 0 Then
        ShowMsgbox "����ȷ��һ��ʱ�䷶Χ,����!"
        Exit Sub
    End If
    Call GetFilter
    RaiseEvent zlRefreshCon(mArrFilter)
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2009-11-18 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    dtpEndDate(0).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEndDate(1).MaxDate = dtpEndDate(0).MaxDate

    dtpEndDate(0).Value = dtpEndDate(0).MaxDate
    dtpEndDate(1).Value = dtpEndDate(0).MaxDate
    
    dtpStartDate(0).Value = Format(DateAdd("d", -1, zlDatabase.Currentdate), "yyyy-mm-dd")  'ȱʡ��7����
    dtpStartDate(1).Value = dtpStartDate(0).Value
    
    On Error GoTo errHandle
    
    '���ؿ�����:���û�а�ȱʡ��־,�����п�Ϊ׼,����ȱʡΪ׼
    gstrSQL = "Select ����, ����, ȱʡ���, ȱʡ�ۿ�, ȱʡ��־ From ���ѿ�����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With cbo����
        .Clear
        .AddItem "���п�"
        .ListIndex = .NewIndex
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
            If Val(Nvl(rsTemp!ȱʡ��־)) = 1 Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub chkDate_Click(Index As Integer)
    Select Case Index
    Case 0
        If chkDate(Index).Value = 0 Then
           If chkDate(1).Value = 0 Then chkDate(1).Value = 1
        End If
    Case 1
        If chkDate(Index).Value = 0 Then
           If chkDate(0).Value = 0 Then chkDate(0).Value = 1
        End If
    End Select
    dtpStartDate(Index).Enabled = chkDate(Index).Value = 1
    dtpEndDate(Index).Enabled = chkDate(Index).Value = 1
    
End Sub

Private Sub chkDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEndDate_Change(Index As Integer)
     If dtpEndDate(Index).Value > dtpStartDate(Index).MaxDate Then dtpEndDate(Index).Value = dtpStartDate(Index).MaxDate
    If dtpEndDate(Index).Value < dtpStartDate(Index).Value Then
        dtpStartDate(Index).Value = dtpEndDate(Index).Value
    End If
End Sub
Private Sub dtpEndDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpStartDate_Change(Index As Integer)
    If dtpStartDate(Index).Value > dtpEndDate(Index).MaxDate Then dtpStartDate(Index).Value = dtpEndDate(Index).MaxDate
    If dtpEndDate(Index).Value < dtpStartDate(Index).Value Then
        dtpEndDate(Index).Value = dtpStartDate(Index).Value
    End If
End Sub

Private Sub dtpStartDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs: mlngModule = glngModul
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picFilter
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight
    End With
End Sub
Private Sub picFilter_Resize()
    Err = 0: On Error Resume Next
    With picFilter
         cmdˢ��.Left = .ScaleLeft + .ScaleWidth - cmdˢ��.Width - 50
    End With
End Sub
Public Sub Init����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���������
    '����:���˺�
    '����:2009-11-18 14:48:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitData
End Sub
Public Property Get GetFilterCon() As Variant
    Call GetFilter
    Set GetFilterCon = mArrFilter
End Property
Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
End Sub
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    Select Case Index
    Case mtxtIdx.idx_������
        If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case mtxtIdx.idx_�쿨��
        If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case Else
        '���ڿ��Ų�֪����,�����޷���λ
    End Select
End Sub
Public Sub ReActionFilter()
    '���½ɻ����
     cmdˢ��_Click
End Sub


