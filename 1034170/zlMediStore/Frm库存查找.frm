VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm������ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ҩƷ"
   ClientHeight    =   2580
   ClientLeft      =   3135
   ClientTop       =   4320
   ClientWidth     =   5985
   Icon            =   "Frm������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Pic���� 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   6135
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton CmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   600
         Picture         =   "Frm������.frx":020A
         TabIndex        =   19
         Top             =   2160
         Width           =   1100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
         Height          =   1575
         Left            =   1170
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "��"
         Height          =   240
         Left            =   5070
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1695
         Width           =   255
      End
      Begin VB.TextBox TxtSelect���� 
         Height          =   300
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1665
         Width           =   4185
      End
      Begin VB.CommandButton Cmd���� 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2820
         Picture         =   "Frm������.frx":0354
         TabIndex        =   15
         Top             =   2160
         Width           =   1100
      End
      Begin VB.CommandButton Cmd���� 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4230
         Picture         =   "Frm������.frx":049E
         TabIndex        =   17
         Top             =   2160
         Width           =   1100
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   2
         Top             =   840
         Width           =   1515
      End
      Begin VB.TextBox Txtͨ������ 
         Height          =   300
         Left            =   3840
         MaxLength       =   40
         TabIndex        =   1
         Top             =   390
         Width           =   1515
      End
      Begin VB.TextBox TxtҩƷ���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   390
         Width           =   1515
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   3
         Top             =   840
         Width           =   1515
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1260
         Width           =   1515
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1290
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblָ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ָ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   14
         Top             =   1350
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   13
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   12
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   11
         Top             =   900
         Width           =   360
      End
      Begin VB.Label LblҩƷ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   10
         Top             =   450
         Width           =   360
      End
      Begin VB.Label Lblͨ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͨ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   9
         Top             =   450
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StrTmp As String
Public StrBit As Byte '�ó�����ҵ�ƥ�䷽ʽ
Dim rsTmp As ADODB.Recordset
Private mfrmMain As Form    '������

Private Type Type_SQLCondition
    strͨ���� As String
    str���� As String
    str���� As String
    str���� As String
    str��� As String
    str���� As String
End Type

Private SQLCondition As Type_SQLCondition

Public Function GetSearch(ByVal FrmMain As Form, _
    ByRef strͨ���� As String, _
    ByRef str���� As String, _
    ByRef str���� As String, _
    ByRef str���� As String, _
    ByRef str��� As String, _
    ByRef str���� As String) As String
    StrTmp = ""
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = StrTmp
    
    strͨ���� = SQLCondition.strͨ����
    str���� = SQLCondition.str����
    str���� = SQLCondition.str����
    str���� = SQLCondition.str����
    str��� = SQLCondition.str���
    str���� = SQLCondition.str����
    
End Function
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub CmdSelect_Click()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ����,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
    Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-ҩƷ������", gstrNodeNo)
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "����ҩƷ������"
        .lngMode = 1
        .Show 1, Me
        If .BlnSuccess = True Then
            TxtSelect����.Tag = 1
            TxtSelect����.Text = .CurrentName
            Cmd����.SetFocus
        End If
    End With
    Unload FrmSelect
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd����_Click()
    If LTrim(Txtͨ������) = "" And LTrim(TxtҩƷ����) = "" And LTrim(Txt����) = "" & _
        LTrim(Txt����) = "" And LTrim(txt���) = "" And LTrim(TxtSelect����) = "" Then MsgBox "����������һ����Ϣ��", vbInformation, gstrSysName
    StrTmp = ""
    If LTrim(Txtͨ������) <> "" Then StrTmp = StrTmp & " And A.���� like [1] "
    If LTrim(TxtҩƷ����) <> "" Then StrTmp = StrTmp & " And A.���� like [2] "
    If LTrim(Txt����) <> "" Then StrTmp = StrTmp & " And B.���� like [3] "
    If LTrim(Txt����) <> "" Then StrTmp = StrTmp & " And B.���� like [4] "
    If LTrim(txt���) <> "" Then StrTmp = StrTmp & " And upper(A.���) like [5] "
    If LTrim(TxtSelect����) <> "" Then StrTmp = StrTmp & " And upper(A.����) like [6] "
    
    SQLCondition.strͨ���� = IIf(StrBit = "0", "%", "") & LTrim(Txtͨ������) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(TxtҩƷ����)) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(Txt����)) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(Txt����)) & "%"
    SQLCondition.str��� = IIf(StrBit = "0", "%", "") & UCase(LTrim(txt���)) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(TxtSelect����)) & "%"
    
    Unload Me
End Sub

Private Sub Cmd����_Click()
    StrTmp = ""
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    StrBit = GetSetting(appName:="ZLSOFT", Section:="����ģ��\����", Key:="����ƥ��", Default:="0")
End Sub

Private Sub Pic����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub TxtSelect����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(TxtSelect����) = "" Then Exit Sub
        TxtSelect���� = UCase(TxtSelect����)
        
        gstrSQL = "Select ����,����,���� From ҩƷ������ Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(����) like [1] or Upper(����) like [2]) Order By ����"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ҩƷ������]", _
                        IIf(gstrMatchMethod = "0", "%", "") & TxtSelect���� & "%", _
                        TxtSelect���� & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = TxtSelect����.Top - .Height
                    .Left = TxtSelect����.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                End With
            Else
                TxtSelect���� = IIf(IsNull(!����), "", !����)
                TxtSelect����.Tag = 1
                Cmd����.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtSelect����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txt���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Txtͨ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub TxtҩƷ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            TxtSelect����.Text = .TextMatrix(.Row, 2)
            TxtSelect����.Tag = 1
            Cmd����.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub
