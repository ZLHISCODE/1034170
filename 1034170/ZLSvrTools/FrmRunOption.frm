VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRunOption 
   BackColor       =   &H80000005&
   Caption         =   "ϵͳ����ѡ��"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "FrmRunOption.frx":0000
   ScaleHeight     =   7335
   ScaleWidth      =   10230
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   255
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
      Begin VB.Image imgMain 
         Height          =   480
         Left            =   0
         Picture         =   "FrmRunOption.frx":04F9
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   6075
      Left            =   930
      TabIndex        =   13
      Top             =   570
      Width           =   5730
      Begin VB.CheckBox chkShutDown 
         BackColor       =   &H80000005&
         Caption         =   "����ر������ĵ���̨"
         Height          =   255
         Left            =   255
         TabIndex        =   30
         Tag             =   "24"
         Top             =   5070
         Width           =   4455
      End
      Begin VB.CheckBox chkSpecial 
         BackColor       =   &H80000005&
         Caption         =   "���Ӷȿ��ƣ����ٰ���һ�����֡���ĸ��������ţ�"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Tag             =   "23"
         Top             =   4200
         Width           =   4455
      End
      Begin VB.CheckBox chkLenCtrl 
         BackColor       =   &H80000005&
         Caption         =   "�������볤�ȿ���"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Tag             =   "20"
         Top             =   3855
         Width           =   1750
      End
      Begin VB.TextBox txtLen 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   23
         Tag             =   "21"
         Text            =   "3"
         Top             =   3840
         Width           =   300
      End
      Begin VB.TextBox txtLen 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   2910
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "22"
         Text            =   "12"
         Top             =   3840
         Width           =   300
      End
      Begin VB.TextBox Txt����·�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2070
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "6"
         Top             =   3090
         Width           =   3195
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "��"
         Height          =   300
         Left            =   5280
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3090
         Width           =   285
      End
      Begin VB.CheckBox Chkʹ����־ 
         BackColor       =   &H80000005&
         Caption         =   "������־��¼(&S)"
         Height          =   210
         Left            =   240
         TabIndex        =   0
         Tag             =   "1"
         Top             =   315
         Width           =   1695
      End
      Begin VB.TextBox Txt������־�����Ŀ�� 
         Height          =   300
         Left            =   1965
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "2"
         Top             =   547
         Width           =   1755
      End
      Begin VB.CheckBox Chk�Ƿ��¼���д��� 
         BackColor       =   &H80000005&
         Caption         =   "��¼���д���(&A)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1230
         Width           =   1695
      End
      Begin VB.TextBox Txt������־�����Ŀ�� 
         Height          =   300
         Left            =   1965
         MaxLength       =   18
         TabIndex        =   5
         Tag             =   "4"
         Top             =   1440
         Width           =   1755
      End
      Begin VB.TextBox Txt��Ϣ�����Ŀ�� 
         Height          =   300
         Left            =   1965
         MaxLength       =   18
         TabIndex        =   7
         Tag             =   "5"
         Top             =   2145
         Width           =   1755
      End
      Begin MSComCtl2.UpDown udLen 
         Height          =   270
         Index           =   1
         Left            =   3210
         TabIndex        =   25
         Top             =   3840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtLen(1)"
         BuddyDispid     =   196615
         BuddyIndex      =   1
         OrigLeft        =   3240
         OrigTop         =   3855
         OrigRight       =   3495
         OrigBottom      =   4110
         Max             =   16
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown udLen 
         Height          =   270
         Index           =   0
         Left            =   2340
         TabIndex        =   29
         Top             =   3840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtLen(0)"
         BuddyDispid     =   196615
         BuddyIndex      =   0
         OrigLeft        =   2370
         OrigTop         =   3855
         OrigRight       =   2625
         OrigBottom      =   4110
         Max             =   16
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.Label lblShutDown 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRunOption.frx":227B
         ForeColor       =   &H8000000D&
         Height          =   510
         Left            =   525
         TabIndex        =   31
         Top             =   5400
         Width           =   4350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRunOption.frx":22BD
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   480
         TabIndex        =   28
         Top             =   4545
         Width           =   3960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000005&
         Caption         =   "-->"
         Height          =   135
         Left            =   2640
         TabIndex        =   26
         Top             =   3915
         Width           =   375
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(��ѡ��������ϵ�ApplyĿ¼��Ϊ����·��)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   21
         Top             =   3480
         Width           =   3510
      End
      Begin VB.Label Lbl����·�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL������·��(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   3150
         Width           =   1800
      End
      Begin VB.Label LblOption1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(�Ƿ��Զ���¼�û���ʹ��ϵͳ�����)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1950
         TabIndex        =   18
         Top             =   315
         Width           =   3060
      End
      Begin VB.Label Lbl������־��Ŀ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��־��ౣ������(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1710
      End
      Begin VB.Label LblOption2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(ʹ����־��ౣ�������������ʱϵͳ���Զ�ɾ����ʱ�ļ�¼)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   870
         Width           =   5040
      End
      Begin VB.Label LblOption3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(�Ƿ��¼ʹ�ù����з����ĸ��ִ���)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1950
         TabIndex        =   16
         Top             =   1230
         Width           =   3060
      End
      Begin VB.Label Lbl������־�����Ŀ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ౣ������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   1490
         Width           =   1710
      End
      Begin VB.Label LblOption4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(�����¼��ౣ�������������ʱϵͳ���Զ�ɾ����ʱ�ļ�¼)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   1770
         Width           =   5040
      End
      Begin VB.Label Lbl��Ϣ�����Ŀ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϣ�����������(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   2205
         Width           =   1710
      End
      Begin VB.Label lblOption5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(��Ϣ����ܱ��������������ʱϵͳ�����Զ�ɾ����       ����Ϊ0ʱ��ʾ���ñ���)"
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   480
         TabIndex        =   14
         Top             =   2520
         Width           =   4680
      End
   End
   Begin VB.CommandButton Cmd��ԭ 
      Cancel          =   -1  'True
      Caption         =   "��ԭ(&R)"
      Height          =   350
      Left            =   2190
      TabIndex        =   12
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����(&O)"
      Height          =   350
      Left            =   900
      TabIndex        =   11
      Top             =   6750
      Width           =   1100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ����ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   20
      Top             =   150
      Width           =   1440
   End
End
Attribute VB_Name = "FrmRunOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RecOption As New ADODB.Recordset

Private Sub chkLenCtrl_Click()
    Dim blnEnabled  As Boolean
    blnEnabled = (chkLenCtrl.value = 1)
    txtLen(0).Enabled = blnEnabled
    txtLen(1).Enabled = blnEnabled
    udLen(0).Enabled = blnEnabled
    udLen(1).Enabled = blnEnabled
    Cmd����.Enabled = True
End Sub

Private Sub chkShutDown_Click()
    Cmd����.Enabled = True
End Sub

Private Sub chkSpecial_Click()
    Cmd����.Enabled = True
End Sub

Private Sub cmdSelect_Click()
    Dim strPath As String
    strPath = OpenFolder(Me, "Excel������·����")
    If strPath = "" Then Exit Sub
    Txt����·�� = strPath
    Cmd����.Enabled = True
End Sub

Private Sub Cmd����_Click()
    If Txt������־�����Ŀ��.Enabled = True And Val(Txt������־�����Ŀ��.Text) > 10 ^ 8 Then
        MsgBox "������־�����Ŀ��̫��", vbInformation, gstrSysName
        Txt������־�����Ŀ��.SetFocus
        Exit Sub
    End If
    If Txt������־�����Ŀ��.Enabled = True And Val(Txt������־�����Ŀ��.Text) > 10 ^ 8 Then
        MsgBox "������־�����Ŀ��̫��", vbInformation, gstrSysName
        Txt������־�����Ŀ��.SetFocus
        Exit Sub
    End If
    If Txt��Ϣ�����Ŀ��.Enabled = True And Val(Txt��Ϣ�����Ŀ��.Text) > 10 ^ 8 Then
        MsgBox "��Ϣ�����Ŀ��̫��", vbInformation, gstrSysName
        Txt��Ϣ�����Ŀ��.SetFocus
        Exit Sub
    End If
    If StrIsValid(Txt����·��.Text, 50) = False Then
        Txt����·��.SetFocus
        Exit Sub
    End If
    If SaveCons = False Then Exit Sub
End Sub

Private Sub Chkʹ����־_Click()
    Cmd����.Enabled = True
    Txt������־�����Ŀ��.Enabled = Chkʹ����־.value = 1
End Sub

Private Sub Chk�Ƿ��¼���д���_Click()
    Cmd����.Enabled = True
    Txt������־�����Ŀ��.Enabled = Chk�Ƿ��¼���д���.value = 1
End Sub

Private Sub Cmd��ԭ_Click()
    Call InitCons
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ActiveControl Is Txt������־�����Ŀ�� Or ActiveControl Is Txt��Ϣ�����Ŀ�� Or ActiveControl Is Txt������־�����Ŀ�� Then
        If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call InitCons
End Sub

Private Sub InitCons()
    Dim ConThis As Control
    '--��ʼ�����ؼ���ֵ--
    
    For Each ConThis In Controls
        If Val(ConThis.Tag) <> 0 Then
            Set RecOption = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Zloption", Val(ConThis.Tag))
            With RecOption
                If Val(ConThis.Tag) = 6 Then
                    ConThis.Enabled = Not (.EOF)
                    CmdSelect.Enabled = Not (.EOF)
                End If
                
                Select Case TypeName(ConThis)
                Case "CheckBox"
                    If .EOF Then
                        ConThis.value = 0
                    Else
                        ConThis.value = IIf(IsNull(!Option_Value), 0, !Option_Value)
                    End If
                Case "TextBox"
                    If .EOF Then
                        ConThis.Text = ""
                    Else
                        ConThis.Text = IIf(IsNull(!Option_Value), "", !Option_Value)
                    End If
                End Select
            End With
        End If
    Next
    Txt������־�����Ŀ��.Enabled = Chkʹ����־.value = 1
    Txt������־�����Ŀ��.Enabled = Chk�Ƿ��¼���д���.value = 1
    
    Cmd����.Enabled = False
End Sub

Private Function SaveCons() As Boolean
    Dim ConThis As Control, StrValue As String
    '--������ؼ���ֵ--
    
    SaveCons = False
    On Error Resume Next
    err = 0
    
    gcnOracle.BeginTrans
    For Each ConThis In Controls
        If Val(ConThis.Tag) <> 0 Then
            Select Case TypeName(ConThis)
            Case "CheckBox"
                StrValue = ConThis.value
            Case "TextBox"
                StrValue = IIf(ConThis.Enabled = True, ConThis.Text, "")
            End Select
            gcnOracle.Execute "Update ZlOptions Set ����ֵ='" & StrValue & "' Where ������=" & Val(ConThis.Tag)
        End If
    Next
    
    If err <> 0 Then
        MsgBox "�������в���ʱ����������", vbInformation, gstrSysName
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    gcnOracle.CommitTrans
    MsgBox "���в����޸ĳɹ���", vbInformation, gstrSysName
    Cmd����.Enabled = False
    SaveCons = True
End Function

Private Sub SelLen(ByVal ConObj As TextBox)
    With ConObj
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLen_Change(Index As Integer)
    Cmd����.Enabled = True
    If Val(txtLen(0).Text) > Val(txtLen(1).Text) Then
        If Index = 0 Then
            txtLen(1).Text = txtLen(0).Text
        Else
            txtLen(0).Text = txtLen(1).Text
        End If
    End If
End Sub

Private Sub txtLen_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLen_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtLen(Index).Text) < udLen(Index).Min Then
        txtLen(Index).Text = udLen(Index).Min
    ElseIf Val(txtLen(Index).Text) > udLen(Index).Max Then
        txtLen(Index).Text = udLen(Index).Max
    End If
    If Val(txtLen(0).Text) > Val(txtLen(1).Text) Then
        If Index = 0 Then
            txtLen(1).Text = txtLen(0).Text
        Else
            txtLen(0).Text = txtLen(1).Text
        End If
    End If
    If Val(txtLen(1 - Index).Text) < udLen(1 - Index).Min Then
        txtLen(1 - Index).Text = udLen(1 - Index).Min
    ElseIf Val(txtLen(1 - Index).Text) > udLen(1 - Index).Max Then
        txtLen(1 - Index).Text = udLen(1 - Index).Max
    End If
End Sub

Private Sub Txt����·��_Change()
    Cmd����.Enabled = True
End Sub

Private Sub Txt����·��_GotFocus()
    SelAll Txt����·��
End Sub

Private Sub Txt������־�����Ŀ��_Change()
    Cmd����.Enabled = True
End Sub

Private Sub Txt������־�����Ŀ��_GotFocus()
    SelLen Txt������־�����Ŀ��
End Sub

Private Sub Txt��Ϣ�����Ŀ��_Change()
    Cmd����.Enabled = True
End Sub

Private Sub Txt��Ϣ�����Ŀ��_GotFocus()
    SelLen Txt��Ϣ�����Ŀ��
End Sub

Private Sub Txt������־�����Ŀ��_Change()
    Cmd����.Enabled = True
End Sub

Private Sub Txt������־�����Ŀ��_GotFocus()
    SelLen Txt������־�����Ŀ��
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

