VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCISBorrowPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmCISBorrowPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   2
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   1
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4935
      TabIndex        =   0
      Top             =   6000
      Width           =   1100
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   0
      Left            =   255
      ScaleHeight     =   5415
      ScaleWidth      =   5880
      TabIndex        =   3
      Top             =   330
      Width           =   5880
      Begin VB.CheckBox chkBorrowAccount 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2685
         TabIndex        =   18
         Top             =   3495
         Width           =   660
      End
      Begin VB.CheckBox chkBorrowReason 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2685
         TabIndex        =   16
         Top             =   3105
         Width           =   660
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   2460
         TabIndex        =   14
         Top             =   2670
         Width           =   1920
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   2460
         TabIndex        =   12
         Top             =   2265
         Width           =   1920
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   3
         Left            =   930
         TabIndex        =   9
         Top             =   1920
         Width           =   4815
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   870
         Width           =   1920
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   930
         TabIndex        =   5
         Top             =   165
         Width           =   4815
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������¼�����ԭ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   825
         TabIndex        =   19
         Top             =   3540
         Width           =   1800
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����¼�������������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   825
         TabIndex        =   17
         Top             =   3150
         Width           =   1800
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ�������Ϊ                      ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   825
         TabIndex        =   15
         Top             =   2730
         Width           =   3780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ������ȱʡΪ                      ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   825
         TabIndex        =   13
         Top             =   2325
         Width           =   3780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   16
         Left            =   180
         TabIndex        =   11
         Top             =   1905
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   195
         TabIndex        =   10
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ���Ӳ������������ȱʡʱ�䷶Χ��"
         Height          =   405
         Left            =   1035
         TabIndex        =   8
         Top             =   555
         Width           =   4065
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmCISBorrowPara.frx":000C
         Top             =   390
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ��Χ(&1)"
         Height          =   180
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Top             =   930
         Width           =   990
      End
   End
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   5790
      Left            =   105
      TabIndex        =   4
      Top             =   30
      Width           =   5970
      _Version        =   589884
      _ExtentX        =   10530
      _ExtentY        =   10213
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmCISBorrowPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mblnOK As Boolean
Private mfrmMain As Object
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mstrPrivs As String
Private mblnBorrowAccount As Boolean '��������¼�����ԭ��

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    If ExecuteCommand("��ȡ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim blnAllowModify As Boolean

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "��ʼ����"
            With tbc
                With .PaintManager
                    .Appearance = xtpTabAppearancePropertyPage2003
                    .BoldSelected = True
                    .COLOR = xtpTabColorDefault
                    .ColorSet.ButtonSelected = COLOR.��ɫ
                    .ShowIcons = True
                End With
                
                .InsertItem 0, "���� ", picPane(0).Hwnd, 0
                .Item(0).Selected = True
            End With
            

            cbo(0).Clear
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "������"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰһ��"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰһ��"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰһ��"
            cbo(0).AddItem "ǰ����"
            
'            lbl(3).ForeColor = COLOR.����ģ��ɫ
'            txt(7).ForeColor = COLOR.����ģ��ɫ
'
'            lbl(0).ForeColor = COLOR.����ģ��ɫ
'            txt(0).ForeColor = COLOR.����ģ��ɫ
            
        '--------------------------------------------------------------------------------------------------------------
        Case "�ؼ�״̬"
            
'            blnAllowModify = IsPrivs(mstrPrivs, "��������")
'            lbl(3).Enabled = blnAllowModify
'            txt(7).Enabled = blnAllowModify
'            lbl(0).Enabled = blnAllowModify
'            txt(0).Enabled = blnAllowModify
        '--------------------------------------------------------------------------------------------------------------
        Case "��ȡ����"
            
            On Error Resume Next
            cbo(0).Text = zlDatabase.GetPara("�Ǽ�ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(0)), IsPrivs(mstrPrivs, "��������"))
            On Error GoTo errHand
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            txt(7).Text = Val(zlDatabase.GetPara("������������", ParamInfo.ϵͳ��, mfrmMain.ģ���, "7", Array(txt(7)), IsPrivs(mstrPrivs, "��������")))
            txt(0).Text = Val(zlDatabase.GetPara("���������", ParamInfo.ϵͳ��, mfrmMain.ģ���, "30", Array(txt(0)), IsPrivs(mstrPrivs, "��������")))
            chkBorrowReason.Value = zlDatabase.GetPara("����¼�����ԭ��", ParamInfo.ϵͳ��, mfrmMain.ģ���, "0", chkBorrowReason, IsPrivs(mstrPrivs, "��������"))
            chkBorrowAccount.Value = zlDatabase.GetPara("��������¼�����ԭ��", ParamInfo.ϵͳ��, mfrmMain.ģ���, "0", chkBorrowAccount, IsPrivs(mstrPrivs, "��������"))
            
        '--------------------------------------------------------------------------------------------------------------
        Case "У������"
            
            If Val(txt(0).Text) < Val(txt(7).Text) Then
                ShowSimpleMsg "���ĵ�����޲���С�ڲ������ĵ�ȱʡ����!"
                Exit Function
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case "��������"
            
            Call SetPara("�Ǽ�ȱʡ��Χ", cbo(0).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("������������", Val(txt(7).Text), mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("���������", Val(txt(0).Text), mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("����¼�����ԭ��", chkBorrowReason.Value, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("��������¼�����ԭ��", chkBorrowAccount.Value, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
        End Select
    Next

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    cmdOK.Tag = IIf(blnData, "Changed", "")
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = (cmdOK.Tag = "Changed")
End Property

'######################################################################################################################

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkBorrowAccount_Click()
    DataChanged = True
End Sub

Private Sub chkBorrowReason_Click()
    DataChanged = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    
    If DataChanged Then
        If ExecuteCommand("У������") = False Then Exit Sub
        
        If ExecuteCommand("��������") Then
            
            DataChanged = False
            
            mblnOK = True
        Else
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("�������޸ĵĲ������뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
    End If
    
    Set mclsVsf = Nothing
End Sub

Private Sub txt_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 5
        zlCommFun.OpenIme True
    End Select
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 5
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


