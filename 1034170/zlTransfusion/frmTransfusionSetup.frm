VERSION 5.00
Begin VB.Form frmTransfusionSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmTransfusionSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6494.574
   ScaleMode       =   0  'User
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkTimeCall 
      Caption         =   "�����ƶ����й���"
      Height          =   255
      Left            =   255
      TabIndex        =   30
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox chk�ӵ����� 
      Caption         =   "�ӵ���ֱ�ӽ��봩��״̬"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   1560
      Width           =   3045
   End
   Begin VB.CheckBox chkAutoReady 
      Caption         =   "ͨ�����ҹ����ҵ����˺��Զ��ӵ�"
      Height          =   180
      Left            =   255
      TabIndex        =   3
      Top             =   1290
      Width           =   3045
   End
   Begin VB.Frame frmCardSet 
      Caption         =   "�豸����"
      Height          =   675
      Left            =   270
      TabIndex        =   28
      Top             =   2955
      Width           =   4470
      Begin VB.CommandButton cmdCardSet 
         Caption         =   "����(&P)"
         Height          =   350
         Left            =   2985
         TabIndex        =   29
         Top             =   210
         Width           =   1100
      End
   End
   Begin VB.Frame fra 
      Caption         =   "��ѡ�񱾹���վ��ʾ�ĵ�������"
      Height          =   660
      Left            =   270
      TabIndex        =   23
      Top             =   2250
      Width           =   4485
      Begin VB.CheckBox chkType 
         Caption         =   "����"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   27
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "��Һ"
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   26
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "ע��"
         Height          =   195
         Index           =   2
         Left            =   2355
         TabIndex        =   25
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Ƥ��"
         Height          =   195
         Index           =   3
         Left            =   3435
         TabIndex        =   24
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
   End
   Begin VB.TextBox txtƤ��Time 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3495
      MaxLength       =   4
      TabIndex        =   21
      ToolTipText     =   "�����ǰʱ��60����"
      Top             =   960
      Width           =   465
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3495
      TabIndex        =   20
      Top             =   1140
      Width           =   465
   End
   Begin VB.TextBox txt��ҺTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3510
      MaxLength       =   4
      TabIndex        =   18
      ToolTipText     =   "�����ǰʱ��60����"
      Top             =   690
      Width           =   465
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3510
      TabIndex        =   17
      Top             =   870
      Width           =   465
   End
   Begin VB.TextBox txt��ϵ�� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3675
      MaxLength       =   4
      TabIndex        =   15
      ToolTipText     =   "��ϵ��Ϊ10,15,20"
      Top             =   420
      Width           =   465
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3675
      TabIndex        =   14
      Top             =   600
      Width           =   465
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3540
      TabIndex        =   13
      Top             =   315
      Width           =   465
   End
   Begin VB.TextBox txt���� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3540
      MaxLength       =   4
      TabIndex        =   11
      ToolTipText     =   "������100��/��"
      Top             =   135
      Width           =   465
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   210
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3795
      Width           =   1100
   End
   Begin VB.CheckBox chkActLog 
      Caption         =   "���������˴���ִ�м�¼"
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   135
      Width           =   2280
   End
   Begin VB.CheckBox chkFinish 
      Caption         =   "����δ�շѲ������ִ��"
      Height          =   195
      Left            =   255
      TabIndex        =   1
      Top             =   420
      Width           =   2280
   End
   Begin VB.CheckBox chkƤ�� 
      Caption         =   "��дƤ�Խ��ʱ��֤���"
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   690
      Width           =   2280
   End
   Begin VB.TextBox txtRefresh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   705
      MaxLength       =   4
      TabIndex        =   5
      ToolTipText     =   "���ˢ�¼��Ϊ 30 �룬����Ϊ 0 ��ʾ���Զ�ˢ��"
      Top             =   960
      Width           =   465
   End
   Begin VB.Frame fraLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   705
      TabIndex        =   8
      Top             =   1140
      Width           =   465
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   6
      Top             =   3795
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   7
      Top             =   3795
      Width           =   1100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ƥ����ǰ      ��������"
      Height          =   180
      Left            =   2745
      TabIndex        =   22
      Top             =   960
      Width           =   1980
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Һ��ǰ      ��������"
      Height          =   180
      Left            =   2760
      TabIndex        =   19
      Top             =   690
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ�ϵ�ϵ��      "
      Height          =   180
      Left            =   2760
      TabIndex        =   16
      ToolTipText     =   "��ϵ��Ϊ10,15,20"
      Top             =   420
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ�ϵ���      ��/��"
      Height          =   180
      Left            =   2790
      TabIndex        =   12
      ToolTipText     =   "������100��/��"
      Top             =   135
      Width           =   1710
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ      ���Զ�ˢ���嵥"
      Height          =   180
      Left            =   480
      TabIndex        =   9
      ToolTipText     =   "���ˢ�¼��Ϊ 30 �룬����Ϊ 0 ��ʾ���Զ�ˢ��"
      Top             =   975
      Width           =   1980
   End
End
Attribute VB_Name = "frmTransfusionSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mstrPrivs As String
Public mlng����ID As Long 'IN:��ǰִ�п���ID
Public mblnOk As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCardSet_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strPar As String, i As Long
    Dim strType As String
    Dim blnModify As Boolean
    'ִ�м䷶Χ
    blnModify = False
    If InStr(mstrPrivs, "��������") > 0 Then blnModify = True
    
    
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
    Call zlDatabase.SetPara("ҽ��ˢ�¼��", Val(txtRefresh.Text), glngSys, 1264, blnModify)
    
    '�Ƿ��������ִ�м�¼
    Call zlDatabase.SetPara("����ִ�м�¼", chkActLog.Value, glngSys, 1264, blnModify)
    '�Ƿ��������δ�շѲ��˵���Ŀ
    Call zlDatabase.SetPara("δ�շ����", chkFinish.Value, glngSys, 1264, blnModify)
    '��дƤ�Խ��ʱ��֤���
    Call zlDatabase.SetPara("Ƥ����֤���", chkƤ��.Value, glngSys, 1264, blnModify)
    '�ӵ���ֱ�ӽ��봩��״̬
    Call zlDatabase.SetPara("�ӵ�ֱ�Ӵ���", chk�ӵ�����.Value, glngSys, 1264)
    '�ƶ�����
    Call zlDatabase.SetPara("�ƶ�����", chkTimeCall.Value, glngSys, 1264)
    
    If Val(txtƤ��Time.Text) < 0 Or Val(txtƤ��Time.Text) > 60 Then txtƤ��Time.Text = 0
    Call zlDatabase.SetPara("Ƥ��������ǰʱ��", Val(txtƤ��Time.Text), glngSys, 1264, blnModify)
    
    
    If Val(txt����.Text) < 10 And Val(txt����.Text) > 100 Then txt����.Text = 40
    Call zlDatabase.SetPara("Ĭ�ϵ���", Val(txt����.Text), glngSys, 1264, blnModify)
    
    If InStr(",10,15,20,", "," & Val(txt��ϵ��.Text) & ",") <= 0 Then txt��ϵ��.Text = 20
    Call zlDatabase.SetPara("Ĭ�ϵ�ϵ��", Val(txt��ϵ��.Text), glngSys, 1264, blnModify)
    
    If Val(txt��ҺTime.Text) < 0 Or Val(txt��ҺTime.Text) > 60 Then txt��ҺTime.Text = 3
    Call zlDatabase.SetPara("��Һ������ǰʱ��", Val(txt��ҺTime.Text), glngSys, 1264, blnModify)
    
    '2008-11-12
    strType = ""
    For i = 0 To chkType.Count - 1
        strType = strType & "," & chkType(i).Value
    Next
    Call zlDatabase.SetPara("��ʾ��������", Mid(strType, 2), glngSys, 1264, blnModify)
    
    '2012-05-14 10.30 sp? ���
    Call zlDatabase.SetPara("������Һ�Զ��ӵ�", chkAutoReady.Value, glngSys, 1264, blnModify)
    
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim strType As String, i As Integer
    Dim intType As Integer '������������
    Dim blnModify As Boolean
    
    mblnOk = False
    blnModify = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    
    cmdCardSet.Enabled = blnModify
    
    '�޸�:���˺�  �޸��˲�����������    ����:2008-06-12 10:58:11,��Ҫ������ Array(..),InStr(mstrPrivs, ";��������;")>0
    txtRefresh.Text = Val(zlDatabase.GetPara("ҽ��ˢ�¼��", glngSys, 1264, "", Array(txtRefresh), blnModify))
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
      '�Ƿ��������ִ�м�¼
    chkActLog.Value = Val(zlDatabase.GetPara("����ִ�м�¼", glngSys, 1264, "", Array(chkActLog), blnModify))
    
    '�Ƿ��������δ�շѲ��˵���Ŀ
    chkFinish.Value = Val(zlDatabase.GetPara("δ�շ����", glngSys, 1264, "", Array(chkFinish), blnModify))
    
    '��дƤ�Խ��ʱ��֤���
    chkƤ��.Value = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, 1264, "", Array(chkƤ��), blnModify))
    
    '�ӵ���ֱ�ӽ��봩��״̬
    chk�ӵ�����.Value = Val(zlDatabase.GetPara("�ӵ�ֱ�Ӵ���", glngSys, 1264, ""))
    
    '�ƶ���ʱ����
    chkTimeCall.Value = Val(zlDatabase.GetPara("�ƶ�����", glngSys, 1264))
        
    txtƤ��Time.Text = Val(zlDatabase.GetPara("Ƥ��������ǰʱ��", glngSys, 1264, "", Array(txtƤ��Time), blnModify))
    If Val(txtƤ��Time.Text) < 0 Or Val(txtƤ��Time.Text) > 60 Then txtƤ��Time.Text = 0
    
    txt����.Text = Val(zlDatabase.GetPara("Ĭ�ϵ���", glngSys, 1264, "", Array(txt����), blnModify))
    If Val(txt����.Text) < 10 Or Val(txt����.Text) > 100 Then txt����.Text = 40
        
    txt��ϵ��.Text = Val(zlDatabase.GetPara("Ĭ�ϵ�ϵ��", glngSys, 1264, "", Array(txt��ϵ��), blnModify))
    If InStr(",10,15,20,", "," & Val(txt��ϵ��.Text) & ",") <= 0 Then txt��ϵ��.Text = 20

    txt��ҺTime.Text = Val(zlDatabase.GetPara("��Һ������ǰʱ��", glngSys, 1264, "", Array(txt��ҺTime), blnModify))
    If Val(txt��ҺTime.Text) < 0 Or Val(txt��ҺTime.Text) > 60 Then txt��ҺTime.Text = 3
        
    '2008-11-12
    strType = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1", Array(Me.chkType(0), Me.chkType(1), Me.chkType(2), Me.chkType(3)), blnModify, intType)
    'strType = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1")
    For i = 0 To chkType.Count - 1
        chkType(i).Value = Val(Split(strType, ",")(i))
    Next
    '2012-05-14
    chkAutoReady.Value = Val(zlDatabase.GetPara("������Һ�Զ��ӵ�", glngSys, 1264, "", Array(chkAutoReady), blnModify))
    
    '�޸�:���˺�  �޸��˲�����������    ����:2008-06-12 10:58:11,��Ҫ�������´���
'    'Ȩ�����õ�Ȩ�޿���
'    If InStr(mstrPrivs, "��������") = 0 And intType = 15 Then
'        chkActLog.Enabled = False
'        chkFinish.Enabled = False
'        chkƤ��.Enabled = False
'
'        txtRefresh.Enabled = False
'        txt����.Enabled = False
'        txt��ϵ��.Enabled = False
'        txt��ҺTime.Enabled = False
'        txtƤ��Time.Enabled = False
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    mstrPrivs = ""
End Sub

Private Sub txtRefresh_GotFocus()
    Call zlControl.TxtSelAll(txtRefresh)
End Sub

Private Sub txtRefresh_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRefresh_Validate(Cancel As Boolean)
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt����_Validate(Cancel As Boolean)
    If Val(txt����.Text) < 10 Or Val(txt����.Text) > 100 Then txt����.Text = 40
End Sub
Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt��ϵ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt��ϵ��_Validate(Cancel As Boolean)
    If InStr(",10,15,20,", "," & Val(txt��ϵ��.Text) & ",") <= 0 Then txt��ϵ��.Text = 20
End Sub
Private Sub txt��ϵ��_GotFocus()
    Call zlControl.TxtSelAll(txt��ϵ��)
End Sub

Private Sub txt��ҺTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt��ҺTime_Validate(Cancel As Boolean)
    If Val(txt��ҺTime.Text) < 0 Or Val(txt��ҺTime.Text) > 60 Then txt��ҺTime.Text = 3
End Sub
Private Sub txt��ҺTime_GotFocus()
    Call zlControl.TxtSelAll(txt��ҺTime)
End Sub

Private Sub txtƤ��Time_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txtƤ��Time_Validate(Cancel As Boolean)
    If Val(txtƤ��Time.Text) < 0 Or Val(txtƤ��Time.Text) > 60 Then txtƤ��Time.Text = 0
End Sub
Private Sub txtƤ��Time_GotFocus()
    Call zlControl.TxtSelAll(txtƤ��Time)
End Sub
