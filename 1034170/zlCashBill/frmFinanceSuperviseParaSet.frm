VERSION 5.00
Begin VB.Form frmFinanceSuperviseParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinanceSuperviseParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraPrintModeDraw 
      Caption         =   "���ý����õ���ӡ��ʽ"
      Height          =   945
      Left            =   150
      TabIndex        =   10
      Top             =   3180
      Width           =   8760
      Begin VB.OptionButton optPrintModeDraw 
         Caption         =   "���ú󲻴�ӡ(&1)"
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   420
         Width           =   1935
      End
      Begin VB.OptionButton optPrintModeDraw 
         Caption         =   "���ú��Զ���ӡ(&2)"
         Height          =   300
         Index           =   1
         Left            =   2070
         TabIndex        =   13
         Top             =   405
         Value           =   -1  'True
         Width           =   2190
      End
      Begin VB.OptionButton optPrintModeDraw 
         Caption         =   "���ú�ѡ���Ƿ��ӡ(&3)"
         Height          =   300
         Index           =   2
         Left            =   4275
         TabIndex        =   12
         Top             =   405
         Width           =   2655
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "��ӡ����(&S)"
         Height          =   375
         Index           =   1
         Left            =   7125
         TabIndex        =   11
         Top             =   360
         Width           =   1530
      End
   End
   Begin VB.Frame fraPrintModeSJ 
      Caption         =   "�տ��վݴ�ӡ��ʽ"
      Height          =   930
      Left            =   120
      TabIndex        =   5
      Top             =   2085
      Width           =   8790
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "��ӡ����(&S)"
         Height          =   375
         Index           =   0
         Left            =   7125
         TabIndex        =   9
         Top             =   330
         Width           =   1530
      End
      Begin VB.OptionButton optPrintModeSJ 
         Caption         =   "�տ��ѡ���Ƿ��ӡ(&3)"
         Height          =   300
         Index           =   2
         Left            =   4275
         TabIndex        =   8
         Top             =   420
         Width           =   2655
      End
      Begin VB.OptionButton optPrintModeSJ 
         Caption         =   "�տ���Զ���ӡ(&2)"
         Height          =   300
         Index           =   1
         Left            =   2070
         TabIndex        =   7
         Top             =   405
         Value           =   -1  'True
         Width           =   2190
      End
      Begin VB.OptionButton optPrintModeSJ 
         Caption         =   "�տ�󲻴�ӡ(&1)"
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   420
         Width           =   1935
      End
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   -855
      TabIndex        =   3
      Top             =   975
      Width           =   9930
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   1
      Left            =   -90
      TabIndex        =   2
      Top             =   4485
      Width           =   9525
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7755
      TabIndex        =   1
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6420
      TabIndex        =   0
      Top             =   4785
      Width           =   1100
   End
   Begin VB.TextBox txtDrawMoney 
      Height          =   330
      Left            =   2235
      TabIndex        =   16
      Top             =   1500
      Width           =   1995
   End
   Begin VB.Label lblDrawMoney 
      AutoSize        =   -1  'True
      Caption         =   "���ý�ȱʡ���ý��                     Ԫ"
      Height          =   210
      Left            =   225
      TabIndex        =   15
      Top             =   1545
      Width           =   4305
   End
   Begin VB.Image imgNotes 
      Height          =   720
      Left            =   195
      Picture         =   "frmFinanceSuperviseParaSet.frx":06EA
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblTittle 
      Caption         =   $"frmFinanceSuperviseParaSet.frx":15B4
      Height          =   645
      Left            =   1080
      TabIndex        =   4
      Top             =   285
      Width           =   7695
   End
End
Attribute VB_Name = "frmFinanceSuperviseParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As String, mstrPrivs As String, mblnOk As Boolean
Public Function ShowMe(ByVal frmMain As Form, _
    ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ز��������
    '����:�������óɹ�������true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 14:33:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOk = False: mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    ShowMe = mblnOk
End Function
Private Sub LoadPara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���
    '����:���˺�
    '����:2013-09-12 15:26:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    i = Val(zlDatabase.GetPara("�տ��վݴ�ӡ��ʽ", glngSys, mlngModule, 0, Array(fraPrintModeSJ, optPrintModeSJ(0), optPrintModeSJ(1), optPrintModeSJ(2)), InStr(1, mstrPrivs, ";��������;") > 0))
    If i > 2 Or i < 0 Then
        optPrintModeSJ(0).Value = True
    Else
        optPrintModeSJ(i).Value = True
    End If
    i = Val(zlDatabase.GetPara("���ý����õ���ӡ��ʽ", glngSys, mlngModule, 0, Array(fraPrintModeDraw, optPrintModeDraw(0), optPrintModeDraw(1), optPrintModeDraw(2)), InStr(1, mstrPrivs, ";��������;") > 0))
    If i > 2 Or i < 0 Then
        optPrintModeDraw(0).Value = True
    Else
        optPrintModeDraw(i).Value = True
    End If
    txtDrawMoney.Text = Val(zlDatabase.GetPara("ȱʡ���ñ��ý��", glngSys, mlngModule, 1000, Array(txtDrawMoney, lblDrawMoney), InStr(1, mstrPrivs, ";��������;") > 0))
End Sub
Private Sub SavePara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2013-09-12 15:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call zlDatabase.SetPara("�տ��վݴ�ӡ��ʽ", IIf(optPrintModeSJ(0).Value, 0, IIf(optPrintModeSJ(1).Value, 1, 2)), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    Call zlDatabase.SetPara("���ý����õ���ӡ��ʽ", IIf(optPrintModeDraw(0).Value, 0, IIf(optPrintModeDraw(1).Value, 1, 2)), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    Call zlDatabase.SetPara("ȱʡ���ñ��ý��", Val(txtDrawMoney.Text), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Call SavePara
    Unload Me
    mblnOk = True
End Sub
Private Sub cmdPrintSetup_Click(Index As Integer)
    If Index = 0 Then
        Call ReportPrintSet(gcnOracle, glngSys, "zl" & Int(glngSys / 100) & "_BILL_1500", Me)
    Else
        Call ReportPrintSet(gcnOracle, glngSys, "zl" & Int(glngSys / 100) & "_BILL_1500_1", Me)
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Call LoadPara
End Sub

