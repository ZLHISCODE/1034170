VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�ظ��ϴ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ϴ�ʱ������"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frm�ظ��ϴ�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   5
      Top             =   810
      Width           =   1100
   End
   Begin VB.Frame fraScope 
      Caption         =   "ʱ�䷶Χ"
      Height          =   1455
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3405
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1830
         TabIndex        =   1
         Top             =   870
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   23789571
         CurrentDate     =   36279
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1830
         TabIndex        =   2
         Top             =   390
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   23789571
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.Label lblTimeStop 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��(&E)"
         Height          =   180
         Left            =   780
         TabIndex        =   4
         Top             =   930
         Width           =   990
      End
      Begin VB.Label lblTimeStart 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��(&B)"
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   450
         Width           =   990
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   150
         Picture         =   "frm�ظ��ϴ�����.frx":000C
         Top             =   420
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   6
      Top             =   270
      Width           =   1100
   End
   Begin VB.Label lbl��ʾ 
      Caption         =   "ע�⣺������ֻ��ҽ�����������ܵ��ƻ��������ݻָ�ʱʹ�á�"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1860
      Width           =   3585
   End
End
Attribute VB_Name = "frm�ظ��ϴ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mdatBegin As Date, mdatEnd As Date

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "��ʼʱ����ڽ���ʱ���ˡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("���ʹ�ò����������ƻ�ҽ�����ĵ����ݡ�" & vbCrLf & "�Ƿ�ȷ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function GetTimeScope(datBegin As Date, datEnd As Date, ByVal datMax As Date) As Boolean
                
    dtpBegin.Value = datBegin
    dtpEnd.Value = datEnd
    
    '�ϸ���������
    dtpBegin.MaxDate = datMax
    dtpEnd.MaxDate = dtpBegin.MaxDate
    
    frm�ظ��ϴ�����.Show vbModal, frm�ϴ�����
    
    GetTimeScope = mblnOK
    If mblnOK = True Then
        datBegin = mdatBegin
        datEnd = mdatEnd
    End If
End Function


