VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSublimeStationSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9585
   Icon            =   "frmSublimeStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt��Ժ���� 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   6555
      MaxLength       =   2
      TabIndex        =   26
      Text            =   "3"
      Top             =   2850
      Width           =   300
   End
   Begin VB.Frame fraFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   3
      Left            =   6570
      TabIndex        =   48
      Top             =   3030
      Width           =   300
   End
   Begin VB.Frame fraSplit 
      Height          =   135
      Left            =   -30
      TabIndex        =   47
      Top             =   4560
      Width           =   9840
   End
   Begin VB.PictureBox picControl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   1770
      ScaleHeight     =   1785
      ScaleWidth      =   2295
      TabIndex        =   38
      Top             =   2670
      Visible         =   0   'False
      Width           =   2295
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1497
         Width           =   200
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   60
         Picture         =   "frmSublimeStationSetup.frx":000C
         ScaleHeight     =   1350
         ScaleWidth      =   2160
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   90
         Width           =   2160
         Begin VB.Shape shpBorder 
            BorderColor     =   &H00C56A31&
            FillColor       =   &H00FF8080&
            Height          =   270
            Left            =   1890
            Top             =   1080
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape shpValue 
            BorderColor     =   &H00C56A31&
            FillColor       =   &H00FF8080&
            Height          =   270
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   270
         End
      End
      Begin VB.Label lblColor 
         Caption         =   "&HFFFFFF"
         Height          =   195
         Left            =   405
         TabIndex        =   42
         Top             =   1500
         UseMnemonic     =   0   'False
         Width           =   1365
      End
   End
   Begin VB.Frame fraMedRec 
      Caption         =   "������鷴������"
      Height          =   600
      Left            =   4995
      TabIndex        =   23
      Top             =   2115
      Width           =   4485
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1025
         TabIndex        =   44
         Top             =   420
         Width           =   300
      End
      Begin VB.TextBox txtMedRec 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   1040
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "1"
         Top             =   240
         Width           =   300
      End
      Begin VB.Label lblMedRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʾ    ���ڵĲ�����鷴����"
         Height          =   180
         Left            =   645
         TabIndex        =   45
         Top             =   255
         Width           =   2520
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " ����ȼ���ɫ"
      Height          =   1530
      Left            =   120
      TabIndex        =   14
      Top             =   2820
      Width           =   4755
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   3
         Left            =   3840
         Picture         =   "frmSublimeStationSetup.frx":0782
         Stretch         =   -1  'True
         Top             =   900
         Width           =   345
      End
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   2
         Left            =   1770
         Picture         =   "frmSublimeStationSetup.frx":0E84
         Stretch         =   -1  'True
         Top             =   900
         Width           =   345
      End
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   3840
         Picture         =   "frmSublimeStationSetup.frx":1586
         Stretch         =   -1  'True
         Top             =   420
         Width           =   345
      End
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1770
         Picture         =   "frmSublimeStationSetup.frx":1C88
         Stretch         =   -1  'True
         Top             =   420
         Width           =   345
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2610
         TabIndex        =   30
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   540
         TabIndex        =   29
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2610
         TabIndex        =   28
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ؼ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   540
         TabIndex        =   27
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7185
      TabIndex        =   36
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8400
      TabIndex        =   37
      Top             =   4785
      Width           =   1100
   End
   Begin VB.Frame fraAdvice 
      Caption         =   " ҽ���������� "
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4755
      Begin VB.CheckBox chkWarn 
         Caption         =   "�걾������δ����"
         Enabled         =   0   'False
         Height          =   555
         Index           =   10
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   11
         Left            =   1395
         TabIndex        =   51
         Top             =   1635
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Ѫ������"
         Height          =   195
         Index           =   12
         Left            =   2520
         TabIndex        =   50
         Top             =   1635
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "ȡѪ֪ͨ"
         Height          =   195
         Index           =   9
         Left            =   3670
         TabIndex        =   49
         Top             =   1380
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RISԤԼ׼��"
         Height          =   195
         Index           =   8
         Left            =   2355
         TabIndex        =   11
         Top             =   1380
         Width           =   1320
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RISԤԼ"
         Height          =   195
         Index           =   7
         Left            =   1395
         TabIndex        =   10
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "����������ʾ"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   2160
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   1860
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Σ��ֵ"
         Height          =   195
         Index           =   4
         Left            =   1395
         TabIndex        =   7
         Top             =   1110
         Width           =   870
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Һ�ܾ�"
         Height          =   195
         Index           =   5
         Left            =   2355
         TabIndex        =   8
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��������"
         Height          =   195
         Index           =   6
         Left            =   3480
         TabIndex        =   9
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox ChkCollate 
         Caption         =   "ҽ��������Զ���λ������ҽ��ҳ��"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   1890
         Width           =   3900
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "����"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   6
         Top             =   855
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�¿�"
         Height          =   195
         Index           =   0
         Left            =   1395
         TabIndex        =   3
         Top             =   855
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��ͣ"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   4
         Top             =   855
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�·�"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   5
         Top             =   855
         Width           =   660
      End
      Begin VB.TextBox txtNotifyAdvice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "10"
         Top             =   315
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdvice 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   32
         Top             =   495
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdviceDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   31
         Top             =   765
         Width           =   300
      End
      Begin VB.TextBox txtNotifyAdviceDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "1"
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chkNotifyAdvice 
         Caption         =   "ÿ    �����Զ�ˢ��ҽ�����������е�����"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   570
         TabIndex        =   35
         Top             =   855
         Width           =   810
      End
      Begin VB.Label lblNotifyAdviceDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ���ڴ����ҽ��������ʾ����������"
         Height          =   180
         Left            =   570
         TabIndex        =   34
         Top             =   600
         Width           =   3420
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Left            =   360
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   41
      Top             =   3165
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   " ���Ի��������� "
      Height          =   690
      Left            =   4995
      TabIndex        =   15
      Top             =   75
      Width           =   4485
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1320
         TabIndex        =   43
         Top             =   495
         Width           =   300
      End
      Begin VB.TextBox txt������� 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   1305
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "3"
         Top             =   315
         Width           =   300
      End
      Begin VB.CheckBox chkPatientFilter 
         Caption         =   "��ȡ���    ���ڵ�סԺ����"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   315
         Width           =   3900
      End
   End
   Begin VB.Frame FraCard 
      Caption         =   " ��Ƭ��ǩ���� "
      Height          =   1095
      Left            =   4995
      TabIndex        =   18
      Top             =   900
      Width           =   4485
      Begin VB.OptionButton optNewCard 
         Caption         =   "��λ��"
         Height          =   180
         Index           =   0
         Left            =   1290
         TabIndex        =   21
         Top             =   675
         Width           =   945
      End
      Begin VB.OptionButton optNewCard 
         Caption         =   "��λ���Ʊ��+��λ��"
         Height          =   180
         Index           =   1
         Left            =   2400
         TabIndex        =   22
         Top             =   675
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CheckBox chkAmount 
         Caption         =   "��Ƭ����������������"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label lblNewCard 
         AutoSize        =   -1  'True
         Caption         =   "��Ƭ����"
         Height          =   180
         Left            =   300
         TabIndex        =   20
         Top             =   660
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList img24 
      Left            =   225
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CheckBox chkNewPati 
      Caption         =   "������б���ʾ    ���ڵǼǵ�סԺ����"
      Height          =   195
      Left            =   5010
      TabIndex        =   25
      Top             =   2850
      Width           =   3900
   End
End
Attribute VB_Name = "frmSublimeStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarColor As OLE_COLOR
Public mstrPrivs As String
Private mlngModual As Long

Private Const ALTERNATE = 1
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" _
    (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" _
    (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Polyline Lib "gdi32" _
    (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'�趨һ�����岶����꣬���������������Ϣ�������ô���
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'ȡ����겶��
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private mlngColor As Long
Private mintIndex As Long
Private mobjFileSys As New FileSystemObject

Public Sub ShowMe()
    '���°�סԺ��ʿ����վ���ã���ʾ��ע��ť
    mintIndex = 0
    Me.Show vbModal
End Sub

Private Sub chkNewPati_Click()
    On Error Resume Next
    If chkNewPati.Value = 1 Then
        txt��Ժ����.Enabled = True
        txt��Ժ����.SetFocus
    Else
        txt��Ժ����.Enabled = False
        txt��Ժ����.Text = ""
    End If
End Sub

Private Sub chkNotifyAdvice_Click()
    txtNotifyAdvice.Enabled = chkNotifyAdvice.Value = 1
    If Visible And txtNotifyAdvice.Enabled Then txtNotifyAdvice.SetFocus
End Sub

Private Sub chkPatientFilter_Click()
    On Error Resume Next
    If chkPatientFilter.Value = 1 Then
        txt�������.Enabled = True
        txt�������.SetFocus
    Else
        txt�������.Enabled = False
        txt�������.Text = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    Dim strTmp As String
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If chkNotifyAdvice.Value = 1 And Val(txtNotifyAdvice.Text) = 0 Then
        If txtNotifyAdvice.Text = "" Then
            MsgBox "������ҽ�����ѵ��Զ�ˢ�¼����", vbInformation, gstrSysName
        Else
            MsgBox "ҽ�����ѵ��Զ�ˢ�¼������ӦΪ1���ӡ�", vbInformation, gstrSysName
        End If
        txtNotifyAdvice.SetFocus: Exit Sub
    End If
    If Val(txtNotifyAdviceDay.Text) = 0 Then
        If txtNotifyAdviceDay.Text = "" Then
            MsgBox "������Ҫ���ѵ�ҽ��������", vbInformation, gstrSysName
        Else
            MsgBox "Ҫ���ѵ�ҽ����������ӦΪ1�졣", vbInformation, gstrSysName
        End If
        txtNotifyAdviceDay.SetFocus: Exit Sub
    End If
    If chkPatientFilter.Value = 1 Then
        If Trim(txt�������.Text) = "" Then
            MsgBox "�������������������", vbInformation, gstrSysName
            txt�������.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt�������.Text) Then
            MsgBox "��������к��зǷ��ַ�����ֻ���������֣�", vbInformation, gstrSysName
            txt�������.SetFocus
            Exit Sub
        End If
        If Val(txt�������.Text) <= 0 Then
            MsgBox "���������������㣡", vbInformation, gstrSysName
            txt�������.SetFocus
            Exit Sub
        End If
    End If
    
    '73646
    If txtMedRec.Text = "" Then
        MsgBox "�����ò�����鷴�����ѵ�������", vbInformation, gstrSysName
        txtMedRec.SetFocus: Exit Sub
    End If
    
    If chkNewPati.Value = 1 Then
        If Trim(txt��Ժ����.Text) = "" Then
            MsgBox "������������ʾ����Ժ�Ǽ�������", vbInformation, gstrSysName
            txt��Ժ����.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt��Ժ����.Text) Then
            MsgBox "�������ʾ����Ժ�Ǽ������к��зǷ��ַ�����ֻ���������֣�", vbInformation, gstrSysName
            txt��Ժ����.SetFocus
            Exit Sub
        End If
        If Val(txt��Ժ����.Text) <= 0 Then
            MsgBox "�������ʾ����Ժ�Ǽ�������������㣡", vbInformation, gstrSysName
            txt��Ժ����.SetFocus
            Exit Sub
        End If
    End If
    
    '�Զ�ˢ��ҽ������
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("�Զ�ˢ��ҽ�����", IIf(chkNotifyAdvice.Value = 1, Val(txtNotifyAdvice.Text), ""), glngSys, pסԺ��ʿվ, blnSetup)
    Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", Val(txtNotifyAdviceDay.Text), glngSys, pסԺ��ʿվ, blnSetup)
    strTmp = ""
    For i = 0 To chkWarn.UBound
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", strTmp, glngSys, pסԺ��ʿվ, blnSetup)
    
    '�����������
    If chkPatientFilter.Value = 1 Then
        Call zlDatabase.SetPara("�������", txt�������.Text, glngSys, 1265, blnSetup)
    Else
        Call zlDatabase.SetPara("�������", "0", glngSys, 1265, blnSetup)
    End If
    '������Ժ���� 111016
    If chkNewPati.Value = 1 Then
        Call zlDatabase.SetPara("��Ժ����", txt��Ժ����.Text, glngSys, 1265, blnSetup)
    Else
        Call zlDatabase.SetPara("��Ժ����", "0", glngSys, 1265, blnSetup)
    End If
    '���滤��ȼ�����ɫ
    Call zlDatabase.SetPara("�ؼ�������ɫ", img����ȼ�(0).Tag, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("һ��������ɫ", img����ȼ�(1).Tag, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("����������ɫ", img����ȼ�(2).Tag, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("����������ɫ", img����ȼ�(3).Tag, glngSys, 1265, blnSetup)
    '--56960:������,2013-01-17,��Ӳ���"��Ƭ���������"
    Call zlDatabase.SetPara("��Ƭ���������", chkAmount.Value, glngSys, 1265, blnSetup)
    '54370:������,2013-05-02,��Ӳ���"ҽ��������Զ���λ��ҽ��ҳ��"
    Call zlDatabase.SetPara("ҽ��������Զ���λ��ҽ��ҳ��", ChkCollate.Value, glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("������鷴������", txtMedRec.Text, glngSys, pסԺ��ʿվ, blnSetup)
    '92852:������,2016-01-21,��λ��������
    Call zlDatabase.SetPara("��λ��Ƭ����ʽ", IIf(optNewCard(0).Value = True, 0, 1), glngSys, 1265, blnSetup)
    Call zlDatabase.SetPara("����������ʾ", chkSound.Value, glngSys, pסԺ��ʿվ, blnSetup)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyCode = vbKeyEscape Then
        ReleaseCapture
        picControl.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strPar As String
    Dim intType As Integer
    
    gblnOK = False
    mlngModual = pסԺ��ʿվ
    
    chkWarn(9).Visible = gblnѪ��ϵͳ
    chkWarn(11).Visible = gblnѪ��ϵͳ
    '�Զ�ˢ��ҽ������
    strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, mlngModual, , Array(chkNotifyAdvice), InStr(mstrPrivs, "��������") > 0, intType)
    If Val(strPar) > 0 Then
        chkNotifyAdvice.Value = 1: txtNotifyAdvice.Text = Val(strPar)
    End If
    'ǰ���¼��л��Զ����ã���˺���ǿ������
    If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0 Then
        txtNotifyAdvice.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, 1, Array(lblNotifyAdviceDay, txtNotifyAdviceDay), InStr(mstrPrivs, "��������") > 0)
    txtNotifyAdviceDay.Text = Val(strPar)
    
    strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, "0000000000000", Array(lbl��������, chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), chkWarn(6), chkWarn(7), chkWarn(8), chkWarn(9), chkWarn(10), chkWarn(11), chkWarn(12)), InStr(mstrPrivs, "��������") > 0)
    For i = 1 To Len(strPar)
        If i - 1 <= chkWarn.UBound Then
            chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
        End If
    Next
    txt�������.Text = zlDatabase.GetPara("�������", glngSys, 1265, "3", Array(chkPatientFilter, txt�������))
    chkPatientFilter.Value = IIf(Val(txt�������.Text) = 0, 0, 1)
    txt�������.Enabled = (chkPatientFilter.Value = 1)
    '111016
    txt��Ժ����.Text = zlDatabase.GetPara("��Ժ����", glngSys, 1265, "0", Array(chkNewPati, txt��Ժ����))
    chkNewPati.Value = IIf(Val(txt��Ժ����.Text) = 0, 0, 1)
    txt��Ժ����.Enabled = (chkNewPati.Value = 1)
    '--56960:������,2013-01-17,��Ӳ���"��Ƭ���������"
    strPar = zlDatabase.GetPara("��Ƭ���������", glngSys, 1265, 0, Array(chkAmount), InStr(mstrPrivs, "��������") > 0)
    chkAmount.Value = IIf(Val(strPar) = 1, 1, 0)
    '54370:������,2013-05-02,��Ӳ���"ҽ��������Զ���λ��ҽ��ҳ��"
    strPar = zlDatabase.GetPara("ҽ��������Զ���λ��ҽ��ҳ��", glngSys, 1265, 0, Array(ChkCollate), InStr(mstrPrivs, "��������") > 0)
    ChkCollate.Value = IIf(Val(strPar) = 1, 1, 0)
    strPar = zlDatabase.GetPara("����������ʾ", glngSys, mlngModual, 0, Array(chkSound, cmdSoundSet), InStr(mstrPrivs, "��������") > 0)
    chkSound.Value = IIf(Val(strPar) = 1, 1, 0)
    txtMedRec.Text = zlDatabase.GetPara("������鷴������", glngSys, mlngModual, "3", Array(lblMedRec, txtMedRec), InStr(mstrPrivs, "��������") > 0)
    strPar = zlDatabase.GetPara("��λ��Ƭ����ʽ", glngSys, 1265, 0, Array(lblNewCard, optNewCard(0), optNewCard(1)), InStr(mstrPrivs, "��������") > 0)
    If Val(strPar) = 0 Then
        optNewCard(0).Value = True
    Else
        optNewCard(1).Value = True
    End If
    Call InitColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DeleteFile
End Sub

Private Sub InitColor()
    Dim strValue As String
    Dim lng�ؼ� As Long, lngһ�� As Long, lng���� As Long, lng���� As Long
    Const c��ɫ As Long = 8388736
    Const c��ɫ As Long = 255
    Const c��ɫ As Long = 16711680
    Const c��ɫ As Long = 16777215
    
    Call DeleteFile
    '��ȡ����ȼ���������(����ȡȱʡ����)
    strValue = zlDatabase.GetPara("�ؼ�������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(0)))
    lng�ؼ� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("һ��������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(1)))
    lngһ�� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("����������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(2)))
    lng���� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("����������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(3)))
    lng���� = IIf(strValue = "", c��ɫ, Val(strValue))
    
    '��ͼ
    mlngColor = lng�ؼ�
    Call DrawPoly
    img����ȼ�(0).Tag = mlngColor
    img����ȼ�(0).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lngһ��
    Call DrawPoly
    img����ȼ�(1).Tag = mlngColor
    img����ȼ�(1).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lng����
    Call DrawPoly
    img����ȼ�(2).Tag = mlngColor
    img����ȼ�(2).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lng����
    Call DrawPoly
    img����ȼ�(3).Tag = mlngColor
    img����ȼ�(3).Picture = img24.ListImages("K_" & mintIndex).Picture
End Sub

Private Sub img����ȼ�_Click(Index As Integer)
    picControl.Tag = Index
    picControl.Visible = True
    Call SetCOLOR(Val(img����ȼ�(Index).Tag))
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > 0 And x < Picture1.ScaleWidth And y > 0 And y < Picture1.ScaleHeight Then
        SetCapture Picture1.hwnd
        shpBorder.Visible = True
    Else
        ReleaseCapture
        shpBorder.Visible = False
    End If

    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = y \ (18 * Screen.TwipsPerPixelY)
    lCol = x \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    shpBorder.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If Picture1.Point(lX, lY) = -1 Then Exit Sub
    picColor.BackColor = Picture1.Point(lX, lY)
    Select Case CStr(Hex(picColor.BackColor))
    Case "0"
        lblColor = "��ɫ"
    Case "3399"
        lblColor = "��ɫ"
    Case "3333"
        lblColor = "���ɫ"
    Case "3300"
        lblColor = "����"
    Case "663300"
        lblColor = "����"
    Case "800000"
        lblColor = "����"
    Case "993333"
        lblColor = "����"
    Case "333333"
        lblColor = "��ɫ-80%"
    Case "80"
        lblColor = "���"
    Case "66FF"
        lblColor = "��ɫ"
    Case "8080"
        lblColor = "���"
    Case "8000"
        lblColor = "��ɫ"
    Case "808000"
        lblColor = "��ɫ"
    Case "FF0000"
        lblColor = "��ɫ"
    Case "996666"
        lblColor = "��-��"
    Case "808080"
        lblColor = "��ɫ-50%"
    Case "FF"
        lblColor = "��ɫ"
    Case "99FF"
        lblColor = "ǳ��ɫ"
    Case "CC99"
        lblColor = "���ɫ"
    Case "669933"
        lblColor = "����"
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
    Case "FF6633"
        lblColor = "ǳ��"
    Case "800080"
        lblColor = "������"
    Case "999999"
        lblColor = "��ɫ-40%"
    Case "FF00FF"
        lblColor = "�ۺ�"
    Case "CCFF"
        lblColor = "��ɫ"
    Case "FFFF"
        lblColor = "��ɫ"
    Case "FF00"
        lblColor = "����"
    Case "FFFF00"
        lblColor = "����"
    Case "FFCC00"
        lblColor = "����"
    Case "663399"
        lblColor = "÷��"
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
    Case "CC99FF"
        lblColor = "õ���"
    Case "99CCFF"
        lblColor = "��ɫ"
    Case "99FFFF"
        lblColor = "ǳ��"
    Case "CCFFCC"
        lblColor = "ǳ��"
    Case "FFFFCC"
        lblColor = "ǳ����"
    Case "FFCC99"
        lblColor = "����"
    Case "FF99CC"
        lblColor = "����"
    Case "FFFFFF"
        lblColor = "��ɫ"
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = y \ (18 * Screen.TwipsPerPixelY)
    lCol = x \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    picControl.Visible = False
    
    '��ָ����ɫ��ͼ
    mlngColor = picColor.BackColor
    img����ȼ�(Val(picControl.Tag)).Tag = mlngColor
    Call DrawPoly
    img����ȼ�(Val(picControl.Tag)).Picture = img24.ListImages("K_" & mintIndex).Picture
End Sub

Private Sub txtMedRec_GotFocus()
    Call zlControl.TxtSelAll(txtMedRec)
End Sub

Private Sub txtMedRec_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdvice_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdvice)
End Sub

Private Sub txtNotifyAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdviceDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdviceDay)
End Sub

Private Sub txtNotifyAdviceDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub SetCOLOR(vData As OLE_COLOR)
    mvarColor = vData
    Dim lRow As Long, lCol As Long
    shpValue.Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor = "���ɫ"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor = "����"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor = "����"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor = "����"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor = "����"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor = "��ɫ-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor = "���"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor = "���"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor = "��-��"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor = "��ɫ-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor = "��ɫ"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor = "ǳ��ɫ"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor = "���ɫ"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor = "����"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor = "ǳ��"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor = "������"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor = "��ɫ-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor = "�ۺ�"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor = "����"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor = "����"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor = "����"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor = "÷��"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor = "õ���"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor = "ǳ����"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor = "����"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor = "����"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 7
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
    shpBorder.Visible = False
    shpValue.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    shpValue.Visible = True
    If vData = tomAutoColor Or vData = -1 Then
    
    Else
        picColor.BackColor = vData
    End If
End Sub

Private Sub AddColor()
    Dim strFile As String
    mintIndex = mintIndex + 1
    '������Ϊ�ļ�,���������ͼƬʱ,���뵽imagelist���ʼ��ֻ�����һ��,Ӧ��������image�б������ͼƬID���
    
    strFile = App.Path & "\HLDJTMP" & mintIndex & ".BMP"
    SavePicture picDraw.Image, strFile
    picDraw.Picture = LoadPicture(strFile)
    img24.ListImages.Add , "K_" & mintIndex, picDraw.Picture
End Sub

Private Sub DrawPoly()
    Dim lngRgn As Long, lngBrush As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim PtInPoly() As POINTAPI

    '������򲢻�����
    ReDim PtInPoly(4) As POINTAPI
    PtInPoly(1).x = 0
    PtInPoly(1).y = 0
    PtInPoly(2).x = picDraw.ScaleWidth
    PtInPoly(2).y = 0
    PtInPoly(3).x = picDraw.ScaleWidth
    PtInPoly(3).y = picDraw.ScaleHeight
    PtInPoly(4).x = PtInPoly(1).x
    PtInPoly(4).y = PtInPoly(1).y
    
    '����ϵͳˢ��
    picDraw.Cls
    lngBrush = CreateSolidBrush(mlngColor)

    '�������ˢ�ӳɹ�,��ѡ��
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn picDraw.hdc, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    End If
    picDraw.Refresh
    
    Call AddColor
End Sub

Private Sub DeleteFile()
    Dim objFile As File
    For Each objFile In mobjFileSys.GetFolder(App.Path).Files
        If Left(objFile.Name, 7) = "HLDJTMP" Then
            mobjFileSys.DeleteFile objFile.Path, True
        End If
    Next
End Sub

Private Sub txt�������_GotFocus()
    txt�������.SelStart = 0
    txt�������.SelLength = txt�������.MaxLength
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ժ����_GotFocus()
    txt��Ժ����.SelStart = 0
    txt��Ժ����.SelLength = txt��Ժ����.MaxLength
End Sub

Private Sub txt��Ժ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
