VERSION 5.00
Begin VB.Form frmMeRec 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "��Ҹ�ҳ����"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10785
   Icon            =   "frmMeRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Appearance      =   0  'Flat
      Caption         =   "��ť����"
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtPic 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Text            =   "��Ҹ�ҳ�ı������"
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblTest 
      Appearance      =   0  'Flat
      Caption         =   "��Ҹ�ҳ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMeRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'���ڹ���:�����Զ��岡����Ҹ�ҳ

'����˵��:1.סԺ��ҳ�����ȹ̶�Ϊ10785,������ҳ�����ȹ̶�11985��������ҳ��ʽ
'         2.�����Caption,������ҳ����ҳǩ����
'         3.ע:���޸ĺ���SavePlugMec��ʱ���벻Ҫд���ʱ�Ĵ���
'         4.����CheckPlugMec:������Ҹ�ҳ������Ч�Լ��
'         5.����SavePlugMec:�齨������Ҹ�ҳ��������SQL
'         6.����LoadPlugMec:������Ҹ�ҳ��������
'         7.�����Tagֵ:���ڱ��洰���Ӧ������ҳҳǩ��index
'         8.�ؼ���Tagֵ:���ڱ������������Ϣ ��ʽ:((����:1/��ֹ:0) | ��ʾ��Ϣ| ����Tagֵ)
'         9.gblnChange:�жϱ�����ؼ�ֵ�Ƿ����ı�


Public gblnChange As Boolean '�Ƿ�ı�ؼ�����

'��ҳ����
Public glngSys As Long
Public glngModule As Long
Public glngPatiID As Long
Public glngPageID As Long
Public glngPatiType As Long


Private Sub cmdTest_Click()
    MsgBox "��Ҹ�ҳ����"
End Sub

Private Sub Form_Load()
    '���ÿ�ȹ̶�Ϊ10785 ����סԺ��ҳ��׼��ʽ
    Me.Width = 10785
    
    '���ÿ�ȹ̶�Ϊ11985 ���ֲ�����ҳ��׼��ʽ
'    Me.Width = 11985
End Sub
'

Public Function CheckPlugMec(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef objTmp As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ҹ�ҳ������Ч�Լ��
    '���أ�True�ǳɹ���False��ʧ��
    '������objTmp ������ʾ�ؼ�
    '      �ؼ���tagֵ:������ʾ��Ϣ  �� : ��ʾ��Ϣ|(����:1/��ֹ:0)|��ҳIndex
    '      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '      lngPageID����ҳID��
    
    '����:���ͨ������true,���򷵻�False
    '����:��͢��
    '����:2017��6��20�� 11:52:48
    On Error GoTo errHandle
    CheckPlugMec = True

    If txtPic.Text = "" Then
        txtPic.Tag = "��Ҳ�����ҳ�ı�����Ϊ��|1" & "|" & Val(Me.Tag)
        Set objTmp = txtPic
        CheckPlugMec = False
        Exit Function
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function SavePlugMec(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�齨������Ҹ�ҳ��������SQL��ͨ������������gOracle����ִ��
    '���أ�True�ǳɹ���False��ʧ��
    '������ lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '       lngPatiID:����id
    '      lngPageID����ҳID��
    '����:����ͨ������true,���򷵻�False
    '����:��͢��
    '����:2017��6��20�� 11:52:48
    Dim strSql As String
    On Error GoTo errHandle
    
    strSql = "zl_������Ϣ�ӱ�_Update(" & lngPatiID & ",'���֤��״̬','��ʧ����')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    SavePlugMec = True
    
    gblnChange = False
    Exit Function
errHandle:
    SavePlugMec = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function LoadPlugMec(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngPatiType As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ҹ�ҳ��������
    '���أ�True�ǳɹ���False��ʧ��
    '      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '      lngPageID����ҳID��
    '      lngPatiType����������:1-����,2-סԺ
    '      lngPatiID-����id
    '����:��͢��
    '����:2017��6��21�� 9:52:48

   On Error GoTo errHandle
    LoadPlugMec = True
    txtPic.Text = lngSys & "|" & lngModule & "|" & lngPatiID & "|" & lngPageID & "|" & lngPatiType
    gblnChange = False
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function





'�Ƿ�ı�ؼ�����
Private Sub txtPic_Change()
    gblnChange = True
End Sub





Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "������Ҹ�ҳ ж���ˣ���������"
End Sub

