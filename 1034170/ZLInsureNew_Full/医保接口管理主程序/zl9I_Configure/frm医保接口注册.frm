VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmҽ���ӿ�ע�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ע��ҽ���ӿ�"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmҽ���ӿ�ע��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt��ռ� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      MaxLength       =   100
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txt�м���û��� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   4
      Top             =   570
      Width           =   2775
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   17
      Top             =   3030
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1860
      TabIndex        =   16
      Top             =   3030
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "ע����Ϣ"
      Height          =   1530
      Left            =   150
      TabIndex        =   7
      Top             =   1380
      Width           =   4065
      Begin VB.TextBox txt˵�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1080
         Width           =   2745
      End
      Begin VB.TextBox txt��Կ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1530
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         MaxLength       =   40
         TabIndex        =   11
         Top             =   690
         Width           =   2745
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   810
         MaxLength       =   3
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lbl˵�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   12
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl��Կ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Կ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   14
         Top             =   1590
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   10
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbl��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   8
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdҽ���ӿ� 
      Caption         =   "��"
      Height          =   285
      Left            =   3900
      TabIndex        =   2
      Top             =   180
      Width           =   285
   End
   Begin VB.TextBox txtҽ���ӿڲ��� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   2475
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl��ռ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���ռ�"
      Height          =   180
      Left            =   450
      TabIndex        =   5
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label lbl�м���û��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�м���û���"
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   630
      Width           =   1080
   End
   Begin VB.Label lblҽ���ӿڲ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ���ӿڲ���"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmҽ���ӿ�ע��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mintRegist As Integer       '0-ʧ�ܻ�ȡ��;1-����;2-����;3-����ע�ᣨָ������ͬ��ҽ�������Լ��м���û��ȣ������಻ͬ��
Private mintInsure As Integer       '����������
Private mstrInsureUser As String    '�м���û���
Private mstrInsureName As String    'ҽ���ӿ�����
Private mstrInsureTablespace As String
Private mstrPath As String          'ע���ļ�·��
Private mstrComponent As String     '��������
Private mstrDemo As String          '˵��
Private mbln����ע�� As Boolean     '����ע���־

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim blnRegist As Boolean        '�Ƿ������ظ�ע��
    Dim strFile As String, strMessage As String
    Dim objTest As Object
    Dim rsTest As New ADODB.Recordset
    '��鱣��������Ƿ��Ѵ��ڸ����࣬������ڣ����������°�װ����ʱֻ����spNew.sql
    
    '���裺
    '1������ļ��ĺϷ��ԣ�zl9I_�ļ�����
    '2������ָ����ҽ���ӿڲ�����ʧ�����˳�
    '3�������Ϸ�����֤��ʧ�����˳�
    '4�������ע���嵥�У��Ƿ���ڲ�ͬ���࣬��ͬ�������������������ʾ
    '----�����ҽ���ӿ��Ѱ�װ���ٴΰ�װ˵����Ҫ���������ű���5-9����----
    '5�����ע���ļ��ĺϷ��ԣ�ʧ�����˳�
    '6�����а�װ�ű���ʧ�����˳�
    
    If Trim(txtҽ���ӿڲ���.Text) = "" Then
        MsgBox "��ѡ����Ҫע���ҽ��������", vbInformation, gstrSysName
        cmdҽ���ӿ�.SetFocus
        Exit Sub
    End If
    strFile = Mid(txtҽ���ӿڲ���.Text, 1, Len(txtҽ���ӿڲ���.Text) - 4)
    
    '����Ƿ������ظ�ע��
    Set objTest = CreateObject(strFile & ".CLS" & Mid(strFile, 4))
    strMessage = objTest.I_Support(I_Support�ظ�ע��)
    blnRegist = (Val(strMessage) = 1)
    
    '4��
    mstrSQL = " Select A.���,A.����,B.���� As ҽ������,B.�û���,B.��ռ�" & _
              " From ������� A,zlInsureComponents B" & _
              " Where A.���=B.���� And Upper(B.����)='" & txtҽ���ӿڲ���.Text & "'"
    Call zlDatabase.OpenRecordset(rsTest, mstrSQL, "װ����ע���ҽ���ӿ�")
    
    mintRegist = 0
    If rsTest.RecordCount <> 0 Then
        rsTest.Filter = "���<>" & txt���.Text
        If rsTest.RecordCount <> 0 Then
            If blnRegist Then
                '�������ע���ԭ���ǣ�����һ��ҽ���ӿڲ���Ӧ���ڶ��������ҽ�������ݵ�����������
                If MsgBox("����������ע���ҽ���ӿڣ��䲿�������뱾��ע��ҽ���ӿڵĲ�������һ�£��Ƿ����ע�᣿", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then rsTest.Filter = 0: Exit Sub
                mintRegist = 3
                '�û������ռ������е�Ϊ׼
                txt�м���û���.Text = Nvl(rsTest!�û���)
                txt��ռ�.Text = Nvl(rsTest!��ռ�)
            Else
                MsgBox "��ҽ���ӿڲ������ظ�ע��ʹ�ã�", vbInformation, gstrSysName
                rsTest.Filter = 0: Exit Sub
            End If
        Else
            rsTest.Filter = "���=" & txt���.Text
            '˵�����ٴ�ע��
            If rsTest.RecordCount <> 0 Then mintRegist = 2
        End If
        rsTest.Filter = 0
    End If
    
    '���������ͬ���൫�������Ʋ�ͬ�ģ�Ҳ��Ϊ������ע��
    If mintRegist = 0 Then
        mstrSQL = " Select A.���,A.����,upper(B.����) As ҽ������" & _
                  " From ������� A,zlInsureComponents B" & _
                  " Where A.���=B.���� And A.���=" & txt���.Text
        Call zlDatabase.OpenRecordset(rsTest, mstrSQL, "װ����ע���ҽ���ӿ�")
        If rsTest.RecordCount <> 0 Then
            MsgBox "����������ע���ҽ���ӿڣ��䱣������뵱ǰ�ӿ�һ�£���ҽ���ӿڲ�����ͬ�����������ע�ᣡ", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If mintRegist = 0 Then mintRegist = 1
    
    '������ݺϷ���
    If Trim(txt���.Text) = "" Then
        MsgBox "���������Ų���Ϊ�գ�", vbInformation, gstrSysName
        txt���.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txt���.Text) Then
        MsgBox "�����������к��зǷ��ַ���", vbInformation, gstrSysName
        txt���.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Me.txt˵��.Text, vbFromUnicode)) > 100 Then
        MsgBox "˵�����ܳ���50�����ֻ�100���ַ���", vbInformation, gstrSysName
        txt˵��.SetFocus
        Exit Sub
    End If
    
    mstrComponent = Me.txtҽ���ӿڲ���.Text
    mstrPath = Me.txtҽ���ӿڲ���.Tag
    mintInsure = Val(txt���.Text)
    mstrInsureUser = Trim(txt�м���û���.Text)
    mstrInsureTablespace = Trim(txt��ռ�.Text)
    mstrInsureName = txt����.Text
    mstrDemo = Trim(txt˵��.Text)
    
    Unload Me
    Exit Sub
End Sub

Private Sub cmdҽ���ӿ�_Click()
    Dim strFile As String, strPath As String
    Dim strMessage As String
    Dim arrMessage
    Dim str��� As String, str���� As String, str˵�� As String
    Dim objTest As Object
    On Error GoTo ErrHand
    
    With CommonDialog1
        .Filter = "ҽ������(*.dll)|*.dll"
        .ShowOpen
        Call GetFileOrPath(.FileName, strFile, strPath)
        strFile = Mid(strFile, 1, Len(strFile) - 4)
    End With

    '1��
    If Mid(strFile, 1, 5) <> "ZL9I_" Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������1", vbInformation, gstrSysName
        Exit Sub
    End If
    '2��
    On Error Resume Next
    Err = 0
    Set objTest = CreateObject(strFile & ".CLS" & Mid(strFile, 4))
    If Err <> 0 Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������2", vbInformation, gstrSysName
        Exit Sub
    End If
    '3��
    Err = 0
    strMessage = objTest.I_RegInfo
    If Err <> 0 Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������3", vbInformation, gstrSysName
        Set objTest = Nothing
        Exit Sub
    End If
    
    '3.1
    arrMessage = Split(strMessage, "|")
    If Not (UBound(arrMessage) >= 1) Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������3.1", vbInformation, gstrSysName
        Exit Sub
    End If
    str��� = Val(arrMessage(0))
    str���� = UCase(arrMessage(1))
    str˵�� = UCase(arrMessage(2))
    If str��� = 0 Or Trim(str����) = "" Then
        MsgBox "ҽ���ӿ�ע����Ϣ��������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '4���ж��Ƿ�֧���м��
    Err = 0
    strMessage = objTest.I_Support(I_Support�м��)
    If Err <> 0 Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������4", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Me.txt���.Text = str���
    Me.txt����.Text = str����
    Me.txt˵��.Text = str˵��

    Me.txtҽ���ӿڲ���.Text = strFile & ".DLL"
    Me.txtҽ���ӿڲ���.Tag = strPath
    
    '֧���м�⣬���û������ռ�ȱʡΪҽ���������ļ���
    If Val(strMessage) = 1 Then
        txt�м���û���.Text = strFile
        txt��ռ�.Text = strFile
    End If
ErrHand:
    Exit Sub
End Sub

Public Function ShowRegist(intInsure As Integer, strInsureUser As String, strInsureTablespace As String, _
    strInsureName As String, strDemo As String, strComponent As String, strPath As String) As Integer
    'intInsure:����������
    'strInsureName:ҽ���ӿ�����,eg��������ҽ�����ġ�
    'strComponent:ҽ����������
    'strPath:ҽ�������ļ�·��
    mintRegist = 0
    mintInsure = 0
    mstrInsureUser = ""
    mstrInsureTablespace = ""
    mstrInsureName = ""
    mstrDemo = ""
    mstrComponent = ""
    mstrPath = ""
    
    Me.Show 1
    
    ShowRegist = mintRegist
    If mintRegist > 0 Then
        strPath = mstrPath
        strComponent = mstrComponent
        intInsure = mintInsure
        strInsureUser = mstrInsureUser
        strInsureTablespace = mstrInsureTablespace
        strInsureName = mstrInsureName
        strDemo = mstrDemo
    End If
End Function

Private Sub GetFileOrPath(ByVal strInput As String, strFile As String, strPath As String)
    Dim intPos As Integer
    '�����������ļ�·���������ļ�·�����ļ�����C:\Appsoft\Apply\zl9Insure.dll�����ص��ļ�����zl9Insure.dll����·������C:\Appsoft\Apply
    intPos = 1
    Do While True
        If InStr(intPos, strInput, "\") = 0 Then Exit Do
        intPos = InStr(intPos, strInput, "\") + 1
    Loop
    If intPos = 1 Then Exit Sub
    
    strPath = UCase(Mid(strInput, 1, intPos - 2))
    strFile = UCase(Mid(strInput, intPos))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt���_GotFocus()
    Call zlControl.TxtSelAll(txt���)
End Sub
