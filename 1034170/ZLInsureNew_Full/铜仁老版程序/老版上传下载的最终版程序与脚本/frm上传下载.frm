VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm�ϴ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ϴ����س���"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frm�ϴ�����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8130
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picĿ¼ 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   300
      ScaleHeight     =   1305
      ScaleWidth      =   7575
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   870
      Width           =   7575
      Begin VB.ComboBox cbo�籣�� 
         Height          =   300
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   930
         Width           =   5355
      End
      Begin VB.CommandButton cmdĿ¼ 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   7200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
      End
      Begin VB.CommandButton cmdĿ¼ 
         Caption         =   "��"
         Height          =   240
         Index           =   1
         Left            =   7200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   585
         Width           =   255
      End
      Begin VB.TextBox txtDown 
         Height          =   300
         Left            =   2130
         MaxLength       =   40
         TabIndex        =   6
         Top             =   540
         Width           =   5355
      End
      Begin VB.TextBox txtUp 
         Height          =   300
         Left            =   2130
         MaxLength       =   40
         TabIndex        =   3
         Top             =   150
         Width           =   5355
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�籣��"
         Height          =   180
         Index           =   0
         Left            =   1530
         TabIndex        =   8
         Top             =   990
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frm�ϴ�����.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������Ŀ¼"
         Height          =   180
         Index           =   7
         Left            =   990
         TabIndex        =   5
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����ϴ�Ŀ¼"
         Height          =   180
         Index           =   6
         Left            =   990
         TabIndex        =   2
         Top             =   210
         Width           =   1080
      End
   End
   Begin MSComctlLib.TabStrip tabHost 
      Height          =   1965
      Left            =   135
      TabIndex        =   0
      Top             =   390
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   3466
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra���� 
      Caption         =   "���ȱ���"
      Height          =   705
      Left            =   135
      TabIndex        =   19
      Top             =   4560
      Width           =   7905
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   255
         Left            =   900
         TabIndex        =   21
         Top             =   300
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl��Ŀ 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.Frame fra�ϴ� 
      Caption         =   "�ϴ�"
      Height          =   1965
      Left            =   4125
      TabIndex        =   11
      Top             =   2490
      Width           =   3915
      Begin VB.CommandButton cmd�ָ� 
         Caption         =   "�ָ����ϴ�"
         Height          =   350
         Left            =   2640
         TabIndex        =   18
         Top             =   780
         Width           =   1100
      End
      Begin VB.CommandButton cmd�ϴ� 
         Caption         =   "��ʼ�ϴ�"
         Height          =   350
         Left            =   2640
         TabIndex        =   17
         Top             =   1440
         Width           =   1100
      End
      Begin VB.Label lbl�ϴ� 
         AutoSize        =   -1  'True
         Caption         =   "����ϴ����ڣ�"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1110
         Width           =   1260
      End
      Begin VB.Label lbl�ϴ�˵�� 
         Caption         =   "   �ϴ�����ÿ��ֻ��ִ��һ�Ρ����ʧ�ܿ�������ִ�С�"
         Height          =   495
         Left            =   1110
         TabIndex        =   13
         Top             =   360
         Width           =   2625
      End
      Begin VB.Image img�ϴ� 
         Height          =   480
         Left            =   210
         Picture         =   "frm�ϴ�����.frx":0884
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   1965
      Left            =   135
      TabIndex        =   10
      Top             =   2490
      Width           =   3915
      Begin VB.CommandButton cmd���� 
         Caption         =   "��ʼ����"
         Height          =   350
         Left            =   2580
         TabIndex        =   15
         Top             =   1440
         Width           =   1100
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����������ڣ�"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   1050
         Width           =   1260
      End
      Begin VB.Image img���� 
         Height          =   480
         Left            =   210
         Picture         =   "frm�ϴ�����.frx":114E
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lbl����˵�� 
         Caption         =   "   �ڵ�һ��ʹ��ҽ���ӿ�ǰ��������Ƚ���һ�����ء����س��������ʱ���С�"
         Height          =   615
         Left            =   1110
         TabIndex        =   12
         Top             =   360
         Width           =   2565
      End
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ע�����ڲ�ͬ����λ�ڲ�ͬ�������У������ϴ�����ֻ����Ե�ǰ�������С�"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   90
      Width           =   6120
   End
End
Attribute VB_Name = "frm�ϴ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������ļ�
Private Enum HostField
    hf�������� = 0
    hf�������� = 1
    hfҽ���� = 2
    hf��Ч�� = 3
    hfװǮ��� = 4
    hf������������� = 5
    hf����������� = 6
    hf��Ŀ������� = 7
    hf����������� = 8
    hf���ݸɲ�������� = 9
    hf������Ա������� = 10
    hf�Ƿ���� = 11
    hf�����·� = 12
End Enum

Private Enum HostParamField
    hp�������� = 0
    hp��ʼ���� = 1
    hp��ֹ���� = 2
    hp�ϴ�IP = 3
    hp�ϴ��û� = 4
    hp�ϴ����� = 5
    hp�ϴ�Ŀ¼ = 6
    hp����IP = 7
    hp�����û� = 8
    hp�������� = 9
    hp����Ŀ¼ = 10
    hpװǮIP��ַ1 = 13
    hpװǮIP��ַ2 = 14
    hp�����·� = 15
End Enum

Private Enum CenterField
    cf�������� = 0
    cf���Ĵ��� = 1
    cf�������� = 2
    cf�����˻���֧�������Ը� = 3
    cf��д����Ʊ = 4
End Enum

Private Enum PolicyField
    pol���Ĵ��� = 0
    pol�㷨 = 1
    pol��ֵ = 2
    pol�������� = 3
    pol�������� = 4
    pol���������� = 5
    pol�����ڶ��� = 6
    polͳ��ⶥ�� = 7
    polͳ����� = 8
    pol��ֵ���� = 9
    pol�ⶥ���� = 10
    polʹ���ۼƱ��� = 11
    pol�������Բ��·� = 12
    pol��չ���䱣�ձ��� = 13
    pol���䱨������ = 14
    pol���䱨���޶� = 15
    pol���䱨���޶����� = 16
    pol���䱨�����𸶽� = 17
    pol��չ�������� = 18
    pol��չ�������� = 19
    pol��չ�󲡱��� = 20
    pol������Ŀ�۸� = 21
    pol�����𸶽����� = 22 '0-��ԭ�𸶽�1�������ۣ�2���𸶽�
    pol��������סԺ���� = 23 '1����һ�Σ�0������
End Enum

Private Enum ParamField
    par���Ĵ��� = 0
    par�������� = 1
    parҽԺ�ȼ� = 2
    parְ����� = 3
    par���� = 4
    par��һ����ʼֵ = 5
    par��һ�α������� = 6
    par�ڶ�����ʼֵ = 7
    par�ڶ��α������� = 8
    par��������ʼֵ = 9
    par�����α������� = 10
    par���Ķ���ʼֵ = 11
    par���Ķα������� = 12
    par�������ʼֵ = 13
    par����α������� = 14
End Enum

Private Enum ItemField
    if��Ŀ��� = 0
    if���ֱ��� = 1
    ifƴ������ = 2
    ifҩ������ = 3
    if��λ = 4
    if���ͱ��� = 5
    if������� = 6
    if�Ƿ���ҩ = 7
    if�Ƿ�ҽ�� = 8
    if���۸����� = 9
    if�����Ը����� = 10
    if�۸� = 11
    if��Ŀ�ں� = 12
    if�������� = 13
    if˵�� = 14
    ifʡ���޼� = 15
    if�м��޼� = 16
    if�ؼ��޼� = 17
    if�缶�޼� = 18
    if�ؼ���Ŀ = 19
    if�ؼ��Ը����� = 20
End Enum

'��Ϊ����ʹ�õ��ļ�ϵͳ����
Dim mobjFileSys As New FileSystemObject

Private mstr�������� As String
Private mstr�������� As String
Private mstr����InOracle As String
Private mstr���InOracle As String
Private mstr����InStr As String

Private mstrԭҽ���� As String
Private mlngװǮ��� As Long
Private mlng������������� As Long
Private mlng����������� As Long
Private mlng��Ŀ������� As Long
Private mlng����������� As Long
Private mlng���ݸɲ���� As Long
Private mlng������Ա������� As Long

Private mstr�ϴ�IP As String
Private mstr�ϴ��û� As String
Private mstr�ϴ����� As String
Private mstr����IP As String
Private mstr�����û� As String
Private mstr�������� As String
Private mstrԶ���ϴ�Ŀ¼ As String
Private mstrԶ������Ŀ¼ As String
Private mstr�����ϴ�Ŀ¼ As String
Private mstr��������Ŀ¼ As String

Private mstrҽԺ���� As String
Private mstrҽԺ���� As String

'�ϴ�����ר��
Private mstr��ʼ���� As String
Private mstr�������� As String
Private mstr�ս����� As String 'Ҳ���Ƿ��÷�������
Private mstrȱʡ��ʼ���� As String

Private mdat��ʼ���� As Date
Private mdat�������� As Date

Private mblnLoad As Boolean

Private Sub cbo�籣��_Click()
    gcnҽ��.Execute "Update �������� A set �籣��='" & cbo�籣��.Text & "' where  A.���� = " & TYPE_ͭ���� & " And A.���� = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    If mblnLoad = False Then Exit Sub
    
    On Error GoTo errHandle
    
    gstrSQL = "SELECT A.����,A.���� FROM �������� A,������������ B " & _
              " Where A.���� = " & TYPE_ͭ���� & " And A.���� = B.���� And A.���� = B.���� " & _
              "    AND nvl(B.��ʼ����,to_date('2000-01-01','yyyy-MM-dd'))<=SYSDATE  AND nvl(B.��ֹ����,to_date('3000-01-01','yyyy-MM-dd'))>=trunc(SYSDATE)" & _
              " Order by A.����"
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�з��ֿ��Խ����ϴ����ص�ҽ�������������ʼ�������Ƿ���ȷ��", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    tabHost.Tabs.Clear
    Do Until rsTemp.EOF
        tabHost.Tabs.Add , "K" & rsTemp("����"), rsTemp("����")
        rsTemp.MoveNext
    Loop
    tabHost.Tabs(1).Selected = True
    
    mblnLoad = False
    
    'ִ����Ӧ�Ĺ���
    If gintType = 1 Then
        '����
        Call cmd����_Click
    ElseIf gintType = 2 Then
        '�ϴ�
        Call cmd�ϴ�_Click
    End If
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    
End Sub

Private Sub tabHost_Click()
    Dim i As Integer, j As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    'ȡ�����������籣��
    Me.cbo�籣��.Clear
    gstrSQL = "Select ����||'-'||���� AS �籣�� from zlyb.��������Ŀ¼ where ��������='" & Mid(tabHost.SelectedItem.Key, 2) & "' order by ���"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    With rsTemp
        Do While Not .EOF
            Me.cbo�籣��.AddItem !�籣��
            .MoveNext
        Loop
    End With
    
    gstrSQL = "SELECT A.�����ϴ���ַ,A.�������ص�ַ,A.�籣�� FROM �������� A " & _
              " Where A.���� = " & TYPE_ͭ���� & " And A.���� = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    txtDown.Text = NVL(rsTemp("�������ص�ַ"))
    txtUp.Text = NVL(rsTemp("�����ϴ���ַ"))
    j = Me.cbo�籣��.ListCount
    For i = 1 To j
        If Me.cbo�籣��.List(i - 1) = NVL(rsTemp!�籣��) Then
            Me.cbo�籣��.ListIndex = i - 1
            Me.cbo�籣��.Enabled = False
            Exit For
        End If
    Next
    
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Sub cmdĿ¼_Click(Index As Integer)
    Dim strTitle As String
    Dim strPath As String
    
    If Index = 0 Then
        strTitle = "��ѡ�񱣴��ϴ��ļ���Ŀ¼��"
    Else
        strTitle = "��ѡ�񱣴������ļ���Ŀ¼��"
    End If
    
    strPath = OpenDir(Me, strTitle)
    If StrIsValid(strPath, 50, , "Ŀ¼��") = False Then
        Exit Sub
    End If
    If strPath <> "" Then
        '����Ŀ¼��
        If Index = 0 Then
            gcnҽ��.Execute "Update �������� A set �����ϴ���ַ='" & strPath & "' where  A.���� = " & TYPE_ͭ���� & " And A.���� = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
            txtUp.Text = strPath
        Else
            gcnҽ��.Execute "Update �������� A set �������ص�ַ='" & strPath & "' where  A.���� = " & TYPE_ͭ���� & " And A.���� = '" & Mid(tabHost.SelectedItem.Key, 2) & "'"
            txtDown.Text = strPath
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    lbl����.Caption = "����������ڣ���"
    lbl�ϴ�.Caption = "����ϴ����ڣ���"
    lbl����.Tag = ""
    lbl�ϴ�.Tag = ""
    
    gstrSQL = "select ����ģʽ,max(��������) as ���� from �ϴ����� group by ����ģʽ"
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Do Until rsTemp.EOF
        If rsTemp("����ģʽ") = 2 Then
            lbl����.Caption = "����������ڣ�" & Format(rsTemp("����"), "yyyy-MM-dd hh:mm:ss")
            lbl����.Tag = Format(rsTemp("����"), "yyyy-MM-dd hh:mm:ss")
        Else
            lbl�ϴ�.Caption = "����ϴ����ڣ�" & Format(rsTemp("����"), "yyyy-MM-dd")
            lbl�ϴ�.Tag = Format(rsTemp("����"), "yyyy-MM-dd")
        End If
        rsTemp.MoveNext
    Loop
    mblnLoad = True
    pgb.Value = 0
End Sub

Private Sub cmd�ָ�_Click()
    Dim datMax As Date
    
    If IsDate(lbl�ϴ�.Tag) = False Then
        MsgBox "��δ���й������ϴ���", vbInformation, gstrSysName
        Exit Sub
    End If
    datMax = CDate(lbl�ϴ�.Tag)
    mdat��ʼ���� = datMax
    mdat�������� = datMax
    If frm�ظ��ϴ�����.GetTimeScope(mdat��ʼ����, mdat��������, datMax) = False Then
        Exit Sub
    End If
    mdat��ʼ���� = mdat��ʼ���� - 1  'Ϊ�˴���ʼ���������
    
    SetEnable False
   
    DoEvents
    Call �ϴ�����(True)
    '���ϴ������п������ύ�ĳ�������ǿ�ƻع�
    On Error Resume Next
    gcnOracle.RollbackTrans
    gcnҽ��.RollbackTrans
    
    Call Form_Load
    
    SetEnable True
    MsgBox "�ָ����ϴ�����������ɡ�", vbInformation, gstrSysName
   
End Sub

Private Sub cmd�ϴ�_Click()
    SetEnable False
    
    Call �ϴ�����
    '���ϴ������п������ύ�ĳ�������ǿ�ƻع�
    On Error Resume Next
    gcnOracle.RollbackTrans
    gcnҽ��.RollbackTrans
    
    Call Form_Load
    
    SetEnable True
    
    If gintType <> 0 Then
        Unload Me
    Else
        MsgBox "�ϴ�����������ɡ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmd����_Click()
    SetEnable False
    
    Call ��������
    Call Form_Load
    
    SetEnable True
    
    If gintType <> 0 Then
        Unload Me
    Else
        MsgBox "���ز���������ɡ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    cmd����.Enabled = blnEnable
    cmd�ϴ�.Enabled = blnEnable
    cmd�ָ�.Enabled = blnEnable
    
    If blnEnable = False Then
        MousePointer = vbHourglass
    Else
        MousePointer = vbDefault
    End If
End Sub

Private Sub ��������()
    Dim rsHost As New ADODB.Recordset
    Dim varHost As Variant

    On Error GoTo errHandle
    If GetҽԺ���� = False Then Exit Sub
    
    '���ڲ�ͬ�����ķ�������ͬ������Ҫ�ֱ𲦺ţ����Էֱ�����
    rsHost.Open "select * from �������� where ����=" & TYPE_ͭ���� & " And ���� = '" & Mid(tabHost.SelectedItem.Key, 2) & "'", gcnҽ��, adOpenStatic, adLockReadOnly

    '���ڿ����ж��ҽ�����ģ������һ��ѭ������
    Do Until rsHost.EOF
        '��ò���
        If Get��������(rsHost) = False Then Exit Sub
        
        '���غ������ݰ�
        If DownHost(rsHost("����"), varHost) = False Then
            MsgBox mstr�������� & "��������������ء����Ӧ���Ĵ����������ļ���û���ҵ���", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Get�����б�(rsHost("����"))
        
        If Is����װǮ(rsHost("����")) = True Then
            '������Ϊ����װǮ�Ŵ���
            If mlngװǮ��� = 0 Or Val(varHost(hfװǮ���)) > mlngװǮ��� Then
                '���װǮ�嵥������
                If DownװǮ(varHost(hfװǮ���), varHost(hfҽ����)) = False Then
                    Exit Sub
                End If
            End If
        End If
        
        If mlng���ݸɲ���� = 0 Or Val(varHost(hf���ݸɲ��������)) > mlng���ݸɲ���� Then
            '���װǮ�嵥������
            If Down����(varHost(hf���ݸɲ��������)) = False Then
                Exit Sub
            End If
        End If
        
        '��ɵ�λ��Ϣ�������޶����ݵ����أ�ÿ�ζ����أ�
        If Down��λ��Ϣ = False Then Exit Sub
        If Down�����޶� = False Then Exit Sub
        
        If mlng������Ա������� = 0 Or Val(varHost(hf������Ա�������)) > mlng������Ա������� Then
            '���װǮ�嵥������
            If Down����(varHost(hf������Ա�������)) = False Then
                Exit Sub
            End If
        End If

        If mlng������������� = 0 Or Val(varHost(hf�������������)) > mlng������������� Then
            '��ɺ�����������
            If Down������(varHost(hf�������������), varHost(hfҽ����)) = False Then
                Exit Sub
            End If
        End If
        
        If mlng����������� = 0 Or Val(varHost(hf�����������)) > mlng����������� Then
            '��ɼ����嵥������
            If Down����(varHost(hf�����������)) = False Then
                Exit Sub
            End If
        End If
        
        If mlng����������� = 0 Or Val(varHost(hf�����������)) > mlng����������� Then
            '�����Ŀ������
            If Down����(varHost(hf�����������), varHost(hfҽ����)) = False Then
                Exit Sub
            End If
        End If
        
        If mlng��Ŀ������� = 0 Or Val(varHost(hf��Ŀ�������)) > mlng��Ŀ������� Then
            '�����Ŀ������
            If Down��Ŀ(varHost(hf��Ŀ�������)) = False Then
                Exit Sub
            End If
        End If

        '��¼������־
        gstrSQL = "insert into �ϴ����� (��������,�û���,����ģʽ,���Ĵ���,�ļ���) " & _
                  "values(sysdate,substr(user,1,20),'2','" & rsHost("����") & "','Center.pak')"
        gcnҽ��.Execute gstrSQL

        rsHost.MoveNext
    Loop

    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Function DownHost(�������� As String, var��� As Variant) As Boolean
'���ܣ�������������ص���Ϣ
'��������������   ��ǰ���ص���������
'      var���    �����ļ��Ƿ���Ҫ���ص��ж����
'���أ��ܳɹ���������True
    Dim objText As TextStream, strLine As String
    Dim str�� As String '��ʾ��ǰ�Ǵ��������ļ�����һ��
    Dim varTemp As Variant, rsTemp As New ADODB.Recordset
    Dim col���� As New Collection, lng��� As Long, bln���� As Boolean
    Dim str��ʼ���� As String, str��ֹ���� As String
    
    '��������Center.pak�����õ�Center�ļ�
    If DownLoadFile("Center.pak") = False Then
        Exit Function
    End If
    
    Set mobjFileSys = New FileSystemObject
    Set objText = mobjFileSys.OpenTextFile(mstr��������Ŀ¼ & "Center")
    gcnҽ��.BeginTrans 'ʹ��������Ƶ�Ŀ����Ҫ������������������ɾ������һ�𣬷�����ܳ���û���������������
    On Error GoTo errHandle
    
    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        If strLine <> "" Then
            If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
                '�α�־
                str�� = UCase(Mid(strLine, 2, Len(strLine) - 2))
            Else
                varTemp = Split(strLine, "|")
                Select Case str��
                    Case "HOSTS"
                        If varTemp(hf��������) = �������� Then
                            '��ǰ������ֱ�Ӹ���һЩû�����������Ĳ���
                            gstrSQL = "Update �������� Set ��Ч��=" & GetDateForOracle(varTemp(hf��Ч��)) & _
                                    ",�Ƿ����=" & varTemp(hf�Ƿ����) & "  Where ����=" & TYPE_ͭ���� & " and ����='" & �������� & "'"
                            gcnҽ��.Execute gstrSQL
                            'ͬʱ�����÷���ֵ
                            bln���� = True
                            var��� = Split(strLine, "|")
                            
                            '----------------------------Ϊ�˱�֤��������������������Ϣֻ��һ�Σ������ڴ�ִ��
                            'ɾ���ϴ�������Ϣ�������ڴ���[HOSPARAMS]���ؽ�
                            gstrSQL = "Delete ������������ Where ����=" & TYPE_ͭ���� & " and ����='" & �������� & "'"
                            gcnҽ��.Execute gstrSQL
                            
                            'ֹͣ�������ģ�Ȼ���ڴ���[CENTERAGENCY]ʱ������
                            gstrSQL = "Update ��������Ŀ¼ Set ����ģʽ=0 Where ����=" & TYPE_ͭ����
                            gcnҽ��.Execute gstrSQL
                        Else
                            '���ڴ����·�������Ҫ�ȼ���Ƿ����Ѿ������·���������Ѿ����ڣ��򲻴���
                            gstrSQL = "Select ���� From �������� Where ����=" & TYPE_ͭ���� & " and ����='" & varTemp(hf��������) & "'"
                            Call OpenRecordset(rsTemp)
                            If rsTemp.RecordCount > 0 Then
                                '�Ѿ������
                                col����.Add True, "K" & varTemp(hf��������)
                            Else
                                '��������������Ϣ
                                col����.Add False, "K" & varTemp(hf��������)
                                gstrSQL = "Insert Into �������� (����,����,����,�Ƿ����,��Ч��) VALUES (" & _
                                    TYPE_ͭ���� & ",'" & varTemp(hf��������) & "','" & varTemp(hf��������) & "'," & varTemp(hf�Ƿ����) & _
                                    "," & GetDateForOracle(varTemp(hf��Ч��)) & ")"
                                gcnҽ��.Execute gstrSQL
                            End If
                        End If
                    Case "HOSTPARAMS"
                        str��ʼ���� = varTemp(hp��ʼ����)
                        If IsDate(str��ʼ����) = True Then
                            str��ʼ���� = "To_date('" & Format(CDate(str��ʼ����), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                        Else
                            str��ʼ���� = "To_date('2003-01-01','yyyy-MM-dd')"
                        End If
                        str��ֹ���� = varTemp(hp��ֹ����)
                        If IsDate(str��ʼ����) = True Then
                            str��ֹ���� = "To_date('" & Format(CDate(str��ֹ����), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                        Else
                            str��ֹ���� = "To_date('3000-01-01','yyyy-MM-dd')"
                        End If
                        If varTemp(hp��������) = �������� Then
                            '��ǰ������ֱ�Ӹ���һЩû�����������Ĳ���
                            gstrSQL = "Insert Into ������������ (����,����,��ʼ����,��ֹ����,�ϴ�IP,�ϴ��û�,�ϴ�����,�ϴ�Ŀ¼,����IP,�����û�,��������,����Ŀ¼,װǮIP��ַ1,װǮIP��ַ2) " & _
                                      " VALUES (" & TYPE_ͭ���� & ",'" & �������� & "'," & str��ʼ���� & "," & str��ֹ���� & _
                                      ",'" & varTemp(hp�ϴ�IP) & "','" & varTemp(hp�ϴ��û�) & "','" & varTemp(hp�ϴ�����) & "','" & varTemp(hp�ϴ�Ŀ¼) & _
                                          "','" & varTemp(hp����IP) & "','" & varTemp(hp�����û�) & "','" & varTemp(hp��������) & "','" & varTemp(hp����Ŀ¼) & _
                                          "','" & varTemp(hpװǮIP��ַ1) & "','" & varTemp(hpװǮIP��ַ2) & "')"
                            gcnҽ��.Execute gstrSQL
                        Else
                            '���ڴ����·�������Ҫ�ȼ���Ƿ����Ѿ������·���������Ѿ����ڣ��򲻴���
                            If col����("K" & varTemp(hp��������)) = False Then
                                gstrSQL = "Insert Into ������������ (����,����,��ʼ����,��ֹ����,�ϴ�IP,�ϴ��û�,�ϴ�����,�ϴ�Ŀ¼,����IP,�����û�,��������,����Ŀ¼,װǮIP��ַ1,װǮIP��ַ2) " & _
                                          " VALUES (" & TYPE_ͭ���� & ",'" & varTemp(hp��������) & "'," & str��ʼ���� & "," & str��ֹ���� & _
                                          ",'" & varTemp(hp�ϴ�IP) & "','" & varTemp(hp�ϴ��û�) & "','" & varTemp(hp�ϴ�����) & "','" & varTemp(hp�ϴ�Ŀ¼) & _
                                          "','" & varTemp(hp����IP) & "','" & varTemp(hp�����û�) & "','" & varTemp(hp��������) & "','" & varTemp(hp����Ŀ¼) & _
                                          "','" & varTemp(hpװǮIP��ַ1) & "','" & varTemp(hpװǮIP��ַ2) & "')"
                                gcnҽ��.Execute gstrSQL
                            End If
                        End If
                    Case "CENTERS"
                        '���ȼ��������Ƿ����
                        gstrSQL = "Select Rowid RID From ��������Ŀ¼ Where  ����=" & TYPE_ͭ���� & " and ����='" & varTemp(cf���Ĵ���) & "'"
                        Call OpenRecordset(rsTemp)
                        If rsTemp.RecordCount > 0 Then
                            '�Ѿ����ڣ�ֻ��Ҫ����
                            gstrSQL = "Update ��������Ŀ¼ Set ��������='" & varTemp(cf��������) & "',�����˻���֧�������Ը�=" & varTemp(cf�����˻���֧�������Ը�) & _
                                        ",��д����Ʊ=" & varTemp(cf��д����Ʊ) & " where RowID='" & rsTemp("RID") & "'"
                            gcnҽ��.Execute gstrSQL
                        Else
                            '���������ģ���Ҫȡ�������ţ�
                            lng��� = GetMax("��������Ŀ¼", "���", 1, " Where ����=" & TYPE_ͭ����)
                            gstrSQL = "Insert Into ��������Ŀ¼ (����,���,����,����,��������,����ģʽ,�����˻���֧�������Ը�,��д����Ʊ) values(" & _
                                TYPE_ͭ���� & "," & lng��� & ",'" & varTemp(cf���Ĵ���) & "','" & varTemp(cf��������) & "','" & varTemp(cf��������) & _
                                "',0," & varTemp(cf�����˻���֧�������Ը�) & "," & varTemp(cf��д����Ʊ) & ")"
                            gcnҽ��.Execute gstrSQL
                            
                            'ͬʱ����HIS�в���
                            gstrSQL = "Insert Into ��������Ŀ¼ (����,���,����,����) values(" & _
                                TYPE_ͭ���� & "," & lng��� & ",'" & varTemp(cf���Ĵ���) & "','" & varTemp(cf��������) & "')"
                            gcnOracle.Execute gstrSQL
                        End If
                    Case "HOSTAGENCY"
                        If varTemp(0) = �������� And varTemp(1) = mstrҽԺ���� Then
                            gstrSQL = "Update �������� Set װǮģʽ=" & varTemp(2) & " Where ����=" & TYPE_ͭ���� & " and ����='" & �������� & "'"
                            gcnҽ��.Execute gstrSQL
                        End If
                    Case "CENTERAGENCY"
                        If varTemp(1) = mstrҽԺ���� Then
                            '�������ô�λ���޼��⣬������ʹ�ø�����
                            gstrSQL = "Update ��������Ŀ¼ Set ����ģʽ=1,ÿ�촲λ���޼�=" & varTemp(2) & " Where ����=" & TYPE_ͭ���� & " and ����='" & varTemp(0) & "'"
                            gcnҽ��.Execute gstrSQL
                            gstrSQL = "Update �������� Set �Ƿ����=1 Where ����=" & TYPE_ͭ���� & " and ����='" & �������� & "'"
                            gcnҽ��.Execute gstrSQL
                        End If
                End Select
            End If
        End If
    Loop
    
    objText.Close  '�ر��ļ��������޷��õ���ѹ��һ��Center�ļ�
    gcnҽ��.CommitTrans
    DownHost = bln����
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    objText.Close
    gcnҽ��.RollbackTrans
End Function

Private Function DownװǮ(ByVal lng��� As Long, ByVal strҽ���� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant
    
    '���ر�
    If DownLoadFile("inmoneylist.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnҽ��.BeginTrans
    On Error GoTo errHandle

    '����ɾ����ǰҽ�����ĵ�����
    lbl��Ŀ.Caption = "װǮ�嵥"
    gcnҽ��.Execute "Delete from װǮ�嵥 where ���Ĵ��� IN (" & mstr����InOracle & ")"

    '�������ļ�
    Call OpenText(mstr��������Ŀ¼ & "inmoneylist", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr����InStr, "," & varFields(0)) > 0 Then
                '����װǮ�嵥,ע��˴��Խ������˼���
                gstrSQL = "insert into װǮ�嵥 (���Ĵ���,����,ҽ����,װǮ�ڴ�,�ʻ�ע��,�����·�) values ('" & varFields(0) & _
                    "','" & varFields(1) & "','" & strҽ���� & "'," & lng��� & ",'" & EncryptStr(varFields(2), "256", True) & "','" & varFields(3) & "')"
                gcnҽ��.Execute gstrSQL
            End If
        End If
    Loop

    '���²�����
    Call Update��������("װǮ���", lng���)

    gcnҽ��.CommitTrans
    DownװǮ = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
End Function

Private Function Down����(ByVal lng��� As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant
    
    '���ر�
    If DownLoadFile("levwork.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnҽ��.BeginTrans
    On Error GoTo errHandle

    '����ɾ����ǰҽ�����ĵ�����
    lbl��Ŀ.Caption = "������Ա"
    gcnҽ��.Execute "Delete from ������Ա where ���Ĵ��� IN (" & mstr����InOracle & ")"

    '�������ļ�
    Call OpenText(mstr��������Ŀ¼ & "levwork", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr����InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into ������Ա (���Ĵ���,����,����,�Ա�,ҽ����,���֤��,��λҽ����,��ݴ���,��λ����,�Ƿ�������ҵ) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "','" & varFields(3) & _
                     "','" & varFields(4) & "','" & varFields(5) & "','" & varFields(6) & "','" & varFields(7) & "','" & varFields(8) & "','" & varFields(9) & "')"
                gcnҽ��.Execute gstrSQL
            End If
        End If
    Loop

    '���²�����
    Call Update��������("���ݸɲ��������", lng���)

    gcnҽ��.CommitTrans
    Down���� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
End Function

Private Function Down��λ��Ϣ() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant

    rsTemp.CursorLocation = adUseClient

    gcnҽ��.BeginTrans
    On Error GoTo errHandle

    '����ɾ����ǰҽ�����ĵ�����
    lbl��Ŀ.Caption = "��λ��Ϣ"
    gcnҽ��.Execute "Delete from ��λ��Ϣ where ���Ĵ��� IN (" & mstr����InOracle & ")"

    '�������ļ�
    Call OpenText(mstr��������Ŀ¼ & "SPECRETIREPAYPARAMS", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr����InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into ��λ��Ϣ (���Ĵ���,����,��λ��������,ͳ�︺������,���˸�������,�Ƿ�������ҵ) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "','" & varFields(3) & _
                     "','" & varFields(4) & "','" & varFields(5) & "')"
                gcnҽ��.Execute gstrSQL
            End If
        End If
    Loop

    gcnҽ��.CommitTrans
    Down��λ��Ϣ = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
End Function

Private Function Down�����޶�() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant

    rsTemp.CursorLocation = adUseClient

    gcnҽ��.BeginTrans
    On Error GoTo errHandle

    '����ɾ����ǰҽ�����ĵ�����
    lbl��Ŀ.Caption = "��λ��Ϣ"
    gcnҽ��.Execute "Delete from �����޶� where ���ĵ�λ IN (" & mstr����InOracle & ")"

    '�������ļ�
    Call OpenText(mstr��������Ŀ¼ & "TENDLEVY", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr����InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into �����޶� (���ĵ�λ,����,����) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "')"
                gcnҽ��.Execute gstrSQL
            End If
        End If
    Loop

    gcnҽ��.CommitTrans
    Down�����޶� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
End Function

Private Function Down����(ByVal lng��� As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant
    
    '���ر�
    If DownLoadFile("experlist.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnҽ��.BeginTrans
    On Error GoTo errHandle

    '����ɾ����ǰҽ�����ĵ�����
    lbl��Ŀ.Caption = "������Ա"
    gcnҽ��.Execute "Delete from ������Ա where ���Ĵ��� IN (" & mstr����InOracle & ")"

    '�������ļ�
    Call OpenText(mstr��������Ŀ¼ & "experlist", objText, lngLines)

    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        SetProgress lngLines, objText.Line

        If strLine <> "" Then
            varFields = Split(strLine, "|")
            If InStr(mstr����InStr, "," & varFields(0)) > 0 Then
                gstrSQL = "insert into ������Ա (���Ĵ���,ְ������) values (" & _
                    "'" & varFields(0) & "','" & varFields(1) & "')"
                gcnҽ��.Execute gstrSQL
            End If
        End If
    Loop

    '���²�����
    Call Update��������("������Ա�������", lng���)

    gcnҽ��.CommitTrans
    Down���� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
End Function

Private Function Down������(ByVal lng��� As Long, ByVal strҽ���� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long
    
    '���ر�
    If DownLoadFile("cardblklist.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnҽ��.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("cardblklist", "unitblklist")
    For lngCount = LBound(varTables) To UBound(varTables)
        '����ɾ����ǰҽ�����ĵ�����
        If varTables(lngCount) = "cardblklist" Then
            lbl��Ŀ.Caption = "��������Ա"
            '�������Ĵ���������һ�㲻ͬ
            If strҽ���� > mstrԭҽ���� Then
                '��ҽ����Ĵ���
                gcnҽ��.Execute "Delete from ������ where ҽ����='" & mstrԭҽ���� & "' and �Ҷ�<>'1' " & _
                                " and ���Ĵ��� IN (" & mstr����InOracle & ")"
            End If
            gcnҽ��.Execute "Delete from ������ where ҽ����='" & strҽ���� & "' And ���Ĵ��� IN (" & mstr����InOracle & ")"
        Else
            lbl��Ŀ.Caption = "��������λ"
            gcnҽ��.Execute "Delete from ��λ������ where ���Ĵ��� IN (" & mstr����InOracle & ")"
        End If
            
        '�������ļ�
        Call OpenText(mstr��������Ŀ¼ & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                If InStr(mstr����InStr, "," & varFields(0)) > 0 Then
                    If varTables(lngCount) = "cardblklist" Then
                        gstrSQL = "insert into ������ (���Ĵ���,����,�Ҷ�,ҽ����) values ('" & varFields(0) & _
                            "','" & varFields(1) & "','" & varFields(2) & "','" & strҽ���� & "')"
                    Else
                        gstrSQL = "insert into ��λ������ (���Ĵ���,����,����,�Ҷ�) values ('" & varFields(0) & _
                            "','" & varFields(1) & "','" & varFields(2) & "','" & varFields(3) & "')"
                    End If
                    gcnҽ��.Execute gstrSQL
                End If
            End If
        Loop
    Next
    
    '���²�����
    Call Update��������("�������������", lng���)
    Call Update��������("ҽ����", strҽ����)

    gcnҽ��.CommitTrans
    Down������ = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
End Function

Private Function Down����(ByVal lng��� As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long, lng��� As Long
    
    '���ر�
    If DownLoadFile("sickdefine.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnOracle.BeginTrans
    gcnҽ��.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("sickdefine", "sickkind")
    For lngCount = LBound(varTables) To UBound(varTables)
        '����ɾ����ǰҽ�����ĵ�����
        If varTables(lngCount) = "sickdefine" Then
            lbl��Ŀ.Caption = "���ղ���"
            gcnҽ��.Execute "Delete from ���ղ��� where ����=" & TYPE_ͭ����
        Else
            lbl��Ŀ.Caption = "��������"
            gcnҽ��.Execute "Delete from ���ղ���֧�� where ���Ĵ��� IN (" & mstr����InOracle & ")"
        End If
            
        '�������ļ�
        Call OpenText(mstr��������Ŀ¼ & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                If varTables(lngCount) = "sickdefine" Then
                    '���ղ���
                    gstrSQL = "insert into ���ղ��� (����,����,����,����,���) values (" & _
                                TYPE_ͭ���� & ",'" & varFields(0) & "','" & _
                                Replace(varFields(3), "'", "''") & "','" & varFields(2) & "','" & varFields(4) & "')"
                    gcnҽ��.Execute gstrSQL
                    
                    'ͬʱ����HIS�Ĳ��֣�ע��HIS�Ĳ�������ֻ֧��3�֣����Խ�1-5����,6-9����Ϊ���ⲡ
                    gstrSQL = "select rowid as RID FROM ���ղ��� where ����=" & TYPE_ͭ���� & " and ����='" & varFields(0) & "'"
                    If Val(varFields(4)) >= 6 Then
                        lng��� = 2
                    ElseIf Val(varFields(4)) >= 1 Then
                        lng��� = 1
                    Else
                        lng��� = 0
                    End If
                    If rsTemp.State = adStateOpen Then rsTemp.Close
                    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                    If rsTemp.RecordCount = 0 Then
                        '�ò��ֲ����ڣ�����
                        gstrSQL = "insert into ���ղ��� (����,ID,����,����,����,���) values (" & _
                                    TYPE_ͭ���� & ",���ղ���_ID.nextval,'" & varFields(0) & "','" & _
                                    Replace(varFields(3), "'", "''") & "','" & varFields(2) & "','" & lng��� & "')"
                    Else
                        '�����Ѿ����ڣ��޸�
                        gstrSQL = "update ���ղ���  set ����='" & Replace(varFields(3), "'", "''") & "',����='" & varFields(2) & _
                            "',���='" & lng��� & "' where  rowid='" & rsTemp("RID") & "'"
                    End If
                    gcnOracle.Execute gstrSQL
                Else
                    '��������
                    If InStr(mstr����InStr, "," & varFields(0)) > 0 Then
                        gstrSQL = "INSERT INTO ���ղ���֧�� (���Ĵ���,�������ʹ���,������������,֧������,�޶�,�޶�����,�ۼƻ�������֧��,�ۼƻ������շ���,���ﱨ��Ӱ���޶�,����ʱ��Ӱ���޶�,סԺ����Ӱ���޶�,�����·�Ӱ���޶�,סԺӰ����,���߽��,�����ʻ�ʹ�÷���,�����㱨��,�����ʻ��㱨��,ͳ��ⶥӰ�챨��)  Values('" & _
                            varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "'," & varFields(3) & _
                            "," & varFields(4) & "," & varFields(5) & "," & varFields(6) & "," & varFields(7) & _
                            "," & varFields(8) & "," & varFields(9) & "," & varFields(10) & "," & varFields(11) & _
                            "," & varFields(12) & "," & varFields(13) & "," & varFields(14) & "," & varFields(15) & _
                            "," & varFields(16) & "," & varFields(17) & ") "
                        gcnҽ��.Execute gstrSQL
                    End If
                End If
            End If
        Loop
    Next
    
    '���²�����
    Call Update��������("�����������", lng���)

    gcnOracle.CommitTrans
    gcnҽ��.CommitTrans
    Down���� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
    gcnOracle.RollbackTrans
End Function

Private Function Down����(ByVal lng��� As Long, ByVal strҽ���� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long
    Dim lng���� As Long, lngְ����� As Long
    Dim cnServer(1 To 2) As ADODB.Connection, lngServer As Long
    
    Set cnServer(1) = gcnOracle
    Set cnServer(2) = gcnҽ��
    
    Dim rs���� As New ADODB.Recordset
    
    rs����.Fields.Append "����", adBigInt, 19, adFldIsNullable
    rs����.Fields.Append "��ְ", adBigInt, 19, adFldIsNullable
    rs����.Fields.Append "����", adSingle, 15, adFldIsNullable
    rs����.CursorLocation = adUseClient
    rs����.LockType = adLockOptimistic
    rs����.CursorType = adOpenStatic
    rs����.Open
    
    '���ر�
    If DownLoadFile("policy.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnOracle.BeginTrans
    gcnҽ��.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("payparams", "paypolicy", "subpayparams", "medikind", "conformation")
    For lngCount = LBound(varTables) To UBound(varTables)
        '����ɾ����ǰҽ�����ĵ�����
        Select Case varTables(lngCount)
            Case "paypolicy"
                lbl��Ŀ.Caption = "��������"
            Case "payparams"
                lbl��Ŀ.Caption = "���ղ���"
            Case "subpayparams"
                lbl��Ŀ.Caption = "��������"
                gcnҽ��.Execute "Delete from ���ղ������� where ����=" & TYPE_ͭ���� & " And ���=" & strҽ���� & _
                                " And ���� IN (" & mstr���InOracle & ")"
            Case "medikind"
                lbl��Ŀ.Caption = "���մ���"
                gcnҽ��.Execute "Delete from ����֧������ where ����=" & TYPE_ͭ����
            Case "conformation"
                lbl��Ŀ.Caption = "����"
                gcnҽ��.Execute "Delete from ����"
        End Select
            
        '�������ļ�
        Call OpenText(mstr��������Ŀ¼ & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                Select Case varTables(lngCount)
                    Case "payparams"
                        '�����ֶα���
                        'Modified by ZYB 20060617
                        '***ֻ���ر�ҽԺ�ȼ�������
                        If varFields(parҽԺ�ȼ�) = Mid(mstrҽԺ����, 1, 1) And varFields(par��������) = 3 Then '����ֻ�����һλ
                            For lngServer = 1 To 2
                                If rsTemp.State = adStateOpen Then rsTemp.Close
                                rsTemp.Open "select ��� from ��������Ŀ¼ where ����=" & TYPE_ͭ���� & " and ����='" & varFields(par���Ĵ���) & "'", cnServer(lngServer), adOpenStatic, adLockReadOnly
                                If rsTemp.RecordCount = 1 Then
                                    '���ڸ�����
                                    lng���� = rsTemp("���")
                                    'Modified by ZYB 20060617
                                    '1*=��ְ;2*=����;����=����
                                    '�·����ļ���,0��ʾ��ְ,1��ʾ����,û���·����ݵ�
                                    lngְ����� = Switch(Left(varFields(parְ�����), 1) = "0", "1", Left(varFields(parְ�����), 1) = "1", 2, True, 3)
                                    
                                    '�������ߣ��Թ�����ҽ������ʹ��
                                    If lngServer = 2 Then 'ֻ����ҽ��������
                                        rs����.AddNew
                                        rs����("����") = lng����
                                        rs����("��ְ") = lngְ�����
                                        rs����("����") = Val(varFields(par����))
                                        rs����.Update
                                    End If
                                     
                                     '��Ȼ���õ����ǲ��ģ���Ҳÿ�ζ�����
                                    strLine = "1;��һ��;0;" & Format(Val(varFields(par�ڶ�����ʼֵ)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "2;�ڶ���;" & Format(Val(varFields(par�ڶ�����ʼֵ)), "########0.00;-########0.00; ; ") & ";" & Format(Val(varFields(par��������ʼֵ)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "3;������;" & Format(Val(varFields(par��������ʼֵ)), "########0.00;-########0.00; ; ") & ";" & Format(Val(varFields(par���Ķ���ʼֵ)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "4;���ĵ�;" & Format(Val(varFields(par���Ķ���ʼֵ)), "########0.00;-########0.00; ; ") & ";" & Format(Val(varFields(par�������ʼֵ)), "########0.00;-########0.00; ; ") & ";"
                                    strLine = strLine & "5;���嵵;" & Format(Val(varFields(par�������ʼֵ)), "########0.00;-########0.00; ; ") & ";0;"
                                    gstrSQL = "zl_���շ��õ�_Update(" & TYPE_ͭ���� & "," & lng���� & ",'" & strLine & "')"
                                    cnServer(lngServer).Execute gstrSQL, , adCmdStoredProc
                                    
                                    '�����
                                    If lngְ����� = 3 Then
                                        gstrSQL = "zl_���������_Update(" & TYPE_ͭ���� & "," & lng���� & ",3,1,1,1,'1;����;0;0;')"
                                    Else
                                        gstrSQL = "zl_���������_Update(" & TYPE_ͭ���� & "," & lng���� & "," & lngְ����� & ",0,0,0,'1;" & IIf(lngְ����� = 1, "��ְ", "����") & ";0;0;')"
                                    End If
                                    cnServer(lngServer).Execute gstrSQL, , adCmdStoredProc
                                    
                                    '����֧������
                                    gstrSQL = "Delete ����֧������ WHERE ����=" & TYPE_ͭ���� & " AND ����=" & lng���� & " AND ���=" & strҽ���� & " and ��ְ=" & lngְ�����
                                    With cnServer(lngServer)
                                        .Execute gstrSQL
                                        .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                            TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & "," & lngְ����� & ",1,1," & Val(varFields(par��һ�α�������)) * 100 & ")"
                                        .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                            TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & "," & lngְ����� & ",1,2," & Val(varFields(par�ڶ��α�������)) * 100 & ")"
                                        .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                            TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & "," & lngְ����� & ",1,3," & Val(varFields(par�����α�������)) * 100 & ")"
                                        .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                            TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & "," & lngְ����� & ",1,4," & Val(varFields(par���Ķα�������)) * 100 & ")"
                                        .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                            TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & "," & lngְ����� & ",1,5," & Val(varFields(par����α�������)) * 100 & ")"
                                
                                        If lngְ����� = 1 Then
                                            'ǿ�д������ݲ��ˣ����������ļ���û�����ݲ��˵����أ�
                                             gstrSQL = "zl_���������_Update(" & TYPE_ͭ���� & "," & lng���� & ",3,1,1,1,'1;����;0;0;')"
                                             .Execute gstrSQL
                                             '����֧������
                                             gstrSQL = "Delete ����֧������ WHERE ����=" & TYPE_ͭ���� & " AND ����=" & lng���� & " AND ���=" & strҽ���� & " and ��ְ=3"
                                             .Execute gstrSQL
                                             .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                                     TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & ",3,1,1,100)"
                                             .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                                     TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & ",3,1,2,100)"
                                             .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                                     TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & ",3,1,3,100)"
                                             .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                                     TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & ",3,1,4,100)"
                                             .Execute "INSERT INTO ����֧������(����,����,���,��ְ,�����,����,����) values(" & _
                                                     TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & ",3,1,5,100)"
                                        
                                        End If
                                    End With
                                End If
                            Next
                        End If
                    Case "paypolicy"
                        '��������
                        Dim cur���� As Double, curʵ������ As Double, cur������ As Double
                        Dim lng�������� As Long, lngסԺ���� As Long
                        
                        '���ɱ���֧���޶��
                        strLine = ""
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "select ��� from ��������Ŀ¼ where ����=" & TYPE_ͭ���� & " and ����='" & varFields(pol���Ĵ���) & "'", gcnҽ��, adOpenStatic, adLockReadOnly
                        If rsTemp.RecordCount = 1 Then
                            '���ڸ�����
                            lng���� = rsTemp("���")
                            
                            rs����.Filter = "����=" & lng����
                            Do Until rs����.EOF
                                cur���� = rs����("����")
                                strLine = strLine & rs����("��ְ") & ";" & "A;" & varFields(polͳ��ⶥ��) & ";" 'ͳ��ⶥ��
                                '�������סԺ6��
                                For lngסԺ���� = 0 To 5
                                    '���ȵõ���Ч��סԺ����
                                    If lngסԺ���� > (Val(varFields(pol��������)) - 1) And Val(varFields(pol��������)) > 0 Then
                                        '********���ô�������
                                        '���ֻ�����⼸��
                                        lng�������� = Val(varFields(pol��������)) - 1
                                    Else
                                        '�������ƿ���Ϊ-1
                                        '��һ��סԺ����ֵΪ0
                                        lng�������� = lngסԺ����
                                    End If
                                    
                                    If varFields(pol�㷨) = "-" Then
                                        '�ݼ��㷨
                                        cur������ = Val(varFields(pol��ֵ)) * lng��������
                                    Else
                                        '����������
                                        cur������ = cur���� * Val(varFields(pol��ֵ)) * lng��������
                                    End If
                                    
                                    If cur������ > Val(varFields(pol����������)) And Val(varFields(pol����������)) > 0 Then
                                        '********���ü���������
                                        '���������ƿ���Ϊ-1
                                        cur������ = Val(varFields(pol����������))
                                    End If
                                    
                                    curʵ������ = cur���� - cur������
                                    
                                    If curʵ������ < Val(varFields(pol��������)) And Val(varFields(pol��������)) > 0 Then
                                        '********������������
                                        '�������ƿ���Ϊ-1
                                        curʵ������ = Val(varFields(pol��������))
                                    End If
                                    
                                    strLine = strLine & rs����("��ְ") & ";" & (lngסԺ���� + 1) & ";" & curʵ������ & ";"
                                Next
                                rs����.MoveNext
                            Loop
                            gstrSQL = "zl_����֧���޶�_Update(" & TYPE_ͭ���� & "," & lng���� & "," & strҽ���� & ",'" & strLine & "')"
                            gcnҽ��.Execute gstrSQL, , adCmdStoredProc
                            
                            For lngServer = 1 To 2
                                gstrSQL = "Delete ���շ��õ� Where ����=" & TYPE_ͭ���� & " and ����=" & lng���� & " And ����>" & varFields(polͳ�����)
                                cnServer(lngServer).Execute gstrSQL
                                
                                gstrSQL = "Update ���շ��õ� Set ����=0 Where ����=" & TYPE_ͭ���� & " and ����=" & lng���� & " And ����=" & varFields(polͳ�����)
                                cnServer(lngServer).Execute gstrSQL
                                
'                                gstrSQL = "Delete ����֧������ Where ����=" & TYPE_ͭ���� & " and ����=" & lng���� & " And ���=" & strҽ���� & " And ����>" & varFields(polͳ�����)
'                                cnServer(lngServer).Execute gstrSQL, , adCmdStoredProc
'
                            Next
                        End If
                        gstrSQL = "Update ��������Ŀ¼ Set " & _
                                  "  �����ڶ���=" & varFields(pol�����ڶ���) & "," & _
                                  "  ��ֵ����=" & varFields(pol��ֵ����) & "," & _
                                  "  �ⶥ����=" & varFields(pol�ⶥ����) & "," & _
                                  "  ʹ���ۼƱ���=" & varFields(polʹ���ۼƱ���) & "," & _
                                  "  �������Բ��·�=" & varFields(pol�������Բ��·�) & "," & _
                                  "  ��չ���䱣�ձ���=" & varFields(pol��չ���䱣�ձ���) & "," & _
                                  "  ���䱨������=" & varFields(pol���䱨������) & "," & _
                                  "  ���䱨���޶�=" & varFields(pol���䱨���޶�) & "," & _
                                  "  ���䱨���޶�����=" & varFields(pol���䱨���޶�����) & "," & _
                                  "  ���䱨�����𸶽�=" & varFields(pol���䱨�����𸶽�) & "," & _
                                  "  ��չ��������=" & varFields(pol��չ��������) & "," & _
                                  "  ��չ��������=" & varFields(pol��չ��������) & "," & _
                                  "  ��չ�󲡱���=" & varFields(pol��չ�󲡱���) & "," & _
                                  "  �����𸶽�����=" & varFields(pol�����𸶽�����) & "," & _
                                  "  ��������סԺ����=" & varFields(pol��������סԺ����) & "," & _
                                  "  ������Ŀ�۸�=" & Val(varFields(pol������Ŀ�۸�)) & "" & _
                                  " Where ����=" & TYPE_ͭ���� & " And ����='" & varFields(0) & "'"
                        gcnҽ��.Execute gstrSQL
                    Case "subpayparams"
                        '������������
                        gstrSQL = "insert into ���ղ�������(����,����,���,��ֵ,����)" & _
                                  " SELECT ����,���," & strҽ���� & " AS ���," & varFields(1) & "," & varFields(2) & _
                                  " FROM ��������Ŀ¼ WHERE ����=" & TYPE_ͭ���� & " And ����='" & varFields(0) & "'"
                        gcnҽ��.Execute gstrSQL
                    Case "medikind"
                        'ҽ������
                        gstrSQL = "insert into ����֧������ (����,����,����,��ҽ������) values (" & _
                                    TYPE_ͭ���� & ",'" & varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "')"
                        gcnҽ��.Execute gstrSQL
                        
                        'ͬʱ����HIS�Ĳ��֣�ע��HIS�Ĳ�������ֻ֧��3�֣����Խ�2-9����Ϊ���ⲡ
                        gstrSQL = "select rowid as RID FROM ����֧������ where ����=" & TYPE_ͭ���� & " and ����='" & varFields(0) & "'"
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                        If rsTemp.RecordCount = 0 Then
                            '�ñ���֧�����಻���ڣ�����
                            gstrSQL = "insert into ����֧������ (����,ID,����,����,����,����,�㷨,ͳ��ȶ�,�Ƿ�ҽ��,�������) values (" & _
                                        TYPE_ͭ���� & ",����֧������_ID.nextval,'" & varFields(0) & "','" & _
                                        varFields(1) & "','',1,1,0,1,3)"
                            gcnOracle.Execute gstrSQL
                        End If
                    Case "conformation"
                        '����
                        gstrSQL = "INSERT INTO ���� (����,����) VALUES ('" & _
                                varFields(0) & "','" & Replace(varFields(1), "'", "") & "')"
                        gcnҽ��.Execute gstrSQL
                End Select
            End If
        Loop
    Next
    
    '���²�����
    Call Update��������("�����������", lng���)

    gcnOracle.CommitTrans
    gcnҽ��.CommitTrans
    Down���� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
    gcnOracle.RollbackTrans
End Function

Private Function Down��Ŀ(ByVal lng��� As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strLine As String, lngLines As Long, objText As TextStream
    Dim varFields As Variant, varTables As Variant, lngCount As Long
    Dim lng�Ƿ�ҽ�� As Long, str������� As String, rs���մ��� As New ADODB.Recordset
    
    '���ر�
    If DownLoadFile("itemcenter.pak") = False Then
        Exit Function
    End If

    rsTemp.CursorLocation = adUseClient

    gcnOracle.BeginTrans
    gcnҽ��.BeginTrans
    On Error GoTo errHandle
    
    varTables = Array("itemcenter", "agencyspecitem", "sickspecitem")
    For lngCount = LBound(varTables) To UBound(varTables)
        '����ɾ����ǰҽ�����ĵ�����
        Select Case varTables(lngCount)
            Case "itemcenter"
                lbl��Ŀ.Caption = "������Ŀ"
                gcnҽ��.Execute "Delete from ������Ŀ where ����=" & TYPE_ͭ����
'                gcnOracle.Execute "Delete from ������Ŀ where ����=" & TYPE_ͭ����

                '�õ���ҽ������ı���
                rs���մ���.Open "Select ����,��ҽ������ From ����֧������", gcnҽ��, adOpenStatic, adLockReadOnly
            Case "agencyspecitem"
                lbl��Ŀ.Caption = "ҽԺ������Ŀ"
            Case "sickspecitem"
                lbl��Ŀ.Caption = "������׼��Ŀ"
                gcnҽ��.Execute "Delete from ���ղ�����Ŀ"
        End Select
            
        '�������ļ�
        Call OpenText(mstr��������Ŀ¼ & varTables(lngCount), objText, lngLines)
    
        Do Until objText.AtEndOfStream
            strLine = Trim(objText.ReadLine)
            SetProgress lngLines, objText.Line
    
            If strLine <> "" Then
                varFields = Split(strLine, "|")
                Select Case varTables(lngCount)
                    Case "itemcenter"
                        '������Ŀ
                            '����ҽ�������жϸ���Ŀ�Ƿ�ҽ�������صĵȼ�����ҽԺ�ȼ���˵����Ժ������������������Ϊҽ����Ŀ
                            lng�Ƿ�ҽ�� = IIf(varFields(if�Ƿ�ҽ��) > Mid(mstrҽԺ����, 1, 1), 0, 1)
                            str������� = varFields(if�������)
                            If lng�Ƿ�ҽ�� = 0 Then
                                '����Ƿ�ҽ����Ŀ����һ������
                                rs���մ���.Filter = "����='" & str������� & "'"
                                If rs���մ���.EOF = False Then
                                    str������� = NVL(rs���մ���("��ҽ������"), str�������)
                                End If
                            End If
                            
                            gstrSQL = "INSERT INTO ������Ŀ (����,����,����,����,��λ,���ͱ���,�������,�Ƿ���ҩ,�Ƿ�ҽ��,���۸�����," & _
                                      "�����Ը�����,�۸�,��Ŀ�ں�,��������,˵��,ʡ���޼�,�м��޼�,�ؼ��޼�,�缶�޼�,�ؼ���Ŀ,�ؼ��Ը�����) VALUES ( " & _
                                      TYPE_ͭ���� & ",'" & varFields(if��Ŀ���) & "','" & Replace(varFields(ifҩ������), "'", "") & "','" & Replace(varFields(ifƴ������), "'", "") & _
                                      "','" & varFields(if��λ) & "','" & varFields(if���ͱ���) & "','" & str������� & "','" & varFields(if�Ƿ���ҩ) & _
                                      "','" & lng�Ƿ�ҽ�� & "','" & varFields(if���۸�����) & "','" & varFields(if�����Ը�����) & "','" & varFields(if�۸�) & _
                                      "','" & varFields(if��Ŀ�ں�) & "','" & varFields(if��������) & "','" & varFields(if˵��) & _
                                      "','" & varFields(ifʡ���޼�) & "','" & varFields(if�м��޼�) & "','" & varFields(if�ؼ��޼�) & "','" & varFields(if�缶�޼�) & "','" & varFields(if�ؼ���Ŀ) & "','" & varFields(if�ؼ��Ը�����) & "')"
                            gcnҽ��.Execute gstrSQL
                            
'                            gstrSQL = "INSERT INTO ������Ŀ (����,����,����,����,������� VALUES ( " & _
'                                            TYPE_ͭ���� & ",'" & varFields(if��Ŀ���) & "','" & varFields(ifҩ������) & "','" & MidUni(varFields(ifƴ������), 1, 10) & _
'                                            "','" & varFields(if�������) & "')"
'                            gcnOracle.Execute gstrSQL
                            
                            '�������ڱ���֧����Ŀ���Ƿ�ҽ������ע
                            gstrSQL = "update ����֧����Ŀ A " & _
                                      "  set A.��Ŀ����='" & varFields(ifҩ������) & "',A.�Ƿ�ҽ��=" & lng�Ƿ�ҽ�� & _
                                      "  where A.��Ŀ����='" & varFields(if��Ŀ���) & "' and A.����=" & TYPE_ͭ����
                            
                            gcnOracle.Execute gstrSQL
                    Case "agencyspecitem"
                        'ҽԺ������Ŀ
                        If varFields(0) = mstrҽԺ���� Then
                            lng�Ƿ�ҽ�� = IIf(varFields(2) = 1, 0, 1)
                            
                            gstrSQL = "update ������Ŀ A " & _
                                      "  set A.�����Ը�����=" & varFields(2) & ",A.�Ƿ�ҽ��=" & lng�Ƿ�ҽ�� & _
                                      "  where A.��Ŀ����='" & varFields(1) & "' and A.����=" & TYPE_ͭ����
                            gcnҽ��.Execute gstrSQL
                            
                            gstrSQL = "update ����֧����Ŀ A " & _
                                      "  set A.�Ƿ�ҽ��=" & lng�Ƿ�ҽ�� & _
                                      "  where A.��Ŀ����='" & varFields(if��Ŀ���) & "' and A.����=" & TYPE_ͭ����
                            gcnOracle.Execute gstrSQL
                        End If
                    Case "sickspecitem"
                        '��������
                        gstrSQL = "INSERT INTO ���ղ�����Ŀ (�������,��Ŀ���,�����Ը�����) VALUES ('" & _
                                varFields(0) & "','" & varFields(1) & "','" & varFields(2) & "')"
                        gcnҽ��.Execute gstrSQL
                End Select
            End If
        Loop
    Next
    
    '����ҽԺ�ĵȼ��������۸��޼�
    Select Case mstrҽԺ����
    Case "33"
        gstrSQL = "Update ������Ŀ Set ���۸�����=��������޼�"
    Case "32"
        gstrSQL = "Update ������Ŀ Set ���۸�����=��������޼�"
    Case "23"
        gstrSQL = "Update ������Ŀ Set ���۸�����=��������޼�"
    Case "22"
        gstrSQL = "Update ������Ŀ Set ���۸�����=��������޼�"
    Case "13", "12"
        gstrSQL = "Update ������Ŀ Set ���۸�����=һ������޼�"
    End Select
    '���²�����
    Call Update��������("��Ŀ�������", lng���)

    gcnOracle.CommitTrans
    gcnҽ��.CommitTrans
    Down��Ŀ = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
    gcnҽ��.RollbackTrans
    gcnOracle.RollbackTrans
End Function

Private Function DownLoadFile(ByVal strFile As String) As Boolean
'���ܣ�����ָ�����ļ���������ɽ�ѹ������
    Dim zipfilesIn As ZIPnames
    Dim zipfilesEx As ZIPnames
    Dim lngReturn As Long
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    '�����ļ�
    lngReturn = FTPDownLoad(mstr����IP, "21", mstr�����û�, mstr��������, mstrԶ������Ŀ¼, strFile, mstr��������Ŀ¼ & strFile)
    If lngReturn <> 0 Then
        MsgBox "���ڡ�" & mstr�������� & "�����ļ�" & strFile & "����ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ļ����н���
    strTemp = mstr��������Ŀ¼ & Mid(strFile, 1, Len(strFile) - 4) & ".zip"
    DecryptFiles mstr��������Ŀ¼ & strFile, strTemp
    
    '��ѹ�ļ�
    If VBUnzip(strTemp, mstr��������Ŀ¼, 1, 1, 0, 0, 0, 0, zipfilesIn, zipfilesEx) = False Then
        Exit Function
    End If
    
    DownLoadFile = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub �ϴ�����(Optional ByVal bln�ָ� As Boolean = False)
    Dim rsHost As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim str�������� As String
    Dim dat�������� As Date, dat��ʼ���� As Date   '��ѯ����ʹ��   ��ʼ����<=�Ǽ�����<��������
    Dim datBegin As Date, datEnd As Date           '��ʱ��������Ҫ���ڷֶδ���
    Dim bln��Ҫ�ϴ� As Boolean
    Dim str�ϴ��ļ�  As String
    
    If GetҽԺ���� = False Then Exit Sub
    rsHost.Open "select * from �������� where ����=" & TYPE_ͭ���� & " And ���� = '" & Mid(tabHost.SelectedItem.Key, 2) & "'", gcnҽ��, adOpenStatic, adLockReadOnly
    
    'ȱʡ���ϴ����������
    dat�������� = CDate(Format(DateAdd("d", 1, Currentdate), "yyyy-MM-dd"))
    If Me.cbo�籣��.ListCount <> 0 Then str�������� = Mid(Me.cbo�籣��.Text, 1, 4)
    
    On Error GoTo errHandle
    '���ڿ����ж��ҽ�����ģ������һ��ѭ������
    Do Until rsHost.EOF
        '��ò���
        Call Get��������(rsHost)
        
        '������������ļ��Ĳ���
        mstr�������� = rsHost("����")
        mstr�������� = rsHost("����")
        
        Call Get�����б�(rsHost("����"))
        
        If bln�ָ� = False Then
            '1���õ����һ���ϴ������
            gstrSQL = "select max(��������) as �ϴ� from �ϴ����� where ����ģʽ=0 and ���Ĵ���='" & mstr�������� & "'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
            
            
            If IsNull(rsTemp("�ϴ�")) = True Then
                'δ���й��κδ�����һ���൱С��ֵ��Ϊ��ʼ
                dat��ʼ���� = CDate("1900-01-01")
                bln��Ҫ�ϴ� = True
            Else
                If rsTemp("�ϴ�") < dat�������� Then
                    '��Ҫ�ϴ����������
                    dat��ʼ���� = rsTemp("�ϴ�")
                    bln��Ҫ�ϴ� = True
                Else
                    '����Ĺ����Ѿ����У�ʲôҲ����Ҫ��
                    MsgBox mstr�������� & "�����ݽ����Ѿ��ϴ���ɣ������ٴ���", vbInformation, gstrSysName
                    bln��Ҫ�ϴ� = False
                End If
            End If
        Else
            '���лָ����ϴ�
            bln��Ҫ�ϴ� = True
            dat��ʼ���� = mdat��ʼ���� + 1
            dat�������� = mdat�������� + 1
        End If
        
        If bln��Ҫ�ϴ� = True Then
            '���Ȳ�������
            If dat��ʼ���� = CDate("1900-01-01") Then
                '��δ���й��ϴ���һ���Դ���
                datBegin = dat��ʼ����
                datEnd = dat��������
            Else
                datBegin = dat��ʼ����
                datEnd = dat��ʼ���� + 1
            End If
            
            Do Until datEnd > dat��������
                mstr��ʼ���� = "to_date('" & Format(datBegin, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                mstr�������� = "to_date('" & Format(datEnd, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                mstr�ս����� = "to_date('" & Format(datEnd - 1, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                mstrȱʡ��ʼ���� = "to_date('" & Format(DateAdd("d", -15, datEnd), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                
                gcnOracle.BeginTrans
                gcnҽ��.BeginTrans
                '����׼���ϴ�������
                If �ս�(bln�ָ�) = False Then
                    '�ս�ʧ��
                    gcnOracle.RollbackTrans
                    gcnҽ��.RollbackTrans
                    Exit Sub
                End If
                
                If bln�ָ� = False Then
                    '��¼�ϴ���־
                    gstrSQL = "insert into �ϴ����� (��������,�û���,����ģʽ,���Ĵ���,�ļ���) " & _
                              "values(" & mstr�������� & ",substr(user,1,20),'0','" & mstr�������� & "','" & mstrҽԺ���� & Format(datEnd - 1, "yyMMdd") & ".pak')"
                    gcnҽ��.Execute gstrSQL
                End If
                
                'Ȼ��ϳ��ļ����ϴ�����
'                If UpLoadFile(mstr�������� & mstrҽԺ���� & Format(datEnd - 1, "yyMMdd") & ".pak") = True Then
                If UpLoadFile(str�������� & mstrҽԺ���� & Format(datEnd - 1, "yyMMdd") & ".pak") = True Then
                    '�Ե�ǰҽ�����ĵ����ݽ����ύ
                    gcnOracle.CommitTrans
                    gcnҽ��.CommitTrans
                Else
                    gcnOracle.RollbackTrans
                    gcnҽ��.RollbackTrans
                End If
                '����ֶδ���
                datBegin = datBegin + 1
                datEnd = datEnd + 1
            Loop
        End If '�����ϴ�
        rsHost.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Function �ս�(Optional ByVal bln�ָ� As Boolean = False) As Boolean
'���ܣ�����������ϴ�����
    Dim rsTemp As New ADODB.Recordset, str��� As String
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency, curͳ�� As Currency
    
    On Error GoTo errHandle
    
    '�ϴ�֮ǰ������õı�����������ȷ��
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.�շ�ϸĿID,C.��Ŀ����,B.����,B.����,A.ʵ�ս�� " & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as ���� " & _
              "  From ���ս����¼ D,���˷��ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              "  where D.���� = 1 And D.���� =" & TYPE_ͭ���� & "  And D.��¼ID = A.����ID And A.�Ǽ�ʱ�� >=" & mstr��ʼ���� & " And A.�Ǽ�ʱ�� <" & mstr�������� & _
              "         AND A.ʵ�ս�� IS NOT NULL and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.���ӱ�־,0)<>9 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= " & TYPE_ͭ���� & _
              "  Order by A.����ID,A.����ʱ��"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        If Calc���÷ָ�(rsTemp, curȫ�Է�, cur�����Ը�, curͳ��) = False Then
            Exit Function
        End If
    End If
        
    rsTemp.Close
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.�շ�ϸĿID,C.��Ŀ����,B.����,B.����,A.ʵ�ս�� " & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as ���� " & _
              "  From ������ҳ D,���˷��ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              "  where D.����ID =A.����ID And D.��ҳID=A.��ҳID And D.���� =" & TYPE_ͭ���� & " And A.�Ǽ�ʱ�� >=" & mstrȱʡ��ʼ���� & "  And A.�Ǽ�ʱ�� <" & mstr�������� & _
              "        AND A.ʵ�ս�� IS NOT NULL and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.���ӱ�־,0)<>9 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= " & TYPE_ͭ���� & _
              "  Order by A.����ID,A.����ʱ��"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        If Calc���÷ָ�(rsTemp, curȫ�Է�, cur�����Ը�, curͳ��) = False Then
            Exit Function
        End If
    End If
    
    gstrSQL = "SELECT ҽ���� FROM �������� WHERE ����=" & TYPE_ͭ���� & " AND ����='" & mstr�������� & "'"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    str��� = rsTemp("ҽ����")
    
    '1����������
    gstrSQL = "select ���Ĵ���,'" & mstrҽԺ���� & "' as ҽԺ����,���,��Ʊ��,����,�Ա�,���� " & _
              "         ,����,ҽ����,���֤��,��λҽ����,��ݴ���,�Ƿ���Ա,�Ƿ�ҽ���չ˶���,�μӲ��䱣��,�ʻ��ۼ�����,�ʻ��ۼ�֧�� " & _
              "         ,ͳ����֧�����,ͳ����֧������,������֧�����,������֧������,�����𸶽���֧��,�������� " & _
              "         ,��������ʻ�֧�����,סԺ�����ʻ�֧�����,�����֧�����,��������,ҽ������,���ִ���,��������,�������� " & _
              "         ,�������ý��,ȫ�Ը����,�����Ը����,�����ʻ�֧��,ͳ����֧��,ͳ�����Ը�,ͳ�����֧��,ͳ������Ը� " & _
              "         ,�������֧��,��������Ը�,�������ⶥ��,������ⶥ��,����ͳ��֧��,����ͳ�����,��������֧��,������������ " & _
              "         ,����סԺ����,������֧��,������������ʻ�֧��,�������Բ��𸶽�,����סԺ�����ʻ�֧�� " & _
              "         ,�ʻ��ۼ�����,�����ʻ�֧��,֧��˳���,���Ҷȼ�,��Ʊ��־,����Ʊ�ݺ�,Ʊ������,A.���," & mstr�ս����� & " as �ս����� " & _
              "  from ���ս����¼ A " & _
              "  Where A.���� = 1 And A.���� =" & TYPE_ͭ���� & " And A.Ʊ������ <" & mstr�������� & _
              IIf(bln�ָ�, " And A.Ʊ������ >= " & mstr��ʼ����, " And Nvl(A.�Ƿ��ϴ�,0)=0 And A.Ʊ������ >=" & mstrȱʡ��ʼ����) & _
              "       and A.���Ĵ��� in (" & mstr����InOracle & ")"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UClinicBill", "��������")
    
    '2��������ϸ
    gstrSQL = "select '" & mstrҽԺ���� & "' as ҽԺ����,B.No||decode(B.��¼״̬,2,'2','1') as ���,substr(E.����,1,8) as ����ҩ�����,trim(substr(E.����,1,40)) as ����ҩ������ " & _
              "         ,trim(substr(G.����,1,40)) as ҽԺҩ������,Round(B.���ʽ��/(B.����*B.����),4) as ʵ�ʼ۸�,B.����*B.����*decode(B.��¼״̬,2,-1,1) as ����,B.���ʽ��*decode(B.��¼״̬,2,-1,1) " & _
              "         ,B.������Ŀ��,E.�������,E.�����Ը�����,E.���ͱ���,nvl(substr(decode(Instr(G.���,'��'),0,G.���,substr(G.���,1,Instr(G.���,'��')-1)),1,40),' ') as ��� " & _
              "  from ���ս����¼ A," & gstrOwner & ".���˷��ü�¼ B," & gstrOwner & ".�����ʻ� C,������Ŀ E," & gstrOwner & ".�շ�ϸĿ G " & _
              "  Where A.���� = 1 And A.���� =" & TYPE_ͭ���� & " And A.��¼ID = B.����ID And Nvl(B.���ӱ�־,0)<>9 " & " And A.Ʊ������ <" & mstr�������� & _
              IIf(bln�ָ�, " And A.Ʊ������ >= " & mstr��ʼ����, " And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(B.ʵ�ս��,0)<>0 And A.Ʊ������ >=" & mstrȱʡ��ʼ����) & _
              "       and B.���ձ���=E.���� and C.����=E.���� and B.�շ�ϸĿID=G.ID and A.����ID=C.����ID and C.����=" & TYPE_ͭ���� & " and C.���� IN (" & mstr���InOracle & ")"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UClinicBillDetail", "������ϸ")
    
    '3���������
    gstrSQL = "select '" & mstrҽԺ���� & "' as ҽԺ����,B.No||decode(B.��¼״̬,2,'2','1') as ���,E.�������,sum(B.���ʽ��)*decode(B.��¼״̬,2,-1,1) as ��� " & _
              "  from ���ս����¼ A," & gstrOwner & ".���˷��ü�¼ B," & gstrOwner & ".�����ʻ� C,������Ŀ E " & _
              "  Where A.���� = 1 And A.���� =" & TYPE_ͭ���� & " And A.��¼ID = B.����ID And Nvl(B.���ӱ�־,0)<>9 And A.Ʊ������ <" & mstr�������� & _
              IIf(bln�ָ�, " And A.Ʊ������ >= " & mstr��ʼ����, " And Nvl(A.�Ƿ��ϴ�,0)=0 And A.Ʊ������ >=" & mstrȱʡ��ʼ����) & _
              "       and A.����ID=C.����ID and C.����=" & TYPE_ͭ���� & " and B.���ձ���=E.���� and C.����=E.���� " & _
              " and C.���� IN (" & mstr���InOracle & ")" & _
              "  group by B.No,B.��¼״̬,E.�������"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UClinicMediKind", "�������")
    
    '4����Ժ�Ǽ�
    '������˵�����Ժ���ֵ����Ժ��������ȥ�������ܡ������ʻ����ġ�����ID��Ϊ�ա�
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    gstrSQL = "select substr(E.����,1,4) as ���Ĵ���,'" & mstrҽԺ���� & "' as ҽԺ����,A.����ID||'_'||A.��ҳID as ���,D.סԺ�� " & _
              "         ,substr(D.����,1,8) as ����,substr(D.�Ա�,1,2) as �Ա�,floor(MONTHS_BETWEEN(A.��Ժ����,D.��������)/12) as ���� " & _
              "         ,substr(C.����,1,8) as ����,substr(C.ҽ����,1,8)  as ҽ����,D.���֤��,substr(C.��λ����,1,5) as ��λҽ���� " & _
              "         ,substr(C.��Ա���,1,2) as ��Ա���,F.����,substr(A.�Ǽ���,1,8) as ҽ��,nvl(substr(G.����,1,50),'�޲���')  as ��Ժ���� " & _
              "         ,trunc(A.��Ժ����)," & str��� & " as ���," & mstr�ս����� & " as �ս����� " & _
              "  from ������ҳ A,�����ʻ� C,������Ϣ D,��������Ŀ¼ E,���ű� F,���ղ��� G " & _
              "  Where A.���� =" & TYPE_ͭ���� & " And A.�Ǽ�ʱ�� <" & mstr�������� & " And A.��Ժ����ID = F.ID And A.��Ժ���� Is Not Null " & _
              "       and A.����ID=C.����ID and C.����=" & TYPE_ͭ���� & " and A.����ID=D.����ID and C.����=E.���� and C.����=E.��� and C.����ID=G.ID(+) " & _
              " and E.��� IN (" & mstr���InOracle & ")" & IIf(bln�ָ�, " And A.�Ǽ�ʱ�� >= " & mstr��ʼ����, " And A.�Ǽ�ʱ�� >=" & mstrȱʡ��ʼ���� & " and nvl(A.�Ƿ��ϴ�,0)=0")
    '������Ǽ�(��ͬ�ļ�¼ֻȡһ��)
    gstrSQL = gstrSQL & vbCrLf & " Union " & vbCrLf & _
              "select substr(E.����,1,4) as ���Ĵ���,'" & mstrҽԺ���� & "' as ҽԺ����,A.����ID||'_'||A.��ҳID as ���,D.סԺ�� " & _
              "         ,substr(D.����,1,8) as ����,substr(D.�Ա�,1,2) as �Ա�,floor(MONTHS_BETWEEN(A.��Ժ����,D.��������)/12) as ���� " & _
              "         ,substr(C.����,1,8) as ����,substr(C.ҽ����,1,8)  as ҽ����,D.���֤��,substr(C.��λ����,1,5) as ��λҽ���� " & _
              "         ,substr(C.��Ա���,1,2) as ��Ա���,F.����,substr(A.�Ǽ���,1,8) as ҽ��,nvl(substr(G.����,1,50),'�޲���') as ��Ժ���� " & _
              "         ,trunc(A.��Ժ����)," & str��� & " as ���," & mstr�ս����� & " as �ս����� " & _
              "  from ������ҳ A,�����ʻ� C,������Ϣ D,��������Ŀ¼ E,���ű� F,���ղ��� G " & _
              "  Where A.���� =" & TYPE_ͭ���� & " And C.����ʱ�� <" & mstr�������� & " And A.��Ժ����ID = F.ID And A.��Ժ���� Is Not Null " & _
              "       and A.����ID=C.����ID and C.����=" & TYPE_ͭ���� & " and A.����ID=D.����ID and C.����=E.���� and C.����=E.��� and C.����ID=G.ID(+)  " & _
              "       and E.��� IN (" & mstr���InOracle & ")" & IIf(bln�ָ�, " And trunc(A.�Ǽ�ʱ��) < " & mstr��ʼ���� & " And C.����ʱ�� >= " & mstr��ʼ����, " and nvl(A.�Ƿ��ϴ�,0)=0")
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UInHosRegister", "��Ժ�Ǽ�")
    gstrSQL = "update ������ҳ A Set A.�Ƿ��ϴ� = 1 " & _
              " where A.����=" & TYPE_ͭ���� & " and A.��Ժ���� Is Not Null and A.�Ǽ�ʱ��<" & mstr�������� & _
              " and exists (select B.���� from �����ʻ� B where A.����ID =B.����ID And B.����=A.���� and B.���� IN (" & mstr���InOracle & "))"
    gcnOracle.Execute gstrSQL
    
    '5����������
    '2003-03-03 ֧�������������Ϊ�˱�֤�ϼƽ�����ȷ  and B.���=1 " &
    gstrSQL = "select '" & mstrҽԺ���� & "' as ҽԺ����,B.����ID||'_'||B.��ҳID as ��Ժ���,B.No||decode(B.��¼״̬,2,'2','1') as ��� " & _
              "         ,F.����,substr(B.����Ա����,1,8) as ҽ�� " & _
              "         ,sum(B.ʵ�ս��)*decode(B.��¼״̬,2,-1,1) as ���,decode(b.��¼״̬,2,-1,1) as ��Ʊ,B.�Ǽ�ʱ��," & str��� & " as ���," & mstr�ս����� & " as �ս����� " & _
              "  from ���˷��ü�¼ B,�����ʻ� C,������ҳ D,��������Ŀ¼ E,���ű� F " & _
              "  where B.��¼���� in (2,3) And Nvl(B.���ӱ�־,0)<>9 and B.�Ǽ�ʱ��<" & mstr�������� & _
              "       and B.����ID=C.����ID and C.����=" & TYPE_ͭ���� & " and B.����ID=D.����ID AND B.��ҳID=D.��ҳID AND D.����=C.���� and C.����=E.���� and C.����=E.��� and B.��������ID=F.ID " & _
              "       and E.��� IN (" & mstr���InOracle & ")" & IIf(bln�ָ�, " and B.�Ǽ�ʱ��>=" & mstr��ʼ����, " And B.�Ǽ�ʱ�� >=" & mstrȱʡ��ʼ���� & " and Nvl(B.�Ƿ��ϴ�,0)<>1") & _
              "  group by B.NO,B.����ID,B.��ҳID,F.����,B.����Ա����,B.��¼״̬,B.�Ǽ�ʱ��"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UInHosBill", "��������")
    
    '6��������ϸ
    gstrSQL = "select '" & mstrҽԺ���� & "' as ҽԺ����,B.����ID||'_'||B.��ҳID as ��Ժ���,B.No||decode(B.��¼״̬,2,'2','1') as ���,substr(E.����,1,8) as ����ҩ�����,trim(substr(E.����,1,40)) as ����ҩ������ " & _
              "         ,trim(substr(G.����,1,40)) as ҽԺҩ������,Round(B.ʵ�ս��/(B.����*B.����),4) as ʵ�ʼ۸�,B.����*B.����*decode(B.��¼״̬,2,-1,1) as ���� " & _
              "         ,B.ʵ�ս��*decode(B.��¼״̬,2,-1,1) as ���" & _
              "         ,B.������Ŀ��,E.�������,E.�����Ը����� " & _
              "         ,E.���ͱ���,nvl(substr(decode(Instr(G.���,'��'),0,G.���,substr(G.���,1,Instr(G.���,'��')-1)),1,40),' ') as ��� " & _
              "  from " & gstrOwner & ".���˷��ü�¼ B," & gstrOwner & ".�����ʻ� C,������Ŀ E," & gstrOwner & ".�շ�ϸĿ G," & gstrOwner & ".������ҳ H " & _
              "  where B.��¼���� in (2,3) And Nvl(B.���ӱ�־,0)<>9 and B.�Ǽ�ʱ��<" & mstr�������� & _
              "       and B.����ID=C.����ID and C.����=" & TYPE_ͭ���� & " and B.����ID=H.����ID AND B.��ҳID=H.��ҳID AND H.����=C.���� " & _
              "       and B.���ձ���=E.���� and E.����=C.���� and B.�շ�ϸĿID=G.ID And Nvl(B.����,0)<>0" & _
              "       and C.���� IN (" & mstr���InOracle & ")" & IIf(bln�ָ�, " and B.�Ǽ�ʱ��>=" & mstr��ʼ����, " And B.�Ǽ�ʱ�� >=" & mstrȱʡ��ʼ���� & "    AND Nvl(B.�Ƿ��ϴ�,0)<>1")
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UInHosBillDetail", "������ϸ")
        
    '7��סԺ����
    gstrSQL = "select ���Ĵ���,'" & mstrҽԺ���� & "' as ҽԺ����,���,��Ʊ��,A.����ID||'_'||A.��ҳID as סԺ�ǼǺ�,B.סԺ��,A.ҽ����,A.���֤��,��λҽ����,A.����,A.�Ա�,A.���� " & _
              "         ,A.����,A.��ݴ���,A.�Ƿ���Ա,�Ƿ�ҽ���չ˶���,�μӲ��䱣��,�ʻ��ۼ�����,�ʻ��ۼ�֧�� " & _
              "         ,ͳ����֧�����,ͳ����֧������ " & _
              "         ,סԺ�����ʻ�֧�����,A.סԺ����,'" & mstrҽԺ���� & "' as ҽԺ�ȼ�,��������,ҽ������,�������,A.��Ժ����,A.��Ժ����,A.סԺ���� " & _
              "         ,�������ý��,ȫ�Ը����,�����Ը����,ת�������Ը�,A.סԺ����,����,ʵ������,ͳ�����Ը�,�����ʻ�֧��,ͳ����֧��,ͳ�����Ը�,ͳ�����֧��,ͳ������Ը� " & _
              "         ,�������֧��,��������Ը�,��������֧��,���������Ը�,��һ��֧��,��һ���Ը�,�ڶ���֧��,�ڶ����Ը�,������֧��,�������Ը�,���Ķ�֧��,���Ķ��Ը�,�����֧��,������Ը�" & _
              "         ,�������ⶥ��,������ⶥ��,����ͳ��֧��,����ͳ�����,��������֧��,������������ " & _
              "         ,����סԺ����,������֧��,������������ʻ�֧��,�������Բ��𸶽�,����סԺ�����ʻ�֧�� " & _
              "         ,�ʻ��ۼ�����,�ʻ��ۼ�֧��+�����ʻ�֧��,֧��˳���,���Ҷȼ�,��;����,��Ʊ��־,����Ʊ�ݺ�,Ʊ������,A.���," & mstr�ս����� & " as �ս����� " & _
              "  from ���ս����¼ A," & gstrOwner & ".������Ϣ B" & _
              "  Where A.���� = 2 And A.���� =" & TYPE_ͭ���� & "  And A.Ʊ������ <" & mstr�������� & _
              "       and A.���Ĵ��� in (" & mstr����InOracle & ") And A.����ID=B.����ID" & IIf(bln�ָ�, "  And A.Ʊ������ >=" & mstr��ʼ����, " And A.Ʊ������ >=" & mstrȱʡ��ʼ���� & "  And Nvl(A.�Ƿ��ϴ�,0)<>1")
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UInHosBalance", "סԺ����")
    
    '8������ҽ������
    gstrSQL = "select '" & mstrҽԺ���� & "' as ҽԺ����,A.���,C.����,sum(B.���ʽ��) as ��� " & _
              "  from ���ս����¼ A," & gstrOwner & ".���˷��ü�¼ B," & gstrOwner & ".����֧������ C " & _
              "  Where A.���� = 2 And A.���� = " & TYPE_ͭ���� & " And A.��¼ID = B.����id+0 And Nvl(B.���ӱ�־,0)<>9 And A.Ʊ������ <" & mstr�������� & _
              "        And B.���մ���ID = C.ID AND A.����ID=B.����ID and C.����=" & TYPE_ͭ���� & IIf(bln�ָ�, " And A.Ʊ������ >= " & mstr��ʼ����, " And A.Ʊ������ >= " & mstrȱʡ��ʼ���� & " And Nvl(A.�Ƿ��ϴ�,0)<>1 ") & _
              "       and A.���Ĵ��� IN (" & mstr����InOracle & ")" & _
              "  group by A.���,C.����"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UInHosMediKind", "����ҽ������")
    
    gstrSQL = "update ���˷��ü�¼ A Set A.�Ƿ��ϴ� = 1 " & _
              " where A.��¼���� in (2,3) and A.�Ǽ�ʱ��>=" & mstrȱʡ��ʼ���� & " and A.�Ǽ�ʱ��<" & mstr�������� & _
              " and exists (select B.���� from ������ҳ B,�����ʻ� C where A.����ID =B.����ID and A.��ҳID =B.��ҳID and B.����=" & TYPE_ͭ���� & _
              " and B.����ID=C.����ID And Nvl(A.�Ƿ��ϴ�,0)<>1 And B.����=C.���� and C.���� IN (" & mstr���InOracle & "))"
    gcnOracle.Execute gstrSQL
    
    '9����Ժ����
    gstrSQL = "select '" & mstrҽԺ���� & "' as ҽԺ����,A.��� " & _
              "         ,B.����,B.����,B.���,B.���� " & _
              "  from ���ս����¼ A," & gstrOwner & ".�����ʻ� C," & gstrOwner & ".���ղ��� B " & _
              "  Where A.���� =" & TYPE_ͭ���� & " And A.Ʊ������ <" & mstr�������� & IIf(bln�ָ�, " And A.Ʊ������ >= " & mstr��ʼ����, " And A.Ʊ������ >= " & mstrȱʡ��ʼ���� & " and Nvl(A.�Ƿ��ϴ�,0)<>1") & _
              "        and C.����ID=B.ID and A.����ID=C.����ID And Nvl(A.�Ƿ��ϴ�,0)<>1 and C.����=" & TYPE_ͭ���� & " and C.���� IN (" & mstr���InOracle & ")"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UInHosSick", "��Ժ����")
    
    gstrSQL = "update ���ս����¼ A Set A.�Ƿ��ϴ� = 1 " & _
              " where Nvl(A.�Ƿ��ϴ�,0)=0 and ����=" & TYPE_ͭ���� & " And Ʊ������ >=" & mstrȱʡ��ʼ���� & " And Ʊ������ <" & mstr��������
    gcnҽ��.Execute gstrSQL
    
    '10����Ŀ��Ӧ�䶯��¼
    gstrSQL = "select '" & mstrҽԺ���� & "' as ҽԺ����,����ҩ�����,trim(����ҩ������) ����ҩ������,trim(ҽԺҩ������) ҽԺҩ������,�������� " & _
              "  from ��Ŀ��Ӧ��־ A " & _
              "  Where A.�������� >= " & mstr��ʼ���� & " And A.�������� <" & mstr��������
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    Call �����ļ�(rsTemp, gstrSQL, "UAssociateItems", "��Ŀ������")
    
    �ս� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub �����ļ�(rsData As ADODB.Recordset, ByVal strSource As String, ByVal strFile As String, ByVal str��Ŀ As String)
'���ݼ�¼�������ļ�
    Dim txtFile As TextStream
    Dim fld As ADODB.Field
    Dim strLine As String, lngLines As Long
    
    
    Set txtFile = mobjFileSys.CreateTextFile(mstr�����ϴ�Ŀ¼ & strFile)
    
    lbl��Ŀ.Caption = str��Ŀ
    lngLines = rsData.RecordCount
    DoEvents
    Do Until rsData.EOF
        SetProgress lngLines, rsData.AbsolutePosition
        
        strLine = ""
        For Each fld In rsData.Fields
            If IsNull(fld.Value) Then
                If fld.Type = adNumeric Then
                    strLine = strLine & "0|"
                Else
                    strLine = strLine & "|"
                End If
            Else
                If fld.Type = adDBTimeStamp Then
                    strLine = strLine & Format(fld.Value, "yyyy-MM-dd HH:mm:ss") & "|"
                Else
                    strLine = strLine & fld.Value & "|"
                End If
            End If
        Next
        '��β�Ա���|����
'        strLine = Mid(strLine, 1, Len(strLine) - 1)
        
        'д��һ�м�¼
        txtFile.WriteLine strLine
        rsData.MoveNext
    Loop
    txtFile.Close
End Sub

Private Function UpLoadFile(ByVal strFile As String) As Boolean
'���ܣ��ϴ�ָ�����ļ���������ɽ�ѹ������
    Dim zipFile As ZIPnames
    Dim lngCount As Integer, zipname As String
    Dim recurse As Integer, updat As Integer, freshen As Integer, junk As Integer
    
    Dim lngReturn As Long
    Dim strPath As String
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    strPath = mstr�����ϴ�Ŀ¼
    strPath = strPath & IIf(Right(strPath, 1) <> "\", "\", "")
    
    '���ȶ��ļ�����ѹ��
    junk = 1    ' 1=throw away path names
    recurse = 0 ' 1=recurse -R 2=recurse -r 2=most useful :)
    updat = 0   ' 1=update only if newer
    freshen = 0 ' 1=freshen - overwrite only
    
    zipFile.s(0) = ""
    zipFile.s(1) = strPath & "UClinicBill"
    zipFile.s(2) = strPath & "UClinicBillDetail"
    zipFile.s(3) = strPath & "UClinicMediKind"
    zipFile.s(4) = strPath & "UInHosRegister"
    zipFile.s(5) = strPath & "UInHosBill"
    zipFile.s(6) = strPath & "UInHosBillDetail"
    zipFile.s(7) = strPath & "UInHosBalance"
    zipFile.s(8) = strPath & "UInHosMediKind"
    zipFile.s(9) = strPath & "UInHosSick"
    zipFile.s(10) = strPath & "UAssociateItems"
    lngCount = 11
    zipname = strPath & strFile & ".zip"
    
    If mobjFileSys.FileExists(zipname) = True Then
        mobjFileSys.DeleteFile zipname, True
    End If
    If VBZip(lngCount, zipname, zipFile, junk, recurse, updat, freshen, strPath) = False Then
        Exit Function
    End If
    
    
    '���ļ����м���
    EncryptFiles zipname, strPath & strFile
    
    '�ϴ��ļ�
    lngReturn = FTPUpLoad(mstr�ϴ�IP, "21", mstr�ϴ��û�, mstr�ϴ�����, strPath & strFile, mstrԶ���ϴ�Ŀ¼, strFile)
    If lngReturn <> 0 Then
        MsgBox "���ڡ�" & mstr�������� & "�����ļ�" & strFile & "�ϴ�ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    UpLoadFile = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Function Get��������(rsHost As ADODB.Recordset) As Boolean
'���ܣ�����������ţ��õ���Ը����ĵĲ�������
    Dim rsTemp As New ADODB.Recordset
    
    '��ʼ������
    mstr�������� = NVL(rsHost("����"))
    mstr�������� = NVL(rsHost("����"))
    
    mstrԭҽ���� = NVL(rsHost("ҽ����"))
    mlngװǮ��� = NVL(rsHost("װǮ���"), 0)
    mlng������������� = NVL(rsHost("�������������"), 0)
    mlng����������� = NVL(rsHost("�����������"), 0)
    mlng��Ŀ������� = NVL(rsHost("��Ŀ�������"), 0)
    mlng����������� = NVL(rsHost("�����������"), 0)
    mlng���ݸɲ���� = NVL(rsHost("���ݸɲ��������"), 0)
    mlng������Ա������� = NVL(rsHost("������Ա�������"), 0)
    
    mstr��������Ŀ¼ = NVL(rsHost("�������ص�ַ"))
    mstr�����ϴ�Ŀ¼ = NVL(rsHost("�����ϴ���ַ"))
    If mstr��������Ŀ¼ = "" Or mstr�����ϴ�Ŀ¼ = "" Then
        MsgBox "������������" & rsHost("����") & "���ı����ϴ�Ŀ¼�ͱ�������Ŀ¼��", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr��������Ŀ¼ = mstr��������Ŀ¼ & IIf(Right(mstr��������Ŀ¼, 1) <> "\", "\", "")
    mstr�����ϴ�Ŀ¼ = mstr�����ϴ�Ŀ¼ & IIf(Right(mstr�����ϴ�Ŀ¼, 1) <> "\", "\", "")
    
    '�õ�Զ����������Ϣ
    gstrSQL = "SELECT B.* FROM ������������ B " & _
              " Where  B.����=" & TYPE_ͭ���� & " And B.����='" & rsHost("����") & "' " & _
              "    AND nvl(B.��ʼ����,to_date('2000-01-01','yyyy-MM-dd'))<=SYSDATE  AND nvl(B.��ֹ����,to_date('3000-01-01','yyyy-MM-dd'))>=trunc(SYSDATE)"
    rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount <> 1 Then
        MsgBox "������" & rsHost("����") & "�����ϴ����ز����д�", vbInformation, gstrSysName
        Exit Function
    End If
    mstr�ϴ�IP = NVL(rsTemp("�ϴ�IP"))
    mstr�ϴ��û� = NVL(rsTemp("�ϴ��û�"))
    mstr�ϴ����� = NVL(rsTemp("�ϴ�����"))
    mstr����IP = NVL(rsTemp("����IP"))
    mstr�����û� = NVL(rsTemp("�����û�"))
    mstr�������� = NVL(rsTemp("��������"))
    mstrԶ���ϴ�Ŀ¼ = NVL(rsTemp("�ϴ�Ŀ¼"))
    mstrԶ������Ŀ¼ = NVL(rsTemp("����Ŀ¼"))
    
    Get�������� = True
End Function

Private Function GetҽԺ����() As Boolean
'���ܣ�����ҽԺ����ز�������ҽԺ���롢ҽԺ�ȼ�
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    '�õ�ҽԺ����
    gstrSQL = "select ҽԺ���� from ������� where ���=" & TYPE_ͭ����
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ʼ��ҽԺ��ҽ�����롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    If IsNull(rsTemp("ҽԺ����")) = True Then
        MsgBox "���ʼ��ҽԺ��ҽ�����롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstrҽԺ���� = Mid(rsTemp("ҽԺ����"), 1, 4)
    
    '�õ�ҽԺ�ȼ�
    gstrSQL = "select ����ֵ from ���ղ��� where ����=" & TYPE_ͭ���� & " And ������='ҽԺ����'"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "����ҽ�������г�ʼ��ҽԺ��ҽԺ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstrҽԺ���� = Mid(rsTemp("����ֵ"), 1, 2)
    GetҽԺ���� = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub SetProgress(lngSum As Long, lngValue As Long)
'��ʾ����ֵ
    If lngSum = 0 Then
        pgb.Value = 0
    Else
        pgb.Value = lngValue / lngSum * 100
    End If
End Sub

Private Function Is����װǮ(ByVal �������� As String) As Boolean
'���ܣ��жϵ�ǰ�����Ƿ�������װǮ
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select װǮģʽ From �������� Where ����=" & TYPE_ͭ���� & " and ����='" & �������� & "'"
    Call OpenRecordset(rsTemp)
    
    Is����װǮ = (NVL(rsTemp("װǮģʽ"), 0) = 2)
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Private Sub Get�����б�(ByVal �������� As String)
'���ܣ��жϵ�ǰ�����Ƿ�������װǮ
    Dim rsTemp As New ADODB.Recordset
    
    mstr����InOracle = ""
    mstr���InOracle = ""
    mstr����InStr = ""
    On Error GoTo errHandle
    
    gstrSQL = "Select ���,���� From ��������Ŀ¼ Where ����=" & TYPE_ͭ���� & " and ��������='" & �������� & "'"
    Call OpenRecordset(rsTemp)
    
    Do Until rsTemp.EOF
        mstr���InOracle = mstr���InOracle & "," & rsTemp("���")
        mstr����InOracle = mstr����InOracle & ",'" & rsTemp("����") & "'"
        mstr����InStr = mstr����InStr & "," & rsTemp("����")
        rsTemp.MoveNext
    Loop
    
    If mstr����InOracle = "" Then
        mstr���InOracle = "''"
        mstr����InOracle = "''"
    Else
        mstr���InOracle = Mid(mstr���InOracle, 2)
        mstr����InOracle = Mid(mstr����InOracle, 2)
    End If
    
    Exit Sub
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Sub

Private Function GetDateForOracle(ByVal ���� As String) As String
'���ܣ�������ͨ�����ڴ��õ�����Oracle������ֵ
    GetDateForOracle = "To_date('" & Format(CDate(AddDate(����)), "yyyy-MM-dd") & "','yyyy-MM-dd')"
End Function

Private Sub OpenText(ByVal FileName As String, TextFile As TextStream, Lines As Long)
'���ܣ���ָ���ļ������õ�������
    Set TextFile = mobjFileSys.OpenTextFile(FileName)
    Do While Not TextFile.AtEndOfStream
        TextFile.ReadLine
    Loop
    Lines = TextFile.Line
    Set TextFile = mobjFileSys.OpenTextFile(FileName)

End Sub

Private Sub Update��������(ByVal �ֶ��� As String, ByVal ֵ As String)
'���ܣ������뱣��������صĲ���
    gstrSQL = "Update �������� Set " & �ֶ��� & "='" & ֵ & "' Where ����= " & TYPE_ͭ���� & " And ����='" & mstr�������� & "'"
    gcnҽ��.Execute gstrSQL
End Sub

Private Function Calc���÷ָ�(rs������ϸ As ADODB.Recordset, _
                 curȫ�Է� As Currency, cur�����Ը� As Currency, curͳ�� As Currency) As Boolean
'���ܣ����ݷ�����ϸ�����¼�����ϸ�з��õı���������õĽ�����ֱ���ϴ�
'������rs������ϸ  ������ϸ���������õ�ϸĿID�����ۡ����������
'      �Ƿ����     �Ƿ���Ҫ�����ݿ��в��˷��ü�¼��ҽ�����ݽ��и��¡�����Ԥ��ʱ������
'      curȫ�Է�    ���������������ȫ�ԷѲ��ֵĽ��
'      cur�����Ը�  ��������������������Ը����ֵĽ��
'      curͳ��      ���������������ͳ�ﲿ�ֵĽ��
'      ���÷ָ�     ���������Ϊ���ʾ�޼۴Ӳ��˷��ü�¼�ж�ȡ�������㵱ǰ�Ǳʼ�¼
'���أ��������ɹ�������й��ܣ�ΪTrue
'����λ�ã�����Ԥ�㡢������㡢סԺ���ʡ�סԺԤ�㡢סԺ���㡢������ϸ�ϴ�

    Dim str���ı��� As String, str���ֱ��� As String, lng����ID As Long
    Dim rs���մ��� As New ADODB.Recordset
    Dim rs������׼ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset, str��Ŀ���� As String, strϸĿ���� As String
    Dim cur��� As Currency, cur���۸� As Currency, cur���� As Currency, cur�Ը����� As Currency, cur��λ�� As Currency, cur������Ŀ As Currency
    Dim curͳ���� As Currency, cur�Ը� As Currency, lng���մ���ID As Long, lng������Ŀ�� As Long
    Dim gcnUpdate As ADODB.Connection
    
    Set gcnUpdate = New ADODB.Connection
    With gcnUpdate
        If .State = 1 Then .Close
        .Open gcnOracle.ConnectionString
        .BeginTrans
    End With
    
    curȫ�Է� = 0
    cur�����Ը� = 0
    curͳ�� = 0
    
    On Error GoTo errHandle
    '�õ�����ҽ������
    gstrSQL = "SELECT A.ID,A.���� FROM ����֧������ A Where A.���� =" & TYPE_ͭ����
    rs���մ���.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    'Modified by zyb ##2003-08-31
    If rs������ϸ.RecordCount > 0 Then rs������ϸ.MoveFirst
    Do Until rs������ϸ.EOF
        If lng����ID <> rs������ϸ("����ID") Then
            lng����ID = rs������ϸ("����ID")
            '��ͬ�Ĳ��ˣ��������ڲ�ͬ�����ģ��䴲λ�޼�Ҳ���ܲ�ͬ������Ҫ��������
            gstrSQL = "SELECT B.���� ����,C.���� AS ���ֱ��� " & _
                "FROM �����ʻ� A,��������Ŀ¼ B,���ղ��� C " & _
                "WHERE A.����ID=" & lng����ID & " AND A.����=" & TYPE_ͭ���� & " AND A.����=B.���� AND nvl(A.����,0)=nvl(B.���,0) AND A.����ID=C.ID(+)"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
            
            '�õ���ҽ�����˵Ĳ�����׼��Ŀ
            gstrSQL = "SELECT A.��Ŀ���,A.�����Ը����� FROM ���ղ�����Ŀ A Where A.������� ='" & rsTemp("���ֱ���") & "'"
            If rs������׼.State = adStateOpen Then rs������׼.Close
            rs������׼.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
            
            '�õ������Ĺ涨�Ĵ�λ���޼�
            str���ı��� = rsTemp("����")
            gstrSQL = "Select ÿ�촲λ���޼�,������Ŀ�۸� From ��������Ŀ¼ Where ����=" & TYPE_ͭ���� & " And ����='" & rsTemp("����") & "'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
            cur��λ�� = rsTemp("ÿ�촲λ���޼�")
            cur������Ŀ = NVL(rsTemp("������Ŀ�۸�"), 0)
        End If
        
        If IsNull(rs������ϸ("��Ŀ����")) = True Then
            MsgBox "��Ϊ" & rs������ϸ("����") & "����ҽ�����롣", vbInformation, gstrSysName
            gcnUpdate.RollbackTrans
            Exit Function
        End If
        str��Ŀ���� = rs������ϸ("��Ŀ����")
        strϸĿ���� = rs������ϸ("����")
        
        '��ñ�����Ŀ����ϸ��Ϣ���������
        gstrSQL = "Select * from ������Ŀ Where ����=" & TYPE_ͭ���� & " And ����='" & str��Ŀ���� & "'"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnҽ��, adOpenStatic, adLockReadOnly
        If rsTemp.EOF Then
            MsgBox strϸĿ���� & "�ı��ձ������󣬲�����ɽ��㡣", vbInformation, gstrSysName
            gcnUpdate.RollbackTrans
            Exit Function
        End If
        
        If rs������ϸ("�շ����") = "J" Then
            '��λ��
            lng������Ŀ�� = 1
            If rs������ϸ("����") <= cur��λ�� Then
                curͳ���� = rs������ϸ("ʵ�ս��")
            Else
                curͳ���� = cur��λ�� * rs������ϸ("����")
            End If
            curͳ�� = curͳ�� + curͳ����
            curȫ�Է� = curȫ�Է� + (rs������ϸ("ʵ�ս��") - curͳ����)
        Else
            '�������Ŀ�������Ա����ļ۸�
            cur���۸� = IIf(NVL(rsTemp("���۸�����"), 0) = 0, NVL(rsTemp("�۸�"), 0), rsTemp("���۸�����"))
            If cur���۸� > 0 And cur���۸� < rs������ϸ("����") Then
                '����Ŀ��������޼ۣ����ұ�ҽԺ�۸�Ҫ��
                cur���� = cur���۸�
            Else
                cur���� = rs������ϸ("����")
            End If
            
            rs������׼.Filter = "��Ŀ���='" & str��Ŀ���� & "'"
            If rs������׼.EOF = False Then
                '�Ƿ�ҽ����Ŀ�����˴���׼
                lng������Ŀ�� = IIf(rs������׼("�����Ը�����") = 1, 0, 1)
                cur�Ը����� = rs������׼("�����Ը�����")
            Else
                '�Ա�����Ŀ�е�ֵΪ׼
                lng������Ŀ�� = rsTemp("�Ƿ�ҽ��")
                cur�Ը����� = rsTemp("�����Ը�����")
                
                If lng������Ŀ�� = 1 And cur������Ŀ > 0 And _
                    (rs������ϸ("�շ����") <> "5" And rs������ϸ("�շ����") <> "6" And rs������ϸ("�շ����") <> "7") Then
                    
                    '���ڰ��۸����ּ����������Ŀ������
                    If rs������ϸ("����") >= cur������Ŀ Then
                        cur�Ը����� = 0.2
                    Else
                        cur�Ը����� = 0
                    End If
                End If
                
                '��Ȼ����Ϊ������Ŀ���������Ը��������Ը�Ϊȫ�Է�
                If lng������Ŀ�� = 1 And rsTemp("�����Ը�����") = 1 Then lng������Ŀ�� = 0
            End If
            
            If lng������Ŀ�� = 0 Then
                'ȫ�Է���Ŀ
                curͳ���� = 0
                curȫ�Է� = curȫ�Է� + rs������ϸ("ʵ�ս��")
            Else
                If cur���۸� = 0 Or rs������ϸ("����") <= cur���۸� Then
                    'û�м۸����ƣ��������Ƶļ۸�û�г���
                    curͳ���� = rs������ϸ("ʵ�ս��") * (1 - cur�Ը�����)
                Else
                    '�м۸����ƣ���ֻ��ȡ���۸�
                    curͳ���� = cur���۸� * rs������ϸ("����") * (1 - cur�Ը�����)
                End If
                curͳ�� = curͳ�� + curͳ����
                
                'Modified by zyb ##2003-08-31
                '���������۸�����ʱ,�������Ը��ļ������Ӧ����(ȫ�Ը�=���޲���+��ҽ����Ŀ�ķ���;ʵ�ս��=ͳ����+�����Ը�+ȫ�Ը�)
                If cur���۸� > 0 And cur���۸� < rs������ϸ("����") Then
                    cur�Ը� = (cur���۸� * rs������ϸ("����") - curͳ����)
                Else
                    cur�Ը� = (rs������ϸ("ʵ�ս��") - curͳ����)
                End If
                cur�����Ը� = cur�����Ը� + cur�Ը�
                curȫ�Է� = curȫ�Է� + (rs������ϸ("ʵ�ս��") - curͳ���� - cur�Ը�)
                'Modified end
            End If
        End If
        
        rs���մ���.Filter = "����='" & rsTemp("�������") & "'"
        If rs���մ���.EOF = False Then
            lng���մ���ID = rs���մ���("ID")
        Else
            lng���մ���ID = 0
        End If
        
        '����������ƣ����������������շѷ���һ�������С�Ȼ��סԺ���ݶ����Ѿ�������˵ģ������ô���㶼����ν
        'Modified by zyb ##2003-09-01(��Ϊͳһ��ΪԤ����ʱȫ������,���Բ������Ƿ��ϴ���־)
        gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs������ϸ("ID") & "," & curͳ���� & "," & _
            lng���մ���ID & "," & lng������Ŀ�� & ",'" & str��Ŀ���� & "',NULL," & cur���۸� & ")"
        gcnUpdate.Execute gstrSQL, , adCmdStoredProc
        
        rs������ϸ.MoveNext
    Loop
    
    gcnUpdate.CommitTrans
    Calc���÷ָ� = True
    Exit Function
errHandle:
    MsgBox "  ���÷ָ�ʱ,�������д���:" & vbCrLf & "  " & Err.Description
    gcnUpdate.RollbackTrans
End Function

