VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ø�������ģ����ʹ�õ��ķ���"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frm��������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgȨ�� 
      Left            =   930
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��������.frx":628A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5550
      TabIndex        =   3
      Top             =   5100
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6780
      TabIndex        =   4
      Top             =   5100
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img���� 
      Left            =   3240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��������.frx":C524
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img�˵� 
      Left            =   870
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��������.frx":D7A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw���� 
      Height          =   4365
      Left            =   3150
      TabIndex        =   2
      Top             =   630
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7699
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img����"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwģ�� 
      Height          =   2025
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   3572
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img�˵�"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ģ��"
         Object.Width           =   4233
      EndProperty
   End
   Begin MSComctlLib.ListView lvwȨ�� 
      Height          =   2295
      Left            =   30
      TabIndex        =   1
      Top             =   2700
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4048
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgȨ��"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ϸ���ø�ģ������ʹ�õ��ķ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   690
      TabIndex        =   5
      Top             =   390
      Width           =   6900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frm��������.frx":E5F8
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mblnFirst As Boolean            '����
Private mblnReturn As Boolean
Private mrsģ�� As New ADODB.Recordset
Private mrs���� As New ADODB.Recordset

Private mlngModul As Long           '�ϴ�ѡ���ģ��
Private mstrPrivs As String         '�ϴ�ѡ���Ȩ�޴�

Public Function ShowEditor(rsģ�� As ADODB.Recordset, ByVal rs���� As ADODB.Recordset) As Boolean
    mblnReturn = False
    Set mrsģ�� = rsģ��
    Set mrs���� = rs����
    
    Me.Show 1
    
    If mblnReturn Then Set rsģ�� = mrsģ��
    ShowEditor = mblnReturn
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Call SavePrivs
    
    mblnReturn = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mblnFirst = True
    
    'װ��ҽ���ӿ����漰����ģ�飨����Һ�1111�������շ�1121��������Ժ����1131�������������1132��סԺ����(ҽ��)1133��סԺ����1137�����˷��ò�ѯ1139��
    mstrSQL = "Select ���,���� From zlPrograms " & _
        " Where ϵͳ=100 And ��� IN (1111,1121,1131,1132,1133,1137,1139,1203,1205,1206) Order By ���"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "��ȡҽ���ӿ����漰����ģ��")
    With rsTemp
        Do While Not .EOF
            lvwģ��.ListItems.Add , "K_" & !���, !����, , 1
            .MoveNext
        Loop
    End With
    
    '����ɹ�ѡ��ķ�����Ҳ���ǽӿ���֧�ֵķ����б�
    With mrs����
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lvw����.ListItems.Add , "K_" & !���, !����, , 1
            .MoveNext
        Loop
    End With
    
    Call lvwģ��_ItemClick(lvwģ��.ListItems(1))
    
    mblnFirst = False
End Sub

Private Sub lvwģ��_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTemp As New ADODB.Recordset
    '��ȡ����ģ��ӵ�е�Ȩ�ޣ�����ģ���¼���ָ���ѡ��ķ���
    lvwȨ��.ListItems.Clear
    If lvwģ��.SelectedItem Is Nothing Then Exit Sub
    
    Call SavePrivs
    
    mstrSQL = "Select ���� From zlProgfuncs Where ϵͳ=100 And ���=" & Mid(Item.Key, 3)
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "��ȡָ��ģ���Ȩ��")
    With rsTemp
        Do While Not .EOF
            lvwȨ��.ListItems.Add , "K_" & .AbsolutePosition, !����, , 1
            .MoveNext
        Loop
    End With
    
    If lvwȨ��.ListItems.Count <> 0 Then Call lvwȨ��_ItemClick(lvwȨ��.ListItems(1))
End Sub

Private Sub lvwȨ��_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '��ȡָ��ģ�飬ָ��Ȩ����ʹ�õķ���
    Call SavePrivs
    
    Call ShowPrivs
End Sub

Private Sub SavePrivs()
    Dim intItem As Integer, intCount As Integer
    Dim strMethod As String
    Dim strField As String, strValue As String
    
    'ɾ��ѡ��ģ��ָ��Ȩ�޵����й���
    If lvwģ��.SelectedItem Is Nothing Then Exit Sub
    If lvwȨ��.SelectedItem Is Nothing Then Exit Sub
    If mblnFirst Then
        mlngModul = Mid(lvwģ��.SelectedItem.Key, 3)
        mstrPrivs = lvwȨ��.SelectedItem.Text
        Exit Sub
    End If
    
    With mrsģ��
        .Filter = "ģ��=" & mlngModul & " And Ȩ�޴�='" & mstrPrivs & "'"
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    '�����ϴ�ѡ���ģ���ָ��Ȩ��
    strField = "ģ��|Ȩ�޴�|����"
    strValue = mlngModul & "|" & mstrPrivs & "|"
    intCount = lvw����.ListItems.Count
    For intItem = 1 To intCount
        If lvw����.ListItems(intItem).Checked Then
            strMethod = lvw����.ListItems(intItem).Text
            Call Record_Add(mrsģ��, strField, strValue & strMethod)
        End If
    Next
    
    mlngModul = Mid(lvwģ��.SelectedItem.Key, 3)
    mstrPrivs = lvwȨ��.SelectedItem.Text
End Sub

Private Sub ShowPrivs()
    Dim lngģ�� As Long, strȨ�� As String
    Dim intItem As Integer, intCount As Integer
    
    lngģ�� = Mid(lvwģ��.SelectedItem.Key, 3)
    strȨ�� = lvwȨ��.SelectedItem.Text
    
    intCount = lvw����.ListItems.Count
    For intItem = 1 To intCount
        lvw����.ListItems(intItem).Checked = False
    Next
    
    With mrsģ��
        .Filter = "ģ��=" & lngģ�� & " And Ȩ�޴�='" & strȨ�� & "'"
        Do While Not .EOF
            If mrs����.RecordCount <> 0 Then mrs����.MoveFirst
            mrs����.Find "����='" & !���� & "'"
            If Not mrs����.EOF Then
                lvw����.ListItems("K_" & mrs����!���).Checked = True
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
End Sub
