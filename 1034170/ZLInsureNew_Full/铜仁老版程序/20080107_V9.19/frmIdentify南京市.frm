VERSION 5.00
Begin VB.Form frmIdentify�Ͼ��� 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������������֤"
   ClientHeight    =   3480
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4635
   Icon            =   "frmIdentify�Ͼ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra���� 
      Caption         =   "ҽ�����˻�����Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4404
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   10
         Top             =   1860
         Width           =   1692
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   8
         Top             =   855
         Width           =   1692
      End
      Begin VB.CommandButton cmd������Ϣ 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3624
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1335
         Width           =   372
      End
      Begin VB.TextBox txt���ﲡ�� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   4
         Top             =   1335
         Width           =   1692
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ȷ����"
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
         Left            =   855
         TabIndex        =   11
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   915
         Width           =   480
      End
      Begin VB.Label lbl���ﲡ�� 
         AutoSize        =   -1  'True
         Caption         =   "���ﲡ��(&F)"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1410
         Width           =   1320
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "���ݺ�(&N)"
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
         Left            =   720
         TabIndex        =   1
         Top             =   420
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3432
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2232
      TabIndex        =   6
      Top             =   2880
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify�Ͼ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstrIdentify As String
Private mlng����ID As Long, mlng����ID As Long
Private mstr�������� As String
Private mstr���ֱ��� As String
Private mstr�������� As String

Private Sub cmdCancle_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim strIdentify As String
    Dim strAddition As String
    Dim lngsequence As String
    On Error GoTo errHandle
    
    '�ж��Ƿ�����ҽ����������
    If Trim(Text1.Text) = "" Then
        MsgBox "δ��ȡ��ҽ����������", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    mstr�������� = Trim(Text1.Text)
    
    If Trim(txt���ﲡ��.Text) = "" Or txt���ﲡ�� <> mstr�������� Then
        MsgBox "���ﲡ��δ¼�������", vbInformation, gstrSysName
        txt���ﲡ��.SetFocus
        Exit Sub
    End If
        
    '�˴��޷�ȡ�ÿ��ź�ҽ����,������ʱ���뱣�ղ�������,�Ժ�õ����ź��ٽ����޸�
    lngsequence = Right(String(20, "0") & Text1.Tag, 20)
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������
'      24��������;25�����ۼ�;26����ͳ���޶�
    
    strIdentify = lngsequence & ";"                                       '0����
    strIdentify = strIdentify & lngsequence & ";"                  '1ҽ���ţ����˱�ţ�
    strIdentify = strIdentify & ";"                                 '2����
    strIdentify = strIdentify & mstr�������� & ";"                   '3����
    strIdentify = strIdentify & ";"                                 '4�Ա�
    strIdentify = strIdentify & ";"                                '5��������
    strIdentify = strIdentify & ";"                                 '6���֤
    strIdentify = strIdentify & ";"                               '7.��λ����(����)
    strAddition = "0;"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                               '9.˳���
    strAddition = strAddition & ";"                            '10��Ա���
    strAddition = strAddition & "10000;"                              '11�ʻ����
    strAddition = strAddition & "0;"                            '12��ǰ״̬
    strAddition = strAddition & mlng����ID & ";"                 '13����ID
    strAddition = strAddition & "1;"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";"                             '18�ʻ������ۼ�
    strAddition = strAddition & ";"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & "0;"                            '20����ͳ���ۼ�
    strAddition = strAddition & "0;"                            '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & "0;"                             '22סԺ�����ۼ�
    strAddition = strAddition & ";"                             '23��������
    
    If mlng����ID = 0 Then
        mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID)
    End If
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrIdentify = strIdentify & mlng����ID & ";" & strAddition
    End If
    If Trim(Text2.Text) <> "" Then
        mstr�������� = Trim(Text2.Text)
    Else
        mstr�������� = Trim(Text1.Text)
    End If
    
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MsgBox mlng����ID & "��" & lngsequence & "��" & Text1.Tag & "��" & strIdentify & strAddition, vbInformation, gstrSysName
End Sub


Private Sub cmd������Ϣ_Click()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select id,����,����,decode(���,1,'���Բ�',2,'���ⲡ','��ͨ��') as ���� from ���ղ��� where ����=" & TYPE_�Ͼ���
    Call OpenRecordset(rsTemp, "ѡ����")
    
    If frmListSel.ShowSelect(rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�") Then
        txt���ﲡ��.Text = rsTemp!����
        mlng����ID = rsTemp!ID
        mstr���ֱ��� = rsTemp!����
        mstr�������� = rsTemp!����
    Else
        txt���ﲡ��.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If mbytType = 0 Then
        txt���ﲡ��.Enabled = True
    Else
        txt���ﲡ��.Enabled = False
    End If
End Sub



Private Sub txt���ﲡ��_GotFocus()
    OpenIme ("")
    Call zlControl.TxtSelAll(txt���ﲡ��)
End Sub

Private Sub txt���ﲡ��_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strText As String
    Dim blnReturn As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errorhandle
    '�������ﲡ��
    
    strText = txt���ﲡ��.Text
    gstrSQL = "select A.id,A.����,A.���� from ���ղ��� A where A.����=" & TYPE_�Ͼ��� & " and (" & _
              zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("A", "����", strText) & ")"
    Call OpenRecordset(rsTemp, "���ﲡ��")
    
    If rsTemp.RecordCount = 1 Then
        blnReturn = True
    Else
        blnReturn = frmListSel.ShowSelect(rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
    End If
    
    If blnReturn Then
        txt���ﲡ��.Text = rsTemp!����
        mlng����ID = rsTemp!ID
        mstr���ֱ��� = rsTemp!����
        mstr�������� = rsTemp!����
        zlCommFun.PressKey (vbKeyTab)
    Else
        txt���ﲡ��_GotFocus
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Public Function Identify(ByVal bytType As Byte) As String
    mbytType = bytType
    Me.Show 1
    Identify = mstrIdentify
    With gPatInfo_�Ͼ���
        .�������� = mstr��������
        .���ֱ��� = mstr���ֱ���
        .�������� = mstr��������
    End With
End Function

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim strInput As String, strSql As String, rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt���� <> "" Then
        If Left(txt����.Text, 1) <> "." Then Exit Sub
        strInput = txt����.Text
        If Not IsNumeric(Mid(strInput, 2)) Then Exit Sub
        If Len(Mid(strInput, 2)) <= 4 Then
            strSql = PreFixNO & Format(CDate("1992-" & Format(zlDatabase.Currentdate, "MM-dd")) - CDate("1992-01-01") + 1, "000") & Format(Mid(strInput, 2), "0000") '����˳����
        Else
            strSql = GetFullNO(Mid(strInput, 2))
        End If
        '�������ʱ����Ҫ�ҺŽ���
        strSql = "Select ����id,����,��ʶ�� From ���˷��ü�¼ Where NO='" & strSql & "' And ��¼����=4 And ��¼״̬=1"
        Set rsTemp = gcnOracle.Execute(strSql)
        If rsTemp.EOF Then
            MsgBox "����ĹҺŵ���", vbInformation, gstrSysName
            Exit Sub
        End If
        strSql = "Select * From ������Ϣ Where ����ID=" & rsTemp!����ID
        Set rsTemp = gcnOracle.Execute(strSql)
        If rsTemp.EOF Then
            MsgBox "��ȡ������Ϣ����", vbInformation, gstrSysName
            Exit Sub
        ElseIf IsNull(rsTemp!�����) Then
            MsgBox "�ò��˵�ҽ������û��¼��", vbInformation, gstrSysName
            Exit Sub
        Else
            Text1.Text = rsTemp!����
            Text1.Tag = rsTemp!�����
        End If
        zlCommFun.PressKey (vbKeyTab)
    End If
    Exit Sub
errHandle:
    MsgBox "��ҽ������û�н��������������¹ҺŲ���������", vbInformation, gstrSysName
End Sub


