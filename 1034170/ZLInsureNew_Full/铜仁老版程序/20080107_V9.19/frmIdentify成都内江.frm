VERSION 5.00
Begin VB.Form frmIdentify�ɶ��ڽ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo��� 
      Height          =   300
      ItemData        =   "frmIdentify�ɶ��ڽ�.frx":0000
      Left            =   915
      List            =   "frmIdentify�ɶ��ڽ�.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1290
      Width           =   2295
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   6
      Top             =   4710
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4710
      TabIndex        =   5
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   4
      Top             =   705
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -465
      TabIndex        =   3
      Top             =   4500
      Width           =   8340
   End
   Begin VB.TextBox TxtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   915
      MaxLength       =   20
      TabIndex        =   2
      Top             =   915
      Width           =   2295
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4605
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   915
      Width           =   2385
   End
   Begin VB.CommandButton cmd�޸����� 
      Caption         =   "�޸�����"
      Height          =   350
      Left            =   300
      TabIndex        =   0
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   915
      TabIndex        =   41
      Top             =   4050
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ƿ���λ"
      Height          =   180
      Index           =   17
      Left            =   180
      TabIndex        =   40
      Top             =   4095
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   4605
      TabIndex        =   39
      Top             =   4050
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ְ���"
      Height          =   180
      Index           =   16
      Left            =   3825
      TabIndex        =   38
      Top             =   4095
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   4605
      TabIndex        =   37
      Top             =   3645
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   15
      Left            =   3825
      TabIndex        =   36
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "�������������벡�˵�IC�������س�,����ȡ���˵������Ϣ��"
      Height          =   180
      Left            =   720
      TabIndex        =   35
      Top             =   465
      Width           =   5130
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmIdentify�ɶ��ڽ�.frx":0004
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   34
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���˱��"
      Height          =   180
      Index           =   1
      Left            =   3825
      TabIndex        =   33
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   540
      TabIndex        =   32
      Top             =   1762
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   4185
      TabIndex        =   31
      Top             =   1755
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   30
      Top             =   2145
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   29
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   6
      Left            =   3825
      TabIndex        =   28
      Top             =   2145
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����Ч��"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   27
      Top             =   3300
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   540
      TabIndex        =   26
      Top             =   2925
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ͳ����"
      Height          =   180
      Index           =   9
      Left            =   3825
      TabIndex        =   25
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   10
      Left            =   3825
      TabIndex        =   24
      Top             =   3300
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ƿ�����"
      Height          =   180
      Index           =   11
      Left            =   3825
      TabIndex        =   23
      Top             =   2925
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   12
      Left            =   180
      TabIndex        =   22
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   21
      Top             =   1305
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   915
      TabIndex        =   20
      Top             =   1710
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4605
      TabIndex        =   19
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   915
      TabIndex        =   18
      Top             =   2100
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4605
      TabIndex        =   17
      Top             =   2100
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   915
      TabIndex        =   16
      Top             =   2490
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4605
      TabIndex        =   15
      Top             =   2490
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   915
      TabIndex        =   14
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4605
      TabIndex        =   13
      Top             =   2880
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   915
      TabIndex        =   12
      Top             =   3255
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   4605
      TabIndex        =   11
      Top             =   3255
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   915
      TabIndex        =   10
      Top             =   3645
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   13
      Left            =   4185
      TabIndex        =   9
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   14
      Left            =   180
      TabIndex        =   8
      Top             =   1350
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify�ɶ��ڽ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
Private mlng����ID As Long
Private mstrReturn As String
Private mblnChange As Boolean
Private mblnFirst As Boolean

Private Sub cbo���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd�޸�����_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    Dim strInPut As String, strOutPut As String
    
    If InitInfor_�ɶ��ڽ�.������_�ڽ� = 0 Then
        '�����޸�����
    
        strNewPassWord = frm�޸�����.ChangePassword(strOldPassWord, strOldPassWord)
        
        If strOldPassWord = strNewPassWord Then Exit Sub
        If strNewPassWord = "" Then Exit Sub
        '    a)  Port�����������ΪͨѶ�˿ںţ�0��1��2��3�ֱ������1��2��3��4;����Ϊ��I/O��ַ����0x378�������齫���������ӵ�����1��
        '    b)  OldPassword�����������Ϊԭ���룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֣�
        '    c)  NewPassword�����������Ϊ�����룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֡�
        strInPut = InitInfor_�ɶ��ڽ�.���ź�_�ڽ�
        strInPut = strInPut & vbTab & strOldPassWord
        strInPut = strInPut & vbTab & strNewPassWord
        
        If ҵ������_�ɶ��ڽ�(��������_�ڽ�, strInPut, strOutPut) = False Then Exit Sub
        txtEdit(1).Text = strNewPassWord
    Else
        '����ֻ�ܶ���
    End If
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '������������Ķ�����������������
    If InitInfor_�ɶ��ڽ�.������_�ڽ� = 0 Then Exit Sub
    txtEdit(1).Enabled = False
    txtEdit(1).BackColor = txtEdit(0).BackColor
    
    '����
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
    Me.cmd�޸�����.Caption = "����(&R)"
    
    
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        txtEdit(Index).Tag = ""
        g�������_�ɶ��ڽ�.���˱�� = ""
        g�������_�ɶ��ڽ�.���� = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    mblnChange = True
    
    If Index = 1 Then
        '�����������
        '���ȡ������Ϣ
         SetOKCtrl False
         If ReadCardInFo = False Then Exit Sub
        '��ʼֵ
        Call LoadCtrlData
        SetOKCtrl True
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
Private Function ReadCardInFo() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������Ϣ
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInPut As String
     '��ȡ������Ϣ
        '   a)  Port�����������ΪͨѶ�˿ںţ�0��1��2��3�ֱ������1��2��3��4;����Ϊ��I/O��ַ����0x378�������齫���������ӵ�����1��
        '   b)  UserPassword�����������Ϊ�û����룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֣�
           
    ReadCardInFo = False
    strInPut = InitInfor_�ɶ��ڽ�.���ź�_�ڽ�
    If InitInfor_�ɶ��ڽ�.������_�ڽ� = 0 Then
        '��������������
        If Trim(txtEdit(1)) = "" Then
            ShowMsgbox "������IC������!"
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        End If
        strInPut = strInPut & vbTab & txtEdit(1).Text
    End If
    
    Err = 0
    On Error GoTo ErrHand:
    
    If ��ȡ�α���Ա��Ϣ_�ɶ��ڽ�(strInPut) = False Then Exit Function
    ReadCardInFo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmdȷ��.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "��û�н��������֤��", vbInformation, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    If Trim(g�������_�ɶ��ڽ�.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    If cbo���.Text = "" Then
        ShowMsgbox "�������δѡ��"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '�����¼ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=" & TYPE_�ɶ��ڽ� & " and ҽ����='" & g�������_�ɶ��ڽ�.���˱�� & "'"
            Call OpenRecordset(rsTemp, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '����
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str��� As String
    Dim int��ǰ״̬ As Integer
    
    
    If IsValid = False Then Exit Sub
    
    
    g�������_�ɶ��ڽ�.������� = Split(cbo���.Text, "-")(0)
    int��ǰ״̬ = 0
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & gintInsure & " and  ҽ����='" & g�������_�ɶ��ڽ�.���˱�� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤(�������);7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��(ͳ���������|�ƿ�����|����Ч����);16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_�ɶ��ڽ�
        
        strIdentify = .����                               '0����
        strIdentify = strIdentify & ";" & .���˱��           '1ҽ����
        strIdentify = strIdentify & ";"                     '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)              '4�Ա�
        strIdentify = strIdentify & ";" & .��������                '5��������
        strIdentify = strIdentify & ";" & .���֤��           '6���֤
        strIdentify = strIdentify & ";" & IIf(.��λ���� = "", "", "(" & .��λ���� & ")")            '7.��λ����(����)
        strAddition = ";0"                                          '8.���Ĵ���
        strAddition = strAddition & ";"                             '9.˳���
        strAddition = strAddition & ";" & .�������                 '10��Ա���
        strAddition = strAddition & ";" & .�ʻ����                 '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬               '12��ǰ״̬
        strAddition = strAddition & ";"                             '13����ID
        strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .ͳ���� & "|" & .�ƿ����� & "|" & .����Ч�� & "|" & .�ƿ���λ & "|" & .��ְ���    '15����֤��
        strAddition = strAddition & ";" & .��������                     '16�����
        strAddition = strAddition & ";" & .�������                            '17�Ҷȼ�
        strAddition = strAddition & ";" & .�ʻ����                             '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID)
    
    g�������_�ɶ��ڽ�.lng����ID = mlng����ID
    
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    
    DebugTool "���������֤,����ʼ���������Ϣ"
    
    If LoadBaseData = False Then
        DebugTool "����ʧ��(�����֤)"
        Exit Function
    End If
    DebugTool "����ɹ�(�����֤)"
    
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Function LoadBaseData() As Boolean
    '���ػ�������
    Dim rsTemp As New ADODB.Recordset
    LoadBaseData = False
    On Error GoTo ErrHand:
      
    If mbytType = 0 Or mbytType = 3 Then
        cbo���.AddItem "0-��ͨ����"
    Else
        cbo���.AddItem "1-��ͨסԺ"
    End If
    cbo���.ListIndex = cbo���.NewIndex
    LoadBaseData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    With g�������_�ɶ��ڽ�
        txtEdit(0) = .����
        lblEdit(1) = .���˱��
        lblEdit(2) = .����
        lblEdit(3) = Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)
        lblEdit(4) = .���֤��
        lblEdit(5) = .�������
        lblEdit(6) = .��������
        lblEdit(7) = .ͳ����
        lblEdit(8) = .����
        lblEdit(9) = .�ƿ�����
        lblEdit(10) = .����Ч��
        lblEdit(11) = .��������
        lblEdit(12) = .��λ����
        lblEdit(13) = Format(.�ʻ����, "####0.00;#####0.00; ;")
        lblEdit(14) = .�ƿ���λ
        lblEdit(15) = .��ְ���
   End With
End Sub
