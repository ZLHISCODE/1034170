VERSION 5.00
Begin VB.Form frmIdentify��Ԫ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   300
      Left            =   810
      MaxLength       =   25
      TabIndex        =   5
      Tag             =   "��ᱣ�Ϻ�"
      Top             =   1320
      Width           =   2265
   End
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "���¶���(&R)"
      Height          =   350
      Left            =   300
      TabIndex        =   25
      Top             =   3705
      Width           =   1305
   End
   Begin VB.ComboBox cbo�籣 
      Height          =   300
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   2265
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5580
      TabIndex        =   23
      Top             =   3705
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4290
      TabIndex        =   22
      Top             =   3705
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   26
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -525
      TabIndex        =   24
      Top             =   3480
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��¼��"
      Height          =   180
      Index           =   2
      Left            =   3780
      TabIndex        =   6
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   4
      Left            =   3600
      TabIndex        =   18
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   4425
      TabIndex        =   19
      Top             =   2565
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   4425
      TabIndex        =   3
      Top             =   900
      Width           =   2265
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����˻�����Ϣ��ʾ������ͨ��[���¶���]��ť���½��ж�ȡ���˻�����Ϣ��"
      Height          =   180
      Left            =   630
      TabIndex        =   27
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify��Ԫ����.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   8
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label lblInf 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��֤��"
      Height          =   180
      Left            =   90
      TabIndex        =   4
      Top             =   1380
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   16
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   6
      Left            =   3600
      TabIndex        =   14
      Top             =   2205
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�籣����"
      Height          =   180
      Index           =   7
      Left            =   90
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   450
      TabIndex        =   12
      Top             =   2205
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   10
      Left            =   90
      TabIndex        =   20
      Top             =   3045
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   810
      TabIndex        =   9
      Top             =   1740
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   4425
      TabIndex        =   7
      Top             =   1320
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   4425
      TabIndex        =   11
      Top             =   1740
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   810
      TabIndex        =   13
      Top             =   2160
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   810
      TabIndex        =   17
      Top             =   2565
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   810
      TabIndex        =   21
      Top             =   3000
      Width           =   5865
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   4425
      TabIndex        =   15
      Top             =   2145
      Width           =   2265
   End
End
Attribute VB_Name = "frmIdentify��Ԫ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '��һ����ϵͳʱ����
Private mblnChange As Boolean
Private Sub cbo�籣_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd�鿨_Click()
   If ��ȡ�α���Ա��Ϣ = False Then
        cmdȷ��.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmdȷ��.Enabled = True
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    cmdȷ��.Enabled = False
End Sub

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
    Dim strInput As String, strOutPut As String
    Dim lng״̬ As Long
    
    IsValid = False
    If Trim(g�������_��Ԫ����.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        If cmd�鿨.Enabled Then cmd�鿨.SetFocus
        Exit Function
    End If
    
     If cbo�籣.Text = "" Then
        ShowMsgbox "�籣������δѡ��"
        Exit Function
    End If
      
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '�����¼ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=" & TYPE_��Ԫ���� & " and ҽ����='" & g�������_��Ԫ����.ҽ��֤�� & "'"
            Call OpenRecordset(rsTemp, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
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
    Dim strInput  As String, strOutPut As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str�籣 As String
    Dim int��ǰ״̬ As Integer
    Dim lng״̬ As Long
    
    
    g�������_��Ԫ����.�������� = Split(cbo�籣.Text, "-")(0)
    g�������_��Ԫ����.�籣���� = cbo�籣.ItemData(cbo�籣.ListIndex)
    If IsValid = False Then Exit Sub
    
    int��ǰ״̬ = 0
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & gintInsure & " and  ҽ����='" & g�������_��Ԫ����.ҽ��֤�� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����id, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_��Ԫ����
        
        strIdentify = .ҽ������                                '0����
        strIdentify = strIdentify & ";" & .ҽ��֤��             '1ҽ����
        strIdentify = strIdentify & ";"                    '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)              '4�Ա�
        strIdentify = strIdentify & ";" & .��������                '5��������
        strIdentify = strIdentify & ";" & .���֤����            '6���֤
        strIdentify = strIdentify & ";" & .��λ����     '7.��λ����(����)
        strAddition = ";0" & .�籣����                                           '8.���Ĵ���
        strAddition = strAddition & ";" & .��¼��                               '9.˳���
        strAddition = strAddition & ";"                                '10��Ա���
        strAddition = strAddition & ";" & .�ʻ����                  '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬                            '12��ǰ״̬
        strAddition = strAddition & ";"             '13����ID
        strAddition = strAddition & ";1"                        '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .��������            '15����֤��
        strAddition = strAddition & ";" & .����                     '16�����
        strAddition = strAddition & ";"                         '17�Ҷȼ�
        strAddition = strAddition & ";"                         '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID)
    If mlng����ID = 0 Then Exit Sub
    
    If mbytType = 0 Or mbytType = 3 Then
    Else
    End If
    g�������_��Ԫ����.����id = mlng����ID
    
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
    If Load�籣���� = False Then
        DebugTool "����ʧ��(�����֤)"
        Exit Function
    End If
    DebugTool "����ɹ�(�����֤)"
    
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g�������_��Ԫ����
        lblEdit(0).Caption = .ҽ������
        txtEdit.Text = .ҽ��֤��
        lblEdit(1).Caption = .����
        lblEdit(2).Caption = .��¼��
        lblEdit(3).Caption = Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)
        lblEdit(4).Caption = .����
        lblEdit(5).Caption = .���֤����
        lblEdit(6).Caption = .��������
        lblEdit(7).Caption = Format(.�ʻ����, "####0.00;-####0.00;;")
        lblEdit(8).Caption = .��λ����
    End With
End Sub
Private Sub Form_Load()
        mblnFirst = True
End Sub

Private Function Load�籣����() As Boolean
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "" & _
        "   Select * From ��������Ŀ¼ " & _
        "   Order by ����"
        
    Err = 0
    On Error GoTo ErrHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption & "�籣����Ŀ¼"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "�����籣����Ŀ¼�����ڲ��������ػ���!"
        Exit Function
    End If
    
    With rsTemp
        cbo�籣.Clear
        Do While Not .EOF
            cbo�籣.AddItem Nvl(!����) & "--" & Nvl(!����)
            cbo�籣.ItemData(cbo�籣.NewIndex) = Nvl(!���, 0)
            .MoveNext
        Loop
    End With
    cbo�籣.ListIndex = 0
    SetDefaultSel
    cbo�籣.Enabled = False
    Load�籣���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo ErrHand:
    Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
    If cbo�籣.ListCount = 0 Then Exit Function
    For i = 0 To cbo�籣.ListCount
        If Split(cbo�籣.List(i), "--")(0) = strReg Then
            cbo�籣.ListIndex = i
            Exit For
        End If
    Next
    If cbo�籣.ListIndex < 0 Then
        cbo�籣.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ��ȡ�α���Ա��Ϣ() As Boolean
    '��ȡ�α���Ա��Ϣ
    Dim strInput As String
    Dim strOutPut As String
    Dim strArr
    
    ��ȡ�α���Ա��Ϣ = False
    
    
    Err = 0
    On Error GoTo ErrHand:
   
    If ҵ������_��Ԫ����(��òα���Ա����, "", strOutPut) = False Then
        Call ClearData
        Exit Function
    End If
    
    strArr = Split(strOutPut, "||")
    '����:ҽ������||ҽ��֤��||���˼�¼��||����||���֤����||��λ����||�Ա�||��������
    
    With g�������_��Ԫ����
        .ҽ������ = strArr(0)
        .ҽ��֤�� = strArr(1)
        .��¼�� = strArr(2)
        .���� = strArr(3)
        .���֤���� = strArr(4)
        .��λ���� = strArr(5)
        .�Ա� = strArr(6)
        .�������� = strArr(7)
        .���� = Get����(.��������)
        .�������� = Split(cbo�籣.Text, "--")(0)
    End With
    
    '��ȡ�ʻ����
    '    YBJGBH  PCHAR   ���ջ������
    '    CPASSWORD   PCHAR   �ֿ��˿�����
    '�����⣬���ݻ��������ô��ȡ.
    strInput = g�������_��Ԫ����.��������
    strInput = strInput & vbTab & g�������_��Ԫ����.����
    If ҵ������_��Ԫ����(��ȡ�ʻ����_����, strInput, strOutPut) = False Then Exit Function
    g�������_��Ԫ����.�ʻ���� = Val(strOutPut)
    
    ��ȡ�α���Ա��Ϣ = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get����(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as ���� from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
        Get���� = Int(Nvl(rsTemp!����, 0))
        Exit Function
    End If
    Exit Function
ErrHand:
End Function
Private Sub ClearData()
    Dim i As Long
    '��������Ϣ
    With g�������_��Ԫ����
        .ҽ������ = ""
        .ҽ��֤�� = ""
        .��¼�� = ""
        .���� = ""
        .���֤���� = ""
        .��λ���� = ""
        .�Ա� = ""
        .�������� = ""
        .���� = 0
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m�ı�ʽ
End Sub

