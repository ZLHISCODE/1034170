VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentify�����山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmIdentify�����山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3525
      Left            =   -165
      TabIndex        =   39
      Top             =   5055
      Visible         =   0   'False
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   6218
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   6255
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3990
      Width           =   255
   End
   Begin VB.CommandButton cmd�޸����� 
      Caption         =   "�޸�����"
      Height          =   350
      Left            =   360
      TabIndex        =   38
      Top             =   4635
      Width           =   1100
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4590
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   908
      Width           =   1980
   End
   Begin VB.TextBox TxtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   930
      MaxLength       =   20
      TabIndex        =   1
      Top             =   908
      Width           =   2385
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -405
      TabIndex        =   25
      Top             =   4425
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   510
      Width           =   8340
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   9
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5490
      TabIndex        =   10
      Top             =   4635
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   930
      TabIndex        =   7
      Top             =   3975
      Width           =   5610
   End
   Begin VB.ComboBox cbo��� 
      Height          =   300
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1275
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "֧�����"
      Height          =   180
      Index           =   14
      Left            =   195
      TabIndex        =   4
      Top             =   1335
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����(&F)"
      Height          =   180
      Left            =   285
      TabIndex        =   6
      Top             =   4035
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   13
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   4590
      TabIndex        =   37
      Top             =   3210
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   4590
      TabIndex        =   36
      Top             =   2805
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   930
      TabIndex        =   35
      Top             =   3615
      Width           =   5625
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4590
      TabIndex        =   34
      Top             =   2430
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   930
      TabIndex        =   33
      Top             =   2805
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   930
      TabIndex        =   32
      Top             =   3210
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   4590
      TabIndex        =   31
      Top             =   2025
      Width           =   1980
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   930
      TabIndex        =   30
      Top             =   2430
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   930
      TabIndex        =   29
      Top             =   2025
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4590
      TabIndex        =   28
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   930
      TabIndex        =   27
      Top             =   1650
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   26
      Top             =   1283
      Width           =   1980
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   12
      Left            =   3840
      TabIndex        =   22
      Top             =   3255
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�Ʋ������"
      Height          =   180
      Index           =   11
      Left            =   3480
      TabIndex        =   19
      Top             =   2850
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   10
      Left            =   210
      TabIndex        =   21
      Top             =   3667
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ���չ����"
      Height          =   180
      Index           =   9
      Left            =   3480
      TabIndex        =   17
      Top             =   2475
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   570
      TabIndex        =   18
      Top             =   2850
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   7
      Left            =   210
      TabIndex        =   20
      Top             =   3262
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ����Ա���"
      Height          =   180
      Index           =   6
      Left            =   3480
      TabIndex        =   15
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   210
      TabIndex        =   16
      Top             =   2482
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   14
      Top             =   2077
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   4200
      TabIndex        =   13
      Top             =   1695
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   570
      TabIndex        =   12
      Top             =   1702
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���˱��"
      Height          =   180
      Index           =   1
      Left            =   3855
      TabIndex        =   11
      Top             =   1335
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify�����山.frx":000C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ͨ��IC����֤��Ա��ݣ�������֤�����Ϣ��ʾ������"
      Height          =   180
      Left            =   630
      TabIndex        =   24
      Top             =   270
      Width           =   4320
   End
End
Attribute VB_Name = "frmIdentify�����山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
'API��ҽ���ӿ�����
Private Type Struct
    lngAppCode  As Long   '��־����ִ��״̬���롣����1ʱ��ʾ����ִ������������С��0ʱ��ʾ����ִ���쳣�����
    strErrMsg  As String  '������ִ��״̬����AppCodС��0ʱ����������ִ�е��쳣�������Ϣ��
End Type
'��ȡ������
Private Declare Function GetAKC190 Lib "YHMdcrAsistntSvr.dll" Alias "_GetAKC190@12" (ByVal strYab003 As String, ByRef strAkc190 As String, ByRef tmpStrut As Struct) As Boolean
     
Dim mblnChange As Boolean
Private Sub cbo���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd����_Click()
     
        Dim rsTemp As New ADODB.Recordset
        
        With rsTemp
            If .State = 1 Then .Close

            gstrSQL = "" & _
                "   Select id, ����, ����, ֧�����, ������, ���ֽ���취, ���칹������ " & _
                "   From ҽ������Ŀ¼"
                
            Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
            .Open gstrSQL, gcnOracle_CQYB
            Call SQLTest
            If .EOF Then
                MsgBox "�������κβ���,�����أ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt����.Top - .Height
                    .Left = txt����.Left + txt����.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = 1000
                    .ColWidth(5) = 1400
                    .ColWidth(6) = 2000
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt���� = "[" & Nvl(!����) & "]" & IIf(IsNull(!����), "", !����)
                txt����.Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
End Sub

Private Sub cmd�޸�����_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    
    strNewPassWord = frm�޸�����.ChangePassword(strOldPassWord, strOldPassWord)
    If strOldPassWord = strNewPassWord Then Exit Sub
    If strNewPassWord = "" Then Exit Sub
      
    If �޸�����_�����山(strOldPassWord, strNewPassWord) = True Then
        g�������_�����山.���� = strNewPassWord
        cmdȷ��_Click
        Unload Me
        Exit Sub
    End If
End Sub



Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        txtEdit(Index).Tag = ""
    End If
    If Index = 0 And mblnChange = False Then
        g�������_�����山.���˱�� = ""
        g�������_�����山.���� = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    mblnChange = True
    If Index = 0 Then
        SetOKCtrl False
        mblnChange = True
    ElseIf Index = 1 Then
        '�����������
        '���ȡ������Ϣ
         SetOKCtrl False
        
        '�������������
        If ������_�����山 = False Then
            Exit Sub
        End If
         If Trim(txtEdit(Index)) = "" Then
            If mbytType = 0 Then
                '���������,�����Ƿ�ǰ�����,���Ѿ����ڸ��ʻ�ʱ,��������������.
                                
                'ȡ����
                 gstrSQL = "Select ����,����ʱ�� From �����ʻ�  where ����=" & TYPE_�����山 & " and ҽ����='" & g�������_�����山.���˱�� & "'"
                 zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                 
                 If rsTemp.RecordCount = 0 Then
                     ShowMsgbox "����������!"
                    txtEdit(Index).SetFocus
                     Exit Sub
                 End If
                 If Format(rsTemp!����ʱ��, "yyyy-mm-dd") <> Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                    ShowMsgbox "����������!"
                    txtEdit(Index).SetFocus
                    Exit Sub
                 End If
                 txtEdit(Index) = Trim(Nvl(rsTemp!����))
                 If txtEdit(Index) = "" Then
                    ShowMsgbox "����������!"
                    txtEdit(Index).SetFocus
                    Exit Sub
                 End If
            Else
                ShowMsgbox "����������!"
                txtEdit(Index).SetFocus
                Exit Sub
            End If
         End If
         
        txtEdit(0).Text = g�������_�����山.����
        lblEdit(1).Caption = g�������_�����山.���˱��
         
         g�������_�����山.���� = Trim(txtEdit(Index))
        If g�������_�����山.���� = "" Then
            g�������_�����山.���� = Trim(txtEdit(0).Text)
        End If
        If ��ݼ���_�����山 = False Then
            Exit Sub
        End If
        
        If g�������_�����山.���� = "" Then
            ShowMsgbox "��Ч���û���֤,��˲�!"
            Exit Sub
        End If
        
        
        '���������,���Ƚ��йҺŴ���,�����ǲ��ܽ�����Ӧ�Ĵ����.
        If mbytType = 0 Then
            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            gstrSQL = "Select 1 From ���˷��ü�¼ " & _
                    "   where ��¼״̬=1 and ��¼����=4  and rownum<=1 and �Ǽ�ʱ�� between to_date('" & strCurrDate & " 00:00:00','yyyy-mm-dd hh24:mi:ss') and to_date('" & strCurrDate & " 23:59:59','yyyy-mm-dd hh24:mi:ss') and ����id in (select ����id From �����ʻ�  where ����=" & TYPE_�����山 & " and ҽ����='" & g�������_�����山.���˱�� & "')"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
            If rsTemp.RecordCount = 0 Then
                ShowMsgbox "��ҽ������δ���йҺ�,���ܽ����������!"
                Exit Sub
            End If
        End If
        '��ʼֵ
        Call LoadCtrlData
        SetOKCtrl True
    End If
    zlCommFun.PressKey vbKeyTab
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
    IsValid = False
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "��û������ҽ�����ţ�", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    
    If Trim(g�������_�����山.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Trim(txt����) <> "" And Val(txt����.Tag) = 0 Then
        ShowMsgbox "����ѡ�����,������ѡ��!"
        txt����.SetFocus
        Exit Function
    End If
    If cbo���.Text = "" Then
        ShowMsgbox "֧�����δѡ��"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '�����¼ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=" & TYPE_�����山 & " and ҽ����='" & g�������_�����山.���˱�� & "'"
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
    
    lng����ID = Val(txt����.Tag)
    
    If lng����ID <> 0 And txt����.Text <> "" Then
        g�������_�����山.���ֱ��� = Mid(txt����.Text, 2, InStr(1, txt����.Text, "]") - 2)
    Else
        g�������_�����山.���ֱ��� = "000000"
    End If
    g�������_�����山.����ID = lng����ID
    
    g�������_�����山.֧����� = Mid(cbo���.Text, 1, InStr(1, cbo���.Text, "-") - 1)
    int��ǰ״̬ = 0
    
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & gintInsure & " and  ҽ����='" & g�������_�����山.���˱�� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_�����山
        
        strIdentify = .����                               '0����
        strIdentify = strIdentify & ";" & .���˱��           '1ҽ����
        strIdentify = strIdentify & ";" & .����                 '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & Decode(.�Ա�, "1", "��", "2", "Ů", "δ֪")              '4�Ա�
        strIdentify = strIdentify & ";" & .��������                '5��������
        strIdentify = strIdentify & ";" & .���֤��           '6���֤
        strIdentify = strIdentify & ";" & .��λ���� & IIf(.��λ���� = 0, "", "(" & .��λ���� & ")")          '7.��λ����(����)
        strAddition = ";0"                                          '8.���Ĵ���
        strAddition = strAddition & ";"                             '9.˳���
        strAddition = strAddition & ";" & .�籣���칹������          '10��Ա���
        strAddition = strAddition & ";" & .�ʻ����                 '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬                            '12��ǰ״̬
        strAddition = strAddition & ";" & IIf(lng����ID = 0, "", lng����ID)             '13����ID
        strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .ҽ����Ա��� & "|" & .ҽ���չ���� & "|" & .ҽ�Ʋ������ & "|" & .�ۼƽɷ�����     '15����֤��
        strAddition = strAddition & ";" & .����                     '16�����
        strAddition = strAddition & ";"                             '17�Ҷȼ�
        strAddition = strAddition & ";" & .�ʻ����                             '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID)
    
    If mbytType = 3 Or mbytType = 1 Then
        '����ǹҺŻ���Ժ�Ǽ�,��ȷ���µľ�����
        g�������_�����山.������ = Get������_�����山
        If g�������_�����山.������ = "" Then
            ShowMsgbox "�ڻ�ȡ������ʱΪ����,����"
            Exit Sub
        End If
        
        '���±����ʻ��������Ϣ
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & gintInsure & ",'������','''" & g�������_�����山.������ & "''')"
        Call ExecuteProcedure("���������")
        
        If mbytType = 1 Then
            'Ϊ�˱�֤�Ȱ���ͨ��Ժ�ٽ��в�����Ժ�ľ���ʱ�������.
             gstrSQL = "Select ��Ժ���� From ������ҳ where ����id=" & mlng����ID & " And ��Ժ���� is null"
             zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
             If Not rsTemp.EOF Then
                    'Ӧ���ǲ���Ǽ�
                    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & gintInsure & ",'����ʱ��','" & Format(rsTemp!��Ժ����, "yyyy-mm-dd HH:MM:SS") & "',1)"
                    Call ExecuteProcedure("�������ʱ��")
             End If
        End If
    Else
        '����ʱ�仹ԭ
        '���±����ʻ��������Ϣ
        If g�������_�����山.����ʱ�� <> "" Then
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & gintInsure & ",'����ʱ��','" & g�������_�����山.����ʱ�� & "',1)"
            Call ExecuteProcedure("�������ʱ��")
        End If
    End If
    
    'ȡ�����ʻ��еľ�����
     gstrSQL = "Select ������,����ʱ�� From �����ʻ�  where ����id=" & mlng����ID & " and ����=" & gintInsure
     zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
     If rsTemp.RecordCount = 0 Then
         ShowMsgbox "�ڱ����ʻ��в����ڸò���"
         Exit Sub
     End If
    g�������_�����山.������ = Nvl(rsTemp!������)
    g�������_�����山.����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    g�������_�����山.lng����ID = mlng����ID
    
    '���±����ʻ��������Ϣ
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & gintInsure & ",'֧�����','''" & g�������_�����山.֧����� & "''')"
    Call ExecuteProcedure("���������")
    
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
    
    With rsTemp
    
        .Open "Select * From ֧����� where ��־=2 or ��־=" & IIf(mbytType = 3, 0, IIf(mbytType = 4, 1, mbytType)) & " order by ����", gcnOracle_CQYB
        Do While Not .EOF
            cbo���.AddItem Nvl(!����) & "-" & Nvl(!����)
            If !ȱʡ = 1 Then
                cbo���.ListIndex = cbo���.NewIndex
            End If
            .MoveNext
        Loop
        If cbo���.ListIndex < 0 Then
            If cbo���.ListCount <> 0 Then
                cbo���.ListIndex = 0
            End If
        End If
    End With
    If cbo���.ListCount = 0 Then
        ShowMsgbox "֧�����δ��ʼ��,����ϵͳ����Ա��ϵ!"
        Exit Function
    End If
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
    With g�������_�����山
        lblEdit(2).Caption = .����
        lblEdit(3).Caption = Decode(.�Ա�, "1", "��", "2", "Ů", "δ֪")
        lblEdit(4).Caption = .���֤��
        lblEdit(5).Caption = .��������
        lblEdit(6).Caption = Get��������_�����山(ҽ����Ա���, .ҽ����Ա���)
        lblEdit(7).Caption = .��λ����
        lblEdit(8).Caption = .����
        'Ŀǰû�����
        lblEdit(9).Caption = ""          'Get��������_�����山(ҽ���չ����, .ҽ���չ����)
        lblEdit(10).Caption = .��λ����
        lblEdit(11).Caption = Get��������_�����山(ҽ�Ʋ������, .ҽ�Ʋ������)
        lblEdit(12).Caption = Format(.�ʻ����, "####0.00;#####0.00; ;")
    End With
    
    gstrSQL = "Select ����ID,֧�����,����ʱ�� from �����ʻ� where ҽ����='" & g�������_�����山.���˱�� & "' and ����=" & TYPE_�����山
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��ز���"
    If rsTemp.EOF Then Exit Sub
    g�������_�����山.֧����� = Nvl(rsTemp!֧�����)
    g�������_�����山.����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    gstrSQL = "Select * From ҽ������Ŀ¼ where ID=" & Nvl(rsTemp!����ID, 0)
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.ProductName, "��ȡ������Ϣ", gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle_CQYB
    Call SQLTest
    If rsTemp.EOF Then
        Exit Sub
    End If
    txt����.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
    txt����.Tag = Nvl(rsTemp!ID, 0)
    Dim i As Long
    For i = 0 To cbo���.ListCount - 1
        If InStr(1, cbo���.List(i), g�������_�����山.֧����� & "-") <> 0 Then
            cbo���.ListIndex = i
            Exit For
        End If
    Next
    
End Sub
Private Sub mshSelect_Click()
    With mshSelect
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            SetColumnSort mshSelect, mintPreCol, mintsort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshSelect
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub


'����ͷ��������
Private Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer)
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
               .Sort = flexSortStringNoCaseAscending
               intPreSort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               intPreSort = flexSortStringNoCaseDescending
            End If
            intPreCol = intCol
            .Row = FindRow(mshFilter, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub


Private Sub txt����_Change()
    txt����.Tag = ""
End Sub

Private Sub txt����_GotFocus()
    OpenIme GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "")
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    If KeyCode = vbKeyReturn Then
        If Me.txt���� = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If Trim(txt����) = "" Then Exit Sub
        If Trim(txt����.Tag) <> "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txt���� = UCase(txt����)
    
        Dim rsTemp As New ADODB.Recordset
        
        With rsTemp
            If .State = 1 Then .Close

            gstrSQL = "" & _
                "   Select id, ����, ����, ֧�����, ������, ���ֽ���취, ���칹������ " & _
                "   From ҽ������Ŀ¼" & _
                "   Where " & zlCommFun.GetLike("", "����", Me.txt����) & " Or " & _
                            zlCommFun.GetLike("", "����", Me.txt����) & " Or " & _
                            zlCommFun.GetLike("", "������", Me.txt����)
            
            Call SQLTest(App.ProductName, Me.Caption, strSql)
            .Open gstrSQL, gcnOracle_CQYB
            Call SQLTest
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt����.Top - .Height
                    .Left = txt����.Left + txt����.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = 1000
                    .ColWidth(5) = 1400
                    .ColWidth(6) = 1400
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt���� = "[" & Nvl(!����) & "]" & IIf(IsNull(!����), "", !����)
                txt����.Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
    End If
End Sub

Private Sub txt����_LostFocus()
    OpenIme ""
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            txt����.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
            txt����.Tag = .TextMatrix(.Row, 0)
            If cmdȷ��.Enabled Then cmdȷ��.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub
'Ѱ����ĳһ��Ԫֵ��ȵ���
Private Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .Rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function
