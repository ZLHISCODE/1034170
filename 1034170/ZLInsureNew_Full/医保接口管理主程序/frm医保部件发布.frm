VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmҽ���������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ����������"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frmҽ����������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8130
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdģ�� 
      Caption         =   "ģ��(&M)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6780
      TabIndex        =   10
      Top             =   4470
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3840
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&F)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6780
      TabIndex        =   11
      Top             =   4920
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img���� 
      Left            =   3210
      Top             =   2520
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
            Picture         =   "frmҽ����������.frx":1CFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img�˵� 
      Left            =   840
      Top             =   2550
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
            Picture         =   "frmҽ����������.frx":2F7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdҽ���ӿ� 
      Caption         =   "��"
      Height          =   285
      Left            =   5940
      TabIndex        =   4
      Top             =   960
      Width           =   285
   End
   Begin MSComctlLib.ListView lvw��֧�ֵķ��� 
      Height          =   3675
      Left            =   2850
      TabIndex        =   9
      Top             =   1710
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   6482
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
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6780
      TabIndex        =   13
      Top             =   990
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&P)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6780
      TabIndex        =   12
      Top             =   510
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5745
      Left            =   6480
      TabIndex        =   14
      Top             =   -300
      Width           =   45
   End
   Begin MSComctlLib.ListView lvw�˵� 
      Height          =   3675
      Left            =   0
      TabIndex        =   7
      Top             =   1710
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   6482
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img�˵�"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -180
      TabIndex        =   5
      Top             =   1440
      Width           =   6705
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2115
      MaxLength       =   20
      TabIndex        =   1
      Top             =   570
      Width           =   2775
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   0
      Top             =   180
      Width           =   405
   End
   Begin VB.TextBox txtҽ������ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2115
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   3825
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   5415
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   635
      SimpleText      =   $"frmҽ����������.frx":3DCE
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҽ����������.frx":3E15
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1425
      TabIndex        =   17
      Top             =   240
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1410
      TabIndex        =   16
      Top             =   630
      Width           =   630
   End
   Begin VB.Label lbl��֧�ֵķ��� 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "��֧�ֵķ���"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   1530
      Width           =   3495
   End
   Begin VB.Label lbl�˵� 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "�˵��嵥"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   1530
      Width           =   2745
   End
   Begin VB.Label lblҽ������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ������(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1065
      TabIndex        =   2
      Top             =   1020
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   180
      Picture         =   "frmҽ����������.frx":46A9
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frmҽ����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mrsPrivs As New ADODB.Recordset      '��������Ȩ��
Private mrsModul As New ADODB.Recordset      '��ģ��ʹ�õ��ķ���
Private mrsMethod As New ADODB.Recordset     '���ӿ���ʹ�õ��ķ���

Private Sub cmd����_Click()
    '���±�������Regist.txt
    Dim intInsure As Integer
    Dim strModuls As String
    Dim strFunctions As String
    Dim strPrivs As String
    Dim strȨ�� As String
    Dim strע���� As String
    '���±�������Regist.sql
    Dim strInsert As String
    
    Dim intItem As Integer, intCount As Integer
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    
    intInsure = Val(txt����.Text)
    If intInsure = 0 Then
        MsgBox "����ѡ��ҽ���ӿڲ�����", vbInformation, gstrSysname
        Exit Sub
    End If
    
    '������ѡ���ģ�顢��������������Ȩ�ޣ�����ע��ű���ע����
    If objFileSys.FileExists(txtҽ������.Tag & "\Regist.txt") Then
        If MsgBox("ҽ����������Ŀ¼�У��Ѿ����ڽӿ�ע���ļ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
    End If
    
    '��ȡʹ��ģ���嵥
    intCount = lvw�˵�.ListItems.Count
    For intItem = 1 To intCount
        If lvw�˵�.ListItems(intItem).Checked Then
            strModuls = strModuls & "," & Mid(lvw�˵�.ListItems(intItem).Key, 3)
        End If
    Next
    If strModuls <> "" Then strModuls = Mid(strModuls, 2)
    
    '��ȡ֧�ֵķ����嵥
    With mrsModul
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strFunctions = strFunctions & vbCrLf & !ģ�� & "|" & !Ȩ�޴� & "|" & !����
            .MoveNext
        Loop
    End With
    If strFunctions <> "" Then strFunctions = Mid(strFunctions, 3)
    
    '��ȡ��������Ȩ��
    With mrsPrivs
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "����,����"
        Do While Not .EOF
            For intItem = 1 To 5
                If Mid(!Ȩ��, intItem, 1) = 1 Then
                    Select Case intItem
                    Case 1
                        strȨ�� = "SELECT"
                    Case 2
                        strȨ�� = "INSERT"
                    Case 3
                        strȨ�� = "UPDATE"
                    Case 4
                        strȨ�� = "DELETE"
                    Case 5
                        strȨ�� = "EXECUTE"
                    End Select
                    strPrivs = strPrivs & vbCrLf & !���� & "|" & !���� & "|" & strȨ��
                End If
            Next
            .MoveNext
        Loop
    End With
    If strPrivs <> "" Then strPrivs = Mid(strPrivs, 3)
    
    'дRegist.txt
    Set objStream = objFileSys.CreateTextFile(txtҽ������.Tag & "\Regist.txt", True)
    objStream.WriteLine "[MODULS]"
    objStream.WriteLine strModuls
    objStream.WriteBlankLines 1
    objStream.WriteLine "[FUNCTIONS]"
    objStream.WriteLine strFunctions
    objStream.WriteBlankLines 1
    objStream.WriteLine "[PRIVS]"
    objStream.WriteLine strPrivs
    objStream.Close
    
    'ת��Ϊʵ�ʵ�Ȩ��SQL
'    With mrsModul
'        If .RecordCount <> 0 Then .MoveFirst
'        Do While Not .EOF
'            mrsPrivs.Filter = "����='" & !���� & "'"
'            Do While Not mrsPrivs.EOF
'                For intItem = 1 To 5
'                    If Mid(mrsPrivs!Ȩ��, intItem, 1) = 1 Then
'                        Select Case intItem
'                        Case 1
'                            strȨ�� = "SELECT"
'                        Case 2
'                            strȨ�� = "INSERT"
'                        Case 3
'                            strȨ�� = "UPDATE"
'                        Case 4
'                            strȨ�� = "DELETE"
'                        Case 5
'                            strȨ�� = "EXECUTE"
'                        End Select
'                        strInsert = strInsert & vbCrLf & "Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100," & _
'                        !ģ�� & ",'" & !Ȩ�޴� & "','USER','" & mrsPrivs!���� & "','" & strȨ�� & "');"
'                    End If
'                Next
'                mrsPrivs.MoveNext
'            Loop
'            mrsPrivs.Filter = 0
'            .MoveNext
'        Loop
'    End With
'    If strInsert <> "" Then strInsert = Mid(strInsert, 3)
    
    '�õ��������ݵĲ���SQL���
    intCount = lvw�˵�.ListItems.Count
    For intItem = 1 To intCount
        If lvw�˵�.ListItems(intItem).Checked Then
            strInsert = strInsert & vbCrLf & _
                "Insert into zlInsureModuls(����,���) Values (" & intInsure & "," & _
                Mid(lvw�˵�.ListItems(intItem).Key, 3) & ");"
        End If
    Next
    With mrsModul
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strInsert = strInsert & vbCrLf & _
                "Insert into zlInsureFuncs(����,���,����,����) Values (" & intInsure & "," & _
                !ģ�� & ",'" & !Ȩ�޴� & "','" & !���� & "');"
            .MoveNext
        Loop
    End With
    With mrsPrivs
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            For intItem = 1 To 5
                If Mid(mrsPrivs!Ȩ��, intItem, 1) = 1 Then
                    Select Case intItem
                    Case 1
                        strȨ�� = "SELECT"
                    Case 2
                        strȨ�� = "INSERT"
                    Case 3
                        strȨ�� = "UPDATE"
                    Case 4
                        strȨ�� = "DELETE"
                    Case 5
                        strȨ�� = "EXECUTE"
                    End Select
                    strInsert = strInsert & vbCrLf & _
                        "Insert Into zlInsurePrivs(����,����,����,Ȩ��) Values(" & intInsure & "," & _
                        "'" & !���� & "','" & !���� & "','" & strȨ�� & "');"
                End If
            Next
            .MoveNext
        Loop
    End With
    If strInsert <> "" Then strInsert = Mid(strInsert, 3)
    
    'дRegist.sql��ʵ�ʵ�Ȩ�޽ű�
    Set objStream = objFileSys.CreateTextFile(txtҽ������.Tag & "\Regist.sql", True)
    objStream.WriteLine strInsert
    objStream.Close
    
    MsgBox "ע���ļ��Ѿ�������", vbInformation, gstrSysname
End Sub

Private Sub cmd����_Click()
    If Val(txt����.Text) = 0 Then
        MsgBox "����ѡ��ҽ���ӿڲ�����", vbInformation, gstrSysname
        Exit Sub
    End If
    Call MakeMethodRecord
    Call frmȨ������.ShowEditor(mrsPrivs, mrsMethod)
End Sub

Private Sub cmdģ��_Click()
    If Val(txt����.Text) = 0 Then
        MsgBox "����ѡ��ҽ���ӿڲ�����", vbInformation, gstrSysname
        Exit Sub
    End If
    Call MakeMethodRecord
    Call frm��������.ShowEditor(mrsModul, mrsMethod)
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdҽ���ӿ�_Click()
    Dim strMessage As String
    Dim strFile As String, strPath As String
    Dim arrMessage
    Dim objTest As Object
    
    cmd����.Enabled = False
    cmdģ��.Enabled = False
    cmd����.Enabled = False
    
    With dlg
        .Filter = "ҽ������(*.dll)|*.dll"
        .ShowOpen
        Call GetFileOrPath(.FileName, strFile, strPath)
        strFile = Mid(strFile, 1, Len(strFile) - 4)
        txtҽ������.Tag = strPath
    End With

    '1��
    If Mid(strFile, 1, 5) <> "ZL9I_" Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������1", vbInformation, gstrSysname
        Exit Sub
    End If
    '2��
    On Error Resume Next
    Err = 0
    Set objTest = CreateObject(strFile & ".CLS" & Mid(strFile, 4))
    If Err <> 0 Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������2", vbInformation, gstrSysname
        Exit Sub
    End If
    '3��
    Err = 0
    strMessage = objTest.I_RegInfo
    If Err <> 0 Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������3", vbInformation, gstrSysname
        Set objTest = Nothing
        Exit Sub
    End If
    
    arrMessage = Split(strMessage, "|")
    If Not (UBound(arrMessage) >= 1) Then
        MsgBox "��ѡ��Ϸ���ҽ���ӿڲ������������3.1", vbInformation, gstrSysname
        Exit Sub
    End If
    
    If Val(arrMessage(0)) = 0 Then
        MsgBox "ҽ���ӿڵ����಻��Ϊ�գ�", vbInformation, gstrSysname
        Exit Sub
    End If
    If Trim(UCase(arrMessage(1))) = "" Then
        MsgBox "ҽ���ӿڵ����Ʋ���Ϊ�գ�", vbInformation, gstrSysname
        Exit Sub
    End If
    
    txt����.Text = Val(arrMessage(0))
    txt����.Text = UCase(arrMessage(1))
    txtҽ������.Text = UCase(strFile) & ".DLL"
    
    cmd����.Enabled = True
    cmdģ��.Enabled = True
    cmd����.Enabled = True
    
    Call ShowRegist
    Exit Sub
ErrHand:
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim intItem As Integer, intCount As Integer
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    '����˵���ϵ
    mstrSQL = "Select ���,���� From zlPrograms Where Upper(����)='ZL9INSURE' Order By ���"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "��ȡ�˵���ϵ")
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvw�˵�.ListItems.Add(, "K_" & rsTemp!���, rsTemp!����, , 1)
            lvwItem.Checked = True
            .MoveNext
        Loop
    End With
    
    '���뷽���б��̶���
    With lvw��֧�ֵķ���
        Call .ListItems.Add(, "K_" & ����.�����֤, "Identify()", , 1)
        Call .ListItems.Add(, "K_" & ����.�����֤_�����Һ�, "Identify2()", , 1)
        Call .ListItems.Add(, "K_" & ����.�ʻ����, "SelfBalance()", , 1)
        Call .ListItems.Add(, "K_" & ����.����Һ�, "RegistSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.����Һ�����, "RegistDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.�����������, "ClinicPreSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.�������, "ClinicSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.�����������, "ClinicDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.�����ʻ�תԤ��, "TransferSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.Ԥ���˸����ʻ�, "TransferDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.סԺ�������, "WipeoffMoney()", , 1)
        Call .ListItems.Add(, "K_" & ����.סԺ����, "SettleSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.סԺ��������, "SettleDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.��Ժ�Ǽ�, "ComeInSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.��Ժ�Ǽǳ���, "ComeInDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.��Ժ�Ǽ�, "LeaveSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.��Ժ�Ǽǳ���, "LeaveDelSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.������ϸ�ϴ�, "TranChargeDetail()", , 1)
        Call .ListItems.Add(, "K_" & ����.סԺ��Ϣ�䶯, "ModiPatiSwap()", , 1)
        Call .ListItems.Add(, "K_" & ����.��ȡҽ����Ŀ��Ϣ, "GetItemInfo()", , 1)
        Call .ListItems.Add(, "K_" & ����.����ѡ��, "ChooseDisease()", , 1)
        
        For intItem = 1 To ����
            Set lvwItem = .ListItems(intItem)
            lvwItem.Checked = True
        Next
    End With
    
End Sub

Private Sub lvw�˵�_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '��������ѡ
    If Item.Key = "K_1600" Then Item.Checked = True
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txtҽ������_GotFocus()
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

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

Private Sub MakeMethodRecord()
    Dim intItem As Integer, intCount As Integer
    Dim strMethod As String
    Dim strField As String, strValue As String
    
    strField = "���," & adDouble & ",18|����," & adLongVarChar & ",50"
    Call Record_Init(mrsMethod, strField)
    
    strField = "���|����"
    intCount = lvw��֧�ֵķ���.ListItems.Count
    For intItem = 1 To intCount
        If lvw��֧�ֵķ���.ListItems(intItem).Checked Then
            strMethod = lvw��֧�ֵķ���.ListItems(intItem).Text
            strMethod = Mid(strMethod, 1, Len(strMethod) - 2)
            strValue = Mid(lvw��֧�ֵķ���.ListItems(intItem).Key, 3) & "|" & strMethod
            Call Record_Add(mrsMethod, strField, strValue)
        End If
    Next
End Sub

Private Sub ShowRegist()
    Dim strTemp As String
    Dim strBase As String
    Dim strFunctions As String
    Dim strPrivs As String
    Dim arrData
    Dim intItem As Integer, intCount As Integer
    Dim strFields As String
    Dim str���� As String, str���� As String, strȨ�� As String
    On Error GoTo ErrHand
    
    '��ʼ��Ȩ�޼�¼��
    strFields = "����," & adLongVarChar & "," & 50 & "|����," & adLongVarChar & "," & 50 & "|Ȩ��," & adLongVarChar & "," & 5
    Call Record_Init(mrsPrivs, strFields)
    strFields = "ģ��," & adDouble & "," & 18 & "|Ȩ�޴�," & adLongVarChar & "," & 50 & "|����," & adLongVarChar & "," & 50
    Call Record_Init(mrsModul, strFields)
    
    If Not ReadFile(strBase, strFunctions, strPrivs) Then Exit Sub
    
    '�������ѡ��
    intCount = lvw�˵�.ListItems.Count
    For intItem = 1 To intCount
        lvw�˵�.ListItems(intItem).Checked = False
    Next
    intCount = lvw��֧�ֵķ���.ListItems.Count
    For intItem = 1 To intCount
        lvw��֧�ֵķ���.ListItems(intItem).Checked = False
    Next
    
    '����ע���ļ���ʾ
    '�˵�
    arrData = Split(strBase, ",")
    intCount = UBound(arrData)
    For intItem = 0 To intCount
        lvw�˵�.ListItems("K_" & arrData(intItem)).Checked = True
    Next
    
    '����
    arrData = Split(strFunctions, vbCrLf)
    intCount = UBound(arrData)
    For intItem = 0 To intCount
        If InStr(1, strTemp & ",", "," & UCase(Split(arrData(intItem), "|")(2)) & ",") = 0 Then
            strTemp = strTemp & "," & UCase(Split(arrData(intItem), "|")(2))
            Select Case UCase(Split(arrData(intItem), "|")(2))
            Case "IDENTIFY"
                lvw��֧�ֵķ���.ListItems("K_" & ����.�����֤).Checked = True
            Case "IDENTIFY2"
                lvw��֧�ֵķ���.ListItems("K_" & ����.�����֤_�����Һ�).Checked = True
            Case "SELFBALANCE"
                lvw��֧�ֵķ���.ListItems("K_" & ����.�ʻ����).Checked = True
            Case "REGISTSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.����Һ�).Checked = True
            Case "REGISTDELSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.����Һ�����).Checked = True
            Case "CLINICPRESWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.�����������).Checked = True
            Case "CLINICSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.�������).Checked = True
            Case "CLINICDELSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.�����������).Checked = True
            Case "TRANSFERSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.�����ʻ�תԤ��).Checked = True
            Case "TRANSFERDELSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.Ԥ���˸����ʻ�).Checked = True
            Case "WIPEOFFMONEY"
                lvw��֧�ֵķ���.ListItems("K_" & ����.סԺ�������).Checked = True
            Case "SETTLESWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.סԺ����).Checked = True
            Case "SETTLEDELSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.סԺ��������).Checked = True
            Case "COMEINSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.��Ժ�Ǽ�).Checked = True
            Case "COMEINDELSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.��Ժ�Ǽǳ���).Checked = True
            Case "LEAVESWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.��Ժ�Ǽ�).Checked = True
            Case "LEAVEDELSWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.��Ժ�Ǽǳ���).Checked = True
            Case "TRANCHARGEDETAIL"
                lvw��֧�ֵķ���.ListItems("K_" & ����.������ϸ�ϴ�).Checked = True
            Case "MODIPATIDWAP"
                lvw��֧�ֵķ���.ListItems("K_" & ����.סԺ��Ϣ�䶯).Checked = True
            Case "GETITEMINFO"
                lvw��֧�ֵķ���.ListItems("K_" & ����.��ȡҽ����Ŀ��Ϣ).Checked = True
            Case "CHOOSEDISEASE"
                lvw��֧�ֵķ���.ListItems("K_" & ����.����ѡ��).Checked = True
            End Select
        End If
    Next
    
    'Ȩ��
    For intItem = 0 To intCount
        Call Record_Add(mrsModul, "ģ��|Ȩ�޴�|����", arrData(intItem))
    Next
    arrData = Split(strPrivs, vbCrLf)
    intCount = UBound(arrData)
    strȨ�� = "00000"
    For intItem = 0 To intCount
        If (str���� <> Split(arrData(intItem), "|")(0) Or str���� <> Split(arrData(intItem), "|")(1)) Then
            If str���� <> "" Then Call Record_Add(mrsPrivs, "����|����|Ȩ��", str���� & "|" & str���� & "|" & strȨ��)
            str���� = Split(arrData(intItem), "|")(0)
            str���� = Split(arrData(intItem), "|")(1)
            strȨ�� = "00000"
        End If
        
        Select Case UCase(Split(arrData(intItem), "|")(2))
        Case "SELECT"
            strȨ�� = "1" & Mid(strȨ��, 2)
        Case "INSERT"
            strȨ�� = Mid(strȨ��, 1, 1) & "1" & Mid(strȨ��, 3)
        Case "UPDATE"
            strȨ�� = Mid(strȨ��, 1, 2) & "1" & Mid(strȨ��, 4)
        Case "DELETE"
            strȨ�� = Mid(strȨ��, 1, 3) & "1" & Mid(strȨ��, 5)
        Case Else
            strȨ�� = Mid(strȨ��, 1, 4) & "1"
        End Select
    Next
    Call Record_Add(mrsPrivs, "����|����|Ȩ��", str���� & "|" & str���� & "|" & strȨ��)
    
    Exit Sub
    
ErrHand:
    MsgBox "װ��ע���ļ�ʱ����δ֪����", vbInformation, gstrSysname
End Sub

Private Function ReadFile(strBase As String, strFunctions As String, strPrivs As String) As Boolean
    Dim intState As Integer
    Dim strLine As String
    Dim strPath As String
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    Const strRegist As String = "Regist.txt"
    '�����ļ�
    strPath = txtҽ������.Tag & "\" & strRegist
    If Not objFileSys.FileExists(strPath) Then Exit Function
    Set objStream = objFileSys.OpenTextFile(strPath, ForReading)
    Do While Not objStream.AtEndOfStream
        strLine = UCase(objStream.ReadLine)
        Select Case strLine
        Case "[MODULS]"
            intState = 1
        Case "[FUNCTIONS]"
            intState = 2
        Case "[PRIVS]"
            intState = 3
        Case Else
            If Trim(strLine) <> "" Then
                Select Case intState
                Case 1  'MODULS
                    strBase = strLine
                Case 2  'FUNCTIONS
                    strFunctions = strFunctions & IIf(strFunctions = "", "", vbCrLf) & strLine
                Case 3  'PRIVS
                    strPrivs = strPrivs & IIf(strPrivs = "", "", vbCrLf) & strLine
                End Select
            End If
        End Select
    Loop
    
    strBase = Trim(strBase)
    strFunctions = Trim(strFunctions)
    strPrivs = Trim(strPrivs)
    
    objStream.Close
    ReadFile = Not (strBase = "" Or strFunctions = "" Or strPrivs = "")
End Function
