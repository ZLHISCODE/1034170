VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillDiscard 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ʊ�ݱ���"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBillDiscard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.OptionButton opt��Χ 
      Caption         =   "���ű���(&M)"
      Height          =   240
      Index           =   1
      Left            =   3300
      TabIndex        =   19
      Top             =   1410
      Width           =   1635
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "���ű���(&S)"
      Height          =   240
      Index           =   0
      Left            =   1545
      TabIndex        =   18
      Top             =   1410
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1530
      TabIndex        =   6
      Top             =   870
      Width           =   1815
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      Index           =   2
      Left            =   4110
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1875
      Width           =   1485
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      Index           =   1
      Left            =   1860
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1875
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   4875
      TabIndex        =   10
      Top             =   2430
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   110690307
      CurrentDate     =   37007
   End
   Begin VB.ComboBox cmb������ 
      Height          =   360
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2415
      Width           =   1830
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   3630
      TabIndex        =   13
      Top             =   5430
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   4980
      TabIndex        =   14
      Top             =   5430
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -270
      TabIndex        =   12
      Top             =   5160
      Width           =   7065
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   270
      TabIndex        =   15
      Top             =   5430
      Width           =   1200
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Index           =   2
      Left            =   3780
      TabIndex        =   17
      Top             =   1875
      Width           =   315
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Index           =   1
      Left            =   1530
      TabIndex        =   16
      Top             =   1875
      Width           =   315
   End
   Begin VB.Label lbl˵�� 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   2085
      Left            =   150
      TabIndex        =   11
      Top             =   2940
      Width           =   6300
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   240
      Index           =   5
      Left            =   3450
      TabIndex        =   5
      Top             =   1935
      Width           =   240
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���뷶Χ(&B)"
      Height          =   240
      Index           =   6
      Left            =   150
      TabIndex        =   2
      Top             =   1935
      Width           =   1320
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ʊ������"
      Height          =   240
      Index           =   4
      Left            =   510
      TabIndex        =   1
      Top             =   930
      Width           =   960
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��(&D)"
      Height          =   240
      Index           =   3
      Left            =   3495
      TabIndex        =   9
      Top             =   2490
      Width           =   1320
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������(&G)"
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   7
      Top             =   2475
      Width           =   1080
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ʊ�ݱ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2130
      TabIndex        =   0
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "frmBillDiscard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnChange As Boolean     'Ϊ��ʱ��ʾ�Ѹı���
Dim mdatCurrnet As Date
Dim mstrID As String
Dim mstrǰ׺ As String
Dim mstr��С���� As String
Dim mstr������ As String
Private mstrPrivs As String
Private mlngƱ�ݳ��� As Long

Private Sub InitContext()
    dtpDate.Value = mdatCurrnet
    dtpDate.MaxDate = mdatCurrnet
    
    txtEdit(0).Text = frmBillSupervise.lvwMain.SelectedItem.Text
    txtEdit(0).Tag = Mid(frmBillSupervise.lvwMain.SelectedItem.Key, 2)

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    mblnChange = True
    Select Case txtEdit(0).Tag
        Case 1      '1-�շ��վ�
            gstrSQL = " And B.��Ա����='�����շ�Ա'"
        Case 2      '2-Ԥ���վ�
            gstrSQL = " And B.��Ա���� in ('Ԥ���տ�Ա','��Ժ�Ǽ�Ա')"
        Case 3      '3-�����վ�
            gstrSQL = " And B.��Ա����='סԺ����Ա'"
        Case 4      '4-�Һ��վ�
            gstrSQL = " And B.��Ա����='����Һ�Ա'"
        Case 5      '5-���￨
            gstrSQL = " And B.��Ա���� in ('�����Ǽ���','��Ժ�Ǽ�Ա')"
        Case Else
            Exit Sub
    End Select
    gstrSQL = "Select A.���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID " & gstrSQL & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) order by A.����"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    cmb������.Clear
    Do Until rsTemp.EOF
        cmb������.AddItem rsTemp("����")
        rsTemp.MoveNext
    Loop
    If cmb������.ListCount > 0 Then cmb������.ListIndex = 0

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmb������_Click()
    mblnChange = True
End Sub

Private Sub cmb������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub dtpDate_Change()
    mblnChange = True
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If ValidateContent() = False Then Exit Sub
    If MsgBox("Ʊ��һ������󣬱������Ͳ�����ʹ���ˡ�" & vbCrLf & "�Ƿ�ȷ��Ҫ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If Save() = False Then Exit Sub
    
    '�޸�
    Call frmBillSupervise.ShowItem(frmBillSupervise.lvw����_S.SelectedItem)
    frmBillSupervise.Fill����
    mblnChange = False
    Unload Me
End Sub

Private Sub opt��Χ_Click(Index As Integer)
    mblnChange = True
    If opt��Χ(0).Value = True Then
        txtEdit(2).Enabled = False
        txtEdit(2).Text = txtEdit(1).Text
    Else
        txtEdit(2).Enabled = True
    End If
    Call ShowSum
End Sub

Private Sub opt��Χ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 And opt��Χ(0).Value = True Then txtEdit(2).Text = txtEdit(1).Text
    Call ShowSum
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    SelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
    If (Index = 1 Or Index = 2) And (KeyAscii >= vbKey0 Or KeyAscii <= vbKey9) And txtEdit(Index).SelLength = 0 Then
        If Len(txtEdit(Index)) >= mlngƱ�ݳ��� Then KeyAscii = 0
    End If
End Sub

Private Function ValidateContent() As Boolean
'����:����������ݵ��Ƿ���Ч
'����:��Ч�򷵻�True,���򷵻�False
    Dim lngCount As Long, i As Integer
    Dim strTemp As String
    
    ValidateContent = False
    '�ַ������
    For lngCount = 1 To 2
        txtEdit(lngCount).Text = Trim(txtEdit(lngCount).Text)
        If StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            txtEdit(lngCount).SetFocus
            SelAll txtEdit(lngCount)
            Exit Function
        End If
        For i = 1 To Len(txtEdit(lngCount).Text)
            strTemp = Mid(txtEdit(lngCount), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox "�����к��з������ַ���", vbExclamation, gstrSysName
                txtEdit(lngCount).SetFocus
                SelAll txtEdit(lngCount)
                Exit Function
            End If
        Next
        If Len(txtEdit(lngCount).Text) <> Len(txtEdit(lngCount).Tag) - Len(mstrǰ׺) Then
            MsgBox "����ĳ��Ȳ��ԡ�", vbExclamation, gstrSysName
            txtEdit(lngCount).SetFocus
            SelAll txtEdit(lngCount)
            Exit Function
        End If
    Next
    
    If mstrǰ׺ & txtEdit(1).Text < txtEdit(1).Tag Then
        MsgBox "���ϵĿ�ʼ�������������õĿ�ʼ���롣", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        SelAll txtEdit(1)
        Exit Function
    End If
    If txtEdit(2).Enabled = True Then
        If mstrǰ׺ & txtEdit(2).Text > txtEdit(2).Tag Then
            MsgBox "���ϵ���ֹ�������С�����õ���ֹ���롣", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            SelAll txtEdit(2)
            Exit Function
        End If
    Else
        If mstrǰ׺ & txtEdit(1).Text > txtEdit(2).Tag Then
            MsgBox "���ϵĺ������С�����õ���ֹ���롣", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            SelAll txtEdit(1)
            Exit Function
        End If
    End If
        
    If txtEdit(1).Text > txtEdit(2).Text Then
        MsgBox "���ϵĿ�ʼ�������С�����ϵ���ֹ���롣", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        SelAll txtEdit(1)
        Exit Function
    End If
    If Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 > 10000 Then
        MsgBox "һ�����ϵ����������ܳ���һ���š�", vbExclamation, gstrSysName
        txtEdit(2).SetFocus
        SelAll txtEdit(2)
        Exit Function
    End If
    If mstr��С���� <> "" Then
        If mstrǰ׺ & txtEdit(1).Text <= mstr��С���� And mstr��С���� <= mstrǰ׺ & txtEdit(2).Text Or _
                mstrǰ׺ & txtEdit(1).Text <= mstr������ And mstr������ <= mstrǰ׺ & txtEdit(2).Text Then
            MsgBox "���ϵĺ����а������Ѿ�ʹ�õ� ��", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            SelAll txtEdit(1)
            Exit Function
        End If
        If mstrǰ׺ & txtEdit(1).Text > mstr��С���� And mstrǰ׺ & txtEdit(2).Text < mstr������ Then
            If MsgBox("���ϵĺ����п��ܰ������Ѿ�ʹ�õģ��Ƿ������", vbYesNo Or vbQuestion Or vbDefaultButton2, gstrSysName) = vbNo Then
                txtEdit(1).SetFocus
                SelAll txtEdit(1)
                Exit Function
            End If
        End If
    End If
    If cmb������.Text = "" Then
        MsgBox "�����˲���Ϊ�ա�", vbExclamation, gstrSysName
        cmb������.SetFocus
        Exit Function
    End If
    
    ValidateContent = True
End Function

Private Function Save() As Boolean
'����:����༭������
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim strTemp As String
    Dim lngID As Long
    
    On Error GoTo errHandle
    Save = False
    
    '�޸�
    gstrSQL = "zl_Ʊ��ʹ����ϸ_damage(" & mstrID & "," & txtEdit(0).Tag & _
        ",'" & mstrǰ׺ & "','" & txtEdit(1).Text & "','" & txtEdit(2).Text & _
        "',to_date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),'" & cmb������.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    If gblnBillPrint Then
        Call gobjBillPrint.zlDiscardBill(mstrID, Val(txtEdit(0).Tag), mstrǰ׺, txtEdit(1).Text, txtEdit(2).Text, dtpDate.Value, cmb������.Text)
    End If
    
    Save = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowSum()
'����:��ʾ������Ϣ
    Dim strTemp As String
    strTemp = " ���ϵĿ�ʼ���룺" & lbl(1).Caption & txtEdit(1).Text & vbCrLf
    strTemp = strTemp & "  ���ϵĽ������룺" & lbl(2).Caption & txtEdit(2).Text & vbCrLf
    If txtEdit(1).Text = "" Or txtEdit(2).Text = "" Then
        strTemp = strTemp & "  ���ϵ�Ʊ����������" & vbCrLf & vbCrLf
    Else
        strTemp = strTemp & "  ���ϵ�Ʊ����������" & Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 & vbCrLf & vbCrLf
    End If
    strTemp = strTemp & "  ���õĿ�ʼ���룺" & Replace(txtEdit(1).Tag, "&", "&&") & vbCrLf
    strTemp = strTemp & "  ���õĽ������룺" & Replace(txtEdit(2).Tag, "&", "&&") & vbCrLf
    If mstr��С���� <> "" Then
        strTemp = strTemp & "  �Ѿ�ʹ�õ���С���룺" & Replace(mstr��С����, "&", "&&") & vbCrLf
        strTemp = strTemp & "  �Ѿ�ʹ�õ������룺" & Replace(mstr������, "&", "&&") & vbCrLf
    End If
    
    lbl˵��.Caption = strTemp
End Sub

Public Function �༭Ʊ�ݱ���(ByVal strPrivs As String, ByVal strID As String) As Boolean
'����:��������õĲ����ش��ڽ���ͨѶ�ĳ���,�������ӽɿ��¼
'����:str�ɿ���     �ɿ��˵�����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim dblCount As Double
    
    On Error GoTo errHandle
        
    mstrPrivs = strPrivs
    
    mdatCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mstrID = strID
    
    Call InitContext
    
    gstrSQL = "Select ������,ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����,ʹ�÷�ʽ " & _
            " From Ʊ�����ü�¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrID)
    
    mstrǰ׺ = IIf(IsNull(rsTemp("ǰ׺�ı�")), "", rsTemp("ǰ׺�ı�"))
    lbl(1).Caption = Replace(mstrǰ׺, "&", "&&")
    lbl(2).Caption = lbl(1).Caption
    txtEdit(1).Tag = rsTemp("��ʼ����")
    txtEdit(2).Text = Mid(rsTemp("��ֹ����"), Len(mstrǰ׺) + 1)
    mlngƱ�ݳ��� = Len(Mid(rsTemp("��ֹ����"), Len(mstrǰ׺) + 1))
    txtEdit(2).Tag = rsTemp("��ֹ����")
    If IsNull(rsTemp("��ǰ����")) Then
        txtEdit(1).Text = Mid(rsTemp("��ʼ����"), Len(mstrǰ׺) + 1)
    Else
        '�Ѿ�ʹ�ã��Ͱ����ֵ��һ
        dblCount = Val(Mid(rsTemp("��ǰ����"), Len(mstrǰ׺) + 1))
        dblCount = dblCount + 1
        txtEdit(1).Text = Format(dblCount, String(Len(txtEdit(2).Text), "0"))
    End If
    
    On Error Resume Next
    If Val(rsTemp!ʹ�÷�ʽ) = 2 Then    '����ʽ��,ֻ��ѡ��Ϊ������Ա:35846
  
        cmb������.Text = UserInfo.����
    Else
        cmb������.Text = rsTemp("������")
    End If
    If Err <> 0 Then
        If Val(rsTemp!ʹ�÷�ʽ) = 2 Then
            cmb������.AddItem UserInfo.����
            cmb������.ListIndex = cmb������.NewIndex
        Else
            cmb������.AddItem rsTemp("������")
            cmb������.ListIndex = cmb������.NewIndex
        End If
    End If
    If InStr(mstrPrivs, "���в���Ա") = 0 Then cmb������.Enabled = False
    On Error GoTo errHandle
    
    gstrSQL = "select nvl(min(����),' ') as ��С����,nvl(max(����),' ')  as ������ from Ʊ��ʹ����ϸ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrID)
    
    mstr��С���� = Trim(rsTemp("��С����"))
    mstr������ = Trim(rsTemp("������"))
    Call opt��Χ_Click(0)
    
    mblnChange = False
    frmBillDiscard.Show vbModal, frmBillSupervise
    �༭Ʊ�ݱ��� = True
    Exit Function
errHandle:
    MsgBox "���ݶ���ʧ�ܡ�", vbExclamation, gstrSysName
    �༭Ʊ�ݱ��� = False
End Function
