VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLabSamplingFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabSamplingFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   3930
      TabIndex        =   19
      Top             =   2925
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   5370
      TabIndex        =   20
      Top             =   2910
      Width           =   1095
   End
   Begin VB.Frame fraFilter 
      Height          =   2775
      Left            =   75
      TabIndex        =   21
      Top             =   75
      Width           =   6855
      Begin VB.ListBox lst���� 
         Height          =   2085
         ItemData        =   "frmLabSamplingFilter.frx":000C
         Left            =   4815
         List            =   "frmLabSamplingFilter.frx":0013
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   540
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Height          =   45
         Left            =   60
         TabIndex        =   24
         Top             =   2010
         Width           =   4785
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   60
         TabIndex        =   23
         Top             =   1560
         Width           =   4785
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   60
         TabIndex        =   22
         Top             =   990
         Width           =   4785
      End
      Begin VB.ComboBox cboCapture 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1140
         Width           =   1395
      End
      Begin VB.ComboBox cboSample 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1140
         Width           =   1365
      End
      Begin VB.CheckBox chkOutPatient 
         Caption         =   "����"
         Height          =   255
         Left            =   990
         TabIndex        =   13
         Top             =   1710
         Width           =   795
      End
      Begin VB.CheckBox chkInpatient 
         Caption         =   "סԺ"
         Height          =   255
         Left            =   2145
         TabIndex        =   14
         Top             =   1710
         Width           =   795
      End
      Begin VB.CheckBox chkPhysical 
         Caption         =   "���"
         Height          =   255
         Left            =   3300
         TabIndex        =   15
         Top             =   1710
         Width           =   795
      End
      Begin VB.TextBox TxtID 
         Height          =   285
         Left            =   990
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox TxtSickCard 
         Height          =   285
         Left            =   3300
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtName 
         Height          =   285
         Left            =   990
         TabIndex        =   5
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox TxtNo 
         Height          =   285
         Left            =   3300
         TabIndex        =   7
         Top             =   630
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   2370
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         Format          =   202571777
         CurrentDate     =   39034
      End
      Begin MSComCtl2.DTPicker DTPEND 
         Height          =   285
         Left            =   2910
         TabIndex        =   18
         Top             =   2370
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         Format          =   202571777
         CurrentDate     =   39034
      End
      Begin VB.Label lbl�������� 
         Caption         =   "��������"
         Height          =   195
         Left            =   4860
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   2100
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�ɼ���ʽ"
         Height          =   195
         Left            =   2460
         TabIndex        =   10
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��         ��"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Դ"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">>>>>>"
         Height          =   195
         Left            =   2100
         TabIndex        =   17
         Top             =   2415
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ��(&1)"
         Height          =   180
         Left            =   150
         TabIndex        =   0
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨(&2)"
         Height          =   180
         Left            =   2460
         TabIndex        =   2
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��(&3)"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   675
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�(&4)"
         Height          =   180
         Left            =   2460
         TabIndex        =   6
         Top             =   675
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmLabSamplingFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDateOldEnd As Date                                            '��¼�ɵ�ʱ��
Private mstrFilter As String                                            '�����ִ�
Private Enum mFilter
    ��ʶ�� = 0
    ���￨
    ����
    ���ݺ�
    �걾
    �ɼ���ʽ
    ����
    סԺ
    ���
    ���ʱ��
    ���ͻ����ʱ��          '=0 ����ʱ�� = 1 ���ʱ��
    ��ʼʱ��
    ����ʱ��
    ��������
End Enum


Private Sub cmdOK_Click()
    Dim dateSpace As Integer
    Dim strFilter As String                             '���������ִ�
    Dim i As Integer, str���� As String
    
    dateSpace = DateDiff("d", Me.DTPBegin.Value, Me.DTPEND.Value)
    
    If dateSpace > 30 Then
        If MsgBox("��ѡ���ʱ�����30�죬���ܵ��²�������Ƿ������", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    strFilter = Me.TxtID & ";" & TxtSickCard & ";" & TxtName & ";" & TxtNo & ";" & Mid(cboSample, InStr(1, cboSample, "-") + 1) & _
                ";" & cboCapture.ItemData(cboCapture.ListIndex) & ";" & IIf(chkOutPatient, 1, "") & ";" & _
                IIf(chkInpatient, 2, "") & ";" & IIf(chkPhysical, 4, "") & ";" & _
                dateSpace & ";0;" & _
                IIf(mDateOldEnd <> DTPEND.Value, DTPBegin.Value, "") & ";" & _
                IIf(mDateOldEnd <> DTPEND.Value, DTPEND.Value, "")
    For i = 0 To lst����.ListCount - 1
        If lst����.Selected(i) Then
            str���� = str���� & "," & lst����.List(i)
        End If
    Next
    strFilter = strFilter & ";" & str����
    
    zlDatabase.SetPara "�ɼ�����վ����", strFilter, 100, 1211
    '���������������
    mstrFilter = strFilter
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mstrFilter = ""
    Unload Me
End Sub
Private Sub DTPBegin_Change()
    If Me.DTPBegin > Me.DTPEND Then
        Me.DTPBegin = Me.DTPEND
    End If
End Sub

Private Sub DTPEND_Change()
    If Me.DTPEND < Me.DTPBegin Then
        Me.DTPEND = Me.DTPBegin
    End If
End Sub

Private Sub Form_Load()
    InitinterFace
End Sub

Private Sub InitinterFace()
    '��ʹ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer                          'ѭ������
    Dim strTmp As String                            '��ʱ�ִ�����
    Dim varFilter As Variant                        '�����ִ��ֽ�
    Dim NowDate As Date                             '��ǰʱ��
    Dim strBloodType As String                      '��Ѫ��������
    Dim strOldType As String                        '�ϰ�LIS�걾����
    Dim objLisInsideComm As Object                  '�°�LIS�ӿڲ���
    
    On Error GoTo errH
    
    strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "")
    strBloodType = zlDatabase.GetPara(273, 100)
    If strTmp <> "" Then
        varFilter = Split(strTmp, ";")
        Me.chkOutPatient = IIf(Val(varFilter(mFilter.����)) = 0, 0, 1)
        Me.chkInpatient = IIf(Val(varFilter(mFilter.סԺ)) = 0, 0, 1)
        Me.chkPhysical = IIf(Val(varFilter(mFilter.���)) = 0, 0, 1)
    Else
        Me.chkOutPatient = 1
        Me.chkInpatient = 1
        Me.chkPhysical = 1
    End If
    
    mDateOldEnd = Me.DTPEND.Value
    
    '===�������ڿ���
'    strSQL = "Select Distinct A.ID,A.����,A.����,B.�������" & _
'        " From ���ű� A,��������˵�� B" & _
'        " Where A.ID=B.����ID And B.�������� IN('�ٴ�','����')" & _
'        " And B.������� IN(3,[1],[2])" & _
'        " And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
'        " Order by A.����"
'
'
'    Set rsTmp =zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(chkOutPatient.Value = 1 Or chkPhysical.Value = 1, 1, -1), IIf(chkInpatient.Value = 1, 2, -1))
'
'    cboPatientDept.Clear
'    cboPatientDept.AddItem "���п���"
'    cboPatientDept.ItemData(cboPatientDept.NewIndex) = 0
'    cboPatientDept.ListIndex = 0
'    Do Until rsTmp.EOF
'        cboPatientDept.AddItem rsTmp!���� & "-" & rsTmp!����
'        cboPatientDept.ItemData(cboPatientDept.NewIndex) = rsTmp!ID
'        If strTmp <> "" Then
'            If rsTmp!ID = CLng(varFilter(mFilter.���˿���)) Then
'                cboPatientDept.ListIndex = cboPatientDept.NewIndex
'            End If
'        End If
'        rsTmp.MoveNext
'    Loop
    
    '===�������걾
    strSQL = "select ����,���� from ���Ƽ���걾 order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    cboSample.Clear
    cboSample.AddItem "���б걾"
    cboSample.ItemData(cboSample.NewIndex) = 0
    Do Until rsTmp.EOF
        cboSample.AddItem rsTmp("����") & "-" & rsTmp("����")
        cboSample.ItemData(cboSample.NewIndex) = rsTmp("����")
        strOldType = strOldType & "," & Trim(rsTmp("����") & "")
        If strTmp <> "" Then
            If rsTmp("����") = varFilter(mFilter.�걾) Then
                cboSample.ListIndex = cboSample.NewIndex
            End If
        End If
        rsTmp.MoveNext
    Loop
    '��ȡ�°�LIS�еı걾���͡�����ʹ�����뵥�¿�ҽ��ʱ��ѡ��ı걾����Ϊ�°�LIS�еı걾���ͣ�
    '���ɼ�����վ����ʹ�õı걾����Ϊ�ϰ�LIS�еı걾���ͣ�
    '������֮��û��ֱ�ӵĹ��������ԣ��ڲɼ�����վ���˹��ܵı걾������������°�ı걾����
    '����LIS�ӿ�
    If objLisInsideComm Is Nothing Then
        Dim strErr As String
        Set objLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        '��ʼ��LIS�ӿڲ���
        If Not objLisInsideComm Is Nothing Then
            If objLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "��ʼ��LIS�ӿ�ʧ�ܣ�" & vbCrLf & strErr
                End If
                Set objLisInsideComm = Nothing
            End If
        End If
    End If
    If Not objLisInsideComm Is Nothing Then
        Set rsTmp = objLisInsideComm.GetSampleTypeNew()   '��ȡ�°�LIS�еı걾����
        With Me.cboSample
            Do Until rsTmp.EOF
                If InStr(strOldType & ",", "," & Trim(rsTmp("����") & "") & ",") <= 0 Then
                    cboSample.AddItem rsTmp("����") & "-" & rsTmp("����")
                    cboSample.ItemData(cboSample.NewIndex) = rsTmp("����")
                End If
                If strTmp <> "" Then
                    If rsTmp("����") = varFilter(mFilter.�걾) Then
                        cboSample.ListIndex = cboSample.NewIndex
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End With
    End If
    If cboSample.Text = "" And cboSample.ListCount > 0 Then cboSample.ListIndex = 0
    
    '===����ɼ���ʽ(������Ѫ�ɼ�)
    strSQL = "select ID,���� from ������ĿĿ¼ where ���='E' and �������� in ('6','9')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    cboCapture.Clear
    cboCapture.AddItem "���з�ʽ"
    cboCapture.ItemData(cboCapture.NewIndex) = 0
    Do Until rsTmp.EOF
        cboCapture.AddItem rsTmp("����")
        cboCapture.ItemData(cboCapture.NewIndex) = rsTmp("ID")
        If strTmp <> "" Then
            If CLng(varFilter(mFilter.�ɼ���ʽ)) = rsTmp("ID") Then
                cboCapture.ListIndex = cboCapture.NewIndex
            End If
        End If
        rsTmp.MoveNext
    Loop
    If cboCapture.Text = "" And cboCapture.ListCount > 0 Then cboCapture.ListIndex = 0
    
    
    NowDate = zlDatabase.Currentdate
    
    Me.DTPBegin.Value = NowDate - 3
    Me.DTPEND.Value = NowDate
    
    '����ʱ��
    If strTmp <> "" Then
        Me.DTPBegin.Value = NowDate - varFilter(mFilter.���ʱ��)
        Me.DTPEND.Value = NowDate
    Else
        Me.DTPBegin.Value = NowDate - 3
        Me.DTPEND.Value = NowDate
    End If
    
    '��������
    strSQL = "select  distinct  ��������  from ������ĿĿ¼ Where ���='C' and �������� is not null  "
    If strBloodType <> "" Then
        strSQL = strSQL & " UNION " & vbNewLine & "Select '" & strBloodType & "' �������� from dual"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    lst����.Clear
    Do Until rsTmp.EOF
        lst����.AddItem Trim("" & rsTmp!��������)
        If strTmp <> "" Then
            If UBound(varFilter) >= mFilter.�������� Then
                If Trim("" & rsTmp!��������) <> "" And varFilter(mFilter.��������) <> "" Then
                    If InStr(varFilter(mFilter.��������) & ",", "," & Trim("" & rsTmp!��������) & ",") > 0 Then
                        lst����.Selected(lst����.NewIndex) = True
                    End If
                ElseIf varFilter(mFilter.��������) = "" Then
                    lst����.Selected(lst����.NewIndex) = True
                End If
            Else
                lst����.Selected(lst����.NewIndex) = True
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowME(Objfrm As Object, ByRef strFilter As String)
    Me.Show vbModal, Objfrm
    strFilter = mstrFilter
End Sub
