VERSION 5.00
Begin VB.Form frmClinicPlanParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6840
   Icon            =   "frmClinicPlanParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo����ȽϷ�ʽ 
      Height          =   300
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox cbo��Դ����վ�� 
      Height          =   300
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox chk����ҽ�������� 
      Caption         =   "����ҽ��ְ�񼶱���"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   2145
   End
   Begin VB.Frame fraSplit 
      Height          =   4485
      Left            =   5280
      TabIndex        =   22
      Top             =   -150
      Width           =   25
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "�ܳ�����ӡ����(&4)"
      Height          =   405
      Index           =   3
      Left            =   2910
      TabIndex        =   21
      Top             =   3870
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "�³�����ӡ����(&3)"
      Height          =   405
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   3870
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "�̶�������ӡ����(&2)"
      Height          =   405
      Index           =   1
      Left            =   2910
      TabIndex        =   19
      Top             =   3420
      Width           =   2145
   End
   Begin VB.CheckBox chkReplaceDoctor 
      Caption         =   "������ҽ��ͬ������ԤԼ�Һŵ�"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   90
      Width           =   2835
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "ԤԼ�嵥��ӡ����(&1)"
      Height          =   405
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   3420
      Width           =   2145
   End
   Begin VB.Frame fraVisitTablePrintMode 
      Caption         =   "������ӡ��ʽ"
      Height          =   735
      Left            =   180
      TabIndex        =   14
      Top             =   2550
      Width           =   4875
      Begin VB.OptionButton optVisitTablePrintMode 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optVisitTablePrintMode 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1035
      End
      Begin VB.OptionButton optVisitTablePrintMode 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   2
         Left            =   3090
         TabIndex        =   17
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Frame fraPrintMode 
      Caption         =   "ԤԼ�嵥��ӡ��ʽ"
      Height          =   1305
      Left            =   3060
      TabIndex        =   10
      Top             =   1110
      Width           =   1995
      Begin VB.OptionButton optPrintMode 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   13
         Top             =   930
         Width           =   1395
      End
      Begin VB.OptionButton optPrintMode 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   615
         Width           =   1035
      End
      Begin VB.OptionButton optPrintMode 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraToExcelMode 
      Caption         =   "ԤԼ�嵥���Ʒ�ʽ"
      Height          =   1305
      Left            =   180
      TabIndex        =   6
      Top             =   1110
      Width           =   2715
      Begin VB.OptionButton optToExcelMode 
         Caption         =   "ѡ���Ƿ������Excel"
         Height          =   225
         Index           =   2
         Left            =   300
         TabIndex        =   9
         Top             =   930
         Width           =   2025
      End
      Begin VB.OptionButton optToExcelMode 
         Caption         =   "�Զ������Excel"
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   8
         Top             =   615
         Width           =   1665
      End
      Begin VB.OptionButton optToExcelMode 
         Caption         =   "�������Excel"
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   330
      Left            =   5550
      TabIndex        =   23
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   330
      Left            =   5550
      TabIndex        =   24
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   330
      Left            =   5580
      TabIndex        =   25
      Top             =   3390
      Width           =   1100
   End
   Begin VB.Label lbl����ȽϷ�ʽ 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��Դ����ıȽϷ�ʽ"
      Height          =   180
      Left            =   210
      TabIndex        =   4
      Top             =   780
      Width           =   2160
   End
   Begin VB.Label lbl��Դ����վ�� 
      AutoSize        =   -1  'True
      Caption         =   "��δ����վ��ĺ�Դ�����                   ���г��ﰲ��"
      Height          =   180
      Left            =   210
      TabIndex        =   2
      Top             =   420
      Width           =   4950
   End
End
Attribute VB_Name = "frmClinicPlanParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String
Private mlngModul As Long
Private mblnOk As Boolean

Public Function ShowMe(frmParent As Form, ByVal lngModul As Long, _
    ByVal strPrivs As String) As Boolean
    '�������
    mstrPrivs = strPrivs: mlngModul = lngModul
    
    On Error Resume Next
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim blnHavePrivs As Boolean
    Dim strValue As String
    
    On Error GoTo ErrHandler
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "����ҽ��������", IIf(chk����ҽ��������.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "������ҽ��ͬ������ԤԼ�Һŵ�", IIf(chkReplaceDoctor.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "δ����վ��ĺ�Դ��ά��վ��", zlStr.NeedCode(cbo��Դ����վ��.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "��������ȽϷ�ʽ", cbo����ȽϷ�ʽ.ItemData(cbo����ȽϷ�ʽ.ListIndex), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "ԤԼ�嵥���Ʒ�ʽ", GetSelectedIndex(optToExcelMode), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ԤԼ�嵥��ӡ��ʽ", GetSelectedIndex(optPrintMode), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "������ӡ��ʽ", GetSelectedIndex(optVisitTablePrintMode), glngSys, mlngModul, blnHavePrivs
    mblnOk = True
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub cmdPrintSet_Click(index As Integer)
    On Error GoTo ErrHandler
    Select Case index
    Case 0: 'ԤԼ�嵥��ӡ��ʽ
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me)
    Case 1: '�̶�������ӡ��ʽ
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_1", Me)
    Case 2: '�³�����ӡ��ʽ
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_2", Me)
    Case 3: '�ܳ�����ӡ��ʽ
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_3", Me)
    Case Else:
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub Form_Load()
    Dim index As Integer, blnHavePrivs As Boolean, strValue As String
    Dim strSQL As String, rsRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    blnHavePrivs = IsHavePrivs(mstrPrivs, "��������")
    
    chk����ҽ��������.Value = Val(zlDatabase.GetPara("����ҽ��������", glngSys, mlngModul, 0, Array(chk����ҽ��������), blnHavePrivs))
    chkReplaceDoctor.Value = Val(zlDatabase.GetPara("������ҽ��ͬ������ԤԼ�Һŵ�", glngSys, mlngModul, 0, Array(chkReplaceDoctor), blnHavePrivs))
    
    strValue = zlDatabase.GetPara("δ����վ��ĺ�Դ��ά��վ��", glngSys, mlngModul, "", Array(lbl��Դ����վ��, cbo��Դ����վ��), blnHavePrivs)
    strSQL = _
        "Select Distinct b.���, b.����" & vbNewLine & _
        "From ���ű� A, Zlnodelist B" & vbNewLine & _
        "Where a.վ�� = b.���" & vbNewLine & _
        "Order By b.���"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "վ���ѯ")
    With cbo��Դ����վ��
        .Clear
        .AddItem ""
        Do While Not rsRecord.EOF
            .AddItem rsRecord!��� & "-" & rsRecord!����
            If strValue = rsRecord!��� Then .ListIndex = .NewIndex
            rsRecord.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    With cbo����ȽϷ�ʽ
        .Clear
        .AddItem "0-���ַ��Ƚ�":  .ItemData(.NewIndex) = 0
        .AddItem "1-����ֵ�Ƚ�": .ItemData(.NewIndex) = 1
    End With
    index = Val(zlDatabase.GetPara("��������ȽϷ�ʽ", glngSys, mlngModul, 0, Array(lbl����ȽϷ�ʽ, cbo����ȽϷ�ʽ), blnHavePrivs))
    zlControl.CboLocate cbo����ȽϷ�ʽ, index, True
    
    index = Val(zlDatabase.GetPara("ԤԼ�嵥���Ʒ�ʽ", glngSys, mlngModul, 0, Array(optToExcelMode(0), optToExcelMode(1), optToExcelMode(2)), blnHavePrivs))
    If index <= optToExcelMode.UBound Then optToExcelMode(index).Value = True
    
    index = Val(zlDatabase.GetPara("ԤԼ�嵥��ӡ��ʽ", glngSys, mlngModul, 0, Array(optPrintMode(0), optPrintMode(1), optPrintMode(2)), blnHavePrivs))
    If index <= optPrintMode.UBound Then optPrintMode(index).Value = True
    
    index = Val(zlDatabase.GetPara("������ӡ��ʽ", glngSys, mlngModul, 0, Array(optVisitTablePrintMode(0), optVisitTablePrintMode(1), optVisitTablePrintMode(2)), blnHavePrivs))
    If index <= optVisitTablePrintMode.UBound Then optVisitTablePrintMode(index).Value = True
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

