VERSION 5.00
Begin VB.Form frm�������㷽ʽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������㷽ʽ"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frm�������㷽ʽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   12
      Top             =   3000
      Width           =   5085
   End
   Begin VB.CommandButton cmd�ָ� 
      Caption         =   "��ԭ(&R)"
      Height          =   350
      Left            =   180
      TabIndex        =   9
      Top             =   2310
      Width           =   1100
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "����(&W)"
      Height          =   350
      Left            =   180
      TabIndex        =   10
      Top             =   3180
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2340
      TabIndex        =   7
      Top             =   3180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3600
      TabIndex        =   8
      Top             =   3180
      Width           =   1100
   End
   Begin VB.TextBox txt���㷽ʽ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1290
      TabIndex        =   6
      Top             =   1710
      Width           =   3525
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1290
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   300
      Left            =   4530
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label lblNote 
      Caption         =   "    ���ѡ�����ͨ������ԭ����ť���Իָ�Ĭ�ϵĲ��ּ����㷽ʽ��Ȼ����ȷ���ύ�����޸�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   2
      Left            =   1290
      TabIndex        =   11
      Top             =   2220
      Width           =   3615
   End
   Begin VB.Label lblNote 
      Caption         =   "    ��ѡ��һ�������֣�����סԺ�����ò��ֶ�Ӧ�����㷽ʽ�Է��ý��н���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   405
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   750
      Width           =   3615
   End
   Begin VB.Label lbl���㷽ʽ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���㷽ʽ(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   210
      TabIndex        =   5
      Top             =   1770
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frm�������㷽ʽ.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ����ǵ�һ��ʹ�û�ҽԺ�ĵ��������ݷ����仯����ʹ�����ع��ܣ��������������������ص����ء�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   150
      Width           =   3645
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&J)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   390
      TabIndex        =   2
      Top             =   1380
      Width           =   810
   End
End
Attribute VB_Name = "frm�������㷽ʽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Private mblnOK As Boolean
Private mint���� As Integer
Private mlng����ID As Long
Private mstr���� As String
Private mstrҽ���� As String
Private mstr�����ı�� As String
Private mstr���� As String
Private mrs���� As New ADODB.Recordset

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101")
    If Not CommServer("GETHOSPSINGLEILLNESS") Then Exit Sub
    MsgBox "���سɹ���", vbIbeam, gstrSysName
End Sub

Private Sub cmdOK_Click()
    If txt������Ϣ.Tag = "" Then
        MsgBox "��ѡ��һ�������֣�", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    
    '��ѡ������㷽ʽ�ϴ���ҽ������
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstrҽ����)
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", mstr�����ı��)
    Call InsertChild(mdomInput.documentElement, "RECKONINGTYPE", txt���㷽ʽ.Tag)
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", txt������Ϣ.Tag)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")) ' ��������
    If CommServer("SETRECKONINGTYPE") = False Then Exit Sub
    
    On Error Resume Next
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'������','''" & txt������Ϣ.Tag & "|" & txt���㷽ʽ.Tag & "''')"
    Call ExecuteProcedure("���浥���ֱ���")
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd�ָ�_Click()
    txt������Ϣ.Text = ""
    txt������Ϣ.Tag = "00000001"
    txt���㷽ʽ.Text = "���������㷽ʽ"
    txt���㷽ʽ.Tag = 1
End Sub

Private Sub cmd������Ϣ_Click()
    Dim blnReturn As Boolean
    blnReturn = frmListSel.ShowSelect(mrs����, "ID", "������ѡ��", "��ѡ�񵥲��֣�")
    If Not blnReturn Then mrs����.Filter = 0: Exit Sub
    
    txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
    txt������Ϣ.Tag = mrs����!����
    txt���㷽ʽ.Tag = mrs����!���㷽ʽ
    Select Case mrs����!���㷽ʽ
    Case 4
        txt���㷽ʽ.Text = "�����ְ�ʱ��������㷽ʽ"
    Case 3
        txt���㷽ʽ.Text = "�����ְ��˴ζ������㷽ʽ"
    Case Else
        txt���㷽ʽ.Text = "���������㷽ʽ"
    End Select
    mrs����.Filter = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '��ȡ�ò��˵�ҽ����Ϣ
    gstrSQL = "Select ������ From �����ʻ� Where ����ID=" & mlng����ID & " And ����=" & mint����
    Call OpenRecordset(rsTemp, "��ȡ�ò��˵�ҽ����Ϣ")
    txt������Ϣ.Text = NVL(rsTemp!������)
    If InStr(1, txt������Ϣ.Text, "|") <> 0 Then txt������Ϣ.Text = Mid(txt������Ϣ.Text, 1, InStr(1, txt������Ϣ.Text, "|") - 1)
    txt������Ϣ.Tag = txt������Ϣ.Text
    
    Call Get��֤_����(mstr����, mstrҽ����, mstr�����ı��, mstr����, mlng����ID)
    
    Call ��ȡ������
    Call ��ʾ������Ϣ
End Sub

Public Function ShowSelect(ByVal lng����ID As Long, ByVal int���� As Integer, ByVal frmParent As Object) As Boolean
    mblnOK = False
    mlng����ID = lng����ID
    mint���� = int����
    Me.Show 1, frmParent
    ShowSelect = mblnOK
End Function

Private Function ��ȡ������() As Boolean
    Dim strFields As String, strValues As String
    Dim str���� As String, str���� As String, str���� As String, str���㷽ʽ As String, str�����׼ As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Set mrs���� = New ADODB.Recordset
    strFields = "ID," & adVarChar & ",30|" & _
                "����," & adLongVarChar & ",30|" & _
                "����," & adLongVarChar & ",200|" & _
                "����," & adLongVarChar & ",30|" & _
                "���㷽ʽ," & adLongVarChar & ",10|" & _
                "�����׼," & adLongVarChar & ",500"
    Call Record_Init(mrs����, strFields)
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101")   '�����ı���̶�
    If CommServer("QUERYHOSPSINGLEILLNESS") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    '���ݱ���õ���������
    strFields = "ID|����|����|����|���㷽ʽ|�����׼"
    For Each nodRow In nodRowset.childNodes
        str���� = GetAttributeValue(nodRow, "SINGLEILLNESSCODE")
        str���� = GetAttributeValue(nodRow, "SINGLEILLNESSNAME")
        str���㷽ʽ = GetAttributeValue(nodRow, "RECKONINGTYPE")
        str�����׼ = GetAttributeValue(nodRow, "PAYSTD")
        str���� = zlCommFun.SpellCode(str����)
        strValues = str���� & "|" & str���� & "|" & str���� & "|" & str���� & "|" & str���㷽ʽ & "|" & str�����׼
        Call Record_Add(mrs����, strFields, strValues)
    Next
    ��ȡ������ = True
End Function

Private Function ��ʾ������Ϣ(Optional ByVal bln����ƥ�� As Boolean = False) As Boolean
    Dim blnReturn As Boolean
    Dim strInput As String, strFilter As String
    
    If Trim(txt������Ϣ.Text) = "" Then Exit Function
    'bln����ƥ��:�����������ƥ�䣬�����Ǵ����ݿ�����ϴ���ѡ��Ĳ��֣���˲�ȡ����ƥ�䣬���б���������Ƶģ�������ͨ���������鲡��ʱ��Ҫ����ƥ��
    If bln����ƥ�� Then
        strInput = UCase("'" & txt������Ϣ.Text & "*'")
        strFilter = "���� Like " & strInput & " Or ���� Like " & strInput & " Or ���� Like " & strInput
    Else
        strInput = UCase("'" & txt������Ϣ.Text & "'")
        strFilter = "����=" & strInput
    End If
    
    With mrs����
        .Filter = strFilter
        If .RecordCount = 0 Then
            If bln����ƥ�� Then
                MsgBox "û���ҵ�ָ���ĵ����֣�[���ֱ���Ϊ:" & UCase(txt������Ϣ.Text) & "]", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txt������Ϣ)
            txt������Ϣ.Text = ""
            txt������Ϣ.Tag = ""
            txt���㷽ʽ.Text = ""
            txt���㷽ʽ.Tag = 1
            .Filter = 0
            Exit Function
        Else
            If mrs����.RecordCount > 1 Then
                blnReturn = frmListSel.ShowSelect(mrs����, "ID", "������ѡ��", "��ѡ�񵥲��֣�")
            Else
                blnReturn = True
            End If
            If blnReturn = False Then
                txt������Ϣ.Text = ""
                txt������Ϣ.Tag = ""
                txt���㷽ʽ.Text = ""
                txt���㷽ʽ.Tag = 1
                Call zlControl.TxtSelAll(txt������Ϣ)
            Else
                txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
                txt������Ϣ.Tag = mrs����!����
                txt���㷽ʽ.Tag = mrs����!���㷽ʽ
                Select Case mrs����!���㷽ʽ
                Case 4
                    txt���㷽ʽ.Text = "�����ְ�ʱ��������㷽ʽ"
                Case 3
                    txt���㷽ʽ.Text = "�����ְ��˴ζ������㷽ʽ"
                Case Else
                    txt���㷽ʽ.Text = "���������㷽ʽ"
                End Select
                ��ʾ������Ϣ = True
            End If
        End If
        .Filter = 0
    End With
End Function

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt������Ϣ.Text) = "" Then Exit Sub
    
    If Not ��ʾ������Ϣ(True) Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
