VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmȨ������ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ø�����������Ҫ���ʵĶ���"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmȨ������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgSelect 
      Left            =   3360
      Top             =   3000
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
            Picture         =   "frmȨ������.frx":628A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   2415
      Left            =   2520
      TabIndex        =   2
      Top             =   2460
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgSelect"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "������"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5580
      TabIndex        =   4
      Top             =   5010
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6810
      TabIndex        =   5
      Top             =   5010
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   4275
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   7541
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin MSComctlLib.ImageList img���� 
      Left            =   420
      Top             =   1410
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
            Picture         =   "frmȨ������.frx":750C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw���� 
      Height          =   4305
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7594
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img����"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmȨ������.frx":878E
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ϸ���ø����������Ȩ�ޣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   690
      TabIndex        =   0
      Top             =   360
      Width           =   6900
   End
End
Attribute VB_Name = "frmȨ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mstrMethod As String
Private mblnFirst As Boolean            '����
Private mblnReturn As Boolean
Private mrsȨ�� As New ADODB.Recordset
Private mrs���� As New ADODB.Recordset

Public Function ShowEditor(rsȨ�� As ADODB.Recordset, ByVal rs���� As ADODB.Recordset) As Boolean
    mblnReturn = False
    Set mrsȨ�� = rsȨ��
    Set mrs���� = rs����
    
    Me.Show 1
    
    If mblnReturn Then Set rsȨ�� = mrsȨ��
    ShowEditor = mblnReturn
End Function

Private Sub Bill_CommandClick()
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    mstrSQL = "SELECT OBJECT_NAME,OBJECT_TYPE FROM ALL_OBJECTS " & _
        " WHERE OWNER='ZLHIS' AND OBJECT_TYPE<>'INDEX'"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "��ȡ�ɹ�ʹ�õĶ���")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���κζ���ɹ�ѡ��", vbInformation, gstrSysname
        Exit Sub
    End If
    
    lvwSelect.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvwSelect.ListItems.Add(, "K_" & .AbsolutePosition, !Object_Name, , 1)
            lvwItem.SubItems(1) = !object_TYPE
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then
            If Bill.Top + Bill.CellTop + Bill.RowHeight(0) + lvwSelect.Height > Me.Height Then
                lvwSelect.Top = Bill.Top + Bill.CellTop - lvwSelect.Height
            Else
                lvwSelect.Top = Bill.Top + Bill.CellTop + Bill.RowHeight(0)
            End If
            lvwSelect.Visible = True
            lvwSelect.ListItems(1).Selected = True
            lvwSelect.SetFocus
        End If
    End With
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    With Bill
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .Col = 0 Then
            If .TxtVisible = False Then Exit Sub
            strInput = UCase(.Text)
            
            mstrSQL = "SELECT OBJECT_NAME,OBJECT_TYPE FROM ALL_OBJECTS " & _
                " WHERE OWNER='ZLHIS' AND OBJECT_TYPE<>'INDEX'" & _
                " And OBJECT_NAME LIKE '" & strInput & "%'"
            Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "��ȡ�ɹ�ʹ�õĶ���")
            
            If rsTemp.RecordCount = 0 Then
                MsgBox "û���κζ���ɹ�ѡ��", vbInformation, gstrSysname
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            
            lvwSelect.ListItems.Clear
            With rsTemp
                Do While Not .EOF
                    Set lvwItem = lvwSelect.ListItems.Add(, "K_" & .AbsolutePosition, !Object_Name, , 1)
                    lvwItem.SubItems(1) = !object_TYPE
                    
                    .MoveNext
                Loop
                If .RecordCount <> 0 Then
                    If Bill.Top + Bill.CellTop + Bill.RowHeight(0) + lvwSelect.Height > Me.Height Then
                        lvwSelect.Top = Bill.Top + Bill.CellTop - lvwSelect.Height
                    Else
                        lvwSelect.Top = Bill.Top + Bill.CellTop + Bill.RowHeight(0)
                    End If
                    lvwSelect.Visible = True
                    lvwSelect.ListItems(1).Selected = True
                    lvwSelect.SetFocus
                    Cancel = True
                End If
            End With
        Else
            '�����������
            '�����ͼ������EXECUTE
            '����ֻ����SELECT
            '��������EXECUTE����������SELECT
        End If
    End With
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Call SavePrivs
    mblnReturn = True
    Unload Me
End Sub

Private Sub Form_Load()
    'װ����֧�ֵķ���
    With mrs����
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lvw����.ListItems.Add , "K_" & !���, !����, , 1
            .MoveNext
        Loop
    End With
    
    '��ʼ�����
    Call SetFormat
    
    mblnFirst = True
    If lvw����.ListItems.Count <> 0 Then Call lvw����_ItemClick(lvw����.ListItems(1))
    mblnFirst = False
End Sub

Private Sub lvwSelect_DblClick()
    Call lvwSelect_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvwSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lvwSelect.SelectedItem Is Nothing Then Exit Sub
    
    With Bill
        .TextMatrix(.Row, 0) = lvwSelect.SelectedItem.Text
        If InStr(1, "PROCEDURE|FUNCTION", lvwSelect.SelectedItem.SubItems(1)) <> 0 Then
            .TextMatrix(.Row, 5) = "��"
            .Col = 5
        Else
            .TextMatrix(.Row, 1) = "��"
            .Col = 1
        End If
    End With
    lvwSelect.Visible = False
End Sub

Private Sub lvwSelect_LostFocus()
    lvwSelect.Visible = False
End Sub

Private Sub lvw����_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    
    Call SavePrivs
    Call ShowPrivs
End Sub

Private Sub SetFormat()
    With Bill
        .Rows = 2
        .Cols = 6
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "SELECT"
        .TextMatrix(0, 2) = "INSERT"
        .TextMatrix(0, 3) = "UPDATE"
        .TextMatrix(0, 4) = "DELETE"
        .TextMatrix(0, 5) = "EXECUTE"
        
        .ColWidth(0) = 2200
        .ColWidth(1) = 650
        .ColWidth(2) = 650
        .ColWidth(3) = 650
        .ColWidth(4) = 650
        .ColWidth(5) = 750
        
        .ColData(0) = 1
        .ColData(1) = -1
        .ColData(2) = -1
        .ColData(3) = -1
        .ColData(4) = -1
        .ColData(5) = -1
        
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        
        .Active = True
        .PrimaryCol = 0
    End With
End Sub

Private Sub SavePrivs()
    Dim strPrivs As String
    Dim intRow As Integer, intRows As Integer
    Dim strField As String, strValue As String
    
    '�����ϴ�ѡ��ķ�����Ȩ��
    If mblnFirst Then
        mstrMethod = lvw����.SelectedItem.Text
        Exit Sub
    End If
    
    With mrsȨ��
        .Filter = "����='" & mstrMethod & "'"
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    intRows = Bill.Rows - 1
    strField = "����|����|Ȩ��"
    strValue = mstrMethod
    For intRow = 1 To intRows
        If Bill.TextMatrix(intRow, 0) <> "" Then
            strPrivs = IIf(Bill.TextMatrix(intRow, 1) = "", 0, 1) & IIf(Bill.TextMatrix(intRow, 2) = "", 0, 1) & _
                IIf(Bill.TextMatrix(intRow, 3) = "", 0, 1) & IIf(Bill.TextMatrix(intRow, 4) = "", 0, 1) & _
                IIf(Bill.TextMatrix(intRow, 5) = "", 0, 1)
            strPrivs = strValue & "|" & Bill.TextMatrix(intRow, 0) & "|" & strPrivs
            Call Record_Add(mrsȨ��, strField, strPrivs)
        End If
    Next
End Sub

Private Sub ShowPrivs()
    Dim intRow As Integer
    Dim intItem As Integer, intCount As Integer
    '��ʾָ��������Ȩ��
    
    mstrMethod = lvw����.SelectedItem.Text
    Bill.ClearBill
    intRow = 1
    
    With mrsȨ��
        .Filter = "����='" & mstrMethod & "'"
        Do While Not .EOF
            Bill.TextMatrix(intRow, 0) = !����
            Call WritePrivs(intRow, Nvl(!Ȩ��, "00000"))
            
            intRow = intRow + 1
            Bill.Rows = Bill.Rows + 1
            .MoveNext
        Loop
        .Filter = 0
    End With
End Sub

Private Sub WritePrivs(ByVal intRow As Integer, ByVal strPrivs As String)
    With Bill
        .TextMatrix(intRow, 1) = IIf(Mid(strPrivs, 1, 1) = 1, "��", "")
        .TextMatrix(intRow, 2) = IIf(Mid(strPrivs, 2, 1) = 1, "��", "")
        .TextMatrix(intRow, 3) = IIf(Mid(strPrivs, 3, 1) = 1, "��", "")
        .TextMatrix(intRow, 4) = IIf(Mid(strPrivs, 4, 1) = 1, "��", "")
        .TextMatrix(intRow, 5) = IIf(Mid(strPrivs, 5, 1) = 1, "��", "")
    End With
End Sub
