VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediUsage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷ�÷�����"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmMediUsage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkAllClass 
      Caption         =   "Ӧ���ڵ�ǰ����"
      Height          =   255
      Left            =   5520
      TabIndex        =   23
      ToolTipText     =   "������ǰ��������ͬ����ҩƷ�ĸ�ҩ;����Ƶ�ʣ�����������������Ϣ"
      Top             =   5243
      Width           =   1935
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "��"
      Height          =   285
      Left            =   7080
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   698
      Width           =   285
   End
   Begin ZL9BillEdit.BillEdit MSFAllergy 
      Height          =   1455
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   2566
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
   Begin VB.Frame frmline 
      Height          =   30
      Left            =   -15
      TabIndex        =   19
      Top             =   3200
      Width           =   7620
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -15
      TabIndex        =   18
      Top             =   5580
      Width           =   7620
   End
   Begin VB.TextBox txtPeriod 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   3720
      MaxLength       =   50
      TabIndex        =   9
      Top             =   5220
      Width           =   945
   End
   Begin VB.TextBox txtLimit 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1635
      MaxLength       =   50
      TabIndex        =   7
      Top             =   5190
      Width           =   1020
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   16
      Top             =   1305
      Width           =   7620
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ��÷�(&R)"
      Height          =   350
      Left            =   2835
      Picture         =   "frmMediUsage.frx":058A
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5650
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "����÷�(&C)"
      Height          =   350
      Left            =   1545
      Picture         =   "frmMediUsage.frx":06D4
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5650
      Width           =   1290
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1590
      MaxLength       =   50
      TabIndex        =   2
      Top             =   690
      Width           =   5505
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2400
      Left            =   240
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4800
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediUsage.frx":081E
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediUsage.frx":0DB8
            Key             =   "Method"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   5175
      TabIndex        =   10
      Top             =   5650
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      Picture         =   "frmMediUsage.frx":1352
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5650
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   6285
      TabIndex        =   11
      Top             =   5650
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit msfUsage 
      Height          =   1530
      Left            =   225
      TabIndex        =   5
      Top             =   3600
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   2699
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
   Begin VB.Label lbl�������� 
      Caption         =   "��������(&A)"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblPeriod 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "�Ƴ�(&T)            ��"
      Height          =   180
      Left            =   3000
      TabIndex        =   8
      Top             =   5280
      Width           =   1890
   End
   Begin VB.Label lblLimit 
      AutoSize        =   -1  'True
      Caption         =   "����������(M)"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   5250
      Width           =   1350
   End
   Begin VB.Label lblComment 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "(*δָ��С������ʱϵͳ�Զ����������㷨����)"
      Height          =   180
      Left            =   2640
      TabIndex        =   17
      Top             =   3360
      Width           =   3870
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "���ͣ�Ƭ��   ������λ��mg    ����"
      Height          =   180
      Left            =   1590
      TabIndex        =   3
      Top             =   1050
      Width           =   3060
   End
   Begin VB.Label lblUsage 
      AutoSize        =   -1  'True
      Caption         =   "�����÷�����(&U)"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   3345
      Width           =   1350
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "ָ��ҩƷƷ��(&I)"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   750
      Width           =   1350
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ���������ָ������ҩ���г�ҩ�ĳ����÷�������Ŀ�����ڸ���ҽ�����ӿ���׼ȷ�����ҩ��ҽ�����´"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   5685
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmMediUsage.frx":149C
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmMediUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngҩƷID As Long '�������մ�������ҩ��id

'---------------------------------------------------
'˵����
'   1����ǰ״̬����me.cmdClose.tag���棬�ֱ�Ϊ"�޸�"��"����"�����ϼ�����ͨ��ShowMe��������
'   2��ָ����Ŀ����me.lblItem.tag���棬���ϼ�����ͨ��ShowMe�������룬���Դ��ݣ�Ҳ���Բ�����
'---------------------------------------------------
Private strInputed As String
Private mblnChoose As Boolean
Dim rsTemp As New ADODB.Recordset
Dim ObjItem As ListItem
Dim strTemp As String
Dim intCount As Integer
Private mlng����id As Long

Public Sub ShowMe(ByVal frmParent As Object, ByVal blnEdit As Boolean, Optional ByVal lng��Ŀid As Long)
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    mlngҩƷID = lng��Ŀid
    Me.cmdClose.Tag = IIf(blnEdit, "�޸�", "����")
    If Me.cmdClose.Tag = "����" Then
        Me.msfUsage.Active = False
        Me.txtLimit.Enabled = False
        Me.txtPeriod.Enabled = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfUsage.Active = True
    End If
    Me.lblItem.Tag = lng��Ŀid
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,i.����id,I.����,I.����,I.���㵥λ,T.ҩƷ����,T.�������,nvl(T.��������,0) as ��������,t.������" & _
            " from ������ĿĿ¼ I,ҩƷ���� T" & _
            " where I.ID=T.ҩ��ID and I.��� in ('5','6') and I.ID=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.lblItem.Tag)
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag
        Else
            Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !���� & "]" & !����: Me.txtItem.Text = Me.txtItem.Tag
            Me.lblInfo.Caption = "ҩƷ���ͣ�" & IIf(IsNull(!ҩƷ����), "", !ҩƷ����) & _
                    "   ������λ��" & IIf(IsNull(!���㵥λ), "", !���㵥λ) & _
                    "   ������ࣺ" & IIf(IsNull(!�������), "", !�������)
            Me.txtLimit.Text = !��������
            mlng����id = !����id
            Call zlUsageRef(lng��Ŀid)
        End If
    End With
    Me.Show 1, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub chkAllergic_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdClear_Click()
    Me.msfUsage.ClearBill
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlUsageRef(Me.lblItem.Tag)
End Sub

Private Sub cmdSave_Click()
    Dim strsql As String
    Dim rscord As Recordset
    Dim str�÷����� As String
    Dim str�����÷� As String
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHand
    strsql = "select ҩ��id from ҩƷ���� where ҩ��id=[1] and ������=1"
    Set rscord = zlDatabase.OpenSQLRecord(strsql, "form_load", mlngҩƷID)
    
    If Val(Me.lblItem.Tag) = 0 Then MsgBox "δ��ȷָ��ҩƷ��", vbExclamation, gstrSysName: Me.txtItem.SetFocus: Exit Sub
    If Val(Me.txtLimit.Text) > 10000000 Then MsgBox "ϵͳ������̫��Ĵ���������Ϊ0��ʾ�����ƣ���", vbExclamation, gstrSysName: Me.txtLimit.SetFocus: Exit Sub
    If Val(Me.txtPeriod.Text) > 100 Then MsgBox "ϵͳ����������̫�����Ƴ̣�Ϊ0��ʾ�������Ƴ̣���", vbExclamation, gstrSysName: Me.txtPeriod.SetFocus: Exit Sub
    strTemp = "": gstrSql = ""
    With Me.msfUsage
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) & Trim(.TextMatrix(intCount, 3)) & Trim(.TextMatrix(intCount, 4)) & Trim(.TextMatrix(intCount, 5)) <> "" Then
                If .TextMatrix(intCount, 1) = "" Then
                    MsgBox "���÷���δ¼�룡", vbInformation, gstrSysName
                    .Col = 1
                    .SetFocus
                    Exit Sub
                End If
                If .TextMatrix(intCount, 3) = "" Then
                    MsgBox "��Ƶ�Ρ�δ¼�룡", vbInformation, gstrSysName
                    .Col = 3
                    .SetFocus
                    Exit Sub
                End If
            End If
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "����Ŀ�������ظ��ĸ�ҩ������", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                
                If .Cols > 7 Then
                    str�÷����� = Trim(.TextMatrix(intCount, 7))
                Else
                    str�÷����� = ""
                End If
                
                strTemp = strTemp & ";" & .RowData(intCount)
                gstrSql = gstrSql & "|" & .RowData(intCount) & "^" & .TextMatrix(intCount, 2) & _
                        "^" & Val(.TextMatrix(intCount, 4)) & "^" & Val(.TextMatrix(intCount, 5)) & "^" & Trim(.TextMatrix(intCount, 6)) & "^" & str�÷�����
            End If
        Next
    End With
    With Me.MSFAllergy
        For intCount = 1 To .Rows - 1
            If Val(.TextMatrix(intCount, 0)) <> 0 Then
                str�����÷� = .TextMatrix(intCount, 0) & "|" & str�����÷�
            End If
        Next
        
    End With
    
    If chkAllClass.Value = 1 Then
        strTemp = mlng����id
    Else
        strTemp = 0
    End If
    If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    gstrSql = "zl_�÷�����_UPDATE(" & Val(Me.lblItem.Tag) & "," & _
            IIf(str�����÷� = "", "NULL", "'" & str�����÷� & "'") & "," & _
            Val(Me.txtLimit.Text) & "," & Val(Me.txtPeriod.Text) & ",'" & gstrSql & "'," & 0 & "," & "0" & "," & strTemp & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    MsgBox Me.txtItem.Text & " �÷���������ɹ���", vbInformation, gstrSysName
    Me.txtItem.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdItem_Click()
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "����", "����", 900
        .Add , "���㵥λ", "��λ", 500
        .Add , "ҩƷ����", "����", 800
        .Add , "�������", "����", 900
        .Add , "������", "����ҩ��", 500
        .Add , "����id", "����id", 0
        .Add , "����", "����", 0
    End With
    Err = 0: On Error GoTo ErrHand
    gstrSql = "select I.ID,i.����id,I.����,I.����,I.���㵥λ,T.ҩƷ����,T.�������,nvl(T.��������,0) as ��������,t.������ " & _
            " from ������ĿĿ¼ I,ҩƷ���� T" & _
            " where I.ID=T.ҩ��ID and I.��� in ('5','6')" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmdItem_Click")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "��δ��������ҩ���г�ҩ��", vbExclamation, gstrSysName
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        mlngҩƷID = !ID
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            ObjItem.Icon = "ItemUse": ObjItem.SmallIcon = "ItemUse"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("ҩƷ����").Index - 1) = IIf(IsNull(!ҩƷ����), "", !ҩƷ����)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("�������").Index - 1) = IIf(IsNull(!�������), "", !�������)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("������").Index - 1) = IIf(IsNull(!������), "", !������)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����id").Index - 1) = IIf(IsNull(!����id), "", !����id)
            ObjItem.Tag = !��������
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Tag = "ҩƷ"
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .Width = Me.txtItem.Width + Me.cmdItem.Width
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        If Me.lvwItems.Tag = Me.txtItem.Name Then
            Me.txtItem.SetFocus
        Else
            Me.msfUsage.SetFocus
        End If
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strsql As String
    Dim rscord As Recordset
    
    On Error GoTo ErrHandle
    strsql = "select ҩ��id from ҩƷ���� where ҩ��id=[1] and ������ in (1,2,3)"
    Set rscord = zlDatabase.OpenSQLRecord(strsql, "form_load", mlngҩƷID)
    With Me.msfUsage
        .Active = True
        
        If Not rscord.EOF Then
            .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 8
        Else
            .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 7
        End If
      
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "�÷�": .TextMatrix(0, 2) = "Ƶ����": .TextMatrix(0, 3) = "Ƶ��"
        .TextMatrix(0, 4) = "���˼���": .TextMatrix(0, 5) = "С������": .TextMatrix(0, 6) = "ҽ������"
        
        
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 5: .ColData(3) = 1: .ColData(4) = 4: .ColData(5) = 4: .ColData(6) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 1200: .ColWidth(2) = 0: .ColWidth(3) = 1200
        .ColWidth(4) = 1000: .ColWidth(5) = 1000: .ColWidth(6) = 1350
        
        .ColAlignment(0) = 1: .ColAlignment(1) = 1: .ColAlignment(2) = 1: .ColAlignment(3) = 1
        .ColAlignment(4) = 7: .ColAlignment(5) = 7: .ColAlignment(6) = 1
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
        
        If Not rscord.EOF Then
            .TextMatrix(0, 7) = "DDDֵ"
            .ColData(7) = 4
            .ColWidth(7) = 1000
        End If
    End With
    
    
     With Me.MSFAllergy
        .Active = True

        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 2
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "����������Ŀ"
        .ColData(0) = 5: .ColData(1) = 1
        .ColWidth(0) = 0: .ColWidth(1) = 3600
  
        .ColAlignment(0) = 1: .ColAlignment(1) = 1
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����id = 0
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To lvwItems.ColumnHeaders.Count
        If InStr(1, lvwItems.ColumnHeaders.Item(i), "����") > 0 Then
            mlng����id = lvwItems.SelectedItem.SubItems(lvwItems.ColumnHeaders("����id").Index - 1)
        End If
    Next
    
    With Me.lvwItems
        Select Case .Tag
        Case "ҩƷ"
            If Me.lblItem.Tag <> Mid(.SelectedItem.Key, 2) Then
                Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
                Me.txtItem.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
                Me.txtItem.Text = Me.txtItem.Tag
                Me.txtPeriod.Text = Val(.SelectedItem.Tag)
                
                If .SelectedItem.SubItems(.ColumnHeaders("������").Index - 1) = "1" Then
                    msfUsage.ColWidth(7) = 1000
                Else
                    msfUsage.ColWidth(7) = 0
                End If
                
                Call zlUsageRef(Me.lblItem.Tag)
            End If
            Me.txtItem.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Case "����"
            For i = 1 To Me.MSFAllergy.Rows - 1
                If Me.MSFAllergy.TextMatrix(i, 0) = Mid(.SelectedItem.Key, 2) And i <> Me.MSFAllergy.Row Then
                    Me.lvwItems.Visible = False
                    Me.MSFAllergy.Text = ""
                    Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 0) = ""
                    Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 1) = ""
                    Exit Sub
                End If
            Next
            
            Me.MSFAllergy.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
            Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 0) = Mid(.SelectedItem.Key, 2)
            Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 1) = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
            Me.MSFAllergy.SetFocus
            Call zlCommFun.PressKey(13)
            Me.lvwItems.Visible = False
        Case "�÷�"
            Me.msfUsage.Text = .SelectedItem.Text
            Me.msfUsage.RowData(Me.msfUsage.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfUsage.TextMatrix(Me.msfUsage.Row, 1) = Me.msfUsage.Text
            Me.msfUsage.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        Case "Ƶ��"
            Me.msfUsage.Text = .SelectedItem.Text
            Me.msfUsage.TextMatrix(Me.msfUsage.Row, 2) = .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
            Me.msfUsage.TextMatrix(Me.msfUsage.Row, 3) = Me.msfUsage.Text
            Me.msfUsage.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        End Select
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub MSFAllergy_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If (Me.MSFAllergy.Row > 1) Or (Me.MSFAllergy.Row = 1 And Me.MSFAllergy.Rows > 2) Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub MSFAllergy_CommandClick()
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "����", "����", 900
        .Add , "���㵥λ", "��λ", 550
        .Add , "����id", "����id", 0
    End With
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "select I.ID,I.����,I.����,I.���㵥λ,i.����id" & _
            " from ������ĿĿ¼ I" & _
            " where I.���='E' and I.��������='1'" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "MSFAllergy_CommandClick")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "��δ����������Ŀ��", vbExclamation, gstrSysName
            Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 0) = "": Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 1) = "": Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            ObjItem.Icon = "Method": ObjItem.SmallIcon = "Method"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����id").Index - 1) = IIf(IsNull(!����id), "", !����id)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Tag = "����"
        .Left = Me.MSFAllergy.Left
        .Top = Me.MSFAllergy.Top + (MSFAllergy.Row - MSFAllergy.MsfObj.TopRow + 1) * MSFAllergy.RowHeight(0) + 300
        .Width = 3600
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MSFAllergy_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim i As Integer
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyCode)) > 0 And KeyCode <> 46 Then KeyCode = 0: Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.MSFAllergy.Text))
    If strTemp = "" Then Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 0) = 0: Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "����", "����", 900
        .Add , "���㵥λ", "��λ", 550
        .Add , "����id", "����id", 0
    End With
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.����,I.����,I.���㵥λ,i.����id" & _
            " from ������ĿĿ¼ I,������Ŀ���� N" & _
            " where I.ID=N.������ĿID and I.���='E' and I.��������='1'" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [1] or N.���� like [2] or N.���� like [2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ���Ĺ���������Ŀ��������ָ��", vbExclamation, gstrSysName
            Me.MSFAllergy.Text = ""
            Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 0) = "":  Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 1) = "": Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            For i = 1 To Me.MSFAllergy.Rows - 1
                If Me.MSFAllergy.TextMatrix(i, 0) = !ID And i <> Me.MSFAllergy.Row Then
                    Me.lvwItems.Visible = False
                    MsgBox "���������ظ���Ŀ��������ָ��", vbExclamation, gstrSysName
                    Me.MSFAllergy.Text = ""
                    Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 0) = ""
                    Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 1) = ""
                    Cancel = True
                    Exit Sub
                End If
            Next
            
            Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 0) = !ID
            Me.MSFAllergy.Text = "[" & !���� & "]" & !����
            Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Row, 1) = "[" & !���� & "]" & !����
            Me.MSFAllergy.SetFocus
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            ObjItem.Icon = "Method": ObjItem.SmallIcon = "Method"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����id").Index - 1) = IIf(IsNull(!����id), "", !����id)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Tag = "����"
        .Left = Me.MSFAllergy.Left
        .Top = Me.MSFAllergy.Top + (MSFAllergy.Row - MSFAllergy.MsfObj.TopRow + 1) * MSFAllergy.RowHeight(0) + 300
        .Width = 3600
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfUsage_AfterAddRow(Row As Long)
    With Me.msfUsage
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfUsage_AfterDeleteRow()
    With Me.msfUsage
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfUsage_CommandClick()
    If Me.msfUsage.Col = 1 Then
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 2000
            .Add , "����", "����", 900
            .Add , "���㵥λ", "��λ", 500
            .Add , "����id", "����id", 0
        End With
        
        Err = 0: On Error GoTo ErrHand
        
        gstrSql = "select I.ID,i.����id,I.����,I.����,I.���㵥λ" & _
                " from ������ĿĿ¼ I" & _
                " where I.���='E' and I.��������='2'" & _
                "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "msfUsage_CommandClick")
'            Call SQLTest
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "�뽨����ҩ;����Ŀ����У�", vbExclamation, gstrSysName: Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
                ObjItem.Icon = "Method": ObjItem.SmallIcon = "Method"
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����id").Index - 1) = IIf(IsNull(!����id), "", !����id)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
            .Tag = "�÷�"
            .Left = Me.msfUsage.Left + 250
            .Top = Me.msfUsage.Top + (msfUsage.Row - msfUsage.MsfObj.TopRow + 1) * msfUsage.RowHeight(0) - .Height
            .Width = 3600
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 1200
            .Add , "����", "����", 500
            .Add , "����", "����", 800
            .Add , "Ӣ������", "Ӣ��", 600
            .Add , "����id", "����id", 0
        End With
        
        gstrSql = "select rownum as ����id,����,����,����,Ӣ������ from ����Ƶ����Ŀ where ���÷�Χ<>2 order by ����"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "msfUsage_CommandClick")
'            Call SQLTest
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "�뽨������Ƶ�ʺ����(�ֵ����)��", vbExclamation, gstrSysName: Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !����, !����)
                ObjItem.Icon = "Method": ObjItem.SmallIcon = "Method"
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("Ӣ������").Index - 1) = IIf(IsNull(!Ӣ������), "", !Ӣ������)
               ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����id").Index - 1) = IIf(IsNull(!����id), "", !����id)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
            .Tag = "Ƶ��"
            .Left = Me.msfUsage.Left + 1500
            .Top = Me.msfUsage.Top + (msfUsage.Row - msfUsage.MsfObj.TopRow + 1) * msfUsage.RowHeight(0) - .Height
            .Width = 3600
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub msfUsage_EditKeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intzheng As Integer '��¼�������ֵĸ���
    
    msfUsage.MaxLength = 16
    If msfUsage.Col = 7 Then
        If KeyAscii = Asc(".") Then
            i = InStr(1, msfUsage.Text, ".") '�ж���ǰ�Ƿ��е�
            If i > 0 Then
             KeyAscii = 0
            End If
        End If
        
        i = InStr(1, msfUsage.Text, ".")
        If i <> 0 Then
            If Len(Mid(msfUsage.Text, i + 1)) > 3 Then
                intzheng = Len(Mid(msfUsage.Text, 1, i - 1))
                msfUsage.MaxLength = intzheng + 6
                Exit Sub
            End If
        Else
            msfUsage.MaxLength = 10
        End If
    End If

End Sub

Private Sub msfUsage_EnterCell(Row As Long, Col As Long)
    Dim i As Integer
    If Col = 4 Or Col = 5 Or Col = 7 Then
        msfUsage.TxtCheck = True
        msfUsage.TextMask = "0123456789."
    Else
        msfUsage.TxtCheck = False
        msfUsage.TextMask = ""
    End If
    strInputed = Me.msfUsage.TextMatrix(Row, Col)
End Sub

Private Sub msfUsage_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfUsage_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfUsage
        If .Active = False Then Exit Sub
        Select Case .Col
        Case 4, 5
            If .TxtVisible = False Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then .TextMatrix(.Row, .Col) = "0"
            Else
                If Trim(.Text) = "" Then .Text = 0: .TextMatrix(.Row, .Col) = "0"
            End If
        Case 6
            If .TxtVisible = False Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then .TextMatrix(.Row, .Col) = Space(1)
            Else
                If Trim(.Text) = "" Then .Text = Space(1): .TextMatrix(.Row, .Col) = Space(1)
            End If
        End Select
        If .Col <> 1 And .Col <> 3 Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then Exit Sub
            strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
        Else
            If Trim(.Text) = "" Then Exit Sub
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strInputed = strTemp Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    If Me.msfUsage.Col = 1 Then
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 2000
            .Add , "����", "����", 900
            .Add , "���㵥λ", "��λ", 500
            .Add , "����id", "����id", 0
        End With
        
        Err = 0: On Error GoTo ErrHand
        
        gstrSql = "select distinct I.ID,I.����,I.����,I.���㵥λ,i.����id" & _
                " from ������ĿĿ¼ I,������Ŀ���� N" & _
                " where I.ID=N.������Ŀid and I.���='E' and I.��������='2'" & _
                "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.���� like [1] or N.���� like [2] or N.���� like [2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "δ�ҵ�ָ���÷�(��ҩ;��)�����������룡", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
            If .RecordCount = 1 Then
                Me.msfUsage.Text = !����
                Me.msfUsage.TextMatrix(Me.msfUsage.Row, 1) = Me.msfUsage.Text
                Me.msfUsage.RowData(Me.msfUsage.Row) = !ID
                Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
                ObjItem.Icon = "Method": ObjItem.SmallIcon = "Method"
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����id").Index - 1) = IIf(IsNull(!����id), "", !����id)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
            .Tag = "�÷�"
            .Left = Me.msfUsage.Left + 260
            .Top = Me.msfUsage.Top + (msfUsage.Row - msfUsage.MsfObj.TopRow + 1) * msfUsage.RowHeight(0) - .Height
            .Width = 3600
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 1200
            .Add , "����", "����", 500
            .Add , "����", "����", 800
            .Add , "Ӣ������", "Ӣ��", 600
        End With
        
        gstrSql = "select ����,����,����,Ӣ������" & _
                " from ����Ƶ����Ŀ" & _
                " where ���÷�Χ<>2 and (���� like [1] or ���� like [2] " & _
                "   or ���� like [2] or upper(Ӣ������) like [2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "δ�ҵ�ָ��Ƶ�ʣ����������룡", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
            If .RecordCount = 1 Then
                Me.msfUsage.Text = !����
                Me.msfUsage.TextMatrix(Me.msfUsage.Row, 2) = !����
                Me.msfUsage.TextMatrix(Me.msfUsage.Row, 3) = Me.msfUsage.Text
                Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !����, !����)
                ObjItem.Icon = "Method": ObjItem.SmallIcon = "Method"
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("Ӣ������").Index - 1) = IIf(IsNull(!Ӣ������), "", !Ӣ������)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
            .Tag = "Ƶ��"
            .Left = Me.msfUsage.Left + 1500
            .Top = Me.msfUsage.Top + (msfUsage.Row - msfUsage.MsfObj.TopRow + 1) * msfUsage.RowHeight(0) - .Height
            .Width = 3600
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End If
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub msfUsage_LeaveCell(Row As Long, Col As Long)
    Dim i As Integer
    Dim strchar As String
    '�ж��Ƿ��зǷ��ַ�����������Զ����
    If msfUsage.Col = 7 Then
        i = InStr(1, msfUsage.TextMatrix(Row, Col), ".")
        If i <> 0 Then
            strchar = Mid(msfUsage.TextMatrix(Row, Col), i + 1)
            If InStr(1, strchar, ".") > 0 Then
                msfUsage.TextMatrix(Row, Col) = ""
                Exit Sub
            End If
        End If
    End If
End Sub


Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtItem.Text))
    If strTemp = "" Then Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "����", "����", 900
        .Add , "���㵥λ", "��λ", 500
        .Add , "ҩƷ����", "����", 800
        .Add , "�������", "����", 900
        .Add , "������", "������", 500
        .Add , "����id", "����id", 0
        .Add , "����", "����", 0
    End With
    
    Err = 0: On Error GoTo ErrHand
        
    gstrSql = "select distinct I.ID,i.����id,I.����,I.����,I.���㵥λ,T.ҩƷ����,T.�������,nvl(T.��������,0) as ��������,t.������" & _
            " from ������ĿĿ¼ I,������Ŀ���� N,ҩƷ���� T" & _
            " where I.ID=N.������ĿID and I.ID=T.ҩ��ID and I.��� in ('5','6')" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [1] or N.���� like [2] or N.���� like [2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
    
    mlngҩƷID = rsTemp!ID
    If rsTemp!������ = "1" Then
        msfUsage.ColWidth(7) = 1000
    Else
        msfUsage.ColWidth(7) = 0
    End If
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ���ĳ�ҩƷ�֣�������ָ��", vbExclamation, gstrSysName
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblItem.Tag <> !ID Then
                Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !���� & "]" & !����: Me.txtItem.Text = Me.txtItem.Tag
                Me.lblInfo.Caption = "ҩƷ���ͣ�" & IIf(IsNull(!ҩƷ����), "", !ҩƷ����) & _
                        "   ������λ��" & IIf(IsNull(!���㵥λ), "", !���㵥λ) & _
                        "   ������ࣺ" & IIf(IsNull(!�������), "", !�������)
                Me.txtLimit.Text = !��������
                Call zlUsageRef(Me.lblItem.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            ObjItem.Icon = "ItemUse": ObjItem.SmallIcon = "ItemUse"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("ҩƷ����").Index - 1) = IIf(IsNull(!ҩƷ����), "", !ҩƷ����)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("�������").Index - 1) = IIf(IsNull(!�������), "", !�������)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("������").Index - 1) = IIf(IsNull(!������), "", !������)
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����id").Index - 1) = IIf(IsNull(!����id), "", !����id)
            ObjItem.Tag = !��������
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Tag = "ҩƷ"
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .Width = Me.txtItem.Width + Me.cmdItem.Width
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_LostFocus()
    Me.txtItem.Text = Me.txtItem.Tag
End Sub

Private Sub txtLimit_GotFocus()
    Me.txtLimit.SelStart = 0: Me.txtLimit.SelLength = 100
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtPeriod_GotFocus()
    Me.txtPeriod.SelStart = 0: Me.txtPeriod.SelLength = 100
End Sub

Private Sub txtPeriod_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub zlUsageRef(lngItemID As Long)
    '--------------------------------------------------------
    '���ܣ�ˢ����ʾҩƷ�÷�����
    '��Σ�lngItemId-ָ����������Ŀid���˴�Ϊ��ҩ��
    '--------------------------------------------------------
    Dim strsql As String
    Dim rsDDD As ADODB.Recordset
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,'['||I.����||']'||I.���� as ����" & _
            " from �����÷����� R,������ĿĿ¼ I" & _
            " where R.�÷�ID=I.ID and R.����=0 and R.��ĿID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    With rsTemp
        Me.MSFAllergy.Rows = 2
        Do While Not .EOF
            Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Rows - 1, 0) = !ID: Me.MSFAllergy.TextMatrix(Me.MSFAllergy.Rows - 1, 1) = !����
            Me.MSFAllergy.Rows = Me.MSFAllergy.Rows + 1
            rsTemp.MoveNext
        Loop
        
    End With
    
    Me.txtPeriod.Text = 3
    gstrSql = "select I.ID,I.���� as ����,P.���� as Ƶ����,P.���� as Ƶ����,r.dddֵ," & _
            " nvl(R.���˼���,0) as ���˼���,nvl(R.С������,0) as С������,R.ҽ������,nvl(R.�Ƴ�,3) as �Ƴ� " & _
            " from �����÷����� R,������ĿĿ¼ I,����Ƶ����Ŀ P" & _
            " where R.�÷�ID=I.ID and R.Ƶ��=P.����(+) and R.����>0 and R.��ĿID=[1] " & _
            " order by R.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
        
    With rsTemp
        Me.msfUsage.ClearBill
        Do While Not .EOF
            If Me.msfUsage.Rows - 1 < .AbsolutePosition Then Me.msfUsage.Rows = Me.msfUsage.Rows + 1
            Me.msfUsage.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfUsage.RowData(.AbsolutePosition) = !ID
            Me.msfUsage.TextMatrix(.AbsolutePosition, 1) = !����
            Me.msfUsage.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!Ƶ����), "", !Ƶ����)
            Me.msfUsage.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!Ƶ����), "", !Ƶ����)
            Me.msfUsage.TextMatrix(.AbsolutePosition, 4) = !���˼���
            If Left(Me.msfUsage.TextMatrix(.AbsolutePosition, 4), 1) = "." Then
                Me.msfUsage.TextMatrix(.AbsolutePosition, 4) = "0" & Me.msfUsage.TextMatrix(.AbsolutePosition, 4)
            End If
            Me.msfUsage.TextMatrix(.AbsolutePosition, 5) = !С������
            If Left(Me.msfUsage.TextMatrix(.AbsolutePosition, 5), 1) = "." Then
                Me.msfUsage.TextMatrix(.AbsolutePosition, 5) = "0" & Me.msfUsage.TextMatrix(.AbsolutePosition, 5)
            End If
            Me.msfUsage.TextMatrix(.AbsolutePosition, 6) = IIf(IsNull(!ҽ������), "", !ҽ������)
            If msfUsage.Cols > 7 Then
                Me.msfUsage.TextMatrix(.AbsolutePosition, 7) = IIf(IsNull(!dddֵ), "", !dddֵ)
                If Val(msfUsage.TextMatrix(.AbsolutePosition, 7)) = 0 Then
                    strsql = "select nvl(dddֵ,0) dddֵ  from ҩƷ��� where ҩ��id=[1]"    '����������÷�������δ����dddֵ����ҩƷ�������ȡһ��dddֵ
                    Set rsDDD = zlDatabase.OpenSQLRecord(strsql, "DDDֵ", lngItemID)
                    Do While Not rsDDD.EOF
                        If rsDDD!dddֵ <> 0 Then
                            msfUsage.TextMatrix(.AbsolutePosition, 7) = rsDDD!dddֵ
                            Exit Do
                        End If
                        rsDDD.MoveNext
                    Loop
                End If
            End If
            Me.txtPeriod.Text = !�Ƴ�
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

