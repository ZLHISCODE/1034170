VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmApparatusItem 
   BorderStyle     =   0  'None
   Caption         =   "������Ŀͨ��"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2490
      Left            =   135
      TabIndex        =   4
      Top             =   105
      Width           =   8145
      _cx             =   14367
      _cy             =   4392
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   135
      ScaleHeight     =   2505
      ScaleWidth      =   8145
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2640
      Width           =   8145
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   450
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CheckBox chkKind 
         Caption         =   "���ּ�������(&K)"
         Height          =   210
         Left            =   6060
         TabIndex        =   8
         Top             =   1935
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   6060
         TabIndex        =   1
         Top             =   720
         Width           =   1755
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "���ҡ�    "
         Height          =   350
         Left            =   6060
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "���ҷ�����������Ŀ"
         Top             =   1065
         Width           =   1185
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ���ӵ�������Ŀ�б���"
         Height          =   350
         Index           =   0
         Left            =   15
         TabIndex        =   5
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ��������Ŀ�б���ɾ��"
         Height          =   350
         Index           =   1
         Left            =   2610
         TabIndex        =   6
         Top             =   45
         Width           =   2535
      End
      Begin VB.CheckBox chkUpper 
         Caption         =   "���ִ�Сд(&U)"
         Height          =   210
         Left            =   6060
         TabIndex        =   7
         Top             =   1605
         Width           =   1755
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6060
         TabIndex        =   0
         Top             =   495
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmApparatusItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLngAptId As Long          '��ǰ��ʾ������id
Private mstr���� As String          '��ǰ��Ŀ�ļ�������

Private Enum mCol
    ID = 0: ���: ����: ������: Ӣ����: ����: ͨ����: ����: ����ֵ: �����: ��������Ŀ
End Enum

Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 12: .FixedCols = 0
        End If
        .ColDataType(mCol.��������Ŀ) = flexDTBoolean
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.���) = "���": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.������) = "������": .TextMatrix(0, mCol.Ӣ����) = "Ӣ����": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.ͨ����) = "ͨ����": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.����ֵ) = "����ֵ": .TextMatrix(0, mCol.�����) = "�����": .TextMatrix(0, mCol.��������Ŀ) = "������Ŀ"
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.���) = 450: .ColWidth(mCol.����) = 800
        .ColWidth(mCol.������) = 1900: .ColWidth(mCol.Ӣ����) = 1400: .ColWidth(mCol.����) = 0
        .ColWidth(mCol.ͨ����) = 720: .ColWidth(mCol.����) = 510
        .ColWidth(mCol.����ֵ) = 630: .ColWidth(mCol.�����) = 630: .ColWidth(mCol.��������Ŀ) = 800
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .ColAlignment(mCol.���) = flexAlignCenterCenter
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.���) = lngCount
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngAptId As Long) As Boolean
    '���ܣ���������idˢ�µ�ǰ��ʾ����
    '��������ǰ��Ŀid
    Dim rsTemp As New ADODB.Recordset
    mLngAptId = lngAptId
    Me.txtFind.Text = ""
    Me.lvwItem.ListItems.Clear
        
    If lngAptId = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select L.������Ŀid As ID, Rownum As ���, I.����, I.���� As ������, L.��д As Ӣ����, L.������� As ����," & vbNewLine & _
            "       C.ͨ������ As ͨ����, C.С��λ�� As ��ȷ��, C.����ֵ, C.�����,C.��������Ŀ" & vbNewLine & _
            "From ����������Ŀ C, ������Ŀ L, ���鱨����Ŀ R, ������ĿĿ¼ I" & vbNewLine & _
            "Where C.��Ŀid = L.������Ŀid And L.������Ŀid = R.������Ŀid And R.������Ŀid = I.ID And I.�����Ŀ <> 1 And" & vbNewLine & _
            "    (I.����ʱ��>sysdate or I.����ʱ�� is null) And   L.��Ŀ��� <> 2 And C.����id = [1] order by i.���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAptId)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False
End Function

Public Function zlEditStart() As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ lngAptId-ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    mstr���� = ""
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select �������� From �������� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mLngAptId)
    If rsTemp.RecordCount > 0 Then mstr���� = "" & rsTemp!��������
    If mstr���� = "" Then Me.chkKind.Value = vbUnchecked
        
    Me.Tag = "�༭": Call Form_Resize
    If Me.Visible Then Me.txtFind.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mLngAptId)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim strLists As String, strItems As String, dblValue As Double
    Dim strListsS As String
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, mCol.ID)) = 0 Then
                MsgBox "��" & lngCount & "����Ŀ��ȷ����", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If Trim(.TextMatrix(lngCount, mCol.ͨ����)) = "" Then
                MsgBox "��" & lngCount & "�С�ͨ���롱δ��д��", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(.TextMatrix(lngCount, mCol.ͨ����)), vbFromUnicode)) > 20 Then
                MsgBox "��" & lngCount & "�С�ͨ���롱��������(20���ַ�)��", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            dblValue = Val(.TextMatrix(lngCount, mCol.����))
            If dblValue > 999999 Or Val(dblValue) - Int(Val(dblValue)) > 0 Then
                MsgBox "��" & lngCount & "�С����ȡ�̫��", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            dblValue = Val(.TextMatrix(lngCount, mCol.����ֵ))
            If dblValue > 999999 Or Val(dblValue * 100000) - Int(Val(dblValue * 100000)) > 0 Then
                MsgBox "��" & lngCount & "�С�����ֵ��̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            dblValue = Val(.TextMatrix(lngCount, mCol.�����))
            If dblValue > 999999 Or Val(dblValue * 100000) - Int(Val(dblValue * 100000)) > 0 Then
                MsgBox "��" & lngCount & "�С�����ȡ�̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            strItems = .TextMatrix(lngCount, mCol.ID)
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.ͨ����))
            If Val(.TextMatrix(lngCount, mCol.����)) = 1 Or Val(.TextMatrix(lngCount, mCol.����)) = 3 Then
                If Trim(.TextMatrix(lngCount, mCol.����)) = "" Then
                    strItems = strItems & ";"
                Else
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.����))
                End If
                If Trim(.TextMatrix(lngCount, mCol.����ֵ)) = "" Then
                    strItems = strItems & ";"
                Else
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.����ֵ))
                End If
                If Trim(.TextMatrix(lngCount, mCol.�����)) = "" Then
                    strItems = strItems & ";"
                Else
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.�����))
                End If
            Else
                strItems = strItems & ";;;"
            End If
            '������Ŀ
            strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.��������Ŀ))
            
            If LenB(strLists) < 3900 Then
                strLists = strLists & "|" & strItems
            Else
                strListsS = strListsS & "|" & strItems
            End If
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)


'    If LenB(gstrSql) > 4000 Then
'        MsgBox "������Ŀ����̫�࣬���ܱ��棡", vbInformation, gstrSysName
'        Me.vfgList.SetFocus: zlEditSave = 0: Exit Function
'    End If

    Err = 0: On Error GoTo ErrHand
   
    
    If strListsS <> "" Then
         '���ݱ���
         '����ַ�������4000���ַ������뵽strListsS��
        If strListsS <> "" Then strListsS = Mid(strListsS, 2)
        gstrSql = "Zl_����������Ŀ_Edit(" & mLngAptId & ",'" & strLists & "',0,'" & strListsS & "')"
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Else
        '���ݱ���
        gstrSql = "Zl_����������Ŀ_Edit(" & mLngAptId & ",'" & strLists & "')"
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    End If

    Me.Tag = "": Call Form_Resize
    zlEditSave = mLngAptId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0
End Function

Private Sub chkKind_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkUpper_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCurRow As Long, blnAdd As Boolean
    Dim strIDs As String '���������ӵ���Ŀ��ID
    Dim i As Long
    
    With Me.vfgList
        Select Case Index
        Case 0         '����
            If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
            Set objItem = Me.lvwItem.SelectedItem
            '���Ѵ�����Ŀ��ID���ӵ�������
            For i = 1 To .Rows
                strIDs = strIDs & "," & .TextMatrix(i - 1, mCol.ID) & ","
            Next
            '���ұ������Ƿ��Ѵ��ڸ���Ŀ
            If InStr(strIDs, "," & Mid(objItem.Key, 2) & ",") Then
                MsgBox """" & objItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1) & """��Ŀ�Ѵ���", vbInformation
                Exit Sub
            End If
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mCol.ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Rows - 1, mCol.����) = objItem.Text
            .TextMatrix(.Rows - 1, mCol.������) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1)
            .TextMatrix(.Rows - 1, mCol.Ӣ����) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1)
            .TextMatrix(.Rows - 1, mCol.����) = Left(objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1), 1)
            If objItem.Tag <> "" Then
                aryTemp = Split(objItem.Tag, "|")
                .TextMatrix(.Rows - 1, mCol.ͨ����) = aryTemp(0)
                .TextMatrix(.Rows - 1, mCol.����) = aryTemp(1)
                .TextMatrix(.Rows - 1, mCol.����ֵ) = aryTemp(2)
                .TextMatrix(.Rows - 1, mCol.�����) = aryTemp(3)
            End If
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
            Me.lvwItem.ListItems.Remove objItem.Key: Me.lvwItem.SetFocus
        Case 1          'ɾ��
            If .Row < .FixedRows Then Exit Sub
            '--  10802 ������Ŀʱ������ҳ�����Ŀ�б��д���Ҫ���ٵ���Ŀʱ����Ŀ�ڼ����в�Ψһ��
            '    ���������Ŀ�б����Ƿ��Ѵ��ڴ���Ŀ,���򲻼���ѡ���б�
            blnAdd = True
            If lvwItem.ListItems.Count > 1 Then
                For Each objItem In lvwItem.ListItems
                    If Val(Mid(objItem.Key, 2)) = Val(.TextMatrix(.Row, mCol.ID)) Then
                        blnAdd = False
                        Exit For
                    End If
                Next
            End If

            If blnAdd Then
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.����))
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1) = .TextMatrix(.Row, mCol.������)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1) = .TextMatrix(.Row, mCol.Ӣ����)
                Select Case Val(.TextMatrix(.Row, mCol.����))
                Case 1: objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = "1-����"
                Case 2: objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = "2-����"
                Case 3: objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = "3-�붨��"
                End Select
                objItem.Tag = .TextMatrix(.Row, mCol.ͨ����) & "|" & .TextMatrix(.Row, mCol.����)
                objItem.Tag = objItem.Tag & "|" & .TextMatrix(.Row, mCol.����ֵ) & "|" & .TextMatrix(.Row, mCol.�����)
                
                objItem.Selected = True
            End If
            .RemoveItem .Row
        End Select
        
        For lngCount = .Row To .Rows - 1
            .TextMatrix(lngCount, mCol.���) = lngCount
        Next
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String, strKind As String
    
    If Me.chkKind.Value = vbChecked Then
        strKind = "And I.�������� = '" & mstr���� & "'"
    Else
        strKind = ""
    End If
    
    If Me.chkUpper.Value = 0 Then
        strFind = DelInvalidChar(Trim(UCase(Me.txtFind.Text)))
        gstrSql = "Select L.������Ŀid As ID, I.����, I.���� As ������, L.��д As Ӣ����, L.������� As ����" & vbNewLine & _
                "From ������ĿĿ¼ I, ���鱨����Ŀ R, ������Ŀ L" & vbNewLine & _
                "Where I.ID = R.������Ŀid And R.������Ŀid = L.������Ŀid And I.�����Ŀ <> 1 " & strKind & " And" & vbNewLine & _
                "   (I.����ʱ��>sysdate or I.����ʱ�� is null) And    (I.���� Like '" & strFind & "%' Or Upper(I.����) Like '" & gstrMatch & strFind & "%' Or Upper(L.��д) Like '" & gstrMatch & strFind & "%')"
    Else
        strFind = DelInvalidChar(Trim(Me.txtFind.Text))
        gstrSql = "Select L.������Ŀid As ID, I.����, I.���� As ������, L.��д As Ӣ����, L.������� As ����" & vbNewLine & _
                "From ������ĿĿ¼ I, ���鱨����Ŀ R, ������Ŀ L" & vbNewLine & _
                "Where I.ID = R.������Ŀid And R.������Ŀid = L.������Ŀid And I.�����Ŀ <> 1 " & strKind & " And" & vbNewLine & _
                "   (I.����ʱ��>sysdate or I.����ʱ�� is null) And    (I.���� Like '" & strFind & "%' Or I.���� Like '" & gstrMatch & strFind & "%' Or L.��д Like '" & gstrMatch & strFind & "%')"
    End If
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF

            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1) = "" & !������
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1) = "" & !Ӣ����
            Select Case Val("" & !����)
            Case 1: objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = "1-����"
            Case 2: objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = "2-����"
            Case 3: objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = "3-�붨��"
            End Select
            objItem.Tag = ""

            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
'    With Me.vfgList
'        For lngCount = .FixedRows To .Rows - 1
'            Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mcol.ID)
'        Next
'    End With
    
    If Me.lvwItem.ListItems.Count = 0 Then
        MsgBox "û��ƥ�����Ŀ��", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vfgList.SetFocus
    End If
    Exit Sub

ErrHand:
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Resume
End Sub

Private Sub Form_Load()
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_����", "����", 900
        .Add , "_������", "������", 2300
        .Add , "_Ӣ����", "Ӣ����", 1500
        .Add , "_����", "����", 1000
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgList.ZOrder 0
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height - 105
    If Me.Tag = "�༭" Then
        Me.vfgList.Height = Me.picEdit.Top - Me.vfgList.Top
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True
        Me.vfgList.Editable = flexEDKbd: Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 105
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False
        Me.vfgList.Editable = flexEDNone: Me.vfgList.FocusRect = flexFocusNone
    End If
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItem
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwItem_DblClick()
    Call cmdEdit_Click(0)
End Sub

Private Sub picEdit_Resize()
    Err = 0: On Error Resume Next
    Me.lvwItem.Height = Me.picEdit.ScaleHeight - Me.lvwItem.Top
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click: Exit Sub
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag <> "�༭" Then Exit Sub
    With Me.vfgList
        If .TextMatrix(.Row, mCol.ͨ����) = "" Then
            Call cmdEdit_Click(1)
        Else
            If MsgBox("����������ͨ���룬��ȷ���Ƿ�Ҫɾ�����У�", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call cmdEdit_Click(1)
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22, vbKeyReturn: Exit Sub
    Case Else
        Select Case Col
        Case mCol.ͨ����
            If InStr(1, "|;'", Chr(KeyAscii)) = 0 Then Exit Sub
        Case mCol.����
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        Case mCol.����ֵ, mCol.�����
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or Chr(KeyAscii) = "." Then Exit Sub
        End Select
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case mCol.ID, mCol.���, mCol.����, mCol.������, mCol.Ӣ����, mCol.����: Cancel = True
    End Select
    If Row < Me.vfgList.FixedRows Then Cancel = True
End Sub


