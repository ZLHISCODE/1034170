VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "frmSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ListView lvw 
      Height          =   2850
      Left            =   2535
      TabIndex        =   2
      Top             =   555
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   5027
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6390
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6390
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   165
         Picture         =   "frmSelect.frx":014A
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   2355
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6390
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3660
      Width           =   6390
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4785
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3540
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   2745
         Top             =   15
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
               Picture         =   "frmSelect.frx":06D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   2760
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   4868
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4905
      TabIndex        =   8
      Top             =   315
      Width           =   435
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�룺SQL���ֶ�����
Public strSQLList As String
Public strSQLTree As String
Public strFLDList As String
Public strFLDTree As String
Public strParName As String '��������
Public bytType As Byte      '������������
Public strMatch As String '����ƥ�������
Public lngSeekHwnd As Long '���ڶ�λ����λ�õĿؼ�
Public mintConnect As Integer           '�������ӱ��

Public mblnMulti As Boolean '�Ƿ��ѡ��
Public mblnOK As Boolean
Public mlngSel As Long  '���е�ֵ�������ֵʱѡ��

'����δ����ʽ���������ԭʼֵ
Public strOutBand As String 'ѡ��İ�ֵ,��Ӧ&B
Public strOutDisp As String 'ѡ�����ʾֵ,��Ӧ&D

Private intPreNode As Long
Private blnItem As Boolean
Private blnSetFlex As Boolean, blnSetLvw As Boolean
Private rsList As ADODB.Recordset
Private strList As String
Private BlnSave As Boolean
Private rParent As RECT

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim strDisp As String, strBand As String
    
    On Error GoTo hErr
    
    strDisp = GetScript(strFLDList, "&D") '��ʾ���ֶ���
    strBand = GetScript(strFLDList, "&B") '�󶨵��ֶ���
    
    If strDisp = "" Or strBand = "" Then
        MsgBox "ѡ������û�ж��������İ󶨼���ʾ�ֶ���Ŀ��", vbInformation, App.Title
        Exit Sub
    End If
    
    If mblnMulti Then
        '��ѡʱ�Զ����ص����
        If Not lvw.Visible And lvw.ListItems.count = 1 Then
            lvw.ListItems(1).Checked = True
        End If
        
        For i = 1 To lvw.ListItems.count
            If lvw.ListItems(i).Checked Then
                If Split(lvw.ListItems(i).Tag, "|")(0) = "" Then
                    lvw.ListItems(i).Selected = True
                    lvw.ListItems(i).EnsureVisible
                    MsgBox "�������ݵ�""" & strDisp & """Ϊ��,����������""" & strParName & """����ʾ��", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
                If Split(lvw.ListItems(i).Tag, "|")(1) = "" Then
                    lvw.ListItems(i).Selected = True
                    lvw.ListItems(i).EnsureVisible
                    MsgBox "�������ݵ�""" & strBand & """Ϊ��,����������""" & strParName & """��󶨣�", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
                '���ͼ��(������,������)
            End If
        Next
        '������ʾ�����󶨴�
        strOutDisp = ""
        strOutBand = ""
        For i = 1 To lvw.ListItems.count
            If lvw.ListItems(i).Checked Then
                strOutDisp = strOutDisp & "," & Split(lvw.ListItems(i).Tag, "|")(0)
                strOutBand = strOutBand & "," & Split(lvw.ListItems(i).Tag, "|")(1)
            End If
        Next
        If strOutDisp = "" Or strOutBand = "" Then
            MsgBox "û��ѡ���κ����ݣ�", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        ElseIf UBound(Split(strOutBand, ",")) > 1000 Then
            MsgBox "ѡ������ݹ��࣡", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        strOutDisp = Mid(strOutDisp, 2)
        strOutBand = " IN (" & Mid(strOutBand, 2) & ") "
    Else
        If lvw.SelectedItem Is Nothing Then
            MsgBox "û��ѡ���κ����ݣ�", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        If InStr(lvw.SelectedItem.Tag, "|") <= 0 Then
            MsgBox "�������ݵ�Ϊ�գ���������Դ��", vbInformation, App.Title
            Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(0) = "" Then
            MsgBox "�������ݵ�""" & strDisp & """Ϊ��,����������""" & strParName & """����ʾ��", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(1) = "" Then
            MsgBox "�������ݵ�""" & strBand & """Ϊ��,����������""" & strParName & """��󶨣�", vbInformation, App.Title
            lvw.SetFocus: Exit Sub
        End If
        
        '���ͼ��
        Select Case bytType
            Case 1
                If Not IsNumeric(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
            Case 2
                If Not IsDate(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    lvw.SetFocus: Exit Sub
                End If
        End Select
    
        strOutDisp = Split(lvw.SelectedItem.Tag, "|")(0)
        strOutBand = Split(lvw.SelectedItem.Tag, "|")(1)
    End If
    
    mblnOK = True
    
    On Error Resume Next
    Hide
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub Form_Activate()
    If tvw_s.Visible Then
        If Not tvw_s.SelectedItem Is Nothing Then
            If tvw_s.SelectedItem.Key = "ALL" Then lvw.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Not lvw.Visible Then Exit Sub
        
        For i = 1 To lvw.ListItems.count
            lvw.ListItems(i).Checked = True
        Next
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Not lvw.Visible Then Exit Sub
        
        For i = 1 To lvw.ListItems.count
            lvw.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim lngW As Long, i As Integer
    
    If Not InDesign Then
        glngSelProc = GetWindowLong(hwnd, GWL_WNDPROC)
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SelMessage)
    End If
    
    mblnOK = False
    BlnSave = True
    blnSetFlex = False '�Ƿ��Ѿ��Ա��ָ����
    blnSetLvw = False
    intPreNode = 0
    
    strOutBand = ""
    strOutDisp = ""
    
    lvw.Tag = strParName
    
    Me.Caption = strParName & "ѡ����"
    
    strSQLList = Replace(strSQLList, "[*]", strMatch)
    strSQLTree = Replace(strSQLTree, "[*]", strMatch)
    
    If strSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
        If Not FillList Then BlnSave = False: Unload Me: Exit Sub
    Else
        tvw_s.Visible = True
        If Not FillTree Then BlnSave = False: Unload Me: Exit Sub
        If tvw_s.Nodes.count > 0 Then
            tvw_s.Nodes(1).Selected = True
            If Not tvw_s.Nodes(1).Child Is Nothing And strMatch = "" Then
                tvw_s.Nodes(1).Child.Selected = True
            End If
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        End If
    End If
    
    lvw.Checkboxes = mblnMulti
    lvw.ToolTipText = "ȫѡ(Ctrl+A),ȫ��(Ctrl+R)"
    
    '����ƥ���Զ�����
    If strMatch <> "" Then
        If rsList.RecordCount = 1 Then
            BlnSave = False
            Call cmdOK_Click
            Unload Me: Exit Sub
        ElseIf rsList.RecordCount = 0 Then
            MsgBox "û���ҵ���ƥ�����Ŀ,���������룡", vbInformation, App.Title
            BlnSave = False
            Call cmdCancel_Click: Exit Sub
        End If
    End If
    
    Call Form_Resize
    
    '���弰�б�ȱʡ���
    If lvw.ColumnHeaders.count = 1 Then
        lvw.ColumnHeaders(1).Width = 2500
        Me.Width = 3000 + IIF(strSQLTree = "", 0, tvw_s.Width + pic.Width)
    Else
        For i = 1 To lvw.ColumnHeaders.count
            lngW = lngW + lvw.ColumnHeaders(i).Width
        Next
        Me.Width = lngW + 500 + IIF(strSQLTree = "", 0, tvw_s.Width + pic.Width)
        If Me.Width < 3000 Then Me.Width = 3000
    End If
    
    If strSQLTree <> "" Then
        If Me.Width < (tvw_s.Width + pic.Width) * 2.2 Then Me.Width = (tvw_s.Width + pic.Width) * 2.2
    End If
    
    RestoreWinState Me, App.ProductName, strParName
    
    If strSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
    Else
        tvw_s.Visible = True
    End If
    
    '��λ
    If lngSeekHwnd <> 0 Then
        Call Form_Resize
        GetWindowRect lngSeekHwnd, rParent
        If rParent.Top >= Me.Height / 15 Then
            Me.Top = rParent.Bottom * 15 - Me.Height + 30
        Else
            Me.Top = (rParent.Bottom - rParent.Top) * 15 + 30
        End If
        If rParent.Left >= Me.Width / 15 Then
            Me.Left = rParent.Right * 15 - Me.Width + 30
        Else
            Me.Left = (rParent.Right - rParent.Left) * 15 + 30
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim lngTVW As Long
    lngTVW = IIF(tvw_s.Visible, tvw_s.Width + pic.Width, 0)
    
    tvw_s.Left = Me.ScaleLeft
    tvw_s.Top = picInfo.Top + picInfo.Height + 15
    tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height - 15
    
    pic.Left = tvw_s.Left + tvw_s.Width
    pic.Top = tvw_s.Top
    pic.Height = tvw_s.Height
    
    lvw.Left = Me.ScaleLeft + lngTVW
    lvw.Top = tvw_s.Top
    lvw.Height = tvw_s.Height
    lvw.Width = Me.ScaleWidth - lngTVW
    
    lbl.Left = lvw.Left
    lbl.Top = lvw.Top
    lbl.Width = lvw.Width
    lbl.Height = lvw.Height
    
    If ScaleWidth - cmdCancel.Width - 300 >= 1445 Then
        cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 45
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strMatch = ""
    lngSeekHwnd = 0
    If BlnSave Then SaveWinState Me, App.ProductName, strParName
    If Not InDesign Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngSelProc)
End Sub

Private Sub lvw_DblClick()
    If blnItem Then Call cmdOK_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnItem = True
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub lvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnItem = False
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or lvw.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        lvw.Left = lvw.Left + X
        lvw.Width = lvw.Width - X
        
        lbl.Left = lbl.Left + X
        lbl.Width = lbl.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Index = intPreNode Then Exit Sub
    intPreNode = Node.Index
    DoEvents
    Call FillList(Node.Tag)
End Sub

Private Function FillTree() As Boolean
'���ܣ����ݶ�������Դ���ֶ����ԣ�������������ʾ��TreeView��
'���أ������Ƿ�ɹ�(�û�����������)
    Dim rstmp As New ADODB.Recordset
    Dim i As Integer, objNode As Node
    Dim strSel As String, strRela As String
    
    On Error GoTo errH
    
    strSel = GetScript(strFLDTree, "&S")
    strRela = GetScript(strFLDTree, "&R")
    
    If strSel = "" Or strRela = "" Then
        MsgBox "δ��������ѡ�������ϸ�б���������ֶ���Ŀ��", vbInformation, App.Title
        Exit Function
    End If
    Call OpenRecord(rstmp, RemoveNote(strSQLTree), Me.Caption & "_FillTree", mintConnect) 'SQLһ��̶�,[*]��SQL��''��,�����޷�����
    
    tvw_s.Nodes.Clear
        
    If InStr("|" & UCase(strFLDTree), "|ID,") > 0 And InStr("|" & UCase(strFLDTree), "|�ϼ�ID,") > 0 Then
        '���������б���ʾ
        Set objNode = tvw_s.Nodes.Add(, , "ALL", "������Ŀ", 1)
        objNode.Tag = "ALL"
        objNode.Expanded = True
        
        For i = 1 To rstmp.RecordCount
            If IsNull(rstmp!�ϼ�ID) Then
                Set objNode = tvw_s.Nodes.Add("ALL", 4, "_" & rstmp!id, IIF(IsNull(rstmp.Fields(strSel).Value), "", rstmp.Fields(strSel).Value), 1)
            Else
                Set objNode = tvw_s.Nodes.Add("_" & rstmp!�ϼ�ID, 4, "_" & rstmp!id, IIF(IsNull(rstmp.Fields(strSel).Value), "", rstmp.Fields(strSel).Value), 1)
            End If
            objNode.Tag = IIF(IsNull(rstmp.Fields(strRela).Value), "", rstmp.Fields(strRela).Value)
            rstmp.MoveNext
        Next
    Else
        '����һ���б���ʾ
        For i = 1 To rstmp.RecordCount
            Set objNode = tvw_s.Nodes.Add(, , , IIF(IsNull(rstmp.Fields(strSel).Value), "", rstmp.Fields(strSel).Value), 1)
            objNode.Tag = IIF(IsNull(rstmp.Fields(strRela).Value), "", rstmp.Fields(strRela).Value)
            rstmp.MoveNext
        Next
    End If

    FillTree = True
    Exit Function
errH:
    If Err.Number = 35601 Then
        MsgBox "�����������������б�����ѡ��������ʹ�ã�", vbExclamation, App.Title
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Private Function GetRelaSQL(ByVal strSQL As String, ByVal strFld As String, ByVal strKey As String) As String
'���ܣ����������SQL
    Dim i As Integer, strRela As String
    
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), "&R") > 0 Then
            strRela = Split(Split(strFld, "|")(i), ",")(0)
            If strKey = "" Then
                GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & " is NULL"
            Else
                Select Case Split(Split(strFld, "|")(i), ",")(1)
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=" & strKey
                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "='" & strKey & "'"
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                        If Format(strKey, "hh:mm:ss") = "00:00:00" Then
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & ">=To_Date('" & Format(strKey, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & strRela & "<=To_Date('" & Format(strKey, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=To_Date('" & Format(strKey, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                End Select
            End If
            Exit Function
        End If
    Next
End Function

Private Function GetScript(strFld As String, strType As String) As String
'���ܣ�����ָ�����ֶ����������ֶ���
'������strType="&S &D &B &R"
'˵����������Ψһ�������ֶ�(����ֶ�)
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), strType) > 0 Then
            GetScript = Split(Split(strFld, "|")(i), ",")(0)
            Exit Function
        End If
    Next
End Function

Private Function HaveScript(strFld As String, strName As String, strType As String) As Boolean
'���ܣ��ж����ֶ������У�ָ�����ֶ��Ƿ����ָ������������
'������strName=�ֶ���,strFld=�ֶ�������,strType="&S &D &B &R"
'���أ�False=δ�����ֶλ��ֶβ�����ָ������
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If Split(Split(strFld, "|")(i), ",")(0) = strName Then
            If InStr(Split(Split(strFld, "|")(i), ",")(2), strType) > 0 Then
                HaveScript = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function FillList(Optional strKey As String, Optional blnSort As Boolean) As Boolean
'���ܣ����ݵ�ǰѡ��ķ�������޷���ʱ�����Ӧ����ϸ�б�
'������strKey=�����б��еĵ�ǰ����ֵ
'˵���������������Ķ��٣�ȷ����ListView����DataGrid
    Dim strSQL As String, i As Long, j As Integer
    Dim objitem As ListItem, strValue As String
    Dim strDisp As String, strBand As String
    
    On Error GoTo errH
    
    lvw.ListItems.Clear
    
    '����Ϊֻ��������
    If Not blnSort Then
        If strSQLTree = "" Then
            strSQL = strSQLList
        Else
            '��̬����ϸ���ݴ���Ϊֻ��ȡ�����ķ��ಿ��(���� Order by �Ӿ�)
            If strKey = "ALL" Then
                strSQL = strSQLList
            Else
                strSQL = GetRelaSQL(RemoveOrderBy(strSQLList), strFLDList, strKey)
            End If
            
            If strSQL = "" Then
                MsgBox "�������ݶ�ȡʧ�ܣ�", vbInformation, App.Title
                Exit Function
            End If
        End If
        
        Screen.MousePointer = 11
        Me.Refresh
        
        Set rsList = New ADODB.Recordset
        Call OpenRecord(rsList, RemoveNote(strSQL), Me.Caption & "_FillList", mintConnect) 'SQLһ��̶�,[*]��SQL��''��,�����޷�����
    End If
    
    If Not rsList.EOF Then
        If lvw.ColumnHeaders.count = 0 Then Call AddListCols
        
        strDisp = GetScript(strFLDList, "&D") '��ʾֵ��Ŀ
        strBand = GetScript(strFLDList, "&B") '��ֵ��Ŀ
        
        For i = 1 To rsList.RecordCount
            strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(1).Text))
            If lvw.ColumnHeaders(1).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(1).Tag)
            Set objitem = lvw.ListItems.Add(, , strValue, , 1)
            For j = 2 To lvw.ColumnHeaders.count
                strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(j).Text))
                If lvw.ColumnHeaders(j).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(j).Tag)
                objitem.SubItems(j - 1) = strValue
            Next
            
            '����ʾֵ����ֵ������TAG��,��Ϊ��һ����Щ�ֶλ�Ϊѡ���ֶ�
            '��ʽΪ"��ʾֵ|��ֵ"
            If strDisp <> "" Then
                objitem.Tag = IIF(IsNull(rsList.Fields(strDisp).Value), "", rsList.Fields(strDisp).Value)
            End If
            objitem.Tag = objitem.Tag & "|"
            If strBand <> "" Then
                objitem.Tag = objitem.Tag & IIF(IsNull(rsList.Fields(strBand).Value), "", rsList.Fields(strBand).Value)
                If mlngSel <> 0 And Val(rsList.Fields(strBand).Value & "") = mlngSel Then objitem.Selected = True: Call objitem.EnsureVisible
            End If
                            
            rsList.MoveNext
        Next
        
        '�Զ������п�
        Call AutoSizeCol(lvw)
        
        If Not Visible Or Not blnSetLvw Then
            Call RestoreListViewState(lvw, App.ProductName & "\" & Me.name & strParName)
            blnSetLvw = True
        End If
        lblInfo.Caption = "�� " & rsList.RecordCount & " ����ϸ��Ŀ."
    Else
        'û������ʱ����ʾ�յ�ListView(����ͷ)
        If lvw.ColumnHeaders.count = 0 Then Call AddListCols
        lblInfo.Caption = "û����ϸ��Ŀ."
    End If
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddListCols()
'���ܣ�����strFLDList�ֶ�����ֵ,ΪListView������ͷ
    Dim i As Integer, j As Integer, strFld As String
    Dim objCol As ColumnHeader
    
    For i = 0 To UBound(Split(strFLDList, "|"))
        strFld = Split(strFLDList, "|")(i)
        If strFld Like "*&S*" Then
            Set objCol = lvw.ColumnHeaders.Add(, "_" & Split(strFld, ",")(0), Split(strFld, ",")(0))
            
            objCol.Width = Me.TextWidth(Split(strFld, ",")(0) & "��")
            
            '�����ֶ������������ö���(��1ֻ�������)
            Select Case Split(strFld, ",")(1)
                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    If rsList.Fields(objCol.Text).NumericScale > 0 Then
                        j = rsList.Fields(objCol.Text).NumericScale
                        objCol.Tag = "0." & String(IIF(j > 2, 2, j), "0; ;")
                        If objCol.Index <> 1 Then objCol.Alignment = lvwColumnRight
                    ElseIf objCol.Index <> 1 Then
                        If rsList.Fields(objCol.Text).Precision < 3 Then
                            objCol.Alignment = lvwColumnCenter
                        Else
                            objCol.Alignment = lvwColumnLeft
                        End If
                    End If
                    If objCol.Text Like "*��" Then objCol.Tag = "0.000"
                    If objCol.Text Like "*��" Then objCol.Tag = "0.00"
                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
                Case Else
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
            End Select
            If objCol.Text Like "*��λ*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
            If objCol.Text Like "*��*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
        End If
    Next
End Sub

Private Function GetValue(objFld As Field) As String
'����:�����ֶ�����ȡ���ʵ���ʾֵ
    Dim strValue As String
    Select Case objFld.type
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            strValue = IIF(IsNull(objFld.Value), 0, objFld.Value)
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
            If Format(strValue, "HH:mm:ss") = "00:00:00" Then
                strValue = Format(strValue, "yyyy-MM-dd")
            Else
                strValue = Format(strValue, "yyyy-MM-dd HH:mm:ss")
            End If
        Case Else
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
    End Select
    GetValue = strValue
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'���ܣ���������
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub
