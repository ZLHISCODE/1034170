VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetBalance 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   ControlBox      =   0   'False
   Icon            =   "frmSetBalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboDiagnose 
      Height          =   300
      Left            =   4935
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   120
      Width           =   2835
   End
   Begin VB.CheckBox chkKind 
      Caption         =   "������"
      Height          =   255
      Index           =   1
      Left            =   5265
      TabIndex        =   24
      Top             =   570
      Width           =   1095
   End
   Begin VB.ComboBox cboBabyFee 
      Height          =   300
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   23
      ToolTipText     =   "Ŀǰ��֧��ÿ���������5��Ӥ��"
      Top             =   547
      Width           =   1410
   End
   Begin VB.CheckBox chkKind 
      Caption         =   "��ͨ����"
      Height          =   255
      Index           =   0
      Left            =   4185
      TabIndex        =   22
      Top             =   570
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����ڼ�"
      Height          =   780
      Left            =   120
      TabIndex        =   16
      Top             =   45
      Width           =   3975
      Begin zl9InExse.ctlDate dtpBegin 
         Height          =   300
         Left            =   885
         TabIndex        =   0
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         Value           =   40212
         MaxDate         =   2958101
         MinDate         =   36526
      End
      Begin VB.CommandButton cmdSyn 
         Caption         =   "ͬ��"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "��ָ��סԺ���������Ժʱ��ͬ��������ֹʱ��"
         Top             =   263
         Width           =   510
      End
      Begin zl9InExse.ctlDate dtpEnd 
         Height          =   300
         Left            =   2505
         TabIndex        =   1
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         Value           =   40212
         MaxDate         =   2958101
         MinDate         =   36526
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   2295
         TabIndex        =   19
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   690
         TabIndex        =   18
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   135
      TabIndex        =   12
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6660
      TabIndex        =   11
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5340
      TabIndex        =   10
      Top             =   4470
      Width           =   1100
   End
   Begin VB.Frame fraWhile 
      Caption         =   "��������"
      Height          =   3360
      Left            =   120
      TabIndex        =   13
      Top             =   870
      Width           =   7635
      Begin MSComctlLib.ListView lvwChargeType 
         Height          =   2400
         Left            =   6015
         TabIndex        =   7
         Top             =   510
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   4233
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdȫѡ 
         Caption         =   "ȫѡ"
         Height          =   300
         Index           =   0
         Left            =   6015
         TabIndex        =   8
         Top             =   2970
         Width           =   510
      End
      Begin VB.CommandButton cmdȫ�� 
         Caption         =   "ȫ��"
         Height          =   300
         Index           =   0
         Left            =   6555
         TabIndex        =   9
         Top             =   2970
         Width           =   510
      End
      Begin VB.CommandButton cmdȫ�� 
         Caption         =   "ȫ��"
         Height          =   300
         Index           =   2
         Left            =   4950
         TabIndex        =   29
         Top             =   2970
         Width           =   510
      End
      Begin VB.CommandButton cmdȫѡ 
         Caption         =   "ȫѡ"
         Height          =   300
         Index           =   2
         Left            =   4410
         TabIndex        =   28
         Top             =   2970
         Width           =   510
      End
      Begin VB.CommandButton cmdȫ�� 
         Caption         =   "ȫ��"
         Height          =   300
         Index           =   1
         Left            =   3465
         TabIndex        =   27
         Top             =   2970
         Width           =   510
      End
      Begin VB.CommandButton cmdȫѡ 
         Caption         =   "ȫѡ"
         Height          =   300
         Index           =   1
         Left            =   2925
         TabIndex        =   26
         Top             =   2970
         Width           =   510
      End
      Begin VB.ListBox lstClass 
         Height          =   2370
         Left            =   4410
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   525
         Width           =   1545
      End
      Begin VB.ListBox lstItem 
         Height          =   2370
         Left            =   2910
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   525
         Width           =   1425
      End
      Begin VB.ListBox lstUnit 
         Height          =   2370
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   525
         Width           =   1545
      End
      Begin VB.ListBox lstTime 
         Height          =   2370
         Left            =   135
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   525
         Width           =   1155
      End
      Begin VB.Label lbl�շ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շ����"
         Height          =   180
         Left            =   6030
         TabIndex        =   6
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   4425
         TabIndex        =   21
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ"
         Height          =   180
         Left            =   2925
         TabIndex        =   17
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ÿ���"
         Height          =   180
         Left            =   1365
         TabIndex        =   15
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         Height          =   180
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Left            =   4185
      TabIndex        =   30
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frmSetBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'��ڲ���
Public mlngInsure As Long '�Ƿ�ҽ����������
Public mbytFuns As Byte '0-���ﲡ��;1-סԺ����
Public mlngPatient As Long '����ID
Public mstrALLChargeType As String '�շ����
Public mstrAllTime As String
Public mstrAllUnit As String
Public mstrALLItem As String
Public mstrAllClass As String
Public mstrAllDiagnose As String
Public mbytKind As Byte  '0-����ͨ����,1-��������,2-��ͨ���ú�������
Public mMinDate As Date, mMaxDate As Date
Public mblnOk As Boolean
Public mblnNOCancel As Boolean
Public mstrUnAuditTime As String 'δ��˵�סԺ����,ȫ��δ���ʱ��������������,�С���δ��˲��˽��ʡ�Ȩ��ʱ�������
Public mbln������ʽ��� As Boolean  'True
Public mstrTime As String
Public mbytFunc As Byte '0-����;1-סԺ
Private mblnDBegin As Boolean   'ҽ�������Ƿ������޸�ʱ�䷶Χ
Private mblnDEnd As Boolean
Private mblnNotClick As Boolean

Private Sub chkKind_Click(Index As Integer)
    If Visible And chkKind(0).Value = 0 And chkKind(1).Value = 0 Then
        chkKind(Index).Value = 1
    End If
    
    '����������ʱ,�����ڼ�
    If chkKind(0).Value = 0 And chkKind(1).Value = 1 Then
        dtpBegin.Enabled = False
        dtpEnd.Enabled = False
    Else
        dtpBegin.Enabled = mblnDBegin
        dtpEnd.Enabled = mblnDEnd
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    mblnOk = True
    Hide
End Sub
Private Function GetInOutDate(lngPati As Long, lngTimes As Long, bytMode As Byte) As Date
'����:��ȡ����ĳ��סԺ����Ժ���Ժʱ��
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select ��Ժ����,��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPati, lngTimes)
    If rsTmp.RecordCount > 0 Then
        If bytMode = 0 Then
            GetInOutDate = rsTmp!��Ժ����
        Else
            GetInOutDate = IIf(IsNull(rsTmp!��Ժ����), CDate("0:00:00"), rsTmp!��Ժ����)
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetBookInDate(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional bytMode As Byte) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵ĵǼ�ʱ��
    '����:�Ǽ�ʱ��,��ʽ:yyyy-mm-dd HH:MM:SS
    '����:���˺�
    '����:2013-10-22 17:16:47
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
        
    On Error GoTo errHandle
    
    strSql = " Select " & IIf(bytMode = 0, "Max", "Min") & IIf(gint����ʱ�� = 0, "(�Ǽ�ʱ��)", "(����ʱ��)")
    strSql = strSql & " As �Ǽ�ʱ�� From סԺ���ü�¼ Where Mod(��¼����,10) In (2,3) And ����ID=[1] And ��ҳID=[2]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳID)
    GetBookInDate = Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdSyn_Click()
    Dim i As Long, lngMax As Long, lngMin As Long, DatTmp As Date
    Dim strBookInDate As String
    If Not lstTime.Visible Then Exit Sub
    
    For i = 0 To lstTime.ListCount - 1
        If lstTime.Selected(i) = True Then
            If lngMax = 0 Then lngMax = lstTime.ItemData(i)
            If lngMin = 0 Then lngMin = lstTime.ItemData(i)
            
            If lngMax < lstTime.ItemData(i) Then
                lngMax = lstTime.ItemData(i)
            End If
            If lngMin > lstTime.ItemData(i) Then
                lngMin = lstTime.ItemData(i)
            End If
        End If
    Next
    
    If lngMin = 0 And lngMax = 0 Then
        MsgBox "����ѡ��סԺ����!", vbInformation, Me.Caption
        Exit Sub
    End If
 
    If lngMin <> 0 Then
        DatTmp = GetInOutDate(mlngPatient, lngMin, 0)
        If DatTmp <> CDate("0:00:00") Then
            '��ȡ�Ǽ�ʱ��,�Ǽ�ʱ�����Ժʱ��ҪС,�ԵǼ�ʱ��Ϊ׼:107022
            strBookInDate = GetBookInDate(mlngPatient, lngMin, 1)
            If strBookInDate <> "" Then
                 If Format(DatTmp, "yyyy-mm-dd HH:MM:SS") > strBookInDate Then
                    DatTmp = CDate(strBookInDate)
                 End If
            End If
            dtpBegin.Value = DatTmp
        Else
            dtpBegin.Value = zlDatabase.Currentdate
        End If
    End If
    
    If lngMax <> 0 Then
        DatTmp = GetInOutDate(mlngPatient, lngMax, 1)
        strBookInDate = GetBookInDate(mlngPatient, lngMax, 0)
        If DatTmp <> CDate("0:00:00") Then
            '��ȡ�Ǽ�ʱ��,�Ǽ�ʱ��ȳ�Ժʱ��Ҫ��,�ԵǼ�ʱ��Ϊ׼:63594
            If strBookInDate <> "" Then
                 If Format(DatTmp, "yyyy-mm-dd,HH:MM:SS") < strBookInDate Then
                    DatTmp = CDate(strBookInDate)
                 End If
            End If
            dtpEnd.Value = DatTmp
        Else
            If strBookInDate <> "" Then
                 dtpEnd.Value = CDate(strBookInDate)
            End If
            If dtpBegin.Value > dtpEnd.Value Then
                dtpEnd.Value = zlDatabase.Currentdate
            End If
        End If
    End If
End Sub

Private Sub cmdȫ��_Click(Index As Integer)
        Select Case Index
        Case 0  '�շ����
            Call SetlvwItem(lvwChargeType, False)
        Case 1  '������Ŀ
            Call SetListbox(lstItem, False)
        Case 2  '��������
            Call SetListbox(lstClass, False)
        End Select
End Sub

Private Sub cmdȫѡ_Click(Index As Integer)
        Select Case Index
        Case 0  '�շ����
            Call SetlvwItem(lvwChargeType, True)
        Case 1  '������Ŀ
            Call SetListbox(lstItem, True)
        Case 2  '��������
            Call SetListbox(lstClass, True)
        End Select
End Sub

Private Sub SetlvwItem(ByVal objLVW As Object, Optional blnAllCheck As Boolean = False)
    Dim i As Long, objItem As ListItem
    mblnNotClick = True
    For Each objItem In objLVW.ListItems
        objItem.Checked = blnAllCheck
    Next
    mblnNotClick = False
End Sub
Private Sub SetListbox(ByVal objList As ListBox, Optional blnAllCheck As Boolean = False)
    Dim i As Long
    mblnNotClick = True
    For i = 0 To objList.ListCount - 1
        objList.Selected(i) = blnAllCheck
    Next
    mblnNotClick = False
    
End Sub
Private Sub dtpBegin_LastDayInput()
        zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub dtpBegin_CmdDownClick()
    Dim dtDate  As Date
    dtDate = dtpBegin.Value
    If frmDownDate.ShowDate(dtpBegin, dtpBegin.MaxDate, dtpBegin.MinDate, dtDate) = False Then Exit Sub
    dtpBegin.Value = dtDate
    If dtpBegin.Enabled Then dtpBegin.SetFocus
End Sub

Private Sub dtpEnd_CmdDownClick()
    Dim dtDate As Date
    dtDate = dtpEnd.Value
    If frmDownDate.ShowDate(dtpEnd, dtpEnd.MaxDate, dtpEnd.MinDate, dtDate) = False Then Exit Sub
    dtpEnd.Value = dtDate
     If dtpEnd.Enabled Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_LastDayInput()
        zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    cmdCancel.Visible = Not mblnNOCancel
    If mblnNOCancel Then
        cmdOK.Left = cmdCancel.Left
    Else
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    End If
    
    If gbln���סԺ������������ Then
        If Not lstTime.Visible Then
            If dtpBegin.Enabled Then dtpBegin.SetFocus
        Else
            If lstTime.Enabled Then lstTime.SetFocus
        End If
    Else
        If dtpBegin.Enabled Then dtpBegin.SetFocus
    End If
    Call cmdSyn_Click
End Sub

Private Sub Form_Load()
    Dim i As Long, rsTemp As ADODB.Recordset
    Dim strTmp As String
    Dim j As Long
    
    mblnOk = False
    
    lstUnit.Clear
    lstTime.Clear
    'סԺ������Χ
    Me.Caption = IIf(mbytFunc = 0, "�����������", "סԺ��������")
    If mbytFunc = 0 Then
         lstTime.AddItem "����"
         lstTime.ItemData(lstTime.NewIndex) = 0
         lstTime.Selected(lstTime.NewIndex) = True
         lbl����.Left = lblTime.Left
         lblTime.Visible = False: lstTime.Visible = False
         lstUnit.Left = lstTime.Left: lstUnit.Width = lstItem.Left - lstUnit.Left - 50
    Else
            If mstrAllTime <> "" Then
                If mstrAllTime = "0" Then cmdSyn.Enabled = False
                j = 0
                For i = 0 To UBound(Split(mstrAllTime, ","))
                    strTmp = Split(mstrAllTime, ",")(i)
                    If strTmp <> 0 Then
                    
                        lstTime.AddItem IIf(strTmp = "0", "����", "��" & strTmp & "��")
                        lstTime.ItemData(lstTime.NewIndex) = strTmp
                        'ҽ������ֻ��ѡ��һ��סԺ�ķ���
                        If mlngInsure = 0 Or j = 0 Then
                            If InStr(1, "," & mstrUnAuditTime & ",", "," & strTmp & ",") > 0 Then
                                lstTime.Selected(j) = False
                            Else
                                lstTime.Selected(j) = True
                            End If
                            mblnNotClick = True
                            If mstrTime <> "" And InStr(1, "," & mstrTime & ",", "," & strTmp & ",") = 0 Then
                                lstTime.Selected(j) = False
                            End If
                            mblnNotClick = False
                        End If
                        j = j + 1
                    End If
                Next
                If lstTime.ListCount > 0 Then lstTime.ListIndex = 0
            End If
    End If
    '���ÿ��ҷ�Χ
    If mstrAllUnit <> "" Then
        For i = 0 To UBound(Split(mstrAllUnit, ","))
            lstUnit.AddItem Split(Split(mstrAllUnit, ",")(i), ":")(1)
            lstUnit.ItemData(lstUnit.ListCount - 1) = Split(Split(mstrAllUnit, ",")(i), ":")(0)
            lstUnit.Selected(i) = True
        Next
        If lstUnit.ListCount > 0 Then lstUnit.ListIndex = 0
    End If
    '�վݷ�Ŀ��Χ
    If mstrALLItem <> "" Then
        For i = 0 To UBound(Split(mstrALLItem, ","))
            lstItem.AddItem Mid(Split(mstrALLItem, ",")(i), 2, Len(Split(mstrALLItem, ",")(i)) - 2)
            lstItem.Selected(i) = True
        Next
        If lstItem.ListCount > 0 Then lstItem.ListIndex = 0
    End If
    '�������ͷ�Χ
    lstClass.AddItem "��������"
    lstClass.Selected(0) = True
    If mstrAllClass <> "" Then
        For i = 0 To UBound(Split(mstrAllClass, ","))
            lstClass.AddItem Mid(Split(mstrAllClass, ",")(i), 2, Len(Split(mstrAllClass, ",")(i)) - 2)
            lstClass.Selected(lstClass.NewIndex) = True
        Next
    End If
    lstClass.ListIndex = 0
    
    '�����Ŀ��Χ
    cboDiagnose.AddItem "�������"
    cboDiagnose.ListIndex = 0
    If mstrAllDiagnose <> "" Then
        For i = 0 To UBound(Split(mstrAllDiagnose, ","))
            cboDiagnose.AddItem Split(mstrAllDiagnose, ",")(i)
        Next
    End If
    
    '�շ����:34260
    Dim objListItem As ListItem
    lvwChargeType.ListItems.Clear
     Set objListItem = lvwChargeType.ListItems.Add(, "ALL", "�������")
     objListItem.Checked = True
     If mstrALLChargeType <> "" Then
        Set rsTemp = zlGet�շ����
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                 If InStr(1, "," & mstrALLChargeType & ",", ",'" & Nvl(rsTemp!����) & "',") > 0 Then
                    Set objListItem = lvwChargeType.ListItems.Add(, "K" & Nvl(rsTemp!����), Nvl(rsTemp!���))
                    objListItem.Checked = True
                 End If
                .MoveNext
            Loop
        End With
    End If
    
    dtpBegin.Value = mMinDate
    dtpEnd.Value = mMaxDate
    dtpBegin.MaxDate = dtpEnd.Value
    
    strTmp = "���˼�Ӥ������|�����˷���|��1��Ӥ������|��2��Ӥ������|��3��Ӥ������|��4��Ӥ������|��5��Ӥ������"
    For i = 0 To UBound(Split(strTmp, "|"))
        cboBabyFee.AddItem Split(strTmp, "|")(i)
    Next
    cboBabyFee.ListIndex = 0
    
    
    'ҽ������ֻ������סԺ�����ͷ����ڼ�
    If mlngInsure > 0 Then
        If mbln������ʽ��� Then    '���˺�:25435
            dtpBegin.Enabled = False
            mblnDBegin = dtpBegin.Enabled
            lstTime.Enabled = False
            lstUnit.Enabled = True
            lstItem.Enabled = True
            lstClass.Enabled = True
            dtpBegin.Enabled = True
            cboBabyFee.Enabled = False
        Else
            dtpBegin.Enabled = False
            cboBabyFee.Enabled = gclsInsure.GetCapability(support����_����Ӥ��������, mlngPatient, mlngInsure)
            lstUnit.Enabled = gclsInsure.GetCapability(support����_ָ������, mlngPatient, mlngInsure)
            lstItem.Enabled = gclsInsure.GetCapability(support����_ָ��������Ŀ, mlngPatient, mlngInsure)
            lstClass.Enabled = gclsInsure.GetCapability(support����_ָ����������, mlngPatient, mlngInsure)
            lstTime.Enabled = gclsInsure.GetCapability(support����_ָ��סԺ����, mlngPatient, mlngInsure)
            dtpEnd.Enabled = gclsInsure.GetCapability(support����_ָ�����ڷ�Χ, mlngPatient, mlngInsure)
        End If
        mblnDBegin = dtpBegin.Enabled
        mblnDEnd = dtpEnd.Enabled
    Else
        mblnDBegin = True
        mblnDEnd = True
    End If
    cmdSyn.Enabled = lstTime.Enabled
    If mbytFunc = 0 Then
        chkKind(0).Value = IIf(mbytKind = 0 Or mbytKind = 2, 1, 0)
        chkKind(1).Value = IIf(mbytKind = 1 Or mbytKind = 2, 1, 0)
    Else
        chkKind(1).Value = 0: chkKind(1).Visible = False
        chkKind(0).Value = 1: chkKind(0).Visible = True
     End If
End Sub

Private Sub lstClass_Click()
    Dim i As Long
    If mblnNotClick Then Exit Sub
    If lstClass.Selected(0) Then
        For i = 1 To lstClass.ListCount - 1
            lstClass.Selected(i) = True
        Next
    Else
        If lstClass.SelCount = 0 Then
            lstClass.Selected(lstClass.ListIndex) = True
        End If
    End If
End Sub

Private Sub lstItem_ItemCheck(Item As Integer)
    If lstItem.SelCount < 1 Then
        MsgBox "����Ҫѡ��һ��������Ŀ��", vbInformation, gstrSysName
        lstItem.Selected(Item) = True
    End If
End Sub

Private Sub lstUnit_ItemCheck(Item As Integer)
    If lstUnit.SelCount < 1 Then
        MsgBox "����Ҫѡ��һ�����ÿ��ң�", vbInformation, gstrSysName
        lstUnit.Selected(Item) = True
    End If
End Sub

Private Sub lstTime_ItemCheck(Item As Integer)
    Dim i As Long
    If mblnNotClick Then Exit Sub
    If lstTime.SelCount < 1 Then
        MsgBox "����Ҫѡ��һ��סԺ��", vbInformation, gstrSysName
        lstTime.Selected(Item) = True
    ElseIf mlngInsure > 0 Then
        'ҽ������ֻ��ѡ��һ��סԺ�ķ���
        For i = 0 To lstTime.ListCount - 1
            If i <> Item Then lstTime.Selected(i) = False
        Next
    End If
    
    If InStr(1, "," & mstrUnAuditTime & ",", "," & lstTime.ItemData(Item) & ",") > 0 Then
        MsgBox "��" & lstTime.ItemData(Item) & "��סԺ����δ��ˣ���û��Ȩ�޶Դ˽��ʣ�", vbInformation, gstrSysName
        lstTime.Selected(Item) = False
    End If
End Sub

Private Sub lvwChargeType_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        Dim objItem As ListItem
        If mblnNotClick Then Exit Sub
        
        With lvwChargeType
            If Item.Key = "ALL" Then
                    For Each objItem In .ListItems
                            If objItem.Key <> "ALL" Then
                                objItem.Checked = Item.Checked
                            End If
                    Next
            Else
                If .ListItems("ALL").Checked = True Then
                    Item.Checked = True
                End If
            End If
        End With
End Sub
