VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppChkRpt 
   Caption         =   "��������"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12840
   Icon            =   "frmAppChkRpt.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   12840
   StartUpPosition =   2  '��Ļ����
   Tag             =   "17500"
   Begin VB.CommandButton cmdSQL 
      Caption         =   "�������SQL"
      Height          =   350
      Left            =   8400
      TabIndex        =   11
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "����(&M)"
      Height          =   350
      Left            =   9960
      TabIndex        =   6
      Top             =   7320
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   11760
      TabIndex        =   5
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "�����Excel"
      Height          =   350
      Left            =   6840
      TabIndex        =   3
      Top             =   7440
      Width           =   1335
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   0
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   30
      Width           =   2205
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   1
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   1500
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   2
      Left            =   12240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfResult 
      Height          =   6255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   10695
      _cx             =   18865
      _cy             =   11033
      Appearance      =   3
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Editable        =   2
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
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "ϵͳ"
      Height          =   180
      Index           =   0
      Left            =   5760
      TabIndex        =   10
      Top             =   75
      Width           =   360
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   8520
      TabIndex        =   9
      Top             =   75
      Width           =   360
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "���س̶�"
      Height          =   180
      Index           =   2
      Left            =   11400
      TabIndex        =   8
      Top             =   75
      Width           =   720
   End
   Begin VB.Label lblRsFilter 
      Caption         =   "Label1"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   7200
      Width           =   5535
   End
End
Attribute VB_Name = "frmAppChkRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_COL = ",300,4;���,500,4;ϵͳ,2000,1;����,1500,1;������,2450,1;��������,6300,1;����˵��,3000,1;���س̶�,930,4;����SQL,0,4;Լ���ֶ�,0,4"
Private Const MSTR_ProCOL = "���,800,4;��������,2450,1;��������,6300,1"
Private mrsProData As New ADODB.Recordset
Private mrsDataFromFile As New ADODB.Recordset
Private mstrSysModul As String
Private Enum enuResult
    Col_ѡ�� = 0
    Col_���
    Col_ϵͳ
    Col_����
    Col_������
    Col_��������
    Col_����˵��
    Col_���س̶�
    Col_����SQL
    Col_Լ���ֶ�
End Enum

Private Enum enuPro
    Procol_��� = 0
    Procol_�������� = 1
    Procol_�������� = 2
End Enum

Private mblnFirst As Boolean
Private mstrPath As String
Private mbytType As Byte   '1-��������������-���̼����
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long '�༶Ŀ¼�����ڣ�Ҳ�ɴ���ָ��Ŀ¼

Private Sub cboFilter_Click(Index As Integer)
    Dim strFilter As String
    
    If mblnFirst = False Then Exit Sub
    
    If cboFilter(0).Text = "����ϵͳ" Then
        strFilter = ""
    Else
        strFilter = "ϵͳ����='" & cboFilter(0).Text & "'"
    End If
    
    If cboFilter(1).Text = "��������" Then
        strFilter = strFilter
    Else
        strFilter = IIf(strFilter = "", "���='" & cboFilter(1).Text & "'", strFilter & " and ���='" & cboFilter(1).Text & "'")
    End If
    
    If cboFilter(2).Text = "���г̶�" Then
        strFilter = strFilter
    Else
        strFilter = IIf(strFilter = "", "���س̶�='" & cboFilter(2).Text & "'", strFilter & " and ���س̶�='" & cboFilter(2).Text & "'")
    End If
    
    Call AddvsfData(strFilter)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSQL_Click()
    Dim strText As String
    Dim strTableName As String
    Dim strFild As String
    Dim varTemp As Variant
    Dim i As Long
    
    If cmdSQL.Visible = False Then Exit Sub
    Clipboard.Clear
    strText = vsfResult.Cell(flexcpText, vsfResult.Row, Col_Լ���ֶ�)
    strTableName = vsfResult.Cell(flexcpText, vsfResult.Row, Col_������)
    strTableName = Mid(strTableName, 1, InStr(strTableName, "_") - 1)
    varTemp = Split(strText, ",")
    For i = 0 To UBound(varTemp)
        strFild = IIf(strFild = "", "", strFild & " And ") & "a." & varTemp(i) & "=b." & varTemp(i)
    Next
    strText = "Delete " & strTableName & " Where Rowid In (Select a.Rowid From " & strTableName & " a,(Select " & strText & ", Max(Rowid) Rid From " & _
           strTableName & " Group By " & strText & ") b Where " & strFild & " And a.Rowid <> b.Rid)"
    Clipboard.SetText strText
End Sub

Private Sub Form_Load()
    
    If mbytType = 1 Then
        Me.Caption = "��������"
        mblnFirst = False
        Call InitTable(vsfResult, MSTR_COL)
        Call InivsfData
        mblnFirst = True
        Me.Tag = 17400
        cmdClose.Caption = "ȡ��(&C)"
        Call vsfResult_Click
    Else
        Me.Caption = "���̼����"
        cmdClose.Caption = "�˳�(&E)"
        cmdModify.Caption = "�鿴����"
        Call InitTable(vsfResult, MSTR_ProCOL)
        With vsfResult
            .Rows = .Rows - 1
            .rowHeight(0) = 500
            mrsProData.Sort = "��������"
            mrsProData.MoveFirst
            Do While Not mrsProData.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Procol_���) = .Rows - 1
                .TextMatrix(.Rows - 1, Procol_��������) = mrsProData!��������
                .TextMatrix(.Rows - 1, Procol_��������) = mrsProData!��������
                .rowHeight(.Rows - 1) = 500
                mrsProData.MoveNext
            Loop
            If .Rows > 1 Then
                .Row = 1
                Call .ShowCell(1, 1)
            End If
        End With
        lblRsFilter.Caption = "�����:��" & mrsProData.RecordCount & "����������"
    End If
End Sub

Public Function ShowMe(ByVal bytType As Byte, ByVal rsProData As ADODB.Recordset, Optional ByVal strPath As String, Optional ByVal rsDataFromFile As ADODB.Recordset) As Boolean
    'bytType,1-�ű�����޸���0-�洢���̼��
    
    mbytType = bytType
    Set mrsProData = rsProData
    If bytType = 1 Then
        Set mrsDataFromFile = rsDataFromFile
        mstrPath = strPath & "\Log\��־����\zlObjCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".Log"
    End If
    SetVisible
    Me.Show 1
End Function

Private Sub SetVisible()
    '���ÿؼ��Ŀɼ���
    Dim i As Long
    
    For i = 0 To lblFilter.UBound
        lblFilter(i).Visible = IIf(mbytType = 1, True, False)
    Next
    For i = 0 To lblFilter.UBound
        cboFilter(i).Visible = IIf(mbytType = 1, True, False)
    Next
    cmdSQL.Visible = False
End Sub

Private Sub InivsfData()
'���ܣ���������״�������ʾ
    Dim i As Long
    Dim strSys As String
    Dim strType2 As String
    Dim strSer As String
    
    With vsfResult
        strSys = "����ϵͳ"
        strType2 = "��������"
        strSer = "���г̶�"
        cboFilter(0).AddItem "����ϵͳ"
        cboFilter(1).AddItem "��������"
        cboFilter(2).AddItem "���г̶�"
        cboFilter(2).AddItem "����"
        cboFilter(2).AddItem "����"
        cboFilter(2).AddItem "��΢"
        .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
        .Rows = .Rows - 1
        Call AddvsfData
        
        For i = 1 To .Rows - 1
            If InStr("|" & strSys & "|", "|" & .TextMatrix(i, Col_ϵͳ) & "|") = 0 Then
                strSys = strSys & "|" & .TextMatrix(i, Col_ϵͳ)
                cboFilter(0).AddItem .TextMatrix(i, Col_ϵͳ)
            End If
            
            If InStr("|" & strType2 & "|", "|" & .TextMatrix(i, Col_����) & "|") = 0 Then
                strType2 = strType2 & "|" & .TextMatrix(i, Col_����)
                cboFilter(1).AddItem .TextMatrix(i, Col_����)
            End If
        Next
    End With
    
    cboFilter(0).ListIndex = 0
    cboFilter(1).ListIndex = 0
    cboFilter(2).ListIndex = 0
End Sub

Private Sub AddvsfData(Optional ByVal strFilter As String)
'���ܣ����������󵽱����
    Dim i As Long
    
    With vsfResult
        .Rows = 1
        .Redraw = flexRDNone
        .ColHidden(Col_����SQL) = True
        mrsProData.Filter = strFilter
        .Rows = mrsProData.RecordCount + 1
        i = 0
        Do While Not mrsProData.EOF
            i = i + 1
            .TextMatrix(i, Col_���) = i
            .TextMatrix(i, Col_ϵͳ) = mrsProData!ϵͳ����
            .TextMatrix(i, Col_����) = mrsProData!���
            .TextMatrix(i, Col_������) = mrsProData!������
            .TextMatrix(i, Col_��������) = mrsProData!��������
            .TextMatrix(i, Col_����˵��) = mrsProData!����˵��
            .TextMatrix(i, Col_���س̶�) = mrsProData!���س̶�
            .TextMatrix(i, Col_����SQL) = mrsProData!����SQL
            .TextMatrix(i, Col_Լ���ֶ�) = "" & mrsProData!Լ���ֶ�
            If .TextMatrix(i, Col_���س̶�) = "��΢" Then
                .Cell(flexcpBackColor, i, Col_���س̶�) = RGB(238, 230, 133)
            ElseIf .TextMatrix(i, Col_���س̶�) = "����" Then
                .Cell(flexcpBackColor, i, Col_���س̶�) = RGB(238, 201, 0)
            ElseIf .TextMatrix(i, Col_���س̶�) = "����" Then
                .Cell(flexcpBackColor, i, Col_���س̶�) = RGB(238, 154, 0)
            End If
            If InStr(.TextMatrix(i, Col_����˵��), "�˹�") > 0 Then
                .TextMatrix(i, Col_ѡ��) = ""
            Else
                .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked
            End If
            mrsProData.MoveNext
        Loop
        .Cell(flexcpAlignment, 0, 0, .Rows - 1) = 4
        .Redraw = flexRDDirect
        If .Rows > 1 Then
            .Row = 1
            Call .ShowCell(1, 1)
        End If
    End With
    lblRsFilter.Caption = "���������" & mrsProData.RecordCount & "�����⡣"
End Sub

Private Sub cmdModify_Click()
'���ܣ�������ѡ�Ķ�������
    Dim i As Long
    Dim j As Long
    Dim lngLine As Long
    Dim varTemp As Variant
    Dim strErr As String
    Dim strTemp As String
    Dim strSQL As String
    Dim blnModify As Boolean
    Dim blnFalse As Boolean
    Dim cnChoose As ADODB.Connection
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
            
    If mbytType = 1 Then
        With vsfResult
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, 0) = flexChecked Then
                    Call ShowFlash("���ڽ��ж�������ݵ����������Ժ�")
                    If .TextMatrix(i, Col_ϵͳ) = "������������" Then
                        If gcnTools Is Nothing Then
                            Set gcnTools = GetConnection("ZLTOOLS")
                        End If
                        Set cnChoose = gcnTools
                    Else
                        Set cnChoose = gcnOracle
                    End If
                    blnFalse = True
                    varTemp = Split(UCase(.TextMatrix(i, Col_����SQL)), "{JM|SQL�ָ���}" & vbNewLine)
                    For j = 0 To UBound(varTemp)
                        strSQL = varTemp(j)
                        If strSQL <> "" Then
                            On Error Resume Next
                            cnChoose.Execute strSQL
                            If err.Number <> 0 Then
                                If strSQL Like "INSERT INTO ZLPARAMETERS*" Then
                                    strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                                    Set rsTemp = SetSelectRecordset(strSQL, strTemp, Split(strTemp, ","), "ZLPARAMETERS")
                                    If InStr(rsTemp!ģ��, "NULL") = 0 And InStr(rsTemp!ϵͳ, "NULL") = 0 Then
                                        If InStr(mstrSysModul, rsTemp!ϵͳ & "&" & rsTemp!ģ��) = 0 Then
                                            mrsDataFromFile.Filter = "���='����'"
                                            Set rsData = CopyNewRec(mrsDataFromFile)
                                            mstrSysModul = mstrSysModul & "|" & rsTemp!ϵͳ & "&" & rsTemp!ģ��
                                            strSQL = "Update Zlparameters Set ������ = -1 * ������ Where ϵͳ =" & rsTemp!ϵͳ & " And ģ�� = " & rsTemp!ģ��
                                            cnChoose.Execute strSQL
                                            rsData.Filter = "���='����' and ����=" & rsTemp!ģ�� & " and ϵͳ���=" & rsTemp!ϵͳ
                                            Do While Not rsData.EOF
                                                mrsDataFromFile.Filter = "���='����' and ����=" & rsTemp!ģ�� & " and ϵͳ���=" & rsTemp!ϵͳ & " and ������='" & rsData!������ & "'"
                                                If mrsDataFromFile.RecordCount > 0 Then
                                                    strSQL = "Update Zlparameters Set ������ = " & rsTemp!������ & " Where ϵͳ =" & rsTemp!ϵͳ & " And ģ�� = " & rsTemp!ģ�� & " and ������='" & rsData!������ & "'"
                                                    cnChoose.Execute strSQL
                                                End If
                                                rsData.MoveNext
                                            Loop
                                            cnChoose.Execute varTemp(j)
    '                                        strSQL = "Update Zlparameters Set ������ = -1 * ������ Where ϵͳ =" & rsTemp!ϵͳ & " And ģ�� = " & rsTemp!ģ��
    '                                        cnChoose.Execute strSQL
                                        End If
                                    Else
                                        blnFalse = False
                                        strErr = IIf(strErr = "", "����ʧ�ܵ�SQL��" & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf, strErr & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf)
                                    End If
                                Else
                                    'ɾ��ʱ���������ʾɾ���ɹ�
                                    If UCase(err.Description) Like "ORA-01418*" Then
                                    Else
                                        blnFalse = False
                                        strErr = IIf(strErr = "", "����ʧ�ܵ�SQL��" & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf, strErr & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf)
                                    End If
                                End If
                            Else
                                'ɾ��������/����ʱ����ɾ����Ӧ�Ĺ���ͬ���
                                If .TextMatrix(i, Col_����) = "ZLTOOL����" Then
                                    If strSQL Like "DROP TABLE*" Or strSQL Like "DROP PROCEDURE*" Or strSQL Like "DROP FUNCTION" Then
                                        gstrSQL = "Select 'Drop Public SYNONYM ' || Synonym_Name ִ��SQL" & vbNewLine & _
                                                    "From All_Synonyms a" & vbNewLine & _
                                                    "Where Table_Owner=[1] And Owner = 'PUBLIC' And Not Exists" & vbNewLine & _
                                                    " (Select 1 From All_Objects b Where a.Table_Name = b.Object_Name And a.Table_Owner = b.Owner) And" & vbNewLine & _
                                                    "      a.Synonym_Name =[2]"
                                        Set rsTemp = gclsBase.OpenSQLRecord(cnChoose, gstrSQL, Me.Caption, UCase(.TextMatrix(i, Col_ϵͳ)), UCase(.TextMatrix(i, Col_������)))
                                        Do While Not rsTemp.EOF
                                            cnChoose.Execute rsTemp!ִ��SQL
                                            rsTemp.MoveNext
                                        Loop
                                    End If
                                End If
                            End If
                        End If
                    Next
                    blnModify = True
                    If blnFalse Then
                        .Cell(flexcpData, i, 0) = 1
                    Else
                        .Cell(flexcpData, i, 0) = 0
                    End If
                End If
            Next
            If blnModify = False Then
                MsgBox "δ��ѡ���Զ����������ݣ�"
                Exit Sub
            End If
            Call ShowFlash("")
            If strErr <> "" Then
                On Error Resume Next
                Call WriteErrorLog(strErr)
                If err.Number = 0 Then
                    MsgBox "������ɣ��в�������δ�ɹ��������������" & mstrPath
                Else
                    MsgBox "������ɣ�������־��¼ʧ�ܣ������Ǹ���־�ļ�(" & mstrPath & ")�Ѵ򿪣����飡"
                End If
                err.Clear: On Error GoTo 0
            Else
                MsgBox "������ɣ�"
            End If
        End With
        Call AfterModify
    Else
        With vsfResult
            If .Rows = 1 Then Exit Sub
            mrsProData.Filter = "��������='" & .TextMatrix(.Row, Procol_��������) & "'"
            strTemp = "Create or Replace " & mrsProData!ԭʼSQL
            If InStr(.TextMatrix(.Row, Procol_��������), "Commit") > 0 Then
                lngLine = GetFirstLine(mrsProData!ԭʼSQL, "COMMIT")
            ElseIf InStr(.TextMatrix(.Row, Procol_��������), "�󶨱���") > 0 Then
                lngLine = GetFirstLine(mrsProData!ԭʼSQL, "EXECUTE IMMEDIATE")
            End If
            Call frmProcEditCommon.ShowMe(0, .TextMatrix(.Row, Procol_��������), strTemp, "", "", "", 1, lngLine)
        End With
    End If
End Sub

Private Function GetFirstLine(ByVal strSQL As String, ByVal strKey As String) As Long
'��ȡָ���ؼ����״γ��ֵ�����
    Dim i As Long
    Dim varTemp As Variant
    
    varTemp = Split(UCase(strSQL), vbLf)
    For i = 0 To UBound(varTemp)
        If InStr(varTemp(i), strKey) > 0 Then
            GetFirstLine = i + 1
            Exit Function
        End If
    Next
End Function

Private Sub AfterModify()
'������ɺ�����ˢ�½�������
    Dim i As Long
    Dim strFilter As String
    Dim lngSelRow As Long
    
    lblRsFilter.Caption = "��������ˢ�½���......"
    With vsfResult
        lngSelRow = .Row
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                strFilter = "��������='" & .TextMatrix(i, Col_��������) & "' and ������='" & .TextMatrix(i, Col_������) & "' and ���='" & .TextMatrix(i, Col_����) & "'"
                Call RecDelete(mrsProData, strFilter)
            End If
        Next
        Call cboFilter_Click(0)
        If .Rows > 1 Then
            If .Rows > lngSelRow Then
                .Row = lngSelRow
                Call .ShowCell(lngSelRow, 1)
            Else
                .Row = .Rows - 1
                Call .ShowCell(.Rows - 1, 1)
            End If
        End If
    End With
    Call vsfResult_AfterEdit(1, 0)
End Sub

Private Sub WriteErrorLog(ByVal strErr As String)
    Dim objFile As Object
    Dim objStream As TextStream
        
    Call MakeSureDirectoryPathExists(mstrPath)
    Set objFile = CreateObject("Scripting.FileSystemObject")
    If objFile.FileExists(mstrPath) = False Then objFile.CreateTextFile mstrPath
    Set objStream = objFile.OpenTextFile(mstrPath)

    Open mstrPath For Append Shared As #1
    Print #1, strErr
    Close #1
End Sub

Private Sub cmdOut_Click()
    
    Call OutExcel
End Sub

Private Sub OutExcel()
'���ܣ���vsf����������Excel��
    Dim spShell, spFolder, spFolderItem, spPath As String
    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = 0

    On Error GoTo errH
    If IsInstallExcel Then
        With vsfResult
            If .Rows < 2 Then
                MsgBox "�����û�����ݣ��޷�������ݣ����飡"
                Exit Sub
            Else
                Set spShell = CreateObject("Shell.Application")
                Set spFolder = spShell.BrowseForFolder(WINDOW_HANDLE, "ѡ��Ŀ¼:", NO_OPTIONS)
                If spFolder Is Nothing Then
                    Exit Sub
                Else
                    Set spFolderItem = spFolder.Self
                    spPath = spFolderItem.Path
                    .SaveGrid Replace(spPath & "\zlObjectCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".xls", "\\", "\"), flexFileExcel, True
                    .BackColorSel = &H8000000D
                     MsgBox "����ɹ���������ѱ�����" & Replace(spPath & "\zlObjectCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".xls", "\\", "\")
                     Exit Sub
                End If
            End If
        End With
    End If
    Exit Sub
errH:
    MsgBox "��ѡ·���ĸ��ļ����ڴ�״̬����ѡ·������"
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If ScaleHeight < 2000 Then Exit Sub
        
    With vsfResult
        .Left = ScaleLeft
        .Width = ScaleWidth
        If mbytType = 1 Then
            .Top = ScaleTop + 600
            .Height = ScaleHeight - cmdModify.Height - 900
            .ColWidth(Col_����) = 1500 + 0.05 * (Me.Width - Me.Tag)
            .ColWidth(Col_������) = 2450 + 0.25 * (Me.Width - Me.Tag)
            .ColWidth(Col_��������) = 6300 + 0.3 * (Me.Width - Me.Tag)
            .ColWidth(Col_����˵��) = 3000 + 0.4 * (Me.Width - Me.Tag)
        Else
            .Top = ScaleTop
            .Height = ScaleHeight - cmdModify.Height - 300
            .ColWidth(Procol_��������) = (.Width - 800) * 0.3
            .ColWidth(Procol_��������) = (.Width - 800) * 0.7
        End If
    End With
    cmdClose.Top = vsfResult.Top + vsfResult.Height + 150
    cmdClose.Left = ScaleWidth - cmdClose.Width - 300
    
    cmdModify.Top = cmdClose.Top
    cmdModify.Left = cmdClose.Left - cmdModify.Width - 500
    
    cmdSQL.Top = cmdClose.Top
    cmdSQL.Left = cmdModify.Left - cmdSQL.Width - 500
    
    cmdOut.Top = cmdClose.Top
    cmdOut.Left = cmdSQL.Left - cmdOut.Width - 500
    
    lblRsFilter.Top = cmdOut.Top + 150
    lblRsFilter.Left = 300
    
    cboFilter(2).Top = 200
    cboFilter(2).Left = ScaleWidth - cboFilter(2).Width - 300
    lblFilter(2).Top = 250
    lblFilter(2).Left = cboFilter(2).Left - lblFilter(2).Width - 150
    
    cboFilter(1).Top = 200
    cboFilter(1).Left = lblFilter(2).Left - cboFilter(1).Width - 300
    lblFilter(1).Top = 250
    lblFilter(1).Left = cboFilter(1).Left - lblFilter(1).Width - 150
    
    cboFilter(0).Top = 200
    cboFilter(0).Left = lblFilter(1).Left - cboFilter(0).Width - 300
    lblFilter(0).Top = 250
    lblFilter(0).Left = cboFilter(0).Left - lblFilter(0).Width - 150

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytType = 1 Then
        Call ReleaseMe
    End If
    Set mrsProData = Nothing
    Set mrsDataFromFile = Nothing
End Sub

Private Sub vsfResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    If mbytType = 1 Then
        With vsfResult
            If Col = Col_ѡ�� Then
                If Row = 0 Then
                    If .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked Then
                        .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked
                        For i = 1 To .Rows - 1
                            If .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked Then
                                .Cell(flexcpChecked, i, Col_ѡ��) = flexChecked
                            End If
                        Next
                    Else
                        .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
                        For i = 1 To .Rows - 1
                            If .Cell(flexcpChecked, i, Col_ѡ��) = flexChecked Then
                                .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked
                            End If
                        Next
                    End If
                Else
                    If .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked Then
                        .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
                    End If
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked Then
                            Exit For
                        Else
                            If i = .Rows - 1 Then
                                .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked
                            End If
                        End If
                    Next
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfResult_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If mbytType = 1 Then
        If Col <> 0 Then Cancel = True
    End If
End Sub

Private Sub vsfResult_Click()
    Dim strTemp As String
    
    If mbytType = 0 Then Exit Sub
    '�������޸�ʱ�Ž��и������SQL�Ŀɼ��Ե���
    strTemp = vsfResult.Cell(flexcpText, vsfResult.Row, Col_������)
    If (strTemp Like "*_PK" Or strTemp Like "*_UQ_*") And Mid(vsfResult.Cell(flexcpText, vsfResult.Row, Col_��������), 1, 6) <> "���ݿ��д���" Then
        cmdSQL.Visible = True
    Else
        cmdSQL.Visible = False
    End If
End Sub

Private Sub vsfResult_DblClick()
    If mbytType = 1 Then Exit Sub
    '���̼�����鿴���̲�ִ�иò���
    Call cmdModify_Click
End Sub

Private Sub vsfResult_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strTip As String
    
    If mbytType = 1 Then
        With vsfResult
            If .MouseRow <> -1 And .MouseRow <> 0 And .MouseCol = Col_����˵�� Then
                If .TextMatrix(.MouseRow, Col_����SQL) <> "" Then
                    strTip = "����SQL:" & vbNewLine & Replace(.TextMatrix(.MouseRow, Col_����SQL), "{JM|SQL�ָ���}", "")
                    Call ShowTipInfo(.hwnd, strTip, True)
                Else
                    Call ShowTipInfo(.hwnd, "")
                End If
            Else
                Call ShowTipInfo(.hwnd, "")
            End If
        End With
    End If
End Sub


