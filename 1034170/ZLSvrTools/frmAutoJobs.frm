VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAutoJobs 
   BackColor       =   &H80000005&
   Caption         =   "��̨��ҵ����"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAutoJobs.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   6465
   WindowState     =   2  'Maximized
   Begin VB.Frame fraComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   885
      TabIndex        =   5
      Top             =   4095
      Width           =   4920
      Begin VB.Label lblPara 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   825
         Width           =   540
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵����"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   195
         Width           =   540
      End
      Begin VB.Label lbl˵�� 
         BackStyle       =   0  'Transparent
         Height          =   525
         Left            =   600
         TabIndex        =   6
         Top             =   210
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����ִ��(&T)��"
      Height          =   350
      Left            =   885
      TabIndex        =   4
      Top             =   5685
      Width           =   1395
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "��������(&T)��"
      Height          =   350
      Left            =   2280
      TabIndex        =   3
      Top             =   5685
      Width           =   1395
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   3945
      TabIndex        =   2
      Top             =   5685
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4845
      TabIndex        =   1
      Top             =   5685
      Width           =   945
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfJobs 
      Height          =   2175
      Left            =   885
      TabIndex        =   0
      Top             =   1455
      Width           =   5415
      _cx             =   9551
      _cy             =   3836
      Appearance      =   1
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
      BackColorBkg    =   16777215
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   285
      Picture         =   "frmAutoJobs.frx":04F9
      Top             =   660
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��̨��ҵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1380
      TabIndex        =   10
      Top             =   6255
      Width           =   4890
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   885
      TabIndex        =   9
      Top             =   5745
      Width           =   4890
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAutoJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MSTR_COL = "ϵͳ,2000,1;���,1200,1;���,450,1;˵��,0,1;����,0,1;����,2000,1;���ù���,3000,1;��ҵ��,0,1;�Զ�ִ��,800,4;״̬,600,4;��ʼִ��ʱ��,1900,1;���ʱ��,900,1;ϵͳ���,0,1;������,0,1"
Private Enum vsfCol
    Col_ϵͳ = 0
    Col_��� = 1
    Col_��� = 2
    Col_˵�� = 3
    Col_���� = 4
    Col_���� = 5
    Col_���ù��� = 6
    Col_��ҵ�� = 7
    Col_�Զ�ִ�� = 8
    Col_״̬ = 9
    Col_��ʼִ��ʱ�� = 10
    Col_���ʱ�� = 11
    Col_ϵͳ��� = 12
    Col_������ = 13
End Enum
Private mlngMaxJobs As Long '�����ݿ����������ҵ��
Private mstrPro As String '��ǰ�����е��ù����ַ���

Private Sub cmdAdd_Click()
    Dim lngSelectRow As Long

    lngSelectRow = vsfJobs.Row
    Call frmAutoJobset.Add(mstrPro)
    Call LoadData(lngSelectRow)
End Sub

Private Sub cmdDel_Click()
    Dim lngSystem As Long
    Dim cnTools As ADODB.Connection
    Dim strTemp As String

    With vsfJobs
        strTemp = "��ȷ��Ҫɾ����" & .TextMatrix(.Row, Col_����) & "����̨��ҵ��"
        If MsgBox(strTemp, vbExclamation + vbDefaultButton1 + vbYesNo) = vbNo Then
            Exit Sub
        End If
        lngSystem = .TextMatrix(.Row, Col_ϵͳ���)
        strTemp = UCase(.TextMatrix(.Row, Col_���ù���))
        If Val(.TextMatrix(.Row, Col_��ҵ��)) <> 0 Then
            If lngSystem = 0 Then
                Set cnTools = GetConnection("ZLTOOLS")
                If cnTools Is Nothing Then Exit Sub
            Else
                Set cnTools = gcnOracle
            End If
            gstrSQL = "zl_JobRemove(" & IIf(lngSystem = 0, "Null", lngSystem) & ",3," & .TextMatrix(.Row, Col_���) & ")"
            err = 0
            cnTools.Execute gstrSQL, , adCmdStoredProc
            If err <> 0 Then
                MsgBox "��ҵɾ��ʧ�ܣ�", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        gstrSQL = "delete zlAutoJobs" & _
                " where Nvl(ϵͳ,0)=" & lngSystem & " and ����=3" & _
                " and ���=" & .TextMatrix(.Row, Col_���)
        err = 0
        On Error Resume Next
        gcnOracle.Execute gstrSQL
        If err <> 0 Then
            MsgBox "��ҵɾ��ʧ�ܣ�", vbExclamation, gstrSysName
            Exit Sub
        Else
            mstrPro = Replace(mstrPro, strTemp & ",", "")
        End If
    End With
    Call LoadData(vsfJobs.Row)
End Sub

Private Sub cmdSet_Click()
    Dim strParas As String
    Dim aryPara() As String
    Dim intCount As Integer
    
    If Val(vsfJobs.TextMatrix(vsfJobs.Row, Col_���)) = 0 Then Exit Sub
    Call frmAutoJobset.RunSet(vsfJobs)
    Call LoadData(vsfJobs.Row)
End Sub

Private Sub cmdTest_Click()
    Dim lngSystem As Long
    Dim cnTools As ADODB.Connection
    Dim lngType As Long
    
    With vsfJobs
        If gblnDBA Then
            'DBA�û�����Ҫ�ж�
        ElseIf .TextMatrix(.Row, Col_ϵͳ) = "������������" Then
            '��Ϊ������������Ϊ�գ�������Ҫ���ж��Ƿ�Ϊ������
        ElseIf .TextMatrix(.Row, Col_������) <> gstrUserName Then
            MsgBox "��ǰ�û����Ǹ�ϵͳ�������ߣ��޷����иò�����"
            Exit Sub
        End If
        lngSystem = .TextMatrix(.Row, Col_ϵͳ���)
        If Val(.TextMatrix(.Row, Col_��ҵ��)) <> 0 Then
            If lngSystem = 0 Then
                Set cnTools = GetConnection("ZLTOOLS")
                If cnTools Is Nothing Then Exit Sub
            Else
                Set cnTools = gcnOracle
            End If
            If .TextMatrix(.Row, Col_���) = "ϵͳ�趨" Then
                lngType = 1
            ElseIf .TextMatrix(.Row, Col_���) = "����ת��" Then
                lngType = 2
            Else
                lngType = 3
            End If
            gstrSQL = "zl_JobRun(" & IIf(lngSystem = 0, "Null", lngSystem) & "," & lngType & "," & .TextMatrix(.Row, Col_���) & ")"
            err = 0
            On Error Resume Next
            cnTools.Execute gstrSQL, , adCmdStoredProc
            If err <> 0 Then
                MsgBox "���Թ��̷����������" & vbNewLine & err.Description, vbExclamation, gstrSysName
                Exit Sub
            End If
            MsgBox "����ִ����ɣ��������ҵ״̬��Ϊ����Ч����˵��ִ�гɹ���", vbInformation, gstrSysName
        End If
    End With
    Call LoadData(vsfJobs.Row)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strSame As String
    Dim i As Integer
    
    On Error GoTo errHandle
    'ת�벻���ڵ�����ת�Ƽ�¼��Ϊ��ҵ��¼
    gstrSQL = "INSERT INTO zlAutoJobs (ϵͳ,����,���,����,˵��,����,����,ִ��ʱ��,���ʱ��)" & _
            " SELECT ϵͳ,2,���,����,˵��,'zl'||floor(ϵͳ/100)||'_DataMoveOut'||���,�����ֶ�||','||ת������,to_date('2000-01-01 01:00:00','YYYY-MM-DD HH24:MI:SS'),30" & _
            " FROM zlDataMove" & _
            " WHERE (ϵͳ,���) not in( select ϵͳ,��� from zlAutoJobs where ����=2)"
    gcnOracle.Execute gstrSQL
    
    lblMain.Caption = "�����������ݿ��̨�Զ���ҵ�����ڶ����������е����ݼ���������޸ĵ�����" & _
        vbCrLf & vbCrLf & "����������ϵͳ�ȽϿ��е�ʱ��ִ�У��Լ��ٺ������������Դ��������֤ǰ̨����������ٶȡ�"
    
    gstrSQL = "select value" & _
            " from v$parameter" & _
            " where name='job_queue_processes'"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset
    mlngMaxJobs = 0
    If Not rsTemp.EOF Then
        mlngMaxJobs = rsTemp.Fields(0).value
        If mlngMaxJobs > 0 Then
            lbl����.Caption = "�������ݿ����job_queue_processes���ã�Ŀǰ��������" & mlngMaxJobs & "���Զ���ҵ"
        Else
            lbl����.Caption = "��ǰ���������Զ���ҵ�����б�Ҫ�����޸����ݿ����job_queue_processes"
        End If
    End If
    If mlngMaxJobs = 0 Then
        cmdTest.Enabled = False
        cmdSet.Enabled = False
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
    End If
    Call InitTable(vsfJobs, MSTR_COL)
    Call LoadData(1)
    mstrPro = ""
    With vsfJobs
        For i = 1 To .Rows - 1
            If InStr(mstrPro, UCase(.TextMatrix(i, Col_���ù���)) & ",") = 0 Then
                mstrPro = mstrPro & UCase(.TextMatrix(i, Col_���ù���)) & ","
            End If
        Next
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
 
End Sub

Private Sub LoadData(ByVal lngRow As Long)
'���ܣ����ؽ���ʱ������ʾ
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim lngColor As Long
    Dim strTemp As String
    Dim strPro As String
    Dim varTemp As Variant
                
    With rsTemp
        gstrSQL = "select Nvl(C.����,'������������') ϵͳ,decode(A.����,1,'ϵͳ�趨',2,'����ת��',3,'�û��Զ���') as ��� ,A.���,A.˵��," & _
                "A.����,A.���� ���ù���,A.��ҵ��,A.����,decode(A.��ҵ��,0,'��',null,'��','��') as �Զ�ִ��," & _
                "decode(B.BROKEN,null,'ȱʧ','Y','��Ч','��Ч') as ״̬,A.ִ��ʱ�� ��ʼִ��ʱ��,A.���ʱ��||Nvl(A.ʱ�䵥λ,'��') as ���ʱ��,Nvl(A.ϵͳ,0) ϵͳ���,C.������ " & _
                "From zlAutoJobs A," & IIf(gblnDBA, "dba_jobs", "user_jobs") & " B,zlsystems C " & _
                "where A.��ҵ��=B.JOB(+) and A.ϵͳ=C.���(+) " & IIf(gblnOwner, " And c.������=user", "") & " order by ϵͳ���,���,���"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset
    End With

    With vsfJobs
        .Rows = 1
        .rowHeight(0) = 300
        .MergeCells = flexMergeRestrictRows
        .MergeCol(Col_ϵͳ) = True
        .MergeCol(Col_���) = True
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, Col_ϵͳ) = rsTemp!ϵͳ & ""
            .TextMatrix(.Rows - 1, Col_���) = rsTemp!��� & ""
            .TextMatrix(.Rows - 1, Col_���) = rsTemp!��� & ""
            .TextMatrix(.Rows - 1, Col_˵��) = rsTemp!˵�� & ""
            .TextMatrix(.Rows - 1, Col_����) = rsTemp!���� & ""
            .TextMatrix(.Rows - 1, Col_����) = rsTemp!���� & ""
            .TextMatrix(.Rows - 1, Col_���ù���) = rsTemp!���ù��� & ""
            .TextMatrix(.Rows - 1, Col_��ҵ��) = rsTemp!��ҵ�� & ""
            .TextMatrix(.Rows - 1, Col_�Զ�ִ��) = rsTemp!�Զ�ִ�� & ""
            .TextMatrix(.Rows - 1, Col_״̬) = rsTemp!״̬ & ""
            .TextMatrix(.Rows - 1, Col_��ʼִ��ʱ��) = rsTemp!��ʼִ��ʱ�� & ""
            .TextMatrix(.Rows - 1, Col_���ʱ��) = rsTemp!���ʱ�� & ""
            .TextMatrix(.Rows - 1, Col_ϵͳ���) = rsTemp!ϵͳ��� & ""
            .TextMatrix(.Rows - 1, Col_������) = rsTemp!������ & ""
            rsTemp.MoveNext
        Loop
        For i = 1 To .Rows - 1
            .rowHeight(i) = 300
            strPro = UCase(.TextMatrix(i, Col_���ù���))
            If InStr(strTemp, strPro & ",") > 0 Then
                varTemp = Split(strTemp, ",")
                For j = 0 To UBound(varTemp)
                    If varTemp(j) = strPro Then
                        .Cell(flexcpBackColor, j + 1, 2, j + 1, .Cols - 1) = RGB(238, 230, 133 + lngColor * 10)
                    End If
                Next
                .Cell(flexcpBackColor, i, 2, i, .Cols - 1) = RGB(238, 230, 133 + lngColor * 10)
                lngColor = lngColor + 1
            End If
            strTemp = strTemp & strPro & ","
        Next
        If .Rows > 1 Then
            If lngRow > .Rows Then lngRow = .Rows
            .Row = lngRow
            Call .ShowCell(lngRow, 1)
            Call vsfJobs_RowColChange
        End If
    End With
End Sub

Private Sub Form_Resize()
    Dim sngBottom As Single
    
    On Error Resume Next
    
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With


    With vsfJobs
        .Top = imgMain.Top + 50
        .Left = imgMain.Left + imgMain.Width + 50
        .Width = ScaleWidth - .Left - 200
        sngBottom = ScaleHeight - lblMain.Height - 420 - cmdTest.Height - fraComment.Height - lbl����.Height
        .Height = IIf(sngBottom - .Top > 2500, sngBottom - .Top, 2500)
    End With
    
    With lblMain
        .Left = vsfJobs.Left
        .Width = vsfJobs.Width

        lbl����.Left = .Left
        lbl����.Width = .Width
    End With
    
    fraComment.Width = vsfJobs.Width
    fraComment.Left = vsfJobs.Left
    fraComment.Top = vsfJobs.Top + vsfJobs.Height
    lbl˵��.Width = fraComment.Width - lbl˵��.Left - 300

    cmdDel.Left = vsfJobs.Left + vsfJobs.Width - cmdDel.Width
    cmdAdd.Left = cmdDel.Left - cmdAdd.Width
    cmdTest.Top = fraComment.Top + fraComment.Height + 60
    cmdTest.Left = vsfJobs.Left
    cmdSet.Top = cmdTest.Top
    cmdSet.Left = cmdTest.Left + cmdTest.Width
    cmdAdd.Top = cmdTest.Top
    cmdDel.Top = cmdTest.Top

    lblMain.Top = cmdTest.Top + cmdTest.Height + 200
    lbl����.Top = lblMain.Top + lblMain.Height + 60
    
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow

    On Error GoTo errHandle
    objPrint.Title.Text = "��̨��ҵ"

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    Set objPrint.Body = vsfJobs
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub vsfJobs_RowColChange()

    With vsfJobs
        lbl˵��.Caption = .TextMatrix(.Row, Col_˵��)
        lblPara.Caption = "������" & .TextMatrix(.Row, Col_����)
        If .TextMatrix(.Row, Col_�Զ�ִ��) = "��" Then
            cmdTest.Enabled = True
        Else
            cmdTest.Enabled = False
        End If
        If .TextMatrix(.Row, Col_���) = "�û��Զ���" Then
            cmdDel.Enabled = True
        Else
            cmdDel.Enabled = False
        End If
    End With
End Sub


