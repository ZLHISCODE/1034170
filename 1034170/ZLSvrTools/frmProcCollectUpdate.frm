VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcCollectUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Ѽ�����"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmProcCollectUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   0
      Left            =   135
      ScaleHeight     =   4500
      ScaleWidth      =   10185
      TabIndex        =   6
      Top             =   1485
      Width           =   10185
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1140
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   105
         Width           =   1935
         _cx             =   3413
         _cy             =   2011
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "frmProcCollectUpdate.frx":6852
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   5
      Top             =   75
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��ʼ�Ѽ�(&S)"
      Height          =   350
      Left            =   7800
      TabIndex        =   4
      Top             =   6075
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   9240
      TabIndex        =   3
      Top             =   6090
      Width           =   1100
   End
   Begin VB.OptionButton opt 
      Caption         =   "��ǰ���ݿ�"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   1125
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.OptionButton opt 
      Caption         =   "�������ݿ�"
      Height          =   255
      Index           =   1
      Left            =   1650
      TabIndex        =   1
      Top             =   1125
      Width           =   1380
   End
   Begin VB.CommandButton cmdConnet 
      Caption         =   "��������(&L)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2970
      TabIndex        =   0
      Top             =   1065
      Width           =   1290
   End
   Begin MSComctlLib.ProgressBar pbr 
      Height          =   105
      Left            =   120
      TabIndex        =   8
      Top             =   6525
      Visible         =   0   'False
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3720
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "test"
      Height          =   180
      Left            =   135
      TabIndex        =   11
      Top             =   6300
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�Ѽ��Ǽǹ���/����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1110
      TabIndex        =   10
      Top             =   105
      Width           =   2820
   End
   Begin VB.Label Label2 
      Caption         =   "�����·�ѡ��ǰ�汾�Ľű������ļ����Ա�͵�ǰ�汾���ݿ��еĹ��̽��бȽϣ��ó��и��ĵĹ��̡�"
      Height          =   210
      Left            =   1140
      TabIndex        =   9
      Top             =   555
      Width           =   9180
   End
End
Attribute VB_Name = "frmProcCollectUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjMain As Object
Private mclsVsf As clsVsf
Private mblnOK As Boolean
Private WithEvents mfrmPageConfigure As frmProcConfigure
Attribute mfrmPageConfigure.VB_VarHelpID = -1

Private mcnOracle As ADODB.Connection

Public Function ShowMe(ByVal objMain As Object) As Boolean
    On Error GoTo errHand
    
    mblnOK = False
    
    Set mobjMain = objMain
    Me.Show 1, mobjMain
    
    ShowMe = mblnOK
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim intRow As Integer
    Dim intFlag As Integer
    Dim strSQL As String
    Dim objItem As Object
    Dim strUpPath As String
    Dim strFlag As String
    
    On Error GoTo errHand
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True)
            Call .ClearColumn
            Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[���]", False, False, False)
            Call .AppendColumn("�汾��", 0, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("ϵͳ����", 1800, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("��װ�ű�", 2700, flexAlignLeftCenter, flexDTString, , "", True)
            
            Call .InitializeEdit(True, False, True)
            Call .InitializeEditColumn(.ColIndex("��װ�ű�"), True, vbVsfEditCommand)

            .IndicatorMode = 2
            .IndicatorCol = .ColIndex("���")
            .ConstCol = .ColIndex("���")
                
            .AppendRows = True
        End With
'        lblState.ForeColor = &HFF&
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"

        With vsf(0)
            strSQL = "Select A.���,A.�汾��,A.���� as ϵͳ����,B.�ļ��� From zlSystems A,zlSysFiles B Where A.��� = B.ϵͳ And B.����=1"
            Set rs = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "")
            If rs.BOF = False Then
                For intRow = 0 To rs.RecordCount - 1
                    intFlag = intFlag + 1
                    If .Rows < intFlag + 1 Then .Rows = intFlag + 1
                    .TextMatrix(intRow + 1, .ColIndex("ϵͳ����")) = rs("ϵͳ����").value
                    .TextMatrix(intRow + 1, .ColIndex("��װ�ű�")) = rs("�ļ���").value
                    
                    strFlag = rs("�汾��").value
                    .TextMatrix(intRow + 1, .ColIndex("�汾��")) = strFlag
                    strFlag = Split(strFlag, ".")(0) & "." & Split(strFlag, ".")(1) & ".0"
                    
                    'ȱʡ�����ű�
                    strUpPath = Split(rs("�ļ���").value, "Ӧ�ýű�")(0) & "�����ű�\" & strFlag & "\zlUpgrade.ini"
                                                            
                    .RowData(intRow + 1) = rs("���").value
                    rs.MoveNext
                Next
            End If
        End With
        mclsVsf.UpdateSerial
        mclsVsf.AppendRows = True
    End Select
    ExecuteCommand = True
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnet_Click()
    If mfrmPageConfigure Is Nothing Then
        Set mfrmPageConfigure = New frmProcConfigure
    End If
    Call mfrmPageConfigure.ShowConfigure(Me)
End Sub

Private Sub cmdOK_Click()
    '1.����������ʱ�ļ���
    Dim strTmp1 As String
    Dim strProcedure As String
    Dim strTmpReports As String
    Dim strFlag As String
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim lngLoop As Long
    Dim i As Integer
    Dim lngLine As Long
    Dim objFileLines As Long
    Dim strLine As String
    Dim strFMT As String
    Dim blnBlock As Boolean
    Dim blnSQL As Boolean
    Dim strPro As String
    Dim strTemp As String
    Dim strFileProName As String
    Dim objFileTemp As TextStream
    Dim objFolder As Folder
    Dim objFolderTemp As Folder
    Dim objCurFolder As Folder
    Dim objFile As File
    Dim objFileFlag As File
    Dim rsInit As ADODB.Recordset
    Dim intSysNumLast As Integer
    Dim lngTemp As Long
    Dim strCommand As String
    Dim lngProcess As Long
    Dim rsSQL As ADODB.Recordset
    Dim blnNew As Boolean
    Dim strOwner As String
    Dim strIniPath As String
    Dim strIni1 As String
    Dim strIniSys As String
    Dim strIniApp As String
    Dim lngSys As Long
    
    Dim objPercent As New clsPercent
    
    On Error GoTo errHand
    
    cmdOK.Enabled = False
    
    Call gclsBase.SQLRecord(rsSQL)
    
    lblTitle.Caption = "���������ʱĿ¼.."
    lblTitle.Visible = True
    DoEvents
    
    strTmp1 = App.Path & "\Tmp1"
    strProcedure = App.Path & "\Procedure"
    strTmpReports = App.Path & "\Reports"
    
    If mcnOracle Is Nothing Then
        MsgBox "���Ƚ����������ã���ȷ���Ѽ���Դ��", vbInformation + vbOKOnly, "�������"
        Exit Sub
    End If
       
    '------------------------------------------------------------------------------------------------------------------
    With vsf(0)
        
'        strSQL = "Delete From zlproceduretext where ���� in (1,2))"
'        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
'        strSQL = "Delete from zlproceduretext where ����id in (select id from zlprocedure where ���� in (1,3))"
'        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
'        strSQL = "Delete from zlprocedure where ���� in (1,3)"
'        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        
                
        strSQL = "Select ���,����,�汾�� From zlSystems a"
        Set rsData = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "")
        If rsData.BOF = True Then
            MsgBox "��ǰ���ݿ�û�а�װ�κ�ϵͳ��", vbInformation + vbOKOnly, "�������"
            GoTo errEnd
        End If
        For i = 1 To .Rows - 1
            rsData.Filter = ""
            rsData.Filter = "���=" & .RowData(i)
            If .TextMatrix(i, vsf(0).ColIndex("��װ�ű�")) = "" Then
                MsgBox "��ѡ��" & .TextMatrix(i, .ColIndex("ϵͳ����")) & "��װ�ű�"
                GoTo errEnd
            End If
            Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")))
            rsInit.Filter = "��Ŀ='�汾��'"
            strIniApp = rsInit("����").value '��װ�ű��汾��

            rsData.Filter = ""
            rsData.Filter = "���=" & .RowData(i)
            strIniSys = Trim(rsData("�汾��").value) '���ݿ�汾��

            If strIniSys <> strIniApp Then
                MsgBox .TextMatrix(i, .ColIndex("ϵͳ����")) & "���ݿ�ϵͳ�汾�������ļ��汾��ƥ�䡣", vbInformation + vbOKOnly, "�������"
                GoTo errEnd
            End If
        Next
    End With
    If gobjFile.FolderExists(strTmp1) Then Call gobjFile.DeleteFolder(strTmp1, True)
    If gobjFile.FolderExists(strProcedure) Then Call gobjFile.DeleteFolder(strProcedure)
    If gobjFile.FolderExists(strTmpReports) Then gobjFile.DeleteFolder (strTmpReports)
        
    DoEvents
    
    Call gobjFile.CreateFolder(strTmpReports)
    Call gobjFile.CreateFolder(strTmp1)
    Call gobjFile.CreateFolder(strProcedure)
    
    '------------------------------------------------------------------------------------------------------------------
    '�����ݿ�������ɵ����ű��ļ�
    lblTitle.Caption = "����׼�����ݿ�.."
    strSQL = "Select Name,Type,Text From user_source Where type in ('PROCEDURE','FUNCTION') Order by Name,Line"
    Set rs = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "")
    
    If rs.BOF = False Then
        pbr.Visible = True
        
        Call objPercent.InitPercent(pbr, rs.RecordCount)
                
        lblTitle.Visible = True
        strFlag = ""
        For lngLoop = 0 To rs.RecordCount - 1
            If strFlag <> Nvl(rs("Name").value) And strFlag <> "" Then
            
                '�ж��ļ����Ƿ�Ƿ�
                If Not (InStr(strFlag, "\") > 0 Or _
                    InStr(strFlag, "/") > 0 Or _
                    InStr(strFlag, ":") > 0 Or _
                    InStr(strFlag, " ") > 0 Or _
                    InStr(strFlag, "*") > 0 Or _
                    InStr(strFlag, "?") > 0 Or _
                    InStr(strFlag, """") > 0 Or _
                    InStr(strFlag, "<") > 0 Or _
                    InStr(strFlag, ">") > 0 Or _
                    InStr(strFlag, "|") > 0) Then
                    
                    '�����������̽ű��ļ�
                    Set objFileTemp = gobjFile.CreateTextFile(strTmp1 & "\" & strFlag & ".sql", True)
                    
                    'д����һ����ѯ���Ĺ���
                    Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                       strTemp = Left(strTemp, Len(strTemp) - 1)
                    Loop
                    
                    DoEvents
                    
                    strTemp = "CREATE OR REPLACE " & strTemp
                    objFileTemp.Write strTemp
                End If
                
                strTemp = ""
                strTemp = strTemp & UCase(Nvl(rs("Text").value))
                    
            ElseIf strFlag = "" Then
                
                strTemp = strTemp & UCase(Nvl(rs("Text").value))
                
            Else
                strTemp = strTemp & Nvl(rs("Text").value)
            End If
            
            strFlag = Nvl(rs("Name").value)
            rs.MoveNext
            
            Call objPercent.LoopPercent
        Next
        
        If strTemp <> "" Then
        
            '�ж��ļ����Ƿ�Ƿ�
            If Not (InStr(strFlag, "\") > 0 Or _
                InStr(strFlag, "/") > 0 Or _
                InStr(strFlag, ":") > 0 Or _
                InStr(strFlag, " ") > 0 Or _
                InStr(strFlag, "*") > 0 Or _
                InStr(strFlag, "?") > 0 Or _
                InStr(strFlag, """") > 0 Or _
                InStr(strFlag, "<") > 0 Or _
                InStr(strFlag, ">") > 0 Or _
                InStr(strFlag, "|") > 0) Then
                
                '�����������̽ű��ļ�
                Set objFileTemp = gobjFile.CreateTextFile(strTmp1 & "\" & strFlag & ".sql", True)
                
                'д����һ����ѯ���Ĺ���
                Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                   strTemp = Left(strTemp, Len(strTemp) - 1)
                Loop
                
                DoEvents
                strTemp = "CREATE OR REPLACE " & strTemp
                objFileTemp.Write strTemp
            End If
            
            strTemp = ""
        End If
        objFileTemp.Close
        pbr.Visible = False
    End If
    
    
    '------------------------------------------------------------------------------------------------------------------
    For i = 1 To vsf(0).Rows - 1
        
        '��ȡ��װ�ű��������ű��Ĺ��������ɵ����ű��ļ�
        '��ȡ��װ�ű�
        If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�"))) Then
            MsgBox "�޷��򿪽ű��ļ�" & vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")) & ",ִ���жϡ�", vbExclamation, gstrSysName
            GoTo errEnd
        Else
            strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�"))) - 11)
            strIniPath = strIniPath & "zlProgram.sql"
        End If
        lblTitle.Caption = "������ȡ��" & vsf(0).TextMatrix(i, vsf(0).ColIndex("ϵͳ����")) & "����װ�ű�.."
        
        Call CheckProcedure(strIniPath, strProcedure)
        
        pbr.value = 0
        pbr.Visible = False
        DoEvents
        
        '��ȡ�����ű�
        strIniSys = vsf(0).TextMatrix(i, vsf(0).ColIndex("�汾��"))
        If Split(strIniSys, ".")(2) = 0 Then
            GoTo errNext
        ElseIf Not gobjFile.FolderExists(Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")), "Ӧ�ýű�")(0) & "�����ű�\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0") Then
            MsgBox "�޷���⵽�����ű��ļ���,ִ���жϡ�", vbExclamation, gstrSysName
            GoTo errEnd
        Else
            strIniPath = Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")), "Ӧ�ýű�")(0) & "�����ű�\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0" & "\"
        End If
'        If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�"))) And vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")) <> "" Then
'            MsgBox "�޷��򿪽ű��ļ�" & vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")) & ",ִ���жϡ�", vbExclamation, gstrSysName
'            GoTo errEnd
'        ElseIf Trim(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�"))) = "" Then
'            GoTo errNext
'        Else
'            strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�"))) - 13)
'        End If

'        Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")))
'        If Not CheckINIValid(rsInit, "ϵͳ��|Ŀ��汾") Then
'            MsgBox "��Ǩ�����ļ���ʽ����ȷ��", vbExclamation, "�������"
'            GoTo errEnd
'        End If
        lblTitle.Caption = "������ȡ" & vsf(0).TextMatrix(i, vsf(0).ColIndex("ϵͳ����")) & "�����ű�.."
'        rsInit.Filter = "��Ŀ='Ŀ��汾'"
'        intSysNumLast = Split(rsInit("����").value, ".")(2) '�õ����������ļ��İ汾��
        intSysNumLast = Split(strIniSys, ".")(2)
        For lngLoop = 10 To intSysNumLast Step 10
            strFlag = Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & "." & CStr(lngLoop)
            Call CheckProcedure(strIniPath & "ZL" & vsf(0).RowData(i) / 100 & "_" & strFlag & ".sql", strProcedure)
        Next
errNext:

    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '�����ݿ��еĹ�����ű����бȶԣ�����html����
    strCommand = GetWinSystemPath & "\wincmp3.exe " & strTmp1 & "\ " & strProcedure & "\ /G:HE " & strTmpReports
    err.Clear
    DoEvents
    lblTitle.Caption = "���ڱȽ�.."
    lngTemp = Shell(strCommand, vbHide)
    DoEvents
    If err <> 0 Then
        err.Clear
         MsgBox "�ļ��Ƚ�ʧ�ܣ�����" & GetWinSystemPath & "\wincmp3.exe�ļ��Ƿ����", vbExclamation, "�������"
        GoTo errEnd
    End If
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    CloseHandle lngProcess
    
    DoEvents
    
    '------------------------------------------------------------------------------------------------------------------
    
    Set objFolder = gobjFile.GetFolder(strTmpReports)
    
    '�����д��ڵļ�Ϊ��Ҫ�����Ĺ���
    For Each objFile In objFolder.Files
        Dim strFileName As String
        Dim lngKey As Long
        Dim strContent As String
        Dim lngMaxLength As Long
        Dim str As String
        Dim lngRow As Long
        Dim strArr() As String
        
        DoEvents
        
        strFileName = Split(objFile.name, ".")(0)
        lblTitle.Caption = "���ڲ����䶯���̣�" & strFileName
        
        
        '��ȡ����������
        strOwner = gclsBase.GetOwnerInfo(strFileName)
        
        Set rs = gclsBase.GetProInfo(strFileName)
        
        '���̲����ڣ��Զ����Ϊ�䶯���̻�հ׹���
        '-------------------------------------------------------------------------------------------------------------------------------------------------
        If rs.BOF = False Then
            lngKey = Nvl(rs("ID").value)
            If rs("����").value = 2 Then
                strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.�հ׹��� & ",'" & strFileName & "'," & ProcState.�ѵ��� & ",'','" & strOwner & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            Else
                strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.�䶯���� & ",'" & strFileName & "'," & ProcState.�ѵ��� & ",'','" & strOwner & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            End If
        Else
            lngKey = gclsBase.GetNextId("zlProcedure")
            strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.�䶯���� & ",'" & strFileName & "'," & ProcState.�ѵ��� & ",'','" & strOwner & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        End If
        
        
        '���汾���Զ������
        Set objFileTemp = gobjFile.OpenTextFile(strTmp1 & "\" & Split(objFile.name, ".")(0) & "." & Split(objFile.name, ".")(1))
        '��ȡ��������
        strContent = ""
        Do While Not objFileTemp.AtEndOfStream
            strLine = objFileTemp.ReadLine
            If strContent = "" Then
                strContent = strContent & Replace(strLine, "'", "''")
            Else
                strContent = strContent & vbCrLf & Replace(strLine, "'", "''")
            End If
        Loop
        objFileTemp.Close
        lngMaxLength = 3900
        If LenB(StrConv(strContent, vbFromUnicode)) > lngMaxLength Then
            strFlag = ""
            str = ""
            For lngRow = 1 To Len(strContent)
                str = str & Mid(strContent, lngRow, 1)
                If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngRow = Len(strContent)) And Mid(strContent, lngRow, 1) <> "'" Then
                    strFlag = strFlag & gstrSplite & str
                    str = ""
                End If
            Next
            strFlag = Mid(strFlag, Len(gstrSplite) + 1)
            strContent = strFlag
        End If
        strArr = Split(strContent, gstrSplite)
        '��ɾ������
        strSQL = "Zl_Zlproceduretext_Delete(" & lngKey & "," & ProcTextType.�����Զ����� & ")"
        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        '�ٲ�������
        For lngRow = 0 To UBound(strArr)
            strSQL = "Zl_Zlproceduretext_Update(" & lngKey & "," & ProcTextType.�����Զ����� & " ," & (lngRow + 1) & ",'" & strArr(lngRow) & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        Next
               
        
        '���α�׼����
        '-------------------------------------------------------------------------------------------------------------------------------------------------
        Set objFileTemp = gobjFile.OpenTextFile(strProcedure & "\" & Split(objFile.name, ".")(0) & "." & Split(objFile.name, ".")(1))
        '��ȡ��������
        strContent = ""
        Do While Not objFileTemp.AtEndOfStream
            strLine = objFileTemp.ReadLine
            If strContent = "" Then
                strContent = strContent & Replace(strLine, "'", "''")
            Else
                strContent = strContent & vbCrLf & Replace(strLine, "'", "''")
            End If
        Loop
        objFileTemp.Close
        lngMaxLength = 3900
        If LenB(StrConv(strContent, vbFromUnicode)) > lngMaxLength Then
            strFlag = ""
            str = ""
            For lngRow = 1 To Len(strContent)
                str = str & Mid(strContent, lngRow, 1)
                If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngRow = Len(strContent)) And Mid(strContent, lngRow, 1) <> "'" Then
                    strFlag = strFlag & gstrSplite & str
                    str = ""
                End If
            Next
            strFlag = Mid(strFlag, Len(gstrSplite) + 1)
            strContent = strFlag
        End If
        strArr = Split(strContent, gstrSplite)
        '��ɾ������
        strSQL = "Zl_Zlproceduretext_Delete(" & lngKey & "," & ProcTextType.���α�׼���� & ")"
        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        For lngRow = 0 To UBound(strArr)
            strSQL = "Zl_Zlproceduretext_Update(" & lngKey & "," & ProcTextType.���α�׼���� & "," & (lngRow + 1) & ",'" & strArr(lngRow) & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
        Next
    Next
    
    '------------------------------------------------------------------------------------------------------------------
    lblTitle.Caption = "���ڲ����û�����.."
    pbr.Visible = True
    Set objFolder = gobjFile.GetFolder(strTmp1)
        
    Call objPercent.InitPercent(pbr, objFolder.Files.Count)

    For Each objFile In objFolder.Files
        lblTitle.Caption = "���ڲ����û����̣�" & objFile.name
        DoEvents
                
        blnNew = False
        If gobjFile.FileExists(strProcedure & "\" & objFile.name) Then
            blnNew = True
        End If
        If blnNew = False Then
            
            '���ݿ��еĹ����ڽű���û�У�˵�����û�����
            '����û�����

            strFileName = Split(objFile.name, ".")(0)
            Set rs = gclsBase.GetProInfo(strFileName)
            If rs.BOF = False Then
                lngKey = Nvl(rs("ID").value)
            Else
                '���̲����ڣ��Զ����Ϊ�û�����
                lngKey = gclsBase.GetNextId("zlProcedure")
            End If
            Set objFileTemp = gobjFile.OpenTextFile(strTmp1 & "\" & Split(objFile.name, ".")(0) & "." & Split(objFile.name, ".")(1))
            
            '��ȡ��������
            strContent = ""
            Do While Not objFileTemp.AtEndOfStream
                strLine = objFileTemp.ReadLine
                If strContent = "" Then
                    strContent = strContent & Replace(strLine, "'", "''")
                Else
                    strContent = strContent & vbCrLf & Replace(strLine, "'", "''")
                End If
            Loop
            
            
            
            objFileTemp.Close
            lngMaxLength = 3900
            If LenB(StrConv(strContent, vbFromUnicode)) > lngMaxLength Then
                strFlag = ""
                str = ""
                For lngRow = 1 To Len(strContent)
                    str = str & Mid(strContent, lngRow, 1)
                    If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngRow = Len(strContent)) And Mid(strContent, lngRow, 1) <> "'" Then
                        strFlag = strFlag & gstrSplite & str
                        str = ""
                    End If
                Next
                strFlag = Mid(strFlag, Len(gstrSplite) + 1)
                strContent = strFlag
            End If
            
            strArr = Split(strContent, gstrSplite)
            
            strOwner = gclsBase.GetOwnerInfo(strFileName)
            strSQL = "Zl_Zlprocedure_Update(" & lngKey & "," & ProcType.�û����� & ",'" & strFileName & "'," & ProcState.�ѵ��� & ",'','" & strOwner & "')"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            
            '��ɾ������
            strSQL = "Zl_Zlproceduretext_Delete(" & lngKey & "," & ProcTextType.�����Զ����� & ")"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            For lngRow = 0 To UBound(strArr)
                strSQL = "Zl_Zlproceduretext_Update(" & lngKey & "," & ProcTextType.�����Զ����� & "," & (lngRow + 1) & ",'" & strArr(lngRow) & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            Next
            
        End If
        
        Call objPercent.LoopPercent
        
    Next
    
    lblTitle.Caption = "�����ύ����..."
    DoEvents
    
    On Error Resume Next
    objFileTemp.Close
    Set objFileTemp = Nothing
    On Error GoTo errHand
    
    Call SQLRecordExecute(rsSQL, "")
    
    lblTitle.Caption = "���������ʱ����..."
    DoEvents
    'ɾ����ʱ�ļ���
    '------------------------------------------------------------------------------------------------------------------
    If gobjFile.FolderExists(strTmp1) Then
        Call gobjFile.DeleteFolder(strTmp1, True)
    End If
    If gobjFile.FolderExists(strProcedure) Then
        Call gobjFile.DeleteFolder(strProcedure, True)
    End If
    If gobjFile.FolderExists(strTmpReports) Then
        Call gobjFile.DeleteFolder(strTmpReports, True)
    End If
    
    MsgBox "�����Ǽ���ɣ�", vbInformation, Me.Caption
    lblTitle.Visible = False
    pbr.Visible = False
    mblnOK = True
    cmdOK.Enabled = True
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errEnd:
    mblnOK = True
    cmdOK.Enabled = True
    Exit Sub
errHand:
    MsgBox "�����Ǽ�ʧ�ܣ�" & vbCrLf & err.Description, vbInformation, Me.Caption
    cmdOK.Enabled = True
End Sub

Private Sub Form_Load()
    Call ExecuteCommand("��ʼ�ؼ�")
    
    Call opt_Click(0)
    
End Sub

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'˵������Ҫ��RunSQLFile���Ӻ���
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    Do While InStr(strText, "  ") > 0
        strText = Replace(strText, "  ", " ")
    Loop
    TrimEx = strText
End Function

Public Function TrimComment(ByVal strSQL As String) As String
'���ܣ�ȥ��д�ڵ���strSQL�������"--"ע��
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim blnStr As Boolean
    Dim i As Long, k As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                k = i: Exit For
            End If
        Next
        If k > 0 Then strSQL = RTrim(Left(strSQL, k - 1))
    End If
    TrimComment = strSQL
End Function

Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'���ܣ���ָ��INI�����ļ������ݶ�ȡ����¼����
'���أ�Nothing�����"��Ŀ,����"�ļ�¼��,����ͬһ��Ŀ�����ж�������
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "��Ŀ", adVarChar, 100
    rsTmp.Fields.Append "����", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = Null
                rsTmp.Update
            End If
            
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!��Ŀ = strItem
            rsTmp!���� = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!��Ŀ = strItem
        rsTmp!���� = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Private Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'���ܣ�����Ӧ�������ļ���ʽ�Ƿ���ȷ
'������rsINI=��������ļ����ݵļ�¼��������"��Ŀ,����"�ֶ�
'      strItem=�����ļ��б���Ҫ�������ݵ���Ŀ��,��"��Ŀ1|��Ŀ2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "��Ŀ='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If IsNull(rsINI!����) Then Exit Function
    Next
    CheckINIValid = True
End Function

Private Function CheckProcedure(ByVal strFile As String, Optional strFilePath As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim lngLine As Long
    Dim strLine As String
    Dim strTemp As String
    Dim strFMT As String
    Dim blnSQL As Boolean
    Dim blnBlock As Boolean
    Dim strFlag As String
    Dim strFileProName As String
    Dim lngFileLines As Long
    Dim objFileTemp As TextStream
    Dim objFile As TextStream
    Dim blnFlag As Boolean
    Dim objPercent As New clsPercent
    Dim lngMsg As Long
    
    On Error GoTo errHand
    
    pbr.value = 0
    pbr.Visible = True

    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    If objFile.AtEndOfStream Then
        objFile.Close
        Exit Function
    End If
        
    Do While Not objFile.AtEndOfStream
        objFile.ReadLine
    Loop
    lngFileLines = objFile.Line
    
    Call objPercent.InitPercent(pbr, lngFileLines)
    
    objFile.Close
    
    Dim blnSpaceProc As Boolean
    
    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objFile.AtEndOfStream
        lngLine = objFile.Line '��ǰ�к�:δ��ȡ��֮ǰ,��ָ��δ�Ƶ���һ��
        strLine = objFile.ReadLine
        strFMT = UCase(TrimComment(TrimEx(strLine)))
        If strFMT Like "PROMPT *" Then GoTo NextLine
        
        
        If blnBlock Then
            If strFMT = "/" Then
                blnSQL = True
                blnBlock = False
                Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                   strTemp = Left(strTemp, Len(strTemp) - 1)
                Loop
                
                
                objFileTemp.Write "CREATE OR REPLACE " & strTemp
                DoEvents
                objFileTemp.Close
                strTemp = ""
                
                If blnSpaceProc = True Then
                    blnSpaceProc = False
                    
                    Set objFileTemp = gobjFile.OpenTextFile(strFilePath & "\" & strFileProName & ".sql")
                    strTemp = objFileTemp.ReadAll
                    objFileTemp.Close
                    strTemp = GetBlankProcedure(strTemp)
                    
                    DoEvents
                    Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
                    objFileTemp.Write strTemp
                    objFileTemp.Close
                    strTemp = ""
                End If
                
            Else
                strTemp = strTemp & vbCrLf & strLine
            End If
        ElseIf strFMT Like "CREATE OR REPLACE PROCEDURE *" Or strFMT Like "CREATE PROCEDURE *" _
            Or strFMT Like "CREATE OR REPLACE FUNCTION *" Or strFMT Like "CREATE FUNCTION *" _
            Or strFMT Like "CREATE OR REPLACE TRIGGER *" Or strFMT Like "CREATE TRIGGER *" _
            Or strFMT Like "CREATE OR REPLACE TYPE *" Or strFMT Like "CREATE TYPE *" _
            Or strFMT Like "CREATE OR REPLACE PACKAGE *" Or strFMT Like "CREATE PACKAGE *" Then
            
            blnBlock = True
            
            '�����������̽ű��ļ�
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            
            If InStr(strFlag, "(") > 0 Then strFlag = Left(strFlag, InStr(strFlag, "(") - 1)
            If InStr(strFlag, ".") > 0 Then strFlag = Split(strFlag, ".")(1)
            strFileProName = Split(strFlag, " ")(1)
            If gobjFile.FileExists(strFilePath & "\" & strFileProName & ".sql") Then
                Call gobjFile.DeleteFile(strFilePath & "\" & strFileProName & ".sql")
            End If
            
            '����Ƿ�Ϊ�հ׹���
            blnSpaceProc = False
            If IsSpaceProcedure("ZLHIS", strFileProName) = True Then
                blnSpaceProc = True
            End If
            
            Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
             
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            strTemp = strTemp & UCase(strFlag)
        End If
        
        Call objPercent.LoopPercent

NextLine:
    Loop
    objFile.Close
    pbr.Visible = False
    pbr.value = 0
'    MsgBox blnFlag
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Function GetWinPath() As String
    '--------------------------------------------------------------------------------------------------------------
    '--����:��ȡϵͳĿ¼
    '--------------------------------------------------------------------------------------------------------------
    Dim Buffer As String
    Dim StrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    StrWinPath = Left(Buffer, rtn)
    GetWinPath = StrWinPath
End Function

Private Sub Form_Resize()
    On Error Resume Next
    vsf(0).Move 15, 15, picPane(0).ScaleWidth - 30, picPane(0).ScaleHeight - 30
    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
    If Not (mfrmPageConfigure Is Nothing) Then
        Unload mfrmPageConfigure
    End If
'    Call InitCommon(gcnOracle)
End Sub

Private Sub mfrmPageConfigure_AfterConn(ByVal cnOracle As ADODB.Connection)
    Set mcnOracle = cnOracle
    
    Call ExecuteCommand("��ʼ����")
'    lblState.Caption = "������"
'    lblState.ForeColor = &HC000&
End Sub

Private Sub opt_Click(Index As Integer)
    cmdConnet.Enabled = (opt(1).value = True)
    
    Select Case Index
    Case 0
        Set mcnOracle = gcnOracle
        Call ExecuteCommand("��ʼ����")
'        lblState.Caption = "������"
'        lblState.ForeColor = &HC000&
    Case 1
        mclsVsf.ClearGrid
    End Select
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsf(0)
        Select Case Col
        '--------------------------------------------------------------------------------------------------------------
        Case .ColIndex("��װ�ű�")
            With dlg
                .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
                .Filter = "(Ӧ�ð�װ�����ļ�)|zlSetup.ini"
                .ShowOpen
                If .FileName = "" Then
                    Exit Sub
                Else
                    vsf(0).TextMatrix(vsf(0).Row, vsf(0).Col) = .FileName
                End If
            End With
        End Select
        
        Call mclsVsf.SetFocus(, , True)
    End With
End Sub

Private Sub vsf_DblClick(Index As Integer)
    mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub




