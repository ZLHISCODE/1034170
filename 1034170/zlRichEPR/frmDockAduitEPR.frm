VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockAduitEPR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picThis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   5865
      ScaleHeight     =   630
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   1
      Left            =   1080
      ScaleHeight     =   705
      ScaleWidth      =   6135
      TabIndex        =   2
      Top             =   3285
      Visible         =   0   'False
      Width           =   6135
      Begin MSComctlLib.ListView lvwThis 
         Height          =   330
         Left            =   495
         TabIndex        =   3
         Top             =   0
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   582
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         _Version        =   393217
         Icons           =   "imgThis"
         SmallIcons      =   "imgThis"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "�ļ�"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblThis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   75
         Width           =   450
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   0
      Left            =   1350
      ScaleHeight     =   1980
      ScaleWidth      =   2685
      TabIndex        =   0
      Top             =   705
      Width           =   2685
      Begin zlRichEditor.Editor edtThis 
         Height          =   1245
         Left            =   315
         TabIndex        =   1
         Top             =   315
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   2196
      End
   End
   Begin MSComctlLib.ImageList imgThis 
      Left            =   5610
      Top             =   885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   360
      Top             =   1905
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmDockAduitEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mblnAddition As Boolean
Private mlngKey As Long

Private Enum FileType
    conPane_RichEpr = 1
    conPane_TablEpr = 2
    conPane_Annex = 3
End Enum
Private Enum ICON_SIZE
    ICON_SMALL = 16
    ICON_LARGE = 32
End Enum

Public Event PrintEpr(ByVal lngRecordId As Long)

Private mObjTabEprView As cTableEPR      '������
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mblnDataMove As Boolean
Private mfrmMain As Object
'######################################################################################################################
Public Function zlRefresh(ByVal frmMain As Object, ByVal lngKey As Long, ByVal blnDataMove As Boolean, Optional ByVal bytKind As Byte = 0) As Boolean
    Dim rs As New ADODB.Recordset
    Set mfrmMain = frmMain
    mlngKey = lngKey
    mblnDataMove = blnDataMove
    
    LockWindowUpdate Me.hWnd

    gstrSQL = "Select 1 From ���Ӳ�����ʽ Where �ļ�id = [1] And ���� Is Not Null"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�����ʽ", "H���Ӳ�����ʽ")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If Not rs.EOF Then
        Call OpenEPR(lngKey)
    Else
        Call OpenSignleEPR(lngKey)
    End If
    
    '����Ƿ��в�������
    gstrSQL = "Select 1 From ���Ӳ������� Where ����id = [1]"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call ShowAdition(lngKey)
        
        dkpMain.ShowPane conPane_Annex
    Else
        dkpMain.FindPane(conPane_Annex).Close
    End If
    
    LockWindowUpdate 0
End Function

Public Function zlPrintDocument(ByVal eDocType As EPRDocTypeEnum, Optional ByVal bytMode As Byte = 2, Optional ByVal lngKey As Long, Optional ByVal strPrintDeviceName As String) As Boolean
    '1-Ԥ��,2-��ӡ
    Dim lngEPRKey As Long
    Dim rs As New ADODB.Recordset
    Dim strReportCode As String
    Dim intOutMode As Integer
    Dim strPrinterName As String
    Dim strPdfFile As String
    
    If InStr(strPrintDeviceName, "TinyPDF|") > 0 Then
        strPrinterName = Split(strPrintDeviceName, "|")(0)
        strPdfFile = Split(strPrintDeviceName, "|")(1)
    Else
        strPrinterName = strPrintDeviceName
    End If

    If eDocType = cpr���Ʊ��� Then
        gstrSQL = "Select f.ͨ��, f.��� From ���Ӳ�����¼ l, �����ļ��б� f Where l.�ļ�id = f.Id And l.Id = [1]"
        If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.EOF Then Exit Function
        intOutMode = Val("" & rs!ͨ��)
        strReportCode = "ZLCISBILL" & Format(rs!���, "00000") & "-2"
        
        gstrSQL = "Select b.��¼����, b.No, b.ҽ��id, c.����id,(Select ������� From ����ҽ����¼ Where ���ID=C.ID and Rownum<2) as �������,B.���ͺ�" & vbNewLine & _
                    "From ����ҽ������ A, ����ҽ������ B, ����ҽ����¼ C" & vbNewLine & _
                    "Where a.����id = [1] And a.ҽ��id = b.ҽ��id And c.Id = b.ҽ��id"
        If mblnDataMove Then
            gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.EOF Then Exit Function
    End If

    If intOutMode = 2 Then
        '�����Զ��屨����ӡ
        If strPrinterName <> "" Then Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & strReportCode, "Printer", strPrinterName)
        If rs!������� = "C" Then '����
            Call Open_LIS_Report(Me, lngKey, strReportCode, rs!ҽ��id, rs!����ID, rs!���ͺ�, rs!��¼����, rs!NO, False, bytMode, strPdfFile)
        Else '���
            Call Open_Pacs_Report(Me, lngKey, strReportCode, rs!ҽ��id, rs!����ID, rs!���ͺ�, rs!��¼����, rs!NO, False, bytMode, strPdfFile)
        End If
    Else
        'EPR��ӡ
        gstrSQL = "select a.����ID,a.��ҳID,a.��������,a.������Դ,a.�༭��ʽ from ���Ӳ�����¼ a where a.ID=[1] "
        If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")

        lngEPRKey = IIf(lngKey > 0, lngKey, mlngKey)

        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRKey)
        If Not rs.EOF Then
            If NVL(rs!�༭��ʽ, 0) = 0 Then
                eDocType = Val(rs!��������)
                Set mfrmPrintPreview = New frmPrintPreview
                If strPrinterName <> "" Then
                    mfrmPrintPreview.DoMultiDocPreview mfrmMain, eDocType, zlCommFun.NVL(rs("����ID").Value, 0), zlCommFun.NVL(rs("��ҳID").Value, 0), zlCommFun.NVL(rs("��������").Value, 1), "", lngEPRKey, IIf(bytMode = 1, False, True), False, True, mblnDataMove, , strPrinterName
                Else
                    mfrmPrintPreview.DoMultiDocPreview mfrmMain, eDocType, zlCommFun.NVL(rs("����ID").Value, 0), zlCommFun.NVL(rs("��ҳID").Value, 0), zlCommFun.NVL(rs("��������").Value, 1), "", lngEPRKey, IIf(bytMode = 1, False, True), False, False, mblnDataMove
                End If
                Unload mfrmPrintPreview 'ByZT:����Load��δ��ʾ��û����Ϊ�رյ������VB�����Զ�Unload
                Set mfrmPrintPreview = Nothing
            Else
                Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_���������, lngEPRKey, False, 0, rs!������Դ)
                mObjTabEprView.zlPrintDoc Me, bytMode = 1, strPrinterName
            End If
        End If
    End If
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    If Not Me.Visible Then Exit Function
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "�˵���"
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    
    Set objBar = cbsMain.Add("��׼", xtpBarTop)
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = objBar.Controls.Add(xtpControlButton, conMenu_File_Open, "�򿪲�����ϸ����...")
    cbsMain.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    
End Function
Private Function OpenSignleEPR(ByVal lngEPRid As Long) As Boolean
    Dim strPath As String, strFile As String, lngs As Long, lnge As Long
    Dim rs As New ADODB.Recordset
    Dim Doc As New cEPRDocument, Elements As New cEPRElements
    Dim lng����ID As Long, lng��ҳID As Long, byt�������� As EPRDocTypeEnum, lng�༭��ʽ As Long, lngKey As Long, blnPrivacy As Boolean
    
    On Error GoTo errHand
    
    lngs = GetTickCount
    Screen.MousePointer = vbHourglass
    DoEvents
    LockWindowUpdate Me.hWnd

    gstrSQL = "select ����ID,��ҳID,��������,�༭��ʽ from ���Ӳ�����¼ where ID=[1]"
    If mblnDataMove Then
        gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
    If Not rs.EOF Then
        lng����ID = NVL(rs("����ID").Value, 0)
        lng��ҳID = NVL(rs("��ҳID").Value, 0)
        byt�������� = NVL(rs("��������").Value, 1)
        lng�༭��ʽ = NVL(rs("�༭��ʽ").Value, 0)
    End If
    rs.Close
    
    edtThis.ForceEdit = True
    
    '������ʱ�ļ�
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    strFile = strPath & "\" & App.hInstance & CLng(Timer) & ".TMP"
    
    Doc.InitEPRDoc cprEM_�޸�, cprET_���������, lngEPRid, byt��������, lng����ID, CStr(lng��ҳID), , , , mblnDataMove
    Doc.OpenEPRDoc Doc.frmEditor.Editor1,mblnDataMove        '�򿪸��ļ�
    
    '�����滻��Ŀ
    If blnPrivacy Then
        '��ȡ���е�Ҫ��
        gstrSQL = "Select A.ID,A.������ From ���Ӳ������� A, ��˽������Ŀ B,����������Ŀ C " & _
            "Where A.�������� = 4 And A.�滻�� = 1 And A.�ļ�id = [1] And A.������� > 0 and B.��Ŀid = C.ID And A.Ҫ������ =C.������ And C.�滻�� = 1 "
            
        If mblnDataMove Then
            gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
        If Not rs.EOF Then
            Do While Not rs.EOF
                lngKey = Elements.Add(NVL(rs("������"), 0))
                Elements("K" & lngKey).GetElementFromDB cprET_�������༭, rs("ID"), True, "���Ӳ�������"
                '�滻Ҫ������
                Elements("K" & lngKey).�����ı� = String(Len(Elements("K" & lngKey).�����ı�), "*")
                Elements("K" & lngKey).Refresh Doc.frmEditor.Editor1
                rs.MoveNext
            Loop
        End If
        rs.Close
    End If
    Doc.frmEditor.SaveDocToFile strFile, False     '�洢�������ʱ�ļ�

    With edtThis
        If lng�༭��ʽ = 0 Then
            dkpMain.FindPane(conPane_TablEpr).Close
            dkpMain.ShowPane conPane_RichEpr
        Else
            dkpMain.FindPane(conPane_RichEpr).Close
            dkpMain.ShowPane conPane_TablEpr
        End If
        .Freeze
        .ReadOnly = False
        .NewDoc
        .ForceEdit = True
        .ViewMode = cprNormal
        .OpenDoc strFile

        '����ҳüҳ��
        Set .Picture = Doc.frmEditor.Editor1.Picture
        .Head = Doc.frmEditor.Editor1.Head
        .Foot = Doc.frmEditor.Editor1.Foot
        
'        Doc.frmEditor.Editor1.DocHeadCopyWithFormat         '����༭����Copy����ʽҳüҳ��
'        .DocHeadPasteWithFormat                             '�ؼ�ҳüҳ��ճ��
'        Doc.frmEditor.Editor1.DocFootCopyWithFormat
'        .DocFootPasteWithFormat
'        Call Doc.GetReplacedHeadFootString(edThis)          '���෽�������ؼ�ҳüҳ���е�Ҫ��
                
        .PaperWidth = Doc.frmEditor.Editor1.PaperWidth
        .PaperHeight = Doc.frmEditor.Editor1.PaperHeight
        .MarginLeft = Doc.frmEditor.Editor1.MarginLeft
        .MarginRight = Doc.frmEditor.Editor1.MarginRight
        .MarginTop = Doc.frmEditor.Editor1.MarginTop
        .MarginBottom = Doc.frmEditor.Editor1.MarginBottom

        '����ҳ���ʽ
        Doc.EPRFileInfo.SetFormat edtThis, Doc.EPRFileInfo.��ʽ
        edtThis.ResetWYSIWYG    'ˢ�����������ã�WYSIWYG����ʾ

        '��ҳ
        .SelectAll
        .AuditMode = True
        .AcceptAuditText
        .ViewMode = cprNormal
        .Range(0, 0).Selected
        .ForceEdit = False
        .UnFreeze
        .ReadOnly = True
    End With

    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile    'ɾ����ʱ�ļ�
 
    Doc.frmEditor.Editor1.Modified = False
    
    Set rs = Nothing
    lnge = GetTickCount
'    Debug.Print "��ȡ��ʱ" & lnge - lngs
    LockWindowUpdate 0
    zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    
    OpenSignleEPR = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function SetRichDocsPos(ByVal lngRecordId As Long) As Boolean
    'ͨ��ID�ȶ�λ���޷���λʱ�ټ���
    On Error GoTo errHand
    Dim lngKSS As Long, lngKSE As Long, lngKES As Long, lngKEE As Long, blnNeed As Boolean, lngKey As Long, lngLen As Long, i As Long
    lngLen = Len(edtThis.Text)
    For i = 0 To lngLen
        If FindNextKey(edtThis, i, "F", lngKey, lngKSS, lngKSE, lngKES, lngKEE, blnNeed) Then
            If edtThis.Range(lngKSE, lngKES).Text = lngRecordId Then
                Call edtThis.Range(lngKSS, lngKEE).ScrollIntoView(cprSPStart)  '  .Selected
                SetRichDocsPos = True
                Exit Function
            End If
            i = lngKEE
        Else
            Exit Function
        End If
    Next
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function OpenEPR(ByVal lngEPRid As Long) As Boolean
'******************************************************************************************************************
'���ܣ�ˢ�²�����ʾ���ݣ�
'������lngEPRId-���Ӳ�����¼ID
'******************************************************************************************************************

Dim mstrPrivs As String, blnPrivacy As Boolean, Elements As New cEPRElements
Dim rs As New ADODB.Recordset, lngKey As Long
Dim strTemp As String, strZipFile As String
    
    gstrSQL = "Select �༭��ʽ From ���Ӳ�����¼ Where ID=[1]"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, lngEPRid)
    
    If NVL(rs!�༭��ʽ, 0) = 1 Then
        dkpMain.FindPane(conPane_RichEpr).Close
        dkpMain.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_���������, lngEPRid, False, 0)
        Call mObjTabEprView.zlRefreshDockfrm
    Else
        dkpMain.FindPane(conPane_TablEpr).Close
        dkpMain.ShowPane conPane_RichEpr
        If SetRichDocsPos(lngEPRid) Then Exit Function
        edtThis.Freeze
        edtThis.ReadOnly = False
        edtThis.ForceEdit = True
        edtThis.NewDoc
        Call ReadRTFFile(lngEPRid)

        If lngEPRid > 0 Then
            '����ҳ���ʽ
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            gstrSQL = "Select c.ID, a.��ʽ From   ����ҳ���ʽ a, �����ļ��б� b, ���Ӳ�����¼ c " & _
                    " Where  c.�ļ�id = b.id And a.���� = b.���� And a.��� = b.ҳ�� And c.ID = [1]"
                    
            If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
            If Not rs.EOF Then
                mEPRFileInfo.��ʽ = zlCommFun.NVL(rs("��ʽ").Value)
                mEPRFileInfo.SetFormat edtThis, mEPRFileInfo.��ʽ
                edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        edtThis.ForceEdit = False
        edtThis.SelStart = 0
        edtThis.UnFreeze
        edtThis.RefreshTargetDC
        edtThis.ViewMode = cprNormal
        edtThis.ReadOnly = True
        Call SetRichDocsPos(lngEPRid)
    End If
    
    OpenEPR = True
    zlCommFun.StopFlash
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    zlCommFun.StopFlash
    Call SaveErrLog
End Function
Private Sub ReadRTFFile(ByVal lngID As Long)
Dim rs As New ADODB.Recordset, strFile As String, strRtf As String, lngLen1 As Long, lngLen2 As Long, lngStart As Long
Dim strIDs As String, varPar() As String, StrKey As String
    On Error GoTo errHand
    gstrSQL = "Select Count(C.Id) As ��Ŀ,f.����, c.Id, c.��������, c.�ļ�id, c.����ʱ��,c.����ID,c.��ҳID, B.ҳ��" & vbNewLine & _
                "From �����ļ��б� F, �����ļ��б� B, ���Ӳ�����¼ C" & vbNewLine & _
                "Where f.���� = b.���� And f.ҳ�� = b.ҳ�� And b.Id = c.�ļ�id And c.Id = [1]" & vbNewLine & _
                "Group By f.����,c.Id, c.��������, c.�ļ�id, c.����ʱ��, c.����id, c.��ҳid, B.ҳ��"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If rs!��Ŀ = 1 Then '����ҳ��ֱ�Ӵ�ӡ
        strFile = zlBlobRead(5, lngID, ,mblnDataMove)
        If gobjFSO.FileExists(strFile) Then
            strRtf = zlFileUnzip(strFile)
            If gobjFSO.FileExists(strRtf) Then
                edtThis.OpenDoc strRtf '���ļ�
                gobjFSO.DeleteFile strRtf, True
            End If
            gobjFSO.DeleteFile strFile, True
        End If
        Call RefreshObject(lngID, edtThis, mblnDataMove)
    Else
        '��ȡ����ҳ����ļ�ID
        strIDs = GetFileRange(rs!�ļ�ID, lngID, Format(rs!����ʱ��, "yyyy-MM-dd HH:mm:ss"), rs!����, rs!����ID, rs!��ҳID, mblnDataMove)
        '��ȡ����ҳ����ļ�ID
        gstrSQL = "Select /*+ rule*/ a.Id, a.�ļ�id, a.��������, a.���汾, a.������, a.���ʱ��, a.����ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ A," & LongIDsTable(strIDs, varPar) & vbNewLine & _
                "Where a.Id = b.Id" & vbNewLine & _
                "Order By a.���, a.����ʱ��"
        If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9))
        gfrmPublic.edtBuff.Freeze
        gfrmPublic.edtBuff.ForceEdit = True
        Do Until rs.EOF
            zlCommFun.ShowFlash "���Դ������ڼ��ص�" & rs.AbsolutePosition & "�ݲ������ݣ�"
            strFile = zlBlobRead(5, rs!ID, ,mblnDataMove)
            If gobjFSO.FileExists(strFile) Then
                strRtf = zlFileUnzip(strFile)
                If gobjFSO.FileExists(strRtf) Then
                    gfrmPublic.edtBuff.OpenDoc strRtf '���ļ�
                    Call RefreshObject(rs!ID, gfrmPublic.edtBuff, mblnDataMove)
                    gobjFSO.DeleteFile strRtf, True
                End If
                gobjFSO.DeleteFile strFile, True
            End If
            
            '��¼�ļ�ID
            StrKey = "FS(" & Format(rs.AbsolutePosition, "00000000") & ",1,0)" & rs!ID & "FE(" & Format(rs.AbsolutePosition, "00000000") & ",1,0)"
            'lngLen2 = Len(edtThis.Text) '���ļ����ӵ����ĵ�ĩβ
            gfrmPublic.edtBuff.Range(0, 0).Selected
            gfrmPublic.edtBuff.Range(0, 0).Text = StrKey
            gfrmPublic.edtBuff.Range(0, 0 + Len(StrKey)).Font.Protected = True
            gfrmPublic.edtBuff.Range(0, 0 + Len(StrKey)).Font.Hidden = True
            
            '׷��RTF�ļ�
            lngLen1 = Len(gfrmPublic.edtBuff.Text) '��¼��ʱ�ļ���ʼ������λ��
            lngLen2 = Len(edtThis.Text) '���ļ����ӵ����ĵ�ĩβ
            edtThis.Range(lngLen2, lngLen2).Selected
            gfrmPublic.edtBuff.SelectAll
            gfrmPublic.edtBuff.CopyWithFormat
            edtThis.PasteWithFormat
            lngStart = Len(edtThis.Text)
            If rs.AbsolutePosition < rs.RecordCount Then
                'ֻҪ�������һ���ļ���ĩβ��֤��һ���س����Ա�׷����һ���ļ�
                If edtThis.Range(lngStart - 2, lngStart) = vbCrLf Then
                    edtThis.Range(lngStart - 2, lngStart).Font.Hidden = False
                Else
                    edtThis.Range(lngStart, lngStart).Text = vbCrLf
                    edtThis.Range(lngStart, lngStart + 2).Font.Hidden = False
                End If
            End If
            edtThis.TOM.TextDocument.Range(lngStart, lngStart).Para = gfrmPublic.edtBuff.TOM.TextDocument.Range(lngLen1, lngLen1).Para '.Duplicate
            rs.MoveNext
        Loop
        gfrmPublic.edtBuff.UnFreeze
        gfrmPublic.edtBuff.ForceEdit = False
        Unload gfrmPublic
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function ShowAdition(ByVal lngRecordId As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ�²��������б���
    '������lngRecordId�����Ӳ�����¼ID
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim objItem As ListItem
    Dim objIcon As StdPicture
    
    Set Me.lvwThis.Icons = Nothing
    Set lvwThis.SmallIcons = Nothing
    lvwThis.ListItems.Clear
    imgThis.ListImages.Clear
    
    gstrSQL = "Select ���, �ļ���, ��С, ������, ���� From ���Ӳ������� Where ����id = [1]"
    If mblnDataMove Then
        gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    With rs
        Do While Not .EOF
            Set objIcon = GetFileIcon(!�ļ���, ICON_SMALL, True)
            imgThis.ListImages.Add , , objIcon
            
            Set lvwThis.Icons = imgThis
            Set lvwThis.SmallIcons = imgThis
            
            Set objItem = lvwThis.ListItems.Add(, "_" & !���, !�ļ��� & "(" & !��С & "KB)")
            objItem.Tag = !�ļ���
            objItem.Icon = imgThis.ListImages.Count
            objItem.SmallIcon = objItem.Icon
            .MoveNext
        Loop
        If lvwThis.ListItems.Count > 0 Then lvwThis.ListItems(1).Selected = True
    End With
    
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function GetFileIcon(ByVal strFile As String, ByVal intSize As ICON_SIZE, Optional blnUntrue As Boolean) As StdPicture
    '******************************************************************************************************************
    '���ܣ�����ָ���ļ��Ĵ�ͼ���Сͼ��
    '˵������Ҫһ��PictureBox�ؼ����ޱ߿�AutoRedraw = True
    '������ strFile��������׺���ļ��������ļ���ʵ�ļ�ʱ��Ӧ�ð���������·����
    '       intSize����ȡͼ��Ĵ�С
    '       blnUntrue������ʵ�ļ�����ʱ��Ҫ�����ļ�����������Ϣ
    '******************************************************************************************************************
    Dim fInfo As SHFILEINFO
    Dim lngRetu As Long
    
    If blnUntrue Then
        strFile = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\" & App.hInstance & CLng(Timer) & strFile
        gobjFSO.CreateTextFile strFile, False
    End If
    If intSize = ICON_LARGE Then
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_LARGEICON)
    Else
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_SMALLICON)
    End If
    If blnUntrue Then gobjFSO.DeleteFile strFile, True
    
    picThis.Width = intSize * Screen.TwipsPerPixelX
    picThis.Height = intSize * Screen.TwipsPerPixelY
    picThis.Cls
    If lngRetu <> 0 Then
        DrawIconEx picThis.hDC, 0, 0, fInfo.hIcon, intSize, intSize, 0, 0, DI_NORMAL
        DestroyIcon fInfo.hIcon
    End If
    Set GetFileIcon = Me.picThis.Image
End Function

Private Sub OpenFile()
    '���ܣ��򿪲��Ÿ���
    Dim strFile As String
    Dim varRetu As Variant, strInfo As String
    
    Screen.MousePointer = vbHourglass
    strFile = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\" & lvwThis.SelectedItem.Tag
    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True
    If zlBlobRead(8, mlngKey & "," & Mid(lvwThis.SelectedItem.Key, 2), strFile) = "" Then
        MsgBox "�ļ���ȡʧ�ܣ���ȷ�ϸ�������Ч�ԣ�", vbInformation, gstrSysName:
        Screen.MousePointer = vbDefault: Exit Sub
    End If
    varRetu = ShellExecute(Me.hWnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If varRetu <= 32 Then
        Select Case varRetu
        Case 2: strInfo = "����Ĺ���"
        Case 29: strInfo = "����ʧ��"
        Case 30: strInfo = "����Ӧ�ó�ʽæµ��..."
        Case 31: strInfo = "û�й����κ�Ӧ�ó�ʽ"
        Case Else: strInfo = "�޷�ʶ��Ĵ���"
        End Select
        MsgBox "�����򿪷�����" & strInfo, vbExclamation, gstrSysName
    End If
    
    Screen.MousePointer = vbDefault
End Sub

'######################################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
        Dim frm As New frmEPRView
        frm.ShowMe Me, mlngKey
    Case ID_EDIT_COPY
        If Control.Enabled And Control.Visible Then '��ݼ�ִ��ʱ��Ҫ�ж�
            Me.edtThis.Copy
        End If
    End Select
End Sub
 
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand
    If Not Me.Visible Then Exit Sub

    Select Case Control.ID
    Case conMenu_File_Open
        Control.Enabled = (mlngKey > 0)
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
        Control.Enabled = edtThis.Selection.getType <> cprSTPicture
        Control.Visible = InStr(gstrPrivsEpr, "���ݸ���") > 0
    End Select

errHand:

End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picPane(0).hWnd
    Case conPane_Annex
        Item.Handle = picPane(1).hWnd
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    End Select
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    'û�����ݸ���Ȩ�޲���������
    If InStr(gstrPrivsEpr, "���ݸ���") = 0 Then Exit Sub
    
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsMain.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub Form_Load()
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    Set Pane1 = dkpMain.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMain.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMain.CreatePane(conPane_Annex, 1200, 15, DockBottomOf, Nothing)
    pane3.MinTrackSize.Height = 0: pane3.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With

    Call InitCommandBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload mfrmPrintPreview
    Set mfrmPrintPreview = Nothing
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Set mfrmMain = Nothing
End Sub

Private Sub lvwThis_DblClick()
    Call OpenFile
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
    '
    RaiseEvent PrintEpr(lngRecordId)
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        edtThis.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 1
        lblThis.Top = (picPane(Index).Height - Me.lblThis.Height) / 2
        lvwThis.Move lvwThis.Left, 15, picPane(Index).Width - lvwThis.Left, picPane(Index).Height - 30
    End Select
    
End Sub

Private Sub RefreshObject(ByVal lngRecordId As Long, ByRef edt As Editor, ByVal blnMoved As Boolean)
'ˢ�½����ϵ�ͼƬ,Ŀǰֻˢ��ͼƬ������Ҫʱ�ٵ���ˢ�±���
Dim Pictures As New cEPRPictures, rsTemp As New ADODB.Recordset, lngKey As Long, Tables As New cEPRTables
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, blnForce As Boolean

    '��ȡ���е�ͼƬ
    gstrSQL = "Select ID, �ļ�id,��ʼ��, ��ֹ��,��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���,Ԥ�����ID " & _
        "From ���Ӳ������� " & _
        "Where �ļ�id = [1] And �������� In (3,5) And ������� Is Not Null" '����ʾ�����е�ͼƬ
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    Do Until rsTemp.EOF
        If rsTemp!�������� = 5 Then
            lngKey = Pictures.Add(NVL(rsTemp!������, 0))
            Call Pictures("K" & lngKey).FillPictureMember(rsTemp, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
            Call Pictures("K" & lngKey).DeleteFromEditor(edt)
            Call Pictures("K" & lngKey).InsertIntoEditor(edt, -1, True)
        ElseIf rsTemp!�������� = 3 Then
            lngKey = Tables.Add(NVL(rsTemp!������, 0))
            Call Tables("K" & lngKey).FillTableMember(rsTemp, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
            
            If Tables("K" & lngKey).Cells.Count = 1 Then
                'һ����Ԫ�񣬿�����PACS�༭����д������
                If FindKey(edt, "T", lngKey, lKSS, lKSE, lKES, lKEE, True) Then
                    '��ɾ��
                    Call Tables("K" & lngKey).DeleteFromEditor(edt)
                    With edt
                        blnForce = .ForceEdit
                        .InProcessing = True
                        .Tag = "TableSingleCell:InsertIntoEditor"
                        .ForceEdit = True
                        .Range(lKSS, lKSS).Font.Protected = False
                        .Range(lKSS, lKSS).Font.Hidden = False
                        .Range(lKSS, lKSS) = Tables("K" & lngKey).Cells(1).�����ı�
                        .ForceEdit = blnForce
                        .UnFreeze
                        .InProcessing = False
                        .Tag = ""
                    End With
                End If
            Else
                '�����Ԫ��
                '��ɾ��
                Call Tables("K" & lngKey).DeleteFromEditor(edt)
                Call Tables("K" & lngKey).InsertIntoEditor(edt, -1)
            End If
        End If
        rsTemp.MoveNext
    Loop
End Sub
Private Sub Open_Pacs_Report(ByVal frmParent As Object, ByVal lngKey As Long, ByVal strReportCode As String, ByVal lngҽ��id As Long, ByVal lng����ID As Long, _
    ByVal lng���ͺ� As Long, ByVal intKind As Integer, ByVal strNo As String, ByVal blnCurrMoved As Boolean, ByVal bytMode As Byte, ByVal strPdfFile As String)
'�Զ��屨����ӡPACS����
        Dim strPicPath As String, strPicFile As String
        Dim cTable As cEPRTable, oPicture As StdPicture
        Dim aryPara(19) As String, intPCount As Integer
        Dim aryFlagPara(1) As String
        Dim intRows As Integer, intCols As Integer
        Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
        Dim i As Integer, rsTemp As New ADODB.Recordset
        
        '��ȡͼ��
        strPicPath = App.Path & "\TmpImage\"
        If gobjFSO.FolderExists(strPicPath) = False Then gobjFSO.CreateFolder strPicPath
        
        '��ȡ����ͼ��(�������ͼ)���ɱ����ļ�,һ����������п������ж������ͼ
        intPCount = 0
        gstrSQL = "Select Id As ����Id From ���Ӳ�������" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
        If mblnDataMove Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If cTable.GetTableFromDB(cprET_���������, lngKey, Val("" & rsTemp!����Id), , IIf(mblnDataMove, "H���Ӳ�������", "���Ӳ�������")) Then
                For i = 1 To cTable.Pictures.Count
                    strPicFile = strPicPath & "PACSPic" & i & ".JPG"
                    If gobjFSO.FileExists(strPicFile) Then gobjFSO.DeleteFile strPicFile, True
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        Set oPicture = cTable.Pictures(i).DrawFinalPic
                    Else
                        Set oPicture = cTable.Pictures(i).OrigPic
                    End If
                    SavePicture oPicture, strPicFile
                    If gobjFSO.FileExists(strPicFile) Then
                        '������ͼ��ͼ���·��
                        If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                            aryFlagPara(0) = strPicFile
                        Else
                            aryPara(intPCount) = strPicFile
                            dcmImages.AddNew
                            dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                            intPCount = intPCount + 1
                            If intPCount > UBound(aryPara) Then Exit Do
                        End If
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
        
        '�ж��Ƿ���Ҫ�Զ����ͼ���Զ��屨����ֻ������һ��ͼ������Զ����ͼ��
        '���²�һ�����ݿ�
        gstrSQL = "Select b.����,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = 1 And b.���� not like '���%'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strReportCode)
        If rsTemp.RecordCount = 1 And intPCount >= 1 Then
            '���ͼ��
            ResizeRegion intPCount, rsTemp("W"), rsTemp("H"), intRows, intCols
            Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
            dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
        End If
        
        '��ȡ�Զ��屨���е�ͼ����
        intPCount = 0
        gstrSQL = "Select b.���� From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = 1" & vbNewLine & _
        "       Order By b.����" 'Trunc(b.y/567),Trunc(b.x/567)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strReportCode)
        Do While Not rsTemp.EOF
            If aryPara(intPCount) = "" Then Exit Do '�����е�ͼ�αȱ����ж�
            '�ֱ�װ�ر��ͼ�ͱ���ͼ��
            If InStr(rsTemp!����, "���") <> 0 Then
                If aryFlagPara(0) <> "" Then aryFlagPara(0) = rsTemp!���� & "=" & aryFlagPara(0)
            Else
                aryPara(intPCount) = rsTemp!���� & "=" & aryPara(intPCount)
                intPCount = intPCount + 1
                If intPCount > UBound(aryPara) Then Exit Do
            End If
            rsTemp.MoveNext
        Loop
        For i = intPCount To UBound(aryPara) '�����е�ͼ�αȱ�������
            If aryPara(i) Like "*=*" Then aryPara(i) = ""
        Next
        
        '���ñ���
       Call ReportOpen(gcnOracle, glngSys, strReportCode, Nothing, _
            "NO=" & strNo, "����=" & intKind, "ҽ��ID=" & lngҽ��id, aryFlagPara(0), _
            aryPara(0), aryPara(1), aryPara(2), aryPara(3), aryPara(4), aryPara(5), _
            aryPara(6), aryPara(7), aryPara(8), aryPara(9), aryPara(10), aryPara(11), _
            aryPara(12), aryPara(13), aryPara(14), aryPara(15), aryPara(16), aryPara(17), _
            aryPara(18), aryPara(19), "PDF=" & strPdfFile, IIf(strPdfFile = "", bytMode, 4))
End Sub
Private Sub Open_LIS_Report(ByVal frmParent As Object, ByVal lngKey As Long, ByVal strReportCode As String, ByVal lngҽ��id As Long, ByVal lng����ID As Long, _
    ByVal lng���ͺ� As Long, ByVal intKind As Integer, ByVal strNo As String, ByVal blnCurrMoved As Boolean, ByVal bytMode As Byte, ByVal strPdfFile As String)
'����LiwWork��ӡ��ͼ�ε�LIS����
'bytMode=2ֱ�Ӵ�ӡ =1 Ԥ��
    Dim strChart(0 To 8) As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer
    Dim objLisWork As Object
    Dim lng�걾id As Long, lngAdviceID As Long
                    
    On Error GoTo ErrHandle
    
    strSQL = "Select a.�걾id ID, b.ҽ��id" & vbNewLine & _
                "From ������Ŀ�ֲ� A, ����걾��¼ B" & vbNewLine & _
                "Where a.ҽ��id = [1] And a.�걾id = b.Id And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�걾ID", lngҽ��id)
    If Not rsTmp.EOF Then
        lng�걾id = NVL(rsTmp!ID, 0)
        lngAdviceID = NVL(rsTmp!ҽ��id, 0)
    Else
        Call ReportOpen(gcnOracle, glngSys, strReportCode, frmParent, "NO=" & CStr(strNo), "����=" & Val(intKind), "ҽ��ID=" & lngҽ��id, _
                    "����ID=" & lng����ID, "���ͺ�=" & lng���ͺ�, "PDF=" & strPdfFile, IIf(strPdfFile = "", bytMode, 4))
        Exit Sub
    End If
    
    Set objLisWork = CreateObject("zl9LisWork.clsLISImg")
    strSQL = "select id from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lng�걾id)
    intLoop = 0
    Do Until rsTmp.EOF
        If Not objLisWork Is Nothing Then
            If objLisWork.Get_Chart2d_File(App.Path, rsTmp("ID")) Then
                strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
            End If
        End If
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    Call ReportOpen(gcnOracle, glngSys, strReportCode, frmParent, "NO=" & CStr(strNo), "����=" & Val(intKind), "ҽ��ID=" & lngAdviceID, _
                    "����ID=" & lng����ID, "�걾ID=" & lng�걾id, _
                    "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), "ͼ��4=" & strChart(3), _
                    "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                    "ͼ��9=" & strChart(8), "PDF=" & strPdfFile, IIf(strPdfFile = "", bytMode, 4))
    'ɾ��ͼ���ļ�
    For intLoop = 0 To 8
        If strChart(intLoop) <> "" Then
            If gobjFSO.FileExists(strChart(intLoop)) Then
                gobjFSO.DeleteFile strChart(intLoop), True
            End If
        End If
    Next
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
