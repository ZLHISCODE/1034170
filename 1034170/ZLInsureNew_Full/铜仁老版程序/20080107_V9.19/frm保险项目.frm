VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm������Ŀ 
   BackColor       =   &H8000000A&
   Caption         =   "ҽ����Ŀ����"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   10110
   Icon            =   "frm������Ŀ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ZL9BillEdit.BillEdit mshSum_S 
      Height          =   2745
      Left            =   3090
      TabIndex        =   4
      Top             =   1020
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4842
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
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2670
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":0E42
            Key             =   "R"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":115C
            Key             =   "C"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":12B6
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   90
      TabIndex        =   7
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   3450
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":1708
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":1924
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":1B40
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":1D5A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":1F76
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   2760
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":2192
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":23AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":25CA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":27E4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ.frx":2A00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1270
      BandCount       =   2
      _CBWidth        =   10110
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "�������"
      Child2          =   "cmb����"
      MinHeight2      =   300
      Width2          =   2325
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   3675
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   635
      SimpleText      =   $"frm������Ŀ.frx":2C1C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm������Ŀ.frx":2C63
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12779
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "��ԭ(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6750
      TabIndex        =   6
      Top             =   4050
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5340
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Visible         =   0   'False
      Begin VB.Menu mnuEditGet 
         Caption         =   "������ȡ��Ŀ�����Ϣ(&G)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolSplit 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "���༭��Ŀ����(&I)"
      End
      Begin VB.Menu mnuViewClass 
         Caption         =   "���༭ҽ������(&C)"
      End
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R) "
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)��"
      End
   End
End
Attribute VB_Name = "frm������Ŀ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private int��˱�־ As Integer
Private classInsure As New clsInsure

Private Enum ColumnEnum
    cOL���� = 0
    cOL���� = 1
    col���� = 2
    COL��� = 3
    COL���� = 4
    COL��λ = 5
    col�۸� = 6
    col�ı䷽ʽ = 7
    col����ID = 8
    COLҽ������ = 9
    colҽ������ = 10
    colҽ������ = 11
    COLҽ����ע = 12
    colԭ���� = 13
    col�������� = 14
    col��ҽ�� = 15
    'Modified By ���� ��������ɳ ԭ��û����ֻ�м���
    colƥ�����к� = 16
    col��˱�־ = 17
End Enum
Private Const mlng���볤�� As Long = 20

Dim mlngListIndex As Long   '�����ϴ��������ѡ������
Dim mblnLoad As Boolean
Dim msngStartX As Single    '�ƶ�ǰ����λ��
Dim mstrȨ�� As String

Dim mstrKey As String       'ǰһ�����ڵ�Ĺؼ�ֵ
Dim mint���� As Integer     '��ǰ��ʾ������
Dim mint���õ��� As Integer '����ר�ã�0��ʾ����������1��ʾ����������ɾ������˵���Ŀ��

Dim mlngCol As Long, mblnDesc As Boolean
Private mcnYB As New ADODB.Connection   'ҽ��ǰ�÷���������

Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub cmdRestore_Click()
    'Modified By ���� ��������ɳ
    If mint���� = TYPE_������ Then
        MsgBox "��ҽ����֧��ȡ�����ܣ��������棡", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("��ȷ��Ҫ�����޸���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    Call FillSum(True)
    mshSum_S.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim lngRow As Long
    
    
    gcnOracle.BeginTrans
    If mint���� = TYPE_������ Then gcn����.BeginTrans
    On Error GoTo errHandle
    
    With mshSum_S
        '��������
        For lngRow = 1 To .Rows - 1
            Select Case .TextMatrix(lngRow, col�ı䷽ʽ)
                Case "����", "�޸�"
                    '���������޸ķ���һ�������д���
'                    �շ�ϸĿID,����,����ID,��Ŀ����,��Ŀ����,��ע
                    'Modified by ZYB 2004-08-17
                    If gintInsure = TYPE_��ɽ Then
                        gstrSQL = "ZL_����֧����Ŀ_Modify(" & .RowData(lngRow) & "," & mint���� & "," & _
                                   IIf(Val(.TextMatrix(lngRow, col����ID)) = 0, "null", .TextMatrix(lngRow, col����ID)) & ",'" & _
                                   .TextMatrix(lngRow, COLҽ������) & "','" & Split(.TextMatrix(lngRow, colҽ������), "-")(1) & "','" & .TextMatrix(lngRow, COLҽ����ע) & _
                                   IIf(mint���� = TYPE_������, "^^" & .TextMatrix(lngRow, colƥ�����к�) & "||" & _
                                   IIf(Trim(.TextMatrix(lngRow, col��˱�־)) = "��", 1, IIf(Trim(.TextMatrix(lngRow, col��˱�־)) = "��", 2, 0)), "") & _
                                   "'," & IIf(Trim(.TextMatrix(lngRow, col��ҽ��)) = "��", 0, 1) & ")"
                    Else
                        gstrSQL = "ZL_����֧����Ŀ_Modify(" & .RowData(lngRow) & "," & mint���� & "," & _
                                   IIf(Val(.TextMatrix(lngRow, col����ID)) = 0, "null", .TextMatrix(lngRow, col����ID)) & ",'" & _
                                   .TextMatrix(lngRow, COLҽ������) & "','" & .TextMatrix(lngRow, colҽ������) & "','" & .TextMatrix(lngRow, COLҽ����ע) & _
                                   IIf(mint���� = TYPE_������, "^^" & .TextMatrix(lngRow, colƥ�����к�) & "||" & _
                                   IIf(Trim(.TextMatrix(lngRow, col��˱�־)) = "��", 1, IIf(Trim(.TextMatrix(lngRow, col��˱�־)) = "��", 2, 0)), "") & _
                                   "'," & IIf(Trim(.TextMatrix(lngRow, col��ҽ��)) = "��", 0, 1) & ")"
                    End If
                    Call DebugTool("׼�����汾���޸�")
                    Call ExecuteProcedure(Me.Caption)
                    Call DebugTool("�޸ĳɹ�")
                    
                    gstrSQL = ""
                    If .TextMatrix(lngRow, COLҽ������) <> .TextMatrix(lngRow, colԭ����) Then
                        '�����޸ļ�¼
                        gstrSQL = "Insert Into ��Ŀ��Ӧ��־(����ҩ�����,����ҩ������,ҽԺҩ������,�޸���,��������) " & _
                        "values('" & .TextMatrix(lngRow, COLҽ������) & "','" & .TextMatrix(lngRow, colҽ������) & "','" & .TextMatrix(lngRow, cOL����) & "','" & gstrUserName & "',sysdate)"
                    End If
                    
                    Call DebugTool("������Ŀ�޸���־:" & gstrSQL)
                    If gstrSQL <> "" Then
                        Select Case mint����
                        Case TYPE_������
                            gcn����.Execute gstrSQL
                        Case TYPE_������
                            gcnOracle.Execute gstrSQL
                        Case TYPE_����������
                            mcnYB.Execute gstrSQL
                        End Select
                    End If
                    Call DebugTool("�޸���־����ɹ���")
                    
                    .TextMatrix(lngRow, colԭ����) = .TextMatrix(lngRow, COLҽ������)
                Case "ɾ��"
                    'ɾ������Ŀ
                    If .TextMatrix(lngRow, colԭ����) <> "" Then
                        gstrSQL = "Insert Into ��Ŀ��Ӧ��־(����ҩ�����,����ҩ������,ҽԺҩ������,�޸���,��������) " & _
                        "values('000000','��ҽ����Ŀ','" & .TextMatrix(lngRow, cOL����) & "','" & gstrUserName & "',sysdate)"
                    End If
                    Select Case mint����
                    Case TYPE_������
                        gcn����.Execute gstrSQL
                    Case TYPE_������
                        gcnOracle.Execute gstrSQL
                    Case TYPE_����������
                        mcnYB.Execute gstrSQL
                    End Select
                    
                    gstrSQL = "ZL_����֧����Ŀ_Delete(" & .RowData(lngRow) & "," & mint���� & ")"
                    Call ExecuteProcedure(Me.Caption)
                    .TextMatrix(lngRow, colԭ����) = .TextMatrix(lngRow, COLҽ������)
            End Select
        Next
        
        '�����ݴ���������������������״̬
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, col�ı䷽ʽ) = ""
        Next
    End With
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    gcnOracle.CommitTrans
    If mint���� = TYPE_������ Then gcn����.CommitTrans
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
    If mint���� = TYPE_������ Then gcn����.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call FillTree
    End If
    
    Call mshSum_S_EnterCell(1, cOL����)
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mstrKey = ""
    mlngCol = 0
    mblnDesc = False
    mblnLoad = True
    
    gstrSQL = "select ���,���� from ������� where nvl(�Ƿ��ֹ,0)<>1  order by ���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With cmb����
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("���")
            If rsTemp("���") = gintInsure Then
                '��ǰҽ����
                'ʹ��API�����Բ�����Click�¼�
                zlControl.CboSetIndex .hwnd, .NewIndex
                Call Fill����
            End If
            
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            'ʹ��API�����Բ�����Click�¼�
            zlControl.CboSetIndex .hwnd, 0
            Call Fill����
        End If
    End With
    mint���� = cmb����.ItemData(cmb����.ListIndex)
    
    Call InitSum
    RestoreWinState Me, App.ProductName
    
    mnuViewItem.Checked = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", "False") <> "False"
    If mnuViewItem.Checked = False Then
        '�����жϴ�����
        mnuViewClass.Checked = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", "False") <> "False"
    End If
    Call SetSkip
    
    zlControl.CboSetHeight cmb����, 3600
    '�õ���ѯ��ʱ�䷶Χ
    If mint���� = TYPE_������ Then
        mint���õ��� = 0
        gstrSQL = "Select ����ֵ From ���ղ��� Where ������='���õ���'"
        Call OpenRecordset(rsTemp, "��ȡ���õ���")
        If Not rsTemp.EOF Then
            mint���õ��� = Nvl(rsTemp!����ֵ, 0)
        End If
        mnuEdit.Visible = True
    End If
End Sub

Private Sub InitSum()
'��ʼ�����ܱ����ʽ
    Dim lngCol As Long
    
    With mshSum_S
        ClearGrid mshSum_S
        
        'Modified By ���� ��������ɳ ԭ�������С���ƥ�����к�
        .Cols = 18
        
        .TextMatrix(0, cOL����) = "����"
        .TextMatrix(0, cOL����) = "�շ�ϸĿ"
        .TextMatrix(0, COL���) = "���"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, COL��λ) = "��λ"
        If mint���� = TYPE_�¶� Then
            .TextMatrix(0, col�۸�) = "�Ը�����"
        Else
            .TextMatrix(0, col�۸�) = "�۸�"
        End If
        .TextMatrix(0, col�ı䷽ʽ) = "�Ƿ��޸�"
        If mint���� = type_���������� Or mint���� = type_������ Then
            .TextMatrix(0, COLҽ������) = "���"
        Else
            .TextMatrix(0, COLҽ������) = "ҽ����Ŀ����"
        End If
        If mint���� = type_������ Then
            .TextMatrix(0, colҽ������) = "��ҵ���ѱ���"
        Else
            .TextMatrix(0, colҽ������) = "ҽ����Ŀ����"
        End If
        .TextMatrix(0, COL����) = "����"
        .TextMatrix(0, colҽ������) = "����"
        .TextMatrix(0, col��˱�־) = "���"
        .TextMatrix(0, COLҽ����ע) = "ҽ����Ŀ��ע"
        .TextMatrix(0, colԭ����) = "ԭҽ����Ŀ����"
        .TextMatrix(0, col����ID) = "����ID"
        .TextMatrix(0, col��������) = "ҽ����������"
        
        If mint���� = TYPE_ǭ�� Then
            .TextMatrix(0, col��ҽ��) = "�������"
        Else
            .TextMatrix(0, col��ҽ��) = "�Ƿ�ҽ��"
        End If
        
        .TextMatrix(0, colƥ�����к�) = "ƥ�����к�"
        
        .ColWidth(cOL����) = 1000
        .ColWidth(cOL����) = 2000
        .ColWidth(COL���) = 1000
        .ColWidth(col����) = 600
        .ColWidth(COL��λ) = 600
        .ColWidth(col�۸�) = 800
        .ColWidth(col�ı䷽ʽ) = 0
        .ColWidth(COLҽ������) = 1200
        .ColWidth(colҽ������) = 1200
        .ColWidth(COLҽ����ע) = 0
        .ColWidth(colԭ����) = 0
        .ColWidth(col����ID) = 0
        .ColWidth(col��������) = 1200
        .ColWidth(col��ҽ��) = 800
        .ColWidth(colƥ�����к�) = 0
        
        If mint���� = TYPE_������ Then
            .ColWidth(COL����) = 700
            .ColWidth(colҽ������) = 700
            .ColWidth(col��˱�־) = 400
        Else
            .ColWidth(COL����) = 0
            .ColWidth(colҽ������) = 0
            .ColWidth(col��˱�־) = 0
        End If
        
        For lngCol = 0 To .Cols - 1
            .ColAlignment(lngCol) = 1
        Next
        .ColAlignment(col�۸�) = 7
        .ColAlignment(col��ҽ��) = 4
        
        '���ø��еı༭����
        .ColData(COL����) = 5
        .ColData(colҽ������) = 5
        .ColData(col��˱�־) = 5
        .ColData(cOL����) = 5 '����ѡ��
        .ColData(cOL����) = 5
        .ColData(COL���) = 5
        .ColData(col����) = 5
        .ColData(COL��λ) = 5
        .ColData(col�۸�) = 5
        .ColData(col�ı䷽ʽ) = 5
        If mint���� = type_���������� Or mint���� = type_������ Then
            .ColData(COLҽ������) = 3
        Else
            .ColData(COLҽ������) = 1
        End If
        If mint���� = type_������ Then
            .ColData(colҽ������) = 4
        Else
            .ColData(colҽ������) = 5
        End If
        .ColData(COLҽ����ע) = 5
        .ColData(colԭ����) = 5
        .ColData(col����ID) = 5
        .ColData(col��������) = 3 'ѡ����
        .ColData(col��ҽ��) = -1 'ѡ����
        .ColData(colƥ�����к�) = 5
        
        .PrimaryCol = cOL����
        
        If gintInsure = TYPE_�ɶ��ϳ� Then
            .TxtCheck = True
            .TextMask = "`"
        End If
                
        Call SetSkip
        .AllowAddRow = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSave.Enabled = True Then
        MsgBox "ҽ����Ŀ�б������ڱ༭״̬�������˳�����", vbInformation, gstrSysName
        Cancel = 1
        Exit Sub
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", mnuViewItem.Checked
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", mnuViewClass.Checked
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbrThis.Height)
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrHeight, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    '�ұ�
    'tvwMain_S��λ��
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    'picV��λ��
    picV.Top = sngTop
    picV.Height = tvwMain_S.Height
    picV.Left = tvwMain_S.Left + tvwMain_S.Width
    
    cmdRestore.Top = sngBottom - cmdRestore.Height - 100
    cmdRestore.Left = ScaleWidth - cmdRestore.Width - 300
    cmdSave.Top = cmdRestore.Top
    cmdSave.Left = cmdRestore.Left - cmdSave.Width - 300
    
    If InStr(mstrȨ��, "��ɾ��") > 0 Then
        '���Ա༭
        sngBottom = cmdRestore.Top - 100
    End If
    
    mshSum_S.Left = picV.Left + picV.Width
    If ScaleWidth - mshSum_S.Left > 0 Then mshSum_S.Width = ScaleWidth - mshSum_S.Left
    mshSum_S.Top = sngTop
    mshSum_S.Height = IIf(sngBottom - mshSum_S.Top > 0, sngBottom - mshSum_S.Top, 0)
    
    Refresh
End Sub

Private Function GetMatch(ByVal rsMatch As ADODB.Recordset, ByVal intType As Integer) As Boolean
    Dim str���� As String, strƥ�����к� As String, strTmp As String, strƥ������ As String
    Dim int��˱�־ As Integer
    '������ȡҽ�����ĵ�ƥ����Ϣ�������±������ݿ�
    'intType=0��������Ŀ;1��ҩƷ��Ŀ
    
    'ȡҩƷ��ƥ����Ϣ
    If Not ���ýӿ�_׼��_������(Function_������.��Ŀƥ��_ȡƥ����Ŀ��Ϣ) Then Exit Function
    gstrField_������ = "hospital_id||audit_status||item_type"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||1||" & intType
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������() Then Exit Function
    If Not ���ýӿ�_ָ����¼��_������("ItemMatch") Then Exit Function
'    ���    �ֶ�    �ֶ�˵��    ��󳤶�    ��ע
'    1   hosp_code   ҽԺĿ¼����    20
'    2   hosp_name   ҽԺĿ¼����    60
'    3   hosp_model  ҽԺĿ¼����    20
'    4   item_name   ����Ŀ¼����    60
'    5   model_name  ����Ŀ¼����    20
'    6   serial_match    ƥ�����к�  12
'    7   valid_flag  ��Ч��־    1   "0"����Ч    "1"����Ч
'    8   audit_flag  ��˱�־    1   "0"��δ���    "1"�����ͨ��    "2"�����δͨ��
'    9   match_type  ƥ������    1   "0"��������Ŀƥ��    "1"����ҩƥ��    "2"���г�ҩƥ��    "3"���в�ҩƥ��
    If ���ýӿ�_��¼��_������ Then
        Do While True
            Call ���ýӿ�_��ȡ����_������("hosp_code", str����)
            Call ���ýӿ�_��ȡ����_������("serial_match", strƥ�����к�)
            Call ���ýӿ�_��ȡ����_������("match_type", strƥ������)
            Call ���ýӿ�_��ȡ����_������("audit_flag", strTmp)
            int��˱�־ = Val(strTmp)
            
            '��λ�ü�¼���ҳ��շ�ϸĿID
            rsMatch.Filter = "����='" & str���� & "'"
            
            If Not rsMatch.EOF Then
                '���±���֧����Ŀ
                gstrSQL = "ZL_����֧����Ŀ_Modify(" & rsMatch!�շ�ϸĿID & "," & TYPE_������ & "," & rsMatch!����ID & ",'" & _
                           rsMatch!��Ŀ���� & "','" & rsMatch!��Ŀ���� & "','" & Split(rsMatch!��ע, "^^")(0) & "^^" & strƥ�����к� & "||" & int��˱�־ & _
                           "'," & rsMatch!�Ƿ�ҽ�� & ")"
                Call ExecuteProcedure("���±���֧����Ŀ")
            Else
                MsgBox "�ӿڷ��ص�ҽԺ������ʶ��[" & str���� & "]�����ڱ��ر���֧����Ŀ�У�δ�ҵ����շ�ϸĿ", vbInformation, gstrSysName
            End If
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    MsgBox "�Ѵ����ĳɹ���ȡ������Ŀ��ƥ����Ϣ��", vbInformation, gstrSysName
    GetMatch = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub mnuEditGet_Click()
    Dim rsMatch As New ADODB.Recordset
    On Error GoTo ErrHand
    If MsgBox("����������ܻỨ�ܳ�ʱ�䣬��ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = " (Select ID �շ�ϸĿID,Decode(TRIM(��ʶ����),NULL,����,'',����,��ʶ����) ���� From �շ�ϸĿ Where ��� Not In ('5','6','7')" & _
              " Union " & _
              " Select ҩƷID �շ�ϸĿID,Decode(Trim(��ʶ��),NULL,����,'',����,��ʶ��) ���� From ҩƷĿ¼)"
    gstrSQL = " Select B.����,A.�շ�ϸĿID,A.����ID,A.��Ŀ����,A.��Ŀ����,A.��ע,A.�Ƿ�ҽ�� " & _
              " From ����֧����Ŀ A," & gstrSQL & " B" & _
              " Where A.�շ�ϸĿID=B.�շ�ϸĿID And A.����=" & TYPE_������
    Call OpenRecordset(rsMatch, "ȡ����֧����Ŀ")
    
    If Not classInsure.InitInsure(gcnOracle, TYPE_������) Then Exit Sub
    gcnOracle.BeginTrans
    
    rsMatch.Filter = 0
    If Not GetMatch(rsMatch, 0) Then
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    rsMatch.Filter = 0
    If Not GetMatch(rsMatch, 1) Then
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    gcnOracle.CommitTrans
    
    '������ʾ��ҳ����Ϣ
    Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuViewFind_Click()
    If cmdSave.Enabled = True Then
        MsgBox "ҽ����Ŀ�б������ڱ༭״̬������ʹ�ò��ҹ��ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    frm������Ŀ����.Show vbModal, Me
End Sub

Private Sub cmb����_Click()
    Call Fill����
    Call FillSum(False)
    
    'Modified By ���� ��������ɳ ԭ�򣺳�ʼ��ҽ���ӿ�
    If cmb����.ItemData(cmb����.ListIndex) <> TYPE_������ Then Exit Sub
    Call classInsure.InitInsure(gcnOracle, TYPE_������)
End Sub

Private Sub mnuViewClass_Click()
    mnuViewItem.Checked = False
    mnuViewClass.Checked = Not mnuViewClass.Checked
    
    Call SetSkip
End Sub

Private Sub mnuViewItem_Click()
    mnuViewClass.Checked = False
    mnuViewItem.Checked = Not mnuViewItem.Checked
    
    Call SetSkip
End Sub

Private Sub SetSkip()
'���ñ�����Ծ����
    With mshSum_S
        If mnuViewItem.Checked = False Then
        
            If mint���� = type_���������� Or mint���� = type_������ Then
            Else
                .ColData(COLҽ������) = 1
            End If
            .LocateCol = COLҽ������
            
            .ColData(col��������) = IIf(mnuViewClass.Checked = True, 5, 3)
        Else
            .ColData(col��������) = 3 'ѡ����
            .LocateCol = col��������
            If mint���� = type_���������� Or mint���� = type_������ Then
            Else
                .ColData(COLҽ������) = 5
            End If
        End If
        If .ColData(.Col) = 5 Then
            '��ǰ���Ѿ�����ѡ�������¶�λ
            .Col = .LocateCol
        End If
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    'ֻˢ���б�����
    Call FillSum
End Sub

Private Sub mshSum_S_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    'ʼ���ǲ�����ɾ����
    Cancel = True
    
    With mshSum_S
        'Modified By ���� ��������ɳ ԭ����ͨ��ҽ��������˵���Ŀ������ɾ��
        If mint���� = TYPE_������ Then
            Call GetItemMatchInfo
            If int��˱�־ = 1 And mint���õ��� = 0 Then
                MsgBox "����Ŀ�Ѿ�ͨ��ҽ��������ˣ�������ɾ��������ҽ��������ϵ��", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
        End If
        
        If .TextMatrix(Row, col�ı䷽ʽ) = "����" Then
            .TextMatrix(Row, col�ı䷽ʽ) = "" '�൱��ʲô��û����
            'Modified By ���� ��������ɳ ԭ�򣺸��ݵ�ǰ����������Ŀƥ����Ϣ
            Call SetItemMatch
        Else
            .TextMatrix(Row, col�ı䷽ʽ) = "ɾ��" '���
            'Modified By ���� ��������ɳ ԭ�򣺸��ݵ�ǰ����������Ŀƥ����Ϣ
            Call SetItemMatch
        End If
        
        .TextMatrix(Row, COLҽ������) = ""
        .TextMatrix(Row, colҽ������) = ""
        .TextMatrix(Row, colҽ������) = ""
        .TextMatrix(Row, COLҽ����ע) = ""
        .TextMatrix(Row, col����ID) = ""
        .TextMatrix(Row, col��������) = ""
        .TextMatrix(Row, col��ҽ��) = ""
        .TextMatrix(Row, col��˱�־) = ""
    End With
    cmdSave.Enabled = True
    cmdRestore.Enabled = True
End Sub

Private Sub mshSum_S_cboClick(ListIndex As Long)
    With mshSum_S
        If mint���� = type_������ Or type_���������� = mint���� Then
            If .Col = COLҽ������ Then
                mlngListIndex = ListIndex
                If .TextMatrix(.Row, COLҽ������) <> .CboText Then
                    .TextMatrix(.Row, COLҽ������) = .CboText
                    Call ��Ǹı�
                End If
                Exit Sub
            End If
        End If
        If .Col <> col�������� Then Exit Sub
        
        If .TextMatrix(.Row, col��������) <> .CboText Then
            '��ֹ�޸ı��մ���,ֻ����ͨ��ѡ����ϸ��ȷ������
            If gintInsure = TYPE_������ Then
                .ListIndex = mlngListIndex
                Exit Sub
            End If
            mlngListIndex = ListIndex
            .TextMatrix(.Row, col��������) = .CboText
            Call ��Ǹı�
        Else
            mlngListIndex = ListIndex
        End If
        
        If .CboText = "" Then
            '����Ϊ��
            .TextMatrix(.Row, col����ID) = ""
            .TextMatrix(.Row, col��������) = ""
        Else
            .TextMatrix(.Row, col����ID) = .ItemData(.ListIndex)
            .TextMatrix(.Row, col��������) = .CboText
        End If
        
    End With
End Sub

Private Sub mshSum_S_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshSum_S
        If KeyCode = vbKeyReturn Then
              '���˺�(200311)
            If mint���� = type_���������� Or mint���� = type_������ Then
                If .Col = COLҽ������ Then
                    If .CboText = "" Then
                            .TextMatrix(.Row, COLҽ������) = " "
                            If mint���� = type_������ Then
                                .Col = colҽ������
                            Else
                                .Col = col��������
                            End If
                    Else
                            .TextMatrix(.Row, COLҽ������) = .CboText
                    End If
                    Call ��Ǹı�
                    Exit Sub
                End If
            End If
            If .TextMatrix(.Row, col��������) <> .CboText Then
                .TextMatrix(.Row, col��������) = .CboText
                Call ��Ǹı�
            End If
            
            If .CboText = "" Then
                '����Ϊ��
                .TextMatrix(.Row, col����ID) = ""
                .TextMatrix(.Row, col��������) = ""
                .Col = col��ҽ��
            Else
                .TextMatrix(.Row, col����ID) = .ItemData(.ListIndex)
                .TextMatrix(.Row, col��������) = .CboText
            End If
        End If
    End With
End Sub

Private Sub mshSum_S_CommandClick()
'���ܣ���ȡҽ����Ŀ��ѡ��
'��������
'���أ�ҽ����Ŀ����
    Dim strCode As String
    Dim strSelected As String
    Dim strName As String
    Dim strlastCode As String
    Dim strMemo As String
    
    With mshSum_S
        strCode = .TextMatrix(.Row, COLҽ������)
        Select Case mint����
            Case TYPE_������, TYPE_����������
                On Error Resume Next
                If frm������Ŀѡ������.GetCode(strCode, strName, Val(.TextMatrix(.Row, col�۸�)), mint����) = True Then
                    strSelected = strCode
                End If
            Case TYPE_����
                If frm������Ŀѡ������.GetCode(strCode, strName, mint����) = True Then
                    strSelected = strCode
                End If
            Case TYPE_�山ũҽ
                If frm������Ŀѡ��_�山ũҽ.GetCode(strCode, strName, mint����) = True Then
                    strSelected = strCode
                End If
            Case TYPE_�㽭
                On Error Resume Next
                If frm������Ŀѡ���㽭.GetCode(strCode, strName, strMemo, Val(.TextMatrix(.Row, col�۸�)), mint����) = True Then
                    strSelected = strCode
                End If
            Case TYPE_��Ҧ
                On Error Resume Next
                If frm������Ŀѡ����Ҧ.GetCode(strCode, strName, Val(.TextMatrix(.Row, col�۸�)), mint����) = True Then
                    strSelected = strCode
                End If
            Case TYPE_�¶�
                On Error Resume Next
                If frm������Ŀѡ���¶�.GetCode(strCode, strName, Val(.TextMatrix(.Row, col�۸�)), mint����) = True Then
                    strSelected = strCode
                End If
                
            Case TYPE_�����山
                '���˺�:20040706
                On Error Resume Next
                If frm������Ŀѡ�������山.GetCode(Me, strCode, strName) = True Then
                    strSelected = Mid(strCode, 2)
                    .TextMatrix(.Row, COLҽ����ע) = Mid(strCode, 1, 1)
                End If
            Case TYPE_ǭ��
                On Error Resume Next
                If frm������Ŀѡ��ǭ��.GetCode(Me, strCode, strName) = True Then
                    strSelected = strCode
                End If
            Case TYPE_�ɶ�����
                'û���ṩ��ȡ����;��
            Case TYPE_�ɶ��ϳ�
                If frm������Ŀѡ���ϳ�.GetCode(strCode, strName) Then
                    strSelected = strCode
                End If
            Case TYPE_����
                strName = .TextMatrix(.Row, colҽ������)
                If frm������Ŀѡ�񱱾�.GetCode(strCode, strName, TYPE_����) = False Then Exit Sub
                strSelected = strCode
                '�����ҩƷ��Ŀ�������Ʒ���ͱ����Ƿ���ҽ�������·���ҩƷ�����У�����ǲ��������ö���
                If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                    If Not CheckTradeName(.RowData(.Row), strCode) Then
                        Exit Sub
                    End If
                End If
            'Modified by ZYB �Ͻ�
            Case TYPE_�Ͻ�
                If frm������Ŀѡ��Ͻ�.GetCode(strCode, strName, mint����) = True Then
                    strSelected = strCode
                End If
            Case Else
                If mint���� = TYPE_������ Then
                    Call GetItemMatchInfo
                    If int��˱�־ = 1 And mint���õ��� = 0 Then
                        MsgBox "����Ŀ�Ѿ�ͨ����ˣ��������޸ģ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                If frm������Ŀѡ��.GetCode(strCode, mint����) = True Then
                    strSelected = strCode
                    If mint���� = TYPE_������ Then
                        Call CheckValid(strCode)
                    End If
                End If
        End Select
        
        If strSelected <> "" Then
            If mint���� = TYPE_ǭ�� Then
                .TextMatrix(.Row, COLҽ������) = Mid(strSelected, 2)
                .TextMatrix(.Row, COLҽ����ע) = Mid(strSelected, 1, 1)
            Else
                .TextMatrix(.Row, COLҽ������) = strSelected
            End If
            If strName = "" Or mint���� = TYPE_���������� Or mint���� = TYPE_�����山 Or mint���� = TYPE_�Ͻ� Or mint���� = TYPE_ǭ�� Then
                Call Get��������
            Else
                '�Ѿ��������ƣ��Ͳ����ٵ���
                .TextMatrix(.Row, colҽ������) = strName
                If mint���� = TYPE_�㽭 Then
                    .TextMatrix(.Row, COLҽ����ע) = strMemo
                Else
                    .TextMatrix(.Row, COLҽ����ע) = ""
                End If
                .TextMatrix(.Row, col��ҽ��) = ""
            End If
            Call ��Ǹı�
            'Modified By ���� ��������ɳ ԭ�򣺸��ݵ�ǰ����������Ŀƥ����Ϣ
            If mint���� = TYPE_������ Then
                .TextMatrix(.Row, colҽ������) = Split(.TextMatrix(.Row, COLҽ����ע), "||")(3)
            End If
            Call SetItemMatch(False)
        End If
    End With
End Sub

Private Sub mshSum_S_DblClick(Cancel As Boolean)
    With mshSum_S
        If .Active = False Then Exit Sub
        Call ��Ǹı�
    End With
End Sub

Private Sub mshSum_S_EnterCell(Row As Long, Col As Long)
    Static lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    
    If Col = col�������� And Trim(mshSum_S.TextMatrix(Row, Col)) = "" Then
        mshSum_S.ListIndex = -1
    End If
    
    '���˺�(200311)
    If type_���������� = mint���� Or type_������ = mint���� Then
        Select Case Col
        Case COLҽ������
            mshSum_S.Clear
            mshSum_S.AddItem ""
            mshSum_S.AddItem "���"
            mshSum_S.AddItem "����"
        Case colҽ������
            If type_������ = mint���� Then
                mshSum_S.TxtCheck = True
                mshSum_S.MaxLength = 11
                mshSum_S.TextMask = ".1234567890"
            End If
        Case col��������
            Dim strSql As String
              gstrSQL = "select ID,����,���� from ����֧������ " & _
              "where ����=" & cmb����.ItemData(cmb����.ListIndex) & " order by ����"
            Call OpenRecordset(rsTemp, Me.Caption)
            mshSum_S.Clear
            Do Until rsTemp.EOF
                mshSum_S.AddItem rsTemp("����") & "." & rsTemp("����")
                mshSum_S.ItemData(mshSum_S.NewIndex) = rsTemp("ID")
                rsTemp.MoveNext
            Loop
            
        End Select
    End If
    'Modified By ���� ��������ɳ ԭ�򣺻�ȡ��Ŀƥ����Ϣ
    If mint���� <> TYPE_������ Then Exit Sub
    If lngRow = Row Then Exit Sub
    lngRow = Row
    Call GetItemMatchInfo
End Sub

Private Sub mshSum_S_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '������Ŀ����
    Dim strǰ As String, strText As String, str���� As String
    Dim rsTemp As New ADODB.Recordset, blnReturn As Boolean
    Dim strLeft As String
    Dim strTemp As String

    strǰ = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0") = "0", "%", "") '˫��ƥ��
    
    On Error GoTo errHandle
    
    With mshSum_S
        If mint���� = type_������ And .Col = colҽ������ And KeyCode = vbKeyReturn Then
            strText = Replace(Trim(.Text), "`", "")
            If Not IsNumeric(strText) And strText <> "" Then
                ShowMsgbox "��ҵ���ѱ�������Ϊ������,�����䣡"
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Val(strText) > 100 Then
                ShowMsgbox "��ҵ���ѱ�������С��100,�����䣡"
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If strText = "" Then
                strText = " "
                .Text = " "
                If Trim(.TextMatrix(.Row, .Col)) = "" Then
                    .TextMatrix(.Row, .Col) = " "
                End If
            End If
            .Text = strText
            Call ��Ǹı�
        End If
        If .Col <> COLҽ������ Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If .TxtVisible = True Then
                strText = Replace(Trim(.Text), "`", "")
                .Text = strText
                If zlCommFun.StrIsValid(strText, mlng���볤��) = False Then
                    Cancel = True
                    Exit Sub
                End If
                If mint���� = TYPE_�ɶ��ϳ� Then Exit Sub
                If Trim(strText) = "" Then
                    '����Ҫ��ȥ����Ƿ���ƥ��ı��룬�൱��ɾ���ñ���
                    .TextMatrix(.Row, COLҽ������) = Trim(strText)
                Else
                    '����SQL���
                    Select Case mint����
                        Case TYPE_����
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '��ҩƷĿ¼���ж�
                                str���� = "ҩƷ"
                                gstrSQL = "" & _
                                    " SELECT YPDM AS ҽ������,ZWM AS ��Ŀ����,PYJM AS ����,YLFL AS ҩ�����," & _
                                    "     DECODE(trim(ZFFL),'01','���ࣨ��ȫ������','02','���ࣨ���ֱ�����','03','���ࣨ��ȫ�Էѣ�','1','���ࣨ��ȫ������','2','���ࣨ���ֱ�����','3','���ࣨ��ȫ�Էѣ�','11','��ͨ����','12','�����Ը�10%','13','�����Ը�15%','14','�����Ը�20%','15','�����Ը�40%','16','�໤����1��5���Ը�30%','17','�໤����6��10���Ը�50%','19','�Է�����','δ֪') AS �Ը�����," & _
                                    "     ZDYYDJ AS ���ҽԺ�ȼ�,YPGG AS ���,YPBZDW AS ��װ��λ,YPJX AS ����,BZYYTS AS ��׼��ҩ����," & _
                                    "     ltrim(to_Char(BZJG,'9999990.00')) As ��׼�۸�, ltrim(to_Char(ZYXE,'9999990.00')) As סԺ�޶�, ltrim(to_Char(MZXE,'9999990.00')) As �����޶�, YPCD As ����,DECODE(SYFW,'0','����','1','סԺ','����סԺ����ʹ��') As ʹ�÷�Χ, BZSM As ��ע" & _
                                    " From SIM_YPML " & _
                                    "Where (upper(YPDM) Like '" & UCase(strText) & "' Or Upper(ZWM) Like '" & UCase(strText) & "' Or Upper(PYJM) Like '" & UCase(strText) & "')"
                            Else
                                '������Ŀ¼���ж�
                                str���� = "����"
                                gstrSQL = "" & _
                                " SELECT ZLDM AS ҽ������,ZLMC AS ��Ŀ����,PYJM AS ����,ZLFL AS ���Ʒ���," & _
                                "     DECODE(trim(ZFFL),'01','���ࣨ��ȫ������','02','���ࣨ���ֱ�����','03','���ࣨ��ȫ�Էѣ�','1','���ࣨ��ȫ������','2','���ࣨ���ֱ�����','3','���ࣨ��ȫ�Էѣ�','11','��ͨ����','12','�����Ը�10%','13','�����Ը�15%','14','�����Ը�20%','15','�����Ը�40%','16','�໤����1��5���Ը�30%','17','�໤����6��10���Ը�50%','19','�Է�����','δ֪') AS �Ը�����," & _
                                "     ltrim(to_Char(BZJG,'9999990.00')) As ��׼�۸�, ltrim(to_Char(ZYXE,'9999990.00')) As סԺ�޶�, ltrim(to_Char(MZXE,'9999990.00')) As �����޶�, JLDW As ������λ, ZDYYDJ As ���ҽԺ�ȼ�,DECODE(SYFW,'0','����','1','סԺ','����סԺ����ʹ��') As ʹ�÷�Χ, BZSM As ��ע" & _
                                " From SIM_ZLML " & _
                                "Where (upper(ZLDM) Like '" & UCase(strText) & "' Or Upper(ZLMC) Like '" & UCase(strText) & "' Or Upper(PYJM) Like '" & UCase(strText) & "')"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                'ǿ��ʹ��¼��Ϊ��״̬
                                gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        Case TYPE_�山ũҽ
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '��ҩƷĿ¼���ж�
                                str���� = "ҩƷ"
                                gstrSQL = "" & _
                                    " SELECT YPLSH AS ��Ŀ����,YPMC AS ��Ŀ����,PY AS ����,GG AS ���,JX AS ����,SCCJ AS ��������," & _
                                    "     CASE WHEN XJFS='0' THEN '����' WHEN xjfs='1' THEN '����' WHEN xjfs='2' THEN '�Է�' END AS ҩƷ���," & _
                                    "     CASE WHEN BXFW ='0' THEN '�弶' WHEN BXFW ='1' THEN '����' WHEN BXFW ='2' THEN '����' END AS ��ҩ��Χ,ZGXJ AS ����޼�," & _
                                    "     CASE WHEN LB='0' THEN '��ҩ' WHEN LB='1' THEN '�г�ҩ' WHEN LB='2' THEN '�в�ҩ' WHEN LB='3' THEN '��������' END AS �������" & _
                                    " From YPML" & _
                                    " Where (upper(YPLSH) Like '" & UCase(strText) & "%' Or Upper(YPMC) Like '" & UCase(strText) & "%' Or Upper(PY) Like '" & UCase(strText) & "%')"
                            Else
                                '������Ŀ¼���ж�
                                str���� = "����"
                                gstrSQL = "" & _
                                    " SELECT XMBM AS ��Ŀ����,XMMC AS ��Ŀ����,PY AS ����,CASE WHEN XJFS='0' THEN '����' WHEN XJFS='1' THEN '����' WHEN XJFS='2' THEN '�Է�' END AS ��Ŀ���," & _
                                    "     CASE WHEN BXFW='0' THEN '�弶' WHEN BXFW='1' THEN '����' WHEN BXFW='2' THEN '����' END AS ��ҩ��Χ,ZGXJ AS ����޼�," & _
                                    "     CASE WHEN XMFL='0' THEN '�Һŷ�' WHEN XMFL='1' THEN '����' WHEN XMFL='2' THEN '���Ʒ�' WHEN XMFL='3' THEN '���Ʒ�' WHEN XMFL='4' THEN '�Ĳķ�' WHEN XMFL='5' THEN '������' WHEN XMFL='6' THEN '�����' WHEN XMFL='7' THEN '��λ��' WHEN XMFL='8' THEN '��ס��'" & _
                                    "          WHEN XMFL='9' THEN '�����' WHEN XMFL='10' THEN '�����' WHEN XMFL='11' THEN '�໤��' WHEN XMFL='12' THEN '���ȷ�' WHEN XMFL='13' THEN 'B����' WHEN XMFL='14' THEN '�ʳ���' WHEN XMFL='15' THEN '������' WHEN XMFL='16' THEN '�����' WHEN XMFL='17' THEN '��ʯ��'" & _
                                    "          WHEN XMFL='18' THEN 'CT��' WHEN XMFL='19' THEN '������' WHEN XMFL='20' THEN '�ĵ�ͼ��' WHEN XMFL='21' THEN '���·�' WHEN XMFL='22' THEN '���Ʒ�' WHEN XMFL='23' THEN '������' WHEN XMFL='24' THEN '�����' WHEN XMFL='25' THEN '����' WHEN XMFL='26' THEN '����' WHEN XMFL='27' THEN '���Ʒ�' WHEN XMFL='28' THEN '������' END AS �������" & _
                                    " From ZLXM" & _
                                " Where (upper(XMBM) Like '" & UCase(strText) & "%' Or Upper(XMMC) Like '" & UCase(strText) & "%' Or Upper(PY) Like '" & UCase(strText) & "%')"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                'ǿ��ʹ��¼��Ϊ��״̬
                                gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        Case TYPE_������
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '��ҩƷĿ¼���ж�
                                str���� = "ҩƷ"
                                gstrSQL = "select YPLSH  ҽ������,YPBM ҩƷ����,TYM ͨ������,SPM ��Ʒ��,SPMZJM ��Ʒ������,YCMC ҩ������,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ� " & _
                                          "      ,PFJ ������,nvl(BZDJ,0) ��׼����,ZFBL �Ը�����,JX ����,BZSL ��װ����,BZDW ��װ��λ,HL ����,HLDW ������λ,RL ����,RLDW ������λ " & _
                                          "      ,DECODE(CFYBZ,1,'��') ����ҩ��־,decode(GMP,1,'��') GMP��־,decode(YPXJFS,1,'�޼�') �޼�,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼���� " & _
                                         "   FROM YPML WHERE YPLSH like '" & strText & "%' or Upper(TYM) like '" & strǰ & UCase(strText) & "%' Or Upper(SPM) like '" & strǰ & UCase(strText) & "%' Or Upper(SPMZJM) like '" & strǰ & UCase(strText) & "%'"
                            Else
                                '������Ŀ¼���ж�
                                str���� = "����"
                                gstrSQL = "Select XMLSH ҽ������,XMBM ���Ʊ���,XMMC ��Ŀ����,ZJM ����,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ�,DW ��λ " & _
                                         "       ,nvl(TPJ,0) ������,nvl(BZJ,0) ��׼����,ZZBL ��ְ�Ը�����,TXBL �����Ը�����,decode(XJFS,1,'ͳһ�޼�',2,'��ҽԺ�ȼ�����',3,'������ҽԺ��׼��������') �޼� " & _
                                         "       ,decode(TPXMBZ,1,'��') ������Ŀ��־,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼����,BZ ��ע " & _
                                         "   FROM ZLXM WHERE XMLSH like '" & strText & "%' or Upper(XMMC) like '" & strǰ & UCase(strText) & "%' Or Upper(ZJM) like '" & strǰ & UCase(strText) & "%'"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                'ǿ��ʹ��¼��Ϊ��״̬
                                gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        Case TYPE_����������
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                '��ҩƷĿ¼���ж�
                                str���� = "ҩƷ"
                                gstrSQL = "select ��ˮ�� ҽ������,���� ҩƷ����,ͨ���� ͨ������,��Ʒ��,��Ʒ�������� ��Ʒ������,ҩ������,decode(��Ŀ�ȼ�,1,'����',2,'����','�Է�') ���õȼ� " & _
                                          "      ,������,nvl(��׼����,0) ��׼����,�Ը�����,����,��װ����,��װ��λ,����,������λ,����,������λ " & _
                                          "      ,DECODE(����ҩ��־,1,'��') ����ҩ��־,decode(GMP��־,1,'��') GMP��־,decode(�޼۷�ʽ,1,'�޼�') �޼� " & _
                                         "   FROM �м��_ҩƷĿ¼ WHERE ��ˮ�� like '" & strText & "%' or Upper(ͨ����) like '" & strǰ & UCase(strText) & "%' Or Upper(��Ʒ��) like '" & strǰ & UCase(strText) & "%' Or Upper(��Ʒ��������) like '" & strǰ & UCase(strText) & "%'"
                            Else
                                '������Ŀ¼���ж�
                                str���� = "����"
                                gstrSQL = "Select ��ˮ�� ҽ������,��Ŀ���� ���Ʊ���,��Ŀ����,������ ����,decode(��Ŀ�ȼ�,1,'����',2,'����','�Է�') ���õȼ�,��λ " & _
                                         "       ,nvl(������,0) ������,nvl(��׼����,0) ��׼����,��ְ���� ��ְ�Ը�����,���ݱ��� �����Ը�����,decode(�޼۷�ʽ,1,'ͳһ�޼�',2,'��ҽԺ�ȼ�����',3,'������ҽԺ��׼��������') �޼� " & _
                                         "       ,decode(������Ŀ��־,1,'��') ������Ŀ��־,��ע " & _
                                         "   FROM �м��_������Ŀ WHERE ��ˮ�� like '" & strText & "%' or Upper(��Ŀ����) like '" & strǰ & UCase(strText) & "%' Or Upper(������) like '" & strǰ & UCase(strText) & "%'"
                            End If
                            If mcnYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
                            Else
                                'ǿ��ʹ��¼��Ϊ��״̬
                                gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                        'Modified by ���� 20031218 ����������
                        Case TYPE_��������, TYPE_����ʡ, TYPE_������, TYPE_��ƽ��
                            '20031229:���,��ֹ�ظ�
                            gstrSQL = "   Select Distinct A.���� as ҽ������,A.����,A.����,B.���� as ����,A.��ע " & _
                                      "   FROM ������Ŀ A,����֧������ B" & _
                                      "   WHERE A.�������=B.���� And A.����=" & mint���� & " And B.����=A.����" & _
                                      " And (A.���� like '" & strText & "%' or Upper(A.����) like '" & strǰ & UCase(strText) & "%' Or Upper(A.����) like '" & strǰ & UCase(strText) & "%')"
                            Call OpenRecordset(rsTemp, Me.Caption)
                        Case TYPE_������
                            gstrSQL = "SELECT A.���� ҽ������,A.����,A.����,A.��λ,B.���� AS ����,C.���� AS ���� " & _
                                      "     ,A.�Ƿ���ҩ,A.�Ƿ�ҽ��,A.���۸�����,A.�����Ը�����,A.�۸�,A.��Ŀ�ں�,A.��������,A.˵�� " & _
                                      "  FROM ������Ŀ A,����֧������ B,���� C " & _
                                      "  WHERE A.����=" & TYPE_������ & " AND A.�������=B.����(+) AND A.���ͱ���=c.����(+) And (" & _
                                      zlCommFun.GetLike("A", "����", strText) & " Or " & zlCommFun.GetLike("A", "����", strText) & " Or " & zlCommFun.GetLike("A", "����", strText) & ")"
                            rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
                        Case type_������, type_����������
                            '200311
                            gstrSQL = "   Select A.����  ҽ������,A.����,A.����,B.���� ����,A.��ע " & _
                                      "   FROM ������Ŀ A,����֧������ B" & _
                                      "   WHERE A.�������=B.���� and b.����=" & mint���� & " And A.����=" & mint���� & " and (A.���� like '" & strText & "%' or Upper(A.����) like '" & strǰ & UCase(strText) & "%' Or Upper(A.����) like '" & strǰ & UCase(strText) & "%')"
                                      
                            Call OpenRecordset(rsTemp, Me.Caption)
                        Case TYPE_����
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "" & _
                                    " Select A.���� AS ҽ������,A.��Ŀ,A.����,A.������,A.������λ AS ��λ,B.���� As ���ⲡ,H.���� AS ��Ŀ�ȼ�,A.��׼����,A.�Ը�����,0 �޼�," & _
                                    " C.���� AS ����ҩ,F.���� AS ����,A.�÷�,A.�ճ�������,D.���� AS ҩƷ����,G.���� AS ����,E.���� AS ʹ�����Ƶȼ�,A.��ע,A.��Ч����" & _
                                    " From ҩƷĿ¼ A," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='������ҩ��ʶ' and A.���=B.���) B," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='����ҩ��־' and A.���=B.���) C," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='ҩƷ����' and A.���=B.���) D," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='ʹ�����Ƶȼ�' and A.���=B.���) E,"
                                gstrSQL = gstrSQL & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='����' and A.���=B.���) F," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='����' and A.���=B.���) G," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='�շ���Ŀ�ȼ�' and A.���=B.���) H" & _
                                    " Where A.���ⲡ =B.����(+) And A.����ҩ=C.����(+) And A.ҩƷ���� =D.����(+)" & _
                                    " And A.ʹ�����Ƶȼ�=E.����(+) And A.����=F.����(+) And A.����=G.����(+) AND A.ҩƷ�ȼ�=H.����(+)" & _
                                    " And (" & zlCommFun.GetLike("A", "����", strText) & " Or " & zlCommFun.GetLike("A", "����", strText) & " Or " & zlCommFun.GetLike("A", "������", strText) & ")"
                            Else
                                '��ǰѡ���ǵ�һ���������
                                gstrSQL = "" & _
                                    " Select A.���� AS ҽ������,A.����,A.������,A.��λ,B.���� AS ���ⲡ,C.���� AS ��Ŀ�ȼ�,A.��׼����,A.�Ը�����,A.�޼�,A.��ע,A.��Ч����" & _
                                    "      From ����Ŀ¼ A," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='������ҩ��ʶ' and A.���=B.���) B," & _
                                    "      (Select B.����,B.����" & _
                                    "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
                                    "      Where A.����='�շ���Ŀ�ȼ�' and A.���=B.���) C" & _
                                    " Where A.���ⲡ =B.����(+) And A.��Ŀ�ȼ�=C.����(+)" & _
                                    " AND (" & zlCommFun.GetLike("A", "����", strText) & " Or " & zlCommFun.GetLike("A", "����", strText) & " Or " & zlCommFun.GetLike("A", "������", strText) & ")"
                            End If
                            If rsTemp.State = 1 Then rsTemp.Close
                            rsTemp.Open gstrSQL, mcnYB
                        Case TYPE_�����山
                                
                                strLeft = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
                                strTemp = "'" & strLeft & strText & "%'"
                                
                                gstrSQL = " select  ��Ʒ���� as ҽ������,  ҽԺ�������, ҩƷͨ��������, ҩƷͨ��Ӣ����,��Ʒ��, ��Ʒ������, ������Ŀ���㷽ʽ, ������ʶ, ҽ����ʶ, �Ƿ񴦷���ҩ, ҩƷ��Ӧ֢, ����ҽ��, ����Ȩ��, ����, ��װ���, " & _
                                         "         ��С��װ��λ, ��С������λ, ÿ���������, ָ���۸�, �б�۸�, ����֧���޼�1, ����֧���޼�2, ����֧���޼�3, ʵ��ִ�м۸�, �Ը�����1, �Ը�����2, �Ը�����3, �Ը�����4, �Ը�����5, �Ը�����6, �Ը�����7, �Ը�����8,  " & _
                                         "         �Ը�����9, �Ը�����10, �Ը�����11, �Ը�����12, ҽԺʹ��״̬, ����ʹ��״̬, ��׼���,  " & _
                                         "         ���������1, ���������2, ���������3, ƴ��������1, ƴ��������2, ƴ��������3, ��ע, ҽ���������,������׼���, ҽ�ƻ������, " & _
                                         "          �޸�ʱ��, Ŀ¼����  " & _
                                         "  from ҽ��������ĿĿ¼" & _
                                         "  where ��Ʒ���� like " & strTemp & " Or ��Ʒ�� like " & strTemp & " Or " & _
                                         "        ���� like " & strTemp & " Or ���������1 like " & UCase(strTemp) & " Or " & _
                                         "        ƴ��������1 like " & UCase(strTemp)
                             Debug.Print Time
                            If gcnOracle_CQYB.State = adStateOpen Then
                                rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
                            Else
                                'ǿ��ʹ��¼��Ϊ��״̬
                                gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                             Debug.Print Time
                             gstrSQL = ""
                        Case TYPE_ǭ��
                                
                                strLeft = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
                                strTemp = "'" & strLeft & strText & "%'"
                                
                                gstrSQL = "" & _
                                    "  Select  ���,decode(���,'1','ҩƷ','2','����','����') as ��Ŀ���,���||���� as ҽ������,����, Ӣ������,�շ����, �շѵȼ�, ������, ��λ, ��׼�۸�, ֧����׼, ����, ���, ��ע, ���ʱ��, ά����־  " & _
                                    "  From ҽ���շ�Ŀ¼" & _
                                    "  Where ���� like " & strTemp & " Or ���� like " & strTemp & " Or " & _
                                    "        �շ���� like " & strTemp & " Or ������ like " & UCase(strTemp) & _
                                    "   order by ���,����"
                                    
                            Debug.Print Time
                            If Not gcnOracle_ǭ�� Is Nothing Then
                                If gcnOracle_ǭ��.State = adStateOpen Then
                                    rsTemp.Open gstrSQL, gcnOracle_ǭ��, adOpenStatic, adLockReadOnly
                                Else
                                    'ǿ��ʹ��¼��Ϊ��״̬
                                    gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
                                    Call OpenRecordset(rsTemp, Me.Caption)
                                End If
                            Else
                                'ǿ��ʹ��¼��Ϊ��״̬
                                gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
                                Call OpenRecordset(rsTemp, Me.Caption)
                            End If
                             Debug.Print Time
                             gstrSQL = ""
                        Case TYPE_�㽭
                            strLeft = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
                            strTemp = "'" & strLeft & UCase(strText) & "%'"
                            If gcn�㽭.State = 0 Then
                                Call openConn�㽭
                            End If
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select AKA060 As ҽ������, AKA061 As ��Ŀ����, AKA069 As �Ը�����, AKA066 As ƴ����, AKA070 As ����, Decode(AKA065,'1','����','2','����','����') As ��� From KA02 " & _
                                    "Where AKA060 Like " & strTemp & " Or AKA061 Like " & strTemp & " Or AKA066 Like " & strTemp
                            Else
                                gstrSQL = "Select AKA090 As ҽ������, AKA091 As ��Ŀ����, AKA069 As �Ը�����, AKA066 As ƴ����, Decode(AKA065,'1','����','2','����','����') As ��� From KA03 " & _
                                    "Where AKA090 Like " & strTemp & " Or AKA091 Like " & strTemp & " Or AKA066 Like " & strTemp
                            End If
                            If gcn�㽭.State = 1 Then Set rsTemp = gcn�㽭.Execute(gstrSQL)
                        Case TYPE_�¶�
                            Dim cn�¶� As New ADODB.Connection
                            
                            strLeft = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
                            strTemp = "'" & strLeft & strText & "%'"
                            
                            cn�¶�.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\YWCS.MDB;Persist Security Info=True;Jet OLEDB:Database Password=yhybv1.1cdb"
                            cn�¶�.CursorLocation = adUseClient
                            cn�¶�.Open

                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select ybxmbm As ҽ������,ybxmmc As ��Ŀ����,zfbl1 As �Ը����� From KYH904 " & _
                                          "Where ybxmbm Like " & UCase(strTemp) & " Or ybxmmc Like " & UCase(strTemp)
                            Else
                                gstrSQL = "Select ybxmbm As ҽ������,ybxmmc As ��Ŀ����,zgxj As һ��ҽԺ�۸�,zgxj1 As ����ҽԺ�۸�,zgxj2 As ����ҽԺ�۸�,zfbl1 As �Ը����� From KYH100 " & _
                                          "Where ybxmbm Like " & UCase(strTemp) & " Or ybxmmc Like " & UCase(strTemp)
                            End If
                            Set rsTemp = cn�¶�.Execute(gstrSQL)
                        Case TYPE_��Ҧ
                            strLeft = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
                            strTemp = "'" & strLeft & UCase(strText) & "%'"
                            If gcn��Ҧ.State = 0 Then
                                Call openConn��Ҧ
                            End If
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select MedicineID As ҽ������,Name As ��Ŀ����,DoseType As ����,ZFBL As �Ը�����,NameJP As ƴ������,NameWB As ����� From hi_Medicine " & _
                                    "Where MedicineID Like " & strTemp & " Or Name Like " & strTemp & " Or NameJP Like " & strTemp & " Or NameWB Like " & strTemp
                            Else
                                gstrSQL = "Select DiagnoseID As ҽ������,Name As ��Ŀ����,'' As ����,ZFBL As �Ը�����,NameJP As ƴ������,NameWB As ����� From hi_Diagnose " & _
                                    "Where DiagnoseID Like " & strTemp & " Or Name Like " & strTemp & " Or NameJP Like " & strTemp & " Or NameWB Like " & strTemp
                            End If
                            If gcn��Ҧ.State = 1 Then Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
                            
                        'Modified by ZYB �Ͻ�
                        Case TYPE_�Ͻ�
                            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                gstrSQL = "Select ҩƷ���� AS ҽ������,�������� AS ��Ŀ����,Ӣ������,��Ʒ����,ҩƷ����,��������,�����Ը�����||'%' AS �����Ը�����,�����𸶽�� From ҩƷĿ¼�� A" & _
                                " Where (" & zlCommFun.GetLike("A", "ҩƷ����", strText) & " Or " & zlCommFun.GetLike("A", "��������", strText) & ")"
                            Else
                                gstrSQL = "Select ������Ŀ���� AS ҽ������,������Ŀ���� AS ��Ŀ����,�������,һ��ҽԺ����,����ҽԺ����,����ҽԺ����,�����Ը�����||'%' AS �����Ը�����,�����𸶽�� From ������Ŀ�� A" & _
                                " Where (" & zlCommFun.GetLike("A", "������Ŀ����", strText) & " Or " & zlCommFun.GetLike("A", "������Ŀ����", strText) & ")"
                            End If
                            If rsTemp.State = 1 Then rsTemp.Close
                            rsTemp.Open gstrSQL, mcnYB
                        Case Else
                            If mint���� = TYPE_������ Then
                                Call GetItemMatchInfo
                                If int��˱�־ = 1 And mint���õ��� = 0 Then
                                    MsgBox "����Ŀ�Ѿ�ͨ����ˣ��������޸ģ�", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                            gstrSQL = "Select ����  ҽ������,����,����,��ע " & _
                                     "   FROM ������Ŀ WHERE ����=" & mint���� & " and (���� like '" & strText & "%' or Upper(����) like '" & strǰ & UCase(strText) & "%' Or Upper(����) like '" & strǰ & UCase(strText) & "%')"
                            Call OpenRecordset(rsTemp, Me.Caption)
                    End Select
                    
                    If rsTemp.RecordCount > 0 Then
                        '����ѡ����
                        If rsTemp.RecordCount >= 1 Or rsTemp.Fields.Count > 3 Then
                            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
                            blnReturn = frmListSel.ShowSelect(rsTemp, "ҽ������", "ҽ����Ŀѡ��", "��ѡ���Ӧ��ҽ����Ŀ��")
                        End If
                    Else
                        If mint���� = TYPE_�ɶ��ڽ� Then
                            MsgBox "����ָ��ҽ����Ŀ��������!"
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    If blnReturn = False Then
                        '��¼����û�п�ѡ�������
                        If rsTemp.RecordCount > 0 Then
                            '��¼�������ݣ���ȡ����ѡ��
                            Cancel = True
                            .TxtVisible = True
                            .TxtSetFocus
                            Exit Sub
                        Else
                            If Not (mint���� = TYPE_������ Or mint���� = TYPE_���� Or mint���� = TYPE_������ Or mint���� = TYPE_�Ͻ� Or mint���� = TYPE_ǭ�� Or mint���� = TYPE_���� Or mint���� = TYPE_�山ũҽ) Then
                                .Text = strText
                                .TextMatrix(.Row, COLҽ������) = strText
                            Else
                                .Text = ""
                                .TextMatrix(.Row, COLҽ������) = ""
                                Cancel = True
                                Exit Sub
                            End If
                        End If
                    Else
                        '�϶����м�¼����
                        If mint���� = TYPE_ǭ�� Then
                            .Text = Mid(rsTemp("ҽ������"), 2)
                        Else
                            .Text = rsTemp("ҽ������")
                        End If
                        Dim str�޼� As String
                        Select Case mint����
                            Case TYPE_������
                                '���������ҽ�����Ƕ���Ŀ�ļ۸�����ж�
                                
                                str�޼� = Nvl(rsTemp("�޼�"), "")
                                If str�޼� <> "" And Val(.TextMatrix(.Row, col�۸�)) > 0 Then
                                    '�������޼�
                                    If str���� = "ҩƷ" Then
                                        'ҩƷû��������
                                        blnReturn = �۸��ж�_����(Val(.TextMatrix(.Row, col�۸�)), rsTemp("��׼����"), str�޼�, False, 0)
                                    Else
                                        blnReturn = �۸��ж�_����(Val(.TextMatrix(.Row, col�۸�)), rsTemp("��׼����"), str�޼�, Nvl(rsTemp("������Ŀ��־"), "") = "��", rsTemp("������"))
                                    End If
                                    If blnReturn = False Then
                                        Cancel = True
                                        .TxtVisible = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                End If
                            Case TYPE_����������
                                '���������ҽ�����Ƕ���Ŀ�ļ۸�����ж�
                                str�޼� = Nvl(rsTemp("�޼�"), "")
                                If str�޼� <> "" And Val(.TextMatrix(.Row, col�۸�)) > 0 Then
                                    '�������޼�
                                    If str���� = "ҩƷ" Then
                                        'ҩƷû��������
                                        blnReturn = �۸��ж�_����������(Val(.TextMatrix(.Row, col�۸�)), rsTemp("��׼����"), str�޼�, False, 0)
                                    Else
                                        blnReturn = �۸��ж�_����������(Val(.TextMatrix(.Row, col�۸�)), rsTemp("��׼����"), str�޼�, Nvl(rsTemp("������Ŀ��־"), "") = "��", rsTemp("������"))
                                    End If
                                    If blnReturn = False Then
                                        Cancel = True
                                        .TxtVisible = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                End If
                            Case TYPE_������
                                If Nvl(rsTemp("���۸�����"), 0) <> 0 And Val(.TextMatrix(.Row, col�۸�)) > 0 Then
                                    If rsTemp("���۸�����") < Val(.TextMatrix(.Row, col�۸�)) Then
                                        If MsgBox("ҽԺ����" & Format(Val(.TextMatrix(.Row, col�۸�)), "0.000") & _
                                            " ����ҽ�����ĺ�׼�ļ۸�" & Format(rsTemp("���۸�����"), "0.000") & "���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                                            Cancel = True
                                            .TxtVisible = True
                                            .TxtSetFocus
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case TYPE_����
                                '�����ҩƷ��Ŀ�����HIS����Ŀ�����Ƿ���ҩƷ������
                                Dim rsCheck As New ADODB.Recordset
                                If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                                    If Not CheckTradeName(.RowData(.Row), rsTemp("ҽ������")) Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                End If
                            Case TYPE_������
                                Call CheckValid(.Text)
                            Case TYPE_�㽭
                                .TextMatrix(.Row, COLҽ����ע) = rsTemp!���
                                .TextMatrix(.Row, colҽ������) = rsTemp("��Ŀ����")
                            Case TYPE_��Ҧ
                                .TextMatrix(.Row, colҽ������) = rsTemp("��Ŀ����")
                            Case TYPE_ǭ��
                                '�ע
                                .TextMatrix(.Row, COLҽ����ע) = rsTemp("���")
                                .TextMatrix(.Row, colҽ������) = rsTemp("����")
                        End Select
                        If mint���� = TYPE_ǭ�� Then
                            .TextMatrix(.Row, COLҽ������) = Mid(rsTemp("ҽ������"), 2)
                        Else
                            .TextMatrix(.Row, COLҽ������) = rsTemp("ҽ������")
                        End If
                    End If
                End If
                Call Get��������
                Call ��Ǹı�
                'Modified By ���� ��������ɳ ԭ�򣺸��ݵ�ǰ����������Ŀƥ����Ϣ
                If mint���� = TYPE_������ Then
                    If .TextMatrix(.Row, COLҽ����ע) <> "" Then
                        .TextMatrix(.Row, colҽ������) = Split(.TextMatrix(.Row, COLҽ����ע), "||")(3)
                    End If
                End If
                Call SetItemMatch(False)
            Else
                If .TextMatrix(.Row, COLҽ������) = "" Then
                    .TextMatrix(.Row, COLҽ������) = " "
                End If
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Cancel = True
End Sub

Private Sub ��Ǹı�()
    '��ǰ�����Ѿ���Ч���������ܷ�õ���������
    cmdRestore.Enabled = True
    cmdSave.Enabled = True
    
    With mshSum_S
        If Trim(.TextMatrix(.Row, COLҽ������)) = "" And Trim(.TextMatrix(.Row, col��������)) = "" Then
            .TextMatrix(.Row, col�ı䷽ʽ) = "ɾ��"
        Else
            If Trim(.TextMatrix(.Row, col�ı䷽ʽ)) <> "�޸�" Then
                'Ϊ�գ����Ѿ��ǡ�������
                .TextMatrix(.Row, col�ı䷽ʽ) = "����"
            End If
        End If
    End With
End Sub

Private Sub Get��������()
'���ܣ����ݵ�ǰ�еı�����Ŀ���룬�õ�������Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long, lngPos As Long
    Dim str������� As String, strTemp As String, varPart As Variant
    
    On Error GoTo errHandle
    With mshSum_S
        If mint���� = TYPE_������ Then
            If mcnYB.State = adStateOpen Then
                
                gstrSQL = "Select SPM ����,'' �������,'' ��ע  From YPML WHERE yplsh='" & .TextMatrix(.Row, COLҽ������) & "' " & _
                           " Union All " & _
                           " Select XMMC ����,'' �������,'' ��ע  From ZLXM WHERE XMLSH='" & .TextMatrix(.Row, COLҽ������) & "'"
                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
            Else
                'ǿ��ʹ��¼��Ϊ��״̬
                gstrSQL = "select ����,�������,��ע from ������Ŀ where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint���� = TYPE_���������� Then
            '��������ҽ�������� 204-03-29
            If mcnYB.State = adStateOpen Then
                
                gstrSQL = "Select ��Ʒ�� ����,lpad(��Ŀ�ȼ�,6,'0') �������,'' ��ע  From �м��_ҩƷĿ¼ WHERE ��ˮ��='" & .TextMatrix(.Row, COLҽ������) & "' " & _
                           " Union All " & _
                           " Select ��Ŀ���� ����,lpad(��Ŀ�ȼ�,6,'0') �������,'' ��ע  From �м��_������Ŀ WHERE ��ˮ��='" & .TextMatrix(.Row, COLҽ������) & "'"
                rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
            Else
                'ǿ��ʹ��¼��Ϊ��״̬
                gstrSQL = "select ����,�������,��ע from ������Ŀ where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint���� = TYPE_�ɶ��ϳ� Then
            gstrSQL = " Select '' DLBM," & gstrCol_ENG & _
                      " From yljcxxk " & _
                      " Where ID=" & Val(.TextMatrix(.Row, COLҽ������))
            Call OpenRecordset(rsTemp, Me.Caption)
        ElseIf mint���� = TYPE_������ Then
            gstrSQL = "SELECT A.���� ҽ������,A.����,A.����,A.�������,A.�Ƿ�ҽ�� " & _
                      "  FROM ������Ŀ A " & _
                      "  WHERE A.����=" & TYPE_������ & " AND A.����='" & .TextMatrix(.Row, COLҽ������) & "'"
            rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
        ElseIf mint���� = type_������ Or mint���� = type_���������� Then
              '���˺�(200311)
            
            gstrSQL = "SELECT A.���� ҽ������,A.����,A.����,A.�������,B.�Ƿ�ҽ�� " & _
                      "  FROM ������Ŀ A,����֧������ B " & _
                      "  WHERE A.�������=B.����(+) and b.����=" & cmb����.ItemData(cmb����.ListIndex) & " and A.����=" & cmb����.ItemData(cmb����.ListIndex) & " AND A.����='" & .TextMatrix(.Row, COLҽ������) & "'"
            OpenRecordset rsTemp, Me.Caption
        ElseIf mint���� = TYPE_���� Then
            gstrSQL = " SELECT ����,���� From ҩƷĿ¼ WHERE ����='" & .TextMatrix(.Row, COLҽ������) & "'" & _
                      " Union " & _
                      " SELECT ����,���� From ����Ŀ¼ WHERE ����='" & .TextMatrix(.Row, COLҽ������) & "'"
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open gstrSQL, mcnYB
        ElseIf mint���� = TYPE_�����山 Then
            '20040706
            gstrSQL = " SELECT ��Ʒ���� ����,��Ʒ�� ���� From ҽ��������ĿĿ¼ WHERE ��Ʒ����='" & .TextMatrix(.Row, COLҽ������) & "'"
            rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
        'Modified by ZYB 2004-08-17
        ElseIf mint���� = TYPE_��ɽ Then
            gstrSQL = "select substr(��ע,Instr(��ע,'|',1,3)+1)||'-'||���� AS ����,�������,��ע from ������Ŀ where ����='" & .TextMatrix(.Row, COLҽ������) & _
                      "' and ����=" & cmb����.ItemData(cmb����.ListIndex)
            Call OpenRecordset(rsTemp, Me.Caption)
        ElseIf mint���� = TYPE_�Ͻ� Then
            gstrSQL = " SELECT ������Ŀ���� ����,������Ŀ���� ����,������� ���� From ������Ŀ�� WHERE ������Ŀ����='" & .TextMatrix(.Row, COLҽ������) & "'" & _
                      " Union " & _
                      " Select ҩƷ���� ����,�������� ����,ҩƷ���� ���� From ҩƷĿ¼�� Where ҩƷ����='" & .TextMatrix(.Row, COLҽ������) & "'"
            rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
        ElseIf mint���� = TYPE_�㽭 Then
            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                gstrSQL = "Select AKA060 As ҽ������, AKA061 As ��Ŀ����, AKA069 As �Ը�����, AKA066 As ƴ����, Decode(AKA065,'1','����','2','����','����') As ��� From KA02 Where AKA060='" & .TextMatrix(.Row, COLҽ������) & "'"
            Else
                gstrSQL = "Select AKA090 As ҽ������, AKA091 As ��Ŀ����, AKA069 As �Ը�����, AKA066 As ƴ����, Decode(AKA065,'1','����','2','����','����') As ��� From KA03 Where AKA090='" & .TextMatrix(.Row, COLҽ������) & "'"
            End If
            Set rsTemp = gcn�㽭.Execute(gstrSQL)
        ElseIf mint���� = TYPE_��Ҧ Then
            If Left(tvwMain_S.SelectedItem.Key, 1) = "D" Or Left(tvwMain_S.SelectedItem.Key, 1) = "E" Or Left(tvwMain_S.SelectedItem.Key, 1) = "F" Then
                gstrSQL = "Select MedicineID As ҽ������,Name As ��Ŀ����,DoseType As ����,ZFBL As �Ը�����,NameJP As ƴ������,NameWB As ����� From hi_Medicine Where MedicineID='" & .TextMatrix(.Row, COLҽ������) & "'"
            Else
                gstrSQL = "Select DiagnoseID As ҽ������,Name As ��Ŀ����,ZFBL As �Ը�����,NameJP As ƴ������,NameWB As ����� From hi_Diagnose Where DiagnoseID='" & .TextMatrix(.Row, COLҽ������) & "'"
            End If
            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
        ElseIf mint���� = TYPE_���� Then
            If mcnYB.State = 1 Then
                gstrSQL = " Select YPDM ����,ZWM ����,PYJM ���� From SIM_YPML " & _
                          " Where (upper(YPDM) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(ZWM) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(PYJM) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "')" & _
                          " UNION " & _
                          " Select ZLDM ����,ZLMC ����,PYJM ���� From SIM_ZLML " & _
                          " Where (upper(ZLDM) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(ZLMC) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(PYJM) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "')"
                If rsTemp.State = 1 Then rsTemp.Close
                rsTemp.Open gstrSQL, mcnYB
            Else
                'ǿ��ʹ��¼��Ϊ��״̬
                gstrSQL = "select ����,�������,��ע from ������Ŀ where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint���� = TYPE_�山ũҽ Then
            If mcnYB.State = 1 Then
                gstrSQL = " Select YPLSH AS ��Ŀ����,YPMC AS ��Ŀ����,PY AS ���� From YPML " & _
                          " Where (upper(YPLSH) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(YPMC) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(PY) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "')" & _
                          " UNION " & _
                          " Select XMBM AS ��Ŀ����,XMMC AS ��Ŀ����,PY AS ���� From ZLXM " & _
                          " Where (upper(XMBM) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(XMMC) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "' Or Upper(PY) Like '" & UCase(.TextMatrix(.Row, COLҽ������)) & "')"
                If rsTemp.State = 1 Then rsTemp.Close
                rsTemp.Open gstrSQL, mcnYB
            Else
                'ǿ��ʹ��¼��Ϊ��״̬
                gstrSQL = "select ����,�������,��ע from ������Ŀ where rownum<1"
                Call OpenRecordset(rsTemp, Me.Caption)
            End If
        ElseIf mint���� = TYPE_ǭ�� Then
            '20040706
            gstrSQL = " SELECT ���,����,���� From ҽ���շ�Ŀ¼ WHERE ���=" & Val(.TextMatrix(.Row, COLҽ����ע)) & " and ����='" & .TextMatrix(.Row, COLҽ������) & "'"
            rsTemp.Open gstrSQL, gcnOracle_ǭ��, adOpenStatic, adLockReadOnly
        Else
            gstrSQL = "select ����,�������,��ע from ������Ŀ where ����='" & .TextMatrix(.Row, COLҽ������) & _
                      "' and ����=" & cmb����.ItemData(cmb����.ListIndex)
            Call OpenRecordset(rsTemp, Me.Caption)
        End If
        
        If rsTemp.RecordCount = 0 Then
            'û�ж�Ӧ�ı�����Ŀ��ֻ�����øñ���
            .TextMatrix(.Row, colҽ������) = ""
            .TextMatrix(.Row, COLҽ����ע) = ""
            .TextMatrix(.Row, col��ҽ��) = ""
        ElseIf mint���� = TYPE_��Ҧ Then
            .TextMatrix(.Row, colҽ������) = Nvl(rsTemp!��Ŀ����, "")
            .TextMatrix(.Row, COLҽ����ע) = ""
            .TextMatrix(.Row, col��ҽ��) = ""
        ElseIf mint���� = TYPE_�㽭 Then
            .TextMatrix(.Row, colҽ������) = Nvl(rsTemp!��Ŀ����, "")
            .TextMatrix(.Row, COLҽ����ע) = Nvl(rsTemp!���, "����")
            .TextMatrix(.Row, col��ҽ��) = ""
        ElseIf mint���� = TYPE_���� Then
            .TextMatrix(.Row, COLҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        ElseIf mint���� = TYPE_�山ũҽ Then
            .TextMatrix(.Row, COLҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        ElseIf mint���� = TYPE_������ Then
            .TextMatrix(.Row, colҽ������) = Nvl(rsTemp("����"))
            .TextMatrix(.Row, col��ҽ��) = IIf(rsTemp("�Ƿ�ҽ��") = 1, "", "��")
            .TextMatrix(.Row, COLҽ����ע) = ""
            str������� = Nvl(rsTemp("�������"))
        ElseIf mint���� = TYPE_�ɶ��ϳ� Then
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp(ExchangeColName("����", False))), "", rsTemp(ExchangeColName("����", False)))
            .TextMatrix(.Row, COLҽ����ע) = IIf(IsNull(rsTemp(ExchangeColName("ҩƷ��Ŀ�ں�", False))), "", rsTemp(ExchangeColName("ҩƷ��Ŀ�ں�", False)))
        ElseIf mint���� = type_������ Or mint���� = type_���������� Then
            .TextMatrix(.Row, col��ҽ��) = IIf(rsTemp("�Ƿ�ҽ��") = 1, "", "��")
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            str������� = Nvl(rsTemp("�������"))
        ElseIf mint���� = TYPE_���� Then
            '.TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        ElseIf mint���� = TYPE_�����山 Then
            .TextMatrix(.Row, COLҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        ElseIf mint���� = TYPE_�Ͻ� Then
            .TextMatrix(.Row, COLҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, COLҽ����ע) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        ElseIf mint���� = TYPE_ǭ�� Then
            .TextMatrix(.Row, COLҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, COLҽ����ע) = IIf(IsNull(rsTemp("���")), "", rsTemp("���"))
        Else
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, COLҽ����ע) = IIf(IsNull(rsTemp("��ע")), "", rsTemp("��ע"))
            str������� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
            '�Թ�ҽ�������õ���ע�еĴ������
            If mint���� = TYPE_�Թ��� Then
                strTemp = .TextMatrix(.Row, COLҽ����ע)
                strTemp = Mid(strTemp, InStr(strTemp, "|") + 1)    'ȥ����һ����ͣ�
                strTemp = Mid(strTemp, 1, InStr(strTemp, "|") - 1) '�õ��ڶ���Ƿ�ҽ����
                .TextMatrix(.Row, col��ҽ��) = IIf(strTemp = 0, "��", "")
            ElseIf mint���� = TYPE_�Ĵ�üɽ Then
                strTemp = .TextMatrix(.Row, COLҽ����ע)
                varPart = Split(strTemp, "|")
                If UBound(varPart) >= 3 Then
                    .TextMatrix(.Row, col��ҽ��) = IIf(varPart(2) = "N", "��", "")
                Else
                    .TextMatrix(.Row, col��ҽ��) = ""
                End If
            'Modified by ���� 20031218 ����������
            ElseIf mint���� = TYPE_�������� Or mint���� = TYPE_����ʡ Or mint���� = TYPE_������ Or mint���� = TYPE_��ƽ�� Then
                strTemp = .TextMatrix(.Row, COLҽ����ע)
                varPart = Split(strTemp, "|")
                If UBound(varPart) >= 3 Then
                    .TextMatrix(.Row, col��ҽ��) = IIf(varPart(3) = "N", "��", "")
                Else
                    .TextMatrix(.Row, col��ҽ��) = ""
                End If
            End If
        End If
        
        For lngIndex = 0 To .ListCount - 1
            lngPos = InStr(.List(lngIndex), ".")
            If lngPos = 0 Then
                strTemp = .List(lngIndex)
            Else
                strTemp = Mid(.List(lngIndex), 1, lngPos - 1)
            End If
            If strTemp = str������� Then
                '�ҵ���ƥ��Ĵ������
                .TextMatrix(.Row, col����ID) = .ItemData(lngIndex)
                .TextMatrix(.Row, col��������) = .List(lngIndex)
                Exit For
            End If
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mshSum_S_KeyPress(KeyAscii As Integer)
    With mshSum_S
        If Not .Active Then Exit Sub
        If .ColData(.Col) = -1 Then Call ��Ǹı�
    End With
End Sub

Private Sub mshSum_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mshSum_S.ToolTipText = mshSum_S.TextMatrix(mshSum_S.MouseRow, mshSum_S.MouseCol)
End Sub

Private Sub mshSum_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rsTemp As New ADODB.Recordset, lngID As Long
    Dim lngRow As Long, lngPos As Long, blnActive As Boolean
    Dim blnEnable As Boolean
    
    If mshSum_S.Active = False Then Exit Sub
    If mshSum_S.MouseRow = 0 Then
        If mlngCol = mshSum_S.MouseCol Then
            mblnDesc = Not mblnDesc
        Else
            mlngCol = mshSum_S.MouseCol
            mblnDesc = False
        End If
        
        blnEnable = cmdRestore.Enabled
        blnActive = mshSum_S.Active
        mshSum_S.Active = False
        mshSum_S.msfObj.MousePointer = vbHourglass
        
        '���ɼ�¼����Ȼ��ˢ�±��
        rsTemp.CursorLocation = adUseClient
        rsTemp.CursorType = adOpenDynamic
        rsTemp.LockType = adLockOptimistic
        With rsTemp.Fields
            .Append "ID", adDouble, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "����", adVarChar, 50, adFldIsNullable
            .Append "���", adVarChar, 80, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "��λ", adVarChar, 20, adFldIsNullable
            .Append "�Ƿ���", adInteger, adFldIsNullable
            .Append "�۸�", adVarNumeric, 20, adFldIsNullable
            .Append "�ı䷽ʽ", adVarChar, 4, adFldIsNullable
            'Modified By ���� 2003-12-09 ��������ɽ
            .Append "��Ŀ����", adVarChar, 50, adFldIsNullable
            .Append "��Ŀ����", adVarChar, 50, adFldIsNullable
            .Append "��ע", adVarChar, 50, adFldIsNullable
            .Append "ԭ����", adVarChar, 20, adFldIsNullable
            .Append "�Ƿ�ҽ��", adInteger
            .Append "����ID", adDouble
            .Append "�������", adVarChar, 10, adFldIsNullable
            .Append "��������", adVarChar, 50, adFldIsNullable
        End With
        
        rsTemp.Open
        With mshSum_S
            For lngRow = 1 To .Rows - 1
                rsTemp.AddNew
                
                rsTemp("ID") = .RowData(lngRow)
                rsTemp("����") = .TextMatrix(lngRow, cOL����)
                rsTemp("����") = .TextMatrix(lngRow, cOL����)
                rsTemp("���") = .TextMatrix(lngRow, COL���)
                rsTemp("����") = .TextMatrix(lngRow, COL����)
                rsTemp("����") = Substr(.TextMatrix(lngRow, col����), 1, 100)
                rsTemp("��λ") = .TextMatrix(lngRow, COL��λ)
                If .TextMatrix(lngRow, col�۸�) = "" Then
                    rsTemp("�Ƿ���") = 1
                    rsTemp("�۸�") = 0
                Else
                    rsTemp("�Ƿ���") = 0
                    rsTemp("�۸�") = Val(.TextMatrix(lngRow, col�۸�))
                End If
                rsTemp("�ı䷽ʽ") = .TextMatrix(lngRow, col�ı䷽ʽ)
                rsTemp("��Ŀ����") = .TextMatrix(lngRow, COLҽ������)
                rsTemp("��Ŀ����") = .TextMatrix(lngRow, colҽ������)
                rsTemp("��ע") = .TextMatrix(lngRow, COLҽ����ע)
                rsTemp("ԭ����") = .TextMatrix(lngRow, colԭ����)
                rsTemp("����ID") = Val(.TextMatrix(lngRow, col����ID))
                rsTemp("�Ƿ�ҽ��") = IIf(.TextMatrix(lngRow, col��ҽ��) = "��", 0, 1)
                
                
                lngPos = InStr(.TextMatrix(lngRow, col��������), ".")
                If lngPos = 0 Then
                    rsTemp("�������") = Null
                    rsTemp("��������") = Null
                Else
                    rsTemp("�������") = Mid(.TextMatrix(lngRow, col��������), 1, lngPos - 1)
                    rsTemp("��������") = Mid(.TextMatrix(lngRow, col��������), lngPos + 1)
                End If
                
                rsTemp.Update
            Next
            lngID = .RowData(.Row)
        End With
        Call FillGrid(rsTemp, lngID)
    
        mshSum_S.Active = blnActive '�ָ�
        mshSum_S.msfObj.MousePointer = vbDefault
        MousePointer = vbDefault
        cmdRestore.Enabled = blnEnable
        cmdSave.Enabled = blnEnable
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    'ֻˢ���б�����
    FillSum
End Sub

Private Sub mshSum_S_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_S_LostFocus()
    mshSum_S.CmdVisible = False
    mshSum_S.CboVisible = False
    
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + x - msngStartX
        If sngTemp > 1500 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwMain_S.SelectedItem
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    
    Set objPrint.Body = mshSum_S.msfObj
    objPrint.Title.Text = nod.Text & "���շ�ϸĿҽ����Ŀ��Ӧ��"
    'objRow.Add "ҽԺ���ƣ�" & gstr��λ����
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & gstrUserName
    objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
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
End Sub
    
Private Sub Fill����()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    'ֻˢ���б�����
    
    '���Ȼ��ҽ������
    mshSum_S.Active = True
    If gintInsure = TYPE_�ɶ��ϳ� Then
        If mcnYB.State = 1 Then mcnYB.Close
        mcnYB.Open GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
        Exit Sub
    End If
    
    gstrSQL = "select ID,����,���� from ����֧������ " & _
              "where ����=" & cmb����.ItemData(cmb����.ListIndex) & " order by ����"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    mshSum_S.Clear
    Do Until rsTemp.EOF
        mshSum_S.AddItem rsTemp("����") & "." & rsTemp("����")
        mshSum_S.ItemData(mshSum_S.NewIndex) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    
    Select Case cmb����.ItemData(cmb����.ListIndex)
        Case TYPE_������, TYPE_����������, TYPE_����, TYPE_�Ͻ�, TYPE_����
            If mcnYB.State = adStateClosed Then
                '���ȶ���������������
                gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & cmb����.ItemData(cmb����.ListIndex)
                Call OpenRecordset(rsTemp, Me.Caption)
                Do Until rsTemp.EOF
                    strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Select Case rsTemp("������")
                        Case "ҽ��������"
                            strServer = strTemp
                        Case "ҽ���û���"
                            strUser = strTemp
                        Case "ҽ���û�����"
                            strPass = strTemp
                    End Select
                    rsTemp.MoveNext
                Loop
                If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
                    Exit Sub
                End If
            End If
        Case TYPE_�山ũҽ
            Dim strDatabase As String
            If mcnYB.State = adStateClosed Then
                '���ȶ���������������
                gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & cmb����.ItemData(cmb����.ListIndex)
                Call OpenRecordset(rsTemp, Me.Caption)
                Do Until rsTemp.EOF
                    strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Select Case rsTemp("������")
                        Case "ҽ��������"
                            strServer = strTemp
                        Case "ҽ���û���"
                            strUser = strTemp
                        Case "ҽ���û�����"
                            strPass = strTemp
                        Case "ҽ��ʵ����"
                            strDatabase = strTemp
                    End Select
                    rsTemp.MoveNext
                Loop
            End If
            If Not OpenSQLServer(mcnYB, strServer, strUser, strPass, strDatabase) Then Exit Sub
        Case TYPE_������
            '������ͨҽ��ǰ�û����Ͳ����޸ġ���Ϊ��Ҫ�����޸ļ�¼
            If ���ҽ��������_���� = False Then mshSum_S.Active = False
        Case TYPE_ͭ��
            '������ͨҽ��ǰ�û����Ͳ����޸ġ���Ϊ��Ҫ�����޸ļ�¼
            If ���ҽ��������_ͭ�� = False Then mshSum_S.Active = False
        Case TYPE_�����山
            If gcnOracle_CQYB Is Nothing Or gcnOracle_CQYB.State <> 1 Then
                Call ҽ����ʼ��_�����山
            End If
        Case TYPE_ǭ��
            If gcnOracle_ǭ�� Is Nothing Then
                
                '�����´�ҽ��
                gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & cmb����.ItemData(cmb����.ListIndex)
                Call OpenRecordset(rsTemp, Me.Caption)
                Do Until rsTemp.EOF
                    strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Select Case rsTemp("������")
                        Case "ҽ��������"
                            strServer = strTemp
                        Case "ҽ���û���"
                            strUser = strTemp
                        Case "ҽ���û�����"
                            strPass = strTemp
                    End Select
                    rsTemp.MoveNext
                Loop
                Set gcnOracle_ǭ�� = New ADODB.Connection
                If OraDataOpen(gcnOracle_ǭ��, strServer, strUser, strPass) = False Then
                    Exit Sub
                End If
            End If
    End Select
End Sub

Private Function FillTree() As Boolean
'����:װ���շ������շ�ϸĿ�����з��ൽtvwMain_S
    '�����������ڵ�����������KEYֵ��һ���ַ������ڶ�λ��������

    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim nod As Node
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    
    MousePointer = vbHourglass
    
    mstrKey = ""     'ȫ��ˢ��ʱ���൱���û�û����κνڵ�
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    gstrSQL = "select ����,��� from �շ���� where ����<>'5' and ����<>'6' and ����<>'7' order by ����"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    LockWindowUpdate tvwMain_S.hwnd
    'ɾ�����нڵ�
    With tvwMain_S.Nodes
        .Clear
        '�������
        Do Until rsTemp.EOF
            .Add , , "R" & rsTemp("����"), "��" & rsTemp("����") & "��" & rsTemp("���"), "R", "R"
            tvwMain_S.Nodes("R" & rsTemp("����")).Sorted = True
            rsTemp.MoveNext
        Loop
        .Add , , "D5", "��5������ҩ", "R", "R"
        tvwMain_S.Nodes("D5").Sorted = True
        .Add , , "E6", "��6���г�ҩ", "R", "R"
        tvwMain_S.Nodes("E6").Sorted = True
        .Add , , "F7", "��7���в�ҩ", "R", "R"
        tvwMain_S.Nodes("F7").Sorted = True
        
        '������ͨ�շ���Ŀ����ڵ�
        gstrSQL = "select id,�ϼ�id,���,����,���� from �շ�ϸĿ  where ���<>'5' and ���<>'6' and ���<>'7' and ĩ�� <> 1 " & _
             " start with �ϼ�ID is null  connect by prior id=�ϼ�ID "
        Call OpenRecordset(rsTemp, Me.Caption)
    
        Do Until rsTemp.EOF
            '��ӽڵ�
            If IsNull(rsTemp("�ϼ�id")) Then
                .Add "R" & rsTemp("���"), tvwChild, "C" & rsTemp("���") & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "C", "C"
            Else
                .Add "C" & rsTemp("���") & rsTemp("�ϼ�id"), tvwChild, "C" & rsTemp("���") & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "C", "C"
            End If
            tvwMain_S.Nodes("C" & rsTemp("���") & rsTemp("ID")).Sorted = True
            rsTemp.MoveNext
        Loop
    
        '��װ��ҩƷ��;���������
        gstrSQL = "select id,�ϼ�id,����,����,���� from ҩƷ��;����  " & _
             " start with �ϼ�ID is null  connect by prior id=�ϼ�ID "
        Call OpenRecordset(rsTemp, Me.Caption)
    
        Do Until rsTemp.EOF
            '��ӽڵ�
            Select Case rsTemp("����")
                Case "�г�ҩ"
                    If IsNull(rsTemp("�ϼ�id")) Then
                        Set nod = .Add("E6", tvwChild, "E6" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    Else
                        Set nod = .Add("E6" & rsTemp("�ϼ�id"), tvwChild, "E6" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    End If
                Case "�в�ҩ"
                    If IsNull(rsTemp("�ϼ�id")) Then
                        Set nod = .Add("F7", tvwChild, "F7" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    Else
                        Set nod = .Add("F7" & rsTemp("�ϼ�id"), tvwChild, "F7" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    End If
                Case Else '����ҩ
                    If IsNull(rsTemp("�ϼ�id")) Then
                        Set nod = .Add("D5", tvwChild, "D5" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    Else
                        Set nod = .Add("D5" & rsTemp("�ϼ�id"), tvwChild, "D5" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    End If
                End Select
            nod.Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    
    LockWindowUpdate 0
    MousePointer = 0
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes(1)
        nod.Selected = True
    Else
        Err.Clear
        nod.Selected = True
        nod.EnsureVisible
    End If
    Call FillSum
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
    MousePointer = 0
End Function

Public Sub FillSum(Optional ByVal blnForce As Boolean = False)
'����:װ�����ͳ������
    Dim rsTemp As New ADODB.Recordset
    Dim nod As Node
    Dim str���ʷ��� As String
    Dim lngID As Long

    If tvwMain_S.SelectedItem Is Nothing Then
        ClearGrid mshSum_S
        Call MenuSet
        Exit Sub
    End If
    
    If blnForce = False Then
        If mstrKey = tvwMain_S.SelectedItem.Key And mint���� = cmb����.ItemData(cmb����.ListIndex) Then
            '��ȫû�иı䣬������ˢ��
            Exit Sub
        End If
        
        If cmdSave.Enabled = True Then
            If mint���� <> TYPE_������ Then
                '�Ѿ��޸ģ���ʾ�Ƿ���Ҫ���浱ǰ������
                If MsgBox("������Ŀ�Ѿ��޸ģ��Ƿ���Ҫ���棿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    Call cmdSave_Click
                End If
            Else
                Call cmdSave_Click
            End If
        End If
    End If
    
    cmdSave = False
    cmdRestore = False
    '��ȡ������Ŀ���������ؼ���
    mstrKey = tvwMain_S.SelectedItem.Key
    mint���� = cmb����.ItemData(cmb����.ListIndex)
    If mint���� = type_���������� Or mint���� = type_������ Then
        Call InitSum
    End If
    Set nod = tvwMain_S.SelectedItem
    
    '���ݲ�ͬ�Ľڵ㣬������ͬ����ʾ
    '�������Ҫ����ʾһ��
    If Mid(nod.Key, 2, 1) = "5" Or Mid(nod.Key, 2, 1) = "6" Or Mid(nod.Key, 2, 1) = "7" Then
        'ҩƷ�Ĵ���Ҫ�鷳һЩ
        mshSum_S.TextMatrix(0, col����) = "����"
        
        Select Case Left(nod.Key, 1)
            Case "D"
                str���ʷ��� = "����ҩ"
            Case "E"
                str���ʷ��� = "�г�ҩ"
            Case "F"
                str���ʷ��� = "�в�ҩ"
        End Select
        
        If nod.Image = "R" Then
            gstrSQL = "select A.ҩƷID as ID,A.����,B.ͨ������||decode(M.����,null,'',b.ͨ������,'',' ��'||M.����||'��') as ����,A.���,A.����,A.�ۼ۵�λ as ��λ,D.�Ƿ���,E.���� ���� " & _
                        "from ҩƷĿ¼ A,ҩƷ��Ϣ B,�շ�ϸĿ D,ҩƷ���� E,(Select distinct ҩƷid,���� from ҩƷ���� ) M " & _
                        "where A.ҩ��ID=B.ҩ��ID and d.id=M.ҩƷID(+) and B.����=E.����(+) and B.���ʷ���='" & str���ʷ��� & "'" & _
                        "      and A.ҩƷID=D.ID and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select A.ҩƷID as ID,A.����,B.ͨ������||decode(M.����,null,'',b.ͨ������,'',' ��'||M.����||'��') as ����,A.���,A.����,A.�ۼ۵�λ as ��λ,D.�Ƿ���,E.���� ���� " & _
                      "from ҩƷĿ¼ A,ҩƷ��Ϣ B,�շ�ϸĿ D,ҩƷ���� E,(Select distinct ҩƷid,���� from ҩƷ����) M ,(select ID from ҩƷ��;���� start with ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID) C " & _
                      "where A.ҩ��ID=B.ҩ��ID and B.����=E.����(+) and d.id=M.ҩƷID(+) and B.���ʷ���='" & str���ʷ��� & "' and B.��;����ID=C.ID" & _
                      "       and A.ҩƷID=D.ID and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
        End If
        
    Else
        '��ҩƷ�����׵ö���
        mshSum_S.TextMatrix(0, col����) = "˵��"
        
        If nod.Image = "R" Then
            gstrSQL = "select id,����,����,���,˵�� as ����,���㵥λ as ��λ,�Ƿ���,'' ���� from �շ�ϸĿ where ĩ��=1 and ���='" & Mid(nod.Key, 2, 1) & "' " & _
                        " and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select id,����,����,���,˵�� as ����,���㵥λ as ��λ,�Ƿ���,'' ���� from �շ�ϸĿ where ĩ��=1 and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
                        " start with �ϼ�ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID "
        End If
    End If
    
    'Modified by ZYB 2004-08-17
    If mint���� = TYPE_��ɽ Then
        gstrSQL = "select A.ID,A.����,A.����,A.���,A.����,A.����,A.��λ,A.�Ƿ���,D.�۸�,'' as �ı䷽ʽ" & _
                   " ,B.��Ŀ����,substr(B.��ע,Instr(B.��ע,'|',1,3)+1)||'-'||B.��Ŀ���� AS ��Ŀ����,B.��ע,B.��Ŀ���� as ԭ����,B.�Ƿ�ҽ��,B.����ID,C.���� as �������,C.���� as �������� " & _
                   " from (" & gstrSQL & ") A,����֧����Ŀ B,����֧������ C," & _
                   "      (select sum(�ּ�) as �۸�,�շ�ϸĿID from �շѼ�Ŀ where ִ������<=sysdate and (��ֹ����>=sysdate or ��ֹ���� is null) group by �շ�ϸĿID) D " & _
                   " Where A.ID=B.�շ�ϸĿID(+) and B.����ID=c.id(+)  and B.����(+)= " & mint���� & _
                   "       and A.ID=D.�շ�ϸĿID(+)  "
    Else
        gstrSQL = "select A.ID,A.����,A.����,A.���,A.����,A.����,A.��λ,A.�Ƿ���,D.�۸�,'' as �ı䷽ʽ" & _
                   " ,B.��Ŀ����,B.��Ŀ����,B.��ע,B.��Ŀ���� as ԭ����,B.�Ƿ�ҽ��,B.����ID,C.���� as �������,C.���� as �������� " & _
                   " from (" & gstrSQL & ") A,����֧����Ŀ B,����֧������ C," & _
                   "      (select sum(�ּ�) as �۸�,�շ�ϸĿID from �շѼ�Ŀ where ִ������<=sysdate and (��ֹ����>=sysdate or ��ֹ���� is null) group by �շ�ϸĿID) D " & _
                   " Where A.ID=B.�շ�ϸĿID(+) and B.����ID=c.id(+)  and B.����(+)= " & mint���� & _
                   "       and A.ID=D.�շ�ϸĿID(+)  "
    End If
    
    MousePointer = 11
    Call OpenRecordset(rsTemp, Me.Caption)
    
    lngID = mshSum_S.RowData(mshSum_S.Row)
    Call FillGrid(rsTemp, lngID)
    
    stbThis.Panels(2).Text = "�����շ���Ŀ" & rsTemp.RecordCount & "��"
    
    MousePointer = 0
    Call MenuSet
End Sub

Private Sub FillGrid(rsTemp As ADODB.Recordset, ByVal lngID As Long)
    Dim strSort As String
    Dim strDemo As String
    Dim intMatch As Integer
    Dim lngRow As Long, lngRowSelect As Long
    
    Select Case mlngCol
        Case cOL����
            strSort = "����"
        Case cOL����
            strSort = "����"
        Case COL���
            strSort = "���"
        Case col����
            strSort = "����"
        Case COL��λ
            strSort = "��λ"
        Case col�۸�
            strSort = "�۸�"
        Case COLҽ������
            strSort = "��Ŀ����"
        Case colҽ������
            strSort = "��Ŀ����"
        Case col��������
            strSort = "��������"
        Case col��ҽ��
            strSort = "�Ƿ�ҽ��"
        Case Else
            rsTemp.Sort = "����"
    End Select
    rsTemp.Sort = strSort & IIf(mblnDesc, " DESC", "")
    
    mshSum_S.TxtVisible = False
    mshSum_S.CboVisible = False
    mshSum_S.Redraw = False
    ClearGrid mshSum_S
    If rsTemp.RecordCount <> 0 Then
        mshSum_S.Rows = rsTemp.RecordCount + 1
    End If
    lngRow = 1
    With mshSum_S
        Do Until rsTemp.EOF
            If rsTemp("ID") = lngID Then
                lngRowSelect = lngRow
            End If
            
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, cOL����) = rsTemp("����")
            .TextMatrix(lngRow, cOL����) = rsTemp("����")
            .TextMatrix(lngRow, COL���) = IIf(IsNull(rsTemp("���")), "", rsTemp("���"))
            .TextMatrix(lngRow, col����) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, COL����) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, COL��λ) = IIf(IsNull(rsTemp("��λ")), "", rsTemp("��λ"))
            .TextMatrix(lngRow, col�۸�) = IIf(rsTemp("�Ƿ���") = 0, Format(rsTemp("�۸�"), "0.000"), "")
            .TextMatrix(lngRow, col�ı䷽ʽ) = IIf(IsNull(rsTemp("�ı䷽ʽ")), "", rsTemp("�ı䷽ʽ"))
            .TextMatrix(lngRow, COLҽ������) = IIf(IsNull(rsTemp("��Ŀ����")), "", rsTemp("��Ŀ����"))
            .TextMatrix(lngRow, colҽ������) = IIf(IsNull(rsTemp("��Ŀ����")), "", rsTemp("��Ŀ����"))
            .TextMatrix(lngRow, colԭ����) = IIf(IsNull(rsTemp("ԭ����")), "", rsTemp("ԭ����"))
            .TextMatrix(lngRow, col����ID) = IIf(IsNull(rsTemp("����ID")), "", rsTemp("����ID"))
            .TextMatrix(lngRow, col��ҽ��) = IIf(rsTemp("�Ƿ�ҽ��") = "0", "��", "")
            If mint���� = TYPE_������ Then
                intMatch = 0
                strDemo = IIf(IsNull(rsTemp("��ע")), "", rsTemp("��ע"))
                If InStr(1, strDemo, "||") <> 0 Then
                    If InStr(1, strDemo, "^^") <> 0 Then
                        .TextMatrix(lngRow, colҽ������) = Split(strDemo, "^^")(0)
                        .TextMatrix(lngRow, colҽ������) = Split(.TextMatrix(lngRow, colҽ������), "||")(3)
                        .TextMatrix(lngRow, COLҽ����ע) = Split(strDemo, "^^")(0)
                    Else
                        .TextMatrix(lngRow, colҽ������) = strDemo
                        .TextMatrix(lngRow, colҽ������) = Split(.TextMatrix(lngRow, colҽ������), "||")(3)
                        .TextMatrix(lngRow, COLҽ����ע) = strDemo
                    End If
                    If InStr(1, strDemo, "^^") <> 0 Then
                        If InStr(1, Split(strDemo, "^^")(1), "||") <> 0 Then
                            .TextMatrix(lngRow, colƥ�����к�) = Split(Split(strDemo, "^^")(1), "||")(0)
                            intMatch = Split(Split(strDemo, "^^")(1), "||")(1)
                        Else
                            .TextMatrix(lngRow, colƥ�����к�) = Split(strDemo, "^^")(1)
                        End If
                    End If
                Else
                    .TextMatrix(lngRow, COLҽ����ע) = strDemo
                End If
                If intMatch = 1 Then
                    .TextMatrix(lngRow, col��˱�־) = "��"
                ElseIf intMatch = 2 Then
                    .TextMatrix(lngRow, col��˱�־) = "��"
                End If
            Else
                .TextMatrix(lngRow, COLҽ����ע) = IIf(IsNull(rsTemp("��ע")), "", rsTemp("��ע"))
                .TextMatrix(lngRow, colƥ�����к�) = ""
            End If
            
            If IsNull(rsTemp("�������")) Or IsNull(rsTemp("��������")) Then
                .TextMatrix(lngRow, col��������) = ""
            Else
                .TextMatrix(lngRow, col��������) = rsTemp("�������") & "." & rsTemp("��������")
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    If lngRowSelect > 0 And lngRowSelect < mshSum_S.Rows - 1 Then
        mshSum_S.msfObj.TopRow = lngRowSelect
        mshSum_S.Row = lngRowSelect
    End If
    mshSum_S.Redraw = True
End Sub

Private Sub ClearGrid(objGrid As Object)
'���ܣ�������,����ɲ��ֳ�ʼ��
    Dim i As Long
    
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    With objGrid.msfObj
        .Rows = 2
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
        Next
    
    End With
End Sub

Private Sub MenuSet()
'����:��ʾ�˵��͹�������״̬(��ӡ)
    Dim blnPrint As Boolean
    
    blnPrint = Not (mshSum_S.Rows = 2 And mshSum_S.TextMatrix(1, 0) = "")

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
    
    If InStr(mstrȨ��, "��ɾ��") > 0 Then
        mshSum_S.Active = blnPrint
        If mint���� = TYPE_������ Then
            'ǿ�Ʋ���ʹ��
            If gcn����.State = adStateClosed Then mshSum_S.Active = False
        End If
    Else
        mshSum_S.Active = False
    End If
End Sub

Public Sub ShowForm(frmParent As Form)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select ���,���� from ������� where nvl(�Ƿ��ֹ,0)<>1  order by ���"
    Call OpenRecordset(rsTemp, "ȡ�������")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��ñ�����𣬲���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm������Ŀ.Visible = True Then
        frm������Ŀ.Show
        Exit Sub
    End If
    
    mstrȨ�� = gstrPrivs
    frm������Ŀ.Show , frmParent
End Sub

'Modified By ���� ��������ɳ ԭ����������������Ŀ��ҽ����Ŀ��ƥ��
Private Sub SetItemMatch(Optional ByVal blnɾ�� As Boolean = True)
    'ҽ����ע�н����������Ϣ
    'intEdit����1����;2�޸�;3ɾ��
    'col�ı䷽ʽ�����ջ�ɾ����ִ��ɾ��ƥ��������޸�ִ����ɾ����������������ִ����������
    Dim strƥ������ As String, str���� As String, str��� As String, strҽԺ���� As String
    Dim rsTemp As New ADODB.Recordset
     
    Select Case mint����
    Case TYPE_������
        '��������ͨ�����������޸Ļ�ɾ��
        If int��˱�־ = 1 And mint���õ��� = 0 Then
            MsgBox "����Ŀ�Ѿ�ͨ��ҽ��������ˣ��������޸Ļ�ɾ��������ҽ��������ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not classInsure.InitInsure(gcnOracle, TYPE_������) Then Exit Sub
        strƥ������ = TranClass
        str���� = "��"

        If Trim(mshSum_S.TextMatrix(mshSum_S.Row, colƥ�����к�)) <> "" Then
            'ɾ��ƥ����Ϣ����������޸ģ�ֱ���˳�
'            1   serial_match    ƥ�����к�  12  ��
'            2   audit_flag  ��˱�־    1   ��  "0"��δ��ˣ�"2"�����δͨ��
'            3   edit_staff  ����Ա����  5   ��
'            4   edit_man    ����Ա����  10  ��
            If Not ���ýӿ�_׼��_������(Function_������.��Ŀƥ��_ɾ��ƥ����Ϣ) Then Exit Sub
            gstrField_������ = "serial_match||audit_flag||edit_staff||edit_man"
            gstrValue_������ = mshSum_S.TextMatrix(mshSum_S.Row, colƥ�����к�) & "||" & int��˱�־ & "||" & gCominfo_������.����Ա���� & "||" & gstrUserName
            If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Sub
            If Not ���ýӿ�_ִ��_������ Then Exit Sub
            mshSum_S.TextMatrix(mshSum_S.Row, colƥ�����к�) = ""
        End If
        
        If Not blnɾ�� Then
            'ִ������ƥ�䶯�����޸�����������ɾ���ˣ�
'            1   hospital_idҽ�ƻ�������    20  ��
'            2   match_type ƥ������        1   ��  "0"��������Ŀƥ�䣻"1"����ҩƥ�䣻"2"���г�ҩƥ�䣻"3"���в�ҩƥ��
'            3   hosp_code  ҽԺĿ¼����    20  ��
'            4   hosp_name  ҽԺĿ¼����    60  ��
'            5   hosp_model ҽԺĿ¼����    20  ��
'            6   price      ����            8   ��
'            7   item_code  ����Ŀ¼����    20  ��
'            8   item_name  ����Ŀ¼����    60  ��
'            9   model_name ����Ŀ¼����    20  ��
'            10  effect_date��Ч����            ��  ��ʽ:YYYY-MM-DD
'            11  expire_dateʧЧ����            ��  ��ʽ:YYYY-MM-DD
'            12  edit_staff ����Ա����      5   ��
'            13  edit_man   ����Ա����      10  ��
            If strƥ������ <> "0" Then
                gstrSQL = "select C.���� ����  " & _
                         " from ҩƷ��Ϣ A,ҩƷĿ¼  B,ҩƷ���� C " & _
                         " where A.ҩ��ID=B.ҩ��ID And A.����=C.���� And B.ҩƷID = " & mshSum_S.RowData(mshSum_S.Row)
                Call OpenRecordset(rsTemp, "��ȡҩƷ�ļ�������")
                str���� = ToVarchar(rsTemp!����, 20)
            End If
            'ȡ�շ�ϸĿ�ı�ʶ����ΪҽԺ�����ϴ�
            If Not (Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "5" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "6" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "7") Then
                gstrSQL = "Select Decode(TRIM(��ʶ����),NULL,����,'',����,��ʶ����) ����,��� From �շ�ϸĿ Where ID=" & mshSum_S.RowData(mshSum_S.Row)
            Else
                gstrSQL = "Select Decode(Trim(��ʶ��),NULL,����,'',����,��ʶ��) ����,��� From ҩƷĿ¼ Where ҩƷID=" & mshSum_S.RowData(mshSum_S.Row)
            End If
            Call OpenRecordset(rsTemp, "ȡҽԺ����")
            strҽԺ���� = Nvl(rsTemp!����)
            str��� = Nvl(rsTemp!���)
            
            If Not ���ýӿ�_׼��_������(Function_������.��Ŀƥ��_��Ŀƥ��) Then Exit Sub
            If Not ���ýӿ�_ָ����¼��_������("MatchInfo") Then Exit Sub
            
            gstrField_������ = "hospital_id||match_type||hosp_code||hosp_name||hosp_model||spec||price||" & _
            "item_code||item_name||model_name||effect_date||expire_date||edit_staff||edit_man"
            gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & strƥ������ & "||" & _
                    strҽԺ���� & "||" & mshSum_S.TextMatrix(mshSum_S.Row, cOL����) & "||" & _
                    str���� & "||" & str��� & "||" & mshSum_S.TextMatrix(mshSum_S.Row, col�۸�) & "||" & _
                    mshSum_S.TextMatrix(mshSum_S.Row, COLҽ������) & "||" & mshSum_S.TextMatrix(mshSum_S.Row, colҽ������) & "||" & _
                    mshSum_S.TextMatrix(mshSum_S.Row, colҽ������) & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & "||" & _
                    "3000-01-01||" & gCominfo_������.����Ա���� & "||" & gstrUserName
            If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Sub
            If Not ���ýӿ�_ִ��_������() Then Exit Sub
            
            '��ȡƥ�����кţ�����
            If Not ���ýӿ�_ָ����¼��_������("MatchItem") Then Exit Sub
            Call ���ýӿ�_��ȡ����_������("serial_match", str����)
            mshSum_S.TextMatrix(mshSum_S.Row, colƥ�����к�) = str����
            
            '���·������ͣ�ҩƷ�Ÿ��£�
            If Not (Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "5" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "6" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "7") Then Exit Sub
            Call ���ýӿ�_��ȡ����_������("Staple_flag", str����)
            If Val(str����) = 1 Then
                str���� = "����ҩƷ"
            ElseIf Val(str����) = 2 Then
                str���� = "����ҩƷ"
            Else
                str���� = "�ǻ���ҩƷ"
            End If
            gstrSQL = "ZL_���·�������('" & mshSum_S.TextMatrix(mshSum_S.Row, COLҽ������) & "','" & str���� & "')"
            Call ExecuteProcedure("���·�������")
        End If
    End Select
End Sub

'Modified By ���� ��������ɳ ԭ���������������·��ҽ������������ҽ�����룬����ʾҽ�������Ƿ���������ƥ����Ϣ
Private Sub GetItemMatchInfo()
    Dim strƥ������ As String, str��Ŀ���� As String, strMatch As String
    Dim intԭ��˱�־ As Integer
    Dim rsTemp As New ADODB.Recordset
    
    intԭ��˱�־ = IIf(mshSum_S.TextMatrix(mshSum_S.Row, col��˱�־) = "��", 1, IIf(mshSum_S.TextMatrix(mshSum_S.Row, col��˱�־) = "��", 2, 0))
    int��˱�־ = 0
    stbThis.Panels(2).Text = ""
    If Trim(mshSum_S.TextMatrix(mshSum_S.Row, COLҽ������)) = "" Then Exit Sub
    
    If mint���� = TYPE_������ Then

        'ȡ�շ�ϸĿ�ı�ʶ����ΪҽԺ�����ϴ�
        If Not (Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "5" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "6" Or Mid(tvwMain_S.SelectedItem.Key, 2, 1) = "7") Then
            gstrSQL = "Select Decode(TRIM(��ʶ����),NULL,����,'',����,��ʶ����) ���� From �շ�ϸĿ Where ID=" & mshSum_S.RowData(mshSum_S.Row)
        Else
            gstrSQL = "Select Decode(Trim(��ʶ��),NULL,����,'',����,��ʶ��) ���� From ҩƷĿ¼ Where ҩƷID=" & mshSum_S.RowData(mshSum_S.Row)
        End If
        Call OpenRecordset(rsTemp, "ȡҽԺ����")
        str��Ŀ���� = Nvl(rsTemp!����)

'        1   hospital_id    ҽ�ƻ�������    20  ��
'        2   his_item_code  ҽԺĿ¼����    20  ��
'        3   medi_item_type ƥ������        1   ��  "0"��������Ŀƥ�䣻"1"����ҩƥ�䣻"2"���г�ҩƥ�䣻"3"���в�ҩƥ��
'        4   fee_date       ���÷���ʱ��        ��  ��ʽ��YYYY-MM-DD
        stbThis.Panels(2).Text = "��ȡ����Ŀ��ƥ����Ϣʧ�ܣ�"
        If Not classInsure.InitInsure(gcnOracle, TYPE_������) Then Exit Sub
        If Not ���ýӿ�_׼��_������(Function_������.��Ŀƥ��_ȡ������Ŀƥ����Ϣ) Then Exit Sub
        strƥ������ = TranClass
        gstrField_������ = "hospital_id||his_item_code||medi_item_type||fee_date"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & Trim(str��Ŀ����) & "||" & _
                strƥ������ & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-DD")
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Sub
        If Not ���ýӿ�_ִ��_������() Then Exit Sub
        'ָ����¼��
        If Not ���ýӿ�_ָ����¼��_������("MatchInfo") Then Exit Sub
        Call ���ýӿ�_��ȡ����_������("audit_flag", strMatch)
        Call DebugTool("��˱�־��" & strMatch)
        If strMatch = "" Then strMatch = "0"
        int��˱�־ = Val(strMatch)
        
        If int��˱�־ = 1 Then
            mshSum_S.TextMatrix(mshSum_S.Row, col��˱�־) = "��"
        ElseIf int��˱�־ = 2 Then
            mshSum_S.TextMatrix(mshSum_S.Row, col��˱�־) = "��"
        Else
            mshSum_S.TextMatrix(mshSum_S.Row, col��˱�־) = ""
        End If
        stbThis.Panels(2).Text = "ƥ����Ϣ��" & IIf(strMatch = "0", "δ���", IIf(strMatch = "1", "���ͨ��", "���δͨ��"))
        
        '���±���֧����Ŀ
        If int��˱�־ <> intԭ��˱�־ Then Call ��Ǹı�
    End If
End Sub

'Modified By ���� ��������ɳ ԭ��ת�����Ϊҽ���ӿ���Ҫ��ƥ������
Private Function TranClass() As String
    Dim strClass As String
    strClass = Mid(tvwMain_S.SelectedItem.Key, 2, 1)
    Select Case strClass
    Case "5"
        TranClass = "1"
    Case "6"
        TranClass = "2"
    Case "7"
        TranClass = "3"
    Case Else
        TranClass = "0"
    End Select
End Function

Private Function CheckValid(ByVal strCode As String) As Boolean
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    '�������Ƿ�ƥ��
    gstrSQL = "Select ��ע From ������Ŀ Where ����='" & strCode & "'"
    Call OpenRecordset(rsTemp, "��ȡ����")
    str���� = Mid(rsTemp!��ע, 1, 1)
    
    If str���� <> TranClass Then
        MsgBox "��ע�⣺��ҽ����Ŀ��������뵱ǰѡ���ҽԺ��Ŀ�����ͬ��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function
